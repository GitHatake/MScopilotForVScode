import * as vscode from "vscode";
import { getConfig, MsCopilotConfig } from "./config";
import { initLogger, log } from "./logger";
import { BrowserSession, BrowserSessionError, createBrowserSession } from "./bridge";
import {
  RawTokenCandidate,
  formatInventory,
  inventoryTokens,
  parseWsUrlToken,
  pickSydneyToken,
  tokenFromSecret,
} from "./bridge/browserSession";
import { SubstrateClient } from "./copilot/substrateClient";
import { HistoryStore } from "./history/store";
import { ChatController } from "./chat/participant";

/**
 * BrowserSession をライフサイクル管理する。接続やページを保持するため使い回すが、
 * 接続に影響する設定が変わったら破棄して作り直す。
 */
class SessionManager {
  private session: BrowserSession | undefined;
  private signature = "";
  private pending: Promise<BrowserSession> | undefined;

  constructor(private readonly storageDir: string) {}

  async get(config: MsCopilotConfig): Promise<BrowserSession> {
    const sig = [
      config.transport,
      config.debugPort,
      config.browserPath,
      config.userDataDir,
      config.startUrl,
    ].join("|");
    // 同時呼び出しでブラウザを二重起動しないよう直列化する。
    while (this.pending) {
      await this.pending.catch(() => undefined);
    }
    if (this.session && sig === this.signature) {
      return this.session;
    }
    this.pending = this.create(config, sig);
    try {
      return await this.pending;
    } finally {
      this.pending = undefined;
    }
  }

  private async create(config: MsCopilotConfig, sig: string): Promise<BrowserSession> {
    if (this.session) {
      await this.session.dispose().catch(() => undefined);
      this.session = undefined;
      this.signature = "";
    }
    const session = createBrowserSession(config, this.storageDir);
    // ensureReady が失敗したら壊れたセッションをキャッシュしない。
    await session.ensureReady();
    this.session = session;
    this.signature = sig;
    return session;
  }

  async dispose(): Promise<void> {
    await this.session?.dispose().catch(() => undefined);
    this.session = undefined;
  }
}

export async function activate(context: vscode.ExtensionContext): Promise<void> {
  initLogger(context);
  log.info("MS Copilot Chat 拡張を起動しました");

  const storageDir = context.globalStorageUri.fsPath;
  const sessions = new SessionManager(storageDir);
  const history = new HistoryStore(context);

  const controller = new ChatController({
    history,
    getMaxMentionBytes: () => getConfig().maxMentionBytes,
    getClient: () => {
      const config = getConfig();
      // ask() 内で getToken → ensureReady が走るよう、遅延で session を渡す。
      const lazy = new LazySession(sessions, config);
      return new SubstrateClient(lazy, config);
    },
  });

  const participant = vscode.chat.createChatParticipant("mscopilot.chat", controller.handle);
  participant.iconPath = new vscode.ThemeIcon("comment-discussion");
  context.subscriptions.push(participant);

  context.subscriptions.push(
    vscode.commands.registerCommand("mscopilot.showLog", () => log.show()),
    vscode.commands.registerCommand("mscopilot.signIn", async () => {
      try {
        await vscode.window.withProgress(
          { location: vscode.ProgressLocation.Notification, title: "MS Copilot に接続中…" },
          async () => {
            const session = await sessions.get(getConfig());
            const token = await session.getToken(true);
            vscode.window.showInformationMessage(
              `MS Copilot への接続に成功しました(テナント: ${token.tenantId.slice(0, 8)}…)。`,
            );
          },
        );
      } catch (e) {
        showSessionError(e);
      }
    }),
    vscode.commands.registerCommand("mscopilot.diagnoseToken", async () => {
      try {
        let candidates: RawTokenCandidate[] = [];
        let href = "";
        let wsUrl = "";
        let bearer = "";
        await vscode.window.withProgress(
          { location: vscode.ProgressLocation.Notification, title: "トークン取得を診断中…" },
          async () => {
            const session = await sessions.get(getConfig());
            href = await session.currentUrl();
            candidates = await session.collectRawTokens();
            wsUrl = (await session.harvestWsTemplate?.()) ?? "";
            bearer = (await session.harvestBearer?.()) ?? "";
          },
        );
        const inv = inventoryTokens(candidates);
        const picked = pickSydneyToken(candidates);
        const wsToken = wsUrl ? parseWsUrlToken(wsUrl) : undefined;
        const bearerToken = bearer ? tokenFromSecret(bearer) : undefined;
        const remain = (t: { expiresAt: number }) => Math.round((t.expiresAt - Date.now()) / 60000);
        log.info("=== トークン取得診断 ===");
        log.info(`現在のページ: ${href || "(不明)"}`);
        log.info(`検出したトークン候補 ${inv.length} 件(★=採用候補, ○=substrate):\n${formatInventory(inv)}`);
        log.info("取得経路の状態(上から優先):");
        log.info(
          picked
            ? `  1) MSAL ストレージ: OK(tenant=${picked.tenantId.slice(0, 8)}… 残り約 ${remain(picked)} 分)`
            : "  1) MSAL ストレージ: なし(このテナントは access_token を永続化しない可能性)",
        );
        log.info(
          bearerToken
            ? `  2) substrate HTTP の Bearer: OK(tenant=${bearerToken.tenantId.slice(0, 8)}… 残り約 ${remain(bearerToken)} 分)`
            : "  2) substrate HTTP の Bearer: 未捕捉(ページの substrate 通信をまだ観測していません)",
        );
        log.info(
          wsToken
            ? `  3) ページ実接続の WebSocket: OK(tenant=${wsToken.tenantId.slice(0, 8)}… 残り約 ${remain(wsToken)} 分)`
            : "  3) ページ実接続の WebSocket: 未捕捉(チャット接続がまだ張られていません)",
        );
        const acquirable = picked || bearerToken || wsToken;
        if (!acquirable) {
          log.warn(
            "→ どの経路でもトークンを採取できていません。対象ブラウザで Copilot チャット" +
              `(${getConfig().startUrl})を開き、ページを表示した状態で数秒待つか、一度メッセージを送ってから再実行してください。`,
          );
        }
        log.show();
        const via = picked ? "MSALストレージ" : bearerToken ? "HTTP Bearer" : wsToken ? "実接続WS" : "";
        vscode.window.showInformationMessage(
          acquirable
            ? `トークンを取得できます(${via} 経由)。詳細はログを参照してください。`
            : `トークンを取得できませんでした。ログの案内に従ってください(候補 ${inv.length} 件)。`,
        );
      } catch (e) {
        showSessionError(e);
      }
    }),
    { dispose: () => void sessions.dispose() },
  );
}

export function deactivate(): void {
  // subscriptions の dispose で SessionManager も破棄される。
}

/**
 * SubstrateClient に渡す BrowserSession のプロキシ。呼び出し時に SessionManager から
 * 実体を取得することで、設定変更や初回接続を吸収する。
 */
class LazySession implements BrowserSession {
  constructor(
    private readonly sessions: SessionManager,
    private readonly config: MsCopilotConfig,
  ) {}

  private real(): Promise<BrowserSession> {
    return this.sessions.get(this.config);
  }

  async ensureReady(): Promise<void> {
    await this.real();
  }

  async getToken(force?: boolean) {
    return (await this.real()).getToken(force);
  }

  async collectRawTokens() {
    return (await this.real()).collectRawTokens();
  }

  async currentUrl() {
    return (await this.real()).currentUrl();
  }

  async harvestWsTemplate() {
    return (await this.real()).harvestWsTemplate?.();
  }

  async harvestBearer() {
    return (await this.real()).harvestBearer?.();
  }

  async runStream(opts: Parameters<BrowserSession["runStream"]>[0]): Promise<void> {
    return (await this.real()).runStream(opts);
  }

  async dispose(): Promise<void> {
    /* 実体は SessionManager が管理するため何もしない */
  }
}

function showSessionError(e: unknown): void {
  log.error("signIn error", e as Error);
  if (e instanceof BrowserSessionError) {
    vscode.window
      .showErrorMessage(e.message, ...(e.guidance ? ["詳細"] : []))
      .then((sel) => {
        if (sel === "詳細" && e.guidance) {
          vscode.window.showInformationMessage(e.guidance);
        }
      });
    return;
  }
  vscode.window.showErrorMessage(`MS Copilot 接続エラー: ${e instanceof Error ? e.message : String(e)}`);
}
