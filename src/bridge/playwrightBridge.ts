import * as path from "node:path";
import { randomUUID } from "node:crypto";
import type { MsCopilotConfig } from "../config";
import { log } from "../logger";
import {
  BrowserSession,
  BrowserSessionError,
  COLLECT_TOKENS_SCRIPT,
  HarvestedBearer,
  READ_BEARER_SCRIPT,
  READ_BEARERS_SCRIPT,
  READ_WS_URL_SCRIPT,
  RawTokenCandidate,
  RunStreamOptions,
  SubstrateToken,
  WS_HOOK_SCRIPT,
  formatInventory,
  formatTokenScope,
  inventoryTokens,
  isTokenFresh,
  parseWsUrlToken,
  pickBearerToken,
  pickSydneyToken,
} from "./browserSession";

/**
 * Playwright 永続コンテキストで専用ブラウザを起動するフォールバック。
 * 企業ポリシーで Edge のリモートデバッグが禁止されている場合などに使用する。
 * playwright は動的 require(未インストールなら分かりやすくエラー)。
 */
export class PlaywrightBridge implements BrowserSession {
  private context: any | undefined;
  private page: any | undefined;
  private readonly streams = new Map<string, (frame: unknown) => void>();
  private readonly profileDir: string;

  constructor(
    private readonly config: MsCopilotConfig,
    storageDir: string,
  ) {
    this.profileDir = config.userDataDir || path.join(storageDir, "pw-profile");
  }

  async ensureReady(): Promise<void> {
    if (this.page && !this.page.isClosed()) {
      return;
    }
    // 前回のセッションが閉じている/壊れている場合は、永続プロファイルのロックを解放するため
    // 一度コンテキストを畳んでから作り直す(開いたままでは同じプロファイルで再起動できない)。
    if (this.context) {
      await this.context.close().catch(() => undefined);
      this.context = undefined;
      this.page = undefined;
    }
    const chromium = loadChromium();
    const launchOpts: Record<string, unknown> = {
      headless: false,
      args: ["--no-first-run", "--no-default-browser-check"],
    };
    if (this.config.browserPath) {
      launchOpts.executablePath = this.config.browserPath;
    } else {
      launchOpts.channel = "msedge"; // SSO のため既定で Edge を優先
    }

    try {
      this.context = await chromium.launchPersistentContext(this.profileDir, launchOpts);
    } catch (e) {
      log.warn("msedge channel 起動に失敗、既定 Chromium で再試行", e as Error);
      delete launchOpts.channel;
      this.context = await chromium.launchPersistentContext(this.profileDir, launchOpts);
    }

    // ユーザーがブラウザ/タブを閉じた・クラッシュした場合、死んだ参照を掴み続けない。
    // クリアしておけば次の ensureReady がクリーンに作り直す。
    this.context.on("close", () => {
      this.context = undefined;
      this.page = undefined;
    });

    this.page = this.context.pages()[0] ?? (await this.context.newPage());
    this.page.on("close", () => {
      this.page = undefined;
    });
    this.page.on("crash", () => {
      this.page = undefined;
    });

    await this.page.exposeFunction("__mscopilotPush", (id: string, s: string) => {
      const handler = this.streams.get(id);
      if (!handler) {
        return;
      }
      try {
        handler(JSON.parse(s));
      } catch {
        /* ignore */
      }
    });

    // ページ読込前に WebSocket フックを仕込む(ユーザー実接続から token を採取するため)。
    try {
      await this.context.addInitScript({ content: WS_HOOK_SCRIPT });
    } catch (e) {
      log.warn("addInitScript(WS hook) に失敗", e as Error);
    }

    await this.page.goto(this.config.startUrl, { waitUntil: "domcontentloaded" });
    log.info("Playwright: page ready");
  }

  async getToken(force = false): Promise<SubstrateToken> {
    await this.ensureReady();
    const token = await this.acquireToken(force);
    if (token) {
      return token;
    }
    await this.logTokenInventory();
    throw new BrowserSessionError(
      "M365 Copilot の認証トークンを取得できませんでした。",
      `起動したブラウザで ${this.config.startUrl} を開き、M365 Copilot にサインイン済みか確認してください。` +
        " コマンド「MS Copilot: トークン取得を診断」で、見つかったトークン候補をログに出力できます。",
    );
  }

  async collectRawTokens(): Promise<RawTokenCandidate[]> {
    await this.ensureReady();
    return ((await this.page.evaluate(COLLECT_TOKENS_SCRIPT)) as RawTokenCandidate[]) ?? [];
  }

  async currentUrl(): Promise<string> {
    try {
      return this.page ? String(this.page.url()) : "";
    } catch {
      return "";
    }
  }

  async harvestWsTemplate(): Promise<string | undefined> {
    await this.ensureReady();
    // 非ブロッキングで読むだけ。実接続 URL は force トークン再取得時の reload や signIn で採取される。
    // ここで能動的に再読込すると発信が数十秒ブロックするため行わない。
    const url = await this.readHarvestedUrl();
    return url || undefined;
  }

  async harvestBearer(): Promise<string | undefined> {
    await this.ensureReady();
    const secret = await this.readHarvestedBearer();
    return secret || undefined;
  }

  private async readHarvestedUrl(): Promise<string> {
    try {
      return this.page ? String((await this.page.evaluate(READ_WS_URL_SCRIPT)) || "") : "";
    } catch {
      return "";
    }
  }

  private async readHarvestedBearer(): Promise<string> {
    try {
      return this.page ? String((await this.page.evaluate(READ_BEARER_SCRIPT)) || "") : "";
    } catch {
      return "";
    }
  }

  /** 採取済みの substrate 宛 Bearer 群(URL付き)を読む。選別は pickBearerToken で行う。 */
  private async readHarvestedBearers(): Promise<HarvestedBearer[]> {
    try {
      if (!this.page) {
        return [];
      }
      const json = String((await this.page.evaluate(READ_BEARERS_SCRIPT)) || "[]");
      const list = JSON.parse(json);
      return Array.isArray(list) ? (list as HarvestedBearer[]) : [];
    } catch {
      return [];
    }
  }

  /** 実テナントでは access_token が永続化されないため、3 経路を制限時間内で走査する。 */
  private async acquireToken(force: boolean): Promise<SubstrateToken | undefined> {
    const attempts = 4;
    for (let attempt = 0; attempt < attempts; attempt++) {
      if ((force && attempt === 0) || attempt > 0) {
        await this.page.reload({ waitUntil: "domcontentloaded" });
      }
      const budgetMs = attempt === 0 && !force ? 4000 : 12000;
      const deadline = Date.now() + budgetMs;
      do {
        const storage = pickSydneyToken(await this.collectRawTokens());
        if (storage && isTokenFresh(storage)) {
          log.info(
            `token acquired (pw storage): exp=${new Date(storage.expiresAt).toISOString()} ${formatTokenScope(storage)}`,
          );
          return storage;
        }
        const bearer = pickBearerToken(await this.readHarvestedBearers());
        if (bearer && isTokenFresh(bearer)) {
          log.info(
            `token acquired (pw http-bearer): exp=${new Date(bearer.expiresAt).toISOString()} ${formatTokenScope(bearer)}`,
          );
          return bearer;
        }
        const wsToken = parseWsUrlToken(await this.readHarvestedUrl());
        if (wsToken && isTokenFresh(wsToken)) {
          log.info(
            `token acquired (pw ws-harvest): exp=${new Date(wsToken.expiresAt).toISOString()} ${formatTokenScope(wsToken)}`,
          );
          return wsToken;
        }
        await delay(750);
      } while (Date.now() < deadline);
    }
    return undefined;
  }

  private async logTokenInventory(): Promise<void> {
    try {
      const href = await this.currentUrl();
      const inv = inventoryTokens(await this.collectRawTokens());
      log.warn(`token 取得失敗。現在のページ: ${href || "(不明)"}`);
      log.warn(`検出したトークン候補 ${inv.length} 件:\n${formatInventory(inv)}`);
    } catch (e) {
      log.warn("token inventory の収集に失敗", e as Error);
    }
  }

  async runStream(opts: RunStreamOptions): Promise<void> {
    await this.ensureReady();
    const id = randomUUID();
    this.streams.set(id, opts.onFrame);

    const abortListener = () => {
      this.page
        .evaluate("try { globalThis.__mscopilotWs && globalThis.__mscopilotWs.close(); } catch (e) {}")
        .catch(() => undefined);
    };
    opts.signal?.addEventListener("abort", abortListener);

    try {
      await this.page.evaluate(
        ({ script, arg, id }: { script: string; arg: unknown; id: string }) => {
          // eslint-disable-next-line no-eval
          const fn = (0, eval)(script);
          return fn(arg, (s: string) => (globalThis as any).__mscopilotPush(id, s));
        },
        { script: opts.script, arg: opts.arg, id },
      );
    } finally {
      opts.signal?.removeEventListener("abort", abortListener);
      this.streams.delete(id);
    }
  }

  async dispose(): Promise<void> {
    try {
      await this.context?.close();
    } catch {
      /* ignore */
    }
    this.context = undefined;
    this.page = undefined;
  }
}

function loadChromium(): any {
  try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const pw = require("playwright");
    return pw.chromium;
  } catch {
    throw new BrowserSessionError(
      "playwright が見つかりません。",
      "playwright を使う場合は拡張のディレクトリで `npm i playwright` を実行し、`npx playwright install msedge` 等でブラウザを用意してください。または設定 mscopilot.transport を cdp に戻してください。",
    );
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
