import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import * as path from "node:path";
import { randomUUID } from "node:crypto";
import CDP from "chrome-remote-interface";
import type { MsCopilotConfig } from "../config";
import { log } from "../logger";
import {
  BrowserSession,
  BrowserSessionError,
  COLLECT_TOKENS_SCRIPT,
  READ_BEARER_SCRIPT,
  READ_WS_URL_SCRIPT,
  RawTokenCandidate,
  RunStreamOptions,
  SubstrateToken,
  WS_HOOK_SCRIPT,
  formatInventory,
  inventoryTokens,
  isTokenFresh,
  parseWsUrlToken,
  pickSydneyToken,
  tokenFromSecret,
} from "./browserSession";

/**
 * ユーザーの Edge/Chrome に CDP で相乗りする実装。
 *
 * 接続戦略:
 *  1) 既にデバッグポートが開いていれば接続(ユーザーがログイン済みプロファイルで
 *     --remote-debugging-port 起動している最良ケース)。
 *  2) 開いていなければ、設定のブラウザ/プロファイルでデバッグポート付き起動。
 *     userDataDir 未指定時は拡張専用プロファイル(初回のみ手動ログインが必要)。
 */
export class CdpBridge implements BrowserSession {
  private client: any | undefined;
  private launched = false;
  private readonly workProfile: string;

  constructor(
    private readonly config: MsCopilotConfig,
    private readonly storageDir: string,
  ) {
    this.workProfile = config.userDataDir || path.join(storageDir, "browser-profile");
  }

  async ensureReady(): Promise<void> {
    if (this.client) {
      return;
    }
    const port = this.config.debugPort;

    if (!(await isPortAlive(port))) {
      await this.launchBrowser(port);
      await waitForPort(port, 20000);
      // 起動直後は CDP.List が about:blank しか返さず、実行コンテキストも未準備なことがある。
      // 目的ページ(またはログインリダイレクト)が現れるまで待ってから接続する。
      await waitForPageReady(port, this.config.startUrl, 20000);
    }

    const target = await this.ensurePageTarget(port);
    const client = await CDP({ port, target: target.id ?? target });
    client.on("disconnect", () => log.warn("CDP: client disconnected"));
    try {
      await client.Page.enable();
      await client.Runtime.enable();
      this.client = client;
      // ページ読込前に WebSocket フックを仕込む(ユーザー実接続から token を採取するため)。
      // 既に読み込み済みのタブにも即時適用しておく。
      try {
        await client.Page.addScriptToEvaluateOnNewDocument({ source: WS_HOOK_SCRIPT });
      } catch {
        /* 一部の環境では未対応。即時 evaluate 側で代替する。 */
      }
      await this.evaluate(WS_HOOK_SCRIPT).catch(() => undefined);
      // 目的の URL でなければ遷移し、読み込みを待つ。
      await this.navigateIfNeeded(this.config.startUrl);
    } catch (e) {
      // 半端な状態を残さない。
      this.client = undefined;
      try {
        await client.close();
      } catch {
        /* ignore */
      }
      throw e;
    }
    log.info("CDP: page ready");
  }

  async getToken(force = false): Promise<SubstrateToken> {
    await this.ensureReady();
    const token = await this.acquireToken(force);
    if (token) {
      return token;
    }
    // 失敗時は「何が見つかったか」を診断ログへ残す(実機での原因切り分け用)。
    await this.logTokenInventory();
    throw new BrowserSessionError(
      "M365 Copilot の認証トークンを取得できませんでした。",
      `対象ブラウザで ${this.config.startUrl} を開き、M365 Copilot にサインイン済みか確認してください。` +
        " コマンド「MS Copilot: トークン取得を診断」で、見つかったトークン候補をログに出力できます。",
    );
  }

  async collectRawTokens(): Promise<RawTokenCandidate[]> {
    await this.ensureReady();
    return (await this.evaluate<RawTokenCandidate[]>(COLLECT_TOKENS_SCRIPT)) ?? [];
  }

  async currentUrl(): Promise<string> {
    if (!this.client) {
      return "";
    }
    return this.evaluate<string>("location.href").catch(() => "");
  }

  async harvestWsTemplate(): Promise<string | undefined> {
    await this.ensureReady();
    const url = await this.readHarvestedUrl();
    return url || undefined;
  }

  async harvestBearer(): Promise<string | undefined> {
    await this.ensureReady();
    const secret = await this.readHarvestedBearer();
    return secret || undefined;
  }

  private async readHarvestedUrl(): Promise<string> {
    if (!this.client) {
      return "";
    }
    return this.evaluate<string>(READ_WS_URL_SCRIPT).catch(() => "");
  }

  private async readHarvestedBearer(): Promise<string> {
    if (!this.client) {
      return "";
    }
    return this.evaluate<string>(READ_BEARER_SCRIPT).catch(() => "");
  }

  /**
   * トークン取得の本体。実テナントでは access_token が web ストレージに残らないため、
   * 「再読込 → 一定時間ポーリング」を繰り返しつつ 3 経路を試す。
   * 有効期限に余裕のあるトークンだけ採用する。
   *
   *  1) MSAL キャッシュ(local/sessionStorage)走査 → pickSydneyToken
   *  2) ページ substrate 宛 HTTP の Bearer 採取(WebSocket を待たずに拾える主経路)
   *  3) ページ実接続の WebSocket URL からの採取(実接続が張られていれば最も確実)
   */
  private async acquireToken(force: boolean): Promise<SubstrateToken | undefined> {
    const attempts = 4;
    for (let attempt = 0; attempt < attempts; attempt++) {
      if ((force && attempt === 0) || attempt > 0) {
        await this.reload();
      }
      // 読込直後は substrate への実トラフィックが立ち上がるまで間がある。
      // 初回は短め、再読込後はページ全体が動くまで長めに待つ。
      const budgetMs = attempt === 0 && !force ? 4000 : 12000;
      const token = await this.pollTokenSources(budgetMs);
      if (token) {
        return token;
      }
    }
    return undefined;
  }

  /** 制限時間内で 3 経路を繰り返し走査し、最初に得られた有効トークンを返す。 */
  private async pollTokenSources(budgetMs: number): Promise<SubstrateToken | undefined> {
    const deadline = Date.now() + budgetMs;
    do {
      const storage = pickSydneyToken(await this.collectRawTokens());
      if (storage && isTokenFresh(storage)) {
        logAcquired("storage", storage);
        return storage;
      }
      const bearer = tokenFromSecret(await this.readHarvestedBearer());
      if (bearer && isTokenFresh(bearer)) {
        logAcquired("http-bearer", bearer);
        return bearer;
      }
      const wsToken = parseWsUrlToken(await this.readHarvestedUrl());
      if (wsToken && isTokenFresh(wsToken)) {
        logAcquired("ws-harvest", wsToken);
        return wsToken;
      }
      await delay(750);
    } while (Date.now() < deadline);
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
    const client = this.client;
    const bindingName = `__mscopilot_${randomUUID().replace(/-/g, "")}`;

    const onBinding = (params: { name: string; payload: string }) => {
      if (params.name !== bindingName) {
        return;
      }
      try {
        opts.onFrame(JSON.parse(params.payload));
      } catch {
        /* 不正 JSON は無視 */
      }
    };

    await client.Runtime.addBinding({ name: bindingName });
    // chrome-remote-interface のイベント購読は解除関数を返す。
    const unsubscribe: () => void = client.Runtime.bindingCalled(onBinding);

    const abortListener = () => {
      // ページ内のアクティブ WS を閉じて実際にストリームを中断する。
      this.evaluate(
        "try { globalThis.__mscopilotWs && globalThis.__mscopilotWs.close(); } catch (e) {}",
      ).catch(() => undefined);
    };
    opts.signal?.addEventListener("abort", abortListener);

    try {
      const expr = `(${opts.script})(${JSON.stringify(opts.arg)}, (s) => globalThis["${bindingName}"](s))`;
      const res = await client.Runtime.evaluate({
        expression: expr,
        awaitPromise: true,
        returnByValue: true,
        userGesture: true,
      });
      if (res.exceptionDetails) {
        throw new Error(
          res.exceptionDetails.exception?.description ??
            res.exceptionDetails.text ??
            "page evaluation error",
        );
      }
    } finally {
      opts.signal?.removeEventListener("abort", abortListener);
      try {
        unsubscribe?.();
      } catch {
        /* ignore */
      }
      try {
        await client.Runtime.removeBinding({ name: bindingName });
      } catch {
        /* ignore */
      }
    }
  }

  async dispose(): Promise<void> {
    try {
      await this.client?.close();
    } catch {
      /* ignore */
    }
    this.client = undefined;
    // 起動したブラウザはユーザーが継続利用できるよう終了させない。
  }

  // ---- 内部ヘルパー -------------------------------------------------------

  private async evaluate<T>(expression: string): Promise<T> {
    const res = await this.client.Runtime.evaluate({
      expression,
      returnByValue: true,
      awaitPromise: true,
    });
    if (res.exceptionDetails) {
      throw new Error(
        res.exceptionDetails.exception?.description ?? res.exceptionDetails.text ?? "eval error",
      );
    }
    return res.result?.value as T;
  }

  private async reload(): Promise<void> {
    await this.client.Page.reload({ ignoreCache: false });
    await this.waitForLoad();
    await delay(1200);
  }

  private async navigateIfNeeded(url: string): Promise<void> {
    const current = await this.evaluate<string>("location.href").catch(() => "");
    // 空/about:blank のときだけ遷移する。既に Microsoft ドメイン(SSO リダイレクト中を含む)
    // にいる場合は遷移させない — ログインフローや別タブを壊さないため。
    if (!current || current === "about:blank" || current.startsWith("chrome://")) {
      await this.client.Page.navigate({ url });
      await this.waitForLoad();
    }
  }

  private waitForLoad(timeoutMs = 30000): Promise<void> {
    return new Promise<void>((resolve) => {
      let done = false;
      let unsubscribe: (() => void) | undefined;
      const finish = () => {
        if (done) {
          return;
        }
        done = true;
        try {
          unsubscribe?.();
        } catch {
          /* ignore */
        }
        resolve();
      };
      unsubscribe = this.client.Page.loadEventFired(finish);
      setTimeout(finish, timeoutMs);
    });
  }

  private async ensurePageTarget(port: number): Promise<any> {
    const targets: any[] = await CDP.List({ port });
    const pages = targets.filter((t) => t.type === "page");
    // startUrl と同じホストのタブだけ再利用する。別の Microsoft タブ(Outlook 等)は
    // 乗っ取らず、専用の新規タブを開く。
    const startHost = safeHost(this.config.startUrl);
    const preferred = pages.find((t) => startHost && safeHost(t.url ?? "") === startHost);
    if (preferred) {
      return preferred;
    }
    return await CDP.New({ port, url: this.config.startUrl });
  }

  private async launchBrowser(port: number): Promise<void> {
    const bin = this.config.browserPath || findBrowser();
    if (!bin) {
      throw new BrowserSessionError(
        "Edge/Chrome の実行ファイルが見つかりませんでした。",
        "設定 mscopilot.browserPath に Edge か Chrome のパスを指定してください。",
      );
    }
    const args = [
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${this.workProfile}`,
      "--no-first-run",
      "--no-default-browser-check",
      "--restore-last-session",
      this.config.startUrl,
    ];
    log.info(`launching browser: ${bin} (profile: ${this.workProfile})`);
    const child = spawn(bin, args, { detached: true, stdio: "ignore" });
    child.unref();
    this.launched = true;
  }
}

// ---- モジュールレベルユーティリティ ---------------------------------------

async function isPortAlive(port: number): Promise<boolean> {
  try {
    await CDP.Version({ port });
    return true;
  } catch {
    return false;
  }
}

async function waitForPort(port: number, timeoutMs: number): Promise<void> {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await isPortAlive(port)) {
      return;
    }
    await delay(500);
  }
  throw new BrowserSessionError(
    "ブラウザのデバッグポートに接続できませんでした。",
    "既にブラウザが起動している場合は一度完全に終了してから再実行するか、mscopilot.debugPort を確認してください。",
  );
}

/**
 * ブラウザ起動直後、接続に足るページが現れるまで待つ(ベストエフォート)。
 * 目的ホストのページが出れば最良。無くても、about:blank 以外へ遷移が始まっていれば
 * 良しとする(SSO で login.microsoftonline.com 等へリダイレクト中の初回ログインを許容)。
 * タイムアウトしても致命ではないため throw せず、後続の遷移処理に委ねる。
 */
async function waitForPageReady(port: number, expectedUrl: string, timeoutMs: number): Promise<void> {
  const expectedHost = safeHost(expectedUrl);
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const pages = (await CDP.List({ port })).filter((t: any) => t.type === "page");
      const onHost = pages.some((t: any) => expectedHost && safeHost(t.url ?? "") === expectedHost);
      const navigated = pages.some((t: any) => {
        const u = t.url ?? "";
        return u && u !== "about:blank" && !u.startsWith("chrome://");
      });
      if (onHost || navigated) {
        // 一覧登録直後は実行コンテキストが未安定なことがあるため少し待つ。
        await delay(750);
        return;
      }
    } catch {
      // 起動途中の一時的失敗は再試行。
    }
    await delay(250);
  }
  log.warn("CDP: Copilot ページの起動確認がタイムアウトしました(続行します)");
}

function safeHost(url: string): string | undefined {
  try {
    return new URL(url).host;
  } catch {
    return undefined;
  }
}

function findBrowser(): string | undefined {
  const candidates: string[] = [];
  if (process.platform === "win32") {
    const pf = process.env["ProgramFiles"] ?? "C:\\Program Files";
    const pfx86 = process.env["ProgramFiles(x86)"] ?? "C:\\Program Files (x86)";
    const local = process.env["LOCALAPPDATA"] ?? "";
    candidates.push(
      path.join(pfx86, "Microsoft", "Edge", "Application", "msedge.exe"),
      path.join(pf, "Microsoft", "Edge", "Application", "msedge.exe"),
      path.join(pf, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(pfx86, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(local, "Google", "Chrome", "Application", "chrome.exe"),
    );
  } else if (process.platform === "darwin") {
    candidates.push(
      "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    );
  } else {
    candidates.push(
      "/usr/bin/microsoft-edge",
      "/usr/bin/microsoft-edge-stable",
      "/usr/bin/google-chrome",
      "/usr/bin/google-chrome-stable",
      "/usr/bin/chromium",
      "/usr/bin/chromium-browser",
    );
  }
  return candidates.find((p) => p && existsSync(p));
}

function logAcquired(via: string, token: SubstrateToken): void {
  log.info(
    `token acquired (${via}): oid=${token.objectId.slice(0, 8)}… tid=${token.tenantId.slice(0, 8)}… ` +
      `exp=${new Date(token.expiresAt).toISOString()}`,
  );
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
