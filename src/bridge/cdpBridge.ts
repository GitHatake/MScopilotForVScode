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
  RawTokenCandidate,
  RunStreamOptions,
  SubstrateToken,
  formatInventory,
  inventoryTokens,
  isTokenFresh,
  pickSydneyToken,
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
    }

    const target = await this.ensurePageTarget(port);
    const client = await CDP({ port, target: target.id ?? target });
    try {
      await client.Page.enable();
      await client.Runtime.enable();
      this.client = client;
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

  /**
   * トークン取得の本体。MSAL の silent 取得が完了するまで時間がかかることがあるため、
   * 「再読込 → 数回ポーリング」を複数回繰り返す。有効期限に余裕のあるトークンだけ採用する。
   */
  private async acquireToken(force: boolean): Promise<SubstrateToken | undefined> {
    const attempts = 4;
    for (let attempt = 0; attempt < attempts; attempt++) {
      if ((force && attempt === 0) || attempt > 0) {
        await this.reload();
      }
      // 1 回の読み込み後も MSAL 反映まで間があるので、数回に分けて走査する。
      for (let poll = 0; poll < 3; poll++) {
        const candidates = await this.collectRawTokens();
        const token = pickSydneyToken(candidates);
        if (token && isTokenFresh(token)) {
          log.info(
            `token acquired: oid=${token.objectId.slice(0, 8)}… tid=${token.tenantId.slice(0, 8)}… exp=${new Date(token.expiresAt).toISOString()}`,
          );
          return token;
        }
        await delay(1000);
      }
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

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
