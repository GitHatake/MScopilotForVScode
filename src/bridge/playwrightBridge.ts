import * as path from "node:path";
import { randomUUID } from "node:crypto";
import type { MsCopilotConfig } from "../config";
import { log } from "../logger";
import {
  BrowserSession,
  BrowserSessionError,
  COLLECT_TOKENS_SCRIPT,
  RunStreamOptions,
  SubstrateToken,
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
    if (this.page) {
      return;
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

    this.page = this.context.pages()[0] ?? (await this.context.newPage());

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

    await this.page.goto(this.config.startUrl, { waitUntil: "domcontentloaded" });
    log.info("Playwright: page ready");
  }

  async getToken(force = false): Promise<SubstrateToken> {
    await this.ensureReady();
    for (let attempt = 0; attempt < 3; attempt++) {
      if ((force && attempt === 0) || attempt > 0) {
        await this.page.reload({ waitUntil: "domcontentloaded" });
      }
      const candidates: string[] = await this.page.evaluate(COLLECT_TOKENS_SCRIPT);
      const token = pickSydneyToken(candidates ?? []);
      if (token) {
        log.info(`token acquired (pw): exp=${new Date(token.expiresAt).toISOString()}`);
        return token;
      }
      await delay(1200);
    }
    throw new BrowserSessionError(
      "M365 Copilot の認証トークンを取得できませんでした。",
      `起動したブラウザで ${this.config.startUrl} を開き、M365 Copilot にサインインしてから再実行してください。`,
    );
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
