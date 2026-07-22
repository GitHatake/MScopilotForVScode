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
  HarvestedBearer,
  PRIME_COPILOT_GUIDANCE,
  READ_BEARER_SCRIPT,
  READ_BEARERS_SCRIPT,
  READ_NET_SCRIPT,
  READ_WS_URL_SCRIPT,
  READ_WS_URLS_SCRIPT,
  RawTokenCandidate,
  RunStreamOptions,
  SubstrateToken,
  WS_HOOK_SCRIPT,
  describeBearers,
  formatInventory,
  formatTokenScope,
  inventoryTokens,
  isTokenFresh,
  parseWsUrlToken,
  pickBearerTokenScored,
  pickSydneyToken,
  scoreToken,
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
  private readonly workProfile: string;
  /**
   * ユーザーが Web の Copilot で送信した時に捕捉した実接続テンプレート(access_token 付き URL)と
   * そこから導いた Copilot 用トークン。ページ再読込で page 側の globalThis は消えるため Node 側で
   * 保持し、失効するまで再利用する(=一度プライムすれば以後は自動で使える)。
   */
  private copilotTemplate: string | undefined;
  private copilotToken: SubstrateToken | undefined;

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
    client.on("disconnect", () => {
      log.warn("CDP: client disconnected");
      // 切断(SSO リダイレクトでのタブ再生成やユーザー操作)を検知したら、
      // 死んだクライアントを掴み続けない。次の操作で ensureReady が再接続する。
      // 既に別クライアントへ張り替え済みなら触らない。
      if (this.client === client) {
        this.client = undefined;
      }
    });
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
    const token = await this.withReconnect(() => this.acquireToken(force));
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
    return this.withReconnect(
      async () => (await this.evaluate<RawTokenCandidate[]>(COLLECT_TOKENS_SCRIPT)) ?? [],
    );
  }

  async currentUrl(): Promise<string> {
    if (!this.client) {
      return "";
    }
    return this.evaluate<string>("location.href").catch(() => "");
  }

  async harvestWsTemplate(): Promise<string | undefined> {
    await this.ensureReady();
    // 保持済みの実接続テンプレートが失効前ならそれを最優先で使う(再読込を跨いでも生き残る)。
    if (this.copilotToken && isTokenFresh(this.copilotToken) && this.copilotTemplate) {
      return this.copilotTemplate;
    }
    // 未保持ならページから採取を試みる(非ブロッキング。能動再読込はしない)。
    await this.captureCopilot();
    if (this.copilotTemplate) {
      return this.copilotTemplate;
    }
    const url = await this.readHarvestedUrl();
    return url || undefined;
  }

  /** 捕捉済みの Chathub 実接続 URL 群(新しい順)を読む。 */
  private async readHarvestedUrls(): Promise<string[]> {
    if (!this.client) {
      return [];
    }
    const json = await this.evaluate<string>(READ_WS_URLS_SCRIPT).catch(() => "[]");
    try {
      const arr = JSON.parse(json || "[]");
      return Array.isArray(arr) ? (arr as string[]) : [];
    } catch {
      return [];
    }
  }

  /**
   * ページが張った Copilot(Chathub)実接続 URL からトークン+テンプレートを採取し Node 側に保持する。
   * 見つかれば true。これがユーザーの「ブラウザで一度送信」で Copilot トークンを得る唯一の経路。
   */
  private async captureCopilot(): Promise<boolean> {
    if (this.copilotToken && isTokenFresh(this.copilotToken)) {
      return true;
    }
    const urls = await this.readHarvestedUrls();
    const single = await this.readHarvestedUrl();
    for (const url of [...urls, single]) {
      const token = parseWsUrlToken(url);
      if (token && isTokenFresh(token)) {
        this.copilotTemplate = url;
        this.copilotToken = token;
        log.info(
          `Copilot 実接続を捕捉しました(以後再利用): ${formatTokenScope(token, url)} exp=${new Date(token.expiresAt).toISOString()}`,
        );
        return true;
      }
    }
    return false;
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

  /** 採取済みの substrate 宛 Bearer 群(URL付き)を読む。選別は pickBearerToken で行う。 */
  private async readHarvestedBearers(): Promise<HarvestedBearer[]> {
    if (!this.client) {
      return [];
    }
    const json = await this.evaluate<string>(READ_BEARERS_SCRIPT).catch(() => "[]");
    try {
      const list = JSON.parse(json || "[]");
      return Array.isArray(list) ? (list as HarvestedBearer[]) : [];
    } catch {
      return [];
    }
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
    // プライム済み(過去にユーザーが Web で送信して捕捉した)Copilot トークンが失効前なら、
    // 再読込せずに最優先で使う。force 再取得でも再読込するとプライムが消えるため、これを先に返す。
    if (await this.captureCopilot()) {
      log.info(`token acquired (copilot-primed): ${formatTokenScope(this.copilotToken!, this.copilotTemplate)}`);
      return this.copilotToken;
    }

    const attempts = 4;
    let fallback: SubstrateToken | undefined; // Copilot 未採取時のみ使う substrate 一般トークン
    for (let attempt = 0; attempt < attempts; attempt++) {
      if ((force && attempt === 0) || attempt > 0) {
        log.info(`token: attempt ${attempt + 1}/${attempts} — ページを再読込します (force=${force})`);
        await this.reload();
      }
      // 読込直後は substrate への実トラフィックが立ち上がるまで間がある。
      // 初回は短め、再読込後はページ全体が動くまで長めに待つ。
      const budgetMs = attempt === 0 && !force ? 4000 : 12000;
      const found = await this.pollTokenSources(budgetMs);
      // 採取状況を毎回スナップショット(Copilot 用トークンが採れているか解析できるように)。
      await this.logHarvestSnapshot(`attempt ${attempt + 1}${found ? " ✓" : ""}`);
      if (found?.isCopilot) {
        return found.token; // Copilot(score>=3)確定 → 即採用
      }
      if (found && !fallback) {
        fallback = found.token; // substrate 一般(/search 等)。より良い物が出るまで保留。
      }
      // fallback を得たら以降の再読込は打ち切る(Copilot トークンはユーザーが Web で送信した時
      // にしか出ず、再読込では出ない)。速やかに送信を試し、失敗時は hint で操作を促す。
      if (fallback) {
        break;
      }
    }
    // Copilot 用トークンは採れなかった。fallback(あれば)を返しつつ、行動指示をログに出す。
    if (fallback) {
      log.warn(
        `token: Copilot 用トークンを採取できませんでした。substrate 一般トークンで送信を試みます` +
          `(スコープ不一致だと "Language model unavailable" になります)。`,
      );
      log.warn(`token: ${PRIME_COPILOT_GUIDANCE}`);
    }
    return fallback;
  }

  /**
   * 採取状況を診断ログへ大量に出す。どの経路で・どのスコープのトークンが採れているか、
   * ページがどんな substrate/認証通信・WebSocket を張ったかを可視化する。
   */
  private async logHarvestSnapshot(tag: string): Promise<void> {
    try {
      const bearers = await this.readHarvestedBearers();
      const wsUrl = await this.readHarvestedUrl();
      const storage = (await this.collectRawTokens()).length;
      log.info(`── harvest snapshot [${tag}] ──`);
      log.info(`  MSAL storage 候補: ${storage} 件 / 実接続WS: ${wsUrl ? "捕捉済み" : "未捕捉"}`);
      log.info(`  Bearer 候補 ${bearers.length} 件(★=Copilot想定 score3 / ○=substrate score2):\n${describeBearers(bearers)}`);
      const net = await this.readNetTrace();
      if (net.length) {
        log.info(`  ネットワークトレース ${net.length} 件:\n${net.map((s) => "    " + s).join("\n")}`);
      }
    } catch (e) {
      log.warn("harvest snapshot 失敗", e as Error);
    }
  }

  private async readNetTrace(): Promise<string[]> {
    if (!this.client) {
      return [];
    }
    const json = await this.evaluate<string>(READ_NET_SCRIPT).catch(() => "[]");
    try {
      const arr = JSON.parse(json || "[]");
      return Array.isArray(arr) ? (arr as string[]) : [];
    } catch {
      return [];
    }
  }

  /**
   * 制限時間内で 3 経路を繰り返し走査する。Copilot(score>=3)が出れば即返す。出なければ
   * substrate 一般(score<=2)を「fallback」として保持し、時間切れ時に返す(即断で /search を
   * 掴んで送信失敗するのを避けつつ、Copilot 採取を待つ)。
   */
  private async pollTokenSources(
    budgetMs: number,
  ): Promise<{ token: SubstrateToken; isCopilot: boolean } | undefined> {
    const deadline = Date.now() + budgetMs;
    let fallback: { token: SubstrateToken; isCopilot: boolean } | undefined;
    do {
      // 1) 実接続 WebSocket(Copilot 確定)。最優先。
      if (await this.captureCopilot()) {
        logAcquired("ws-harvest", this.copilotToken!, this.copilotTemplate);
        return { token: this.copilotToken!, isCopilot: true };
      }
      // 2) HTTP/WS-URL の Bearer 群。採取元 URL 込みでスコアリングし Copilot を優先。
      const scored = pickBearerTokenScored(await this.readHarvestedBearers());
      if (scored && isTokenFresh(scored.token)) {
        const via = scored.score >= 3 ? "http-bearer(copilot)" : "http-bearer";
        logAcquired(via, scored.token, scored.url);
        if (scored.score >= 3) {
          return { token: scored.token, isCopilot: true };
        }
        fallback = fallback ?? { token: scored.token, isCopilot: false };
      }
      // 3) MSAL ストレージ。
      const storage = pickSydneyToken(await this.collectRawTokens());
      if (storage && isTokenFresh(storage)) {
        const score = scoreToken(storage);
        logAcquired(score >= 3 ? "storage(copilot)" : "storage", storage);
        if (score >= 3) {
          return { token: storage, isCopilot: true };
        }
        fallback = fallback ?? { token: storage, isCopilot: false };
      }
      await delay(750);
    } while (Date.now() < deadline);
    return fallback;
  }

  private async logTokenInventory(): Promise<void> {
    try {
      const href = await this.currentUrl();
      const inv = inventoryTokens(await this.collectRawTokens());
      log.warn(`token 取得失敗。現在のページ: ${href || "(不明)"}`);
      log.warn(`MSAL 候補 ${inv.length} 件:\n${formatInventory(inv)}`);
      const bearers = await this.readHarvestedBearers();
      log.warn(`HTTP/WS Bearer 候補 ${bearers.length} 件:\n${describeBearers(bearers)}`);
      log.warn(PRIME_COPILOT_GUIDANCE);
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

  /**
   * 操作中に CDP 接続が切れた場合(SSO リダイレクトでタブが再生成される等)、
   * 一度だけ再接続してから操作をやり直す。二度目も切断されたら諦めて送出する。
   * 接続断以外のエラーはそのまま送出する(症状を握り潰さない)。
   */
  private async withReconnect<T>(op: () => Promise<T>): Promise<T> {
    try {
      return await op();
    } catch (e) {
      if (!isConnectionClosed(e)) {
        throw e;
      }
      log.warn("CDP: 接続が切れたため再接続して再試行します");
      this.client = undefined;
      await this.ensureReady();
      return await op();
    }
  }

  private async evaluate<T>(expression: string): Promise<T> {
    // 切断で client が失われていれば再接続する(未定義参照ではなく接続断として扱う)。
    if (!this.client) {
      await this.ensureReady();
    }
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
    if (!this.client) {
      await this.ensureReady();
    }
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
    const client = this.client;
    return new Promise<void>((resolve) => {
      let done = false;
      let unsubscribe: (() => void) | undefined;
      let poll: ReturnType<typeof setInterval> | undefined;
      let timer: ReturnType<typeof setTimeout> | undefined;
      const finish = () => {
        if (done) {
          return;
        }
        done = true;
        if (poll) {
          clearInterval(poll);
        }
        if (timer) {
          clearTimeout(timer);
        }
        try {
          unsubscribe?.();
        } catch {
          /* ignore */
        }
        resolve();
      };
      unsubscribe = client.Page.loadEventFired(finish);
      // 接続断(タブ再生成)では loadEventFired が来ないため、client の張り替えを検知して
      // 満了を待たず早期に抜ける。再接続は呼び出し側(withReconnect)に委ねる。
      poll = setInterval(() => {
        if (this.client !== client) {
          finish();
        }
      }, 250);
      timer = setTimeout(finish, timeoutMs);
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

/**
 * CDP の WebSocket が閉じている/対象が破棄されたことを示すエラーかどうか。
 * 例: "WebSocket is not open: readyState 3 (CLOSED)" / "Target closed"。
 */
function isConnectionClosed(e: unknown): boolean {
  const msg = e instanceof Error ? e.message : String(e ?? "");
  return /WebSocket is not open|readyState 3|Target.*closed|not connected|Session closed|Inspected target/i.test(
    msg,
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

function logAcquired(via: string, token: SubstrateToken, target?: string): void {
  log.info(
    `token acquired (${via}): oid=${token.objectId.slice(0, 8)}… tid=${token.tenantId.slice(0, 8)}… ` +
      `exp=${new Date(token.expiresAt).toISOString()} ${formatTokenScope(token, target)}`,
  );
}

function delay(ms: number): Promise<void> {
  return new Promise((r) => setTimeout(r, ms));
}
