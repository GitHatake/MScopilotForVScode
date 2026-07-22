/**
 * ブラウザ接続の共通インターフェース。
 * CDP(ユーザーの Edge に相乗り)と Playwright(専用ブラウザ)の実装差異を吸収する。
 *
 * 設計方針(計画より):
 *  - 認証は「ブラウザにログイン済み」の M365 セッションから sydney トークンを読み取る。
 *  - チャット発信は可能な限り「ページ内」で行い、実 UA/TLS/セッションと一致させる
 *    (runStream)。フレームはページ→Node へ push 関数(binding)経由で返す。
 */

export interface SubstrateToken {
  /** substrate 用アクセストークン(JWT 本体) */
  accessToken: string;
  /** ユーザーの Entra Object ID (JWT の oid) */
  objectId: string;
  /** テナント ID (JWT の tid) */
  tenantId: string;
  /** 失効時刻 (epoch ms) */
  expiresAt: number;
}

export interface RunStreamOptions {
  /**
   * ページ内で実行する関数式。`(arg, push) => Promise<void> | void` の形。
   * `push(frameJsonString)` を呼ぶとフレームが Node 側 onFrame に渡る。
   */
  script: string;
  /** script の第一引数に渡す JSON 直列化可能な値 */
  arg: unknown;
  /** ページから push された各フレーム(JSON.parse 済み) */
  onFrame: (frame: unknown) => void;
  /** 中断用シグナル */
  signal?: AbortSignal;
}

/**
 * ページの MSAL キャッシュから収集した生のトークン候補。
 * secret(JWT)に加え、選別・診断に使う MSAL メタデータを持つ。
 */
export interface RawTokenCandidate {
  /** JWT 本体 */
  secret: string;
  /** MSAL の target(要求スコープ、空白区切り)。resource を最も確実に示す。 */
  target?: string;
  /** トークンを取得したアプリの clientId */
  clientId?: string;
  /** テナント(realm) */
  realm?: string;
  /** 認可サーバーの環境(login.microsoftonline.com 等) */
  environment?: string;
  /** MSAL の credentialType(AccessToken / AccessToken_With_AuthScheme / IdToken 等) */
  credentialType?: string;
  /** 収集元ストレージ(localStorage / sessionStorage) */
  store?: string;
}

export interface BrowserSession {
  /** ブラウザへ接続し、M365 ページを開いてサインイン状態を整える。 */
  ensureReady(): Promise<void>;

  /**
   * ログイン済みセッションから substrate(audience: .../sydney)トークンを取得する。
   * force=true で再読込して最新トークンを取り直す。
   */
  getToken(force?: boolean): Promise<SubstrateToken>;

  /**
   * ページの MSAL キャッシュにある全トークン候補を収集する(選別前の生データ)。
   * トークン取得診断コマンド用。
   */
  collectRawTokens(): Promise<RawTokenCandidate[]>;

  /** 現在ブラウザで開いている URL(診断表示用)。取得できなければ空文字。 */
  currentUrl(): Promise<string>;

  /**
   * ページの実接続(ユーザー自身の substrate WebSocket)から捕捉した URL テンプレートを返す。
   * access_token / variants / scenario など実テナントの実値を含む。捕捉していなければ undefined。
   * 実装は window.WebSocket をフックして記録する(WS_HOOK_SCRIPT)。
   */
  harvestWsTemplate?(): Promise<string | undefined>;

  /**
   * ページの substrate 宛 HTTP から捕捉した Bearer トークン(JWT 本体)を返す。
   * チャット WebSocket が開く前でも採取できる主経路。捕捉していなければ undefined。
   */
  harvestBearer?(): Promise<string | undefined>;

  /** ページ内でストリーミング処理を実行する。script 完了で resolve。 */
  runStream(opts: RunStreamOptions): Promise<void>;

  dispose(): Promise<void>;
}

/** 認証・接続に関する分かりやすいエラー。UI にガイダンスを表示するために型を分ける。 */
export class BrowserSessionError extends Error {
  constructor(
    message: string,
    readonly guidance?: string,
  ) {
    super(message);
    this.name = "BrowserSessionError";
  }
}

/**
 * チャット(sydney/BizChat)用トークンを強く示すヒント。target(スコープ)/ aud / scp の
 * いずれかに含まれれば最優先で採用する。実テナントでは aud が GUID のことも多く、
 * audience 単独では判別できないため、スコープ側の語も併せて照合する。
 */
const CHAT_HINTS = [
  "sydney",
  "m365chat",
  "m365 chat",
  "chathub",
  "m365copilot",
  "m365 copilot",
  "bizchat",
  "officeweb",
];

/** substrate リソースを示すヒント。sydney が見つからない場合のフォールバック用。 */
const SUBSTRATE_HINTS = [
  "substrate.office.com",
  "substrate.svc.cloud.microsoft",
  "substrate",
];

/** 有効期限にこれだけ余裕が無いトークンは「実質失効」とみなす(ミリ秒)。 */
const EXPIRY_MARGIN_MS = 60_000;

interface JwtPayload {
  aud?: string;
  oid?: string;
  tid?: string;
  exp?: number;
  scp?: string;
  appid?: string;
}

function decodeJwt(token: string): JwtPayload | undefined {
  const parts = token.split(".");
  if (parts.length < 2) {
    return undefined;
  }
  try {
    const json = Buffer.from(parts[1], "base64url").toString("utf8");
    return JSON.parse(json) as JwtPayload;
  } catch {
    return undefined;
  }
}

function toCandidate(c: RawTokenCandidate | string): RawTokenCandidate {
  return typeof c === "string" ? { secret: c } : c;
}

/**
 * 候補の「照合対象テキスト」を組み立てる。MSAL の target(スコープ)を最重視しつつ、
 * JWT の aud / scp も含める。いずれか一つでも substrate/sydney を示せば拾える。
 */
function matchText(cand: RawTokenCandidate, payload: JwtPayload | undefined): string {
  return [cand.target, cand.environment, payload?.aud, payload?.scp]
    .filter((x): x is string => typeof x === "string" && x.length > 0)
    .join(" ")
    .toLowerCase();
}

/** 3=チャット用トークン確度大、2=substrate リソース、0=対象外。 */
function scoreText(text: string): number {
  if (CHAT_HINTS.some((h) => text.includes(h))) {
    return 3;
  }
  if (SUBSTRATE_HINTS.some((h) => text.includes(h))) {
    return 2;
  }
  return 0;
}

interface Scored {
  token: SubstrateToken;
  score: number;
  valid: boolean;
}

function betterThan(cand: Scored, cur: Scored | undefined): boolean {
  if (!cur) {
    return true;
  }
  // 有効(期限に余裕あり)を最優先。失効済みは実際に使えないため。
  if (cand.valid !== cur.valid) {
    return cand.valid;
  }
  // 次にスコア(sydney > substrate)。
  if (cand.score !== cur.score) {
    return cand.score > cur.score;
  }
  // 最後に有効期限が新しい方。
  return cand.token.expiresAt > cur.token.expiresAt;
}

/**
 * ブラウザから収集したトークン候補群から、substrate(sydney)向けの最良トークンを 1 つ選ぶ。
 *
 * 選別基準(優先度順):
 *  1. 期限に余裕がある(実際に使える)
 *  2. スコア(target/aud/scp に sydney 等のチャットヒント > substrate ホスト)
 *  3. 有効期限が新しい
 *
 * 文字列(secret のみ)でも RawTokenCandidate でも受け付ける。
 */
export function pickSydneyToken(
  candidates: Array<RawTokenCandidate | string>,
): SubstrateToken | undefined {
  const now = Date.now();
  let best: Scored | undefined;
  for (const raw of candidates) {
    const cand = toCandidate(raw);
    if (!cand?.secret) {
      continue;
    }
    const payload = decodeJwt(cand.secret);
    // URL 構築に oid/tid/exp が必須。
    if (!payload?.oid || !payload.tid || !payload.exp) {
      continue;
    }
    const score = scoreText(matchText(cand, payload));
    if (score === 0) {
      continue;
    }
    const token: SubstrateToken = {
      accessToken: cand.secret,
      objectId: payload.oid,
      tenantId: payload.tid,
      expiresAt: payload.exp * 1000,
    };
    const scored: Scored = {
      token,
      score,
      valid: token.expiresAt - now > EXPIRY_MARGIN_MS,
    };
    if (betterThan(scored, best)) {
      best = scored;
    }
  }
  return best?.token;
}

/** トークンに有効期限の余裕があるか(取得ループの再試行判定に使う)。 */
export function isTokenFresh(token: SubstrateToken): boolean {
  return token.expiresAt - Date.now() > EXPIRY_MARGIN_MS;
}

/**
 * ページ内で実行し、実テナントの substrate トークンを複数経路で採取するフック。
 * 実テナントでは access_token が web ストレージに永続化されないため、ページの実トラフィックから拾う。
 *
 *  1) window.WebSocket: substrate 実接続 URL(access_token 付き)を __mscopilotWsUrl に記録。
 *  2) fetch / XMLHttpRequest の **リクエスト** Authorization: Bearer を採取(__mscopilotBearers)。
 *  3) fetch / XMLHttpRequest の **レスポンス本文** に含まれる JWT を採取(__mscopilotBearers)。
 *     実測(feedback_5)で、HTTP ヘッダに乗る substrate トークンは /search 用のみで、Copilot
 *     (Chathub)用トークンは WebSocket URL のクエリにしか現れないと判明した。Copilot 用トークンは
 *     MSAL のトークン取得(login.microsoftonline 等への通信)の**レスポンス本文**を経由するため、
 *     WebSocket が開く前でも本文走査で拾える可能性がある。スコープ選別は Node 側(pickBearerToken)。
 *
 * さらに、解析用に substrate/認証系の通信と WebSocket を __mscopilotNet に**大量にトレース**する。
 *
 * CDP の addScriptToEvaluateOnNewDocument / Playwright の addInitScript でページ読込前に仕込むこと。
 * 拡張側が張る発信用 WebSocket は __mscopilotSelfConnecting フラグで除外する。
 */
export const WS_HOOK_SCRIPT = `(() => {
  try {
    if (window.__mscopilotHooked) return;
    window.__mscopilotHooked = true;
    var SUB = /substrate\\.(office\\.com|svc\\.cloud\\.microsoft)/i;
    // トレース対象(広め): substrate/認証系/Copilot 関連の通信を可視化する。
    var TOKENISH = /(login\\.microsoftonline|login\\.windows\\.net|\\/oauth2\\/|\\/token|substrate\\.(office\\.com|svc\\.cloud\\.microsoft)|copilot|sydney|m365chat|authgateway|\\.auth\\.|getaccesstoken|issuetoken)/i;
    // レスポンス本文を走査する対象(狭め): トークン発行系のみ。substrate の一般 API 本文は
    // 走査しない(巨大かつ Copilot トークンを含まないため)。Copilot 用トークンは MSAL の
    // トークン取得(oauth2/token 等)の本文を経由しうる。
    var AUTHISH = /(login\\.microsoftonline|login\\.windows\\.net|\\/oauth2\\/|\\/common\\/|issuetoken|getaccesstoken|authgateway|\\.auth\\.|\\/token)/i;
    var JWT_RE = /eyJ[A-Za-z0-9_\\-]+\\.[A-Za-z0-9_\\-]+\\.[A-Za-z0-9_\\-]+/g;

    // --- 解析用ネットワークトレース(トークンは載せない) ---
    var trace = function (s) {
      try {
        var a = globalThis.__mscopilotNet || (globalThis.__mscopilotNet = []);
        if (a.length < 400) a.push(String((Date.now() % 100000) + " " + s));
      } catch (e) {}
    };
    var shortUrl = function (u) {
      try { var s = String(u).split("?")[0]; return s.length > 140 ? s.slice(0, 140) + "…" : s; } catch (e) { return ""; }
    };

    // --- Bearer 蓄積(request header / response body 共通) ---
    var addBearer = function (tok, url, via) {
      try {
        if (!tok || String(tok).indexOf("eyJ") !== 0) return;
        globalThis.__mscopilotBearer = tok; // 後方互換: 最後に観測した Bearer
        var list = globalThis.__mscopilotBearers || (globalThis.__mscopilotBearers = []);
        for (var i = 0; i < list.length; i++) {
          if (list[i].t === tok) { list[i].u = String(url || list[i].u || ""); return; }
        }
        if (list.length < 60) {
          list.push({ t: tok, u: String(url || ""), via: String(via || "") });
          trace("TOKEN+ via=" + (via || "") + " " + shortUrl(url));
        }
      } catch (e) {}
    };
    var captureAuthHeader = function (auth, url) {
      try {
        if (!auth) return;
        var m = /Bearer\\s+(eyJ[A-Za-z0-9\\-_.]+\\.[A-Za-z0-9\\-_.]+\\.[A-Za-z0-9\\-_.]+)/i.exec(String(auth));
        if (m) addBearer(m[1], url, "req-header");
      } catch (e) {}
    };
    var scanBody = function (text, url) {
      try {
        if (!text || typeof text !== "string" || text.length > 3000000) return;
        if (text.indexOf("eyJ") === -1) return;
        var m, n = 0;
        JWT_RE.lastIndex = 0;
        while ((m = JWT_RE.exec(text)) !== null && n < 12) { addBearer(m[0], url, "resp-body"); n++; }
      } catch (e) {}
    };
    var headerGet = function (h, name) {
      try {
        if (!h) return "";
        var lower = name.toLowerCase();
        if (typeof h.get === "function") return h.get(name) || "";
        if (Array.isArray(h)) {
          for (var i = 0; i < h.length; i++) if (String(h[i][0]).toLowerCase() === lower) return h[i][1];
          return "";
        }
        for (var k in h) if (k.toLowerCase() === lower) return h[k];
      } catch (e) {}
      return "";
    };

    // 1) WebSocket
    var OW = window.WebSocket;
    var WrapWS = function (url, protocols) {
      try {
        var u = String(url);
        var self = !!globalThis.__mscopilotSelfConnecting;
        trace("WS" + (self ? "(self)" : "") + " " + shortUrl(u));
        // 拡張自身の発信接続(__mscopilotSelfConnecting)は記録しない(自作URLの再利用を防ぐ)。
        if (!self && u.indexOf("/Chathub/") !== -1 && u.indexOf("access_token=") !== -1) {
          if (!globalThis.__mscopilotWsUrl) globalThis.__mscopilotWsUrl = u;
          var wl = globalThis.__mscopilotWsUrls || (globalThis.__mscopilotWsUrls = []);
          if (wl.indexOf(u) === -1 && wl.length < 10) wl.push(u);
          // WS URL のクエリから Copilot 用 access_token も Bearer 候補として拾う。
          try {
            var mm = /[?&]access_token=([^&]+)/.exec(u);
            if (mm) addBearer(decodeURIComponent(mm[1]), u, "ws-url");
          } catch (e2) {}
        }
      } catch (e) {}
      return protocols !== undefined ? new OW(url, protocols) : new OW(url);
    };
    WrapWS.prototype = OW.prototype;
    WrapWS.CONNECTING = OW.CONNECTING; WrapWS.OPEN = OW.OPEN;
    WrapWS.CLOSING = OW.CLOSING; WrapWS.CLOSED = OW.CLOSED;
    window.WebSocket = WrapWS;

    // 2) fetch(request header + response body)
    var of = window.fetch;
    if (typeof of === "function") {
      window.fetch = function (input, init) {
        var url = "";
        try {
          url = typeof input === "string" ? input : (input && input.url) || "";
          if (SUB.test(url)) {
            var auth = init && init.headers ? headerGet(init.headers, "authorization") : "";
            if (!auth && input && input.headers && typeof input.headers.get === "function") {
              auth = input.headers.get("authorization") || "";
            }
            captureAuthHeader(auth, url);
          }
          if (TOKENISH.test(url)) trace("FETCH " + shortUrl(url));
        } catch (e) {}
        var res = of.apply(this, arguments);
        try {
          if (url && AUTHISH.test(url) && res && typeof res.then === "function") {
            res.then(function (resp) {
              try {
                trace("FETCH<- " + (resp && resp.status) + " " + shortUrl(url));
                if (resp && typeof resp.clone === "function") {
                  resp.clone().text().then(function (t) { scanBody(t, url); }, function () {});
                }
              } catch (e) {}
            }, function () {});
          }
        } catch (e) {}
        return res;
      };
    }

    // 3) XMLHttpRequest(request header + response body)
    var XP = XMLHttpRequest.prototype;
    var oOpen = XP.open, oSet = XP.setRequestHeader;
    XP.open = function (method, url) {
      try {
        this.__mscUrl = String(url || "");
        if (TOKENISH.test(this.__mscUrl)) {
          trace("XHR " + shortUrl(this.__mscUrl));
          var self = this;
          var scan = AUTHISH.test(this.__mscUrl);
          this.addEventListener("load", function () {
            try {
              trace("XHR<- " + self.status + " " + shortUrl(self.__mscUrl));
              if (!scan) return;
              var rt = (self.responseType === "" || self.responseType === "text") ? self.responseText : "";
              if (rt) scanBody(rt, self.__mscUrl);
            } catch (e) {}
          });
        }
      } catch (e) {}
      return oOpen.apply(this, arguments);
    };
    XP.setRequestHeader = function (key, value) {
      try {
        if (String(key).toLowerCase() === "authorization" && SUB.test(this.__mscUrl || "")) captureAuthHeader(value, this.__mscUrl);
      } catch (e) {}
      return oSet.apply(this, arguments);
    };
  } catch (e) {}
})()`;

/** 捕捉済みの substrate 接続 URL を返すスクリプト式(無ければ空文字)。 */
export const READ_WS_URL_SCRIPT = `(globalThis.__mscopilotWsUrl || "")`;

/** 捕捉済みの substrate 宛 Bearer トークン(JWT 本体)を返すスクリプト式(無ければ空文字)。 */
export const READ_BEARER_SCRIPT = `(globalThis.__mscopilotBearer || "")`;

/** 解析用ネットワークトレースを取り出して消費する(読むたびにクリア)スクリプト式。 */
export const READ_NET_SCRIPT = `(function(){try{var a=globalThis.__mscopilotNet||[];globalThis.__mscopilotNet=[];return JSON.stringify(a);}catch(e){return "[]";}})()`;

/**
 * 捕捉済みの substrate 宛 Bearer 群(URL付き)を JSON 文字列で返すスクリプト式。
 * サービスごとにスコープが異なるため、Node 側で Copilot(Chathub)用を選別する。
 */
export const READ_BEARERS_SCRIPT = `JSON.stringify(globalThis.__mscopilotBearers || [])`;

/** ページから採取した Bearer 1 件(t=JWT本体, u=採取元URL, via=採取経路)。 */
export interface HarvestedBearer {
  t: string;
  u: string;
  via?: string;
}

/**
 * 採取した Bearer 群を診断ログ用に 1 件 1 行で整形する(secret は含めない)。
 * どのスコープのトークンが・どの経路で採れているかを可視化し、Copilot 用トークンの
 * 有無を一目で判断できるようにする。score 降順。
 */
export function describeBearers(bearers: Array<HarvestedBearer | string>): string {
  const now = Date.now();
  const rows = bearers
    .map((raw) => {
      const b = typeof raw === "string" ? { t: raw, u: "", via: "" } : raw;
      const payload = decodeJwt(b.t);
      const exp = payload?.exp ? payload.exp * 1000 : undefined;
      const score = scoreText(matchText({ secret: b.t, target: b.u }, payload));
      return {
        score,
        via: b.via || "?",
        aud: payload?.aud || "?",
        scp: (payload?.scp || "").slice(0, 60),
        exp,
        expired: exp !== undefined ? exp <= now : undefined,
        url: (b.u || "").split("?")[0].slice(0, 100),
      };
    })
    .sort((a, b) => b.score - a.score);
  if (rows.length === 0) {
    return "  (Bearer 候補は 1 件も採取していません)";
  }
  return rows
    .map((r, i) => {
      const mark = r.score >= 3 ? "★" : r.score === 2 ? "○" : "  ";
      const exp = r.exp ? `${new Date(r.exp).toISOString()}${r.expired ? "(失効)" : ""}` : "?";
      return (
        `  ${mark} [${i}] score=${r.score} via=${r.via}\n` +
        `        aud=${r.aud} scp=${r.scp}\n` +
        `        exp=${exp} url=${r.url}`
      );
    })
    .join("\n");
}

/**
 * ページの substrate 宛 HTTP から採取した Bearer 群から、Copilot(Chathub)用として
 * 最良のトークンを 1 つ選ぶ。substrate は同一ホストで多数のサービスを提供し、サービスごとに
 * スコープの異なるトークンが飛ぶ。誤ったスコープのトークンでも WebSocket 認証は通るが、
 * モデル起動が拒否され本文に "Language model unavailable" が返るため、スコープ選別が要。
 *
 * 各 Bearer をリクエスト URL を target とした候補に変換し、既存の pickSydneyToken に委譲する
 * (URL の chathub/m365copilot 等 + JWT の aud/scp でスコアリングし、chat 用 > substrate 一般)。
 */
export function pickBearerToken(bearers: Array<HarvestedBearer | string>): SubstrateToken | undefined {
  const candidates: RawTokenCandidate[] = [];
  for (const b of bearers) {
    if (typeof b === "string") {
      if (b) candidates.push({ secret: b });
    } else if (b && typeof b.t === "string" && b.t) {
      candidates.push({ secret: b.t, target: typeof b.u === "string" ? b.u : "" });
    }
  }
  return pickSydneyToken(candidates);
}

/** トークンの素性(aud / scp / 選別スコア)を診断ログ用に整形する。secret は含めない。 */
export function formatTokenScope(token: SubstrateToken): string {
  const payload = decodeJwt(token.accessToken);
  const aud = payload?.aud || "?";
  const scp = (payload?.scp || "").slice(0, 80);
  const score = scoreText(matchText({ secret: token.accessToken }, payload));
  return `aud=${aud} scp=${scp} score=${score}`;
}

/**
 * JWT 本体(secret)から SubstrateToken を組み立てる。oid/tid が読めない JWT は対象外。
 * exp が読めない場合は保守的に 50 分後を仮の失効時刻とする。
 */
export function tokenFromSecret(secret: string | undefined): SubstrateToken | undefined {
  if (!secret) {
    return undefined;
  }
  const payload = decodeJwt(secret);
  if (!payload?.oid || !payload.tid) {
    return undefined;
  }
  const expiresAt = payload.exp ? payload.exp * 1000 : Date.now() + 50 * 60_000;
  return { accessToken: secret, objectId: payload.oid, tenantId: payload.tid, expiresAt };
}

/**
 * 捕捉した WebSocket URL から access_token を取り出し SubstrateToken を組み立てる。
 * oid/tid は JWT claim を優先し、無ければ URL パス({oid}@{tid})から補う。
 * exp が読めない場合は保守的に 50 分後を仮の失効時刻とする。
 */
export function parseWsUrlToken(url: string): SubstrateToken | undefined {
  const m = /[?&]access_token=([^&]+)/.exec(url);
  if (!m) {
    return undefined;
  }
  const secret = decodeURIComponent(m[1]);
  const payload = decodeJwt(secret);
  const pathMatch = /\/Chathub\/([^@/?]+)@([^/?]+)/.exec(url);
  const objectId = payload?.oid || (pathMatch ? decodeURIComponent(pathMatch[1]) : undefined);
  const tenantId = payload?.tid || (pathMatch ? decodeURIComponent(pathMatch[2]) : undefined);
  if (!objectId || !tenantId) {
    return undefined;
  }
  const expiresAt = payload?.exp ? payload.exp * 1000 : Date.now() + 50 * 60_000;
  return { accessToken: secret, objectId, tenantId, expiresAt };
}

export interface TokenInventoryEntry {
  credentialType: string;
  audience: string;
  scopes: string;
  target: string;
  store: string;
  expiresAt?: number;
  expired?: boolean;
  /** 選別スコア(3/2/0)。0 は対象外。 */
  score: number;
}

/**
 * 収集した全候補を診断用に一覧化する(選別に失敗した際、何が見つかったかを可視化する)。
 * secret は含めない。score 降順で返す。
 */
export function inventoryTokens(
  candidates: Array<RawTokenCandidate | string>,
): TokenInventoryEntry[] {
  const now = Date.now();
  const list: TokenInventoryEntry[] = candidates.map((raw) => {
    const cand = toCandidate(raw);
    const payload = decodeJwt(cand.secret);
    const exp = payload?.exp ? payload.exp * 1000 : undefined;
    return {
      credentialType: cand.credentialType || "?",
      audience: payload?.aud || "?",
      scopes: payload?.scp || "",
      target: cand.target || "",
      store: cand.store || "",
      expiresAt: exp,
      expired: exp !== undefined ? exp <= now : undefined,
      score: scoreText(matchText(cand, payload)),
    };
  });
  return list.sort((a, b) => b.score - a.score);
}

/** 診断一覧を読みやすいテキストへ整形する。 */
export function formatInventory(entries: TokenInventoryEntry[]): string {
  if (entries.length === 0) {
    return "  (トークン候補が 1 件も見つかりませんでした)";
  }
  return entries
    .map((e, i) => {
      const exp = e.expiresAt
        ? `${new Date(e.expiresAt).toISOString()}${e.expired ? " (失効)" : ""}`
        : "?";
      const mark = e.score >= 3 ? "★" : e.score === 2 ? "○" : "  ";
      const scope = (e.target || e.scopes || "").slice(0, 120);
      return (
        `  ${mark} [${i}] score=${e.score} type=${e.credentialType} store=${e.store}\n` +
        `        aud=${e.audience}\n` +
        `        scope=${scope}\n` +
        `        exp=${exp}`
      );
    })
    .join("\n");
}

/**
 * ページ内で実行し、MSAL の local/sessionStorage キャッシュから
 * トークン候補(secret + メタデータ)をすべて返すスクリプト式。
 *
 * credentialType で絞らず、JWT らしい secret を持つエントリを広く収集する。
 * MSAL のバージョン差(AccessToken_With_AuthScheme 等)やキャッシュ形状の違いに強くするため、
 * 選別・診断は Node 側(pickSydneyToken / inventoryTokens)に委ねる。
 */
export const COLLECT_TOKENS_SCRIPT = `(() => {
  const out = [];
  const looksLikeJwt = (s) =>
    typeof s === "string" && s.length > 40 && s.split(".").length === 3;
  const scan = (store, storeName) => {
    if (!store) return;
    let n = 0;
    try { n = store.length; } catch (e) { return; }
    for (let i = 0; i < n; i++) {
      let key, raw;
      try { key = store.key(i); raw = key ? store.getItem(key) : null; } catch (e) { continue; }
      if (!raw) continue;
      const t = raw.trim();
      if (t.charAt(0) !== "{") continue;
      let obj;
      try { obj = JSON.parse(t); } catch (e) { continue; }
      if (!obj || !looksLikeJwt(obj.secret)) continue;
      out.push({
        secret: obj.secret,
        target: typeof obj.target === "string" ? obj.target : "",
        clientId: typeof obj.clientId === "string" ? obj.clientId : "",
        realm: typeof obj.realm === "string" ? obj.realm : "",
        environment: typeof obj.environment === "string" ? obj.environment : "",
        credentialType: typeof obj.credentialType === "string" ? obj.credentialType : "",
        store: storeName
      });
    }
  };
  try { scan(window.localStorage, "localStorage"); } catch (e) {}
  try { scan(window.sessionStorage, "sessionStorage"); } catch (e) {}
  return out;
})()`;
