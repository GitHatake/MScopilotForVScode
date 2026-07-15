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
