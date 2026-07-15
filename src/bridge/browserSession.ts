/**
 * ブラウザ接続の共通インターフェース。
 * CDP(ユーザーの Edge に相乗り)と Playwright(専用ブラウザ)の実装差異を吸収する。
 *
 * 設計方針(計画より):
 *  - 認証は「ブラウザにログイン済み」の M365 セッションから sydney トークンを読み取る。
 *  - チャット発信は可能な限り「ページ内」で行い、実 UA/TLS/セッションと一致させる
 *    (runStream)。フレームはページ→Node へ push 関数(binding)経由で返す。
 */

import { log } from "../logger";

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

export interface BrowserSession {
  /** ブラウザへ接続し、M365 ページを開いてサインイン状態を整える。 */
  ensureReady(): Promise<void>;

  /**
   * ログイン済みセッションから substrate(audience: .../sydney)トークンを取得する。
   * force=true で再読込して最新トークンを取り直す。
   */
  getToken(force?: boolean): Promise<SubstrateToken>;

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

const SYDNEY_AUD_HINTS = ["substrate.office.com/sydney", "/sydney", "sydney"];

interface JwtPayload {
  aud?: string;
  oid?: string;
  tid?: string;
  exp?: number;
  scp?: string;
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

/**
 * ブラウザから収集した AccessToken 候補群から、substrate(sydney)向けの
 * 有効なトークンを 1 つ選ぶ。JWT を復号して aud / oid / tid / exp を得る。
 */
export function pickSydneyToken(candidates: string[]): SubstrateToken | undefined {
  // 優先度: sydney audience > substrate ホスト。各グループ内では exp が新しい方。
  let sydney: SubstrateToken | undefined;
  let substrate: SubstrateToken | undefined;
  for (const secret of candidates) {
    if (!secret || typeof secret !== "string") {
      continue;
    }
    log.info(
      `token shape: length=${secret.length} parts=${secret.split(".").length} ` +
        `startsWithEyJ=${secret.startsWith("eyJ")}`,
    );
    const payload = decodeJwt(secret);

    log.info(
      `token shape: length=${secret.length} parts=${secret.split(".").length} ` +
        `startsWithEyJ=${secret.startsWith("eyJ")}`,
    );
    
    log.info(
      `token metadata: aud=${payload?.aud ?? "(none)"} ` +
        `oid=${payload?.oid ? "present" : "missing"} ` +
        `tid=${payload?.tid ? "present" : "missing"} ` +
        `exp=${payload?.exp ?? "(none)"} ` +
        `scp=${payload?.scp ?? "(none)"}`,
    );
    if (!payload?.aud || !payload.oid || !payload.tid || !payload.exp) {
      continue;
    }
    const aud = payload.aud.toLowerCase();
    const token: SubstrateToken = {
      accessToken: secret,
      objectId: payload.oid,
      tenantId: payload.tid,
      expiresAt: payload.exp * 1000,
    };
    if (SYDNEY_AUD_HINTS.some((h) => aud.includes(h))) {
      if (!sydney || token.expiresAt > sydney.expiresAt) {
        sydney = token;
      }
    } else if (aud.includes("substrate")) {
      // 新系統(cloud.microsoft)で audience が異なる場合のフォールバック。
      if (!substrate || token.expiresAt > substrate.expiresAt) {
        substrate = token;
      }
    }
  }
  return sydney ?? substrate;
}

/**
 * ページ内で実行し、MSAL の localStorage キャッシュから AccessToken の
 * secret 文字列をすべて返すスクリプト式。トークンの選別は Node 側で行う。
 */
export const COLLECT_TOKENS_SCRIPT = `(() => {
  const out = [];
  const scan = (store) => {
    if (!store) return;
    try {
      for (let i = 0; i < store.length; i++) {
        const key = store.key(i);
        if (!key) continue;
        const raw = store.getItem(key);
        if (!raw || raw[0] !== "{") continue;
        let obj;
        try { obj = JSON.parse(raw); } catch { continue; }
        if (obj && obj.credentialType === "AccessToken" && typeof obj.secret === "string") {
          out.push(obj.secret);
        }
      }
    } catch (e) { /* ignore */ }
  };
  try { scan(window.localStorage); } catch (e) {}
  try { scan(window.sessionStorage); } catch (e) {}
  return out;
})()`;
