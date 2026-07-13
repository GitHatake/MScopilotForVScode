import { randomUUID } from "node:crypto";
import type { EndpointVariant } from "../config";
import type { SubstrateToken } from "../bridge/browserSession";

/**
 * substrate(BizChat)は SignalR JSON hub protocol を使う。
 * フレームは 0x1e(record separator)区切りの JSON。
 *
 * 注意: これは非公式・未文書のプロトコルの再現である。実トラフィックにより
 * 微調整が必要な箇所は本ファイルに集約してある(payload の optionsSets など)。
 * substrateClient は生フレームをログ出力するので、実機で確認しながら調整できる。
 */

export const RECORD_SEP = "\u001e";

/** SignalR メッセージ種別 */
export const MSG = {
  INVOCATION: 1,
  STREAM_ITEM: 2,
  COMPLETION: 3,
  STREAM_INVOCATION: 4,
  CANCEL: 5,
  PING: 6,
  CLOSE: 7,
} as const;

export interface WsEndpointInfo {
  url: string;
  clientRequestId: string;
  sessionId: string;
}

/** WebSocket URL を組み立てる(トークンはクエリ文字列で渡す仕様)。 */
export function buildWsUrl(
  variant: EndpointVariant,
  token: SubstrateToken,
  conversationId?: string,
): WsEndpointInfo {
  const clientRequestId = randomUUID();
  const sessionId = randomUUID();
  const userKey = `${token.objectId}@${token.tenantId}`;

  const base =
    variant === "cloud"
      ? `wss://substrate.svc.cloud.microsoft/m365Copilot/Chathub/${userKey}`
      : `wss://substrate.office.com/m365chat/SecuredChathub/${userKey}`;

  const params = new URLSearchParams();
  params.set("X-ClientRequestId", clientRequestId);
  params.set("X-SessionId", sessionId);
  if (conversationId) {
    params.set("ConversationId", conversationId);
  }
  // access_token は必ず最後に付与(ヘッダではなくクエリで渡す)。
  params.set("access_token", token.accessToken);

  return { url: `${base}?${params.toString()}`, clientRequestId, sessionId };
}

export interface InvocationParams {
  prompt: string;
  conversationId: string;
  /** 会話の最初の送信か(true で新規セッション扱い) */
  isStartOfSession: boolean;
  /** 追加コンテキスト(#file/#folder の解決結果など)。プロンプト先頭に注入済みでも可。 */
  invocationId: string;
}

/**
 * type 4(StreamInvocation)ペイロードを構築する。
 * optionsSets / source / scenario などは実機トラフィックに合わせて調整する箇所。
 */
export function buildInvocation(p: InvocationParams): Record<string, unknown> {
  return {
    arguments: [
      {
        source: "officeweb",
        scenario: "office",
        optionsSets: [
          "enterprise_toolbox",
          "enterprise_flux",
          "streamingsupport",
          "gpt4",
        ],
        allowedMessageTypes: [
          "Chat",
          "InternalSearchQuery",
          "InternalSearchResult",
          "InternalLoaderMessage",
          "RenderCardRequest",
        ],
        isStartOfSession: p.isStartOfSession,
        message: {
          author: "user",
          inputMethod: "Keyboard",
          messageType: "Chat",
          text: p.prompt,
        },
        conversationId: p.conversationId,
        participant: { id: "" },
      },
    ],
    invocationId: p.invocationId,
    target: "chat",
    type: MSG.STREAM_INVOCATION,
  };
}

export function encodeFrame(obj: unknown): string {
  return JSON.stringify(obj) + RECORD_SEP;
}

export function handshakeFrame(): string {
  return encodeFrame({ protocol: "json", version: 1 });
}

export function pingFrame(): string {
  return encodeFrame({ type: MSG.PING });
}

/** 受信バッファを 0x1e で分割し JSON.parse する。ハンドシェイク応答 `{}` も含む。 */
export function parseFrames(buffer: string): unknown[] {
  const frames: unknown[] = [];
  for (const chunk of buffer.split(RECORD_SEP)) {
    if (!chunk) {
      continue;
    }
    try {
      frames.push(JSON.parse(chunk));
    } catch {
      // 不完全/未知のフレームは無視(呼び出し側で残余をバッファする)
    }
  }
  return frames;
}

export interface FrameInterpretation {
  /** これまでの累積本文(存在する場合)。差分ではなくスナップショットのことが多い。 */
  fullText?: string;
  /** 会話が完了したか(type 2/3、または final フラグ)。 */
  done: boolean;
  /** エラー本文(あれば)。 */
  error?: string;
  /** ping フレームか(呼び出し側で pong 返信)。 */
  isPing: boolean;
}

/**
 * 1 フレームを解釈して本文/完了/エラーを取り出す。
 * BizChat/Sydney は messages 配列の text をスナップショットで送る傾向があるため、
 * 差分ではなく「最新の累積本文」を返す設計にしている。防御的に多形を許容する。
 */
export function interpretFrame(frame: unknown): FrameInterpretation {
  const f = frame as Record<string, any>;
  const type = typeof f?.type === "number" ? (f.type as number) : undefined;

  if (type === MSG.PING) {
    return { done: false, isPing: true };
  }

  // 完了系: type 3 (Completion) もしくは type 2 でストリーム終端が示される。
  if (type === MSG.COMPLETION) {
    const err = extractError(f);
    return { done: true, isPing: false, error: err };
  }

  const fullText = extractText(f);
  const done = isTerminal(f);
  const err = extractError(f);
  return { fullText, done, isPing: false, error: err };
}

function extractText(f: Record<string, any>): string | undefined {
  // 代表的な形: { arguments: [ { messages: [ { author:"bot", text:"..." } ] } ] }
  const args = f?.arguments;
  const item = f?.item;
  const containers = [
    ...(Array.isArray(args) ? args : []),
    ...(item ? [item] : []),
  ];
  let best: string | undefined;
  for (const c of containers) {
    const messages = c?.messages;
    if (Array.isArray(messages)) {
      for (const m of messages) {
        if (
          m &&
          (m.author === "bot" || m.author === "assistant") &&
          typeof m.text === "string" &&
          m.text.length > 0 &&
          isVisibleChatMessage(m)
        ) {
          best = m.text;
        }
      }
    }
    // フォールバック: 直接 text を持つ形
    if (typeof c?.text === "string" && c.text.length > 0) {
      best = c.text;
    }
  }
  return best;
}

/**
 * 表示対象の本文メッセージだけを対象にする。検索クエリ/ローダ/カード等の
 * 内部メッセージは本文として扱わない(誤ってストリームに出さないため)。
 */
const HIDDEN_MESSAGE_TYPES = new Set([
  "InternalSearchQuery",
  "InternalSearchResult",
  "InternalLoaderMessage",
  "RenderCardRequest",
  "Progress",
  "Disengaged",
]);

function isVisibleChatMessage(m: Record<string, any>): boolean {
  const type = m?.messageType;
  if (typeof type === "string" && HIDDEN_MESSAGE_TYPES.has(type)) {
    return false;
  }
  if (m?.hiddenText && !m?.text) {
    return false;
  }
  return true;
}

function isTerminal(f: Record<string, any>): boolean {
  const args = f?.arguments;
  const item = f?.item;
  const containers = [
    ...(Array.isArray(args) ? args : []),
    ...(item ? [item] : []),
  ];
  for (const c of containers) {
    if (c?.messages === undefined && c?.result) {
      return true;
    }
    if (c?.isFinal === true || c?.done === true) {
      return true;
    }
  }
  return false;
}

/** フレームから会話 ID を拾えれば返す(サーバーが採番/エコーする場合に備える)。 */
export function extractConversationId(frame: unknown): string | undefined {
  const f = frame as Record<string, any>;
  const args = f?.arguments;
  const item = f?.item;
  const containers = [
    ...(Array.isArray(args) ? args : []),
    ...(item ? [item] : []),
  ];
  for (const c of containers) {
    if (typeof c?.conversationId === "string" && c.conversationId) {
      return c.conversationId;
    }
  }
  if (typeof f?.conversationId === "string" && f.conversationId) {
    return f.conversationId;
  }
  return undefined;
}

function extractError(f: Record<string, any>): string | undefined {
  if (typeof f?.error === "string" && f.error.length > 0) {
    return f.error;
  }
  const args = f?.arguments;
  if (Array.isArray(args)) {
    for (const c of args) {
      if (typeof c?.error === "string") {
        return c.error;
      }
      const value = c?.result?.value;
      if (typeof value === "string" && value !== "Success" && value !== "") {
        return `${value}: ${c?.result?.message ?? ""}`.trim();
      }
    }
  }
  return undefined;
}
