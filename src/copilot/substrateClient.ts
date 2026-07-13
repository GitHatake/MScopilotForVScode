import { randomUUID } from "node:crypto";
import type { BrowserSession } from "../bridge/browserSession";
import type { MsCopilotConfig } from "../config";
import { log } from "../logger";
import { RUN_STREAM_SCRIPT } from "./injectedClient";
import {
  RECORD_SEP,
  buildInvocation,
  buildWsUrl,
  encodeFrame,
  extractConversationId,
  handshakeFrame,
  interpretFrame,
  parseFrames,
  pingFrame,
} from "./protocol";

export interface AskParams {
  prompt: string;
  /** 既存会話の継続なら会話 ID。新規なら未指定。 */
  conversationId?: string;
  /** 差分テキストを逐次受け取るコールバック。 */
  onDelta: (delta: string) => void;
  signal?: AbortSignal;
  /** トークンを強制的に取り直してから接続する(期限切れ再試行用)。 */
  forceToken?: boolean;
}

export interface AskResult {
  text: string;
  conversationId: string;
  error?: string;
}

/**
 * substrate(BizChat)への 1 リクエストを実行する。
 * WebSocket 接続と送受信はブラウザページ内で行い、フレーム解釈は Node 側で行う。
 */
export class SubstrateClient {
  constructor(
    private readonly session: BrowserSession,
    private readonly config: MsCopilotConfig,
  ) {}

  async ask(params: AskParams): Promise<AskResult> {
    const token = await this.session.getToken(params.forceToken);
    const conversationId = params.conversationId ?? randomUUID();
    const isStartOfSession = !params.conversationId;

    const endpoint = buildWsUrl(this.config.endpointVariant, token, conversationId);
    const invocation = buildInvocation({
      prompt: params.prompt,
      conversationId,
      isStartOfSession,
      invocationId: "0",
    });

    const arg = {
      wsUrl: endpoint.url,
      handshake: handshakeFrame(),
      invocation: encodeFrame(invocation),
      ping: pingFrame(),
      sep: RECORD_SEP,
      timeoutMs: 300000,
    };

    log.info(
      `ask: convId=${conversationId} start=${isStartOfSession} variant=${this.config.endpointVariant}`,
    );
    log.info(`ws endpoint: ${redactToken(endpoint.url)}`);

    let lastText = "";
    let resolvedConversationId = conversationId;
    let errorText: string | undefined;

    const onFrame = (frame: unknown) => {
      const meta = frame as {
        __transport?: string;
        raw?: string;
        code?: number;
        reason?: string;
      };
      if (meta?.__transport) {
        log.info(
          `transport: ${meta.__transport}` +
            (meta.raw ? ` ${meta.raw}` : "") +
            (meta.code !== undefined ? ` code=${meta.code} reason=${meta.reason ?? ""}` : ""),
        );
        // 正常終了(1000)以外でクローズし、本文が無ければエラーとして扱う。
        // これにより認証切れ等を検知し、上位でのトークン再取得リトライにつながる。
        if (
          meta.__transport === "closed" &&
          meta.code !== undefined &&
          meta.code !== 1000 &&
          lastText === "" &&
          !errorText
        ) {
          errorText = `WebSocket が異常終了しました (code ${meta.code}${meta.reason ? `: ${meta.reason}` : ""})`;
        }
        if (meta.__transport === "error" && lastText === "" && !errorText) {
          errorText = "WebSocket 接続エラー(CSP/ネットワーク/認証を確認してください)";
        }
        return;
      }

      logRawFrame(frame);

      const echoed = extractConversationId(frame);
      if (echoed) {
        resolvedConversationId = echoed;
      }

      const interp = interpretFrame(frame);
      if (interp.isPing) {
        return; // pong はページ内で処理済み
      }
      if (interp.error) {
        errorText = interp.error;
        log.warn(`frame error: ${interp.error}`);
      }
      if (interp.fullText !== undefined) {
        const delta = diff(lastText, interp.fullText);
        if (delta) {
          lastText = interp.fullText;
          params.onDelta(delta);
        }
      }
    };

    try {
      await this.session.runStream({
        script: RUN_STREAM_SCRIPT,
        arg,
        onFrame,
        signal: params.signal,
      });
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      log.error("runStream failed", msg);
      if (!errorText) {
        errorText = msg;
      }
    }

    return { text: lastText, conversationId: resolvedConversationId, error: errorText };
  }
}

/**
 * スナップショット方式(累積本文が毎回届く)を前提に差分を計算する。
 * prev が next の接頭辞なら末尾差分、そうでなければ next 全体を新規本文として返す。
 */
function diff(prev: string, next: string): string {
  if (next.startsWith(prev)) {
    return next.slice(prev.length);
  }
  // 稀に本文が置き換わる場合。UI では前回分の後に続けて出す。
  return next;
}

function redactToken(url: string): string {
  return url.replace(/access_token=[^&]+/, "access_token=***");
}

function logRawFrame(frame: unknown): void {
  try {
    const s = JSON.stringify(frame);
    log.info(`frame: ${s.length > 600 ? s.slice(0, 600) + "…" : s}`);
  } catch {
    /* ignore */
  }
}

// 参照維持(将来 Node 側で生フレームを再パースする用途)。
export { parseFrames };
