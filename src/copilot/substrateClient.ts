import { randomUUID } from "node:crypto";
import type { BrowserSession } from "../bridge/browserSession";
import { formatTokenScope } from "../bridge/browserSession";
import type { MsCopilotConfig } from "../config";
import { log } from "../logger";
import { RUN_STREAM_SCRIPT } from "./injectedClient";
import { ResponseAssembler } from "./responseAssembler";
import {
  RECORD_SEP,
  buildInvocation,
  buildWsUrl,
  deriveWsUrlFromTemplate,
  encodeFrame,
  extractConversationId,
  handshakeFrame,
  interpretFrame,
  newSessionIds,
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
    const ids = newSessionIds();
    const reqId = ids.correlationId.slice(0, 8);
    const t0 = Date.now();
    log.info(`[${reqId}] ask 開始: forceToken=${!!params.forceToken} prompt=${params.prompt.length}字`);

    const token = await this.session.getToken(params.forceToken);
    log.info(`[${reqId}] token: ${formatTokenScope(token)} exp=${new Date(token.expiresAt).toISOString()}`);
    const scope = formatTokenScope(token);
    if (/score=[01]\b/.test(scope)) {
      log.warn(
        `[${reqId}] ⚠ 採用トークンが Copilot 用と判別できません(score<2)。"Language model unavailable" の恐れ。`,
      );
    } else if (/score=2\b/.test(scope)) {
      log.warn(
        `[${reqId}] ⚠ 採用トークンは substrate 一般スコープ(score=2)。Copilot(Chathub)専用ではない可能性。`,
      );
    }

    const conversationId = params.conversationId ?? randomUUID();
    const isStartOfSession = !params.conversationId;

    // ページの実接続テンプレートが取れればそれを流用する(実 variants/token/scenario を使え、
    // 仕様変更に強い)。取れなければ実測値ベースの buildWsUrl でフォールバック。
    const template = (await this.session.harvestWsTemplate?.()) || undefined;
    const endpoint =
      (template && deriveWsUrlFromTemplate(template, conversationId, ids)) ||
      buildWsUrl(this.config.endpointVariant, token, conversationId, ids);
    log.info(
      `[${reqId}] endpoint: ${template ? "実接続テンプレート流用" : "合成(buildWsUrl)"} host=${hostOf(endpoint.url)}`,
    );

    const locale = localeInfo();
    const invocation = buildInvocation({
      prompt: params.prompt,
      isStartOfSession,
      invocationId: "0",
      sessionId: ids.sessionId,
      correlationId: ids.correlationId,
      locale: locale.locale,
      timeZone: locale.timeZone,
      timeZoneOffset: locale.timeZoneOffset,
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
      `[${reqId}] ask: convId=${conversationId} start=${isStartOfSession} variant=${this.config.endpointVariant} ` +
        `handshake=${arg.handshake.length}B invocation=${arg.invocation.length}B`,
    );
    log.info(`[${reqId}] ws url: ${redactToken(endpoint.url)}`);

    const assembler = new ResponseAssembler(params.onDelta);
    let resolvedConversationId = conversationId;
    let errorText: string | undefined;
    let frameCount = 0;
    let opened = false;
    let handshakeAt = 0;

    const onFrame = (frame: unknown) => {
      const meta = frame as {
        __transport?: string;
        raw?: string;
        code?: number;
        reason?: string;
      };
      if (meta?.__transport) {
        if (meta.__transport === "open") {
          opened = true;
        }
        if (meta.__transport === "handshake") {
          handshakeAt = Date.now() - t0;
        }
        log.info(
          `[${reqId}] transport: ${meta.__transport} (+${Date.now() - t0}ms)` +
            (meta.raw ? ` ${truncate(meta.raw, 200)}` : "") +
            (meta.code !== undefined ? ` code=${meta.code} reason=${meta.reason ?? ""}` : ""),
        );
        // 正常終了(1000)以外でクローズし、本文が無ければエラーとして扱う。
        // これにより認証切れ等を検知し、上位でのトークン再取得リトライにつながる。
        if (
          meta.__transport === "closed" &&
          meta.code !== undefined &&
          meta.code !== 1000 &&
          !assembler.produced &&
          !errorText
        ) {
          errorText = `WebSocket が異常終了しました (code ${meta.code}${meta.reason ? `: ${meta.reason}` : ""})`;
        }
        if (meta.__transport === "error" && !assembler.produced && !errorText) {
          errorText = "WebSocket 接続エラー(CSP/ネットワーク/認証を確認してください)";
        }
        return;
      }

      frameCount++;
      logRawFrame(reqId, frameCount, frame);

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
        log.warn(`[${reqId}] frame error: ${interp.error}`);
      }
      // writeAtCursor は差分なので末尾へ追記する。
      if (interp.appendText) {
        assembler.appendDelta(interp.appendText);
      }
      // messages[].text は累積スナップショットなので差分を取り出す。
      if (interp.fullText !== undefined) {
        assembler.setSnapshot(interp.fullText);
      }
      if (interp.done) {
        assembler.end();
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
      log.error(`[${reqId}] runStream failed`, msg);
      if (!errorText) {
        errorText = msg;
      }
    }

    // ストリーム終了。保留中の本文があれば最終判定(失敗フレーズなら error 化)して確定する。
    assembler.end();
    if (assembler.failure) {
      log.warn(`[${reqId}] service failure(実応答ではない): ${assembler.failure}`);
    }

    const error = errorText ?? assembler.failure;
    log.info(
      `[${reqId}] ask 完了 (+${Date.now() - t0}ms): opened=${opened} handshake=${handshakeAt ? handshakeAt + "ms" : "無"} ` +
        `frames=${frameCount} textLen=${assembler.text.length} produced=${assembler.produced} ` +
        `failure=${assembler.failure ? "有" : "無"} error=${error ? `"${truncate(error, 120)}"` : "無"}`,
    );
    if (!opened && !error) {
      log.warn(`[${reqId}] ⚠ WebSocket が open にならないまま終了(接続到達性/CSP/認証を確認)`);
    }

    return {
      text: assembler.text,
      conversationId: resolvedConversationId,
      error,
    };
  }
}

function redactToken(url: string): string {
  return url.replace(/access_token=[^&]+/, "access_token=***");
}

function truncate(s: string, n: number): string {
  return s.length > n ? s.slice(0, n) + "…" : s;
}

function hostOf(url: string): string {
  const m = /^wss?:\/\/([^/]+)/i.exec(url);
  return m ? m[1] : "?";
}

/** 実行環境の locale / timezone を invocation.message 用に取得する(取れなければ既定)。 */
function localeInfo(): { locale: string; timeZone: string; timeZoneOffset: number } {
  try {
    const resolved = Intl.DateTimeFormat().resolvedOptions();
    const locale = (resolved.locale || "en-us").toLowerCase();
    const timeZone = resolved.timeZone || "UTC";
    // JS の getTimezoneOffset は「UTC−ローカル」の分。substrate は「ローカル−UTC」の時。
    const timeZoneOffset = -new Date().getTimezoneOffset() / 60;
    return { locale, timeZone, timeZoneOffset };
  } catch {
    return { locale: "en-us", timeZone: "UTC", timeZoneOffset: 0 };
  }
}

function logRawFrame(reqId: string, n: number, frame: unknown): void {
  try {
    const s = JSON.stringify(frame);
    log.info(`[${reqId}] frame#${n}: ${truncate(s, 800)}`);
  } catch {
    /* ignore */
  }
}

// 参照維持(将来 Node 側で生フレームを再パースする用途)。
export { parseFrames };
