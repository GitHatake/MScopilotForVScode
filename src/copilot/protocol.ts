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
  /** 相関 ID(ダッシュ無し)。invocation の traceId/requestId/clientCorrelationId に流用する。 */
  correlationId: string;
  /** セッション ID(ダッシュあり)。invocation の sessionId/clientSessionId に流用する。 */
  sessionId: string;
}

export interface SessionIds {
  /** ダッシュ無し(chatsessionid / clientrequestid / traceId 等) */
  correlationId: string;
  /** ダッシュあり(X-SessionId / clientSessionId) */
  sessionId: string;
}

/** 1 リクエスト分のセッション ID 群を生成する。dashed/compact は同一 UUID 由来。 */
export function newSessionIds(): SessionIds {
  const sessionId = randomUUID();
  return { sessionId, correlationId: sessionId.replace(/-/g, "") };
}

/**
 * 実トラフィック(officeweb, 2026/7 実測)で観測された variants フラグ列(そのまま)。
 * ページの実接続テンプレートが取れない場合のフォールバックとして、実値を丸ごと使う
 * (既知の通る組み合わせを再現するのが一発成功に最も近い)。
 */
const DEFAULT_VARIANTS =
  "EnableMcpServerWidgets,feature.EnableMcpServerWidgets,feature.EnableImageGenInsufficientTokensThrottled," +
  "feature.EnableImageGenSystemCapacityThrottled,feature.EnableLuForChatCIQ,feature.enableChatCIQPlugin," +
  "EnableRequestPlugins,feature.EnableSensitivityLabels,EnableUnsupportedUrlDetector," +
  "feature.IsCustomEngineCopilotEnabled,feature.bizchatfluxv3,feature.enablechatpages,feature.enableCodeCanvas," +
  "feature.turnOnDARecommendation,feature.IsStreamingModeInChatRequestEnabled,IncludeSourceAttributionsConcise," +
  "SkipPublishEmptyMessage,feature.EnableDeduplicatingSourceAttributions,feature.IsCitationsReferencesOutputEnabled," +
  "feature.enableDeltaStreamingForReferences,feature.enableIncludeReferencesInDeltaResponse," +
  "feature.enablereferencesforagents,Enable3PActionProgressMessages,feature.enableClientWebRtc," +
  "feature.EnableMeetingRecapOfSeriesMeetingWithCiq,feature.EnableReferencesListCompleteSignal," +
  "feature.StorageMessageSplitDisabled,feature.EnableCuaTakeControlApi,SingletonEnvOn," +
  "agt_bizchat_enablePagesCitations,agt_module_canvasSetup_enablePagesCitations," +
  "agt_bizchat_enablePagesCitationsForMultiturn,agt_module_canvasSetup_enablePagesCitationsForMultiturn," +
  "cdxenablefccinmainline,EnableComposeWidget,-agt_researcheragent_enableMemoryRead,feature.cwcallowedos," +
  "feature.EnableMergingPureDeltas,feature.disabledisallowedmsgs,feature.enableCitationsForSynthesisData," +
  "feature.EnableConversationShareApis,feature.enableGenerateGraphicArtOptionsSet,cdximagen," +
  "feature.EnableUpdatedUXForConfirmationDialog,feature.EnableContentApiandDocTypeHtmlInRichAnswers," +
  "cdxgrounding_api_v2_rich_web_answers_reference_bottom_force,cdxenablerenderforisocomp," +
  "feature.EnableClientFileURLSupportForOfficeWebPaidCopilot,feature.EnableDesignEditorImageGrounding," +
  "feature.EnableDesignerEditor,feature.EnableSkipRehydrationForSpeCIdImages,feature.EnablePersonalization," +
  "rich_responses,feature.EnableBase64DataInMessageAnnotations,feature.EnableSkipEmittingMessageOnFlush," +
  "feature.EnableRemoveEmptySourceAttributions,feature.EnableRemoveStreamingMode,feature.OfficeWebToHelix," +
  "feature.OfficeDesktopToHelix,feature.M365TeamsHubToHelix,feature.OwaHubToHelix,feature.MonarchHubToHelix," +
  "feature.Win32OutlookHubToHelix,feature.MacOutlookHubToHelix,Agt_bizchat_enableGpt5ForHelix";

/** type4 の optionsSets(実測、officeweb をそのまま)。 */
const OPTIONS_SETS = [
  "search_result_progress_messages_with_search_queries",
  "update_textdoc_response_after_streaming",
  "deepleo_networking_timeout_10minutes_canmore",
  "cwc_flux_image",
  "cwc_code_interpreter",
  "cwc_code_interpreter_amsfix",
  "cwcfluxgptv",
  "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
  "gptvnorm2048",
  "cwc_code_interpreter_citation_fix",
  "code_interpreter_interactive_charts",
  "cwc_code_interpreter_interactive_charts_inline_image",
  "code_interpreter_matplotlib_patching",
  "cwc_fileupload_odb",
  "update_memory_plugin",
  "add_custom_instructions",
  "cwc_flux_v3",
  "flux_v3_progress_messages",
  "enable_batch_token_processing",
  "enable_gg_gpt",
  "flux_v3_references",
  "flux_v3_references_entities",
  "flux_v3_image_gen_enable_dimensions",
  "flux_v3_image_gen_enable_non_watermarked_storage",
  "flux_v3_image_gen_enable_icon_dimensions",
  "flux_v3_image_gen_enable_system_text_with_params",
  "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
  "flux_v3_image_gen_enable_story",
  "rich_responses",
  "pages_citations",
  "pages_citations_multiturn",
];

/** type4 の allowedMessageTypes(実測、officeweb をそのまま)。 */
const ALLOWED_MESSAGE_TYPES = [
  "Chat",
  "Suggestion",
  "InternalSearchQuery",
  "Disengaged",
  "InternalLoaderMessage",
  "Progress",
  "GeneratedCode",
  "RenderCardRequest",
  "AdsQuery",
  "SemanticSerp",
  "GenerateContentQuery",
  "GenerateGraphicArt",
  "SearchQuery",
  "ConfirmationCard",
  "AuthError",
  "DeveloperLogs",
  "TriggerPlugin",
  "HintInvocation",
  "MemoryUpdate",
  "EndOfRequest",
  "TriggerConfirmation",
  "ResumeInvokeAction",
  "ResumeUserInputRequest",
  "TriggerUserInputRequest",
  "EscapeHatch",
  "TriggerPluginAuth",
  "ResumePluginAuth",
  "SideBySide",
  "ReferencesListComplete",
  "SwitchRespondingEndpoint",
];

/**
 * WebSocket URL を組み立てる(トークンはクエリ文字列で渡す仕様)。
 *
 * 実測(2026/7, officeweb)に合わせたパス/クエリ:
 *   wss://substrate.office.com/m365Copilot/Chathub/{oid}@{tid}?chatsessionid=…&access_token=…&…
 * ページの実接続を捕捉できる場合は deriveWsUrlFromTemplate を優先すること
 * (variants 等が実テナントの実値になり、仕様変更にも強い)。
 */
export function buildWsUrl(
  variant: EndpointVariant,
  token: SubstrateToken,
  conversationId?: string,
  ids: SessionIds = newSessionIds(),
): WsEndpointInfo {
  const userKey = `${token.objectId}@${token.tenantId}`;
  const host = variant === "cloud" ? "substrate.svc.cloud.microsoft" : "substrate.office.com";
  const base = `wss://${host}/m365Copilot/Chathub/${userKey}`;

  // 実測の並び・エンコードに合わせて手組みする(commas を維持、source は "officeweb" のまま)。
  const parts = [
    `chatsessionid=${ids.correlationId}`,
    `XRoutingParameterSessionKey=${ids.correlationId}`,
    `clientrequestid=${ids.correlationId}`,
    `X-SessionId=${ids.sessionId}`,
    conversationId ? `ConversationId=${conversationId}` : undefined,
    `access_token=${token.accessToken}`,
    `variants=${DEFAULT_VARIANTS}`,
    `source=%22officeweb%22`,
    `product=Office`,
    `agentHost=Bizchat.FullScreen`,
    `licenseType=Starter`,
    `isEdu=false`,
    `agent=web`,
    `scenario=OfficeWebIncludedCopilot`,
  ].filter((x): x is string => typeof x === "string");

  return { url: `${base}?${parts.join("&")}`, correlationId: ids.correlationId, sessionId: ids.sessionId };
}

/**
 * ページの実接続で捕捉した WebSocket URL をテンプレートとして、自分のリクエスト用 URL を作る。
 * access_token / variants / source / scenario などは実値をそのまま流用し、
 * セッション ID 群と ConversationId だけ差し替える(ユーザーの実接続と衝突させない)。
 */
export function deriveWsUrlFromTemplate(
  template: string,
  conversationId: string,
  ids: SessionIds = newSessionIds(),
): WsEndpointInfo | undefined {
  const qIndex = template.indexOf("?");
  if (qIndex < 0) {
    return undefined;
  }
  const base = template.slice(0, qIndex);
  const overrides: Record<string, string> = {
    chatsessionid: ids.correlationId,
    XRoutingParameterSessionKey: ids.correlationId,
    clientrequestid: ids.correlationId,
    "X-SessionId": ids.sessionId,
    ConversationId: conversationId,
  };
  const seen = new Set<string>();
  const rebuilt = template
    .slice(qIndex + 1)
    .split("&")
    .map((pair) => {
      const eq = pair.indexOf("=");
      const key = eq >= 0 ? pair.slice(0, eq) : pair;
      if (key in overrides) {
        seen.add(key);
        return `${key}=${overrides[key]}`;
      }
      return pair;
    });
  if (!seen.has("ConversationId")) {
    rebuilt.push(`ConversationId=${conversationId}`);
  }
  return {
    url: `${base}?${rebuilt.join("&")}`,
    correlationId: ids.correlationId,
    sessionId: ids.sessionId,
  };
}

export interface InvocationParams {
  prompt: string;
  /** 会話の最初の送信か(true で新規セッション扱い) */
  isStartOfSession: boolean;
  invocationId: string;
  /** URL と揃えるセッション ID(ダッシュあり)。 */
  sessionId: string;
  /** URL と揃える相関 ID(ダッシュ無し)。requestId/traceId 等に使う。 */
  correlationId: string;
  locale?: string;
  timeZone?: string;
  timeZoneOffset?: number;
}

/**
 * type 4(StreamInvocation)ペイロードを構築する。実測(officeweb)の形に合わせてある。
 * conversationId は URL 側のクエリで渡すため arguments には含めない。
 */
export function buildInvocation(p: InvocationParams): Record<string, unknown> {
  const locale = p.locale ?? "en-us";
  const timeZone = p.timeZone ?? "UTC";
  const timeZoneOffset = p.timeZoneOffset ?? 0;
  return {
    arguments: [
      {
        source: "officeweb",
        clientCorrelationId: p.correlationId,
        sessionId: p.sessionId,
        optionsSets: OPTIONS_SETS,
        streamingMode: "ConciseWithPadding",
        options: {},
        extraExtensionParameters: {},
        allowedMessageTypes: ALLOWED_MESSAGE_TYPES,
        sliceIds: [],
        threadLevelGptId: {},
        traceId: p.correlationId,
        isStartOfSession: p.isStartOfSession,
        clientInfo: {
          clientPlatform: "mcmcopilot-web",
          clientAppName: "Office",
          clientEntrypoint: "mcmcopilot-officeweb",
          clientSessionId: p.sessionId,
          ProductCategory: "Chat",
          clientAppType: "Web",
          productEntryPoint: "ChatPanel",
          deviceOS: "Windows",
          deviceType: "Desktop",
          clientPlatformVersion: "10",
        },
        message: {
          author: "user",
          inputMethod: "Keyboard",
          text: p.prompt,
          entityAnnotationTypes: ["People", "File", "Event", "Email", "TeamsMessage"],
          requestId: p.correlationId,
          locationInfo: { timeZoneOffset, timeZone },
          locale,
          messageType: "Chat",
          experienceType: "Default",
          adaptiveCards: [],
          clientPreferences: {},
          connectedFederatedConnections: ["dummyId"],
        },
        plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }],
        isSbsSupported: true,
        tone: "Magic",
        renderReferencesBehindEOS: true,
        disconnectBehavior: "continue",
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
  /** これまでの累積本文(スナップショット。存在する場合は「置き換え」)。 */
  fullText?: string;
  /** 差分本文(writeAtCursor。存在する場合は「末尾へ追記」)。 */
  appendText?: string;
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

  const appendText = extractAppend(f);
  const fullText = extractText(f);
  const done = isTerminal(f);
  const err = extractError(f);
  return { fullText, appendText, done, isPing: false, error: err };
}

/**
 * writeAtCursor(逐次差分)を取り出す。substrate はスナップショット(messages[].text)と
 * 差分(writeAtCursor)を混在させて送るため、差分は「追記」として別に扱う。
 */
function extractAppend(f: Record<string, any>): string | undefined {
  const args = f?.arguments;
  if (!Array.isArray(args)) {
    return undefined;
  }
  for (const c of args) {
    if (typeof c?.writeAtCursor === "string" && c.writeAtCursor.length > 0) {
      return c.writeAtCursor;
    }
  }
  return undefined;
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
