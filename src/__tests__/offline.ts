/* オフライン検証: ブラウザ/ネットワーク無しで純ロジックを実証する。
 * `npm run test:offline` でビルドして実行する。 */
import assert from "node:assert";
import {
  pickSydneyToken,
  pickBearerToken,
  pickBearerTokenScored,
  scoreToken,
  isCopilotWsUrl,
  describeBearers,
  formatTokenScope,
  inventoryTokens,
  parseWsUrlToken,
  tokenFromSecret,
  COLLECT_TOKENS_SCRIPT,
  WS_HOOK_SCRIPT,
  RawTokenCandidate,
} from "../bridge/browserSession";
import { RUN_STREAM_SCRIPT } from "../copilot/injectedClient";
import { ResponseAssembler } from "../copilot/responseAssembler";
import {
  buildWsUrl,
  buildInvocation,
  deriveWsUrlFromTemplate,
  encodeFrame,
  handshakeFrame,
  parseFrames,
  interpretFrame,
  extractConversationId,
  RECORD_SEP,
} from "../copilot/protocol";

function b64url(obj: unknown): string {
  return Buffer.from(JSON.stringify(obj)).toString("base64url");
}
function jwt(payload: unknown): string {
  return `${b64url({ alg: "RS256", typ: "JWT" })}.${b64url(payload)}.sig`;
}

let passed = 0;
function ok(name: string, fn: () => void) {
  fn();
  passed++;
  console.log(`  ✓ ${name}`);
}

console.log("token 選別:");
ok("sydney audience のトークンを oid/tid/exp 付きで選ぶ", () => {
  const exp = 1893456000; // 2030
  const sydney = jwt({ aud: "https://substrate.office.com/sydney", oid: "OID-1", tid: "TID-1", exp });
  const graph = jwt({ aud: "https://graph.microsoft.com", oid: "OID-1", tid: "TID-1", exp });
  const token = pickSydneyToken([graph, sydney, "garbage", ""]);
  assert.ok(token, "トークンが選ばれること");
  assert.equal(token!.objectId, "OID-1");
  assert.equal(token!.tenantId, "TID-1");
  assert.equal(token!.expiresAt, exp * 1000);
  assert.equal(token!.accessToken, sydney);
});
ok("sydney が無ければ undefined", () => {
  const graph = jwt({ aud: "https://graph.microsoft.com", oid: "o", tid: "t", exp: 1893456000 });
  assert.equal(pickSydneyToken([graph]), undefined);
});
ok("複数 sydney は exp が新しい方を選ぶ", () => {
  const older = jwt({ aud: ".../sydney", oid: "o", tid: "t", exp: 1000 });
  const newer = jwt({ aud: ".../sydney", oid: "o", tid: "t", exp: 2000 });
  assert.equal(pickSydneyToken([older, newer])!.expiresAt, 2000 * 1000);
});
ok("sydney 不在時は substrate ホストのトークンにフォールバック", () => {
  const sub = jwt({ aud: "https://substrate.svc.cloud.microsoft/", oid: "o", tid: "t", exp: 3000 });
  const graph = jwt({ aud: "https://graph.microsoft.com", oid: "o", tid: "t", exp: 9999 });
  const picked = pickSydneyToken([graph, sub]);
  assert.ok(picked);
  assert.equal(picked!.accessToken, sub);
});
ok("sydney があれば substrate より優先(同じ有効性)", () => {
  const sub = jwt({ aud: "https://substrate.svc.cloud.microsoft/", oid: "o", tid: "t", exp: 9999 });
  const syd = jwt({ aud: "https://substrate.office.com/sydney", oid: "o", tid: "t", exp: 1 });
  assert.equal(pickSydneyToken([sub, syd])!.accessToken, syd);
});
const FUTURE = Math.floor(Date.now() / 1000) + 3600; // 1時間後
const PAST = Math.floor(Date.now() / 1000) - 3600; // 1時間前
ok("aud が GUID でも target(スコープ)に sydney があれば拾う", () => {
  // 実テナントでは aud が GUID のことが多い。scope 側の語で判別できることを確認。
  const secret = jwt({ aud: "00000003-0000-0000-c000-000000000000", oid: "o", tid: "t", exp: FUTURE });
  const cand: RawTokenCandidate = {
    secret,
    target: "https://substrate.office.com/sydney.readwrite openid profile",
    credentialType: "AccessToken",
  };
  const token = pickSydneyToken([cand]);
  assert.ok(token, "target 経由で拾えること");
  assert.equal(token!.accessToken, secret);
});
ok("有効なトークンを失効済みより優先(スコアが下でも)", () => {
  const expiredSydney = jwt({ aud: ".../sydney", oid: "o", tid: "t", exp: PAST });
  const validSubstrate = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const picked = pickSydneyToken([expiredSydney, validSubstrate]);
  assert.equal(picked!.accessToken, validSubstrate, "失効した sydney より有効な substrate");
});
ok("scp に sydney スコープがあれば aud 非 substrate でも拾う", () => {
  const secret = jwt({ aud: "api://guid", scp: "M365Chat.ReadWrite", oid: "o", tid: "t", exp: FUTURE });
  const token = pickSydneyToken([{ secret, credentialType: "AccessToken_With_AuthScheme" }]);
  assert.ok(token, "scp の m365chat で拾えること");
});
ok("inventoryTokens は候補を score 付きで一覧化する", () => {
  const syd = jwt({ aud: ".../sydney", oid: "o", tid: "t", exp: FUTURE });
  const graph = jwt({ aud: "https://graph.microsoft.com", oid: "o", tid: "t", exp: FUTURE });
  const inv = inventoryTokens([graph, syd]);
  assert.equal(inv.length, 2);
  assert.equal(inv[0].score, 3, "sydney が先頭(score 降順)");
  assert.equal(inv[1].score, 0, "graph は対象外 score0");
});
ok("COLLECT_TOKENS_SCRIPT は local/sessionStorage を走査し候補(メタ付き)を返す", () => {
  const mkStore = (d: Record<string, string>) => ({
    _d: d,
    get length() {
      return Object.keys(this._d).length;
    },
    key(i: number) {
      return Object.keys(this._d)[i];
    },
    getItem(k: string) {
      return this._d[k];
    },
  });
  const jwtLike = "aaaa.bbbb.cccc"; // JWT らしさ(ドット2つ・十分な長さ)
  const s1 = jwtLike + "1111111111111111111111111111";
  const s2 = jwtLike + "2222222222222222222222222222";
  const s3 = jwtLike + "3333333333333333333333333333";
  (globalThis as any).window = {
    localStorage: mkStore({
      k1: JSON.stringify({ credentialType: "AccessToken", secret: s1, target: "scope-a" }),
      k2: JSON.stringify({ credentialType: "AccessToken", secret: "no-dots-short" }), // JWT でない → 除外
      k3: "not-json",
    }),
    sessionStorage: mkStore({
      s1: JSON.stringify({ credentialType: "AccessToken_With_AuthScheme", secret: s3, target: "scope-c" }),
    }),
  };
  const result = eval(COLLECT_TOKENS_SCRIPT) as RawTokenCandidate[];
  const secrets = result.map((r) => r.secret).sort();
  assert.deepEqual(secrets, [s1, s3].sort(), "JWT らしい secret のみ収集");
  const a = result.find((r) => r.secret === s1)!;
  assert.equal(a.target, "scope-a");
  assert.equal(a.store, "localStorage");
  const c = result.find((r) => r.secret === s3)!;
  assert.equal(c.credentialType, "AccessToken_With_AuthScheme");
  assert.equal(c.store, "sessionStorage");
  void s2;
  delete (globalThis as any).window;
});

console.log("Bearer 選別(スコープ別):");
ok("複数 substrate Bearer から Copilot(Chathub)用を URL で選ぶ", () => {
  // substrate は同一ホストで多数サービスを提供。URL とスコープで Copilot 用を見分ける。
  const copilot = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const optics = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const picked = pickBearerToken([
    { t: optics, u: "https://substrate.office.com/optics/v2/upload" },
    { t: copilot, u: "https://substrate.office.com/m365Copilot/conversations" },
  ]);
  assert.ok(picked, "選ばれること");
  assert.equal(picked!.accessToken, copilot, "m365Copilot URL のトークンを優先(score3)");
});
ok("scp が sydney の Bearer を URL 非Copilotでも拾う", () => {
  const secret = jwt({ aud: "api://guid", scp: "M365Chat.ReadWrite", oid: "o", tid: "t", exp: FUTURE });
  const picked = pickBearerToken([{ t: secret, u: "https://substrate.office.com/search" }]);
  assert.ok(picked);
  assert.equal(picked!.accessToken, secret);
});
ok("文字列だけの Bearer リストも受け付ける(後方互換)", () => {
  const secret = jwt({ aud: "https://substrate.office.com/sydney", oid: "o", tid: "t", exp: FUTURE });
  assert.equal(pickBearerToken([secret])!.accessToken, secret);
});
ok("空/不正な Bearer リストは undefined", () => {
  assert.equal(pickBearerToken([]), undefined);
  assert.equal(pickBearerToken([{ t: "", u: "x" }]), undefined);
});
ok("formatTokenScope は aud/scp/score を出す(secret は含めない)", () => {
  const secret = jwt({ aud: "https://substrate.office.com/sydney", scp: "Chat.RW", oid: "o", tid: "t", exp: FUTURE });
  const s = formatTokenScope({ accessToken: secret, objectId: "o", tenantId: "t", expiresAt: FUTURE * 1000 });
  assert.ok(s.includes("aud=https://substrate.office.com/sydney"));
  assert.ok(s.includes("scp=Chat.RW"));
  assert.ok(s.includes("score=3"));
  assert.ok(!s.includes(secret), "secret を漏らさない");
});
ok("describeBearers は score 降順で整形し secret を漏らさない(feedback_5 の可視化)", () => {
  const search = jwt({ aud: "https://substrate.office.com/search", scp: "SubstrateSearch-Internal.ReadWrite", oid: "o", tid: "t", exp: FUTURE });
  const copilot = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const out = describeBearers([
    { t: search, u: "https://substrate.office.com/search/api", via: "req-header" },
    { t: copilot, u: "https://substrate.office.com/m365Copilot/conversations", via: "ws-url" },
  ]);
  // score3(★, Copilot)が先頭、score2(○, search)が後。
  assert.ok(out.indexOf("★") < out.indexOf("○"), "score 降順(Copilot が先)");
  assert.ok(out.includes("via=ws-url"), "採取経路を出す");
  assert.ok(out.includes("scp=SubstrateSearch-Internal.ReadWrite"), "scp を出す");
  assert.ok(!out.includes(search) && !out.includes(copilot), "secret を漏らさない");
});
ok("describeBearers は候補ゼロでも安全なメッセージを返す", () => {
  assert.ok(describeBearers([]).includes("1 件も採取していません"));
});
ok("scoreToken は採取元 URL 込みで Copilot を score3 と判定(feedback_6 の核心)", () => {
  // claim だけでは substrate 一般(score2)でも、Chathub URL 由来なら Copilot(score3)。
  const tok = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const wsUrl = "wss://substrate.office.com/m365Copilot/Chathub/o@t?access_token=x";
  assert.equal(scoreToken({ accessToken: tok, objectId: "o", tenantId: "t", expiresAt: FUTURE * 1000 }), 2);
  assert.equal(scoreToken({ accessToken: tok, objectId: "o", tenantId: "t", expiresAt: FUTURE * 1000 }, wsUrl), 3);
});
ok("pickBearerTokenScored は選んだトークンのスコアと採取元URLを返す", () => {
  const search = jwt({ aud: "https://substrate.office.com/search", oid: "o", tid: "t", exp: FUTURE });
  // /search のみ → score2(Copilot ではない=送信前に掴んでも Language model unavailable)。
  const only = pickBearerTokenScored([{ t: search, u: "https://substrate.office.com/search/api" }]);
  assert.ok(only);
  assert.equal(only!.score, 2);
  // Copilot URL の Bearer が加わると score3 の方を選ぶ。
  const copilot = jwt({ aud: "https://substrate.office.com", oid: "o", tid: "t", exp: FUTURE });
  const both = pickBearerTokenScored([
    { t: search, u: "https://substrate.office.com/search/api" },
    { t: copilot, u: "https://substrate.office.com/m365Copilot/Chathub/o@t?access_token=x" },
  ]);
  assert.equal(both!.score, 3);
  assert.equal(both!.token.accessToken, copilot);
});
ok("isCopilotWsUrl は Chathub 実接続 URL のみ true", () => {
  assert.ok(isCopilotWsUrl("wss://substrate.office.com/m365Copilot/Chathub/o@t?access_token=x"));
  assert.ok(!isCopilotWsUrl("wss://substrate.office.com/search?foo=1"));
  assert.ok(!isCopilotWsUrl("wss://substrate.office.com/m365Copilot/Chathub/o@t")); // token 無しは対象外
});

console.log("応答組み立て / 失敗検出:");
ok("スナップショット差分を正しく onDelta へ流す", () => {
  const out: string[] = [];
  const a = new ResponseAssembler((d) => out.push(d));
  a.setSnapshot("Hello");
  a.setSnapshot("Hello world");
  a.end();
  assert.equal(out.join(""), "Hello world");
  assert.equal(a.text, "Hello world");
  assert.equal(a.produced, true);
  assert.equal(a.failure, undefined);
});
ok("writeAtCursor 差分を追記できる", () => {
  const out: string[] = [];
  const a = new ResponseAssembler((d) => out.push(d));
  a.setSnapshot("Hello! How can");
  a.appendDelta(" I help you today");
  a.appendDelta("?");
  a.end();
  assert.equal(out.join(""), "Hello! How can I help you today?");
});
ok("分割到着した失敗フレーズを本文に出さず error 化する(feedback_4 の症状)", () => {
  const out: string[] = [];
  const a = new ResponseAssembler((d) => out.push(d));
  a.appendDelta("Language model");
  a.appendDelta(" unavailable");
  a.end();
  assert.equal(out.length, 0, "本文は一切出さない");
  assert.equal(a.produced, false, "produced=false でリトライ対象");
  assert.equal(a.failure, "Language model unavailable");
  assert.equal(a.text, "", "失敗は応答本文にしない");
});
ok("一括到着の失敗フレーズも error 化する", () => {
  const out: string[] = [];
  const a = new ResponseAssembler((d) => out.push(d));
  a.setSnapshot("Language model unavailable");
  a.end();
  assert.equal(out.length, 0);
  assert.equal(a.failure, "Language model unavailable");
});
ok("失敗フレーズと同じ書き出しでも本物の応答は出す", () => {
  const out: string[] = [];
  const a = new ResponseAssembler((d) => out.push(d));
  // 失敗フレーズの前方一致で一旦保留 → 分岐した時点でフラッシュ。
  a.appendDelta("Language model");
  a.appendDelta("s are neural networks.");
  a.end();
  assert.equal(a.failure, undefined);
  assert.equal(out.join(""), "Language models are neural networks.");
  assert.equal(a.produced, true);
});

console.log("WebSocket URL:");
ok("office/cloud で正しいホストとクエリを組む", () => {
  const token = { accessToken: "TKN", objectId: "OID", tenantId: "TID", expiresAt: Date.now() + 1e6 };
  const office = buildWsUrl("office", token, "CONV-1");
  assert.ok(office.url.startsWith("wss://substrate.office.com/m365Copilot/Chathub/OID@TID?"));
  assert.ok(office.url.includes("ConversationId=CONV-1"));
  assert.ok(office.url.includes("access_token=TKN"));
  assert.ok(office.url.includes("source=%22officeweb%22"), "source は officeweb");
  assert.ok(office.url.includes("scenario=OfficeWebIncludedCopilot"));
  const cloud = buildWsUrl("cloud", token);
  assert.ok(cloud.url.startsWith("wss://substrate.svc.cloud.microsoft/m365Copilot/Chathub/OID@TID?"));
});
ok("テンプレートからは token/variants を流用し session/ConversationId のみ差し替え", () => {
  const template =
    "wss://substrate.office.com/m365Copilot/Chathub/OID@TID?chatsessionid=OLD&X-SessionId=OLD-D&" +
    "ConversationId=PAGE-CONV&access_token=REALTOKEN&variants=a,b,c&source=%22officeweb%22";
  const ep = deriveWsUrlFromTemplate(template, "MY-CONV")!;
  assert.ok(ep, "テンプレートから生成できること");
  assert.ok(ep.url.includes("access_token=REALTOKEN"), "実 token を流用");
  assert.ok(ep.url.includes("variants=a,b,c"), "実 variants を流用");
  assert.ok(ep.url.includes("ConversationId=MY-CONV"), "ConversationId は差し替え");
  assert.ok(!ep.url.includes("PAGE-CONV"), "ページの ConversationId は残さない");
  assert.ok(!ep.url.includes("chatsessionid=OLD"), "session id は差し替え");
});
ok("parseWsUrlToken は URL から token を取り出す(oid/tid は path 補完)", () => {
  const secret = jwt({ exp: FUTURE }); // oid/tid を含まない token
  const url = `wss://substrate.office.com/m365Copilot/Chathub/OID-P@TID-P?access_token=${secret}&x=1`;
  const t = parseWsUrlToken(url)!;
  assert.ok(t, "取り出せること");
  assert.equal(t.accessToken, secret);
  assert.equal(t.objectId, "OID-P", "oid は path から");
  assert.equal(t.tenantId, "TID-P", "tid は path から");
  assert.equal(t.expiresAt, FUTURE * 1000);
});
ok("tokenFromSecret は oid/tid/exp を持つ JWT を SubstrateToken 化する", () => {
  const secret = jwt({ oid: "O", tid: "T", exp: FUTURE });
  const t = tokenFromSecret(secret)!;
  assert.ok(t);
  assert.equal(t.objectId, "O");
  assert.equal(t.tenantId, "T");
  assert.equal(t.expiresAt, FUTURE * 1000);
  // oid/tid が無い JWT は対象外(URL パス補完が無いと WS URL を組めないため)。
  assert.equal(tokenFromSecret(jwt({ exp: FUTURE })), undefined);
  assert.equal(tokenFromSecret(""), undefined);
});
ok("WS_HOOK_SCRIPT は有効な JS 式として解析できる", () => {
  // ブラウザ globals は無いが、正規表現エスケープ等の構文エラーは検出できる。
  assert.doesNotThrow(() => new Function(`return (${WS_HOOK_SCRIPT})`));
});

console.log("SignalR フレーム:");
ok("handshake は 0x1e 終端", () => {
  const h = handshakeFrame();
  assert.ok(h.endsWith(RECORD_SEP));
  assert.deepEqual(JSON.parse(h.slice(0, -1)), { protocol: "json", version: 1 });
});
ok("invocation は type4 target=chat(officeweb 形)", () => {
  const inv = buildInvocation({
    prompt: "こんにちは",
    isStartOfSession: true,
    invocationId: "0",
    sessionId: "SID-D",
    correlationId: "CID",
  });
  assert.equal((inv as any).type, 4);
  assert.equal((inv as any).target, "chat");
  const a = (inv as any).arguments[0];
  assert.equal(a.message.text, "こんにちは");
  assert.equal(a.source, "officeweb");
  assert.equal(a.tone, "Magic");
  assert.equal(a.sessionId, "SID-D");
  assert.equal(a.traceId, "CID");
  assert.equal(a.message.requestId, "CID");
  // conversationId は URL 側で渡すため arguments には含めない。
  assert.equal(a.conversationId, undefined);
});
ok("parseFrames は 0x1e 区切りを分解", () => {
  const buf = encodeFrame({ a: 1 }) + encodeFrame({ b: 2 }) + '{"partial":';
  const frames = parseFrames(buf);
  assert.equal(frames.length, 2); // 不完全フレームは無視
  assert.deepEqual(frames[0], { a: 1 });
});

console.log("フレーム解釈:");
ok("streaming フレームから bot 本文を抽出", () => {
  const frame = { type: 1, arguments: [{ messages: [{ author: "bot", text: "途中まで" }] }] };
  const r = interpretFrame(frame);
  assert.equal(r.fullText, "途中まで");
  assert.equal(r.done, false);
});
ok("内部メッセージ(検索クエリ等)は本文にしない", () => {
  const frame = {
    type: 1,
    arguments: [
      {
        messages: [
          { author: "bot", messageType: "InternalSearchQuery", text: "検索中…" },
          { author: "bot", messageType: "Chat", text: "本当の答え" },
        ],
      },
    ],
  };
  assert.equal(interpretFrame(frame).fullText, "本当の答え");
});
ok("内部メッセージのみなら本文なし", () => {
  const frame = {
    type: 1,
    arguments: [{ messages: [{ author: "bot", messageType: "InternalLoaderMessage", text: "…" }] }],
  };
  assert.equal(interpretFrame(frame).fullText, undefined);
});
ok("writeAtCursor は差分(appendText)として取り出す", () => {
  const frame = { type: 1, target: "update", arguments: [{ writeAtCursor: " I help you today" }] };
  const r = interpretFrame(frame);
  assert.equal(r.appendText, " I help you today");
  assert.equal(r.fullText, undefined, "差分フレームでは fullText は出さない");
});
ok("スナップショット→差分→最終スナップショットで正しく累積", () => {
  // 実測の hello 応答の流れを再現し、累積本文が壊れないことを確認。
  let text = "";
  const apply = (frame: unknown) => {
    const i = interpretFrame(frame);
    if (i.appendText) {
      text += i.appendText;
    }
    if (i.fullText !== undefined) {
      text = i.fullText.startsWith(text) ? i.fullText : text + i.fullText;
    }
  };
  const snap = (t: string) => ({ type: 1, arguments: [{ messages: [{ author: "bot", text: t }] }] });
  const delta = (t: string) => ({ type: 1, arguments: [{ writeAtCursor: t }] });
  apply(snap("Hello! How can"));
  apply(delta(" I help you today"));
  apply(delta("?"));
  apply(snap("Hello! How can I help you today? 😊"));
  assert.equal(text, "Hello! How can I help you today? 😊");
});
ok("completion(type3)で done", () => {
  assert.equal(interpretFrame({ type: 3, invocationId: "0" }).done, true);
});
ok("type2 の item.conversationId と bot 本文を拾う", () => {
  const frame = {
    type: 2,
    invocationId: "0",
    item: {
      messages: [
        { author: "user", text: "hello" },
        { author: "bot", text: "Hello! How can I help you today? 😊", turnState: "Completed" },
      ],
      conversationId: "SRV-CONV",
      result: { value: "Success" },
    },
  };
  assert.equal(extractConversationId(frame), "SRV-CONV");
  assert.equal(interpretFrame(frame).fullText, "Hello! How can I help you today? 😊");
});
ok("ping(type6)を検出", () => {
  assert.equal(interpretFrame({ type: 6 }).isPing, true);
});
ok("conversationId をフレームから拾う", () => {
  const id = extractConversationId({ type: 2, item: { conversationId: "SRV-CONV" } });
  assert.equal(id, "SRV-CONV");
});

console.log("ページ内スクリプト:");
ok("RUN_STREAM_SCRIPT が有効な JS 式として解析できる", () => {
  // ブラウザ globals は無いが構文エラーは検出できる。
  const fn = new Function(`return (${RUN_STREAM_SCRIPT})`)();
  assert.equal(typeof fn, "function");
  assert.equal(fn.length, 2); // (arg, push)
});

console.log(`\n${passed} 件のオフライン検証にすべて合格しました。`);
