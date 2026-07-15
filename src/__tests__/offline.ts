/* オフライン検証: ブラウザ/ネットワーク無しで純ロジックを実証する。
 * `npm run test:offline` でビルドして実行する。 */
import assert from "node:assert";
import {
  pickSydneyToken,
  inventoryTokens,
  COLLECT_TOKENS_SCRIPT,
  RawTokenCandidate,
} from "../bridge/browserSession";
import { RUN_STREAM_SCRIPT } from "../copilot/injectedClient";
import {
  buildWsUrl,
  buildInvocation,
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

console.log("WebSocket URL:");
ok("office/cloud で正しいホストとクエリを組む", () => {
  const token = { accessToken: "TKN", objectId: "OID", tenantId: "TID", expiresAt: Date.now() + 1e6 };
  const office = buildWsUrl("office", token, "CONV-1");
  assert.ok(office.url.startsWith("wss://substrate.office.com/m365chat/SecuredChathub/OID@TID?"));
  assert.ok(office.url.includes("ConversationId=CONV-1"));
  assert.ok(office.url.includes("access_token=TKN"));
  const cloud = buildWsUrl("cloud", token);
  assert.ok(cloud.url.startsWith("wss://substrate.svc.cloud.microsoft/m365Copilot/Chathub/OID@TID?"));
});

console.log("SignalR フレーム:");
ok("handshake は 0x1e 終端", () => {
  const h = handshakeFrame();
  assert.ok(h.endsWith(RECORD_SEP));
  assert.deepEqual(JSON.parse(h.slice(0, -1)), { protocol: "json", version: 1 });
});
ok("invocation は type4 target=chat", () => {
  const inv = buildInvocation({ prompt: "こんにちは", conversationId: "C", isStartOfSession: true, invocationId: "0" });
  assert.equal((inv as any).type, 4);
  assert.equal((inv as any).target, "chat");
  assert.equal((inv as any).arguments[0].message.text, "こんにちは");
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
ok("completion(type3)で done", () => {
  assert.equal(interpretFrame({ type: 3, invocationId: "0" }).done, true);
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
