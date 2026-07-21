/**
 * ページ内(ブラウザ)で実行される WebSocket ランナー。
 * `(arg, push) => Promise<string>` の関数式。arg で接続情報とフレーム文字列を受け取り、
 * 受信フレームを push(文字列)で Node 側へ返す。
 *
 * ページ origin(m365.cloud.microsoft 等)から substrate へ接続することで、
 * 実ブラウザの UA/TLS/CSP に沿った「Web と同じ」リクエストになる。
 *
 * arg の形:
 *   { wsUrl, handshake, invocation, ping, sep, timeoutMs }
 */
export const RUN_STREAM_SCRIPT = String.raw`(arg, push) => new Promise((resolve, reject) => {
  let ws;
  // この接続は拡張側の発信。WS フックがこの URL をテンプレートとして記録しないよう印を付ける
  // (記録するのはページ自身の実接続だけにする)。new WebSocket は同期的にフックを通るため、
  // 直後に false へ戻せば安全に自分の接続だけ除外できる。
  try { globalThis.__mscopilotSelfConnecting = true; } catch (e) {}
  try { ws = new WebSocket(arg.wsUrl); } catch (e) { reject(String(e)); return; }
  finally { try { globalThis.__mscopilotSelfConnecting = false; } catch (e) {} }
  try { ws.binaryType = "arraybuffer"; } catch (e) {}
  // 中断(VSCode の Cancel)用にアクティブな WS をグローバルへ公開する。
  try { globalThis.__mscopilotWs = ws; } catch (e) {}
  let handshakeDone = false;
  let buffer = "";
  let settled = false;
  const finish = (ok, info) => {
    if (settled) return;
    settled = true;
    try { if (globalThis.__mscopilotWs === ws) globalThis.__mscopilotWs = undefined; } catch (e) {}
    try { ws.close(); } catch (e) {}
    if (ok) resolve(info || "done"); else reject(info || "ws error");
  };
  ws.onopen = () => {
    try { ws.send(arg.handshake); } catch (e) { finish(false, String(e)); }
  };
  ws.onerror = () => { push(JSON.stringify({ __transport: "error" })); };
  ws.onclose = (ev) => {
    var code = ev && typeof ev.code === "number" ? ev.code : 0;
    var reason = ev && ev.reason ? String(ev.reason) : "";
    push(JSON.stringify({ __transport: "closed", code: code, reason: reason, handshakeDone: handshakeDone }));
    finish(true, "closed");
  };
  ws.onmessage = (ev) => {
    let data = ev.data;
    if (typeof data !== "string") {
      try { data = new TextDecoder().decode(data); } catch (e) { return; }
    }
    buffer += data;
    const parts = buffer.split(arg.sep);
    buffer = parts.pop();
    for (const part of parts) {
      if (!part) continue;
      if (!handshakeDone) {
        handshakeDone = true;
        push(JSON.stringify({ __transport: "handshake", raw: part }));
        try { ws.send(arg.invocation); } catch (e) { finish(false, String(e)); return; }
        continue;
      }
      if (part.indexOf('"type":6') !== -1) {
        try { ws.send(arg.ping); } catch (e) {}
      }
      push(part);
      if (part.indexOf('"type":3') !== -1) {
        finish(true, "completed");
        return;
      }
    }
  };
  const timeoutMs = typeof arg.timeoutMs === "number" ? arg.timeoutMs : 300000;
  setTimeout(() => finish(true, "timeout"), timeoutMs);
})`;
