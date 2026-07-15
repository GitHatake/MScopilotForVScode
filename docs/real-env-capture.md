# 実環境データ採取ガイド

このドキュメントは、**ログイン済みの M365 Copilot(BizChat)実テナント**から、拡張の
推測部分を実測値へ置き換えるために必要なデータを採取する手順です。順番に実施し、最後の
[渡すもの チェックリスト](#5-渡すもの-チェックリスト)の内容を共有してください。

採取するのは大きく 3 つです。

| # | データ | 解決する未知数 | 主な手段 |
|---|--------|----------------|----------|
| A | トークンのメタデータ(`aud`/`scp`/`target`/`exp`) | トークン選別ヒットの調整 | 拡張の診断コマンド or DevTools スニペット |
| B | WebSocket の接続 URL と送受信フレーム | エンドポイント/ハンドシェイク/type4 ペイロード/conversationId | DevTools コンソールの WS タップ |
| C | ページ origin・CSP・クローズコード | ページ内発信の可否切り分け | B の副産物 + 補助スニペット |

---

## 0. 事前準備

1. **Windows 10 + Microsoft Edge**(実運用と同じ環境)。
2. Edge で **M365 Copilot にサインイン済み**にし、`https://m365.cloud.microsoft/chat`
   (またはあなたの環境で Copilot チャットが開ける URL)を開く。
3. Copilot のチャット画面で、実際に **短いメッセージを1回送って応答が返る**ことを確認する
   (例: `hello`)。※ これでトークンが発行され、WebSocket も張られます。
4. `F12` で DevTools を開ける状態にする。

> ⚠️ **秘密情報の扱い**: アクセストークン本体(JWT の `secret`、URL の `access_token=...`)は
> 生きた認証情報です。**そのまま共有しないでください。** 本ガイドのスニペットは secret を
> 出さない/自動でマスクするように作ってあります。手動採取した場合は
> [渡すもの チェックリスト](#5-渡すもの-チェックリスト)の伏字化を必ず行ってください。

---

## A. トークンのメタデータ

### A-1. 拡張の診断コマンド(推奨)

拡張をインストール/デバッグ起動している場合はこれが最も簡単です。

1. コマンドパレット(`Ctrl+Shift+P`)→ **「MS Copilot: トークン取得を診断」** を実行。
2. 出力チャンネル(自動で開きます)の内容を**すべてコピー**。
   - 「現在のページ」「検出したトークン候補 N 件(aud/scope/exp/score)」が出ます。secret は含まれません。

### A-2. DevTools スニペット(拡張なしでも可 / フォールバック)

Copilot のタブで DevTools の **Console** を開き、以下を貼り付けて実行します。
**secret は出力しません**(メタデータのみ)。

```js
(() => {
  const dec = (jwt) => {
    try {
      const b = jwt.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      return JSON.parse(decodeURIComponent(escape(atob(b))));
    } catch (e) { return {}; }
  };
  const rows = [];
  const scan = (store, name) => {
    for (let i = 0; i < store.length; i++) {
      const k = store.key(i);
      const raw = k ? store.getItem(k) : null;
      if (!raw || raw.trim().charAt(0) !== '{') continue;
      let o; try { o = JSON.parse(raw); } catch (e) { continue; }
      if (typeof o.secret !== 'string' || o.secret.split('.').length !== 3) continue;
      const p = dec(o.secret);
      rows.push({
        store: name,
        credentialType: o.credentialType || '',
        clientId: o.clientId || '',
        aud: p.aud || '',
        scp: p.scp || (Array.isArray(p.roles) ? p.roles.join(' ') : p.roles || ''),
        target: o.target || '',
        exp: p.exp ? new Date(p.exp * 1000).toISOString() : '',
      });
    }
  };
  scan(localStorage, 'local');
  scan(sessionStorage, 'session');
  console.table(rows);
  copy(JSON.stringify(rows, null, 2)); // クリップボードにも入ります
  return rows.length + ' candidates copied (secrets omitted)';
})();
```

- 画面の表(`console.table`)を確認しつつ、`copy(...)` でクリップボードに入った JSON を
  テキストとして保存してください。**secret は含まれていません。**
- どの候補が `substrate` / `sydney` を含むか(`aud`・`scp`・`target` のどれかに)に注目してください。

---

## B. WebSocket の接続 URL と送受信フレーム(最重要)

実際のプロトコルを確定させる中心データです。**DevTools コンソールに「WS タップ」を仕込んでから
メッセージを1回送る**方式が最も確実です(接続 URL のトークンは自動で伏字化されます)。

### B-1. WS タップを仕込む(推奨)

1. Copilot タブの DevTools **Console** を開く。
2. 以下を貼り付けて実行(まだメッセージは送らない)。

```js
(() => {
  if (window.__wsTap) return 'already installed';
  window.__wsTap = [];
  const OW = window.WebSocket;
  const Wrapped = function (url, protocols) {
    const safeUrl = String(url).replace(/access_token=[^&]+/i, 'access_token=***');
    const ws = protocols !== undefined ? new OW(url, protocols) : new OW(url);
    window.__wsTap.push({ dir: 'url', data: safeUrl });
    console.log('[WS open]', safeUrl);
    const origSend = ws.send.bind(ws);
    ws.send = (d) => {
      try { window.__wsTap.push({ dir: 'send', data: typeof d === 'string' ? d : '[binary]' }); } catch (e) {}
      return origSend(d);
    };
    ws.addEventListener('message', (e) => {
      window.__wsTap.push({ dir: 'recv', data: typeof e.data === 'string' ? e.data : '[binary]' });
    });
    ws.addEventListener('close', (e) => {
      window.__wsTap.push({ dir: 'close', code: e.code, reason: e.reason });
      console.log('[WS close]', e.code, e.reason);
    });
    return ws;
  };
  Wrapped.prototype = OW.prototype;
  Wrapped.CONNECTING = OW.CONNECTING; Wrapped.OPEN = OW.OPEN;
  Wrapped.CLOSING = OW.CLOSING; Wrapped.CLOSED = OW.CLOSED;
  window.WebSocket = Wrapped;
  return 'WS tap installed. Copilot で hello を送信 → その後に copy(JSON.stringify(window.__wsTap, null, 2)) を実行';
})();
```

3. **ここで Copilot に短いメッセージ(例: `hello`)を送信**し、応答が最後まで返るのを待つ。
   - もし `[WS open]` がコンソールに出ない場合、接続はページ読み込み時に張られています。
     その場合はタップを貼った後に**ページを再読込 → もう一度送信**してください
     (再読込するとタップは消えるので、再読込後に B-1 のスニペットを貼り直す)。
4. 送受信が終わったら、コンソールで次を実行して結果を取り出す。

```js
copy(JSON.stringify(window.__wsTap, null, 2));
```

- クリップボードに入った JSON を `ws-capture.json` として保存してください。
- 中身は SignalR の JSON フレーム(`0x1e` は `` として見えます)で、**接続 URL のトークンは
  `***` に伏字化済み**、フレーム本文にトークンは含まれません。
- メッセージ本文はあなたの入力/応答テキストなので、差し支えなければ `hello` のような無害な内容で。

### B-2. HAR エクスポート(B-1 が使えない場合のフォールバック)

1. DevTools **Network** タブ → フィルタで `WS` を選択。
2. Copilot にメッセージを1回送る。
3. 一覧に出た **WebSocket の接続行を右クリック → Copy → Copy as HAR**(その1リクエストのみ)。
4. 保存した HAR を開き、**`access_token=` に続くトークン文字列を `***` に置換**してから共有
   (`_webSocketMessages` にフレーム、`request.url` に接続 URL が入っています)。

---

## C. origin / CSP / クローズコード

多くは B の副産物として得られますが、補助的に以下も採ってください。

1. **origin と現在 URL**(コンソールで):
   ```js
   console.log(location.href, '|', location.origin);
   ```
2. **CSP の確認**: B-1 のタップで **substrate への WS 接続が成功している**なら、その origin の CSP は
   既に substrate への接続を許可しています(=拡張のページ内発信が成立する前提が満たされる)。
   参考までにレスポンスヘッダの CSP を見るには Network タブでドキュメント(最初の HTML)を選び、
   **Headers → Response Headers** の `content-security-policy` を確認。
3. **異常時のクローズコード**: 接続が即切れる/応答が返らない場合、B の `close` エントリの
   `code`/`reason`(例: `1006`, `1008` 等)を控えてください。認証やポリシー起因の切り分けに使います。

---

## 4. 補足: どの未知数がどのデータで埋まるか

- **A(トークン)** → 選別ヒット `CHAT_HINTS`/`SUBSTRATE_HINTS`(`src/bridge/browserSession.ts`)を実測 `aud`/`scp`/`target` に合わせる。
- **B の接続 URL** → `buildWsUrl`(`src/copilot/protocol.ts`)のホスト/パス/クエリ、`negotiate` の要否。
- **B の `send` フレーム** → `buildInvocation` の type4 ペイロード実体(`source`/`scenario`/`optionsSets`/`allowedMessageTypes`)。
- **B の `recv` フレーム** → `interpretFrame`/`extractText` の応答本文の取り出し方、type1/2/3 の使われ方、`conversationId` の採番元。
- **C** → ページ内発信(CSP)の可否と、失敗時の原因切り分け。

---

## 5. 渡すもの チェックリスト

以下をまとめて共有してください(**すべてトークン本体を含まない**ことを確認)。

- [ ] **A**: 診断コマンドのログ全文 **または** A-2 スニペットの JSON(secret 無し)
- [ ] **B**: `ws-capture.json`(B-1)**または** 伏字化済み HAR(B-2)
  - [ ] 接続 URL の `access_token` が `***` になっている
  - [ ] `send` フレーム(type4)が1件以上含まれる
  - [ ] `recv` フレーム(type1/2/3)が含まれ、応答本文が確認できる
  - [ ] `close` の `code`/`reason`(あれば)
- [ ] **C**: `location.href` / `location.origin`、必要なら CSP ヘッダ、異常時のクローズコード
- [ ] 使った **startUrl** と、`endpointVariant`(`office`/`cloud`)のどちらで通ったか(分かれば)

### 伏字化の最終確認(共有前に必ず)

- [ ] `access_token=` の後ろが実トークンのまま残っていない(`***` になっている)
- [ ] `secret` / `Authorization` / `Cookie` 等の値を貼っていない
- [ ] JWT らしい `xxxxx.yyyyy.zzzzz` の長い文字列を貼っていない

これらが揃えば、`protocol.ts` の推測箇所を実測値へ置き換え、選別ヒットも実テナントに合わせられます。
