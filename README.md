# MS Copilot Chat (VSCode 拡張)

社内契約の **Microsoft 365 Copilot Chat(企業向け / BizChat)** を、Web UI ではなく
VSCode 上から GitHub Copilot Chat のように利用する拡張機能です。ブラウザのログイン済み
セッションに相乗りし、Web フロントと同じ substrate リクエストをプログラムから直接送ります。

> ⚠️ **注意**: これは非公式・未文書のプロトコルの再現です。Microsoft 側の仕様変更で動作
> しなくなる可能性があり、社内アカウントの利用規約・異常検知の観点でもリスクがあります。
> エンドポイントやフレーム構造は設定・コードで調整できるようにしてあります。

## 機能

- VSCode の Chat ビューで `@mscopilot` に話しかけてチャット(逐次ストリーム表示)
- `#file` / `#folder` でファイル・フォルダをメンションして参照
- `/resume` でローカル保存した過去のチャットを一覧・復元
- 認証はブラウザのログイン済み M365 セッションから取得(Entra / 企業テナント想定)

## 動作要件

- Windows 10、Microsoft Edge(または Chrome)
- ブラウザで M365 Copilot(`https://m365.cloud.microsoft/chat` 等)にサインイン済み
- Node.js 18+(ビルド用)、VSCode 1.95+

## セットアップ / 実行

```bash
npm install
npm run compile
```

VSCode でこのフォルダを開き、`F5`(Run Extension)で拡張開発ホストを起動します。
Chat ビューを開き `@mscopilot` を選んで質問してください。

### 認証(接続方式)

既定は **CDP** 方式です。接続の流れ:

1. `mscopilot.debugPort`(既定 9222)でリモートデバッグ中のブラウザがあれば、それに接続。
2. 無ければ、`mscopilot.browserPath`(未指定なら既定パスを自動探索)のブラウザを
   `--remote-debugging-port` 付きで起動します。

**ログイン済みの実プロファイルに相乗りしたい場合**(SSO / 条件付きアクセスを通すのに有利):

- Edge を完全に終了してから、以下のように自分でデバッグポート付き起動しておくのが最も確実です。
  ```
  msedge.exe --remote-debugging-port=9222 --user-data-dir="C:\Users\<you>\AppData\Local\Microsoft\Edge\User Data"
  ```
  もしくは `mscopilot.userDataDir` に上記プロファイルパスを設定します(Edge は要終了)。
- 専用プロファイル(既定)を使う場合は、起動したブラウザで一度 M365 Copilot にサインインしてください。

うまくいかない場合はコマンド **「MS Copilot: ブラウザ接続を初期化 / サインイン確認」** で接続を検証、
**「MS Copilot: ログを表示」** で詳細ログを確認できます。

### トークンが取得できないとき(トラブルシューティング)

「認証トークンを取得できませんでした」と出る場合は、コマンド
**「MS Copilot: トークン取得を診断」** を実行してください。2 系統の取得経路の状態を
**ログ出力** します(現在のページ URL、MSAL 候補の一覧、ページ実接続からの採取可否。
secret は出しません)。

トークンの取得経路は 3 つあり、上から順に試します。**テナントによっては access_token が
web ストレージに永続化されない**(`cacheLocation: memoryStorage` 等)ため、実測では 2) が主経路です。

1. **MSAL キャッシュ走査**: local/sessionStorage の候補を `target`/`aud`/`scp` で照合(`pickSydneyToken`)。
2. **substrate HTTP の Bearer 採取**: `fetch` / `XMLHttpRequest` をフックし、ページが substrate へ
   送る `Authorization: Bearer` を採取します(`tokenFromSecret`)。チャット接続を待たずに拾えるため、
   **対象ページを開いて数秒待つだけ**で取得できることが多いです。
3. **ページ実接続の WebSocket 採取**: `window.WebSocket` をフックし、実接続 URL から `access_token` を
   採取(`parseWsUrlToken`)。実接続が張られていれば最も確実(URL テンプレートも同時に得られます)。

診断ログの読み方(「取得経路の状態」に 3 経路の OK / 未捕捉が出ます):

- **3 経路すべて「なし / 未捕捉」**: 対象ブラウザで Copilot(`mscopilot.startUrl`)を開いた状態で
  **数秒待って**再実行してください。それでも駄目なら**一度メッセージを送る**と実接続が張られます。
- **MSAL 候補が 0 件 / IdToken のみ**: そのテナントは access_token をストレージに残さないタイプです
  (正常)。2) か 3) が OK なら発信できます。
- **候補はあるが ★/○ が付かない(score 0)**: 一覧の `aud` / スコープを確認し、
  `src/bridge/browserSession.ts` の `CHAT_HINTS` / `SUBSTRATE_HINTS` に該当語を追加すると拾えます。
- **「残り約 N 分」が短い / 失効**: トークン寿命(約 60 分)切れです。ブラウザで Copilot を操作するか、
  拡張側の再取得(`getToken(force)`)で更新されます。

> 発信 URL は、ページ実接続を捕捉できた場合はそのテンプレートを流用します
> (`deriveWsUrlFromTemplate`)。実テナントの `variants` / `scenario` / `access_token` を
> そのまま使い、セッション ID と `ConversationId` だけ差し替えるため、Microsoft 側の
> パラメータ変更にも強くなります。捕捉できない場合は実測値ベースの `buildWsUrl` を使います。

### フォールバック(Playwright)

企業ポリシーで Edge のリモートデバッグが禁止されている場合は Playwright 方式を使えます。

```bash
npm i playwright
npx playwright install msedge
```

設定 `mscopilot.transport` を `playwright` に変更してください(専用ブラウザで初回サインインが必要)。

## 設定項目

| 設定 | 既定 | 説明 |
| --- | --- | --- |
| `mscopilot.transport` | `cdp` | 接続方式(`cdp` / `playwright`) |
| `mscopilot.endpointVariant` | `office` | substrate エンドポイント系統(`office` / `cloud`) |
| `mscopilot.browserPath` | (自動) | Edge/Chrome の実行ファイルパス |
| `mscopilot.debugPort` | `9222` | CDP デバッグポート |
| `mscopilot.userDataDir` | (専用) | ブラウザのユーザーデータディレクトリ |
| `mscopilot.startUrl` | `https://m365.cloud.microsoft/chat` | トークン取得に使う URL |
| `mscopilot.maxMentionBytes` | `131072` | メンション注入の最大バイト数 |

## 実機での調整が必要な箇所(重要)

> 📄 実環境からの**データ採取手順**は [`docs/real-env-capture.md`](docs/real-env-capture.md) に
> まとめてあります(トークンのメタデータ、WebSocket の URL/フレーム、CSP/クローズコードを
> secret を出さずに採る方法)。まずこれに沿ってデータを採取すると、下記の調整を実測値で行えます。

プロトコル値は **2026/7 の実トラフィック(officeweb / `m365.cloud.microsoft`)を実測**して反映済みです。
それでも Microsoft 側の変更で調整が要る場合、以下を実機のブラウザ DevTools で確認してください。
ログには送受信フレームがそのまま出ます。

- **`src/copilot/protocol.ts`**(実測反映済み)
  - `buildWsUrl`: `wss://substrate.office.com/m365Copilot/Chathub/{oid}@{tid}?…`(実測パス/クエリ)。
    ただし通常は下記テンプレート流用が優先されます。
  - `deriveWsUrlFromTemplate`: ページ実接続 URL を流用し session/ConversationId のみ差し替え(推奨経路)。
  - `buildInvocation`: type4 ペイロード(`source:"officeweb"` / `optionsSets` / `tone:"Magic"` 等)。
  - `interpretFrame` / `extractText` / `extractAppend`: `messages[].text`(スナップショット)と
    `writeAtCursor`(差分)の両対応。完了は type2/3。
- **`src/copilot/injectedClient.ts`**: ページ内 WS ランナー。ハンドシェイクや ping 応答の扱い。
- **`src/bridge/browserSession.ts`**
  - `WS_HOOK_SCRIPT`: `WebSocket` / `fetch` / `XMLHttpRequest` をフックし、substrate の
    `access_token`(WS URL)と `Authorization: Bearer`(HTTP)を採取(主経路)。
  - `tokenFromSecret` / `parseWsUrlToken`: 採取した JWT を SubstrateToken 化。
  - `CHAT_HINTS` / `SUBSTRATE_HINTS`: MSAL 走査でのトークン選別語(補助経路)。
- **`mscopilot.startUrl`**: 相乗りするページ origin(実測 `https://m365.cloud.microsoft/chat`。
  CSP が substrate への接続を許可し、かつその origin で Copilot 実接続が張られる必要あり)。

## アーキテクチャ

```
VSCode拡張(TS)  ──▶  BrowserSession  ──CDP/Playwright──▶  実ブラウザ(M365ログイン済)
  Chat Participant     (cdpBridge /                        ページ内で substrate WS を発信
  #file 解決 / /resume   playwrightBridge)                 (injectedClient)
  ローカル履歴          SubstrateClient(フレーム構築/解析: protocol.ts)
```

- 認証: 2 系統。(1) MSAL キャッシュ(local/sessionStorage)から候補を収集し
  `target`/`aud`/`scp` で照合(`pickSydneyToken`)。(2) それが空/失効なら、`window.WebSocket` を
  フックしてユーザー自身の Copilot 実接続 URL から `access_token` を採取(`parseWsUrlToken`)。
  **実テナントでは access_token がストレージに永続化されないことがあり、(2) が主経路**になる。
  取得失敗時は両経路の状態を診断ログへ出力する。
- 発信: WebSocket 接続と送受信は**ブラウザページ内**で実行し、実 UA/TLS/CSP に沿わせる。
  接続 URL は捕捉した実接続テンプレートを流用(`deriveWsUrlFromTemplate`)し、実 `variants` や
  `access_token` をそのまま使う。フレームは binding 経由で Node 側へ返し、`protocol.ts` が
  `messages[].text`(スナップショット)と `writeAtCursor`(差分)を解釈する。

## 開発

```bash
npm run typecheck     # 型チェック
npm run compile       # esbuild バンドル
npm run watch         # ウォッチビルド
npm run test:offline  # ブラウザ無しで純ロジックを検証
```

`npm run test:offline` はトークン選別・URL 構築・SignalR フレームの構築/解析/解釈など、
ネットワークやブラウザ無しで検証できる部分を実行します。
