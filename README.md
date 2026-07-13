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

このプロトコルは実トラフィックでの検証が前提です。以下は実機のブラウザ DevTools(Network → WS)で
実際のフレームを確認し、必要に応じて調整してください。ログには送受信フレームがそのまま出ます。

- **`src/copilot/protocol.ts`**
  - `buildWsUrl`: エンドポイントのパス/クエリ。`endpointVariant` で 2 系統を切替。
  - `buildInvocation`: type4 ペイロードの `source` / `scenario` / `optionsSets` / `message` 形。
  - `interpretFrame` / `extractText`: 応答本文の取り出し方(messages 配列の想定)。
  - 完了判定に使う `type`(2/3)。
- **`src/copilot/injectedClient.ts`**: ページ内 WS ランナー。ハンドシェイクや ping 応答の扱い。
- **`mscopilot.startUrl`**: 相乗りするページ origin(CSP が substrate への接続を許可する必要あり)。

## アーキテクチャ

```
VSCode拡張(TS)  ──▶  BrowserSession  ──CDP/Playwright──▶  実ブラウザ(M365ログイン済)
  Chat Participant     (cdpBridge /                        ページ内で substrate WS を発信
  #file 解決 / /resume   playwrightBridge)                 (injectedClient)
  ローカル履歴          SubstrateClient(フレーム構築/解析: protocol.ts)
```

- 認証: ログイン済みセッションの MSAL キャッシュ(local/sessionStorage)から
  audience `.../sydney` のアクセストークンを読み取る(無ければ substrate ホストの
  トークンにフォールバック。`browserSession.ts` の `pickSydneyToken`)。
- 発信: WebSocket 接続と送受信は**ブラウザページ内**で実行し、実 UA/TLS/CSP に沿わせる。
  フレームは binding 経由で Node 側へ返し、`protocol.ts` が解釈する。

## 開発

```bash
npm run typecheck     # 型チェック
npm run compile       # esbuild バンドル
npm run watch         # ウォッチビルド
npm run test:offline  # ブラウザ無しで純ロジックを検証
```

`npm run test:offline` はトークン選別・URL 構築・SignalR フレームの構築/解析/解釈など、
ネットワークやブラウザ無しで検証できる部分を実行します。
