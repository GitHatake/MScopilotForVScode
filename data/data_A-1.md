[2026-07-16T01:17:37.387Z] [info] MS Copilot Chat 拡張を起動しました
[2026-07-16T01:17:44.272Z] [info] CDP: page ready
[2026-07-16T01:17:44.294Z] [info] === トークン取得診断 ===
[2026-07-16T01:17:44.294Z] [info] 現在のページ: https://m365.cloud.microsoft/chat
[2026-07-16T01:17:44.294Z] [info] 検出したトークン候補 1 件(★=採用候補, ○=substrate):
     [0] score=0 type=IdToken store=sessionStorage
        aud=a2760c41-63c9-42b5-8d58-bfa1fd9e2eb3
        scope=
        exp=2026-07-16T00:48:40.000Z (失効)
[2026-07-16T01:17:44.294Z] [warn] → substrate/sydney に一致する有効なトークンが見つかりませんでした。 上の一覧に候補がある場合は、その aud / scope を browserSession.ts の CHAT_HINTS / SUBSTRATE_HINTS に追加してください。
[2026-07-16T02:50:54.164Z] [warn] token 取得失敗。現在のページ: https://m365.cloud.microsoft/chat
[2026-07-16T02:50:54.164Z] [warn] 検出したトークン候補 1 件:
     [0] score=0 type=IdToken store=sessionStorage
        aud=a2760c41-63c9-42b5-8d58-bfa1fd9e2eb3
        scope=
        exp=2026-07-16T00:48:40.000Z (失効)
[2026-07-16T02:50:54.166Z] [error] signIn error M365 Copilot の認証トークンを取得できませんでした。
BrowserSessionError: M365 Copilot の認証トークンを取得できませんでした。
	at CdpBridge.getToken (c:\Users\htyk26\Documents\project\MScopilotForVScode\dist\extension.js:31886:11)
	at process.processTicksAndRejections (node:internal/process/task_queues:104:5)
	at async c:\Users\htyk26\Documents\project\MScopilotForVScode\dist\extension.js:33037:27
[2026-07-16T02:51:43.538Z] [info] === トークン取得診断 ===
[2026-07-16T02:51:43.538Z] [info] 現在のページ: https://m365.cloud.microsoft/chat
[2026-07-16T02:51:43.538Z] [info] 検出したトークン候補 1 件(★=採用候補, ○=substrate):
     [0] score=0 type=IdToken store=sessionStorage
        aud=a2760c41-63c9-42b5-8d58-bfa1fd9e2eb3
        scope=
        exp=2026-07-16T03:51:00.000Z
[2026-07-16T02:51:43.538Z] [warn] → substrate/sydney に一致する有効なトークンが見つかりませんでした。 上の一覧に候補がある場合は、その aud / scope を browserSession.ts の CHAT_HINTS / SUBSTRATE_HINTS に追加してください。