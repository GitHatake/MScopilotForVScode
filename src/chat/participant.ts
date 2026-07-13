import * as vscode from "vscode";
import { randomUUID } from "node:crypto";
import { BrowserSessionError } from "../bridge";
import type { SubstrateClient } from "../copilot/substrateClient";
import { resolveReferences } from "../mentions/resolveReferences";
import { Conversation, HistoryStore, makeTitle } from "../history/store";
import { log } from "../logger";

export interface ControllerDeps {
  history: HistoryStore;
  getClient: () => SubstrateClient;
  getMaxMentionBytes: () => number;
}

/**
 * Chat Participant のハンドラ本体。request→SubstrateClient→stream 逐次出力を担う。
 * 会話は in-memory の current を保持しつつローカル履歴へ永続化する。
 */
export class ChatController {
  private current: Conversation | undefined;

  constructor(private readonly deps: ControllerDeps) {}

  readonly handle: vscode.ChatRequestHandler = async (request, context, stream, token) => {
    try {
      if (request.command === "resume") {
        await this.handleResume(stream);
        return {};
      }
      await this.handleChat(request, context, stream, token);
      return {};
    } catch (e) {
      this.renderError(stream, e);
      return { errorDetails: { message: describe(e) } };
    }
  };

  private async handleChat(
    request: vscode.ChatRequest,
    context: vscode.ChatContext,
    stream: vscode.ChatResponseStream,
    token: vscode.CancellationToken,
  ): Promise<void> {
    // 新規チャット(履歴が空)なら会話を作り直す。
    if (context.history.length === 0 || !this.current) {
      this.current = newConversation();
    }
    const conv = this.current;

    // 参照(#file/#folder)を解決してプロンプトへ注入。
    const ctx = await resolveReferences(request.references, this.deps.getMaxMentionBytes());
    for (const u of ctx.used) {
      if (u.uri) {
        stream.reference(u.uri);
      }
    }
    const prompt = ctx.text ? `${ctx.text}\n\n${request.prompt}` : request.prompt;

    stream.progress("Microsoft Copilot に問い合わせています…");

    const controller = new AbortController();
    const sub = token.onCancellationRequested(() => controller.abort());

    let received = false;
    const onDelta = (delta: string) => {
      received = true;
      stream.markdown(delta);
    };
    let result;
    try {
      result = await this.deps.getClient().ask({
        prompt,
        conversationId: conv.substrateConversationId,
        signal: controller.signal,
        onDelta,
      });
      // 応答が無くエラーだけ返った場合はトークン期限切れの可能性が高いので、
      // トークンを強制再取得して 1 回だけ再試行する。
      if (result.error && !received && !controller.signal.aborted) {
        log.info("応答なし+エラーのためトークンを再取得して再試行します");
        stream.progress("認証を更新して再試行しています…");
        result = await this.deps.getClient().ask({
          prompt,
          conversationId: conv.substrateConversationId,
          signal: controller.signal,
          forceToken: true,
          onDelta,
        });
      }
    } finally {
      sub.dispose();
    }

    if (result.error && !received) {
      stream.markdown(
        `\n\n⚠️ Copilot からの応答取得に失敗しました: \`${result.error}\`\n\n` +
          "「MS Copilot: ログを表示」で詳細を確認できます。",
      );
    }

    // 会話状態を更新して永続化。
    conv.substrateConversationId = result.conversationId;
    conv.updatedAt = Date.now();
    if (conv.messages.length === 0) {
      conv.title = makeTitle(request.prompt);
    }
    conv.messages.push({ role: "user", text: request.prompt, ts: Date.now() });
    conv.messages.push({ role: "assistant", text: result.text, ts: Date.now() });
    await this.deps.history.upsert(conv);
  }

  private async handleResume(stream: vscode.ChatResponseStream): Promise<void> {
    const items = await this.deps.history.list();
    if (items.length === 0) {
      stream.markdown("保存された過去のチャットはありません。");
      return;
    }
    const picked = await vscode.window.showQuickPick(
      items.map((c) => ({
        label: c.title,
        description: new Date(c.updatedAt).toLocaleString(),
        detail: `${c.messages.length} メッセージ`,
        conv: c,
      })),
      { placeHolder: "再開する過去のチャットを選択してください" },
    );
    if (!picked) {
      stream.markdown("再開をキャンセルしました。");
      return;
    }
    this.current = picked.conv;
    stream.markdown(`### 「${picked.conv.title}」を再開しました\n\n`);
    for (const m of picked.conv.messages) {
      if (m.role === "user") {
        stream.markdown(`\n**🧑 あなた:** ${m.text}\n`);
      } else {
        stream.markdown(`\n**🤖 Copilot:**\n\n${m.text}\n`);
      }
    }
    stream.markdown("\n\n---\n続けてメッセージを送ると、この会話の続きから応答します。");
  }

  private renderError(stream: vscode.ChatResponseStream, e: unknown): void {
    log.error("chat handler error", e as Error);
    if (e instanceof BrowserSessionError) {
      stream.markdown(`⚠️ ${e.message}`);
      if (e.guidance) {
        stream.markdown(`\n\n${e.guidance}`);
      }
      return;
    }
    stream.markdown(
      `⚠️ エラーが発生しました: \`${describe(e)}\`\n\n「MS Copilot: ログを表示」で詳細を確認できます。`,
    );
  }
}

function newConversation(): Conversation {
  const now = Date.now();
  return {
    id: randomUUID(),
    title: "(新規チャット)",
    createdAt: now,
    updatedAt: now,
    messages: [],
  };
}

function describe(e: unknown): string {
  return e instanceof Error ? e.message : String(e);
}
