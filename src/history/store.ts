import * as vscode from "vscode";
import { log } from "../logger";

export interface StoredMessage {
  role: "user" | "assistant";
  text: string;
  ts: number;
}

export interface Conversation {
  /** ローカル会話 ID(拡張が採番)。 */
  id: string;
  /** 一覧表示用タイトル(最初のユーザー発話から生成)。 */
  title: string;
  /** substrate 側の会話 ID(継続に使用、無い場合あり)。 */
  substrateConversationId?: string;
  createdAt: number;
  updatedAt: number;
  messages: StoredMessage[];
}

/**
 * 会話履歴をローカル(globalStorage)に JSON 保存する。
 * /resume で一覧・復元する。サーバー履歴の取得は対象外(計画の合意による)。
 */
export class HistoryStore {
  private readonly fileUri: vscode.Uri;
  private cache: Conversation[] | undefined;

  constructor(private readonly context: vscode.ExtensionContext) {
    this.fileUri = vscode.Uri.joinPath(context.globalStorageUri, "history.json");
  }

  async list(): Promise<Conversation[]> {
    const all = await this.load();
    return [...all].sort((a, b) => b.updatedAt - a.updatedAt);
  }

  async get(id: string): Promise<Conversation | undefined> {
    const all = await this.load();
    return all.find((c) => c.id === id);
  }

  async upsert(conv: Conversation): Promise<void> {
    const all = await this.load();
    const idx = all.findIndex((c) => c.id === conv.id);
    if (idx >= 0) {
      all[idx] = conv;
    } else {
      all.push(conv);
    }
    this.cache = all;
    await this.persist(all);
  }

  async delete(id: string): Promise<void> {
    const all = (await this.load()).filter((c) => c.id !== id);
    this.cache = all;
    await this.persist(all);
  }

  private async load(): Promise<Conversation[]> {
    if (this.cache) {
      return this.cache;
    }
    try {
      const bytes = await vscode.workspace.fs.readFile(this.fileUri);
      const parsed = JSON.parse(Buffer.from(bytes).toString("utf8"));
      this.cache = Array.isArray(parsed) ? (parsed as Conversation[]) : [];
    } catch {
      this.cache = [];
    }
    return this.cache;
  }

  private async persist(all: Conversation[]): Promise<void> {
    try {
      await vscode.workspace.fs.createDirectory(this.context.globalStorageUri);
      const data = Buffer.from(JSON.stringify(all, null, 2), "utf8");
      await vscode.workspace.fs.writeFile(this.fileUri, data);
    } catch (e) {
      log.error("履歴の保存に失敗", e as Error);
    }
  }
}

export function makeTitle(firstUserText: string): string {
  const oneLine = firstUserText.replace(/\s+/g, " ").trim();
  return oneLine.length > 60 ? oneLine.slice(0, 60) + "…" : oneLine || "(無題)";
}
