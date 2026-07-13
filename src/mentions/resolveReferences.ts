import * as vscode from "vscode";
import { log } from "../logger";

export interface ResolvedContext {
  /** プロンプト先頭に注入するコンテキスト本文(空なら注入なし)。 */
  text: string;
  /** 実際に取り込んだ参照の表示名。UI の参照表示に使う。 */
  used: { name: string; uri?: vscode.Uri }[];
}

/**
 * Chat の #file / #folder 参照(request.references)を解決してテキスト化する。
 * フォルダは配下ファイルを列挙し、全体のバイト上限で切り詰める。
 */
export async function resolveReferences(
  references: readonly vscode.ChatPromptReference[],
  maxBytes: number,
): Promise<ResolvedContext> {
  const used: { name: string; uri?: vscode.Uri }[] = [];
  const blocks: string[] = [];
  let budget = maxBytes;

  for (const ref of references) {
    if (budget <= 0) {
      break;
    }
    try {
      const value = ref.value;
      if (value instanceof vscode.Uri) {
        budget = await appendUri(value, blocks, used, budget);
      } else if (isLocation(value)) {
        budget = await appendLocation(value, blocks, used, budget);
      } else if (typeof value === "string") {
        const block = section(ref.id, value);
        blocks.push(clip(block, budget));
        used.push({ name: ref.id });
        budget -= Buffer.byteLength(block);
      }
    } catch (e) {
      log.warn(`reference 解決に失敗: ${ref.id}`, e as Error);
    }
  }

  if (blocks.length === 0) {
    return { text: "", used };
  }
  const header =
    "以下は参照として添付されたファイル/フォルダの内容です。回答時の参考にしてください。\n";
  return { text: header + blocks.join("\n"), used };
}

async function appendUri(
  uri: vscode.Uri,
  blocks: string[],
  used: { name: string; uri?: vscode.Uri }[],
  budget: number,
): Promise<number> {
  const stat = await vscode.workspace.fs.stat(uri);
  if (stat.type & vscode.FileType.Directory) {
    const files = await listFiles(uri, 200);
    used.push({ name: `${basename(uri)}/`, uri });
    for (const file of files) {
      if (budget <= 0) {
        break;
      }
      budget = await appendFile(file, blocks, budget);
    }
    return budget;
  }
  used.push({ name: basename(uri), uri });
  return appendFile(uri, blocks, budget);
}

async function appendFile(uri: vscode.Uri, blocks: string[], budget: number): Promise<number> {
  const bytes = await vscode.workspace.fs.readFile(uri);
  if (isProbablyBinary(bytes)) {
    return budget;
  }
  const content = Buffer.from(bytes).toString("utf8");
  const block = section(vscode.workspace.asRelativePath(uri), content);
  const clipped = clip(block, budget);
  blocks.push(clipped);
  return budget - Buffer.byteLength(clipped);
}

async function appendLocation(
  loc: vscode.Location,
  blocks: string[],
  used: { name: string; uri?: vscode.Uri }[],
  budget: number,
): Promise<number> {
  const doc = await vscode.workspace.openTextDocument(loc.uri);
  const text = doc.getText(loc.range);
  const name = `${vscode.workspace.asRelativePath(loc.uri)}:${loc.range.start.line + 1}-${loc.range.end.line + 1}`;
  used.push({ name, uri: loc.uri });
  const block = section(name, text);
  const clipped = clip(block, budget);
  blocks.push(clipped);
  return budget - Buffer.byteLength(clipped);
}

async function listFiles(dir: vscode.Uri, limit: number): Promise<vscode.Uri[]> {
  const out: vscode.Uri[] = [];
  const stack: vscode.Uri[] = [dir];
  while (stack.length > 0 && out.length < limit) {
    const cur = stack.pop()!;
    let entries: [string, vscode.FileType][];
    try {
      entries = await vscode.workspace.fs.readDirectory(cur);
    } catch {
      continue;
    }
    for (const [name, type] of entries) {
      if (name === "node_modules" || name === ".git" || name.startsWith(".")) {
        continue;
      }
      const child = vscode.Uri.joinPath(cur, name);
      if (type & vscode.FileType.Directory) {
        stack.push(child);
      } else if (type & vscode.FileType.File) {
        out.push(child);
        if (out.length >= limit) {
          break;
        }
      }
    }
  }
  return out;
}

function section(name: string, content: string): string {
  return `\n===== ${name} =====\n${content}\n===== END =====\n`;
}

function clip(s: string, budget: number): string {
  if (Buffer.byteLength(s) <= budget) {
    return s;
  }
  const truncated = Buffer.from(s, "utf8").subarray(0, Math.max(0, budget)).toString("utf8");
  return truncated + "\n…(切り詰め)…\n";
}

function isLocation(v: unknown): v is vscode.Location {
  return !!v && typeof v === "object" && "uri" in v && "range" in v;
}

function basename(uri: vscode.Uri): string {
  const parts = uri.path.split("/").filter(Boolean);
  return parts[parts.length - 1] ?? uri.path;
}

function isProbablyBinary(bytes: Uint8Array): boolean {
  const n = Math.min(bytes.length, 4096);
  for (let i = 0; i < n; i++) {
    if (bytes[i] === 0) {
      return true;
    }
  }
  return false;
}
