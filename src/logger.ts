import * as vscode from "vscode";

let channel: vscode.OutputChannel | undefined;

export function initLogger(context: vscode.ExtensionContext): vscode.OutputChannel {
  if (!channel) {
    channel = vscode.window.createOutputChannel("MS Copilot Chat");
    context.subscriptions.push(channel);
  }
  return channel;
}

function ts(): string {
  return new Date().toISOString();
}

export const log = {
  info(msg: string, ...args: unknown[]): void {
    channel?.appendLine(`[${ts()}] [info] ${format(msg, args)}`);
  },
  warn(msg: string, ...args: unknown[]): void {
    channel?.appendLine(`[${ts()}] [warn] ${format(msg, args)}`);
  },
  error(msg: string, ...args: unknown[]): void {
    channel?.appendLine(`[${ts()}] [error] ${format(msg, args)}`);
  },
  show(): void {
    channel?.show(true);
  },
};

function format(msg: string, args: unknown[]): string {
  if (args.length === 0) {
    return msg;
  }
  const rendered = args
    .map((a) => {
      if (a instanceof Error) {
        return `${a.message}\n${a.stack ?? ""}`;
      }
      if (typeof a === "string") {
        return a;
      }
      try {
        return JSON.stringify(a);
      } catch {
        return String(a);
      }
    })
    .join(" ");
  return `${msg} ${rendered}`;
}
