import * as vscode from "vscode";

export type Transport = "cdp" | "playwright";
export type EndpointVariant = "office" | "cloud";

export interface MsCopilotConfig {
  transport: Transport;
  endpointVariant: EndpointVariant;
  browserPath: string;
  debugPort: number;
  userDataDir: string;
  startUrl: string;
  maxMentionBytes: number;
}

export function getConfig(): MsCopilotConfig {
  const c = vscode.workspace.getConfiguration("mscopilot");
  return {
    transport: c.get<Transport>("transport", "cdp"),
    endpointVariant: c.get<EndpointVariant>("endpointVariant", "office"),
    browserPath: c.get<string>("browserPath", ""),
    debugPort: c.get<number>("debugPort", 9222),
    userDataDir: c.get<string>("userDataDir", ""),
    startUrl: c.get<string>("startUrl", "https://m365.cloud.microsoft/chat"),
    maxMentionBytes: c.get<number>("maxMentionBytes", 131072),
  };
}
