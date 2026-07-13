import type { MsCopilotConfig } from "../config";
import { BrowserSession } from "./browserSession";
import { CdpBridge } from "./cdpBridge";
import { PlaywrightBridge } from "./playwrightBridge";

/**
 * 設定に応じた BrowserSession を生成する。
 * 呼び出し側で 1 インスタンスを使い回す想定(接続/ページを保持するため)。
 */
export function createBrowserSession(
  config: MsCopilotConfig,
  storageDir: string,
): BrowserSession {
  if (config.transport === "playwright") {
    return new PlaywrightBridge(config, storageDir);
  }
  return new CdpBridge(config, storageDir);
}

export { BrowserSession, BrowserSessionError } from "./browserSession";
