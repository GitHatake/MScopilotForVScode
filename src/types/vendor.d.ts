// chrome-remote-interface / playwright は最小限の型で扱う(playwright はオプションの
// 動的依存で未インストールのことがある)。詳細な型は実行時挙動に依存するため any とする。
declare module "chrome-remote-interface" {
  const CDP: any;
  export default CDP;
}

declare module "playwright" {
  export const chromium: any;
  const _default: any;
  export default _default;
}
