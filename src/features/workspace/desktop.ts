// Window controls on the desktop app (electron/preload.ts); absent in the browser.
export interface DesktopWindow {
  setFullScreen?: (on: boolean) => Promise<boolean>;
  isFullScreen?: () => Promise<boolean>;
  arrangeWindows?: (how: 'tiled' | 'horizontal' | 'vertical' | 'cascade') => Promise<void>;
  newWindow?: () => void;
  windowCount?: () => Promise<number>;
}
export const desktopWindow = (globalThis as { clerdo?: DesktopWindow }).clerdo;
