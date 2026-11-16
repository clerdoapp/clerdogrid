declare const __APP_VERSION__: string;
/// <reference types="vite/client" />

interface ImportMetaEnv {
  /** '1' in the build served at grid.clerdo.app, which signs in with Clerdo Account and saves to the cloud. */
  readonly VITE_CLERDO_CLOUD?: string;
}
