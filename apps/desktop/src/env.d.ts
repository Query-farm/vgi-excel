declare const __APP_VERSION__: string;
declare const __BUILD_ID__: string;

interface ImportMetaEnv {
  readonly VITE_SENTRY_DESKTOP_DSN?: string;
  readonly VITE_SENTRY_ENABLED?: string;
  readonly VITE_SENTRY_ENABLE_LOCAL?: string;
  readonly VITE_SENTRY_ENVIRONMENT?: string;
}

interface ImportMeta { readonly env: ImportMetaEnv; }
