import * as Sentry from "@sentry/react";
import { classifyTelemetryError, sanitizeTelemetryText } from "@query-farm/vgi-excel-core";

const DEFAULT_DSN = "https://e124c1f57fececb77135645c9afa6349@o4511299556081664.ingest.us.sentry.io/4511948908134400";

export function initializeTelemetry(): void {
  const enabled = import.meta.env.VITE_SENTRY_ENABLED !== "0"
    && (!import.meta.env.DEV || import.meta.env.VITE_SENTRY_ENABLE_LOCAL === "1");
  Sentry.init({
    dsn: import.meta.env.VITE_SENTRY_OFFICE_DSN || DEFAULT_DSN,
    enabled,
    environment: import.meta.env.VITE_SENTRY_ENVIRONMENT || "production",
    release: `cupola-excel@${__APP_VERSION__}+${__BUILD_ID__}`,
    dist: "office",
    sendDefaultPii: false,
    tracesSampleRate: 0,
    maxBreadcrumbs: 0,
    sendClientReports: false,
    integrations: (defaults) => defaults.filter((integration) => integration.name !== "BrowserSession"),
    beforeBreadcrumb: () => null,
    beforeSend: scrubEvent,
    initialScope: {
      tags: { product: "cupola-excel", host: "microsoft-365", version: __APP_VERSION__, build: __BUILD_ID__, transport: "https" },
    },
  });
}

export function captureError(error: unknown, operation: string): void {
  Sentry.captureException(error instanceof Error ? error : new Error(sanitizeTelemetryText(error)), {
    tags: { operation: sanitizeTelemetryText(operation) },
  });
}

export const TelemetryErrorBoundary = Sentry.ErrorBoundary;

function scrubEvent(event: Sentry.ErrorEvent): Sentry.ErrorEvent {
  event.user = undefined;
  event.request = undefined;
  event.server_name = undefined;
  event.transaction = undefined;
  event.logger = undefined;
  event.fingerprint = undefined;
  event.contexts = undefined;
  event.extra = undefined;
  event.breadcrumbs = undefined;
  if (event.message) event.message = classifyTelemetryError(event.message);
  if (event.logentry?.message) event.logentry.message = classifyTelemetryError(event.logentry.message);
  if (event.logentry) event.logentry.params = undefined;
  if (event.tags) event.tags = Object.fromEntries(Object.entries(event.tags).filter(([key]) => ["product", "host", "version", "build", "transport", "operation"].includes(key)).map(([key, value]) => [key, sanitizeTelemetryText(value)]));
  for (const exception of event.exception?.values ?? []) {
    if (exception.value) exception.value = classifyTelemetryError(exception.value);
    exception.type = /^(?:Error|TypeError|RangeError|ReferenceError|SyntaxError|EvalError|URIError|AggregateError|DOMException)$/i.test(exception.type ?? "") ? exception.type : "Error";
    exception.module = undefined;
    exception.mechanism = undefined;
    for (const frame of exception.stacktrace?.frames ?? []) {
      frame.filename = frame.filename ? sanitizeTelemetryText(frame.filename) : frame.filename;
      frame.abs_path = frame.abs_path ? sanitizeTelemetryText(frame.abs_path) : frame.abs_path;
      frame.context_line = undefined;
      frame.pre_context = undefined;
      frame.post_context = undefined;
      frame.vars = undefined;
      frame.module = undefined;
      frame.module_metadata = undefined;
    }
  }
  return event;
}
