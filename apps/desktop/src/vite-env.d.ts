/// <reference types="vite/client" />

interface Window {
  chrome?: { webview?: { postMessage(value: unknown): void; addEventListener(type: "message", handler: (event: MessageEvent) => void): void } };
  vgiReceiveHostResponse?: (value: unknown) => void;
  vgiSelectTab?: (tab: string) => void;
}
