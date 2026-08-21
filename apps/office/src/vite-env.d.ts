/// <reference types="vite/client" />

declare module "*.wasm?url" { const value: string; export default value; }
declare module "*.js?url" { const value: string; export default value; }

declare namespace CustomFunctions {
  enum ErrorCode {
    invalidValue = "invalidValue",
    notAvailable = "notAvailable"
  }
  class Error {
    constructor(code: ErrorCode, message?: string);
  }
  interface CancelableInvocation {
    onCanceled: () => void;
  }
  function associate(id: string, handler: (...args: any[]) => any): void;
}
