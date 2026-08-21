import { formulaService } from "./runtime";
import { captureError } from "./telemetry";

type CancelableInvocation = CustomFunctions.CancelableInvocation;

export async function query(
  sql: string,
  connection?: string,
  includeHeaders?: boolean,
  _refreshKey?: unknown,
  invocation?: CancelableInvocation,
): Promise<(string | number | boolean)[][]> {
  const controller = new AbortController();
  if (invocation) invocation.onCanceled = () => controller.abort();
  try {
    const result = await formulaService.query(sql, blankToUndefined(connection), includeHeaders !== false, controller.signal);
    return excelResult(result);
  } catch (error) {
    captureError(error, "custom-function.query");
    throw excelError(error);
  }
}

export async function value(
  sql: string,
  connection?: string,
  _refreshKey?: unknown,
  invocation?: CancelableInvocation,
): Promise<string | number | boolean> {
  const controller = new AbortController();
  if (invocation) invocation.onCanceled = () => controller.abort();
  try {
    return (await formulaService.value(sql, blankToUndefined(connection), controller.signal)) ?? "";
  } catch (error) {
    captureError(error, "custom-function.value");
    throw excelError(error);
  }
}

export async function call(functionName: string, ...args: unknown[]): Promise<(string | number | boolean)[][]> {
  try {
    return excelResult(await formulaService.call(functionName, args));
  } catch (error) {
    captureError(error, "custom-function.call");
    throw excelError(error);
  }
}

function blankToUndefined(value?: string): string | undefined {
  return value?.trim() || undefined;
}

function excelResult(matrix: Array<Array<string | number | boolean | null>>): Array<Array<string | number | boolean>> {
  return matrix.map((row) => row.map((cell) => cell ?? ""));
}

function excelError(error: unknown): CustomFunctions.Error {
  const message = error instanceof Error ? error.message : String(error);
  const unavailable = /auth|connect|network|fetch/i.test(message);
  return new CustomFunctions.Error(
    unavailable ? CustomFunctions.ErrorCode.notAvailable : CustomFunctions.ErrorCode.invalidValue,
    message.slice(0, 255),
  );
}

if (typeof CustomFunctions !== "undefined") {
  CustomFunctions.associate("QUERY", query);
  CustomFunctions.associate("VALUE", value);
  CustomFunctions.associate("CALL", call);
}
