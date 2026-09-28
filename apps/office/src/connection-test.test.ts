import { afterEach, expect, it, vi } from "vitest";
import { CONNECTION_TEST_TIMEOUT_MS, runConnectionTest } from "./connection-test";

afterEach(() => vi.useRealTimers());

it("ends a stalled attachment at the deadline and releases its worker", async () => {
  vi.useFakeTimers();
  const release = vi.fn();
  const result = runConnectionTest(signal => {
    signal.addEventListener("abort", release);
    return new Promise<never>(() => {});
  });
  const assertion = expect(result).rejects.toThrow("timed out after 20 seconds");
  await vi.advanceTimersByTimeAsync(CONNECTION_TEST_TIMEOUT_MS);
  await assertion;
  expect(release).toHaveBeenCalledOnce();
  expect(vi.getTimerCount()).toBe(0);
  await expect(runConnectionTest(async () => "retry succeeded")).resolves.toBe("retry succeeded");
});

it("releases successful probes and reports early failures without waiting", async () => {
  vi.useFakeTimers();
  let successfulSignal: AbortSignal | undefined;
  await expect(runConnectionTest(async signal => { successfulSignal = signal; return 42; })).resolves.toBe(42);
  expect(successfulSignal?.aborted).toBe(true);
  const release = vi.fn();
  await expect(runConnectionTest(async signal => {
    signal.addEventListener("abort", release);
    throw new Error("network unavailable");
  })).rejects.toThrow("network unavailable");
  expect(release).toHaveBeenCalledOnce();
  expect(vi.getTimerCount()).toBe(0);
});

it("ignores a late completion after timeout", async () => {
  vi.useFakeTimers();
  let complete!: (value: number) => void;
  const result = runConnectionTest(() => new Promise<number>(resolve => { complete = resolve; }));
  const assertion = expect(result).rejects.toThrow("timed out");
  await vi.advanceTimersByTimeAsync(CONNECTION_TEST_TIMEOUT_MS);
  await assertion;
  complete(42);
  await expect(result).rejects.toThrow("timed out");
});
