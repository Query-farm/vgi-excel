/** Connection probes must not leave a blocked WASM worker behind. */
export const CONNECTION_TEST_TIMEOUT_MS = 20_000;

export async function runConnectionTest<T>(operation: (signal: AbortSignal) => Promise<T>): Promise<T> {
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    const deadline = new Promise<never>((_resolve, reject) => {
      timer = setTimeout(() => {
        reject(new Error("Connection test timed out after 20 seconds. Check the HTTPS address and your network, then try again."));
        controller.abort();
      }, CONNECTION_TEST_TIMEOUT_MS);
    });
    return await Promise.race([operation(controller.signal), deadline]);
  } finally {
    clearTimeout(timer);
    controller.abort();
  }
}
