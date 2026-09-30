import { test as base, expect } from '@playwright/test';

export const officeOrigin = process.env.CUPOLA_OFFICE_BASE_URL ?? 'https://127.0.0.1:4184';
export const test = base.extend({
  context: async ({ context }, use) => {
    // Hosted smoke tests exercise production bundles without sending test errors.
    await context.route(/^https:\/\/[^/]*sentry\.io\//, route => route.abort());
    await use(context);
  },
});
export { expect };
