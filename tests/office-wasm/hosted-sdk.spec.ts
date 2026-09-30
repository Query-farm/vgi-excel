import { expect, test, officeOrigin } from './test';

test('hosted bootstrap loads the real Microsoft Office SDK under isolation headers', async ({ page }) => {
  test.skip(!process.env.CUPOLA_OFFICE_BASE_URL, 'Requires the production host.');
  const sdk = page.waitForResponse('https://appsforoffice.microsoft.com/lib/1/hosted/office.js');
  await page.goto(`${officeOrigin}/taskpane.html`);
  expect((await sdk).ok()).toBe(true);
  await expect.poll(() => page.evaluate(() => typeof Office !== 'undefined' && typeof Office.onReady === 'function')).toBe(true);
  expect(await page.evaluate(() => globalThis.crossOriginIsolated)).toBe(true);
  await expect(page.getByRole('tab', { name: 'Query Editor', exact: true })).toBeVisible();
});
