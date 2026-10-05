import { expect, test, officeOrigin } from './test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

const lock = JSON.parse(readFileSync(resolve('vgi-extensions.lock.json'), 'utf8'));
const assetPath = (platform: string) => `vgi/${lock.artifacts[platform].sha256}/${platform}/vgi.duckdb_extension.wasm`;

const endpoint = process.env.CUPOLA_LIVE_VGI_ENDPOINT ?? 'https://vgi-open-meteo.rusty-bb6.workers.dev';

for (const isolated of [true, false]) {
  test(`packaged VGI loads without its upstream repository (isolated=${isolated})`, async ({ page, context }) => {
    const upstream: string[] = [], packaged: string[] = [];
    await context.route(/https:\/\/haybarn-extensions\.query\.farm\/community\/.*vgi/, route => { upstream.push(route.request().url()); return route.abort(); });
    await page.route('https://appsforoffice.microsoft.com/**', route => route.abort());
    if (!isolated) await page.route('**/taskpane.html', async route => {
      const response = await route.fetch(); const headers = response.headers();
      delete headers['cross-origin-opener-policy']; delete headers['cross-origin-embedder-policy'];
      await route.fulfill({ response, headers });
    });
    page.on('response', response => { if (response.url().includes('/vgi/')) packaged.push(response.url()); });
    await page.addInitScript(endpoint => localStorage.setItem('vgi.excel.connections.v1', JSON.stringify([{
      name: 'packaged', catalog: 'open_meteo', location: endpoint,
    }])), endpoint);
    await page.goto(`${officeOrigin}/taskpane.html`);
    expect(await page.evaluate(() => crossOriginIsolated)).toBe(isolated);
    await page.getByLabel('SQL query').fill("SELECT open_meteo.main.weather_code_text(0) AS weather");
    await page.getByRole('button', { name: 'Run', exact: true }).click();
    await expect(page.getByRole('table', { name: 'Query results' })).toContainText('Clear sky');
    expect(packaged).toContain(`${officeOrigin}/${assetPath(isolated ? 'wasm_threads' : 'wasm_eh')}`);
    expect(upstream).toEqual([]);
  });
}

test('missing packaged VGI fails without falling back to community', async ({ page, context }) => {
  const upstream: string[] = [];
  await context.route(/https:\/\/haybarn-extensions\.query\.farm\/community\/.*vgi/, route => { upstream.push(route.request().url()); return route.abort(); });
  await page.route('https://appsforoffice.microsoft.com/**', route => route.abort());
  await page.route('**/vgi/**', route => route.fulfill({ status: 404, body: 'Missing packaged artifact' }));
  await page.addInitScript(endpoint => localStorage.setItem('vgi.excel.connections.v1', JSON.stringify([{
    name: 'packaged', catalog: 'open_meteo', location: endpoint,
  }])), endpoint);
  await page.goto(`${officeOrigin}/taskpane.html`);
  await page.getByLabel('SQL query').fill('SELECT 42');
  await page.getByRole('button', { name: 'Run', exact: true }).click();
  await expect(page.locator('.notice.error')).toBeVisible();
  await expect(page.getByRole('table', { name: 'Query results' })).toHaveCount(0);
  expect(upstream).toEqual([]);
});
