import { test, expect } from '@playwright/test';

test('desktop updates require download and a separate install action, with persistent status at narrow width', async ({ page }) => {
  await page.setViewportSize({ width: 360, height: 480 });
  await page.addInitScript(() => {
    const state = { supported: true, daily: true, phase: 'available', message: 'A new Cupola release is available.', version: '0.5.2', build: '20261001.1', notes: 'https://github.com/Query-farm/vgi-excel/releases/tag/v0.5.2-20261001.1', lastChecked: '2026-09-30T12:00:00Z' };
    (window as any).updateCalls = [];
    Object.defineProperty(window, 'chrome', { value: { webview: { addEventListener() {}, postMessage(request: any) {
      let result: any = true;
      if (request.method.startsWith('updates.')) {
        (window as any).updateCalls.push(request.method);
        if (request.method === 'updates.download') { state.phase = 'ready'; state.message = 'The installer is downloaded and verified.'; }
        if (request.method === 'updates.install') { state.phase = 'handoff'; state.message = 'The updater is open. Save your work and close every Excel window to continue.'; }
        if (request.method === 'updates.preference') state.daily = request.params.daily;
        result = { ...state };
      } else if (request.method === 'connections.list') result = [];
      else if (request.method === 'agent.key.load') result = null;
      else if (request.method === 'app.diagnostics') result = 'Cupola diagnostics';
      setTimeout(() => window.vgiReceiveHostResponse?.({ id: request.id, result }), 0);
    } } }, configurable: true });
  });
  await page.goto('http://127.0.0.1:4173');
  await page.getByRole('button', { name: 'Settings', exact: true }).click();
  await page.getByRole('tab', { name: 'About', exact: true }).click();
  const updates = page.getByRole('region', { name: 'Updates' });
  await expect(updates).toContainText('20261001.1');
  await expect(updates.getByRole('button', { name: 'Install update' })).toHaveCount(0);
  await updates.getByRole('button', { name: 'Download update' }).click();
  await expect(updates).toContainText('close all Excel windows');
  expect(await page.evaluate(() => (window as any).updateCalls.includes('updates.install'))).toBe(false);
  await updates.getByRole('checkbox', { name: 'Check for updates daily' }).click();
  await expect(updates.getByRole('checkbox', { name: 'Check for updates daily' })).not.toBeChecked();
  await updates.getByRole('button', { name: 'Install update' }).click();
  await expect(updates.getByRole('status')).toContainText('The updater is open');
  await expect(updates.getByRole('status')).not.toContainText('was updated');
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
});
