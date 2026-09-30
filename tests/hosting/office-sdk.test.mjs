import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

test('Office SDK uses CORS so require-corp does not block its cross-origin script', async () => {
 for (const file of ['apps/office/taskpane.html','apps/office/oauth-dialog.html']) {
  const html=await readFile(file,'utf8');
  assert.match(html,/<script\s+crossorigin="anonymous"\s+src="https:\/\/appsforoffice\.microsoft\.com\/lib\/1\/hosted\/office\.js"><\/script>/);
 }
});
