import { test, expect } from '@playwright/test';

test('ページが正常に表示される', async ({ page }) => {
  await page.goto('/');

  await expect(page).toHaveTitle(/Excel/);
  await expect(page.getByRole('heading', { name: 'Excel ダウンロード' })).toBeVisible();
});

test('ダウンロードボタンが4つ表示される', async ({ page }) => {
  await page.goto('/');

  await expect(page.getByRole('button', { name: 'Excel をダウンロード', exact: true })).toBeVisible();
  await expect(page.getByRole('button', { name: '5シート Excel をダウンロード' })).toBeVisible();
  await expect(page.getByRole('button', { name: '入力シート Excel をダウンロード' })).toBeVisible();
  await expect(page.getByRole('button', { name: '保護シート Excel をダウンロード' })).toBeVisible();
});

test('売上レポート（1シート）をダウンロードできる', async ({ page }) => {
  await page.goto('/');

  // download イベントを待ち受けてからボタンをクリック
  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Excel をダウンロード', exact: true }).click();
  const download = await downloadPromise;

  // ファイル名が .xlsx で終わる
  expect(download.suggestedFilename()).toMatch(/\.xlsx$/);

  // ファイルサイズが 0 より大きい
  const path = await download.path();
  const { stat } = await import('fs/promises');
  const { size } = await stat(path);
  expect(size).toBeGreaterThan(0);
});
