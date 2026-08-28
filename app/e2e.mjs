import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const require = createRequire(import.meta.url);
const appDir = path.dirname(fileURLToPath(import.meta.url));
const repoDir = path.dirname(appDir);
const playwrightPath = process.env.GP_CATS_PLAYWRIGHT_PATH || 'playwright';
const browserPath = process.env.GP_CATS_BROWSER_PATH;
const { chromium } = require(playwrightPath);
const XLSX = require(path.join(appDir, 'vendor', 'xlsx.full.min.js'));

const browser = await chromium.launch({ headless: true, ...(browserPath ? { executablePath: browserPath } : {}) });
const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
const pageErrors = [];
page.on('pageerror', error => pageErrors.push(error.message));
const screenshotsDir = path.join(appDir, 'screenshots');
fs.mkdirSync(screenshotsDir, { recursive: true });

try {
  const appUrl = pathToFileURL(path.join(repoDir, 'dist', 'category-permissions.html')).href;
  await page.goto(appUrl);
  await page.locator('#permission-file-empty').setInputFiles(path.join(appDir, 'samples', 'permissions.xlsx'));
  await page.locator('.file-button').filter({ hasText: 'permissions.xlsx' }).waitFor();
  await page.locator('#category-file-empty').setInputFiles(path.join(appDir, 'samples', 'categories.json'));
  await page.locator('.tree-row').first().waitFor();

  assert.match(await page.locator('.summary-strip').innerText(), /17\s+категорий/);
  assert.match(await page.locator('.summary-strip').innerText(), /4\s+партнеров/);
  assert.match(await page.locator('.alert-banner').innerText(), /99999/);

  await page.getByRole('button', { name: 'С ограничениями', exact: true }).click();
  const restrictedRows = await page.locator('.tree-row').allTextContents();
  assert(restrictedRows.some(row => row.includes('Электроника')));
  assert(!restrictedRows.some(row => row.includes('Красота')));

  await page.getByRole('button', { name: 'Предупреждения', exact: true }).click();
  const warningRows = await page.locator('.tree-row').allTextContents();
  assert(warningRows.some(row => row.includes('Телефоны')));
  assert(!warningRows.some(row => row.includes('Бытовая техника')));

  await page.getByRole('button', { name: 'Все', exact: true }).click();
  await page.getByRole('button', { name: 'Партнеры', exact: true }).click();
  await page.locator('#partner-select').selectOption('partner-alfa');
  assert.match(await page.locator('.section-title').innerText(), /partner-alfa/);
  await page.getByRole('button', { name: 'Категории', exact: true }).click();
  await page.screenshot({ path: path.join(screenshotsDir, 'real-app-desktop.png'), fullPage: true });

  const downloadPromise = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Экспорт XLSX', exact: true }).click();
  const download = await downloadPromise;
  const reportPath = path.join(os.tmpdir(), `category-permissions-report-${process.pid}.xlsx`);
  await download.saveAs(reportPath);

  const report = XLSX.read(fs.readFileSync(reportPath), { type: 'buffer' });
  assert.deepEqual(report.SheetNames, ['Категории', 'Партнеры', 'Ошибки', 'Предупреждения']);
  fs.unlinkSync(reportPath);

  const mobile = await browser.newPage({ viewport: { width: 390, height: 844 } });
  await mobile.goto(appUrl);
  await mobile.locator('#permission-file-empty').setInputFiles(path.join(appDir, 'samples', 'permissions.xlsx'));
  await mobile.locator('.file-button').filter({ hasText: 'permissions.xlsx' }).waitFor();
  await mobile.locator('#category-file-empty').setInputFiles(path.join(appDir, 'samples', 'categories.json'));
  await mobile.locator('.tree-row').first().waitFor();
  const dimensions = await mobile.evaluate(() => ({ width: innerWidth, scrollWidth: document.documentElement.scrollWidth }));
  assert.equal(dimensions.scrollWidth, dimensions.width);
  await mobile.screenshot({ path: path.join(screenshotsDir, 'real-app-mobile.png'), fullPage: true });
  await mobile.close();

  assert.deepEqual(pageErrors, []);
  console.log('E2E PASS: import, validation, filters, partner mode, XLSX export');
} finally {
  await browser.close();
}
