import { test, expect } from '@playwright/test';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';

async function upload(page, contents, name = 'primary.csv', index = 0) {
  await page.locator('input[type="file"]').nth(index).setInputFiles({
    name, mimeType: name.endsWith('.json') ? 'application/json' : 'text/csv', buffer: Buffer.from(contents),
  });
}
async function nav(page, name) {
  await page.getByRole('button', { name, exact: true }).first().click();
}

for (const [mode, url] of [
  ['hosted', 'http://127.0.0.1:4173'],
  ['standalone', pathToFileURL(resolve('dist/index.html')).href],
]) {
  test.describe(mode, () => {
    test.beforeEach(async ({ page }) => {
      const failures = [];
      page.on('pageerror', error => failures.push(error.message));
      page.on('console', message => { if (message.type() === 'error') failures.push(message.text()); });
      page.on('request', request => {
        if (!/^(file:|data:|blob:|http:\/\/127\.0\.0\.1:4173(?:\/|$))/.test(request.url())) failures.push(`External request: ${request.url()}`);
      });
      await page.goto(url);
      await expect(page.getByText('Upload your data to get started', { exact: true })).toBeVisible();
      page.__reviewFailures = failures;
    });
    test.afterEach(async ({ page }) => {
      expect(page.__reviewFailures).toEqual([]);
    });

    test('single-column CSV loads into the dashboard', async ({ page }) => {
      await upload(page, 'sales\n10\n20');
      await expect(page.getByText('Review your import', { exact: true })).toBeVisible();
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await expect(page.getByText('Portfolio Dashboard', { exact: true })).toBeVisible();
      await expect(page.getByLabel('sales aggregation', { exact: true })).toHaveValue('sum');
    });

    test('reviewed categories and invalid date counts survive confirmation', async ({ page }) => {
      await upload(page, JSON.stringify([
        { id: 1001, label: '2026-01-01', date: '13/10/2026', sales: 10 },
        { id: 1002, label: '2026-02-01', date: '31/02/2026', sales: 20 },
      ]), 'primary.json');
      await page.getByLabel('id type', { exact: true }).selectOption('category');
      await page.getByLabel('label type', { exact: true }).selectOption('category');
      await page.getByLabel('date type', { exact: true }).selectOption('date-dmy');
      await page.getByLabel('date column name', { exact: true }).fill('Transaction date');
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await expect(page.getByLabel('id aggregation', { exact: true })).toHaveCount(0);
      await expect(page.getByLabel('label aggregation', { exact: true })).toHaveCount(0);
      await nav(page, 'Insights');
      await expect(page.getByTestId('unparseable-date-count')).toHaveText('1');
    });

    test('a saved setup cannot override a new category choice', async ({ page }) => {
      const csv = 'id,sales\n1001,10\n1002,20';
      await upload(page, csv);
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await page.getByRole('button', { name: 'Remember this setup', exact: true }).click();
      await expect(page.getByRole('button', { name: 'Saved', exact: true })).toBeVisible();
      await upload(page, csv);
      await page.getByLabel('id type', { exact: true }).selectOption('category');
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await expect(page.getByText('Applied your saved setup for files shaped like this.', { exact: false })).toBeVisible();
      await expect(page.getByLabel('id aggregation', { exact: true })).toHaveCount(0);
    });

    test('both comparison totals refresh after aggregation changes', async ({ page }) => {
      const csv = 'segment,sales\nx,10\nx,20';
      await upload(page, csv);
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await nav(page, 'Compare');
      await upload(page, csv, 'secondary.csv', 1);
      await expect(page.getByTestId('compare-a-sales')).toHaveText('$30');
      await expect(page.getByTestId('compare-b-sales')).toHaveText('$30');
      for (const [aggregation, value] of [['average', '$15'], ['count', '$2.00'], ['min', '$10'], ['max', '$20']]) {
        await nav(page, 'Dashboard');
        await page.getByLabel('sales aggregation', { exact: true }).selectOption(aggregation);
        await nav(page, 'Compare');
        await expect(page.getByTestId('compare-a-sales')).toHaveText(value);
        await expect(page.getByTestId('compare-b-sales')).toHaveText(value);
      }
    });

    test('switching to an unreviewed workbook sheet requires review', async ({ page }) => {
      const workbook = XLSX.utils.book_new();
      for (const name of ['First', 'Second']) XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([{ id: 1001, sales: 10 }, { id: 1002, sales: 20 }]), name);
      await page.locator('input[type="file"]').first().setInputFiles({ name: 'sheets.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }) });
      await page.getByLabel('id type', { exact: true }).selectOption('category');
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await page.getByRole('button', { name: 'Second', exact: true }).click();
      await expect(page.getByText('Review your import', { exact: true })).toBeVisible();
      await page.getByLabel('id type', { exact: true }).selectOption('category');
      await page.getByRole('button', { name: 'Analyze data', exact: true }).click();
      await page.getByRole('button', { name: 'First', exact: true }).click();
      await expect(page.getByText('Portfolio Dashboard', { exact: true })).toBeVisible();
      await expect(page.getByLabel('id aggregation', { exact: true })).toHaveCount(0);
    });
  });
}
