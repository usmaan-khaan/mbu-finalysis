const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { buildSync } = require('esbuild');

const output = path.resolve('node_modules/.cache/mbu-analysis-test.cjs');
fs.mkdirSync(path.dirname(output), { recursive: true });
buildSync({ entryPoints: ['src/App.jsx'], bundle: true, platform: 'node', format: 'cjs', outfile: output, logLevel: 'silent' });
const app = require(output);

test('formatted numbers and missing values are parsed without invented zeros', () => {
  assert.equal(app.parseNumericValue('$1,200.50'), 1200.5);
  assert.equal(app.parseNumericValue('(€2,000)'), -2000);
  assert.equal(app.parseNumericValue('12%'), 0.12);
  assert.equal(app.parseNumericValue(''), null);
  assert.equal(app.parseNumericValue('N/A'), null);
  assert.equal(app.detectSchema([{ amount: '$1,000' }, { amount: '$2,000' }], ['amount'])[0].type, 'number');
  assert.deepEqual(app.aggregateMetrics([{ value: 10 }, { value: null }], [{ key: 'value', aggregation: 'average' }]), { count: 2, value: 10 });
});

test('metric aggregations and special category names work', () => {
  const rows = [{ segment: '__proto__', value: 10 }, { segment: '__proto__', value: 20 }];
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'sum' }]).value, 30);
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'average' }]).value, 15);
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'count' }]).value, 2);
  assert.equal(app.groupByGeneric(rows, 'segment', ['value'], [])[0].value, 30);
});

test('correlation ignores blank pairs', () => {
  const rows = [{ a: null, b: null }, { a: 1, b: 1 }, { a: 2, b: 2 }, { a: 3, b: 3 }, { a: 4, b: 4 }];
  assert.deepEqual(app.computeCorrelations(rows, ['a', 'b']), []);
});

test('calendar gaps do not masquerade as adjacent periods', () => {
  const trend = app.monthlyGroupGeneric([
    { date: '2026-01-10', sales: 100 }, { date: '2026-03-10', sales: 121 }
  ], 'date', ['sales'], []);
  assert.equal(app.periodDelta(trend, 'sales'), null);
  assert.ok(Math.abs(app.computeAvgGrowth(trend, 'sales') - 0.1) < 0.000001);
  assert.equal(app.parseDateValue('2026-02-30'), null);
  assert.equal(app.parseDateValue('08/10/2026', 'dmy'), '2026-10-08');
  assert.equal(app.parseDateValue('08/10/2026', 'mdy'), '2026-08-10');
});

test('backtesting reports actual historical error only when observations exist', () => {
  const trend = [100, 110, 121, 133.1, 146.41, 161.051].map((value, i) => ({
    month: `2026-${String(i + 1).padStart(2, '0')}`, sales: value
  }));
  const result = app.backtestForecast(trend, 'sales');
  assert.equal(result.samples, 3);
  assert.ok(result.meanAbsolutePercentError < 0.000001);
});

test('CSV exports neutralize spreadsheet formulas and saved setup keys do not truncate', () => {
  assert.match(app.rowsToCSV(['name'], [{ name: '=HYPERLINK("https://example.com","click")' }]), /'=/);
  assert.notEqual(app.schemaSignature(['a'.repeat(151) + 'x']), app.schemaSignature(['a'.repeat(151) + 'y']));
});
