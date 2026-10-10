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

test('reviewed categories override detection for numeric IDs and date-looking text', () => {
  const sheet = app.prepareImportSheet({
    fields: ['id', 'label', 'amount'],
    rows: [{ id: 1001, label: '2026-01-01', amount: '$10' }, { id: 1002, label: '2026-02-01', amount: '$20' }],
  }, { id: 'category', label: 'category', amount: 'number' }, { id: 'Customer ID', label: 'label', amount: 'amount' });
  assert.deepEqual(sheet.schema.map(col => [col.key, col.type]), [['Customer ID', 'category'], ['label', 'category'], ['amount', 'number']]);
  assert.equal(sheet.rows[0]['Customer ID'], 1001);
  assert.equal(sheet.rows[0].label, '2026-01-01');
  assert.equal(sheet.rows[0].amount, 10);
  assert.throws(() => app.prepareImportSheet({ fields: ['a', 'b'], rows: [{ a: 1, b: 2 }] }, { a: 'number', b: 'number' }, { a: 'Same', b: 'same' }), /unique/);
});

test('invalid dates remain diagnosable after normalization, renaming, and day-first review', () => {
  const sheet = app.prepareImportSheet({
    fields: ['date', 'amount'],
    rows: [{ date: '13/10/2026', amount: 10 }, { date: '31/02/2026', amount: 20 }, { date: '', amount: 30 }],
  }, { date: 'date-dmy', amount: 'number' }, { date: 'Transaction date', amount: 'amount' });
  assert.equal(sheet.rows[0]['Transaction date'], '2026-10-13');
  assert.equal(sheet.rows[1]['Transaction date'], null);
  const quality = app.computeDataQuality(sheet.rows, sheet.fields, 'Transaction date', sheet.sourceRows, sheet.columnTypes);
  assert.equal(quality.badDateCount, 1);
  assert.equal(quality.nullStats.find(col => col.field === 'Transaction date').nulls, 2);
});

test('reviewed numeric and date columns retain their types even with no valid values', () => {
  const sheet = app.prepareImportSheet({ fields: ['date', 'amount'], rows: [{ date: 'bad', amount: 'N/A' }] },
    { date: 'date', amount: 'number' }, { date: 'date', amount: 'amount' });
  assert.deepEqual(sheet.schema.map(col => col.type), ['date', 'number']);
  assert.deepEqual(sheet.rows, [{ date: null, amount: null }]);
});

test('single-column CSVs import but malformed records still fail', () => {
  assert.deepEqual(app.parseDelimitedText('sales\n10\n20'), { fields: ['sales'], rows: [{ sales: 10 }, { sales: 20 }] });
  assert.deepEqual(app.parseDelimitedText('sales\n10\n20', '\t').rows, [{ sales: 10 }, { sales: 20 }]);
  assert.deepEqual(app.parseDelimitedText('segment\tsales\nx\t10', '\t').rows, [{ segment: 'x', sales: 10 }]);
  assert.throws(() => app.parseDelimitedText('segment,sales\nx,10,extra'), /CSV parse error/);
  assert.throws(() => app.parseDelimitedText('segment,sales\nx'), /CSV parse error/);
  assert.throws(() => app.parseDelimitedText('segment,sales\n"unterminated,10'), /CSV parse error/);
  assert.throws(() => app.parseDelimitedText('sales'), /Could not find columns or rows/);
});

test('minimum, maximum, and count exclude blank and invalid values', () => {
  const rows = [{ value: 10 }, { value: -5 }, { value: null }, { value: 'N/A' }];
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'min' }]).value, -5);
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'max' }]).value, 10);
  assert.equal(app.aggregateMetrics(rows, [{ key: 'value', aggregation: 'count' }]).value, 2);
});

test('saved setups cannot undo reviewed types and keep compatible budget preferences', () => {
  const config = { metrics: [{ key: 'sales', aggregation: 'sum', on: true }, { key: 'budget', on: true }], dimensions: [{ key: 'id', on: true }], dateField: 'date', ratios: [], budgetPairs: [] };
  const saved = { metrics: [{ key: 'id', on: true }, { key: 'sales', aggregation: 'average', on: true }], dimensions: [], dateField: 'id', ratios: [{ key: 'bad', num: 'id', den: 'sales' }, { key: 'good', num: 'sales', den: 'budget' }], budgetPairs: [{ key: 'pair', actual: 'sales', budget: 'budget' }] };
  const restored = app.reconcileSavedConfig(config, saved, [{ key: 'date', type: 'date' }, { key: 'id', type: 'category' }]);
  assert.deepEqual(restored.metrics.map(metric => metric.key), ['sales', 'budget']);
  assert.equal(restored.metrics[0].aggregation, 'average');
  assert.equal(restored.dimensions[0].key, 'id');
  assert.equal(restored.dateField, 'date');
  assert.deepEqual(restored.ratios.map(ratio => ratio.key), ['good']);
  assert.deepEqual(restored.budgetPairs, saved.budgetPairs);
});
