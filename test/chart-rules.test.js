import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import vm from 'node:vm';

const source = await readFile(new URL('../chart-rules.js', import.meta.url), 'utf8');
const context = { globalThis: {} };
vm.runInNewContext(source, context);
const ChartRules = context.globalThis.ChartRules;

test('includes answered questions with no more than 50 unique values', () => {
  const columns = ['Question A', 'Question B', 'Question C'];
  const rows = Array.from({ length: 51 }, (_, index) => ({
    'Question A': index % 2 ? 'Yes' : 'No',
    'Question B': `Value ${index + 1}`,
    'Question C': index === 0 ? 'Only answer' : ''
  }));

  assert.deepEqual(
    Array.from(ChartRules.getEligibleChartColumns(rows, columns)),
    ['Question A', 'Question C']
  );
});

test('counts normalized response labels when applying the unique-value limit', () => {
  const rows = [
    { Question: '  YES ' },
    { Question: 'yes' },
    { Question: 'No' }
  ];

  assert.deepEqual(Array.from(ChartRules.getEligibleChartColumns(rows, ['Question'])), ['Question']);
  assert.equal(ChartRules.countUniqueAnswers(rows, 'Question'), 2);
});

test('excludes empty and likely metadata columns from automatic charts', () => {
  const rows = [
    { 'Student ID': '1', Email: 'a@example.com', Timestamp: '2026-01-01', Name: 'A', Empty: '', Rating: 'Good' },
    { 'Student ID': '2', Email: 'b@example.com', Timestamp: '2026-01-02', Name: 'B', Empty: '', Rating: 'Great' }
  ];

  assert.deepEqual(
    Array.from(ChartRules.getEligibleChartColumns(rows, Object.keys(rows[0]))),
    ['Rating']
  );
});

test('returns only eligible columns that do not already have a chart', () => {
  const eligible = ['Question A', 'Question B', 'Question C'];
  const charts = [{ primaryColumn: 'Question A' }, { primaryColumn: 'Question C' }];

  assert.deepEqual(
    Array.from(ChartRules.getMissingChartColumns(eligible, charts)),
    ['Question B']
  );
});

test('supports linked multi-select values when determining eligibility', () => {
  const rows = [
    { Strategies: ['Tutoring', 'Mentoring'] },
    { Strategies: ['Mentoring'] }
  ];

  assert.equal(ChartRules.countUniqueAnswers(rows, 'Strategies'), 2);
  assert.deepEqual(Array.from(ChartRules.getEligibleChartColumns(rows, ['Strategies'])), ['Strategies']);
});
