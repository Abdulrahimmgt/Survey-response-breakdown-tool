(function (root, factory) {
  root.ChartRules = factory();
})(typeof globalThis !== 'undefined' ? globalThis : window, function () {
  'use strict';

  const DEFAULT_MAX_UNIQUE_VALUES = 50;

  function text(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toISOString().slice(0, 10);
    return String(value).trim();
  }

  function answerLabels(value) {
    const values = Array.isArray(value) ? value.flatMap(answerLabels) : [text(value)];
    return values.filter(Boolean);
  }

  function normalizeAnswer(value) {
    return text(value)
      .toLocaleLowerCase()
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function countUniqueAnswers(rows, column) {
    const values = new Set();
    rows.forEach(row => {
      answerLabels(row[column]).forEach(value => values.add(normalizeAnswer(value)));
    });
    values.delete('');
    return values.size;
  }

  function isLikelyMetadataColumn(column) {
    return /^(id|student id|response id|tempid)$/i.test(text(column))
      || /date|timestamp|email|name|phone|address/i.test(text(column));
  }

  function getEligibleChartColumns(rows, columns, options = {}) {
    const maxUniqueValues = options.maxUniqueValues ?? DEFAULT_MAX_UNIQUE_VALUES;
    const hiddenColumns = options.hiddenColumns instanceof Set ? options.hiddenColumns : new Set(options.hiddenColumns || []);
    const includeMetadata = options.includeMetadata === true;

    return columns.filter(column => {
      if (hiddenColumns.has(column)) return false;
      if (!includeMetadata && isLikelyMetadataColumn(column)) return false;
      const uniqueCount = countUniqueAnswers(rows, column);
      return uniqueCount > 0 && uniqueCount <= maxUniqueValues;
    });
  }

  function getMissingChartColumns(eligibleColumns, charts) {
    const existingColumns = new Set(charts.map(chart => chart.primaryColumn).filter(Boolean));
    return eligibleColumns.filter(column => !existingColumns.has(column));
  }

  return {
    DEFAULT_MAX_UNIQUE_VALUES,
    countUniqueAnswers,
    getEligibleChartColumns,
    getMissingChartColumns,
    isLikelyMetadataColumn
  };
});
