(function () {
  'use strict';

  const NO_RESPONSE = 'No Response';
  const COLORS = [
    '#006b5f', '#d99b22', '#4b7f9f', '#c75050', '#6b8e4e',
    '#7f5aa2', '#2f9c95', '#9b6a35', '#4c647a', '#d06b9a',
    '#6c8fbd', '#b6a136', '#3a8d5d', '#875c74'
  ];

  const state = {
    workbook: null,
    fileName: '',
    sheetName: '',
    rows: [],
    columns: [],
    charts: [],
    nextChartNumber: 1
  };

  const els = {
    fileInput: document.getElementById('fileInput'),
    statusMessage: document.getElementById('statusMessage'),
    fileDetails: document.getElementById('fileDetails'),
    sheetPickerWrap: document.getElementById('sheetPickerWrap'),
    sheetSelect: document.getElementById('sheetSelect'),
    fileStats: document.getElementById('fileStats'),
    previewSection: document.getElementById('previewSection'),
    previewTable: document.getElementById('previewTable'),
    dashboardSection: document.getElementById('dashboardSection'),
    chartGrid: document.getElementById('chartGrid'),
    addChartBtn: document.getElementById('addChartBtn'),
    emptyState: document.getElementById('emptyState'),
    chartTemplate: document.getElementById('chartCardTemplate'),
    filterTemplate: document.getElementById('filterTemplate')
  };

  Chart.register(ChartDataLabels);
  Chart.defaults.font.family = 'Arial, Helvetica, sans-serif';

  els.fileInput.addEventListener('change', handleFileUpload);
  els.sheetSelect.addEventListener('change', () => loadSheet(els.sheetSelect.value));
  els.addChartBtn.addEventListener('click', () => addChart());

  async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const extension = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(extension)) {
      showStatus('Please choose an .xlsx, .xls, or .csv file.', 'error');
      return;
    }

    showStatus('Reading your file...', '');
    resetDataset();

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: 'array',
        cellDates: true,
        raw: false
      });

      if (!workbook.SheetNames.length) {
        throw new Error('No sheets were found in this file.');
      }

      state.workbook = workbook;
      state.fileName = file.name;
      populateSheetSelector(workbook.SheetNames);
      loadSheet(workbook.SheetNames[0]);

      if (state.rows.length > 50000) {
        showStatus('Large file warning: this app is designed for normal files up to about 50,000 rows. It may take longer to update charts.', 'warning');
      } else {
        showStatus('File loaded. Your data stays in this browser.', '');
      }
    } catch (error) {
      console.error(error);
      resetDataset();
      showStatus('This file could not be opened. It may be damaged or in an unsupported format.', 'error');
    }
  }

  function populateSheetSelector(sheetNames) {
    els.sheetSelect.innerHTML = sheetNames.map(name => `<option value="${escapeAttr(name)}">${escapeHtml(name)}</option>`).join('');
    els.sheetPickerWrap.classList.toggle('hidden', sheetNames.length <= 1);
  }

  function loadSheet(sheetName) {
    const sheet = state.workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false
    }).filter(row => row.some(cell => normalizeValue(cell) !== ''));

    state.sheetName = sheetName;
    state.rows = [];
    state.columns = [];

    if (!rawRows.length) {
      renderDataset();
      showStatus('The selected sheet does not contain usable rows.', 'warning');
      return;
    }

    state.columns = makeUniqueHeaders(rawRows[0]);
    state.rows = rawRows.slice(1)
      .filter(row => row.some(cell => normalizeValue(cell) !== ''))
      .map(row => {
        const record = {};
        state.columns.forEach((column, index) => {
          record[column] = row[index] === undefined ? '' : row[index];
        });
        return record;
      });

    state.charts = [];
    state.nextChartNumber = 1;
    renderDataset();
    if (state.columns.length && state.rows.length) addChart();
  }

  function makeUniqueHeaders(headerRow) {
    const seen = new Map();
    return headerRow.map((header, index) => {
      const base = normalizeValue(header) || `Column ${index + 1}`;
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      return count ? `${base} (${count + 1})` : base;
    });
  }

  function renderDataset() {
    const hasData = state.rows.length > 0 && state.columns.length > 0;
    els.emptyState.classList.toggle('hidden', hasData);
    els.fileDetails.classList.toggle('hidden', !state.workbook);
    els.previewSection.classList.toggle('hidden', !state.workbook);
    els.dashboardSection.classList.toggle('hidden', !hasData);

    renderFileStats();
    renderPreview();
    renderAllCharts();
  }

  function renderFileStats() {
    if (!state.workbook) return;
    const stats = [
      ['File name', state.fileName],
      ['Selected sheet', state.sheetName || 'None'],
      ['Total rows', formatNumber(state.rows.length)],
      ['Total columns', formatNumber(state.columns.length)]
    ];
    els.fileStats.innerHTML = stats.map(([label, value]) => `
      <div class="stat">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </div>
    `).join('');
  }

  function renderPreview() {
    if (!state.workbook) return;
    if (!state.rows.length || !state.columns.length) {
      els.previewTable.innerHTML = '<div class="chart-empty">No data is available in this sheet.</div>';
      return;
    }

    const rows = state.rows.slice(0, 20);
    els.previewTable.innerHTML = `
      <table>
        <thead><tr>${state.columns.map(column => `<th>${escapeHtml(column)}</th>`).join('')}</tr></thead>
        <tbody>
          ${rows.map(row => `
            <tr>${state.columns.map(column => `<td>${escapeHtml(displayCell(row[column]))}</td>`).join('')}</tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  function addChart(sourceConfig) {
    const chart = createChartConfig(sourceConfig);
    state.charts.push(chart);
    renderAllCharts();
  }

  function createChartConfig(sourceConfig) {
    const firstColumn = state.columns[0] || '';
    const config = sourceConfig ? cloneChartConfig(sourceConfig) : {
      id: makeId(),
      title: `Chart ${state.nextChartNumber}`,
      collapsed: false,
      primaryColumn: firstColumn,
      compareColumn: '',
      chartType: 'bar',
      compareType: 'grouped',
      compareValueMode: 'counts',
      sortMode: 'desc',
      topMode: 'all',
      showCounts: true,
      showPercentages: true,
      includeBlanks: true,
      filters: [],
      summarySearch: '',
      selectedResponses: new Set(),
      hiddenResponses: new Set(),
      merges: []
    };

    config.id = makeId();
    config.title = sourceConfig ? `${sourceConfig.title} copy` : config.title;
    state.nextChartNumber += 1;
    return config;
  }

  function cloneChartConfig(config) {
    return {
      ...config,
      filters: config.filters.map(filter => ({
        id: makeId(),
        column: filter.column,
        selected: new Set(filter.selected),
        search: filter.search || ''
      })),
      selectedResponses: new Set(),
      hiddenResponses: new Set(config.hiddenResponses),
      merges: config.merges.map(merge => ({ name: merge.name, sources: new Set(merge.sources) })),
      chartInstance: null
    };
  }

  function renderAllCharts() {
    els.chartGrid.innerHTML = '';
    state.charts.forEach(chart => {
      const card = renderChartCard(chart);
      els.chartGrid.appendChild(card);
      updateChartCard(chart, card);
    });
  }

  function renderChartCard(chart) {
    const fragment = els.chartTemplate.content.cloneNode(true);
    const card = fragment.querySelector('.chart-card');
    card.dataset.chartId = chart.id;
    card.classList.toggle('collapsed', chart.collapsed);

    const titleInput = card.querySelector('.chart-title-input');
    titleInput.value = chart.title;
    titleInput.addEventListener('input', event => {
      chart.title = event.target.value || 'Untitled chart';
    });

    card.querySelector('.duplicate-chart').addEventListener('click', () => addChart(chart));
    card.querySelector('.collapse-chart').addEventListener('click', () => {
      chart.collapsed = !chart.collapsed;
      renderAllCharts();
    });
    card.querySelector('.delete-chart').addEventListener('click', () => {
      if (confirm(`Delete "${chart.title}"?`)) {
        if (chart.chartInstance) chart.chartInstance.destroy();
        state.charts = state.charts.filter(item => item.id !== chart.id);
        renderAllCharts();
      }
    });

    populateColumnSelect(card.querySelector('.primary-column'), chart.primaryColumn, false);
    populateColumnSelect(card.querySelector('.compare-column'), chart.compareColumn, true);

    bindSelect(card, '.primary-column', chart, 'primaryColumn');
    bindSelect(card, '.compare-column', chart, 'compareColumn');
    bindSelect(card, '.chart-type', chart, 'chartType');
    bindSelect(card, '.compare-type', chart, 'compareType');
    bindSelect(card, '.compare-value-mode', chart, 'compareValueMode');
    bindSelect(card, '.sort-mode', chart, 'sortMode');
    bindSelect(card, '.top-mode', chart, 'topMode');
    bindCheckbox(card, '.show-counts', chart, 'showCounts');
    bindCheckbox(card, '.show-percentages', chart, 'showPercentages');
    bindCheckbox(card, '.include-blanks', chart, 'includeBlanks');

    card.querySelector('.summary-search').value = chart.summarySearch;
    card.querySelector('.summary-search').addEventListener('input', event => {
      chart.summarySearch = event.target.value;
      updateChartCard(chart, card);
    });

    card.querySelector('.add-filter').addEventListener('click', () => {
      chart.filters.push({
        id: makeId(),
        column: state.columns[0] || '',
        selected: new Set(),
        search: ''
      });
      renderAllCharts();
    });

    card.querySelector('.clear-filters').addEventListener('click', () => {
      chart.filters = [];
      renderAllCharts();
    });

    card.querySelector('.merge-selected').addEventListener('click', () => {
      const name = normalizeValue(card.querySelector('.merge-name').value);
      if (!name || !chart.selectedResponses.size) return;
      const sources = new Set();
      chart.selectedResponses.forEach(label => {
        sources.add(label);
        chart.merges.forEach(merge => {
          if (merge.name === label) merge.sources.forEach(source => sources.add(source));
        });
      });
      chart.merges = chart.merges.filter(merge => !chart.selectedResponses.has(merge.name));
      chart.merges.push({ name, sources });
      chart.selectedResponses.clear();
      card.querySelector('.merge-name').value = '';
      updateChartCard(chart, card);
    });

    card.querySelector('.hide-selected').addEventListener('click', () => {
      chart.selectedResponses.forEach(value => chart.hiddenResponses.add(value));
      chart.selectedResponses.clear();
      updateChartCard(chart, card);
    });

    card.querySelector('.reset-responses').addEventListener('click', () => {
      chart.hiddenResponses.clear();
      chart.merges = [];
      chart.selectedResponses.clear();
      updateChartCard(chart, card);
    });

    card.querySelector('.export-png').addEventListener('click', () => exportChartPng(chart));
    card.querySelector('.export-summary').addEventListener('click', () => exportSummaryCsv(chart));
    card.querySelector('.export-filtered').addEventListener('click', () => exportFilteredDataCsv(chart));

    return card;
  }

  function bindSelect(card, selector, chart, key) {
    const input = card.querySelector(selector);
    input.value = chart[key];
    input.addEventListener('change', event => {
      chart[key] = event.target.value;
      chart.selectedResponses.clear();
      renderAllCharts();
    });
  }

  function bindCheckbox(card, selector, chart, key) {
    const input = card.querySelector(selector);
    input.checked = chart[key];
    input.addEventListener('change', event => {
      chart[key] = event.target.checked;
      updateChartCard(chart, card);
    });
  }

  function populateColumnSelect(select, selectedValue, includeNone) {
    const options = includeNone ? ['<option value="">No comparison</option>'] : [];
    options.push(...state.columns.map(column => `<option value="${escapeAttr(column)}">${escapeHtml(column)}</option>`));
    select.innerHTML = options.join('');
    select.value = selectedValue;
  }

  function updateChartCard(chart, card) {
    if (!state.rows.length || !state.columns.length) return;

    const isComparison = Boolean(chart.compareColumn);
    card.querySelectorAll('.single-setting').forEach(el => el.classList.toggle('hidden', isComparison));
    card.querySelectorAll('.compare-setting').forEach(el => el.classList.toggle('hidden', !isComparison));
    card.querySelector('.collapse-chart').textContent = chart.collapsed ? 'Show' : 'Hide';

    renderFilters(chart, card);

    const filteredRows = applyFilters(state.rows, chart.filters);
    card.querySelector('.row-count').textContent = `Showing ${formatNumber(filteredRows.length)} of ${formatNumber(state.rows.length)} rows`;

    const result = isComparison
      ? buildComparisonResult(filteredRows, chart)
      : buildSingleColumnResult(filteredRows, chart);

    renderChart(chart, card, result);
    renderSummaryTable(chart, card, result);
  }

  function renderFilters(chart, card) {
    const list = card.querySelector('.filters-list');
    list.innerHTML = '';

    chart.filters.forEach(filter => {
      if (!filter.column || !state.columns.includes(filter.column)) filter.column = state.columns[0] || '';
      const fragment = els.filterTemplate.content.cloneNode(true);
      const filterCard = fragment.querySelector('.filter-card');
      const columnSelect = filterCard.querySelector('.filter-column');
      populateColumnSelect(columnSelect, filter.column, false);

      columnSelect.addEventListener('change', event => {
        filter.column = event.target.value;
        filter.selected.clear();
        filter.search = '';
        renderAllCharts();
      });

      filterCard.querySelector('.remove-filter').addEventListener('click', () => {
        chart.filters = chart.filters.filter(item => item.id !== filter.id);
        renderAllCharts();
      });

      const search = filterCard.querySelector('.filter-search');
      search.value = filter.search || '';
      search.addEventListener('input', event => {
        filter.search = event.target.value;
        renderAllCharts();
      });

      const values = getUniqueValues(state.rows, filter.column);
      const searchText = (filter.search || '').toLowerCase();
      const visibleValues = values.filter(value => value.toLowerCase().includes(searchText));
      const valuesWrap = filterCard.querySelector('.filter-values');
      valuesWrap.innerHTML = visibleValues.map(value => {
        const id = `${filter.id}-${hashString(value)}`;
        return `
          <label for="${id}">
            <input id="${id}" type="checkbox" value="${escapeAttr(value)}" ${filter.selected.has(value) ? 'checked' : ''}>
            <span>${escapeHtml(value)}</span>
          </label>
        `;
      }).join('') || '<p class="muted">No matching values.</p>';

      valuesWrap.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', event => {
          if (event.target.checked) filter.selected.add(event.target.value);
          else filter.selected.delete(event.target.value);
          updateChartCard(chart, card);
        });
      });

      list.appendChild(filterCard);
    });
  }

  function applyFilters(rows, filters) {
    const activeFilters = filters.filter(filter => filter.column && filter.selected.size);
    if (!activeFilters.length) return rows;

    return rows.filter(row => activeFilters.every(filter => {
      const value = getResponseLabel(row[filter.column]);
      return filter.selected.has(value);
    }));
  }

  function buildSingleColumnResult(rows, chart) {
    const counts = new Map();
    let nonBlank = 0;

    rows.forEach(row => {
      const originalLabel = getResponseLabel(row[chart.primaryColumn]);
      if (originalLabel !== NO_RESPONSE) nonBlank += 1;
      if (!chart.includeBlanks && originalLabel === NO_RESPONSE) return;
      const label = getMergedLabel(originalLabel, chart.merges);
      if (chart.hiddenResponses.has(label)) return;
      counts.set(label, (counts.get(label) || 0) + 1);
    });

    let items = Array.from(counts, ([response, count]) => ({
      response,
      count,
      rowPercent: rows.length ? roundOne((count / rows.length) * 100) : 0,
      nonBlankPercent: nonBlank ? roundOne((count / nonBlank) * 100) : 0
    }));

    items = sortItems(items, chart.sortMode);
    items = applyTopGrouping(items, chart.topMode, rows.length, nonBlank);

    return {
      type: 'single',
      rows,
      labels: items.map(item => item.response),
      values: items.map(item => item.count),
      items
    };
  }

  function buildComparisonResult(rows, chart) {
    const matrix = new Map();
    const comparisonLabels = new Set();
    let total = 0;

    rows.forEach(row => {
      const primary = getMergedLabel(getResponseLabel(row[chart.primaryColumn]), chart.merges);
      const comparison = getResponseLabel(row[chart.compareColumn]);
      if (chart.hiddenResponses.has(primary)) return;
      if (!matrix.has(primary)) matrix.set(primary, new Map());
      matrix.get(primary).set(comparison, (matrix.get(primary).get(comparison) || 0) + 1);
      comparisonLabels.add(comparison);
      total += 1;
    });

    const primaryLabels = Array.from(matrix.keys()).sort((a, b) => a.localeCompare(b));
    const compareLabels = Array.from(comparisonLabels).sort((a, b) => a.localeCompare(b));
    const primaryTotals = new Map(primaryLabels.map(label => [
      label,
      compareLabels.reduce((sum, compare) => sum + (matrix.get(label).get(compare) || 0), 0)
    ]));
    const compareTotals = new Map(compareLabels.map(label => [
      label,
      primaryLabels.reduce((sum, primary) => sum + (matrix.get(primary).get(label) || 0), 0)
    ]));

    const valueMode = chart.compareType === 'stacked100' ? 'primaryPercent' : chart.compareValueMode;
    const datasets = compareLabels.map((compare, index) => ({
      label: compare,
      data: primaryLabels.map(primary => {
        const count = matrix.get(primary).get(compare) || 0;
        if (valueMode === 'primaryPercent') return primaryTotals.get(primary) ? roundOne((count / primaryTotals.get(primary)) * 100) : 0;
        if (valueMode === 'comparePercent') return compareTotals.get(compare) ? roundOne((count / compareTotals.get(compare)) * 100) : 0;
        if (valueMode === 'totalPercent') return total ? roundOne((count / total) * 100) : 0;
        return count;
      }),
      backgroundColor: COLORS[index % COLORS.length],
      borderColor: COLORS[index % COLORS.length],
      borderWidth: 1
    }));

    return {
      type: 'comparison',
      rows,
      labels: primaryLabels,
      compareLabels,
      matrix,
      primaryTotals,
      compareTotals,
      total,
      datasets,
      valueMode
    };
  }

  function renderChart(chart, card, result) {
    const canvas = card.querySelector('canvas');
    const empty = card.querySelector('.chart-empty');
    const chartArea = card.querySelector('.chart-area');
    const hasData = result.type === 'single'
      ? result.items.length > 0
      : result.labels.length > 0 && result.compareLabels.length > 0;

    if (chart.chartInstance) {
      chart.chartInstance.destroy();
      chart.chartInstance = null;
    }

    empty.classList.toggle('hidden', hasData);
    const tableOnly = result.type === 'comparison' ? chart.compareType === 'table' : chart.chartType === 'table';
    chartArea.classList.toggle('hidden', tableOnly);
    if (!hasData || tableOnly) return;

    const context = canvas.getContext('2d');
    const config = result.type === 'comparison'
      ? getComparisonChartConfig(chart, result)
      : getSingleChartConfig(chart, result);
    chart.chartInstance = new Chart(context, config);
  }

  function getSingleChartConfig(chart, result) {
    const type = chart.chartType === 'horizontalBar' ? 'bar' : chart.chartType;
    const values = result.values;

    return {
      type,
      data: {
        labels: result.labels,
        datasets: [{
          label: 'Responses',
          data: values,
          backgroundColor: result.labels.map((_, index) => COLORS[index % COLORS.length]),
          borderColor: result.labels.map((_, index) => COLORS[index % COLORS.length]),
          borderWidth: 1,
          tension: 0.25,
          fill: chart.chartType === 'line' ? false : true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: chart.chartType === 'horizontalBar' ? 'y' : 'x',
        plugins: {
          legend: { display: ['pie', 'doughnut'].includes(chart.chartType) },
          tooltip: {
            callbacks: {
              label: context => {
                const item = result.items[context.dataIndex];
                return `${item.response}: ${formatNumber(item.count)} (${item.rowPercent}%)`;
              }
            }
          },
          datalabels: {
            color: '#1d2733',
            anchor: 'end',
            align: type === 'line' ? 'top' : 'end',
            formatter: (value, context) => {
              const item = result.items[context.dataIndex];
              const parts = [];
              if (chart.showCounts) parts.push(formatNumber(value));
              if (chart.showPercentages) parts.push(`${item.rowPercent}%`);
              return parts.join(' | ');
            }
          }
        },
        scales: ['pie', 'doughnut'].includes(chart.chartType) ? {} : {
          x: { beginAtZero: true },
          y: { beginAtZero: true }
        }
      }
    };
  }

  function getComparisonChartConfig(chart, result) {
    const stacked = chart.compareType === 'stacked' || chart.compareType === 'stacked100';
    return {
      type: 'bar',
      data: {
        labels: result.labels,
        datasets: result.datasets
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'bottom' },
          tooltip: {
            callbacks: {
              label: context => `${context.dataset.label}: ${context.formattedValue}${result.valueMode === 'counts' ? '' : '%'}`
            }
          },
          datalabels: {
            color: '#1d2733',
            anchor: 'end',
            align: 'end',
            formatter: value => value ? `${value}${result.valueMode === 'counts' ? '' : '%'}` : ''
          }
        },
        scales: {
          x: { stacked },
          y: {
            stacked,
            beginAtZero: true,
            max: chart.compareType === 'stacked100' || result.valueMode !== 'counts' ? 100 : undefined
          }
        }
      }
    };
  }

  function renderSummaryTable(chart, card, result) {
    if (result.type === 'comparison') {
      renderComparisonTable(chart, card, result);
      return;
    }

    const searchText = chart.summarySearch.trim().toLowerCase();
    const visibleItems = result.items.filter(item => item.response.toLowerCase().includes(searchText));
    const html = `
      <table>
        <thead>
          <tr>
            <th>Select</th>
            <th>Response</th>
            <th class="number">Count</th>
            <th class="number">Percentage</th>
          </tr>
        </thead>
        <tbody>
          ${visibleItems.map(item => `
            <tr>
              <td>
                <label class="selected-cell">
                  <input type="checkbox" class="response-select" value="${escapeAttr(item.response)}" ${chart.selectedResponses.has(item.response) ? 'checked' : ''}>
                  <span>Select</span>
                </label>
              </td>
              <td>${escapeHtml(item.response)}</td>
              <td class="number">${formatNumber(item.count)}</td>
              <td class="number">${item.rowPercent}%</td>
            </tr>
          `).join('') || '<tr><td colspan="4">No matching responses.</td></tr>'}
        </tbody>
      </table>
    `;
    card.querySelector('.summary-table').innerHTML = html;
    card.querySelectorAll('.response-select').forEach(input => {
      input.addEventListener('change', event => {
        if (event.target.checked) chart.selectedResponses.add(event.target.value);
        else chart.selectedResponses.delete(event.target.value);
      });
    });
  }

  function renderComparisonTable(chart, card, result) {
    const mode = result.valueMode || chart.compareValueMode;
    const suffix = mode === 'counts' ? '' : '%';
    const rows = result.labels.map(primary => {
      const total = result.primaryTotals.get(primary) || 0;
      const cells = result.compareLabels.map(compare => {
        const count = result.matrix.get(primary).get(compare) || 0;
        let value = count;
        if (mode === 'primaryPercent') value = total ? roundOne((count / total) * 100) : 0;
        if (mode === 'comparePercent') value = result.compareTotals.get(compare) ? roundOne((count / result.compareTotals.get(compare)) * 100) : 0;
        if (mode === 'totalPercent') value = result.total ? roundOne((count / result.total) * 100) : 0;
        return `<td class="number">${escapeHtml(formatTableValue(value, suffix))}</td>`;
      }).join('');
      return `
        <tr>
          <td>${escapeHtml(primary)}</td>
          ${cells}
          <td class="number">${formatNumber(total)}</td>
        </tr>
      `;
    }).join('');

    card.querySelector('.summary-table').innerHTML = `
      <table>
        <thead>
          <tr>
            <th>${escapeHtml(chart.primaryColumn)}</th>
            ${result.compareLabels.map(label => `<th class="number">${escapeHtml(label)}</th>`).join('')}
            <th class="number">Total</th>
          </tr>
        </thead>
        <tbody>${rows || `<tr><td colspan="${result.compareLabels.length + 2}">No data is available.</td></tr>`}</tbody>
      </table>
    `;
  }

  function sortItems(items, sortMode) {
    return [...items].sort((a, b) => {
      if (sortMode === 'asc') return a.count - b.count || a.response.localeCompare(b.response);
      if (sortMode === 'alpha') return a.response.localeCompare(b.response);
      return b.count - a.count || a.response.localeCompare(b.response);
    });
  }

  function applyTopGrouping(items, topMode, totalRows, nonBlankRows) {
    const topCount = Number(topMode);
    if (!topCount || items.length <= topCount) return items;

    const topItems = items.slice(0, topCount);
    const otherCount = items.slice(topCount).reduce((sum, item) => sum + item.count, 0);
    if (otherCount > 0) {
      topItems.push({
        response: 'Other',
        count: otherCount,
        rowPercent: totalRows ? roundOne((otherCount / totalRows) * 100) : 0,
        nonBlankPercent: nonBlankRows ? roundOne((otherCount / nonBlankRows) * 100) : 0
      });
    }
    return topItems;
  }

  function getUniqueValues(rows, column) {
    return Array.from(new Set(rows.map(row => getResponseLabel(row[column]))))
      .sort((a, b) => a.localeCompare(b));
  }

  function getResponseLabel(value) {
    const normalized = normalizeValue(value);
    return normalized === '' ? NO_RESPONSE : normalized;
  }

  function getMergedLabel(label, merges) {
    const match = merges.find(merge => merge.sources.has(label));
    return match ? match.name : label;
  }

  function normalizeValue(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return value.toLocaleDateString();
    if (typeof value === 'string') return value.trim();
    return String(value).trim();
  }

  function displayCell(value) {
    const normalized = normalizeValue(value);
    return normalized === '' ? '' : normalized;
  }

  function exportChartPng(chart) {
    if (!chart.chartInstance) {
      alert('This chart is currently shown as a table only.');
      return;
    }
    const link = document.createElement('a');
    link.download = `${safeFileName(chart.title)}.png`;
    link.href = chart.chartInstance.toBase64Image('image/png', 1);
    link.click();
  }

  function exportSummaryCsv(chart) {
    const rows = applyFilters(state.rows, chart.filters);
    const result = chart.compareColumn
      ? buildComparisonResult(rows, chart)
      : buildSingleColumnResult(rows, chart);
    const csvRows = [];

    if (result.type === 'comparison') {
      csvRows.push([chart.primaryColumn, ...result.compareLabels, 'Total']);
      result.labels.forEach(primary => {
        csvRows.push([
          primary,
          ...result.compareLabels.map(compare => result.matrix.get(primary).get(compare) || 0),
          result.primaryTotals.get(primary) || 0
        ]);
      });
    } else {
      csvRows.push(['Response', 'Count', 'Percentage']);
      result.items.forEach(item => csvRows.push([item.response, item.count, `${item.rowPercent}%`]));
    }

    downloadCsv(csvRows, `${safeFileName(chart.title)}-summary.csv`);
  }

  function exportFilteredDataCsv(chart) {
    const rows = applyFilters(state.rows, chart.filters);
    const csvRows = [state.columns, ...rows.map(row => state.columns.map(column => displayCell(row[column])))];
    downloadCsv(csvRows, `${safeFileName(chart.title)}-filtered-data.csv`);
  }

  function downloadCsv(rows, fileName) {
    const csv = rows.map(row => row.map(csvEscape).join(',')).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  function csvEscape(value) {
    const text = String(value ?? '');
    return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  function resetDataset() {
    state.workbook = null;
    state.fileName = '';
    state.sheetName = '';
    state.rows = [];
    state.columns = [];
    state.charts.forEach(chart => {
      if (chart.chartInstance) chart.chartInstance.destroy();
    });
    state.charts = [];
    state.nextChartNumber = 1;
    els.chartGrid.innerHTML = '';
    els.fileStats.innerHTML = '';
    els.previewTable.innerHTML = '';
    renderDataset();
  }

  function showStatus(message, type) {
    els.statusMessage.textContent = message;
    els.statusMessage.className = `status-message ${type || ''}`.trim();
  }

  function makeId() {
    return `id-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  }

  function hashString(value) {
    let hash = 0;
    for (let index = 0; index < value.length; index += 1) {
      hash = ((hash << 5) - hash) + value.charCodeAt(index);
      hash |= 0;
    }
    return Math.abs(hash).toString(36);
  }

  function formatNumber(value) {
    return Number(value || 0).toLocaleString();
  }

  function roundOne(value) {
    return Math.round(value * 10) / 10;
  }

  function formatTableValue(value, suffix) {
    return suffix ? `${value}${suffix}` : formatNumber(value);
  }

  function safeFileName(value) {
    return normalizeValue(value).replace(/[\\/:*?"<>|]+/g, '-').slice(0, 80) || 'chart';
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }
})();
