(function () {
  'use strict';

  const NO_RESPONSE = 'No Response';
  const TABLE_ROW_LIMIT = 10;
  const REPORT_UNIQUE_VALUE_LIMIT = 15;
  const REPORT_FILTER_UNIQUE_VALUE_LIMIT = 500;
  const UPLOADED_SOURCE_ID = 'uploaded-workbook';
  const COLORS = [
    '#006b5f', '#d99b22', '#4b7f9f', '#c75050', '#6b8e4e',
    '#7f5aa2', '#2f9c95', '#9b6a35', '#4c647a', '#d06b9a',
    '#6c8fbd', '#b6a136', '#3a8d5d', '#875c74'
  ];
  const ANSWER_SETS = [
    ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree'],
    ['Not at all', 'Not really', 'Kind of', 'Definitely', 'Absolutely'],
    ['Never', 'Almost never', 'Sometimes', 'Lots of times', 'All the time'],
    ['1: Bad', '2: Okay', '3: Good', '4: Great', '5: Amazing'],
    ['1: Not confident at all', '2: Slightly confident', '3: Somewhat confident', '4: Very confident', '5: Extremely confident'],
    ['Poor', 'Fair', 'Good', 'Very Good', 'Excellent'],
    ['Not well at all', 'Slightly well', 'Somewhat well', 'Quite well', 'Extremely well'],
    ['Not at all', 'A little bit', 'Somewhat well', 'Quite well', 'Extremely well'],
    Array.from({ length: 10 }, (_, index) => String(index + 1)),
    ['Yes', 'No', "I don't know"],
    ['Yes', 'Maybe', 'No'],
    ['Very low extent', 'Low extent', 'Moderate extent', 'Great extent', 'Very great extent'],
    ['Not informed at all', 'Slightly informed', 'Somewhat informed', 'Quite informed', 'Extremely informed'],
    ['Not effective at all', 'Slightly effective', 'Somewhat effective', 'Quite effective', 'Extremely effective'],
    ['No growth at all', 'Slight growth', 'Some growth', 'A lot of growth', 'Substantial growth', "I don't know"],
    ['Never', 'Rarely', 'Occasionally', 'A moderate amount', 'A great deal'],
    ['Strongly disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly agree', 'Not applicable'],
    ['Not at all', 'Not really', 'Kind of', 'Definitely', 'Absolutely', "I didn't have any of these classes or activities"],
    ['0-10%', '11-25%', '26-50%', '51-75%', '76-100%'],
    ['I did not provide this support', '1 time during the program', '2-3 times during the program', 'Weekly', 'Daily'],
    ['Not helpful', 'Somewhat helpful', 'Helpful', 'Very Helpful'],
    ['Strongly disagree', 'Disagree', 'Neither agree nor disagree', 'Agree', 'Strongly agree'],
    ['Not Challenging', 'Somewhat Challenging', 'Very Challenging'],
    ['Selected', 'Not Selected'],
    ['Community or local organizations', 'Education newsletter or email blast', 'Grant database or website', 'Professional network', 'Social media', 'Other [please specify]'],
    ['Not at all', 'A little', 'Somewhat', 'Quite a bit', 'A great deal']
  ];

  const state = {
    workbook: null,
    fileName: '',
    sheetName: '',
    rows: [],
    columns: [],
    charts: [],
    nextChartNumber: 1,
    sources: [],
    reportResult: null
  };

  const els = {
    fileInput: document.getElementById('fileInput'),
    statusMessage: document.getElementById('statusMessage'),
    fileDetails: document.getElementById('fileDetails'),
    sheetPickerWrap: document.getElementById('sheetPickerWrap'),
    sheetSelect: document.getElementById('sheetSelect'),
    fileStats: document.getElementById('fileStats'),
    dashboardSection: document.getElementById('dashboardSection'),
    chartGrid: document.getElementById('chartGrid'),
    addChartBtn: document.getElementById('addChartBtn'),
    emptyState: document.getElementById('emptyState'),
    chartTemplate: document.getElementById('chartCardTemplate'),
    filterTemplate: document.getElementById('filterTemplate'),
    reportSourceSelect: document.getElementById('reportSourceSelect'),
    publicSheetUrl: document.getElementById('publicSheetUrl'),
    loadPublicSheetBtn: document.getElementById('loadPublicSheetBtn'),
    reportStatus: document.getElementById('reportStatus'),
    reportNameInput: document.getElementById('reportNameInput'),
    reportDataSheetSelect: document.getElementById('reportDataSheetSelect'),
    questionChecklist: document.getElementById('questionChecklist'),
    reportColumnNote: document.getElementById('reportColumnNote'),
    primaryBreakdownSelect: document.getElementById('primaryBreakdownSelect'),
    reportFilterColumnSelect: document.getElementById('reportFilterColumnSelect'),
    reportFilterValues: document.getElementById('reportFilterValues'),
    reportFilterNote: document.getElementById('reportFilterNote'),
    generateReportBtn: document.getElementById('generateReportBtn'),
    downloadReportCsvBtn: document.getElementById('downloadReportCsvBtn'),
    downloadReportXlsxBtn: document.getElementById('downloadReportXlsxBtn'),
    reportOutputTitle: document.getElementById('reportOutputTitle'),
    reportOutputMeta: document.getElementById('reportOutputMeta'),
    distributionOutput: document.getElementById('distributionOutput')
  };

  Chart.register(ChartDataLabels);
  Chart.defaults.font.family = 'Arial, Helvetica, sans-serif';

  els.fileInput.addEventListener('change', handleFileUpload);
  els.sheetSelect.addEventListener('change', () => loadSheet(els.sheetSelect.value));
  els.addChartBtn.addEventListener('click', () => addChart());
  els.loadPublicSheetBtn.addEventListener('click', loadPublicGoogleSheet);
  els.reportSourceSelect.addEventListener('change', renderReportControls);
  els.reportDataSheetSelect.addEventListener('change', renderReportColumns);
  els.primaryBreakdownSelect.addEventListener('change', syncBreakdownQuestionSelection);
  els.reportFilterColumnSelect.addEventListener('change', renderReportFilterValues);
  els.generateReportBtn.addEventListener('click', generateDistributionReport);
  els.downloadReportCsvBtn.addEventListener('click', downloadDistributionCsv);
  els.downloadReportXlsxBtn.addEventListener('click', downloadDistributionXlsx);

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

      upsertReportSource(UPLOADED_SOURCE_ID, `Uploaded: ${file.name}`, workbook);
      activateWorkbook(workbook, file.name, workbook.SheetNames[0]);

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

  function activateWorkbook(workbook, fileName, sheetName) {
    state.workbook = workbook;
    state.fileName = fileName;
    populateSheetSelector(workbook.SheetNames);
    loadSheet(sheetName || workbook.SheetNames[0]);
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
    els.dashboardSection.classList.toggle('hidden', !hasData);

    renderFileStats();
    renderAllCharts();
    renderReportControls();
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
      topMode: '10',
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
      if (!chart.includeBlanks && (primary === NO_RESPONSE || comparison === NO_RESPONSE)) return;
      if (chart.hiddenResponses.has(primary)) return;
      if (!matrix.has(primary)) matrix.set(primary, new Map());
      matrix.get(primary).set(comparison, (matrix.get(primary).get(comparison) || 0) + 1);
      comparisonLabels.add(comparison);
      total += 1;
    });

    const rawPrimaryLabels = Array.from(matrix.keys());
    const rawCompareLabels = Array.from(comparisonLabels);
    const rawPrimaryTotals = new Map(rawPrimaryLabels.map(label => [
      label,
      rawCompareLabels.reduce((sum, compare) => sum + (matrix.get(label).get(compare) || 0), 0)
    ]));
    const rawCompareTotals = new Map(rawCompareLabels.map(label => [
      label,
      rawPrimaryLabels.reduce((sum, primary) => sum + (matrix.get(primary).get(label) || 0), 0)
    ]));
    const primaryLabels = capLabels(rawPrimaryLabels, rawPrimaryTotals, TABLE_ROW_LIMIT);
    const compareLabels = capLabels(rawCompareLabels, rawCompareTotals, TABLE_ROW_LIMIT);
    const cappedMatrix = new Map(primaryLabels.map(label => [label, new Map(compareLabels.map(compare => [compare, 0]))]));

    rawPrimaryLabels.forEach(primary => {
      const primaryLabel = primaryLabels.includes(primary) ? primary : 'Other';
      rawCompareLabels.forEach(compare => {
        const compareLabel = compareLabels.includes(compare) ? compare : 'Other';
        const count = matrix.get(primary).get(compare) || 0;
        cappedMatrix.get(primaryLabel).set(compareLabel, cappedMatrix.get(primaryLabel).get(compareLabel) + count);
      });
    });

    const primaryTotals = new Map(primaryLabels.map(label => [
      label,
      compareLabels.reduce((sum, compare) => sum + (cappedMatrix.get(label).get(compare) || 0), 0)
    ]));
    const compareTotals = new Map(compareLabels.map(label => [
      label,
      primaryLabels.reduce((sum, primary) => sum + (cappedMatrix.get(primary).get(label) || 0), 0)
    ]));

    const valueMode = chart.compareType === 'stacked100' ? 'primaryPercent' : chart.compareValueMode;
    const datasets = compareLabels.map((compare, index) => ({
      label: compare,
      data: primaryLabels.map(primary => {
        const count = cappedMatrix.get(primary).get(compare) || 0;
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
      matrix: cappedMatrix,
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
            color: ['pie', 'doughnut'].includes(chart.chartType) ? '#ffffff' : '#1d2733',
            anchor: ['pie', 'doughnut'].includes(chart.chartType) ? 'center' : 'end',
            align: ['pie', 'doughnut'].includes(chart.chartType) ? 'center' : (type === 'line' ? 'top' : 'end'),
            clamp: true,
            clip: false,
            textAlign: 'center',
            font: context => ({
              weight: '700',
              size: ['pie', 'doughnut'].includes(chart.chartType) && context.dataset.data.length > 8 ? 10 : 11
            }),
            formatter: (value, context) => {
              const item = result.items[context.dataIndex];
              if (['pie', 'doughnut'].includes(chart.chartType) && item.rowPercent < 4) return '';
              const parts = [];
              if (chart.showCounts) parts.push(formatNumber(value));
              if (chart.showPercentages) parts.push(`${item.rowPercent}%`);
              return ['pie', 'doughnut'].includes(chart.chartType) ? parts.join('\n') : parts.join(' | ');
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
    const matchingItems = result.items.filter(item => item.response.toLowerCase().includes(searchText));
    const visibleItems = matchingItems.slice(0, TABLE_ROW_LIMIT);
    renderTableNote(card, matchingItems.length, 'responses');
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
    renderTableNote(card, result.labels.length, 'comparison rows');
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

  function renderTableNote(card, totalRows, label) {
    const note = card.querySelector('.table-note');
    if (!note) return;
    note.textContent = totalRows > TABLE_ROW_LIMIT
      ? `Showing first ${TABLE_ROW_LIMIT} ${label}.`
      : '';
  }

  function capLabels(labels, totals, limit) {
    const sorted = [...labels].sort((a, b) => (totals.get(b) || 0) - (totals.get(a) || 0) || a.localeCompare(b));
    if (sorted.length <= limit) return sorted;
    const topLabels = sorted.slice(0, limit - 1);
    return topLabels.includes('Other') ? topLabels : [...topLabels, 'Other'];
  }

  function sortItems(items, sortMode) {
    return [...items].sort((a, b) => {
      if (sortMode === 'asc') return a.count - b.count || a.response.localeCompare(b.response);
      if (sortMode === 'alpha') return a.response.localeCompare(b.response);
      return b.count - a.count || a.response.localeCompare(b.response);
    });
  }

  function applyTopGrouping(items, topMode, totalRows, nonBlankRows) {
    const topCount = Math.min(Number(topMode) || TABLE_ROW_LIMIT, TABLE_ROW_LIMIT);
    if (items.length <= topCount) return items;

    const topItems = items.slice(0, topCount - 1);
    const otherCount = items.slice(topCount - 1).reduce((sum, item) => sum + item.count, 0);
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

  async function loadPublicGoogleSheet() {
    const url = normalizeValue(els.publicSheetUrl.value);
    const sheetId = extractGoogleSheetId(url);
    if (!sheetId) {
      showReportStatus('Paste a valid public Google Sheets link.', 'error');
      return;
    }

    showReportStatus('Loading public Google Sheet...', '');
    try {
      const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
      const response = await fetch(exportUrl);
      if (!response.ok) throw new Error(`Google Sheets returned ${response.status}`);
      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: 'array',
        cellDates: true,
        raw: false
      });
      if (!workbook.SheetNames.length) throw new Error('No sheets found');
      const sourceName = `Google Sheet: ${sheetId.slice(0, 8)}...`;
      upsertReportSource(`google-${sheetId}`, sourceName, workbook);
      activateWorkbook(workbook, sourceName, workbook.SheetNames[0]);
      els.reportSourceSelect.value = `google-${sheetId}`;
      renderReportControls();
      showStatus('Public Google Sheet loaded. Your charts now use this data source.', '');
      showReportStatus('Public Google Sheet loaded. It is used only in this browser session.', '');
    } catch (error) {
      console.error(error);
      showReportStatus('Could not load that Google Sheet. Make sure it is public or shared with anyone who has the link.', 'error');
    }
  }

  function upsertReportSource(id, name, workbook) {
    const existing = state.sources.find(source => source.id === id);
    if (existing) {
      existing.name = name;
      existing.workbook = workbook;
    } else {
      state.sources.push({ id, name, workbook });
    }
    state.reportResult = null;
  }

  function getSelectedReportSource() {
    return state.sources.find(source => source.id === els.reportSourceSelect.value) || state.sources[0] || null;
  }

  function renderReportControls() {
    const previousSource = els.reportSourceSelect.value;
    els.reportSourceSelect.innerHTML = state.sources.length
      ? state.sources.map(source => `<option value="${escapeAttr(source.id)}">${escapeHtml(source.name)}</option>`).join('')
      : '<option value="">No source loaded</option>';

    if (state.sources.some(source => source.id === previousSource)) {
      els.reportSourceSelect.value = previousSource;
    } else if (state.sources.length) {
      els.reportSourceSelect.value = state.sources[0].id;
    }

    const source = getSelectedReportSource();
    const disabled = !source;
    [
      els.reportNameInput, els.reportDataSheetSelect,
      els.primaryBreakdownSelect, els.reportFilterColumnSelect,
      els.generateReportBtn, els.downloadReportCsvBtn, els.downloadReportXlsxBtn
    ].forEach(control => {
      control.disabled = disabled;
    });

    if (!source) {
      els.reportDataSheetSelect.innerHTML = '<option value="">No sheets</option>';
      els.primaryBreakdownSelect.innerHTML = '<option value="">No breakdown</option>';
      els.reportFilterColumnSelect.innerHTML = '<option value="">No filter</option>';
      renderCheckboxList(els.questionChecklist, [], { emptyText: 'No response columns found' });
      renderCheckboxList(els.reportFilterValues, [], { emptyText: 'No filter values found' });
      els.reportFilterValues.classList.add('hidden');
      els.reportFilterNote.textContent = '';
      updateReportColumnNote([]);
      return;
    }

    const sheetNames = source.workbook.SheetNames;
    populateSelect(els.reportDataSheetSelect, sheetNames, pickDataSheet(sheetNames), false);
    if (!normalizeValue(els.reportNameInput.value)) els.reportNameInput.value = 'Distribution';
    renderReportColumns();
  }

  function renderReportColumns() {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const columns = source && sheetName ? getSheetColumns(source.workbook, sheetName) : [];
    const dataRows = source && sheetName ? getSheetRecords(source.workbook, sheetName) : [];
    const columnStats = columns.map(column => ({
      column,
      uniqueCount: getReportUniqueValues(dataRows, column).length
    }));
    const filterColumnStats = columns.map(column => ({
      column,
      uniqueCount: getReportFilterValues(dataRows, column).length
    }));
    const eligibleColumns = columnStats
      .filter(item => item.uniqueCount <= REPORT_UNIQUE_VALUE_LIMIT)
      .map(item => item.column);
    const eligibleFilterColumns = filterColumnStats
      .filter(item => item.uniqueCount > 0 && item.uniqueCount < REPORT_FILTER_UNIQUE_VALUE_LIMIT)
      .map(item => item.column);
    const ignoredColumns = columnStats.filter(item => item.uniqueCount > REPORT_UNIQUE_VALUE_LIMIT);
    const responseColumns = columns.filter(column => !isLikelyMetadataColumn(column));
    const defaultResponseColumns = responseColumns.filter(column => eligibleColumns.includes(column));
    renderCheckboxList(els.questionChecklist, eligibleColumns.map(column => ({ value: column, label: column })), {
      checkedValues: defaultResponseColumns.length ? defaultResponseColumns : eligibleColumns,
      emptyText: 'No response columns found'
    });
    populateSelect(els.primaryBreakdownSelect, eligibleColumns, '', true, 'No main breakdown');
    els.primaryBreakdownSelect.value = '';
    populateSelect(els.reportFilterColumnSelect, eligibleFilterColumns, '', true, 'No filter');
    els.reportFilterColumnSelect.value = '';
    renderReportFilterValues();
    updateReportColumnNote(ignoredColumns);
  }

  function renderReportFilterValues() {
    const source = getSelectedReportSource();
    const sheetName = els.reportDataSheetSelect.value;
    const filterColumn = els.reportFilterColumnSelect.value;
    const dataRows = source && sheetName ? getSheetRecords(source.workbook, sheetName) : [];

    if (!filterColumn || !dataRows.length) {
      renderCheckboxList(els.reportFilterValues, [], { emptyText: 'No filter values found' });
      els.reportFilterValues.classList.add('hidden');
      els.reportFilterNote.textContent = `Filter columns must have fewer than ${REPORT_FILTER_UNIQUE_VALUE_LIMIT} unique values.`;
      return;
    }

    const values = getReportFilterValues(dataRows, filterColumn);
    renderCheckboxList(els.reportFilterValues, values.map(value => ({ value, label: value })), {
      checkedValues: values,
      emptyText: 'No filter values found',
      onChange: updateReportFilterSelectionNote
    });
    els.reportFilterValues.classList.toggle('hidden', !values.length);
    updateReportFilterSelectionNote();
  }

  function updateReportFilterSelectionNote() {
    const inputs = Array.from(els.reportFilterValues.querySelectorAll('input[type="checkbox"]'));
    if (!inputs.length) {
      els.reportFilterNote.textContent = 'No filter values found for this column.';
      return;
    }

    const selectedCount = inputs.filter(input => input.checked).length;
    els.reportFilterNote.textContent = `${selectedCount} of ${inputs.length} filter value${inputs.length === 1 ? '' : 's'} selected.`;
  }

  function syncBreakdownQuestionSelection() {
    const breakdownColumn = els.primaryBreakdownSelect.value;
    if (!breakdownColumn) return;
    Array.from(els.questionChecklist.querySelectorAll('input[type="checkbox"]'))
      .filter(input => input.value === breakdownColumn)
      .forEach(input => {
        input.checked = false;
      });
  }

  function updateReportColumnNote(ignoredColumns) {
    if (!els.reportColumnNote) return;
    if (!ignoredColumns.length) {
      els.reportColumnNote.textContent = `Only columns with ${REPORT_UNIQUE_VALUE_LIMIT} or fewer unique responses are shown here.`;
      return;
    }

    const names = ignoredColumns.slice(0, 4).map(item => item.column).join(', ');
    const extra = ignoredColumns.length > 4 ? `, and ${ignoredColumns.length - 4} more` : '';
    const reason = ignoredColumns.length === 1 ? 'it has' : 'they have';
    els.reportColumnNote.textContent = `${ignoredColumns.length} column${ignoredColumns.length === 1 ? '' : 's'} hidden because ${reason} more than ${REPORT_UNIQUE_VALUE_LIMIT} unique responses: ${names}${extra}.`;
  }

  function generateDistributionReport() {
    const source = getSelectedReportSource();
    if (!source) {
      showReportStatus('Load or upload a source first.', 'error');
      return;
    }

    try {
      const dataRows = getSheetRecords(source.workbook, els.reportDataSheetSelect.value);
      const reportName = normalizeValue(els.reportNameInput.value) || 'Distribution';
      if (!dataRows.length) throw new Error('The raw data sheet has no rows.');

      const reportFilter = getReportFilter(dataRows);
      const filteredRows = applyReportFilter(dataRows, reportFilter);
      const breakdownColumns = uniqueList([
        els.primaryBreakdownSelect.value
      ]).filter(column => column && column in (dataRows[0] || {}));
      const breakdownSet = new Set(breakdownColumns);
      const questions = getCheckedItems(els.questionChecklist).map(item => ({
        column: item.value,
        display: item.label
      })).filter(question => !breakdownSet.has(question.column));

      if (!filteredRows.length) throw new Error('No rows match the selected report filter.');
      if (!questions.length) throw new Error('Select at least one response column to include.');

      const output = buildDistributionOutput(filteredRows, questions, breakdownColumns, reportFilter.label);
      state.reportResult = {
        title: reportName,
        aoa: output.aoa,
        rows: output.rows,
        skipped: output.skipped,
        questionCount: output.questionCount,
        breakdownLabel: output.breakdownLabel,
        filterLabel: output.filterLabel,
        sourceName: source.name
      };
      renderDistributionOutput(state.reportResult);
      showReportStatus('Distribution report generated.', '');
    } catch (error) {
      console.error(error);
      showReportStatus(error.message || 'Could not generate the report.', 'error');
    }
  }

  function getReportFilter(dataRows) {
    const column = els.reportFilterColumnSelect.value;
    if (!column) return { column: '', values: new Set(), label: 'All rows' };
    const selectedValues = getCheckedItems(els.reportFilterValues).map(item => item.value);
    const totalValues = getReportFilterValues(dataRows, column).length;
    return {
      column,
      values: new Set(selectedValues.map(normalizeForMatch)),
      label: formatReportFilterLabel(column, selectedValues, totalValues)
    };
  }

  function formatReportFilterLabel(column, selectedValues, totalValues) {
    if (!selectedValues.length) return `${column}: no values selected`;
    if (selectedValues.length === totalValues) return `${column}: all ${totalValues} value${totalValues === 1 ? '' : 's'}`;
    const preview = selectedValues.slice(0, 5).join(', ');
    const extra = selectedValues.length > 5 ? `, and ${selectedValues.length - 5} more` : '';
    return `${column}: ${preview}${extra}`;
  }

  function applyReportFilter(dataRows, filter) {
    if (!filter.column) return dataRows;
    if (!filter.values.size) return [];
    return dataRows.filter(row => filter.values.has(normalizeForMatch(getResponseLabel(row[filter.column]))));
  }

  function buildDistributionOutput(dataRows, questions, breakdownColumns, filterLabel = 'All rows') {
    const aoa = [];
    const rows = [];
    const skipped = [];
    let questionCount = 0;
    const breakdownLabel = breakdownColumns.length ? breakdownColumns[0] : 'No main breakdown selected';
    const contextWidth = getReportSectionWidth(dataRows, breakdownColumns);
    addReportContextRow(aoa, rows, contextWidth, 'Breakdown', breakdownLabel);
    addReportContextRow(aoa, rows, contextWidth, 'Filter', filterLabel);
    const contextSpacer = Array(contextWidth).fill('');
    aoa.push(contextSpacer);
    rows.push(contextSpacer.map(value => ({ value, type: 'spacer' })));

    questions.forEach(question => {
      if (!(question.column in (dataRows[0] || {}))) {
        skipped.push(question.column);
        return;
      }
      const section = buildQuestionSection(dataRows, question.display || question.column, question.column, breakdownColumns);
      aoa.push(...section.aoa);
      rows.push(...section.rows);
      questionCount += 1;
    });

    return { aoa, rows, skipped, questionCount, breakdownLabel, filterLabel };
  }

  function addReportContextRow(aoa, rows, width, label, value) {
    const contextRow = Array(width).fill('');
    contextRow[0] = label;
    contextRow[1] = value;
    aoa.push(contextRow);
    rows.push([
      { value: label, type: 'meta-label' },
      { value, type: 'meta-value', colspan: Math.max(1, width - 1) }
    ]);
  }

  function buildQuestionSection(dataRows, displayName, questionColumn, breakdownColumns) {
    const answers = sortReportAnswers(getReportUniqueValues(dataRows, questionColumn), questionColumn);
    const combos = createBreakdownCombos(dataRows, breakdownColumns);
    const headerRows = createReportHeaderRows(combos, breakdownColumns);
    const width = Math.max(1, headerRows[0].length) + 1;
    const aoa = [];
    const rows = [];

    const titleRow = Array(width).fill('');
    titleRow[0] = displayName;
    aoa.push(titleRow);
    rows.push(titleRow.map((value, index) => ({ value, type: index === 0 ? 'question' : 'blank' })));

    headerRows.forEach(row => {
      const fullRow = ['', ...row];
      aoa.push(fullRow);
      rows.push(fullRow.map(value => ({ value, type: 'header' })));
    });

    answers.forEach(answer => {
      const row = Array(width).fill('');
      const renderedRow = Array(width).fill(null).map(() => ({ value: '', type: 'plain' }));
      row[0] = answer;
      renderedRow[0] = { value: answer, type: 'answer' };

      combos.forEach((combo, comboIndex) => {
        const filtered = dataRows.filter(dataRow => combo.conditions.every(condition => getReportValue(dataRow[condition.column]) === condition.value));
        const counts = countAnswers(filtered, questionColumn);
        const total = Array.from(counts.values()).reduce((sum, count) => sum + count, 0);
        const count = counts.get(normalizeForMatch(answer)) || 0;
        const percent = total ? count / total : 0;
        const countIndex = 1 + (comboIndex * 2);
        row[countIndex] = count;
        row[countIndex + 1] = percent;
        renderedRow[countIndex] = { value: count, type: 'count' };
        renderedRow[countIndex + 1] = { value: percent, type: 'percent' };
      });

      aoa.push(row);
      rows.push(renderedRow);
    });

    const spacer = Array(width).fill('');
    aoa.push(spacer);
    rows.push(spacer.map(value => ({ value, type: 'spacer' })));
    return { aoa, rows };
  }

  function getReportSectionWidth(dataRows, breakdownColumns) {
    const combos = createBreakdownCombos(dataRows, breakdownColumns);
    const headerRows = createReportHeaderRows(combos, breakdownColumns);
    return Math.max(3, headerRows[0].length + 1);
  }

  function createReportHeaderRows(combos, breakdownColumns) {
    if (!breakdownColumns.length) return [['Count', 'Percentage']];

    const rows = breakdownColumns.map(column => {
      const row = [];
      combos.forEach(combo => {
        const value = combo.conditions.find(condition => condition.column === column)?.value || '';
        row.push(value, '');
      });
      return row;
    });
    rows.push(combos.flatMap(() => ['Count', 'Percentage']));
    return rows;
  }

  function createBreakdownCombos(dataRows, breakdownColumns) {
    if (!breakdownColumns.length) return [{ label: 'All rows', conditions: [] }];
    const valuesByColumn = breakdownColumns.map(column => ({
      column,
      values: getReportUniqueValues(dataRows, column)
    })).filter(item => item.values.length);

    if (!valuesByColumn.length) return [{ label: 'All rows', conditions: [] }];

    return cartesianProduct(valuesByColumn.map(item => item.values)).map(values => ({
      label: values.join(' / '),
      conditions: values.map((value, index) => ({ column: valuesByColumn[index].column, value }))
    }));
  }

  function countAnswers(rows, questionColumn) {
    const counts = new Map();
    rows.forEach(row => {
      const value = getReportValue(row[questionColumn]);
      if (!value) return;
      const key = normalizeForMatch(value);
      counts.set(key, (counts.get(key) || 0) + 1);
    });
    return counts;
  }

  function renderDistributionOutput(result) {
    els.reportOutputTitle.textContent = result.title;
    const skippedText = result.skipped.length ? `, skipped ${result.skipped.length} missing columns` : '';
    els.reportOutputMeta.textContent = `${result.questionCount} response columns from ${result.sourceName}; breakdown: ${result.breakdownLabel}; filter: ${result.filterLabel}${skippedText}`;
    els.distributionOutput.innerHTML = `
      <table>
        <tbody>
          ${result.rows.map(row => {
            const spacer = row.every(cell => cell.type === 'spacer' || normalizeValue(cell.value) === '');
            return `<tr class="${spacer ? 'report-spacer' : ''}">
              ${row.map(cell => renderReportCell(cell)).join('')}
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    `;
  }

  function renderReportCell(cell) {
    const colspan = cell.colspan ? ` colspan="${cell.colspan}"` : '';
    if (cell.type === 'spacer') return '<td></td>';
    if (cell.type === 'meta-label') return `<td${colspan} class="report-meta-label">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'meta-value') return `<td${colspan} class="report-meta-value">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'question') return `<td${colspan} class="question-title">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'header') return `<td${colspan} class="report-header">${escapeHtml(cell.value)}</td>`;
    if (cell.type === 'count') return `<td${colspan} class="number">${formatNumber(cell.value)}</td>`;
    if (cell.type === 'percent') {
      const percent = Number(cell.value) || 0;
      return `<td${colspan} class="number heat-cell" style="background:${getHeatColor(percent)}">${formatPercent(percent)}</td>`;
    }
    return `<td${colspan}>${escapeHtml(cell.value)}</td>`;
  }

  function downloadDistributionCsv() {
    if (!state.reportResult) {
      alert('Generate a distribution report first.');
      return;
    }
    downloadCsv(state.reportResult.aoa.map(row => row.map(value => typeof value === 'number' ? value : displayCell(value))), `${safeFileName(state.reportResult.title)}.csv`);
  }

  function downloadDistributionXlsx() {
    if (!state.reportResult) {
      alert('Generate a distribution report first.');
      return;
    }
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet(state.reportResult.aoa);
    XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(state.reportResult.title));
    XLSX.writeFile(workbook, `${safeFileName(state.reportResult.title)}.xlsx`);
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

  function populateSelect(select, values, selectedValue, includeNone, noneLabel = 'None') {
    const previous = select.value;
    const options = includeNone ? [`<option value="">${escapeHtml(noneLabel)}</option>`] : [];
    options.push(...values.map(value => `<option value="${escapeAttr(value)}">${escapeHtml(value)}</option>`));
    select.innerHTML = options.join('');

    if (selectedValue && values.includes(selectedValue)) select.value = selectedValue;
    else if (previous && (values.includes(previous) || (includeNone && previous === ''))) select.value = previous;
    else if (includeNone) select.value = '';
    else if (values.length) select.value = values[0];
  }

  function pickSheet(sheetNames, pattern) {
    return sheetNames.find(name => pattern.test(name)) || sheetNames[0] || '';
  }

  function pickDataSheet(sheetNames) {
    return sheetNames.find(name => !/question|config|input|generation|site|scs|lookup|mapping/i.test(name)) || sheetNames[0] || '';
  }

  function pickColumn(columns, pattern) {
    return columns.find(column => pattern.test(column)) || columns[0] || '';
  }

  function isLikelyMetadataColumn(column) {
    return /^(id|student id|response id|tempid)$/i.test(column)
      || /date|timestamp|email|name|phone|address/i.test(column);
  }

  function getSheetColumns(workbook, sheetName) {
    const rows = getSheetMatrix(workbook, sheetName);
    const headerRow = findHeaderRow(rows);
    return rows[headerRow] ? makeUniqueHeaders(rows[headerRow]) : [];
  }

  function getSheetRecords(workbook, sheetName) {
    const rows = getSheetMatrix(workbook, sheetName);
    const headerRow = findHeaderRow(rows);
    const headers = rows[headerRow] ? makeUniqueHeaders(rows[headerRow]) : [];
    return rows.slice(headerRow + 1)
      .filter(row => row.some(cell => normalizeValue(cell) !== ''))
      .map(row => {
        const record = {};
        headers.forEach((header, index) => {
          record[header] = row[index] === undefined ? '' : row[index];
        });
        return record;
      });
  }

  function getSheetMatrix(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false
    }).filter(row => row.some(cell => normalizeValue(cell) !== ''));
  }

  function findHeaderRow(rows) {
    if (!rows.length) return 0;
    const surveyHeaderIndex = rows.slice(0, 6).findIndex(row => row.some(cell => /^survey$/i.test(normalizeValue(cell))));
    if (surveyHeaderIndex >= 0) return surveyHeaderIndex;
    return 0;
  }

  function renderCheckboxList(container, items, options = {}) {
    const normalizedItems = items.map(item => typeof item === 'string' ? { value: item, label: item } : item);
    const checkedValues = options.checkedValues === undefined
      ? normalizedItems.map(item => item.value)
      : options.checkedValues;
    const checked = new Set(checkedValues.map(normalizeValue));

    if (!normalizedItems.length) {
      container.innerHTML = `<div class="checklist-empty">${escapeHtml(options.emptyText || 'No options found')}</div>`;
      return;
    }

    container.innerHTML = normalizedItems.map((item, index) => {
      const id = `${container.id}-${hashString(item.value)}-${index}`;
      const isChecked = checked.has(normalizeValue(item.value));
      return `
        <label for="${id}">
          <input id="${id}" type="checkbox" value="${escapeAttr(item.value)}" data-label="${escapeAttr(item.label || item.value)}" ${isChecked ? 'checked' : ''}>
          <span>${escapeHtml(item.label || item.value)}</span>
        </label>
      `;
    }).join('');

    if (options.onChange) {
      container.querySelectorAll('input[type="checkbox"]').forEach(input => {
        input.addEventListener('change', options.onChange);
      });
    }
  }

  function getCheckedItems(container) {
    return Array.from(container.querySelectorAll('input[type="checkbox"]:checked'))
      .map(input => ({
        value: input.value,
        label: input.dataset.label || input.value
      }))
      .filter(item => item.value);
  }

  function getReportUniqueValues(rows, column) {
    const seen = new Set();
    const values = [];
    rows.forEach(row => {
      const value = getReportValue(row[column]);
      if (!value || seen.has(value)) return;
      seen.add(value);
      values.push(value);
    });
    return values;
  }

  function getReportFilterValues(rows, column) {
    return Array.from(new Set(rows.map(row => getResponseLabel(row[column]))))
      .sort((a, b) => a.localeCompare(b));
  }

  function getReportValue(value) {
    const normalized = normalizeValue(value);
    return normalized === NO_RESPONSE ? '' : normalized;
  }

  function sortReportAnswers(values, questionText) {
    const observed = values.filter(Boolean);
    const normalizedObserved = new Set(observed.map(normalizeForMatch));
    const hardSet = getHardCodedAnswerSet(questionText);
    if (hardSet) return hardSet;

    const matchingSet = ANSWER_SETS.find(set => {
      const normalizedSet = new Set(set.map(normalizeForMatch));
      return Array.from(normalizedObserved).every(value => normalizedSet.has(value));
    });

    if (matchingSet) return matchingSet;
    return [...observed].sort((a, b) => a.localeCompare(b));
  }

  function getHardCodedAnswerSet(questionText) {
    if (/How frequently did you engage with your students.*families this summer/i.test(questionText)) {
      return ['Never', 'Rarely', 'Occasionally', 'A moderate amount', 'A great deal'];
    }
    if (/To what extent did your child\/children enjoy participating/i.test(questionText)) {
      return ['Not at all', 'A little bit', 'Somewhat well', 'Quite well', 'Extremely well'];
    }
    if (/If given the opportunity, would you return/i.test(questionText)) {
      return ['Yes', 'Maybe', 'No'];
    }
    return null;
  }

  function cartesianProduct(arrays) {
    return arrays.reduce((acc, values) => acc.flatMap(prefix => values.map(value => [...prefix, value])), [[]]);
  }

  function uniqueList(values) {
    return Array.from(new Set(values));
  }

  function normalizeForMatch(value) {
    return normalizeValue(value)
      .toLowerCase()
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getHeatColor(percent) {
    const value = Math.max(0, Math.min(1, Number(percent) || 0));
    const lightness = 97 - (value * 35);
    const saturation = 58 + (value * 16);
    return `hsl(151, ${saturation}%, ${lightness}%)`;
  }

  function formatPercent(value) {
    return `${roundOne((Number(value) || 0) * 100)}%`;
  }

  function extractGoogleSheetId(url) {
    const match = normalizeValue(url).match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

  function safeSheetName(value) {
    return (normalizeValue(value).replace(/[\\/?*:[\]]+/g, ' ').trim() || 'Distribution').slice(0, 31);
  }

  function showReportStatus(message, type) {
    els.reportStatus.textContent = message;
    els.reportStatus.className = `status-message ${type || ''}`.trim();
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
    state.sources = state.sources.filter(source => source.id !== UPLOADED_SOURCE_ID);
    state.reportResult = null;
    els.chartGrid.innerHTML = '';
    els.fileStats.innerHTML = '';
    els.reportOutputTitle.textContent = 'No report generated yet';
    els.reportOutputMeta.textContent = 'Select a source and generate a report.';
    els.distributionOutput.innerHTML = '<div class="report-empty">Upload a workbook or load a public Google Sheet, then choose the input settings.</div>';
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
