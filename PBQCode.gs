// Display labels for all 15 non-empty answer combinations (index = bitmask value).
// Bit 0 = option A, bit 1 = option B, bit 2 = option C, bit 3 = option D.
const COMBO_LABELS = [
  ' ', 'A', 'B', 'AB', 'C', 'AC', 'BC', 'ABC',
  'D', 'AD', 'BD', 'ABD', 'CD', 'ACD', 'BCD', 'ABCD'
];

// ─── Menu ────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PBQ Analysis')
    .addItem('Open PBQ Panel', 'showSidebar')
    .addSeparator()
    .addItem('Remove PBQ Charts from This Sheet', 'clearPBQCharts')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('PBQ Analysis')
    .setWidth(300);  // note: setWidth is ignored for sidebars; only works for dialogs
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── Per-tab form URL storage ─────────────────────────────────────────────────
// Each sheet tab can be linked to a specific Google Form by storing its URL
// in DocumentProperties keyed by the tab's permanent gid (sheetId).
// This overrides the spreadsheet-level getFormUrl(), which only ever returns
// one form even when multiple forms route to different tabs.

function getTabFormInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const key   = 'tab_form_' + sheet.getSheetId();
  const url   = PropertiesService.getDocumentProperties().getProperty(key) || '';
  if (!url) return { url: '', title: '' };
  try {
    const form = FormApp.openByUrl(url);
    return { url, title: form.getTitle() };
  } catch (_) {
    return { url, title: '(could not open form)' };
  }
}

function setTabFormUrl(url) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const key   = 'tab_form_' + sheet.getSheetId();
  const trimmed = (url || '').trim();
  if (trimmed) {
    // Validate before saving
    try { FormApp.openByUrl(trimmed); } catch (e) {
      throw new Error('Could not open that form URL. Check the URL and try again.\n(' + e.message + ')');
    }
    const form = FormApp.openByUrl(trimmed);
    PropertiesService.getDocumentProperties().setProperty(key, trimmed);
    return form.getTitle();
  } else {
    PropertiesService.getDocumentProperties().deleteProperty(key);
    return '';
  }
}

// ─── Called from sidebar ──────────────────────────────────────────────────────

// Returns column metadata for the active sheet so the sidebar can build its dropdown.
function getQuestionColumns() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  if (lastCol === 0 || lastRow === 0) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const columns = [];

  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).trim();
    if (!header) continue;

    const responseCount = lastRow > 1
      ? sheet.getRange(2, i + 1, lastRow - 1, 1)
          .getValues()
          .filter(row => String(row[0]).trim() !== '')
          .length
      : 0;

    // Detect PBQ-style checkbox responses: cells that contain multiple comma-separated
    // phrases (multi-select). Check that >40% of sampled rows have at least one ", " —
    // that's the separator Google Forms uses between selected checkbox options.
    let looksLikePBQ = false;
    if (lastRow > 1 && responseCount > 0) {
      const sampleSize = Math.min(lastRow - 1, 20);
      const sample = sheet.getRange(2, i + 1, sampleSize, 1).getValues();
      const nonEmpty = sample.filter(r => String(r[0]).trim() !== '');
      if (nonEmpty.length > 0) {
        const plausible = nonEmpty.filter(r => {
          const v = String(r[0]);
          return v.includes(', ') || v.length > 10;
        }).length;
        looksLikePBQ = plausible / nonEmpty.length > 0.4;
      }
    }

    columns.push({
      index: i + 1,
      colLetter: columnIndexToLetter(i + 1),
      header,
      responseCount,
      looksLikePBQ
    });
  }

  return columns;
}

// Returns up to 4 answer option texts for the given column.
// Returns { options: string[4], warning: string|null, correctIdxs: number[]|null,
//           source: 'form'|'responses'|'none' }.
//
// Strategy:
//  1. If the spreadsheet is linked to a Google Form, read options directly from the
//     form definition — exact, authoritative, order-preserving, includes unselected options.
//  2. If unlinked (e.g. a saved copy), fall back to inferring options by splitting
//     response text on the Forms checkbox separator (", ").
function discoverOptions(colIndex) {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const empty = { options: ['', '', '', ''], warning: null, correctIdxs: null, source: 'none' };
  if (lastRow < 2) return Object.assign(empty, { warning: 'No responses found.' });

  const columnHeader = String(sheet.getRange(1, colIndex, 1, 1).getValue()).trim();

  // ── Path 1: read directly from the linked form ───────────────────────────
  // Prefer the per-tab stored URL (set by the user in the sidebar) over the
  // spreadsheet-level getFormUrl(), which returns only one form even when
  // multiple forms route responses to different tabs.
  const tabFormUrl = PropertiesService.getDocumentProperties()
    .getProperty('tab_form_' + sheet.getSheetId());
  const formUrl = tabFormUrl || SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  if (formUrl) {
    try {
      const form          = FormApp.openByUrl(formUrl);
      const checkboxItems = form.getItems(FormApp.ItemType.CHECKBOX);

      let match = checkboxItems.find(item => item.getTitle().trim() === columnHeader);
      if (!match) {
        match = checkboxItems.find(item => columnHeader.startsWith(item.getTitle().trim()));
      }

      if (match) {
        const checkboxItem = match.asCheckboxItem();
        const allChoices   = checkboxItem.getChoices();
        const choices      = allChoices.slice(0, 4);
        const options      = choices.map(c => c.getValue());
        while (options.length < 4) options.push('');

        let correctIdxs = null;
        try {
          const correct = choices
            .map((c, i) => c.isCorrectAnswer() ? i : -1)
            .filter(i => i >= 0);
          if (correct.length > 0) correctIdxs = correct;
        } catch (_) {}

        let warning = null;
        if (allChoices.length > 4) {
          warning = `Form has ${allChoices.length} options — only first 4 shown. Edit below if needed.`;
        } else if (allChoices.length < 2) {
          warning = 'Fewer than 2 options found in the form — check the question type.';
        }

        return { options, warning, correctIdxs, source: 'form' };
      }
      // No matching checkbox question found in the linked form —
      // fall through to response-based detection below.
    } catch (_) {
      // FormApp access failed — fall through to response-based detection.
    }
  }

  // ── Path 2: infer options by splitting response text ─────────────────────
  const responses = sheet.getRange(2, colIndex, lastRow - 1, 1)
    .getValues()
    .map(row => String(row[0]).trim())
    .filter(r => r !== '');

  if (responses.length === 0) {
    return Object.assign(empty, { warning: 'No responses found in this column.' });
  }

  const tokenCounts   = new Map();
  const tokenFirstIdx = new Map();
  let seenOrder = 0;

  responses.forEach(resp => {
    resp.split(', ').map(t => t.trim()).filter(t => t.length > 0).forEach(token => {
      tokenCounts.set(token, (tokenCounts.get(token) || 0) + 1);
      if (!tokenFirstIdx.has(token)) tokenFirstIdx.set(token, seenOrder++);
    });
  });

  const minCount   = Math.max(2, Math.floor(responses.length * 0.05));
  const candidates = [...tokenCounts.entries()]
    .filter(([, count]) => count >= minCount)
    .sort((a, b) => a[0].localeCompare(b[0]))  // alphabetical matches A./B./C./D. prefix convention
    .map(([text]) => text);

  const options = candidates.slice(0, 4);
  while (options.length < 4) options.push('');

  let warning = null;
  if (candidates.length === 0) {
    warning = 'Could not detect options automatically. Enter them manually below.';
  } else if (candidates.length < 4) {
    warning = `Only ${candidates.length} option${candidates.length > 1 ? 's' : ''} detected — fill in any missing ones.`;
  } else if (candidates.length > 4) {
    warning = `${candidates.length} candidates found — showing first 4. Edit below if needed.`;
  }

  return { options, warning, correctIdxs: null, source: 'responses' };
}

// Main chart-generation entry point.
// options     – array of 4 option texts (empty string = unused slot)
// correctIdxs – 0-based indices of correct options, e.g. [0, 2] means A and C
function generatePBQChart(colIndex, options, correctIdxs) {
  const activeOptions = options.map(o => String(o).trim()).filter(o => o !== '');
  if (activeOptions.length === 0) {
    throw new Error('Please enter at least one answer option text.');
  }
  if (!correctIdxs || correctIdxs.length === 0) {
    throw new Error('Please select at least one correct answer.');
  }

  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No response data found in the selected column.');

  const colData       = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  const questionTitle = String(colData[0][0]).trim();
  const correctBinary = correctIdxs.reduce((acc, i) => acc | (1 << i), 0);

  // Encode each response as a bitmask via substring-matching against option texts.
  const binaryResponses = [];
  let unmatched = 0;

  for (let i = 1; i < colData.length; i++) {
    const resp = String(colData[i][0]).trim();
    if (!resp) continue;

    let bitmask = 0;
    options.forEach((opt, j) => {
      if (opt.trim() && resp.includes(opt.trim())) bitmask |= (1 << j);
    });

    if (bitmask > 0) { binaryResponses.push(bitmask); } else { unmatched++; }
  }

  if (binaryResponses.length === 0) {
    throw new Error(
      'No responses matched the provided option texts. ' +
      'Check that the option text in the sidebar exactly matches what is in the form.'
    );
  }

  // Build frequency table, sorted most-wrong first.
  const comboData = [];
  for (let j = 1; j <= 15; j++) {
    const delta = getDeltaCorrect(j, correctBinary);
    const freq  = binaryResponses.filter(b => b === j).length;
    comboData.push({ label: COMBO_LABELS[j], delta, freq });
  }
  comboData.sort((a, b) => b.delta - a.delta || a.label.localeCompare(b.label));

  // Remove previous PBQ chart and its data columns from this sheet.
  // Must happen before getLastColumn() so vacated columns aren't counted.
  removePreviousPBQCharts(sheet);

  // Write chart data to this sheet, well to the right of any existing content.
  // Start at ROW 2, not row 1 — Google Sheets treats row 1 of a chart range as a
  // header and skips it as a data point. Since comboData is sorted most-wrong-first,
  // the AC (or whatever is furthest from correct) always lands first and would be
  // silently dropped if written to row 1.
  const tableData  = comboData.map(row => [row.label, row.freq]);
  const dataCol    = sheet.getLastColumn() + 3;
  const dataStart  = 2;
  const labelRange = sheet.getRange(dataStart, dataCol,     tableData.length, 1);
  const valueRange = sheet.getRange(dataStart, dataCol + 1, tableData.length, 1);
  sheet.getRange(dataStart, dataCol, tableData.length, 2).setValues(tableData);
  sheet.getRange(dataStart, dataCol, tableData.length, 2).setFontColor('#aaaaaa').setFontSize(9);
  sheet.getRange(dataStart, dataCol, 1, 2).setNote('PBQ chart data — do not edit');

  // Flush ensures the values are committed to the spreadsheet before the chart
  // builder snapshots the range — without this the chart can render as blank.
  SpreadsheetApp.flush();

  const correctLabel = correctIdxs.sort().map(i => 'ABCD'[i]).join(', ');
  const chartTitle   = questionTitle.length > 55
    ? questionTitle.slice(0, 52) + '…'
    : questionTitle;

  const chartsBefore = sheet.getCharts().map(c => c.getChartId());

  const hAxisTitle = `Correct: ${correctLabel}  ·  ${binaryResponses.length} responses` +
                     (unmatched > 0 ? `  ·  ${unmatched} unmatched` : '');

  // Add label column first, then value column as two separate addRange calls.
  // GAS rule: if the first addRange is text-only, it becomes the X-axis domain.
  // Adding both columns as one 2-column range bypasses this detection and leaves
  // the chart blank because it can't determine which column is labels vs. series.
  const builtChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(labelRange)
    .addRange(valueRange)
    .setOption('title', chartTitle)
    .setOption('hAxis.title', hAxisTitle)
    .setOption('vAxis.title', 'Responses')
    .setOption('vAxis.minValue', 0)
    .setPosition(3, colIndex + 1, 10, 10)
    .build();

  sheet.insertChart(builtChart);

  // Flush after insertChart so the chart list reflects the new chart before we search.
  // Without this, getCharts() may not yet include the newly inserted chart.
  SpreadsheetApp.flush();
  const newChart = sheet.getCharts().find(c => !chartsBefore.includes(c.getChartId()));
  if (newChart) storePBQChartMeta(sheet.getName(), newChart.getChartId(), dataCol);

  const correctCount = binaryResponses.filter(b => b === correctBinary).length;
  const pctCorrect   = Math.round((correctCount / binaryResponses.length) * 100);

  return { questionTitle, totalResponses: binaryResponses.length, unmatchedCount: unmatched,
           correctCount, pctCorrect, correctLabel };
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function columnIndexToLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// Counts how many bits differ between two bitmask answers (Hamming distance).
function getDeltaCorrect(source, target) {
  const delta = source ^ target;
  return [1, 2, 4, 8].filter(bit => (delta & bit) !== 0).length;
}

// ─── Chart + data tracking ────────────────────────────────────────────────────
// Each entry stores both the chart ID and the starting column of its data,
// so removal cleans up both the chart and its data columns on that sheet tab.

function storePBQChartMeta(sheetName, chartId, dataCol) {
  const props = PropertiesService.getDocumentProperties();
  const store = JSON.parse(props.getProperty('pbq_chart_meta') || '{}');
  if (!store[sheetName]) store[sheetName] = [];
  store[sheetName].push({ chartId, dataCol });
  props.setProperty('pbq_chart_meta', JSON.stringify(store));
}

function removePreviousPBQCharts(sheet) {
  const props   = PropertiesService.getDocumentProperties();
  const store   = JSON.parse(props.getProperty('pbq_chart_meta') || '{}');
  const entries = store[sheet.getName()] || [];

  // Remove tracked charts by stored chart ID.
  const trackedIds = entries.map(e => e.chartId);
  sheet.getCharts()
    .filter(c => trackedIds.includes(c.getChartId()))
    .forEach(c => sheet.removeChart(c));

  // Clear tracked data columns. Data now starts at row 2 (15 rows = rows 2–16).
  entries.forEach(e => {
    try { sheet.getRange(2, e.dataCol, 15, 2).clear(); } catch (_) {}
  });

  // Fallback: scan every column for the PBQ data note and clear any found.
  // Check both row 1 and row 2 to handle data written by older versions of the script.
  const lastCol = sheet.getLastColumn();
  for (let c = 1; c <= lastCol; c++) {
    try {
      const n1 = sheet.getRange(1, c, 1, 1).getNote();
      const n2 = sheet.getRange(2, c, 1, 1).getNote();
      if (n1.includes('PBQ chart data') || n2.includes('PBQ chart data')) {
        sheet.getRange(1, c, 17, 2).clear();  // rows 1–17 covers both old and new layouts
        c++;
      }
    } catch (_) {}
  }

  SpreadsheetApp.flush();  // commit all clears before getLastColumn() is used for new dataCol

  store[sheet.getName()] = [];
  props.setProperty('pbq_chart_meta', JSON.stringify(store));
}

function clearPBQCharts() {
  removePreviousPBQCharts(SpreadsheetApp.getActiveSheet());
}

