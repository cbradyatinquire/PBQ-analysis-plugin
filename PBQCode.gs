// Display labels for all 15 non-empty answer combinations (index = bitmask value).
// Bit 0 = option A, bit 1 = option B, bit 2 = option C, bit 3 = option D.
const COMBO_LABELS = [
  ' ', 'A', 'B', 'AB', 'C', 'AC', 'BC', 'ABC',
  'D', 'AD', 'BD', 'ABD', 'CD', 'ACD', 'BCD', 'ABCD'
];

const DATA_SHEET_NAME = '_PBQ_Data_';

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
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
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
    if (!header || header === DATA_SHEET_NAME) continue;

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
        // A multi-select response either contains ", " (multiple options chosen) or is
        // a single option text longer than a typical name/number (>10 chars).
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
// Returns { options: string[4], warning: string|null, source: 'form'|'responses'|'none' }.
//
// Strategy:
//  1. If the spreadsheet is linked to a Google Form, read options directly from the
//     form definition — exact, authoritative, no parsing needed.
//  2. If unlinked (e.g. a saved copy), fall back to inferring options by splitting
//     response text on the Forms checkbox separator (", ").
function discoverOptions(colIndex) {
  const sheet  = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  const empty = { options: ['', '', '', ''], warning: null, source: 'none' };
  if (lastRow < 2) return Object.assign(empty, { warning: 'No responses found.' });

  const columnHeader = String(sheet.getRange(1, colIndex, 1, 1).getValue()).trim();

  // ── Path 1: read directly from the linked form ───────────────────────────
  const formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
  if (formUrl) {
    try {
      const form          = FormApp.openByUrl(formUrl);
      const checkboxItems = form.getItems(FormApp.ItemType.CHECKBOX);

      // Match by exact title first, then by the column header being a prefix
      // (Forms truncates long titles in some UI contexts but Sheets stores the full text).
      let match = checkboxItems.find(item => item.getTitle().trim() === columnHeader);
      if (!match) {
        match = checkboxItems.find(item => columnHeader.startsWith(item.getTitle().trim()));
      }

      if (match) {
        const choices = match.asCheckboxItem().getChoices().map(c => c.getValue());
        const options = choices.slice(0, 4);
        while (options.length < 4) options.push('');

        let warning = null;
        if (choices.length > 4) {
          warning = `Form has ${choices.length} options — only 4 are supported. Edit below if needed.`;
        } else if (choices.length < 2) {
          warning = 'Fewer than 2 options found in the form — check the question type.';
        }

        return { options, warning, source: 'form' };
      }
      // Linked form found but no checkbox question matched this column header —
      // fall through to response-based detection.
    } catch (_) {
      // FormApp access failed (permissions, deleted form, etc.) — fall through.
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

  const minCount  = Math.max(2, Math.floor(responses.length * 0.05));
  const candidates = [...tokenCounts.entries()]
    .filter(([, count]) => count >= minCount)
    .sort((a, b) => tokenFirstIdx.get(a[0]) - tokenFirstIdx.get(b[0]))
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

  return { options, warning, source: 'responses' };
}

// Main chart-generation entry point.
// options       – array of 4 option texts (may be empty strings if fewer than 4 options)
// correctIdxs   – 0-based indices of correct options, e.g. [0, 2] means A and C
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

  const colData      = sheet.getRange(1, colIndex, lastRow, 1).getValues();
  const questionTitle = String(colData[0][0]).trim();

  // Correct-answer bitmask
  const correctBinary = correctIdxs.reduce((acc, i) => acc | (1 << i), 0);

  // Encode each response as a bitmask by substring-matching against option texts.
  // This works regardless of whether options were prefixed "A."/"B." in the form.
  const binaryResponses = [];
  let unmatched = 0;

  for (let i = 1; i < colData.length; i++) {
    const resp = String(colData[i][0]).trim();
    if (!resp) continue;

    let bitmask = 0;
    options.forEach((opt, j) => {
      if (opt.trim() && resp.includes(opt.trim())) bitmask |= (1 << j);
    });

    if (bitmask > 0) {
      binaryResponses.push(bitmask);
    } else {
      unmatched++;
    }
  }

  if (binaryResponses.length === 0) {
    throw new Error(
      'No responses matched the provided option texts. ' +
      'Check that the option text in the sidebar exactly matches what is in the form.'
    );
  }

  // Build frequency table for all 15 non-empty combinations, sorted most-wrong first.
  const comboData = [];
  for (let j = 1; j <= 15; j++) {
    const delta = getDeltaCorrect(j, correctBinary);
    const freq  = binaryResponses.filter(b => b === j).length;
    comboData.push({ label: COMBO_LABELS[j], delta, freq });
  }
  comboData.sort((a, b) => b.delta - a.delta || a.label.localeCompare(b.label));

  // Write chart data to hidden helper sheet
  const dataSheet = getOrCreateDataSheet();
  dataSheet.clearContents();
  const tableData = comboData.map(row => [row.label, row.freq]);
  dataSheet.getRange(1, 1, tableData.length, 2).setValues(tableData);

  // Remove any previously generated PBQ charts on this sheet
  removePreviousPBQCharts(sheet);

  // Correct-answer label for chart subtitle, e.g. "A, C"
  const correctLabel = correctIdxs.sort().map(i => 'ABCD'[i]).join(', ');

  const chartTitle = questionTitle.length > 55
    ? questionTitle.slice(0, 52) + '…'
    : questionTitle;

  const chartsBefore = sheet.getCharts().map(c => c.getChartId());

  const builtChart = sheet.newChart()
    .addRange(dataSheet.getRange(1, 1, tableData.length, 2))
    .setChartType(Charts.ChartType.COLUMN)
    .asColumnChart()
    .setColors(['#1a73e8'])
    .setOption('title', chartTitle)
    .setOption('subtitle', `Correct: ${correctLabel}  ·  ${binaryResponses.length} responses${unmatched > 0 ? `  ·  ${unmatched} unmatched` : ''}`)
    .setOption('hAxis', { title: 'Answer combination', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { title: 'Responses', minValue: 0, format: '0' })
    .setOption('legend', { position: 'none' })
    .setOption('bar',   { groupWidth: '75%' })
    .setOption('chartArea', { left: 55, top: 55, width: '82%', height: '65%' })
    .setPosition(3, Math.min(colIndex + 1, sheet.getLastColumn() + 1), 10, 10)
    .build();

  sheet.insertChart(builtChart);

  const newChart = sheet.getCharts().find(c => !chartsBefore.includes(c.getChartId()));
  if (newChart) storePBQChartId(sheet.getName(), newChart.getChartId());

  const correctCount = binaryResponses.filter(b => b === correctBinary).length;
  const pctCorrect   = Math.round((correctCount / binaryResponses.length) * 100);

  return {
    questionTitle,
    totalResponses: binaryResponses.length,
    unmatchedCount: unmatched,
    correctCount,
    pctCorrect,
    correctLabel
  };
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

function getOrCreateDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(DATA_SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(DATA_SHEET_NAME);
    s.hideSheet();
  }
  return s;
}

// ─── Chart-ID tracking ────────────────────────────────────────────────────────

function storePBQChartId(sheetName, chartId) {
  const props = PropertiesService.getDocumentProperties();
  const store = JSON.parse(props.getProperty('pbq_chart_ids') || '{}');
  if (!store[sheetName]) store[sheetName] = [];
  store[sheetName].push(chartId);
  props.setProperty('pbq_chart_ids', JSON.stringify(store));
}

function removePreviousPBQCharts(sheet) {
  const props = PropertiesService.getDocumentProperties();
  const store = JSON.parse(props.getProperty('pbq_chart_ids') || '{}');
  const ids   = store[sheet.getName()] || [];

  sheet.getCharts()
    .filter(c => ids.includes(c.getChartId()))
    .forEach(c => sheet.removeChart(c));

  store[sheet.getName()] = [];
  props.setProperty('pbq_chart_ids', JSON.stringify(store));
}

function clearPBQCharts() {
  removePreviousPBQCharts(SpreadsheetApp.getActiveSheet());
}
