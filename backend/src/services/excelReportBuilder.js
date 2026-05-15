const ExcelJS = require('exceljs');

// ─── Color Palette ───
const COLORS = {
  primary:     '0078D4',
  primaryDark: '005A9E',
  green:       '107C10',
  greenLight:  'DFF6DD',
  red:         'A4262C',
  redLight:    'FDE7E9',
  orange:      'CA5010',
  orangeLight: 'FFF4CE',
  purple:      '8764B8',
  purpleLight: 'E8E0F0',
  white:       'FFFFFF',
  lightGray:   'F3F2F1',
  darkGray:    '605E5C',
  medGray:     'D2D0CE',
  black:       '323130',
  altRow:      'F8F8FA',
};

const BORDER_THIN = { style: 'thin', color: { argb: 'FFD0D0D0' } };
const BORDERS_ALL = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };

/**
 * Build a rich Excel workbook (async – uses exceljs).
 * Writes directly to outputPath.
 */
async function buildReport({ assessed, summary, mode, sheetName, outputPath }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Azure CSP Migration Tool';
  wb.created = new Date();

  const statusKey =
    mode === 'jio'    ? 'JIO REGION AVAILABLE'
    : mode === 'region' ? 'REGION MOVE SUPPORTED'
    : 'SUBSCRIPTION MOVE SUPPORTED';

  const modeLabel =
    mode === 'jio'    ? 'Jio Availability'
    : mode === 'region' ? 'Region Move'
    : 'Subscription Move';

  _buildDataSheet(wb, assessed, statusKey, sheetName || 'Assessment Data');
  _buildDashboard(wb, assessed, summary, mode, modeLabel, statusKey);
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'PROVIDER');
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'RESOURCE GROUP');
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'LOCATION');
  _buildStatusSheet(wb, assessed, statusKey, modeLabel);
  _buildActionSheet(wb, assessed, statusKey, modeLabel);

  await wb.xlsx.writeFile(outputPath);
}

// ──────────────────────────────────────────────────────────
// Sheet 1 — Assessment Data
// ──────────────────────────────────────────────────────────

function _buildDataSheet(wb, assessed, statusKey, title) {
  if (assessed.length === 0) return;
  const ws = wb.addWorksheet(title, { views: [{ state: 'frozen', ySplit: 1 }] });
  const keys = Object.keys(assessed[0]);

  const headerRow = ws.addRow(keys);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primary } };
    cell.font = { bold: true, color: { argb: COLORS.white }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDERS_ALL;
  });
  headerRow.height = 28;

  const statusColIdx = keys.indexOf(statusKey) + 1;

  assessed.forEach((row, i) => {
    const dr = ws.addRow(keys.map(k => row[k] ?? ''));
    const isAlt = i % 2 === 1;
    dr.eachCell(cell => {
      cell.font = { size: 10, color: { argb: COLORS.black } };
      cell.border = BORDERS_ALL;
      cell.alignment = { vertical: 'middle', wrapText: true };
      if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.altRow } };
    });
    if (statusColIdx > 0) {
      const sc = dr.getCell(statusColIdx);
      const c = _statusColors(String(sc.value || ''));
      sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c.bg } };
      sc.font = { bold: true, size: 10, color: { argb: c.fg } };
    }
  });

  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: false }, cell => {
      const len = String(cell.value || '').length;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 4, 55);
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: assessed.length + 1, column: keys.length } };
}

// ──────────────────────────────────────────────────────────
// Sheet 2 — Summary Dashboard
// ──────────────────────────────────────────────────────────

function _buildDashboard(wb, assessed, summary, mode, modeLabel) {
  const ws = wb.addWorksheet('Summary Dashboard');
  const dateStr = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' });

  // We use columns A-L (12 cols); F-L used for charts on the right side
  const BAR_START = 6; // column F
  const BAR_COLS = 20; // F through Y (20 cells for 100%)

  // Set narrow widths for chart bar columns (F onward)
  ws.getColumn(1).width = 28;
  ws.getColumn(2).width = 14;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 3; // spacer
  for (let c = BAR_START; c < BAR_START + BAR_COLS + 2; c++) ws.getColumn(c).width = 3.2;

  const providers = new Set(), rgs = new Set(), locs = new Set();
  const providerCounts = {};
  assessed.forEach(r => {
    const t = (r['NORMALIZED TYPE'] || _findField(r, 'TYPE') || '').toLowerCase().split('/');
    if (t.length >= 2) {
      providers.add(t[0]);
      const pName = t[0].replace('microsoft.', '');
      providerCounts[pName] = (providerCounts[pName] || 0) + 1;
    }
    rgs.add(_findField(r, 'RESOURCE GROUP') || 'Unknown');
    locs.add(_findField(r, 'LOCATION') || 'Unknown');
  });

  const cond = summary.conditional || 0;
  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';

  // ── Title ──
  ws.mergeCells('A1:D1');
  const tc = ws.getCell('A1');
  tc.value = `AZURE CSP MIGRATION ASSESSMENT — ${modeLabel.toUpperCase()}`;
  tc.font = { bold: true, size: 16, color: { argb: COLORS.white } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primary } };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 40;

  // Right side: chart title
  const chartTitleEnd = BAR_START + BAR_COLS - 1;
  ws.mergeCells(1, BAR_START, 1, chartTitleEnd);
  const ctCell = ws.getCell(1, BAR_START);
  ctCell.value = 'DISTRIBUTION CHART';
  ctCell.font = { bold: true, size: 14, color: { argb: COLORS.white } };
  ctCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
  ctCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // ── Assessment Details ──
  let row = 3;
  _sectionHeader(ws, row, 'ASSESSMENT DETAILS', 'A', 'D'); row++;
  _labelValue(ws, row++, 'Assessment Type', modeLabel);
  _labelValue(ws, row++, 'Assessment Date', dateStr);
  _labelValue(ws, row++, 'Data Source', 'Microsoft Learn (Real-time)');
  _labelValue(ws, row++, 'Total Resources', String(summary.total));

  // ── Right side: Stacked Distribution Bar (rows 3-4) ──
  // Label row
  ws.mergeCells(3, BAR_START, 3, chartTitleEnd);
  const distLabel = ws.getCell(3, BAR_START);
  distLabel.value = 'MIGRATION READINESS DISTRIBUTION';
  distLabel.font = { bold: true, size: 11, color: { argb: COLORS.white } };
  distLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
  distLabel.alignment = { horizontal: 'center', vertical: 'middle' };

  // Stacked bar: calculate how many cells each status gets
  const statusItems = [
    { count: summary.yes, color: COLORS.green, label: 'Yes' },
    { count: cond, color: COLORS.purple, label: 'Cond' },
    { count: summary.review, color: COLORS.orange, label: 'Review' },
    { count: summary.no, color: COLORS.red, label: 'No' },
  ];
  if (mode !== 'subscription') statusItems.splice(1, 1); // remove conditional for non-subscription

  let colOffset = BAR_START;
  ws.getRow(4).height = 28;
  statusItems.forEach(item => {
    if (item.count === 0) return;
    const cellCount = Math.max(1, Math.round((item.count / summary.total) * BAR_COLS));
    const endCol = Math.min(colOffset + cellCount - 1, BAR_START + BAR_COLS - 1);
    for (let c = colOffset; c <= endCol; c++) {
      const cell = ws.getCell(4, c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: item.color } };
      cell.border = BORDERS_ALL;
    }
    // Put label in first cell of segment
    const labelCell = ws.getCell(4, colOffset);
    labelCell.value = item.count;
    labelCell.font = { bold: true, size: 10, color: { argb: COLORS.white } };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };
    colOffset = endCol + 1;
  });

  // Legend row
  ws.getRow(5).height = 20;
  const legendItems = [
    { label: `■ Yes (${pct(summary.yes)}%)`, color: COLORS.green },
    { label: `■ No (${pct(summary.no)}%)`, color: COLORS.red },
    { label: `■ Review (${pct(summary.review)}%)`, color: COLORS.orange },
  ];
  if (mode === 'subscription') {
    legendItems.push({ label: `■ Conditional (${pct(cond)}%)`, color: COLORS.purple });
  }
  const legendCellWidth = Math.floor(BAR_COLS / legendItems.length);
  legendItems.forEach((item, idx) => {
    const startCol = BAR_START + idx * legendCellWidth;
    const endCol = BAR_START + (idx + 1) * legendCellWidth - 1;
    ws.mergeCells(5, startCol, 5, endCol);
    const cell = ws.getCell(5, startCol);
    cell.value = item.label;
    cell.font = { bold: true, size: 9, color: { argb: item.color } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // ── Summary Table (left side) ──
  row += 1;
  _sectionHeader(ws, row, 'OVERALL SUMMARY', 'A', 'D'); row++;
  const summHdr = ws.getRow(row);
  ['Metric', 'Count', 'Percentage', 'Visual'].forEach((h, i) => {
    const c = summHdr.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 11, color: { argb: COLORS.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = BORDERS_ALL;
  });
  summHdr.height = 24;
  row++;

  const summaryRows = [
    ['Total Assessed', summary.total, '100%', '', COLORS.primary],
    [mode === 'jio' ? 'Available (Yes)' : 'Can Move (Yes)', summary.yes, `${pct(summary.yes)}%`, _bar(pct(summary.yes)), COLORS.green],
    [mode === 'jio' ? 'Not Available (No)' : 'Cannot Move (No)', summary.no, `${pct(summary.no)}%`, _bar(pct(summary.no)), COLORS.red],
    ['Needs Review', summary.review, `${pct(summary.review)}%`, _bar(pct(summary.review)), COLORS.orange],
  ];
  if (mode === 'subscription') {
    summaryRows.push(['Conditional', cond, `${pct(cond)}%`, _bar(pct(cond)), COLORS.purple]);
  }

  const summTableStart = row;
  summaryRows.forEach(([label, count, percent, bar, color]) => {
    const r = ws.getRow(row);
    r.getCell(1).value = label;
    r.getCell(1).font = { bold: true, size: 11, color: { argb: COLORS.black } };
    r.getCell(2).value = count;
    r.getCell(2).alignment = { horizontal: 'center' };
    r.getCell(2).font = { bold: true, size: 12, color: { argb: color } };
    r.getCell(3).value = percent;
    r.getCell(3).alignment = { horizontal: 'center' };
    r.getCell(4).value = bar;
    r.getCell(4).font = { size: 10 };
    [1,2,3,4].forEach(c => { r.getCell(c).border = BORDERS_ALL; });
    row++;
  });

  // ── Right side: Horizontal Bar Chart (aligned with summary rows) ──
  const barChartLabelRow = summTableStart - 1; // on the header row
  ws.mergeCells(barChartLabelRow, BAR_START, barChartLabelRow, chartTitleEnd);
  const barChartTitle = ws.getCell(barChartLabelRow, BAR_START);
  barChartTitle.value = 'HORIZONTAL BAR CHART';
  barChartTitle.font = { bold: true, size: 11, color: { argb: COLORS.white } };
  barChartTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
  barChartTitle.alignment = { horizontal: 'center', vertical: 'middle' };

  // Skip "Total" row, draw bars for Yes/No/Review/Conditional
  const barItems = [
    { label: mode === 'jio' ? 'Available' : 'Can Move', count: summary.yes, color: COLORS.green, lightColor: COLORS.greenLight },
    { label: mode === 'jio' ? 'Not Available' : 'Cannot Move', count: summary.no, color: COLORS.red, lightColor: COLORS.redLight },
    { label: 'Needs Review', count: summary.review, color: COLORS.orange, lightColor: COLORS.orangeLight },
  ];
  if (mode === 'subscription') {
    barItems.push({ label: 'Conditional', count: cond, color: COLORS.purple, lightColor: COLORS.purpleLight });
  }

  // Total row — leave blank on chart side
  const totalRow = summTableStart; // "Total Assessed" row
  ws.mergeCells(totalRow, BAR_START, totalRow, chartTitleEnd);
  const totalChartCell = ws.getCell(totalRow, BAR_START);
  totalChartCell.value = `Total: ${summary.total} resources`;
  totalChartCell.font = { bold: true, size: 10, color: { argb: COLORS.primary } };
  totalChartCell.alignment = { horizontal: 'center', vertical: 'middle' };
  totalChartCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  totalChartCell.border = BORDERS_ALL;

  barItems.forEach((item, idx) => {
    const barRow = summTableStart + 1 + idx;
    const maxBarCols = BAR_COLS;
    const filledCols = summary.total > 0 ? Math.max(item.count > 0 ? 1 : 0, Math.round((item.count / summary.total) * maxBarCols)) : 0;

    ws.getRow(barRow).height = 22;

    // Draw filled portion
    for (let c = 0; c < filledCols; c++) {
      const cell = ws.getCell(barRow, BAR_START + c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: item.color } };
      cell.border = BORDERS_ALL;
    }
    // Draw empty portion (light background)
    for (let c = filledCols; c < maxBarCols; c++) {
      const cell = ws.getCell(barRow, BAR_START + c);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: item.lightColor } };
      cell.border = { top: BORDER_THIN, bottom: BORDER_THIN, left: { style: 'hair', color: { argb: 'FFE0E0E0' } }, right: { style: 'hair', color: { argb: 'FFE0E0E0' } } };
    }
    // Value label at end
    const labelCol = BAR_START + maxBarCols;
    const cell = ws.getCell(barRow, labelCol);
    cell.value = `${item.count} (${pct(item.count)}%)`;
    cell.font = { bold: true, size: 9, color: { argb: item.color } };
    cell.alignment = { horizontal: 'left', vertical: 'middle' };
  });

  // ── Scope ──
  row += 1;
  _sectionHeader(ws, row, 'SCOPE', 'A', 'D'); row++;
  _labelValue(ws, row++, 'Unique Resource Providers', String(providers.size));
  _labelValue(ws, row++, 'Unique Resource Groups', String(rgs.size));
  _labelValue(ws, row++, 'Unique Locations', String(locs.size));

  // ── Readiness Score ──
  row += 1;
  _sectionHeader(ws, row, 'READINESS SCORE', 'A', 'D'); row++;
  const readiness = parseFloat(pct(summary.yes));
  const readinessColor = readiness >= 80 ? COLORS.green : readiness >= 50 ? COLORS.orange : COLORS.red;

  ws.mergeCells(`A${row}:D${row}`);
  const sc = ws.getCell(`A${row}`);
  sc.value = `Migration Readiness: ${pct(summary.yes)}%  —  ${summary.yes} of ${summary.total} resources can proceed`;
  sc.font = { bold: true, size: 14, color: { argb: readinessColor } };
  sc.alignment = { horizontal: 'center', vertical: 'middle' };
  sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  sc.border = BORDERS_ALL;
  ws.getRow(row).height = 36;
  row++;

  ws.mergeCells(`A${row}:D${row}`);
  const bc = ws.getCell(`A${row}`);
  bc.value = `Blockers: ${summary.no}  |  Review Required: ${summary.review + cond}  —  See "Action Items" sheet`;
  bc.font = { size: 11, color: { argb: COLORS.darkGray } };
  bc.alignment = { horizontal: 'center', vertical: 'middle' };
  bc.border = BORDERS_ALL;
  row++;

  // ── Readiness Gauge Bar (full width) ──
  row += 1;
  _sectionHeader(ws, row, 'READINESS GAUGE', 'A', 'D'); row++;

  // Right side header
  ws.mergeCells(row - 1, BAR_START, row - 1, chartTitleEnd);
  const gaugeHdr = ws.getCell(row - 1, BAR_START);
  gaugeHdr.value = '';
  gaugeHdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };

  // Draw gauge: green cells for readiness %, gray for remainder
  ws.getRow(row).height = 30;
  const gaugeTotal = BAR_COLS + 4; // use A through all bar cols
  const greenCells = Math.round((readiness / 100) * gaugeTotal);

  // Label on left
  ws.mergeCells(`A${row}:D${row}`);
  const gaugeLabel = ws.getCell(`A${row}`);
  gaugeLabel.value = readiness >= 80 ? '✅ HIGH READINESS' : readiness >= 50 ? '⚠️ MODERATE READINESS' : '❌ LOW READINESS';
  gaugeLabel.font = { bold: true, size: 14, color: { argb: readinessColor } };
  gaugeLabel.alignment = { horizontal: 'center', vertical: 'middle' };
  gaugeLabel.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
  gaugeLabel.border = BORDERS_ALL;

  for (let c = 0; c < BAR_COLS; c++) {
    const cell = ws.getCell(row, BAR_START + c);
    if (c < greenCells) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: readinessColor } };
    } else {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.medGray } };
    }
    cell.border = BORDERS_ALL;
  }
  // Percentage label
  const gaugePctCell = ws.getCell(row, BAR_START + BAR_COLS);
  gaugePctCell.value = `${pct(summary.yes)}%`;
  gaugePctCell.font = { bold: true, size: 12, color: { argb: readinessColor } };
  gaugePctCell.alignment = { horizontal: 'left', vertical: 'middle' };
  row++;

  // ── Top Resource Providers Chart (right side, below gauge) ──
  row += 1;
  _sectionHeader(ws, row, 'TOP RESOURCE PROVIDERS', 'A', 'D');
  ws.mergeCells(row, BAR_START, row, chartTitleEnd);
  const tpHdr = ws.getCell(row, BAR_START);
  tpHdr.value = 'PROVIDER DISTRIBUTION';
  tpHdr.font = { bold: true, size: 11, color: { argb: COLORS.white } };
  tpHdr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
  tpHdr.alignment = { horizontal: 'center', vertical: 'middle' };
  row++;

  // Sort providers by count, take top 8
  const sortedProviders = Object.entries(providerCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  const maxProviderCount = sortedProviders.length > 0 ? sortedProviders[0][1] : 1;
  const providerColors = [COLORS.primary, COLORS.green, COLORS.orange, COLORS.purple, COLORS.red, COLORS.primaryDark, '4B8BBE', '306998'];

  sortedProviders.forEach(([provider, count], idx) => {
    const r = ws.getRow(row);
    r.height = 20;

    // Left: Provider name and count
    ws.mergeCells(row, 1, row, 2);
    const nameCell = ws.getCell(row, 1);
    nameCell.value = provider;
    nameCell.font = { bold: true, size: 10, color: { argb: COLORS.black } };
    nameCell.alignment = { horizontal: 'left', vertical: 'middle' };
    nameCell.border = BORDERS_ALL;

    ws.mergeCells(row, 3, row, 4);
    const countCell = ws.getCell(row, 3);
    countCell.value = `${count} resources`;
    countCell.font = { size: 10, color: { argb: COLORS.darkGray } };
    countCell.alignment = { horizontal: 'right', vertical: 'middle' };
    countCell.border = BORDERS_ALL;

    // Right: Bar
    const barColor = providerColors[idx % providerColors.length];
    const filledCols = Math.max(1, Math.round((count / maxProviderCount) * BAR_COLS));
    for (let c = 0; c < BAR_COLS; c++) {
      const cell = ws.getCell(row, BAR_START + c);
      if (c < filledCols) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: barColor } };
      } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.lightGray } };
      }
      cell.border = { top: BORDER_THIN, bottom: BORDER_THIN, left: { style: 'hair', color: { argb: 'FFE0E0E0' } }, right: { style: 'hair', color: { argb: 'FFE0E0E0' } } };
    }
    // Count at end
    const endCell = ws.getCell(row, BAR_START + BAR_COLS);
    endCell.value = count;
    endCell.font = { bold: true, size: 9, color: { argb: barColor } };
    endCell.alignment = { horizontal: 'left', vertical: 'middle' };

    row++;
  });
}

// ──────────────────────────────────────────────────────────
// Sheets 3-5 — Pivot Sheets
// ──────────────────────────────────────────────────────────

function _buildPivotSheet(wb, assessed, statusKey, modeLabel, pivotBy) {
  const sheetNames = { PROVIDER: 'By Provider', 'RESOURCE GROUP': 'By Resource Group', LOCATION: 'By Location' };
  const ws = wb.addWorksheet(sheetNames[pivotBy], { views: [{ state: 'frozen', ySplit: 3 }] });

  const pivot = {};
  assessed.forEach(r => {
    let key;
    if (pivotBy === 'PROVIDER') {
      const parts = (r['NORMALIZED TYPE'] || _findField(r, 'TYPE') || '').toLowerCase().split('/');
      key = parts.length >= 2 ? parts[0].replace('microsoft.', '') : 'unknown';
    } else if (pivotBy === 'RESOURCE GROUP') {
      key = _findField(r, 'RESOURCE GROUP') || 'Unknown';
    } else {
      key = _findField(r, 'LOCATION') || 'Unknown';
    }
    const status = r[statusKey] || 'Review';
    if (!pivot[key]) pivot[key] = { total: 0, Yes: 0, No: 0, Conditional: 0, Review: 0 };
    pivot[key].total++;
    if (pivot[key][status] !== undefined) pivot[key][status]++;
    else pivot[key].Review++;
  });

  // Title
  ws.mergeCells('A1:G1');
  const tc = ws.getCell('A1');
  tc.value = `PIVOT BY ${pivotBy} — ${modeLabel.toUpperCase()}`;
  tc.font = { bold: true, size: 14, color: { argb: COLORS.white } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primary } };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 32;

  ws.mergeCells('A2:G2');
  const sub = ws.getCell('A2');
  sub.value = `${Object.keys(pivot).length} unique ${pivotBy.toLowerCase()}s`;
  sub.font = { size: 10, italic: true, color: { argb: COLORS.darkGray } };
  sub.alignment = { horizontal: 'center' };

  const colLabel = pivotBy === 'PROVIDER' ? 'Resource Provider' : pivotBy === 'RESOURCE GROUP' ? 'Resource Group' : 'Location';
  const headers = [colLabel, 'Total', 'Yes', 'No', 'Conditional', 'Review', 'Move Rate (%)'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 11, color: { argb: COLORS.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = BORDERS_ALL;
  });
  hdrRow.height = 26;

  const sorted = Object.entries(pivot).sort((a, b) => b[1].total - a[1].total);
  let rowNum = 4;
  const totals = { total: 0, Yes: 0, No: 0, Conditional: 0, Review: 0 };

  sorted.forEach(([key, counts], i) => {
    totals.total += counts.total; totals.Yes += counts.Yes; totals.No += counts.No;
    totals.Conditional += counts.Conditional; totals.Review += counts.Review;
    const moveRate = counts.total > 0 ? ((counts.Yes / counts.total) * 100).toFixed(1) : '0.0';
    const r = ws.getRow(rowNum);
    [key, counts.total, counts.Yes, counts.No, counts.Conditional, counts.Review, `${moveRate}%`].forEach((v, ci) => {
      const cell = r.getCell(ci + 1);
      cell.value = v;
      cell.border = BORDERS_ALL;
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
      cell.font = { size: 10, color: { argb: COLORS.black } };
      if (i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.altRow } };
    });
    _colorCountCell(r.getCell(3), counts.Yes, COLORS.greenLight, COLORS.green);
    _colorCountCell(r.getCell(4), counts.No, COLORS.redLight, COLORS.red);
    _colorCountCell(r.getCell(5), counts.Conditional, COLORS.purpleLight, COLORS.purple);
    _colorCountCell(r.getCell(6), counts.Review, COLORS.orangeLight, COLORS.orange);
    const rv = parseFloat(moveRate);
    r.getCell(7).font = { bold: true, size: 10, color: { argb: rv >= 80 ? COLORS.green : rv >= 50 ? COLORS.orange : COLORS.red } };
    rowNum++;
  });

  rowNum++;
  const totalRate = totals.total > 0 ? ((totals.Yes / totals.total) * 100).toFixed(1) : '0.0';
  const totRow = ws.getRow(rowNum);
  ['TOTAL', totals.total, totals.Yes, totals.No, totals.Conditional, totals.Review, `${totalRate}%`].forEach((v, i) => {
    const c = totRow.getCell(i + 1);
    c.value = v;
    c.font = { bold: true, size: 11, color: { argb: COLORS.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDERS_ALL;
  });
  totRow.height = 26;

  ws.getColumn(1).width = 35;
  for (let c = 2; c <= 7; c++) ws.getColumn(c).width = 16;
}

// ──────────────────────────────────────────────────────────
// Sheet 6 — By Status
// ──────────────────────────────────────────────────────────

function _buildStatusSheet(wb, assessed, statusKey, modeLabel) {
  const ws = wb.addWorksheet('By Status');

  ws.mergeCells('A1:F1');
  const tc = ws.getCell('A1');
  tc.value = `RESOURCES BY STATUS — ${modeLabel.toUpperCase()}`;
  tc.font = { bold: true, size: 14, color: { argb: COLORS.white } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primary } };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 32;

  const groups = { Yes: [], No: [], Conditional: [], Review: [] };
  assessed.forEach(r => {
    const s = r[statusKey] || 'Review';
    if (!groups[s]) groups[s] = [];
    groups[s].push(r);
  });

  const statusCfg = {
    Yes:         { label: modeLabel.includes('Jio') ? 'AVAILABLE' : 'CAN MOVE', bg: COLORS.green, light: COLORS.greenLight },
    No:          { label: modeLabel.includes('Jio') ? 'NOT AVAILABLE' : 'CANNOT MOVE', bg: COLORS.red, light: COLORS.redLight },
    Conditional: { label: 'CONDITIONAL', bg: COLORS.purple, light: COLORS.purpleLight },
    Review:      { label: 'NEEDS REVIEW', bg: COLORS.orange, light: COLORS.orangeLight },
  };

  let rowNum = 3;
  for (const [status, resources] of Object.entries(groups)) {
    if (resources.length === 0) continue;
    const cfg = statusCfg[status] || statusCfg.Review;

    ws.mergeCells(`A${rowNum}:F${rowNum}`);
    const gh = ws.getCell(`A${rowNum}`);
    gh.value = `${cfg.label} (${resources.length} resources)`;
    gh.font = { bold: true, size: 12, color: { argb: COLORS.white } };
    gh.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cfg.bg } };
    gh.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(rowNum).height = 28;
    rowNum++;

    const hdr = ws.getRow(rowNum);
    ['#', 'Resource Name', 'Resource Type', 'Resource Group', 'Location', 'Remarks'].forEach((h, i) => {
      const c = hdr.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, size: 10, color: { argb: COLORS.black } };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cfg.light } };
      c.border = BORDERS_ALL;
      c.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    hdr.height = 22;
    rowNum++;

    resources.forEach((r, idx) => {
      const dr = ws.getRow(rowNum);
      [idx + 1, _findField(r, 'NAME'), r['NORMALIZED TYPE'] || _findField(r, 'TYPE'), _findField(r, 'RESOURCE GROUP'), _findField(r, 'LOCATION'), r['REMARKS'] || ''].forEach((v, ci) => {
        const cell = dr.getCell(ci + 1);
        cell.value = v;
        cell.font = { size: 10, color: { argb: COLORS.black } };
        cell.border = BORDERS_ALL;
        cell.alignment = { vertical: 'middle', wrapText: ci === 5 };
        if (idx % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.altRow } };
      });
      rowNum++;
    });
    rowNum++;
  }

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 40;
  ws.getColumn(4).width = 25;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 55;
}

// ──────────────────────────────────────────────────────────
// Sheet 7 — Action Items
// ──────────────────────────────────────────────────────────

function _buildActionSheet(wb, assessed, statusKey, modeLabel) {
  const ws = wb.addWorksheet('Action Items');
  const actionable = assessed.filter(r => (r[statusKey] || 'Review') !== 'Yes');

  ws.mergeCells('A1:G1');
  const tc = ws.getCell('A1');
  tc.value = `ACTION ITEMS — ${modeLabel.toUpperCase()}`;
  tc.font = { bold: true, size: 14, color: { argb: COLORS.white } };
  tc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.red } };
  tc.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 32;

  ws.mergeCells('A2:G2');
  const sub = ws.getCell('A2');
  sub.value = `${actionable.length} resources require attention before migration`;
  sub.font = { size: 11, italic: true, color: { argb: COLORS.darkGray } };
  sub.alignment = { horizontal: 'center' };

  const hdrRow = ws.getRow(3);
  ['#', 'Resource Name', 'Resource Type', 'Resource Group', 'Location', 'Status', 'Action Required'].forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 11, color: { argb: COLORS.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
    c.border = BORDERS_ALL;
  });
  hdrRow.height = 26;

  if (actionable.length === 0) {
    ws.mergeCells('A4:G4');
    const ni = ws.getCell('A4');
    ni.value = 'No action items — all resources can be migrated!';
    ni.font = { bold: true, size: 12, color: { argb: COLORS.green } };
    ni.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(4).height = 30;
  } else {
    const priority = { No: 1, Conditional: 2, Review: 3 };
    actionable.sort((a, b) => (priority[a[statusKey] || 'Review'] || 9) - (priority[b[statusKey] || 'Review'] || 9));

    let rowNum = 4;
    actionable.forEach((r, idx) => {
      const status = r[statusKey] || 'Review';
      const dr = ws.getRow(rowNum);
      [idx + 1, _findField(r, 'NAME'), r['NORMALIZED TYPE'] || _findField(r, 'TYPE'), _findField(r, 'RESOURCE GROUP'), _findField(r, 'LOCATION'), status, r['REMARKS'] || 'Manual review required'].forEach((v, ci) => {
        const cell = dr.getCell(ci + 1);
        cell.value = v;
        cell.font = { size: 10, color: { argb: COLORS.black } };
        cell.border = BORDERS_ALL;
        cell.alignment = { vertical: 'middle', wrapText: ci === 6 };
        if (idx % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.altRow } };
      });
      const sc = dr.getCell(6);
      const c = _statusColors(status);
      sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: c.bg } };
      sc.font = { bold: true, size: 10, color: { argb: c.fg } };
      sc.alignment = { horizontal: 'center', vertical: 'middle' };
      rowNum++;
    });

    ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: 7 } };
  }

  ws.getColumn(1).width = 6;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 40;
  ws.getColumn(4).width = 25;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 60;
}

// ──────────────────────────────────────────────────────────
// Helpers
// ──────────────────────────────────────────────────────────

function _findField(row, field) {
  const keys = Object.keys(row);
  const upper = field.toUpperCase();
  const match = keys.find(k => k.toUpperCase() === upper);
  if (match) return String(row[match] || '');
  const partial = keys.find(k => k.toUpperCase().includes(upper) || upper.includes(k.toUpperCase()));
  return partial ? String(row[partial] || '') : '';
}

function _statusColors(val) {
  switch (val) {
    case 'Yes':         return { bg: COLORS.greenLight, fg: COLORS.green };
    case 'No':          return { bg: COLORS.redLight, fg: COLORS.red };
    case 'Conditional': return { bg: COLORS.purpleLight, fg: COLORS.purple };
    default:            return { bg: COLORS.orangeLight, fg: COLORS.orange };
  }
}

function _colorCountCell(cell, count, bgColor, fgColor) {
  if (count > 0) {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.font = { bold: true, size: 10, color: { argb: fgColor } };
  }
}

function _sectionHeader(ws, row, text, colStart, colEnd) {
  ws.mergeCells(`${colStart}${row}:${colEnd}${row}`);
  const cell = ws.getCell(`${colStart}${row}`);
  cell.value = text;
  cell.font = { bold: true, size: 12, color: { argb: COLORS.white } };
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLORS.primaryDark } };
  cell.alignment = { horizontal: 'left', vertical: 'middle' };
  cell.border = BORDERS_ALL;
  ws.getRow(row).height = 28;
}

function _labelValue(ws, row, label, value) {
  const lc = ws.getCell(`A${row}`);
  lc.value = label;
  lc.font = { bold: true, size: 11, color: { argb: COLORS.darkGray } };
  lc.border = BORDERS_ALL;
  ws.mergeCells(`B${row}:D${row}`);
  const vc = ws.getCell(`B${row}`);
  vc.value = value;
  vc.font = { size: 11, color: { argb: COLORS.black } };
  vc.border = BORDERS_ALL;
}

function _bar(pct) {
  const filled = Math.round(parseFloat(pct) / 5);
  return '█'.repeat(filled) + '░'.repeat(20 - filled) + ` ${pct}%`;
}

module.exports = { buildReport };
