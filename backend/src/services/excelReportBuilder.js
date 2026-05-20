const ExcelJS = require('exceljs');

// ─── Modern Color Palette (flat, contemporary) ───
const C = {
  // Primary
  brand:       '2563EB',
  brandDark:   '1D4ED8',
  brandLight:  'EFF6FF',
  // Status
  success:     '059669',
  successBg:   'ECFDF5',
  danger:      'DC2626',
  dangerBg:    'FEF2F2',
  warning:     'D97706',
  warningBg:   'FFFBEB',
  info:        '7C3AED',
  infoBg:      'F5F3FF',
  // Neutrals
  white:       'FFFFFF',
  gray50:      'F9FAFB',
  gray100:     'F3F4F6',
  gray200:     'E5E7EB',
  gray300:     'D1D5DB',
  gray400:     '9CA3AF',
  gray500:     '6B7280',
  gray700:     '374151',
  gray800:     '1F2937',
  gray900:     '111827',
};

const BORDER_THIN = { style: 'thin', color: { argb: 'FFD1D5DB' } };
const BORDER_BOTTOM = { bottom: BORDER_THIN };
const BORDER_ALL = { top: BORDER_THIN, bottom: BORDER_THIN, left: BORDER_THIN, right: BORDER_THIN };

/**
 * Build a modern Excel report.
 */
async function buildReport({ assessed, summary, mode, sheetName, outputPath }) {
  // AWS mode gets a completely different report format
  if (mode === 'aws') {
    return _buildAwsReport({ assessed, summary, sheetName, outputPath });
  }

  // GCP mode gets a report format similar to AWS
  if (mode === 'gcp') {
    return _buildGcpReport({ assessed, summary, sheetName, outputPath });
  }

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Resource Migration Assessment Tool';
  wb.created = new Date();

  const statusKey =
    mode === 'jio'    ? 'JIO REGION AVAILABLE'
    : mode === 'region' ? 'REGION MOVE SUPPORTED'
    : 'SUBSCRIPTION MOVE SUPPORTED';

  const modeLabel =
    mode === 'jio'    ? 'Jio Availability'
    : mode === 'region' ? 'Region Move'
    : 'Subscription Move';

  _buildCoverSheet(wb, assessed, summary, mode, modeLabel, statusKey);
  _buildDashboardSheet(wb, assessed, summary, mode, modeLabel, statusKey);
  _buildDataSheet(wb, assessed, statusKey, sheetName || 'Assessment Data');
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'PROVIDER');
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'RESOURCE GROUP');
  _buildPivotSheet(wb, assessed, statusKey, modeLabel, 'LOCATION');
  _buildActionSheet(wb, assessed, statusKey, modeLabel);

  await wb.xlsx.writeFile(outputPath);
}

// ══════════════════════════════════════════════════════════
// SHEET: Executive Summary (Cover)
// ══════════════════════════════════════════════════════════

function _buildCoverSheet(wb, assessed, summary, mode, modeLabel, statusKey) {
  const ws = wb.addWorksheet('Executive Summary', {
    properties: { tabColor: { argb: C.brand } },
    views: [{ showGridLines: false }]
  });
  const dateStr = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'long', timeStyle: 'short' });
  const version = '1.1.0';
  const cond = summary.conditional || 0;
  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';
  const readiness = parseFloat(pct(summary.yes));
  const readinessColor = readiness >= 80 ? C.success : readiness >= 50 ? C.warning : C.danger;

  // Column widths
  ws.getColumn(1).width = 3;
  ws.getColumn(2).width = 5;
  ws.getColumn(3).width = 20;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 18;
  ws.getColumn(7).width = 18;
  ws.getColumn(8).width = 18;
  ws.getColumn(9).width = 5;
  ws.getColumn(10).width = 3;

  let row = 1;
  ws.getRow(row).height = 15; row++;

  // ── Brand Header ──
  ws.mergeCells(row, 2, row, 9);
  const h1 = ws.getCell(row, 2);
  h1.value = 'Resource Migration Assessment';
  h1.font = { bold: true, size: 26, color: { argb: C.gray900 } };
  h1.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(row).height = 40;
  row++;

  ws.mergeCells(row, 2, row, 9);
  const sub = ws.getCell(row, 2);
  sub.value = `${modeLabel} Report  \u2022  ${dateStr}  \u2022  v${version}`;
  sub.font = { size: 11, color: { argb: C.gray500 } };
  sub.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(row).height = 22;
  row++;

  // Accent line
  ws.getRow(row).height = 5;
  for (let c = 2; c <= 9; c++) {
    ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.brand } };
  }
  row++;
  ws.getRow(row).height = 20; row++;

  // ── Readiness Score ──
  ws.mergeCells(row, 2, row, 5);
  const scoreLabel = ws.getCell(row, 2);
  scoreLabel.value = 'Migration Readiness';
  scoreLabel.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 22;
  row++;

  ws.mergeCells(row, 2, row + 1, 4);
  const scoreVal = ws.getCell(row, 2);
  scoreVal.value = `${pct(summary.yes)}%`;
  scoreVal.font = { bold: true, size: 48, color: { argb: readinessColor } };
  scoreVal.alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 5, row + 1, 9);
  const scoreNote = ws.getCell(row, 5);
  const readinessText = readiness >= 80 ? 'Excellent \u2014 Most resources are ready to migrate'
    : readiness >= 50 ? 'Moderate \u2014 Some resources need attention before migration'
    : 'Low \u2014 Significant blockers need to be resolved';
  scoreNote.value = readinessText;
  scoreNote.font = { size: 12, color: { argb: C.gray500 } };
  scoreNote.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  ws.getRow(row).height = 30;
  ws.getRow(row + 1).height = 30;
  row += 2;
  ws.getRow(row).height = 24; row++;

  // ── KPI Cards ──
  const kpis = [
    { label: 'Total', value: summary.total, color: C.brand, bg: C.brandLight },
    { label: mode === 'jio' ? 'Available' : 'Can Move', value: summary.yes, color: C.success, bg: C.successBg },
    { label: mode === 'jio' ? 'Not Available' : 'Cannot Move', value: summary.no, color: C.danger, bg: C.dangerBg },
    { label: 'Review', value: summary.review, color: C.warning, bg: C.warningBg },
  ];
  if (mode === 'subscription' && cond > 0) {
    kpis.push({ label: 'Conditional', value: cond, color: C.info, bg: C.infoBg });
  }

  const kpiCols = [3, 4, 5, 6, 7];

  // KPI header row
  ws.getRow(row).height = 26;
  kpis.forEach((kpi, idx) => {
    if (idx >= kpiCols.length) return;
    const col = kpiCols[idx];
    const cell = ws.getCell(row, col);
    cell.value = kpi.label;
    cell.font = { bold: true, size: 10, color: { argb: C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.color } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  row++;

  // KPI values row
  ws.getRow(row).height = 48;
  kpis.forEach((kpi, idx) => {
    if (idx >= kpiCols.length) return;
    const col = kpiCols[idx];
    const cell = ws.getCell(row, col);
    cell.value = kpi.value;
    cell.font = { bold: true, size: 28, color: { argb: kpi.color } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  row++;
  ws.getRow(row).height = 28; row++;

  // ── Scope ──
  ws.mergeCells(row, 2, row, 9);
  const scopeHdr = ws.getCell(row, 2);
  scopeHdr.value = 'Assessment Scope';
  scopeHdr.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 26;
  row++;

  for (let c = 2; c <= 9; c++) ws.getCell(row, c).border = { top: BORDER_THIN };
  ws.getRow(row).height = 4; row++;

  const providers = new Set(), rgs = new Set(), locs = new Set();
  assessed.forEach(r => {
    const t = (r['NORMALIZED TYPE'] || _findField(r, 'TYPE') || '').toLowerCase().split('/');
    if (t.length >= 2) providers.add(t[0].replace('microsoft.', ''));
    rgs.add(_findField(r, 'RESOURCE GROUP') || 'Unknown');
    locs.add(_findField(r, 'LOCATION') || 'Unknown');
  });

  const scopeItems = [
    ['Mode', modeLabel],
    ['Resources Analyzed', String(summary.total)],
    ['Resource Providers', `${providers.size}`],
    ['Resource Groups', `${rgs.size}`],
    ['Locations', `${locs.size}`],
    ['Data Source', 'Microsoft Learn (Real-time)'],
  ];

  scopeItems.forEach(([label, value], idx) => {
    const rowBg = idx % 2 === 0 ? C.gray50 : C.white;
    ws.mergeCells(row, 2, row, 4);
    const lCell = ws.getCell(row, 2);
    lCell.value = label;
    lCell.font = { bold: true, size: 11, color: { argb: C.gray700 } };
    lCell.alignment = { vertical: 'middle' };
    lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    lCell.border = BORDER_ALL;

    ws.mergeCells(row, 5, row, 9);
    const vCell = ws.getCell(row, 5);
    vCell.value = value;
    vCell.font = { size: 11, color: { argb: C.gray800 } };
    vCell.alignment = { vertical: 'middle' };
    vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    vCell.border = BORDER_ALL;
    ws.getRow(row).height = 22;
    row++;
  });

  ws.getRow(row).height = 28; row++;

  // ── Next Steps ──
  ws.mergeCells(row, 2, row, 9);
  const nsHdr = ws.getCell(row, 2);
  nsHdr.value = 'Recommended Actions';
  nsHdr.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 26;
  row++;

  for (let c = 2; c <= 9; c++) ws.getCell(row, c).border = { top: BORDER_THIN };
  ws.getRow(row).height = 4; row++;

  const steps = [];
  if (summary.yes > 0) steps.push({ icon: '\u25CF', text: `${summary.yes} resources ready \u2014 proceed with migration`, color: C.success });
  if (summary.review > 0) steps.push({ icon: '\u25CF', text: `${summary.review} resources need review \u2014 see Action Items`, color: C.warning });
  if (cond > 0) steps.push({ icon: '\u25CF', text: `${cond} conditional \u2014 verify prerequisites`, color: C.info });
  if (summary.no > 0) steps.push({ icon: '\u25CF', text: `${summary.no} blocked \u2014 plan alternatives`, color: C.danger });

  steps.forEach((step, idx) => {
    const stepBg = idx % 2 === 0 ? C.gray50 : C.white;
    ws.mergeCells(row, 3, row, 9);
    const icon = ws.getCell(row, 2);
    icon.value = step.icon;
    icon.font = { size: 12, color: { argb: step.color } };
    icon.alignment = { horizontal: 'center', vertical: 'middle' };
    icon.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: stepBg } };
    icon.border = BORDER_ALL;

    const txt = ws.getCell(row, 3);
    txt.value = step.text;
    txt.font = { size: 11, color: { argb: C.gray700 } };
    txt.alignment = { vertical: 'middle' };
    txt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: stepBg } };
    txt.border = BORDER_ALL;
    ws.getRow(row).height = 24;
    row++;
  });

  // Footer
  row += 2;
  ws.mergeCells(row, 2, row, 9);
  const ft = ws.getCell(row, 2);
  ft.value = `Generated by Resource Migration Assessment Tool v${version}`;
  ft.font = { size: 9, italic: true, color: { argb: C.gray400 } };
  ft.alignment = { horizontal: 'left' };
}

// ══════════════════════════════════════════════════════════
// SHEET: Dashboard
// ══════════════════════════════════════════════════════════

function _buildDashboardSheet(wb, assessed, summary, mode, modeLabel, statusKey) {
  const ws = wb.addWorksheet('Dashboard', {
    properties: { tabColor: { argb: C.brandDark } },
    views: [{ showGridLines: false }]
  });
  const cond = summary.conditional || 0;
  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';

  ws.getColumn(1).width = 3;
  ws.getColumn(2).width = 24;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 16;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 3;
  for (let c = 8; c <= 28; c++) ws.getColumn(c).width = 3;
  ws.getColumn(29).width = 12;

  let row = 1;
  ws.getRow(row).height = 8; row++;

  // Title
  ws.mergeCells(row, 2, row, 6);
  const title = ws.getCell(row, 2);
  title.value = `${modeLabel} \u2014 Dashboard`;
  title.font = { bold: true, size: 18, color: { argb: C.gray900 } };
  ws.getRow(row).height = 32;
  row++;

  // Accent line
  for (let c = 2; c <= 28; c++) ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.brand } };
  ws.getRow(row).height = 3; row++;
  ws.getRow(row).height = 16; row++;

  // ── Distribution Bar ──
  ws.mergeCells(row, 2, row, 6);
  const distTitle = ws.getCell(row, 2);
  distTitle.value = 'Status Distribution';
  distTitle.font = { bold: true, size: 11, color: { argb: C.gray700 } };
  ws.getRow(row).height = 22; row++;

  const BAR_START = 2;
  const BAR_COLS = 27;
  const statusItems = [
    { count: summary.yes, color: C.success, label: mode === 'jio' ? 'Available' : 'Can Move' },
    { count: summary.no, color: C.danger, label: mode === 'jio' ? 'Not Available' : 'Cannot Move' },
    { count: summary.review, color: C.warning, label: 'Review' },
  ];
  if (mode === 'subscription' && cond > 0) {
    statusItems.push({ count: cond, color: C.info, label: 'Conditional' });
  }

  ws.getRow(row).height = 22;
  let colOffset = BAR_START;
  statusItems.forEach(item => {
    if (item.count === 0) return;
    const cellCount = Math.max(1, Math.round((item.count / summary.total) * BAR_COLS));
    const endCol = Math.min(colOffset + cellCount - 1, BAR_START + BAR_COLS - 1);
    for (let c = colOffset; c <= endCol; c++) {
      ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: item.color } };
    }
    colOffset = endCol + 1;
  });
  row++;

  // Legend
  ws.getRow(row).height = 18;
  let legendCol = BAR_START;
  statusItems.forEach(item => {
    if (item.count === 0) return;
    const cell = ws.getCell(row, legendCol);
    cell.value = `\u25CF ${item.label}: ${item.count} (${pct(item.count)}%)`;
    cell.font = { size: 9, color: { argb: item.color } };
    legendCol += 7;
  });
  row++;
  ws.getRow(row).height = 20; row++;

  // ── Summary Table ──
  ws.mergeCells(row, 2, row, 6);
  const sumTitle = ws.getCell(row, 2);
  sumTitle.value = 'Breakdown';
  sumTitle.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 24; row++;

  const headers = ['Status', 'Count', 'Percentage', 'Visual', ''];
  ws.getRow(row).height = 26;
  headers.forEach((h, i) => {
    const cell = ws.getCell(row, i + 2);
    cell.value = h;
    cell.font = { bold: true, size: 10, color: { argb: C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.brand } };
    cell.border = BORDER_ALL;
    cell.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
  });
  row++;

  const tableRows = [
    [mode === 'jio' ? 'Available' : 'Can Move', summary.yes, pct(summary.yes), C.success],
    [mode === 'jio' ? 'Not Available' : 'Cannot Move', summary.no, pct(summary.no), C.danger],
    ['Needs Review', summary.review, pct(summary.review), C.warning],
  ];
  if (mode === 'subscription') tableRows.push(['Conditional', cond, pct(cond), C.info]);

  tableRows.forEach(([label, count, percent, color], idx) => {
    ws.getRow(row).height = 24;
    const rowBg = idx % 2 === 0 ? C.white : C.gray50;

    ws.getCell(row, 2).value = label;
    ws.getCell(row, 2).font = { bold: true, size: 11, color: { argb: color } };
    ws.getCell(row, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    ws.getCell(row, 2).border = BORDER_ALL;

    ws.getCell(row, 3).value = count;
    ws.getCell(row, 3).font = { bold: true, size: 12, color: { argb: color } };
    ws.getCell(row, 3).alignment = { horizontal: 'center' };
    ws.getCell(row, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    ws.getCell(row, 3).border = BORDER_ALL;

    ws.getCell(row, 4).value = `${percent}%`;
    ws.getCell(row, 4).font = { bold: true, size: 10, color: { argb: C.gray700 } };
    ws.getCell(row, 4).alignment = { horizontal: 'center' };
    ws.getCell(row, 4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    ws.getCell(row, 4).border = BORDER_ALL;

    const barLen = Math.max(1, Math.round(parseFloat(percent) / 10));
    ws.getCell(row, 5).value = '\u2588'.repeat(barLen) + '\u2591'.repeat(10 - barLen);
    ws.getCell(row, 5).font = { size: 10, color: { argb: color } };
    ws.getCell(row, 5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    ws.getCell(row, 5).border = BORDER_ALL;

    ws.getCell(row, 6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    ws.getCell(row, 6).border = BORDER_ALL;
    row++;
  });

  // Total row
  ws.getRow(row).height = 26;
  ws.getCell(row, 2).value = 'Total';
  ws.getCell(row, 2).font = { bold: true, size: 11, color: { argb: C.white } };
  ws.getCell(row, 2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
  ws.getCell(row, 2).border = BORDER_ALL;
  ws.getCell(row, 3).value = summary.total;
  ws.getCell(row, 3).font = { bold: true, size: 13, color: { argb: C.white } };
  ws.getCell(row, 3).alignment = { horizontal: 'center' };
  ws.getCell(row, 3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
  ws.getCell(row, 3).border = BORDER_ALL;
  ws.getCell(row, 4).value = '100%';
  ws.getCell(row, 4).font = { bold: true, size: 10, color: { argb: C.white } };
  ws.getCell(row, 4).alignment = { horizontal: 'center' };
  ws.getCell(row, 4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
  ws.getCell(row, 4).border = BORDER_ALL;
  for (let c = 5; c <= 6; c++) {
    ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
    ws.getCell(row, c).border = BORDER_ALL;
  }
  row++;
  ws.getRow(row).height = 24; row++;

  // ── Top Providers ──
  ws.mergeCells(row, 2, row, 6);
  const provTitle = ws.getCell(row, 2);
  provTitle.value = 'Top Resource Providers';
  provTitle.font = { bold: true, size: 11, color: { argb: C.gray700 } };
  ws.getRow(row).height = 24; row++;

  const providerCounts = {};
  assessed.forEach(r => {
    const t = (r['NORMALIZED TYPE'] || _findField(r, 'TYPE') || '').toLowerCase().split('/');
    if (t.length >= 2) {
      const pName = t[0].replace('microsoft.', '');
      providerCounts[pName] = (providerCounts[pName] || 0) + 1;
    }
  });

  const sortedProviders = Object.entries(providerCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);
  const maxCount = sortedProviders.length > 0 ? sortedProviders[0][1] : 1;
  const provColors = [C.brand, C.success, C.warning, C.info, C.danger, C.brandDark, '6366F1', '0891B2'];

  sortedProviders.forEach(([name, count], idx) => {
    ws.getRow(row).height = 20;
    ws.getCell(row, 2).value = name;
    ws.getCell(row, 2).font = { size: 10, color: { argb: C.gray700 } };
    ws.getCell(row, 3).value = count;
    ws.getCell(row, 3).font = { bold: true, size: 10, color: { argb: provColors[idx % provColors.length] } };
    ws.getCell(row, 3).alignment = { horizontal: 'center' };

    const barLen = Math.max(1, Math.round((count / maxCount) * 15));
    ws.getCell(row, 4).value = '\u2501'.repeat(barLen);
    ws.getCell(row, 4).font = { size: 9, color: { argb: provColors[idx % provColors.length] } };

    for (let c = 2; c <= 6; c++) ws.getCell(row, c).border = BORDER_BOTTOM;
    row++;
  });
}

// ══════════════════════════════════════════════════════════
// SHEET: Assessment Data
// ══════════════════════════════════════════════════════════

function _buildDataSheet(wb, assessed, statusKey, title) {
  if (assessed.length === 0) return;
  const ws = wb.addWorksheet(title, {
    views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
  });
  const keys = Object.keys(assessed[0]);

  const headerRow = ws.addRow(keys);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.brand } };
    cell.font = { bold: true, color: { argb: C.white }, size: 10 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_ALL;
  });
  headerRow.height = 30;

  const statusColIdx = keys.indexOf(statusKey) + 1;

  assessed.forEach((row, i) => {
    const dr = ws.addRow(keys.map(k => row[k] ?? ''));
    const isAlt = i % 2 === 1;
    dr.eachCell(cell => {
      cell.font = { size: 10, color: { argb: C.gray700 } };
      cell.alignment = { vertical: 'middle', wrapText: true };
      cell.border = BORDER_ALL;
      if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
    });
    if (statusColIdx > 0) {
      const sc = dr.getCell(statusColIdx);
      const colors = _statusColors(String(sc.value || ''));
      sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } };
      sc.font = { bold: true, size: 10, color: { argb: colors.fg } };
      sc.alignment = { horizontal: 'center', vertical: 'middle' };
    }
    dr.height = 22;
  });

  ws.columns.forEach(col => {
    let maxLen = 10;
    col.eachCell({ includeEmpty: false }, cell => {
      const len = String(cell.value || '').length;
      if (len > maxLen) maxLen = len;
    });
    col.width = Math.min(maxLen + 3, 45);
  });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: assessed.length + 1, column: keys.length } };
}

// ══════════════════════════════════════════════════════════
// SHEET: Pivot Analysis
// ══════════════════════════════════════════════════════════

function _buildPivotSheet(wb, assessed, statusKey, modeLabel, pivotBy) {
  const sheetNames = { PROVIDER: 'By Provider', 'RESOURCE GROUP': 'By Resource Group', LOCATION: 'By Location' };
  const ws = wb.addWorksheet(sheetNames[pivotBy], {
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

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

  const colLabel = pivotBy === 'PROVIDER' ? 'Resource Provider' : pivotBy === 'RESOURCE GROUP' ? 'Resource Group' : 'Location';

  // Title
  ws.mergeCells('A1:G1');
  const tc = ws.getCell('A1');
  tc.value = `${colLabel} Analysis`;
  tc.font = { bold: true, size: 16, color: { argb: C.gray900 } };
  tc.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:G2');
  const st = ws.getCell('A2');
  st.value = `${Object.keys(pivot).length} unique ${pivotBy.toLowerCase()}s  \u2022  ${modeLabel}`;
  st.font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  // Header
  const headers = [colLabel, 'Total', 'Yes', 'No', 'Conditional', 'Review', 'Readiness'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.brand } };
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  const sorted = Object.entries(pivot).sort((a, b) => b[1].total - a[1].total);
  let rowNum = 4;

  sorted.forEach(([key, counts], i) => {
    const moveRate = counts.total > 0 ? ((counts.Yes / counts.total) * 100).toFixed(1) : '0.0';
    const r = ws.getRow(rowNum);
    r.height = 22;

    [key, counts.total, counts.Yes, counts.No, counts.Conditional, counts.Review, `${moveRate}%`].forEach((v, ci) => {
      const cell = r.getCell(ci + 1);
      cell.value = v;
      cell.alignment = { horizontal: ci === 0 ? 'left' : 'center', vertical: 'middle' };
      cell.font = { size: 10, color: { argb: C.gray700 } };
      cell.border = BORDER_ALL;
      if (i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
    });

    if (counts.Yes > 0) r.getCell(3).font = { bold: true, size: 10, color: { argb: C.success } };
    if (counts.No > 0) r.getCell(4).font = { bold: true, size: 10, color: { argb: C.danger } };
    if (counts.Conditional > 0) r.getCell(5).font = { bold: true, size: 10, color: { argb: C.info } };
    if (counts.Review > 0) r.getCell(6).font = { bold: true, size: 10, color: { argb: C.warning } };

    const rv = parseFloat(moveRate);
    r.getCell(7).font = { bold: true, size: 10, color: { argb: rv >= 80 ? C.success : rv >= 50 ? C.warning : C.danger } };
    rowNum++;
  });

  // Total row
  rowNum++;
  const totals = sorted.reduce((acc, [, c]) => {
    acc.total += c.total; acc.Yes += c.Yes; acc.No += c.No; acc.Conditional += c.Conditional; acc.Review += c.Review;
    return acc;
  }, { total: 0, Yes: 0, No: 0, Conditional: 0, Review: 0 });
  const totalRate = totals.total > 0 ? ((totals.Yes / totals.total) * 100).toFixed(1) : '0.0';

  const totRow = ws.getRow(rowNum);
  totRow.height = 28;
  ['TOTAL', totals.total, totals.Yes, totals.No, totals.Conditional, totals.Review, `${totalRate}%`].forEach((v, i) => {
    const c = totRow.getCell(i + 1);
    c.value = v;
    c.font = { bold: true, size: 11, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });

  ws.getColumn(1).width = 35;
  for (let c = 2; c <= 7; c++) ws.getColumn(c).width = 14;
}

// ══════════════════════════════════════════════════════════
// SHEET: Action Items
// ══════════════════════════════════════════════════════════

function _buildActionSheet(wb, assessed, statusKey, modeLabel) {
  const ws = wb.addWorksheet('Action Items', {
    properties: { tabColor: { argb: C.danger } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

  const actionable = assessed.filter(r => (r[statusKey] || 'Review') !== 'Yes');

  ws.mergeCells('A1:G1');
  const tc = ws.getCell('A1');
  tc.value = 'Action Items';
  tc.font = { bold: true, size: 16, color: { argb: C.gray900 } };
  tc.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:G2');
  const st = ws.getCell('A2');
  st.value = `${actionable.length} resources require attention  \u2022  ${modeLabel}`;
  st.font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  const hdrRow = ws.getRow(3);
  ['#', 'Resource Name', 'Resource Type', 'Resource Group', 'Location', 'Status', 'Action Required'].forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.danger } };
    c.alignment = { horizontal: i === 0 ? 'center' : 'left', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  if (actionable.length === 0) {
    ws.mergeCells('A4:G4');
    const ni = ws.getCell('A4');
    ni.value = '\u2713  All resources can be migrated \u2014 no action items';
    ni.font = { size: 12, color: { argb: C.success } };
    ni.alignment = { horizontal: 'center', vertical: 'middle' };
    ws.getRow(4).height = 40;
  } else {
    const priority = { No: 1, Conditional: 2, Review: 3 };
    actionable.sort((a, b) => (priority[a[statusKey] || 'Review'] || 9) - (priority[b[statusKey] || 'Review'] || 9));

    let rowNum = 4;
    actionable.forEach((r, idx) => {
      const status = r[statusKey] || 'Review';
      const dr = ws.getRow(rowNum);
      dr.height = 22;

      [idx + 1, _findField(r, 'NAME'), r['NORMALIZED TYPE'] || _findField(r, 'TYPE'), _findField(r, 'RESOURCE GROUP'), _findField(r, 'LOCATION'), status, r['REMARKS'] || 'Manual review required'].forEach((v, ci) => {
        const cell = dr.getCell(ci + 1);
        cell.value = v;
        cell.font = { size: 10, color: { argb: C.gray700 } };
        cell.alignment = { vertical: 'middle', wrapText: ci === 6 };
        cell.border = BORDER_ALL;
        if (idx % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
      });

      const sc = dr.getCell(6);
      const colors = _statusColors(status);
      sc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.bg } };
      sc.font = { bold: true, size: 9, color: { argb: colors.fg } };
      sc.alignment = { horizontal: 'center', vertical: 'middle' };
      rowNum++;
    });

    ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: 7 } };
  }

  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 28;
  ws.getColumn(3).width = 35;
  ws.getColumn(4).width = 22;
  ws.getColumn(5).width = 16;
  ws.getColumn(6).width = 14;
  ws.getColumn(7).width = 50;
}

// ══════════════════════════════════════════════════════════
// Helpers
// ══════════════════════════════════════════════════════════

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
    case 'Yes':         return { bg: C.successBg, fg: C.success };
    case 'No':          return { bg: C.dangerBg, fg: C.danger };
    case 'Conditional': return { bg: C.infoBg, fg: C.info };
    default:            return { bg: C.warningBg, fg: C.warning };
  }
}

// ══════════════════════════════════════════════════════════
// AWS-Specific Report (Different from Subscription/Region)
// ══════════════════════════════════════════════════════════

const AWS_COLORS = {
  orange:      'F59E0B',
  orangeDark:  'D97706',
  orangeLight: 'FFFBEB',
  azure:       '0078D4',
  azureLight:  'EBF5FF',
  teal:        '0D9488',
  tealLight:   'F0FDFA',
  indigo:      '4F46E5',
  indigoLight: 'EEF2FF',
};

// Derive migration strategy from similarity
function _migrationStrategy(similarity) {
  switch (similarity) {
    case 'Direct Equivalent': return { strategy: 'Rehost / Lift & Shift', effort: 'Low', weeks: '1-2', risk: 'Low' };
    case 'Similar':           return { strategy: 'Replatform', effort: 'Medium', weeks: '2-4', risk: 'Medium' };
    case 'Partial':           return { strategy: 'Refactor', effort: 'High', weeks: '4-8', risk: 'High' };
    case 'No Direct Mapping': return { strategy: 'Re-architect / Replace', effort: 'Very High', weeks: '6-12+', risk: 'Critical' };
    default:                  return { strategy: 'Review Required', effort: 'TBD', weeks: 'TBD', risk: 'Unknown' };
  }
}

function _effortColor(effort) {
  switch (effort) {
    case 'Low':       return C.success;
    case 'Medium':    return C.info;
    case 'High':      return C.warning;
    case 'Very High': return C.danger;
    default:          return C.gray500;
  }
}

function _riskColor(risk) {
  switch (risk) {
    case 'Low':      return C.success;
    case 'Medium':   return C.info;
    case 'High':     return C.warning;
    case 'Critical': return C.danger;
    default:         return C.gray500;
  }
}

async function _buildAwsReport({ assessed, summary, sheetName, outputPath }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Resource Migration Assessment Tool';
  wb.created = new Date();

  _buildAwsCoverSheet(wb, assessed, summary);
  _buildAwsDataSheet(wb, assessed, sheetName || 'Service Mapping');
  _buildAwsCategorySheet(wb, assessed, summary);
  _buildAwsServiceGroupSheet(wb, assessed);
  _buildAwsMigrationRoadmapSheet(wb, assessed, summary);
  _buildAwsRiskMatrixSheet(wb, assessed);

  await wb.xlsx.writeFile(outputPath);
}

function _buildAwsCoverSheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('Executive Summary', {
    properties: { tabColor: { argb: AWS_COLORS.orange } },
    views: [{ showGridLines: false }]
  });
  const dateStr = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'long', timeStyle: 'short' });
  const version = '1.2.0';
  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';
  const coverage = summary.total > 0 ? (((summary.directEquivalent + summary.similar) / summary.total) * 100).toFixed(1) : '0.0';
  const coverageVal = parseFloat(coverage);
  const coverageColor = coverageVal >= 80 ? C.success : coverageVal >= 50 ? C.warning : C.danger;

  // Readiness rating
  const readiness = coverageVal >= 80 ? { label: 'READY TO MIGRATE', color: C.success, bg: C.successBg }
    : coverageVal >= 50 ? { label: 'NEEDS PLANNING', color: C.warning, bg: C.warningBg }
    : { label: 'COMPLEX MIGRATION', color: C.danger, bg: C.dangerBg };

  ws.getColumn(1).width = 3;
  ws.getColumn(2).width = 5;
  ws.getColumn(3).width = 22;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 18;
  ws.getColumn(7).width = 18;
  ws.getColumn(8).width = 18;
  ws.getColumn(9).width = 5;
  ws.getColumn(10).width = 3;

  let row = 1;
  ws.getRow(row).height = 15; row++;

  // Brand Header
  ws.mergeCells(row, 2, row, 9);
  const h1 = ws.getCell(row, 2);
  h1.value = 'AWS to Azure Migration Assessment';
  h1.font = { bold: true, size: 26, color: { argb: C.gray900 } };
  h1.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(row).height = 40;
  row++;

  ws.mergeCells(row, 2, row, 9);
  const sub = ws.getCell(row, 2);
  sub.value = `Cross-Cloud Service Mapping Report  \u2022  ${dateStr}  \u2022  v${version}`;
  sub.font = { size: 11, color: { argb: C.gray500 } };
  sub.alignment = { horizontal: 'left', vertical: 'middle' };
  ws.getRow(row).height = 22;
  row++;

  // Accent line (orange)
  ws.getRow(row).height = 5;
  for (let c = 2; c <= 9; c++) {
    ws.getCell(row, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AWS_COLORS.orange } };
  }
  row++;
  ws.getRow(row).height = 20; row++;

  // ── Migration Readiness Badge ──
  ws.mergeCells(row, 2, row, 4);
  const readyLabel = ws.getCell(row, 2);
  readyLabel.value = 'Migration Readiness';
  readyLabel.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.mergeCells(row, 5, row, 7);
  const readyBadge = ws.getCell(row, 5);
  readyBadge.value = `  ${readiness.label}  `;
  readyBadge.font = { bold: true, size: 14, color: { argb: readiness.color } };
  readyBadge.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: readiness.bg } };
  readyBadge.alignment = { horizontal: 'center', vertical: 'middle' };
  readyBadge.border = BORDER_ALL;
  ws.getRow(row).height = 32;
  row++;
  ws.getRow(row).height = 16; row++;

  // ── Coverage Score + Description ──
  ws.mergeCells(row, 2, row, 5);
  const scoreLabel = ws.getCell(row, 2);
  scoreLabel.value = 'Service Mapping Coverage';
  scoreLabel.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 22;
  row++;

  ws.mergeCells(row, 2, row + 1, 4);
  const scoreVal = ws.getCell(row, 2);
  scoreVal.value = `${coverage}%`;
  scoreVal.font = { bold: true, size: 48, color: { argb: coverageColor } };
  scoreVal.alignment = { horizontal: 'left', vertical: 'middle' };

  ws.mergeCells(row, 5, row + 1, 9);
  const scoreNote = ws.getCell(row, 5);
  const coverageText = coverageVal >= 80 ? 'Excellent \u2014 Most AWS services have direct or similar Azure equivalents. Migration can proceed with confidence.'
    : coverageVal >= 50 ? 'Moderate \u2014 Some services need alternative solutions. Plan refactoring for partial/unmapped services.'
    : 'Low \u2014 Many services require custom migration approaches. Conduct deep-dive architecture review.';
  scoreNote.value = coverageText;
  scoreNote.font = { size: 11, color: { argb: C.gray500 } };
  scoreNote.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
  ws.getRow(row).height = 30;
  ws.getRow(row + 1).height = 30;
  row += 2;
  ws.getRow(row).height = 20; row++;

  // ── KPI Cards ──
  const kpis = [
    { label: 'Total Services', value: summary.total, color: AWS_COLORS.orange, bg: AWS_COLORS.orangeLight },
    { label: 'Direct Equivalent', value: summary.directEquivalent, color: C.success, bg: C.successBg },
    { label: 'Similar', value: summary.similar, color: C.info, bg: C.infoBg },
    { label: 'Partial', value: summary.partial, color: C.warning, bg: C.warningBg },
    { label: 'No Mapping', value: summary.noMapping, color: C.danger, bg: C.dangerBg },
  ];

  const kpiCols = [3, 4, 5, 6, 7];
  ws.getRow(row).height = 26;
  kpis.forEach((kpi, idx) => {
    if (idx >= kpiCols.length) return;
    const col = kpiCols[idx];
    const cell = ws.getCell(row, col);
    cell.value = kpi.label;
    cell.font = { bold: true, size: 10, color: { argb: C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.color } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  row++;

  ws.getRow(row).height = 48;
  kpis.forEach((kpi, idx) => {
    if (idx >= kpiCols.length) return;
    const col = kpiCols[idx];
    const cell = ws.getCell(row, col);
    cell.value = kpi.value;
    cell.font = { bold: true, size: 28, color: { argb: kpi.color } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  row++;

  // Percentage row
  ws.getRow(row).height = 20;
  kpis.forEach((kpi, idx) => {
    if (idx >= kpiCols.length || idx === 0) return;
    const col = kpiCols[idx];
    const cell = ws.getCell(row, col);
    cell.value = `${pct(kpi.value)}%`;
    cell.font = { size: 9, color: { argb: C.gray500 } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });
  row++;
  ws.getRow(row).height = 24; row++;

  // ── Migration Effort Breakdown ──
  ws.mergeCells(row, 2, row, 9);
  const effortHdr = ws.getCell(row, 2);
  effortHdr.value = 'Estimated Migration Effort';
  effortHdr.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 26;
  row++;

  for (let c = 2; c <= 9; c++) ws.getCell(row, c).border = { top: BORDER_THIN };
  ws.getRow(row).height = 4; row++;

  const effortData = [
    { label: 'Phase 1: Lift & Shift', count: summary.directEquivalent, timeline: '1-2 weeks per service', color: C.success },
    { label: 'Phase 2: Replatform', count: summary.similar, timeline: '2-4 weeks per service', color: C.info },
    { label: 'Phase 3: Refactor', count: summary.partial, timeline: '4-8 weeks per service', color: C.warning },
    { label: 'Phase 4: Re-architect', count: summary.noMapping, timeline: '6-12+ weeks per service', color: C.danger },
  ];

  effortData.forEach((phase, idx) => {
    if (phase.count === 0) return;
    const rowBg = idx % 2 === 0 ? C.gray50 : C.white;
    ws.mergeCells(row, 2, row, 4);
    const lbl = ws.getCell(row, 2);
    lbl.value = phase.label;
    lbl.font = { bold: true, size: 11, color: { argb: phase.color } };
    lbl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    lbl.alignment = { vertical: 'middle' };
    lbl.border = BORDER_ALL;

    const cnt = ws.getCell(row, 5);
    cnt.value = `${phase.count} service${phase.count !== 1 ? 's' : ''}`;
    cnt.font = { bold: true, size: 11, color: { argb: C.gray800 } };
    cnt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    cnt.alignment = { horizontal: 'center', vertical: 'middle' };
    cnt.border = BORDER_ALL;

    ws.mergeCells(row, 6, row, 9);
    const tl = ws.getCell(row, 6);
    tl.value = phase.timeline;
    tl.font = { size: 10, italic: true, color: { argb: C.gray500 } };
    tl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    tl.alignment = { vertical: 'middle' };
    tl.border = BORDER_ALL;
    ws.getRow(row).height = 24;
    row++;
  });
  ws.getRow(row).height = 24; row++;

  // ── Assessment Scope ──
  ws.mergeCells(row, 2, row, 9);
  const scopeHdr = ws.getCell(row, 2);
  scopeHdr.value = 'Assessment Scope';
  scopeHdr.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 26;
  row++;

  for (let c = 2; c <= 9; c++) ws.getCell(row, c).border = { top: BORDER_THIN };
  ws.getRow(row).height = 4; row++;

  const awsNamespaces = new Set();
  const azureServices = new Set();
  assessed.forEach(r => {
    const svc = (r['AWS SERVICE'] || '').toLowerCase();
    if (svc.startsWith('aws::')) {
      const parts = svc.split('::');
      if (parts.length >= 2) awsNamespaces.add(parts[1]);
    }
    const azSvc = (r['AZURE EQUIVALENT'] || '').trim();
    if (azSvc && azSvc !== 'No Azure Equivalent Found' && azSvc !== 'Review Required') azureServices.add(azSvc);
  });

  const scopeItems = [
    ['Migration Type', 'AWS \u2192 Azure (Cross-Cloud)'],
    ['AWS Services Analyzed', String(summary.total)],
    ['AWS Namespaces', `${awsNamespaces.size || 'N/A'} (${[...awsNamespaces].slice(0, 5).join(', ')}${awsNamespaces.size > 5 ? '...' : ''})`],
    ['Azure Services Mapped', `${azureServices.size} unique Azure services`],
    ['Mapping Coverage', `${coverage}% (Direct + Similar)`],
    ['Data Source', 'Curated AWS-Azure Service Mapping (microsoft-learn)'],
  ];

  scopeItems.forEach(([label, value], idx) => {
    const rowBg = idx % 2 === 0 ? C.gray50 : C.white;
    ws.mergeCells(row, 2, row, 4);
    const lCell = ws.getCell(row, 2);
    lCell.value = label;
    lCell.font = { bold: true, size: 11, color: { argb: C.gray700 } };
    lCell.alignment = { vertical: 'middle' };
    lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    lCell.border = BORDER_ALL;

    ws.mergeCells(row, 5, row, 9);
    const vCell = ws.getCell(row, 5);
    vCell.value = value;
    vCell.font = { size: 11, color: { argb: C.gray800 } };
    vCell.alignment = { vertical: 'middle' };
    vCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    vCell.border = BORDER_ALL;
    ws.getRow(row).height = 22;
    row++;
  });

  ws.getRow(row).height = 24; row++;

  // ── Recommended Actions ──
  ws.mergeCells(row, 2, row, 9);
  const nsHdr = ws.getCell(row, 2);
  nsHdr.value = 'Recommended Next Steps';
  nsHdr.font = { bold: true, size: 12, color: { argb: C.gray700 } };
  ws.getRow(row).height = 26;
  row++;

  for (let c = 2; c <= 9; c++) ws.getCell(row, c).border = { top: BORDER_THIN };
  ws.getRow(row).height = 4; row++;

  const steps = [
    { icon: '1', text: 'Review "Service Mapping" sheet for complete AWS\u2192Azure mapping details', color: AWS_COLORS.orange },
    { icon: '2', text: 'Check "Migration Roadmap" sheet for phased migration timeline', color: AWS_COLORS.teal },
    { icon: '3', text: 'Assess "Risk Matrix" sheet to identify high-risk services requiring attention', color: C.danger },
  ];
  if (summary.directEquivalent > 0) steps.push({ icon: '4', text: `Start with ${summary.directEquivalent} Direct Equivalent services \u2014 lowest effort, quickest wins`, color: C.success });
  if (summary.noMapping > 0) steps.push({ icon: '5', text: `Plan architecture review for ${summary.noMapping} unmapped services`, color: C.danger });

  steps.forEach((step, idx) => {
    const stepBg = idx % 2 === 0 ? C.gray50 : C.white;
    const numCell = ws.getCell(row, 2);
    numCell.value = step.icon;
    numCell.font = { bold: true, size: 12, color: { argb: C.white } };
    numCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: step.color } };
    numCell.alignment = { horizontal: 'center', vertical: 'middle' };
    numCell.border = BORDER_ALL;

    ws.mergeCells(row, 3, row, 9);
    const txt = ws.getCell(row, 3);
    txt.value = step.text;
    txt.font = { size: 11, color: { argb: C.gray700 } };
    txt.alignment = { vertical: 'middle' };
    txt.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: stepBg } };
    txt.border = BORDER_ALL;
    ws.getRow(row).height = 26;
    row++;
  });

  // Footer
  row += 2;
  ws.mergeCells(row, 2, row, 9);
  const ft = ws.getCell(row, 2);
  ft.value = `Generated by Resource Migration Assessment Tool v${version}  \u2022  Report contains ${wb.worksheets.length || 6} sheets`;
  ft.font = { size: 9, italic: true, color: { argb: C.gray400 } };
  ft.alignment = { horizontal: 'left' };
}

function _buildAwsDataSheet(wb, assessed, title) {
  if (assessed.length === 0) return;
  const ws = wb.addWorksheet(title, {
    views: [{ state: 'frozen', ySplit: 1, showGridLines: false }]
  });

  const columns = ['#', 'NAME', 'AWS SERVICE', 'AZURE EQUIVALENT', 'AZURE RESOURCE TYPE', 'CATEGORY', 'SIMILARITY', 'MIGRATION STRATEGY', 'EFFORT', 'DESCRIPTION', 'MIGRATION NOTES'];

  const headerRow = ws.addRow(columns);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AWS_COLORS.orange } };
    cell.font = { bold: true, color: { argb: C.white }, size: 10 };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = BORDER_ALL;
  });
  headerRow.height = 32;

  assessed.forEach((row, i) => {
    const ms = _migrationStrategy(row['SIMILARITY']);
    const rowData = [
      i + 1,
      row['NAME'] ?? '',
      row['AWS SERVICE'] ?? '',
      row['AZURE EQUIVALENT'] ?? '',
      row['AZURE RESOURCE TYPE'] ?? '',
      row['CATEGORY'] ?? '',
      row['SIMILARITY'] ?? '',
      ms.strategy,
      ms.effort,
      row['DESCRIPTION'] ?? '',
      row['MIGRATION NOTES'] ?? ''
    ];
    const dr = ws.addRow(rowData);
    const isAlt = i % 2 === 1;
    dr.eachCell((cell, colNum) => {
      cell.font = { size: 10, color: { argb: C.gray700 } };
      cell.alignment = { vertical: 'middle', wrapText: colNum >= 10 };
      cell.border = BORDER_ALL;
      if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
    });

    // # column (col 1) - subtle
    dr.getCell(1).font = { size: 9, color: { argb: C.gray400 } };
    dr.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    // Color the SIMILARITY column (column 7)
    const simCell = dr.getCell(7);
    const simColors = _awsSimilarityColors(String(simCell.value || ''));
    simCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: simColors.bg } };
    simCell.font = { bold: true, size: 10, color: { argb: simColors.fg } };
    simCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Color the CATEGORY column (column 6)
    const catCell = dr.getCell(6);
    catCell.font = { bold: true, size: 10, color: { argb: AWS_COLORS.azure } };
    catCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Color MIGRATION STRATEGY column (col 8)
    dr.getCell(8).font = { size: 10, color: { argb: AWS_COLORS.teal } };
    dr.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };

    // Color EFFORT column (col 9)
    dr.getCell(9).font = { bold: true, size: 10, color: { argb: _effortColor(ms.effort) } };
    dr.getCell(9).alignment = { horizontal: 'center', vertical: 'middle' };

    dr.height = 24;
  });

  // Set column widths
  const widths = [5, 22, 32, 28, 32, 16, 18, 22, 12, 40, 40];
  widths.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: assessed.length + 1, column: columns.length } };
}

function _buildAwsCategorySheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('By Category', {
    properties: { tabColor: { argb: C.info } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';

  // Title
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'Mapping Category Breakdown';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: C.gray900 } };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:H2');
  ws.getCell('A2').value = `${summary.total} AWS services assessed  \u2022  ${pct(summary.directEquivalent + summary.similar)}% have clear Azure path`;
  ws.getCell('A2').font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  // Header row
  const headers = ['Category', 'Count', '%', 'Visual', 'Strategy', 'Effort', 'Timeline', 'Example Services'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AWS_COLORS.orange } };
    c.alignment = { horizontal: i === 0 || i === 7 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  // Get example services per category
  const examplesByCategory = {};
  assessed.forEach(r => {
    const sim = r['SIMILARITY'] || 'No Direct Mapping';
    if (!examplesByCategory[sim]) examplesByCategory[sim] = [];
    if (examplesByCategory[sim].length < 3) {
      const name = r['AWS SERVICE'] || r['NAME'] || '';
      if (name) examplesByCategory[sim].push(name.replace(/^aws::/i, '').replace(/::/g, ' '));
    }
  });

  const categories = [
    { label: 'Direct Equivalent', count: summary.directEquivalent, color: C.success, strategy: 'Rehost / Lift & Shift', effort: 'Low', timeline: '1-2 weeks' },
    { label: 'Similar', count: summary.similar, color: C.info, strategy: 'Replatform', effort: 'Medium', timeline: '2-4 weeks' },
    { label: 'Partial', count: summary.partial, color: C.warning, strategy: 'Refactor', effort: 'High', timeline: '4-8 weeks' },
    { label: 'No Direct Mapping', count: summary.noMapping, color: C.danger, strategy: 'Re-architect', effort: 'Very High', timeline: '6-12+ weeks' },
  ];

  let rowNum = 4;
  categories.forEach((cat, idx) => {
    const r = ws.getRow(rowNum);
    r.height = 28;
    const rowBg = idx % 2 === 0 ? C.white : C.gray50;

    r.getCell(1).value = cat.label;
    r.getCell(1).font = { bold: true, size: 11, color: { argb: cat.color } };
    r.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(1).border = BORDER_ALL;

    r.getCell(2).value = cat.count;
    r.getCell(2).font = { bold: true, size: 14, color: { argb: cat.color } };
    r.getCell(2).alignment = { horizontal: 'center' };
    r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(2).border = BORDER_ALL;

    r.getCell(3).value = `${pct(cat.count)}%`;
    r.getCell(3).font = { bold: true, size: 10, color: { argb: C.gray700 } };
    r.getCell(3).alignment = { horizontal: 'center' };
    r.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(3).border = BORDER_ALL;

    // Visual bar - wider (20 blocks)
    const barLen = Math.max(1, Math.round(parseFloat(pct(cat.count)) / 5));
    r.getCell(4).value = '\u2588'.repeat(barLen) + '\u2591'.repeat(20 - barLen);
    r.getCell(4).font = { size: 9, color: { argb: cat.color } };
    r.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(4).border = BORDER_ALL;
    r.getCell(4).alignment = { horizontal: 'center' };

    r.getCell(5).value = cat.strategy;
    r.getCell(5).font = { bold: true, size: 10, color: { argb: AWS_COLORS.teal } };
    r.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(5).border = BORDER_ALL;

    r.getCell(6).value = cat.effort;
    r.getCell(6).font = { bold: true, size: 10, color: { argb: _effortColor(cat.effort) } };
    r.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(6).border = BORDER_ALL;

    r.getCell(7).value = cat.timeline;
    r.getCell(7).font = { size: 10, italic: true, color: { argb: C.gray500 } };
    r.getCell(7).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(7).border = BORDER_ALL;

    // Example services
    const examples = (examplesByCategory[cat.label] || []).join(', ');
    r.getCell(8).value = examples || '\u2014';
    r.getCell(8).font = { size: 9, italic: true, color: { argb: C.gray500 } };
    r.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    r.getCell(8).alignment = { vertical: 'middle', wrapText: true };
    r.getCell(8).border = BORDER_ALL;

    rowNum++;
  });

  // Total row
  rowNum++;
  const totRow = ws.getRow(rowNum);
  totRow.height = 28;
  ['TOTAL', summary.total, '100%', '', '', '', '', ''].forEach((v, i) => {
    const c = totRow.getCell(i + 1);
    c.value = v;
    c.font = { bold: true, size: 11, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 24;
  ws.getColumn(5).width = 22;
  ws.getColumn(6).width = 12;
  ws.getColumn(7).width = 14;
  ws.getColumn(8).width = 40;
}

function _buildAwsServiceGroupSheet(wb, assessed) {
  const ws = wb.addWorksheet('By AWS Namespace', {
    properties: { tabColor: { argb: AWS_COLORS.orangeDark } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

  // Group by AWS namespace (e.g., ec2, s3, lambda)
  const groups = {};
  assessed.forEach(r => {
    const svc = (r['AWS SERVICE'] || '').toLowerCase();
    let namespace = 'other';
    if (svc.startsWith('aws::')) {
      const parts = svc.split('::');
      if (parts.length >= 2) namespace = parts[1];
    } else if (svc.includes('amazon') || svc.includes('aws')) {
      namespace = svc.replace(/amazon\s*/i, '').replace(/aws\s*/i, '').split(/\s/)[0].toLowerCase() || 'other';
    } else {
      namespace = svc.split(/[.\s:/]/)[0].toLowerCase() || 'other';
    }
    if (!groups[namespace]) groups[namespace] = { total: 0, 'Direct Equivalent': 0, 'Similar': 0, 'Partial': 0, 'No Direct Mapping': 0, services: [] };
    groups[namespace].total++;
    const sim = r['SIMILARITY'] || 'No Direct Mapping';
    if (groups[namespace][sim] !== undefined) groups[namespace][sim]++;
    else groups[namespace]['No Direct Mapping']++;
    if (groups[namespace].services.length < 5) {
      groups[namespace].services.push(r['AZURE EQUIVALENT'] || 'N/A');
    }
  });

  // Title
  ws.mergeCells('A1:I1');
  ws.getCell('A1').value = 'AWS Namespace Analysis';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: C.gray900 } };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:I2');
  ws.getCell('A2').value = `${Object.keys(groups).length} AWS namespaces detected  \u2022  Sorted by total service count`;
  ws.getCell('A2').font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  const headers = ['AWS Namespace', 'Total', 'Direct', 'Similar', 'Partial', 'No Map', 'Coverage', 'Readiness', 'Azure Services'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AWS_COLORS.orange } };
    c.alignment = { horizontal: i === 0 || i === 8 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  const sorted = Object.entries(groups).sort((a, b) => b[1].total - a[1].total);
  let rowNum = 4;

  sorted.forEach(([ns, counts], i) => {
    const coverageRate = counts.total > 0 ? (((counts['Direct Equivalent'] + counts['Similar']) / counts.total) * 100).toFixed(0) : '0';
    const cv = parseInt(coverageRate);
    const readinessLabel = cv >= 80 ? '\u2713 Ready' : cv >= 50 ? '\u25CB Partial' : '\u2717 Needs Work';
    const readinessColor = cv >= 80 ? C.success : cv >= 50 ? C.warning : C.danger;
    const uniqueAzure = [...new Set(counts.services.filter(s => s !== 'N/A' && s !== 'No Azure Equivalent Found'))].slice(0, 3).join(', ');

    const r = ws.getRow(rowNum);
    r.height = 24;

    const values = [ns.toUpperCase(), counts.total, counts['Direct Equivalent'], counts['Similar'], counts['Partial'], counts['No Direct Mapping'], `${coverageRate}%`, readinessLabel, uniqueAzure || '\u2014'];
    values.forEach((v, ci) => {
      const cell = r.getCell(ci + 1);
      cell.value = v;
      cell.alignment = { horizontal: ci === 0 || ci === 8 ? 'left' : 'center', vertical: 'middle' };
      cell.font = { size: 10, color: { argb: C.gray700 } };
      cell.border = BORDER_ALL;
      if (i % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
    });

    // Namespace name in bold
    r.getCell(1).font = { bold: true, size: 10, color: { argb: AWS_COLORS.orangeDark } };

    if (counts['Direct Equivalent'] > 0) r.getCell(3).font = { bold: true, size: 10, color: { argb: C.success } };
    if (counts['Similar'] > 0) r.getCell(4).font = { bold: true, size: 10, color: { argb: C.info } };
    if (counts['Partial'] > 0) r.getCell(5).font = { bold: true, size: 10, color: { argb: C.warning } };
    if (counts['No Direct Mapping'] > 0) r.getCell(6).font = { bold: true, size: 10, color: { argb: C.danger } };

    r.getCell(7).font = { bold: true, size: 10, color: { argb: cv >= 80 ? C.success : cv >= 50 ? C.warning : C.danger } };
    r.getCell(8).font = { bold: true, size: 10, color: { argb: readinessColor } };
    r.getCell(9).font = { size: 9, italic: true, color: { argb: C.gray500 } };
    rowNum++;
  });

  // Summary row
  rowNum++;
  const sumRow = ws.getRow(rowNum);
  sumRow.height = 26;
  const totals = sorted.reduce((acc, [, c]) => {
    acc.total += c.total; acc.direct += c['Direct Equivalent']; acc.similar += c['Similar'];
    acc.partial += c['Partial']; acc.noMap += c['No Direct Mapping']; return acc;
  }, { total: 0, direct: 0, similar: 0, partial: 0, noMap: 0 });
  const totalCov = totals.total > 0 ? (((totals.direct + totals.similar) / totals.total) * 100).toFixed(0) : '0';

  ['TOTAL', totals.total, totals.direct, totals.similar, totals.partial, totals.noMap, `${totalCov}%`, '', ''].forEach((v, i) => {
    const c = sumRow.getCell(i + 1);
    c.value = v;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray800 } };
    c.alignment = { horizontal: i === 0 ? 'left' : 'center', vertical: 'middle' };
    c.border = BORDER_ALL;
  });

  ws.getColumn(1).width = 20;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 10;
  ws.getColumn(4).width = 10;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 10;
  ws.getColumn(7).width = 12;
  ws.getColumn(8).width = 14;
  ws.getColumn(9).width = 40;

  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: 9 } };
}

function _buildAwsMigrationRoadmapSheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('Migration Roadmap', {
    properties: { tabColor: { argb: AWS_COLORS.teal } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

  // Title
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'Migration Roadmap \u2014 Phased Approach';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: C.gray900 } };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:H2');
  ws.getCell('A2').value = 'Services grouped by migration phase with recommended strategy and timeline';
  ws.getCell('A2').font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  const phases = [
    { num: 1, title: 'Phase 1: Quick Wins (Lift & Shift)', filter: 'Direct Equivalent', color: C.success, bg: C.successBg, desc: 'Services with direct Azure equivalents \u2014 minimal code changes required' },
    { num: 2, title: 'Phase 2: Replatform', filter: 'Similar', color: C.info, bg: C.infoBg, desc: 'Services with similar Azure counterparts \u2014 evaluate feature differences' },
    { num: 3, title: 'Phase 3: Refactor', filter: 'Partial', color: C.warning, bg: C.warningBg, desc: 'Partial mappings \u2014 architecture modifications needed' },
    { num: 4, title: 'Phase 4: Re-architect', filter: 'No Direct Mapping', color: C.danger, bg: C.dangerBg, desc: 'No direct equivalent \u2014 redesign using native Azure services' },
  ];

  // Header
  const headers = ['#', 'AWS Service', 'Azure Equivalent', 'Category', 'Strategy', 'Effort', 'Est. Timeline', 'Migration Notes'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AWS_COLORS.teal } };
    c.alignment = { horizontal: i === 0 ? 'center' : 'left', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  let rowNum = 4;
  let itemNum = 1;

  phases.forEach(phase => {
    const items = assessed.filter(r => r['SIMILARITY'] === phase.filter);
    if (items.length === 0) return;

    // Phase header row
    const phRow = ws.getRow(rowNum);
    phRow.height = 30;
    ws.mergeCells(rowNum, 1, rowNum, 8);
    const phCell = phRow.getCell(1);
    phCell.value = `${phase.title}  \u2022  ${items.length} service${items.length !== 1 ? 's' : ''}  \u2022  ${phase.desc}`;
    phCell.font = { bold: true, size: 11, color: { argb: phase.color } };
    phCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: phase.bg } };
    phCell.alignment = { vertical: 'middle' };
    phCell.border = BORDER_ALL;
    rowNum++;

    items.forEach((r, idx) => {
      const ms = _migrationStrategy(r['SIMILARITY']);
      const dr = ws.getRow(rowNum);
      dr.height = 24;
      const isAlt = idx % 2 === 1;

      const values = [
        itemNum++,
        r['AWS SERVICE'] || '',
        r['AZURE EQUIVALENT'] || '',
        r['CATEGORY'] || '',
        ms.strategy,
        ms.effort,
        ms.weeks,
        r['MIGRATION NOTES'] || '\u2014'
      ];

      values.forEach((v, ci) => {
        const cell = dr.getCell(ci + 1);
        cell.value = v;
        cell.font = { size: 10, color: { argb: C.gray700 } };
        cell.alignment = { vertical: 'middle', wrapText: ci === 7, horizontal: ci === 0 ? 'center' : 'left' };
        cell.border = BORDER_ALL;
        if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
      });

      // Effort coloring
      dr.getCell(6).font = { bold: true, size: 10, color: { argb: _effortColor(ms.effort) } };
      dr.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

      // Timeline
      dr.getCell(7).font = { size: 10, italic: true, color: { argb: C.gray500 } };
      dr.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };

      rowNum++;
    });

    // Spacing between phases
    ws.getRow(rowNum).height = 8;
    rowNum++;
  });

  // Footer summary
  rowNum++;
  ws.mergeCells(rowNum, 1, rowNum, 8);
  const footNote = ws.getCell(rowNum, 1);
  footNote.value = `Total: ${summary.total} services  \u2022  Estimated total timeline depends on team size and parallelization  \u2022  Phase 1 services can start immediately`;
  footNote.font = { size: 10, italic: true, color: { argb: C.gray500 } };
  footNote.alignment = { vertical: 'middle' };

  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 28;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 22;
  ws.getColumn(6).width = 12;
  ws.getColumn(7).width = 14;
  ws.getColumn(8).width = 50;
}

function _buildAwsRiskMatrixSheet(wb, assessed) {
  const ws = wb.addWorksheet('Risk Matrix', {
    properties: { tabColor: { argb: C.danger } },
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }]
  });

  // Title
  ws.mergeCells('A1:G1');
  ws.getCell('A1').value = 'Migration Risk Assessment';
  ws.getCell('A1').font = { bold: true, size: 16, color: { argb: C.gray900 } };
  ws.getRow(1).height = 34;

  ws.mergeCells('A2:G2');
  ws.getCell('A2').value = 'Services ranked by migration risk level \u2014 Critical and High risk items need immediate attention';
  ws.getCell('A2').font = { size: 10, color: { argb: C.gray500 } };
  ws.getRow(2).height = 20;

  const headers = ['Priority', 'AWS Service', 'Azure Target', 'Risk Level', 'Impact', 'Strategy', 'Action Required'];
  const hdrRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const c = hdrRow.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, size: 10, color: { argb: C.white } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.danger } };
    c.alignment = { horizontal: i === 0 || i === 3 ? 'center' : 'left', vertical: 'middle' };
    c.border = BORDER_ALL;
  });
  hdrRow.height = 28;

  // Rank all services by risk (highest risk first)
  const riskOrder = { 'No Direct Mapping': 1, 'Partial': 2, 'Similar': 3, 'Direct Equivalent': 4 };
  const sortedByRisk = [...assessed].sort((a, b) => {
    const aRisk = riskOrder[a['SIMILARITY']] || 5;
    const bRisk = riskOrder[b['SIMILARITY']] || 5;
    return aRisk - bRisk;
  });

  let rowNum = 4;
  sortedByRisk.forEach((r, idx) => {
    const ms = _migrationStrategy(r['SIMILARITY']);
    const dr = ws.getRow(rowNum);
    dr.height = 26;
    const isAlt = idx % 2 === 1;

    // Impact based on category
    const category = r['CATEGORY'] || '';
    const impact = category.toLowerCase().includes('compute') || category.toLowerCase().includes('database')
      ? 'High' : category.toLowerCase().includes('storage') || category.toLowerCase().includes('network')
      ? 'Medium' : 'Low';

    // Action required
    let action = '';
    switch (r['SIMILARITY']) {
      case 'Direct Equivalent': action = 'Validate config compatibility'; break;
      case 'Similar': action = 'Feature parity assessment needed'; break;
      case 'Partial': action = 'Architecture review + POC required'; break;
      case 'No Direct Mapping': action = 'Deep-dive research + alternative design'; break;
      default: action = 'Review required';
    }

    const values = [
      `P${idx + 1}`,
      r['AWS SERVICE'] || '',
      r['AZURE EQUIVALENT'] || '',
      ms.risk,
      impact,
      ms.strategy,
      action
    ];

    values.forEach((v, ci) => {
      const cell = dr.getCell(ci + 1);
      cell.value = v;
      cell.font = { size: 10, color: { argb: C.gray700 } };
      cell.alignment = { vertical: 'middle', horizontal: ci === 0 || ci === 3 ? 'center' : 'left' };
      cell.border = BORDER_ALL;
      if (isAlt) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C.gray50 } };
    });

    // Priority badge
    dr.getCell(1).font = { bold: true, size: 9, color: { argb: C.white } };
    dr.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: _riskColor(ms.risk) } };

    // Risk level coloring
    dr.getCell(4).font = { bold: true, size: 10, color: { argb: _riskColor(ms.risk) } };

    // Impact coloring
    const impactColor = impact === 'High' ? C.danger : impact === 'Medium' ? C.warning : C.success;
    dr.getCell(5).font = { bold: true, size: 10, color: { argb: impactColor } };
    dr.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    rowNum++;
  });

  // Risk summary at bottom
  rowNum += 2;
  ws.mergeCells(rowNum, 1, rowNum, 7);
  const sumCell = ws.getCell(rowNum, 1);
  const critical = assessed.filter(r => r['SIMILARITY'] === 'No Direct Mapping').length;
  const high = assessed.filter(r => r['SIMILARITY'] === 'Partial').length;
  const medium = assessed.filter(r => r['SIMILARITY'] === 'Similar').length;
  const low = assessed.filter(r => r['SIMILARITY'] === 'Direct Equivalent').length;
  sumCell.value = `Risk Summary:  Critical: ${critical}  |  High: ${high}  |  Medium: ${medium}  |  Low: ${low}`;
  sumCell.font = { bold: true, size: 11, color: { argb: C.gray700 } };
  sumCell.alignment = { vertical: 'middle' };

  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 28;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 22;
  ws.getColumn(7).width = 42;

  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 2, column: 7 } };
}

function _awsSimilarityColors(val) {
  switch (val) {
    case 'Direct Equivalent': return { bg: C.successBg, fg: C.success };
    case 'Similar':           return { bg: C.infoBg, fg: C.info };
    case 'Partial':           return { bg: C.warningBg, fg: C.warning };
    case 'No Direct Mapping': return { bg: C.dangerBg, fg: C.danger };
    default:                  return { bg: C.warningBg, fg: C.warning };
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// GCP-TO-AZURE REPORT (similar structure to AWS report)
// ══════════════════════════════════════════════════════════════════════════════
const GCP_COLORS = {
  blue: '4285F4',
  blueDark: '1A73E8',
  blueLight: 'E8F0FE',
  red: 'EA4335',
  redLight: 'FCE8E6',
  green: '34A853',
  greenLight: 'E6F4EA',
  yellow: 'FBBC04',
  yellowLight: 'FEF7E0',
  azure: '0078D4',
  azureLight: 'E8F4FD',
  teal: '00897B',
  tealLight: 'E0F2F1',
  indigo: '3F51B5',
  indigoLight: 'E8EAF6'
};

async function _buildGcpReport({ assessed, summary, sheetName, outputPath }) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Resource Migration Assessment Tool';
  wb.created = new Date();

  _buildGcpCoverSheet(wb, assessed, summary);
  _buildGcpDataSheet(wb, assessed, sheetName || 'Service Mapping');
  _buildGcpCategorySheet(wb, assessed, summary);
  _buildGcpServiceGroupSheet(wb, assessed);
  _buildGcpMigrationRoadmapSheet(wb, assessed, summary);
  _buildGcpRiskMatrixSheet(wb, assessed);

  await wb.xlsx.writeFile(outputPath);
}

function _buildGcpCoverSheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('Executive Summary', {
    properties: { tabColor: { argb: GCP_COLORS.blue } },
    views: [{ showGridLines: false }]
  });
  const dateStr = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'long', timeStyle: 'short' });
  const version = '1.2.1';
  const pct = n => summary.total > 0 ? ((n / summary.total) * 100).toFixed(1) : '0.0';
  const coverage = summary.total > 0 ? (((summary.directEquivalent + summary.similar) / summary.total) * 100).toFixed(1) : '0.0';
  const coverageVal = parseFloat(coverage);
  const coverageColor = coverageVal >= 80 ? C.success : coverageVal >= 50 ? C.warning : C.danger;

  const readiness = coverageVal >= 80 ? { label: 'READY TO MIGRATE', color: C.success, bg: C.successBg }
    : coverageVal >= 50 ? { label: 'NEEDS PLANNING', color: C.warning, bg: C.warningBg }
    : { label: 'COMPLEX MIGRATION', color: C.danger, bg: C.dangerBg };

  ws.getColumn(1).width = 3;
  ws.getColumn(2).width = 5;
  ws.getColumn(3).width = 22;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 18;
  ws.getColumn(7).width = 18;
  ws.getColumn(8).width = 18;
  ws.getColumn(9).width = 5;
  ws.getColumn(10).width = 3;

  let row = 1;
  ws.getRow(row).height = 15; row++;

  // Brand Header
  ws.mergeCells(row, 2, row, 9);
  const h1 = ws.getCell(row, 2);
  h1.value = 'GCP → Azure Migration Assessment';
  h1.font = { name: 'Segoe UI', size: 22, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  h1.alignment = { vertical: 'middle' };
  ws.getRow(row).height = 36; row++;

  ws.mergeCells(row, 2, row, 9);
  const sub = ws.getCell(row, 2);
  sub.value = `Generated: ${dateStr}  |  Tool v${version}  |  Source: Microsoft Learn GCP Professional`;
  sub.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray500 } };
  row += 2;

  // Readiness badge
  ws.mergeCells(row, 2, row, 4);
  const badge = ws.getCell(row, 2);
  badge.value = `Migration Readiness: ${readiness.label}`;
  badge.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FF' + readiness.color } };
  badge.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + readiness.bg } };
  badge.border = BORDER_ALL;
  badge.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(row).height = 28; row += 2;

  // KPI cards
  const kpis = [
    { label: 'Total Services', value: summary.total, color: GCP_COLORS.blue },
    { label: 'Direct Equivalent', value: `${summary.directEquivalent} (${pct(summary.directEquivalent)}%)`, color: C.success },
    { label: 'Similar', value: `${summary.similar} (${pct(summary.similar)}%)`, color: C.info },
    { label: 'Partial', value: `${summary.partial} (${pct(summary.partial)}%)`, color: C.warning },
    { label: 'No Mapping', value: `${summary.noMapping} (${pct(summary.noMapping)}%)`, color: C.danger },
  ];
  for (let i = 0; i < kpis.length; i++) {
    const col = 2 + i * 2;
    ws.mergeCells(row, col, row, col + 1);
    const lbl = ws.getCell(row, col);
    lbl.value = kpis[i].label;
    lbl.font = { name: 'Segoe UI', size: 9, color: { argb: 'FF' + C.gray500 } };
    lbl.alignment = { horizontal: 'center' };

    ws.mergeCells(row + 1, col, row + 1, col + 1);
    const val = ws.getCell(row + 1, col);
    val.value = kpis[i].value;
    val.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + kpis[i].color } };
    val.alignment = { horizontal: 'center' };
  }
  row += 3;

  // Coverage bar (text)
  ws.mergeCells(row, 2, row, 9);
  const covLabel = ws.getCell(row, 2);
  covLabel.value = `Azure Coverage: ${coverage}% of GCP services have Direct Equivalent or Similar mapping`;
  covLabel.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF' + coverageColor } };
  row += 2;

  // Migration effort breakdown
  ws.mergeCells(row, 2, row, 5);
  const effortTitle = ws.getCell(row, 2);
  effortTitle.value = 'Estimated Migration Effort';
  effortTitle.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FF' + C.gray800 } };
  row++;
  const efforts = [
    { phase: 'Lift & Shift (Direct)', count: summary.directEquivalent, effort: 'Low', color: C.success },
    { phase: 'Refactor (Similar)', count: summary.similar, effort: 'Medium', color: C.info },
    { phase: 'Re-architect (Partial)', count: summary.partial, effort: 'High', color: C.warning },
    { phase: 'Replace / Custom (None)', count: summary.noMapping, effort: 'Very High', color: C.danger },
  ];
  for (const e of efforts) {
    ws.getCell(row, 3).value = e.phase;
    ws.getCell(row, 3).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
    ws.getCell(row, 5).value = `${e.count} services`;
    ws.getCell(row, 5).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + e.color } };
    ws.getCell(row, 6).value = `Effort: ${e.effort}`;
    ws.getCell(row, 6).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + e.color } };
    row++;
  }
  row += 2;

  // Category breakdown
  ws.mergeCells(row, 2, row, 5);
  const catTitle = ws.getCell(row, 2);
  catTitle.value = 'Services by Category';
  catTitle.font = { name: 'Segoe UI', size: 13, bold: true, color: { argb: 'FF' + C.gray800 } };
  row++;
  const sortedCats = Object.entries(summary.categories).sort((a, b) => b[1] - a[1]);
  for (const [cat, count] of sortedCats) {
    ws.getCell(row, 3).value = cat;
    ws.getCell(row, 3).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
    ws.getCell(row, 5).value = count;
    ws.getCell(row, 5).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
    row++;
  }
}

function _buildGcpDataSheet(wb, assessed, sheetName) {
  const ws = wb.addWorksheet(sheetName || 'Service Mapping', {
    properties: { tabColor: { argb: GCP_COLORS.azure } }
  });

  const headers = ['GCP Service', 'Name', 'Azure Equivalent', 'Azure Resource Type', 'Category', 'Similarity', 'SKU Recommendation', 'Migration Notes'];
  const widths = [28, 24, 30, 32, 16, 18, 22, 50];

  // Title row
  ws.mergeCells(1, 1, 1, headers.length);
  const title = ws.getCell(1, 1);
  title.value = 'GCP → Azure Service Mapping Details';
  title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  ws.getRow(1).height = 28;

  // Header row
  const headerRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GCP_COLORS.blueDark } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
    ws.getColumn(i + 1).width = widths[i];
  });
  ws.getRow(3).height = 24;

  // Data rows
  let rowNum = 4;
  for (const item of assessed) {
    const r = ws.getRow(rowNum);
    r.getCell(1).value = item['GCP SERVICE'] || '';
    r.getCell(2).value = item['NAME'] || '';
    r.getCell(3).value = item['AZURE EQUIVALENT'] || '';
    r.getCell(4).value = item['AZURE RESOURCE TYPE'] || '';
    r.getCell(5).value = item['CATEGORY'] || '';
    r.getCell(6).value = item['SIMILARITY'] || '';
    r.getCell(7).value = item['SKU RECOMMENDATION'] || '';
    r.getCell(8).value = item['MIGRATION NOTES'] || '';

    // Color-code similarity
    const simColors = _gcpSimilarityColors(item['SIMILARITY']);
    r.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + simColors.bg } };
    r.getCell(6).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + simColors.fg } };

    for (let c = 1; c <= headers.length; c++) {
      r.getCell(c).border = BORDER_ALL;
      if (c !== 6) r.getCell(c).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
      r.getCell(c).alignment = { vertical: 'middle', wrapText: c === 8 };
    }

    // Zebra striping
    if (rowNum % 2 === 0) {
      for (let c = 1; c <= headers.length; c++) {
        if (c !== 6) r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.gray50 } };
      }
    }
    rowNum++;
  }

  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: headers.length } };
}

function _buildGcpCategorySheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('By Category', {
    properties: { tabColor: { argb: GCP_COLORS.green } }
  });

  ws.mergeCells(1, 1, 1, 6);
  const title = ws.getCell(1, 1);
  title.value = 'GCP Services Grouped by Category';
  title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  ws.getRow(1).height = 28;

  const headers = ['Category', 'Count', 'Direct Equivalent', 'Similar', 'Partial', 'No Mapping'];
  const headerRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GCP_COLORS.green } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  ws.getColumn(1).width = 24;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 18;
  ws.getColumn(4).width = 12;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 14;

  // Group by category
  const catStats = {};
  for (const item of assessed) {
    const cat = item['CATEGORY'] || 'Unknown';
    if (!catStats[cat]) catStats[cat] = { total: 0, direct: 0, similar: 0, partial: 0, none: 0 };
    catStats[cat].total++;
    const sim = (item['SIMILARITY'] || '').toLowerCase();
    if (sim === 'direct equivalent') catStats[cat].direct++;
    else if (sim === 'similar') catStats[cat].similar++;
    else if (sim === 'partial') catStats[cat].partial++;
    else catStats[cat].none++;
  }

  let rowNum = 4;
  const sortedCats = Object.entries(catStats).sort((a, b) => b[1].total - a[1].total);
  for (const [cat, stats] of sortedCats) {
    const r = ws.getRow(rowNum);
    r.getCell(1).value = cat;
    r.getCell(2).value = stats.total;
    r.getCell(3).value = stats.direct;
    r.getCell(4).value = stats.similar;
    r.getCell(5).value = stats.partial;
    r.getCell(6).value = stats.none;
    for (let c = 1; c <= 6; c++) {
      r.getCell(c).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
      r.getCell(c).border = BORDER_ALL;
      r.getCell(c).alignment = { horizontal: c === 1 ? 'left' : 'center', vertical: 'middle' };
    }
    if (rowNum % 2 === 0) {
      for (let c = 1; c <= 6; c++) r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.gray50 } };
    }
    rowNum++;
  }
}

function _buildGcpServiceGroupSheet(wb, assessed) {
  const ws = wb.addWorksheet('By GCP Service', {
    properties: { tabColor: { argb: GCP_COLORS.yellow } }
  });

  ws.mergeCells(1, 1, 1, 7);
  const title = ws.getCell(1, 1);
  title.value = 'GCP Services by Source Namespace';
  title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  ws.getRow(1).height = 28;

  // Group by GCP namespace
  const groups = {};
  for (const item of assessed) {
    const svc = item['GCP SERVICE'] || 'Unknown';
    if (!groups[svc]) groups[svc] = { items: [], azure: new Set() };
    groups[svc].items.push(item);
    if (item['AZURE EQUIVALENT'] && item['AZURE EQUIVALENT'] !== 'No Azure Equivalent Found') {
      groups[svc].azure.add(item['AZURE EQUIVALENT']);
    }
  }

  const headers = ['GCP Service', 'Count', 'Azure Mappings', 'Categories', 'Direct', 'Similar/Partial', 'No Mapping'];
  const headerRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GCP_COLORS.blueDark } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });
  ws.getColumn(1).width = 28;
  ws.getColumn(2).width = 10;
  ws.getColumn(3).width = 36;
  ws.getColumn(4).width = 22;
  ws.getColumn(5).width = 10;
  ws.getColumn(6).width = 16;
  ws.getColumn(7).width = 14;

  let rowNum = 4;
  const sortedGroups = Object.entries(groups).sort((a, b) => b[1].items.length - a[1].items.length);
  for (const [svc, data] of sortedGroups) {
    const r = ws.getRow(rowNum);
    const cats = [...new Set(data.items.map(i => i['CATEGORY']))].join(', ');
    let direct = 0, similarPartial = 0, none = 0;
    for (const it of data.items) {
      const s = (it['SIMILARITY'] || '').toLowerCase();
      if (s === 'direct equivalent') direct++;
      else if (s === 'similar' || s === 'partial') similarPartial++;
      else none++;
    }

    r.getCell(1).value = svc;
    r.getCell(2).value = data.items.length;
    r.getCell(3).value = [...data.azure].join(', ') || 'None';
    r.getCell(4).value = cats;
    r.getCell(5).value = direct;
    r.getCell(6).value = similarPartial;
    r.getCell(7).value = none;

    // Readiness indicator
    const readinessPct = data.items.length > 0 ? ((direct + similarPartial) / data.items.length * 100) : 0;
    const readColor = readinessPct >= 80 ? C.success : readinessPct >= 50 ? C.warning : C.danger;
    r.getCell(1).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + readColor } };

    for (let c = 2; c <= 7; c++) {
      r.getCell(c).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
      r.getCell(c).border = BORDER_ALL;
      r.getCell(c).alignment = { horizontal: c <= 4 ? 'left' : 'center', vertical: 'middle', wrapText: c === 3 };
    }
    r.getCell(1).border = BORDER_ALL;
    if (rowNum % 2 === 0) {
      for (let c = 1; c <= 7; c++) r.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.gray50 } };
    }
    rowNum++;
  }

  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: 7 } };
}

function _buildGcpMigrationRoadmapSheet(wb, assessed, summary) {
  const ws = wb.addWorksheet('Migration Roadmap', {
    properties: { tabColor: { argb: GCP_COLORS.teal } }
  });

  ws.mergeCells(1, 1, 1, 7);
  const title = ws.getCell(1, 1);
  title.value = 'Migration Roadmap — Phased Approach';
  title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  ws.getRow(1).height = 28;

  const phases = [
    { name: 'Phase 1: Lift & Shift', filter: 'direct equivalent', effort: 'Low', timeline: '1-3 months', color: C.success },
    { name: 'Phase 2: Refactor', filter: 'similar', effort: 'Medium', timeline: '3-6 months', color: C.info },
    { name: 'Phase 3: Re-architect', filter: 'partial', effort: 'High', timeline: '6-12 months', color: C.warning },
    { name: 'Phase 4: Replace / Custom', filter: 'no direct mapping', effort: 'Very High', timeline: '6-18 months', color: C.danger },
  ];

  let rowNum = 3;
  for (const phase of phases) {
    const items = assessed.filter(i => (i['SIMILARITY'] || '').toLowerCase() === phase.filter);
    if (items.length === 0) continue;

    // Phase header
    ws.mergeCells(rowNum, 1, rowNum, 7);
    const phCell = ws.getCell(rowNum, 1);
    phCell.value = `${phase.name}  |  ${items.length} services  |  Effort: ${phase.effort}  |  Timeline: ${phase.timeline}`;
    phCell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF' + C.white } };
    phCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + phase.color } };
    phCell.alignment = { vertical: 'middle' };
    ws.getRow(rowNum).height = 24;
    rowNum++;

    // Column headers
    const colHeaders = ['GCP Service', 'Azure Equivalent', 'Category', 'SKU', 'Strategy', 'Risk', 'Notes'];
    const hRow = ws.getRow(rowNum);
    colHeaders.forEach((h, i) => {
      const cell = hRow.getCell(i + 1);
      cell.value = h;
      cell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF' + C.gray700 } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + C.gray100 } };
      cell.border = BORDER_ALL;
    });
    rowNum++;

    for (const item of items) {
      const r = ws.getRow(rowNum);
      r.getCell(1).value = item['GCP SERVICE'] || '';
      r.getCell(2).value = item['AZURE EQUIVALENT'] || '';
      r.getCell(3).value = item['CATEGORY'] || '';
      r.getCell(4).value = item['SKU RECOMMENDATION'] || '';
      r.getCell(5).value = _migrationStrategy(item['SIMILARITY']);
      r.getCell(6).value = phase.effort;
      r.getCell(7).value = item['MIGRATION NOTES'] || '';

      const efColor = _effortColor(phase.effort);
      r.getCell(6).font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF' + efColor } };

      for (let c = 1; c <= 7; c++) {
        if (c !== 6) r.getCell(c).font = { name: 'Segoe UI', size: 9, color: { argb: 'FF' + C.gray700 } };
        r.getCell(c).border = BORDER_ALL;
        r.getCell(c).alignment = { vertical: 'middle', wrapText: c === 7 };
      }
      rowNum++;
    }
    rowNum++; // gap between phases
  }

  ws.getColumn(1).width = 26;
  ws.getColumn(2).width = 30;
  ws.getColumn(3).width = 16;
  ws.getColumn(4).width = 20;
  ws.getColumn(5).width = 18;
  ws.getColumn(6).width = 12;
  ws.getColumn(7).width = 44;
}

function _buildGcpRiskMatrixSheet(wb, assessed) {
  const ws = wb.addWorksheet('Risk Matrix', {
    properties: { tabColor: { argb: GCP_COLORS.red } }
  });

  ws.mergeCells(1, 1, 1, 7);
  const title = ws.getCell(1, 1);
  title.value = 'Migration Risk Assessment';
  title.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + GCP_COLORS.blueDark } };
  ws.getRow(1).height = 28;

  // Sort by risk: No Direct Mapping first, then Partial, Similar, Direct
  const riskOrder = { 'no direct mapping': 0, 'partial': 1, 'similar': 2, 'direct equivalent': 3 };
  const sorted = [...assessed].sort((a, b) => {
    const aRisk = riskOrder[(a['SIMILARITY'] || '').toLowerCase()] ?? 1;
    const bRisk = riskOrder[(b['SIMILARITY'] || '').toLowerCase()] ?? 1;
    return aRisk - bRisk;
  });

  const headers = ['Priority', 'GCP Service', 'Azure Equivalent', 'Similarity', 'Risk Level', 'Category', 'Mitigation Notes'];
  const headerRow = ws.getRow(3);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + C.white } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GCP_COLORS.red } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = BORDER_ALL;
  });

  let rowNum = 4;
  for (let idx = 0; idx < sorted.length; idx++) {
    const item = sorted[idx];
    const r = ws.getRow(rowNum);
    const sim = (item['SIMILARITY'] || '').toLowerCase();
    const risk = sim === 'no direct mapping' ? 'Critical' : sim === 'partial' ? 'High' : sim === 'similar' ? 'Medium' : 'Low';
    const riskColor = _riskColor(risk);

    r.getCell(1).value = idx + 1;
    r.getCell(2).value = item['GCP SERVICE'] || '';
    r.getCell(3).value = item['AZURE EQUIVALENT'] || '';
    r.getCell(4).value = item['SIMILARITY'] || '';
    r.getCell(5).value = risk;
    r.getCell(6).value = item['CATEGORY'] || '';
    r.getCell(7).value = item['MIGRATION NOTES'] || '';

    r.getCell(5).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + riskColor } };
    const simColors = _gcpSimilarityColors(item['SIMILARITY']);
    r.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + simColors.bg } };
    r.getCell(4).font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF' + simColors.fg } };

    for (let c = 1; c <= 7; c++) {
      if (c !== 4 && c !== 5) r.getCell(c).font = { name: 'Segoe UI', size: 10, color: { argb: 'FF' + C.gray700 } };
      r.getCell(c).border = BORDER_ALL;
      r.getCell(c).alignment = { vertical: 'middle', wrapText: c === 7 };
    }
    rowNum++;
  }

  ws.getColumn(1).width = 8;
  ws.getColumn(2).width = 28;
  ws.getColumn(3).width = 30;
  ws.getColumn(4).width = 18;
  ws.getColumn(5).width = 12;
  ws.getColumn(6).width = 16;
  ws.getColumn(7).width = 44;

  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: rowNum - 1, column: 7 } };
}

function _gcpSimilarityColors(val) {
  switch (val) {
    case 'Direct Equivalent': return { bg: C.successBg, fg: C.success };
    case 'Similar':           return { bg: C.infoBg, fg: C.info };
    case 'Partial':           return { bg: C.warningBg, fg: C.warning };
    case 'No Direct Mapping': return { bg: C.dangerBg, fg: C.danger };
    default:                  return { bg: C.warningBg, fg: C.warning };
  }
}

module.exports = { buildReport };
