const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = require('docx');
const fs = require('fs');

const md = fs.readFileSync('DOCUMENTATION.docx.md', 'utf-8');
const lines = md.split('\n');
const children = [];

function makePara(text, opts = {}) {
  const runs = [];
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  for (const p of parts) {
    if (p.startsWith('**') && p.endsWith('**')) {
      runs.push(new TextRun({ text: p.slice(2, -2), bold: true, font: 'Calibri', size: opts.size || 22 }));
    } else if (p) {
      runs.push(new TextRun({ text: p, font: 'Calibri', size: opts.size || 22 }));
    }
  }
  return new Paragraph({ children: runs, heading: opts.heading, spacing: { after: 120 } });
}

let inCodeBlock = false;
let codeLines = [];
let tableRows = [];

for (let i = 0; i < lines.length; i++) {
  const line = lines[i];

  if (line.startsWith('```')) {
    if (inCodeBlock) {
      for (const cl of codeLines) {
        children.push(new Paragraph({ children: [new TextRun({ text: cl, font: 'Consolas', size: 18 })], spacing: { after: 40 }, indent: { left: 400 } }));
      }
      codeLines = [];
      inCodeBlock = false;
    } else {
      inCodeBlock = true;
    }
    continue;
  }
  if (inCodeBlock) { codeLines.push(line); continue; }

  if (line.startsWith('|') && line.includes('|')) {
    const cells = line.split('|').slice(1, -1).map(c => c.trim());
    if (cells.every(c => /^[-:]+$/.test(c))) continue;
    tableRows.push(cells);
    if (i + 1 >= lines.length || !lines[i + 1].startsWith('|')) {
      const tRows = tableRows.map((row, ri) => new TableRow({
        children: row.map(cell => new TableCell({
          children: [new Paragraph({ children: [new TextRun({ text: cell, bold: ri === 0, font: 'Calibri', size: 20 })], spacing: { after: 60 } })],
          width: { size: Math.floor(9000 / row.length), type: WidthType.DXA }
        }))
      }));
      children.push(new Table({ rows: tRows, width: { size: 9000, type: WidthType.DXA } }));
      children.push(new Paragraph({ spacing: { after: 120 } }));
      tableRows = [];
    }
    continue;
  }

  if (line.startsWith('# ')) { children.push(makePara(line.slice(2), { heading: HeadingLevel.HEADING_1 })); continue; }
  if (line.startsWith('## ')) { children.push(makePara(line.slice(3), { heading: HeadingLevel.HEADING_2 })); continue; }
  if (line.startsWith('### ')) { children.push(makePara(line.slice(4), { heading: HeadingLevel.HEADING_3 })); continue; }
  if (line.startsWith('---')) { children.push(new Paragraph({ spacing: { after: 200 } })); continue; }

  if (line.startsWith('- ') || line.startsWith('* ')) {
    children.push(new Paragraph({ children: [new TextRun({ text: '• ' + line.slice(2), font: 'Calibri', size: 22 })], spacing: { after: 80 }, indent: { left: 400 } }));
    continue;
  }

  const numMatch = line.match(/^(\d+)\.\s(.+)/);
  if (numMatch) {
    children.push(new Paragraph({ children: [new TextRun({ text: numMatch[1] + '. ' + numMatch[2], font: 'Calibri', size: 22 })], spacing: { after: 80 }, indent: { left: 400 } }));
    continue;
  }

  if (line.startsWith('> ')) {
    children.push(new Paragraph({ children: [new TextRun({ text: line.slice(2), italics: true, font: 'Calibri', size: 22 })], spacing: { after: 80 }, indent: { left: 400 } }));
    continue;
  }

  if (line.trim() === '') { children.push(new Paragraph({ spacing: { after: 80 } })); continue; }

  children.push(makePara(line));
}

const doc = new Document({
  sections: [{ children }],
  creator: 'Azure CSP Migration Tool',
  title: 'Azure CSP Migration Assessment Tool - Documentation'
});

Packer.toBuffer(doc).then(buf => {
  const desktopPath = 'D:/OneDrive - Cloud 9 Infosystems, Inc/Desktop/AzureCSP-Migration-Documentation.docx';
  fs.writeFileSync(desktopPath, buf);
  console.log('Word doc saved to Desktop: ' + desktopPath);

  fs.writeFileSync('DOCUMENTATION.docx', buf);
  console.log('Also saved: DOCUMENTATION.docx in project root');
}).catch(err => {
  console.error('Error:', err.message);
  process.exit(1);
});
