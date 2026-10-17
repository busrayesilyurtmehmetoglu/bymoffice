// The sample files of the demo site, made with the engine's public API from synthetic content:
// a workbook (values, formulas, number formats, a chart) and a document (a title, headings, lists,
// a table), each in Turkish and in English. Each file is read back and checked before it is kept.

const TEXT = {
  tr: {
    sheet: 'Satışlar',
    date: 'dd.mm.yyyy',
    header: ['Ay', 'Tarih', 'Gelir', 'Gider', 'Kâr', 'Kâr oranı'],
    months: ['Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran'],
    total: 'Toplam',
    chart: 'Gelir ve gider',
    trend: 'Kâr',
    title: 'Örnek belge',
    intro: 'Bu belge BYM Office ile tarayıcıda üretildi. Metni değiştirebilir, biçimlendirebilir ve docx olarak indirebilirsiniz.',
    listHeading: 'Neler var',
    bullets: ['Başlıklar ve paragraf stilleri', 'Madde işaretli ve numaralı listeler', 'Hücreleri olan bir tablo'],
    tableHeading: 'Tablo',
    table: [['Ürün', 'Adet', 'Fiyat'], ['Kalem', '12', '4,50'], ['Defter', '5', '18,00']],
    stepsHeading: 'Adımlar',
    steps: ['Bir paragrafa tıklayın ve yazın', 'Araç çubuğundan biçim verin', 'Dosyayı indirin'],
    end: 'Belgenin sonu.',
  },
  en: {
    sheet: 'Sales',
    date: 'mm/dd/yyyy',
    header: ['Month', 'Date', 'Income', 'Expenses', 'Profit', 'Margin'],
    months: ['January', 'February', 'March', 'April', 'May', 'June'],
    total: 'Total',
    chart: 'Income and expenses',
    trend: 'Profit',
    title: 'Sample document',
    intro: 'This document was made in the browser with BYM Office. Change the text, format it and download it as a docx.',
    listHeading: 'What is in it',
    bullets: ['Headings and paragraph styles', 'Bulleted and numbered lists', 'A table with cells'],
    tableHeading: 'Table',
    table: [['Item', 'Quantity', 'Price'], ['Pen', '12', '4.50'], ['Notebook', '5', '18.00']],
    stepsHeading: 'Steps',
    steps: ['Click a paragraph and type', 'Format it from the toolbar', 'Download the file'],
    end: 'The end of the document.',
  },
};

const INCOME = [1200, 1500, 1350, 1800, 1650, 2100];
const EXPENSES = [800, 950, 700, 1100, 900, 1250];

function check(condition, message) {
  if (!condition) throw new Error(`site sample: ${message}`);
}

async function sheetSample(B, lang) {
  const t = TEXT[lang];
  const file = await B.open(await B.createBlank('xlsx', { locale: lang, sheetName: t.sheet }));
  const workbook = await file.workbook();
  const apply = (ops) => {
    for (const op of ops) B.applyOp(workbook, op);
  };
  const put = (row, col, text) => apply(B.setCellInput(workbook, 0, row, col, text));
  t.header.forEach((text, col) => put(0, col, text));
  const last = INCOME.length; // the row of the last month (0-based: the header is row 0)
  for (let i = 0; i < INCOME.length; i++) {
    const row = i + 1;
    const r = row + 1; // as the formulas name it
    put(row, 0, t.months[i]);
    put(row, 1, `2026-${String(i + 1).padStart(2, '0')}-01`);
    put(row, 2, String(INCOME[i]));
    put(row, 3, String(EXPENSES[i]));
    put(row, 4, `=C${r}-D${r}`);
    put(row, 5, `=IF(C${r}=0,0,E${r}/C${r})`);
  }
  const total = last + 1;
  put(total, 0, t.total);
  for (const [col, letter] of [[2, 'C'], [3, 'D'], [4, 'E']]) put(total, col, `=SUM(${letter}2:${letter}${last + 1})`);
  put(total, 5, `=IF(C${total + 1}=0,0,E${total + 1}/C${total + 1})`);

  apply(B.formatRange(workbook, 0, { top: 0, left: 0, bottom: 0, right: 5 }, { bold: true, fill: { rgb: 'FFDCE6F2' }, border: { preset: 'bottom' } }));
  apply(B.formatRange(workbook, 0, { top: total, left: 0, bottom: total, right: 5 }, { bold: true, border: { preset: 'top' } }));
  apply(B.formatRange(workbook, 0, { top: 1, left: 1, bottom: last, right: 1 }, { numFmt: t.date }));
  apply(B.formatRange(workbook, 0, { top: 1, left: 2, bottom: total, right: 4 }, { numFmt: '#,##0.00' }));
  apply(B.formatRange(workbook, 0, { top: 1, left: 5, bottom: total, right: 5 }, { numFmt: '0.0%' }));
  apply(B.setColumnWidth(workbook, 0, 0, 0, 14));
  apply(B.setColumnWidth(workbook, 0, 1, 5, 12));

  // Income and expenses by month, and the profit as a line, beside the table.
  const ref = (col, from, to) => `'${t.sheet}'!$${col}$${from}${to === undefined ? '' : `:$${col}$${to}`}`;
  const series = (col) => ({ nameRef: ref(col, 1), categories: ref('A', 2, last + 1), values: ref(col, 2, last + 1) });
  apply(B.addChart(workbook, 0, { type: 'column', title: t.chart, series: [series('C'), series('D')], legend: 'b', anchor: { top: 0, left: 7, bottom: 15, right: 14 } }));
  apply(B.addChart(workbook, 0, { type: 'line', title: t.trend, series: [series('E')], legend: 'none', anchor: { top: 16, left: 7, bottom: 30, right: 14 } }));
  const engine = new B.CalcEngine(workbook);
  engine.recalculate();
  check(engine.uncalculated().length === 0, 'every formula of the workbook has a value');
  const bytes = await file.save();

  const again = await (await B.open(bytes)).workbook();
  const sheet = again.sheets[0];
  const sum = INCOME.reduce((a, b) => a + b, 0) - EXPENSES.reduce((a, b) => a + b, 0);
  check(B.getCell(sheet, total, 4) && B.getCell(sheet, total, 4).value === sum, `the total profit is ${sum}`);
  check(B.getCell(sheet, 1, 1) && B.getCell(sheet, 1, 1).kind === 'number', 'the dates are dates');
  check(B.sheetCharts(again, 0).length === 2, 'the workbook has two charts');
  return bytes;
}

async function documentSample(B, lang) {
  const t = TEXT[lang];
  const file = await B.open(await B.createBlank('docx', { locale: lang }));
  const doc = await file.document();
  const history = new B.docOps.DocHistory(doc);
  const at = (path, offset) => ({ story: 'body', path, offset });
  const edit = (made) => history.applyEdit(made, undefined);
  const paragraphs = [t.title, t.intro, t.listHeading, ...t.bullets, t.tableHeading, '', t.stepsHeading, ...t.steps, t.end];
  edit(B.docOps.insertText(doc, at([0], 0), paragraphs.join('\n')));
  const index = (text, from = 0) => paragraphs.indexOf(text, from);
  edit(B.docOps.setParaFormat(doc, at([0], 0), { style: 'Title' }));
  for (const heading of [t.listHeading, t.tableHeading, t.stepsHeading]) edit(B.docOps.setParaFormat(doc, at([index(heading)], 0), { style: 'Heading1' }));
  const first = index(t.bullets[0]);
  edit(B.docOps.toggleList(doc, { anchor: at([first], 0), focus: at([first + t.bullets.length - 1], 1) }, 'bullet'));
  const step = index(t.steps[0]);
  edit(B.docOps.toggleList(doc, { anchor: at([step], 0), focus: at([step + t.steps.length - 1], 1) }, 'number'));
  // "BYM Office" in the introduction in bold.
  const name = t.intro.indexOf('BYM Office');
  edit(B.docOps.setRunFormat(doc, { anchor: at([1], name), focus: at([1], name + 'BYM Office'.length) }, { bold: true }));

  // The table, in the empty paragraph after its heading.
  const place = index('', index(t.tableHeading));
  edit(B.docOps.insertTable(doc, at([place], 0), t.table.length, t.table[0].length));
  const table = doc.body.blocks.findIndex((block) => block.type === 'table');
  check(table >= 0, 'the document has a table');
  t.table.forEach((row, r) =>
    row.forEach((text, c) => {
      edit(B.docOps.insertText(doc, at([table, r, c, 0], 0), text));
      if (r === 0) edit(B.docOps.setRunFormat(doc, { anchor: at([table, r, c, 0], 0), focus: at([table, r, c, 0], text.length) }, { bold: true }));
    }),
  );
  const bytes = await file.save();

  const again = await (await B.open(bytes)).document();
  const written = B.documentText(again);
  for (const text of [t.title, t.intro, ...t.bullets, ...t.steps, t.table[2][0], t.end]) check(written.includes(text), `the document says "${text}"`);
  check(again.body.blocks.some((block) => block.type === 'table'), 'the table is saved');
  return bytes;
}

/** name -> bytes. */
export async function makeSamples(B) {
  const out = {};
  for (const lang of Object.keys(TEXT)) {
    out[`sample-sheet-${lang}.xlsx`] = await sheetSample(B, lang);
    out[`sample-doc-${lang}.docx`] = await documentSample(B, lang);
  }
  return out;
}
