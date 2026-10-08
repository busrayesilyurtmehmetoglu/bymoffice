// buildXlsx() on models made in code: the edges of what a model can hold, the
// order of the elements it writes, and what it reports when a model holds
// something a new file cannot carry.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { png } from '../e2e/fixtures/docx';
import { buildXlsx, open, parseXml, type XlsxBuildReport } from '../src/index';
import { resolveCellStyle } from '../src/model/sheet/color';
import {
  MAX_COLS,
  MAX_ROWS,
  createSheet,
  emptyStyleSheet,
  getCell,
  localName,
  type Cell,
  type DrawingPicture,
  type DrawingShape,
  type Sheet,
  type Workbook,
  type XmlElement,
} from '../src/index';
import { OpcPackage } from '../src/opc';
import { buildXlsx as xlsxFixture, sampleWorkbook, worksheetXml } from './fixtures';
import { canonical, workbookDifference } from './workbook-content';
import { allNamed, bookOf, build, put, sheetOf, text } from './xlsx-build-fixtures';

const elementNames = (el: XmlElement) => el.children.filter((node): node is XmlElement => node.type === 'element').map((node) => localName(node.name));

/** Fails unless the children of `el` come in the order of `schema`, each at most once unless it may repeat. */
function assertOrder(el: XmlElement, schema: string[], repeatable: string[] = []): void {
  const names = elementNames(el);
  let last = -1;
  const seen = new Set<string>();
  for (const name of names) {
    const rank = schema.indexOf(name);
    assert.ok(rank >= 0, `<${name}> is not a child of <${el.name}>`);
    assert.ok(rank >= last, `<${name}> comes too late in <${el.name}>: ${names.join(', ')}`);
    assert.ok(!seen.has(name) || repeatable.includes(name), `<${name}> is repeated in <${el.name}>`);
    seen.add(name);
    last = rank;
  }
}

// The sequences of the schema (ECMA-376 part 1, sml.xsd), written out here independently of the writer.
const CT_WORKSHEET = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
const CT_WORKBOOK = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
const CT_STYLESHEET = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
const CT_FONT = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
const CT_BORDER = ['start', 'end', 'left', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];
const CT_RPRELT = ['rFont', 'charset', 'family', 'b', 'i', 'strike', 'outline', 'shadow', 'condense', 'extend', 'color', 'sz', 'u', 'vertAlign', 'scheme'];

test('buildXlsx: a workbook without sheets gets one, and an empty sheet is a sheet (the one the file opens on)', async () => {
  const empty = await build(bookOf());
  assert.deepEqual(empty.after.sheets.map((sheet) => [sheet.name, sheet.state, sheet.rowCount]), [['Sheet1', 'visible', 0]]);
  assert.deepEqual(empty.warnings, ['Workbook: it has no sheets; an empty one is written']);

  const blank = await build(bookOf(sheetOf('Boş')));
  assert.deepEqual(blank.warnings, []);
  assert.deepEqual(blank.after.sheets.map((sheet) => [sheet.name, sheet.rows.size, sheet.dimension]), [['Boş', 0, { top: 0, left: 0, bottom: 0, right: 0 }]]);
  assert.deepEqual(blank.pkg.partNames(), ['/_rels/.rels', '/xl/workbook.xml', '/xl/_rels/workbook.xml.rels', '/xl/worksheets/sheet1.xml', '/xl/styles.xml']);
  assert.equal(text(await blank.pkg.readPart('/xl/worksheets/sheet1.xml')).split('\n')[1], '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/></worksheet>');
});

const AWKWARD = [
  'a < b & c > d "tırnak" \'tek\'',
  '  baştaki ve sondaki boşluk  ',
  '\tsekme\tve\nsatır sonu',
  'satır\r\nbaşı ve yalnız \r',
  '_x000D_ gerçek metin, kaçış değil; _x005F_ de öyle',
  'denetim \u0001\u0002\u001f karakterleri',
  'yarım vekil \ud800 ve tam 😀 \udc00',
  ']]> <![CDATA[ <!-- --> <?xml?> &amp; &#10;',
  'İıĞğŞşÇçÖöÜü — TÜRKÇE KARAKTERLER',
  '',
  ' ',
];

test('buildXlsx: text that needs escaping comes back as it was, in cells, names, links and comments', async () => {
  const sheet = sheetOf('A&B <1> "x"');
  const second = sheetOf(" baştaki boşluk 'ile'", 2);
  AWKWARD.forEach((value, i) => put(sheet, i, 0, { kind: 'string', value }));
  // A formula is literal text: "_x0041_" in it is a name, and its text result is escaped like any text.
  put(sheet, 0, 1, { kind: 'string', value: AWKWARD[5], formula: { text: 'IF(A1<"b&c",_x0041_&"<\'>",A6)', type: 'normal' } });
  put(sheet, 1, 1, { kind: 'error', value: '#N/A' });
  put(sheet, 2, 1, { kind: 'date', value: '2024-02-29T12:30:00' });
  put(sheet, 3, 1, { kind: 'string', value: 'zengin <metin>', runs: [{ text: 'zengin ', font: { name: 'A&B', bold: true, italic: false, strike: false, underline: 'double', size: 12 } }, { text: '<metin>' }] });
  sheet.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/?a=1&b="2"<3>', display: 'a & b', tooltip: '<ipucu> "x"\nsatır' });
  sheet.hyperlinks.push({ range: { top: 1, left: 0, bottom: 2, right: 1 }, location: "'A&B <1> \"x\"'!A1" });
  sheet.comments = [{ row: 0, col: 0, author: 'Ö & <Ş>', text: AWKWARD[0] + '\n' + AWKWARD[5] }];
  const workbook = bookOf(sheet, second);
  workbook.definedNames.push({ name: 'Ad_Ğ', value: '\'A&B <1> "x"\'!$A$1:$B$2', hidden: false, localSheet: undefined }, { name: 'Yerel', value: '"<&>"', hidden: true, localSheet: 1 });

  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, []);
  assert.deepEqual(after.sheets.map((s) => s.name), ['A&B <1> "x"', " baştaki boşluk 'ile'"]);
  AWKWARD.forEach((value, i) => assert.equal(getCell(after.sheets[0], i, 0)?.value, value, `row ${i}`));
  assert.deepEqual(getCell(after.sheets[0], 0, 1)?.formula, { text: 'IF(A1<"b&c",_x0041_&"<\'>",A6)', type: 'normal' });
  assert.equal(getCell(after.sheets[0], 0, 1)?.value, AWKWARD[5]);
  assert.deepEqual([getCell(after.sheets[0], 2, 1)?.kind, getCell(after.sheets[0], 2, 1)?.value], ['date', '2024-02-29T12:30:00']);
  assert.equal(await workbookDifference(workbook, after), undefined);
  assert.deepEqual(after.sheets[0].comments?.map((comment) => [comment.author, comment.text]), [['Ö & <Ş>', AWKWARD[0] + '\n' + AWKWARD[5]]]);
  assert.deepEqual(after.definedNames, workbook.definedNames);
  assert.equal(after.sheets[0].hyperlinks[0].target, 'https://example.com/?a=1&b="2"<3>');

  // Nothing of the model is markup in the file: every part is parsed again, and no part has a document type.
  for (const part of pkg.partNames()) {
    const xml = text(await pkg.readPart(part));
    assert.doesNotMatch(xml, /<!DOCTYPE|<!ENTITY/i, part);
    parseXml(xml);
  }
  // Texts with white space at their ends say that it counts.
  assert.match(text(await pkg.readPart('/xl/sharedStrings.xml')), /<si><t xml:space="preserve"> {2}baştaki ve sondaki boşluk {2}<\/t><\/si>/);
  for (const part of ['/xl/sharedStrings.xml', '/xl/comments1.xml']) {
    for (const run of parseXml(await pkg.readPart(part)).root.children.flatMap((node) => (node.type === 'element' ? allNamed(node, 'rPr') : []))) assertOrder(run, CT_RPRELT);
  }
});


test('buildXlsx: a character XML cannot carry in a name is replaced, and the report says so', async () => {
  const sheet = sheetOf('Ad\u0001ı');
  sheet.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, location: 'A1', tooltip: 'x\u0002y' });
  const { after, warnings } = await build(bookOf(sheet));
  assert.equal(after.sheets[0].name, 'Ad�ı');
  assert.equal(after.sheets[0].hyperlinks[0].tooltip, 'x�y');
  assert.deepEqual(warnings, ['Text: a character XML cannot carry (in a formula, an error value, a name or another attribute) is written as U+FFFD (2 times)']);
  // The count is that of one build.
  assert.deepEqual((await build(bookOf(sheetOf('Temiz')))).warnings, []);
});

test('buildXlsx: the last row and the last column of a sheet', async () => {
  const sheet = sheetOf('Kenar');
  put(sheet, MAX_ROWS - 1, MAX_COLS - 1, { kind: 'number', value: 1 });
  put(sheet, 0, 0, { kind: 'number', value: 2 });
  put(sheet, MAX_ROWS - 1, 0, { kind: 'string', value: 'son satır', formula: { text: 'XFD1048576&"x"', type: 'normal' } });
  sheet.merges.push({ top: MAX_ROWS - 2, left: MAX_COLS - 2, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 });
  sheet.columns.push({ min: MAX_COLS - 1, max: MAX_COLS - 1, width: 30, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false });
  sheet.hyperlinks.push({ range: { top: MAX_ROWS - 1, left: MAX_COLS - 1, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 }, location: 'A1' });
  sheet.comments = [{ row: MAX_ROWS - 1, col: MAX_COLS - 1, text: 'köşe' }];
  sheet.view.frozenRows = 3;
  sheet.view.activeCell = 'XFD1048576';
  const workbook = bookOf(sheet);
  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, []);
  assert.equal(await workbookDifference(workbook, after), undefined);
  assert.equal(getCell(after.sheets[0], MAX_ROWS - 1, MAX_COLS - 1)?.value, 1);
  assert.deepEqual(after.sheets[0].dimension, { top: 0, left: 0, bottom: MAX_ROWS - 1, right: MAX_COLS - 1 });
  assert.deepEqual([after.sheets[0].view.frozenRows, after.sheets[0].view.topLeftCell, after.sheets[0].view.activeCell], [3, 'A4', 'XFD1048576']);
  const xml = text(await pkg.readPart('/xl/worksheets/sheet1.xml'));
  assert.match(xml, /<dimension ref="A1:XFD1048576"\/>/);
  assert.match(xml, /<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"\/><selection pane="bottomLeft" activeCell="XFD1048576" sqref="XFD1048576"\/>/);
  assert.match(xml, /<col min="16384" max="16384" width="30" customWidth="1"\/>/);
  // The box of a comment in the last cell stays on the sheet.
  assert.match(text(await pkg.readPart('/xl/drawings/vmlDrawing1.vml')), /<x:Anchor>16383, 15, 1048574, 10, 16383, 15, 1048575, 4<\/x:Anchor>/);

  // What lies outside the sheet is left out, and said.
  const outside = sheetOf('Dışarı');
  put(outside, MAX_ROWS, 0, { kind: 'number', value: 1 });
  put(outside, 0, MAX_COLS, { kind: 'number', value: 1 });
  put(outside, 0, 0, { kind: 'number', value: 7 });
  outside.merges.push({ top: 0, left: 0, bottom: MAX_ROWS, right: 0 });
  outside.comments = [{ row: 0, col: MAX_COLS, text: 'x' }, { row: 0, col: 0, text: 'ilk' }, { row: 0, col: 0, text: 'ikinci' }];
  const cut = await build(bookOf(outside));
  assert.deepEqual(cut.warnings, [
    'Sheet: a row outside the sheet is left out',
    'Sheet: cells outside the sheet are left out',
    'Sheet: a merged range outside the sheet is left out',
    'Comment: a comment outside the sheet is left out',
    'Comment: a second comment on the same cell is left out',
  ]);
  assert.deepEqual([cut.after.sheets[0].rows.size, getCell(cut.after.sheets[0], 0, 0)?.value, cut.after.sheets[0].merges.length, cut.after.sheets[0].comments?.[0].text], [1, 7, 0, 'ilk']);
});

test('buildXlsx: thousands of formats, each resolved as in the model', async () => {
  const sheet = sheetOf('Stiller');
  const workbook = bookOf(sheet);
  const styles = workbook.styles;
  // The model's second fill is a colour: in a file the second fill is always "gray125".
  styles.numFmts.set(164, '0.000" <birim> & "');
  styles.numFmts.set(7, '#,##0.00 ₺');
  const count = 3000;
  for (let i = 0; i < count; i++) {
    styles.fonts.push({ name: i % 2 ? 'Arial' : 'Çağdaş "Yazı"', size: 8 + (i % 20) / 2, bold: i % 3 === 0, italic: i % 5 === 0, strike: i % 7 === 0, underline: i % 11 === 0 ? 'doubleAccounting' : 'none', color: { rgb: `FF${(i * 4099).toString(16).padStart(6, '0').slice(-6).toUpperCase()}` }, vertAlign: i % 13 === 0 ? 'superscript' : undefined, family: 2, charset: 162, scheme: i % 17 === 0 ? 'minor' : undefined });
    styles.fills.push({ type: 'pattern', pattern: i % 4 === 0 ? 'darkGrid' : 'solid', fg: { theme: i % 10, tint: -0.25 }, bg: { indexed: 8 + (i % 50) } });
    styles.borders.push({ left: { style: 'thin', color: { auto: true } }, bottom: i % 2 ? { style: 'dashDot', color: { rgb: 'FF00FF00' } } : undefined, diagonal: { style: 'hair' }, diagonalUp: i % 2 === 0, diagonalDown: i % 3 === 0 });
    styles.cellXfs.push({
      numFmtId: [0, 14, 164, 7][i % 4],
      fontId: i + 1,
      fillId: i + 1,
      borderId: i + 1,
      xfId: 0,
      quotePrefix: i % 9 === 0,
      alignment: i % 2 ? { horizontal: 'centerContinuous', vertical: 'distributed', wrapText: true, textRotation: 255, indent: 3, readingOrder: 2, shrinkToFit: false } : undefined,
      protection: i % 3 ? { locked: false, hidden: true } : undefined,
    });
    put(sheet, Math.floor(i / 50), i % 50, { kind: 'number', value: i, style: i + 1 });
  }
  styles.cellStyles.push({ name: 'Normal', xfId: 0, builtinId: 0 });
  sheet.rows.get(3)!.style = 17;
  sheet.columns.push({ min: 60, max: 70, style: 29, customWidth: false, hidden: false, outlineLevel: 0, collapsed: false });
  styles.themeColors = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080'];
  styles.indexedColors = Array.from({ length: 64 }, (_, i) => (i * 0x030507).toString(16).padStart(6, '0').slice(-6).toUpperCase());

  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, []);
  assert.equal(await workbookDifference(workbook, after), undefined);
  for (const i of [0, 1, 2, 1499, 2999]) {
    const was = resolveCellStyle(styles, i + 1);
    const is = resolveCellStyle(after.styles, getCell(after.sheets[0], Math.floor(i / 50), i % 50)!.style);
    assert.equal(canonical(is), canonical(was), `format ${i}`);
  }
  assert.equal(after.styles.cellXfs.length, count + 1);
  // The fixed fill stands between the first fill and the model's second one.
  assert.deepEqual(after.styles.fills.slice(0, 2), [{ type: 'pattern', pattern: 'none' }, { type: 'pattern', pattern: 'gray125' }]);
  assert.equal(after.styles.fills.length, count + 2);
  assert.deepEqual(after.styles.themeColors, styles.themeColors);
  assert.deepEqual(after.styles.indexedColors, styles.indexedColors);
  assert.equal(pkg.contentTypeOf('/xl/theme/theme1.xml'), 'application/vnd.openxmlformats-officedocument.theme+xml');

  const root = parseXml(await pkg.readPart('/xl/styles.xml')).root;
  assertOrder(root, CT_STYLESHEET);
  assert.deepEqual(elementNames(root), ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'colors']);
  for (const font of allNamed(root, 'font')) assertOrder(font, CT_FONT);
  for (const border of allNamed(root, 'border')) assertOrder(border, CT_BORDER);
  for (const xf of allNamed(root, 'xf')) assertOrder(xf, ['alignment', 'protection', 'extLst']);
  for (const table of ['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'numFmts']) {
    const el = allNamed(root, table)[0];
    assert.equal(el.attrs.get('count'), String(elementNames(el).length), table);
  }
  const theme = parseXml(await pkg.readPart('/xl/theme/theme1.xml')).root;
  assertOrder(allNamed(theme, 'clrScheme')[0], ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']);
  assertOrder(allNamed(theme, 'themeElements')[0], ['clrScheme', 'fontScheme', 'fmtScheme']);
});

test('buildXlsx: very long texts and many strings', async () => {
  const sheet = sheetOf('Uzun');
  const long = 'ğüşiöç<&>"\n'.repeat(20000);
  put(sheet, 0, 0, { kind: 'string', value: long });
  put(sheet, 0, 1, { kind: 'string', value: 'x'.repeat(32767) });
  for (let i = 0; i < 5000; i++) put(sheet, 1 + i, 0, { kind: 'string', value: `metin ${i % 100}` });
  const workbook = bookOf(sheet);
  const { after, pkg } = await build(workbook);
  assert.equal(getCell(after.sheets[0], 0, 0)?.value, long);
  assert.equal((getCell(after.sheets[0], 0, 1)?.value as string).length, 32767);
  assert.equal(await workbookDifference(workbook, after), undefined);
  // A text used by many cells is one string of the table.
  assert.equal(after.sharedStrings.length, 102);
  assert.match(text(await pkg.readPart('/xl/sharedStrings.xml')), /<sst [^>]*count="5002" uniqueCount="102">/);
});

test('buildXlsx: strings the model shares keep their indexes; a cell whose text changed gets a new one', async () => {
  const sheet = sheetOf('S');
  const workbook = bookOf(sheet);
  workbook.sharedStrings = [{ text: 'sıfır' }, { text: 'bir', runs: [{ text: 'b', font: { bold: true, italic: false, strike: false, underline: 'none' } }, { text: 'ir' }] }, { text: 'kullanılmayan' }];
  put(sheet, 0, 0, { kind: 'string', value: 'bir', si: 1, runs: workbook.sharedStrings[1].runs });
  put(sheet, 0, 1, { kind: 'string', value: 'sıfır', si: 0 });
  // Edited since it was read: the text is no longer the one of its string.
  put(sheet, 0, 2, { kind: 'string', value: 'değişti', si: 0 });
  put(sheet, 0, 3, { kind: 'string', value: 'yok', si: 99 });
  const { after } = await build(workbook);
  assert.deepEqual(after.sharedStrings.map((item) => item.text), ['sıfır', 'bir', 'kullanılmayan', 'değişti', 'yok']);
  assert.deepEqual([0, 1, 2, 3].map((col) => [getCell(after.sheets[0], 0, col)?.value, getCell(after.sheets[0], 0, col)?.si]), [['bir', 1], ['sıfır', 0], ['değişti', 3], ['yok', 4]]);
  assert.deepEqual(getCell(after.sheets[0], 0, 0)?.runs?.map((run) => [run.text, run.font?.bold]), [['b', true], ['ir', undefined]]);
});

test('buildXlsx: rows are written in the order of the sheet, whatever order the model got them in', async () => {
  const sheet = sheetOf('Sıra');
  for (const row of [9, 2, 5, 0]) put(sheet, row, 3, { kind: 'number', value: row });
  put(sheet, 2, 0, { kind: 'number', value: -1 });
  const { after, pkg } = await build(bookOf(sheet));
  assert.deepEqual([...after.sheets[0].rows.keys()], [0, 2, 5, 9]);
  assert.deepEqual(text(await pkg.readPart('/xl/worksheets/sheet1.xml')).match(/<(row|c) r="[A-Z0-9]+"/g), ['<row r="1"', '<c r="D1"', '<row r="3"', '<c r="A3"', '<c r="D3"', '<row r="6"', '<c r="D6"', '<row r="10"', '<c r="D10"']);
});

test('buildXlsx: shared formulas are kept where a file can state them and written out elsewhere', async () => {
  const sheet = sheetOf('F');
  // Group 0: the first cell in the order of the sheet is not the one with the text.
  put(sheet, 0, 0, { kind: 'number', value: 1, formula: { text: '', type: 'shared', si: 0 } });
  put(sheet, 1, 0, { kind: 'number', value: 2, formula: { text: 'B2+$C$1', type: 'shared', si: 0, ref: 'A1:A3' } });
  put(sheet, 2, 0, { kind: 'number', value: 3, formula: { text: '', type: 'shared', si: 0 } });
  // Group 1 has a text without a range; group 2 has no text at all.
  put(sheet, 0, 5, { kind: 'number', value: 4, formula: { text: 'A1*2', type: 'shared', si: 1 } });
  put(sheet, 1, 5, { kind: 'number', value: 5, formula: { text: '', type: 'shared', si: 1 } });
  put(sheet, 4, 5, { kind: 'number', value: 6, formula: { text: '', type: 'shared', si: 2 } });
  // An array formula, a data table, a formula without a result, and metadata the model cannot write.
  put(sheet, 6, 0, { kind: 'number', value: 7, formula: { text: 'SUM(A1:A3*2)', type: 'array', ref: 'A7:B7' } });
  put(sheet, 7, 0, { kind: 'number', value: 8, formula: { text: '', type: 'dataTable', ref: 'A8:B9', attrs: new Map([['dt2D', '1'], ['dtr', '1'], ['r1', 'C1'], ['r2', 'C2'], ['xr:uid', '{1}']]) } });
  put(sheet, 8, 5, { kind: 'blank', value: null, formula: { text: 'NOW()', type: 'normal' } });
  put(sheet, 9, 5, { kind: 'number', value: 9, cm: 1, vm: 2, ph: true });
  const workbook = bookOf(sheet);
  const { after, warnings, pkg } = await build(workbook);
  const formulaOf = (row: number, col: number) => getCell(after.sheets[0], row, col)?.formula;
  assert.deepEqual(formulaOf(0, 0), { text: 'B1+$C$1', type: 'normal' });
  assert.deepEqual(formulaOf(1, 0), { text: 'B2+$C$1', type: 'shared', si: 0, ref: 'A1:A3' });
  assert.deepEqual(formulaOf(2, 0), { text: '', type: 'shared', si: 0 });
  assert.deepEqual(formulaOf(0, 5), { text: 'A1*2', type: 'normal' });
  assert.deepEqual(formulaOf(1, 5), { text: 'A2*2', type: 'normal' });
  assert.equal(formulaOf(4, 5), undefined);
  assert.equal(getCell(after.sheets[0], 4, 5)?.value, 6);
  assert.deepEqual(formulaOf(6, 0), { text: 'SUM(A1:A3*2)', type: 'array', ref: 'A7:B7' });
  assert.deepEqual(formulaOf(7, 0), { text: '', type: 'dataTable', ref: 'A8:B9', attrs: new Map([['dt2D', '1'], ['dtr', '1'], ['r1', 'C1'], ['r2', 'C2']]) });
  assert.deepEqual(getCell(after.sheets[0], 8, 5), { kind: 'blank', value: null, style: 0, formula: { text: 'NOW()', type: 'normal' } });
  assert.deepEqual(getCell(after.sheets[0], 9, 5), { kind: 'number', value: 9, style: 0, ph: true });
  assert.deepEqual(warnings, [
    'Formula: a cell of a shared formula whose first cell is missing keeps its value and has no formula',
    'Formula: an attribute of another namespace is left out',
    'Cell: metadata (a dynamic array or a rich value) is not in the model and its index is left out',
  ]);
  // A formula without a stored result is calculated when the file is opened.
  const book = parseXml(await pkg.readPart('/xl/workbook.xml')).root;
  assert.equal(allNamed(book, 'calcPr')[0].attrs.get('fullCalcOnLoad'), '1');
  assertOrder(book, CT_WORKBOOK);
  // The model is not changed by writing it.
  assert.deepEqual(getCell(sheet, 0, 0)?.formula, { text: '', type: 'shared', si: 0 });
  assert.equal(getCell(sheet, 9, 5)?.cm, 1);
});

test('buildXlsx: the parts of a workbook and of a sheet come in the order of the schema', async () => {
  const sheet = sheetOf('Her şey');
  sheet.tabColor = { rgb: 'FFFF0000' };
  sheet.view = { ...sheet.view, showGridLines: false, showRowColHeaders: false, showZeros: false, rightToLeft: true, tabSelected: true, zoomScale: 85, frozenRows: 1, frozenCols: 2, firstVisibleCell: 'B3', activeCell: 'D5' };
  sheet.baseColWidth = 12;
  sheet.defaultColWidth = 14.5;
  sheet.defaultRowHeight = 18;
  sheet.columns.push({ min: 0, max: 1, width: 20, customWidth: true, hidden: false, outlineLevel: 1, collapsed: false }, { min: 4, max: 4, hidden: true, customWidth: false, outlineLevel: 0, collapsed: true });
  for (let r = 0; r < 6; r++) for (let c = 0; c < 4; c++) put(sheet, r, c, { kind: 'number', value: r * 10 + c });
  Object.assign(sheet.rows.get(2)!, { height: 30, customHeight: true, hidden: true, outlineLevel: 2, collapsed: true, attrs: [['thickBot', '1'], ['x14ac:dyDescent', '0.25']] });
  sheet.merges.push({ top: 7, left: 0, bottom: 8, right: 1 });
  sheet.hyperlinks.push({ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/' });
  sheet.autoFilter = {
    range: { top: 0, left: 0, bottom: 5, right: 3 },
    columns: [
      { colId: 0, kind: 'values', values: ['1', '<2>'], blank: true },
      { colId: 1, kind: 'custom', and: true, conditions: [{ operator: 'greaterThan', value: '3' }, { operator: 'lessThan', value: '50' }] },
      { colId: 2, kind: 'top10', top: false, percent: true, value: 10 },
      { colId: 3, kind: 'color', raw: '<x:filterColumn colId="3"><x:colorFilter dxfId="0"/></x:filterColumn>' },
    ],
    sortState: { range: { top: 1, left: 0, bottom: 5, right: 3 }, conditions: [{ range: { top: 1, left: 1, bottom: 5, right: 1 }, descending: true }], attrs: [['xr:uid', '{2}']] },
    attrs: [['xr:uid', '{1}']],
  };
  sheet.conditionalFormats = [
    {
      ranges: [{ top: 0, left: 0, bottom: 5, right: 0 }, { top: 0, left: 2, bottom: 5, right: 2 }],
      rules: [
        { type: 'cellIs', priority: 2, stopIfTrue: true, dxfId: 0, operator: 'between', formulas: ['1', '$D$1<5'], percent: false, bottom: false, aboveAverage: true, equalAverage: false },
        { type: 'top10', priority: 1, stopIfTrue: false, formulas: [], rank: 3, percent: true, bottom: true, aboveAverage: true, equalAverage: false },
        { type: 'aboveAverage', priority: Infinity, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: false, equalAverage: true, stdDev: 2 },
        { type: 'containsText', priority: 7, stopIfTrue: false, formulas: ['NOT(ISERROR(SEARCH("a&b",A1)))'], operator: 'containsText', text: 'a&b', percent: false, bottom: false, aboveAverage: true, equalAverage: false },
      ],
    },
    {
      ranges: [{ top: 0, left: 1, bottom: 5, right: 1 }],
      rules: [
        { type: 'colorScale', priority: 4, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, colorScale: { values: [{ type: 'min', gte: true }, { type: 'percentile', value: '50', gte: true }, { type: 'max', gte: true }], colors: [{ rgb: 'FFF8696B' }, { theme: 4 }, { rgb: 'FF63BE7B' }] } },
        { type: 'dataBar', priority: 5, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, dataBar: { min: { type: 'num', value: '0', gte: true }, max: { type: 'formula', value: '$D$6', gte: true }, color: { rgb: 'FF638EC6' }, showValue: false, minLength: 5, maxLength: 95, extended: false, gradient: true, negativeFillSameAsPositive: false, negativeBorderSameAsPositive: true, axisPosition: 'automatic', direction: 'context' } },
        { type: 'iconSet', priority: 6, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, iconSet: { set: '4Rating', showValue: false, reverse: true, values: [{ type: 'percent', value: '0', gte: true }, { type: 'percent', value: '25', gte: false }, { type: 'num', value: '3', gte: true }, { type: 'percent', value: '75', gte: true }] } },
        { type: 'dataBar', priority: 9, stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, extended: true },
      ],
    },
  ];
  sheet.dataValidations = [
    { ranges: [{ top: 1, left: 0, bottom: 5, right: 0 }], type: 'list', operator: 'between', formula1: '"a,b<c>,ç"', allowBlank: true, hideDropDown: true, showInputMessage: true, promptTitle: 'Başlık & <x>', prompt: 'Seçin\n"biri"', showErrorMessage: true, errorStyle: 'warning', errorTitle: 'Hata', error: 'Olmaz' },
    { ranges: [{ top: 1, left: 1, bottom: 5, right: 1 }, { top: 9, left: 9, bottom: 9, right: 9 }], type: 'whole', operator: 'notBetween', formula1: '1', formula2: 'Öteki!$A$1', allowBlank: false, hideDropDown: false, showInputMessage: false, showErrorMessage: false, errorStyle: 'stop' },
  ];
  sheet.comments = [{ row: 1, col: 1, author: 'A', text: 'not', visible: true }];
  sheet.drawings = [{ anchor: { kind: 'absolute', x: 9525, y: 19050, width: 952500, height: 476250 }, object: { type: 'shape', geometry: 'rect', fill: '#FF0000' } }];
  const other = sheetOf('Öteki', 1);
  other.state = 'veryHidden';
  const workbook = bookOf(sheet, other);
  workbook.date1904 = true;
  workbook.activeSheet = 0;
  workbook.definedNames.push({ name: '_xlnm._FilterDatabase', value: "'Her şey'!$A$1:$D$6", localSheet: 0, hidden: true });
  workbook.styles.dxfs.push({ type: 'element', name: 'x:dxf', attrs: new Map(), children: [{ type: 'element', name: 'x:font', attrs: new Map(), children: [{ type: 'element', name: 'x:b', attrs: new Map(), children: [] }] }] });

  const { after, warnings, pkg, bytes } = await build(workbook);
  const root = parseXml(await pkg.readPart('/xl/worksheets/sheet1.xml')).root;
  assertOrder(root, CT_WORKSHEET, ['conditionalFormatting']);
  assert.deepEqual(elementNames(root), ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'autoFilter', 'mergeCells', 'conditionalFormatting', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'drawing', 'legacyDrawing']);
  assertOrder(allNamed(root, 'sheetView')[0], ['pane', 'selection', 'pivotSelection', 'extLst'], ['selection']);
  assertOrder(allNamed(root, 'autoFilter')[0], ['filterColumn', 'sortState', 'extLst'], ['filterColumn']);
  for (const rule of allNamed(root, 'cfRule')) assertOrder(rule, ['formula', 'colorScale', 'dataBar', 'iconSet', 'extLst'], ['formula']);
  for (const scale of allNamed(root, 'colorScale')) assertOrder(scale, ['cfvo', 'color'], ['cfvo', 'color']);
  for (const validation of allNamed(root, 'dataValidation')) assertOrder(validation, ['formula1', 'formula2']);
  const book = parseXml(await pkg.readPart('/xl/workbook.xml')).root;
  assertOrder(book, CT_WORKBOOK);
  assert.deepEqual(elementNames(book), ['workbookPr', 'bookViews', 'sheets', 'definedNames']);
  const styles = parseXml(await pkg.readPart('/xl/styles.xml')).root;
  assertOrder(styles, CT_STYLESHEET);
  // A record kept from a file whose elements had a prefix is written without it.
  assert.equal(text(await pkg.readPart('/xl/styles.xml')).includes('<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>'), true);
  const drawing = parseXml(await pkg.readPart('/xl/drawings/drawing1.xml')).root;
  assertOrder(allNamed(drawing, 'absoluteAnchor')[0], ['pos', 'ext', 'sp', 'grpSp', 'graphicFrame', 'cxnSp', 'pic', 'contentPart', 'clientData']);
  assertOrder(allNamed(drawing, 'sp')[0], ['nvSpPr', 'spPr', 'style', 'txBody']);
  assertOrder(allNamed(drawing, 'spPr')[0], ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst']);
  const comments = parseXml(await pkg.readPart('/xl/comments1.xml')).root;
  assertOrder(comments, ['authors', 'commentList', 'extLst']);

  // Sheet ids differ in a file, whatever the model says.
  assert.deepEqual(allNamed(book, 'sheet').map((el) => el.attrs.get('sheetId')), ['1', '2']);
  assert.deepEqual(warnings, [
    'Row: an attribute that cannot be written as it is (one the writer states itself, one of another namespace, or one whose name is not a name) is left out',
    'Filter: a criterion the model only keeps as source text (a colour, an icon) is left out',
    'Filter: an attribute or an extension of another namespace (or a sort state outside the sheet) is left out (2 times)',
    'Conditional format: a rule of the 2010 extension is left out',
  ]);

  const read = after.sheets[0];
  assert.deepEqual(read.view, { showGridLines: false, showRowColHeaders: false, showZeros: false, rightToLeft: true, tabSelected: true, zoomScale: 85, frozenRows: 1, frozenCols: 2, firstVisibleCell: 'B3', topLeftCell: 'D4', activeCell: 'D5' });
  assert.deepEqual([read.tabColor, read.baseColWidth, read.defaultColWidth, read.defaultRowHeight], [{ rgb: 'FFFF0000' }, 12, 14.5, 18]);
  assert.deepEqual(read.columns, sheet.columns);
  // Attributes of a namespace the new part does not declare are hints, and left out.
  assert.deepEqual({ ...read.rows.get(2), cells: undefined }, { height: 30, customHeight: true, hidden: true, outlineLevel: 2, collapsed: true, attrs: [['thickBot', '1']], cells: undefined });
  assert.deepEqual(read.autoFilter?.range, sheet.autoFilter.range);
  assert.deepEqual(read.autoFilter?.columns.map((column) => ({ ...column, raw: undefined })), [
    { colId: 0, kind: 'values', values: ['1', '<2>'], blank: true, raw: undefined },
    { colId: 1, kind: 'custom', and: true, conditions: [{ operator: 'greaterThan', value: '3' }, { operator: 'lessThan', value: '50' }], raw: undefined },
    { colId: 2, kind: 'top10', top: false, percent: true, value: 10, raw: undefined },
  ]);
  assert.deepEqual([read.autoFilter?.sortState?.range, read.autoFilter?.sortState?.conditions.map((condition) => [condition.range, condition.descending])], [{ top: 1, left: 0, bottom: 5, right: 3 }, [[{ top: 1, left: 1, bottom: 5, right: 1 }, true]]]);
  const expected = sheet.conditionalFormats.map((format) => ({ ranges: format.ranges, rules: format.rules.filter((rule) => !rule.extended) }));
  // A rule without a priority gets the next one after the highest.
  expected[0].rules[2] = { ...expected[0].rules[2], priority: 10 };
  // Where a rule is in the part that was written (`src`) is not what the model said.
  const placeless = <T extends { src?: unknown }>(item: T): T => ({ ...item, src: undefined });
  assert.equal(canonical(read.conditionalFormats!.map((format) => ({ ...placeless(format), rules: format.rules.map(placeless) }))), canonical(expected));
  assert.equal(canonical(read.dataValidations!.map(placeless)), canonical(sheet.dataValidations));
  assert.equal(after.styles.dxfFormats?.[0].font?.bold, true);
  assert.deepEqual([after.date1904, after.sheets[1].state, after.definedNames], [true, 'veryHidden', workbook.definedNames]);

  // The same model gives the same bytes, and a save leaves them alone.
  assert.deepEqual(await buildXlsx(workbook), bytes);
  const reopened = await open(bytes);
  await reopened.workbook();
  assert.deepEqual(await reopened.save(), bytes);
});

test('buildXlsx: drawings of every kind of anchor, and what is not in the model', async () => {
  const image = png(4, 4, [0, 0, 200]);
  const sheet = sheetOf('Çizim');
  sheet.columns.push({ min: 1, max: 1, width: 20, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false }, { min: 3, max: 3, hidden: true, customWidth: false, outlineLevel: 0, collapsed: false });
  const text0: DrawingShape['text'] = {
    anchor: 'middle',
    wrap: false,
    insets: [4, 2, 4, 2],
    paragraphs: [
      { align: 'center', runs: [{ text: 'kalın <a&b>', bold: true, size: 14, color: '#112233', font: 'Çağrı "x"' }, { text: ' ince\nikinci satır', italic: true, underline: true, strike: true }] },
      { runs: [] },
      { align: 'right', runs: [{ text: '  sağda  ' }] },
    ],
  };
  sheet.drawings = [
    { anchor: { kind: 'oneCell', from: { col: 1, colOff: 95250, row: 2, rowOff: 47625 }, width: 1905000, height: 952500 }, object: { type: 'picture', media: 'bir', name: 'Resim', description: 'açıklama', hidden: true, flipV: true, crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.05 } } },
    { anchor: { kind: 'absolute', x: 0, y: 0, width: 1270000, height: 635000 }, object: { type: 'shape', geometry: 'roundRect', fill: '#00FF00', line: { color: '#0000FF', width: 3 }, rotation: 270, text: text0 } },
    { anchor: { kind: 'twoCell', editAs: 'absolute', from: { col: 0, colOff: 0, colShare: 0.5, row: 10, rowOff: 0, rowShare: 0.25 }, to: { col: 3, colOff: 0, colShare: 0.9, row: 12, rowOff: 0, rowShare: 1 } }, object: { type: 'shape', geometry: 'bentConnector3', connector: true, line: { color: '#000000', width: 1, dash: true }, flipH: true } },
    {
      anchor: { kind: 'twoCell', from: { col: 5, colOff: 0, row: 0, rowOff: 0 }, to: { col: 9, colOff: 0, row: 10, rowOff: 0 } },
      object: {
        type: 'group',
        name: 'Küme',
        children: [
          { type: 'picture', media: 'bir', box: { x: 0, y: 0, width: 0.5, height: 0.5 } },
          { type: 'group', box: { x: 0.5, y: 0.5, width: 0.5, height: 0.5 }, children: [{ type: 'shape', geometry: 'ellipse', box: { x: 0.25, y: 0.25, width: 0.5, height: 0.5 } }, { type: 'placeholder', what: 'diagram', box: { x: 0, y: 0, width: 1, height: 1 } }] },
          { type: 'shape', geometry: 'rect' },
        ],
      },
    },
    { anchor: { kind: 'twoCell', from: { col: 0, colOff: 0, row: 20, rowOff: 0 }, to: { col: 2, colOff: 0, row: 25, rowOff: 0 } }, object: { type: 'placeholder', what: 'chart', name: 'Grafik 1' } },
    { anchor: { kind: 'twoCell', from: { col: 0, colOff: 0, row: 30, rowOff: 0 }, to: { col: 2, colOff: 0, row: 35, rowOff: 0 } }, object: { type: 'picture', external: true } },
    { anchor: { kind: 'twoCell', from: { col: 0, colOff: 0, row: 40, rowOff: 0 }, to: { col: 2, colOff: 0, row: 45, rowOff: 0 } }, object: { type: 'picture', media: 'yok' } },
    { anchor: { kind: 'twoCell', from: { col: MAX_COLS, colOff: 0, row: 0, rowOff: 0 }, to: { col: 2, colOff: 0, row: 45, rowOff: 0 } }, object: { type: 'shape', geometry: 'rect' } },
  ];
  sheet.comments = [
    { row: 0, col: 0, text: 'ilk ileti', author: 'A', threaded: true, replies: [{ author: 'B', text: 'yanıt' }, { text: 'bir yanıt daha' }] },
    { row: 3, col: 3, text: 'kutulu', box: { kind: 'twoCell', from: { col: 4, colOff: 0, colShare: 0.5, row: 3, rowOff: 95250 }, to: { col: 6, colOff: 190500, row: 8, rowOff: 0, rowShare: 0.5 } } },
  ];
  const workbook = bookOf(sheet);
  workbook.media = { readPart: (key) => (key === 'bir' ? Promise.resolve(image) : Promise.reject(new Error('no such picture'))) };

  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, [
    'Drawing: a diagram is not in the model and is left out',
    'Drawing: a member of a group without a place in the group is left out',
    'Drawing: a chart is not in the model and is left out',
    'Drawing: the address of a linked picture is not in the model; it is written as an empty frame',
    'Drawing: a picture whose bytes cannot be read (a metafile, or one missing from the file) is written as an empty frame',
    'Drawing: an object whose anchor names no cell of the sheet is left out',
    'Comment: a conversation is written as a note with its first message; replies are left out (3 times)',
  ]);
  const drawings = after.sheets[0].drawings!;
  assert.deepEqual(drawings.map((drawing) => [drawing.anchor.kind, drawing.object.type]), [['oneCell', 'picture'], ['absolute', 'shape'], ['twoCell', 'shape'], ['twoCell', 'group'], ['twoCell', 'picture'], ['twoCell', 'picture']]);
  assert.deepEqual(drawings[0].anchor, sheet.drawings[0].anchor);
  assert.deepEqual(drawings[1].anchor, sheet.drawings[1].anchor);
  // Shares of a cell are offsets in a file: half of column A (64 px), a quarter of row 11 (20 px); nothing of the hidden column D.
  assert.deepEqual(drawings[2].anchor, { kind: 'twoCell', editAs: 'absolute', from: { col: 0, colOff: 32 * 9525, row: 10, rowOff: 5 * 9525 }, to: { col: 3, colOff: 0, row: 12, rowOff: 20 * 9525 } });
  const picture = drawings[0].object as DrawingPicture;
  assert.deepEqual(picture, { type: 'picture', name: 'Resim', description: 'açıklama', hidden: true, flipV: true, media: '/xl/media/image1.png', crop: { left: 0.1, top: 0.2, right: 0.3, bottom: 0.05 } });
  assert.deepEqual(await after.media!.readPart('/xl/media/image1.png'), image);
  const shape = drawings[1].object as DrawingShape;
  assert.deepEqual([shape.geometry, shape.fill, shape.line, shape.rotation], ['roundRect', '#00FF00', { color: '#0000FF', width: 3 }, 270]);
  assert.deepEqual([shape.text?.anchor, shape.text?.wrap, shape.text?.insets], ['middle', false, [4, 2, 4, 2]]);
  assert.deepEqual(shape.text?.paragraphs.map((paragraph) => [paragraph.align, paragraph.runs.map((run) => run.text)]), [
    ['center', ['kalın <a&b>', ' ince', '\n', 'ikinci satır']],
    [undefined, []],
    ['right', ['  sağda  ']],
  ]);
  assert.deepEqual(shape.text?.paragraphs[0].runs[0], { text: 'kalın <a&b>', bold: true, size: 14, color: '#112233', font: 'Çağrı "x"' });
  assert.deepEqual(shape.text?.paragraphs[0].runs[3], { text: 'ikinci satır', italic: true, underline: true, strike: true });
  const connector = drawings[2].object as DrawingShape;
  assert.deepEqual([connector.geometry, connector.connector, connector.fill, connector.line, connector.flipH], ['bentConnector3', true, undefined, { color: '#000000', width: 1, dash: true }, true]);
  const group = drawings[3].object;
  assert.ok(group.type === 'group');
  assert.equal(group.name, 'Küme');
  assert.deepEqual(group.children.map((child) => [child.type, child.box]), [['picture', { x: 0, y: 0, width: 0.5, height: 0.5 }], ['group', { x: 0.5, y: 0.5, width: 0.5, height: 0.5 }]]);
  const inner = group.children[1];
  assert.ok(inner.type === 'group');
  assert.deepEqual(inner.children.map((child) => [child.type, child.box]), [['shape', { x: 0.25, y: 0.25, width: 0.5, height: 0.5 }]]);
  assert.deepEqual([(drawings[4].object as DrawingPicture).media, (drawings[4].object as DrawingPicture).external, (drawings[5].object as DrawingPicture).media], [undefined, undefined, undefined]);
  // Ids of the objects of a drawing differ.
  const ids = allNamed(parseXml(await pkg.readPart('/xl/drawings/drawing1.xml')).root, 'cNvPr').map((el) => el.attrs.get('id'));
  assert.equal(new Set(ids).size, ids.length);
  assert.equal((await pkg.relationships('/xl/drawings/drawing1.xml')).length, 1);

  assert.deepEqual(after.sheets[0].comments?.map((comment) => [comment.row, comment.col, comment.author, comment.text, comment.threaded]), [[0, 0, 'A', 'ilk ileti', undefined], [3, 3, undefined, 'kutulu', undefined]]);
  // The box of a comment is written in pixels: half of column E, 10 px into row 4; 20 px into column G, half of row 9.
  assert.deepEqual(after.sheets[0].comments?.[1].box, { kind: 'twoCell', from: { col: 4, colOff: 32 * 9525, row: 3, rowOff: 10 * 9525 }, to: { col: 6, colOff: 20 * 9525, row: 8, rowOff: 10 * 9525 } });
  assert.equal(await workbookDifference({ ...workbook, sheets: [{ ...sheet, drawings: sheet.drawings.slice(0, 3), comments: undefined }] }, { ...after, sheets: [{ ...after.sheets[0], drawings: drawings.slice(0, 3), comments: undefined }] }), undefined);
});

test('buildXlsx: more comments than one block of shape ids holds', async () => {
  const first = sheetOf('Çok');
  first.comments = Array.from({ length: 2500 }, (_, i) => ({ row: i, col: 0, text: `not ${i}`, author: `yazar ${i % 3}` }));
  const second = sheetOf('Az', 2);
  second.comments = [{ row: 0, col: 0, text: 'tek' }];
  const { after, pkg } = await build(bookOf(first, second));
  assert.equal(after.sheets[0].comments?.length, 2500);
  assert.deepEqual([after.sheets[0].comments?.[2499].text, after.sheets[0].comments?.[2499].author], ['not 2499', 'yazar 0']);
  const vml = text(await pkg.readPart('/xl/drawings/vmlDrawing1.vml'));
  assert.match(vml, /<o:idmap v:ext="edit" data="1,2,3"\/>/);
  const ids = vml.match(/_x0000_s\d+/g)!;
  assert.equal(new Set(ids).size, 2500);
  // The next drawing starts in a block of its own.
  assert.match(text(await pkg.readPart('/xl/drawings/vmlDrawing2.vml')), /data="4".*_x0000_s4097"/s);
  assert.equal(allNamed(parseXml(await pkg.readPart('/xl/comments1.xml')).root, 'author').length, 3);
});

test('buildXlsx: names and sheets a file cannot have are left out or reported', async () => {
  const hidden = sheetOf('Gizli');
  hidden.state = 'hidden';
  hidden.sheetId = 5;
  const twin = sheetOf('İkiz');
  twin.sheetId = 5;
  twin.state = 'veryHidden';
  const workbook = bookOf(hidden, twin);
  workbook.activeSheet = 7;
  workbook.definedNames.push({ name: 'Boş', value: '', hidden: false }, { name: '', value: '1', hidden: false }, { name: 'Uzak', value: '1', localSheet: 9, hidden: false }, { name: 'Kalan', value: 'Gizli!$A$1', hidden: false });
  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, [
    'Workbook: no sheet is visible; applications refuse such a file',
    'Defined name: a name without a name or without a value is left out (2 times)',
    'Defined name: a name of a sheet the workbook does not have is left out',
  ]);
  assert.deepEqual(after.definedNames, [{ name: 'Kalan', value: 'Gizli!$A$1', localSheet: undefined, hidden: false }]);
  assert.deepEqual(after.sheets.map((sheet) => [sheet.name, sheet.sheetId, sheet.state]), [['Gizli', 5, 'hidden'], ['İkiz', 1, 'veryHidden']]);
  assert.equal(after.activeSheet, 0);
  assert.doesNotMatch(text(await pkg.readPart('/xl/workbook.xml')), /activeTab/);
});

test('buildXlsx: sheet names are written as they are; one an application refuses is reported', async () => {
  const names = ['Normal', 'normal', 'a/b', 'x'.repeat(32), ''];
  const { after, warnings } = await build(bookOf(...names.map((name, i) => sheetOf(name, i + 1))));
  assert.deepEqual(after.sheets.map((sheet) => sheet.name), names);
  assert.deepEqual(warnings, ['Workbook: a sheet has a name applications refuse (empty, longer than 31 characters, one of []:*?/\\ in it, or used twice) (4 times)']);
  // A shape whose text has no paragraphs is written without a text body, which has to have one.
  const sheet = sheetOf('Ş');
  sheet.drawings = [{ anchor: { kind: 'absolute', x: 0, y: 0, width: 9525, height: 9525 }, object: { type: 'shape', geometry: 'rect', text: { paragraphs: [] } } }];
  const built = await build(bookOf(sheet));
  assert.doesNotMatch(text(await built.pkg.readPart('/xl/drawings/drawing1.xml')), /txBody/);
});

test('buildXlsx: a model read from an xlsx file is written as what the model holds of it', async () => {
  const source = await (await open(await sampleWorkbook())).workbook();
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const bytes = await buildXlsx(source, { report });
  const after = await (await open(bytes)).workbook();
  assert.equal(await workbookDifference(source, after), undefined);
  assert.deepEqual(await buildXlsx(source), bytes);

  // What stands in parts the model does not hold is named.
  const withTable = await xlsxFixture({ sheets: [{ name: 'T', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId9"/></tableParts>') }] });
  const zipped = await OpcPackage.open(withTable);
  zipped.writePart('/xl/worksheets/_rels/sheet1.xml.rels', new TextEncoder().encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>'));
  const tabled = await (await open(await zipped.save())).workbook();
  const tableReport: XlsxBuildReport = { parts: [], warnings: [] };
  await buildXlsx(tabled, { report: tableReport });
  assert.deepEqual(tableReport.warnings, ['Sheet: "table" content of the source file is not in the model and is left out']);
});

test('buildXlsx: numbers a file cannot hold, and values of every kind', async () => {
  const sheet = sheetOf('Değer');
  const values: [Cell['kind'], Cell['value']][] = [['number', 0], ['number', -0], ['number', 1e308], ['number', 5e-324], ['number', 0.1 + 0.2], ['number', -123456789012345680000], ['boolean', true], ['boolean', false], ['error', '#REF!'], ['error', '#NAME?'], ['string', '0123'], ['blank', null]];
  values.forEach(([kind, value], i) => put(sheet, i, 0, { kind, value, style: 0 }));
  put(sheet, 20, 0, { kind: 'number', value: NaN });
  put(sheet, 21, 0, { kind: 'number', value: Infinity });
  const { after } = await build(bookOf(sheet));
  values.slice(0, 11).forEach(([kind, value], i) => {
    const cell = getCell(after.sheets[0], i, 0)!;
    assert.equal(cell.kind, kind, `row ${i}`);
    assert.ok(Object.is(cell.value, value) || (value === 0 && cell.value === 0), `row ${i}: ${cell.value}`);
  });
  // A blank cell without a format says nothing and is not written.
  assert.equal(getCell(after.sheets[0], 11, 0), undefined);
  assert.deepEqual([getCell(after.sheets[0], 20, 0), getCell(after.sheets[0], 21, 0)], [{ kind: 'error', value: '#NUM!', style: 0 }, { kind: 'error', value: '#NUM!', style: 0 }]);
});
