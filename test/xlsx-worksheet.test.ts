import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseRunFont, parseSharedStrings, parseStringItem, decodeXstring } from '../src/formats/xlsx/shared-strings';
import { parseWorksheet } from '../src/formats/xlsx/worksheet';
import { MAX_COLS, createSheet, getCell, type SharedString, type Sheet } from '../src/model/sheet/types';
import { XmlScanner } from '../src/xml-scan';

// --- helpers ---------------------------------------------------------------

const utf8 = (text: string) => new TextEncoder().encode(text);
const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const sst = (items: string) => utf8(`${HEADER}<sst ${NS} count="1" uniqueCount="1">${items}</sst>`);
const worksheet = (body: string) =>
  utf8(`${HEADER}<worksheet ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${body}</worksheet>`);

function emptySheet(): Sheet {
  return createSheet({ name: 'Sayfa1', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
}

function read(body: string, shared: SharedString[] = []): Sheet {
  const sheet = emptySheet();
  parseWorksheet(worksheet(body), shared, sheet);
  return sheet;
}

const data = (rows: string) => `<sheetData>${rows}</sheetData>`;

// --- shared strings ----------------------------------------------------------

test('shared strings: plain items keep their text and spaces', () => {
  const strings = parseSharedStrings(
    sst(
      '<si><t>Gemi</t></si>' +
        '<si><t xml:space="preserve">  baş ve son  </t></si>' +
        '<si><t>  korunur  </t></si>' +
        '<si><t/></si>' +
        '<si/>' +
        '<si><t>Ağrı İstanbul Şırnak çöğü 😀🚢 &amp; &lt;etiket&gt;</t></si>' +
        '<si><t>satır 1\nsatır 2</t></si>',
    ),
  );
  assert.deepEqual(strings, [
    { text: 'Gemi' },
    { text: '  baş ve son  ' },
    { text: '  korunur  ' },
    { text: '' },
    { text: '' },
    { text: 'Ağrı İstanbul Şırnak çöğü 😀🚢 & <etiket>' },
    { text: 'satır 1\nsatır 2' },
  ]);
});

test('shared strings: _xHHHH_ escapes are decoded', () => {
  assert.equal(decodeXstring('a_x000D__x000A_b'), 'a\r\nb');
  assert.equal(decodeXstring('_x005F_x000D_'), '_x000D_');
  assert.equal(decodeXstring('_x0009_tab _x00e7_'), '\ttab ç');
  assert.equal(decodeXstring('not_x12_an_xZZZZ_escape_x'), 'not_x12_an_xZZZZ_escape_x');
  const strings = parseSharedStrings(sst('<si><t>bir_x000D_iki</t></si><si><r><t>_x005F_x0041_</t></r></si>'));
  assert.equal(strings[0].text, 'bir\riki');
  assert.equal(strings[1].text, '_x0041_');
});

test('shared strings: rich text yields runs with fonts; phonetic text is left out', () => {
  const strings = parseSharedStrings(
    sst(
      '<si>' +
        '<r><t xml:space="preserve">Düz </t></r>' +
        '<r><rPr><b/><i val="1"/><strike val="0"/><u/><sz val="14"/><color rgb="ffff0000"/>' +
        '<rFont val="Arial Narrow"/><family val="2"/><charset val="162"/><scheme val="minor"/><vertAlign val="superscript"/>' +
        '<outline/></rPr><t>kalın</t></r>' +
        '<r><rPr><u val="double"/><color theme="4" tint="-0.249977111117893"/></rPr><t> çift</t></r>' +
        '<r><rPr><u val="none"/><b val="false"/><color indexed="10"/><vertAlign val="baseline"/></rPr><t>x</t></r>' +
        '<r><rPr><color auto="1"/></rPr><t>y</t></r>' +
        '<rPh sb="0" eb="1"><t>フリガナ</t></rPh><phoneticPr fontId="1"/>' +
        '</si>' +
        '<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1" type="noConversion"/></si>',
    ),
  );
  assert.equal(strings.length, 2);
  assert.equal(strings[0].text, 'Düz kalın çiftxy');
  assert.deepEqual(strings[0].runs, [
    { text: 'Düz ' },
    {
      text: 'kalın',
      font: {
        bold: true,
        italic: true,
        strike: false,
        underline: 'single',
        size: 14,
        color: { rgb: 'FFFF0000' },
        name: 'Arial Narrow',
        family: 2,
        charset: 162,
        scheme: 'minor',
        vertAlign: 'superscript',
      },
    },
    {
      text: ' çift',
      font: { bold: false, italic: false, strike: false, underline: 'double', color: { theme: 4, tint: -0.249977111117893 } },
    },
    { text: 'x', font: { bold: false, italic: false, strike: false, underline: 'none', color: { indexed: 10 } } },
    { text: 'y', font: { bold: false, italic: false, strike: false, underline: 'none', color: { auto: true } } },
  ]);
  assert.deepEqual(strings[1], { text: '東京' });
  assert.equal('runs' in strings[1], false);
});

test('shared strings: parseRunFont and parseStringItem work on a scanner positioned at the element', () => {
  const sc = new XmlScanner('<x><rPr><b/><sz val="9.5"/></rPr><is><t>a</t></is><end/></x>');
  sc.next();
  sc.next();
  assert.deepEqual(parseRunFont(sc), { bold: true, italic: false, strike: false, underline: 'none', size: 9.5 });
  sc.next();
  assert.deepEqual(parseStringItem(sc), { text: 'a' });
  sc.next();
  assert.equal(sc.name, 'end');

  const empty = new XmlScanner('<rPr/>');
  empty.next();
  assert.deepEqual(parseRunFont(empty), { bold: false, italic: false, strike: false, underline: 'none' });
});

test('shared strings: prefixed elements, empty table and unknown children', () => {
  assert.deepEqual(parseSharedStrings(utf8('<x:sst xmlns:x="urn:x"><x:si><x:t>önek</x:t></x:si><x:extLst><x:ext/></x:extLst></x:sst>')), [
    { text: 'önek' },
  ]);
  assert.deepEqual(parseSharedStrings(sst('')), []);
  assert.deepEqual(parseSharedStrings(utf8(`<sst ${NS}/>`)), []);
});

test('shared strings: malformed parts fail with a clear error', () => {
  assert.throws(() => parseSharedStrings(utf8('<sst><si><t>kesik')), /Invalid XML.*unclosed element/);
  assert.throws(() => parseSharedStrings(utf8('<sst><si><t>a</si></sst>')), /unexpected closing tag/);
  assert.throws(() => parseSharedStrings(utf8('<styleSheet/>')), /expected <sst>/);
  assert.throws(() => parseSharedStrings(utf8('')), /no root element/);
  assert.throws(() => parseSharedStrings(utf8('<!DOCTYPE sst><sst/>')), /DTD/);
  assert.throws(() => parseSharedStrings(utf8('<sst><si><t>&bilinmeyen;</t></si></sst>')), /Unknown XML entity/);
});

// --- worksheet: sheet-level properties ---------------------------------------------

test('worksheet: an empty sheet keeps the defaults', () => {
  for (const bytes of [worksheet(''), worksheet('<sheetData/>'), utf8(`<worksheet ${NS}/>`)]) {
    const sheet = emptySheet();
    parseWorksheet(bytes, [], sheet);
    assert.deepEqual(sheet, emptySheet());
    assert.equal(sheet.rowCount, 0);
    assert.equal(sheet.colCount, 0);
  }
});

test('worksheet: dimension, format defaults and tab colour', () => {
  const sheet = read(
    '<sheetPr codeName="Sayfa1"><tabColor rgb="ff00b050"/><pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<dimension ref="B2:D10"/>' +
      '<sheetFormatPr baseColWidth="10" defaultColWidth="12.5" defaultRowHeight="15.75" x14ac:dyDescent="0.25"/>' +
      '<sheetData/>',
  );
  assert.deepEqual(sheet.dimension, { top: 1, left: 1, bottom: 9, right: 3 });
  assert.equal(sheet.baseColWidth, 10);
  assert.equal(sheet.defaultColWidth, 12.5);
  assert.equal(sheet.defaultRowHeight, 15.75);
  assert.deepEqual(sheet.tabColor, { rgb: 'FF00B050' });

  const themed = read('<sheetPr><tabColor theme="5" tint="0.4"/></sheetPr><dimension ref="A1"/>');
  assert.deepEqual(themed.tabColor, { theme: 5, tint: 0.4 });
  assert.deepEqual(themed.dimension, { top: 0, left: 0, bottom: 0, right: 0 });

  const bad = read('<dimension ref="nonsense"/><sheetFormatPr defaultRowHeight="abc"/>');
  assert.equal(bad.dimension, undefined);
  assert.equal(bad.defaultRowHeight, undefined);
});

test('worksheet: sheet view flags come from the first view only', () => {
  const sheet = read(
    '<sheetViews>' +
      '<sheetView showGridLines="0" showRowColHeaders="false" showZeros="0" rightToLeft="1" tabSelected="1" zoomScale="85" workbookViewId="0">' +
      '<selection activeCell="C5" sqref="C5"/></sheetView>' +
      '<sheetView zoomScale="200" workbookViewId="1"><selection activeCell="Z9" sqref="Z9"/></sheetView>' +
      '</sheetViews>',
  );
  assert.deepEqual(sheet.view, {
    showGridLines: false,
    showRowColHeaders: false,
    showZeros: false,
    rightToLeft: true,
    tabSelected: true,
    zoomScale: 85,
    frozenRows: 0,
    frozenCols: 0,
    activeCell: 'C5',
  });

  const plain = read('<sheetViews><sheetView workbookViewId="0"/></sheetViews>');
  assert.deepEqual(plain.view, emptySheet().view);
  assert.equal(read('<sheetViews><sheetView zoomScale="9999"/></sheetViews>').view.zoomScale, 100);
});

test('worksheet: frozen panes', () => {
  const both = read(
    '<sheetViews><sheetView workbookViewId="0">' +
      '<pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>' +
      '<selection pane="topRight" activeCell="C1" sqref="C1"/>' +
      '<selection pane="bottomLeft" activeCell="A4" sqref="A4"/>' +
      '<selection pane="bottomRight" activeCell="E9" sqref="E9"/>' +
      '</sheetView></sheetViews>',
  );
  assert.equal(both.view.frozenCols, 2);
  assert.equal(both.view.frozenRows, 3);
  assert.equal(both.view.topLeftCell, 'C4');
  assert.equal(both.view.activeCell, 'E9');

  const rowsOnly = read(
    '<sheetViews><sheetView><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
      '<selection pane="bottomLeft" activeCell="B7" sqref="B7"/></sheetView></sheetViews>',
  );
  assert.deepEqual([rowsOnly.view.frozenRows, rowsOnly.view.frozenCols], [1, 0]);
  assert.equal(rowsOnly.view.topLeftCell, 'A2');
  assert.equal(rowsOnly.view.activeCell, 'B7');

  const colsOnly = read('<sheetViews><sheetView><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozenSplit"/></sheetView></sheetViews>');
  assert.deepEqual([colsOnly.view.frozenRows, colsOnly.view.frozenCols], [0, 1]);
  assert.equal(colsOnly.view.topLeftCell, 'B1');
  assert.equal(colsOnly.view.activeCell, undefined);

  // A split that is not frozen states its position in twips, not in cells.
  for (const state of ['state="split"', '']) {
    const split = read(
      `<sheetViews><sheetView><pane xSplit="2400" ySplit="1800" topLeftCell="D7" activePane="bottomRight" ${state}/>` +
        '<selection activeCell="A1"/><selection pane="bottomRight" activeCell="D8"/></sheetView></sheetViews>',
    );
    assert.deepEqual([split.view.frozenRows, split.view.frozenCols], [0, 0]);
    assert.equal(split.view.topLeftCell, undefined);
    assert.equal(split.view.activeCell, 'D8');
  }

  // No selection for the active pane: fall back to the first one.
  const fallback = read(
    '<sheetViews><sheetView><pane ySplit="1" activePane="bottomLeft" state="frozen"/><selection activeCell="B2"/></sheetView></sheetViews>',
  );
  assert.equal(fallback.view.activeCell, 'B2');
});

test('worksheet: column ranges are 0-based, sorted and never overlap', () => {
  const sheet = read(
    '<cols>' +
      '<col min="5" max="6" width="20.5" customWidth="1"/>' +
      '<col min="1" max="1" width="3" hidden="1" style="4" outlineLevel="2" collapsed="1"/>' +
      '<col min="3" max="8" width="9" bestFit="1"/>' +
      '<col min="6" max="5" width="1"/>' +
      '<col min="0" max="1" width="2"/>' +
      '<col max="4" width="2"/>' +
      '<col min="16380" max="99999" width="7"/>' +
      '<col min="20000" max="20001" width="7"/>' +
      '</cols>',
  );
  const plain = { customWidth: false, hidden: false, outlineLevel: 0, collapsed: false };
  assert.deepEqual(sheet.columns, [
    { min: 0, max: 0, width: 3, customWidth: false, hidden: true, style: 4, outlineLevel: 2, collapsed: true },
    { ...plain, min: 2, max: 3, width: 9 },
    { ...plain, min: 4, max: 5, width: 20.5, customWidth: true },
    { ...plain, min: 6, max: 7, width: 9 },
    { ...plain, min: 16379, max: 16383, width: 7 },
  ]);
  assert.equal(sheet.colCount, 0);
});

test('worksheet: a column range over the whole sheet stays one record', () => {
  const sheet = read(`<cols><col min="1" max="${MAX_COLS}" width="12" customWidth="1"/></cols><sheetData/>`);
  assert.equal(sheet.columns.length, 1);
  assert.deepEqual([sheet.columns[0].min, sheet.columns[0].max], [0, MAX_COLS - 1]);
  assert.equal(sheet.rows.size, 0);
});

test('worksheet: an absurd number of column records is refused', () => {
  const cols = '<col min="1" max="16384" width="1"/>'.repeat(MAX_COLS + 1);
  assert.throws(() => read(`<cols>${cols}</cols>`), /more than 16384 <col> elements/);
});

// --- worksheet: rows and cells -----------------------------------------------------

test('worksheet: row properties', () => {
  const sheet = read(
    data(
      '<row r="2" spans="1:3" ht="30" customHeight="1" x14ac:dyDescent="0.25"><c r="A2"><v>1</v></c></row>' +
        '<row r="4" hidden="1" outlineLevel="1" collapsed="1" s="7" customFormat="1"/>' +
        '<row r="5" s="7"/>' +
        '<row r="6" ht="12.75"><c r="C6" s="2"/></row>',
    ),
  );
  assert.deepEqual([...sheet.rows.keys()], [1, 3, 4, 5]);
  const r2 = sheet.rows.get(1)!;
  assert.deepEqual([r2.height, r2.customHeight, r2.hidden, r2.style, r2.outlineLevel, r2.collapsed], [30, true, false, undefined, 0, false]);
  const r4 = sheet.rows.get(3)!;
  assert.deepEqual([r4.height, r4.customHeight, r4.hidden, r4.style, r4.outlineLevel, r4.collapsed], [undefined, false, true, 7, 1, true]);
  assert.deepEqual(r4.cells, []);
  // "s" without customFormat is not a row style.
  assert.equal(sheet.rows.get(4)!.style, undefined);
  assert.equal(sheet.rows.get(5)!.height, 12.75);
  // Blank but formatted cells count towards the used area; empty rows do not.
  assert.equal(sheet.rowCount, 6);
  assert.equal(sheet.colCount, 3);
  assert.deepEqual(getCell(sheet, 5, 2), { kind: 'blank', value: null, style: 2 });
});

test('worksheet: every cell type', () => {
  const shared: SharedString[] = [{ text: 'Çıkış limanı' }, { text: 'kalın düz', runs: [{ text: 'kalın', font: { bold: true, italic: false, strike: false, underline: 'none' } }, { text: ' düz' }] }];
  const sheet = read(
    data(
      '<row r="1">' +
        '<c r="A1"><v>42</v></c>' +
        '<c r="B1" t="n" s="3"><v>-1.5E-3</v></c>' +
        '<c r="C1" t="s"><v>0</v></c>' +
        '<c r="D1" t="s" s="1"><v>1</v></c>' +
        '<c r="E1" t="str"><f>A1&amp;"ş"</f><v>42ş_x000A_</v></c>' +
        '<c r="F1" t="inlineStr"><is><t>satır içi 😀</t></is></c>' +
        '<c r="G1" t="inlineStr"><is><r><rPr><i/></rPr><t>eğik</t></r><r><t> düz</t></r></is></c>' +
        '<c r="H1" t="b"><v>1</v></c>' +
        '<c r="I1" t="b"><v>0</v></c>' +
        '<c r="J1" t="e"><f>1/0</f><v>#DIV/0!</v></c>' +
        '<c r="K1" t="d"><v>2026-10-01T12:30:00</v></c>' +
        '<c r="L1" s="5"/>' +
        '<c r="M1"><f>SUM(A1:B1)</f></c>' +
        '<c r="N1"><v></v></c>' +
        '<c r="O1" t="str"><v></v></c>' +
        '<c r="P1"><v> 7 </v></c>' +
        '<c r="Q1" t="s"/>' +
        '</row>',
    ),
    shared,
  );
  const at = (col: number) => getCell(sheet, 0, col);
  assert.deepEqual(at(0), { kind: 'number', value: 42, style: 0 });
  assert.deepEqual(at(1), { kind: 'number', value: -0.0015, style: 3 });
  assert.deepEqual(at(2), { kind: 'string', value: 'Çıkış limanı', style: 0 });
  assert.equal('runs' in at(2)!, false);
  assert.equal(at(3)!.value, 'kalın düz');
  assert.equal(at(3)!.style, 1);
  assert.equal(at(3)!.runs, shared[1].runs, 'runs are shared by reference');
  assert.deepEqual(at(4), { kind: 'string', value: '42ş\n', style: 0, formula: { text: 'A1&"ş"', type: 'normal' } });
  assert.deepEqual(at(5), { kind: 'string', value: 'satır içi 😀', style: 0 });
  assert.deepEqual(at(6), {
    kind: 'string',
    value: 'eğik düz',
    style: 0,
    runs: [{ text: 'eğik', font: { bold: false, italic: true, strike: false, underline: 'none' } }, { text: ' düz' }],
  });
  assert.deepEqual(at(7), { kind: 'boolean', value: true, style: 0 });
  assert.deepEqual(at(8), { kind: 'boolean', value: false, style: 0 });
  assert.deepEqual(at(9), { kind: 'error', value: '#DIV/0!', style: 0, formula: { text: '1/0', type: 'normal' } });
  assert.deepEqual(at(10), { kind: 'date', value: '2026-10-01T12:30:00', style: 0 });
  assert.deepEqual(at(11), { kind: 'blank', value: null, style: 5 });
  assert.deepEqual(at(12), { kind: 'blank', value: null, style: 0, formula: { text: 'SUM(A1:B1)', type: 'normal' } });
  assert.deepEqual(at(13), { kind: 'blank', value: null, style: 0 });
  assert.deepEqual(at(14), { kind: 'string', value: '', style: 0 });
  assert.deepEqual(at(15), { kind: 'number', value: 7, style: 0 });
  assert.deepEqual(at(16), { kind: 'blank', value: null, style: 0 });
  assert.equal(sheet.rowCount, 1);
  assert.equal(sheet.colCount, 17);
});

test('worksheet: shared, array and data table formulas', () => {
  const sheet = read(
    data(
      '<row r="1">' +
        '<c r="A1"><f t="shared" ref="A1:A3" si="0">B1*2</f><v>2</v></c>' +
        '<c r="B1"><f t="array" ref="B1:B2" aca="1" ca="1">TRANSPOSE(D1:E1)</f><v>1</v></c>' +
        '<c r="C1"><f t="dataTable" ref="C1:C2" dt2D="1" dtr="1" r1="A1" r2="B1"/><v>5</v></c>' +
        '<c r="D1"><f t="normal" xml:space="preserve"> 1+1</f><v>2</v></c>' +
        '<c r="E1"><f t="future">X()</f><v>2</v></c>' +
        '</row>' +
        '<row r="2"><c r="A2"><f t="shared" si="0"/><v>4</v></c></row>' +
        '<row r="3"><c r="A3"><f t="shared" si="0"></f><v>6</v></c></row>',
    ),
  );
  assert.deepEqual(getCell(sheet, 0, 0)!.formula, { text: 'B1*2', type: 'shared', ref: 'A1:A3', si: 0 });
  assert.deepEqual(getCell(sheet, 1, 0), { kind: 'number', value: 4, style: 0, formula: { text: '', type: 'shared', si: 0 } });
  assert.deepEqual(getCell(sheet, 2, 0)!.formula, { text: '', type: 'shared', si: 0 });
  assert.deepEqual(getCell(sheet, 0, 1)!.formula, {
    text: 'TRANSPOSE(D1:E1)',
    type: 'array',
    ref: 'B1:B2',
    attrs: new Map([
      ['aca', '1'],
      ['ca', '1'],
    ]),
  });
  assert.deepEqual(getCell(sheet, 0, 2)!.formula, {
    text: '',
    type: 'dataTable',
    ref: 'C1:C2',
    attrs: new Map([
      ['dt2D', '1'],
      ['dtr', '1'],
      ['r1', 'A1'],
      ['r2', 'B1'],
    ]),
  });
  assert.deepEqual(getCell(sheet, 0, 3)!.formula, { text: ' 1+1', type: 'normal', attrs: new Map([['xml:space', 'preserve']]) });
  // An unknown formula type is kept to be written back.
  assert.deepEqual(getCell(sheet, 0, 4)!.formula, { text: 'X()', type: 'normal', attrs: new Map([['t', 'future']]) });
});

test('worksheet: rows and cells without "r" follow the previous one', () => {
  const sheet = read(
    data(
      '<row><c><v>1</v></c><c><v>2</v></c></row>' +
        '<row><c><v>3</v></c></row>' +
        '<row r="10"><c r="D10"><v>4</v></c><c><v>5</v></c><c r="B10"><v>6</v></c><c><v>7</v></c></row>' +
        '<row><c><v>8</v></c></row>',
    ),
  );
  assert.deepEqual([...sheet.rows.keys()], [0, 1, 9, 10]);
  assert.equal(getCell(sheet, 0, 0)!.value, 1);
  assert.equal(getCell(sheet, 0, 1)!.value, 2);
  assert.equal(getCell(sheet, 1, 0)!.value, 3);
  assert.equal(getCell(sheet, 9, 3)!.value, 4);
  assert.equal(getCell(sheet, 9, 4)!.value, 5);
  assert.equal(getCell(sheet, 9, 1)!.value, 6);
  assert.equal(getCell(sheet, 9, 2)!.value, 7);
  assert.equal(getCell(sheet, 10, 0)!.value, 8);
  assert.equal(sheet.rowCount, 11);
  assert.equal(sheet.colCount, 5);
});

test('worksheet: a repeated cell or row keeps the last value', () => {
  const sheet = read(
    data(
      '<row r="1" ht="10"><c r="A1"><v>1</v></c><c r="A1" s="2"><v>2</v></c><c r="B1"><v>3</v></c></row>' +
        '<row r="1" ht="20"><c r="B1"><v>4</v></c></row>',
    ),
  );
  assert.equal(sheet.rows.size, 1);
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'number', value: 2, style: 2 });
  assert.equal(getCell(sheet, 0, 1)!.value, 4);
  assert.equal(sheet.rows.get(0)!.height, 20);
});

test('worksheet: cells at the sheet limits and absolute references', () => {
  const sheet = read(data('<row r="1048576"><c r="XFD1048576"><v>9</v></c><c r="$A$1048576"><v>1</v></c></row>'));
  assert.equal(getCell(sheet, 1048575, 16383)!.value, 9);
  assert.equal(getCell(sheet, 1048575, 0)!.value, 1);
  assert.equal(sheet.rowCount, 1048576);
  assert.equal(sheet.colCount, 16384);
});

test('worksheet: a prefixed document reads the same', () => {
  const sheet = emptySheet();
  parseWorksheet(
    utf8(
      '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:rel="urn:rel">' +
        '<x:dimension ref="A1:B1"/>' +
        '<x:sheetViews><x:sheetView workbookViewId="0"><x:pane ySplit="1" state="frozen"/></x:sheetView></x:sheetViews>' +
        '<x:cols><x:col min="2" max="2" width="30" customWidth="1"/></x:cols>' +
        '<x:sheetData><x:row r="1"><x:c r="A1" t="s"><x:v>0</x:v></x:c><x:c r="B1" t="inlineStr"><x:is><x:t>iç</x:t></x:is></x:c>' +
        '<x:c r="C1"><x:f>A1</x:f><x:v>3</x:v></x:c></x:row></x:sheetData>' +
        '<x:mergeCells count="1"><x:mergeCell ref="A2:B3"/></x:mergeCells>' +
        '<x:hyperlinks><x:hyperlink ref="A1" rel:id="rId7"/></x:hyperlinks>' +
        '</x:worksheet>',
    ),
    [{ text: 'önekli' }],
    sheet,
  );
  assert.deepEqual(sheet.dimension, { top: 0, left: 0, bottom: 0, right: 1 });
  assert.equal(sheet.view.frozenRows, 1);
  assert.equal(sheet.columns.length, 1);
  assert.equal(getCell(sheet, 0, 0)!.value, 'önekli');
  assert.equal(getCell(sheet, 0, 1)!.value, 'iç');
  assert.deepEqual(getCell(sheet, 0, 2), { kind: 'number', value: 3, style: 0, formula: { text: 'A1', type: 'normal' } });
  assert.deepEqual(sheet.merges, [{ top: 1, left: 0, bottom: 2, right: 1 }]);
  assert.deepEqual(sheet.hyperlinks, [{ range: { top: 0, left: 0, bottom: 0, right: 0 }, rId: 'rId7' }]);
});

test('worksheet: merges and hyperlinks', () => {
  const sheet = read(
    '<sheetData/>' +
      '<mergeCells count="4"><mergeCell ref="A1:C2"/><mergeCell ref="bozuk"/><mergeCell/><mergeCell ref="E5:D4"/><mergeCell ref="A1:XFE1"/></mergeCells>' +
      '<hyperlinks>' +
      '<hyperlink ref="A1" r:id="rId1" tooltip="Aç &amp; gör"/>' +
      '<hyperlink ref="B2:C3" location="Sayfa2!A1" display="İkinci sayfa"/>' +
      '<hyperlink ref="!!" r:id="rId2"/>' +
      '<hyperlink r:id="rId3"/>' +
      '</hyperlinks>',
  );
  assert.deepEqual(sheet.merges, [
    { top: 0, left: 0, bottom: 1, right: 2 },
    { top: 3, left: 3, bottom: 4, right: 4 },
  ]);
  assert.deepEqual(sheet.hyperlinks, [
    { range: { top: 0, left: 0, bottom: 0, right: 0 }, rId: 'rId1', tooltip: 'Aç & gör' },
    { range: { top: 1, left: 1, bottom: 2, right: 2 }, location: 'Sayfa2!A1', display: 'İkinci sayfa' },
  ]);
});

test('worksheet: unknown elements are skipped wherever they appear', () => {
  const sheet = read(
    '<sheetPr><outlinePr summaryBelow="0"/></sheetPr>' +
      '<sheetData><gelecek><row r="9"><c r="A9"><v>9</v></c></row></gelecek>' +
      '<row r="1"><yeni a="1"><c r="Z1"><v>1</v></c></yeni><c r="A1" cm="1" vm="2"><v>5</v><extLst><ext uri="x"><v>99</v></ext></extLst></c></row></sheetData>' +
      '<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>3</formula></cfRule></conditionalFormatting>' +
      '<dataValidations count="1"><dataValidation type="list" sqref="B1"><formula1>"a,b"</formula1></dataValidation></dataValidations>' +
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
      '<drawing r:id="rId1"/><legacyDrawing r:id="rId2"/>' +
      '<extLst><ext uri="{x}"><x14:sheetData xmlns:x14="urn:x14"><x14:row r="50"><x14:c r="A50"><x14:v>1</x14:v></x14:c></x14:row></x14:sheetData></ext></extLst>',
  );
  assert.equal(sheet.rows.size, 1);
  assert.deepEqual(getCell(sheet, 0, 0), { kind: 'number', value: 5, style: 0 });
  assert.equal(sheet.rowCount, 1);
  assert.equal(sheet.colCount, 1);
});

test('worksheet: pretty-printed parts read the same', () => {
  const sheet = read(
    '\n  <sheetData>\n    <row r="1">\n      <c r="A1" t="inlineStr">\n        <is>\n          <t> a </t>\n        </is>\n      </c>\n' +
      '      <c r="B1">\n        <v>12</v>\n      </c>\n    </row>\n  </sheetData>\n',
  );
  assert.equal(getCell(sheet, 0, 0)!.value, ' a ');
  assert.equal(getCell(sheet, 0, 1)!.value, 12);
});

// --- worksheet: malformed input ------------------------------------------------------

test('worksheet: malformed parts fail with a clear error', () => {
  const fails = (body: string, pattern: RegExp, shared: SharedString[] = []) => assert.throws(() => read(body, shared), pattern, body);

  // Out-of-range and malformed references.
  fails(data('<row r="1048577"><c r="A1048577"><v>1</v></c></row>'), /row 1048577 is past the last row/);
  fails(data('<row r="0"/>'), /row number "0"/);
  fails(data('<row r="-3"/>'), /row number "-3"/);
  fails(data('<row r="1e3"/>'), /row number "1e3"/);
  fails(data('<row r="99999999999999999999"/>'), /row number/);
  fails(data('<row r="1048576"/><row/>'), /past the last row/);
  fails(data('<row r="1"><c r="XFE1"><v>1</v></c></row>'), /cell reference "XFE1"/);
  fails(data('<row r="1"><c r="A0"><v>1</v></c></row>'), /cell reference "A0"/);
  fails(data('<row r="1"><c r="A1048577"><v>1</v></c></row>'), /cell reference "A1048577"/);
  fails(data('<row r="1"><c r="11"><v>1</v></c></row>'), /cell reference "11"/);
  fails(data('<row r="1"><c r=""><v>1</v></c></row>'), /cell reference ""/);
  fails(data('<row r="1"><c r="XFD1"/><c/></row>'), /past the last column/);

  // Shared string indexes.
  fails(data('<row r="1"><c r="A1" t="s"><v>2</v></c></row>'), /shared string index "2" in cell A1 is out of range \(2 strings\)/, [{ text: 'a' }, { text: 'b' }]);
  fails(data('<row r="1"><c r="A1" t="s"><v>0</v></c></row>'), /shared string index "0" in cell A1 is out of range \(0 strings\)/);
  fails(data('<row r="1"><c r="B3" t="s"><v>-1</v></c></row>'), /shared string index "-1"/, [{ text: 'a' }]);
  fails(data('<row r="1"><c r="A1" t="s"><v>1.5</v></c></row>'), /shared string index "1.5"/, [{ text: 'a' }, { text: 'b' }]);

  // Numbers.
  fails(data('<row r="7"><c r="C7"><v>12,5</v></c></row>'), /"12,5" in cell C7 is not a number/);
  fails(data('<row r="1"><c r="A1"><v>abc</v></c></row>'), /"abc" in cell A1 is not a number/);
  fails(data('<row r="1"><c r="A1"><v>NaN</v></c></row>'), /not a number/);
  fails(data('<row r="1"><c r="A1"><v>Infinity</v></c></row>'), /not a number/);
  fails(data('<row r="1"><c r="A1"><v>1e999</v></c></row>'), /not a number/);

  // Broken XML.
  const bytes = (xml: string) => () => parseWorksheet(utf8(xml), [], emptySheet());
  assert.throws(bytes('<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>'), /Invalid XML.*unclosed element <row>/);
  assert.throws(bytes('<worksheet><sheetData><row r="1"><c r="A1"><v>1</c></row></sheetData></worksheet>'), /unexpected closing tag <\/c>/);
  assert.throws(bytes('<worksheet><sheetData><row r="1"><c r="A1'), /Invalid XML/);
  assert.throws(bytes('<worksheet><conditionalFormatting><cfRule>'), /unclosed element/);
  assert.throws(bytes('<worksheet/><worksheet/>'), /more than one root element/);
  assert.throws(bytes('<!DOCTYPE worksheet [<!ENTITY a "b">]><worksheet/>'), /DTD/);
  assert.throws(bytes('<worksheet><sheetData><row><c><v>&lol;</v></c></row></sheetData></worksheet>'), /Unknown XML entity/);
  assert.throws(bytes(''), /no root element/);
  assert.throws(bytes('<styleSheet/>'), /root element is <styleSheet>, expected <worksheet>/);
  assert.throws(() => parseWorksheet(new Uint8Array([0x50, 0x4b, 3, 4, 0, 0, 0xff]), [], emptySheet()), /Invalid XML/);
});

test('worksheet: every truncation of a part fails instead of hanging', () => {
  const xml =
    `<worksheet ${NS}><dimension ref="A1:B2"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/>` +
    '<selection activeCell="A1"/></sheetView></sheetViews><cols><col min="1" max="2" width="9"/></cols>' +
    '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>x</t></r></is></c><c r="B1"><f t="shared" si="0" ref="B1:B2">A1</f><v>1</v></c></row></sheetData>' +
    '<mergeCells><mergeCell ref="A1:B1"/></mergeCells><hyperlinks><hyperlink ref="A1" location="A2"/></hyperlinks></worksheet>';
  parseWorksheet(utf8(xml), [], emptySheet());
  for (let length = 0; length < xml.length; length++) {
    assert.throws(() => parseWorksheet(utf8(xml.slice(0, length)), [], emptySheet()), /Invalid XML/, `length ${length}`);
  }
});

// --- speed ------------------------------------------------------------------

test('worksheet: 100,000 rows x 10 columns read in reasonable time', () => {
  const rowCount = 100000;
  const shared: SharedString[] = [];
  for (let i = 0; i < 1000; i++) shared.push({ text: `Liman ${i} İskenderun` });

  let started = Date.now();
  const parts: string[] = [`${HEADER}<worksheet ${NS}><dimension ref="A1:J${rowCount}"/><sheetData>`];
  for (let row = 1; row <= rowCount; row++) {
    let line = `<row r="${row}" spans="1:10">`;
    line += `<c r="A${row}" s="1"><v>${row}</v></c>`;
    line += `<c r="B${row}" t="s"><v>${row % 1000}</v></c>`;
    line += `<c r="C${row}" s="2"><v>${row * 1.25}</v></c>`;
    line += `<c r="D${row}"><f>A${row}*C${row}</f><v>${row * row * 1.25}</v></c>`;
    line += `<c r="E${row}" t="b"><v>${row % 2}</v></c>`;
    line += `<c r="F${row}" t="inlineStr"><is><t>satır ${row}</t></is></c>`;
    line += `<c r="G${row}" s="4"><v>45000.${row % 100}</v></c>`;
    line += `<c r="H${row}" t="str"><f>"x"&amp;A${row}</f><v>x${row}</v></c>`;
    line += `<c r="I${row}" s="3"/>`;
    line += `<c r="J${row}"><v>-${row}E-2</v></c>`;
    parts.push(line + '</row>');
  }
  parts.push('</sheetData></worksheet>');
  const bytes = utf8(parts.join(''));
  parts.length = 0;
  const buildMs = Date.now() - started;

  const sheet = emptySheet();
  started = Date.now();
  parseWorksheet(bytes, shared, sheet);
  const elapsed = Date.now() - started;
  console.log(`worksheet: ${(bytes.length / 1e6).toFixed(1)} MB, ${rowCount * 10} cells parsed in ${elapsed} ms (fixture built in ${buildMs} ms)`);

  assert.equal(sheet.rows.size, rowCount);
  assert.equal(sheet.rowCount, rowCount);
  assert.equal(sheet.colCount, 10);
  assert.deepEqual(sheet.dimension, { top: 0, left: 0, bottom: rowCount - 1, right: 9 });
  for (const row of [1, 2, 777, 50000, rowCount]) {
    const r = row - 1;
    assert.deepEqual(getCell(sheet, r, 0), { kind: 'number', value: row, style: 1 });
    assert.deepEqual(getCell(sheet, r, 1), { kind: 'string', value: `Liman ${row % 1000} İskenderun`, style: 0 });
    assert.deepEqual(getCell(sheet, r, 2), { kind: 'number', value: row * 1.25, style: 2 });
    assert.deepEqual(getCell(sheet, r, 3), {
      kind: 'number',
      value: row * row * 1.25,
      style: 0,
      formula: { text: `A${row}*C${row}`, type: 'normal' },
    });
    assert.deepEqual(getCell(sheet, r, 4), { kind: 'boolean', value: row % 2 === 1, style: 0 });
    assert.deepEqual(getCell(sheet, r, 5), { kind: 'string', value: `satır ${row}`, style: 0 });
    assert.deepEqual(getCell(sheet, r, 6), { kind: 'number', value: Number(`45000.${row % 100}`), style: 4 });
    assert.deepEqual(getCell(sheet, r, 7), { kind: 'string', value: `x${row}`, style: 0, formula: { text: `"x"&A${row}`, type: 'normal' } });
    assert.deepEqual(getCell(sheet, r, 8), { kind: 'blank', value: null, style: 3 });
    assert.deepEqual(getCell(sheet, r, 9), { kind: 'number', value: Number(`-${row}E-2`), style: 0 });
  }
  assert.ok(elapsed < 10000, `took ${elapsed} ms`);
});
