// Saving an xls file as xlsx: the model read from the xls stream is written as a
// new package, which has to read back as the same model. Every xls file is built
// here, record by record.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { png } from '../e2e/fixtures/docx';
import { buildXlsx, open, type XlsxBuildReport } from '../src/index';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { resolveCellStyle } from '../src/model/sheet/color';
import { getCell, type DrawingGroup, type DrawingPicture, type DrawingShape, type Workbook } from '../src/model/sheet/types';
import { OpcPackage } from '../src/opc';
import { anchorRect } from '../src/render/sheet-drawings';
import { bse, clientAnchor, childAnchor, drawingGroup, drawingRecords, groupOf, note, shape, txo } from './xls-art-fixtures';
import {
  arrayRec,
  blank,
  boolErr,
  cat,
  definedName,
  defaultStyles,
  f64,
  font,
  format,
  formula,
  labelSst,
  number,
  ptg,
  rec,
  rk,
  rkInt,
  shrFmla,
  sst,
  str16,
  tokens,
  u16,
  u32,
  wide,
  workbookStream,
  xf,
  xlsFile,
  type Bytes,
  type WorkbookSpec,
} from './xls-fixtures';
import { canonical, mediaDigests, workbookContent, workbookDifference } from './workbook-content';

const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

/** A compound file around the stream, of whatever size the stream needs. */
const fileOf = (spec: WorkbookSpec) => xlsFile(workbookStream(spec));

interface Converted {
  before: Workbook;
  after: Workbook;
  bytes: Uint8Array;
  pkg: OpcPackage;
}

/** xls -> saveAsXlsx() -> open() -> workbook(). */
async function convert(spec: WorkbookSpec): Promise<Converted> {
  const file = await open(fileOf(spec));
  assert.equal(file.format, 'xls');
  const before = await file.workbook();
  const bytes = await file.saveAsXlsx!();
  const reopened = await open(bytes);
  assert.equal(reopened.format, 'xlsx');
  return { before, after: await reopened.workbook(), bytes, pkg: reopened.package! };
}

const row = (index: number, height: number, flags: number, xfIndex = 0) => rec(0x0208, u16(index), u16(0), u16(0), u16(height), u16(0), u16(0), u16(flags | 0x0100), u16(xfIndex));
const colInfo = (first: number, last: number, width: number, xfIndex = 15, flags = 0) => rec(0x007d, u16(first), u16(last), u16(width), u16(xfIndex), u16(flags), u16(0));
const window1 = (activeTab: number) => rec(0x003d, u16(0), u16(0), u16(100), u16(100), u16(0x38), u16(activeTab), u16(0), u16(1), u16(600));
const window2 = (flags: number, top = 0, left = 0) => rec(0x023e, u16(flags), u16(top), u16(left), u16(64), u16(0), u16(0), u16(0), u32(0));
const selection = (pane: number, r: number, c: number) => rec(0x001d, [pane], u16(r), u16(c), u16(0), u16(1), u16(r), u16(r), [c, c]);

const STD_LINK = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const URL_MONIKER = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const linkString = (value: string) => [...u32(value.length + 1), ...wide(value), 0, 0];
const hlink = (top: number, bottom: number, left: number, right: number, flags: number, ...body: Bytes[]) => rec(0x01b8, u16(top), u16(bottom), u16(left), u16(right), STD_LINK, u32(2), u32(flags), ...body);
const urlLink = (r: number, c: number, url: string, flags = 0, ...more: Bytes[]) => hlink(r, r, c, c, 0x01 | 0x02 | flags, ...more.slice(0, flags & 0x10 ? 1 : 0), URL_MONIKER, [...u32(url.length * 2 + 2), ...wide(url), 0, 0], ...more.slice(flags & 0x10 ? 1 : 0));

/** A string of the shared string table; `runs` are pairs of first character and font index. */
function sstString(value: string, runs?: [number, number][]): number[] {
  if (!runs) return [...u16(value.length), 1, ...wide(value)];
  return [...u16(value.length), 1 | 0x08, ...u16(runs.length), ...wide(value), ...runs.flatMap(([at, fontIndex]) => [...u16(at), ...u16(fontIndex)])];
}

/** Fonts 0..3 plain, 5 bold Tahoma, 6 red italic; XF 15 default, then the formats the tests use. */
function styledHead(extra: Bytes = []): Uint8Array {
  return cat(
    extra,
    font(),
    font(),
    font(),
    font(),
    font({ bold: true, name: 'Tahoma', size: 9 }),
    font({ italic: true, color: 10, underline: 1, strike: true, size: 14, name: 'Çağrı Yazı' }),
    format(164, '#,##0.00" TL"'),
    format(165, 'dd.mm.yyyy'),
    ...Array.from({ length: 15 }, () => xf({ style: true })),
    xf(), // 15
    xf({ font: 5 }), // 16: bold
    xf({ format: 164, horizontal: 3 }), // 17: money, right
    xf({ format: 165 }), // 18: date
    xf({ font: 6, horizontal: 2, vertical: 1, wrap: true, lines: [1, 2, 5, 6], lineColors: [10, 12, 8, 17], pattern: 1, fg: 13, bg: 9 }), // 19
    xf({ locked: false, hidden: true, rotation: 45, indent: 2, shrink: true, pattern: 4, fg: 22, bg: 23, diagonal: { style: 1, color: 10, up: true } }), // 20
    xf({ format: 9 }), // 21: percent (built in)
  );
}

function richSheet(): Uint8Array {
  const sharedTokens = tokens(ptg.refN(0, 0xff), ptg.int(2), ptg.mul);
  return cat(
    rec(0x0055, u16(10)),
    rec(0x0225, u16(0), u16(300)),
    colInfo(0, 0, 4736, 15, 0x02),
    colInfo(2, 3, 2304, 16, 0x01),
    colInfo(5, 7, 1024, 15, 0x0200 | 0x1000),
    row(0, 480, 0x40),
    row(1, 300, 0x20),
    row(6, 300, 0x80 | 0x10 | 0x02, 19),
    labelSst(0, 0, 0, 16),
    labelSst(0, 1, 1),
    labelSst(0, 2, 2, 19),
    labelSst(0, 4, 3),
    number(1, 0, 1234.5, 17),
    rk(1, 1, rkInt(-7)),
    number(1, 2, 45000, 18),
    number(1, 3, 0.25, 21),
    boolErr(2, 0, 1, false),
    boolErr(2, 1, 0, false),
    boolErr(2, 2, 0x07, true),
    boolErr(2, 3, 0x2a, true, 19),
    blank(2, 5, 20),
    formula(3, 0, tokens(ptg.ref(1, 0), ptg.ref(1, 1), ptg.add), 1227.5, 17),
    formula(3, 1, tokens(ptg.str('a<b & "c"'), ptg.ref(0, 1), ptg.concat), { text: 'a<b & "c"Şişli' }),
    formula(3, 2, tokens(ptg.ref(1, 0), ptg.int(0), ptg.gt), true),
    formula(3, 3, tokens(ptg.int(1), ptg.int(0), ptg.div), { error: 0x07 }),
    formula(3, 4, tokens(ptg.str('')), 'emptyText'),
    // A shared formula down three cells, and an array formula over two.
    formula(4, 1, ptg.exp(4, 1), 2),
    shrFmla(4, 6, 1, 1, sharedTokens),
    formula(5, 1, ptg.exp(4, 1), 4),
    formula(6, 1, ptg.exp(4, 1), 6),
    formula(8, 3, ptg.exp(8, 3), 2),
    arrayRec(8, 9, 3, 3, tokens(ptg.area(0, 1, 0, 0), ptg.int(2), ptg.mul)),
    formula(9, 3, ptg.exp(8, 3), 4),
    number(20, 255, 9),
    rec(0x00e5, u16(2), u16(10), u16(11), u16(0), u16(2), u16(12), u16(12), u16(4), u16(5)),
    urlLink(0, 1, 'https://example.com/çay?a=1&b=<2>', 0x10, linkString('Örnek "bağlantı"')),
    hlink(10, 11, 0, 2, 0x08, linkString("'Gizli & <Özel>'!A1")),
    rec(0x0800, u16(0x0800), u16(10), u16(11), u16(0), u16(2), wide('İpucu & <ek>'), [0, 0]),
    window2(0x0008 | 0x0100 | 0x0200 | 0x0004 | 0x0010 | 0x0002, 0, 0),
    rec(0x00a0, u16(3), u16(2)),
    rec(0x0041, u16(1), u16(2), u16(9), u16(5), [0]),
    selection(3, 0, 0),
    selection(0, 11, 6),
  );
}

function richBook(): WorkbookSpec {
  return {
    head: styledHead(cat(rec(0x0022, u16(1)), window1(1), rec(0x0092, u16(3), [0x11, 0x22, 0x33, 0], [0xaa, 0xbb, 0xcc, 0], [0x01, 0x02, 0x03, 0]))),
    tail: cat(
      definedName({ name: 'Toplam', tokens: tokens(ptg.ref3d(0, 3, 0)) }),
      definedName({ name: '', builtin: 6, sheet: 1, tokens: tokens(ptg.area3d(0, 0, 9, 0, 3)) }),
      definedName({ name: 'Gizli_Ad', hidden: true, tokens: tokens(ptg.int(42)) }),
      rec(0x01ae, u16(3), [0x01, 0x04]),
      rec(0x0017, u16(1), u16(0), u16(0), u16(0)),
      sst(sstString('Başlık'), sstString('Şişli'), sstString('kalın ve ince', [[0, 5], [5, 6]]), sstString('  boşluklu \n satır  ')),
    ),
    sheets: [
      { name: 'Veri', records: richSheet() },
      { name: 'Gizli & <Özel>', visibility: 1, records: cat(number(0, 0, 1), window2(0x0200 | 0x0002 | 0x0004 | 0x0010 | 0x0040), rec(0x00a0, u16(1), u16(2))) },
      { name: "Çok 'gizli'", visibility: 2, records: number(2, 2, -0.5) },
    ],
  };
}

test('xls to xlsx: cells, formulas, formats, sizes, panes, links and names read back as they were', async () => {
  const { before, after, pkg } = await convert(richBook());
  assert.equal(await workbookDifference(before, after), undefined);

  // The comparison says the two models agree; these say what they agree on.
  assert.deepEqual(after.sheets.map((sheet) => [sheet.name, sheet.state, sheet.kind]), [
    ['Veri', 'visible', 'worksheet'],
    ['Gizli & <Özel>', 'hidden', 'worksheet'],
    ["Çok 'gizli'", 'veryHidden', 'worksheet'],
  ]);
  // The xls file opens on its second sheet, which is hidden: the new file opens on the first one that is shown.
  assert.equal(before.activeSheet, 1);
  assert.equal(after.activeSheet, 0);
  assert.equal(after.date1904, true);
  const sheet = after.sheets[0];
  const cell = (r: number, c: number) => getCell(sheet, r, c)!;
  assert.deepEqual([cell(0, 0).kind, cell(0, 0).value], ['string', 'Başlık']);
  assert.equal(cell(0, 4).value, '  boşluklu \n satır  ');
  assert.deepEqual(cell(0, 2).runs?.map((run) => [run.text, run.font?.bold, run.font?.name]), [['kalın', true, 'Tahoma'], [' ve ince', false, 'Çağrı Yazı']]);
  assert.deepEqual([cell(1, 0).value, cell(1, 1).value, cell(1, 2).value, cell(1, 3).value], [1234.5, -7, 45000, 0.25]);
  assert.deepEqual([cell(2, 0).value, cell(2, 1).value, cell(2, 2).value, cell(2, 3).value], [true, false, '#DIV/0!', '#N/A']);
  assert.deepEqual([cell(2, 0).kind, cell(2, 2).kind, cell(2, 5).kind], ['boolean', 'error', 'blank']);
  assert.deepEqual([cell(3, 0).formula?.text, cell(3, 0).value], ['A2+B2', 1227.5]);
  assert.deepEqual([cell(3, 1).formula?.text, cell(3, 1).value, cell(3, 1).kind], ['"a<b & ""c"""&B1', 'a<b & "c"Şişli', 'string']);
  assert.deepEqual([cell(3, 2).value, cell(3, 3).value, cell(3, 4).value], [true, '#DIV/0!', '']);
  // The shared formula stays shared: one text, the range, and cells that point to it.
  assert.deepEqual(cell(4, 1).formula, { text: 'A5*2', type: 'shared', ref: 'B5:B7', si: 0 });
  assert.deepEqual(cell(5, 1).formula, { text: '', type: 'shared', si: 0 });
  assert.deepEqual([cell(4, 1).value, cell(5, 1).value, cell(6, 1).value], [2, 4, 6]);
  assert.deepEqual(cell(8, 3).formula, { text: 'A1:A2*2', type: 'array', ref: 'D9:D10' });
  assert.equal(cell(9, 3).formula, undefined);
  assert.equal(cell(20, 255).value, 9);

  const styles = after.styles;
  const money = resolveCellStyle(styles, cell(1, 0).style);
  assert.equal(money.numFmt, '#,##0.00" TL"');
  assert.equal(money.alignment.horizontal, 'right');
  const fancy = resolveCellStyle(styles, cell(0, 2).style);
  assert.deepEqual([fancy.font.name, fancy.font.size, fancy.font.italic, fancy.font.strike, fancy.font.underline], ['Çağrı Yazı', 14, true, true, 'single']);
  assert.deepEqual([fancy.alignment.horizontal, fancy.alignment.vertical, fancy.alignment.wrapText], ['center', 'center', true]);
  assert.deepEqual([fancy.border.left?.style, fancy.border.right?.style, fancy.border.top?.style, fancy.border.bottom?.style], ['thin', 'medium', 'thick', 'double']);
  assert.equal(fancy.fill.type === 'pattern' && fancy.fill.pattern, 'solid');
  const guarded = resolveCellStyle(styles, cell(2, 5).style);
  assert.deepEqual(styles.cellXfs[cell(2, 5).style].protection, { locked: false, hidden: true });
  assert.deepEqual([guarded.alignment.textRotation, guarded.alignment.indent, guarded.alignment.shrinkToFit], [45, 2, true]);
  assert.equal(guarded.border.diagonalUp, true);
  // The palette of the file is the palette of the new one.
  assert.deepEqual(styles.indexedColors?.slice(8, 11), ['112233', 'AABBCC', '010203']);

  assert.deepEqual(sheet.merges, [
    { top: 10, left: 0, bottom: 11, right: 2 },
    { top: 12, left: 4, bottom: 12, right: 5 },
  ]);
  assert.deepEqual(sheet.view, { showGridLines: true, showRowColHeaders: true, showZeros: true, rightToLeft: false, tabSelected: true, zoomScale: 150, frozenRows: 2, frozenCols: 1, topLeftCell: 'F10', activeCell: 'G12' });
  assert.equal(after.sheets[1].view.rightToLeft, true);
  assert.equal(after.sheets[1].view.zoomScale, 50);
  assert.deepEqual(sheet.columns.map((range) => [range.min, range.max, range.width, range.hidden, range.customWidth, range.outlineLevel, range.collapsed]), [
    [0, 0, 18.5, false, true, 0, false],
    [2, 3, 9, true, false, 0, false],
    [5, 7, 4, false, false, 2, true],
  ]);
  assert.deepEqual([sheet.rows.get(0)?.height, sheet.rows.get(0)?.customHeight, sheet.rows.get(1)?.hidden, sheet.rows.get(6)?.outlineLevel, sheet.rows.get(6)?.collapsed], [24, true, true, 2, true]);
  assert.deepEqual([sheet.baseColWidth, sheet.defaultRowHeight], [10, 15]);
  assert.deepEqual(
    sheet.hyperlinks.map((link) => [link.range, link.target, link.location, link.display, link.tooltip]),
    [
      [{ top: 0, left: 1, bottom: 0, right: 1 }, 'https://example.com/çay?a=1&b=<2>', undefined, 'Örnek "bağlantı"', undefined],
      [{ top: 10, left: 0, bottom: 11, right: 2 }, undefined, "'Gizli & <Özel>'!A1", undefined, 'İpucu & <ek>'],
    ],
  );
  assert.deepEqual(after.definedNames, [
    { name: 'Toplam', value: 'Veri!$A$4', localSheet: undefined, hidden: false },
    { name: '_xlnm.Print_Area', value: 'Veri!$A$1:$D$10', localSheet: 0, hidden: false },
    { name: 'Gizli_Ad', value: '42', localSheet: undefined, hidden: true },
  ]);

  // The shared strings of the file keep their indexes; the shared formula is written as one.
  assert.equal(cell(0, 4).si, 3);
  const sheetXml = text(await pkg.readPart('/xl/worksheets/sheet1.xml'));
  assert.match(sheetXml, /<c r="B5"><f t="shared" ref="B5:B7" si="0">A5\*2<\/f><v>2<\/v><\/c>/);
  assert.match(sheetXml, /<c r="B6"><f t="shared" si="0"\/><v>4<\/v><\/c>/);
  // Stored results are kept as they are: nothing asks for a recalculation.
  assert.doesNotMatch(text(await pkg.readPart('/xl/workbook.xml')), /calcPr/);
});

test('xls to xlsx: the file is a package an application can open', async () => {
  const { pkg } = await convert(richBook());
  assert.deepEqual(pkg.partNames(), [
    '/_rels/.rels',
    '/xl/workbook.xml',
    '/xl/_rels/workbook.xml.rels',
    '/xl/worksheets/sheet1.xml',
    '/xl/worksheets/_rels/sheet1.xml.rels',
    '/xl/worksheets/sheet2.xml',
    '/xl/worksheets/sheet3.xml',
    '/xl/styles.xml',
    '/xl/sharedStrings.xml',
  ]);
  assert.equal(pkg.contentTypeOf('/xl/workbook.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
  assert.equal(pkg.contentTypeOf('/xl/worksheets/sheet2.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
  assert.equal(pkg.contentTypeOf('/xl/styles.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  assert.equal(pkg.contentTypeOf('/xl/sharedStrings.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml');
  // Every part is well-formed XML without a document type.
  for (const part of pkg.partNames()) {
    const xml = text(await pkg.readPart(part));
    assert.doesNotMatch(xml, /<!DOCTYPE/i, part);
    await pkg.readXml(part);
  }
  // The link is a relationship to somewhere outside, escaped once.
  const rels = await pkg.relationships('/xl/worksheets/sheet1.xml');
  assert.deepEqual(rels.map((rel) => [rel.target, rel.external]), [['https://example.com/çay?a=1&b=<2>', true]]);
});

test('xls to xlsx: the same file gives the same bytes, and they survive a save untouched', async () => {
  const first = await convert(richBook());
  const second = await convert(richBook());
  assert.deepEqual(first.bytes, second.bytes);
  // Written twice from one model as well.
  assert.deepEqual(await buildXlsx(first.before), first.bytes);

  // open() -> save() -> open(): nothing was changed, so nothing is written again.
  const reopened = await open(first.bytes);
  await reopened.workbook();
  const saved = await reopened.save();
  assert.deepEqual(saved, first.bytes);
  assert.equal(canonical(workbookContent(await (await open(saved)).workbook())), canonical(workbookContent(first.after)));
  // For an xlsx, saveAsXlsx() is save().
  assert.deepEqual(await reopened.saveAsXlsx!(), first.bytes);
});

test('xls to xlsx: save() still gives the bytes of the xls file', async () => {
  const bytes = fileOf(richBook());
  const file = await open(bytes);
  await file.saveAsXlsx!();
  assert.deepEqual(await file.save(), bytes);
});

test('xls to xlsx: a chart sheet and a macro sheet keep their names and places as empty worksheets', async () => {
  const spec: WorkbookSpec = {
    tail: cat(definedName({ name: 'Yerel', sheet: 3, tokens: tokens(ptg.int(1)) })),
    sheets: [
      { name: 'Grafik', type: 2 },
      { name: 'Makro', type: 1, records: number(0, 0, 5) },
      { name: 'Veri', records: cat(number(0, 0, 5), rec(0x0862, u16(0x0862), new Array(10).fill(0), u32(20), u32(10))) },
    ],
  };
  const file = await open(fileOf(spec));
  const before = await file.workbook();
  assert.deepEqual(before.sheets.map((sheet) => sheet.kind), ['chartsheet', 'macrosheet', 'worksheet']);
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const after = await (await open(await buildXlsx(before, { report }))).workbook();
  assert.deepEqual(after.sheets.map((sheet) => [sheet.name, sheet.kind, sheet.rowCount]), [['Grafik', 'worksheet', 0], ['Makro', 'worksheet', 0], ['Veri', 'worksheet', 1]]);
  // The name stays local to the third sheet.
  assert.deepEqual(after.definedNames, [{ name: 'Yerel', value: '1', localSheet: 2, hidden: false }]);
  assert.deepEqual(after.sheets[2].tabColor, { indexed: 10 });
  assert.deepEqual(report.warnings, [
    'Sheet: the content of a chart sheet is not in the model; it is written as an empty worksheet',
    // The chart sheet of the fixture holds no records of a chart.
    'Sheet "Grafik" was read with a warning: The chart of the sheet is not shown: it holds no Chart record',
    'Sheet: the content of a macro sheet is not in the model; it is written as an empty worksheet',
  ]);
  assert.equal(await workbookDifference(before, after), undefined);
});

test('xls to xlsx: a link without a target is left out and reported', async () => {
  const unknown = hlink(0, 0, 0, 0, 0x01 | 0x10 | 0x08, linkString('Görünen'), new Array(16).fill(9), [1, 2, 3, 4]);
  const before = await (await open(fileOf({ sheets: [{ name: 'S', records: cat(number(0, 0, 1), unknown) }] }))).workbook();
  assert.equal(before.sheets[0].hyperlinks.length, 1);
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const after = await (await open(await buildXlsx(before, { report }))).workbook();
  assert.equal(after.sheets[0].hyperlinks.length, 0);
  assert.deepEqual(report.warnings, ['Hyperlink: a link without a target (or outside the sheet) is left out']);
});

test('xls to xlsx: a cell of a shared formula outside the range of its group gets its own text', async () => {
  // The group says it covers B1:B2, and a third cell points to it from B4.
  const records = cat(
    formula(0, 1, ptg.exp(0, 1), 1),
    shrFmla(0, 1, 1, 1, tokens(ptg.refN(0, 0xff), ptg.int(1), ptg.add)),
    formula(1, 1, ptg.exp(0, 1), 2),
    formula(3, 1, ptg.exp(0, 1), 3),
  );
  const { before, after, pkg } = await convert({ sheets: [{ name: 'S', records }] });
  assert.deepEqual(getCell(before.sheets[0], 3, 1)?.formula, { text: '', type: 'shared', si: 0 });
  assert.deepEqual(getCell(after.sheets[0], 0, 1)?.formula, { text: 'A1+1', type: 'shared', ref: 'B1:B2', si: 0 });
  assert.deepEqual(getCell(after.sheets[0], 1, 1)?.formula, { text: '', type: 'shared', si: 0 });
  assert.deepEqual(getCell(after.sheets[0], 3, 1)?.formula, { text: 'A4+1', type: 'normal' });
  assert.match(text(await pkg.readPart('/xl/worksheets/sheet1.xml')), /<c r="B4"><f>A4\+1<\/f><v>3<\/v><\/c>/);
  assert.equal(await workbookDifference(before, after), undefined);
});

test('xls to xlsx: a formula that refers to another workbook keeps its text', async () => {
  const spec: WorkbookSpec = {
    tail: cat(rec(0x01ae, u16(1), str16('other.xls'), str16('Sheet1')), rec(0x0017, u16(1), u16(0), u16(0), u16(0))),
    sheets: [{ name: 'S', records: formula(0, 0, tokens(ptg.ref3d(0, 1, 1)), 7) }],
  };
  const { before, after } = await convert(spec);
  assert.equal(getCell(before.sheets[0], 0, 0)?.formula?.text, '[1]Sheet1!$B$2');
  assert.deepEqual([getCell(after.sheets[0], 0, 0)?.formula?.text, getCell(after.sheets[0], 0, 0)?.value], ['[1]Sheet1!$B$2', 7]);
});

// --- drawings and comments ------------------------------------------------------------------

const RED: [number, number, number] = [220, 0, 0];

function artBook(): WorkbookSpec {
  const image = png(8, 4, RED);
  const head = cat(defaultStyles(), font({ bold: true, name: 'Tahoma', size: 9 }), rec(0x00eb, drawingGroup([bse(image), bse(new Uint8Array([1, 2, 3, 4]), 0xf01a, 2)])));
  const records = cat(
    colInfo(1, 1, 4096),
    row(2, 600, 0x40),
    number(0, 0, 1),
    drawingRecords([
      shape({ kind: 75, ot: 8, id: 1, props: [[0x4104, 1], [0x0380, 'Logo & <ek>'], [0x0381, 'Şirket "logosu"'], [0x0102, 0x4000], [0x0004, 30 << 16]], anchor: clientAnchor([1, 512, 2, 64], [3, 256, 6, 128], 2), flags: 0xa40 }),
      shape({ kind: 3, ot: 3, id: 2, props: [[0x0181, 0x0000ff], [0x01c0, 0x00ff00], [0x01cb, 19050], [0x01ce, 1]], anchor: clientAnchor([0, 0, 8, 0], [2, 0, 12, 0]), text: txo('ilk çizgi\nikinci <çizgi> & son', [[0, 5], [9, 0]]) }),
      shape({ kind: 20, ot: 1, id: 3, anchor: clientAnchor([4, 0, 0, 0], [6, 0, 3, 0], 1) }),
      groupOf(4, [0, 0, 1000, 500], clientAnchor([0, 0, 14, 0], [4, 0, 18, 0]), [
        shape({ kind: 1, ot: 2, id: 5, anchor: childAnchor(0, 0, 500, 500), props: [[0x01bf, 0x100000], [0x01ff, 0x80000]] }),
        shape({ kind: 75, ot: 8, id: 6, anchor: childAnchor(500, 100, 1000, 400), props: [[0x4104, 1]] }),
      ]),
      // A picture kept as a metafile, which is never unpacked, and a chart.
      shape({ kind: 75, ot: 8, id: 7, props: [[0x4104, 2]], anchor: clientAnchor([8, 0, 0, 0], [9, 0, 2, 0]) }),
      shape({ kind: 201, ot: 5, id: 8, anchor: clientAnchor([8, 0, 5, 0], [9, 0, 9, 0]) }),
      shape({ kind: 202, ot: 0x19, id: 9, anchor: clientAnchor([2, 512, 0, 128], [4, 0, 4, 0]), text: txo('Ayse:\nKontrol <edildi> & tamam', [[0, 5], [5, 0]]) }),
      shape({ kind: 202, ot: 0x19, id: 10, text: txo('ikinci') }),
    ]),
    note(1, 1, 9, 'Ayşe & <Bora>'),
    note(4, 2, 10, '', true),
  );
  return { head, sheets: [{ name: 'Çizim', records }, { name: 'Boş' }, { name: 'Not', records: cat(drawingRecords([shape({ kind: 202, ot: 0x19, id: 1, text: txo('öteki sayfa') })]), note(0, 0, 1, 'Can')) }] };
}

test('xls to xlsx: pictures, shapes, groups and comments are written into drawing parts', async () => {
  const file = await open(fileOf(artBook()));
  const before = await file.workbook();
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const bytes = await buildXlsx(before, { report });
  const reopened = await open(bytes);
  const after = await reopened.workbook();
  const pkg = reopened.package!;

  assert.deepEqual(report.warnings, [
    'Drawing: a picture whose bytes cannot be read (a metafile, or one missing from the file) is written as an empty frame',
    'Drawing: a chart is not in the model and is left out',
    // The chart object of the fixture has no records of a chart after it.
    'Sheet "Çizim" was read with a warning: Chart is not shown: no records of a chart follow its object',
  ]);
  assert.deepEqual(report.parts, pkg.partNames().concat('/[Content_Types].xml').sort((a, b) => report.parts.indexOf(a) - report.parts.indexOf(b)));
  assert.deepEqual(
    report.parts.filter((part) => !/sheet\d|workbook|styles|Content_Types|\/_rels\/\.rels/.test(part)),
    ['/xl/drawings/drawing1.xml', '/xl/drawings/_rels/drawing1.xml.rels', '/xl/comments1.xml', '/xl/drawings/vmlDrawing1.vml', '/xl/comments2.xml', '/xl/drawings/vmlDrawing2.vml', '/xl/media/image1.png'],
  );
  assert.equal(pkg.contentTypeOf('/xl/media/image1.png'), 'image/png');
  assert.equal(pkg.contentTypeOf('/xl/drawings/drawing1.xml'), 'application/vnd.openxmlformats-officedocument.drawing+xml');
  assert.equal(pkg.contentTypeOf('/xl/comments2.xml'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml');
  assert.equal(pkg.contentTypeOf('/xl/drawings/vmlDrawing1.vml'), 'application/vnd.openxmlformats-officedocument.vmlDrawing');
  for (const part of pkg.partNames()) if (/\.(xml|vml|rels)$/.test(part)) await pkg.readXml(part);

  // Same places (in pixels), same pictures (by their bytes), same texts.
  assert.equal(await workbookDifference(before, after), undefined);

  const sheet = after.sheets[0];
  const source = before.sheets[0];
  assert.equal(source.drawings?.length, 6);
  assert.equal(sheet.drawings?.length, 5);
  const metrics = createSheetMetrics(sheet);
  for (let i = 0; i < 5; i++) {
    const was = anchorRect(source.drawings![i].anchor, createSheetMetrics(source));
    const is = anchorRect(sheet.drawings![i].anchor, metrics);
    for (const side of ['x', 'y', 'width', 'height'] as const) assert.ok(Math.abs(was[side] - is[side]) < 0.01, `drawing ${i} ${side}: ${was[side]} / ${is[side]}`);
  }
  assert.deepEqual(sheet.drawings?.map((drawing) => [drawing.object.type, drawing.anchor.kind, drawing.anchor.editAs]), [
    ['picture', 'twoCell', 'oneCell'],
    ['shape', 'twoCell', undefined],
    ['shape', 'twoCell', 'absolute'],
    ['group', 'twoCell', undefined],
    ['picture', 'twoCell', undefined],
  ]);
  const picture = sheet.drawings![0].object as DrawingPicture;
  assert.deepEqual([picture.name, picture.description, picture.rotation, picture.flipH, picture.crop], ['Logo & <ek>', 'Şirket "logosu"', 30, true, { left: 0.25, top: 0, right: 0, bottom: 0 }]);
  assert.equal(picture.media, '/xl/media/image1.png');
  assert.deepEqual(await after.media!.readPart(picture.media!), png(8, 4, RED));
  const ellipse = sheet.drawings![1].object as DrawingShape;
  assert.deepEqual([ellipse.geometry, ellipse.fill, ellipse.line], ['ellipse', '#FF0000', { color: '#00FF00', width: 2, dash: true }]);
  assert.equal(ellipse.text?.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('|'), 'ilk çizgi\nikinci <çizgi> & son');
  assert.equal(ellipse.text?.paragraphs[0].runs[0].bold, true);
  const line = sheet.drawings![2].object as DrawingShape;
  assert.deepEqual([line.geometry, line.connector, line.fill], ['line', true, undefined]);
  const group = sheet.drawings![3].object as DrawingGroup;
  assert.deepEqual(group.children.map((child) => [child.type, child.box]), [
    ['shape', { x: 0, y: 0, width: 0.5, height: 1 }],
    ['picture', { x: 0.5, y: 0.2, width: 0.5, height: 0.6 }],
  ]);
  // One picture, shown twice, is one part.
  assert.equal((group.children[1] as DrawingPicture).media, '/xl/media/image1.png');
  assert.deepEqual([(group.children[0] as DrawingShape).fill, (group.children[0] as DrawingShape).line], [undefined, undefined]);
  // The metafile: a frame without a picture.
  assert.equal((sheet.drawings![4].object as DrawingPicture).media, undefined);

  assert.deepEqual(
    sheet.comments?.map((comment) => [comment.row, comment.col, comment.author, comment.text, comment.visible]),
    [
      [1, 1, 'Ayşe & <Bora>', 'Ayse:\nKontrol <edildi> & tamam', undefined],
      [4, 2, undefined, 'ikinci', true],
    ],
  );
  assert.deepEqual(sheet.comments?.[0].runs?.map((run) => [run.text, run.font?.bold]), [['Ayse:', true], ['\nKontrol <edildi> & tamam', false]]);
  const box = (comment: { box?: Parameters<typeof anchorRect>[0] }, of: typeof sheet) => anchorRect(comment.box!, createSheetMetrics(of));
  const was = box(source.comments![0], source);
  const is = box(sheet.comments![0], sheet);
  for (const side of ['x', 'y', 'width', 'height'] as const) assert.ok(Math.abs(was[side] - is[side]) <= 1, `comment box ${side}: ${was[side]} / ${is[side]}`);
  assert.deepEqual(after.sheets[2].comments?.map((comment) => [comment.row, comment.col, comment.author, comment.text]), [[0, 0, 'Can', 'öteki sayfa']]);
  assert.equal(after.sheets[1].comments, undefined);

  // Shape ids of the two legacy drawings do not meet.
  assert.match(text(await pkg.readPart('/xl/drawings/vmlDrawing1.vml')), /<o:idmap v:ext="edit" data="1"\/>.*_x0000_s1025.*_x0000_s1026/s);
  assert.match(text(await pkg.readPart('/xl/drawings/vmlDrawing2.vml')), /<o:idmap v:ext="edit" data="2"\/>.*_x0000_s2049/s);

  // Deterministic, and unchanged by a save.
  assert.deepEqual(await buildXlsx(before), bytes);
  assert.deepEqual(await reopened.save(), bytes);
  assert.equal((await mediaDigests(after)).size, 1);
});

test('xls to xlsx: an older workbook (BIFF5) is written too', async () => {
  const label = (r: number, c: number, value: string) => rec(0x0204, u16(r), u16(c), u16(15), u16(value.length), [...value].map((ch) => ch.charCodeAt(0)));
  // A font of the older format: its name is a byte string.
  const font5 = rec(0x0031, u16(200), u16(0), u16(0x7fff), u16(400), u16(0), [0, 0, 0, 0], [5, ...[...'Arial'].map((ch) => ch.charCodeAt(0))]);
  const spec: WorkbookSpec = { version: 0x0500, head: cat(font5, ...Array.from({ length: 15 }, () => xf({ style: true })), xf()), sheets: [{ name: 'Eski', records: cat(label(0, 0, 'metin'), number(1, 0, 2.5), rec(0x0006, u16(2), u16(0), u16(15), f64(5), u16(0), u32(0), u16(3), [0x1e, 5, 0])) }] };
  const file = await open(fileOf(spec));
  const before = await file.workbook();
  const after = await (await open(await file.saveAsXlsx!())).workbook();
  assert.equal(await workbookDifference(before, after), undefined);
  // A text that came without a shared string gets one in the new file.
  assert.deepEqual([getCell(before.sheets[0], 0, 0)?.si, getCell(after.sheets[0], 0, 0)?.si, getCell(after.sheets[0], 0, 0)?.value], [undefined, 0, 'metin']);
  assert.deepEqual([getCell(after.sheets[0], 2, 0)?.formula?.text, getCell(after.sheets[0], 2, 0)?.value], ['5', 5]);
});

test('xls to xlsx: the comparison the tests rely on notices what differs', async () => {
  const { before, after } = await convert(richBook());
  const differs = async (change: (book: Workbook) => void) => {
    const copy = await (await open(await buildXlsx(before))).workbook();
    change(copy);
    return (await workbookDifference(before, copy)) !== undefined;
  };
  assert.equal(await workbookDifference(before, after), undefined);
  assert.ok(await differs((book) => (getCell(book.sheets[0], 1, 0)!.value = 1234.6)));
  assert.ok(await differs((book) => (getCell(book.sheets[0], 2, 0)!.kind = 'number')));
  assert.ok(await differs((book) => (getCell(book.sheets[0], 3, 0)!.formula!.text = 'A2+B3')));
  assert.ok(await differs((book) => (getCell(book.sheets[0], 5, 1)!.formula = { text: 'A6*3', type: 'normal' })));
  assert.ok(await differs((book) => (book.styles.fonts[book.styles.cellXfs[getCell(book.sheets[0], 0, 0)!.style].fontId].size = 9.5)));
  assert.ok(await differs((book) => (book.styles.cellXfs[getCell(book.sheets[0], 1, 0)!.style].numFmtId = 2)));
  assert.ok(await differs((book) => book.sheets[0].merges.pop()));
  assert.ok(await differs((book) => (book.sheets[0].columns[0].width = 18)));
  assert.ok(await differs((book) => (book.sheets[0].rows.get(1)!.hidden = false)));
  assert.ok(await differs((book) => (book.sheets[0].view.frozenRows = 3)));
  assert.ok(await differs((book) => (book.sheets[0].hyperlinks[0].target = 'https://example.com/')));
  assert.ok(await differs((book) => (book.definedNames[0].value = 'Veri!$A$5')));
  assert.ok(await differs((book) => (book.sheets[1].state = 'visible')));
  assert.ok(await differs((book) => (book.sheets[2].name = 'x')));
  assert.ok(await differs((book) => (book.activeSheet = 2)));
  assert.ok(await differs((book) => (book.styles.indexedColors![9] = '000000')));
  assert.ok(await differs((book) => book.sheets[0].rows.delete(20)));
});
