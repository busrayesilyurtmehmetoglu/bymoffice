// Page setup (margins, paper, scaling, headers and footers, print options, page
// breaks): read from xls records and from xlsx elements into the model, and
// written by buildXlsx(). Every file is built here, record by record or as XML.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readXlsStream } from '../src/formats/xls/workbook';
import { applyOp, buildXlsx, open, parseXml, localName, type PageBreak, type SheetPageSetup, type Workbook, type XlsxBuildReport, type XmlElement } from '../src/index';
import { OpcPackage } from '../src/opc';
import { buildXlsx as makeXlsx, worksheetXml } from './fixtures';
import { workbookDifference } from './workbook-content';
import { cat, f64, narrow, number, rec, str16, u16, workbookStream, xlsFile } from './xls-fixtures';
import { bookOf, build, put, sheetOf, text } from './xlsx-build-fixtures';

// ---------------------------------------------------------------------------
// Records

const margin = (id: number, inches: number) => rec(id, f64(inches));
const flag = (id: number, on: boolean) => rec(id, u16(on ? 1 : 0));
const wsBool = (flags: number) => rec(0x0081, u16(flags));

interface SetupSpec {
  paper?: number;
  scale?: number;
  start?: number;
  fitWidth?: number;
  fitHeight?: number;
  flags?: number;
  hRes?: number;
  vRes?: number;
  header?: number;
  footer?: number;
  copies?: number;
}

const setup = (spec: SetupSpec = {}) =>
  rec(
    0x00a1,
    u16(spec.paper ?? 9),
    u16(spec.scale ?? 100),
    u16(spec.start ?? 1),
    u16(spec.fitWidth ?? 1),
    u16(spec.fitHeight ?? 1),
    u16(spec.flags ?? 0x0002),
    u16(spec.hRes ?? 600),
    u16(spec.vRes ?? 600),
    f64(spec.header ?? 0.5),
    f64(spec.footer ?? 0.5),
    u16(spec.copies ?? 1),
  );

/** HEADERFOOTER: a 12-byte header of its record family, the id of a custom view (zeros for the sheet), flags, four counts, the texts. */
function headerFooter(flags: number, texts: [evenHeader: string, evenFooter: string, firstHeader: string, firstFooter: string], view = 0): Uint8Array {
  return rec(0x089c, u16(0x089c), u16(0), new Array(8).fill(0), new Array(16).fill(view), u16(flags), ...texts.map((item) => u16(item.length)), ...texts.map((item) => (item === '' ? [] : str16(item))));
}

/** Page breaks of BIFF8: where, and the columns or rows each one spans. */
const breaks8 = (id: number, ...items: [at: number, min: number, max: number][]) => rec(id, u16(items.length), ...items.map(([at, min, max]) => [...u16(at), ...u16(min), ...u16(max)]));

const read8 = (records: Uint8Array) => readXlsStream(workbookStream({ sheets: [{ name: 'S', records }] }));
const read5 = (records: Uint8Array, head: Uint8Array | number[] = []) => readXlsStream(workbookStream({ version: 0x0500, head, sheets: [{ name: 'S', records }] }), 'Book');

const PAGE_RECORDS = [0x0014, 0x0015, 0x001a, 0x001b, 0x0026, 0x0027, 0x0028, 0x0029, 0x002a, 0x002b, 0x0082, 0x0083, 0x0084, 0x00a1, 0x089c];

// ---------------------------------------------------------------------------
// xls -> model

test('xls page setup: every record of BIFF8 is read into the model and none is counted as skipped', () => {
  const records = cat(
    wsBool(0x05c1),
    breaks8(0x001b, [40, 0, 255], [20, 0, 255], [60, 2, 5]),
    breaks8(0x001a, [7, 0, 65535], [3, 10, 20]),
    rec(0x0014, str16('&L&"Arial,Bold"&12Gemi adı: İŞĞÜÖÇ ığüşöç&C&&Co &P/&N&R&D')),
    rec(0x0015, str16('&CSayfa &P')),
    flag(0x0083, true),
    flag(0x0084, true),
    margin(0x0026, 0.25),
    margin(0x0027, 0.3),
    margin(0x0028, 0.4),
    margin(0x0029, 0.45),
    flag(0x002a, true),
    flag(0x002b, true),
    flag(0x0082, false),
    // The data of a printer driver: not read, and counted.
    rec(0x004d, u16(0), new Array(20).fill(7)),
    // Over then down, landscape, black and white, draft, comments at the end, a first page number, errors as dashes.
    setup({ paper: 8, scale: 85, start: 5, fitWidth: 2, fitHeight: 0, flags: 0x0001 | 0x0008 | 0x0010 | 0x0020 | 0x0080 | 0x0200 | 0x0800, hRes: 300, vRes: 0xfffd, header: 0.1, footer: 0.2, copies: 3 }),
    headerFooter(0x0003, ['&CÇift', '&Leven', '&Cİlk', '']),
    number(0, 0, 1),
  );
  const { workbook, report } = read8(records);
  assert.deepEqual(workbook.sheets[0].pageSetup, {
    margins: { left: 0.25, right: 0.3, top: 0.4, bottom: 0.45, header: 0.1, footer: 0.2 },
    page: {
      paperSize: 8,
      scale: 85,
      orientation: 'landscape',
      horizontalDpi: 300,
      // A quality level of the driver (-3), as an xlsx file states it.
      verticalDpi: 4294967293,
      copies: 3,
      fitToWidth: 2,
      fitToHeight: 0,
      firstPageNumber: 5,
      useFirstPageNumber: true,
      pageOrder: 'overThenDown',
      blackAndWhite: true,
      draft: true,
      cellComments: 'atEnd',
      errors: 'dash',
    },
    print: { horizontalCentered: true, verticalCentered: true, headings: true, gridLines: true },
    headerFooter: {
      oddHeader: '&L&"Arial,Bold"&12Gemi adı: İŞĞÜÖÇ ığüşöç&C&&Co &P/&N&R&D',
      oddFooter: '&CSayfa &P',
      evenHeader: '&CÇift',
      evenFooter: '&Leven',
      firstHeader: '&Cİlk',
      differentOddEven: true,
      differentFirst: true,
      scaleWithDoc: false,
      alignWithMargins: false,
    },
    fitToPage: true,
    rowBreaks: [
      // A break across the whole of the old sheet spans the whole sheet; in ascending order.
      { at: 20, min: 0, max: 16383, manual: true },
      { at: 40, min: 0, max: 16383, manual: true },
      { at: 60, min: 2, max: 5, manual: true },
    ],
    colBreaks: [
      { at: 3, min: 10, max: 20, manual: true },
      { at: 7, min: 0, max: 1048575, manual: true },
    ],
  } satisfies SheetPageSetup);
  for (const id of PAGE_RECORDS) assert.equal(report.skippedRecords.get(id), undefined, `record 0x${id.toString(16)}`);
  assert.equal(report.skippedRecords.get(0x004d), 1);
});

test('xls page setup: a sheet as an application writes it by default states only what differs from the defaults', () => {
  const records = cat(
    wsBool(0x04c1),
    rec(0x0014),
    rec(0x0015),
    flag(0x0083, false),
    flag(0x0084, false),
    margin(0x0026, 0.7),
    margin(0x0027, 0.7),
    margin(0x0028, 0.75),
    margin(0x0029, 0.75),
    flag(0x002a, false),
    flag(0x002b, false),
    flag(0x0082, true),
    setup({ header: 0.3, footer: 0.3 }),
    headerFooter(0x000c, ['', '', '', '']),
  );
  const { workbook, report } = read8(records);
  assert.deepEqual(workbook.sheets[0].pageSetup, {
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
    page: { paperSize: 9, scale: 100, orientation: 'portrait', horizontalDpi: 600, verticalDpi: 600, copies: 1 },
  });
  assert.equal(report.skippedRecords.size, 0);
});

test('xls page setup: a sheet without any of the records has no page setup', () => {
  assert.equal(read8(number(0, 0, 1)).workbook.sheets[0].pageSetup, undefined);
  // Automatic page breaks that are not shown are a statement of their own.
  assert.deepEqual(read8(wsBool(0x04c0)).workbook.sheets[0].pageSetup, { autoPageBreaks: false });
});

test('xls page setup: SETUP that says its printer settings were never set gives no paper, scale, orientation, resolution or copies', () => {
  // fNoPls: the fields hold whatever was in memory.
  const never = read8(setup({ paper: 77, scale: 3, flags: 0x0004 | 0x0002, hRes: 1, vRes: 2, copies: 9, fitWidth: 3, header: 0.2, footer: 0.25 })).workbook.sheets[0].pageSetup;
  assert.deepEqual(never, { margins: { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.2, footer: 0.25 }, page: { fitToWidth: 3 } });
  // fNoOrient: only the orientation is not set.
  const noOrientation = read8(setup({ flags: 0x0040 | 0x0002, paper: 1 })).workbook.sheets[0].pageSetup;
  assert.deepEqual(noOrientation?.page, { paperSize: 1, scale: 100, horizontalDpi: 600, verticalDpi: 600, copies: 1 });
  assert.equal(read8(setup({ flags: 0 })).workbook.sheets[0].pageSetup?.page?.orientation, 'landscape');
  // The first page number only counts with its flag; it may be negative.
  assert.equal(read8(setup({ start: 7 })).workbook.sheets[0].pageSetup?.page?.firstPageNumber, undefined);
  assert.deepEqual(
    [read8(setup({ start: 0xfffe, flags: 0x0082 })).workbook.sheets[0].pageSetup?.page?.firstPageNumber, read8(setup({ start: 0xfffe, flags: 0x0082 })).workbook.sheets[0].pageSetup?.page?.useFirstPageNumber],
    [-2, true],
  );
  // Comments as displayed, and the other ways of printing errors.
  assert.equal(read8(setup({ flags: 0x0022 })).workbook.sheets[0].pageSetup?.page?.cellComments, 'asDisplayed');
  // UNSURE (see the reader): "blank" is what every real file says, and is taken as nothing said.
  assert.equal(read8(setup({ flags: 0x0402 })).workbook.sheets[0].pageSetup?.page?.errors, undefined);
  assert.equal(read8(setup({ flags: 0x0802 })).workbook.sheets[0].pageSetup?.page?.errors, 'dash');
  assert.equal(read8(setup({ flags: 0x0c02 })).workbook.sheets[0].pageSetup?.page?.errors, 'NA');
});

test('xls page setup: values that mean nothing are left out, the sheet is still read', () => {
  const records = cat(
    margin(0x0026, NaN),
    margin(0x0027, -1),
    margin(0x0028, Infinity),
    margin(0x0029, 0),
    // A scale outside 10..400, no paper, no copies; margins of the header and footer that are no lengths.
    setup({ paper: 0, scale: 401, copies: 0, header: -0.5, footer: NaN, fitWidth: 40000, fitHeight: 32767 }),
    number(2, 1, 5),
  );
  const { workbook, report } = read8(records);
  const sheet = workbook.sheets[0];
  assert.deepEqual(sheet.pageSetup, {
    margins: { left: 0.75, right: 0.75, top: 1, bottom: 0, header: 0.5, footer: 0.5 },
    page: { orientation: 'portrait', horizontalDpi: 600, verticalDpi: 600, fitToHeight: 32767 },
  });
  assert.equal(sheet.rows.get(2)?.cells[1].value, 5);
  // What could not be used is counted with the records that were not read.
  assert.deepEqual([report.skippedRecords.get(0x0026), report.skippedRecords.get(0x0027), report.skippedRecords.get(0x0028), report.skippedRecords.get(0x0029)], [1, 1, 1, undefined]);
  assert.equal(read8(setup({ scale: 9 })).workbook.sheets[0].pageSetup?.page?.scale, undefined);
  assert.equal(read8(setup({ scale: 10 })).workbook.sheets[0].pageSetup?.page?.scale, 10);
  assert.equal(read8(setup({ scale: 400 })).workbook.sheets[0].pageSetup?.page?.scale, 400);
});

test('xls page setup: records that are cut short are left out one by one', () => {
  const records = cat(
    rec(0x0026, [1, 2, 3]),
    rec(0x0083),
    rec(0x0084, [1]),
    rec(0x00a1, u16(9), u16(100), u16(1)),
    // A header whose text runs past the record, and a footer that is fine.
    rec(0x0014, u16(50), [0], narrow('short')),
    rec(0x0015, str16('&Cok')),
    // Three breaks announced, one and a half present.
    rec(0x001b, u16(3), u16(10), u16(0), u16(255), u16(20), u16(0)),
    rec(0x001a, [1]),
    rec(0x089c, u16(0x089c), new Array(20).fill(0)),
    number(0, 0, 1),
  );
  const { workbook, report } = read8(records);
  assert.deepEqual(workbook.sheets[0].pageSetup, { headerFooter: { oddFooter: '&Cok' }, rowBreaks: [{ at: 10, min: 0, max: 16383, manual: true }] });
  assert.equal(workbook.sheets[0].rows.get(0)?.cells[0].value, 1);
  for (const id of [0x0026, 0x0083, 0x0084, 0x00a1, 0x0014, 0x001a, 0x089c]) assert.equal(report.skippedRecords.get(id), 1, `record 0x${id.toString(16)}`);
  // SETUP that ends after its flags, or in the middle of its later fields, gives what it has.
  const partial = read8(rec(0x00a1, u16(9), u16(50), u16(1), u16(1), u16(1), u16(0x0002), u16(300))).workbook.sheets[0].pageSetup;
  assert.deepEqual(partial, { page: { paperSize: 9, scale: 50, orientation: 'portrait' } });
  // HEADERFOOTER whose texts are cut off is left out as a whole, flags included.
  const cut = read8(rec(0x089c, u16(0x089c), u16(0), new Array(24).fill(0), u16(0x0001), u16(9), u16(0), u16(0), u16(0), u16(9), [0], narrow('abc')));
  assert.equal(cut.workbook.sheets[0].pageSetup, undefined);
  assert.equal(cut.report.skippedRecords.get(0x089c), 1);
});

test('xls page setup: of a record stated twice the first one counts; the settings of a custom view are not the sheet\'s', () => {
  const records = cat(
    margin(0x0026, 0.1),
    margin(0x0026, 0.9),
    rec(0x0014, str16('first')),
    rec(0x0014, str16('second')),
    setup({ scale: 50 }),
    setup({ scale: 60 }),
    // One that could not be read does not stand in the way of the next.
    rec(0x0027, [0]),
    margin(0x0027, 0.2),
    breaks8(0x001b, [5, 0, 255], [5, 1, 2], [0, 0, 255], [9, 300, 400], [8, 5, 2]),
    // A custom view: its own header, margins, breaks and paper.
    rec(0x01aa, new Array(20).fill(1)),
    rec(0x0015, str16('view footer')),
    margin(0x0028, 3),
    breaks8(0x001a, [4, 0, 65535]),
    flag(0x0083, true),
    rec(0x01ab, u16(1)),
    margin(0x0029, 0.4),
    // HEADERFOOTER of a custom view names the view; the one of the sheet does not.
    headerFooter(0x000f, ['view', '', '', ''], 7),
    headerFooter(0x0005, ['sheet', '', '', '']),
  );
  const { workbook, report } = read8(records);
  assert.deepEqual(workbook.sheets[0].pageSetup, {
    margins: { left: 0.1, right: 0.2, top: 1, bottom: 0.4, header: 0.5, footer: 0.5 },
    page: { paperSize: 9, scale: 50, orientation: 'portrait', horizontalDpi: 600, verticalDpi: 600, copies: 1 },
    headerFooter: { oddHeader: 'first', evenHeader: 'sheet', differentOddEven: true, alignWithMargins: false },
    // One break per row; none above the first row or outside the sheet.
    rowBreaks: [{ at: 5, min: 0, max: 16383, manual: true }],
  });
  const skipped = (id: number) => report.skippedRecords.get(id);
  assert.deepEqual([skipped(0x0026), skipped(0x0014), skipped(0x00a1), skipped(0x0027), skipped(0x0015), skipped(0x0028), skipped(0x001a), skipped(0x0083), skipped(0x089c)], [1, 1, 1, 1, 1, 1, 1, 1, 1]);
  assert.deepEqual([skipped(0x01aa), skipped(0x01ab)], [1, 1]);
});

test('xls page setup: a header longer than 255 characters and one that goes on in a CONTINUE record are read whole', () => {
  const long = `&L${'Uzun başlık '.repeat(40)}&R&P`;
  assert.equal(read8(rec(0x0014, str16(long))).workbook.sheets[0].pageSetup?.headerFooter?.oddHeader, long);
  // The text goes on after a byte that states the width again.
  const split = cat(rec(0x0015, u16(8), [0], narrow('abcd')), rec(0x003c, [1], [0x5e, 0x01, 0x65, 0, 0x66, 0, 0x67, 0]));
  assert.equal(read8(split).workbook.sheets[0].pageSetup?.headerFooter?.oddFooter, 'abcdŞefg');
});

test('xls page setup: BIFF5 has byte strings in the code page, breaks without a span and fewer flags', () => {
  const records = cat(
    wsBool(0x0101),
    // Places only: each break spans the whole sheet.
    rec(0x001b, u16(2), u16(30), u16(12)),
    rec(0x001a, u16(1), u16(4)),
    // "&CŞirket İşi" in Windows-1254.
    rec(0x0014, [12], narrow('&C'), [0xde], narrow('irket '), [0xdd, 0xfe], narrow('i')),
    rec(0x0015, [0]),
    flag(0x0083, true),
    margin(0x0026, 0.5),
    flag(0x002b, true),
    // Bits that only BIFF8 gives a meaning (comments at the end, errors) are not read.
    setup({ paper: 9, scale: 70, flags: 0x0020 | 0x0200 | 0x0c00, copies: 2 }),
    // Not a record of this format.
    headerFooter(0x0003, ['x', '', '', '']),
    number(0, 0, 1),
  );
  const { workbook, report } = read5(records, rec(0x0042, u16(1254)));
  assert.equal(report.biff, 5);
  assert.deepEqual(workbook.sheets[0].pageSetup, {
    margins: { left: 0.5, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
    page: { paperSize: 9, scale: 70, orientation: 'landscape', horizontalDpi: 600, verticalDpi: 600, copies: 2, cellComments: 'asDisplayed' },
    print: { horizontalCentered: true, gridLines: true },
    headerFooter: { oddHeader: '&CŞirket İşi' },
    fitToPage: true,
    rowBreaks: [
      { at: 12, min: 0, max: 16383, manual: true },
      { at: 30, min: 0, max: 16383, manual: true },
    ],
    colBreaks: [{ at: 4, min: 0, max: 1048575, manual: true }],
  });
  assert.equal(report.skippedRecords.get(0x089c), 1);
  for (const id of [0x0014, 0x0015, 0x001a, 0x001b, 0x0026, 0x002b, 0x0083, 0x00a1]) assert.equal(report.skippedRecords.get(id), undefined);
  // A byte string that runs past its record leaves the header out.
  const cut = read5(rec(0x0014, [40], narrow('abc')));
  assert.equal(cut.workbook.sheets[0].pageSetup, undefined);
  assert.equal(cut.report.skippedRecords.get(0x0014), 1);
});

test('xls page setup: a hostile break count reads no further than the record', () => {
  const many = read8(rec(0x001b, u16(0xffff), ...Array.from({ length: 2000 }, (_, i) => [...u16((i % 500) + 1), ...u16(0), ...u16(255)]))).workbook.sheets[0].pageSetup;
  assert.equal(many?.rowBreaks?.length, 500);
  assert.deepEqual(many?.rowBreaks?.slice(0, 2).map((item) => item.at), [1, 2]);
});

// ---------------------------------------------------------------------------
// xlsx -> model

const PRINT_XML =
  '<printOptions horizontalCentered="1" verticalCentered="0" headings="true" gridLines="1" gridLinesSet="0"/>' +
  '<pageMargins left="0.25" right="0.3" top="0.4" bottom="0.45" header="0.1" footer="1.5e-1"/>' +
  '<pageSetup paperSize="9" scale="85" firstPageNumber="5" fitToWidth="2" fitToHeight="0" pageOrder="overThenDown" orientation="landscape" usePrinterDefaults="0" blackAndWhite="1" draft="1" cellComments="atEnd" useFirstPageNumber="1" errors="NA" horizontalDpi="4294967293" verticalDpi="300" copies="2" r:id="rId7"/>' +
  '<headerFooter differentOddEven="1" differentFirst="1" scaleWithDoc="0" alignWithMargins="0">' +
  '<oddHeader>&amp;L&amp;"Arial,Bold"İŞĞ &lt;a&gt;&amp;C&amp;&amp;Co_x000D_\nline 2</oddHeader><oddFooter>&amp;CSayfa &amp;P / &amp;N</oddFooter>' +
  '<evenHeader>even</evenHeader><evenFooter> padded </evenFooter><firstHeader>first</firstHeader><firstFooter>last</firstFooter></headerFooter>' +
  '<rowBreaks count="2" manualBreakCount="1"><brk id="40" max="16383" man="1"/><brk id="20" min="2" max="5" pt="1"/></rowBreaks>' +
  '<colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>';

const EXPECTED: SheetPageSetup = {
  fitToPage: true,
  autoPageBreaks: false,
  print: { horizontalCentered: true, verticalCentered: false, headings: true, gridLines: true, gridLinesSet: false },
  margins: { left: 0.25, right: 0.3, top: 0.4, bottom: 0.45, header: 0.1, footer: 0.15 },
  page: {
    paperSize: 9,
    scale: 85,
    firstPageNumber: 5,
    fitToWidth: 2,
    fitToHeight: 0,
    pageOrder: 'overThenDown',
    orientation: 'landscape',
    usePrinterDefaults: false,
    blackAndWhite: true,
    draft: true,
    cellComments: 'atEnd',
    useFirstPageNumber: true,
    errors: 'NA',
    horizontalDpi: 4294967293,
    verticalDpi: 300,
    copies: 2,
  },
  headerFooter: {
    differentOddEven: true,
    differentFirst: true,
    scaleWithDoc: false,
    alignWithMargins: false,
    oddHeader: '&L&"Arial,Bold"İŞĞ <a>&C&&Co\r\nline 2',
    oddFooter: '&CSayfa &P / &N',
    evenHeader: 'even',
    evenFooter: ' padded ',
    firstHeader: 'first',
    firstFooter: 'last',
  },
  rowBreaks: [
    { at: 20, min: 2, max: 5, manual: false, pivot: true },
    { at: 40, min: 0, max: 16383, manual: true },
  ],
  colBreaks: [{ at: 3, min: 0, max: 1048575, manual: true }],
};

const sheetWith = (before: string, after: string) => worksheetXml(`${before}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${after}`);

async function readSheet(xml: string): Promise<Workbook> {
  const file = await open(await makeXlsx({ sheets: [{ name: 'S', xml }] }));
  return file.workbook();
}

test('xlsx page setup: the elements of a worksheet are read into the model', async () => {
  const workbook = await readSheet(sheetWith('<sheetPr><tabColor rgb="FFFF0000"/><pageSetUpPr autoPageBreaks="0" fitToPage="1"/></sheetPr>', PRINT_XML));
  // Listed out of order in the file: each break knows its place there (see PageBreak.fileOrder).
  assert.deepEqual(workbook.sheets[0].pageSetup, { ...EXPECTED, rowBreaks: [{ ...EXPECTED.rowBreaks![0], fileOrder: 1 }, { ...EXPECTED.rowBreaks![1], fileOrder: 0 }] });
  assert.deepEqual(workbook.sheets[0].tabColor, { rgb: 'FFFF0000' });
  // A sheet that says nothing has none.
  assert.equal((await readSheet(sheetWith('', ''))).sheets[0].pageSetup, undefined);
  assert.equal((await readSheet(sheetWith('<sheetPr/>', '<rowBreaks count="0"/>'))).sheets[0].pageSetup, undefined);
});

test('xlsx page setup: attributes that hold no value of their type state nothing; an element stated twice counts once', async () => {
  const xml = sheetWith(
    '<sheetPr><pageSetUpPr fitToPage="yes"/></sheetPr>',
    '<printOptions gridLines="2" headings="1"/><printOptions headings="0"/>' +
      '<pageMargins left="-1" right="NaN" top="0x10" header="INF" footer="0"/><pageMargins left="9" right="9" top="9" bottom="9" header="9" footer="9"/>' +
      '<pageSetup paperSize="-1" scale="401" fitToWidth="1.5" orientation="sideways" pageOrder="" errors="na" cellComments="all" draft="maybe" copies="4294967296" horizontalDpi="4294967295" paperWidth="210mm" paperHeight="tall"/>' +
      '<pageSetup scale="50"/>' +
      '<headerFooter><oddHeader></oddHeader><oddFooter/><evenHeader>a</evenHeader><evenHeader>b</evenHeader><unknown>x</unknown></headerFooter><headerFooter><oddHeader>second</oddHeader></headerFooter>' +
      '<rowBreaks><brk id="0" man="1"/><brk id="1048576"/><brk id="5" min="9" max="3"/><brk id="6" max="16384"/><brk id="7" max="16383"/><brk id="7" min="1" max="2"/><brk/><other id="9"/></rowBreaks>' +
      '<rowBreaks><brk id="99"/></rowBreaks>' +
      '<colBreaks><brk id="16384"/><brk id="16383" max="1048575" man="true"/></colBreaks>' +
      // The page setup of a custom view is not the sheet's.
      '<customSheetViews><customSheetView guid="{00000000-0000-0000-0000-000000000001}"><pageMargins left="5" right="5" top="5" bottom="5" header="5" footer="5"/><pageSetup scale="33"/></customSheetView></customSheetViews>',
  );
  const setupOf = (await readSheet(xml)).sheets[0].pageSetup;
  assert.deepEqual(setupOf, {
    print: { headings: true },
    // A margin that is none is the one applications give a new sheet.
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0 },
    page: { horizontalDpi: 4294967295, paperWidth: '210mm' },
    headerFooter: { evenHeader: 'a' },
    rowBreaks: [{ at: 7, min: 0, max: 16383, manual: false }],
    colBreaks: [{ at: 16383, min: 0, max: 1048575, manual: true }],
  });
  for (const scale of ['9', '10', '400', '1e2', ' 100']) {
    const page = (await readSheet(sheetWith('', `<pageSetup scale="${scale}"/>`))).sheets[0].pageSetup?.page;
    assert.deepEqual(page, scale === '10' || scale === '400' ? { scale: Number(scale) } : {}, `scale ${scale}`);
  }
});

test('xlsx page setup: saving an edited workbook keeps the elements byte for byte', async () => {
  // With what the model has no field for: the relationship of the printer settings, attributes of other namespaces, odd spacing.
  const elements =
    '<printOptions  horizontalCentered="1" gridLines="1"/>\n  <pageMargins left="0.25" right="0.3" top="0.4" bottom="0.45" header="0.1" footer="0.15"/>' +
    '<pageSetup paperSize="9" scale="85" orientation="landscape" horizontalDpi="4294967293" verticalDpi="300" r:id="rId7" x:future="1" xmlns:x="urn:test"/>' +
    "<headerFooter differentFirst='1'><oddHeader>&amp;L&amp;\"Arial,Bold\"İŞĞ &lt;a&gt;_x000D_</oddHeader><oddFooter>&amp;CSayfa &amp;P</oddFooter><firstHeader>&#x130;lk</firstHeader></headerFooter>" +
    '<rowBreaks count="1" manualBreakCount="1"><brk id="40" max="16383" man="1"/></rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="3" max="1048575" man="1"/></colBreaks>';
  const properties = '<sheetPr codeName="Sayfa1"><tabColor rgb="FF00FF00"/><pageSetUpPr fitToPage="1"/></sheetPr>';
  const original = await makeXlsx({ sheets: [{ name: 'S', xml: sheetWith(properties, elements) }] });
  const file = await open(original);
  const workbook = await file.workbook();
  assert.equal(workbook.sheets[0].pageSetup?.page?.scale, 85);
  assert.equal(workbook.sheets[0].pageSetup?.headerFooter?.firstHeader, 'İlk');

  applyOp(workbook, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 42, style: 0 } }, { row: 5, col: 2, cell: { kind: 'string', value: 'yeni', style: 0 } }] });
  const saved = await OpcPackage.open(await file.save());
  const part = text(await saved.readPart('/xl/worksheets/sheet1.xml'));
  assert.ok(part.includes('<v>42</v>'), 'the edit is in the saved part');
  assert.ok(part.includes(elements), 'the page setup elements are the bytes of the source');
  assert.ok(part.includes(properties), 'so is <sheetPr>');
  // And the file reads back with the same page setup.
  const again = await (await open(await file.save())).workbook();
  assert.deepEqual(again.sheets[0].pageSetup, workbook.sheets[0].pageSetup);
});

// ---------------------------------------------------------------------------
// model -> buildXlsx

/** Children of CT_Worksheet in the order the schema gives them. */
const WORKSHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

const childNames = (el: XmlElement) => el.children.filter((node): node is XmlElement => node.type === 'element').map((node) => localName(node.name));

async function sheetRoot(pkg: OpcPackage): Promise<XmlElement> {
  return parseXml(text(await pkg.readPart('/xl/worksheets/sheet1.xml'))).root;
}

/** A sheet that has something in every element around the page setup. */
function fullSheet(setupOf: SheetPageSetup) {
  const sheet = sheetOf('Form');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  put(sheet, 3, 2, { kind: 'string', value: 'x' });
  sheet.tabColor = { rgb: 'FF336699' };
  sheet.columns = [{ min: 0, max: 0, width: 20, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false }];
  sheet.merges = [{ top: 5, left: 0, bottom: 5, right: 2 }];
  sheet.hyperlinks = [{ range: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'https://example.com/' }];
  sheet.autoFilter = { range: { top: 0, left: 0, bottom: 3, right: 2 }, columns: [] };
  sheet.dataValidations = [{ ranges: [{ top: 1, left: 1, bottom: 1, right: 1 }], type: 'whole', operator: 'between', formula1: '1', formula2: '5', allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: true, errorStyle: 'stop' }];
  sheet.comments = [{ row: 2, col: 2, author: 'A', text: 'note' }];
  sheet.drawings = [{ anchor: { kind: 'absolute', x: 0, y: 0, width: 95250, height: 95250 }, object: { type: 'shape', geometry: 'rect' } }];
  sheet.pageSetup = setupOf;
  return sheet;
}

test('buildXlsx page setup: what the model states reads back the same, in the order of the schema', async () => {
  const model = bookOf(fullSheet(EXPECTED));
  const { after, pkg, warnings } = await build(model);
  assert.deepEqual(after.sheets[0].pageSetup, EXPECTED);
  assert.deepEqual(warnings, []);
  assert.equal(await workbookDifference(model, after), undefined);

  const root = await sheetRoot(pkg);
  const names = childNames(root);
  for (const name of ['sheetPr', 'autoFilter', 'mergeCells', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'drawing', 'legacyDrawing']) assert.ok(names.includes(name), `<${name}> is written`);
  const ranks = names.map((name) => WORKSHEET_ORDER.indexOf(name));
  assert.ok(ranks.every((rank) => rank >= 0), `only elements of the schema: ${names.join(' ')}`);
  assert.deepEqual(ranks, [...ranks].sort((a, b) => a - b), `in schema order: ${names.join(' ')}`);
  // CT_SheetPr: tabColor, outlinePr, pageSetUpPr; CT_HeaderFooter: odd, even, first.
  const sheetPr = root.children.find((node): node is XmlElement => node.type === 'element' && node.name === 'sheetPr')!;
  assert.deepEqual(childNames(sheetPr), ['tabColor', 'pageSetUpPr']);
  const header = root.children.find((node): node is XmlElement => node.type === 'element' && node.name === 'headerFooter')!;
  assert.deepEqual(childNames(header), ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']);
  const rowBreaks = root.children.find((node): node is XmlElement => node.type === 'element' && node.name === 'rowBreaks')!;
  assert.deepEqual([rowBreaks.attrs.get('count'), rowBreaks.attrs.get('manualBreakCount')], ['2', '1']);
});

test('buildXlsx page setup: a sheet without one gets no element, and each member stands for its own element', async () => {
  const plain = sheetOf('Plain');
  put(plain, 0, 0, { kind: 'number', value: 1 });
  const none = await build(bookOf(plain));
  assert.deepEqual(childNames(await sheetRoot(none.pkg)), ['dimension', 'sheetViews', 'sheetData']);

  const cases: [SheetPageSetup, string[]][] = [
    [{}, []],
    [{ fitToPage: true }, ['sheetPr']],
    [{ margins: { left: 1, right: 1, top: 1, bottom: 1, header: 0, footer: 0 } }, ['pageMargins']],
    [{ page: {} }, ['pageSetup']],
    [{ page: { orientation: 'landscape' } }, ['pageSetup']],
    [{ print: { gridLines: true } }, ['printOptions']],
    [{ headerFooter: {} }, ['headerFooter']],
    [{ headerFooter: { oddFooter: '&P' } }, ['headerFooter']],
    [{ rowBreaks: [{ at: 3, min: 0, max: 16383, manual: true }] }, ['rowBreaks']],
    [{ colBreaks: [{ at: 3, min: 0, max: 1048575, manual: true }] }, ['colBreaks']],
    [{ rowBreaks: [], colBreaks: [] }, []],
  ];
  for (const [setupOf, elements] of cases) {
    const sheet = sheetOf('S');
    put(sheet, 0, 0, { kind: 'number', value: 1 });
    sheet.pageSetup = setupOf;
    const model = bookOf(sheet);
    const built = await build(model);
    const names = childNames(await sheetRoot(built.pkg)).filter((name) => !['dimension', 'sheetViews', 'sheetData'].includes(name));
    assert.deepEqual(names, elements, JSON.stringify(setupOf));
    assert.deepEqual(built.warnings, []);
    assert.equal(await workbookDifference(model, built.after), undefined, JSON.stringify(setupOf));
    // An empty record and an empty list of breaks are nothing in a file.
    const expected = elements.length === 0 ? undefined : setupOf;
    assert.deepEqual(built.after.sheets[0].pageSetup, expected, JSON.stringify(setupOf));
  }
});

test('buildXlsx page setup: headers and footers keep every character (escapes, Turkish letters, control characters, long texts)', async () => {
  const long = `&L${'Çok uzun üstbilgi – '.repeat(30)}&R&P/&N`;
  assert.ok(long.length > 255);
  const texts = {
    oddHeader: '&L&"Arial,Bold"&12İŞĞÜÖÇ ığüşöç&C&&Co <b> "q" \'s\' ]]> &amp;&R&D &T',
    oddFooter: 'line 1\nline 2\r\nline 3\ttab',
    evenHeader: long,
    evenFooter: ' leading and trailing ',
    // Looks like an escape and has to stay text; then characters XML cannot carry.
    firstHeader: '_x0041_ and _x000D_',
    firstFooter: 'bell\u0007 null\u0000 lone\ud800 pair😀 ￾',
  };
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  sheet.pageSetup = { headerFooter: { ...texts, differentFirst: true, differentOddEven: true } };
  const model = bookOf(sheet);
  const { after, pkg, warnings } = await build(model);
  assert.deepEqual(after.sheets[0].pageSetup?.headerFooter, { ...texts, differentFirst: true, differentOddEven: true });
  assert.deepEqual(warnings, []);
  assert.equal(await workbookDifference(model, after), undefined);
  const part = text(await pkg.readPart('/xl/worksheets/sheet1.xml'));
  assert.ok(part.includes('<oddHeader>&amp;L&amp;"Arial,Bold"&amp;12İŞĞÜÖÇ ığüşöç&amp;C&amp;&amp;Co &lt;b&gt; "q" \'s\' ]]&gt; &amp;amp;&amp;R&amp;D &amp;T</oddHeader>'), part);
  assert.ok(part.includes('<firstHeader>_x005F_x0041_ and _x005F_x000D_</firstHeader>'), part);
  // The part is well-formed whatever the text holds.
  assert.doesNotThrow(() => parseXml(part));
});

test('buildXlsx page setup: what a file cannot state is left out and counted in the warnings', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  const hostile = {
    fitToPage: 'yes',
    autoPageBreaks: false,
    margins: { left: NaN, right: -0.5, top: Infinity, bottom: 0, header: '1', footer: 0.25 },
    page: {
      paperSize: 9.5,
      scale: 401,
      fitToWidth: -1,
      fitToHeight: 3,
      firstPageNumber: -2,
      useFirstPageNumber: true,
      orientation: 'sideways" injected="1',
      pageOrder: 'overThenDown',
      cellComments: 'everywhere',
      errors: 'NA',
      horizontalDpi: 4294967296,
      verticalDpi: 4294967295,
      copies: NaN,
      draft: 1,
      paperWidth: '21 cm',
      paperHeight: '297mm',
    },
    print: { gridLines: 'true', headings: true },
    headerFooter: { oddHeader: 42, oddFooter: '', evenHeader: 'ok', scaleWithDoc: 0 },
    rowBreaks: [
      { at: 0, min: 0, max: 16383, manual: true },
      { at: 5, min: 0, max: 16383, manual: true },
      { at: 5, min: 0, max: 3, manual: false },
      { at: 1048576, min: 0, max: 16383, manual: true },
      { at: 2.5, min: 0, max: 16383, manual: true },
      { at: 9, min: 4, max: 2, manual: true },
      { at: 3, min: 0, max: 16384, manual: true },
      null,
      { at: 2, min: 1, max: 1, manual: false },
    ],
    colBreaks: [{ at: 16384, min: 0, max: 5, manual: true }],
  };
  sheet.pageSetup = hostile as unknown as SheetPageSetup;
  const model = bookOf(sheet);
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const after = await (await open(await buildXlsx(model, { report }))).workbook();
  assert.deepEqual(after.sheets[0].pageSetup, {
    autoPageBreaks: false,
    margins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0, header: 0.3, footer: 0.25 },
    // The first page number cannot be stated, so "use it" is not stated either.
    page: { fitToHeight: 3, pageOrder: 'overThenDown', errors: 'NA', verticalDpi: 4294967295, paperHeight: '297mm' },
    print: { headings: true },
    headerFooter: { evenHeader: 'ok' },
    rowBreaks: [
      { at: 2, min: 1, max: 1, manual: false },
      { at: 5, min: 0, max: 16383, manual: true },
    ],
  });
  assert.deepEqual(report.warnings, [
    // fitToPage, then ten of <pageSetup>, one of <printOptions>, two of <headerFooter>.
    'Page setup: a value a file cannot state (a number out of range, a name that is none) is left out (14 times)',
    'Page setup: a margin that is no length is written as the default one (4 times)',
    'Page setup: a page break outside the sheet, or a second one at the same place, is left out (8 times)',
  ]);
  // The comparison of a written model with the one read back expects exactly this.
  assert.equal(await workbookDifference(model, after), undefined);
});

test('buildXlsx page setup: more breaks than applications take are written, with a warning', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  const rowBreaks: PageBreak[] = Array.from({ length: 1027 }, (_, i) => ({ at: 2000 - i, min: 0, max: 16383, manual: true }));
  sheet.pageSetup = { rowBreaks };
  const { after, warnings } = await build(bookOf(sheet));
  assert.equal(after.sheets[0].pageSetup?.rowBreaks?.length, 1027);
  assert.deepEqual(after.sheets[0].pageSetup?.rowBreaks?.slice(0, 2).map((item) => item.at), [974, 975]);
  assert.deepEqual(warnings, ['Page setup: a sheet has more than 1026 page breaks in one direction, which applications may refuse']);
});

test('buildXlsx page setup: page breaks follow the rows and columns moved since the sheet was read', async () => {
  const file = await open(await makeXlsx({ sheets: [{ name: 'S', xml: sheetWith('', '<rowBreaks count="2" manualBreakCount="2"><brk id="2" max="16383" man="1"/><brk id="10" min="1" max="4" man="1"/></rowBreaks><colBreaks count="1"><brk id="3" max="1048575" man="1"/></colBreaks>') }] }));
  const model = await file.workbook();
  applyOp(model, { type: 'insert', sheet: 0, axis: 'row', at: 5, count: 3 });
  applyOp(model, { type: 'insert', sheet: 0, axis: 'col', at: 0, count: 2 });
  // The model keeps the breaks as the file has them.
  assert.deepEqual(model.sheets[0].pageSetup?.rowBreaks?.map((item) => item.at), [2, 10]);
  const { after } = await build(model);
  assert.deepEqual(after.sheets[0].pageSetup, {
    rowBreaks: [
      { at: 2, min: 0, max: 16383, manual: true },
      { at: 13, min: 3, max: 6, manual: true },
    ],
    colBreaks: [{ at: 5, min: 0, max: 1048575, manual: true }],
  });
  // A save of the file itself moves them the same way.
  const saved = await (await open(await file.save())).workbook();
  assert.deepEqual(saved.sheets[0].pageSetup, after.sheets[0].pageSetup);
});

// ---------------------------------------------------------------------------
// xls -> xlsx

test('xls -> xlsx: the page setup of an xls file is in the converted file', async () => {
  const records = cat(
    wsBool(0x05c1),
    breaks8(0x001b, [40, 0, 255], [60, 2, 5]),
    breaks8(0x001a, [7, 0, 65535]),
    rec(0x0014, str16('&L&"Arial,Bold"Gemi: İŞĞ &C&&<Co>&R&D')),
    rec(0x0015, str16('&CSayfa &P / &N')),
    flag(0x0083, true),
    flag(0x0084, false),
    margin(0x0026, 0.25),
    margin(0x0027, 0.25),
    margin(0x0028, 0.5),
    margin(0x0029, 0.5),
    rec(0x004d, u16(0), new Array(20).fill(7)),
    setup({ paper: 9, scale: 56, flags: 0x0400, hRes: 0xfffd, vRes: 180, header: 0.3, footer: 0.3 }),
    headerFooter(0x000e, ['', '', '&Cİlk sayfa', '']),
    number(0, 0, 1),
  );
  const stream = workbookStream({ sheets: [{ name: 'Form', records }, { name: 'Boş' }] });
  const file = await open(xlsFile(stream));
  assert.equal(file.format, 'xls');
  const model = await file.workbook();
  const expected: SheetPageSetup = {
    margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    page: { paperSize: 9, scale: 56, orientation: 'landscape', horizontalDpi: 4294967293, verticalDpi: 180, copies: 1 },
    print: { horizontalCentered: true },
    headerFooter: { oddHeader: '&L&"Arial,Bold"Gemi: İŞĞ &C&&<Co>&R&D', oddFooter: '&CSayfa &P / &N', firstHeader: '&Cİlk sayfa', differentFirst: true },
    fitToPage: true,
    rowBreaks: [
      { at: 40, min: 0, max: 16383, manual: true },
      { at: 60, min: 2, max: 5, manual: true },
    ],
    colBreaks: [{ at: 7, min: 0, max: 1048575, manual: true }],
  };
  assert.deepEqual(model.sheets[0].pageSetup, expected);
  assert.equal(model.sheets[1].pageSetup, undefined);

  const converted = await open(await file.saveAsXlsx!());
  assert.equal(converted.format, 'xlsx');
  // saveAsXlsx() is buildXlsx() of the model; built once more for what the builder says about it.
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  await buildXlsx(model, { report });
  assert.ok(report.parts.includes('/xl/worksheets/sheet1.xml'));
  const after = await converted.workbook();
  assert.deepEqual(after.sheets[0].pageSetup, expected);
  assert.equal(after.sheets[1].pageSetup, undefined);
  assert.deepEqual(report.warnings, []);
  assert.equal(await workbookDifference(model, after), undefined);
  // No part for the data of the printer driver, and no relationship to one.
  assert.ok(!report.parts.some((part) => /printerSettings/i.test(part)));
  assert.ok(!text(await converted.package!.readPart('/xl/worksheets/sheet1.xml')).includes('r:id'));
});

test('workbookDifference: a page setup that was lost or changed is a difference', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  sheet.pageSetup = { page: { orientation: 'landscape', scale: 80 }, headerFooter: { oddFooter: '&P' }, rowBreaks: [{ at: 4, min: 0, max: 16383, manual: true }] };
  const model = bookOf(sheet);
  const { after } = await build(model);
  assert.equal(await workbookDifference(model, after), undefined);
  for (const change of [
    (target: SheetPageSetup) => delete target.page!.orientation,
    (target: SheetPageSetup) => (target.page!.scale = 81),
    (target: SheetPageSetup) => (target.headerFooter!.oddFooter = '&N'),
    (target: SheetPageSetup) => (target.rowBreaks![0].at = 5),
    (target: SheetPageSetup) => (target.margins = { left: 1, right: 1, top: 1, bottom: 1, header: 1, footer: 1 }),
  ]) {
    const copy = structuredClone(after.sheets[0].pageSetup!);
    change(copy);
    const difference = await workbookDifference(model, { ...after, sheets: [{ ...after.sheets[0], pageSetup: copy }] });
    assert.match(difference || '', /^sheet 1: \/pageSetup/);
  }
  const lost = await workbookDifference(model, { ...after, sheets: [{ ...after.sheets[0], pageSetup: undefined }] });
  assert.match(lost || '', /^sheet 1: \/pageSetup/);
});

test('review: the first WSBOOL counts, one inside a custom view does not; GRIDSET does not stop the grid lines from printing', () => {
  const wsbool = (flags: number) => {
    const bytes = new Uint8Array(6);
    new DataView(bytes.buffer).setUint16(0, 0x0081, true);
    new DataView(bytes.buffer).setUint16(2, 2, true);
    new DataView(bytes.buffer).setUint16(4, flags, true);
    return bytes;
  };
  const flag = (id: number, value: number) => {
    const bytes = new Uint8Array(6);
    new DataView(bytes.buffer).setUint16(0, id, true);
    new DataView(bytes.buffer).setUint16(2, 2, true);
    new DataView(bytes.buffer).setUint16(4, value, true);
    return bytes;
  };
  // Fit to page in the first record, not in the second.
  assert.equal(read8(cat(wsbool(0x05c1), wsbool(0x04c1))).workbook.sheets[0].pageSetup?.fitToPage, true);
  // Grid lines are printed, and GRIDSET says 0: the model keeps the lines and states nothing that would stop them.
  assert.deepEqual(read8(cat(flag(0x002b, 1), flag(0x0082, 0))).workbook.sheets[0].pageSetup?.print, { gridLines: true });
});

test('review: a header that begins or ends with a space keeps it in the file', async () => {
  const sheet = sheetOf('S');
  sheet.pageSetup = { headerFooter: { oddHeader: '  &CSpaced  ', oddFooter: '&CPlain' } };
  const bytes = await buildXlsx(bookOf(sheet));
  const xml = new TextDecoder().decode(await (await open(bytes)).package!.readPart('/xl/worksheets/sheet1.xml'));
  assert.ok(xml.indexOf('<oddHeader xml:space="preserve">  &amp;CSpaced  </oddHeader>') > 0, xml.slice(xml.indexOf('<headerFooter')));
  assert.ok(xml.indexOf('<oddFooter>&amp;CPlain</oddFooter>') > 0);
  assert.equal((await (await open(bytes)).workbook()).sheets[0].pageSetup?.headerFooter?.oddHeader, '  &CSpaced  ');
});
