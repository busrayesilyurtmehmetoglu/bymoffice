import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readXlsStream } from '../src/formats/xls/workbook';
import { formatCodeOf, resolveCellStyle, resolveColor } from '../src/model/sheet/color';
import { getCell } from '../src/model/sheet/types';
import { createSheetMetrics } from '../src/layout/sheet-metrics';
import { ZipArchive } from '../src/zip';
import { blank, cat, font, format, narrow, number, rec, str16, u16, u32, wide, workbookStream, xf } from './xls-fixtures';

const one = (records: Uint8Array, head?: Uint8Array) => readXlsStream(workbookStream({ head, sheets: [{ name: 'Sheet1', records }] })).workbook;

const row = (index: number, height: number, flags: number, xfIndex = 0) => rec(0x0208, u16(index), u16(0), u16(0), u16(height), u16(0), u16(0), u16(flags | 0x0100), u16(xfIndex));
const colInfo = (first: number, last: number, width: number, xfIndex = 15, flags = 0) => rec(0x007d, u16(first), u16(last), u16(width), u16(xfIndex), u16(flags), u16(0));

test('ROW: height, hidden, custom height, default format, outline', () => {
  const sheet = one(cat(row(0, 480, 0x40), row(1, 300, 0x20), row(2, 255 | 0x8000, 0), row(3, 300, 0x80 | 0x10 | 0x03, 15), number(3, 0, 1))).sheets[0];
  assert.deepEqual(sheet.rows.get(0), { height: 24, customHeight: true, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
  assert.equal(sheet.rows.get(1)?.hidden, true);
  assert.equal(sheet.rows.get(1)?.height, 15);
  // The default height is not a height of the row.
  assert.equal(sheet.rows.get(2)?.height, undefined);
  const fourth = sheet.rows.get(3);
  assert.deepEqual([fourth?.style, fourth?.outlineLevel, fourth?.collapsed], [0, 3, true]);
  assert.equal(fourth?.cells[0].value, 1);
  // Rows without cells do not count as used.
  assert.equal(sheet.rowCount, 4);
});

test('COLINFO, DEFCOLWIDTH, STANDARDWIDTH, DEFAULTROWHEIGHT', () => {
  const sheet = one(
    cat(
      rec(0x0055, u16(8)),
      rec(0x0225, u16(0), u16(255)),
      colInfo(0, 0, 4736, 15, 0x02),
      colInfo(2, 3, 2304, 15, 0x01),
      // Overlaps the range before it: only the part not covered is added. A last column of 256 means 255.
      colInfo(3, 256, 512, 16, 0x0300 | 0x1000),
    ),
    cat(font(), ...Array.from({ length: 15 }, () => xf({ style: true })), xf(), xf({ horizontal: 2 })),
  ).sheets[0];
  assert.equal(sheet.baseColWidth, 8);
  assert.equal(sheet.defaultColWidth, undefined);
  assert.equal(sheet.defaultRowHeight, 12.75);
  assert.deepEqual(sheet.columns, [
    { min: 0, max: 0, width: 18.5, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false },
    { min: 2, max: 3, width: 9, customWidth: false, hidden: true, outlineLevel: 0, collapsed: false },
    { min: 4, max: 255, width: 2, customWidth: false, hidden: false, outlineLevel: 3, collapsed: true, style: 1 },
  ]);
  const metrics = createSheetMetrics(sheet);
  assert.equal(metrics.cols.size(0), 129);
  assert.equal(metrics.cols.size(2), 0);

  const wide12 = one(rec(0x0055, u16(12))).sheets[0];
  // 12 characters of 7 pixels and 5 of padding, rounded up to a multiple of 8: 96 pixels.
  assert.equal(createSheetMetrics(wide12).cols.size(0), 96);
  const standard = one(cat(rec(0x0055, u16(12)), rec(0x0099, u16(2560)))).sheets[0];
  assert.equal(standard.defaultColWidth, 10);
});

test('MERGECELLS and DIMENSION', () => {
  const sheet = one(
    cat(
      rec(0x0200, u32(1), u32(10), u16(2), u16(5), u16(0)),
      rec(0x00e5, u16(3), u16(0), u16(1), u16(0), u16(2), u16(4), u16(4), u16(1), u16(1), u16(9), u16(3), u16(0), u16(0)),
      rec(0x00e5, u16(1), u16(7), u16(8), u16(0), u16(255)),
    ),
  ).sheets[0];
  assert.deepEqual(sheet.dimension, { top: 1, left: 2, bottom: 9, right: 4 });
  // A range whose corners are the wrong way round is left out.
  assert.deepEqual(sheet.merges, [
    { top: 0, bottom: 1, left: 0, right: 2 },
    { top: 4, bottom: 4, left: 1, right: 1 },
    { top: 7, bottom: 8, left: 0, right: 255 },
  ]);
  assert.equal(one(rec(0x0200, u32(0), u32(0), u16(0), u16(0), u16(0))).sheets[0].dimension, undefined);
});

test('WINDOW2, PANE, SELECTION, SCL: view options, frozen panes, active cell, zoom', () => {
  const window2 = (flags: number, top = 0, left = 0) => rec(0x023e, u16(flags), u16(top), u16(left), u16(64), u16(0), u16(0), u16(0), u32(0));
  const selection = (pane: number, r: number, c: number) => rec(0x001d, [pane], u16(r), u16(c), u16(0), u16(1), u16(r), u16(r), [c, c]);
  const frozen = one(cat(window2(0x0008 | 0x0100 | 0x0200 | 0x0004 | 0x0010, 4, 2), rec(0x00a0, u16(3), u16(2)), rec(0x0041, u16(1), u16(2), u16(9), u16(5), [0]), selection(3, 0, 0), selection(0, 11, 6))).sheets[0];
  assert.deepEqual(frozen.view, {
    showGridLines: false,
    showRowColHeaders: true,
    showZeros: true,
    rightToLeft: false,
    tabSelected: true,
    zoomScale: 150,
    frozenRows: 2,
    frozenCols: 1,
    firstVisibleCell: 'C5',
    topLeftCell: 'F10',
    activeCell: 'G12',
  });
  const plain = one(cat(window2(0x0002 | 0x0004 | 0x0040), rec(0x0041, u16(1000), u16(2000), u16(0), u16(0), [3]), selection(3, 2, 1))).sheets[0];
  // A split that is not frozen is not modelled.
  assert.deepEqual([plain.view.frozenRows, plain.view.frozenCols, plain.view.showGridLines, plain.view.showZeros, plain.view.rightToLeft, plain.view.activeCell], [0, 0, true, false, true, 'B3']);
  // A zoom outside what a view can show is ignored.
  assert.equal(one(rec(0x00a0, u16(1), u16(0))).sheets[0].view.zoomScale, 100);
  assert.equal(one(rec(0x00a0, u16(50), u16(1))).sheets[0].view.zoomScale, 100);
});

const STD_LINK = [0xd0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const URL_MONIKER = [0xe0, 0xc9, 0xea, 0x79, 0xf9, 0xba, 0xce, 0x11, 0x8c, 0x82, 0x00, 0xaa, 0x00, 0x4b, 0xa9, 0x0b];
const FILE_MONIKER = [0x03, 0x03, 0, 0, 0, 0, 0, 0, 0xc0, 0, 0, 0, 0, 0, 0, 0x46];
const linkString = (text: string) => [...u32(text.length + 1), ...wide(text), 0, 0];
const hlink = (top: number, bottom: number, left: number, right: number, flags: number, ...body: number[][]) => rec(0x01b8, u16(top), u16(bottom), u16(left), u16(right), STD_LINK, u32(2), u32(flags), ...body);

test('HLINK: URL, file, location inside the workbook, display text and tooltip', () => {
  const url = 'https://example.com/çay?a=1';
  const sheet = one(
    cat(
      hlink(0, 0, 0, 0, 0x01 | 0x02 | 0x10, linkString('Örnek'), URL_MONIKER, [...u32(url.length * 2 + 2), ...wide(url), 0, 0]),
      // A URL followed by the GUID and flags some writers add inside the same length.
      hlink(1, 1, 0, 1, 0x01 | 0x02 | 0x08, URL_MONIKER, [...u32(14 * 2 + 24), ...wide('http://a.test'), 0, 0, ...new Array(24).fill(7)], linkString('bölüm')),
      hlink(2, 2, 0, 0, 0x01, FILE_MONIKER, [...u16(2), ...u32(8), ...narrow('veri.xls'), 0].slice(0, 6 + 8), [...narrow('ls'), 0].slice(0, 0), [0xff, 0xff, 0xad, 0xde], new Array(20).fill(0), [...u32(4 + 2 + 20), ...u32(20), ...u16(3), ...wide('alt\\veri ğ.xls').slice(0, 20)]),
      hlink(3, 4, 1, 2, 0x08, linkString("'Sayfa 2'!A1")),
      rec(0x0800, u16(0x0800), u16(3), u16(4), u16(1), u16(2), wide('İpucu'), [0, 0]),
      // Outside the sheet: left out.
      hlink(5, 5, 300, 300, 0x08, linkString('A1')),
    ),
  ).sheets[0];
  assert.deepEqual(sheet.hyperlinks, [
    { range: { top: 0, left: 0, bottom: 0, right: 0 }, display: 'Örnek', target: url },
    { range: { top: 1, left: 0, bottom: 1, right: 1 }, target: 'http://a.test', location: 'bölüm' },
    { range: { top: 2, left: 0, bottom: 2, right: 0 }, target: '../../alt/veri ğ' },
    { range: { top: 3, left: 1, bottom: 4, right: 2 }, location: "'Sayfa 2'!A1", tooltip: 'İpucu' },
  ]);
});

test('HLINK: a link that cannot be read is left out and counted; the sheet still opens', () => {
  for (const body of [[...u32(30000), 65, 0], [...u32(0x7fffffff), 65, 0]]) {
    const { workbook, report } = readXlsStream(workbookStream({ sheets: [{ name: 'S', records: cat(number(0, 0, 1), hlink(0, 0, 0, 0, 0x08, body)) }] }));
    assert.equal(workbook.sheets[0].hyperlinks.length, 0);
    assert.equal(getCell(workbook.sheets[0], 0, 0)?.value, 1);
    assert.equal(report.skippedRecords.get(0x01b8), 1);
  }
  // A moniker of an unknown kind: the display text is kept, the target is not known.
  const unknown = hlink(0, 0, 0, 0, 0x01 | 0x10 | 0x08, linkString('Görünen'), new Array(16).fill(9), [1, 2, 3, 4]);
  assert.deepEqual(one(unknown).sheets[0].hyperlinks, [{ range: { top: 0, left: 0, bottom: 0, right: 0 }, display: 'Görünen' }]);
});

function styledWorkbook() {
  const head = cat(
    font({ name: 'Arial', size: 10 }),
    font({ name: 'Times New Roman', size: 14.5, bold: true, italic: true, strike: true, underline: 2, color: 10, script: 1, family: 1, charset: 162 }),
    font({ underline: 0x21, script: 2 }),
    font(),
    // The index 4 does not exist: this record is the font XF records call 5.
    font({ name: 'Courier New', color: 12 }),
    format(5, '#,##0\\ "TL"'),
    format(164, 'dd/mm/yyyy'),
    ...Array.from({ length: 15 }, () => xf({ style: true })),
    xf({ parent: 0 }), // 15: the default cell format
    xf({ font: 1, format: 164, horizontal: 2, vertical: 1, wrap: true, rotation: 45, indent: 3, shrink: true, pattern: 1, fg: 13, bg: 64, lines: [1, 2, 5, 6], lineColors: [8, 10, 12, 64], locked: false, hidden: true, parent: 3 }),
    xf({ font: 5, format: 2, pattern: 17, fg: 9, bg: 10, diagonal: { style: 7, color: 11, up: true } }),
    xf({ font: 99, format: 5, rotation: 255, vertical: 0 }),
    // The same fill and border as record 16: the tables hold each once.
    xf({ pattern: 1, fg: 13, bg: 64, lines: [1, 2, 5, 6], lineColors: [8, 10, 12, 64] }),
    rec(0x0293, u16(0x8000), [0, 0xff]),
    rec(0x0293, u16(0x8003), [3, 0xff]),
    rec(0x0293, u16(0x0005), str16('Özel stil')),
    rec(0x0092, u16(3), [0x11, 0x22, 0x33, 0], [0xaa, 0xbb, 0xcc, 0], [0x01, 0x02, 0x03, 0]),
  );
  return readXlsStream(workbookStream({ head, sheets: [{ name: 'S', records: cat(number(0, 0, 1, 16), number(0, 1, 2, 17), number(0, 2, 3, 18), blank(0, 3, 15), number(0, 4, 5, 999), number(0, 5, 6, 3), number(0, 6, 7, 19)) }] })).workbook;
}

test('FONT, FORMAT, XF, PALETTE: cell styles resolve as they do for xlsx', () => {
  const workbook = styledWorkbook();
  const { styles } = workbook;
  const sheet = workbook.sheets[0];
  assert.equal(styles.cellStyleXfs.length, 15);
  assert.equal(styles.cellXfs.length, 5 + 15);
  assert.equal(styles.fonts.length, 5);

  const first = getCell(sheet, 0, 0);
  assert.equal(first?.style, 1);
  const a = resolveCellStyle(styles, first?.style ?? -1);
  assert.deepEqual(a.font, { name: 'Times New Roman', size: 14.5, bold: true, italic: true, strike: true, underline: 'double', color: { indexed: 10 }, vertAlign: 'superscript', family: 1, charset: 162 });
  assert.equal(a.numFmt, 'dd/mm/yyyy');
  assert.equal(a.numFmtId, 164);
  assert.deepEqual(a.alignment, { horizontal: 'center', vertical: 'center', wrapText: true, shrinkToFit: true, textRotation: 45, indent: 3 });
  assert.deepEqual(a.fill, { type: 'pattern', pattern: 'solid', fg: { indexed: 13 }, bg: { indexed: 64 } });
  assert.equal(resolveColor(a.fill.type === 'pattern' ? a.fill.fg : undefined, styles), '#FFFF00');
  assert.deepEqual(a.border, {
    diagonalUp: false,
    diagonalDown: false,
    left: { style: 'thin', color: { indexed: 8 } },
    right: { style: 'medium', color: { indexed: 10 } },
    top: { style: 'thick', color: { indexed: 12 } },
    bottom: { style: 'double', color: { indexed: 64 } },
  });
  // The palette record replaces the colours from index 8 on.
  assert.equal(resolveColor(a.border.left?.color, styles), '#112233');
  assert.equal(resolveColor(a.border.right?.color, styles), '#010203');
  assert.equal(resolveColor(a.border.top?.color, styles), '#0000FF');
  assert.equal(resolveColor(a.border.bottom?.color, styles), undefined);
  assert.equal(resolveColor(a.font.color, styles), '#010203');
  assert.deepEqual(styles.cellXfs[1].protection, { locked: false, hidden: true });
  assert.equal(styles.cellXfs[1].xfId, 3);
  assert.equal(styles.cellXfs[0].protection, undefined);

  const b = resolveCellStyle(styles, getCell(sheet, 0, 1)?.style ?? -1);
  assert.equal(b.font.name, 'Courier New');
  assert.equal(b.numFmt, '0.00');
  assert.deepEqual(b.fill, { type: 'pattern', pattern: 'gray125', fg: { indexed: 9 }, bg: { indexed: 10 } });
  assert.deepEqual(b.border, { diagonalUp: true, diagonalDown: false, diagonal: { style: 'hair', color: { indexed: 11 } } });
  assert.deepEqual(b.alignment, {});

  const c = resolveCellStyle(styles, getCell(sheet, 0, 2)?.style ?? -1);
  // A font index outside the table reads as the first font; a FORMAT record replaces a built-in code.
  assert.equal(c.font.name, 'Arial');
  assert.equal(c.numFmt, '#,##0\\ "TL"');
  assert.equal(formatCodeOf(styles, 14), 'mm-dd-yy');
  assert.deepEqual(c.alignment, { vertical: 'top', textRotation: 255 });

  // The default cell format of the file is the first format of the model.
  assert.equal(getCell(sheet, 0, 3)?.style, 0);
  // An XF index outside the table reads as the default format; a style XF named by a cell gets a cell format of its own.
  assert.equal(getCell(sheet, 0, 4)?.style, 0);
  const viaStyle = getCell(sheet, 0, 5)?.style ?? -1;
  assert.ok(viaStyle >= 5);
  assert.equal(styles.cellXfs[viaStyle].xfId, 3);
  // Equal fills and borders are stored once.
  assert.equal(styles.cellXfs[getCell(sheet, 0, 6)?.style ?? -1].fillId, styles.cellXfs[1].fillId);
  assert.equal(styles.cellXfs[getCell(sheet, 0, 6)?.style ?? -1].borderId, styles.cellXfs[1].borderId);
  assert.deepEqual(styles.fills.slice(0, 2), [{ type: 'pattern', pattern: 'none' }, { type: 'pattern', pattern: 'gray125' }]);

  assert.deepEqual(styles.cellStyles, [
    { name: 'Normal', xfId: 0, builtinId: 0 },
    { name: 'Comma', xfId: 3, builtinId: 3 },
    { name: 'Özel stil', xfId: 5 },
  ]);
  assert.equal(styles.indexedColors?.length, 64);
  assert.equal(styles.indexedColors?.[8], '112233');
  assert.equal(styles.indexedColors?.[11], '00FF00');
});

const fullColor = (type: number, value: number, tint = 0) => [...u16(type), ...u16(tint & 0xffff), ...u32(value), 0, 0, 0, 0, 0, 0, 0, 0];
const xfExt = (index: number, ...props: [type: number, body: number[]][]) =>
  rec(0x087d, u16(0x087d), new Array(10).fill(0), u16(0), u16(index), u16(0), u16(props.length), ...props.map(([type, body]) => [...u16(type), ...u16(body.length + 4), ...body]));

test('XFEXT: exact colours replace the palette indexes; theme colours resolve with the Office theme', () => {
  const head = cat(
    font({ color: 8 }),
    ...Array.from({ length: 15 }, () => xf({ style: true })),
    xf(),
    xf({ pattern: 1, fg: 22, bg: 64, lines: [1, 1, 1, 1], lineColors: [8, 8, 8, 8] }),
    xf({ pattern: 1, fg: 22 }),
    xfExt(16, [4, fullColor(2, 0xff563412)], [13, fullColor(3, 1)], [7, fullColor(3, 4, -16384)], [9, fullColor(1, 10)], [15, u16(7)], [0x99, [1, 2, 3]], [5, fullColor(0, 0)]),
    xfExt(17, [4, fullColor(3, 4, 13107)]),
    // An index outside the XF table, and a record cut short: ignored.
    xfExt(999, [4, fullColor(2, 0)]),
    rec(0x087d, u16(0x087d), new Array(10).fill(0), u16(0), u16(16), u16(0), u16(5), u16(4), u16(20), [1, 2]),
  );
  const workbook = readXlsStream(workbookStream({ head, sheets: [{ name: 'S', records: cat(number(0, 0, 1, 16), number(0, 1, 1, 17), number(0, 2, 1, 15)) }] })).workbook;
  const { styles } = workbook;
  const a = resolveCellStyle(styles, getCell(workbook.sheets[0], 0, 0)?.style ?? -1);
  assert.deepEqual(a.fill, { type: 'pattern', pattern: 'solid', fg: { rgb: 'FF123456' }, bg: { indexed: 64 } });
  assert.deepEqual(a.font.color, { theme: 1 });
  assert.equal(resolveColor(a.font.color, styles), '#000000');
  assert.deepEqual(a.border.top?.color, { theme: 4, tint: -16384 / 32767 });
  assert.deepEqual(a.border.left?.color, { indexed: 10 });
  assert.deepEqual(a.border.bottom?.color, { indexed: 8 });
  assert.equal(a.alignment.indent, 7);
  // The font of the other cells keeps its own colour.
  assert.deepEqual(resolveCellStyle(styles, getCell(workbook.sheets[0], 0, 2)?.style ?? -1).font.color, { indexed: 8 });
  const b = resolveCellStyle(styles, getCell(workbook.sheets[0], 0, 1)?.style ?? -1);
  assert.equal(b.fill.type === 'pattern' && b.fill.fg?.theme, 4);
  assert.equal(resolveColor({ theme: 4 }, styles), '#4F81BD');
  assert.equal(styles.themeColors.length, 12);
});

test('THEME: the colours of a custom theme are read from the archive in the record', async () => {
  const zip = new ZipArchive();
  const slot = (name: string, rgb: string) => `<a:${name}><a:srgbClr val="${rgb}"/></a:${name}>`;
  const xml =
    '<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1>' +
    slot('dk2', '222222') + slot('lt2', 'EEEEEE') + slot('accent1', 'AA0000') + slot('accent2', '00AA00') + slot('accent3', '0000AA') + slot('accent4', 'AAAA00') + slot('accent5', '00AAAA') + slot('accent6', 'AA00AA') + slot('hlink', '0000FF') + slot('folHlink', '800080') +
    '</a:clrScheme></a:themeElements></a:theme>';
  zip.set('theme/theme/themeManager.xml', new TextEncoder().encode('<a:themeManager xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'));
  zip.set('theme/theme/theme1.xml', new TextEncoder().encode(xml + ' '.repeat(3000)));
  const archive = await zip.write();
  // The archive is split between the THEME record and a continuation, as larger ones are in real files.
  const cut = Math.floor(archive.length / 2);
  const theme = cat(rec(0x0896, u16(0x0896), new Array(10).fill(0), u32(0), archive.subarray(0, cut)), rec(0x087f, u16(0x087f), new Array(10).fill(0), archive.subarray(cut)));
  const head = cat(font(), ...Array.from({ length: 15 }, () => xf({ style: true })), xf(), theme);
  const { styles } = readXlsStream(workbookStream({ head, sheets: [{ name: 'S' }] })).workbook;
  assert.deepEqual(styles.themeColors, ['FEFEFE', '111111', 'EEEEEE', '222222', 'AA0000', '00AA00', '0000AA', 'AAAA00', '00AAAA', 'AA00AA', '0000FF', '800080']);

  // A version other than 0 means the Office theme; a damaged archive falls back to it as well.
  const office = readXlsStream(workbookStream({ head: cat(font(), xf({ style: true }), xf(), rec(0x0896, u16(0x0896), new Array(10).fill(0), u32(124226))), sheets: [{ name: 'S' }] })).workbook.styles.themeColors;
  assert.equal(office[4], '4F81BD');
  const broken = cat(font(), xf({ style: true }), xf(), rec(0x0896, u16(0x0896), new Array(10).fill(0), u32(0), archive.subarray(0, 100)), xfExt(1, [4, fullColor(3, 5)]));
  assert.equal(readXlsStream(workbookStream({ head: broken, sheets: [{ name: 'S' }] })).workbook.styles.themeColors[5], 'C0504D');
});

test('a workbook without style records gets the default style sheet', () => {
  const workbook = readXlsStream(workbookStream({ head: new Uint8Array(0), sheets: [{ name: 'S', records: number(0, 0, 1, 7) }] })).workbook;
  assert.equal(getCell(workbook.sheets[0], 0, 0)?.style, 0);
  assert.equal(resolveCellStyle(workbook.styles, 0).numFmt, 'General');
});
