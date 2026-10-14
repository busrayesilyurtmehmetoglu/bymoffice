import assert from 'node:assert/strict';
import { test } from 'node:test';
import { isDeepStrictEqual } from 'node:util';
import { decodeXstring, parseRunFont } from '../src/formats/xlsx/shared-strings';
import { parseWorksheet } from '../src/formats/xlsx/worksheet';
import { encodeXstring, runPropertiesXml, writeStyles, writeWorksheet, writeXlsx } from '../src/formats/xlsx/write';
import { MAX_COLS, MAX_ROWS, createSheet, emptyStyleSheet, getCell, open, parseCellRef, parseRange, parseXml, ZipArchive } from '../src/index';
import type { Cell, Font, OfficeFile, Sheet, Workbook, XmlElement } from '../src/index';
import { resolveCellStyle } from '../src/model/sheet/color';
import { setCellInput, styleAt } from '../src/ops/commands';
import { formatRange, mergeRange, setColumnWidth, setColumnsHidden, setRowHeight, setRowsHidden, unmergeRange } from '../src/ops/format';
import { History, applyOp, ensureCellFormat, ensureNumberFormat, type Op } from '../src/ops/ops';
import { XmlScanner } from '../src/xml-scan';
import { buildXlsx, sampleWorkbook, utf8, worksheetXml, type XlsxSpec } from './fixtures';

// --- helpers ---------------------------------------------------------------

// node:assert builds a diff for the message of a failed comparison, which for large values (a
// package, a model, a 40 MB part) takes minutes. Everything here is compared first and described
// briefly when it differs.

const brief = (text: string, at: number) => JSON.stringify(text.slice(Math.max(0, at - 60), at + 100));

function sameText(actual: string, expected: string, label = 'text'): void {
  if (actual === expected) return;
  let at = 0;
  while (at < actual.length && at < expected.length && actual.charCodeAt(at) === expected.charCodeAt(at)) at++;
  assert.fail(`${label}: differs at ${at} (lengths ${actual.length} and ${expected.length})\n  actual:   ${brief(actual, at)}\n  expected: ${brief(expected, at)}`);
}

function sameBytes(actual: Uint8Array, expected: Uint8Array, label = 'bytes'): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

/** JSON with Maps and Sets spelled out and undefined members left out, as a canonical text of a value. */
const canonical = (value: unknown) =>
  JSON.stringify(value, (_key, member) => (member instanceof Map ? { map: [...member] } : member instanceof Set ? { set: [...member] } : member));

/**
 * assert.ok without its message: for a failed assert.ok, node:assert looks the call up in the
 * source to quote it, which on this file can take minutes. The stack names the line.
 */
function ok(condition: unknown, label?: string): void {
  if (!condition) assert.fail(label === undefined ? 'the condition is false' : label.length > 300 ? `${label.slice(0, 300)}…` : label);
}

/** assert.equal without the diff. */
function eq(actual: unknown, expected: unknown, label?: string): void {
  if (typeof actual === 'string' && typeof expected === 'string') sameText(actual, expected, label);
  else if (!Object.is(actual, expected)) assert.fail(`${label ?? 'value'}: ${String(actual)} is not ${String(expected)}`);
}

/** assert.deepEqual without the diff. */
function deq(actual: unknown, expected: unknown, label?: string): void {
  if (actual instanceof Uint8Array && expected instanceof Uint8Array) sameBytes(actual, expected, label);
  else if (!isDeepStrictEqual(actual, expected)) sameText(canonical(actual) ?? 'undefined', canonical(expected) ?? 'undefined', label ?? 'value');
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const decoder = new TextDecoder();

interface Loaded {
  bytes: Uint8Array;
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { bytes, file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const reopen = async (bytes: Uint8Array) => (await open(bytes)).workbook();

async function partText(bytes: Uint8Array, name: string): Promise<string> {
  return decoder.decode(await ZipArchive.read(bytes).read(name));
}

const at = (ref: string) => parseCellRef(ref)!;

function put(doc: Loaded, sheet: number, ref: string, cell: Cell | undefined): void {
  doc.history.apply([{ type: 'setCells', sheet, cells: [{ ...at(ref), cell }] }], undefined, undefined);
}

function type(doc: Loaded, sheet: number, ref: string, text: string): void {
  const { row, col } = at(ref);
  doc.history.apply(setCellInput(doc.wb, sheet, row, col, text), undefined, undefined);
}

const structural = (doc: Loaded, op: Op) => doc.history.apply([op], undefined, undefined);

/** Drops undefined members, which deepEqual would tell apart from missing ones. */
function defined<T extends object>(value: T): T {
  const out: Record<string, unknown> = {};
  for (const [key, member] of Object.entries(value)) if (member !== undefined) out[key] = member;
  return out as T;
}

/** What a saved file has to bring back: everything the model holds about cells, rows and ranges. */
function picture(wb: Workbook): unknown {
  return {
    names: wb.definedNames.map((name) => [name.name, name.value, name.localSheet]),
    sheets: wb.sheets.map((sheet) => ({
      name: sheet.name,
      columns: sheet.columns.map(defined),
      merges: sheet.merges,
      hyperlinks: sheet.hyperlinks.map(defined),
      rows: [...sheet.rows.entries()]
        .sort((a, b) => a[0] - b[0])
        // A row with neither cells nor properties is not written, and need not be.
        .filter(([, row]) => row.cells.some(() => true) || row.height !== undefined || row.customHeight || row.hidden || row.style !== undefined || row.outlineLevel > 0 || row.collapsed || row.attrs !== undefined)
        .map(([index, row]) => {
          const cells: unknown[] = [];
          row.cells.forEach((cell, col) =>
            cells.push(
              defined({
                col,
                kind: cell.kind,
                value: cell.value,
                style: cell.style,
                formula: cell.formula && defined({ ...cell.formula, attrs: cell.formula.attrs && [...cell.formula.attrs] }),
                runs: cell.runs,
                cm: cell.cm,
                vm: cell.vm,
                ph: cell.ph,
              }),
            ),
          );
          return defined({
            index,
            height: row.height,
            customHeight: row.customHeight,
            hidden: row.hidden,
            style: row.style,
            outlineLevel: row.outlineLevel,
            collapsed: row.collapsed,
            attrs: row.attrs,
            cells,
          });
        }),
    })),
  };
}

/** Saves, reopens and checks that the reopened model is the edited one. Returns the saved bytes. */
async function roundTrip(doc: Loaded): Promise<Uint8Array> {
  const saved = await save(doc);
  deq(picture(await reopen(saved)), picture(doc.wb));
  return saved;
}

/** Every entry but the named ones still has the compressed bytes it had. */
function assertOtherPartsUntouched(before: Uint8Array, after: Uint8Array, changed: string[]): void {
  const a = ZipArchive.read(before);
  const b = ZipArchive.read(after);
  deq(b.names(), a.names());
  for (const name of a.names()) {
    if (changed.indexOf(name) >= 0) continue;
    deq(b.entry(name)!.raw, a.entry(name)!.raw, name);
  }
}

const SHEET1 = 'xl/worksheets/sheet1.xml';
const SHEET2 = 'xl/worksheets/sheet2.xml';

const STYLES =
  `${XML}<styleSheet xmlns="${NS}">` +
  '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₺&quot;"/></numFmts>' +
  '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><name val="Calibri"/></font></fonts>' +
  '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" justifyLastLine="1"/></xf>' +
  '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>' +
  '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
  '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs>' +
  '<extLst><ext uri="{x}"><gelecek a="1"/></ext></extLst>' +
  '</styleSheet>';

const ROW3 = '<row r="3" spans="1:4" thickBot="1" bilinmeyen="x&amp;y"><c r="A3" gelecek="1"><v>7</v><extLst><ext uri="x"/></extLst></c><!-- yorum --></row>';
const ROW1 = '<row r="1" spans="1:4" ht="24" customHeight="1" x14ac:dyDescent="0.25"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c></row>';
const TAIL_ELEMENTS = [
  '<conditionalFormatting sqref="B2:B4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>$B$2</formula></cfRule></conditionalFormatting>',
  '<conditionalFormatting sqref="A3"><cfRule type="expression" dxfId="0" priority="2"><formula>A3&gt;1</formula></cfRule></conditionalFormatting>',
  '<dataValidations count="2"><dataValidation type="list" sqref="C2:C4"><formula1>"a,b"</formula1></dataValidation>' +
    '<dataValidation type="whole" sqref="D3"><formula1>1</formula1><formula2>B4</formula2></dataValidation></dataValidations>',
  '<hyperlinks><hyperlink ref="A2" r:id="rId1" tooltip="Çay"/><hyperlink ref="A4" location="Diğer!A1"/></hyperlinks>',
  '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
  '<pageSetup  orientation="landscape" />',
  '<rowBreaks count="2" manualBreakCount="2"><brk id="3" max="16383" man="1"/><brk id="5" max="16383" man="1"/></rowBreaks>',
  '<extLst><ext uri="{x}"><x14:gelecek xmlns:x14="urn:x14"><!-- yorum --><![CDATA[ham <veri>]]></x14:gelecek></ext></extLst>',
];

const RICH_SHEET =
  '<sheetPr><tabColor rgb="FF00B050"/></sheetPr>' +
  '<dimension ref="A1:D6"/>' +
  '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
  '<selection pane="bottomLeft" activeCell="B3" sqref="B3"/></sheetView></sheetViews>' +
  '<sheetFormatPr defaultRowHeight="15"/>' +
  '<cols><col min="1" max="1" width="18.5" customWidth="1" bestFit="1"/><col min="3" max="4" width="9" hidden="1"/></cols>' +
  '<sheetData>' +
  ROW1 +
  '<row r="2" spans="1:4" x14ac:dyDescent="0.25"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>12.5</v></c><c r="C2" t="b"><v>1</v></c><c r="D2" t="e"><v>#DIV/0!</v></c></row>' +
  ROW3 +
  '<row r="4"><c r="A4" t="inlineStr"><is><t>Toplam</t></is></c><c r="B4" s="2"><f>SUM(B2:B3)</f><v>19.5</v></c></row>' +
  '<row r="6"><c r="A6"><f t="shared" ref="A6:C6" si="0">A2*2</f><v>0</v></c><c r="B6"><f t="shared" si="0"/><v>25</v></c><c r="C6"><f t="shared" si="0"/><v>2</v></c></row>' +
  '</sheetData>' +
  '<autoFilter ref="A1:D4"/>' +
  '<mergeCells count="1"><mergeCell ref="A8:C9"/></mergeCells>' +
  TAIL_ELEMENTS.join('');

const RICH: XlsxSpec = {
  workbookHead: '<workbookPr date1904="1"/><bookViews><workbookView activeTab="0"/></bookViews>',
  workbookTail: '<definedNames><definedName name="Alan">Veri!$B$2:$B$4</definedName><definedName name="Yerel" localSheetId="1">Veri!$A$6</definedName></definedNames><calcPr calcId="191029"/><extLst><ext uri="{w}"/></extLst>',
  sharedStrings:
    '<si><t>Ürün</t></si><si><t xml:space="preserve"> Çay </t></si>' +
    '<si><r><t>Kalın</t></r><r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t> şeker</t></r>' +
    '<rPh sb="0" eb="1"><t>ふりがな</t></rPh><phoneticPr fontId="1"/></si>',
  styles: STYLES,
  sheets: [
    { name: 'Veri', links: { rId1: 'https://example.com/çay' }, xml: worksheetXml(RICH_SHEET) },
    { name: 'Diğer', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><f>Veri!B2+1</f><v>13.5</v></c></row></sheetData>') },
  ],
};

/** A sheet alone in a workbook. */
const single = (inner: string, extra: Partial<XlsxSpec> = {}) => buildXlsx({ sheets: [{ name: 'Sayfa', xml: worksheetXml(inner) }], ...extra });

function sheetFromXml(xml: string, shared: { text: string }[] = []): Sheet {
  const sheet = createSheet({ name: 'Sayfa', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(utf8(xml), shared, sheet);
  return sheet;
}

const bareWorkbook = (sheet: Sheet, shared: { text: string }[] = []): Workbook => ({
  sheets: [sheet],
  styles: emptyStyleSheet(),
  sharedStrings: shared,
  definedNames: [],
  date1904: false,
  activeSheet: 0,
});

/** Deterministic PRNG (mulberry32). */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** Text made of everything that is awkward to write. */
function nastyText(rand: () => number): string {
  const pieces = ['_', 'x', '_x', '000D', '005F', '0041', 'd83d', '_x000D_', '_x005F_', '\r', '\n', '\t', ' ', '\u0000', '\u0001', '\u001f', '\ud83d', '\ude00', '😀', '\uFFFE', '\uFFFF', '&', '<', '>', '"', "'", ']]>', 'ğ', 'İ', 'a', 'F'];
  let out = '';
  const length = Math.floor(rand() * 12);
  for (let i = 0; i < length; i++) out += pieces[Math.floor(rand() * pieces.length)];
  return out;
}

// --- nothing changed ----------------------------------------------------------

test('write: saving an unmodified workbook returns the original bytes', async () => {
  for (const bytes of [await sampleWorkbook(), await buildXlsx(RICH)]) {
    const doc = await load(bytes);
    deq(await save(doc), bytes);
    // Not even after the model was read and saved twice.
    deq(await save(doc), bytes);
  }
});

test('write: a change that was undone before the first save still writes an equivalent file', async () => {
  const doc = await load(await buildXlsx(RICH));
  const before = picture(doc.wb);
  type(doc, 0, 'B2', '99');
  doc.history.undo();
  const saved = await save(doc);
  deq(picture(await reopen(saved)), before);
});

// --- one edited cell ------------------------------------------------------------

test('write: editing one cell leaves everything else as it was', async () => {
  const doc = await load(await buildXlsx(RICH));
  type(doc, 0, 'B2', '99');
  const saved = await roundTrip(doc);

  // The other sheet, the strings, the styles and the sheet relationships keep their compressed bytes.
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml']);

  const before = await partText(doc.bytes, SHEET1);
  const after = await partText(saved, SHEET1);
  // Everything up to the changed row and everything after it is the original text.
  const changedRow = '<row r="2" x14ac:dyDescent="0.25"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>99</v></c><c r="C2" t="b"><v>1</v></c><c r="D2" t="e"><v>#DIV/0!</v></c></row>';
  const start = before.indexOf('<row r="2"');
  const end = before.indexOf(ROW3);
  eq(after, before.slice(0, start) + changedRow + before.slice(end));
  // In particular: rows with unknown attributes and children, and the elements the model does not hold, in order.
  ok(after.includes(ROW1) && after.includes(ROW3));
  ok(after.includes(TAIL_ELEMENTS.join('') + '</worksheet>'));
  eq(getCell((await reopen(saved)).sheets[0], 1, 1)!.value, 99);
});

test('write: prolog, root attributes, comments and white space of the sheet part are kept', async () => {
  const xml =
    '<?xml version="1.0" encoding="UTF-8"?>\n<!-- baş -->\n<x:worksheet xmlns:x="' + NS + '" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"  mc:Ignorable="x14ac">\n' +
    '  <!-- görünüm -->\n  <x:dimension ref="A1:A2"/>\n  <x:sheetData kayit="1">\n    <x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>\n    <!-- ara -->\n    <x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>\n  </x:sheetData>\n' +
    '  <?pi bilgi?>\n  <x:pageMargins left="1"/>\n</x:worksheet>\n<!-- son -->\n';
  const sheet = sheetFromXml(xml);
  const wb = bareWorkbook(sheet);
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 1, col: 1, cell: { kind: 'string', value: 'yeni', style: 0 } }] });
  const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
  eq(
    out,
    xml
      .replace('ref="A1:A2"', 'ref="A1:B2"')
      .replace('<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>', '<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c><x:c r="B2" t="inlineStr"><x:is><x:t>yeni</x:t></x:is></x:c></x:row>'),
  );
  const again = sheetFromXml(out);
  eq(getCell(again, 1, 1)!.value, 'yeni');
});

test('write: UTF-16 input and CRLF line endings come out as UTF-8', () => {
  const xml = '<?xml version="1.0" encoding="UTF-16"?>\r\n<worksheet>\r\n<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>İş</t></is></c></row></sheetData>\r\n</worksheet>';
  const bytes = new Uint8Array(2 + xml.length * 2);
  bytes.set([0xff, 0xfe]);
  for (let i = 0; i < xml.length; i++) {
    bytes[2 + i * 2] = xml.charCodeAt(i) & 0xff;
    bytes[3 + i * 2] = xml.charCodeAt(i) >> 8;
  }
  const sheet = sheetFromXml(xml);
  const wb = bareWorkbook(sheet);
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 4, col: 0, cell: { kind: 'number', value: 5, style: 0 } }] });
  const out = decoder.decode(writeWorksheet(bytes, sheet, wb));
  ok(out.startsWith('<?xml version="1.0" encoding="UTF-8"?>\n<worksheet>'));
  const again = sheetFromXml(out);
  eq(getCell(again, 0, 0)!.value, 'İş');
  eq(getCell(again, 4, 0)!.value, 5);
});

// --- cell and formula kinds -------------------------------------------------------

test('write: every kind of cell and formula is written and read back', async () => {
  const doc = await load(await buildXlsx(RICH));
  const bold: Font = { bold: true, italic: false, strike: false, underline: 'none' };
  const cells: [string, Cell][] = [
    ['A10', { kind: 'number', value: 42, style: 0 }],
    ['B10', { kind: 'number', value: -1.5e-7, style: 2 }],
    ['C10', { kind: 'number', value: 0.1 + 0.2, style: 0 }],
    ['D10', { kind: 'number', value: 1e21, style: 0 }],
    ['E10', { kind: 'number', value: 1.7976931348623157e308, style: 0 }],
    ['F10', { kind: 'boolean', value: true, style: 0 }],
    ['G10', { kind: 'boolean', value: false, style: 1 }],
    ['H10', { kind: 'error', value: '#N/A', style: 0 }],
    ['I10', { kind: 'date', value: '2026-10-01T12:30:00', style: 0 }],
    ['J10', { kind: 'blank', value: null, style: 2 }],
    ['K10', { kind: 'string', value: 'Ağrı İstanbul Şırnak çöğü 😀🚢 & <etiket> "tırnak"', style: 0 }],
    ['L10', { kind: 'string', value: '', style: 0 }],
    ['A11', { kind: 'number', value: 3, style: 0, formula: { text: 'A10/14&"<ş>"', type: 'normal' } }],
    ['B11', { kind: 'string', value: 'sonuç ', style: 0, formula: { text: 'K10&" "', type: 'normal' } }],
    ['C11', { kind: 'blank', value: null, style: 0, formula: { text: 'NOW()', type: 'normal' } }],
    ['D11', { kind: 'error', value: '#DIV/0!', style: 0, formula: { text: '1/0', type: 'normal' } }],
    ['E11', { kind: 'boolean', value: true, style: 0, formula: { text: 'A10>1', type: 'normal' } }],
    ['A12', { kind: 'number', value: 1, style: 0, formula: { text: 'A10*2', type: 'shared', ref: 'A12:C12', si: 7 } }],
    ['B12', { kind: 'number', value: 2, style: 0, formula: { text: '', type: 'shared', si: 7 } }],
    ['C12', { kind: 'number', value: 3, style: 0, formula: { text: '', type: 'shared', si: 7 } }],
    ['A13', { kind: 'number', value: 1, style: 0, formula: { text: 'TRANSPOSE(A10:B10)', type: 'array', ref: 'A13:A14', attrs: new Map([['aca', '1'], ['ca', '1']]) } }],
    ['B13', { kind: 'number', value: 5, style: 0, formula: { text: '', type: 'dataTable', ref: 'B13:B14', attrs: new Map([['dt2D', '1'], ['r1', 'A1'], ['r2', 'B1']]) } }],
    ['C13', { kind: 'number', value: 5, style: 0, formula: { text: ' 1+1', type: 'normal', attrs: new Map([['xml:space', 'preserve']]) } }],
    ['D13', { kind: 'number', value: 5, style: 0, formula: { text: 'X()', type: 'normal', attrs: new Map([['t', 'gelecek']]) } }],
    ['A15', { kind: 'string', value: 'kalın ve düz', style: 0, runs: [{ text: 'kalın', font: bold }, { text: ' ve düz' }] }],
    ['B15', { kind: 'number', value: 8, style: 0, cm: 1, vm: 2, ph: true }],
    ['C15', { kind: 'blank', value: null, style: 0, cm: 3 }],
  ];
  doc.history.apply([{ type: 'setCells', sheet: 0, cells: cells.map(([ref, cell]) => ({ ...at(ref), cell })) }], undefined, undefined);
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<c r="A10"><v>42</v></c><c r="B10" s="2"><v>-1.5e-7</v></c><c r="C10"><v>0.30000000000000004</v></c><c r="D10"><v>1e+21</v></c>'));
  ok(xml.includes('<c r="J10" s="2"/>'));
  ok(xml.includes('<c r="A12"><f t="shared" ref="A12:C12" si="7">A10*2</f><v>1</v></c><c r="B12"><f t="shared" si="7"/><v>2</v></c>'));
  ok(xml.includes('<c r="C11"><f>NOW()</f></c>'));
  ok(xml.includes('<c r="B15" cm="1" vm="2" ph="1"><v>8</v></c><c r="C15" cm="3"/>'));
  ok(xml.includes('<c r="B11" t="str"><f>K10&amp;" "</f><v>sonuç </v></c>'));
});

test('write: a number that a file cannot hold is written as an error, not dropped', () => {
  const xml = worksheetXml('<sheetData/>');
  const sheet = sheetFromXml(xml);
  const wb = bareWorkbook(sheet);
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: NaN, style: 0 } }, { row: 0, col: 1, cell: { kind: 'number', value: Infinity, style: 0 } }] });
  const again = sheetFromXml(decoder.decode(writeWorksheet(utf8(xml), sheet, wb)));
  deq(getCell(again, 0, 0), { kind: 'error', value: '#NUM!', style: 0 });
  deq(getCell(again, 0, 1), { kind: 'error', value: '#NUM!', style: 0 });
});

test('write: shared strings are referred to while the text is the same, and written inline once it changed', async () => {
  const doc = await load(await buildXlsx(RICH));
  // B1 is the rich string with phonetic text; touching another cell of its row regenerates the row.
  type(doc, 0, 'C1', '5');
  type(doc, 0, 'A2', 'Kahve');
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<c r="A1" s="1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c><c r="C1"><v>5</v></c>'));
  ok(xml.includes('<c r="A2" t="inlineStr"><is><t>Kahve</t></is></c>'));
  // The table itself is never touched, so the phonetic run is still there.
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml']);
  ok((await partText(saved, 'xl/sharedStrings.xml')).includes('<rPh sb="0" eb="1">'));
  const wb = await reopen(saved);
  eq(getCell(wb.sheets[0], 0, 1)!.si, 2);
  eq(getCell(wb.sheets[0], 0, 1)!.runs!.length, 2);
  eq(getCell(wb.sheets[0], 1, 0)!.si, undefined);
});

test('write: a cell copied from a shared string keeps the reference; a stale index does not', () => {
  const shared = [{ text: 'bir' }, { text: 'iki' }];
  const xml = worksheetXml('<sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData>');
  const sheet = sheetFromXml(xml, shared);
  const wb = bareWorkbook(sheet, shared);
  applyOp(wb, {
    type: 'setCells',
    sheet: 0,
    cells: [
      { row: 0, col: 1, cell: { ...getCell(sheet, 0, 0)! } },
      { row: 0, col: 2, cell: { kind: 'string', value: 'üç', style: 0, si: 1 } },
      { row: 0, col: 3, cell: { kind: 'string', value: 'dört', style: 0, si: 99 } },
    ],
  });
  const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
  ok(out.includes('<c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>üç</t></is></c><c r="D1" t="inlineStr"><is><t>dört</t></is></c>'));
});

test('write: awkward text survives: control characters, escapes, spaces, line breaks, long text', async () => {
  const texts = [
    'bir\r\niki\rüç\ndört\tbeş',
    '_x000D_ görünümlü _x005F_ ve _xABCD_ ile _x00e7_',
    '_x0041',
    '_x0041\r',
    '\u0000\u0001\u0008\u000b\u000c\u001f',
    '\ud83d yalnız \ude00 vekil',
    '\uFFFE\uFFFF',
    '  baş ve son  ',
    ' ',
    '\n',
    'iki  boşluk',
    ']]> &amp; <!-- --> <?pi?> &#13;',
    'Ağrı İstanbul Şırnak ığüşöç IĞÜŞÖÇ 😀🚢👩‍👩‍👧',
    'u'.repeat(32766) + 'ş',
  ];
  const doc = await load(await single('<sheetData/>'));
  texts.forEach((text, i) => put(doc, 0, `A${i + 1}`, { kind: 'string', value: text, style: 0 }));
  texts.forEach((text, i) => put(doc, 0, `B${i + 1}`, { kind: 'string', value: text, style: 0, formula: { text: `A${i + 1}`, type: 'normal' } }));
  const saved = await roundTrip(doc);
  const wb = await reopen(saved);
  texts.forEach((text, i) => {
    eq(getCell(wb.sheets[0], i, 0)!.value, text);
    eq(getCell(wb.sheets[0], i, 1)!.value, text);
  });
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<t xml:space="preserve">bir_x000D_\niki_x000D_üç\ndört\tbeş</t>'));
  ok(xml.includes('<t>_x005F_x000D_ görünümlü _x005F_x005F_ ve _x005F_xABCD_ ile _x005F_x00e7_</t>'));
  ok(xml.includes('<t>_x0041</t>'));
  ok(xml.includes('<t xml:space="preserve">  baş ve son  </t>'));
  ok(xml.includes('<t>iki  boşluk</t>'));
});

test('write: encodeXstring is the inverse of decodeXstring', () => {
  eq(encodeXstring('düz metin'), 'düz metin');
  eq(encodeXstring('a\rb'), 'a_x000D_b');
  eq(encodeXstring('_x000D_'), '_x005F_x000D_');
  eq(encodeXstring('_x000D'), '_x000D');
  eq(encodeXstring('_x000D\r'), '_x005F_x000D_x000D_');
  eq(encodeXstring('😀'), '😀');
  eq(encodeXstring('\ud83d'), '_xD83D_');
  const rand = prng(11);
  for (let i = 0; i < 20000; i++) {
    const text = nastyText(rand);
    const encoded = encodeXstring(text);
    eq(decodeXstring(encoded), text, JSON.stringify(text));
    // What is left can be put into XML as it is, once & < > are escaped.
    ok(!/[\u0000-\u0008\u000B-\u001F\uFFFE\uFFFF]/.test(encoded), JSON.stringify(text));
    ok(!/[\uD800-\uDBFF](?![\uDC00-\uDFFF])/.test(encoded) && !/(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(encoded), JSON.stringify(text));
  }
});

test('write: any text gives well-formed XML that reads back the same', () => {
  const rand = prng(5);
  const xml = worksheetXml('<sheetData/>');
  for (let round = 0; round < 40; round++) {
    const sheet = sheetFromXml(xml);
    const wb = bareWorkbook(sheet);
    const texts: string[] = [];
    const cells = [];
    for (let i = 0; i < 50; i++) {
      const text = nastyText(rand);
      texts.push(text);
      cells.push({ row: i, col: 0, cell: { kind: 'string', value: text, style: 0 } as Cell });
      cells.push({ row: i, col: 1, cell: { kind: 'string', value: text, style: 0, runs: [{ text: text.slice(0, 2) }, { text: text.slice(2), font: { bold: true, italic: false, strike: false, underline: 'none', name: text } }] } as Cell });
      cells.push({ row: i, col: 2, cell: { kind: 'string', value: text, style: 0, formula: { text: 'A1', type: 'normal' } } as Cell });
    }
    applyOp(wb, { type: 'setCells', sheet: 0, cells });
    const out = writeWorksheet(utf8(xml), sheet, wb);
    parseXml(out);
    const again = sheetFromXml(decoder.decode(out));
    texts.forEach((text, i) => {
      for (let col = 0; col < 3; col++) eq(getCell(again, i, col)!.value, text, JSON.stringify(text));
    });
  }
});

test('write: text without an escape of its own stays well-formed', () => {
  const xml = worksheetXml('<sheetData/>');
  const sheet = sheetFromXml(xml);
  const wb = bareWorkbook(sheet);
  applyOp(wb, {
    type: 'setCells',
    sheet: 0,
    cells: [
      { row: 0, col: 0, cell: { kind: 'number', value: 1, style: 0, formula: { text: '"a\u0001b\ud83d<&>"&_x0041_', type: 'normal', attrs: new Map([['ek', 'x"\u0002\n<']]) } } },
      { row: 0, col: 1, cell: { kind: 'error', value: '#\u0000<', style: 0 } },
    ],
  });
  const out = writeWorksheet(utf8(xml), sheet, wb);
  parseXml(out);
  const again = sheetFromXml(decoder.decode(out));
  // What XML cannot carry and the formula syntax has no escape for becomes U+FFFD; "_x0041_" in a formula is a name.
  eq(getCell(again, 0, 0)!.formula!.text, '"a\uFFFDb\uFFFD<&>"&_x0041_');
  eq(getCell(again, 0, 0)!.formula!.attrs!.get('ek'), 'x"\uFFFD\n<');
  eq(getCell(again, 0, 1)!.value, '#\uFFFD<');
});

test('write: run properties are the inverse of parseRunFont', () => {
  const fonts: Font[] = [
    { bold: false, italic: false, strike: false, underline: 'none' },
    { bold: true, italic: true, strike: true, underline: 'single' },
    { bold: false, italic: false, strike: false, underline: 'double', name: 'Arial "Dar" & <x>', size: 9.5, family: 2, charset: 162, scheme: 'minor', vertAlign: 'superscript' },
    { bold: false, italic: false, strike: false, underline: 'singleAccounting', color: { rgb: 'FFFF0000' }, vertAlign: 'subscript', scheme: 'none' },
    { bold: false, italic: false, strike: false, underline: 'doubleAccounting', color: { theme: 4, tint: -0.249977111117893 }, scheme: 'major' },
    { bold: false, italic: false, strike: false, underline: 'none', color: { indexed: 10 } },
    { bold: false, italic: false, strike: false, underline: 'none', color: { auto: true } },
    { bold: false, italic: false, strike: false, underline: 'none', color: { auto: false, theme: 1 } },
  ];
  for (const prefix of ['', 'x:']) {
    for (const font of fonts) {
      const xml = runPropertiesXml(font, prefix);
      const sc = new XmlScanner(xml);
      sc.next();
      deq(parseRunFont(sc), font, xml);
    }
  }
  // Children in the order the schema requires.
  eq(
    runPropertiesXml({ bold: true, italic: true, strike: true, underline: 'double', name: 'A', charset: 1, family: 2, color: { rgb: 'FF000000' }, size: 10, vertAlign: 'subscript', scheme: 'minor' }),
    '<rPr><rFont val="A"/><charset val="1"/><family val="2"/><b/><i/><strike/><color rgb="FF000000"/><sz val="10"/><u val="double"/><vertAlign val="subscript"/><scheme val="minor"/></rPr>',
  );
});

// --- rows -----------------------------------------------------------------------

test('write: row properties and unknown row attributes are written back when the row is regenerated', async () => {
  const doc = await load(
    await single(
      '<sheetData>' +
        '<row r="1" spans="1:2" s="1" customFormat="1" ht="30" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" thickTop="1" ph="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row>' +
        '<row r="2" spans="1:2" x14ac:dyDescent="0.25"><c r="A2"><v>2</v></c></row>' +
        '<row r="3" ht="12.75"/>' +
        '</sheetData>',
      { styles: STYLES },
    ),
  );
  type(doc, 0, 'B1', '5');
  type(doc, 0, 'B2', '6');
  type(doc, 0, 'A3', '7');
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<row r="1" s="1" customFormat="1" ht="30" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" thickTop="1" ph="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1" s="1"><v>5</v></c></row>'));
  ok(xml.includes('<row r="2" x14ac:dyDescent="0.25"><c r="A2"><v>2</v></c><c r="B2"><v>6</v></c></row>'));
  ok(xml.includes('<row r="3" ht="12.75"><c r="A3"><v>7</v></c></row>'));
});

test('write: a blank cell with the default format is written where the row or column has a format', async () => {
  const doc = await load(
    await single(
      '<cols><col min="2" max="2" width="9" style="1"/></cols><sheetData><row r="1" s="2" customFormat="1"><c r="A1" s="2"><v>1</v></c></row><row r="2"><c r="B2" s="1"><v>2</v></c><c r="C2"><v>3</v></c></row></sheetData>',
      { styles: STYLES },
    ),
  );
  put(doc, 0, 'A1', { kind: 'blank', value: null, style: 0 });
  put(doc, 0, 'B2', { kind: 'blank', value: null, style: 0 });
  put(doc, 0, 'C2', undefined);
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<row r="1" s="2" customFormat="1"><c r="A1"/></row><row r="2"><c r="B2"/></row>'));
});

test('write: cells outside the used area, emptied rows and an emptied sheet', async () => {
  const doc = await load(await buildXlsx(RICH));
  type(doc, 0, 'F3', 'yan');
  type(doc, 0, 'B5', '1');
  type(doc, 0, 'XFD1048576', 'son');
  let saved = await roundTrip(doc);
  let xml = await partText(saved, SHEET1);
  ok(xml.includes('<dimension ref="A1:XFD1048576"/>'));
  // The new row 5 sits between the original rows 4 and 6; the last row comes last.
  ok(/<row r="4">.*<\/row><row r="5"><c r="B5"><v>1<\/v><\/c><\/row><row r="6">.*<\/row><row r="1048576"><c r="XFD1048576" t="inlineStr">/.test(xml));
  // A regenerated row keeps what the model holds of it: the unknown attribute and child of A3 are not part of that.
  ok(xml.includes('<row r="3" thickBot="1" bilinmeyen="x&amp;y"><c r="A3"><v>7</v></c><c r="F3" t="inlineStr"><is><t>yan</t></is></c></row>'));

  // Emptying a row removes it; a row with properties of its own stays.
  put(doc, 0, 'XFD1048576', undefined);
  put(doc, 0, 'B5', undefined);
  put(doc, 0, 'A1', undefined);
  put(doc, 0, 'B1', undefined);
  saved = await roundTrip(doc);
  xml = await partText(saved, SHEET1);
  ok(xml.includes('<dimension ref="A2:F6"/>'));
  ok(xml.includes('<sheetData><row r="1" ht="24" customHeight="1" x14ac:dyDescent="0.25"/><row r="2"'));
  ok(!xml.includes('<row r="5"') && !xml.includes('<row r="1048576"'));

  // Everything gone.
  const all: Op = { type: 'setCells', sheet: 0, cells: [] };
  for (const [index, row] of doc.wb.sheets[0].rows) row.cells.forEach((_cell, col) => all.cells.push({ row: index, col, cell: undefined }));
  doc.history.apply([all], undefined, undefined);
  saved = await roundTrip(doc);
  xml = await partText(saved, SHEET1);
  ok(xml.includes('<dimension ref="A1"/>'));
  ok(xml.includes('<sheetData><row r="1" ht="24" customHeight="1" x14ac:dyDescent="0.25"/><row r="2" x14ac:dyDescent="0.25"/><row r="3" thickBot="1" bilinmeyen="x&amp;y"/></sheetData><autoFilter ref="A1:D4"/>'));

  const empty = await load(await single('<dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><pageMargins left="1"/>'));
  put(empty, 0, 'A1', undefined);
  ok((await partText(await roundTrip(empty), SHEET1)).includes('<dimension ref="A1"/><sheetData/><pageMargins left="1"/>'));
});

test('write: a self-closing sheetData and a sheet without dimension', async () => {
  const doc = await load(await single('<sheetData/><pageMargins left="1"/>'));
  type(doc, 0, 'C3', '1');
  const xml = await partText(await roundTrip(doc), SHEET1);
  ok(xml.endsWith('<sheetData><row r="3"><c r="C3"><v>1</v></c></row></sheetData><pageMargins left="1"/></worksheet>'));
  ok(!xml.includes('dimension'));
});

test('write: out-of-order, repeated and unnumbered rows take the safe path', async () => {
  const cases = [
    '<row r="3"><c r="A3"><v>3</v></c></row><row r="1"><c r="A1"><v>1</v></c></row>',
    '<row r="1"><c r="A1"><v>1</v></c></row><row r="1"><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row>',
    '<row><c><v>1</v></c><c><v>2</v></c></row><row><c r="B2"><v>3</v></c><c><v>4</v></c></row><row r="5"><c><v>5</v></c></row><row><c><v>6</v></c></row>',
    '<row r="2"><c r="A2"><v>1</v></c></row><bilinmeyen a="1"/><row r="2"><c r="A2"><v>2</v></c></row>',
  ];
  for (const rows of cases) {
    const doc = await load(await single(`<dimension ref="A1"/><sheetData>${rows}</sheetData>`));
    type(doc, 0, 'D9', '99');
    const saved = await roundTrip(doc);
    const xml = await partText(saved, SHEET1);
    // Every row is numbered and in order now.
    const numbers = [...xml.matchAll(/<row r="(\d+)"/g)].map((m) => Number(m[1]));
    deq(numbers, [...numbers].sort((a, b) => a - b));
    eq(new Set(numbers).size, numbers.length);
    ok(!/<row>|<row [^r]/.test(xml));
  }
  // Rows in order whose cells have no references are copied as they are: the positions follow from the row.
  const doc = await load(await single('<sheetData><row r="1"><c><v>1</v></c><c><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData>'));
  type(doc, 0, 'B2', '4');
  ok((await partText(await roundTrip(doc), SHEET1)).includes('<row r="1"><c><v>1</v></c><c><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>'));
});

test('write: unexpected parts are refused with a clear error', () => {
  const sheet = sheetFromXml(worksheetXml('<sheetData/>'));
  const wb = bareWorkbook(sheet);
  assert.throws(() => writeWorksheet(utf8(`<chartsheet xmlns="${NS}"><sheetViews/></chartsheet>`), sheet, wb), /root element <chartsheet>, not a worksheet/);
  assert.throws(() => writeWorksheet(utf8(`<worksheet xmlns="${NS}"><sheetViews/></worksheet>`), sheet, wb), /has no <sheetData>/);
  assert.throws(() => writeWorksheet(utf8(`<worksheet xmlns="${NS}"/>`), sheet, wb), /has no <sheetData>/);
  assert.throws(() => writeWorksheet(utf8('<worksheet><sheetData/><sheetData/></worksheet>'), sheet, wb), /more than one <sheetData>/);
  assert.throws(() => writeWorksheet(utf8('<worksheet><sheetData><row r="1">'), sheet, wb), /Invalid XML/);
  assert.throws(() => writeWorksheet(utf8('<worksheet><sheetData/></worksheet><x/>'), sheet, wb), /more than one root element/);
  // Dialog sheets and macro sheets hold cells the same way.
  for (const root of ['dialogsheet', 'macrosheet']) {
    const xml = `<${root}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></${root}>`;
    const s = sheetFromXml(xml);
    const w = bareWorkbook(s);
    applyOp(w, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }] });
    eq(decoder.decode(writeWorksheet(utf8(xml), s, w)), xml.replace('<v>1</v>', '<v>2</v>'));
  }
});

// --- structural edits ---------------------------------------------------------------

test('write: inserting rows moves cells, formulas and every reference in the sheet', async () => {
  const doc = await load(await buildXlsx(RICH));
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 2, count: 2 });
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);

  ok(xml.includes('<dimension ref="A1:D8"/>'));
  // Regenerated rows keep what the model holds of them; "spans" and unknown cell attributes are gone.
  ok(xml.includes('<row r="5" thickBot="1" bilinmeyen="x&amp;y"><c r="A5"><v>7</v></c></row>'));
  ok(xml.includes('<c r="B6" s="2"><f>SUM(B2:B5)</f><v>19.5</v></c>'));
  // The shared formula was written out for each of its cells.
  ok(xml.includes('<row r="8"><c r="A8"><f>A2*2</f><v>0</v></c><c r="B8"><f>B2*2</f><v>25</v></c><c r="C8"><f>C2*2</f><v>2</v></c></row>'));
  ok(xml.includes('<autoFilter ref="A1:D6"/>'));
  ok(xml.includes('<mergeCells count="1"><mergeCell ref="A10:C11"/></mergeCells>'));
  ok(xml.includes('<conditionalFormatting sqref="B2:B6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>$B$2</formula></cfRule></conditionalFormatting>'));
  ok(xml.includes('<conditionalFormatting sqref="A5"><cfRule type="expression" dxfId="0" priority="2"><formula>A5&gt;1</formula></cfRule></conditionalFormatting>'));
  ok(xml.includes('<dataValidation type="list" sqref="C2:C6"><formula1>"a,b"</formula1></dataValidation><dataValidation type="whole" sqref="D5"><formula1>1</formula1><formula2>B6</formula2></dataValidation>'));
  ok(xml.includes('<hyperlinks><hyperlink ref="A2" r:id="rId1" tooltip="Çay"/><hyperlink ref="A6" location="Diğer!A1"/></hyperlinks>'));
  ok(xml.includes('<selection pane="bottomLeft" activeCell="B5" sqref="B5"/>'));
  ok(xml.includes('<rowBreaks count="2" manualBreakCount="2"><brk id="5" max="16383" man="1"/><brk id="7" max="16383" man="1"/></rowBreaks>'));
  // Elements without references are still the original text.
  for (const raw of [TAIL_ELEMENTS[4], TAIL_ELEMENTS[5], TAIL_ELEMENTS[7], '<sheetPr><tabColor rgb="FF00B050"/></sheetPr>', '<sheetFormatPr defaultRowHeight="15"/>']) ok(xml.includes(raw), raw);
  // Column ranges come from the model now.
  ok(xml.includes('<cols><col min="1" max="1" width="18.5" customWidth="1" bestFit="1"/><col min="3" max="4" width="9" hidden="1"/></cols>'));

  // The other sheet's formula and the defined names follow.
  ok((await partText(saved, SHEET2)).includes('<f>Veri!B2+1</f>'));
  const workbookXml = await partText(saved, 'xl/workbook.xml');
  ok(workbookXml.includes('<definedName name="Alan">Veri!$B$2:$B$6</definedName><definedName name="Yerel" localSheetId="1">Veri!$A$8</definedName>'));
  const wb = await reopen(saved);
  eq(wb.sheets[0].view.activeCell, 'B5');
  deq(wb.sheets[0].dimension, { top: 0, left: 0, bottom: 7, right: 3 });
});

test('write: deleting rows drops what applied only to them', async () => {
  const doc = await load(await buildXlsx(RICH));
  // Rows 2 and 3 go: the rule and the validation on A3/D3 have nothing left, the active cell is gone.
  structural(doc, { type: 'delete', sheet: 0, axis: 'row', at: 1, count: 2 });
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  ok(xml.includes('<autoFilter ref="A1:D2"/>'));
  ok(xml.includes('<conditionalFormatting sqref="B2">'));
  ok(!xml.includes('type="expression"'));
  ok(xml.includes('<dataValidations count="1"><dataValidation type="list" sqref="C2"><formula1>"a,b"</formula1></dataValidation></dataValidations>'));
  ok(xml.includes('<hyperlinks><hyperlink ref="A2" location="Diğer!A1"/></hyperlinks>'));
  ok(xml.includes('<selection pane="bottomLeft" activeCell="A1" sqref="A1"/>'));
  // The break after row 3 goes to the row before the deleted ones (the one rule of page-breaks.ts); the one after row 5 moves up.
  ok(xml.includes('<rowBreaks count="2" manualBreakCount="2"><brk id="1" max="16383" man="1"/><brk id="3" max="16383" man="1"/></rowBreaks>'));
  ok(xml.includes('<c r="B2" s="2"><f>SUM(#REF!)</f><v>19.5</v></c>'));
  ok((await partText(saved, SHEET2)).includes('<f>Veri!#REF!+1</f>'));
  parseXml(xml);

  // Deleting everything a container applied to removes the container.
  const all = await load(await buildXlsx(RICH));
  structural(all, { type: 'delete', sheet: 0, axis: 'row', at: 0, count: 9 });
  const gone = await partText(await roundTrip(all), SHEET1);
  for (const name of ['autoFilter', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'mergeCells', 'rowBreaks']) ok(!gone.includes(`<${name}`), name);
  ok(gone.includes('<sheetData/>'));
  ok(gone.includes(TAIL_ELEMENTS[7] + '</worksheet>'));
});

test('write: inserting and deleting columns', async () => {
  const doc = await load(await buildXlsx(RICH));
  structural(doc, { type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 });
  let saved = await roundTrip(doc);
  let xml = await partText(saved, SHEET1);
  ok(xml.includes('<cols><col min="1" max="1" width="18.5" customWidth="1" bestFit="1"/><col min="4" max="5" width="9" hidden="1"/></cols>'));
  ok(xml.includes('<c r="C4" s="2"><f>SUM(C2:C3)</f><v>19.5</v></c>'));
  ok(xml.includes('<conditionalFormatting sqref="C2:C4">'));
  ok(xml.includes('<formula>$C$2</formula>'));
  ok(xml.includes('<autoFilter ref="A1:E4"/>'));
  ok(xml.includes('<mergeCell ref="A8:D9"/>'));
  ok(xml.includes('activeCell="C3" sqref="C3"'));
  // Row breaks do not move with columns.
  ok(xml.includes(TAIL_ELEMENTS[6]));
  ok((await partText(saved, 'xl/workbook.xml')).includes('Veri!$C$2:$C$4'));

  structural(doc, { type: 'delete', sheet: 0, axis: 'col', at: 0, count: 2 });
  saved = await roundTrip(doc);
  xml = await partText(saved, SHEET1);
  ok(xml.includes('<cols><col min="2" max="3" width="9" hidden="1"/></cols>'));
  ok(xml.includes('<conditionalFormatting sqref="A2:A4">'));
  ok(!xml.includes('type="expression"'));
  ok(!xml.includes('<hyperlinks'));
  ok(xml.includes('<mergeCell ref="A8:B9"/>'));
});

test('write: regenerated elements are put where the schema wants them', async () => {
  const doc = await load(
    await single(
      '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><sheetProtection sheet="1"/><pageMargins left="1"/><extLst><ext uri="x"/></extLst>',
    ),
  );
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  // Ranges the file did not have: the model gets them here directly, as a future command would set them.
  const sheet = doc.wb.sheets[0];
  sheet.columns = [{ min: 1, max: 2, width: 20, customWidth: true, hidden: false, outlineLevel: 1, collapsed: false, style: 0 }];
  sheet.merges = [{ top: 3, left: 0, bottom: 4, right: 1 }];
  sheet.hyperlinks = [{ range: { top: 1, left: 0, bottom: 1, right: 0 }, location: 'Sayfa!A9', display: 'git "<&>"' }];
  const xml = await partText(await roundTrip(doc), SHEET1);
  ok(
    xml.includes(
      '</sheetViews><cols><col min="2" max="3" width="20" customWidth="1" style="0" outlineLevel="1"/></cols><sheetData><row r="2"><c r="A2"><v>1</v></c></row></sheetData>' +
        '<sheetProtection sheet="1"/><mergeCells count="1"><mergeCell ref="A4:B5"/></mergeCells><hyperlinks><hyperlink ref="A2" location="Sayfa!A9" display="git &quot;&lt;&amp;>&quot;"/></hyperlinks><pageMargins left="1"/>',
    ),
  );

  // Nothing after sheetData at all: they go to the end.
  const bare = await load(await single('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'));
  structural(bare, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  bare.wb.sheets[0].merges = [{ top: 0, left: 0, bottom: 0, right: 1 }];
  ok((await partText(await roundTrip(bare), SHEET1)).endsWith('</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>'));
});

test('write: the relationship prefix of regenerated links is the one in scope', () => {
  const links = [{ range: { top: 0, left: 0, bottom: 0, right: 0 }, rId: 'rId5' }];
  const write = (xml: string) => {
    const sheet = sheetFromXml(xml);
    sheet.name = 'S';
    const wb = bareWorkbook(sheet);
    applyOp(wb, { type: 'insert', sheet: 0, axis: 'col', at: 5, count: 1 });
    sheet.hyperlinks = links;
    const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
    parseXml(out);
    return out;
  };
  // Declared on the root under another name.
  ok(write('<worksheet xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/></worksheet>').includes('<hyperlink ref="A1" rel:id="rId5"/>'));
  // Declared on the links themselves: the prefix the original links use.
  ok(write('<worksheet><sheetData/><hyperlinks xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><hyperlink ref="A1" q:id="rId5"/></hyperlinks></worksheet>').includes('<hyperlink ref="A1" q:id="rId5"/>'));
  // Not declared anywhere: the link declares it.
  ok(write('<worksheet><sheetData/></worksheet>').includes('<hyperlink ref="A1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId5"/>'));
});

test('write: references in extension lists and column breaks follow structural edits', () => {
  const xml =
    '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
    '<protectedRanges><protectedRange sqref="A1:B2" name="bir"/><protectedRange sqref="C5" name="iki"/></protectedRanges>' +
    '<colBreaks count="2" manualBreakCount="1"><brk id="2" max="1048575" man="1"/><brk id="6" min="2" max="9"/></colBreaks>' +
    '<ignoredErrors><ignoredError sqref="C5" numberStoredAsText="1"/></ignoredErrors>' +
    '<extLst><ext uri="{a}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar"><x14:cfvo type="formula"><xm:f>$C$5</xm:f></x14:cfvo></x14:cfRule><xm:sqref>C5:C9</xm:sqref></x14:conditionalFormatting>' +
    '<x14:conditionalFormatting><x14:cfRule type="iconSet"/><xm:sqref>C5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>' +
    '<ext uri="{b}"><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>Diğer!$A$1:$A$3</xm:f></x14:formula1><xm:sqref>C5</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>';
  const run = (op: Op) => {
    const sheet = sheetFromXml(xml);
    sheet.name = 'S';
    const wb = bareWorkbook(sheet);
    applyOp(wb, op);
    const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
    parseXml(out);
    return out;
  };
  let out = run({ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 2 });
  ok(out.includes('<protectedRange sqref="A1:D2" name="bir"/><protectedRange sqref="E5" name="iki"/>'));
  ok(out.includes('<colBreaks count="2" manualBreakCount="1"><brk id="4" max="1048575" man="1"/><brk id="8" min="2" max="9"/></colBreaks>'));
  ok(out.includes('<xm:f>$E$5</xm:f>') && out.includes('<xm:sqref>E5:E9</xm:sqref>') && out.includes('<xm:f>Diğer!$A$1:$A$3</xm:f>'));

  out = run({ type: 'delete', sheet: 0, axis: 'row', at: 4, count: 1 });
  ok(out.includes('<protectedRanges><protectedRange sqref="A1:B2" name="bir"/></protectedRanges>'));
  ok(!out.includes('ignoredError'));
  // Column breaks: the bounds on the other axis shrink with the rows.
  ok(out.includes('<brk id="2" max="1048575" man="1"/><brk id="6" min="2" max="8"/>'));
  ok(out.includes('<xm:sqref>C5:C8</xm:sqref>'));
  ok(!out.includes('iconSet') && !out.includes('{b}'));

  out = run({ type: 'delete', sheet: 0, axis: 'col', at: 1, count: 2 });
  // The break after column B goes to the column before the deleted ones (A).
  ok(out.includes('<colBreaks count="2" manualBreakCount="1"><brk id="1" max="1048575" man="1"/><brk id="4" min="2" max="9"/></colBreaks>'));
  ok(out.endsWith('</colBreaks></worksheet>'));
});

// --- styles -------------------------------------------------------------------------

test('write: formats added by an edit are appended to the styles and nothing else changes', async () => {
  const doc = await load(await buildXlsx(RICH));
  type(doc, 0, 'B2', '12%');
  // A custom number format on top of the bold, centred format.
  const custom = ensureCellFormat(doc.wb.styles, 1, ensureNumberFormat(doc.wb.styles, 'yyyy\\ "yılı" <&>'));
  put(doc, 0, 'A1', { kind: 'number', value: 45000, style: custom });
  eq(doc.wb.styles.modified, true);
  const saved = await roundTrip(doc);
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml', 'xl/styles.xml']);

  const xml = await partText(saved, 'xl/styles.xml');
  const added =
    '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" justifyLastLine="1"/></xf>';
  eq(
    xml,
    STYLES.replace('?>', '?>\n')
      .replace('<numFmts count="1">', '<numFmts count="2">')
      .replace('</numFmts>', '<numFmt numFmtId="165" formatCode="yyyy\\ &quot;yılı&quot; &lt;&amp;>"/></numFmts>')
      .replace('<cellXfs count="3">', '<cellXfs count="5">')
      .replace('</cellXfs>', added + '</cellXfs>')
      .replace('<alignment horizontal="center" justifyLastLine="1"/></xf><xf numFmtId="164"', '<alignment horizontal="center" justifyLastLine="1"/></xf><xf numFmtId="164"'),
  );

  const wb = await reopen(saved);
  deq(wb.styles.cellXfs, doc.wb.styles.cellXfs);
  deq([...wb.styles.numFmts], [...doc.wb.styles.numFmts]);
  deq(getCell(wb.sheets[0], 1, 1), { kind: 'number', value: 0.12, style: 3 });
  eq(wb.styles.cellXfs[getCell(wb.sheets[0], 0, 0)!.style].numFmtId, 165);

  // Saving again starts from the original styles, not from the saved ones.
  eq(await partText(await save(doc), 'xl/styles.xml'), xml);
});

test('write: writeStyles on its own', () => {
  const styles = emptyStyleSheet();
  eq(writeStyles(utf8(STYLES), styles).length, utf8(STYLES).length);

  // A part without number formats or cell formats gets both, in schema order.
  const sparse = `<x:styleSheet xmlns:x="${NS}"><x:fonts count="1"><x:font/></x:fonts><x:dxfs/></x:styleSheet>`;
  const id = ensureNumberFormat(styles, '0.000');
  const format = ensureCellFormat(styles, 0, id);
  eq(format, 1);
  const out = decoder.decode(writeStyles(utf8(sparse), styles));
  eq(
    out,
    `<x:styleSheet xmlns:x="${NS}"><x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="0.000"/></x:numFmts><x:fonts count="1"><x:font/></x:fonts>` +
      '<x:fills count="1"><x:fill><x:patternFill patternType="none"/></x:fill></x:fills><x:borders count="1"><x:border><x:left/><x:right/><x:top/><x:bottom/><x:diagonal/></x:border></x:borders>' +
      '<x:cellXfs count="2"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><x:xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></x:cellXfs><x:dxfs/></x:styleSheet>',
  );
  assert.throws(() => writeStyles(utf8('<worksheet/>'), styles), /root element <worksheet>/);

  // Alignment and protection of a format that has no element to start from.
  const fresh = emptyStyleSheet();
  fresh.cellXfs.push({ numFmtId: 14, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: true, alignment: { horizontal: 'right', vertical: 'top', wrapText: true, shrinkToFit: false, textRotation: 45, indent: 2, readingOrder: 1 }, protection: { locked: false, hidden: true } });
  fresh.modified = true;
  const made = decoder.decode(writeStyles(undefined, fresh));
  ok(
    made.includes(
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1" applyNumberFormat="1" applyAlignment="1" applyProtection="1">' +
        '<alignment horizontal="right" vertical="top" textRotation="45" wrapText="1" indent="2" shrinkToFit="0" readingOrder="1"/><protection locked="0" hidden="1"/></xf></cellXfs>',
    ),
  );
});

test('write: a workbook without styles gets a styles part when a format is needed', async () => {
  const doc = await load(await single('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'));
  type(doc, 0, 'A2', '12%');
  eq(doc.wb.styles.modified, true);
  const saved = await roundTrip(doc);
  const names = ZipArchive.read(saved).names();
  ok(names.includes('xl/styles.xml'));
  ok((await partText(saved, '[Content_Types].xml')).includes('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'));
  ok((await partText(saved, 'xl/_rels/workbook.xml.rels')).includes('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'));
  const wb = await reopen(saved);
  const cell = getCell(wb.sheets[0], 1, 0)!;
  eq(cell.value, 0.12);
  eq(wb.styles.cellXfs[cell.style].numFmtId, 9);
  eq(wb.styles.cellXfs.length, doc.wb.styles.cellXfs.length);

  // A second save does not add the part or the relationship twice.
  type(doc, 0, 'A3', '5');
  const again = await roundTrip(doc);
  deq(ZipArchive.read(again).names(), names);
  eq((await partText(again, 'xl/_rels/workbook.xml.rels')).split('/styles"').length, 2);
  eq((await partText(again, '[Content_Types].xml')).split('/xl/styles.xml').length, 2);
});

// --- workbook part ----------------------------------------------------------------

/** Adds a calculation chain to a package, the way Excel stores one. */
async function withCalcChain(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  const text = async (name: string) => decoder.decode(await zip.read(name));
  zip.set('xl/calcChain.xml', utf8(`${XML}<calcChain xmlns="${NS}"><c r="B4" i="1"/></calcChain>`));
  zip.set(
    'xl/_rels/workbook.xml.rels',
    utf8((await text('xl/_rels/workbook.xml.rels')).replace('</Relationships>', '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>')),
  );
  zip.set(
    '[Content_Types].xml',
    utf8((await text('[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>')),
  );
  return zip.write();
}

test('write: the workbook part asks for recalculation and loses its calculation chain', async () => {
  const doc = await load(await withCalcChain(await buildXlsx(RICH)));
  type(doc, 0, 'B2', '99');
  const saved = await roundTrip(doc);
  const names = ZipArchive.read(saved).names();
  ok(!names.includes('xl/calcChain.xml'));
  const rels = await partText(saved, 'xl/_rels/workbook.xml.rels');
  ok(!rels.includes('calcChain'));
  eq(rels.split('<Relationship ').length - 1, 4);
  ok(!(await partText(saved, '[Content_Types].xml')).includes('calcChain'));

  const before = parseXml(await partText(doc.bytes, 'xl/workbook.xml'));
  const after = parseXml(await partText(saved, 'xl/workbook.xml'));
  const calcPr = after.root.children.find((node): node is XmlElement => node.type === 'element' && node.name === 'calcPr')!;
  deq([...calcPr.attrs], [['calcId', '191029'], ['fullCalcOnLoad', '1']]);
  // Nothing else in the workbook part changed.
  calcPr.attrs.delete('fullCalcOnLoad');
  deq(after, before);

  // Saving again gives the same package.
  deq(ZipArchive.read(await save(doc)).names(), names);
});

test('write: calcPr is created in schema order when the workbook has none', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }],
      workbookHead: '<workbookPr/>',
      workbookTail: '<definedNames><definedName name="Ad">S!$A$1</definedName></definedNames><pivotCaches/><extLst/>',
    }),
  );
  type(doc, 0, 'A1', '1');
  const xml = await partText(await roundTrip(doc), 'xl/workbook.xml');
  ok(xml.includes('</definedNames><calcPr fullCalcOnLoad="1"/><pivotCaches/><extLst/></workbook>'));

  const bare = await load(await single('<sheetData/>'));
  type(bare, 0, 'A1', '1');
  ok((await partText(await roundTrip(bare), 'xl/workbook.xml')).includes('</sheets><calcPr fullCalcOnLoad="1"/></workbook>'));
});

test('write: defined names are matched by name when the lists do not line up, and a missing one is an error', async () => {
  const doc = await load(await buildXlsx(RICH));
  type(doc, 0, 'B2', '1');
  doc.wb.definedNames = [{ ...doc.wb.definedNames[1], value: 'Veri!$A$7' }, { ...doc.wb.definedNames[0], value: 'Veri!$B$2:$B$9 <&>' }];
  const xml = await partText(await save(doc), 'xl/workbook.xml');
  ok(xml.includes('<definedName name="Alan">Veri!$B$2:$B$9 &lt;&amp;&gt;</definedName><definedName name="Yerel" localSheetId="1">Veri!$A$7</definedName>'));

  doc.wb.definedNames = [{ name: 'Yok', value: 'A1', hidden: false }];
  const before = await doc.file.package!.save();
  await assert.rejects(writeXlsx(doc.file.package!, doc.wb), /defined name "Yok" is not in the file/);
  // A failed write leaves the package as it was.
  deq(await doc.file.package!.save(), before);
});

// --- saving more than once -----------------------------------------------------------

test('write: every save starts from the original parts', async () => {
  const doc = await load(await withCalcChain(await buildXlsx(RICH)));
  const original = picture(doc.wb);

  type(doc, 0, 'B2', '99');
  type(doc, 0, 'A20', 'yeni');
  const first = await roundTrip(doc);

  // Undo after saving: the next save has to bring the original content back.
  doc.history.undo();
  doc.history.undo();
  const second = await roundTrip(doc);
  deq(picture(await reopen(second)), original);
  eq(await partText(second, SHEET1), (await partText(doc.bytes, SHEET1)).replace('<row r="2" spans="1:4" x14ac:dyDescent', '<row r="2" x14ac:dyDescent'));

  // Redo and save: the same as the first save.
  doc.history.redo();
  doc.history.redo();
  const third = await roundTrip(doc);
  eq(await partText(third, SHEET1), await partText(first, SHEET1));
  eq(await partText(third, 'xl/workbook.xml'), await partText(first, 'xl/workbook.xml'));

  // More edits on top, structural ones included, saved and undone across saves.
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 });
  type(doc, 0, 'A1', 'üst');
  await roundTrip(doc);
  structural(doc, { type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 });
  await roundTrip(doc);
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  // More rows count as changed by now, so the text may differ from the first save; the content may not.
  const fourth = await roundTrip(doc);
  deq(picture(await reopen(fourth)), picture(await reopen(first)));
  while (doc.history.canUndo) doc.history.undo();
  deq(picture(await reopen(await save(doc))), original);
});

test('write: a saved file can be opened, edited and saved again', async () => {
  const doc = await load(await buildXlsx(RICH));
  type(doc, 0, 'B2', '99');
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 1, count: 1 });
  const next = await load(await roundTrip(doc));
  type(next, 0, 'A2', 'araya');
  type(next, 1, 'B1', '=Veri!B3*2');
  structural(next, { type: 'delete', sheet: 0, axis: 'col', at: 2, count: 1 });
  const saved = await roundTrip(next);
  for (const name of [SHEET1, SHEET2, 'xl/workbook.xml']) parseXml(await partText(saved, name));
});

// --- speed ------------------------------------------------------------------

test('write: 100,000 rows x 10 columns: one changed cell, and full regeneration', () => {
  const rowCount = 100000;
  const shared: { text: string }[] = [];
  for (let i = 0; i < 1000; i++) shared.push({ text: `Liman ${i}` });
  const parts: string[] = [`${XML}<worksheet xmlns="${NS}"><dimension ref="A1:J${rowCount}"/><sheetData>`];
  for (let row = 1; row <= rowCount; row++) {
    let line = `<row r="${row}" spans="1:10" x14ac:dyDescent="0.25">`;
    line += `<c r="A${row}" s="1"><v>${row}</v></c><c r="B${row}" t="s"><v>${row % 1000}</v></c><c r="C${row}"><v>${row * 1.25}</v></c>`;
    line += `<c r="D${row}"><f>A${row}*C${row}</f><v>${row * row * 1.25}</v></c><c r="E${row}" t="b"><v>${row % 2}</v></c>`;
    line += `<c r="F${row}" t="inlineStr"><is><t>satır ${row}</t></is></c><c r="G${row}"><v>45000.${row % 100}</v></c>`;
    line += `<c r="H${row}" t="str"><f>"x"&amp;A${row}</f><v>x${row}</v></c><c r="I${row}" s="3"/><c r="J${row}"><v>-${row}E-2</v></c>`;
    parts.push(line + '</row>');
  }
  parts.push('</sheetData><pageMargins left="1"/></worksheet>');
  const source = parts.join('');
  const bytes = utf8(source);
  const sheet = createSheet({ name: 'Büyük', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(bytes, shared, sheet);
  const wb = bareWorkbook(sheet, shared);

  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 49999, col: 2, cell: { kind: 'number', value: -1, style: 0 } }] });
  let started = Date.now();
  let out = writeWorksheet(bytes, sheet, wb);
  const single = Date.now() - started;
  // Everything before and after the changed row is the source text. (Compared as booleans: a failure must not print 40 MB.)
  const changed = '<row r="50000" x14ac:dyDescent="0.25"><c r="A50000" s="1"><v>50000</v></c><c r="B50000" t="s"><v>0</v></c><c r="C50000"><v>-1</v></c><c r="D50000">';
  const from = source.indexOf('<row r="50000"');
  const to = source.indexOf('<row r="50001"');
  const text = decoder.decode(out);
  const after = text.indexOf('<row r="50001"');
  ok(text.slice(0, from) === source.slice(0, from), 'text before the changed row');
  ok(text.slice(after) === source.slice(to), 'text after the changed row');
  ok(text.slice(from, after).startsWith(changed));

  applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 10, count: 1 });
  started = Date.now();
  out = writeWorksheet(bytes, sheet, wb);
  const full = Date.now() - started;
  console.log(`write: ${(bytes.length / 1e6).toFixed(1)} MB sheet: one changed cell ${single} ms, full regeneration ${full} ms (${(out.length / 1e6).toFixed(1)} MB)`);

  const again = createSheet({ name: 'Büyük', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(out, shared, again);
  eq(again.rows.size, rowCount);
  deq(again.dimension, { top: 0, left: 0, bottom: rowCount, right: 9 });
  deq(getCell(again, 50000, 2), { kind: 'number', value: -1, style: 0 });
  deq(getCell(again, 50000, 3), { kind: 'number', value: 50000 * 50000 * 1.25, style: 0, formula: { text: 'A50001*C50001', type: 'normal' } });
  deq(getCell(again, rowCount, 1), { kind: 'string', value: 'Liman 0', style: 0, si: 0 });
  deq(getCell(again, rowCount, 5), { kind: 'string', value: `satır ${rowCount}`, style: 0 });
  deq(again.rows.get(7)!.attrs, [['x14ac:dyDescent', '0.25']]);
  ok(single < 5000, `one changed cell took ${single} ms`);
  ok(full < 10000, `full regeneration took ${full} ms`);
});

// --- formatting --------------------------------------------------------------------

const range = (ref: string) => parseRange(ref)!;
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

/** The records of a style table, for comparing what a save kept. */
function tableRecords(xml: string, table: string): XmlElement[] {
  const root = parseXml(xml).root;
  const container = root.children.find((node): node is XmlElement => node.type === 'element' && node.name === table);
  return container ? container.children.filter((node): node is XmlElement => node.type === 'element') : [];
}

test('write: formatting commands are saved: fonts, fills, borders, alignment, number formats', async () => {
  const doc = await load(await buildXlsx(RICH));
  const wb = doc.wb;
  apply(doc, formatRange(wb, 0, range('B2:C3'), { bold: true, italic: true, underline: 'double', fontName: 'Arial', fontSize: 14, color: { rgb: 'FFFF0000' } }));
  apply(doc, formatRange(wb, 0, range('B2:C3'), { fill: { rgb: 'FFFFFF00' }, border: { preset: 'all', edge: { style: 'medium', color: { rgb: 'FF0000FF' } } } }));
  apply(doc, formatRange(wb, 0, range('B2:B2'), { horizontal: 'right', vertical: 'top', wrapText: true, numFmt: '0.00%' }));
  apply(doc, formatRange(wb, 0, range('A1:A1'), { strike: true, numFmt: 'yyyy-mm-dd' }));
  apply(doc, formatRange(wb, 0, range('E12:F13'), { border: { preset: 'outline' }, fill: { theme: 4, tint: 0.4 } }));
  // Whole columns and whole rows are formatted as such.
  apply(doc, formatRange(wb, 0, { top: 0, bottom: MAX_ROWS - 1, left: 5, right: 6 }, { italic: true }));
  apply(doc, formatRange(wb, 0, { top: 14, bottom: 15, left: 0, right: MAX_COLS - 1 }, { fill: { rgb: 'FF00FF00' } }));
  eq(wb.styles.modified, true);

  const saved = await roundTrip(doc);
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml', 'xl/styles.xml']);
  const again = await reopen(saved);

  // Every table reads back as the model has it, and so does every position.
  deq(again.styles.fonts, wb.styles.fonts);
  deq(again.styles.fills, wb.styles.fills);
  deq(again.styles.borders, wb.styles.borders);
  deq(again.styles.cellXfs, wb.styles.cellXfs);
  deq([...again.styles.numFmts], [...wb.styles.numFmts]);
  for (const ref of ['A1', 'B1', 'B2', 'C2', 'B3', 'C3', 'A3', 'E12', 'F12', 'E13', 'F13', 'F1', 'G30', 'A15', 'Z16', 'F15', 'H20']) {
    const { row, col } = at(ref);
    deq(resolveCellStyle(again.styles, styleAt(again.sheets[0], row, col)), resolveCellStyle(wb.styles, styleAt(wb.sheets[0], row, col)), ref);
  }
  const b2 = resolveCellStyle(again.styles, getCell(again.sheets[0], 1, 1)!.style);
  deq(b2.font, { name: 'Arial', size: 14, bold: true, italic: true, strike: false, underline: 'double', color: { rgb: 'FFFF0000' } });
  deq(b2.fill, { type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFFFF00' }, bg: { indexed: 64 } });
  deq(b2.border.left, { style: 'medium', color: { rgb: 'FF0000FF' } });
  deq(b2.alignment, { horizontal: 'right', vertical: 'top', wrapText: true });
  eq(b2.numFmt, '0.00%');
  eq(getCell(again.sheets[0], 1, 1)!.value, 12.5);

  // The style part only grew: every original record is still there, as it was, in its place.
  const before = await partText(doc.bytes, 'xl/styles.xml');
  const after = await partText(saved, 'xl/styles.xml');
  for (const table of ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'extLst']) {
    const original = tableRecords(before, table);
    const now = tableRecords(after, table);
    deq(now.slice(0, original.length), original, table);
    eq(tableRecords(after, table).length, { numFmts: wb.styles.numFmts.size, fonts: wb.styles.fonts.length, fills: wb.styles.fills.length, borders: wb.styles.borders.length, cellXfs: wb.styles.cellXfs.length }[table] ?? original.length, table);
  }
  ok(after.includes(`<fonts count="${wb.styles.fonts.length}">`) && after.includes(`<fills count="${wb.styles.fills.length}">`));
  ok(after.includes(`<borders count="${wb.styles.borders.length}">`) && after.includes(`<cellXfs count="${wb.styles.cellXfs.length}">`));
  ok(after.includes('<font><b/><i/><u val="double"/><sz val="14"/><color rgb="FFFF0000"/><name val="Arial"/></font>'));
  ok(after.includes('<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>'));
  ok(after.includes('<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/><bgColor indexed="64"/></patternFill></fill>'));
  const edge = '<left style="medium"><color rgb="FF0000FF"/></left><right style="medium"><color rgb="FF0000FF"/></right><top style="medium"><color rgb="FF0000FF"/></top><bottom style="medium"><color rgb="FF0000FF"/></bottom><diagonal/>';
  ok(after.includes(`<border>${edge}</border>`));
  ok(after.includes('<border><left style="thin"/><right/><top style="thin"/><bottom/><diagonal/></border>'));

  // A1 had the bold, centred format: its copy keeps what the model does not hold and flags what changed.
  const a1 = tableRecords(after, 'cellXfs')[getCell(again.sheets[0], 0, 0)!.style];
  eq(a1.attrs.get('applyAlignment'), '1');
  eq(a1.attrs.get('applyFont'), '1');
  eq(a1.attrs.get('applyNumberFormat'), '1');
  deq([...(a1.children[0] as XmlElement).attrs], [['horizontal', 'center'], ['justifyLastLine', '1']]);
  // B2 came from the default format: every table it now points into is flagged.
  const b2xf = tableRecords(after, 'cellXfs')[getCell(again.sheets[0], 1, 1)!.style];
  for (const flag of ['applyFont', 'applyFill', 'applyBorder', 'applyAlignment', 'applyNumberFormat']) eq(b2xf.attrs.get(flag), '1', flag);
  deq([...(b2xf.children[0] as XmlElement).attrs], [['horizontal', 'right'], ['vertical', 'top'], ['wrapText', '1']]);

  // Whole columns and rows in the sheet part.
  const xml = await partText(saved, SHEET1);
  const italic = wb.sheets[0].columns.find((c) => c.min === 5)!.style;
  ok(xml.includes(`<col min="6" max="7" style="${italic}"/>`));
  const green = wb.sheets[0].rows.get(14)!.style;
  ok(xml.includes(`<row r="15" s="${green}" customFormat="1"/><row r="16" s="${green}" customFormat="1"/>`));
});

test('write: an alignment that is removed or changed on a copied format', () => {
  const styles = emptyStyleSheet();
  styles.fonts.push({ ...styles.fonts[0], bold: true });
  styles.cellXfs.push({ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, alignment: { horizontal: 'center' }, protection: { locked: false } });
  const original =
    `<styleSheet xmlns="${NS}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>` +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>' +
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1" applyProtection="1" pivotButton="1"><alignment horizontal="center" justifyLastLine="1"/><protection locked="0"/><extLst><ext uri="x"/></extLst></xf></cellXfs></styleSheet>';
  // No alignment any more; alignment changed; protection changed; alignment added to the plain format.
  styles.cellXfs.push({ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, protection: { locked: false } });
  styles.cellXfs.push({ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, alignment: { horizontal: 'left', wrapText: true }, protection: { locked: false } });
  styles.cellXfs.push({ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, alignment: { horizontal: 'center' }, protection: { locked: true, hidden: true } });
  styles.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, alignment: { vertical: 'top' } });
  // A format with another parent style has nothing to start from.
  styles.cellXfs.push({ numFmtId: 2, fontId: 1, fillId: 0, borderId: 0, xfId: 3, quotePrefix: false });
  styles.modified = true;
  const out = decoder.decode(writeStyles(utf8(original), styles));
  ok(
    out.endsWith(
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyProtection="1" pivotButton="1"><protection locked="0"/><extLst><ext uri="x"/></extLst></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1" applyProtection="1" pivotButton="1"><alignment horizontal="left" justifyLastLine="1" wrapText="1"/><protection locked="0"/><extLst><ext uri="x"/></extLst></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1" applyProtection="1" pivotButton="1"><alignment horizontal="center" justifyLastLine="1"/><protection locked="1" hidden="1"/><extLst><ext uri="x"/></extLst></xf>' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>' +
        '<xf numFmtId="2" fontId="1" fillId="0" borderId="0" xfId="3" applyNumberFormat="1" applyFont="1"/></cellXfs></styleSheet>',
    ),
    out,
  );
  ok(out.includes('<cellXfs count="7">'));
});

test('write: fonts, fills and borders of every shape', () => {
  const styles = emptyStyleSheet();
  styles.fonts.push({ bold: true, italic: true, strike: true, underline: 'single', vertAlign: 'superscript', size: 9.5, color: { theme: 1, tint: -0.5 }, name: 'A & "B"', family: 2, charset: 162, scheme: 'minor' });
  styles.fonts.push({ bold: false, italic: false, strike: false, underline: 'doubleAccounting', color: { indexed: 8 } });
  styles.fonts.push({ bold: false, italic: false, strike: false, underline: 'none', color: { auto: true } });
  styles.fills.push({ type: 'pattern', pattern: 'gray125' });
  styles.fills.push({ type: 'pattern', pattern: 'darkGrid', fg: { rgb: 'FF112233' }, bg: { theme: 2 } });
  const gradient = parseXml('<gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color rgb="FF00FF00"/></stop></gradientFill>').root;
  styles.fills.push({ type: 'gradient', raw: gradient });
  styles.borders.push({ diagonalUp: true, diagonalDown: true, left: { style: 'thin' }, right: { style: 'dashed', color: { auto: true } }, top: { style: 'double', color: { indexed: 64 } }, bottom: { style: 'hair', color: { theme: 3, tint: 0.25 } }, diagonal: { style: 'medium', color: { rgb: 'FFAA0000' } } });
  styles.modified = true;
  const out = decoder.decode(writeStyles(undefined, styles));
  ok(out.includes('<font><b/><i/><strike/><u/><vertAlign val="superscript"/><sz val="9.5"/><color theme="1" tint="-0.5"/><name val="A &amp; &quot;B&quot;"/><family val="2"/><charset val="162"/><scheme val="minor"/></font>'));
  ok(out.includes('<font><u val="doubleAccounting"/><color indexed="8"/></font><font><color auto="1"/></font></fonts>'));
  ok(out.includes('<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGrid"><fgColor rgb="FF112233"/><bgColor theme="2"/></patternFill></fill>'));
  ok(out.includes('<fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color rgb="FF00FF00"/></stop></gradientFill></fill></fills>'));
  ok(
    out.includes(
      '<border diagonalUp="1" diagonalDown="1"><left style="thin"/><right style="dashed"><color auto="1"/></right><top style="double"><color indexed="64"/></top>' +
        '<bottom style="hair"><color theme="3" tint="0.25"/></bottom><diagonal style="medium"><color rgb="FFAA0000"/></diagonal></border></borders>',
    ),
  );
  // Tables in schema order.
  deq([...out.matchAll(/<(fonts|fills|borders|cellStyleXfs|cellXfs|cellStyles) /g)].map((m) => m[1]), ['fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles']);
});

test('write: formatting a workbook that has no styles part', async () => {
  const doc = await load(await single('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'));
  // A second fill first, so that the solid fill does not land on the index reserved for "gray125".
  doc.wb.styles.fills.push({ type: 'pattern', pattern: 'gray125' });
  apply(doc, formatRange(doc.wb, 0, range('A1:B2'), { bold: true, fill: { rgb: 'FFFFCC00' }, border: { preset: 'outline' }, horizontal: 'center', numFmt: '0.0' }));
  const saved = await roundTrip(doc);
  const again = await reopen(saved);
  deq(again.styles.fonts, doc.wb.styles.fonts);
  deq(again.styles.fills, doc.wb.styles.fills);
  deq(again.styles.borders, doc.wb.styles.borders);
  deq(again.styles.cellXfs, doc.wb.styles.cellXfs);
  for (const ref of ['A1', 'B1', 'A2', 'B2', 'C3']) {
    const { row, col } = at(ref);
    deq(resolveCellStyle(again.styles, styleAt(again.sheets[0], row, col)), resolveCellStyle(doc.wb.styles, styleAt(doc.wb.sheets[0], row, col)), ref);
  }
  const xml = await partText(saved, 'xl/styles.xml');
  ok(xml.includes('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'));
});

test('write: column widths, row heights, hidden rows and columns, merges', async () => {
  const doc = await load(await buildXlsx(RICH));
  const wb = doc.wb;
  apply(doc, setColumnWidth(wb, 0, 1, 1, 33.5));
  apply(doc, setColumnsHidden(wb, 0, 6, 7, true));
  apply(doc, setRowHeight(0, 1, 1, 40));
  apply(doc, setRowHeight(0, 19, 20, 12.5));
  apply(doc, setRowsHidden(0, 3, 3, true));
  apply(doc, setRowsHidden(0, 29, 29, true));
  apply(doc, unmergeRange(wb, 0, range('A8:A8')));
  apply(doc, mergeRange(wb, 0, range('E2:F4')));
  const saved = await roundTrip(doc);
  const xml = await partText(saved, SHEET1);
  const before = await partText(doc.bytes, SHEET1);

  ok(xml.includes('<cols><col min="1" max="1" width="18.5" customWidth="1" bestFit="1"/><col min="2" max="2" width="33.5" customWidth="1"/><col min="3" max="4" width="9" hidden="1"/><col min="7" max="8" hidden="1"/></cols>'));
  ok(xml.includes('<row r="2" ht="40" customHeight="1" x14ac:dyDescent="0.25"><c r="A2" t="s"><v>1</v></c>'));
  ok(xml.includes('<row r="4" hidden="1"><c r="A4" t="inlineStr"><is><t>Toplam</t></is></c>'));
  // Rows the file did not have, with nothing but a height or a flag.
  ok(xml.includes('<row r="20" ht="12.5" customHeight="1"/><row r="21" ht="12.5" customHeight="1"/><row r="30" hidden="1"/></sheetData>'));
  ok(xml.includes('<autoFilter ref="A1:D4"/><mergeCells count="1"><mergeCell ref="E2:F4"/></mergeCells><conditionalFormatting'));
  // Rows that were not touched and the elements the model does not hold are the original text.
  ok(xml.includes(ROW1) && xml.includes(ROW3) && xml.includes(TAIL_ELEMENTS.join('') + '</worksheet>'));
  ok(xml.startsWith(before.slice(0, before.indexOf('<cols>'))));

  const again = await reopen(saved);
  deq(again.sheets[0].columns, wb.sheets[0].columns);
  deq(again.sheets[0].merges, [range('E2:F4')]);
  deq([again.sheets[0].rows.get(1)!.height, again.sheets[0].rows.get(1)!.customHeight], [40, true]);
  eq(again.sheets[0].rows.get(3)!.hidden, true);
  eq(again.sheets[0].rows.get(29)!.hidden, true);
  eq(again.sheets[0].rows.get(20)!.height, 12.5);

  // Removing the last merge removes the element; undoing everything brings the original ranges back.
  apply(doc, unmergeRange(wb, 0, range('E2:E2')));
  ok(!(await partText(await roundTrip(doc), SHEET1)).includes('mergeCells'));
  while (doc.history.canUndo) doc.history.undo();
  const restored = await reopen(await save(doc));
  deq(restored.sheets[0].merges, [range('A8:C9')]);
  deq(restored.sheets[0].columns, (await reopen(doc.bytes)).sheets[0].columns);
});

test('write: columns and merges are added in schema order to a sheet that has none', async () => {
  const doc = await load(await single('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><sheetProtection sheet="1"/><pageMargins left="1"/>'));
  apply(doc, setColumnWidth(doc.wb, 0, 2, 3, 20));
  apply(doc, mergeRange(doc.wb, 0, range('B2:C3')));
  const xml = await partText(await roundTrip(doc), SHEET1);
  ok(
    xml.endsWith(
      '</sheetViews><cols><col min="3" max="4" width="20" customWidth="1"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
        '<sheetProtection sheet="1"/><mergeCells count="1"><mergeCell ref="B2:C3"/></mergeCells><pageMargins left="1"/></worksheet>',
    ),
  );
});

// --- pictures and comments ------------------------------------------------------------

const NS_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const marker = (tag: string, col: number, row: number, colOff = 9525, rowOff = 19050) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`;
const shape = (name: string) => `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="${name}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/>`;
const twoCell = (name: string, from: [number, number], to: [number, number], attrs = '') => `<xdr:twoCellAnchor${attrs}>${marker('from', from[0], from[1])}${marker('to', to[0], to[1])}${shape(name)}</xdr:twoCellAnchor>`;

const DRAWING =
  `${XML}<xdr:wsDr xmlns:xdr="${NS_XDR}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
  twoCell('önde', [1, 1], [3, 4]) +
  twoCell('arkada', [1, 8], [3, 10]) +
  twoCell('üstünde', [0, 4], [2, 7]) +
  `<xdr:oneCellAnchor>${marker('from', 5, 9)}<xdr:ext cx="100" cy="200"/>${shape('tek')}</xdr:oneCellAnchor>` +
  `<xdr:absoluteAnchor><xdr:pos x="0" y="0"/><xdr:ext cx="100" cy="200"/>${shape('mutlak')}</xdr:absoluteAnchor>` +
  twoCell('sabit', [1, 9], [3, 11], ' editAs="absolute"') +
  twoCell('taşınır', [1, 9], [3, 11], ' editAs="oneCell"') +
  `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a14">${twoCell('yeni', [4, 9], [5, 10])}</mc:Choice><mc:Fallback>${twoCell('yedek', [4, 9], [5, 10])}</mc:Fallback></mc:AlternateContent>` +
  '</xdr:wsDr>';

const COMMENTS =
  `${XML}<comments xmlns="${NS}"><authors><author>Yazar</author></authors><commentList>` +
  '<comment ref="A1" authorId="0"><text><r><t>ilk</t></r></text></comment>' +
  '<comment ref="B6" authorId="0"><text><t>ortadaki</t></text></comment>' +
  '<comment ref="C9" authorId="0"><text><t>sondaki</t></text></comment>' +
  '</commentList></comments>';

const note = (id: number, row: number, col: number, anchor: string, extra = '') =>
  `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;visibility:hidden' fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>` +
  `<v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div>${extra}</v:textbox>` +
  `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`;

const BUTTON = '<v:shape id="_x0000_s1028" type="#_x0000_t201"><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 5, 0, 2, 0, 6, 0</x:Anchor><x:Row>5</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>';

const vml = (extra = '', button = '') =>
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
  note(1025, 0, 0, '\n    1, 15, 0, 2, 3, 31, 4, 1', extra) +
  note(1026, 5, 1, '2, 15, 4, 10, 4, 31, 8, 4') +
  note(1027, 8, 2, ' 3, 15, 7, 10, 5, 31, 11, 4 ') +
  button +
  '</xml>';

const DRAWING_PART = 'xl/drawings/drawing1.xml';
const COMMENTS_PART = 'xl/comments1.xml';
const VML_PART = 'xl/drawings/vmlDrawing1.vml';
const SHEET1_RELS = 'xl/worksheets/_rels/sheet1.xml.rels';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A sheet with a drawing, comments, their boxes and two links, laid out the way Excel does. */
async function anchoredWorkbook(vmlText = vml(), secondSheetShares = false, drawing = DRAWING): Promise<Uint8Array> {
  const sheet = (extra: string) =>
    worksheetXml(
      '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
        '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="6"><c r="B6"><v>6</v></c></row><row r="9"><c r="C9"><v>9</v></c></row></sheetData>' +
        extra,
    );
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Resimli', xml: sheet('<hyperlinks><hyperlink ref="B6" r:id="rId9"/><hyperlink ref="C9" r:id="rId8"/></hyperlinks><drawing r:id="rId1"/><legacyDrawing r:id="rId2"/>') },
      { name: 'Öteki', xml: sheet(secondSheetShares ? '<drawing r:id="rId1"/>' : '') },
    ],
  });
  const zip = ZipArchive.read(bytes);
  zip.set(DRAWING_PART, utf8(drawing));
  zip.set(COMMENTS_PART, utf8(COMMENTS));
  zip.set(VML_PART, utf8(vmlText));
  zip.set(
    SHEET1_RELS,
    utf8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/>` +
        `<Relationship Id="rId2" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>` +
        `<Relationship Id="rId3" Type="${REL}/comments" Target="../comments1.xml"/>` +
        `<Relationship Id="rId8" Type="${REL}/hyperlink" Target="https://example.com/c9" TargetMode="External"/>` +
        `<Relationship Id="rId9" Type="${REL}/hyperlink" Target="https://example.com/b6" TargetMode="External"/>` +
        '</Relationships>',
    ),
  );
  if (secondSheetShares) {
    zip.set('xl/worksheets/_rels/sheet2.xml.rels', utf8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`));
  }
  zip.set(
    '[Content_Types].xml',
    utf8(
      decoder
        .decode(await zip.read('[Content_Types].xml'))
        .replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>'),
    ),
  );
  return zip.write();
}

/** Name of each object of a drawing -> [from col, from colOff, from row, from rowOff, to ...]. */
function anchorsOf(xml: string): Record<string, number[]> {
  const out: Record<string, number[]> = {};
  const visit = (el: XmlElement) => {
    for (const node of el.children) {
      if (node.type !== 'element') continue;
      if (/:(twoCellAnchor|oneCellAnchor)$/.test(node.name)) {
        const numbers: number[] = [];
        let name = '';
        for (const part of node.children) {
          if (part.type !== 'element') continue;
          if (/:(from|to)$/.test(part.name)) for (const value of part.children) if (value.type === 'element') numbers.push(Number((value.children[0] as { value: string }).value));
          const match = /name="([^"]*)"/.exec(JSON.stringify([...((part.children.find((n) => n.type === 'element') as XmlElement | undefined)?.children ?? [])].map((n) => (n.type === 'element' ? [...n.attrs].map(([k, v]) => `${k}="${v}"`).join(' ') : ''))).replace(/\\"/g, '"'));
          if (match && /:sp$/.test(part.name)) name = match[1];
        }
        out[name] = numbers;
      } else visit(node);
    }
  };
  visit(parseXml(xml).root);
  return out;
}

/** Row, column and anchor of each comment box of a legacy drawing, by shape id. */
function notesOf(text: string): Record<string, [number, number, string]> {
  const out: Record<string, [number, number, string]> = {};
  for (const block of text.split('</v:shape>')) {
    if (!block.includes('ObjectType="Note"')) continue;
    const m = /<v:shape id="_x0000_s(\d+)"[\s\S]*?<x:Anchor>([^<]*)<\/x:Anchor>[\s\S]*?<x:Row>(\d+)<\/x:Row><x:Column>(\d+)<\/x:Column>/.exec(block);
    if (m) out[m[1]] = [Number(m[3]), Number(m[4]), m[2].replace(/\s/g, '')];
  }
  return out;
}

const commentRefs = (xml: string) => [...xml.matchAll(/<comment ref="([^"]*)"/g)].map((m) => m[1]);

test('write: pictures and comments stay byte for byte while no rows or columns moved', async () => {
  const doc = await load(await anchoredWorkbook());
  deq(doc.wb.sheets[0].anchored, ['comments', 'drawing', 'vmlDrawing']);
  type(doc, 0, 'A1', '5');
  apply(doc, setColumnWidth(doc.wb, 0, 0, 0, 30));
  const saved = await roundTrip(doc);
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml']);
});

test('write: inserting rows moves pictures, comments and their boxes', async () => {
  // The legacy drawing also holds a button. The reader makes such a sheet refuse structural edits;
  // the writer must still leave the button alone if it ever gets one, so the refusal is lifted here.
  const doc = await load(await anchoredWorkbook(vml('', BUTTON)));
  deq(doc.wb.sheets[0].anchored, ['comments', 'drawing', 'formControl', 'vmlDrawing']);
  doc.wb.sheets[0].anchored = ['comments', 'drawing', 'vmlDrawing'];
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 5, count: 2 });
  const saved = await roundTrip(doc);
  assertOtherPartsUntouched(doc.bytes, saved, [SHEET1, 'xl/workbook.xml', DRAWING_PART, COMMENTS_PART, VML_PART]);

  deq(anchorsOf(await partText(saved, DRAWING_PART)), {
    önde: [1, 9525, 1, 19050, 3, 9525, 4, 19050],
    arkada: [1, 9525, 10, 19050, 3, 9525, 12, 19050],
    // Rows went in under its top edge: it grows.
    üstünde: [0, 9525, 4, 19050, 2, 9525, 9, 19050],
    tek: [5, 9525, 11, 19050],
    sabit: [1, 9525, 9, 19050, 3, 9525, 11, 19050],
    taşınır: [1, 9525, 11, 19050, 3, 9525, 13, 19050],
    yeni: [4, 9525, 11, 19050, 5, 9525, 12, 19050],
    yedek: [4, 9525, 11, 19050, 5, 9525, 12, 19050],
  });
  // Everything in the drawing but the numbers is as it was.
  const strip = (xml: string) => xml.replace(/<xdr:(row|col)>\d+<\/xdr:\1>/g, '').replace('?>\n', '?>');
  eq(strip(await partText(saved, DRAWING_PART)), strip(DRAWING));

  deq(commentRefs(await partText(saved, COMMENTS_PART)), ['A1', 'B8', 'C11']);
  eq(await partText(saved, COMMENTS_PART), COMMENTS.replace('?>', '?>\n').replace('"B6"', '"B8"').replace('"C9"', '"C11"'));
  const text = await partText(saved, VML_PART);
  parseXml(text);
  deq(notesOf(text), { 1025: [0, 0, '1,15,0,2,3,31,4,1'], 1026: [7, 1, '2,15,4,10,4,31,10,4'], 1027: [10, 2, '3,15,9,10,5,31,13,4'] });
  // An anchor that did not move keeps its own spacing; a button is not a comment box.
  ok(text.includes('<x:Anchor>\n    1, 15, 0, 2, 3, 31, 4, 1</x:Anchor>'));
  ok(text.includes('<x:ClientData ObjectType="Button"><x:Anchor>1, 0, 5, 0, 2, 0, 6, 0</x:Anchor><x:Row>5</x:Row><x:Column>1</x:Column></x:ClientData>'));

  // The frozen panes of the model and of the file agree.
  const again = await reopen(saved);
  deq([again.sheets[0].view.frozenRows, again.sheets[0].view.frozenCols], [doc.wb.sheets[0].view.frozenRows, doc.wb.sheets[0].view.frozenCols]);
  deq([again.sheets[0].view.frozenRows, again.sheets[0].view.frozenCols], [2, 1]);
  deq(again.sheets[0].anchored, ['comments', 'drawing', 'formControl', 'vmlDrawing']);

  // Undoing the insert and saving again puts the parts back as the file had them.
  doc.history.undo();
  const restored = await roundTrip(doc);
  eq(await partText(restored, DRAWING_PART), DRAWING);
  eq(await partText(restored, COMMENTS_PART), COMMENTS);
  eq(await partText(restored, VML_PART), vml('', BUTTON));
  eq(await partText(restored, SHEET1_RELS), await partText(doc.bytes, SHEET1_RELS));
});

test('write: deleting rows moves, shrinks or removes what was on them', async () => {
  const doc = await load(await anchoredWorkbook());
  // Rows 6 to 9 go: the comments on B6 and C9 with them, and the link on each.
  structural(doc, { type: 'delete', sheet: 0, axis: 'row', at: 5, count: 4 });
  const saved = await roundTrip(doc);
  deq(anchorsOf(await partText(saved, DRAWING_PART)), {
    önde: [1, 9525, 1, 19050, 3, 9525, 4, 19050],
    // Its top edge was on a deleted row: it starts where the rows were.
    arkada: [1, 9525, 5, 0, 3, 9525, 6, 19050],
    üstünde: [0, 9525, 4, 19050, 2, 9525, 5, 0],
    tek: [5, 9525, 5, 19050],
    sabit: [1, 9525, 9, 19050, 3, 9525, 11, 19050],
    taşınır: [1, 9525, 5, 19050, 3, 9525, 7, 19050],
    yeni: [4, 9525, 5, 19050, 5, 9525, 6, 19050],
    yedek: [4, 9525, 5, 19050, 5, 9525, 6, 19050],
  });
  deq(commentRefs(await partText(saved, COMMENTS_PART)), ['A1']);
  ok((await partText(saved, COMMENTS_PART)).includes('<authors><author>Yazar</author></authors>'));
  const text = await partText(saved, VML_PART);
  parseXml(text);
  deq(notesOf(text), { 1025: [0, 0, '1,15,0,2,3,31,4,1'] });
  ok(text.includes('<v:shapetype id="_x0000_t202"'));

  // The relationships of the links that are gone are gone too; the others stay.
  const rels = await partText(saved, SHEET1_RELS);
  ok(!rels.includes('hyperlink'));
  deq([...rels.matchAll(/Id="([^"]*)"/g)].map((m) => m[1]), ['rId1', 'rId2', 'rId3']);
  ok(!(await partText(saved, SHEET1)).includes('<hyperlinks'));
  deq((await reopen(saved)).sheets[0].anchored, ['comments', 'drawing', 'vmlDrawing']);

  // A picture whose rows are all deleted is kept, with no height, where the rows were.
  const flat = await load(await anchoredWorkbook());
  structural(flat, { type: 'delete', sheet: 0, axis: 'row', at: 8, count: 3 });
  const anchors = anchorsOf(await partText(await roundTrip(flat), DRAWING_PART));
  deq(anchors.arkada, [1, 9525, 8, 0, 3, 9525, 8, 0]);
  deq(anchors.tek, [5, 9525, 8, 0]);
  // One link is left, and so is its relationship.
  const kept = await partText(await save(flat), SHEET1_RELS);
  ok(kept.includes('Id="rId9"') && !kept.includes('Id="rId8"'));
});

test('write: inserting and deleting columns moves pictures, comments and their boxes', async () => {
  const doc = await load(await anchoredWorkbook());
  structural(doc, { type: 'insert', sheet: 0, axis: 'col', at: 1, count: 3 });
  let saved = await roundTrip(doc);
  let anchors = anchorsOf(await partText(saved, DRAWING_PART));
  deq(anchors['önde'], [4, 9525, 1, 19050, 6, 9525, 4, 19050]);
  deq(anchors['üstünde'], [0, 9525, 4, 19050, 5, 9525, 7, 19050]);
  deq(anchors.tek, [8, 9525, 9, 19050]);
  deq(anchors.sabit, [1, 9525, 9, 19050, 3, 9525, 11, 19050]);
  deq(commentRefs(await partText(saved, COMMENTS_PART)), ['A1', 'E6', 'F9']);
  deq(notesOf(await partText(saved, VML_PART)), { 1025: [0, 0, '4,15,0,2,6,31,4,1'], 1026: [5, 4, '5,15,4,10,7,31,8,4'], 1027: [8, 5, '6,15,7,10,8,31,11,4'] });

  // Then columns B to E go again, with the comment that moved to E6.
  structural(doc, { type: 'delete', sheet: 0, axis: 'col', at: 1, count: 4 });
  saved = await roundTrip(doc);
  anchors = anchorsOf(await partText(saved, DRAWING_PART));
  deq(anchors['önde'], [1, 0, 1, 19050, 2, 9525, 4, 19050]);
  deq(anchors['üstünde'], [0, 9525, 4, 19050, 1, 9525, 7, 19050]);
  deq(commentRefs(await partText(saved, COMMENTS_PART)), ['A1', 'B9']);
  deq(notesOf(await partText(saved, VML_PART)), { 1025: [0, 0, '1,0,0,2,2,31,4,1'], 1027: [8, 1, '2,15,7,10,4,31,11,4'] });
});

test('write: a legacy drawing that is not well-formed XML is adjusted on its text', async () => {
  // An unclosed <br> inside a text box: fine for the application, not for an XML parser.
  const text = vml('<br>');
  assert.throws(() => parseXml(text));
  const doc = await load(await anchoredWorkbook(text));
  structural(doc, { type: 'delete', sheet: 0, axis: 'row', at: 5, count: 1 });
  const saved = await roundTrip(doc);
  const out = await partText(saved, VML_PART);
  deq(notesOf(out), { 1025: [0, 0, '1,15,0,2,3,31,4,1'], 1027: [7, 2, '3,15,6,10,5,31,10,4'] });
  // Everything but the removed box and the three changed values is the original text.
  eq(out, text.replace(note(1026, 5, 1, '2, 15, 4, 10, 4, 31, 8, 4'), '').replace(note(1027, 8, 2, ' 3, 15, 7, 10, 5, 31, 11, 4 '), note(1027, 7, 2, ' 3, 15, 6, 10, 5, 31, 10, 4 ')));
  ok(out.includes('<v:shapetype id="_x0000_t202"') && out.includes('<br>'));
  deq(commentRefs(await partText(saved, COMMENTS_PART)), ['A1', 'C8']);
});

test('write: a drawing shared by two sheets is refused rather than moved for one of them', async () => {
  const doc = await load(await anchoredWorkbook(vml(), true));
  structural(doc, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  const before = await doc.file.package!.save();
  await assert.rejects(writeXlsx(doc.file.package!, doc.wb), /drawing1\.xml is shared with sheet "Öteki"/);
  deq(await doc.file.package!.save(), before);
});

// --- raw content that refers to cells ----------------------------------------------------

/** What a structural edit makes of raw elements placed after the rows of a sheet called "Veri". */
function adjusted(inner: string, op: Op, head = ''): string {
  const xml = `<worksheet>${head}<sheetData><row r="5"><c r="A5"><v>1</v></c></row></sheetData>${inner}</worksheet>`;
  const sheet = sheetFromXml(xml);
  sheet.name = 'Veri';
  const wb = bareWorkbook(sheet);
  applyOp(wb, op);
  const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
  parseXml(out);
  return out.slice(out.indexOf(head === '' ? 'sheetData' : '<worksheet>') + (head === '' ? 0 : 11)).replace(/^sheetData\/?>(.*<\/sheetData>)?/, '').replace('</worksheet>', '');
}

const rows = (kind: 'insert' | 'delete', at: number, count: number): Op => ({ type: kind, sheet: 0, axis: 'row', at, count });
const cols = (kind: 'insert' | 'delete', at: number, count: number): Op => ({ type: kind, sheet: 0, axis: 'col', at, count });
const rule = (sqref: string, formula: string, extra = '') => `<conditionalFormatting sqref="${sqref}"><cfRule type="expression" priority="1"><formula>${formula}</formula></cfRule>${extra}</conditionalFormatting>`;
const validation = (sqref: string, formula: string) => `<dataValidations count="1"><dataValidation type="custom" sqref="${sqref}"><formula1>${formula}</formula1></dataValidation></dataValidations>`;

test('write: formulas of conditional formats and validations follow the corner of their range', () => {
  // Relative: "the cell this rule is on, compared with its neighbour". Deleting the first row of the
  // range leaves the rule saying the same thing; it must not become #REF!.
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), rows('delete', 4, 1)), rule('A5:A9', 'A5&gt;B5'));
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), rows('insert', 2, 2)), rule('A7:A12', 'A7&gt;B7'));
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), rows('insert', 6, 1)), rule('A5:A11', 'A5&gt;B5'));
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), cols('insert', 0, 1)), rule('B5:B10', 'B5&gt;C5'));
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), cols('insert', 1, 1)), rule('A5:A10', 'A5&gt;B5'));
  eq(adjusted(rule('B5:B10', 'B5&gt;C5'), cols('delete', 0, 1)), rule('A5:A10', 'A5&gt;B5'));
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), rows('delete', 4, 6)), '');
  eq(adjusted(rule('A5:A10', 'A5&gt;B5'), cols('delete', 0, 1)), '');

  // Absolute: these name particular cells and follow them like any formula.
  eq(adjusted(rule('C5:C10', '$A$5&gt;$B$7'), rows('delete', 4, 1)), rule('C5:C9', '#REF!&gt;$B$6'));
  eq(adjusted(rule('C5:C10', '$A$5&gt;$B$7'), rows('insert', 2, 2)), rule('C7:C12', '$A$7&gt;$B$9'));
  eq(adjusted(rule('C5:C10', '$A$5&gt;$B$7'), cols('insert', 0, 1)), rule('D5:D10', '$B$5&gt;$C$7'));
  eq(adjusted(rule('C5:C10', '$A$5&gt;$B$7'), cols('delete', 0, 1)), rule('B5:B10', '#REF!&gt;$A$7'));

  // Mixed, a second range, and a threshold given as a formula.
  const bar = (val: string) => `<cfRule type="dataBar" priority="4"><dataBar><cfvo type="formula" val="${val}"/><cfvo type="max"/><cfvo type="num" val="5"/><color rgb="FF638EC6"/></dataBar></cfRule>`;
  const mixed = (sqref: string, formula: string, val: string) => rule(sqref, formula, bar(val));
  const source = mixed('D5:D10 F2', 'AND($A5&gt;1,D$3&gt;0,D4&lt;9)', '$B$6');
  eq(adjusted(source, rows('delete', 4, 1)), mixed('D5:D9 F2', 'AND($A5&gt;1,D$3&gt;0,D4&lt;9)', '$B$5'));
  eq(adjusted(source, rows('insert', 2, 2)), mixed('D7:D12 F2', 'AND($A7&gt;1,D$5&gt;0,D6&lt;9)', '$B$8'));
  eq(adjusted(source, cols('insert', 0, 1)), mixed('E5:E10 G2', 'AND($B5&gt;1,E$3&gt;0,E4&lt;9)', '$C$6'));
  eq(adjusted(source, cols('delete', 0, 1)), mixed('C5:C10 E2', 'AND(#REF!&gt;1,C$3&gt;0,C4&lt;9)', '$A$6'));
  // The first range goes: the rule is now written for the corner of what is left, F2.
  eq(adjusted(source, rows('delete', 4, 6)), mixed('F2', 'AND($A2&gt;1,F$3&gt;0,F1&lt;9)', '#REF!'));

  // Validations, the same way.
  eq(adjusted(validation('A5:A10', 'A5&gt;B4'), rows('delete', 4, 1)), validation('A5:A9', 'A5&gt;B4'));
  eq(adjusted(validation('A5:A10', 'A5&gt;B4'), rows('insert', 2, 2)), validation('A7:A12', 'A7&gt;B6'));
  eq(adjusted(validation('E6', '$G$5:$G$9'), rows('delete', 4, 1)), validation('E5', '$G$5:$G$8'));
  eq(adjusted(validation('E6', '$G$5:$G$9'), cols('insert', 0, 1)), validation('F6', '$H$5:$H$9'));
  eq(adjusted(validation('E6', '$G$5:$G$9'), rows('delete', 5, 1)), '');
  const both = '<dataValidations count="1"><dataValidation type="whole" operator="between" sqref="B5"><formula1>A5</formula1><formula2>$D$9</formula2></dataValidation></dataValidations>';
  eq(adjusted(both, rows('insert', 0, 1)), both.replace('B5', 'B6').replace('A5', 'A6').replace('$D$9', '$D$10'));

  // Two edits one after the other: each step uses the corner as that step moved it.
  const sheetXml = `<worksheet><sheetData/>${rule('A5:A10', 'A5&gt;$B$5')}</worksheet>`;
  const sheet = sheetFromXml(sheetXml);
  const wb = bareWorkbook(sheet);
  applyOp(wb, rows('insert', 0, 3));
  applyOp(wb, cols('insert', 0, 2));
  applyOp(wb, rows('delete', 8, 1));
  ok(decoder.decode(writeWorksheet(utf8(sheetXml), sheet, wb)).includes(rule('C8:C12', 'C8&gt;$D$8')));
});

test('write: the extension list forms of rules, validations and sparklines', () => {
  const x14 = (cf: string, cfRef: string, dv: string, dvRef: string, spark: string) =>
    '<extLst><ext uri="{c}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>' + cf + '</xm:f></x14:cfRule>' +
    `<xm:sqref>${cfRef}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>` +
    `<ext uri="{d}"><x14:dataValidations count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>${dv}</xm:f></x14:formula1><xm:sqref>${dvRef}</xm:sqref></x14:dataValidation></x14:dataValidations></ext>` +
    `<ext uri="{s}"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines>${spark}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>`;
  const spark = (f: string, sqref: string) => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${sqref}</xm:sqref></x14:sparkline>`;
  const source = x14('A5&gt;$B$7', 'A5:A10', 'B5:B9', 'C5:C6', spark('Veri!A5:D5', 'E5') + spark('Veri!A7:D7', 'E7'));
  // A sparkline's source is a plain range: it follows its cells; the sparkline goes with its own cell.
  eq(adjusted(source, rows('delete', 4, 1)), x14('A5&gt;$B$6', 'A5:A9', 'B5:B9', 'C5', spark('Veri!A6:D6', 'E6')));
  eq(adjusted(source, rows('insert', 0, 2)), x14('A7&gt;$B$9', 'A7:A12', 'B7:B11', 'C7:C8', spark('Veri!A7:D7', 'E7') + spark('Veri!A9:D9', 'E9')));
  eq(adjusted(source, cols('insert', 0, 1)), x14('B5&gt;$C$7', 'B5:B10', 'C5:C9', 'D5:D6', spark('Veri!B5:E5', 'F5') + spark('Veri!B7:E7', 'F7')));
});

test('write: a filter keeps its criteria on the columns they were set on', () => {
  const filter = (ref: string, first: number | undefined, second: number | undefined, sort = '') =>
    `<autoFilter ref="${ref}">` +
    (first === undefined ? '' : `<filterColumn colId="${first}"><filters><filter val="x"/></filters></filterColumn>`) +
    (second === undefined ? '' : `<filterColumn colId="${second}"><customFilters><customFilter operator="greaterThan" val="5"/></customFilters></filterColumn>`) +
    `${sort}</autoFilter>`;
  const sort = (ref: string, condition: string) => `<sortState ref="${ref}"><sortCondition ref="${condition}"/></sortState>`;
  // Criteria on C and E of B1:E9.
  const source = filter('B1:E9', 1, 3, sort('B2:E9', 'C2:C9'));
  eq(adjusted(source, cols('delete', 1, 1)), filter('B1:D9', 0, 2, sort('B2:D9', 'B2:B9')));
  // The filtered column itself goes: its criterion goes, and so does the sort that was on it.
  eq(adjusted(source, cols('delete', 2, 1)), filter('B1:D9', undefined, 2));
  eq(adjusted(source, cols('delete', 4, 1)), filter('B1:D9', 1, undefined, sort('B2:D9', 'C2:C9')));
  eq(adjusted(source, cols('insert', 0, 1)), filter('C1:F9', 1, 3, sort('C2:F9', 'D2:D9')));
  eq(adjusted(source, cols('insert', 2, 1)), filter('B1:F9', 2, 4, sort('B2:F9', 'D2:D9')));
  eq(adjusted(source, cols('insert', 5, 1)), source);
  eq(adjusted(source, cols('delete', 0, 1)), filter('A1:D9', 1, 3, sort('A2:D9', 'B2:B9')));
  eq(adjusted(source, rows('insert', 0, 2)), filter('B3:E11', 1, 3, sort('B4:E11', 'C4:C11')));
  eq(adjusted(source, cols('delete', 1, 4)), '');
});

test('write: only what is known to be a reference is moved', () => {
  const source =
    '<scenarios current="0"><scenario name="s" count="2"><inputCells r="A5" val="1"/><inputCells r="B7" val="2"/></scenario></scenarios>' +
    '<dataConsolidate><dataRefs count="2"><dataRef ref="A1:B5" sheet="veri"/><dataRef ref="A1:B5" sheet="Başka"/></dataRefs></dataConsolidate>' +
    '<cellWatches><cellWatch r="A5"/><cellWatch r="C9"/></cellWatches>' +
    // Elements this writer knows nothing about: "ref" and "r" may mean anything there.
    '<bilinmeyen ref="A5:B6" r="A5" activeCell="A5" topLeftCell="A5"/><gelecek><öğe ref="A5"/></gelecek>';
  eq(
    adjusted(source, rows('delete', 4, 1)),
    '<scenarios current="0"><scenario name="s" count="1"><inputCells r="B6" val="2"/></scenario></scenarios>' +
      '<dataConsolidate><dataRefs count="2"><dataRef ref="A1:B4" sheet="veri"/><dataRef ref="A1:B5" sheet="Başka"/></dataRefs></dataConsolidate>' +
      '<cellWatches><cellWatch r="C8"/></cellWatches>' +
      '<bilinmeyen ref="A5:B6" r="A5" activeCell="A5" topLeftCell="A5"/><gelecek><öğe ref="A5"/></gelecek>',
  );
  eq(
    adjusted(source, rows('insert', 0, 2)),
    source.replace('r="A5" val', 'r="A7" val').replace('"B7"', '"B9"').replace('ref="A1:B5" sheet="veri"', 'ref="A3:B7" sheet="veri"').replace('<cellWatch r="A5"/><cellWatch r="C9"/>', '<cellWatch r="A7"/><cellWatch r="C11"/>'),
  );
  // Every watched cell and every input cell gone: the containers go too.
  eq(adjusted('<scenarios><scenario name="s"><inputCells r="A5" val="1"/></scenario></scenarios><cellWatches><cellWatch r="A5"/></cellWatches><pageMargins left="1"/>', rows('delete', 4, 1)), '<pageMargins left="1"/>');
});

test('write: the position of a view stays; its selection follows the cells', () => {
  const view = (active: string) =>
    '<sheetViews><sheetView topLeftCell="C4" workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozen"/>' +
    `<selection pane="topRight" activeCell="B1" sqref="B1"/><selection pane="bottomRight" activeCell="${active}" sqref="${active}"/></sheetView></sheetViews>`;
  const run = (op: Op) => {
    const xml = `<worksheet>${view('B5')}<sheetData><row r="5"><c r="B5"><v>1</v></c></row></sheetData></worksheet>`;
    const sheet = sheetFromXml(xml);
    const wb = bareWorkbook(sheet);
    applyOp(wb, op);
    const out = decoder.decode(writeWorksheet(utf8(xml), sheet, wb));
    const again = sheetFromXml(out);
    // The file and the model agree on the frozen panes.
    deq([again.view.frozenRows, again.view.frozenCols, again.view.topLeftCell, again.view.firstVisibleCell], [sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.topLeftCell, sheet.view.firstVisibleCell]);
    return out.slice(out.indexOf('<sheetViews>'), out.indexOf('<sheetData'));
  };
  // Rows above the frozen part: the split and the first scrollable cell are where they were.
  eq(run(rows('insert', 0, 1)), view('B6').replace('activeCell="B1" sqref="B1"', 'activeCell="B2" sqref="B2"'));
  eq(run(rows('insert', 3, 2)), view('B7'));
  eq(run(cols('insert', 0, 1)), view('C5').replace('activeCell="B1" sqref="B1"', 'activeCell="C1" sqref="C1"'));
  // The selected cell is deleted: a view still needs a selection.
  eq(run(rows('delete', 4, 1)), view('A1'));
});

test('write: links regenerated from the model keep the attributes it does not hold', async () => {
  const sheetXml = worksheetXml(
    '<sheetData><row r="2"><c r="A2"><v>1</v></c></row><row r="4"><c r="A4"><v>2</v></c></row></sheetData>' +
      '<hyperlinks><hyperlink ref="A2" r:id="rId1" tooltip="bir" xr:uid="{AAAA}"/><hyperlink ref="A4" location="Veri!A9" display="iki" xr:uid="{BBBB}" gelecek="x&amp;y"/><hyperlink ref="A6" location="Veri!A1"/></hyperlinks>',
  );
  const build = async () => load(await buildXlsx({ sheets: [{ name: 'Veri', links: { rId1: 'https://example.com/' }, xml: sheetXml }] }));

  // The model changed its links without any rows or columns moving.
  let doc = await build();
  const sheet = doc.wb.sheets[0];
  sheet.hyperlinks = sheet.hyperlinks.map((link) => (link.location === 'Veri!A9' ? { ...link, location: 'Veri!A12' } : link)).filter((link) => link.location !== 'Veri!A1');
  sheet.hyperlinksModified = true;
  sheet.modified = true;
  let xml = await partText(await roundTrip(doc), SHEET1);
  ok(xml.includes('<hyperlinks><hyperlink ref="A2" r:id="rId1" tooltip="bir" xr:uid="{AAAA}"/><hyperlink ref="A4" location="Veri!A12" display="iki" xr:uid="{BBBB}" gelecek="x&amp;y"/></hyperlinks>'), xml);
  // The rows were not regenerated for it.
  ok(xml.includes('<sheetData><row r="2"><c r="A2"><v>1</v></c></row><row r="4"><c r="A4"><v>2</v></c></row></sheetData>'));

  // After rows moved, each link still finds the element it came from.
  doc = await build();
  structural(doc, rows('insert', 2, 3));
  xml = await partText(await roundTrip(doc), SHEET1);
  ok(xml.includes('<hyperlink ref="A2" r:id="rId1" tooltip="bir" xr:uid="{AAAA}"/><hyperlink ref="A7" location="Veri!A12" display="iki" xr:uid="{BBBB}" gelecek="x&amp;y"/><hyperlink ref="A9" location="Veri!A1"/>'), xml);
});

test('write: column settings regenerated from the model keep the attributes it does not hold', async () => {
  const sheetXml = worksheetXml('<cols><col min="1" max="1" width="5" bestFit="1"/><col min="2" max="5" width="9" customWidth="1" bestFit="1" phonetic="1" x:y="z&amp;"/><col min="8" max="8" width="3"/></cols><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>');
  const extra = ' bestFit="1" phonetic="1" x:y="z&amp;"';
  // A width changed in the middle of a range: its pieces are still that range.
  let doc = await load(await single(sheetXml.slice(sheetXml.indexOf('<cols>'), sheetXml.indexOf('</worksheet>'))));
  apply(doc, setColumnWidth(doc.wb, 0, 2, 2, 20));
  let xml = await partText(await roundTrip(doc), SHEET1);
  ok(
    xml.includes(
      `<cols><col min="1" max="1" width="5" bestFit="1"/><col min="2" max="2" width="9" customWidth="1"${extra}/><col min="3" max="3" width="20" customWidth="1"${extra}/>` +
        `<col min="4" max="5" width="9" customWidth="1"${extra}/><col min="8" max="8" width="3"/></cols>`,
    ),
    xml,
  );
  // Columns that had no settings get none of them.
  apply(doc, setColumnWidth(doc.wb, 0, 9, 9, 7));
  ok((await partText(await roundTrip(doc), SHEET1)).includes('<col min="8" max="8" width="3"/><col min="10" max="10" width="7" customWidth="1"/></cols>'));

  // After columns moved, the ranges are found where they went.
  doc = await load(await single(sheetXml.slice(sheetXml.indexOf('<cols>'), sheetXml.indexOf('</worksheet>'))));
  structural(doc, cols('insert', 2, 2));
  structural(doc, cols('delete', 0, 1));
  xml = await partText(await roundTrip(doc), SHEET1);
  ok(xml.includes(`<cols><col min="1" max="6" width="9" customWidth="1"${extra}/><col min="9" max="9" width="3"/></cols>`), xml);
});

test('write: a row format index without customFormat is written back', async () => {
  const doc = await load(
    await single('<sheetData><row r="1" spans="1:2" s="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c></row><row r="2" s="2" ht="20"><c r="A2"><v>2</v></c></row></sheetData>', { styles: STYLES }),
  );
  deq(doc.wb.sheets[0].rows.get(0)!.attrs, [['s', '1'], ['x14ac:dyDescent', '0.25']]);
  eq(doc.wb.sheets[0].rows.get(0)!.style, undefined);
  type(doc, 0, 'B1', '5');
  // The second row is given a format of its own: that one is written, once.
  apply(doc, formatRange(doc.wb, 0, { top: 1, bottom: 1, left: 0, right: MAX_COLS - 1 }, { bold: true }));
  // (The model keeps the old index among the row's attributes, so the two are compared through the file only.)
  const xml = await partText(await save(doc), SHEET1);
  const style = doc.wb.sheets[0].rows.get(1)!.style;
  ok(xml.includes('<row r="1" s="1" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c></row>'), xml);
  ok(xml.includes(`<row r="2" s="${style}" customFormat="1" ht="20">`), xml);
  eq((xml.match(/<row r="2"[^>]*>/) as RegExpMatchArray)[0].split(' s="').length, 2);
});

// --- other sheets ----------------------------------------------------------------------

test('write: raw content of other sheets follows the rows of the sheet it refers to', async () => {
  const other =
    '<sheetData><row r="1" spans="1:1"><c r="A1"><v>1</v></c></row></sheetData>' +
    rule('A1:A3', 'Veri!$A$5&gt;1') +
    '<dataValidations count="2"><dataValidation type="list" sqref="B1:B9"><formula1>veri!$A$1:$A$5</formula1></dataValidation>' +
    '<dataValidation type="custom" sqref="C1"><formula1>Veri!A5&gt;C1</formula1></dataValidation></dataValidations>' +
    '<pageMargins left="1"/>';
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Veri', xml: worksheetXml('<sheetData><row r="5"><c r="A5"><v>5</v></c></row></sheetData>') },
      { name: 'Diğer', xml: worksheetXml(other) },
      { name: 'Üçüncü', xml: worksheetXml(`<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Veri</t></is></c></row></sheetData>${rule('A1', 'A1&gt;1')}`) },
      { name: 'Veri 2', xml: worksheetXml(`<sheetData/>${rule('A1', "'Veri 2'!$A$5&gt;1")}`) },
    ],
  });
  const doc = await load(bytes);
  structural(doc, rows('insert', 0, 2));
  eq(doc.wb.sheets[1].modified, undefined);
  let saved = await roundTrip(doc);

  // The second sheet was not modified, yet what it says about the first one moved with it: the
  // absolute references follow the rows, the relative one keeps its distance from its own cell.
  const xml = await partText(saved, SHEET2);
  const expected = worksheetXml(other.replace('Veri!$A$5&gt;1', 'Veri!$A$7&gt;1').replace('veri!$A$1:$A$5', 'veri!$A$3:$A$7'));
  eq(xml, expected);
  // The sheets that name it only in a cell's text, or name another sheet, keep their bytes.
  assertOtherPartsUntouched(bytes, saved, [SHEET1, SHEET2, 'xl/workbook.xml']);

  // A second save starts from the original again; undoing puts the original back.
  eq(await partText(await save(doc), SHEET2), expected);
  doc.history.undo();
  saved = await roundTrip(doc);
  eq(await partText(saved, SHEET2), await partText(bytes, SHEET2));

  // A sheet that is modified for its own reasons gets the same treatment.
  doc.history.redo();
  type(doc, 1, 'A2', '7');
  ok((await partText(await roundTrip(doc), SHEET2)).includes('<row r="1" spans="1:1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>7</v></c></row></sheetData>' + rule('A1:A3', 'Veri!$A$7&gt;1')));
});

// --- drawings and tables -----------------------------------------------------------------

test('write: text links of shapes follow the cells they show', async () => {
  const linked =
    `${XML}<xdr:wsDr xmlns:xdr="${NS_XDR}">` +
    `<xdr:twoCellAnchor>${marker('from', 1, 1)}${marker('to', 3, 4)}<xdr:sp macro="" textlink="Resimli!$A$9"><xdr:nvSpPr><xdr:cNvPr id="2" name="bir"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>` +
    `<xdr:twoCellAnchor editAs="absolute">${marker('from', 1, 1)}${marker('to', 3, 4)}<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="3" name="grup"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/>` +
    '<xdr:sp textlink="$B$6"><xdr:nvSpPr><xdr:cNvPr id="4" name="iki"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:cxnSp textlink="Öteki!$A$9"><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="üç"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr/></xdr:cxnSp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>' +
    '</xdr:wsDr>';
  const doc = await load(await anchoredWorkbook(vml(), false, linked));
  structural(doc, rows('insert', 5, 2));
  const xml = await partText(await roundTrip(doc), DRAWING_PART);
  // The anchors of the first object did not move (it lies above the new rows); its link did.
  eq(xml, linked.replace('?>', '?>\n').replace('Resimli!$A$9', 'Resimli!$A$11').replace('textlink="$B$6"', 'textlink="$B$8"'));
});

const TABLE_PART = 'xl/tables/table1.xml';
const TABLE =
  `${XML}<table xmlns="${NS}" id="1" name="Tablo1" displayName="Tablo1" ref="B2:D4" totalsRowShown="0"><autoFilter ref="B2:D4"/>` +
  '<tableColumns count="3"><tableColumn id="1" name="Ad"/><tableColumn id="2" name="Fiyat" dataDxfId="0"/><tableColumn id="3" name="Not"><calculatedColumnFormula>Tablo1[[#This Row],[Fiyat]]*2</calculatedColumnFormula></tableColumn></tableColumns>' +
  '<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>';

async function tableWorkbook(): Promise<Uint8Array> {
  const bytes = await buildXlsx({
    sharedStrings: '<si><t>Ad</t></si><si><t>Fiyat</t></si><si><t>Not</t></si>',
    sheets: [
      {
        name: 'Tablolu',
        xml: worksheetXml(
          '<sheetData><row r="2"><c r="B2" t="s"><v>0</v></c><c r="C2" t="s"><v>1</v></c><c r="D2" t="s"><v>2</v></c></row><row r="3"><c r="B3"><v>1</v></c><c r="C3"><v>2</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts>',
        ),
      },
    ],
  });
  const zip = ZipArchive.read(bytes);
  zip.set(TABLE_PART, utf8(TABLE));
  zip.set(SHEET1_RELS, utf8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/></Relationships>`));
  zip.set('[Content_Types].xml', utf8(decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>')));
  return zip.write();
}

const columnNames = (xml: string) => [...xml.matchAll(/<tableColumn id="\d+" name="([^"]*)"/g)].map((m) => m[1]);

test('write: the columns of a table are named after its header cells', async () => {
  const bytes = await tableWorkbook();
  let doc = await load(bytes);
  deq(doc.wb.sheets[0].anchored, ['table']);

  // Editing below the header leaves the table part alone.
  type(doc, 0, 'B3', '9');
  let saved = await roundTrip(doc);
  assertOtherPartsUntouched(bytes, saved, [SHEET1, 'xl/workbook.xml']);

  // A new header text is the new column name; everything else in the table stays.
  type(doc, 0, 'B2', 'Ürün & "Ad"');
  saved = await roundTrip(doc);
  assertOtherPartsUntouched(bytes, saved, [SHEET1, 'xl/workbook.xml', TABLE_PART]);
  eq(await partText(saved, TABLE_PART), TABLE.replace('?>', '?>\n').replace('name="Ad"', 'name="Ürün &amp; &quot;Ad&quot;"'));
  ok((await partText(saved, SHEET1)).includes('<row r="2"><c r="B2" t="inlineStr"><is><t>Ürün &amp; "Ad"</t></is></c><c r="C2" t="s"><v>1</v></c><c r="D2" t="s"><v>2</v></c></row>'));

  // Undone: the table is as the file had it again.
  doc.history.undo();
  eq(await partText(await roundTrip(doc), TABLE_PART), TABLE);

  // An empty header, a repeated one (whatever its case) and a number: the names are made valid the
  // way the application makes them, and the header cells are written to match.
  doc = await load(bytes);
  put(doc, 0, 'B2', { kind: 'string', value: 'not', style: 0 });
  put(doc, 0, 'C2', undefined);
  saved = await save(doc);
  deq(columnNames(await partText(saved, TABLE_PART)), ['not', 'Column2', 'Not2']);
  ok((await partText(saved, SHEET1)).includes('<row r="2"><c r="B2" t="inlineStr"><is><t>not</t></is></c><c r="C2" t="inlineStr"><is><t>Column2</t></is></c><c r="D2" t="inlineStr"><is><t>Not2</t></is></c></row>'));
  const again = await reopen(saved);
  deq([getCell(again.sheets[0], 1, 1)!.value, getCell(again.sheets[0], 1, 2)!.value, getCell(again.sheets[0], 1, 3)!.value], ['not', 'Column2', 'Not2']);

  doc = await load(bytes);
  put(doc, 0, 'C2', { kind: 'number', value: 2024, style: 0 });
  put(doc, 0, 'D2', { kind: 'number', value: 2024, style: 0 });
  saved = await save(doc);
  deq(columnNames(await partText(saved, TABLE_PART)), ['Ad', '2024', '20242']);
  ok((await partText(saved, SHEET1)).includes('<c r="C2" t="inlineStr"><is><t>2024</t></is></c><c r="D2" t="inlineStr"><is><t>20242</t></is></c>'));

  // The whole header row emptied: the row is written for the names alone.
  doc = await load(bytes);
  for (const ref of ['B2', 'C2', 'D2']) put(doc, 0, ref, undefined);
  saved = await save(doc);
  deq(columnNames(await partText(saved, TABLE_PART)), ['Column1', 'Column2', 'Column3']);
  ok((await partText(saved, SHEET1)).includes('<sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>Column1</t></is></c><c r="C2" t="inlineStr"><is><t>Column2</t></is></c><c r="D2" t="inlineStr"><is><t>Column3</t></is></c></row><row r="3">'));
});

// --- hostile and unusual input --------------------------------------------------------------

test('write: raw content nested too deeply to adjust is refused with a clear error', async () => {
  const depth = 100000;
  const deep = `<extLst>${'<ext>'.repeat(depth)}${'</ext>'.repeat(depth)}</extLst>`;
  const doc = await load(await single(`<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>${deep}`));
  // Without a structural edit the element is copied, however deep it is.
  type(doc, 0, 'A1', '2');
  ok((await partText(await save(doc), SHEET1)).endsWith(`${deep}</worksheet>`));
  structural(doc, rows('insert', 0, 1));
  const before = await doc.file.package!.save();
  await assert.rejects(writeXlsx(doc.file.package!, doc.wb), /<extLst> is nested too deeply to adjust/);
  deq(await doc.file.package!.save(), before);
});

test('write: a styles part added to a strict workbook gets a strict relationship', async () => {
  const strict = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
  const zip = ZipArchive.read(await single('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'));
  zip.set('xl/_rels/workbook.xml.rels', utf8(decoder.decode(await zip.read('xl/_rels/workbook.xml.rels')).split(`${REL}/worksheet`).join(`${strict}/worksheet`)));
  const doc = await load(await zip.write());
  type(doc, 0, 'A2', '12%');
  const saved = await roundTrip(doc);
  ok((await partText(saved, 'xl/_rels/workbook.xml.rels')).includes(`<Relationship Id="rId2" Type="${strict}/styles" Target="styles.xml"/>`));
  eq((await reopen(saved)).styles.cellXfs.length, doc.wb.styles.cellXfs.length);
});

// --- whole columns and rows in reference lists ---------------------------------------------

test('write: whole columns and rows of rules and validations stay whole through structural edits', () => {
  // A column inserted before a whole column moves it; "C" alone would not be a reference.
  eq(adjusted(rule('B:B', 'B1&gt;0'), cols('insert', 0, 1)), rule('C:C', 'C1&gt;0'));
  eq(adjusted(rule('$B:$B', 'B1&gt;0'), cols('insert', 0, 1)), rule('$C:$C', 'C1&gt;0'));
  eq(adjusted(rule('3:3', 'A3&gt;0'), rows('insert', 0, 1)), rule('4:4', 'A4&gt;0'));
  eq(adjusted(rule('D1:D4 B:B', 'D1&gt;0'), cols('insert', 0, 1)), rule('E1:E4 C:C', 'E1&gt;0'));
  eq(adjusted(rule('B:B D1:D4', 'B1&gt;0'), cols('insert', 0, 1)), rule('C:C E1:E4', 'C1&gt;0'));
  // Rows do not move a whole column, columns do not move a whole row; what is anchored still follows.
  eq(adjusted(rule('B:B', 'B1&gt;$A$5'), rows('insert', 0, 1)), rule('B:B', 'B1&gt;$A$6'));
  eq(adjusted(rule('B:B', 'B1&gt;0'), rows('delete', 0, 3)), rule('B:B', 'B1&gt;0'));
  eq(adjusted(rule('D1:D4 B:B', 'D1&gt;0'), rows('insert', 0, 1)), rule('D2:D5 B:B', 'D2&gt;0'));
  eq(adjusted(rule('3:3', 'A3&gt;$B$1'), cols('insert', 0, 2)), rule('3:3', 'A3&gt;$D$1'));
  // The same column written as a range keeps that form.
  eq(adjusted(rule('B1:B1048576', 'B1&gt;0'), cols('insert', 0, 1)), rule('C1:C1048576', 'C1&gt;0'));
  eq(adjusted(rule('B1:B1048576', 'B1&gt;0'), rows('insert', 0, 1)), rule('B1:B1048576', 'B1&gt;0'));
  eq(adjusted(rule('B1:B1048576', 'B1&gt;0'), rows('delete', 2, 4)), rule('B1:B1048576', 'B1&gt;0'));
  eq(adjusted(rule('A3:XFD3', 'A3&gt;0'), cols('insert', 0, 1)), rule('A3:XFD3', 'A3&gt;0'));
  // Deleting: a span shrinks, a range that is gone leaves the list, the last one takes the element along.
  eq(adjusted(rule('B:C', 'B1&gt;0'), cols('delete', 1, 1)), rule('B:B', 'B1&gt;0'));
  eq(adjusted(rule('2:3', 'A2&gt;0'), rows('delete', 1, 1)), rule('2:2', 'A2&gt;0'));
  eq(adjusted(rule('B:B D1:D3', 'B1&gt;0'), cols('delete', 1, 1)), rule('C1:C3', 'C1&gt;0'));
  eq(adjusted(rule('B:B', 'B1&gt;0'), cols('delete', 1, 1)), '');
  eq(adjusted(rule('3:3', 'A3&gt;0'), rows('delete', 2, 1)), '');
  // A range that shrinks to one cell is still written as that cell.
  eq(adjusted(rule('B2:B3', 'B2&gt;0'), rows('delete', 2, 1)), rule('B2', 'B2&gt;0'));

  // Validations and every other element that carries a reference list.
  eq(adjusted(validation('B:B', 'B1&gt;0'), cols('insert', 0, 1)), validation('C:C', 'C1&gt;0'));
  eq(adjusted(validation('B:B', 'B1&gt;0'), rows('insert', 0, 1)), validation('B:B', 'B1&gt;0'));
  eq(adjusted(validation('3:3 B:B', 'A3&gt;0'), rows('insert', 0, 1)), validation('4:4 B:B', 'A4&gt;0'));
  const others = (a: string, b: string, c: string) =>
    `<protectedRanges><protectedRange sqref="${a}" name="p"/></protectedRanges><ignoredErrors><ignoredError sqref="${b}" numberStoredAsText="1"/></ignoredErrors>` +
    `<extLst><ext uri="{c}"><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="expression"><xm:f>B1&gt;0</xm:f></x14:cfRule><xm:sqref>${c}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`;
  eq(adjusted(others('B:B', '2:2 B:C', 'B:B'), cols('insert', 0, 1)), others('C:C', '2:2 C:D', 'C:C').replace('B1&gt;0', 'C1&gt;0'));
  eq(adjusted(others('B:B', '2:2 B:C', 'B:B'), rows('insert', 0, 1)), others('B:B', '3:3 B:C', 'B:B'));
});

test('write: a relative reference of a rule that leaves the sheet goes round its edge', () => {
  // "The cell above is positive", and row 1 goes: for the new first row that is the last row of the sheet.
  eq(adjusted(rule('B2:B5', 'B1&gt;0'), rows('delete', 0, 1)), rule('B1:B4', 'B1048576&gt;0'));
  eq(adjusted(rule('B2:B5', 'A2&gt;0'), cols('delete', 0, 1)), rule('A2:A5', 'XFD2&gt;0'));
  // And back again when a row comes in above.
  eq(adjusted(rule('B1:B4', 'B1048576&gt;0'), rows('insert', 0, 1)), rule('B2:B5', 'B1&gt;0'));
  eq(adjusted(validation('B2:B5', 'B1&gt;0'), rows('delete', 0, 1)), validation('B1:B4', 'B1048576&gt;0'));
  // What is anchored with "$" named a cell that is gone.
  eq(adjusted(rule('B2:B5', '$B$1&gt;0'), rows('delete', 0, 1)), rule('B1:B4', '#REF!&gt;0'));
});
