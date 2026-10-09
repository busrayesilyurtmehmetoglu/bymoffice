// Fourth review of the sheets of a workbook: what a copy of a sheet carries and says it loses,
// the filter after cells were cleared, and what a save keeps of a part with other line ends.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { ZipArchive, open } from '../src/index';
import type { OfficeFile, Workbook } from '../src/index';
import { clearRange } from '../src/ops/commands';
import { filterDataRows, filterSummary, filterValues, setAutoFilter, setFilterColumn } from '../src/ops/filter';
import { History, type Op } from '../src/ops/ops';
import { addConditionalRule, removeConditionalRule, setValidation } from '../src/ops/rules';
import { copySheet, copySheetLosses, renameSheet } from '../src/ops/sheets';
import { getUiStrings } from '../src/ui/strings';
import { buildXlsx, worksheetXml } from './fixtures';

const decoder = new TextDecoder();

interface Loaded {
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const reopen = async (bytes: Uint8Array) => (await open(bytes)).workbook();
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const text = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));

const cells = (rows: number): string => {
  let out = '';
  for (let r = 1; r <= rows; r++) out += `<row r="${r}"><c r="A${r}"><v>${r}</v></c></row>`;
  return `<sheetData>${out}</sheetData>`;
};

// ------------------------------------------------------------------------- B1

const PROTECTION = '<sheetProtection sheet="1" objects="1" scenarios="1"/>';
const RANGES =
  '<protectedRanges><protectedRange sqref="A1:A2" name="Open" securityDescriptor="O:WDG:WD"/><protectedRange password="CC1A" sqref="C8 A3:B3" name="Two"/>' +
  '<protectedRange sqref="D5:D6" name="Gone"><securityDescriptor>O:WDG:WDD:(A;;CC;;;S-1-5-21)</securityDescriptor></protectedRange></protectedRanges>';

const guarded = () => buildXlsx({ sheets: [{ name: 'Kor', xml: worksheetXml(`${cells(8)}${PROTECTION}${RANGES}`) }] });

const sqrefs = (xml: string): string[] => Array.from(xml.matchAll(/<protectedRange\b[^>]*?\bsqref="([^"]*)"/g), (match) => match[1]);

test('review 4 (B1): the open ranges of a protected sheet are, in its copy, where the rows and columns moved them', async () => {
  const doc = await load(await guarded());
  // Three rows above everything, then the rows the third range is in go.
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 }]);
  apply(doc, [{ type: 'delete', sheet: 0, axis: 'row', at: 7, count: 2 }]);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 }]);
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  const source = await text(saved, 'xl/worksheets/sheet1.xml');
  const copy = await text(saved, 'xl/worksheets/sheet2.xml');
  // The sheet of the file, as the writer moves its ranges; the copy says the same.
  assert.deepEqual(sqrefs(source), ['A4:A5', 'D9 A6:C6']);
  assert.deepEqual(sqrefs(copy), sqrefs(source));
  assert.ok(copy.indexOf(PROTECTION) >= 0, copy);
  assert.ok(/<protectedRange sqref="A4:A5" name="Open" securityDescriptor="O:WDG:WD"\/>/.test(copy), copy);
  assert.ok(/<protectedRange password="CC1A" sqref="D9 A6:C6" name="Two"\/>/.test(copy), copy);
  // The range all of whose cells were deleted is gone, in both.
  assert.equal(copy.indexOf('Gone'), -1, copy);
  assert.equal(source.indexOf('Gone'), -1, source);
  const again = await reopen(saved);
  assert.deepEqual(again.sheets[1].protection, again.sheets[0].protection);
  // The model of the source is not changed by the copy: its protection is still what the file said.
  assert.deepEqual(
    doc.wb.sheets[0].protection!.ranges!.map((range) => range.attrs.find((pair) => pair[0] === 'sqref')![1]),
    ['A1:A2', 'C8 A3:B3', 'D5:D6'],
  );
  // A file built from the model says what the saved one says, for the sheet of the file as for its copy.
  const built = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.deepEqual(built.sheets[1].protection, again.sheets[1].protection);
  assert.deepEqual(built.sheets[0].protection, again.sheets[0].protection);
  // Rows moved in the copy afterwards move its ranges too, in the file saved and in the file built.
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 2 }]);
  const later = await save(doc);
  assert.deepEqual(sqrefs(await text(later, 'xl/worksheets/sheet2.xml')), ['A6:A7', 'D11 A8:C8']);
  assert.deepEqual(sqrefs(await text(later, 'xl/worksheets/sheet1.xml')), ['A4:A5', 'D9 A6:C6']);
  const builtLater = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.deepEqual(builtLater.sheets[1].protection, (await reopen(later)).sheets[1].protection);
});

test('review 4 (B1): without rows or columns moved the copy shares the protection of its source', async () => {
  const doc = await load(await guarded());
  apply(doc, copySheet(doc.wb, 0));
  assert.equal(doc.wb.sheets[1].protection, doc.wb.sheets[0].protection);
  // A descriptor that is a child element travels with its range.
  const copy = await text(await save(doc), 'xl/worksheets/sheet2.xml');
  assert.ok(copy.indexOf('<securityDescriptor>O:WDG:WDD:(A;;CC;;;S-1-5-21)</securityDescriptor>') > 0, copy);
});

// ------------------------------------------------------------------------- B4

const X14AC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"';
const excelSheet = (inner: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ${X14AC}>${inner}</worksheet>`;

/** Rows as Excel writes them: the span of their cells and the descent of their font on every one. */
const excelRows = (extra = '') =>
  '<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/><sheetData>' +
  '<row r="1" spans="1:2" x14ac:dyDescent="0.25"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>' +
  `<row r="2" spans="1:2"${extra} x14ac:dyDescent="0.25"><c r="A2"><v>3</v></c></row>` +
  '<row r="3" spans="1:2" ht="30" customHeight="1" x14ac:dyDescent="0.25"><c r="A3"><v>4</v></c></row></sheetData>';

test('review 4 (B4): the copy of a sheet as Excel writes it loses nothing: the hints on its rows are no content', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Veri', xml: excelSheet(excelRows()) }] }));
  assert.deepEqual(copySheetLosses(doc.wb, 0), []);
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  const copy = await text(saved, 'xl/worksheets/sheet2.xml');
  // The new part does not declare the namespace of the hint: it is not written there.
  assert.equal(copy.indexOf('dyDescent'), -1, copy);
  assert.ok(/<row r="3" ht="30" customHeight="1">/.test(copy), copy);
  const again = await reopen(saved);
  assert.equal(again.sheets[1].rows.get(2)!.height, 30);
});

test('review 4 (B4): a setting of a row that the model has no field for goes into the copy as the file has it', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Veri', xml: excelSheet(excelRows(' s="0" thickBot="1" ph="1"')) }] }));
  // Nothing is lost: the copy carries the attributes that need no namespace.
  assert.deepEqual(copySheetLosses(doc.wb, 0), []);
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  const copy = await text(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(/<row r="2" s="0" thickBot="1" ph="1">/.test(copy), copy);
  const again = await reopen(saved);
  assert.deepEqual(again.sheets[1].rows.get(1)!.attrs, [
    ['s', '0'],
    ['thickBot', '1'],
    ['ph', '1'],
  ]);
  // The rows of the source are as they were: the copy has lists of its own.
  assert.deepEqual(
    doc.wb.sheets[0].rows.get(1)!.attrs!.map((pair) => pair[0]),
    ['s', 'thickBot', 'ph', 'x14ac:dyDescent'],
  );
  const built = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.deepEqual(built.sheets[1].rows.get(1)!.attrs, again.sheets[1].rows.get(1)!.attrs);
});

test('review 4 (B4): an attribute of another namespace that is no hint is still counted as lost', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Veri', xml: excelSheet(excelRows(' x14ac:somethingElse="1"')) }] }));
  assert.deepEqual(copySheetLosses(doc.wb, 0), [{ what: 'rowAttributes', count: 1 }]);
});

// ------------------------------------------------------------------------- B5

test('review 4 (B5): the picture of a header or footer and the background picture of a sheet are named as what a copy loses', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [
        {
          name: 'Veri',
          xml: worksheetXml(`${cells(2)}<headerFooter><oddHeader>&amp;C&amp;G</oddHeader></headerFooter><legacyDrawingHF r:id="rIdHF"/><picture r:id="rIdBack"/>`),
          links: { rIdHF: 'https://example.org/hf.vml', rIdBack: 'https://example.org/back.png' },
        },
      ],
    }),
  );
  assert.deepEqual(doc.wb.sheets[0].unmodelled, ['legacyDrawingHF', 'picture']);
  assert.deepEqual(copySheetLosses(doc.wb, 0), [
    { what: 'legacyDrawingHF', count: 1 },
    { what: 'picture', count: 1 },
  ]);
  for (const locale of ['tr', 'en']) {
    const strings = getUiStrings(locale);
    for (const what of ['legacyDrawingHF', 'picture']) {
      const said = strings.copyLoss(what, 1);
      // Named in the language, not by the name of the element.
      assert.ok(/^1 \S+ \S+/.test(said) && said !== what, `${locale} ${what}: ${said}`);
    }
  }
  // The sheet of the file keeps both; the copy has neither.
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  assert.ok((await text(saved, 'xl/worksheets/sheet1.xml')).indexOf('<legacyDrawingHF r:id="rIdHF"/><picture r:id="rIdBack"/>') > 0);
  const copy = await text(saved, 'xl/worksheets/sheet2.xml');
  assert.equal(copy.indexOf('legacyDrawingHF'), -1, copy);
  assert.equal(copy.indexOf('<picture'), -1, copy);
});

// ------------------------------------------------------------------------- B6

const listed = (filter = '') => {
  let rows = '<row r="1"><c r="A1" t="inlineStr"><is><t>Ad</t></is></c><c r="B1" t="inlineStr"><is><t>Not</t></is></c></row>';
  for (let r = 2; r <= 7; r++) rows += `<row r="${r}"><c r="A${r}"><v>${r}</v></c>${r === 4 ? '' : `<c r="B${r}"><v>${r * 10}</v></c>`}</row>`;
  return buildXlsx({ sheets: [{ name: 'Liste', xml: worksheetXml(`<sheetData>${rows}</sheetData>${filter}`) }] });
};

const hiddenOf = (wb: Workbook): number[] =>
  Array.from(wb.sheets[0].rows.entries())
    .filter(([, row]) => row.hidden)
    .map(([index]) => index)
    .sort((a, b) => a - b);

test('review 4 (B6): a filter switched on over whole columns ends at the last row that holds something, cleared rows aside', async () => {
  const doc = await load(await listed());
  // The last two rows are emptied: the sheet still remembers that it once had seven rows.
  apply(doc, clearRange(doc.wb, 0, { top: 5, left: 0, bottom: 6, right: 1 }));
  assert.equal(doc.wb.sheets[0].rowCount, 7);
  const again = await load(await save(doc));
  const whole = { top: 0, left: 0, bottom: 1048575, right: 1 };
  apply(doc, setAutoFilter(doc.wb, 0, whole));
  apply(again, setAutoFilter(again.wb, 0, whole));
  assert.deepEqual(doc.wb.sheets[0].autoFilter!.range, { top: 0, left: 0, bottom: 4, right: 1 });
  assert.deepEqual(again.wb.sheets[0].autoFilter!.range, doc.wb.sheets[0].autoFilter!.range);
  // An emptied sheet: the filter is the header row, as in a sheet that never held more.
  apply(doc, setAutoFilter(doc.wb, 0, undefined));
  apply(doc, clearRange(doc.wb, 0, { top: 1, left: 0, bottom: 6, right: 1 }));
  apply(doc, setAutoFilter(doc.wb, 0, whole));
  assert.deepEqual(doc.wb.sheets[0].autoFilter!.range, { top: 0, left: 0, bottom: 0, right: 1 });
});

test('review 4 (B6): the same criterion hides the same rows in the workbook and in the file saved from it', async () => {
  // The filter of the file runs further down than its cells do.
  const doc = await load(await listed('<autoFilter ref="A1:B12"/>'));
  apply(doc, clearRange(doc.wb, 0, { top: 5, left: 0, bottom: 6, right: 1 }));
  const again = await load(await save(doc));
  for (const criterion of [{ kind: 'values', values: ['20', '30'], blank: false }, { kind: 'values', values: [], blank: true }, { kind: 'top10', top: true, value: 1 }] as const) {
    apply(doc, setFilterColumn(doc.wb, 0, 1, criterion));
    apply(again, setFilterColumn(again.wb, 0, 1, criterion));
    assert.deepEqual(hiddenOf(doc.wb), hiddenOf(again.wb), JSON.stringify(criterion));
    assert.deepEqual(filterSummary(doc.wb.sheets[0]), filterSummary(again.wb.sheets[0]), JSON.stringify(criterion));
    assert.deepEqual(filterValues(doc.wb, 0, 1), filterValues(again.wb, 0, 1), JSON.stringify(criterion));
    // And what is saved now reads as what the model shows.
    const read = await reopen(await save(doc));
    assert.deepEqual(hiddenOf(read), hiddenOf(doc.wb), JSON.stringify(criterion));
  }
  // The rows under the cells are no rows of the list: none of them is hidden or counted.
  assert.ok(hiddenOf(doc.wb).every((index) => index <= 4), String(hiddenOf(doc.wb)));
  assert.deepEqual(filterDataRows(doc.wb.sheets[0], doc.wb.sheets[0].autoFilter!), { top: 1, bottom: 4 });
});

test('review 4 (B6): a row without cells below the list is no row of the list', async () => {
  const doc = await load(await listed('<autoFilter ref="A1:B12"/>'));
  const sheet = doc.wb.sheets[0];
  assert.deepEqual(filterDataRows(sheet, sheet.autoFilter!), { top: 1, bottom: 6 });
  // A row object without cells (one whose height was set) holds nothing.
  apply(doc, [{ type: 'setRows', sheet: 0, rows: [{ row: 9, props: { height: 30, customHeight: true } }] }]);
  assert.deepEqual(filterDataRows(sheet, sheet.autoFilter!), { top: 1, bottom: 6 });
});

// ------------------------------------------------------------------------- B7

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const BOM = '﻿';
/** A sheet part as an editor on Windows may leave it: a byte order mark and CRLF line ends, indented. */
const crlfSheet = (rules = '') =>
  `${BOM}<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="${NS}">\r\n  <sheetViews>\r\n    <sheetView workbookViewId="0"/>\r\n  </sheetViews>\r\n` +
  '  <sheetData>\r\n    <row r="1">\r\n      <c r="A1"><v>5</v></c>\r\n      <c r="B1" t="inlineStr"><is><t xml:space="preserve">two\r\nlines</t></is></c>\r\n    </row>\r\n  </sheetData>\r\n' +
  `${rules}  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\r\n  <!-- kept -->\r\n</worksheet>\r\n`;

const bytesOf = async (bytes: Uint8Array, name: string) => ZipArchive.read(bytes).read(name);

/** What two texts share at their start and at their end, and what each has in between. */
function between(before: string, after: string): { head: string; tail: string; was: string; is: string } {
  let p = 0;
  while (p < before.length && p < after.length && before.charCodeAt(p) === after.charCodeAt(p)) p++;
  let q = 0;
  while (q < before.length - p && q < after.length - p && before.charCodeAt(before.length - 1 - q) === after.charCodeAt(after.length - 1 - q)) q++;
  return { head: before.slice(0, p), tail: before.slice(before.length - q), was: before.slice(p, before.length - q), is: after.slice(p, after.length - q) };
}

const exact = (bytes: Uint8Array): string => new TextDecoder('utf-8', { ignoreBOM: true }).decode(bytes);

test('review 4 (B7): a rule added to a part with a byte order mark and CRLF line ends changes the bytes of its element alone', async () => {
  const source = crlfSheet();
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: source }] }));
  apply(doc, addConditionalRule(doc.wb, 0, { top: 0, left: 0, bottom: 2, right: 0 }, { type: 'duplicateValues' }));
  const saved = await save(doc);
  const written = exact(await bytesOf(saved, 'xl/worksheets/sheet1.xml'));
  const diff = between(source, written);
  assert.equal(diff.was, '', JSON.stringify(diff.was));
  assert.ok(/^<conditionalFormatting sqref="A1:A3"><cfRule type="duplicateValues"[^>]*\/><\/conditionalFormatting>$/.test(diff.is), JSON.stringify(diff.is));
  // Everything else is the file: the mark, every line end, the text of the cell.
  assert.equal(diff.head + diff.tail, source);
  assert.equal(written.charCodeAt(0), 0xfeff);
  assert.equal((await reopen(saved)).sheets[0].conditionalFormats!.length, 1);
  // A validation as well: two places change, nothing between them or around them.
  apply(doc, setValidation(doc.wb, 0, { top: 0, left: 2, bottom: 0, right: 2 }, { type: 'list', formula1: '"a,b"' }));
  const both = exact(await bytesOf(await save(doc), 'xl/worksheets/sheet1.xml'));
  const stripped = both.replace(/<conditionalFormatting .*?<\/conditionalFormatting>/, '').replace(/<dataValidations .*?<\/dataValidations>/, '');
  assert.equal(stripped, source);
  // Undone, the part is the file again.
  doc.history.undo();
  doc.history.undo();
  assert.equal(exact(await bytesOf(await save(doc), 'xl/worksheets/sheet1.xml')), source);
});

test('review 4 (B7): a rule changed and one removed in such a part: the other elements keep their bytes, line ends inside them included', async () => {
  const rules =
    '  <conditionalFormatting sqref="A1:A3">\r\n    <cfRule type="cellIs" priority="1" operator="equal">\r\n      <formula>5</formula>\r\n    </cfRule>\r\n  </conditionalFormatting>\r\n' +
    '  <conditionalFormatting sqref="D1:D3">\r\n    <cfRule type="cellIs" priority="2" operator="equal">\r\n      <formula>7</formula>\r\n    </cfRule>\r\n  </conditionalFormatting>\r\n' +
    '  <dataValidations count="1">\r\n    <dataValidation type="whole" sqref="E1">\r\n      <formula1>1</formula1>\r\n      <formula2>2</formula2>\r\n    </dataValidation>\r\n  </dataValidations>\r\n';
  const source = crlfSheet(rules);
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: source }] }));
  apply(doc, removeConditionalRule(doc.wb, 0, 0));
  const written = exact(await bytesOf(await save(doc), 'xl/worksheets/sheet1.xml'));
  // The element of the first format is gone, and the rule that is left is the first now. Every other byte is the
  // file's: the line ends between the elements and inside the ones that stayed.
  const first = '<conditionalFormatting sqref="A1:A3">\r\n    <cfRule type="cellIs" priority="1" operator="equal">\r\n      <formula>5</formula>\r\n    </cfRule>\r\n  </conditionalFormatting>';
  assert.ok(source.indexOf(first) > 0);
  assert.equal(written, source.replace(first, '').replace('priority="2"', 'priority="1"'));
  assert.ok(written.indexOf('two\r\nlines') > 0);
});

test('review 4 (B7): a sheet renamed elsewhere changes the name in such a part and nothing else', async () => {
  const rules = '  <dataValidations count="1">\r\n    <dataValidation type="list" sqref="E1">\r\n      <formula1>Kaynak!$A$1:$A$3</formula1>\r\n    </dataValidation>\r\n  </dataValidations>\r\n';
  const source = crlfSheet(rules);
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: source }, { name: 'Kaynak', xml: worksheetXml(cells(3)) }] }));
  apply(doc, renameSheet(doc.wb, 1, 'Liste'));
  const written = exact(await bytesOf(await save(doc), 'xl/worksheets/sheet1.xml'));
  assert.equal(written, source.replace('Kaynak!$A$1:$A$3', 'Liste!$A$1:$A$3'));
});

test('review 4 (B7): a part without a mark and with plain line ends is written as before', async () => {
  const source = crlfSheet().replace(/\r\n/g, '\n').slice(1);
  const doc = await load(await buildXlsx({ sheets: [{ name: 'Data', xml: source }] }));
  apply(doc, addConditionalRule(doc.wb, 0, { top: 0, left: 0, bottom: 2, right: 0 }, { type: 'duplicateValues' }));
  const written = exact(await bytesOf(await save(doc), 'xl/worksheets/sheet1.xml'));
  const diff = between(source, written);
  assert.equal(diff.was, '');
  assert.equal(diff.head + diff.tail, source);
});
