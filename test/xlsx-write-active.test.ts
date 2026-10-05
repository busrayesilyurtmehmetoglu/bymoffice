// Saving the sheet a workbook opens on (Workbook.activeSheet): written with the rest of a
// change, never on its own - a file that was only looked at is saved byte for byte.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { ZipArchive, open, parseCellRef, parseXml } from '../src/index';
import type { OfficeFile, Workbook, XmlElement } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, worksheetXml } from './fixtures';

const decoder = new TextDecoder();
const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;

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

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

function type(doc: Loaded, sheet: number, ref: string, text: string): void {
  const { row, col } = parseCellRef(ref)!;
  apply(doc, setCellInput(doc.wb, sheet, row, col, text));
}

async function content(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = ZipArchive.read(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of zip.names()) out.set(name, await zip.read(name));
  return out;
}

const sameBytes = (a: Uint8Array | undefined, b: Uint8Array | undefined) => !!a && !!b && a.length === b.length && a.every((value, i) => value === b[i]);

/** The names of the entries whose content is not what it was. */
async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = await content(before);
  const b = await content(after);
  const out: string[] = [];
  for (const name of new Set([...a.keys(), ...b.keys()])) if (!sameBytes(a.get(name), b.get(name))) out.push(name);
  return out.sort();
}

const all = (el: XmlElement, local: string, out: XmlElement[] = []): XmlElement[] => {
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (node.name.replace(/^.*:/, '') === local) out.push(node);
    all(node, local, out);
  }
  return out;
};

/** What the file says about the sheet it opens on: activeTab, and tabSelected of each sheet part in order. */
async function selection(bytes: Uint8Array, count = 3): Promise<{ activeTab: string | undefined; selected: boolean[] }> {
  const parts = await content(bytes);
  const view = all(parseXml(parts.get('xl/workbook.xml')!).root, 'workbookView')[0];
  const selected: boolean[] = [];
  for (let n = 1; n <= count; n++) {
    const views = all(parseXml(parts.get(`xl/worksheets/sheet${n}.xml`)!).root, 'sheetView');
    selected.push(views.some((el) => el.attrs.get('tabSelected') === '1'));
  }
  return { activeTab: view ? view.attrs.get('activeTab') : undefined, selected };
}

interface BookOptions {
  /** Inner XML before <sheets>; by default one window on the first sheet. */
  head?: string;
  /** Index of the sheet that says tabSelected. */
  selected?: number;
  hidden?: number;
}

/** Three sheets A, B, C with a number each; the first is selected. */
const book = (options: BookOptions = {}) =>
  buildXlsx({
    workbookHead: options.head === undefined ? '<bookViews><workbookView xWindow="10" firstSheet="0"/></bookViews>' : options.head,
    sheets: ['A', 'B', 'C'].map((name, index) => ({
      name,
      state: index === options.hidden ? ('hidden' as const) : undefined,
      xml: worksheetXml(
        `<sheetViews><sheetView${index === (options.selected ?? 0) ? ' tabSelected="1"' : ''} workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>` +
          `<sheetData><row r="1">${num('A1', index + 1)}</row></sheetData>`,
      ),
    })),
  });

test('active sheet: showing another sheet and nothing else leaves the file byte for byte', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 2;
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('active sheet: with a cell edited, the file opens on the sheet of the model and only the two sheets concerned say so', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 2;
  // The edit is on the sheet in the middle: it is neither the old nor the new active sheet.
  type(doc, 1, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '2', selected: [false, false, true] });
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml']);
  // Whatever else the views say is kept.
  const parts = await content(saved);
  const view = all(parseXml(parts.get('xl/workbook.xml')!).root, 'workbookView')[0];
  assert.deepEqual([...view.attrs].sort(), [['activeTab', '2'], ['firstSheet', '0'], ['xWindow', '10']]);
  assert.ok(decoder.decode(parts.get('xl/worksheets/sheet1.xml')).includes('<sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView>'));
  assert.ok(decoder.decode(parts.get('xl/worksheets/sheet3.xml')).includes('<sheetView workbookViewId="0" tabSelected="1"><selection activeCell="B2" sqref="B2"/></sheetView>'));
  const again = await (await open(saved)).workbook();
  assert.equal(again.activeSheet, 2);
  assert.deepEqual(
    again.sheets.map((sheet) => sheet.view.tabSelected),
    [false, false, true],
  );
  assert.equal(again.sheets[1].rows.get(1)!.cells[0].value, 5);
});

test('active sheet: the sheet that is edited is the one shown; an unedited sheet is rewritten for its tab alone', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 1;
  type(doc, 1, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '1', selected: [false, true, false] });
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
});

test('active sheet: back on the sheet of the file, a second save says what the file said', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 2;
  type(doc, 1, 'A2', '5');
  await save(doc);
  doc.wb.activeSheet = 0;
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: undefined, selected: [true, false, false] });
  // The sheets that were rewritten for their tab are the file's again.
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/workbook.xml', 'xl/worksheets/sheet2.xml']);
});

test('active sheet: a file without bookViews gets one when it opens on another sheet', async () => {
  const doc = await load(await book({ head: '' }));
  doc.wb.activeSheet = 1;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '1', selected: [false, true, false] });
  assert.equal((await (await open(saved)).workbook()).activeSheet, 1);
});

test('active sheet: an edit without another sheet shown leaves views and tabs alone', async () => {
  // The file is not consistent: it opens on B and says that C is selected. Nobody asked for that to change.
  const doc = await load(await book({ head: '<bookViews><workbookView activeTab="1"/></bookViews>', selected: 2 }));
  assert.equal(doc.wb.activeSheet, 1);
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '1', selected: [false, false, true] });
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
});

test('active sheet: a hidden sheet is never the one a file opens on', async () => {
  const doc = await load(await book({ hidden: 1 }));
  doc.wb.activeSheet = 1;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  // The first visible sheet, which is the one the file opened on: nothing to change.
  assert.deepEqual(await selection(saved), { activeTab: undefined, selected: [true, false, false] });
});

test('active sheet: styles alone are a change that takes the active sheet along', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 1;
  doc.wb.styles.modified = true;
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '1', selected: [false, true, false] });
  const text = decoder.decode((await content(saved)).get('xl/workbook.xml'));
  // No cell changed: nothing asks for a recalculation.
  assert.ok(!text.includes('fullCalcOnLoad'));
});

test('active sheet: with a sheet renamed, the file opens on the sheet of the model', async () => {
  const doc = await load(await book());
  apply(doc, renameSheet(doc.wb, 1, 'Middle'));
  doc.wb.activeSheet = 1;
  const saved = await save(doc);
  assert.deepEqual(await selection(saved), { activeTab: '1', selected: [false, true, false] });
});

test('active sheet: a sheet inserted and taken back leaves the file as it was, whatever sheet is shown', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 2;
  apply(doc, insertSheet(doc.wb, 3, { name: 'New' }));
  doc.history.undo();
  doc.wb.activeSheet = 1;
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('active sheet: a workbook that is in no file is built to open on the active sheet of the model', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 2;
  const built = await buildFromModel(doc.wb);
  const again = await (await open(built)).workbook();
  assert.equal(again.activeSheet, 2);
  // The sheet marked as selected is the active one, whatever the sheets of the model say of themselves.
  assert.deepEqual(
    again.sheets.map((sheet) => sheet.view.tabSelected),
    [false, false, true],
  );
  // A hidden active sheet: the first visible one is the sheet the file opens on, and the one that is marked.
  doc.wb.activeSheet = 0;
  doc.wb.sheets[0].state = 'hidden';
  const hidden = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.equal(hidden.activeSheet, 1);
  assert.deepEqual(
    hidden.sheets.map((sheet) => sheet.view.tabSelected),
    [false, true, false],
  );
});

// --- chart sheets ---------------------------------------------------------------

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const chartXml = (views: string) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr/>${views}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><!-- kept --></chartsheet>`;

/** A, a chart sheet, C. `active` is the sheet the file opens on and the one it marks; `views` replaces the views of the chart sheet. */
const chartBook = (active: number, views?: string) =>
  buildXlsx({
    workbookHead: `<bookViews><workbookView activeTab="${active}"/></bookViews>`,
    sheets: [
      { name: 'A', xml: worksheetXml(`<sheetViews><sheetView${active === 0 ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews><sheetData><row r="1">${num('A1', 1)}</row></sheetData>`) },
      {
        name: 'Chart',
        kind: 'chartsheet' as const,
        xml: chartXml(views !== undefined ? views : `<sheetViews><sheetView${active === 1 ? ' tabSelected="1"' : ''} zoomScale="85" workbookViewId="0"/></sheetViews>`),
      },
      { name: 'C', xml: worksheetXml(`<sheetViews><sheetView${active === 2 ? ' tabSelected="1"' : ''} workbookViewId="0"/></sheetViews><sheetData><row r="1">${num('A1', 3)}</row></sheetData>`) },
    ],
  });

/** activeTab, and for each sheet in order whether a view of its part says tabSelected. */
async function chartSelection(bytes: Uint8Array): Promise<{ activeTab: string | undefined; selected: boolean[] }> {
  const parts = await content(bytes);
  const view = all(parseXml(parts.get('xl/workbook.xml')!).root, 'workbookView')[0];
  const selected = ['xl/worksheets/sheet1.xml', 'xl/chartsheets/sheet2.xml', 'xl/worksheets/sheet3.xml'].map((name) =>
    all(parseXml(parts.get(name)!).root, 'sheetView').some((el) => el.attrs.get('tabSelected') === '1' || el.attrs.get('tabSelected') === 'true'),
  );
  return { activeTab: view ? view.attrs.get('activeTab') : undefined, selected };
}

test('active sheet: from a chart sheet to a worksheet, exactly one sheet is marked, and nothing else of the chart sheet changes', async () => {
  const doc = await load(await chartBook(1));
  assert.equal(doc.wb.activeSheet, 1);
  doc.wb.activeSheet = 2;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await chartSelection(saved), { activeTab: '2', selected: [false, false, true] });
  const before = decoder.decode((await content(doc.bytes)).get('xl/chartsheets/sheet2.xml'));
  const after = decoder.decode((await content(saved)).get('xl/chartsheets/sheet2.xml'));
  assert.equal(after, before.replace(' tabSelected="1"', ''));
  // Back on the chart sheet: its part is the file's again.
  doc.wb.activeSheet = 1;
  const back = await save(doc);
  assert.deepEqual(await chartSelection(back), { activeTab: '1', selected: [false, true, false] });
  assert.deepEqual(await changedParts(doc.bytes, back), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
});

test('active sheet: from a worksheet to a chart sheet, exactly one sheet is marked', async () => {
  const doc = await load(await chartBook(2));
  doc.wb.activeSheet = 1;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await chartSelection(saved), { activeTab: '1', selected: [false, true, false] });
  const before = decoder.decode((await content(doc.bytes)).get('xl/chartsheets/sheet2.xml'));
  const after = decoder.decode((await content(saved)).get('xl/chartsheets/sheet2.xml'));
  // The attribute is added to the view that is there; every other byte is the file's.
  assert.equal(after.replace(' tabSelected="1"', ''), before);
  assert.ok(after.includes('<!-- kept -->'));
  assert.equal((await (await open(saved)).workbook()).activeSheet, 1);
});

test('active sheet: a chart sheet without views, with an empty list of views and with a prefix gets or loses its mark', async () => {
  for (const views of ['', '<sheetViews/>', '<sheetViews></sheetViews>']) {
    const doc = await load(await chartBook(2, views));
    doc.wb.activeSheet = 1;
    type(doc, 0, 'A2', '5');
    const saved = await save(doc);
    assert.deepEqual(await chartSelection(saved), { activeTab: '1', selected: [false, true, false] }, JSON.stringify(views));
    const after = decoder.decode((await content(saved)).get('xl/chartsheets/sheet2.xml'));
    // Where the format has the views: behind the properties of the sheet.
    assert.ok(/<sheetPr\/><sheetViews><sheetView tabSelected="1" workbookViewId="0"\/><\/sheetViews><pageMargins/.test(after), after);
  }
  // Two windows on the chart sheet, written with a prefix: both lose the mark.
  const prefixed = `<?xml version="1.0"?><x:chartsheet xmlns:x="${NS}"><x:sheetViews><x:sheetView tabSelected="1" workbookViewId="0"/><x:sheetView tabSelected="true" workbookViewId="1"/></x:sheetViews></x:chartsheet>`;
  const bytes = await buildXlsx({
    workbookHead: '<bookViews><workbookView activeTab="1"/></bookViews>',
    sheets: [
      { name: 'A', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row></sheetData>`) },
      { name: 'Chart', kind: 'chartsheet' as const, xml: prefixed },
      { name: 'C', xml: worksheetXml('<sheetData/>') },
    ],
  });
  const doc = await load(bytes);
  doc.wb.activeSheet = 0;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await chartSelection(saved), { activeTab: undefined, selected: [true, false, false] });
  assert.equal(
    decoder.decode((await content(saved)).get('xl/chartsheets/sheet2.xml')),
    prefixed.replace(' tabSelected="1"', '').replace(' tabSelected="true"', ''),
  );
});

test('active sheet: with a sheet added, a chart sheet that was the active one loses its mark, and one that is keeps it', async () => {
  const doc = await load(await chartBook(1));
  apply(doc, insertSheet(doc.wb, 3, { name: 'New' }));
  doc.wb.activeSheet = 3;
  const saved = await save(doc);
  assert.deepEqual((await chartSelection(saved)).selected, [false, false, false]);
  assert.equal((await (await open(saved)).workbook()).activeSheet, 3);
  doc.wb.activeSheet = 1;
  const kept = await save(doc);
  assert.deepEqual(await chartSelection(kept), { activeTab: '1', selected: [false, true, false] });
});

// --- several windows, and a hidden active sheet -----------------------------------

test('active sheet: in a file with several windows another sheet shown is not written', async () => {
  const two = (selected: boolean[]) => `<sheetViews>${selected.map((on, id) => `<sheetView${on ? ' tabSelected="1"' : ''} workbookViewId="${id}"/>`).join('')}</sheetViews>`;
  // The first window shows A, the second C.
  const bytes = await buildXlsx({
    workbookHead: '<bookViews><workbookView/><workbookView activeTab="2"/></bookViews>',
    sheets: [
      { name: 'A', xml: worksheetXml(`${two([true, false])}<sheetData><row r="1">${num('A1', 1)}</row></sheetData>`) },
      { name: 'B', xml: worksheetXml(`${two([false, false])}<sheetData/>`) },
      { name: 'C', xml: worksheetXml(`${two([false, true])}<sheetData/>`) },
    ],
  });
  const marks = async (saved: Uint8Array) => {
    const parts = await content(saved);
    return [1, 2, 3].map((n) => all(parseXml(parts.get(`xl/worksheets/sheet${n}.xml`)!).root, 'sheetView').map((el) => el.attrs.get('tabSelected') === '1'));
  };
  const doc = await load(bytes);
  doc.wb.activeSheet = 1;
  type(doc, 0, 'A2', '5');
  const saved = await save(doc);
  assert.deepEqual(await marks(saved), [[true, false], [false, false], [false, true]]);
  assert.deepEqual(
    all(parseXml((await content(saved)).get('xl/workbook.xml')!).root, 'workbookView').map((el) => el.attrs.get('activeTab')),
    [undefined, '2'],
  );
  assert.deepEqual(await changedParts(bytes, saved), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  // A sheet is added: every window opens on the active sheet, and every view of every sheet says the same.
  apply(doc, insertSheet(doc.wb, 3, { name: 'New' }));
  doc.wb.activeSheet = 1;
  const added = await save(doc);
  assert.deepEqual(await marks(added), [[false, false], [true, true], [false, false]]);
  assert.deepEqual(
    all(parseXml((await content(added)).get('xl/workbook.xml')!).root, 'workbookView').map((el) => el.attrs.get('activeTab')),
    ['1', '1'],
  );
});

test('active sheet: a hidden sheet as the active one, with styles alone changed, leaves the workbook part as it is', async () => {
  const doc = await load(await book({ hidden: 1 }));
  doc.wb.activeSheet = 1;
  doc.wb.styles.modified = true;
  const saved = await save(doc);
  assert.ok(sameBytes((await content(saved)).get('xl/workbook.xml'), (await content(doc.bytes)).get('xl/workbook.xml')));
  assert.ok(!(await changedParts(doc.bytes, saved)).some((name) => name.startsWith('xl/worksheets/')));
});
