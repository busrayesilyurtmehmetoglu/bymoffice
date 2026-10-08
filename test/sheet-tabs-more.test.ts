// The rest of what is done to the sheets of a workbook: a sheet moved, hidden
// and shown, copied, and the colour of its tab. Each as an operation with its
// undo, and as what a save makes of it in the package.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { ZipArchive, cellName, open, parseCellRef, parseXml } from '../src/index';
import type { OfficeFile, Sheet, Workbook, XmlElement } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { canCopySheet, canHideSheet, canMoveSheet, copySheet, copySheetLosses, copySheetName, deleteSheet, insertSheet, moveSheet, renameSheet, setSheetState, setTabColor } from '../src/ops/sheets';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { buildXlsx, utf8, worksheetXml } from './fixtures';
import { comment, commentsXml, vmlNote, vmlXml } from './sheet-drawing-fixtures';

const decoder = new TextDecoder();
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

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
const names = (wb: Workbook) => wb.sheets.map((sheet) => sheet.name);
const states = (wb: Workbook) => wb.sheets.map((sheet) => `${sheet.name}:${sheet.state}`);

function same(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

/** Every entry of the archive but the ones named, byte for byte. */
async function sameBut(actual: Uint8Array, expected: Uint8Array, changed: string[], label: string): Promise<void> {
  const a = ZipArchive.read(actual);
  const b = ZipArchive.read(expected);
  assert.deepEqual(a.names().slice().sort(), b.names().slice().sort(), `${label}: the same parts`);
  for (const name of b.names()) {
    if (changed.indexOf(name) >= 0) continue;
    same(await a.read(name), await b.read(name), `${label}: ${name}`);
  }
}

const childrenOf = (el: XmlElement, name: string) => el.children.filter((node): node is XmlElement => node.type === 'element' && node.name === name);

/**
 * What a package of a workbook has to hold: each relationship leads to a part, each part has
 * a content type, each sheet of the list has a part, a name scoped to a sheet names one that is
 * there, the file opens on a visible sheet, and exactly that sheet is marked as selected.
 */
async function checkPackage(bytes: Uint8Array): Promise<void> {
  const zip = ZipArchive.read(bytes);
  const all = new Set(zip.names());
  const types = parseXml(await zip.read('[Content_Types].xml')).root;
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  for (const node of types.children) {
    if (node.type !== 'element') continue;
    if (node.name === 'Default') defaults.add((node.attrs.get('Extension') || '').toLowerCase());
    if (node.name === 'Override') {
      const part = (node.attrs.get('PartName') || '').slice(1);
      assert.ok(all.has(part), `an override names ${part}, which is not in the package`);
      assert.ok(!overrides.has(part), `two overrides for ${part}`);
      overrides.add(part);
    }
  }
  const targets = new Map<string, Map<string, string>>();
  for (const name of all) {
    if (name === '[Content_Types].xml') continue;
    assert.ok(overrides.has(name) || defaults.has(name.slice(name.lastIndexOf('.') + 1).toLowerCase()), `${name} has no content type`);
    if (!/\.rels$/.test(name)) continue;
    const source = name.replace(/_rels\/([^/]*)\.rels$/, '$1');
    const ids = new Map<string, string>();
    for (const rel of parseXml(await zip.read(name)).root.children) {
      if (rel.type !== 'element') continue;
      const id = rel.attrs.get('Id') || '';
      assert.ok(!ids.has(id), `${name}: the id ${id} is used twice`);
      if (rel.attrs.get('TargetMode') === 'External') continue;
      const target = rel.attrs.get('Target') || '';
      const segments: string[] = target.charAt(0) === '/' ? [] : source.split('/').slice(0, -1);
      for (const segment of target.split('/')) {
        if (segment === '..') segments.pop();
        else if (segment !== '' && segment !== '.') segments.push(segment);
      }
      assert.ok(all.has(segments.join('/')), `${name}: ${id} leads to ${segments.join('/')}, which is not in the package`);
      ids.set(id, segments.join('/'));
    }
    targets.set(source, ids);
  }
  const book = parseXml(await zip.read('xl/workbook.xml')).root;
  const sheets = childrenOf(childrenOf(book, 'sheets')[0], 'sheet');
  const rels = targets.get('xl/workbook.xml') || new Map<string, string>();
  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();
  const parts: string[] = [];
  for (const sheet of sheets) {
    const part = rels.get(sheet.attrs.get('r:id') || '');
    assert.ok(part, `sheet ${sheet.attrs.get('name')} has a part`);
    assert.ok(parts.indexOf(part as string) < 0, `${part} is the part of one sheet`);
    parts.push(part as string);
    assert.ok(!sheetIds.has(sheet.attrs.get('sheetId') || ''), 'sheet ids are unique');
    sheetIds.add(sheet.attrs.get('sheetId') || '');
    assert.ok(!sheetNames.has((sheet.attrs.get('name') || '').toLowerCase()), 'sheet names are unique');
    sheetNames.add((sheet.attrs.get('name') || '').toLowerCase());
  }
  const names = childrenOf(book, 'definedNames')[0];
  for (const name of names ? childrenOf(names, 'definedName') : []) {
    const local = name.attrs.get('localSheetId');
    if (local !== undefined) assert.ok(Number(local) < sheets.length, `the name ${name.attrs.get('name')} is local to sheet ${local} of ${sheets.length}`);
  }
  const views = childrenOf(book, 'bookViews')[0];
  const view = views ? childrenOf(views, 'workbookView')[0] : undefined;
  const active = Number((view && view.attrs.get('activeTab')) || '0');
  assert.ok(active < sheets.length, 'the active tab is a sheet');
  assert.ok((sheets[active].attrs.get('state') || 'visible') === 'visible', `the file opens on a visible sheet (${active})`);
  const first = view ? view.attrs.get('firstSheet') : undefined;
  if (first !== undefined) assert.ok(Number(first) < sheets.length, 'the first tab shown is a sheet');
  assert.ok(sheets.some((sheet) => (sheet.attrs.get('state') || 'visible') === 'visible'), 'a sheet is visible');
  // Where sheets say which of them is selected, it is the active one.
  const selected: number[] = [];
  for (let i = 0; i < parts.length; i++) if (/tabSelected="(1|true)"/.test(decoder.decode(await zip.read(parts[i])))) selected.push(i);
  if (selected.length > 0) assert.deepEqual(selected, [active], 'the selected sheet is the active one');
}

const row = (r: number, cells: string) => `<row r="${r}">${cells}</row>`;
const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const formula = (ref: string, f: string, value = 0) => `<c r="${ref}"><f>${f}</f><v>${value}</v></c>`;

const APP =
  `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application>` +
  '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
  '<TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Bir</vt:lpstr><vt:lpstr>İki</vt:lpstr><vt:lpstr>Üç</vt:lpstr></vt:vector></TitlesOfParts></Properties>';

/** Three sheets: "Bir" (formulas, a print area, a note, a tab colour by theme), "İki" (hidden), "Üç" (a filter name). */
async function three(): Promise<Uint8Array> {
  const bytes = await buildXlsx({
    workbookHead: '<bookViews><workbookView xWindow="0" yWindow="0" firstSheet="0" activeTab="0"/></bookViews>',
    workbookTail:
      "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Bir!$A$1:$B$2</definedName><definedName name=\"Hepsi\">Bir!$A$1+'Üç'!$A$1</definedName><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"2\">'Üç'!$1:$1</definedName></definedNames>",
    sheets: [
      {
        name: 'Bir',
        xml: worksheetXml(
          '<sheetPr codeName="Sayfa1"><tabColor theme="4" tint="-0.249977111117893"/><pageSetUpPr fitToPage="1"/></sheetPr>' +
            '<sheetViews><sheetView tabSelected="1" zoomScale="120" workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="18"/><cols><col min="2" max="2" width="20" customWidth="1"/></cols>' +
            `<sheetData>${row(1, num('A1', 1) + formula('B1', 'A1*2', 2))}${row(2, formula('A2', "Bir!A1+'Üç'!A1+SUM(Bir:Üç!A1)", 12) + formula('B2', 'İki!A1', 5))}</sheetData>` +
            '<mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells><hyperlinks><hyperlink ref="A1" location="Bir!B2"/></hyperlinks>' +
            '<pageMargins left="0.5" right="0.5" top="1" bottom="1" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape"/><legacyDrawing r:id="rIdV"/>',
        ),
      },
      { name: 'İki', state: 'hidden', xml: worksheetXml(`<sheetData>${row(1, num('A1', 5))}</sheetData>`) },
      { name: 'Üç', xml: worksheetXml(`<sheetData>${row(1, num('A1', 10) + formula('B1', 'Bir!A1', 1))}</sheetData>`) },
    ],
  });
  const zip = ZipArchive.read(bytes);
  zip.set('xl/comments1.xml', utf8(commentsXml(['Ada'], comment('B2', 0, '<t>not</t>'))));
  zip.set('xl/drawings/vmlDrawing1.vml', utf8(vmlXml(vmlNote(1, 1))));
  zip.set(
    'xl/worksheets/_rels/sheet1.xml.rels',
    utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rIdC" Type="${NS_REL}/comments" Target="../comments1.xml"/><Relationship Id="rIdV" Type="${NS_REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/></Relationships>`),
  );
  zip.set('docProps/app.xml', utf8(APP));
  const types = decoder.decode(await zip.read('[Content_Types].xml'));
  zip.set(
    '[Content_Types].xml',
    utf8(
      types.replace(
        '</Types>',
        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',
      ),
    ),
  );
  const rels = decoder.decode(await zip.read('_rels/.rels'));
  zip.set('_rels/.rels', utf8(rels.replace('</Relationships>', `<Relationship Id="rIdApp" Type="${NS_REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`)));
  return zip.write();
}

const formulaOf = (sheet: Sheet, ref: string): string | undefined => {
  const { row: r, col } = parseCellRef(ref)!;
  return sheet.rows.get(r)?.cells[col]?.formula?.text;
};
const scoped = (wb: Workbook) => wb.definedNames.map((name) => `${name.name}@${name.localSheet === undefined ? '-' : wb.sheets[name.localSheet].name}=${name.value}`).sort();
const shown = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => `${cellName(placed.row, placed.col)} by ${placed.comment.author ?? ''}: ${placed.comment.text}`)
    .sort();

/** The model as far as these operations change it, to compare what was saved with what is in memory. */
const digest = (wb: Workbook) => ({
  sheets: wb.sheets.map((sheet) => ({
    name: sheet.name,
    state: sheet.state,
    tabColor: sheet.tabColor,
    cells: [...sheet.rows.entries()]
      .sort((a, b) => a[0] - b[0])
      .map(([index, data]) => `${index}:${data.cells.map((cell, col) => `${col}=${cell.formula ? `=${cell.formula.text}` : String(cell.value)}`).filter(Boolean).join(',')}`),
    notes: shown(sheet),
    merges: sheet.merges,
    frozen: [sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.zoomScale],
    columns: sheet.columns.map((range) => `${range.min}-${range.max}:${range.width}`),
    links: sheet.hyperlinks.map((link) => link.location),
  })),
  names: scoped(wb),
  active: wb.sheets[wb.activeSheet] ? wb.sheets[wb.activeSheet].name : undefined,
});

// --- B1: moving ------------------------------------------------------------------------------

test('sheets are moved: the order changes, local names and the active sheet follow their sheets, formulas stay as they are', async () => {
  const original = await three();
  const doc = await load(original);
  const wb = doc.wb;
  assert.equal(canMoveSheet(wb, 0, 2), undefined);
  assert.deepEqual(moveSheet(wb, 1, 1), []);
  assert.throws(() => moveSheet(wb, 0, 3), /^RangeError: Cannot move a sheet to index 3$/);
  assert.throws(() => moveSheet(wb, 5, 0), /^RangeError: No sheet at index 5$/);
  const before = scoped(wb);
  apply(doc, moveSheet(wb, 0, 2));
  assert.deepEqual(names(wb), ['İki', 'Üç', 'Bir']);
  assert.equal(wb.activeSheet, 2);
  assert.equal(wb.sheetsModified, true);
  assert.deepEqual(scoped(wb), before);
  // A 3D reference is text: it reads the sheets between its ends, whichever those are now.
  assert.equal(formulaOf(wb.sheets[2], 'A2'), "Bir!A1+'Üç'!A1+SUM(Bir:Üç!A1)");

  const saved = await save(doc);
  await checkPackage(saved);
  const again = await reopen(saved);
  assert.deepEqual(digest(again), digest(wb));
  // Parts keep their names and sheets their ids; nothing but the workbook part, the list of titles and the selected marks changed.
  const book = await text(saved, 'xl/workbook.xml');
  assert.match(book, /<sheet name="İki" sheetId="2" state="hidden" r:id="rId2"\/><sheet name="Üç" sheetId="3" r:id="rId3"\/><sheet name="Bir" sheetId="1" r:id="rId1"\/>/);
  assert.match(book, /activeTab="2"/);
  assert.match(book, /name="_xlnm.Print_Area" localSheetId="2"/);
  assert.match(book, /name="_xlnm.Print_Titles" localSheetId="1"/);
  assert.match(await text(saved, 'docProps/app.xml'), /<vt:lpstr>İki<\/vt:lpstr><vt:lpstr>Üç<\/vt:lpstr><vt:lpstr>Bir<\/vt:lpstr>/);
  await sameBut(saved, original, ['xl/workbook.xml', 'docProps/app.xml'], 'moved');

  doc.history.undo();
  assert.deepEqual(names(wb), ['Bir', 'İki', 'Üç']);
  assert.equal(wb.activeSheet, 0);
  same(await save(doc), original, 'undone');
  doc.history.redo();
  same(await save(doc), saved, 'done again');
});

// --- B2: hiding and showing --------------------------------------------------------------------

test('sheets are hidden and shown: the last visible sheet stays, the active sheet moves on, the state is written', async () => {
  const original = await three();
  const doc = await load(original);
  const wb = doc.wb;
  assert.deepEqual(setSheetState(wb, 1, 'hidden'), []);
  assert.equal(canHideSheet(wb, 0), undefined);
  apply(doc, setSheetState(wb, 0, 'hidden'));
  assert.deepEqual(states(wb), ['Bir:hidden', 'İki:hidden', 'Üç:visible']);
  assert.equal(wb.activeSheet, 2);
  assert.equal(canHideSheet(wb, 2), 'Cannot hide the last visible sheet');
  assert.throws(() => setSheetState(wb, 2, 'hidden'), /^RangeError: Cannot hide the last visible sheet$/);
  assert.throws(() => setSheetState(wb, 2, 'gone' as never), /^RangeError: A sheet is visible, hidden or veryHidden$/);
  apply(doc, setSheetState(wb, 1, 'visible'));
  assert.equal(wb.activeSheet, 2);

  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(digest(await reopen(saved)), digest(wb));
  const book = await text(saved, 'xl/workbook.xml');
  assert.match(book, /<sheet name="Bir" sheetId="1" r:id="rId1" state="hidden"\/><sheet name="İki" sheetId="2" r:id="rId2"\/>/);
  assert.match(book, /activeTab="2"/);
  await sameBut(saved, original, ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml'], 'hidden');
  assert.doesNotMatch(await text(saved, 'xl/worksheets/sheet1.xml'), /tabSelected/);
  assert.match(await text(saved, 'xl/worksheets/sheet3.xml'), /tabSelected="1"/);

  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(states(wb), ['Bir:visible', 'İki:hidden', 'Üç:visible']);
  assert.equal(wb.activeSheet, 0);
  same(await save(doc), original, 'undone');

  // A sheet that is very hidden is kept so, and can be shown.
  wb.sheets[1].state = 'veryHidden';
  apply(doc, setSheetState(wb, 1, 'visible'));
  assert.equal(wb.sheets[1].state, 'visible');
  doc.history.undo();
  assert.equal(wb.sheets[1].state, 'veryHidden');
});

// --- B4: the colour of a tab ---------------------------------------------------------------------

test('tabs are coloured: only <tabColor> of the sheet part changes, a colour of the file stays as it is until it is changed', async () => {
  const original = await three();
  const doc = await load(original);
  const wb = doc.wb;
  assert.deepEqual(wb.sheets[0].tabColor, { theme: 4, tint: -0.249977111117893 });
  assert.throws(() => setTabColor(wb, 0, 'red'), /^RangeError: A tab colour is written #RRGGBB$/);
  assert.deepEqual(setTabColor(wb, 2, undefined), []);
  // Another sheet gets a colour: the sheet that has one is not written.
  apply(doc, setTabColor(wb, 2, '#ff8800'));
  assert.deepEqual(wb.sheets[2].tabColor, { rgb: 'FFFF8800' });
  let saved = await save(doc);
  await checkPackage(saved);
  await sameBut(saved, original, ['xl/worksheets/sheet3.xml'], 'coloured');
  assert.equal(await text(saved, 'xl/worksheets/sheet3.xml'), (await text(original, 'xl/worksheets/sheet3.xml')).replace('<sheetData>', '<sheetPr><tabColor rgb="FFFF8800"/></sheetPr><sheetData>'));
  assert.deepEqual(digest(await reopen(saved)), digest(wb));

  // The colour of the file replaced, then removed: the other children and attributes of <sheetPr> stay.
  apply(doc, setTabColor(wb, 0, '#00AA00'));
  saved = await save(doc);
  const sheet1 = await text(original, 'xl/worksheets/sheet1.xml');
  assert.equal(await text(saved, 'xl/worksheets/sheet1.xml'), sheet1.replace('<tabColor theme="4" tint="-0.249977111117893"/>', '<tabColor rgb="FF00AA00"/>'));
  apply(doc, setTabColor(wb, 0, undefined));
  saved = await save(doc);
  assert.equal(await text(saved, 'xl/worksheets/sheet1.xml'), sheet1.replace('<tabColor theme="4" tint="-0.249977111117893"/>', ''));
  assert.equal((await reopen(saved)).sheets[0].tabColor, undefined);
  await checkPackage(saved);

  while (doc.history.canUndo) doc.history.undo();
  assert.deepEqual(wb.sheets[0].tabColor, { theme: 4, tint: -0.249977111117893 });
  same(await save(doc), original, 'undone');

  // Together with a change of cells: the sheet is written from the model, with its colour.
  apply(doc, setCellInput(wb, 2, 4, 0, 'x'));
  apply(doc, setTabColor(wb, 2, '#112233'));
  saved = await save(doc);
  assert.match(await text(saved, 'xl/worksheets/sheet3.xml'), /<worksheet [^>]*><sheetPr><tabColor rgb="FF112233"\/><\/sheetPr>/);
  assert.deepEqual(digest(await reopen(saved)), digest(wb));
});

// --- B3: copying ------------------------------------------------------------------------------------

test('sheets are copied: cells, formats, sizes, merged cells, links, the view, the page, notes and local names come along under a new name', async () => {
  const original = await three();
  const doc = await load(original);
  const wb = doc.wb;
  assert.equal(copySheetName(wb, 0), 'Bir (2)');
  assert.equal(canCopySheet(wb, 0), undefined);
  apply(doc, copySheet(wb, 0));
  assert.deepEqual(names(wb), ['Bir', 'Bir (2)', 'İki', 'Üç']);
  const copy = wb.sheets[1];
  assert.equal(copy.part, '');
  // References without a sheet name stay; those that name the copied sheet name the copy.
  assert.equal(formulaOf(copy, 'B1'), 'A1*2');
  assert.equal(formulaOf(copy, 'A2'), "'Bir (2)'!A1+'Üç'!A1+SUM('Bir (2):Üç'!A1)");
  assert.equal(formulaOf(wb.sheets[0], 'A2'), "Bir!A1+'Üç'!A1+SUM(Bir:Üç!A1)");
  assert.deepEqual(copy.hyperlinks.map((link) => link.location), ["'Bir (2)'!B2"]);
  assert.deepEqual(copy.merges, [{ top: 0, left: 3, bottom: 0, right: 4 }]);
  assert.deepEqual([copy.view.frozenRows, copy.view.frozenCols, copy.view.zoomScale], [1, 1, 120]);
  assert.deepEqual(copy.tabColor, { theme: 4, tint: -0.249977111117893 });
  assert.deepEqual(shown(copy), ['B2 by Ada: not']);
  assert.deepEqual(scoped(wb), ["Hepsi@-=Bir!$A$1+'Üç'!$A$1", "_xlnm.Print_Area@Bir (2)='Bir (2)'!$A$1:$B$2", '_xlnm.Print_Area@Bir=Bir!$A$1:$B$2', "_xlnm.Print_Titles@Üç='Üç'!$1:$1"]);
  assert.equal(copy.pageSetup!.page!.orientation, 'landscape');
  // The code name of the sheet (<sheetPr codeName>) is its own: the copy has none, and that is said.
  assert.deepEqual(copySheetLosses(wb, 0), [{ what: 'codeName', count: 1 }]);

  const saved = await save(doc);
  await checkPackage(saved);
  const again = await reopen(saved);
  assert.deepEqual(digest(again), digest(wb));
  assert.equal(again.sheets[1].pageSetup!.page!.orientation, 'landscape');
  assert.equal(again.sheets[1].pageSetup!.margins!.left, 0.5);
  assert.equal(again.sheets[1].pageSetup!.fitToPage, true);
  assert.equal(again.sheets[1].defaultRowHeight, 18);
  // The sheets of the file are as they were, but for the mark of the selected one.
  const zip = ZipArchive.read(saved);
  for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml']) {
    same(await zip.read(name), await ZipArchive.read(original).read(name), name);
  }
  // A file made from the model alone holds the same.
  assert.deepEqual(digest(await (await open(await buildFromModel(wb))).workbook()), digest(wb));

  doc.history.undo();
  assert.deepEqual(names(wb), ['Bir', 'İki', 'Üç']);
  same(await save(doc), original, 'undone');
  doc.history.redo();
  same(await save(doc), saved, 'done again');

  // The next copy, a copy at a place, a copy under a name, and names at the limit of 31 characters.
  assert.equal(copySheetName(wb, 0), 'Bir (3)');
  apply(doc, copySheet(wb, 3, { at: 0, name: 'Başa' }));
  assert.deepEqual(names(wb), ['Başa', 'Bir', 'Bir (2)', 'İki', 'Üç']);
  assert.equal(wb.activeSheet, 1);
  assert.throws(() => copySheet(wb, 1, { name: 'İki' }), /another sheet has that name/);
  apply(doc, renameSheet(wb, 4, 'a'.repeat(31)));
  assert.equal(copySheetName(wb, 4), `${'a'.repeat(27)} (2)`);
  await checkPackage(await save(doc));
});

test('sheets are copied: rules go with the copy; what is kept as the file has it is not copied, and is counted', async () => {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'Veri',
        xml: worksheetXml(
          `<sheetData>${row(1, num('A1', 1))}</sheetData>` +
            '<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>' +
            '<dataValidations count="1"><dataValidation type="whole" sqref="B1"><formula1>1</formula1><formula2>9</formula2></dataValidation></dataValidations>' +
            '<hyperlinks><hyperlink ref="A1" r:id="rIdL"/></hyperlinks>',
        ),
        links: { rIdL: 'https://example.com/' },
      },
      { name: 'Grafik', kind: 'chartsheet', xml: `${XML}<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>` },
    ],
  });
  const doc = await load(bytes);
  assert.deepEqual(copySheetLosses(doc.wb, 0), [
    { what: 'externalLinks', count: 1 },
  ]);
  assert.match(canCopySheet(doc.wb, 1) || '', /^Cannot copy a chartsheet: not supported yet$/);
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  await checkPackage(saved);
  const again = await reopen(saved);
  assert.deepEqual(names(again), ['Veri', 'Veri (2)', 'Grafik']);
  // Conditional formats and data validations are copied: the copy has the rules the sheet has.
  assert.equal(again.sheets[0].conditionalFormats!.length, 1);
  assert.deepEqual(again.sheets[1].conditionalFormats!.map((format) => [format.ranges, format.rules.map((rule) => [rule.type, rule.operator, rule.formulas])]), [[[{ top: 0, left: 0, bottom: 8, right: 0 }], [['cellIs', 'greaterThan', ['0']]]]]);
  assert.deepEqual(again.sheets[1].dataValidations!.map((rule) => [rule.type, rule.formula1, rule.formula2, rule.ranges]), [['whole', '1', '9', [{ top: 0, left: 1, bottom: 0, right: 1 }]]]);
});

// --- B5: together ------------------------------------------------------------------------------------

test('sheets: a workbook whose sheet names differ only by case refuses each of these, and so does an index that is none', async () => {
  const doc = await load(await buildXlsx({ sheets: [{ name: 'IL', xml: worksheetXml('<sheetData/>') }, { name: 'İL', xml: worksheetXml('<sheetData/>') }] }));
  const wb = doc.wb;
  const clash = /differ only by case/;
  assert.throws(() => moveSheet(wb, 0, 1), clash);
  assert.throws(() => setSheetState(wb, 0, 'hidden'), clash);
  assert.throws(() => copySheet(wb, 0), clash);
  assert.throws(() => setTabColor(wb, 0, '#000000'), clash);
  for (const run of [() => moveSheet(wb, 9, 0), () => setSheetState(wb, 9, 'hidden'), () => copySheet(wb, 9), () => setTabColor(wb, 9, undefined)]) assert.throws(run, /No sheet at index 9/);
});

function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

test('sheets: after any mix of moving, hiding, copying, colouring, inserting, deleting, renaming and notes, what is saved is what is in memory', async () => {
  let saves = 0;
  const done = new Map<string, number>();
  for (let seed = 1; seed <= 24; seed++) {
    const original = await three();
    const doc = await load(original);
    const wb = doc.wb;
    const next = random(seed * 15485863);
    const pick = (n: number) => Math.floor(next() * n);
    const trail: string[] = [];
    const check = async (stage: string) => {
      const where = `seed ${seed} ${stage}: ${trail.join('; ')}`;
      const saved = await save(doc);
      saves++;
      await checkPackage(saved);
      assert.deepEqual(digest(await reopen(saved)), digest(wb), where);
      assert.deepEqual(digest(await (await open(await buildFromModel(wb))).workbook()).sheets, digest(wb).sheets, `${where} (built)`);
    };
    for (let step = 0; step < 14; step++) {
      const roll = next();
      const count = wb.sheets.length;
      const at = pick(count);
      let what = '';
      try {
        if (roll < 0.2) {
          what = 'move';
          const to = pick(count);
          trail.push(`move ${at}>${to}`);
          apply(doc, moveSheet(wb, at, to));
        } else if (roll < 0.35) {
          what = 'state';
          const state = wb.sheets[at].state === 'visible' ? 'hidden' : 'visible';
          trail.push(`${state} ${at}`);
          apply(doc, setSheetState(wb, at, state));
        } else if (roll < 0.47) {
          what = 'copy';
          trail.push(`copy ${at}`);
          apply(doc, copySheet(wb, at, next() < 0.5 ? {} : { at: pick(count + 1) }));
        } else if (roll < 0.57) {
          what = 'colour';
          const color = next() < 0.3 ? undefined : `#${(0x100000 + pick(0xefffff)).toString(16)}`;
          trail.push(`colour ${at} ${color}`);
          apply(doc, setTabColor(wb, at, color));
        } else if (roll < 0.65) {
          what = 'insert';
          trail.push(`insert ${at}`);
          apply(doc, insertSheet(wb, at));
        } else if (roll < 0.72) {
          what = 'delete';
          trail.push(`delete ${at}`);
          apply(doc, deleteSheet(wb, at));
        } else if (roll < 0.78) {
          what = 'rename';
          trail.push(`rename ${at}`);
          apply(doc, renameSheet(wb, at, `Ad ${seed}.${step}`));
        } else if (roll < 0.86) {
          what = 'note';
          trail.push(`note ${at}`);
          apply(doc, setComment(wb, at, pick(4), pick(3), next() < 0.3 ? undefined : `n${step}`));
        } else if (roll < 0.95 && doc.history.canUndo) {
          what = 'undo';
          trail.push('undo');
          doc.history.undo();
        } else if (doc.history.canRedo) {
          what = 'redo';
          trail.push('redo');
          doc.history.redo();
        }
        if (what) done.set(what, (done.get(what) || 0) + 1);
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
        trail.push('(refused)');
      }
      if (step % 3 === 2) await check(`step ${step}`);
    }
    await check('end');
    while (doc.history.canUndo) doc.history.undo();
    const undone = await save(doc);
    // A sheet whose cells were rewritten (a formula named a sheet that was renamed or deleted) is written from
    // the model for good, and the workbook with it; everything else is the file again.
    if (wb.sheets.some((sheet) => sheet.modified)) await sameBut(undone, original, ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/_rels/workbook.xml.rels', '[Content_Types].xml'], `seed ${seed}: everything undone: ${trail.join('; ')}`);
    else same(undone, original, `seed ${seed}: everything undone: ${trail.join('; ')}`);
  }
  assert.ok(saves >= 110, `${saves} saves`);
  for (const what of ['move', 'state', 'copy', 'colour', 'insert', 'delete', 'rename', 'note', 'undo']) assert.ok((done.get(what) || 0) >= 5, `${what}: ${done.get(what)}`);
});
