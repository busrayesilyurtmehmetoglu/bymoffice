// Notes written into the package: the first note of a sheet, a note edited
// or deleted beside notes that stay as they are, a sheet with form controls
// or with threaded comments, undo back to the file, sheets that are not in
// the file, and a random mix with rows and columns that move.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { writeXlsx } from '../src/formats/xlsx/write';
import { MAX_COLS, MAX_ROWS, cellName, open, parseCellRef, parseXml, ZipArchive } from '../src/index';
import type { OfficeFile, Sheet, Workbook, XmlElement } from '../src/index';
import { setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { buildXlsx, utf8, worksheetXml } from './fixtures';
import { comment, commentsXml, drawingWorkbook as bareWorkbook, vmlNote, vmlXml, type DrawingBookSpec } from './sheet-drawing-fixtures';

const decoder = new TextDecoder();
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The fixture with the content type of its legacy drawing declared, as every file an application writes has it. */
async function drawingWorkbook(spec: DrawingBookSpec): Promise<Uint8Array> {
  const zip = ZipArchive.read(await bareWorkbook(spec));
  const types = decoder.decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', utf8(types.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')));
  return zip.write();
}

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
const names = (bytes: Uint8Array) => ZipArchive.read(bytes).names().sort();
const text = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));

function same(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

/**
 * The file as it was, as far as notes go: the same parts, each with its bytes. The workbook
 * part and the sheets are left out: once rows moved or cells changed they are written again
 * for good (the application is asked to calculate), whatever is undone.
 */
async function sameAround(actual: Uint8Array, expected: Uint8Array, label: string): Promise<void> {
  assert.deepEqual(names(actual), names(expected), label);
  for (const name of names(expected)) {
    if (name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
    same(await ZipArchive.read(actual).read(name), await ZipArchive.read(expected).read(name), `${label}: ${name}`);
  }
  for (const name of names(expected)) {
    if (!/^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
    const drawing = (xml: string) => /<legacyDrawing [^>]*>/.exec(xml)?.[0];
    assert.equal(drawing(await text(actual, name)), drawing(await text(expected, name)), `${label}: <legacyDrawing> of ${name}`);
  }
}

const note = (doc: Loaded, sheet: number, ref: string, value: string | undefined, author?: string) => {
  const { row, col } = parseCellRef(ref)!;
  return doc.history.apply(setComment(doc.wb, sheet, row, col, value, author === undefined ? {} : { author }), undefined, undefined);
};

const apply = (doc: Loaded, op: Op) => doc.history.apply([op], undefined, undefined);

/** The comments of a sheet as they are shown: "B2 by Ada: text", sorted. */
const shown = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => `${cellName(placed.row, placed.col)} by ${placed.comment.author ?? ''}: ${placed.comment.text}`)
    .sort();

const WORKSHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'autoFilter', 'mergeCells', 'hyperlinks', 'pageMargins', 'pageSetup', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'tableParts', 'extLst'];

/**
 * What every package has to hold: each relationship leads to a part, each part has a content
 * type, each override names a part, and a sheet's <legacyDrawing> names a legacy drawing and
 * stands where the schema has it.
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
  for (const name of all) {
    if (name === '[Content_Types].xml') continue;
    assert.ok(overrides.has(name) || defaults.has(name.slice(name.lastIndexOf('.') + 1).toLowerCase()), `${name} has no content type`);
    if (!/\.rels$/.test(name)) continue;
    const source = name.replace(/_rels\/([^/]*)\.rels$/, '$1');
    const ids = new Set<string>();
    for (const rel of parseXml(await zip.read(name)).root.children) {
      if (rel.type !== 'element') continue;
      const id = rel.attrs.get('Id') || '';
      assert.ok(!ids.has(id), `${name}: the id ${id} is used twice`);
      ids.add(id);
      if (rel.attrs.get('TargetMode') === 'External') continue;
      const target = rel.attrs.get('Target') || '';
      const segments: string[] = target.charAt(0) === '/' ? [] : source.split('/').slice(0, -1);
      for (const segment of target.split('/')) {
        if (segment === '..') segments.pop();
        else if (segment !== '' && segment !== '.') segments.push(segment);
      }
      assert.ok(all.has(segments.join('/')), `${name}: ${id} leads to ${segments.join('/')}, which is not in the package`);
    }
  }
  for (const name of all) {
    if (!/^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
    const root = parseXml(await zip.read(name)).root;
    let last = -1;
    for (const node of root.children) {
      if (node.type !== 'element') continue;
      const rank = WORKSHEET_ORDER.indexOf(node.name);
      if (rank < 0) continue;
      assert.ok(rank >= last, `${name}: <${node.name}> stands behind an element that comes after it`);
      last = rank;
      if (node.name !== 'legacyDrawing') continue;
      const rels = new TextDecoder().decode(await zip.read(name.replace(/([^/]*)$/, '_rels/$1.rels')));
      const id = node.attrs.get('r:id');
      assert.match(rels, new RegExp(`<Relationship Id="${id}" Type="[^"]*/vmlDrawing"`), `${name}: <legacyDrawing> names no legacy drawing`);
    }
  }
}

/** Shape ids of a legacy drawing, and the blocks its <o:idmap> names. */
function shapeIds(vml: string): { ids: number[]; blocks: number[] } {
  const ids = [...vml.matchAll(/<v:shape [^>]*id="_x0000_s(\d+)"/g)].map((match) => Number(match[1]));
  const map = /<o:idmap [^>]*data="([^"]*)"/.exec(vml);
  return { ids, blocks: map ? map[1].split(',').map(Number) : [] };
}

function checkIds(vml: string): void {
  const { ids, blocks } = shapeIds(vml);
  assert.equal(new Set(ids).size, ids.length, `shape ids are unique: ${ids.join(' ')}`);
  for (const id of ids) assert.ok(blocks.indexOf(Math.floor(id / 1024)) >= 0, `shape id ${id} is in a block of the drawing (${blocks.join(',')})`);
}

const plainSheet = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Bir',
        xml: worksheetXml(
          '<sheetData><row r="2"><c r="B2"><v>1</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><extLst><ext uri="{X}"/></extLst>',
        ),
      },
      { name: 'İki', xml: worksheetXml('<sheetData/>') },
    ],
  });

// --- (a) the first note ----------------------------------------------------------------

test('notes are written: the first note of a sheet makes the comments part, the legacy drawing, their relationships and <legacyDrawing>', async () => {
  const original = await plainSheet();
  const doc = await load(original);
  note(doc, 0, 'B2', 'İlk not', 'Ada');
  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(
    names(saved).filter((name) => names(original).indexOf(name) < 0),
    ['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml', 'xl/worksheets/_rels/sheet1.xml.rels'],
  );
  const sheetXml = await text(saved, 'xl/worksheets/sheet1.xml');
  assert.match(sheetXml, /<pageMargins [^>]*\/><legacyDrawing r:id="rId2"\/><extLst>/);
  const rels = await text(saved, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.ok(rels.includes(`<Relationship Id="rId1" Type="${NS_REL}/comments" Target="../comments1.xml"/>`), rels);
  assert.ok(rels.includes(`<Relationship Id="rId2" Type="${NS_REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>`), rels);
  const types = await text(saved, '[Content_Types].xml');
  assert.ok(types.includes('PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"'));
  assert.ok(types.includes('PartName="/xl/drawings/vmlDrawing1.vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"'));
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.includes('<authors><author>Ada</author></authors>'), comments);
  assert.ok(comments.includes('<comment ref="B2" authorId="0" shapeId="0"><text><t>İlk not</t></text></comment>'), comments);
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  checkIds(vml);
  assert.match(vml, /<x:Anchor>2, 15, 0, 10, 4, 15, 4, 4<\/x:Anchor>/);
  assert.match(vml, /<x:Row>1<\/x:Row><x:Column>1<\/x:Column>/);
  assert.ok(vml.includes('<v:shapetype id="_x0000_t202"'));

  const again = await reopen(saved);
  assert.deepEqual(again.sheets[0].comments, [{ row: 1, col: 1, text: 'İlk not', author: 'Ada', box: again.sheets[0].comments![0].box }]);
  assert.ok(again.sheets[0].comments![0].box, 'the box is read back');
  assert.equal(again.sheets[1].comments, undefined);
  assert.deepEqual(again.sheets[0].anchored, ['comments', 'vmlDrawing']);

  // A second save writes the same file; a second note goes into the same parts.
  same(await save(doc), saved, 'saved twice');
  note(doc, 0, 'D7', 'İkinci', 'Grace');
  const second = await save(doc);
  await checkPackage(second);
  assert.deepEqual(names(second), names(saved));
  assert.deepEqual(shown((await reopen(second)).sheets[0]), ['B2 by Ada: İlk not', 'D7 by Grace: İkinci']);
  checkIds(await text(second, 'xl/drawings/vmlDrawing1.vml'));
});

// --- (b) notes beside notes --------------------------------------------------------------

const RICH = '<comment ref="A1" authorId="0" shapeId="0" xr:uid="{11111111-2222-3333-4444-555555555555}"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Ada:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\nzengin  </t></r></text></comment>';
const SECOND = '<comment ref="B2" authorId="1"><text><t>ikinci</t></text><commentPr locked="0"><anchor/></commentPr></comment>';
const THIRD = '<comment  ref="C3"   authorId="0" future="1"><text><t>üçüncü</t></text></comment>';
const boxOf = (ref: string, id: number, extra = '') => {
  const { row, col } = parseCellRef(ref)!;
  return (
    `<v:shape id="_x0000_s${id}" type="#_x0000_t202" style='position:absolute;margin-left:59.25pt;z-index:${id};visibility:hidden' fillcolor="infoBackground [80]" o:insetmode="auto"${extra}>\r\n` +
    '  <v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>\r\n' +
    `  <x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${col + 1}, 15, ${row}, 2, ${col + 3}, 31, ${row + 4}, 1</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData>\r\n` +
    ' </v:shape>'
  );
};
const BOXES = [boxOf('A1', 1025, ' o:custom="kept"'), boxOf('B2', 1026), boxOf('C3', 1027)];
const threeNotesVml = () =>
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">\r\n' +
  ' <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>\r\n' +
  ' <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>\r\n' +
  ` ${BOXES.join('\r\n ')}\r\n</xml>`;
const threeNotes = () =>
  drawingWorkbook({
    comments: `${XML}\r\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors><author>Ada</author><author>Grace &amp; Co</author></authors><commentList>${RICH}${SECOND}${THIRD}</commentList></comments>`,
    vml: threeNotesVml(),
  });

test('notes are written: an edited note gets its text; the other comments and every box keep their bytes', async () => {
  const original = await threeNotes();
  const doc = await load(original);
  assert.deepEqual(shown(doc.wb.sheets[0]), ['A1 by Ada: Ada:\nzengin  ', 'B2 by Grace & Co: ikinci', 'C3 by Ada: üçüncü']);
  note(doc, 0, 'B2', 'değişti <&> "tırnak"\n  ikinci satır ');
  const saved = await save(doc);
  await checkPackage(saved);
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.includes(RICH), 'the comment with runs is as it was');
  assert.ok(comments.includes(THIRD), 'the comment with unknown spacing and attributes is as it was');
  assert.ok(comments.includes('<comment ref="B2" authorId="1"><text><t xml:space="preserve">değişti &lt;&amp;&gt; "tırnak"\n  ikinci satır </t></text><commentPr locked="0"><anchor/></commentPr></comment>'), comments);
  assert.ok(comments.startsWith(`${XML}\r\n<comments `), 'what stands before the root is as it was');
  assert.ok(comments.includes('<authors><author>Ada</author><author>Grace &amp; Co</author></authors>'));
  same(await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml'), await ZipArchive.read(original).read('xl/drawings/vmlDrawing1.vml'), 'the legacy drawing');
  same(await ZipArchive.read(saved).read('xl/worksheets/_rels/sheet1.xml.rels'), await ZipArchive.read(original).read('xl/worksheets/_rels/sheet1.xml.rels'), 'the relationships');
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['A1 by Ada: Ada:\nzengin  ', 'B2 by Grace & Co: değişti <&> "tırnak"\n  ikinci satır ', 'C3 by Ada: üçüncü']);
  // The runs of the note that was not edited come back as runs.
  assert.equal((await reopen(saved)).sheets[0].comments![0].runs!.length, 2);
});

test('notes are written: a deleted note takes its comment and its box along, a new one is appended; the others keep their bytes', async () => {
  const original = await threeNotes();
  const doc = await load(original);
  note(doc, 0, 'C3', undefined);
  note(doc, 0, 'E5', 'yeni', 'Grace & Co');
  note(doc, 0, 'F6', 'başkası', 'Yeni <Yazar>');
  const saved = await save(doc);
  await checkPackage(saved);
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.includes(`<commentList>${RICH}${SECOND}<comment ref="E5" authorId="1" shapeId="0"><text><t>yeni</t></text></comment><comment ref="F6" authorId="2" shapeId="0"><text><t>başkası</t></text></comment></commentList>`), comments);
  assert.ok(comments.includes('<authors><author>Ada</author><author>Grace &amp; Co</author><author>Yeni &lt;Yazar&gt;</author></authors>'), comments);
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml.includes(BOXES[0]) && vml.includes(BOXES[1]), 'the boxes that stay are as they were, line breaks included');
  assert.ok(!vml.includes('<x:Row>2</x:Row><x:Column>2</x:Column>'), 'the box of the deleted note is gone');
  checkIds(vml);
  assert.deepEqual(shapeIds(vml), { ids: [1025, 1026, 1027, 1028], blocks: [1] });
  assert.equal(vml.split('<v:shapetype').length, 2, 'the shape type is declared once');
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['A1 by Ada: Ada:\nzengin  ', 'B2 by Grace & Co: ikinci', 'E5 by Grace & Co: yeni', 'F6 by Yeni <Yazar>: başkası']);

  // Deleted and added again on the same cell: a new note, not the old comment with another text.
  note(doc, 0, 'C3', 'yeniden', 'Ada');
  const again = await save(doc);
  const reread = await text(again, 'xl/comments1.xml');
  assert.ok(!reread.includes('future="1"') && reread.includes('<comment ref="C3" authorId="0" shapeId="0"><text><t>yeniden</t></text></comment>'), reread);
  checkIds(await text(again, 'xl/drawings/vmlDrawing1.vml'));
});

// --- (c) the last note -------------------------------------------------------------------

test('notes are written: deleting the last note takes the comment parts, their relationships and <legacyDrawing> out', async () => {
  const original = await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>tek</t>')), vml: vmlXml(vmlNote(1, 1)) });
  const doc = await load(original);
  note(doc, 0, 'B2', undefined);
  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(
    names(original).filter((name) => names(saved).indexOf(name) < 0),
    ['xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml'],
  );
  assert.ok(!(await text(saved, 'xl/worksheets/sheet1.xml')).includes('legacyDrawing'));
  assert.ok(!(await text(saved, 'xl/worksheets/_rels/sheet1.xml.rels')).includes('<Relationship '));
  const again = await reopen(saved);
  assert.equal(again.sheets[0].comments, undefined);
  assert.equal(again.sheets[0].anchored, undefined);

  // A note after that: the parts come back under their names, with the new note alone.
  note(doc, 0, 'C4', 'sonra');
  const later = await save(doc);
  await checkPackage(later);
  assert.deepEqual(names(later), names(original));
  assert.deepEqual(shown((await reopen(later)).sheets[0]), ['C4 by Author: sonra']);
});

// --- (d) form controls ---------------------------------------------------------------------

const BUTTON =
  '<v:shape id="_x0000_s1030" type="#_x0000_t201" style="position:absolute" o:button="t"><x:ClientData ObjectType="Button"><x:Anchor>1, 0, 5, 0, 2, 0, 6, 0</x:Anchor><x:Row>5</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>';
const BUTTON_TYPE = '<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201" path="m,l,21600r21600,l21600,xe"><v:path shadowok="f"/></v:shapetype>';
const controlsVml = (notes = '', type = '') =>
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>${BUTTON_TYPE}${type}${BUTTON}${notes}</xml>`;

test('notes are written: in a sheet with form controls a note is added beside them, and deleting it leaves their drawing', async () => {
  const original = await drawingWorkbook({ vml: controlsVml() });
  const doc = await load(original);
  assert.deepEqual(doc.wb.sheets[0].anchored, ['formControl', 'vmlDrawing']);
  note(doc, 0, 'B6', 'düğmenin yanında');
  const saved = await save(doc);
  await checkPackage(saved);
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml.includes(BUTTON) && vml.includes(BUTTON_TYPE), 'the control keeps its bytes');
  assert.deepEqual(shapeIds(vml), { ids: [1030, 1031], blocks: [1] });
  assert.ok(vml.includes('<v:shapetype id="_x0000_t202"'), 'the type of note boxes is declared');
  assert.ok((await text(saved, 'xl/worksheets/sheet1.xml')).includes('<legacyDrawing r:id="rIdV"/>'), 'the sheet names the same drawing');
  assert.deepEqual(names(saved).filter((name) => names(original).indexOf(name) < 0), ['xl/comments1.xml']);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['B6 by Author: düğmenin yanında']);
  assert.deepEqual((await reopen(saved)).sheets[0].anchored, ['comments', 'formControl', 'vmlDrawing']);

  // Deleted again (not undone): what the model says is what the file had.
  note(doc, 0, 'B6', undefined);
  same(await save(doc), original, 'the file as it was');
});

test('notes are written: the last note of a sheet with form controls goes, the drawing of the controls stays', async () => {
  const original = await drawingWorkbook({ comments: commentsXml(['A'], comment('C2', 0, '<t>tek</t>')), vml: controlsVml(vmlNote(1, 2)) });
  const doc = await load(original);
  note(doc, 0, 'C2', undefined);
  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(names(original).filter((name) => names(saved).indexOf(name) < 0), ['xl/comments1.xml']);
  assert.equal(await text(saved, 'xl/drawings/vmlDrawing1.vml'), controlsVml());
  assert.ok((await text(saved, 'xl/worksheets/sheet1.xml')).includes('<legacyDrawing r:id="rIdV"/>'));
  assert.equal((await reopen(saved)).sheets[0].comments, undefined);
});

test('notes are written: a new box gets an id from a block no drawing of the workbook uses when its own block is full', async () => {
  // Sheet 1 has a drawing whose block 1 is used up to its last id; sheet 2 has one that uses blocks 2 and 5.
  const full = controlsVml().replace('_x0000_s1030', '_x0000_s2047');
  const bytes = await buildXlsx({ sheets: [{ name: 'Bir', xml: worksheetXml('<sheetData/><legacyDrawing r:id="rId1"/>') }, { name: 'İki', xml: worksheetXml('<sheetData/><legacyDrawing r:id="rId1"/>') }] });
  const zip = ZipArchive.read(bytes);
  const rel = (n: number) => utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/vmlDrawing" Target="../drawings/vmlDrawing${n}.vml"/></Relationships>`);
  zip.set('xl/worksheets/_rels/sheet1.xml.rels', rel(1));
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', rel(2));
  zip.set('xl/drawings/vmlDrawing1.vml', utf8(full));
  zip.set('xl/drawings/vmlDrawing2.vml', utf8(controlsVml().replace('data="1"', 'data="2,5"').replace('_x0000_s1030', '_x0000_s2050')));
  zip.set('[Content_Types].xml', utf8(decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')));
  const doc = await load(await zip.write());
  note(doc, 0, 'A1', 'bir');
  note(doc, 1, 'A1', 'iki');
  const saved = await save(doc);
  await checkPackage(saved);
  const first = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  const second = await text(saved, 'xl/drawings/vmlDrawing2.vml');
  assert.deepEqual(shapeIds(first), { ids: [2047, 6 * 1024 + 1], blocks: [1, 6] });
  assert.deepEqual(shapeIds(second), { ids: [2050, 5 * 1024 + 1], blocks: [2, 5] });
  assert.ok(!(await text(saved, '[Content_Types].xml')).includes('vmlDrawing1.vml'), 'a part with a default content type gets no override');
});

// --- (e) threaded comments -----------------------------------------------------------------

const NS_TC = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
const PLACEHOLDER = comment('B2', 0, '<t>[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment.\n\nComment:\n    İlk ileti</t>');
const threadedBook = () =>
  drawingWorkbook({
    comments: commentsXml(['tc={AAAA}', 'Ada'], PLACEHOLDER + comment('D4', 1, '<t>eski not</t>')),
    vml: vmlXml(vmlNote(1, 1) + vmlNote(3, 3)),
    threaded:
      `${XML}<ThreadedComments xmlns="${NS_TC}">` +
      '<threadedComment ref="B2" dT="2024-01-01T10:00:00.00" personId="{P1}" id="{AAAA}"><text>İlk ileti</text></threadedComment>' +
      '<threadedComment ref="B2" dT="2024-01-01T11:00:00.00" personId="{P1}" id="{BBBB}" parentId="{AAAA}"><text>Yanıt</text></threadedComment></ThreadedComments>',
    persons: `${XML}<personList xmlns="${NS_TC}"><person displayName="Ada Lovelace" id="{P1}" userId="ada" providerId="None"/></personList>`,
  });

test('notes are written: a threaded comment is refused and kept byte for byte while the notes beside it are edited', async () => {
  const original = await threadedBook();
  const doc = await load(original);
  assert.equal(doc.wb.sheets[0].comments![0].threaded, true);
  assert.throws(() => note(doc, 0, 'B2', 'benim'), /threaded comment/);
  assert.throws(() => note(doc, 0, 'B2', undefined), /threaded comment/);
  assert.equal(doc.wb.sheets[0].modified, undefined, 'a refused change changes nothing');
  same(await save(doc), original, 'nothing was written');

  note(doc, 0, 'D4', 'yeni metin');
  note(doc, 0, 'F6', 'ek not', 'Ada');
  const saved = await save(doc);
  await checkPackage(saved);
  for (const part of ['xl/threadedComments/threadedComment1.xml', 'xl/persons/person.xml']) {
    same(await ZipArchive.read(saved).read(part), await ZipArchive.read(original).read(part), part);
  }
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.includes(PLACEHOLDER), 'the note written for older applications is as it was');
  assert.ok(comments.includes('<authors><author>tc={AAAA}</author><author>Ada</author></authors>'));
  assert.ok(comments.includes('<comment ref="F6" authorId="1" shapeId="0">'));
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml.includes(vmlNote(1, 1)) && vml.includes(vmlNote(3, 3)));
  const again = (await reopen(saved)).sheets[0];
  assert.deepEqual(shown(again), ['B2 by Ada Lovelace: İlk ileti', 'D4 by Ada: yeni metin', 'F6 by Ada: ek not']);
  assert.deepEqual(again.comments![0].replies, [{ author: 'Ada Lovelace', text: 'Yanıt' }]);

  // Deleting every note leaves the conversation and the note that stands for it.
  note(doc, 0, 'D4', undefined);
  note(doc, 0, 'F6', undefined);
  const less = await save(doc);
  await checkPackage(less);
  assert.ok((await text(less, 'xl/comments1.xml')).includes(`<commentList>${PLACEHOLDER}</commentList>`));
  assert.deepEqual(shown((await reopen(less)).sheets[0]), ['B2 by Ada Lovelace: İlk ileti']);
});

// --- (f) save, undo, save ------------------------------------------------------------------

test('notes are written: save, undo, save gives the file that was opened, byte for byte', async () => {
  for (const [label, make, steps] of [
    ['first note', plainSheet, [['B2', 'a']]],
    ['edited', threeNotes, [['B2', 'b']]],
    ['deleted and added', threeNotes, [['A1', undefined], ['Z9', 'c']]],
    ['last note deleted', () => drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>tek</t>')), vml: vmlXml(vmlNote(1, 1)) }), [['B2', undefined]]],
    ['beside form controls', () => drawingWorkbook({ vml: controlsVml() }), [['A1', 'd']]],
    ['beside a conversation', threadedBook, [['D4', undefined], ['A9', 'e']]],
  ] as [string, () => Promise<Uint8Array>, [string, string | undefined][]][]) {
    const original = await make();
    const doc = await load(original);
    for (const [ref, value] of steps) note(doc, 0, ref, value);
    const saved = await save(doc);
    await checkPackage(saved);
    assert.notDeepEqual(names(saved).length === names(original).length && Buffer.compare(saved, original) === 0, true, `${label}: the save changed the file`);
    for (let i = 0; i < steps.length; i++) doc.history.undo();
    same(await save(doc), original, `${label}: undone`);
    for (let i = 0; i < steps.length; i++) doc.history.redo();
    same(await save(doc), saved, `${label}: done again`);
    for (let i = 0; i < steps.length; i++) doc.history.undo();
    same(await save(doc), original, `${label}: undone again`);
  }
});

test('notes are written: rows that move take the notes along; undoing everything gives the file back', async () => {
  const original = await threeNotes();
  const doc = await load(original);
  apply(doc, { type: 'insert', sheet: 0, axis: 'row', at: 1, count: 2 });
  note(doc, 0, 'B4', 'kaydı'); // The note of B2, which is in B4 now.
  note(doc, 0, 'B2', 'araya');
  apply(doc, { type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 });
  assert.deepEqual(shown(doc.wb.sheets[0]), ['A2 by Author: araya', 'A4 by Grace & Co: kaydı', 'B5 by Ada: üçüncü']);
  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), shown(doc.wb.sheets[0]));
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  checkIds(vml);
  assert.equal(vml.split('ObjectType="Note"').length - 1, 3);
  for (let i = 0; i < 4; i++) doc.history.undo();
  await sameAround(await save(doc), original, 'undone');
});

// --- (h) sheets that are not in the file -----------------------------------------------------

test('notes are written: a sheet added since gets its notes in parts of its own, and loses them with the sheet', async () => {
  const original = await threeNotes();
  const doc = await load(original);
  apply(doc, { type: 'insertSheet', sheet: 1, name: 'Yeni' });
  note(doc, 1, 'B2', 'yeni sayfada', 'Ada');
  apply(doc, { type: 'insert', sheet: 1, axis: 'row', at: 0, count: 3 });
  note(doc, 1, 'A1', 'üstte');
  const saved = await save(doc);
  await checkPackage(saved);
  assert.deepEqual(
    names(saved).filter((name) => names(original).indexOf(name) < 0),
    ['xl/comments2.xml', 'xl/drawings/vmlDrawing2.vml', 'xl/worksheets/_rels/sheet2.xml.rels', 'xl/worksheets/sheet2.xml'],
  );
  const again = await reopen(saved);
  assert.deepEqual(shown(again.sheets[1]), ['A1 by Author: üstte', 'B5 by Ada: yeni sayfada']);
  assert.deepEqual(shown(again.sheets[0]), shown(doc.wb.sheets[0]));
  // Its boxes take a block of ids of their own.
  assert.deepEqual(shapeIds(await text(saved, 'xl/drawings/vmlDrawing2.vml')), { ids: [2049, 2050], blocks: [2] });
  same(await save(doc), saved, 'saved twice');

  // The notes go: so do their parts. The sheet stays.
  note(doc, 1, 'A1', undefined);
  note(doc, 1, 'B5', undefined);
  const bare = await save(doc);
  await checkPackage(bare);
  assert.deepEqual(names(bare).filter((name) => names(original).indexOf(name) < 0), ['xl/worksheets/sheet2.xml']);
  assert.ok(!(await text(bare, 'xl/worksheets/sheet2.xml')).includes('legacyDrawing'));

  // Undo to the notes, then all the way: the file as it was.
  doc.history.undo();
  doc.history.undo();
  await checkPackage(await save(doc));
  assert.deepEqual(shown((await reopen(await save(doc))).sheets[1]), ['A1 by Author: üstte', 'B5 by Ada: yeni sayfada']);
  for (let i = 0; i < 4; i++) doc.history.undo();
  same(await save(doc), original, 'undone');
});

test('notes are written: deleting a sheet takes the parts made for its notes along; undo brings them back', async () => {
  const original = await plainSheet();
  const doc = await load(original);
  note(doc, 0, 'A1', 'gidecek');
  note(doc, 1, 'A1', 'kalacak');
  const both = await save(doc);
  await checkPackage(both);
  apply(doc, { type: 'deleteSheet', sheet: 0 });
  const one = await save(doc);
  await checkPackage(one);
  assert.deepEqual(names(one).filter((name) => /comments|vml/.test(name)), ['xl/comments2.xml', 'xl/drawings/vmlDrawing2.vml']);
  assert.deepEqual(shown((await reopen(one)).sheets[0]), ['A1 by Author: kalacak']);
  doc.history.undo();
  // The content types list the same parts in another order.
  const back = await save(doc);
  await checkPackage(back);
  assert.deepEqual(names(back), names(both));
  for (const name of names(both)) if (name !== '[Content_Types].xml') same(await ZipArchive.read(back).read(name), await ZipArchive.read(both).read(name), `the sheet is back with its note: ${name}`);
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), original, 'undone');
});

// --- (i) what a text may hold ----------------------------------------------------------------

test('notes are written: markup characters, line breaks, emoji, control characters and long text come back as they went', async () => {
  const texts = [
    '<b>kalın</b> & "tırnak" \'tek\' ]]> <!-- -->',
    'satır 1\nsatır 2\n\n  girintili\tsekme ',
    '\u{1F600} emoji \u{1F469}‍\u{1F4BB} ve İıŞşĞğ',
    '_x000D_ gibi görünen metin ve \u0001 denetim',
    'x'.repeat(32767),
  ];
  const doc = await load(await plainSheet());
  texts.forEach((value, index) => note(doc, 0, cellName(index, 0), value, index === 0 ? 'A & <B> "C" \u{1F600}' : 'İkinci\tYazar'));
  const saved = await save(doc);
  await checkPackage(saved);
  const again = (await reopen(saved)).sheets[0];
  assert.deepEqual(
    again.comments!.map((item) => item.text),
    texts,
  );
  assert.deepEqual(
    again.comments!.map((item) => item.author),
    ['A & <B> "C" \u{1F600}', 'İkinci\tYazar', 'İkinci\tYazar', 'İkinci\tYazar', 'İkinci\tYazar'],
  );
  // The parts are well-formed for a reader that knows nothing of this writer.
  parseXml(await ZipArchive.read(saved).read('xl/comments1.xml'));
  parseXml(await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml'));
});

// --- what the reader does not take as a comment stays ------------------------------------------

test('notes are written: comments the reader passed over (a second one on a cell, one without a cell) are not taken for deleted ones', async () => {
  const twice = comment('B2', 0, '<t>ikincisi</t>');
  const lost = '<comment authorId="0"><text><t>hücresiz</t></text></comment>';
  const original = await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>') + twice + lost + comment('C3', 0, '<t>üç</t>')), vml: vmlXml(vmlNote(1, 1) + vmlNote(2, 2)) });
  const doc = await load(original);
  note(doc, 0, 'C3', undefined);
  note(doc, 0, 'B2', undefined);
  const saved = await save(doc);
  await checkPackage(saved);
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.includes(`<commentList>${twice}${lost}</commentList>`), comments);
});

test('notes are written: a legacy drawing that is not well-formed still takes and loses boxes', async () => {
  const broken = vmlXml(vmlNote(1, 1) + '<v:imagedata src="x.png" o:title=unquoted><br>').replace('</xml>', '</xml>');
  const original = await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>')), vml: broken });
  assert.throws(() => parseXml(broken));
  const doc = await load(original);
  note(doc, 0, 'E5', 'yeni');
  note(doc, 0, 'B2', undefined);
  const saved = await save(doc);
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml.includes('<v:imagedata src="x.png" o:title=unquoted><br>'));
  assert.ok(!vml.includes('<x:Row>1</x:Row>') && vml.includes('<x:Row>4</x:Row><x:Column>4</x:Column>'));
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['E5 by Author: yeni']);
});

test('notes are written: a comments part two sheets share is not changed for one of them', async () => {
  const bytes = await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ortak</t>')), vml: vmlXml(vmlNote(1, 1)) });
  const zip = ZipArchive.read(bytes);
  const workbook = decoder.decode(await zip.read('xl/workbook.xml')).replace('</sheets>', '<sheet name="Öteki" sheetId="2" r:id="rId9"/></sheets>');
  zip.set('xl/workbook.xml', utf8(workbook));
  zip.set('xl/_rels/workbook.xml.rels', utf8(decoder.decode(await zip.read('xl/_rels/workbook.xml.rels')).replace('</Relationships>', `<Relationship Id="rId9" Type="${NS_REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)));
  zip.set('xl/worksheets/sheet2.xml', utf8(worksheetXml('<sheetData/>')));
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/comments" Target="../comments1.xml"/></Relationships>`));
  const original = await zip.write();
  const doc = await load(original);
  // The reader marks such sheets and the operations refuse; the writer is the last defence, reached here without the mark.
  assert.equal(doc.wb.sheets[0].notesBlocked, 'shared');
  delete doc.wb.sheets[0].notesBlocked;
  note(doc, 0, 'B2', 'değişti');
  await assert.rejects(save(doc), /is shared with sheet "Öteki"/);
  // The package is as it was.
  same(await doc.file.package!.save(), original, 'after the refused save');
});

// --- xls and other models: buildXlsx ---------------------------------------------------------

test('notes are written: buildXlsx places notes added since by the same rule', async () => {
  const doc = await load(await threeNotes());
  apply(doc, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  note(doc, 0, 'F1', 'sonradan');
  apply(doc, { type: 'insert', sheet: 0, axis: 'col', at: 0, count: 2 });
  const built = await reopen(await buildFromModel(doc.wb));
  assert.deepEqual(shown(built.sheets[0]), shown(doc.wb.sheets[0]));
  assert.deepEqual(shown(built.sheets[0]), ['C2 by Ada: Ada:\nzengin  ', 'D3 by Grace & Co: ikinci', 'E4 by Ada: üçüncü', 'H1 by Author: sonradan']);
});

// --- (g) a random mix --------------------------------------------------------------------------

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

test('notes are written: after any mix of notes, moved rows and columns, undo and redo, what is shown is what is read back', async () => {
  const fixtures: [string, () => Promise<Uint8Array>][] = [
    ['three notes', threeNotes],
    ['no notes', plainSheet],
    ['form controls', () => drawingWorkbook({ comments: commentsXml(['A'], comment('C2', 0, '<t>tek</t>')), vml: controlsVml(vmlNote(1, 2)) })],
  ];
  let saves = 0;
  for (let seed = 1; seed <= 24; seed++) {
    const [label, make] = fixtures[seed % fixtures.length];
    const original = await make();
    const doc = await load(original);
    const next = random(seed * 7919);
    const pick = (n: number) => Math.floor(next() * n);
    const trail: string[] = [];
    let depth = 0;
    const check = async (stage: string) => {
      const where = `seed ${seed} (${label}) ${stage}: ${trail.join('; ')}`;
      const saved = await save(doc);
      saves++;
      await checkPackage(saved);
      const again = await reopen(saved);
      assert.deepEqual(shown(again.sheets[0]), shown(doc.wb.sheets[0]), where);
      if (await ZipArchive.read(saved).has('xl/drawings/vmlDrawing1.vml')) {
        const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
        checkIds(vml);
        // One box per comment that is shown.
        assert.equal(vml.split('ObjectType="Note"').length - 1, shown(doc.wb.sheets[0]).length, where);
      } else assert.equal(shown(doc.wb.sheets[0]).length, 0, where);
      // The boxes are where their comments are.
      for (const item of again.sheets[0].comments || []) assert.ok(item.box, `${where}: ${cellName(item.row, item.col)} has a box`);
      return saved;
    };
    for (let step = 0; step < 14; step++) {
      const roll = next();
      const sheet = doc.wb.sheets[0];
      try {
        if (roll < 0.4) {
          const ref = cellName(pick(8), pick(6));
          trail.push(`note ${ref}`);
          note(doc, 0, ref, `n${seed}.${step}`, next() < 0.5 ? 'Ada' : undefined).length > 0 && depth++;
        } else if (roll < 0.55) {
          const current = new SheetDrawings(sheet).currentComments();
          if (current.length === 0) continue;
          const target = current[pick(current.length)];
          trail.push(`delete ${cellName(target.row, target.col)}`);
          note(doc, 0, cellName(target.row, target.col), undefined).length > 0 && depth++;
        } else if (roll < 0.85) {
          const op: Op = { type: next() < 0.5 ? 'insert' : 'delete', sheet: 0, axis: next() < 0.5 ? 'row' : 'col', at: pick(6), count: 1 + pick(2) };
          trail.push(`${op.type} ${op.axis} ${op.at}+${op.count}`);
          apply(doc, op);
          depth++;
        } else if (roll < 0.95 && depth > 0) {
          trail.push('undo');
          doc.history.undo();
          depth--;
        } else if (doc.history.canRedo) {
          trail.push('redo');
          doc.history.redo();
          depth++;
        }
      } catch (error) {
        // A sheet with form controls refuses to move rows; nothing changed then.
        if (!(error instanceof RangeError)) throw error;
        trail.push('(refused)');
      }
      if (step % 4 === 3) await check(`step ${step}`);
    }
    await check('end');
    while (doc.history.canUndo) doc.history.undo();
    const undone = await save(doc);
    if (doc.wb.sheets[0].modified) await sameAround(undone, original, `seed ${seed} (${label}): everything undone: ${trail.join('; ')}`);
    else same(undone, original, `seed ${seed} (${label}): everything undone: ${trail.join('; ')}`);
  }
  assert.ok(saves > 90);
  assert.ok(MAX_ROWS > 0 && MAX_COLS > 0);
});

const rootOf = (xml: string): XmlElement => parseXml(xml).root;

test('notes are written: a strict workbook gets its new parts in its own namespaces', async () => {
  const strictMain = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
  const strictRel = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
  const bytes = await plainSheet();
  const zip = ZipArchive.read(bytes);
  for (const name of zip.names()) {
    if (!/\.(xml|rels)$/.test(name) || name === '[Content_Types].xml') continue;
    const before = decoder.decode(await zip.read(name));
    zip.set(name, utf8(before.split('http://schemas.openxmlformats.org/spreadsheetml/2006/main').join(strictMain).split(NS_REL).join(strictRel)));
  }
  const doc = await load(await zip.write());
  note(doc, 0, 'B2', 'katı');
  const saved = await save(doc);
  assert.equal(rootOf(await text(saved, 'xl/comments1.xml')).attrs.get('xmlns'), strictMain);
  assert.ok((await text(saved, 'xl/worksheets/_rels/sheet1.xml.rels')).includes(`Type="${strictRel}/comments"`));
});

// --- review ------------------------------------------------------------------------------------

test('review: a note with a carriage return that is given back as a text field holds it is not written', async () => {
  const original = await drawingWorkbook({
    comments: commentsXml(['Ada'], comment('B2', 0, '<r><rPr><b/></rPr><t>Ada:</t></r><r><t xml:space="preserve">_x000D_\nbody</t></r>')),
    vml: vmlXml(vmlNote(1, 1)),
  });
  const doc = await load(original);
  assert.equal(doc.wb.sheets[0].comments![0].text, 'Ada:\r\nbody');
  assert.deepEqual(note(doc, 0, 'B2', 'Ada:\nbody'), []);
  assert.equal(doc.wb.sheets[0].comments![0].runs!.length, 2);
  same(await save(doc), original, 'nothing was written');
});

/** UTF-16 (little endian, with its mark) of a text. */
function utf16(value: string): Uint8Array {
  const out = new Uint8Array(2 + value.length * 2);
  out[0] = 0xff;
  out[1] = 0xfe;
  for (let i = 0; i < value.length; i++) {
    out[2 + i * 2] = value.charCodeAt(i) & 0xff;
    out[3 + i * 2] = value.charCodeAt(i) >> 8;
  }
  return out;
}

test('review: a legacy drawing in UTF-16 is written with a declaration that says what its bytes are', async () => {
  const vml = `<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\r\n${vmlXml(vmlNote(1, 1))}`;
  const original = await drawingWorkbook({ comments: commentsXml(['Ada'], comment('B2', 0, '<t>ilk</t>')), extra: { 'xl/drawings/vmlDrawing1.vml': utf16(vml) }, vml: '' });
  const doc = await load(original);
  assert.equal(doc.wb.sheets[0].notesBlocked, undefined);
  note(doc, 0, 'D4', 'yeni');
  const saved = await save(doc);
  const bytes = await ZipArchive.read(saved).read('xl/drawings/vmlDrawing1.vml');
  // The bytes are UTF-8 (no mark, one byte per ASCII character), and the declaration says so.
  assert.notEqual(bytes[0], 0xff);
  const written = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  assert.match(written, /^<\?xml version="1.0" encoding="UTF-8" standalone="yes"\?>/);
  assert.ok(!/UTF-16/i.test(written));
  parseXml(bytes);
  assert.ok(written.includes(vmlNote(1, 1)), 'the box that was there is as it was');
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['B2 by Ada: ilk', 'D4 by Author: yeni']);
});

test('review: line breaks of the file stay as they are in a comments part and in a sheet part that are written for a note', async () => {
  const first = '<comment ref="A1" authorId="0">\r\n  <text>\r\n    <t>bir\r\niki</t>\r\n  </text>\r\n</comment>';
  const second = '<comment ref="B2" authorId="0">\r\n  <text><t>düzenlenecek</t></text>\r\n</comment>';
  const comments = `${XML}\r\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\r\n<authors>\r\n<author>Ada</author>\r\n</authors>\r\n<commentList>\r\n${first}\r\n${second}\r\n</commentList>\r\n</comments>\r\n`;
  const original = await drawingWorkbook({ comments, vml: vmlXml(vmlNote(0, 0) + vmlNote(1, 1)) });
  const doc = await load(original);
  note(doc, 0, 'B2', 'değişti');
  note(doc, 0, 'C3', 'eklendi', 'Yeni');
  const saved = await save(doc);
  const written = await text(saved, 'xl/comments1.xml');
  assert.ok(written.includes(first), 'the comment that was not touched has its bytes');
  assert.equal(
    written,
    comments
      .replace('<text><t>düzenlenecek</t></text>', '<text><t>değişti</t></text>')
      .replace('<author>Ada</author>\r\n', '<author>Ada</author>\r\n<author>Yeni</author>')
      .replace(`${second.replace('düzenlenecek', 'değişti')}\r\n`, `${second.replace('düzenlenecek', 'değişti')}\r\n<comment ref="C3" authorId="1" shapeId="0"><text><t>eklendi</t></text></comment>`),
  );

  // The sheet part of a sheet that gets its first note: only <legacyDrawing> is new.
  const sheetXml = `${XML}\r\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_REL}">\r\n<sheetData>\r\n<row r="1"><c r="A1"><v>1</v></c></row>\r\n</sheetData>\r\n<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>\r\n<extLst><ext uri="{X}"/></extLst>\r\n</worksheet>\r\n`;
  const plain = await buildXlsx({ sheets: [{ name: 'Bir', xml: sheetXml }] });
  const second2 = await load(plain);
  note(second2, 0, 'A1', 'ilk');
  const out = await save(second2);
  assert.equal(await text(out, 'xl/worksheets/sheet1.xml'), sheetXml.replace('<extLst>', '<legacyDrawing r:id="rId2"/><extLst>'));
  await checkPackage(out);
  second2.history.undo();
  same(await save(second2), plain, 'undone');
});

test('review: a new author does not take an index that comments of the file use without an author being listed for it', async () => {
  const original = await drawingWorkbook({
    comments: `${XML}<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors/><commentList>${comment('A1', 0, '<t>adsız</t>')}${comment('A2', 2, '<t>adsız 2</t>')}</commentList></comments>`,
    vml: vmlXml(vmlNote(0, 0) + vmlNote(1, 0)),
  });
  const doc = await load(original);
  assert.deepEqual(shown(doc.wb.sheets[0]), ['A1 by : adsız', 'A2 by : adsız 2']);
  note(doc, 0, 'C3', 'yeni', 'Ada');
  const saved = await save(doc);
  const written = await text(saved, 'xl/comments1.xml');
  assert.ok(written.includes('<authors><author></author><author></author><author></author><author>Ada</author></authors>'), written);
  assert.ok(written.includes('<comment ref="C3" authorId="3" shapeId="0">'), written);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), ['A1 by : adsız', 'A2 by : adsız 2', 'C3 by Ada: yeni']);
});

test('review: <legacyDrawing> stands before controls that are wrapped in mc:AlternateContent', async () => {
  const controls =
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><controls><control shapeId="1025" r:id="rId9" name="Check Box 1"/></controls></mc:Choice></mc:AlternateContent>';
  const drawing = '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x14"><drawing r:id="rId8"/></mc:Choice></mc:AlternateContent>';
  const original = await buildXlsx({ sheets: [{ name: 'Bir', xml: worksheetXml(`<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${drawing}${controls}<extLst/>`) }] });
  // Once with the sheet itself untouched, once with a cell written as well.
  for (const withCell of [false, true]) {
    const doc = await load(original);
    note(doc, 0, 'B2', 'not');
    if (withCell) doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] }], undefined, undefined);
    const sheetXml = await text(await save(doc), 'xl/worksheets/sheet1.xml');
    assert.ok(sheetXml.includes(`${drawing}<legacyDrawing r:id="rId2"/>${controls}<extLst/>`), `${withCell}: ${sheetXml}`);
  }
});

test('review: a sheet whose comment parts cannot take notes is marked by the reader, and its notes are not changed', async () => {
  // A legacy drawing that is empty, and one without an end.
  for (const vml of ['', '<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="x"']) {
    const doc = await load(await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>')), vml }));
    assert.equal(doc.wb.sheets[0].notesBlocked, 'drawing', JSON.stringify(vml));
    assert.throws(() => note(doc, 0, 'C3', 'yeni'), /its legacy drawing cannot be read/);
    assert.throws(() => note(doc, 0, 'B2', undefined), /its legacy drawing cannot be read/);
    // Everything else of the sheet is edited and saved.
    doc.history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] }], undefined, undefined);
    assert.deepEqual(shown((await reopen(await save(doc))).sheets[0]), ['B2 by A: ilk']);
  }
  // A legacy drawing that is not well-formed but has its root is none of that.
  const loose = await load(await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>')), vml: vmlXml(`${vmlNote(1, 1)}<br>`) }));
  assert.equal(loose.wb.sheets[0].notesBlocked, undefined);

  // Two comment parts on one sheet.
  const two = ZipArchive.read(await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>')), vml: vmlXml(vmlNote(1, 1)) }));
  two.set('xl/comments2.xml', utf8(commentsXml(['B'], comment('C3', 0, '<t>iki</t>'))));
  two.set('xl/worksheets/_rels/sheet1.xml.rels', utf8(decoder.decode(await two.read('xl/worksheets/_rels/sheet1.xml.rels')).replace('</Relationships>', `<Relationship Id="rIdC2" Type="${NS_REL}/comments" Target="../comments2.xml"/></Relationships>`)));
  const parts = await load(await two.write());
  assert.equal(parts.wb.sheets[0].notesBlocked, 'parts');
  assert.throws(() => note(parts, 0, 'B2', 'x'), /more than one comments part/);

  // A comments part two sheets show.
  const zip = ZipArchive.read(await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ortak</t>')), vml: vmlXml(vmlNote(1, 1)) }));
  zip.set('xl/workbook.xml', utf8(decoder.decode(await zip.read('xl/workbook.xml')).replace('</sheets>', '<sheet name="Öteki" sheetId="2" r:id="rId9"/></sheets>')));
  zip.set('xl/_rels/workbook.xml.rels', utf8(decoder.decode(await zip.read('xl/_rels/workbook.xml.rels')).replace('</Relationships>', `<Relationship Id="rId9" Type="${NS_REL}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>`)));
  zip.set('xl/worksheets/sheet2.xml', utf8(worksheetXml('<sheetData/>')));
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/comments" Target="../comments1.xml"/></Relationships>`));
  const shared = await load(await zip.write());
  assert.deepEqual(shared.wb.sheets.map((sheet) => sheet.notesBlocked), ['shared', 'shared']);
  assert.throws(() => note(shared, 0, 'B2', 'x'), /shared with another sheet/);
  assert.throws(() => note(shared, 1, 'D4', 'x'), /shared with another sheet/);
});

test('review: deleting a note takes one box of its cell along: a form control on the same cell stays', async () => {
  const control = '<v:shape id="_x0000_s1040" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"><x:Anchor>1, 0, 1, 0, 2, 0, 2, 0</x:Anchor><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape>';
  // Two note boxes on B2 (as a damaged file has them) and a control there as well.
  const original = await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>ilk</t>') + comment('D4', 0, '<t>kalan</t>')), vml: vmlXml(control + vmlNote(1, 1) + vmlNote(1, 1).replace('_x0000_s1126', '_x0000_s1999') + vmlNote(3, 3)) });
  const doc = await load(original);
  note(doc, 0, 'B2', undefined);
  const vml = await text(await save(doc), 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml.includes(control), 'the control is there');
  assert.equal(vml.split('ObjectType="Note"').length - 1, 2, 'one of the two boxes of the cell went');
  assert.ok(vml.includes('_x0000_s1999') && !vml.includes('_x0000_s1126'), 'the first one');
});

test('review 2: comment parts whose relationships spell their names in another case come back under the names of the archive', async () => {
  const zip = ZipArchive.read(await drawingWorkbook({ comments: commentsXml(['A'], comment('B2', 0, '<t>tek</t>')), vml: vmlXml(vmlNote(1, 1)) }));
  const rels = 'xl/worksheets/_rels/sheet1.xml.rels';
  const before = decoder.decode(await zip.read(rels));
  const spelled = before.replace('../comments1.xml', '../Comments1.xml').replace('../drawings/vmlDrawing1.vml', '../Drawings/VMLDrawing1.vml');
  assert.notEqual(spelled, before);
  zip.set(rels, utf8(spelled));
  const original = await zip.write();
  const doc = await load(original);
  assert.deepEqual(shown(doc.wb.sheets[0]), ['B2 by A: tek']);
  note(doc, 0, 'B2', undefined);
  const saved = await save(doc);
  await checkPackage(saved);
  assert.ok(names(saved).indexOf('xl/comments1.xml') < 0 && names(saved).indexOf('xl/drawings/vmlDrawing1.vml') < 0);
  doc.history.undo();
  const back = await save(doc);
  assert.deepEqual(names(back), names(original));
  same(back, original, 'undone');
});
