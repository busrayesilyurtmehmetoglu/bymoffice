// Notes that travel with their cells: a copy inside the editor takes the notes
// of the copied cells along, a cut takes them away, and a sort moves each note
// with its row. What the writer makes of a note that a sort moved: the comment
// of the file with another cell, its box beside it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { writeXlsx } from '../src/formats/xlsx/write';
import { cellName, open, parseCellRef, parseXml, ZipArchive } from '../src/index';
import type { CellRange, OfficeFile, Sheet, Workbook } from '../src/index';
import { clearRange, copyRange, cutRange, pasteAt } from '../src/ops/commands';
import { setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { SheetDrawings, adjustDrawingAnchor, commentEdits } from '../src/render/sheet-drawings';
import { utf8 } from './fixtures';
import { comment, commentsXml, drawingWorkbook as bareWorkbook, vmlNote, vmlXml, type DrawingBookSpec } from './sheet-drawing-fixtures';

const decoder = new TextDecoder();

async function drawingWorkbook(spec: DrawingBookSpec): Promise<Uint8Array> {
  const zip = ZipArchive.read(await bareWorkbook(spec));
  const types = decoder.decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', utf8(types.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')));
  return zip.write();
}

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
const names = (bytes: Uint8Array) => ZipArchive.read(bytes).names().sort();
const text = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));
const range = (ref: string): CellRange => {
  const [a, b] = ref.split(':');
  const from = parseCellRef(a)!;
  const to = parseCellRef(b || a)!;
  return { top: from.row, left: from.col, bottom: to.row, right: to.col };
};

function same(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

/** The file as it was around the sheets: see the test of notes that are written. */
async function sameAround(actual: Uint8Array, expected: Uint8Array, label: string): Promise<void> {
  assert.deepEqual(names(actual), names(expected), label);
  for (const name of names(expected)) {
    if (name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
    same(await ZipArchive.read(actual).read(name), await ZipArchive.read(expected).read(name), `${label}: ${name}`);
  }
}

const note = (doc: Loaded, ref: string, value: string | undefined, author?: string, sheet = 0) => {
  const { row, col } = parseCellRef(ref)!;
  return doc.history.apply(setComment(doc.wb, sheet, row, col, value, author === undefined ? {} : { author }), undefined, undefined);
};

/** The comments of a sheet as they are shown: "B2 by Ada: text", sorted. */
const shown = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => `${cellName(placed.row, placed.col)} by ${placed.comment.author ?? ''}: ${placed.comment.text}`)
    .sort();

/** Where the boxes of the notes are, by the rows of their anchors: "B2@1-5". */
const boxes = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => {
      const box = placed.comment.box ? adjustDrawingAnchor(placed.comment.box, commentEdits(placed.comment, sheet.structuralEdits)) : undefined;
      return `${cellName(placed.row, placed.col)}@${box && box.from && box.to ? `${box.from.row}-${box.to.row}` : '?'}`;
    })
    .sort();

const values = (sheet: Sheet, col: number, top: number, bottom: number) => {
  const out: unknown[] = [];
  for (let r = top; r <= bottom; r++) {
    const row = sheet.rows.get(r);
    const cell = row ? row.cells[col] : undefined;
    out.push(cell ? cell.value : null);
  }
  return out;
};

const rows = (cells: (number | string)[][]) =>
  `<sheetData>${cells
    .map(
      (line, r) =>
        `<row r="${r + 1}">${line.map((value, c) => (typeof value === 'number' ? `<c r="${cellName(r, c)}"><v>${value}</v></c>` : `<c r="${cellName(r, c)}" t="inlineStr"><is><t>${value}</t></is></c>`)).join('')}</row>`,
    )
    .join('')}</sheetData>`;

const RICH = '<comment ref="A1" authorId="0" shapeId="0" xr:uid="{11111111-2222-3333-4444-555555555555}"><text><r><rPr><b/><sz val="9"/></rPr><t>Ada:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve">\nzengin  </t></r></text></comment>';
const PLAIN = '<comment ref="A3" authorId="1" future="1"><text><t>üç</t></text></comment>';
const OUTSIDE = '<comment ref="D2" authorId="0"><text><t>dışarıda</t></text></comment>';

/** A1..B4 with numbers to sort by in column A; notes on A1 (rich), A3 and, beside the data, D2. */
const sortable = () =>
  drawingWorkbook({
    sheet: rows([
      [3, 'c'],
      [1, 'a'],
      [4, 'd'],
      [2, 'b'],
    ]),
    comments: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><authors><author>Ada</author><author>Grace</author></authors><commentList>${RICH}${PLAIN}${OUTSIDE}</commentList></comments>`,
    vml: vmlXml(vmlNote(0, 0) + vmlNote(2, 0) + vmlNote(1, 3)),
  });

const sortA = (doc: Loaded, ref = 'A1:B4') => doc.history.apply(sortRange(doc.wb, 0, range(ref), [{ col: 0 }]), undefined, undefined);

// --- A1: copy, cut and paste -------------------------------------------------------------

test('notes are copied: a copy inside the editor takes the notes of its cells along, a paste puts them on the cells it writes', async () => {
  const doc = await load(await sortable());
  const sheet = doc.wb.sheets[0];
  const copied = copyRange(doc.wb, 0, range('A1:B2'));
  assert.deepEqual(copied.notes, [{ row: 0, col: 0, text: 'Ada:\nzengin  ', author: 'Ada' }]);
  // The target has a note where the copy has none (E6) and one where the copy has one (D5).
  note(doc, 'D5', 'eski', 'Zed');
  note(doc, 'E6', 'kalır', 'Zed');
  const result = pasteAt(doc.wb, 0, range('D5'), copied.text, copied);
  assert.equal(result.skippedNotes, undefined);
  doc.history.apply(result.ops, undefined, undefined);
  assert.deepEqual(shown(sheet), ['A1 by Ada: Ada:\nzengin  ', 'A3 by Grace: üç', 'D2 by Ada: dışarıda', 'D5 by Ada: Ada:\nzengin  ', 'E6 by Zed: kalır']);
  assert.deepEqual(values(sheet, 3, 4, 5), [3, 1]);
  // The copy is a note of its own: plain text, placed like a new one.
  const pasted = (sheet.comments || []).find((item) => item.text === 'Ada:\nzengin  ' && item.since !== undefined)!;
  assert.equal(pasted.runs, undefined);
  assert.equal(pasted.box, undefined);

  const saved = await save(doc);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), shown(sheet));
  // One step: the cells and the notes go together.
  doc.history.undo();
  assert.deepEqual(shown(sheet), ['A1 by Ada: Ada:\nzengin  ', 'A3 by Grace: üç', 'D2 by Ada: dışarıda', 'D5 by Zed: eski', 'E6 by Zed: kalır']);
  assert.deepEqual(values(sheet, 3, 4, 5), [null, null]);
  doc.history.redo();
  assert.deepEqual(shown(sheet).length, 5);
});

test('notes are copied: a copy repeated over a larger selection repeats its notes; text from elsewhere brings none', async () => {
  const doc = await load(await sortable());
  const sheet = doc.wb.sheets[0];
  const copied = copyRange(doc.wb, 0, range('A1:A2'));
  doc.history.apply(pasteAt(doc.wb, 0, range('F1:G4'), copied.text, copied).ops, undefined, undefined);
  assert.deepEqual(
    shown(sheet).filter((line) => /^[FG]/.test(line)),
    ['F1 by Ada: Ada:\nzengin  ', 'F3 by Ada: Ada:\nzengin  ', 'G1 by Ada: Ada:\nzengin  ', 'G3 by Ada: Ada:\nzengin  '],
  );
  // The same text without the copy (another program had it): cells only.
  const before = shown(sheet);
  doc.history.apply(pasteAt(doc.wb, 0, range('J1'), copied.text).ops, undefined, undefined);
  assert.deepEqual(shown(sheet), before);
  assert.deepEqual(values(sheet, 9, 0, 1), [3, 1]);
});

test('notes are copied: a cut takes the notes of its cells away in the same step; clearing cells leaves notes alone', async () => {
  const doc = await load(await sortable());
  const sheet = doc.wb.sheets[0];
  doc.history.apply(clearRange(doc.wb, 0, range('A1:B4')), undefined, undefined);
  assert.equal(shown(sheet).length, 3);
  doc.history.undo();

  const copied = copyRange(doc.wb, 0, range('A1:B3'));
  doc.history.apply(cutRange(doc.wb, 0, range('A1:B3')), undefined, undefined);
  assert.deepEqual(shown(sheet), ['D2 by Ada: dışarıda']);
  assert.deepEqual(values(sheet, 0, 0, 3), [null, null, null, 2]);
  doc.history.apply(pasteAt(doc.wb, 0, range('A6'), copied.text, copied).ops, undefined, undefined);
  assert.deepEqual(shown(sheet), ['A6 by Ada: Ada:\nzengin  ', 'A8 by Grace: üç', 'D2 by Ada: dışarıda']);
  const saved = await save(doc);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), shown(sheet));
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(shown(sheet), ['A1 by Ada: Ada:\nzengin  ', 'A3 by Grace: üç', 'D2 by Ada: dışarıda']);
  assert.deepEqual(values(sheet, 0, 0, 3), [3, 1, 4, 2]);
  await sameAround(await save(doc), await sortable(), 'undone');
});

test('notes are copied: a conversation is not copied and not written over, a sheet whose notes cannot be changed gets the cells alone', async () => {
  const threaded =
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B2" dT="2024-01-01T00:00:00.00" personId="{P1}" id="{T1}"><text>İlk ileti</text></threadedComment></ThreadedComments>';
  const doc = await load(
    await drawingWorkbook({
      sheet: rows([[1, 2]]),
      comments: commentsXml(['Ada'], comment('A1', 0, '<t>not</t>') + comment('B2', 0, '<t>[Threaded comment]</t>')),
      vml: vmlXml(vmlNote(0, 0) + vmlNote(1, 1)),
      threaded,
    }),
  );
  const sheet = doc.wb.sheets[0];
  // The conversation of B2 is not in the copy.
  const copied = copyRange(doc.wb, 0, range('A1:B2'));
  assert.deepEqual(copied.notes, [{ row: 0, col: 0, text: 'not', author: 'Ada' }]);
  // Pasted so that the note lands on the conversation: the cells are written, the note is left out and counted.
  const result = pasteAt(doc.wb, 0, range('B2'), copied.text, copied);
  assert.equal(result.skippedNotes, 1);
  doc.history.apply(result.ops, undefined, undefined);
  assert.deepEqual(values(sheet, 1, 1, 1), [1]);
  assert.equal((sheet.comments || []).filter((item) => item.threaded).length, 1);
  assert.equal(shown(sheet).length, 2);

  sheet.notesBlocked = 'shared';
  const blocked = pasteAt(doc.wb, 0, range('E5'), copied.text, copied);
  assert.equal(blocked.skippedNotes, 1);
  assert.ok(blocked.ops.every((op) => op.type !== 'setComments'));
  // A cut leaves the notes it cannot take.
  assert.ok(cutRange(doc.wb, 0, range('A1:B2')).every((op) => op.type !== 'setComments'));
});

test('notes are copied: out of a filtered list only the notes of the rows that are shown', async () => {
  const doc = await load(await sortable());
  const sheet = doc.wb.sheets[0];
  doc.history.apply([{ type: 'setAutoFilter', sheet: 0, autoFilter: { range: range('A1:B4'), columns: [{ colId: 0, kind: 'values', values: ['1', '2'] }] }, hide: [2] }], undefined, undefined);
  const copied = copyRange(doc.wb, 0, range('A1:B4'));
  assert.deepEqual(copied.rows, [0, 1, 3]);
  assert.deepEqual(copied.notes, [{ row: 0, col: 0, text: 'Ada:\nzengin  ', author: 'Ada' }]);
  assert.equal(sheet.rows.get(2)!.hidden, true);
});

// --- A2: sorting ---------------------------------------------------------------------------

test('notes are sorted: each note goes with its row, notes outside the range stay, undo puts them back', async () => {
  const original = await sortable();
  const doc = await load(original);
  const sheet = doc.wb.sheets[0];
  sortA(doc);
  assert.deepEqual(values(sheet, 0, 0, 3), [1, 2, 3, 4]);
  // 3 (A1) went to row 3, 4 (A3) to row 4; D2 is outside the columns that were sorted.
  assert.deepEqual(shown(sheet), ['A3 by Ada: Ada:\nzengin  ', 'A4 by Grace: üç', 'D2 by Ada: dışarıda']);
  assert.deepEqual(boxes(sheet), ['A3@2-5', 'A4@2-6', 'D2@0-4']);

  const saved = await save(doc);
  const again = (await reopen(saved)).sheets[0];
  assert.deepEqual(shown(again), shown(sheet));
  assert.deepEqual(boxes(again), boxes(sheet));
  // The comments are the ones of the file with other cells: formats, attributes and authors stay.
  const comments = await text(saved, 'xl/comments1.xml');
  assert.ok(comments.indexOf(RICH.replace('ref="A1"', 'ref="A3"')) >= 0, comments);
  assert.ok(comments.indexOf('<comment ref="A4" authorId="1" future="1"><text><t>üç</t></text></comment>') >= 0, comments);
  assert.ok(comments.indexOf(OUTSIDE) >= 0, comments);
  assert.equal(again.comments![0].runs!.length, 2);
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.equal(vml.split('ObjectType="Note"').length - 1, 3);
  assert.match(vml, /<x:Anchor>1, 15, 2, 2, 3, 15, 5, 16<\/x:Anchor><x:AutoFill>False<\/x:AutoFill><x:Row>2<\/x:Row><x:Column>0<\/x:Column>/);
  assert.match(vml, /<x:Row>3<\/x:Row><x:Column>0<\/x:Column>/);
  assert.match(vml, /<x:Row>1<\/x:Row><x:Column>3<\/x:Column>/);

  doc.history.undo();
  assert.deepEqual(shown(sheet), ['A1 by Ada: Ada:\nzengin  ', 'A3 by Grace: üç', 'D2 by Ada: dışarıda']);
  await sameAround(await save(doc), original, 'undone');
  doc.history.redo();
  same(await save(doc), saved, 'done again');
});

test('notes are sorted: a note edited, deleted or added after the sort, and rows moved after it, are written where they are shown', async () => {
  const doc = await load(await sortable());
  const sheet = doc.wb.sheets[0];
  sortA(doc);
  note(doc, 'A3', 'başka'); // The rich note of the file, moved and then edited: its author stays.
  note(doc, 'A4', undefined); // The plain one, moved and then deleted.
  note(doc, 'A1', 'yeni', 'Zed'); // Where the rich one was.
  doc.history.apply([{ type: 'delete', sheet: 0, axis: 'row', at: 0, count: 1 }], undefined, undefined);
  doc.history.apply([{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }], undefined, undefined);
  assert.deepEqual(shown(sheet), ['B2 by Ada: başka', 'E1 by Ada: dışarıda']);
  const saved = await save(doc);
  const again = (await reopen(saved)).sheets[0];
  assert.deepEqual(shown(again), shown(sheet));
  assert.deepEqual(boxes(again), boxes(sheet));
  const vml = await text(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.equal(vml.split('ObjectType="Note"').length - 1, 2);
  // Sorted again, with a note added since: both kinds move.
  note(doc, 'B3', 'eklenen', 'Zed');
  doc.history.apply(sortRange(doc.wb, 0, range('B1:C3'), [{ col: 1, descending: true }]), undefined, undefined);
  assert.deepEqual(values(sheet, 1, 0, 2), [4, 3, 2]);
  assert.deepEqual(shown(sheet), ['B1 by Zed: eklenen', 'B2 by Ada: başka', 'E1 by Ada: dışarıda']);
  const third = await save(doc);
  assert.deepEqual(shown((await reopen(third)).sheets[0]), shown(sheet));
  // The model alone says the same to the writer of new files.
  const built = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.deepEqual(shown(built.sheets[0]), shown(sheet));
});

test('notes are sorted: a range with a conversation, and a sheet whose comments could not all be read, still refuse the sort', async () => {
  const threaded =
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="A2" dT="2024-01-01T00:00:00.00" personId="{P1}" id="{T1}"><text>İlk ileti</text></threadedComment></ThreadedComments>';
  const doc = await load(
    await drawingWorkbook({
      sheet: rows([[3], [1], [2], [9], [8]]),
      comments: commentsXml(['Ada'], comment('A2', 0, '<t>[Threaded comment]</t>') + comment('A5', 0, '<t>not</t>')),
      vml: vmlXml(vmlNote(1, 0) + vmlNote(4, 0)),
      threaded,
    }),
  );
  const sheet = doc.wb.sheets[0];
  assert.throws(() => sortA(doc, 'A1:A3'), /^RangeError: Cannot sort a range that contains threaded comments$/);
  assert.deepEqual(values(sheet, 0, 0, 2), [3, 1, 2]);
  // Rows without the conversation sort, and take their note along.
  sortA(doc, 'A4:A5');
  assert.deepEqual(values(sheet, 0, 3, 4), [8, 9]);
  assert.deepEqual(shown(sheet).filter((line) => /not$/.test(line)), ['A4 by Ada: not']);
  const saved = await save(doc);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), shown(sheet));
  assert.match(await text(saved, 'xl/threadedComments/threadedComment1.xml'), /ref="A2"/);

  // A second comment on a cell is one the reader passes over: the model cannot move what it does not hold.
  const twice = await load(
    await drawingWorkbook({ sheet: rows([[3], [1]]), comments: commentsXml(['Ada'], comment('A1', 0, '<t>bir</t>') + comment('A1', 0, '<t>iki</t>')), vml: vmlXml(vmlNote(0, 0)) }),
  );
  assert.throws(() => sortA(twice, 'A1:A2'), /^RangeError: Cannot sort in a sheet with comments: not supported yet$/);
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

/** What every package has to hold, as far as it matters here: each relationship leads to a part, each part has a content type. */
async function checkPackage(bytes: Uint8Array): Promise<void> {
  const zip = ZipArchive.read(bytes);
  const all = new Set(zip.names());
  const types = parseXml(await zip.read('[Content_Types].xml')).root;
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  for (const node of types.children) {
    if (node.type !== 'element') continue;
    if (node.name === 'Default') defaults.add((node.attrs.get('Extension') || '').toLowerCase());
    if (node.name === 'Override') overrides.add((node.attrs.get('PartName') || '').slice(1));
  }
  for (const part of overrides) assert.ok(all.has(part), `an override names ${part}, which is not in the package`);
  for (const name of all) {
    if (name === '[Content_Types].xml') continue;
    assert.ok(overrides.has(name) || defaults.has(name.slice(name.lastIndexOf('.') + 1).toLowerCase()), `${name} has no content type`);
    if (!/\.rels$/.test(name)) continue;
    const source = name.replace(/_rels\/([^/]*)\.rels$/, '$1');
    for (const rel of parseXml(await zip.read(name)).root.children) {
      if (rel.type !== 'element' || rel.attrs.get('TargetMode') === 'External') continue;
      const target = rel.attrs.get('Target') || '';
      const segments: string[] = target.charAt(0) === '/' ? [] : source.split('/').slice(0, -1);
      for (const segment of target.split('/')) {
        if (segment === '..') segments.pop();
        else if (segment !== '' && segment !== '.') segments.push(segment);
      }
      assert.ok(all.has(segments.join('/')), `${name}: leads to ${segments.join('/')}, which is not in the package`);
    }
  }
}

test('notes are sorted: after any mix of notes, sorts, copies, moved rows, undo and redo, what is shown is what is read back', async () => {
  let saves = 0;
  let sorts = 0;
  for (let seed = 1; seed <= 30; seed++) {
    const original = seed % 3 === 0 ? await drawingWorkbook({ sheet: rows([[5, 1], [3, 2], [9, 3], [1, 4], [7, 5], [2, 6]]) }) : await sortable();
    const doc = await load(original);
    const next = random(seed * 104729);
    const pick = (n: number) => Math.floor(next() * n);
    const trail: string[] = [];
    const sheet = doc.wb.sheets[0];
    const check = async (stage: string) => {
      const where = `seed ${seed} ${stage}: ${trail.join('; ')}`;
      const saved = await save(doc);
      saves++;
      await checkPackage(saved);
      const again = (await reopen(saved)).sheets[0];
      assert.deepEqual(shown(again), shown(sheet), where);
      const has = await ZipArchive.read(saved).has('xl/drawings/vmlDrawing1.vml');
      if (has) assert.equal((await text(saved, 'xl/drawings/vmlDrawing1.vml')).split('ObjectType="Note"').length - 1, shown(sheet).length, where);
      else assert.equal(shown(sheet).length, 0, where);
      for (const item of again.comments || []) assert.ok(item.box, `${where}: ${cellName(item.row, item.col)} has a box`);
      // The box of a note is beside its cell: its rows hold the row of the note.
      for (const item of again.comments || []) {
        const box = item.box!;
        if (box.from && box.to) assert.ok(box.from.row <= item.row + 1 && box.to.row >= item.row, `${where}: the box of ${cellName(item.row, item.col)} spans rows ${box.from.row}-${box.to.row}`);
      }
      // A file made from the model alone says the same.
      assert.deepEqual(shown((await (await open(await buildFromModel(doc.wb))).workbook()).sheets[0]), shown(sheet), `${where} (built)`);
    };
    for (let step = 0; step < 16; step++) {
      const roll = next();
      try {
        if (roll < 0.25) {
          const ref = cellName(pick(7), pick(3));
          trail.push(`note ${ref}`);
          note(doc, ref, `n${seed}.${step}`, next() < 0.5 ? 'Ada' : undefined);
        } else if (roll < 0.35) {
          const current = new SheetDrawings(sheet).currentComments();
          if (current.length === 0) continue;
          const target = current[pick(current.length)];
          trail.push(`delete ${cellName(target.row, target.col)}`);
          note(doc, cellName(target.row, target.col), undefined);
        } else if (roll < 0.6) {
          const top = pick(3);
          const area = { top, left: 0, bottom: top + 2 + pick(4), right: 1 + pick(2) };
          const key = { col: pick(area.right + 1), descending: next() < 0.5 };
          trail.push(`sort ${cellName(area.top, area.left)}:${cellName(area.bottom, area.right)} by ${key.col}${key.descending ? ' desc' : ''}`);
          if (doc.history.apply(sortRange(doc.wb, 0, area, [key]), undefined, undefined).length > 0) sorts++;
        } else if (roll < 0.7) {
          const from = { top: pick(4), left: 0, bottom: 0, right: pick(2) };
          from.bottom = from.top + pick(2);
          const copied = copyRange(doc.wb, 0, from);
          const at = { row: pick(6), col: pick(3) };
          trail.push(`${next() < 0.5 ? 'cut' : 'copy'} ${cellName(from.top, from.left)}:${cellName(from.bottom, from.right)} to ${cellName(at.row, at.col)}`);
          if (trail[trail.length - 1].slice(0, 3) === 'cut') doc.history.apply(cutRange(doc.wb, 0, from), undefined, undefined);
          doc.history.apply(pasteAt(doc.wb, 0, { top: at.row, left: at.col, bottom: at.row, right: at.col }, copied.text, copied).ops, undefined, undefined);
        } else if (roll < 0.85) {
          const op: Op = { type: next() < 0.5 ? 'insert' : 'delete', sheet: 0, axis: next() < 0.6 ? 'row' : 'col', at: pick(5), count: 1 + pick(2) };
          trail.push(`${op.type} ${op.axis} ${op.at}+${op.count}`);
          doc.history.apply([op], undefined, undefined);
        } else if (roll < 0.95 && doc.history.canUndo) {
          trail.push('undo');
          doc.history.undo();
        } else if (doc.history.canRedo) {
          trail.push('redo');
          doc.history.redo();
        }
      } catch (error) {
        if (!(error instanceof RangeError)) throw error;
        trail.push('(refused)');
      }
      if (step % 4 === 3) await check(`step ${step}`);
    }
    await check('end');
    while (doc.history.canUndo) doc.history.undo();
    await sameAround(await save(doc), original, `seed ${seed}: everything undone: ${trail.join('; ')}`);
  }
  assert.ok(saves >= 140, `${saves} saves`);
  assert.ok(sorts >= 20, `${sorts} sorts moved rows`);
});
