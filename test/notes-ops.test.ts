// Notes (legacy comments) as operations: adding, editing and deleting one,
// undo and redo, what is refused, and where a note is after rows and columns moved.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createSheet, emptyStyleSheet, type Sheet, type SheetComment, type Workbook } from '../src/index';
import { DEFAULT_NOTE_AUTHOR, MAX_NOTE_TEXT, commentAtCell, noteProblem, setComment } from '../src/ops/notes';
import { History, applyOp, type Op } from '../src/ops/ops';
import { sortRange } from '../src/ops/sort';
import { SheetDrawings, commentCellNow } from '../src/render/sheet-drawings';

function book(comments?: SheetComment[]): Workbook {
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  if (comments) sheet.comments = comments;
  return { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

const shown = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => `${placed.row},${placed.col}:${placed.comment.text}`)
    .sort();

test('notes: a note is added with its author, edited and deleted, each as one step that undo takes back exactly', () => {
  const wb = book();
  const sheet = wb.sheets[0];
  const history = new History(wb);
  history.apply(setComment(wb, 0, 1, 1, 'first', { author: 'Ada' }), undefined, undefined);
  assert.deepEqual(sheet.comments, [{ row: 1, col: 1, text: 'first', author: 'Ada', since: 0 }]);
  assert.equal(sheet.commentsModified, true);
  assert.equal(sheet.modified, undefined, 'no cell changed');
  const added = sheet.comments;

  history.apply(setComment(wb, 0, 1, 1, 'second'), undefined, undefined);
  // Editing keeps who wrote the note.
  assert.deepEqual(sheet.comments, [{ row: 1, col: 1, text: 'second', author: 'Ada', since: 0 }]);
  assert.notEqual(sheet.comments, added, 'the list is replaced, not altered');
  assert.deepEqual(added, [{ row: 1, col: 1, text: 'first', author: 'Ada', since: 0 }]);

  history.apply(setComment(wb, 0, 1, 1, undefined), undefined, undefined);
  assert.equal(sheet.comments, undefined);

  history.undo();
  assert.deepEqual(sheet.comments, [{ row: 1, col: 1, text: 'second', author: 'Ada', since: 0 }]);
  history.undo();
  assert.equal(sheet.comments, added, 'undo puts the very list back');
  history.undo();
  assert.equal(sheet.comments, undefined);
  assert.equal(history.canUndo, false);
  history.redo();
  history.redo();
  history.redo();
  assert.equal(sheet.comments, undefined);
  history.undo();
  assert.deepEqual(shown(sheet), ['1,1:second']);
});

test('notes: without an author the note is by "Author"; an author is trimmed and cut at 255 characters', () => {
  const wb = book();
  applyOp(wb, setComment(wb, 0, 0, 0, 'a')[0]);
  assert.equal(DEFAULT_NOTE_AUTHOR, 'Author');
  assert.equal(wb.sheets[0].comments![0].author, 'Author');
  applyOp(wb, setComment(wb, 0, 0, 1, 'b', { author: '   ' })[0]);
  assert.equal(wb.sheets[0].comments![1].author, 'Author');
  applyOp(wb, setComment(wb, 0, 0, 2, 'c', { author: `  ${'x'.repeat(300)} ` })[0]);
  assert.equal(wb.sheets[0].comments![2].author, 'x'.repeat(255));
});

test('notes: text that is empty or only white space deletes; the same text and a note that is not there are no operation', () => {
  const wb = book([{ row: 2, col: 3, text: 'kept', author: 'A' }]);
  assert.deepEqual(setComment(wb, 0, 2, 3, 'kept'), []);
  assert.deepEqual(setComment(wb, 0, 5, 5, undefined), []);
  assert.deepEqual(setComment(wb, 0, 5, 5, ' \n\t '), []);
  const ops = setComment(wb, 0, 2, 3, '  \r\n ');
  assert.equal(ops.length, 1);
  applyOp(wb, ops[0]);
  assert.equal(wb.sheets[0].comments, undefined);
});

test('notes: line breaks are "\\n"; text is kept as typed otherwise (markup, emoji)', () => {
  const wb = book();
  applyOp(wb, setComment(wb, 0, 0, 0, 'a\r\nb\rc <b>&amp;</b> \u{1F600} ')[0]);
  assert.equal(wb.sheets[0].comments![0].text, 'a\nb\nc <b>&amp;</b> \u{1F600} ');
});

test('notes: editing a note with formatted runs leaves one plain text; its box, author and visibility stay', () => {
  const box = { kind: 'twoCell' as const, from: { col: 1, colOff: 0, row: 0, rowOff: 0 }, to: { col: 3, colOff: 0, row: 4, rowOff: 0 } };
  const wb = book([{ row: 0, col: 0, text: 'Ada:\nold', author: 'Ada', runs: [{ text: 'Ada:', font: { bold: true } }, { text: '\nold' }], visible: true, box }]);
  applyOp(wb, setComment(wb, 0, 0, 0, 'Ada:\nnew')[0]);
  assert.deepEqual(wb.sheets[0].comments, [{ row: 0, col: 0, text: 'Ada:\nnew', author: 'Ada', visible: true, box }]);
});

test('notes: what is refused: too long a text, a threaded comment, a cell outside the sheet, a sheet that is not a worksheet', () => {
  const wb = book([{ row: 0, col: 0, text: 'talk', threaded: true, replies: [{ text: 'yes' }] }, { row: 1, col: 0, text: 'note' }]);
  const before = wb.sheets[0].comments;
  assert.equal(MAX_NOTE_TEXT, 32767);
  assert.throws(() => setComment(wb, 0, 3, 3, 'x'.repeat(32768)), RangeError);
  assert.equal(setComment(wb, 0, 3, 3, 'x'.repeat(32767)).length, 1);
  assert.throws(() => setComment(wb, 0, 0, 0, 'mine'), /Cannot change the comment in A1: it is a threaded comment/);
  assert.throws(() => setComment(wb, 0, 0, 0, undefined), /Cannot change the comment in A1: it is a threaded comment/);
  assert.match(noteProblem(wb, 0, 0, 0) || '', /threaded comment/);
  assert.equal(noteProblem(wb, 0, 1, 0), undefined);
  assert.throws(() => setComment(wb, 0, -1, 0, 'x'), RangeError);
  assert.throws(() => setComment(wb, 0, 0, 16384, 'x'), RangeError);
  assert.throws(() => setComment(wb, 0, 1.5, 0, 'x'), RangeError);
  assert.throws(() => setComment(wb, 3, 0, 0, 'x'), RangeError);
  wb.sheets[0].kind = 'chartsheet';
  assert.throws(() => setComment(wb, 0, 5, 5, 'x'), RangeError);
  assert.equal(wb.sheets[0].comments, before);
});

test('notes: in a merged cell the note belongs to its top-left cell', () => {
  const wb = book();
  wb.sheets[0].merges = [{ top: 1, left: 1, bottom: 2, right: 3 }];
  applyOp(wb, setComment(wb, 0, 2, 3, 'merged')[0]);
  assert.deepEqual(shown(wb.sheets[0]), ['1,1:merged']);
  assert.equal(commentAtCell(wb.sheets[0], 2, 2)!.comment.text, 'merged');
  applyOp(wb, setComment(wb, 0, 2, 2, undefined)[0]);
  assert.equal(wb.sheets[0].comments, undefined);
});

test('notes: a note added after rows moved stands where it was put, and moves with what comes after', () => {
  const wb = book([{ row: 4, col: 0, text: 'file' }]);
  const sheet = wb.sheets[0];
  const history = new History(wb);
  history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }], undefined, undefined);
  assert.deepEqual(shown(sheet), ['6,0:file']);
  history.apply(setComment(wb, 0, 3, 1, 'new'), undefined, undefined);
  assert.deepEqual(sheet.comments![1], { row: 3, col: 1, text: 'new', author: 'Author', since: 1 });
  assert.deepEqual(shown(sheet), ['3,1:new', '6,0:file']);
  // The note of the file is found where it is now, and stays a comment of the file when edited.
  history.apply(setComment(wb, 0, 6, 0, 'file2'), undefined, undefined);
  assert.deepEqual(sheet.comments![0], { row: 4, col: 0, text: 'file2' });
  history.apply([{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }], undefined, undefined);
  assert.deepEqual(shown(sheet), ['3,2:new', '6,1:file2']);
  history.apply([{ type: 'delete', sheet: 0, axis: 'row', at: 3, count: 1 }], undefined, undefined);
  // The note went with its row; the cell that is there now has none and can take one.
  assert.deepEqual(shown(sheet), ['5,1:file2']);
  assert.equal(commentAtCell(sheet, 3, 2), undefined);
  history.apply(setComment(wb, 0, 3, 2, 'again'), undefined, undefined);
  assert.deepEqual(shown(sheet), ['3,2:again', '5,1:file2']);
  history.undo();
  history.undo();
  assert.deepEqual(shown(sheet), ['3,2:new', '6,1:file2']);
  history.undo();
  history.undo();
  history.undo();
  history.undo();
  assert.deepEqual(sheet.comments, [{ row: 4, col: 0, text: 'file' }]);
  assert.equal(sheet.structuralEdits, undefined);
});

test('notes: commentCellNow() applies only the structural edits made after a note was added', () => {
  const edits = [
    { kind: 'insert' as const, axis: 'row' as const, at: 0, count: 2 },
    { kind: 'insert' as const, axis: 'col' as const, at: 0, count: 1 },
  ];
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '' }, edits), { row: 3, col: 2 });
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '', since: 1 }, edits), { row: 1, col: 2 });
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '', since: 2 }, edits), { row: 1, col: 1 });
  // A count that is none (from outside) moves the note by nothing rather than by everything.
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '', since: 9 }, edits), { row: 1, col: 1 });
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '', since: -3 }, edits), { row: 3, col: 2 });
  assert.deepEqual(commentCellNow({ row: 1, col: 1, text: '' }, undefined), { row: 1, col: 1 });
  assert.equal(commentCellNow({ row: 1, col: 1, text: '', since: 0 }, [{ kind: 'delete', axis: 'row', at: 1, count: 1 }]), undefined);
});

test('notes: a sheet that got a note is sorted, and the note goes with its row', () => {
  const wb = book();
  const sheet = wb.sheets[0];
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }, { row: 1, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] });
  const range = { top: 0, left: 0, bottom: 1, right: 0 };
  const sorted = sortRange(wb, 0, range, [{ col: 0 }]);
  applyOp(wb, setComment(wb, 0, 0, 0, 'n')[0]);
  const before = sheet.comments;
  const undo = sorted.map((op) => applyOp(wb, op));
  assert.equal(sheet.rows.get(0)!.cells[0].value, 1);
  assert.deepEqual(shown(sheet), ['1,0:n']);
  // A note added since has no cell in the file to remember.
  assert.equal(sheet.comments![0].from, undefined);
  for (const op of undo.reverse()) applyOp(wb, op);
  assert.equal(sheet.comments, before);
  assert.deepEqual(shown(sheet), ['0,0:n']);
});

test('notes: clearing cells and deleting a sheet leave the model of notes as it is; "setComments" is its own inverse', () => {
  const wb = book();
  const op = setComment(wb, 0, 0, 0, 'n')[0];
  const undo = applyOp(wb, op);
  assert.deepEqual(undo, { type: 'setComments', sheet: 0, comments: undefined });
  const list = wb.sheets[0].comments;
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: undefined }] });
  assert.equal(wb.sheets[0].comments, list);
  const redo = applyOp(wb, undo);
  assert.equal((redo as Extract<Op, { type: 'setComments' }>).comments, list);
  assert.throws(() => applyOp(wb, { type: 'setComments', sheet: 0, comments: 'x' as unknown as SheetComment[] }), RangeError);
});

test('notes: no more notes than a sheet is read with', () => {
  const many: SheetComment[] = [];
  for (let i = 0; i < 100000; i++) many.push({ row: i, col: 0, text: 'x' });
  const wb = book(many);
  assert.throws(() => setComment(wb, 0, 0, 5, 'one more'), /too many/i);
  // Editing and deleting still work.
  assert.equal(setComment(wb, 0, 7, 0, 'edited').length, 1);
  assert.equal(setComment(wb, 0, 7, 0, undefined).length, 1);
});

// --- review ------------------------------------------------------------------------------------

test('review: text that differs from the note only in how its line breaks are written is no change', () => {
  const runs = [{ text: 'Ada:', font: { bold: true } }, { text: '\r\nbody' }];
  const wb = book([{ row: 0, col: 0, text: 'Ada:\r\nbody', author: 'Ada', runs }]);
  // What a text field hands back for it.
  assert.deepEqual(setComment(wb, 0, 0, 0, 'Ada:\nbody'), []);
  assert.deepEqual(setComment(wb, 0, 0, 0, 'Ada:\r\nbody'), []);
  assert.equal(setComment(wb, 0, 0, 0, 'Ada:\nbody!').length, 1);
});

test('review: one rule for text: undefined deletes; the text the note has is no change; other blank text deletes', () => {
  const wb = book([{ row: 0, col: 0, text: '' }, { row: 1, col: 0, text: '  ' }, { row: 2, col: 0, text: 'x' }]);
  assert.deepEqual(setComment(wb, 0, 0, 0, ''), [], 'an empty note keeps being one');
  assert.deepEqual(setComment(wb, 0, 1, 0, '  '), []);
  assert.equal(setComment(wb, 0, 0, 0, undefined).length, 1);
  assert.equal(setComment(wb, 0, 1, 0, '').length, 1);
  assert.equal(setComment(wb, 0, 2, 0, '').length, 1);
  assert.equal(setComment(wb, 0, 2, 0, ' \n').length, 1);
  applyOp(wb, setComment(wb, 0, 0, 0, undefined)[0]);
  assert.deepEqual(wb.sheets[0].comments!.map((item) => item.row), [1, 2]);
  // A cell without a note gets none from blank text.
  assert.deepEqual(setComment(wb, 0, 9, 9, ''), []);
});

test('review: a note whose cell was deleted does not count towards the limit, and comments the reader could not show block sorting', () => {
  const wb = book([{ row: 5, col: 0, text: 'gone' }]);
  const sheet = wb.sheets[0];
  sheet.anchored = ['comments', 'vmlDrawing'];
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }, { row: 1, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] });
  const range = { top: 0, left: 0, bottom: 1, right: 0 };
  // (A note that is there does not block a sort either: it goes with its row, or stays outside the range.)
  applyOp(wb, { type: 'delete', sheet: 0, axis: 'row', at: 5, count: 1 });
  assert.deepEqual(shown(sheet), []);
  sortRange(wb, 0, range, [{ col: 0 }]).forEach((op) => applyOp(wb, op));
  assert.equal(sheet.rows.get(0)!.cells[0].value, 1);

  // Deleted by hand: the same.
  const wb2 = book([{ row: 5, col: 0, text: 'x' }]);
  wb2.sheets[0].anchored = ['comments', 'vmlDrawing'];
  applyOp(wb2, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }, { row: 1, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] });
  applyOp(wb2, setComment(wb2, 0, 5, 0, undefined)[0]);
  sortRange(wb2, 0, range, [{ col: 0 }]).forEach((op) => applyOp(wb2, op));
  assert.equal(wb2.sheets[0].rows.get(0)!.cells[0].value, 1);
  // Comments the reader could not show may still be in the file: such a sheet is not sorted.
  const wb3 = book();
  wb3.sheets[0].anchored = ['comments'];
  wb3.sheets[0].warnings = ['Comments: 1 comment(s) could not be read and are not shown'];
  applyOp(wb3, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 2, style: 0 } }, { row: 1, col: 0, cell: { kind: 'number', value: 1, style: 0 } }] });
  assert.throws(() => sortRange(wb3, 0, range, [{ col: 0 }]).forEach((op) => applyOp(wb3, op)), /Cannot sort in a sheet with comments/);

  // The limit counts the notes that are there.
  const many: SheetComment[] = [];
  for (let i = 0; i < 100000; i++) many.push({ row: i, col: 0, text: 'x' });
  const full = book(many);
  applyOp(full, { type: 'delete', sheet: 0, axis: 'row', at: 0, count: 10 });
  assert.equal(setComment(full, 0, 0, 5, 'room again').length, 1);
});

test('review: a note anywhere inside a merged cell is the note of the merged cell; no second one is added', () => {
  const wb = book([{ row: 1, col: 1, text: 'inner', author: 'A' }]);
  const sheet = wb.sheets[0];
  sheet.merges = [{ top: 0, left: 0, bottom: 2, right: 2 }];
  assert.equal(commentAtCell(sheet, 0, 0)!.comment.text, 'inner');
  assert.equal(commentAtCell(sheet, 2, 2)!.comment.text, 'inner');
  assert.equal(commentAtCell(sheet, 3, 3), undefined);
  applyOp(wb, setComment(wb, 0, 0, 0, 'edited')[0]);
  assert.deepEqual(sheet.comments, [{ row: 1, col: 1, text: 'edited', author: 'A' }]);
  applyOp(wb, setComment(wb, 0, 2, 0, undefined)[0]);
  assert.equal(sheet.comments, undefined);
  // The top-left one is preferred when there are several.
  sheet.comments = [{ row: 1, col: 1, text: 'inner' }, { row: 0, col: 0, text: 'corner' }];
  assert.equal(commentAtCell(sheet, 2, 2)!.comment.text, 'corner');
});

test('review: a sheet whose comment parts cannot take notes refuses every change of a note, with the reason', () => {
  for (const [kind, reason] of [
    ['drawing', /its legacy drawing cannot be read/],
    ['shared', /its comment parts are shared with another sheet/],
    ['parts', /it has more than one comments part/],
  ] as const) {
    const wb = book([{ row: 0, col: 0, text: 'x' }]);
    wb.sheets[0].notesBlocked = kind;
    assert.match(noteProblem(wb, 0, 5, 5) || '', reason);
    assert.throws(() => setComment(wb, 0, 5, 5, 'new'), reason);
    assert.throws(() => setComment(wb, 0, 0, 0, 'edit'), reason);
    assert.throws(() => setComment(wb, 0, 0, 0, undefined), reason);
  }
});
