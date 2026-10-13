// The rules of rebasing, one by one: an operation read through a shift that was applied
// before it, a shift read through an operation that comes after it, and the two together for
// a step. And what the protocol takes for a message.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cleanUser, contentVersion, parseClientMessage, parseServerMessage, utf8Length } from '../src/collab/protocol';
import { inverseShift, mapPosition, mapRec, mapShift, transformStep, type Carry } from '../src/collab/rebase';
import type { Rec, Shift, WireOp } from '../src/collab/types';

const rows = (ins: boolean, at: number, n: number, s = 0, name = 'Data'): Shift => ({ k: 'rc', s, axis: 'row', ins, at, n, name });
const cols = (ins: boolean, at: number, n: number, s = 0): Shift => ({ k: 'rc', s, axis: 'col', ins, at, n, name: 'Data' });
const sort = (top: number, target: number[], wide = true, left = 0, right = 5): Shift => ({ k: 'perm', s: 0, top, left, bottom: top + target.length - 1, right, target, wide });
const cell = (row: number, col: number, f?: string, s = 0): WireOp => ({ t: 'cells', s, c: [[row, col, f === undefined ? { v: { k: 'number', v: 1 } } : { v: { k: 'blank', v: null, f: { x: f } } }]] });
const del = (at: number, n: number, x: 'row' | 'col' = 'row'): Rec => ({ w: { t: 'del', s: 0, x, at, n }, nm: 'Data' });
const ins = (at: number, n: number, x: 'row' | 'col' = 'row'): Rec => ({ w: { t: 'ins', s: 0, x, at, n }, nm: 'Data' });
const at = (op: WireOp | undefined): unknown => (op && op.t === 'cells' ? [op.s, op.c.map((entry) => [entry[0], entry[1]])] : op && (op.t === 'ins' || op.t === 'del') ? [op.at, op.n] : op);
const through = (op: WireOp, shift: Shift, first = false) => mapRec({ w: op, nm: 'Data' }, shift, first);

test('cells follow rows and columns that are inserted and deleted, and go with the ones that go', () => {
  assert.deepEqual(at(through(cell(5, 2), rows(true, 5, 2)).rec?.w), [0, [[7, 2]]], 'rows put in at its row push it down');
  assert.deepEqual(at(through(cell(5, 2), rows(true, 6, 2)).rec?.w), [0, [[5, 2]]], 'rows put in below leave it');
  assert.deepEqual(at(through(cell(5, 2), rows(false, 2, 3)).rec?.w), [0, [[2, 2]]], 'rows deleted above pull it up');
  assert.equal(through(cell(5, 2), rows(false, 5, 1)).conflict, true, 'its row deleted: gone');
  assert.equal(through(cell(5, 2), rows(false, 3, 3)).conflict, true, 'the last row of a run that is deleted: gone');
  assert.deepEqual(at(through(cell(5, 2), cols(true, 2, 1)).rec?.w), [0, [[5, 3]]], 'a column put in at its column');
  assert.equal(through(cell(5, 2), cols(false, 2, 1)).conflict, true, 'its column deleted');
  assert.deepEqual(at(through(cell(5, 2), rows(true, 0, 9, 1)).rec?.w), [0, [[5, 2]]], 'rows of another sheet move nothing');
  const two: WireOp = { t: 'cells', s: 0, c: [[1, 0, 0], [5, 0, 0]] };
  const part = through(two, rows(false, 5, 1));
  assert.deepEqual([at(part.rec?.w), part.partial], [[0, [[1, 0]]], true], 'part of it goes, and that is said');
});

test('formulas are written for the cells they meant, on their sheet and by its name', () => {
  const text = (mapped: ReturnType<typeof through>): unknown => (mapped.rec && mapped.rec.w.t === 'cells' && mapped.rec.w.c[0][2] !== 0 ? mapped.rec.w.c[0][2].v?.f?.x : undefined);
  assert.equal(text(through(cell(0, 0, 'B5+C5'), rows(true, 2, 1))), 'B6+C6', 'on the sheet the rows were put into');
  assert.equal(text(through(cell(0, 0, 'B5+Data!B5', 1), rows(true, 2, 1))), 'B5+Data!B6', 'on another sheet only what names that sheet moves');
  assert.equal(text(through(cell(0, 0, 'B5'), rows(false, 4, 1))), '#REF!', 'what it read is gone');
  assert.equal(text(through(cell(0, 0, 'Data!A1+1', 1), { k: 'sName', s: 0, old: 'Data', name: 'Yeni Ad' })), "'Yeni Ad'!A1+1", 'a sheet that got another name');
  assert.equal(text(through(cell(0, 0, 'Data!A1+1', 1), { k: 'sDel', at: 0, name: 'Data' })), '#REF!+1', 'a sheet that went');
});

test('rows that go: read through rows that came and went beside them', () => {
  assert.deepEqual(at(mapRec(del(3, 3), rows(true, 1, 2)).rec?.w), [5, 3], 'rows put in above move the run');
  assert.deepEqual(at(mapRec(del(3, 3), rows(true, 6, 2)).rec?.w), [3, 3], 'rows put in below leave it');
  assert.deepEqual((mapRec(del(3, 3), rows(true, 4, 2)).recs || []).map((rec) => at(rec.w)), [[6, 2], [3, 1]], 'rows that someone else put in inside it stay: the rows below them go, then the row above them');
  assert.deepEqual(at(mapRec(del(3, 3), rows(true, 3, 2)).rec?.w), [5, 3], 'rows put in at its first row are above it');
  assert.deepEqual(at(mapRec(del(3, 3), rows(false, 4, 4)).rec?.w), [3, 1], 'what someone else deleted of it is done');
  assert.equal(mapRec(del(3, 3), rows(false, 2, 5)).noop, true, 'all of it deleted already: nothing left to do, nothing lost');
  assert.deepEqual(at(mapRec(del(3, 3), cols(false, 0, 9)).rec?.w), [3, 3], 'columns do nothing to rows');
  assert.deepEqual(at(mapRec(del(2, 2, 'col'), cols(true, 0, 1)).rec?.w), [3, 2], 'columns are the same rule');
});

test('rows that come: where they are put, also inside a run that went, and who stands first at one place', () => {
  assert.deepEqual(at(mapRec(ins(4, 1), rows(true, 4, 2), false).rec?.w), [4, 1], 'two at one place: the later one stands above');
  assert.deepEqual(at(mapRec(ins(4, 1), rows(true, 4, 2), true).rec?.w), [6, 1], 'and the earlier one below it');
  assert.deepEqual(at(mapRec(ins(4, 1), rows(false, 3, 3)).rec?.w), [3, 1], 'inside a run that went: where the run was');
  assert.deepEqual(at(mapRec(ins(4, 1), rows(false, 4, 2)).rec?.w), [4, 1], 'at the first row of a run that went');
  assert.deepEqual(at(mapRec(ins(9, 1), rows(false, 3, 3)).rec?.w), [6, 1], 'below a run that went');
});

test('what was applied beside a step is read through the step, the same way round', () => {
  // The step deleted rows 3..5; beside it rows were put in at 4: they went with the run, and are no more.
  assert.deepEqual(mapShift(rows(true, 4, 2), rows(false, 3, 3)), [rows(true, 3, 2)], 'rows put in inside a run the step deleted stay, where the run was');
  assert.deepEqual(mapShift(rows(true, 3, 2), rows(false, 3, 3)), [rows(true, 3, 2)], 'at its first row they stay');
  assert.deepEqual(mapShift(rows(true, 6, 2), rows(false, 3, 3)), [rows(true, 3, 2)], 'below it they follow');
  // The step put a row in at 4; beside it rows 3..5 went: the row that was put in stays, the rows on either side of it went.
  assert.deepEqual(mapShift(rows(false, 3, 3), rows(true, 4, 1)), [rows(false, 5, 2), rows(false, 3, 1)], 'a run that went around rows the step put in is two runs');
  assert.deepEqual(mapShift(rows(false, 3, 3), rows(true, 3, 1)), [rows(false, 4, 3)], 'rows put in above the run move it');
  assert.deepEqual(mapShift(rows(false, 3, 3), rows(false, 4, 4)), [rows(false, 3, 1)], 'two runs that went: what is left of the one');
  assert.deepEqual(mapShift(rows(false, 3, 3), rows(false, 0, 9)), [], 'nothing left');
  assert.deepEqual(mapShift(rows(true, 4, 1), rows(true, 4, 2)), [rows(true, 6, 1)], 'two at one place: the one that was applied first is below');
  assert.deepEqual(mapShift(rows(true, 4, 1), { k: 'sName', s: 0, old: 'Data', name: 'Veri' }), [rows(true, 4, 1, 0, 'Veri')], 'the name of its sheet follows');
  assert.deepEqual(mapShift(rows(true, 4, 1, 2), { k: 'sDel', at: 0, name: 'x' }), [rows(true, 4, 1, 1)], 'and its place among the sheets');
  assert.deepEqual(mapShift(rows(true, 4, 1, 2), { k: 'sDel', at: 2, name: 'x' }), [], 'a sheet that went takes it along');
});

test('a sort that was applied beside a step is read through rows the step put in or took away', () => {
  const reversed = sort(1, [3, 2, 1, 0]);
  assert.deepEqual(mapShift(reversed, rows(true, 0, 2)), [sort(3, [3, 2, 1, 0])], 'rows above move the range');
  assert.deepEqual(mapShift(reversed, rows(true, 5, 2)), [reversed], 'rows below leave it');
  // A row put in before the third row of the range: it stands there before the sort and after it.
  assert.deepEqual(mapShift(reversed, rows(true, 3, 1)), [sort(1, [4, 3, 2, 1, 0])], 'a row inside stays where it was put, the others go where the sort put them');
  // The second row of the range went: the three that are left are reversed among themselves.
  assert.deepEqual(mapShift(reversed, rows(false, 2, 1)), [sort(1, [2, 1, 0])], 'rows that went leave the sort of the others');
  assert.deepEqual(mapShift(reversed, rows(false, 0, 9)), [], 'all of them gone');
  assert.deepEqual(mapShift(reversed, cols(false, 0, 9)), [], 'its columns gone');
});

test('rows that go are found where a sort put them, in runs; half a row that was sorted is nobody\'s to delete', () => {
  // Rows 2 and 3 of a range from row 1 that was reversed are rows 3 and 2 now: one run.
  const one = mapRec(del(2, 2), sort(1, [3, 2, 1, 0]));
  assert.deepEqual(at(one.rec?.w), [2, 2], 'one run');
  // Rows 1 and 2, sorted apart.
  const apart = mapRec(del(1, 2), sort(1, [3, 0, 1, 2]));
  assert.deepEqual((apart.recs || []).map((rec) => at(rec.w)), [[4, 1], [1, 1]], 'two runs, the lower one first');
  assert.equal(mapRec(del(1, 2), sort(1, [3, 0, 1, 2], false)).conflict, true, 'a sort of some columns only: dropped');
  assert.deepEqual(at(mapRec(del(7, 2), sort(1, [3, 2, 1, 0])).rec?.w), [7, 2], 'rows outside the range are where they were');
  assert.deepEqual(at(through(cell(1, 2), sort(1, [3, 2, 1, 0])).rec?.w), [0, [[4, 2]]], 'a cell goes with its row');
  assert.deepEqual(at(through(cell(1, 9), sort(1, [3, 2, 1, 0])).rec?.w), [0, [[1, 9]]], 'unless its column was not sorted');
});

test('a step keeps or drops as a whole, and what was applied beside it comes out as the next step has to read it', () => {
  const step: Rec[] = [del(3, 3), { w: cell(5, 1), nm: undefined }];
  const beside: Carry[] = [{ sh: rows(true, 4, 2), w: { t: 'ins', s: 0, x: 'row', at: 4, n: 2 } }];
  const moved = transformStep(step, beside);
  assert.deepEqual((moved.recs || []).map((rec) => at(rec.w)), [[6, 2], [3, 1], [0, [[7, 1]]]], 'the run goes around the new rows, and the cell below it is below them');
  assert.deepEqual(moved.carry.map((item) => item.sh), [rows(true, 3, 2)], 'for what comes after the step, the rows that were put in stand where the run was');
  const lost = transformStep([{ w: cell(5, 1) }, { w: cell(9, 1) }], [{ sh: rows(false, 9, 1) }]);
  assert.deepEqual([lost.recs, lost.noop], [undefined, false], 'one operation without a place drops the step');
  const done = transformStep([del(3, 2)], [{ sh: rows(false, 0, 9) }]);
  assert.deepEqual([done.recs, done.noop], [undefined, true], 'a step with nothing left to do is no loss');
  // Undoing gives way where someone else wrote since; a step does not.
  const other: Carry[] = [{ w: cell(5, 1) }];
  assert.equal(transformStep([{ w: cell(5, 1) }], other, false).recs?.length, 1, 'a step writes over what was written');
  assert.equal(transformStep([{ w: cell(5, 1) }], other, true).recs, undefined, 'an undo does not');
  const filters: Carry[] = [{ w: { t: 'filter', s: 0, f: 0 } }];
  assert.equal(transformStep([{ w: { t: 'filter', s: 0, f: 0 } }], filters).recs, undefined, 'two filters of one sheet: the later one gives way');
  assert.equal(transformStep([{ w: { t: 'filter', s: 1, f: 0 } }], filters).recs?.length, 1, 'of two sheets: both stay');
});

test('sheets: an operation follows its sheet, and goes with it', () => {
  assert.equal(through(cell(0, 0, undefined, 2), { k: 'sIns', at: 1, name: 'x' }).rec?.w.s, 3, 'a sheet put in before it');
  assert.equal(through(cell(0, 0, undefined, 2), { k: 'sIns', at: 3, name: 'x' }).rec?.w.s, 2, 'one put in behind it');
  assert.equal(through(cell(0, 0, undefined, 2), { k: 'sDel', at: 0, name: 'x' }).rec?.w.s, 1, 'one deleted before it');
  assert.equal(through(cell(0, 0, undefined, 2), { k: 'sDel', at: 2, name: 'x' }).conflict, true, 'its own deleted');
  assert.equal(through({ t: 'sheet-', s: 2 }, { k: 'sDel', at: 2, name: 'x' }).noop, true, 'a sheet that is deleted twice is deleted');
  assert.equal(through(cell(0, 0, undefined, 0), { k: 'sMove', from: 0, to: 2 }).rec?.w.s, 2, 'a sheet that was moved');
  assert.equal(through(cell(0, 0, undefined, 1), { k: 'sMove', from: 0, to: 2 }).rec?.w.s, 0, 'and one it passed');
  assert.equal(through({ t: 'sheet+', s: 1, name: 'n' }, { k: 'sIns', at: 1, name: 'x' }, false).rec?.w.s, 1, 'two sheets put in at one place: the later one first');
  assert.equal(through({ t: 'sheet+', s: 1, name: 'n' }, { k: 'sIns', at: 1, name: 'x' }, true).rec?.w.s, 2, 'the earlier one behind it');
  assert.equal(through({ t: 'move', s: 0, to: 2 }, { k: 'sDel', at: 1, name: 'x' }).rec?.w.t === 'move' && (through({ t: 'move', s: 0, to: 2 }, { k: 'sDel', at: 1, name: 'x' }).rec?.w as { to: number }).to, 1, 'where a sheet is moved to follows as well');
});

test('rules and charts: ranges and formulas follow; what is named by its place gives way', () => {
  const validation: WireOp = { t: 'cmd', s: 0, name: 'setValidation', a: [{ top: 4, left: 1, bottom: 5, right: 1 }, { type: 'whole', formula1: 'B2' }] };
  const moved = through(validation, rows(true, 0, 3)).rec?.w as { a: unknown[] };
  assert.deepEqual(moved.a, [{ top: 7, left: 1, bottom: 8, right: 1 }, { type: 'whole', formula1: 'B5' }], 'a validation moves with its cells');
  assert.equal(through(validation, rows(false, 4, 2)).conflict, true, 'all its cells gone');
  assert.equal(through({ t: 'cmd', s: 0, name: 'removeConditionalRule', a: [0] }, rows(true, 0, 1)).conflict, true, 'a rule named by its place, on a sheet whose rows moved');
  assert.equal(through({ t: 'cmd', s: 0, name: 'removeConditionalRule', a: [0] }, rows(true, 0, 1, 1)).rec !== undefined, true, 'rows of another sheet do not matter to it');
  assert.equal(through({ t: 'cmd', s: 0, name: 'addChart', a: [{ type: 'column' }] }, rows(true, 0, 1, 1)).conflict, true, 'a new chart reads cells that may have moved anywhere');
  assert.equal(through({ t: 'cmd', s: 0, name: 'removeChart', a: [0] }, rows(true, 0, 1)).rec !== undefined, true, 'taking a chart away reads none');
  const beside: Carry[] = [{ w: { t: 'cmd', s: 0, name: 'addChart', a: [{ type: 'line' }] } }];
  assert.equal(transformStep([{ w: { t: 'cmd', s: 0, name: 'removeChart', a: [0] } }], beside).recs, undefined, 'a chart named by its place gives way to another change of the charts');
  assert.equal(transformStep([{ w: { t: 'cmd', s: 0, name: 'addChart', a: [{ type: 'pie' }] } }], beside).recs?.length, 1, 'two new charts are two charts');
});

test('a shift taken back, and a position read through a shift', () => {
  for (const shift of [rows(true, 3, 2), rows(false, 3, 2), sort(1, [2, 0, 1]), { k: 'sIns', at: 1, name: 'n' } as Shift, { k: 'sMove', from: 0, to: 2 } as Shift, { k: 'sName', s: 0, old: 'a', name: 'b' } as Shift]) {
    assert.deepEqual(inverseShift(inverseShift(shift)), shift, `taken back twice is the shift: ${shift.k}`);
  }
  assert.deepEqual(inverseShift(sort(1, [2, 0, 1])), sort(1, [1, 2, 0]), 'a sort taken back');
  assert.deepEqual(mapPosition({ s: 0, row: 5, col: 1 }, rows(false, 5, 1)), { s: 0, row: 5, col: 1 }, 'a selection on a row that went is on the row that took its place');
  assert.equal(mapPosition({ s: 0, row: 5, col: 1 }, rows(false, 5, 1), true), undefined, 'a cell that is typed in is gone with it');
  assert.equal(mapPosition({ s: 1, row: 5, col: 1 }, { k: 'sDel', at: 1, name: 'x' }), undefined, 'and with its sheet');
});

test('messages: what is one and what is none', () => {
  const none = (text: string, read: (text: string) => unknown, what: string): void => assert.throws(() => read(text), { name: 'WireError' }, what);
  const hello = { t: 'hello', rev: 3, from: 1, ep: 'room1', you: 'c1', ro: false, v: [], peers: [] };
  assert.equal((parseServerMessage(JSON.stringify(hello)) as { rev: number }).rev, 3, 'a hello');
  none(JSON.stringify({ ...hello, from: 4 }), parseServerMessage, 'a log that starts behind where the room is');
  none(JSON.stringify({ ...hello, rev: -1 }), parseServerMessage, 'a negative step');
  none(JSON.stringify({ ...hello, rev: 1.5 }), parseServerMessage, 'half a step');
  none(JSON.stringify({ ...hello, you: '' }), parseServerMessage, 'nobody');
  none(JSON.stringify({ t: 'op', rev: 1, id: '', from: 'c', ops: [] }), parseServerMessage, 'a step without a name');
  none(JSON.stringify({ t: 'op', rev: 1, id: 'x'.repeat(101), from: 'c', ops: [] }), parseServerMessage, 'a name that is too long');
  none(JSON.stringify({ t: 'op', rev: 1, id: 'a:1', from: 'c', ops: {} }), parseServerMessage, 'operations that are no list');
  none('[1]', parseServerMessage, 'no object');
  none('{', parseServerMessage, 'no JSON');
  none(JSON.stringify({ t: 'saved', from: 'c', v: [] }), parseServerMessage, 'a save of no version');
  assert.equal((parseServerMessage(JSON.stringify({ t: 'saved', from: 'c', v: ['a'], rev: -2 })) as { rev?: number }).rev, undefined, 'a step that is none is left out of "saved"');
  assert.deepEqual(parseServerMessage('{"t":"sync"}'), { t: 'sync' }, 'sync');
  none(JSON.stringify({ t: 'join', v: ['x'.repeat(201)] }), parseClientMessage, 'a version that is too long');
  none(JSON.stringify({ t: 'join', v: Array(9).fill('v') }), parseClientMessage, 'too many versions');
  none(JSON.stringify({ t: 'join', v: [], key: 'a b' }), parseClientMessage, 'a key with a space');
  none(JSON.stringify({ t: 'join', v: [], pv: 0 }), parseClientMessage, 'version 0 of the contract');
  none(JSON.stringify({ t: 'join', v: [], room: 'a#b' }), parseClientMessage, 'a room with "#"');
  none(JSON.stringify({ t: 'op', base: 0, id: 'a:1', ops: [] }), parseClientMessage, 'a step of nothing');
  none(JSON.stringify({ t: 'op', base: 0, id: 'a:1', ops: Array(10002).fill(0) }), parseClientMessage, 'a step of too much');
  assert.equal((parseClientMessage(JSON.stringify({ t: 'op', base: 0, id: 'a:1', ops: Array(10001).fill(0) })) as { ops: unknown[] }).ops.length, 10001, 'ten thousand operations and their formats');
  assert.deepEqual(cleanUser({ name: '  A\u0001d a  ', color: '#12ab34' }), { name: 'A d a', color: '#12AB34' }, 'a name is one line of text, a colour six hex digits');
  assert.deepEqual(cleanUser({ name: '\u0000', color: 'red' }), {}, 'what is neither is left out');
  assert.equal(cleanUser({ name: 'x'.repeat(100) }).name?.length, 64, 'a name has a length');
  assert.deepEqual([utf8Length('a'), utf8Length('ç'), utf8Length('€'), utf8Length('😀')], [1, 2, 3, 4], 'bytes of UTF-8');
  assert.notEqual(contentVersion(new Uint8Array([1, 2, 3])), contentVersion(new Uint8Array([1, 2, 4])), 'other bytes, other version');
  assert.match(contentVersion(new Uint8Array([1, 2, 3])), /^h:3:[0-9a-f]{16}$/, 'the length and sixteen hex digits');
});

test('the edges of the rules', () => {
  assert.equal(through(cell(0, 0, undefined, 1), { k: 'sIns', at: 1, name: 'x' }).rec?.w.s, 2, 'a sheet put in at the place of its sheet moves it');
  assert.equal(through({ t: 'sheet+', s: 2, name: 'n' }, { k: 'sDel', at: 0, name: 'x' }).rec?.w.s, 1, 'a new sheet follows a sheet that went before it');
  const note: WireOp = { t: 'notes', s: 0, n: [[1, 2, { x: 'n' }]] };
  const noted = through(note, sort(1, [3, 2, 1, 0])).rec?.w as { n: [number, number, unknown][] };
  assert.deepEqual([noted.n[0][0], noted.n[0][1]], [4, 2], 'a note goes with the row a sort moved');
  assert.equal(through({ t: 'rows', s: 0, r: [[5, { hid: true }]] }, rows(false, 5, 1)).conflict, true, 'the height of a row that went is a change without a place');
  assert.equal(through({ t: 'merges', s: 0, add: [[4, 0, 5, 0]], del: [] }, rows(false, 5, 1)).conflict, true, 'a merged range that is one cell after rows went is none');
  const filter: WireOp = { t: 'filter', s: 0, f: { range: { top: 2, left: 0, bottom: 9, right: 3 }, columns: [] } };
  assert.equal(through(filter, rows(false, 2, 1)).conflict, true, 'a filter whose header row went');
  assert.equal((through(filter, rows(true, 0, 2)).rec?.w as { f: { range: { top: number } } }).f.range.top, 4, 'a filter follows rows put in above it');
  assert.equal(mapRec(del(3, 1), { k: 'sName', s: 0, old: 'Data', name: 'Veri' }).rec?.nm, 'Veri', 'rows that go are read by the name their sheet has now');
  assert.equal((through({ t: 'move', s: 0, to: 1 }, { k: 'sIns', at: 1, name: 'x' }).rec?.w as { to: number }).to, 2, 'where a sheet is moved to follows a sheet put in there');
  assert.equal(through({ t: 'move', s: 2, to: 1 }, { k: 'sDel', at: 1, name: 'x' }).noop, true, 'a sheet that is where it was to go has nowhere to go');
  const sorting: WireOp = { t: 'sort', s: 0, g: [1, 0, 5, 3], o: [4, 3, 2, 1, 0] };
  assert.equal(transformStep([{ w: sorting }], [{ w: { t: 'sort', s: 0, g: [3, 0, 8, 3], o: [0, 1, 2, 3, 4, 5] } }]).recs, undefined, 'two sorts that meet: the later one gives way');
  assert.equal(transformStep([{ w: sorting }], [{ w: { t: 'sort', s: 0, g: [6, 0, 8, 3], o: [0, 1, 2] } }]).recs?.length, 1, 'two that do not meet both stay');
  assert.equal(through(sorting, rows(true, 3, 1)).conflict, true, 'a sort with rows put into its range has no order to keep');
  assert.deepEqual((through(sorting, rows(true, 1, 2)).rec?.w as { g: number[] }).g, [3, 0, 7, 3], 'rows put in at its first row move it');
  for (const op of [{ t: 'name', s: 0, name: 'a' }, { t: 'tab', s: 0, color: 0 }, { t: 'state', s: 0, state: 'hidden' }] as WireOp[]) {
    assert.equal(transformStep([{ w: op }], [{ w: op }], true).recs, undefined, `the undoing of "${op.t}" gives way to someone who did the same since`);
    assert.equal(transformStep([{ w: op }], [{ w: op }], false).recs?.length, 1, `a step of "${op.t}" does not`);
  }
});

test('what is not known to come out right is dropped, either way round: a sort and rows, merged cells and rows, merged cells and a cell', () => {
  const sorted: Rec = { w: { t: 'sort', s: 0, g: [1, 0, 3, 5], o: [2, 0, 1] }, wide: true };
  const merged: Rec = { w: { t: 'merges', s: 0, add: [[2, 0, 3, 1]], del: [] } };
  const carried = (shift: Shift): Carry[] => [{ sh: shift }];
  // A sort beside rows or columns of its sheet, whichever of the two waits: dropped, and so is what was made after it.
  for (const [name, step, beside] of [
    ['rows that wait beside a sort', [ins(2, 1)], carried(sort(1, [2, 0, 1]))],
    ['rows that go beside a sort of some columns', [del(2, 1)], carried(sort(1, [2, 0, 1], false, 0, 1))],
    ['columns that wait beside a sort', [ins(1, 1, 'col')], carried(sort(1, [2, 0, 1]))],
    ['a sort that waits beside rows', [sorted], carried(rows(true, 2, 1))],
    ['a sort that waits beside columns', [sorted], carried(cols(false, 1, 1))],
  ] as [string, Rec[], Carry[]][]) {
    const done = transformStep(step, beside);
    assert.deepEqual([done.recs, done.poison, done.noop], [undefined, true, false], `${name}: dropped, with what follows it`);
  }
  // On another sheet they have nothing to do with each other.
  assert.equal(transformStep([ins(2, 1)], [{ sh: { ...sort(1, [2, 0, 1]), s: 1 } as Shift }]).recs !== undefined, true, 'a sort of another sheet moves nothing here');
  // Merged cells beside rows of their sheet: the merge is dropped; rows that wait beside a merge take what follows with them.
  const mergeBesideRows = transformStep([merged], carried(rows(true, 3, 1)));
  assert.deepEqual([mergeBesideRows.recs, mergeBesideRows.poison === true], [undefined, false], 'a merge that waits beside rows is dropped alone');
  const rowsBesideMerge = transformStep([ins(3, 1)], [{ w: merged.w }]);
  assert.deepEqual([rowsBesideMerge.recs, rowsBesideMerge.poison], [undefined, true], 'rows that wait beside a merge are dropped with what follows');
  assert.equal(transformStep([merged], carried(rows(true, 3, 1, 1))).recs !== undefined, true, 'rows of another sheet leave a merge alone');
  // Merged cells and a cell that someone else wrote inside them.
  assert.equal(transformStep([merged], [{ w: cell(3, 1) }]).recs, undefined, 'a merge over a cell someone else wrote meanwhile is dropped');
  assert.equal(transformStep([merged], [{ w: cell(9, 4) }]).recs !== undefined, true, 'a cell elsewhere leaves it alone');
  // A sheet that is moved, beside anything that names a sheet by its place.
  const moved = transformStep([{ w: cell(1, 1) }], [{ sh: { k: 'sMove', from: 0, to: 1 } }]);
  assert.deepEqual([moved.recs, moved.poison], [undefined, true], 'a cell beside a sheet that was moved is dropped with what follows');
});

test('an undoing gives way whole, a deletion grows by nothing, and what was said is said', () => {
  // Rows that someone else put in inside a run that is to go: the run goes around them, in two parts, and is no longer.
  const around = transformStep([del(2, 4)], [{ sh: rows(true, 4, 3) }]);
  assert.deepEqual((around.recs || []).map((rec) => at(rec.w)), [[7, 2], [2, 2]], 'four rows go, in two runs, and the three that came stay');
  assert.equal((around.recs || []).reduce((sum, rec) => sum + (rec.w as { n: number }).n, 0), 4, 'no more rows than were to go');
  // The undoing of rows beside rows of someone else on the sheet: not done at all.
  assert.equal(transformStep([del(2, 1)], [{ sh: rows(true, 9, 1) }], true).recs, undefined, 'rows are not taken back after someone else put rows into the sheet');
  assert.equal(transformStep([del(2, 1)], [{ sh: rows(true, 9, 1) }], false).recs !== undefined, true, 'a deletion that is no undoing is');
  // A formula that read what someone else took away is written as the error it became, and that is said.
  const broken = transformStep([{ w: cell(1, 5, 'B4+1') }], [{ sh: rows(false, 3, 1) }]);
  assert.deepEqual([broken.partial, JSON.stringify(broken.recs).indexOf('#REF!') >= 0], [true, true], 'the formula is an error now, and it is said');
  assert.equal(transformStep([{ w: cell(1, 5, 'B9+1') }], [{ sh: rows(false, 3, 1) }]).partial, false, 'one that reads a row that only moved says nothing');
  // Rows that were to go, some of which someone else took first and then brought back: they stay, and it is said.
  const first = transformStep([del(2, 3)], [{ sh: rows(false, 3, 1) }]);
  assert.deepEqual([(first.recs || []).map((rec) => at(rec.w)), first.partial], [[[2, 2]], false], 'the row that is gone already is not deleted again, and nothing is said yet');
  const second = transformStep(first.recs as Rec[], [{ sh: rows(true, 3, 1) }]);
  assert.equal(second.partial, true, 'when it comes back, it is said that not everything that was to go is gone');
  // Rows put in where a run went stand behind the rows that come back there, as they stood behind the run.
  const folded = transformStep([ins(4, 1)], [{ sh: rows(false, 2, 2) }]);
  assert.deepEqual((folded.recs || []).map((rec) => at(rec.w)), [[2, 1]], 'at the place of the run');
  const unfolded = transformStep(folded.recs as Rec[], [{ sh: rows(true, 2, 2) }]);
  assert.deepEqual((unfolded.recs || []).map((rec) => at(rec.w)), [[4, 1]], 'and behind the run again when it comes back');
});
