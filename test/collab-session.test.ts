// Collaborative editing, case by case: what becomes of two changes that were made at once,
// of a connection that comes and goes, of a file that was saved on the way, and of what
// arrives that is no change at all.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getCell, type Workbook } from '../src/index';
import { CollabUnsupportedError } from '../src/collab/types';
import { contentVersion, parseClientMessage, parseServerMessage } from '../src/collab/protocol';
import { decodeStep } from '../src/collab/wire';
import { addChart, removeChart, sheetCharts } from '../src/ops/charts';
import { copyRange, pasteAt, setCellInput } from '../src/ops/commands';
import { deleteRows, setAutoFilter } from '../src/ops/filter';
import { formatRange, mergeRange } from '../src/ops/format';
import { cellNow, setComment } from '../src/ops/notes';
import type { Op } from '../src/ops/ops';
import { addConditionalRule, conditionalRules, removeConditionalRule, setValidation, validationsIn } from '../src/ops/rules';
import { copySheet, deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { insertPageBreak, setPageSetup } from '../src/ops/page-setup';
import { sortRange } from '../src/ops/sort';
import { Net, act, assertSame, join, saveOf, seeded, settleAll, startFile, type Member } from './collab-fixtures';
import { PROTOCOL_VERSION } from '../src/collab/room';

const value = (wb: Workbook, sheet: number, row: number, col: number): unknown => {
  const cell = getCell(wb.sheets[sheet], row, col);
  return cell ? cell.value : undefined;
};
const formula = (wb: Workbook, sheet: number, row: number, col: number): string | undefined => getCell(wb.sheets[sheet], row, col)?.formula?.text;
const codes = (member: Member): string[] => member.notices.map((notice) => notice.code);

/** Two participants on the starting file, both online. */
async function two(seed = 1): Promise<{ net: Net; random: () => number; a: Member; b: Member }> {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(seed);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  return { net, random, a, b };
}

/** Lets what `a` did reach the relay first, then everything else. */
function firstThen(net: Net, random: () => number, first: Member, then: Member, make: () => void): void {
  then.transport.suspend();
  net.drain(random);
  make();
  net.drain(random);
  then.transport.resume();
  net.drain(random);
}

const insertRows = (sheet: number, at: number, count: number): Op[] => [{ type: 'insert', sheet, axis: 'row', at, count }];
const insertColumns = (sheet: number, at: number, count: number): Op[] => [{ type: 'insert', sheet, axis: 'col', at, count }];

test('a value written into a row that someone else deleted is dropped, and its author is told', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 3, 1, 'late'));
  act(a, () => deleteRows(a.wb, 0, 3, 1));
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assertSame([a, b], 'after the row went');
  assert.deepEqual(codes(b), ['dropped'], 'b was told');
  assert.deepEqual(codes(a), [], 'a was told nothing');
  assert.equal(value(a.wb, 0, 3, 0), 'Row 4', 'the row below moved up');
  assert.equal(b.session.canUndo, false, 'the step that was dropped is not one to undo');
});

test('a value written below rows that someone else inserted lands in the row it was written in', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setCellInput(b.wb, 0, 5, 1, 'kept'));
    act(a, () => insertRows(0, 2, 2));
  });
  assertSame([a, b], 'after rows came');
  assert.equal(value(a.wb, 0, 7, 1), 'kept', 'two rows further down');
  assert.equal(value(a.wb, 0, 7, 0), 'Row 5', 'beside the row it belongs to');
  assert.deepEqual(codes(b), [], 'nothing was dropped');
});

test('a formula written beside a column that someone else inserted reads the cells it meant', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setCellInput(b.wb, 0, 12, 0, '=B2+C2'));
    act(a, () => insertColumns(0, 1, 1));
  });
  assertSame([a, b], 'after a column came');
  assert.equal(formula(a.wb, 0, 12, 0), 'C2+D2', 'the references moved with their cells');
});

test('a formula that names a sheet someone else renamed names it by its new name', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setCellInput(b.wb, 0, 12, 0, '=Other!A1+1'));
    act(a, () => renameSheet(a.wb, 1, 'Depo'));
  });
  assertSame([a, b], 'after the rename');
  assert.equal(formula(a.wb, 0, 12, 0), 'Depo!A1+1', 'the formula follows the name');
});

test('a value written on a sheet that someone else deleted is dropped; one on a sheet behind it lands on that sheet', async () => {
  const { net, random, a, b } = await two();
  act(a, () => insertSheet(a.wb, 2, { name: 'Third' }));
  net.drain(random);
  firstThen(net, random, a, b, () => {
    act(b, () => setCellInput(b.wb, 1, 0, 0, 'gone'));
    act(b, () => setCellInput(b.wb, 2, 0, 0, 'stays'));
    act(a, () => deleteSheet(a.wb, 1));
  });
  assertSame([a, b], 'after the sheet went');
  assert.deepEqual(codes(b), ['dropped'], 'b was told of the one change');
  assert.equal(a.wb.sheets[1].name, 'Third', 'the sheet behind it took its place');
  assert.equal(value(a.wb, 1, 0, 0), 'stays', 'and has what was written on it');
});

test('rows that two participants insert and delete at once are all inserted and deleted', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(a, () => insertRows(0, 2, 1));
    act(b, () => deleteRows(b.wb, 0, 6, 2));
    act(b, () => insertRows(0, 2, 1));
  });
  assertSame([a, b], 'after both');
  assert.equal(value(a.wb, 0, 4, 0), 'Row 2', 'two rows were inserted above the second row');
  assert.equal(value(a.wb, 0, 8, 0), 'Row 8', 'and rows 6 and 7 are gone');
});

test('rows that two participants delete at once are deleted once', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(a, () => deleteRows(a.wb, 0, 2, 3));
    act(b, () => deleteRows(b.wb, 0, 3, 3));
  });
  assertSame([a, b], 'after both');
  assert.deepEqual([value(a.wb, 0, 1, 0), value(a.wb, 0, 2, 0)], ['Row 1', 'Row 6'], 'rows 2 to 5 are gone, and nothing else');
  assert.deepEqual(codes(b), [], 'nothing is said: all that b wanted is done');
});

test('a value written into a range that someone else sorted goes with its row', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setCellInput(b.wb, 0, 1, 2, '999'));
    act(a, () => sortRange(a.wb, 0, { top: 1, left: 0, bottom: 8, right: 4 }, [{ col: 1, descending: true }]));
  });
  assertSame([a, b], 'after the sort');
  assert.equal(value(a.wb, 0, 8, 0), 'Row 1', 'the first row is the last one now');
  assert.equal(value(a.wb, 0, 8, 2), 999, 'and what was written into it went along');
  assert.equal(value(a.wb, 0, 1, 2), 44, 'the row that stands there now has its own value');
});

test('notes: two on different cells both stay, two on one cell leave the later one, one on a deleted row is dropped', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(a, () => setComment(a.wb, 0, 1, 1, 'from a', { author: 'a' }));
    act(b, () => setComment(b.wb, 0, 2, 2, 'from b', { author: 'b' }));
    act(b, () => setComment(b.wb, 0, 1, 1, 'b on the same cell', { author: 'b' }));
  });
  assertSame([a, b], 'after the notes');
  const notes = (a.wb.sheets[0].comments || []).map((comment) => `${cellNow(a.wb.sheets[0], comment)?.row},${cellNow(a.wb.sheets[0], comment)?.col}: ${comment.text}`).sort();
  assert.deepEqual(notes, ['1,1: b on the same cell', '2,2: from b'], 'one note per cell');
  firstThen(net, random, a, b, () => {
    act(b, () => setComment(b.wb, 0, 5, 0, 'on a row that goes'));
    act(a, () => deleteRows(a.wb, 0, 5, 1));
  });
  assertSame([a, b], 'after the row of a note went');
  assert.deepEqual(codes(b), ['dropped'], 'b was told');
});

test('rules: a validation follows rows that were inserted; a rule named by its place gives way to a change of the list', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setValidation(b.wb, 0, { top: 4, left: 1, bottom: 5, right: 1 }, { type: 'whole', operator: 'greaterThan', formula1: 'B2' }));
    act(a, () => insertRows(0, 0, 3));
  });
  assertSame([a, b], 'after the validation');
  const rule = validationsIn(a.wb, 0, { top: 7, left: 1, bottom: 8, right: 1 })[0];
  assert.deepEqual([rule.ranges, rule.formula1], [[{ top: 7, left: 1, bottom: 8, right: 1 }], 'B5'], 'the cells and the formula moved down');

  act(a, () => addConditionalRule(a.wb, 0, { top: 0, left: 0, bottom: 3, right: 0 }, { type: 'cellIs', operator: 'greaterThan', formulas: ['1'], format: { bold: true } }));
  net.drain(random);
  firstThen(net, random, a, b, () => {
    act(b, () => removeConditionalRule(b.wb, 0, 0));
    act(a, () => addConditionalRule(a.wb, 0, { top: 0, left: 1, bottom: 3, right: 1 }, { type: 'cellIs', operator: 'lessThan', formulas: ['5'], format: { italic: true } }));
  });
  assertSame([a, b], 'after the rules');
  assert.equal(conditionalRules(a.wb, 0).length, 2, 'the rule that was to be removed is another one now: nothing was removed');
  assert.deepEqual(codes(b), ['dropped'], 'and b was told');
});

test('charts: two that are added at once are both there; one that is removed while another is added is not the wrong one', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(a, () => addChart(a.wb, 0, { type: 'column', range: { top: 0, left: 0, bottom: 4, right: 2 }, title: 'A' }));
    act(b, () => addChart(b.wb, 0, { type: 'line', range: { top: 0, left: 0, bottom: 4, right: 2 }, title: 'B' }));
  });
  assertSame([a, b], 'after two charts');
  assert.deepEqual(sheetCharts(a.wb, 0).map((chart) => chart.title), ['A', 'B'], 'both charts, in the order of the relay');
  firstThen(net, random, a, b, () => {
    act(b, () => removeChart(b.wb, 0, 1));
    act(a, () => removeChart(a.wb, 0, 0));
  });
  assertSame([a, b], 'after one was removed twice');
  assert.deepEqual(sheetCharts(a.wb, 0).map((chart) => chart.title), ['B'], 'the chart b meant is still there: its removal gave way');
  assert.deepEqual(codes(b), ['dropped'], 'and b was told');
});

test('a merge across a range someone else merged is dropped', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => mergeRange(b.wb, 1, { top: 5, left: 0, bottom: 6, right: 2 }));
    act(a, () => mergeRange(a.wb, 1, { top: 6, left: 1, bottom: 8, right: 1 }));
  });
  assertSame([a, b], 'after the merges');
  assert.deepEqual(a.wb.sheets[1].merges, [{ top: 6, left: 1, bottom: 8, right: 1 }], 'the one the relay placed first');
  assert.deepEqual(codes(b), ['dropped'], 'the other one was dropped and said');
});

test('two filters set at once: the later one gives way', async () => {
  const { net, random, a, b } = await two();
  firstThen(net, random, a, b, () => {
    act(b, () => setAutoFilter(b.wb, 0, { top: 0, left: 0, bottom: 9, right: 2 }));
    act(a, () => setAutoFilter(a.wb, 0, { top: 0, left: 0, bottom: 9, right: 3 }));
  });
  assertSame([a, b], 'after the filters');
  assert.equal(a.wb.sheets[0].autoFilter?.range.right, 3, 'the filter of a');
  assert.deepEqual(codes(b), ['dropped'], 'b was told');
});

test('undo takes back the own step only, also with steps of others after it', async () => {
  const { net, random, a, b } = await two();
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'a1'));
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 12, 1, 'b1'));
  act(b, () => insertRows(0, 0, 2));
  net.drain(random);
  assert.equal(value(a.wb, 0, 14, 0), 'a1', 'the cell of a moved down');
  assert.notEqual(a.session.undo(), undefined, 'a undid its step');
  net.drain(random);
  assertSame([a, b], 'after the undo');
  assert.equal(value(b.wb, 0, 14, 0), undefined, 'the cell of a is empty again, where it is now');
  assert.equal(value(b.wb, 0, 14, 1), 'b1', 'what b wrote stays');
  assert.equal(a.session.canUndo, false, 'a has nothing more to undo: the steps of b are not its own');
  assert.notEqual(a.session.redo(), undefined, 'and can make its step again');
  net.drain(random);
  assertSame([a, b], 'after the redo');
  assert.equal(value(b.wb, 0, 14, 0), 'a1', 'the value is back');
});

test('undo does not take a cell back that someone else has written since', async () => {
  const { net, random, a, b } = await two();
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'first'));
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 12, 0, 'second'));
  net.drain(random);
  assert.equal(a.session.undo(), undefined, 'nothing was undone');
  assert.deepEqual(codes(a), ['undo-gone'], 'and a was told why');
  net.drain(random);
  assert.equal(value(a.wb, 0, 12, 0), 'second', 'the value of b stays');
  assertSame([a, b], 'after the undo that was refused');
});

test('undoing a deletion of rows brings the rows, their cells and the formulas that read them back', async () => {
  const { net, random, a, b } = await two();
  const before = formula(a.wb, 0, 9, 1);
  act(a, () => formatRange(a.wb, 0, { top: 3, left: 0, bottom: 3, right: 3 }, { bold: true }));
  act(a, () => deleteRows(a.wb, 0, 7, 2));
  net.drain(random);
  assert.equal(formula(b.wb, 0, 7, 1), 'SUM(B2:B7)', 'the sum got shorter');
  assert.notEqual(a.session.undo(), undefined, 'a undid the deletion');
  net.drain(random);
  assertSame([a, b], 'after the undo');
  assert.equal(value(b.wb, 0, 8, 0), 'Row 8', 'the rows are back');
  assert.equal(formula(b.wb, 0, 8, 3), 'B9+C9', 'with their formulas');
  assert.equal(formula(b.wb, 0, 9, 1), before, 'and the sum reads them again');
});

test('a deleted sheet cannot be brought back by undo, which says so', async () => {
  const { net, random, a, b } = await two();
  act(a, () => deleteSheet(a.wb, 1));
  net.drain(random);
  assert.equal(a.session.undo(), undefined, 'nothing was undone');
  assert.deepEqual(codes(a), ['undo-none'], 'and a was told');
  assertSame([a, b], 'after it');
});

test('the page setup is refused in a room, before anything is changed', async () => {
  const { a } = await two();
  const before = a.wb.sheets[0].pageSetup;
  const names = a.wb.definedNames;
  assert.throws(() => a.session.submit(setPageSetup(a.wb, 0, { orientation: 'landscape' })), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'other');
  assert.throws(() => a.session.submit(setPageSetup(a.wb, 0, { printArea: [{ top: 0, left: 0, bottom: 1, right: 1 }] })), (error: unknown) => error instanceof CollabUnsupportedError);
  assert.throws(() => a.session.submit(insertPageBreak(a.wb, 0, { row: 3 })), (error: unknown) => error instanceof CollabUnsupportedError);
  assert.equal(a.wb.sheets[0].pageSetup, before, 'the setup is as it was');
  assert.equal(a.wb.definedNames, names, 'and so are the names');
});

test('a sheet cannot be copied, and rules do not come with a paste: both are said', async () => {
  const { net, random, a, b } = await two();
  assert.throws(() => a.session.submit(copySheet(a.wb, 0)), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'copySheet', 'copying a sheet is refused');
  assert.equal(a.wb.sheets.length, 2, 'and nothing was changed');
  act(a, () => setValidation(a.wb, 0, { top: 1, left: 1, bottom: 1, right: 1 }, { type: 'whole', operator: 'greaterThan', formula1: '1' }));
  const copied = copyRange(a.wb, 0, { top: 1, left: 1, bottom: 1, right: 1 });
  const done = a.session.submit(pasteAt(a.wb, 0, { top: 12, left: 5, bottom: 12, right: 5 }, copied.text, copied).ops);
  assert.equal(done.skipped, 1, 'the rule of the copy was left out');
  net.drain(random);
  assertSame([a, b], 'after the paste');
  assert.equal(value(b.wb, 0, 12, 5), 10, 'the cell came');
  assert.equal(validationsIn(b.wb, 0, { top: 12, left: 5, bottom: 12, right: 5 }).length, 0, 'its rule did not');
});

test('a participant who was offline sends what it changed when it is back, and gets what it missed', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  assert.equal(b.session.state, 'offline', 'b knows that it is offline');
  for (let i = 0; i < 5; i++) act(b, () => setCellInput(b.wb, 1, 5 + i, 0, `offline ${i}`));
  for (let i = 0; i < 5; i++) act(a, () => setCellInput(a.wb, 1, 5 + i, 1, `online ${i}`));
  net.drain(random);
  assert.equal(b.session.pending, 5, 'five steps wait');
  assert.equal(value(b.wb, 1, 5, 1), undefined, 'b has nothing of a yet');
  b.transport.resume();
  net.drain(random);
  assert.equal(b.session.state, 'online', 'b is back');
  assert.equal(b.session.pending, 0, 'and has sent everything');
  assertSame([a, b], 'after b came back');
  assert.equal(net.server.rooms()[0].rev, 10, 'ten steps in all');
});

test('a connection that is lost after a step was sent does not send the step twice', async () => {
  const { net, random, a, b } = await two();
  act(b, () => setCellInput(b.wb, 1, 5, 0, 'once'));
  // The step reaches the relay; the answer does not reach b.
  for (let i = 0; i < 1; i++) net.step(() => 0.999);
  b.transport.suspend();
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assertSame([a, b], 'after the reconnection');
  assert.equal(net.server.rooms()[0].rev, 1, 'the step is in the order once');
  assert.equal(b.session.pending, 0, 'and b knows that it was placed');
});

test('a participant whose file is not the one of the room loads it again; without a way to load it the session closes', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(7);
  const a = await join(net, 'a', bytes);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'changed'));
  net.drain(random);
  const saved = await saveOf(a);
  a.session.announceSaved([contentVersion(saved)], a.session.savableRev);
  act(a, () => setCellInput(a.wb, 0, 12, 1, 'after the save'));
  net.drain(random);
  // Someone with a file nobody in the room knows: an older copy that was changed elsewhere.
  let loads = 0;
  const b = await join(net, 'b', bytes, { versions: ['some other version'] }, async () => {
    loads++;
    return saved;
  });
  await settleAll(net, random, [a, b]);
  assert.equal(loads, 1, 'the file was loaded again once');
  assert.equal(b.session.state, 'online', 'and b joined with it');
  assertSame([a, b], 'after the reload');
  assert.equal(value(b.wb, 0, 12, 1), 'after the save', 'what came after the save was sent to b');
  const c = await join(net, 'c', bytes, { versions: ['yet another version'] });
  net.drain(random);
  assert.equal(c.session.state, 'closed', 'without a way to load the file the session is over');
  assert.deepEqual(codes(c), ['stale'], 'and it is said');
});

test('a participant who joins with the file that was saved gets only what came after the save', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(8);
  const a = await join(net, 'a', bytes);
  net.drain(random);
  act(a, () => insertRows(0, 0, 1));
  net.drain(random);
  const saved = await saveOf(a);
  a.session.announceSaved([contentVersion(saved)], a.session.savableRev);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 0, 0, 'top'));
  net.drain(random);
  const b = await join(net, 'b', saved);
  net.drain(random);
  assert.equal(b.session.rev, 2, 'b is at the step of the room');
  assertSame([a, b], 'after joining with the saved file');
  assert.equal(value(b.wb, 0, 1, 0), 'Item', 'the row was inserted once, by the file');
});

test('"saved" tells the others whether their workbook is the one that was saved', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(9);
  const heard: { clean: boolean; versions: string[] }[] = [];
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes, { onSaved: (saved) => heard.push({ clean: saved.clean, versions: saved.versions }) });
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'x'));
  net.drain(random);
  a.session.announceSaved(['e:"tag-1"', 'h:1:0'], a.session.savableRev);
  net.drain(random);
  assert.deepEqual(heard, [{ clean: true, versions: ['e:"tag-1"', 'h:1:0'] }], 'b is at the step that was saved');
  act(b, () => setCellInput(b.wb, 0, 12, 1, 'y'));
  a.session.announceSaved(['e:"tag-2"'], 1);
  net.drain(random);
  assert.equal(heard[1].clean, false, 'b has changed something the saved file does not hold');
  assert.equal(b.session.savableRev, 2, 'with nothing pending b knows the step its own file would hold');
});

test('the log of a room that was cut: who comes back from before the cut loads the file again', async () => {
  const bytes = await startFile();
  const net = new Net({ maxLogEntries: 3 });
  const random = seeded(10);
  const a = await join(net, 'a', bytes);
  let loads = 0;
  let latest = bytes;
  const b = await join(net, 'b', bytes, {}, async () => {
    loads++;
    return latest;
  });
  net.drain(random);
  b.transport.suspend();
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 13, 0, 'never sent'));
  for (let i = 0; i < 6; i++) act(a, () => setCellInput(a.wb, 0, 12, i, `v${i}`));
  net.drain(random);
  latest = await saveOf(a);
  a.session.announceSaved([contentVersion(latest)], a.session.savableRev);
  net.drain(random);
  b.transport.resume();
  await settleAll(net, random, [a, b]);
  // b holds a change that is in no file: its workbook is not given up for the file without being asked.
  assert.equal(loads, 0, 'the file was not loaded again');
  assert.equal(b.session.state, 'detached', 'b is detached from the room');
  assert.deepEqual(codes(b), ['detached'], 'and was told');
  assert.equal(value(b.wb, 0, 13, 0), 'never sent', 'what it had changed is still there');
  assert.equal(act(b, () => setCellInput(b.wb, 0, 14, 0, 'still editing')), true, 'and it can go on editing');
  assert.equal(b.session.pending, 0, 'what it changes now is told to nobody');
  // Its user says: give it up for the file as it is stored.
  b.session.reload();
  await settleAll(net, random, [a, b]);
  assert.equal(loads, 1, 'now the file is loaded again');
  assert.equal(b.session.state, 'online', 'and the room joined');
  assertSame([a, b], 'after the reload');
  assert.equal(value(b.wb, 0, 13, 0), undefined, 'the change that was never sent is gone');
});

test('a read-only participant follows; the relay takes no step from it', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(11);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes, { readOnly: true });
  net.drain(random);
  assert.equal(a.session.peers[0].readOnly, true, 'the others know');
  act(a, () => setCellInput(a.wb, 0, 12, 0, 'seen'));
  net.drain(random);
  assert.equal(value(b.wb, 0, 12, 0), 'seen', 'b follows');
  assert.equal(b.session.readOnly, true, 'b knows that it only reads');
});

test('selections travel, at most one at a time, and follow rows that are inserted above them', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(12);
  const timers: (() => void)[] = [];
  const a = await join(net, 'a', bytes, { user: { name: 'Ada <b>', color: '#12ab34' } });
  const b = await join(net, 'b', bytes, { user: {}, timer: (callback) => (timers.push(callback), () => undefined) });
  net.drain(random);
  assert.deepEqual([b.session.peers[0].name, b.session.peers[0].color], ['Ada <b>', '#12AB34'], 'the name is the text that was given, the colour the one that was given');
  assert.match(a.session.peers[0].color, /^#[0-9A-F]{6}$/, 'a participant without a colour gets one');
  assert.equal(a.session.peers[0].name, undefined, 'and one without a name has none');
  const before = timers.length;
  for (let i = 0; i < 20; i++) b.session.setPresence({ sheet: 0, row: 5 + i, col: 1, range: { top: 5 + i, left: 1, bottom: 5 + i, right: 2 } });
  assert.equal(timers.length, before + 1, 'twenty selections in one moment set one timer going');
  timers[timers.length - 1]();
  net.drain(random);
  assert.deepEqual(a.session.peers[0].presence, { sheet: 0, row: 24, col: 1, range: { top: 24, left: 1, bottom: 24, right: 2 } }, 'the last one arrived');
  act(a, () => insertRows(0, 3, 2));
  assert.equal(a.session.peers[0].presence?.row, 26, 'the selection of b moved with its rows');
});

test('a workbook without changes of its own is loaded again when the room is at another file; a room that started anew detaches one with changes', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(21);
  let unsaved = true;
  const a = await join(net, 'a', bytes, {}, async () => bytes);
  (a.session as unknown as { host: { unsaved(): boolean } }).host.unsaved = () => unsaved;
  net.drain(random);
  for (let i = 0; i < 5; i++) act(a, () => setCellInput(a.wb, 0, 20 + i, 0, `work ${i}`));
  net.drain(random);
  assert.equal(a.session.rev, 5, 'five steps were placed');
  // The relay forgets the room (it was started anew, or let the empty room go) while the connection is down.
  a.transport.suspend();
  net.drain(random);
  (net.server as unknown as { roomsByKey: Map<string, unknown> }).roomsByKey.clear();
  a.transport.resume();
  net.drain(random);
  assert.equal(a.session.state, 'detached', 'the workbook holds steps no file has: it is detached, not thrown away');
  assert.equal(value(a.wb, 0, 20, 0), 'work 0', 'the work is there');
  assert.deepEqual(codes(a), ['detached'], 'and its user was told');
  // The user saves: the file holds everything, and the room is joined with it.
  const saved = await saveOf(a);
  unsaved = false;
  a.session.rejoin([contentVersion(saved)]);
  net.drain(random);
  assert.equal(a.session.state, 'online', 'joined again');
  assert.equal(a.session.rev, 0, 'in a room that starts with the saved file');
  const b = await join(net, 'b', saved);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 30, 0, 'together again'));
  net.drain(random);
  assertSame([a, b], 'after joining again');
});

test('steps that are not as the contract says are refused, whatever they hold', async () => {
  const { a } = await two();
  const styles = a.wb.styles;
  const bad: unknown[] = [
    null,
    {},
    [],
    [{ t: 'cells', s: 0 }],
    [{ t: 'cells', s: -1, c: [] }],
    [{ t: 'cells', s: 0, c: [[0, 0]] }],
    [{ t: 'cells', s: 0, c: [[1048576, 0, 0]] }],
    [{ t: 'cells', s: 0, c: [[0, 16384, 0]] }],
    [{ t: 'cells', s: 0, c: [[0.5, 0, 0]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'number', v: '1' } }]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'number', v: null } }]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'string', v: 'x'.repeat(40000) } }]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'nothing', v: null } }]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'blank', v: null, f: { x: 5 } } }]] }],
    [{ t: 'cells', s: 0, c: [[0, 0, { s: 3 }]] }],
    [{ t: 'st', f: [{ n: 0 }] }, { t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }],
    [{ t: 'st', f: [{ n: 500, f: {}, l: {}, b: {} }] }, { t: 'cells', s: 0, c: [[0, 0, { s: 0 }]] }],
    [{ t: 'rows', s: 0, r: [[0, { h: -1 }]] }],
    [{ t: 'rows', s: 0, r: [[0, { hid: 'yes' }]] }],
    [{ t: 'cols', s: 0, p: [[5, 2, 0]] }],
    [{ t: 'merges', s: 0, add: [[2, 0, 1, 0]], del: [] }],
    [{ t: 'merges', s: 0, add: 'all', del: [] }],
    [{ t: 'filter', s: 0, f: { range: { top: 0, left: 0, bottom: 1, right: 1 }, columns: [{ colId: 9, kind: 'values' }] } }],
    [{ t: 'filter', s: 0, f: { range: { top: 0, left: 0, bottom: 1, right: 1 }, columns: [{ colId: 0, kind: 'values', values: 'a' }] } }],
    [{ t: 'sort', s: 0, g: [0, 0, 1], o: [0, 1] }],
    [{ t: 'ins', s: 0, x: 'diagonal', at: 0, n: 1 }],
    [{ t: 'del', s: 0, x: 'row', at: 0, n: 0 }],
    [{ t: 'notes', s: 0, n: [[0, 0, { x: 5 }]] }],
    [{ t: 'sheet+', s: 0, name: 'x'.repeat(32) }],
    [{ t: 'state', s: 0, state: 'gone' }],
    [{ t: 'tab', s: 0, color: { rgb: 'red' } }],
    [{ t: 'cmd', s: 0, name: 'constructor', a: [] }],
    [{ t: 'cmd', s: 0, name: 'setValidation', a: [JSON.parse('{"__proto__": {"polluted": true}}')] }],
    [{ t: 'restore', s: 0 }],
    [{ t: 'copySheet', s: 0 }],
  ];
  for (const raw of bad) assert.throws(() => decodeStep(raw, styles), { name: 'WireError' }, `refused: ${JSON.stringify(raw).slice(0, 80)}`);
  assert.equal(({} as { polluted?: boolean }).polluted, undefined, 'no prototype was touched');
  assert.equal(a.wb.styles.cellXfs.length, styles.cellXfs.length, 'a step that is refused adds no format');
});

test('random bytes and random values are no message and no step: reading them throws WireError and nothing else', async () => {
  const { a } = await two();
  const random = seeded(99);
  const junk = (depth: number): unknown => {
    const roll = random();
    if (depth > 4 || roll < 0.25) return [0, -1, 1.5, 1e308, 'x', '', null, true, 'cells', 'op', 'hello', 1048576][Math.floor(random() * 12)];
    if (roll < 0.55) return Array.from({ length: Math.floor(random() * 4) }, () => junk(depth + 1));
    const out: Record<string, unknown> = {};
    for (const key of ['t', 's', 'c', 'r', 'p', 'v', 'k', 'f', 'x', 'a', 'n', 'rev', 'id', 'from', 'ops', 'base', 'name', 'you', 'peers', 'add', 'del', 'g', 'o', 'at']) if (random() < 0.3) out[key] = junk(depth + 1);
    if (random() < 0.5) out.t = ['cells', 'rows', 'cols', 'merges', 'filter', 'sort', 'ins', 'del', 'notes', 'sheet+', 'sheet-', 'name', 'move', 'state', 'tab', 'cmd', 'st', 'op', 'hello', 'nack', 'join', 'saved', 'presence'][Math.floor(random() * 23)];
    return out;
  };
  let steps = 0;
  for (let i = 0; i < 4000; i++) {
    const value = junk(0);
    const text = random() < 0.1 ? String.fromCharCode(...Array.from({ length: 20 }, () => Math.floor(random() * 256))) : JSON.stringify(value);
    for (const read of [parseServerMessage, parseClientMessage]) {
      try {
        read(text);
      } catch (error) {
        assert.equal((error as Error).name, 'WireError', `a message that is none is refused as one: ${text.slice(0, 80)}`);
      }
    }
    try {
      decodeStep(Array.isArray(value) ? value : [value], a.wb.styles);
      steps++;
    } catch (error) {
      assert.equal((error as Error).name, 'WireError', `a step that is none is refused as one: ${text.slice(0, 80)}`);
    }
  }
  assert.ok(steps < 4000, 'most of it was refused');
});

// ------------------------------------------- the session and a relay it is handed by the test

/** A session on a transport whose other end is the test. */
async function alone(options: { reload?: () => Promise<string[]>; pausable?: boolean; heard?: (code: string) => void } = {}) {
  const wb = await (await (await import('../src/index')).open(await startFile())).workbook();
  const sent: { t: string; [name: string]: unknown }[] = [];
  const notices: string[] = [];
  let listener: (text: string) => void = () => undefined;
  let state: (state: 'open' | 'closed') => void = () => undefined;
  let reloads = 0;
  let readOnly = 0;
  const transport: Record<string, unknown> = { send: (text: string) => sent.push(JSON.parse(text)), onMessage: (given: (text: string) => void) => (listener = given), close: () => undefined };
  if (options.pausable) {
    transport.onState = (given: (state: 'open' | 'closed') => void) => {
      state = given;
      given('open');
    };
    transport.suspend = () => state('closed');
    transport.resume = () => state('open');
  }
  const { CollabSession } = await import('../src/collab/session');
  const session = new CollabSession(
    { workbook: () => wb, changed: () => undefined, notice: (notice) => void (notices.push(notice.code), options.heard && options.heard(notice.code)), reload: options.reload ? async () => (reloads++, options.reload!()) : undefined },
    { transport: transport as never, versions: ['h:1:x'], timer: () => () => undefined, onReadOnly: () => readOnly++ },
  );
  const tell = (message: unknown): void => listener(JSON.stringify(message));
  const hello = (from: number, rev: number, more: object = {}): void => tell({ t: 'hello', rev, from, ep: 'room1', you: 'me', ro: false, v: [], peers: [], ...more });
  const step = (rev: number, more: object = {}): void => tell({ t: 'op', rev, id: `q:${rev}`, from: 'q', ops: [{ t: 'cells', s: 0, c: [[40 + rev, 0, { v: { k: 'number', v: rev } }]] }], ...more });
  return { wb, session, sent, notices, tell, hello, step, transport: transport as { suspend(): void; resume(): void }, reloads: () => reloads, readOnly: () => readOnly };
}

test('a step is this participant\'s own only when it says what this participant sent', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'mine'));
  const mine = s.sent.filter((message) => message.t === 'op')[0] as { id: string; ops: unknown[] };
  // The relay places a step under that very name that says something else.
  s.tell({ t: 'op', rev: 1, id: mine.id, from: 'someone', ops: [{ t: 'cells', s: 0, c: [[0, 0, { v: { k: 'string', v: 'forged' } }]] }] });
  assert.equal(s.session.pending, 1, 'the step of this participant still waits');
  assert.equal(value(s.wb, 0, 0, 0), 'forged', 'what arrived was applied as the step of someone else');
  assert.equal(value(s.wb, 0, 30, 0), 'mine', 'under what this participant did');
  s.tell({ t: 'nack', id: mine.id, rev: 1 });
  const again = s.sent.filter((message) => message.t === 'op')[1] as { id: string; base: number; ops: unknown[] };
  assert.equal(again.base, 1, 'it is sent again on the step the room is at');
  s.tell({ t: 'op', rev: 2, id: again.id, from: 'me', ops: again.ops });
  assert.equal(s.session.pending, 0, 'and is placed when what comes back is what was sent');
  assert.equal(s.session.rev, 2, 'two steps');
});

test('a participant the relay makes read-only is told, changes nothing from then on, and what it had waiting is taken back', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'before'));
  s.hello(0, 0, { ro: true });
  assert.equal(s.session.readOnly, true, 'read-only');
  assert.equal(s.readOnly(), 1, 'whoever shows the workbook was told');
  assert.deepEqual(s.notices, ['refused', 'read-only'], 'and so was the user');
  assert.equal(value(s.wb, 0, 30, 0), undefined, 'what waited was taken back: nobody else will ever have it');
  assert.throws(() => s.session.submit(setCellInput(s.wb, 0, 31, 0, 'typed by a reader')), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'readOnly', 'a change is refused');
  assert.equal(value(s.wb, 0, 31, 0), undefined, 'and not made');
});

test('steps that arrive out of turn make the session ask again from where it is, once, and never load the file again', async () => {
  const s = await alone({ reload: async () => ['h:1:x'], pausable: true });
  s.hello(0, 0);
  s.step(2);
  s.step(3);
  const joins = s.sent.filter((message) => message.t === 'join');
  assert.equal(joins.length, 2, 'one more "join" for the two steps that came out of turn');
  assert.equal(joins[1].since, 0, 'from the step that is known');
  s.hello(0, 3);
  s.step(1);
  s.step(2);
  s.step(3);
  assert.deepEqual([s.session.rev, s.session.state], [3, 'online'], 'caught up');
  // The answer to a "join" that was sent twice, and steps that are known already.
  s.hello(0, 3);
  s.step(1);
  s.step(3);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual([s.session.rev, s.session.state, s.reloads(), s.notices], [3, 'online', 0, []], 'nothing was loaded again and nothing was said');
  // A log that starts behind where this side is leaves steps out: that is another matter.
  s.hello(5, 6);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual([s.reloads(), s.session.state, s.notices], [0, 'detached', ['detached']], 'the workbook holds three steps nobody is known to have saved: it is detached, not loaded over');
});

test('a relay of another version of the contract ends the session; a room that says no name, or another room than the one this side was in, is not trusted', async () => {
  const s = await alone();
  assert.deepEqual([s.sent[0].pv, typeof s.sent[0].key, (s.sent[0] as { since?: number }).since], [PROTOCOL_VERSION, 'string', undefined], '"join" names the version; a first join names no step');
  assert.equal(PROTOCOL_VERSION, 5);
  assert.equal(s.session.reason, undefined, 'an open session has no reason');
  s.tell({ t: 'error', code: 'version' });
  assert.deepEqual([s.session.state, s.notices, s.session.reason], ['closed', ['version'], 'version'], 'a room of another version is told apart from a refusal');
  // Any other refusal is a refusal.
  const refused = await alone();
  refused.tell({ t: 'error', code: 'unauthorized' });
  assert.deepEqual([refused.session.state, refused.notices, refused.session.reason], ['closed', ['refused'], 'refused']);
  // The relay says the version of the room in "hello": another one is not joined, whatever the relay let in.
  for (const pv of [PROTOCOL_VERSION - 1, PROTOCOL_VERSION + 1]) {
    const older = await alone();
    older.hello(0, 0, { pv });
    assert.deepEqual([older.session.state, older.notices, older.session.reason], ['closed', ['version'], 'version'], `a "hello" of version ${pv}`);
  }
  const same = await alone();
  same.hello(0, 0, { pv: PROTOCOL_VERSION });
  assert.deepEqual([same.session.state, same.notices], ['online', []], 'a "hello" of this version is joined');
  const silent = await alone();
  silent.hello(0, 0);
  assert.equal(silent.session.state, 'online', 'a relay that names no version is taken at its word that it let this side in');
  // A relay of another language may write the version as a text of digits; nothing else is a version.
  const digits = await alone();
  digits.hello(0, 0, { pv: String(PROTOCOL_VERSION) });
  assert.deepEqual([digits.session.state, digits.notices], ['online', []], 'a version written as digits is that version');
  const otherDigits = await alone();
  otherDigits.hello(0, 0, { pv: String(PROTOCOL_VERSION - 1) });
  assert.deepEqual([otherDigits.session.state, otherDigits.session.reason], ['closed', 'version'], 'and another one written so is another version');
  for (const pv of [`${PROTOCOL_VERSION}.0`, 'x', '', ' 5', '05', '1001', 3.5, true]) {
    const odd = await alone();
    odd.hello(0, 0, { pv });
    assert.deepEqual([odd.session.state, odd.session.reason], ['closed', 'peer'], `a version that is no number makes the "hello" no message: ${JSON.stringify(pv)}`);
  }
  // Whoever hears the notice of the end can read why already.
  const box: { s?: Awaited<ReturnType<typeof alone>> } = {};
  const heard: (string | undefined)[] = [];
  box.s = await alone({ heard: () => void heard.push(box.s!.session.reason) });
  box.s.tell({ t: 'error', code: 'version' });
  assert.deepEqual(heard, ['version'], 'the reason is there when the notice is given');
  // Left by its own will: no reason.
  const left = await alone();
  left.hello(0, 0);
  left.session.leave();
  assert.deepEqual([left.session.state, left.session.reason], ['closed', undefined]);
  // A relay that names no room is none of this contract.
  const bare = await alone();
  bare.tell({ t: 'hello', rev: 0, from: 0, you: 'me', ro: false, v: [], peers: [] });
  assert.equal(bare.session.state, 'closed', 'a "hello" without the name of the room is no message');
  // Without the word of a file there is only the start of a room to join at.
  const late = await alone();
  late.hello(2, 5);
  assert.equal(late.session.state, 'closed', 'a relay that lets a file in at a later step on nobody\'s word is not followed');
  // Coming back, the join names the room and the step; another room in answer is another room.
  const other = await alone({ pausable: true });
  other.hello(0, 0);
  other.session.submit(setCellInput(other.wb, 0, 30, 0, 'x'));
  const sentOp = other.sent.filter((message) => message.t === 'op')[0] as unknown as { id: string; ops: unknown[] };
  other.tell({ t: 'op', rev: 1, id: sentOp.id, from: 'me', ops: sentOp.ops });
  other.transport.suspend();
  other.transport.resume();
  const joins = other.sent.filter((message) => message.t === 'join');
  assert.equal(joins.length, 2, `it joins again: ${other.sent.map((message) => message.t).join()}`);
  const again = joins[1] as { t: string; since?: number; ep?: string };
  assert.deepEqual([again.t, again.since, again.ep], ['join', 1, 'room1'], 'the join after a lost connection names the room this side was in, and its step');
  other.hello(1, 1, { ep: 'room2' });
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual([other.session.state, other.notices], ['detached', ['detached']], 'a room of another name is not the room the step was of: the workbook, with a step nobody saved, is kept apart');
});

test('what the others say of where they are is passed on a few times a second, however often they say it', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(31);
  const timers: (() => void)[] = [];
  let told = 0;
  const a = await join(net, 'a', bytes, { timer: (callback) => (timers.push(callback), () => undefined), onPeers: () => told++ });
  const b = await join(net, 'b', bytes);
  net.drain(random);
  const before = told;
  const raw = b.transport;
  for (let i = 0; i < 2000; i++) raw.send(JSON.stringify({ t: 'presence', data: { s: 0, a: [i % 50, 1], r: [i % 50, 1, i % 50, 1] } }));
  net.drain(random);
  assert.equal(told, before, 'nothing was passed on yet');
  assert.equal(a.session.peers[0].presence?.row, 49, 'whoever asks gets the latest');
  for (const run of timers.splice(0)) run();
  assert.equal(told, before + 1, 'two thousand messages, one telling');
});

test('a step that was refused waits until the workbook is at the step the relay named, and is sent once', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'mine'));
  const ops = (): { base: number }[] => s.sent.filter((message) => message.t === 'op') as never;
  s.tell({ t: 'nack', id: (ops()[0] as unknown as { id: string }).id, rev: 2 });
  assert.equal(ops().length, 1, 'not sent again while two steps are missing');
  s.step(1);
  assert.equal(ops().length, 1, 'nor with one missing');
  s.step(2);
  assert.equal(ops().length, 2, 'sent when the workbook is there');
  assert.equal(ops()[1].base, 2, 'on that step');
  s.session.submit(setCellInput(s.wb, 0, 31, 0, 'second'));
  assert.equal(ops().length, 2, 'one step at a time is on its way');
});

test('what arrives while a file is written waits, and a change made here meanwhile is told to whoever writes', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.pause();
  s.step(1);
  assert.equal(s.session.rev, 0, 'held back');
  assert.equal(s.session.resume(), false, 'nothing was changed here meanwhile');
  assert.equal(s.session.rev, 1, 'let in');
  s.session.pause();
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'typed while the file was written'));
  assert.equal(s.session.resume(), true, 'the file that was written is of no one step');
  assert.equal(s.session.savableRev, undefined, 'and with a step waiting there is no step a file would hold');
});

test('a new change forgets what was undone; a step that is too large is refused whole', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'one'));
  const sentOne = s.sent.filter((message) => message.t === 'op')[0] as { id: string; ops: unknown[] };
  assert.equal(s.session.undo(), undefined, 'an undo waits while the step is on its way');
  assert.equal(value(s.wb, 0, 30, 0), 'one', 'nothing is undone yet');
  s.tell({ t: 'op', rev: 1, id: sentOne.id, from: 'me', ops: sentOne.ops });
  assert.equal(value(s.wb, 0, 30, 0), undefined, 'it is undone once the step is placed');
  assert.equal(s.session.canRedo, true, 'what was undone can be made again');
  s.session.submit(setCellInput(s.wb, 0, 31, 0, 'two'));
  assert.equal(s.session.canRedo, false, 'until something else is done');
  const big = 'x'.repeat(30000);
  const cells = Array.from({ length: 40 }, (_x, i) => ({ row: 50 + i, col: 0, cell: { kind: 'string' as const, value: big + i, style: 0 } }));
  assert.throws(() => s.session.submit([{ type: 'setCells', sheet: 0, cells }]), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'size', 'more than the relay takes in one message');
  assert.equal(value(s.wb, 0, 50, 0), undefined, 'nothing of it was applied');
  assert.throws(() => s.session.submit(Array.from({ length: 10001 }, () => ({ type: 'setMerges' as const, sheet: 0, merges: [] }))), (error: unknown) => error instanceof CollabUnsupportedError && error.what === 'size', 'more operations than a step holds');
});

test('what is said: part of a change that went, an undo that brings not everything back, a file that was loaded over changes', async () => {
  const { net, random, a, b } = await two(41);
  firstThen(net, random, a, b, () => {
    b.session.submit([{ type: 'setCells', sheet: 0, cells: [{ row: 3, col: 6, cell: { kind: 'number', value: 1, style: 0 } }, { row: 5, col: 6, cell: { kind: 'number', value: 2, style: 0 } }] }]);
    act(a, () => deleteRows(a.wb, 0, 3, 1));
  });
  assert.deepEqual(codes(b), ['partial'], 'one of two cells went with its row');
  assert.equal(value(a.wb, 0, 4, 6), 2, 'the other one is where its row is');
  act(a, () => setComment(a.wb, 0, 6, 0, 'a note'));
  act(a, () => deleteRows(a.wb, 0, 6, 1));
  net.drain(random);
  a.notices.length = 0;
  assert.notEqual(a.session.undo(), undefined, 'the deletion is undone');
  assert.deepEqual(codes(a), ['undo-lossy'], 'and it is said that the note did not come back');
  assert.deepEqual([a.session.stats.dropped, b.session.stats.quiet], [0, 0], 'nothing was dropped');

  const bytes = await startFile();
  const c = await join(net, 'c', bytes, { room: 'another' }, async () => bytes);
  net.drain(random);
  c.transport.suspend();
  net.drain(random);
  act(c, () => setCellInput(c.wb, 0, 30, 0, 'never sent'));
  c.session.reload();
  await c.loading;
  assert.deepEqual(c.notices, [{ code: 'reloaded', count: 1 }], 'a change that was given up with the file is counted');
});

test('a file that the room never takes is not loaded again for ever', async () => {
  const timers: (() => void)[] = [];
  const wb = await (await (await import('../src/index')).open(await startFile())).workbook();
  let listener: (text: string) => void = () => undefined;
  let loads = 0;
  const notices: string[] = [];
  const { CollabSession } = await import('../src/collab/session');
  const session = new CollabSession(
    { workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice.code), reload: async () => (loads++, ['h:1:x']) },
    { transport: { send: () => undefined, onMessage: (given) => (listener = given), close: () => undefined }, versions: ['h:1:x'], timer: (callback) => (timers.push(callback), () => undefined) },
  );
  for (let i = 0; i < 20 && session.state !== 'detached'; i++) {
    listener(JSON.stringify({ t: 'reset', rev: 4, v: ['other'] }));
    for (const run of timers.splice(0)) run();
    for (let k = 0; k < 5; k++) await new Promise((resolve) => setImmediate(resolve));
  }
  assert.equal(session.state, 'detached', 'the session gives up, and what is edited stays');
  assert.equal(loads, 6, 'after six tries');
  assert.deepEqual(notices, ['waiting', 'stale', 'detached'], 'having said that it waits, and then that it is over');
});

test('a room that takes the file and sends it away again, time after time, is not joined for ever', async () => {
  const timers: (() => void)[] = [];
  const wb = await (await (await import('../src/index')).open(await startFile())).workbook();
  let listener: (text: string) => void = () => undefined;
  let loads = 0;
  const notices: string[] = [];
  const { CollabSession } = await import('../src/collab/session');
  const session = new CollabSession(
    { workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice.code), reload: async () => (loads++, ['h:1:x']) },
    { transport: { send: () => undefined, onMessage: (given) => (listener = given), close: () => undefined }, versions: ['h:1:x'], timer: (callback) => (timers.push(callback), () => undefined) },
  );
  const settle = async (): Promise<void> => {
    for (const run of timers.splice(0)) run();
    for (let k = 0; k < 5; k++) await new Promise((resolve) => setImmediate(resolve));
  };
  for (let i = 0; i < 20 && session.state !== 'closed' && session.state !== 'detached'; i++) {
    // The room says welcome, and then that it is at another file: nothing was done in between.
    listener(JSON.stringify({ t: 'hello', rev: 0, from: 0, ep: `room${i}`, you: 'me', ro: false, v: [], peers: [] }));
    listener(JSON.stringify({ t: 'reset', rev: 4, v: ['other'] }));
    await settle();
  }
  assert.equal(session.state, 'detached', 'what is edited is kept, apart from the room');
  assert.equal(loads, 6, 'after six tries');
  assert.deepEqual(notices, ['waiting', 'stale', 'detached']);
  // A step that arrives is progress: the count starts again.
  const again = new CollabSession(
    { workbook: () => wb, changed: () => undefined, notice: () => undefined, unsaved: () => false, reload: async () => (loads++, ['h:1:x']) },
    { transport: { send: () => undefined, onMessage: (given) => (listener = given), close: () => undefined }, versions: ['h:1:x'], timer: (callback) => (timers.push(callback), () => undefined) },
  );
  for (let i = 0; i < 9; i++) {
    listener(JSON.stringify({ t: 'hello', rev: 0, from: 0, ep: `again${i}`, you: 'me', ro: false, v: [], peers: [] }));
    listener(JSON.stringify({ t: 'op', rev: 1, id: 'q:1', from: 'q', ops: [{ t: 'cells', s: 0, c: [[41, 0, { v: { k: 'number', v: i } }]] }] }));
    listener(JSON.stringify({ t: 'reset', rev: 4, v: ['other'] }));
    await settle();
  }
  assert.notEqual(again.state, 'closed');
  assert.notEqual(again.state, 'detached', 'a room in which steps arrive is joined again as often as it takes');
});

test('"saved" is no news of a clean workbook for who has a step waiting', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(51);
  const heard: boolean[] = [];
  // The save is said again a moment after someone came in: the moment passes when the test says so.
  let timers: (() => void)[] = [];
  const a = await join(net, 'a', bytes, { timer: (run) => (timers.push(run), () => undefined) });
  const b = await join(net, 'b', bytes, { onSaved: (saved) => heard.push(saved.clean) });
  net.drain(random);
  b.transport.suspend();
  net.drain(random);
  act(b, () => setCellInput(b.wb, 0, 30, 0, 'waits'));
  b.transport.resume();
  a.session.announceSaved(['e:"t"'], 0);
  // The announcement arrives while the step of b is still on its way.
  for (let i = 0; i < 200 && heard.length === 0; i++) {
    if (!net.step(random)) {
      const due = timers;
      timers = [];
      for (const run of due) run();
    }
  }
  net.drain(random);
  assert.equal(heard.length >= 1, true, 'b heard of the save');
  assert.equal(b.session.rev === 0 ? heard[0] : false, false, 'and is not told that its workbook is the one that was saved');
});

test('the relay of the page: a full room, a presence that is too large, a stranger in a room that has somebody', async () => {
  const { MemoryServer } = await import('../src/collab/server');
  const queue: (() => void)[] = [];
  const server = new MemoryServer({ maxConnections: 1, deliver: (run) => queue.push(run) });
  const drain = (): void => {
    while (queue.length > 0) queue.shift()!();
  };
  const connect = (join: object): { got: { t: string; code?: string }[]; send(message: object): void } => {
    const raw = server.connect();
    const got: { t: string; code?: string }[] = [];
    raw.onMessage((text) => got.push(JSON.parse(text)));
    raw.onState!((state) => {
      if (state === 'open') raw.send(JSON.stringify(join));
    });
    return { got, send: (message) => raw.send(JSON.stringify(message)) };
  };
  const first = connect({ t: 'join', v: ['v1'], key: 'a' });
  drain();
  assert.equal(first.got[0].t, 'hello', 'the first one is in');
  const second = connect({ t: 'join', v: ['v1'], key: 'b' });
  drain();
  assert.deepEqual(second.got.map((message) => [message.t, message.code]), [['error', 'room-full']], 'a room that is full takes nobody');
  const other = new MemoryServer({ deliver: (run) => queue.push(run) });
  const raw = other.connect();
  const seen: { t: string }[] = [];
  raw.onMessage((text) => seen.push(JSON.parse(text)));
  raw.onState!(() => raw.send(JSON.stringify({ t: 'join', v: ['v1'], key: 'a' })));
  const watcher = other.connect();
  const watched: { t: string }[] = [];
  watcher.onMessage((text) => watched.push(JSON.parse(text)));
  watcher.onState!(() => watcher.send(JSON.stringify({ t: 'join', v: ['v1'], key: 'b' })));
  while (queue.length > 0) queue.shift()!();
  raw.send(JSON.stringify({ t: 'presence', data: 'x'.repeat(20000) }));
  raw.send(JSON.stringify({ t: 'presence', data: 'small' }));
  while (queue.length > 0) queue.shift()!();
  assert.deepEqual(watched.filter((message) => message.t === 'presence').length, 1, 'a presence that is too large is not passed on');
  const stranger = other.connect();
  const told: { t: string }[] = [];
  stranger.onMessage((text) => told.push(JSON.parse(text)));
  stranger.onState!(() => stranger.send(JSON.stringify({ t: 'join', v: ['another file'], key: 'c' })));
  while (queue.length > 0) queue.shift()!();
  assert.deepEqual(told.map((message) => message.t), ['reset'], 'a room without steps that has somebody in it is not started anew for a stranger');
  assert.equal(watched.some((message) => message.t === 'sync'), true, 'and those in it are asked to save');
});

test('the relay of the page: a key that joins again takes the place of its connection; who is asked to save, and how often; a "saved" of a step nobody saw', async () => {
  const { MemoryServer } = await import('../src/collab/server');
  const queue: (() => void)[] = [];
  let clock = 0;
  const server = new MemoryServer({ deliver: (run) => queue.push(run), now: () => clock });
  const drain = (): void => {
    while (queue.length > 0) queue.shift()!();
  };
  type Got = { t: string; code?: string; rev?: number; v?: string[]; from?: string; you?: string };
  const connect = (join: object): { got: Got[]; send(message: object): void } => {
    const raw = server.connect();
    const got: Got[] = [];
    raw.onMessage((text) => got.push(JSON.parse(text)));
    raw.onState!((state) => {
      if (state === 'open') raw.send(JSON.stringify(join));
    });
    return { got, send: (message) => raw.send(JSON.stringify(message)) };
  };
  const writer = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'secret-of-a' });
  const reader = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'secret-of-b', ro: true });
  drain();
  assert.deepEqual([writer.got[0].t, reader.got[0].t], ['hello', 'hello'], 'both are in');
  assert.equal(JSON.stringify(writer.got.concat(reader.got)).indexOf('secret-of-'), -1, 'the key of a participant is told to nobody');
  // The same key again: the old connection is told that it was replaced, and the new one is in.
  const again = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'secret-of-a' });
  drain();
  assert.equal(again.got[0].t, 'hello', 'the key joins again');
  assert.deepEqual(writer.got.filter((message) => message.t === 'error').map((message) => message.code), ['replaced'], 'and the connection it had is told so');
  // Another version of the contract is no participant.
  const old = connect({ t: 'join', pv: PROTOCOL_VERSION + 1, v: ['v1'], key: 'c' });
  drain();
  assert.equal(old.got[0].t, 'error', 'a contract the relay does not speak');
  // Who cannot join: those who write are asked to save, those who only read are not; and a reader who cannot join asks nobody.
  const asked = (who: { got: Got[] }): number => who.got.filter((message) => message.t === 'sync').length;
  connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['other'], key: 'd', ro: true });
  drain();
  assert.deepEqual([asked(again), asked(reader)], [0, 0], 'someone who only reads makes nobody save');
  connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['other'], key: 'e' });
  drain();
  assert.deepEqual([asked(again), asked(reader)], [1, 0], 'someone who writes does; who only reads is not asked');
  connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['other'], key: 'f' });
  connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['other'], key: 'g' });
  drain();
  assert.equal(asked(again), 2, 'twice a minute and no more');
  clock += 61000;
  connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['other'], key: 'h' });
  drain();
  assert.equal(asked(again), 3, 'and again a minute later');
  // "saved": of a step the room does not have, the step is not passed on; nor may a version move to another step.
  const told = (): Got[] => reader.got.filter((message) => message.t === 'saved');
  again.send({ t: 'saved', v: ['v2'], rev: 7 });
  drain();
  assert.deepEqual(told().map((message) => message.rev), [undefined], 'a step nobody has seen is not named to the others');
  again.send({ t: 'op', id: 'x:1', base: 0, ops: [{ t: 'cells', s: 0, c: [] }] });
  drain();
  again.send({ t: 'saved', v: ['v3'], rev: 1 });
  again.send({ t: 'saved', v: ['v3'], rev: 0 });
  again.send({ t: 'saved', v: ['v1'], rev: 1 });
  drain();
  assert.deepEqual(told().slice(1).map((message) => message.rev), [1, undefined, undefined], 'a version stays at the step it was first saved at, also the one the room started with');
  // Who joins: with the file the room was started with from the start; with a later one only on the word of the
  // file itself (the room and the step that were written into it), never on what was said of it.
  const ep = (again.got[0] as { ep?: string }).ep;
  assert.equal(typeof ep, 'string', 'the room says what it is called');
  const said = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v3'], key: 'i' });
  const stamped = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v3'], key: 'k', since: 1, ep });
  const stranger = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v3'], key: 'l', since: 1, ep: 'another-room' });
  const bare = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v3'], key: 'm', since: 1 });
  const first = connect({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'j' });
  drain();
  assert.equal(said.got[0].t, 'reset', 'what was said of a file is no ground to join on');
  assert.deepEqual([stamped.got[0].t, stamped.got[0].rev, stamped.got[0].from, stamped.got[0].v], ['hello', 1, 1, ['v3']], 'who has the saved file starts behind the step it names, and hears that the room knows the file');
  assert.deepEqual([stranger.got[0].t, bare.got[0].t], ['reset', 'reset'], 'a step of another room, or of none, is the step of nothing');
  assert.deepEqual([first.got[0].t, first.got[0].from], ['hello', 0], 'who has the first file gets every step');
});

test('one participant saves for the others: who writes and whose name of the relay comes first', async () => {
  const net = new Net();
  const random = seeded(3);
  const bytes = await startFile();
  const members = [await join(net, 'a', bytes), await join(net, 'b', bytes), await join(net, 'c', bytes, { readOnly: true })];
  net.drain(random);
  const writers = members.filter((member) => !member.session.readOnly);
  assert.equal(writers.length, 2, 'two of the three write');
  assert.deepEqual(members.filter((member) => member.session.leads).length, 1, 'exactly one of them saves for the others');
  const leader = members.find((member) => member.session.leads)!;
  assert.equal(leader.session.readOnly, false, 'and it is one who writes');
  const ids = writers.map((member) => (member.session as unknown as { me: string }).me).sort();
  assert.equal((leader.session as unknown as { me: string }).me, ids[0], 'the one whose name comes first');
  leader.session.leave();
  net.drain(random);
  const next = members.filter((member) => member !== leader && member.session.leads);
  assert.deepEqual(next.map((member) => member.name), writers.filter((member) => member !== leader).map((member) => member.name), 'when it leaves, the other writer does');
});

test('a file that is loaded again takes the place of what waited: nothing of the old workbook is sent after it', async () => {
  const net = new Net();
  const random = seeded(5);
  const bytes = await startFile();
  const a = await join(net, 'a', bytes);
  let stored = bytes;
  const b = await join(net, 'b', bytes, {}, async () => stored);
  net.drain(random);
  act(a, () => setCellInput(a.wb, 0, 2, 0, 'kept'));
  net.drain(random);
  stored = await saveOf(a);
  a.session.announceSaved([contentVersion(stored)], a.session.savableRev);
  net.drain(random);
  // b goes offline with a step waiting and comes back with the file as it is saved now.
  b.transport.suspend();
  act(b, () => setCellInput(b.wb, 0, 3, 0, 'never sent'));
  assert.equal(b.session.pending, 1, 'a step waits');
  const sent: string[] = [];
  const send = b.transport.send.bind(b.transport);
  b.transport.send = (text: string) => {
    sent.push(text);
    send(text);
  };
  b.session.reload();
  await settleAll(net, random, [a, b]);
  b.transport.resume();
  await settleAll(net, random, [a, b]);
  assert.equal(b.session.pending, 0, 'nothing waits after the load');
  assert.equal(sent.some((text) => text.indexOf('never sent') >= 0), false, 'and the step of the workbook that went was sent to nobody');
  assert.deepEqual(b.notices.map((notice) => notice.code), ['reloaded'], 'which is said');
  assertSame([a, b], 'after the load');
});

test('a relay that writes a step out anew (other order of the names, same step) hands this participant its own step', async () => {
  const s = await alone();
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, '=A1+1'));
  const sent = s.sent.filter((message) => message.t === 'op')[0] as unknown as { id: string; ops: unknown[] };
  // Every object with its names the other way round, as a relay that parses and writes again may send it.
  const turned = (value: unknown): unknown => {
    if (Array.isArray(value)) return value.map(turned);
    if (value === null || typeof value !== 'object') return value;
    const out: { [name: string]: unknown } = {};
    for (const name of Object.keys(value as object).reverse()) out[name] = turned((value as { [name: string]: unknown })[name]);
    return out;
  };
  assert.notEqual(JSON.stringify(turned(sent.ops)), JSON.stringify(sent.ops), 'the step is written differently (or the test tests nothing)');
  s.tell({ t: 'op', rev: 1, id: sent.id, from: 'me', ops: turned(sent.ops) });
  assert.deepEqual([s.session.pending, s.session.rev, s.session.state, s.notices], [0, 1, 'online', []], 'it is the step this side sent: placed, and nothing waits');
  // Another step under that name is someone else's.
  s.session.submit(setCellInput(s.wb, 0, 31, 0, 'mine'));
  const next = s.sent.filter((message) => message.t === 'op')[1] as unknown as { id: string };
  s.tell({ t: 'op', rev: 2, id: next.id, from: 'q', ops: [{ t: 'cells', s: 0, c: [[32, 0, { v: { k: 'string', v: 'theirs' } }]] }] });
  assert.equal(s.session.pending, 1, 'a step that says something else is not this side\'s step, whatever it is called');
});

test('undo waits for what is on its way, is told when it does not get its turn, and a step that was not placed is undone like any other once the session is on its own', async () => {
  const s = await alone({ pausable: true });
  s.hello(0, 0);
  s.session.submit(setCellInput(s.wb, 0, 30, 0, 'first'));
  assert.equal(s.session.undo(), undefined, 'a step that is on its way is not undone yet');
  assert.equal(value(s.wb, 0, 30, 0), 'first', 'and the cell is as it was typed');
  // Something else is typed before the undo had its turn: it does not get it, which is said.
  s.session.submit(setCellInput(s.wb, 0, 31, 0, 'second'));
  assert.deepEqual(s.notices, ['undo-wait'], 'the undo that waited is given up, and that is said');
  // While a file is loaded again nothing is taken.
  const loading = await alone({ reload: () => new Promise<string[]>(() => undefined) });
  loading.hello(0, 0);
  loading.session.reload();
  assert.throws(() => loading.session.submit(setCellInput(loading.wb, 0, 30, 0, 'x')), /loaded again/, 'a change made while the file is loaded again is refused');
  assert.equal(loading.session.undo(), undefined, 'and so is an undo');
  // Out of the room with two steps that were never placed: they are steps of the workbook, to be undone one by one.
  const sentOp = s.sent.filter((message) => message.t === 'op')[0] as unknown as { id: string };
  s.tell({ t: 'nack', id: sentOp.id, rev: 0 });
  s.transport.suspend();
  s.transport.resume();
  s.hello(0, 0, { ep: 'another-room' });
  assert.equal(s.session.state, 'detached', 'another room: the workbook is kept apart with what waited');
  assert.equal(s.session.canUndo, true, 'what waited can be undone');
  assert.ok(s.session.undo(), 'the last step is undone at once');
  assert.deepEqual([value(s.wb, 0, 31, 0), value(s.wb, 0, 30, 0)], [undefined, 'first'], 'the second step is taken back, the first is there');
  assert.ok(s.session.undo(), 'and the one before it');
  assert.equal(value(s.wb, 0, 30, 0), undefined, 'the first step is taken back as well');
  assert.ok(s.session.redo(), 'what was undone is made again');
  assert.equal(value(s.wb, 0, 30, 0), 'first', 'the first step is back');
});

test('the relay of the page: an empty room that is started again on another file is another room; on its own file it is the room it was', async () => {
  const { MemoryServer } = await import('../src/collab/server');
  const queue: (() => void)[] = [];
  const server = new MemoryServer({ deliver: (run) => queue.push(run) });
  const visit = (join: object): { ep?: string; from?: number; t: string } => {
    const raw = server.connect();
    const got: { t: string; ep?: string; from?: number }[] = [];
    raw.onMessage((text) => got.push(JSON.parse(text)));
    raw.onState!((state) => {
      if (state === 'open') raw.send(JSON.stringify(join));
    });
    while (queue.length > 0) queue.shift()!();
    raw.close();
    while (queue.length > 0) queue.shift()!();
    return got[0];
  };
  const first = visit({ t: 'join', pv: PROTOCOL_VERSION, v: ['file-1'], key: 'a' });
  const same = visit({ t: 'join', pv: PROTOCOL_VERSION, v: ['file-1'], key: 'b' });
  const other = visit({ t: 'join', pv: PROTOCOL_VERSION, v: ['file-2'], key: 'c' });
  const back = visit({ t: 'join', pv: PROTOCOL_VERSION, v: ['file-1'], key: 'a', since: 0, ep: first.ep });
  assert.deepEqual([first.t, same.t, other.t], ['hello', 'hello', 'hello'], 'each of them starts the empty room');
  assert.equal(same.ep, first.ep, 'on the file it was started with, it is the room it was');
  assert.notEqual(other.ep, first.ep, 'on another file it is another room');
  assert.notEqual(back.ep, first.ep, 'and who comes back to the room before it is not in that room: the name says so');
});

test('the relay of the page takes only so many joins a minute of one connection; and a save is said again once for however many come in', async () => {
  const { MemoryServer, MAX_JOINS_PER_MINUTE } = await import('../src/collab/server');
  const queue: (() => void)[] = [];
  const server = new MemoryServer({ deliver: (run) => queue.push(run), now: () => 5000 });
  const raw = server.connect();
  const got: { t: string; code?: string }[] = [];
  raw.onMessage((text) => got.push(JSON.parse(text)));
  raw.onState!((state) => {
    if (state === 'open') raw.send(JSON.stringify({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'a' }));
  });
  const drain = (): void => {
    while (queue.length > 0) queue.shift()!();
  };
  drain();
  for (let i = 1; i < MAX_JOINS_PER_MINUTE + 1; i++) {
    raw.send(JSON.stringify({ t: 'leave' }));
    raw.send(JSON.stringify({ t: 'join', pv: PROTOCOL_VERSION, v: ['v1'], key: 'a' }));
    drain();
  }
  assert.deepEqual(got.filter((message) => message.t === 'hello').length, MAX_JOINS_PER_MINUTE, 'ten joins get in');
  assert.deepEqual(got[got.length - 1], { t: 'error', code: 'join-rate' }, 'the eleventh ends the connection');
  // Three come in at once: what was last saved is said to them once, a moment later.
  const s = await alone();
  s.hello(0, 0);
  s.session.announceSaved(['e:"v2"'], 0);
  const said = (): number => s.sent.filter((message) => message.t === 'saved').length;
  assert.equal(said(), 1, 'the save is said');
  let timers: (() => void)[] = [];
  (s.session as unknown as { timer: (run: () => void, ms: number) => () => void }).timer = (run) => (timers.push(run), () => undefined);
  for (const id of ['p1', 'p2', 'p3']) s.tell({ t: 'peer', id, user: {}, ro: false });
  assert.equal(said(), 1, 'not again at once, for each who comes');
  const due = timers;
  timers = [];
  for (const run of due) run();
  assert.equal(said(), 2, 'once, for all three');
});

test('a save that was made while the room could not be reached is said when the room is back', async () => {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(61);
  const heard: string[][] = [];
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes, { onSaved: (saved) => heard.push(saved.versions) });
  net.drain(random);
  a.transport.suspend();
  net.drain(random);
  a.session.announceSaved(['e:"saved-offline"'], a.session.savableRev);
  net.drain(random);
  assert.deepEqual(heard, [], 'nobody hears of it while its participant is out of reach');
  a.transport.resume();
  net.drain(random);
  assert.deepEqual(heard, [['e:"saved-offline"']], 'the others hear of it when it is back: the file is known by another name now, and their next save is no conflict');
  a.transport.suspend();
  a.transport.resume();
  net.drain(random);
  assert.equal(heard.length, 1, 'once');
});
