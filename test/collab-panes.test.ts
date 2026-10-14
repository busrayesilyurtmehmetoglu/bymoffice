// Frozen panes where a workbook is edited together: they travel as what they are ("pane"), the
// last one the relay ordered is what everybody has, a sheet that moves takes them along, and an
// undo takes back only one's own.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { decodeStep, encodeOp, encodeStep } from '../src/collab/wire';
import { WireError } from '../src/collab/types';
import { freezePanes, frozenPanes } from '../src/ops/panes';
import { insertSheet, moveSheet } from '../src/ops/sheets';
import { Net, act, assertSame, join, saveOf, seeded, startFile, type Member } from './collab-fixtures';
import { ZipArchive } from '../src/index';
import { buildXlsx, worksheetXml } from './fixtures';

async function two(seed = 1): Promise<{ net: Net; random: () => number; a: Member; b: Member }> {
  const bytes = await startFile();
  const net = new Net();
  const random = seeded(seed);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  return { net, random, a, b };
}

const panes = (member: Member) => member.wb.sheets.map((sheet) => `${sheet.name}:${sheet.view.frozenRows},${sheet.view.frozenCols},${sheet.view.topLeftCell || ''}`).join(' ');

test('collab panes: a freeze reaches the others, and an undo takes it back for everybody', async () => {
  const { net, random, a, b } = await two();
  assert.ok(act(a, () => freezePanes(a.wb, 0, 1, 2)));
  net.drain(random);
  assert.deepEqual(frozenPanes(b.wb, 0), { rows: 1, cols: 2 });
  assert.equal(b.wb.sheets[0].view.topLeftCell, 'C2');
  assert.equal(panes(a), panes(b));
  assert.ok(a.session.undo());
  net.drain(random);
  assert.deepEqual(frozenPanes(b.wb, 0), { rows: 0, cols: 0 });
  assert.equal(panes(a), panes(b));
  assertSame([a, b], 'after the undo');
});

test('collab panes: two freezes at once end as the one the relay took last, on both sides', async () => {
  for (let seed = 1; seed <= 6; seed++) {
    const { net, random, a, b } = await two(seed);
    b.transport.suspend();
    net.drain(random);
    act(a, () => freezePanes(a.wb, 0, 3, 0));
    act(b, () => freezePanes(b.wb, 0, 0, 1));
    net.drain(random);
    b.transport.resume();
    net.drain(random);
    assert.equal(panes(a), panes(b), `seed ${seed}`);
    // The relay had a's first: b's comes after it and stands.
    assert.deepEqual(frozenPanes(a.wb, 0), { rows: 0, cols: 1 }, `seed ${seed}`);
  }
});

test('collab panes: a sheet inserted by someone else before the freeze arrives leaves the panes on the sheet they were for', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => freezePanes(b.wb, 1, 2, 0));
  act(a, () => insertSheet(a.wb, 0, { name: 'First' }));
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assert.equal(panes(a), panes(b));
  const other = a.wb.sheets.find((sheet) => sheet.name === 'Other')!;
  assert.deepEqual([other.view.frozenRows, other.view.frozenCols], [2, 0]);
  assert.deepEqual(b.notices, []);
});

test('collab panes: a sheet moved by someone else meanwhile - the freeze is dropped and that is said, as for the colour of a tab', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => freezePanes(b.wb, 1, 2, 0));
  act(a, () => moveSheet(a.wb, 1, 0));
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assert.equal(panes(a), panes(b));
  const other = a.wb.sheets.find((sheet) => sheet.name === 'Other')!;
  const carried = other.view.frozenRows === 2;
  assert.ok(carried || b.notices.some((notice) => notice.code === 'dropped'), 'carried, or dropped and said');
});

test('collab panes: the wire form is checked like everything that arrives', async () => {
  const { a } = await two();
  const ops = freezePanes(a.wb, 0, 4, 1);
  const encoded = encodeOp(a.wb, ops[0]);
  assert.deepEqual(encoded.op, { t: 'pane', s: 0, r: 4, c: 1, tl: 'B5' });
  const raw = encodeStep([encoded.op!], a.wb.styles);
  assert.deepEqual(decodeStep(JSON.parse(JSON.stringify(raw)), a.wb.styles), [encoded.op]);
  // A pane scrolled further down than the frozen rows (as Excel writes it) travels.
  assert.deepEqual(decodeStep([{ t: 'pane', s: 0, r: 2, c: 0, tl: 'A30' }], a.wb.styles), [{ t: 'pane', s: 0, r: 2, c: 0, tl: 'A30' }]);
  for (const bad of [
    { t: 'pane', s: 0, r: -1, c: 0 },
    { t: 'pane', s: 0, r: 0, c: 16384 },
    { t: 'pane', s: 0, r: 1.5, c: 0 },
    { t: 'pane', s: 0, r: 1, c: 0, tl: 'not a cell' },
    // The first cell that scrolls is below and right of the frozen ones; with nothing frozen there is none.
    { t: 'pane', s: 0, r: 5, c: 0, tl: 'A1' },
    { t: 'pane', s: 0, r: 5, c: 2, tl: 'B9' },
    { t: 'pane', s: 0, r: 0, c: 0, tl: 'B2' },
    { t: 'pane', s: 0, r: 1, c: 0, fv: 7 },
    { t: 'pane', s: 0, r: '1', c: 0 },
  ]) {
    assert.throws(() => decodeStep([bad], a.wb.styles), WireError, JSON.stringify(bad));
  }
});

test('collab panes: an undo of panes someone else changed since is not made, and that is said; both keep the later panes', async () => {
  const { net, random, a, b } = await two();
  assert.ok(act(a, () => freezePanes(a.wb, 0, 1, 0)));
  net.drain(random);
  assert.ok(act(b, () => freezePanes(b.wb, 0, 2, 0)));
  net.drain(random);
  a.session.undo();
  net.drain(random);
  assert.deepEqual(frozenPanes(a.wb, 0), { rows: 2, cols: 0 });
  assert.equal(panes(a), panes(b));
  assert.ok(a.notices.length > 0, 'the undo that was not made is said');
  // Without someone else in between, the undo is made.
  assert.ok(act(a, () => freezePanes(a.wb, 0, 3, 3)));
  net.drain(random);
  a.session.undo();
  net.drain(random);
  assert.deepEqual(frozenPanes(b.wb, 0), { rows: 2, cols: 0 });
  assert.equal(b.wb.sheets[0].view.topLeftCell, 'A3');
});

test('collab panes: frozen and undone in a room, a file whose pane was scrolled is written back as it was, by everybody', async () => {
  const views = '<sheetView workbookViewId="0" topLeftCell="A20"><pane ySplit="2" topLeftCell="A30" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A31" sqref="A31"/></sheetView>';
  const bytes = await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(`<sheetViews>${views}</sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>`) }] });
  const net = new Net();
  const random = seeded(3);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  const viewsOf = async (saved: Uint8Array) => (/<sheetViews>.*<\/sheetViews>/.exec(new TextDecoder().decode(await ZipArchive.read(saved).read('xl/worksheets/sheet1.xml'))) || [''])[0];
  const original = await viewsOf(bytes);
  assert.ok(act(a, () => freezePanes(a.wb, 0, 1, 1)));
  net.drain(random);
  assert.ok(a.session.undo());
  net.drain(random);
  for (const member of [a, b]) {
    assert.deepEqual([member.wb.sheets[0].view.frozenRows, member.wb.sheets[0].view.topLeftCell, member.wb.sheets[0].view.firstVisibleCell], [2, 'A30', 'A20'], member.name);
    assert.equal(await viewsOf(await saveOf(member)), original, member.name);
  }
});
