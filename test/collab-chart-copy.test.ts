// Charts copied and pasted while a workbook is edited together: the paste travels as a command
// that names the chart it copies, every participant makes the copy from the chart part of the
// file it read itself, and what cannot be said that way is not done by anybody.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive } from '../src/index';
import { chartRefs } from '../src/model/sheet/chart';
import { copyChart, cutChart, duplicateChart, pasteChart, pasteChartOf, removeChart, updateChart } from '../src/ops/charts';
import type { Op } from '../src/ops/ops';
import { richBook } from './chart-copy-fixtures';
import { pythonHere, readByPython } from './chart-copy-python';
import { bare, chartsOf, packageProblems } from './chart-edit-fixtures';
import { Net, act, assertSame, join, saveOf, seeded, type Member } from './collab-fixtures';

const codes = (member: Member): string[] => member.notices.map((notice) => notice.code);

/** The charts of a workbook as what is said of them: what they show, what they are made from, where they stand. */
const charts = (member: Member): string =>
  JSON.stringify(
    member.wb.sheets.map((sheet) =>
      (sheet.drawings || []).map((drawing) => (drawing.object.type === 'chart' ? { chart: bare(drawing.object.chart), source: drawing.object.source, frame: drawing.object.frame, known: drawing.object.known, name: drawing.object.name, anchor: drawing.anchor } : drawing.object.type)),
    ),
  );

async function two(seed = 1): Promise<{ net: Net; random: () => number; a: Member; b: Member; bytes: Uint8Array }> {
  const bytes = await richBook();
  const net = new Net();
  const random = seeded(seed);
  const a = await join(net, 'a', bytes);
  const b = await join(net, 'b', bytes);
  net.drain(random);
  return { net, random, a, b, bytes };
}

test('collab chart copy: a chart that is duplicated and one that is pasted on another sheet reach everybody, each made from the part of the file they read', async (t) => {
  const { net, random, a, b } = await two();
  assert.equal(act(a, () => duplicateChart(a.wb, 0, 0).ops), true);
  net.drain(random);
  assert.equal(act(b, () => pasteChart(b.wb, 1, copyChart(b.wb, 0, 1), { row: 3, col: 2 }).ops), true);
  net.drain(random);
  assert.equal(act(a, () => pasteChart(a.wb, 1, copyChart(a.wb, 0, 0)).ops), true);
  net.drain(random);
  assertSame([a, b], 'after the copies');
  assert.equal(charts(a), charts(b), 'the charts are the same');
  assert.deepEqual([chartsOf(a.wb.sheets[0]).length, chartsOf(a.wb.sheets[1]).length], [3, 2]);
  assert.deepEqual([codes(a), codes(b)], [[], []], 'nobody was told anything');
  // What b got of a's copy is made from b's own file: the same part, by its name.
  assert.deepEqual(chartsOf(b.wb.sheets[0])[2].source, { part: '/xl/charts/chart1.xml' });
  assert.deepEqual(chartRefs(chartsOf(a.wb.sheets[1])[0].chart), ['Sayfa1!$B$1', 'Sayfa1!$A$2:$A$4', 'Sayfa1!$B$2:$B$4'], 'on the other sheet the copy names the sheet it shows');
  // A copy is changed like any chart, by whoever.
  assert.equal(act(b, () => updateChart(b.wb, 0, 2, { title: 'Kopya' })), true);
  net.drain(random);
  assert.equal(charts(a), charts(b));
  const files: { name: string; bytes: Uint8Array; charts: number[] }[] = [];
  const parts: string[][] = [];
  for (const member of [a, b]) {
    const saved = await saveOf(member);
    assert.deepEqual(await packageProblems(saved), [], `${member.name}: the package is whole`);
    files.push({ name: member.name, bytes: saved, charts: [3, 2] });
    const zip = ZipArchive.read(saved);
    const list: string[] = [];
    for (const name of zip.names().filter((entry) => /^xl\/(charts|media|drawings)\//.test(entry)).sort()) list.push(`${name}\n${new TextDecoder('latin1').decode(await zip.read(name))}`);
    parts.push(list);
  }
  assert.deepEqual(parts[1], parts[0], 'the two files hold the same charts, part by part');
  if (!pythonHere()) {
    t.diagnostic('python3 is not available: the files were not read by another reader');
    return;
  }
  const found = readByPython(files, 'collab');
  assert.deepEqual(found.problems, []);
});

test('collab chart copy: a paste whose chart somebody else took away at the same time is dropped, and its author is told', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => pasteChart(b.wb, 1, copyChart(b.wb, 0, 1)).ops);
  act(a, () => removeChart(a.wb, 0, 0));
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assertSame([a, b], 'after the chart went');
  assert.equal(charts(a), charts(b));
  assert.deepEqual(codes(b), ['dropped']);
  assert.deepEqual([chartsOf(a.wb.sheets[0]).length, chartsOf(a.wb.sheets[1]).length], [1, 0], 'the copy of the chart that moved up was not made in its place');
});

test('collab chart copy: of two copies made at the same time the second gives way, and both workbooks hold the first', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => pasteChart(b.wb, 1, copyChart(b.wb, 0, 0)).ops);
  act(a, () => duplicateChart(a.wb, 1 - 1, 1).ops);
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assertSame([a, b], 'after both');
  assert.equal(charts(a), charts(b));
  assert.deepEqual(codes(b), ['dropped'], 'the place a chart is named by may be another chart\'s once the charts of any sheet changed');
  assert.deepEqual([chartsOf(a.wb.sheets[0]).length, chartsOf(a.wb.sheets[1]).length], [3, 0]);
});

test('collab chart copy: a paste beside rows that somebody else put in at the same time is dropped', async () => {
  const { net, random, a, b } = await two();
  b.transport.suspend();
  net.drain(random);
  act(b, () => duplicateChart(b.wb, 0, 0).ops);
  act(a, () => [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }] as Op[]);
  net.drain(random);
  b.transport.resume();
  net.drain(random);
  assertSame([a, b], 'after the row');
  assert.equal(charts(a), charts(b));
  assert.deepEqual(codes(b), ['dropped']);
  assert.equal(chartsOf(b.wb.sheets[0]).length, 2);
});

test('collab chart copy: a copy that is no longer the chart it was taken from does not travel, and is pasted by nobody', async () => {
  const { net, random, a, b } = await two();
  const held = copyChart(a.wb, 0, 0);
  // The chart is changed after it was copied: what is held is what it was, which no command says.
  act(a, () => updateChart(a.wb, 0, 0, { title: 'Sonra' }));
  net.drain(random);
  const pasted = pasteChart(a.wb, 1, held);
  assert.equal(pasted.ops[0].type === 'setDrawings' && pasted.ops[0].by, undefined, 'no command stands for it');
  assert.equal(a.session.submit(pasted.ops).skipped, 1);
  net.drain(random);
  assert.equal(charts(a), charts(b));
  assert.equal(chartsOf(a.wb.sheets[1]).length, 0, 'it was pasted nowhere');
  // The same for a chart that was cut: it is gone for everybody, and its copy travels with no command.
  const cut = cutChart(a.wb, 0, 1);
  act(a, () => cut.ops);
  net.drain(random);
  assert.equal(pasteChart(a.wb, 1, cut.copy).ops[0].type === 'setDrawings' && (pasteChart(a.wb, 1, cut.copy).ops[0] as { by?: unknown }).by, undefined);
  assert.equal(charts(a), charts(b));
  // Copied again from the chart as it is, it travels.
  assert.equal(act(a, () => pasteChart(a.wb, 1, copyChart(a.wb, 0, 0)).ops), true);
  net.drain(random);
  assert.equal(charts(a), charts(b));
  assert.equal(chartsOf(b.wb.sheets[1]).length, 1);
});

test('collab chart copy: the command names a sheet and a chart that are there, and a place that is a cell', async () => {
  const { a } = await two();
  assert.equal(pasteChartOf(a.wb, 1, 'Sayfa1', 0, null).length, 1);
  assert.equal(pasteChartOf(a.wb, 1, 'sayfa1', 1, { row: 2, col: 2 }).length, 1, 'sheets are named without regard to case');
  for (const args of [
    ['Nowhere', 0, null],
    [7, 0, null],
    ['Sayfa1', 9, null],
    ['Sayfa1', '0', null],
    ['Sayfa1', 0.5, null],
    ['Sayfa1', 0, { row: -1, col: 0 }],
    ['Sayfa1', 0, 'A1'],
    [null, null, null],
  ] as unknown[][]) {
    assert.throws(() => pasteChartOf(a.wb, 1, args[0] as never, args[1] as never, args[2] as never), RangeError, JSON.stringify(args));
  }
  assert.throws(() => pasteChartOf(a.wb, 9, 'Sayfa1', 0, null), RangeError);
});
