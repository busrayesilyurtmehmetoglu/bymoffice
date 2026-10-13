// Charts copied, cut, pasted, duplicated, changed and deleted in random order, among rows and
// columns that are inserted and deleted and sheets that are renamed and deleted: whatever was
// done, the file that is saved is whole, says what the model says, and is read by a reader that
// is not this engine; with everything undone it is the file that was opened.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, type Chart, type Sheet, type SheetDrawing } from '../src/index';
import { chartRefAreas, chartRefs } from '../src/model/sheet/chart';
import { addChart, chartCopyProblem, chartProblem, copyChart, cutChart, duplicateChart, moveChart, pasteChart, removeChart, sheetCharts, updateChart, CHART_TYPES, type ChartCopy, type ChartPatch } from '../src/ops/charts';
import type { Op } from '../src/ops/ops';
import { deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { drawingAnchorNow } from '../src/render/sheet-drawings';
import { UNKNOWN, latinChart, richBook } from './chart-copy-fixtures';
import { pythonHere, readByPython } from './chart-copy-python';
import { bare, chartsOf, load, packageProblems, reopen, salesBook, same, save, type Loaded } from './chart-edit-fixtures';

/** How many runs per starting file; more when asked for (CHART_COPY_SEEDS=500). */
const SEEDS = Math.max(1, Math.min(Number(process.env.CHART_COPY_SEEDS) || 56, 5000));

function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const pick = <T>(random: () => number, list: readonly T[]): T => list[Math.floor(random() * list.length)];
const int = (random: () => number, low: number, high: number): number => low + Math.floor(random() * (high - low + 1));

/** A value as text, with the members of every object in the order of their names. */
function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`;
  if (!value || typeof value !== 'object') return JSON.stringify(value === undefined ? null : value);
  const object = value as Record<string, unknown>;
  return `{${Object.keys(object)
    .filter((name) => object[name] !== undefined)
    .sort()
    .map((name) => `${JSON.stringify(name)}:${stable(object[name])}`)
    .join(',')}}`;
}

interface Run extends Loaded {
  /** What was last copied or cut. */
  held?: ChartCopy;
  /**
   * The chart of the file it was copied from, by the drawing the file has it as, and the sheet it
   * is on; `clean` for as long as nothing but rows, columns and sheets can have changed what
   * that chart says: then a copy that is pasted says what the chart says at that time.
   */
  heldFrom?: { sheet: Sheet; key: SheetDrawing; clean: boolean };
  /** Pastes that were held against the chart they came from. */
  compared?: number;
}

/** A copy that was pasted is held against the chart it was copied from: it shows the cells the chart shows now. */
function compare(doc: Run, target: Sheet): void {
  const from = doc.heldFrom;
  if (!from || !from.clean) return;
  const source = (from.sheet.drawings || []).find((drawing) => (drawing.from || drawing) === from.key);
  const pasted = (target.drawings || [])[(target.drawings || []).length - 1];
  if (!source || !pasted || source.object.type !== 'chart' || pasted.object.type !== 'chart' || doc.wb.sheets.indexOf(from.sheet) < 0) return;
  const now = chartRefs(source.object.chart);
  // On another sheet the references of the copy name the sheet of the chart: held against those that named it already.
  if (target !== from.sheet && now.some((ref) => ref.indexOf('!') < 0)) return;
  assert.deepEqual(chartRefs(pasted.object.chart), now, 'a copy that is pasted says what the chart it was copied from says now');
  doc.compared = (doc.compared || 0) + 1;
}

/** One random step; undefined when it was refused (which changes nothing). */
function step(doc: Run, random: () => number, chartsOnly: boolean): string | undefined {
  const wb = doc.wb;
  const sheet = int(random, 0, wb.sheets.length - 1);
  const charts = sheetCharts(wb, sheet);
  const roll = chartsOnly ? 0.3 + random() * 0.7 : random();
  let ops: Op[] = [];
  let what = '';
  try {
    if (roll < 0.06) {
      what = 'cell';
      ops = [{ type: 'setCells', sheet, cells: [{ row: int(random, 0, 6), col: int(random, 0, 4), cell: { kind: 'number', value: int(random, -50, 500), style: 0 } }] }];
    } else if (roll < 0.18) {
      what = 'rows or columns';
      ops = [{ type: random() < 0.6 ? 'insert' : 'delete', sheet, axis: random() < 0.6 ? 'row' : 'col', at: int(random, 0, 6), count: int(random, 1, 2) }];
    } else if (roll < 0.24) {
      what = 'rename';
      ops = renameSheet(wb, sheet, pick(random, ['Veri', "It's", 'Yeni Ad', 'S&P <1>', 'Sayfa1', 'Other', 'Çizelge']));
    } else if (roll < 0.27) {
      what = 'sheet-';
      ops = deleteSheet(wb, sheet);
    } else if (roll < 0.3) {
      what = 'sheet+';
      ops = insertSheet(wb, int(random, 0, wb.sheets.length), { name: `New ${int(random, 0, 999)}` });
    } else if (roll < 0.36) {
      what = 'add';
      const top = int(random, 0, 2);
      ops = addChart(wb, sheet, { type: pick(random, CHART_TYPES), range: { top, left: 0, bottom: top + int(random, 1, 4), right: int(random, 1, 3) }, title: random() < 0.5 ? `T${int(random, 0, 99)}` : undefined });
    } else if (roll < 0.46 && charts.length > 0) {
      const id = pick(random, charts).id;
      if (chartCopyProblem(wb, sheet, id)) return undefined;
      doc.held = copyChart(wb, sheet, id);
      const taken = (wb.sheets[sheet].drawings || [])[id];
      doc.heldFrom = taken.object.type === 'chart' && taken.object.part !== undefined ? { sheet: wb.sheets[sheet], key: taken.from || taken, clean: true } : undefined;
      return 'copy';
    } else if (roll < 0.62 && doc.held) {
      what = 'paste';
      ops = pasteChart(wb, sheet, doc.held, random() < 0.4 ? { row: int(random, 0, 40), col: int(random, 0, 12) } : undefined).ops;
    } else if (roll < 0.7 && charts.length > 0) {
      what = 'duplicate';
      ops = duplicateChart(wb, sheet, pick(random, charts).id).ops;
    } else if (roll < 0.75 && charts.length > 0) {
      what = 'cut';
      const cut = cutChart(wb, sheet, pick(random, charts).id);
      doc.held = cut.copy;
      doc.heldFrom = undefined;
      ops = cut.ops;
    } else if (roll < 0.84 && charts.length > 0) {
      what = 'update';
      const chart = pick(random, charts);
      const patch: ChartPatch = {};
      const kind = int(random, 0, 4);
      if (kind === 0) patch.title = random() < 0.3 ? null : `Başlık ${int(random, 0, 99)}`;
      if (kind === 1) patch.legend = pick(random, ['r', 'l', 't', 'b', 'tr', 'none'] as const);
      if (kind === 2) patch.type = pick(random, CHART_TYPES);
      if (kind === 3) patch.dataLabels = random() < 0.5;
      if (kind === 4 && chart.series.length > 0) {
        const order = chart.series.map((_s, i) => i).sort(() => random() - 0.5);
        patch.series = order.slice(0, Math.max(1, order.length - int(random, 0, 1))).map((from) => (random() < 0.4 ? { from, color: `#${int(random, 0x100000, 0xffffff).toString(16).toUpperCase()}` } : { from }));
      }
      if (chartProblem(wb, sheet, chart.id, patch)) return undefined;
      ops = updateChart(wb, sheet, chart.id, patch);
    } else if (roll < 0.88 && charts.length > 0) {
      what = 'move';
      const top = int(random, 0, 30);
      const left = int(random, 0, 8);
      ops = moveChart(wb, sheet, pick(random, charts).id, { top, left, bottom: top + int(random, 3, 12), right: left + int(random, 2, 6) });
    } else if (roll < 0.92 && charts.length > 0) {
      what = 'remove';
      ops = removeChart(wb, sheet, pick(random, charts).id);
    } else if (roll < 0.97) {
      // A step that is undone may be one that changed the chart before it was copied.
      if (doc.heldFrom) doc.heldFrom.clean = false;
      return doc.history.undo() ? 'undo' : undefined;
    } else {
      if (doc.heldFrom) doc.heldFrom.clean = false;
      return doc.history.redo() ? 'redo' : undefined;
    }
    if (ops.length === 0) return undefined;
    const target = wb.sheets[sheet];
    doc.history.apply(ops, undefined, undefined);
    if (what === 'update' && doc.heldFrom) doc.heldFrom.clean = false;
    if (what === 'paste') compare(doc, target);
    return what;
  } catch (error) {
    // What a command refuses it refuses whole: the checks below see a model that is as it was.
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

/** Saves, and holds the file against the model; returns the file and the charts each worksheet has. */
async function check(doc: Run, label: string): Promise<{ bytes: Uint8Array; charts: number[] }> {
  let saved: Uint8Array;
  try {
    saved = await save(doc);
  } catch (error) {
    // Whatever was applied as an operation can be written: a save never throws.
    assert.fail(`${label}: the save throws: ${(error as Error).message}`);
  }
  assert.deepEqual(await packageProblems(saved), [], `${label}: the package is whole`);
  const again = await reopen(saved);
  assert.equal(again.sheets.length, doc.wb.sheets.length, `${label}: the sheets`);
  const counts: number[] = [];
  doc.wb.sheets.forEach((sheet: Sheet, s: number) => {
    const model = (sheet.drawings || []).filter((drawing) => drawing.object.type === 'chart');
    const read = (again.sheets[s].drawings || []).filter((drawing) => drawing.object.type === 'chart');
    assert.equal(read.length, model.length, `${label}: sheet ${s} has as many charts in the file as in the model`);
    if (sheet.kind === 'worksheet') counts.push(model.length);
    const key = (chart: Chart) => stable(bare(chart));
    const left = read.slice();
    model.forEach((drawing: SheetDrawing, i: number) => {
      const chart = chartsOf({ drawings: [drawing] } as Sheet)[0].chart;
      const place = JSON.stringify(drawingAnchorNow(drawing, sheet.structuralEdits));
      const at = left.findIndex((other) => key(chartsOf({ drawings: [other] } as Sheet)[0].chart) === key(chart) && JSON.stringify(other.anchor) === place);
      if (at < 0) {
        const wanted = `${key(chart)} at ${place}`;
        const got = left.map((other) => `${key(chartsOf({ drawings: [other] } as Sheet)[0].chart)} at ${JSON.stringify(other.anchor)}`);
        let from = 0;
        while (got[0] !== undefined && from < wanted.length && wanted.charAt(from) === got[0].charAt(from)) from++;
        assert.fail(`${label}: chart ${i} of sheet ${s} is not read back as it is in the model: ...${wanted.slice(Math.max(0, from - 60), from + 120)}... but the file has ...${(got[0] || 'no chart').slice(Math.max(0, from - 60), from + 120)}...`);
      }
      left.splice(at, 1);
      // A chart that follows names cells of the workbook, or says that they are gone.
      const object = chartsOf({ drawings: [drawing] } as Sheet)[0];
      if (chart.live) for (const ref of chartRefs(chart)) assert.ok(ref.indexOf('#REF!') >= 0 || ref.indexOf('!') < 0 || chartRefAreas(doc.wb, ref, true), `${label}: the reference ${ref} names cells`);
      // A copy that is made from a part is a chart that is in no file.
      if (object.source) assert.equal(object.part, undefined, `${label}: a copy has no part of the file`);
    });
  });
  return { bytes: saved, charts: counts };
}

async function startingFiles(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const rich = await richBook();
  const doc = await load(rich);
  doc.history.apply(addChart(doc.wb, 1, { type: 'line', series: [{ values: 'Sayfa1!$C$2:$C$4' }], title: 'Own' }), undefined, undefined);
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: { top: 0, left: 0, bottom: 3, right: 2 } }), undefined, undefined);
  return [
    { name: 'charts with styles, colours and a picture', bytes: rich },
    { name: 'a chart with shapes drawn over it', bytes: await richBook({ shapes: true }) },
    { name: 'a chart part that declares Latin-1', bytes: await richBook({ first: latinChart() }) },
    { name: 'charts of the file and charts the engine wrote', bytes: await save(doc) },
    { name: 'a chart that is not copied', bytes: await richBook({ first: UNKNOWN }) },
    { name: 'no charts', bytes: await salesBook({ sheets: [{ name: 'Other', xml: '<sheetData/>' }] }) },
  ];
}

/** The last file of every run, for the reader that is not this engine. */
const kept: { name: string; bytes: Uint8Array; charts: number[] }[] = [];

/** The starting files by what they are; each has a test of its own, so that none runs long. */
const STARTS = ['charts with styles, colours and a picture', 'a chart with shapes drawn over it', 'a chart part that declares Latin-1', 'charts of the file and charts the engine wrote', 'a chart that is not copied', 'no charts'];
let pastes = 0;
let compared = 0;

for (const wanted of STARTS) {
  test(`chart copy random: whatever is copied, cut, pasted and undone among other edits, the file is whole and says what the model says (${wanted})`, async () => {
    const start = (await startingFiles()).find((file) => file.name === wanted);
    assert.ok(start, 'the starting file is made');
    for (let seed = 1; seed <= SEEDS; seed++) {
      const random = seeded(seed * 7919 + start.name.length * 31);
      const doc: Run = await load(start.bytes);
      const done: string[] = [];
      for (let i = 0; i < 36; i++) {
        const did = step(doc, random, false);
        if (!did) continue;
        done.push(did);
        if (did === 'paste' || did === 'duplicate') pastes++;
        if (random() < 0.2) await check(doc, `${start.name}, seed ${seed}, after ${done.join(', ')}`);
      }
      const last = await check(doc, `${start.name}, seed ${seed}, at the end (${done.join(', ')})`);
      kept.push({ name: `${start.name} ${seed}`, bytes: last.bytes, charts: last.charts });
      compared += doc.compared || 0;
      // Everything undone: the charts and drawings of the file are the bytes of the file.
      while (doc.history.undo());
      const undone = await save(doc);
      const before = ZipArchive.read(start.bytes);
      const after = ZipArchive.read(undone);
      assert.deepEqual(after.names().sort(), before.names().sort(), `${start.name}, seed ${seed}: with everything undone the file has the parts it had (${done.join(', ')})`);
      for (const name of before.names()) {
        // Once cells changed or rows moved a sheet is written from the model, and the workbook asks for a calculation.
        if (name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
        same(await after.read(name), await before.read(name), `${start.name}, seed ${seed}, everything undone: ${name}`);
      }
    }
  });
}

test('chart copy random: charts alone, copied, pasted, cut, changed and undone, leave the file byte for byte', async () => {
  let steps = 0;
  for (const start of await startingFiles()) {
    for (let seed = 1; seed <= Math.max(8, Math.floor(SEEDS / 4)); seed++) {
      const random = seeded(seed * 104729 + start.name.length);
      const doc: Run = await load(start.bytes);
      for (let i = 0; i < 30; i++) {
        const did = step(doc, random, true);
        // Saved in between as well: what a save made for a copy goes again with the copy.
        if (did && random() < 0.3) await save(doc);
        if (did === 'paste' || did === 'duplicate' || did === 'cut') steps++;
      }
      while (doc.history.undo());
      same(await save(doc), start.bytes, `${start.name}, seed ${seed}: everything undone`);
    }
  }
  assert.ok(steps > 100, `charts were pasted, duplicated and cut: ${steps} times`);
});

test('chart copy random: the files of the runs are read by a reader that is not this engine', async (t) => {
  if (!pythonHere()) {
    t.skip('python3 is not available');
    return;
  }
  assert.ok(kept.length >= SEEDS * STARTS.length, 'the runs kept their files');
  assert.ok(pastes > SEEDS * STARTS.length, `charts were pasted and duplicated: ${pastes} times`);
  assert.ok(compared > SEEDS, `pastes were held against the charts they came from: ${compared} times`);
  const found = readByPython(kept, 'random');
  assert.deepEqual(found.problems.slice(0, 40), []);
  assert.equal(found.files, kept.length);
  let charts = 0;
  for (const file of kept) for (const count of file.charts) charts += count;
  t.diagnostic(`${found.files} files with ${charts} charts read by Python${found.openpyxl ? ' and openpyxl' : ' (openpyxl is not installed: charts were not counted by it)'}`);
});
