// Random runs over charts: cells edited, rows and columns inserted and deleted, charts added,
// changed, moved and deleted, undo and redo, with the file saved and read again along the
// way. What has to hold whatever the order: the package is whole, what is read back is the
// model that was written, the references of the charts name the cells the model says (and the
// values the file cached are those cells), and with everything undone the file is the file.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, type Chart, type Sheet, type SheetDrawing, type Workbook } from '../src/index';
import { chartRefAreas, chartRefs } from '../src/model/sheet/chart';
import { addChart, chartProblem, chartRangeRef, moveChart, removeChart, sheetCharts, updateChart, CHART_TYPES, type ChartPatch } from '../src/ops/charts';
import { setComment } from '../src/ops/notes';
import { addConditionalRule, setValidation } from '../src/ops/rules';
import { copySheet, deleteSheet, insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import type { Op } from '../src/ops/ops';
import { drawingAnchorNow } from '../src/render/sheet-drawings';
import { FOREIGN, bare, chartsOf, load, packageProblems, reopen, salesBook, same, save, type Loaded } from './chart-edit-fixtures';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, strRef, valAx } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

/** How many runs per starting file; more when asked for (CHART_SEEDS=200). */
const SEEDS = Math.max(1, Math.min(Number(process.env.CHART_SEEDS) || 12, 5000));

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

/** One random step; undefined when it was refused (which changes nothing). */
function step(doc: Loaded, random: () => number, options: { chartsOnly?: boolean } = {}): string | undefined {
  const wb = doc.wb;
  const sheet = int(random, 0, wb.sheets.length - 1);
  const charts = sheetCharts(wb, sheet);
  let roll = random();
  // One step in five is about the sheets, a note or a rule: what stands around the charts in the same parts.
  if (roll < 0.2 && !options.chartsOnly) {
    const kind = int(random, 0, 7);
    try {
      let ops: Op[];
      if (kind === 0) ops = insertSheet(wb, int(random, 0, wb.sheets.length), { name: `New ${int(random, 0, 999)}` });
      else if (kind === 1) ops = deleteSheet(wb, sheet);
      else if (kind === 2) ops = renameSheet(wb, sheet, pick(random, ['Veri', "It's", 'Yeni Ad', 'S&P <1>', 'Sayfa1', 'Other']));
      else if (kind === 3) ops = copySheet(wb, sheet);
      else if (kind === 4) ops = moveSheet(wb, sheet, int(random, 0, wb.sheets.length - 1));
      else if (kind === 5) ops = setComment(wb, sheet, int(random, 0, 5), int(random, 0, 3), random() < 0.3 ? undefined : `note ${int(random, 0, 9)}`);
      else if (kind === 6) ops = setValidation(wb, sheet, { top: int(random, 0, 3), left: 5, bottom: 6, right: 6 }, random() < 0.3 ? undefined : { type: 'whole', operator: 'greaterThan', formula1: '1' });
      else ops = addConditionalRule(wb, sheet, { top: 0, left: 1, bottom: 5, right: 2 }, { type: 'cellIs', operator: 'greaterThan', value1: '15', format: { fill: '#FFFF00' } });
      if (ops.length === 0) return undefined;
      doc.history.apply(ops, undefined, undefined);
      return ['sheet+', 'sheet-', 'rename', 'copy', 'sheet move', 'note', 'validation', 'conditional'][kind];
    } catch (error) {
      if (error instanceof RangeError) return undefined;
      throw error;
    }
  }
  roll = options.chartsOnly ? 0.36 + roll * 0.64 : random();
  let ops: Op[] = [];
  let what = '';
  try {
    if (roll < 0.2) {
      what = 'cell';
      ops = [{ type: 'setCells', sheet, cells: [{ row: int(random, 0, 6), col: int(random, 0, 4), cell: random() < 0.8 ? { kind: 'number', value: int(random, -50, 500), style: 0 } : { kind: 'string', value: `t${int(random, 0, 9)}`, style: 0 } }] }];
    } else if (roll < 0.35) {
      what = 'rows or columns';
      ops = [{ type: random() < 0.6 ? 'insert' : 'delete', sheet, axis: random() < 0.6 ? 'row' : 'col', at: int(random, 0, 6), count: int(random, 1, 2) }];
    } else if (roll < 0.5) {
      what = 'add';
      const top = int(random, 0, 2);
      const left = int(random, 0, 1);
      ops = addChart(wb, sheet, { type: pick(random, CHART_TYPES), grouping: pick(random, ['standard', 'stacked', 'percentStacked'] as const), range: { top, left, bottom: top + int(random, 1, 4), right: left + int(random, 1, 3) }, title: random() < 0.5 ? `T${int(random, 0, 99)}` : undefined, legend: pick(random, ['r', 'b', 'none'] as const) });
    } else if (roll < 0.75 && charts.length > 0) {
      what = 'update';
      const chart = pick(random, charts);
      const patch: ChartPatch = {};
      const kind = int(random, 0, 6);
      if (kind === 0) patch.title = random() < 0.3 ? null : `Title ${int(random, 0, 99)}\nline`;
      if (kind === 1) patch.legend = pick(random, ['r', 'l', 't', 'b', 'tr', 'none'] as const);
      if (kind === 2) patch.type = pick(random, CHART_TYPES);
      if (kind === 3 && chart.type !== 'pie' && chart.type !== 'doughnut') {
        patch.axisTitles = { category: random() < 0.3 ? null : 'Cat', value: random() < 0.3 ? null : 'Val' };
        patch.axisBounds = random() < 0.5 ? { min: 0, max: 1000 } : { min: null, max: null };
      }
      if (kind === 4) patch.dataLabels = random() < 0.5;
      if (kind === 5 && chart.series.length > 0) {
        const order = chart.series.map((_s, i) => i).sort(() => random() - 0.5);
        const list: NonNullable<ChartPatch['series']> = order.slice(0, Math.max(1, order.length - int(random, 0, 1))).map((from) => (random() < 0.4 ? { from, color: `#${int(random, 0x100000, 0xffffff).toString(16).toUpperCase()}` } : random() < 0.3 ? { from, name: `S${from}` } : { from }));
        if (random() < 0.4) list.push({ values: chartRangeRef(pick(random, wb.sheets).name, { top: 1, left: 1, bottom: 3, right: 1 }), name: 'Added' });
        patch.series = list;
      }
      if (kind === 6 && chart.type !== 'pie' && chart.type !== 'doughnut' && chart.type !== 'scatter' && chart.type !== undefined) patch.grouping = pick(random, ['standard', 'stacked', 'percentStacked'] as const);
      if (chartProblem(wb, sheet, chart.id, patch)) return undefined;
      ops = updateChart(wb, sheet, chart.id, patch);
    } else if (roll < 0.85 && charts.length > 0) {
      what = 'move';
      const top = int(random, 0, 30);
      const left = int(random, 0, 8);
      ops = moveChart(wb, sheet, pick(random, charts).id, { top, left, bottom: top + int(random, 3, 12), right: left + int(random, 2, 6) });
    } else if (roll < 0.9 && charts.length > 0) {
      what = 'remove';
      ops = removeChart(wb, sheet, pick(random, charts).id);
    } else if (roll < 0.96) {
      return doc.history.undo() ? 'undo' : undefined;
    } else {
      return doc.history.redo() ? 'redo' : undefined;
    }
    if (ops.length === 0) return undefined;
    doc.history.apply(ops, undefined, undefined);
    return what;
  } catch (error) {
    // What a command refuses it refuses whole: the checks below see a model that is as it was.
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

/** The numbers the cells of a reference hold, as a chart part caches them. */
function cellNumbers(wb: Workbook, chart: Chart, ref: string): { index: number[]; values: number[] } | undefined {
  const areas = chartRefAreas(wb, ref, chart.live === true);
  if (!areas) return undefined;
  const index: number[] = [];
  const values: number[] = [];
  let at = 0;
  for (const { sheet, range } of areas) {
    for (let row = range.top; row <= range.bottom; row++) {
      for (let col = range.left; col <= range.right; col++, at++) {
        const held = sheet.rows.get(row);
        const cell = held ? held.cells[col] : undefined;
        if (cell && cell.kind === 'number') {
          index.push(at);
          values.push(cell.value as number);
        }
      }
    }
  }
  return { index, values };
}

async function check(doc: Loaded, label: string): Promise<void> {
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
  doc.wb.sheets.forEach((sheet: Sheet, s: number) => {
    const model = (sheet.drawings || []).filter((drawing) => drawing.object.type === 'chart');
    const read = (again.sheets[s].drawings || []).filter((drawing) => drawing.object.type === 'chart');
    assert.equal(read.length, model.length, `${label}: sheet ${s} has as many charts in the file as in the model`);
    // The frames stand in the order of the part: those of the file first, then those that were added.
    const key = (chart: Chart) => stable(bare(chart));
    const left = read.slice();
    model.forEach((drawing: SheetDrawing, i: number) => {
      const chart = chartsOf({ drawings: [drawing] } as Sheet)[0].chart;
      const at = left.findIndex((other) => key(chartsOf({ drawings: [other] } as Sheet)[0].chart) === key(chart) && JSON.stringify(other.anchor) === JSON.stringify(drawingAnchorNow(drawing, sheet.structuralEdits)));
      if (at < 0) {
        // Said with what differs, as far as one of the charts that were read tells.
        const wanted = `${key(chart)} at ${JSON.stringify(drawingAnchorNow(drawing, sheet.structuralEdits))}`;
        const got = left.map((other) => `${key(chartsOf({ drawings: [other] } as Sheet)[0].chart)} at ${JSON.stringify(other.anchor)}`);
        let from = 0;
        while (got[0] !== undefined && from < wanted.length && wanted.charAt(from) === got[0].charAt(from)) from++;
        assert.fail(`${label}: chart ${i} of sheet ${s} is not read back as it is in the model: ...${wanted.slice(Math.max(0, from - 60), from + 120)}... but the file has ...${(got[0] || 'no chart').slice(Math.max(0, from - 60), from + 120)}...`);
      }
      const found = chartsOf({ drawings: [left[at]] } as Sheet)[0].chart;
      left.splice(at, 1);
      // Its references name cells of the workbook, or say that the cells are gone; what the file cached of them is what they hold.
      for (const ref of chartRefs(chart)) assert.ok(ref.indexOf('#REF!') >= 0 || chartRefAreas(doc.wb, ref, chart.live === true), `${label}: the reference ${ref} names cells`);
      chart.groups.forEach((group, g) =>
        group.series.forEach((series, n) => {
          const ref = series.values ? series.values.ref : undefined;
          const cells = ref !== undefined ? cellNumbers(doc.wb, chart, ref) : undefined;
          // A chart of the file that was not changed (moved at most) keeps its part, and with it the cache of the file; every other one caches the cells.
          const written = drawing.from ? drawing.from.object !== drawing.object : drawing.since !== undefined;
          if (cells && written) assert.deepEqual(found.groups[g].series[n].values?.cache, cells, `${label}: the cached values of ${ref} are the cells`);
        }),
      );
    });
  });
}

const frame = twoCell([4, 1], [9, 12], chartFrameOf('rIdC1', 'Grafik 1'));

async function startingFiles(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const plain = await salesBook({ sheets: [{ name: 'Other', xml: '<sheetData/>' }] });
  const foreign = await salesBook({ drawing: frame, charts: [FOREIGN], sheets: [{ name: 'Other', xml: '<sheetData/>' }] });
  // A file the engine wrote charts into.
  const doc = await load(plain);
  doc.history.apply(addChart(doc.wb, 0, { type: 'column', range: { top: 0, left: 0, bottom: 3, right: 2 }, title: 'Own' }), undefined, undefined);
  doc.history.apply(addChart(doc.wb, 1, { type: 'line', series: [{ values: 'Sayfa1!$C$2:$C$4' }] }), undefined, undefined);
  const own = await save(doc);
  // Parts spelled the way other writers spell them: the default namespace for charts (openpyxl), and one with
  // entries of a legend, formats of points and caches that say what the cells say.
  const zip = ZipArchive.read(foreign);
  const decoder = new TextDecoder();
  zip.set('xl/charts/chart1.xml', new TextEncoder().encode(decoder.decode(await zip.read('xl/charts/chart1.xml')).replace(/<(\/?)c:/g, '<$1').replace('xmlns:c=', 'xmlns=')));
  const one = (i: number, col: string) =>
    ser(i, { tx: strRef(`Sayfa1!$${col}$1`, [col]), cat: strRef('Sayfa1!$A$2:$A$4', ['Jan', 'Feb', 'Mar']), val: numRef(`Sayfa1!$${col}$2:$${col}$4`, col === 'B' ? [10, 20, 30] : [30, 20, 10]).replace(/<c:pt idx="(\d)">/g, '<c:pt idx="$1" formatCode="0.0">') });
  const entries = chartXml({
    plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>', one(0, 'B') + one(1, 'C') + one(2, 'B'), '<c:gapWidth val="150"/>') + catAx() + valAx(),
    tail: '<c:legend><c:legendPos val="r"/><c:legendEntry><c:idx val="2"/><c:delete val="1"/></c:legendEntry><c:legendEntry><c:idx val="0"/><c:txPr><a:bodyPr/><a:p><a:endParaRPr/></a:p></c:txPr></c:legendEntry><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>',
  });
  return [
    { name: 'no charts', bytes: plain },
    { name: 'a chart of another application', bytes: foreign },
    { name: 'charts the engine wrote', bytes: own },
    { name: 'a chart in the default namespace', bytes: await zip.write() },
    { name: 'a chart with entries of a legend', bytes: await salesBook({ drawing: frame, charts: [entries], sheets: [{ name: 'Other', xml: '<sheetData/>' }] }) },
  ];
}

test('chart random: whatever is done to cells, rows, columns and charts, the file that is saved is whole and says what the model says', async () => {
  for (const start of await startingFiles()) {
    for (let seed = 1; seed <= SEEDS; seed++) {
      const random = seeded(seed * 7919 + start.name.length);
      const doc = await load(start.bytes);
      const done: string[] = [];
      for (let i = 0; i < 40; i++) {
        const did = step(doc, random);
        if (did) done.push(did);
        if (did && random() < 0.25) await check(doc, `${start.name}, seed ${seed}, after ${done.join(', ')}`);
      }
      await check(doc, `${start.name}, seed ${seed}, at the end (${done.join(', ')})`);
      // Everything undone: the charts and drawings of the file are the bytes of the file.
      while (doc.history.undo());
      const undone = await save(doc);
      const before = ZipArchive.read(start.bytes);
      const after = ZipArchive.read(undone);
      assert.deepEqual(after.names().sort(), before.names().sort(), `${start.name}, seed ${seed}: with everything undone the file has the parts it had`);
      for (const name of before.names()) {
        // Once cells changed or rows moved a sheet is written from the model, and the workbook asks for a calculation.
        if (name === 'xl/workbook.xml' || /^xl\/worksheets\/[^/]*\.xml$/.test(name)) continue;
        same(await after.read(name), await before.read(name), `${start.name}, seed ${seed}, everything undone: ${name}`);
      }
    }
  }
});

test('chart random: charts alone, added, changed, moved, deleted and undone, leave the file byte for byte', async () => {
  for (const start of await startingFiles()) {
    for (let seed = 1; seed <= 8; seed++) {
      const random = seeded(seed * 104729);
      const doc = await load(start.bytes);
      for (let i = 0; i < 30; i++) {
        // No cells and no rows: only steps on charts.
        step(doc, random, { chartsOnly: true });
        if (random() < 0.2) await save(doc);
      }
      while (doc.history.undo());
      same(await save(doc), start.bytes, `${start.name}, seed ${seed}: everything undone`);
    }
  }
});
