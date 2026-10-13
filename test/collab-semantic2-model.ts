// A second model of what a workbook that is edited together has to end up as, written by a reviewer without
// looking at the first one (collab-semantic.test.ts) and without reading how changes are rebased: rows,
// columns, cells and sheets carry markers, every step is noted by the markers its author saw, and what the
// workbook has to hold follows from the order of the relay alone. run() answers with what does not hold.
import { buildXlsx } from '../src/formats/xlsx/build';
import { open, type Workbook } from '../src/index';
import { MemoryServer } from '../src/collab/server';
import { CollabSession } from '../src/collab/session';
import { contentVersion } from '../src/collab/protocol';
import { setCellInput } from '../src/ops/commands';
import { deleteColumns, deleteRows } from '../src/ops/filter';
import { mergeRange, setRowHeight } from '../src/ops/format';
import { setComment, cellNow } from '../src/ops/notes';
import { deleteSheet, insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import type { Op } from '../src/ops/ops';
import { bookOf } from './chart-edit-fixtures';

export function rng(seed: number) { let s = seed >>> 0; return () => { s = (s + 0x9e3779b9) >>> 0; let t = s; t = Math.imul(t ^ (t >>> 16), 0x85ebca6b); t = Math.imul(t ^ (t >>> 13), 0xc2b2ae35); return ((t ^ (t >>> 16)) >>> 0) / 4294967296; }; }
const ri = (r: () => number, a: number, b: number) => a + Math.floor(r() * (b - a + 1));
const colName = (c: number) => { let s = ''; c++; while (c > 0) { s = String.fromCharCode(65 + ((c - 1) % 26)) + s; c = Math.floor((c - 1) / 26); } return s; };

let nextId = 1;
const rowHome = new Map<number, number>(); // row marker id -> sheet uid
const colHome = new Map<number, number>();
const innerHome = new Map<number, number>();
const colSide = new Map<number, 'in' | 'out'>();
const sheetName0 = new Map<number, string>();
const innerColOf = new Map<number, number>();

function sheetRows(uid: number, n: number): (string | null)[][] {
  const rows: (string | null)[][] = [];
  const hdr: (string | null)[] = [`R${uid}`];
  for (let c = 1; c <= 5; c++) { const id = nextId++; colHome.set(id, uid); colSide.set(id, c <= 3 ? 'in' : 'out'); if (c === 3) innerColOf.set(uid, id); hdr.push(`C${id}`); }
  rows.push(hdr);
  for (let r = 1; r <= n; r++) { const id = nextId++; const m = nextId++; rowHome.set(id, uid); innerHome.set(m, uid); rows.push([`R${id}`, null, null, `M${m}`, null, null]); }
  return rows;
}

interface Scan { index: number; name: string; uid: number; rowPos: Map<number, number>; colPos: Map<number, number>; innerPos: Map<number, number>; rowAt: Map<number, number>; innerAt: Map<number, number>; colAt: Map<number, number>; innerCol: number; lastCol: number; lastRow: number; tokens: Map<string, [number, number]>; formulas: Map<string, { row: number; col: number; text: string }>; heights: Map<number, number>; notes: Map<string, [number, number]>; merges: number[][]; problems: string[]; rowOrder: number[]; colOrder: number[] }

export function scan(wb: Workbook): Scan[] {
  return wb.sheets.map((sheet, index) => {
    const sc: Scan = { index, name: sheet.name, uid: -1, rowPos: new Map(), colPos: new Map(), innerPos: new Map(), rowAt: new Map(), innerAt: new Map(), colAt: new Map(), innerCol: -1, lastCol: sheet.colCount - 1, lastRow: 0, tokens: new Map(), formulas: new Map(), heights: new Map(), notes: new Map(), merges: sheet.merges.map((m) => [m.top, m.left, m.bottom, m.right]), problems: [], rowOrder: [], colOrder: [] };
    const mcols = new Set<number>(); const icols = new Set<number>(); const hrows = new Set<number>();
    for (const [r, row] of sheet.rows) {
      if (row.cells.length - 1 > sc.lastCol) sc.lastCol = row.cells.length - 1;
      if (row.height !== undefined) sc.heights.set(r, row.height);
      row.cells.forEach((cell, c) => {
        if (!cell) return;
        if (r > sc.lastRow) sc.lastRow = r;
        if (cell.formula) { const m = /\+(\d{7,})$/.exec(cell.formula.text); if (m) { if (sc.formulas.has(m[1]) || sc.tokens.has('F' + m[1])) sc.problems.push(`formula tag ${m[1]} twice`); sc.formulas.set(m[1], { row: r, col: c, text: cell.formula.text }); sc.tokens.set('F' + m[1], [r, c]); } else sc.problems.push(`untagged formula ${cell.formula.text} at ${r},${c}`); return; }
        if (cell.kind !== 'string') { if (cell.kind !== 'blank') sc.problems.push(`odd cell ${cell.kind} at ${r},${c}`); return; }
        const v = String(cell.value); const id = Number(v.slice(1));
        const dup = (map: Map<number, number>) => { if (map.has(id)) sc.problems.push(`marker ${v} twice`); };
        if (v[0] === 'R') { dup(sc.rowPos); sc.rowPos.set(id, r); if (sc.rowAt.has(r)) sc.problems.push(`two row markers in row ${r}`); sc.rowAt.set(r, id); mcols.add(c); }
        else if (v[0] === 'C') { dup(sc.colPos); sc.colPos.set(id, c); if (sc.colAt.has(c)) sc.problems.push(`two col markers in col ${c}`); sc.colAt.set(c, id); hrows.add(r); }
        else if (v[0] === 'M') { dup(sc.innerPos); sc.innerPos.set(id, r); if (sc.innerAt.has(r)) sc.problems.push(`two inner markers in row ${r}`); sc.innerAt.set(r, id); icols.add(c); }
        else if (v[0] === 'V') { if (sc.tokens.has(v)) sc.problems.push(`value ${v} twice`); sc.tokens.set(v, [r, c]); }
        else sc.problems.push(`odd text ${v}`);
      });
    }
    if (mcols.size !== 1 || !mcols.has(0)) sc.problems.push(`row markers in columns ${[...mcols]}`);
    if (icols.size > 1) sc.problems.push(`inner markers in columns ${[...icols]}`);
    if (hrows.size !== 1 || !hrows.has(0)) sc.problems.push(`col markers in rows ${[...hrows]}`);
    const homes = new Set<number>(); sc.rowPos.forEach((_p, id) => homes.add(id === sc.rowAt.get(0) && !rowHome.has(id) ? id : (rowHome.get(id) as number)));
    const head = sc.rowAt.get(0);
    sc.uid = head === undefined ? -1 : head;
    { const ic = innerColOf.get(sc.uid); sc.innerCol = ic !== undefined && sc.colPos.has(ic) ? sc.colPos.get(ic)! : -1; if (sc.innerCol < 0) sc.problems.push('inner column is gone'); if (icols.size && !icols.has(sc.innerCol)) sc.problems.push('inner markers outside inner column'); }
    sc.rowPos.forEach((_p, id) => { if (id !== head && rowHome.get(id) !== head) sc.problems.push(`row marker R${id} of sheet ${rowHome.get(id)} on sheet ${head}`); });
    sc.innerPos.forEach((_p, id) => { if (innerHome.get(id) !== head) sc.problems.push(`inner marker M${id} on foreign sheet`); });
    sc.colPos.forEach((_p, id) => { if (colHome.get(id) !== head) sc.problems.push(`col marker C${id} on foreign sheet`); });
    sc.rowOrder = [...sc.rowPos.entries()].sort((a, b) => a[1] - b[1]).map((e) => e[0]);
    sc.colOrder = [...sc.colPos.entries()].sort((a, b) => a[1] - b[1]).map((e) => e[0]);
    for (const comment of sheet.comments || []) { const at = cellNow(sheet, comment); if (at) { if (sc.notes.has(String(comment.text))) sc.problems.push(`note ${comment.text} twice`); sc.notes.set(String(comment.text), [at.row, at.col]); } }
    return sc;
  });
}

export function dump(wb: Workbook): string {
  return JSON.stringify(wb.sheets.map((sheet) => {
    const cells: unknown[] = [];
    for (const r of [...sheet.rows.keys()].sort((a, b) => a - b)) { const row = sheet.rows.get(r)!; row.cells.forEach((cell, c) => { if (cell && (cell.kind !== 'blank' || cell.formula)) cells.push([r, c, cell.kind, cell.formula ? '=' + cell.formula.text : cell.value]); }); if (row.height !== undefined) cells.push([r, 'h', row.height]); }
    const notes = (sheet.comments || []).map((c) => { const at = cellNow(sheet, c); return at ? [at.row, at.col, c.text] : null; }).filter(Boolean).sort((a: any, b: any) => a[0] - b[0] || a[1] - b[1]);
    return [sheet.name, cells, sheet.merges.map((m) => [m.top, m.left, m.bottom, m.right]).sort((a, b) => a[0] - b[0] || a[1] - b[1]), notes];
  }));
}

type RowKey = string; // 'R12' or 'M13'
interface Intent {
  kind: string; sheet: number; // sheet uid
  cell?: { rk: RowKey; col: number; token: string };
  note?: { rk: RowKey; col: number; text: string };
  height?: { row: number; h: number };
  f?: { form: string; tag: string; rows?: number[]; cols?: number[]; target: number; snap?: number[] };
  del?: { rows?: number[]; inners?: number[]; cols?: number[] };
  ins?: { rows?: number[]; inners?: number[]; cols?: number[] };
  merge?: { rows: number[]; cols: number[] };
  rowSnap?: number[]; colSnap?: number[];
  sdel?: number; sname?: string; snew?: number;
  concurrentWith?: number; // rev known to author at creation
}

export interface Config { seed: number; members: number; ticks: number; sorts: boolean; formulas: boolean; burst: boolean; cuts: boolean; sheets: boolean; cols: boolean; merges: boolean; quiet?: boolean; trace?: boolean }
export interface Result { classes: string[]; violations: string[]; made: number; placed: number; dropped: number; quiet: number; overdrop: string[]; collateral: string[]; byKind: Record<string, { made: number; dropped: number; quiet: number }>; partial: number }

export async function run(cfg: Config): Promise<Result> {
  nextId = 1; rowHome.clear(); colHome.clear(); innerHome.clear(); colSide.clear(); sheetName0.clear(); innerColOf.clear();
  const r = rng(cfg.seed);
  const u1 = nextId++; const rows1 = sheetRows(u1, 10); const u2 = nextId++; const rows2 = sheetRows(u2, 6);
  const book = bookOf(rows1, [{ name: 'Other', rows: rows2 }]);
  sheetName0.set(u1, 'Data'); sheetName0.set(u2, 'Other');
  const bytes = await buildXlsx(book);
  const lanes = new Map<string, (() => void)[]>();
  const server = new MemoryServer({ deliver: (fn, lane) => { let q = lanes.get(lane); if (!q) lanes.set(lane, (q = [])); q.push(fn); } });
  let timers: { cb: () => void; off: boolean }[] = [];
  const dropInfo = new Map<string, any>(); const lastText = new Map<any, string>(); const pendingJson = new Map<any, string>();
  const dropAt = new Map<string, number>(); const pendingSeen = new Map<any, Set<string>>(); let watch: any[] = [];
  const stepNet = (): boolean => { const busy = [...lanes.values()].filter((q) => q.length > 0); if (!busy.length) return false; busy[Math.floor(r() * busy.length)].shift()!();
    if (watch.length) look();
    for (const m of watch) { const now = new Set<string>((m.session as any).pendingSteps.map((x: any) => x.id)); const was = pendingSeen.get(m) || new Set<string>(); was.forEach((id) => { if (!now.has(id)) { dropAt.set(id, m.session.rev); dropInfo.set(id, { last: lastText.get(m), before: pendingJson.get(m), who: m.name }); } }); pendingSeen.set(m, now); pendingJson.set(m, JSON.stringify((m.session as any).pendingSteps.map((x: any) => [x.id, x.recs.map((q: any) => q.w)]))); }
    return true; };
  // observer
  const order: string[] = [];
  const obs = server.connect();
  obs.onMessage((text) => { const m = JSON.parse(text); if (m.t === 'op') order.push(m.id); });
  obs.onState!((st) => { if (st === 'open') obs.send(JSON.stringify({ t: 'join', pv: 1, key: 'obs', v: [contentVersion(bytes)] })); });
  while (stepNet());
  interface M { name: string; wb: Workbook; session: CollabSession; t: any; notices: { code: string; count?: number }[]; cut: boolean; key: string }
  const members: M[] = [];
  for (let i = 0; i < cfg.members; i++) {
    const file = await open(bytes); const wb = await file.workbook(); const t0 = server.connect(); const holder: any = {}; const t: any = { send: (x: string) => { if ((cfg as any).log) { const q = JSON.parse(x); if (q.t === 'op') console.log(`  ${holder.m.name} SENDS base ${q.base} ${q.id.split(':')[1]} ${JSON.stringify(q.ops).slice(0, 260)}`); } t0.send(x); }, onMessage: (l: any) => t0.onMessage((text) => { lastText.set(holder.m, text); if ((cfg as any).log) { const q = JSON.parse(text); if (q.t === 'op' || q.t === 'nack') console.log(`  ${holder.m.name} GETS ${q.t} rev ${q.rev} ${q.id.split(':')[1]} pending ${holder.m.session.pending}`); } l(text); if ((cfg as any).follow) { const q = JSON.parse(text); const st = (holder.m.session as any).pendingSteps.find((x: any) => x.id.split(':')[1] === String((cfg as any).follow[1])); if (q.t === 'op' && holder.m.name === (cfg as any).follow[0]) console.log(`rev ${q.rev} remote ${JSON.stringify(q.ops).slice(0, 200)}\n      -> mine ${st ? JSON.stringify(st.recs.map((y: any) => y.w)) : 'gone'}  col0 ${[...holder.m.wb.sheets[(cfg as any).follow[2]].rows.entries()].sort((a: any, b: any) => a[0] - b[0]).map((e: any) => e[1].cells[0] ? e[1].cells[0].value : '-').join(' ')}`); } if ((cfg as any).log) { const q = JSON.parse(text); if (q.t === 'op') console.log(`     -> sheets ${holder.m.wb.sheets.map((x: any) => x.name).join(',')} pending ${JSON.stringify((holder.m.session as any).pendingSteps.map((x: any) => x.recs.map((y: any) => y.w))).slice(0, 400)}`); } }), onState: (l: any) => t0.onState!(l), close: () => t0.close(), suspend: () => t0.suspend(), resume: () => t0.resume() };
    const m: M = { name: 'm' + i, wb, session: undefined as any, t, notices: [], cut: false, key: '' };
    m.session = new CollabSession({ workbook: () => m.wb, changed: () => undefined, notice: (n) => m.notices.push(n) }, { transport: t, versions: [contentVersion(bytes)], user: { name: m.name }, timer: (cb) => { const e = { cb, off: false }; timers.push(e); return () => { e.off = true; }; } });
    m.key = (m.session as any).tag; holder.m = m; members.push(m);
  }
  while (stepNet());
  watch = members;
  const goneAt = new Map<string, number>(); const bornAt = new Map<string, number>(); let seenRev = 0; let seenSet = new Set<string>();
  const wfile = await open(bytes); const wwb = await wfile.workbook(); const wt = server.connect();
  const ws = new CollabSession({ workbook: () => wwb, changed: () => { }, notice: () => undefined }, { transport: wt, versions: [contentVersion(bytes)], timer: () => () => undefined });
  const look = () => { if (ws.rev === seenRev && seenSet.size) return; const now = new Set<string>(); for (const sc of scan(wwb)) { sc.rowPos.forEach((_p, id) => now.add('R' + id)); sc.innerPos.forEach((_p, id) => now.add('M' + id)); sc.colPos.forEach((_p, id) => now.add('C' + id)); } seenSet.forEach((k) => { if (!now.has(k)) goneAt.set(k, ws.rev); }); now.forEach((k) => { if (!seenSet.has(k)) bornAt.set(k, ws.rev); }); seenSet = now; seenRev = ws.rev; };
  while (stepNet()); look();
  const intents = new Map<string, Intent>();
  const kindOf = new Map<string, string>();
  let token = 1; let tag = 1000000;
  const submit = (m: M, ops: Op[], intent: Intent): boolean => {
    if (!ops.length) return false;
    const before = (m.session as any).seq;
    try { m.session.submit(ops); } catch (e) { if (e instanceof RangeError) return false; throw e; }
    if ((m.session as any).seq === before) return false;
    const id = `${m.key}:${(m.session as any).seq}`; if ((cfg as any).log) console.log(`${m.name} MAKES ${(m.session as any).seq} ${intent.kind} on sheet uid ${intent.sheet} ${JSON.stringify(intent.cell || intent.f || intent.sname || '')} [sheets ${m.wb.sheets.map((x: any) => x.name).join(',')}]`); pendingSeen.set(m, new Set<string>((m.session as any).pendingSteps.map((x: any) => x.id))); pendingJson.set(m, JSON.stringify((m.session as any).pendingSteps.map((x: any) => [x.id, x.recs.map((q: any) => q.w)]))); intent.concurrentWith = m.session.rev; intents.set(id, intent); kindOf.set(id, intent.kind); return true;
  };
  const str = (v: string) => ({ kind: 'string' as const, value: v, style: 0 });
  const act = (m: M): void => {
    const scs = scan(m.wb); const sc = scs[ri(r, 0, scs.length - 1)]; const s = sc.index;
    if (sc.uid < 0 || sc.problems.length) return; // local view broken: caught at the end
    const dataRows = sc.rowOrder.filter((id) => sc.rowPos.get(id)! > 0);
    const rowSnap = sc.rowOrder.slice(); const colSnap = sc.colOrder.slice();
    const roll = r() * 100;
    const pickRow = () => dataRows.length ? sc.rowPos.get(dataRows[ri(r, 0, dataRows.length - 1)])! : -1;
    const dataCols = sc.colOrder.filter((id) => sc.colPos.get(id) !== sc.innerCol);
    const rkOf = (row: number, col: number): RowKey | undefined => { const inside = col >= 1 && col <= sc.innerCol; const id = inside ? sc.innerAt.get(row) : sc.rowAt.get(row); return id === undefined ? undefined : (inside ? 'M' : 'R') + id; };
    if (roll < 34) { // write
      const row = pickRow(); if (row < 0 || !dataCols.length) return; const cid = dataCols[ri(r, 0, dataCols.length - 1)]; const col = sc.colPos.get(cid)!; const rk = rkOf(row, col); if (!rk) return;
      const v = `V${token++}`; submit(m, setCellInput(m.wb, s, row, col, v), { kind: 'write', sheet: sc.uid, cell: { rk, col: cid, token: v }, rowSnap, colSnap });
    } else if (roll < 42 && cfg.formulas) { // formula, written outside the sorted columns
      const outs = dataCols.filter((id) => sc.colPos.get(id)! > sc.innerCol); const row = pickRow(); if (row < 0 || !outs.length || dataRows.length < 2) return;
      const cid = outs[ri(r, 0, outs.length - 1)]; const col = sc.colPos.get(cid)!; const rk = rkOf(row, col)!;
      const tsc = cfg.sheets && r() < 0.3 ? scs[ri(r, 0, scs.length - 1)] : sc; if (tsc.uid < 0 || tsc.problems.length) return;
      const trows = tsc.rowOrder.filter((id) => tsc.rowPos.get(id)! > 0); if (trows.length < 2 || !tsc.colOrder.length) return;
      const a = ri(r, 0, trows.length - 1); const b = ri(r, a, trows.length - 1); const tc = tsc.colOrder[ri(r, 0, tsc.colOrder.length - 1)]; const tcp = tsc.colPos.get(tc)!;
      const pre = tsc === sc ? '' : `${tsc.name}!`; const t = String(tag++); const form = ['cell', 'range', 'row', 'col'][ri(r, 0, 3)];
      const ra = tsc.rowPos.get(trows[a])!; const rb = tsc.rowPos.get(trows[b])!;
      let x = ''; let frows: number[] = []; let fcols: number[] = [tc];
      if (form === 'cell') { x = `${pre}${colName(tcp)}${ra + 1}`; frows = [trows[a]]; }
      else if (form === 'range') { x = `SUM(${pre}${colName(tcp)}${ra + 1}:${colName(tcp)}${rb + 1})`; frows = trows.slice(a, b + 1); }
      else if (form === 'row') { x = `SUM(${pre}${ra + 1}:${ra + 1})`; frows = [trows[a]]; fcols = []; }
      else { x = `SUM(${pre}${colName(tcp)}:${colName(tcp)})`; frows = []; }
      submit(m, setCellInput(m.wb, s, row, col, `=${x}+${t}`), { kind: 'formula', sheet: sc.uid, cell: { rk, col: cid, token: 'F' + t }, f: { form, tag: t, rows: frows, cols: fcols, target: tsc.uid, snap: tsc.rowOrder.slice() }, rowSnap, colSnap });
    } else if (roll < 47) { // note
      const row = pickRow(); if (row < 0 || !dataCols.length) return; const cid = dataCols[ri(r, 0, dataCols.length - 1)]; const col = sc.colPos.get(cid)!; const rk = rkOf(row, col); if (!rk) return;
      const text = `N${token++}`; submit(m, setComment(m.wb, s, row, col, text, { author: m.name }), { kind: 'note', sheet: sc.uid, note: { rk, col: cid, text }, rowSnap, colSnap });
    } else if (roll < 51) { // height
      const row = pickRow(); if (row < 0) return; const h = 10 + (token++ % 30000) * 0.01; submit(m, setRowHeight(s, row, row, h), { kind: 'height', sheet: sc.uid, height: { row: sc.rowAt.get(row)!, h }, rowSnap, colSnap });
    } else if (roll < 63) { // insert rows
      const at = ri(r, 1, sc.lastRow + 1); const n = ri(r, 1, 3); const rowsNew: number[] = []; const inn: number[] = []; const cells: any[] = [];
      for (let i = 0; i < n; i++) { const id = nextId++; const mm = nextId++; rowsNew.push(id); inn.push(mm); cells.push({ row: at + i, col: 0, cell: str(`R${id}`) }); if (sc.innerCol > 0) cells.push({ row: at + i, col: sc.innerCol, cell: str(`M${mm}`) }); }
      const snap: number[] = []; let put = false; for (const id of rowSnap) { if (!put && sc.rowPos.get(id)! >= at) { snap.push(...rowsNew); put = true; } snap.push(id); } if (!put) snap.push(...rowsNew);
      if (submit(m, [{ type: 'insert', sheet: s, axis: 'row', at, count: n }, { type: 'setCells', sheet: s, cells }] as Op[], { kind: 'insRow', sheet: sc.uid, ins: { rows: rowsNew, inners: sc.innerCol > 0 ? inn : [] }, rowSnap: snap, colSnap })) { rowsNew.forEach((id) => rowHome.set(id, sc.uid)); inn.forEach((id) => innerHome.set(id, sc.uid)); }
    } else if (roll < 74) { // delete rows
      if (dataRows.length < 3) return; const at = ri(r, 1, sc.lastRow); const n = Math.min(ri(r, 1, 3), sc.lastRow - at + 1); const rowsGone: number[] = []; const inn: number[] = [];
      for (let i = at; i < at + n; i++) { if (sc.rowAt.has(i)) rowsGone.push(sc.rowAt.get(i)!); if (sc.innerAt.has(i)) inn.push(sc.innerAt.get(i)!); }
      let ops: Op[]; try { ops = deleteRows(m.wb, s, at, n); } catch (e) { if (e instanceof RangeError) return; throw e; }
      submit(m, ops, { kind: 'delRow', sheet: sc.uid, del: { rows: rowsGone, inners: inn }, rowSnap: rowSnap.filter((id) => rowsGone.indexOf(id) < 0), colSnap });
    } else if (roll < 79 && cfg.cols) { // insert column
      const at = ri(r, cfg.sorts ? 2 : 1, sc.lastCol + 1); const id = nextId++;
      const snap: number[] = []; let put = false; for (const c of colSnap) { if (!put && sc.colPos.get(c)! >= at) { snap.push(id); put = true; } snap.push(c); } if (!put) snap.push(id);
      if (submit(m, [{ type: 'insert', sheet: s, axis: 'col', at, count: 1 }, { type: 'setCells', sheet: s, cells: [{ row: 0, col: at, cell: str(`C${id}`) }] }] as Op[], { kind: 'insCol', sheet: sc.uid, ins: { cols: [id] }, rowSnap, colSnap: snap })) { colHome.set(id, sc.uid); colSide.set(id, at <= sc.innerCol ? 'in' : 'out'); }
    } else if (roll < 83 && cfg.cols) { // delete column
      if (dataCols.length < 3) return; const cid = dataCols[ri(r, 0, dataCols.length - 1)]; const at = sc.colPos.get(cid)!;
      let ops: Op[]; try { ops = deleteColumns(m.wb, s, at, 1); } catch (e) { if (e instanceof RangeError) return; throw e; }
      submit(m, ops, { kind: 'delCol', sheet: sc.uid, del: { cols: [cid] }, rowSnap, colSnap: colSnap.filter((c) => c !== cid) });
    } else if (roll < 90 && cfg.sorts) { // sort
      if (sc.lastRow < 3) return; const top = ri(r, 1, sc.lastRow - 1); const bottom = ri(r, top + 1, sc.lastRow); const wide = r() < 0.6 || sc.innerCol < 1;
      const o: number[] = []; for (let i = 0; i <= bottom - top; i++) o.push(i); for (let i = o.length - 1; i > 0; i--) { const j = ri(r, 0, i); const x = o[i]; o[i] = o[j]; o[j] = x; }
      submit(m, [{ type: 'sortRange', sheet: s, range: { top, left: wide ? 0 : 1, bottom, right: wide ? sc.lastCol : sc.innerCol }, order: o }] as Op[], { kind: wide ? 'sortWide' : 'sortPart', sheet: sc.uid });
    } else if (roll < 93 && cfg.merges) {
      const outs = sc.colOrder.filter((id) => sc.colPos.get(id)! > sc.innerCol); if (outs.length < 2 || dataRows.length < 2) return;
      const top = pickRow(); const bottom = Math.min(top + ri(r, 0, 2), sc.lastRow); const left = sc.colPos.get(outs[0])!; const right = sc.colPos.get(outs[outs.length - 1])!;
      const rr: number[] = []; for (let i = top; i <= bottom; i++) if (sc.rowAt.has(i)) rr.push(sc.rowAt.get(i)!); const cc: number[] = []; for (let i = left; i <= right; i++) if (sc.colAt.has(i)) cc.push(sc.colAt.get(i)!);
      let ops: Op[]; try { ops = mergeRange(m.wb, s, { top, left, bottom, right }); } catch (e) { if (e instanceof RangeError) return; throw e; }
      submit(m, ops, { kind: 'merge', sheet: sc.uid, merge: { rows: rr, cols: cc }, rowSnap, colSnap });
    } else if (cfg.sheets) {
      const k = ri(r, 0, 9);
      if (k < 3) { const name = `Sh${token++}`; const at = ri(r, 0, m.wb.sheets.length); const save = nextId; const uid = nextId++; const rows = sheetRows(uid, 3); const cells: any[] = []; rows.forEach((row, rr) => row.forEach((v, c) => { if (v !== null) cells.push({ row: rr, col: c, cell: str(v) }); }));
        let ops: Op[]; try { ops = insertSheet(m.wb, at, { name }); } catch (e) { if (e instanceof RangeError) { nextId = save; return; } throw e; }
        if (submit(m, ops.concat([{ type: 'setCells', sheet: at, cells }] as Op[]), { kind: 'sheet+', sheet: uid, snew: uid })) sheetName0.set(uid, name); }
      else if (k < 5) { if (m.wb.sheets.length < 2) return; let ops: Op[]; try { ops = deleteSheet(m.wb, s); } catch (e) { if (e instanceof RangeError) return; throw e; } submit(m, ops, { kind: 'sheet-', sheet: sc.uid, sdel: sc.uid }); }
      else if (k < 8 || (cfg as any).noMove) { const name = `Nm${token++}`; let ops: Op[]; try { ops = renameSheet(m.wb, s, name); } catch (e) { if (e instanceof RangeError) return; throw e; } submit(m, ops, { kind: 'rename', sheet: sc.uid, sname: name }); }
      else { let ops: Op[]; try { ops = moveSheet(m.wb, s, ri(r, 0, m.wb.sheets.length - 1)); } catch (e) { if (e instanceof RangeError) return; throw e; } submit(m, ops, { kind: 'move', sheet: sc.uid }); }
    }
  };
  for (let t = 0; t < cfg.ticks; t++) {
    const x = r();
    if (cfg.burst && x < 0.01) { const m = members[ri(r, 0, members.length - 1)]; const n = ri(r, 20, 30); for (let i = 0; i < n; i++) act(m); }
    else if (x < 0.3) act(members[ri(r, 0, members.length - 1)]);
    else if (cfg.cuts && x < 0.33) { const m = members[ri(r, 0, members.length - 1)]; if (m.cut) { m.t.resume(); m.cut = false; } else { m.t.suspend(); m.cut = true; } }
    else stepNet();
  }
  for (const m of members) if (m.cut) { m.t.resume(); m.cut = false; }
  for (let round = 0; round < 50; round++) { while (stepNet()); const due = timers.filter((e) => !e.off); timers = []; if (!due.length) break; for (const e of due) e.cb(); }
  // ---- checks
  const V: string[] = []; const overdrop: string[] = []; const collateral: string[] = [];
  const d0 = dump(members[0].wb);
  for (const m of members) { if (m.session.state !== 'online') V.push(`${m.name} state ${m.session.state} notices ${JSON.stringify(m.notices)}`); if (m.session.pending) V.push(`${m.name} pending ${m.session.pending}`); if (dump(m.wb) !== d0) V.push(`DIVERGED ${m.name}`); }
  const placed = new Set(order);
  const byKind: Result['byKind'] = {}; let made = 0, dropped = 0, quiet = 0, partial = 0;
  for (const m of members) { const st = m.session.stats; made += st.made; dropped += st.dropped; quiet += st.quiet; const said = m.notices.filter((n) => n.code === 'dropped').reduce((a, n) => a + (n.count || 0), 0); partial += m.notices.filter((n) => n.code === 'partial').reduce((a, n) => a + (n.count || 0), 0); if (said !== st.dropped) V.push(`${m.name}: dropped ${st.dropped} but told ${said}`); if (st.made !== st.placed + st.dropped + st.quiet) V.push(`${m.name}: made ${st.made} != placed ${st.placed} + dropped ${st.dropped} + quiet ${st.quiet}`); }
  let notPlaced = 0; intents.forEach((it, id) => { const k = (byKind[it.kind] = byKind[it.kind] || { made: 0, dropped: 0, quiet: 0 }); k.made++; if (!placed.has(id)) { notPlaced++; k.dropped++; } });
  if (notPlaced !== dropped + quiet) V.push(`not placed ${notPlaced} != dropped ${dropped} + quiet ${quiet}`);
  const fin = scan(members[0].wb); const byUid = new Map<number, Scan>(); for (const sc of fin) { for (const p of sc.problems) V.push(`final sheet ${sc.name}: ${p}`); byUid.set(sc.uid, sc); }
  // sheets
  const sheetGone = new Set<number>(); const lastName = new Map<number, string>(sheetName0); const sheetsMade = new Set<number>([u1, u2]);
  let sortsPlaced = new Set<number>(); let partSorts = new Set<number>();
  for (const id of order) { const it = intents.get(id); if (!it) continue; if (it.sdel !== undefined) sheetGone.add(it.sdel); if (it.sname) lastName.set(it.sheet, it.sname); if (it.snew !== undefined) sheetsMade.add(it.snew); if (it.kind === 'sortWide' || it.kind === 'sortPart') sortsPlaced.add(it.sheet); if (it.kind === 'sortPart') partSorts.add(it.sheet); }
  intents.forEach((it) => { if (it.kind === 'sortWide' || it.kind === 'sortPart') sortsPlaced.add(it.sheet); });
  sheetsMade.forEach((uid) => { const here = byUid.has(uid); if (here && sheetGone.has(uid)) V.push(`sheet ${uid} deleted by a placed step but is there`); if (!here && !sheetGone.has(uid)) V.push(`sheet ${uid} is gone and nobody deleted it`); if (here && byUid.get(uid)!.name !== lastName.get(uid)) V.push(`sheet ${uid} is named ${byUid.get(uid)!.name}, last placed name ${lastName.get(uid)}`); });
  byUid.forEach((_sc, uid) => { if (!sheetsMade.has(uid)) V.push(`sheet ${uid} exists but its step was not placed`); });
  // rows / cols: who is gone
  const rowsMade = new Map<number, number>(); const innersMade = new Map<number, number>(); const colsMade = new Map<number, number>(); const insertedAt = new Map<number, string>();
  rowHome.forEach((uid, id) => { if (uid === u1 || uid === u2) rowsMade.set(id, uid); }); innerHome.forEach((uid, id) => { if (uid === u1 || uid === u2) innersMade.set(id, uid); }); colHome.forEach((uid, id) => { if (uid === u1 || uid === u2) colsMade.set(id, uid); });
  // ids made before run started only: the ones of initial sheets created before any insert. Narrow them to those in the file.
  { const init = scan((await (await open(bytes)).workbook())); rowsMade.clear(); innersMade.clear(); colsMade.clear(); for (const sc of init) { sc.rowPos.forEach((_p, id) => { if (id !== sc.uid) rowsMade.set(id, sc.uid); }); sc.innerPos.forEach((_p, id) => innersMade.set(id, sc.uid)); sc.colPos.forEach((_p, id) => colsMade.set(id, sc.uid)); } }
  const rowsDel = new Set<number>(); const innersDel = new Set<number>(); const colsDel = new Set<number>(); const delStep = new Map<number, string>();
  for (const id of order) { const it = intents.get(id); if (!it) continue;
    if (it.ins) { (it.ins.rows || []).forEach((x) => { rowsMade.set(x, it.sheet); insertedAt.set(x, id); }); (it.ins.inners || []).forEach((x) => innersMade.set(x, it.sheet)); (it.ins.cols || []).forEach((x) => colsMade.set(x, it.sheet)); }
    if (it.snew !== undefined) { rowHome.forEach((uid, x) => { if (uid === it.snew) rowsMade.set(x, uid); }); innerHome.forEach((uid, x) => { if (uid === it.snew) innersMade.set(x, uid); }); colHome.forEach((uid, x) => { if (uid === it.snew) colsMade.set(x, uid); }); }
    if (it.del) { (it.del.rows || []).forEach((x) => { rowsDel.add(x); delStep.set(x, id); }); (it.del.inners || []).forEach((x) => innersDel.add(x)); (it.del.cols || []).forEach((x) => colsDel.add(x)); } }
  // only rows that the insert steps of new sheets made after the sheet+ step
  const check = (label: string, madeMap: Map<number, number>, del: Set<number>, pos: (sc: Scan) => Map<number, number>, strict: (uid: number) => boolean) => {
    madeMap.forEach((uid, id) => { const sc = byUid.get(uid); if (!sc) return; const here = pos(sc).has(id); if (here && del.has(id) && strict(uid)) V.push(`${label}${id}: deleted by placed step ${kindOf.get(delStep.get(id) || '')} but survives`); if (!here && !del.has(id)) { (strict(uid) ? collateral : collateral).push(`${label}${id} gone, nobody meant it (inserted by ${insertedAt.get(id) ? 'a step' : 'file'})`); } });
    byUid.forEach((sc) => pos(sc).forEach((_p, id) => { if (id !== sc.uid && !madeMap.has(id)) V.push(`${label}${id} exists but the step that made it was not placed`); }));
  };
  check('R', rowsMade, rowsDel, (sc) => sc.rowPos, () => true);
  check('M', innersMade, innersDel, (sc) => sc.innerPos, (uid) => !partSorts.has(uid));
  check('C', colsMade, colsDel, (sc) => sc.colPos, () => true);
  byUid.forEach((sc) => { if (sc.rowPos.size - 1 !== sc.innerPos.size && sc.innerCol > 0) V.push(`sheet ${sc.name}: ${sc.rowPos.size - 1} row markers, ${sc.innerPos.size} inner markers`); sc.colPos.forEach((p, id) => { if (sc.innerCol > 0 && (p <= sc.innerCol ? 'in' : 'out') !== colSide.get(id)) V.push(`C${id} changed side`); }); });
  // order snapshots
  for (const id of order) { const it = intents.get(id); if (!it) continue; const sc = byUid.get(it.sheet); if (!sc) continue;
    if (it.rowSnap && !sortsPlaced.has(it.sheet)) { const want = it.rowSnap.filter((x) => sc.rowPos.has(x)); const got = sc.rowOrder.filter((x) => it.rowSnap!.indexOf(x) >= 0); if (want.join() !== got.join()) V.push(`row order after ${it.kind}: wanted ${want.join()} got ${got.join()}`); }
    if (it.colSnap) { const want = it.colSnap.filter((x) => sc.colPos.has(x)); const got = sc.colOrder.filter((x) => it.colSnap!.indexOf(x) >= 0); if (want.join() !== got.join()) V.push(`col order after ${it.kind}: wanted ${want.join()} got ${got.join()}`); } }
  // cells and notes
  const last = new Map<string, string>(); const lastNote = new Map<string, string>(); const where = new Map<string, string>(); const lastH = new Map<number, number>();
  for (const id of order) { const it = intents.get(id); if (!it) continue; if (it.cell) { const k = `${it.sheet}|${it.cell.rk}|${it.cell.col}`; last.set(k, it.cell.token); where.set(it.cell.token, k); } if (it.note) { const k = `${it.sheet}|${it.note.rk}|${it.note.col}`; lastNote.set(k, it.note.text); where.set(it.note.text, k); } if (it.height) lastH.set(it.height.row, it.height.h); }
  const keyAt = (sc: Scan, row: number, col: number): string => { const inside = col >= 1 && col <= sc.innerCol; const rid = inside ? sc.innerAt.get(row) : sc.rowAt.get(row); return `${sc.uid}|${rid === undefined ? '?' : (inside ? 'M' : 'R') + rid}|${sc.colAt.get(col)}`; };
  byUid.forEach((sc) => {
    sc.tokens.forEach((at, tok) => { const k = keyAt(sc, at[0], at[1]); if (!where.has(tok)) V.push(`${tok} is in the sheet but its step was not placed`); else if (where.get(tok) !== k) V.push(`WRONG CELL ${tok}: meant ${where.get(tok)}, stands at ${k}`); else if (last.get(k) !== tok) V.push(`${tok} at ${k} but last placed write is ${last.get(k)}`); });
    sc.notes.forEach((at, text) => { const k = keyAt(sc, at[0], at[1]); if (!where.has(text)) V.push(`note ${text} exists but its step was not placed`); else if (where.get(text) !== k) V.push(`WRONG CELL note ${text}: meant ${where.get(text)}, stands at ${k}`); else if (lastNote.get(k) !== text) V.push(`note ${text} at ${k} but last placed is ${lastNote.get(k)}`); });
  });
  const present = (k: string, map: Map<string, [number, number]> | undefined): boolean => { const [uid, rk, col] = k.split('|'); const sc = byUid.get(Number(uid)); if (!sc) return false; const rid = Number(rk.slice(1)); return (rk[0] === 'M' ? sc.innerPos.has(rid) : sc.rowPos.has(rid)) && sc.colPos.has(Number(col)); };
  last.forEach((tok, k) => { if (!present(k, undefined)) return; const sc = byUid.get(Number(k.split('|')[0]))!; if (!sc.tokens.has(tok)) V.push(`LOST WRITE ${tok} meant for ${k}: the cell is there, the value is not`); });
  lastNote.forEach((text, k) => { if (!present(k, undefined)) return; const sc = byUid.get(Number(k.split('|')[0]))!; if (!sc.notes.has(text)) V.push(`LOST NOTE ${text} meant for ${k}`); });
  byUid.forEach((sc) => { if (sortsPlaced.has(sc.uid)) return; sc.rowPos.forEach((p, id) => { const want = lastH.get(id); const got = sc.heights.get(p); if (want !== undefined && got !== want) V.push(`height of R${id}: wanted ${want} got ${got}`); if (want === undefined && got !== undefined) V.push(`height ${got} on R${id} nobody set`); }); });
  // formulas
  if (cfg.formulas) for (const id of order) { const it = intents.get(id); if (!it || !it.f) continue; const home = byUid.get(it.sheet); if (!home) continue; const f = home.formulas.get(it.f.tag); if (!f) continue; // overwritten or row gone: covered above
    if (sortsPlaced.has(it.sheet) || sortsPlaced.has(it.f.target)) continue;
    const tsc = byUid.get(it.f.target); const pre = it.f.target === it.sheet ? '' : tsc ? `${tsc.name}!` : '';
    const rows = tsc ? (it.f.rows || []).filter((x) => tsc.rowPos.has(x)).map((x) => tsc.rowPos.get(x)!) : []; const cols = tsc ? (it.f.cols || []).filter((x) => tsc.colPos.has(x)).map((x) => tsc.colPos.get(x)!) : [];
    const got = f.text.slice(0, f.text.length - it.f.tag.length - 1).replace(/'/g, '');
    if (it.f.form === 'range') {
      // survivors of the range must be inside; rows the author saw outside the range must stay outside
      const m = /^SUM\((?:([A-Za-z0-9]+)!)?([A-Z]+)(\d+):([A-Z]+)(\d+)\)$/.exec(got);
      if (!tsc || !cols.length) { if (got.indexOf('#REF!') < 0) V.push(`FORMULA range tag ${it.f.tag}: wanted #REF! (sheet or column gone) got ${got}`); continue; }
      if (!m) { if (rows.length) V.push(`FORMULA range tag ${it.f.tag}: got ${got}, rows ${rows.join()} survive`); continue; }
      const top = Number(m[3]) - 1; const bottom = Number(m[5]) - 1;
      if ((m[1] || '') !== (it.f.target === it.sheet ? '' : tsc.name) || m[2] !== colName(cols[0]) || m[4] !== colName(cols[0])) V.push(`FORMULA range tag ${it.f.tag}: got ${got}, wanted sheet ${tsc.name} column ${colName(cols[0])}`);
      for (const p of rows) if (p < top || p > bottom) V.push(`FORMULA range tag ${it.f.tag}: got ${got}, row ${p + 1} of the range is outside`);
      const seen = new Set(it.f.snap || []); const meant = new Set(it.f.rows || []);
      for (let p = top; p <= bottom; p++) { const rid = tsc.rowAt.get(p); if (rid !== undefined && seen.has(rid) && !meant.has(rid)) V.push(`FORMULA range tag ${it.f.tag}: got ${got}, holds row ${p + 1} (R${rid}) that the author saw outside the range`); }
      continue;
    }
    let want: string | undefined;
    if (tsc) { if (it.f.form === 'cell' && rows.length && cols.length) want = `${pre}${colName(cols[0])}${rows[0] + 1}`; else if (it.f.form === 'row' && rows.length) want = `SUM(${pre}${rows[0] + 1}:${rows[0] + 1})`; else if (it.f.form === 'col' && cols.length) want = `SUM(${pre}${colName(cols[0])}:${colName(cols[0])})`; }
    if (want === undefined ? got.indexOf('#REF!') < 0 : got !== want) V.push(`FORMULA ${it.f.form} tag ${it.f.tag}: wanted ${want === undefined ? '#REF!' : want} got ${got}`); }
  // merges
  if (cfg.merges) byUid.forEach((sc) => { if (sortsPlaced.has(sc.uid)) return; const wants: string[] = []; for (const id of order) { const it = intents.get(id); if (!it || !it.merge || it.sheet !== sc.uid) continue; const rr = it.merge.rows.filter((x) => sc.rowPos.has(x)).map((x) => sc.rowPos.get(x)!); const cc = it.merge.cols.filter((x) => sc.colPos.has(x)).map((x) => sc.colPos.get(x)!); if (!rr.length || !cc.length) continue; const rect = [Math.min(...rr), Math.min(...cc), Math.max(...rr), Math.max(...cc)]; if (rect[0] === rect[2] && rect[1] === rect[3]) continue; wants.push(rect.join()); }
    for (const m of sc.merges) if (wants.indexOf(m.join()) < 0) V.push(`MERGE ${m.join()} on ${sc.name} matches no placed merge (${wants.join(' ; ')})`);
  });
  // over-dropping: a step that was dropped though what it was about was there when it was dropped
  const revOf = new Map<string, number>(); order.forEach((id, i) => revOf.set(id, i + 1));
  const rowDelAt = new Map<number, number>(); const innerDelAt = new Map<number, number>(); const colDelAt = new Map<number, number>(); const sheetDelAt = new Map<number, number>(); const madeBy = new Map<string, string>();
  intents.forEach((it, id) => { if (it.ins) { (it.ins.rows || []).forEach((x) => madeBy.set('R' + x, id)); (it.ins.inners || []).forEach((x) => madeBy.set('M' + x, id)); (it.ins.cols || []).forEach((x) => madeBy.set('C' + x, id)); } if (it.snew !== undefined) madeBy.set('S' + it.snew, id); });
  order.forEach((id, i) => { const it = intents.get(id); if (!it) return; if (it.del) { const first = (map: Map<number, number>, x: number) => { if (!map.has(x)) map.set(x, i + 1); }; (it.del.rows || []).forEach((x) => first(rowDelAt, x)); (it.del.inners || []).forEach((x) => first(innerDelAt, x)); (it.del.cols || []).forEach((x) => first(colDelAt, x)); } if (it.sdel !== undefined && !sheetDelAt.has(it.sdel)) sheetDelAt.set(it.sdel, i + 1); });
  const collateralIds = new Set<string>(collateral.map((c) => c.split(' ')[0]));
  const classes: string[] = [];
  collateralIds.forEach((k) => { const d = goneAt.get(k); const b = bornAt.get(k) || 0; if (d === undefined) { classes.push('never seen going'); return; } const killer = intents.get(order[d - 1]); if (!killer) { classes.push('killer unknown'); return; }
    const c = killer.concurrentWith as number; classes.push(`${k[0]} killed by ${killer.kind} ${c < b ? 'made before the row was placed (concurrent)' : 'MADE AFTER THE ROW WAS PLACED'}`); if (c >= b && cfg.trace) console.log(`NONCONCURRENT seed ${cfg.seed} ${k} born ${b} gone ${d} killer ${order[d - 1]} made at ${c} intent ${JSON.stringify(killer.del || killer)}`); });
  (cfg as any).classes = classes;
  intents.forEach((it, id) => { if (placed.has(id)) return; const at = dropAt.get(id); if (at === undefined) { V.push(`step ${it.kind} neither placed nor seen dropped`); return; }
    const alive = (label: string, x: number, delAt: Map<number, number>): boolean => { const maker = madeBy.get(label + x); if (maker !== undefined && !placed.has(maker)) return false; if (collateralIds.has(label + x)) return false; const d = delAt.get(x); return d === undefined || d > at; };
    const sheetAlive = alive('S', it.sheet, sheetDelAt); let why = '';
    const rkAlive = (rk: string) => (rk[0] === 'M' ? (partSorts.has(it.sheet) ? false : alive('M', Number(rk.slice(1)), innerDelAt)) : alive('R', Number(rk.slice(1)), rowDelAt));
    if (it.cell || it.note) { const c = (it.cell || it.note)!; if (sheetAlive && rkAlive(c.rk) && alive('C', c.col, colDelAt)) why = 'cell was there'; }
    else if (it.height) { if (sheetAlive && alive('R', it.height.row, rowDelAt)) why = 'row was there'; }
    else if (it.kind === 'insRow' || it.kind === 'insCol' || it.kind === 'rename' || it.kind === 'move') { if (sheetAlive) why = 'sheet was there'; }
    else if (it.kind === 'sheet+') why = 'always possible';
    else if (it.kind === 'delRow') { if (sheetAlive && it.del!.rows!.some((x) => alive('R', x, rowDelAt))) why = 'rows were there'; }
    else if (it.kind === 'delCol') { if (sheetAlive && it.del!.cols!.some((x) => alive('C', x, colDelAt))) why = 'col was there'; }
    else if (it.kind === 'sheet-') { if (sheetAlive) why = 'sheet was there'; }
    else if (it.kind.slice(0, 4) === 'sort') { if (sheetAlive) why = 'sort'; }
    if (why) { overdrop.push(`${it.kind}: ${why}`); if (cfg.trace && it.kind !== 'sortWide' && it.kind !== 'sortPart') { const info = dropInfo.get(id); console.log(`OVERDROP seed ${cfg.seed} ${info.who} ${it.kind} id ${id} at rev ${at} intent ${JSON.stringify(it.cell || it.note || it.height || it.del || it.ins || {})}\n   remote: ${String(info.last).slice(0, 400)}\n   pending before: ${String(info.before).slice(0, 1500)}`); } } });
  return { classes, violations: V, made, placed: order.length, dropped, quiet, overdrop, collateral, byKind, partial };
}
