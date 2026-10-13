// What should have happened, worked out without the code that rebases: rows and columns of a
// sheet carry markers (the text in their first cell), every change is noted by the markers its
// author saw when making it, and what the workbook has to hold in the end follows from the
// order of the relay's log alone. The random runs of collab-random.test.ts only show that
// everybody ends up the same; these show that it is the right workbook.
//
// Nothing here reads src/collab/rebase.ts or wire.ts: the model stands on the session, the
// relay of the page and the commands.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx } from '../src/formats/xlsx/build';
import { cellName, open, type OfficeFile, type Workbook } from '../src/index';
import { contentVersion } from '../src/collab/protocol';
import { MemoryServer, type MemoryTransport } from '../src/collab/server';
import { CollabSession, type CollabNotice } from '../src/collab/session';
import { setCellInput } from '../src/ops/commands';
import { deleteColumns, deleteRows } from '../src/ops/filter';
import { renameSheet } from '../src/ops/sheets';
import { bookOf } from './chart-edit-fixtures';

function rng(seed: number): () => number {
  let a = (seed * 2654435761) >>> 0;
  return () => {
    a |= 0;
    a = (a + 0x9e3779b9) | 0;
    let t = a ^ (a >>> 16);
    t = Math.imul(t, 0x21f0aaad);
    t ^= t >>> 15;
    t = Math.imul(t, 0x735a2d97);
    return ((t ^ (t >>> 15)) >>> 0) / 4294967296;
  };
}
const ri = (r: () => number, lo: number, hi: number) => lo + Math.floor(r() * (hi - lo + 1));
const rp = <T>(r: () => number, l: readonly T[]): T => l[Math.floor(r() * l.length)];

/** The relay of the page with its deliveries held back, by connection and direction. */
class Wire {
  server: MemoryServer;
  lanes = new Map<string, (() => void)[]>();
  constructor() {
    this.server = new MemoryServer({
      deliver: (run, lane) => {
        let q = this.lanes.get(lane);
        if (!q) this.lanes.set(lane, (q = []));
        q.push(run);
      },
    });
  }
  busy(): string[] {
    return Array.from(this.lanes.entries()).filter(([, q]) => q.length > 0).map(([k]) => k);
  }
  step(r: () => number, skip?: (lane: string) => boolean): boolean {
    let b = this.busy();
    if (skip) b = b.filter((l) => !skip(l));
    if (b.length === 0) return false;
    this.lanes.get(rp(r, b))!.shift()!();
    return true;
  }
  drain(r: () => number): void {
    let n = 0;
    while (this.step(r)) if (++n > 5e6) throw new Error('never drained');
  }
  log(): { id: string }[] {
    return ((this.server as unknown as { roomsByKey: Map<string, { log: { id: string }[] }> }).roomsByKey.get('') || { log: [] }).log;
  }
}

interface Peer { name: string; file: OfficeFile; wb: Workbook; s: CollabSession; t: MemoryTransport; notices: CollabNotice[]; lane: number; sent: string[] }

async function mk(net: Wire, name: string, bytes: Uint8Array): Promise<Peer> {
  const file = await open(bytes);
  const wb = await file.workbook();
  const t = net.server.connect();
  const lane = (net.server as unknown as { lanes: number }).lanes;
  const p: Peer = { name, file, wb, s: undefined as unknown as CollabSession, t, notices: [], lane, sent: [] };
  const wrapped = { ...t, send: (text: string) => { p.sent.push(text); t.send(text); } };
  p.s = new CollabSession({ workbook: () => p.wb, changed: () => undefined, notice: (n) => p.notices.push(n) }, { transport: wrapped, versions: [contentVersion(bytes)], user: { name }, timer: () => () => undefined });
  return p;
}

const ROWS = 12, COLS = 6;
async function file(): Promise<Uint8Array> {
  const rows: (number | string | null)[][] = [];
  for (let r = 0; r < ROWS; r++) { const row: (number | string | null)[] = []; for (let c = 0; c < COLS; c++) row.push(r === 0 ? (c === 0 ? 'H' : `C${c}`) : c === 0 ? `R${r}` : r * 100 + c); rows.push(row); }
  const other: (number | string | null)[][] = [['o']];
  return buildXlsx(bookOf(rows, [{ name: 'Other', rows: other }]));
}
const val = (wb: Workbook, s: number, r: number, c: number): any => { const row = wb.sheets[s].rows.get(r); const cell = row ? row.cells[c] : undefined; return cell ? (cell.formula ? '=' + cell.formula.text : cell.value) : null; };
const a1 = (r: number, c: number) => cellName(r, c);
function markers(wb: Workbook) {
  const sh = wb.sheets[0]; let maxR = 0; for (const k of sh.rows.keys()) if (k > maxR) maxR = k;
  const rowM: (string | null)[] = []; for (let r = 0; r <= maxR; r++) { const v = val(wb, 0, r, 0); rowM.push(typeof v === 'string' && /^R\d+$/.test(v) ? v : null); }
  let maxC = 0; const head = sh.rows.get(0); if (head) head.cells.forEach((_c, i) => { if (i > maxC) maxC = i; });
  const colM: (string | null)[] = []; for (let c = 0; c <= maxC; c++) { const v = val(wb, 0, 0, c); colM.push(typeof v === 'string' && /^C\d+$/.test(v) ? v : null); }
  return { rowM, colM };
}
type Intent =
  | { k: 'w'; rm: string; cm: string; v: string }
  | { k: 'f'; rm: string; cm: string; trm: string; tcm: string }
  | { k: 'of'; r: number; c: number; trm: string; tcm: string }
  | { k: 'ins'; axis: 'row' | 'col'; above: string | null; below: string | null; n: number }
  | { k: 'del'; axis: 'row' | 'col'; ms: string[] }
  | { k: 'name'; name: string }
  | { k: 'sort' };

export async function run(seed: number, mode: 'rc' | 'sort', verbose = false): Promise<string | undefined> {
  const r = rng(seed);
  const bytes = await file();
  const net = new Wire();
  const ps: Peer[] = [];
  const n = ri(r, 2, 5);
  for (let i = 0; i < n; i++) ps.push(await mk(net, `p${i}`, bytes));
  net.drain(r);
  const intents = new Map<string, Intent>(); // `${peer}:${seq}`
  const order: string[] = [];
  let counter = 0;
  const frozen = new Set<number>();
  const skip = (lane: string) => frozen.has(Number(lane.split(':')[0])) && lane.endsWith(':down');
  const submit = (p: Peer, ops: any[], intent: Intent): void => {
    if (ops.length === 0) return;
    const before = p.s.stats.made;
    try { p.s.submit(ops); } catch (e) { if (e instanceof RangeError) return; throw e; }
    if (p.s.stats.made !== before + 1) return;
    intents.set(`${p.name}:${p.s.stats.made}`, intent); order.push(`${p.name}:${p.s.stats.made} ${JSON.stringify(intent)}`);
  };
  const act = (p: Peer): void => {
    const wb = p.wb; const { rowM, colM } = markers(wb);
    const rIdx = rowM.map((m, i) => (m ? i : -1)).filter((i) => i >= 0); const cIdx = colM.map((m, i) => (m ? i : -1)).filter((i) => i >= 0);
    const roll = r();
    if (roll < 0.34) {
      if (!rIdx.length || !cIdx.length) return; const rr = rp(r, rIdx), cc = rp(r, cIdx); const v = `v${++counter}`;
      submit(p, setCellInput(wb, 0, rr, cc, v), { k: 'w', rm: rowM[rr]!, cm: colM[cc]!, v });
    } else if (roll < 0.46 && mode === 'rc') {
      if (rIdx.length < 2 || !cIdx.length) return; const rr = rp(r, rIdx), cc = rp(r, cIdx); let tr = rp(r, rIdx), tc = rp(r, cIdx); if (tr === rr && tc === cc) return;
      submit(p, setCellInput(wb, 0, rr, cc, `=${a1(tr, tc)}`), { k: 'f', rm: rowM[rr]!, cm: colM[cc]!, trm: rowM[tr]!, tcm: colM[tc]! });
    } else if (roll < 0.54 && mode === 'rc') {
      if (!rIdx.length || !cIdx.length) return; const tr = rp(r, rIdx), tc = rp(r, cIdx); const or = ri(r, 3, 12), oc = ri(r, 0, 3);
      const nm = wb.sheets[0].name; const q = /^[A-Za-z_][A-Za-z0-9_]*$/.test(nm) ? nm : `'${nm.replace(/'/g, "''")}'`;
      submit(p, setCellInput(wb, 1, or, oc, `=${q}!${a1(tr, tc)}`), { k: 'of', r: or, c: oc, trm: rowM[tr]!, tcm: colM[tc]! });
    } else if (roll < 0.70) {
      const axis = mode === 'sort' || r() < 0.6 ? 'row' : 'col'; const M = axis === 'row' ? rowM : colM; const at = ri(r, 1, M.length); const cnt = ri(r, 1, 3);
      let above: string | null = null, below: string | null = null;
      for (let i = at - 1; i >= 1; i--) if (M[i]) { above = M[i]; break; }
      for (let i = at; i < M.length; i++) if (M[i]) { below = M[i]; break; }
      submit(p, [{ type: 'insert', sheet: 0, axis, at, count: cnt }], { k: 'ins', axis, above, below, n: cnt });
    } else if (roll < 0.88) {
      const axis = mode === 'sort' || r() < 0.6 ? 'row' : 'col'; const M = axis === 'row' ? rowM : colM; if (M.length < 3) return;
      const at = ri(r, 1, M.length - 1); const cnt = Math.min(ri(r, 1, 3), M.length - at);
      const ms = M.slice(at, at + cnt); if (ms.some((m) => !m)) return; // only bands made of marked rows
      submit(p, axis === 'row' ? deleteRows(wb, 0, at, cnt) : deleteColumns(wb, 0, at, cnt), { k: 'del', axis, ms: ms as string[] });
    } else if (roll < 0.94 && mode === 'rc') {
      const name = rp(r, ['Data', 'Veri', 'Yeni Ad', "A'b", 'X1']);
      submit(p, renameSheet(wb, 0, name), { k: 'name', name });
    } else if (mode === 'sort') {
      if (rowM.length < 4) return; const top = ri(r, 1, rowM.length - 2); const bottom = ri(r, top + 1, rowM.length - 1); const len = bottom - top + 1;
      const perm = Array.from({ length: len }, (_x, i) => i); for (let i = len - 1; i > 0; i--) { const j = ri(r, 0, i); [perm[i], perm[j]] = [perm[j], perm[i]]; }
      submit(p, [{ type: 'sortRange', sheet: 0, range: { top, left: 0, bottom, right: colM.length - 1 }, order: perm }], { k: 'sort' });
    }
  };
  const steps = ri(r, 10, 60);
  for (let i = 0; i < steps; i++) {
    const roll = r();
    if (roll < 0.5) act(rp(r, ps));
    else if (roll < 0.56) { for (const p of ps) act(p); }
    else if (roll < 0.60) { const p = rp(r, ps); for (let k = ri(r, 3, 20); k > 0; k--) act(p); }
    else if (roll < 0.66) { const p = rp(r, ps); if (frozen.has(p.lane)) frozen.delete(p.lane); else frozen.add(p.lane); }
    else if (roll < 0.70) rp(r, ps).t.suspend();
    else if (roll < 0.75) rp(r, ps).t.resume();
    else for (let k = ri(r, 1, 12); k > 0; k--) net.step(r, skip);
  }
  frozen.clear(); for (const p of ps) p.t.resume(); net.drain(r);

  // ---- expected, from the log alone
  const keyOf = new Map<string, string>(); // session key -> peer name
  for (const p of ps) for (const t of p.sent) { const m = JSON.parse(t); if (m.t === 'op') { keyOf.set(m.id.split(':')[0], p.name); break; } }
  const placed: { who: string; it: Intent }[] = []; const placedIds = new Set<string>();
  for (const e of net.log()) { const [key, seq] = e.id.split(':'); const who = `${keyOf.get(key)}:${seq}`; const it = intents.get(who); if (!it) return `log entry ${e.id} has no intent (${who})`; placed.push({ who, it }); placedIds.add(who); }
  const goneR = new Set<string>(), goneC = new Set<string>();
  for (const [, it] of intents) if (it.k === 'del') for (const m of it.ms) (it.axis === 'row' ? goneR : goneC).add(m);
  const wb = ps[0].wb; const { rowM, colM } = markers(wb);
  const fail = (s: string) => `${s}\n   steps: ${verbose ? order.join('\n     ') : order.length}\n   placed: ${placed.map((x) => x.who).join(' ')}`;
  // survivors
  for (let i = 1; i < ROWS; i++) { const m = `R${i}`; const here = rowM.indexOf(m) >= 0; if (here === goneR.has(m)) return fail(`row ${m}: ${here ? 'survived a deletion' : 'is gone though nobody deleted it'}`); }
  for (let i = 1; i < COLS; i++) { const m = `C${i}`; const here = colM.indexOf(m) >= 0; if (here === goneC.has(m)) return fail(`col ${m}: ${here ? 'survived a deletion' : 'is gone though nobody deleted it'}`); }
  // order of markers
  const seqR = rowM.filter(Boolean).map((m) => Number(m!.slice(1))); if (mode === 'rc') for (let i = 1; i < seqR.length; i++) if (seqR[i] < seqR[i - 1]) return fail(`rows out of order ${seqR}`);
  // unplaced steps: must be justified
  for (const [who, it] of intents) {
    if (placedIds.has(who)) continue;
    if (it.k === 'w' || it.k === 'f') { if (!goneR.has(it.rm) && !goneC.has(it.cm)) return fail(`LOST WRITE ${who} ${JSON.stringify(it)}: never placed though its row and column are there`); }
    else if (it.k === 'del') { /* fine: all its rows are gone (checked above) */ }
    else if (it.k === 'ins') return fail(`insertion ${who} ${JSON.stringify(it)} was never placed`);
    else if (it.k === 'of') { /* its target may be gone -> still should be placed as #REF!? report softly */ if (!goneR.has(it.trm) && !goneC.has(it.tcm)) return fail(`LOST formula on Other ${who} ${JSON.stringify(it)}`); }
    else if (it.k === 'name') return fail(`rename ${who} never placed`);
  }
  // cell contents
  const exp = new Map<string, Intent>(); const expO = new Map<string, Intent>(); let name = 'Sayfa1';
  let firstName = true;
  for (const { it } of placed) { if (it.k === 'w' || it.k === 'f') exp.set(`${it.rm}|${it.cm}`, it); else if (it.k === 'of') expO.set(`${it.r}|${it.c}`, it); else if (it.k === 'name') name = it.name; }
  const finalName = wb.sheets[0].name;
  if (placed.some((x) => x.it.k === 'name') && finalName !== name) return fail(`sheet name ${finalName} != last placed rename ${name}`);
  const addr = (trm: string, tcm: string): string | null => { const rr = rowM.indexOf(trm), cc = colM.indexOf(tcm); return rr < 0 || cc < 0 ? null : a1(rr, cc); };
  for (let rr = 1; rr < rowM.length; rr++) for (let cc = 1; cc < colM.length; cc++) {
    const got = val(wb, 0, rr, cc); const rm = rowM[rr], cm = colM[cc];
    if (!rm || !cm) { if (got !== null && got !== undefined && got !== '') return fail(`content ${JSON.stringify(got)} in an inserted row/column at ${a1(rr, cc)}`); continue; }
    const it = exp.get(`${rm}|${cm}`);
    if (!it) { const init = Number(rm.slice(1)) * 100 + Number(cm.slice(1)); if (got !== init) return fail(`${rm}/${cm} (${a1(rr, cc)}): nobody wrote, has ${JSON.stringify(got)} not ${init}`); }
    else if (it.k === 'w') { if (got !== it.v) return fail(`${rm}/${cm} (${a1(rr, cc)}): last placed write is ${it.v}, has ${JSON.stringify(got)}`); }
    else if (it.k === 'f') { const t = addr(it.trm, it.tcm); const want = t ? `=${t}` : '#REF!'; if (t ? got !== want : !String(got).includes('#REF!')) return fail(`${rm}/${cm} (${a1(rr, cc)}): formula should point at ${it.trm}/${it.tcm} = ${want}, has ${JSON.stringify(got)}`); }
  }
  for (const [k, it] of expO) { if (it.k !== 'of') continue; const got = String(val(wb, 1, it.r, it.c)); const t = addr(it.trm, it.tcm);
    if (!t) { if (!got.includes('#REF!')) return fail(`Other ${k}: target ${it.trm}/${it.tcm} is gone, has ${got}`); continue; }
    const m = /^=(.*)!([A-Z]+\d+)$/.exec(got); if (!m) return fail(`Other ${k}: unreadable ${got}`);
    const nm = m[1].replace(/^'|'$/g, '').replace(/''/g, "'"); if (nm !== finalName || m[2] !== t) return fail(`Other ${k}: should be ${finalName}!${t} (${it.trm}/${it.tcm}), has ${got}`); }
  // inserted rows between neighbours that both survive and were neighbours
  if (mode === 'rc') for (const axis of ['row', 'col'] as const) {
    const M = axis === 'row' ? rowM : colM; const gone = axis === 'row' ? goneR : goneC; const pre = axis === 'row' ? 'R' : 'C'; const N = axis === 'row' ? ROWS : COLS;
    for (let i = 1; i < N - 1; i++) { const a = `${pre}${i}`, b = `${pre}${i + 1}`; if (gone.has(a) || gone.has(b)) continue;
      let want = 0; for (const { it } of placed) if (it.k === 'ins' && it.axis === axis && it.above === a && it.below === b) want += it.n;
      const have = M.indexOf(b) - M.indexOf(a) - 1; if (have !== want) return fail(`${axis}s between ${a} and ${b}: ${have} inserted, placed insertions say ${want}`); }
  }
  // told
  for (const p of ps) for (const nn of p.notices) if (nn.code !== 'dropped' && nn.code !== 'partial') return fail(`${p.name} notice ${JSON.stringify(nn)}`);
  return undefined;
}

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));

for (const mode of ['rc', 'sort'] as const) {
  test(`what ends up in the workbook is what the order of the relay says (${mode === 'rc' ? 'rows, columns, formulas, names' : 'sorting'}; ${SEEDS} seeds)`, async () => {
    const bad: string[] = [];
    for (let seed = 1; seed <= SEEDS; seed++) {
      const result = await run(seed, mode, true);
      if (result) bad.push(`seed ${seed}: ${result}`);
      if (bad.length >= 3) break;
    }
    assert.deepEqual(bad, [], `${bad.length} of the runs ended with the wrong workbook`);
  });
}
