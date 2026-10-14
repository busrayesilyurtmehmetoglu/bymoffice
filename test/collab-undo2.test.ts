// Undo with steps of others in between, checked against a replay by identity: rows carry markers, and what the
// sheet has to hold after every step, undoing included, follows from the order of the relay. Written by a
// reviewer; an undo is asked for only when nothing of its participant is on its way (it would wait otherwise).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx } from '../src/formats/xlsx/build';
import { open, type Workbook } from '../src/index';
import { MemoryServer } from '../src/collab/server';
import { CollabSession } from '../src/collab/session';
import { contentVersion } from '../src/collab/protocol';
import { setCellInput } from '../src/ops/commands';
import { deleteRows } from '../src/ops/filter';
import type { Op } from '../src/ops/ops';
import { bookOf } from './chart-edit-fixtures';
import { rng } from './collab-semantic2-model';
import { PROTOCOL_VERSION } from '../src/collab/room';
const ri = (r: () => number, a: number, b: number) => a + Math.floor(r() * (b - a + 1));
const look = (wb: Workbook) => { const rows: { id: string; pos: number; vals: string[] }[] = []; for (const [p, row] of wb.sheets[0].rows) { const c = row.cells[0]; if (c && /^R\d+$/.test(String(c.value))) rows.push({ id: String(c.value), pos: p, vals: [1, 2, 3].map((k) => (row.cells[k] ? String(row.cells[k].value) : '')) }); else if (row.cells.some((x) => x && x.kind !== 'blank')) rows.push({ id: '?', pos: p, vals: [] }); } return rows.sort((a, b) => a.pos - b.pos); };
type It = { k: 'w'; row: string; col: number; v: string } | { k: 'ins'; rows: string[]; snap: string[] } | { k: 'del'; rows: string[]; snap: string[]; content: Map<string, string[]> } | { k: 'undo'; of: string; snap: string[] };
export async function run(seed: number, withDel: boolean, trace = false) {
  const r = rng(seed); let next = 1; const init: (string | null)[][] = []; for (let i = 0; i < 8; i++) init.push([`R${next++}`, null, null, null]);
  const bytes = await buildXlsx(bookOf(init)); const lanes = new Map<string, (() => void)[]>(); const server = new MemoryServer({ deliver: (fn, lane) => { let q = lanes.get(lane); if (!q) lanes.set(lane, (q = [])); q.push(fn); } });
  const log: string[] = []; const stepNet = () => { const busy = [...lanes.entries()].filter((q) => q[1].length); if (!busy.length) return false; const pick = busy[Math.floor(r() * busy.length)]; if (trace) log.push(`   deliver ${pick[0]}`); pick[1].shift()!(); return true; };
  const order: string[] = []; const obs = server.connect(); obs.onMessage((t) => { const q = JSON.parse(t); if (q.t === 'op') order.push(q.id); }); obs.onState!((s) => { if (s === 'open') obs.send(JSON.stringify({ t: 'join', pv: PROTOCOL_VERSION, key: 'obs', v: [contentVersion(bytes)] })); }); while (stepNet());
  const ms: any[] = []; for (let i = 0; i < 2 + (seed % 2); i++) { const wb = await (await open(bytes)).workbook(); const m: any = { name: 'm' + i, wb, notices: [] }; m.session = new CollabSession({ workbook: () => m.wb, changed: () => undefined, notice: (n) => m.notices.push(n) }, { transport: server.connect(), versions: [contentVersion(bytes)], timer: () => () => undefined }); m.key = m.session.tag; ms.push(m); }
  while (stepNet());
  const its = new Map<string, It>(); let tok = 1; log.length = 0;
  const sub = (m: any, ops: Op[], it: It): boolean => { const before = m.session.seq; try { m.session.submit(ops, `${m.key}:${before + 1}`); } catch (e) { if (e instanceof RangeError) return false; throw e; } if (m.session.seq === before) return false; its.set(`${m.key}:${m.session.seq}`, it); if (trace) log.push(`${m.name} ${JSON.stringify(it, (k, v) => (v instanceof Map ? [...v] : v)).slice(0, 200)}`); return true; };
  for (let t = 0; t < Number(process.env.TICKS || 220); t++) { const x = r(); const m = ms[ri(r, 0, ms.length - 1)]; const rows = look(m.wb); const real = rows.filter((q) => q.id !== '?');
    if (x < 0.14 && !process.env.NOWRITE) { if (!real.length) continue; const row = real[ri(r, 0, real.length - 1)]; const col = ri(r, 1, 3); const v = `V${tok++}`; sub(m, setCellInput(m.wb, 0, row.pos, col, v), { k: 'w', row: row.id, col, v }); }
    else if (x < 0.20) { const at = ri(r, 0, rows.length ? rows[rows.length - 1].pos + 1 : 0); const n = ri(r, 1, 2); const ids: string[] = []; const cells: any[] = []; for (let i = 0; i < n; i++) { ids.push(`R${next++}`); cells.push({ row: at + i, col: 0, cell: { kind: 'string', value: ids[i], style: 0 } }); } const snap: string[] = []; let put = false; for (const q of real) { if (!put && q.pos >= at) { snap.push(...ids); put = true; } snap.push(q.id); } if (!put) snap.push(...ids); sub(m, [{ type: 'insert', sheet: 0, axis: 'row', at, count: n }, { type: 'setCells', sheet: 0, cells }] as Op[], { k: 'ins', rows: ids, snap }); }
    else if (x < 0.25 && withDel) { if (real.length < 3) continue; const a = ri(r, 0, real.length - 1); const n = ri(r, 1, 2); const gone = real.filter((q) => q.pos >= real[a].pos && q.pos < real[a].pos + n); let ops: Op[]; try { ops = deleteRows(m.wb, 0, real[a].pos, n); } catch (e) { if (e instanceof RangeError) continue; throw e; } sub(m, ops, { k: 'del', rows: gone.map((q) => q.id), snap: real.filter((q) => gone.indexOf(q) < 0).map((q) => q.id), content: new Map(gone.map((q) => [q.id, q.vals])) }); }
    else if (x < 0.34 && m.session.pending === 0) { const before = m.session.seq; const pendingBefore = m.session.pending; let went; try { went = m.session.undo(); } catch (e: any) { throw new Error('undo threw ' + e.message); } if (went && m.session.seq !== before) { its.set(`${m.key}:${m.session.seq}`, { k: 'undo', of: String(went.meta), snap: look(m.wb).filter((q) => q.id !== '?').map((q) => q.id) }); if (trace) log.push(`${m.name} UNDO of ${went.meta} [${(its.get(String(went.meta)) || { k: '?' }).k}]`); } else if (went && m.session.state !== 'online') { /* applied but not queued */ } }
    else for (let k = ri(r, 1, 8); k > 0; k--) stepNet(); }
  while (stepNet());
  const V: string[] = []; const a = JSON.stringify(look(ms[0].wb)); for (const m of ms) { if (JSON.stringify(look(m.wb)) !== a) V.push('DIVERGED'); if (m.session.pending) V.push('pending'); }
  // replay
  const present = new Set<string>(); for (let i = 1; i <= 8; i++) present.add(`R${i}`); const cell = new Map<string, { v: string; by: string }[]>(); const placed = new Set(order); const lossy = ms.some((m) => m.notices.some((n: any) => n.code === 'undo-lossy'));
  const top = (k: string) => { const s = cell.get(k); return s && s.length ? s[s.length - 1] : undefined; }; const delContent = new Map<string, string>();
  for (const id of order) { const it = its.get(id); if (!it) continue;
    if (it.k === 'w') { const k = `${it.row}|${it.col}`; if (!cell.has(k)) cell.set(k, []); cell.get(k)!.push({ v: it.v, by: id }); }
    else if (it.k === 'ins') it.rows.forEach((x) => present.add(x)); else if (it.k === 'del') it.rows.forEach((x) => present.delete(x));
    else { const of = its.get(it.of); if (!of) continue; if (of.k === 'w') { const k = `${of.row}|${of.col}`; const s = cell.get(k) || []; if (s.length && s[s.length - 1].by === it.of) s.pop(); /* else: someone wrote since, it stays */ } else if (of.k === 'ins') of.rows.forEach((x) => present.delete(x)); else if (of.k === 'del') of.rows.forEach((x) => present.add(x)); } }
  const fin = look(ms[0].wb); const pos = new Map(fin.map((q) => [q.id, q.pos])); const cls = (s: string) => { V.push(s); };
  if (fin.some((q) => q.id === '?')) cls('a row with content and no marker');
  if (new Set(fin.map((q) => q.id)).size !== fin.length) cls('a marker twice');
  present.forEach((x) => { if (!pos.has(x)) cls(`row ${x} should be there (replay) and is not`); }); fin.forEach((q) => { if (q.id !== '?' && !present.has(q.id)) cls(`row ${q.id} is there and should not be (replay)`); });
  for (const id of order) { const it = its.get(id); if (!it || it.k === 'w') continue; const want = it.snap.filter((x) => pos.has(x)); const got = fin.map((q) => q.id).filter((x) => it.snap.indexOf(x) >= 0); if (want.join() !== got.join()) cls(`order after ${it.k}${it.k === 'undo' ? ' of ' + (its.get(it.of) || { k: '?' }).k : ''}: wanted ${want.join()} got ${got.join()}`); }
  for (const q of fin) for (let c = 1; c <= 3; c++) { const t = top(`${q.id}|${c}`); const want = t ? t.v : ''; if (q.vals[c - 1] !== want) cls(`cell ${q.id}|${c}: wanted ${want || 'empty'} got ${q.vals[c - 1] || 'empty'}`); }
  return { V, log, notices: ms.map((m) => m.notices.map((n: any) => n.code)), stats: ms.map((m) => m.session.stats) };
}

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));

for (const withDeletes of [false, true]) {
  test(`undo beside the steps of others, ${withDeletes ? 'with' : 'without'} rows that are deleted (${SEEDS} runs)`, async () => {
    const bad: string[] = [];
    for (let seed = 1; seed <= SEEDS && bad.length < 3; seed++) {
      const result = await run(seed, withDeletes);
      if (result.V.length > 0) bad.push(`seed ${seed}: ${result.V.slice(0, 2).join(' | ').slice(0, 300)}`);
    }
    assert.deepEqual(bad, [], 'runs that ended with the wrong rows or cells');
  });
}

test('undo: rows that were put in at one place keep their order after an undo (once seed 155)', async () => {
  const result = await run(155, true);
  assert.deepEqual(result.V, [], 'seed 155');
});
