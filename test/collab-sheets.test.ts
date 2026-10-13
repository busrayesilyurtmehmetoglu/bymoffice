// Every meeting of one or two waiting steps on the list of sheets (insert, delete, move, rename) with one
// step of someone else, written by a reviewer as the whole space of small cases: after each, a token that was
// written on every sheet stands on the sheet it was meant for, or its step was dropped and that was said.
// Never on another sheet, never lost unsaid, and both participants have the same sheets.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx } from '../src/formats/xlsx/build';
import { open, type Workbook } from '../src/index';
import { MemoryServer } from '../src/collab/server';
import { CollabSession } from '../src/collab/session';
import { contentVersion } from '../src/collab/protocol';
import { setCellInput } from '../src/ops/commands';
import { deleteSheet, insertSheet, moveSheet, renameSheet } from '../src/ops/sheets';
import type { Op } from '../src/ops/ops';
import { bookOf } from './chart-edit-fixtures';

type SOp = { k: 'ins'; at: number } | { k: 'del'; at: number } | { k: 'move'; from: number; to: number } | { k: 'name'; at: number };
const all = (n: number): SOp[] => { const out: SOp[] = []; for (let i = 0; i <= n; i++) out.push({ k: 'ins', at: i }); for (let i = 0; i < n; i++) out.push({ k: 'del', at: i }); for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) if (i !== j) out.push({ k: 'move', from: i, to: j }); for (let i = 0; i < n; i++) out.push({ k: 'name', at: i }); return out; };
let uniq = 0;
const make = (wb: Workbook, op: SOp): Op[] => op.k === 'ins' ? insertSheet(wb, op.at, { name: `New${++uniq}` }) : op.k === 'del' ? deleteSheet(wb, op.at) : op.k === 'move' ? moveSheet(wb, op.from, op.to) : renameSheet(wb, op.at, `Ren${++uniq}`);
const idOf = (wb: Workbook, i: number): string => { const row = wb.sheets[i].rows.get(0); const c = row && row.cells[0]; return c ? String(c.value) : `new:${wb.sheets[i].name}`; };
async function space(two: boolean): Promise<{ runs: number; bad: Map<string, string> }> {
  const bytes = await buildXlsx(bookOf([['ID-a']], [{ name: 'B', rows: [['ID-b']] }, { name: 'C', rows: [['ID-c']] }]));
  const bad = new Map<string, string>(); let runs = 0; 
  const scenario = async (as: SOp[], b: SOp) => {
    const lanes: (() => void)[][] = []; const byLane = new Map<string, (() => void)[]>();
    const server = new MemoryServer({ deliver: (fn, lane) => { let q = byLane.get(lane); if (!q) { byLane.set(lane, (q = [])); lanes.push(q); } q.push(fn); } });
    const drain = () => { for (let again = true; again; ) { again = false; for (const q of lanes) while (q.length) { q.shift()!(); again = true; } } };
    const mk = async (name: string) => { const wb = await (await open(bytes)).workbook(); const notices: any[] = []; const session = new CollabSession({ workbook: () => wb, changed: () => undefined, notice: (n) => notices.push(n) }, { transport: server.connect(), versions: [contentVersion(bytes)], user: { name }, timer: () => () => undefined }); return { wb, session, notices }; };
    const x = await mk('x'); const y = await mk('y'); drain();
    // both make their steps before anything is delivered; Y's is sent first
    try { y.session.submit(make(y.wb, b)); } catch (e) { if (e instanceof RangeError) return; throw e; }
    const yUp = [...byLane.entries()];
    try { for (const a of as) x.session.submit(make(x.wb, a)); } catch (e) { if (e instanceof RangeError) return; throw e; }
    const targets: { id: string; tok: string }[] = [];
    for (let i = 0; i < x.wb.sheets.length; i++) { const tok = `T${i}`; targets.push({ id: idOf(x.wb, i), tok }); x.session.submit(setCellInput(x.wb, i, 5, 1, tok)); }
    // deliver Y's op to the relay first: find y's up lane (the one whose queue was filled by y)
    { const q = byLane.get('2:up')!; if (q.length !== 1) throw new Error('lane'); q.shift()!(); }
    drain(); runs++;
    const label = `${as.map((a) => JSON.stringify(a)).join(' then ')} | remote ${JSON.stringify(b)}`; const cls = `${as.map((a) => a.k).join('+')} x ${b.k}`;
    const say = (what: string) => { if (!bad.has(cls + ': ' + what.replace(/T\d|ID-.|new:\w+/g, '#'))) bad.set(cls + ': ' + what.replace(/T\d|ID-.|new:\w+/g, '#'), `${label} -> ${what}`); (bad as any).n = ((bad as any).n || 0) + 1; };
    const names = (wb: Workbook) => wb.sheets.map((s, i) => idOf(wb, i) + '=' + s.name).join(',');
    if (names(x.wb) !== names(y.wb)) say(`DIVERGED ${names(x.wb)} / ${names(y.wb)}`);
    const dropped = x.notices.filter((n) => n.code === 'dropped').reduce((a, n) => a + n.count, 0);
    for (const t of targets) { let found = ''; y.wb.sheets.forEach((s, i) => { const row = s.rows.get(5); if (row && row.cells[1] && row.cells[1].value === t.tok) found = idOf(y.wb, i); });
      const there = y.wb.sheets.some((_s, i) => idOf(y.wb, i) === t.id);
      if (found && found !== t.id) say(`token meant for ${t.id} stands in ${found}`);
      else if (!found && there && !dropped) say(`token for ${t.id} lost though the sheet is there and nothing was said`);
      else if (!found && there) say(`(dropped and said) token for ${t.id} though the sheet is there`); }
  };
  const ops = all(3);
  for (const a of ops) for (const b of ops) { await scenario([a], b); if (two) for (const a2 of all(a.k === 'ins' ? 4 : a.k === 'del' ? 2 : 3)) await scenario([a, a2], b); }
  return { runs, bad };
}

for (const two of [false, true]) {
  test(`sheets: ${two ? 'two waiting steps' : 'one waiting step'} beside a step of someone else, every case`, async () => {
    const { runs, bad } = await space(two);
    const wrong = Array.from(bad.entries()).filter(([kind]) => kind.indexOf('(dropped and said)') < 0);
    assert.deepEqual(wrong.map(([kind, example]) => `${kind} - e.g. ${example}`), [], `of ${runs} cases, those that end with a token on the wrong sheet, lost unsaid, or with other sheets for the two`);
    assert.ok(runs > (two ? 4000 : 200), `the whole space was run: ${runs} cases`);
  });
}
