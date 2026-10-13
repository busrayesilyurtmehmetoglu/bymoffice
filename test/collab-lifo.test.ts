// The rule of this round with others in the room: an entry kept for undo is read through what others did, as it
// reads on the workbook the entry was written for (through the entries above it, undone). A writes cells (named by
// sheet name + row tag) and inserts/renames sheets; B puts rows in and takes rows out on any sheet and writes in
// columns A never uses. Then A undoes everything. Every undo that is done must put back exactly the cell it changed.
// Written by a reviewer, as a run of many seeds; kept as it was written.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx } from '../src/formats/xlsx/build';
import { bookOf } from './chart-edit-fixtures';
import { setCellInput } from '../src/ops/commands';
import { deleteRows } from '../src/ops/filter';
import { insertSheet, renameSheet } from '../src/ops/sheets';
import { Net, join, seeded, int, canon } from './collab-fixtures';
const tot = { runs: 0, steps: 0, undone: 0, refused: 0, wrong: 0, diverge: 0, thrown: 0, unsaid: 0 };
const out: string[] = [];
async function fuzz(first: number, last: number): Promise<{ tot: { wrong: number; diverge: number; thrown: number }; out: string[] }> {
  tot.runs = tot.steps = tot.undone = tot.refused = tot.wrong = tot.diverge = tot.thrown = tot.unsaid = 0;
  out.length = 0;
  for (let seed = first; seed <= last; seed++) {
    tot.runs++;
    try {
      const rnd = seeded(seed * 97 + 3);
      const mk = (p: string) => Array.from({ length: 8 }, (_, r) => [`${p}${r}`, `${p}${r}b`, `${p}${r}c`, `${p}${r}d`] as (string | null)[]);
      const bytes = await buildXlsx(bookOf(mk('D'), [{ name: 'Other', rows: mk('O') }, { name: 'Third', rows: mk('T') }]));
      const net = new Net();
      const a = await join(net, 'a', bytes), b = await join(net, 'b', bytes);
      net.drain(rnd);
      const find = (wb: any, tag: string): { sheet: any; s: number; r: number } | undefined => { for (let s = 0; s < wb.sheets.length; s++) { let hit: number | undefined; wb.sheets[s].rows.forEach((row: any, r: number) => { if (row.cells[0] && row.cells[0].value === tag) hit = r; }); if (hit !== undefined) return { sheet: wb.sheets[s], s, r: hit }; } return undefined; };
      const cellOf = (wb: any, tag: string, c: number): unknown => { const at = find(wb, tag); if (!at) return undefined; const cell = at.sheet.rows.get(at.r).cells[c]; return cell ? cell.value : null; };
      type Step = { t: 'w'; tag: string; c: number; old: unknown; v: string } | { t: 'sheet'; name: string } | { t: 'name'; old: string; name: string };
      const mine: Step[] = [];
      const script: string[] = [];
      let uniq = 0;
      const n = int(rnd, 3, 14);
      for (let k = 0; k < n; k++) {
        const roll = rnd();
        try {
          if (roll < 0.45) {
            const s = int(rnd, 0, a.wb.sheets.length - 1), r = int(rnd, 0, 7), c = int(rnd, 1, 3);
            const row = a.wb.sheets[s].rows.get(r); const tag = row && row.cells[0] ? String(row.cells[0].value) : undefined;
            if (tag && /^[DOT]\d$/.test(tag)) { const v = `a${++uniq}`; const old = cellOf(a.wb, tag, c); const made = a.session.stats.made; a.session.submit(setCellInput(a.wb, s, r, c, v)); if (a.session.stats.made > made) { mine.push({ t: 'w', tag, c, old, v }); script.push(`A:w(${tag},${c})`); } }
          } else if (roll < 0.58) { const name = `S${++uniq}`; a.session.submit(insertSheet(a.wb, int(rnd, 0, a.wb.sheets.length), { name })); mine.push({ t: 'sheet', name }); script.push(`A:sheet+(${name})`); }
          else if (roll < 0.64) { const s = int(rnd, 0, a.wb.sheets.length - 1); const old = a.wb.sheets[s].name; const name = `N${++uniq}`; a.session.submit(renameSheet(a.wb, s, name)); mine.push({ t: 'name', old, name }); script.push(`A:name(${old}->${name})`); }
          else if (roll < 0.82) { const s = int(rnd, 0, b.wb.sheets.length - 1); const at = int(rnd, 0, 7); b.session.submit([{ type: 'insert', sheet: s, axis: 'row', at, count: int(rnd, 1, 2) }]); script.push(`B:ins(${b.wb.sheets[s].name}@${at})`); }
          else if (roll < 0.92) { const s = int(rnd, 0, b.wb.sheets.length - 1); const at = int(rnd, 0, 7); b.session.submit(deleteRows(b.wb, s, at, 1)); script.push(`B:del(${b.wb.sheets[s].name}@${at})`); }
          else { const s = int(rnd, 0, b.wb.sheets.length - 1); b.session.submit(setCellInput(b.wb, s, int(rnd, 0, 9), 6, `b${++uniq}`)); script.push('B:w'); }
        } catch (e) { if (!(e instanceof RangeError)) throw e; }
        if (rnd() < 0.7) net.drain(rnd); else for (let j = int(rnd, 0, 3); j > 0; j--) net.step(rnd);
      }
      net.drain(rnd);
      // which of A's steps were placed at all (a dropped one is said, and is no step)
      const dropped = a.notices.filter((x) => x.code === 'dropped').reduce((s, x) => s + (x.count || 0), 0);
      if (dropped > 0) continue; // keep the oracle simple: runs in which A lost nothing on the way
      tot.steps += mine.length;
      let bad = '';
      for (let i = mine.length - 1; i >= 0 && !bad; i--) {
        const step = mine[i];
        const before = a.notices.length;
        const done = a.session.undo();
        net.drain(rnd);
        const said = a.notices.slice(before).map((x) => x.code);
        if (done) {
          tot.undone++;
          if (step.t === 'w') { const now = cellOf(a.wb, step.tag, step.c); if (now !== undefined && now !== step.old) bad = `undo of w(${step.tag},${step.c}) gave ${now}, was ${step.old}`; }
          else if (step.t === 'sheet') { if (a.wb.sheets.some((s: any) => s.name === step.name)) bad = `undo of sheet+ left ${step.name}`; }
          else if (!a.wb.sheets.some((s: any) => s.name === step.old)) bad = `undo of the name did not bring ${step.old} back`;
        } else {
          tot.refused++;
          if (said.length === 0) { tot.unsaid++; bad = `undo of step ${i} (${step.t}) did nothing and said nothing`; }
        }
        // whatever became of this undo, no other cell of A's columns may have changed: all tags in cols 1..3 hold the file's value or a value of a step of A not yet undone
        if (!bad) for (const p of ['D', 'O', 'T']) for (let r = 0; r < 8 && !bad; r++) for (let c = 1; c <= 3 && !bad; c++) {
          const tag = `${p}${r}`; const now = cellOf(a.wb, tag, c); if (now === undefined) continue;
          const orig = `${tag}${' bcd'[c]}`;
          const standing = mine.slice(0, i + (done ? 0 : 1)).filter((x) => x.t === 'w' && (x as any).tag === tag && (x as any).c === c).map((x: any) => x.v);
          const refusedLater = mine.slice(i).filter((x) => x.t === 'w' && (x as any).tag === tag && (x as any).c === c).map((x: any) => x.v);
          if (now !== orig && standing.indexOf(now as string) < 0 && refusedLater.indexOf(now as string) < 0) bad = `after undo of step ${i}, ${tag}:${c} is ${now} (file ${orig}, standing ${standing})`;
        }
      }
      if (bad) { tot.wrong++; if (out.length < 8) out.push(`WRONG seed ${seed}: ${bad}; notices ${a.notices.map((x) => x.code)} :: ${script.join(' ')}`); }
      if (canon(a.wb) !== canon(b.wb)) { tot.diverge++; if (out.length < 8) out.push(`DIVERGE seed ${seed}`); }
    } catch (e) { tot.thrown++; if (out.length < 8) out.push(`THROWN seed ${seed}: ${(e as Error).stack?.split('\n').slice(0, 4).join(' | ')}`); }
  }
  return { tot: tot as { wrong: number; diverge: number; thrown: number }, out };
}

const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));

test(`an undo beside rows that others put in and took out puts back the cell it changed, and no other (${SEEDS} runs)`, async () => {
  const { tot, out } = await fuzz(1, SEEDS);
  assert.deepEqual([tot.wrong, tot.diverge, tot.thrown, out], [0, 0, 0, []], 'no cell that is wrong, no workbooks that differ, nothing thrown');
});
