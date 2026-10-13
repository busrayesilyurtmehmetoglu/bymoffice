// Random runs of collaborative editing: two to five participants on one relay make changes of
// every kind at once, while the test delivers the messages in whatever order it likes and
// cuts connections. What has to hold when the messages have run out: everybody has the same
// workbook, every step that was dropped was said, the files that are saved are whole and read
// back the same, and it still holds after everybody has undone everything.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { clearRange, copyRange, pasteAt, setCellInput } from '../src/ops/commands';
import { addChart, moveChart, removeChart, sheetCharts, updateChart } from '../src/ops/charts';
import { deleteColumns, deleteRows, setAutoFilter, setFilterColumn } from '../src/ops/filter';
import { formatRange, mergeRange, setColumnWidth, setColumnsHidden, setRowHeight, setRowsHidden, unmergeRange } from '../src/ops/format';
import { setComment } from '../src/ops/notes';
import type { Op } from '../src/ops/ops';
import { addConditionalRule, clearConditionalRules, conditionalRules, moveConditionalRule, removeConditionalRule, setValidation, updateConditionalRule } from '../src/ops/rules';
import { deleteSheet, insertSheet, moveSheet, renameSheet, setSheetState, setTabColor } from '../src/ops/sheets';
import { sortRange } from '../src/ops/sort';
import { contentVersion } from '../src/collab/protocol';
import { open } from '../src/index';
import { packageProblems } from './chart-edit-fixtures';
import { Net, act, assertSame, canon, firstDifference, int, join, pick, saveOf, seeded, startFile, type Member } from './collab-fixtures';

/** How many runs; more when asked for (COLLAB_SEEDS=2000). */
const SEEDS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));
const FIRST = Number(process.env.COLLAB_FIRST) || 1;

const TEXTS = ['12', '3.5', 'text', '=A1+1', '=SUM(B2:B5)', '=Other!A1*2', '=Data!C3', '', 'x y', '=B3&"!"', '=SUM(A:A)', '100%', '2024-01-05', '=IF(B2>3,1,0)'];
const NAMES = ['Veri', 'Yeni Ad', 'Sayfa1', 'Other', 'Data', 'S 2', 'Rapor'];

/** One random change of a participant; the name of what was done, undefined when the workbook took nothing. */
export function change(member: Pick<Member, 'wb' | 'name'>, random: () => number, apply: (make: () => Op[]) => boolean = (make) => act(member as Member, make)): string | undefined {
  const wb = member.wb;
  const s = int(random, 0, wb.sheets.length - 1);
  const row = int(random, 0, 11);
  const col = int(random, 0, 5);
  const range = { top: row, left: col, bottom: Math.min(row + int(random, 0, 3), 14), right: Math.min(col + int(random, 0, 2), 7) };
  const kinds: [string, number, () => Op[]][] = [
    ['type', 30, () => setCellInput(wb, s, row, col, pick(random, TEXTS))],
    ['clear', 4, () => clearRange(wb, s, range)],
    ['format', 8, () => formatRange(wb, s, range, pick(random, [{ bold: true }, { italic: true }, { fill: { rgb: 'FFFFFF00' } }, { numFmt: '0.00' }, { numFmt: 'dd.mm.yyyy' }, { horizontal: 'center' as const }, { border: { preset: 'all' as const } }, { fontSize: 14 }, { bold: false }]))],
    ['paste', 3, () => pasteAt(wb, s, { top: row, left: col, bottom: row, right: col }, 'unused', copyRange(wb, s, { top: int(random, 0, 5), left: int(random, 0, 3), bottom: int(random, 5, 7), right: int(random, 3, 4) })).ops],
    ['paste text', 2, () => pasteAt(wb, s, { top: row, left: col, bottom: row, right: col }, '1\t2\n=A1\ttext').ops],
    ['width', 3, () => setColumnWidth(wb, s, col, col + int(random, 0, 2), random() < 0.2 ? undefined : int(random, 4, 30))],
    ['hide columns', 2, () => setColumnsHidden(wb, s, col, col, random() < 0.6)],
    ['height', 3, () => setRowHeight(s, row, row + int(random, 0, 2), random() < 0.2 ? undefined : int(random, 10, 40))],
    ['hide rows', 2, () => setRowsHidden(s, row, row + int(random, 0, 1), random() < 0.6)],
    ['merge', 3, () => mergeRange(wb, s, { ...range, bottom: range.bottom + 1 })],
    ['unmerge', 2, () => unmergeRange(wb, s, range)],
    ['insert rows', 8, () => [{ type: 'insert', sheet: s, axis: 'row', at: row, count: int(random, 1, 3) }]],
    ['delete rows', 8, () => deleteRows(wb, s, row, int(random, 1, 3))],
    ['insert columns', 5, () => [{ type: 'insert', sheet: s, axis: 'col', at: col, count: int(random, 1, 2) }]],
    ['delete columns', 5, () => deleteColumns(wb, s, col, int(random, 1, 2))],
    ['sheet+', 3, () => insertSheet(wb, int(random, 0, wb.sheets.length), { name: `${pick(random, NAMES)} ${int(random, 0, 9)}` })],
    ['sheet-', 2, () => deleteSheet(wb, s)],
    ['rename', 3, () => renameSheet(wb, s, pick(random, NAMES))],
    ['move sheet', 2, () => moveSheet(wb, s, int(random, 0, wb.sheets.length - 1))],
    ['hide sheet', 1, () => setSheetState(wb, s, pick(random, ['visible', 'hidden'] as const))],
    ['tab', 1, () => setTabColor(wb, s, random() < 0.3 ? undefined : pick(random, ['#FF0000', '#00AA00', '#0000FF']))],
    ['note', 6, () => setComment(wb, s, row, col, random() < 0.25 ? undefined : `note ${int(random, 0, 99)}`, { author: member.name })],
    ['validation', 4, () => setValidation(wb, s, range, random() < 0.25 ? undefined : pick(random, [{ type: 'whole' as const, operator: 'greaterThan' as const, formula1: '1' }, { type: 'list' as const, formula1: '$A$1:$A$4' }, { type: 'decimal' as const, operator: 'between' as const, formula1: 'B2', formula2: '100' }]))],
    ['rule+', 4, () => addConditionalRule(wb, s, range, pick(random, [{ type: 'cellIs' as const, operator: 'greaterThan' as const, formulas: ['15'], format: { fill: '#FFFF00' } }, { type: 'expression' as const, formulas: ['$B2>C2'], format: { bold: true } }, { type: 'top10' as const, rank: 3, format: { color: '#FF0000' } }]))],
    ['rule-', 2, () => removeConditionalRule(wb, s, int(random, 0, Math.max(conditionalRules(wb, s).length - 1, 0)))],
    ['rule move', 1, () => moveConditionalRule(wb, s, 0, Math.max(conditionalRules(wb, s).length - 1, 0))],
    ['rule change', 2, () => updateConditionalRule(wb, s, 0, { stopIfTrue: random() < 0.5, format: { italic: true } })],
    ['rules clear', 1, () => clearConditionalRules(wb, s, random() < 0.5 ? undefined : range)],
    ['filter', 3, () => setAutoFilter(wb, s, wb.sheets[s].autoFilter && random() < 0.5 ? undefined : { top: 0, left: 0, bottom: 9, right: 3 })],
    ['filter column', 3, () => setFilterColumn(wb, s, int(random, 0, 3), random() < 0.3 ? undefined : pick(random, [{ kind: 'values' as const, values: ['10', '20', 'Row 1', 'Row 2'], blank: true }, { kind: 'custom' as const, conditions: [{ operator: 'greaterThan', value: '30' }] }]))],
    ['sort', 4, () => sortRange(wb, s, { top: 1, left: 0, bottom: int(random, 3, 9), right: 3 }, [{ col: int(random, 0, 3), descending: random() < 0.5 }])],
    ['chart+', 3, () => addChart(wb, s, { type: pick(random, ['column', 'line', 'pie'] as const), range: { top: 0, left: 0, bottom: 4, right: 2 }, title: `Chart ${int(random, 0, 9)}` })],
    ['chart-', 1, () => removeChart(wb, s, int(random, 0, Math.max(sheetCharts(wb, s).length - 1, 0)))],
    ['chart move', 1, () => moveChart(wb, s, (sheetCharts(wb, s)[0] || { id: 0 }).id, { top: row, left: col, bottom: row + 8, right: col + 5 })],
    ['chart change', 1, () => updateChart(wb, s, (sheetCharts(wb, s)[0] || { id: 0 }).id, { title: `Title ${int(random, 0, 9)}`, legend: pick(random, ['r', 'b', 'none'] as const) })],
  ];
  let total = 0;
  for (const kind of kinds) total += kind[1];
  let roll = random() * total;
  for (const [name, weight, make] of kinds) {
    roll -= weight;
    if (roll < 0) return apply(make) ? name : undefined;
  }
  return undefined;
}

function settle(net: Net, members: Member[], random: () => number): void {
  for (const member of members) member.transport.resume();
  net.drain(random);
  for (const member of members) {
    assert.equal(member.session.state, 'online', `${member.name} is online at the end`);
    assert.equal(member.session.pending, 0, `${member.name} has nothing left to send`);
  }
}

/** Every step a participant made was placed, or dropped and said. */
function accounted(members: Member[], label: string): void {
  for (const member of members) {
    // A step that was dropped is said as that, and an undoing that was dropped as an undoing that is gone (with how many).
    const said = member.notices.filter((notice) => notice.code === 'dropped' || (notice.code === 'undo-gone' && notice.count !== undefined)).reduce((sum, notice) => sum + (notice.count || 0), 0);
    const stats = member.session.stats;
    assert.equal(said, stats.dropped, `${label}: ${member.name} was told of every dropped step`);
    assert.equal(stats.placed + stats.dropped + stats.quiet, stats.made, `${label}: every step of ${member.name} was placed or dropped`);
    assert.equal(member.notices.filter((notice) => notice.code === 'peer' || notice.code === 'stale' || notice.code === 'refused').length, 0, `${label}: ${member.name} stayed in the room`);
  }
}

/** What the runs did in all: said at the end with COLLAB_STATS=1, to see that they do what they are for. */
export const totals: { [name: string]: number } = {};
const tally = (name: string, by = 1): void => {
  totals[name] = (totals[name] || 0) + by;
};

export async function run(seed: number, options: { files?: boolean; late?: boolean } = {}): Promise<void> {
  const random = seeded(seed);
  const bytes = await startFile();
  const net = new Net();
  const members: Member[] = [];
  const count = int(random, 2, 5);
  for (let i = 0; i < count; i++) members.push(await join(net, `p${i}`, bytes));
  const steps = int(random, 10, 60);
  const done: string[] = [];
  const lateAt = options.late ? int(random, 3, steps - 1) : -1;
  for (let i = 0; i < steps; i++) {
    const roll = random();
    if (roll < 0.55) {
      const member = pick(random, members);
      const what = change(member, random);
      if (what) done.push(`${member.name} ${what}`);
      tally(what ? `did ${what}` : 'refused');
    } else if (roll < 0.63) {
      const member = pick(random, members);
      if (random() < 0.5) member.session.undo();
      else member.session.redo();
    } else if (roll < 0.68) pick(random, members).transport.suspend();
    else if (roll < 0.74) pick(random, members).transport.resume();
    else for (let k = int(random, 1, 12); k > 0; k--) net.step(random);
    if (i === lateAt) {
      // Someone saves, says so, and a participant who reads that file joins.
      settle(net, members, random);
      const saver = pick(random, members);
      const rev = saver.session.savableRev;
      assert.notEqual(rev, undefined, 'a participant with nothing pending knows the step its file holds');
      const saved = await saveOf(saver);
      saver.session.announceSaved([contentVersion(saved)], rev);
      net.drain(random);
      members.push(await join(net, 'late', saved));
      net.drain(random);
      assertSame(members, `seed ${seed}: after a participant joined with the saved file (${done.join(', ')})`);
    }
  }
  settle(net, members, random);
  const label = `seed ${seed} (${done.join(', ')})`;
  assertSame(members, label);
  accounted(members, label);
  if (options.files) {
    const texts: string[] = [];
    for (const member of members) {
      const saved = await saveOf(member);
      assert.deepEqual(await packageProblems(saved), [], `${label}: the file of ${member.name} is whole`);
      texts.push(canon(await (await open(saved)).workbook()));
    }
    for (let i = 1; i < texts.length; i++) if (texts[i] !== texts[0]) assert.fail(`${label}: the file of ${members[i].name} reads differently ${firstDifference(texts[0], texts[i])}`);
  }
  for (const member of members) {
    tally('made', member.session.stats.made);
    tally('placed', member.session.stats.placed);
    tally('dropped', member.session.stats.dropped);
    tally('quiet', member.session.stats.quiet);
    for (const notice of member.notices) tally(`notice ${notice.code}`);
  }
  // Everybody takes back what can be taken back.
  // An undoing waits until the steps of its participant are placed: one at a time, with the messages delivered in between.
  for (let round = 0; round < 400 && members.some((member) => member.session.canUndo); round++) {
    for (const member of members) if (member.session.canUndo) member.session.undo();
    net.drain(random);
  }
  settle(net, members, random);
  assertSame(members, `${label}, undone`);
  accounted(members, `${label}, undone`);
}

test(`random runs end with the same workbook for everybody (${SEEDS} seeds)`, async () => {
  for (let seed = FIRST; seed < FIRST + SEEDS; seed++) await run(seed);
});

test('the runs did what they are for', () => {
  if (process.env.COLLAB_STATS) console.log(JSON.stringify(totals, null, 1));
  assert.ok(totals.placed > 0 && totals.dropped > 0, `steps were placed and steps were dropped: ${JSON.stringify(totals)}`);
});

test('random runs whose files are saved and read again', async () => {
  for (let seed = 5000; seed < 5000 + Math.max(10, Math.floor(SEEDS / 3)); seed++) await run(seed, { files: true });
});

test('random runs in which a participant joins with a file that was saved on the way', async () => {
  for (let seed = 9000; seed < 9000 + Math.max(10, Math.floor(SEEDS / 5)); seed++) await run(seed, { late: true });
});
