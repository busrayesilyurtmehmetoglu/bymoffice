// Random editing of a workbook and of a document with saves in between, for the tests of writing
// a file in one go. The engine is handed in (`load`), so that the same run can be made with two
// builds of it and their files compared byte for byte.

import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import type { DocEdit, DocFragment } from '../src/ops/doc/index';
import type { Op } from '../src/ops/ops';
import type { Workbook } from '../src/index';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type Mod = any;

/** The modules of an engine that the runs use, from the folder that holds its `src` and `test`. */
export interface Engine {
  root: string;
  index: Mod;
  commands: Mod;
  charts: Mod;
  filter: Mod;
  format: Mod;
  notes: Mod;
  ops: Mod;
  rules: Mod;
  sheets: Mod;
  sort: Mod;
  doc: Mod;
  docxRead: Mod;
  docxWrite: Mod;
  xlsxWrite: Mod;
  opc: Mod;
  collabFixtures: Mod;
  docFixtures: Mod;
  commentFixtures: Mod;
  docxFixtures: Mod;
}

export async function load(root: string): Promise<Engine> {
  const from = (path: string): Promise<Mod> => import(`${root}/${path}`);
  return {
    root,
    index: await from('src/index'),
    commands: await from('src/ops/commands'),
    charts: await from('src/ops/charts'),
    filter: await from('src/ops/filter'),
    format: await from('src/ops/format'),
    notes: await from('src/ops/notes'),
    ops: await from('src/ops/ops'),
    rules: await from('src/ops/rules'),
    sheets: await from('src/ops/sheets'),
    sort: await from('src/ops/sort'),
    doc: await from('src/ops/doc/index'),
    docxRead: await from('src/formats/docx/document'),
    docxWrite: await from('src/formats/docx/write'),
    xlsxWrite: await from('src/formats/xlsx/write'),
    opc: await from('src/opc'),
    collabFixtures: await from('test/collab-fixtures'),
    docFixtures: await from('test/doc-ops-fixtures'),
    commentFixtures: await from('test/doc-comments-fixtures'),
    docxFixtures: await from('test/docx-fixtures'),
  };
}

const TEXTS = ['12', '3.5', 'text', '=A1+1', '=SUM(B2:B5)', '=Other!A1*2', '=Data!C3', '', 'x y', '=B3&"!"', '=SUM(A:A)', '100%', '2024-01-05', '=IF(B2>3,1,0)'];
const NAMES = ['Veri', 'Yeni Ad', 'Sayfa1', 'Other', 'Data', 'S 2', 'Rapor'];

/** The kinds of change a workbook takes, each with how often it is picked. */
export function sheetChanges(e: Engine, wb: Workbook, random: () => number): [name: string, weight: number, make: () => Op[]][] {
  const { int, pick } = e.collabFixtures;
  const { setCellInput, clearRange, copyRange, pasteAt } = e.commands;
  const { addChart, copyChart, duplicateChart, moveChart, pasteChart, removeChart, sheetCharts, updateChart } = e.charts;
  const { deleteColumns, deleteRows, setAutoFilter } = e.filter;
  const { formatRange, mergeRange, setColumnWidth, setRowHeight, unmergeRange } = e.format;
  const { setComment } = e.notes;
  const { addConditionalRule, conditionalRules, removeConditionalRule, setValidation, updateConditionalRule } = e.rules;
  const { copySheet, deleteSheet, insertSheet, moveSheet, renameSheet, setSheetState, setTabColor } = e.sheets;
  const { sortRange } = e.sort;
  const s = int(random, 0, wb.sheets.length - 1);
  const row = int(random, 0, 11);
  const col = int(random, 0, 5);
  const range = { top: row, left: col, bottom: Math.min(row + int(random, 0, 3), 14), right: Math.min(col + int(random, 0, 2), 7) };
  const first = (): number => (sheetCharts(wb, s)[0] || { id: 0 }).id;
  return [
    ['type', 24, () => setCellInput(wb, s, row, col, pick(random, TEXTS))],
    ['clear', 3, () => clearRange(wb, s, range)],
    ['style', 10, () => formatRange(wb, s, range, pick(random, [{ bold: true }, { italic: true }, { fill: { rgb: 'FFFFFF00' } }, { numFmt: '0.00' }, { numFmt: 'dd.mm.yyyy' }, { horizontal: 'center' as const }, { border: { preset: 'all' as const } }, { fontSize: 14 + int(random, 0, 9) }, { fill: { rgb: `FF00${int(random, 10, 99)}00` } }]))],
    ['paste', 3, () => pasteAt(wb, s, { top: row, left: col, bottom: row, right: col }, 'unused', copyRange(wb, s, { top: int(random, 0, 5), left: int(random, 0, 3), bottom: int(random, 5, 7), right: int(random, 3, 4) })).ops],
    ['width', 2, () => setColumnWidth(wb, s, col, col + int(random, 0, 2), random() < 0.2 ? undefined : int(random, 4, 30))],
    ['height', 2, () => setRowHeight(s, row, row + int(random, 0, 2), random() < 0.2 ? undefined : int(random, 10, 40))],
    ['merge', 4, () => mergeRange(wb, s, { ...range, bottom: range.bottom + 1 })],
    ['unmerge', 2, () => unmergeRange(wb, s, range)],
    ['insert rows', 7, () => [{ type: 'insert', sheet: s, axis: 'row', at: row, count: int(random, 1, 3) } as Op]],
    ['delete rows', 6, () => deleteRows(wb, s, row, int(random, 1, 3))],
    ['insert columns', 5, () => [{ type: 'insert', sheet: s, axis: 'col', at: col, count: int(random, 1, 2) } as Op]],
    ['delete columns', 4, () => deleteColumns(wb, s, col, int(random, 1, 2))],
    ['sheet+', 4, () => insertSheet(wb, int(random, 0, wb.sheets.length), { name: `${pick(random, NAMES)} ${int(random, 0, 9)}` })],
    ['sheet-', 3, () => deleteSheet(wb, s)],
    ['rename', 4, () => renameSheet(wb, s, pick(random, NAMES))],
    ['move sheet', 3, () => moveSheet(wb, s, int(random, 0, wb.sheets.length - 1))],
    ['copy sheet', 3, () => copySheet(wb, s, { at: int(random, 0, wb.sheets.length) })],
    ['hide sheet', 1, () => setSheetState(wb, s, pick(random, ['visible', 'hidden'] as const))],
    ['tab', 2, () => setTabColor(wb, s, random() < 0.3 ? undefined : pick(random, ['#FF0000', '#00AA00', '#0000FF']))],
    ['note', 7, () => setComment(wb, s, row, col, random() < 0.25 ? undefined : `note ${int(random, 0, 99)}`, { author: 'me' })],
    ['validation', 4, () => setValidation(wb, s, range, random() < 0.25 ? undefined : pick(random, [{ type: 'whole' as const, operator: 'greaterThan' as const, formula1: '1' }, { type: 'list' as const, formula1: '$A$1:$A$4' }, { type: 'decimal' as const, operator: 'between' as const, formula1: 'B2', formula2: '100' }]))],
    ['rule+', 5, () => addConditionalRule(wb, s, range, pick(random, [{ type: 'cellIs' as const, operator: 'greaterThan' as const, formulas: ['15'], format: { fill: '#FFFF00' } }, { type: 'expression' as const, formulas: ['$B2>C2'], format: { bold: true } }, { type: 'top10' as const, rank: 3, format: { color: '#FF0000' } }]))],
    ['rule-', 2, () => removeConditionalRule(wb, s, int(random, 0, Math.max(conditionalRules(wb, s).length - 1, 0)))],
    ['rule change', 2, () => updateConditionalRule(wb, s, 0, { stopIfTrue: random() < 0.5, format: { italic: true } })],
    ['filter', 2, () => setAutoFilter(wb, s, wb.sheets[s].autoFilter && random() < 0.5 ? undefined : { top: 0, left: 0, bottom: 9, right: 3 })],
    ['sort', 4, () => sortRange(wb, s, { top: 1, left: 0, bottom: int(random, 3, 9), right: 3 }, [{ col: int(random, 0, 3), descending: random() < 0.5 }])],
    ['chart+', 5, () => addChart(wb, s, { type: pick(random, ['column', 'line', 'pie'] as const), range: { top: 0, left: 0, bottom: 4, right: 2 }, title: `Chart ${int(random, 0, 9)}` })],
    ['chart-', 2, () => removeChart(wb, s, int(random, 0, Math.max(sheetCharts(wb, s).length - 1, 0)))],
    ['chart copy', 3, () => duplicateChart(wb, s, first()).ops],
    ['chart paste', 2, () => pasteChart(wb, int(random, 0, wb.sheets.length - 1), copyChart(wb, s, first()), random() < 0.5 ? { row, col } : undefined).ops],
    ['chart move', 1, () => moveChart(wb, s, first(), { top: row, left: col, bottom: row + 8, right: col + 5 })],
    ['chart change', 1, () => updateChart(wb, s, first(), { title: `Title ${int(random, 0, 9)}`, legend: pick(random, ['r', 'b', 'none'] as const) })],
  ];
}

/** One random change of a workbook, applied; the name of what was done and what takes it back, undefined when the workbook took nothing. */
export function changeSheet(e: Engine, wb: Workbook, random: () => number, only?: string): { name: string; undo: Op[] } | undefined {
  const kinds = sheetChanges(e, wb, random);
  let picked = only === undefined ? undefined : kinds.find((kind) => kind[0] === only);
  if (!picked) {
    let total = 0;
    for (const kind of kinds) total += kind[1];
    let roll = random() * total;
    for (const kind of kinds) {
      roll -= kind[1];
      if (roll < 0) {
        picked = kind;
        break;
      }
    }
  }
  if (!picked) return undefined;
  let ops: Op[];
  try {
    ops = picked[2]();
  } catch (error) {
    // A command that does not apply where it was tried says so before it changes anything.
    return undefined;
  }
  if (!ops || ops.length === 0) return undefined;
  const undo: Op[] = [];
  try {
    for (const op of ops) undo.unshift(e.ops.applyOp(wb, op));
  } catch (error) {
    // An operation the workbook refuses: what was applied of the step is taken back.
    if ((error as Error).name !== 'RangeError') throw error;
    for (const op of undo) e.ops.applyOp(wb, op);
    return undefined;
  }
  return { name: picked[0], undo };
}

/** The file the runs of a workbook start with. */
export function sheetStart(e: Engine): Promise<Uint8Array> {
  return e.collabFixtures.startFile();
}

/**
 * A run of a workbook: changes, undoing, saves, and going on with the file that was saved. Answers
 * every file that was saved, in order, and what was done.
 */
export async function sheetRun(e: Engine, seed: number, start: Uint8Array, saved?: (model: object, trail: string[]) => void): Promise<{ files: Uint8Array[]; trail: string[] }> {
  const random: () => number = e.collabFixtures.seeded(seed);
  const { int } = e.collabFixtures;
  let file = await e.index.open(start);
  let wb: Workbook = await file.workbook();
  let undo: Op[][] = [];
  const files: Uint8Array[] = [];
  const trail: string[] = [];
  const steps = int(random, 8, 36);
  for (let step = 0; step < steps; step++) {
    const roll = random();
    if (roll < 0.08 && undo.length > 0) {
      for (const op of undo.pop() as Op[]) e.ops.applyOp(wb, op);
      trail.push('undo');
    } else if (roll < 0.2) {
      files.push(await file.save());
      if (saved) saved(wb, trail);
      trail.push('SAVE');
    } else if (roll < 0.25) {
      const bytes = await file.save();
      if (saved) saved(wb, trail);
      files.push(bytes);
      file = await e.index.open(bytes);
      wb = await file.workbook();
      undo = [];
      trail.push('REOPEN');
    } else {
      const made = changeSheet(e, wb, random);
      if (made) {
        undo.push(made.undo);
        trail.push(made.name);
      }
    }
  }
  files.push(await file.save());
  // Everything undone: the file as it was opened last, if the writer gives back what it took.
  while (undo.length > 0) for (const op of undo.pop() as Op[]) e.ops.applyOp(wb, op);
  files.push(await file.save());
  return { files, trail };
}

const NOW = (): Date => new Date('2026-10-12T08:09:10Z');
const hex = (n: number): string => ('0000000' + n.toString(16).toUpperCase()).slice(-8);
const COMMENT_TEXTS = ['Short.', 'Two\nlines', 'With <markup> & "quotes"', 'çığ 😀', '', 'tab\there'];

/** The file the runs of a document start with: five kinds (no comments; comments; with answers; with every part beside them; tracked). */
export function docStart(e: Engine, kind: number): Promise<Uint8Array> {
  const { p, r, tbl, tr, tc } = e.docFixtures;
  const { cStart, cEnd, cRef, commentedSpec } = e.commentFixtures;
  const ids = [1, 2, 3, 4];
  const body =
    p(r('One ') + cStart(1) + cStart(2) + r('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + r(' three ') + cStart(3) + r('four') + cEnd(3) + cRef(3) + r(' five.')) +
    p(r('A second paragraph with several words in it.')) +
    tbl(tr(tc('Cell A') + tc('Cell B')) + tr(tc('Cell C') + tc('Cell D')), [2000, 2000]) +
    cStart(4) +
    p('A paragraph that a comment holds whole.') +
    cEnd(4) +
    p(cRef(4) + r('The last paragraph.'));
  const plain = p('One two three four five.') + p('A second paragraph.') + tbl(tr(tc('Cell A') + tc('Cell B')) + tr(tc('Cell C') + tc('Cell D')), [2000, 2000]) + p('The last paragraph.');
  const comments = ids.map((id) => ({ id, author: 'A' + id, text: 'T' + id, paraId: kind === 1 ? undefined : hex(id + 9) }));
  const extended = ids.map((id): [string, string?, boolean?] => (id === 2 ? [hex(11), hex(10)] : id === 3 ? [hex(12), undefined, true] : [hex(id + 9)]));
  let spec: unknown;
  if (kind === 0) spec = { body: plain };
  else if (kind === 1) spec = { body, comments };
  else if (kind === 2) spec = { body, comments, extended };
  else if (kind === 4) spec = { body, comments, extended, settings: '<w:trackRevisions/>' };
  else spec = { body, comments, extended, ids: ids.map((id): [string, string] => [hex(id + 9), hex(0x10000000 + id)]), extensible: [[hex(0x10000001), '2026-10-01T09:00:00Z']] };
  return e.docxFixtures.buildDocx(commentedSpec(spec));
}

/** One random edit of a document; undefined when nothing applies where it was tried. */
export function docEdit(e: Engine, doc: unknown, d: { next(): number; pick<T>(list: T[]): T; chance(p: number): boolean }, clipboard: { fragment?: DocFragment }, only?: string): { name: string; edit: DocEdit } | undefined {
  const { paragraphPaths, randomEdit, randomPosition } = e.docFixtures;
  const { acceptAll, addComment, deleteRange, editComment, fragmentOf, insertFragment, listComments, rejectAll, removeComment, replyToComment, resolveComment, setTrackRevisions } = e.doc;
  const paths = paragraphPaths(doc, 'body').filter((path: number[]) => path.indexOf(-1) < 0);
  const roll = d.next();
  const kind = only !== undefined ? only : roll < 0.3 ? 'comment' : roll < 0.45 ? 'other' : 'edit';
  try {
    if (kind === 'comment' || kind === 'add' || kind === 'reply' || kind === 'remove') {
      const list = listComments(doc);
      const name = kind !== 'comment' ? kind : list.length === 0 ? 'add' : d.pick(['add', 'add', 'edit', 'reply', 'reply', 'resolve', 'remove']);
      if (list.length === 0 && name !== 'add') return undefined;
      const any = (): number => d.pick(list as { id: number }[]).id;
      if (name === 'add') return { name, edit: addComment(doc, { anchor: randomPosition(doc, d, 'body', paths), focus: randomPosition(doc, d, 'body', paths) }, d.pick(COMMENT_TEXTS)) };
      if (name === 'edit') return { name, edit: editComment(doc, any(), d.pick(COMMENT_TEXTS)) };
      if (name === 'reply') return { name, edit: replyToComment(doc, any(), d.pick(COMMENT_TEXTS)) };
      if (name === 'resolve') return { name, edit: resolveComment(doc, any(), d.chance(0.6)) };
      return { name, edit: removeComment(doc, any()) };
    }
    if (kind === 'other' || kind === 'track') {
      const name = kind === 'track' ? 'track' : d.pick(['track', 'acceptAll', 'rejectAll', 'copy', 'cut', 'paste']);
      if (name === 'track') return { name, edit: setTrackRevisions(doc, !(doc as { settings: { trackRevisions?: boolean } }).settings.trackRevisions) };
      if (name === 'acceptAll') return { name, edit: acceptAll(doc) };
      if (name === 'rejectAll') return { name, edit: rejectAll(doc) };
      const selection = { anchor: randomPosition(doc, d, 'body', paths), focus: randomPosition(doc, d, 'body', paths) };
      if (name === 'copy') {
        clipboard.fragment = fragmentOf(doc, selection);
        return undefined;
      }
      if (name === 'cut') {
        clipboard.fragment = fragmentOf(doc, selection);
        return { name, edit: deleteRange(doc, selection) };
      }
      return clipboard.fragment ? { name, edit: insertFragment(doc, selection.anchor, clipboard.fragment) } : undefined;
    }
    return randomEdit(doc, d, ['body']);
  } catch (error) {
    // A place where the command does not apply: nothing was changed.
    if ((error as Error).name === 'RangeError') return undefined;
    throw error;
  }
}

export interface DocSession {
  pkg: { save(): Promise<Uint8Array> };
  doc: unknown;
  history: { canUndo: boolean; canRedo: boolean; undo(): void; redo(): void; applyEdit(edit: DocEdit, meta: undefined): void; seal(): void };
}

export async function docOpen(e: Engine, bytes: Uint8Array): Promise<DocSession> {
  const pkg = await e.opc.OpcPackage.open(bytes);
  const doc = await e.docxRead.readDocx(pkg);
  e.doc.setTracking(doc, { author: 'Me', now: NOW });
  return { pkg, doc, history: new e.doc.DocHistory(doc) };
}

/** A run of a document: edits, comments, tracking, undoing, saves, and going on with the file that was saved. */
export async function docRun(e: Engine, seed: number, kind: number, saved?: (model: object, trail: string[]) => void): Promise<{ files: Uint8Array[]; trail: string[] }> {
  const d = e.docFixtures.dice(seed * 977 + kind * 13);
  let o = await docOpen(e, await docStart(e, kind));
  const files: Uint8Array[] = [];
  const trail: string[] = [];
  const clipboard: { fragment?: DocFragment } = {};
  const save = async (): Promise<Uint8Array> => {
    await e.docxWrite.writeDocx(o.pkg, o.doc);
    if (saved) saved(o.doc as object, trail);
    const bytes = await o.pkg.save();
    files.push(bytes);
    return bytes;
  };
  const steps = 10 + Math.floor(d.next() * 30);
  for (let step = 0; step < steps; step++) {
    const roll = d.next();
    if (roll < 0.1 && o.history.canUndo) {
      o.history.undo();
      trail.push('undo');
    } else if (roll < 0.14 && o.history.canRedo) {
      o.history.redo();
      trail.push('redo');
    } else if (roll < 0.19) {
      o = await docOpen(e, await save());
      trail.push('REOPEN');
    } else if (roll < 0.3) {
      await save();
      trail.push('SAVE');
    } else {
      const made = docEdit(e, o.doc, d, clipboard);
      if (!made || made.edit.ops.length === 0) continue;
      o.history.applyEdit(made.edit, undefined);
      o.history.seal();
      trail.push(made.name);
    }
  }
  await save();
  while (o.history.canUndo) o.history.undo();
  await save();
  return { files, trail };
}

export function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

/**
 * Waits `n` moments: below 100 that many turns of the promises; from 100 on that many turns of the event loop
 * (what reads a part comes back within those); from 1000 on that many milliseconds (the file is done by then).
 */
export async function moments(n: number): Promise<void> {
  if (n < 100) for (let i = 0; i < n; i++) await Promise.resolve();
  else if (n < 1000) for (let i = 100; i < n; i++) await new Promise<void>((resolve) => setImmediate(resolve));
  else await new Promise<void>((resolve) => setTimeout(resolve, n - 1000));
}

/** How long after a save has started something is changed, in the runs that change a model while its file is written. */
export const RACE_DELAYS = [0, 1, 2, 3, 5, 8, 13, 20, 30, 50, 90, 101, 102, 103, 104, 106, 108, 112, 116, 124, 140, 170, 1001, 1003];

/**
 * What a reader that knows nothing of this engine finds wrong with files (test/tools/package_check.py:
 * Python's zipfile and ElementTree); an empty list when Python is not there.
 */
export function independent(files: { [name: string]: Uint8Array }): string[] {
  try {
    execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
  } catch {
    return [];
  }
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.race-'));
  try {
    for (const name of Object.keys(files)) writeFileSync(joinPath(dir, `${name}.bin`), files[name]);
    return JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 300000, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'inherit'] }));
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}
