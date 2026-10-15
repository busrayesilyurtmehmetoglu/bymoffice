// What the tests of collaborative editing share: a starting file, participants on a relay of
// the page whose messages the test delivers in whatever order it likes (each connection keeps
// its own order, as a socket does), and a text that says what a workbook holds whatever file
// its participant read.

import assert from 'node:assert/strict';
import { buildXlsx } from '../src/formats/xlsx/build';
import { readCollabStamp, writeCollabStamp } from '../src/formats/xlsx/collab-stamp';
import { writeXlsx } from '../src/formats/xlsx/write';
import { open, type OfficeFile, type Sheet, type Workbook } from '../src/index';
import { contentVersion } from '../src/collab/protocol';
import { MemoryServer, type MemoryTransport } from '../src/collab/server';
import { CollabSession, type CollabNotice, type SessionChange, type SessionOptions } from '../src/collab/session';
import { stable, wireFormatOf } from '../src/collab/wire';
import { sheetCharts } from '../src/ops/charts';
import { cellNow } from '../src/ops/notes';
import { explicitFormula, sharedMasters, type Op } from '../src/ops/ops';
import { conditionalFormatNow, dataValidationNow, ruleSpace } from '../src/ops/rules';
import { bookOf } from './chart-edit-fixtures';

export function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export const pick = <T>(random: () => number, list: readonly T[]): T => list[Math.floor(random() * list.length)];
export const int = (random: () => number, low: number, high: number): number => low + Math.floor(random() * (high - low + 1));

/** A file with numbers, texts and formulas on two sheets. */
export async function startFile(): Promise<Uint8Array> {
  const rows: (number | string | null)[][] = [['Item', 'North', 'South', 'Sum']];
  for (let r = 1; r <= 8; r++) rows.push([`Row ${r}`, r * 10, 100 - r * 7, null]);
  const book = bookOf(rows, [{ name: 'Other', rows: [[1, 2, 3], [4, 5, 6]] }]);
  const data = book.sheets[0];
  for (let r = 1; r <= 8; r++) {
    const row = data.rows.get(r)!;
    row.cells[3] = { kind: 'number', value: r * 10 + 100 - r * 7, style: 0, formula: { text: `B${r + 1}+C${r + 1}`, type: 'normal' } };
  }
  data.rows.set(9, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'string', value: 'Total', style: 0 }, { kind: 'number', value: 360, style: 0, formula: { text: 'SUM(B2:B9)', type: 'normal' } }] });
  data.rowCount = 10;
  book.sheets[1].rows.get(0)!.cells[3] = { kind: 'number', value: 360, style: 0, formula: { text: 'Data!B10*2', type: 'normal' } };
  book.sheets[1].colCount = 4;
  return buildXlsx(book);
}

/** The messages on their way, by connection and direction. */
export class Net {
  readonly server: MemoryServer;
  private lanes = new Map<string, (() => void)[]>();

  constructor(options: { maxLogEntries?: number } = {}) {
    this.server = new MemoryServer({
      ...options,
      deliver: (run, lane) => {
        let waiting = this.lanes.get(lane);
        if (!waiting) this.lanes.set(lane, (waiting = []));
        waiting.push(run);
      },
    });
  }

  get waiting(): number {
    let n = 0;
    this.lanes.forEach((lane) => (n += lane.length));
    return n;
  }

  /** Delivers one message of a lane picked by `random`; false when nothing waits. */
  step(random: () => number): boolean {
    const busy = Array.from(this.lanes.values()).filter((lane) => lane.length > 0);
    if (busy.length === 0) return false;
    const run = pick(random, busy).shift()!;
    run();
    return true;
  }

  /** Delivers everything that is on its way to the relay, and nothing the relay sends. */
  deliverUp(): void {
    for (;;) {
      const busy = Array.from(this.lanes.entries()).filter(([name, lane]) => name.slice(-3) === ':up' && lane.length > 0);
      if (busy.length === 0) return;
      busy[0][1].shift()!();
    }
  }

  drain(random: () => number, limit = 1000000): void {
    let n = 0;
    while (this.step(random)) if (++n > limit) assert.fail('The messages never ran out');
  }
}

export interface Member {
  name: string;
  file: OfficeFile;
  wb: Workbook;
  session: CollabSession;
  transport: MemoryTransport;
  notices: CollabNotice[];
  changes: SessionChange[];
  /** Steps this participant made that were applied here (whether or not they stayed). */
  made: number;
  /** The file being loaded again, while it is: for a test to wait for, in place of guessing how long it takes. */
  loading?: Promise<void>;
}

/**
 * A participant: reads `bytes` and joins.
 */
export async function join(net: Net, name: string, bytes: Uint8Array, options: Partial<SessionOptions> = {}, reload?: () => Promise<Uint8Array>, unsaved?: () => boolean): Promise<Member> {
  const file = await open(bytes);
  const wb = await file.workbook();
  const transport = net.server.connect();
  const member: Member = { name, file, wb, session: undefined as unknown as CollabSession, transport, notices: [], changes: [], made: 0 };
  member.session = new CollabSession(
    {
      workbook: () => member.wb,
      changed: (change) => member.changes.push(change),
      notice: (notice) => member.notices.push(notice),
      unsaved,
      reload: reload
        ? () => {
            const load = (async () => {
              const next = await reload();
              member.file = await open(next);
              member.wb = await member.file.workbook();
              return { versions: [contentVersion(next)], stamp: await readCollabStamp(member.file.package!) };
            })();
            member.loading = load.then(
              () => new Promise<void>((resolve) => setImmediate(resolve)),
              () => undefined,
            );
            return load;
          }
        : undefined,
    },
    { transport, versions: [contentVersion(bytes)], stamp: await readCollabStamp(file.package!), user: { name }, timer: () => () => undefined, ...options },
  );
  return member;
}

export async function saveOf(member: Member): Promise<Uint8Array> {
  // As mount() saves a workbook that is edited together: the room and the step go into the file.
  await writeCollabStamp(member.file.package!, member.session.stampNow);
  await writeXlsx(member.file.package!, member.wb);
  return member.file.package!.save();
}

/** Applies operations as a step of a participant; undefined when the workbook refused them. */
export function act(member: Member, make: () => Op[]): boolean {
  try {
    const ops = make();
    if (ops.length === 0) return false;
    member.session.submit(ops);
    member.made++;
    return true;
  } catch (error) {
    if (error instanceof RangeError) return false;
    throw error;
  }
}

function sheetText(wb: Workbook, sheet: Sheet, index: number): unknown {
  const masters = sharedMasters(sheet);
  const format = (style: number | undefined): string => (style === undefined ? '' : stable(wireFormatOf(wb.styles, style)));
  const plain = format(0);
  const rows: unknown[] = [];
  for (const r of Array.from(sheet.rows.keys()).sort((a, b) => a - b)) {
    const row = sheet.rows.get(r)!;
    const cells: unknown[] = [];
    row.cells.forEach((cell, c) => {
      const explicit = explicitFormula(cell, r, c, masters);
      const look = format(cell.style);
      // A cell that holds nothing and looks like none is no cell.
      if (explicit.kind === 'blank' && !explicit.formula && look === plain) return;
      cells.push([c, explicit.kind, explicit.value, explicit.formula ? [explicit.formula.text, explicit.formula.type === 'shared' ? 'normal' : explicit.formula.type, explicit.formula.type === 'array' || explicit.formula.type === 'dataTable' || explicit.formula.dynamic === true ? explicit.formula.ref || null : null, explicit.formula.dynamic === true, explicit.formula.attrs ? Array.from(explicit.formula.attrs.entries()).filter(([name]) => name === 'aca' || name === 'ca' || name === 'bx').sort() : []] : null, look === plain ? '' : look, explicit.runs || null, explicit.ph === true, explicit.spilled === true]);
    });
    const props = [row.height === undefined ? null : row.height, row.customHeight, row.hidden, row.style === undefined ? '' : format(row.style), row.outlineLevel || 0, row.collapsed === true];
    if (cells.length > 0 || stable(props) !== stable([null, false, false, '', 0, false])) rows.push([r, props, cells]);
  }
  const columns: unknown[] = [];
  for (const range of sheet.columns) columns.push([range.min, range.max, range.width === undefined ? null : range.width, range.customWidth, range.hidden, range.style === undefined ? '' : format(range.style), range.outlineLevel || 0, range.collapsed === true]);
  const notes: unknown[] = [];
  for (const comment of sheet.comments || []) {
    const cell = cellNow(sheet, comment);
    if (cell) notes.push([cell.row, cell.col, comment.text, comment.author || '', comment.runs || null, comment.visible === true, comment.threaded === true]);
  }
  notes.sort((a, b) => (stable(a) < stable(b) ? -1 : 1));
  const space = ruleSpace(wb, sheet);
  const validations = (sheet.dataValidations || []).map((rule) => dataValidationNow(rule, space)).filter((rule) => rule !== undefined).map((rule) => ({ ...rule, src: undefined, live: undefined, kept: undefined }));
  const conditional = (sheet.conditionalFormats || []).map((item) => conditionalFormatNow(item, space)).filter((item) => item !== undefined).map((item) => ({ ranges: item!.ranges, rules: item!.rules.map((rule) => ({ ...rule, src: undefined, kept: undefined, dxfId: undefined, dxf: rule.dxf || (rule.dxfId !== undefined && wb.styles.dxfFormats ? wb.styles.dxfFormats[rule.dxfId] : undefined) })) }));
  const filter = sheet.autoFilter ? { range: sheet.autoFilter.range, // What a file says by default and what a command leaves unsaid are the same: "and" is false unless it is true.
        columns: sheet.autoFilter.columns.map((column) => ({ ...column, raw: undefined, and: column.and === true ? true : undefined })), sort: sheet.autoFilter.sortState ? { range: sheet.autoFilter.sortState.range, conditions: sheet.autoFilter.sortState.conditions.map((condition) => [condition.range, condition.descending]) } : null, held: (sheet.autoFilter.held || []).slice().sort((a, b) => a - b) } : null;
  return {
    name: sheet.name,
    kind: sheet.kind,
    links: sheet.hyperlinks.map((link) => [link.range, link.location || null, link.target || null, link.display || null]).sort((a, b) => (stable(a) < stable(b) ? -1 : 1)),
    drawings: (sheet.drawings || []).map((drawing) => drawing.object.type),
    state: sheet.state,
    tab: sheet.tabColor || null,
    rows,
    columns: mergedColumns(columns),
    merges: sheet.merges.map((m) => [m.top, m.left, m.bottom, m.right]).sort((a, b) => a[0] - b[0] || a[1] - b[1]),
    notes,
    validations,
    conditional,
    filter,
    charts: sheet.kind === 'worksheet' ? sheetCharts(wb, index).map((chart) => ({ ...chart, ofFile: undefined, editable: undefined })) : [],
  };
}

/** Runs of columns with the same settings joined: how they are cut into runs says nothing. */
function mergedColumns(columns: unknown[]): unknown[] {
  const out: unknown[][] = [];
  for (const item of columns as unknown[][]) {
    const last = out[out.length - 1];
    if (last && last[1] === (item[0] as number) - 1 && stable(last.slice(2)) === stable(item.slice(2))) last[1] = item[1];
    else out.push(item.slice());
  }
  return out;
}

/** What a workbook holds, as a text that is the same for the same content. */
export function canon(wb: Workbook): string {
  return stable({
    sheets: wb.sheets.map((sheet, index) => sheetText(wb, sheet, index)),
    names: wb.definedNames.map((name) => [name.name, name.value, name.localSheet === undefined ? null : name.localSheet, name.hidden]).sort((a, b) => (stable(a) < stable(b) ? -1 : 1)),
  });
}

/** Where two texts first differ, with what stands around it: for a failure that can be read. */
export function firstDifference(a: string, b: string): string {
  let at = 0;
  while (at < a.length && at < b.length && a[at] === b[at]) at++;
  return `at ${at}: …${a.slice(Math.max(0, at - 120), at + 120)}… against …${b.slice(Math.max(0, at - 120), at + 120)}…`;
}

export function assertSame(members: Member[], label: string): void {
  const texts = members.map((member) => canon(member.wb));
  for (let i = 1; i < texts.length; i++) {
    if (texts[i] !== texts[0]) assert.fail(`${label}: ${members[i].name} differs from ${members[0].name} ${firstDifference(texts[0], texts[i])}`);
  }
}

/**
 * Delivers every message and waits for every file that is being loaded again, until nothing
 * is on its way and nothing is loading: without guessing how many turns a load takes.
 */
export async function settleAll(net: Net, random: () => number, members: Member[]): Promise<void> {
  for (let round = 0; round < 200; round++) {
    net.drain(random);
    const loads = members.map((member) => member.loading).filter((load): load is Promise<void> => load !== undefined);
    for (const member of members) member.loading = undefined;
    if (loads.length > 0) {
      await Promise.all(loads);
      continue;
    }
    await new Promise<void>((resolve) => setImmediate(resolve));
    if (net.waiting === 0 && !members.some((member) => member.loading !== undefined)) return;
  }
  assert.fail('The loads never ran out');
}
