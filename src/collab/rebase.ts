// Rebasing: an operation that was made beside another one is read through that one before it
// is applied after it. Positions follow rows, columns and sheets that were inserted, deleted
// or moved; formulas follow them and the names of sheets; an operation whose place is gone is
// dropped, and whoever made it is told.
//
// None of this decides what the participants end up with: an operation is only ever applied in
// the one order the relay gave, on the workbook its author applied it to (see session.ts). A
// rule here that is not the best one moves a change to a place its author did not mean.

import { adjustFormula, adjustRange, renameSheetsInFormula, type StructuralEdit } from '../calc/refs';
import { MAX_COLS, MAX_ROWS, foldSheetName, type CellRange, type SheetNameMap } from '../model/sheet/types';
import { shiftAutoFilter } from '../ops/ops';
import type { Rec, Shift, WireCellPatch, WireColumn, WireNote, WireOp, WireRange, WireRowProps } from './types';
import { commandFamily, shiftOf } from './wire';

/** The name of no sheet: the home of a formula on a sheet other than the one that was edited. */
const NO_SHEET = '\u0000';

/** What reading an operation through a shift gave. */
export interface Mapped {
  /** The operation as it is to be applied now; missing when nothing is left to do or it cannot be done. */
  rec?: Rec;
  /** In place of `rec`: the operations it became, each to be applied after the one before it (rows that a sort took apart are deleted run by run). */
  recs?: Rec[];
  /** Nothing is left to do, and nothing was lost: what it wanted is so already. */
  noop?: true;
  /** Its place is gone. */
  conflict?: true;
  /** Part of it went with a place that is gone. */
  partial?: true;
  /** Rows or columns that are put in stand behind this many of those the shift put in at the same place (and not before them all, as they do otherwise). */
  behind?: number;
}

function mapSheet(s: number, shift: Shift): number | undefined {
  if (shift.k === 'sIns') return s >= shift.at ? s + 1 : s;
  if (shift.k === 'sDel') return s === shift.at ? undefined : s > shift.at ? s - 1 : s;
  if (shift.k === 'sMove') {
    if (s === shift.from) return shift.to;
    const without = s > shift.from ? s - 1 : s;
    return without >= shift.to ? without + 1 : without;
  }
  return s;
}

type RowCol = Extract<Shift, { k: 'rc' }>;

const editOf = (shift: RowCol): StructuralEdit => ({ kind: shift.ins ? 'insert' : 'delete', axis: shift.axis, at: shift.at, count: shift.n, sheetName: shift.name });

function mapPoint(p: number, shift: RowCol): number | undefined {
  if (shift.ins) return p >= shift.at ? p + shift.n : p;
  if (p < shift.at) return p;
  return p >= shift.at + shift.n ? p - shift.n : undefined;
}

/** A place rows, columns or a sheet are put in at. `first`: what is put in comes before what the shift put in at the same place. */
function mapInsertion(at: number, shift: { ins: boolean; at: number; n: number }, first: boolean): number {
  if (shift.ins) return at > shift.at || (at === shift.at && first) ? at + shift.n : at;
  if (at <= shift.at) return at;
  return at >= shift.at + shift.n ? at - shift.n : shift.at;
}

/** A run of rows or columns that goes, read through a shift on the same axis; undefined when nothing of it is left. */
function mapBand(at: number, n: number, shift: RowCol): [number, number] | undefined {
  if (shift.ins) {
    if (shift.at <= at) return [at + shift.n, n];
    // Rows put in inside the run go with it.
    return shift.at >= at + n ? [at, n] : [at, n + shift.n];
  }
  const end = at + n;
  const gone = Math.max(0, Math.min(end, shift.at + shift.n) - Math.max(at, shift.at));
  if (gone >= n) return undefined;
  const start = at <= shift.at ? at : at >= shift.at + shift.n ? at - shift.n : shift.at;
  return [start, n - gone];
}

const toRange = (range: WireRange): CellRange => ({ top: range[0], left: range[1], bottom: range[2], right: range[3] });
const fromRange = (range: CellRange): WireRange => [range.top, range.left, range.bottom, range.right];

function mapWireRange(range: WireRange, shift: RowCol): WireRange | undefined {
  const moved = adjustRange(toRange(range), editOf(shift));
  return moved ? fromRange(moved) : undefined;
}

/** A formula read through a shift. `same`: the formula is on the sheet the rows or columns moved in. */
function mapFormula(text: string, same: boolean, shift: Shift): string {
  if (shift.k === 'rc') return adjustFormula(text, same ? shift.name : NO_SHEET, editOf(shift));
  if (shift.k === 'sName' || shift.k === 'sDel') {
    if (text.indexOf('!') < 0) return text;
    const old = foldSheetName(shift.k === 'sName' ? shift.old : shift.name);
    const to = shift.k === 'sName' ? shift.name : null;
    const map: SheetNameMap = (written) => (foldSheetName(written) === old ? to : undefined);
    return renameSheetsInFormula(text, map);
  }
  return text;
}

function mapPatchFormula(patch: WireCellPatch, same: boolean, shift: Shift): WireCellPatch {
  if (patch === 0 || !patch.v || !patch.v.f) return patch;
  const text = mapFormula(patch.v.f.x, same, shift);
  return text === patch.v.f.x ? patch : { ...patch, v: { ...patch.v, f: { ...patch.v.f, x: text } } };
}

/** How often a formula that is written names something that is gone. */
function brokenRefs(patch: WireCellPatch): number {
  if (patch === 0 || !patch.v || !patch.v.f) return 0;
  return patch.v.f.x.split('#REF!').length - 1;
}

/** Whether a shift changes what formulas say. */
const movesFormulas = (shift: Shift): boolean => shift.k === 'rc' || shift.k === 'sName' || shift.k === 'sDel';

/** The commands that name a rule or a chart by its place in a list. */
const BY_PLACE = ['updateConditionalRule', 'removeConditionalRule', 'moveConditionalRule', 'updateChart', 'moveChart', 'removeChart', 'pasteChart'];

function mapModelRange(value: unknown, shift: RowCol): CellRange | undefined {
  const range = value as CellRange;
  if (!range || typeof range !== 'object' || typeof range.top !== 'number' || typeof range.left !== 'number' || typeof range.bottom !== 'number' || typeof range.right !== 'number') return undefined;
  return adjustRange(range, editOf(shift));
}

/** The ranges a command is about (one, or a list), read through a shift; undefined when none is left. */
function mapRanges(value: unknown, shift: RowCol): unknown {
  if (Array.isArray(value)) {
    const left = value.map((range) => mapModelRange(range, shift)).filter((range) => range !== undefined);
    return left.length > 0 ? left : undefined;
  }
  return mapModelRange(value, shift);
}

function mapCommand(op: Extract<WireOp, { t: 'cmd' }>, shift: Shift): Mapped {
  const family = commandFamily(op.name);
  const same = shift.k === 'rc' && shift.s === op.s;
  if (family === 'chart') {
    // What a chart shows and where it stands are references and anchors that are not read through here.
    return op.name === 'removeChart' || shift.k === 'perm' ? { rec: { w: op } } : { conflict: true };
  }
  if (BY_PLACE.indexOf(op.name) >= 0) {
    // Rows or columns that go may take a rule with them, and the places of the others change.
    return same ? { conflict: true } : { rec: { w: op } };
  }
  const args = op.a.slice();
  if (same && shift.k === 'rc' && args[0] !== null && args[0] !== undefined) {
    const ranges = mapRanges(args[0], shift);
    if (ranges === undefined) return { conflict: true };
    args[0] = ranges;
  }
  const settings = args[1];
  if (settings && typeof settings === 'object' && !Array.isArray(settings) && (op.name === 'setValidation' || op.name === 'addConditionalRule')) {
    const next = { ...(settings as Record<string, unknown>) };
    for (const key of ['formula1', 'formula2']) if (typeof next[key] === 'string') next[key] = mapFormula(next[key] as string, same, shift);
    if (Array.isArray(next.formulas)) next.formulas = next.formulas.map((text) => (typeof text === 'string' ? mapFormula(text, same, shift) : text));
    args[1] = next;
  }
  return { rec: { w: { ...op, a: args } } };
}

/**
 * Reads an operation through a shift that was applied before it. `first`: the operation
 * comes before the shift in the order they end up in (it is read through a later one only to
 * know where it is from then on), which decides who stands first when both put something in
 * at one place.
 */
export function mapRec(rec: Rec, shift: Shift, first = false): Mapped {
  const op = rec.w;
  // The sheet the operation is about.
  let s = op.s;
  if (op.t === 'sheet+') {
    if (shift.k === 'sIns') s = mapInsertion(op.s, { ins: true, at: shift.at, n: 1 }, first);
    else if (shift.k === 'sDel') s = op.s > shift.at ? op.s - 1 : op.s;
    return { rec: { w: s === op.s ? op : { ...op, s }, nm: rec.nm } };
  }
  if (shift.k === 'sIns' || shift.k === 'sDel' || shift.k === 'sMove') {
    const mapped = mapSheet(op.s, shift);
    if (mapped === undefined) return op.t === 'sheet-' ? { noop: true } : { conflict: true };
    s = mapped;
  }
  const moved = (next: WireOp, nm = rec.nm): Mapped => ({ rec: rec.wide ? { w: next, nm, wide: true } : { w: next, nm } });
  const onSheet = shift.k === 'rc' && shift.s === op.s;

  switch (op.t) {
    case 'cells': {
      const perm = shift.k === 'perm' && shift.s === op.s ? shift : undefined;
      if (!onSheet && !perm && !movesFormulas(shift)) return moved(s === op.s ? op : { ...op, s });
      const cells: [number, number, WireCellPatch][] = [];
      let partial = false;
      for (const entry of op.c) {
        let row = entry[0];
        let col = entry[1];
        if (onSheet && shift.k === 'rc') {
          const at = mapPoint(shift.axis === 'row' ? row : col, shift);
          if (at === undefined) {
            partial = true;
            continue;
          }
          if (shift.axis === 'row') row = at;
          else col = at;
        } else if (perm && row >= perm.top && row <= perm.bottom && col >= perm.left && col <= perm.right) row = perm.top + perm.target[row - perm.top];
        const patch = movesFormulas(shift) ? mapPatchFormula(entry[2], onSheet, shift) : entry[2];
        // A formula that read what someone else took away meanwhile is written as the error it has become: said.
        if (patch !== entry[2] && brokenRefs(patch) > brokenRefs(entry[2])) partial = true;
        cells.push(row === entry[0] && col === entry[1] && patch === entry[2] ? entry : [row, col, patch]);
      }
      if (cells.length === 0) return { conflict: true };
      const out = moved({ ...op, s, c: cells });
      if (partial) out.partial = true;
      return out;
    }
    case 'notes': {
      const perm = shift.k === 'perm' && shift.s === op.s ? shift : undefined;
      if (!onSheet && !perm) return moved(s === op.s ? op : { ...op, s });
      const notes: [number, number, WireNote | 0][] = [];
      let partial = false;
      for (const entry of op.n) {
        let row = entry[0];
        let col = entry[1];
        if (onSheet && shift.k === 'rc') {
          const at = mapPoint(shift.axis === 'row' ? row : col, shift);
          if (at === undefined) {
            partial = true;
            continue;
          }
          if (shift.axis === 'row') row = at;
          else col = at;
        } else if (perm && row >= perm.top && row <= perm.bottom && col >= perm.left && col <= perm.right) row = perm.top + perm.target[row - perm.top];
        notes.push([row, col, entry[2]]);
      }
      if (notes.length === 0) return { conflict: true };
      const out = moved({ ...op, s, n: notes });
      if (partial) out.partial = true;
      return out;
    }
    case 'rows': {
      if (!onSheet || shift.k !== 'rc' || shift.axis !== 'row') return moved(s === op.s ? op : { ...op, s });
      const rows: [number, WireRowProps][] = [];
      for (const entry of op.r) {
        const at = mapPoint(entry[0], shift);
        if (at !== undefined) rows.push([at, entry[1]]);
      }
      if (rows.length === 0) return { conflict: true };
      const out = moved({ ...op, s, r: rows });
      if (rows.length < op.r.length) out.partial = true;
      return out;
    }
    case 'cols': {
      if (!onSheet || shift.k !== 'rc' || shift.axis !== 'col') return moved(s === op.s ? op : { ...op, s });
      const columns: [number, number, WireColumn | 0][] = [];
      for (const entry of op.p) {
        const range = adjustRange({ top: 0, bottom: 0, left: entry[0], right: entry[1] }, editOf(shift));
        if (range) columns.push([range.left, Math.min(range.right, MAX_COLS - 1), entry[2]]);
      }
      if (columns.length === 0) return { conflict: true };
      const out = moved({ ...op, s, p: columns });
      if (columns.length < op.p.length) out.partial = true;
      return out;
    }
    case 'merges': {
      if (!onSheet || shift.k !== 'rc') return moved(s === op.s ? op : { ...op, s });
      const through = (list: WireRange[]): WireRange[] => {
        const out: WireRange[] = [];
        for (const range of list) {
          const next = mapWireRange(range, shift);
          if (next && (next[0] !== next[2] || next[1] !== next[3])) out.push(next);
        }
        return out;
      };
      const add = through(op.add);
      const del = through(op.del);
      if (add.length === 0 && del.length === 0) return op.add.length > 0 ? { conflict: true } : { noop: true };
      const out = moved({ ...op, s, add, del });
      if (add.length < op.add.length) out.partial = true;
      return out;
    }
    case 'filter': {
      if (!onSheet || shift.k !== 'rc') return moved(s === op.s ? op : { ...op, s });
      let f = op.f;
      if (f !== 0) {
        const next = shiftAutoFilter(f, { kind: shift.ins ? 'insert' : 'delete', axis: shift.axis, at: shift.at, count: shift.n });
        if (!next) return { conflict: true };
        f = next;
      }
      const rows = (list: number[] | undefined): number[] | undefined => {
        if (!list || shift.axis !== 'row') return list;
        const out: number[] = [];
        for (const row of list) {
          const at = mapPoint(row, shift);
          if (at !== undefined) out.push(at);
        }
        return out;
      };
      const next: WireOp = { t: 'filter', s, f };
      const show = rows(op.show);
      const hide = rows(op.hide);
      if (show && show.length > 0) next.show = show;
      if (hide && hide.length > 0) next.hide = hide;
      return moved(next);
    }
    case 'sort': {
      if (shift.k === 'perm' && shift.s === op.s) {
        const apart = op.g[2] < shift.top || op.g[0] > shift.bottom || op.g[3] < shift.left || op.g[1] > shift.right;
        return apart ? moved(op) : { conflict: true };
      }
      if (!onSheet || shift.k !== 'rc') return moved(s === op.s ? op : { ...op, s });
      if (shift.axis === 'col') {
        const range = mapWireRange(op.g, shift);
        return range ? moved({ ...op, s, g: range }) : { conflict: true };
      }
      // Rows above the range move it; rows that come or go inside it leave no order to keep.
      const end = shift.at + shift.n - 1;
      if (shift.at > op.g[2]) return moved(s === op.s ? op : { ...op, s });
      if (shift.ins ? shift.at <= op.g[0] : end < op.g[0]) {
        const by = shift.ins ? shift.n : -shift.n;
        if (op.g[2] + by >= MAX_ROWS) return { conflict: true };
        return moved({ ...op, s, g: [op.g[0] + by, op.g[1], op.g[2] + by, op.g[3]] });
      }
      return { conflict: true };
    }
    case 'ins':
    case 'del': {
      let nm = rec.nm;
      if (shift.k === 'sName' && shift.s === op.s) nm = shift.name;
      if (op.t === 'del' && op.x === 'row' && shift.k === 'perm' && shift.s === op.s && op.at <= shift.bottom && op.at + op.n - 1 >= shift.top) {
        // The rows that were to go stand elsewhere after the sort. Where only some columns were sorted there is
        // no saying which row is meant: what stood in one row stands in two.
        if (!shift.wide) return { conflict: true };
        const rows: number[] = [];
        for (let row = op.at; row < op.at + op.n; row++) rows.push(row >= shift.top && row <= shift.bottom ? shift.top + shift.target[row - shift.top] : row);
        rows.sort((a, b) => b - a);
        // Run by run, the lowest rows first: what is deleted further down moves nothing above it.
        const recs: Rec[] = [];
        for (let i = 0; i < rows.length; ) {
          let k = i;
          while (k + 1 < rows.length && rows[k + 1] === rows[k] - 1) k++;
          recs.push({ w: { ...op, s, at: rows[k], n: k - i + 1 }, nm });
          i = k + 1;
        }
        return recs.length === 1 ? { rec: recs[0] } : { recs };
      }
      if (!onSheet || shift.k !== 'rc' || shift.axis !== op.x) return moved(s === op.s ? op : { ...op, s }, nm);
      if (op.t === 'ins') {
        if (!shift.ins && op.at > shift.at && op.at <= shift.at + shift.n) {
          // Into a run that is gone, or right behind its last row: at the place of the run, and how far in is remembered.
          return { rec: { w: { ...op, s, at: shift.at }, nm, fold: op.at - shift.at } };
        }
        if (shift.ins && rec.fold !== undefined && shift.at === op.at && !first) {
          // Rows came (back) where the run went: as far behind the first of them as it was behind the first of the run.
          const behind = Math.min(rec.fold, shift.n);
          return { rec: { w: { ...op, s, at: op.at + behind }, nm }, behind };
        }
        const done = moved({ ...op, s, at: mapInsertion(op.at, shift, first) }, nm);
        if (rec.fold !== undefined && done.rec) done.rec = { ...done.rec, fold: rec.fold };
        return done;
      }
      // Rows of the run that someone else took away first, and that come back now (their deletion was undone): they
      // were to go and stay. Which is said.
      const back = rec.cut === true && shift.ins && shift.at >= op.at && shift.at <= op.at + op.n && !first;
      if (shift.ins && shift.at > op.at && shift.at < op.at + op.n && !first) {
        // Rows that someone else put in between the rows that are to go were never seen by whoever deletes: they
        // stay, and the rows on either side of them go - the lower ones first, which moves nothing above them.
        const split: Mapped = { recs: [{ w: { ...op, s, at: shift.at + shift.n, n: op.at + op.n - shift.at }, nm }, { w: { ...op, s, n: shift.at - op.at }, nm }] };
        if (back) split.partial = true;
        return split;
      }
      const band = mapBand(op.at, op.n, shift);
      if (!band) return { noop: true };
      const out = moved({ ...op, s, at: band[0], n: band[1] }, nm);
      if (back) out.partial = true;
      else if (out.rec && (rec.cut === true || band[1] < op.n)) out.rec = { ...out.rec, cut: true };
      return out;
    }
    case 'sheet-':
      return moved(s === op.s ? op : { ...op, s }, shift.k === 'sName' && shift.s === op.s ? shift.name : rec.nm);
    case 'name':
      return moved(s === op.s ? op : { ...op, s }, shift.k === 'sName' && shift.s === op.s ? shift.name : rec.nm);
    case 'state':
    case 'tab':
    case 'pane':
      return moved(s === op.s ? op : { ...op, s });
    case 'move': {
      let to = op.to;
      if (shift.k === 'sIns' && to >= shift.at) to++;
      else if (shift.k === 'sDel' && to > shift.at) to--;
      if (to === s) return { noop: true };
      return moved(s === op.s && to === op.to ? op : { ...op, s, to });
    }
    case 'cmd': {
      const mapped = movesFormulas(shift) || shift.k === 'perm' ? mapCommand(op, shift) : { rec: { w: op as WireOp } };
      if (!mapped.rec) return mapped;
      const next = mapped.rec.w;
      return moved(s === next.s ? next : { ...next, s });
    }
  }
}

/**
 * A sort read through rows that were inserted into or deleted from the sheet beside it: the
 * rows that are left go where the sort put them, counted in the sheet with those rows in it
 * (or without them). Undefined when nothing is left to say.
 */
function permThroughRows(perm: Extract<Shift, { k: 'perm' }>, over: RowCol): Shift | undefined {
  const top = perm.top;
  const length = perm.target.length;
  const bottom = top + length - 1;
  if (over.ins) {
    if (over.at <= top) return { ...perm, top: top + over.n, bottom: bottom + over.n };
    if (over.at > bottom) return perm;
    // The new rows stand where they were put, before the sort and after it.
    const at = over.at - top;
    const target: number[] = new Array(length + over.n);
    for (let i = 0; i < length + over.n; i++) {
      if (i >= at && i < at + over.n) target[i] = i;
      else {
        const to = perm.target[i < at ? i : i - over.n];
        target[i] = to < at ? to : to + over.n;
      }
    }
    return { ...perm, bottom: bottom + over.n, target };
  }
  const end = over.at + over.n - 1;
  if (end < top) return { ...perm, top: top - over.n, bottom: bottom - over.n };
  if (over.at > bottom) return perm;
  const first = Math.max(over.at, top) - top;
  const last = Math.min(end, bottom) - top;
  // Where the rows that went were put by the sort: they go from there.
  const gone: number[] = [];
  for (let i = first; i <= last; i++) gone.push(perm.target[i]);
  gone.sort((a, b) => a - b);
  const target: number[] = [];
  for (let i = 0; i < length; i++) {
    if (i >= first && i <= last) continue;
    const to = perm.target[i];
    let below = 0;
    while (below < gone.length && gone[below] < to) below++;
    target.push(to - below);
  }
  if (target.length === 0) return undefined;
  const above = Math.max(0, top - over.at);
  return { ...perm, top: top - above, bottom: top - above + target.length - 1, target };
}

/**
 * A shift read through another one that was applied beside it and comes after it in the order
 * they end up in: how it moves things as counted behind that other one. Several shifts when
 * it takes more than one to say (rows that go on either side of rows that were put in), none
 * when nothing of it is left.
 */
export function mapShift(shift: Shift, over: Shift): Shift[] {
  const one = (next: Shift | undefined): Shift[] => (next ? [next] : []);
  switch (shift.k) {
    case 'rc': {
      if (over.k === 'rc') {
        if (over.s !== shift.s || over.axis !== shift.axis) return [shift];
        if (shift.ins) {
          // Rows put in inside a run that the step deleted stay (the step goes around them): they stand where the run was.
          return [{ ...shift, at: mapInsertion(shift.at, over, true) }];
        }
        if (over.ins && over.at > shift.at && over.at < shift.at + shift.n) {
          // Rows that were put in inside the run that went stay: the rows below them went, and the rows above them.
          return [
            { ...shift, at: over.at + over.n, n: shift.at + shift.n - over.at },
            { ...shift, n: over.at - shift.at },
          ];
        }
        const band = mapBand(shift.at, shift.n, over);
        return band ? [{ ...shift, at: band[0], n: band[1] }] : [];
      }
      if (over.k === 'sName') return [over.s === shift.s ? { ...shift, name: over.name } : shift];
      if (over.k === 'perm') return [shift];
      const s = mapSheet(shift.s, over);
      return s === undefined ? [] : [s === shift.s ? shift : { ...shift, s }];
    }
    case 'perm': {
      if (over.k === 'rc') {
        if (over.s !== shift.s) return [shift];
        if (over.axis === 'col') {
          const range = adjustRange({ top: 0, bottom: 0, left: shift.left, right: shift.right }, editOf(over));
          return range ? [{ ...shift, left: range.left, right: range.right }] : [];
        }
        return one(permThroughRows(shift, over));
      }
      if (over.k === 'sName' || over.k === 'perm') return [shift];
      const s = mapSheet(shift.s, over);
      return s === undefined ? [] : [s === shift.s ? shift : { ...shift, s }];
    }
    case 'sIns':
      if (over.k === 'sIns') return [{ ...shift, at: mapInsertion(shift.at, { ins: true, at: over.at, n: 1 }, true) }];
      if (over.k === 'sDel') return [shift.at > over.at ? { ...shift, at: shift.at - 1 } : shift];
      return [shift];
    case 'sDel': {
      if (over.k === 'sName') return [over.s === shift.at ? { ...shift, name: over.name } : shift];
      if (over.k === 'rc' || over.k === 'perm') return [shift];
      const at = mapSheet(shift.at, over);
      return at === undefined ? [] : [at === shift.at ? shift : { ...shift, at }];
    }
    case 'sMove': {
      if (over.k === 'rc' || over.k === 'perm' || over.k === 'sName') return [shift];
      const from = mapSheet(shift.from, over);
      if (from === undefined) return [];
      let to = shift.to;
      if (over.k === 'sIns' && to >= over.at) to++;
      else if (over.k === 'sDel' && to > over.at) to--;
      return from === to ? [] : [{ ...shift, from, to }];
    }
    case 'sName': {
      if (over.k === 'rc' || over.k === 'perm') return [shift];
      // The sheet got another name since: that one stands.
      if (over.k === 'sName') return over.s === shift.s ? [] : [shift];
      const s = mapSheet(shift.s, over);
      return s === undefined ? [] : [s === shift.s ? shift : { ...shift, s }];
    }
  }
}

/** The shift that takes a shift back. */
export function inverseShift(shift: Shift): Shift {
  switch (shift.k) {
    case 'rc':
      return { ...shift, ins: !shift.ins };
    case 'perm': {
      const target: number[] = new Array(shift.target.length);
      for (let i = 0; i < shift.target.length; i++) target[shift.target[i]] = i;
      return { ...shift, target };
    }
    case 'sIns':
      return { k: 'sDel', at: shift.at, name: shift.name };
    case 'sDel':
      return { k: 'sIns', at: shift.at, name: shift.name };
    case 'sMove':
      return { k: 'sMove', from: shift.to, to: shift.from };
    case 'sName':
      return { k: 'sName', s: shift.s, old: shift.name, name: shift.old };
  }
}

const key = (row: number, col: number): number => row * MAX_COLS + col;

/**
 * What an operation has to give way to in one that was applied before it and is about the
 * same thing. `undo`: the operation takes back an earlier change of whoever keeps it, and
 * gives way wherever someone else changed the same thing since.
 */
function clash(op: WireOp, other: WireOp, undo: boolean): { drop: true } | { op: WireOp } | undefined {
  // A chart that is copied is named by its place on a sheet that may be another one than the copy is put on:
  // whatever was done to the charts of any sheet first may have made that place another chart's.
  if (op.t === 'cmd' && op.name === 'pasteChart' && other.t === 'cmd' && commandFamily(other.name) === 'chart') return { drop: true };
  if (op.s !== other.s) return undefined;
  if (op.t === 'filter' && other.t === 'filter') return { drop: true };
  // Cells that are merged and something that is written into them, either way round: what a merge empties is what
  // its author saw, and what is written into a merged range may be hidden under it. No rule is known to be right.
  const merged = op.t === 'merges' ? op : other.t === 'merges' ? other : undefined;
  const written = op.t === 'cells' || op.t === 'notes' ? op : other.t === 'cells' || other.t === 'notes' ? other : undefined;
  if (merged && written) {
    // A note of a merged cell is the note of its first cell: where a note belongs depends on every merged range of its sheet.
    if (written.t === 'notes') return { drop: true };
    const places: [number, number][] = written.c.map((entry) => [entry[0], entry[1]]);
    for (const range of merged.add) {
      for (const [row, col] of places) if (row >= range[0] && row <= range[2] && col >= range[1] && col <= range[3]) return { drop: true };
    }
  }
  if (op.t === 'sort' && other.t === 'sort') {
    const apart = op.g[2] < other.g[0] || op.g[0] > other.g[2] || op.g[3] < other.g[1] || op.g[1] > other.g[3];
    return apart ? undefined : { drop: true };
  }
  if (op.t === 'cmd' && other.t === 'cmd') {
    return BY_PLACE.indexOf(op.name) >= 0 && commandFamily(op.name) === commandFamily(other.name) ? { drop: true } : undefined;
  }
  if (!undo) return undefined;
  if (op.t === 'cells' && other.t === 'cells') {
    const taken = new Set<number>();
    for (const entry of other.c) taken.add(key(entry[0], entry[1]));
    const cells = op.c.filter((entry) => !taken.has(key(entry[0], entry[1])));
    if (cells.length === op.c.length) return undefined;
    return cells.length > 0 ? { op: { ...op, c: cells } } : { drop: true };
  }
  if (op.t === 'notes' && other.t === 'notes') {
    const taken = new Set<number>();
    for (const entry of other.n) taken.add(key(entry[0], entry[1]));
    const notes = op.n.filter((entry) => !taken.has(key(entry[0], entry[1])));
    if (notes.length === op.n.length) return undefined;
    return notes.length > 0 ? { op: { ...op, n: notes } } : { drop: true };
  }
  if ((op.t === 'name' || op.t === 'tab' || op.t === 'state' || op.t === 'pane') && other.t === op.t) return { drop: true };
  return undefined;
}

/** What was applied beside the operations that are rebased: how it moves things, and what it changed. */
export interface Carry {
  sh?: Shift;
  w?: WireOp;
}

export function carryOf(rec: Rec): Carry {
  return { sh: shiftOf(rec.w, rec.nm, rec.wide), w: rec.w };
}

export interface Transformed {
  /** The operations as they are to be applied after the carry; missing when the step is dropped. */
  recs?: Rec[];
  /** The carry as it is to be read after the step. */
  carry: Carry[];
  /** Part of the step went with places that are gone. */
  partial: boolean;
  /** The step is dropped and nothing was lost: what it wanted is so already. */
  noop: boolean;
  /**
   * The step is dropped, and with it everything that was made after it: where those later steps
   * belong cannot be told any more (a sheet was moved, or rows and a sort of one sheet met).
   */
  poison?: boolean;
}

/**
 * Reads the operations of one step (each made after the one before it) through a carry (each
 * applied after the one before it, the first on what the step was made on), and the carry
 * through the step. A step is kept whole or dropped whole.
 */
export function transformStep(recs: readonly Rec[], carry: readonly Carry[], undo = false): Transformed {
  let partial = false;
  let dropped = false;
  let poison = false;
  /**
   * The meetings for which no rule is known to be right: a sheet that is moved beside anything that
   * names a sheet by its place, and rows or columns that come or go on a sheet that is sorted beside
   * them. The step is dropped, and what was made after it as well.
   */
  const unruled = (mine: Shift | undefined, other: Shift | undefined): boolean => {
    if ((mine && mine.k === 'sMove') || (other && other.k === 'sMove')) return true;
    if (!mine || !other) return false;
    if (mine.k === 'rc' && other.k === 'perm') return mine.s === other.s;
    if (mine.k === 'perm' && other.k === 'rc') return mine.s === other.s;
    return false;
  };
  /** One operation through a list of carried items; the items come back as they are to be read behind it. */
  const pass = (rec: Rec, items: readonly Carry[]): { recs: Rec[]; items: Carry[] } => {
    let cur = rec;
    const behind: Carry[] = [];
    for (let k = 0; k < items.length; k++) {
      const item = items[k];
      if (item.w) {
        const hit = clash(cur.w, item.w, undo);
        if (hit && 'drop' in hit) {
          dropped = true;
          return { recs: [], items: [] };
        }
        if (hit) {
          cur = { ...cur, w: hit.op };
          partial = true;
        }
      }
      const mine = shiftOf(cur.w, cur.nm, cur.wide);
      if (unruled(mine, item.sh)) {
        dropped = true;
        poison = true;
        return { recs: [], items: [] };
      }
      // Taking rows or columns back (or bringing them back) after someone else put rows or columns into that sheet or
      // took some away: where they belong, and what belongs into them, is not known to come out right. Not undone.
      if (undo && mine && mine.k === 'rc' && item.sh && item.sh.k === 'rc' && item.sh.s === mine.s) {
        dropped = true;
        return { recs: [], items: [] };
      }
      // Merged ranges and rows or columns that come or go on their sheet, either way round: dropped as well.
      if (cur.w.t === 'merges' && item.sh && item.sh.k === 'rc' && item.sh.s === cur.w.s) {
        dropped = true;
        return { recs: [], items: [] };
      }
      if (mine && mine.k === 'rc' && item.w && item.w.t === 'merges' && item.w.s === mine.s) {
        dropped = true;
        poison = true;
        return { recs: [], items: [] };
      }
      let result: Mapped = { rec: cur };
      if (item.sh) {
        result = mapRec(cur, item.sh, false);
        if (result.conflict) {
          dropped = true;
          return { recs: [], items: [] };
        }
        if (result.partial) partial = true;
      }
      // The carried item is read through the operation as it was beside it.
      if (mine) {
        // Put in behind (some of) what the item put in: the item is read through it where it stands now, and is
        // two insertions when it stands between them.
        const through = result.behind !== undefined && result.rec ? shiftOf(result.rec.w, result.rec.nm, result.rec.wide) || mine : mine;
        const w = item.w ? mapRec({ w: item.w }, through, true).rec : undefined;
        let shifts = item.sh ? mapShift(item.sh, through) : [];
        if (result.behind !== undefined && item.sh && item.sh.k === 'rc' && through.k === 'rc') {
          const k = result.behind;
          shifts = k >= item.sh.n ? [item.sh] : [{ ...item.sh, n: k }, { ...item.sh, at: item.sh.at + k + through.n, n: item.sh.n - k }];
        }
        if (shifts.length === 0 && w) behind.push({ w: w.w });
        shifts.forEach((sh, i) => behind.push(i === 0 && w ? { sh, w: w.w } : { sh }));
      } else behind.push(item);
      if (result.recs) {
        // It became several: each goes through what is left of the carry, one behind the other.
        let rest: Carry[] = items.slice(k + 1);
        const out: Rec[] = [];
        for (const part of result.recs) {
          const done = pass(part, rest);
          if (dropped) return { recs: [], items: [] };
          for (const made of done.recs) out.push(made);
          rest = done.items;
        }
        return { recs: out, items: behind.concat(rest) };
      }
      if (!result.rec) {
        // Nothing is left of it to do: the rest of the carry stands as it is.
        return { recs: [], items: behind.concat(items.slice(k + 1)) };
      }
      cur = result.rec;
    }
    return { recs: [cur], items: behind };
  };
  let over: Carry[] = carry.slice();
  const out: Rec[] = [];
  for (const rec of recs) {
    const done = pass(rec, over);
    if (dropped) return { carry: carry.slice(), partial: false, noop: false, poison };
    for (const made of done.recs) out.push(made);
    over = done.items;
  }
  if (out.length === 0) return { carry: over, partial: false, noop: true };
  return { recs: out, carry: over, partial, noop: false };
}

/**
 * A position read through a shift: for the selections of participants. Undefined when its
 * sheet is gone; a cell that went gives the one that took its place, or with `strict` undefined.
 */
export function mapPosition(at: { s: number; row: number; col: number }, shift: Shift, strict = false): { s: number; row: number; col: number } | undefined {
  const s = mapSheet(at.s, shift);
  if (s === undefined) return undefined;
  if (shift.k === 'rc' && shift.s === at.s) {
    const p = mapPoint(shift.axis === 'row' ? at.row : at.col, shift);
    if (p === undefined && strict) return undefined;
    const clamped = p === undefined ? Math.min(shift.at, (shift.axis === 'row' ? MAX_ROWS : MAX_COLS) - 1) : p;
    return shift.axis === 'row' ? { s, row: clamped, col: at.col } : { s, row: at.row, col: clamped };
  }
  return { s, row: at.row, col: at.col };
}
