// What collaborative editing sends for a change of a workbook, and what it keeps of one.
//
// An operation of the model (see ../ops/ops.ts) cannot travel as it is: it holds cells with
// indexes into the style tables of one participant, lists by reference, and for rules and
// drawings things of the file that participant read (another participant may have read a
// later version of the file). A wire operation says the same change in what every
// participant has: positions as they are now, formats written out, and for rules and charts
// the command that was run, which each participant runs on its own workbook.

import type { AutoFilter, CellKind, Color, RichRun, SheetState } from '../model/sheet/types';

/** top, left, bottom, right. */
export type WireRange = [number, number, number, number];

export interface WireFormula {
  /** The text, without "=". A shared formula travels written out for its cell. */
  x: string;
  t?: 'array' | 'dataTable';
  /** The range of an array formula. */
  r?: string;
  /** Attributes of <f> the model has no field for. */
  a?: [string, string][];
  /** A dynamic array formula: its array result spills (where to is worked out by every participant). */
  d?: 1;
}

/** What a cell holds beside its format. */
export interface WireContent {
  k: CellKind;
  v: number | string | boolean | null;
  f?: WireFormula;
  r?: RichRun[];
  ph?: true;
}

/**
 * A change of one cell: 0 takes the cell away; otherwise `v` is what it holds from now on and
 * `s` its format, and whichever is left out stays as it is. In memory `s` is an index into
 * the cell formats of the workbook; on the wire it is an index into the formats the message
 * brings along (see encodeStep()).
 */
export type WireCellPatch = 0 | { v?: WireContent; s?: number };

/** Properties of a row; null takes a height or a format away. Only what is named changes. */
export interface WireRowProps {
  h?: number | null;
  ch?: boolean;
  hid?: boolean;
  s?: number | null;
}

/** The settings of a run of columns, all of them. */
export interface WireColumn {
  w?: number;
  cw: boolean;
  hid: boolean;
  s?: number;
  ol: number;
  col: boolean;
}

export interface WireNote {
  /** The text. */
  x: string;
  /** The author. */
  a?: string;
  r?: RichRun[];
  /** The box is always shown. */
  vis?: true;
}

export type WireOp =
  | { t: 'cells'; s: number; c: [number, number, WireCellPatch][] }
  | { t: 'rows'; s: number; r: [number, WireRowProps][] }
  /** Runs of columns (first, last) and the settings they get; 0: none. */
  | { t: 'cols'; s: number; p: [number, number, WireColumn | 0][] }
  | { t: 'merges'; s: number; add: WireRange[]; del: WireRange[] }
  /** The filter of a sheet (0: none) with the rows it shows and hides, worked out by whoever sent it. */
  | { t: 'filter'; s: number; f: AutoFilter | 0; show?: number[]; hide?: number[] }
  /** Row `g.top + i` takes the cells of row `g.top + o[i]`. */
  | { t: 'sort'; s: number; g: WireRange; o: number[] }
  | { t: 'ins' | 'del'; s: number; x: 'row' | 'col'; at: number; n: number }
  /** Notes by the cells they are on now; 0 takes the note of a cell away. */
  | { t: 'notes'; s: number; n: [number, number, WireNote | 0][] }
  | { t: 'sheet+'; s: number; name: string }
  | { t: 'sheet-'; s: number }
  | { t: 'name'; s: number; name: string }
  | { t: 'move'; s: number; to: number }
  | { t: 'state'; s: number; state: SheetState }
  | { t: 'tab'; s: number; color: Color | 0 }
  /**
   * The frozen panes of a sheet: rows and columns, the first cell that scrolls and the first cell of the view (as
   * SheetView has them). The first cell that scrolls is below and right of the frozen ones, or the step is no step;
   * without it, it is the one right below and beside them (see paneOfWire()).
   */
  | { t: 'pane'; s: number; r: number; c: number; tl?: string; fv?: string }
  /** A command on the rules or the charts of a sheet, with what it was called with after the workbook and the sheet. */
  | { t: 'cmd'; s: number; name: string; a: unknown[] };

/**
 * How an operation moves what stands around it: the positions and the names that operations
 * made beside it have to be read through it.
 */
export type Shift =
  | { k: 'rc'; s: number; axis: 'row' | 'col'; ins: boolean; at: number; n: number; name: string }
  /** A sort: row `range.top + i` goes to row `range.top + target[i]`. `wide`: every used column of the sheet was sorted, so that whole rows moved. */
  | { k: 'perm'; s: number; top: number; left: number; bottom: number; right: number; target: number[]; wide: boolean }
  | { k: 'sIns'; at: number; name: string }
  | { k: 'sDel'; at: number; name: string }
  | { k: 'sMove'; from: number; to: number }
  | { k: 'sName'; s: number; old: string; name: string };

/**
 * A wire operation with the name it needs to be read by others: the name of its sheet for rows
 * and columns that move and for a sheet that goes, the name the sheet had for a rename.
 */
export interface Rec {
  w: WireOp;
  nm?: string;
  /** A sort that took every used column of its sheet along (see Shift). */
  wide?: boolean;
  /**
   * For rows or columns that are put in at a place where a run was taken away by someone else: how far
   * into that run they were put in. When rows come back at that place (the deletion was undone), they are
   * put in as far behind the first of them.
   */
  fold?: number;
  /** For rows or columns that are taken away: someone else took some of them away first. When rows come back there, not all that was to go is gone. */
  cut?: boolean;
}

/** A message or an operation that is not what the contract says: whoever sent it is not trusted further. */
export class WireError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'WireError';
  }
}

/** Something that cannot be done while a workbook is edited together. `what` names it for the message of the view. */
export class CollabUnsupportedError extends RangeError {
  readonly what: 'copySheet' | 'undo' | 'size' | 'readOnly' | 'picture' | 'other';
  constructor(what: CollabUnsupportedError['what'], message: string) {
    super(message);
    this.name = 'CollabUnsupportedError';
    this.what = what;
  }
}
