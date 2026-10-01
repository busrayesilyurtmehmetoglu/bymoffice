// Cell references inside formula text: "A1", "$B$2:C3", "A:C", "1:3", "'My Sheet'!A1", ...
//
// Nothing is evaluated here. A formula is scanned once, left to right, and only the
// references it contains are rewritten; every other character is copied as it stands, and a
// reference that does not change keeps its own spelling ("$" signs and letter case included).
// The scanner never throws and never backtracks over more than one token: text it cannot make
// sense of (an unclosed quote or bracket) is returned untouched from that point on.

import { MAX_COLS, MAX_ROWS, columnName, type CellRange } from '../model/sheet/types';

export interface StructuralEdit {
  kind: 'insert' | 'delete';
  axis: 'row' | 'col';
  /** First affected index, 0-based. */
  at: number;
  count: number;
  /** Name of the sheet the rows or columns were inserted into or deleted from. */
  sheetName: string;
}

const REF_ERROR = '#REF!';

/** Error constants, upper case. A "#" that starts none of these is copied as one character ("A1#"). */
const ERROR_LITERALS = [
  REF_ERROR,
  '#DIV/0!',
  '#VALUE!',
  '#NAME?',
  '#NUM!',
  '#NULL!',
  '#N/A',
  '#SPILL!',
  '#CALC!',
  '#GETTING_DATA',
  '#FIELD!',
  '#BLOCKED!',
  '#CONNECT!',
  '#UNKNOWN!',
  '#BUSY!',
  '#PYTHON!',
];
const MAX_ERROR_LENGTH = 13;

// What a word (a run of name characters) spells.
const K_CELL = 1;
const K_COL = 2;
const K_ROW = 3;

/** One side of a reference: "$B$7", "$B" or "$7". Indexes are 0-based. */
interface Word {
  kind: number;
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

/** A cell, a range, a column range or a row range. A single cell repeats its corner in the second one. */
interface Ref {
  kind: number;
  /** Written with a colon. */
  pair: boolean;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  absR1: boolean;
  absC1: boolean;
  absR2: boolean;
  absC2: boolean;
  /** Sheet prefix written again before the second corner ("Sheet1!A1:Sheet1!B2"), kept as written. */
  mid?: string;
}

interface Body {
  /** Offset just past what was read; equal to the start when there was no word at all. */
  end: number;
  ref?: Ref;
}

/** A token that may be a reference, with its sheet prefix. */
interface Item {
  start: number;
  /** Offset of the reference itself, after the "!" of a prefix. */
  bodyStart: number;
  /** Always greater than `start`. */
  end: number;
  ref?: Ref;
  /** Sheet names of the prefix (two for a 3D reference); undefined without a prefix. */
  sheets?: string[];
  /** Refers to another workbook: never rewritten. */
  external: boolean;
}

/** Returns the new text of a reference, undefined to keep it as written. */
type Transform = (ref: Ref, sheets: string[] | undefined, external: boolean) => string | undefined;

// ---------------------------------------------------------------------------
// Scanning

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * Characters of names, references and numbers: letters, digits, "_", ".", "$", "\", "?" and
 * everything outside ASCII. A reference is only recognized when it is a whole word, which is
 * what keeps "E5" in "1E5" and "A1" in "A1B2C" from being taken for cells.
 */
function isWordChar(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  if (c >= 65) return c <= 90 || c === 95 || c === 92;
  if (c >= 48) return c <= 57 || c === 63;
  return c === 36 || c === 46;
}

function skipWord(s: string, i: number): number {
  const n = s.length;
  while (i < n && isWordChar(s.charCodeAt(i))) i++;
  return i;
}

/** Offset just past the closing quote of the literal that starts at `i` (a doubled quote is an escape), -1 if unclosed. */
function skipQuoted(s: string, i: number, quote: string): number {
  let from = i + 1;
  for (;;) {
    const k = s.indexOf(quote, from);
    if (k < 0) return -1;
    if (s[k + 1] !== quote) return k + 1;
    from = k + 2;
  }
}

/**
 * Offset just past the "]" matching the "[" at `i`, -1 if unclosed. Brackets nest in
 * structured references ("Table1[[#This Row],[Total]]"), where "'" escapes the next character.
 */
function skipBrackets(s: string, i: number): number {
  const n = s.length;
  let depth = 0;
  for (let j = i; j < n; j++) {
    const c = s.charCodeAt(j);
    if (c === 39) j++;
    else if (c === 91) depth++;
    else if (c === 93 && --depth === 0) return j + 1;
  }
  return -1;
}

/** Reads s[a..b) as "$B$7", "$B" or "$7"; undefined for anything else or outside the sheet limits. */
function readWord(s: string, a: number, b: number): Word | undefined {
  let i = a;
  let absCol = false;
  let absRow = false;
  if (s.charCodeAt(i) === 36) {
    absCol = true;
    i++;
  }
  let col = 0;
  const lettersStart = i;
  for (; i < b; i++) {
    const c = s.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    if (col > MAX_COLS) return undefined;
  }
  const hasLetters = i > lettersStart;
  if (!hasLetters) {
    // The "$" belonged to the row.
    absRow = absCol;
    absCol = false;
  } else if (s.charCodeAt(i) === 36) {
    absRow = true;
    i++;
  }
  let row = 0;
  const digitsStart = i;
  for (; i < b; i++) {
    const c = s.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    row = row * 10 + (c - 48);
    if (row > MAX_ROWS) return undefined;
  }
  const hasDigits = i > digitsStart;
  if (hasDigits) {
    if (row === 0) return undefined;
    return { kind: hasLetters ? K_CELL : K_ROW, row: row - 1, col: col - 1, absRow, absCol };
  }
  if (!hasLetters || absRow) return undefined;
  return { kind: K_COL, row: 0, col: col - 1, absRow: false, absCol };
}

/** "(" makes a function call, "[" a table, "!" a sheet prefix: the word before them is not a reference. */
const endsName = (c: number) => c === 40 || c === 91 || c === 33;

/** Reads a reference without its sheet prefix: "A1", "A1:B2", "A:C" or "1:3". */
function readBody(s: string, p: number): Body {
  const e1 = skipWord(s, p);
  if (e1 === p) return { end: p };
  let c = s.charCodeAt(e1);
  if (endsName(c)) return { end: e1 };
  const w1 = readWord(s, p, e1);
  if (!w1) return { end: e1 };
  if (c === 58) {
    const e2 = skipWord(s, e1 + 1);
    if (e2 > e1 + 1) {
      c = s.charCodeAt(e2);
      const w2 = endsName(c) ? undefined : readWord(s, e1 + 1, e2);
      if (w2 && w2.kind === w1.kind) {
        return {
          end: e2,
          ref: {
            kind: w1.kind,
            pair: true,
            r1: w1.row,
            c1: w1.col,
            r2: w2.row,
            c2: w2.col,
            absR1: w1.absRow,
            absC1: w1.absCol,
            absR2: w2.absRow,
            absC2: w2.absCol,
          },
        };
      }
    }
  }
  // A lone "A" or "5" is a name or a number.
  if (w1.kind !== K_CELL) return { end: e1 };
  return {
    end: e1,
    ref: {
      kind: K_CELL,
      pair: false,
      r1: w1.row,
      c1: w1.col,
      r2: w1.row,
      c2: w1.col,
      absR1: w1.absRow,
      absC1: w1.absCol,
      absR2: w1.absRow,
      absC2: w1.absCol,
    },
  };
}

/**
 * Sheet names compare without regard to case. The four Turkish letters I, İ, ı, i count as
 * one, without relying on locale data: "İ" lower-cases to "i" plus a combining dot (U+0307).
 */
function foldSheetName(name: string): string {
  return String(name).toLowerCase().replace(/\u0307/g, '').replace(/\u0131/g, 'i');
}

function sameSheets(a: string[], b: string[]): boolean {
  if (a.length !== b.length) return false;
  for (let k = 0; k < a.length; k++) if (foldSheetName(a[k]) !== foldSheetName(b[k])) return false;
  return true;
}

function prefixed(s: string, start: number, bodyStart: number, sheets: string[] | undefined, external: boolean, chain: boolean): Item {
  const body = readBody(s, bodyStart);
  const item: Item = { start, bodyStart, end: body.end, ref: body.ref, sheets, external };
  // "Sheet1!A1:Sheet1!B2" is one range when both corners name the same sheet.
  const ref = body.ref;
  if (chain && ref && !ref.pair && sheets && !external && s.charCodeAt(body.end) === 58) {
    const p = body.end + 1;
    const c = s.charCodeAt(p);
    const second = c === 39 || isWordChar(c) ? readItem(s, p, false) : undefined;
    const corner = second && second.ref;
    if (second && corner && !corner.pair && second.sheets && !second.external && second.bodyStart > p && sameSheets(sheets, second.sheets)) {
      item.end = second.end;
      item.ref = { ...ref, pair: true, r2: corner.r1, c2: corner.c1, absR2: corner.absR1, absC2: corner.absC1, mid: s.slice(p, second.bodyStart) };
    }
  }
  return item;
}

/**
 * Reads the token at `i`, which starts with "'", "[" or a word character. Returns undefined
 * when a quote or bracket is never closed; the caller then leaves the rest of the text alone.
 * `chain` is false while reading the second corner of a range whose corners both carry a prefix.
 */
function readItem(s: string, i: number, chain = true): Item | undefined {
  const c = s.charCodeAt(i);
  if (c === 39) {
    const q = skipQuoted(s, i, "'");
    if (q < 0) return undefined;
    if (s.charCodeAt(q) !== 33) return { start: i, bodyStart: q, end: q, external: false };
    const name = s.slice(i + 1, q - 1).replace(/''/g, "'");
    // Sheet names cannot hold "[", "\" or "/": those come from a workbook name or path.
    const external = name.indexOf('[') >= 0 || name.indexOf('\\') >= 0 || name.indexOf('/') >= 0;
    return prefixed(s, i, q + 1, name.split(':'), external, chain);
  }
  if (c === 91) {
    const j = skipBrackets(s, i);
    if (j < 0) return undefined;
    // "[1]Sheet1!A1", "[1]Sheet1:Sheet3!A1", "[1]!Name": another workbook.
    const e = skipWord(s, j);
    if (s.charCodeAt(e) === 33) return prefixed(s, i, e + 1, undefined, true, false);
    if (e > j && s.charCodeAt(e) === 58) {
      const e2 = skipWord(s, e + 1);
      if (e2 > e + 1 && s.charCodeAt(e2) === 33) return prefixed(s, i, e2 + 1, undefined, true, false);
    }
    return { start: i, bodyStart: j, end: j, external: false };
  }
  const e1 = skipWord(s, i);
  const next = s.charCodeAt(e1);
  if (next === 33) return prefixed(s, i, e1 + 1, [s.slice(i, e1)], false, chain);
  if (next === 58) {
    const e2 = skipWord(s, e1 + 1);
    if (e2 > e1 + 1 && s.charCodeAt(e2) === 33) {
      return prefixed(s, i, e2 + 1, [s.slice(i, e1), s.slice(e1 + 1, e2)], false, chain);
    }
  }
  return prefixed(s, i, i, undefined, false, false);
}

/** Offset just past the error constant at `i`, or past the "#" alone. */
function skipError(s: string, i: number): number {
  const head = s.slice(i, i + MAX_ERROR_LENGTH).toUpperCase();
  for (let k = 0; k < ERROR_LITERALS.length; k++) {
    const literal = ERROR_LITERALS[k];
    if (!head.startsWith(literal)) continue;
    const end = i + literal.length;
    // "#REF!A1" is what is left of a reference to a deleted sheet; the cell is part of it.
    return literal === REF_ERROR ? readBody(s, end).end : end;
  }
  return i + 1;
}

function rewrite(s: string, transform: Transform): string {
  const n = s.length;
  let out = '';
  let copied = 0;
  let i = 0;
  while (i < n) {
    const c = s.charCodeAt(i);
    if (c === 34) {
      const q = skipQuoted(s, i, '"');
      if (q < 0) break;
      i = q;
    } else if (c === 35) {
      i = skipError(s, i);
    } else if (c === 39 || c === 91 || isWordChar(c)) {
      const item = readItem(s, i);
      if (!item) break;
      if (item.ref) {
        const replacement = transform(item.ref, item.sheets, item.external);
        if (replacement !== undefined) {
          out += s.slice(copied, item.bodyStart) + replacement;
          copied = item.end;
        }
      }
      i = item.end;
    } else {
      i++;
    }
  }
  return copied === 0 ? s : out + s.slice(copied);
}

// ---------------------------------------------------------------------------
// Writing

function formatRef(ref: Ref): string {
  const col1 = (ref.absC1 ? '$' : '') + columnName(ref.c1);
  const col2 = (ref.absC2 ? '$' : '') + columnName(ref.c2);
  const row1 = (ref.absR1 ? '$' : '') + (ref.r1 + 1);
  const row2 = (ref.absR2 ? '$' : '') + (ref.r2 + 1);
  if (ref.kind === K_COL) return `${col1}:${col2}`;
  if (ref.kind === K_ROW) return `${row1}:${row2}`;
  return ref.pair ? `${col1}${row1}:${ref.mid || ''}${col2}${row2}` : col1 + row1;
}

// ---------------------------------------------------------------------------
// Relative shift

/** What comes back for a text argument that is not a string: nothing for null and undefined. */
function toText(value: unknown): string {
  return value === null || value === undefined ? '' : String(value);
}

/** NaN and fractions would otherwise end up in the text. */
function toInt(value: number): number {
  return value === value ? Math.trunc(value) : 0;
}

/**
 * Moves relative references by (dRow, dCol), as copying or filling a formula does; parts
 * anchored with "$" stay. A reference pushed off the sheet becomes #REF!. References into
 * other workbooks move like any other.
 */
export function shiftFormula(text: string, dRow: number, dCol: number): string {
  if (typeof text !== 'string') return toText(text);
  dRow = toInt(dRow);
  dCol = toInt(dCol);
  if (dRow === 0 && dCol === 0) return text;
  return rewrite(text, (ref) => {
    let { r1, c1, r2, c2 } = ref;
    if (ref.kind !== K_ROW) {
      if (!ref.absC1) c1 += dCol;
      if (!ref.absC2) c2 += dCol;
      if (c1 < 0 || c2 < 0 || c1 >= MAX_COLS || c2 >= MAX_COLS) return REF_ERROR;
    }
    if (ref.kind !== K_COL) {
      if (!ref.absR1) r1 += dRow;
      if (!ref.absR2) r2 += dRow;
      if (r1 < 0 || r2 < 0 || r1 >= MAX_ROWS || r2 >= MAX_ROWS) return REF_ERROR;
    }
    if (r1 === ref.r1 && c1 === ref.c1 && r2 === ref.r2 && c2 === ref.c2) return undefined;
    return formatRef({ ...ref, r1, c1, r2, c2 });
  });
}

// ---------------------------------------------------------------------------
// Structural edits

type AxisEdit = Pick<StructuralEdit, 'kind' | 'axis' | 'at' | 'count'>;

/** An edit that moves nothing (bad numbers, zero count) leaves every reference alone. */
function isEffective(edit: AxisEdit): boolean {
  return !!edit && edit.at >= 0 && edit.count >= 1 && (edit.kind === 'insert' || edit.kind === 'delete');
}

/**
 * Where the inclusive span lo..hi of rows or columns ends up; undefined when nothing of it
 * is left. `limit` is the size of the sheet along the axis.
 */
function adjustSpan(lo: number, hi: number, limit: number, edit: AxisEdit): [number, number] | undefined {
  const at = Math.trunc(edit.at);
  const count = Math.trunc(edit.count);
  if (edit.kind === 'insert') {
    if (at <= lo) {
      lo += count;
      hi += count;
    } else if (at <= hi) {
      // Inserted inside: the span grows.
      hi += count;
    }
    if (lo >= limit) return undefined;
    // Whatever is pushed past the last row or column is cut off.
    if (hi >= limit) hi = limit - 1;
    return [lo, hi];
  }
  const end = at + count - 1;
  if (hi < at) return [lo, hi];
  if (lo > end) return [lo - count, hi - count];
  if (lo >= at && hi <= end) return undefined;
  return [lo < at ? lo : at, hi > end ? hi - count : at - 1];
}

function adjustRef(ref: Ref, edit: AxisEdit): string | undefined {
  const rows = edit.axis === 'row';
  // A whole-column reference has no rows to move, and the other way round.
  if (ref.kind === (rows ? K_COL : K_ROW)) return undefined;
  const a = rows ? ref.r1 : ref.c1;
  const b = rows ? ref.r2 : ref.c2;
  const span = adjustSpan(Math.min(a, b), Math.max(a, b), rows ? MAX_ROWS : MAX_COLS, edit);
  if (!span) return REF_ERROR;
  // Corners written in reverse order ("B5:A1") stay that way.
  const first = a <= b ? span[0] : span[1];
  const second = a <= b ? span[1] : span[0];
  if (first === a && second === b) return undefined;
  return formatRef(rows ? { ...ref, r1: first, r2: second } : { ...ref, c1: first, c2: second });
}

/** Tells per reference whether it points at the edited sheet. `ownSheetName` is the sheet of the formula. */
function sheetMatcher(ownSheetName: string, edit: StructuralEdit): (sheets: string[] | undefined, external: boolean) => boolean {
  const edited = foldSheetName(edit.sheetName);
  const ownMatches = foldSheetName(ownSheetName) === edited;
  return (sheets, external) => {
    // Rows and columns of another workbook did not move.
    if (external) return false;
    if (!sheets) return ownMatches;
    // A 3D reference is adjusted when the edited sheet is one of its two ends; the sheets
    // between them are not known here.
    for (let k = 0; k < sheets.length; k++) if (foldSheetName(sheets[k]) === edited) return true;
    return false;
  };
}

function editTransform(ownSheetName: string, edit: StructuralEdit): Transform {
  const onEditedSheet = sheetMatcher(ownSheetName, edit);
  return (ref, sheets, external) => (onEditedSheet(sheets, external) ? adjustRef(ref, edit) : undefined);
}

/**
 * Rewrites the references of a formula after a structural edit. `formulaSheetName` is the
 * sheet the formula is in: references without a sheet name belong to it.
 */
export function adjustFormula(text: string, formulaSheetName: string, edit: StructuralEdit): string {
  if (typeof text !== 'string') return toText(text);
  if (!isEffective(edit)) return text;
  return rewrite(text, editTransform(formulaSheetName, edit));
}

/**
 * One axis of a reference in an anchored formula: the two ends a and b, each either anchored
 * with "$" (it follows `edit`, when the edit is about this axis of its sheet) or relative (it
 * moves by `delta` with the anchor cell). Undefined when the result is #REF!.
 */
function adjustAnchoredAxis(a: number, b: number, absA: boolean, absB: boolean, delta: number, limit: number, edit: AxisEdit | undefined): [number, number] | undefined {
  if (absA && absB) {
    // Fully anchored: a span, with the growing and shrinking of adjustFormula.
    if (!edit) return [a, b];
    const span = adjustSpan(Math.min(a, b), Math.max(a, b), limit, edit);
    if (!span) return undefined;
    return a <= b ? span : [span[1], span[0]];
  }
  if (!absA) a += delta;
  else if (edit) {
    const moved = adjustSpan(a, a, limit, edit);
    if (!moved) return undefined;
    a = moved[0];
  }
  if (!absB) b += delta;
  else if (edit) {
    const moved = adjustSpan(b, b, limit, edit);
    if (!moved) return undefined;
    b = moved[0];
  }
  if (a < 0 || b < 0 || a >= limit || b >= limit) return undefined;
  return [a, b];
}

/**
 * Adjusts a formula that is written relative to an anchor cell (conditional formats, data
 * validations) after a structural edit. Parts of a reference anchored with "$" follow the
 * edit as in adjustFormula; relative parts keep their distance from the anchor, which moved
 * from `before` to `after`.
 */
export function adjustAnchoredFormula(
  text: string,
  formulaSheetName: string,
  edit: StructuralEdit,
  before: { row: number; col: number },
  after: { row: number; col: number },
): string {
  if (typeof text !== 'string') return toText(text);
  const dRow = before && after ? toInt(after.row - before.row) : 0;
  const dCol = before && after ? toInt(after.col - before.col) : 0;
  const effective = isEffective(edit);
  if (!effective && dRow === 0 && dCol === 0) return text;
  const onEditedSheet = effective ? sheetMatcher(formulaSheetName, edit) : () => false;
  return rewrite(text, (ref, sheets, external) => {
    // Relative parts move with the anchor whatever sheet they point at; anchored parts only
    // follow an edit of their own sheet.
    const edited = onEditedSheet(sheets, external);
    let { r1, c1, r2, c2 } = ref;
    if (ref.kind !== K_COL) {
      const rows = adjustAnchoredAxis(r1, r2, ref.absR1, ref.absR2, dRow, MAX_ROWS, edited && edit.axis === 'row' ? edit : undefined);
      if (!rows) return REF_ERROR;
      r1 = rows[0];
      r2 = rows[1];
    }
    if (ref.kind !== K_ROW) {
      const cols = adjustAnchoredAxis(c1, c2, ref.absC1, ref.absC2, dCol, MAX_COLS, edited && edit.axis === 'col' ? edit : undefined);
      if (!cols) return REF_ERROR;
      c1 = cols[0];
      c2 = cols[1];
    }
    if (r1 === ref.r1 && c1 === ref.c1 && r2 === ref.r2 && c2 === ref.c2) return undefined;
    return formatRef({ ...ref, r1, c1, r2, c2 });
  });
}

/** The same adjustment for a plain range (merged cells, hyperlinks); undefined when all of it was deleted. */
export function adjustRange(range: CellRange, edit: Pick<StructuralEdit, 'kind' | 'axis' | 'at' | 'count'>): CellRange | undefined {
  if (!range) return undefined;
  const top = Math.min(range.top, range.bottom);
  const bottom = Math.max(range.top, range.bottom);
  const left = Math.min(range.left, range.right);
  const right = Math.max(range.left, range.right);
  if (!isEffective(edit)) return { top, left, bottom, right };
  if (edit.axis === 'row') {
    const span = adjustSpan(top, bottom, MAX_ROWS, edit);
    return span && { top: span[0], left, bottom: span[1], right };
  }
  const span = adjustSpan(left, right, MAX_COLS, edit);
  return span && { top, left: span[0], bottom, right: span[1] };
}

/** Splits "A1:B2 D5" or "Sheet1!$A$1" into references; undefined when the text is anything more than such a list. */
function readList(s: string): Item[] | undefined {
  const n = s.length;
  const items: Item[] = [];
  let i = 0;
  for (;;) {
    while (i < n && isSpace(s.charCodeAt(i))) i++;
    if (i >= n) break;
    const c = s.charCodeAt(i);
    if (c !== 39 && c !== 91 && !isWordChar(c)) return undefined;
    const item = readItem(s, i);
    if (!item || !item.ref) return undefined;
    if (item.end < n && !isSpace(s.charCodeAt(item.end))) return undefined;
    items.push(item);
    i = item.end;
  }
  return items.length > 0 ? items : undefined;
}

/** Adjusts a reference list such as an sqref attribute or a defined name value ("A1:B2 D5", "Sheet1!$A$1:$B$2"). */
export function adjustReferenceText(text: string, defaultSheetName: string, edit: StructuralEdit): string {
  if (typeof text !== 'string') return toText(text);
  if (!isEffective(edit)) return text;
  const transform = editTransform(defaultSheetName, edit);
  const items = readList(text);
  // Not a plain list: a formula, where a space is the intersection operator.
  if (!items) return rewrite(text, transform);
  const kept: string[] = [];
  let out = '';
  let copied = 0;
  let dropped = false;
  for (let k = 0; k < items.length; k++) {
    const item = items[k];
    const replacement = transform(item.ref as Ref, item.sheets, item.external);
    if (replacement === REF_ERROR) {
      dropped = true;
      continue;
    }
    if (replacement === undefined) {
      kept.push(text.slice(item.start, item.end));
      continue;
    }
    kept.push(text.slice(item.start, item.bodyStart) + replacement);
    out += text.slice(copied, item.bodyStart) + replacement;
    copied = item.end;
  }
  if (dropped) return kept.join(' ');
  return copied === 0 ? text : out + text.slice(copied);
}
