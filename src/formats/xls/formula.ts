// Parsed formulas of BIFF records (token arrays, "rgce") -> formula text.
//
// The model keeps formulas as text, written as xlsx files write them: English
// function names, "," between arguments, A1 references. The tokens are in
// postfix order; the text is built on a stack of operands. A token array that
// this reader cannot turn into text for certain gives no text at all: the
// caller keeps the cell's value and drops the formula rather than show a
// wrong one.

import { columnName } from '../../model/sheet/types';
import { decodeBytes } from './codepage';
import { FUNC_USER_DEFINED, fixedArgCount, functionName } from './functions';

/** A workbook the formulas of this one refer to (SUPBOOK), or this workbook itself. */
export interface XlsSupBook {
  /** "self": this workbook; "addin": functions of add-ins; "external": another workbook or a DDE/OLE source. */
  kind: 'self' | 'addin' | 'external';
  sheetNames: string[];
  /** Names defined in that workbook (EXTERNNAME), in file order. */
  externNames: string[];
  /** Position among the external workbooks, counted from 1: the "[1]" of "[1]Sheet1!A1". */
  ordinal: number;
}

/** One entry of the EXTERNSHEET table: a range of sheets of a workbook. */
export interface XlsXti {
  book: number;
  /** Sheet indexes; -1 for a deleted sheet, -2 for "no sheet" (a workbook-level name). */
  first: number;
  last: number;
}

export interface FormulaContext {
  biff8: boolean;
  /** Code page of BIFF5 byte strings. */
  codepage: number;
  sheetNames: string[];
  /** Names of the NAME records, in file order; PtgName counts from 1. */
  names: string[];
  xti: XlsXti[];
  supBooks: XlsSupBook[];
  /** Why formulas could not be turned into text: reason -> number of formulas. */
  failures: Map<string, number>;
}

const P_COMPARE = 1;
const P_CONCAT = 2;
const P_ADD = 3;
const P_MULTIPLY = 4;
const P_POWER = 5;
const P_PERCENT = 6;
const P_UNARY = 7;
const P_UNION = 8;
const P_INTERSECT = 9;
const P_RANGE = 10;
const P_ATOM = 11;

/** Binary operators by token, 0x03..0x11. */
const BINARY: readonly [text: string, precedence: number][] = [
  ['+', P_ADD],
  ['-', P_ADD],
  ['*', P_MULTIPLY],
  ['/', P_MULTIPLY],
  ['^', P_POWER],
  ['&', P_CONCAT],
  ['<', P_COMPARE],
  ['<=', P_COMPARE],
  ['=', P_COMPARE],
  ['>=', P_COMPARE],
  ['>', P_COMPARE],
  ['<>', P_COMPARE],
  [' ', P_INTERSECT],
  [',', P_UNION],
  [':', P_RANGE],
];

const ERRORS: Readonly<Record<number, string>> = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

/** Text of an error code, as cells and formulas hold it; undefined for a code that is not an error value. */
export function errorText(code: number): string | undefined {
  return ERRORS[code];
}

/** A sheet name as a formula writes it: quoted unless it is a plain ASCII word that cannot be mistaken for a reference. */
export function quoteSheetName(name: string): string {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}[0-9]+$/.test(name) && !/^[Rr][0-9]*[Cc]?[0-9]*$/.test(name) && !/^(true|false)$/i.test(name)) {
    return name;
  }
  return `'${name.replace(/'/g, "''")}'`;
}

/** A number as a formula writes it, without the sign. */
function numberText(value: number): string {
  return String(value).replace('e', 'E');
}

class Failure {
  constructor(readonly reason: string) {}
}

const fail = (reason: string): never => {
  throw new Failure(reason);
};

const signed = (value: number, bits: number): number => (value >= 1 << (bits - 1) ? value - (1 << bits) : value);

interface Decoder {
  ctx: FormulaContext;
  bytes: Uint8Array;
  view: DataView;
  /** Read position and end of the tokens. */
  p: number;
  end: number;
  /** Read position and end of the data that follows the tokens (array constants). */
  extra: number;
  extraEnd: number;
  /** The cell the formula is read for; relative tokens count from it. */
  row: number;
  col: number;
  texts: string[];
  precs: number[];
}

function need(d: Decoder, size: number): void {
  if (d.p + size > d.end) fail('truncated token');
}

function push(d: Decoder, text: string, precedence: number): void {
  d.texts.push(text);
  d.precs.push(precedence);
}

/** Row and column limits of the BIFF version: references wrap around inside them. */
const rowMask = (d: Decoder) => (d.ctx.biff8 ? 0xffff : 0x3fff);

function cellText(row: number, col: number, rowRelative: boolean, colRelative: boolean): string {
  return (colRelative ? '' : '$') + columnName(col) + (rowRelative ? '' : '$') + (row + 1);
}

function areaText(d: Decoder, r1: number, r2: number, c1: number, c2: number, flags: number): string {
  // flags: bit 0 first row relative, 1 first column relative, 2 last row relative, 3 last column relative.
  const r1Rel = (flags & 1) !== 0;
  const c1Rel = (flags & 2) !== 0;
  const r2Rel = (flags & 4) !== 0;
  const c2Rel = (flags & 8) !== 0;
  // All rows of the old sheet size stand for the whole column, however many rows a sheet has now; likewise for rows.
  if (r1 === 0 && r2 === rowMask(d) && !(c1 === 0 && c2 === 0xff)) {
    return (c1Rel ? '' : '$') + columnName(c1) + ':' + (c2Rel ? '' : '$') + columnName(c2);
  }
  if (c1 === 0 && c2 === 0xff && !(r1 === 0 && r2 === rowMask(d))) {
    return (r1Rel ? '' : '$') + (r1 + 1) + ':' + (r2Rel ? '' : '$') + (r2 + 1);
  }
  return cellText(r1, c1, r1Rel, c1Rel) + ':' + cellText(r2, c2, r2Rel, c2Rel);
}

/** Reads the cell address of a reference token. `offsets`: relative parts are distances from the formula's cell. */
function readRef(d: Decoder, offsets: boolean): string {
  let row: number;
  let col: number;
  let rowRel: boolean;
  let colRel: boolean;
  if (d.ctx.biff8) {
    need(d, 4);
    row = d.view.getUint16(d.p, true);
    const field = d.view.getUint16(d.p + 2, true);
    d.p += 4;
    col = field & 0xff;
    colRel = (field & 0x4000) !== 0;
    rowRel = (field & 0x8000) !== 0;
    if (offsets) {
      if (rowRel) row = (d.row + signed(row, 16)) & 0xffff;
      if (colRel) col = (d.col + signed(col, 8)) & 0xff;
    }
  } else {
    need(d, 3);
    const field = d.view.getUint16(d.p, true);
    col = d.bytes[d.p + 2];
    d.p += 3;
    row = field & 0x3fff;
    colRel = (field & 0x4000) !== 0;
    rowRel = (field & 0x8000) !== 0;
    if (offsets) {
      if (rowRel) row = (d.row + signed(row, 14)) & 0x3fff;
      if (colRel) col = (d.col + signed(col, 8)) & 0xff;
    }
  }
  return cellText(row, col, rowRel, colRel);
}

function readArea(d: Decoder, offsets: boolean): string {
  let r1: number;
  let r2: number;
  let c1: number;
  let c2: number;
  let flags: number;
  if (d.ctx.biff8) {
    need(d, 8);
    r1 = d.view.getUint16(d.p, true);
    r2 = d.view.getUint16(d.p + 2, true);
    const f1 = d.view.getUint16(d.p + 4, true);
    const f2 = d.view.getUint16(d.p + 6, true);
    d.p += 8;
    c1 = f1 & 0xff;
    c2 = f2 & 0xff;
    flags = (f1 & 0x8000 ? 1 : 0) | (f1 & 0x4000 ? 2 : 0) | (f2 & 0x8000 ? 4 : 0) | (f2 & 0x4000 ? 8 : 0);
    if (offsets) {
      if (flags & 1) r1 = (d.row + signed(r1, 16)) & 0xffff;
      if (flags & 4) r2 = (d.row + signed(r2, 16)) & 0xffff;
    }
  } else {
    need(d, 6);
    const f1 = d.view.getUint16(d.p, true);
    const f2 = d.view.getUint16(d.p + 2, true);
    c1 = d.bytes[d.p + 4];
    c2 = d.bytes[d.p + 5];
    d.p += 6;
    r1 = f1 & 0x3fff;
    r2 = f2 & 0x3fff;
    flags = (f1 & 0x8000 ? 1 : 0) | (f1 & 0x4000 ? 2 : 0) | (f2 & 0x8000 ? 4 : 0) | (f2 & 0x4000 ? 8 : 0);
    if (offsets) {
      if (flags & 1) r1 = (d.row + signed(r1, 14)) & 0x3fff;
      if (flags & 4) r2 = (d.row + signed(r2, 14)) & 0x3fff;
    }
  }
  if (offsets) {
    if (flags & 2) c1 = (d.col + signed(c1, 8)) & 0xff;
    if (flags & 8) c2 = (d.col + signed(c2, 8)) & 0xff;
  }
  return areaText(d, r1, r2, c1, c2, flags);
}

/** "Sheet1!", "'My sheet'!", "Sheet1:Sheet3!", "[1]Sheet1!"; fails for sheets that are gone. */
function sheetPrefix(d: Decoder, what: string): string {
  const ctx = d.ctx;
  if (ctx.biff8) {
    need(d, 2);
    const index = d.view.getUint16(d.p, true);
    d.p += 2;
    const xti = ctx.xti[index];
    const book = xti ? ctx.supBooks[xti.book] : undefined;
    if (!xti || !book) return fail(`${what}: no EXTERNSHEET entry`);
    if (book.kind === 'addin') return fail(`${what}: add-in reference`);
    const names = book.kind === 'self' ? ctx.sheetNames : book.sheetNames;
    const first = names[xti.first];
    const last = names[xti.last];
    if (xti.first < 0 || xti.last < 0 || first === undefined || last === undefined) return fail(`${what}: deleted or unknown sheet`);
    const range = first === last ? first : `${first}:${last}`;
    if (book.kind === 'self') return range === first ? `${quoteSheetName(first)}!` : `${quoteRange(first, last)}!`;
    const plain = quoteSheetName(first) === first && quoteSheetName(last) === last;
    return plain ? `[${book.ordinal}]${range}!` : `'[${book.ordinal}]${range.replace(/'/g, "''")}'!`;
  }
  // BIFF5: a negative index means this workbook, with the sheet indexes in the token itself.
  need(d, 14);
  const index = d.view.getInt16(d.p, true);
  const firstTab = d.view.getUint16(d.p + 10, true);
  const lastTab = d.view.getUint16(d.p + 12, true);
  d.p += 14;
  if (index >= 0) {
    const book = ctx.supBooks[index - 1];
    if (book && book.kind === 'self' && book.sheetNames.length === 1) return `${quoteSheetName(book.sheetNames[0])}!`;
    return fail(`${what}: external reference (BIFF5)`);
  }
  const first = ctx.sheetNames[firstTab];
  const last = ctx.sheetNames[lastTab];
  if (first === undefined || last === undefined) return fail(`${what}: deleted or unknown sheet`);
  return first === last ? `${quoteSheetName(first)}!` : `${quoteRange(first, last)}!`;
}

function quoteRange(first: string, last: string): string {
  if (quoteSheetName(first) === first && quoteSheetName(last) === last) return `${first}:${last}`;
  return `'${first.replace(/'/g, "''")}:${last.replace(/'/g, "''")}'`;
}

function stringLiteral(text: string): string {
  return `"${text.replace(/"/g, '""')}"`;
}

/** Text of `count` characters at `at`, for a BIFF8 string with its flags byte already read. */
function wideText(d: Decoder, at: number, count: number, wide: boolean, limit: number): string {
  if (at + count * (wide ? 2 : 1) > limit) return fail('truncated string');
  let out = '';
  if (wide) {
    for (let i = 0; i < count; i++) out += String.fromCharCode(d.view.getUint16(at + i * 2, true));
  } else {
    for (let i = 0; i < count; i++) out += String.fromCharCode(d.bytes[at + i]);
  }
  return out;
}

/** The array constant of a PtgArray, read from the data after the tokens: "{1,2;3,4}". */
function readArrayConstant(d: Decoder): string {
  const biff8 = d.ctx.biff8;
  const limit = d.extraEnd;
  let p = d.extra;
  if (p + 3 > limit) return fail('PtgArray: no data');
  let cols = d.bytes[p];
  let rows = d.view.getUint16(p + 1, true);
  p += 3;
  if (biff8) {
    cols += 1;
    rows += 1;
  } else if (cols === 0) cols = 256;
  let out = '{';
  for (let r = 0; r < rows; r++) {
    if (r > 0) out += ';';
    for (let c = 0; c < cols; c++) {
      if (c > 0) out += ',';
      if (p + 1 > limit) return fail('PtgArray: truncated data');
      const type = d.bytes[p++];
      if (type === 0x01) {
        if (p + 8 > limit) return fail('PtgArray: truncated data');
        const value = d.view.getFloat64(p, true);
        if (!isFinite(value)) return fail('PtgArray: not a number');
        out += value < 0 ? `-${numberText(-value)}` : numberText(value);
        p += 8;
      } else if (type === 0x02) {
        if (biff8) {
          if (p + 3 > limit) return fail('PtgArray: truncated data');
          const count = d.view.getUint16(p, true);
          const wide = (d.bytes[p + 2] & 1) !== 0;
          out += stringLiteral(wideText(d, p + 3, count, wide, limit));
          p += 3 + count * (wide ? 2 : 1);
        } else {
          const count = d.bytes[p];
          if (p + 1 + count > limit) return fail('PtgArray: truncated data');
          out += stringLiteral(decodeBytes(d.bytes, p + 1, p + 1 + count, d.ctx.codepage));
          p += 1 + count;
        }
      } else if (type === 0x04) {
        if (p + 8 > limit) return fail('PtgArray: truncated data');
        out += d.bytes[p] ? 'TRUE' : 'FALSE';
        p += 8;
      } else if (type === 0x10) {
        if (p + 8 > limit) return fail('PtgArray: truncated data');
        const error = ERRORS[d.bytes[p]];
        if (error === undefined) return fail('PtgArray: unknown error value');
        out += error;
        p += 8;
      } else {
        // An empty element has no way to be written in formula text.
        return fail(type === 0x00 ? 'PtgArray: empty element' : 'PtgArray: unknown element type');
      }
    }
  }
  d.extra = p;
  return out + '}';
}

/** Skips the rectangles a PtgMemArea keeps after the tokens. */
function skipMemAreaData(d: Decoder): void {
  if (d.extra + 2 > d.extraEnd) return;
  const count = d.view.getUint16(d.extra, true);
  d.extra = Math.min(d.extraEnd, d.extra + 2 + count * (d.ctx.biff8 ? 8 : 6));
}

function popArgs(d: Decoder, count: number, what: string): string[] {
  if (d.texts.length < count) return fail(`${what}: too few operands`);
  const at = d.texts.length - count;
  const args = d.texts.splice(at, count);
  const precs = d.precs.splice(at, count);
  // A union as an argument would read as two arguments without its parentheses.
  for (let i = 0; i < count; i++) if (precs[i] === P_UNION) args[i] = `(${args[i]})`;
  return args;
}

function callFunction(d: Decoder, index: number, count: number): void {
  if (index === FUNC_USER_DEFINED) {
    // The first operand names the function: an add-in function or one of a later Excel version.
    if (count < 1) return fail('user-defined function without a name');
    const args = popArgs(d, count, 'PtgFuncVar');
    const name = args[0];
    if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name)) return fail('user-defined function with an unusable name');
    push(d, `${name}(${args.slice(1).join(',')})`, P_ATOM);
    return;
  }
  const name = functionName(index);
  if (name === undefined) return fail(`unknown function ${index}`);
  push(d, `${name}(${popArgs(d, count, name).join(',')})`, P_ATOM);
}

function nameText(d: Decoder, index: number): string {
  const name = d.ctx.names[index - 1];
  if (name === undefined || name === '') return fail('PtgName: unknown name');
  return name;
}

function decode(d: Decoder): string {
  const ctx = d.ctx;
  const biff8 = ctx.biff8;
  const bytes = d.bytes;
  const view = d.view;
  while (d.p < d.end) {
    const ptg = bytes[d.p++];
    // Tokens from 0x20 on come in three classes (reference, value, array) that read the same.
    const base = ptg >= 0x20 ? (ptg & 0x1f) | 0x20 : ptg;
    if (ptg >= 0x80) return fail(`unknown token 0x${ptg.toString(16)}`);

    if (base >= 0x03 && base <= 0x11) {
      if (d.texts.length < 2) return fail('operator without operands');
      const [op, precedence] = BINARY[base - 0x03];
      const right = d.texts.pop() as string;
      const rightPrec = d.precs.pop() as number;
      const left = d.texts.pop() as string;
      const leftPrec = d.precs.pop() as number;
      // Operators group from the left, so an operand of the same precedence on the right needs parentheses.
      push(d, (leftPrec < precedence ? `(${left})` : left) + op + (rightPrec <= precedence ? `(${right})` : right), precedence);
      continue;
    }

    switch (base) {
      case 0x01:
        return fail('PtgExp inside a formula');
      case 0x02:
        return fail('PtgTbl inside a formula');
      case 0x12:
      case 0x13: {
        if (d.texts.length < 1) return fail('operator without operands');
        const at = d.texts.length - 1;
        d.texts[at] = (base === 0x12 ? '+' : '-') + (d.precs[at] < P_UNARY ? `(${d.texts[at]})` : d.texts[at]);
        d.precs[at] = P_UNARY;
        break;
      }
      case 0x14: {
        if (d.texts.length < 1) return fail('operator without operands');
        const at = d.texts.length - 1;
        d.texts[at] = (d.precs[at] < P_PERCENT ? `(${d.texts[at]})` : d.texts[at]) + '%';
        d.precs[at] = P_PERCENT;
        break;
      }
      case 0x15: {
        if (d.texts.length < 1) return fail('parentheses without an operand');
        const at = d.texts.length - 1;
        d.texts[at] = `(${d.texts[at]})`;
        d.precs[at] = P_ATOM;
        break;
      }
      case 0x16:
        push(d, '', P_ATOM);
        break;
      case 0x17: {
        need(d, 1);
        const count = bytes[d.p];
        if (biff8) {
          need(d, 2);
          const wide = (bytes[d.p + 1] & 1) !== 0;
          push(d, stringLiteral(wideText(d, d.p + 2, count, wide, d.end)), P_ATOM);
          d.p += 2 + count * (wide ? 2 : 1);
        } else {
          need(d, 1 + count);
          push(d, stringLiteral(decodeBytes(bytes, d.p + 1, d.p + 1 + count, ctx.codepage)), P_ATOM);
          d.p += 1 + count;
        }
        break;
      }
      case 0x19: {
        need(d, 3);
        const flags = bytes[d.p];
        const data = view.getUint16(d.p + 1, true);
        d.p += 3;
        if (flags & 0x04) {
          // CHOOSE: a table of jump offsets follows, one more than there are choices.
          need(d, (data + 1) * 2);
          d.p += (data + 1) * 2;
        } else if (flags & 0x10) {
          if (d.texts.length < 1) return fail('SUM without an operand');
          push(d, `SUM(${popArgs(d, 1, 'SUM')[0]})`, P_ATOM);
        } else if (flags & 0x20) {
          return fail('PtgAttrBaxcel');
        }
        // Volatile, IF and GOTO marks, and spaces say nothing the text needs.
        break;
      }
      case 0x1c: {
        need(d, 1);
        const error = ERRORS[bytes[d.p++]];
        if (error === undefined) return fail('PtgErr: unknown error value');
        push(d, error, P_ATOM);
        break;
      }
      case 0x1d:
        need(d, 1);
        push(d, bytes[d.p++] ? 'TRUE' : 'FALSE', P_ATOM);
        break;
      case 0x1e:
        need(d, 2);
        push(d, String(view.getUint16(d.p, true)), P_ATOM);
        d.p += 2;
        break;
      case 0x1f: {
        need(d, 8);
        const value = view.getFloat64(d.p, true);
        d.p += 8;
        if (!isFinite(value)) return fail('PtgNum: not a number');
        if (value < 0 || (value === 0 && 1 / value < 0)) push(d, `-${numberText(-value)}`, P_UNARY);
        else push(d, numberText(value), P_ATOM);
        break;
      }
      case 0x20:
        need(d, 7);
        d.p += 7;
        push(d, readArrayConstant(d), P_ATOM);
        break;
      case 0x21: {
        need(d, 2);
        const index = view.getUint16(d.p, true);
        d.p += 2;
        const count = fixedArgCount(index);
        if (count === undefined) return fail(`PtgFunc: function ${functionName(index) || index} has no fixed argument count`);
        callFunction(d, index, count);
        break;
      }
      case 0x22: {
        need(d, 3);
        const count = bytes[d.p] & 0x7f;
        const index = view.getUint16(d.p + 1, true);
        d.p += 3;
        if (index & 0x8000) return fail('PtgFuncVar: command equivalent');
        callFunction(d, index, count);
        break;
      }
      case 0x23: {
        need(d, biff8 ? 4 : 14);
        const index = view.getUint16(d.p, true);
        d.p += biff8 ? 4 : 14;
        push(d, nameText(d, index), P_ATOM);
        break;
      }
      case 0x24:
        push(d, readRef(d, false), P_ATOM);
        break;
      case 0x25:
        push(d, readArea(d, false), P_ATOM);
        break;
      case 0x26:
        // The reference this stands for is spelled out by the tokens that follow.
        need(d, 6);
        d.p += 6;
        skipMemAreaData(d);
        break;
      case 0x27:
      case 0x28:
        need(d, 6);
        d.p += 6;
        break;
      case 0x29:
      case 0x2e:
      case 0x2f:
        need(d, 2);
        d.p += 2;
        break;
      case 0x2a:
        need(d, biff8 ? 4 : 3);
        d.p += biff8 ? 4 : 3;
        push(d, '#REF!', P_ATOM);
        break;
      case 0x2b:
        need(d, biff8 ? 8 : 6);
        d.p += biff8 ? 8 : 6;
        push(d, '#REF!', P_ATOM);
        break;
      case 0x2c:
        push(d, readRef(d, true), P_ATOM);
        break;
      case 0x2d:
        push(d, readArea(d, true), P_ATOM);
        break;
      case 0x39: {
        // A name of another sheet or workbook, or a function of an add-in.
        if (biff8) {
          need(d, 6);
          const xti = ctx.xti[view.getUint16(d.p, true)];
          const index = view.getUint16(d.p + 2, true);
          d.p += 6;
          const book = xti ? ctx.supBooks[xti.book] : undefined;
          if (!xti || !book) return fail('PtgNameX: no EXTERNSHEET entry');
          if (book.kind === 'self') {
            const name = nameText(d, index);
            const sheet = xti.first >= 0 ? ctx.sheetNames[xti.first] : undefined;
            push(d, sheet === undefined ? name : `${quoteSheetName(sheet)}!${name}`, P_ATOM);
          } else {
            const name = book.externNames[index - 1];
            if (name === undefined || name === '') return fail('PtgNameX: unknown external name');
            if (book.kind === 'addin') push(d, name, P_ATOM);
            else {
              const sheet = xti.first >= 0 ? book.sheetNames[xti.first] : undefined;
              if (sheet === undefined) push(d, `[${book.ordinal}]!${name}`, P_ATOM);
              else if (quoteSheetName(sheet) === sheet) push(d, `[${book.ordinal}]${sheet}!${name}`, P_ATOM);
              else push(d, `'[${book.ordinal}]${sheet.replace(/'/g, "''")}'!${name}`, P_ATOM);
            }
          }
        } else {
          need(d, 24);
          const bookIndex = view.getInt16(d.p, true);
          const index = view.getUint16(d.p + 10, true);
          d.p += 24;
          if (bookIndex < 0) push(d, nameText(d, index), P_ATOM);
          else {
            const book = ctx.supBooks[bookIndex - 1];
            const name = book && book.kind === 'addin' ? book.externNames[index - 1] : undefined;
            if (name === undefined || name === '') return fail('PtgNameX: external name (BIFF5)');
            push(d, name, P_ATOM);
          }
        }
        break;
      }
      case 0x3a: {
        const prefix = sheetPrefix(d, 'PtgRef3d');
        push(d, prefix + readRef(d, false), P_ATOM);
        break;
      }
      case 0x3b: {
        const prefix = sheetPrefix(d, 'PtgArea3d');
        push(d, prefix + readArea(d, false), P_ATOM);
        break;
      }
      case 0x3c:
        need(d, biff8 ? 6 : 17);
        d.p += biff8 ? 6 : 17;
        push(d, '#REF!', P_ATOM);
        break;
      case 0x3d:
        need(d, biff8 ? 10 : 20);
        d.p += biff8 ? 10 : 20;
        push(d, '#REF!', P_ATOM);
        break;
      default:
        return fail(`unsupported token 0x${ptg.toString(16)}`);
    }
  }
  if (d.texts.length !== 1) return fail(d.texts.length === 0 ? 'empty formula' : 'operands left over');
  const text = d.texts[0];
  // A missing argument on its own is not a formula.
  if (text === '') return fail('empty formula');
  return text;
}

// Formulas of one sheet all come from the same stream; its view is made once.
let viewedBytes: Uint8Array | undefined;
let viewOfBytes: DataView | undefined;

/**
 * Formula text (without "=") for the tokens in bytes[start..end). `extraStart..extraEnd` is the
 * data that follows the tokens in the record (array constants). `row` and `col` are the cell
 * the formula is read for: the relative references of a shared formula count from it.
 *
 * Returns undefined when the tokens cannot be turned into text for certain; the reason is
 * counted in ctx.failures.
 */
export function formulaText(
  ctx: FormulaContext,
  bytes: Uint8Array,
  start: number,
  end: number,
  extraStart: number,
  extraEnd: number,
  row: number,
  col: number,
): string | undefined {
  if (bytes !== viewedBytes) {
    viewedBytes = bytes;
    viewOfBytes = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }
  const decoder: Decoder = {
    ctx,
    bytes,
    view: viewOfBytes as DataView,
    p: start,
    end,
    extra: extraStart,
    extraEnd,
    row,
    col,
    texts: [],
    precs: [],
  };
  try {
    return decode(decoder);
  } catch (error) {
    if (!(error instanceof Failure)) throw error;
    ctx.failures.set(error.reason, (ctx.failures.get(error.reason) || 0) + 1);
    return undefined;
  }
}

/** Lets go of the stream formulaText() last worked on. */
export function releaseFormulaBytes(): void {
  viewedBytes = undefined;
  viewOfBytes = undefined;
}
