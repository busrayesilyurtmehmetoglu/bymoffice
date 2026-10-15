// Formula text -> syntax tree.
//
// One pass over the text, by character code, without regular expressions. Operators are
// read by precedence climbing; the only recursion is into parentheses, function arguments
// and array constants, and it is counted, so no text can overflow the stack. A tree is
// never changed once it is built: the same tree is handed out again for the same text.

import { MAX_COLS, MAX_ROWS, type CellRange } from '../model/sheet/types';
import type { ErrorCode } from './values';

export type Node =
  | { type: 'number'; value: number }
  | { type: 'string'; value: string }
  | { type: 'boolean'; value: boolean }
  | { type: 'error'; code: ErrorCode }
  | {
      type: 'ref';
      /** Sheet name as written, without quotes. For a reference into another workbook, the whole prefix. */
      sheet?: string;
      /** Last sheet of a 3D reference ("Sheet1:Sheet3!A1"). */
      sheetEnd?: string;
      /** Corners in order: top <= bottom, left <= right. */
      range: CellRange;
      /** Which sides were written with "$". */
      abs: { top: boolean; left: boolean; bottom: boolean; right: boolean };
      /** "1:3": every column of the rows. */
      wholeRows?: boolean;
      /** "A:C": every row of the columns. */
      wholeCols?: boolean;
      /** Refers to another workbook. */
      external?: boolean;
    }
  | { type: 'name'; name: string; sheet?: string }
  /**
   * `null` is an argument left empty, as in IF(A1,,2). Three operators of dynamic arrays are
   * calls as well, the way files store them: "A1#" is ANCHORARRAY(A1), "@A1:A3" is
   * SINGLE(A1:A3), and a call of what a call gives ("LAMBDA(x,x+1)(2)") is APPLY_CALL with
   * the callee as its first argument.
   */
  | { type: 'call'; name: string; args: (Node | null)[] }
  /** "%" is written after its operand. */
  | { type: 'unary'; op: '-' | '+' | '%'; operand: Node }
  /** ":" makes a range, "," a union (only inside parentheses), " " an intersection. */
  | { type: 'binary'; op: BinaryOp; left: Node; right: Node }
  /** {1,2;3,4}: constants only. */
  | { type: 'array'; rows: Node[][] };

export type BinaryOp = '+' | '-' | '*' | '/' | '^' | '&' | '=' | '<>' | '<' | '>' | '<=' | '>=' | ':' | ',' | ' ';

export type RefNode = Extract<Node, { type: 'ref' }>;

/**
 * Name of the call node that calls what another call gives: LAMBDA(x,x+1)(2). It is no name
 * that can be typed, so no function of the library or of a file can be called that.
 */
export const APPLY_CALL = '()';

/** Strips the prefixes files give to names of functions newer than the file format, and to the parameters of LET and LAMBDA. */
function unprefixed(upper: string): string {
  for (let guard = 0; guard < 4; guard++) {
    if (upper.lastIndexOf('_XLFN.', 0) === 0 || upper.lastIndexOf('_XLWS.', 0) === 0 || upper.lastIndexOf('_XLPM.', 0) === 0) upper = upper.slice(6);
    else break;
  }
  return upper;
}

export class FormulaParseError extends Error {
  /** Offset into the text given to parseFormula() where the problem was found. */
  position: number;

  constructor(message: string, position: number) {
    super(message);
    this.name = 'FormulaParseError';
    this.position = position;
  }
}

/** Excel's own limit on the length of a formula. */
export const MAX_FORMULA_LENGTH = 8192;

/** Parentheses, calls and array constants inside one another; also the longest run of signs or "%". */
const MAX_DEPTH = 256;

const MAX_CACHED = 8192;

const ERROR_CODES: ErrorCode[] = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'];

// What a word spells when it is (one side of) a reference.
const K_CELL = 1;
const K_COL = 2;
const K_ROW = 3;

interface Word {
  kind: number;
  row: number;
  col: number;
  absRow: boolean;
  absCol: boolean;
}

interface Parser {
  s: string;
  n: number;
  i: number;
  depth: number;
}

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;
const isDigit = (c: number) => c >= 48 && c <= 57;

/** Characters of names and references: letters, digits, "_", ".", "$", "\", "?" and everything outside ASCII. */
function isWordChar(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  if (c >= 65) return c <= 90 || c === 95 || c === 92;
  if (c >= 48) return c <= 57 || c === 63;
  return c === 36 || c === 46;
}

/** A name starts with a letter, "_", "\" or a character outside ASCII. */
function isNameStart(c: number): boolean {
  if (c >= 128) return true;
  if (c >= 97) return c <= 122;
  return (c >= 65 && c <= 90) || c === 95 || c === 92;
}

function skipWord(s: string, i: number): number {
  const n = s.length;
  while (i < n && isWordChar(s.charCodeAt(i))) i++;
  return i;
}

function skipSpaces(p: Parser): boolean {
  const start = p.i;
  while (p.i < p.n && isSpace(p.s.charCodeAt(p.i))) p.i++;
  return p.i > start;
}

function fail(message: string, position: number): never {
  throw new FormulaParseError(message, position);
}

function enter(p: Parser): void {
  if (++p.depth > MAX_DEPTH) fail('The formula is nested too deeply', p.i);
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
  if (i > digitsStart) {
    if (row === 0) return undefined;
    return { kind: hasLetters ? K_CELL : K_ROW, row: row - 1, col: col - 1, absRow, absCol };
  }
  if (!hasLetters || absRow) return undefined;
  return { kind: K_COL, row: 0, col: col - 1, absRow: false, absCol };
}

function refNode(first: Word, second: Word, sheet: string | undefined, sheetEnd: string | undefined, external: boolean): RefNode {
  const kind = first.kind;
  const flipRows = first.row > second.row;
  const flipCols = first.col > second.col;
  const node: RefNode = {
    type: 'ref',
    range:
      kind === K_COL
        ? { top: 0, left: Math.min(first.col, second.col), bottom: MAX_ROWS - 1, right: Math.max(first.col, second.col) }
        : kind === K_ROW
          ? { top: Math.min(first.row, second.row), left: 0, bottom: Math.max(first.row, second.row), right: MAX_COLS - 1 }
          : { top: Math.min(first.row, second.row), left: Math.min(first.col, second.col), bottom: Math.max(first.row, second.row), right: Math.max(first.col, second.col) },
    abs: {
      top: flipRows ? second.absRow : first.absRow,
      left: flipCols ? second.absCol : first.absCol,
      bottom: flipRows ? first.absRow : second.absRow,
      right: flipCols ? first.absCol : second.absCol,
    },
  };
  if (sheet !== undefined) node.sheet = sheet;
  if (sheetEnd !== undefined) node.sheetEnd = sheetEnd;
  if (kind === K_COL) node.wholeCols = true;
  if (kind === K_ROW) node.wholeRows = true;
  if (external) node.external = true;
  return node;
}

/** The error constant at `i`, or undefined. */
function errorAt(s: string, i: number): ErrorCode | undefined {
  for (let k = 0; k < ERROR_CODES.length; k++) {
    const code = ERROR_CODES[k];
    if (s.length - i >= code.length && s.substr(i, code.length).toUpperCase() === code) return code;
  }
  return undefined;
}

/**
 * Reads a reference or a name whose text starts at `start` (after any sheet prefix):
 * "A1", "A:C", "1:3", "TRUE", "Total". `end` is where the word stops.
 */
function parseWordBody(p: Parser, start: number, end: number, sheet: string | undefined, sheetEnd: string | undefined, external: boolean): Node {
  const s = p.s;
  const word = readWord(s, start, end);
  if (word) {
    if (word.kind === K_CELL) {
      p.i = end;
      return refNode(word, word, sheet, sheetEnd, external);
    }
    // A lone "A" or "5" is a name or a number; with a second one after a colon it is whole columns or rows.
    if (s.charCodeAt(end) === 58) {
      const end2 = skipWord(s, end + 1);
      const second = end2 > end + 1 && s.charCodeAt(end2) !== 40 ? readWord(s, end + 1, end2) : undefined;
      if (second && second.kind === word.kind) {
        p.i = end2;
        return refNode(word, second, sheet, sheetEnd, external);
      }
    }
  }
  if (!isNameStart(s.charCodeAt(start))) fail('Expected a reference or a name', start);
  for (let i = start; i < end; i++) if (s.charCodeAt(i) === 36) fail('Unexpected "$"', i);
  if (sheetEnd !== undefined) fail('A name cannot span sheets', start);
  let text = s.slice(start, end);
  p.i = end;
  if (sheet === undefined) {
    if (end - start === 4 && text.toUpperCase() === 'TRUE') return { type: 'boolean', value: true };
    if (end - start === 5 && text.toUpperCase() === 'FALSE') return { type: 'boolean', value: false };
    // "_xlpm.x": a parameter of LET or LAMBDA, the way files store it.
    if (text.charCodeAt(0) === 95 && text.length > 6 && text.substring(0, 6).toUpperCase() === '_XLPM.') text = text.slice(6);
    return { type: 'name', name: text };
  }
  return { type: 'name', name: text, sheet };
}

/** What follows the "!" of a sheet prefix. */
function parseAfterSheet(p: Parser, sheet: string, sheetEnd: string | undefined, external: boolean): Node {
  const s = p.s;
  const start = p.i;
  if (s.charCodeAt(start) === 35) {
    // "Sheet1!#REF!": what is left of a reference whose cells were deleted.
    const code = errorAt(s, start);
    if (!code) fail('Unknown error constant', start);
    p.i = start + (code as string).length;
    return { type: 'error', code: code as ErrorCode };
  }
  const end = skipWord(s, start);
  if (end === start) fail('Expected a reference after the sheet name', start);
  return parseWordBody(p, start, end, sheet, sheetEnd, external);
}

function parseQuotedSheet(p: Parser): Node {
  const s = p.s;
  const start = p.i;
  let from = start + 1;
  let close: number;
  for (;;) {
    close = s.indexOf("'", from);
    if (close < 0) fail('The quote is never closed', start);
    if (s.charCodeAt(close + 1) !== 39) break;
    from = close + 2;
  }
  if (s.charCodeAt(close + 1) !== 33) fail('Expected "!" after the sheet name', close + 1);
  const name = s.slice(start + 1, close).replace(/''/g, "'");
  if (name === '') fail('The sheet name is empty', start);
  p.i = close + 2;
  // Sheet names cannot hold "[", "\" or "/": those come from a workbook name or path.
  if (name.indexOf('[') >= 0 || name.indexOf('\\') >= 0 || name.indexOf('/') >= 0) return parseAfterSheet(p, name, undefined, true);
  const colon = name.indexOf(':');
  if (colon > 0 && colon < name.length - 1) return parseAfterSheet(p, name.slice(0, colon), name.slice(colon + 1), false);
  return parseAfterSheet(p, name, undefined, false);
}

/** "[1]Sheet1!A1", "[1]Sheet1:Sheet3!A1", "[1]!Name": another workbook. */
function parseExternal(p: Parser): Node {
  const s = p.s;
  const start = p.i;
  const close = s.indexOf(']', start);
  if (close < 0) fail('The bracket is never closed', start);
  let end = skipWord(s, close + 1);
  if (s.charCodeAt(end) === 58) {
    const end2 = skipWord(s, end + 1);
    if (end2 > end + 1 && s.charCodeAt(end2) === 33) end = end2;
  }
  // Anything else in brackets belongs to a table reference.
  if (s.charCodeAt(end) !== 33) fail('Structured references are not supported', start);
  p.i = end + 1;
  return parseAfterSheet(p, s.slice(start, end), undefined, true);
}

function parseCall(p: Parser, start: number, end: number): Node {
  const s = p.s;
  // Functions newer than the file format are stored with a prefix.
  const name = unprefixed(s.slice(start, end).toUpperCase());
  if (name === '' || !isNameStart(name.charCodeAt(0))) fail('Expected a function name', start);
  p.i = end;
  return { type: 'call', name, args: parseArguments(p) };
}

/** With "(" current: the arguments of a call up to the ")" that closes them. */
function parseArguments(p: Parser): (Node | null)[] {
  const s = p.s;
  p.i++;
  enter(p);
  const args: (Node | null)[] = [];
  skipSpaces(p);
  if (s.charCodeAt(p.i) === 41) {
    p.i++;
    p.depth--;
    return args;
  }
  for (;;) {
    skipSpaces(p);
    let c = s.charCodeAt(p.i);
    args.push(c === 44 || c === 41 ? null : parseBinary(p, 1));
    skipSpaces(p);
    c = s.charCodeAt(p.i);
    if (c === 44) {
      p.i++;
      continue;
    }
    if (c === 41) {
      p.i++;
      break;
    }
    fail(p.i >= p.n ? 'Missing ")"' : 'Unexpected character', p.i);
  }
  p.depth--;
  return args;
}

function parseWord(p: Parser): Node {
  const s = p.s;
  const start = p.i;
  const end = skipWord(s, start);
  const c = s.charCodeAt(end);
  if (c === 40) return parseCall(p, start, end);
  if (c === 33) {
    p.i = end + 1;
    return parseAfterSheet(p, s.slice(start, end), undefined, false);
  }
  if (c === 58) {
    const end2 = skipWord(s, end + 1);
    if (end2 > end + 1 && s.charCodeAt(end2) === 33) {
      p.i = end2 + 1;
      return parseAfterSheet(p, s.slice(start, end), s.slice(end + 1, end2), false);
    }
  }
  if (c === 91) fail('Structured references are not supported', end);
  return parseWordBody(p, start, end, undefined, undefined, false);
}

/** Whether "1:3" or "$1:$3" starts at `i`. */
function rowRangeAt(s: string, i: number): boolean {
  const end = skipWord(s, i);
  if (s.charCodeAt(end) !== 58) return false;
  const first = readWord(s, i, end);
  if (!first || first.kind !== K_ROW) return false;
  const end2 = skipWord(s, end + 1);
  const second = end2 > end + 1 ? readWord(s, end + 1, end2) : undefined;
  return !!second && second.kind === K_ROW;
}

function parseNumber(p: Parser): Node {
  const s = p.s;
  const start = p.i;
  let i = start;
  while (isDigit(s.charCodeAt(i))) i++;
  let digits = i - start;
  if (s.charCodeAt(i) === 46) {
    const fraction = ++i;
    while (isDigit(s.charCodeAt(i))) i++;
    digits += i - fraction;
  }
  if (digits === 0) fail('Expected a number', start);
  const e = s.charCodeAt(i);
  if (e === 69 || e === 101) {
    let j = i + 1;
    const sign = s.charCodeAt(j);
    if (sign === 43 || sign === 45) j++;
    if (isDigit(s.charCodeAt(j))) {
      while (isDigit(s.charCodeAt(j))) j++;
      i = j;
    }
  }
  const value = Number(s.slice(start, i));
  if (!isFinite(value)) fail('The number is too large', start);
  p.i = i;
  return { type: 'number', value };
}

function parseString(p: Parser): Node {
  const s = p.s;
  const start = p.i;
  let from = start + 1;
  let doubled = false;
  for (;;) {
    const close = s.indexOf('"', from);
    if (close < 0) fail('The text is never closed', start);
    if (s.charCodeAt(close + 1) !== 34) {
      p.i = close + 1;
      const text = s.slice(start + 1, close);
      return { type: 'string', value: doubled ? text.replace(/""/g, '"') : text };
    }
    doubled = true;
    from = close + 2;
  }
}

function parseErrorLiteral(p: Parser): Node {
  const code = errorAt(p.s, p.i);
  if (!code) fail('Unknown error constant', p.i);
  p.i += (code as string).length;
  return { type: 'error', code: code as ErrorCode };
}

function parseArrayItem(p: Parser): Node {
  const s = p.s;
  const c = s.charCodeAt(p.i);
  if (c === 34) return parseString(p);
  if (c === 35) return parseErrorLiteral(p);
  if (c === 45 || c === 43) {
    const sign = p.i++;
    skipSpaces(p);
    const next = s.charCodeAt(p.i);
    if (!isDigit(next) && next !== 46) fail('Expected a number', sign + 1);
    const number = parseNumber(p) as { type: 'number'; value: number };
    return c === 45 ? { type: 'number', value: number.value === 0 ? 0 : -number.value } : number;
  }
  if (isDigit(c) || c === 46) return parseNumber(p);
  const end = skipWord(s, p.i);
  const word = s.slice(p.i, end).toUpperCase();
  if (word === 'TRUE' || word === 'FALSE') {
    p.i = end;
    return { type: 'boolean', value: word === 'TRUE' };
  }
  return fail('An array constant holds only numbers, text, booleans and errors', p.i);
}

function parseArray(p: Parser): Node {
  const s = p.s;
  const start = p.i++;
  const rows: Node[][] = [];
  let row: Node[] = [];
  for (;;) {
    skipSpaces(p);
    row.push(parseArrayItem(p));
    skipSpaces(p);
    const c = s.charCodeAt(p.i);
    if (c === 44) p.i++;
    else if (c === 59) {
      p.i++;
      rows.push(row);
      row = [];
    } else if (c === 125) {
      p.i++;
      rows.push(row);
      break;
    } else fail(p.i >= p.n ? 'Missing "}"' : 'Unexpected character', p.i);
  }
  for (let k = 1; k < rows.length; k++) if (rows[k].length !== rows[0].length) fail('The rows of an array constant differ in length', start);
  return { type: 'array', rows };
}

function parsePrimary(p: Parser): Node {
  const s = p.s;
  if (p.i >= p.n) fail('The formula ends too early', p.i);
  const c = s.charCodeAt(p.i);
  if (isDigit(c)) {
    if (rowRangeAt(s, p.i)) return parseWordBody(p, p.i, skipWord(s, p.i), undefined, undefined, false);
    return parseNumber(p);
  }
  if (c === 46) return parseNumber(p);
  if (c === 34) return parseString(p);
  if (c === 35) return parseErrorLiteral(p);
  if (c === 39) return parseQuotedSheet(p);
  if (c === 91) return parseExternal(p);
  if (c === 123) return parseArray(p);
  if (c === 40) {
    p.i++;
    enter(p);
    let node = parseBinary(p, 1);
    skipSpaces(p);
    // A comma inside parentheses that are not a function's joins references.
    while (s.charCodeAt(p.i) === 44) {
      p.i++;
      node = { type: 'binary', op: ',', left: node, right: parseBinary(p, 1) };
      skipSpaces(p);
    }
    if (s.charCodeAt(p.i) !== 41) fail(p.i >= p.n ? 'Missing ")"' : 'Unexpected character', p.i);
    p.i++;
    p.depth--;
    return node;
  }
  if (isNameStart(c) || c === 36) return parseWord(p);
  return fail('Unexpected character', p.i);
}

function sameSheetName(a: string | undefined, b: string | undefined): boolean {
  return a === b || (a !== undefined && b !== undefined && a.toUpperCase() === b.toUpperCase());
}

/** "A1:B2" and "Sheet1!A1:Sheet1!B2" are one reference; anything else joined by a colon is worked out when calculating. */
function joinRange(left: Node, right: Node): Node {
  // "Sheet1!A1:B2#": the spill of B2 of the sheet the range is on, as B2 of "Sheet1!A1:B2" is.
  if (left.type === 'ref' && left.sheet !== undefined && left.sheetEnd === undefined && !left.external && right.type === 'call' && right.name === 'ANCHORARRAY') {
    const inner = right.args[0];
    if (inner && inner.type === 'ref' && inner.sheet === undefined) right = { type: 'call', name: 'ANCHORARRAY', args: [{ ...inner, sheet: left.sheet }] };
  }
  if (
    left.type === 'ref' &&
    right.type === 'ref' &&
    !left.external &&
    !right.external &&
    right.sheetEnd === undefined &&
    (right.sheet === undefined || (left.sheetEnd === undefined && sameSheetName(left.sheet, right.sheet)))
  ) {
    const a = left.range;
    const b = right.range;
    const node: RefNode = {
      type: 'ref',
      range: { top: Math.min(a.top, b.top), left: Math.min(a.left, b.left), bottom: Math.max(a.bottom, b.bottom), right: Math.max(a.right, b.right) },
      abs: {
        top: b.top < a.top ? right.abs.top : left.abs.top,
        left: b.left < a.left ? right.abs.left : left.abs.left,
        bottom: b.bottom >= a.bottom ? right.abs.bottom : left.abs.bottom,
        right: b.right >= a.right ? right.abs.right : left.abs.right,
      },
    };
    if (left.sheet !== undefined) node.sheet = left.sheet;
    if (left.sheetEnd !== undefined) node.sheetEnd = left.sheetEnd;
    if (left.wholeCols && right.wholeCols) node.wholeCols = true;
    if (left.wholeRows && right.wholeRows) node.wholeRows = true;
    return node;
  }
  return { type: 'binary', op: ':', left, right };
}

/** What follows an operand without a space: "#" of a spill reference, the arguments of a call of a call. */
function parsePostfix(p: Parser, node: Node): Node {
  const s = p.s;
  for (let guard = 0; ; guard++) {
    if (guard >= MAX_DEPTH) fail('The formula is nested too deeply', p.i);
    const c = s.charCodeAt(p.i);
    // "A1#": the cells a dynamic array formula in A1 spills into.
    if (c === 35 && ((node.type === 'ref' && !node.wholeRows && !node.wholeCols && node.range.top === node.range.bottom && node.range.left === node.range.right) || node.type === 'name')) {
      p.i++;
      node = { type: 'call', name: 'ANCHORARRAY', args: [node] };
    } else if (c === 40 && node.type === 'call') {
      // LAMBDA(x,x+1)(2): what the call gives is called in turn.
      const args = parseArguments(p);
      node = { type: 'call', name: APPLY_CALL, args: [node as Node | null].concat(args) };
    } else return node;
  }
}

function parseOperand(p: Parser): Node {
  // "@": implicit intersection, asked for by name; it takes what follows up to the next operator.
  if (p.s.charCodeAt(p.i) === 64) {
    p.i++;
    enter(p);
    skipSpaces(p);
    const node: Node = { type: 'call', name: 'SINGLE', args: [parseRange(p)] };
    p.depth--;
    return node;
  }
  return parsePostfix(p, parsePrimary(p));
}

function parseRange(p: Parser): Node {
  let left = parseOperand(p);
  while (p.s.charCodeAt(p.i) === 58) {
    p.i++;
    left = joinRange(left, parseOperand(p));
  }
  return left;
}

/** What the reference operators can work on. */
function isRefLike(node: Node): boolean {
  return node.type === 'ref' || node.type === 'name' || node.type === 'call' || (node.type === 'binary' && (node.op === ':' || node.op === ' ' || node.op === ','));
}

/** A space between two references is the intersection operator; everywhere else it means nothing. */
function parseIntersection(p: Parser): Node {
  let left = parseRange(p);
  for (;;) {
    if (!skipSpaces(p) || p.i >= p.n || !isRefLike(left)) return left;
    const c = p.s.charCodeAt(p.i);
    const operand = isNameStart(c) || c === 36 || c === 39 || c === 40 || c === 91 || c === 64 || (isDigit(c) && rowRangeAt(p.s, p.i));
    if (!operand) return left;
    const at = p.i;
    const right = parseRange(p);
    if (!isRefLike(right)) fail('Expected an operator', at);
    left = { type: 'binary', op: ' ', left, right };
  }
}

/** Signs before the operand bind tighter than "%" after it, and both tighter than "^": -2^2 is 4. */
function parseUnary(p: Parser): Node {
  const s = p.s;
  skipSpaces(p);
  let signs = '';
  for (;;) {
    const c = s.charCodeAt(p.i);
    if (c !== 45 && c !== 43) break;
    if (signs.length >= MAX_DEPTH) fail('Too many signs', p.i);
    signs += c === 45 ? '-' : '+';
    p.i++;
    skipSpaces(p);
  }
  let node = parseIntersection(p);
  for (let k = signs.length - 1; k >= 0; k--) node = { type: 'unary', op: signs[k] as '-' | '+', operand: node };
  let percents = 0;
  while (s.charCodeAt(p.i) === 37) {
    if (++percents > MAX_DEPTH) fail('Too many percent signs', p.i);
    p.i++;
    node = { type: 'unary', op: '%', operand: node };
    skipSpaces(p);
  }
  return node;
}

// Precedence of the value operators, lowest first: comparisons, "&", "+ -", "* /", "^".
// All of them group from the left, "^" included (2^3^2 is 64).
function parseBinary(p: Parser, minPrecedence: number): Node {
  const s = p.s;
  let left = parseUnary(p);
  for (;;) {
    skipSpaces(p);
    const c = s.charCodeAt(p.i);
    let op: BinaryOp;
    let precedence: number;
    let length = 1;
    if (c === 43) (op = '+'), (precedence = 3);
    else if (c === 45) (op = '-'), (precedence = 3);
    else if (c === 42) (op = '*'), (precedence = 4);
    else if (c === 47) (op = '/'), (precedence = 4);
    else if (c === 94) (op = '^'), (precedence = 5);
    else if (c === 38) (op = '&'), (precedence = 2);
    else if (c === 61) (op = '='), (precedence = 1);
    else if (c === 60) {
      const next = s.charCodeAt(p.i + 1);
      precedence = 1;
      if (next === 62) (op = '<>'), (length = 2);
      else if (next === 61) (op = '<='), (length = 2);
      else op = '<';
    } else if (c === 62) {
      precedence = 1;
      if (s.charCodeAt(p.i + 1) === 61) (op = '>='), (length = 2);
      else op = '>';
    } else return left;
    if (precedence < minPrecedence) return left;
    p.i += length;
    left = { type: 'binary', op, left, right: parseBinary(p, precedence + 1) };
  }
}

let cache = new Map<string, Node>();

/**
 * Parses a formula; the leading "=" is optional. Throws FormulaParseError for text that is
 * not a formula this engine understands, structured (table) references included. The tree
 * returned is shared between callers and must not be changed.
 */
export function parseFormula(text: string): Node {
  if (typeof text !== 'string') throw new FormulaParseError('The formula is not text', 0);
  const cached = cache.get(text);
  if (cached) return cached;
  if (text.length > MAX_FORMULA_LENGTH) throw new FormulaParseError(`The formula is longer than ${MAX_FORMULA_LENGTH} characters`, MAX_FORMULA_LENGTH);
  const p: Parser = { s: text, n: text.length, i: 0, depth: 0 };
  skipSpaces(p);
  if (text.charCodeAt(p.i) === 61) p.i++;
  skipSpaces(p);
  if (p.i >= p.n) fail('The formula is empty', p.i);
  const node = parseBinary(p, 1);
  skipSpaces(p);
  if (p.i < p.n) fail('Unexpected character', p.i);
  if (cache.size >= MAX_CACHED) cache = new Map();
  cache.set(text, node);
  return node;
}
