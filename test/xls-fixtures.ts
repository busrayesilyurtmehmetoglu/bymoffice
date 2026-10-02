// Builds synthetic BIFF record streams for the xls reader tests. Everything is
// generated here; no file comes from anywhere else.

import { minimalCfb } from './fixtures';

export type Bytes = number[] | Uint8Array;

export function cat(...parts: Bytes[]): Uint8Array {
  let size = 0;
  for (const part of parts) size += part.length;
  const out = new Uint8Array(size);
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

export const u8 = (value: number): number[] => [value & 0xff];
export const u16 = (value: number): number[] => [value & 0xff, (value >> 8) & 0xff];
export const u32 = (value: number): number[] => [value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff];

export function f64(value: number): number[] {
  const view = new DataView(new ArrayBuffer(8));
  view.setFloat64(0, value, true);
  return [...new Uint8Array(view.buffer)];
}

/** UTF-16LE code units. */
export function wide(text: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < text.length; i++) out.push(text.charCodeAt(i) & 0xff, text.charCodeAt(i) >> 8);
  return out;
}

/** One byte per character; every character has to be below U+0100. */
export function narrow(text: string): number[] {
  const out: number[] = [];
  for (let i = 0; i < text.length; i++) out.push(text.charCodeAt(i));
  return out;
}

const isNarrow = (text: string) => /^[\u0000-ÿ]*$/.test(text);

/** XLUnicodeString: 16-bit count, flags, characters. Compressed when the text allows it. */
export function str16(text: string, forceWide = false): number[] {
  const compressed = !forceWide && isNarrow(text);
  return [...u16(text.length), compressed ? 0 : 1, ...(compressed ? narrow(text) : wide(text))];
}

/** ShortXLUnicodeString: 8-bit count, flags, characters. */
export function str8(text: string, forceWide = false): number[] {
  const compressed = !forceWide && isNarrow(text);
  return [text.length, compressed ? 0 : 1, ...(compressed ? narrow(text) : wide(text))];
}

/** One record: id, length, content. */
export function rec(id: number, ...content: Bytes[]): Uint8Array {
  const data = cat(...content);
  if (data.length > 0xffff) throw new Error('record too long');
  return cat(u16(id), u16(data.length), data);
}

export const bof = (type: number, version = 0x0600): Uint8Array => rec(0x0809, u16(version), u16(type), u16(0x0dbb), u16(0x07cc), u32(0), u32(6));
export const eof = (): Uint8Array => rec(0x000a);

// ---------------------------------------------------------------------------
// Cell records

export const number = (row: number, col: number, value: number, xf = 15) => rec(0x0203, u16(row), u16(col), u16(xf), f64(value));
export const rk = (row: number, col: number, value: number, xf = 15) => rec(0x027e, u16(row), u16(col), u16(xf), u32(value));
export const blank = (row: number, col: number, xf = 15) => rec(0x0201, u16(row), u16(col), u16(xf));
export const labelSst = (row: number, col: number, index: number, xf = 15) => rec(0x00fd, u16(row), u16(col), u16(xf), u32(index));
export const boolErr = (row: number, col: number, value: number, isError: boolean, xf = 15) => rec(0x0205, u16(row), u16(col), u16(xf), [value, isError ? 1 : 0]);

/** RK encodings: an integer, an integer in hundredths, a double, a double in hundredths. */
export const rkInt = (value: number) => ((value << 2) | 2) >>> 0;
export const rkInt100 = (value: number) => ((value << 2) | 3) >>> 0;
export function rkFloat(value: number, hundredths = false): number {
  const view = new DataView(new ArrayBuffer(8));
  view.setFloat64(0, value, true);
  if (view.getUint32(0, true) !== 0 || view.getUint32(4, true) & 3) throw new Error('not an RK number');
  return (view.getUint32(4, true) | (hundredths ? 1 : 0)) >>> 0;
}

export type FormulaResult = number | boolean | { error: number } | { text: string } | 'emptyText';

/** FORMULA, and the STRING record that carries a text result. */
export function formula(row: number, col: number, tokens: Bytes, result: FormulaResult = 0, xf = 15, extra: Bytes = [], flags = 0): Uint8Array {
  let value: number[];
  let after: Uint8Array = new Uint8Array(0);
  if (typeof result === 'number') value = f64(result);
  else if (typeof result === 'boolean') value = [1, 0, result ? 1 : 0, 0, 0, 0, 0xff, 0xff];
  else if (result === 'emptyText') value = [3, 0, 0, 0, 0, 0, 0xff, 0xff];
  else if ('error' in result) value = [2, 0, result.error, 0, 0, 0, 0xff, 0xff];
  else {
    value = [0, 0, 0, 0, 0, 0, 0xff, 0xff];
    after = rec(0x0207, str16(result.text));
  }
  return cat(rec(0x0006, u16(row), u16(col), u16(xf), value, u16(flags), u32(0), u16(tokens.length), tokens, extra), after);
}

export const shrFmla = (top: number, bottom: number, left: number, right: number, tokens: Bytes) => rec(0x04bc, u16(top), u16(bottom), [left, right, 0, 0], u16(tokens.length), tokens);
export const arrayRec = (top: number, bottom: number, left: number, right: number, tokens: Bytes, extra: Bytes = []) =>
  rec(0x0221, u16(top), u16(bottom), [left, right], u16(0), u32(0), u16(tokens.length), tokens, extra);

// ---------------------------------------------------------------------------
// Formula tokens (BIFF8)

const colField = (col: number, rowRelative: boolean, colRelative: boolean) => (col & 0xff) | (colRelative ? 0x4000 : 0) | (rowRelative ? 0x8000 : 0);

export const ptg = {
  exp: (row: number, col: number) => [0x01, ...u16(row), ...u16(col)],
  tbl: (row: number, col: number) => [0x02, ...u16(row), ...u16(col)],
  add: [0x03],
  sub: [0x04],
  mul: [0x05],
  div: [0x06],
  power: [0x07],
  concat: [0x08],
  lt: [0x09],
  le: [0x0a],
  eq: [0x0b],
  ge: [0x0c],
  gt: [0x0d],
  ne: [0x0e],
  isect: [0x0f],
  union: [0x10],
  range: [0x11],
  uplus: [0x12],
  uminus: [0x13],
  percent: [0x14],
  paren: [0x15],
  missArg: [0x16],
  str: (text: string) => [0x17, ...str8(text)],
  attrVolatile: [0x19, 0x01, 0, 0],
  attrIf: (distance = 0) => [0x19, 0x02, ...u16(distance)],
  attrChoose: (offsets: number[]) => [0x19, 0x04, ...u16(offsets.length - 1), ...offsets.flatMap((offset) => u16(offset))],
  attrGoto: (distance = 0) => [0x19, 0x08, ...u16(distance)],
  attrSum: [0x19, 0x10, 0, 0],
  attrSpace: (type = 0, count = 1) => [0x19, 0x40, type, count],
  err: (code: number) => [0x1c, code],
  bool: (value: boolean) => [0x1d, value ? 1 : 0],
  int: (value: number) => [0x1e, ...u16(value)],
  num: (value: number) => [0x1f, ...f64(value)],
  array: (cls = 0x60) => [cls, 0, 0, 0, 0, 0, 0, 0],
  func: (index: number, cls = 0x41) => [cls, ...u16(index)],
  funcVar: (index: number, args: number, cls = 0x42) => [cls, args, ...u16(index)],
  name: (index: number, cls = 0x43) => [cls, ...u16(index), 0, 0],
  /** Relative unless told otherwise, as a formula typed without "$" is. */
  ref: (row: number, col: number, rowRelative = true, colRelative = true, cls = 0x44) => [cls, ...u16(row), ...u16(colField(col, rowRelative, colRelative))],
  area: (r1: number, r2: number, c1: number, c2: number, relative = true, cls = 0x25) => [cls, ...u16(r1), ...u16(r2), ...u16(colField(c1, relative, relative)), ...u16(colField(c2, relative, relative))],
  memArea: (size: number) => [0x26, 0, 0, 0, 0, ...u16(size)],
  memErr: (size: number) => [0x27, 0, 0, 0, 0, ...u16(size)],
  memFunc: (size: number) => [0x29, ...u16(size)],
  refErr: [0x2a, 0, 0, 0, 0],
  areaErr: [0x2b, 0, 0, 0, 0, 0, 0, 0, 0],
  /** Offsets from the cell of the formula, for the parts that are relative. */
  refN: (row: number, col: number, rowRelative = true, colRelative = true, cls = 0x4c) => [cls, ...u16(row), ...u16(colField(col, rowRelative, colRelative))],
  areaN: (r1: number, r2: number, c1: number, c2: number, cls = 0x2d) => [cls, ...u16(r1), ...u16(r2), ...u16(colField(c1, true, true)), ...u16(colField(c2, true, true))],
  nameX: (xti: number, index: number, cls = 0x39) => [cls, ...u16(xti), ...u16(index), 0, 0],
  ref3d: (xti: number, row: number, col: number, relative = false, cls = 0x3a) => [cls, ...u16(xti), ...u16(row), ...u16(colField(col, relative, relative))],
  area3d: (xti: number, r1: number, r2: number, c1: number, c2: number, relative = false, cls = 0x3b) => [cls, ...u16(xti), ...u16(r1), ...u16(r2), ...u16(colField(c1, relative, relative)), ...u16(colField(c2, relative, relative))],
  refErr3d: (xti: number) => [0x3c, ...u16(xti), 0, 0, 0, 0],
};

/** Joins tokens. */
export const tokens = (...parts: number[][]): number[] => parts.flat();

// ---------------------------------------------------------------------------
// Style records

export interface FontSpec {
  name?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  strike?: boolean;
  underline?: number;
  color?: number;
  script?: number;
  family?: number;
  charset?: number;
}

export function font(spec: FontSpec = {}): Uint8Array {
  const flags = (spec.italic ? 0x02 : 0) | (spec.strike ? 0x08 : 0);
  return rec(0x0031, u16(Math.round((spec.size ?? 10) * 20)), u16(flags), u16(spec.color ?? 0x7fff), u16(spec.bold ? 700 : 400), u16(spec.script ?? 0), [spec.underline ?? 0, spec.family ?? 0, spec.charset ?? 0, 0], str8(spec.name ?? 'Arial'));
}

export const format = (id: number, code: string) => rec(0x041e, u16(id), str16(code));

export interface XfSpec {
  font?: number;
  format?: number;
  style?: boolean;
  parent?: number;
  locked?: boolean;
  hidden?: boolean;
  horizontal?: number;
  vertical?: number;
  wrap?: boolean;
  rotation?: number;
  indent?: number;
  shrink?: boolean;
  /** Line styles and colours: left, right, top, bottom. */
  lines?: [number, number, number, number];
  lineColors?: [number, number, number, number];
  diagonal?: { style: number; color: number; up?: boolean; down?: boolean };
  pattern?: number;
  fg?: number;
  bg?: number;
}

export function xf(spec: XfSpec = {}): Uint8Array {
  const type = ((spec.locked ?? true) ? 1 : 0) | (spec.hidden ? 2 : 0) | (spec.style ? 4 : 0) | ((spec.style ? 0xfff : (spec.parent ?? 0)) << 4);
  const align = (spec.horizontal ?? 0) | (spec.wrap ? 8 : 0) | ((spec.vertical ?? 2) << 4);
  const lines = spec.lines ?? [0, 0, 0, 0];
  const colors = spec.lineColors ?? [0, 0, 0, 0];
  const diagonal = spec.diagonal;
  const border1 = (lines[0] | (lines[1] << 4) | (lines[2] << 8) | (lines[3] << 12) | (colors[0] << 16) | (colors[1] << 23) | (diagonal?.down ? 0x40000000 : 0) | (diagonal?.up ? 0x80000000 : 0)) >>> 0;
  const border2 = (colors[2] | (colors[3] << 7) | ((diagonal?.color ?? 0) << 14) | ((diagonal?.style ?? 0) << 21) | ((spec.pattern ?? 0) << 26)) >>> 0;
  const fill = (spec.fg ?? 64) | ((spec.bg ?? 65) << 7);
  return rec(0x00e0, u16(spec.font ?? 0), u16(spec.format ?? 0), u16(type), [align, spec.rotation ?? 0, (spec.indent ?? 0) | (spec.shrink ? 0x10 : 0), 0], u32(border1), u32(border2), u16(fill));
}

/** The fonts and XF records every workbook written by Excel starts with: 16 XF records, of which 0..14 are style XFs. */
export function defaultStyles(): Uint8Array {
  const parts: Uint8Array[] = [];
  for (let i = 0; i < 4; i++) parts.push(font());
  for (let i = 0; i < 15; i++) parts.push(xf({ style: true }));
  parts.push(xf({ parent: 0 }));
  return cat(...parts);
}

// ---------------------------------------------------------------------------
// Workbook assembly

export interface SheetSpec {
  name: string;
  /** Records between BOF and EOF. */
  records?: Bytes;
  /** 0 visible, 1 hidden, 2 very hidden. */
  visibility?: number;
  /** BOUNDSHEET type: 0 worksheet, 1 macro sheet, 2 chart. */
  type?: number;
  /** Type field of the BOF record; follows `type` when omitted. */
  bofType?: number;
}

export interface WorkbookSpec {
  /** Records of the globals before the sheet list (fonts, formats, XF records, ...). Default: defaultStyles(). */
  head?: Bytes;
  /** Records of the globals after the sheet list (SUPBOOK, EXTERNSHEET, NAME, SST, ...). */
  tail?: Bytes;
  sheets: SheetSpec[];
  /** 0x0600 (BIFF8) unless stated. */
  version?: number;
}

const BOF_TYPES: Record<number, number> = { 0: 0x0010, 1: 0x0040, 2: 0x0020, 6: 0x0006 };

function boundSheet(offset: number, sheet: SheetSpec, biff8: boolean): Uint8Array {
  const name = biff8 ? str8(sheet.name) : [sheet.name.length, ...narrow(sheet.name)];
  return rec(0x0085, u32(offset), [sheet.visibility ?? 0, sheet.type ?? 0], name);
}

/** The bytes of a Workbook stream: globals, then one substream per sheet, with the sheet offsets filled in. */
export function workbookStream(spec: WorkbookSpec): Uint8Array {
  const version = spec.version ?? 0x0600;
  const biff8 = version === 0x0600;
  const head = cat(bof(0x0005, version), spec.head ?? defaultStyles());
  const tail = cat(spec.tail ?? [], eof());
  const bodies = spec.sheets.map((sheet) => cat(bof(sheet.bofType ?? BOF_TYPES[sheet.type ?? 0], version), sheet.records ?? [], eof()));
  // The sheet list has the same size whatever the offsets are, so it can be measured first.
  const listSize = spec.sheets.reduce((sum, sheet) => sum + boundSheet(0, sheet, biff8).length, 0);
  let offset = head.length + listSize + tail.length;
  const list: Uint8Array[] = [];
  spec.sheets.forEach((sheet, i) => {
    list.push(boundSheet(offset, sheet, biff8));
    offset += bodies[i].length;
  });
  return cat(head, ...list, tail, ...bodies);
}

/** A compound file with the stream as "Workbook". The stream is padded with zeros to the size the fixture needs. */
export function xlsFile(stream: Uint8Array): Uint8Array {
  const size = 9 * 512;
  if (stream.length > size) throw new Error('stream too long for the fixture');
  const padded = new Uint8Array(size);
  padded.set(stream);
  return minimalCfb(padded, new Uint8Array(0));
}

/** SST from complete strings (each already encoded as an XLUnicodeRichExtendedString). */
export const sst = (...strings: Bytes[]) => rec(0x00fc, u32(strings.length), u32(strings.length), ...strings);

export const externSheet = (...entries: [book: number, first: number, last: number][]) => rec(0x0017, u16(entries.length), ...entries.map(([book, first, last]) => [...u16(book), ...u16(first), ...u16(last)]));
export const supBookSelf = (sheetCount: number) => rec(0x01ae, u16(sheetCount), [0x01, 0x04]);
export const supBookAddin = () => rec(0x01ae, u16(1), [0x01, 0x3a]);
export const supBookExternal = (path: string, sheetNames: string[]) => rec(0x01ae, u16(sheetNames.length), str16(path), ...sheetNames.map((name) => str16(name)));
export const externName = (name: string) => rec(0x0023, u16(0), u32(0), str8(name), [0x02, 0x00, 0x1c, 0x17]);

export interface NameSpec {
  name: string;
  tokens: Bytes;
  /** 1-based sheet the name is local to. */
  sheet?: number;
  hidden?: boolean;
  /** Number of a built-in name; `name` is then ignored. */
  builtin?: number;
  extra?: Bytes;
}

export function definedName(spec: NameSpec): Uint8Array {
  const text = spec.builtin === undefined ? spec.name : String.fromCharCode(spec.builtin);
  const flags = (spec.hidden ? 0x0001 : 0) | (spec.builtin === undefined ? 0 : 0x0020);
  const compressed = isNarrow(text);
  return rec(0x0018, u16(flags), [0, text.length], u16(spec.tokens.length), u16(0), u16(spec.sheet ?? 0), [0, 0, 0, 0], [compressed ? 0 : 1], compressed ? narrow(text) : wide(text), spec.tokens, spec.extra ?? []);
}
