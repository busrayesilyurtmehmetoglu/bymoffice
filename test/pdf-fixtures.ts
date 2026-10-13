// Checking PDF files the engine wrote, without the engine: a reader written
// from the PDF specification (it shares no code with src/pdf and inflates
// with node:zlib), and the command line tools of Poppler where the machine
// has them.

import { execFileSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { test as nodeTest } from 'node:test';
import { inflateSync } from 'node:zlib';

// ---------------------------------------------------------------------------
// The reader

export class Name {
  constructor(readonly name: string) {}
}
export class Ref {
  constructor(readonly id: number) {}
}
export class Str {
  constructor(
    readonly bytes: Uint8Array,
    readonly hex: boolean,
  ) {}
}
export type Value = number | boolean | null | Name | Ref | Str | Value[] | Dict;
export interface Dict {
  [key: string]: Value;
}
export interface Obj {
  id: number;
  offset: number;
  value: Value;
  /** The bytes of the stream as they are in the file. */
  stream?: Uint8Array;
}

const WHITE = [0, 9, 10, 12, 13, 32];
const DELIMITERS = '()<>[]{}/%';

const isWhite = (code: number): boolean => WHITE.indexOf(code) >= 0;
const isDelimiter = (code: number): boolean => DELIMITERS.indexOf(String.fromCharCode(code)) >= 0;

function ascii(bytes: Uint8Array, from: number, to: number): string {
  let out = '';
  for (let i = from; i < to && i < bytes.length; i++) out += String.fromCharCode(bytes[i]);
  return out;
}

class Parser {
  at: number;
  constructor(
    readonly bytes: Uint8Array,
    at: number,
  ) {
    this.at = at;
  }

  fail(message: string): never {
    throw new Error(`${message} at offset ${this.at}`);
  }

  skipWhite(): void {
    const bytes = this.bytes;
    for (;;) {
      while (this.at < bytes.length && isWhite(bytes[this.at])) this.at++;
      if (bytes[this.at] !== 0x25) return;
      while (this.at < bytes.length && bytes[this.at] !== 10 && bytes[this.at] !== 13) this.at++;
    }
  }

  /** A run of regular characters. */
  word(): string {
    const start = this.at;
    while (this.at < this.bytes.length && !isWhite(this.bytes[this.at]) && !isDelimiter(this.bytes[this.at])) this.at++;
    return ascii(this.bytes, start, this.at);
  }

  expect(text: string): void {
    if (ascii(this.bytes, this.at, this.at + text.length) !== text) this.fail(`Expected "${text}"`);
    this.at += text.length;
  }

  value(depth = 0): Value {
    if (depth > 64) this.fail('Nested too deeply');
    this.skipWhite();
    const bytes = this.bytes;
    const code = bytes[this.at];
    if (code === undefined) this.fail('The file ends inside an object');
    if (code === 0x2f) {
      this.at++;
      const start = this.at;
      while (this.at < bytes.length && !isWhite(bytes[this.at]) && !isDelimiter(bytes[this.at])) this.at++;
      const raw = ascii(bytes, start, this.at);
      if (/#(?![0-9a-fA-F]{2})/.test(raw)) this.fail('A name with a bad # escape');
      // The escapes are UTF-8 bytes.
      const out: number[] = [];
      for (let i = 0; i < raw.length; i++) {
        if (raw.charAt(i) === '#') {
          out.push(parseInt(raw.substr(i + 1, 2), 16));
          i += 2;
        } else out.push(raw.charCodeAt(i));
      }
      return new Name(Buffer.from(out).toString('utf8'));
    }
    if (code === 0x28) return this.literal();
    if (code === 0x3c) {
      if (bytes[this.at + 1] === 0x3c) {
        this.at += 2;
        const dict: Dict = {};
        for (;;) {
          this.skipWhite();
          if (bytes[this.at] === 0x3e && bytes[this.at + 1] === 0x3e) {
            this.at += 2;
            return dict;
          }
          const key = this.value(depth + 1);
          if (!(key instanceof Name)) this.fail('A dictionary key that is no name');
          if (Object.prototype.hasOwnProperty.call(dict, (key as Name).name)) this.fail(`The key /${(key as Name).name} is there twice`);
          dict[(key as Name).name] = this.value(depth + 1);
        }
      }
      this.at++;
      const start = this.at;
      while (this.at < bytes.length && bytes[this.at] !== 0x3e) this.at++;
      if (this.at >= bytes.length) this.fail('A hexadecimal string without an end');
      const digits = ascii(bytes, start, this.at).replace(/\s+/g, '');
      this.at++;
      if (!/^[0-9a-fA-F]*$/.test(digits)) this.fail('A hexadecimal string with other characters');
      const padded = digits.length % 2 ? digits + '0' : digits;
      const out = new Uint8Array(padded.length / 2);
      for (let i = 0; i < out.length; i++) out[i] = parseInt(padded.substr(i * 2, 2), 16);
      return new Str(out, true);
    }
    if (code === 0x5b) {
      this.at++;
      const out: Value[] = [];
      for (;;) {
        this.skipWhite();
        if (bytes[this.at] === 0x5d) {
          this.at++;
          return out;
        }
        if (this.at >= bytes.length) this.fail('An array without an end');
        out.push(this.value(depth + 1));
      }
    }
    const start = this.at;
    const word = this.word();
    if (word === 'true') return true;
    if (word === 'false') return false;
    if (word === 'null') return null;
    if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) {
      this.at = start;
      this.fail(`Unexpected "${word || String.fromCharCode(code)}"`);
    }
    // "12 0 R"?
    if (/^\d+$/.test(word)) {
      const back = this.at;
      this.skipWhite();
      const generation = this.word();
      if (/^\d+$/.test(generation)) {
        this.skipWhite();
        if (bytes[this.at] === 0x52 && (this.at + 1 >= bytes.length || isWhite(bytes[this.at + 1]) || isDelimiter(bytes[this.at + 1]))) {
          this.at++;
          return new Ref(Number(word));
        }
      }
      this.at = back;
    }
    return Number(word);
  }

  private literal(): Str {
    const bytes = this.bytes;
    this.at++;
    const out: number[] = [];
    let depth = 1;
    for (;;) {
      if (this.at >= bytes.length) this.fail('A string without an end');
      const code = bytes[this.at++];
      if (code === 0x5c) {
        const next = bytes[this.at++];
        if (next === 0x6e) out.push(10);
        else if (next === 0x72) out.push(13);
        else if (next === 0x74) out.push(9);
        else if (next === 0x62) out.push(8);
        else if (next === 0x66) out.push(12);
        else if (next >= 0x30 && next <= 0x37) {
          let octal = next - 0x30;
          for (let i = 0; i < 2 && bytes[this.at] >= 0x30 && bytes[this.at] <= 0x37; i++) octal = octal * 8 + (bytes[this.at++] - 0x30);
          out.push(octal & 255);
        } else if (next === 10) {
          // A line that goes on.
        } else if (next === 13) {
          if (bytes[this.at] === 10) this.at++;
        } else out.push(next);
      } else if (code === 0x28) {
        depth++;
        out.push(code);
      } else if (code === 0x29) {
        depth--;
        if (depth === 0) return new Str(Uint8Array.from(out), false);
        out.push(code);
      } else out.push(code);
    }
  }
}

export interface CheckedText {
  text: string;
  /** Start of the baseline in points from the bottom left corner of the page. */
  x: number;
  y: number;
  /** The text matrix. */
  matrix: number[];
  size: number;
  /** Horizontal scaling in percent. */
  scaling: number;
  /** Text rendering mode. */
  mode: number;
  /** Characters (codes of the font) the string holds. */
  codes: number;
  /** The advance of the string in points: the widths of its glyphs and the spacing between them, times the size, the scaling and the stretch of the text matrix. */
  advance: number;
  /** Character spacing (Tc), in units of text space. */
  spacing: number;
  /** /BaseFont of the font the string is written in. */
  font: string;
  /** The colour the text is filled with: grey or RGB, 0 to 1. */
  fill: number[];
  /** The glyph numbers of the string. */
  glyphs: number[];
  /** Where each glyph starts, from the start of the string, in units of text space (before the text matrix). */
  starts: number[];
}

/** Something a page fills or strokes. */
export interface CheckedShape {
  /** "f", "f*", "S" or the like. */
  operator: string;
  /** The operators that made the path, with their numbers: ["re", x, y, w, h], ["m", x, y], ... */
  path: (string | number)[][];
  /** The colour of the fill (for a fill) or of the stroke: grey or RGB, 0 to 1. */
  colour: number[];
  /** Opacity, from the graphics state the page chose. */
  alpha: number;
  lineWidth: number;
  dash: number[];
}

export interface CheckedImage {
  width: number;
  height: number;
  filter: string;
  colorSpace: string;
  bits: number;
  /** The samples after decoding (and after the predictor was undone), for an image with the FlateDecode filter. */
  samples?: Uint8Array;
  /** The samples were stored with the PNG predictors. */
  predictor?: boolean;
  /** /Decode of the image, when it has one. */
  decode?: number[];
  /** The bytes as they are in the file. */
  data: Uint8Array;
  mask?: CheckedImage;
  /** Where the image is on the page: the matrix it was drawn with. */
  matrix?: number[];
}

export interface CheckedPage {
  /** In points. */
  width: number;
  height: number;
  texts: CheckedText[];
  /** All strings of the page in the order of the content stream. */
  text: string;
  images: CheckedImage[];
  /** The content stream, decoded. */
  content: string;
  /** What the page fills and strokes, in the order of the content stream. */
  shapes: CheckedShape[];
  /** The paths the page clips with. */
  clips: (string | number)[][][];
}

export interface CheckedFont {
  baseFont: string;
  tables: string[];
  glyphs: number;
  /** Codes with an entry in the ToUnicode map. */
  mapped: number;
  /** The font file as it is in the PDF: a TrueType file, or a CFF table. */
  file: Uint8Array;
  kind: 'truetype' | 'cff';
  /** The widths /W states, by glyph. */
  widths: Map<number, number>;
  /** Flags of the font descriptor. */
  flags: number;
}

export interface CheckedPdf {
  version: string;
  objects: Map<number, Obj>;
  trailer: Dict;
  info: { [key: string]: string };
  pages: CheckedPage[];
  fonts: CheckedFont[];
}

function decodeTextString(value: Str): string {
  const bytes = value.bytes;
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    let out = '';
    for (let i = 2; i + 1 < bytes.length; i += 2) out += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
    return out;
  }
  let out = '';
  for (let i = 0; i < bytes.length; i++) out += String.fromCharCode(bytes[i]);
  return out;
}

function u16(bytes: Uint8Array, at: number): number {
  return (bytes[at] << 8) | bytes[at + 1];
}
function u32(bytes: Uint8Array, at: number): number {
  return ((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0;
}

/** Checks a TrueType file as the OpenType specification describes it: directory, checksums, the tables a PDF needs. */
export function checkTrueType(file: Uint8Array): { tables: string[]; glyphs: number; unitsPerEm: number; advance: number; table(tag: string): Uint8Array } {
  if (file.length < 12) throw new Error('The font is shorter than its header');
  if (u32(file, 0) !== 0x00010000) throw new Error('The font does not start as a TrueType file');
  const count = u16(file, 4);
  if (count === 0 || 12 + count * 16 > file.length) throw new Error('The table directory of the font does not fit');
  // searchRange, entrySelector, rangeShift.
  let power = 1;
  let selector = 0;
  while (power * 2 <= count) {
    power *= 2;
    selector++;
  }
  if (u16(file, 6) !== power * 16 || u16(file, 8) !== selector || u16(file, 10) !== count * 16 - power * 16) throw new Error('The search fields of the font directory are wrong');
  const tables = new Map<string, Uint8Array>();
  const sumOf = (from: number, length: number, skip = -1): number => {
    let sum = 0;
    for (let i = 0; i < length; i += 4) {
      if (from + i === skip) continue;
      let word = 0;
      for (let b = 0; b < 4; b++) word = word * 256 + (i + b < length ? file[from + i + b] : 0);
      sum = (sum + word) >>> 0;
    }
    return sum;
  };
  let previous = '';
  let headAt = -1;
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    const tag = ascii(file, at, at + 4);
    if (tag <= previous) throw new Error(`The tables of the font are not in order: "${tag}" after "${previous}"`);
    previous = tag;
    const offset = u32(file, at + 8);
    const length = u32(file, at + 12);
    if (offset % 4 !== 0) throw new Error(`Table ${tag} does not start at a multiple of four`);
    if (offset + length > file.length) throw new Error(`Table ${tag} lies beyond the end of the font`);
    if (tag === 'head') headAt = offset;
    const sum = sumOf(offset, length, tag === 'head' ? offset + 8 : -1);
    if (sum !== u32(file, at + 4)) throw new Error(`The checksum of table ${tag} is wrong`);
    tables.set(tag, file.subarray(offset, offset + length));
  }
  for (const tag of ['glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp']) if (!tables.has(tag)) throw new Error(`The font has no ${tag} table`);
  if (file.length % 4 !== 0) throw new Error('The font does not end at a multiple of four');
  const head = tables.get('head') as Uint8Array;
  if (head.length !== 54) throw new Error('The head table is not 54 bytes long');
  if (u32(head, 12) !== 0x5f0f3cf5) throw new Error('The magic number of the head table is wrong');
  // The whole file, with the adjustment as it is, sums to 0xB1B0AFBA.
  if (sumOf(0, file.length) !== 0xb1b0afba) throw new Error('checkSumAdjustment of the head table is wrong');
  if (headAt < 0) throw new Error('No head table');
  const unitsPerEm = u16(head, 18);
  if (unitsPerEm < 16 || unitsPerEm > 16384) throw new Error('unitsPerEm is out of range');
  const maxp = tables.get('maxp') as Uint8Array;
  if (u32(maxp, 0) !== 0x00010000 || maxp.length !== 32) throw new Error('The maxp table is no version 1.0 table');
  const glyphs = u16(maxp, 4);
  const hhea = tables.get('hhea') as Uint8Array;
  if (hhea.length !== 36) throw new Error('The hhea table is not 36 bytes long');
  const metrics = u16(hhea, 34);
  if (metrics < 1 || metrics > glyphs) throw new Error('numberOfHMetrics does not go with the number of glyphs');
  const hmtx = tables.get('hmtx') as Uint8Array;
  if (hmtx.length !== metrics * 4 + (glyphs - metrics) * 2) throw new Error('The hmtx table has the wrong length');
  const long = u16(head, 50) === 1;
  const loca = tables.get('loca') as Uint8Array;
  if (loca.length < (glyphs + 1) * (long ? 4 : 2)) throw new Error('The loca table is too short for the glyphs');
  const glyf = tables.get('glyf') as Uint8Array;
  let last = 0;
  for (let i = 0; i <= glyphs; i++) {
    const at = long ? u32(loca, i * 4) : u16(loca, i * 2) * 2;
    if (at < last || at > glyf.length) throw new Error(`Glyph ${i} lies outside the glyf table`);
    last = at;
  }
  return {
    tables: Array.from(tables.keys()),
    glyphs,
    unitsPerEm,
    advance: u16(hmtx, (Math.min(1, metrics - 1)) * 4),
    table: (tag: string) => {
      const table = tables.get(tag);
      if (!table) throw new Error(`No ${tag} table`);
      return table;
    },
  };
}

/**
 * Undoes the PNG predictors (PDF /Predictor 15, as the PNG specification
 * describes them): every row starts with the number of its filter.
 */
export function unfilterPng(data: Uint8Array, width: number, height: number, colors: number): Uint8Array {
  const stride = width * colors;
  if (data.length !== (stride + 1) * height) throw new Error(`Predicted samples of ${width} x ${height} with ${colors} components have ${data.length} bytes, not ${(stride + 1) * height}`);
  const out = new Uint8Array(stride * height);
  for (let y = 0; y < height; y++) {
    const filter = data[y * (stride + 1)];
    if (filter > 4) throw new Error(`Row ${y} of an image has the unknown filter ${filter}`);
    for (let i = 0; i < stride; i++) {
      const raw = data[y * (stride + 1) + 1 + i];
      const a = i >= colors ? out[y * stride + i - colors] : 0;
      const b = y > 0 ? out[(y - 1) * stride + i] : 0;
      const c = i >= colors && y > 0 ? out[(y - 1) * stride + i - colors] : 0;
      let predicted = 0;
      if (filter === 1) predicted = a;
      else if (filter === 2) predicted = b;
      else if (filter === 3) predicted = (a + b) >> 1;
      else if (filter === 4) {
        const guess = a + b - c;
        const pa = Math.abs(guess - a);
        const pb = Math.abs(guess - b);
        const pc = Math.abs(guess - c);
        predicted = pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      }
      out[y * stride + i] = (raw + predicted) & 255;
    }
  }
  return out;
}

/** The codes a ToUnicode CMap maps, each to its text. */
export function parseToUnicode(cmap: string): Map<number, string> {
  const map = new Map<number, string>();
  const utf16 = (hex: string): string => {
    let out = '';
    for (let i = 0; i + 3 < hex.length; i += 4) out += String.fromCharCode(parseInt(hex.substr(i, 4), 16));
    return out;
  };
  if (!/begincodespacerange\s*<0000>\s*<[fF]{4}>\s*endcodespacerange/.test(cmap)) throw new Error('The ToUnicode map has no two byte code space');
  const chars = /(\d+)\s+beginbfchar([\s\S]*?)endbfchar/g;
  for (let block = chars.exec(cmap); block; block = chars.exec(cmap)) {
    const pairs = block[2].match(/<[0-9a-fA-F]+>\s*<[0-9a-fA-F]+>/g) || [];
    if (pairs.length !== Number(block[1])) throw new Error('A bfchar block states another number of entries than it has');
    if (pairs.length > 100) throw new Error('A bfchar block has more than 100 entries');
    for (const pair of pairs) {
      const parts = pair.match(/[0-9a-fA-F]+/g) as string[];
      if (parts[0].length !== 4) throw new Error('A code of the ToUnicode map is not two bytes long');
      if (parts[1].length % 4 !== 0) throw new Error('A target of the ToUnicode map is no UTF-16');
      map.set(parseInt(parts[0], 16), utf16(parts[1]));
    }
  }
  const ranges = /(\d+)\s+beginbfrange([\s\S]*?)endbfrange/g;
  for (let block = ranges.exec(cmap); block; block = ranges.exec(cmap)) {
    const triples = block[2].match(/<[0-9a-fA-F]+>\s*<[0-9a-fA-F]+>\s*<[0-9a-fA-F]+>/g) || [];
    if (triples.length !== Number(block[1])) throw new Error('A bfrange block states another number of entries than it has');
    for (const triple of triples) {
      const parts = triple.match(/[0-9a-fA-F]+/g) as string[];
      const first = parseInt(parts[0], 16);
      const lastCode = parseInt(parts[1], 16);
      const target = parseInt(parts[2], 16);
      for (let code = first; code <= lastCode && code - first < 65536; code++) map.set(code, String.fromCharCode(target + code - first));
    }
  }
  return map;
}

/**
 * Reads a PDF the way the specification lays it out and throws at the first
 * thing that is not as it must be: the header, `startxref`, every entry of
 * the cross reference table (20 bytes, pointing at "n 0 obj"), nothing
 * between the objects, /Length of every stream, the page tree and its
 * counts, streams that inflate, images whose samples have the stated size,
 * fonts whose file is a sound TrueType file, text whose codes the ToUnicode
 * map knows.
 */
export function checkPdf(file: Uint8Array): CheckedPdf {
  const head = ascii(file, 0, 16);
  const version = /^%PDF-(1\.[0-7])[\r\n]/.exec(head);
  if (!version) throw new Error('The file does not start with a PDF header');
  const tail = ascii(file, Math.max(0, file.length - 64), file.length);
  const end = /startxref\r?\n(\d+)\r?\n%%EOF\r?\n?$/.exec(tail);
  if (!end) throw new Error('The file does not end with startxref and %%EOF');
  const xrefAt = Number(end[1]);
  if (ascii(file, xrefAt, xrefAt + 5) !== 'xref\n') throw new Error('startxref does not point at the cross reference table');
  const section = /^xref\n(\d+) (\d+)\n/.exec(ascii(file, xrefAt, xrefAt + 40));
  if (!section || section[1] !== '0') throw new Error('The cross reference table does not start at object 0');
  const count = Number(section[2]);
  let at = xrefAt + section[0].length;
  const offsets: number[] = [];
  for (let i = 0; i < count; i++) {
    const entry = ascii(file, at, at + 20);
    const parsed = /^(\d{10}) (\d{5}) ([nf])[ \r][\r\n]$/.exec(entry);
    if (!parsed) throw new Error(`Entry ${i} of the cross reference table is not 20 bytes of the right form: ${JSON.stringify(entry)}`);
    if (i === 0) {
      if (parsed[3] !== 'f' || parsed[2] !== '65535') throw new Error('Object 0 is not the head of the free list');
      offsets.push(-1);
    } else {
      if (parsed[3] !== 'n' || parsed[2] !== '00000') throw new Error(`Object ${i} is not in use`);
      offsets.push(Number(parsed[1]));
    }
    at += 20;
  }
  const trailerParser = new Parser(file, at);
  trailerParser.expect('trailer');
  const trailer = trailerParser.value() as Dict;
  trailerParser.skipWhite();
  if (ascii(file, trailerParser.at, trailerParser.at + 9) !== 'startxref') throw new Error('Something stands between the trailer and startxref');
  if (trailer.Size !== count) throw new Error(`/Size of the trailer (${String(trailer.Size)}) is not the number of entries (${count})`);
  if (!(trailer.Root instanceof Ref)) throw new Error('The trailer has no /Root');

  // Every object, at its offset; in the order of the file nothing lies between them.
  const objects = new Map<number, Obj>();
  const ends = new Map<number, number>();
  for (let id = 1; id < count; id++) {
    const offset = offsets[id];
    if (!(offset > 0) || offset >= xrefAt) throw new Error(`Object ${id} has an offset outside the body of the file`);
    const start = new RegExp(`^${id} 0 obj[\\s<\\[(/]`);
    if (!start.test(ascii(file, offset, offset + 24))) throw new Error(`The offset of object ${id} (${offset}) does not point at "${id} 0 obj"`);
    const parser = new Parser(file, offset + `${id} 0 obj`.length);
    const value = parser.value();
    parser.skipWhite();
    const object: Obj = { id, offset, value };
    if (ascii(file, parser.at, parser.at + 6) === 'stream') {
      parser.at += 6;
      if (file[parser.at] === 13) parser.at++;
      if (file[parser.at] !== 10) throw new Error(`The keyword stream of object ${id} is not followed by a line feed`);
      parser.at++;
      if (value === null || typeof value !== 'object' || Array.isArray(value) || value instanceof Name || value instanceof Ref || value instanceof Str) throw new Error(`Stream ${id} has no dictionary`);
      const length = (value as Dict).Length;
      if (typeof length !== 'number' || length < 0 || Math.floor(length) !== length) throw new Error(`Stream ${id} has no direct /Length`);
      if (parser.at + length > file.length) throw new Error(`Stream ${id} is longer than the file`);
      object.stream = file.subarray(parser.at, parser.at + length);
      parser.at += length;
      // An end of line stands between the data and the keyword; it is no part of the data.
      const dataEnd = parser.at;
      if (file[parser.at] === 13) parser.at++;
      if (file[parser.at] === 10) parser.at++;
      if (ascii(file, parser.at, parser.at + 9) !== 'endstream') throw new Error(`/Length of stream ${id} (${length}) does not end at endstream`);
      if (parser.at === dataEnd) throw new Error(`No end of line stands before the endstream of stream ${id}`);
      parser.at += 9;
      parser.skipWhite();
    }
    if (ascii(file, parser.at, parser.at + 6) !== 'endobj') throw new Error(`Object ${id} does not end with endobj`);
    ends.set(id, parser.at + 6);
    objects.set(id, object);
  }
  const inOrder = Array.from(objects.values()).sort((a, b) => a.offset - b.offset);
  for (let i = 0; i < inOrder.length; i++) {
    const from = ends.get(inOrder[i].id) as number;
    const to = i + 1 < inOrder.length ? inOrder[i + 1].offset : xrefAt;
    if (to < from) throw new Error(`Object ${inOrder[i].id} overlaps what follows it`);
    for (let k = from; k < to; k++) if (!isWhite(file[k])) throw new Error(`Bytes that belong to no object after object ${inOrder[i].id} (offset ${k})`);
  }
  // Before the first object: the header and its comment line only.
  if (inOrder.length > 0) {
    const first = new Parser(file, 0);
    first.skipWhite();
    if (first.at !== inOrder[0].offset) throw new Error('Something stands between the header and the first object');
  }

  const resolve = (value: Value | undefined): Value | undefined => {
    for (let i = 0; i < 16 && value instanceof Ref; i++) {
      const object = objects.get(value.id);
      if (!object) throw new Error(`A reference to object ${value.id}, which the file does not have`);
      value = object.value;
    }
    return value;
  };
  const dictOf = (value: Value | undefined, what: string): Dict => {
    const resolved = resolve(value);
    if (resolved === undefined || resolved === null || typeof resolved !== 'object' || Array.isArray(resolved) || resolved instanceof Name || resolved instanceof Str || resolved instanceof Ref) throw new Error(`${what} is no dictionary`);
    return resolved as Dict;
  };
  const nameOf = (value: Value | undefined): string => {
    const resolved = resolve(value);
    return resolved instanceof Name ? resolved.name : '';
  };
  const streamOf = (value: Value | undefined, what: string): { dict: Dict; raw: Uint8Array; data: Uint8Array } => {
    if (!(value instanceof Ref)) throw new Error(`${what} is not a reference to a stream`);
    const object = objects.get(value.id);
    if (!object || !object.stream) throw new Error(`${what} is no stream`);
    const dict = object.value as Dict;
    const filter = nameOf(dict.Filter);
    if (filter === 'FlateDecode') {
      let data: Uint8Array;
      try {
        data = inflateSync(object.stream);
      } catch (error) {
        throw new Error(`${what} (object ${value.id}) does not inflate: ${(error as Error).message}`);
      }
      return { dict, raw: object.stream, data };
    }
    if (filter !== '' && filter !== 'DCTDecode') throw new Error(`${what} has a filter this check does not know: ${filter}`);
    return { dict, raw: object.stream, data: object.stream };
  };

  // Every reference anywhere must lead to an object.
  const visit = (value: Value, depth: number): void => {
    if (depth > 64) throw new Error('A value is nested too deeply');
    if (value instanceof Ref) {
      if (!objects.has(value.id)) throw new Error(`A reference to object ${value.id}, which the file does not have`);
    } else if (Array.isArray(value)) value.forEach((entry) => visit(entry, depth + 1));
    else if (value !== null && typeof value === 'object' && !(value instanceof Name) && !(value instanceof Str)) for (const key of Object.keys(value)) visit((value as Dict)[key], depth + 1);
  };
  objects.forEach((object) => visit(object.value, 0));
  visit(trailer, 0);

  const info: { [key: string]: string } = {};
  if (trailer.Info !== undefined) {
    const dict = dictOf(trailer.Info, '/Info');
    for (const key of Object.keys(dict)) {
      const entry = dict[key];
      if (entry instanceof Str) info[key] = decodeTextString(entry);
    }
  }
  if (trailer.ID !== undefined) {
    const id = trailer.ID;
    if (!Array.isArray(id) || id.length !== 2 || !(id[0] instanceof Str) || !(id[1] instanceof Str) || (id[0] as Str).bytes.length < 16) throw new Error('/ID is not a pair of strings of 16 bytes');
  }

  const catalog = dictOf(trailer.Root, '/Root');
  if (nameOf(catalog.Type) !== 'Catalog') throw new Error('/Root is no /Catalog');
  interface FontEntry {
    font: CheckedFont;
    map: Map<number, string>;
    /** The width of a code, in thousandths of the font size. */
    width(code: number): number;
  }
  const fonts = new Map<number, FontEntry>();
  const fontOf = (ref: Value | undefined): FontEntry => {
    if (!(ref instanceof Ref)) throw new Error('A font of a page is not a reference');
    const known = fonts.get(ref.id);
    if (known) return known;
    const font = dictOf(ref, 'A font');
    if (nameOf(font.Type) !== 'Font' || nameOf(font.Subtype) !== 'Type0') throw new Error('The font is no Type0 font');
    if (nameOf(font.Encoding) !== 'Identity-H') throw new Error('The font is not encoded with Identity-H');
    const descendants = resolve(font.DescendantFonts);
    if (!Array.isArray(descendants) || descendants.length !== 1) throw new Error('The font has not exactly one descendant');
    const cid = dictOf(descendants[0], 'The descendant font');
    const subtype = nameOf(cid.Subtype);
    if (subtype !== 'CIDFontType2' && subtype !== 'CIDFontType0') throw new Error('The descendant font is neither a CIDFontType2 nor a CIDFontType0');
    const system = dictOf(cid.CIDSystemInfo, '/CIDSystemInfo');
    if (!(system.Registry instanceof Str) || !(system.Ordering instanceof Str) || typeof system.Supplement !== 'number') throw new Error('/CIDSystemInfo is incomplete');
    // The codes of the font are glyph numbers and nothing else: the collection Identity-H goes with.
    if (decodeTextString(system.Registry) !== 'Adobe' || decodeTextString(system.Ordering) !== 'Identity' || system.Supplement !== 0) throw new Error('/CIDSystemInfo is not Adobe-Identity-0');
    const baseFont = nameOf(font.BaseFont);
    if (baseFont === '' || /[^!-~]|[\[\](){}<>/%#]/.test(baseFont)) throw new Error(`The name of the font (${JSON.stringify(baseFont)}) is none a PostScript name may be`);
    if (nameOf(cid.BaseFont) !== baseFont) throw new Error('The two fonts have different names');
    const descriptor = dictOf(cid.FontDescriptor, '/FontDescriptor');
    for (const key of ['FontName', 'Flags', 'FontBBox', 'ItalicAngle', 'Ascent', 'Descent', 'CapHeight', 'StemV']) if (descriptor[key] === undefined) throw new Error(`The font descriptor has no /${key}`);
    if (nameOf(descriptor.FontName) !== baseFont) throw new Error('The font descriptor has another name than the font');
    const box = resolve(descriptor.FontBBox);
    if (!Array.isArray(box) || box.length !== 4 || box.some((entry) => typeof entry !== 'number') || !((box[2] as number) > (box[0] as number)) || !((box[3] as number) > (box[1] as number))) throw new Error('/FontBBox of the font descriptor is no box');
    // A font whose codes are no Latin letters is symbolic (bit 3) and not nonsymbolic (bit 6).
    const flags = descriptor.Flags as number;
    if (typeof flags !== 'number' || (flags & 4) === 0 || (flags & 32) !== 0) throw new Error(`The flags of the font descriptor (${String(flags)}) do not say that the font is symbolic`);
    let glyphs: number;
    let file: Uint8Array;
    let tables: string[] = [];
    let kind: 'truetype' | 'cff';
    /** The advance of a glyph as the font file states it, in thousandths of an em; undefined where the file is not asked. */
    let advanceOf: ((glyph: number) => number) | undefined;
    if (subtype === 'CIDFontType2') {
      const fontFile = streamOf(descriptor.FontFile2, '/FontFile2');
      if (fontFile.dict.Length1 !== fontFile.data.length) throw new Error('/Length1 of the font file is not its length');
      const checked = checkTrueType(fontFile.data);
      glyphs = checked.glyphs;
      file = fontFile.data;
      tables = checked.tables;
      kind = 'truetype';
      const hmtx = checked.table('hmtx');
      const metrics = u16(checked.table('hhea'), 34);
      advanceOf = (glyph) => (u16(hmtx, Math.min(glyph, metrics - 1) * 4) * 1000) / checked.unitsPerEm;
    } else {
      const fontFile = streamOf(descriptor.FontFile3, '/FontFile3');
      if (nameOf(fontFile.dict.Subtype) !== 'CIDFontType0C') throw new Error('/FontFile3 of a CIDFontType0 is no /CIDFontType0C');
      // A CFF table: version 1, a header of four bytes, an offset size from 1 to 4.
      const data = fontFile.data;
      if (data.length < 4 || data[0] !== 1 || data[2] < 4 || data[3] < 1 || data[3] > 4) throw new Error('/FontFile3 does not start as a CFF table');
      // The glyphs are counted by the font program itself.
      glyphs = 65536;
      file = data;
      kind = 'cff';
    }
    const dw = typeof cid.DW === 'number' ? cid.DW : 1000;
    const widths = new Map<number, number>();
    if (cid.W !== undefined) {
      const w = resolve(cid.W);
      if (!Array.isArray(w)) throw new Error('/W of the font is no array');
      for (let i = 0; i < w.length; ) {
        const first = w[i];
        const next = w[i + 1];
        if (typeof first !== 'number' || Math.floor(first) !== first || first < 0) throw new Error('/W of the font does not start a run with a code');
        if (Array.isArray(next)) {
          next.forEach((entry, index) => {
            if (typeof entry !== 'number' || entry < 0) throw new Error('/W of the font holds a width that is none');
            if (widths.has(first + index)) throw new Error(`/W of the font states the width of code ${first + index} twice`);
            widths.set(first + index, entry);
          });
          i += 2;
        } else if (typeof next === 'number' && typeof w[i + 2] === 'number') {
          for (let code = first; code <= next && code - first < 65536; code++) widths.set(code, w[i + 2] as number);
          i += 3;
        } else throw new Error('/W of the font is not made of runs');
      }
    }
    if (advanceOf) {
      const advance = advanceOf;
      widths.forEach((width, code) => {
        if (code >= glyphs) throw new Error(`/W states the width of glyph ${code}, which the font does not have`);
        if (Math.abs(advance(code) - width) > 0.51) throw new Error(`/W states ${width} for glyph ${code}; the font file says ${advance(code)}`);
      });
      if (widths.size === 0 && Math.abs(advance(1) - dw) > 0.5) throw new Error('/DW is not the advance of the glyph in the font file');
    }
    let mapped = 0;
    let map = new Map<number, string>();
    if (font.ToUnicode !== undefined) {
      const cmap = streamOf(font.ToUnicode, '/ToUnicode');
      map = parseToUnicode(Buffer.from(cmap.data).toString('latin1'));
      mapped = map.size;
    }
    // Every code the map knows has a glyph: through /CIDToGIDMap, a glyph the font has.
    const toGid = cid.CIDToGIDMap;
    if (toGid instanceof Ref) {
      const table = streamOf(toGid, '/CIDToGIDMap').data;
      map.forEach((_text, code) => {
        if (code * 2 + 1 >= table.length) throw new Error(`Code ${code} lies beyond /CIDToGIDMap`);
        if (u16(table, code * 2) >= glyphs) throw new Error(`Code ${code} is mapped to a glyph the font does not have`);
        // Glyph 0 is "no such character": readers may leave text that is made of it out.
        if (u16(table, code * 2) === 0) throw new Error(`Code ${code} is mapped to the glyph for missing characters`);
      });
    } else if (subtype === 'CIDFontType2') {
      if (nameOf(toGid) !== 'Identity') throw new Error('The font has no /CIDToGIDMap');
      map.forEach((_text, code) => {
        if (code >= glyphs) throw new Error(`Code ${code} is a glyph the font does not have`);
        if (code === 0) throw new Error('Code 0, the glyph for missing characters, is in the ToUnicode map');
      });
    } else if (toGid !== undefined) throw new Error('A CIDFontType0 has a /CIDToGIDMap');
    const entry: FontEntry = { font: { baseFont, tables, glyphs, mapped, file, kind, widths, flags }, map, width: (code) => (widths.has(code) ? (widths.get(code) as number) : dw) };
    fonts.set(ref.id, entry);
    return entry;
  };

  const imageOf = (ref: Value | undefined): CheckedImage => {
    const stream = streamOf(ref, 'An image');
    const dict = stream.dict;
    if (nameOf(dict.Subtype) !== 'Image') throw new Error('An XObject of a page is no image');
    const width = dict.Width as number;
    const height = dict.Height as number;
    if (!(width >= 1) || !(height >= 1) || Math.floor(width) !== width || Math.floor(height) !== height) throw new Error('An image has no size');
    const colorSpace = nameOf(dict.ColorSpace);
    const colors = colorSpace === 'DeviceGray' ? 1 : colorSpace === 'DeviceRGB' ? 3 : colorSpace === 'DeviceCMYK' ? 4 : 0;
    if (colors === 0) throw new Error(`An image has the colour space "${colorSpace}"`);
    if (dict.BitsPerComponent !== 8) throw new Error('An image has not 8 bits per component');
    const filter = nameOf(dict.Filter);
    const image: CheckedImage = { width, height, filter, colorSpace, bits: 8, data: stream.raw };
    if (filter === 'DCTDecode') {
      if (stream.raw[0] !== 0xff || stream.raw[1] !== 0xd8) throw new Error('A DCTDecode image does not start as a JPEG file');
      if (stream.raw[stream.raw.length - 2] !== 0xff || stream.raw[stream.raw.length - 1] !== 0xd9) throw new Error('A DCTDecode image does not end as a JPEG file');
    } else if (filter === 'FlateDecode') {
      let samples = stream.data;
      if (dict.DecodeParms !== undefined) {
        const parms = dictOf(dict.DecodeParms, '/DecodeParms');
        if (parms.Predictor !== 15) throw new Error(`An image has the predictor ${String(parms.Predictor)}, which this check does not know`);
        if (parms.Colors !== colors || parms.Columns !== width || (parms.BitsPerComponent !== undefined && parms.BitsPerComponent !== 8)) throw new Error('The parameters of the predictor of an image do not go with the image');
        samples = unfilterPng(samples, width, height, colors);
        image.predictor = true;
      }
      if (samples.length !== width * height * colors) throw new Error(`An image of ${width} x ${height} with ${colors} components has ${samples.length} bytes of samples`);
      image.samples = samples;
    } else throw new Error(`An image has the filter "${filter}"`);
    if (dict.Decode !== undefined) {
      const decode = resolve(dict.Decode);
      if (!Array.isArray(decode) || decode.length !== colors * 2 || decode.some((entry) => typeof entry !== 'number')) throw new Error('/Decode of an image does not go with its colour space');
      image.decode = decode as number[];
    }
    if (dict.SMask !== undefined) {
      const mask = imageOf(dict.SMask);
      if (mask.colorSpace !== 'DeviceGray' || mask.width !== width || mask.height !== height) throw new Error('The soft mask of an image does not go with it');
      image.mask = mask;
    }
    return image;
  };

  const pages: CheckedPage[] = [];
  const seen = new Set<number>();
  const walk = (ref: Value | undefined, parent: number, depth: number): number => {
    if (!(ref instanceof Ref)) throw new Error('A node of the page tree is not a reference');
    if (depth > 32 || seen.has(ref.id)) throw new Error('The page tree turns back on itself');
    seen.add(ref.id);
    const node = dictOf(ref, 'A node of the page tree');
    if (parent !== 0 && !(node.Parent instanceof Ref && node.Parent.id === parent)) throw new Error(`Node ${ref.id} of the page tree does not name its parent`);
    const type = nameOf(node.Type);
    if (type === 'Pages') {
      const kids = resolve(node.Kids);
      if (!Array.isArray(kids)) throw new Error('A /Pages node has no /Kids');
      let leaves = 0;
      for (const kid of kids) leaves += walk(kid, ref.id, depth + 1);
      if (node.Count !== leaves) throw new Error(`/Count of node ${ref.id} (${String(node.Count)}) is not its number of pages (${leaves})`);
      return leaves;
    }
    if (type !== 'Page') throw new Error(`Node ${ref.id} of the page tree is neither /Pages nor /Page`);
    const box = resolve(node.MediaBox);
    if (!Array.isArray(box) || box.length !== 4 || box.some((entry) => typeof entry !== 'number')) throw new Error('A page has no /MediaBox');
    const numbers = box as number[];
    const page: CheckedPage = { width: numbers[2] - numbers[0], height: numbers[3] - numbers[1], texts: [], text: '', images: [], content: '', shapes: [], clips: [] };
    if (!(page.width > 0) || !(page.height > 0)) throw new Error('A page has no size');
    const resources = dictOf(node.Resources, '/Resources');
    const xobjects = resources.XObject === undefined ? {} : dictOf(resources.XObject, '/XObject');
    const fontNames = resources.Font === undefined ? {} : dictOf(resources.Font, '/Font');
    const states = resources.ExtGState === undefined ? {} : dictOf(resources.ExtGState, '/ExtGState');
    const content = node.Contents === undefined ? '' : Buffer.from(streamOf(node.Contents, '/Contents').data).toString('latin1');
    page.content = content;
    // The operators of the content stream, with their operands.
    const parser = new Parser(Buffer.from(content, 'latin1'), 0);
    const operands: Value[] = [];
    let inText = false;
    // What q saves and Q gives back.
    interface Graphics {
      ctm: number[];
      font: FontEntry | undefined;
      size: number;
      scaling: number;
      mode: number;
      spacing: number;
      fill: number[];
      stroke: number[];
      ca: number;
      CA: number;
      lineWidth: number;
      dash: number[];
    }
    let gs: Graphics = { ctm: [1, 0, 0, 1, 0, 0], font: undefined, size: 0, scaling: 100, mode: 0, spacing: 0, fill: [0], stroke: [0], ca: 1, CA: 1, lineWidth: 1, dash: [] };
    let matrix = [1, 0, 0, 1, 0, 0];
    const stack: Graphics[] = [];
    let path: (string | number)[][] = [];
    /** A clip was asked for: the operator that ends the path makes it. */
    let clipping = false;
    for (;;) {
      parser.skipWhite();
      if (parser.at >= parser.bytes.length) break;
      const code = parser.bytes[parser.at];
      const isOperand = code === 0x2f || code === 0x28 || code === 0x3c || code === 0x5b || code === 0x2b || code === 0x2d || code === 0x2e || (code >= 0x30 && code <= 0x39);
      if (isOperand) {
        operands.push(parser.value());
        continue;
      }
      const operator = parser.word();
      if (operator === '') parser.fail('Something in a content stream that is neither an operand nor an operator');
      const numbersOf = (wanted: number): number[] => {
        if (operands.length !== wanted || operands.some((entry) => typeof entry !== 'number')) throw new Error(`The operator ${operator} needs ${wanted} numbers`);
        return operands as number[];
      };
      const unit = (count: number): number[] => {
        const values = numbersOf(count).slice();
        if (values.some((value) => value < 0 || value > 1)) throw new Error(`The operator ${operator} is given a colour outside 0 to 1`);
        return values;
      };
      const paint = (stroke: boolean): void => {
        if (path.length === 0) throw new Error(`The operator ${operator} paints a path that is empty`);
        if (clipping) page.clips.push(path);
        if (operator !== 'n') page.shapes.push({ operator, path, colour: (stroke ? gs.stroke : gs.fill).slice(), alpha: stroke ? gs.CA : gs.ca, lineWidth: gs.lineWidth, dash: gs.dash.slice() });
        path = [];
        clipping = false;
      };
      const construct = (count: number): void => {
        if (inText) throw new Error(`The operator ${operator} builds a path inside a text object`);
        if (operator !== 'm' && operator !== 're' && path.length === 0) throw new Error(`The operator ${operator} goes on with a path that was not started`);
        path.push([operator as string | number].concat(numbersOf(count)));
      };
      const pathless = (): void => {
        if (path.length !== 0) throw new Error(`The operator ${operator} stands inside a path`);
      };
      if (operator === 'q') {
        pathless();
        if (inText) throw new Error('A q inside a text object');
        stack.push({ ...gs, ctm: gs.ctm.slice(), fill: gs.fill.slice(), stroke: gs.stroke.slice(), dash: gs.dash.slice() });
        if (stack.length > 28) throw new Error('The graphics state is saved more than 28 times over');
      } else if (operator === 'Q') {
        pathless();
        if (inText) throw new Error('A Q inside a text object');
        const popped = stack.pop();
        if (!popped) throw new Error('A Q without its q');
        gs = popped;
      } else if (operator === 'cm') {
        pathless();
        const m = numbersOf(6);
        const ctm = gs.ctm;
        gs.ctm = [m[0] * ctm[0] + m[1] * ctm[2], m[0] * ctm[1] + m[1] * ctm[3], m[2] * ctm[0] + m[3] * ctm[2], m[2] * ctm[1] + m[3] * ctm[3], m[4] * ctm[0] + m[5] * ctm[2] + ctm[4], m[4] * ctm[1] + m[5] * ctm[3] + ctm[5]];
      } else if (operator === 'Do') {
        pathless();
        const key = operands[0];
        if (!(key instanceof Name) || xobjects[key.name] === undefined) throw new Error('A page draws an XObject its resources do not have');
        const image = { ...imageOf(xobjects[key.name]) };
        image.matrix = gs.ctm.slice();
        page.images.push(image);
      } else if (operator === 'BT') {
        pathless();
        if (inText) throw new Error('A BT inside a text object');
        inText = true;
        matrix = [1, 0, 0, 1, 0, 0];
      } else if (operator === 'ET') {
        if (!inText) throw new Error('An ET without its BT');
        inText = false;
      } else if (operator === 'Tf') {
        const key = operands[0];
        if (!(key instanceof Name) || fontNames[key.name] === undefined || typeof operands[1] !== 'number') throw new Error('A page uses a font its resources do not have');
        gs.font = fontOf(fontNames[key.name]);
        gs.size = operands[1];
      } else if (operator === 'Tz') gs.scaling = numbersOf(1)[0];
      else if (operator === 'Tr') gs.mode = numbersOf(1)[0];
      else if (operator === 'Tc') gs.spacing = numbersOf(1)[0];
      else if (operator === 'Tm') {
        if (!inText) throw new Error('A text matrix outside a text object');
        matrix = numbersOf(6).slice();
      } else if (operator === 'Td') {
        const d = numbersOf(2);
        matrix = [matrix[0], matrix[1], matrix[2], matrix[3], d[0] * matrix[0] + d[1] * matrix[2] + matrix[4], d[0] * matrix[1] + d[1] * matrix[3] + matrix[5]];
      } else if (operator === 'Tj' || operator === 'TJ') {
        if (!inText) throw new Error('Text outside a text object');
        // The strings that are shown, and the numbers between them (TJ): thousandths of the font size the next glyph stands further left by.
        let pieces: Value[];
        if (operator === 'Tj') {
          if (operands.length !== 1 || !(operands[0] instanceof Str)) throw new Error('Tj without a string');
          pieces = [operands[0]];
        } else {
          const array = operands[0];
          if (operands.length !== 1 || !Array.isArray(array) || array.length === 0 || !(array[0] instanceof Str) || !(array[array.length - 1] instanceof Str)) throw new Error('TJ without an array that starts and ends with a string');
          pieces = array;
        }
        const font = gs.font;
        if (!font) throw new Error('Text before a font was chosen');
        let text = '';
        // In units of text space.
        let width = 0;
        const glyphs: number[] = [];
        const starts: number[] = [];
        for (const piece of pieces) {
          if (typeof piece === 'number') {
            width -= (piece * gs.size) / 1000;
            continue;
          }
          if (!(piece instanceof Str)) throw new Error('TJ with something that is neither a string nor a number');
          if (piece.bytes.length % 2 !== 0) throw new Error('A string of a two byte font with an odd number of bytes');
          for (let i = 0; i < piece.bytes.length; i += 2) {
            const cid = u16(piece.bytes, i);
            const mappedText = font.map.get(cid);
            if (mappedText === undefined) throw new Error(`Code ${cid} of a string is not in the ToUnicode map`);
            text += mappedText;
            starts.push(width);
            width += (font.width(cid) * gs.size) / 1000 + gs.spacing;
            glyphs.push(cid);
          }
        }
        const codes = glyphs.length;
        const stretch = Math.sqrt(matrix[0] * matrix[0] + matrix[1] * matrix[1]);
        const advance = (width * gs.scaling * stretch) / 100;
        page.texts.push({ text, x: matrix[4], y: matrix[5], matrix: matrix.slice(), size: gs.size, scaling: gs.scaling, mode: gs.mode, codes, advance, spacing: gs.spacing, font: font.font.baseFont, fill: gs.fill.slice(), glyphs, starts });
        page.text += text;
      } else if (operator === 'rg') gs.fill = unit(3);
      else if (operator === 'g') gs.fill = unit(1);
      else if (operator === 'RG') gs.stroke = unit(3);
      else if (operator === 'G') gs.stroke = unit(1);
      else if (operator === 'gs') {
        pathless();
        const key = operands[0];
        if (operands.length !== 1 || !(key instanceof Name) || states[key.name] === undefined) throw new Error('A page uses a graphics state its resources do not have');
        const state = dictOf(states[key.name], 'A graphics state');
        if (nameOf(state.Type) !== 'ExtGState') throw new Error('A graphics state is no /ExtGState');
        for (const entry of Object.keys(state)) {
          if (entry === 'Type') continue;
          const value = state[entry];
          if ((entry !== 'ca' && entry !== 'CA') || typeof value !== 'number' || value < 0 || value > 1) throw new Error(`A graphics state sets /${entry}, which this check does not know, or to a value that is no opacity`);
          if (entry === 'ca') gs.ca = value;
          else gs.CA = value;
        }
      } else if (operator === 'w') {
        gs.lineWidth = numbersOf(1)[0];
        if (gs.lineWidth < 0) throw new Error('A line has a width below nothing');
      } else if (operator === 'J' || operator === 'j') {
        const value = numbersOf(1)[0];
        if (value !== 0 && value !== 1 && value !== 2) throw new Error(`The operator ${operator} is given ${value}`);
      } else if (operator === 'M') {
        if (numbersOf(1)[0] < 1) throw new Error('A miter limit below 1');
      } else if (operator === 'd') {
        const lengths = operands[0];
        if (operands.length !== 2 || !Array.isArray(lengths) || typeof operands[1] !== 'number' || lengths.some((entry) => typeof entry !== 'number' || entry < 0)) throw new Error('The operator d needs an array of lengths and a phase');
        if (lengths.length > 0 && !lengths.some((entry) => (entry as number) > 0)) throw new Error('The dashes of a line are all of no length');
        gs.dash = lengths as number[];
      } else if (operator === 'm') construct(2);
      else if (operator === 'l') construct(2);
      else if (operator === 'c') construct(6);
      else if (operator === 're') construct(4);
      else if (operator === 'h') construct(0);
      else if (operator === 'W' || operator === 'W*') {
        if (path.length === 0) throw new Error('A clip without a path');
        clipping = true;
      } else if (operator === 'f' || operator === 'f*' || operator === 'n') paint(false);
      else if (operator === 'S') paint(true);
      else throw new Error(`A content stream uses the operator "${operator}", which this check does not know`);
      operands.length = 0;
    }
    if (inText) throw new Error('A text object without its ET');
    if (stack.length !== 0) throw new Error('A q without its Q');
    if (path.length !== 0) throw new Error('A path of a page is never painted');
    pages.push(page);
    return 1;
  };
  walk(catalog.Pages, 0, 0);

  return { version: version[1], objects, trailer, info, pages, fonts: Array.from(fonts.values()).map((entry) => entry.font) };
}

// ---------------------------------------------------------------------------
// Poppler

let popplerFound: boolean | undefined;

/** Whether pdfinfo, pdftotext, pdfimages and pdffonts can be run here. */
export function hasPoppler(): boolean {
  if (popplerFound === undefined) {
    try {
      for (const tool of ['pdfinfo', 'pdftotext', 'pdfimages', 'pdffonts']) execFileSync(tool, ['-v'], { stdio: 'pipe' });
      popplerFound = true;
    } catch {
      popplerFound = false;
    }
  }
  return popplerFound;
}

/** With BYM_REQUIRE_POPPLER=1 a machine without Poppler fails the checks that need it instead of skipping them. */
export const requirePoppler = (): boolean => process.env.BYM_REQUIRE_POPPLER === '1';

export const POPPLER_MISSING = 'poppler not found: independent PDF checks skipped';

let saidMissing = false;

/**
 * What to do about a check that needs Poppler on a machine without it: throws
 * with BYM_REQUIRE_POPPLER=1; otherwise says once, on the output of the run,
 * that the independent checks are skipped, and returns false.
 */
export function popplerOrSay(): boolean {
  if (hasPoppler()) return true;
  if (requirePoppler()) throw new Error('BYM_REQUIRE_POPPLER=1, and Poppler (pdfinfo, pdftotext, pdfimages, pdffonts) is not installed: the independent PDF checks cannot run');
  if (!saidMissing) {
    saidMissing = true;
    console.log(`# ${POPPLER_MISSING}`);
  }
  return false;
}

/**
 * A test that reads a PDF with Poppler. Without Poppler it is skipped, with
 * the reason in its line of the report; with BYM_REQUIRE_POPPLER=1 it fails.
 */
export function popplerTest(name: string, run: () => Promise<void> | void): void {
  if (hasPoppler() || requirePoppler()) {
    nodeTest(name, async () => {
      popplerOrSay();
      await run();
    });
    return;
  }
  popplerOrSay();
  nodeTest(name, { skip: POPPLER_MISSING }, () => undefined);
}

let serial = 0;

export interface PopplerReport {
  /** Lines of pdfinfo, by their label. */
  info: { [label: string]: string };
  pages: number;
  /** Size of every page in points. */
  sizes: { width: number; height: number }[];
  /** The text of every page, as pdftotext gives it (in the order of the content stream: -raw). */
  text: string[];
  /** The text of every page as pdftotext orders it itself (its reading order, which is also where it turns lines from right to left around). */
  reading: string[];
  /** One entry per image: page, size and how it is stored. */
  images: { page: number; width: number; height: number; color: string; encoding: string }[];
  fonts: { name: string; type: string; embedded: boolean; subset: boolean; unicode: boolean }[];
  /** What the tools wrote to their error stream: syntax errors and warnings. */
  complaints: string;
}

/**
 * Runs the Poppler tools over a PDF. The file lives in a folder of its own
 * inside the working directory for as long as the tools read it.
 */
export function poppler(file: Uint8Array): PopplerReport {
  const folder = join(process.cwd(), `tmp-pdf-check-${process.pid}-${++serial}`);
  mkdirSync(folder, { recursive: true });
  const path = join(folder, 'file.pdf');
  let complaints = '';
  const run = (tool: string, args: string[]): string => {
    // The error stream is caught with the output of each tool.
    const errors = join(folder, 'stderr.txt');
    const out = execFileSync('/bin/sh', ['-c', '"$0" "$@" 2>"$ERRORS"', tool, ...args], { encoding: 'utf8', env: { ...process.env, ERRORS: errors }, maxBuffer: 256 * 1024 * 1024 });
    complaints += execFileSync('/bin/cat', [errors], { encoding: 'utf8' });
    return out;
  };
  try {
    writeFileSync(path, file);
    const info: { [label: string]: string } = {};
    const lines = run('pdfinfo', ['-l', '100000', path]).split('\n');
    const sizes: { width: number; height: number }[] = [];
    for (const line of lines) {
      const colon = line.indexOf(':');
      if (colon < 0) continue;
      const label = line.slice(0, colon).trim();
      const value = line.slice(colon + 1).trim();
      const size = /^Page\s+(\d+) size$/.exec(label);
      if (size) {
        const parts = /^([\d.]+) x ([\d.]+) pts/.exec(value);
        if (parts) sizes[Number(size[1]) - 1] = { width: Number(parts[1]), height: Number(parts[2]) };
      } else if (!/^Page\s+\d+ rot$/.test(label)) info[label] = value;
    }
    const pages = Number(info.Pages);
    const text = run('pdftotext', ['-raw', '-enc', 'UTF-8', path, '-']).split('\f');
    if (text.length > 0 && text[text.length - 1] === '') text.pop();
    const reading = run('pdftotext', ['-enc', 'UTF-8', path, '-']).split('\f');
    if (reading.length > 0 && reading[reading.length - 1] === '') reading.pop();
    const images: PopplerReport['images'] = [];
    for (const line of run('pdfimages', ['-list', path]).split('\n').slice(2)) {
      const cells = line.trim().split(/\s+/);
      if (cells.length < 9) continue;
      images.push({ page: Number(cells[0]), width: Number(cells[3]), height: Number(cells[4]), color: cells[5], encoding: cells[8] });
    }
    const fonts: PopplerReport['fonts'] = [];
    for (const line of run('pdffonts', [path]).split('\n').slice(2)) {
      const parsed = /^(\S+)\s+(.+?)\s+(\S+)\s+(yes|no)\s+(yes|no)\s+(yes|no)\s+\d+\s+\d+\s*$/.exec(line);
      if (parsed) fonts.push({ name: parsed[1], type: parsed[2].trim(), embedded: parsed[4] === 'yes', subset: parsed[5] === 'yes', unicode: parsed[6] === 'yes' });
    }
    return { info, pages, sizes, text, reading, images, fonts, complaints: complaints.trim() };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

/**
 * The first page of a PDF as Poppler paints it, at `dpi` dots per inch: the
 * colour of every pixel. For asking what a picture looks like to a reader
 * that is not ours.
 */
export function popplerPixels(file: Uint8Array, dpi = 72): { width: number; height: number; at(x: number, y: number): [number, number, number] } {
  const folder = join(process.cwd(), `tmp-pdf-check-${process.pid}-${++serial}`);
  mkdirSync(folder, { recursive: true });
  try {
    const path = join(folder, 'file.pdf');
    writeFileSync(path, file);
    const ppm = execFileSync('pdftoppm', ['-r', String(dpi), '-f', '1', '-l', '1', path], { maxBuffer: 256 * 1024 * 1024 });
    // "P6", width, height, 255, then the samples.
    const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(ppm.subarray(0, 40).toString('latin1'));
    if (!head) throw new Error('pdftoppm gave no picture');
    const width = Number(head[1]);
    const height = Number(head[2]);
    const data = ppm.subarray(head[0].length);
    return { width, height, at: (x, y) => [data[(y * width + x) * 3], data[(y * width + x) * 3 + 1], data[(y * width + x) * 3 + 2]] };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

export interface PopplerWord {
  page: number;
  text: string;
  /** The box of the word in points, from the top left corner of the page. */
  xMin: number;
  yMin: number;
  xMax: number;
  yMax: number;
}

/** The words of a PDF with the boxes Poppler gives them (pdftotext -bbox), page by page, and what the tool complained of. */
export function popplerWords(file: Uint8Array): { words: PopplerWord[]; complaints: string } {
  const folder = join(process.cwd(), `tmp-pdf-check-${process.pid}-${++serial}`);
  mkdirSync(folder, { recursive: true });
  try {
    const path = join(folder, 'file.pdf');
    const errors = join(folder, 'stderr.txt');
    writeFileSync(path, file);
    const html = execFileSync('/bin/sh', ['-c', '"$0" "$@" 2>"$ERRORS"', 'pdftotext', '-bbox', '-enc', 'UTF-8', path, '-'], { encoding: 'utf8', env: { ...process.env, ERRORS: errors }, maxBuffer: 256 * 1024 * 1024 });
    const complaints = execFileSync('/bin/cat', [errors], { encoding: 'utf8' }).trim();
    const words: PopplerWord[] = [];
    let page = 0;
    const decode = (text: string): string =>
      text
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&#39;/g, "'")
        .replace(/&amp;/g, '&');
    for (const line of html.split('\n')) {
      if (/<page /.test(line)) page++;
      const word = /<word xMin="([-\d.]+)" yMin="([-\d.]+)" xMax="([-\d.]+)" yMax="([-\d.]+)">(.*)<\/word>/.exec(line);
      if (word) words.push({ page, xMin: Number(word[1]), yMin: Number(word[2]), xMax: Number(word[3]), yMax: Number(word[4]), text: decode(word[5]) });
    }
    return { words, complaints };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}

/** A page of a PDF as Poppler paints it (see popplerPixels), by its number from 1. */
export function popplerPage(file: Uint8Array, pageNumber: number, dpi = 72): { width: number; height: number; at(x: number, y: number): [number, number, number]; data: Uint8Array } {
  const folder = join(process.cwd(), `tmp-pdf-check-${process.pid}-${++serial}`);
  mkdirSync(folder, { recursive: true });
  try {
    const path = join(folder, 'file.pdf');
    writeFileSync(path, file);
    const ppm = execFileSync('pdftoppm', ['-r', String(dpi), '-f', String(pageNumber), '-l', String(pageNumber), path], { maxBuffer: 512 * 1024 * 1024 });
    const head = /^P6\s+(\d+)\s+(\d+)\s+255\s/.exec(ppm.subarray(0, 40).toString('latin1'));
    if (!head) throw new Error('pdftoppm gave no picture');
    const width = Number(head[1]);
    const height = Number(head[2]);
    const data = ppm.subarray(head[0].length);
    return { width, height, data, at: (x, y) => [data[(y * width + x) * 3], data[(y * width + x) * 3 + 1], data[(y * width + x) * 3 + 2]] };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}
