// Builds synthetic Word binary (.doc) files for the tests of the doc reader.
//
// Nothing here comes from a real document: a DocBuilder is told the text,
// which characters and paragraphs have which sprms, and the structures of the
// table stream as bytes; it lays out the WordDocument stream (header, text,
// formatted disk pages), the table stream and the Data stream, and wraps them
// in a compound file.

import { CfbFile } from '../src/cfb';
import { readDoc } from '../src/formats/doc/document';
import { FC } from '../src/formats/doc/fib';
import type { Block, DocModel, Paragraph, Run, Table } from '../src/model/doc/types';
import { paragraphText } from '../src/model/doc/types';

export { FC };

// ------------------------------------------------------------------- bytes

export const u8s = (...values: number[]): Uint8Array => new Uint8Array(values);
export const u16le = (value: number): Uint8Array => new Uint8Array([value & 0xff, (value >> 8) & 0xff]);
export const u32le = (value: number): Uint8Array => new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]);

export function cat(...parts: (Uint8Array | number[])[]): Uint8Array {
  let length = 0;
  for (const part of parts) length += part.length;
  const out = new Uint8Array(length);
  let p = 0;
  for (const part of parts) {
    out.set(part, p);
    p += part.length;
  }
  return out;
}

export function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    out[i * 2] = text.charCodeAt(i) & 0xff;
    out[i * 2 + 1] = text.charCodeAt(i) >> 8;
  }
  return out;
}

/** A sprm: the code, then the operand as bytes or as numbers that are bytes. */
export const sprm = (op: number, ...operand: (number | Uint8Array)[]): Uint8Array =>
  cat(u16le(op), ...operand.map((part) => (typeof part === 'number' ? u8s(part) : part)));

/** A PLC: n+1 positions and n data elements. */
export function plc(cps: number[], data: Uint8Array[] = []): Uint8Array {
  return cat(...cps.map(u32le), ...data);
}

/** An STTB of 16-bit strings, each with `extra` bytes after it. */
export function sttb(strings: string[], extra: Uint8Array[] = [], cbExtra = 0): Uint8Array {
  const parts: Uint8Array[] = [u16le(0xffff), u16le(strings.length), u16le(cbExtra)];
  strings.forEach((text, i) => {
    parts.push(u16le(text.length), utf16le(text));
    if (cbExtra) parts.push(extra[i] || new Uint8Array(cbExtra));
  });
  return cat(...parts);
}

// --------------------------------------------------------------- compound file

/** A version 3 compound file with the streams in regular sectors (the mini stream cutoff is 0). */
export function buildCfb(streams: [name: string, bytes: Uint8Array][]): Uint8Array {
  const SECTOR = 512;
  const END = 0xfffffffe;
  const FREE = 0xffffffff;
  const dirSectors = Math.ceil((streams.length + 1) / 4);
  const sizes = streams.map(([, bytes]) => Math.ceil(bytes.length / SECTOR));
  const content = dirSectors + sizes.reduce((a, b) => a + b, 0);
  let fatSectors = 1;
  while (fatSectors * 128 < fatSectors + content) fatSectors++;
  if (fatSectors > 109) throw new Error('fixture too large for a compound file without DIFAT sectors');
  const total = fatSectors + content;
  const out = new Uint8Array((total + 1) * SECTOR);
  const view = new DataView(out.buffer);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, fatSectors, true); // first directory sector
  view.setUint32(56, 0, true); // mini stream cutoff: nothing goes into the mini stream
  view.setUint32(60, END, true);
  view.setUint32(68, END, true);
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, i < fatSectors ? i : FREE, true);

  const at = (sector: number) => (sector + 1) * SECTOR;
  const fat = (sector: number, value: number) => view.setUint32(at(0) + sector * 4, value, true);
  for (let i = 0; i < fatSectors * 128; i++) fat(i, FREE);
  for (let i = 0; i < fatSectors; i++) fat(i, 0xfffffffd);
  for (let i = 0; i < dirSectors; i++) fat(fatSectors + i, i === dirSectors - 1 ? END : fatSectors + i + 1);

  const entry = (index: number, name: string, type: number, child: number, right: number, start: number, size: number) => {
    const p = at(fatSectors) + index * 128;
    for (let i = 0; i < name.length; i++) view.setUint16(p + i * 2, name.charCodeAt(i), true);
    view.setUint16(p + 64, (name.length + 1) * 2, true);
    view.setUint8(p + 66, type);
    view.setUint32(p + 68, FREE, true);
    view.setUint32(p + 72, right, true);
    view.setUint32(p + 76, child, true);
    view.setUint32(p + 116, start, true);
    view.setUint32(p + 120, size, true);
  };
  entry(0, 'Root Entry', 5, streams.length ? 1 : FREE, FREE, END, 0);
  let sector = fatSectors + dirSectors;
  streams.forEach(([name, bytes], i) => {
    const count = sizes[i];
    entry(i + 1, name, 2, FREE, i + 1 < streams.length ? i + 2 : FREE, count ? sector : END, bytes.length);
    out.set(bytes, at(sector));
    for (let k = 0; k < count; k++) fat(sector + k, k === count - 1 ? END : sector + k + 1);
    sector += count;
  });
  return out;
}

// -------------------------------------------------------------------- document

interface PieceSpec {
  cpStart: number;
  cpEnd: number;
  fc: number;
  compressed: boolean;
  prm: number;
}

interface RunSpec {
  from: number;
  to: number;
  data: Uint8Array;
}

/** Packs property runs into formatted disk pages; returns the pages and the bin table that lists them. */
function packPages(runs: { fcStart: number; fcEnd: number; data: Uint8Array }[], paragraph: boolean, firstPage: number): { pages: Uint8Array; bte: Uint8Array } {
  const entry = paragraph ? 13 : 1;
  const pages: Uint8Array[] = [];
  const firsts: number[] = [];
  let i = 0;
  while (i < runs.length) {
    const page = new Uint8Array(512);
    const view = new DataView(page.buffer);
    let count = 0;
    let top = 511; // data is packed downwards from here
    const places: number[] = [];
    while (i + count < runs.length && count < 100) {
      const data = runs[i + count].data;
      // What the data takes: its size bytes and itself, at an even offset.
      let need = 0;
      let encoded: Uint8Array | undefined;
      if (data.length) {
        if (paragraph) encoded = data.length % 2 ? cat([(data.length + 1) / 2], data) : cat([0, data.length / 2], data);
        else encoded = cat([data.length], data);
        need = encoded.length;
      }
      let place = top - need;
      if (place % 2) place--;
      const front = (count + 2) * 4 + (count + 1) * entry;
      if (need ? place < front : top < front) break;
      if (encoded) {
        page.set(encoded, place);
        top = place;
        places.push(place / 2);
      } else places.push(0);
      count++;
    }
    if (count === 0) throw new Error('fixture: a property list does not fit into a page');
    for (let k = 0; k < count; k++) {
      view.setUint32(k * 4, runs[i + k].fcStart, true);
      page[(count + 1) * 4 + k * entry] = places[k];
    }
    view.setUint32(count * 4, runs[i + count - 1].fcEnd, true);
    page[511] = count;
    firsts.push(runs[i].fcStart);
    pages.push(page);
    i += count;
  }
  const lastEnd = runs.length ? runs[runs.length - 1].fcEnd : 0;
  const bte = cat(...firsts.map(u32le), u32le(lastEnd), ...pages.map((_, k) => u32le(firstPage + k)));
  return { pages: cat(...pages), bte };
}

export interface BuildOptions {
  nFib?: number;
  wIdent?: number;
  /** Extra bits for the flags of the header (0x0100 encrypted, 0x0004 complex, ...). */
  flags?: number;
  /** Write the table stream under this name instead of "1Table". */
  tableName?: string;
  /** Leave the piece table out. */
  noClx?: boolean;
  /** Called with the finished streams before they are wrapped, to damage them. */
  patch?: (streams: { word: Uint8Array; table: Uint8Array; data: Uint8Array; fcLcb(index: number): [number, number] }) => void;
}

export class DocBuilder {
  /** Where text and blobs go in the WordDocument stream: after the header. */
  private readonly wordParts: Uint8Array[] = [];
  private wordLength = 0x800;
  private readonly pieces: PieceSpec[] = [];
  private cp = 0;
  private readonly chars: RunSpec[] = [];
  private readonly paras: RunSpec[] = [];
  private readonly tables = new Map<number, Uint8Array>();
  private readonly lists: Uint8Array[] = [];
  private data: Uint8Array = new Uint8Array(0);
  /** Lengths of the parts of the text; the main text is what is left. */
  ccp: { ftn?: number; hdd?: number; atn?: number; edn?: number; txbx?: number; hdrTxbx?: number } = {};

  /** Appends text as a piece of its own; returns the positions it covers. */
  text(text: string, options: { compressed?: boolean; prm?: number } = {}): { from: number; to: number } {
    const compressed = options.compressed === true;
    let bytes: Uint8Array;
    if (compressed) {
      bytes = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i);
    } else bytes = utf16le(text);
    const fc = this.blob(bytes);
    const from = this.cp;
    this.cp += text.length;
    this.pieces.push({ cpStart: from, cpEnd: this.cp, fc, compressed, prm: options.prm || 0 });
    return { from, to: this.cp };
  }

  /** Appends bytes to the WordDocument stream; returns their offset. */
  blob(bytes: Uint8Array): number {
    const at = this.wordLength;
    this.wordParts.push(bytes);
    this.wordLength += bytes.length;
    return at;
  }

  /** Character properties for the positions [from, to). */
  chpx(from: number, to: number, ...sprms: Uint8Array[]): this {
    this.chars.push({ from, to, data: cat(...sprms) });
    return this;
  }

  /** Paragraph properties for the paragraph whose text and mark are [from, to); `istd` is its style. */
  papx(from: number, to: number, istd = 0, ...sprms: Uint8Array[]): this {
    this.paras.push({ from, to, data: cat(u16le(istd), ...sprms) });
    return this;
  }

  /** Gives every paragraph of the text so far that has no properties yet the same ones. */
  papxAll(text: string, istd = 0, ...sprms: Uint8Array[]): this {
    let start = 0;
    for (let i = 0; i < text.length; i++) {
      const code = text.charCodeAt(i);
      if (code !== 13 && code !== 7) continue;
      this.papx(start, i + 1, istd, ...sprms);
      start = i + 1;
    }
    return this;
  }

  /** A structure of the table stream, by its index in the header (FC.*). */
  table(index: number, bytes: Uint8Array): this {
    this.tables.set(index, bytes);
    return this;
  }

  /** A property list of the piece table (a Prc); returns the Prm that refers to it. */
  prc(...sprms: Uint8Array[]): number {
    this.lists.push(cat(...sprms));
    return ((this.lists.length - 1) << 1) | 1;
  }

  /** Appends bytes to the Data stream; returns their offset. */
  dataBlob(bytes: Uint8Array): number {
    const at = this.data.length;
    this.data = cat(this.data, bytes);
    return at;
  }

  get length(): number {
    return this.cp;
  }

  private fcRuns(runs: RunSpec[]): { fcStart: number; fcEnd: number; data: Uint8Array }[] {
    const out: { fcStart: number; fcEnd: number; data: Uint8Array }[] = [];
    for (const run of runs) {
      for (const piece of this.pieces) {
        const from = Math.max(run.from, piece.cpStart);
        const to = Math.min(run.to, piece.cpEnd);
        if (to <= from) continue;
        const width = piece.compressed ? 1 : 2;
        out.push({ fcStart: piece.fc + (from - piece.cpStart) * width, fcEnd: piece.fc + (to - piece.cpStart) * width, data: run.data });
      }
    }
    out.sort((a, b) => a.fcStart - b.fcStart);
    // The offsets of a page are shared by neighbouring runs: gaps get runs without properties.
    const filled: typeof out = [];
    for (const run of out) {
      const last = filled[filled.length - 1];
      if (last && run.fcStart > last.fcEnd) filled.push({ fcStart: last.fcEnd, fcEnd: run.fcStart, data: new Uint8Array(0) });
      filled.push(run);
    }
    return filled;
  }

  /** The streams wrapped in a compound file. */
  build(options: BuildOptions = {}): Uint8Array {
    // Formatted disk pages go after the text, on a page boundary.
    const pad = (512 - (this.wordLength % 512)) % 512;
    let wordLength = this.wordLength + pad;
    const chpx = packPages(this.fcRuns(this.chars), false, wordLength / 512);
    wordLength += chpx.pages.length;
    const papx = packPages(this.fcRuns(this.paras), true, wordLength / 512);
    wordLength += papx.pages.length;

    const tables = new Map(this.tables);
    if (!tables.has(FC.PlcfBteChpx) && this.chars.length) tables.set(FC.PlcfBteChpx, chpx.bte);
    if (!tables.has(FC.PlcfBtePapx) && this.paras.length) tables.set(FC.PlcfBtePapx, papx.bte);
    if (!tables.has(FC.Clx) && !options.noClx) {
      const pcds = this.pieces.map((piece) => cat(u16le(0), u32le(piece.compressed ? (piece.fc * 2) | 0x40000000 : piece.fc), u16le(piece.prm)));
      const cps = [0, ...this.pieces.map((piece) => piece.cpEnd)];
      const body = cat(...cps.map(u32le), ...pcds);
      tables.set(FC.Clx, cat(...this.lists.map((list) => cat([1], u16le(list.length), list)), [2], u32le(body.length), body));
    }

    // The table stream: the structures one after the other, from offset 16 on.
    const tableParts: Uint8Array[] = [new Uint8Array(16)];
    let tableLength = 16;
    const pairs = new Map<number, [number, number]>();
    for (const [index, bytes] of tables) {
      pairs.set(index, [tableLength, bytes.length]);
      tableParts.push(bytes);
      tableLength += bytes.length;
    }
    // Room behind the last structure: lists keep their levels after the structure the header names.
    const table = cat(...tableParts, this.tail);

    const word = new Uint8Array(wordLength);
    const view = new DataView(word.buffer);
    view.setUint16(0, options.wIdent === undefined ? 0xa5ec : options.wIdent, true);
    view.setUint16(2, options.nFib === undefined ? 0x00c1 : options.nFib, true);
    view.setUint16(6, 0x0409, true);
    view.setUint16(10, 0x0200 | (options.flags || 0), true);
    view.setUint16(32, 14, true);
    view.setUint16(62, 22, true);
    const ccp = this.ccp;
    const other = (ccp.ftn || 0) + (ccp.hdd || 0) + (ccp.atn || 0) + (ccp.edn || 0) + (ccp.txbx || 0) + (ccp.hdrTxbx || 0);
    // With other parts the text ends with one more paragraph mark, which belongs to none of them.
    view.setUint32(76, this.cp - other - (other ? 1 : 0), true);
    view.setUint32(80, ccp.ftn || 0, true);
    view.setUint32(84, ccp.hdd || 0, true);
    view.setUint32(92, ccp.atn || 0, true);
    view.setUint32(96, ccp.edn || 0, true);
    view.setUint32(100, ccp.txbx || 0, true);
    view.setUint32(104, ccp.hdrTxbx || 0, true);
    view.setUint16(152, 93, true);
    for (const [index, [fc, lcb]] of pairs) {
      view.setUint32(154 + index * 8, fc, true);
      view.setUint32(154 + index * 8 + 4, lcb, true);
    }
    let p = 0x800;
    for (const part of this.wordParts) {
      word.set(part, p);
      p += part.length;
    }
    word.set(chpx.pages, this.wordLength + pad);
    word.set(papx.pages, this.wordLength + pad + chpx.pages.length);

    const streams = { word, table, data: this.data, fcLcb: (index: number): [number, number] => pairs.get(index) || [0, 0] };
    if (options.patch) options.patch(streams);
    const list: [string, Uint8Array][] = [
      ['WordDocument', streams.word],
      [options.tableName || '1Table', streams.table],
    ];
    if (streams.data.length) list.push(['Data', streams.data]);
    return buildCfb(list);
  }

  /** Bytes that follow the last structure of the table stream. */
  tail: Uint8Array = new Uint8Array(0);

  read(options?: BuildOptions): DocModel {
    return readDoc(CfbFile.read(this.build(options)));
  }
}

/** A document of plain paragraphs: `text` with "\r" after each paragraph. */
export function simpleDoc(text: string, options: { compressed?: boolean } = {}): DocBuilder {
  const builder = new DocBuilder();
  builder.text(text, options);
  builder.papxAll(text);
  return builder;
}

// ------------------------------------------------------------------ structures

export interface StyleSpec {
  /** Built-in style number; 0xFFE for a style of the user. */
  sti?: number;
  /** 1 paragraph, 2 character, 3 table, 4 list. */
  kind?: number;
  base?: number;
  next?: number;
  name: string;
  /** Paragraph sprms (without the style index in front). */
  papx?: Uint8Array;
  chpx?: Uint8Array;
  tapx?: Uint8Array;
}

/** A style sheet; `undefined` entries are empty slots. The fonts are the defaults for Latin, East Asian and other text. */
export function stsh(styles: (StyleSpec | undefined)[], fonts: [number, number, number] = [0, 0, 0]): Uint8Array {
  const header = cat(u16le(styles.length), u16le(10), u16le(1), u16le(0x5b), u16le(15), u16le(0), u16le(fonts[0]), u16le(fonts[1]), u16le(fonts[2]));
  const parts: Uint8Array[] = [u16le(header.length), header];
  styles.forEach((style, istd) => {
    if (!style) {
      parts.push(u16le(0));
      return;
    }
    const kind = style.kind || 1;
    const lists: Uint8Array[] = [];
    const withIstd = (bytes: Uint8Array | undefined) => cat(u16le(istd), bytes || new Uint8Array(0));
    if (kind === 1) lists.push(withIstd(style.papx), style.chpx || new Uint8Array(0));
    else if (kind === 2) lists.push(style.chpx || new Uint8Array(0));
    else if (kind === 3) lists.push(style.tapx || new Uint8Array(0), withIstd(style.papx), style.chpx || new Uint8Array(0));
    else lists.push(withIstd(style.papx));
    const base = style.base === undefined ? 0x0fff : style.base;
    const next = style.next === undefined ? istd : style.next;
    let std = cat(
      u16le(style.sti === undefined ? 0x0ffe : style.sti),
      u16le(kind | (base << 4)),
      u16le(lists.length | (next << 4)),
      u16le(0),
      u16le(0),
      u16le(style.name.length),
      utf16le(style.name),
      u16le(0),
    );
    for (const list of lists) {
      if (std.length % 2) std = cat(std, [0]);
      std = cat(std, u16le(list.length), list);
    }
    parts.push(u16le(std.length), std);
  });
  return cat(...parts);
}

/** A font table with the given names. */
export function fontTable(names: string[]): Uint8Array {
  const parts: Uint8Array[] = [u16le(names.length), u16le(0)];
  for (const name of names) {
    const ffn = cat(u8s(0x21), u16le(400), u8s(0, 0), new Uint8Array(10), new Uint8Array(24), utf16le(name), u16le(0));
    parts.push(u8s(ffn.length), ffn);
  }
  return cat(...parts);
}

/** sprmTDefTable for cells with the given edges (one more than there are cells); `cells` are their 20-byte descriptions. */
export function defTable(edges: number[], cells: Uint8Array[] = []): Uint8Array {
  const count = edges.length - 1;
  const body = cat(u8s(count), ...edges.map((edge) => u16le(edge & 0xffff)), ...cells);
  return cat(u16le(0xd608), u16le(body.length + 1), body);
}

/** A cell description (TC80): flags, preferred width and four borders (top, left, bottom, right) as 32-bit values. */
export function tc80(flags = 0, borders: number[] = [0, 0, 0, 0], width = 0): Uint8Array {
  return cat(u16le(flags), u16le(width), ...borders.map(u32le));
}

/** A Brc80: width in eighths of a point, type, colour index, distance. */
export const brc80 = (width: number, type: number, ico = 0, space = 0): number => (width | (type << 8) | (ico << 16) | (space << 24)) >>> 0;

// --------------------------------------------------------------------- reading

export const IN_TABLE = sprm(0x2416, 1);
export const ROW_END = sprm(0x2417, 1);

export function paragraphs(blocks: Block[]): Paragraph[] {
  return blocks.filter((block): block is Paragraph => block.type === 'paragraph');
}

export function tablesOf(blocks: Block[]): Table[] {
  return blocks.filter((block): block is Table => block.type === 'table');
}

export function runsOf(paragraph: Paragraph): Run[] {
  return paragraph.children.filter((inline): inline is Run => inline.type === 'run');
}

export const textOf = (blocks: Block[]): string[] => paragraphs(blocks).map(paragraphText);

/** The text of the cells of a table, row by row; paragraphs of a cell joined with "|". */
export const cellTexts = (table: Table): string[][] => table.rows.map((row) => row.cells.map((cell) => textOf(cell.blocks).join('|')));
