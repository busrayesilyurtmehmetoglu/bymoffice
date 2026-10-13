// The font of the invisible text of a PDF. The pages are pictures; the text
// lies over them, unseen, so that it can be found, selected and copied. That
// text needs a font, and the font needs no shapes: it is a TrueType file of a
// few hundred bytes, made here, with one empty glyph that stands for every
// character. Which character a code of the text is says the ToUnicode map.
//
// Codes are handed out as characters come: the first character of a document
// is code 1, the next new one code 2. A character outside the BMP is one
// code, and the map gives its pair of surrogates, which is how readers of
// PDF files expect it.

import { zlibDeflate } from './flate';
import { PdfString, latin1, name, type PdfRef, type PdfWriter } from './writer';

/** Advance of the one glyph, in thousandths of the font size: a character of the text is one em wide before it is scaled. */
export const GLYPH_WIDTH = 1000;
export const FONT_NAME = 'BYMGlyphless';
/** Codes one font holds (two bytes, 0 is not used). Characters beyond them are written as U+FFFD. */
export const MAX_FONT_CODES = 65535;

const UNITS_PER_EM = 1000;
const ASCENT = 800;
const DESCENT = -200;

class Bytes {
  readonly out: number[] = [];
  u8(value: number): this {
    this.out.push(value & 255);
    return this;
  }
  u16(value: number): this {
    this.out.push((value >> 8) & 255, value & 255);
    return this;
  }
  u32(value: number): this {
    this.out.push((value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255);
    return this;
  }
  zeros(count: number): this {
    for (let i = 0; i < count; i++) this.out.push(0);
    return this;
  }
  tag(text: string): this {
    for (let i = 0; i < 4; i++) this.out.push(text.charCodeAt(i));
    return this;
  }
}

/** Sum of the big endian 32 bit words of a table, the last one padded with zeros. */
function checksum(bytes: readonly number[]): number {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) {
    const word = (bytes[i] || 0) * 16777216 + (bytes[i + 1] || 0) * 65536 + (bytes[i + 2] || 0) * 256 + (bytes[i + 3] || 0);
    sum = (sum + word) >>> 0;
  }
  return sum;
}

let cachedFont: Uint8Array | undefined;

/**
 * A TrueType font of two glyphs without an outline (".notdef" and the one
 * every character is mapped to), each one em wide. The tables are those a
 * font file must have, so that programs that look at the font on its own
 * take it too; a PDF reads glyf, head, hhea, hmtx, loca and maxp.
 */
export function glyphlessFont(): Uint8Array {
  if (cachedFont) return cachedFont.slice();
  const tables: { [tag: string]: number[] } = {};

  tables['OS/2'] = new Bytes()
    .u16(0) // version
    .u16(GLYPH_WIDTH) // xAvgCharWidth
    .u16(400) // usWeightClass
    .u16(5) // usWidthClass
    .u16(0) // fsType: may be embedded
    .zeros(20) // sub- and superscript, strikeout
    .u16(0) // sFamilyClass
    .zeros(10) // panose
    .zeros(16) // ulUnicodeRange
    .tag('BYM ')
    .u16(0x40) // fsSelection: regular
    .u16(0) // usFirstCharIndex
    .u16(0xffff) // usLastCharIndex
    .u16(ASCENT)
    .u16(DESCENT)
    .u16(0) // sTypoLineGap
    .u16(ASCENT) // usWinAscent
    .u16(-DESCENT).out; // usWinDescent

  // One table for all of Unicode: every character is glyph 1 (format 13, "many to one").
  tables.cmap = new Bytes().u16(0).u16(1).u16(3).u16(10).u32(12).u16(13).u16(0).u32(28).u32(0).u32(1).u32(0).u32(0x10ffff).u32(1).out;

  // No glyph has an outline.
  tables.glyf = [];

  tables.head = new Bytes()
    .u32(0x00010000) // version
    .u32(0x00010000) // fontRevision
    .u32(0) // checkSumAdjustment, set below
    .u32(0x5f0f3cf5) // magicNumber
    .u16(0x000b) // flags
    .u16(UNITS_PER_EM)
    .zeros(16) // created, modified
    .zeros(8) // xMin, yMin, xMax, yMax
    .u16(0) // macStyle
    .u16(8) // lowestRecPPEM
    .u16(2) // fontDirectionHint
    .u16(0) // indexToLocFormat: short
    .u16(0).out; // glyphDataFormat

  tables.hhea = new Bytes()
    .u32(0x00010000)
    .u16(ASCENT)
    .u16(DESCENT)
    .u16(0) // lineGap
    .u16(GLYPH_WIDTH) // advanceWidthMax
    .zeros(6) // minLeftSideBearing, minRightSideBearing, xMaxExtent
    .u16(1) // caretSlopeRise
    .zeros(4) // caretSlopeRun, caretOffset
    .zeros(8)
    .u16(0) // metricDataFormat
    .u16(2).out; // numberOfHMetrics

  tables.hmtx = new Bytes().u16(GLYPH_WIDTH).u16(0).u16(GLYPH_WIDTH).u16(0).out;
  tables.loca = new Bytes().u16(0).u16(0).u16(0).out;
  // Version 1.0; two glyphs; maxZones 2; everything else is 0.
  tables.maxp = new Bytes().u32(0x00010000).u16(2).zeros(8).u16(2).zeros(16).out;

  const names: [number, string][] = [
    [1, FONT_NAME],
    [2, 'Regular'],
    [4, FONT_NAME],
    [6, FONT_NAME],
  ];
  const records = new Bytes().u16(0).u16(names.length).u16(6 + names.length * 12);
  const strings = new Bytes();
  for (let i = 0; i < names.length; i++) {
    const text = names[i][1];
    records.u16(3).u16(1).u16(0x0409).u16(names[i][0]).u16(text.length * 2).u16(strings.out.length);
    for (let c = 0; c < text.length; c++) strings.u16(text.charCodeAt(c));
  }
  tables.name = records.out.concat(strings.out);

  // Version 3: no names of glyphs.
  tables.post = new Bytes().u32(0x00030000).u32(0).u16(-100).u16(50).u32(1).zeros(16).out;

  // Tags in the order of their bytes: capitals come first.
  const tags = Object.keys(tables).sort();
  const count = tags.length;
  let power = 1;
  let selector = 0;
  while (power * 2 <= count) {
    power *= 2;
    selector++;
  }
  const header = new Bytes().u32(0x00010000).u16(count).u16(power * 16).u16(selector).u16(count * 16 - power * 16);
  const body: number[] = [];
  let offset = 12 + count * 16;
  let headAt = 0;
  for (let i = 0; i < count; i++) {
    const table = tables[tags[i]];
    header.tag(tags[i]).u32(checksum(table)).u32(offset).u32(table.length);
    if (tags[i] === 'head') headAt = offset;
    for (let b = 0; b < table.length; b++) body.push(table[b]);
    while (body.length % 4 !== 0) body.push(0);
    offset = 12 + count * 16 + body.length;
  }
  const file = header.out.concat(body);
  // With the adjustment in place the whole file sums to 0xB1B0AFBA.
  const adjustment = (0xb1b0afba - checksum(file)) >>> 0;
  file[headAt + 8] = (adjustment >>> 24) & 255;
  file[headAt + 9] = (adjustment >>> 16) & 255;
  file[headAt + 10] = (adjustment >>> 8) & 255;
  file[headAt + 11] = adjustment & 255;
  cachedFont = Uint8Array.from(file);
  return cachedFont.slice();
}

const HEX = '0123456789abcdef';
const hex4 = (value: number): string => HEX.charAt((value >> 12) & 15) + HEX.charAt((value >> 8) & 15) + HEX.charAt((value >> 4) & 15) + HEX.charAt(value & 15);

const REPLACEMENT = 0xfffd;

/**
 * What a character of a document is in the text of a PDF: itself; a space
 * for a tab or a line end; U+FFFD for half a surrogate pair and for a code
 * point that is no character; -1 for a control, which is left out.
 */
function textCodePoint(code: number): number {
  if (code < 0x20) return code === 9 || code === 10 || code === 13 ? 0x20 : -1;
  if (code >= 0x7f && code < 0xa0) return -1;
  if (code >= 0xd800 && code <= 0xdfff) return REPLACEMENT;
  if ((code & 0xfffe) === 0xfffe || (code >= 0xfdd0 && code <= 0xfdef)) return REPLACEMENT;
  return code;
}

/** The font of the text of one file: hands out the codes and writes itself once all text is known. */
export class PdfTextFont {
  private readonly byCodePoint = new Map<number, number>();
  /** The code point of every code; index 0 is not used. */
  private readonly codePoints: number[] = [0];

  /** Codes handed out so far. */
  get codes(): number {
    return this.codePoints.length - 1;
  }

  private codeOf(codePoint: number): number {
    const known = this.byCodePoint.get(codePoint);
    if (known !== undefined) return known;
    // The last code is kept for U+FFFD, which stands for every character that comes after the font is full.
    if (codePoint !== REPLACEMENT && this.codePoints.length >= MAX_FONT_CODES) return this.codeOf(REPLACEMENT);
    const code = this.codePoints.length;
    this.codePoints.push(codePoint);
    this.byCodePoint.set(codePoint, code);
    return code;
  }

  /**
   * Text as the string of a text showing operator: four hexadecimal digits
   * per character. `count` is the number of characters written, which is what
   * the width of the string is made of.
   */
  encode(text: string): { hex: string; count: number } {
    const parts: string[] = [];
    let piece = '';
    let count = 0;
    for (let i = 0; i < text.length; i++) {
      let code = text.charCodeAt(i);
      if (code >= 0xd800 && code <= 0xdbff && i + 1 < text.length) {
        const next = text.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
          i++;
        }
      }
      const codePoint = textCodePoint(code);
      if (codePoint < 0) continue;
      piece += hex4(this.codeOf(codePoint));
      count++;
      if (piece.length >= 8192) {
        parts.push(piece);
        piece = '';
      }
    }
    parts.push(piece);
    return { hex: parts.join(''), count };
  }

  /** The CMap that says which character every code is: UTF-16BE, a pair of surrogates for a character outside the BMP. */
  toUnicode(): string {
    const lines: string[] = [
      '/CIDInit /ProcSet findresource begin',
      '12 dict begin',
      'begincmap',
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
      '/CMapName /Adobe-Identity-UCS def',
      '/CMapType 2 def',
      '1 begincodespacerange',
      '<0000> <ffff>',
      'endcodespacerange',
    ];
    const total = this.codes;
    for (let first = 1; first <= total; first += 100) {
      const last = Math.min(first + 99, total);
      lines.push(last - first + 1 + ' beginbfchar');
      for (let code = first; code <= last; code++) {
        const codePoint = this.codePoints[code];
        const target = codePoint < 0x10000 ? hex4(codePoint) : hex4(0xd800 + ((codePoint - 0x10000) >> 10)) + hex4(0xdc00 + ((codePoint - 0x10000) & 0x3ff));
        lines.push('<' + hex4(code) + '> <' + target + '>');
      }
      lines.push('endbfchar');
    }
    lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
    return lines.join('\n') + '\n';
  }

  /** Writes the font under the number of `ref`: the Type 0 font, its descendant, the descriptor, the font file and the two maps. */
  async write(writer: PdfWriter, ref: PdfRef, native = true): Promise<void> {
    const flate = name('FlateDecode');
    const font = glyphlessFont();
    const file = writer.addStream({ Filter: flate, Length1: font.length }, await zlibDeflate(font, native));
    const descriptor = writer.add({
      Type: name('FontDescriptor'),
      FontName: name(FONT_NAME),
      // Fixed pitch, symbolic.
      Flags: 5,
      FontBBox: [0, DESCENT, GLYPH_WIDTH, ASCENT],
      ItalicAngle: 0,
      Ascent: ASCENT,
      Descent: DESCENT,
      CapHeight: ASCENT,
      StemV: 80,
      FontFile2: file,
    });
    // Every code is glyph 1; code 0, which no text uses, is ".notdef".
    const glyphs = new Uint8Array((this.codes + 1) * 2);
    for (let code = 1; code <= this.codes; code++) glyphs[code * 2 + 1] = 1;
    const toGlyph = writer.addStream({ Filter: flate }, await zlibDeflate(glyphs, native));
    const descendant = writer.add({
      Type: name('Font'),
      Subtype: name('CIDFontType2'),
      BaseFont: name(FONT_NAME),
      CIDSystemInfo: { Registry: new PdfString(latin1('Adobe')), Ordering: new PdfString(latin1('Identity')), Supplement: 0 },
      FontDescriptor: descriptor,
      DW: GLYPH_WIDTH,
      CIDToGIDMap: toGlyph,
    });
    const toUnicode = writer.addStream({ Filter: flate }, await zlibDeflate(latin1(this.toUnicode()), native));
    writer.add({ Type: name('Font'), Subtype: name('Type0'), BaseFont: name(FONT_NAME), Encoding: name('Identity-H'), DescendantFonts: [descendant], ToUnicode: toUnicode }, ref);
  }
}
