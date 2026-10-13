// A font of the text of a PDF, with its font file in the file: a Type 0
// font over a CID font whose codes are the glyph numbers of the font
// (Identity-H), the widths of the glyphs that are used, and the map that
// says which character each of them is, so that the text can be found and
// copied.
//
// A TrueType font is cut down to the glyphs that are used (subset.ts) unless
// its licence forbids that; it then goes in whole. An OpenType font with CFF
// outlines goes in whole, as its CFF table: nothing of it is rewritten.
// UNSURE: whether a subset of the font of the package may be handed on under
// its licence was not settled, so no font with CFF outlines is cut down.

import { plainerText } from './canvas';
import { adler32, zlibDeflate } from './flate';
import { fontLicence, fontTable, type PdfFontFace } from './sfnt';
import { subsetTrueType } from './subset';
import { PdfString, latin1, name, pdfNumber, type PdfDict, type PdfRef, type PdfValue, type PdfWriter } from './writer';

const HEX = '0123456789abcdef';
const hex4 = (value: number): string => HEX.charAt((value >> 12) & 15) + HEX.charAt((value >> 8) & 15) + HEX.charAt((value >> 4) & 15) + HEX.charAt(value & 15);

const digests = new WeakMap<PdfFontFace, number>();

/** A number that changes with the bytes of a font file. */
function faceDigest(face: PdfFontFace): number {
  let digest = digests.get(face);
  if (digest === undefined) {
    digest = adler32(face.bytes) ^ face.bytes.length;
    digests.set(face, digest);
  }
  return digest;
}

/**
 * Six capital letters that change with the glyphs of a subset and with the
 * font it was cut from: what stands before the name of a font that is not
 * whole. Two fonts of one name, cut down to the same glyphs, have two tags.
 */
export function subsetTag(glyphs: readonly number[], face?: PdfFontFace): string {
  let hash = (0x811c9dc5 ^ (face ? faceDigest(face) : 0)) >>> 0;
  for (let i = 0; i < glyphs.length; i++) {
    hash ^= glyphs[i] & 255;
    hash = Math.imul(hash, 0x01000193) >>> 0;
    hash ^= glyphs[i] >> 8;
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  let tag = '';
  for (let i = 0; i < 6; i++) {
    tag += String.fromCharCode(65 + (hash % 26));
    hash = Math.floor(hash / 26) + (i + 1) * 7919;
  }
  return tag;
}

/** How many characters one glyph may stand for in the map to Unicode. */
const MAX_GLYPH_TEXT = 16;

/** The map from the codes of a font (its glyph numbers) to the text each one stands for, as a CMap: UTF-16, several characters for a ligature. */
export function glyphsToUnicode(glyphs: ReadonlyMap<number, string>): string {
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
  const codes = Array.from(glyphs.keys())
    .filter((glyph) => glyph > 0 && glyph <= 0xffff)
    .sort((a, b) => a - b);
  for (let first = 0; first < codes.length; first += 100) {
    const last = Math.min(first + 100, codes.length);
    lines.push(last - first + ' beginbfchar');
    for (let i = first; i < last; i++) {
      const text = (glyphs.get(codes[i]) as string).slice(0, MAX_GLYPH_TEXT);
      let target = '';
      for (let k = 0; k < text.length; k++) target += hex4(text.charCodeAt(k));
      lines.push('<' + hex4(codes[i]) + '> <' + (target || 'fffd') + '>');
    }
    lines.push('endbfchar');
  }
  lines.push('endcmap', 'CMapName currentdict /CMap defineresource pop', 'end', 'end');
  return lines.join('\n') + '\n';
}

/** A font of a file: the glyphs the pages use of it, and how it is written once all pages are. */
export class PdfEmbeddedFont {
  /** Glyph to the text it stands for (the plainest it was used for: see plainerText). */
  readonly glyphs = new Map<number, string>();

  constructor(
    readonly face: PdfFontFace,
    readonly ref: PdfRef,
  ) {}

  /** Takes note of glyphs a page uses. */
  use(glyphs: ReadonlyMap<number, string>): void {
    glyphs.forEach((text, glyph) => {
      // The plainest text of a glyph, whatever page used it first: the file does not depend on the order of its pages.
      const known = this.glyphs.get(glyph);
      this.glyphs.set(glyph, known === undefined ? text : plainerText(known, text));
    });
  }

  /**
   * Writes the font under its number: the Type 0 font, its descendant, the
   * descriptor, the font file and the map to Unicode. `taken` are the names
   * the fonts of the file have so far: this one gets a name of its own, and
   * adds it. Returns the text of the map to Unicode.
   */
  async write(writer: PdfWriter, native = true, taken?: Set<string>): Promise<string> {
    const face = this.face;
    const flate = name('FlateDecode');
    const used = Array.from(this.glyphs.keys()).sort((a, b) => a - b);
    const perEm = 1000 / face.unitsPerEm;
    const scaled = (value: number): number => Math.round(value * perEm);
    let baseFont = face.name;
    const descriptor: PdfDict = {
      Type: name('FontDescriptor'),
      // Symbolic: the codes of the font are no characters of a standard set.
      Flags: 4 | (face.fixedPitch ? 1 : 0) | (face.serif ? 2 : 0) | (face.italicAngle !== 0 ? 64 : 0),
      FontBBox: [scaled(face.bbox[0]), scaled(face.bbox[1]), scaled(face.bbox[2]), scaled(face.bbox[3])],
      ItalicAngle: Number(pdfNumber(face.italicAngle, 2)),
      Ascent: scaled(face.ascent),
      Descent: scaled(face.descent),
      CapHeight: scaled(face.capHeight),
      // What the font does not say: a guess from its weight, as other writers of PDF files make it.
      StemV: Math.round(50 + (face.weight / 65) * (face.weight / 65)),
    };
    const descendant: PdfDict = { Type: name('Font'), CIDSystemInfo: { Registry: new PdfString(latin1('Adobe')), Ordering: new PdfString(latin1('Identity')), Supplement: 0 } };
    // The font program first: whatever is wrong with the font is found before anything of it is in the file.
    let file: Uint8Array | undefined;
    if (face.kind === 'truetype') {
      if (fontLicence(face).subset) {
        file = subsetTrueType(face, used);
        baseFont = subsetTag(used, face) + '+' + face.name;
      } else file = face.bytes;
    }
    // No two fonts of a file under one name: readers of PDF files take fonts of one name for one font.
    if (taken) {
      const wanted = baseFont;
      for (let n = 2; taken.has(baseFont); n++) baseFont = wanted + '-' + n;
      taken.add(baseFont);
    }
    if (file) {
      descriptor.FontFile2 = writer.addStream({ Filter: flate, Length1: file.length }, await zlibDeflate(file, native));
      descendant.Subtype = name('CIDFontType2');
      descendant.CIDToGIDMap = name('Identity');
    } else {
      // The glyph programs of the font as they are in its file.
      const table = fontTable(face, 'CFF ');
      const cff = face.bytes.subarray(table.start, table.end);
      descriptor.FontFile3 = writer.addStream({ Filter: flate, Subtype: name('CIDFontType0C') }, await zlibDeflate(cff, native));
      descendant.Subtype = name('CIDFontType0');
    }
    descriptor.FontName = name(baseFont);
    descendant.BaseFont = name(baseFont);
    descendant.FontDescriptor = writer.add(descriptor);
    // The widths of the glyphs that are used: runs of neighbouring glyphs, each with its list.
    const widths: PdfValue[] = [];
    for (let i = 0; i < used.length; ) {
      const run: number[] = [];
      const first = used[i];
      while (i < used.length && used[i] === first + run.length) {
        run.push(Number(pdfNumber(face.advance(used[i]) * perEm, 2)));
        i++;
      }
      widths.push(first, run);
    }
    descendant.DW = Number(pdfNumber(face.advance(0) * perEm, 2));
    if (widths.length > 0) descendant.W = widths;
    const toUnicode = glyphsToUnicode(this.glyphs);
    const map = writer.addStream({ Filter: flate }, await zlibDeflate(latin1(toUnicode), native));
    writer.add({ Type: name('Font'), Subtype: name('Type0'), BaseFont: name(baseFont), Encoding: name('Identity-H'), DescendantFonts: [writer.add(descendant)], ToUnicode: map }, this.ref);
    return toUnicode;
  }
}
