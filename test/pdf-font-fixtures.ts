// Font files made by code, for the tests of the PDF fonts: a small TrueType
// font with real outlines (boxes and a triangle, so that what is painted can
// be found in a picture), built table by table from the OpenType
// specification. It shares no code with src/pdf.

import { readFileSync } from 'node:fs';
import { join } from 'node:path';

class Out {
  readonly bytes: number[] = [];
  u8(value: number): this {
    this.bytes.push(value & 255);
    return this;
  }
  u16(value: number): this {
    this.bytes.push((value >> 8) & 255, value & 255);
    return this;
  }
  i16(value: number): this {
    return this.u16(value < 0 ? value + 0x10000 : value);
  }
  u32(value: number): this {
    this.bytes.push((value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255);
    return this;
  }
  zeros(count: number): this {
    for (let i = 0; i < count; i++) this.bytes.push(0);
    return this;
  }
  text(value: string): this {
    for (let i = 0; i < value.length; i++) this.bytes.push(value.charCodeAt(i) & 255);
    return this;
  }
  pad(): this {
    while (this.bytes.length % 4 !== 0) this.bytes.push(0);
    return this;
  }
}

export interface TestGlyph {
  advance: number;
  /** Closed outlines of points on the curve, in units of the font. */
  contours?: [number, number][][];
  /** A glyph made of other glyphs, each moved by (dx, dy). */
  parts?: { glyph: number; dx: number; dy: number }[];
  /** Bytes of instructions put into the glyph: what a subset must carry along untouched. */
  instructions?: number[];
}

export interface TestFontOptions {
  glyphs?: TestGlyph[];
  /** Character to glyph. */
  map?: { [codePoint: number]: number };
  /** The format of the character map: 4 (default) or 12. */
  cmapFormat?: 4 | 12;
  fsType?: number;
  /** False leaves the OS/2 table out. */
  os2?: boolean;
  name?: string;
  unitsPerEm?: number;
  weight?: number;
  italicAngle?: number;
  /** Tables added as they are (cvt, fpgm, prep, or anything else). */
  extra?: { [tag: string]: number[] };
  /** Tables left out. */
  without?: string[];
}

const box = (x0: number, y0: number, x1: number, y1: number): [number, number][] => [
  [x0, y0],
  [x0, y1],
  [x1, y1],
  [x1, y0],
];

/**
 * The glyphs of the default test font, 1000 units per em:
 *   0 .notdef (an empty box), 1 space (nothing, 300 wide),
 *   2 "A": a full box from the baseline to 700, 600 wide in an advance of 700,
 *   3 "B": a triangle, advance 500,
 *   4 "C": a box below the baseline (a descender), advance 400,
 *   5 "D": made of glyphs 2 and 4, advance 800,
 *   6 "ş", 7 "İ", 8 "ğ": boxes of their own widths, 9: a glyph no character leads to.
 */
export const TEST_GLYPHS: TestGlyph[] = [
  { advance: 500, contours: [box(50, 0, 450, 700), box(100, 50, 400, 650).reverse() as [number, number][]] },
  { advance: 300 },
  { advance: 700, contours: [box(50, 0, 650, 700)], instructions: [0xb0, 0x01] },
  {
    advance: 500,
    contours: [
      [
        [50, 0],
        [250, 700],
        [450, 0],
      ],
    ],
  },
  { advance: 400, contours: [box(50, -200, 350, 300)] },
  {
    advance: 800,
    parts: [
      { glyph: 2, dx: 0, dy: 0 },
      { glyph: 4, dx: 350, dy: 0 },
    ],
  },
  { advance: 450, contours: [box(50, 0, 400, 500)] },
  { advance: 350, contours: [box(50, 0, 300, 800)] },
  { advance: 550, contours: [box(50, -150, 500, 500)] },
  { advance: 900, contours: [box(0, 0, 900, 900)] },
];

export const TEST_MAP: { [codePoint: number]: number } = { 0x20: 1, 0x41: 2, 0x42: 3, 0x43: 4, 0x44: 5, 0x15f: 6, 0x130: 7, 0x11f: 8 };

function glyphBytes(glyph: TestGlyph, all: TestGlyph[], depth = 0): { data: number[]; box: [number, number, number, number] } {
  if (glyph.parts) {
    let bounds: [number, number, number, number] = [0, 0, 0, 0];
    let first = true;
    for (const part of glyph.parts) {
      const inner = depth < 8 ? glyphBytes(all[part.glyph], all, depth + 1).box : ([0, 0, 0, 0] as [number, number, number, number]);
      const moved: [number, number, number, number] = [inner[0] + part.dx, inner[1] + part.dy, inner[2] + part.dx, inner[3] + part.dy];
      bounds = first ? moved : [Math.min(bounds[0], moved[0]), Math.min(bounds[1], moved[1]), Math.max(bounds[2], moved[2]), Math.max(bounds[3], moved[3])];
      first = false;
    }
    const out = new Out().i16(-1).i16(bounds[0]).i16(bounds[1]).i16(bounds[2]).i16(bounds[3]);
    glyph.parts.forEach((part, index) => {
      // Arguments are words and offsets; all but the last say that more follow.
      out.u16(0x0003 | (index + 1 < (glyph.parts as unknown[]).length ? 0x0020 : 0)).u16(part.glyph).i16(part.dx).i16(part.dy);
    });
    return { data: out.pad().bytes, box: bounds };
  }
  const contours = glyph.contours || [];
  if (contours.length === 0) return { data: [], box: [0, 0, 0, 0] };
  const points = contours.reduce((list, contour) => list.concat(contour), [] as [number, number][]);
  const bounds: [number, number, number, number] = [Math.min(...points.map((p) => p[0])), Math.min(...points.map((p) => p[1])), Math.max(...points.map((p) => p[0])), Math.max(...points.map((p) => p[1]))];
  const out = new Out().i16(contours.length).i16(bounds[0]).i16(bounds[1]).i16(bounds[2]).i16(bounds[3]);
  let end = -1;
  for (const contour of contours) {
    end += contour.length;
    out.u16(end);
  }
  const instructions = glyph.instructions || [];
  out.u16(instructions.length);
  for (const byte of instructions) out.u8(byte);
  // Every point on the curve, its coordinates as words.
  for (let i = 0; i < points.length; i++) out.u8(1);
  let last = 0;
  for (const point of points) {
    out.i16(point[0] - last);
    last = point[0];
  }
  last = 0;
  for (const point of points) {
    out.i16(point[1] - last);
    last = point[1];
  }
  return { data: out.pad().bytes, box: bounds };
}

function cmap4(map: { [codePoint: number]: number }): number[] {
  const codes = Object.keys(map)
    .map(Number)
    .filter((code) => code < 0xffff)
    .sort((a, b) => a - b);
  // One segment per character: the simplest table that is right.
  const segments = codes.map((code) => ({ start: code, end: code, delta: (map[code] - code) & 0xffff }));
  segments.push({ start: 0xffff, end: 0xffff, delta: 1 });
  const count = segments.length;
  let power = 1;
  let selector = 0;
  while (power * 2 <= count) {
    power *= 2;
    selector++;
  }
  const out = new Out()
    .u16(4)
    .u16(16 + count * 8)
    .u16(0)
    .u16(count * 2)
    .u16(power * 2)
    .u16(selector)
    .u16(count * 2 - power * 2);
  for (const segment of segments) out.u16(segment.end);
  out.u16(0);
  for (const segment of segments) out.u16(segment.start);
  for (const segment of segments) out.u16(segment.delta);
  for (let i = 0; i < count; i++) out.u16(0);
  return out.bytes;
}

function cmap12(map: { [codePoint: number]: number }): number[] {
  const codes = Object.keys(map)
    .map(Number)
    .sort((a, b) => a - b);
  const out = new Out()
    .u16(12)
    .u16(0)
    .u32(16 + codes.length * 12)
    .u32(0)
    .u32(codes.length);
  for (const code of codes) out.u32(code).u32(code).u32(map[code]);
  return out.bytes;
}

function checksum(bytes: readonly number[]): number {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) sum = (sum + (bytes[i] || 0) * 16777216 + (bytes[i + 1] || 0) * 65536 + (bytes[i + 2] || 0) * 256 + (bytes[i + 3] || 0)) >>> 0;
  return sum;
}

/** A font file of the given tables: the directory, the checksums and the adjustment of the head table. */
export function sfntFile(tables: { [tag: string]: number[] }, version = 0x00010000): Uint8Array {
  const tags = Object.keys(tables).sort();
  let power = 1;
  let selector = 0;
  while (power * 2 <= tags.length) {
    power *= 2;
    selector++;
  }
  const out = new Out()
    .u32(version)
    .u16(tags.length)
    .u16(power * 16)
    .u16(selector)
    .u16(tags.length * 16 - power * 16);
  let offset = 12 + tags.length * 16;
  let headAt = -1;
  for (const tag of tags) {
    const data = tables[tag];
    if (tag === 'head') headAt = offset;
    out.text(tag).u32(checksum(data)).u32(offset).u32(data.length);
    offset += (data.length + 3) & ~3;
  }
  for (const tag of tags) {
    for (const byte of tables[tag]) out.u8(byte);
    out.pad();
  }
  if (headAt >= 0) {
    const adjustment = (0xb1b0afba - checksum(out.bytes)) >>> 0;
    out.bytes[headAt + 8] = (adjustment >>> 24) & 255;
    out.bytes[headAt + 9] = (adjustment >>> 16) & 255;
    out.bytes[headAt + 10] = (adjustment >>> 8) & 255;
    out.bytes[headAt + 11] = adjustment & 255;
  }
  return Uint8Array.from(out.bytes);
}

/** The tables of the test font, each as its bytes: for tests that break one of them. */
export function testFontTables(options: TestFontOptions = {}): { [tag: string]: number[] } {
  const glyphs = options.glyphs || TEST_GLYPHS;
  const map = options.map || TEST_MAP;
  const unitsPerEm = options.unitsPerEm || 1000;
  const made = glyphs.map((glyph) => glyphBytes(glyph, glyphs));
  const glyf = new Out();
  const loca = new Out();
  for (const glyph of made) {
    loca.u16(glyf.bytes.length / 2);
    for (const byte of glyph.data) glyf.u8(byte);
  }
  loca.u16(glyf.bytes.length / 2);
  const bounds: [number, number, number, number] = [Math.min(...made.map((g) => g.box[0])), Math.min(...made.map((g) => g.box[1])), Math.max(...made.map((g) => g.box[2])), Math.max(...made.map((g) => g.box[3]))];
  const tables: { [tag: string]: number[] } = {};
  tables.head = new Out()
    .u32(0x00010000)
    .u32(0x00010000)
    .u32(0)
    .u32(0x5f0f3cf5)
    .u16(0x000b)
    .u16(unitsPerEm)
    .zeros(16)
    .i16(bounds[0])
    .i16(bounds[1])
    .i16(bounds[2])
    .i16(bounds[3])
    .u16(0)
    .u16(8)
    .i16(2)
    .i16(0)
    .i16(0).bytes;
  tables.hhea = new Out()
    .u32(0x00010000)
    .i16(800)
    .i16(-200)
    .i16(0)
    .u16(Math.max(...glyphs.map((g) => g.advance)))
    .zeros(22)
    .u16(glyphs.length).bytes;
  tables.maxp = new Out().u32(0x00010000).u16(glyphs.length).u16(64).u16(4).u16(64).u16(4).u16(2).zeros(12).u16(2).u16(2).bytes;
  const hmtx = new Out();
  glyphs.forEach((glyph, index) => hmtx.u16(glyph.advance).i16(made[index].box[0]));
  tables.hmtx = hmtx.bytes;
  tables.loca = loca.bytes;
  tables.glyf = glyf.bytes;
  const sub = options.cmapFormat === 12 ? cmap12(map) : cmap4(map);
  tables.cmap = new Out()
    .u16(0)
    .u16(1)
    .u16(3)
    .u16(options.cmapFormat === 12 ? 10 : 1)
    .u32(12).bytes.concat(sub);
  if (options.os2 !== false) {
    tables['OS/2'] = new Out()
      .u16(2)
      .i16(500)
      .u16(options.weight || 400)
      .u16(5)
      .u16(options.fsType || 0)
      .zeros(20)
      .i16(0)
      .zeros(10)
      .zeros(16)
      .text('TEST')
      .u16(0x0040)
      .u16(0x20)
      .u16(0x15f)
      .i16(800)
      .i16(-200)
      .i16(0)
      .u16(900)
      .u16(200)
      .zeros(8)
      .i16(500)
      .i16(700)
      .u16(0)
      .u16(0x20)
      .u16(1).bytes;
  }
  const psName = options.name === undefined ? 'BymTest-Regular' : options.name;
  const name = new Out().u16(0).u16(1).u16(18).u16(3).u16(1).u16(0x0409).u16(6).u16(psName.length * 2).u16(0);
  for (let i = 0; i < psName.length; i++) name.u16(psName.charCodeAt(i));
  tables.name = name.bytes;
  tables.post = new Out()
    .u32(0x00030000)
    .u32(Math.round((options.italicAngle || 0) * 65536) >>> 0)
    .i16(-100)
    .i16(50)
    .u32(0)
    .zeros(16).bytes;
  if (options.extra) for (const tag of Object.keys(options.extra)) tables[tag] = options.extra[tag];
  if (options.without) for (const tag of options.without) delete tables[tag];
  return tables;
}

/** The test font as a file. */
export function testFont(options: TestFontOptions = {}): Uint8Array {
  return sfntFile(testFontTables(options));
}

export interface TestLookup {
  /** The type of the lookup: 4 for ligatures. With `extension`, the lookup is of type 7 and wraps one of this type. */
  type: number;
  extension?: boolean;
  /** For type 4: [the glyph of the ligature, its first glyph, the glyphs after it...], in the order the font tries them. */
  ligatures?: number[][];
}

export interface TestGsub {
  lookups: TestLookup[];
  /** The features, each with the lookups it names. */
  features: { tag: string; lookups: number[] }[];
  /** The scripts; without them one "DFLT" whose default language has every feature. `languages` are further language systems. */
  scripts?: { tag: string; features: number[]; languages?: { tag: string; features: number[] }[] }[];
}

/** A GSUB table, as the OpenType specification lays it out. A lookup of another type than 4 holds one subtable of two zero bytes after its format. */
export function gsubTable(gsub: TestGsub): number[] {
  const u16 = (value: number): number[] => [(value >> 8) & 255, value & 255];
  const tag = (text: string): number[] => [0, 1, 2, 3].map((i) => text.charCodeAt(i));
  const scripts = gsub.scripts || [{ tag: 'DFLT', features: gsub.features.map((_feature, index) => index) }];
  const langSys = (features: number[]): number[] => [...u16(0), ...u16(0xffff), ...u16(features.length), ...features.reduce((list, index) => list.concat(u16(index)), [] as number[])];
  // The script list.
  let scriptList = u16(scripts.length);
  let scriptTables: number[] = [];
  const scriptHead = 2 + scripts.length * 6;
  for (const script of scripts) {
    scriptList = scriptList.concat(tag(script.tag), u16(scriptHead + scriptTables.length));
    const languages = script.languages || [];
    const head = 4 + languages.length * 6;
    const own = langSys(script.features);
    let table = [...u16(head), ...u16(languages.length)];
    let after = own.slice();
    for (const language of languages) {
      table = table.concat(tag(language.tag), u16(head + after.length));
      after = after.concat(langSys(language.features));
    }
    scriptTables = scriptTables.concat(table, after);
  }
  scriptList = scriptList.concat(scriptTables);
  // The feature list.
  let featureList = u16(gsub.features.length);
  let featureTables: number[] = [];
  const featureHead = 2 + gsub.features.length * 6;
  for (const feature of gsub.features) {
    featureList = featureList.concat(tag(feature.tag), u16(featureHead + featureTables.length));
    featureTables = featureTables.concat(u16(0), u16(feature.lookups.length), feature.lookups.reduce((list, index) => list.concat(u16(index)), [] as number[]));
  }
  featureList = featureList.concat(featureTables);
  // The lookups.
  const ligatureSubtable = (ligatures: number[][]): number[] => {
    const firsts = Array.from(new Set(ligatures.map((ligature) => ligature[1]))).sort((a, b) => a - b);
    const sets = firsts.map((first) => {
      const own = ligatures.filter((ligature) => ligature[1] === first);
      let head = u16(own.length);
      let body: number[] = [];
      for (const ligature of own) {
        head = head.concat(u16(2 + own.length * 2 + body.length));
        body = body.concat(u16(ligature[0]), u16(ligature.length - 1), ligature.slice(2).reduce((list, glyph) => list.concat(u16(glyph)), [] as number[]));
      }
      return head.concat(body);
    });
    let out = [...u16(1), ...u16(0), ...u16(firsts.length)];
    let at = 6 + firsts.length * 2;
    let bodies: number[] = [];
    for (const set of sets) {
      out = out.concat(u16(at + bodies.length));
      bodies = bodies.concat(set);
    }
    const coverageAt = at + bodies.length;
    out[2] = coverageAt >> 8;
    out[3] = coverageAt & 255;
    return out.concat(bodies, u16(1), u16(firsts.length), firsts.reduce((list, glyph) => list.concat(u16(glyph)), [] as number[]));
  };
  let lookupList = u16(gsub.lookups.length);
  let lookupTables: number[] = [];
  const lookupHead = 2 + gsub.lookups.length * 2;
  for (const lookup of gsub.lookups) {
    lookupList = lookupList.concat(u16(lookupHead + lookupTables.length));
    let subtable = lookup.type === 4 ? ligatureSubtable(lookup.ligatures || []) : [...u16(1), 0, 0];
    // An extension: its format, the type it stands for, and where the subtable lies (right behind it).
    if (lookup.extension) subtable = [...u16(1), ...u16(lookup.type), 0, 0, 0, 8, ...subtable];
    lookupTables = lookupTables.concat(u16(lookup.extension ? 7 : lookup.type), u16(0), u16(1), u16(8), subtable);
  }
  lookupList = lookupList.concat(lookupTables);
  const head = [0, 1, 0, 0, ...u16(10), ...u16(10 + scriptList.length), ...u16(10 + scriptList.length + featureList.length)];
  return head.concat(scriptList, featureList, lookupList);
}

/** A face of the font the package brings along (OpenType with CFF outlines). */
export function bonum(face: 'regular' | 'bold' | 'italic' | 'bolditalic' = 'regular'): Uint8Array {
  return new Uint8Array(readFileSync(join(process.cwd(), 'assets', 'fonts', 'tex-gyre-bonum', `texgyrebonum-${face}.otf`)));
}

/** A generator of numbers that is the same in every run. */
export function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}
