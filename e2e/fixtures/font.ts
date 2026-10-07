// A TrueType font made in code, for tests: no font file is kept in the
// repository. Every character from U+0020 to U+007E is the same box, and
// every glyph is as wide as the font is high (the advance is one em), so a
// text of n characters at s px is exactly n * s px wide in every browser.

const UNITS_PER_EM = 1000;
const FIRST_CHAR = 0x20;
const LAST_CHAR = 0x7e;
const ASCENT = 800;
const DESCENT = 200;

class Writer {
  readonly bytes: number[] = [];
  u8(value: number): this {
    this.bytes.push(value & 255);
    return this;
  }
  u16(value: number): this {
    this.bytes.push((value >>> 8) & 255, value & 255);
    return this;
  }
  i16(value: number): this {
    return this.u16(value < 0 ? value + 0x10000 : value);
  }
  u32(value: number): this {
    this.bytes.push((value >>> 24) & 255, (value >>> 16) & 255, (value >>> 8) & 255, value & 255);
    return this;
  }
  tag(text: string): this {
    for (let i = 0; i < 4; i++) this.u8(text.charCodeAt(i));
    return this;
  }
  zeros(count: number): this {
    for (let i = 0; i < count; i++) this.bytes.push(0);
    return this;
  }
}

function checksum(bytes: number[]): number {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) {
    sum = (sum + (((bytes[i] || 0) << 24) | ((bytes[i + 1] || 0) << 16) | ((bytes[i + 2] || 0) << 8) | (bytes[i + 3] || 0))) >>> 0;
  }
  return sum;
}

export interface TestFontOptions {
  /** The family name inside the file. Browsers use the name the font is registered with; this one only has to be there. */
  name?: string;
  /** Advance width of every glyph, in 1/1000 em. Default 1000: a character is as wide as the font size. */
  advance?: number;
}

/** The bytes of the font. */
export function testFont(options: TestFontOptions = {}): Uint8Array {
  const family = options.name || 'Bym Test';
  const advance = options.advance === undefined ? UNITS_PER_EM : options.advance;
  const glyphs = LAST_CHAR - FIRST_CHAR + 2; // .notdef and one per character
  const right = Math.max(200, advance - 100);

  // One contour, four points on the curve, clockwise: a box from (100, 0) to (right, 700).
  const glyph = new Writer().i16(1).i16(100).i16(0).i16(right).i16(700).u16(3).u16(0);
  for (let i = 0; i < 4; i++) glyph.u8(0x01);
  glyph.i16(100).i16(0).i16(right - 100).i16(0);
  glyph.i16(0).i16(700).i16(0).i16(-700);
  while (glyph.bytes.length % 4) glyph.u8(0);

  const glyf = new Writer();
  const loca = new Writer();
  for (let i = 0; i < glyphs; i++) {
    loca.u32(glyf.bytes.length);
    for (let k = 0; k < glyph.bytes.length; k++) glyf.u8(glyph.bytes[k]);
  }
  loca.u32(glyf.bytes.length);

  const head = new Writer()
    .u32(0x00010000)
    .u32(0x00010000)
    .u32(0) // checkSumAdjustment, set at the end
    .u32(0x5f0f3cf5)
    .u16(0x0003)
    .u16(UNITS_PER_EM)
    .zeros(16) // created, modified
    .i16(100)
    .i16(0)
    .i16(right)
    .i16(700)
    .u16(0) // macStyle
    .u16(8)
    .i16(2)
    .i16(1) // long offsets in loca
    .i16(0);

  const hhea = new Writer()
    .u32(0x00010000)
    .i16(ASCENT)
    .i16(-DESCENT)
    .i16(0)
    .u16(advance)
    .i16(100)
    .i16(advance - right)
    .i16(right)
    .i16(1)
    .i16(0)
    .i16(0)
    .zeros(8)
    .i16(0)
    .u16(glyphs);

  const maxp = new Writer().u32(0x00010000).u16(glyphs).u16(4).u16(1).u16(0).u16(0).u16(2).zeros(16);

  const os2 = new Writer()
    .u16(3)
    .i16(advance)
    .u16(400)
    .u16(5)
    .u16(0)
    .i16(650)
    .i16(600)
    .i16(0)
    .i16(75)
    .i16(650)
    .i16(600)
    .i16(0)
    .i16(350)
    .i16(50)
    .i16(300)
    .i16(0)
    .zeros(10) // panose
    .u32(1) // Basic Latin
    .u32(0)
    .u32(0)
    .u32(0)
    .tag('TEST')
    .u16(0x0040)
    .u16(FIRST_CHAR)
    .u16(LAST_CHAR)
    .i16(ASCENT)
    .i16(-DESCENT)
    .i16(0)
    .u16(ASCENT)
    .u16(DESCENT)
    .u32(1)
    .u32(0)
    .i16(500)
    .i16(700)
    .u16(0)
    .u16(FIRST_CHAR)
    .u16(1);

  const hmtx = new Writer();
  for (let i = 0; i < glyphs; i++) hmtx.u16(advance).i16(100);

  // Format 4: one segment for the characters (glyph = code - 31) and the closing one.
  const cmap = new Writer()
    .u16(0)
    .u16(1)
    .u16(3)
    .u16(1)
    .u32(12)
    .u16(4)
    .u16(32)
    .u16(0)
    .u16(4)
    .u16(4)
    .u16(1)
    .u16(0)
    .u16(LAST_CHAR)
    .u16(0xffff)
    .u16(0)
    .u16(FIRST_CHAR)
    .u16(0xffff)
    .i16(1 - FIRST_CHAR)
    .i16(1)
    .u16(0)
    .u16(0);

  const names: [number, string][] = [
    [1, family],
    [2, 'Regular'],
    [3, family + ' Regular 1.0'],
    [4, family],
    [5, 'Version 1.0'],
    [6, family.replace(/[^A-Za-z0-9]/g, '') + '-Regular'],
  ];
  const strings = new Writer();
  const name = new Writer().u16(0).u16(names.length).u16(6 + names.length * 12);
  for (const [id, text] of names) {
    name.u16(3).u16(1).u16(0x0409).u16(id).u16(text.length * 2).u16(strings.bytes.length);
    for (let i = 0; i < text.length; i++) strings.u16(text.charCodeAt(i));
  }
  for (let i = 0; i < strings.bytes.length; i++) name.u8(strings.bytes[i]);

  const post = new Writer().u32(0x00030000).u32(0).i16(-100).i16(50).u32(1).zeros(16);

  // In the order of their tags, as the table directory must be.
  const tables: [string, number[]][] = [
    ['OS/2', os2.bytes],
    ['cmap', cmap.bytes],
    ['glyf', glyf.bytes],
    ['head', head.bytes],
    ['hhea', hhea.bytes],
    ['hmtx', hmtx.bytes],
    ['loca', loca.bytes],
    ['maxp', maxp.bytes],
    ['name', name.bytes],
    ['post', post.bytes],
  ];
  const count = tables.length;
  let power = 1;
  let exponent = 0;
  while (power * 2 <= count) {
    power *= 2;
    exponent++;
  }
  const file = new Writer().u32(0x00010000).u16(count).u16(power * 16).u16(exponent).u16(count * 16 - power * 16);
  let offset = 12 + count * 16;
  let headAt = 0;
  for (const [tag, bytes] of tables) {
    if (tag === 'head') headAt = offset;
    file.tag(tag).u32(checksum(bytes)).u32(offset).u32(bytes.length);
    offset += (bytes.length + 3) & ~3;
  }
  for (const [, bytes] of tables) {
    for (let i = 0; i < bytes.length; i++) file.u8(bytes[i]);
    while (file.bytes.length % 4) file.u8(0);
  }
  const adjustment = (0xb1b0afba - checksum(file.bytes)) >>> 0;
  file.bytes[headAt + 8] = (adjustment >>> 24) & 255;
  file.bytes[headAt + 9] = (adjustment >>> 16) & 255;
  file.bytes[headAt + 10] = (adjustment >>> 8) & 255;
  file.bytes[headAt + 11] = adjustment & 255;
  return Uint8Array.from(file.bytes);
}
