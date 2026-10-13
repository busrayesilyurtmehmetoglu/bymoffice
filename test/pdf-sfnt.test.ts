// Reading font files for a PDF: what a font says of its glyphs and of its
// licence, and what happens to a file that is broken on purpose. The
// TrueType fonts are made by code (pdf-font-fixtures.ts); the OpenType font
// with CFF outlines is the one the package brings along.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { FontError, MAX_FONT_BYTES, fontLicence, readFont } from '../src/pdf/sfnt';
import { TEST_GLYPHS, bonum, seeded, sfntFile, testFont, testFontTables } from './pdf-font-fixtures';

test('font: a TrueType file says which glyph a character is and how wide the glyphs are', () => {
  const face = readFont(testFont());
  assert.equal(face.kind, 'truetype');
  assert.equal(face.unitsPerEm, 1000);
  assert.equal(face.glyphCount, TEST_GLYPHS.length);
  assert.equal(face.name, 'BymTest-Regular');
  assert.equal(face.glyph(0x41), 2, 'the glyph of "A"');
  assert.equal(face.glyph(0x15f), 6, 'the glyph of "ş"');
  assert.equal(face.glyph(0x130), 7, 'the glyph of "İ"');
  assert.equal(face.glyph(0x5a), 0, 'a character the font does not have');
  assert.equal(face.glyph(0x1f600), 0, 'a character beyond the BMP in a font with a map of the BMP');
  assert.equal(face.glyph(-1), 0, 'no character at all');
  assert.equal(face.glyph(NaN), 0, 'no number at all');
  assert.equal(face.advance(2), 700);
  assert.equal(face.advance(5), 800);
  assert.equal(face.advance(60000), 900, 'a glyph beyond the advances has the last advance');
  assert.equal(face.ascent, 800);
  assert.equal(face.descent, -200);
  assert.equal(face.capHeight, 700);
  assert.equal(face.fsType, 0);
  assert.equal(face.weight, 400);
  assert.deepEqual(face.bbox.slice(), [0, -200, 900, 900]);
});

test('font: a character map of format 12 reaches characters beyond the BMP', () => {
  const face = readFont(testFont({ cmapFormat: 12, map: { 0x41: 2, 0x1f600: 3, 0x10ffff: 4 } }));
  assert.equal(face.glyph(0x41), 2);
  assert.equal(face.glyph(0x1f600), 3);
  assert.equal(face.glyph(0x10ffff), 4);
  assert.equal(face.glyph(0x1f601), 0);
  assert.equal(face.glyph(0x110000), 0, 'beyond Unicode');
});

test('font: a character map that leads to a glyph the font does not have gives no glyph', () => {
  const face = readFont(testFont({ map: { 0x41: 2, 0x42: 500 } }));
  assert.equal(face.glyph(0x41), 2);
  assert.equal(face.glyph(0x42), 0);
});

test('font: the font of the package is read: OpenType with CFF outlines, Turkish letters, a licence that allows embedding', () => {
  for (const which of ['regular', 'bold', 'italic', 'bolditalic'] as const) {
    const face = readFont(bonum(which));
    assert.equal(face.kind, 'cff', which);
    assert.equal(face.unitsPerEm, 1000, which);
    assert.match(face.name, /^TeXGyreBonum-/, which);
    for (const char of 'ğüşıöçĞÜŞİÖÇ AaZz09.,') assert.ok(face.glyph(char.codePointAt(0) as number) > 0, `${which} has a glyph for "${char}"`);
    assert.equal(face.glyph(0x20ba), 0, `${which} has no lira sign`);
    assert.ok(face.advance(face.glyph(0x41)) > 500 && face.advance(face.glyph(0x41)) < 900, `${which}: the advance of "A" is that of a letter`);
    assert.equal(face.fsType, 12, which);
    assert.deepEqual(fontLicence(face), { embed: true, subset: true }, which);
    assert.equal(face.italicAngle !== 0, which === 'italic' || which === 'bolditalic', `${which}: the angle of the font`);
  }
  assert.ok(readFont(bonum('bold')).weight >= 600, 'the bold face says that it is bold');
});

test('font: the licence bits decide whether a font may be embedded', () => {
  assert.deepEqual(fontLicence({ fsType: 0 }), { embed: true, subset: true });
  assert.deepEqual(fontLicence({ fsType: 0x0004 }), { embed: true, subset: true }, 'preview and print');
  assert.deepEqual(fontLicence({ fsType: 0x0008 }), { embed: true, subset: true }, 'editable');
  assert.equal(fontLicence({ fsType: 0x0002 }).embed, false, 'restricted licence embedding');
  assert.equal(fontLicence({ fsType: 0x0002 | 0x0008 }).embed, false, 'restricted wins over what else is said');
  assert.equal(fontLicence({ fsType: 0x0200 }).embed, false, 'bitmap embedding only');
  assert.equal(fontLicence({ fsType: 0x0200 | 0x0004 }).embed, false, 'bitmap embedding only, with another bit');
  assert.deepEqual(fontLicence({ fsType: 0x0100 }), { embed: true, subset: false }, 'no subsetting: the whole font or none');
  assert.equal(fontLicence({ fsType: undefined }).embed, false, 'a font that states no licence is not embedded');
  assert.match(String(fontLicence({ fsType: 0x0002 }).reason), /licence/);
  assert.equal(readFont(testFont({ os2: false })).fsType, undefined);
  assert.equal(readFont(testFont({ fsType: 0x0002 })).fsType, 2);
});

test('font: what is no TrueType or OpenType font is refused with a reason', () => {
  const refuse = (bytes: Uint8Array, why: RegExp, what: string) => assert.throws(() => readFont(bytes), (error: unknown) => error instanceof FontError && why.test(error.message), what);
  refuse(new Uint8Array(4), /shorter/, 'four bytes');
  refuse(Uint8Array.from([0x77, 0x4f, 0x46, 0x46, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), /WOFF file/, 'a WOFF file');
  refuse(Uint8Array.from([0x77, 0x4f, 0x46, 0x32, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), /WOFF2 file/, 'a WOFF2 file');
  refuse(Uint8Array.from([0x74, 0x74, 0x63, 0x66, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]), /collection/, 'a collection');
  refuse(new Uint8Array(64), /no TrueType or OpenType/, 'zeros');
  refuse('not bytes' as unknown as Uint8Array, /must be bytes/, 'a string');
  for (const tag of ['head', 'hhea', 'hmtx', 'maxp', 'cmap', 'loca']) refuse(testFont({ without: [tag] }), new RegExp(`no ${tag} table`), `a font without ${tag}`);
  refuse(testFont({ without: ['glyf'] }), /neither TrueType nor CFF/, 'a font without outlines');
  // Only its length is looked at before a file is refused for it.
  const huge = { length: MAX_FONT_BYTES + 1 };
  Object.setPrototypeOf(huge, Uint8Array.prototype);
  refuse(huge as unknown as Uint8Array, /is not read/, 'a file beyond the limit');
});

test('font: a font whose glyphs are not what its outlines alone say is refused: variations, colours, bitmaps', () => {
  // A browser draws such a font otherwise than its plain outlines: at another weight, in colours, or from pictures.
  for (const [tag, why] of [['fvar', /variations/], ['COLR', /colour/], ['SVG ', /colour/], ['sbix', /colour/], ['CBDT', /colour/]] as [string, RegExp][]) {
    assert.throws(() => readFont(testFont({ extra: { [tag]: [0, 0, 0, 0] } })), (error: unknown) => error instanceof FontError && why.test(error.message), tag);
  }
});

test('font: tables that lie about themselves are refused, not followed', () => {
  const refuse = (change: (tables: { [tag: string]: number[] }) => void, why: RegExp, what: string) => {
    const tables = testFontTables();
    change(tables);
    assert.throws(() => readFont(sfntFile(tables)), (error: unknown) => error instanceof FontError && why.test(error.message), what);
  };
  refuse((t) => (t.head[12] = 0), /head table/, 'a head table without its magic number');
  refuse((t) => ((t.head[18] = 0), (t.head[19] = 1)), /units per em/, 'one unit per em');
  refuse((t) => ((t.maxp[4] = 0), (t.maxp[5] = 0)), /no glyphs/, 'no glyphs');
  refuse((t) => ((t.maxp[4] = 0xff), (t.maxp[5] = 0xff)), /loca table of the font is too short/, 'more glyphs than the loca table has');
  refuse((t) => ((t.hhea[34] = 0), (t.hhea[35] = 0)), /number of advances/, 'no advances');
  refuse((t) => ((t.hhea[34] = 0xff), (t.hhea[35] = 0xff)), /number of advances/, 'more advances than glyphs');
  refuse((t) => (t.hmtx.length = 8), /hmtx table of the font is shorter/, 'an hmtx table cut short');
  refuse((t) => ((t.head[50] = 0), (t.head[51] = 7)), /how its glyphs are indexed/, 'an unknown loca format');
  // Segments the table has no room for.
  refuse((t) => ((t.cmap[18] = 0xff), (t.cmap[19] = 0xfe)), /character map/, 'a character map with more segments than bytes');
  refuse((t) => (t.cmap[7] = 0), /no character map for Unicode/, 'a character map of symbols only');
  // A directory that points beyond the file.
  const file = testFont();
  const cut = file.slice(0, file.length - 40);
  assert.throws(() => readFont(cut), (error: unknown) => error instanceof FontError && /beyond the end/.test(error.message), 'a file cut short');
});

test('font: a font file with random bytes changed never hangs and never reads outside itself', () => {
  const random = seeded(20261013);
  const sources = [testFont(), testFont({ cmapFormat: 12 }), bonum('regular')];
  let read = 0;
  let refused = 0;
  const started = Date.now();
  for (let round = 0; round < 600; round++) {
    const source = sources[round % sources.length];
    const broken = source.slice();
    // A few bytes anywhere, or in the directory and the first tables, where most of the structure is.
    const changes = 1 + Math.floor(random() * 6);
    for (let i = 0; i < changes; i++) {
      const at = random() < 0.5 ? Math.floor(random() * Math.min(broken.length, 4096)) : Math.floor(random() * broken.length);
      broken[at] = random() < 0.3 ? 0xff : Math.floor(random() * 256);
    }
    const file = random() < 0.15 ? broken.slice(0, Math.floor(random() * broken.length)) : broken;
    try {
      const face = readFont(file);
      // What was read answers every question without throwing anything but a FontError.
      for (const code of [0x20, 0x41, 0x130, 0xffff, 0x10000, 0x10ffff]) {
        const glyph = face.glyph(code);
        assert.ok(glyph >= 0 && glyph < face.glyphCount, `round ${round}: a glyph of the font`);
        assert.ok(face.advance(glyph) >= 0 && face.advance(glyph) <= 0xffff, `round ${round}: an advance of the font`);
      }
      read++;
    } catch (error) {
      assert.ok(error instanceof FontError, `round ${round}: a broken font is refused with a FontError, not with ${String(error)}`);
      refused++;
    }
  }
  assert.ok(refused > 20, `some of the broken files are refused (${refused})`);
  assert.ok(read > 20, `some of them can still be read (${read})`);
  assert.ok(Date.now() - started < 20000, 'no broken file takes long');
});
