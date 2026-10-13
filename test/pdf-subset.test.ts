// A TrueType font cut down to the glyphs a file uses: what stays in it, that
// it is still a sound font file, and what happens with a font whose glyphs
// are broken on purpose.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { FontError, readFont } from '../src/pdf/sfnt';
import { glyphClosure, subsetTrueType } from '../src/pdf/subset';
import { checkTrueType } from './pdf-fixtures';
import { TEST_GLYPHS, seeded, sfntFile, testFont, testFontTables } from './pdf-font-fixtures';

const u16 = (bytes: Uint8Array, at: number): number => (bytes[at] << 8) | bytes[at + 1];
const u32 = (bytes: Uint8Array, at: number): number => bytes[at] * 16777216 + ((bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]);

/** The bytes of every glyph of a font file that passed checkTrueType(). */
function glyphsOf(file: Uint8Array): Uint8Array[] {
  const checked = checkTrueType(file);
  const long = u16(checked.table('head'), 50) === 1;
  const loca = checked.table('loca');
  const glyf = checked.table('glyf');
  const out: Uint8Array[] = [];
  for (let i = 0; i < checked.glyphs; i++) {
    const from = long ? u32(loca, i * 4) : u16(loca, i * 2) * 2;
    const to = long ? u32(loca, (i + 1) * 4) : u16(loca, (i + 1) * 2) * 2;
    out.push(glyf.subarray(from, to));
  }
  return out;
}

test('subset: only the glyphs that are used keep their outlines, under the numbers they had', () => {
  const file = testFont();
  const original = glyphsOf(file);
  const subset = subsetTrueType(readFont(file), [2, 6]);
  const checked = checkTrueType(subset);
  assert.equal(checked.glyphs, TEST_GLYPHS.length, 'the glyphs keep their numbers');
  const kept = glyphsOf(subset);
  for (let i = 0; i < kept.length; i++) {
    if (i === 0 || i === 2 || i === 6) assert.deepEqual(Array.from(kept[i]), Array.from(original[i]), `glyph ${i} is in the subset as it was, its instructions too`);
    else assert.equal(kept[i].length, 0, `glyph ${i} is not in the subset`);
  }
  assert.ok(subset.length < file.length, 'the subset is smaller than the font');
  // The advances are those of the font.
  const hmtx = checked.table('hmtx');
  TEST_GLYPHS.forEach((glyph, index) => assert.equal(u16(hmtx, index * 4), glyph.advance, `the advance of glyph ${index}`));
  assert.deepEqual(checked.tables, ['glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp'], 'the tables a PDF reads, and no others');
});

test('subset: a glyph made of other glyphs brings them along', () => {
  const face = readFont(testFont());
  assert.deepEqual(glyphClosure(face, [5]), [0, 2, 4, 5]);
  const kept = glyphsOf(subsetTrueType(face, [5]));
  assert.ok(kept[2].length > 0 && kept[4].length > 0 && kept[5].length > 0, 'the glyph and its two parts');
  assert.equal(kept[3].length, 0, 'and not the one next to them');
});

test('subset: the same glyphs give the same bytes, in whatever order they are named', () => {
  const face = readFont(testFont());
  assert.deepEqual(Array.from(subsetTrueType(face, [6, 2, 2, 3])), Array.from(subsetTrueType(face, [3, 6, 2])));
});

test('subset: the programs the instructions of the glyphs call stay in the font', () => {
  const file = testFont({ extra: { 'cvt ': [0, 10, 0, 20], fpgm: [0xb0, 0x00, 0x2c, 0x2d], prep: [0xb0, 0x01, 0x1d] } });
  const checked = checkTrueType(subsetTrueType(readFont(file), [2]));
  assert.deepEqual(Array.from(checked.table('cvt ')), [0, 10, 0, 20]);
  assert.deepEqual(Array.from(checked.table('fpgm')), [0xb0, 0x00, 0x2c, 0x2d]);
  assert.deepEqual(Array.from(checked.table('prep')), [0xb0, 0x01, 0x1d]);
});

test('subset: glyph numbers the font does not have are refused', () => {
  const face = readFont(testFont());
  assert.throws(() => subsetTrueType(face, [2, 400]), (error: unknown) => error instanceof FontError && /glyph 400/.test(error.message));
  assert.throws(() => subsetTrueType(face, [-1]), (error: unknown) => error instanceof FontError);
  assert.throws(() => subsetTrueType(face, [1.5]), (error: unknown) => error instanceof FontError);
});

test('subset: glyphs that are made of each other in a circle end', () => {
  const glyphs = TEST_GLYPHS.slice();
  glyphs[5] = { advance: 800, parts: [{ glyph: 9, dx: 0, dy: 0 }] };
  glyphs[9] = { advance: 900, parts: [{ glyph: 5, dx: 0, dy: 0 }] };
  const face = readFont(testFont({ glyphs }));
  assert.deepEqual(glyphClosure(face, [5]), [0, 5, 9]);
  checkTrueType(subsetTrueType(face, [5]));
});

test('subset: a glyph that names a part the font does not have is refused', () => {
  const glyphs = TEST_GLYPHS.slice();
  glyphs[5] = { advance: 800, parts: [{ glyph: 2, dx: 0, dy: 0 }] };
  const tables = testFontTables({ glyphs });
  // The number of the part: the last glyph record that is made of parts, bytes 12 and 13.
  const face0 = readFont(sfntFile(tables));
  assert.deepEqual(glyphClosure(face0, [5]), [0, 2, 5]);
  const loca = tables.loca;
  const at = ((loca[10] << 8) | loca[11]) * 2;
  tables.glyf[at + 12] = 0x7f;
  tables.glyf[at + 13] = 0xff;
  const face = readFont(sfntFile(tables));
  assert.throws(() => subsetTrueType(face, [5]), (error: unknown) => error instanceof FontError && /part/.test(error.message));
});

test('subset: glyphs that lie outside the glyf table are refused: when the font is read, and again should the font change afterwards', () => {
  const tables = testFontTables();
  // The end of glyph 2 before its start.
  tables.loca[6] = 0;
  tables.loca[7] = 1;
  assert.throws(() => readFont(sfntFile(tables)), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  const beyond = testFontTables();
  beyond.loca[6] = 0xff;
  beyond.loca[7] = 0xff;
  assert.throws(() => readFont(sfntFile(beyond)), (error: unknown) => error instanceof FontError && /loca/.test(error.message));
  // A font that was sound when it was read and is not any more: the subset is not made of what lies outside.
  for (const [high, low] of [[0, 1], [0xff, 0xff]]) {
    const bytes = sfntFile(testFontTables());
    const face = readFont(bytes);
    const loca = (face.tables.get('loca') as { offset: number }).offset;
    bytes[loca + 6] = high;
    bytes[loca + 7] = low;
    assert.throws(() => subsetTrueType(face, [2]), (error: unknown) => error instanceof FontError && /glyf/.test(error.message));
  }
});

test('subset: a font with broken glyphs gives a sound subset or is refused, and never hangs', () => {
  const random = seeded(77);
  const started = Date.now();
  let made = 0;
  let refused = 0;
  for (let round = 0; round < 400; round++) {
    const tables = testFontTables();
    // One of the two tables, a byte or two: where the glyphs are, or what they are made of.
    const table = tables[round % 4 === 0 ? 'loca' : 'glyf'];
    const changes = 1 + Math.floor(random() * 2);
    for (let i = 0; i < changes; i++) table[Math.floor(random() * table.length)] = random() < 0.3 ? 0xff : Math.floor(random() * 256);
    try {
      const face = readFont(sfntFile(tables));
      const subset = subsetTrueType(face, [2, 3, 5, 9]);
      checkTrueType(subset);
      made++;
    } catch (error) {
      assert.ok(error instanceof FontError, `round ${round}: refused with a FontError, not with ${String(error)}`);
      refused++;
    }
  }
  assert.ok(made > 10 && refused > 10, `both happen (${made} made, ${refused} refused)`);
  assert.ok(Date.now() - started < 20000, 'no broken font takes long');
});
