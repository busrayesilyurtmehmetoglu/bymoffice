import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { LayoutFont } from '../src/layout/doc/types';
import { CanvasTextMeasurer, type MeasureContext } from '../src/render/doc-measure';
import { baseFontFamily, sanitizeFontFamily, splitFontFamily } from '../src/render/fonts';
import { cssFont, FONT_FALLBACK } from '../src/render/text';
import type { Font } from '../src/model/sheet/types';

const font = (family: string, size = 16, bold = false, italic = false): LayoutFont => ({ family, size, bold, italic });
const cellFont = (name: string): Font => ({ name, size: 12 }) as Font;

/**
 * A canvas of a device that has the `installed` families: a text is as wide
 * as its length times the size times the factor of the first family of the
 * font string that is installed (0.5 for the generic family at the end).
 */
function deviceContext(installed: { [family: string]: number }): MeasureContext {
  return {
    font: '',
    measureText(text: string) {
      const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1] || 10);
      let factor = 0.5;
      const quoted = /"([^"]+)"/g;
      for (let match = quoted.exec(this.font); match; match = quoted.exec(this.font)) {
        if (Object.prototype.hasOwnProperty.call(installed, match[1])) {
          factor = installed[match[1]];
          break;
        }
      }
      return { width: text.length * size * factor, fontBoundingBoxAscent: size * (factor + 0.3), fontBoundingBoxDescent: size * 0.3 };
    },
  };
}

test('font names: the code page suffixes of old Windows documents are told from the family', () => {
  const table: [string, string][] = [
    ['Bookman Old Style Tur', 'Bookman Old Style'],
    ['Arial Tur', 'Arial'],
    ['Arial TUR', 'Arial'],
    ['Times New Roman CE', 'Times New Roman'],
    ['Arial Cyr', 'Arial'],
    ['Arial CYR', 'Arial'],
    ['Courier New Greek', 'Courier New'],
    ['Tahoma Baltic', 'Tahoma'],
    ['Comic Sans MS Tur', 'Comic Sans MS'],
    ['Calibri Tur', 'Calibri'],
    ['Times New Roman (Turkish)', 'Times New Roman'],
    ['Arial (Hebrew)', 'Arial'],
    ['Arial (Arabic)', 'Arial'],
    ['Arial (Vietnamese)', 'Arial'],
    ['Tahoma (Thai)', 'Tahoma'],
    ['Arial (Greek)', 'Arial'],
    ['Arial (Cyrillic)', 'Arial'],
    ['Arial (Baltic)', 'Arial'],
    ['Arial (Central European)', 'Arial'],
    ['Arial (Western)', 'Arial'],
    ['Arial(Hebrew)', 'Arial'],
    ['  Arial   Tur  ', 'Arial'],
  ];
  for (const [name, base] of table) assert.equal(baseFontFamily(name), base, name);
});

test('font names: only a whole known suffix at the end of a longer name is one', () => {
  const same = [
    'Arial',
    'Futura',
    'Centaur',
    'Tur',
    'CE',
    ' Cyr',
    '(Hebrew)',
    'Arial Turkish',
    'Arial Hebrew',
    'Noto Sans Thai',
    'Noto Sans Arabic',
    'Arial Tur Bold',
    'ArialTur',
    'Arial-Tur',
    'Lettuce',
    'Arial (Bold)',
    'Arial (Hebrew',
    'Segoe UI',
    'MS Mincho',
    'Arial Tur2',
  ];
  // Nothing is taken off: what is left is the name as it is put into a font string.
  for (const name of same) assert.equal(baseFontFamily(name), sanitizeFontFamily(name), name);
  // One suffix only: what is left is a name a font may really have.
  assert.equal(baseFontFamily('Arial CE Tur'), 'Arial CE');
  // Names that are not names.
  assert.equal(baseFontFamily(''), '');
  assert.equal(baseFontFamily(undefined), '');
  assert.equal(baseFontFamily(42), '');
  assert.equal(baseFontFamily('";'), '');
  // Nothing of a family in front of the suffix: the name is left as it is, never an empty family.
  assert.deepEqual(splitFontFamily('; CE'), { full: 'CE', base: 'CE' });
  assert.deepEqual(splitFontFamily('"" (Hebrew)'), { full: 'Hebrew', base: 'Hebrew' });
  assert.deepEqual(splitFontFamily('() Tur'), { full: 'Tur', base: 'Tur' });
  assert.deepEqual(splitFontFamily(' Tur'), { full: 'Tur', base: 'Tur' });
  assert.equal(new CanvasTextMeasurer().cssFamilies('; CE'), '"CE", sans-serif');
  // The result is as safe as a sanitized name.
  assert.equal(baseFontFamily('x"; } body { color: red Tur'), 'x body color red');
  assert.deepEqual(splitFontFamily('Arial Tur'), { full: 'Arial Tur', base: 'Arial' });
  assert.deepEqual(splitFontFamily('Arial (Hebrew)'), { full: 'Arial Hebrew', base: 'Arial' });
  assert.deepEqual(splitFontFamily('Arial'), { full: 'Arial', base: 'Arial' });
  assert.ok(baseFontFamily('a'.repeat(100000) + ' Tur').length <= 128);
});

test('doc measure: a suffixed name is followed by its family and by what follows that family', () => {
  const measurer = new CanvasTextMeasurer();
  const bookman = '"Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif';
  assert.equal(measurer.cssFamilies('Bookman Old Style'), bookman);
  // The name of the file comes first (a font may really be called so); everything after it is the same.
  assert.equal(measurer.cssFamilies('Bookman Old Style Tur'), '"Bookman Old Style Tur", ' + bookman);
  assert.equal(measurer.cssFamilies('Times New Roman CE'), '"Times New Roman CE", "Times New Roman", "Liberation Serif", "Tinos", "Times", serif');
  assert.equal(measurer.cssFamilies('Arial (Hebrew)'), '"Arial Hebrew", "Arial", "Liberation Sans", "Arimo", "Helvetica", sans-serif');
  assert.equal(measurer.cssFamilies('Courier New Tur'), '"Courier New Tur", "Courier New", "Liberation Mono", "Cousine", "Courier", monospace');
  assert.equal(measurer.cssFamilies('Calibri Tur'), '"Calibri Tur", "Calibri", "Carlito", sans-serif');
  // Not a suffix: nothing changes.
  assert.equal(measurer.cssFamilies('Futura'), '"Futura", sans-serif');
  assert.equal(measurer.cssFamilies('Some Font'), '"Some Font", sans-serif');
});

test('doc measure: a suffixed name measures as its family does', () => {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({ 'Bookman Old Style': 0.7, Arial: 0.6 }) });
  assert.equal(measurer.width(font('Bookman Old Style', 20), 'abcd'), 56);
  assert.equal(measurer.width(font('Bookman Old Style Tur', 20), 'abcd'), 56);
  assert.equal(measurer.width(font('Arial CE', 20), 'abcd'), 48);
  // The substitute is found through the family as well.
  const other = new CanvasTextMeasurer({ context: deviceContext({ 'TeX Gyre Bonum': 0.8 }) });
  assert.equal(other.width(font('Bookman Old Style Tur', 10), 'ab'), 16);
  assert.deepEqual(other.metrics(font('Bookman Old Style Tur', 10)), other.metrics(font('Bookman Old Style', 10)));
});

test('doc measure: the line of a suffixed name is the line of its family', () => {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  assert.deepEqual(measurer.metrics(font('Times New Roman Tur', 2048)), { ascent: 1825, descent: 443, lineGap: 87 });
  assert.deepEqual(measurer.metrics(font('Arial CE', 2048)), { ascent: 1854, descent: 434, lineGap: 67 });
  assert.deepEqual(measurer.metrics(font('Tahoma (Thai)', 2048)), { ascent: 2049, descent: 423, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('Comic Sans MS Tur', 2048)), measurer.metrics(font('Comic Sans MS', 2048)));
});

test('doc measure: hasFont asks for the family of a suffixed name too', () => {
  const asked: string[] = [];
  const measurer = new CanvasTextMeasurer({
    hasFamily: (family) => {
      asked.push(family);
      return family === 'Arial';
    },
  });
  assert.equal(measurer.hasFont('Arial Tur'), true);
  assert.deepEqual(asked, ['Arial Tur', 'Arial']);
  assert.equal(new CanvasTextMeasurer({ hasFamily: () => false }).hasFont('Arial Tur'), false);
});

test('doc measure: the alternative name of the font table follows the family, before the substitutes', () => {
  const fonts = new Map([
    ['Bookman Old Style Tur', { family: 'roman', altName: 'Bookman' }],
    ['Odd Tur', { family: 'roman', altName: 'Odd Greek' }],
  ]);
  const measurer = new CanvasTextMeasurer({ fonts });
  assert.equal(measurer.cssFamilies('Bookman Old Style Tur'), '"Bookman Old Style Tur", "Bookman Old Style", "Bookman", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", serif');
  assert.equal(measurer.cssFamilies('Odd Tur'), '"Odd Tur", "Odd", "Odd Greek", serif');
});

test('doc measure: families that Apple and Android devices lack are followed by fonts with their metrics', () => {
  const measurer = new CanvasTextMeasurer();
  assert.equal(measurer.cssFamilies('Book Antiqua'), '"Book Antiqua", "TeX Gyre Pagella", "Palatino Linotype", "Palatino", "URW Palladio L", serif');
  assert.equal(measurer.cssFamilies('Palatino Linotype'), '"Palatino Linotype", serif');
  assert.equal(measurer.cssFamilies('Century Schoolbook'), '"Century Schoolbook", "TeX Gyre Schola", "Century Schoolbook L", serif');
  assert.equal(measurer.cssFamilies('Century Gothic'), '"Century Gothic", "TeX Gyre Adventor", "URW Gothic", sans-serif');
  // No font with the same widths is known for these: they stay as they were.
  assert.equal(measurer.cssFamilies('Tahoma'), '"Tahoma", sans-serif');
  assert.equal(measurer.cssFamilies('Garamond'), '"Garamond", serif');
  assert.equal(measurer.cssFamilies('Verdana'), '"Verdana", sans-serif');
});

test('doc measure: a family that is not in the table of line metrics takes the font box of the canvas', () => {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({ 'URW Bookman': 0.6 }) });
  // Bookman Old Style is not in the table, and neither is the font it is shown with here: the box of that font, no
  // leading. (Shown with a font of the table it takes the row of that font: fonts-embedded.test.ts.)
  const box = measurer.metrics(font('Bookman Old Style', 10));
  assert.ok(Math.abs(box.ascent - 9) < 1e-9 && Math.abs(box.descent - 3) < 1e-9 && box.lineGap === 0, JSON.stringify(box));
  const unknown = new CanvasTextMeasurer({ context: deviceContext({}) }).metrics(font('No Such Font', 10));
  assert.ok(Math.abs(unknown.ascent - 8) < 1e-9 && Math.abs(unknown.descent - 3) < 1e-9 && unknown.lineGap === 0, JSON.stringify(unknown));
});

test('doc measure: line metrics read from font files', () => {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  assert.deepEqual(measurer.metrics(font('Rockwell', 2048)), { ascent: 2099, descent: 600, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('Arial Rounded MT Bold', 2048)), { ascent: 1938, descent: 432, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('Brush Script MT', 2048)), { ascent: 1820, descent: 692, lineGap: 0 });
});

test('grid text: a suffixed name is followed by its family', () => {
  assert.equal(cssFont(cellFont('Arial Tur')), `16px "Arial Tur", "Arial", ${FONT_FALLBACK}`);
  assert.equal(cssFont(cellFont('Times New Roman (Hebrew)')), `16px "Times New Roman Hebrew", "Times New Roman", ${FONT_FALLBACK}`);
  assert.equal(cssFont(cellFont('Arial')), `16px "Arial", ${FONT_FALLBACK}`);
  assert.equal(cssFont(cellFont('Futura')), `16px "Futura", ${FONT_FALLBACK}`);
});

test('sources: no range of a regular expression ends in a raw U+FFFF', async () => {
  // A regular expression class written with the raw characters U+00C0 and U+FFFF cannot be read back
  // when the file is taken for another encoding; such ranges are written with escapes.
  const { readdir, readFile } = await import('node:fs/promises');
  const { join } = await import('node:path');
  const root = new URL('../src/', import.meta.url).pathname;
  const found: string[] = [];
  const walk = async (dir: string): Promise<void> => {
    for (const entry of await readdir(dir, { withFileTypes: true })) {
      const path = join(dir, entry.name);
      if (entry.isDirectory()) await walk(path);
      else if (/\.ts$/.test(entry.name)) {
        const text = await readFile(path, 'utf8');
        if (/-[\uFFFE\uFFFF]/.test(text)) found.push(path.slice(root.length));
      }
    }
  };
  await walk(root);
  assert.deepEqual(found, []);
});
