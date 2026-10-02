// Tests for defects found by the second independent review of the document
// layout engine: each of them failed before its fix.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { mapSymbols, symbolChar, symbolFont } from '../src/layout/doc/symbols';
import { STYLES, lay, p, r, round, texts } from './doc-layout-fixtures';

const mark = (type: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:fldChar w:fldCharType="${type}"/></w:r>`;
const instr = (code: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:instrText xml:space="preserve">${code}</w:instrText></w:r>`;

const BIG = '<w:b/><w:sz w:val="40"/>';
const SMALL = '<w:rFonts w:ascii="Other" w:hAnsi="Other"/><w:sz w:val="16"/>';

test('a live field is shown in the formatting of its result, not of the mark that begins it', async () => {
  // The mark that begins the field is bold and 20 pt; the page number in the file is 8 pt in another font.
  const { pages } = await lay(p(r('Page ') + mark('begin', BIG) + instr(' PAGE ', BIG) + mark('separate', BIG) + r('7', SMALL) + mark('end', BIG)));
  const number = texts(pages[0]).filter((item) => item.text === '1');
  assert.equal(number.length, 1);
  assert.deepEqual([number[0].font.family, round(number[0].font.size), number[0].font.bold], ['Other', 10.67, false]);
  // The text around it keeps its own.
  assert.equal(texts(pages[0])[0].font.size, 16);

  // The first run of the result decides, whatever follows it.
  const two = await lay(p(mark('begin') + instr('PAGE') + mark('separate') + r('7', SMALL) + r('8', BIG) + mark('end')));
  assert.deepEqual([texts(two.pages[0])[0].text, round(texts(two.pages[0])[0].font.size)], ['1', 10.67]);
});

test('a live field without a result takes the formatting of its code, and without that of its first mark', async () => {
  const noResult = await lay(p(mark('begin', BIG) + instr(' PAGE ', SMALL) + mark('end', BIG)));
  assert.deepEqual([texts(noResult.pages[0])[0].text, texts(noResult.pages[0])[0].font.family, round(texts(noResult.pages[0])[0].font.size)], ['1', 'Other', 10.67]);
  // A separator with nothing after it is no result either.
  const empty = await lay(p(mark('begin', BIG) + instr(' PAGE ', SMALL) + mark('separate', BIG) + mark('end', BIG)));
  assert.equal(round(texts(empty.pages[0])[0].font.size), 10.67);
  // A field in the result does not lend its code: the result of that field does.
  const nested = await lay(p(mark('begin', BIG) + instr(' PAGE ', BIG) + mark('separate', BIG) + mark('begin', BIG) + instr(' QUOTE x ', BIG) + mark('separate', BIG) + r('x', SMALL) + mark('end', BIG) + mark('end', BIG)));
  assert.equal(round(texts(nested.pages[0])[0].font.size), 10.67);
});

// ---------------------------------------------------------------------------
// Symbol fonts

const SYM = '<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Symbol"/>';
const pua = (...codes: number[]): string => String.fromCharCode(...codes.map((code) => 0xf000 + code));

test('the Symbol font is mapped in full: Greek letters, operators, arrows and the rest', () => {
  const expected: [number, string][] = [
    [0x22, '∀'], [0x24, '∃'], [0x27, '∋'], [0x2d, '−'], [0x40, '≅'], [0x41, 'Α'], [0x44, 'Δ'], [0x57, 'Ω'], [0x5c, '∴'], [0x5e, '⊥'],
    [0x61, 'α'], [0x62, 'β'], [0x6d, 'μ'], [0x70, 'π'], [0x7a, 'ζ'], [0x7e, '∼'], [0xa0, '€'], [0xa2, '′'], [0xa3, '≤'], [0xa5, '∞'],
    [0xab, '↔'], [0xac, '←'], [0xad, '↑'], [0xae, '→'], [0xaf, '↓'], [0xb0, '°'], [0xb1, '±'], [0xb2, '″'], [0xb3, '≥'], [0xb4, '×'],
    [0xb6, '∂'], [0xb7, '•'], [0xb8, '÷'], [0xb9, '≠'], [0xba, '≡'], [0xbb, '≈'], [0xbc, '…'], [0xc5, '⊕'], [0xc6, '∅'], [0xc7, '∩'],
    [0xc8, '∪'], [0xc9, '⊃'], [0xcc, '⊂'], [0xce, '∈'], [0xcf, '∉'], [0xd0, '∠'], [0xd1, '∇'], [0xd5, '∏'], [0xd6, '√'], [0xd7, '⋅'],
    [0xd8, '¬'], [0xd9, '∧'], [0xda, '∨'], [0xdb, '⇔'], [0xdc, '⇐'], [0xde, '⇒'], [0xe0, '◊'], [0xe5, '∑'], [0xf2, '∫'],
  ];
  for (const [code, char] of expected) {
    assert.equal(symbolChar(1, code), char, code.toString(16));
    // Through the private use area as well.
    assert.equal(symbolChar(1, 0xf000 + code), char, code.toString(16));
  }
  // Digits and punctuation of the font are those of any font.
  assert.equal(mapSymbols(1, pua(0x31, 0x35, 0x3a, 0x30, 0x30)).text, '15:00');
  // Every code of the font has a character, but for the two Unicode has none for and the codes the font leaves empty.
  const missing: number[] = [];
  for (let code = 0x21; code <= 0xfe; code++) if (!(code >= 0x7f && code <= 0x9f) && symbolChar(1, code) === undefined) missing.push(code);
  assert.deepEqual(missing, [0x60, 0xf0]);
  // One character each: text keeps its length.
  for (let code = 0x21; code <= 0xff; code++) for (const kind of [1, 2, 3, 4, 5] as const) assert.ok((symbolChar(kind, code) || 'x').length === 1);
});

test('Wingdings and its relatives: the marks documents use', () => {
  const wingdings: [number, string][] = [
    [0x22, '✂'], [0x28, '☎'], [0x2a, '✉'], [0x45, '☜'], [0x46, '☞'], [0x4a, '☺'], [0x6c, '●'], [0x6e, '■'], [0x6f, '□'], [0x71, '❑'],
    [0x72, '❒'], [0x75, '◆'], [0x76, '❖'], [0x81, '①'], [0x8a, '⑩'], [0x8c, '❶'], [0x9f, '•'], [0xa1, '○'], [0xa7, '▪'], [0xa8, '◻'],
    [0xab, '★'], [0xd8, '➢'], [0xdf, '←'], [0xe0, '→'], [0xe1, '↑'], [0xe2, '↓'], [0xe8, '➔'], [0xf0, '⇨'], [0xfb, '✗'], [0xfc, '✓'],
    [0xfd, '☒'], [0xfe, '☑'],
  ];
  for (const [code, char] of wingdings) assert.equal(symbolChar(2, 0xf000 + code), char, code.toString(16));
  assert.deepEqual([symbolChar(3, 0x50), symbolChar(3, 0x52), symbolChar(3, 0x53), symbolChar(3, 0xa3)], ['✓', '☑', '☒', '□']);
  assert.deepEqual([symbolChar(4, 0x70), symbolChar(4, 0x71), symbolChar(4, 0x74), symbolChar(4, 0x75), symbolChar(4, 0x7d)], ['▲', '▼', '◀', '▶', '▸']);
  assert.deepEqual([symbolChar(5, 0x61), symbolChar(5, 0x72), symbolChar(5, 0x6e)], ['✔', '✖', '●']);
  // The table is chosen by the name of the font, whatever its case.
  assert.deepEqual(['Symbol', 'WINGDINGS', 'Wingdings 2', 'wingdings 3', 'Webdings', 'Arial', undefined].map(symbolFont), [1, 2, 3, 4, 5, 0, 0]);
  // A font the tables do not know: only the bullets Word writes most, and only from the private use area.
  assert.equal(mapSymbols(0, pua(0xb7) + 'o').text, '•o');
});

test('mapped symbols are shown in an ordinary font, glyphs the tables do not know in the font of the document', async () => {
  // A square root, written as Word keeps it.
  const root = await lay(p(r('x ') + r(pua(0xd6), SYM)));
  const sign = texts(root.pages[0]).find((item) => item.text === '√')!;
  assert.ok(sign, 'no square root on the page');
  assert.notEqual(sign.font.family, 'Symbol');
  // The ordinary font is the one the host falls back to.
  const other = await lay(p(r(pua(0xd6), SYM)), { fallbackFamily: 'Fallback' });
  assert.equal(texts(other.pages[0])[0].font.family, 'Fallback');

  // 0x60 has no character: it stays what it is, in Symbol; what is mapped around it is ordinary text.
  const mixed = await lay(p(r(pua(0x61, 0x60, 0x62), SYM)));
  assert.deepEqual(texts(mixed.pages[0]).map((item) => [item.text, item.font.family === 'Symbol']), [
    ['α', false],
    [pua(0x60), true],
    ['β', false],
  ]);
  // The pieces stand one after the other.
  const [a, b, c] = texts(mixed.pages[0]);
  assert.ok(a.x < b.x && b.x < c.x);

  // Nothing mapped: all of it in the font of the document.
  const none = await lay(p(r(pua(0x60), SYM)));
  assert.deepEqual(texts(none.pages[0]).map((item) => item.font.family), ['Symbol']);
});

// ---------------------------------------------------------------------------
// Numbers of lists

test('the letter of a list item is made upper case in the language of the text', async () => {
  const numbering =
    '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="9"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
  const item = (mark = '') => p('x', `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>${mark ? `<w:rPr>${mark}</w:rPr>` : ''}`);
  const labels = async (body: string, settings?: string, styles?: string) => {
    const { pages } = await lay({ body, numbering, settings, styles });
    return texts(pages[0]).filter((text) => !text.source).map((text) => text.text);
  };
  // The language of the paragraph mark, which the number takes its formatting from.
  assert.deepEqual(await labels(item('<w:lang w:val="tr-TR"/>') + item('<w:lang w:val="en-US"/>') + item()), ['İ.', 'J.', 'K.']);
  // Nothing stated: the language of the document.
  assert.deepEqual(await labels(item(), '<w:themeFontLang w:val="tr-TR"/>'), ['İ.']);
  assert.deepEqual(await labels(item(), '<w:themeFontLang w:val="de-DE"/>'), ['I.']);
  assert.deepEqual(await labels(item()), ['I.']);
  // The language the styles give by default counts as stated.
  const turkish = STYLES.replace('<w:sz w:val="24"/>', '<w:sz w:val="24"/><w:lang w:val="tr-TR"/>');
  assert.deepEqual(await labels(item(), '<w:themeFontLang w:val="de-DE"/>', turkish), ['İ.']);
});
