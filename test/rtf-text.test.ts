// Text of an RTF file: paragraphs, the characters control words and symbols
// stand for, code pages, "\u" and what follows it, and which groups are text
// at all.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { decodeByte, highTable, isKnownRtfCodepage } from '../src/formats/rtf/codepage';
import { readRtf } from '../src/formats/rtf/document';
import { bytesOf, doc, itemsOf, paragraphs, read, runsOf, texts } from './rtf-fixtures';

/** "\uN" with what an old reader shows instead. */
const u = (code: number, instead = '?'): string => '\\u' + code + instead;
const chars = (...codes: number[]): string => String.fromCharCode(...codes);

test('a file that does not start with "{\\rtf" is refused', () => {
  assert.throws(() => read('Hello'), /^Error: Invalid rtf: /);
  assert.throws(() => read('{\\rt'), /Invalid rtf/);
  assert.throws(() => read(''), /Invalid rtf/);
  assert.throws(() => read('{\\pict}'), /Invalid rtf/);
  assert.throws(() => readRtf(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0])), /Invalid rtf/);
});

test('a byte order mark and white space before the first group are let pass', () => {
  assert.deepEqual(texts(readRtf(bytesOf(chars(0xef, 0xbb, 0xbf) + ' \r\n{\\rtf1 A\\par}'))), ['A']);
});

test('paragraphs end with "\\par"; text without one after it is a paragraph too', () => {
  assert.deepEqual(texts(read(doc(String.raw`One\par Two\par Three`))), ['One', 'Two', 'Three']);
  assert.deepEqual(texts(read(doc(String.raw`One\par Two\par`))), ['One', 'Two']);
  assert.deepEqual(texts(read(doc(String.raw`\par\par`))), ['', '']);
  // A document without any text has no paragraph.
  assert.deepEqual(texts(read(doc(''))), []);
});

test('line ends of the file are not text; spaces are', () => {
  assert.deepEqual(texts(read(doc('One\r\ntwo  three\n\\par'))), ['Onetwo  three']);
  // A backslash before a line end is a paragraph mark.
  assert.deepEqual(texts(read(doc('One\\\nTwo'))), ['One', 'Two']);
});

test('control symbols and the characters control words stand for', () => {
  const model = read(doc(String.raw`a\{b\}c\\d\~e\emdash\endash\lquote\rquote\ldblquote\rdblquote\bullet\emspace\enspace\qmspace\zwj\zwnj x\par`));
  assert.equal(texts(model)[0], 'a{b}c\\d' + chars(0xa0) + 'e' + chars(0x2014, 0x2013, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2003, 0x2002, 0x2005, 0x200d, 0x200c) + 'x');
});

test('tabs, breaks and hyphens are items of their own', () => {
  const model = read(doc('a\\tab b\tc\\line d\\page e\\column f\\-g\\_h\\lbr2 i\\par'));
  const types = itemsOf(paragraphs(model)[0]).map((item) => (item.type === 'text' ? item.text : item.type === 'break' ? item.kind + (item.clear ? ':' + item.clear : '') : item.type));
  assert.deepEqual(types, ['a', 'tab', 'b', 'tab', 'c', 'line', 'd', 'page', 'e', 'column', 'f', 'softHyphen', 'g', 'noBreakHyphen', 'h', 'line:right', 'i']);
});

test('runs of text with the same formatting are one run', () => {
  const model = read(doc(String.raw`{\insrsid1 Hel}{\insrsid2 lo }{\b bold}{\insrsid3  end}\par`));
  const runs = runsOf(paragraphs(model)[0]);
  assert.equal(runs.length, 3);
  assert.deepEqual(runs[0].content, [{ type: 'text', text: 'Hello ' }]);
});

test('bytes are read in the code page of the document', () => {
  const turkish = String.raw`{\rtf1\ansi\ansicpg1254 \'fd\'fe\'f0\'dd\'de\'d0\par}`;
  assert.deepEqual(texts(read(turkish)), [chars(0x131, 0x15f, 0x11f, 0x130, 0x15e, 0x11e)]);
  // Without "\ansicpg" the bytes are Windows-1252.
  assert.deepEqual(texts(read(String.raw`{\rtf1\ansi \'fd\'fe\'f0\par}`)), [chars(0xfd, 0xfe, 0xf0)]);
  assert.deepEqual(texts(read(String.raw`{\rtf1\ansi\ansicpg1250 \'b9\'e8\par}`)), [chars(0x105, 0x10d)]);
  assert.deepEqual(texts(read(String.raw`{\rtf1\ansi\ansicpg1251 \'c0\'ff\par}`)), [chars(0x410, 0x44f)]);
});

test('"\\mac", "\\pc" and "\\pca" name code pages; a stated "\\ansicpg" wins over them', () => {
  assert.deepEqual(texts(read(String.raw`{\rtf1\mac \'8a\'d2\par}`)), [chars(0xe4, 0x201c)]);
  assert.deepEqual(texts(read(String.raw`{\rtf1\pc \'81\par}`)), [chars(0xfc)]);
  assert.deepEqual(texts(read(String.raw`{\rtf1\pca \'9b\par}`)), [chars(0xf8)]);
  assert.deepEqual(texts(read(String.raw`{\rtf1\ansicpg1254\ansi \'fd\par}`)), [chars(0x131)]);
});

test('the character set of the font in effect says how its bytes are read', () => {
  // f0 is Western (the code page of the document), f1 Turkish, f2 a font of symbols: its bytes above 127 are kept in
  // the private use area, where Word keeps them.
  const model = read(doc(String.raw`\f0 \'fd{\f1 \'fd}\'fd\f1 \'fd\f2 \'b7A\par`));
  assert.equal(texts(model)[0], chars(0xfd, 0x131, 0xfd, 0x131, 0xf0b7) + 'A');
});

test('a font without a character set of its own takes the code page of the document', () => {
  const source = String.raw`{\rtf1\ansi\ansicpg1254{\fonttbl{\f0 Plain;}{\f1\fcharset0 Western;}{\f2\fcharset204 Cyrillic;}{\f3\fcharset1 Default;}}\f0 \'fd\f1 \'fd\f2 \'e4\f3 \'fd\par}`;
  assert.equal(texts(read(source))[0], chars(0x131, 0x131, 0x434, 0x131));
});

test('"\\cpg" of a font wins over its character set', () => {
  const source = String.raw`{\rtf1\ansi{\fonttbl{\f0\fcharset0\cpg1253 Greek;}}\'e1\par}`;
  assert.equal(texts(read(source))[0], chars(0x3b1));
});

test('"\\loch", "\\hich" and "\\dbch" say which associated font the text is in', () => {
  // The high half of the code page is in f1 (Turkish), the low half in f0.
  const model = read(doc(String.raw`\loch\af0\hich\af1\dbch\af0 \loch\f0 a\hich\af1 \'fd\loch\f0 \'fd\par`));
  assert.equal(texts(model)[0], 'a' + chars(0x131, 0xfd));
});

test('"\\u" gives a character by its number and passes over what stands in for it', () => {
  assert.equal(texts(read(doc(u(305) + 'x\\par')))[0], chars(0x131) + 'x');
  // What stands in may be a byte given as hex digits.
  assert.equal(texts(read(doc('\\u305\\' + "'fdx\\par")))[0], chars(0x131) + 'x');
  // A space after the number belongs to the control word, not to what is passed over.
  assert.equal(texts(read(doc('\\u305 ?x\\par')))[0], chars(0x131) + 'x');
});

test('"\\uc" says how much stands in for a "\\u", for the group it is in', () => {
  assert.equal(texts(read(doc('\\uc0 ' + u(305, '') + 'x\\par')))[0], chars(0x131) + 'x');
  assert.equal(texts(read(doc('\\uc2 ' + u(305, 'ab') + 'x\\par')))[0], chars(0x131) + 'x');
  assert.equal(texts(read(doc('{\\uc0 ' + u(305, '') + 'a}' + u(305) + 'b\\par')))[0], chars(0x131) + 'a' + chars(0x131) + 'b');
  // Nothing beyond the group is passed over.
  assert.equal(texts(read(doc('{\\uc3 ' + u(305, 'a') + '}bcd\\par')))[0], chars(0x131) + 'bcd');
});

test('a huge "\\uc" does not swallow the document', () => {
  const model = read(doc('\\uc2147483647 ' + u(305, '') + 'a'.repeat(40) + 'tail\\par'));
  assert.equal(texts(model)[0], chars(0x131) + 'aaaaaaaatail');
});

test('a control word after "\\u" is not passed over', () => {
  assert.deepEqual(texts(read(doc('\\u305\\par x\\par'))), [chars(0x131), 'x']);
  assert.equal(texts(read(doc('\\u305\\u351 ?x\\par')))[0], chars(0x131, 0x15f) + 'x');
});

test('negative numbers of "\\u" and surrogate pairs', () => {
  assert.equal(texts(read(doc(u(-3913) + '\\par')))[0], chars(0xf0b7));
  assert.equal(texts(read(doc(u(-10179) + u(-8704) + '\\par')))[0], chars(0xd83d, 0xde00));
  assert.equal(texts(read(doc(u(55357) + u(56832) + '\\par')))[0], chars(0xd83d, 0xde00));
});

test('two bytes of a double-byte code page are one character', () => {
  const model = read(String.raw`{\rtf1\ansi\ansicpg932 \'82\'a0x\'b1\par}`);
  const text = texts(model)[0];
  // Hiragana "a" where the host can decode it, U+FFFD where it cannot; then a half-width katakana, which is one byte.
  assert.equal(text.length, 3);
  assert.ok(text.charAt(0) === chars(0x3042) || text.charAt(0) === chars(0xfffd));
  assert.equal(text.slice(1), 'x' + chars(0xff71));
  // The second byte may be written as it is.
  const plain = texts(read(String.raw`{\rtf1\ansi\ansicpg932 \'83Ax\par}`))[0];
  assert.equal(plain.length, 2);
  assert.equal(plain.charAt(1), 'x');
});

test('bytes above 127 written as they are take the code page too', () => {
  const bytes = bytesOf('{\\rtf1\\ansi\\ansicpg1254 a' + chars(0xfd) + 'b\\par}');
  assert.equal(texts(readRtf(bytes))[0], 'a' + chars(0x131) + 'b');
});

test('an unknown code page is read as Windows-1252, with a warning', () => {
  const model = read(String.raw`{\rtf1\ansi\ansicpg9999 \'e9\par}`);
  assert.equal(texts(model)[0], chars(0xe9));
  assert.match(model.warnings.join('\n'), /code page 9999/);
});

test('the tables of the code pages agree with the decoder of the host where it has one', () => {
  const labels: [number, string][] = [
    [874, 'windows-874'],
    [1250, 'windows-1250'],
    [1251, 'windows-1251'],
    [1253, 'windows-1253'],
    [1254, 'windows-1254'],
    [1255, 'windows-1255'],
    [1256, 'windows-1256'],
    [1257, 'windows-1257'],
    [1258, 'windows-1258'],
    [10000, 'macintosh'],
  ];
  let checked = 0;
  for (const [codepage, label] of labels) {
    assert.ok(isKnownRtfCodepage(codepage));
    assert.equal(highTable(codepage).length, 128);
    let decoder: TextDecoder;
    try {
      decoder = new TextDecoder(label);
    } catch (error) {
      continue;
    }
    checked++;
    for (let byte = 0x80; byte <= 0xff; byte++) {
      const expected = decoder.decode(new Uint8Array([byte]));
      const got = decodeByte(codepage, byte);
      // Bytes the code page leaves empty: the host gives U+FFFD or the control character of the same number.
      if (expected === chars(0xfffd) || expected.charCodeAt(0) === byte) continue;
      if (got === chars(0xfffd) || got.charCodeAt(0) === byte) continue;
      assert.equal(got, expected, `code page ${codepage}, byte ${byte.toString(16)}`);
    }
  }
  assert.ok(checked === 0 || checked >= 5);
});

test('a control word the reader does not know is passed over, and the text of its group is read', () => {
  const model = read(doc(String.raw`a{\unknownword1 b}\otherunknown c\par`));
  assert.deepEqual(texts(model), ['abc']);
  assert.equal(model.unknown.get('rtf/unknownword'), 1);
  assert.equal(model.unknown.get('rtf/otherunknown'), 1);
});

test('a group marked "\\*" that the reader does not know is passed over whole', () => {
  const model = read(doc(String.raw`a{\*\unknowndest hidden{\b nested} \'fd}b{\*\generator Some Writer 1.0;}c\par`));
  assert.deepEqual(texts(model), ['abc']);
  assert.equal(model.unknown.size, 0);
});

test('destinations the reader has no use for are passed over, with or without the mark', () => {
  const body = String.raw`{\info{\title T}{\author A}}{\*\themedata 504b0304}{\*\colorschememapping 3c3f}{\*\latentstyles\lsdstimax371{\lsdlockedexcept \lsdpriority0 Normal;}}{\*\datastore 0105}{\*\rsidtbl \rsid1}{\*\xmlnstbl {\xmlns1 http://x}}{\*\mmathPr\mmathFont34}{\*\wgrffmtfilter 2450}{\*\pgdsctbl{\pgdsc0 x}}{\xe index}{\tc toc}{\*\pn\pnlvlbody{\pntxta .}}a\par`;
  const model = read(doc(body));
  assert.deepEqual(texts(model), ['a']);
});

test('"\\upr" holds a text twice; the Unicode form is the one that is read', () => {
  const model = read(doc(String.raw`a{\upr{old}{\*\ud{new}}}b\par`));
  assert.deepEqual(texts(model), ['anewb']);
});

test('data after "\\bin" outside of a picture is passed over', () => {
  assert.deepEqual(texts(read(doc(String.raw`a\bin3 {}{b\par`))), ['ab']);
});

test('what follows the group of the document is not read', () => {
  assert.deepEqual(texts(read(String.raw`{\rtf1 a\par}b\par{c}`)), ['a']);
});

test('control characters are not text, however they are written; a tab is', () => {
  const bytes = bytesOf('{\\rtf1 a' + chars(0, 1, 0x0b, 0x1f) + 'b\\' + "'0d\\'0a\\'01c\\u7?d\\u9?e\\'09f\\par}");
  assert.deepEqual(texts(readRtf(bytes)), ['abcd\te\tf']);
});
