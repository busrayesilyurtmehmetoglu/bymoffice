// The tokenizer of the RTF reader: groups, control words with their numbers
// and the space that belongs to them, control symbols, bytes given as hex
// digits, data after "\bin", and the passing over of a group.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { RtfTokenizer, T_BIN, T_CLOSE, T_EOF, T_HEX, T_OPEN, T_SYMBOL, T_TEXT, T_WORD, hexValue } from '../src/formats/rtf/tokenizer';
import { bytesOf } from './rtf-fixtures';

/** The tokens of a source as short texts: "{", "}", "\word", "\word=N", "\'hh", "\x" for a symbol, "text", "bin(...)". */
function tokens(source: string): string[] {
  const tk = new RtfTokenizer(bytesOf(source));
  const out: string[] = [];
  // No source has more tokens than bytes.
  for (let i = 0; i <= source.length; i++) {
    const kind = tk.next();
    if (kind === T_EOF) return out;
    if (kind === T_OPEN) out.push('{');
    else if (kind === T_CLOSE) out.push('}');
    else if (kind === T_WORD) out.push('\\' + tk.word + (tk.hasParam ? '=' + tk.param : ''));
    else if (kind === T_HEX) out.push("\\'" + tk.code.toString(16));
    else if (kind === T_SYMBOL) out.push('sym:' + String.fromCharCode(tk.code));
    else if (kind === T_TEXT) out.push('"' + source.slice(tk.start, tk.stop) + '"');
    else if (kind === T_BIN) out.push('bin(' + source.slice(tk.start, tk.stop) + ')');
  }
  throw new Error('the tokenizer did not reach the end');
}

test('groups, control words, numbers and text', () => {
  assert.deepEqual(tokens(String.raw`{\rtf1\ansi Hello {\b bold}\par}`), ['{', '\\rtf=1', '\\ansi', '"Hello "', '{', '\\b', '"bold"', '}', '\\par', '}']);
});

test('one space after a control word belongs to it; a second one is text', () => {
  assert.deepEqual(tokens(String.raw`\b  x`), ['\\b', '" x"']);
  assert.deepEqual(tokens(String.raw`\fs24 x`), ['\\fs=24', '"x"']);
  // Anything else that ends a word is not swallowed.
  assert.deepEqual(tokens('\\b\r\nx'), ['\\b', '"\r\nx"']);
  assert.deepEqual(tokens(String.raw`\b;x`), ['\\b', '";x"']);
});

test('a number may be negative, and a minus without digits is text', () => {
  assert.deepEqual(tokens(String.raw`\li-720\fi0\b0 x`), ['\\li=-720', '\\fi=0', '\\b=0', '"x"']);
  assert.deepEqual(tokens(String.raw`\b-x`), ['\\b', '"-x"']);
});

test('the letters of a control word may be capitals; a digit ends them', () => {
  assert.deepEqual(tokens(String.raw`\trftsWidth3\clwWidth1200`), ['\\trftsWidth=3', '\\clwWidth=1200']);
});

test('a number that does not fit into 32 bits stays at the largest that does', () => {
  assert.deepEqual(tokens(String.raw`\fs99999999999999999999999\li-99999999999999999999 x`), ['\\fs=2147483647', '\\li=-2147483647', '"x"']);
  assert.deepEqual(tokens(String.raw`\fs2147483647\fs2147483648`), ['\\fs=2147483647', '\\fs=2147483647']);
});

test('a control word is cut after 32 letters', () => {
  const long = 'a'.repeat(40);
  const out = tokens('\\' + long + ' x');
  assert.equal(out[0], '\\' + 'a'.repeat(32));
  assert.equal(out[1], '"aaaaaaaa x"');
});

test('control symbols and bytes as hex digits', () => {
  assert.deepEqual(tokens(String.raw`\{\}\\\~\-\_\*\'fd\'FE`), ['sym:{', 'sym:}', 'sym:\\', 'sym:~', 'sym:-', 'sym:_', 'sym:*', "\\'fd", "\\'fe"]);
  // One digit is taken as the value; none makes the symbol stand for nothing.
  assert.deepEqual(tokens(String.raw`\'4`), ["\\'4"]);
  assert.deepEqual(tokens(String.raw`\'gx`), ["sym:'", '"gx"']);
  // A backslash before a line end is a symbol of its own.
  assert.deepEqual(tokens('a\\\nb'), ['"a"', 'sym:\n', '"b"']);
});

test('a backslash that ends the file stands for nothing', () => {
  assert.deepEqual(tokens('abc\\'), ['"abc"']);
});

test('"\\bin" takes the bytes after it as data, braces and backslashes included', () => {
  assert.deepEqual(tokens(String.raw`{\pict\bin5 {}\{\x}`), ['{', '\\pict', 'bin({}\\{\\)', '"x"', '}']);
  // The space after the number belongs to the word, as for every control word.
  assert.deepEqual(tokens(String.raw`\bin2 ab c`), ['bin(ab)', '" c"']);
});

test('the length of "\\bin" is cut to what the file holds; a negative or missing one is nothing', () => {
  assert.deepEqual(tokens(String.raw`\bin999999 abc`), ['bin(abc)']);
  assert.deepEqual(tokens(String.raw`\bin99999999999999999999 abc`), ['bin(abc)']);
  assert.deepEqual(tokens(String.raw`\bin-5 abc`), ['bin()', '"abc"']);
  assert.deepEqual(tokens(String.raw`\bin abc`), ['bin()', '"abc"']);
});

test('skipGroup passes over the rest of a group, groups inside it included', () => {
  const source = String.raw`{\*\skip a{b{c}d}\{\}e \bin3 }{x}after}`;
  const tk = new RtfTokenizer(bytesOf(source));
  assert.equal(tk.next(), T_OPEN);
  assert.equal(tk.next(), T_SYMBOL);
  tk.skipGroup();
  assert.equal(tk.next(), T_CLOSE);
  // The three bytes after "\\bin3 " were data: the brace among them closed nothing.
  assert.equal(source.slice(tk.pos), 'after}');
});

test('skipGroup ends at the end of a file that is cut short', () => {
  const source = String.raw`{\info{\title abc`;
  const tk = new RtfTokenizer(bytesOf(source));
  tk.next();
  tk.next();
  tk.skipGroup();
  assert.equal(tk.pos, source.length);
  assert.equal(tk.next(), T_EOF);
});

test('skipGroup does not take a word that only starts with "bin" for data', () => {
  const source = String.raw`{\binfsxn1 a}b`;
  const tk = new RtfTokenizer(bytesOf(source));
  tk.next();
  tk.skipGroup();
  assert.equal(tk.next(), T_CLOSE);
  assert.equal(tk.next(), T_TEXT);
  assert.equal(source.slice(tk.start, tk.stop), 'b');
});

test('the name of a control word is one string however often it occurs', () => {
  const tk = new RtfTokenizer(bytesOf(String.raw`\par\pard\par`));
  tk.next();
  const first = tk.word;
  tk.next();
  assert.equal(tk.word, 'pard');
  tk.next();
  assert.equal(tk.word, first);
});

test('hexValue', () => {
  assert.equal(hexValue(0x30), 0);
  assert.equal(hexValue(0x66), 15);
  assert.equal(hexValue(0x46), 15);
  assert.equal(hexValue(0x67), -1);
  assert.equal(hexValue(0x20), -1);
});
