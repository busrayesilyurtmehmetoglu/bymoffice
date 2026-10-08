import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_PASTE_CELLS, parseTsv, parseTsvLimited, toHtmlTable, toTsv } from '../src/edit/clipboard';

test('clipboard: grids become tab separated text and back', () => {
  const grid = [
    ['a', 'b c', ''],
    ['tab\there', 'line\nbreak', 'say "hi"'],
    ['ğüşiöç', '', '"'],
  ];
  const text = toTsv(grid);
  assert.equal(text, 'a\tb c\t\n"tab\there"\t"line\nbreak"\t"say ""hi"""\nğüşiöç\t\t""""');
  assert.deepEqual(parseTsv(text), grid);
});

test('clipboard: text from other programs', () => {
  assert.deepEqual(parseTsv('1\t2\r\n3\t4\r\n'), [['1', '2'], ['3', '4']]);
  assert.deepEqual(parseTsv('1\t2\n3'), [['1', '2'], ['3']]);
  assert.deepEqual(parseTsv('a\rb'), [['a'], ['b']]);
  assert.deepEqual(parseTsv(''), [['']]);
  assert.deepEqual(parseTsv('\n'), [['']]);
  assert.deepEqual(parseTsv('\t'), [['', '']]);
  assert.deepEqual(parseTsv('a\n\nb'), [['a'], [''], ['b']]);
  // A quote that is not at the start of a field is an ordinary character.
  assert.deepEqual(parseTsv('5" disk\t"x"'), [['5" disk', 'x']]);
  // Unclosed quote: the rest of the text is the field.
  assert.deepEqual(parseTsv('"abc\tdef\nghi'), [['abc\tdef\nghi']]);
  assert.deepEqual(parseTsv('"a"b\tc'), [['ab', 'c']]);
});

test('clipboard: a huge paste is cut off instead of stalling', () => {
  const text = 'x\t'.repeat(MAX_PASTE_CELLS + 5000);
  const started = Date.now();
  const rows = parseTsv(text);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].length, MAX_PASTE_CELLS);
  assert.ok(Date.now() - started < 30000);
  assert.equal(parseTsvLimited(text).truncated, true);
  assert.equal(parseTsvLimited('a\tb\nc').truncated, false);
  assert.equal(parseTsvLimited('x\t'.repeat(MAX_PASTE_CELLS - 1) + 'x').truncated, false);
});

test('clipboard: HTML tables escape everything', () => {
  assert.equal(
    toHtmlTable([['<b>&"x"</b>', 'a\nb'], ['', '<script>alert(1)</script>']]),
    '<table><tr><td>&lt;b&gt;&amp;&quot;x&quot;&lt;/b&gt;</td><td>a<br>b</td></tr><tr><td></td><td>&lt;script&gt;alert(1)&lt;/script&gt;</td></tr></table>',
  );
});
