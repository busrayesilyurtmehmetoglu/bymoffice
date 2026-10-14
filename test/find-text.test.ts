// Finding text in text: plain text only, case folded by the locale (Turkish "I"/"ı",
// "İ"/"i"), whole words and whole texts, replacing, and what is refused.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_FIND_TEXT, type CompiledFind, type FindQuery, checkReplacement, compileFind, findIn, findLastIn, foldFor, matchesIn, replaceIn } from '../src/edit/find';
import { LOCALE_TR } from '../src/format/locale';

const q = (text: string, extra: Partial<FindQuery> = {}): CompiledFind => {
  const find = compileFind({ text, ...extra });
  if (!find) throw new Error('nothing to look for');
  return find;
};

test('find: an empty text looks for nothing; a text that is not a string or is too long is refused', () => {
  assert.equal(compileFind({ text: '' }), undefined);
  assert.throws(() => compileFind({ text: 5 as unknown as string }), RangeError);
  assert.throws(() => compileFind({ text: 'x'.repeat(MAX_FIND_TEXT + 1) }), RangeError);
  assert.ok(compileFind({ text: 'x'.repeat(MAX_FIND_TEXT) }));
  assert.throws(() => checkReplacement(null), RangeError);
  assert.throws(() => checkReplacement('y'.repeat(MAX_FIND_TEXT + 1)), RangeError);
  assert.equal(checkReplacement(''), '');
});

test('find: the text is plain text - characters of regular expressions and wildcards mean themselves', () => {
  for (const text of ['.', '*', '?', '(a', '[x]', 'a+b', '^$', '\\d', '~*', '$1']) {
    const hay = `before ${text} after`;
    assert.equal(findIn(q(text), hay), 7, text);
    assert.equal(findIn(q(text), 'nothing of it'), -1, text);
  }
  assert.deepEqual(replaceIn(q('a.c'), 'abc a.c', '$&'), { text: 'abc $&', count: 1 });
});

test('find: without regard to case by default, with it when asked', () => {
  assert.equal(findIn(q('apple'), 'An APPLE a day'), 3);
  assert.equal(findIn(q('apple', { matchCase: true }), 'An APPLE a day'), -1);
  assert.equal(findIn(q('APPLE', { matchCase: true }), 'An APPLE a day'), 3);
  assert.equal(findLastIn(q('a'), 'a A a'), 4);
  assert.equal(findLastIn(q('a'), 'a A a', 4), 2);
});

test('find: Turkish folds "I" with "ı" and "İ" with "i"; other languages fold "I" with "i"', () => {
  // Turkish.
  assert.equal(findIn(q('istanbul', { locale: 'tr' }), 'İSTANBUL'), 0);
  assert.equal(findIn(q('ISPARTA', { locale: LOCALE_TR }), 'ısparta'), 0);
  assert.equal(findIn(q('ısparta', { locale: 'tr' }), 'ISPARTA'), 0);
  assert.equal(findIn(q('istanbul', { locale: 'tr' }), 'ISTANBUL'), -1, 'I is the capital of ı in Turkish');
  assert.equal(findIn(q('ığdır', { locale: 'tr' }), 'IĞDIR'), 0);
  // English.
  assert.equal(findIn(q('istanbul'), 'ISTANBUL'), 0);
  assert.equal(findIn(q('istanbul', { locale: 'en' }), 'İSTANBUL'), -1, 'İ is a letter of its own outside Turkish');
  // Folding keeps every character where it was: what is found is found at the same place.
  const text = 'xİyIzı';
  for (const locale of ['tr', 'en'] as const) assert.equal(foldFor(q('a', { locale }), text).length, text.length);
  assert.equal(findIn(q('ı', { locale: 'tr' }), 'aİbIc'), 3);
});

test('find: whole words: no letter or digit right before or after; Turkish letters are letters', () => {
  const word = (text: string) => q(text, { whole: 'word' });
  assert.equal(findIn(word('cat'), 'concatenate cat.'), 12);
  assert.equal(findIn(word('cat'), 'cats'), -1);
  assert.equal(findIn(word('kedi'), 'kediş kedi'), 6);
  assert.equal(findIn(word('kedi'), 'çkedi'), -1);
  assert.equal(findIn(word('5'), '15 5'), 3);
  assert.equal(findIn(word('a b'), 'xa b a b'), 5);
  assert.deepEqual(matchesIn(word('ab'), 'ab ab_ab (ab)'), [0, 10]);
  assert.equal(findLastIn(word('ab'), 'ab abc ab'), 7);
});

test('find: the whole text, as "match entire cell contents"', () => {
  const all = (text: string, extra = {}) => q(text, { whole: 'all', ...extra });
  assert.equal(findIn(all('total'), 'Total'), 0);
  assert.equal(findIn(all('total'), 'Totals'), -1);
  assert.equal(findIn(all('total'), ' total'), -1);
  assert.equal(findIn(all('total', { matchCase: true }), 'Total'), -1);
  assert.deepEqual(replaceIn(all('a'), 'A', 'b'), { text: 'b', count: 1 });
  assert.deepEqual(replaceIn(all('a'), 'aa', 'b'), { text: 'aa', count: 0 });
});

test('find: matches do not overlap; replacing replaces each, the replacement is not searched again', () => {
  assert.deepEqual(matchesIn(q('aa'), 'aaaaa'), [0, 2]);
  assert.deepEqual(replaceIn(q('a'), 'banana', 'aa'), { text: 'baanaanaa', count: 3 });
  assert.deepEqual(replaceIn(q('x'), 'abc', 'y'), { text: 'abc', count: 0 });
  assert.deepEqual(replaceIn(q('İ', { locale: 'tr' }), 'işçi İŞÇİ', '-'), { text: '-şç- -ŞÇ-', count: 4 });
  assert.deepEqual(matchesIn(q('a'), 'aaaa', 2), [0, 1]);
});

test('find: a long text with many near misses is searched in time linear in its length', () => {
  const text = 'ab'.repeat(200000) + 'abc';
  const started = Date.now();
  assert.equal(findIn(q('abc', { whole: 'word' }), text), -1);
  assert.equal(findIn(q('abc'), text), 400000);
  assert.equal(matchesIn(q('ab', { whole: 'word' }), text).length, 0);
  // A guard against a loop, not a measure of speed.
  assert.ok(Date.now() - started < 5000);
});
