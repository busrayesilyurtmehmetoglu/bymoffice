import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_HEADER_FOOTER_CHARS, headerFooterIsEmpty, headerFooterLines, parseHeaderFooter, type HeaderFooterRun, type HeaderFooterValues } from '../src/layout/header-footer';

const VALUES: HeaderFooterValues = { page: 3, pages: 12, date: '04.10.2026', time: '14:05', file: 'report.xlsx', path: '', sheet: 'Summary' };

/** The text of a section on a page, lines joined with "|". */
const text = (runs: HeaderFooterRun[], values: HeaderFooterValues = VALUES): string =>
  headerFooterLines(runs, values)
    .map((line) => line.map((piece) => piece.text).join(''))
    .join('|');

const all = (code: string, values: HeaderFooterValues = VALUES): [string, string, string] => {
  const sections = parseHeaderFooter(code);
  return [text(sections.left, values), text(sections.center, values), text(sections.right, values)];
};

test('header and footer: the three sections; text before a section code belongs to the centre', () => {
  assert.deepEqual(all('&LLeft&CCentre&RRight'), ['Left', 'Centre', 'Right']);
  assert.deepEqual(all('Just text'), ['', 'Just text', '']);
  assert.deepEqual(all('&RRight only'), ['', '', 'Right only']);
  // A section named twice goes on where it was; lower case letters are codes too.
  assert.deepEqual(all('&La&cb&La2&rc'), ['aa2', 'b', 'c']);
  assert.equal(headerFooterIsEmpty(parseHeaderFooter('')), true);
  assert.equal(headerFooterIsEmpty(parseHeaderFooter(undefined)), true);
  assert.equal(headerFooterIsEmpty(parseHeaderFooter('&L&C&R')), true);
  assert.equal(headerFooterIsEmpty(parseHeaderFooter('&C ')), false);
});

test('header and footer: fields are replaced by the values of the page', () => {
  assert.deepEqual(all('&LPage &P of &N&C&A&R&D &T'), ['Page 3 of 12', 'Summary', '04.10.2026 14:05']);
  assert.deepEqual(all('&C&F'), ['', 'report.xlsx', '']);
  // An unknown file name and path are empty, not a code.
  assert.deepEqual(all('&C[&F][&Z]', { ...VALUES, file: '' }), ['', '[][]', '']);
  // "&P+n" and "&P-n" add to the page number; a sign that no digit follows is text.
  assert.deepEqual(all('&C&P+2 &P-1 &P+ &P-&N'), ['', '5 2 3+ 3-12', '']);
  // "&&" is an ampersand.
  assert.deepEqual(all('&CTom && Jerry &&P'), ['', 'Tom & Jerry &P', '']);
});

test('header and footer: format codes never become text, and what they say is kept', () => {
  const sections = parseHeaderFooter('&L&"Arial,Bold Italic"&14Big&"-,Regular" plain&C&BB&B n &Ii&I &Uu&U &Ee&E &Ss&S &Xx&X &Yy&Y&R&KFF0000red&K01+000 theme&8small');
  assert.deepEqual(
    sections.left.map((run) => [run.text, run.format]),
    [
      ['Big', { font: 'Arial', bold: true, italic: true, size: 14 }],
      [' plain', { size: 14 }],
    ],
  );
  assert.deepEqual(
    sections.center.map((run) => [run.text, run.format]),
    [
      ['B', { bold: true }],
      [' n ', {}],
      ['i', { italic: true }],
      [' ', {}],
      ['u', { underline: 1 }],
      [' ', {}],
      ['e', { underline: 2 }],
      [' ', {}],
      ['s', { strike: true }],
      [' ', {}],
      ['x', { vertical: 'super' }],
      [' ', {}],
      ['y', { vertical: 'sub' }],
    ],
  );
  assert.deepEqual(
    sections.right.map((run) => [run.text, run.format]),
    [
      ['red', { color: '#FF0000' }],
      [' theme', {}],
      ['small', { size: 8 }],
    ],
  );
  // The format of one section does not reach the next.
  assert.deepEqual(parseHeaderFooter('&L&Ba&Cb').center[0].format, {});
  // A font without a style part keeps bold and italic as they are.
  assert.deepEqual(parseHeaderFooter('&C&B&"Tahoma"x').center[0].format, { bold: true, font: 'Tahoma' });
});

test('header and footer: codes that are not known, broken codes and codes at the end leave no trace in the text', () => {
  assert.deepEqual(all('&CA&GB&OC&HD&QE&?F'), ['', 'ABCDEF', '']);
  // A font name without its closing quote takes the rest; a single "&" at the end is dropped.
  assert.deepEqual(all('&Cbefore&"Arial,Bold after'), ['', 'before', '']);
  assert.deepEqual(all('&Ctext&'), ['', 'text', '']);
  // A size of zero or beyond the largest is ignored, its digits are not text; at most three digits are a size.
  assert.deepEqual(parseHeaderFooter('&C&0a&999b&12c').center.map((run) => [run.text, run.format.size]), [['ab', undefined], ['c', 12]]);
  // All the digits are the code: none of them is printed.
  assert.deepEqual(all('&C&1234'), ['', '', '']);
  // "&K" without a colour is dropped alone.
  assert.deepEqual(all('&C&Kzzzzzz'), ['', 'zzzzzz', '']);
  // Markup stays text: nothing is ever interpreted.
  assert.deepEqual(all('&C<b>&P</b><script>'), ['', '<b>3</b><script>', '']);
});

test('header and footer: lines, and the limits on what is read', () => {
  assert.deepEqual(all('&LOne\nTwo\r\nThree&CA\n\nB\n'), ['One|Two|Three', 'A||B', '']);
  // The value of a field is one piece of its line.
  assert.deepEqual(all('&C&A', { ...VALUES, sheet: 'Two\nlines' }), ['', 'Two lines', '']);
  // Pieces of one format are joined across fields; a new format starts a new piece.
  const lines = headerFooterLines(parseHeaderFooter('&CPage &P&B!&B').center, VALUES);
  assert.deepEqual(lines.map((line) => line.map((piece) => piece.text)), [['Page 3', '!']]);
  // A code longer than the limit is cut; thousands of runs and lines are not all kept.
  const long = parseHeaderFooter('&C' + 'x'.repeat(MAX_HEADER_FOOTER_CHARS * 2));
  assert.equal(text(long.center).length, MAX_HEADER_FOOTER_CHARS - 2);
  const many = parseHeaderFooter('&C' + '&Ba'.repeat(2000));
  assert.equal(many.center.length, 512);
  const tall = headerFooterLines(parseHeaderFooter('&C' + 'a\n'.repeat(2000)).center, VALUES);
  assert.equal(tall.length, 64);
  // Something that is no text is no header.
  assert.equal(headerFooterIsEmpty(parseHeaderFooter(42 as unknown as string)), true);
});

test('header and footer: digits after a size code never become text, however many there are', () => {
  for (const source of ['&C&999999x', '&C&12345x', '&C&1x']) {
    assert.deepEqual(parseHeaderFooter(source).center.map((run) => run.text), ['x'], source);
  }
  // A size beyond the limits is left out, one within them applies.
  assert.equal(parseHeaderFooter('&C&999999x').center[0].format.size, undefined);
  assert.equal(parseHeaderFooter('&C&14x').center[0].format.size, 14);
});
