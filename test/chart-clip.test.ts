// What the clipboard gets for a chart, and how a paste knows the chart that is held: by the
// mark in the HTML alone, never by what the text says.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { chartClip, chartPasteKind } from '../src/ui/chart-clip';

test('chart clip: the clipboard gets a word and a mark, and nothing else', () => {
  const put = chartClip('Grafik', 'abc123');
  assert.deepEqual(put, { text: 'Grafik', html: '<span data-bym-chart="abc123">Grafik</span>' });
  // A word that would be markup is not written as markup.
  assert.equal(chartClip('<b>&"x"', 'm1').html, '<span data-bym-chart="m1">bx</span>');
});

test('chart clip: a paste is the chart that is held only by the mark of its HTML', () => {
  const put = chartClip('Chart', 'abc123');
  assert.equal(chartPasteKind(put.html, 'abc123'), 'chart');
  // A browser may wrap what it hands out, and may quote otherwise.
  assert.equal(chartPasteKind(`<html><body><!--StartFragment-->${put.html}<!--EndFragment--></body></html>`, 'abc123'), 'chart');
  assert.equal(chartPasteKind("<span data-bym-chart='abc123'>Chart</span>", 'abc123'), 'chart');
  // The mark of another copy (another view, this page before it was loaded again, a copy that is no longer held): no chart, and no text either.
  assert.equal(chartPasteKind(put.html, 'zzz'), 'stale');
  assert.equal(chartPasteKind(put.html, undefined), 'stale');
  assert.equal(chartPasteKind('<span data-bym-chart="abc1234">Chart</span>', 'abc123'), 'stale');
  assert.equal(chartPasteKind('<span data-bym-chart="">Chart</span>', ''), 'stale');
  // The word alone is text, whatever is held: text is never taken for the chart.
  assert.equal(chartPasteKind(undefined, 'abc123'), 'other');
  assert.equal(chartPasteKind('', 'abc123'), 'other');
  assert.equal(chartPasteKind('<span>Chart</span>', 'abc123'), 'other');
  assert.equal(chartPasteKind('<table data-bym-copy="abc123"><tr><td>Chart</td></tr></table>', 'abc123'), 'other');
  // Text that speaks of the mark is text.
  assert.equal(chartPasteKind('<p>data-bym-chart=&quot;abc123&quot;</p>', 'abc123'), 'other');
});
