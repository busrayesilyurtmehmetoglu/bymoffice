// How tracked changes show on the page: what was inserted is underlined, what
// was deleted is struck through and takes no place in the text, both in a
// colour that tells the authors apart.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { revisionColor } from '../src/layout/doc/prep';
import { paragraphText, type Paragraph } from '../src/model/doc/types';
import { deleteBackward, insertText } from '../src/ops/doc/index';
import { applyDocOps } from '../src/ops/doc/ops';
import { DocLayout } from '../src/layout/doc/layout';
import { fakeMeasurer, lay, p, pageText, r, texts } from './doc-layout-fixtures';

const BODY = p(
  r('keep ') +
    '<w:del w:id="1" w:author="Ann"><w:r><w:delText xml:space="preserve">gone </w:delText></w:r></w:del>' +
    `<w:ins w:id="2" w:author="Ann">${r('new ')}</w:ins>` +
    `<w:moveFrom w:id="3" w:author="Bob">${r('from ')}</w:moveFrom>` +
    `<w:moveTo w:id="4" w:author="Bob">${r('to')}</w:moveTo>`,
);

test('deleted and moved-away text is shown struck through, without a place in the text', async () => {
  const { pages, doc } = await lay(BODY);
  const items = texts(pages[0]);
  assert.equal(items.map((item) => item.text).join(''), 'keep gone new from to');
  // The offsets are those of the text that is there.
  assert.equal(pageText(pages[0]), 'keep new to');
  assert.equal(pageText(pages[0]), paragraphText(doc.body.blocks[0] as Paragraph));
  const gone = items.filter((item) => item.text.indexOf('gone') >= 0)[0];
  assert.equal(gone.source, undefined);
  assert.equal(gone.strike, true);
  assert.equal(gone.color, revisionColor('Ann'));
  const from = items.filter((item) => item.text.indexOf('from') >= 0)[0];
  assert.equal(from.source, undefined);
  assert.equal(from.strike, true);
  assert.equal(from.color, revisionColor('Bob'));
  // In the order of the file, each after what is before it.
  for (let i = 1; i < items.length; i++) assert.ok(items[i].x >= items[i - 1].x + items[i - 1].width - 0.01, `${items[i - 1].text}|${items[i].text}`);
});

test('inserted and moved-in text is underlined in the colour of its author', async () => {
  const { pages } = await lay(BODY);
  const items = texts(pages[0]);
  const fresh = items.filter((item) => item.text.indexOf('new') >= 0)[0];
  assert.deepEqual(fresh.underline, { color: revisionColor('Ann'), style: 'single' });
  assert.equal(fresh.color, revisionColor('Ann'));
  assert.deepEqual(fresh.source && fresh.source.offset, 5);
  const to = items.filter((item) => item.text === 'to')[0];
  assert.deepEqual(to.underline, { color: revisionColor('Bob'), style: 'single' });
  const keep = items[0];
  assert.equal(keep.underline, undefined);
  assert.equal(keep.strike, undefined);
  assert.equal(keep.color, '#000000');
});

test('the colour of an author is one of a few, always the same for the same name', () => {
  const seen = new Set<string>();
  for (const name of ['Ann', 'Bob', 'Çağla', '', 'x'.repeat(5000), 'Author']) {
    const color = revisionColor(name);
    assert.ok(/^#[0-9A-F]{6}$/.test(color), color);
    assert.equal(revisionColor(name), color);
    seen.add(color);
  }
  assert.equal(revisionColor(undefined), revisionColor(''));
  assert.ok(seen.size > 1);
});

test('a deleted tab, a deleted field result and a paragraph of deleted text alone are laid out', async () => {
  const field =
    '<w:del w:id="5" w:author="Ann"><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText> DATE </w:delInstrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:delText>today</w:delText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:del>';
  const { pages } = await lay(p('<w:del w:id="1" w:author="Ann"><w:r><w:delText>a</w:delText><w:tab/><w:delText>b</w:delText></w:r></w:del>') + p(r('x') + field + r('y')) + p('after'));
  // The tab itself is an item without text, for the line through it.
  const items = texts(pages[0]).filter((item) => item.text !== '');
  assert.deepEqual(
    items.map((item) => item.text),
    ['a', 'b', 'x', 'today', 'y', 'after'],
  );
  assert.ok(items[1].x > items[0].x + items[0].width, 'the tab takes room');
  assert.ok(!items.some((item) => /DATE/.test(item.text)));
  // The paragraph of deleted text takes a line of its own.
  assert.ok(items[2].y > items[0].y);
});

test('an edit that is tracked shows as it is made', async () => {
  const { doc } = await lay({ body: p('abcd'), settings: '<w:trackRevisions/>' });
  const who = { author: 'Ann', date: '2026-01-01T00:00:00Z' };
  applyDocOps(doc, deleteBackward(doc, { story: 'body', path: [0], offset: 4 }, 'char', who).ops);
  applyDocOps(doc, insertText(doc, { story: 'body', path: [0], offset: 0 }, 'Z', undefined, who).ops);
  const layout = new DocLayout(doc, fakeMeasurer);
  layout.layout();
  const items = texts(layout.pages[0]);
  assert.deepEqual(
    items.map((item) => [item.text, item.strike === true, item.underline !== undefined, item.source !== undefined]),
    [
      ['Z', false, true, true],
      ['abc', false, false, true],
      ['d', true, false, false],
    ],
  );
});

test('the code of a deleted field is never shown, whether its run holds code alone or code beside text', async () => {
  // One run that holds the mark, the code and the result: the flags of the run say which item is code.
  const mixed =
    '<w:del w:id="5" w:author="Ann"><w:r><w:delText>before</w:delText><w:fldChar w:fldCharType="begin"/><w:delInstrText> AUTHOR </w:delInstrText><w:delText>HIDDEN</w:delText><w:fldChar w:fldCharType="separate"/><w:delText>shown</w:delText><w:fldChar w:fldCharType="end"/></w:r></w:del>';
  const { pages } = await lay(p(r('x') + mixed + r('y')));
  const all = texts(pages[0])
    .map((item) => item.text)
    .join('|');
  assert.ok(all.indexOf('before') >= 0 && all.indexOf('shown') >= 0, all);
  assert.equal(all.indexOf('HIDDEN'), -1, all);
  assert.equal(all.indexOf('AUTHOR'), -1, all);
});

test('deleted text says that it is deleted, so that whoever reads the page as text can leave it out', async () => {
  const { pages } = await lay(BODY);
  const items = texts(pages[0]);
  assert.deepEqual(
    items.map((item) => [item.text.trim(), item.deleted === true]),
    [
      ['keep', false],
      ['gone', true],
      ['new', false],
      ['from', true],
      ['to', false],
    ],
  );
});
