// Find and Replace in documents: the paragraphs that are searched and their order (the body,
// its tables, then headers and notes), the way round the end, what a match may not run through,
// what is found but not replaced, deleted text, replacing as one undo step, replacing where
// changes are tracked, and a long document searched in slices.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { paragraphText, type DocModel } from '../src/model/doc/types';
import { DocHistory, DocMatchScan, DocSearch, compileDocFind, docParagraphs, findAllInDocument, findInDocument, insertText, revisionsOf, type DocMatch, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { STYLES, open, p, paragraphAt, pos, r, tbl, tc, texts, tr } from './doc-ops-fixtures';
import { buildDocx } from './docx-fixtures';

const where = (match: DocMatch | undefined): string => (match ? `${match.range.anchor.story}:${match.range.anchor.path.join('.')}@${match.range.anchor.offset}-${match.range.focus.offset}` : 'none');

/** Every match a search reaches from a place, in order, until it comes round to the first again. */
function walk(doc: DocModel, text: string, options: Parameters<typeof findInDocument>[3] = {}, backward = false): string[] {
  const out: string[] = [];
  let from = undefined as DocRange['anchor'] | undefined;
  for (let i = 0; i < 100; i++) {
    const found = findInDocument(doc, text, from, options, backward);
    if (!found) break;
    const label = where(found);
    if (out.indexOf(label) >= 0) break;
    out.push(label);
    from = backward ? found.range.anchor : found.range.focus;
  }
  return out;
}

/** Replaces every match the way the view does: the last first, each through insertText, as one step of the history. */
function replaceAll(doc: DocModel, history: DocHistory, text: string, replacement: string, options = {}): { replaced: number; skipped: number } {
  const matches = findAllInDocument(doc, text, options);
  const usable = matches.filter((match) => match.replaceable);
  let skipped = matches.length - usable.length;
  let replaced = 0;
  history.seal();
  for (let i = usable.length - 1; i >= 0; i--) {
    try {
      const edit = insertText(doc, usable[i].range, replacement);
      history.applyEdit({ ops: edit.ops, selection: edit.selection, mergeKey: 'replace-all' }, undefined);
      replaced++;
    } catch {
      skipped++;
    }
  }
  history.seal();
  return { replaced, skipped };
}

const field = (code: string, result: string) =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r(result)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

test('find in a document: paragraph after paragraph, into tables cell by cell, round the end and backwards', async () => {
  const doc = await open(p('one cat') + tbl(tr(tc('cat a') + tc('b cat')) + tr(tc('c') + tc('cat')), [2000, 2000]) + p('cat cat'));
  assert.deepEqual(walk(doc, 'cat'), ['body:0@4-7', 'body:1.0.0.0@0-3', 'body:1.0.1.0@2-5', 'body:1.1.1.0@0-3', 'body:2@0-3', 'body:2@4-7']);
  assert.deepEqual(walk(doc, 'cat', {}, true), ['body:2@4-7', 'body:2@0-3', 'body:1.1.1.0@0-3', 'body:1.0.1.0@2-5', 'body:1.0.0.0@0-3', 'body:0@4-7']);
  // From the middle: on from the end of the selection, round to the start.
  const from = findInDocument(doc, 'cat', pos([1, 0, 1, 0], 5));
  assert.equal(where(from), 'body:1.1.1.0@0-3');
  const back = findInDocument(doc, 'cat', pos([1, 0, 1, 0], 2), {}, true);
  assert.equal(where(back), 'body:1.0.0.0@0-3');
  // One match only: found again from itself.
  assert.equal(where(findInDocument(doc, 'one', pos(0, 3))), 'body:0@0-3');
  assert.equal(findInDocument(doc, 'dog'), undefined);
  assert.equal(findInDocument(doc, ''), undefined);
  // A match never runs into the next paragraph.
  assert.equal(findInDocument(doc, 'cat cat c'), undefined);
});

test('find in a document: case, Turkish letters and whole words', async () => {
  const doc = await open(p('İstanbul ISTANBUL istanbulda'));
  assert.deepEqual(walk(doc, 'istanbul', { locale: 'tr' }), ['body:0@0-8', 'body:0@18-26']);
  assert.deepEqual(walk(doc, 'istanbul', { locale: 'tr', wholeWord: true }), ['body:0@0-8']);
  assert.deepEqual(walk(doc, 'istanbul'), ['body:0@9-17', 'body:0@18-26']);
  assert.deepEqual(walk(doc, 'ISTANBUL', { matchCase: true }), ['body:0@9-17']);
});

test('find in a document: headers and notes come after the body; only the body when asked', async () => {
  const bytes = await buildDocx({
    body: p('body word' + '') + p(r('note here') + '<w:r><w:footnoteReference w:id="2"/></w:r>') + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>',
    styles: STYLES,
    headers: { rIdH1: p('header word') },
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:t>word in a note</w:t></w:r></w:p></w:footnote>',
  });
  const doc = await readDocx(await OpcPackage.open(bytes));
  const all = walk(doc, 'word');
  assert.equal(all[0], 'body:0@5-9');
  assert.ok(all.some((place) => place.startsWith('header:')), all.join());
  assert.ok(all.some((place) => place.startsWith('footnote:2')), all.join());
  assert.deepEqual(walk(doc, 'word', { stories: 'body' }), ['body:0@5-9']);
  assert.ok(docParagraphs(doc).length > docParagraphs(doc, 'body').length);
});

test('find in a document: a match does not run through a note reference or a field; a field result is found, not replaced', async () => {
  const doc = await open({
    body: p(r('ab') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r('cd')) + p(r('page ') + field('REF mark \\h', 'target') + r(' end')),
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:t>n</w:t></w:r></w:p></w:footnote>',
  });
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'abcd');
  assert.equal(findInDocument(doc, 'bc', undefined, { stories: 'body' }), undefined);
  assert.equal(where(findInDocument(doc, 'ab', undefined, { stories: 'body' })), 'body:0@0-2');
  assert.equal(paragraphText(paragraphAt(doc, 1)), 'page target end');
  const result = findInDocument(doc, 'target', undefined, { stories: 'body' });
  assert.equal(where(result), 'body:1@5-11');
  assert.equal(result!.replaceable, false);
  // Across the edge of the field: not found at all.
  assert.equal(findInDocument(doc, 'page t', undefined, { stories: 'body' }), undefined);
  // Text beside a note reference is replaced, and the reference stays.
  const history = new DocHistory(doc);
  assert.deepEqual(replaceAll(doc, history, 'cd', 'XY', { stories: 'body' }), { replaced: 1, skipped: 0 });
  assert.deepEqual(replaceAll(doc, history, 'ab', '', { stories: 'body' }), { replaced: 1, skipped: 0 });
  const kinds = doc.body.blocks[0].type === 'paragraph' ? doc.body.blocks[0].children.flatMap((inline) => (inline.type === 'run' ? inline.content.map((item) => item.type) : [])) : [];
  assert.ok(kinds.indexOf('footnoteReference') >= 0, kinds.join());
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'XY');
  // The field result is passed over and counted.
  assert.deepEqual(replaceAll(doc, history, 'target', 'x', { stories: 'body' }), { replaced: 0, skipped: 1 });
});

test('find in a document: deleted text is not found; inserted text is', async () => {
  const doc = await open(p(r('keep ') + '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>' + '<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z">' + r('new') + '</w:ins>'));
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'keep new');
  assert.equal(findInDocument(doc, 'gone'), undefined);
  assert.equal(where(findInDocument(doc, 'new')), 'body:0@5-8');
  assert.equal(findInDocument(doc, 'keep new')!.replaceable, true);
});

test('replace all: one undo step that takes everything back; the last match first keeps the places of the others', async () => {
  const doc = await open(p('a cat and a cat') + tbl(tr(tc('cat')), [2000]) + p('Cat!'));
  const history = new DocHistory(doc);
  const before = texts(doc);
  assert.deepEqual(replaceAll(doc, history, 'cat', 'dog'), { replaced: 4, skipped: 0 });
  assert.deepEqual(texts(doc), ['a dog and a dog', 'dog', 'dog!']);
  history.undo();
  assert.deepEqual(texts(doc), before);
  history.redo();
  assert.deepEqual(texts(doc), ['a dog and a dog', 'dog', 'dog!']);
  // Longer and shorter replacements in one paragraph.
  history.undo();
  replaceAll(doc, history, 'a', 'aaa');
  assert.equal(texts(doc)[0], 'aaa caaat aaand aaa caaat');
});

test('replace all: where changes are tracked each replacement is a tracked deletion and insertion', async () => {
  const doc = await open({ body: p('red apple, red pear'), settings: '<w:trackRevisions/>' });
  const history = new DocHistory(doc);
  assert.deepEqual(replaceAll(doc, history, 'red', 'green'), { replaced: 2, skipped: 0 });
  assert.equal(paragraphText(paragraphAt(doc, 0)), 'green apple, green pear');
  const kinds = revisionsOf(doc).map((revision) => revision.kind).sort();
  assert.deepEqual(kinds, ['delete', 'delete', 'insert', 'insert']);
  // The deleted text is no match any more.
  assert.equal(findInDocument(doc, 'red'), undefined);
});

test('find in a document: a document of 500 pages is searched in slices, none of which takes long', async () => {
  // About 40 paragraphs of 300 characters a page.
  const line = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. '.repeat(3);
  let body = '';
  for (let i = 0; i < 20000; i++) body += i === 19999 ? p(line + 'needle') : i % 1000 === 0 ? tbl(tr(tc(line) + tc(line)), [2000, 2000]) : p(line);
  const doc = await open(body);
  const list = docParagraphs(doc);
  assert.ok(list.length >= 20000, String(list.length));
  const find = compileDocFind('needle')!;
  const search = new DocSearch(find, list, undefined);
  let slices = 0;
  let longest = 0;
  let found: DocMatch | null | undefined;
  for (;;) {
    const t = Date.now();
    found = search.step(256);
    longest = Math.max(longest, Date.now() - t);
    slices++;
    if (found !== undefined) break;
  }
  assert.ok(found && found.range.anchor.path[0] === 19999);
  assert.ok(slices > 50, String(slices));
  // A guard against a slice that does everything at once, not a measure of speed.
  assert.ok(longest < 1000, String(longest));
  const scan = new DocMatchScan(compileDocFind('lorem')!, list, 10000);
  let steps = 0;
  while (!scan.step(256)) steps++;
  assert.equal(scan.matches.length, 10000);
  assert.equal(scan.more, true);
  assert.ok(steps > 5);
});
