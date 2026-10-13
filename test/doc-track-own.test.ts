// What a tracked deletion takes out for good: only what the same author put in
// since the document was opened. An insertion a file brought is nobody's,
// whatever name it carries: somebody may be waiting to reject it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, deleteBackward, deleteRange, insertText, splitParagraph, type DocEditOptions, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { STYLES, docPicture, p, paragraphAt, pos, r, range, run, texts } from './doc-ops-fixtures';
import { buildDocx as buildSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const ME: DocEditOptions = { author: 'Me', date: '2026-10-05T10:00:00Z' };
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
const ATTRS = ' w:author="Me" w:date="2026-10-05T10:00:00Z"';
const OLD = ' w:author="Me" w:date="2020-01-01T00:00:00Z"';
const t = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;
const dt = (text: string) => `<w:r><w:delText>${text}</w:delText></w:r>`;

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
}

async function open(body: string): Promise<Opened> {
  const pkg = await OpcPackage.open(await buildSpec({ body: body + SECT, styles: STYLES, settings: '<w:trackRevisions/>' }));
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc) };
}

async function body(o: Opened): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  const text = decoder.decode(await o.pkg.readPart('/word/document.xml'));
  return text.slice(text.indexOf('<w:body>') + 8, text.indexOf('<w:sectPr>'));
}

test('text in an insertion the file brought under the name of the editor is deleted as a tracked change, not taken out', async () => {
  const o = await open(p(r('ab') + `<w:ins w:id="7"${OLD}>${r('XY')}</w:ins>`));
  run(o.history, deleteBackward(o.doc, pos(0, 4), 'char', ME));
  // The text stays, deleted, inside the insertion somebody may still reject.
  assert.deepEqual(texts(o.doc), ['abX']);
  assert.equal(await body(o), `<w:p>${t('ab')}<w:ins w:id="7"${OLD}>${t('X')}<w:del w:id="8"${ATTRS}>${dt('Y')}</w:del></w:ins></w:p>`);
});

test('text the editor typed in this session still goes for good, also beside an insertion of the file under the same name', async () => {
  const o = await open(p(r('ab') + `<w:ins w:id="7"${OLD}>${r('XY')}</w:ins>`));
  const before = docPicture(o.doc);
  const at = run(o.history, insertText(o.doc, pos(0, 4), 'Q', undefined, ME));
  assert.deepEqual(texts(o.doc), ['abXYQ']);
  // What was typed is an insertion of its own: it does not continue the one of the file.
  const made = paragraphAt(o.doc, 0).children[2];
  assert.equal(made.wrap && made.wrap.kind === 'ins' ? made.wrap.id : undefined, 8, 'the typed text has an insertion of its own');
  run(o.history, deleteBackward(o.doc, at.focus, 'char', ME));
  assert.deepEqual(texts(o.doc), ['abXY']);
  assert.equal((await body(o)).indexOf('w:del'), -1, 'nothing is left of what was typed and taken back');
  while (o.history.canUndo) o.history.undo();
  assert.equal(docPicture(o.doc), before);
});

test('a paragraph mark the file brought as an insertion of the editor is deleted as a tracked change; one made in this session goes', async () => {
  const o = await open(p('ab', `<w:rPr><w:ins w:id="7"${OLD}/></w:rPr>`) + p('cd'));
  run(o.history, deleteBackward(o.doc, pos(1, 0), 'char', ME));
  assert.deepEqual(texts(o.doc), ['ab', 'cd']);
  const mark = paragraphAt(o.doc, 0).props.markProps;
  assert.equal(mark && mark.deleted && mark.deleted.author, 'Me', 'the mark of the file is deleted, not removed');
  assert.equal(mark && mark.inserted && mark.inserted.id, 7, 'and it still says who inserted it');

  const mine = await open(p('abcd'));
  const at = run(mine.history, splitParagraph(mine.doc, pos(0, 2), ME));
  run(mine.history, deleteBackward(mine.doc, at.focus, 'char', ME));
  assert.deepEqual(texts(mine.doc), ['abcd']);
});

test('a range over paragraphs of the file that carry the name of the editor keeps their marks and text, deleted', async () => {
  const o = await open(p(r('ab') + `<w:ins w:id="5"${OLD}>${r('X')}</w:ins>`, `<w:rPr><w:ins w:id="6"${OLD}/></w:rPr>`) + p(`<w:ins w:id="7"${OLD}>${r('Y')}</w:ins>` + r('cd')));
  run(o.history, deleteRange(o.doc, range(pos(0, 2), pos(1, 1)), ME));
  assert.deepEqual(texts(o.doc), ['ab', 'cd']);
  const xml = await body(o);
  assert.equal((xml.match(/<w:delText/g) || []).length, 2, xml);
  assert.ok(paragraphAt(o.doc, 0).props.markProps?.deleted !== undefined, 'the mark between them is deleted');
});
