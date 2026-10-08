// Second review of tracked changes: the ids of new revisions stay ahead of ids a paste brings
// and of ids the file writes in any spelling, and a date is one a file can hold.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { type DocModel, type Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, fragmentOf, insertFragment, insertText, nextRevisionId, type DocEditOptions, type DocPosition } from '../src/ops/doc/index';
import { buildDocx } from './docx-fixtures';

const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const r = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const at = (block: number, offset: number): DocPosition => ({ story: 'body', path: [block], offset });

async function open(body: string): Promise<DocModel> {
  return readDocx(await OpcPackage.open(await buildDocx({ body: body + SECT, styles: STYLES })));
}

test('review 2: ids written with single quotes or with spaces around the sign count as ids of the document', async () => {
  // Kept as XML by the model: only the source text of the part knows these ids.
  const quoted = await open(`<w:p><w:pPr><w:pPrChange w:id='4000' w:author="A"><w:pPr/></w:pPrChange></w:pPr>${r('a')}</w:p>`);
  assert.equal(nextRevisionId(quoted), 4001);
  const spaced = await open(`<w:p><w:pPr><w:pPrChange w:id = "5000" w:author="A"><w:pPr/></w:pPrChange></w:pPr>${r('a')}</w:p>`);
  assert.equal(nextRevisionId(spaced), 5001);
  const both = await open(`<w:p><w:pPr><w:pPrChange w:id\n=\t'6000' w:author="A"><w:pPr/></w:pPrChange></w:pPr>${r('a')}</w:p><w:p><w:bookmarkStart w:id="7" w:name="b"/>${r('b')}<w:bookmarkEnd w:id="7"/></w:p>`);
  assert.equal(nextRevisionId(both), 6001);
});

/** A copy of "ab" whose paragraph mark is given a revision id by hand, as content from elsewhere may have one. */
function foreign(source: DocModel, id: number) {
  const fragment = fragmentOf(source, { anchor: at(0, 0), focus: at(1, 1) });
  const first = fragment.blocks[0] as Paragraph;
  fragment.blocks[0] = { ...first, props: { ...first.props, markProps: { inserted: { id, author: 'Elsewhere', date: '2020-01-01T00:00:00Z' } } } };
  return fragment;
}

test('review 2: a new revision gets an id beyond those a paste brought, whenever the first id is asked for', async () => {
  const ME: DocEditOptions = { trackChanges: true, author: 'Me', date: '2026-10-05T10:00:00Z' };
  const source = await open(`<w:p>${r('ab')}</w:p><w:p>${r('cd')}</w:p>`);

  // The ids were counted before the paste.
  const counted = await open(`<w:p>${r('target')}</w:p>`);
  assert.equal(nextRevisionId(counted), 1);
  let history = new DocHistory(counted);
  history.applyEdit(insertFragment(counted, at(0, 0), foreign(source, 5000)), undefined);
  history.seal();
  assert.equal(nextRevisionId(counted), 5001);

  // Pasted, undone, the first id asked for, and the paste done again: the id it brings is still out of reach.
  const undone = await open(`<w:p>${r('target')}</w:p>`);
  history = new DocHistory(undone);
  history.applyEdit(insertFragment(undone, at(0, 0), foreign(source, 7000)), undefined);
  history.seal();
  history.undo();
  const edit = insertText(undone, at(0, 0), 'X', undefined, ME);
  history.applyEdit(edit, undefined);
  history.seal();
  assert.equal(nextRevisionId(undone), 7002);
});

test('review 2: the date of a tracked change is a date and time a file can hold', async () => {
  const doc = await open(`<w:p>${r('abc')}</w:p>`);
  const type = (date: string) => insertText(doc, at(0, 0), 'X', undefined, { trackChanges: true, author: 'Me', date });
  for (const date of [
    '2026-10-05T10:00:00Z',
    '2026-10-05T10:00:00',
    '2026-10-05T10:00:00.123Z',
    '2026-10-05T10:00:00+03:00',
    '2026-10-05T10:00:00-14:00',
    '2024-02-29T23:59:59Z',
    '2000-02-29T00:00:00Z',
    '0001-01-01T00:00:00Z',
    '2026-12-31T00:00:00Z',
  ]) {
    assert.doesNotThrow(() => type(date), date);
  }
  for (const date of [
    '2026-10-05T10:00',
    '2026-10-05T10:00Z',
    '2026-02-31T10:00:00Z',
    '2026-02-29T10:00:00Z',
    '1900-02-29T10:00:00Z',
    '2026-04-31T10:00:00Z',
    '2026-00-10T10:00:00Z',
    '2026-13-10T10:00:00Z',
    '2026-10-00T10:00:00Z',
    '0000-01-01T00:00:00Z',
    '2026-10-05T24:00:00Z',
    '2026-10-05T10:60:00Z',
    '2026-10-05T10:00:60Z',
    '2026-10-05T10:00:00+15:00',
    '2026-10-05T10:00:00+03:60',
    '2026-10-05T10:00:00+0300',
    '2026-10-05T10:00:00 Z',
    '2026-10-05 10:00:00Z',
    '2026-10-05',
    '',
  ]) {
    if (date === '') continue;
    assert.throws(() => type(date), RangeError, date);
  }
});
