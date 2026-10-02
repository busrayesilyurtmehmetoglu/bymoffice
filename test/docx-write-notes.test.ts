// writeDocx after a note reference was deleted with the editing commands: the
// note leaves the notes part with it, the other notes and their references
// stay as they were, and undoing brings the file back.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type DocModel, type Note, type Paragraph, type StoryCollection } from '../src/model/doc/types';
import { DocHistory, deleteBackward, deleteRange, type DocPosition } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { buildDocx } from './docx-fixtures';

const at = (block: number, offset: number): DocPosition => ({ story: 'body', path: [block], offset });

const SEPARATORS =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
const note = (kind: 'footnote' | 'endnote', id: number) => `<w:${kind} w:id="${id}"><w:p><w:r><w:${kind}Ref/></w:r><w:r><w:t xml:space="preserve"> ${kind} ${id}</w:t></w:r></w:p></w:${kind}>`;
const ref = (kind: 'footnote' | 'endnote', id: number) => `<w:r><w:${kind}Reference w:id="${id}"/></w:r>`;
const run = (text: string) => `<w:r><w:t>${text}</w:t></w:r>`;

/** The references of the body to notes of a kind, in order. */
function references(doc: DocModel, kind: 'footnoteReference' | 'endnoteReference'): number[] {
  const out: number[] = [];
  for (const block of doc.body.blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) if (inline.type === 'run') for (const item of inline.content) if (item.type === kind) out.push(item.id);
  }
  return out;
}

const texts = (collection: StoryCollection<Note> | undefined): [number, string][] =>
  Array.from((collection as StoryCollection<Note>).items.values()).map((item) => [item.id, item.blocks.map((block) => paragraphText(block as Paragraph)).join('|')]);

test('a deleted note reference takes its note out of the file; the other notes and references stay consistent', async () => {
  const bytes = await buildDocx({
    body: `<w:p>${run('one')}${ref('footnote', 1)}${run('two')}${ref('footnote', 2)}${run('three')}${ref('footnote', 3)}</w:p><w:p>${run('end')}${ref('endnote', 1)}</w:p>`,
    footnotes: SEPARATORS + note('footnote', 1) + note('footnote', 2) + note('footnote', 3),
    endnotes: '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>' + note('endnote', 1),
  });
  const pkg = await OpcPackage.open(bytes);
  const original = new Map<string, Uint8Array>();
  for (const part of pkg.partNames()) original.set(part, await pkg.readPart(part));
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  assert.deepEqual(texts(doc.footnotes), [
    [-1, ''],
    [0, ''],
    [1, ' footnote 1'],
    [2, ' footnote 2'],
    [3, ' footnote 3'],
  ]);
  // Backspace after the second reference ("onetwo|"), then a range over the endnote reference.
  history.applyEdit(deleteBackward(doc, at(0, 6)), undefined);
  history.seal();
  history.applyEdit(deleteRange(doc, { anchor: at(1, 1), focus: at(1, 3) }), undefined);
  assert.deepEqual(references(doc, 'footnoteReference'), [1, 3]);
  await writeDocx(pkg, doc);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  assert.deepEqual(texts(again.footnotes), [
    [-1, ''],
    [0, ''],
    [1, ' footnote 1'],
    [3, ' footnote 3'],
  ]);
  assert.deepEqual(references(again, 'footnoteReference'), [1, 3]);
  for (const id of references(again, 'footnoteReference')) assert.ok(again.footnotes!.items.has(id), `reference ${id} has its note`);
  assert.deepEqual(references(again, 'endnoteReference'), []);
  assert.deepEqual(texts(again.endnotes), [[-1, '']]);
  assert.deepEqual(again.warnings, doc.warnings);

  // Undone, the file is what it was.
  while (history.undo());
  await writeDocx(pkg, doc);
  const back = await OpcPackage.open(await pkg.save());
  assert.deepEqual(back.partNames().sort(), Array.from(original.keys()).sort());
  for (const [part, content] of original) assert.deepEqual(await back.readPart(part), content, part);
});
