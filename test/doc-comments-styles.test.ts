// The styles Word gives a comment (its mark in the text, and the text of the comment itself) are
// brought along with the first comment of a document that has none of them.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createBlank } from '../src/formats/blank';
import { open as openFile } from '../src/index';
import type { DocModel, Paragraph, Run } from '../src/model/doc/types';
import { DocHistory, addComment, editComment, insertText, replyToComment, type DocRange } from '../src/ops/doc/index';
import { ZipArchive } from '../src/zip';
import { openCommented } from './doc-comments-fixtures';
import { STYLES, p, pos, range, run } from './doc-ops-fixtures';

const decoder = new TextDecoder();
const partText = async (bytes: Uint8Array, name: string): Promise<string> => decoder.decode(await ZipArchive.read(bytes).read(name));
const count = (text: string, needle: string): number => text.split(needle).length - 1;

/** The style of the run that refers to a comment in the text, and the styles of the first paragraph of the comment and of its mark there. */
function used(doc: DocModel, id: number): { reference?: string; text?: string; mark?: string } {
  let reference: string | undefined;
  for (const block of doc.body.blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) if (inline.type === 'run' && inline.content.some((item) => item.type === 'commentReference' && item.id === id)) reference = inline.props.style;
  }
  const first = doc.comments!.items.get(id)!.blocks[0] as Paragraph;
  const mark = first.children.find((inline): inline is Run => inline.type === 'run' && inline.content.some((item) => item.type === 'annotationRef'));
  return { reference, text: first.props.style, mark: mark && mark.props.style };
}

test('the first comment of a new document brings the two styles of a comment, once', async () => {
  const file = await openFile(await createBlank('docx'));
  const doc = await file.document!();
  const history = new DocHistory<DocRange | undefined>(doc);
  run(history, insertText(doc, pos(0, 0), 'One two three.'));
  assert.equal(doc.styles.styles.has('CommentReference'), false);
  const before = doc.styles;
  const first = addComment(doc, range(pos(0, 4), pos(0, 7)), 'Why two?', { author: 'Ann' });
  run(history, first);
  const reference = doc.styles.styles.get('CommentReference')!;
  const text = doc.styles.styles.get('CommentText')!;
  assert.deepEqual([reference.type, reference.name, reference.basedOn, reference.run], ['character', 'annotation reference', 'DefaultParagraphFont', { size: 16, sizeCs: 16 }]);
  assert.deepEqual([text.type, text.name, text.basedOn, text.run], ['paragraph', 'annotation text', 'Normal', { size: 20, sizeCs: 20 }]);
  assert.deepEqual(used(doc, first.id as number), { reference: 'CommentReference', text: 'CommentText', mark: 'CommentReference' });

  // Another comment and an answer add nothing more.
  const table = doc.styles;
  const second = replyToComment(doc, first.id as number, 'Because.', { author: 'Bob' });
  run(history, second);
  run(history, addComment(doc, range(pos(0, 0), pos(0, 3)), 'And one?', { author: 'Cid' }));
  assert.equal(doc.styles, table, 'the styles are not replaced again');
  assert.deepEqual(used(doc, second.id as number), { reference: 'CommentReference', text: 'CommentText', mark: 'CommentReference' });

  const bytes = await file.save();
  const styles = await partText(bytes, 'word/styles.xml');
  assert.equal(count(styles, 'w:styleId="CommentReference"'), 1);
  assert.equal(count(styles, 'w:styleId="CommentText"'), 1);
  assert.match(styles, /<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"\/><w:basedOn w:val="DefaultParagraphFont"\/><w:uiPriority w:val="99"\/><w:semiHidden\/><w:unhideWhenUsed\/><w:rPr><w:sz w:val="16"\/><w:szCs w:val="16"\/><\/w:rPr><\/w:style>/);
  assert.match(styles, /<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"\/><w:basedOn w:val="Normal"\/><w:uiPriority w:val="99"\/><w:semiHidden\/><w:unhideWhenUsed\/><w:rPr><w:sz w:val="20"\/><w:szCs w:val="20"\/><\/w:rPr><\/w:style>/);
  const body = await partText(bytes, 'word/document.xml');
  assert.equal(count(body, '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference '), 3, 'every reference in the text states the style');
  const comments = await partText(bytes, 'word/comments.xml');
  assert.equal(count(comments, '<w:pStyle w:val="CommentText"/>'), 3);
  assert.equal(count(comments, '<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/>'), 3);
  // Read again: the styles are those of the file, and a comment more adds none.
  const again = await (await openFile(bytes)).document!();
  assert.equal(again.styles.styles.get('CommentText')!.name, 'annotation text');
  const kept = again.styles;
  run(new DocHistory<DocRange | undefined>(again), addComment(again, range(pos(0, 8), pos(0, 13)), 'Three.', { author: 'Dee' }));
  assert.equal(again.styles, kept);

  // Taken back, the document has the styles it had.
  history.undo();
  history.undo();
  history.undo();
  assert.equal(doc.styles, before);
  assert.equal(count(await partText(await file.save(), 'word/styles.xml'), 'CommentReference'), 0);
});

test('styles the document has are used as they are, whatever their ids, and an id that is taken is not used', async () => {
  // Word knows the styles by their names: other ids (a translated Word writes them) are the same styles.
  const named =
    STYLES +
    '<w:style w:type="paragraph" w:styleId="Aciklama"><w:name w:val="annotation text"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:styleId="Basvuru"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="14"/></w:rPr></w:style>';
  let o = await openCommented({ body: p('One two three.'), styles: named });
  let table = o.doc.styles;
  let edit = addComment(o.doc, range(pos(0, 0), pos(0, 3)), 'x');
  run(new DocHistory<DocRange | undefined>(o.doc), edit);
  assert.equal(o.doc.styles, table);
  assert.deepEqual(used(o.doc, edit.id as number), { reference: 'Basvuru', text: 'Aciklama', mark: 'Basvuru' });

  // The ids Word uses name styles of another kind here: they are left alone, and the new styles get ids of their own.
  const taken =
    STYLES +
    '<w:style w:type="paragraph" w:styleId="CommentReference"><w:name w:val="Mine"/></w:style>' +
    '<w:style w:type="character" w:styleId="CommentText"><w:name w:val="Mine too"/></w:style>' +
    '<w:style w:type="character" w:styleId="CommentText1"><w:name w:val="And mine"/></w:style>';
  o = await openCommented({ body: p('One two three.'), styles: taken });
  edit = addComment(o.doc, range(pos(0, 0), pos(0, 3)), 'x');
  run(new DocHistory<DocRange | undefined>(o.doc), edit);
  assert.deepEqual(used(o.doc, edit.id as number), { reference: 'CommentReference1', text: 'CommentText2', mark: 'CommentReference1' });
  assert.equal(o.doc.styles.styles.get('CommentReference')!.name, 'Mine');
  assert.equal(o.doc.styles.styles.get('CommentText')!.name, 'Mine too');
  // No default style to be based on: none is named.
  assert.equal(o.doc.styles.styles.get('CommentReference1')!.basedOn, undefined);
  assert.equal(o.doc.styles.styles.get('CommentText2')!.basedOn, 'Normal');

  // Only one of the two is missing: only that one is made.
  o = await openCommented({ body: p('One two three.'), styles: STYLES + '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="14"/></w:rPr></w:style>' });
  edit = addComment(o.doc, range(pos(0, 0), pos(0, 3)), 'x');
  run(new DocHistory<DocRange | undefined>(o.doc), edit);
  assert.equal(o.doc.styles.styles.get('CommentReference')!.run!.size, 14);
  assert.equal(o.doc.styles.styles.get('CommentText')!.name, 'annotation text');
});

test('changing the text of a comment of a file without those styles brings none', async () => {
  const o = await openCommented({
    body: p('One two three.'),
    styles: STYLES,
    comments: [{ id: 1, content: '<w:p><w:r><w:annotationRef/></w:r><w:r><w:t>plain</w:t></w:r></w:p>' }],
  });
  const table = o.doc.styles;
  run(new DocHistory<DocRange | undefined>(o.doc), editComment(o.doc, 1, 'changed'));
  assert.equal(o.doc.styles, table);
});
