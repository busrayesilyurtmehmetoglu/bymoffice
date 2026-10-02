// What an RTF file holds besides its main text: fields, headers and footers,
// notes, comments, bookmarks, tracked changes and embedded objects.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import { findSections, inheritedPart } from '../src/layout/doc/page';
import { paragraphText, safeHyperlinkTarget, type FieldCharItem, type Paragraph, type RunItem } from '../src/model/doc/types';
import { fakeMeasurer, texts as pageTexts } from './doc-layout-fixtures';
import { doc, itemsOf, paragraphs, paragraphsOf, read, runWith, runsOf, texts } from './rtf-fixtures';

/** The items of a paragraph as short texts, field code text in brackets. */
function shape(paragraph: Paragraph): string[] {
  return itemsOf(paragraph).map((item: RunItem) => {
    if (item.type === 'text') return item.text;
    if (item.type === 'instrText') return '[' + item.text + ']';
    if (item.type === 'fieldChar') return item.charType;
    return item.type;
  });
}

const begins = (paragraph: Paragraph): FieldCharItem[] => itemsOf(paragraph).filter((item): item is FieldCharItem => item.type === 'fieldChar' && item.charType === 'begin');

test('a field is a complex field: begin, code, separator, result, end', () => {
  const model = read(doc(String.raw`Page {\field{\*\fldinst { PAGE }}{\fldrslt {\b 7}}} of {\field\flddirty{\*\fldinst NUMPAGES}{\fldrslt 9}}.\par`));
  const p = paragraphs(model)[0];
  assert.deepEqual(shape(p), ['Page ', 'begin', '[ PAGE ]', 'separate', '7', 'end', ' of ', 'begin', '[NUMPAGES]', 'separate', '9', 'end', '.']);
  assert.equal(paragraphText(p), 'Page 7 of 9.');
  assert.deepEqual(begins(p).map((item) => item.field), [
    { instruction: ' PAGE ', result: '7' },
    { instruction: 'NUMPAGES', result: '9' },
  ]);
  // Each mark is a run of its own, and the code is flagged as not shown.
  const runs = runsOf(p);
  assert.equal(runs[1].content.length, 1);
  assert.equal(runs[2].fieldCode, true);
  assert.equal(runWith(p, '7').props.bold, true);
});

test('a hyperlink field keeps its target as written', () => {
  const model = read(doc(String.raw`{\field{\*\fldinst{HYPERLINK "https://example.org/a b" \\o "tip"}}{\fldrslt{\ul\cf2 link text}}}\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'link text');
  const field = begins(p)[0].field!;
  assert.equal(field.instruction, 'HYPERLINK "https://example.org/a b" \\o "tip"');
  assert.equal(field.result, 'link text');
  // The reader does not judge targets: a script stays in the model as text, and the view refuses it.
  const script = read(doc(String.raw`{\field{\*\fldinst HYPERLINK "javascript:alert(1)"}{\fldrslt x}}\par`));
  assert.equal(begins(paragraphs(script)[0])[0].field!.instruction, 'HYPERLINK "javascript:alert(1)"');
  assert.equal(safeHyperlinkTarget('javascript:alert(1)'), undefined);
});

test('a field without a result, a field in a field, and what a field holds besides code and result', () => {
  const none = paragraphs(read(doc(String.raw`a{\field{\*\fldinst TOC \\o}}b\par`)))[0];
  assert.deepEqual(shape(none), ['a', 'begin', '[TOC \\o]', 'end', 'b']);
  assert.equal(paragraphText(none), 'ab');
  const nested = paragraphs(read(doc(String.raw`{\field{\*\fldinst IF {\field{\*\fldinst PAGE}{\fldrslt 2}} = 2 "yes"}{\fldrslt yes}}\par`)))[0];
  assert.deepEqual(shape(nested), ['begin', '[IF ]', 'begin', '[PAGE]', 'separate', '2', 'end', '[ = 2 "yes"]', 'separate', 'yes', 'end']);
  assert.equal(paragraphText(nested), 'yes');
  // The code of a field is its own text; that of a field inside it belongs to that field.
  assert.equal(begins(nested)[0].field!.instruction, 'IF  = 2 "yes"');
  assert.equal(begins(nested)[1].field!.instruction, 'PAGE');
  // Data of the field and text that is neither code nor result do not show.
  const data = paragraphs(read(doc(String.raw`{\field stray{\*\fldinst REF x{\*\datafield 0102ff}}{\fldrslt r}}\par`)))[0];
  assert.equal(paragraphText(data), 'r');
  assert.equal(begins(data)[0].field!.instruction, 'REF x');
});

test('a field may span paragraphs', () => {
  const model = read(doc(String.raw`{\field{\*\fldinst TOC}{\fldrslt one\par two\par}}three\par`));
  assert.deepEqual(texts(model), ['one', 'two', 'three']);
  assert.equal(begins(paragraphs(model)[0])[0].field!.result, 'onetwo');
});

test('the page number field shows the number of its page', () => {
  const model = read(doc(String.raw`{\footer\pard\qc Page {\field{\*\fldinst PAGE}{\fldrslt 99}}/{\field{\*\fldinst NUMPAGES}{\fldrslt 99}}\par}one\page two\par`));
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  assert.equal(layout.pages.length, 2);
  const footer = (index: number) =>
    pageTexts(layout.pages[index])
      .filter((item) => item.y > layout.pages[index].height / 2)
      .map((item) => item.text)
      .join('');
  assert.equal(footer(0), 'Page 1/2');
  assert.equal(footer(1), 'Page 2/2');
});

test('headers and footers: one story for each kind, named by its section', () => {
  const body =
    String.raw`\facingp\sectd\titlepg{\header\pard H default\par}{\headerl\pard H even\par}{\headerf\pard H first\par}{\footerr\pard F default\par}{\footerf\pard F first\par}` +
    String.raw`One\par\sect\sectd{\headerr\pard H two\par}Two\par\sect\sectd Three\par`;
  const model = read(doc(body));
  assert.deepEqual(Array.from(model.headers.keys()), ['hdr:0:default', 'hdr:0:even', 'hdr:0:first', 'hdr:1:default']);
  assert.deepEqual(Array.from(model.footers.keys()), ['ftr:0:default', 'ftr:0:first']);
  assert.equal(model.headers.get('hdr:0:even')!.part, 'rtf:header');
  assert.equal(model.footers.get('ftr:0:first')!.part, 'rtf:footer');
  assert.deepEqual(paragraphsOf(model.headers.get('hdr:0:even')!.blocks).map(paragraphText), ['H even']);
  assert.deepEqual(texts(model), ['One', '', 'Two', '', 'Three']);
  const sections = findSections(model);
  assert.deepEqual(sections[0].props.headers, { default: 'hdr:0:default', even: 'hdr:0:even', first: 'hdr:0:first' });
  assert.deepEqual(sections[0].props.footers, { default: 'ftr:0:default', first: 'ftr:0:first' });
  assert.deepEqual(sections[1].props.headers, { default: 'hdr:1:default' });
  assert.equal(sections[1].props.footers, undefined);
  assert.equal(sections[2].props.headers, undefined);
  // A section that names no header has that of the section before it.
  assert.equal(inheritedPart(sections, 2, 'headers', 'default'), 'hdr:1:default');
  assert.equal(inheritedPart(sections, 2, 'footers', 'default'), 'ftr:0:default');
  assert.equal(model.settings.evenAndOddHeaders, true);
});

test('a header does not take the formatting of the text around it, nor leave its own behind', () => {
  const model = read(doc(String.raw`\pard\qc\b\intbl{\header H\par}body\par`));
  const header = paragraphsOf(model.headers.get('hdr:0:default')!.blocks)[0];
  assert.deepEqual(header.props, { widowControl: false });
  assert.deepEqual(runWith(header, 'H').props, {});
  assert.equal(model.headers.get('hdr:0:default')!.blocks[0].type, 'paragraph');
  assert.equal(paragraphsOf(model.body.blocks)[0].props.align, 'center');
});

test('a header with tables, fields and nothing at all', () => {
  const model = read(doc(String.raw`{\header\trowd\cellx3000\pard\intbl A {\field{\*\fldinst PAGE}{\fldrslt 1}}\cell\row}{\footer}body\par`));
  const header = model.headers.get('hdr:0:default')!;
  assert.deepEqual(header.blocks.map((block) => block.type), ['table']);
  assert.equal(paragraphText(paragraphsOf(header.blocks)[0]), 'A 1');
  // An empty footer is a footer: it holds one empty paragraph.
  assert.deepEqual(model.footers.get('ftr:0:default')!.blocks, [{ type: 'paragraph', props: {}, children: [] }]);
});

test('footnotes and endnotes: the reference in the text, the note in its collection', () => {
  const body =
    String.raw`Text{\super\chftn}{\footnote\pard\plain{\super\chftn} First note\par second paragraph}` +
    String.raw` more{\super\chftn}{\footnote\ftnalt\pard\plain{\super\chftn} An endnote}` +
    String.raw` end{\super\chftn}{\footnote\pard\plain{\super\chftn} Third}\par`;
  const model = read(doc(body));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'Text more end');
  const refs = itemsOf(p).filter((item) => item.type === 'footnoteReference' || item.type === 'endnoteReference');
  assert.deepEqual(refs, [
    { type: 'footnoteReference', id: 1 },
    { type: 'endnoteReference', id: 1 },
    { type: 'footnoteReference', id: 2 },
  ]);
  assert.deepEqual(Array.from(model.footnotes!.items.keys()), [1, 2]);
  assert.deepEqual(Array.from(model.endnotes!.items.keys()), [1]);
  const note = model.footnotes!.items.get(1)!;
  assert.deepEqual(paragraphsOf(note.blocks).map(paragraphText), [' First note', 'second paragraph']);
  assert.equal(itemsOf(paragraphsOf(note.blocks)[0])[0].type, 'footnoteRef');
  assert.equal(itemsOf(paragraphsOf(model.endnotes!.items.get(1)!.blocks)[0])[0].type, 'endnoteRef');
  assert.equal(model.footnotes!.part, 'rtf:footnotes');
  // The reference keeps the formatting of the mark.
  assert.equal(runsOf(p)[1].props.vertAlign, 'superscript');
});

test('a note whose mark is text of its own, and a mark without a note', () => {
  const model = read(doc(String.raw`a*{\footnote\pard\plain * custom}b\chftn c\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'a*bc');
  assert.deepEqual(itemsOf(p).filter((item) => item.type === 'footnoteReference'), [{ type: 'footnoteReference', id: 1, customMark: true }]);
  assert.equal(paragraphText(paragraphsOf(model.footnotes!.items.get(1)!.blocks)[0]), '* custom');
});

test('the separators of the notes are notes of a kind of their own, kept when there are notes', () => {
  const separators = String.raw`{\*\ftnsep\pard\plain\chftnsep\par}{\*\ftnsepc\pard\plain\chftnsepc\par}{\*\aftnsep\pard\plain\chftnsep\par}`;
  const withNotes = read(doc(separators + String.raw`a\chftn{\footnote\pard\plain\chftn n}\par`));
  const notes = withNotes.footnotes!.items;
  assert.deepEqual(Array.from(notes.keys()).sort(), [-1, 0, 1]);
  assert.equal(notes.get(-1)!.kind, 'separator');
  assert.equal(itemsOf(paragraphsOf(notes.get(-1)!.blocks)[0])[0].type, 'separator');
  assert.equal(itemsOf(paragraphsOf(notes.get(0)!.blocks)[0])[0].type, 'continuationSeparator');
  assert.equal(withNotes.endnotes, undefined);
  const without = read(doc(separators + String.raw`a\par`));
  assert.equal(without.footnotes, undefined);
  assert.deepEqual(texts(without), ['a']);
});

test('comments: the reference, the author and the text', () => {
  const body =
    String.raw`{\*\atrfstart 7}commented{\*\atrfend 7}{\*\atnid AB}{\*\atnauthor Ada Byron}\chatn{\*\annotation{\*\atnref 7}{\*\atndate 1234}\pard\plain{\chatn} First comment\par}` +
    String.raw` and{\*\atnid CD}{\*\atnauthor Someone Else}\chatn{\*\annotation\pard\plain\chatn Second}\par`;
  const model = read(doc(body));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'commented and');
  assert.deepEqual(itemsOf(p).filter((item) => item.type === 'commentReference'), [
    { type: 'commentReference', id: 0 },
    { type: 'commentReference', id: 1 },
  ]);
  const comments = model.comments!.items;
  assert.equal(comments.get(0)!.author, 'Ada Byron');
  assert.equal(comments.get(0)!.initials, 'AB');
  assert.equal(comments.get(1)!.author, 'Someone Else');
  assert.deepEqual(paragraphsOf(comments.get(0)!.blocks).map(paragraphText), [' First comment']);
  assert.equal(itemsOf(paragraphsOf(comments.get(0)!.blocks)[0])[0].type, 'annotationRef');
  // The marks of the text the comment is about carry its id.
  const marks = p.children.filter((inline) => inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd');
  assert.deepEqual(marks, [
    { type: 'commentRangeStart', id: '0' },
    { type: 'commentRangeEnd', id: '0' },
  ]);
});

test('bookmarks: a start and an end with the same id', () => {
  const model = read(doc(String.raw`a{\*\bkmkstart first}b{\*\bkmkstart second}c{\*\bkmkend first}d{\*\bkmkend second}{\*\bkmkend never}e\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'abcde');
  const marks = p.children.filter((inline) => inline.type !== 'run');
  assert.deepEqual(marks, [
    { type: 'bookmarkStart', id: '0', name: 'first' },
    { type: 'bookmarkStart', id: '1', name: 'second' },
    { type: 'bookmarkEnd', id: '0' },
    { type: 'bookmarkEnd', id: '1' },
  ]);
});

test('tracked changes: deleted text does not show, inserted text does', () => {
  const model = read(doc(String.raw`keep {\deleted gone }{\revised new }end\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'keep new end');
  const gone = runWith(p, 'gone');
  assert.equal(gone.wrap!.kind, 'del');
  assert.deepEqual(gone.content, [{ type: 'deletedText', text: 'gone ' }]);
  assert.equal(runWith(p, 'new').wrap!.kind, 'ins');
  assert.equal(runWith(p, 'keep').wrap, undefined);
});

test('an embedded object shows its result; its data is never read', () => {
  const body = String.raw`a{\object\objemb\objw100\objh100{\*\objclass Package}{\*\objdata 0105000002000000080000005061636b61676500}{\result shown}}b\par`;
  const model = read(doc(body));
  assert.deepEqual(texts(model), ['ashownb']);
  assert.equal(model.unknown.size, 0);
});

test('hidden text is text with a property, as in a docx file', () => {
  const model = read(doc(String.raw`a{\v hidden}b\par`));
  const p = paragraphs(model)[0];
  assert.equal(paragraphText(p), 'ahiddenb');
  assert.equal(runWith(p, 'hidden').props.vanish, true);
});

test('the page number of old files and a page break', () => {
  const model = read(doc(String.raw`{\footer\pard\chpgn\par}a\par`));
  assert.equal(itemsOf(paragraphsOf(model.footers.get('ftr:0:default')!.blocks)[0])[0].type, 'pageNumber');
});
