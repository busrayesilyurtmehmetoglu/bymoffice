// The parts of a Word binary file beside its main text: sections, headers
// and footers, notes, comments, lists, fields, bookmarks and the document
// properties.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { NumberingCounters } from '../src/formats/docx/numbering';
import { resolveParaProps } from '../src/formats/docx/resolve';
import { DocLayout } from '../src/layout/doc/layout';
import { findWrapper, inFieldCode, paragraphText, type DocModel, type Paragraph, type RunItem } from '../src/model/doc/types';
import { fakeMeasurer, texts } from './doc-layout-fixtures';
import { DocBuilder, FC, cat, fontTable, paragraphs, plc, runsOf, sprm, stsh, sttb, textOf, u16le, u32le, u8s, utf16le } from './doc-binary-fixtures';

/** A SEPX in the WordDocument stream; returns the PlcfSed entry that points to it. */
function sed(builder: DocBuilder, ...sprms: Uint8Array[]): Uint8Array {
  const list = cat(...sprms);
  const at = builder.blob(cat(u16le(list.length), list));
  return cat(u16le(0), u32le(at), u16le(0), u32le(0xffffffff));
}

const PAGE_400 = [sprm(0xb01f, u16le(6000)), sprm(0xb020, u16le(6000)), sprm(0xb021, u16le(600)), sprm(0xb022, u16le(600)), sprm(0x9023, u16le(600)), sprm(0x9024, u16le(600)), sprm(0xb017, u16le(150)), sprm(0xb018, u16le(150))];

const items = (paragraph: Paragraph): RunItem[] => runsOf(paragraph).flatMap((run) => run.content);

test('a document without section data gets the page Word assumes', () => {
  const builder = new DocBuilder();
  builder.text('x\r');
  builder.papx(0, 2, 0);
  const section = builder.read().body.section!;
  assert.deepEqual(section.page, { width: 12240, height: 15840 });
  assert.deepEqual(section.margins, { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 });
});

test('sections: the paragraph that ends one carries its properties, the last one is that of the body', () => {
  const builder = new DocBuilder();
  const text = 'first\r\u000csecond\r';
  builder.text(text);
  builder.papx(0, 6, 0).papx(6, 7, 0).papx(7, 14, 0);
  builder.table(
    FC.PlcfSed,
    plc(
      [0, 7, 14],
      [
        sed(builder, ...PAGE_400, sprm(0x300a, 1), sprm(0x500b, u16le(1)), sprm(0x900c, u16le(300)), sprm(0x3019, 1)),
        sed(
          builder,
          sprm(0x3009, 0), // continuous
          sprm(0x301d, 2), // landscape
          sprm(0xb01f, u16le(16838)),
          sprm(0xb020, u16le(11906)),
          sprm(0x300e, 2), // lower roman page numbers
          sprm(0x3011, 1),
          sprm(0x501c, u16le(5)),
          sprm(0x301a, 1),
          sprm(0xb025, u16le(200)),
          sprm(0x5015, u16le(5)),
          sprm(0x3013, 1),
          sprm(0x5040, u16le(4)),
          sprm(0x503f, u16le(3)),
          sprm(0x303c, 2),
          sprm(0x9031, u16le(360)),
          sprm(0x5032, u16le(2)),
          sprm(0x702b, u32le(4 | (1 << 8))),
          sprm(0x522f, u16le(1 | (1 << 5))),
        ),
      ],
    ),
  );
  const model = builder.read();
  const list = paragraphs(model.body.blocks);
  assert.deepEqual(textOf(model.body.blocks), ['first', '', 'second']);
  assert.equal(list[0].props.section, undefined);
  const first = list[1].props.section!;
  assert.deepEqual(first.page, { width: 6000, height: 6000 });
  assert.deepEqual(first.margins, { top: 600, right: 600, bottom: 600, left: 600, header: 150, footer: 150, gutter: 0 });
  assert.equal(first.titlePage, true);
  assert.deepEqual(first.columns, { count: 2, space: 300, separator: true });
  assert.equal(first.type, undefined);
  const last = model.body.section!;
  assert.equal(list[2].props.section, undefined);
  assert.equal(last.type, 'continuous');
  assert.deepEqual(last.page, { width: 16838, height: 11906, orientation: 'landscape' });
  assert.deepEqual(last.pageNumbering, { format: 'lowerRoman', start: 5 });
  assert.equal(last.vAlign, 'center');
  assert.equal(last.margins!.gutter, 200);
  assert.deepEqual(last.lineNumbering, { countBy: 5, restart: 'newSection' });
  assert.deepEqual(last.footnotes, { format: 'lowerLetter', start: 3, restart: 'eachPage' });
  assert.deepEqual(last.docGrid, { linePitch: 360, type: 'lines' });
  assert.deepEqual(last.pageBorders, { top: { style: 'single', size: 4, space: 0, color: { auto: true } }, display: 'firstPage', zOrder: 'front', offsetFrom: 'page' });
});

test('a section that ends in the middle of a paragraph ends with that paragraph; sections without text are passed over', () => {
  const builder = new DocBuilder();
  const text = 'first\rsecond\rthird\r';
  builder.text(text);
  builder.papxAll(text);
  // The first section ends inside "first", the second and the third both inside "second".
  builder.table(
    FC.PlcfSed,
    plc([0, 3, 8, 9, 19], [sed(builder, sprm(0xb01f, u16le(1000))), sed(builder, sprm(0xb01f, u16le(2000))), sed(builder, sprm(0xb01f, u16le(3000))), sed(builder, sprm(0xb01f, u16le(4000)))]),
  );
  const model = builder.read();
  const widths = paragraphs(model.body.blocks).map((paragraph) => paragraph.props.section && paragraph.props.section.page!.width);
  assert.deepEqual(widths, [1000, 3000, undefined]);
  assert.equal(model.body.section!.page!.width, 4000);
  assert.deepEqual(textOf(model.body.blocks), ['first', 'second', 'third']);
});

test('a page break inside a section is a break, not the end of a section', () => {
  const builder = new DocBuilder();
  const text = 'a\u000cb\r';
  builder.text(text);
  builder.papx(0, 4, 0);
  builder.table(FC.PlcfSed, plc([0, 4], [sed(builder, ...PAGE_400)]));
  const model = builder.read();
  assert.equal(paragraphs(model.body.blocks).length, 1);
  assert.equal(paragraphText(paragraphs(model.body.blocks)[0]), 'a\nb');
});

test('columns of unequal width', () => {
  const builder = new DocBuilder();
  builder.text('x\r');
  builder.papx(0, 2, 0);
  builder.table(
    FC.PlcfSed,
    plc([0, 2], [sed(builder, sprm(0x500b, u16le(1)), sprm(0x3005, 0), sprm(0xf203, 0, u16le(3000)), sprm(0xf204, 0, u16le(400)), sprm(0xf203, 1, u16le(5000)))]),
  );
  assert.deepEqual(builder.read().body.section!.columns, { count: 2, space: 720, equalWidth: false, columns: [{ width: 3000, space: 400 }, { width: 5000 }] });
});

/**
 * A document with headers: `stories` are the texts of the header part in
 * file order (six for the note separators, then six for each section: even
 * header, odd header, even footer, odd footer, first header, first footer).
 * A story is given without the mark that ends it; "" is an empty story.
 */
function withHeaders(body: string, sections: number[], stories: string[], options: { titlePage?: boolean; facing?: boolean } = {}): DocBuilder {
  const builder = new DocBuilder();
  let text = body;
  const cps = [0];
  for (const content of stories) {
    if (content !== '') text += content + '\r';
    cps.push(text.length - body.length);
  }
  const header = text.length - body.length;
  // The end of the header part, and the mark that ends the text of the document.
  cps.push(header + 1);
  text += '\r\r';
  builder.text(text);
  builder.papxAll(text);
  builder.ccp.hdd = header + 1;
  builder.table(FC.PlcfHdd, plc(cps));
  const ends = [0, ...sections];
  builder.table(
    FC.PlcfSed,
    plc(
      ends,
      sections.map(() => sed(builder, ...PAGE_400, ...(options.titlePage ? [sprm(0x300a, 1)] : []))),
    ),
  );
  if (options.facing) {
    const dop = new Uint8Array(84);
    dop[0] = 1;
    builder.table(FC.Dop, dop);
  }
  return builder;
}

test('headers and footers: six stories for each section, ids made up from section and kind', () => {
  const none = ['', '', '', '', '', ''];
  const model = withHeaders('body\r', [5], [...none, 'even head', 'odd head', '', 'odd foot\rsecond line', 'first head', '']).read();
  assert.deepEqual([...model.headers.keys()], ['hdr:0:even', 'hdr:0:default', 'hdr:0:first']);
  assert.deepEqual([...model.footers.keys()], ['ftr:0:default']);
  assert.deepEqual(textOf(model.headers.get('hdr:0:default')!.blocks), ['odd head']);
  assert.deepEqual(textOf(model.footers.get('ftr:0:default')!.blocks), ['odd foot', 'second line']);
  assert.deepEqual(model.body.section!.headers, { even: 'hdr:0:even', default: 'hdr:0:default', first: 'hdr:0:first' });
  assert.deepEqual(model.body.section!.footers, { default: 'ftr:0:default' });
  assert.deepEqual(textOf(model.body.blocks), ['body']);
  assert.equal(model.headers.get('hdr:0:default')!.part, 'doc:header');
});

test('an empty header story means "as in the section before"; the layout engine follows that', () => {
  const none = ['', '', '', '', '', ''];
  const body = 'one\r\u000ctwo\r\u000cthree\r';
  const model = withHeaders(body, [5, 10, 16], [...none, '', 'HEAD A', '', 'FOOT A', '', '', ...none, '', 'HEAD C', '', '', '', '']).read();
  const [first, second] = paragraphs(model.body.blocks)
    .map((paragraph) => paragraph.props.section)
    .filter(Boolean);
  assert.deepEqual(first!.headers, { default: 'hdr:0:default' });
  assert.equal(second!.headers, undefined);
  assert.deepEqual(model.body.section!.headers, { default: 'hdr:2:default' });
  assert.equal(model.body.section!.footers, undefined);

  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  const shown = layout.pages.map((page) => texts(page).map((item) => item.text));
  assert.deepEqual(shown, [
    ['HEAD A', 'one', 'FOOT A'],
    ['HEAD A', 'two', 'FOOT A'],
    ['HEAD C', 'three', 'FOOT A'],
  ]);
});

test('first page and even page headers are used when the section and the document ask for them', () => {
  const none = ['', '', '', '', '', ''];
  const body = 'p1\u000cp2\u000cp3\r';
  const stories = [...none, 'EVEN', 'ODD', '', '', 'FIRST', ''];
  const plain = new DocLayout(withHeaders(body, [9], stories).read(), fakeMeasurer);
  plain.layout();
  assert.deepEqual(plain.pages.map((page) => texts(page)[0].text), ['ODD', 'ODD', 'ODD']);
  const model = withHeaders(body, [9], stories, { titlePage: true, facing: true }).read();
  assert.equal(model.settings.evenAndOddHeaders, true);
  const layout = new DocLayout(model, fakeMeasurer);
  layout.layout();
  assert.deepEqual(layout.pages.map((page) => texts(page)[0].text), ['FIRST', 'EVEN', 'ODD']);
});

test('fields: marks around code and result become a complex field, as in docx', () => {
  const builder = new DocBuilder();
  const text = 'Page \u0013 PAGE \\* roman \u00147\u0015 of \u0013NUMPAGES\u001412\u0015.\r';
  builder.text(text);
  builder.papxAll(text);
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) >= 0x13 && text.charCodeAt(i) <= 0x15) builder.chpx(i, i + 1, sprm(0x0855, 1));
  const [paragraph] = paragraphs(builder.read().body.blocks);
  assert.equal(paragraphText(paragraph), 'Page 7 of 12.');
  const begins = items(paragraph).filter((item) => item.type === 'fieldChar' && item.charType === 'begin');
  assert.deepEqual(
    begins.map((item) => (item.type === 'fieldChar' ? item.field : undefined)),
    [
      { instruction: ' PAGE \\* roman ', result: '7' },
      { instruction: 'NUMPAGES', result: '12' },
    ],
  );
  // The code is instruction text and is flagged as not shown.
  const code = runsOf(paragraph).find((run) => run.content.some((item) => item.type === 'instrText'))!;
  const index = code.content.findIndex((item) => item.type === 'instrText');
  assert.equal(inFieldCode(code, index), true);
});

test('the layout engine computes the page fields of a doc file and follows its hyperlinks', () => {
  const none = ['', '', '', '', '', ''];
  const footer = 'Page \u0013 PAGE \u00149\u0015 / \u0013 NUMPAGES \u00149\u0015';
  const body = 'See \u0013 HYPERLINK "https://example.com/a" \u0014the site\u0015.\u000cnext\r';
  const model = withHeaders(body, [body.length], [...none, '', '', '', footer, '', '']).read();
  const layout = new DocLayout(model, fakeMeasurer);
  layout.layout();
  assert.equal(layout.pages.length, 2);
  const footers = layout.pages.map((page) =>
    texts(page)
      .filter((item) => item.y > 300)
      .map((item) => item.text)
      .join(''),
  );
  assert.deepEqual(footers, ['Page 1 / 2', 'Page 2 / 2']);
  const link = texts(layout.pages[0]).filter((item) => item.link !== undefined);
  assert.deepEqual(link.map((item) => [item.text, item.link]), [
    ['the site', 'https://example.com/a'],
  ]);
});

test('nested fields and fields that span paragraphs', () => {
  const text = 'a\u0013 IF \u0013 PAGE \u00143\u0015 = 3 "x" \u0014res\rult\u0015b\r';
  const builder = new DocBuilder();
  builder.text(text);
  builder.papxAll(text);
  const list = paragraphs(builder.read().body.blocks);
  assert.deepEqual(list.map(paragraphText), ['ares', 'ultb']);
  const outer = items(list[0]).find((item) => item.type === 'fieldChar' && item.charType === 'begin');
  assert.deepEqual(outer && outer.type === 'fieldChar' && outer.field, { instruction: ' IF  = 3 "x" ', result: 'result' });
});

test('field marks that do not pair up are taken out; the text after them still shows', () => {
  for (const [text, shown] of [
    ['a\u0013 never closed\rb\r', ['a never closed', 'b']],
    ['a\u0014b\u0015c\r', ['abc']],
    ['a\u0013x\u0014y\u0014z\u0015c\u0015\r', ['ayzc']],
  ] as [string, string[]][]) {
    const builder = new DocBuilder();
    builder.text(text);
    builder.papxAll(text);
    assert.deepEqual(textOf(builder.read().body.blocks), shown);
  }
});

test('bookmarks become range marks at their positions', () => {
  const builder = new DocBuilder();
  const text = 'one two three\rnext\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(0, 7, sprm(0x0835, 1));
  builder.table(FC.SttbfBkmk, sttb(['word', 'whole', 'broken']));
  // "word" covers "two", "whole" both paragraphs; the third refers to an end that is not there.
  builder.table(FC.PlcfBkf, plc([4, 0, 2, 19], [cat(u16le(1), u16le(0)), cat(u16le(0), u16le(0)), cat(u16le(9), u16le(0))]));
  builder.table(FC.PlcfBkl, plc([19, 7, 19]));
  const [first, second] = paragraphs(builder.read().body.blocks);
  const marks = (paragraph: Paragraph) =>
    paragraph.children.map((inline) => (inline.type === 'run' ? paragraphText({ ...paragraph, children: [inline] }) : inline.type === 'bookmarkStart' ? `<${inline.name}` : inline.type === 'bookmarkEnd' ? `>${inline.id}` : '?'));
  assert.deepEqual(marks(first), ['<whole', 'one ', '<word', 'two', '>0', ' three']);
  assert.deepEqual(marks(second), ['next', '>1']);
  assert.equal(paragraphText(first), 'one two three');
});

test('a bookmark at the border of two stories is in the model once', () => {
  const builder = withNotes('Text\u0002\u0002\r', ['one', 'two'], 'footnote');
  builder.table(FC.SttbfBkmk, sttb(['edge']));
  // The bookmark starts and ends where the first note ends and the second begins.
  const border = 7 + 6;
  builder.table(FC.PlcfBkf, plc([border, 100], [cat(u16le(0), u16le(0))]));
  builder.table(FC.PlcfBkl, plc([border, 100]));
  const model = builder.read();
  let marks = 0;
  for (const note of model.footnotes!.items.values()) {
    for (const paragraph of paragraphs(note.blocks)) marks += paragraph.children.filter((inline) => inline.type === 'bookmarkStart' || inline.type === 'bookmarkEnd').length;
  }
  assert.equal(marks, 2);
});

/** A document with notes: the references are 0x02 characters of the main text; `notes` are the note texts. */
function withNotes(body: string, notes: string[], kind: 'footnote' | 'endnote'): DocBuilder {
  const builder = new DocBuilder();
  let text = body;
  const refs: number[] = [];
  for (let i = 0; i < body.length; i++) if (body.charCodeAt(i) === 2) refs.push(i);
  const cps = [0];
  for (const note of notes) {
    text += '\u0002 ' + note + '\r';
    cps.push(text.length - body.length);
  }
  // The part ends with a paragraph mark of its own, and so does the document.
  cps.push(text.length - body.length + 1);
  const length = text.length - body.length + 1;
  text += '\r\r';
  builder.text(text);
  builder.papxAll(text);
  for (let i = 0; i < text.length; i++) if (text.charCodeAt(i) === 2) builder.chpx(i, i + 1, sprm(0x0855, 1));
  if (kind === 'footnote') builder.ccp.ftn = length;
  else builder.ccp.edn = length;
  builder.table(kind === 'footnote' ? FC.PlcffndRef : FC.PlcfendRef, plc([...refs, body.length], refs.map(() => u16le(1))));
  builder.table(kind === 'footnote' ? FC.PlcffndTxt : FC.PlcfendTxt, plc(cps));
  return builder;
}

test('footnotes: references in the text, the notes as stories with their own mark', () => {
  const model = withNotes('First\u0002 and second\u0002.\r', ['note one', 'note two\rmore'], 'footnote').read();
  const [paragraph] = paragraphs(model.body.blocks);
  assert.equal(paragraphText(paragraph), 'First and second.');
  assert.deepEqual(
    items(paragraph).filter((item) => item.type === 'footnoteReference'),
    [
      { type: 'footnoteReference', id: 1 },
      { type: 'footnoteReference', id: 2 },
    ],
  );
  const notes = model.footnotes!;
  assert.deepEqual([...notes.items.keys()], [1, 2]);
  assert.deepEqual(textOf(notes.items.get(1)!.blocks), [' note one']);
  assert.deepEqual(textOf(notes.items.get(2)!.blocks), [' note two', 'more']);
  assert.equal(items(paragraphs(notes.items.get(1)!.blocks)[0])[0].type, 'footnoteRef');
  assert.equal(model.endnotes, undefined);
  assert.deepEqual(textOf(model.body.blocks), ['First and second.']);

  // The layout engine numbers the references and puts the notes at the foot of the page.
  const layout = new DocLayout(model, fakeMeasurer);
  layout.layout();
  const shown = texts(layout.pages[0]).map((item) => item.text);
  assert.deepEqual(shown, ['First', '1', ' and second', '2', '.', '1', ' note one', '2', ' note two', 'more']);
});

test('endnotes are read the same way', () => {
  const model = withNotes('Text\u0002\r', ['the end'], 'endnote').read();
  assert.deepEqual(items(paragraphs(model.body.blocks)[0]).filter((item) => item.type === 'endnoteReference'), [{ type: 'endnoteReference', id: 1 }]);
  assert.deepEqual(textOf(model.endnotes!.items.get(1)!.blocks), [' the end']);
  assert.equal(items(paragraphs(model.endnotes!.items.get(1)!.blocks)[0])[0].type, 'endnoteRef');
  assert.equal(model.footnotes, undefined);
});

test('a note with a mark of its own: the reference stands before the mark text', () => {
  const builder = new DocBuilder();
  const body = 'See*\r';
  const text = body + 'the note\r\r\r';
  builder.text(text);
  builder.papxAll(text);
  builder.ccp.ftn = 10;
  builder.table(FC.PlcffndRef, plc([3, 5], [u16le(0)]));
  builder.table(FC.PlcffndTxt, plc([0, 9, 10]));
  const model = builder.read();
  const list = items(paragraphs(model.body.blocks)[0]);
  assert.deepEqual(list, [
    { type: 'text', text: 'See' },
    { type: 'footnoteReference', id: 1, customMark: true },
    { type: 'text', text: '*' },
  ]);
  assert.deepEqual(textOf(model.footnotes!.items.get(1)!.blocks), ['the note']);
});

test('comments: the reference in the text, author and initials, the comment as a story', () => {
  const builder = new DocBuilder();
  const body = 'Commented\u0005 text\r';
  const text = body + '\u0005A remark\r\r\r';
  builder.text(text);
  builder.papxAll(text);
  builder.chpx(9, 10, sprm(0x0855, 1)).chpx(body.length, body.length + 1, sprm(0x0855, 1));
  builder.ccp.atn = 11;
  const ref = new Uint8Array(30);
  ref.set(cat(u16le(2), utf16le('ZK')));
  ref.set(u16le(1), 20);
  builder.table(FC.PlcfandRef, plc([9, 16], [ref]));
  builder.table(FC.PlcfandTxt, plc([0, 10, 11]));
  builder.table(FC.GrpXstAtnOwners, cat(u16le(5), utf16le('Other'), u16le(11), utf16le('Zeynep Kaya')));
  const model = builder.read();
  assert.deepEqual(items(paragraphs(model.body.blocks)[0]).filter((item) => item.type === 'commentReference'), [{ type: 'commentReference', id: 0 }]);
  const comment = model.comments!.items.get(0)!;
  assert.equal(comment.author, 'Zeynep Kaya');
  assert.equal(comment.initials, 'ZK');
  assert.deepEqual(textOf(comment.blocks), ['A remark']);
  assert.equal(items(paragraphs(comment.blocks)[0])[0].type, 'annotationRef');
});

/** One level of a list: start, number format, flags, the characters that follow the number, properties, text. */
function lvl(start: number, nfc: number, jc: number, follow: number, text: string, papx: Uint8Array = new Uint8Array(0), chpx: Uint8Array = new Uint8Array(0), flags = 0, restart = 0): Uint8Array {
  return cat(u32le(start), u8s(nfc, jc | flags), new Uint8Array(9), u8s(follow), u32le(0), u32le(0), u8s(chpx.length, papx.length, restart, 0), papx, chpx, u16le(text.length), utf16le(text));
}

/** A list definition: its id, the styles of its levels and whether it has one level only. */
const lstf = (id: number, simple: boolean, styles: number[] = []): Uint8Array => {
  const levels = new Uint8Array(18).fill(0xff);
  for (let i = 0; i < 9; i++) levels.set(u16le(styles[i] === undefined ? 0x0fff : styles[i]), i * 2);
  return cat(u32le(id), u32le(0), levels, u8s(simple ? 1 : 0, 0));
};

function listDoc(): DocBuilder {
  const builder = new DocBuilder();
  const text = 'one\rtwo\rsub a\rsub b\rthree\rbullet\rrestart\r';
  builder.text(text);
  const list = (ilfo: number, level: number) => [sprm(0x460b, u16le(ilfo)), sprm(0x260a, level)];
  let at = 0;
  const para = (length: number, ...sprms: Uint8Array[]) => {
    builder.papx(at, at + length, 0, ...sprms);
    at += length;
  };
  para(4, ...list(1, 0));
  para(4, ...list(1, 0));
  para(6, ...list(1, 1));
  para(6, ...list(1, 1));
  para(6, ...list(1, 0));
  para(7, ...list(2, 0));
  para(8, ...list(3, 0));
  builder.table(FC.SttbfFfn, fontTable(['Times New Roman', 'Symbol']));
  builder.table(FC.Stshf, stsh([{ sti: 0, name: 'Normal' }, { sti: 1, name: 'heading 1', base: 0 }]));
  // A list of nine levels ("1.", "1.a)", ...) and a bullet list of one level. Their levels follow the definitions.
  const levels: Uint8Array[] = [];
  levels.push(lvl(1, 0, 0, 0, '\u0000.', sprm(0x845e, u16le(720)), sprm(0x0835, 1)));
  levels.push(lvl(1, 4, 2, 1, '\u0000.\u0001)', cat(sprm(0x845e, u16le(1440)), sprm(0x8460, u16le(0x10000 - 360)))));
  for (let i = 2; i < 9; i++) levels.push(lvl(1, 2, 0, 2, String.fromCharCode(i) + '.', undefined, undefined, i === 2 ? 0x04 | 0x08 : 0, i === 2 ? 1 : 0));
  levels.push(lvl(1, 23, 0, 0, '', undefined, sprm(0x4a4f, u16le(1))));
  builder.table(FC.PlfLst, cat(u16le(2), lstf(1000, false, [1]), lstf(2000, true), ...levels));
  // Three overrides: the first two plain; the third uses the first list again and restarts its first level at 7.
  const lfo = (id: number, levelCount: number) => cat(u32le(id), u32le(0), u32le(0), u8s(levelCount, 0, 0, 0));
  builder.table(
    FC.PlfLfo,
    cat(u32le(3), lfo(1000, 0), lfo(2000, 0), lfo(1000, 1), u32le(0xffffffff), u32le(0xffffffff), u32le(0xffffffff), u32le(7), u32le(0 | 0x10)),
  );
  return builder;
}

test('lists: definitions and overrides become the numbering of the model', () => {
  const model = listDoc().read();
  const { numbering } = model;
  assert.deepEqual([...numbering.abstracts.keys()], [0, 1]);
  assert.deepEqual([...numbering.nums.entries()].map(([id, num]) => [id, num.abstractId]), [
    [1, 0],
    [2, 1],
    [3, 0],
  ]);
  const first = numbering.abstracts.get(0)!;
  assert.equal(first.levels.size, 9);
  assert.equal(first.multiLevelType, 'multilevel');
  assert.deepEqual(first.levels.get(0), { level: 0, start: 1, format: 'decimal', text: '%1.', suffix: 'tab', align: 'left', paraStyle: 'Heading1', para: { indent: { left: 720 } }, run: { bold: true } });
  assert.deepEqual(first.levels.get(1), { level: 1, start: 1, format: 'lowerLetter', text: '%1.%2)', suffix: 'space', align: 'right', para: { indent: { left: 1440, hanging: 360 } } });
  assert.deepEqual(first.levels.get(2), { level: 2, start: 1, format: 'lowerRoman', text: '%3.', suffix: 'nothing', align: 'left', legal: true, restart: 1 });
  const bullets = numbering.abstracts.get(1)!;
  assert.equal(bullets.multiLevelType, 'singleLevel');
  assert.deepEqual(bullets.levels.get(0), { level: 0, start: 1, format: 'bullet', text: '', suffix: 'tab', align: 'left', run: { fonts: { ascii: 'Symbol' } } });
  assert.deepEqual([...numbering.nums.get(3)!.overrides], [[0, { start: 7 }]]);
  assert.equal(model.warnings.length, 0);
});

test('the counters written for docx give the labels of the paragraphs of a doc file', () => {
  const model = listDoc().read();
  const counters = new NumberingCounters(model.numbering, model.styles);
  const labels = paragraphs(model.body.blocks).map((paragraph) => {
    const numbering = resolveParaProps(model, paragraph.props).numbering!;
    return counters.next(numbering.numId!, numbering.level!)!.text;
  });
  assert.deepEqual(labels, ['1.', '2.', '2.a)', '2.b)', '3.', '', '7.']);
  // The indent of the level reaches the paragraph through the resolver.
  const [, , sub] = paragraphs(model.body.blocks);
  assert.deepEqual(resolveParaProps(model, sub.props).indent, { left: 1440, hanging: 360 });

  const layout = new DocLayout(model, fakeMeasurer);
  layout.layout();
  const shown = texts(layout.pages[0]).map((item) => item.text);
  assert.deepEqual(shown.slice(0, 4), ['1.', 'one', '2.', 'two']);
  assert.ok(shown.includes('2.a)'));
  assert.ok(shown.includes('•'), 'the bullet of the Symbol font is shown as a bullet');
});

test('a paragraph that names a list the file does not have is still read, without a label', () => {
  const builder = listDoc();
  const model = builder.read({
    patch: ({ table, fcLcb }) => {
      // Leave only the first override.
      const [at] = fcLcb(FC.PlfLfo);
      new DataView(table.buffer, table.byteOffset).setUint32(at, 1, true);
    },
  });
  assert.deepEqual([...model.numbering.nums.keys()], [1]);
  const layout = new DocLayout(model, fakeMeasurer);
  assert.equal(layout.layout().done, true);
  assert.equal(textOf(model.body.blocks).length, 7);
});

test('list levels that are cut short leave the lists read so far and a warning', () => {
  const model = listDoc().read({
    patch: ({ table, fcLcb }) => {
      // The text length of the second level reaches past the end of the stream.
      const [at] = fcLcb(FC.PlfLst);
      const firstLevel = at + 2 + 2 * 28;
      const second = firstLevel + 28 + 4 + 3 + 2 + 4;
      const view = new DataView(table.buffer, table.byteOffset);
      const text = second + 28 + view.getUint8(second + 24) + view.getUint8(second + 25);
      view.setUint16(text, 0xffff, true);
    },
  });
  assert.equal(model.numbering.abstracts.get(0)!.levels.size, 1);
  assert.match(model.warnings.join('\n'), /list definitions are cut short/);
  assert.equal(textOf(model.body.blocks).length, 7);
});

test('document properties: tab distance, facing pages, mirrored margins, hyphenation, tracked changes', () => {
  const builder = new DocBuilder();
  builder.text('x\r');
  builder.papx(0, 2, 0);
  const dop = new Uint8Array(84);
  dop[0] = 0x01;
  dop[5] = 0x10 | 0x80;
  dop[6] = 0x20;
  dop.set(u16le(708), 10);
  builder.table(FC.Dop, dop);
  const model: DocModel = builder.read();
  assert.equal(model.settings.defaultTabStop, 708);
  assert.equal(model.settings.evenAndOddHeaders, true);
  assert.equal(model.settings.mirrorMargins, true);
  assert.equal(model.settings.autoHyphenation, true);
  assert.equal(model.settings.trackRevisions, true);
  assert.equal(model.settings.compatSettings.get('compatibilityMode'), '11');
  // Without document properties the model has the defaults.
  const plain = new DocBuilder();
  plain.text('x\r');
  plain.papx(0, 2, 0);
  assert.equal(plain.read().settings.defaultTabStop, 720);
  assert.equal(findWrapper({}, 'del'), undefined);
});
