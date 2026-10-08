// Third review of the document formats: tracked changes read from RTF and doc files carry
// their authors and dates and get the ids a docx needs, deleted fields and paragraph marks of
// an RTF file are deleted in the model, a field whose end was deleted still ends, and a
// protection is read as loosely as files write it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, emptyWriteReport } from '../src/formats/docx/write';
import { enforcedProtection, findWrapper, paragraphText, tracksChanges, type DocModel, type Run } from '../src/model/doc/types';
import { DEFAULT_AUTHOR } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { DocBuilder, FC, paragraphs as docParagraphs, runsOf as docRuns, sprm, sttb, u16le, u32le } from './doc-binary-fixtures';
import { lay, p, pageText, r, texts } from './doc-layout-fixtures';
import { readSpec } from './docx-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';
import { paragraphs, read, runsOf } from './rtf-fixtures';

const decoder = new TextDecoder();

async function built(source: DocModel): Promise<{ xml: string; warnings: string[]; again: DocModel }> {
  const report = emptyWriteReport();
  const pkg = await OpcPackage.open(await buildDocx(source, { report }));
  const data = await pkg.readPart('/word/document.xml');
  parseXml(data);
  assert.deepEqual(schemaOrderViolations(data), []);
  return { xml: decoder.decode(data), warnings: report.warnings, again: await readDocx(pkg) };
}

/** The opening tags of the revisions in a part: wrappers, and the marks of paragraphs. */
const revisionTags = (xml: string): string[] => Array.from(xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*>/g), (match) => match[0]);

function checkRevisions(xml: string): string[] {
  const tags = revisionTags(xml);
  const ids: string[] = [];
  for (const tag of tags) {
    const id = /\sw:id="(\d+)"/.exec(tag);
    assert.ok(id, `no w:id: ${tag}`);
    assert.ok(/\sw:author="[^"]*"/.test(tag), `no w:author: ${tag}`);
    ids.push(id[1]);
  }
  // No id twice, and none that a bookmark of the part has.
  for (const match of xml.matchAll(/<w:bookmarkStart\b[^>]*\sw:id="(\d+)"/g)) ids.push(match[1]);
  assert.equal(new Set(ids).size, ids.length, ids.join(','));
  return tags;
}

// 2026-10-08 14:30 as a DTTM: minutes, hours << 6, day << 11, month << 16, (year - 1900) << 20.
const DTTM = 30 | (14 << 6) | (8 << 11) | (10 << 16) | (126 << 20);
const WHEN = '2026-10-08T14:30:00Z';

// ------------------------------------------------------------------------ RTF

test('review 3: RTF revisions carry the author of the revision table and the date', async () => {
  const source = read(`{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}{Bob;}}keep {\\deleted\\revauthdel1\\revdttmdel${DTTM} gone }{\\revised\\revauth2\\revdttm${DTTM} new }tail\\par}`);
  const [paragraph] = paragraphs(source);
  assert.equal(paragraphText(paragraph), 'keep new tail');
  const [, gone, added] = runsOf(paragraph);
  assert.deepEqual(findWrapper(gone, 'del'), { kind: 'del', author: 'Ann', date: WHEN });
  assert.deepEqual(findWrapper(added, 'ins'), { kind: 'ins', author: 'Bob', date: WHEN });
  // The table and its words are read, not counted as unknown.
  for (const word of ['revtbl', 'revauth', 'revauthdel', 'revdttm', 'revdttmdel']) assert.equal(source.unknown.get('rtf/' + word), undefined, word);

  const out = await built(source);
  const tags = checkRevisions(out.xml);
  assert.equal(tags.length, 2, out.xml);
  assert.ok(/^<w:del w:id="\d+" w:author="Ann" w:date="2026-10-08T14:30:00Z">$/.test(tags[0]), tags[0]);
  assert.ok(/^<w:ins w:id="\d+" w:author="Bob" w:date="2026-10-08T14:30:00Z">$/.test(tags[1]), tags[1]);
  assert.deepEqual(out.warnings.filter((line) => /tracked/i.test(line)), []);
  // Read again, the revisions are what they were, with their ids.
  const [again] = out.again.body.blocks;
  assert.equal(again.type === 'paragraph' && paragraphText(again), 'keep new tail');
});

test('review 3: the input of the review: a date that is none is left out, the author stays', async () => {
  // 1717969 says 27 hours: no date. The ids and the authors are there all the same.
  const source = read('{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}}keep {\\deleted\\revauthdel1\\revdttmdel1717969 gone }{\\revised\\revauth1 new }tail\\par}');
  const out = await built(source);
  const tags = checkRevisions(out.xml);
  assert.deepEqual(
    tags.map((tag) => tag.replace(/ w:id="\d+"/, '')),
    ['<w:del w:author="Ann">', '<w:ins w:author="Ann">'],
  );
  assert.deepEqual(out.warnings.filter((line) => /tracked/i.test(line)), []);
});

test('review 3: revisions without an author get the stand-in name and one line in the report', async () => {
  // No table at all, an index beyond the table, and no index.
  const source = read('{\\rtf1{\\*\\revtbl{Unknown;}}a{\\deleted b}c{\\revised\\revauth9 d}e{\\deleted\\revauthdel0 f}\\par{\\revised g\\par}h\\par}');
  const first = paragraphs(source)[0];
  assert.equal(findWrapper(runsOf(first)[1], 'del')!.author, undefined);
  assert.equal(findWrapper(runsOf(first)[3], 'ins')!.author, undefined);
  assert.equal(findWrapper(runsOf(first)[5], 'del')!.author, 'Unknown');
  const out = await built(source);
  const tags = checkRevisions(out.xml);
  // Three wrappers in the first paragraph, one in the second and its mark.
  assert.equal(tags.length, 5, out.xml);
  assert.equal(tags.filter((tag) => tag.indexOf(`w:author="${DEFAULT_AUTHOR}"`) > 0).length, 4, tags.join('\n'));
  assert.equal(tags.filter((tag) => tag.indexOf('w:date=') > 0).length, 0);
  const lines = out.warnings.filter((line) => /tracked/i.test(line));
  assert.equal(lines.length, 1, out.warnings.join('\n'));
  assert.ok(lines[0].indexOf('4') >= 0, lines[0]);
});

test('review 3: names of the revision table are text, long tables are read whole, and a table without end does no harm', () => {
  let table = '';
  for (let i = 0; i < 600; i++) table += `{Author number ${i};}`;
  const many = read(`{\\rtf1{\\*\\revtbl${table}}{\\revised\\revauth599 x}\\par}`);
  assert.equal(findWrapper(runsOf(paragraphs(many)[0])[0], 'ins')!.author, 'Author number 599');
  const odd = read('{\\rtf1{\\*\\revtbl{A<b>&"c";}{\\b D\\u350?;}loose{}}{\\revised\\revauth0 x}{\\revised\\revauth1 y}{\\revised\\revauth-5 z}\\par}');
  const runs = runsOf(paragraphs(odd)[0]);
  assert.equal(findWrapper(runs[0], 'ins')!.author, 'A<b>&"c"');
  assert.equal(findWrapper(runs[1], 'ins')!.author, 'DŞ');
  assert.equal(findWrapper(runs[2], 'ins')!.author, undefined);
  // Cut short inside the table: what was read is a document.
  assert.doesNotThrow(() => read('{\\rtf1 a\\par{\\*\\revtbl{Ann;}{Bo'));
  // Dates that are none.
  for (const value of [0, -1, 63, 31 << 11, 13 << 16, (2 << 16) | (30 << 11) | (126 << 20)]) {
    const model = read(`{\\rtf1{\\revised\\revdttm${value} x}\\par}`);
    assert.equal(findWrapper(runsOf(paragraphs(model)[0])[0], 'ins')!.date, undefined, String(value));
  }
});

test('review 3: a deleted field of an RTF file is deleted whole', async () => {
  const source = read('{\\rtf1 a{\\deleted\\revauthdel0 {\\field{\\*\\fldinst PAGE}{\\fldrslt 7}}}b\\par}');
  const [paragraph] = paragraphs(source);
  assert.equal(paragraphText(paragraph), 'ab');
  const marks = runsOf(paragraph).filter((run) => run.content.some((item) => item.type === 'fieldChar'));
  assert.equal(marks.length, 3);
  for (const mark of marks) assert.ok(findWrapper(mark, 'del'), 'a field mark outside the deletion');
  const out = await built(source);
  assert.equal(out.xml.indexOf('<w:instrText'), -1, out.xml);
  assert.ok(/<w:delInstrText[^>]*>\s*PAGE\s*<\/w:delInstrText>/.test(out.xml), out.xml);
  assert.ok(/<w:delText[^>]*>7<\/w:delText>/.test(out.xml), out.xml);
  // Every field mark is inside a deletion.
  const body = out.xml.slice(out.xml.indexOf('<w:body>'));
  const outside = body.replace(/<w:del\b[^>]*>.*?<\/w:del>/g, '');
  assert.equal(outside.indexOf('<w:fldChar'), -1, body);
  checkRevisions(out.xml);
  // The page shows the result the file has, struck through, and no live page number.
  const laid = await lay({ document: out.xml });
  const shown = texts(laid.pages[0]);
  assert.equal(shown.map((item) => item.text).join(''), 'a7b');
  assert.equal(shown.filter((item) => item.text === '7')[0].deleted, true);
  assert.equal(pageText(laid.pages[0]), 'ab');
});

test('review 3: a field whose code and result alone are deleted has deleted marks too', () => {
  const source = read('{\\rtf1 a{\\field{\\*\\fldinst {\\deleted PAGE}}{\\fldrslt {\\deleted 7}}}b\\par}');
  const marks = runsOf(paragraphs(source)[0]).filter((run) => run.content.some((item) => item.type === 'fieldChar'));
  assert.equal(marks.length, 3);
  for (const mark of marks) assert.ok(findWrapper(mark, 'del'));
  // A live field keeps live marks.
  const live = read('{\\rtf1 a{\\field{\\*\\fldinst PAGE}{\\fldrslt 7}}b\\par}');
  for (const run of runsOf(paragraphs(live)[0])) assert.equal(run.wrap, undefined);
});

test('review 3: the paragraph mark of an RTF file says that it was deleted or inserted', async () => {
  const source = read(`{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}}one{\\deleted\\revauthdel1\\revdttmdel${DTTM}\\par}two{\\revised\\revauth1\\par}three\\par}`);
  const [one, two, three] = paragraphs(source);
  assert.deepEqual(one.props.markProps && one.props.markProps.deleted, { author: 'Ann', date: WHEN });
  assert.equal(one.props.markProps!.inserted, undefined);
  assert.deepEqual(two.props.markProps && two.props.markProps.inserted, { author: 'Ann' });
  assert.equal(three.props.markProps === undefined || (three.props.markProps.deleted === undefined && three.props.markProps.inserted === undefined), true);
  // Formatting of the mark is still there beside the revision.
  const bold = read('{\\rtf1 one{\\b\\deleted\\par}two\\par}');
  const mark = paragraphs(bold)[0].props.markProps!;
  assert.equal(mark.bold, true);
  assert.deepEqual(mark.deleted, {});
  // And the next paragraph with the same formatting is not deleted with it.
  assert.equal(paragraphs(read('{\\rtf1 one{\\deleted\\par}two\\par three\\par}'))[1].props.markProps, undefined);
  const out = await built(source);
  const tags = checkRevisions(out.xml);
  assert.ok(/<w:rPr><w:del w:id="\d+" w:author="Ann" w:date="2026-10-08T14:30:00Z"\/><\/w:rPr><\/w:pPr>/.test(out.xml), out.xml);
  assert.equal(tags.length, 2);
});

// ------------------------------------------------------------------------ doc

test('review 3: doc revisions get ids in the docx, their authors and dates stay, and deleted field code is deleted code', async () => {
  const builder = new DocBuilder();
  const text = 'kept gone new \u0013 PAGE \u00147\u0015 end\rnext\r';
  builder.text(text);
  builder.papxAll(text);
  builder.table(FC.SttbfRMark, sttb(['Unknown', 'Ayşe']));
  const del = [sprm(0x0800, 1), sprm(0x4863, u16le(1)), sprm(0x6864, u32le(DTTM))];
  builder.chpx(5, 10, ...del);
  builder.chpx(10, 14, sprm(0x0801, 1), sprm(0x4804, u16le(1)));
  // The field, all of it deleted; its marks are special characters.
  for (let i = 14; i < 24; i++) {
    const special = text.charCodeAt(i) >= 0x13 && text.charCodeAt(i) <= 0x15;
    builder.chpx(i, i + 1, ...(special ? [sprm(0x0855, 1)] : []), ...del);
  }
  // The mark of the first paragraph was deleted too; no author is named for it.
  const mark = text.indexOf('\r');
  builder.chpx(mark, mark + 1, sprm(0x0800, 1));
  const source = builder.read();
  const [first, second] = docParagraphs(source.body.blocks);
  assert.equal(paragraphText(first), 'kept new  end');
  assert.deepEqual(first.props.markProps && first.props.markProps.deleted, {});
  assert.equal(second.props.markProps, undefined);
  const code = docRuns(first).filter((run: Run) => run.content.some((item) => item.type === 'instrText'));
  assert.ok(code.length > 0);
  for (const run of code) for (const item of run.content) if (item.type === 'instrText') assert.equal(item.deleted, true);

  const out = await built(source);
  const tags = checkRevisions(out.xml);
  assert.ok(tags.length >= 3, out.xml);
  assert.ok(/^<w:del w:id="\d+" w:author="Ayşe" w:date="2026-10-08T14:30:00Z">$/.test(tags.filter((tag) => tag.indexOf('<w:del') === 0 && tag.charAt(tag.length - 2) !== '/')[0]), tags.join('\n'));
  assert.ok(tags.some((tag) => /^<w:ins w:id="\d+" w:author="Ayşe">$/.test(tag)), tags.join('\n'));
  // The mark of the paragraph: an id, the stand-in author, no date.
  assert.ok(new RegExp(`<w:rPr><w:del w:id="\\d+" w:author="${DEFAULT_AUTHOR}"/></w:rPr>`).test(out.xml), out.xml);
  assert.equal(out.xml.indexOf('<w:instrText'), -1, out.xml);
  assert.ok(out.xml.indexOf('<w:delInstrText') > 0, out.xml);
  const lines = out.warnings.filter((line) => /tracked/i.test(line));
  assert.equal(lines.length, 1, out.warnings.join('\n'));
  assert.ok(lines[0].indexOf('1') >= 0, lines[0]);
});

test('review 3: a day its month does not have is no date (doc and RTF)', () => {
  const builder = new DocBuilder();
  const text = 'ab\r';
  builder.text(text);
  builder.papxAll(text);
  // 31 February 2026.
  builder.chpx(0, 1, sprm(0x0801, 1), sprm(0x6805, u32le((31 << 11) | (2 << 16) | (126 << 20))));
  // 29 February 2024: a leap year.
  builder.chpx(1, 2, sprm(0x0800, 1), sprm(0x6864, u32le((29 << 11) | (2 << 16) | (124 << 20))));
  const runs = docRuns(docParagraphs(builder.read().body.blocks)[0]);
  assert.equal(findWrapper(runs[0], 'ins')!.date, undefined);
  assert.equal(findWrapper(runs[1], 'del')!.date, '2024-02-29T00:00:00Z');
});

// --------------------------------------------------------------------- layout

test('review 3: a field whose result and end were deleted ends there: the text after it shows', async () => {
  const fld = (type: string) => `<w:r><w:fldChar w:fldCharType="${type}"/></w:r>`;
  const del = (inner: string) => `<w:del w:id="1" w:author="Ann">${inner}</w:del>`;
  const body = p(r('a') + fld('begin') + '<w:r><w:instrText> PAGE </w:instrText></w:r>' + fld('separate') + del('<w:r><w:delText>7</w:delText></w:r>' + fld('end')) + r('after')) + p('next');
  const laid = await lay(body);
  const page = laid.pages[0];
  const shown = texts(page).map((item) => item.text).join('|');
  assert.ok(shown.indexOf('after') >= 0, shown);
  assert.ok(pageText(page).indexOf('after') >= 0, pageText(page));
  // A field deleted whole inside the result of a live field closes nothing but itself.
  const nested =
    p(
      r('a') +
        fld('begin') +
        '<w:r><w:instrText> PAGE </w:instrText></w:r>' +
        fld('separate') +
        del(fld('begin') + '<w:r><w:delInstrText> DATE </w:delInstrText></w:r>' + fld('separate') + '<w:r><w:delText>x</w:delText></w:r>' + fld('end')) +
        r('9') +
        fld('end') +
        r('z'),
    ) + p('next');
  const inner = await lay(nested);
  // The live field shows the page number once, in place of its result "9".
  assert.equal(pageText(inner.pages[0]).indexOf('9'), -1, pageText(inner.pages[0]));
  assert.ok(pageText(inner.pages[0]).indexOf('z') >= 0, pageText(inner.pages[0]));
});

// ----------------------------------------------------------------- protection

const BODY = '<w:p><w:r><w:t>text</w:t></w:r></w:p>';

test('review 3: the enforcement and the kind of a protection are read whatever their case and the space around them', async () => {
  for (const value of ['TRUE', 'True', 'ON', 'On', ' 1 ', '\ttrue\n']) {
    const doc = await readSpec({ body: BODY, settings: `<w:documentProtection w:edit="readOnly" w:enforcement="${value.replace('\t', '&#9;').replace('\n', '&#10;')}"/>` });
    assert.equal(doc.settings.protection && doc.settings.protection.enforced, true, JSON.stringify(value));
    assert.equal(enforcedProtection(doc), 'readOnly', JSON.stringify(value));
  }
  for (const value of ['FALSE', 'Off', ' 0 ', 'yes', '']) {
    const doc = await readSpec({ body: BODY, settings: `<w:documentProtection w:edit="readOnly" w:enforcement="${value}"/>` });
    assert.equal(enforcedProtection(doc), undefined, JSON.stringify(value));
  }
  const kinds: [string, ReturnType<typeof enforcedProtection>][] = [
    ['TrackedChanges', 'trackedChanges'],
    [' trackedchanges ', 'trackedChanges'],
    ['READONLY', 'readOnly'],
    ['Comments', 'comments'],
    ['FORMS', 'forms'],
    ['None', undefined],
    [' none ', undefined],
  ];
  for (const [edit, expected] of kinds) {
    const doc = await readSpec({ body: BODY, settings: `<w:documentProtection w:edit="${edit}" w:enforcement="1"/>` });
    assert.equal(enforcedProtection(doc), expected, edit);
    assert.equal(tracksChanges(doc), expected === 'trackedChanges', edit);
  }
});

test('review 3: of several protections the first one that is enforced counts; the elements stay as they are', async () => {
  const settings = '<w:documentProtection/><w:documentProtection w:edit="none" w:enforcement="0"/><w:documentProtection w:edit="forms" w:enforcement="1" w:hash="q83vEjRWeJA="/><w:documentProtection w:edit="readOnly" w:enforcement="1"/>';
  const doc = await readSpec({ body: BODY, settings });
  assert.equal(enforcedProtection(doc), 'forms');
  assert.equal(doc.settings.protection!.xml!.attrs.get('w:hash'), 'q83vEjRWeJA=');
  // None is enforced: the first one is the protection of the document, as before.
  const none = await readSpec({ body: BODY, settings: '<w:documentProtection w:edit="comments"/><w:documentProtection w:edit="forms" w:enforcement="0"/>' });
  assert.equal(none.settings.protection!.edit, 'comments');
  assert.equal(enforcedProtection(none), undefined);
});
