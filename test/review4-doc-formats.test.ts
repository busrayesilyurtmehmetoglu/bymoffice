// Fourth review of the document formats: text of an RTF or doc file that was inserted and then
// deleted is a deletion inside an insertion, as a docx has it, and two changes of such a file
// that stand apart are two objects of the model, resolved one without the other.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { buildDocx, emptyWriteReport } from '../src/formats/docx/write';
import { paragraphText, type DocModel, type Paragraph, type Run } from '../src/model/doc/types';
import { DocHistory, acceptAll, acceptRevision, rejectAll, rejectRevision, revisionsOf, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { DocBuilder, FC, paragraphs as docParagraphs, sprm, sttb, u16le, u32le } from './doc-binary-fixtures';
import { schemaOrderViolations } from './docx-write-helpers.test';
import { paragraphs, read } from './rtf-fixtures';

const decoder = new TextDecoder();

async function built(source: DocModel): Promise<{ xml: string; warnings: string[]; again: DocModel; bytes: Uint8Array }> {
  const report = emptyWriteReport();
  const bytes = await buildDocx(source, { report });
  const pkg = await OpcPackage.open(bytes);
  const data = await pkg.readPart('/word/document.xml');
  parseXml(data);
  assert.deepEqual(schemaOrderViolations(data), []);
  return { xml: decoder.decode(data), warnings: report.warnings, again: await readDocx(pkg), bytes };
}

const bodyText = (doc: DocModel): string[] => doc.body.blocks.filter((block): block is Paragraph => block.type === 'paragraph').map(paragraphText);

function resolved(doc: DocModel, how: 'accept' | 'reject'): string[] {
  const history = new DocHistory<DocRange | undefined>(doc);
  const edit = how === 'accept' ? acceptAll(doc) : rejectAll(doc);
  assert.equal(edit.skipped, 0);
  history.applyEdit(edit);
  assert.deepEqual(revisionsOf(doc), []);
  return bodyText(doc);
}

/** The wrappers around a run, innermost first, as "kind author". */
const chainOf = (run: Run): string[] => {
  const out: string[] = [];
  for (let w = run.wrap; w; w = w.parent) out.push(`${w.kind} ${(w as { author?: string }).author}`);
  return out;
};

const runsIn = (paragraph: Paragraph): Run[] => paragraph.children.filter((inline): inline is Run => inline.type === 'run');

// 2026-10-08 14:30 and 15:45 as DTTM: minutes, hours << 6, day << 11, month << 16, (year - 1900) << 20.
const DTTM = 30 | (14 << 6) | (8 << 11) | (10 << 16) | (126 << 20);
const LATER = 45 | (15 << 6) | (8 << 11) | (10 << 16) | (126 << 20);

const RTF_BOTH = `{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}{Bob;}}keep {\\revised\\revauth1\\revdttm${DTTM} insA {\\deleted\\revauthdel2\\revdttmdel${LATER} insA-delB }more }tail\\par}`;

// ------------------------------------------------------------------------- A3

test('review 4 (A3): RTF text that was inserted and then deleted is a deletion inside the insertion', async () => {
  const source = read(RTF_BOTH);
  const [paragraph] = paragraphs(source);
  assert.equal(paragraphText(paragraph), 'keep insA more tail');
  const runs = runsIn(paragraph);
  assert.deepEqual(runs.map(chainOf), [[], ['ins Ann'], ['del Bob', 'ins Ann'], ['ins Ann'], []]);
  // One insertion holds all three: the deletion is inside it.
  assert.equal(runs[1].wrap, runs[3].wrap);
  assert.equal(runs[2].wrap && runs[2].wrap.parent, runs[1].wrap);
  assert.deepEqual(runs[2].content, [{ type: 'deletedText', text: 'insA-delB ' }]);

  const out = await built(source);
  assert.ok(
    /<w:ins w:id="(\d+)" w:author="Ann" w:date="2026-10-08T14:30:00Z"><w:r>.*?insA .*?<\/w:r><w:del w:id="(\d+)" w:author="Bob" w:date="2026-10-08T15:45:00Z"><w:r>.*?<w:delText[^>]*>insA-delB <\/w:delText><\/w:r><\/w:del><w:r>.*?more .*?<\/w:r><\/w:ins>/.test(out.xml),
    out.xml,
  );
  const ids = Array.from(out.xml.matchAll(/<w:(?:ins|del)\b[^>]*? w:id="(\d+)"/g), (match) => match[1]);
  assert.equal(ids.length, 2, out.xml);
  assert.equal(new Set(ids).size, 2);
  const list = revisionsOf(out.again);
  assert.deepEqual(
    list.map((revision) => [revision.kind, revision.author, revision.text]),
    [
      ['insert', 'Ann', 'insA more '],
      ['delete', 'Bob', 'insA-delB '],
    ],
  );
  assert.deepEqual(resolved(out.again, 'reject'), ['keep tail']);
  assert.deepEqual(resolved((await built(source)).again, 'accept'), ['keep insA more tail']);
  // The model as read from the RTF file resolves the same way.
  assert.deepEqual(resolved(read(RTF_BOTH), 'reject'), ['keep tail']);
  assert.deepEqual(resolved(read(RTF_BOTH), 'accept'), ['keep insA more tail']);
  // Rejecting the deletion alone gives the text back, inside the insertion.
  const one = read(RTF_BOTH);
  new DocHistory<DocRange | undefined>(one).applyEdit(rejectRevision(one, revisionsOf(one)[1]));
  assert.deepEqual(bodyText(one), ['keep insA insA-delB more tail']);
  assert.equal(revisionsOf(one).length, 1);
});

test('review 4 (A3): a field of an RTF file that was inserted and then deleted', () => {
  const source = read(`{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}{Bob;}}a{\\revised\\revauth1\\deleted\\revauthdel2 {\\field{\\*\\fldinst PAGE}{\\fldrslt 7}}}b\\par}`);
  assert.deepEqual(resolved(source, 'reject'), ['ab']);
  const kept = read(`{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}{Bob;}}a{\\revised\\revauth1\\deleted\\revauthdel2 {\\field{\\*\\fldinst PAGE}{\\fldrslt 7}}}b\\par}`);
  for (const run of runsIn(paragraphs(kept)[0]).slice(1, -1)) assert.deepEqual(chainOf(run), ['del Bob', 'ins Ann']);
});

function docBoth(): DocModel {
  const builder = new DocBuilder();
  const text = 'keep insA gone more tail\r';
  builder.text(text);
  builder.papxAll(text);
  builder.table(FC.SttbfRMark, sttb(['Unknown', 'Ann', 'Bob']));
  const ins = [sprm(0x0801, 1), sprm(0x4804, u16le(1)), sprm(0x6805, u32le(DTTM))];
  const del = [sprm(0x0800, 1), sprm(0x4863, u16le(2)), sprm(0x6864, u32le(LATER))];
  builder.chpx(5, 10, ...ins);
  builder.chpx(10, 15, ...ins, ...del);
  builder.chpx(15, 20, ...ins);
  return builder.read();
}

test('review 4 (A3): doc text that was inserted and then deleted is a deletion inside the insertion', async () => {
  const source = docBoth();
  const [paragraph] = docParagraphs(source.body.blocks);
  assert.equal(paragraphText(paragraph), 'keep insA more tail');
  const runs = runsIn(paragraph);
  assert.deepEqual(runs.map(chainOf), [[], ['ins Ann'], ['del Bob', 'ins Ann'], ['ins Ann'], []]);
  assert.equal(runs[1].wrap, runs[3].wrap);
  assert.equal(runs[2].wrap && runs[2].wrap.parent, runs[1].wrap);
  const out = await built(source);
  assert.ok(
    /<w:ins w:id="\d+" w:author="Ann" w:date="2026-10-08T14:30:00Z"><w:r>.*?insA .*?<\/w:r><w:del w:id="\d+" w:author="Bob" w:date="2026-10-08T15:45:00Z"><w:r>.*?<w:delText[^>]*>gone <\/w:delText><\/w:r><\/w:del><w:r>.*?more .*?<\/w:r><\/w:ins>/.test(out.xml),
    out.xml,
  );
  assert.deepEqual(resolved(out.again, 'reject'), ['keep tail']);
  assert.deepEqual(resolved(docBoth(), 'reject'), ['keep tail']);
  assert.deepEqual(resolved(docBoth(), 'accept'), ['keep insA more tail']);
});

// ------------------------------------------------------------------------- A6

test('review 4 (A6): two changes of an RTF file by one author at one time, apart, are resolved one without the other', () => {
  const source = `{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}}{\\revised\\revauth1\\revdttm${DTTM} one }mid {\\revised\\revauth1\\revdttm${DTTM} two }tail\\par{\\revised\\revauth1\\revdttm${DTTM} three }\\par}`;
  const doc = read(source);
  const [first, second] = paragraphs(doc);
  const runs = runsIn(first);
  assert.notEqual(runs[0].wrap, runs[2].wrap);
  assert.notEqual(runs[0].wrap, runsIn(second)[0].wrap);
  const list = revisionsOf(doc);
  assert.equal(list.length, 3);
  const history = new DocHistory<DocRange | undefined>(doc);
  history.applyEdit(acceptRevision(doc, list[1]));
  const left = revisionsOf(doc);
  assert.deepEqual(
    left.map((revision) => revision.text),
    ['one ', 'three '],
  );
  history.applyEdit(rejectRevision(doc, left[0]));
  assert.deepEqual(bodyText(doc), ['mid two tail', 'three ']);
  assert.equal(revisionsOf(doc).length, 1);
});

test('review 4 (A6): runs of one change that follow each other in different formatting are still one change', () => {
  const doc = read(`{\\rtf1{\\*\\revtbl{Unknown;}{Ann;}}a{\\revised\\revauth1\\revdttm${DTTM} one {\\b bold} more}z\\par}`);
  const runs = runsIn(paragraphs(doc)[0]);
  assert.equal(runs[1].wrap, runs[2].wrap);
  assert.equal(runs[2].wrap, runs[3].wrap);
  assert.equal(revisionsOf(doc).length, 1);
});

test('review 4 (A6): two changes of a doc file by one author at one time, apart, are resolved one without the other', () => {
  const builder = new DocBuilder();
  const text = 'one mid two tail\r';
  builder.text(text);
  builder.papxAll(text);
  builder.table(FC.SttbfRMark, sttb(['Unknown', 'Ann']));
  const ins = [sprm(0x0801, 1), sprm(0x4804, u16le(1)), sprm(0x6805, u32le(DTTM))];
  builder.chpx(0, 4, ...ins);
  builder.chpx(8, 12, ...ins);
  const doc = builder.read();
  const list = revisionsOf(doc);
  assert.equal(list.length, 2);
  new DocHistory<DocRange | undefined>(doc).applyEdit(rejectRevision(doc, list[1]));
  assert.deepEqual(bodyText(doc), ['one mid tail']);
  assert.equal(revisionsOf(doc).length, 1);
});
