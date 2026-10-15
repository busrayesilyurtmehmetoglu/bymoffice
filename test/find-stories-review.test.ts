// Find and Replace beyond the body, after its review: a text box is changed only where the writer
// can write the change back in every branch of alternate content (one rule for both, decided when
// the edit is made); the reader and the writer count the text boxes of a graphic alike; a nested
// chain of text boxes is written in time linear in its size; what may not be replaced (locked
// content controls on rows and cells, alternate content around blocks, a box in a locked control)
// is marked so when it is found; a match does not run through a floating drawing; the replacement
// stays inside the content control, bookmark or comment range the match was in.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel, Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { READ_ONLY } from '../src/ops/doc/pieces';
import { DocHistory, findAllInDocument, insertTextInTextBox, replaceMatchEdit, replaceMatches, type DocEdit } from '../src/ops/doc/index';
import { SECT, STYLES, p, r, tbl, tc, tr } from './doc-ops-fixtures';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const decoder = new TextDecoder();

const ANCHOR_HEAD =
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/>';
const drawing = (inner: string, id = 1) =>
  `<w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
const pict = (inner: string) => `<w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict>`;
const alt = (choice: string, fallback: string, requires = 'wps') => `<w:r><mc:AlternateContent><mc:Choice Requires="${requires}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent></w:r>`;
const wordBox = (inner: string, id = 1) => alt(drawing(inner, id), pict(inner));

interface Opened {
  bytes: Uint8Array;
  pkg: OpcPackage;
  doc: DocModel;
}

async function openSpec(spec: DocxSpec | string): Promise<Opened> {
  const bytes = await buildDocx(typeof spec === 'string' ? { body: spec + SECT, styles: STYLES } : spec);
  const pkg = await OpcPackage.open(bytes);
  return { bytes, pkg, doc: await readDocx(pkg) };
}

async function saved(opened: Opened): Promise<Opened & { xml: string }> {
  await writeDocx(opened.pkg, opened.doc);
  const bytes = await opened.pkg.save();
  const pkg = await OpcPackage.open(bytes);
  return { bytes, pkg, doc: await readDocx(pkg), xml: decoder.decode(await pkg.readPart('/word/document.xml')) };
}

const documentXml = async (bytes: Uint8Array): Promise<string> => decoder.decode(await (await OpcPackage.open(bytes)).readPart('/word/document.xml'));
const body = (xml: string): string => xml.slice(xml.indexOf('<w:body>'), xml.indexOf('<w:sectPr'));
const count = (text: string, part: string): number => text.split(part).length - 1;

function replaceAll(doc: DocModel, text: string, replacement: string) {
  const history = new DocHistory(doc);
  const result = replaceMatches(doc, findAllInDocument(doc, text), replacement, { apply: (edit: DocEdit) => history.applyEdit({ ...edit, mergeKey: 'all' }, undefined) });
  history.seal();
  return result;
}

let python: boolean | undefined;
function packageProblems(bytes: Uint8Array): unknown[] {
  if (python === undefined) {
    try {
      execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
      python = true;
    } catch {
      python = false;
    }
  }
  if (!python) return [];
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.find-review-'));
  try {
    writeFileSync(joinPath(dir, 'file.docx'), bytes);
    return JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 120000 })) as unknown[];
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

// ----------------------------------------------------- what the writer can write

test('a text box whose other branch could not take the change is found, marked, not replaced, and the file stays as it was', async () => {
  const cases: [string, string][] = [
    ['the Fallback holds other text', alt(drawing(p('cat one')), pict(p('OLD cat FALLBACK')))],
    ['the Fallback holds no text box', alt(drawing(p('cat one')), '<w:pict><v:rect style="width:1pt"/></w:pict>')],
    ['the Fallback holds two text boxes', alt(drawing(p('cat one')), pict(p('cat one')) + pict(p('x')))],
    ['a Choice that is not taken holds no box', `<w:r><mc:AlternateContent><mc:Choice Requires="zzz"><w:t>no</w:t></mc:Choice><mc:Fallback>${pict(p('cat fb'))}</mc:Fallback></mc:AlternateContent></w:r>`],
  ];
  for (const [what, run] of cases) {
    const opened = await openSpec(p(r('anchor ') + run));
    const matches = findAllInDocument(opened.doc, 'cat');
    assert.equal(matches.length, 1, what);
    assert.equal(matches[0].replaceable, false, what);
    assert.equal(matches[0].fixed, 'textBox', what);
    assert.throws(() => insertTextInTextBox(opened.doc, matches[0].range, 'dog'), (error: Error) => error instanceof RangeError && error.message === READ_ONLY, what);
    const result = replaceAll(opened.doc, 'cat', 'dog');
    assert.deepEqual([result.replaced, result.skipped, result.textBoxesKept], [0, 1, 1], what);
    const after = await saved(opened);
    assert.equal(after.xml, await documentXml(opened.bytes), what);
  }
});

test('a box Word wrote with the same text in both branches: both take the replacement; a content control in it keeps its id in both', async () => {
  const inner = p('cat one') + `<w:p><w:sdt><w:sdtPr><w:id w:val="77"/></w:sdtPr><w:sdtContent>${r('cat two')}</w:sdtContent></w:sdt></w:p>`;
  const opened = await openSpec(p(r('anchor ') + wordBox(inner)));
  assert.equal(replaceAll(opened.doc, 'cat', 'dog').replaced, 2);
  const after = await saved(opened);
  assert.equal(count(after.xml, '<mc:AlternateContent>'), 1);
  assert.equal(count(after.xml, 'cat'), 0);
  assert.equal(count(after.xml, 'dog one'), 2);
  assert.equal(count(after.xml, '<w:id w:val="77"/>'), 2);
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('the reader and the writer count the text boxes of a drawing alike: one outside its graphic is no box of it', async () => {
  const odd =
    `<w:r><w:drawing>${ANCHOR_HEAD}<wp:docPr id="1" name="B"/><x:ext xmlns:x="urn:x"><w:txbxContent><w:p><w:r><w:t>hidden</w:t></w:r></w:p></w:txbxContent></x:ext>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>${p('cat seen')}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
  const opened = await openSpec(p(r('a ') + odd));
  assert.equal(replaceAll(opened.doc, 'cat', 'dog').replaced, 1);
  const after = await saved(opened);
  assert.match(after.xml, /dog seen/);
  assert.match(after.xml, /hidden/);
  assert.equal(findAllInDocument(after.doc, 'dog').length, 1);
});

test('should the writer meet a box it cannot write in every branch, it writes the element as it was read: nothing of the file is lost', async () => {
  const opened = await openSpec(p(r('anchor ') + alt(drawing(p('cat one')), pict(p('cat one')) + pict(p('x')))));
  // The model changed behind the rules (not through Find and Replace).
  const paragraph = opened.doc.body.blocks[0] as Paragraph;
  const run = paragraph.children[1];
  assert.ok(run.type === 'run' && run.content[0].type === 'alternate');
  const holder = run.content[0] as Extract<typeof run.content[0], { type: 'alternate' }>;
  const graphic = holder.items[0] as Extract<typeof holder.items[0], { type: 'drawing' }>;
  const box = graphic.textBoxes![0];
  const changed = { ...(box[0] as Paragraph), children: [{ type: 'run', props: (box[0] as Paragraph).children[0].type === 'run' ? ((box[0] as Paragraph).children[0] as { props: never }).props : {}, content: [{ type: 'text', text: 'dog one' }] }] } as unknown as Paragraph;
  const next = { ...holder, items: [{ ...graphic, textBoxes: [[changed]] }] };
  opened.doc.body.blocks[0] = { type: 'paragraph', props: paragraph.props, children: [paragraph.children[0], { ...run, content: [next] }] } as Paragraph;
  opened.doc.body.modified = true;
  const after = await saved(opened);
  assert.equal(count(after.xml, '<mc:Fallback>'), 1);
  assert.equal(count(after.xml, '<v:textbox>'), 2);
  assert.match(after.xml, />x<\/w:t>/);
});

// ------------------------------------------------------------------ nesting

test('a chain of text boxes in text boxes is replaced and written in time that follows its size, not doubling with each level', async () => {
  const chain = (depth: number, kind: 'word' | 'vml' | 'wps'): string => {
    let inner = p('cat');
    for (let k = 0; k < depth; k++) inner = p(r('L') + (kind === 'word' ? wordBox(inner, k + 1) : kind === 'vml' ? `<w:r>${pict(inner)}</w:r>` : `<w:r>${drawing(inner, k + 1)}</w:r>`));
    return inner;
  };
  for (const [depth, kind] of [
    [16, 'vml'],
    [16, 'wps'],
    [9, 'word'],
  ] as [number, 'word' | 'vml' | 'wps'][]) {
    const source = chain(depth, kind);
    const opened = await openSpec(source);
    const started = Date.now();
    const result = replaceAll(opened.doc, 'cat', 'dog');
    const after = await saved(opened);
    const took = Date.now() - started;
    assert.equal(result.replaced, 1, `${kind} ${depth}`);
    // Every copy of the innermost box (a Word box holds each level twice) has the new text, and the part is no larger than it was.
    assert.equal(count(after.xml, 'cat'), 0, `${kind} ${depth}`);
    assert.equal(count(after.xml, 'dog'), count(source, 'cat'), `${kind} ${depth}`);
    assert.ok(after.xml.length < source.length * 1.2 + 4000, `${kind} ${depth}: ${after.xml.length} for ${source.length}`);
    // A guard against a blow-up, not a measure of speed (the review measured 35 s for 17.5 KB).
    assert.ok(took < 20000, `${kind} ${depth}: ${took} ms`);
    assert.deepEqual(packageProblems(after.bytes), [], `${kind} ${depth}`);
  }
});

// --------------------------------------------------------- what may not be replaced

test('locked content controls on rows, on cells and around a text box, alternate content around blocks: found, not replaceable', async () => {
  const grid = '<w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>';
  const cases: [string, string, string][] = [
    ['row', `<w:tbl>${grid}<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent><w:tr><w:tc>${p('cat row')}</w:tc></w:tr></w:sdtContent></w:sdt></w:tbl>`, 'locked'],
    ['cell', `<w:tbl>${grid}<w:tr><w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent><w:tc>${p('cat cell')}</w:tc></w:sdtContent></w:sdt></w:tr></w:tbl>`, 'locked'],
    ['box in a locked control', `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${p(r('a ') + wordBox(p('cat inbox')))}</w:sdtContent></w:sdt>`, 'locked'],
    ['box in a locked inline control', p(r('a ') + `<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>${`<w:r>${pict(p('cat inbox'))}</w:r>`}</w:sdtContent></w:sdt>`), 'locked'],
    ['blocks in alternate content', `<mc:AlternateContent><mc:Choice Requires="w14">${p('cat ch')}</mc:Choice><mc:Fallback>${p('cat fb')}</mc:Fallback></mc:AlternateContent>`, 'locked'],
    ['blocks in alternate content in a box', p(r('a ') + `<w:r>${pict(`<mc:AlternateContent><mc:Choice Requires="w14">${p('cat ch')}</mc:Choice><mc:Fallback>${p('cat fb')}</mc:Fallback></mc:AlternateContent>`)}</w:r>`), 'locked'],
  ];
  for (const [what, xml, why] of cases) {
    const opened = await openSpec(xml + p('free cat'));
    const matches = findAllInDocument(opened.doc, 'cat');
    assert.equal(matches.length, 2, what);
    const kept = matches.filter((match) => !match.replaceable);
    assert.equal(kept.length, 1, what);
    assert.equal(kept[0].fixed, why, what);
    const result = replaceAll(opened.doc, 'cat', 'dog');
    assert.deepEqual([result.replaced, result.skipped], [1, 1], what);
    const after = await saved(opened);
    assert.equal(findAllInDocument(after.doc, 'cat').length, 1, what);
  }
});

test('a match does not run through a floating text box or a drawing in alternate content: replacing would take it away', async () => {
  const opened = await openSpec(p(r('ca') + wordBox(p('box text')) + r('t')) + p(r('ca') + `<w:r>${pict(p('vml'))}</w:r>` + r('t')));
  assert.deepEqual(findAllInDocument(opened.doc, 'cat'), []);
  assert.equal(findAllInDocument(opened.doc, 'ca').length, 2);
});

// -------------------------------------------------- where the replacement goes

test('the replacement stays in the content control whose whole text it replaces, with its properties, also replaced by nothing', async () => {
  for (const [lock, replacement] of [
    ['', 'dog'],
    ['<w:lock w:val="sdtLocked"/>', 'dog'],
    ['<w:lock w:val="sdtLocked"/>', ''],
  ]) {
    const opened = await openSpec(p(r('x ') + `<w:sdt><w:sdtPr><w:tag w:val="T1"/>${lock}</w:sdtPr><w:sdtContent>${r('cat')}</w:sdtContent></w:sdt>`));
    assert.equal(replaceAll(opened.doc, 'cat', replacement).replaced, 1);
    const after = await saved(opened);
    const text = body(after.xml);
    assert.equal(count(text, '<w:sdt>'), 1, `${lock} "${replacement}"`);
    assert.match(text, new RegExp(`<w:sdtContent><w:r><w:t(?: [^>]*)?(?:/>|>${replacement}</w:t>)</w:r></w:sdtContent>`), `${lock} "${replacement}": ${text}`);
    if (lock) assert.match(text, /<w:lock w:val="sdtLocked"\/>/);
    assert.deepEqual(packageProblems(after.bytes), []);
  }
});

test('the replacement stays inside the bookmark and the comment range the match begins or ends with, in the body and in a text box', async () => {
  const inner = '<w:p><w:bookmarkStart w:id="5" w:name="bm"/><w:r><w:t xml:space="preserve">cat </w:t></w:r><w:bookmarkEnd w:id="5"/><w:commentRangeStart w:id="0"/><w:r><w:t>see cat</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>';
  const comments = '<w:comment w:id="0" w:author="A"><w:p><w:r><w:t>c</w:t></w:r></w:p></w:comment>';
  for (const where of ['body', 'box']) {
    const opened = await openSpec({ body: (where === 'body' ? inner : p(r('a ') + wordBox(inner))) + SECT, styles: STYLES, comments });
    assert.equal(replaceAll(opened.doc, 'cat', 'dog').replaced, 2, where);
    const after = await saved(opened);
    const text = body(after.xml);
    assert.match(text, /<w:bookmarkStart w:id="5" w:name="bm"\/><w:r><w:t xml:space="preserve">dog <\/w:t><\/w:r><w:bookmarkEnd w:id="5"\/>/, where);
    assert.match(text, /<w:commentRangeStart w:id="0"\/><w:r><w:t>see dog<\/w:t><\/w:r><w:commentRangeEnd w:id="0"\/>/, where);
  }
});

test('the replacement takes the formatting of the first matched character, in the body and in a text box', async () => {
  const runs = r('bold ', '<w:b/>') + r('cat', '<w:i/>') + r(' tail');
  const opened = await openSpec(p(runs) + p(r('a ') + `<w:r>${pict(p(runs))}</w:r>`));
  assert.equal(replaceAll(opened.doc, 'cat', 'dog').replaced, 2);
  const after = await saved(opened);
  assert.equal(count(after.xml, '<w:i/></w:rPr><w:t>dog</w:t>'), 2);
});

test('replaceMatchEdit: the edit of Replace, in the body and in a text box', async () => {
  const opened = await openSpec(p(r('one cat ') + `<w:r>${pict(p('box cat'))}</w:r>`));
  const history = new DocHistory(opened.doc);
  for (const match of findAllInDocument(opened.doc, 'cat').reverse()) history.applyEdit(replaceMatchEdit(opened.doc, match.range, 'owl'), undefined);
  assert.deepEqual(
    findAllInDocument(opened.doc, 'owl').map((match) => match.range.anchor.path.join('.')),
    ['0', '0.-1.0.0.0'],
  );
  // With a paragraph end the replacement is insertText(): two paragraphs.
  const second = await openSpec(p('a cat b'));
  const edit = replaceMatchEdit(second.doc, findAllInDocument(second.doc, 'cat')[0].range, 'x\ny');
  new DocHistory(second.doc).applyEdit(edit, undefined);
  assert.equal(second.doc.body.blocks.filter((block) => block.type === 'paragraph').length, 2);
});
