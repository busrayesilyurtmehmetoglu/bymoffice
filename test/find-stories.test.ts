// Find and Replace beyond the body: the text boxes (DrawingML in alternate content with its VML
// fallback, VML alone, in a table, in a text box, in a header), headers, footers, footnotes and
// endnotes, content controls. The order of the search, replacing one match and all of them as
// one undo step, refusing where changes are tracked or where the edit would not travel, and what
// the saved file holds: the changed part right (read again by the engine and by Python's
// ElementTree), every other part byte for byte as it was.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { join as joinPath } from 'node:path';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type DocModel } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import {
  DocHistory,
  DocMatchScan,
  DocSearch,
  TRACK_UNSUPPORTED,
  compileDocFind,
  docParagraphs,
  findAllInDocument,
  findInDocument,
  insertTextInTextBox,
  replaceInTextBox,
  replaceMatches,
  resolvePosition,
  revisionsOf,
  type DocEdit,
  type DocMatch,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { SECT, STYLES, p, r, tbl, tc, tr } from './doc-ops-fixtures';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const where = (match: DocMatch | undefined): string => (match ? `${match.range.anchor.story}:${match.range.anchor.path.join('.')}@${match.range.anchor.offset}-${match.range.focus.offset}` : 'none');

/** Every match a search reaches from the start, in order, until it comes round again. */
function walk(doc: DocModel, text: string, options: Parameters<typeof findInDocument>[3] = {}, backward = false): string[] {
  const out: string[] = [];
  let from: DocPosition | undefined;
  for (let i = 0; i < 200; i++) {
    const found = findInDocument(doc, text, from, options, backward);
    if (!found) break;
    const label = where(found);
    if (out.indexOf(label) >= 0) break;
    out.push(label);
    from = backward ? found.range.anchor : found.range.focus;
  }
  return out;
}

const textAt = (doc: DocModel, story: string, path: number[]): string => paragraphText(resolvePosition(doc, { story, path, offset: 0 }).paragraph);

// ------------------------------------------------------------------- fixtures

const ANCHOR_HEAD =
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/>';

/** A DrawingML text box (wps) without alternate content. */
const wpsBox = (inner: string, id = 1) =>
  `<w:r><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** A text box the way Word writes it: the DrawingML shape in a Choice, the same text in a VML Fallback. */
const wordBox = (inner: string, id = 1) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

/** A VML text box alone. */
const vmlBox = (inner: string) => `<w:r><w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;

const HEADER_SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

/** Every kind of story with "cat" in it. */
const everyStory = (): DocxSpec => ({
  body:
    p('cat in the body') +
    p(r('anchor ') + wordBox(p('cat in a box') + p('another cat'))) +
    tbl(tr(tc(p(r('cell ') + vmlBox(p('cat in a cell box'))))), [3000]) +
    p(r('ref') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r(' end') + '<w:r><w:endnoteReference w:id="3"/></w:r>') +
    HEADER_SECT,
  styles: STYLES,
  headers: { rIdH1: p(r('head cat ') + wpsBox(p('cat in a header box'), 7)) },
  footers: { rIdF1: p('foot cat') },
  footnotes: '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:t>cat in a footnote</w:t></w:r></w:p></w:footnote>',
  endnotes: '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="3"><w:p><w:r><w:t>cat in an endnote</w:t></w:r></w:p></w:endnote>',
});

interface Opened {
  bytes: Uint8Array;
  pkg: OpcPackage;
  doc: DocModel;
}

async function openSpec(spec: DocxSpec): Promise<Opened> {
  const bytes = await buildDocx(spec);
  const pkg = await OpcPackage.open(bytes);
  return { bytes, pkg, doc: await readDocx(pkg) };
}

/** Writes the model into its package, saves it and opens what was saved. */
async function saved(opened: Opened): Promise<Opened> {
  await writeDocx(opened.pkg, opened.doc);
  const bytes = await opened.pkg.save();
  const pkg = await OpcPackage.open(bytes);
  return { bytes, pkg, doc: await readDocx(pkg) };
}

/** The text of every part of a package by name. */
async function partsOf(bytes: Uint8Array): Promise<Map<string, string>> {
  const pkg = await OpcPackage.open(bytes);
  const out = new Map<string, string>();
  for (const name of pkg.partNames()) out.set(name, decoder.decode(await pkg.readPart(name)));
  return out;
}

/** The parts whose text differs between two packages. */
async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = await partsOf(before);
  const b = await partsOf(after);
  const out: string[] = [];
  const names = new Set([...a.keys(), ...b.keys()]);
  names.forEach((name) => (a.get(name) !== b.get(name) ? out.push(name) : 0));
  return out.sort();
}

let python: boolean | undefined;
/** Python's zipfile and ElementTree on the saved file (test/tools/package_check.py): what is wrong, [] when nothing is. Undefined without python3. */
function packageProblems(bytes: Uint8Array): unknown[] | undefined {
  if (python === undefined) {
    try {
      execFileSync('python3', ['-I', '-c', 'import zipfile, xml.etree.ElementTree'], { stdio: 'ignore', timeout: 60000 });
      python = true;
    } catch {
      python = false;
    }
  }
  if (!python) return undefined;
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.find-stories-'));
  try {
    writeFileSync(joinPath(dir, 'file.docx'), bytes);
    return JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 120000 })) as unknown[];
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

/** Replaces every match through a history, as one step, the way the view does. */
function replaceAll(doc: DocModel, history: DocHistory, text: string, replacement: string, options: { textBoxes?: 'replace' | 'skip'; stories?: 'all' | 'body' } = {}) {
  history.seal();
  const result = replaceMatches(doc, findAllInDocument(doc, text, { stories: options.stories }), replacement, {
    apply: (edit) => history.applyEdit({ ops: edit.ops, selection: edit.selection, mergeKey: 'replace-all' }, undefined),
    textBoxes: options.textBoxes,
  });
  history.seal();
  return result;
}

// ---------------------------------------------------------------------- order

test('find in every story: the body and its tables, the text boxes of the body, then headers, footers and notes with theirs', async () => {
  const { doc } = await openSpec(everyStory());
  const order = walk(doc, 'cat');
  assert.deepEqual(order, [
    'body:0@0-3',
    // The box of paragraph 1 (graphic 0, box 0), then the box in the cell of the table.
    'body:1.-1.0.0.0@0-3',
    'body:1.-1.0.0.1@8-11',
    'body:2.0.0.0.-1.0.0.0@0-3',
    'header:rIdH1:0@5-8',
    'header:rIdH1:0.-1.0.0.0@0-3',
    'footer:rIdF1:0@5-8',
    'footnote:2:0@0-3',
    'endnote:3:0@0-3',
  ]);
  // Backwards: the same places the other way round.
  assert.deepEqual(walk(doc, 'cat', {}, true), order.slice().reverse());
  // "Main text only": the body and its tables, no text box.
  assert.deepEqual(walk(doc, 'cat', { stories: 'body' }), ['body:0@0-3']);
  // From a place in a text box the search goes on from there.
  assert.equal(where(findInDocument(doc, 'cat', { story: 'body', path: [1, -1, 0, 0, 0], offset: 3 })), 'body:1.-1.0.0.1@8-11');
  // What is found in a text box can be replaced.
  assert.equal(findAllInDocument(doc, 'cat').every((match) => match.replaceable), true);
});

test('find: a text box in a text box comes right after the box it is in; the box of a deleted run is not searched', async () => {
  const deleted = '<w:del w:id="9" w:author="A" w:date="2026-01-01T00:00:00Z">' + vmlBox(p('cat gone')) + '</w:del>';
  const { doc } = await openSpec({ body: p(r('a ') + wpsBox(p(r('outer cat ') + vmlBox(p('inner cat'))) + p('outer two cat'))) + p(deleted) + p('last cat') + SECT, styles: STYLES });
  assert.deepEqual(walk(doc, 'cat'), ['body:2@5-8', 'body:0.-1.0.0.0@6-9', 'body:0.-1.0.0.1@10-13', 'body:0.-1.0.0.0.-1.0.0.0@6-9']);
});

// -------------------------------------------------------------------- replace

test('replace in a VML text box: one edit, undone and redone; the saved file has the new text in the box, every other part as it was', async () => {
  const opened = await openSpec({ body: p('cat outside') + p(r('anchor ') + vmlBox(p('a cat in a box'))) + SECT, styles: STYLES });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const match = findInDocument(doc, 'cat', { story: 'body', path: [0], offset: 3 })!;
  assert.equal(where(match), 'body:1.-1.0.0.0@2-5');
  const edit = insertTextInTextBox(doc, match.range, 'dog');
  assert.equal(edit.ops.length, 1);
  assert.deepEqual(edit.selection.focus, { story: 'body', path: [1, -1, 0, 0, 0], offset: 5 });
  history.applyEdit(edit, undefined);
  history.seal();
  assert.equal(textAt(doc, 'body', [1, -1, 0, 0, 0]), 'a dog in a box');
  assert.equal(textAt(doc, 'body', [1]), 'anchor ');
  history.undo();
  assert.equal(textAt(doc, 'body', [1, -1, 0, 0, 0]), 'a cat in a box');
  // Edit, undo, save: the bytes that were read.
  const back = await saved({ ...opened });
  assert.deepEqual(await changedParts(opened.bytes, back.bytes), []);
  history.redo();
  const after = await saved(opened);
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/document.xml']);
  assert.equal(textAt(after.doc, 'body', [1, -1, 0, 0, 0]), 'a dog in a box');
  assert.equal(textAt(after.doc, 'body', [0]), 'cat outside');
  assert.deepEqual(packageProblems(after.bytes) || [], []);
});

test('replace all in a text box Word writes (a Choice and a VML Fallback): both branches get the new text and the alternate content stays', async () => {
  const opened = await openSpec({ body: p(r('anchor cat ') + wordBox(p('cat one') + p('two cat cat'))) + SECT, styles: STYLES });
  const history = new DocHistory(opened.doc);
  const result = replaceAll(opened.doc, history, 'cat', 'dog');
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped }, { replaced: 4, skipped: 0 });
  assert.equal(textAt(opened.doc, 'body', [0]), 'anchor dog ');
  assert.equal(textAt(opened.doc, 'body', [0, -1, 0, 0, 0]), 'dog one');
  assert.equal(textAt(opened.doc, 'body', [0, -1, 0, 0, 1]), 'two dog dog');
  // One undo step takes back the body and the box.
  history.undo();
  assert.equal(textAt(opened.doc, 'body', [0]), 'anchor cat ');
  assert.equal(textAt(opened.doc, 'body', [0, -1, 0, 0, 1]), 'two cat cat');
  history.redo();
  const after = await saved(opened);
  const xml = (await partsOf(after.bytes)).get('/word/document.xml')!;
  assert.match(xml, /<mc:AlternateContent>/);
  const choice = xml.slice(xml.indexOf('<mc:Choice'), xml.indexOf('</mc:Choice>'));
  const fallback = xml.slice(xml.indexOf('<mc:Fallback'), xml.indexOf('</mc:Fallback>'));
  for (const branch of [choice, fallback]) {
    assert.match(branch, /dog one/);
    assert.match(branch, /two dog dog/);
    assert.doesNotMatch(branch, /cat/);
  }
  assert.match(fallback, /<v:shape style="width:144pt;height:36pt">/);
  assert.equal(textAt(after.doc, 'body', [0, -1, 0, 0, 1]), 'two dog dog');
  assert.deepEqual(packageProblems(after.bytes) || [], []);
});

test('replace all in every story at once: one undo step; the saved file changes the parts with matches and no other', async () => {
  const opened = await openSpec(everyStory());
  const { doc } = opened;
  const history = new DocHistory(doc);
  const result = replaceAll(doc, history, 'cat', 'lynx');
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped, inTextBoxes: result.inTextBoxes || 0 }, { replaced: 9, skipped: 0, inTextBoxes: 0 });
  assert.deepEqual(walk(doc, 'cat'), []);
  assert.equal(walk(doc, 'lynx').length, 9);
  history.undo();
  assert.equal(walk(doc, 'cat').length, 9);
  assert.deepEqual(await changedParts(opened.bytes, (await saved({ ...opened })).bytes), []);
  history.redo();
  const after = await saved(opened);
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/document.xml', '/word/endnotes.xml', '/word/footer1.xml', '/word/footnotes.xml', '/word/header1.xml']);
  assert.deepEqual(walk(after.doc, 'cat'), []);
  assert.equal(textAt(after.doc, 'header:rIdH1', [0, -1, 0, 0, 0]), 'lynx in a header box');
  assert.equal(textAt(after.doc, 'body', [2, 0, 0, 0, -1, 0, 0, 0]), 'lynx in a cell box');
  assert.deepEqual(packageProblems(after.bytes) || [], []);
});

test('replace only in a header: the body and the notes stay byte for byte', async () => {
  const opened = await openSpec(everyStory());
  const history = new DocHistory(opened.doc);
  const result = replaceAll(opened.doc, history, 'head cat', 'head dog');
  assert.equal(result.replaced, 1);
  const after = await saved(opened);
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/header1.xml']);
  assert.equal(textAt(after.doc, 'header:rIdH1', [0]), 'head dog ');
  assert.equal(textAt(after.doc, 'header:rIdH1', [0, -1, 0, 0, 0]), 'cat in a header box');
});

test('replace in a text box in a text box: the outer box takes the inner one back; one undo step', async () => {
  const opened = await openSpec({ body: p(r('a ') + wpsBox(p(r('outer ') + vmlBox(p('inner cat'))))) + SECT, styles: STYLES });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const match = findInDocument(doc, 'cat')!;
  assert.equal(where(match), 'body:0.-1.0.0.0.-1.0.0.0@6-9');
  const edit = insertTextInTextBox(doc, match.range, 'owl');
  history.applyEdit(edit, undefined);
  history.seal();
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0, -1, 0, 0, 0]), 'inner owl');
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0]), 'outer ');
  history.undo();
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0, -1, 0, 0, 0]), 'inner cat');
  history.redo();
  const after = await saved(opened);
  assert.equal(textAt(after.doc, 'body', [0, -1, 0, 0, 0, -1, 0, 0, 0]), 'inner owl');
  assert.deepEqual(packageProblems(after.bytes) || [], []);
});

test('replace in a text box treats the runs as a replacement in the body does', async () => {
  const runs = r('bold ', '<w:b/>') + r('cat', '<w:i/>') + r(' tail');
  const opened = await openSpec({ body: p(runs) + p(r('a ') + vmlBox(p(runs))) + SECT, styles: STYLES });
  const history = new DocHistory(opened.doc);
  assert.equal(replaceAll(opened.doc, history, 'cat', 'dog').replaced, 2);
  const after = await saved(opened);
  const xml = (await partsOf(after.bytes)).get('/word/document.xml')!;
  const body = xml.slice(xml.indexOf('<w:body>') + 8, xml.indexOf('</w:p>') + 6);
  const box = xml.slice(xml.indexOf('<w:txbxContent>') + 15, xml.indexOf('</w:txbxContent>'));
  assert.equal(box, body);
  assert.match(box, /dog/);
  assert.match(box, /<w:b\/>/);
});

test('where changes are tracked: the body is replaced as tracked changes, a text box is not replaced and is counted', async () => {
  const opened = await openSpec({ body: p(r('anchor cat ') + vmlBox(p('box cat'))) + SECT, styles: STYLES, settings: '<w:trackRevisions/>' });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const box = findAllInDocument(doc, 'cat')[1];
  assert.throws(() => insertTextInTextBox(doc, box.range, 'dog'), (error: Error) => error instanceof RangeError && error.message === TRACK_UNSUPPORTED);
  const result = replaceAll(doc, history, 'cat', 'dog');
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped, inTextBoxes: result.inTextBoxes || 0 }, { replaced: 1, skipped: 1, inTextBoxes: 1 });
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0]), 'box cat');
  assert.deepEqual(
    revisionsOf(doc).map((revision) => revision.kind).sort(),
    ['delete', 'insert'],
  );
});

test('where changes are tracked in every story: headers, footers and notes take tracked changes, the four text boxes are counted; saved and read again', async () => {
  const opened = await openSpec({ ...everyStory(), settings: '<w:trackRevisions/>' });
  const history = new DocHistory(opened.doc);
  const result = replaceAll(opened.doc, history, 'cat', 'dog');
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped, inTextBoxes: result.inTextBoxes || 0 }, { replaced: 5, skipped: 4, inTextBoxes: 4 });
  const after = await saved(opened);
  assert.equal(textAt(after.doc, 'header:rIdH1', [0]), 'head dog ');
  assert.equal(textAt(after.doc, 'footnote:2', [0]), 'dog in a footnote');
  assert.equal(textAt(after.doc, 'body', [1, -1, 0, 0, 0]), 'cat in a box');
  assert.equal(revisionsOf(after.doc).filter((revision) => revision.kind === 'insert').length, 5);
  assert.deepEqual(packageProblems(after.bytes) || [], []);
});

test('in a room the matches of text boxes are skipped and counted; the others are replaced', async () => {
  const { doc } = await openSpec(everyStory());
  const history = new DocHistory(doc);
  const result = replaceAll(doc, history, 'cat', 'dog', { textBoxes: 'skip' });
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped, inTextBoxes: result.inTextBoxes || 0 }, { replaced: 5, skipped: 4, inTextBoxes: 4 });
  assert.equal(textAt(doc, 'body', [1, -1, 0, 0, 0]), 'cat in a box');
  assert.equal(textAt(doc, 'footnote:2', [0]), 'dog in a footnote');
});

test('in a room the limit counts the replacements that are made, not the matches of text boxes that are skipped', async () => {
  const { doc } = await openSpec(everyStory());
  const history = new DocHistory(doc);
  const matches = findAllInDocument(doc, 'cat');
  const apply = (edit: DocEdit) => history.applyEdit(edit, undefined);
  assert.equal(replaceMatches(doc, matches, 'dog', { apply, limit: 4, textBoxes: 'skip' }).refused, true);
  const done = replaceMatches(doc, matches, 'dog', { apply, limit: 5, textBoxes: 'skip' });
  assert.deepEqual([done.replaced, done.skipped, done.inTextBoxes], [5, 4, 4]);
});

test('a text box anchored in a run inside alternate content is not replaced (its other branch would be lost); nothing changes', async () => {
  const run = `<mc:AlternateContent><mc:Choice Requires="w14">${vmlBox(p('cat here'))}</mc:Choice><mc:Fallback>${r('fallback')}</mc:Fallback></mc:AlternateContent>`;
  const { doc } = await openSpec({ body: p(r('a ') + run) + SECT, styles: STYLES });
  const matches = findAllInDocument(doc, 'cat');
  assert.equal(matches.length, 1);
  assert.throws(() => insertTextInTextBox(doc, matches[0].range, 'dog'), RangeError);
  const history = new DocHistory(doc);
  assert.deepEqual(replaceAll(doc, history, 'cat', 'dog').replaced, 0);
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0]), 'cat here');
});

test('replaceInTextBox: the matches under one paragraph in one edit, in all its boxes; ranges under two paragraphs are refused', async () => {
  const { doc } = await openSpec({ body: p(r('a ') + vmlBox(p('cat cat') + p('cat')) + vmlBox(p('cat'))) + p(r('b ') + vmlBox(p('cat'))) + SECT, styles: STYLES });
  const matches = findAllInDocument(doc, 'cat');
  assert.equal(matches.length, 5);
  assert.throws(() => replaceInTextBox(doc, [matches[4].range, matches[0].range], 'x'), RangeError);
  const own = matches.slice(0, 4).reverse().map((match) => match.range);
  const result = replaceInTextBox(doc, own, 'dog');
  assert.equal(result.replaced, 4);
  assert.equal(result.failed, 0);
  assert.equal(result.edit!.ops.length, 1);
  // The selection of the replacement of the first match.
  assert.deepEqual(result.edit!.selection.focus, { story: 'body', path: [0, -1, 0, 0, 0], offset: 3 });
  const history = new DocHistory(doc);
  history.applyEdit(result.edit!, undefined);
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 0]), 'dog dog');
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 1]), 'dog');
  assert.equal(textAt(doc, 'body', [0, -1, 1, 0, 0]), 'dog');
  assert.equal(textAt(doc, 'body', [1, -1, 0, 0, 0]), 'cat');
});

test('content controls: text in one is found and replaced; in one whose content is locked it is found and not replaced', async () => {
  const control = (lock: string, text: string) => `<w:sdt><w:sdtPr>${lock}<w:id w:val="5"/></w:sdtPr><w:sdtContent>${r(text)}</w:sdtContent></w:sdt>`;
  const block = `<w:sdt><w:sdtPr><w:lock w:val="sdtContentLocked"/></w:sdtPr><w:sdtContent>${p('cat in a locked block')}</w:sdtContent></w:sdt>`;
  const { doc } = await openSpec({ body: p(control('', 'cat free')) + p(control('<w:lock w:val="contentLocked"/>', 'cat fixed')) + p(control('<w:lock w:val="sdtLocked"/>', 'cat movable')) + block + SECT, styles: STYLES });
  const matches = findAllInDocument(doc, 'cat');
  assert.deepEqual(
    matches.map((match) => match.replaceable),
    [true, false, true, false],
  );
  const history = new DocHistory(doc);
  const result = replaceAll(doc, history, 'cat', 'dog');
  assert.deepEqual({ replaced: result.replaced, skipped: result.skipped }, { replaced: 2, skipped: 2 });
  assert.equal(textAt(doc, 'body', [1]), 'cat fixed');
});

// ----------------------------------------------------------------------- speed

test('speed: 10,000 paragraphs, 500 text boxes, a header and a footer; one match at the end, then all of them, in slices', async () => {
  let body = '';
  const line = 'Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor. ';
  for (let i = 0; i < 10000; i++) body += i % 20 === 0 ? p(r(line) + vmlBox(p(`box ${i} text`) + p('more box text'))) : p(line + line);
  body += p('needle at the end') + HEADER_SECT;
  const t0 = Date.now();
  const { doc } = await openSpec({ body, styles: STYLES, headers: { rIdH1: p(r('head ') + vmlBox(p('head box'))) }, footers: { rIdF1: p('foot') } });
  const opened = Date.now() - t0;
  let t = Date.now();
  const list = docParagraphs(doc);
  const listing = Date.now() - t;
  assert.equal(list.length, 10001 + 1000 + 2 + 1);
  t = Date.now();
  const search = new DocSearch(compileDocFind('needle')!, list, undefined);
  let found: DocMatch | null | undefined;
  let longest = 0;
  for (;;) {
    const s = Date.now();
    found = search.step(256);
    longest = Math.max(longest, Date.now() - s);
    if (found !== undefined) break;
  }
  const finding = Date.now() - t;
  assert.equal(where(found!), 'body:10000@0-6');
  t = Date.now();
  const scan = new DocMatchScan(compileDocFind('box')!, list);
  while (!scan.step(256));
  const scanning = Date.now() - t;
  assert.equal(scan.matches.length, 1001);
  t = Date.now();
  const history = new DocHistory(doc);
  const result = replaceMatches(doc, scan.matches, 'BOX', { apply: (edit) => history.applyEdit({ ops: edit.ops, selection: edit.selection, mergeKey: 'all' }, undefined) });
  history.seal();
  const replacing = Date.now() - t;
  assert.equal(result.replaced, 1001);
  t = Date.now();
  history.undo();
  const undoing = Date.now() - t;
  assert.equal(textAt(doc, 'body', [0, -1, 0, 0, 1]), 'more box text');
  console.log(`  open ${opened} ms, list ${listing} ms, find ${finding} ms (longest slice ${longest} ms), scan ${scanning} ms, replace all ${replacing} ms, undo ${undoing} ms`);
  // Guards against something that takes far too long, not measures of speed.
  assert.ok(longest < 1000, String(longest));
  assert.ok(replacing < 30000, String(replacing));
});

test('speed with hostile input: one text box of 20,000 matching paragraphs, one paragraph with 5,000 text boxes', async () => {
  let inner = '';
  for (let i = 0; i < 20000; i++) inner += p('cat');
  const big = await openSpec({ body: p(r('a ') + vmlBox(inner)) + SECT, styles: STYLES });
  let t = Date.now();
  const all = findAllInDocument(big.doc, 'cat');
  assert.equal(all.length, 20000);
  const history = new DocHistory(big.doc);
  const result = replaceMatches(big.doc, all, 'dog', { apply: (edit) => history.applyEdit(edit, undefined) });
  const oneBox = Date.now() - t;
  assert.equal(result.replaced, 20000);
  assert.equal(textAt(big.doc, 'body', [0, -1, 0, 0, 19999]), 'dog');

  let boxes = '';
  for (let i = 0; i < 5000; i++) boxes += vmlBox(p('cat'));
  const many = await openSpec({ body: p(r('a ') + boxes) + SECT, styles: STYLES });
  t = Date.now();
  const found = findAllInDocument(many.doc, 'cat');
  const listing = Date.now() - t;
  assert.equal(found.length, 5000);
  t = Date.now();
  const history2 = new DocHistory(many.doc);
  const result2 = replaceMatches(many.doc, found, 'dog', { apply: (edit) => history2.applyEdit(edit, undefined) });
  const manyBoxes = Date.now() - t;
  assert.equal(result2.replaced, 5000);
  console.log(`  one box of 20,000 matches: ${oneBox} ms; 5,000 boxes in a paragraph: find ${listing} ms, replace all ${manyBoxes} ms`);
  assert.ok(oneBox < 60000 && manyBoxes < 60000, `${oneBox} ${manyBoxes}`);
});

test('a place in a text box: the search starts there; a range over two text boxes is refused', async () => {
  const { doc } = await openSpec({ body: p(r('a ') + vmlBox(p('x cat')) + vmlBox(p('y cat'))) + SECT, styles: STYLES });
  const second: DocRange = { anchor: { story: 'body', path: [0, -1, 0, 0, 0], offset: 0 }, focus: { story: 'body', path: [0, -1, 1, 0, 0], offset: 1 } };
  assert.throws(() => insertTextInTextBox(doc, second, 'z'), RangeError);
  assert.equal(where(findInDocument(doc, 'cat', { story: 'body', path: [0, -1, 1, 0, 0], offset: 0 })), 'body:0.-1.1.0.0@2-5');
});
