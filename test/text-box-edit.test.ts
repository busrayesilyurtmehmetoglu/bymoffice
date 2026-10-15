// Editing in text boxes (editInTextBox): typing, deleting, Enter, formatting and undo in a
// DrawingML box, a box Word writes (a Choice and its VML Fallback), a VML box, a box in a table
// cell, in a header and in a box; what the saved file holds (the changed part right, both
// branches the same, every other part as it was, edit + undo + save gives the bytes that were
// read); and where a box is not edited: changes tracked, a locked shape or control, a box whose
// other branch could not take the change, a command that would change more than the box.

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
  TEXT_BOX_OUTSIDE,
  TRACK_UNSUPPORTED,
  deleteBackward,
  deleteForward,
  deleteRange,
  editInTextBox,
  expandRangeInTextBox,
  findAllInDocument,
  insertRow,
  insertText,
  movePosition,
  rangeFormat,
  readInTextBox,
  resolvePosition,
  setParaFormat,
  splitParagraph,
  textBoxExit,
  textBoxProblem,
  toggleList,
  toggleRunFormat,
  type BoxCommand,
  type DocEdit,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import { setTracking } from '../src/ops/doc/track';
import { READ_ONLY } from '../src/ops/doc/pieces';
import { tracing, wiresOf } from '../src/collab/doc/trace';
import * as shared from '../src/collab/doc/commands';
import { NUMBERING, SECT, STYLES, p, r, tbl, tc, tr } from './doc-ops-fixtures';
import { buildDocx, type DocxSpec } from './docx-fixtures';
import { fakeMeasurer } from './doc-layout-fixtures';
import { cutLoses, sliceOf } from '../src/ui/doc-clipboard';
import { originOf } from '../src/ops/doc/fragment';
import { takenByDocument } from '../src/ops/doc/structure';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutPage, LineMark, TextItem } from '../src/layout/doc/types';

const decoder = new TextDecoder();

// ------------------------------------------------------------------- fixtures

const ANCHOR_HEAD =
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="457200"/><wp:wrapNone/>';

/** A DrawingML text box (wps) without alternate content; `locks`: the shape's own locks. */
const wpsBox = (inner: string, id = 1, locks = '') =>
  `<w:r><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>${locks ? `<wps:cNvSpPr txBox="1">${locks}</wps:cNvSpPr>` : ''}<wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

/** A text box the way Word writes it: the DrawingML shape in a Choice, the same text in a VML Fallback (or another one). */
const wordBox = (inner: string, id = 1, fallback = inner) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${ANCHOR_HEAD}<wp:docPr id="${id}" name="Box ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>` +
  `<mc:Fallback><w:pict><v:shape style="width:144pt;height:36pt"><v:textbox><w:txbxContent>${fallback}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

/** A VML text box alone; `lock`: an o:lock element in the shape. */
const vmlBox = (inner: string, lock = '') => `<w:r><w:pict><v:shape style="width:144pt;height:36pt">${lock}<v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>`;

const HEADER_SECT = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';

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

async function saved(opened: Opened): Promise<Opened> {
  await writeDocx(opened.pkg, opened.doc);
  const bytes = await opened.pkg.save();
  const pkg = await OpcPackage.open(bytes);
  return { bytes, pkg, doc: await readDocx(pkg) };
}

async function partsOf(bytes: Uint8Array): Promise<Map<string, string>> {
  const pkg = await OpcPackage.open(bytes);
  const out = new Map<string, string>();
  for (const name of pkg.partNames()) out.set(name, decoder.decode(await pkg.readPart(name)));
  return out;
}

async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = await partsOf(before);
  const b = await partsOf(after);
  const out: string[] = [];
  new Set([...a.keys(), ...b.keys()]).forEach((name) => (a.get(name) !== b.get(name) ? out.push(name) : 0));
  return out.sort();
}

let python: boolean | undefined;
/** Python's zipfile and ElementTree on the saved file (test/tools/package_check.py): what is wrong, [] when nothing is or without python3. */
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
  const dir = mkdtempSync(joinPath(process.cwd(), 'test', '.text-box-edit-'));
  try {
    writeFileSync(joinPath(dir, 'file.docx'), bytes);
    return JSON.parse(execFileSync('python3', ['-I', joinPath(process.cwd(), 'test', 'tools', 'package_check.py'), dir], { encoding: 'utf8', timeout: 120000 })) as unknown[];
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

const at = (path: number[], offset: number, story = 'body'): DocPosition => ({ story, path, offset });
const span = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });
const textAt = (doc: DocModel, path: number[], story = 'body'): string => paragraphText(resolvePosition(doc, at(path, 0, story)).paragraph);

/** Runs a command in the box (as the view does) and applies it through the history; the selection it leaves. */
function step(doc: DocModel, history: DocHistory, where: DocPosition | DocRange, command: BoxCommand): DocRange {
  const range = 'anchor' in where ? where : span(where, where);
  const edit = editInTextBox(doc, range, command);
  history.applyEdit(edit, range);
  return edit.selection;
}

/** Types text one character at a time, as keys do. */
function type(doc: DocModel, history: DocHistory, from: DocPosition, text: string): DocPosition {
  let caret = from;
  for (const char of text) caret = step(doc, history, caret, (d, range) => insertText(d, range, char)).focus;
  return caret;
}

const isError = (message: string) => (error: unknown) => error instanceof RangeError && error.message === message;

// ---------------------------------------------------------------------- editing

test('typing, Backspace, Delete and Enter in a DrawingML text box: the box changes, the paragraph that anchors it keeps its text', async () => {
  const opened = await openSpec(p('before') + p(r('anchor ') + wpsBox(p('hello') + p('world'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const box = [1, -1, 0, 0];
  let caret = type(doc, history, at(box.concat([0]), 5), ' there');
  assert.deepEqual(caret, at(box.concat([0]), 11));
  assert.equal(textAt(doc, box.concat([0])), 'hello there');
  assert.equal(textAt(doc, [1]), 'anchor ');
  caret = step(doc, history, caret, (d, range) => deleteBackward(d, range.focus)).focus;
  assert.equal(textAt(doc, box.concat([0])), 'hello ther');
  caret = step(doc, history, at(box.concat([0]), 0), (d, range) => deleteForward(d, range.focus)).focus;
  assert.equal(textAt(doc, box.concat([0])), 'ello ther');
  // Enter: a new paragraph in the box.
  caret = step(doc, history, at(box.concat([0]), 4), (d, range) => splitParagraph(d, range.focus)).focus;
  assert.deepEqual(caret, at(box.concat([1]), 0));
  assert.deepEqual([textAt(doc, box.concat([0])), textAt(doc, box.concat([1])), textAt(doc, box.concat([2]))], ['ello', ' ther', 'world']);
  // Backspace at the start of a paragraph of the box joins it with the one before.
  caret = step(doc, history, caret, (d, range) => deleteBackward(d, range.focus)).focus;
  assert.deepEqual(caret, at(box.concat([0]), 4));
  assert.equal(textAt(doc, box.concat([0])), 'ello ther');
  // Backspace at the very start of the box: nothing to join, nothing changes, the caret stays in the box.
  const edit = editInTextBox(doc, span(at(box.concat([0]), 0), at(box.concat([0]), 0)), (d, range) => deleteBackward(d, range.focus));
  assert.equal(edit.ops.length, 0);
  assert.deepEqual(edit.selection.focus, at(box.concat([0]), 0));
  // A selection over two paragraphs of the box, deleted.
  step(doc, history, span(at(box.concat([0]), 2), at(box.concat([1]), 3)), (d, range) => deleteRange(d, range));
  assert.equal(textAt(doc, box.concat([0])), 'elld');
  assert.equal(textAt(doc, [0]), 'before');
  assert.equal(textAt(doc, [1]), 'anchor ');
  const after = await saved(opened);
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/document.xml']);
  assert.equal(textAt(after.doc, box.concat([0])), 'elld');
  assert.equal(textAt(after.doc, [1]), 'anchor ');
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('typed characters are one step of the history; edit, undo and save gives the bytes that were read', async () => {
  const opened = await openSpec(p(r('anchor ') + wpsBox(p('one'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  type(doc, history, at([0, -1, 0, 0, 0], 3), '-two-three');
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'one-two-three');
  const undone = history.undo();
  assert.ok(undone);
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'one');
  assert.equal(history.canUndo, false);
  const back = await saved({ ...opened });
  assert.deepEqual(back.bytes, opened.bytes);
  history.redo();
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'one-two-three');
});

test('typing in two boxes is two steps, and the step in a box is not joined with typing in the text', async () => {
  const opened = await openSpec(p(r('a ') + wpsBox(p('x'), 1) + wpsBox(p('y'), 2)));
  const { doc } = opened;
  const history = new DocHistory(doc);
  type(doc, history, at([0], 2), 'B');
  type(doc, history, at([0, -1, 0, 0, 0], 1), '1');
  type(doc, history, at([0, -1, 1, 0, 0], 1), '2');
  assert.deepEqual([textAt(doc, [0]), textAt(doc, [0, -1, 0, 0, 0]), textAt(doc, [0, -1, 1, 0, 0])], ['a B', 'x1', 'y2']);
  history.undo();
  assert.deepEqual([textAt(doc, [0]), textAt(doc, [0, -1, 0, 0, 0]), textAt(doc, [0, -1, 1, 0, 0])], ['a B', 'x1', 'y']);
  history.undo();
  assert.deepEqual([textAt(doc, [0, -1, 0, 0, 0]), textAt(doc, [0, -1, 1, 0, 0])], ['x', 'y']);
  assert.equal(textAt(doc, [0]), 'a B');
});

test('bold, italic and alignment in a box Word writes: both branches of its alternate content take them, and read back', async () => {
  const opened = await openSpec(p(r('anchor ') + wordBox(p('make this bold') + p('second'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const box = [0, -1, 0, 0];
  step(doc, history, span(at(box.concat([0]), 10), at(box.concat([0]), 14)), (d, range) => toggleRunFormat(d, range, 'bold'));
  step(doc, history, span(at(box.concat([0]), 0), at(box.concat([0]), 4)), (d, range) => toggleRunFormat(d, range, 'italic'));
  step(doc, history, at(box.concat([1]), 2), (d, range) => setParaFormat(d, range, { align: 'center' }));
  // The format of a selection in the box is read as anywhere else.
  assert.equal(rangeFormat(doc, span(at(box.concat([0]), 10), at(box.concat([0]), 14))).bold, true);
  assert.equal(rangeFormat(doc, at(box.concat([1]), 1)).align, 'center');
  const after = await saved(opened);
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/document.xml']);
  const xml = (await partsOf(after.bytes)).get('/word/document.xml')!;
  const choice = xml.slice(xml.indexOf('<mc:Choice'), xml.indexOf('</mc:Choice>'));
  const fallback = xml.slice(xml.indexOf('<mc:Fallback'), xml.indexOf('</mc:Fallback>'));
  for (const branch of [choice, fallback]) {
    assert.match(branch, /<w:b\/>[^]*bold/);
    assert.match(branch, /<w:i\/>[^]*make/);
    assert.match(branch, /<w:jc w:val="center"\/>/);
  }
  assert.match(fallback, /<v:shape style="width:144pt;height:36pt">/);
  assert.equal(rangeFormat(after.doc, span(at(box.concat([0]), 10), at(box.concat([0]), 14))).bold, true);
  assert.equal(textAt(after.doc, box.concat([0])), 'make this bold');
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('Enter and typing in a box Word writes: the new paragraph is in both branches', async () => {
  const opened = await openSpec(p(r('anchor ') + wordBox(p('first line'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const caret = step(doc, history, at([0, -1, 0, 0, 0], 10), (d, range) => splitParagraph(d, range.focus)).focus;
  type(doc, history, caret, 'second line');
  const after = await saved(opened);
  const xml = (await partsOf(after.bytes)).get('/word/document.xml')!;
  const choice = xml.slice(xml.indexOf('<mc:Choice'), xml.indexOf('</mc:Choice>'));
  const fallback = xml.slice(xml.indexOf('<mc:Fallback'), xml.indexOf('</mc:Fallback>'));
  for (const branch of [choice, fallback]) {
    assert.equal(branch.split('<w:p>').length + branch.split('<w:p ').length - 2, 2, branch);
    assert.match(branch, /first line[^]*second line/);
  }
  assert.deepEqual([textAt(after.doc, [0, -1, 0, 0, 0]), textAt(after.doc, [0, -1, 0, 0, 1])], ['first line', 'second line']);
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('a VML text box alone, one in a table cell and one in a header are edited; only the part that holds the box changes', async () => {
  const opened = await openSpec({
    body: p('text') + p(r('anchor ') + vmlBox(p('vml box'))) + tbl(tr(tc(p(r('cell ') + wpsBox(p('cell box'))))), [3000]) + HEADER_SECT,
    styles: STYLES,
    headers: { rIdH1: p(r('head ') + wpsBox(p('header box'), 7)) },
  });
  const { doc } = opened;
  const history = new DocHistory(doc);
  type(doc, history, at([1, -1, 0, 0, 0], 7), '!');
  type(doc, history, at([2, 0, 0, 0, -1, 0, 0, 0], 0), '>');
  const after = await saved({ ...opened });
  assert.deepEqual(await changedParts(opened.bytes, after.bytes), ['/word/document.xml']);
  assert.equal(textAt(after.doc, [1, -1, 0, 0, 0]), 'vml box!');
  assert.equal(textAt(after.doc, [2, 0, 0, 0, -1, 0, 0, 0]), '>cell box');
  assert.equal(textAt(after.doc, [2, 0, 0, 0]), 'cell ');
  assert.deepEqual(packageProblems(after.bytes), []);

  // In the header: only the header part changes.
  const second = await openSpec({
    body: p('text') + HEADER_SECT,
    styles: STYLES,
    headers: { rIdH1: p(r('head ') + wpsBox(p('header box'), 7)) },
  });
  const h = new DocHistory(second.doc);
  type(second.doc, h, at([0, -1, 0, 0, 0], 6, 'header:rIdH1'), 'X');
  assert.equal(textAt(second.doc, [0, -1, 0, 0, 0], 'header:rIdH1'), 'headerX box');
  const saved2 = await saved(second);
  assert.deepEqual(await changedParts(second.bytes, saved2.bytes), ['/word/header1.xml']);
  assert.equal(textAt(saved2.doc, [0, -1, 0, 0, 0], 'header:rIdH1'), 'headerX box');
  assert.deepEqual(packageProblems(saved2.bytes), []);
});

test('a text box in a text box: the inner box is edited, the outer one takes it back; one undo step each', async () => {
  const opened = await openSpec(p(r('a ') + wpsBox(p(r('outer ') + vmlBox(p('inner'))) + p('outer two'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const inner = [0, -1, 0, 0, 0, -1, 0, 0, 0];
  let caret = type(doc, history, at(inner, 5), '-text');
  assert.deepEqual(caret, at(inner, 10));
  caret = step(doc, history, caret, (d, range) => splitParagraph(d, range.focus)).focus;
  assert.deepEqual(caret, at([0, -1, 0, 0, 0, -1, 0, 0, 1], 0));
  assert.equal(textAt(doc, inner), 'inner-text');
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'outer ');
  assert.equal(textAt(doc, [0, -1, 0, 0, 1]), 'outer two');
  // Moving stays in the inner box.
  assert.deepEqual(movePosition(doc, caret, 'char', 1), caret);
  assert.deepEqual(movePosition(doc, at(inner, 0), 'char', -1), at(inner, 0));
  history.undo();
  assert.equal(textAt(doc, inner), 'inner-text');
  const after = await saved({ ...opened });
  assert.equal(textAt(after.doc, inner), 'inner-text');
  assert.equal(textAt(after.doc, [0, -1, 0, 0, 1]), 'outer two');
  assert.deepEqual(packageProblems(after.bytes), []);
  history.undo();
  const back = await saved({ ...opened });
  assert.deepEqual(back.bytes, opened.bytes);
});

test('a table in a text box: its cells are typed in as cells are anywhere', async () => {
  const opened = await openSpec(p(r('a ') + wpsBox(tbl(tr(tc('one') + tc('two')), [1000, 1000]) + p('after'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  type(doc, history, at([0, -1, 0, 0, 0, 0, 1, 0], 3), '!');
  assert.equal(textAt(doc, [0, -1, 0, 0, 0, 0, 1, 0]), 'two!');
  const after = await saved(opened);
  assert.equal(textAt(after.doc, [0, -1, 0, 0, 0, 0, 1, 0]), 'two!');
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('a row added to a table in a text box, and Tab past its last cell, as anywhere', async () => {
  const opened = await openSpec(p(r('a ') + wpsBox(tbl(tr(tc('one') + tc('two')), [1000, 1000]) + p('after'))));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const caret = step(doc, history, at([0, -1, 0, 0, 0, 0, 1, 0], 1), (d, range) => insertRow(d, range.focus, 'below')).focus;
  const table = resolvePosition(doc, at([0, -1, 0, 0, 0, 1, 0, 0], 0));
  assert.equal(table.length, 0);
  assert.ok(caret.path.length >= 8);
  const after = await saved(opened);
  assert.equal(textAt(after.doc, [0, -1, 0, 0, 0, 0, 1, 0]), 'two');
  assert.equal(textAt(after.doc, [0, -1, 0, 0, 0, 1, 1, 0]), '');
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('a selection that goes into a table of a text box covers whole cells, as in the body: the copy holds what a cut deletes', async () => {
  const T = tbl(tr(tc('c1') + tc('c2')) + tr(tc('c3') + tc('c4')), [1000, 1000]);
  const { doc } = await openSpec(p(r('anchor ') + wpsBox(p('head') + T + p('tail'))));
  const selection = span(at([0, -1, 0, 0, 0], 2), at([0, -1, 0, 0, 1, 0, 1, 0], 1));
  const shown = expandRangeInTextBox(doc, selection);
  // As expandRangeForTables() widens the same range in the body: to the end of the cell c2.
  assert.deepEqual(shown.focus, at([0, -1, 0, 0, 1, 0, 1, 0], 2));
  assert.deepEqual(shown.anchor, selection.anchor);
  const slice = readInTextBox(doc, shown, (d, lowered) => sliceOf(d, lowered.anchor, lowered.focus));
  assert.deepEqual(
    slice!.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')),
    ['ad', 'c1', 'c2'],
  );
  // The cut: what goes is what the copy holds, and pasting it back gives the text again.
  assert.equal(readInTextBox(doc, shown, (d, lowered) => cutLoses(d, lowered, slice)), false);
  const history = new DocHistory(doc);
  const caret = step(doc, history, selection, (d, range) => deleteRange(d, range)).focus;
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'he');
  const text = slice!.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n');
  step(doc, history, caret, (d, range) => insertText(d, range, text));
  const all: string[] = [];
  for (let i = 0; i < 8; i++) {
    try {
      all.push(textAt(doc, [0, -1, 0, 0, i]));
    } catch {
      break;
    }
  }
  assert.ok(all.join('|').indexOf('c2') >= 0, all.join('|'));
  // Outside text boxes it is expandRangeForTables() itself.
  assert.deepEqual(expandRangeInTextBox(doc, at([0], 1)), span(at([0], 1), at([0], 1)));
});

test('copy and cut in a text box: the copy is read in the box, and a cut of its text loses nothing', async () => {
  const { doc } = await openSpec(p(r('a ') + wpsBox(p(r('bold', '<w:b/>') + r(' plain')) + p('second'))));
  const range = span(at([0, -1, 0, 0, 0], 0), at([0, -1, 0, 0, 1], 3));
  const slice = readInTextBox(doc, range, (d, lowered) => sliceOf(d, lowered.anchor, lowered.focus));
  assert.ok(slice);
  assert.deepEqual(
    slice!.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')),
    ['bold plain', 'sec'],
  );
  assert.equal(slice!.paragraphs[0].runs[0].format && slice!.paragraphs[0].runs[0].format.bold, true);
  assert.equal(readInTextBox(doc, range, (d, lowered) => cutLoses(d, lowered, slice)), false);
  // Nothing of the document changed by reading.
  assert.equal(textAt(doc, [0, -1, 0, 0, 0]), 'bold plain');
  assert.throws(() => readInTextBox(doc, span(at([0], 0), at([0, -1, 0, 0, 0], 1)), () => 0), RangeError);
  // Outside text boxes it is the reader itself.
  assert.equal(readInTextBox(doc, at([0], 1), (d, lowered) => d === doc && lowered.focus.offset), 1);
});

test('ids of revisions a command in a box gives are new to the whole document, not to the box alone', async () => {
  const boxInner = '<w:p><w:ins w:id="3" w:author="Other" w:date="2020-01-01T00:00:00Z"><w:r><w:t>inserted text</w:t></w:r></w:ins></w:p>';
  const opened = await openSpec(p('body text') + p(r('anchor ') + wpsBox(boxInner)));
  const { doc } = opened;
  const history = new DocHistory(doc);
  // Tracked typing in the body takes the next id (4).
  setTracking(doc, { force: true, author: 'Me' });
  const c0 = at([0], 4);
  history.applyEdit(insertText(doc, c0, 'X'), span(c0, c0));
  setTracking(doc, { force: false });
  // Enter inside the insertion in the box: the second half gets an id of its own, which must not be 4 again.
  step(doc, history, at([1, -1, 0, 0, 0], 4), (d, range) => splitParagraph(d, range.focus));
  const after = await saved(opened);
  const xml = (await partsOf(after.bytes)).get('/word/document.xml')!;
  const ids = [...xml.matchAll(/<w:(?:ins|del)\b[^>]*w:id="(\d+)"/g)].map((match) => match[1]);
  assert.equal(new Set(ids).size, ids.length, ids.join(' '));
  assert.deepEqual(packageProblems(after.bytes), []);
  // A copy taken in a box is a copy of this document, as one taken anywhere else in it.
  assert.equal(readInTextBox(doc, at([1, -1, 0, 0, 0], 0), (d) => originOf(d)), originOf(doc));
  // The ids of drawings a command in a box would give are new to the whole document too.
  const two = await openSpec(p(r('a ') + wpsBox(p('x'), 1)) + p(r('b ') + wpsBox(p('y'), 41)));
  assert.equal(readInTextBox(two.doc, at([0, -1, 0, 0, 0], 0), (d) => takenByDocument(d).drawing), 41);
});

// --------------------------------------------------------------------- refused

test('where changes are tracked a box is not edited (TRACK_UNSUPPORTED); the text around it takes tracked changes', async () => {
  const opened = await openSpec(p(r('anchor ') + wpsBox(p('box'))));
  const { doc } = opened;
  setTracking(doc, { force: true, author: 'Me' });
  assert.throws(() => editInTextBox(doc, at([0, -1, 0, 0, 0], 3), (d, range) => insertText(d, range, 'x')), isError(TRACK_UNSUPPORTED));
  const edit = editInTextBox(doc, at([0], 0), (d, range) => insertText(d, range, 'x'));
  assert.ok(edit.ops.length > 0);
  // Options asking for tracking are taken too.
  const plain = await openSpec(p(r('anchor ') + wpsBox(p('box'))));
  assert.throws(() => editInTextBox(plain.doc, at([0, -1, 0, 0, 0], 3), (d, range) => insertText(d, range, 'x'), { trackChanges: true }), isError(TRACK_UNSUPPORTED));
});

test('a shape that locks its text, a locked content control, a box the writer could not write back: refused, and why', async () => {
  const locked = [
    ['DrawingML noTextEdit', p(r('a ') + wpsBox(p('box'), 1, '<a:spLocks noTextEdit="1"/>'))],
    ['VML lock of the text', p(r('a ') + vmlBox(p('box'), '<o:lock v:ext="edit" text="t"/>'))],
    ['locked content control around the run', p('<w:sdt><w:sdtPr><w:lock w:val="contentLocked"/></w:sdtPr><w:sdtContent>' + r('a ') + wpsBox(p('box')) + '</w:sdtContent></w:sdt>')],
  ];
  for (const [what, body] of locked) {
    const { doc } = await openSpec(body);
    const place = at([0, -1, 0, 0, 0], 1);
    assert.equal(textBoxProblem(doc, place), 'locked', what);
    assert.throws(() => editInTextBox(doc, place, (d, range) => insertText(d, range, 'x')), isError(READ_ONLY), what);
    // Find and Replace follows the same rule: the match is found and kept.
    const matches = findAllInDocument(doc, 'box');
    assert.equal(matches.length, 1, what);
    assert.equal(matches[0].replaceable, false, what);
  }
  const { doc } = await openSpec(p(r('a ') + wordBox(p('box one'), 1, p('OTHER TEXT'))));
  assert.equal(textBoxProblem(doc, at([0, -1, 0, 0, 0], 1)), 'unwritable');
  assert.throws(() => editInTextBox(doc, at([0, -1, 0, 0, 0], 1), (d, range) => insertText(d, range, 'x')), isError(READ_ONLY));
  // A box that can be edited has no problem, nor a place outside boxes.
  const fine = await openSpec(p(r('a ') + wordBox(p('box one'))));
  assert.equal(textBoxProblem(fine.doc, at([0, -1, 0, 0, 0], 1)), undefined);
  assert.equal(textBoxProblem(fine.doc, at([0], 1)), undefined);
  // An o:lock that does not lock the text does not count.
  const position = await openSpec(p(r('a ') + vmlBox(p('box'), '<o:lock v:ext="edit" aspectratio="t"/>')));
  assert.equal(textBoxProblem(position.doc, at([0, -1, 0, 0, 0], 1)), undefined);
});

test('a command that would change more than the blocks of the box (a new list: the numbering) is refused and changes nothing', async () => {
  const opened = await openSpec({ body: p(r('a ') + wpsBox(p('item'))) + SECT, styles: STYLES, numbering: NUMBERING });
  const { doc } = opened;
  const before = doc.numbering;
  let refused = false;
  try {
    editInTextBox(doc, at([0, -1, 0, 0, 0], 1), (d, range) => toggleList(d, range, 'bullet'));
  } catch (error) {
    refused = error instanceof RangeError && error.message === TEXT_BOX_OUTSIDE;
  }
  // A new list needs a new numbering instance: a change outside the box, refused, and nothing changed.
  assert.equal(refused, true);
  assert.equal(doc.numbering, before);
  const after = await saved(opened);
  assert.deepEqual(packageProblems(after.bytes), []);
});

test('the ends of a range in two boxes, or one in a box and one outside, are refused', async () => {
  const { doc } = await openSpec(p(r('a ') + wpsBox(p('x'), 1) + wpsBox(p('y'), 2)));
  assert.throws(() => editInTextBox(doc, span(at([0, -1, 0, 0, 0], 0), at([0, -1, 1, 0, 0], 1)), (d, range) => deleteRange(d, range)), RangeError);
  assert.throws(() => editInTextBox(doc, span(at([0], 0), at([0, -1, 1, 0, 0], 1)), (d, range) => deleteRange(d, range)), RangeError);
});

test('in a room a command in a text box does not travel: its operations carry no command, so the room refuses them', async () => {
  const { doc } = await openSpec(p(r('a ') + wpsBox(p('x'))));
  tracing(true);
  try {
    const outside = editInTextBox(doc, at([0], 2), (d, range) => shared.insertText(d, range, 'b'));
    assert.ok(wiresOf(outside.ops));
    const inside = editInTextBox(doc, at([0, -1, 0, 0, 0], 1), (d, range) => shared.insertText(d, range, 'b'));
    assert.ok(inside.ops.length > 0);
    assert.equal(wiresOf(inside.ops), undefined);
  } finally {
    tracing(false);
  }
});

// ---------------------------------------------------------------- leaving a box

test('textBoxExit: the caret leaves to just before the drawing of the outermost box, in the paragraph that anchors it', async () => {
  const { doc } = await openSpec(p(r('one ') + wpsBox(p('x'), 1) + r('two ') + wordBox(p(r('y ') + vmlBox(p('inner'))), 2) + r('end')));
  assert.deepEqual(textBoxExit(doc, at([0, -1, 0, 0, 0], 1)), { story: 'body', path: [0], offset: 4, before: true });
  // Alternate content is kept whole: the place before it is no side of a picture.
  assert.deepEqual(textBoxExit(doc, at([0, -1, 1, 0, 0], 1)), { story: 'body', path: [0], offset: 8 });
  assert.deepEqual(textBoxExit(doc, at([0, -1, 1, 0, 0, -1, 0, 0, 0], 2)), { story: 'body', path: [0], offset: 8 });
  assert.deepEqual(textBoxExit(doc, at([0], 2)), at([0], 2));
});

// ------------------------------------------------------------------------ speed

test('speed: typing in one of 1,000 text boxes takes time that does not follow the number of boxes', async () => {
  let body = '';
  for (let i = 0; i < 1000; i++) body += p(r('anchor ' + i + ' ') + wordBox(p('box ' + i + ' text') + p('second paragraph ' + i), i + 1));
  const opened = await openSpec(body);
  const { doc } = opened;
  const history = new DocHistory(doc);
  const t = Date.now();
  type(doc, history, at([500, -1, 0, 0, 0], 3), 'abcdefghij'.repeat(5));
  const each = (Date.now() - t) / 50;
  console.log(`  one key in a box of 1,000: ${each.toFixed(2)} ms`);
  assert.equal(textAt(doc, [500, -1, 0, 0, 0]), 'box' + 'abcdefghij'.repeat(5) + ' 500 text');
  // A hang guard, not a measure of speed: ten times what an idle machine takes.
  assert.ok(each < 50, `${each} ms a key`);
  const t2 = Date.now();
  const after = await saved(opened);
  console.log(`  save: ${Date.now() - t2} ms`);
  assert.equal(textAt(after.doc, [500, -1, 0, 0, 0]), 'box' + 'abcdefghij'.repeat(5) + ' 500 text');
});

// ----------------------------------------------------------------------- layout

/** Everything the pages show, as text, to compare two layouts. */
const pictureOf = (pages: LayoutPage[]): string => JSON.stringify(pages);

function freshPages(doc: DocModel): LayoutPage[] {
  const layout = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  assert.equal(layout.layout().done, true);
  return layout.pages;
}

/** Brings a layout up to date after an edit as the view does: follows it, or lays out anew where it cannot. */
let followed = 0;
function follow(engine: DocLayout, doc: DocModel): DocLayout {
  const progress = engine.edited();
  let next = progress ? engine : new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  let guard = 100000;
  while (!next.layout().done && guard-- > 0);
  // The engine finds out while it lays the changed blocks out whether it could follow the edit after all.
  if (progress && !next.broken) followed++;
  else if (next.broken) {
    next = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
    while (!next.layout().done && guard-- > 0);
  }
  return next;
}

test('the layout follows an edit in a text box: the pages show the new text where a new layout shows it', async () => {
  const { doc } = await openSpec(p('first') + p(r('anchor ') + wordBox(p('box text') + p('two'))) + p('after'));
  const history = new DocHistory(doc);
  let engine = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  engine.layout();
  let caret = type(doc, history, at([1, -1, 0, 0, 0], 3), 'XYZ');
  engine = follow(engine, doc);
  assert.equal(pictureOf(engine.pages), pictureOf(freshPages(doc)));
  const shown = engine.pages[0].items.filter((item): item is TextItem => item.type === 'text' && item.source !== undefined && item.source.path.indexOf(-1) >= 0);
  assert.ok(shown.some((item) => item.text.indexOf('XYZ') >= 0));
  caret = step(doc, history, caret, (d, range) => splitParagraph(d, range.focus)).focus;
  engine = follow(engine, doc);
  assert.equal(pictureOf(engine.pages), pictureOf(freshPages(doc)));
  // A line mark for the new paragraph of the box: the caret can be drawn there.
  const marks = engine.pages[0].items.filter((item): item is LineMark => item.type === 'mark');
  assert.ok(marks.some((mark) => mark.source.path.join('.') === [1, -1, 0, 0, 1].join('.')));
  history.undo();
  engine = follow(engine, doc);
  assert.equal(pictureOf(engine.pages), pictureOf(freshPages(doc)));
  // Each of the edits was followed, not laid out anew.
  assert.equal(followed, 3);
});

test('random edits in text boxes, undone and redone: the layout that follows them shows what a new layout shows', async () => {
  const body =
    p('top') +
    p(r('a ') + wordBox(p('one box') + p('second'), 1) + r(' b ') + wpsBox(p('two box'), 2)) +
    tbl(tr(tc(p(r('cell ') + vmlBox(p('cell box'))))), [3000]) +
    p(r('outer ') + wpsBox(p(r('nest ') + vmlBox(p('inner box'))), 3)) +
    p('end');
  const boxes = [
    [1, -1, 0, 0],
    [1, -1, 1, 0],
    [2, 0, 0, 0, -1, 0, 0],
    [3, -1, 0, 0],
    [3, -1, 0, 0, 0, -1, 0, 0],
  ];
  let seed = 7;
  const next = (n: number): number => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed % n;
  };
  const { doc } = await openSpec(body);
  const history = new DocHistory(doc, 100000);
  let engine = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  engine.layout();
  const before = followed;
  for (let i = 0; i < 120; i++) {
    const roll = next(10);
    if (roll === 0 && history.canUndo) history.undo();
    else if (roll === 1 && history.canRedo) history.redo();
    else {
      const box = boxes[next(boxes.length)];
      // A paragraph of the box: the first one, or one Enter made.
      const blocks = (() => {
        try {
          return resolvePosition(doc, at(box.concat([1]), 0));
        } catch {
          return undefined;
        }
      })();
      const path = box.concat([blocks && next(2) === 0 ? 1 : 0]);
      const length = resolvePosition(doc, at(path, 0)).length;
      const offset = next(length + 1);
      const caret = at(path, offset);
      const kind = next(4);
      history.seal();
      if (kind === 0) step(doc, history, caret, (d, range) => insertText(d, range, 'w' + i));
      else if (kind === 1) step(doc, history, caret, (d, range) => splitParagraph(d, range.focus));
      else if (kind === 2) step(doc, history, caret, (d, range) => deleteBackward(d, range.focus));
      else step(doc, history, span(at(path, 0), at(path, Math.min(length, 2))), (d, range) => toggleRunFormat(d, range, 'bold'));
    }
    engine = follow(engine, doc);
    assert.equal(pictureOf(engine.pages), pictureOf(freshPages(doc)), `step ${i}`);
  }
  // Most of the edits were followed, not laid out anew.
  console.log(`  ${followed - before} of 120 random edits in boxes followed by the layout`);
  assert.ok(followed - before > 60, `${followed - before} followed`);
});

test('near the limit of work for text boxes: edits in boxes and in the text, undone and redone, laid out as a new layout lays them out', async () => {
  const HEADER = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
  let body = p('top ' + 'w '.repeat(30));
  for (let i = 0; i < 6; i++) body += p(r('a' + i + ' ' + 'word '.repeat(20)) + (i % 3 === 0 ? wordBox(p('box ' + i + ' ' + 'x'.repeat(40)), i + 1) : i % 3 === 1 ? wpsBox(p('box ' + i), i + 1) : vmlBox(p('vml ' + i))));
  body += tbl(tr(tc(p(r('cell ') + wpsBox(p('cell box'), 50)))), [3000]);
  body += p(r('outer ') + wpsBox(p(r('nest ') + vmlBox(p('inner box'))), 60)) + p('end');
  const boxes: [string, number[]][] = [];
  for (let i = 0; i < 6; i++) boxes.push(['body', [1 + i, -1, 0, 0]]);
  boxes.push(['body', [7, 0, 0, 0, -1, 0, 0]], ['body', [8, -1, 0, 0]], ['body', [8, -1, 0, 0, 0, -1, 0, 0]], ['header:rIdH1', [0, -1, 0, 0]]);
  for (const [seed0, scale] of [
    [1, 1.2],
    [2, 1.05],
    [3, 1.5],
  ]) {
    const { doc } = await openSpec({ body: body + HEADER, styles: STYLES, headers: { rIdH1: p(r('head ') + wpsBox(p('header box'), 70)) } });
    let seed = seed0;
    const rnd = (n: number): number => {
      seed = (seed * 1103515245 + 12345) & 0x7fffffff;
      return seed % n;
    };
    const base = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
    while (!base.layout().done);
    const options = { lineMarks: true, maxTextBoxWork: Math.floor((base as unknown as { ctx: { boxWork: number } }).ctx.boxWork * scale) };
    const fresh = (): DocLayout => {
      const layout = new DocLayout(doc, fakeMeasurer, options);
      while (!layout.layout().done);
      return layout;
    };
    let engine = fresh();
    const history = new DocHistory(doc, 100000);
    for (let i = 0; i < 160; i++) {
      const roll = rnd(12);
      try {
        if (roll === 0 && history.canUndo) history.undo();
        else if (roll === 1 && history.canRedo) history.redo();
        else if (roll === 2) {
          const index = rnd(9);
          const caret = at([index], rnd(resolvePosition(doc, at([index], 0)).length + 1));
          history.seal();
          history.applyEdit(rnd(2) ? splitParagraph(doc, caret) : insertText(doc, caret, 'body' + i), span(caret, caret));
        } else {
          const [story, box] = boxes[rnd(boxes.length)];
          let count = 1;
          try {
            while (count < 50) {
              resolvePosition(doc, at(box.concat([count]), 0, story));
              count++;
            }
          } catch {
            // The paragraphs of the box end here.
          }
          const path = box.concat([rnd(count)]);
          const length = resolvePosition(doc, at(path, 0, story)).length;
          const caret = at(path, rnd(length + 1), story);
          const kind = rnd(4);
          const command: BoxCommand =
            kind === 0 ? (d, range) => insertText(d, range, 'w' + i) : kind === 1 ? (d, range) => splitParagraph(d, range.focus) : kind === 2 ? (d, range) => deleteBackward(d, range.focus) : (d, range) => toggleRunFormat(d, range, 'bold');
          const range = kind === 3 ? span(at(path, 0, story), at(path, Math.min(length, 3), story)) : span(caret, caret);
          if (rnd(3) === 0) history.seal();
          step(doc, history, range, command);
        }
      } catch {
        // A refused edit changes nothing.
      }
      const progress = engine.edited();
      if (!progress) engine = new DocLayout(doc, fakeMeasurer, options);
      while (!engine.layout().done);
      if (engine.broken) engine = fresh();
      assert.equal(pictureOf(engine.pages), pictureOf(fresh().pages), `seed ${seed0} x${scale} step ${i}`);
    }
  }
});

test('speed: a key in one of 1,000 text boxes, with the layout brought up to date, takes time that does not follow the number of boxes', async () => {
  let body = '';
  for (let i = 0; i < 1000; i++) body += p(r('anchor ' + i + ' ') + wordBox(p('box ' + i + ' text') + p('second paragraph ' + i), i + 1));
  const { doc } = await openSpec(body);
  const history = new DocHistory(doc);
  let engine = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
  let guard = 100000;
  while (!engine.layout().done && guard-- > 0);
  let caret = at([500, -1, 0, 0, 0], 3);
  const keys = 30;
  let worst = 0;
  const t = Date.now();
  for (let i = 0; i < keys; i++) {
    const t0 = Date.now();
    caret = step(doc, history, caret, (d, range) => insertText(d, range, 'k')).focus;
    engine = follow(engine, doc);
    worst = Math.max(worst, Date.now() - t0);
  }
  const each = (Date.now() - t) / keys;
  assert.ok(followed >= keys, 'the keys were followed by the layout');
  console.log(`  a key and its layout in a box of 1,000 (${engine.pages.length} pages): ${each.toFixed(1)} ms on average, ${worst} ms at most`);
  // A hang guard, not a measure of speed.
  assert.ok(each < 500, `${each} ms a key`);
});

test('speed with hostile input: typing in a box of 5,000 paragraphs (DrawingML alone, and as Word writes it)', async () => {
  let inner = '';
  for (let i = 0; i < 5000; i++) inner += p('paragraph ' + i + ' of a long box');
  for (const [what, box] of [
    ['DrawingML', wpsBox(inner)],
    ['Word', wordBox(inner)],
  ]) {
    const opened = await openSpec(p(r('a ') + box));
    const { doc } = opened;
    const history = new DocHistory(doc);
    const t = Date.now();
    type(doc, history, at([0, -1, 0, 0, 2500], 0), 'abcdefghij');
    const each = (Date.now() - t) / 10;
    console.log(`  one key in a ${what} box of 5,000 paragraphs: ${each.toFixed(1)} ms`);
    assert.equal(textAt(doc, [0, -1, 0, 0, 2500]), 'abcdefghijparagraph 2500 of a long box');
    // A hang guard, not a measure of speed.
    assert.ok(each < 2000, `${what}: ${each} ms a key`);
    // The box is one block of the body: the layout lays all of it out again.
    let engine = new DocLayout(doc, fakeMeasurer, { lineMarks: true });
    while (!engine.layout().done);
    const t2 = Date.now();
    for (let i = 0; i < 3; i++) {
      step(doc, history, at([0, -1, 0, 0, 2500], 0), (d, range) => insertText(d, range, 'k'));
      engine = follow(engine, doc);
    }
    console.log(`  and with the layout: ${((Date.now() - t2) / 3).toFixed(0)} ms a key`);
  }
});

test('a box Word writes, and one in a header: edit, undo and save gives the bytes that were read', async () => {
  const opened = await openSpec({
    body: p(r('a ') + wordBox(p('word box'))) + HEADER_SECT,
    styles: STYLES,
    headers: { rIdH1: p(r('head ') + wpsBox(p('header box'), 7)) },
  });
  const { doc } = opened;
  const history = new DocHistory(doc);
  type(doc, history, at([0, -1, 0, 0, 0], 4), 'X');
  history.seal();
  type(doc, history, at([0, -1, 0, 0, 0], 3, 'header:rIdH1'), 'Y');
  history.undo();
  history.undo();
  const back = await saved({ ...opened });
  assert.deepEqual(back.bytes, opened.bytes);
});

// Silence the unused import checker for helpers that only some tests need.
void ((x: DocEdit) => x);
