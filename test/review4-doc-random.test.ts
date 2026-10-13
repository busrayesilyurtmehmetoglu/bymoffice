// Editing with and without tracking, accepting, rejecting, undoing and redoing on random
// documents (seeded): after every step the model is what a reader of the file written from it
// has, the file is well formed with ids that are unique, and an edit that is not tracked makes
// no tracked change of a paragraph mark or of formatting. Among the edits: tables, rows, columns
// and links, which are tracked changes where changes are tracked.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { Block, DocModel } from '../src/model/doc/types';
import {
  DocHistory,
  acceptRevision,
  deleteBackward,
  deleteColumn,
  deleteForward,
  deleteRange,
  deleteRow,
  insertColumn,
  insertHyperlink,
  insertRow,
  insertTable,
  insertText,
  listRevisions,
  rejectRevision,
  splitParagraph,
  storyIds,
  storyOf,
  type DocEdit,
  type DocRange,
} from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { NUMBERING, STYLES, checkInvariants, dice, paragraphAt, paragraphPaths, picture, texts, type Dice } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';

const decoder = new TextDecoder();
const AUTHORS = ['Ann', 'Bob', 'Me'];
const WORDS = ['alpha', 'beta', 'gamma', 'delta'];

/** A document whose changes sit in and around content controls, links and fields, with marks and formatting that are changes too. */
export function generateWrapped(seed: number, duplicateIds = false): DocxSpec {
  const d: Dice = dice(seed);
  let id = 100;
  let serial = 0;
  const who = (author = d.pick(AUTHORS)): string => ` w:id="${duplicateIds && d.chance(0.3) ? 100 + d.int(4) : id++}" w:author="${author}" w:date="2026-10-01T09:00:00Z"`;
  const word = (): string => d.pick(WORDS) + serial++ + ' ';
  const run = (text = word(), rPr = ''): string => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const gone = (text = word()): string => `<w:r><w:delText xml:space="preserve">${text}</w:delText></w:r>`;
  const ins = (inner: string): string => `<w:ins${who()}>${inner}</w:ins>`;
  const del = (inner: string): string => `<w:del${who()}>${inner}</w:del>`;
  const control = (inner: string): string => `<w:sdt><w:sdtPr><w:id w:val="${9000 + serial++}"/></w:sdtPr><w:sdtContent>${inner}</w:sdtContent></w:sdt>`;
  const link = (inner: string): string => `<w:hyperlink w:anchor="here">${inner}</w:hyperlink>`;
  const content = (): string => {
    let xml = '';
    const count = d.int(5);
    for (let i = 0; i < count; i++) {
      const kind = d.int(14);
      if (kind < 2) xml += run();
      else if (kind === 2) xml += ins(run());
      else if (kind === 3) xml += del(gone());
      else if (kind === 4) xml += control(ins(run()));
      else if (kind === 5) xml += control(del(gone()));
      else if (kind === 6) xml += control(run() + ins(run()) + del(gone()));
      else if (kind === 7) xml += control(ins(run() + del(gone())));
      else if (kind === 8) xml += link(ins(run()));
      else if (kind === 9) xml += control(link(del(gone())) + (d.chance(0.5) ? run() : ''));
      else if (kind === 10) xml += control('');
      else if (kind === 11) xml += ins(control(run()));
      else if (kind === 12) xml += run(word(), `<w:b/><w:rPrChange${who()}><w:rPr><w:i/></w:rPr></w:rPrChange>`);
      else xml += `<w:fldSimple w:instr=" DOCPROPERTY Title ">${ins(run())}</w:fldSimple>`;
    }
    return xml;
  };
  let body = '';
  const count = 3 + d.int(4);
  for (let i = 0; i < count; i++) {
    let pPr = '';
    if (d.chance(0.5)) pPr += `<w:jc w:val="${d.pick(['center', 'right', 'both'])}"/>`;
    if (d.chance(0.2)) pPr = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/><w:numberingChange${who()} w:original="1."/></w:numPr>` + pPr;
    const kind = d.int(6);
    if (kind === 0) pPr += `<w:rPr><w:ins${who()}/></w:rPr>`;
    else if (kind === 1) pPr += `<w:rPr><w:del${who()}/></w:rPr>`;
    else if (kind === 2) pPr += `<w:rPr><w:b/><w:rPrChange${who()}><w:rPr/></w:rPrChange></w:rPr>`;
    if (d.chance(0.2)) pPr += `<w:pPrChange${who()}><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange>`;
    body += `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content()}</w:p>`;
  }
  body += '<w:p><w:r><w:t>end</w:t></w:r></w:p>';
  return { body: body + '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>', styles: STYLES, numbering: NUMBERING };
}

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

async function open(spec: DocxSpec): Promise<Opened> {
  const bytes = await buildSpec(spec);
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

/**
 * The blocks of every story as a text that is equal for two models that hold the same content.
 * Left out: where a block stood in its source, and the id of a content control (the writer gives
 * the second half of a control that a paragraph break cut in two no id, so that no id is twice).
 */
function shapeOf(doc: DocModel): string {
  const blocks = new Map<string, Block[]>();
  for (const id of storyIds(doc)) blocks.set(id, storyOf(doc, id).blocks);
  return picture(blocks)
    .replace(/"src(Start|End)":\d+,?/g, '')
    .replace(/\{"attrs":\{"map":\[\["w:val","9\d+"\]\]\},"children":\[\],"name":"w:id","type":"element"\}/g, '')
    .replace(/,}/g, '}');
}

const REVISION = /<w:(ins|del|moveFrom|moveTo|rPrChange|pPrChange|numberingChange|cellIns|cellDel|tblGridChange)\b([^>]*)>/g;

async function check(o: Opened, label: string, uniqueIds: boolean): Promise<void> {
  await writeDocx(o.pkg, o.doc);
  const text = decoder.decode(await o.pkg.readPart('/word/document.xml'));
  parseXml(text);
  if (uniqueIds) {
    const ids = Array.from(text.matchAll(REVISION), (match) => (/\sw:id="(\d+)"/.exec(match[2]) || ['', ''])[1]);
    assert.equal(new Set(ids).size, ids.length, `${label}: an id twice (${ids.join(',')})`);
  }
  for (const deletion of text.matchAll(/<w:del [^>]*[^/]>(.*?)<\/w:del>/g)) assert.ok(!/<w:t[ >]/.test(deletion[1]), `${label}: text in a deletion: ${deletion[0]}`);
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  const a = shapeOf(o.doc);
  const b = shapeOf(again);
  if (a === b) return;
  let i = 0;
  while (i < a.length && a.charAt(i) === b.charAt(i)) i++;
  assert.fail(`${label}: the model is not what the file reads as, from ${i}\nmodel: ${a.slice(Math.max(0, i - 300), i + 300)}\nfile:  ${b.slice(Math.max(0, i - 300), i + 300)}`);
}

/**
 * The tracked changes of paragraph marks and the records of changed paragraph formatting in the body: what an edit
 * without tracking never makes. (The record of a run is left out: text typed into the run cuts it in two, each
 * half with the record.)
 */
function marksAndRecords(doc: DocModel): number {
  let count = 0;
  for (const block of doc.body.blocks) {
    if (block.type !== 'paragraph') continue;
    const props = block.props;
    if (props.change) count++;
    const mark = props.markProps;
    if (mark) for (const key of ['inserted', 'deleted', 'moveFrom', 'moveTo', 'change'] as const) if (mark[key]) count++;
    for (const extra of props.extras || []) count += (JSON.stringify(picture(extra.el)).match(/numberingChange/g) || []).length;
  }
  return count;
}

async function walk(seed: number, duplicateIds: boolean, steps: number): Promise<{ done: number; resolved: number; undone: number; tables: number }> {
  const o = await open(generateWrapped(seed, duplicateIds));
  const d = dice(seed * 31 + 7);
  const me = { author: 'Me', date: '2026-10-09T10:00:00Z' };
  let done = 0;
  let resolved = 0;
  let undone = 0;
  let tables = 0;
  await check(o, `seed ${seed} start`, !duplicateIds);
  for (let step = 0; step < steps; step++) {
    const label = `seed ${seed}${duplicateIds ? ' (ids twice)' : ''} step ${step}`;
    const kind = d.int(16);
    if (kind === 0 && o.history.canUndo) {
      o.history.undo();
      undone++;
      await check(o, label + ' undo', !duplicateIds);
      continue;
    }
    if (kind === 1 && o.history.canRedo) {
      o.history.redo();
      await check(o, label + ' redo', !duplicateIds);
      continue;
    }
    const paths = paragraphPaths(o.doc);
    const path = d.pick(paths);
    const length = texts(o.doc)[paths.indexOf(path)].length;
    const at = { story: 'body', path, offset: d.int(length + 1) };
    const to = { story: 'body', path, offset: Math.min(length, at.offset + 1 + d.int(4)) };
    const tracked = d.chance(0.5);
    const options = tracked ? { ...me, trackChanges: true } : undefined;
    const before = marksAndRecords(o.doc);
    let edit: DocEdit | undefined;
    let name = '';
    try {
      if (kind <= 3) [name, edit] = ['insertText', insertText(o.doc, at, d.pick(WORDS), undefined, options)];
      else if (kind === 4) [name, edit] = ['deleteRange', deleteRange(o.doc, { anchor: at, focus: to }, options)];
      else if (kind === 5) [name, edit] = ['splitParagraph', splitParagraph(o.doc, at, options)];
      else if (kind === 6) [name, edit] = ['deleteBackward', deleteBackward(o.doc, at, 'char', options)];
      else if (kind === 7) [name, edit] = ['deleteForward', deleteForward(o.doc, at, 'char', options)];
      else if (kind === 8) [name, edit] = ['lines', insertText(o.doc, at, 'x\ny', undefined, options)];
      else if (kind === 12) [name, edit] = ['table', insertTable(o.doc, at, 1 + d.int(2), 1 + d.int(2), options)];
      else if (kind === 13 || kind === 14) {
        // In a cell of a table, when the document has one by now.
        const cells = paths.filter((other) => other.length >= 4);
        if (!cells.length) continue;
        const cell = { story: 'body', path: d.pick(cells), offset: 0 };
        const choice = d.int(2);
        if (kind === 13) [name, edit] = choice ? ['insertRow', insertRow(o.doc, cell, d.pick(['above', 'below'] as const), options)] : ['deleteRow', deleteRow(o.doc, cell, options)];
        else [name, edit] = choice ? ['insertColumn', insertColumn(o.doc, cell, d.pick(['left', 'right'] as const), options)] : ['deleteColumn', deleteColumn(o.doc, cell, options)];
        tables++;
      } else if (kind === 15) [name, edit] = ['link', insertHyperlink(o.doc, { anchor: at, focus: to }, '#here', options)];
      else {
        const list = listRevisions(o.doc).list.filter((revision) => revision.kind !== 'other');
        if (!list.length) continue;
        const revision = d.pick(list);
        [name, edit] = kind % 2 ? ['accept', acceptRevision(o.doc, revision)] : ['reject', rejectRevision(o.doc, revision)];
        resolved++;
      }
    } catch (error) {
      // A place that cannot be edited, a change that cannot be tracked or resolved: nothing happened.
      if (!(error instanceof RangeError)) throw error;
      continue;
    }
    if (!edit.ops.length) continue;
    o.history.applyEdit(edit);
    o.history.seal();
    done++;
    checkInvariants(o.doc);
    const tag = `${label} ${name}${tracked ? ' tracked' : ''} at ${path.join('.')}:${at.offset}`;
    await check(o, tag, !duplicateIds);
    if (!tracked && kind <= 8) {
      const after = marksAndRecords(o.doc);
      // Typing and Enter remove nothing: the count is what it was. A deletion may take marks and records with it.
      if (name === 'insertText' || name === 'lines') assert.equal(after, before, tag);
      // Enter in an empty item of a list takes the item out of the list, and the record of its changed number with it.
      else if (name === 'splitParagraph') assert.ok(after === before || after === before - 1, `${tag}: ${before} -> ${after}`);
      else assert.ok(after <= before, `${tag}: ${before} -> ${after}`);
    }
  }
  // Every step taken back: the file that was read.
  while (o.history.canUndo) o.history.undo();
  await writeDocx(o.pkg, o.doc);
  assert.deepEqual(await o.pkg.save(), o.bytes, `seed ${seed}: undo`);
  return { done, resolved, undone, tables };
}

const SEEDS = [1, 2, 3, 5, 8, 13, 21, 34, 55, 89, 144, 217, 233, 377];

test('random review 4: the model is what the file reads as after every edit, resolution, undo and redo', async () => {
  let done = 0;
  let resolved = 0;
  let undone = 0;
  let tables = 0;
  for (const seed of SEEDS) {
    const result = await walk(seed, false, 70);
    done += result.done;
    resolved += result.resolved;
    undone += result.undone;
    tables += result.tables;
  }
  assert.ok(done > SEEDS.length * 25 && resolved > SEEDS.length * 3 && undone > SEEDS.length * 2, `${done} edits, ${resolved} resolved, ${undone} undone`);
  assert.ok(tables > SEEDS.length, `${tables} edits of rows and columns`);
});

test('random review 4: the same in files that come with ids twice', async () => {
  let done = 0;
  for (const seed of SEEDS) done += (await walk(seed, true, 50)).done;
  assert.ok(done > SEEDS.length * 15, String(done));
});

test('random review 4: the paragraphs of the generator hold what is tested', async () => {
  let controls = 0;
  let changes = 0;
  for (const seed of SEEDS) {
    const spec = generateWrapped(seed);
    controls += spec.body.split('<w:sdt>').length - 1;
    changes += spec.body.split('w:author=').length - 1;
    const o = await open(spec);
    assert.ok(paragraphAt(o.doc, 0));
  }
  assert.ok(controls > SEEDS.length * 2 && changes > SEEDS.length * 8, `${controls} controls, ${changes} changes`);
});
