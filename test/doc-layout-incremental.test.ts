// Laying a document out again after an edit (DocLayout.edited) must give
// exactly the pages a new engine gives for the edited document. Checked here
// on random documents with random edits, after every step, and on documents
// of many pages where the point is how little is laid out again.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { DocLayout } from '../src/layout/doc/layout';
import type { LayoutPage, TextMeasurer } from '../src/layout/doc/types';
import type { DocModel } from '../src/model/doc/types';
import {
  DocHistory,
  deleteBackward,
  deleteForward,
  insertPageBreak,
  insertText,
  setParaFormat,
  setRunFormat,
  splitParagraph,
  toggleList,
  type DocEdit,
  type DocOp,
  type DocRange,
} from '../src/ops/doc/index';
import { fakeMeasurer, p, r, sectPr, STYLES, tbl, tc, tr, words } from './doc-layout-fixtures';
import { dice, measurer as opsMeasurer, NUMBERING, paragraphAt, paragraphPaths, pos, randomDoc, randomEdit, type Dice } from './doc-ops-fixtures';
import { paragraphLength } from '../src/ops/doc/index';
import { readSpec } from './docx-fixtures';

const SEEDS = Number(process.env.DOC_LAYOUT_SEEDS || 40);
const STEPS = Number(process.env.DOC_LAYOUT_STEPS || 60);

/** As the editor lays documents out: with a mark for every line. */
const OPTIONS = { lineMarks: true };

const picture = (pages: LayoutPage[]): string => JSON.stringify(pages);

function fresh(doc: DocModel, measurer: TextMeasurer): LayoutPage[] {
  const layout = new DocLayout(doc, measurer, OPTIONS);
  assert.equal(layout.layout().done, true);
  return layout.pages;
}

/**
 * Only edits of the blocks of the body can be followed (pictures and
 * relationships that come with them change nothing that is laid out
 * already); everything else needs a new engine, as the view does it.
 */
const bodyOnly = (ops: DocOp[]): boolean =>
  ops.every((op) => (op.type === 'replaceBlocks' && op.story === 'body') || op.type === 'setMedia' || ((op.type === 'setRelationship' || op.type === 'setRelationships') && op.story === 'body'));

interface Follower {
  doc: DocModel;
  measurer: TextMeasurer;
  engine: DocLayout;
  /** Edits followed by edited(), edits that needed a new engine, pages laid out again by followed edits, pages there were. */
  followed: number;
  restarted: number;
  relaid: number;
  total: number;
}

function follower(doc: DocModel, measurer: TextMeasurer): Follower {
  const engine = new DocLayout(doc, measurer, OPTIONS);
  engine.layout();
  return { doc, measurer, engine, followed: 0, restarted: 0, relaid: 0, total: 0 };
}

/** Counts the pages that are other objects than before: the ones laid out again. */
function countPages(engine: DocLayout): () => number {
  const before = engine.pages.slice();
  return () => {
    let made = 0;
    for (let i = 0; i < engine.pages.length; i++) if (engine.pages[i] !== before[i]) made++;
    return made;
  };
}

/** Brings the engine up to date after an edit and compares its pages with those of a new engine. `budget`: pages per call of layout(). */
function follow(f: Follower, ops: DocOp[], label: string, budget?: number): void {
  // The engine finds out by itself what it cannot follow: nothing tells it which story an edit was in.
  const progress = f.engine.edited();
  assert.ok(bodyOnly(ops) || !progress, `${label}: an edit outside the blocks of the body was followed`);
  if (!progress) {
    f.restarted++;
    f.engine = new DocLayout(f.doc, f.measurer, OPTIONS);
  } else f.followed++;
  const expected = picture(fresh(f.doc, f.measurer));
  // The pages before the first one laid out again are those of the old layout: they must be right already.
  if (progress) {
    const kept = f.engine.pages.slice(0, progress.from);
    const wanted = (JSON.parse(expected) as LayoutPage[]).slice(0, progress.from);
    // Totals in them are filled in at the end.
    if (progress.done) assert.equal(picture(kept), picture(wanted), `${label}: pages kept`);
  }
  let guard = 100000;
  while (!f.engine.layout(budget === undefined ? undefined : { pages: budget }).done && guard-- > 0);
  assert.ok(guard > 0, `${label}: layout did not end`);
  const actual = picture(f.engine.pages);
  if (actual !== expected) {
    const a = JSON.parse(actual) as LayoutPage[];
    const b = JSON.parse(expected) as LayoutPage[];
    let page = 0;
    while (page < a.length && page < b.length && JSON.stringify(a[page]) === JSON.stringify(b[page])) page++;
    const left = a[page] ? a[page].items : [];
    const right = b[page] ? b[page].items : [];
    let item = 0;
    while (item < left.length && item < right.length && JSON.stringify(left[item]) === JSON.stringify(right[item])) item++;
    assert.fail(
      `${label}: page ${page} of ${a.length}/${b.length} differs at item ${item} (${progress ? 'followed from page ' + progress.from : 'new engine'}):\n` +
        `  got      ${JSON.stringify(left[item])}\n  expected ${JSON.stringify(right[item])}`,
    );
  }
}

test('random edits of random documents: the engine that follows them shows what a new engine shows', async () => {
  let followed = 0;
  let restarted = 0;
  for (let seed = 1; seed <= SEEDS; seed++) {
    const doc = await randomDoc(seed);
    const d = dice(seed * 104729);
    const history = new DocHistory<DocRange | undefined>(doc, 100000);
    const f = follower(doc, opsMeasurer);
    // The operations of every step, as the view keeps them, to know what an undo touches.
    const undone: DocOp[][] = [];
    const redone: DocOp[][] = [];
    for (let step = 0; step < STEPS; step++) {
      const label = `seed ${seed} step ${step}`;
      const roll = d.int(10);
      if (roll === 0 && history.canUndo) {
        history.undo();
        const ops = undone.pop() as DocOp[];
        redone.push(ops);
        follow(f, ops, `${label} undo`);
        continue;
      }
      if (roll === 1 && history.canRedo) {
        history.redo();
        const ops = redone.pop() as DocOp[];
        undone.push(ops);
        follow(f, ops, `${label} redo`);
        continue;
      }
      let made;
      try {
        // Mostly the body; the other stories (a header, the notes) now and then.
        made = d.chance(0.8) ? randomEdit(doc, d, ['body']) : randomEdit(doc, d);
      } catch (error) {
        assert.ok(error instanceof RangeError, `${label}: ${error instanceof Error ? error.stack : error}`);
        continue;
      }
      if (!made || made.edit.ops.length === 0) continue;
      history.applyEdit(made.edit, undefined);
      history.seal();
      undone.push(made.edit.ops);
      redone.length = 0;
      follow(f, made.edit.ops, `${label} ${made.name}`, d.chance(0.3) ? 1 : undefined);
    }
    followed += f.followed;
    restarted += f.restarted;
  }
  // Most edits are followed; the ones that change lists, notes or other stories are not.
  assert.ok(followed > restarted, `followed ${followed}, restarted ${restarted}`);
});

/** A long document: paragraphs of a few lines, a page break now and then, list items, a table, headings, a footer with the page numbers. */
async function longDoc(seed: number, paragraphs: number): Promise<DocModel> {
  const d = dice(seed);
  let body = '';
  for (let i = 0; i < paragraphs; i++) {
    const kind = d.int(40);
    const text = words(8 + d.int(40), 2 + d.int(5), i * 50);
    if (kind === 0) body += p(text, '<w:pageBreakBefore/>');
    else if (kind === 1) body += p(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
    else if (kind === 2) body += tbl(tr(tc(text) + tc(words(5, 3, i))) + tr(tc('x') + tc(text)), [2400, 2400]) + p('');
    else if (kind === 3) body += p(text, '<w:keepNext/><w:pStyle w:val="Heading1"/>');
    else if (kind === 4) body += p(r(text, '<w:b/>') + r(' tail'), '<w:jc w:val="both"/><w:widowControl/>');
    else if (kind === 5) body += p(text, '<w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:space="1" w:color="000000"/></w:pBdr>');
    else if (kind === 6) body += p(text, '<w:spacing w:before="120" w:after="240"/><w:contextualSpacing/>');
    else body += p(text);
  }
  const field = (code: string) =>
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${code} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${r('1')}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  body += sectPr('<w:footerReference w:type="default" r:id="f1"/>');
  return readSpec({
    styles: STYLES + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>',
    numbering: NUMBERING,
    body,
    footers: { f1: p(r('Page ') + field('PAGE') + r(' of ') + field('NUMPAGES')) },
  });
}

/** An edit as someone typing makes it, near a place in the document. */
function typingEdit(doc: DocModel, d: Dice, near: number): DocEdit | undefined {
  const paths = paragraphPaths(doc, 'body');
  const path = paths[Math.max(0, Math.min(paths.length - 1, near + d.int(7) - 3))];
  const length = paragraphLength(paragraphAt(doc, path));
  const at = pos(path, d.int(length + 1));
  const roll = d.int(20);
  try {
    if (roll < 9) return insertText(doc, at, d.pick(['a', 'b', ' ', 'word ', 'ğ', 'longerword']));
    if (roll < 12) return deleteBackward(doc, at);
    if (roll < 14) return deleteForward(doc, at, 'word');
    if (roll < 16) return splitParagraph(doc, at);
    if (roll === 16) return setRunFormat(doc, { anchor: pos(path, 0), focus: pos(path, length) }, { bold: true });
    if (roll === 17) return setParaFormat(doc, at, { align: d.pick(['center', 'right', 'both']) });
    if (roll === 18) return insertPageBreak(doc, at);
    return toggleList(doc, at, 'number');
  } catch (error) {
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

test('typing in a long document: the same pages as a new engine, and few of them laid out again', async () => {
  for (let seed = 1; seed <= 6; seed++) {
    const doc = await longDoc(seed, 260);
    const d = dice(seed * 31337);
    const history = new DocHistory<DocRange | undefined>(doc, 100000);
    const f = follower(doc, fakeMeasurer);
    const pages = f.engine.pages.length;
    assert.ok(pages > 30, `seed ${seed}: the document has ${pages} pages`);
    let near = d.int(paragraphPaths(doc, 'body').length);
    const undone: DocOp[][] = [];
    for (let step = 0; step < 120; step++) {
      if (d.chance(0.1)) near = d.int(paragraphPaths(doc, 'body').length);
      let ops: DocOp[];
      if (d.chance(0.08) && history.canUndo) {
        history.undo();
        ops = undone.pop() as DocOp[];
      } else {
        const edit = typingEdit(doc, d, near);
        if (!edit || edit.ops.length === 0) continue;
        history.applyEdit(edit, undefined);
        history.seal();
        undone.push(edit.ops);
        ops = edit.ops;
      }
      follow(f, ops, `seed ${seed} step ${step}`, d.chance(0.5) ? 2 : undefined);
    }
    assert.ok(f.followed > f.restarted * 3, `seed ${seed}: followed ${f.followed}, restarted ${f.restarted}`);
  }
});

test('a character typed in the middle of a long document lays out a page or two, not the document', async () => {
  let body = '';
  // Twenty lines a page; every paragraph is three lines, and every tenth one starts a page.
  for (let i = 0; i < 600; i++) body += p(words(22, 4, i * 30), i % 10 === 0 && i > 0 ? '<w:pageBreakBefore/>' : '');
  const doc = await readSpec({ styles: STYLES, body: body + sectPr() });
  const engine = new DocLayout(doc, fakeMeasurer, OPTIONS);
  engine.layout();
  const count = engine.pages.length;
  assert.ok(count >= 60, `${count} pages`);
  const history = new DocHistory<DocRange | undefined>(doc);
  let relaid = 0;
  for (let step = 0; step < 30; step++) {
    const made = countPages(engine);
    history.applyEdit(insertText(doc, pos([301], 3 + step), 'x'), undefined);
    const progress = engine.edited();
    assert.ok(progress, 'the edit is followed');
    while (!engine.layout({ pages: 1 }).done);
    const closed = made();
    // Up to the next page that starts with a page break, which is at most three pages on.
    assert.ok(closed <= 4, `step ${step}: ${closed} pages were laid out again`);
    relaid += closed;
    assert.equal(engine.pages.length, count);
  }
  assert.ok(relaid <= 30 * 2.5, `${relaid} pages were laid out again in 30 steps`);
  assert.equal(picture(engine.pages), picture(fresh(doc, fakeMeasurer)));
  // Enter moves every block after it: the pages after it are made again, from lines that are kept.
  history.applyEdit(splitParagraph(doc, pos([301], 5)), undefined);
  const progress = engine.edited();
  assert.ok(progress && progress.from >= 58 && progress.from <= 60, `from page ${progress && progress.from}`);
  while (!engine.layout({ pages: 3 }).done);
  assert.equal(picture(engine.pages), picture(fresh(doc, fakeMeasurer)));
});

test('an edit beyond what is laid out so far costs nothing; one in a table that is being fed needs a new engine', async () => {
  let body = '';
  for (let i = 0; i < 200; i++) body += p(words(30, 4, i * 40));
  const doc = await readSpec({ styles: STYLES, body: body + sectPr() });
  const engine = new DocLayout(doc, fakeMeasurer, OPTIONS);
  engine.layout({ pages: 2 });
  const history = new DocHistory<DocRange | undefined>(doc);
  history.applyEdit(insertText(doc, pos([150], 0), 'later '), undefined);
  const progress = engine.edited();
  assert.deepEqual(progress, { from: 2, done: false, updated: [] });
  while (!engine.layout({ pages: 5 }).done);
  assert.equal(picture(engine.pages), picture(fresh(doc, fakeMeasurer)));

  // An edit while layout is under way, before the place it has reached.
  const second = new DocLayout(doc, fakeMeasurer, OPTIONS);
  second.layout({ pages: 6 });
  history.applyEdit(insertText(doc, pos([3], 2), 'early '), undefined);
  const again = second.edited();
  assert.ok(again && again.from <= 1, 'laid out again from the first pages');
  while (!second.layout({ pages: 4 }).done);
  assert.equal(picture(second.pages), picture(fresh(doc, fakeMeasurer)));
});

test('edits that change what the document counts are refused, the others keep list numbers and notes', async () => {
  const item = (text: string) => p(text, '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>');
  let body = item('first') + item('second') + p('between') + item('third');
  for (let i = 0; i < 60; i++) body += p(words(20, 4, i * 20));
  body += item('fourth') + p(r('with a note') + '<w:r><w:footnoteReference w:id="2"/></w:r>') + p('end');
  const doc = await readSpec({
    styles: STYLES,
    numbering: NUMBERING,
    body: body + sectPr(),
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> the note</w:t></w:r></w:p></w:footnote>',
  });
  const history = new DocHistory<DocRange | undefined>(doc);
  const f = follower(doc, fakeMeasurer);
  const step = (edit: DocEdit, followed: boolean, label: string) => {
    history.applyEdit(edit, undefined);
    const before = f.followed;
    follow(f, edit.ops, label);
    assert.equal(f.followed > before, followed, label);
  };
  // Typing in a list item and in the paragraph with the note reference: the number and the note stay.
  step(insertText(doc, pos([1], 3), 'XY'), true, 'typing in a list item');
  step(insertText(doc, pos([65], 2), 'Z'), true, 'typing beside a note reference');
  step(setRunFormat(doc, { anchor: pos([1], 0), focus: pos([1], 4) }, { italic: true }), true, 'formatting a list item');
  // A new item changes the numbers after it.
  step(splitParagraph(doc, pos([1], 2)), false, 'a new list item');
  step(insertText(doc, pos([2], 0), 'more '), true, 'typing after the new engine');
  step(toggleList(doc, pos([3], 0), 'number'), false, 'a paragraph made a list item');
  // The undo of typing puts the old paragraph back: followed.
  step(insertText(doc, pos([10], 0), 'q'), true, 'typing');
  history.undo();
  const before = f.followed;
  follow(f, [], 'undo of typing');
  assert.equal(f.followed, before + 1);
});
