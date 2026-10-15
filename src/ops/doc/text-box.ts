// Editing in text boxes. The commands of a document do not edit inside text boxes (a path through
// one is refused everywhere); Find and Replace and the editor do, through the functions here:
// replaceInTextBox() and insertTextInTextBox() for Find and Replace, editInTextBox() for any
// command the editor runs where the caret is in a text box (typing, deleting, Enter, formatting).
//
// A text box is taken out as a list of blocks of its own and made the body of a stand-in
// document that shares everything else with the real one; replaceText() runs on that, so the
// text is replaced by the same rules as anywhere else. The blocks it leaves go back into the
// drawing, as a new drawing item in a new copy of the paragraph that anchors it, and that
// paragraph is what changes in the document: one replaceBlocks operation of the list it is in,
// undone like any other edit. A text box in a text box goes up the same way, one box at a time.
// All matches under one anchoring paragraph (in all of its boxes, the nested ones too) are
// replaced together, so that "Replace all" over a large box, or over a paragraph with many
// boxes, copies each box and the paragraph once and not once per match.
//
// editInTextBox() runs the command itself on the stand-in, with the range lowered into it. Of what
// the command returns only changes of blocks of the stand-in's body are taken: a command that would
// change anything else (the numbering, the relationships, the styles, the comments) is refused with
// RangeError(TEXT_BOX_OUTSIDE), since that change would belong to the stand-in alone and be lost.
//
// Refused (RangeError, nothing changed): where changes are tracked (TRACK_UNSUPPORTED: the
// revisions of text boxes cannot be accepted or rejected here); RangeError(READ_ONLY) where the
// rules of find-edit.ts say the box may not be changed - the writer could not write the change
// back in every branch of alternate content (textBoxesWritable(), the rule the writer follows),
// the run or paragraph that holds it is in alternate content or a locked content control, the
// shape locks its text (textLocked()) - and whatever replaceText() or the command itself refuses.

import { findWrapper, type Block, type DocModel, type Paragraph, type RunItem, type Wrapper } from '../../model/doc/types';
import { buildEdit, type DocEdit, type DocEditOptions } from './draft';
import { applyDocOp, type DocOp } from './ops';
import { READ_ONLY } from './pieces';
import {
  asRange,
  containerBlocks,
  eachGraphic,
  hasAtomAt,
  inTextBox,
  itemLength,
  paragraphLength,
  resolvePosition,
  textBoxRootLength,
  type DocPosition,
  type DocRange,
  type GraphicPlace,
  type ParagraphGraphic,
} from './position';
import { noteStandIn } from './stand-in';
import type { BodyContext } from '../../formats/docx/write-body';
import { fixedWrap, graphicFixed, paragraphFixed, replaceText, storyContext, textLocked } from './find-edit';
import { expandRangeForTables, insertText } from './text';
import { TRACK_UNSUPPORTED, trackingFor } from './track';

/** Deepest text box in a text box that is followed. */
const MAX_BOX_DEPTH = 16;

const bad = (message: string): never => {
  throw new RangeError(message);
};

function inAlternate(wrap: Wrapper | undefined): boolean {
  for (let w = wrap, steps = 0; w && steps < 4096; w = w.parent, steps++) if (w.kind === 'alternate') return true;
  return false;
}

/** A document that is the real one except for its body, which is a copy of the list `blocks`. */
function standIn(doc: DocModel, blocks: Block[]): DocModel {
  const body = { ...doc.body, blocks: blocks.slice() };
  delete body.sectionIndex;
  delete body.section;
  delete body.sectionSpan;
  delete body.modified;
  const out: DocModel = { ...doc, body };
  // Ids and what else is kept beside a document are those of the real one.
  noteStandIn(out, doc);
  return out;
}

/** Thrown (as a RangeError) for a command in a text box that would change more than the blocks of the box: the numbering, the styles, a relationship. */
export const TEXT_BOX_OUTSIDE = 'This change reaches outside the text box and cannot be made in it';

/** Applies the operations of an edit of a stand-in to it: only changes of its body may be made there (anything else would belong to the stand-in alone). */
function applyInside(inside: DocModel, ops: DocOp[]): void {
  for (const op of ops) if (op.type !== 'replaceBlocks' || op.story !== 'body') bad(TEXT_BOX_OUTSIDE);
  for (const op of ops) applyDocOp(inside, op);
}

/** The paragraph with some of its graphics (by their number) replaced; undefined when one of them sits where it cannot be changed. */
function withGraphics(paragraph: Paragraph, replaced: Map<ParagraphGraphic, ParagraphGraphic>): Paragraph | undefined {
  let left = replaced.size;
  const swap = (items: RunItem[], depth: number): RunItem[] | undefined => {
    let out: RunItem[] | undefined;
    for (let i = 0; i < items.length && left > 0; i++) {
      const item = items[i];
      let next: RunItem | undefined;
      if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        next = replaced.get(item);
        if (next) left--;
      } else if (item.type === 'alternate' && depth < 32) {
        const inner = swap(item.items, depth + 1);
        if (inner) next = { ...item, items: inner };
      }
      if (next) {
        if (!out) out = items.slice();
        out[i] = next;
      }
    }
    return out;
  };
  const children = paragraph.children.slice();
  for (let i = 0; i < children.length && left > 0; i++) {
    const inline = children[i];
    if (inline.type !== 'run') continue;
    const content = swap(inline.content, 0);
    if (!content) continue;
    // Alternate content around the run is written as it was read while its branch is unchanged: a change would lose the other branch.
    if (inAlternate(inline.wrap)) return undefined;
    children[i] = { ...inline, content };
  }
  if (left > 0) return undefined;
  const out: Paragraph = { ...paragraph, children };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

/** A position of a document as a position of the stand-in for the text box its path enters at `from`, and back from a stand-in under `base`. */
function lower(position: DocPosition, from: number): DocPosition {
  const out: DocPosition = { story: 'body', path: position.path.slice(from), offset: position.offset };
  if (position.before === true) out.before = true;
  return out;
}

function lift(position: DocPosition, story: string, base: number[]): DocPosition {
  const out: DocPosition = { story, path: base.concat(position.path), offset: position.offset };
  if (position.before === true) out.before = true;
  return out;
}

const liftRange = (range: DocRange, story: string, base: number[]): DocRange => ({ anchor: lift(range.anchor, story, base), focus: lift(range.focus, story, base) });

/** Length of the path of the paragraph that anchors the outermost text box of a path; -1 outside text boxes. */
const anchorLength = (path: number[]): number => path.indexOf(-1);

/** What a replacement under one anchoring paragraph came to. */
interface Outcome {
  ops: DocOp[];
  replaced: number;
  failed: number;
  /** The selection the last replacement made (the one of the first range in document order), in the coordinates of the document the work was done in. */
  selection?: DocRange;
}

/**
 * Replaces `ranges` (all under the paragraph at `anchor` of `story`, in its text boxes, last first)
 * with `text`. `strict`: a range that cannot be replaced throws; otherwise it is counted in `failed`.
 * `ctx`: the writer's context of the part of the story (the namespaces of the kept trees).
 */
function replaceUnder(ctx: BodyContext, doc: DocModel, story: string, anchor: number[], ranges: readonly DocRange[], text: string, strict: boolean, depth: number): Outcome {
  if (depth > MAX_BOX_DEPTH) bad('The text box is nested too deep');
  const blocks = containerBlocks(doc, story, anchor.slice(0, -1));
  const paragraph = blocks[anchor[anchor.length - 1]];
  if (!paragraph || paragraph.type !== 'paragraph') return bad('The path does not lead to a text box');
  if (paragraphFixed(paragraph)) bad(READ_ONLY);
  const graphics: ParagraphGraphic[] = [];
  const fixed: boolean[] = [];
  eachGraphic(paragraph, (item, _index, place) => {
    graphics.push(item);
    fixed.push(item.textBoxes !== undefined && graphicFixed(ctx, item, place));
  });
  const at = anchor.length;
  // The boxes the ranges are in, in the order the ranges come.
  const boxes = new Map<string, { graphic: number; box: number; ranges: DocRange[] }>();
  for (const range of ranges) {
    const path = range.anchor.path;
    const key = path[at + 1] + ':' + path[at + 2];
    let entry = boxes.get(key);
    if (!entry) boxes.set(key, (entry = { graphic: path[at + 1], box: path[at + 2], ranges: [] }));
    entry.ranges.push(range);
  }
  const outcome: Outcome = { ops: [], replaced: 0, failed: 0 };
  const changed = new Map<number, Block[][]>();
  boxes.forEach((entry) => {
    const item = graphics[entry.graphic];
    const contents = item && item.textBoxes;
    if (!contents || !(entry.box >= 0 && entry.box < contents.length && Math.floor(entry.box) === entry.box)) bad('The path does not lead to a text box');
    if (fixed[entry.graphic]) {
      // The writer could not write the change back without losing something: not one range of the box is replaced.
      if (strict) bad(READ_ONLY);
      outcome.failed += entry.ranges.length;
      return;
    }
    const base = anchor.concat([-1, entry.graphic, entry.box]);
    const inside = standIn(doc, (contents as Block[][])[entry.box]);
    const own = replaceInside(ctx, inside, entry.ranges.map((range) => ({ anchor: lower(range.anchor, base.length), focus: lower(range.focus, base.length) })), text, strict, depth);
    outcome.replaced += own.replaced;
    outcome.failed += own.failed;
    // The boxes come last first: the last one's is the selection of the first range.
    if (own.selection) outcome.selection = liftRange(own.selection, story, base);
    if (own.replaced === 0) return;
    let next = changed.get(entry.graphic);
    if (!next) changed.set(entry.graphic, (next = (contents as Block[][]).slice()));
    next[entry.box] = inside.body.blocks;
  });
  if (changed.size === 0) return outcome;
  const replaced = new Map<ParagraphGraphic, ParagraphGraphic>();
  changed.forEach((textBoxes, graphic) => replaced.set(graphics[graphic], { ...graphics[graphic], textBoxes } as ParagraphGraphic));
  const next = withGraphics(paragraph, replaced);
  if (!next) return bad(READ_ONLY);
  outcome.ops = buildEdit(doc, (draft) => {
    draft.setParagraph(story, anchor, next);
    return { story, path: anchor, offset: 0 };
  }).ops;
  return outcome;
}

/** Replaces ranges of a stand-in (its body is a text box), last first: text of the box itself with replaceText(), text of boxes in it through replaceUnder(). The stand-in is changed. */
function replaceInside(ctx: BodyContext, inside: DocModel, ranges: readonly DocRange[], text: string, strict: boolean, depth: number): { replaced: number; failed: number; selection?: DocRange } {
  let replaced = 0;
  let failed = 0;
  let selection: DocRange | undefined;
  for (let i = 0; i < ranges.length; i++) {
    const range = ranges[i];
    const anchor = anchorLength(range.anchor.path);
    if (anchor < 0) {
      try {
        const edit = replaceText(inside, range, text);
        applyInside(inside, edit.ops);
        if (edit.ops.length > 0) replaced++;
        selection = edit.selection;
      } catch (error) {
        if (strict) throw error;
        failed++;
      }
      continue;
    }
    // The ranges under the same paragraph of the box, one after the other.
    const path = range.anchor.path.slice(0, anchor);
    const group = [range];
    while (i + 1 < ranges.length && sameAnchor(ranges[i + 1], path)) group.push(ranges[++i]);
    try {
      const outcome = replaceUnder(ctx, inside, 'body', path, group, text, strict, depth + 1);
      applyInside(inside, outcome.ops);
      replaced += outcome.replaced;
      failed += outcome.failed;
      if (outcome.selection) selection = outcome.selection;
    } catch (error) {
      if (strict) throw error;
      failed += group.length;
    }
  }
  return selection ? { replaced, failed, selection } : { replaced, failed };
}

function sameAnchor(range: DocRange, anchor: number[]): boolean {
  const path = range.anchor.path;
  if (anchorLength(path) !== anchor.length) return false;
  for (let i = 0; i < anchor.length; i++) if (path[i] !== anchor[i]) return false;
  return true;
}

/** Checks that both ends of a range are in the same text box; false for a range outside text boxes. */
function inOneBox(range: DocRange): boolean {
  if (!range || !range.anchor || !range.focus || !Array.isArray(range.anchor.path) || !Array.isArray(range.focus.path)) return bad('A range needs two positions');
  if (!inTextBox(range.anchor.path) && !inTextBox(range.focus.path)) return false;
  const root = textBoxRootLength(range.anchor.path);
  if (range.focus.story !== range.anchor.story || textBoxRootLength(range.focus.path) !== root) return bad('The positions are in different text boxes');
  for (let i = 0; i < root; i++) if (range.focus.path[i] !== range.anchor.path[i]) bad('The positions are in different text boxes');
  return true;
}

/**
 * A key that is the same for two ranges under the same paragraph that anchors text boxes (the
 * paragraph of the story, not of a box, that holds the outermost box); '' outside text boxes.
 * The ranges of one key are replaced together (replaceInTextBox).
 */
export function textBoxKey(range: DocRange): string {
  if (!range || !range.anchor || !Array.isArray(range.anchor.path)) return '';
  const at = anchorLength(range.anchor.path);
  return at < 0 ? '' : range.anchor.story + '\u0001' + range.anchor.path.slice(0, at).join(',');
}

/**
 * Replaces a range of a text box with text, as insertText() does elsewhere (where the range is
 * not in a text box, it is insertText()). One edit; its operation replaces the paragraph that
 * anchors the outermost text box. Throws what insertText() throws, RangeError(TRACK_UNSUPPORTED)
 * where changes are tracked and RangeError(READ_ONLY) for a drawing in alternate content around runs.
 */
export function insertTextInTextBox(doc: DocModel, at: DocPosition | DocRange, text: string, options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  if (!inOneBox(range)) return insertText(doc, range, text, undefined, options);
  if (trackingFor(doc, options)) bad(TRACK_UNSUPPORTED);
  const anchor = range.anchor.path.slice(0, anchorLength(range.anchor.path));
  const outcome = replaceUnder(storyContext(doc, range.anchor.story), doc, range.anchor.story, anchor, [range], text, true, 0);
  return { ops: outcome.ops, selection: outcome.selection || range };
}

/**
 * Replaces ranges in the text boxes under one anchoring paragraph (all of one textBoxKey; the
 * matches of "Replace all", last first, so that the places of the ones before stay where they
 * are) with the same text, as one edit. `replaced`: how many were; `failed`: ranges insertText()
 * refused, which stay as they are. `edit` is undefined when none was replaced. Throws
 * RangeError(TRACK_UNSUPPORTED) where changes are tracked, RangeError(READ_ONLY) for a drawing in
 * alternate content around runs, and RangeError for ranges not all of one key; nothing is
 * changed then.
 */
export function replaceInTextBox(doc: DocModel, ranges: readonly DocRange[], text: string, options?: DocEditOptions): { edit?: DocEdit; replaced: number; failed: number } {
  if (ranges.length === 0) return { replaced: 0, failed: 0 };
  const key = textBoxKey(ranges[0]);
  if (key === '') return bad('The range is not in a text box');
  for (const range of ranges) if (!inOneBox(range) || textBoxKey(range) !== key) bad('The ranges are under different paragraphs');
  if (trackingFor(doc, options)) bad(TRACK_UNSUPPORTED);
  const first = ranges[0].anchor;
  const outcome = replaceUnder(storyContext(doc, first.story), doc, first.story, first.path.slice(0, anchorLength(first.path)), ranges, text, false, 0);
  if (outcome.replaced === 0 || !outcome.selection) return { replaced: 0, failed: outcome.failed };
  return { edit: { ops: outcome.ops, selection: outcome.selection }, replaced: outcome.replaced, failed: outcome.failed };
}

/** A command as editInTextBox() runs it: on a document (the real one, or the stand-in of a text box) and a range of it. */
export type BoxCommand = (doc: DocModel, range: DocRange) => DocEdit;

/** The graphic number `wanted` of a paragraph (as eachGraphic() numbers them), with where it is. */
function graphicAt(paragraph: Paragraph, wanted: number): { item: ParagraphGraphic; place: GraphicPlace } | undefined {
  let found: { item: ParagraphGraphic; place: GraphicPlace } | undefined;
  eachGraphic(paragraph, (item, index, place) => {
    if (index !== wanted) return false;
    found = { item, place };
    return true;
  });
  return found;
}

/** The paragraph at `anchor` of `story` and the text box the path `path` enters there (path[anchor.length + 1] is the graphic, the next the box). */
function boxOf(doc: DocModel, story: string, anchor: number[], path: number[]): { paragraph: Paragraph; item: ParagraphGraphic; place: GraphicPlace; boxes: Block[][]; box: number } {
  const blocks = containerBlocks(doc, story, anchor.slice(0, -1));
  const paragraph = blocks[anchor[anchor.length - 1]];
  if (!paragraph || paragraph.type !== 'paragraph') return bad('The path does not lead to a text box');
  const at = anchor.length;
  const found = graphicAt(paragraph, path[at + 1]);
  const boxes = found && found.item.textBoxes;
  const box = path[at + 2];
  if (!found || !boxes || !(box >= 0 && box < boxes.length && Math.floor(box) === box)) return bad('The path does not lead to a text box');
  return { paragraph, item: found.item, place: found.place, boxes, box };
}

/** Runs `command` in the text box the path of `range` enters at `anchor` of `story` (and in the boxes in that one, down to the box the range is in). */
function editUnder(ctx: BodyContext, doc: DocModel, story: string, anchor: number[], range: DocRange, command: BoxCommand, depth: number): DocEdit {
  if (depth > MAX_BOX_DEPTH) bad('The text box is nested too deep');
  const { paragraph, item, place, boxes, box } = boxOf(doc, story, anchor, range.anchor.path);
  // The writer could not write the change back without losing something, or the shape locks its text.
  if (paragraphFixed(paragraph) || graphicFixed(ctx, item, place)) bad(READ_ONLY);
  const base = anchor.concat([-1, range.anchor.path[anchor.length + 1], box]);
  const inside = standIn(doc, boxes[box]);
  const lowered: DocRange = { anchor: lower(range.anchor, base.length), focus: lower(range.focus, base.length) };
  const inner = anchorLength(lowered.anchor.path);
  const edit = inner >= 0 ? editUnder(ctx, inside, 'body', lowered.anchor.path.slice(0, inner), lowered, command, depth + 1) : command(inside, lowered);
  if (!edit || !Array.isArray(edit.ops) || !edit.selection) return bad('A command must return an edit');
  const out: DocEdit = { ops: [], selection: liftRange(edit.selection, story, base) };
  // Typing in one box is one step of the history; typing in another box is another one.
  if (edit.mergeKey !== undefined) out.mergeKey = 'box:' + base.join(',') + '\u0001' + edit.mergeKey;
  if (edit.seal) out.seal = true;
  if (edit.ops.length === 0) return out;
  applyInside(inside, edit.ops);
  const textBoxes = boxes.slice();
  textBoxes[box] = inside.body.blocks;
  const next = withGraphics(paragraph, new Map([[item, { ...item, textBoxes } as ParagraphGraphic]]));
  if (!next) return bad(READ_ONLY);
  out.ops = buildEdit(doc, (draft) => {
    draft.setParagraph(story, anchor, next);
    return { story, path: anchor, offset: 0 };
  }).ops;
  return out;
}

/**
 * Runs a command of the editor where `at` is: in the text box it is in, or, where it is not in a
 * text box, on the document itself (command(doc, range)). In a text box the command runs on a
 * stand-in whose body is the box (see the head of this file) with the range lowered into it, and
 * what it changes there goes back into the box as one operation that replaces the paragraph that
 * anchors the outermost box: one step of the history. The selection the command leaves comes back
 * in the coordinates of the document; typing in one box is merged as it is elsewhere.
 *
 * Throws RangeError(TRACK_UNSUPPORTED) where changes are tracked (the revisions of text boxes cannot
 * be accepted or rejected here), RangeError(READ_ONLY) where the box may not be changed (see the
 * head of this file), RangeError(TEXT_BOX_OUTSIDE) where the command changes more than blocks of
 * the box, RangeError when the two
 * ends of the range are not in one text box, and whatever the command throws. Nothing is changed
 * then.
 */
export function editInTextBox(doc: DocModel, at: DocPosition | DocRange, command: BoxCommand, options?: DocEditOptions): DocEdit {
  const range = asRange(at);
  if (!inOneBox(range)) return command(doc, range);
  if (trackingFor(doc, options)) bad(TRACK_UNSUPPORTED);
  const anchor = range.anchor.path.slice(0, anchorLength(range.anchor.path));
  return editUnder(storyContext(doc, range.anchor.story), doc, range.anchor.story, anchor, range, command, 0);
}

/**
 * Reads what a range of a text box holds the way `read` reads a range of a document: `read` gets a
 * stand-in whose body is the box (see the head of this file) and the range lowered into it. For
 * what does not depend on where the range is (a copy of its content, whether a deletion would lose
 * something); nothing is changed, and what `read` does to the stand-in stays there. Outside text
 * boxes it is read(doc, range). Throws RangeError when the ends are not in one text box.
 */
export function readInTextBox<T>(doc: DocModel, at: DocPosition | DocRange, read: (doc: DocModel, range: DocRange) => T): T {
  const range = asRange(at);
  if (!inOneBox(range)) return read(doc, range);
  let current = doc;
  let story = range.anchor.story;
  let lowered = range;
  for (let depth = 0; depth <= MAX_BOX_DEPTH; depth++) {
    const at = anchorLength(lowered.anchor.path);
    if (at < 0) return read(current, lowered);
    const anchor = lowered.anchor.path.slice(0, at);
    const { boxes, box } = boxOf(current, story, anchor, lowered.anchor.path);
    current = standIn(current, boxes[box]);
    story = 'body';
    lowered = { anchor: lower(lowered.anchor, at + 3), focus: lower(lowered.focus, at + 3) };
  }
  return bad('The text box is nested too deep');
}

/**
 * expandRangeForTables() for a range of a text box: the range a deletion of it removes (whole cells
 * where it goes into or out of a table of the box), in the coordinates of the document. The
 * deletion of a command run in the box (editInTextBox) widens the range the same way, so a copy of
 * this range holds what a cut removes. Outside text boxes it is expandRangeForTables().
 */
export function expandRangeInTextBox(doc: DocModel, range: DocRange): DocRange {
  const given = asRange(range);
  if (!inOneBox(given)) return expandRangeForTables(doc, given);
  const root = given.anchor.path.slice(0, textBoxRootLength(given.anchor.path));
  const wide = readInTextBox(doc, given, (inside, lowered) => expandRangeForTables(inside, lowered));
  return liftRange(wide, given.anchor.story, root);
}

/**
 * Why the text box a position is in cannot be edited: 'locked' (a locked content control or
 * alternate content holds it, or the shape locks its text), 'unwritable' (the writer could not
 * write a change back in every branch of its alternate content); undefined where it can be, and
 * outside text boxes. The boxes around it are asked too, the outermost first. Throws RangeError for
 * a path that does not lead to a text box.
 */
export function textBoxProblem(doc: DocModel, position: DocPosition): 'locked' | 'unwritable' | undefined {
  if (!position || !Array.isArray(position.path) || anchorLength(position.path) < 0) return undefined;
  const ctx = storyContext(doc, position.story);
  let current = doc;
  let story = position.story;
  let path = position.path;
  for (let depth = 0; depth <= MAX_BOX_DEPTH; depth++) {
    const at = anchorLength(path);
    if (at < 0) return undefined;
    const anchor = path.slice(0, at);
    const { paragraph, item, place, boxes, box } = boxOf(current, story, anchor, path);
    if (paragraphFixed(paragraph) || fixedWrap(place.run.wrap) || textLocked(item)) return 'locked';
    if (graphicFixed(ctx, item, place)) return 'unwritable';
    current = standIn(current, boxes[box]);
    story = 'body';
    path = path.slice(at + 3);
  }
  return bad('The text box is nested too deep');
}

/**
 * Where the caret goes when it leaves the text box a position is in (Escape): just before the
 * drawing that holds the outermost box, in the paragraph that anchors it. A position outside text
 * boxes comes back as it is.
 */
export function textBoxExit(doc: DocModel, position: DocPosition): DocPosition {
  const at = anchorLength(position.path);
  if (at < 0) return position;
  const anchor = position.path.slice(0, at);
  const paragraph = resolvePosition(doc, { story: position.story, path: anchor, offset: 0 }).paragraph;
  const wanted = position.path[at + 1];
  // The offset the graphic sits at: the length of what comes before it in paragraphText(), in the order eachGraphic() goes.
  let offset = 0;
  let seen = 0;
  let done = false;
  const walk = (items: RunItem[], code: true | boolean[] | undefined, depth: number): void => {
    for (let i = 0; i < items.length && !done; i++) {
      if (code !== undefined && (code === true || code[i])) continue;
      const item = items[i];
      if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        if (seen++ === wanted) done = true;
      } else if (item.type === 'alternate') {
        if (depth < 32) walk(item.items, undefined, depth + 1);
      } else offset += itemLength(item);
    }
  };
  for (const inline of paragraph.children) {
    if (done) break;
    const gone = inline.wrap !== undefined && (findWrapper(inline, 'del') !== undefined || findWrapper(inline, 'moveFrom') !== undefined);
    if (gone) continue;
    if (inline.type === 'math') offset += inline.text.length;
    else if (inline.type === 'run' && inline.fieldCode !== true) walk(inline.content, inline.fieldCode, 0);
  }
  const out: DocPosition = { story: position.story, path: anchor, offset: Math.min(offset, paragraphLength(paragraph)) };
  if (done && hasAtomAt(paragraph, out.offset)) out.before = true;
  return out;
}
