// Find and Replace in a document. The text that is searched is the text a
// paragraph shows (paragraphText: deleted text and field codes are not in it,
// the results of fields are), one paragraph at a time: a match never runs from
// one paragraph into the next. The paragraphs are those of the stories in the
// order of storyIds() - the body, the headers, the footers, the footnotes, the
// endnotes - each story with its tables, cell by cell, and then the text boxes
// anchored in it, in the order of their anchors (a text box in a text box right
// after the box it is in). That follows Word as far as Word has an order: the
// main text, then the text boxes of the main text, then the other stories.
//
// A match does not run through what shows without a width of its own and a
// deletion would take (a picture, a note reference, a drawing in alternate
// content, a floating text box) or through the skeleton of a field: replacing
// it would take those away. Of the matches that are found, the ones in the
// result of a field, in a formula, in content that is kept as it is, in a
// content control whose content is locked (on the run, the paragraph, the row
// or the cell), in alternate content and in a text box the writer could not
// write back in every branch (find-edit.ts) are not replaced: replacing passes
// over them and counts them. A match is replaced by replaceText() (in the
// place of the matched characters), in a text box through
// insertTextInTextBox(), which refuses where changes are tracked.
//
// A search goes a few paragraphs at a time (DocSearch.step), so that it never
// holds the page up.

import { compileFind, findIn, findLastIn, foldFor, type CompiledFind, type FindQuery } from '../../edit/find';
import type { FormatLocale } from '../../format/locale';
import { paragraphText, type Block, type DocModel, type Paragraph } from '../../model/doc/types';
import type { DocEdit } from './draft';
import type { BodyContext } from '../../formats/docx/write-body';
import { fixedWrap, graphicFixed, lockedControl, replaceText, storyContext } from './find-edit';
import { fieldPairs, flatten, isAtom, isRemovable } from './pieces';
import { insertTextInTextBox, replaceInTextBox, textBoxKey } from './text-box';
import { TRACK_UNSUPPORTED } from './track';
import { eachGraphic, inTextBox, storyIds, storyOf, type DocPosition, type DocRange } from './position';

export interface DocFindOptions {
  matchCase?: boolean;
  /** Only where the text is a word of its own (no letter or digit right before or after it). */
  wholeWord?: boolean;
  /** How case is folded (Turkish "I" and "ı", "İ" and "i"). */
  locale?: FormatLocale | 'tr' | 'en';
  /** 'all' (the default): every story and the text boxes in it (see the head of this file); 'body': the body and its tables only (Word's "Main document"). */
  stories?: 'all' | 'body';
}

/** Why nothing of a paragraph can be replaced: it is in a locked content control or in alternate content ('locked'), or in a text box the writer could not write back ('textBox'). */
export type DocFixed = 'locked' | 'textBox';

/** A paragraph that is searched, and where it is. */
export interface DocParagraphRef {
  story: string;
  path: number[];
  paragraph: Paragraph;
  /** Nothing of it can be replaced, and why. */
  fixed?: DocFixed;
}

/** A match: its range, and whether it can be replaced (see the head of this file). */
export interface DocMatch {
  range: DocRange;
  replaceable: boolean;
  /** Not replaceable because its whole paragraph is not (DocParagraphRef.fixed). */
  fixed?: DocFixed;
}

/** Most paragraphs that are searched; tables are not followed deeper than MAX_TABLE_DEPTH, text boxes in text boxes not deeper than MAX_BOX_DEPTH. */
export const MAX_FIND_PARAGRAPHS = 2000000;
const MAX_TABLE_DEPTH = 64;
const MAX_BOX_DEPTH = 16;

/** A query for a document: undefined for an empty text. Throws RangeError for one that is not valid (see compileFind). */
export function compileDocFind(text: string, options: DocFindOptions = {}): CompiledFind | undefined {
  const query: FindQuery = { text, matchCase: options.matchCase === true, locale: options.locale };
  if (options.wholeWord === true) query.whole = 'word';
  return compileFind(query);
}

/** Whether the content a table, a row or a cell wrapper holds may not be replaced. */
const fixedTable = (wrap: Block['wrap']): boolean => fixedWrap(wrap);

/** The paragraphs of a list of blocks, in its tables cell by cell. `fixed`: what the list is in may not be replaced in. */
function addBlocks(out: DocParagraphRef[], story: string, blocks: Block[], prefix: number[], depth: number, fixed: DocFixed | undefined): void {
  for (let i = 0; i < blocks.length && out.length < MAX_FIND_PARAGRAPHS; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      const ref: DocParagraphRef = { story, path: prefix.concat([i]), paragraph: block };
      const why = fixed || (fixedWrap(block.wrap) ? 'locked' : undefined);
      if (why) ref.fixed = why;
      out.push(ref);
    } else if (block.type === 'table' && depth < MAX_TABLE_DEPTH) {
      const table = fixed || (fixedTable(block.wrap) ? 'locked' : undefined);
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        const inRow = table || (fixedTable(row.wrap) ? 'locked' : undefined);
        for (let c = 0; c < row.cells.length; c++) addBlocks(out, story, row.cells[c].blocks, prefix.concat([i, r, c]), depth + 1, inRow || (fixedTable(row.cells[c].wrap) ? 'locked' : undefined));
      }
    }
  }
}

/**
 * The text boxes anchored in a list of blocks (in its tables too), in the order of their anchors: the paragraphs
 * of each box, then the boxes anchored in those. `prefix`: the path of the list; `fixed`: as for addBlocks(); a
 * box the writer could not write back, or whose graphic is in a locked control or alternate content, is fixed
 * with all it holds (find-edit.ts).
 */
function addBoxes(out: DocParagraphRef[], story: string, blocks: Block[], prefix: number[], tables: number, boxes: number, fixed: DocFixed | undefined, context: () => BodyContext): void {
  for (let i = 0; i < blocks.length && out.length < MAX_FIND_PARAGRAPHS; i++) {
    const block = blocks[i];
    if (block.type === 'paragraph') {
      if (boxes >= MAX_BOX_DEPTH) continue;
      const anchor = fixed || (fixedWrap(block.wrap) ? 'locked' : undefined);
      eachGraphic(block, (item, graphic, place) => {
        const contents = item.textBoxes;
        if (!contents) return false;
        const inBox = anchor || (fixedWrap(place.run.wrap) ? 'locked' : graphicFixed(context(), item, place) ? 'textBox' : undefined);
        for (let b = 0; b < contents.length && out.length < MAX_FIND_PARAGRAPHS; b++) {
          const base = prefix.concat([i, -1, graphic, b]);
          addBlocks(out, story, contents[b], base, 0, inBox);
          addBoxes(out, story, contents[b], base, 0, boxes + 1, inBox, context);
        }
        return out.length >= MAX_FIND_PARAGRAPHS;
      });
    } else if (block.type === 'table' && tables < MAX_TABLE_DEPTH) {
      const table = fixed || (fixedTable(block.wrap) ? 'locked' : undefined);
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        const inRow = table || (fixedTable(row.wrap) ? 'locked' : undefined);
        for (let c = 0; c < row.cells.length; c++) addBoxes(out, story, row.cells[c].blocks, prefix.concat([i, r, c]), tables + 1, boxes, inRow || (fixedTable(row.cells[c].wrap) ? 'locked' : undefined), context);
      }
    }
  }
}

/** The paragraphs a search goes through, in its order (see the head of this file). */
export function docParagraphs(doc: DocModel, stories: 'all' | 'body' = 'all'): DocParagraphRef[] {
  const out: DocParagraphRef[] = [];
  if (stories === 'body') {
    addBlocks(out, 'body', doc.body.blocks, [], 0, undefined);
    return out;
  }
  const ids = storyIds(doc);
  for (let i = 0; i < ids.length && out.length < MAX_FIND_PARAGRAPHS; i++) {
    const id = ids[i];
    const blocks = storyOf(doc, id).blocks;
    addBlocks(out, id, blocks, [], 0, undefined);
    // The writer's context of the part, made once and only for a story with text boxes.
    let context: BodyContext | undefined;
    addBoxes(out, id, blocks, [], 0, 0, undefined, () => context || (context = storyContext(doc, id)));
  }
  return out;
}

/** What is known of the pieces of a paragraph: which places a match may not run through, which characters may not be replaced. */
interface Shape {
  /** Offsets with something at them that a match may not have inside it. */
  stops: number[];
  /** Ranges [start, end) of characters that cannot be replaced. */
  fixed: [number, number][];
}

function shapeOf(paragraph: Paragraph): Shape {
  const pieces = flatten(paragraph);
  const { owner } = fieldPairs(pieces);
  const stops: number[] = [];
  const fixed: [number, number][] = [];
  let offset = 0;
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    const item = piece.item;
    if (piece.len === 0) {
      // Something a deletion would take with the match: a picture, a note reference, a drawing in alternate content, a floating text box.
      const taken = isRemovable(piece) && !(item && (item.type === 'text' || item.type === 'lastRenderedPageBreak' || item.type === 'softHyphen'));
      if (isAtom(piece) || taken || (item && (item.type === 'fieldChar' || piece.code))) stops.push(offset);
      continue;
    }
    // Text that is not plain text of a run: a formula, content kept whole, the cached result of a field.
    const plain = item !== undefined && !piece.frozen && !piece.code && piece.lock === undefined && owner[i] < 0 && !inSimpleField(piece.wrap) && !lockedControl(piece.wrap);
    if (!plain) {
      const last = fixed[fixed.length - 1];
      if (last && last[1] === offset) last[1] = offset + piece.len;
      else fixed.push([offset, offset + piece.len]);
    }
    offset += piece.len;
  }
  return { stops, fixed };
}

function inSimpleField(wrap: { kind: string; parent?: unknown } | undefined): boolean {
  for (let w = wrap as { kind: string; parent?: unknown } | undefined, depth = 0; w && depth < 256; w = w.parent as typeof w, depth++) if (w.kind === 'simpleField') return true;
  return false;
}

/** Whether a match [start, end) may be found (nothing without width inside it) and replaced (no fixed character in it). */
function judge(shape: Shape, start: number, end: number): { found: boolean; replaceable: boolean } {
  // Both lists are in the order of the offsets (and the fixed ranges do not overlap): one look each, found by halving.
  const stops = shape.stops;
  let low = 0;
  let high = stops.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    docFindCounters.compared++;
    if (stops[mid] <= start) low = mid + 1;
    else high = mid;
  }
  if (low < stops.length && stops[low] < end) return { found: false, replaceable: false };
  const fixed = shape.fixed;
  low = 0;
  high = fixed.length;
  while (low < high) {
    const mid = (low + high) >> 1;
    docFindCounters.compared++;
    if (fixed[mid][1] <= start) low = mid + 1;
    else high = mid;
  }
  if (low < fixed.length && fixed[low][0] < end) return { found: true, replaceable: false };
  return { found: true, replaceable: true };
}

/** How many comparisons the judging of matches made since the module was loaded: for tests, which check that it grows with the logarithm of what a paragraph holds. */
export const docFindCounters = { compared: 0 };

/** The range of a match: after what has no width at its start, before what has none at its end, so that neither is in it. */
function matchOf(ref: DocParagraphRef, start: number, length: number, replaceable: boolean): DocMatch {
  const match: DocMatch = { range: matchRange(ref, start, length), replaceable: replaceable && ref.fixed === undefined };
  if (ref.fixed !== undefined) match.fixed = ref.fixed;
  return match;
}

export function matchRange(ref: { story: string; path: number[] }, start: number, length: number): DocRange {
  const anchor: DocPosition = { story: ref.story, path: ref.path.slice(), offset: start };
  const focus: DocPosition = { story: ref.story, path: ref.path.slice(), offset: start + length, before: true };
  return { anchor, focus };
}

/** Text and shape of the paragraphs a search came to, kept while it runs. */
class ParagraphCache {
  private readonly shapes = new Map<Paragraph, Shape>();
  shape(paragraph: Paragraph): Shape {
    let shape = this.shapes.get(paragraph);
    if (!shape) {
      shape = shapeOf(paragraph);
      // Only the paragraphs with a match are looked at this closely; enough of them are remembered.
      if (this.shapes.size > 4096) this.shapes.clear();
      this.shapes.set(paragraph, shape);
    }
    return shape;
  }
}

/**
 * The matches of one paragraph that may be found, in order, at most `limit` of them. A place
 * that cannot be a match (something without width inside it) is passed by one character, as
 * the search does: the text after its start may still hold a match.
 */
export function paragraphMatches(find: CompiledFind, ref: DocParagraphRef, cache?: ParagraphCache, limit = Infinity): DocMatch[] {
  const text = paragraphText(ref.paragraph);
  const folded = foldFor(find, text);
  const n = find.needle.length;
  let start = findIn(find, text, 0, folded);
  if (start < 0) return [];
  const shape = (cache || new ParagraphCache()).shape(ref.paragraph);
  const out: DocMatch[] = [];
  for (let guard = 0; start >= 0 && out.length < limit && guard <= text.length; guard++) {
    const verdict = judge(shape, start, start + n);
    if (verdict.found) out.push(matchOf(ref, start, n, verdict.replaceable));
    start = findIn(find, text, verdict.found ? start + n : start + 1, folded);
  }
  return out;
}

const keyOf = (story: string, path: number[]): string => story + '\u0001' + path.join(',');

/**
 * A search from a place in one direction, round the end of the document back to that place,
 * a few paragraphs at a time. Made for the document as it is; the caller starts a new one
 * after a change. Forward it starts at `from` (the end of the selection), backward it ends
 * there (the start of the selection).
 */
export class DocSearch {
  private readonly cache = new ParagraphCache();
  private readonly first: number;
  private readonly offset: number;
  /** Paragraphs looked at, from 0 to the length of the list and one more: the one the search started in comes twice. */
  private visited = 0;
  private done = false;
  /** About how much the last call of step() looked at, in paragraphs. */
  used = 0;

  constructor(
    private readonly find: CompiledFind,
    private readonly list: DocParagraphRef[],
    from: DocPosition | undefined,
    private readonly backward = false,
  ) {
    let first = -1;
    if (from) {
      const key = keyOf(from.story, from.path);
      for (let i = 0; i < list.length; i++) {
        if (keyOf(list[i].story, list[i].path) === key) {
          first = i;
          break;
        }
      }
    }
    // A place in no paragraph that is searched (a story or text box left out): from the start of the document, or back from its end.
    if (first < 0) {
      this.first = backward ? Math.max(0, list.length - 1) : 0;
      this.offset = backward ? Number.MAX_SAFE_INTEGER : 0;
    } else {
      this.first = first;
      this.offset = from ? from.offset : 0;
    }
  }

  get finished(): boolean {
    return this.done;
  }

  /** Looks at about `budget` paragraphs more: the next match, null when there is none all the way round, undefined when the budget ran out first. */
  step(budget: number): DocMatch | null | undefined {
    this.used = 0;
    if (this.done || this.list.length === 0) {
      this.done = true;
      return null;
    }
    const total = this.list.length;
    while (this.used < Math.max(1, budget)) {
      if (this.visited > total) {
        this.done = true;
        return null;
      }
      const k = this.visited;
      this.visited++;
      const index = this.backward ? (this.first - k + total * 2) % total : (this.first + k) % total;
      const ref = this.list[index];
      const text = paragraphText(ref.paragraph);
      this.used += 1 + (text.length >> 10);
      const found = this.inParagraph(ref, text, k === 0 ? 'first' : k === total ? 'last' : 'whole');
      if (found) return found;
    }
    return undefined;
  }

  /** The match to go to in one paragraph: in the one the search started in only on the side of `offset` that is the turn's. */
  private inParagraph(ref: DocParagraphRef, text: string, part: 'first' | 'whole' | 'last'): DocMatch | undefined {
    const find = this.find;
    const n = find.needle.length;
    const folded = foldFor(find, text);
    if (!this.backward) {
      let from = part === 'first' ? this.offset : 0;
      for (let guard = 0; guard <= text.length; guard++) {
        const start = findIn(find, text, from, folded);
        if (start < 0 || (part === 'last' && start >= this.offset)) return undefined;
        const verdict = judge(this.cache.shape(ref.paragraph), start, start + n);
        if (verdict.found) return matchOf(ref, start, n, verdict.replaceable);
        from = start + 1;
      }
      return undefined;
    }
    let before = part === 'first' ? this.offset : text.length;
    for (let guard = 0; guard <= text.length; guard++) {
      const start = findLastIn(find, text, before, folded);
      if (start < 0 || (part === 'last' && start + n <= this.offset)) return undefined;
      const verdict = judge(this.cache.shape(ref.paragraph), start, start + n);
      if (verdict.found) return matchOf(ref, start, n, verdict.replaceable);
      before = start + n - 1;
    }
    return undefined;
  }
}

/** Collects every match of the document, a few paragraphs at a time (for counting, and for "Replace all"). */
export class DocMatchScan {
  private readonly cache = new ParagraphCache();
  private next = 0;
  readonly matches: DocMatch[] = [];
  /** Matches counted beyond `limit`: there are more. */
  more = false;

  constructor(
    private readonly find: CompiledFind,
    private readonly list: DocParagraphRef[],
    private readonly limit = Infinity,
  ) {}

  /** Goes on for about `budget` paragraphs; true when it is done. */
  step(budget: number): boolean {
    let used = 0;
    while (used < Math.max(1, budget)) {
      if (this.next >= this.list.length || this.more) return true;
      const ref = this.list[this.next++];
      // One more than the room that is left: that one says there are more.
      const found = paragraphMatches(this.find, ref, this.cache, this.limit - this.matches.length + 1);
      used += 1 + found.length;
      for (let i = 0; i < found.length; i++) {
        if (this.matches.length >= this.limit) {
          this.more = true;
          return true;
        }
        this.matches.push(found[i]);
      }
    }
    return this.next >= this.list.length || this.more;
  }
}

export interface ReplaceMatchesOptions {
  /** Makes the edit that replaces one match outside text boxes: replaceText() by default (in a room the view passes insertText(), the command that travels). */
  insert?: (doc: DocModel, range: DocRange, text: string) => DocEdit;
  /** Applies one edit, through a history; what it throws counts the match (all matches of a text box) as not replaced. */
  apply(edit: DocEdit): void;
  /** Most matches that are replaced at once: with more nothing is replaced (`refused`). */
  limit?: number;
  /** 'skip': the matches in text boxes are not replaced (counted in `skipped` and `inTextBoxes`): a room, where such a change does not travel. */
  textBoxes?: 'replace' | 'skip';
}

export interface ReplaceMatchesResult {
  replaced: number;
  skipped: number;
  /** Of `skipped`, the matches in text boxes that were not replaced because they were skipped or the change could not be tracked; left out when there are none. */
  inTextBoxes?: number;
  /** Of `skipped`, the matches in text boxes the writer could not write back (DocMatch.fixed 'textBox'); left out when there are none. */
  textBoxesKept?: number;
  last?: DocRange;
  refused?: true;
}

/**
 * "Replace all" with the matches of a document in their order (see DocMatchScan): the ones that can be replaced are,
 * the last first, so that the places of the ones before it stay where they are; each through `insert` and `apply`
 * (one edit each: a history merges them into one step by a key of their own), except that the matches of one text
 * box are replaced together, as one edit (replaceInTextBox). Where changes are tracked the replacements are tracked;
 * in text boxes they cannot be, and are not made. `skipped`: matches that cannot be replaced and edits that were
 * refused; `last`: the selection the replacement of the first match left.
 */
export function replaceMatches(doc: DocModel, matches: readonly DocMatch[], replacement: string, options: ReplaceMatchesOptions): ReplaceMatchesResult {
  const usable = matches.filter((match) => match.replaceable);
  let skipped = matches.length - usable.length;
  if (options.limit !== undefined) {
    // The ones that are skipped are not made: they do not count.
    const made = options.textBoxes === 'skip' ? usable.filter((match) => textBoxKey(match.range) === '').length : usable.length;
    if (made > options.limit) return { replaced: 0, skipped: 0, refused: true };
  }
  const insert = options.insert || replaceText;
  let replaced = 0;
  let inTextBoxes = 0;
  let last: DocRange | undefined;
  for (let i = usable.length - 1; i >= 0; i--) {
    const key = textBoxKey(usable[i].range);
    if (key !== '') {
      // The matches of the same text box that come one after the other.
      const ranges: DocRange[] = [usable[i].range];
      while (i > 0 && textBoxKey(usable[i - 1].range) === key) ranges.push(usable[--i].range);
      if (options.textBoxes === 'skip') {
        skipped += ranges.length;
        inTextBoxes += ranges.length;
        continue;
      }
      try {
        const result = replaceInTextBox(doc, ranges, replacement);
        skipped += result.failed;
        if (!result.edit) continue;
        options.apply(result.edit);
        replaced += result.replaced;
        last = result.edit.selection;
      } catch (error) {
        skipped += ranges.length;
        if (error instanceof RangeError && error.message === TRACK_UNSUPPORTED) inTextBoxes += ranges.length;
      }
      continue;
    }
    try {
      const edit = insert(doc, usable[i].range, replacement);
      if (edit.ops.length === 0) continue;
      options.apply(edit);
      replaced++;
      last = edit.selection;
    } catch {
      skipped++;
    }
  }
  const out: ReplaceMatchesResult = { replaced, skipped };
  if (inTextBoxes > 0) out.inTextBoxes = inTextBoxes;
  const kept = matches.filter((match) => !match.replaceable && match.fixed === 'textBox').length;
  if (kept > 0) out.textBoxesKept = kept;
  if (last) out.last = last;
  return out;
}

/** The edit that replaces one match (Replace): replaceText(), or insertTextInTextBox() for a match in a text box. */
export function replaceMatchEdit(doc: DocModel, range: DocRange, text: string): DocEdit {
  return inTextBox(range.anchor.path) ? insertTextInTextBox(doc, range, text) : replaceText(doc, range, text);
}

/** Finds the next match in one go, for a document small enough not to need steps (and for Node). */
export function findInDocument(doc: DocModel, text: string, from?: DocPosition, options: DocFindOptions = {}, backward = false): DocMatch | undefined {
  const find = compileDocFind(text, options);
  if (!find) return undefined;
  const search = new DocSearch(find, docParagraphs(doc, options.stories), from, backward);
  for (;;) {
    const found = search.step(1 << 20);
    if (found !== undefined) return found === null ? undefined : found;
  }
}

/** Every match of the document in order, at most `limit`. */
export function findAllInDocument(doc: DocModel, text: string, options: DocFindOptions = {}, limit = Infinity): DocMatch[] {
  const find = compileDocFind(text, options);
  if (!find) return [];
  const scan = new DocMatchScan(find, docParagraphs(doc, options.stories), limit);
  while (!scan.step(1 << 20));
  return scan.matches;
}
