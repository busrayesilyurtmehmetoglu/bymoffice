// Positions in a document and how to move between them.
//
// A position uses the coordinates the layout engine gives its text items
// (TextItem.source): the story, the path of the paragraph and an offset in
// paragraphText() of that paragraph. Offsets count UTF-16 code units. What
// paragraphText() leaves out (field codes, deleted text) has no offset; what
// it does not count (drawings, note references, bookmarks) sits between two
// offsets and has no width.
//
// A caret at an offset stands after everything without width that sits at that
// offset, except that it does not step over the beginning of a field or of a
// bookmark or comment range: what is typed there stays outside of them.

import { findWrapper, paragraphText, type Block, type DocModel, type Note, type Paragraph, type RunItem, type Story, type StoryCollection, type Table, type Wrapper } from '../../model/doc/types';

export interface DocPosition {
  /** "body", "header:<rId>", "footer:<rId>", "footnote:<id>", "endnote:<id>". */
  story: string;
  /** Path of the paragraph: see TextItem.source of the layout engine. */
  path: number[];
  /** Offset in paragraphText() of the paragraph, in UTF-16 code units. */
  offset: number;
}

export interface DocRange {
  anchor: DocPosition;
  focus: DocPosition;
}

/** Deepest path that is followed; the readers nest tables far less deep than this. */
const MAX_PATH = 1024;

/** A story as the operations see it: something that holds blocks and says whether it was edited. */
export interface StoryHandle {
  id: string;
  kind: 'body' | 'header' | 'footer' | 'footnote' | 'endnote';
  /** The part the story is in: carries `modified` and the relationships. */
  host: Story | StoryCollection<Note>;
  /** Notes only. */
  note?: Note;
  noteId?: number;
  blocks: Block[];
}

const bad = (message: string): never => {
  throw new RangeError(message);
};

/** The story a position names. Throws RangeError when the document has no such story. */
export function storyOf(doc: DocModel, id: string): StoryHandle {
  if (typeof id !== 'string') return bad('The story of a position must be a string');
  if (id === 'body') return { id, kind: 'body', host: doc.body, blocks: doc.body.blocks };
  const colon = id.indexOf(':');
  const kind = colon < 0 ? '' : id.slice(0, colon);
  const key = id.slice(colon + 1);
  if (kind === 'header' || kind === 'footer') {
    const story = (kind === 'header' ? doc.headers : doc.footers).get(key);
    if (!story) return bad(`The document has no story "${id}"`);
    return { id, kind, host: story, blocks: story.blocks };
  }
  if (kind === 'footnote' || kind === 'endnote') {
    const collection = kind === 'footnote' ? doc.footnotes : doc.endnotes;
    const noteId = /^-?\d{1,9}$/.test(key) ? Number(key) : NaN;
    const note = collection && noteId === noteId ? collection.items.get(noteId) : undefined;
    if (!collection || !note) return bad(`The document has no story "${id}"`);
    return { id, kind, host: collection, note, noteId, blocks: note.blocks };
  }
  return bad(`The document has no story "${id}"`);
}

/** Ids of the stories of a document that hold text, the body first. */
export function storyIds(doc: DocModel): string[] {
  const out = ['body'];
  doc.headers.forEach((_story, id) => out.push('header:' + id));
  doc.footers.forEach((_story, id) => out.push('footer:' + id));
  if (doc.footnotes) doc.footnotes.items.forEach((note, id) => (note.kind === undefined ? out.push('footnote:' + id) : 0));
  if (doc.endnotes) doc.endnotes.items.forEach((note, id) => (note.kind === undefined ? out.push('endnote:' + id) : 0));
  return out;
}

const isIndex = (value: unknown, length: number): value is number => typeof value === 'number' && value >= 0 && value < length && Math.floor(value) === value;

/** The drawing or legacy graphic number `wanted` of a paragraph, counted the way the layout engine numbers them. */
function graphicOf(paragraph: Paragraph, wanted: number): { textBoxes?: Block[][] } | undefined {
  let seen = 0;
  let found: { textBoxes?: Block[][] } | undefined;
  const walk = (items: RunItem[], code: true | boolean[] | undefined, depth: number): boolean => {
    for (let i = 0; i < items.length; i++) {
      if (code !== undefined && (code === true || code[i])) continue;
      const item = items[i];
      if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        if (seen++ === wanted) {
          found = item;
          return true;
        }
      } else if (item.type === 'alternate' && depth < 32 && walk(item.items, undefined, depth + 1)) return true;
    }
    return false;
  };
  for (const inline of paragraph.children) {
    if (inline.type !== 'run' || inline.fieldCode === true) continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    if (walk(inline.content, inline.fieldCode, 0)) break;
  }
  return found;
}

/** The list of blocks a container path names; undefined when the path leads nowhere. */
function blocksAt(top: Block[], path: number[]): Block[] | undefined {
  if (path.length > MAX_PATH) return undefined;
  let blocks = top;
  let i = 0;
  while (i < path.length) {
    const block = blocks[path[i]];
    if (!isIndex(path[i], blocks.length) || !block) return undefined;
    if (block.type === 'table') {
      const row = block.rows[path[i + 1]];
      const cell = row && isIndex(path[i + 1], block.rows.length) && isIndex(path[i + 2], row.cells.length) ? row.cells[path[i + 2]] : undefined;
      if (!cell) return undefined;
      blocks = cell.blocks;
      i += 3;
    } else if (block.type === 'paragraph' && path[i + 1] === -1) {
      const graphic = graphicOf(block, path[i + 2]);
      const boxes = graphic && graphic.textBoxes;
      if (!boxes || !isIndex(path[i + 3], boxes.length)) return undefined;
      blocks = boxes[path[i + 3]];
      i += 4;
    } else return undefined;
  }
  return blocks;
}

/** Blocks of the container at `path` in a story ([] is the story itself). Throws RangeError for a path that leads nowhere. */
export function containerBlocks(doc: DocModel, story: string, path: number[]): Block[] {
  const blocks = Array.isArray(path) ? blocksAt(storyOf(doc, story).blocks, path) : undefined;
  return blocks || bad('The path does not lead to a list of blocks');
}

/** True when the path goes through a text box: such content can be read but not edited. */
export const inTextBox = (path: number[]): boolean => path.indexOf(-1) >= 0;

/** Where a position is in the model. */
export interface ResolvedPosition {
  story: StoryHandle;
  /** Path of the list of blocks the paragraph is in. */
  container: number[];
  blocks: Block[];
  /** Index of the paragraph in `blocks`. */
  index: number;
  paragraph: Paragraph;
  /** Length of paragraphText() of the paragraph. */
  length: number;
  offset: number;
}

/** Offsets run items take in paragraphText(); `code` flags the items that sit in a field code and take none. */
export function itemsLength(items: RunItem[], code: true | boolean[] | undefined, depth: number): number {
  let length = 0;
  for (let i = 0; i < items.length; i++) {
    if (code !== undefined && (code === true || code[i])) continue;
    length += itemLength(items[i], depth);
  }
  return length;
}

/** Offsets one run item takes in paragraphText(). */
export function itemLength(item: RunItem, depth = 0): number {
  switch (item.type) {
    case 'text':
      return item.text.length;
    case 'tab':
    case 'break':
    case 'noBreakHyphen':
      return 1;
    case 'symbol':
      return item.char.length;
    case 'alternate':
      // Nested as deep as the reader lets them, which is far less than this.
      return depth > 64 ? 0 : itemsLength(item.items, undefined, depth + 1);
    default:
      return 0;
  }
}

/** A place in the content of a paragraph. */
export interface InlineLocation {
  /** Index in Paragraph.children; their number for the end of the paragraph. */
  inline: number;
  /** Index in the content of the run; -1 for an inline that is not a run (a formula). */
  item: number;
  /** Offset inside the item: in its text, or 0. */
  char: number;
}

const notShown = (inline: { wrap?: Wrapper }): boolean => inline.wrap !== undefined && (findWrapper(inline, 'del') !== undefined || findWrapper(inline, 'moveFrom') !== undefined);

/**
 * The item of the paragraph that holds the character at `offset` of
 * paragraphText(); the end of the paragraph for an offset equal to its
 * length; undefined beyond.
 */
export function locateOffset(paragraph: Paragraph, offset: number): InlineLocation | undefined {
  if (!(offset >= 0)) return undefined;
  let at = 0;
  const children = paragraph.children;
  for (let i = 0; i < children.length; i++) {
    const inline = children[i];
    if ((inline.type !== 'run' && inline.type !== 'math') || notShown(inline)) continue;
    if (inline.type === 'math') {
      if (offset < at + inline.text.length) return { inline: i, item: -1, char: offset - at };
      at += inline.text.length;
      continue;
    }
    const code = inline.fieldCode;
    if (code === true) continue;
    for (let k = 0; k < inline.content.length; k++) {
      if (code !== undefined && code[k]) continue;
      const length = itemLength(inline.content[k]);
      if (offset < at + length) return { inline: i, item: k, char: offset - at };
      at += length;
    }
  }
  return offset === at ? { inline: children.length, item: 0, char: 0 } : undefined;
}

/** The offset in paragraphText() of a place in the content of a paragraph: the reverse of locateOffset(). Content that is not shown has the offset of what follows it. */
export function offsetOfLocation(paragraph: Paragraph, location: InlineLocation): number {
  let at = 0;
  const children = paragraph.children;
  for (let i = 0; i < children.length && i <= location.inline; i++) {
    const inline = children[i];
    if ((inline.type !== 'run' && inline.type !== 'math') || notShown(inline)) continue;
    const last = i === location.inline;
    if (inline.type === 'math') {
      at += last ? Math.max(0, Math.min(location.char, inline.text.length)) : inline.text.length;
      continue;
    }
    const code = inline.fieldCode;
    if (code === true) continue;
    for (let k = 0; k < inline.content.length; k++) {
      if (code !== undefined && code[k]) continue;
      const length = itemLength(inline.content[k]);
      if (last && k >= location.item) return at + (k === location.item ? Math.max(0, Math.min(location.char, length)) : 0);
      at += length;
    }
  }
  return at;
}

/** Number of offsets the content of a paragraph takes: the length of paragraphText(). */
export function paragraphLength(paragraph: Paragraph): number {
  let cached = LENGTHS.get(paragraph);
  if (cached === undefined) LENGTHS.set(paragraph, (cached = paragraphText(paragraph).length));
  return cached;
}

// Paragraphs are never changed in place, so their length can be kept.
const LENGTHS = new WeakMap<Paragraph, number>();

/** Finds the paragraph of a position. Throws RangeError when the position does not name a place in the document. */
export function resolvePosition(doc: DocModel, position: DocPosition): ResolvedPosition {
  if (!position || !Array.isArray(position.path) || position.path.length === 0) return bad('A position needs a path');
  const story = storyOf(doc, position.story);
  const container = position.path.slice(0, -1);
  const index = position.path[position.path.length - 1];
  const blocks = blocksAt(story.blocks, container);
  if (!blocks || !isIndex(index, blocks.length)) return bad('The path of the position does not lead to a paragraph');
  const paragraph = blocks[index];
  if (paragraph.type !== 'paragraph') return bad('The path of the position does not lead to a paragraph');
  const length = paragraphLength(paragraph);
  if (!isIndex(position.offset, length + 1)) return bad('The offset of the position is outside its paragraph');
  return { story, container, blocks, index, paragraph, length, offset: position.offset };
}

/** The part of a path up to and including the last text box it enters: positions with different roots cannot form a range. */
function rootLength(path: number[]): number {
  for (let i = path.length - 1; i >= 0; i--) if (path[i] === -1) return i + 3;
  return 0;
}

/**
 * Compares two positions of one story in document order: negative when `a`
 * comes first. Throws RangeError for positions of different stories or text
 * boxes.
 */
export function comparePositions(a: DocPosition, b: DocPosition): number {
  if (a.story !== b.story) return bad('The positions are in different stories');
  const root = rootLength(a.path);
  if (root !== rootLength(b.path)) return bad('The positions are in different text boxes');
  for (let i = 0; i < root; i++) if (a.path[i] !== b.path[i]) return bad('The positions are in different text boxes');
  const shared = Math.min(a.path.length, b.path.length);
  for (let i = root; i < shared; i++) if (a.path[i] !== b.path[i]) return a.path[i] - b.path[i];
  if (a.path.length !== b.path.length) return a.path.length - b.path.length;
  return a.offset - b.offset;
}

export const samePosition = (a: DocPosition, b: DocPosition): boolean =>
  a.story === b.story && a.offset === b.offset && a.path.length === b.path.length && a.path.every((value, i) => value === b.path[i]);

/** The ends of a range in document order. Throws RangeError when they cannot form a range. */
export function orderRange(range: DocRange): { start: DocPosition; end: DocPosition } {
  return comparePositions(range.anchor, range.focus) <= 0 ? { start: range.anchor, end: range.focus } : { start: range.focus, end: range.anchor };
}

export const isCollapsed = (range: DocRange): boolean => samePosition(range.anchor, range.focus);

export const collapsed = (position: DocPosition): DocRange => ({ anchor: position, focus: position });

/** A position or a range as a range. */
export const asRange = (at: DocPosition | DocRange): DocRange => ('anchor' in at ? at : collapsed(at));

// ------------------------------------------------------------------ characters

const isHigh = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLow = (code: number) => code >= 0xdc00 && code <= 0xdfff;

/** Marks that belong to the character before them: combining marks, variation selectors, the joiners. */
function isExtender(code: number): boolean {
  if (code < 0x300) return false;
  return (
    (code >= 0x300 && code <= 0x36f) ||
    (code >= 0x483 && code <= 0x489) ||
    (code >= 0x591 && code <= 0x5bd) ||
    (code >= 0x610 && code <= 0x61a) ||
    (code >= 0x64b && code <= 0x65f) ||
    (code >= 0x1ab0 && code <= 0x1aff) ||
    (code >= 0x1dc0 && code <= 0x1dff) ||
    code === 0x200c ||
    code === 0x200d ||
    (code >= 0x20d0 && code <= 0x20ff) ||
    (code >= 0xfe00 && code <= 0xfe0f) ||
    (code >= 0xfe20 && code <= 0xfe2f)
  );
}

/**
 * The offset one character after `offset` in `text`: a surrogate pair and a
 * base character with its marks are one step. A practical subset of the
 * Unicode grapheme rules.
 */
export function nextCharOffset(text: string, offset: number): number {
  if (offset >= text.length) return text.length;
  let at = offset + (isHigh(text.charCodeAt(offset)) && isLow(text.charCodeAt(offset + 1)) ? 2 : 1);
  // A long run of marks is one character however long it is; the loop ends with the text.
  while (at < text.length) {
    const code = text.charCodeAt(at);
    if (isExtender(code)) {
      // What a joiner joins belongs to the same character.
      at += code === 0x200d && at + 1 < text.length ? (isHigh(text.charCodeAt(at + 1)) && isLow(text.charCodeAt(at + 2)) ? 3 : 2) : 1;
    } else break;
  }
  return Math.min(at, text.length);
}

/** The offset one character before `offset`. */
export function previousCharOffset(text: string, offset: number): number {
  let at = Math.min(offset, text.length);
  while (at > 0) {
    at -= at > 1 && isLow(text.charCodeAt(at - 1)) && isHigh(text.charCodeAt(at - 2)) ? 2 : 1;
    if (at === 0) break;
    const code = text.charCodeAt(at);
    // Still on a mark, or on what a joiner before it joins: the character starts further back.
    if (!isExtender(code) && text.charCodeAt(at - 1) !== 0x200d) break;
  }
  return at;
}

/** The nearest offset at or after `offset` that is not inside a character. */
export function snapOffset(text: string, offset: number): number {
  if (offset <= 0) return 0;
  if (offset >= text.length) return text.length;
  let start = offset;
  while (start > 0) {
    const code = text.charCodeAt(start);
    if ((isLow(code) && isHigh(text.charCodeAt(start - 1))) || isExtender(code) || text.charCodeAt(start - 1) === 0x200d) start--;
    else break;
  }
  return start === offset ? offset : nextCharOffset(text, start);
}

const SPACE = 0;
const WORD = 1;
const PUNCT = 2;

function charClass(code: number): number {
  if (code <= 0x20 || code === 0xa0 || code === 0x3000 || (code >= 0x2000 && code <= 0x200a)) return SPACE;
  if (code < 0x80) {
    return (code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95 || code === 39 ? WORD : PUNCT;
  }
  // General punctuation and symbols; everything else beyond ASCII counts as a letter.
  if ((code >= 0xa1 && code <= 0xbf && code !== 0xaa && code !== 0xb5 && code !== 0xba) || (code >= 0x2010 && code <= 0x205e) || (code >= 0x3001 && code <= 0x303f)) return PUNCT;
  return WORD;
}

/** Start of the word after `offset`: past the rest of this word and the spaces after it. */
export function nextWordOffset(text: string, offset: number): number {
  let at = offset;
  if (at >= text.length) return text.length;
  const first = charClass(text.charCodeAt(at));
  if (first !== SPACE) while (at < text.length && charClass(text.charCodeAt(at)) === first) at++;
  while (at < text.length && charClass(text.charCodeAt(at)) === SPACE) at++;
  return at;
}

/** Start of the word before `offset`: past the spaces before it and back to where that word begins. */
export function previousWordOffset(text: string, offset: number): number {
  let at = Math.min(offset, text.length);
  while (at > 0 && charClass(text.charCodeAt(at - 1)) === SPACE) at--;
  if (at === 0) return 0;
  const kind = charClass(text.charCodeAt(at - 1));
  while (at > 0 && charClass(text.charCodeAt(at - 1)) === kind) at--;
  return at;
}

/** The word around an offset as [start, end): what a double click selects. */
export function wordAt(text: string, offset: number): [start: number, end: number] {
  if (text.length === 0) return [0, 0];
  const at = Math.min(Math.max(0, offset), text.length - 1);
  const kind = charClass(text.charCodeAt(at));
  let start = at;
  let end = at + 1;
  while (start > 0 && charClass(text.charCodeAt(start - 1)) === kind) start--;
  while (end < text.length && charClass(text.charCodeAt(end)) === kind) end++;
  return [start, end];
}

// ------------------------------------------------------------------ paragraphs

/** Path of the first (step 1) or last (step -1) paragraph in a list of blocks from `from` on; tables are entered. `base` is the path of the list. */
export function paragraphIn(blocks: Block[], base: number[], from: number, step: 1 | -1, depth = 0): number[] | undefined {
  if (depth > MAX_PATH) return undefined;
  for (let i = from; i >= 0 && i < blocks.length; i += step) {
    const block = blocks[i];
    if (block.type === 'paragraph') return base.concat([i]);
    if (block.type === 'table') {
      const found = paragraphInTable(block, base.concat([i]), step === 1 ? 0 : block.rows.length - 1, step === 1 ? 0 : -1, step, depth + 1);
      if (found) return found;
    }
  }
  return undefined;
}

/** The same inside a table, starting at a cell (`cell` -1: the last one of the row). */
function paragraphInTable(table: Table, base: number[], row: number, cell: number, step: 1 | -1, depth: number): number[] | undefined {
  for (let r = row; r >= 0 && r < table.rows.length; r += step) {
    const cells = table.rows[r].cells;
    let c = r === row ? (cell < 0 ? cells.length - 1 : cell) : step === 1 ? 0 : cells.length - 1;
    for (; c >= 0 && c < cells.length; c += step) {
      const blocks = cells[c].blocks;
      const found = paragraphIn(blocks, base.concat([r, c]), step === 1 ? 0 : blocks.length - 1, step, depth + 1);
      if (found) return found;
    }
  }
  return undefined;
}

/**
 * Path of the paragraph after (step 1) or before (step -1) the one at `path`,
 * in document order: through the cells of tables, nested ones included. Stays
 * inside the text box a path is in. Undefined at the end of the story.
 */
export function adjacentParagraph(doc: DocModel, story: string, path: number[], step: 1 | -1): number[] | undefined {
  const top = storyOf(doc, story).blocks;
  const stop = rootLength(path);
  let at = path.slice();
  // Each turn either finds a paragraph or leaves one level of the path.
  for (let guard = 0; guard <= MAX_PATH && at.length > stop; guard++) {
    const container = at.slice(0, -1);
    const blocks = blocksAt(top, container);
    if (!blocks) return bad('The path does not lead to a paragraph');
    const found = paragraphIn(blocks, container, at[at.length - 1] + step, step, 0);
    if (found) return found;
    if (container.length <= stop) return undefined;
    // Out of the cell: on with the next cell of the table.
    const tablePath = container.slice(0, -2);
    const tables = blocksAt(top, tablePath.slice(0, -1));
    const table = tables && tables[tablePath[tablePath.length - 1]];
    if (!table || table.type !== 'table') return bad('The path does not lead to a paragraph');
    const row = container[container.length - 2];
    const cell = container[container.length - 1] + step;
    let inTable: number[] | undefined;
    if (cell >= 0 && cell < table.rows[row].cells.length) inTable = paragraphInTable(table, tablePath, row, cell, step, 0);
    else if (row + step >= 0 && row + step < table.rows.length) inTable = paragraphInTable(table, tablePath, row + step, step === 1 ? 0 : -1, step, 0);
    if (inTable) return inTable;
    at = tablePath;
  }
  return undefined;
}

/** Start of the first paragraph of a story; undefined when it has none. */
export function storyStart(doc: DocModel, story: string): DocPosition | undefined {
  const path = paragraphIn(storyOf(doc, story).blocks, [], 0, 1, 0);
  return path && { story, path, offset: 0 };
}

/** End of the last paragraph of a story. */
export function storyEnd(doc: DocModel, story: string): DocPosition | undefined {
  const blocks = storyOf(doc, story).blocks;
  const path = paragraphIn(blocks, [], blocks.length - 1, -1, 0);
  return path && { story, path, offset: resolvePosition(doc, { story, path, offset: 0 }).length };
}

export type MoveUnit = 'char' | 'word' | 'paragraphStart' | 'paragraphEnd' | 'paragraph' | 'story';

/**
 * The position one unit before (direction -1) or after (1) a position.
 * "char" and "word" go on into the neighbouring paragraph at the ends of
 * one; "paragraph" goes to the start of the next or previous paragraph;
 * "story" to the very start or end. At the end of the story the position
 * comes back unchanged.
 */
export function movePosition(doc: DocModel, position: DocPosition, unit: MoveUnit, direction: 1 | -1): DocPosition {
  const at = resolvePosition(doc, position);
  const { story, path } = position;
  const here = (offset: number): DocPosition => ({ story, path: path.slice(), offset });
  const into = (step: 1 | -1, end: boolean): DocPosition | undefined => {
    const next = adjacentParagraph(doc, story, path, step);
    if (!next) return undefined;
    return { story, path: next, offset: end ? resolvePosition(doc, { story, path: next, offset: 0 }).length : 0 };
  };
  if (unit === 'paragraphStart') return here(0);
  if (unit === 'paragraphEnd') return here(at.length);
  if (unit === 'story') {
    if (rootLength(path) === 0) return (direction === 1 ? storyEnd(doc, story) : storyStart(doc, story)) || here(at.offset);
    let last = path;
    // Inside a text box the walk stays in the box, which holds few paragraphs.
    for (let next = adjacentParagraph(doc, story, last, direction); next; next = adjacentParagraph(doc, story, last, direction)) last = next;
    return { story, path: last.slice(), offset: direction === 1 ? resolvePosition(doc, { story, path: last, offset: 0 }).length : 0 };
  }
  if (unit === 'paragraph') {
    if (direction === -1 && at.offset > 0) return here(0);
    return into(direction, false) || here(direction === 1 ? at.length : 0);
  }
  const text = paragraphText(at.paragraph);
  if (direction === 1) {
    if (at.offset >= at.length) return into(1, false) || here(at.length);
    return here(unit === 'word' ? nextWordOffset(text, at.offset) : nextCharOffset(text, at.offset));
  }
  if (at.offset <= 0) return into(-1, true) || here(0);
  return here(unit === 'word' ? previousWordOffset(text, at.offset) : previousCharOffset(text, at.offset));
}

/**
 * Every paragraph from the one at `from` to the one at `to` (both included),
 * in document order, as paths. Throws RangeError when there are more than
 * `limit`.
 */
export function paragraphsBetween(doc: DocModel, story: string, from: number[], to: number[], limit = 2000000): number[][] {
  const out: number[][] = [from.slice()];
  let at = from;
  const same = (a: number[], b: number[]) => a.length === b.length && a.every((value, i) => value === b[i]);
  while (!same(at, to)) {
    const next = adjacentParagraph(doc, story, at, 1);
    if (!next) return bad('The end of the range is not after its start');
    if (out.length >= limit) return bad('The range holds too many paragraphs');
    out.push(next);
    at = next;
  }
  return out;
}
