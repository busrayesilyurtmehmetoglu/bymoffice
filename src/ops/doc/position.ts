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
// bookmark or comment range: what is typed there stays outside of them. A
// position with `before` set stands before what shows without width at its
// offset instead (a picture, a note reference), so that a caret can be put on
// either side of a picture and a range can hold a picture alone.

import { resolveRunProps, type TableContext } from '../../formats/docx/resolve';
import {
  findWrapper,
  paragraphText,
  type Block,
  type DocModel,
  type Note,
  type Paragraph,
  type RunItem,
  type Run,
  type RunProps,
  type Story,
  type StoryCollection,
  type Table,
  type Wrapper,
} from '../../model/doc/types';

export interface DocPosition {
  /** "body", "header:<rId>", "footer:<rId>", "footnote:<id>", "endnote:<id>". */
  story: string;
  /** Path of the paragraph: see TextItem.source of the layout engine. */
  path: number[];
  /** Offset in paragraphText() of the paragraph, in UTF-16 code units. */
  offset: number;
  /**
   * The position is before what shows without width at this offset (an inline
   * picture or drawing, a note reference) instead of after it. Where nothing
   * of the kind sits at the offset the flag says nothing.
   */
  before?: true;
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

/** A drawing (<w:drawing>) or legacy graphic (<w:pict>, <w:object>) of a paragraph, which may hold text boxes. */
export type ParagraphGraphic = Extract<RunItem, { type: 'drawing' | 'pict' | 'object' }>;

/** Where a graphic is: its run, the run item of alternate content it is in (the innermost), and whether that is in alternate content too. */
export interface GraphicPlace {
  run: Run;
  holder?: Extract<RunItem, { type: 'alternate' }>;
  nested: boolean;
}

/**
 * Visits the drawings and legacy graphics of a paragraph in the order the layout engine numbers
 * them (the number is the one a path into a text box names): not those in the code of a field
 * or in deleted text, those of alternate content in its selected branch. Stops when `visit`
 * answers true.
 */
export function eachGraphic(paragraph: Paragraph, visit: (item: ParagraphGraphic, index: number, where: GraphicPlace) => boolean | void): void {
  let seen = 0;
  const walk = (items: RunItem[], code: true | boolean[] | undefined, depth: number, run: Run, holder: GraphicPlace['holder']): boolean => {
    for (let i = 0; i < items.length; i++) {
      if (code !== undefined && (code === true || code[i])) continue;
      const item = items[i];
      if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object') {
        if (visit(item, seen++, { run, holder, nested: depth > 1 }) === true) return true;
      } else if (item.type === 'alternate' && depth < 32 && walk(item.items, undefined, depth + 1, run, item)) return true;
    }
    return false;
  };
  for (const inline of paragraph.children) {
    if (inline.type !== 'run' || inline.fieldCode === true) continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    if (walk(inline.content, inline.fieldCode, 0, inline, undefined)) break;
  }
}

/** The drawing or legacy graphic number `wanted` of a paragraph, counted the way the layout engine numbers them. */
function graphicOf(paragraph: Paragraph, wanted: number): { textBoxes?: Block[][] } | undefined {
  let found: { textBoxes?: Block[][] } | undefined;
  eachGraphic(paragraph, (item, index) => {
    if (index !== wanted) return false;
    found = item;
    return true;
  });
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
export function textBoxRootLength(path: number[]): number {
  for (let i = path.length - 1; i >= 0; i--) if (path[i] === -1) return i + 3;
  return 0;
}

/**
 * Compares two positions of one story in document order: negative when `a`
 * comes first. Throws RangeError for positions of different stories or text
 * boxes.
 */
export function comparePositions(a: DocPosition, b: DocPosition): number {
  if (!a || !b || !Array.isArray(a.path) || !Array.isArray(b.path)) return bad('A position needs a path');
  if (a.story !== b.story) return bad('The positions are in different stories');
  const root = textBoxRootLength(a.path);
  if (root !== textBoxRootLength(b.path)) return bad('The positions are in different text boxes');
  for (let i = 0; i < root; i++) if (a.path[i] !== b.path[i]) return bad('The positions are in different text boxes');
  const shared = Math.min(a.path.length, b.path.length);
  for (let i = root; i < shared; i++) if (a.path[i] !== b.path[i]) return a.path[i] - b.path[i];
  if (a.path.length !== b.path.length) return a.path.length - b.path.length;
  if (a.offset !== b.offset) return a.offset - b.offset;
  // At one offset the position before what has no width there comes first.
  return (a.before === true ? 0 : 1) - (b.before === true ? 0 : 1);
}

/** Do two positions name the same place: story, paragraph, offset and side. */
export function samePosition(a: DocPosition, b: DocPosition): boolean {
  if (!a || !b || !Array.isArray(a.path) || !Array.isArray(b.path)) return bad('A position needs a path');
  return a.story === b.story && a.offset === b.offset && (a.before === true) === (b.before === true) && a.path.length === b.path.length && a.path.every((value, i) => value === b.path[i]);
}

/** The ends of a range in document order. Throws RangeError when they cannot form a range. */
export function orderRange(range: DocRange): { start: DocPosition; end: DocPosition } {
  if (!range || typeof range !== 'object') return bad('A range needs two positions');
  return comparePositions(range.anchor, range.focus) <= 0 ? { start: range.anchor, end: range.focus } : { start: range.focus, end: range.anchor };
}

export function isCollapsed(range: DocRange): boolean {
  if (!range || typeof range !== 'object') return bad('A range needs two positions');
  return samePosition(range.anchor, range.focus);
}

export const collapsed = (position: DocPosition): DocRange => ({ anchor: position, focus: position });

/** A position or a range as a range. Throws RangeError for what is neither. */
export function asRange(at: DocPosition | DocRange): DocRange {
  if (!at || typeof at !== 'object') return bad('A position or a range is needed');
  if (!('anchor' in at) && !('focus' in at)) return collapsed(at);
  const range = at as DocRange;
  if (!range.anchor || !range.focus || typeof range.anchor !== 'object' || typeof range.focus !== 'object') return bad('A range needs two positions');
  return range;
}

// ------------------------------------------------------------------ characters

const isHigh = (code: number) => code >= 0xd800 && code <= 0xdbff;
const isLow = (code: number) => code >= 0xdc00 && code <= 0xdfff;

/** The code point that ends at `i`. */
function pointBefore(text: string, i: number): number {
  const low = text.charCodeAt(i - 1);
  if (isLow(low) && i > 1) {
    const high = text.charCodeAt(i - 2);
    if (isHigh(high)) return (high - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
  }
  return low;
}

/** The code point that begins at `i`. */
function pointAt(text: string, i: number): number {
  const high = text.charCodeAt(i);
  if (isHigh(high)) {
    const low = text.charCodeAt(i + 1);
    if (isLow(low)) return (high - 0xd800) * 0x400 + (low - 0xdc00) + 0x10000;
  }
  return high;
}

/** Marks that belong to the character before them: combining marks, variation selectors, the joiners, skin tones. */
function isExtender(code: number): boolean {
  if (code < 0x300) return false;
  return (
    (code >= 0x300 && code <= 0x36f) ||
    (code >= 0x483 && code <= 0x489) ||
    (code >= 0x591 && code <= 0x5bd) ||
    (code >= 0x610 && code <= 0x61a) ||
    (code >= 0x64b && code <= 0x65f) ||
    code === 0x670 ||
    (code >= 0x6d6 && code <= 0x6dc) ||
    code === 0xe31 ||
    (code >= 0xe34 && code <= 0xe3a) ||
    (code >= 0xe47 && code <= 0xe4e) ||
    (code >= 0x1ab0 && code <= 0x1aff) ||
    (code >= 0x1dc0 && code <= 0x1dff) ||
    code === 0x200c ||
    code === 0x200d ||
    (code >= 0x20d0 && code <= 0x20ff) ||
    (code >= 0x3099 && code <= 0x309a) ||
    (code >= 0xfe00 && code <= 0xfe0f) ||
    (code >= 0xfe20 && code <= 0xfe2f) ||
    (code >= 0x1f3fb && code <= 0x1f3ff) ||
    (code >= 0xe0020 && code <= 0xe007f) ||
    (code >= 0xe0100 && code <= 0xe01ef)
  );
}

/** Regional indicators: two of them are a flag. */
const isRegional = (code: number) => code >= 0x1f1e6 && code <= 0x1f1ff;

const HANGUL_L = 1;
const HANGUL_V = 2;
const HANGUL_T = 3;
const HANGUL_LV = 4;
const HANGUL_LVT = 5;

/** The part a code point plays in a Hangul syllable; 0 for everything else. */
function hangul(code: number): number {
  if (code < 0x1100) return 0;
  if (code <= 0x115f || (code >= 0xa960 && code <= 0xa97c)) return HANGUL_L;
  if ((code >= 0x1160 && code <= 0x11a7) || (code >= 0xd7b0 && code <= 0xd7c6)) return HANGUL_V;
  if ((code >= 0x11a8 && code <= 0x11ff) || (code >= 0xd7cb && code <= 0xd7fb)) return HANGUL_T;
  if (code >= 0xac00 && code <= 0xd7a3) return (code - 0xac00) % 28 === 0 ? HANGUL_LV : HANGUL_LVT;
  return 0;
}

/**
 * May a caret stand at `offset` of `text`: is it not inside a character. A
 * surrogate pair, a base character with its marks, what a joiner joins, a
 * letter with its skin tone or variation selector, the two halves of a flag
 * and the letters of a Hangul syllable are one character each. A practical
 * subset of the Unicode grapheme rules, and the one place they are written
 * down: stepping forward, stepping back and snapping all ask this.
 */
export function isCharBreak(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return true;
  if (isLow(text.charCodeAt(offset)) && isHigh(text.charCodeAt(offset - 1))) return false;
  const prev = pointBefore(text, offset);
  const next = pointAt(text, offset);
  if (prev < 0x300 && next < 0x300) return !(prev === 0x0d && next === 0x0a);
  // A line end or another control character is always a character of its own.
  if (prev < 0x20 || next < 0x20) return true;
  if (isExtender(next)) return false;
  // What a joiner joins belongs to the same character.
  if (prev === 0x200d) return false;
  if (isRegional(prev) && isRegional(next)) {
    // Flags pair up from the start of the run of indicators.
    let before = 0;
    for (let i = offset; i >= 2 && isRegional(pointBefore(text, i)); i -= 2) before++;
    return before % 2 === 0;
  }
  const left = hangul(prev);
  if (left !== 0) {
    const right = hangul(next);
    if (left === HANGUL_L) return right === 0 || right === HANGUL_T;
    if (left === HANGUL_LV || left === HANGUL_V) return !(right === HANGUL_V || right === HANGUL_T);
    return right !== HANGUL_T;
  }
  return true;
}

/** Every offset of `text` a caret may stand at, in order: 0, the borders between its characters, its length. */
export function graphemeBreaks(text: string): number[] {
  const out: number[] = [0];
  for (let i = 1; i < text.length; i++) if (isCharBreak(text, i)) out.push(i);
  if (text.length > 0) out.push(text.length);
  return out;
}

/** The offset one character after `offset` in `text` (see isCharBreak for what a character is). */
export function nextCharOffset(text: string, offset: number): number {
  if (offset >= text.length) return text.length;
  let at = Math.max(0, offset) + 1;
  // A long run of marks is one character however long it is; the loop ends with the text.
  while (at < text.length && !isCharBreak(text, at)) at++;
  return at;
}

/** The offset one character before `offset`. */
export function previousCharOffset(text: string, offset: number): number {
  let at = Math.min(offset, text.length) - 1;
  while (at > 0 && !isCharBreak(text, at)) at--;
  return Math.max(0, at);
}

/** The nearest offset at or after `offset` that is not inside a character. */
export function snapOffset(text: string, offset: number): number {
  if (offset <= 0) return 0;
  let at = offset;
  while (at < text.length && !isCharBreak(text, at)) at++;
  return Math.min(at, text.length);
}

const SPACE = 0;
const WORD = 1;
const PUNCT = 2;
/** Ideographs and kana: written without spaces, so each is a word of its own. */
const IDEOGRAPH = 3;

function pointClass(code: number): number {
  if (code <= 0x20 || code === 0xa0 || code === 0x3000 || (code >= 0x2000 && code <= 0x200a)) return SPACE;
  if (code < 0x80) {
    return (code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95 ? WORD : PUNCT;
  }
  // General punctuation and symbols; everything else beyond ASCII counts as a letter.
  if ((code >= 0xa1 && code <= 0xbf && code !== 0xaa && code !== 0xb5 && code !== 0xba) || (code >= 0x2010 && code <= 0x205e) || (code >= 0x3001 && code <= 0x303f)) return PUNCT;
  if (
    (code >= 0x3040 && code <= 0x30ff) ||
    (code >= 0x3400 && code <= 0x4dbf) ||
    (code >= 0x4e00 && code <= 0x9fff) ||
    (code >= 0xf900 && code <= 0xfaff) ||
    (code >= 0xff66 && code <= 0xff9f) ||
    (code >= 0x20000 && code <= 0x3ffff)
  ) {
    return isExtender(code) ? WORD : IDEOGRAPH;
  }
  return WORD;
}

const isApostrophe = (code: number) => code === 39 || code === 0x2019;

/** The class of the character `offset` is in: of its base, whatever marks follow it. */
function classAt(text: string, offset: number): number {
  let start = offset;
  while (start > 0 && !isCharBreak(text, start)) start--;
  const code = pointAt(text, start);
  if (!isApostrophe(code)) return pointClass(code);
  // An apostrophe between two letters is part of the word ("don't", "İstanbul’da"); anywhere else it is a quote.
  if (start === 0 || start + 1 >= text.length) return PUNCT;
  const before = pointBefore(text, start);
  const after = pointAt(text, start + 1);
  return pointClass(before) === WORD && pointClass(after) === WORD ? WORD : PUNCT;
}

/** Does a word, a run of spaces or a run of punctuation begin at `offset` (0 < offset < length). */
function isWordBreak(text: string, offset: number): boolean {
  if (!isCharBreak(text, offset)) return false;
  const kind = classAt(text, offset);
  return kind === IDEOGRAPH || kind !== classAt(text, offset - 1);
}

/** Start of the word after `offset`: past the rest of this word and the spaces after it. */
export function nextWordOffset(text: string, offset: number): number {
  for (let at = Math.max(0, offset) + 1; at < text.length; at++) {
    if (isWordBreak(text, at) && classAt(text, at) !== SPACE) return at;
  }
  return text.length;
}

/** Start of the word before `offset`: past the spaces before it and back to where that word begins. */
export function previousWordOffset(text: string, offset: number): number {
  for (let at = Math.min(offset, text.length) - 1; at > 0; at--) {
    if (isWordBreak(text, at) && classAt(text, at) !== SPACE) return at;
  }
  return 0;
}

/** The word around an offset as [start, end): what a double click selects. */
export function wordAt(text: string, offset: number): [start: number, end: number] {
  if (text.length === 0) return [0, 0];
  const at = Math.min(Math.max(0, offset), text.length - 1);
  let start = at;
  let end = at + 1;
  while (start > 0 && !isWordBreak(text, start)) start--;
  while (end < text.length && !isWordBreak(text, end)) end++;
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
  const stop = textBoxRootLength(path);
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

/**
 * Items of a run that show but take no offset of paragraphText(): a caret can
 * stand before or after them (DocPosition.before), a deletion takes them as a
 * unit of their own.
 */
export const ATOM_ITEMS = new Set<string>(['drawing', 'pict', 'object', 'footnoteReference', 'endnoteReference', 'separator', 'continuationSeparator', 'pageNumber', 'positionalTab']);

/** Does something that shows without width (a picture, a note reference) sit at an offset of a paragraph. */
export function hasAtomAt(paragraph: Paragraph, offset: number): boolean {
  let at = 0;
  for (const inline of paragraph.children) {
    if (at > offset) return false;
    if (inline.type === 'math') {
      if (!notShown(inline)) at += inline.text.length;
      continue;
    }
    if (inline.type !== 'run' || inline.fieldCode === true || notShown(inline)) continue;
    const code = inline.fieldCode;
    // Content of mc:AlternateContent is kept whole: nothing in it is a unit of its own.
    const whole = inline.wrap !== undefined && findWrapper(inline, 'alternate') !== undefined;
    for (let k = 0; k < inline.content.length; k++) {
      if (code !== undefined && code[k]) continue;
      const item = inline.content[k];
      const length = itemLength(item);
      if (length > 0) at += length;
      else if (at === offset && !whole && ATOM_ITEMS.has(item.type)) return true;
    }
  }
  return false;
}

/** Where a paragraph sits when its path goes through a table: what the style resolution needs. */
export function tableContextOf(doc: DocModel, story: string, path: number[]): TableContext | undefined {
  if (path.length < 4 || path[path.length - 4] === -1) return undefined;
  const tablePath = path.slice(0, -3);
  const blocks = containerBlocks(doc, story, tablePath.slice(0, -1));
  const table = blocks[tablePath[tablePath.length - 1]];
  if (!table || table.type !== 'table') return undefined;
  return { table, row: path[path.length - 3], cell: path[path.length - 2] };
}

/** The stretches of a paragraph that are hidden text (w:vanish, from the run or a style), as [start, end) offsets. */
function hiddenRanges(doc: DocModel, paragraph: Paragraph, where: TableContext | undefined): [number, number][] {
  const out: [number, number][] = [];
  const memo = new Map<RunProps, boolean>();
  let at = 0;
  for (const inline of paragraph.children) {
    if (inline.type === 'math') {
      if (!notShown(inline)) at += inline.text.length;
      continue;
    }
    if (inline.type !== 'run' || inline.fieldCode === true || notShown(inline)) continue;
    const length = itemsLength(inline.content, inline.fieldCode, 0);
    if (length === 0) continue;
    let hidden = memo.get(inline.props);
    if (hidden === undefined) memo.set(inline.props, (hidden = resolveRunProps(doc, paragraph.props, inline.props, where).vanish === true));
    if (hidden) {
      const last = out.length ? out[out.length - 1] : undefined;
      if (last && last[1] === at) last[1] = at + length;
      else out.push([at, at + length]);
    }
    at += length;
  }
  return out;
}

export type MoveUnit = 'char' | 'word' | 'paragraphStart' | 'paragraphEnd' | 'paragraph' | 'story';

const MOVE_UNITS: readonly string[] = ['char', 'word', 'paragraphStart', 'paragraphEnd', 'paragraph', 'story'];

export interface MoveOptions {
  /** Stop inside hidden text (w:vanish) too; without it a stretch of hidden text is passed in one step. */
  includeHidden?: boolean;
}

/**
 * The position one unit before (direction -1) or after (1) a position.
 * "char" and "word" go on into the neighbouring paragraph at the ends of
 * one; "paragraph" goes to the start of the next or previous paragraph;
 * "story" to the very start or end. At the end of the story the position
 * comes back unchanged.
 *
 * What shows without width (a picture, a note reference) is a step of its
 * own for "char": first the position before it (`before: true`), then the
 * one after it. Moving forward, and to the start of a paragraph, arrives
 * before such a thing; moving back arrives after it. Several of them at one
 * offset are passed together. Hidden text is passed in one step, and no
 * position inside it is returned, unless `options.includeHidden` is set.
 */
export function movePosition(doc: DocModel, position: DocPosition, unit: MoveUnit, direction: 1 | -1, options?: MoveOptions): DocPosition {
  if (direction !== 1 && direction !== -1) return bad('The direction of a move is 1 or -1');
  if (MOVE_UNITS.indexOf(unit) < 0) return bad('The unit of a move is not valid');
  const at = resolvePosition(doc, position);
  const { story, path } = position;
  const here = (offset: number): DocPosition => ({ story, path: path.slice(), offset });
  /** The position at an offset of a paragraph that is before what has no width there. */
  const first = (to: number[], paragraph: Paragraph, offset: number): DocPosition => (hasAtomAt(paragraph, offset) ? { story, path: to, offset, before: true } : { story, path: to, offset });
  const unchanged = (): DocPosition => (position.before === true ? first(path.slice(), at.paragraph, at.offset) : here(at.offset));
  const into = (step: 1 | -1, end: boolean): DocPosition | undefined => {
    const next = adjacentParagraph(doc, story, path, step);
    if (!next) return undefined;
    const target = resolvePosition(doc, { story, path: next, offset: 0 });
    return end ? { story, path: next, offset: target.length } : first(next, target.paragraph, 0);
  };
  const startOf = (to: number[]): DocPosition => first(to.slice(), resolvePosition(doc, { story, path: to, offset: 0 }).paragraph, 0);
  if (unit === 'paragraphStart') return first(path.slice(), at.paragraph, 0);
  if (unit === 'paragraphEnd') return here(at.length);
  if (unit === 'story') {
    if (textBoxRootLength(path) === 0) {
      const edge = direction === 1 ? storyEnd(doc, story) : storyStart(doc, story);
      return edge ? (direction === 1 ? edge : startOf(edge.path)) : unchanged();
    }
    let last = path;
    // Inside a text box the walk stays in the box, which holds few paragraphs.
    for (let next = adjacentParagraph(doc, story, last, direction); next; next = adjacentParagraph(doc, story, last, direction)) last = next;
    return direction === 1 ? { story, path: last.slice(), offset: resolvePosition(doc, { story, path: last, offset: 0 }).length } : startOf(last);
  }
  if (unit === 'paragraph') {
    if (direction === -1 && at.offset > 0) return first(path.slice(), at.paragraph, 0);
    return into(direction, false) || (direction === 1 ? here(at.length) : first(path.slice(), at.paragraph, 0));
  }
  const atom = unit === 'char' && hasAtomAt(at.paragraph, at.offset);
  // Over the picture at the offset, from one side of it to the other.
  if (atom && direction === 1 && position.before === true) return here(at.offset);
  if (atom && direction === -1 && position.before !== true) return { ...here(at.offset), before: true };
  const text = paragraphText(at.paragraph);
  const hidden = options && options.includeHidden ? [] : hiddenRanges(doc, at.paragraph, tableContextOf(doc, story, path));
  const shown = (offset: number): number => {
    for (const [from, to] of hidden) if (offset > from && offset < to) return direction === 1 ? to : from;
    return offset;
  };
  if (direction === 1) {
    if (at.offset >= at.length) return into(1, false) || unchanged();
    return first(path.slice(), at.paragraph, shown(unit === 'word' ? nextWordOffset(text, at.offset) : nextCharOffset(text, at.offset)));
  }
  if (at.offset <= 0) return into(-1, true) || unchanged();
  return here(shown(unit === 'word' ? previousWordOffset(text, at.offset) : previousCharOffset(text, at.offset)));
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
