// What a copy of a part of a document puts on the clipboard, and what the
// editor keeps of it to paste it back as it was.
//
// The HTML is built here, piece by piece, from text that is escaped: nothing
// of a document reaches the clipboard as markup of its own. HTML that comes
// from the clipboard is never parsed; it is only searched for the mark that
// says the copy was made by an editor of this page.
//
// The copy itself (pictures, links, fields, tables: a DocFragment) never
// leaves the page. It is kept here, for every editor of the page, under the
// mark its HTML carries; a paste whose HTML has the mark and whose text is
// the text of that copy gets the copy, any other paste the plain text of the
// clipboard. Only the last few copies are kept, up to a number of bytes.

import { findWrapper, paragraphText, type Block, type DocModel, type Paragraph, type ParaProps, type Wrapper } from '../model/doc/types';
import {
  applyDocOps,
  asRange,
  deleteRange,
  fragmentOf,
  itemLength,
  paragraphsBetween,
  rangeFormat,
  resolvePosition,
  storyOf,
  type DocFragment,
  type DocPosition,
  type DocRange,
  type RunFormatPatch,
} from '../ops/doc/index';

/** Text in one formatting. A line break inside a paragraph is "\v", as insertText() takes it. */
export interface SliceRun {
  text: string;
  format: RunFormatPatch;
}

export interface SliceParagraph {
  runs: SliceRun[];
  /** Alignment and style of the paragraph, for the paragraphs whose end is part of the copy. */
  align?: string;
  style?: string;
}

/**
 * A part of a document as a copy. `paragraphs` (the last one without its end)
 * and `table` are its text in its formatting: what the HTML on the clipboard
 * is made of. `fragment` is the copy the editor pastes: the content itself,
 * with its pictures, links, fields and tables. A slice without one is pasted
 * as formatted text.
 */
export interface DocSlice {
  paragraphs: SliceParagraph[];
  /** The cells of a table that were selected, by row. */
  table?: string[][];
  fragment?: DocFragment;
}

/** A copy larger than this is kept as plain text only. */
const MAX_SLICE_PARAGRAPHS = 5000;
const MAX_SLICE_RUNS = 20000;

const escapeHtml = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A font name as it may stand in a style attribute: letters, digits, spaces and hyphens only. */
const safeFont = (name: string): string => name.replace(/[^A-Za-z0-9 \-À-￿]/g, '').slice(0, 64);

/** The runs of a paragraph between two offsets, each with the formatting in effect for it. */
function runsOf(doc: DocModel, story: string, path: number[], paragraph: Paragraph, from: number, to: number, budget: { runs: number }): SliceRun[] | undefined {
  const text = paragraphText(paragraph);
  const out: SliceRun[] = [];
  let at = 0;
  const add = (start: number, end: number): boolean => {
    const a = Math.max(start, from);
    const b = Math.min(end, to);
    if (b <= a) return true;
    if (--budget.runs < 0) return false;
    const values = rangeFormat(doc, { anchor: { story, path, offset: a }, focus: { story, path, offset: b } });
    const format: RunFormatPatch = {
      bold: values.bold === true,
      italic: values.italic === true,
      underline: values.underline === true,
      strike: values.strike === true,
    };
    if (values.font !== undefined) format.font = values.font;
    if (values.size !== undefined) format.size = values.size;
    if (values.color !== undefined) format.color = values.color;
    if (values.highlight !== undefined) format.highlight = values.highlight;
    if (values.vertAlign !== undefined) format.vertAlign = values.vertAlign;
    const piece = text.slice(a, b).replace(/\n/g, '\v');
    const last = out[out.length - 1];
    if (last && JSON.stringify(last.format) === JSON.stringify(format)) last.text += piece;
    else out.push({ text: piece, format });
    return true;
  };
  for (const inline of paragraph.children) {
    if (inline.type !== 'run' && inline.type !== 'math') continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    let length = 0;
    if (inline.type === 'math') length = inline.text.length;
    else if (inline.fieldCode !== true) {
      const code = inline.fieldCode;
      for (let k = 0; k < inline.content.length; k++) if (!(code !== undefined && code[k])) length += itemLength(inline.content[k]);
    }
    if (length > 0 && !add(at, at + length)) return undefined;
    at += length;
    if (at >= to) break;
  }
  return out;
}

/**
 * The part of a story between two positions (`start` before `end`) with its
 * formatting; undefined when it is too large to keep or cannot be read.
 */
export function sliceOf(doc: DocModel, start: DocPosition, end: DocPosition): DocSlice | undefined {
  try {
    const paths = paragraphsBetween(doc, start.story, start.path, end.path, MAX_SLICE_PARAGRAPHS);
    const budget = { runs: MAX_SLICE_RUNS };
    const paragraphs: SliceParagraph[] = [];
    for (let i = 0; i < paths.length; i++) {
      const at = resolvePosition(doc, { story: start.story, path: paths[i], offset: 0 });
      const from = i === 0 ? start.offset : 0;
      const to = i === paths.length - 1 ? end.offset : at.length;
      const runs = runsOf(doc, start.story, paths[i], at.paragraph, from, to, budget);
      if (!runs) return undefined;
      const paragraph: SliceParagraph = { runs };
      if (i < paths.length - 1) {
        const values = rangeFormat(doc, { story: start.story, path: paths[i], offset: 0 });
        if (values.align !== undefined) paragraph.align = values.align;
        if (values.paraStyle !== undefined && values.paraStyle !== '') paragraph.style = values.paraStyle;
      }
      paragraphs.push(paragraph);
    }
    const slice: DocSlice = { paragraphs };
    const fragment = fragmentOfRange(doc, start, end);
    if (fragment) slice.fragment = fragment;
    return slice;
  } catch {
    return undefined;
  }
}

/** The copy of a range as the editor pastes it; undefined when it is too large to keep or cannot be read. */
export function fragmentOfRange(doc: DocModel, start: DocPosition, end: DocPosition): DocFragment | undefined {
  try {
    return fragmentOf(doc, { anchor: start, focus: end });
  } catch {
    return undefined;
  }
}

/** What blocks hold, counted: the length of their text, and everything in them that is more than text in some formatting. */
interface Census {
  text: number;
  other: Map<string, number>;
}

/** Items a copy carries as characters of its text. */
const AS_TEXT = new Set<string>(['text', 'tab', 'noBreakHyphen']);
/** Inlines that are never removed with the text around them, or mean nothing by themselves. */
const NO_CONTENT = new Set<string>(['proofError', 'empty']);

const chain = (wrap: Wrapper | undefined): string => {
  let out = '';
  for (let w = wrap; w; w = w.parent) out += w.kind + ' ';
  return out;
};

/** A number for each bag of paragraph properties, so that bags can be counted like everything else. Bags are never changed in place. */
const PROPS_IDS = new WeakMap<ParaProps, number>();
let propsIds = 0;

/** The key paragraph properties are counted under; undefined for properties that state nothing. */
function propsKey(props: ParaProps): string | undefined {
  let stated = false;
  for (const key in props) {
    if (Object.prototype.hasOwnProperty.call(props, key)) {
      stated = true;
      break;
    }
  }
  if (!stated) return undefined;
  let id = PROPS_IDS.get(props);
  if (id === undefined) PROPS_IDS.set(props, (id = ++propsIds));
  return 'paragraph properties ' + id;
}

function census(blocks: Block[]): Census {
  const counts = new Map<string, number>();
  const count = (key: string, by = 1): void => {
    counts.set(key, (counts.get(key) || 0) + by);
  };
  let text = 0;
  const walk = (list: Block[], depth: number): void => {
    for (const block of list) {
      // A block in a content control or another wrapper of blocks is more than the block.
      if (block.wrap && (block.type === 'paragraph' || block.type === 'table')) count('block in ' + chain(block.wrap));
      if (block.type === 'table') {
        count('table');
        for (const row of block.rows) {
          count('row');
          if (row.wrap) count('row in ' + chain(row.wrap));
          for (const cell of row.cells) {
            if (cell.wrap) count('cell in ' + chain(cell.wrap));
            if (depth < 64) walk(cell.blocks, depth + 1);
          }
        }
        continue;
      }
      if (block.type !== 'paragraph') continue;
      text += paragraphText(block).length;
      if (block.props.section) count('section');
      // The formatting of a paragraph (its style, its list, its kept properties) goes with its end.
      const formatted = propsKey(block.props);
      if (formatted !== undefined) count(formatted);
      for (const inline of block.children) {
        // Text inside a link, a field, a content control or a tracked change is more than its characters.
        const wrapped = chain(inline.wrap);
        if (inline.type !== 'run') {
          if (!NO_CONTENT.has(inline.type)) count(inline.type);
          if (inline.type === 'math' && wrapped !== '') count('in ' + wrapped, inline.text.length);
          continue;
        }
        const code = inline.fieldCode;
        for (let k = 0; k < inline.content.length; k++) {
          const item = inline.content[k];
          if (code === true || (code !== undefined && code[k])) {
            count('code');
            continue;
          }
          // A hint of the last layout stands for nothing.
          if (item.type === 'lastRenderedPageBreak') continue;
          if (item.type === 'break') {
            if (item.kind !== 'line') count('break ' + item.kind);
          } else if (!AS_TEXT.has(item.type)) count(item.type);
          if (wrapped !== '') count('in ' + wrapped, Math.max(1, itemLength(item)));
        }
      }
    }
  };
  walk(blocks, 0);
  return { text, other: counts };
}

/**
 * Whether a cut of a range would take something out of the document that the
 * copy of it does not hold (`slice`; undefined when only plain text is on the
 * clipboard): a note reference, an equation, a content control, a tracked
 * change, a drawing with text, and whatever else a copy does not carry (see
 * fragmentOf of ops/doc). The formatting of paragraphs counts: a deletion
 * that joins two paragraphs keeps the formatting of one of them, and the copy
 * has to hold that of the other (a copy that is plain text holds none, so a
 * cut over the end of a formatted paragraph needs a copy that is more). The deletion itself is the judge: it is made, what
 * the story holds is counted before and after, and it is taken back, so that
 * nothing here has to know what a deletion removes. What went is then looked
 * for in the copy, kind by kind: the copy has to hold at least as much of
 * each, and all of the text. Throws the RangeError of a place that cannot be
 * edited.
 */
export function cutLoses(doc: DocModel, range: DocRange, slice: DocSlice | undefined): boolean {
  const story = asRange(range).focus.story;
  const before = census(storyOf(doc, story).blocks);
  const edit = deleteRange(doc, range);
  if (edit.ops.length === 0) return false;
  const inverse = applyDocOps(doc, edit.ops);
  let after: Census;
  try {
    after = census(storyOf(doc, story).blocks);
  } finally {
    applyDocOps(doc, inverse);
  }
  const fragment = slice ? slice.fragment : undefined;
  const carried: Census = fragment ? census(fragment.blocks) : { text: 0, other: new Map() };
  // Of the paragraph a range ends in the copy holds the formatting too, beside the blocks.
  const ending = fragment && fragment.endProps ? propsKey(fragment.endProps) : undefined;
  if (ending !== undefined) carried.other.set(ending, (carried.other.get(ending) || 0) + 1);
  if (!fragment && slice && !slice.table) for (const paragraph of slice.paragraphs) for (const run of paragraph.runs) carried.text += run.text.length;
  let lost = false;
  before.other.forEach((had, key) => {
    if (had - (after.other.get(key) || 0) > (carried.other.get(key) || 0)) lost = true;
  });
  if (lost) return true;
  // Every character that goes must be one the copy holds.
  if (slice && (fragment || !slice.table) && before.text - after.text > carried.text) return true;
  return false;
}

const ALIGN: { [name: string]: string } = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

function runHtml(run: SliceRun): string {
  let html = escapeHtml(run.text).replace(/\t/g, '&#9;').replace(/\v/g, '<br>');
  if (html === '') return '';
  const format = run.format;
  let style = '';
  if (typeof format.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(format.color)) style += `color:#${format.color};`;
  if (typeof format.size === 'number' && format.size > 0 && format.size < 3277) style += `font-size:${format.size / 2}pt;`;
  if (typeof format.font === 'string' && safeFont(format.font) !== '') style += `font-family:'${safeFont(format.font)}';`;
  if (format.vertAlign === 'superscript') html = `<sup>${html}</sup>`;
  else if (format.vertAlign === 'subscript') html = `<sub>${html}</sub>`;
  if (format.strike) html = `<s>${html}</s>`;
  if (format.underline) html = `<u>${html}</u>`;
  if (format.italic) html = `<i>${html}</i>`;
  if (format.bold) html = `<b>${html}</b>`;
  return style ? `<span style="${style}">${html}</span>` : html;
}

/** A slice as HTML: paragraphs with bold, italic, underline, colour and size; a selection of cells as a table. */
export function sliceHtml(slice: DocSlice, mark: string): string {
  let html = '';
  if (slice.table) {
    html += '<table>';
    for (const row of slice.table) {
      html += '<tr>';
      for (const cell of row) html += `<td>${escapeHtml(cell).replace(/\n/g, '<br>')}</td>`;
      html += '</tr>';
    }
    html += '</table>';
  } else {
    for (const paragraph of slice.paragraphs) {
      const align = paragraph.align !== undefined ? ALIGN[paragraph.align] : undefined;
      let inner = '';
      for (const run of paragraph.runs) inner += runHtml(run);
      html += `<p${align ? ` style="text-align:${align}"` : ''}>${inner === '' ? '<br>' : inner}</p>`;
    }
  }
  // The mark is a string of letters and digits made by the editor.
  return `<div data-bym-doc-copy="${escapeHtml(mark)}">${html}</div>`;
}

/** Whether HTML from the clipboard carries the mark of a copy. The HTML is searched, never parsed. */
export function hasCopyMark(html: string | undefined, mark: string): boolean {
  return typeof html === 'string' && mark !== '' && html.indexOf(`data-bym-doc-copy="${mark}"`) >= 0;
}

// ------------------------------------------------------- the copies of the page

/** A copy an editor of the page made, as it is kept for a paste. */
export interface StoredCopy {
  /** What the HTML on the clipboard carries of it. */
  mark: string;
  /** The plain text that went on the clipboard with it. */
  text: string;
  slice: DocSlice;
  /** Bytes the copy is counted with: its text, its blocks, its pictures. */
  bytes: number;
  /** Resolves when the pictures of the copy that were still being read have their bytes (or will not get them). */
  ready?: Promise<void>;
  /** The copy was made by a cut: see InsertFragmentOptions.moved. */
  cut?: boolean;
  /** The copy held more than the page keeps: its pictures were dropped, and none is taken for it any more. */
  full?: boolean;
}

/** How many copies the page keeps, and how many bytes they may hold together (pictures included). */
export const MAX_STORED_COPIES = 8;
export const MAX_STORED_BYTES = 64 * 1024 * 1024;

/** Newest last. */
let copies: StoredCopy[] = [];
let sequence = 0;

const normalText = (text: string): string => text.replace(/\r\n?/g, '\n');

function sizeOf(text: string, slice: DocSlice): number {
  let bytes = 256 + text.length * 2;
  const fragment = slice.fragment;
  if (fragment) {
    bytes += fragment.textLength * 4 + fragment.blocks.length * 128;
    fragment.media.forEach((media) => (bytes += media.bytes ? media.bytes.length : 0));
  }
  return bytes;
}

/** Drops the oldest copies while there are too many or they hold too much; a newest copy that is too large alone loses its pictures, then its fragment. */
function trim(): void {
  let total = 0;
  for (const copy of copies) total += copy.bytes;
  while (copies.length > 1 && (copies.length > MAX_STORED_COPIES || total > MAX_STORED_BYTES)) total -= (copies.shift() as StoredCopy).bytes;
  const last = copies[0];
  if (last && copies.length === 1 && last.bytes > MAX_STORED_BYTES && last.slice.fragment) {
    const fragment = last.slice.fragment;
    // Without the bytes the copy can still be pasted into the document it came from.
    fragment.media.forEach((media) => delete media.bytes);
    last.full = true;
    last.bytes = sizeOf(last.text, last.slice);
    if (last.bytes > MAX_STORED_BYTES) {
      last.slice = { paragraphs: [] };
      last.bytes = 256 + last.text.length * 2;
    }
  }
}

/** Keeps a copy for the editors of the page and returns it; its mark is new and made of letters and digits only. */
export function storeCopy(text: string, slice: DocSlice, cut = false): StoredCopy {
  const mark = Math.random().toString(36).slice(2) + Date.now().toString(36) + (++sequence).toString(36);
  const copy: StoredCopy = { mark, text: normalText(text), slice, bytes: sizeOf(text, slice) };
  if (cut) copy.cut = true;
  copies.push(copy);
  trim();
  return copy;
}

/** Hands a picture that was read for a copy to it: the bytes count against what the page keeps. */
export function setCopyMedia(copy: StoredCopy, name: string, bytes: Uint8Array, type?: string): void {
  const fragment = copy.slice.fragment;
  const media = fragment ? fragment.media.get(name) : undefined;
  if (!media || media.bytes || copy.full || copies.indexOf(copy) < 0) return;
  // A picture that does not fit is not kept (see trim): the copy can then be pasted only where the picture already is.
  media.bytes = bytes;
  if (type !== undefined) media.type = type;
  copy.bytes += bytes.length;
  trim();
}

/** The mark in HTML from the clipboard; undefined when it has none. The HTML is searched, never parsed. */
export function copyMarkOf(html: string | undefined): string | undefined {
  if (typeof html !== 'string') return undefined;
  const key = 'data-bym-doc-copy="';
  const at = html.indexOf(key);
  if (at < 0) return undefined;
  const end = html.indexOf('"', at + key.length);
  const mark = end < 0 ? '' : html.slice(at + key.length, end);
  return /^[a-z0-9]{1,64}$/.test(mark) ? mark : undefined;
}

/** The copy a paste is one of: the HTML carries its mark and the plain text is its text. */
export function storedCopy(html: string | undefined, text: string): StoredCopy | undefined {
  const mark = copyMarkOf(html);
  if (mark === undefined) return undefined;
  for (let i = copies.length - 1; i >= 0; i--) {
    if (copies[i].mark === mark) return copies[i].text === normalText(text) ? copies[i] : undefined;
  }
  return undefined;
}

/** How many copies the page keeps right now and how many bytes they count. */
export function storedCopies(): { count: number; bytes: number } {
  let bytes = 0;
  for (const copy of copies) bytes += copy.bytes;
  return { count: copies.length, bytes };
}

/** Forgets every copy. */
export function clearStoredCopies(): void {
  copies = [];
}
