// The content of a paragraph as a flat list of pieces, which is what the
// editing commands work on: one piece per item of a run, one per inline that
// is not a run. Cutting, removing and inserting happen on that list; rebuild()
// then turns it back into runs, keeping every run nothing happened to as the
// object it was and joining neighbours that are formatted alike.

import { isLiveField, parseFieldCode } from '../../layout/doc/symbols';
import {
  EMPTY_RUN_PROPS,
  type AttrList,
  type Block,
  type ExtraXml,
  type Inline,
  type MiscXml,
  type Paragraph,
  type RangeMark,
  type Run,
  type RunItem,
  type RunProps,
  type Table,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import { ATOM_ITEMS, itemLength, itemsLength, snapOffset } from './position';

export const READ_ONLY = 'This part of the document cannot be edited';

export interface Piece {
  /** An item of a run; `props`, `group`, `index` and `code` describe the run it is in. */
  item?: RunItem;
  /** Or an inline that is kept whole: everything that is not a run, and runs that are not shown or cannot be edited. */
  inline?: Inline;
  props: RunProps;
  wrap?: Wrapper;
  /** Pieces with the same group (and formatting) that follow each other become one run. The run they came from, or a token. */
  group: object;
  /** Index in the content of the run the piece came from; -1 for a piece made by an edit. */
  index: number;
  /** The item sits in a field code: it is not shown and has no offsets. */
  code: boolean;
  /** Offsets the piece takes in paragraphText(). */
  len: number;
  /** Pieces with the same lock are one unit: the cached result of a field the engine computes (PAGE, NUMPAGES). */
  lock?: object;
  /** The piece cannot be removed or reformatted: it sits in an element that is written back whole. */
  frozen?: boolean;
}

function chainHas(wrap: Wrapper | undefined, test: (wrapper: Wrapper) => boolean): boolean {
  for (let w = wrap; w; w = w.parent) if (test(w)) return true;
  return false;
}

const isRevisionHidden = (w: Wrapper) => w.kind === 'del' || w.kind === 'moveFrom';
const isAlternate = (w: Wrapper) => w.kind === 'alternate';

const LIVE_SIMPLE = new WeakMap<Wrapper, boolean>();

/** A simple field whose result the engine computes. */
function isLiveSimple(w: Wrapper): boolean {
  if (w.kind !== 'simpleField') return false;
  let live = LIVE_SIMPLE.get(w);
  if (live === undefined) LIVE_SIMPLE.set(w, (live = isLiveField(parseFieldCode(w.instruction).name)));
  return live;
}

/** The pieces of a paragraph, with the units that cannot be cut marked. */
export function flatten(paragraph: Paragraph): Piece[] {
  const pieces: Piece[] = [];
  for (const inline of paragraph.children) {
    const wrap = inline.wrap;
    const hidden = wrap !== undefined && chainHas(wrap, isRevisionHidden);
    if (inline.type === 'run' && !hidden && inline.content.length > 0 && !(wrap && chainHas(wrap, isAlternate))) {
      const content = inline.content;
      const flags = inline.fieldCode;
      for (let k = 0; k < content.length; k++) {
        const code = flags === true || (flags !== undefined && flags[k] === true);
        pieces.push({ item: content[k], props: inline.props, wrap, group: inline, index: k, code, len: code ? 0 : itemLength(content[k]) });
      }
      continue;
    }
    let len = 0;
    if (!hidden) {
      if (inline.type === 'math') len = inline.text.length;
      else if (inline.type === 'run' && inline.fieldCode !== true) len = itemsLength(inline.content, inline.fieldCode, 0);
    }
    const piece: Piece = { inline, props: EMPTY_RUN_PROPS, wrap, group: inline, index: -1, code: false, len };
    if (hidden || (wrap && chainHas(wrap, isAlternate))) piece.frozen = true;
    pieces.push(piece);
  }
  markLocks(pieces);
  return pieces;
}

export interface FieldPair {
  begin: number;
  /** -1 when the field has no separator. */
  separate: number;
  /** -1 when the field does not end in the list. */
  end: number;
  live: boolean;
}

/** The complex fields that begin in a list of pieces, and for each piece the innermost of them it belongs to (-1: none). */
export function fieldPairs(pieces: Piece[]): { pairs: FieldPair[]; owner: number[] } {
  const pairs: FieldPair[] = [];
  const owner: number[] = [];
  const stack: number[] = [];
  for (let i = 0; i < pieces.length; i++) {
    const item = pieces[i].item;
    if (item && item.type === 'fieldChar') {
      if (item.charType === 'begin') {
        const field = item.field;
        const live = !pieces[i].code && field !== undefined && !field.unclosed && isLiveField(parseFieldCode(field.instruction).name);
        stack.push(pairs.length);
        pairs.push({ begin: i, separate: -1, end: -1, live });
        owner.push(pairs.length - 1);
        continue;
      }
      const top = stack.length ? stack[stack.length - 1] : -1;
      owner.push(top);
      if (top >= 0) {
        if (item.charType === 'separate') {
          if (pairs[top].separate < 0) pairs[top].separate = i;
        } else {
          pairs[top].end = i;
          stack.pop();
        }
      }
      continue;
    }
    owner.push(stack.length ? stack[stack.length - 1] : -1);
  }
  return { pairs, owner };
}

function markLocks(pieces: Piece[]): void {
  let fields = false;
  for (const piece of pieces) {
    if (piece.item && piece.item.type === 'fieldChar') fields = true;
    if (piece.wrap) {
      for (let w: Wrapper | undefined = piece.wrap; w; w = w.parent) if (isLiveSimple(w)) piece.lock = w;
    }
  }
  if (!fields) return;
  // Pairs are in the order they begin, so an outer field marks its pieces before the fields inside it.
  for (const pair of fieldPairs(pieces).pairs) {
    if (!pair.live || pair.end < 0) continue;
    for (let i = pair.begin; i <= pair.end; i++) if (pieces[i].lock === undefined) pieces[i].lock = pair;
  }
}

const RANGE_STARTS = new Set<string>(['bookmarkStart', 'commentRangeStart', 'moveFromRangeStart', 'moveToRangeStart', 'permStart']);

/** May a caret step over this piece without width: not over the beginning of a field or of a marked range. */
function passable(piece: Piece): boolean {
  if (piece.item) return !(piece.item.type === 'fieldChar' && piece.item.charType === 'begin');
  return !RANGE_STARTS.has((piece.inline as Inline).type);
}

/**
 * How an offset becomes a place between two pieces:
 *   "caret"  after what has no width at the offset, but not into a field or marked range
 *            (and, with `beforeAtoms`, not past what shows without width: a picture, a note reference);
 *   "start"  the same, for the start of a range: inside a unit that cannot be cut it is the start of the unit;
 *   "after"  directly after the character before the offset;
 *   "before" directly before the character at the offset.
 */
export type BoundaryMode = 'caret' | 'start' | 'after' | 'before';

/**
 * The index of the piece an offset falls before. A text piece the offset is
 * inside of is cut in two (the list changes, the offsets of its pieces do
 * not). An offset inside a unit that cannot be cut moves to its end, or to
 * its start for "start" and "before". `beforeAtoms` is DocPosition.before:
 * it matters for "caret" and "start" only.
 */
export function boundary(pieces: Piece[], offset: number, mode: BoundaryMode, beforeAtoms = false): number {
  let acc = 0;
  let k = 0;
  for (; k < pieces.length; k++) {
    const len = pieces[k].len;
    if (len > 0 && acc + len > offset) break;
    acc += len;
  }
  const back = mode === 'start' || mode === 'before';
  let i: number;
  if (k < pieces.length && acc < offset) {
    const piece = pieces[k];
    const item = piece.item;
    if (item && item.type === 'text' && !piece.lock && !piece.frozen) {
      const cut = snapOffset(item.text, offset - acc);
      if (cut < item.text.length) {
        const left: Piece = { ...piece, item: { type: 'text', text: item.text.slice(0, cut) }, index: -1, len: cut };
        const right: Piece = { ...piece, item: { type: 'text', text: item.text.slice(cut) }, index: -1, len: item.text.length - cut };
        pieces.splice(k, 1, left, right);
      }
      i = k + 1;
    } else i = back ? k : k + 1;
  } else {
    let first = k;
    while (first > 0 && pieces[first - 1].len === 0) first--;
    if (mode === 'after') i = first;
    else if (mode === 'before') i = k;
    else {
      i = first;
      while (i < k && passable(pieces[i]) && !(beforeAtoms && isAtom(pieces[i]))) i++;
    }
  }
  if (i > 0 && i < pieces.length && pieces[i - 1].lock !== undefined && pieces[i - 1].lock === pieces[i].lock) {
    const lock = pieces[i].lock;
    if (back) while (i > 0 && pieces[i - 1].lock === lock) i--;
    else while (i < pieces.length && pieces[i].lock === lock) i++;
  }
  return i;
}

/** Where a caret at a position stands among the pieces. */
export const caretIndex = (pieces: Piece[], position: { offset: number; before?: true }): number => boundary(pieces, position.offset, 'caret', position.before === true);

/** Where a range that starts at a position begins among the pieces. */
export const startIndex = (pieces: Piece[], position: { offset: number; before?: true }): number => boundary(pieces, position.offset, 'start', position.before === true);

/** Does something that shows without width stand between the piece at `index` and the next piece with width. */
export function atomFollows(pieces: Piece[], index: number): boolean {
  for (let i = index; i < pieces.length && pieces[i].len === 0; i++) if (isAtom(pieces[i])) return true;
  return false;
}

/** The offset at which the piece at `index` begins. */
export function offsetOf(pieces: Piece[], index: number): number {
  let offset = 0;
  for (let i = 0; i < index && i < pieces.length; i++) offset += pieces[i].len;
  return offset;
}

const REMOVABLE = new Set<string>([
  'text',
  'tab',
  'break',
  'noBreakHyphen',
  'softHyphen',
  'symbol',
  'alternate',
  'positionalTab',
  'pageNumber',
  'lastRenderedPageBreak',
  'drawing',
  'pict',
  'object',
  'footnoteReference',
  'endnoteReference',
  'separator',
  'continuationSeparator',
]);

/** Content a deletion takes. Marks of ranges, comments, unknown elements and everything not shown stay. */
export function isRemovable(piece: Piece): boolean {
  if (piece.frozen || piece.code) return false;
  if (piece.inline) return piece.inline.type === 'math';
  return REMOVABLE.has((piece.item as RunItem).type);
}

/**
 * Shows, but without width in the offsets, and so is taken by Backspace and
 * Delete as a unit of its own: a drawing, a note reference. A soft hyphen and
 * the hints of the last layout are not: they do not show.
 */
export const isAtom = (piece: Piece): boolean => piece.len === 0 && piece.item !== undefined && !piece.frozen && !piece.code && ATOM_ITEMS.has(piece.item.type);

/**
 * A hint an earlier layout or spelling check left behind (where a page broke,
 * where a misspelt word began): stale once the text beside it changes, so
 * Backspace and Delete drop the ones they pass.
 */
export const isStaleHint = (piece: Piece): boolean =>
  !piece.frozen && (piece.item !== undefined ? !piece.code && piece.item.type === 'lastRenderedPageBreak' : (piece.inline as Inline).type === 'proofError');

/** Part of the skeleton of a field: its marks and its code. */
export const isStructural = (piece: Piece): boolean => piece.item !== undefined && (piece.code || piece.item.type === 'fieldChar' || piece.item.type === 'instrText');

/** A wrapper chain without the wrapper `gone`; wrappers inside it are copied, once each. */
function withoutWrapper(wrap: Wrapper | undefined, gone: Wrapper, memo: Map<Wrapper, Wrapper | undefined>): Wrapper | undefined {
  if (!wrap) return undefined;
  if (wrap === gone) return wrap.parent;
  if (memo.has(wrap)) return memo.get(wrap);
  let inside = false;
  for (let w = wrap.parent; w; w = w.parent) if (w === gone) inside = true;
  const out = inside ? ({ ...wrap, parent: withoutWrapper(wrap.parent, gone, memo) } as Wrapper) : wrap;
  if (out !== wrap && out.parent === undefined) delete out.parent;
  memo.set(wrap, out);
  return out;
}

/**
 * The pieces without the content of [from, to). A field goes as a whole when
 * all of it is in the range or all of its result is; otherwise its marks and
 * code stay. A hyperlink that is left without content goes too, and what it
 * still held (bookmarks) stays in its place.
 */
export function removePieces(pieces: Piece[], from: number, to: number): Piece[] {
  if (from >= to) return pieces;
  const flags = removedPieces(pieces, from, to);
  const out: Piece[] = [];
  const links = new Set<Wrapper>();
  for (let q = 0; q < pieces.length; q++) {
    const piece = pieces[q];
    if (!flags[q]) {
      out.push(piece);
      continue;
    }
    for (let w = piece.wrap; w; w = w.parent) if (w.kind === 'hyperlink') links.add(w);
  }
  links.forEach((link) => {
    let used = false;
    for (const piece of out) {
      if ((piece.len > 0 || isAtom(piece)) && chainHas(piece.wrap, (w) => w === link)) used = true;
    }
    if (used) return;
    const memo = new Map<Wrapper, Wrapper | undefined>();
    for (let q = 0; q < out.length; q++) {
      if (chainHas(out[q].wrap, (w) => w === link)) out[q] = { ...out[q], wrap: withoutWrapper(out[q].wrap, link, memo) };
    }
  });
  return out;
}

/**
 * Which pieces a removal of [from, to) takes, a flag for each: the judgement
 * removePieces() makes, for those who have to know what goes (the copy a cut
 * puts on the clipboard holds exactly these).
 */
export function removedPieces(pieces: Piece[], from: number, to: number): boolean[] {
  const flags: boolean[] = [];
  if (from >= to) {
    for (let q = 0; q < pieces.length; q++) flags.push(false);
    return flags;
  }
  const { pairs, owner } = fieldPairs(pieces);
  const doomed: boolean[] = [];
  for (const pair of pairs) {
    let gone = false;
    if (pair.end >= 0) {
      if (from <= pair.begin && pair.end < to) gone = true;
      else if (pair.separate >= 0 && pair.end > from && pair.separate < to) {
        let any = false;
        gone = true;
        for (let q = pair.separate + 1; q < pair.end && gone; q++) {
          if (pieces[q].len === 0) continue;
          any = true;
          if (q < from || q >= to || !isRemovable(pieces[q])) gone = false;
        }
        if (!any) gone = false;
      }
    }
    doomed.push(gone);
  }
  for (let q = 0; q < pieces.length; q++) {
    const piece = pieces[q];
    flags.push(isStructural(piece) ? owner[q] >= 0 && doomed[owner[q]] && !piece.frozen : q >= from && q < to && isRemovable(piece));
  }
  return flags;
}

/** Characters XML cannot hold: they are dropped from typed and pasted text. */
function isIllegal(text: string, i: number): boolean {
  const code = text.charCodeAt(i);
  if (code < 0x20) return code !== 0x09 && code !== 0x0a && code !== 0x0b && code !== 0x0d;
  if (code < 0xd800) return false;
  if (code <= 0xdbff) {
    const next = text.charCodeAt(i + 1);
    return !(next >= 0xdc00 && next <= 0xdfff);
  }
  if (code <= 0xdfff) {
    const prev = i > 0 ? text.charCodeAt(i - 1) : 0;
    return !(prev >= 0xd800 && prev <= 0xdbff);
  }
  return code === 0xfffe || code === 0xffff;
}

/** Text as it may go into a document: line ends as "\n", without the characters XML cannot hold. */
export function cleanText(text: string): string {
  let out = '';
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    const crlf = code === 0x0d;
    if (!crlf && !isIllegal(text, i)) continue;
    out += text.slice(from, i);
    if (crlf && text.charCodeAt(i + 1) !== 0x0a) out += '\n';
    from = i + 1;
  }
  return from === 0 ? text : out + text.slice(from);
}

const TAB_ITEM: RunItem = Object.freeze({ type: 'tab' as const });
const LINE_BREAK_ITEM: RunItem = Object.freeze({ type: 'break' as const, kind: 'line' as const });

/** Pieces for one line of text: a tab character becomes a tab, a vertical tab (Shift+Enter) a line break. */
export function textPieces(line: string, props: RunProps, wrap: Wrapper | undefined, group: object): Piece[] {
  const out: Piece[] = [];
  const push = (item: RunItem, len: number) => out.push({ item, props, wrap, group, index: -1, code: false, len });
  let from = 0;
  for (let i = 0; i <= line.length; i++) {
    const code = i < line.length ? line.charCodeAt(i) : -1;
    if (code !== -1 && code !== 0x09 && code !== 0x0b) continue;
    if (i > from) push({ type: 'text', text: line.slice(from, i) }, i - from);
    if (code === 0x09) push(TAB_ITEM, 1);
    else if (code === 0x0b) push(LINE_BREAK_ITEM, 1);
    from = i + 1;
  }
  return out;
}

/** A piece for one item that an edit makes. */
export const itemPiece = (item: RunItem, props: RunProps, wrap: Wrapper | undefined, group: object = {}): Piece => ({ item, props, wrap, group, index: -1, code: false, len: itemLength(item) });

/** The innermost wrapper both chains share. */
export function commonWrap(a: Wrapper | undefined, b: Wrapper | undefined): Wrapper | undefined {
  if (!a || !b) return undefined;
  const seen = new Set<Wrapper>();
  for (let w: Wrapper | undefined = a; w; w = w.parent) seen.add(w);
  for (let w: Wrapper | undefined = b; w; w = w.parent) if (seen.has(w)) return w;
  return undefined;
}

/** The wrapper new content takes between two neighbours: the one both are in, outside of anything that is not for new text. */
export function wrapBetween(left: Wrapper | undefined, right: Wrapper | undefined): Wrapper | undefined {
  let wrap = commonWrap(left, right);
  for (let w = wrap; w; w = w.parent) if (isRevisionHidden(w) || isAlternate(w) || isLiveSimple(w)) wrap = w.parent;
  return wrap;
}

const MARK_ONLY = ['inserted', 'deleted', 'moveFrom', 'moveTo', 'change'] as const;

/** Run formatting of a paragraph mark as text takes it over: without what says the mark itself was inserted or deleted. */
export function markRunProps(mark: RunProps | undefined): RunProps {
  if (!mark) return EMPTY_RUN_PROPS;
  let out: RunProps | undefined;
  for (const key of MARK_ONLY) {
    if (mark[key] === undefined) continue;
    if (!out) out = { ...mark };
    delete out[key];
  }
  if (!out) return mark;
  return Object.keys(out).length ? out : EMPTY_RUN_PROPS;
}

/** Wrappers whose formatting belongs to what they hold: text typed beside them does not take it over. */
const isContainer = (w: Wrapper): boolean => w.kind === 'hyperlink' || w.kind === 'simpleField' || w.kind === 'sdt';

/**
 * Formatting text typed before the piece at `index` takes: that of the text
 * to its left (Word's rule at the border between two runs), of the text to
 * its right at the start of the paragraph, of the paragraph mark in an empty
 * paragraph.
 *
 * Text in a hyperlink, a content control or the result of a field hands its
 * formatting on only to what is typed inside the same: after a link the text
 * is formatted like the last text before the link, and like the paragraph
 * mark when there is none (so the blue and the underline end with the link).
 */
export function inheritedProps(pieces: Piece[], index: number, mark: RunProps | undefined): RunProps {
  const wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
  const around = new Set<Wrapper>();
  for (let w = wrap; w; w = w.parent) around.add(w);
  let fields: { pairs: FieldPair[]; owner: number[] } | undefined;
  /** Is the piece in a container the new text is not in. */
  const apart = (i: number): boolean => {
    for (let w = pieces[i].wrap; w; w = w.parent) if (isContainer(w) && !around.has(w)) return true;
    if (!fields) fields = fieldPairs(pieces);
    const owner = fields.owner[i];
    if (owner < 0) return false;
    const pair = fields.pairs[owner];
    return !(pair.begin < index && (pair.end < 0 || index <= pair.end));
  };
  const text = (piece: Piece): boolean => piece.item !== undefined && piece.len > 0 && !piece.frozen;
  let passed = false;
  for (let i = index - 1; i >= 0; i--) {
    if (!text(pieces[i])) continue;
    if (!apart(i)) return pieces[i].props;
    passed = true;
  }
  if (!passed) {
    for (let i = index; i < pieces.length; i++) {
      if (!text(pieces[i])) continue;
      if (apart(i)) break;
      return pieces[i].props;
    }
  }
  return markRunProps(mark);
}

const PLAIN_ITEMS = new Set<string>(['text', 'tab', 'break', 'noBreakHyphen', 'softHyphen', 'symbol']);

/** The group text with this formatting joins when it is put before the piece at `index`: the run beside it when that is formatted alike. */
export function groupFor(pieces: Piece[], index: number, props: RunProps, wrap: Wrapper | undefined): object {
  // Only text joins text: a run that holds the mark of a field or a drawing stays as it is.
  const fits = (piece: Piece | undefined) => piece !== undefined && piece.item !== undefined && !piece.code && PLAIN_ITEMS.has(piece.item.type) && piece.props === props && piece.wrap === wrap;
  if (fits(pieces[index - 1])) return pieces[index - 1].group;
  if (fits(pieces[index])) return pieces[index].group;
  return {};
}

/** Structural equality of property bags. Kept XML trees and maps are equal only to themselves. */
export function sameProps(a: unknown, b: unknown, depth = 0): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null || depth > 8) return false;
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!sameProps(a[i], b[i], depth + 1)) return false;
    return true;
  }
  if (a instanceof Map || b instanceof Map || (a as { type?: unknown }).type === 'element') return false;
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = Object.keys(left);
  if (keys.length !== Object.keys(right).length) return false;
  for (const key of keys) if (!sameProps(left[key], right[key], depth + 1)) return false;
  return true;
}

/** A run of nothing but text: such runs are joined when they are formatted alike. */
function isPlain(run: Run): boolean {
  if (run.fieldCode !== undefined || run.content.length === 0) return false;
  for (const item of run.content) if (!PLAIN_ITEMS.has(item.type)) return false;
  return true;
}

/** Items with neighbouring texts joined; `codes` (a flag per item) is kept in step. */
function joinTexts(items: RunItem[], codes: boolean[] | undefined): { items: RunItem[]; codes: boolean[] | undefined } {
  let out: RunItem[] | undefined;
  let flags: boolean[] | undefined;
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    const last = out ? out[out.length - 1] : i > 0 ? items[i - 1] : undefined;
    const join = item.type === 'text' && last !== undefined && last.type === 'text' && (!codes || (!codes[i] && !codes[i - 1]));
    if (join && !out) {
      out = items.slice(0, i);
      if (codes) flags = codes.slice(0, i);
    }
    if (!out) continue;
    if (join) out[out.length - 1] = { type: 'text', text: (out[out.length - 1] as { text: string }).text + (item as { text: string }).text };
    else {
      out.push(item);
      if (flags && codes) flags.push(codes[i]);
    }
  }
  return out ? { items: out, codes: flags } : { items, codes };
}

function makeRun(pieces: Piece[], from: number, to: number): Run {
  const first = pieces[from];
  const origin = (first.group as Run).type === 'run' ? (first.group as Run) : undefined;
  if (origin && to - from === origin.content.length && origin.props === first.props && origin.wrap === first.wrap) {
    let intact = true;
    for (let k = from; k < to && intact; k++) if (pieces[k].index !== k - from || pieces[k].item !== origin.content[k - from]) intact = false;
    if (intact) return origin;
  }
  let items: RunItem[] = [];
  let codes: boolean[] | undefined = [];
  let anyCode = false;
  let allCode = true;
  for (let k = from; k < to; k++) {
    items.push(pieces[k].item as RunItem);
    codes.push(pieces[k].code);
    if (pieces[k].code) anyCode = true;
    else allCode = false;
  }
  if (!anyCode) codes = undefined;
  ({ items, codes } = joinTexts(items, codes));
  const run: Run = { type: 'run', props: first.props, content: items };
  if (first.wrap) run.wrap = first.wrap;
  if (anyCode) run.fieldCode = allCode ? true : (codes as boolean[]);
  let attrs: AttrList | undefined;
  if (origin && origin.attrs) attrs = origin.attrs;
  if (attrs) run.attrs = attrs;
  return run;
}

/** The inlines of a paragraph from its pieces. */
export function rebuild(pieces: Piece[]): Inline[] {
  const out: Inline[] = [];
  let k = 0;
  while (k < pieces.length) {
    const piece = pieces[k];
    let next: Inline;
    if (piece.inline) {
      next = piece.inline;
      if (next.wrap !== piece.wrap) {
        next = { ...next, wrap: piece.wrap } as Inline;
        if (piece.wrap === undefined) delete next.wrap;
      }
      k++;
    } else {
      let end = k + 1;
      while (end < pieces.length && pieces[end].item && pieces[end].group === piece.group && pieces[end].props === piece.props && pieces[end].wrap === piece.wrap) end++;
      next = makeRun(pieces, k, end);
      k = end;
    }
    const last = out.length ? out[out.length - 1] : undefined;
    if (last && last.type === 'run' && next.type === 'run' && last.wrap === next.wrap && isPlain(last) && isPlain(next) && sameProps(last.props, next.props)) {
      const merged: Run = { type: 'run', props: last.props, content: joinTexts(last.content.concat(next.content), undefined).items };
      if (last.wrap) merged.wrap = last.wrap;
      out[out.length - 1] = merged;
    } else out.push(next);
  }
  return out;
}

/** A paragraph like `paragraph` with other content (and properties): a new object without a source span. */
export function paragraphWith(paragraph: Paragraph, children: Inline[], props = paragraph.props, keepAttrs = true): Paragraph {
  const out: Paragraph = { type: 'paragraph', props, children };
  if (paragraph.wrap) out.wrap = paragraph.wrap;
  if (keepAttrs && paragraph.attrs) out.attrs = paragraph.attrs;
  return out;
}

/** Have two lists of inlines the same members. */
export function sameInlines(a: Inline[], b: Inline[]): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

// ------------------------------------------------------- what a deletion leaves

const RANGE_MARK_NAMES = new Set<string>([
  'bookmarkStart',
  'bookmarkEnd',
  'commentRangeStart',
  'commentRangeEnd',
  'moveFromRangeStart',
  'moveFromRangeEnd',
  'moveToRangeStart',
  'moveToRangeEnd',
  'permStart',
  'permEnd',
]);

const localName = (name: string): string => name.slice(name.indexOf(':') + 1);

const loosePiece = (inline: Inline): Piece => ({ inline, props: EMPTY_RUN_PROPS, group: inline, index: -1, code: false, len: 0 });

/**
 * A range mark that was kept as a child of a table or a row (a bookmark
 * between two rows), as the mark it is. Other kept children of a table belong
 * to its structure and go with it.
 */
export function looseMark(extra: ExtraXml): Inline | undefined {
  const el = extra.el;
  const type = localName(el.name);
  if (extra.partial || extra.index < 0 || !RANGE_MARK_NAMES.has(type)) return undefined;
  let id: string | undefined;
  let name: string | undefined;
  let attrs: AttrList | undefined;
  el.attrs.forEach((value, key) => {
    const local = localName(key);
    if (local === 'id' && id === undefined) id = value;
    else if (local === 'name' && name === undefined) name = value;
    else (attrs || (attrs = [])).push([key, value]);
  });
  if (id === undefined || el.children.length) return { type: 'unknown', xml: el };
  const mark: RangeMark = { type: type as RangeMark['type'], id };
  if (name !== undefined) mark.name = name;
  if (attrs) mark.attrs = attrs;
  return mark;
}

/** What blocks leave behind when they are removed: the pieces of their paragraphs, and the blocks that are not content. */
export interface Harvest {
  pieces: Piece[];
  blocks: Block[];
}

function harvestLoose(extras: ExtraXml[] | undefined, misc: MiscXml[] | undefined, into: Harvest): void {
  if (extras) {
    for (const extra of extras) {
      const mark = looseMark(extra);
      if (mark) into.pieces.push(loosePiece(mark));
    }
  }
  if (misc) for (const note of misc) into.pieces.push(loosePiece({ type: 'misc', node: note.node }));
}

function refuseAlternate(wrap: Wrapper | undefined): void {
  if (wrap && chainHas(wrap, isAlternate)) throw new RangeError(READ_ONLY);
}

/** Adds what a row of a table leaves behind. */
export function harvestRow(row: TableRow, into: Harvest, depth = 0): void {
  refuseAlternate(row.wrap);
  harvestLoose(row.extras, row.misc, into);
  for (const cell of row.cells) {
    refuseAlternate(cell.wrap);
    harvestBlocks(cell.blocks, into, depth + 1);
  }
}

/** Adds what a table leaves behind. */
export function harvestTable(table: Table, into: Harvest, depth = 0): void {
  refuseAlternate(table.wrap);
  harvestLoose(table.extras, table.misc, into);
  for (const row of table.rows) harvestRow(row, into, depth);
}

/**
 * Adds what blocks leave behind when they are removed. Throws RangeError for
 * content that is written back whole (mc:AlternateContent) and so cannot be
 * taken apart.
 */
export function harvestBlocks(blocks: Block[], into: Harvest, depth = 0): void {
  // Tables nest as deep as the reader lets them, which is far less than this.
  if (depth > 256) throw new RangeError('The tables are nested too deep');
  for (const block of blocks) {
    refuseAlternate(block.wrap);
    if (block.type === 'paragraph') {
      const pieces = flatten(block);
      for (let i = 0; i < pieces.length; i++) into.pieces.push(pieces[i]);
    } else if (block.type === 'table') harvestTable(block, into, depth + 1);
    else into.blocks.push(block);
  }
}

/** The inlines that stay of harvested pieces once all their content is removed: marks of ranges, comment marks, kept XML, tracked deletions. */
export function leftInlines(pieces: Piece[]): Inline[] {
  return rebuild(removePieces(pieces, 0, pieces.length));
}
