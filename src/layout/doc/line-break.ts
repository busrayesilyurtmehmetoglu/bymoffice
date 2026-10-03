// Breaks the atoms of a prepared paragraph into one line at a time.
//
// Greedy, like Word: a line takes what fits and ends at the last break
// opportunity before the first atom that does not. Spaces at the end of a
// line hang out of it. A word that is wider than the line is cut between
// characters. Text runs left to right; there is no bidi reordering and no
// shaping beyond what the measurer does for the strings it is given.

import type { FieldValues, LayoutContext, RunStyle } from './context';
import type { Atom, ParaPrep, PrepTab, TextAtom } from './prep';
import { formatPageNumber, type LiveField } from './symbols';
import type { LayoutFont, PageItem } from './types';
import { EPSILON } from './units';

/** A piece of a line: an atom, or the part of one that is on the line, at its place. */
export interface Piece {
  kind: 'text' | 'tab' | 'box';
  /** Left edge, relative to the left edge of the column. */
  x: number;
  /** Advance, trailing spaces included. */
  width: number;
  /** Width and number of the spaces at the end. */
  trail: number;
  spaces: number;
  text: string;
  font: LayoutFont;
  style: RunStyle;
  /** Offset in paragraphText(), -1 without a source. */
  offset: number;
  link?: string;
  note?: number;
  field?: LiveField;
  fieldFormat?: string;
  /** Tabs: the leader name. */
  leader?: string;
  /** Boxes. */
  height?: number;
  items?: PageItem[];
  /** The piece takes no room of its own (a number hanging left of the text). */
  label?: boolean;
}

export interface Line {
  pieces: Piece[];
  /** Where the next line starts. */
  atom: number;
  ch: number;
  /** The break the line ends with. */
  brk?: 'line' | 'page' | 'column';
  /** Font of the break, which gives an otherwise empty line its height. */
  brkFont?: LayoutFont;
  /** Nothing of the paragraph is left. */
  last: boolean;
  /** Right edge of the content, hanging spaces left out. */
  right: number;
  /** Index of the first piece after the last tab: what justification may stretch. */
  free: number;
}

/** Text of a field the engine computes: the value when it is known, the cached result otherwise. */
export function fieldText(field: LiveField, format: string | undefined, cached: string, values: FieldValues | undefined): string {
  let value: number | undefined;
  if (values) value = field === 'PAGE' ? values.page : field === 'NUMPAGES' ? values.numPages : values.sectionPages;
  if (value === undefined) return cached;
  return formatPageNumber(value, format !== undefined ? format : field === 'PAGE' && values ? values.pageFormat : undefined);
}

/** The tab stop a tab at x runs to. `fixed` stops are those the paragraph defines; the others repeat at the default distance. */
function nextStop(ctx: LayoutContext, prep: ParaPrep, x: number): PrepTab & { fixed: boolean } {
  let best: PrepTab | undefined;
  for (let i = 0; i < prep.tabs.length; i++) {
    if (prep.tabs[i].pos > x + EPSILON) {
      best = prep.tabs[i];
      break;
    }
  }
  // A hanging indent is a tab stop where the text starts.
  if (prep.first < 0 && prep.left > x + EPSILON && (!best || prep.left < best.pos)) best = { pos: prep.left, align: 'left' };
  if (best) return { pos: best.pos, align: best.align, leader: best.leader, fixed: true };
  const step = ctx.defaultTab;
  return { pos: (Math.floor((x + EPSILON) / step) + 1) * step, align: 'left', fixed: false };
}

function pieceOf(ctx: LayoutContext, atom: TextAtom, from: number, to: number): Piece {
  const whole = from === 0 && to === atom.text.length;
  const word = Math.max(0, Math.min(atom.word, to) - from);
  const spaces = to - from - word;
  let width: number;
  let trail: number;
  if (whole) {
    trail = atom.trail;
    width = atom.width + trail;
  } else {
    const style = atom.style;
    width = ctx.width(atom.font, atom.text.slice(from, from + word)) * style.scale + style.spacing * word;
    trail = spaces === 0 ? 0 : (atom.trail / (atom.text.length - atom.word)) * spaces;
    width = Math.max(0, width) + trail;
  }
  const piece: Piece = {
    kind: 'text',
    x: 0,
    width,
    trail,
    spaces,
    text: whole ? atom.text : atom.text.slice(from, to),
    font: atom.font,
    style: atom.style,
    offset: atom.offset < 0 ? -1 : atom.offset + from,
  };
  if (atom.link !== undefined) piece.link = atom.link;
  if (atom.note !== undefined && from === 0) piece.note = atom.note;
  return piece;
}

/** How many characters of the word of the atom, from `from`, fit into `available`; at least one when `must`. */
function fitCharacters(ctx: LayoutContext, atom: TextAtom, from: number, available: number, must: boolean): number {
  const style = atom.style;
  let used = 0;
  let count = from;
  while (count < atom.word) {
    const code = atom.text.charCodeAt(count);
    const step = code >= 0xd800 && code <= 0xdbff && count + 1 < atom.word ? 2 : 1;
    const width = ctx.width(atom.font, atom.text.substr(count, step)) * style.scale + style.spacing * step;
    if (used + width > available + EPSILON && (count > from || !must)) break;
    used += width;
    count += step;
  }
  return count - from;
}

/**
 * The rest of a word that was cut at the end of the line before is not
 * measured as a whole when it is longer than this: every line of a very long
 * word would measure all that is left of it, which is quadratic in its length.
 */
const LONG_REST = 256;

interface OpenTab {
  /** Index of the tab piece. */
  index: number;
  stop: number;
  align: string;
  /** x where the tab starts. */
  start: number;
}

/**
 * Width of the text of the pieces up to the decimal separator; undefined when
 * there is none. The separator is that of the language of the text ("," in
 * Turkish or German, "." in English): "1.250.000,5" aligns on its comma there.
 * Text of no known language aligns on the last "." or "," after a digit, the
 * one a number written either way ends its whole part with.
 */
function widthToDecimal(ctx: LayoutContext, pieces: Piece[], from: number): number | undefined {
  let width = 0;
  let guess: number | undefined;
  for (let i = from; i < pieces.length; i++) {
    const piece = pieces[i];
    if (piece.kind === 'text') {
      const text = piece.text;
      const separator = piece.style.decimal;
      let at = -1;
      if (separator !== undefined) {
        for (let k = text.indexOf(separator, 1); k > 0; k = text.indexOf(separator, k + 1)) {
          const before = text.charCodeAt(k - 1);
          if (before >= 0x30 && before <= 0x39) {
            at = k;
            break;
          }
        }
      } else {
        for (let k = text.length - 1; k > 0; k--) {
          const code = text.charCodeAt(k);
          const before = text.charCodeAt(k - 1);
          if ((code === 0x2c || code === 0x2e) && before >= 0x30 && before <= 0x39) {
            at = k;
            break;
          }
        }
      }
      if (at > 0) {
        const upTo = width + ctx.width(piece.font, text.slice(0, at)) * piece.style.scale + piece.style.spacing * at;
        if (separator !== undefined) return upTo;
        guess = upTo;
      }
    }
    width += piece.width;
  }
  return guess;
}

/**
 * Breaks one line. The line starts at character `ch` of atom `atomIndex`
 * and lies between x0 and x1 (relative to the left edge of the column).
 * `markStays`: a page or column break that ends the paragraph ends it on
 * this line (the paragraph mark does not go to the next page).
 * Without `force` a line that cannot take its first word whole stays empty
 * (the caller has another place for the word); with it the word is placed
 * anyway, cut if need be, so that every line takes something.
 */
export function breakLine(
  ctx: LayoutContext,
  prep: ParaPrep,
  atomIndex: number,
  ch: number,
  x0: number,
  x1: number,
  fields: FieldValues | undefined,
  markStays: boolean,
  force = true,
): Line {
  const atoms = prep.atoms;
  const pieces: Piece[] = [];
  let x = x0;
  let i = atomIndex;
  let open: OpenTab | undefined;
  let free = 0;
  // The last place the line may end at: number of pieces, the atom to go on with, and the tab open then.
  let breakCount = -1;
  let breakAtom = 0;
  let breakOpen: OpenTab | undefined;
  const line: Line = { pieces, atom: atoms.length, ch: 0, last: true, right: x0, free: 0 };

  /** Moves what follows an aligning tab so that it sits at the stop the tab runs to. */
  const settle = (atLineEnd: boolean) => {
    if (!open) return;
    const tab = open;
    open = undefined;
    if (tab.index >= pieces.length) return;
    let end = x;
    if (atLineEnd && pieces.length > tab.index + 1) end -= pieces[pieces.length - 1].trail;
    const width = end - tab.start;
    let shift: number;
    if (tab.align === 'center') shift = tab.stop - tab.start - width / 2;
    else {
      const upTo = tab.align === 'decimal' ? widthToDecimal(ctx, pieces, tab.index + 1) : undefined;
      shift = tab.stop - tab.start - (upTo === undefined ? width : upTo);
    }
    if (!(shift > 0)) return;
    pieces[tab.index].width = shift;
    for (let k = tab.index + 1; k < pieces.length; k++) pieces[k].x += shift;
    x += shift;
  };
  /** Where content ending at `end` would end once the open tab is settled. */
  const settledEnd = (end: number) => {
    if (!open) return end;
    if (open.align === 'center') return Math.max(end, open.stop + (end - open.start) / 2);
    return Math.max(end, open.stop);
  };
  const addTab = (stop: PrepTab & { fixed: boolean }, font: LayoutFont, style: RunStyle, offset: number, link: string | undefined) => {
    settle(false);
    const piece: Piece = { kind: 'tab', x, width: 0, trail: 0, spaces: 0, text: '', font, style, offset };
    if (link !== undefined) piece.link = link;
    if (stop.leader !== undefined) piece.leader = stop.leader;
    pieces.push(piece);
    if (stop.align === 'center' || stop.align === 'right' || stop.align === 'decimal') {
      open = { index: pieces.length - 1, stop: stop.pos, align: stop.align, start: x };
    } else {
      piece.width = Math.max(0, stop.pos - x);
      x += piece.width;
    }
    free = pieces.length;
  };

  if (atomIndex === 0 && ch === 0 && prep.label) {
    const label = prep.label;
    const hang = label.align === 'right' ? label.width : label.align === 'center' ? label.width / 2 : 0;
    pieces.push({ kind: 'text', x: x - hang, width: label.width, trail: 0, spaces: 0, text: label.text, font: label.font, style: label.style, offset: -1, label: true });
    x += label.width - hang;
    if (label.suffix === 'tab') {
      const stop = nextStop(ctx, prep, x);
      // A number that runs past the right edge is followed by a space, not by a tab to the next line.
      if (stop.pos <= x1 || stop.fixed) addTab({ pos: stop.pos, align: 'left', fixed: stop.fixed }, label.font, label.style, -1, undefined);
      else x += ctx.width(label.font, ' ');
    } else if (label.suffix === 'space') x += ctx.width(label.font, ' ');
    free = pieces.length;
  }
  /** Pieces that are content of the paragraph: the label does not count for "the line is still empty". */
  const base = pieces.length;

  let ended = false;
  while (i < atoms.length) {
    const atom: Atom = atoms[i];
    if (atom.t === 'break') {
      line.brk = atom.kind;
      line.brkFont = atom.font;
      i++;
      ch = 0;
      // After a line break the paragraph mark has a line of its own; after a page break only in older documents.
      line.last = i >= atoms.length && atom.kind !== 'line' && markStays;
      ended = true;
      break;
    }
    if (atom.t === 'tab') {
      // What an earlier aligning tab holds is complete now, so a break kept below never has to undo it.
      settle(false);
      let stop: PrepTab & { fixed: boolean };
      if (atom.absolute) {
        const align = atom.absolute.align;
        stop = { pos: align === 'right' ? x1 : align === 'center' ? (x0 + x1) / 2 : x0, align, leader: atom.absolute.leader, fixed: true };
        if (stop.pos < x) stop.pos = x;
      } else stop = nextStop(ctx, prep, x);
      // A tab that would run to a default stop beyond the right edge starts the next line instead.
      if (!stop.fixed && stop.pos > x1 + EPSILON) {
        if (pieces.length > base) {
          ended = true;
          break;
        }
        stop.pos = Math.max(x, x1);
      }
      if (pieces.length > base) {
        breakCount = pieces.length;
        breakAtom = i;
        breakOpen = open;
      }
      addTab(stop, atom.font, atom.style, atom.offset, atom.link);
      i++;
      ch = 0;
      continue;
    }

    if (ch === 0 && atom.bb && pieces.length > base) {
      breakCount = pieces.length;
      breakAtom = i;
      breakOpen = open;
    }
    let piece: Piece;
    if (atom.t === 'text') {
      if (ch > 0 && atom.word - ch > LONG_REST && pieces.length === base) {
        // What is left of a long word: only as many characters as the line takes are looked at.
        const count = fitCharacters(ctx, atom, ch, x1 - settledEnd(x), force);
        if (count < atom.word - ch) {
          if (count > 0) {
            const part = pieceOf(ctx, atom, ch, ch + count);
            part.x = x;
            pieces.push(part);
            x += part.width;
            ch += count;
          } else pieces.length = 0;
          ended = true;
          break;
        }
      }
      piece = ch === 0 ? pieceOf(ctx, atom, 0, atom.text.length) : pieceOf(ctx, atom, ch, atom.text.length);
    } else if (atom.t === 'box') {
      piece = { kind: 'box', x: 0, width: atom.width, trail: 0, spaces: 0, text: '', font: atom.font, style: atom.style, offset: atom.offset, height: atom.height, items: atom.items };
    } else {
      const text = fieldText(atom.field, atom.format, atom.cached, fields);
      const style = atom.style;
      piece = { kind: 'text', x: 0, width: Math.max(0, ctx.width(atom.font, text) * style.scale + style.spacing * text.length), trail: 0, spaces: 0, text, font: atom.font, style, offset: atom.offset, field: atom.field };
      if (atom.format !== undefined) piece.fieldFormat = atom.format;
      if (atom.link !== undefined) piece.link = atom.link;
    }

    if (settledEnd(x + piece.width - piece.trail) > x1 + EPSILON) {
      if (breakCount >= 0) {
        // Back to the last place a line may end.
        pieces.length = breakCount;
        i = breakAtom;
        ch = 0;
        open = breakOpen && breakOpen.index < breakCount ? breakOpen : undefined;
        if (free > breakCount) free = breakCount;
        x = breakCount > 0 ? pieces[breakCount - 1].x + pieces[breakCount - 1].width : x0;
        ended = true;
        break;
      }
      if (!force && pieces.length === base) {
        // Nothing of the paragraph is on the line and its next word does not fit: the line stays empty.
        pieces.length = 0;
        ended = true;
        break;
      }
      if (atom.t === 'text' && atom.word - ch > 1) {
        // No place to break: the word is cut where the line is full.
        const count = fitCharacters(ctx, atom, ch, x1 - settledEnd(x), pieces.length === base);
        if (count > 0) {
          const part = pieceOf(ctx, atom, ch, ch + count);
          part.x = x;
          pieces.push(part);
          x += part.width;
          ch += count;
        }
        ended = true;
        break;
      }
      if (pieces.length > base) {
        ended = true;
        break;
      }
      // The first thing on the line and it cannot be made smaller: it goes here whatever its width.
    }
    piece.x = x;
    pieces.push(piece);
    x += piece.width;
    i++;
    ch = 0;
  }

  settle(true);
  if (ended) {
    line.atom = i;
    line.ch = ch;
    if (!line.brk) line.last = false;
    if (i >= atoms.length && !line.brk) line.last = true;
  }
  if (!line.last && !line.brk && pieces.length > base && ch === 0 && i > 0) {
    // A line that ends at a soft hyphen shows it.
    const before = atoms[i - 1];
    const tail = pieces[pieces.length - 1];
    if (before.t === 'text' && before.shy && tail.kind === 'text' && tail.spaces === 0) {
      // The hyphen is not in the text of the paragraph: a piece of its own, without a place in the text.
      const width = ctx.width(tail.font, '-') * tail.style.scale;
      pieces.push({ kind: 'text', x: tail.x + tail.width, width, trail: 0, spaces: 0, text: '-', font: tail.font, style: tail.style, offset: -1 });
      x = tail.x + tail.width + width;
    }
  }
  line.free = Math.min(free, pieces.length);
  if (pieces.length) {
    const tail = pieces[pieces.length - 1];
    line.right = tail.x + tail.width - tail.trail;
  }
  return line;
}
