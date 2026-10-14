// How a change of a document that waits for its place in the common order is read through a
// change that got its place first. Nothing here knows a document: a place is a story, the path
// of a paragraph and an offset in it, and what a change did to the places around it is said in
// a few words - text came in at a place (with so many paragraph ends in it), the text between
// two places went, some blocks were replaced by others.
//
// A list of blocks (the body, a header, the content of a table cell) is read as lines of text:
// a paragraph is a line, the end of a paragraph a character like any other. Typing, deleting,
// Enter and joining two paragraphs are then insertions and deletions of one long text, and the
// rules are the ones of that text:
//
//   - Two insertions at the same place both stay; the one that got its place LATER stands first.
//   - A deletion deletes what its author SAW: text someone else put into the range stays, and
//     the deletion is read as the two ranges around it.
//   - Text put into a range someone else deleted stays, where the range was.
//   - A range of character formatting covers what its author saw, like a deletion.
//
// Whatever is not known to come out right this way gives way: a change that is only known as
// "these blocks were replaced" (a table that got a row) or "something in this story", and a
// waiting change that is about the same blocks or story, do not go together, and the waiting
// one is dropped, which its author is told. See docs/collab-sozlesmesi.md.

import type { ParaFormatPatch, RunFormatPatch, ToggleKey } from '../../ops/doc/format';

/** A place in a document: the story, the path of the paragraph, the offset in its text; `b`: before what shows without width there. */
export interface WirePos {
  s: string;
  p: number[];
  o: number;
  b?: true;
}

/** A change of a document as it travels: the command that was run and where; every participant runs it on its own document. */
export type DocWire =
  /** Text typed or pasted at a place, or over a range (`to`). */
  | { t: 'text'; at: WirePos; to?: WirePos; x: string; f?: RunFormatPatch }
  | { t: 'del'; at: WirePos; to: WirePos }
  /** Enter. */
  | { t: 'split'; at: WirePos }
  | { t: 'run'; at: WirePos; to: WirePos; f: RunFormatPatch }
  /** A property that all of the range has is turned off. */
  | { t: 'unrun'; at: WirePos; to: WirePos; k: ToggleKey }
  | { t: 'para'; at: WirePos; to: WirePos; f: ParaFormatPatch }
  /** A command that is named (see COMMANDS in wire.ts), with what it was called with after the place. */
  | { t: 'cmd'; name: string; at: WirePos; to?: WirePos; a: unknown[] }
  /**
   * The `n` blocks from the block `at` names (its offset is 0) are replaced by the paragraphs `x`: how a step that
   * cannot be said as a command is taken back (text that was deleted, with how it looked). Everybody, its author
   * too, builds the paragraphs from what travelled.
   */
  | { t: 'put'; at: WirePos; n: number; x: PutParagraph[]; /** putDigest() of the blocks it replaces. */ h: string };

/** A paragraph as it travels in a `put`: its properties and its runs, each a list of texts, tabs and line breaks. See wire.ts for what is taken. */
export interface PutParagraph {
  /** Its formatting, as a change of paragraph formatting (ParaFormatPatch). */
  p?: Record<string, unknown>;
  /** The formatting of its mark, as the formatting of a run. */
  m?: Record<string, unknown>;
  a?: [string, string][];
  r: PutInline[];
}

export type PutInline = { p?: Record<string, unknown>; a?: [string, string][]; c: (string | [string, string] | number)[] } | { e: string } | { b: [0 | 1, string] | [0 | 1, string, string] };

/** What a change did to the places around it. */
export type Mark =
  /** Text came in at (row, off) of the list of blocks `c` of story `s`: `rows` paragraph ends, and `last` characters after the last of them (all of the text when there is none). */
  | { k: 'ins'; s: string; c: number[]; row: number; off: number; rows: number; last: number }
  /** The text from (r1, o1) up to (r2, o2) went. */
  | { k: 'del'; s: string; c: number[]; r1: number; o1: number; r2: number; o2: number }
  /** Formatting changed in the blocks r1..r2 of a list; nothing moved. */
  | { k: 'tch'; s: string; c: number[]; r1: number; r2: number }
  /** The `n` blocks from `at` of a list were replaced by `m` blocks, and what happened in them is not known place by place. */
  | {
      k: 'blk';
      s: string;
      c: number[];
      at: number;
      n: number;
      m: number;
      /** Where known: the plain text of the blocks before (`was`) and after (`now`), "" for one that is no paragraph. Carets in them follow the text (see the view). */
      was?: string[];
      now?: string[];
    }
  /** Something changed in the story that is not known place by place. "*": in the document. */
  | { k: 'opq'; s: string };

type TextMark = Extract<Mark, { k: 'ins' | 'del' }>;
type Moving = Extract<Mark, { k: 'ins' | 'del' | 'blk' }>;

const samePath = (a: readonly number[], b: readonly number[], length: number): boolean => {
  for (let i = 0; i < length; i++) if (a[i] !== b[i]) return false;
  return true;
};

const copy = (pos: WirePos, p: number[], o: number): WirePos => (pos.b ? { s: pos.s, p, o, b: true } : { s: pos.s, p, o });

/** Where the block at `row` of the list a mark is about stands after the mark; undefined when it is one of those that went. */
function rowThrough(row: number, mark: Moving): number | undefined {
  if (mark.k === 'ins') return row > mark.row ? row + mark.rows : row;
  if (mark.k === 'del') return row > mark.r2 ? row - (mark.r2 - mark.r1) : row > mark.r1 ? undefined : row;
  return row < mark.at ? row : row >= mark.at + mark.n ? row + mark.m - mark.n : undefined;
}

/** A path that leads through the list of blocks a mark is about (it is longer than the path of that list), after the mark. */
function pathThrough(path: readonly number[], mark: Moving): number[] | undefined {
  const depth = mark.c.length;
  const next = rowThrough(path[depth], mark);
  if (next === undefined) return undefined;
  const out = path.slice();
  out[depth] = next;
  return out;
}

const through = (s: string, path: readonly number[], mark: Moving): boolean => s === mark.s && path.length > mark.c.length && samePath(path, mark.c, mark.c.length);

/**
 * A place read through a mark. `right`: text that came in exactly at the place stands before
 * it (the place moves); otherwise the place stays where it is, before that text. `gone`: the
 * place was inside what went, and is where that was.
 */
export function mapPos(pos: WirePos, mark: Mark, right: boolean): { pos: WirePos; gone: boolean } {
  if (mark.k === 'tch' || mark.k === 'opq') return { pos, gone: false };
  if (!through(pos.s, pos.p, mark)) return { pos, gone: false };
  const depth = mark.c.length;
  if (pos.p.length > depth + 1 || mark.k === 'blk') {
    // In a block that stands in the list (a table), or beside blocks that were replaced: only the block moves.
    const path = pathThrough(pos.p, mark);
    return path ? { pos: path[depth] === pos.p[depth] ? pos : copy(pos, path, pos.o), gone: false } : { pos, gone: true };
  }
  const row = pos.p[depth];
  const off = pos.o;
  const at = (r: number, o: number, gone = false): { pos: WirePos; gone: boolean } => {
    if (r === row && o === off) return { pos, gone };
    const path = pos.p.slice();
    path[depth] = r;
    return { pos: copy(pos, path, o), gone };
  };
  if (mark.k === 'ins') {
    if (row < mark.row || (row === mark.row && (off < mark.off || (off === mark.off && !right)))) return { pos, gone: false };
    if (row > mark.row) return at(row + mark.rows, off);
    return mark.rows === 0 ? at(row, off + mark.last) : at(row + mark.rows, mark.last + off - mark.off);
  }
  if (row < mark.r1 || (row === mark.r1 && off <= mark.o1)) return { pos, gone: false };
  if (row > mark.r2) return at(row - (mark.r2 - mark.r1), off);
  if (row === mark.r2 && off >= mark.o2) return at(mark.r1, mark.o1 + off - mark.o2);
  return at(mark.r1, mark.o1, true);
}

/** -1, 0, 1 for two places of one list of blocks; undefined for places that are not in the same list. */
export function orderOf(a: WirePos, b: WirePos): number | undefined {
  if (a.s !== b.s || a.p.length !== b.p.length || !samePath(a.p, b.p, a.p.length - 1)) return undefined;
  const ra = a.p[a.p.length - 1];
  const rb = b.p[b.p.length - 1];
  return ra !== rb ? (ra < rb ? -1 : 1) : a.o !== b.o ? (a.o < b.o ? -1 : 1) : 0;
}

/** The place of a mark that put text in, and the place its text ends at. */
const insAt = (mark: Extract<Mark, { k: 'ins' }>): WirePos => ({ s: mark.s, p: mark.c.concat([mark.row]), o: mark.off });
const insEnd = (mark: Extract<Mark, { k: 'ins' }>): WirePos =>
  mark.rows === 0 ? { s: mark.s, p: mark.c.concat([mark.row]), o: mark.off + mark.last } : { s: mark.s, p: mark.c.concat([mark.row + mark.rows]), o: mark.last };

interface MappedRange {
  /** What is left of the range, in order; a range that holds nothing any more is among them with both ends at one place. */
  ranges: [WirePos, WirePos][];
  /** An end of the range was inside what went. */
  gone: boolean;
  /** Paragraphs were joined where an end of the range was: the paragraph it stood in is not the one it was. */
  joined: boolean;
}

/**
 * A range read through a mark. Text that came in at its start stands before it, text that
 * came in at its end after it, and text that came in inside it splits it (`split`; without it
 * the range grows around that text).
 */
export function mapRange(a: WirePos, b: WirePos, mark: Mark, split: boolean): MappedRange {
  if (mark.k === 'tch' || mark.k === 'opq') return { ranges: [[a, b]], gone: false, joined: false };
  const sorted = orderOf(a, b);
  if (sorted !== -1) {
    // A caret, or two ends that are not in one list of blocks: each is read by itself.
    const first = mapPos(a, mark, false);
    const last = sorted === 0 ? first : mapPos(b, mark, false);
    const gone = first.gone || last.gone;
    return { ranges: [[first.pos, last.pos]], gone, joined: gone && (mark.k !== 'del' || mark.r1 !== mark.r2) };
  }
  const first = mapPos(a, mark, true);
  const last = mapPos(b, mark, false);
  const gone = first.gone || last.gone;
  const joined = first.gone && last.gone && (mark.k !== 'del' || mark.r1 !== mark.r2);
  if (split && mark.k === 'ins') {
    const at = insAt(mark);
    if (orderOf(a, at) === -1 && orderOf(at, b) === -1) return { ranges: [[a, at], [insEnd(mark), last.pos]], gone: false, joined: false };
  }
  return { ranges: [[first.pos, last.pos]], gone, joined };
}

/** The list of blocks a mark is about and the blocks of it. */
function rowsOf(mark: Exclude<Mark, { k: 'opq' }>): { c: number[]; lo: number; hi: number } {
  if (mark.k === 'ins') return { c: mark.c, lo: mark.row, hi: mark.row };
  if (mark.k === 'blk') return { c: mark.c, lo: mark.at, hi: mark.at + mark.n - 1 };
  return { c: mark.c, lo: mark.r1, hi: mark.r2 };
}

/** Whether a mark is about something in (or over) the blocks another one replaced. */
function touches(mark: Mark, blk: Extract<Mark, { k: 'blk' }>): boolean {
  if (mark.k === 'opq') return mark.s === '*' || mark.s === blk.s;
  if (mark.s !== blk.s) return false;
  const rows = rowsOf(mark);
  const depth = blk.c.length;
  if (rows.c.length >= depth) {
    if (!samePath(rows.c, blk.c, depth)) return false;
    // In the same list: the same blocks. Deeper: inside one of them.
    if (rows.c.length === depth) return rows.lo < blk.at + Math.max(blk.n, 1) && rows.hi >= blk.at && (blk.n > 0 || mark.k !== 'ins');
    return rows.c[depth] >= blk.at && rows.c[depth] < blk.at + blk.n;
  }
  // The blocks stand inside a block of the list the mark is about: text there holds paragraphs only; formatting and replaced blocks may be over it.
  if (!samePath(blk.c, rows.c, rows.c.length)) return false;
  const row = blk.c[rows.c.length];
  return (mark.k === 'tch' || mark.k === 'blk') && row >= rows.lo && row <= rows.hi;
}

/** The places a change names. */
function placesOf(wire: DocWire): WirePos[] {
  if (wire.t === 'split' || wire.t === 'put') return [wire.at];
  return wire.to ? [wire.at, wire.to] : [wire.at];
}

/**
 * Whether a waiting change and a mark of a change that got its place do not go together: one
 * of them is only known as "something in this story" (or in these blocks), and the other is
 * about the same story (or blocks). `effect`: what the waiting change did when it was applied.
 */
function clash(places: readonly WirePos[], effect: readonly Mark[], mark: Mark): boolean {
  if (mark.k === 'opq') {
    if (mark.s === '*' || places.some((pos) => pos.s === mark.s)) return true;
  } else if (mark.k === 'blk') {
    for (const pos of places) if (through(pos.s, pos.p, mark) && rowThrough(pos.p[mark.c.length], mark) === undefined) return true;
  }
  for (const own of effect) {
    if (own.k === 'opq' && (own.s === '*' || mark.s === '*' || own.s === mark.s)) return true;
    if (own.k === 'blk' && touches(mark, own)) return true;
  }
  return false;
}

/** A mark read through a mark that the other side made: `pending` says this mark is of the waiting change (it stays first at a place both put text in). */
function markThrough(mark: Mark, over: Mark, pending: boolean): Mark[] {
  if (over.k === 'tch' || over.k === 'opq' || mark.k === 'opq' || mark.s !== over.s) return [mark];
  const lost: Mark[] = [{ k: 'opq', s: mark.s }];
  // The list of blocks the mark is about may itself stand in a block of the list the other is about.
  let c = mark.c;
  if (c.length > over.c.length && samePath(c, over.c, over.c.length)) {
    const moved = pathThrough(c, over);
    if (!moved) return lost;
    c = moved;
  } else if (c.length !== over.c.length || !samePath(c, over.c, c.length)) return [mark];
  else if (over.k === 'blk' && touches(mark, over)) return lost;
  const same = c.length === over.c.length;
  if (mark.k === 'blk') {
    const at = same ? rowThrough(mark.at, over) : mark.at;
    // Blocks that were replaced where the other took blocks away, or put text in between: not known any more.
    if (at === undefined) return lost;
    if (same && mark.n > 0 && rowThrough(mark.at + mark.n - 1, over) !== at + mark.n - 1) return lost;
    const out: Mark = { k: 'blk', s: mark.s, c, at, n: mark.n, m: mark.m };
    if (mark.was && mark.now) {
      out.was = mark.was;
      out.now = mark.now;
    }
    return [out];
  }
  if (mark.k === 'tch') {
    if (!same) return [{ k: 'tch', s: mark.s, c, r1: mark.r1, r2: mark.r2 }];
    const first = mapPos({ s: mark.s, p: c.concat([mark.r1]), o: 0 }, over, false).pos;
    const last = mapPos({ s: mark.s, p: c.concat([mark.r2]), o: Number.MAX_SAFE_INTEGER }, over, true).pos;
    return [{ k: 'tch', s: mark.s, c, r1: first.p[c.length], r2: Math.max(first.p[c.length], last.p[c.length]) }];
  }
  if (mark.k === 'ins') {
    if (!same) return [{ ...mark, c }];
    const at = mapPos({ s: mark.s, p: c.concat([mark.row]), o: mark.off }, over, !pending);
    if (at.gone && over.k === 'blk') return lost;
    return [{ k: 'ins', s: mark.s, c, row: at.pos.p[c.length], off: at.pos.o, rows: mark.rows, last: mark.last }];
  }
  if (!same) return [{ ...mark, c }];
  const out: Mark[] = [];
  const mapped = mapRange({ s: mark.s, p: c.concat([mark.r1]), o: mark.o1 }, { s: mark.s, p: c.concat([mark.r2]), o: mark.o2 }, over, true);
  if (mapped.gone && over.k === 'blk') return lost;
  // In the order they are applied in: the later range first, so that the earlier one is where it was.
  for (let i = mapped.ranges.length - 1; i >= 0; i--) {
    const [from, to] = mapped.ranges[i];
    if (orderOf(from, to) === 0) continue;
    out.push({ k: 'del', s: mark.s, c, r1: from.p[c.length], o1: from.o, r2: to.p[c.length], o2: to.o });
  }
  return out;
}

/** Two runs of marks, each read through the other: `a` is of the waiting change. */
export function marksThrough(a: readonly Mark[], b: readonly Mark[]): [Mark[], Mark[]] {
  if (a.length === 0 || b.length === 0) return [a.slice(), b.slice()];
  if (a.length === 1 && b.length === 1) return [markThrough(a[0], b[0], true), markThrough(b[0], a[0], false)];
  if (a.length === 1) {
    const [a1, b1] = marksThrough(a, [b[0]]);
    const [a2, rest] = marksThrough(a1, b.slice(1));
    return [a2, b1.concat(rest)];
  }
  const [a1, b1] = marksThrough([a[0]], b);
  const [rest, b2] = marksThrough(a.slice(1), b1);
  return [a1.concat(rest), b2];
}

/**
 * Ranges that went one after the other and lie side by side are one range that went: a
 * deletion that was read as two ranges around text which is gone as well by now.
 */
function joinedMarks(marks: readonly Mark[]): Mark[] {
  const out: Mark[] = [];
  for (const mark of marks) {
    const last = out[out.length - 1];
    if (mark.k === 'del' && last && last.k === 'del' && last.s === mark.s && last.c.length === mark.c.length && samePath(last.c, mark.c, mark.c.length)) {
      if (mark.r2 === last.r1 && mark.o2 === last.o1) {
        // The one that went second ends where the first began.
        out[out.length - 1] = { k: 'del', s: mark.s, c: mark.c, r1: mark.r1, o1: mark.o1, r2: last.r2, o2: last.o2 };
        continue;
      }
      if (mark.r1 === last.r1 && mark.o1 === last.o1) {
        // The one that went second begins where the first did: it ends as far behind the end of the first.
        const r2 = mark.r2 + (last.r2 - last.r1);
        const o2 = mark.r2 === last.r1 ? last.o2 + mark.o2 - last.o1 : mark.o2;
        out[out.length - 1] = { k: 'del', s: mark.s, c: mark.c, r1: last.r1, o1: last.o1, r2, o2 };
        continue;
      }
    }
    out.push(mark);
  }
  return out;
}

export interface Rebased {
  /** The change as it is to be applied now: none when it is dropped or nothing is left of it. */
  wires: DocWire[];
  /** The change does not go together with what got its place, and is dropped. */
  drop?: true;
  /** Nothing is left to do: what it was about went. */
  noop?: true;
  /** Part (or all) of what it was to delete or to format went with a deletion that got its place first. */
  cut?: true;
  /** What got its place, read through the change as it was: what the changes that wait behind it are read through. */
  carry: Mark[];
}

/**
 * One waiting change read through what got its place first (`carry`, in order). `effect`: what
 * the change did when it was applied last, in the places of that time.
 *
 * Whether the two go together is asked mark by mark, with the places of the change as they
 * are after the marks before it: a mark is about the document the marks before it left.
 */
export function rebaseWire(wire: DocWire, effect: readonly Mark[], arrivedMarks: readonly Mark[]): Rebased {
  const carry = joinedMarks(arrivedMarks);
  const dropped = (): Rebased => ({ wires: [], drop: true, carry: carry.slice() });
  const moved = marksThrough(effect, carry)[1];
  /** What the change did, read through the marks that were gone through. */
  let own: Mark[] = effect.slice();
  const on = (mark: Mark): void => {
    own = marksThrough(own, [mark])[0];
  };
  if (wire.t === 'put') {
    let at = wire.at;
    for (const mark of carry) {
      const depth = at.p.length - 1;
      // Whatever is about the blocks it replaces, or inside them: the two do not go together.
      if (touches(mark, { k: 'blk', s: at.s, c: at.p.slice(0, depth), at: at.p[depth], n: wire.n, m: wire.n }) || clash([at], own, mark)) return dropped();
      if (mark.k === 'ins' || mark.k === 'del' || mark.k === 'blk') {
        if (through(at.s, at.p, mark)) {
          const path = pathThrough(at.p, mark);
          if (!path) return dropped();
          at = { s: at.s, p: path, o: 0 };
        }
      }
      on(mark);
    }
    return { wires: [{ ...wire, at }], carry: moved };
  }
  if (wire.t === 'split' || ((wire.t === 'text' || wire.t === 'cmd') && !wire.to)) {
    let at = wire.at;
    let gone = false;
    for (const mark of carry) {
      if (clash([at], own, mark)) return dropped();
      const next = mapPos(at, mark, false);
      at = next.pos;
      if (next.gone) gone = true;
      on(mark);
    }
    // Text and the end of a paragraph stay where the text around them went; a command is about a place that is no more.
    if (gone && wire.t === 'cmd') return dropped();
    return { wires: [{ ...wire, at }], carry: moved };
  }
  let ranges: [WirePos, WirePos][] = [[wire.at, wire.to as WirePos]];
  /** Where the text of a replacement goes: where the range began, before whatever came in there. */
  let start = wire.at;
  const sorted = orderOf(wire.at, wire.to as WirePos);
  // Character formatting and deletions cover what their author saw; what is about paragraphs covers the paragraphs,
  // and so does what is about a caret (formatting at a place where nothing is selected).
  const split = (wire.t === 'del' || wire.t === 'run' || wire.t === 'unrun' || wire.t === 'text') && (sorted === -1 || sorted === 1);
  if (sorted === 1) {
    ranges = [[wire.to as WirePos, wire.at]];
    start = wire.to as WirePos;
  }
  let joined = false;
  let cut = false;
  // For a caret: what arrived took away what stood right before it, and what stood right behind it.
  let before = false;
  let behind = false;
  let rows = false;
  /** The caret as it stands behind text that came in exactly where it is: what went right behind it went behind that text. */
  let far = wire.at;
  for (const mark of carry) {
    const places: WirePos[] = [start];
    for (const [a, b] of ranges) places.push(a, b);
    if (clash(places, own, mark)) return dropped();
    // Blocks that were replaced between the two ends of a range: what the range was about is no longer there to be told apart.
    if (mark.k === 'blk') {
      for (const [a, b] of ranges) {
        const depth = a.p.length - 1;
        if (orderOf(a, b) !== undefined && touches(mark, { k: 'blk', s: a.s, c: a.p.slice(0, depth), at: a.p[depth], n: b.p[depth] - a.p[depth] + 1, m: 0 })) return dropped();
      }
    }
    if (sorted === 0 && mark.k === 'del' && ranges.length === 1) {
      const at = ranges[0][0];
      const depth = mark.c.length;
      if (at.s === mark.s && at.p.length === depth + 1 && samePath(at.p, mark.c, depth)) {
        const first: WirePos = { s: mark.s, p: mark.c.concat([mark.r1]), o: mark.o1 };
        const last: WirePos = { s: mark.s, p: mark.c.concat([mark.r2]), o: mark.o2 };
        // Something went right before the caret (it is inside the range, or at its end) or right behind it.
        const ends = orderOf(first, at) === -1 && orderOf(at, last) !== -1 && orderOf(at, last) !== undefined && orderOf(at, last) !== 1;
        const begins = orderOf(first, far) !== 1 && orderOf(first, far) !== undefined && orderOf(far, last) === -1;
        if (ends) before = true;
        if (begins) behind = true;
        if ((ends || begins) && mark.r1 !== mark.r2) rows = true;
      }
    }
    if (sorted === 0) far = mapPos(far, mark, true).pos;
    start = mapPos(start, mark, false).pos;
    const next: [WirePos, WirePos][] = [];
    for (const [a, b] of ranges) {
      if (split && mark.k === 'del') {
        // The two ranges share text: what the other one took away is not this one's to take away, or to format, any more.
        const first: WirePos = { s: mark.s, p: mark.c.concat([mark.r1]), o: mark.o1 };
        const last: WirePos = { s: mark.s, p: mark.c.concat([mark.r2]), o: mark.o2 };
        if (orderOf(a, last) === -1 && orderOf(first, b) === -1) cut = true;
      }
      const mapped = mapRange(a, b, mark, split);
      if (mapped.joined) joined = true;
      for (const range of mapped.ranges) if (!split || orderOf(range[0], range[1]) !== 0) next.push(range);
    }
    ranges = next;
    on(mark);
  }
  if (wire.t === 'text') {
    const wires: DocWire[] = [];
    // The text keeps its place before what came in where the range began; when the range still begins there it is one replacement as before.
    const whole = ranges.length === 1 && orderOf(start, ranges[0][0]) === 0;
    for (let i = ranges.length - 1; i >= (whole ? 1 : 0); i--) wires.push({ t: 'del', at: ranges[i][0], to: ranges[i][1] });
    const text: DocWire = whole ? { t: 'text', at: ranges[0][0], to: ranges[0][1], x: wire.x } : { t: 'text', at: start, x: wire.x };
    if (wire.f !== undefined) text.f = wire.f;
    wires.push(text);
    return cut ? { wires, cut: true, carry: moved } : { wires, carry: moved };
  }
  if (!split) {
    // Every paragraph it was about went into another one: what stands there now is not what was meant.
    if (joined || (before && behind && rows) || ranges.length !== 1) return dropped();
    return { wires: [{ ...wire, at: ranges[0][0], to: ranges[0][1] } as DocWire], carry: moved };
  }
  if (ranges.length === 0) return { wires: [], noop: true, cut: true, carry: moved };
  const wires: DocWire[] = [];
  // Later ranges first: what goes there does not move what stands before it.
  for (let i = ranges.length - 1; i >= 0; i--) wires.push({ ...wire, at: ranges[i][0], to: ranges[i][1] } as DocWire);
  return cut ? { wires, cut: true, carry: moved } : { wires, carry: moved };
}

/**
 * The marks of a step that got its place, as the changes that wait are read through them.
 *
 * A step that replaces text says "this went, then that came in where it began". For a change
 * that waits at a place inside what went (or right behind it), that order loses what stood
 * where: the place falls onto the start of the range, and so does the new text. What is meant
 * is that the new text stands where the range ENDED - behind whatever a waiting change put
 * into the range, before whatever stood behind it - so the text is said first, at the end of
 * the range, and the range goes after it. (The other way round, when the step itself waits
 * and what arrived is read through it, the order it was applied in is the right one: its text
 * stands first at the place.)
 */
export function arrived(marks: readonly Mark[]): Mark[] {
  const out: Mark[] = [];
  for (const mark of marks) {
    if (mark.k !== 'ins') {
      out.push(mark);
      continue;
    }
    let row = mark.row;
    let off = mark.off;
    let at = out.length;
    // Over the ranges that went right before it, to stand before them all.
    while (at > 0) {
      const before = out[at - 1];
      if (before.k !== 'del' || before.s !== mark.s || before.c.length !== mark.c.length || !samePath(before.c, mark.c, mark.c.length)) break;
      const after = before.r1 > row || (before.r1 === row && before.o1 > off);
      if (before.r1 === row && before.o1 === off) {
        // The range began where the text comes in: the text stands where it ended.
        row = before.r2;
        off = before.o2;
      } else if (after) {
        // The range lies behind the text, and moves with it.
        const from = mapPos({ s: before.s, p: before.c.concat([before.r1]), o: before.o1 }, { ...mark, row, off }, true).pos;
        const to = mapPos({ s: before.s, p: before.c.concat([before.r2]), o: before.o2 }, { ...mark, row, off }, true).pos;
        const depth = before.c.length;
        out[at - 1] = { k: 'del', s: before.s, c: before.c, r1: from.p[depth], o1: from.o, r2: to.p[depth], o2: to.o };
      } else if (row === before.r1) {
        // The range lies before the text: with it still there the text stands further on.
        off = before.o2 + off - before.o1;
        row = before.r2;
      } else row += before.r2 - before.r1;
      at--;
    }
    out.splice(at, 0, { ...mark, row, off });
  }
  return out;
}
