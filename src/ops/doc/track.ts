// Tracked changes: what the editing commands need to make their changes as
// revisions instead of making them outright. A document that asks for it
// (DocSettings.trackRevisions, or a protection that allows nothing but
// tracked changes) is edited this way, and so is any document when the
// caller says so.
//
// What a tracked change looks like in the model:
//   - text put in sits in an `ins` wrapper; the mark of a paragraph that was
//     made says `inserted` in its mark properties;
//   - text taken out stays, in a `del` wrapper innermost around its run, its
//     text as <w:delText> items; the mark of a paragraph that was joined to
//     the next says `deleted`, and the two paragraphs stay two;
//   - text (and tabs and line breaks) that sits directly in a tracked
//     insertion the same author made since the document was opened (see
//     Tracking.made) is taken out for good, as Word does:
//     nobody needs the history of a change to a change. Everything else in
//     such an insertion (a note reference, a field, a picture, an equation,
//     and text inside a simple field, a link, a content control or a smart
//     tag) is deleted like anything else, inside the insertion, so that
//     nothing but typed text ever leaves a document that tracks its changes,
//     and no element that held it goes with it. The mark of a paragraph the
//     author made goes for good only between two paragraphs formatted alike. An editor without a name owns no
//     insertion: the stand-in name is also what Word writes for anybody whose
//     name was removed from a file.
//   - formatting that is changed stays changed, and what it was is recorded:
//     a run says so in its properties (`change`, a <w:rPrChange> that holds
//     the run properties as they were), a paragraph in its own (a
//     <w:pPrChange> with the paragraph properties as they were: alignment,
//     indents, spacing, style, list - not the mark, not the section). The
//     records are elements, spelled as the part of their story spells its
//     elements, because that is what the model keeps of the ones a file
//     brings: listing, accepting and rejecting know one kind only. Rules, as
//     Word has them (see trackedRunProps):
//       - formatting that carries a record already, whoever made it, gets no
//         second one: what it was first is what a rejection brings back;
//       - the author who, in the same session, brings formatting back to
//         exactly what their own record says takes the record away;
//       - text in a tracked insertion the same author made in this session
//         is formatted without a record: what is new has no formatting it
//         used to have. A paragraph (and its mark) is new in that sense only
//         when all of it is: the mark and everything before it. A paragraph
//         the author split still holds text that was there;
//       - deleted text is not formatted at all.
// Content in a `del` wrapper is not content any more: it takes no offsets
// (see position.ts) and no command touches it again.
//
// Rows, cells and tables that are put in or taken out, and content that is
// put in whole (a paste, a picture, a linked text): see track-blocks.ts.

import { createContext, createReaderState } from '../../formats/docx/common';
import { paraPropsElement, runPropsElement, wEl, type PropsContext } from '../../formats/docx/write-props';
import { EMPTY_PARA_PROPS, EMPTY_RUN_PROPS, tracksChanges, type AttrList, type Block, type DocModel, type DocRelationship, type ExtraXml, type Inline, type ParaProps, type RevisionInfo, type RunItem, type RunProps, type Wrapper } from '../../model/doc/types';
import { ownerOf } from './stand-in';
import type { XmlElement } from '../../xml';
import { fieldPairs, makeRun, removedPieces, sameProps, type Piece } from './pieces';

/** What a command throws for a change it cannot make as a tracked one: the change is not made at all. */
export const TRACK_UNSUPPORTED = 'This change cannot be tracked yet';

/** The author new revisions name when nobody says who is editing. UNSURE: Word uses the name of its user; "Author" is what it shows for a document whose personal data was removed. */
export const DEFAULT_AUTHOR = 'Author';

export interface TrackOptions {
  /** Track the changes of this command even though the document does not ask for it. False does not turn tracking off for a document that asks for it. */
  trackChanges?: boolean;
  /** Who makes the change (w:author). */
  author?: string;
  /**
   * When, as ISO 8601 text ("2026-10-05T12:00:00Z"; UTC to the second is
   * what the editor writes), which is what w:date gets as it is. Anything
   * else is refused with a RangeError.
   */
  date?: string;
}

/** How the changes of one document are tracked when a command is not told otherwise. */
export interface TrackingSetup {
  /** Track although the document does not ask for it. */
  force?: boolean;
  author?: string;
  /** The clock revisions take their date from. */
  now?: () => Date;
}

const SETUPS = new WeakMap<DocModel, TrackingSetup>();

/** Sets how the changes of a document are tracked from now on; see TrackingSetup. */
export function setTracking(doc: DocModel, setup: TrackingSetup): void {
  SETUPS.set(ownerOf(doc), { ...setup });
}

/** Did the application ask for the changes of this document to be tracked whatever the document says (setTracking with `force`). */
export function forcesTracking(doc: DocModel): boolean {
  const setup = SETUPS.get(ownerOf(doc));
  return setup !== undefined && setup.force === true;
}

const two = (value: number): string => (value < 10 ? '0' : '') + value;

/** A moment as w:date holds it: UTC, to the second. */
export function revisionDate(at: Date): string {
  const time = at.getTime();
  // A clock that tells no time must not put "NaN" into a file.
  const date = time === time ? at : new Date(0);
  return `${date.getUTCFullYear()}-${two(date.getUTCMonth() + 1)}-${two(date.getUTCDate())}T${two(date.getUTCHours())}:${two(date.getUTCMinutes())}:${two(date.getUTCSeconds())}Z`;
}

/** The tracking one command works with. */
export interface Tracking {
  author: string;
  /** The author was named. Without a name nothing in the document counts as the editor's own. */
  named: boolean;
  date: string;
  /** A revision of this command with an id no other has. */
  revision(): RevisionInfo;
  /**
   * Is a revision one this author made since the document was opened: it carries the name, and its id is one that
   * was given for that name in this session. A revision that came with the file is nobody's, whatever name it
   * carries (somebody may be waiting to reject it), and neither is any revision for an editor without a name.
   */
  made(info: RevisionInfo | undefined): boolean;
  /** An id for the second of two elements a revision with the id `from` was cut into: see renewedId(). */
  renew(from: number | undefined): number;
}

/** The ids given to the revisions of each author since a document was opened. */
const SESSION_IDS = new WeakMap<DocModel, Map<number, string>>();

const systemClock = (): Date => new Date();

/**
 * A date and time as xsd:dateTime writes it (w:date is one): the seconds are part of it, the
 * day is one its month has in its year, there is no year 0000, and a time zone is "Z" or an
 * offset of hours and minutes of at most 14:00. Stricter than the type in one place: the hour
 * 24:00:00, which it allows for the end of a day, is not taken.
 */
export function isIsoDate(text: string): boolean {
  const match = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d):(\d\d)(?:\.\d{1,9})?(?:Z|[+-](\d\d):(\d\d))?$/.exec(text);
  if (!match || text.length > 40) return false;
  const [year, month, day, hour, minute, second] = [Number(match[1]), Number(match[2]), Number(match[3]), Number(match[4]), Number(match[5]), Number(match[6])];
  if (year < 1 || month < 1 || month > 12 || day < 1 || hour > 23 || minute > 59 || second > 59) return false;
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const days = month === 2 ? (leap ? 29 : 28) : month === 4 || month === 6 || month === 9 || month === 11 ? 30 : 31;
  if (day > days) return false;
  if (match[7] === undefined) return true;
  const [zoneHours, zoneMinutes] = [Number(match[7]), Number(match[8])];
  return zoneMinutes <= 59 && (zoneHours < 14 || (zoneHours === 14 && zoneMinutes === 0));
}

/** Does a command with these options track its changes in this document. */
export function isTracking(doc: DocModel, options?: TrackOptions): boolean {
  if (options && options.trackChanges === true) return true;
  const setup = SETUPS.get(ownerOf(doc));
  return (setup !== undefined && setup.force === true) || tracksChanges(doc);
}

/** The tracking for a command; undefined when its changes are not tracked. */
export function trackingFor(doc: DocModel, options?: TrackOptions): Tracking | undefined {
  if (!isTracking(doc, options)) return undefined;
  const setup = SETUPS.get(ownerOf(doc)) || {};
  const named = options && typeof options.author === 'string' && options.author !== '' ? options.author : setup.author;
  const known = typeof named === 'string' && named !== '';
  const author = known ? (named as string) : DEFAULT_AUTHOR;
  let date: string;
  if (options && options.date !== undefined && options.date !== '') {
    if (typeof options.date !== 'string' || !isIsoDate(options.date)) throw new RangeError('The date of a tracked change must be ISO 8601 text');
    date = options.date;
  } else date = revisionDate((setup.now || systemClock)());
  let given = SESSION_IDS.get(ownerOf(doc));
  if (!given) SESSION_IDS.set(ownerOf(doc), (given = new Map<number, string>()));
  const ids = given;
  return {
    author,
    named: known,
    date,
    revision: () => {
      const id = nextRevisionId(doc);
      ids.set(id, author);
      return { id, author, date };
    },
    made: (info) => known && info !== undefined && info.id !== undefined && info.author === author && ids.get(info.id) === author,
    renew: (from) => renewedId(doc, from),
  };
}

/**
 * Says that the revisions with these ids are ones `author` made since the document was opened (see Tracking.made).
 * Nothing a file brought is the editor's by itself; this is for a caller that knows better, and for tests of what
 * happens to an own insertion that holds more than typing can put there.
 */
export function claimRevisions(doc: DocModel, author: string, ids: number[]): void {
  let given = SESSION_IDS.get(ownerOf(doc));
  if (!given) SESSION_IDS.set(ownerOf(doc), (given = new Map<number, string>()));
  for (const id of ids) given.set(id, author);
}

/**
 * A new id for an element that is a part of the revision with the id `from` (an edit cut that in two). The part is
 * as much of this session as the revision was: whoever made the one made the other.
 */
function renewedId(doc: DocModel, from: number | undefined): number {
  const id = nextRevisionId(doc);
  const given = SESSION_IDS.get(ownerOf(doc));
  const author = given && from !== undefined ? given.get(from) : undefined;
  if (given && author !== undefined) given.set(id, author);
  return id;
}

/**
 * A revision like `info` for another element of the same change (the cell of a new row in a column that is a pending
 * change): the same author and date under an id of its own, and as much of this session as the one it is like.
 */
export function revisionLike(doc: DocModel, info: RevisionInfo): RevisionInfo {
  return { ...info, id: renewedId(doc, info.id) };
}

// ------------------------------------------------------------------------ ids

const NEXT_IDS = new WeakMap<DocModel, { next: number }>();
/** The largest id content brought into a document before its ids were first counted: see noteForeignIds(). */
const FOREIGN_IDS = new WeakMap<DocModel, number>();
/** The largest id an annotation can have (a signed 32-bit number, as the reader takes it). */
const MAX_ID = 0x7fffffff;

/** Wrappers around one inline that are looked at for ids. */
const MAX_WRAPPER_CHAIN = 4096;

function idsOfBlocks(blocks: Block[], note: (id: number | undefined) => void, depth: number): void {
  if (depth > 64) return;
  const ofProps = (props: RunProps | undefined): void => {
    if (!props) return;
    for (const key of ['inserted', 'deleted', 'moveFrom', 'moveTo'] as const) if (props[key]) note((props[key] as RevisionInfo).id);
  };
  for (const block of blocks) {
    if (block.type === 'mark') note(Number(block.mark.id));
    else if (block.type === 'paragraph') {
      ofProps(block.props.markProps);
      for (const inline of block.children) {
        // The chain is as deep as the source nesting, which the reader limits; a model made by hand may hold one that never ends.
        let steps = 0;
        for (let w = inline.wrap; w && steps < MAX_WRAPPER_CHAIN; w = w.parent, steps++) if (w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo') note(w.id);
        if (inline.type === 'run') {
          for (const item of inline.content) if (item.type === 'commentReference') note(item.id);
        } else if (inline.type !== 'math' && inline.type !== 'proofError' && inline.type !== 'empty' && inline.type !== 'misc' && inline.type !== 'unknown') note(Number(inline.id));
      }
    } else if (block.type === 'table') {
      if (block.extras) for (const extra of block.extras) idsOfElement(extra.el, note, 0);
      for (const row of block.rows) {
        if (row.props.inserted) note(row.props.inserted.id);
        if (row.props.deleted) note(row.props.deleted.id);
        if (row.extras) for (const extra of row.extras) idsOfElement(extra.el, note, 0);
        for (const cell of row.cells) {
          if (cell.props.inserted) note(cell.props.inserted.id);
          if (cell.props.deleted) note(cell.props.deleted.id);
          idsOfBlocks(cell.blocks, note, depth + 1);
        }
      }
      if (block.gridChange) note(block.gridChange.id);
    }
  }
}

/** The ids in an element the model keeps as read (a bookmark between the rows of a table, say). */
function idsOfElement(el: { attrs: Map<string, string>; children: unknown[] }, note: (id: number | undefined) => void, depth: number): void {
  if (depth > 64) return;
  el.attrs.forEach((value, name) => {
    if (name.slice(name.indexOf(':') + 1) === 'id' && /^\d{1,10}$/.test(value)) note(Number(value));
  });
  for (const node of el.children) {
    const child = node as { type?: string; attrs: Map<string, string>; children: unknown[] };
    if (child && child.type === 'element') idsOfElement(child, note, depth + 1);
  }
}

/**
 * The largest id the annotations of a document have (revisions, bookmarks,
 * comments, in every story). For a part read from a docx its source text is
 * searched, which also finds the ids in what the model keeps as XML; the
 * model is searched as well, for parts without a source.
 */
export function largestAnnotationId(doc: DocModel): number {
  let max = 0;
  const note = (id: number | undefined): void => {
    if (id !== undefined && id === id && id > max && id <= MAX_ID) max = id;
  };
  const sources: string[] = [];
  const lists: Block[][] = [doc.body.blocks];
  sources.push(doc.body.source);
  const story = (item: { source: string; blocks: Block[] }): void => {
    sources.push(item.source);
    lists.push(item.blocks);
  };
  doc.headers.forEach(story);
  doc.footers.forEach(story);
  for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) {
    if (!collection) continue;
    sources.push(collection.source);
    collection.items.forEach((item: { id: number; blocks: Block[] }) => {
      lists.push(item.blocks);
    });
  }
  if (doc.comments) doc.comments.items.forEach((_comment, id) => note(id));
  // As a file may write an attribute: either kind of quote, white space around the sign.
  const pattern = /:id\s*=\s*(?:"(\d{1,10})"|'(\d{1,10})')/g;
  for (const source of sources) {
    if (!source) continue;
    pattern.lastIndex = 0;
    for (let match = pattern.exec(source); match; match = pattern.exec(source)) note(Number(match[1] !== undefined ? match[1] : match[2]));
  }
  for (const blocks of lists) idsOfBlocks(blocks, note, 0);
  return max;
}

/**
 * Tells the counter of revision ids about content that comes into the document from elsewhere
 * (a paste): whatever ids it carries are out of reach for new revisions from here on - also
 * when the paste is undone and done again later, and whether or not an id was asked for before.
 * Costs by the size of the content, not of the document.
 */
export function noteForeignIds(doc: DocModel, blocks: Block[]): void {
  let max = 0;
  idsOfBlocks(
    blocks,
    (id) => {
      if (id !== undefined && id === id && id > max && id <= MAX_ID) max = id;
    },
    0,
  );
  if (max === 0) return;
  const state = NEXT_IDS.get(ownerOf(doc));
  if (state) {
    if (max >= state.next) state.next = max + 1;
  } else if (max > (FOREIGN_IDS.get(ownerOf(doc)) || 0)) FOREIGN_IDS.set(ownerOf(doc), max);
}

/**
 * An id for a new revision: larger than every id the document had when it
 * was first asked, than every id content brought since (noteForeignIds()),
 * and than every id given since. An id is never given twice, also not after
 * the change it was made for is undone.
 */
export function nextRevisionId(doc: DocModel): number {
  let state = NEXT_IDS.get(ownerOf(doc));
  if (!state) NEXT_IDS.set(ownerOf(doc), (state = { next: Math.max(largestAnnotationId(ownerOf(doc)), FOREIGN_IDS.get(ownerOf(doc)) || 0) + 1 }));
  if (state.next > MAX_ID) throw new RangeError('The document has no room for another revision');
  return state.next++;
}

// ------------------------------------------------------------------- wrappers

type RevisionWrapper = Extract<Wrapper, { kind: 'ins' | 'del' | 'moveFrom' | 'moveTo' }>;

const isRevision = (w: Wrapper): w is RevisionWrapper => w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo';

/** Is the content in an insertion this author tracked: such content is taken out for good. */
export function inOwnInsertion(wrap: Wrapper | undefined, author: string): boolean {
  for (let w = wrap; w; w = w.parent) if (w.kind === 'ins' && w.author === author) return true;
  return false;
}

function newWrapper(kind: 'ins' | 'del', revision: RevisionInfo, parent: Wrapper | undefined): RevisionWrapper {
  const wrapper = { kind, ...revision } as RevisionWrapper;
  if (parent) wrapper.parent = parent;
  return wrapper;
}

/**
 * A wrapper chain in which `target` is replaced by `replacement`; the
 * wrappers inside it are copied (once each, see `memo`), as they belong to
 * another element now.
 */
function swapped(wrap: Wrapper | undefined, target: Wrapper, replacement: Wrapper, memo: Map<Wrapper, Wrapper>): Wrapper | undefined {
  if (!wrap) return undefined;
  if (wrap === target) return replacement;
  const known = memo.get(wrap);
  if (known) return known;
  let inside = false;
  for (let w = wrap.parent; w; w = w.parent) if (w === target) inside = true;
  if (!inside) return wrap;
  const out = { ...wrap, parent: swapped(wrap.parent, target, replacement, memo) } as Wrapper;
  memo.set(wrap, out);
  return out;
}

const chainHas = (wrap: Wrapper | undefined, target: Wrapper): boolean => {
  for (let w = wrap; w; w = w.parent) if (w === target) return true;
  return false;
};

/**
 * The pieces with every revision wrapper of `shared` replaced by a copy with
 * an id of its own: what the second half of a split paragraph, or of a split
 * insertion, gets, so that no two elements of the file have the same id.
 */
export function renumbered(pieces: Piece[], shared: Set<Wrapper>, tracking: Tracking): Piece[] {
  if (shared.size === 0) return pieces;
  const memo = new Map<Wrapper, Wrapper | undefined>();
  const copy = (wrap: Wrapper | undefined): Wrapper | undefined => {
    if (!wrap) return undefined;
    if (memo.has(wrap)) return memo.get(wrap);
    const parent = copy(wrap.parent);
    let out = wrap;
    if (shared.has(wrap)) out = { ...wrap, id: tracking.renew((wrap as RevisionWrapper).id) } as Wrapper;
    else if (parent !== wrap.parent) out = { ...wrap } as Wrapper;
    if (out !== wrap) {
      if (parent) out.parent = parent;
      else delete out.parent;
    }
    memo.set(wrap, out);
    return out;
  };
  let out: Piece[] | undefined;
  for (let i = 0; i < pieces.length; i++) {
    const wrap = copy(pieces[i].wrap);
    if (wrap === pieces[i].wrap) continue;
    if (!out) out = pieces.slice();
    out[i] = retargeted(pieces[i], wrap);
  }
  return out || pieces;
}

/** A piece in another wrapper chain. */
function retargeted(piece: Piece, wrap: Wrapper | undefined): Piece {
  const out: Piece = { ...piece, wrap };
  if (piece.inline) {
    const inline = { ...piece.inline, wrap } as Inline;
    if (!wrap) delete inline.wrap;
    out.inline = inline;
    out.group = inline;
  }
  return out;
}

/** The revision wrappers (with an id) around the pieces. */
export function revisionsAround(pieces: Piece[]): Set<Wrapper> {
  const out = new Set<Wrapper>();
  for (const piece of pieces) for (let w = piece.wrap; w; w = w.parent) if (isRevision(w)) out.add(w);
  return out;
}

/** The pieces of a paragraph after a cut, without revision ids the pieces before the cut already have. */
export function afterCut(head: Piece[], tail: Piece[], tracking: Tracking, keep?: Wrapper): Piece[] {
  const before = revisionsAround(head);
  if (before.size === 0) return tail;
  // What is put in between is in these: they are not cut.
  for (let w = keep; w; w = w.parent) before.delete(w);
  const shared = new Set<Wrapper>();
  revisionsAround(tail).forEach((w) => (before.has(w) ? shared.add(w) : 0));
  return renumbered(tail, shared, tracking);
}

/** Wrappers that are one thing with an identity and data of its own: writing one twice makes two of it. */
const isIdentified = (w: Wrapper): boolean => w.kind === 'sdt' || w.kind === 'customXml';

/** Refuses a cut that would copy a content control: the wrappers from `from` up to (not including) `until`. */
function refuseCopies(from: Wrapper | undefined, until: Wrapper): void {
  for (let w = from; w && w !== until; w = w.parent) if (isIdentified(w)) throw new RangeError(TRACK_UNSUPPORTED);
}

/**
 * Does text the editor types go on in an insertion: one they made since the document was opened. An insertion a file
 * brought under the same name is not theirs to add to (see Tracking.made): what is typed beside it is an insertion of
 * its own, which can be taken back for good. An editor without a name owns none, and goes on in any that carries the
 * stand-in name, so that typing makes one insertion and not one for every key.
 */
function continues(tracking: Tracking, wrapper: RevisionWrapper): boolean {
  if (wrapper.author !== tracking.author) return false;
  return !tracking.named || tracking.made(wrapper);
}

/**
 * The wrapper text put before the piece at `index` goes into, given the one
 * it would take untracked (`base`, see wrapBetween): an insertion of this
 * author beside it is continued (it keeps its date), text inside the
 * insertion or move of somebody else cuts that in two (the second half gets
 * an id of its own; refused with TRACK_UNSUPPORTED when a content control
 * would be cut with it), anything else gets a new insertion. Returns the pieces
 * too: the list itself unless a cut changed what follows.
 */
export function insertionWrap(tracking: Tracking, pieces: Piece[], index: number, base: Wrapper | undefined): { wrap: Wrapper; pieces: Piece[] } {
  let around: RevisionWrapper | undefined;
  for (let w = base; w; w = w.parent) {
    if (w.kind === 'ins' || w.kind === 'moveTo') {
      around = w;
      break;
    }
  }
  if (around) {
    if (around.kind === 'ins' && continues(tracking, around)) return { wrap: base as Wrapper, pieces };
    // Every wrapper between the text and the insertion that is cut is cut with it.
    refuseCopies(base, around);
    const mine = newWrapper('ins', tracking.revision(), around.parent);
    const wrap = swapped(base, around, mine, new Map()) as Wrapper;
    // What follows in the same element is another element now.
    let out = pieces;
    let any = false;
    for (let i = index; i < pieces.length && !any; i++) any = chainHas(pieces[i].wrap, around);
    if (any) {
      const second = { ...around, id: tracking.renew(around.id) } as Wrapper;
      const memo = new Map<Wrapper, Wrapper>();
      out = pieces.slice();
      for (let i = index; i < out.length; i++) {
        if (chainHas(out[i].wrap, around)) out[i] = retargeted(out[i], swapped(out[i].wrap, around, second, memo));
      }
    }
    return { wrap, pieces: out };
  }
  for (const near of [index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined]) {
    for (let w = near; w; w = w.parent) {
      if (w.kind === 'ins' && continues(tracking, w) && w.parent === base) return { wrap: w, pieces };
    }
  }
  return { wrap: newWrapper('ins', tracking.revision(), base), pieces };
}

/** A new insertion like `wrap` (which is, or is in, an insertion of this command) for the next paragraph: the same nesting, another id. */
export function insertionLike(tracking: Tracking, wrap: Wrapper): Wrapper {
  let around: Wrapper | undefined;
  for (let w: Wrapper | undefined = wrap; w; w = w.parent) {
    if (w.kind === 'ins') {
      around = w;
      break;
    }
  }
  if (!around) return newWrapper('ins', tracking.revision(), wrap);
  // A content control inside the insertion would be written once for every paragraph.
  refuseCopies(wrap, around);
  return swapped(wrap, around, newWrapper('ins', tracking.revision(), around.parent), new Map()) as Wrapper;
}

// ------------------------------------------------------------------ deletions

const isOwnDeletion = (wrap: Wrapper | undefined, author: string): wrap is RevisionWrapper => wrap !== undefined && wrap.kind === 'del' && wrap.author === author;

/** An item as it is inside a deletion: text becomes <w:delText>, field code text <w:delInstrText>. */
function struckItem(item: RunItem): RunItem {
  if (item.type === 'text') return item.source === undefined ? { type: 'deletedText', text: item.text } : { type: 'deletedText', text: item.text, source: item.source };
  if (item.type === 'instrText' && item.deleted !== true) return { ...item, deleted: true };
  return item;
}

const DELETED_PLAIN = new Set<string>(['deletedText', 'tab', 'break', 'noBreakHyphen', 'softHyphen', 'symbol']);

const sameBag = (a: unknown, b: unknown): boolean => {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) return false;
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = Object.keys(left);
  if (keys.length !== Object.keys(right).length) return false;
  // One level: formatting that differs deeper down keeps its runs apart, which costs nothing but a longer file.
  for (const key of keys) if (left[key] !== right[key]) return false;
  return true;
};

/** Neighbouring deleted runs of text in one deletion, formatted alike, become one run. */
function joinStruck(pieces: Piece[]): Piece[] {
  const out: Piece[] = [];
  for (const piece of pieces) {
    const last = out.length ? out[out.length - 1] : undefined;
    const a = last && last.frozen ? last.inline : undefined;
    const b = piece.frozen ? piece.inline : undefined;
    if (
      a &&
      b &&
      a.type === 'run' &&
      b.type === 'run' &&
      a.wrap !== undefined &&
      a.wrap === b.wrap &&
      a.wrap.kind === 'del' &&
      a.fieldCode === undefined &&
      b.fieldCode === undefined &&
      a.content.every((item) => DELETED_PLAIN.has(item.type)) &&
      b.content.every((item) => DELETED_PLAIN.has(item.type)) &&
      sameBag(a.props, b.props) &&
      sameBag(a.attrs, b.attrs)
    ) {
      const content = a.content.slice();
      for (const item of b.content) {
        const end = content[content.length - 1];
        if (item.type === 'deletedText' && end && end.type === 'deletedText' && item.source === undefined && end.source === undefined) content[content.length - 1] = { type: 'deletedText', text: end.text + item.text };
        else content.push(item);
      }
      const run: Inline = { ...a, content };
      out[out.length - 1] = { ...last, inline: run, group: run } as Piece;
    } else out.push(piece);
  }
  return out;
}

/**
 * The pieces with the content of [from, to) deleted as a tracked change:
 * what removePieces() would take is put into a deletion (and takes no
 * offsets any more), except for the text, tabs and line breaks this author
 * had inserted as a tracked change since the document was opened, which go - unless a link, a control or
 * the like around the insertion would be left with nothing: then all of the
 * author's text in it stays, deleted. `changed` is false when there was
 * nothing to delete.
 */
export function strikePieces(tracking: Tracking, pieces: Piece[], from: number, to: number): { pieces: Piece[]; changed: boolean } {
  if (from >= to) return { pieces, changed: false };
  const flags = removedPieces(pieces, from, to);
  const author = tracking.author;
  const fields = tracking.named ? fieldPairs(pieces).owner : undefined;
  /**
   * Typed text that sits directly in this author's own insertion: the only thing that goes for
   * good. Text in anything else inside the insertion (a simple field, a link, a content control,
   * a smart tag) is deleted where it is, so that what holds it stays.
   */
  const candidate = (q: number): boolean => {
    const piece = pieces[q];
    const item = piece.item;
    const wrap = piece.wrap;
    if (!fields || !item || piece.code || fields[q] >= 0 || !wrap || wrap.kind !== 'ins' || !tracking.made(wrap)) return false;
    return item.type === 'text' || item.type === 'tab' || (item.type === 'break' && item.kind === 'line');
  };
  const gone: boolean[] = [];
  let anyGone = false;
  for (let q = 0; q < pieces.length; q++) {
    gone.push(flags[q] && candidate(q));
    if (gone[q]) anyGone = true;
  }
  if (anyGone) {
    // An insertion typed into a link or a control must not take that with it: every wrapper around
    // the insertion has to hold something that stays, or the text stays too, deleted. The answer is
    // one for each wrapper, whatever the runs of the text are: a wrapper that would be left with
    // nothing keeps all of the text it holds, as a deletion, and from one that keeps something else
    // all of it goes. (Key by key that means: the characters go for good until the last one, which
    // stays, deleted.)
    const held = new Set<Wrapper>();
    for (let q = 0; q < pieces.length; q++) {
      if (gone[q]) continue;
      for (let w = pieces[q].wrap; w && !held.has(w); w = w.parent) held.add(w);
    }
    for (let q = 0; q < pieces.length; q++) {
      if (!gone[q]) continue;
      let steps = 0;
      for (let w = (pieces[q].wrap as Wrapper).parent; w && steps < MAX_WRAPPER_CHAIN; w = w.parent, steps++) {
        if (held.has(w)) continue;
        gone[q] = false;
        break;
      }
    }
  }
  const own = (q: number): boolean => gone[q];
  const out: Piece[] = [];
  let changed = false;
  /** The deletion of this author the piece before ends in. */
  let open: RevisionWrapper | undefined;
  /** Item pieces being collected for one deleted run. */
  let items: Piece[] = [];
  const flush = (): void => {
    if (items.length === 0) return;
    const run = makeRun(items, 0, items.length);
    out.push({ inline: run, props: run.props, wrap: run.wrap, group: run, index: -1, code: false, len: 0, frozen: true });
    items = [];
  };
  for (let q = 0; q < pieces.length; q++) {
    const piece = pieces[q];
    if (!flags[q]) {
      flush();
      out.push(piece);
      open = piece.frozen && isOwnDeletion(piece.wrap, author) ? piece.wrap : undefined;
      continue;
    }
    changed = true;
    if (own(q)) continue;
    const parent = piece.wrap;
    let del: RevisionWrapper;
    if (open && open.parent === parent) del = open;
    else {
      // A deletion of this author that follows what is deleted here takes it in.
      let n = q + 1;
      while (n < pieces.length && flags[n] && pieces[n].wrap === parent) n++;
      const next = n < pieces.length ? pieces[n] : undefined;
      del = next && next.frozen && isOwnDeletion(next.wrap, author) && next.wrap.parent === parent ? next.wrap : newWrapper('del', tracking.revision(), parent);
    }
    open = del;
    if (piece.inline) {
      flush();
      const inline = { ...piece.inline, wrap: del } as Inline;
      out.push({ ...piece, inline, group: inline, wrap: del, len: 0, frozen: true });
      continue;
    }
    const last = items.length ? items[items.length - 1] : undefined;
    if (last && (last.group !== piece.group || last.props !== piece.props || last.wrap !== del)) flush();
    items.push({ ...piece, item: struckItem(piece.item as RunItem), wrap: del, len: 0 });
  }
  flush();
  return changed ? { pieces: joinStruck(out), changed } : { pieces, changed };
}

// ------------------------------------------------------------ paragraph marks

const MARK_FLAGS = ['inserted', 'deleted', 'moveFrom', 'moveTo'] as const;

/** Paragraph properties whose mark says it was inserted or deleted (or, with undefined, does not say so). */
export function withMarkRevision(props: ParaProps, key: 'inserted' | 'deleted', revision: RevisionInfo | undefined): ParaProps {
  const mark: RunProps = { ...(props.markProps || {}) };
  if (revision) mark[key] = revision;
  else delete mark[key];
  const out: ParaProps = { ...props };
  if (Object.keys(mark).length) out.markProps = mark;
  else delete out.markProps;
  return out;
}

/** Paragraph properties whose mark is a revision exactly as far as the mark of `from` is one. */
export function withMarkFlagsOf(props: ParaProps, from: ParaProps): ParaProps {
  const have = props.markProps || {};
  const want = from.markProps || {};
  let same = true;
  for (const key of MARK_FLAGS) if (have[key] !== want[key]) same = false;
  if (same) return props;
  const mark: RunProps = { ...have };
  for (const key of MARK_FLAGS) {
    if (want[key]) mark[key] = want[key];
    else delete mark[key];
  }
  const out: ParaProps = { ...props };
  if (Object.keys(mark).length) out.markProps = mark;
  else delete out.markProps;
  return out;
}

/** Paragraph properties for a paragraph an edit makes: its mark is no revision yet. */
export function withPlainMark(props: ParaProps): ParaProps {
  return withMarkFlagsOf(props, {});
}

const withoutRevisionFlags = (mark: RunProps | undefined): RunProps | undefined => {
  if (!mark) return undefined;
  const out: RunProps = { ...mark };
  for (const key of MARK_FLAGS) delete out[key];
  return Object.keys(out).length ? out : undefined;
};

/**
 * Would joining the paragraph of `first` onto the one of `second` lose no formatting: the two
 * are formatted alike (style, list, alignment, indents, the formatting of their marks, kept
 * XML), whatever their marks say of revisions, and the first ends no section. Kept XML is alike
 * only when it is the same object.
 */
export function joinsWithoutLoss(first: ParaProps, second: ParaProps): boolean {
  if (first.section !== undefined) return false;
  const a: ParaProps = { ...first };
  const b: ParaProps = { ...second };
  delete a.markProps;
  delete b.markProps;
  delete b.section;
  return sameProps(a, b) && sameProps(withoutRevisionFlags(first.markProps), withoutRevisionFlags(second.markProps));
}

/** Was the mark of a paragraph inserted by this author as a tracked change since the document was opened: joining it to a paragraph formatted alike takes the mark out for good. */
export function isOwnMark(props: ParaProps, tracking: Tracking): boolean {
  const mark = props.markProps;
  return mark !== undefined && tracking.made(mark.inserted);
}

/** A kept <w:pPrChange> or <w:rPrChange> with an id of its own. */
function freshChange(change: XmlElement, tracking: Tracking): XmlElement {
  const attrs = new Map(change.attrs);
  attrs.forEach((_value, name) => {
    if (name.slice(name.indexOf(':') + 1) === 'id') attrs.set(name, String(tracking.renew(undefined)));
  });
  return sameRecord(change, { ...change, attrs });
}

/**
 * Kept children with the ids of the tracked changes in them replaced: `renew` is asked for
 * every id and answers with the id to write instead, or with undefined to leave it. The list
 * itself when nothing changes.
 */
function renumberedExtras(extras: ExtraXml[] | undefined, renew: (id: number) => number | undefined): ExtraXml[] | undefined {
  if (!extras) return extras;
  const walk = (el: XmlElement, depth: number): XmlElement => {
    if (depth > MAX_PLAIN_DEPTH) return el;
    let attrs: Map<string, string> | undefined;
    const local = localOfName(el.name);
    if (REVISION_RECORDS.has(local) || REVISION_MARKS.has(local)) {
      el.attrs.forEach((value, name) => {
        if (attrs || localOfName(name) !== 'id' || !/^\d{1,10}$/.test(value)) return;
        const next = renew(Number(value));
        if (next === undefined) return;
        attrs = new Map(el.attrs);
        attrs.set(name, String(next));
      });
    }
    let children: XmlElement['children'] | undefined;
    for (let i = 0; i < el.children.length; i++) {
      const node = el.children[i];
      if (node.type !== 'element') continue;
      const next = walk(node, depth + 1);
      if (next === node) continue;
      if (!children) children = el.children.slice();
      children[i] = next;
    }
    if (!attrs && !children) return el;
    return { ...el, attrs: attrs || el.attrs, children: children || el.children };
  };
  let out: ExtraXml[] | undefined;
  for (let i = 0; i < extras.length; i++) {
    const el = walk(extras[i].el, 0);
    if (el === extras[i].el) continue;
    if (!out) out = extras.slice();
    out[i] = { ...extras[i], el };
  }
  return out || extras;
}

/**
 * Paragraph properties for a second paragraph with the same properties: the
 * formatting changes they keep (of the paragraph and of its mark, and the
 * changes inside what they keep as XML, a changed list number say) get ids
 * of their own, so that no id is in the file twice.
 */
export function withFreshChanges(props: ParaProps, tracking: Tracking): ParaProps {
  const mark = props.markProps;
  const renew = (): number => tracking.renew(undefined);
  const extras = renumberedExtras(props.extras, renew);
  const markExtras = mark ? renumberedExtras(mark.extras, renew) : undefined;
  if (!props.change && !(mark && mark.change) && extras === props.extras && !(mark && markExtras !== mark.extras)) return props;
  const out: ParaProps = { ...props };
  if (props.change) out.change = freshChange(props.change, tracking);
  if (extras !== props.extras) out.extras = extras;
  if (mark && (mark.change || markExtras !== mark.extras)) {
    const next: RunProps = { ...mark };
    if (mark.change) next.change = freshChange(mark.change, tracking);
    if (markExtras !== mark.extras) next.extras = markExtras;
    out.markProps = next;
  }
  return out;
}

/**
 * Paragraph properties whose mark is, as a tracked change, what the mark of `from` is: the
 * revisions it says of itself (inserted, deleted, moved) and the record of a change of its
 * formatting. The formatting itself stays what `props` has.
 */
export function withMarkRevisionsOf(props: ParaProps, from: ParaProps): ParaProps {
  const flagged = withMarkFlagsOf(props, from);
  const have = flagged.markProps ? flagged.markProps.change : undefined;
  const want = from.markProps ? from.markProps.change : undefined;
  if (have === want) return flagged;
  const mark: RunProps = { ...(flagged.markProps || {}) };
  if (want) mark.change = want;
  else delete mark.change;
  const out: ParaProps = { ...flagged };
  if (Object.keys(mark).length) out.markProps = mark;
  else delete out.markProps;
  return Object.keys(out).length ? out : EMPTY_PARA_PROPS;
}

export const isDeletedMark = (props: ParaProps): boolean => props.markProps !== undefined && props.markProps.deleted !== undefined;

// ---------------------------------------------------------- formatting changes

/** What is known of a record of changed formatting this session made: who made it, and the properties it holds, as the model has them. */
interface OwnRecord {
  author: string;
  old: object;
}

/**
 * The records of changed formatting made since the document was opened. A record a file brought is not among them:
 * whoever it names, it is not taken away by formatting.
 */
const OWN_RECORDS = new WeakMap<XmlElement, OwnRecord>();

/** A copy of a record (the same change, another id: its run or paragraph was split) is as much the author's as the record. */
function sameRecord(from: XmlElement, copy: XmlElement): XmlElement {
  const own = OWN_RECORDS.get(from);
  if (own) OWN_RECORDS.set(copy, own);
  return copy;
}

/** What a part needs to be written into: how it spells its elements. */
interface PartLike {
  part: string;
  rootAttrs: AttrList;
  relationships: Map<string, DocRelationship>;
}

const WRITERS = new WeakMap<object, PropsContext>();

/** A writer of properties for the part of a story: records are spelled with the prefixes that part declares. */
function writerFor(host: PartLike): PropsContext {
  let ctx = WRITERS.get(host);
  if (!ctx) {
    const state = createReaderState();
    const reader = createContext(state, host.part, host.rootAttrs, host.relationships);
    // As the writer of the package finds it: the prefix bound to WordprocessingML, none when it is the default namespace.
    const w = reader.wPrefixes.length ? reader.wPrefixes[0] : reader.prefixes.get('') === 'w' ? '' : 'w:';
    ctx = { w, a: w, r: reader.rPrefixes[0], strict: state.strict, reader };
    WRITERS.set(host, ctx);
  }
  return ctx;
}

function recordElement(tracking: Tracking, ctx: PropsContext, local: 'rPrChange' | 'pPrChange', inner: XmlElement, old: object): XmlElement {
  const revision = tracking.revision();
  const el = wEl(
    ctx,
    local,
    [
      ['id', revision.id],
      ['author', revision.author],
      ['date', revision.date],
    ],
    [inner],
  );
  OWN_RECORDS.set(el, { author: tracking.author, old });
  return el;
}

/** Is the record one this author made in this session, and are the properties back at what it holds. */
function undoes(tracking: Tracking, record: XmlElement, now: object): boolean {
  const own = OWN_RECORDS.get(record);
  return own !== undefined && tracking.named && own.author === tracking.author && sameProps(now, own.old);
}

function withoutChange<T extends { change?: XmlElement }>(props: T, empty: T): T {
  const out = { ...props };
  delete out.change;
  return Object.keys(out).length ? out : empty;
}

/**
 * The run properties a tracked change of formatting leaves: `after` is what the command made of `before` (the same
 * object when it changed nothing), `host` the part the run is in. `own` says that the text is in a tracked insertion
 * the author made in this session (or the properties are those of the mark of a paragraph that is all the author's,
 * see isOwnParagraph): no record then. Properties that
 * carry a record keep it and get no other (Word keeps one record: the first "as it was"); the author's own record of
 * this session goes when the properties are back at what it holds. UNSURE: Word's rule for taking its own record
 * away was not checked against Word; this one never loses a record somebody else may want to reject.
 */
export function trackedRunProps(tracking: Tracking, host: PartLike, before: RunProps, after: RunProps, own: boolean): RunProps {
  if (after === before || own) return after;
  const record = before.change;
  if (record) {
    if (after.change !== record) return after;
    const bare = withoutChange(after, EMPTY_RUN_PROPS);
    return undoes(tracking, record, plainProps(bare)) ? bare : after;
  }
  // What it was, as formatting and nothing else: a mark that is a revision itself says so outside the record.
  const old = plainProps(before);
  const ctx = writerFor(host);
  return { ...after, change: recordElement(tracking, ctx, 'rPrChange', runPropsElement(ctx, old) || wEl(ctx, 'rPr'), old) };
}

/** Paragraph properties as a record of their change holds them: without the mark, the section the paragraph ends and any tracked change. */
function recordedParaProps(props: ParaProps): ParaProps {
  let bare = props;
  if (props.markProps !== undefined || props.section !== undefined || props.change !== undefined) {
    bare = { ...props };
    delete bare.markProps;
    delete bare.section;
    delete bare.change;
  }
  return plainProps(bare);
}

/** Was the content put in by this author, as a tracked change, since the document was opened. */
export function inSessionInsertion(wrap: Wrapper | undefined, tracking: Tracking): boolean {
  let steps = 0;
  for (let w = wrap; w && steps < MAX_WRAPPER_CHAIN; w = w.parent, steps++) if (w.kind === 'ins' && tracking.made(w)) return true;
  return false;
}

/**
 * Is a paragraph all the author's, made in this session: its mark is a tracked insertion of theirs and so is
 * everything in it. Such a paragraph has no formatting it used to have. One the author only split is not: the half
 * that got the new mark holds text that was there, and a change of its formatting has to be one that can be rejected.
 */
export function isOwnParagraph(paragraph: { props: ParaProps; children: Inline[] }, tracking: Tracking): boolean {
  const mark = paragraph.props.markProps;
  if (!mark || !tracking.made(mark.inserted)) return false;
  for (const inline of paragraph.children) if (!inSessionInsertion(inline.wrap, tracking)) return false;
  return true;
}

/** The kept children that hold a tracked change (a list number somebody inserted or changed, say). */
export function revisionExtras(extras: ExtraXml[] | undefined): ExtraXml[] {
  const out: ExtraXml[] = [];
  if (!extras) return out;
  for (const extra of extras) {
    const one = [extra];
    // Without its tracked changes it would be another element, or none.
    if (plainExtras(one, true) !== one) out.push(extra);
  }
  return out;
}

/**
 * The paragraph properties a tracked change of formatting leaves (alignment, indents, spacing, style, list); see
 * trackedRunProps. `own`: the paragraph is all the author's (isOwnParagraph), and gets no record. A change that would
 * take a tracked change of somebody else out of the properties (a paragraph leaves a list whose number carries one)
 * is refused with TRACK_UNSUPPORTED: no record could bring that back.
 */
export function trackedParaProps(tracking: Tracking, host: PartLike, before: ParaProps, after: ParaProps, own: boolean): ParaProps {
  if (after === before) return after;
  const kept = revisionExtras(before.extras);
  if (kept.length) {
    const still = after.extras || [];
    for (const extra of kept) if (still.indexOf(extra) < 0) throw new RangeError(TRACK_UNSUPPORTED);
  }
  if (own) return after;
  const record = before.change;
  if (record) {
    if (after.change !== record || !undoes(tracking, record, recordedParaProps(after))) return after;
    return withoutChange(after, EMPTY_PARA_PROPS);
  }
  const old = recordedParaProps(before);
  const ctx = writerFor(host);
  return { ...after, change: recordElement(tracking, ctx, 'pPrChange', paraPropsElement(ctx, old) || wEl(ctx, 'pPr'), old) };
}

// ------------------------------------------------------- one id, one element

const changeId = (change: XmlElement): number | undefined => {
  let id: number | undefined;
  change.attrs.forEach((value, name) => {
    if (id === undefined && name.slice(name.indexOf(':') + 1) === 'id' && /^\d{1,10}$/.test(value)) id = Number(value);
  });
  return id;
};

function renumberedChange(doc: DocModel, change: XmlElement): XmlElement {
  const attrs = new Map(change.attrs);
  let done = false;
  attrs.forEach((_value, name) => {
    if (!done && name.slice(name.indexOf(':') + 1) === 'id') {
      attrs.set(name, String(nextRevisionId(doc)));
      done = true;
    }
  });
  return sameRecord(change, { ...change, attrs });
}

/**
 * The blocks an edit puts into the document, with every tracked change in
 * their paragraphs an element of its own: what the edit split in two (a run
 * whose formatting was changed, an insertion, a deletion or a move something
 * was put into, a paragraph that carries a change of its formatting or of its
 * mark, or one inside the properties it keeps as XML) would otherwise be
 * written as two elements with one id. The second of two such elements gets a
 * new id. Holds for every edit, tracked or not. Paragraphs inside the tables
 * of the blocks are not looked at.
 *
 * `untouched` are the blocks among them that the document has already, as
 * they are: the writer copies those from the source of their part, so they
 * are never changed here (a file that came with an id twice keeps it there),
 * and the ids they have are taken for all the others. Returns the list itself
 * when nothing had to change.
 */
export function distinctRevisions(doc: DocModel, blocks: Block[], untouched?: Set<Block>): Block[] {
  let out: Block[] | undefined;
  /** The ids the elements written so far have. */
  const seen = new Set<number>();
  const taken = (id: number | undefined): boolean => {
    if (id === undefined) return false;
    if (seen.has(id)) return true;
    seen.add(id);
    return false;
  };
  const noteExtras = (extras: ExtraXml[] | undefined): void => {
    renumberedExtras(extras, (id) => {
      seen.add(id);
      return undefined;
    });
  };
  if (untouched && untouched.size) {
    for (const block of blocks) {
      if (block.type !== 'paragraph' || !untouched.has(block)) continue;
      for (const inline of block.children) {
        let steps = 0;
        for (let w = inline.wrap; w && steps < MAX_WRAPPER_CHAIN; w = w.parent, steps++) if (isRevision(w) && w.id !== undefined) seen.add(w.id);
        if (inline.type === 'run' && inline.props.change) taken(changeId(inline.props.change));
      }
      const props = block.props;
      if (props.change) taken(changeId(props.change));
      noteExtras(props.extras);
      const mark = props.markProps;
      if (mark) {
        if (mark.change) taken(changeId(mark.change));
        for (const key of MARK_FLAGS) {
          const info: RevisionInfo | undefined = mark[key];
          if (info) taken(info.id);
        }
        noteExtras(mark.extras);
      }
    }
  }
  const renew = (id: number): number | undefined => (taken(id) ? nextRevisionId(doc) : undefined);
  for (let b = 0; b < blocks.length; b++) {
    const block = blocks[b];
    if (block.type !== 'paragraph' || (untouched && untouched.has(block))) continue;
    let children: Inline[] | undefined;
    /** The wrappers around the inline before, and what stands for each in the element that is open. */
    let previous = new Set<Wrapper>();
    const alias = new Map<Wrapper, Wrapper>();
    for (let i = 0; i < block.children.length; i++) {
      const inline = block.children[i];
      let next = inline;
      const chain: Wrapper[] = [];
      for (let w = inline.wrap; w && chain.length < MAX_WRAPPER_CHAIN; w = w.parent) chain.push(w);
      let use: Wrapper | undefined;
      // Outermost first: a wrapper that was not around the inline before begins a new element here.
      for (let k = chain.length - 1; k >= 0; k--) {
        const w = chain[k];
        if (previous.has(w)) {
          use = alias.get(w) || w;
          continue;
        }
        const fresh = isRevision(w) && taken(w.id);
        if (fresh || use !== w.parent) {
          const copy = { ...w } as Wrapper;
          if (use) copy.parent = use;
          else delete copy.parent;
          if (fresh) (copy as RevisionWrapper).id = renewedId(doc, (w as RevisionWrapper).id);
          alias.set(w, copy);
          use = copy;
        } else {
          alias.delete(w);
          use = w;
        }
      }
      previous = new Set<Wrapper>(chain);
      if (use !== inline.wrap) next = { ...inline, wrap: use } as Inline;
      if (next.type === 'run' && next.props.change && taken(changeId(next.props.change))) {
        next = { ...next, props: { ...next.props, change: renumberedChange(doc, next.props.change) } };
      }
      if (next !== inline) {
        if (!children) children = block.children.slice();
        children[i] = next;
      }
    }
    let props = block.props;
    if (props.change && taken(changeId(props.change))) props = { ...props, change: renumberedChange(doc, props.change) };
    const extras = renumberedExtras(props.extras, renew);
    if (extras !== props.extras) props = { ...props, extras };
    let mark = props.markProps;
    if (mark) {
      const before = mark;
      if (mark.change && taken(changeId(mark.change))) mark = { ...mark, change: renumberedChange(doc, mark.change) };
      for (const key of MARK_FLAGS) {
        const info: RevisionInfo | undefined = mark[key];
        if (info && taken(info.id)) mark = { ...mark, [key]: { ...info, id: renewedId(doc, info.id) } };
      }
      const markExtras = renumberedExtras(mark.extras, renew);
      if (markExtras !== mark.extras) mark = { ...mark, extras: markExtras };
      if (mark !== before) props = { ...props, markProps: mark };
    }
    if (!children && props === block.props) continue;
    if (!out) out = blocks.slice();
    out[b] = { ...block, children: children || block.children, props };
  }
  return out || blocks;
}

// ------------------------------------------------- properties without changes

const MAX_PLAIN_DEPTH = 300;
const localOfName = (name: string): string => name.slice(name.indexOf(':') + 1);

/** Elements that record a tracked change of formatting or of a table, wherever they stand. */
const REVISION_RECORDS = new Set(['rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange', 'tblPrExChange', 'trPrChange', 'tcPrChange', 'tblGridChange', 'numberingChange', 'cellIns', 'cellDel', 'cellMerge']);
/** Inside properties these mark what the properties belong to (a paragraph mark, a row, a numbering) as inserted, deleted or moved. */
const REVISION_MARKS = new Set(['ins', 'del', 'moveFrom', 'moveTo']);
/** The fields of property bags that hold a tracked change. */
const REVISION_FIELDS = ['inserted', 'deleted', 'moveFrom', 'moveTo', 'change'];

/** Property bags without their tracked changes, by the bag they were made from: the same bag gives the same object. */
const PLAIN = new WeakMap<object, object>();

/** A kept element without the records of tracked changes in it; the element itself when it holds none. */
function plainElement(el: XmlElement, marks: boolean, depth = 0): XmlElement {
  if (depth > MAX_PLAIN_DEPTH) return el;
  let children: XmlElement['children'] | undefined;
  for (let i = 0; i < el.children.length; i++) {
    const node = el.children[i];
    if (node.type !== 'element') {
      if (children) children.push(node);
      continue;
    }
    const local = localOfName(node.name);
    const next = REVISION_RECORDS.has(local) || (marks && REVISION_MARKS.has(local)) ? undefined : plainElement(node, marks, depth + 1);
    if (next === node) {
      if (children) children.push(node);
      continue;
    }
    if (!children) children = el.children.slice(0, i);
    if (next) children.push(next);
  }
  return children ? { ...el, children } : el;
}

/**
 * Kept children without the records of tracked changes. `marks`: the children are those of
 * properties, where <w:ins> and <w:del> are marks too (in a table they would wrap rows).
 */
export function plainExtras(extras: ExtraXml[] | undefined, marks: boolean): ExtraXml[] | undefined {
  if (!extras) return undefined;
  let out: ExtraXml[] | undefined;
  for (let i = 0; i < extras.length; i++) {
    const extra = extras[i];
    const local = localOfName(extra.el.name);
    // The parent itself (index -1) is kept for what it carries; only what is in it goes.
    const gone = extra.index >= 0 && (REVISION_RECORDS.has(local) || (marks && REVISION_MARKS.has(local)));
    const el = gone ? undefined : plainElement(extra.el, marks);
    if (el === extra.el) {
      if (out) out.push(extra);
      continue;
    }
    if (!out) out = extras.slice(0, i);
    // A list number that was kept for the change recorded in it is, without that, what the fields of the model say.
    if (el && !(extra.partial === true && isPlainNumPr(el))) out.push({ ...extra, el });
  }
  if (!out) return extras;
  return out.length ? out : undefined;
}

/** Is the element a <w:numPr> that says nothing but a level and a list: all of it is in ParaProps.numbering then. */
function isPlainNumPr(el: XmlElement): boolean {
  if (localOfName(el.name) !== 'numPr' || el.attrs.size) return false;
  const seen = new Set<string>();
  for (const node of el.children) {
    if (node.type !== 'element') return false;
    const local = localOfName(node.name);
    if ((local !== 'ilvl' && local !== 'numId') || seen.has(local) || node.children.length || node.attrs.size !== 1) return false;
    let value: string | undefined;
    node.attrs.forEach((text, name) => {
      if (localOfName(name) === 'val') value = text;
    });
    if (value === undefined || !/^\d{1,10}$/.test(value)) return false;
    seen.add(local);
  }
  return true;
}

/**
 * Properties (of a paragraph, a run, a section, a table, a row, a cell) as a copy carries them:
 * without what makes them a tracked change (see the head of fragment.ts). It is also what the
 * properties of anything an edit makes anew are: a new paragraph mark, a new row, a new cell.
 */
export function plainProps<T extends object>(props: T): T {
  if ((props as unknown) === EMPTY_PARA_PROPS) return props;
  const memo = PLAIN.get(props);
  if (memo) return memo as T;
  const record = props as unknown as { [name: string]: unknown };
  let copy: { [name: string]: unknown } | undefined;
  const set = (name: string, value: unknown) => {
    if (!copy) copy = { ...record };
    if (value === undefined) delete copy[name];
    else copy[name] = value;
  };
  for (const name of REVISION_FIELDS) if (record[name] !== undefined) set(name, undefined);
  if (Array.isArray(record.extras)) {
    const extras = plainExtras(record.extras as ExtraXml[], true);
    if (extras !== record.extras) set('extras', extras);
  }
  // The properties of the paragraph mark, and the section a paragraph ends.
  for (const name of ['markProps', 'section']) {
    const inner = record[name];
    if (inner && typeof inner === 'object') {
      const next = plainProps(inner as object);
      // A mark that was nothing but a revision is a mark without properties.
      if (next !== inner) set(name, name === 'markProps' && Object.keys(next).length === 0 ? undefined : next);
    }
  }
  const out = (copy || record) as unknown as T;
  PLAIN.set(props, out);
  // What is plain stays what it is when asked again.
  if (out !== props) PLAIN.set(out, out);
  return out;
}
