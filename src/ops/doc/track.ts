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
//   - text (and tabs and line breaks) the same author put in as a tracked
//     insertion is taken out for good, as Word does: nobody needs the history
//     of a change to a change. Everything else in such an insertion (a note
//     reference, a field, a picture, an equation) is deleted like anything
//     else, inside the insertion, so that nothing but typed text ever leaves
//     a document that tracks its changes. An editor without a name owns no
//     insertion: the stand-in name is also what Word writes for anybody whose
//     name was removed from a file.
// Content in a `del` wrapper is not content any more: it takes no offsets
// (see position.ts) and no command touches it again.

import { tracksChanges, type Block, type DocModel, type Inline, type ParaProps, type RevisionInfo, type RunItem, type RunProps, type Wrapper } from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { fieldPairs, makeRun, removedPieces, type Piece } from './pieces';

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
  SETUPS.set(doc, { ...setup });
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
}

const systemClock = (): Date => new Date();

/** A date and time as xsd:dateTime writes it, with fields that are in range. */
function isIsoDate(text: string): boolean {
  const match = /^(\d{4})-(\d\d)-(\d\d)T(\d\d):(\d\d)(?::(\d\d)(?:\.\d{1,9})?)?(?:Z|[+-](\d\d):(\d\d))?$/.exec(text);
  if (!match || text.length > 40) return false;
  const [month, day, hour, minute, second] = [Number(match[2]), Number(match[3]), Number(match[4]), Number(match[5]), Number(match[6] || 0)];
  if (month < 1 || month > 12 || day < 1 || day > 31 || hour > 23 || minute > 59 || second > 59) return false;
  return match[7] === undefined || (Number(match[7]) <= 14 && Number(match[8]) <= 59);
}

/** Does a command with these options track its changes in this document. */
export function isTracking(doc: DocModel, options?: TrackOptions): boolean {
  if (options && options.trackChanges === true) return true;
  const setup = SETUPS.get(doc);
  return (setup !== undefined && setup.force === true) || tracksChanges(doc);
}

/** The tracking for a command; undefined when its changes are not tracked. */
export function trackingFor(doc: DocModel, options?: TrackOptions): Tracking | undefined {
  if (!isTracking(doc, options)) return undefined;
  const setup = SETUPS.get(doc) || {};
  const named = options && typeof options.author === 'string' && options.author !== '' ? options.author : setup.author;
  const known = typeof named === 'string' && named !== '';
  const author = known ? (named as string) : DEFAULT_AUTHOR;
  let date: string;
  if (options && options.date !== undefined && options.date !== '') {
    if (typeof options.date !== 'string' || !isIsoDate(options.date)) throw new RangeError('The date of a tracked change must be ISO 8601 text');
    date = options.date;
  } else date = revisionDate((setup.now || systemClock)());
  return { author, named: known, date, revision: () => ({ id: nextRevisionId(doc), author, date }) };
}

// ------------------------------------------------------------------------ ids

const NEXT_IDS = new WeakMap<DocModel, { next: number }>();
/** The largest id an annotation can have (a signed 32-bit number, as the reader takes it). */
const MAX_ID = 0x7fffffff;

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
        // The chain is as deep as the source nesting, which the reader limits.
        for (let w = inline.wrap; w; w = w.parent) if (w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo') note(w.id);
        if (inline.type === 'run') {
          for (const item of inline.content) if (item.type === 'commentReference') note(item.id);
        } else if (inline.type !== 'math' && inline.type !== 'proofError' && inline.type !== 'empty' && inline.type !== 'misc' && inline.type !== 'unknown') note(Number(inline.id));
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) {
        if (row.props.inserted) note(row.props.inserted.id);
        if (row.props.deleted) note(row.props.deleted.id);
        for (const cell of row.cells) idsOfBlocks(cell.blocks, note, depth + 1);
      }
    }
  }
}

/**
 * The largest id the annotations of a document have (revisions, bookmarks,
 * comments, in every story). For a part read from a docx its source text is
 * searched, which also finds the ids in what the model keeps as XML; the
 * model is searched as well, for parts without a source.
 */
function largestId(doc: DocModel): number {
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
  const pattern = /:id="(\d{1,10})"/g;
  for (const source of sources) {
    if (!source) continue;
    pattern.lastIndex = 0;
    for (let match = pattern.exec(source); match; match = pattern.exec(source)) note(Number(match[1]));
  }
  for (const blocks of lists) idsOfBlocks(blocks, note, 0);
  return max;
}

/**
 * An id for a new revision: larger than every id the document had when it
 * was first asked, and than every id given since. An id is never given
 * twice, also not after the change it was made for is undone.
 */
export function nextRevisionId(doc: DocModel): number {
  let state = NEXT_IDS.get(doc);
  if (!state) NEXT_IDS.set(doc, (state = { next: largestId(doc) + 1 }));
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
    if (shared.has(wrap)) out = { ...wrap, id: tracking.revision().id } as Wrapper;
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
    if (around.kind === 'ins' && around.author === tracking.author) return { wrap: base as Wrapper, pieces };
    // Every wrapper between the text and the insertion that is cut is cut with it.
    refuseCopies(base, around);
    const mine = newWrapper('ins', tracking.revision(), around.parent);
    const wrap = swapped(base, around, mine, new Map()) as Wrapper;
    // What follows in the same element is another element now.
    let out = pieces;
    let any = false;
    for (let i = index; i < pieces.length && !any; i++) any = chainHas(pieces[i].wrap, around);
    if (any) {
      const second = { ...around, id: tracking.revision().id } as Wrapper;
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
      if (w.kind === 'ins' && w.author === tracking.author && w.parent === base) return { wrap: w, pieces };
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
 * had inserted as a tracked change, which go. `changed` is false when there
 * was nothing to delete.
 */
export function strikePieces(tracking: Tracking, pieces: Piece[], from: number, to: number): { pieces: Piece[]; changed: boolean } {
  if (from >= to) return { pieces, changed: false };
  const flags = removedPieces(pieces, from, to);
  const author = tracking.author;
  const fields = tracking.named ? fieldPairs(pieces).owner : undefined;
  /** Typed text of this author's own insertion: the only thing that goes for good. */
  const own = (q: number): boolean => {
    const piece = pieces[q];
    const item = piece.item;
    if (!fields || !item || piece.code || fields[q] >= 0 || !inOwnInsertion(piece.wrap, author)) return false;
    return item.type === 'text' || item.type === 'tab' || (item.type === 'break' && item.kind === 'line');
  };
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

/** Was the mark of a paragraph inserted by this author as a tracked change: joining it to the next takes the mark out for good. */
export function isOwnMark(props: ParaProps, tracking: Tracking): boolean {
  const mark = props.markProps;
  return tracking.named && mark !== undefined && mark.inserted !== undefined && mark.inserted.author === tracking.author;
}

/** A kept <w:pPrChange> or <w:rPrChange> with an id of its own. */
function freshChange(change: XmlElement, tracking: Tracking): XmlElement {
  const attrs = new Map(change.attrs);
  attrs.forEach((_value, name) => {
    if (name.slice(name.indexOf(':') + 1) === 'id') attrs.set(name, String(tracking.revision().id));
  });
  return { ...change, attrs };
}

/**
 * Paragraph properties for a second paragraph with the same properties: the
 * formatting changes they keep as XML (of the paragraph and of its mark) get
 * ids of their own, so that no id is in the file twice.
 */
export function withFreshChanges(props: ParaProps, tracking: Tracking): ParaProps {
  const mark = props.markProps;
  if (!props.change && !(mark && mark.change)) return props;
  const out: ParaProps = { ...props };
  if (props.change) out.change = freshChange(props.change, tracking);
  if (mark && mark.change) out.markProps = { ...mark, change: freshChange(mark.change, tracking) };
  return out;
}

export const isDeletedMark = (props: ParaProps): boolean => props.markProps !== undefined && props.markProps.deleted !== undefined;
