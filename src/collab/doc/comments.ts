// Comments of a document that is edited together: the commands of comments as they travel
// (CommentWire in rebase.ts), read from what arrived, run on a document, and made of what the
// view did.
//
// A command of comments is run by the commands of ../../ops/doc/comments.ts, as everywhere else;
// what those decide from the state of the participant they run at is decided here instead, from
// what the document holds, so that every participant - and one who joins later from a saved file -
// makes the same:
//
//   - the number of a new comment: one above the largest number of a comment, or of a mark of
//     one, that the file of the document holds (namesOfDocument; when there is no number above,
//     the least that is free). The commands themselves count on from a counter of the
//     participant (it goes on from numbers that were given and taken back), which two
//     participants do not share;
//   - the paragraph id a comment is known by (w14:paraId; Comment.paraId): picked at random by the
//     author of the command and sent with it, and taken only where the document has no paragraph
//     and no comment of that id (else the command is not applied, by anybody);
//   - who wrote it and when: as its author's side decided them, sent with it. Nobody puts in a
//     name or a clock of their own.
//
// A comment is named by the paragraph id it is known by: that stays the same whatever comes in
// before it (a comment added at the same time by someone else takes the number this one had on
// its author's side). A comment that has none - one of a file that gave it none - is named by its
// number and by who wrote it and when, so that a comment that took its number after it was
// removed is not taken for it.

import { addComment, COMMENT_GONE, commentText, editComment, listComments, removeComment, replyToComment, resolveComment, MAX_COMMENT_TEXT, type CommentEdit } from '../../ops/doc/comments';
import { threadRoots } from '../../formats/docx/comments';
import { cleanText } from '../../ops/doc/pieces';
import { isIsoDate, noteForeignIds } from '../../ops/doc/track';
import type { DocOp } from '../../ops/doc/ops';
import { asRange, samePosition, type DocPosition, type DocRange } from '../../ops/doc/position';
import type { Block, Comment, DocModel, Inline } from '../../model/doc/types';
import { WireError } from '../types';
import { stable, xmlText } from '../wire';
import type { CommentRef, CommentWire, DocWire, Mark, WirePos } from './rebase';
import { hashOf, namesOfDocument, readPos, toPosition, toWirePos } from './wire';

/** Most commands of comments in one step: each reads the whole document for what it names. */
export const MAX_STEP_COMMENTS = 8;
/** Most paragraphs of the text of one comment that travels. */
export const MAX_COMMENT_PARAGRAPHS = 1000;
/**
 * Most characters and paragraphs the comments of a document hold while it is edited together: a
 * command that would go beyond is refused, by everybody alike (the document decides, not who
 * runs it). A file that holds more takes no new text of comments in a room.
 */
export const MAX_ROOM_COMMENT_CHARS = 4000000;
export const MAX_ROOM_COMMENT_PARAGRAPHS = 50000;
/** Longest name of an author, and longest initials. */
const MAX_AUTHOR = 255;
const MAX_INITIALS = 32;
/** The largest number of a comment (ST_DecimalNumber as Word reads it). */
const MAX_COMMENT_ID = 0x7fffffff;

const PARA_ID = /^[0-9A-F]{8}$/;

const fail = (what: string): never => {
  throw new WireError(`Invalid ${what}`);
};

/** A paragraph id a command may give or name: eight upper case hex digits, not 0, below 80000000 (as the schema of Word 2010 says). */
const isParaId = (value: unknown): value is string => typeof value === 'string' && PARA_ID.test(value) && value !== '00000000' && parseInt(value, 16) < 0x80000000;

const isInt = (value: unknown, low: number, high: number): value is number => typeof value === 'number' && Math.floor(value) === value && value >= low && value <= high;

/** A name (of an author, initials) as a file holds it: one line of text. */
function readName(value: unknown, most: number, empty: boolean): string {
  if (typeof value !== 'string' || value.length > most || (!empty && value.length === 0) || !xmlText(value) || /[\t\n\r]/.test(value)) return fail('comment');
  return value;
}

/** The text of a comment: what the commands of comments write, "\n" between paragraphs (cleaned as they clean it: no carriage return). */
function readText(value: unknown): string {
  if (typeof value !== 'string' || value.length > MAX_COMMENT_TEXT || !xmlText(value) || value.indexOf('\r') >= 0 || paragraphsOf(value) > MAX_COMMENT_PARAGRAPHS) return fail('comment');
  return value;
}

/** How many paragraphs a text of a comment makes: one, and one more for every "\n". */
function paragraphsOf(text: string): number {
  let n = 1;
  for (let at = text.indexOf('\n'); at >= 0; at = text.indexOf('\n', at + 1)) n++;
  return n;
}

const digest = (value: unknown): string => (typeof value === 'string' && /^[0-9a-f]{8}$/.test(value) ? value : fail('comment'));

function readRef(value: unknown): CommentRef {
  if (isParaId(value)) return value;
  if (Array.isArray(value) && value.length === 2 && isInt(value[0], 0, MAX_COMMENT_ID) && typeof value[1] === 'string' && /^[0-9a-f]{8}$/.test(value[1])) return [value[0], value[1]];
  return fail('comment');
}

/** A command of comments, read from what arrived. Throws WireError for what is none. */
export function readComment(from: Record<string, unknown>): CommentWire {
  const has = (key: string): boolean => Object.prototype.hasOwnProperty.call(from, key) && from[key] !== undefined;
  const date = (): string => (typeof from.d === 'string' && isIsoDate(from.d) ? from.d : fail('comment'));
  const paraId = (value: unknown): string => (isParaId(value) ? value : fail('comment'));
  switch (from.k) {
    case 'add': {
      const out: Extract<CommentWire, { k: 'add' }> = { t: 'cmt', k: 'add', at: readPos(from.at), x: readText(from.x), u: readName(from.u, MAX_AUTHOR, false), i: readName(from.i, MAX_INITIALS, true), d: date(), q: paraId(from.q) };
      if (has('to')) out.to = readPos(from.to);
      return out;
    }
    case 'reply': {
      return { t: 'cmt', k: 'reply', c: readRef(from.c), x: readText(from.x), u: readName(from.u, MAX_AUTHOR, false), i: readName(from.i, MAX_INITIALS, true), d: date(), q: paraId(from.q) };
    }
    case 'edit': {
      const out: Extract<CommentWire, { k: 'edit' }> = { t: 'cmt', k: 'edit', c: readRef(from.c), x: readText(from.x) };
      if (has('h')) out.h = digest(from.h);
      return out;
    }
    case 'done': {
      const out: Extract<CommentWire, { k: 'done' }> = { t: 'cmt', k: 'done', c: readRef(from.c), v: from.v === 1 ? 1 : from.v === 0 ? 0 : fail('comment') };
      if (has('h')) out.h = digest(from.h);
      return out;
    }
    case 'del': {
      const out: Extract<CommentWire, { k: 'del' }> = { t: 'cmt', k: 'del', c: readRef(from.c) };
      if (has('h')) out.h = digest(from.h);
      return out;
    }
  }
  return fail('comment');
}

/** A command of comments as it goes to the relay. */
export function writeComment(wire: CommentWire): Record<string, unknown> {
  const out: Record<string, unknown> = { ...wire };
  if (wire.k === 'add') {
    out.at = wire.at.b ? [wire.at.s, wire.at.p, wire.at.o, 1] : [wire.at.s, wire.at.p, wire.at.o];
    if (wire.to) out.to = wire.to.b ? [wire.to.s, wire.to.p, wire.to.o, 1] : [wire.to.s, wire.to.p, wire.to.o];
  }
  return out;
}

// ------------------------------------------------------------------ naming

/** Who wrote a comment and when, in eight hex digits: what names a comment that has no paragraph id, beside its number. */
export function commentDigest(comment: Comment): string {
  return hashOf(stable([comment.author === undefined ? '' : comment.author, comment.date === undefined ? '' : comment.date]));
}

/** How a comment of the document is named in a command. */
export function refOf(doc: DocModel, id: number): CommentRef {
  const comment = doc.comments ? doc.comments.items.get(id) : undefined;
  if (!comment) throw new RangeError(COMMENT_GONE);
  return comment.paraId !== undefined ? comment.paraId : [id, commentDigest(comment)];
}

/** The number of the comment a command names. Throws RangeError(COMMENT_GONE) for one the document does not have. */
export function commentNamed(doc: DocModel, ref: CommentRef): number {
  const items = doc.comments ? doc.comments.items : undefined;
  if (items) {
    if (typeof ref === 'string') {
      let found: number | undefined;
      items.forEach((comment, id) => {
        if (found === undefined && comment.paraId === ref) found = id;
      });
      if (found !== undefined) return found;
    } else {
      const comment = items.get(ref[0]);
      if (comment && commentDigest(comment) === ref[1]) return ref[0];
    }
  }
  throw new RangeError(COMMENT_GONE);
}

/**
 * What a command that undoes finds of the comments it is about, in eight hex digits: for a
 * removal, every comment it takes away (a first comment takes its answers along); for an edit,
 * the text of the comment; for a change of resolution, which comments the thread has and which of
 * them are resolved. An undoing that does not find them as its step left them - someone else
 * answered, edited, resolved them since - is not applied (it would take their change away too).
 */
export function commentState(doc: DocModel, id: number, kind: 'edit' | 'done' | 'del'): string {
  const items = doc.comments ? doc.comments.items : undefined;
  const comment = items ? items.get(id) : undefined;
  if (!items || !comment) throw new RangeError(COMMENT_GONE);
  if (kind === 'edit') return hashOf(stable(['edit', commentText(comment)]));
  const roots = threadRoots(items).roots;
  const root = roots.get(comment) as Comment;
  const answer = kind === 'del' && root !== comment;
  const parts: unknown[] = [kind];
  items.forEach((member, key) => {
    if (answer ? member !== comment : roots.get(member) !== root) return;
    parts.push(kind === 'done' ? [member.paraId || key, member.done === true] : [member.paraId || key, commentText(member), member.author || '', member.date || '', member.done === true]);
  });
  return hashOf(stable(parts));
}

/** What the file of a document holds that a new comment must not take: the numbers of comments and of their marks, and the paragraph ids. */
interface Taken {
  numbers: Set<number>;
  most: number;
  paraIds: Set<string>;
}

/** Read off what the document will be written as (see namesOfDocument): the same for everybody who holds the same document, whatever file they read it from. */
function takenIn(doc: DocModel): Taken {
  const out: Taken = { numbers: new Set(), most: -1, paraIds: new Set() };
  const number = (id: number): void => {
    if (id > MAX_COMMENT_ID) return;
    out.numbers.add(id);
    if (id > out.most) out.most = id;
  };
  namesOfDocument(
    doc,
    [],
    (thing) => {
      if (thing.kind === 'para') out.paraIds.add(thing.id);
    },
    undefined,
    undefined,
    (id) => {
      if (/^[0-9]{1,10}$/.test(id)) number(Number(id));
    },
  );
  if (doc.comments)
    doc.comments.items.forEach((comment, id) => {
      number(id);
      if (comment.paraId !== undefined) out.paraIds.add(comment.paraId.toUpperCase());
    });
  return out;
}

/**
 * The commands of comments give a new comment the next number of a counter of this participant,
 * which may lag behind numbers that others gave since: it is moved on beyond every number the
 * document holds, so that the number the command gives is none of a comment or a mark there (and
 * can be renamed alone, see settled()).
 */
function beyond(doc: DocModel, taken: Taken): void {
  if (taken.most >= 0) noteForeignIds(doc, [{ type: 'mark', mark: { type: 'commentRangeStart', id: String(taken.most) } } as unknown as Block]);
}

/**
 * The number a new comment gets: one above the largest the file holds. Where that is the largest a
 * number can be there is none above it, and no comment is added (the commands of comments do not
 * add one there either, in a room or not).
 */
function freeNumber(taken: Taken): number {
  if (taken.most < MAX_COMMENT_ID) return taken.most + 1;
  throw new RangeError('The document has no room for another comment');
}

/**
 * The paragraph id a comment of the file that has none gets when a command gives it one (an answer
 * to its thread, a resolution): made of the comment itself - its number, who wrote it and when -
 * so that two participants who answer it at once give it the same, and what waits behind either
 * names it alike. The first of a row of such ids that the document does not hold (`taken`, `given`).
 */
function derivedParaId(id: number, comment: Comment, taken: Taken, given: Set<string>): string {
  for (let k = 0; k < 1000; k++) {
    const value = parseInt(hashOf(stable(['paraId', id, commentDigest(comment), k])), 16) & 0x7fffffff;
    if (value === 0) continue;
    const text = ('0000000' + value.toString(16).toUpperCase()).slice(-8);
    if (!taken.paraIds.has(text) && !given.has(text)) return text;
  }
  throw new RangeError('The comment cannot be given a paragraph id');
}

/**
 * Refuses text that would make the comments of the document larger than they are taken to be in a
 * room (see MAX_ROOM_COMMENT_CHARS). `old`: the text it replaces, counted as the comments are (an
 * empty comment is one paragraph); none for a new comment.
 */
function withinCommentBudget(doc: DocModel, text: string, old?: string): void {
  let chars = text.length - (old === undefined ? 0 : old.length);
  let paragraphs = paragraphsOf(text) - (old === undefined ? 0 : paragraphsOf(old));
  if (chars <= 0 && paragraphs <= 0) return;
  if (doc.comments)
    doc.comments.items.forEach((comment) => {
      const said = commentText(comment);
      chars += said.length;
      paragraphs += paragraphsOf(said);
    });
  if (chars > MAX_ROOM_COMMENT_CHARS || paragraphs > MAX_ROOM_COMMENT_PARAGRAPHS) throw new RangeError('The comments of the document would be too large to be edited together');
}

// ----------------------------------------------------------------- running

/** A mark of a comment, or a run holding the reference of one, with the number `from` renamed `to`; the same object where it has none. */
function renamedInline(inline: Inline, from: number, to: number): Inline {
  if ((inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') && inline.id === String(from)) return { ...inline, id: String(to) };
  if (inline.type === 'run' && inline.content.some((item) => item.type === 'commentReference' && item.id === from)) {
    return { ...inline, content: inline.content.map((item) => (item.type === 'commentReference' && item.id === from ? { ...item, id: to } : item)) };
  }
  return inline;
}

function renamedBlock(block: Block, from: number, to: number, depth = 0): Block {
  if (block.type === 'table' && depth < 64) {
    // An edit in a cell may be said as the table it is in, made anew.
    let changed = false;
    const rows = block.rows.map((row) => {
      let rowChanged = false;
      const cells = row.cells.map((cell) => {
        let cellChanged = false;
        const blocks = cell.blocks.map((inner) => {
          const next = renamedBlock(inner, from, to, depth + 1);
          if (next !== inner) cellChanged = true;
          return next;
        });
        if (!cellChanged) return cell;
        rowChanged = true;
        return { ...cell, blocks };
      });
      if (!rowChanged) return row;
      changed = true;
      return { ...row, cells };
    });
    return changed ? { ...block, rows } : block;
  }
  if (block.type !== 'paragraph') return block;
  let changed = false;
  const children = block.children.map((inline) => {
    const next = renamedInline(inline, from, to);
    if (next !== inline) changed = true;
    return next;
  });
  return changed ? { ...block, children } : block;
}

/**
 * What a command of comments made, with what the participant it ran at decided taken from the
 * document and the command instead: the new comment (`made`, the number the command gave it)
 * gets `number`, and every comment that got a paragraph id gets the one `give` says (undefined:
 * the command is not applied; `had`: the comment as it was, for one that was there). The ids
 * given are checked: none of them names a paragraph or a comment of the document, nor do two of
 * them name one thing.
 */
function settled(doc: DocModel, edit: CommentEdit, taken: Taken, made: number | undefined, number: number | undefined, give: (comment: Comment, fresh: boolean, given: Set<string>) => string | undefined): CommentEdit {
  const before = doc.comments ? doc.comments.items : new Map<number, Comment>();
  const given = new Set<string>();
  const ops: DocOp[] = edit.ops.map((op): DocOp => {
    if (op.type === 'setComments' && op.items) {
      const items = new Map<number, Comment>();
      op.items.forEach((comment, key) => {
        const fresh = made !== undefined && key === made;
        let out = comment;
        const id = fresh && number !== undefined ? number : key;
        if (id !== key) out = { ...out, id };
        const had = before.get(key);
        if (out.paraId !== undefined && (fresh || !had || had.paraId === undefined)) {
          const paraId = give(had || out, fresh, given);
          if (paraId === undefined || !isParaId(paraId) || taken.paraIds.has(paraId) || given.has(paraId)) throw new RangeError('The comment cannot be known by the paragraph id it was given');
          given.add(paraId);
          out = { ...out, paraId };
        }
        items.set(id, out);
      });
      return { ...op, items };
    }
    if (op.type === 'replaceBlocks' && made !== undefined && number !== undefined && made !== number) return { ...op, insert: op.insert.map((block) => renamedBlock(block, made, number)) };
    return op;
  });
  const out: CommentEdit = { ops, selection: edit.selection };
  if (made !== undefined) out.id = number !== undefined ? number : made;
  return out;
}

/**
 * The operations of a command of comments on a document. Throws RangeError where the document
 * does not take it: the comment is gone, the place takes no comment, a paragraph id is taken.
 */
export function commentEdit(doc: DocModel, wire: CommentWire): CommentEdit {
  switch (wire.k) {
    case 'add': {
      withinCommentBudget(doc, wire.x);
      const taken = takenIn(doc);
      const where: DocPosition | DocRange = wire.to ? { anchor: toPosition(wire.at), focus: toPosition(wire.to) } : toPosition(wire.at);
      const number = freeNumber(taken);
      beyond(doc, taken);
      const edit = addComment(doc, where, wire.x, { author: wire.u, initials: wire.i, date: wire.d });
      if (edit.id === undefined) throw new RangeError('The comment was not made');
      return settled(doc, edit, taken, edit.id, number, (_comment, fresh) => (fresh ? wire.q : undefined));
    }
    case 'reply': {
      const target = commentNamed(doc, wire.c);
      withinCommentBudget(doc, wire.x);
      const taken = takenIn(doc);
      const number = freeNumber(taken);
      beyond(doc, taken);
      const edit = replyToComment(doc, target, wire.x, { author: wire.u, initials: wire.i, date: wire.d });
      if (edit.id === undefined) throw new RangeError('The comment was not made');
      // The first comment of the thread gets the id it is to be known by, when it has none yet: made of itself.
      return settled(doc, edit, taken, edit.id, number, (comment, fresh, given) => (fresh ? wire.q : derivedParaId(comment.id, comment, taken, given)));
    }
    case 'edit': {
      const target = commentNamed(doc, wire.c);
      if (wire.h !== undefined && commentState(doc, target, 'edit') !== wire.h) throw new RangeError('The comment was changed since');
      withinCommentBudget(doc, wire.x, commentText((doc.comments as NonNullable<DocModel['comments']>).items.get(target) as Comment));
      return editComment(doc, target, wire.x);
    }
    case 'done': {
      const target = commentNamed(doc, wire.c);
      if (wire.h !== undefined && commentState(doc, target, 'done') !== wire.h) throw new RangeError('The comment was changed since');
      const edit = resolveComment(doc, target, wire.v === 1);
      // Only where a comment of the thread gets an id is the document read for the ones it has. Each gets one made of itself.
      if (!edit.ops.some((op) => op.type === 'setComments')) return edit;
      const taken = takenIn(doc);
      return settled(doc, edit, taken, undefined, undefined, (comment, _fresh, given) => derivedParaId(comment.id, comment, taken, given));
    }
    case 'del': {
      const target = commentNamed(doc, wire.c);
      if (wire.h !== undefined && commentState(doc, target, 'del') !== wire.h) throw new RangeError('The comment was changed since');
      return removeComment(doc, target);
    }
  }
}

/**
 * What a command of comments did to the places around it, read off its operations: the
 * paragraphs its marks went into or out of are touched (their text is what it was); a mark that
 * stood between blocks and went is a block that went.
 */
export function commentMarks(ops: readonly DocOp[]): Mark[] {
  const out: Mark[] = [];
  for (const op of ops) {
    if (op.type !== 'replaceBlocks') continue;
    if (op.remove === op.insert.length) {
      if (op.remove > 0) out.push({ k: 'tch', s: op.story, c: op.path.slice(), r1: op.at, r2: op.at + op.remove - 1 });
    } else out.push({ k: 'blk', s: op.story, c: op.path.slice(), at: op.at, n: op.remove, m: op.insert.length });
  }
  return out;
}

/** What is known before a command of comments runs, to take it back afterwards. */
export interface CommentPrior {
  /** The comment it names, by its number. */
  id?: number;
  /** For an edit: the text of the comment, and whether it is nothing but paragraphs of text. */
  text?: string;
  plain?: boolean;
  /** For a change of resolution: every comment of the thread was the other way. */
  uniform?: boolean;
}

export function commentPrior(doc: DocModel, wire: CommentWire): CommentPrior {
  if (wire.k === 'add' || wire.k === 'reply') return {};
  let id: number;
  try {
    id = commentNamed(doc, wire.c);
  } catch {
    return {};
  }
  if (wire.k === 'edit') {
    const found = listComments(doc).filter((comment) => comment.id === id)[0];
    return found ? { id, text: found.text, plain: found.plain } : { id };
  }
  if (wire.k === 'done') {
    const items = (doc.comments as NonNullable<DocModel['comments']>).items;
    const roots = threadRoots(items).roots;
    const root = roots.get(items.get(id) as Comment);
    let uniform = true;
    items.forEach((comment) => {
      if (roots.get(comment) === root && (comment.done === true) !== (wire.v === 0)) uniform = false;
    });
    return { id, uniform };
  }
  return { id };
}

/**
 * The commands that take a command of comments back, on the document as it is after it: a
 * comment that was added or an answer is removed; an edit of a comment of plain paragraphs is
 * edited back; a thread that was all resolved (or all open) is so again. Null where it cannot be
 * taken back by commands: a comment that was removed (what it was, and where its marks stood,
 * does not travel), the edit of a comment of more than text, a thread that was neither.
 */
export function commentBack(doc: DocModel, wire: CommentWire, prior: CommentPrior, ops: readonly DocOp[]): DocWire[] | null {
  if (ops.length === 0) return [];
  switch (wire.k) {
    case 'add':
    case 'reply': {
      const made = commentNamed(doc, wire.q);
      return [{ t: 'cmt', k: 'del', c: wire.q, h: commentState(doc, made, 'del') }];
    }
    case 'edit':
      if (prior.id === undefined || prior.text === undefined || prior.plain !== true) return null;
      return [{ t: 'cmt', k: 'edit', c: refOf(doc, prior.id), x: prior.text, h: commentState(doc, prior.id, 'edit') }];
    case 'done':
      if (prior.id === undefined || prior.uniform !== true) return null;
      return [{ t: 'cmt', k: 'done', c: refOf(doc, prior.id), v: wire.v === 1 ? 0 : 1, h: commentState(doc, prior.id, 'done') }];
    case 'del':
      return null;
  }
}

// ------------------------------------------------------------- the view's side

/** A paragraph id nobody holds yet: random (the command is refused by everybody in the rare case that the document has it). */
function randomParaId(): string {
  for (;;) {
    const value = Math.floor(Math.random() * 0x80000000);
    if (value !== 0) return ('0000000' + value.toString(16).toUpperCase()).slice(-8);
  }
}

/** The comments of an edit of comments, and the ones among them that got a paragraph id they did not have. */
function madeBy(doc: DocModel, edit: CommentEdit): { items: Map<number, Comment>; given: number[] } | undefined {
  const op = edit.ops.filter((one) => one.type === 'setComments')[0] as Extract<DocOp, { type: 'setComments' }> | undefined;
  if (!op || !op.items) return undefined;
  const before = doc.comments ? doc.comments.items : new Map<number, Comment>();
  const given: number[] = [];
  op.items.forEach((comment, id) => {
    const had = before.get(id);
    if (comment.paraId !== undefined && (!had || had.paraId === undefined)) given.push(id);
  });
  return { items: op.items, given };
}

const ends = (at: DocPosition | DocRange): { at: WirePos; to?: WirePos } => {
  const range = asRange(at);
  const same = samePosition(range.anchor, range.focus) && (range.anchor.before === true) === (range.focus.before === true);
  return same ? { at: toWirePos(range.anchor) } : { at: toWirePos(range.anchor), to: toWirePos(range.focus) };
};

/**
 * The command of comments a call of the view stands for: `name` is the command of
 * ../../ops/doc/comments.ts that was called, `args` what it was called with after the document,
 * `edit` what it returned there. Undefined where the call cannot travel.
 */
export function commentWireOf(doc: DocModel, name: string, args: unknown[], edit: CommentEdit): CommentWire | undefined {
  switch (name) {
    case 'addComment': {
      const made = madeBy(doc, edit);
      const comment = made && edit.id !== undefined ? made.items.get(edit.id) : undefined;
      if (!comment || comment.author === undefined || comment.date === undefined || typeof args[1] !== 'string') return undefined;
      return { t: 'cmt', k: 'add', ...ends(args[0] as DocPosition | DocRange), x: cleanText(args[1]), u: comment.author, i: comment.initials === undefined ? '' : comment.initials, d: comment.date, q: randomParaId() };
    }
    case 'replyToComment': {
      const made = madeBy(doc, edit);
      const reply = made && edit.id !== undefined ? made.items.get(edit.id) : undefined;
      if (!made || !reply || reply.author === undefined || reply.date === undefined || typeof args[0] !== 'number' || typeof args[1] !== 'string') return undefined;
      return { t: 'cmt', k: 'reply', c: refOf(doc, args[0]), x: cleanText(args[1]), u: reply.author, i: reply.initials === undefined ? '' : reply.initials, d: reply.date, q: randomParaId() };
    }
    case 'editComment':
      return typeof args[0] === 'number' && typeof args[1] === 'string' ? { t: 'cmt', k: 'edit', c: refOf(doc, args[0]), x: cleanText(args[1]) } : undefined;
    case 'resolveComment': {
      if (typeof args[0] !== 'number' || typeof args[1] !== 'boolean') return undefined;
      return { t: 'cmt', k: 'done', c: refOf(doc, args[0]), v: args[1] ? 1 : 0 };
    }
    case 'removeComment':
      return typeof args[0] === 'number' ? { t: 'cmt', k: 'del', c: refOf(doc, args[0]) } : undefined;
  }
  return undefined;
}

/** The commands of ../../ops/doc/comments.ts that travel as commands of comments. */
export const COMMENT_COMMANDS = ['addComment', 'replyToComment', 'editComment', 'resolveComment', 'removeComment'];
