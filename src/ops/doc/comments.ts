// Comments of a document: listing them, and adding, editing, answering,
// resolving and removing them as edits that can be undone.
//
// What a comment looks like in the model:
//   - the comment itself is an item of DocModel.comments: who wrote it, when,
//     and its content as blocks. A comment read from a file keeps whatever
//     content it has; one written here is plain paragraphs;
//   - the text it is about is marked in its story: a <w:commentRangeStart>
//     and a <w:commentRangeEnd> (inlines of paragraphs, or blocks between
//     them) and, after the end, a run with a <w:commentReference>. The
//     editing commands leave all three where they are: text that is deleted
//     takes none of them along, so a comment whose text is gone stays, about
//     a point (UNSURE: Word removes the comment with the last of its text);
//   - an answer is a comment of its own with the marks of the comment it
//     answers repeated beside them, and says which comment it answers
//     (Comment.parent, always the first comment of the thread, as Word
//     writes it); resolving sets Comment.done on every comment of a thread
//     (UNSURE: Word files seen state it on each comment of the thread).
//
// None of this is ever a tracked change: Word does not track comments, and
// neither do these commands, whatever the document says about tracking
// (UNSURE for removal: Word removes a comment outright also while tracking).
// A comment is not copied either: a copy carries text, never the marks of a
// comment (see fragment.ts).

import { paraIdsInSource, threadRoots } from '../../formats/docx/comments';
import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  findWrapper,
  paragraphText,
  partNamespaces,
  setPartNamespaces,
  type Block,
  type Comment,
  type DocModel,
  type DocStyle,
  type DocStyles,
  type ExtraXml,
  type Inline,
  type Paragraph,
  type ParaProps,
  type RangeMark,
  type Run,
  type RunItem,
  type RunProps,
  type Wrapper,
} from '../../model/doc/types';
import { Draft, type DocEdit } from './draft';
import { boundary, cleanText, flatten, itemPiece, paragraphWith, rebuild, textPieces, wrapBetween, type Piece } from './pieces';
import {
  asRange,
  collapsed,
  comparePositions,
  inTextBox,
  isCollapsed,
  itemsLength,
  orderRange,
  paragraphIn,
  resolvePosition,
  storyIds,
  storyOf,
  storyStart,
  wordAt,
  type DocPosition,
  type DocRange,
} from './position';
import { replacedParagraphRanges } from './find-edit';
import { ownerOf } from './stand-in';
import { nextRevisionId, trackingFor } from './track';

/** Thrown for a comment the document does not have (any more). */
export const COMMENT_GONE = 'The document has no such comment';
/** Thrown where no comment can be put: a header or footer, a text box, content that is kept whole. */
export const COMMENT_NOT_HERE = 'A comment cannot be put here';
/** Thrown for a comment that cannot be changed: its marks sit in content that is kept whole, or it has no paragraph to be known by. */
export const COMMENT_FIXED = 'This comment cannot be changed';
/** Longest text of a comment, UTF-16 code units. */
export const MAX_COMMENT_TEXT = 32767;
/** Most comments one document takes through these commands. */
export const MAX_COMMENTS = 100000;

export interface CommentOptions {
  /** Who writes the comment. Default: the author the document is edited as (setTracking), else "Author". */
  author?: string;
  /** The initials of the author. Default: the first letters of the words of the name. */
  initials?: string;
  /** When, as ISO 8601 text. Default: now, by the clock of setTracking(). */
  date?: string;
}

/** A comment as the document has it now. */
export interface DocComment {
  id: number;
  /** "" for a comment that names nobody. */
  author: string;
  initials?: string;
  /** ISO 8601 text as the file has it. */
  date?: string;
  /** The text of the comment, its paragraphs separated by "\n". */
  text: string;
  /** The comment it answers: the first comment of its thread. */
  parent?: number;
  /** The thread it belongs to is resolved. */
  done: boolean;
  /** The story of the text it is about; undefined for a comment nothing refers to. */
  story?: string;
  /** Where that text begins and ends; the same place for a comment about a point. */
  start?: DocPosition;
  end?: DocPosition;
  /** The content is plain paragraphs: editing it loses nothing but the text that is replaced. */
  plain: boolean;
  /** Its marks can be taken out of the document and others put beside them: it can be answered and removed. */
  canChange: boolean;
}

/** What an edit of comments hands back besides the edit. */
export interface CommentEdit extends DocEdit {
  /** The comment that was added. */
  id?: number;
}

const bad = (message: string): never => {
  throw new RangeError(message);
};

// ------------------------------------------------------------------ finding

/** A mark of a comment in a story. */
interface Place {
  story: string;
  /** Path of the list of blocks. */
  container: number[];
  /** Index of the block: the paragraph the mark is in, or the mark itself when it stands between blocks. */
  index: number;
  /** Index among the inlines of the paragraph; -1 for a mark between blocks. */
  inline: number;
  /** For a mark in a paragraph: the offset in the text of the paragraph it stands at. */
  offset: number;
}

interface Anchors {
  starts: Place[];
  ends: Place[];
  refs: Place[];
  /** A mark sits where it cannot be taken out: in a text box, in kept XML, in content that is written back whole. */
  fixed: boolean;
  /** When the first of the marks was met, in reading order. */
  order: number;
}

const MAX_DEPTH = 256;
const COMMENT_ID = /^\d{1,10}$/;

const markId = (id: string): number | undefined => (COMMENT_ID.test(id) ? Number(id) : undefined);

function refsOfItems(items: RunItem[], note: (id: number, fixed: boolean) => void, top: boolean, depth: number): void {
  for (const item of items) {
    if (item.type === 'commentReference') note(item.id, !top);
    else if (depth < 64) {
      if (item.type === 'alternate') refsOfItems(item.items, note, false, depth + 1);
      else if ((item.type === 'drawing' || item.type === 'pict' || item.type === 'object') && item.textBoxes) {
        for (const box of item.textBoxes) fixedMarks(box, note, depth + 1);
      }
    }
  }
}

/** The marks of comments in content that is never edited (a text box): they only say that their comments are fixed. */
function fixedMarks(blocks: Block[], note: (id: number, fixed: boolean) => void, depth: number): void {
  if (depth > MAX_DEPTH) return;
  for (const block of blocks) {
    if (block.type === 'mark') {
      const id = block.mark.type === 'commentRangeStart' || block.mark.type === 'commentRangeEnd' ? markId(block.mark.id) : undefined;
      if (id !== undefined) note(id, true);
    } else if (block.type === 'paragraph') {
      for (const inline of block.children) {
        if (inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') {
          const id = markId(inline.id);
          if (id !== undefined) note(id, true);
        } else if (inline.type === 'run') refsOfItems(inline.content, (id) => note(id, true), false, depth);
      }
    } else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) fixedMarks(cell.blocks, note, depth + 1);
    }
  }
}

function looseIds(extras: ExtraXml[] | undefined, note: (id: number, fixed: boolean) => void): void {
  if (!extras) return;
  for (const extra of extras) {
    const name = extra.el.name;
    const local = name.slice(name.indexOf(':') + 1);
    if (local !== 'commentRangeStart' && local !== 'commentRangeEnd') continue;
    extra.el.attrs.forEach((value, key) => {
      if (key.slice(key.indexOf(':') + 1) === 'id' && COMMENT_ID.test(value)) note(Number(value), true);
    });
  }
}

const SCANS = new WeakMap<DocModel, { blocks: Map<string, Block[]>; value: Map<number, Anchors> }>();

/** Where the marks of every comment are. */
function scanAnchors(doc: DocModel): Map<number, Anchors> {
  const out = new Map<number, Anchors>();
  let order = 0;
  const of = (id: number): Anchors => {
    let anchors = out.get(id);
    if (!anchors) out.set(id, (anchors = { starts: [], ends: [], refs: [], fixed: false, order: order++ }));
    return anchors;
  };
  const fixed = (id: number): void => {
    of(id).fixed = true;
  };
  const walk = (story: string, blocks: Block[], container: number[], depth: number): void => {
    if (depth > MAX_DEPTH) return;
    for (let index = 0; index < blocks.length; index++) {
      const block = blocks[index];
      if (block.type === 'mark') {
        const mark = block.mark;
        if (mark.type !== 'commentRangeStart' && mark.type !== 'commentRangeEnd') continue;
        const id = markId(mark.id);
        if (id === undefined) continue;
        const anchors = of(id);
        if (block.wrap && findWrapper(block, 'alternate')) anchors.fixed = true;
        else (mark.type === 'commentRangeStart' ? anchors.starts : anchors.ends).push({ story, container, index, inline: -1, offset: 0 });
      } else if (block.type === 'paragraph') {
        const whole = block.wrap !== undefined && findWrapper(block, 'alternate') !== undefined;
        // The offset each inline begins at, counted as paragraphText() counts: a paragraph may hold very many marks.
        let offset = 0;
        for (let i = 0; i < block.children.length; i++) {
          const inline = block.children[i];
          const at = offset;
          if ((inline.type === 'run' || inline.type === 'math') && !(inline.wrap !== undefined && (findWrapper(inline, 'del') !== undefined || findWrapper(inline, 'moveFrom') !== undefined))) {
            if (inline.type === 'math') offset += inline.text.length;
            else if (inline.fieldCode !== true) offset += itemsLength(inline.content, inline.fieldCode, 0);
          }
          const frozen = whole || (inline.wrap !== undefined && findWrapper(inline, 'alternate') !== undefined);
          if (inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') {
            const id = markId(inline.id);
            if (id === undefined) continue;
            const anchors = of(id);
            if (frozen) anchors.fixed = true;
            else (inline.type === 'commentRangeStart' ? anchors.starts : anchors.ends).push({ story, container, index, inline: i, offset: at });
          } else if (inline.type === 'run') {
            refsOfItems(
              inline.content,
              (id, inner) => {
                const anchors = of(id);
                if (frozen || inner) anchors.fixed = true;
                else anchors.refs.push({ story, container, index, inline: i, offset: at });
              },
              true,
              depth,
            );
          }
        }
      } else if (block.type === 'table') {
        looseIds(block.extras, fixed);
        for (let r = 0; r < block.rows.length; r++) {
          const row = block.rows[r];
          looseIds(row.extras, fixed);
          for (let c = 0; c < row.cells.length; c++) walk(story, row.cells[c].blocks, container.concat([index, r, c]), depth + 1);
        }
      }
    }
  };
  for (const story of storyIds(doc)) walk(story, storyOf(doc, story).blocks, [], 0);
  return out;
}

/**
 * The marks of the comments, found once for each state of the document: the
 * lists of blocks of the stories are what an edit replaces or changes, so
 * the same lists with the same blocks at their ends and of the same length
 * are taken for the same document only together with every block being the
 * same object, which is checked.
 */
function anchorsOf(doc: DocModel): Map<number, Anchors> {
  const cached = SCANS.get(doc);
  const ids = storyIds(doc);
  if (cached && cached.blocks.size === ids.length) {
    let same = true;
    for (const id of ids) {
      const blocks = storyOf(doc, id).blocks;
      const kept = cached.blocks.get(id);
      if (!kept || kept.length !== blocks.length) {
        same = false;
        break;
      }
      for (let i = 0; i < blocks.length && same; i++) if (kept[i] !== blocks[i]) same = false;
      if (!same) break;
    }
    if (same) return cached.value;
  }
  const value = scanAnchors(doc);
  const blocks = new Map<string, Block[]>();
  for (const id of ids) blocks.set(id, storyOf(doc, id).blocks.slice());
  SCANS.set(doc, { blocks, value });
  return value;
}

/** The position a place stands for: where the text a start mark opens begins (`edge` 1), or where the text an end mark or a reference closes ends (-1). */
function positionOf(doc: DocModel, place: Place, edge: 1 | -1): DocPosition | undefined {
  let blocks: Block[];
  try {
    blocks = place.container.length ? resolveContainer(doc, place.story, place.container) : storyOf(doc, place.story).blocks;
  } catch {
    return undefined;
  }
  const block = blocks[place.index];
  if (!block) return undefined;
  if (place.inline >= 0) {
    if (block.type !== 'paragraph') return undefined;
    return { story: place.story, path: place.container.concat([place.index]), offset: place.offset };
  }
  // A mark between blocks: the start of the paragraph after it, or the end of the one before it.
  const after = paragraphIn(blocks, place.container, place.index + 1, 1);
  const before = paragraphIn(blocks, place.container, place.index - 1, -1);
  const path = edge === 1 ? after || before : before || after;
  if (!path) return undefined;
  const atEnd = path === before;
  let length = 0;
  if (atEnd) {
    try {
      length = resolvePosition(doc, { story: place.story, path, offset: 0 }).length;
    } catch {
      return undefined;
    }
  }
  return { story: place.story, path, offset: length };
}

function resolveContainer(doc: DocModel, story: string, container: number[]): Block[] {
  let blocks = storyOf(doc, story).blocks;
  for (let i = 0; i + 2 < container.length; i += 3) {
    const table = blocks[container[i]];
    const row = table && table.type === 'table' ? table.rows[container[i + 1]] : undefined;
    const cell = row ? row.cells[container[i + 2]] : undefined;
    if (!cell) return bad(COMMENT_GONE);
    blocks = cell.blocks;
  }
  return blocks;
}

// ---------------------------------------------------------------- the threads

interface Threads {
  /** The first comment of the thread of each comment. */
  roots: Map<Comment, Comment>;
  /** The answers of each first comment, in the order of the comments. */
  answers: Map<Comment, Comment[]>;
}

// A table of comments is never changed in place (see the operation "setComments"), so its threads can be kept.
const THREADS = new WeakMap<Map<number, Comment>, Threads>();

/** The threads of a table of comments: answers of any depth belong to the first comment they lead to (see threadRoots). */
function threadsOf(items: Map<number, Comment>): Threads {
  let threads = THREADS.get(items);
  if (!threads) {
    const roots = threadRoots(items).roots;
    const answers = new Map<Comment, Comment[]>();
    items.forEach((comment) => {
      const root = roots.get(comment) as Comment;
      if (root === comment) return;
      const thread = answers.get(root);
      if (thread) thread.push(comment);
      else answers.set(root, [comment]);
    });
    THREADS.set(items, (threads = { roots, answers }));
  }
  return threads;
}

/** The first comment of the thread a comment belongs to; the comment itself when it answers none (or one that is gone, or is part of a ring of answers). */
function rootOf(items: Map<number, Comment>, comment: Comment): Comment {
  return threadsOf(items).roots.get(comment) || comment;
}

/** A comment and its answers, in the order of the comments. */
function threadOf(items: Map<number, Comment>, root: Comment): Comment[] {
  const answers = threadsOf(items).answers.get(root);
  return answers ? [root].concat(answers) : [root];
}

const PLAIN_ITEMS = new Set<string>(['text', 'tab', 'break', 'noBreakHyphen', 'softHyphen', 'annotationRef', 'lastRenderedPageBreak']);

function blocksText(blocks: Block[], out: string[], depth: number): void {
  if (depth > MAX_DEPTH) return;
  for (const block of blocks) {
    if (block.type === 'paragraph') out.push(paragraphText(block));
    else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) blocksText(cell.blocks, out, depth + 1);
  }
}

// A comment is never changed in place, so what is read out of it can be kept: a list of the comments is asked
// for after every change of the document, and most comments are the objects they were.
const TEXTS = new WeakMap<object, string>();
const PLAIN = new WeakMap<object, boolean>();

/** The text of a comment: its paragraphs (those in tables too) separated by "\n". */
export function commentText(comment: { blocks: Block[] }): string {
  let text = TEXTS.get(comment);
  if (text === undefined) {
    const out: string[] = [];
    blocksText(comment.blocks, out, 0);
    TEXTS.set(comment, (text = out.join('\n')));
  }
  return text;
}

/** Is the content of a comment nothing but paragraphs of text: what editComment() writes, and replaces without a loss. */
function isPlain(comment: Comment): boolean {
  let plain = PLAIN.get(comment);
  if (plain === undefined) PLAIN.set(comment, (plain = hasPlainContent(comment)));
  return plain;
}

function hasPlainContent(comment: Comment): boolean {
  for (const block of comment.blocks) {
    if (block.type !== 'paragraph' || block.wrap) return false;
    for (const inline of block.children) {
      if (inline.type === 'proofError') continue;
      if (inline.type !== 'run' || inline.wrap || inline.fieldCode !== undefined) return false;
      for (const item of inline.content) if (!PLAIN_ITEMS.has(item.type)) return false;
    }
  }
  return true;
}

/**
 * The comments of the document as it is now: those about text in the order
 * of that text (the body, then the notes), each followed by its answers in
 * the order they were written, then the ones nothing refers to. After an
 * edit, ask again.
 */
export function listComments(doc: DocModel): DocComment[] {
  const collection = doc.comments;
  if (!collection || collection.items.size === 0) return [];
  const items = collection.items;
  const anchors = anchorsOf(doc);
  const describe = (comment: Comment): DocComment => {
    const root = rootOf(items, comment);
    const out: DocComment = { id: comment.id, author: comment.author === undefined ? '' : comment.author, text: commentText(comment), done: root.done === true || comment.done === true, plain: isPlain(comment), canChange: true };
    if (comment.initials !== undefined) out.initials = comment.initials;
    if (comment.date !== undefined) out.date = comment.date;
    if (root !== comment) out.parent = root.id;
    const at = anchors.get(comment.id);
    if (at) {
      if (at.fixed) out.canChange = false;
      const first = at.starts[0] || at.refs[0] || at.ends[0];
      const last = at.ends[0] || at.refs[0] || at.starts[0];
      const start = first ? positionOf(doc, first, 1) : undefined;
      let end = last ? positionOf(doc, last, -1) : undefined;
      if (start && end) {
        // Marks the wrong way round, or in different stories, are a comment about the place of the first.
        try {
          if (start.story !== end.story || comparePositions(start, end) > 0) end = start;
        } catch {
          end = start;
        }
        out.story = start.story;
        out.start = start;
        out.end = end;
      }
    }
    return out;
  };
  // Every comment once, with the thread it belongs to: a file may hold very many of them.
  const roots: { comment: Comment; order: number }[] = [];
  const answers = threadsOf(items).answers;
  let unplaced = anchors.size;
  items.forEach((comment) => {
    if (rootOf(items, comment) !== comment) return;
    const at = anchors.get(comment.id);
    roots.push({ comment, order: at ? at.order : unplaced++ });
  });
  roots.sort((a, b) => a.order - b.order);
  const out: DocComment[] = [];
  for (const entry of roots) {
    out.push(describe(entry.comment));
    const thread = answers.get(entry.comment);
    if (thread) for (const comment of thread) out.push(describe(comment));
  }
  return out;
}

// ------------------------------------------------------------------- content

/** The id of the style Word names so, when the document has it: by that name, else by the id Word gives it. */
function styleNamed(doc: DocModel, type: 'paragraph' | 'character', name: string, id: string): string | undefined {
  let found: string | undefined;
  doc.styles.styles.forEach((style, key) => {
    if (found === undefined && style.type === type && typeof style.name === 'string' && style.name.toLowerCase() === name) found = key;
  });
  if (found !== undefined) return found;
  const style = doc.styles.styles.get(id);
  return style && style.type === type ? id : undefined;
}

const STYLED = new WeakMap<object, Map<string, object>>();

/** Properties that state a style and nothing else; the same object for the same style of the same styles. */
function styleProps<T extends ParaProps | RunProps>(doc: DocModel, style: string | undefined, empty: T): T {
  if (style === undefined) return empty;
  let known = STYLED.get(doc.styles);
  if (!known) STYLED.set(doc.styles, (known = new Map()));
  const key = (empty === (EMPTY_PARA_PROPS as unknown) ? 'p:' : 'r:') + style;
  let props = known.get(key);
  if (!props) known.set(key, (props = Object.freeze({ style })));
  return props as T;
}

/** An id no style of the document has: the one Word uses, or that with the least number that is free. */
function freeStyleId(styles: Map<string, DocStyle>, wanted: string): string {
  let id = wanted;
  for (let n = 1; styles.has(id); n++) id = wanted + n;
  return id;
}

/**
 * Brings the two styles Word gives a comment, each only when the document has no style Word
 * would take for it (see styleNamed): the mark of a comment ("annotation reference", 8 pt) and
 * the text of a comment ("annotation text", 10 pt), based on the default styles of their kinds
 * when the document has those. Styles the document has are not touched, and an id that is
 * taken (by a style of another kind) is left to its style. Without them a comment is shown in
 * the size of the text, its mark as large as a letter.
 *
 * UNSURE: this is the least Word needs, not all it writes for a first comment: its "annotation
 * text" is linked to a character style ("Comment Text Char") and sets the line spacing, and
 * the styles of the subject of a comment and of the text of a balloon come with it.
 */
function ensureCommentStyles(doc: DocModel, draft: Draft): void {
  const had = doc.styles.styles;
  const reference = styleNamed(doc, 'character', 'annotation reference', 'CommentReference');
  const text = styleNamed(doc, 'paragraph', 'annotation text', 'CommentText');
  if (reference !== undefined && text !== undefined) return;
  const styles = new Map(had);
  const base = (type: 'paragraph' | 'character'): string | undefined => {
    const id = doc.styles.defaultIds[type];
    const style = id === undefined ? undefined : had.get(id);
    return style && style.type === type ? id : undefined;
  };
  const make = (id: string, type: 'paragraph' | 'character', name: string, size: number): void => {
    const style: DocStyle = { id, type, name, isDefault: false, uiPriority: 99, semiHidden: true, unhideWhenUsed: true, run: { size, sizeCs: size } };
    const basedOn = base(type);
    if (basedOn !== undefined) style.basedOn = basedOn;
    styles.set(id, style);
  };
  if (reference === undefined) make(freeStyleId(styles, 'CommentReference'), 'character', 'annotation reference', 16);
  if (text === undefined) make(freeStyleId(styles, 'CommentText'), 'paragraph', 'annotation text', 20);
  const next: DocStyles = { ...doc.styles, styles };
  // The new table stands for the same part.
  const spelt = partNamespaces(doc.styles);
  if (spelt) setPartNamespaces(next, spelt);
  draft.apply({ type: 'setStyles', styles: next });
}

const referenceProps = (doc: DocModel): RunProps => styleProps<RunProps>(doc, styleNamed(doc, 'character', 'annotation reference', 'CommentReference'), EMPTY_RUN_PROPS);

/** The content of a comment that says `text`: a paragraph for each line, the first with the mark of the comment before it. */
function contentOf(doc: DocModel, text: string): Block[] {
  const props = styleProps<ParaProps>(doc, styleNamed(doc, 'paragraph', 'annotation text', 'CommentText'), EMPTY_PARA_PROPS);
  const lines = text.split('\n');
  const out: Block[] = [];
  for (let i = 0; i < lines.length; i++) {
    const children: Inline[] = [];
    if (i === 0) children.push({ type: 'run', props: referenceProps(doc), content: [{ type: 'annotationRef' }] });
    const pieces = textPieces(lines[i], EMPTY_RUN_PROPS, undefined, {});
    if (pieces.length) for (const inline of rebuild(pieces)) children.push(inline);
    out.push({ type: 'paragraph', props, children });
  }
  return out;
}

function checkedText(text: unknown): string {
  if (typeof text !== 'string') return bad('The text of a comment must be a string');
  const clean = cleanText(text);
  if (clean.length > MAX_COMMENT_TEXT) return bad('The text of a comment is too long');
  return clean;
}

/** The initials of a name: the first letter of each of its words, at most nine. */
export function initialsOf(author: string): string {
  let out = '';
  for (const word of author.split(/[\s.\-_]+/)) {
    if (word === '' || out.length >= 9) continue;
    const code = word.codePointAt(0) as number;
    out += String.fromCodePoint(code).toUpperCase();
  }
  return out;
}

const PARA_IDS = new WeakMap<DocModel, Set<string>>();

/** A paragraph id no paragraph of the document has and none was given before, also not to a comment that was taken back. */
function freshParaId(doc: DocModel, seed: number): string {
  let taken = PARA_IDS.get(ownerOf(doc));
  if (!taken) {
    // Of every part that holds paragraphs: the body, the comments, the headers and footers, the notes. Paragraphs
    // that an edit made carry no id, so the texts of the parts as they were read say all there is.
    const ids = (taken = paraIdsInSource(doc.body.source));
    const of = (story: { source: string } | undefined): void => {
      if (story) paraIdsInSource(story.source).forEach((id) => ids.add(id));
    };
    of(doc.comments);
    doc.headers.forEach(of);
    doc.footers.forEach(of);
    of(doc.footnotes);
    of(doc.endnotes);
    PARA_IDS.set(ownerOf(doc), taken);
  }
  if (doc.comments) doc.comments.items.forEach((comment) => (comment.paraId !== undefined ? (taken as Set<string>).add(comment.paraId) : 0));
  let value = (Math.imul(seed | 0, 0x9e3779b1) ^ 0x2545f491) >>> 0;
  for (let tries = 0; tries < 1000000; tries++) {
    value = (Math.imul(value, 1664525) + 1013904223) >>> 0;
    const candidate = value & 0x7fffffff;
    if (candidate === 0) continue;
    let text = candidate.toString(16).toUpperCase();
    while (text.length < 8) text = '0' + text;
    if (taken.has(text)) continue;
    taken.add(text);
    return text;
  }
  return bad('The document has no room for another comment');
}

/** A changed comment: a new object, without the place it had in the source of its part. */
function changed(comment: Comment, change: Partial<Comment>): Comment {
  const out: Comment = { ...comment, ...change };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

function hasLastParagraph(comment: Comment): boolean {
  for (let i = comment.blocks.length - 1; i >= 0; i--) {
    const type = comment.blocks[i].type;
    if (type === 'paragraph') return true;
    if (type === 'table') return false;
  }
  return false;
}

// ------------------------------------------------------------------ building

/** Runs a command on a draft that tracks nothing, and hands back its operations; the document is as before when this returns or throws. */
function build(doc: DocModel, command: (draft: Draft) => { selection: DocRange; id?: number }): CommentEdit {
  const draft = new Draft(doc);
  try {
    const result = command(draft);
    draft.finish();
    const edit: CommentEdit = { ops: draft.ops, selection: result.selection };
    if (result.id !== undefined) edit.id = result.id;
    return edit;
  } finally {
    draft.rollback();
  }
}

/** The selection an edit of comments leaves: the one given when it still names places of the document, else the start of the body. */
function kept(doc: DocModel, selection: DocRange | undefined): DocRange {
  if (selection) {
    try {
      resolvePosition(doc, selection.anchor);
      resolvePosition(doc, selection.focus);
      return { anchor: { ...selection.anchor, path: selection.anchor.path.slice() }, focus: { ...selection.focus, path: selection.focus.path.slice() } };
    } catch {
      // The place is gone (a mark between blocks was taken out before it).
    }
  }
  const start = storyStart(doc, 'body');
  return collapsed(start || { story: 'body', path: [0], offset: 0 });
}

const markPiece = (mark: RangeMark): Piece => ({ inline: mark, props: EMPTY_RUN_PROPS, wrap: mark.wrap, group: mark, index: -1, code: false, len: 0 });

function makeMark(type: 'commentRangeStart' | 'commentRangeEnd', id: number, wrap: Wrapper | undefined): RangeMark {
  const mark: RangeMark = { type, id: String(id) };
  if (wrap) mark.wrap = wrap;
  return mark;
}

const wrapAt = (pieces: Piece[], index: number): Wrapper | undefined => wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);

function commentsOf(doc: DocModel): Map<number, Comment> {
  return doc.comments ? doc.comments.items : new Map<number, Comment>();
}

function commentAt(doc: DocModel, id: unknown): Comment {
  const comment = typeof id === 'number' && doc.comments ? doc.comments.items.get(id) : undefined;
  return comment || bad(COMMENT_GONE);
}

/**
 * The range a comment asked for at a caret is about: the word the caret is
 * in or at the end of, as Word takes it; the caret itself between words and
 * in an empty paragraph. UNSURE: Word's own rule for a caret between words
 * (it takes the word after it) was not compared.
 */
function wordRange(doc: DocModel, position: DocPosition): DocRange {
  const at = resolvePosition(doc, position);
  const text = paragraphText(at.paragraph);
  const isWord = (span: [number, number]): boolean => span[1] > span[0] && !/^\s*$/.test(text.slice(span[0], span[1]));
  let span = wordAt(text, at.offset);
  if (!isWord(span) || at.offset >= text.length) {
    const before = at.offset > 0 ? wordAt(text, at.offset - 1) : span;
    span = isWord(before) && before[1] === at.offset ? before : isWord(span) && at.offset < text.length ? span : [at.offset, at.offset];
  }
  if (span[0] === span[1]) return collapsed({ story: position.story, path: position.path.slice(), offset: at.offset });
  return { anchor: { story: position.story, path: position.path.slice(), offset: span[0] }, focus: { story: position.story, path: position.path.slice(), offset: span[1] } };
}

/**
 * Adds a comment about a range of text, or at a caret (then about the word
 * the caret is in; between words, about the place itself). The comment gets
 * a new id, the author and the date of `options` (default: who the document
 * is edited as, now), and `text` as plain paragraphs ("\n" separates them).
 * Throws RangeError(COMMENT_NOT_HERE) for a header, a footer and a text box.
 */
export function addComment(doc: DocModel, at: DocPosition | DocRange, text: string, options?: CommentOptions): CommentEdit {
  const clean = checkedText(text);
  const given = asRange(at);
  const story = given.focus.story;
  if (given.anchor.story !== story) bad('The positions are in different stories');
  if (story !== 'body' && story.slice(0, 9) !== 'footnote:' && story.slice(0, 8) !== 'endnote:') bad(COMMENT_NOT_HERE);
  if (inTextBox(given.anchor.path) || inTextBox(given.focus.path)) bad(COMMENT_NOT_HERE);
  resolvePosition(doc, given.anchor);
  resolvePosition(doc, given.focus);
  const range = isCollapsed(given) ? wordRange(doc, given.focus) : given;
  const { start, end } = orderRange(range);
  if (commentsOf(doc).size >= MAX_COMMENTS) bad('The document has no room for another comment');
  const who = trackingFor(doc, { trackChanges: true, author: options && options.author, date: options && options.date });
  if (!who) return bad(COMMENT_NOT_HERE);
  return build(doc, (draft) => {
    ensureCommentStyles(doc, draft);
    const id = nextRevisionId(doc);
    const comment: Comment = { id, author: who.author, date: who.date, blocks: contentOf(doc, clean), paraId: freshParaId(doc, id) };
    const initials = options && typeof options.initials === 'string' ? options.initials : initialsOf(who.author);
    if (initials !== '') comment.initials = initials;
    const items = new Map(commentsOf(doc));
    items.set(id, comment);
    draft.apply({ type: 'setComments', items });

    // The end first: putting it in moves nothing that stands before it.
    const reference = referenceProps(doc);
    const put = (position: DocPosition, edge: 'start' | 'end'): void => {
      const paragraph = resolvePosition(doc, position).paragraph;
      const pieces = flatten(paragraph);
      // A range begins before the character at its start and ends after the one before its end: what has no width there stays outside.
      const index = boundary(pieces, position.offset, edge === 'start' ? 'before' : 'after');
      const wrap = wrapAt(pieces, index);
      if (edge === 'start') pieces.splice(index, 0, markPiece(makeMark('commentRangeStart', id, wrap)));
      else pieces.splice(index, 0, markPiece(makeMark('commentRangeEnd', id, wrap)), itemPiece({ type: 'commentReference', id }, reference, wrap));
      draft.setParagraph(position.story, position.path, paragraphWith(paragraph, rebuild(pieces)));
    };
    put(end, 'end');
    if (comparePositions(start, end) === 0) {
      // About a point: the start goes right before the end that was just put there.
      const paragraph = resolvePosition(doc, end).paragraph;
      const children = paragraph.children.slice();
      let index = -1;
      for (let i = 0; i < children.length; i++) if (children[i].type === 'commentRangeEnd' && (children[i] as RangeMark).id === String(id)) index = i;
      if (index < 0) return bad(COMMENT_NOT_HERE);
      children.splice(index, 0, makeMark('commentRangeStart', id, children[index].wrap));
      draft.setParagraph(end.story, end.path, paragraphWith(paragraph, children));
    } else put(start, 'start');
    return { selection: kept(doc, given), id };
  });
}

/** Replaces the text of a comment: plain paragraphs, "\n" between them. Who wrote it and when stay (UNSURE: Word keeps both too, as far as its files show). No operations when the comment says that already. */
export function editComment(doc: DocModel, id: number, text: string, selection?: DocRange): CommentEdit {
  const clean = checkedText(text);
  const comment = commentAt(doc, id);
  return build(doc, (draft) => {
    if (!(isPlain(comment) && commentText(comment) === clean)) {
      const items = new Map(commentsOf(doc));
      items.set(comment.id, changed(comment, { blocks: contentOf(doc, clean) }));
      draft.apply({ type: 'setComments', items });
    }
    return { selection: kept(doc, selection) };
  });
}

// ------------------------------------------------------- text of a comment

/**
 * The story Find says a paragraph of the text of a comment is in: "comment:" and the id of the
 * comment. Its paths are those of the blocks of the comment (Comment.blocks), as in any other
 * story. No other command takes such a position.
 */
export const COMMENT_STORY = 'comment:';

export const commentStory = (id: number): string => COMMENT_STORY + id;

/** The id of the comment a story of Find names; undefined for any other story. */
export function commentOfStory(story: unknown): number | undefined {
  if (typeof story !== 'string' || story.slice(0, COMMENT_STORY.length) !== COMMENT_STORY) return undefined;
  const id = story.slice(COMMENT_STORY.length);
  return /^-?\d{1,10}$/.test(id) ? Number(id) : undefined;
}

/** The paragraph of the blocks of a comment at a path ([block] or, in a table, [table, row, cell, ...]); undefined for none. */
function paragraphOfBlocks(blocks: Block[], path: readonly number[]): Paragraph | undefined {
  if (!Array.isArray(path) || path.length % 3 !== 1 || path.length > MAX_DEPTH * 3) return undefined;
  let list = blocks;
  for (let i = 0; i + 1 < path.length; i += 3) {
    const table = list[path[i]];
    const row = table && table.type === 'table' ? table.rows[path[i + 1]] : undefined;
    const cell = row ? row.cells[path[i + 2]] : undefined;
    if (!cell) return undefined;
    list = cell.blocks;
  }
  const block = list[path[path.length - 1]];
  return block && block.type === 'paragraph' ? block : undefined;
}

/** The blocks with the paragraph at a path (one paragraphOfBlocks() finds) in the place of the one that is there; the tables on the way are copies. */
function withParagraphAt(blocks: Block[], path: readonly number[], paragraph: Paragraph, at = 0): Block[] {
  const out = blocks.slice();
  if (at === path.length - 1) {
    out[path[at]] = paragraph;
    return out;
  }
  const table = blocks[path[at]];
  if (!table || table.type !== 'table') return bad(COMMENT_GONE);
  const rows = table.rows.slice();
  const row = rows[path[at + 1]];
  const cells = row.cells.slice();
  const cell = cells[path[at + 2]];
  cells[path[at + 2]] = { ...cell, blocks: withParagraphAt(cell.blocks, path, paragraph, at + 3) };
  rows[path[at + 1]] = { ...row, cells };
  out[path[at]] = { ...table, rows };
  return out;
}

/**
 * Where a place of a paragraph of a comment (a path and an offset, as Find gives them) is in
 * the text of the comment (commentText(): its paragraphs separated by "\n"); undefined for a
 * place the comment does not have.
 */
export function commentTextOffset(comment: { blocks: Block[] }, path: readonly number[], offset: number): number | undefined {
  let before = 0;
  let found: number | undefined;
  const walk = (blocks: Block[], prefix: number[], depth: number): boolean => {
    if (depth > MAX_DEPTH) return false;
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      if (block.type === 'paragraph') {
        const length = paragraphText(block).length;
        const here = prefix.length + 1 === path.length && path[path.length - 1] === i && prefix.every((step, k) => step === path[k]);
        if (here) {
          if (offset >= 0 && offset <= length) found = before + offset;
          return true;
        }
        before += length + 1;
      } else if (block.type === 'table') {
        for (let r = 0; r < block.rows.length; r++) {
          const row = block.rows[r];
          for (let c = 0; c < row.cells.length; c++) if (walk(row.cells[c].blocks, prefix.concat([i, r, c]), depth + 1)) return true;
        }
      }
    }
    return false;
  };
  walk(comment.blocks, [], 0);
  return found;
}

/** How many passes over a paragraph of a comment its replacements made since the module was loaded: for tests, which check that the matches of a paragraph take one. */
export const commentReplaceCounters = { passes: 0 };

/** What replaceInComment() made of a comment. */
export interface CommentReplacement {
  /** The edit that changes the comment; undefined when nothing was replaced. */
  edit?: CommentEdit;
  replaced: number;
  /** Ranges that could not be replaced; they stay as they are. */
  failed: number;
}

/**
 * Replaces ranges of the text of one comment with the same text (Find and Replace): each range
 * is one paragraph of the comment (the story is commentStory(id), see COMMENT_STORY), and the
 * new text takes the place of the replaced characters as replaceText() does - with the
 * formatting of the first of them, in their run - so that nothing else of the comment changes:
 * not its other text, its formatting, its fields or its paragraphs. A range that holds anything
 * but plain text of runs, or begins in a tracked change, is not replaced (`failed`); nor is any
 * when the comment would grow beyond MAX_COMMENT_TEXT. The text has no paragraph end: one with
 * one replaces nothing. One edit, never a tracked change (as editComment()); `selection` is the
 * selection of the document it leaves.
 */
export function replaceInComment(doc: DocModel, id: number, ranges: readonly DocRange[], text: string, selection?: DocRange): CommentReplacement {
  if (typeof text !== 'string') bad('The text of a comment must be a string');
  commentAt(doc, id);
  const story = commentStory(id);
  const mine = ranges.filter((range) => range && range.anchor && range.anchor.story === story);
  const out = replaceInComments(doc, mine, text, selection);
  out.failed += ranges.length - mine.length;
  return out;
}

/**
 * replaceInComment() for ranges of any number of comments (each in the story of its comment), as one edit: the
 * comments of a "Replace all" are one change of the table of comments, not one for each comment (which would copy
 * the table once for each). A range in no comment the document has fails.
 */
export function replaceInComments(doc: DocModel, ranges: readonly DocRange[], text: string, selection?: DocRange): CommentReplacement {
  if (typeof text !== 'string') bad('The text of a comment must be a string');
  const items = commentsOf(doc);
  const clean = cleanText(text);
  const byComment = new Map<number, { path: number[]; start: number; end: number }[]>();
  let failed = 0;
  for (const range of ranges) {
    const a = range && range.anchor;
    const f = range && range.focus;
    const id = a ? commentOfStory(a.story) : undefined;
    const ok =
      id !== undefined && items.has(id) && f && f.story === a.story && Array.isArray(a.path) && Array.isArray(f.path) && a.path.length === f.path.length && a.path.every((step, k) => step === f.path[k]) && Number.isInteger(a.offset) && Number.isInteger(f.offset);
    if (!ok || a.offset === f.offset) {
      failed++;
      continue;
    }
    let places = byComment.get(id as number);
    if (!places) byComment.set(id as number, (places = []));
    places.push({ path: a.path.slice(), start: Math.min(a.offset, f.offset), end: Math.max(a.offset, f.offset) });
  }
  let replaced = 0;
  const made: Comment[] = [];
  byComment.forEach((places, id) => {
    const comment = items.get(id) as Comment;
    if (clean.indexOf('\n') >= 0) {
      failed += places.length;
      return;
    }
    // The places of each paragraph, in the order of the text: a paragraph is gone over once for all of them.
    const byParagraph = new Map<string, { path: number[]; places: { start: number; end: number }[] }>();
    for (const place of places) {
      const key = place.path.join(',');
      let group = byParagraph.get(key);
      if (!group) byParagraph.set(key, (group = { path: place.path, places: [] }));
      group.places.push({ start: place.start, end: place.end });
    }
    const before = commentText(comment).length;
    let length = before;
    let blocks = comment.blocks;
    let here = 0;
    byParagraph.forEach((group) => {
      group.places.sort((x, y) => x.start - y.start || x.end - y.end);
      const paragraph = paragraphOfBlocks(comment.blocks, group.path);
      if (!paragraph) {
        failed += group.places.length;
        return;
      }
      commentReplaceCounters.passes++;
      // A place that overlaps the one before it is not replaced (its text would be gone): replacedParagraphRanges() says so.
      const result = replacedParagraphRanges(paragraph, group.places, clean);
      for (let i = 0; i < group.places.length; i++) {
        if (!result.done[i]) {
          failed++;
          continue;
        }
        here++;
        length += clean.length - (group.places[i].end - group.places[i].start);
      }
      if (result.paragraph) blocks = withParagraphAt(blocks, group.path, result.paragraph);
    });
    if (here === 0) return;
    // A comment that would be longer than a comment may be, and longer than it was, is not changed at all; one that a
    // file made longer may still be made shorter.
    if (length > MAX_COMMENT_TEXT && length > before) {
      failed += here;
      return;
    }
    const next = changed(comment, { blocks });
    made.push(next);
    replaced += here;
  });
  if (made.length === 0) return { replaced, failed };
  const edit = build(doc, (draft) => {
    const next = new Map(items);
    for (const comment of made) next.set(comment.id, comment);
    draft.apply({ type: 'setComments', items: next });
    return { selection: kept(doc, selection) };
  });
  return { edit, replaced, failed };
}

/** The paragraphs a set of places is in, each once, with what happens there. */
function groupPlaces<T>(entries: { place: Place; value: T }[]): Map<string, { place: Place; values: { inline: number; value: T }[] }> {
  const out = new Map<string, { place: Place; values: { inline: number; value: T }[] }>();
  for (const entry of entries) {
    const key = entry.place.story + '\u0000' + entry.place.container.join(',') + '\u0000' + entry.place.index;
    let group = out.get(key);
    if (!group) out.set(key, (group = { place: entry.place, values: [] }));
    group.values.push({ inline: entry.place.inline, value: entry.value });
  }
  return out;
}

function paragraphAt(doc: DocModel, place: Place): Paragraph {
  const blocks = resolveContainer(doc, place.story, place.container);
  const block = blocks[place.index];
  return block && block.type === 'paragraph' ? block : bad(COMMENT_GONE);
}

/**
 * Adds an answer to a comment: a comment of its own that says which thread
 * it belongs to (the first comment of the thread of `id`), with marks beside
 * those of that comment, after the marks of the answers the thread has
 * there already. Throws RangeError(COMMENT_FIXED) for a comment
 * whose marks cannot be followed (nothing refers to it, or they sit in
 * content that is kept whole).
 */
export function replyToComment(doc: DocModel, id: number, text: string, options?: CommentOptions, selection?: DocRange): CommentEdit {
  const clean = checkedText(text);
  const items = commentsOf(doc);
  const root = rootOf(items, commentAt(doc, id));
  const anchors = anchorsOf(doc).get(root.id);
  if (!anchors || anchors.fixed || anchors.refs.length === 0 || anchors.refs[0].inline < 0 || !hasLastParagraph(root)) bad(COMMENT_FIXED);
  if (items.size >= MAX_COMMENTS) bad('The document has no room for another comment');
  const who = trackingFor(doc, { trackChanges: true, author: options && options.author, date: options && options.date });
  if (!who) return bad(COMMENT_FIXED);
  const at = anchors as Anchors;
  return build(doc, (draft) => {
    ensureCommentStyles(doc, draft);
    const replyId = nextRevisionId(doc);
    const next = new Map(items);
    // The thread is known by the paragraph id of its first comment.
    if (root.paraId === undefined) next.set(root.id, changed(root, { paraId: freshParaId(doc, root.id) }));
    const reply: Comment = { id: replyId, author: who.author, date: who.date, blocks: contentOf(doc, clean), paraId: freshParaId(doc, replyId), parent: root.id };
    const initials = options && typeof options.initials === 'string' ? options.initials : initialsOf(who.author);
    if (initials !== '') reply.initials = initials;
    // An answer to a thread that is resolved belongs to a resolved thread.
    if (root.done) reply.done = true;
    next.set(replyId, reply);
    draft.apply({ type: 'setComments', items: next });

    // The marks of the answer, each right after the last one the thread has of its kind beside the one of the first
    // comment, so that the answers stand in the text in the order they were given; a range only when the first
    // comment has both of its ends in the text.
    const all = anchorsOf(doc);
    const thread = threadOf(items, root);
    const last = (kind: 'starts' | 'ends' | 'refs'): Place => {
      let place = at[kind][0];
      for (const member of thread) {
        const other = member === root ? undefined : all.get(member.id);
        const found = other && !other.fixed ? other[kind][0] : undefined;
        // Only a mark in the same paragraph: one that a file has elsewhere is not a place beside the first comment.
        if (!found || found.inline <= place.inline || found.index !== place.index || found.story !== place.story) continue;
        if (found.container.length !== place.container.length || found.container.some((step, k) => step !== place.container[k])) continue;
        place = found;
      }
      return place;
    };
    const entries: { place: Place; value: 'start' | 'end' | 'ref' }[] = [{ place: last('refs'), value: 'ref' }];
    if (at.starts.length && at.ends.length && at.starts[0].inline >= 0 && at.ends[0].inline >= 0) {
      entries.push({ place: last('starts'), value: 'start' }, { place: last('ends'), value: 'end' });
    }
    groupPlaces(entries).forEach((group) => {
      const paragraph = paragraphAt(doc, group.place);
      const children = paragraph.children.slice();
      // From the back, so that the places before stay what they are.
      group.values.sort((a, b) => b.inline - a.inline);
      for (const entry of group.values) {
        const beside = children[entry.inline];
        let inline: Inline;
        if (entry.value === 'ref') {
          const run: Run = { type: 'run', props: beside.type === 'run' ? beside.props : referenceProps(doc), content: [{ type: 'commentReference', id: replyId }] };
          if (beside.wrap) run.wrap = beside.wrap;
          inline = run;
        } else inline = makeMark(entry.value === 'start' ? 'commentRangeStart' : 'commentRangeEnd', replyId, beside.wrap);
        children.splice(entry.inline + 1, 0, inline);
      }
      draft.setParagraph(group.place.story, group.place.container.concat([group.place.index]), paragraphWith(paragraph, children));
    });
    return { selection: kept(doc, selection), id: replyId };
  });
}

/**
 * Says that the thread a comment belongs to is resolved, or that it is open
 * again: every comment of the thread says so. No operations when they do
 * already. Throws RangeError(COMMENT_FIXED) for a comment without a
 * paragraph to be known by.
 */
export function resolveComment(doc: DocModel, id: number, done: boolean, selection?: DocRange): CommentEdit {
  if (typeof done !== 'boolean') bad('A comment is either resolved or not');
  const items = commentsOf(doc);
  const root = rootOf(items, commentAt(doc, id));
  const thread = threadOf(items, root);
  return build(doc, (draft) => {
    let next: Map<number, Comment> | undefined;
    for (const comment of thread) {
      const is = comment.done === true;
      if (is === done && (comment.paraId !== undefined || !done)) continue;
      if (!hasLastParagraph(comment)) bad(COMMENT_FIXED);
      const change: Partial<Comment> = {};
      if (comment.paraId === undefined) change.paraId = freshParaId(doc, comment.id);
      const out = changed(comment, change);
      if (done) out.done = true;
      else delete out.done;
      if (!next) next = new Map(items);
      next.set(comment.id, out);
    }
    if (next) draft.apply({ type: 'setComments', items: next });
    return { selection: kept(doc, selection) };
  });
}

/**
 * Removes a comment with its marks; the first comment of a thread takes its
 * answers along. Throws RangeError(COMMENT_FIXED) when one of the marks sits
 * where it cannot be taken out (a text box, content that is kept whole): a
 * mark without its comment must not stay.
 */
export function removeComment(doc: DocModel, id: number, selection?: DocRange): CommentEdit {
  const items = commentsOf(doc);
  const comment = commentAt(doc, id);
  const going = comment.parent !== undefined && rootOf(items, comment) !== comment ? [comment] : threadOf(items, comment);
  const anchors = anchorsOf(doc);
  const ids = new Set<number>();
  const entries: { place: Place; value: number }[] = [];
  const between: Place[] = [];
  for (const one of going) {
    ids.add(one.id);
    const at = anchors.get(one.id);
    if (!at) continue;
    if (at.fixed) bad(COMMENT_FIXED);
    for (const place of at.starts.concat(at.ends, at.refs)) {
      if (place.inline < 0) between.push(place);
      else entries.push({ place, value: one.id });
    }
  }
  return build(doc, (draft) => {
    const next = new Map(items);
    ids.forEach((gone) => next.delete(gone));
    draft.apply({ type: 'setComments', items: next });
    groupPlaces(entries).forEach((group) => {
      const paragraph = paragraphAt(doc, group.place);
      const children: Inline[] = [];
      for (const inline of paragraph.children) {
        if ((inline.type === 'commentRangeStart' || inline.type === 'commentRangeEnd') && COMMENT_ID.test(inline.id) && ids.has(Number(inline.id))) continue;
        if (inline.type !== 'run' || !inline.content.some((item) => item.type === 'commentReference' && ids.has(item.id))) {
          children.push(inline);
          continue;
        }
        const content: RunItem[] = [];
        const flags: boolean[] = [];
        const codes = inline.fieldCode;
        inline.content.forEach((item, k) => {
          if (item.type === 'commentReference' && ids.has(item.id)) return;
          content.push(item);
          flags.push(codes === true || (codes !== undefined && codes[k] === true));
        });
        // A run that held nothing but the reference goes with it.
        if (content.length === 0) continue;
        const run: Run = { ...inline, content };
        if (codes !== undefined && codes !== true) run.fieldCode = flags;
        children.push(run);
      }
      draft.setParagraph(group.place.story, group.place.container.concat([group.place.index]), paragraphWith(paragraph, children));
    });
    // Marks between blocks, the last first: taking one out moves the ones after it.
    const listOf = (place: Place): string => place.story + '\u0000' + place.container.join(',');
    between.sort((a, b) => (listOf(a) === listOf(b) ? b.index - a.index : listOf(a) < listOf(b) ? -1 : 1));
    for (const place of between) {
      const blocks = resolveContainer(doc, place.story, place.container);
      const block = blocks[place.index];
      if (!block || block.type !== 'mark' || !ids.has(Number(block.mark.id))) bad(COMMENT_GONE);
      draft.replace(place.story, place.container, place.index, 1, []);
    }
    return { selection: kept(doc, selection) };
  });
}

/** The comment whose text a position is in, or that is about that very place; the innermost (shortest) when several are. Undefined when there is none. */
export function commentAtPosition(list: DocComment[], position: DocPosition): DocComment | undefined {
  let best: DocComment | undefined;
  let bestSpan = Infinity;
  for (const comment of list) {
    if (comment.parent !== undefined || !comment.start || !comment.end || comment.story !== position.story) continue;
    const at: DocPosition = { story: position.story, path: position.path, offset: position.offset };
    let inside: boolean;
    try {
      inside = comparePositions(comment.start, at) <= 0 && comparePositions(at, comment.end) <= 0;
    } catch {
      inside = false;
    }
    if (!inside) continue;
    const span = comment.start.path.join(',') === comment.end.path.join(',') ? comment.end.offset - comment.start.offset : 1e9;
    if (span < bestSpan) {
      best = comment;
      bestSpan = span;
    }
  }
  return best;
}
