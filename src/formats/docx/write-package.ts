// Parts of a docx package from the document model: the text of a story part
// (body, header, footer, notes, comments) and the supporting parts a new
// package needs (styles, numbering, settings, font table).
//
// A story is written with as much of its source text as still stands. A block
// that carries its source span is copied from the source, together with the
// white space before it; a block without one (an edit, see Story.modified) is
// written from the model. A wrapper around top-level blocks (a content
// control, say) carries one span for all of them, and its blocks carry none:
// it is copied only when reading its source again gives what the model holds.
// The start of the part (declaration, root element with its namespace
// declarations) and its end are copied too as long as the model still says
// what they say.

import type {
  AbstractNum,
  AttrList,
  Block,
  Body,
  Comment,
  DocModel,
  DocRelationship,
  DocStyle,
  DocStyles,
  ExtraXml,
  FontInfo,
  MiscXml,
  Note,
  NumInstance,
  NumLevel,
  Numbering,
  Story,
  StoryCollection,
  TableStyleCondition,
} from '../../model/doc/types';
import { parseXml, type XmlElement } from '../../xml';
import { XmlScanner } from '../../xml-scan';
import { readBlock } from './body';
import { bounded, createContext, createReaderState, kindOf } from './common';
import { SourceCursor } from './source';
import {
  Group,
  cleanText,
  groupByWrapper,
  treeText,
  writeBlockNode,
  writeMisc,
  type BodyContext,
  type Node,
} from './write-body';
import {
  cellPropsElement,
  keptElement,
  paraPropsElement,
  rowPropsElement,
  runPropsElement,
  sectionPropsElement,
  tablePropsElement,
  wEl,
} from './write-props';

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export interface DocxNamespaces {
  w: string;
  r: string;
  /** Relationship types start with this. */
  rel: string;
}

export const TRANSITIONAL_NS: DocxNamespaces = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/',
};

export const STRICT_NS: DocxNamespaces = {
  w: 'http://purl.oclc.org/ooxml/wordprocessingml/main',
  r: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  rel: 'http://purl.oclc.org/ooxml/officeDocument/relationships/',
};

/** Namespace declarations of a part written from nothing. */
export function generatedRootAttrs(strict: boolean): AttrList {
  const ns = strict ? STRICT_NS : TRANSITIONAL_NS;
  return [
    ['xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006'],
    ['xmlns:o', 'urn:schemas-microsoft-com:office:office'],
    ['xmlns:r', ns.r],
    ['xmlns:m', strict ? 'http://purl.oclc.org/ooxml/officeDocument/math' : 'http://schemas.openxmlformats.org/officeDocument/2006/math'],
    ['xmlns:v', 'urn:schemas-microsoft-com:vml'],
    ['xmlns:wp', strict ? 'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing' : 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'],
    ['xmlns:w10', 'urn:schemas-microsoft-com:office:word'],
    ['xmlns:w', ns.w],
    ['xmlns:w14', 'http://schemas.microsoft.com/office/word/2010/wordml'],
    ['xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml'],
    ['xmlns:wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'],
    ['mc:Ignorable', 'w14 w15 wp14'],
  ];
}

export interface WriteCounters {
  dropped: { count: number };
  drawingIds: { next: number };
}

/** What the writers need for one part: its prefixes, and a reader for the trees the model kept from it. */
export function createBodyContext(
  part: string,
  rootAttrs: AttrList,
  relationships: Map<string, DocRelationship>,
  shared?: WriteCounters,
): BodyContext {
  const state = createReaderState();
  const reader = createContext(state, part, rootAttrs, relationships);
  // The prefix bound to WordprocessingML; none when it is the default namespace of the part.
  const w = reader.wPrefixes.length ? reader.wPrefixes[0] : reader.prefixes.get('') === 'w' ? '' : 'w:';
  return {
    w,
    a: w,
    r: reader.rPrefixes[0],
    strict: state.strict,
    reader,
    plain: false,
    dropped: shared ? shared.dropped : { count: 0 },
    runCache: new WeakMap(),
    paraCache: new WeakMap(),
    drawingIds: shared ? shared.drawingIds : { next: 1 },
    controls: new Set(),
    splitControls: { count: 0 },
    boxTexts: new WeakMap(),
  };
}

export interface StoryWriteOptions {
  /** Every block is written from the model, whatever source span it carries. */
  regenerate?: boolean;
  /** Counts top-level blocks copied from the source and written from the model. */
  counts?: { copied: number; written: number };
  /** Root element to write when the story has none of its own (a model that was not read from a package). */
  root?: { name: string; attrs: AttrList };
}

const isBlank = (text: string, from: number, to: number): boolean => {
  for (let i = from; i < to; i++) {
    const c = text.charCodeAt(i);
    if (c !== 32 && c !== 9 && c !== 10 && c !== 13) return false;
  }
  return true;
};

const isSpaceAt = (text: string, at: number): boolean => {
  const c = text.charCodeAt(at);
  return c === 32 || c === 9 || c === 10 || c === 13;
};

/**
 * Collects the text of a part: slices of its source, joined where they are
 * adjacent, and text written from the model. The white space around a copied
 * slice goes with it (the containers written this way hold elements only, so
 * it means nothing, but a part that was indented stays indented): the space
 * before it, and the space after it when text from the model follows.
 */
class Emitter {
  private readonly out: string[] = [];
  private runStart = -1;
  private runEnd = -1;
  /** Offset up to which the source was copied; white space before it is not copied again. */
  private copied = 0;
  /** Offset in the source just past the last thing copied from it; undefined after text written from the model. */
  at: number | undefined;

  constructor(readonly source: string) {}

  /** Copies source[start, end). */
  copy(start: number, end: number): void {
    if (this.runEnd >= 0 && start >= this.runEnd && isBlank(this.source, this.runEnd, start)) this.runEnd = end;
    else {
      this.flush();
      let lead = start;
      const floor = start >= this.copied ? this.copied : 0;
      while (lead > floor && isSpaceAt(this.source, lead - 1)) lead--;
      this.runStart = lead;
      this.runEnd = end;
    }
    if (end > this.copied) this.copied = end;
    this.at = end;
  }

  text(text: string): void {
    if (this.runEnd >= 0) {
      // The white space after what was copied last: it stood before the element this text stands for.
      while (this.runEnd < this.source.length && isSpaceAt(this.source, this.runEnd)) this.runEnd++;
      if (this.runEnd > this.copied) this.copied = this.runEnd;
    }
    this.flush();
    this.out.push(text);
    this.at = undefined;
  }

  private flush(): void {
    if (this.runEnd >= 0) this.out.push(this.source.slice(this.runStart, this.runEnd));
    this.runStart = this.runEnd = -1;
  }

  done(): string {
    this.flush();
    return this.out.join('');
  }
}

/** True when [start, end) can be the source of an element or comment. */
function spanOk(source: string, start: number | undefined, end: number | undefined): boolean {
  return (
    start !== undefined &&
    end !== undefined &&
    start >= 0 &&
    end > start + 1 &&
    end <= source.length &&
    source.charCodeAt(start) === 60 &&
    source.charCodeAt(end - 1) === 62
  );
}

function countBlocks(node: Node<Block>): number {
  if (!(node instanceof Group)) return 1;
  let count = 0;
  for (const child of node.children) count += countBlocks(child);
  return count;
}

/** What the writer makes of the source of a wrapper as the reader reads it now; undefined when it cannot be read. */
function rereadGroup(ctx: BodyContext, raw: string): string | undefined {
  try {
    const blocks: Block[] = [];
    readBlock(ctx.reader, bounded(parseXml(raw).root), blocks, undefined, 0);
    const plain: BodyContext = { ...ctx, plain: true };
    const out: string[] = [];
    for (const node of groupByWrapper(blocks)) writeBlockNode(plain, node, out, 0);
    return out.join('');
  } catch (error) {
    return undefined;
  }
}

interface SectionMark {
  /** Number of blocks before the mark; undefined when it comes last. */
  index: number | undefined;
  emit(): void;
}

function emitBlocks(ctx: BodyContext, em: Emitter, blocks: readonly Block[], options: StoryWriteOptions, section?: SectionMark): void {
  const source = em.source;
  const copying = !options.regenerate && source.length > 0;
  const counts = options.counts;
  let wrapped = false;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].wrap) {
      wrapped = true;
      break;
    }
  }
  const nodes: readonly Node<Block>[] = wrapped ? groupByWrapper(blocks) : blocks;
  let before = 0;
  let marked = !section;
  const part: string[] = [];
  for (const node of nodes) {
    if (!marked && section && section.index !== undefined && before >= section.index) {
      section.emit();
      marked = true;
    }
    if (node instanceof Group) {
      before += countBlocks(node);
      part.length = 0;
      writeBlockNode(ctx, node, part, 0);
      const text = part.join('');
      const wrapper = node.wrapper;
      if (copying && spanOk(source, wrapper.srcStart, wrapper.srcEnd) && rereadGroup(ctx, source.slice(wrapper.srcStart, wrapper.srcEnd)) === text) {
        em.copy(wrapper.srcStart as number, wrapper.srcEnd as number);
        if (counts) counts.copied++;
      } else {
        em.text(text);
        if (counts) counts.written++;
      }
      continue;
    }
    before++;
    if (copying && spanOk(source, node.srcStart, node.srcEnd)) {
      em.copy(node.srcStart as number, node.srcEnd as number);
      if (counts) counts.copied++;
    } else {
      part.length = 0;
      writeBlockNode(ctx, node, part, 0);
      em.text(part.join(''));
      if (counts) counts.written++;
    }
  }
  if (!marked && section) section.emit();
}

const escapeRegExp = (text: string): string => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * Ends a container: the rest of the source when it is nothing but the given
 * closing tags (and what the story has after its root), else those tags
 * written out.
 */
function emitEnd(em: Emitter, closers: string[], between: string, epilog: string | undefined, to: number, copying: boolean): void {
  if (copying && em.at !== undefined && between === '' && em.at <= to) {
    const rest = em.source.slice(em.at, to);
    const match = new RegExp('^\\s*' + closers.map((name) => `</${escapeRegExp(name)}\\s*>`).join('\\s*')).exec(rest);
    if (match) {
      const after = rest.slice(match[0].length);
      if (epilog === undefined ? isBlank(after, 0, after.length) : after === epilog) {
        em.copy(em.at, to);
        return;
      }
    }
  }
  let text = '';
  closers.forEach((name, i) => {
    text += `</${name}>`;
    if (i === 0) text += between;
  });
  em.text(text + (epilog || ''));
}

function openTag(ctx: BodyContext, name: string, attrs: Iterable<[string, string]>): string {
  let text = '<' + name;
  for (const [key, value] of attrs) text += ` ${key}="${cleanText(ctx, value, true)}"`;
  return text + '>';
}

interface RootScan {
  sc: XmlScanner;
  cursor: SourceCursor;
  /** Offset just past the root start tag. */
  rootEnd: number;
  selfClosing: boolean;
}

/** Finds the root start tag of the source and says whether the story still has it as written. */
function scanRoot(story: { source: string; prolog?: string; rootName: string; rootAttrs: AttrList }): RootScan | undefined {
  try {
    const sc = new XmlScanner(story.source);
    sc.next();
    if (sc.kind !== 'open' || sc.name !== story.rootName) return undefined;
    const attrs = sc.attrs();
    if (attrs.length !== story.rootAttrs.length) return undefined;
    for (let i = 0; i < attrs.length; i++) {
      if (attrs[i][0] !== story.rootAttrs[i][0] || attrs[i][1] !== story.rootAttrs[i][1]) return undefined;
    }
    const cursor = new SourceCursor(story.source);
    cursor.toTag();
    if (cursor.lost || story.source.slice(0, cursor.pos) !== (story.prolog || '')) return undefined;
    cursor.enter();
    if (cursor.lost) return undefined;
    return { sc, cursor, rootEnd: cursor.pos, selfClosing: sc.selfClosing };
  } catch (error) {
    return undefined;
  }
}

interface StoryRoot {
  source: string;
  prolog?: string;
  epilog?: string;
  rootName: string;
  rootAttrs: AttrList;
}

/** The declaration and root start tag written from the model. */
function rootText(ctx: BodyContext, story: StoryRoot, options: StoryWriteOptions): { name: string; text: string } {
  if (story.rootName === '' && options.root) return { name: options.root.name, text: XML_DECLARATION + openTag(ctx, options.root.name, options.root.attrs) };
  if (story.rootName === '') throw new Error('Cannot write the document: a story has no root element');
  return { name: story.rootName, text: (story.prolog || '') + openTag(ctx, story.rootName, story.rootAttrs) };
}

/** Kept children and comments as text, in the order `index` gives them. */
function looseText(ctx: BodyContext, extras: ExtraXml[], misc: MiscXml[]): string {
  const items: { index: number; order: number; text: string }[] = [];
  extras.forEach((extra, i) => items.push({ index: extra.index, order: i + misc.length, text: treeText(ctx, extra.el) }));
  misc.forEach((note, i) => {
    const out: string[] = [];
    writeMisc(ctx, note.node, out);
    // A comment stands before the element with its index.
    items.push({ index: note.index - 0.5, order: i, text: out.join('') });
  });
  items.sort((x, y) => x.index - y.index || x.order - y.order);
  return items.map((item) => item.text).join('');
}

/** XML text of the main document part. */
export function bodyText(ctx: BodyContext, body: Body, options: StoryWriteOptions = {}): string {
  const source = body.source;
  const copying = !options.regenerate && source.length > 0;
  const em = new Emitter(source);
  const extras = (body.extras || []).filter((extra) => extra.index >= 0);
  const misc = body.misc || [];
  // The body is the child of the document that is not among the kept ones.
  const taken = new Set(extras.map((extra) => extra.index));
  let bodyIndex = 0;
  while (taken.has(bodyIndex)) bodyIndex++;
  const before = looseText(
    ctx,
    extras.filter((extra) => extra.index < bodyIndex),
    misc.filter((note) => note.index <= bodyIndex),
  );
  const after = looseText(
    ctx,
    extras.filter((extra) => extra.index > bodyIndex),
    misc.filter((note) => note.index > bodyIndex),
  );

  let bodyName = ctx.w + 'body';
  let rootName = body.rootName;
  let head = false;
  const scan = copying ? scanRoot(body) : undefined;
  if (scan && !scan.selfClosing) {
    try {
      const { sc, cursor } = scan;
      while (sc.next()) {
        if (sc.kind === 'text') continue;
        cursor.toTag();
        if (sc.kind === 'close' || cursor.lost) break;
        if (kindOf(ctx.reader, sc.name) !== 'body') {
          if (!cursor.skip(sc.readRaw())) break;
          continue;
        }
        bodyName = sc.name;
        const bodyStart = cursor.pos;
        cursor.enter();
        if (cursor.lost) break;
        // What stands between the root tag and the body, as the writer would write it.
        const between: string[] = [];
        for (const node of parseXml(`<x>${source.slice(scan.rootEnd, bodyStart)}</x>`).root.children) {
          if (node.type === 'element') between.push(treeText(ctx, node));
          else if (node.type !== 'text') writeMisc(ctx, node, between);
        }
        if (between.join('') !== before) break;
        if (sc.selfClosing) {
          em.copy(0, bodyStart);
          em.text(`<${bodyName}>`);
        } else em.copy(0, cursor.pos);
        head = true;
        break;
      }
    } catch (error) {
      head = false;
    }
  }
  if (!head) {
    const root = rootText(ctx, body, options);
    rootName = root.name;
    em.text(root.text + before + `<${bodyName}>`);
  }

  const section: SectionMark | undefined = body.section
    ? {
        index: body.sectionIndex,
        emit: () => {
          const span = body.sectionSpan;
          if (copying && span && spanOk(source, span.srcStart, span.srcEnd)) em.copy(span.srcStart as number, span.srcEnd as number);
          else if (body.section) em.text(treeText(ctx, sectionPropsElement(ctx, body.section)));
        },
      }
    : undefined;
  emitBlocks(ctx, em, body.blocks, options, section);
  emitEnd(em, [bodyName, rootName], after, body.epilog, source.length, copying);
  return em.done();
}

/** Starts a part whose root holds the content directly; returns the name of the root. */
function emitRoot(ctx: BodyContext, em: Emitter, story: StoryRoot, options: StoryWriteOptions): string {
  const scan = !options.regenerate && story.source.length > 0 ? scanRoot(story) : undefined;
  if (scan && !scan.selfClosing) {
    em.copy(0, scan.rootEnd);
    return story.rootName;
  }
  const root = rootText(ctx, story, options);
  em.text(root.text);
  return root.name;
}

/** XML text of a header or footer part. */
export function storyText(ctx: BodyContext, story: Story, options: StoryWriteOptions = {}): string {
  const copying = !options.regenerate && story.source.length > 0;
  const em = new Emitter(story.source);
  const rootName = emitRoot(ctx, em, story, options);
  emitBlocks(ctx, em, story.blocks, options);
  emitEnd(em, [rootName], '', story.epilog, story.source.length, copying);
  return em.done();
}

/** With the offset of the "<" of a start tag: the offset just past the tag, its name, and whether it closes itself. */
function startTagAt(source: string, start: number, limit: number): { end: number; name: string; selfClosing: boolean } | undefined {
  let i = start + 1;
  while (i < limit) {
    const c = source.charCodeAt(i);
    if (c === 62) {
      let nameEnd = start + 1;
      while (nameEnd < i && !/[\s/>]/.test(source.charAt(nameEnd))) nameEnd++;
      return { end: i + 1, name: source.slice(start + 1, nameEnd), selfClosing: source.charCodeAt(i - 1) === 47 };
    }
    if (c === 34 || c === 39) {
      const close = source.indexOf(c === 34 ? '"' : "'", i + 1);
      if (close < 0 || close >= limit) return undefined;
      i = close + 1;
    } else i++;
  }
  return undefined;
}

/**
 * XML text of a footnotes, endnotes or comments part. `local` is the local
 * name of the item elements ("footnote", "endnote", "comment").
 */
export function collectionText(
  ctx: BodyContext,
  collection: StoryCollection<Note> | StoryCollection<Comment>,
  local: 'footnote' | 'endnote' | 'comment',
  options: StoryWriteOptions = {},
): string {
  const source = collection.source;
  const copying = !options.regenerate && source.length > 0;
  const em = new Emitter(source);
  const rootName = emitRoot(ctx, em, collection, options);

  const item = (entry: Note | Comment): void => {
    const name = ctx.w + local;
    const a = ctx.a;
    if (copying && spanOk(source, entry.srcStart, entry.srcEnd)) {
      const start = entry.srcStart as number;
      const end = entry.srcEnd as number;
      const tag = startTagAt(source, start, end);
      if (tag && tag.selfClosing && entry.blocks.length === 0) {
        em.copy(start, end);
        return;
      }
      if (tag && !tag.selfClosing) {
        em.copy(start, tag.end);
        emitBlocks(ctx, em, entry.blocks, options);
        emitEnd(em, [tag.name], '', '', end, true);
        return;
      }
    }
    const attrs: [string, string][] = [];
    if (local === 'comment') {
      const comment = entry as Comment;
      attrs.push([a + 'id', String(entry.id)]);
      // The schema requires an author.
      attrs.push([a + 'author', comment.author === undefined ? '' : comment.author]);
      if (comment.date !== undefined) attrs.push([a + 'date', comment.date]);
      if (comment.initials !== undefined) attrs.push([a + 'initials', comment.initials]);
    } else {
      const note = entry as Note;
      if (note.kind !== undefined) attrs.push([a + 'type', note.kind]);
      attrs.push([a + 'id', String(entry.id)]);
    }
    for (const pair of entry.attrs || []) attrs.push(pair);
    em.text(openTag(ctx, name, attrs));
    // The blocks of a note that an edit replaced still carry their spans where they were not touched.
    emitBlocks(ctx, em, entry.blocks, options);
    em.text(`</${name}>`);
  };

  // Items in their order, with the kept children of the root at the places they had.
  const extras = (collection.extras || []).filter((extra) => extra.index >= 0);
  const notes = (collection.misc || []).slice();
  const entries = Array.from(collection.items.values() as Iterable<Note | Comment>);
  const total = entries.length + extras.length;
  const slots: (XmlElement | undefined)[] = new Array(total);
  const late: XmlElement[] = [];
  for (const extra of extras) {
    if (extra.index < total && slots[extra.index] === undefined) slots[extra.index] = extra.el;
    else late.push(extra.el);
  }
  const comments = (slot: number) => {
    for (let i = 0; i < notes.length; i++) {
      if (notes[i].index > slot) continue;
      const out: string[] = [];
      writeMisc(ctx, notes[i].node, out);
      em.text(out.join(''));
      notes.splice(i--, 1);
    }
  };
  let next = 0;
  for (let slot = 0; slot < total; slot++) {
    comments(slot);
    const kept = slots[slot];
    if (kept) em.text(treeText(ctx, kept));
    else if (next < entries.length) item(entries[next++]);
  }
  while (next < entries.length) item(entries[next++]);
  for (const el of late) em.text(treeText(ctx, el));
  comments(Infinity);
  emitEnd(em, [rootName], '', collection.epilog, source.length, copying);
  return em.done();
}

// ------------------------------------------------------------------ numbering

function ranks(names: string[]): Map<string, number> {
  const map = new Map<string, number>();
  names.forEach((name, i) => map.set(name, i));
  return map;
}

/** CT_Lvl. */
export const LVL_ORDER = ranks(['start', 'numFmt', 'lvlRestart', 'pStyle', 'isLgl', 'suff', 'lvlText', 'lvlPicBulletId', 'legacy', 'lvlJc', 'pPr', 'rPr']);

/** CT_Style. */
export const STYLE_ORDER = ranks([
  'name',
  'aliases',
  'basedOn',
  'next',
  'link',
  'autoRedefine',
  'hidden',
  'uiPriority',
  'semiHidden',
  'unhideWhenUsed',
  'qFormat',
  'locked',
  'personal',
  'personalCompose',
  'personalReply',
  'rsid',
  'pPr',
  'rPr',
  'tblPr',
  'trPr',
  'tcPr',
  'tblStylePr',
]);

type Gen = { kind: string; el: XmlElement };

function levelElement(ctx: BodyContext, level: NumLevel): XmlElement {
  const gens: Gen[] = [];
  const add = (kind: string, el: XmlElement) => gens.push({ kind, el });
  if (level.start !== undefined) add('start', wEl(ctx, 'start', [['val', level.start]]));
  if (level.format !== undefined) add('numFmt', wEl(ctx, 'numFmt', [['val', level.format]]));
  if (level.restart !== undefined) add('lvlRestart', wEl(ctx, 'lvlRestart', [['val', level.restart]]));
  if (level.paraStyle !== undefined) add('pStyle', wEl(ctx, 'pStyle', [['val', level.paraStyle]]));
  if (level.legal !== undefined) add('isLgl', wEl(ctx, 'isLgl', level.legal ? undefined : [['val', '0']]));
  if (level.suffix !== undefined) add('suff', wEl(ctx, 'suff', [['val', level.suffix]]));
  if (level.text !== undefined) add('lvlText', wEl(ctx, 'lvlText', [['val', level.text]]));
  if (level.align !== undefined) add('lvlJc', wEl(ctx, 'lvlJc', [['val', level.align]]));
  if (level.para) add('pPr', paraPropsElement(ctx, level.para) || wEl(ctx, 'pPr'));
  if (level.run) add('rPr', runPropsElement(ctx, level.run) || wEl(ctx, 'rPr'));
  return keptElement(ctx, 'lvl', LVL_ORDER, gens, level.extras, [['ilvl', level.level]]);
}

function abstractElement(ctx: BodyContext, abstract: AbstractNum): XmlElement {
  const children: XmlElement[] = [];
  if (abstract.multiLevelType !== undefined) children.push(wEl(ctx, 'multiLevelType', [['val', abstract.multiLevelType]]));
  if (abstract.styleLink !== undefined) children.push(wEl(ctx, 'styleLink', [['val', abstract.styleLink]]));
  if (abstract.numStyleLink !== undefined) children.push(wEl(ctx, 'numStyleLink', [['val', abstract.numStyleLink]]));
  const levels = Array.from(abstract.levels.values()).sort((x, y) => x.level - y.level);
  for (const level of levels) children.push(levelElement(ctx, level));
  return wEl(ctx, 'abstractNum', [['abstractNumId', abstract.id]], children);
}

function numElement(ctx: BodyContext, num: NumInstance): XmlElement {
  const children: XmlElement[] = [wEl(ctx, 'abstractNumId', [['val', num.abstractId]])];
  const levels = Array.from(num.overrides.keys()).sort((x, y) => x - y);
  for (const index of levels) {
    const override = num.overrides.get(index);
    if (!override) continue;
    const inner: XmlElement[] = [];
    if (override.start !== undefined) inner.push(wEl(ctx, 'startOverride', [['val', override.start]]));
    if (override.level) inner.push(levelElement(ctx, override.level));
    children.push(wEl(ctx, 'lvlOverride', [['ilvl', index]], inner));
  }
  return wEl(ctx, 'num', [['numId', num.id]], children);
}

/** The definitions of the model as XML text: abstract definitions first, as the schema wants them. */
export function numberingTexts(ctx: BodyContext, numbering: Numbering, only?: { abstracts: Set<number>; nums: Set<number> }): { abstracts: string; nums: string } {
  let abstracts = '';
  let nums = '';
  for (const abstract of numbering.abstracts.values()) if (!only || only.abstracts.has(abstract.id)) abstracts += treeText(ctx, abstractElement(ctx, abstract));
  for (const num of numbering.nums.values()) if (!only || only.nums.has(num.id)) nums += treeText(ctx, numElement(ctx, num));
  return { abstracts, nums };
}

// --------------------------------------------------------------------- styles

function formattingGens(ctx: BodyContext, from: TableStyleCondition, gens: Gen[]): void {
  if (from.para) gens.push({ kind: 'pPr', el: paraPropsElement(ctx, from.para) || wEl(ctx, 'pPr') });
  if (from.run) gens.push({ kind: 'rPr', el: runPropsElement(ctx, from.run) || wEl(ctx, 'rPr') });
  if (from.table) gens.push({ kind: 'tblPr', el: tablePropsElement(ctx, from.table) || wEl(ctx, 'tblPr') });
  if (from.row) gens.push({ kind: 'trPr', el: rowPropsElement(ctx, from.row) || wEl(ctx, 'trPr') });
  if (from.cell) gens.push({ kind: 'tcPr', el: cellPropsElement(ctx, from.cell) || wEl(ctx, 'tcPr') });
}

export function styleElement(ctx: BodyContext, style: DocStyle): XmlElement {
  const gens: Gen[] = [];
  const add = (kind: string, el: XmlElement) => gens.push({ kind, el });
  const flag = (kind: string, value: boolean | undefined) => {
    if (value !== undefined) add(kind, wEl(ctx, kind, value ? undefined : [['val', '0']]));
  };
  if (style.name !== undefined) add('name', wEl(ctx, 'name', [['val', style.name]]));
  if (style.basedOn !== undefined) add('basedOn', wEl(ctx, 'basedOn', [['val', style.basedOn]]));
  if (style.next !== undefined) add('next', wEl(ctx, 'next', [['val', style.next]]));
  if (style.link !== undefined) add('link', wEl(ctx, 'link', [['val', style.link]]));
  flag('hidden', style.hidden);
  if (style.uiPriority !== undefined) add('uiPriority', wEl(ctx, 'uiPriority', [['val', style.uiPriority]]));
  flag('semiHidden', style.semiHidden);
  flag('unhideWhenUsed', style.unhideWhenUsed);
  flag('qFormat', style.quickFormat);
  flag('locked', style.locked);
  formattingGens(ctx, style, gens);
  if (style.conditions) {
    for (const [type, condition] of style.conditions) {
      const inner: Gen[] = [];
      formattingGens(ctx, condition, inner);
      add('tblStylePr', wEl(ctx, 'tblStylePr', [['type', type]], inner.map((gen) => gen.el)));
    }
  }
  return keptElement(ctx, 'style', STYLE_ORDER, gens, style.extras, [
    ['type', style.type],
    ['default', style.isDefault ? '1' : undefined],
    ['customStyle', style.custom ? '1' : undefined],
    ['styleId', style.id],
  ]);
}

/** <w:docDefaults> of the style sheet; empty text when it states nothing. */
export function docDefaultsText(ctx: BodyContext, styles: DocStyles): string {
  const run = runPropsElement(ctx, styles.defaults.run);
  const para = paraPropsElement(ctx, styles.defaults.para);
  if (!run && !para) return '';
  const children: XmlElement[] = [];
  if (run) children.push(wEl(ctx, 'rPrDefault', undefined, [run]));
  if (para) children.push(wEl(ctx, 'pPrDefault', undefined, [para]));
  return treeText(ctx, wEl(ctx, 'docDefaults', undefined, children));
}

// ------------------------------------------------- settings and the font table

const COMPAT_URI = 'http://schemas.microsoft.com/office/word';

/**
 * <w:documentProtection> for a new package. The element a docx was read with
 * goes back with every attribute it names by the prefix of the element (the
 * password hash among them); a protection without one (read from a doc or RTF
 * file) is written as its kind and whether it is enforced, without a password.
 */
function protectionElement(ctx: BodyContext, settings: DocModel['settings']): XmlElement | undefined {
  const protection = settings.protection;
  if (!protection) return undefined;
  if (protection.xml) return keptAttrs(ctx, 'documentProtection', protection.xml);
  const attrs: [string, string][] = [];
  if (protection.edit !== '') attrs.push(['edit', protection.edit]);
  attrs.push(['enforcement', protection.enforced ? '1' : '0']);
  return wEl(ctx, 'documentProtection', attrs);
}

/** An element of the settings with the attributes it was read with, those it names by its own prefix. */
function keptAttrs(ctx: BodyContext, local: string, xml: XmlElement): XmlElement {
  const attrs: [string, string][] = [];
  const colon = xml.name.indexOf(':');
  const prefix = colon < 0 ? '' : xml.name.slice(0, colon + 1);
  xml.attrs.forEach((value, name) => {
    const at = name.indexOf(':');
    // An attribute of another namespace has no declaration in the new part.
    if (at < 0 || name.slice(0, at + 1) === prefix) attrs.push([name.slice(at + 1), value]);
  });
  return wEl(ctx, local, attrs);
}

/** Inner XML of <w:settings> for a new package. Notes that are separators are listed, as Word does. */
export function settingsText(ctx: BodyContext, doc: DocModel): string {
  const s = doc.settings;
  const out: XmlElement[] = [];
  // In the order of CT_Settings.
  if (s.writeProtection) out.push(keptAttrs(ctx, 'writeProtection', s.writeProtection));
  if (s.mirrorMargins) out.push(wEl(ctx, 'mirrorMargins'));
  if (s.trackRevisions) out.push(wEl(ctx, 'trackRevisions'));
  const protection = protectionElement(ctx, s);
  if (protection) out.push(protection);
  out.push(wEl(ctx, 'defaultTabStop', [['val', s.defaultTabStop]]));
  if (s.autoHyphenation) out.push(wEl(ctx, 'autoHyphenation'));
  if (s.evenAndOddHeaders) out.push(wEl(ctx, 'evenAndOddHeaders'));
  const separators = (collection: StoryCollection<Note> | undefined, local: string, item: string) => {
    if (!collection) return;
    const refs: XmlElement[] = [];
    for (const note of collection.items.values()) if (note.kind !== undefined) refs.push(wEl(ctx, item, [['id', note.id]]));
    if (refs.length) out.push(wEl(ctx, local, undefined, refs));
  };
  separators(doc.footnotes, 'footnotePr', 'footnote');
  separators(doc.endnotes, 'endnotePr', 'endnote');
  const compat: XmlElement[] = [];
  for (const [name, value] of s.compat) {
    if (/^[A-Za-z][A-Za-z0-9]{0,63}$/.test(name)) compat.push(wEl(ctx, name, value ? undefined : [['val', '0']]));
  }
  for (const [name, value] of s.compatSettings) {
    compat.push(
      wEl(ctx, 'compatSetting', [
        ['name', name],
        ['uri', COMPAT_URI],
        ['val', value],
      ]),
    );
  }
  if (compat.length) out.push(wEl(ctx, 'compat', undefined, compat));
  if (s.themeFontLang) {
    out.push(
      wEl(ctx, 'themeFontLang', [
        ['val', s.themeFontLang.val],
        ['eastAsia', s.themeFontLang.eastAsia],
        ['bidi', s.themeFontLang.bidi],
      ]),
    );
  }
  return out.map((el) => treeText(ctx, el)).join('');
}

/** Inner XML of <w:fonts>. Embedded font files are not carried over. */
export function fontTableText(ctx: BodyContext, fonts: Map<string, FontInfo>): string {
  let text = '';
  for (const font of fonts.values()) {
    const children: XmlElement[] = [];
    if (font.altName !== undefined) children.push(wEl(ctx, 'altName', [['val', font.altName]]));
    if (font.panose !== undefined) children.push(wEl(ctx, 'panose1', [['val', font.panose]]));
    if (font.charset !== undefined) children.push(wEl(ctx, 'charset', [['val', font.charset]]));
    if (font.family !== undefined) children.push(wEl(ctx, 'family', [['val', font.family]]));
    if (font.pitch !== undefined) children.push(wEl(ctx, 'pitch', [['val', font.pitch]]));
    text += treeText(ctx, wEl(ctx, 'font', [['name', font.name]], children));
  }
  return text;
}

// ------------------------------------------------------------------- scanning

export interface ScannedChild {
  /** Name the reader switches on ("style", "abstractNum", ...). */
  kind: string;
  attrs: AttrList;
  start: number;
  end: number;
}

export interface ScannedPart {
  rootName: string;
  rootAttrs: AttrList;
  /** Offset just past the root start tag. */
  contentStart: number;
  selfClosing: boolean;
  children: ScannedChild[];
  /** Offset of the root end tag; undefined for a root that closes itself. */
  closeStart: number | undefined;
}

/** The children of the root of a part with their places in its text; undefined when the text cannot be followed. */
export function scanPart(text: string): ScannedPart | undefined {
  try {
    const sc = new XmlScanner(text);
    sc.next();
    if (sc.kind !== 'open') return undefined;
    const rootName = sc.name;
    const rootAttrs = sc.attrs();
    const reader = createContext(createReaderState(), '', rootAttrs, new Map());
    const cursor = new SourceCursor(text);
    cursor.toTag();
    cursor.enter();
    if (cursor.lost) return undefined;
    const part: ScannedPart = { rootName, rootAttrs, contentStart: cursor.pos, selfClosing: sc.selfClosing, children: [], closeStart: undefined };
    if (sc.selfClosing) return part;
    // The scanner moves on; what it is on is no longer what the check above saw.
    const token = (): string => sc.kind;
    while (sc.next()) {
      if (token() === 'text') continue;
      cursor.toTag();
      if (cursor.lost) return undefined;
      if (token() === 'close') {
        part.closeStart = cursor.pos;
        break;
      }
      const start = cursor.pos;
      const kind = kindOf(reader, sc.name);
      const attrs = sc.attrs();
      if (!cursor.skip(sc.readRaw())) return undefined;
      part.children.push({ kind, attrs, start, end: cursor.pos });
    }
    return part.closeStart === undefined ? undefined : part;
  } catch (error) {
    return undefined;
  }
}

/** The text of a part as bytes: UTF-8, or UTF-16 with a byte order mark when its declaration says so. */
export function encodePart(text: string): Uint8Array {
  const declaration = /^<\?xml[^>]*\?>/.exec(text);
  if (declaration && /encoding\s*=\s*["']utf-16/i.test(declaration[0])) {
    const bytes = new Uint8Array(2 + text.length * 2);
    bytes[0] = 0xff;
    bytes[1] = 0xfe;
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      bytes[2 + i * 2] = c & 0xff;
      bytes[3 + i * 2] = c >> 8;
    }
    return bytes;
  }
  return new TextEncoder().encode(text);
}
