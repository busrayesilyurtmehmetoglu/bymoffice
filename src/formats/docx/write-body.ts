// Blocks of the document model -> WordprocessingML text: the reverse of
// body.ts.
//
// The model lists wrapped content flat, each item pointing to its wrapper
// (hyperlink, tracked change, content control, ...). The writer groups
// adjacent items that share a wrapper object back into one element.
//
// Trees the model kept whole (drawings, VML pictures, objects, math, unknown
// elements, alternate content) go back as they are. Where the model also
// holds something it read out of such a tree (the blocks of a text box, the
// items of an alternate content branch, the size of a drawing), the tree is
// read again and compared with the model: an unchanged tree is written as it
// is, a changed one gets the content of the model.
//   - A changed <mc:AlternateContent> is written as the content of the branch
//     the reader selected, without the element around it: the other branches
//     describe the old content and cannot be kept in step.
//   - A drawing without a DrawingML tree (a picture inserted by an edit, or a
//     drawing read from a binary doc or RTF file) is written as a standard
//     picture, text box or group.
//
// Characters XML 1.0 cannot hold (control characters, lone surrogates,
// U+FFFE, U+FFFF) are left out and counted: Word has no escape for them in
// document text.

import type {
  Block,
  DrawingAnchor,
  DrawingItem,
  ExtraXml,
  FieldCharItem,
  ImageCrop,
  ImageRef,
  Inline,
  MiscXml,
  Paragraph,
  ParaProps,
  RangeMark,
  Run,
  RunItem,
  RunProps,
  Table,
  TableCell,
  TableRow,
  Wrapper,
} from '../../model/doc/types';
import { parseXml, type XmlComment, type XmlElement, type XmlNode, type XmlPI } from '../../xml';
import { findSelected, readBlock } from './body';
import { kindOf, parseOnOff } from './common';
import {
  cellPropsElement,
  mergeKept,
  paraPropsElement,
  rowPropsElement,
  runPropsElement,
  tablePropsElement,
  wEl,
  type PropsContext,
} from './write-props';

/** Content nested deeper than this is not written: the reader would refuse it long before. */
export const MAX_WRITE_DEPTH = 200;
/** Depth limit for kept XML trees. */
const MAX_TREE_DEPTH = 300;
/** Longest chain of wrappers around one item. */
const MAX_WRAPPERS = 256;

export interface BodyContext extends PropsContext {
  /** Kept trees are written as they are, without checking what the model read out of them. */
  plain: boolean;
  /** Characters that XML cannot hold and that were left out. */
  dropped: { count: number };
  /** Property elements by the bag they were written from: bags are shared between many items. */
  runCache: WeakMap<RunProps, string>;
  paraCache: WeakMap<ParaProps, string>;
  /** The relationship id to write for a picture of a generated drawing; undefined writes the picture without an image. */
  image?(ref: ImageRef): string | undefined;
  /** The relationship id and anchor to write for a hyperlink; null writes its content without the link. */
  link?(wrapper: Extract<Wrapper, { kind: 'hyperlink' }>): { rId?: string; anchor?: string } | null;
  /** Ids for generated drawings, unique in the document. */
  drawingIds: { next: number };
  /**
   * Content controls written so far. One wrapper written a second time (an
   * edit split what it wraps over two paragraphs) gives a second control,
   * which must not repeat the id of the first.
   */
  controls: Set<Wrapper>;
  splitControls: { count: number };
}

// ------------------------------------------------------------------- escaping

const UNSAFE_TEXT = /[&<>\r\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/;
const UNSAFE_ATTR = /[&<"\t\n\r\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/;
const INVALID = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/;

/**
 * Text without the characters XML cannot hold, escaped for element content
 * ("text"), for an attribute value ("attr") or not at all ("raw": the inside
 * of a comment or a CDATA section).
 */
function clean(ctx: { dropped: { count: number } }, text: string, mode: 'text' | 'attr' | 'raw'): string {
  if (!(mode === 'attr' ? UNSAFE_ATTR : mode === 'text' ? UNSAFE_TEXT : INVALID).test(text)) return text;
  const escape = mode !== 'raw';
  const attribute = mode === 'attr';
  let out = '';
  let from = 0;
  for (let i = 0; i < text.length; i++) {
    const c = text.charCodeAt(i);
    let put: string | undefined;
    if (c === 38) {
      if (escape) put = '&amp;';
    } else if (c === 60) {
      if (escape) put = '&lt;';
    } else if (c === 13) {
      if (escape) put = '&#13;';
    } else if (c === 62) {
      if (escape && !attribute) put = '&gt;';
    } else if (c === 34) {
      if (attribute) put = '&quot;';
    } else if (c === 9 || c === 10) {
      if (attribute) put = c === 9 ? '&#9;' : '&#10;';
    } else if (c < 32 || c === 0xfffe || c === 0xffff) {
      put = '';
      ctx.dropped.count++;
    } else if (c >= 0xd800 && c <= 0xdbff) {
      const next = i + 1 < text.length ? text.charCodeAt(i + 1) : 0;
      if (next >= 0xdc00 && next <= 0xdfff) i++;
      else {
        put = '';
        ctx.dropped.count++;
      }
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      put = '';
      ctx.dropped.count++;
    }
    if (put !== undefined) {
      out += text.slice(from, i) + put;
      from = i + 1;
    }
  }
  return out + text.slice(from);
}

/** Text escaped for element content or an attribute value, without the characters XML cannot hold. */
export function cleanText(ctx: { dropped: { count: number } }, text: string, attribute = false): string {
  return clean(ctx, text, attribute ? 'attr' : 'text');
}

function writeAttrs(ctx: BodyContext, attrs: Iterable<[string, string]> | undefined, out: string[]): void {
  if (!attrs) return;
  for (const [name, value] of attrs) out.push(' ', name, '="', cleanText(ctx, value, true), '"');
}

export function writeMisc(ctx: BodyContext, node: XmlComment | XmlPI, out: string[]): void {
  if (node.type === 'comment') {
    // "--" cannot be in a comment, nor "-" at its end; one read from a file has neither.
    let value = clean(ctx, node.value, 'raw');
    if (value.indexOf('--') >= 0) value = value.replace(/--+/g, '-');
    if (value.charAt(value.length - 1) === '-') value += ' ';
    out.push('<!--', value, '-->');
  } else {
    const value = clean(ctx, node.value, 'raw').replace(/\?>/g, '? >');
    out.push('<?', node.target, value ? ' ' + value : '', '?>');
  }
}

/** XML text of a kept tree. */
export function writeTree(ctx: BodyContext, el: XmlElement, out: string[], depth = 0): void {
  if (depth > MAX_TREE_DEPTH) throw new Error(`Cannot write the document: XML is nested deeper than ${MAX_TREE_DEPTH} levels`);
  out.push('<', el.name);
  writeAttrs(ctx, el.attrs, out);
  if (el.children.length === 0) {
    out.push('/>');
    return;
  }
  out.push('>');
  for (const node of el.children) {
    if (node.type === 'element') writeTree(ctx, node, out, depth + 1);
    else if (node.type === 'text') {
      if (node.cdata && node.value.indexOf(']]>') < 0 && !INVALID.test(node.value)) out.push('<![CDATA[', node.value, ']]>');
      else out.push(cleanText(ctx, node.value));
    } else writeMisc(ctx, node, out);
  }
  out.push('</', el.name, '>');
}

export function treeText(ctx: BodyContext, el: XmlElement): string {
  const out: string[] = [];
  writeTree(ctx, el, out);
  return out.join('');
}

// ------------------------------------------------------------------- wrappers

/** Adjacent items that share a wrapper: one wrapping element. */
export class Group<T> {
  children: (T | Group<T>)[] = [];
  constructor(readonly wrapper: Wrapper) {}
}

export type Node<T> = T | Group<T>;

function chainOf(wrap: Wrapper): Wrapper[] {
  const chain: Wrapper[] = [];
  for (let w: Wrapper | undefined = wrap; w; w = w.parent) {
    if (chain.length >= MAX_WRAPPERS) throw new Error(`Cannot write the document: more than ${MAX_WRAPPERS} wrappers around one item`);
    chain.push(w);
  }
  return chain.reverse();
}

/** The flat list of the model as the tree of elements it stands for. */
export function groupByWrapper<T extends { wrap?: Wrapper }>(items: readonly T[]): Node<T>[] {
  const root: Node<T>[] = [];
  const open: Group<T>[] = [];
  for (const item of items) {
    if (!item.wrap) {
      if (open.length) open.length = 0;
      root.push(item);
      continue;
    }
    // Most items follow one with the same wrapper.
    if (open.length && open[open.length - 1].wrapper === item.wrap) {
      open[open.length - 1].children.push(item);
      continue;
    }
    const chain = chainOf(item.wrap);
    let same = 0;
    while (same < open.length && same < chain.length && open[same].wrapper === chain[same]) same++;
    open.length = same;
    for (let i = same; i < chain.length; i++) {
      const group = new Group<T>(chain[i]);
      (open.length ? open[open.length - 1].children : root).push(group);
      open.push(group);
    }
    open[open.length - 1].children.push(item);
  }
  return root;
}

/** Kept children and comments of a table or row, by the wrapper they sat in. */
interface Loose {
  extras: Map<Wrapper | undefined, ExtraXml[]>;
  misc: Map<Wrapper | undefined, MiscXml[]>;
  seen: Set<Wrapper>;
}

function looseOf(extras: ExtraXml[] | undefined, misc: MiscXml[] | undefined, skip?: ExtraXml): Loose | undefined {
  if (!extras && !misc) return undefined;
  const loose: Loose = { extras: new Map(), misc: new Map(), seen: new Set() };
  for (const extra of extras || []) {
    if (extra === skip) continue;
    const list = loose.extras.get(extra.wrap);
    if (list) list.push(extra);
    else loose.extras.set(extra.wrap, [extra]);
  }
  for (const item of misc || []) {
    const list = loose.misc.get(item.wrap);
    if (list) list.push(item);
    else loose.misc.set(item.wrap, [item]);
  }
  return loose;
}

interface Level<T> {
  name: 'block' | 'inline' | 'row' | 'cell';
  /** Writes the item; true when it wrote an element (false for a comment, or nothing). */
  item(ctx: BodyContext, item: T, out: string[], depth: number): boolean;
}

interface Piece {
  text: string;
  element: boolean;
}

/**
 * Pieces of content with the kept children put back at the place they had
 * among the element children (`index`), and the comments before the element
 * they stood before.
 */
function interleave(ctx: BodyContext, pieces: Piece[], extras: ExtraXml[] | undefined, misc: MiscXml[] | undefined, out: string[]): void {
  let elements = 0;
  for (const piece of pieces) if (piece.element) elements++;
  const kept = (extras || []).filter((extra) => extra.index >= 0);
  const total = elements + kept.length;
  const slots: (XmlElement | undefined)[] = new Array(total);
  const late: XmlElement[] = [];
  for (const extra of kept) {
    if (extra.index < total && slots[extra.index] === undefined) slots[extra.index] = extra.el;
    else late.push(extra.el);
  }
  const notes = misc ? misc.slice() : [];
  let slot = 0;
  const comments = () => {
    for (let i = 0; i < notes.length; i++) {
      const note = notes[i];
      if (note.index > slot) continue;
      writeMisc(ctx, note.node, out);
      notes.splice(i--, 1);
    }
  };
  const keptUpTo = () => {
    while (slot < total && slots[slot] !== undefined) {
      comments();
      writeTree(ctx, slots[slot] as XmlElement, out);
      slot++;
    }
  };
  for (const piece of pieces) {
    if (piece.element) {
      keptUpTo();
      comments();
      slot++;
    }
    out.push(piece.text);
  }
  keptUpTo();
  for (const el of late) writeTree(ctx, el, out);
  slot = Infinity;
  comments();
}

/**
 * Writes the nodes of one container. `lead` are elements written before them
 * that count among the children (properties); `extras` and `misc` are what
 * the model kept of the container besides its content.
 */
function writeContent<T>(
  ctx: BodyContext,
  nodes: Node<T>[],
  level: Level<T>,
  out: string[],
  depth: number,
  extras: ExtraXml[] | undefined,
  misc: MiscXml[] | undefined,
  loose: Loose | undefined,
  lead?: string[],
): void {
  if (!(extras && extras.length) && !(misc && misc.length)) {
    if (lead) for (const text of lead) out.push(text);
    for (const node of nodes) writeNode(ctx, node, level, out, depth, loose);
    return;
  }
  const pieces: Piece[] = [];
  if (lead) for (const text of lead) pieces.push({ text, element: true });
  for (const node of nodes) {
    const part: string[] = [];
    const element = writeNode(ctx, node, level, part, depth, loose);
    pieces.push({ text: part.join(''), element });
  }
  interleave(ctx, pieces, extras, misc, out);
}

function writeNode<T>(ctx: BodyContext, node: Node<T>, level: Level<T>, out: string[], depth: number, loose: Loose | undefined): boolean {
  if (node instanceof Group) {
    writeGroup(ctx, node, level, out, depth, loose);
    return true;
  }
  return level.item(ctx, node, out, depth);
}

const bool = (value: boolean): string => (value ? '1' : '0');

/** The opening tag of a wrapper that holds its content directly, and its name. */
function wrapperTag(ctx: BodyContext, wrapper: Wrapper, out: string[]): string | null {
  const w = ctx.w;
  const a = ctx.a;
  let name: string;
  const attrs: [string, string][] = [];
  switch (wrapper.kind) {
    case 'hyperlink': {
      name = w + 'hyperlink';
      let rId = wrapper.rId;
      let anchor = wrapper.anchor;
      if (ctx.link) {
        const mapped = ctx.link(wrapper);
        if (mapped === null) return null;
        rId = mapped.rId;
        anchor = mapped.anchor;
      }
      if (rId !== undefined) attrs.push([ctx.r + 'id', rId]);
      if (anchor !== undefined) attrs.push([a + 'anchor', anchor]);
      if (wrapper.tooltip !== undefined) attrs.push([a + 'tooltip', wrapper.tooltip]);
      break;
    }
    case 'ins':
    case 'del':
    case 'moveFrom':
    case 'moveTo':
      name = w + wrapper.kind;
      if (wrapper.id !== undefined) attrs.push([a + 'id', String(wrapper.id)]);
      if (wrapper.author !== undefined) attrs.push([a + 'author', wrapper.author]);
      if (wrapper.date !== undefined) attrs.push([a + 'date', wrapper.date]);
      break;
    case 'simpleField':
      name = w + 'fldSimple';
      attrs.push([a + 'instr', wrapper.instruction]);
      if (wrapper.locked !== undefined) attrs.push([a + 'fldLock', bool(wrapper.locked)]);
      if (wrapper.dirty !== undefined) attrs.push([a + 'dirty', bool(wrapper.dirty)]);
      break;
    case 'dir':
    case 'bdo':
      name = w + wrapper.kind;
      if (wrapper.val !== undefined) attrs.push([a + 'val', wrapper.val]);
      break;
    case 'smartTag':
    case 'customXml':
      name = w + wrapper.kind;
      break;
    default:
      throw new Error('Cannot write the document: unknown kind of wrapper');
  }
  out.push('<', name);
  writeAttrs(ctx, attrs, out);
  writeAttrs(ctx, wrapper.attrs, out);
  out.push('>');
  return name;
}

/** Writes the element of a wrapper around its content, already written. */
function writeShell(ctx: BodyContext, wrapper: Wrapper, inner: string, out: string[]): void {
  if (wrapper.kind === 'alternate') {
    writeTree(ctx, wrapper.xml, out);
    return;
  }
  if (wrapper.kind === 'sdt') {
    const w = ctx.w;
    const self = wrapper.extras ? wrapper.extras.find((extra) => extra.index < 0) : undefined;
    out.push('<', w, 'sdt');
    if (self) writeAttrs(ctx, self.el.attrs, out);
    out.push('>');
    const pieces: Piece[] = [];
    let pr = wrapper.pr;
    if (!ctx.plain) {
      if (ctx.controls.has(wrapper)) {
        ctx.splitControls.count++;
        // Without the id: Word gives the control a new one.
        if (pr) pr = { type: 'element', name: pr.name, attrs: pr.attrs, children: pr.children.filter((node) => node.type !== 'element' || kindOf(ctx.reader, node.name) !== 'id') };
      } else ctx.controls.add(wrapper);
    }
    if (pr) pieces.push({ text: treeText(ctx, pr), element: true });
    if (wrapper.endPr) pieces.push({ text: treeText(ctx, wrapper.endPr), element: true });
    pieces.push({ text: inner ? `<${w}sdtContent>${inner}</${w}sdtContent>` : `<${w}sdtContent/>`, element: true });
    interleave(ctx, pieces, wrapper.extras, wrapper.misc, out);
    out.push('</', w, 'sdt>');
    return;
  }
  const name = wrapperTag(ctx, wrapper, out);
  if (name === null) {
    out.push(inner);
    return;
  }
  if ((wrapper.kind === 'smartTag' || wrapper.kind === 'customXml') && wrapper.pr) writeTree(ctx, wrapper.pr, out);
  out.push(inner, '</', name, '>');
}

function writeGroup<T>(ctx: BodyContext, group: Group<T>, level: Level<T>, out: string[], depth: number, loose: Loose | undefined): void {
  if (depth > MAX_WRITE_DEPTH) throw new Error(`Cannot write the document: content is nested deeper than ${MAX_WRITE_DEPTH} levels`);
  const wrapper = group.wrapper;
  if (loose) loose.seen.add(wrapper);
  if (wrapper.kind === 'alternate' && ctx.plain) {
    writeTree(ctx, wrapper.xml, out);
    return;
  }
  const inner: string[] = [];
  // Kept children of the wrapper element itself take their place among its content, except for a
  // content control, whose content sits one level down.
  const own = wrapper.kind === 'sdt' || !wrapper.extras ? undefined : wrapper.extras;
  const inside = loose ? loose.extras.get(wrapper) : undefined;
  writeContent(ctx, group.children, level, inner, depth + 1, own && inside ? own.concat(inside) : own || inside, loose ? loose.misc.get(wrapper) : undefined, loose);
  const text = inner.join('');
  if (wrapper.kind === 'alternate') {
    // The element as it was read while the branch the reader took still says what the model says.
    if (rereadAlternate(ctx, wrapper.xml, level.name) === text) writeTree(ctx, wrapper.xml, out);
    else out.push(text);
    return;
  }
  writeShell(ctx, wrapper, text, out);
}

/** What the model kept inside wrappers that hold no row or cell (and so were not written with one). */
function writeOrphans(ctx: BodyContext, loose: Loose | undefined, out: string[]): void {
  if (!loose) return;
  const wrappers = new Set<Wrapper>();
  for (const key of loose.extras.keys()) if (key && !loose.seen.has(key)) wrappers.add(key);
  for (const key of loose.misc.keys()) if (key && !loose.seen.has(key)) wrappers.add(key);
  for (const wrapper of wrappers) {
    const inner: string[] = [];
    interleave(ctx, [], loose.extras.get(wrapper), loose.misc.get(wrapper), inner);
    let text = inner.join('');
    const chain = chainOf(wrapper);
    for (let i = chain.length - 1; i >= 0; i--) {
      const shell: string[] = [];
      writeShell(ctx, chain[i], text, shell);
      text = shell.join('');
      // A wrapper written with its rows already holds the rest of the chain.
      if (i > 0 && loose.seen.has(chain[i - 1])) break;
    }
    out.push(text);
  }
}

const plainContext = (ctx: BodyContext): BodyContext => (ctx.plain ? ctx : { ...ctx, plain: true });

const synthetic = (ctx: BodyContext, local: string, children: XmlNode[]): XmlElement => ({ type: 'element', name: ctx.w + local, attrs: new Map(), children });

/** The first block the reader makes of an element; undefined when it refuses it. */
function rereadBlock(ctx: BodyContext, el: XmlElement): Block | undefined {
  const blocks: Block[] = [];
  try {
    readBlock(ctx.reader, el, blocks, undefined, 0);
  } catch (error) {
    return undefined;
  }
  return blocks[0];
}

/**
 * What the writer makes of the selected branch of an <mc:AlternateContent>
 * as the reader reads it now; undefined when it cannot be read at this level.
 */
function rereadAlternate(ctx: BodyContext, xml: XmlElement, level: 'block' | 'inline' | 'row' | 'cell'): string | undefined {
  const plain = plainContext(ctx);
  const out: string[] = [];
  const content = <T extends { wrap?: Wrapper }>(items: readonly T[], kind: Level<T>, loose?: Loose): boolean => {
    const nodes = groupByWrapper(items);
    const group = nodes[0];
    if (nodes.length !== 1 || !(group instanceof Group) || group.wrapper.kind !== 'alternate') return false;
    writeContent(plain, group.children, kind, out, 1, loose ? loose.extras.get(group.wrapper) : undefined, loose ? loose.misc.get(group.wrapper) : undefined, loose);
    return true;
  };
  let ok = false;
  try {
    if (level === 'block') {
      const blocks: Block[] = [];
      readBlock(ctx.reader, xml, blocks, undefined, 0);
      ok = content(blocks, BLOCK);
    } else if (level === 'inline') {
      const paragraph = rereadBlock(ctx, synthetic(ctx, 'p', [xml]));
      ok = !!paragraph && paragraph.type === 'paragraph' && content(paragraph.children, INLINE);
    } else {
      const table = rereadBlock(ctx, synthetic(ctx, 'tbl', [level === 'row' ? xml : synthetic(ctx, 'tr', [xml])]));
      if (table && table.type === 'table') {
        if (level === 'row') ok = content(table.rows, ROW, looseOf(table.extras, table.misc));
        else if (table.rows.length === 1) ok = content(table.rows[0].cells, CELL, looseOf(table.rows[0].extras, table.rows[0].misc));
      }
    }
  } catch (error) {
    return undefined;
  }
  return ok ? out.join('') : undefined;
}

// ----------------------------------------------------------------- run items

const isXmlSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

function writeTextItem(ctx: BodyContext, local: string, item: { text: string; source?: string }, out: string[]): void {
  const w = ctx.w;
  const text = item.text;
  if (item.source !== undefined) {
    // The text as it was written, with the outer whitespace the reader took off, as long as it still gives the text.
    const source = item.source;
    let start = 0;
    let end = source.length;
    while (start < end && isXmlSpace(source.charCodeAt(start))) start++;
    while (end > start && isXmlSpace(source.charCodeAt(end - 1))) end--;
    if (source.length - (end - start) > 0 && end - start === text.length && source.indexOf(text, start) === start) {
      out.push('<', w, local, '>', cleanText(ctx, source), '</', w, local, '>');
      return;
    }
  }
  if (text === '') {
    out.push('<', w, local, '/>');
    return;
  }
  const outer = isXmlSpace(text.charCodeAt(0)) || isXmlSpace(text.charCodeAt(text.length - 1));
  out.push('<', w, local, outer ? ' xml:space="preserve">' : '>', cleanText(ctx, text), '</', w, local, '>');
}

/** The attribute of a kept element by its WordprocessingML local name. */
function keptAttr(ctx: BodyContext, el: XmlElement, local: string): string | undefined {
  for (const name of el.attrs.keys()) if (kindOf(ctx.reader, name) === local) return name;
  return undefined;
}

function writeFieldChar(ctx: BodyContext, item: FieldCharItem, out: string[]): void {
  if (item.xml) {
    // The element as read, with what the model holds of it on top.
    const el: XmlElement = { type: 'element', name: item.xml.name, attrs: new Map(item.xml.attrs), children: item.xml.children };
    const typeName = keptAttr(ctx, el, 'fldCharType');
    if (typeName === undefined || el.attrs.get(typeName) !== item.charType) el.attrs.set(typeName || ctx.a + 'fldCharType', item.charType);
    const flag = (local: string, value: boolean | undefined) => {
      const name = keptAttr(ctx, el, local);
      const now = name === undefined ? undefined : parseOnOff(el.attrs.get(name));
      if (value === undefined) {
        // A value that is no boolean is not in the model and stays; a boolean the model no longer has goes.
        if (name !== undefined && now !== undefined) el.attrs.delete(name);
      } else if (now !== value) el.attrs.set(name || ctx.a + local, bool(value));
    };
    flag('dirty', item.dirty);
    flag('fldLock', item.locked);
    writeTree(ctx, el, out);
    return;
  }
  const w = ctx.w;
  const a = ctx.a;
  out.push('<', w, 'fldChar ', a, 'fldCharType="', item.charType, '"');
  if (item.dirty !== undefined) out.push(' ', a, 'dirty="', bool(item.dirty), '"');
  if (item.locked !== undefined) out.push(' ', a, 'fldLock="', bool(item.locked), '"');
  if (item.data && item.data.length) {
    out.push('>');
    for (const el of item.data) writeTree(ctx, el, out);
    out.push('</', w, 'fldChar>');
  } else out.push('/>');
}

const SIMPLE_NAMES: { [type: string]: string } = {
  tab: 'tab',
  noBreakHyphen: 'noBreakHyphen',
  softHyphen: 'softHyphen',
  footnoteRef: 'footnoteRef',
  endnoteRef: 'endnoteRef',
  annotationRef: 'annotationRef',
  separator: 'separator',
  continuationSeparator: 'continuationSeparator',
  pageNumber: 'pgNum',
  lastRenderedPageBreak: 'lastRenderedPageBreak',
};

function writeRunItem(ctx: BodyContext, item: RunItem, out: string[], depth: number): void {
  const w = ctx.w;
  const a = ctx.a;
  switch (item.type) {
    case 'text':
      writeTextItem(ctx, 't', item, out);
      return;
    case 'deletedText':
      writeTextItem(ctx, 'delText', item, out);
      return;
    case 'instrText':
      writeTextItem(ctx, item.deleted ? 'delInstrText' : 'instrText', item, out);
      return;
    case 'break':
      if (item.kind === 'page') out.push('<', w, 'br ', a, 'type="page"/>');
      else if (item.kind === 'column') out.push('<', w, 'br ', a, 'type="column"/>');
      else if (item.clear !== undefined) out.push('<', w, 'br ', a, 'type="textWrapping" ', a, 'clear="', cleanText(ctx, item.clear, true), '"/>');
      else out.push('<', w, item.cr ? 'cr/>' : 'br/>');
      return;
    case 'symbol': {
      out.push('<', w, 'sym');
      if (item.font !== undefined) out.push(' ', a, 'font="', cleanText(ctx, item.font, true), '"');
      out.push(' ', a, 'char="', ('000' + (item.char.charCodeAt(0) || 0).toString(16).toUpperCase()).slice(-4), '"/>');
      return;
    }
    case 'fieldChar':
      writeFieldChar(ctx, item, out);
      return;
    case 'footnoteReference':
    case 'endnoteReference':
      out.push('<', w, item.type);
      if (item.customMark !== undefined) out.push(' ', a, 'customMarkFollows="', bool(item.customMark), '"');
      out.push(' ', a, 'id="', String(item.id), '"/>');
      return;
    case 'commentReference':
      out.push('<', w, 'commentReference ', a, 'id="', String(item.id), '"/>');
      return;
    case 'positionalTab':
      out.push('<', w, 'ptab');
      if (item.relativeTo !== undefined) out.push(' ', a, 'relativeTo="', cleanText(ctx, item.relativeTo, true), '"');
      if (item.alignment !== undefined) out.push(' ', a, 'alignment="', cleanText(ctx, item.alignment, true), '"');
      if (item.leader !== undefined) out.push(' ', a, 'leader="', cleanText(ctx, item.leader, true), '"');
      out.push('/>');
      return;
    case 'drawing':
      writeDrawing(ctx, item, out, depth);
      return;
    case 'pict':
    case 'object':
      writeTree(ctx, ctx.plain ? item.xml : syncTextBoxes(ctx, item.xml, item.textBoxes, depth), out);
      return;
    case 'alternate': {
      if (ctx.plain) {
        writeTree(ctx, item.xml, out);
        return;
      }
      if (depth > MAX_WRITE_DEPTH) throw new Error(`Cannot write the document: content is nested deeper than ${MAX_WRITE_DEPTH} levels`);
      const inner: string[] = [];
      for (const sub of item.items) writeRunItem(ctx, sub, inner, depth + 1);
      const text = inner.join('');
      // As for the wrapper of the same name: the element as read while its branch is unchanged.
      let was: string | undefined;
      const paragraph = rereadBlock(ctx, synthetic(ctx, 'p', [synthetic(ctx, 'r', [item.xml])]));
      const run = paragraph && paragraph.type === 'paragraph' ? paragraph.children[0] : undefined;
      const again = run && run.type === 'run' ? run.content[0] : undefined;
      if (again && again.type === 'alternate') {
        const plain = plainContext(ctx);
        const parts: string[] = [];
        for (const sub of again.items) writeRunItem(plain, sub, parts, depth + 1);
        was = parts.join('');
      }
      if (was === text) writeTree(ctx, item.xml, out);
      else out.push(text);
      return;
    }
    case 'misc':
      writeMisc(ctx, item.node, out);
      return;
    case 'unknown':
      writeTree(ctx, item.xml, out);
      return;
  }
  const name = SIMPLE_NAMES[item.type];
  if (name === undefined) throw new Error('Cannot write the document: unknown kind of run content');
  out.push('<', w, name, '/>');
}

// ------------------------------------------------------------------ drawings

const localOf = (name: string): string => name.slice(name.indexOf(':') + 1);

function childNamed(el: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const node of el.children) if (node.type === 'element' && localOf(node.name) === local) return node;
  return undefined;
}

/** A copy of the tree with one element replaced; the tree itself when the element is not in it. */
function replaceNode(root: XmlElement, target: XmlElement, replacement: XmlElement, depth = 0): XmlElement {
  if (root === target) return replacement;
  if (depth > MAX_TREE_DEPTH) return root;
  for (let i = 0; i < root.children.length; i++) {
    const node = root.children[i];
    if (node.type !== 'element') continue;
    const next = replaceNode(node, target, replacement, depth + 1);
    if (next === node) continue;
    const children = root.children.slice();
    children[i] = next;
    return { type: 'element', name: root.name, attrs: root.attrs, children };
  }
  return root;
}

const withAttrs = (el: XmlElement, pairs: [string, string][]): XmlElement => {
  const attrs = new Map(el.attrs);
  for (const [name, value] of pairs) attrs.set(name, value);
  return { type: 'element', name: el.name, attrs, children: el.children };
};

const EMU_TEXT = /^\s*-?\d{1,10}\s*$/;

/** The tree of a drawing with the size of the model, when the two differ. */
function syncSize(item: DrawingItem, xml: XmlElement): XmlElement {
  const frame = childNamed(xml, 'inline') || childNamed(xml, 'anchor');
  const extent = childNamed(frame, 'extent');
  if (!extent || item.width === undefined || item.height === undefined) return xml;
  const cx = extent.attrs.get('cx');
  const cy = extent.attrs.get('cy');
  const same = (text: string | undefined, value: number) => text !== undefined && EMU_TEXT.test(text) && Number(text) === value;
  if (same(cx, item.width) && same(cy, item.height)) return xml;
  const size: [string, string][] = [
    ['cx', String(item.width)],
    ['cy', String(item.height)],
  ];
  let out = replaceNode(xml, extent, withAttrs(extent, size));
  // A picture states its size once more, on the picture itself.
  const picture = childNamed(childNamed(childNamed(frame, 'graphic'), 'graphicData'), 'pic');
  const inner = childNamed(childNamed(childNamed(picture, 'spPr'), 'xfrm'), 'ext');
  if (inner) out = replaceNode(out, inner, withAttrs(inner, size));
  return out;
}

/** What the writer makes of the blocks the reader reads from a <w:txbxContent> now. */
function rereadTextBox(ctx: BodyContext, content: XmlElement): string | undefined {
  const table = rereadBlock(ctx, synthetic(ctx, 'tbl', [synthetic(ctx, 'tr', [synthetic(ctx, 'tc', content.children)])]));
  if (!table || table.type !== 'table' || table.rows.length !== 1 || table.rows[0].cells.length !== 1) return undefined;
  const out: string[] = [];
  try {
    writeBlocks(plainContext(ctx), table.rows[0].cells[0].blocks, out, 1);
  } catch (error) {
    return undefined;
  }
  return out.join('');
}

/** The kept tree of a drawing or VML picture with the text boxes of the model, where those were edited. */
function syncTextBoxes(ctx: BodyContext, xml: XmlElement, boxes: Block[][] | undefined, depth: number): XmlElement {
  if (!boxes || !boxes.length) return xml;
  let contents: XmlElement[];
  try {
    contents = findSelected(ctx.reader, xml, 'txbxContent');
  } catch (error) {
    return xml;
  }
  if (contents.length !== boxes.length) return xml;
  let out = xml;
  for (let i = 0; i < boxes.length; i++) {
    const parts: string[] = [];
    writeBlocks(ctx, boxes[i], parts, depth + 1);
    const text = parts.join('');
    const was = rereadTextBox(ctx, contents[i]);
    if (was === undefined || was === text) continue;
    // Namespace prefixes are not resolved by the parser, so the text parses without the declarations of the part.
    const parsed = parseXml(`<x>${text}</x>`).root.children;
    out = replaceNode(out, contents[i], { type: 'element', name: contents[i].name, attrs: contents[i].attrs, children: parsed });
  }
  return out;
}

interface DrawingNames {
  wp: string;
  a: string;
  pic: string;
  r: string;
}

const TRANSITIONAL_DRAWING: DrawingNames = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

const STRICT_DRAWING: DrawingNames = {
  wp: 'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
  a: 'http://purl.oclc.org/ooxml/drawingml/main',
  pic: 'http://purl.oclc.org/ooxml/drawingml/picture',
  r: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
};

const NS_WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape';
const NS_WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup';

const WRAP_ELEMENTS: { [wrap: string]: string } = {
  none: 'wrapNone',
  square: 'wrapSquare',
  tight: 'wrapTight',
  through: 'wrapThrough',
  topAndBottom: 'wrapTopAndBottom',
};

const emuText = (value: number | undefined, fallback = 0): string => String(value === undefined || !isFinite(value) ? fallback : Math.round(value));

function cropText(crop: ImageCrop | undefined): string {
  if (!crop) return '';
  let out = '<a:srcRect';
  if (crop.left !== undefined) out += ` l="${emuText(crop.left)}"`;
  if (crop.top !== undefined) out += ` t="${emuText(crop.top)}"`;
  if (crop.right !== undefined) out += ` r="${emuText(crop.right)}"`;
  if (crop.bottom !== undefined) out += ` b="${emuText(crop.bottom)}"`;
  return out + '/>';
}

function blipText(ctx: BodyContext, ns: DrawingNames, image: ImageRef | undefined): string {
  const rId = image ? (ctx.image ? ctx.image(image) : image.rId) : undefined;
  if (rId === undefined) return '<a:blip/>';
  return `<a:blip xmlns:r="${ns.r}" r:${image && image.linked ? 'link' : 'embed'}="${cleanText(ctx, rId, true)}"/>`;
}

function transformText(item: DrawingItem, width: string, height: string, group: boolean): string {
  const t = item.transform;
  let out = '<a:xfrm';
  if (t && t.rotation) out += ` rot="${emuText(t.rotation)}"`;
  if (t && t.flipH) out += ' flipH="1"';
  if (t && t.flipV) out += ' flipV="1"';
  out += `><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/>`;
  if (group) out += `<a:chOff x="0" y="0"/><a:chExt cx="${width}" cy="${height}"/>`;
  return out + '</a:xfrm>';
}

/**
 * A standard DrawingML drawing for an item that has no tree of its own: a
 * picture, a text box, or a group of them. What the model does not know of a
 * shape (its outline and fill, the place of each part of a group) is not
 * invented: text boxes are written without outline and fill, the parts of a
 * group over its whole area.
 */
function generateDrawing(ctx: BodyContext, item: DrawingItem, out: string[], depth: number): void {
  if (depth > MAX_WRITE_DEPTH) throw new Error(`Cannot write the document: content is nested deeper than ${MAX_WRITE_DEPTH} levels`);
  const ns = ctx.strict ? STRICT_DRAWING : TRANSITIONAL_DRAWING;
  const w = ctx.w;
  const id = ctx.drawingIds.next++;
  const width = emuText(item.width);
  const height = emuText(item.height);
  const name = cleanText(ctx, item.name !== undefined ? item.name : `Drawing ${id}`, true);
  const descr = item.description !== undefined ? ` descr="${cleanText(ctx, item.description, true)}"` : '';
  const anchored = item.placement === 'anchor';

  const pictures = item.images && item.images.length ? item.images : item.image ? [{ image: item.image, crop: item.crop }] : [];
  const boxes = item.textBoxes || [];
  const insets = childNamed(item.xml, 'bodyPr');
  let bodyPr = '<wps:bodyPr';
  if (insets) for (const side of ['lIns', 'tIns', 'rIns', 'bIns']) if (/^\d{1,9}$/.test(insets.attrs.get(side) || '')) bodyPr += ` ${side}="${insets.attrs.get(side)}"`;
  bodyPr += '/>';

  const picture = (image: ImageRef | undefined, crop: ImageCrop | undefined, outer: boolean): string =>
    `<pic:pic xmlns:pic="${ns.pic}"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"${descr}/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill>${blipText(ctx, ns, image)}${cropText(crop)}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr>${outer ? transformText(item, width, height, false) : `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`}` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>';
  const shape = (blocks: Block[] | undefined, fill: ImageRef | undefined, outer: boolean): string => {
    let text = `<wps:wsp xmlns:wps="${NS_WPS}"><wps:cNvSpPr${blocks ? ' txBox="1"' : ''}/><wps:spPr>`;
    text += outer ? transformText(item, width, height, false) : `<a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`;
    text += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>';
    text += fill ? `<a:blipFill>${blipText(ctx, ns, fill)}<a:stretch><a:fillRect/></a:stretch></a:blipFill>` : '<a:noFill/>';
    text += '<a:ln><a:noFill/></a:ln></wps:spPr>';
    if (blocks) {
      const inner: string[] = [];
      writeBlocks(ctx, blocks, inner, depth + 1);
      text += `<wps:txbx><${w}txbxContent>${inner.length ? inner.join('') : `<${w}p/>`}</${w}txbxContent></wps:txbx>`;
    }
    return text + bodyPr + '</wps:wsp>';
  };

  let uri: string;
  let graphic: string;
  const parts = pictures.length + boxes.length;
  if (item.kind === 'group' || parts > 2 || (parts === 2 && !(pictures.length === 1 && boxes.length === 1))) {
    uri = NS_WPG;
    graphic = `<wpg:wgp xmlns:wpg="${NS_WPG}"><wpg:cNvGrpSpPr/><wpg:grpSpPr>${transformText(item, width, height, true)}</wpg:grpSpPr>`;
    for (const part of pictures) graphic += picture(part.image, part.crop, false);
    for (const blocks of boxes) graphic += shape(blocks, undefined, false);
    graphic += '</wpg:wgp>';
  } else if (boxes.length) {
    uri = NS_WPS;
    graphic = shape(boxes[0], pictures.length ? pictures[0].image : undefined, true);
  } else if (pictures.length || item.kind === 'picture') {
    uri = ns.pic;
    graphic = picture(pictures.length ? pictures[0].image : undefined, pictures.length ? pictures[0].crop || item.crop : item.crop, true);
  } else {
    uri = NS_WPS;
    graphic = shape(undefined, undefined, true);
  }

  const effect = item.effectExtent;
  const effectText = effect
    ? `<wp:effectExtent l="${emuText(effect.left)}" t="${emuText(effect.top)}" r="${emuText(effect.right)}" b="${emuText(effect.bottom)}"/>`
    : '';
  const tail =
    `<wp:docPr id="${id}" name="${name}"${descr}/><wp:cNvGraphicFramePr/>` +
    `<a:graphic xmlns:a="${ns.a}"><a:graphicData uri="${uri}">${graphic}</a:graphicData></a:graphic>`;
  const distances = (from: { distT?: number; distB?: number; distL?: number; distR?: number } | undefined): string => {
    let text = '';
    if (from) for (const side of ['distT', 'distB', 'distL', 'distR'] as const) if (from[side] !== undefined) text += ` ${side}="${emuText(from[side])}"`;
    return text;
  };

  out.push('<', w, 'drawing>');
  if (!anchored) {
    out.push(`<wp:inline xmlns:wp="${ns.wp}"${distances(item.inline)}><wp:extent cx="${width}" cy="${height}"/>`, effectText, tail, '</wp:inline>');
  } else {
    const anchor: DrawingAnchor = item.anchor || {};
    const position = (local: string, place: DrawingAnchor['horizontal'], fallback: string): string => {
      const from = cleanText(ctx, (place && place.relativeFrom) || fallback, true);
      const inner = place && place.align !== undefined ? `<wp:align>${cleanText(ctx, place.align)}</wp:align>` : `<wp:posOffset>${emuText(place ? place.offset : 0)}</wp:posOffset>`;
      return `<wp:${local} relativeFrom="${from}">${inner}</wp:${local}>`;
    };
    const wrapName = WRAP_ELEMENTS[anchor.wrap || 'none'] || 'wrapNone';
    let wrap = `<wp:${wrapName}`;
    const sided = wrapName === 'wrapSquare' || wrapName === 'wrapTight' || wrapName === 'wrapThrough';
    if (sided) wrap += ` wrapText="${cleanText(ctx, anchor.wrapText || 'bothSides', true)}"`;
    if (anchor.wrapDist) {
      const d = anchor.wrapDist;
      if (d.top !== undefined && wrapName !== 'wrapTight' && wrapName !== 'wrapThrough') wrap += ` distT="${emuText(d.top)}"`;
      if (d.bottom !== undefined && wrapName !== 'wrapTight' && wrapName !== 'wrapThrough') wrap += ` distB="${emuText(d.bottom)}"`;
      if (d.left !== undefined && wrapName !== 'wrapTopAndBottom') wrap += ` distL="${emuText(d.left)}"`;
      if (d.right !== undefined && wrapName !== 'wrapTopAndBottom') wrap += ` distR="${emuText(d.right)}"`;
    }
    if (wrapName === 'wrapTight' || wrapName === 'wrapThrough') {
      // The outline is required; without one in the model it is the rectangle of the drawing.
      const polygon = anchor.wrapPolygon && anchor.wrapPolygon.points.length >= 2 ? anchor.wrapPolygon.points : ([[0, 0], [0, 21600], [21600, 21600], [21600, 0], [0, 0]] as [number, number][]);
      wrap += `><wp:wrapPolygon edited="${anchor.wrapPolygon && anchor.wrapPolygon.edited ? '1' : '0'}">`;
      polygon.forEach(([x, y], i) => {
        wrap += `<wp:${i === 0 ? 'start' : 'lineTo'} x="${emuText(x)}" y="${emuText(y)}"/>`;
      });
      wrap += `</wp:wrapPolygon></wp:${wrapName}>`;
    } else wrap += '/>';
    const simple = anchor.simplePos;
    out.push(
      `<wp:anchor xmlns:wp="${ns.wp}"${distances(anchor)} simplePos="${simple ? '1' : '0'}"`,
      ` relativeHeight="${emuText(anchor.relativeHeight, 251658240 + id)}" behindDoc="${anchor.behindDoc ? '1' : '0'}" locked="${anchor.locked ? '1' : '0'}"`,
      ` layoutInCell="${anchor.layoutInCell === false ? '0' : '1'}" allowOverlap="${anchor.allowOverlap === false ? '0' : '1'}">`,
      `<wp:simplePos x="${emuText(simple ? simple.x : 0)}" y="${emuText(simple ? simple.y : 0)}"/>`,
      position('positionH', anchor.horizontal, 'column'),
      position('positionV', anchor.vertical, 'paragraph'),
      `<wp:extent cx="${width}" cy="${height}"/>`,
      effectText,
      wrap,
      tail,
      '</wp:anchor>',
    );
  }
  out.push('</', w, 'drawing>');
}

/** True when the tree is a <w:drawing> with a frame, which is written as it is. */
export function isDrawingTree(ctx: BodyContext, xml: XmlElement | undefined): boolean {
  return !!xml && kindOf(ctx.reader, xml.name) === 'drawing' && !!(childNamed(xml, 'inline') || childNamed(xml, 'anchor'));
}

function writeDrawing(ctx: BodyContext, item: DrawingItem, out: string[], depth: number): void {
  if (!isDrawingTree(ctx, item.xml)) {
    generateDrawing(ctx, item, out, depth);
    return;
  }
  let xml = item.xml;
  if (!ctx.plain) xml = syncTextBoxes(ctx, syncSize(item, xml), item.textBoxes, depth);
  writeTree(ctx, xml, out);
}

// ------------------------------------------------------------ inline content

function runPropsText(ctx: BodyContext, props: RunProps): string {
  let text = ctx.runCache.get(props);
  if (text === undefined) {
    const el = runPropsElement(ctx, props);
    text = el ? treeText(ctx, el) : '';
    ctx.runCache.set(props, text);
  }
  return text;
}

function paraPropsText(ctx: BodyContext, props: ParaProps): string {
  // Section properties name headers by relationship and may differ between two writes of the same bag.
  if (props.section) {
    const el = paraPropsElement(ctx, props);
    return el ? treeText(ctx, el) : '';
  }
  let text = ctx.paraCache.get(props);
  if (text === undefined) {
    const el = paraPropsElement(ctx, props);
    text = el ? treeText(ctx, el) : '';
    ctx.paraCache.set(props, text);
  }
  return text;
}

function writeRun(ctx: BodyContext, run: Run, out: string[], depth: number): void {
  const w = ctx.w;
  out.push('<', w, 'r');
  writeAttrs(ctx, run.attrs, out);
  const props = runPropsText(ctx, run.props);
  if (!props && !run.content.length) {
    out.push('/>');
    return;
  }
  out.push('>', props);
  for (const item of run.content) writeRunItem(ctx, item, out, depth);
  out.push('</', w, 'r>');
}

function writeRangeMark(ctx: BodyContext, mark: RangeMark, out: string[]): void {
  out.push('<', ctx.w, mark.type);
  writeAttrs(ctx, mark.attrs, out);
  out.push(' ', ctx.a, 'id="', cleanText(ctx, mark.id, true), '"');
  if (mark.name !== undefined) out.push(' ', ctx.a, 'name="', cleanText(ctx, mark.name, true), '"');
  out.push('/>');
}

const INLINE: Level<Inline> = {
  name: 'inline',
  item(ctx, item, out, depth) {
    switch (item.type) {
      case 'run':
        writeRun(ctx, item, out, depth);
        return true;
      case 'math':
      case 'unknown':
        writeTree(ctx, item.xml, out);
        return true;
      case 'proofError':
        out.push('<', ctx.w, 'proofErr ', ctx.a, 'type="', cleanText(ctx, item.kind, true), '"/>');
        return true;
      case 'empty':
        return false;
      case 'misc':
        writeMisc(ctx, item.node, out);
        return false;
    }
    writeRangeMark(ctx, item, out);
    return true;
  },
};

function writeParagraph(ctx: BodyContext, paragraph: Paragraph, out: string[], depth: number): void {
  const w = ctx.w;
  out.push('<', w, 'p');
  writeAttrs(ctx, paragraph.attrs, out);
  const props = paraPropsText(ctx, paragraph.props);
  const children = paragraph.children;
  if (!props && !children.length) {
    out.push('/>');
    return;
  }
  out.push('>', props);
  // Most paragraphs hold plain runs: no grouping needed.
  let wrapped = false;
  for (let i = 0; i < children.length; i++) {
    if (children[i].wrap) {
      wrapped = true;
      break;
    }
  }
  if (!wrapped) for (const inline of children) INLINE.item(ctx, inline, out, depth);
  else writeContent(ctx, groupByWrapper(children), INLINE, out, depth, undefined, undefined, undefined);
  out.push('</', w, 'p>');
}

// --------------------------------------------------------------------- tables

function gridElement(ctx: BodyContext, grid: number[]): XmlElement {
  return wEl(
    ctx,
    'tblGrid',
    undefined,
    grid.map((width) => wEl(ctx, 'gridCol', [['w', width]])),
  );
}

function gridText(ctx: BodyContext, table: Table, kept: ExtraXml | undefined): string {
  const wanted = gridElement(ctx, table.grid);
  if (!kept) return treeText(ctx, wanted);
  // The grid as it was read (with its change record, say) while it still says what the model says.
  const again = rereadBlock(ctx, synthetic(ctx, 'tbl', [kept.el]));
  const understood = again && again.type === 'table' ? gridElement(ctx, again.grid) : undefined;
  if (understood && treeText(ctx, understood) === treeText(ctx, wanted)) return treeText(ctx, kept.el);
  return treeText(ctx, mergeKept(kept.el, understood, wanted));
}

const CELL: Level<TableCell> = {
  name: 'cell',
  item(ctx, cell, out, depth) {
    const w = ctx.w;
    out.push('<', w, 'tc');
    writeAttrs(ctx, cell.attrs, out);
    out.push('>');
    const props = cellPropsElement(ctx, cell.props);
    if (props) writeTree(ctx, props, out);
    writeBlocks(ctx, cell.blocks, out, depth + 1);
    out.push('</', w, 'tc>');
    return true;
  },
};

const ROW: Level<TableRow> = {
  name: 'row',
  item(ctx, row, out, depth) {
    const w = ctx.w;
    out.push('<', w, 'tr');
    writeAttrs(ctx, row.attrs, out);
    out.push('>');
    const lead: string[] = [];
    if (row.tableProps) {
      const exceptions = tablePropsElement(ctx, row.tableProps, 'tblPrEx');
      if (exceptions) lead.push(treeText(ctx, exceptions));
    }
    const props = rowPropsElement(ctx, row.props);
    if (props) lead.push(treeText(ctx, props));
    const loose = looseOf(row.extras, row.misc);
    writeContent(ctx, groupByWrapper(row.cells), CELL, out, depth, loose ? loose.extras.get(undefined) : undefined, loose ? loose.misc.get(undefined) : undefined, loose, lead);
    writeOrphans(ctx, loose, out);
    out.push('</', w, 'tr>');
    return true;
  },
};

function writeTable(ctx: BodyContext, table: Table, out: string[], depth: number): void {
  if (depth > MAX_WRITE_DEPTH) throw new Error(`Cannot write the document: content is nested deeper than ${MAX_WRITE_DEPTH} levels`);
  const w = ctx.w;
  out.push('<', w, 'tbl');
  writeAttrs(ctx, table.attrs, out);
  out.push('>');
  const props = tablePropsElement(ctx, table.props);
  const keptGrid = table.extras ? table.extras.find((extra) => extra.partial === true && !extra.wrap && kindOf(ctx.reader, extra.el.name) === 'tblGrid') : undefined;
  // Both are required by the schema.
  const lead = [props ? treeText(ctx, props) : `<${w}tblPr/>`, gridText(ctx, table, keptGrid)];
  const loose = looseOf(table.extras, table.misc, keptGrid);
  writeContent(ctx, groupByWrapper(table.rows), ROW, out, depth + 1, loose ? loose.extras.get(undefined) : undefined, loose ? loose.misc.get(undefined) : undefined, loose, lead);
  writeOrphans(ctx, loose, out);
  out.push('</', w, 'tbl>');
}

const BLOCK: Level<Block> = {
  name: 'block',
  item(ctx, block, out, depth) {
    switch (block.type) {
      case 'paragraph':
        writeParagraph(ctx, block, out, depth);
        return true;
      case 'table':
        writeTable(ctx, block, out, depth);
        return true;
      case 'mark':
        writeRangeMark(ctx, block.mark, out);
        return true;
      case 'unknown':
        writeTree(ctx, block.xml, out);
        return true;
      case 'misc':
        writeMisc(ctx, block.node, out);
        return false;
    }
    return false;
  },
};

/** Writes one node of a grouped list of blocks (see groupByWrapper): a block, or a wrapper with what it wraps. */
export function writeBlockNode(ctx: BodyContext, node: Node<Block>, out: string[], depth: number): void {
  writeNode(ctx, node, BLOCK, out, depth, undefined);
}

/** Writes blocks from the model. */
export function writeBlocks(ctx: BodyContext, blocks: readonly Block[], out: string[], depth: number): void {
  let wrapped = false;
  for (let i = 0; i < blocks.length; i++) {
    if (blocks[i].wrap) {
      wrapped = true;
      break;
    }
  }
  if (!wrapped) for (const block of blocks) BLOCK.item(ctx, block, out, depth);
  else for (const node of groupByWrapper(blocks)) writeNode(ctx, node, BLOCK, out, depth, undefined);
}
