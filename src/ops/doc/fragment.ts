// A part of a document as a copy of its own (what the clipboard of the editor
// holds), and the command that puts such a copy into a document.
//
// fragmentOf() takes what a deletion of the same range would remove
// (deleteSpan, removePieces): the copy and the deletion are made by the same
// rules, so a cut that is pasted back is a move. A fragment holds new blocks;
// what it shares with the document it was taken from (runs items, property
// bags, kept XML trees, the bytes of pictures) is never changed in place by
// any command, so later edits of that document do not reach it.
//
// What a fragment carries: text in its formatting, tabs, breaks of lines,
// pages and columns, soft and non-breaking hyphens, symbols; pictures and
// drawings without text whose kept XML refers to nothing but images;
// hyperlinks; simple fields, and complex fields that are in the range from
// their beginning to their end; paragraphs with their properties, style and
// list; tables, whole or the rows and cells of the range, with their grid,
// properties, merges and everything in their cells.
//
// What it does not carry, because it cannot be put into another place without
// more than the fragment knows: references to notes and comments, tracked
// changes, content controls, equations, text boxes and drawings with text,
// charts, diagrams and embedded objects, VML pictures, alternate content,
// fields that go on over several paragraphs, smart tags and custom XML. Text
// inside such wrappers is carried without them. The caller (cutLoses of the
// editor) refuses a cut that would remove something the fragment lacks.
//
// Tracked changes, in full: what is pasted is plain content, in the document
// it came from as in any other. A copy would otherwise bring revisions under
// ids, authors and dates of its own, and a second paste would bring them twice.
//   - Text inside a tracked insertion or move is carried as text; text of a
//     tracked deletion is not carried (it is not part of what the document
//     says).
//   - A paragraph mark that is a tracked insertion, deletion or move is a
//     plain mark in the copy. For a deleted mark that means the two
//     paragraphs it would have joined stay two.
//   - The records of changed formatting are left out: of paragraphs, runs and
//     paragraph marks, sections, tables, rows and cells (w:pPrChange,
//     w:rPrChange, w:sectPrChange, w:tblPrChange, w:tblPrExChange,
//     w:trPrChange, w:tcPrChange, w:tblGridChange, w:numberingChange), and so
//     are the marks of inserted, deleted and merged rows and cells. The
//     formatting itself is carried as it is now.
//   - A cut is a copy: a paragraph that is cut and pasted back has lost these
//     records. (UNSURE: Word keeps them when tracking is off.)
//
// insertFragment() makes the fragment fit the document it goes into:
//   - relationships of pictures and links are found in the part of the story
//     or added to it, and the ids in the kept XML follow;
//   - a picture the document does not have is added (DocModel.media) under a
//     name that is free; one it has is used again;
//   - drawings get ids no other drawing of the document has;
//   - in the document the fragment came from, styles and lists keep their
//     ids. In another document a style of the same id and a list are the
//     document's own where it has the style ("use the styles of the
//     destination", as Word does by default) and a copy of the definition is
//     added where it has not; a list always comes as a new list;
//   - kept XML that uses a namespace prefix the part of the story does not
//     declare the same way gets the declaration itself.
// All of it is one edit: undoing it takes the pictures, relationships, styles
// and lists it added out again.
//
// UNSURE, where Word was not compared or is known to do otherwise:
//   - Word pastes cells over the cells at the caret (or adds rows); here a
//     table pasted in a cell is always a table inside that cell.
//   - Word lets the user choose between the formatting of the source and of
//     the destination, and may continue a list that is at the caret; here a
//     style of the same id is the destination's and a list is a new list.
//   - A section break pasted into another document keeps its page settings
//     and loses the headers and footers it named; Word copies those too.
//   - A namespace kept XML declares for itself is not added to mc:Ignorable
//     of the part. Word reads its own extension namespaces (w14, wp14, ...)
//     without that; a strict validator may not.
//   - A merged cell whose first row is not part of a copy of rows begins the
//     merge in the copy.

import {
  EMPTY_PARA_PROPS,
  paragraphText,
  partNamespaces,
  setPartNamespaces,
  type AbstractNum,
  type AttrList,
  type Block,
  type DocModel,
  type DocRelationship,
  type DocStyle,
  type DrawingItem,
  type ExtraXml,
  type ImageRef,
  type Inline,
  type NumInstance,
  type Paragraph,
  type ParaProps,
  type RowProps,
  type Run,
  type RunItem,
  type RunProps,
  type Table,
  type TableCell,
  type TableProps,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { buildEdit, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { caretIndex, fieldPairs, flatten, isAtom, isStructural, looseMark, paragraphWith, rebuild, removedPieces, sameProps, startIndex, wrapBetween, type Piece } from './pieces';
import {
  asRange,
  comparePositions,
  containerBlocks,
  inTextBox,
  itemLength,
  orderRange,
  paragraphIn,
  resolvePosition,
  samePosition,
  storyOf,
  type DocPosition,
  type DocRange,
} from './position';
import { freeRelId, rowSpans, takenByDocument, type Taken } from './structure';
import { cellsBetween, deleteSpan, joinedProps, neighbour, parting, withoutSection } from './text';
import { noteForeignIds } from './track';

const bad = (message: string): never => {
  throw new RangeError(message);
};

/** A picture a fragment refers to. The bytes are missing while they are still being read from the package of the document. */
export interface FragmentMedia {
  /** Media type ("image/png"). */
  type: string;
  bytes?: Uint8Array;
}

export interface DocFragment {
  /** Paragraphs and tables, new objects all of them. */
  blocks: Block[];
  /** The first block is a paragraph that continues the paragraph it is pasted into. */
  openStart: boolean;
  /** The last block is a paragraph without its end: what follows the caret continues it. */
  openEnd: boolean;
  /** Stands for the document the fragment was taken from (see originOf); the document itself is not held. */
  origin: object;
  /** The part of the story it was taken from, with the namespaces its root declares and its kind of document. */
  part: string;
  rootAttrs: AttrList;
  strict: boolean;
  /** The relationships of that part the fragment refers to, by id. */
  relationships: Map<string, DocRelationship>;
  /** The pictures it refers to, by the name of their part. */
  media: Map<string, FragmentMedia>;
  /** The definitions of the styles and lists it uses, and of those these depend on. */
  styles: Map<string, DocStyle>;
  abstracts: Map<number, AbstractNum>;
  nums: Map<number, NumInstance>;
  /** The namespaces the roots of the styles part and of the numbering part declare: what the prefixes in the kept trees of the definitions stand for. */
  styleNamespaces: AttrList;
  numberingNamespaces: AttrList;
  /**
   * For a range from one paragraph into another: the properties of the
   * paragraph it ends in, whose end is not part of the copy, and the
   * properties of the one paragraph a deletion of the range makes of the two.
   * A deletion keeps the formatting of one of the two paragraphs only; a cut
   * that is pasted back gives the other its formatting again (see
   * InsertFragmentOptions.moved).
   */
  endProps?: ParaProps;
  joinedProps?: ParaProps;
  /** Length of the text it holds, and how deep its tables nest (0: no table). */
  textLength: number;
  tableDepth: number;
  /**
   * The copy names more styles or lists than it holds the definitions of (there is a limit).
   * It can go back into the document it came from, which has them; in another document its
   * ids would name whatever that document has under them, and it is refused there.
   */
  definitionsCut?: boolean;
}

/** A copy larger than this is not made (the clipboard then holds plain text only). */
export const FRAGMENT_LIMITS = Object.freeze({ paragraphs: 5000, items: 20000, cells: 20000 });
/** Deepest nesting of tables a paste may leave. */
export const MAX_TABLE_DEPTH = 16;
const MAX_DEFINITIONS = 256;
const MAX_TREE_DEPTH = 300;

const ORIGINS = new WeakMap<DocModel, object>();

/** What stands for a document in the fragments taken from it: the same object for the same document, and nothing that keeps the document alive. */
export function originOf(doc: DocModel): object {
  let origin = ORIGINS.get(doc);
  if (!origin) ORIGINS.set(doc, (origin = {}));
  return origin;
}

const REL_NAMESPACES = ['http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'];
const isRelNamespace = (uri: string): boolean => REL_NAMESPACES.indexOf(uri) >= 0;
const isWordNamespace = (uri: string): boolean => /wordprocessingml\/(2006\/)?main$/.test(uri);

const localOf = (name: string): string => name.slice(name.indexOf(':') + 1);
const prefixOf = (name: string): string => (name.indexOf(':') < 0 ? '' : name.slice(0, name.indexOf(':')));
const endsWith = (text: string, suffix: string): boolean => text.slice(-suffix.length) === suffix;

const MEDIA_TYPES: { [extension: string]: string } = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jpe: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  svg: 'image/svg+xml',
  pict: 'image/pict',
};

function extensionOf(name: string): string {
  const dot = name.lastIndexOf('.');
  const extension = dot < 0 || dot < name.lastIndexOf('/') ? '' : name.slice(dot + 1).toLowerCase();
  return /^[a-z0-9]{1,8}$/.test(extension) ? extension : '';
}

/** The prefixes a root element binds to the namespace of relationships. */
function relPrefixes(rootAttrs: AttrList): Set<string> {
  const out = new Set<string>();
  for (const [name, value] of rootAttrs) if (name.indexOf('xmlns:') === 0 && isRelNamespace(value)) out.add(name.slice(6));
  return out;
}

/**
 * Calls `visit` for every attribute of a kept tree that names a relationship:
 * one whose prefix is bound to the namespace of relationships, by the root of
 * the part or by the tree itself. Returns the tree, or a copy in which the
 * values `visit` returned replace the ones that were there.
 */
function mapRelRefs(el: XmlElement, prefixes: Set<string>, visit: (id: string, el: XmlElement) => string, depth = 0): XmlElement {
  if (depth > MAX_TREE_DEPTH) return bad('The XML is nested too deep');
  let scope = prefixes;
  el.attrs.forEach((value, name) => {
    if (name.indexOf('xmlns:') !== 0) return;
    const prefix = name.slice(6);
    if (isRelNamespace(value) === scope.has(prefix)) return;
    if (scope === prefixes) scope = new Set(prefixes);
    if (isRelNamespace(value)) scope.add(prefix);
    else scope.delete(prefix);
  });
  let attrs: Map<string, string> | undefined;
  el.attrs.forEach((value, name) => {
    const colon = name.indexOf(':');
    if (colon < 0 || name.indexOf('xmlns:') === 0 || !scope.has(name.slice(0, colon))) return;
    const next = visit(value, el);
    if (next === value) return;
    if (!attrs) attrs = new Map(el.attrs);
    attrs.set(name, next);
  });
  let children: XmlElement['children'] | undefined;
  for (let i = 0; i < el.children.length; i++) {
    const node = el.children[i];
    if (node.type !== 'element') continue;
    const next = mapRelRefs(node, scope, visit, depth + 1);
    if (next === node) continue;
    if (!children) children = el.children.slice();
    children[i] = next;
  }
  if (!attrs && !children) return el;
  return { type: 'element', name: el.name, attrs: attrs || el.attrs, children: children || el.children };
}

function childNamed(el: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const node of el.children) if (node.type === 'element' && localOf(node.name) === local) return node;
  return undefined;
}

/** A <w:drawing> with a frame: the tree is what a writer writes. Anything else stands in for a drawing the model describes by its fields. */
const isDrawingTree = (xml: XmlElement | undefined): boolean => !!xml && localOf(xml.name) === 'drawing' && !!(childNamed(xml, 'inline') || childNamed(xml, 'anchor'));

// ------------------------------------------------------------------ the copy

interface CopyContext {
  doc: DocModel;
  relationships: Map<string, DocRelationship>;
  prefixes: Set<string>;
  out: DocFragment;
  wraps: Map<Wrapper, Wrapper | undefined>;
  groups: Map<object, object>;
  styleIds: Set<string>;
  numIds: Set<number>;
  budget: { paragraphs: number; items: number; cells: number };
}

function noteMedia(ctx: CopyContext, name: string): void {
  if (ctx.out.media.has(name)) return;
  const own = ctx.doc.media ? ctx.doc.media.get(name) : undefined;
  if (own) ctx.out.media.set(name, { type: own.type, bytes: own.bytes });
  else ctx.out.media.set(name, { type: MEDIA_TYPES[extensionOf(name)] || 'application/octet-stream' });
}

/** Is a relationship one a fragment can take along: a picture of the package, or an address outside of it (a linked picture, a link). */
function carriedRel(rel: DocRelationship | undefined): rel is DocRelationship {
  if (!rel) return false;
  if (rel.external) return endsWith(rel.type, '/image') || endsWith(rel.type, '/hyperlink');
  return rel.part !== null && endsWith(rel.type, '/image');
}

function noteRel(ctx: CopyContext, rel: DocRelationship): void {
  ctx.out.relationships.set(rel.id, rel);
  if (!rel.external && rel.part !== null) noteMedia(ctx, rel.part);
}

function imageRefs(item: DrawingItem): ImageRef[] {
  const out: ImageRef[] = [];
  if (item.images) for (const picture of item.images) out.push(picture.image);
  if (item.image && out.indexOf(item.image) < 0) out.push(item.image);
  return out;
}

/** The drawing when a fragment can carry it: it holds no text, and what its XML refers to are pictures and addresses only. */
function carryDrawing(ctx: CopyContext, item: DrawingItem): DrawingItem | undefined {
  if (item.textBoxes && item.textBoxes.length) return undefined;
  if (isDrawingTree(item.xml)) {
    const rels: DocRelationship[] = [];
    let ok = true;
    mapRelRefs(item.xml, ctx.prefixes, (id) => {
      const rel = ctx.relationships.get(id);
      if (carriedRel(rel)) rels.push(rel);
      else ok = false;
      return id;
    });
    if (!ok) return undefined;
    for (const rel of rels) noteRel(ctx, rel);
    return item;
  }
  // A drawing read from a binary file: its pictures are named by their part, and are in the model.
  if (item.kind !== 'picture' && !(item.kind === 'shape' && item.shape)) return undefined;
  const refs = imageRefs(item);
  for (const ref of refs) if (ref.target === undefined || ref.external || ref.linked) return undefined;
  for (const ref of refs) noteMedia(ctx, ref.target as string);
  return item;
}

const PLAIN_ITEMS = new Set<string>(['text', 'tab', 'break', 'noBreakHyphen', 'softHyphen', 'symbol', 'positionalTab', 'pageNumber', 'separator', 'continuationSeparator', 'instrText', 'fieldChar']);

function carryItem(ctx: CopyContext, item: RunItem): RunItem | undefined {
  if (PLAIN_ITEMS.has(item.type)) return item;
  if (item.type === 'drawing') return carryDrawing(ctx, item);
  return undefined;
}

/** Wrappers that are data and nothing else: they go with the text they hold. */
const CARRIED_WRAPPERS = new Set<string>(['hyperlink', 'simpleField', 'dir', 'bdo']);

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
  if (depth > MAX_TREE_DEPTH) return el;
  let children: XmlElement['children'] | undefined;
  for (let i = 0; i < el.children.length; i++) {
    const node = el.children[i];
    if (node.type !== 'element') {
      if (children) children.push(node);
      continue;
    }
    const local = localOf(node.name);
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
function plainExtras(extras: ExtraXml[] | undefined, marks: boolean): ExtraXml[] | undefined {
  if (!extras) return undefined;
  let out: ExtraXml[] | undefined;
  for (let i = 0; i < extras.length; i++) {
    const extra = extras[i];
    const local = localOf(extra.el.name);
    // The parent itself (index -1) is kept for what it carries; only what is in it goes.
    const gone = extra.index >= 0 && (REVISION_RECORDS.has(local) || (marks && REVISION_MARKS.has(local)));
    const el = gone ? undefined : plainElement(extra.el, marks);
    if (el === extra.el) {
      if (out) out.push(extra);
      continue;
    }
    if (!out) out = extras.slice(0, i);
    if (el) out.push({ ...extra, el });
  }
  if (!out) return extras;
  return out.length ? out : undefined;
}

/**
 * Properties (of a paragraph, a run, a section, a table, a row, a cell) as a copy carries them:
 * without what makes them a tracked change. See the head of this module.
 */
function plainProps<T extends object>(props: T): T {
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
      if (next !== inner) set(name, next);
    }
  }
  const out = (copy || record) as unknown as T;
  PLAIN.set(props, out);
  // What is plain stays what it is when asked again.
  if (out !== props) PLAIN.set(out, out);
  return out;
}

/** The chain of wrappers of an item as the fragment has it: copies of those it carries, without the others. */
function carryWrap(ctx: CopyContext, wrap: Wrapper | undefined, depth = 0): Wrapper | undefined {
  if (!wrap || depth > 256) return undefined;
  if (ctx.wraps.has(wrap)) return ctx.wraps.get(wrap);
  const parent = carryWrap(ctx, wrap.parent, depth + 1);
  let out = parent;
  let carried = CARRIED_WRAPPERS.has(wrap.kind) && !wrap.extras && !wrap.misc;
  let rel: DocRelationship | undefined;
  if (carried && wrap.kind === 'hyperlink' && wrap.rId !== undefined) {
    rel = ctx.relationships.get(wrap.rId);
    // A link through a relationship that is no address is not one the fragment can make again elsewhere.
    if (rel && !(rel.external && carriedRel(rel))) carried = false;
  }
  if (carried) {
    const copy = { ...wrap } as Wrapper;
    delete copy.srcStart;
    delete copy.srcEnd;
    if (parent) copy.parent = parent;
    else delete copy.parent;
    if (rel) {
      noteRel(ctx, rel);
      if (copy.kind === 'hyperlink' && copy.target === undefined) copy.target = rel.target;
    }
    out = copy;
  }
  ctx.wraps.set(wrap, out);
  return out;
}

function noteRunProps(ctx: CopyContext, props: RunProps | undefined): void {
  if (props && props.style !== undefined) ctx.styleIds.add(props.style);
}

function noteParaProps(ctx: CopyContext, props: ParaProps): void {
  if (props.style !== undefined) ctx.styleIds.add(props.style);
  if (props.numbering && props.numbering.numId !== undefined && props.numbering.numId !== 0) ctx.numIds.add(props.numbering.numId);
  noteRunProps(ctx, props.markProps);
}

/**
 * What a removal of the pieces [from, to) takes, as inlines of the fragment.
 * A complex field is carried as a field when all of it goes and all of it
 * can be carried; of any other field only the text of its result is.
 */
function carryPieces(ctx: CopyContext, pieces: Piece[], from: number, to: number): Inline[] {
  const flags = removedPieces(pieces, from, to);
  const kept: (Piece | undefined)[] = [];
  for (let q = 0; q < pieces.length; q++) {
    const piece = pieces[q];
    let carried: Piece | undefined;
    if (flags[q] && piece.item) {
      const item = carryItem(ctx, piece.item);
      if (item) {
        if (--ctx.budget.items < 0) bad('The range holds too much to copy');
        let group = ctx.groups.get(piece.group);
        if (!group) ctx.groups.set(piece.group, (group = {}));
        carried = { item, props: piece.props && plainProps(piece.props), wrap: carryWrap(ctx, piece.wrap), group, index: -1, code: piece.code, len: piece.len };
      }
    }
    kept.push(carried);
  }
  // The outermost fields, one after the other: whole, or without their marks and code.
  let any = false;
  for (const piece of pieces) if (piece.item && (piece.code || piece.item.type === 'fieldChar' || piece.item.type === 'instrText')) any = true;
  if (any) {
    const { pairs } = fieldPairs(pieces);
    let done = -1;
    for (const pair of pairs) {
      if (pair.begin <= done) continue;
      const end = pair.end >= 0 ? pair.end : pieces.length - 1;
      done = end;
      let whole = pair.end >= 0;
      for (let q = pair.begin; q <= end && whole; q++) {
        if (isStructural(pieces[q]) ? !flags[q] || !kept[q] : flags[q] && !kept[q]) whole = false;
      }
      if (whole) continue;
      for (let q = pair.begin; q <= end; q++) if (isStructural(pieces[q])) kept[q] = undefined;
    }
  }
  const out: Piece[] = [];
  for (const piece of kept) {
    if (!piece) continue;
    noteRunProps(ctx, piece.props);
    out.push(piece);
  }
  return rebuild(out);
}

function carryParagraph(ctx: CopyContext, paragraph: Paragraph, pieces: Piece[], from: number, to: number, withMark: boolean): Paragraph {
  if (--ctx.budget.paragraphs < 0) bad('The range holds too much to copy');
  const props = withMark ? plainProps(paragraph.props) : EMPTY_PARA_PROPS;
  noteParaProps(ctx, props);
  return { type: 'paragraph', props, children: carryPieces(ctx, pieces, from, to) };
}

function wholeParagraph(ctx: CopyContext, paragraph: Paragraph): Paragraph {
  const pieces = flatten(paragraph);
  const out = carryParagraph(ctx, paragraph, pieces, 0, pieces.length, true);
  // A paragraph that goes whole keeps the attributes of its element (see pasteAttrs for those that name it).
  if (paragraph.attrs) out.attrs = paragraph.attrs;
  return out;
}

/** What goes with a table or a row that is removed whole: its kept children, except the marks of ranges (which stay in the document). */
function carryExtras(extras: ExtraXml[] | undefined): ExtraXml[] | undefined {
  if (!extras) return undefined;
  // And without the record of a changed grid; an element that wraps rows is not looked into.
  const out = plainExtras(
    extras.filter((extra) => !extra.wrap && looseMark(extra) === undefined),
    false,
  );
  return out && out.length ? out : undefined;
}

function carryCell(ctx: CopyContext, cell: TableCell, depth: number, blank: boolean): TableCell {
  if (--ctx.budget.cells < 0) bad('The range holds too much to copy');
  const blocks: Block[] = blank ? [] : carryBlocks(ctx, cell.blocks, depth);
  let paragraph = false;
  for (const block of blocks) if (block.type === 'paragraph') paragraph = true;
  if (!paragraph || blocks[blocks.length - 1].type !== 'paragraph') {
    // A cell ends with a paragraph; the one a cell without content keeps is formatted like its first.
    let props: ParaProps = EMPTY_PARA_PROPS;
    for (const block of cell.blocks) {
      if (block.type !== 'paragraph') continue;
      props = plainProps(withoutSection(block.props));
      break;
    }
    if (!paragraph) noteParaProps(ctx, props);
    blocks.push({ type: 'paragraph', props: paragraph ? EMPTY_PARA_PROPS : props, children: [] });
  }
  const out: TableCell = { props: plainProps(cell.props), blocks };
  if (cell.attrs) out.attrs = cell.attrs;
  return out;
}

function carryRow(row: TableRow, cells: TableCell[], props: RowProps = row.props): TableRow {
  const out: TableRow = { props: plainProps(props), cells };
  if (row.attrs) out.attrs = row.attrs;
  if (row.tableProps) out.tableProps = plainProps(row.tableProps);
  const extras = carryExtras(row.extras);
  if (extras) out.extras = extras;
  return out;
}

function noteTable(ctx: CopyContext, props: TableProps, depth: number): void {
  if (props.style !== undefined) ctx.styleIds.add(props.style);
  if (depth > ctx.out.tableDepth) ctx.out.tableDepth = depth;
}

function wholeTable(ctx: CopyContext, table: Table, depth: number): Table {
  if (depth > 64) return bad('The tables are nested too deep');
  noteTable(ctx, table.props, depth);
  const rows = table.rows.map((row) =>
    carryRow(
      row,
      row.cells.map((cell) => carryCell(ctx, cell, depth, false)),
    ),
  );
  const out: Table = { type: 'table', props: plainProps(table.props), grid: table.grid.slice(), rows };
  if (table.attrs) out.attrs = table.attrs;
  const extras = carryExtras(table.extras);
  if (extras) out.extras = extras;
  return out;
}

/** A merged cell that continued one of a row the copy does not have begins the merge in the copy. */
function restartMerges(row: TableRow): TableRow {
  let cells: TableCell[] | undefined;
  for (let c = 0; c < row.cells.length; c++) {
    if (row.cells[c].props.vMerge !== 'continue') continue;
    if (!cells) cells = row.cells.slice();
    cells[c] = { ...row.cells[c], props: { ...row.cells[c].props, vMerge: 'restart' } };
  }
  return cells ? { ...row, cells } : row;
}

/**
 * The rows of a table that hold cells `takes` picks, as a table. With
 * `window` only those cells, over the grid columns they cover (a rectangle
 * of cells); without it the rows whole, the other cells of them empty (what
 * a range that leaves the table takes: see deleteSpan).
 */
function partTable(ctx: CopyContext, table: Table, takes: (row: number, cell: number) => boolean, window: boolean, depth: number): Table {
  noteTable(ctx, table.props, depth);
  const picked: { row: TableRow; r: number; first: number; last: number; spans: [number, number][] }[] = [];
  let left = Infinity;
  let right = -Infinity;
  for (let r = 0; r < table.rows.length; r++) {
    const row = table.rows[r];
    let first = -1;
    let last = -1;
    for (let c = 0; c < row.cells.length; c++) {
      if (!takes(r, c)) continue;
      if (first < 0) first = c;
      last = c;
    }
    if (first < 0) continue;
    const spans = rowSpans(row);
    left = Math.min(left, spans[first][0]);
    right = Math.max(right, spans[last][1]);
    picked.push({ row, r, first, last, spans });
  }
  if (!picked.length) return bad('The range holds no cell of the table');
  const whole = !window || (left <= 0 && right >= table.grid.length);
  const rows: TableRow[] = [];
  for (const entry of picked) {
    const row = entry.row;
    if (whole && !window) {
      rows.push(carryRow(row, row.cells.map((cell, c) => carryCell(ctx, cell, depth, !takes(entry.r, c)))));
      continue;
    }
    const cells: TableCell[] = [];
    for (let c = entry.first; c <= entry.last; c++) cells.push(carryCell(ctx, row.cells[c], depth, !takes(entry.r, c)));
    const props: RowProps = { ...row.props };
    const before = entry.spans[entry.first][0] - left;
    const after = right - entry.spans[entry.last][1];
    if (before > 0) props.gridBefore = before;
    else delete props.gridBefore;
    if (after > 0) props.gridAfter = after;
    else delete props.gridAfter;
    if (before !== (row.props.gridBefore || 0)) delete props.widthBefore;
    if (after !== (row.props.gridAfter || 0)) delete props.widthAfter;
    rows.push(carryRow(row, cells, props));
  }
  if (picked[0].r > 0) rows[0] = restartMerges(rows[0]);
  let props = table.props;
  let grid = table.grid.slice();
  if (!whole) {
    grid = table.grid.slice(Math.max(0, left), Math.max(0, right));
    // The width the table states is that of all its columns.
    if (props.width && props.width.type !== 'auto') props = { ...props, width: { type: 'auto', value: 0 } };
  }
  const out: Table = { type: 'table', props: plainProps(props), grid, rows };
  if (table.attrs) out.attrs = table.attrs;
  const extras = whole && picked.length === table.rows.length ? carryExtras(table.extras) : undefined;
  if (extras) out.extras = extras;
  return out;
}

/** The paragraphs and tables of a list of blocks, whole; what is not content (marks, kept XML between blocks) stays where it is. */
function carryBlocks(ctx: CopyContext, blocks: Block[], depth: number): Block[] {
  const out: Block[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') out.push(wholeParagraph(ctx, block));
    else if (block.type === 'table') out.push(wholeTable(ctx, block, depth + 1));
  }
  return out;
}

/** The definitions of the styles and lists the fragment names, and of those these name. */
function collectDefinitions(ctx: CopyContext): void {
  const doc = ctx.doc;
  const out = ctx.out;
  const styles: string[] = [];
  const nums: number[] = [];
  ctx.styleIds.forEach((id) => styles.push(id));
  ctx.numIds.forEach((id) => nums.push(id));
  const style = (id: string | undefined) => {
    if (id !== undefined && !out.styles.has(id)) styles.push(id);
  };
  // Each turn takes one name off a list; what it adds is bounded by the number of definitions kept.
  while (styles.length || nums.length) {
    if (styles.length) {
      const id = styles.pop() as string;
      const found = doc.styles.styles.get(id);
      if (!found || out.styles.has(id)) continue;
      if (out.styles.size >= MAX_DEFINITIONS) {
        out.definitionsCut = true;
        continue;
      }
      out.styles.set(id, found);
      style(found.basedOn);
      style(found.link);
      style(found.next);
      const numbering = found.para ? found.para.numbering : undefined;
      if (numbering && numbering.numId !== undefined && numbering.numId !== 0) nums.push(numbering.numId);
      continue;
    }
    const id = nums.pop() as number;
    const num = doc.numbering.nums.get(id);
    if (!num || out.nums.has(id)) continue;
    if (out.nums.size >= MAX_DEFINITIONS) {
      out.definitionsCut = true;
      continue;
    }
    out.nums.set(id, num);
    const abstract = doc.numbering.abstracts.get(num.abstractId);
    if (!abstract || out.abstracts.has(abstract.id)) continue;
    out.abstracts.set(abstract.id, abstract);
    style(abstract.numStyleLink);
    style(abstract.styleLink);
    abstract.levels.forEach((level) => style(level.paraStyle));
    num.overrides.forEach((override) => style(override.level ? override.level.paraStyle : undefined));
  }
}

function measure(blocks: Block[], depth = 0): number {
  let text = 0;
  if (depth > 256) return 0;
  for (const block of blocks) {
    if (block.type === 'paragraph') text += paragraphText(block).length + 1;
    else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) text += measure(cell.blocks, depth + 1);
  }
  return text;
}

/** The part of a path up to and including the last text box it enters. */
function rootOf(path: number[]): number {
  for (let i = path.length - 1; i >= 0; i--) if (path[i] === -1) return i + 3;
  return 0;
}

/**
 * A copy of what lies between the two ends of a range: what deleteRange()
 * would remove of it, as far as a fragment can carry it (see the head of this
 * module). A range that crosses the border of a table is taken as
 * expandRangeForTables() gives it, like the deletion. Throws RangeError for a
 * range that is not one of the document or holds more than FRAGMENT_LIMITS.
 */
export function fragmentOf(doc: DocModel, range: DocRange): DocFragment {
  const { start: first, end: last } = orderRange(asRange(range));
  const a = resolvePosition(doc, first);
  const b = resolvePosition(doc, last);
  comparePositions(first, last);
  const host = storyOf(doc, first.story).host;
  const out: DocFragment = {
    blocks: [],
    openStart: false,
    openEnd: false,
    origin: originOf(doc),
    part: host.part,
    rootAttrs: host.rootAttrs,
    strict: doc.strict,
    relationships: new Map(),
    media: new Map(),
    styles: new Map(),
    abstracts: new Map(),
    nums: new Map(),
    styleNamespaces: partNamespaces(doc.styles) || declarations(host.rootAttrs),
    numberingNamespaces: partNamespaces(doc.numbering) || declarations(host.rootAttrs),
    textLength: 0,
    tableDepth: 0,
  };
  const ctx: CopyContext = {
    doc,
    relationships: host.relationships,
    prefixes: relPrefixes(host.rootAttrs),
    out,
    wraps: new Map(),
    groups: new Map(),
    styleIds: new Set(),
    numIds: new Set(),
    budget: { paragraphs: FRAGMENT_LIMITS.paragraphs, items: FRAGMENT_LIMITS.items, cells: FRAGMENT_LIMITS.cells },
  };
  // Inside a text box the paths are followed from the box on.
  const root = rootOf(first.path);
  const base = first.path.slice(0, root);
  const start: DocPosition = { ...first, path: first.path.slice(root) };
  const end: DocPosition = { ...last, path: last.path.slice(root) };
  const story = first.story;
  const { depth, level } = parting(start, end);
  const story0 = (path: number[]) => containerBlocks(doc, story, base.concat(path));

  if (samePosition(first, last)) {
    // Nothing between the two ends: a deletion removes nothing either (a caret inside a field that is one unit covers no part of it).
  } else if (depth === start.path.length && depth === end.path.length) {
    const pieces = flatten(a.paragraph);
    const from = startIndex(pieces, start);
    const to = caretIndex(pieces, end);
    if (from < to) {
      out.blocks.push(carryParagraph(ctx, a.paragraph, pieces, from, to, false));
      out.openStart = out.openEnd = true;
    }
  } else if (depth % 3 !== 0) {
    const table = story0(start.path.slice(0, level))[start.path[level]] as Table;
    out.blocks.push(partTable(ctx, table, cellsBetween(table, start.path[level + 1], start.path[level + 2], end.path[level + 1], end.path[level + 2]), true, 1));
  } else {
    const container = start.path.slice(0, level);
    const blocks = story0(container);
    const from = start.path[level];
    const to = end.path[level];
    const samePath = (x: number[] | undefined, y: number[]) => x !== undefined && x.length === y.length && x.every((value, i) => value === y[i]);
    if (start.path.length === level + 1) {
      const pieces = flatten(a.paragraph);
      out.blocks.push(carryParagraph(ctx, a.paragraph, pieces, startIndex(pieces, start), pieces.length, true));
      out.openStart = true;
    } else {
      const table = blocks[from] as Table;
      const goes = start.offset === 0 && samePath(paragraphIn(blocks, container, from, 1), start.path);
      const [r1, c1] = [start.path[level + 1], start.path[level + 2]];
      out.blocks.push(goes ? wholeTable(ctx, table, 1) : partTable(ctx, table, (r, c) => r > r1 || (r === r1 && c >= c1), false, 1));
    }
    for (let i = from + 1; i < to; i++) {
      const block = blocks[i];
      if (block.type === 'paragraph') out.blocks.push(wholeParagraph(ctx, block));
      else if (block.type === 'table') out.blocks.push(wholeTable(ctx, block, 1));
    }
    if (end.path.length === level + 1) {
      const pieces = flatten(b.paragraph);
      out.blocks.push(carryParagraph(ctx, b.paragraph, pieces, 0, caretIndex(pieces, end), false));
      out.openEnd = true;
      if (out.openStart) {
        out.endProps = plainProps(b.paragraph.props);
        // Only ever compared with the paragraph a cut left behind, which has whatever tracked changes it had.
        out.joinedProps = joinedProps(a.paragraph, b.paragraph);
        noteParaProps(ctx, out.endProps);
      }
    } else {
      const table = blocks[to] as Table;
      const goes = end.offset === b.length && samePath(paragraphIn(blocks, container, to, -1), end.path);
      const [r2, c2] = [end.path[level + 1], end.path[level + 2]];
      out.blocks.push(goes ? wholeTable(ctx, table, 1) : partTable(ctx, table, (r, c) => r < r2 || (r === r2 && c <= c2), false, 1));
    }
  }
  collectDefinitions(ctx);
  out.textLength = Math.max(0, measure(out.blocks) - 1);
  return out;
}

/** Does a fragment hold anything at all: text, or something that shows without text (a picture, a table). */
export function fragmentHasContent(fragment: DocFragment): boolean {
  for (const block of fragment.blocks) {
    if (block.type === 'table') return true;
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) if (inline.type === 'run' && inline.content.length > 0) return true;
  }
  return fragment.blocks.length > 1;
}

// ------------------------------------------------------------------ the paste

export interface InsertFragmentOptions extends DocEditOptions {
  /** Is a part of this name in the package already? See InsertImageOptions.taken. */
  taken?: (part: string) => boolean;
  /**
   * The fragment was cut, not copied. A cut from one paragraph into another
   * joins the two and leaves the formatting of one of them (see deleteSpan).
   * When such a fragment is pasted into a paragraph that has the formatting
   * the cut left (the place it was cut from, in the first place), the text
   * after the caret gets the formatting of the paragraph the fragment ended
   * in, so that a cut and a paste at the same place leave every paragraph
   * formatted as it was. Without it, and in any other paragraph, the text
   * after the caret keeps the formatting of the paragraph it is in, as it
   * does in Word.
   */
  moved?: boolean;
}

interface PasteContext {
  doc: DocModel;
  draft: Draft;
  story: string;
  fragment: DocFragment;
  sameDoc: boolean;
  samePart: boolean;
  /**
   * The same context with the declarations the definitions of styles and of
   * lists have to bring: those go into the styles part and the numbering
   * part, whose roots declare what they declare, whatever the part of the
   * story does.
   */
  forStyles: PasteContext;
  forNumbering: PasteContext;
  /** Namespace declarations kept XML has to bring itself: [attribute name, uri]. */
  decls: [string, string][];
  declPrefixes: Set<string>;
  prefixes: Set<string>;
  /** Section breaks can be pasted here: into the body, outside of tables. */
  sections: boolean;
  options: InsertFragmentOptions | undefined;
  memo: Map<object, unknown>;
  /** Paragraph properties as they are pasted, without and with their section break. */
  paras: Map<ParaProps, (ParaProps | undefined)[]>;
  rels: Map<string, string>;
  media: Map<string, string>;
  nums: Map<number, number>;
  abstracts: Map<number, number>;
  addStyles: Map<string, DocStyle>;
  addAbstracts: [number, AbstractNum][];
  addNums: [number, NumInstance][];
  seenStyles: Set<string>;
  taken: Taken | undefined;
  drawing: number;
}

const isElement = (value: object): value is XmlElement => (value as XmlElement).type === 'element' && (value as XmlElement).attrs instanceof Map && Array.isArray((value as XmlElement).children);

/** The prefixes a tree uses and does not declare itself, of those in `wanted`. */
function freePrefixes(el: XmlElement, wanted: Set<string>, declared: Set<string>, out: Set<string>, depth = 0): void {
  if (depth > MAX_TREE_DEPTH) return;
  let scope = declared;
  el.attrs.forEach((_value, name) => {
    const prefix = name === 'xmlns' ? '' : name.indexOf('xmlns:') === 0 ? name.slice(6) : undefined;
    if (prefix === undefined || scope.has(prefix)) return;
    if (scope === declared) scope = new Set(declared);
    scope.add(prefix);
  });
  const use = (prefix: string) => {
    if (wanted.has(prefix) && !scope.has(prefix)) out.add(prefix);
  };
  use(prefixOf(el.name));
  el.attrs.forEach((_value, name) => {
    if (name.indexOf(':') > 0 && name.indexOf('xmlns:') !== 0) use(prefixOf(name));
  });
  for (const node of el.children) if (node.type === 'element') freePrefixes(node, wanted, scope, out, depth + 1);
}

/** A kept tree as it can stand in the part it is pasted into: with the declarations of the namespaces that part does not declare the same way. */
function adaptTree(ctx: PasteContext, el: XmlElement): XmlElement {
  if (!ctx.decls.length) return el;
  const known = ctx.memo.get(el);
  if (known !== undefined) return known as XmlElement;
  const free = new Set<string>();
  freePrefixes(el, ctx.declPrefixes, new Set(), free);
  let out = el;
  if (free.size) {
    const attrs = new Map(el.attrs);
    for (const [name, uri] of ctx.decls) if (free.has(name === 'xmlns' ? '' : name.slice(6))) attrs.set(name, uri);
    out = { type: 'element', name: el.name, attrs, children: el.children };
  }
  ctx.memo.set(el, out);
  return out;
}

/**
 * A value of the model (a property bag, an item) with every kept tree in it
 * adapted: the value itself when nothing in it changes, so bags that were
 * shared stay shared.
 */
function adaptValue<T>(ctx: PasteContext, value: T, depth = 0): T {
  if (!ctx.decls.length || value === null || typeof value !== 'object' || depth > 12 || value instanceof Uint8Array) return value;
  const object = value as unknown as object;
  if (isElement(object)) return adaptTree(ctx, object) as unknown as T;
  const known = ctx.memo.get(object);
  if (known !== undefined) return known as T;
  let out: unknown = value;
  if (Array.isArray(value)) {
    let copy: unknown[] | undefined;
    for (let i = 0; i < value.length; i++) {
      const next = adaptValue(ctx, value[i], depth + 1);
      if (next === value[i]) continue;
      if (!copy) copy = value.slice();
      copy[i] = next;
    }
    if (copy) out = copy;
  } else if (value instanceof Map) {
    let copy: Map<unknown, unknown> | undefined;
    value.forEach((entry, key) => {
      const next = adaptValue(ctx, entry, depth + 1);
      if (next === entry) return;
      if (!copy) copy = new Map(value);
      copy.set(key, next);
    });
    if (copy) out = copy;
  } else {
    const record = value as unknown as { [key: string]: unknown };
    let copy: { [key: string]: unknown } | undefined;
    for (const key of Object.keys(record)) {
      const next = adaptValue(ctx, record[key], depth + 1);
      if (next === record[key]) continue;
      if (!copy) copy = { ...record };
      copy[key] = next;
    }
    if (copy) out = copy;
  }
  ctx.memo.set(object, out);
  return out as T;
}

function adaptAttrs(ctx: PasteContext, attrs: AttrList | undefined): AttrList | undefined {
  if (!attrs || !ctx.decls.length) return attrs;
  let out: AttrList | undefined;
  for (const [name, uri] of ctx.decls) {
    const prefix = name === 'xmlns' ? undefined : name.slice(6);
    if (prefix === undefined || !attrs.some(([attr]) => prefixOf(attr) === prefix)) continue;
    if (!out) out = attrs.slice();
    out.push([name, uri]);
  }
  return out || attrs;
}

function takenOf(ctx: PasteContext): Taken {
  if (!ctx.taken) {
    ctx.taken = takenByDocument(ctx.doc);
    ctx.drawing = ctx.taken.drawing;
  }
  return ctx.taken;
}

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean => {
  if (a === b) return true;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
};

/** The names of parts the stories of a document refer to, and of the pictures it carries. */
function hasPart(ctx: PasteContext, name: string): boolean {
  return takenOf(ctx).parts.has(name.toLowerCase());
}

/** The name of the part a picture of the fragment has in the document: the one it had there, or that of a part added for it. */
function ensureMedia(ctx: PasteContext, name: string): string {
  const known = ctx.media.get(name);
  if (known !== undefined) return known;
  let part: string | undefined;
  if (ctx.sameDoc && hasPart(ctx, name)) part = name;
  else {
    const media = ctx.fragment.media.get(name);
    const bytes = media ? media.bytes : undefined;
    if (!media || !(bytes instanceof Uint8Array)) return bad('A picture of the copy is not available');
    // The same picture pasted before is one part.
    if (ctx.doc.media) ctx.doc.media.forEach((own, key) => (part === undefined && own.type === media.type && sameBytes(own.bytes, bytes) ? (part = key) : 0));
    if (part === undefined) {
      const taken = takenOf(ctx);
      const isTaken = ctx.options ? ctx.options.taken : undefined;
      const extension = extensionOf(name) || 'bin';
      let number = taken.image + 1;
      let made = `/word/media/image${number}.${extension}`;
      // Each turn tries the next number; the caller is asked a bounded number of times.
      for (let tries = 0; taken.parts.has(made.toLowerCase()) || (typeof isTaken === 'function' && isTaken(made) === true); tries++) {
        if (tries >= 10000) bad('No name is free for the picture');
        number++;
        made = `/word/media/image${number}.${extension}`;
      }
      taken.image = number;
      taken.parts.add(made.toLowerCase());
      ctx.draft.apply({ type: 'setMedia', name: made, media: { bytes, type: media.type } });
      part = made;
    }
  }
  ctx.media.set(name, part);
  return part;
}

const kindOfRel = (type: string): string => type.slice(type.lastIndexOf('/') + 1);

/** The id a relationship of the fragment has in the part of the story: one that part has for the same target, or a new one. */
function mapRel(ctx: PasteContext, id: string, rel: DocRelationship | undefined = ctx.fragment.relationships.get(id)): string {
  const known = ctx.rels.get(id);
  if (known !== undefined) return known;
  if (!rel) return id;
  const host = storyOf(ctx.doc, ctx.story).host;
  const kind = kindOfRel(rel.type);
  const part = rel.external ? null : ensureMedia(ctx, rel.part as string);
  const fits = (own: DocRelationship): boolean => kindOfRel(own.type) === kind && own.external === rel.external && (rel.external ? own.target === rel.target : own.part === part);
  let mapped: string | undefined;
  const own = ctx.samePart ? host.relationships.get(id) : undefined;
  if (own && fits(own)) mapped = id;
  else host.relationships.forEach((other) => (mapped === undefined && fits(other) ? (mapped = other.id) : 0));
  if (mapped === undefined) {
    mapped = freeRelId(host.relationships);
    const base = ctx.doc.strict ? REL_NAMESPACES[1] : REL_NAMESPACES[0];
    // Relative to the part of the story when that is beside the media folder, as Word writes it.
    const beside = /^\/word\/[^/]+$/.test(host.part);
    const target = part === null ? rel.target : beside && part.indexOf('/word/') === 0 ? part.slice('/word/'.length) : part;
    ctx.draft.apply({ type: 'setRelationship', story: ctx.story, id: mapped, relationship: { id: mapped, type: `${base}/${kind}`, target, external: rel.external, part } });
  }
  ctx.rels.set(id, mapped);
  return mapped;
}

function adaptRef(ctx: PasteContext, ref: ImageRef, tree: boolean): ImageRef {
  if (!tree) {
    const target = ensureMedia(ctx, ref.target as string);
    return target === ref.target ? ref : { ...ref, target };
  }
  const rel = ctx.fragment.relationships.get(ref.rId);
  if (!rel) return ref;
  const rId = mapRel(ctx, ref.rId);
  const target = rel.external ? ref.target : ensureMedia(ctx, rel.part as string);
  return rId === ref.rId && target === ref.target ? ref : { ...ref, rId, target };
}

function adaptDrawing(ctx: PasteContext, item: DrawingItem): DrawingItem {
  const tree = isDrawingTree(item.xml);
  const out: DrawingItem = { ...item };
  if (tree) {
    let xml = mapRelRefs(item.xml, ctx.prefixes, (id) => mapRel(ctx, id));
    // An id no other drawing of the document has: the drawing is a new one here.
    const frame = childNamed(xml, 'inline') || childNamed(xml, 'anchor');
    const docPr = childNamed(frame, 'docPr');
    if (frame && docPr) {
      takenOf(ctx);
      if (ctx.drawing >= 0x7ffffffe) bad('The document has no room for another drawing');
      const id = String(++ctx.drawing);
      const attrs = new Map(docPr.attrs);
      attrs.set('id', id);
      const fresh: XmlElement = { type: 'element', name: docPr.name, attrs, children: docPr.children };
      const inner: XmlElement = { type: 'element', name: frame.name, attrs: frame.attrs, children: frame.children.map((node) => (node === docPr ? fresh : node)) };
      xml = { type: 'element', name: xml.name, attrs: xml.attrs, children: xml.children.map((node) => (node === frame ? inner : node)) };
    }
    out.xml = adaptTree(ctx, xml);
  }
  const refs = new Map<ImageRef, ImageRef>();
  const ref = (old: ImageRef): ImageRef => {
    let next = refs.get(old);
    if (!next) refs.set(old, (next = adaptRef(ctx, old, tree)));
    return next;
  };
  if (item.images) out.images = item.images.map((picture) => ({ ...picture, image: ref(picture.image) }));
  if (item.image) out.image = ref(item.image);
  return out;
}

function adaptItem(ctx: PasteContext, item: RunItem): RunItem {
  if (item.type === 'drawing') return adaptDrawing(ctx, item);
  if (item.type === 'fieldChar') return adaptValue(ctx, item);
  return item;
}

function adaptWrap(ctx: PasteContext, wrap: Wrapper | undefined, depth = 0): Wrapper | undefined {
  if (!wrap || depth > 256) return undefined;
  const known = ctx.memo.get(wrap);
  if (known !== undefined) return known as Wrapper;
  const parent = adaptWrap(ctx, wrap.parent, depth + 1);
  let out = wrap;
  const change = (): Wrapper => (out === wrap ? (out = { ...wrap } as Wrapper) : out);
  if (parent !== wrap.parent) {
    const copy = change();
    if (parent) copy.parent = parent;
    else delete copy.parent;
  }
  if (wrap.kind === 'hyperlink') {
    let rId = wrap.rId;
    if (rId !== undefined && ctx.fragment.relationships.has(rId)) rId = mapRel(ctx, rId);
    else if (rId !== undefined && !ctx.samePart) rId = undefined;
    if (rId === undefined && wrap.target !== undefined && wrap.target.charAt(0) !== '#' && storyOf(ctx.doc, ctx.story).host.rootName !== '') {
      // A link that names its address only (one read from a binary file) gets the relationship a package needs.
      const key = '\u0000' + wrap.target;
      rId = mapRel(ctx, key, { id: key, type: REL_NAMESPACES[0] + '/hyperlink', target: wrap.target, external: true, part: null });
    }
    if (rId !== wrap.rId) {
      const copy = change() as Extract<Wrapper, { kind: 'hyperlink' }>;
      if (rId === undefined) delete copy.rId;
      else copy.rId = rId;
    }
  }
  if ('attrs' in wrap && wrap.attrs) {
    const attrs = adaptAttrs(ctx, wrap.attrs);
    if (attrs !== wrap.attrs) (change() as { attrs?: AttrList }).attrs = attrs;
  }
  ctx.memo.set(wrap, out);
  return out;
}

/** Brings the definition of a style along when the document has no style of that id. */
function ensureStyle(ctx: PasteContext, id: string | undefined, depth = 0): void {
  if (id === undefined || ctx.seenStyles.has(id) || depth > MAX_DEFINITIONS) return;
  ctx.seenStyles.add(id);
  if (ctx.doc.styles.styles.has(id)) return;
  const style = ctx.fragment.styles.get(id);
  if (!style) return;
  const styles = ctx.forStyles;
  let copy = adaptValue(styles, style);
  // The document has its own default styles.
  if (copy.isDefault) copy = { ...copy, isDefault: false };
  if (copy.para) {
    const para = adaptParaProps(styles, copy.para, false);
    if (para !== copy.para) copy = { ...copy, para };
  }
  ctx.addStyles.set(id, copy);
  ensureStyle(ctx, style.basedOn, depth + 1);
  ensureStyle(ctx, style.link, depth + 1);
  ensureStyle(ctx, style.next, depth + 1);
}

/** The id a list of the fragment has in the document: its own in the document it came from, a new list anywhere else. */
function mapNum(ctx: PasteContext, id: number): number {
  const known = ctx.nums.get(id);
  if (known !== undefined) return known;
  let mapped = id;
  const numbering = ctx.doc.numbering;
  const num = ctx.fragment.nums.get(id);
  if (!(ctx.sameDoc && numbering.nums.has(id)) && num) {
    const free = (map: Map<number, unknown>, added: [number, unknown][]): number => {
      let next = 1;
      map.forEach((_value, key) => (key >= next ? (next = key + 1) : 0));
      for (const [key] of added) if (key >= next) next = key + 1;
      return next;
    };
    const lists = ctx.forNumbering;
    // The ids are settled before anything else is looked at: a style a level names may name this list again.
    mapped = free(numbering.nums, ctx.addNums);
    ctx.nums.set(id, mapped);
    const entry: [number, NumInstance] = [mapped, num];
    ctx.addNums.push(entry);
    let abstractId = ctx.abstracts.get(num.abstractId);
    if (abstractId === undefined) {
      const abstract = ctx.fragment.abstracts.get(num.abstractId);
      abstractId = abstract ? free(numbering.abstracts, ctx.addAbstracts) : num.abstractId;
      ctx.abstracts.set(num.abstractId, abstractId);
      if (abstract) {
        ctx.addAbstracts.push([abstractId, { ...adaptValue(lists, abstract), id: abstractId }]);
        ensureStyle(ctx, abstract.numStyleLink);
        ensureStyle(ctx, abstract.styleLink);
        abstract.levels.forEach((level) => ensureStyle(ctx, level.paraStyle));
      }
    }
    entry[1] = { ...adaptValue(lists, num), id: mapped, abstractId };
    num.overrides.forEach((override) => ensureStyle(ctx, override.level ? override.level.paraStyle : undefined));
  }
  ctx.nums.set(id, mapped);
  return mapped;
}

function adaptParaProps(ctx: PasteContext, props: ParaProps, sections = ctx.sections): ParaProps {
  if (props === EMPTY_PARA_PROPS) return props;
  const memo = ctx.paras.get(props);
  const key = sections ? 1 : 0;
  if (memo && memo[key]) return memo[key] as ParaProps;
  ensureStyle(ctx, props.style);
  if (props.markProps) ensureStyle(ctx, props.markProps.style);
  let out = props;
  if (ctx.decls.length) {
    // Property by property: the memo of the bag itself holds its adapted forms.
    let copy: { [name: string]: unknown } | undefined;
    const record = props as unknown as { [name: string]: unknown };
    for (const name of Object.keys(record)) {
      const next = adaptValue(ctx, record[name], 1);
      if (next === record[name]) continue;
      if (!copy) copy = { ...record };
      copy[name] = next;
    }
    if (copy) out = copy as unknown as ParaProps;
  }
  const numId = props.numbering ? props.numbering.numId : undefined;
  if (numId !== undefined && numId !== 0) {
    const mapped = mapNum(ctx, numId);
    if (mapped !== numId) out = { ...out, numbering: { ...props.numbering, numId: mapped } };
  }
  if (out.section) {
    if (!sections) out = withoutSection(out);
    else if (!ctx.sameDoc && (out.section.headers || out.section.footers)) {
      // The headers and footers a section names are parts of the document it came from.
      const section = { ...out.section };
      delete section.headers;
      delete section.footers;
      out = { ...out, section };
    }
  }
  const slot: (ParaProps | undefined)[] = memo || [undefined, undefined];
  slot[key] = out;
  if (!memo) ctx.paras.set(props, slot);
  return out;
}

function adaptRunProps(ctx: PasteContext, props: RunProps): RunProps {
  ensureStyle(ctx, props.style);
  return adaptValue(ctx, props);
}

/** The content of a paragraph of the fragment as pieces of the document it goes into. */
function pastePieces(ctx: PasteContext, paragraph: Paragraph): Piece[] {
  const out: Piece[] = [];
  for (const inline of paragraph.children) {
    if (inline.type !== 'run') continue;
    const run: Run = inline;
    const props = adaptRunProps(ctx, run.props);
    const wrap = adaptWrap(ctx, run.wrap);
    const group = {};
    for (let k = 0; k < run.content.length; k++) {
      const code = run.fieldCode === true || (run.fieldCode !== undefined && run.fieldCode[k] === true);
      const item = adaptItem(ctx, run.content[k]);
      out.push({ item, props, wrap, group, index: -1, code, len: code ? 0 : itemLength(item) });
    }
  }
  return out;
}

/**
 * Pieces of the fragment as they stand at the caret: inside the wrappers the
 * caret is in (a link, a content control), as typed text would be. A link of
 * the fragment that would come to stand inside a link is left out: its text
 * joins the link it is pasted into.
 */
function inside(pieces: Piece[], base: Wrapper | undefined): Piece[] {
  if (!base) return pieces;
  let linked = false;
  for (let w: Wrapper | undefined = base; w; w = w.parent) if (w.kind === 'hyperlink') linked = true;
  const memo = new Map<Wrapper, Wrapper>();
  const rebase = (wrap: Wrapper | undefined, depth: number): Wrapper => {
    if (!wrap || depth > 256) return base;
    const known = memo.get(wrap);
    if (known) return known;
    const parent = rebase(wrap.parent, depth + 1);
    const out = wrap.kind === 'hyperlink' && linked ? parent : ({ ...wrap, parent } as Wrapper);
    memo.set(wrap, out);
    return out;
  };
  return pieces.map((piece) => ({ ...piece, wrap: rebase(piece.wrap, 0) }));
}

const lengthOf = (pieces: Piece[]): number => {
  let length = 0;
  for (const piece of pieces) length += piece.len;
  return length;
};

/** Attributes that name the one element they are on (w14:paraId, w14:textId): a paste makes another element, which must not repeat them. */
const IDENTITY_ATTRS = new Set<string>(['paraId', 'textId']);

/**
 * The attributes of the element of a paragraph, table, row or cell as the
 * pasted one has them: all of them (revision save ids too), without those
 * that have to be unique in a document. They go for a cut that is pasted
 * back as well: nothing tells such a paste from the second paste of the same
 * copy, and Word gives an element that has none a new one.
 */
function pasteAttrs(ctx: PasteContext, attrs: AttrList | undefined): AttrList | undefined {
  if (!attrs) return undefined;
  let out = attrs;
  if (attrs.some(([name]) => IDENTITY_ATTRS.has(localOf(name)))) out = attrs.filter(([name]) => !IDENTITY_ATTRS.has(localOf(name)));
  return out.length ? adaptAttrs(ctx, out) : undefined;
}

function pasteBlock(ctx: PasteContext, block: Block, depth: number): Block | undefined {
  if (block.type === 'paragraph') {
    const made: Paragraph = { type: 'paragraph', props: adaptParaProps(ctx, block.props, ctx.sections && depth === 0), children: rebuild(pastePieces(ctx, block)) };
    const attrs = pasteAttrs(ctx, block.attrs);
    if (attrs) made.attrs = attrs;
    return made;
  }
  if (block.type !== 'table') return undefined;
  if (depth > MAX_TABLE_DEPTH) return bad('The tables would be nested too deep');
  ensureStyle(ctx, block.props.style);
  const rows = block.rows.map((row): TableRow => {
    const cells = row.cells.map((cell): TableCell => {
      const blocks: Block[] = [];
      for (const inner of cell.blocks) {
        const made = pasteBlock(ctx, inner, depth + 1);
        if (made) blocks.push(made);
      }
      if (!blocks.length || blocks[blocks.length - 1].type !== 'paragraph') blocks.push({ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] });
      const made: TableCell = { props: adaptValue(ctx, cell.props), blocks };
      const attrs = pasteAttrs(ctx, cell.attrs);
      if (attrs) made.attrs = attrs;
      return made;
    });
    const out: TableRow = { props: adaptValue(ctx, row.props), cells };
    const attrs = pasteAttrs(ctx, row.attrs);
    if (attrs) out.attrs = attrs;
    if (row.tableProps) out.tableProps = adaptValue(ctx, row.tableProps);
    if (row.extras) out.extras = adaptValue(ctx, row.extras);
    return out;
  });
  const out: Table = { type: 'table', props: adaptValue(ctx, block.props), grid: block.grid.slice(), rows };
  const attrs = pasteAttrs(ctx, block.attrs);
  if (attrs) out.attrs = attrs;
  if (block.extras) out.extras = adaptValue(ctx, block.extras);
  return out;
}

function tableDepth(blocks: Block[], depth = 0): number {
  if (depth > 64) return depth;
  let deepest = 0;
  for (const block of blocks) {
    if (!block || block.type !== 'table' || !Array.isArray(block.rows)) continue;
    let inner = 0;
    for (const row of block.rows) for (const cell of row.cells) inner = Math.max(inner, tableDepth(cell.blocks, depth + 1));
    deepest = Math.max(deepest, 1 + inner);
  }
  return deepest;
}

/** The namespace declarations of a root element: [attribute name, uri]. */
const declarations = (attrs: AttrList): [string, string][] => attrs.filter(([name]) => name === 'xmlns' || name.indexOf('xmlns:') === 0);

function wordPrefix(attrs: AttrList): string | undefined {
  for (const [name, value] of attrs) if (name.indexOf('xmlns:') === 0 && isWordNamespace(value)) return name.slice(6);
  return undefined;
}

function checkFragment(fragment: DocFragment): void {
  const ok =
    fragment &&
    typeof fragment === 'object' &&
    Array.isArray(fragment.blocks) &&
    Array.isArray(fragment.rootAttrs) &&
    fragment.relationships instanceof Map &&
    fragment.media instanceof Map &&
    fragment.styles instanceof Map &&
    fragment.abstracts instanceof Map &&
    fragment.nums instanceof Map;
  if (!ok) bad('A fragment of a document is needed');
  for (const block of fragment.blocks) if (!block || (block.type !== 'paragraph' && block.type !== 'table')) bad('A fragment holds paragraphs and tables');
}

/**
 * The pictures of a fragment a document needs the bytes of to take the
 * fragment, and the fragment does not have (yet): none for the document it
 * came from while that still has the pictures.
 */
export function fragmentNeeds(doc: DocModel, fragment: DocFragment): string[] {
  checkFragment(fragment);
  const out: string[] = [];
  let parts: Set<string> | undefined;
  fragment.media.forEach((media, name) => {
    if (media.bytes instanceof Uint8Array) return;
    if (fragment.origin === originOf(doc)) {
      if (!parts) parts = takenByDocument(doc).parts;
      if (parts.has(name.toLowerCase())) return;
    }
    out.push(name);
  });
  return out;
}

/**
 * Puts a fragment at a position, or over a range (whose content goes first).
 * A fragment of text alone goes into the paragraph at the caret. Otherwise
 * the paragraph is cut in two there: the part before the caret takes the
 * first paragraph of the fragment and its paragraph formatting (as the end of
 * that paragraph was copied), the whole paragraphs and tables follow, and the
 * part after the caret continues the last paragraph of the fragment, with the
 * formatting it had. In a table cell a table of the fragment becomes a table
 * inside the cell. A table never touches another one and never ends a list
 * of blocks: an empty paragraph is added where it would. The caret is left
 * after what was put in.
 *
 * Throws RangeError for a place that cannot be edited; for tables that would
 * nest deeper than MAX_TABLE_DEPTH; for a fragment of a document of the other
 * kind (strict or transitional) or one that names WordprocessingML by another
 * prefix; and for a picture whose bytes neither the fragment nor the document
 * has (see fragmentNeeds). The document is unchanged then.
 */
export function insertFragment(doc: DocModel, at: DocPosition | DocRange, fragment: DocFragment, options?: InsertFragmentOptions): DocEdit {
  checkFragment(fragment);
  const range = asRange(at);
  return buildEdit(doc, (draft) => {
    const { start, end } = orderRange(range);
    resolvePosition(doc, start);
    resolvePosition(doc, end);
    const caret = samePosition(start, end) ? start : deleteSpan(draft, start, end);
    if (fragment.blocks.length === 0) return caret;
    return place(draft, caret, fragment, options);
  }, options);
}

function place(draft: Draft, caret: DocPosition, fragment: DocFragment, options: InsertFragmentOptions | undefined): DocPosition {
  const doc = draft.doc;
  const at = resolvePosition(doc, caret);
  if (inTextBox(caret.path)) bad('The content of a text box cannot be edited');
  const story = caret.story;
  const host = storyOf(doc, story).host;
  const container = caret.path.slice(0, -1);
  const nested = tableDepth(fragment.blocks);
  if (nested > 0 && container.length / 3 + nested > MAX_TABLE_DEPTH) bad('The tables would be nested too deep');

  const sameDoc = fragment.origin === originOf(doc);
  // A copy carries no revision; whatever id it holds all the same is not given to a revision made later.
  noteForeignIds(doc, fragment.blocks);
  // Part of it would be bound to styles and lists that only share an id with those it was made with.
  if (!sameDoc && fragment.definitionsCut === true) bad('The copy names more styles or lists than it carries');
  const samePart = sameDoc && fragment.part === host.part;
  const decls: [string, string][] = [];
  const source = declarations(fragment.rootAttrs);
  if (!samePart && source.length && host.rootAttrs.length) {
    if (fragment.strict !== doc.strict) bad('The copy comes from a document of another kind');
    if (wordPrefix(fragment.rootAttrs) !== wordPrefix(host.rootAttrs)) bad('The copy comes from a document that names its elements differently');
    const own = new Map<string, string>(declarations(host.rootAttrs));
    for (const [name, uri] of source) if (own.get(name) !== uri) decls.push([name, uri]);
  }
  const ctx: PasteContext = {
    doc,
    draft,
    story,
    fragment,
    sameDoc,
    samePart,
    decls,
    declPrefixes: new Set(decls.map(([name]) => (name === 'xmlns' ? '' : name.slice(6)))),
    prefixes: relPrefixes(fragment.rootAttrs),
    sections: story === 'body' && caret.path.length === 1,
    options,
    memo: new Map(),
    paras: new Map(),
    rels: new Map(),
    media: new Map(),
    nums: new Map(),
    abstracts: new Map(),
    addStyles: new Map(),
    addAbstracts: [],
    addNums: [],
    seenStyles: new Set(),
    taken: undefined,
    drawing: 0,
  } as unknown as PasteContext;
  // What a definition names by a prefix its own part declared goes with the declaration, unless the part it goes into
  // is known to declare the same. (A part that is yet to be made is not known.)
  const forDefinitions = (from: AttrList, own: AttrList | undefined): PasteContext => {
    const list: [string, string][] = [];
    if (!sameDoc) {
      const have = new Map<string, string>(own ? declarations(own) : []);
      for (const [name, uri] of declarations(from)) if (!isWordNamespace(uri) && have.get(name) !== uri) list.push([name, uri]);
    }
    return { ...ctx, decls: list, declPrefixes: new Set(list.map(([name]) => (name === 'xmlns' ? '' : name.slice(6)))), memo: new Map(), paras: new Map() };
  };
  const forStyles = forDefinitions(fragment.styleNamespaces || [], partNamespaces(doc.styles));
  const forNumbering = forDefinitions(fragment.numberingNamespaces || [], partNamespaces(doc.numbering));
  for (const each of [ctx, forStyles, forNumbering]) {
    each.forStyles = forStyles;
    each.forNumbering = forNumbering;
  }

  const paragraph = at.paragraph;
  const pieces = flatten(paragraph);
  const index = caretIndex(pieces, caret);
  const base = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
  let head = pieces.slice(0, index);
  let tail = pieces.slice(index);
  const blocks = fragment.blocks;
  const first = blocks[0];
  const last = blocks[blocks.length - 1];
  const openStart = fragment.openStart && first.type === 'paragraph';
  const openEnd = fragment.openEnd && last.type === 'paragraph';
  const definitions = (): void => {
    if (ctx.addStyles.size) {
      const styles = new Map(doc.styles.styles);
      ctx.addStyles.forEach((style, id) => styles.set(id, style));
      const next = { ...doc.styles, styles };
      // The new table stands for the same part.
      const spelt = partNamespaces(doc.styles);
      if (spelt) setPartNamespaces(next, spelt);
      draft.apply({ type: 'setStyles', styles: next });
    }
    if (ctx.addAbstracts.length || ctx.addNums.length) draft.apply({ type: 'setNumbering', abstracts: ctx.addAbstracts, nums: ctx.addNums });
  };

  if (blocks.length === 1 && openStart && openEnd) {
    const added = inside(pastePieces(ctx, first as Paragraph), base);
    definitions();
    draft.setParagraph(story, caret.path, paragraphWith(paragraph, rebuild(head.concat(added, tail))));
    return { story, path: caret.path.slice(), offset: lengthOf(head) + lengthOf(added) };
  }

  const shows = (list: Piece[]): boolean => list.some((piece) => piece.len > 0 || isAtom(piece));
  const bound = <T extends Block>(block: T): T => (paragraph.wrap ? { ...block, wrap: paragraph.wrap } : block);
  const out: Block[] = [];
  let headUsed = false;
  let from = 0;
  let to = blocks.length;
  if (openStart) {
    const made: Paragraph = { type: 'paragraph', props: adaptParaProps(ctx, (first as Paragraph).props), children: rebuild(head.concat(inside(pastePieces(ctx, first as Paragraph), base))) };
    if (paragraph.attrs && head.length) made.attrs = paragraph.attrs;
    // UNSURE: the text before the caret takes the paragraph formatting of the first paragraph pasted, whose end was
    // copied, as in Word; whether Word does so for every kind of paragraph formatting was not compared.
    out.push(bound(made));
    headUsed = true;
    from = 1;
  } else if (shows(head)) {
    headUsed = true;
  } else {
    // What stands before the caret without showing (the end of a bookmark) stays with the text after it.
    tail = head.concat(tail);
    head = [];
  }
  if (openEnd) to--;
  const middle: Block[] = [];
  for (let i = from; i < to; i++) {
    const made = pasteBlock(ctx, blocks[i], 0);
    if (made) middle.push(bound(made));
  }
  let target: Block | undefined;
  let offset = 0;
  let after = false;
  if (openEnd) {
    if (!openStart && headUsed) out.push(paragraphWith(paragraph, rebuild(head), withoutSection(paragraph.props)));
    for (const block of middle) out.push(block);
    const inner = inside(pastePieces(ctx, last as Paragraph), base);
    let props = paragraph.props;
    const left = fragment.joinedProps;
    if (options && options.moved === true && openStart && fragment.endProps && left && (props === left || sameProps(props, left))) props = adaptParaProps(ctx, fragment.endProps);
    target = paragraphWith(paragraph, rebuild(inner.concat(tail)), props, !(headUsed && head.length > 0));
    offset = lengthOf(inner);
    out.push(target);
  } else {
    const next = neighbour(at.blocks, at.index, 1);
    const follows = next >= 0 && at.blocks[next].type === 'paragraph';
    if (!openStart && headUsed && !shows(tail) && follows && paragraph.props.section === undefined) {
      // At the end of a paragraph that a paragraph follows: no paragraph is made for the nothing after the caret.
      out.push(paragraph);
      for (const block of middle) out.push(block);
      after = true;
    } else {
      if (!openStart && headUsed) out.push(paragraphWith(paragraph, rebuild(head), withoutSection(paragraph.props)));
      for (const block of middle) out.push(block);
      // The paragraph itself when nothing of it is before the caret: it keeps its place in the source of the story.
      target = head.length === 0 ? paragraph : paragraphWith(paragraph, rebuild(tail), paragraph.props, false);
      out.push(target);
    }
  }
  definitions();

  // A table never touches another one and never ends its list.
  const before = neighbour(at.blocks, at.index, -1);
  const behind = neighbour(at.blocks, at.index, 1);
  const empty = (): Paragraph => bound({ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] });
  const placed: Block[] = [];
  let previous = before >= 0 ? at.blocks[before].type : '';
  for (const block of out) {
    if (block.type === 'table' && previous === 'table') placed.push(empty());
    placed.push(block);
    previous = block.type;
  }
  if (previous === 'table' && (behind < 0 || at.blocks[behind].type === 'table')) placed.push(empty());

  draft.replace(story, container, at.index, 1, placed);
  if (target) return { story, path: container.concat([at.index + placed.indexOf(target)]), offset };
  const now = containerBlocks(doc, story, container);
  const beyond = after ? paragraphIn(now, container, at.index + placed.length, 1) : undefined;
  if (beyond) return { story, path: beyond, offset: 0 };
  const back = paragraphIn(now, container, at.index + placed.length - 1, -1);
  if (!back) return bad('The story holds no paragraph');
  return { story, path: back, offset: resolvePosition(doc, { story, path: back, offset: 0 }).length };
}
