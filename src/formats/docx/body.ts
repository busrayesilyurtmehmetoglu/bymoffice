// Block content of a WordprocessingML part (body, header, footer, note,
// comment, text box) -> blocks of the document model.
//
// A part is scanned as a stream. Each block directly under the story element
// is taken as source text, parsed into a small tree and converted; the tree is
// dropped afterwards, so memory is bounded by the largest single block, not by
// the part. The block keeps where in the source of the part it was read from.
//
// Wrapping elements do not become containers. Their content is listed flat
// and each item points to its Wrapper:
//   - hyperlink, fldSimple, ins/del/moveFrom/moveTo, sdt, smartTag, customXml,
//     dir, bdo around inline content;
//   - sdt and customXml around blocks, rows and cells;
//   - mc:AlternateContent anywhere: the first Choice whose requirements the
//     reader knows is selected, else the Fallback; the whole element stays on
//     the wrapper. The same selection is made wherever the reader looks into
//     a drawing, so that a shape and its fallback are not both read.
// A wrapper without content leaves an item of type "empty".
// Drawings, VML pictures, OLE objects and math are kept whole as trees, with
// the facts layout needs (size, position, wrapping, pictures, text boxes)
// read out.
// Comments and processing instructions become items of type "misc" in place.

import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  type AttrList,
  type Block,
  type ComplexField,
  type DrawingAnchor,
  type DrawingItem,
  type DrawingPicture,
  type DrawingPosition,
  type DrawingTransform,
  type ExtraXml,
  type FieldCharItem,
  type ImageCrop,
  type ImageRef,
  type Inline,
  type LegacyGraphicItem,
  type Paragraph,
  type ParaProps,
  type RangeMark,
  type Run,
  type RunItem,
  type RunProps,
  type SectionProps,
  type SourceSpan,
  type Table,
  type TableCell,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import { parseXml, type XmlComment, type XmlElement, type XmlNode, type XmlPI } from '../../xml';
import type { XmlScanner } from '../../xml-scan';
import {
  MAX_EMU,
  MAX_NESTING,
  MAX_TWIPS,
  CHARTEX_SPACE,
  MAX_CHART_FRAME_WARNINGS,
  MAX_XML_DEPTH,
  anyAttr,
  attr,
  attrList,
  boolAttr,
  bounded,
  findDescendants,
  hasContent,
  hasLooseContent,
  intAttr,
  isMisc,
  keep,
  kindOf,
  parseMeasure,
  parseOnOff,
  pushExtra,
  pushMisc,
  quote,
  relAttr,
  textOf,
  type PartContext,
} from './common';
import type { SourceCursor } from './source';
import { readCellProps, readParaProps, readRevision, readRowProps, readRunProps, readSectionProps, readTableProps } from './props';

/** Markup-compatibility requirements the reader can take the Choice branch for. */
const KNOWN_REQUIREMENTS = ['wps', 'wpg', 'wpc', 'wpi', 'wp14', 'w14', 'w15', 'w16', 'w16se', 'w16cid', 'w16cex', 'w16sdtdh', 'a14'];

/** Fields nested deeper than this are not assembled (their items are still kept). */
const MAX_FIELD_DEPTH = 32;

// Items without any state are shared.
const TAB: RunItem = Object.freeze({ type: 'tab' as const });
const NO_BREAK_HYPHEN: RunItem = Object.freeze({ type: 'noBreakHyphen' as const });
const SOFT_HYPHEN: RunItem = Object.freeze({ type: 'softHyphen' as const });
const LINE_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'line' as const });
const PAGE_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'page' as const });
const COLUMN_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'column' as const });
const CARRIAGE_RETURN: RunItem = Object.freeze({ type: 'break' as const, kind: 'line' as const, cr: true });

const SIMPLE_ITEMS = new Map<string, RunItem>(
  (['footnoteRef', 'endnoteRef', 'annotationRef', 'separator', 'continuationSeparator', 'lastRenderedPageBreak'] as const).map(
    (type) => [type, Object.freeze({ type })] as [string, RunItem],
  ),
);
SIMPLE_ITEMS.set('pgNum', Object.freeze({ type: 'pageNumber' as const }));
SIMPLE_ITEMS.set('tab', TAB);
SIMPLE_ITEMS.set('noBreakHyphen', NO_BREAK_HYPHEN);
SIMPLE_ITEMS.set('softHyphen', SOFT_HYPHEN);
SIMPLE_ITEMS.set('cr', CARRIAGE_RETURN);

const RANGE_MARKS = new Set<string>([
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

const REVISION_KINDS = new Set<string>(['ins', 'del', 'moveFrom', 'moveTo']);

/**
 * The array without the spare room it got while it was filled. An array grown
 * item by item reserves room for well over a dozen; a document has an array
 * per run, paragraph, cell and row, most of them with one or two items, and
 * that room would be most of what they take.
 */
const fitted = <T>(items: T[]): T[] => (items.length ? items.slice() : items);

function nested(depth: number): number {
  if (depth >= MAX_NESTING) throw new Error(`Invalid docx: content is nested deeper than ${MAX_NESTING} levels`);
  return depth + 1;
}

const isXmlSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

/**
 * Item for <w:t> and its relatives. Outer whitespace only counts with
 * xml:space="preserve"; when it is removed, the text as written is kept too.
 * The ends are found by hand: a regular expression for this back-tracks over
 * every run of whitespace inside the text.
 */
function textItem<T extends 'text' | 'deletedText' | 'instrText'>(type: T, el: XmlElement): { type: T; text: string; source?: string } {
  const source = textOf(el);
  if (el.attrs.get('xml:space') === 'preserve') return { type, text: source };
  let start = 0;
  let end = source.length;
  while (start < end && isXmlSpace(source.charCodeAt(start))) start++;
  while (end > start && isXmlSpace(source.charCodeAt(end - 1))) end--;
  if (start === 0 && end === source.length) return { type, text: source };
  return { type, text: source.slice(start, end), source };
}

/** True when a text element has nothing but text and, at most, xml:space. */
function plainText(el: XmlElement): boolean {
  if (el.attrs.size > 1 || (el.attrs.size === 1 && !el.attrs.has('xml:space'))) return false;
  for (const node of el.children) if (node.type !== 'text') return false;
  return true;
}

function imageRef(ctx: PartContext, rId: string, linked: boolean): ImageRef {
  const ref: ImageRef = { rId };
  if (linked) ref.linked = true;
  const rel = ctx.relationships.get(rId);
  if (rel) {
    ref.target = rel.part === null ? rel.target : rel.part;
    if (rel.external) ref.external = true;
  }
  return ref;
}

/** Signed EMU value from element text or an attribute; undefined when it is not a sane integer. */
function emu(text: string | undefined, min = -MAX_EMU): number | undefined {
  if (text === undefined) return undefined;
  // Element text may be padded; the ends are found by hand (see textItem).
  let start = 0;
  let end = text.length;
  while (start < end && isXmlSpace(text.charCodeAt(start))) start++;
  while (end > start && isXmlSpace(text.charCodeAt(end - 1))) end--;
  if (end - start > 11) return undefined;
  const digits = text.slice(start, end);
  if (!/^-?\d{1,10}$/.test(digits)) return undefined;
  const value = Number(digits);
  return value >= min && value <= MAX_EMU ? value : undefined;
}

const localOf = (name: string) => name.slice(name.indexOf(':') + 1);

const child = (el: XmlElement | undefined, local: string): XmlElement | undefined => {
  if (!el) return undefined;
  for (const node of el.children) {
    if (node.type === 'element' && localOf(node.name) === local) return node;
  }
  return undefined;
};

const GRAPHIC_KINDS: [suffix: string, kind: DrawingItem['kind']][] = [
  ['/picture', 'picture'],
  ['/wordprocessingShape', 'shape'],
  ['/wordprocessingGroup', 'group'],
  ['/wordprocessingCanvas', 'canvas'],
  ['/chart', 'chart'],
  ['/chartex', 'chart'],
  ['/diagram', 'diagram'],
];

const WRAP_KINDS: [local: string, wrap: NonNullable<DrawingAnchor['wrap']>][] = [
  ['wrapNone', 'none'],
  ['wrapSquare', 'square'],
  ['wrapTight', 'tight'],
  ['wrapThrough', 'through'],
  ['wrapTopAndBottom', 'topAndBottom'],
];

/** Whether a prefix a Choice requires stands for a namespace of the newer chart parts, declared on the part or on the Choice. */
function chartExPrefix(ctx: PartContext, choice: XmlElement, prefix: string): boolean {
  const own = choice.attrs.get(`xmlns:${prefix}`);
  return own !== undefined ? CHARTEX_SPACE.test(own) : ctx.prefixes.get(prefix) === 'cx';
}

/**
 * Whether every frame of a chart of the newer types in a Choice shows a chart that was read
 * (ReaderState.chartExDrawn), and there is one. Frames nested deeper than a drawing holds are not looked for.
 */
function chartExDrawn(ctx: PartContext, choice: XmlElement): boolean {
  const drawn = ctx.state.chartExDrawn;
  if (!drawn) return false;
  let found = 0;
  const stack: { el: XmlElement; depth: number }[] = [{ el: choice, depth: 0 }];
  while (stack.length > 0) {
    const { el, depth } = stack.pop() as { el: XmlElement; depth: number };
    if (localOf(el.name) === 'graphicData') {
      if (!/\/chartex$/.test(el.attrs.get('uri') || '')) continue;
      const reference = child(el, 'chart');
      const rId = reference ? anyAttr(reference, 'id') : undefined;
      const rel = rId === undefined ? undefined : ctx.relationships.get(rId);
      if (!rel || rel.external || rel.part === null || rel.type.slice(rel.type.lastIndexOf('/') + 1) !== 'chartEx') return false;
      const part = ctx.state.partName ? ctx.state.partName(rel.part) : rel.part;
      if (part === undefined || !drawn.has(part)) return false;
      found++;
      continue;
    }
    if (depth >= 16) continue;
    for (const node of el.children) if (node.type === 'element') stack.push({ el: node, depth: depth + 1 });
  }
  return found > 0;
}

/** Choices of alternate content that hold a chart of a newer type: whether the reader took them (see selectBranch()). */
const CHARTEX_CHOICES = new WeakMap<XmlElement, boolean>();

/** The branch of an <mc:AlternateContent> the reader goes with. */
export function selectBranch(ctx: PartContext, el: XmlElement): { selected: 'choice' | 'fallback' | 'none'; branch?: XmlElement } {
  let fallback: XmlElement | undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const kind = kindOf(ctx, node.name);
    if (kind === 'mc:Choice') {
      const requires = (node.attrs.get('Requires') || '').split(/\s+/).filter(Boolean);
      if (requires.length && requires.every((prefix) => KNOWN_REQUIREMENTS.indexOf(prefix) >= 0)) {
        return { selected: 'choice', branch: node };
      }
      // A chart of a newer type: the Choice is taken where its chart can be drawn, else what stands in for it.
      const extended = requires.filter((prefix) => KNOWN_REQUIREMENTS.indexOf(prefix) < 0);
      if (extended.length > 0 && extended.every((prefix) => chartExPrefix(ctx, node, prefix))) {
        // Decided once, when the document is read: the writer reads the tree again without the charts at hand.
        let drawn = CHARTEX_CHOICES.get(node);
        if (drawn === undefined && ctx.state.chartExDrawn) {
          drawn = chartExDrawn(ctx, node);
          CHARTEX_CHOICES.set(node, drawn);
        }
        if (drawn) return { selected: 'choice', branch: node };
      }
    } else if (kind === 'mc:Fallback' && !fallback) fallback = node;
  }
  return fallback ? { selected: 'fallback', branch: fallback } : { selected: 'none' };
}

/**
 * Descendants (not the element itself) with the given local name, without
 * entering matches. Of an <mc:AlternateContent> on the way only the selected
 * branch is searched, as everywhere else in the reader. Bounded in depth.
 */
export function findSelected(ctx: PartContext, el: XmlElement, local: string, out: XmlElement[] = [], depth = 0): XmlElement[] {
  if (depth > MAX_XML_DEPTH) throw new Error(`Invalid docx: XML is nested deeper than ${MAX_XML_DEPTH} levels`);
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (localOf(node.name) === local) out.push(node);
    else if (kindOf(ctx, node.name) === 'mc:AlternateContent') {
      const { branch } = selectBranch(ctx, node);
      if (branch) findSelected(ctx, branch, local, out, depth + 1);
    } else findSelected(ctx, node, local, out, depth + 1);
  }
  return out;
}

/** Child with the given local name, also when it sits in the selected branch of alternate content. */
function childThrough(ctx: PartContext, el: XmlElement | undefined, local: string, depth = 0): XmlElement | undefined {
  if (!el || depth > 8) return undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (localOf(node.name) === local) return node;
    if (kindOf(ctx, node.name) !== 'mc:AlternateContent') continue;
    const found = childThrough(ctx, selectBranch(ctx, node).branch, local, depth + 1);
    if (found) return found;
  }
  return undefined;
}

/** The first element child, taken from the selected branch when that child is alternate content. */
function firstElement(ctx: PartContext, el: XmlElement | undefined, depth = 0): XmlElement | undefined {
  if (!el || depth > 8) return undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    return kindOf(ctx, node.name) === 'mc:AlternateContent' ? firstElement(ctx, selectBranch(ctx, node).branch, depth + 1) : node;
  }
  return undefined;
}

function readPosition(ctx: PartContext, el: XmlElement | undefined, pctLocal: string): DrawingPosition | undefined {
  if (!el) return undefined;
  const position: DrawingPosition = {};
  const relativeFrom = el.attrs.get('relativeFrom');
  if (relativeFrom !== undefined) position.relativeFrom = relativeFrom;
  const offset = childThrough(ctx, el, 'posOffset');
  const value = offset ? emu(textOf(offset)) : undefined;
  if (value !== undefined) position.offset = value;
  const pct = childThrough(ctx, el, pctLocal);
  const pctValue = pct ? emu(textOf(pct)) : undefined;
  if (pctValue !== undefined) position.pctOffset = pctValue;
  const align = childThrough(ctx, el, 'align');
  if (align) position.align = textOf(align).trim();
  return position;
}

function readRelativeSize(ctx: PartContext, frame: XmlElement, local: string, pctLocal: string): { relativeFrom?: string; pct?: number } | undefined {
  const el = childThrough(ctx, frame, local);
  if (!el) return undefined;
  const size: { relativeFrom?: string; pct?: number } = {};
  const relativeFrom = el.attrs.get('relativeFrom');
  if (relativeFrom !== undefined) size.relativeFrom = relativeFrom;
  const pct = child(el, pctLocal);
  const value = pct ? emu(textOf(pct)) : undefined;
  if (value !== undefined) size.pct = value;
  return size;
}

/** <a:xfrm>; undefined when there is none or it says nothing the reader understands. */
function readTransform(el: XmlElement | undefined): DrawingTransform | undefined {
  if (!el) return undefined;
  const transform: DrawingTransform = {};
  let any = false;
  const set = (field: 'x' | 'y' | 'width' | 'height' | 'rotation' | 'childX' | 'childY' | 'childWidth' | 'childHeight', value: number | undefined) => {
    if (value === undefined) return;
    transform[field] = value;
    any = true;
  };
  const offset = child(el, 'off');
  set('x', emu(offset?.attrs.get('x')));
  set('y', emu(offset?.attrs.get('y')));
  const extent = child(el, 'ext');
  set('width', emu(extent?.attrs.get('cx'), 0));
  set('height', emu(extent?.attrs.get('cy'), 0));
  set('rotation', emu(el.attrs.get('rot')));
  for (const flip of ['flipH', 'flipV'] as const) {
    const text = el.attrs.get(flip);
    const value = text === undefined ? undefined : parseOnOff(text);
    if (value === undefined) continue;
    transform[flip] = value;
    any = true;
  }
  const childOffset = child(el, 'chOff');
  set('childX', emu(childOffset?.attrs.get('x')));
  set('childY', emu(childOffset?.attrs.get('y')));
  const childExtent = child(el, 'chExt');
  set('childWidth', emu(childExtent?.attrs.get('cx'), 0));
  set('childHeight', emu(childExtent?.attrs.get('cy'), 0));
  return any ? transform : undefined;
}

const CROP_SIDES = [
  ['l', 'left'],
  ['t', 'top'],
  ['r', 'right'],
  ['b', 'bottom'],
] as const;

/** The picture a <blipFill> holds: the image and the cropping stated next to it, in the same fill. */
function readPicture(
  ctx: PartContext,
  fill: XmlElement | undefined,
  xfrm: XmlElement | undefined,
  groups: DrawingTransform[],
): DrawingPicture | undefined {
  const blip = child(fill, 'blip');
  if (!blip) return undefined;
  const embed = anyAttr(blip, 'embed');
  const link = embed ? undefined : anyAttr(blip, 'link');
  if (!embed && !link) return undefined;
  const picture: DrawingPicture = { image: imageRef(ctx, (embed || link) as string, !embed) };
  const rect = child(fill, 'srcRect');
  if (rect) {
    const crop: ImageCrop = {};
    let any = false;
    for (const [attribute, side] of CROP_SIDES) {
      const text = rect.attrs.get(attribute);
      if (text === undefined || !/^-?\d{1,7}$/.test(text)) continue;
      crop[side] = Number(text);
      any = true;
    }
    if (any) picture.crop = crop;
  }
  const transform = readTransform(xfrm);
  if (transform) picture.transform = transform;
  if (groups.length) picture.groups = groups;
  return picture;
}

/**
 * Every picture under `el`, in document order: pictures of their own
 * (<pic:pic>) and images that fill a shape. Text boxes are not entered (the
 * drawings in them are read with their paragraphs), and of alternate content
 * only the selected branch is. `groups` are the transforms of the groups
 * around `el`, outermost first.
 */
function collectPictures(ctx: PartContext, el: XmlElement, groups: DrawingTransform[], out: DrawingPicture[], depth: number): void {
  if (depth > MAX_XML_DEPTH) throw new Error(`Invalid docx: XML is nested deeper than ${MAX_XML_DEPTH} levels`);
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const local = localOf(node.name);
    if (local === 'txbxContent') continue;
    if (kindOf(ctx, node.name) === 'mc:AlternateContent') {
      const { branch } = selectBranch(ctx, node);
      if (branch) collectPictures(ctx, branch, groups, out, depth + 1);
    } else if (local === 'pic') {
      const picture = readPicture(ctx, child(node, 'blipFill'), child(child(node, 'spPr'), 'xfrm'), groups);
      if (!picture) continue;
      const properties = child(child(node, 'nvPicPr'), 'cNvPr');
      const name = properties?.attrs.get('name');
      if (name !== undefined) picture.name = name;
      const description = properties?.attrs.get('descr');
      if (description !== undefined) picture.description = description;
      out.push(picture);
    } else if (local === 'blipFill') {
      // The fill of a shape: `el` is the shape properties, which hold the transform too.
      const picture = readPicture(ctx, node, child(el, 'xfrm'), groups);
      if (!picture) continue;
      picture.fill = true;
      out.push(picture);
    } else if (local === 'wgp' || local === 'grpSp') {
      const transform = readTransform(child(child(node, 'grpSpPr'), 'xfrm'));
      collectPictures(ctx, node, transform ? groups.concat([transform]) : groups, out, depth + 1);
    } else collectPictures(ctx, node, groups, out, depth + 1);
  }
}

/**
 * The <w:txbxContent> elements a drawing (<w:drawing>) or legacy graphic (<w:pict>, <w:object>)
 * holds as its text boxes, in order: of a drawing those of its graphic (<a:graphicData>), of a
 * legacy graphic all of them. The one count of the text boxes of a graphic, for the reader and
 * the writer alike.
 */
export function graphicTextBoxContents(ctx: PartContext, type: 'drawing' | 'pict' | 'object', el: XmlElement): XmlElement[] {
  if (type !== 'drawing') return findSelected(ctx, el, 'txbxContent');
  const data = child(child(child(el, 'inline') || child(el, 'anchor'), 'graphic'), 'graphicData');
  return data ? findSelected(ctx, data, 'txbxContent') : [];
}

function textBoxes(ctx: PartContext, contents: XmlElement[], depth: number): Block[][] | undefined {
  if (!contents.length) return undefined;
  const inner = nested(depth);
  return contents.map((content) => {
    const blocks: Block[] = [];
    readBlocks(ctx, content.children, blocks, undefined, inner);
    linkFields(blocks);
    return blocks;
  });
}

const DISTANCES = ['distT', 'distB', 'distL', 'distR'] as const;

/**
 * The chart part the <c:chart> of a frame names, under the name the package
 * has it: a part of this package its relationship says is a chart. A frame that leads out of the file, to a part
 * that is not there, or to something that is no chart has none (it shows a
 * placeholder); the first two are said in the warnings.
 */
function chartPart(ctx: PartContext, el: XmlElement | undefined): string | undefined {
  if (!el || localOf(el.name) !== 'chart') return undefined;
  const rId = anyAttr(el, 'id');
  const rel = rId === undefined ? undefined : ctx.relationships.get(rId);
  const type = rel ? rel.type.slice(rel.type.lastIndexOf('/') + 1) : '';
  if (!rel || (type !== 'chart' && type !== 'chartEx')) return undefined;
  const state = ctx.state;
  const warn = (text: string) => {
    if (state.chartWarnings++ < MAX_CHART_FRAME_WARNINGS) state.warnings.push(text);
  };
  if (rel.external || rel.part === null) {
    warn(`${ctx.part}: a chart with an external target is not shown`);
    return undefined;
  }
  // The part as the package spells it: frames that spell one part differently show one chart, read once.
  const part = state.partName ? state.partName(rel.part) : rel.part;
  if (part === undefined) {
    warn(`The chart part ${quote(rel.part)} is missing`);
    return undefined;
  }
  state.chartParts.add(part);
  return part;
}

function readDrawing(ctx: PartContext, el: XmlElement, depth: number): RunItem {
  const frame = child(el, 'inline') || child(el, 'anchor');
  if (!frame) return { type: 'unknown', xml: keep(ctx, el, 'r') };
  const anchored = localOf(frame.name) === 'anchor';
  const item: DrawingItem = { type: 'drawing', placement: anchored ? 'anchor' : 'inline', kind: 'other', xml: el };

  const extent = child(frame, 'extent');
  const width = emu(extent?.attrs.get('cx'), 0);
  const height = emu(extent?.attrs.get('cy'), 0);
  if (width !== undefined) item.width = width;
  if (height !== undefined) item.height = height;
  const docPr = child(frame, 'docPr');
  const name = docPr?.attrs.get('name');
  if (name !== undefined) item.name = name;
  const description = docPr?.attrs.get('descr');
  if (description !== undefined) item.description = description;
  const effect = child(frame, 'effectExtent');
  if (effect) {
    const room: NonNullable<DrawingItem['effectExtent']> = {};
    for (const [attribute, side] of CROP_SIDES) {
      const value = emu(effect.attrs.get(attribute));
      if (value !== undefined) room[side] = value;
    }
    item.effectExtent = room;
  }

  // Both kinds of frame state distances from the text; they are kept with the kind they belong to.
  const distances: NonNullable<DrawingItem['inline']> = {};
  let anyDistance = false;
  for (const side of DISTANCES) {
    const value = emu(frame.attrs.get(side), 0);
    if (value === undefined) continue;
    distances[side] = value;
    anyDistance = true;
  }
  if (!anchored && anyDistance) item.inline = distances;

  if (anchored) {
    const anchor: DrawingAnchor = distances;
    const relativeHeight = emu(frame.attrs.get('relativeHeight'), 0);
    if (relativeHeight !== undefined) anchor.relativeHeight = relativeHeight;
    for (const flag of ['behindDoc', 'allowOverlap', 'layoutInCell', 'locked'] as const) {
      const text = frame.attrs.get(flag);
      const value = text === undefined ? undefined : parseOnOff(text);
      if (value !== undefined) anchor[flag] = value;
    }
    if (parseOnOff(frame.attrs.get('simplePos') || '0')) {
      const simple = child(frame, 'simplePos');
      const x = emu(simple?.attrs.get('x'));
      const y = emu(simple?.attrs.get('y'));
      if (x !== undefined && y !== undefined) anchor.simplePos = { x, y };
    }
    const horizontal = readPosition(ctx, child(frame, 'positionH'), 'pctPosHOffset');
    if (horizontal) anchor.horizontal = horizontal;
    const vertical = readPosition(ctx, child(frame, 'positionV'), 'pctPosVOffset');
    if (vertical) anchor.vertical = vertical;
    for (const [local, wrap] of WRAP_KINDS) {
      const wrapElement = child(frame, local);
      if (!wrapElement) continue;
      anchor.wrap = wrap;
      const wrapText = wrapElement.attrs.get('wrapText');
      if (wrapText !== undefined) anchor.wrapText = wrapText;
      const wrapDist: NonNullable<DrawingAnchor['wrapDist']> = {};
      let any = false;
      for (const [attribute, side] of [['distT', 'top'], ['distB', 'bottom'], ['distL', 'left'], ['distR', 'right']] as const) {
        const value = emu(wrapElement.attrs.get(attribute), 0);
        if (value === undefined) continue;
        wrapDist[side] = value;
        any = true;
      }
      if (any) anchor.wrapDist = wrapDist;
      const polygon = child(wrapElement, 'wrapPolygon');
      if (polygon) {
        const points: [number, number][] = [];
        for (const point of polygon.children) {
          if (point.type !== 'element') continue;
          const x = emu(point.attrs.get('x'));
          const y = emu(point.attrs.get('y'));
          if (x !== undefined && y !== undefined) points.push([x, y]);
        }
        anchor.wrapPolygon = { points };
        const edited = polygon.attrs.get('edited');
        const value = edited === undefined ? undefined : parseOnOff(edited);
        if (value !== undefined) anchor.wrapPolygon.edited = value;
      }
      break;
    }
    const sizeRelH = readRelativeSize(ctx, frame, 'sizeRelH', 'pctWidth');
    if (sizeRelH) anchor.sizeRelH = sizeRelH;
    const sizeRelV = readRelativeSize(ctx, frame, 'sizeRelV', 'pctHeight');
    if (sizeRelV) anchor.sizeRelV = sizeRelV;
    item.anchor = anchor;
  }

  const data = child(child(frame, 'graphic'), 'graphicData');
  const uri = data?.attrs.get('uri') || '';
  for (const [suffix, kind] of GRAPHIC_KINDS) if (uri.endsWith(suffix)) item.kind = kind;
  if (data) {
    // The outermost object: a picture or shape (properties in spPr) or a group (grpSpPr).
    const top = firstElement(ctx, data);
    if (item.kind === 'chart') {
      const part = chartPart(ctx, top);
      if (part !== undefined) item.chartPart = part;
    }
    const transform = readTransform(child(child(top, 'spPr'), 'xfrm') || child(child(top, 'grpSpPr'), 'xfrm'));
    if (transform) item.transform = transform;

    const pictures: DrawingPicture[] = [];
    collectPictures(ctx, data, [], pictures, 0);
    if (pictures.length) {
      item.images = pictures;
      item.image = pictures[0].image;
      if (pictures[0].crop) item.crop = pictures[0].crop;
    }
    const boxes = textBoxes(ctx, graphicTextBoxContents(ctx, 'drawing', el), depth);
    if (boxes) item.textBoxes = boxes;
  }
  return item;
}

function readLegacyGraphic(ctx: PartContext, el: XmlElement, type: 'pict' | 'object', depth: number): LegacyGraphicItem {
  const item: LegacyGraphicItem = { type, xml: keep(ctx, el, 'r') };
  const imageData = findSelected(ctx, el, 'imagedata')[0];
  const rId = imageData ? anyAttr(imageData, 'id') : undefined;
  if (rId) item.image = imageRef(ctx, rId, false);
  const boxes = textBoxes(ctx, graphicTextBoxContents(ctx, type, el), depth);
  if (boxes) item.textBoxes = boxes;
  return item;
}

function readRunItem(ctx: PartContext, el: XmlElement, out: RunItem[], depth: number): void {
  const kind = kindOf(ctx, el.name);
  const unknown = () => out.push({ type: 'unknown', xml: keep(ctx, el, 'r') });
  switch (kind) {
    case 't':
      if (!plainText(el)) unknown();
      else out.push(textItem('text', el));
      return;
    case 'delText':
      if (!plainText(el)) unknown();
      else out.push(textItem('deletedText', el));
      return;
    case 'instrText':
    case 'delInstrText':
      if (!plainText(el)) unknown();
      else if (kind === 'instrText') out.push(textItem('instrText', el));
      else out.push({ ...textItem('instrText', el), deleted: true });
      return;
    case 'br': {
      ctx.hits = 0;
      const type = attr(ctx, el, 'type');
      const clear = attr(ctx, el, 'clear');
      if (ctx.hits !== el.attrs.size || el.children.length) unknown();
      else if (clear !== undefined && (type === undefined || type === 'textWrapping')) out.push({ type: 'break', kind: 'line', clear });
      else if (clear !== undefined) unknown();
      else if (type === undefined || type === 'textWrapping') out.push(LINE_BREAK);
      else if (type === 'page') out.push(PAGE_BREAK);
      else if (type === 'column') out.push(COLUMN_BREAK);
      else unknown();
      return;
    }
    case 'sym': {
      ctx.hits = 0;
      const font = attr(ctx, el, 'font');
      const code = attr(ctx, el, 'char');
      if (code === undefined || !/^[0-9A-Fa-f]{1,4}$/.test(code) || ctx.hits !== el.attrs.size || hasContent(el)) unknown();
      else {
        const item: RunItem = { type: 'symbol', char: String.fromCharCode(parseInt(code, 16)) };
        if (font !== undefined) item.font = font;
        out.push(item);
      }
      return;
    }
    case 'fldChar': {
      ctx.hits = 0;
      const charType = attr(ctx, el, 'fldCharType');
      const dirty = boolAttr(ctx, el, 'dirty');
      const locked = boolAttr(ctx, el, 'fldLock');
      if (charType !== 'begin' && charType !== 'separate' && charType !== 'end') {
        unknown();
        return;
      }
      const item: FieldCharItem = { type: 'fieldChar', charType };
      if (dirty !== undefined) item.dirty = dirty;
      if (locked !== undefined) item.locked = locked;
      // A field character must stay one whatever else it says: taken out of the
      // stream it would leave its field open. What the item has no field for
      // (another attribute, a value that is not a boolean, a comment) is kept with it.
      if (ctx.hits !== el.attrs.size || hasLooseContent(el)) item.xml = keep(ctx, el, 'r', true);
      for (const node of el.children) {
        if (node.type === 'element') (item.data || (item.data = [])).push(keep(ctx, node, 'fldChar'));
      }
      out.push(item);
      return;
    }
    case 'footnoteReference':
    case 'endnoteReference': {
      ctx.hits = 0;
      const id = intAttr(ctx, el, 'id', -0x80000000, 0x7fffffff);
      const customMark = boolAttr(ctx, el, 'customMarkFollows');
      if (id === undefined || ctx.hits !== el.attrs.size || hasContent(el)) unknown();
      else if (customMark === undefined) out.push({ type: kind, id });
      else out.push({ type: kind, id, customMark });
      return;
    }
    case 'commentReference': {
      ctx.hits = 0;
      const id = intAttr(ctx, el, 'id', -0x80000000, 0x7fffffff);
      if (id === undefined || ctx.hits !== el.attrs.size || hasContent(el)) unknown();
      else out.push({ type: 'commentReference', id });
      return;
    }
    case 'ptab': {
      ctx.hits = 0;
      const item: RunItem = { type: 'positionalTab' };
      const alignment = attr(ctx, el, 'alignment');
      if (alignment !== undefined) item.alignment = alignment;
      const relativeTo = attr(ctx, el, 'relativeTo');
      if (relativeTo !== undefined) item.relativeTo = relativeTo;
      const leader = attr(ctx, el, 'leader');
      if (leader !== undefined) item.leader = leader;
      if (ctx.hits !== el.attrs.size || hasContent(el)) unknown();
      else out.push(item);
      return;
    }
    case 'drawing':
      out.push(readDrawing(ctx, el, depth));
      return;
    case 'pict':
    case 'object':
      out.push(readLegacyGraphic(ctx, el, kind, depth));
      return;
    case 'mc:AlternateContent': {
      const { selected, branch } = selectBranch(ctx, el);
      const items: RunItem[] = [];
      const inner = nested(depth);
      if (branch) for (const node of branch.children) if (node.type === 'element') readRunItem(ctx, node, items, inner);
      out.push({ type: 'alternate', selected, items, xml: el });
      return;
    }
  }
  const simple = SIMPLE_ITEMS.get(kind);
  if (simple && el.attrs.size === 0 && el.children.length === 0) out.push(simple);
  else unknown();
}

function readRun(ctx: PartContext, el: XmlElement, wrap: Wrapper | undefined, depth: number): Run {
  let props: RunProps | undefined;
  const content: RunItem[] = [];
  for (const node of el.children) {
    if (node.type !== 'element') {
      if (isMisc(node)) content.push({ type: 'misc', node });
    } else if (props === undefined && kindOf(ctx, node.name) === 'rPr') props = readRunProps(ctx, node);
    else readRunItem(ctx, node, content, depth);
  }
  const run: Run = { type: 'run', props: props || EMPTY_RUN_PROPS, content: fitted(content) };
  if (wrap) run.wrap = wrap;
  const attrs = attrList(el);
  if (attrs) run.attrs = attrs;
  return run;
}

function readRangeMark(ctx: PartContext, el: XmlElement, type: RangeMark['type'], wrap: Wrapper | undefined): RangeMark | undefined {
  let id: string | undefined;
  let name: string | undefined;
  let attrs: AttrList | undefined;
  for (const [attribute, value] of el.attrs) {
    const kind = kindOf(ctx, attribute);
    if (kind === 'id' && id === undefined) id = value;
    else if (kind === 'name' && name === undefined) name = value;
    else (attrs || (attrs = [])).push([attribute, value]);
  }
  if (id === undefined || el.children.length) return undefined;
  const mark: RangeMark = { type, id };
  if (name !== undefined) mark.name = name;
  if (attrs) mark.attrs = attrs;
  if (wrap) mark.wrap = wrap;
  return mark;
}

/** Attributes of a wrapper element that its fields do not hold. */
function restAttrs(ctx: PartContext, el: XmlElement, known: readonly string[]): AttrList | undefined {
  let attrs: AttrList | undefined;
  for (const [name, value] of el.attrs) {
    if (known.indexOf(kindOf(ctx, name)) < 0) (attrs || (attrs = [])).push([name, value]);
  }
  return attrs;
}

/**
 * Builds the wrapper for a wrapping element and returns it with the nodes
 * that are its content: elements, and the comments and processing
 * instructions among them. Undefined when the element is not a wrapper.
 */
function readWrapper(
  ctx: PartContext,
  el: XmlElement,
  kind: string,
  parent: Wrapper | undefined,
  context: string,
): { wrapper: Wrapper; content: (XmlElement | XmlComment | XmlPI)[] } | undefined {
  let wrapper: Wrapper;
  const content: (XmlElement | XmlComment | XmlPI)[] = [];
  let extras: ExtraXml[] | undefined;
  let index = -1;
  const takeAll = (nodes: XmlNode[]) => {
    for (const node of nodes) if (node.type === 'element' || isMisc(node)) content.push(node);
  };
  const children = (special: (kind: string, node: XmlElement) => boolean) => {
    for (const node of el.children) {
      if (node.type !== 'element') {
        if (isMisc(node)) content.push(node);
        continue;
      }
      index++;
      if (!special(kindOf(ctx, node.name), node)) content.push(node);
    }
  };

  if (kind === 'sdt') {
    const sdt: Wrapper = { kind: 'sdt' };
    for (const node of el.children) {
      if (node.type !== 'element') {
        // Not content: the content of a content control sits one level down.
        if (isMisc(node)) sdt.misc = pushMisc(sdt.misc, index + 1, node);
        continue;
      }
      index++;
      const name = kindOf(ctx, node.name);
      // The properties of a content control are kept, not interpreted.
      if (name === 'sdtPr' && !sdt.pr) sdt.pr = keep(ctx, node, 'sdt');
      else if (name === 'sdtEndPr' && !sdt.endPr) sdt.endPr = keep(ctx, node, 'sdt');
      else if (name === 'sdtContent') takeAll(node.children);
      else extras = pushExtra(ctx, extras, index, node, context, false);
    }
    if (el.attrs.size) extras = pushExtra(ctx, extras, -1, el, context, true);
    wrapper = sdt;
  } else if (kind === 'smartTag' || kind === 'customXml') {
    const tag: Wrapper = { kind };
    const attrs = attrList(el);
    if (attrs) tag.attrs = attrs;
    children((name, node) => {
      if ((name === 'smartTagPr' || name === 'customXmlPr') && !tag.pr) tag.pr = keep(ctx, node, kind);
      else return false;
      return true;
    });
    wrapper = tag;
  } else if (kind === 'mc:AlternateContent') {
    const { selected, branch } = selectBranch(ctx, el);
    if (branch) takeAll(branch.children);
    wrapper = { kind: 'alternate', selected, xml: el };
  } else if (context !== 'p') {
    // The remaining wrappers only exist around inline content.
    return undefined;
  } else if (kind === 'hyperlink') {
    const link: Wrapper = { kind: 'hyperlink' };
    const rId = relAttr(ctx, el, 'id');
    if (rId !== undefined) {
      link.rId = rId;
      const rel = ctx.relationships.get(rId);
      if (rel) link.target = rel.target;
      else ctx.state.warnings.push(`${ctx.part}: hyperlink relationship is missing`);
    }
    const anchor = attr(ctx, el, 'anchor');
    if (anchor !== undefined) link.anchor = anchor;
    const tooltip = attr(ctx, el, 'tooltip');
    if (tooltip !== undefined) link.tooltip = tooltip;
    const attrs = restAttrs(ctx, el, ['r:id', 'anchor', 'tooltip']);
    if (attrs) link.attrs = attrs;
    children(() => false);
    wrapper = link;
  } else if (REVISION_KINDS.has(kind)) {
    const revision: Wrapper = { kind: kind as 'ins' | 'del' | 'moveFrom' | 'moveTo', ...readRevision(ctx, el) };
    const attrs = restAttrs(ctx, el, revision.id === undefined ? ['author', 'date'] : ['id', 'author', 'date']);
    if (attrs) revision.attrs = attrs;
    children(() => false);
    wrapper = revision;
  } else if (kind === 'fldSimple') {
    const field: Wrapper = { kind: 'simpleField', instruction: attr(ctx, el, 'instr') || '' };
    const dirty = boolAttr(ctx, el, 'dirty');
    if (dirty !== undefined) field.dirty = dirty;
    const locked = boolAttr(ctx, el, 'fldLock');
    if (locked !== undefined) field.locked = locked;
    const attrs = restAttrs(ctx, el, ['instr'].concat(dirty === undefined ? [] : ['dirty'], locked === undefined ? [] : ['fldLock']));
    if (attrs) field.attrs = attrs;
    children((name, node) => {
      if (name !== 'fldData') return false;
      extras = pushExtra(ctx, extras, index, node, context, false);
      return true;
    });
    wrapper = field;
  } else if (kind === 'dir' || kind === 'bdo') {
    const direction: Wrapper = { kind };
    const val = attr(ctx, el, 'val');
    if (val !== undefined) direction.val = val;
    const attrs = restAttrs(ctx, el, ['val']);
    if (attrs) direction.attrs = attrs;
    children(() => false);
    wrapper = direction;
  } else return undefined;

  if (parent) wrapper.parent = parent;
  if (extras) wrapper.extras = extras;
  return { wrapper, content };
}

/** Plain text of a math formula: the <m:t> texts in order. */
function mathText(el: XmlElement): string {
  let text = '';
  for (const t of findDescendants(el, 't')) text += textOf(t);
  return text;
}

function readInline(ctx: PartContext, el: XmlElement, out: Inline[], wrap: Wrapper | undefined, depth: number): void {
  const kind = kindOf(ctx, el.name);
  if (kind === 'r') {
    out.push(readRun(ctx, el, wrap, depth));
    return;
  }
  if (RANGE_MARKS.has(kind)) {
    const mark = readRangeMark(ctx, el, kind as RangeMark['type'], wrap);
    if (mark) {
      out.push(mark);
      return;
    }
  } else if (kind === 'proofErr') {
    ctx.hits = 0;
    const type = attr(ctx, el, 'type');
    if (type !== undefined && ctx.hits === el.attrs.size && !el.children.length) {
      out.push(wrap ? { type: 'proofError', kind: type, wrap } : { type: 'proofError', kind: type });
      return;
    }
  } else if (kind === 'm:oMath' || kind === 'm:oMathPara') {
    // Known and kept whole; not counted among the unknown elements.
    const math: Inline = { type: 'math', text: mathText(el), xml: el };
    if (wrap) math.wrap = wrap;
    out.push(math);
    return;
  } else {
    const wrapped = readWrapper(ctx, el, kind, wrap, 'p');
    if (wrapped) {
      const inner = nested(depth);
      const before = out.length;
      for (const node of wrapped.content) {
        if (node.type === 'element') readInline(ctx, node, out, wrapped.wrapper, inner);
        else out.push({ type: 'misc', node, wrap: wrapped.wrapper });
      }
      if (out.length === before) out.push({ type: 'empty', wrap: wrapped.wrapper });
      return;
    }
  }
  const unknown: Inline = { type: 'unknown', xml: keep(ctx, el, 'p') };
  if (wrap) unknown.wrap = wrap;
  out.push(unknown);
}

function readParagraph(ctx: PartContext, el: XmlElement, wrap: Wrapper | undefined, depth: number): Paragraph {
  let props: ParaProps | undefined;
  const children: Inline[] = [];
  for (const node of el.children) {
    if (node.type !== 'element') {
      if (isMisc(node)) children.push({ type: 'misc', node });
    } else if (props === undefined && kindOf(ctx, node.name) === 'pPr') props = readParaProps(ctx, node);
    else readInline(ctx, node, children, undefined, depth);
  }
  const paragraph: Paragraph = { type: 'paragraph', props: props || EMPTY_PARA_PROPS, children: fitted(children) };
  if (wrap) paragraph.wrap = wrap;
  const attrs = attrList(el);
  if (attrs) paragraph.attrs = attrs;
  return paragraph;
}

function readCell(ctx: PartContext, el: XmlElement, wrap: Wrapper | undefined, depth: number): TableCell {
  let propsElement: XmlElement | undefined;
  const blocks: Block[] = [];
  for (const node of el.children) {
    if (node.type !== 'element') {
      if (isMisc(node)) blocks.push({ type: 'misc', node });
    } else if (!propsElement && kindOf(ctx, node.name) === 'tcPr') propsElement = node;
    else readBlock(ctx, node, blocks, undefined, depth);
  }
  const cell: TableCell = { props: readCellProps(ctx, propsElement), blocks: fitted(blocks) };
  if (wrap) cell.wrap = wrap;
  const attrs = attrList(el);
  if (attrs) cell.attrs = attrs;
  return cell;
}

/**
 * Cells of a row, looking through cell-level wrappers. `index` is the place
 * of the element among the children of the row or, inside a wrapper, among
 * the content of that wrapper.
 */
function readCells(ctx: PartContext, el: XmlElement, row: TableRow, index: number, wrap: Wrapper | undefined, depth: number): void {
  const kind = kindOf(ctx, el.name);
  if (kind === 'tc') {
    row.cells.push(readCell(ctx, el, wrap, depth));
    return;
  }
  const wrapped = readWrapper(ctx, el, kind, wrap, 'tr');
  if (wrapped && wrapped.content.length) {
    const inner = nested(depth);
    let position = -1;
    for (const node of wrapped.content) {
      if (node.type === 'element') readCells(ctx, node, row, ++position, wrapped.wrapper, inner);
      else row.misc = pushMisc(row.misc, position + 1, node, wrapped.wrapper);
    }
    return;
  }
  // Not a cell: kept in its place, with the wrapper it sits in.
  row.extras = pushExtra(ctx, row.extras, index, el, 'tr', false, wrap);
}

function readRow(ctx: PartContext, el: XmlElement, wrap: Wrapper | undefined, depth: number): TableRow {
  const row: TableRow = { props: {}, cells: [] };
  let propsElement: XmlElement | undefined;
  let index = -1;
  for (const node of el.children) {
    if (node.type !== 'element') {
      if (isMisc(node)) row.misc = pushMisc(row.misc, index + 1, node);
      continue;
    }
    index++;
    const kind = kindOf(ctx, node.name);
    if (kind === 'trPr' && !propsElement) propsElement = node;
    else if (kind === 'tblPrEx' && !row.tableProps) row.tableProps = readTableProps(ctx, node);
    else readCells(ctx, node, row, index, undefined, depth);
  }
  row.cells = fitted(row.cells);
  row.props = readRowProps(ctx, propsElement);
  if (wrap) row.wrap = wrap;
  const attrs = attrList(el);
  if (attrs) row.attrs = attrs;
  return row;
}

/** Rows of a table, looking through row-level wrappers; `index` as for readCells. */
function readRows(ctx: PartContext, el: XmlElement, table: Table, index: number, wrap: Wrapper | undefined, depth: number): void {
  const kind = kindOf(ctx, el.name);
  if (kind === 'tr') {
    table.rows.push(readRow(ctx, el, wrap, depth));
    return;
  }
  const wrapped = readWrapper(ctx, el, kind, wrap, 'tbl');
  if (wrapped && wrapped.content.length) {
    const inner = nested(depth);
    let position = -1;
    for (const node of wrapped.content) {
      if (node.type === 'element') readRows(ctx, node, table, ++position, wrapped.wrapper, inner);
      else table.misc = pushMisc(table.misc, position + 1, node, wrapped.wrapper);
    }
    return;
  }
  table.extras = pushExtra(ctx, table.extras, index, el, 'tbl', false, wrap);
}

/** <w:tblGridChange w:id=""><w:tblGrid><w:gridCol/>...</w:tblGrid></w:tblGridChange>, when it holds nothing else. */
function readGridChange(ctx: PartContext, el: XmlElement): { id?: number; grid: number[] } | undefined {
  if (hasLooseContent(el)) return undefined;
  ctx.hits = 0;
  const id = intAttr(ctx, el, 'id', 0, 0x7fffffff);
  if (ctx.hits !== el.attrs.size) return undefined;
  let inner: XmlElement | undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (inner || kindOf(ctx, node.name) !== 'tblGrid') return undefined;
    inner = node;
  }
  // The schema asks for the grid: a record without one says nothing that could be brought back.
  if (!inner || inner.attrs.size || hasLooseContent(inner)) return undefined;
  const grid: number[] = [];
  for (const col of inner.children) {
    if (col.type !== 'element') continue;
    if (kindOf(ctx, col.name) !== 'gridCol' || col.children.length) return undefined;
    ctx.hits = 0;
    const text = attr(ctx, col, 'w');
    const width = text === undefined ? undefined : parseMeasure(text, 20, 0, MAX_TWIPS);
    if (width === undefined || ctx.hits !== col.attrs.size) return undefined;
    grid.push(width);
  }
  if (grid.length > 1000) return undefined;
  const out: { id?: number; grid: number[] } = { grid };
  if (id !== undefined) out.id = id;
  return out;
}

function readTable(ctx: PartContext, el: XmlElement, wrap: Wrapper | undefined, depth: number): Table {
  const table: Table = { type: 'table', props: {}, grid: [], rows: [] };
  const inner = nested(depth);
  let propsElement: XmlElement | undefined;
  let gridSeen = false;
  let index = -1;
  for (const node of el.children) {
    if (node.type !== 'element') {
      if (isMisc(node)) table.misc = pushMisc(table.misc, index + 1, node);
      continue;
    }
    index++;
    const kind = kindOf(ctx, node.name);
    if (kind === 'tblPr' && !propsElement) propsElement = node;
    else if (kind === 'tblGrid' && !gridSeen) {
      gridSeen = true;
      let plain = node.attrs.size === 0 && !hasLooseContent(node);
      let change: XmlElement | undefined;
      for (const col of node.children) {
        if (col.type !== 'element') continue;
        // The grid as it was before a tracked change, when it is the last child and holds plain columns.
        if (!change && kindOf(ctx, col.name) === 'tblGridChange') {
          const was = readGridChange(ctx, col);
          if (was) {
            change = col;
            table.gridChange = was;
            continue;
          }
        }
        // Something follows the record: the grid is not what the schema says.
        if (change) plain = false;
        ctx.hits = 0;
        const text = kindOf(ctx, col.name) === 'gridCol' ? attr(ctx, col, 'w') : undefined;
        const width = text === undefined ? undefined : parseMeasure(text, 20, 0, MAX_TWIPS);
        if (kindOf(ctx, col.name) === 'gridCol') table.grid.push(width === undefined ? 0 : width);
        if (width === undefined || ctx.hits !== col.attrs.size || col.children.length) plain = false;
      }
      if (!plain) {
        // Kept as it was read, record and all: the record is not in the model too, or it would be two changes.
        delete table.gridChange;
        table.extras = pushExtra(ctx, table.extras, index, node, 'tbl', true);
      }
    } else readRows(ctx, node, table, index, undefined, inner);
  }
  table.rows = fitted(table.rows);
  table.grid = fitted(table.grid);
  table.props = readTableProps(ctx, propsElement);
  if (wrap) table.wrap = wrap;
  const attrs = attrList(el);
  if (attrs) table.attrs = attrs;
  return table;
}

/** Converts the nodes of a block container (a cell, a text box, the content of a wrapper) and appends the results to `out`. */
function readBlocks(ctx: PartContext, nodes: readonly XmlNode[], out: Block[], wrap: Wrapper | undefined, depth: number): void {
  for (const node of nodes) {
    if (node.type === 'element') readBlock(ctx, node, out, wrap, depth);
    else if (isMisc(node)) out.push(wrap ? { type: 'misc', node, wrap } : { type: 'misc', node });
  }
}

/**
 * Converts one block-level element and appends the result to `out`. `span`
 * says where the element is in the source of the part, when it sits directly
 * in a story.
 */
export function readBlock(
  ctx: PartContext,
  el: XmlElement,
  out: Block[],
  wrap: Wrapper | undefined,
  depth: number,
  span?: SourceSpan,
): void {
  const kind = kindOf(ctx, el.name);
  let block: Block;
  if (kind === 'p') block = readParagraph(ctx, el, wrap, depth);
  else if (kind === 'tbl') block = readTable(ctx, el, wrap, depth);
  else {
    let mark: RangeMark | undefined;
    if (RANGE_MARKS.has(kind)) mark = readRangeMark(ctx, el, kind as RangeMark['type'], undefined);
    if (mark) block = { type: 'mark', mark };
    else {
      const wrapped = readWrapper(ctx, el, kind, wrap, 'body');
      if (wrapped) {
        if (span) {
          wrapped.wrapper.srcStart = span.srcStart;
          wrapped.wrapper.srcEnd = span.srcEnd;
        }
        const before = out.length;
        readBlocks(ctx, wrapped.content, out, wrapped.wrapper, nested(depth));
        if (out.length === before) out.push({ type: 'empty', wrap: wrapped.wrapper });
        return;
      }
      block = { type: 'unknown', xml: keep(ctx, el, 'body') };
    }
    if (wrap) block.wrap = wrap;
  }
  if (span) {
    block.srcStart = span.srcStart;
    block.srcEnd = span.srcEnd;
  }
  out.push(block);
}

interface OpenField {
  field: ComplexField;
  inResult: boolean;
  /** No "end" in the story: closed at the end of the paragraph it begins in. */
  unclosed: boolean;
}

interface LinkState {
  /** Open fields, innermost last. */
  stack: OpenField[];
  /** Open fields beyond MAX_FIELD_DEPTH, which are not assembled. */
  skipped: number;
  /** The "begin" items without an "end", once that is known (second pass). */
  unclosed: Set<FieldCharItem> | undefined;
  /** Unclosed fields on the stack. */
  pending: number;
}

function linkRun(run: Run, state: LinkState): void {
  const stack = state.stack;
  const content = run.content;
  // One flag per item, as an array only when the items of the run differ.
  let first = false;
  let flags: boolean[] | undefined;
  for (let i = 0; i < content.length; i++) {
    const item = content[i];
    // An item is in a field code when a field around it has not reached its result yet.
    let code = false;
    for (let k = 0; k < stack.length; k++) {
      if (!stack[k].inResult) {
        code = true;
        break;
      }
    }
    if (i === 0) first = code;
    else if (flags) flags.push(code);
    else if (code !== first) {
      flags = [];
      for (let k = 0; k < i; k++) flags.push(first);
      flags.push(code);
    }

    if (item.type === 'fieldChar') {
      if (item.charType === 'begin') {
        const unclosed = state.unclosed !== undefined && state.unclosed.has(item);
        if (stack.length >= MAX_FIELD_DEPTH) {
          // Not assembled. One that never ends must not be counted: there is no "end" to count it down.
          if (!unclosed) state.skipped++;
          if (item.field) delete item.field;
        } else {
          item.field = unclosed ? { instruction: '', result: '', unclosed: true } : { instruction: '', result: '' };
          stack.push({ field: item.field, inResult: false, unclosed });
          if (unclosed) state.pending++;
        }
      } else if (state.skipped) {
        if (item.charType === 'end') state.skipped--;
      } else if (stack.length) {
        if (item.charType === 'separate') stack[stack.length - 1].inResult = true;
        else if (stack.pop()!.unclosed) state.pending--;
      }
    } else if (state.skipped || !stack.length) continue;
    else if (item.type === 'instrText') {
      const top = stack[stack.length - 1];
      if (!top.inResult) top.field.instruction += item.text;
    } else if (!code && (item.type === 'text' || item.type === 'tab')) {
      // Shown text belongs to the result of every field around it.
      for (const open of stack) open.field.result += item.type === 'tab' ? '\t' : item.text;
    }
  }
  if (flags) run.fieldCode = flags;
  else if (first && content.length) run.fieldCode = true;
  else if (run.fieldCode !== undefined) delete run.fieldCode;
}

function linkBlocks(blocks: Block[], state: LinkState): void {
  for (const block of blocks) {
    if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) linkBlocks(cell.blocks, state);
      continue;
    }
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) if (inline.type === 'run') linkRun(inline, state);
    if (state.pending) {
      // The fields that never end, all begun in this paragraph, end here.
      const stack = state.stack;
      let kept = 0;
      for (const open of stack) if (!open.unclosed) stack[kept++] = open;
      stack.length = kept;
      state.pending = 0;
    }
  }
}

/** The "begin" items of the blocks that no "end" answers, by plain nesting. */
function unclosedFields(blocks: Block[], open: FieldCharItem[]): void {
  for (const block of blocks) {
    if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) unclosedFields(cell.blocks, open);
      continue;
    }
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) {
      if (inline.type !== 'run') continue;
      for (const item of inline.content) {
        if (item.type !== 'fieldChar') continue;
        if (item.charType === 'begin') open.push(item);
        else if (item.charType === 'end') open.pop();
      }
    }
  }
}

/**
 * Assembles complex fields: the "begin" item of each gets the code and the
 * cached result, and the items of runs that sit in a field code are flagged
 * (Run.fieldCode). A field may start in one paragraph and end in a later one,
 * so this runs over a whole story.
 *
 * A field without an "end" would hide everything after it. Such a field is
 * taken to end with the paragraph it begins in: what follows in that
 * paragraph is its code (or its result, after a separator), later paragraphs
 * are not touched by it. Whether a field ends is only known at the end of the
 * story, so a story that has such fields is gone through a second time.
 */
export function linkFields(blocks: Block[]): void {
  const state: LinkState = { stack: [], skipped: 0, unclosed: undefined, pending: 0 };
  linkBlocks(blocks, state);
  if (!state.stack.length && !state.skipped) return;
  const open: FieldCharItem[] = [];
  unclosedFields(blocks, open);
  linkBlocks(blocks, { stack: [], skipped: 0, unclosed: new Set(open), pending: 0 });
}

export interface ScannedBlocks {
  /** <w:sectPr> found among the blocks (only the body has one). */
  section?: SectionProps;
  /** Where it is in the source, and how many blocks came before it. */
  sectionSpan?: SourceSpan;
  sectionIndex?: number;
}

/**
 * Reads the block children of the element the scanner is on (an open tag)
 * up to its closing tag. The cursor is on the same tag and follows: it gives
 * each block its place in the source, and the comments and processing
 * instructions between the blocks, which the scanner passes over.
 */
export function scanBlocks(ctx: PartContext, sc: XmlScanner, cursor: SourceCursor, out: Block[]): ScannedBlocks {
  const result: ScannedBlocks = {};
  cursor.enter();
  if (sc.selfClosing) return result;
  const misc = (node: XmlComment | XmlPI, srcStart: number, srcEnd: number) => {
    out.push({ type: 'misc', node, srcStart, srcEnd });
  };
  while (sc.next()) {
    if (sc.kind === 'text') continue;
    cursor.toTag(misc);
    // Children are consumed whole below, so a closing tag seen here is the container's.
    if (sc.kind === 'close') {
      cursor.leave();
      break;
    }
    const raw = sc.readRaw();
    const srcStart = cursor.pos;
    const span = cursor.skip(raw) ? { srcStart, srcEnd: cursor.pos } : undefined;
    // Everything kept from the block (drawings, math, unknown elements) is bounded in depth by this check.
    const el = bounded(parseXml(raw).root);
    if (kindOf(ctx, el.name) === 'sectPr' && !result.section) {
      result.section = readSectionProps(ctx, el);
      if (span) result.sectionSpan = span;
      result.sectionIndex = out.length;
    } else readBlock(ctx, el, out, undefined, 0, span);
  }
  return result;
}
