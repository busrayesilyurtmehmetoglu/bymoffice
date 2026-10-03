// Drawings: what an inline or anchored drawing shows (a picture, the content
// of its text boxes, or a placeholder) and where an anchored one goes.

import type { Block, DrawingAnchor, DrawingItem, DrawingTransform, LegacyGraphicItem, ShapeDash, ShapeLook } from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { stackBands, type FieldValues, type FloatSpec, type LayoutContext, type StoryRef } from './context';
import type { ImageItem, PageItem, ShapeItem } from './types';
import { MAX_LAYOUT_DEPTH, MAX_PAGE_SIDE, MAX_TEXT_BOX_WORK, clamp, emu, finite } from './units';

export interface BoxContent {
  width: number;
  height: number;
  items: PageItem[];
}

/** Insets Word gives a text box that states none: 0.1 inch left and right, 0.05 inch top and bottom; px. */
const INSET_X = 9.6;
const INSET_Y = 4.8;

const size = (value: number) => clamp(finite(value), 0, MAX_PAGE_SIDE);

interface Insets {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const DEFAULT_INSETS: Insets = { left: INSET_X, top: INSET_Y, right: INSET_X, bottom: INSET_Y };

/** The first element with the local name under `el`; the reader bounds how deep the trees it keeps are. */
function findElement(el: XmlElement, local: string, depth: number): XmlElement | undefined {
  if (depth > 64) return undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const colon = node.name.indexOf(':');
    if ((colon < 0 ? node.name : node.name.slice(colon + 1)) === local) return node;
    const inner = findElement(node, local, depth + 1);
    if (inner) return inner;
  }
  return undefined;
}

/** Insets of the text box of a DrawingML shape (<wps:bodyPr lIns= tIns= rIns= bIns=>, EMU). */
function bodyInsets(xml: XmlElement): Insets {
  const body = findElement(xml, 'bodyPr', 0);
  if (!body) return DEFAULT_INSETS;
  const read = (name: string, fallback: number) => {
    const text = body.attrs.get(name);
    if (text === undefined || !/^\d{1,9}$/.test(text)) return fallback;
    return size(emu(Number(text)));
  };
  return { left: read('lIns', INSET_X), top: read('tIns', INSET_Y), right: read('rIns', INSET_X), bottom: read('bIns', INSET_Y) };
}

const localName = (name: string): string => {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
};

/** The first child element with the local name. */
function childElement(el: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!el) return undefined;
  for (const node of el.children) if (node.type === 'element' && localName(node.name) === local) return node;
  return undefined;
}

/** What a shape shows around its content, whichever format it was read from; lengths in px. */
interface ShapeView {
  geometry: 'rect' | 'roundRect' | 'ellipse' | 'line';
  /** For "roundRect": the radius of the corners as a share of the shorter side. */
  corner?: number;
  fill?: string;
  line?: { color: string; width: number; dash?: ShapeDash };
}

/** DrawingML names of theme colours -> the names resolveColor knows. */
const SCHEME_COLORS: { [name: string]: string } = {
  tx1: 'text1',
  tx2: 'text2',
  bg1: 'background1',
  bg2: 'background2',
  dk1: 'dark1',
  dk2: 'dark2',
  lt1: 'light1',
  lt2: 'light2',
  hlink: 'hyperlink',
  folHlink: 'followedHyperlink',
  accent1: 'accent1',
  accent2: 'accent2',
  accent3: 'accent3',
  accent4: 'accent4',
  accent5: 'accent5',
  accent6: 'accent6',
};

/** The colour names shapes are most often given, as "RRGGBB". */
const NAMED_COLORS: { [name: string]: string } = {
  black: '000000',
  white: 'FFFFFF',
  red: 'FF0000',
  green: '008000',
  lime: '00FF00',
  blue: '0000FF',
  yellow: 'FFFF00',
  cyan: '00FFFF',
  aqua: '00FFFF',
  magenta: 'FF00FF',
  fuchsia: 'FF00FF',
  gray: '808080',
  grey: '808080',
  silver: 'C0C0C0',
  maroon: '800000',
  navy: '000080',
  olive: '808000',
  purple: '800080',
  teal: '008080',
  orange: 'FFA500',
};

const HEX6 = /^[0-9A-Fa-f]{6}$/;

/** A share stated in thousandths of a percent ("85000" is 0.85); undefined when it is not one. */
function share(text: string | undefined): number | undefined {
  if (text === undefined || !/^\d{1,7}$/.test(text)) return undefined;
  return Math.min(Number(text) / 100000, 1);
}

/**
 * The colour of a DrawingML colour element (<a:srgbClr>, <a:schemeClr>,
 * <a:sysClr>, <a:prstClr>) with the commonest changes to it applied, as
 * "#RRGGBB"; undefined when it states none or is fully transparent.
 * Luminance changes are applied to the channels, which is exact for greys
 * and close for colours.
 */
function drawingColor(ctx: LayoutContext, el: XmlElement | undefined): string | undefined {
  if (!el) return undefined;
  const kind = localName(el.name);
  const value = el.attrs.get('val');
  let rgb: string | undefined;
  if (kind === 'srgbClr') rgb = value !== undefined && HEX6.test(value) ? value : undefined;
  else if (kind === 'sysClr') {
    const last = el.attrs.get('lastClr');
    rgb = last !== undefined && HEX6.test(last) ? last : value === 'window' ? 'FFFFFF' : value === 'windowText' ? '000000' : undefined;
  } else if (kind === 'prstClr') rgb = value !== undefined && Object.prototype.hasOwnProperty.call(NAMED_COLORS, value) ? NAMED_COLORS[value] : undefined;
  else if (kind === 'schemeClr') {
    const name = value !== undefined && Object.prototype.hasOwnProperty.call(SCHEME_COLORS, value) ? SCHEME_COLORS[value] : undefined;
    const resolved = name === undefined ? '' : ctx.color({ theme: name }, '');
    rgb = resolved.charAt(0) === '#' && HEX6.test(resolved.slice(1)) ? resolved.slice(1) : undefined;
  }
  if (rgb === undefined) return undefined;
  const channels = [parseInt(rgb.slice(0, 2), 16), parseInt(rgb.slice(2, 4), 16), parseInt(rgb.slice(4, 6), 16)];
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const amount = share(node.attrs.get('val'));
    if (amount === undefined) continue;
    const change = localName(node.name);
    if (change === 'alpha' && amount === 0) return undefined;
    for (let i = 0; i < 3; i++) {
      if (change === 'lumMod' || change === 'shade') channels[i] *= amount;
      else if (change === 'lumOff') channels[i] += 255 * amount;
      else if (change === 'tint') channels[i] += (255 - channels[i]) * (1 - amount);
    }
  }
  let out = '#';
  for (let i = 0; i < 3; i++) {
    const channel = Math.round(clamp(channels[i], 0, 255));
    out += (channel < 16 ? '0' : '') + channel.toString(16).toUpperCase();
  }
  return out;
}

/** The geometry a preset shape of DrawingML is drawn with; the ones that are not here are not drawn. */
const PRESETS: { [prst: string]: ShapeView['geometry'] } = {
  rect: 'rect',
  round1Rect: 'rect',
  round2SameRect: 'rect',
  round2DiagRect: 'rect',
  snip1Rect: 'rect',
  snip2SameRect: 'rect',
  snip2DiagRect: 'rect',
  snipRoundRect: 'rect',
  flowChartProcess: 'rect',
  flowChartAlternateProcess: 'rect',
  roundRect: 'roundRect',
  ellipse: 'ellipse',
  line: 'line',
  straightConnector1: 'line',
};

const PRESET_DASHES: { [name: string]: true } = {
  dash: true,
  dot: true,
  dashDot: true,
  lgDash: true,
  lgDashDot: true,
  lgDashDotDot: true,
  sysDash: true,
  sysDot: true,
  sysDashDot: true,
  sysDashDotDot: true,
};

/** 0.75 pt, the outline of a shape that states a colour but no width; px. */
const DEFAULT_OUTLINE = 1;
const MAX_OUTLINE = 16;
/** An outline of the width 0 is the thinnest line the device draws; the painter makes it one pixel of the canvas. px. */
const HAIRLINE = 0.5;
/** The corners of a rounded rectangle that states none: a sixth of its shorter side. */
const DEFAULT_CORNER = 16667;

const outlineWidth = (px: number): number => (px > 0 ? Math.min(px, MAX_OUTLINE) : HAIRLINE);

const cornerShare = (value: number | undefined): number => clamp(finite(value === undefined ? DEFAULT_CORNER : value), 0, 50000) / 100000;

/** A shape of DrawingML as it is drawn. */
interface DrawingMlShape {
  view: ShapeView;
  /** It states a fill that is not shown: a gradient, a pattern, a picture. */
  hiddenFill: boolean;
  /** Its geometry is stated, as a preset. */
  preset: boolean;
}

/**
 * Geometry, fill and outline of the shape of a DrawingML drawing: what
 * <wps:spPr> states (<a:prstGeom>, <a:solidFill>, <a:ln>), and for what it
 * leaves open the colours of the shape style. The caller hands in a drawing
 * that is one shape; gradients, patterns and pictures as fills are not shown.
 * Undefined when there is no shape, or its geometry is one that is not drawn.
 */
function drawingMlShape(ctx: LayoutContext, xml: XmlElement): DrawingMlShape | undefined {
  const shape = findElement(xml, 'wsp', 0);
  if (!shape) return undefined;
  const props = childElement(shape, 'spPr');
  const geometry = childElement(props, 'prstGeom');
  const preset = geometry ? geometry.attrs.get('prst') : undefined;
  if (preset !== undefined && !Object.prototype.hasOwnProperty.call(PRESETS, preset)) return undefined;
  const view: ShapeView = { geometry: preset === undefined ? 'rect' : PRESETS[preset] };
  if (view.geometry === 'roundRect') {
    // <a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst>
    const guide = childElement(childElement(geometry, 'avLst'), 'gd');
    const match = guide && guide.attrs.get('name') === 'adj' ? /^val (\d{1,6})$/.exec(guide.attrs.get('fmla') || '') : null;
    view.corner = cornerShare(match ? Number(match[1]) : undefined);
  }
  const style = childElement(shape, 'style');
  const styleColor = (name: string): string | undefined => {
    const ref = childElement(style, name);
    const idx = ref ? ref.attrs.get('idx') : undefined;
    if (!ref || idx === undefined || idx === '0') return undefined;
    for (const node of ref.children) if (node.type === 'element') return drawingColor(ctx, node);
    return undefined;
  };
  const colorOf = (holder: XmlElement | undefined): string | undefined => {
    if (!holder) return undefined;
    for (const node of holder.children) if (node.type === 'element') return drawingColor(ctx, node);
    return undefined;
  };
  let hiddenFill = false;
  /** A fill stated among the children: a colour, null for a fill that shows nothing here, undefined when none is stated. */
  const stated = (holder: XmlElement | undefined, own: boolean): string | null | undefined => {
    if (!holder) return undefined;
    for (const node of holder.children) {
      if (node.type !== 'element') continue;
      const local = localName(node.name);
      if (local === 'solidFill') return colorOf(node) || null;
      if (local === 'noFill') return null;
      if (local === 'gradFill' || local === 'pattFill' || local === 'blipFill' || local === 'grpFill') {
        if (own) hiddenFill = true;
        return null;
      }
    }
    return undefined;
  };
  const fill = stated(props, true);
  const fillColor = fill === undefined ? styleColor('fillRef') : fill;
  if (fillColor && view.geometry !== 'line') view.fill = fillColor;
  const outline = childElement(props, 'ln');
  const stroke = stated(outline, false);
  const strokeColor = stroke === undefined ? styleColor('lnRef') : stroke;
  if (strokeColor) {
    const text = outline ? outline.attrs.get('w') : undefined;
    const width = text !== undefined && /^\d{1,9}$/.test(text) ? emu(Number(text)) : DEFAULT_OUTLINE;
    view.line = { color: strokeColor, width: outlineWidth(width) };
    const dash = childElement(outline, 'prstDash');
    const name = dash ? dash.attrs.get('val') : undefined;
    if (name !== undefined && Object.prototype.hasOwnProperty.call(PRESET_DASHES, name)) view.line.dash = name as ShapeDash;
  }
  return { view, hiddenFill, preset: preset !== undefined };
}

/** The look of a shape of a doc or RTF file as it is drawn; a geometry the reader does not know is drawn as a box. */
function lookView(look: ShapeLook): ShapeView {
  const view: ShapeView = { geometry: look.geometry === 'other' ? 'rect' : look.geometry };
  if (view.geometry === 'roundRect') view.corner = cornerShare(look.corner);
  if (typeof look.fill === 'string' && HEX6.test(look.fill) && view.geometry !== 'line') view.fill = '#' + look.fill.toUpperCase();
  const line = look.line;
  if (line && typeof line.color === 'string' && HEX6.test(line.color)) {
    view.line = { color: '#' + line.color.toUpperCase(), width: outlineWidth(size(emu(line.width))) };
    if (line.dash !== undefined && Object.prototype.hasOwnProperty.call(PRESET_DASHES, line.dash)) view.line.dash = line.dash;
  }
  return view;
}

/** A colour as VML writes it: "#rrggbb", "#rgb" or a name, at times followed by a palette index ("#4f81bd [3204]"). */
function vmlColor(text: string | undefined): string | undefined {
  if (text === undefined) return undefined;
  const match = /^\s*(#[0-9A-Fa-f]{6}|#[0-9A-Fa-f]{3}|[A-Za-z]{3,20})\b/.exec(text);
  if (!match) return undefined;
  const value = match[1];
  if (value.charAt(0) !== '#') {
    const name = value.toLowerCase();
    return Object.prototype.hasOwnProperty.call(NAMED_COLORS, name) ? '#' + NAMED_COLORS[name] : undefined;
  }
  if (value.length === 7) return value.toUpperCase();
  return ('#' + value.charAt(1) + value.charAt(1) + value.charAt(2) + value.charAt(2) + value.charAt(3) + value.charAt(3)).toUpperCase();
}

const vmlOff = (text: string | undefined): boolean => text === 'f' || text === 'false' || text === '0';

/**
 * Fill and stroke of a VML shape with a text box. A VML shape is stroked in
 * black, 0.75 pt wide, unless it says otherwise; its default white fill is
 * not painted (white paper shows the same, and nothing behind it is covered).
 */
function vmlStyle(shape: XmlElement | undefined): ShapeView {
  const out: ShapeView = { geometry: 'rect' };
  if (!shape) return out;
  const fillElement = childElement(shape, 'fill');
  const strokeElement = childElement(shape, 'stroke');
  if (!vmlOff(shape.attrs.get('filled')) && !(fillElement && vmlOff(fillElement.attrs.get('on')))) {
    const fill = vmlColor(shape.attrs.get('fillcolor')) || (fillElement ? vmlColor(fillElement.attrs.get('color')) : undefined);
    if (fill !== undefined) out.fill = fill;
  }
  if (!vmlOff(shape.attrs.get('stroked')) && !(strokeElement && vmlOff(strokeElement.attrs.get('on')))) {
    const stated = cssLength(shape.attrs.get('strokeweight'));
    const width = stated === undefined ? DEFAULT_OUTLINE : stated;
    if (width > 0) out.line = { color: vmlColor(shape.attrs.get('strokecolor')) || '#000000', width: Math.min(width, MAX_OUTLINE) };
  }
  return out;
}

/**
 * The fill of a shape under its content and its outline over it. A box with
 * a solid outline is a filled rectangle and four lines; everything else is a
 * shape the painter draws. A shape with content is not turned: its text
 * would not turn with it.
 */
function framed(view: ShapeView | undefined, transform: DrawingTransform | undefined, width: number, height: number, content: PageItem[], out: PageItem[]): void {
  const push = () => {
    for (let i = 0; i < content.length; i++) out.push(content[i]);
  };
  if (!view) return push();
  const line = view.geometry === 'line';
  const visible = line ? width > 0 || height > 0 : width > 0 && height > 0;
  const rotation = transform && transform.rotation && !content.length ? finite(transform.rotation / 60000) % 360 : 0;
  if (!visible) return push();
  if (view.geometry === 'rect' && !(view.line && view.line.dash) && rotation === 0) {
    if (view.fill !== undefined) out.push({ type: 'rect', x: 0, y: 0, width, height, fill: view.fill });
    push();
    if (view.line) {
      const { color, width: thickness } = view.line;
      const side = (x1: number, y1: number, x2: number, y2: number) => out.push({ type: 'line', x1, y1, x2, y2, color, width: thickness, style: 'single' });
      side(0, 0, width, 0);
      side(0, height, width, height);
      side(0, 0, 0, height);
      side(width, 0, width, height);
    }
    return;
  }
  const shape = (): ShapeItem => {
    const item: ShapeItem = { type: 'shape', x: 0, y: 0, width, height, geometry: view.geometry };
    if (view.geometry === 'roundRect') item.radius = Math.min(width, height) * (view.corner === undefined ? DEFAULT_CORNER / 100000 : view.corner);
    if (rotation !== 0) item.rotation = rotation;
    if (line && transform) {
      // Only a line looks different when it is mirrored.
      if (transform.flipH) item.flipH = true;
      if (transform.flipV) item.flipV = true;
    }
    return item;
  };
  if (!content.length) {
    if (view.fill === undefined && !view.line) return;
    const item = shape();
    if (view.fill !== undefined) item.fill = view.fill;
    if (view.line) item.stroke = { ...view.line };
    out.push(item);
    return;
  }
  if (view.fill !== undefined) {
    const item = shape();
    item.fill = view.fill;
    out.push(item);
  }
  push();
  if (view.line) {
    const item = shape();
    item.stroke = { ...view.line };
    out.push(item);
  }
}

/** What a paragraph or a table row counts for beside its characters. */
const BLOCK_WEIGHT = 16;

/** What the blocks cost to lay out: their characters, and something for each paragraph and table row. */
function blocksWeight(blocks: Block[], depth: number): number {
  let weight = 0;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    weight += BLOCK_WEIGHT;
    if (block.type === 'paragraph') {
      for (const inline of block.children) {
        if (inline.type !== 'run') continue;
        for (const item of inline.content) if (item.type === 'text' || item.type === 'deletedText') weight += item.text.length;
      }
    } else if (block.type === 'table' && depth < MAX_LAYOUT_DEPTH) {
      for (const row of block.rows) {
        weight += BLOCK_WEIGHT;
        for (const cell of row.cells) weight += blocksWeight(cell.blocks, depth + 1);
      }
    }
  }
  return weight;
}

/**
 * Whether the text of a box is laid out for the drawing `owner`. The text of a box belongs
 * to one drawing: a second drawing that shows the same blocks (a model may share them) gets
 * an empty box, so the text is not laid out, and kept, once for each of them. And all boxes
 * of a document together are laid out only up to a limit; the ones after it stay empty.
 */
function takesBox(ctx: LayoutContext, blocks: Block[], owner: object): boolean {
  return ctx.boxEvent(blocks, owner, () => decideBox(ctx, blocks, owner));
}

function decideBox(ctx: LayoutContext, blocks: Block[], owner: object): boolean {
  const first = ctx.boxOwners.get(blocks);
  if (first === undefined) ctx.boxOwners.set(blocks, owner);
  else if (first !== owner) return false;
  const stated = ctx.options.maxTextBoxWork;
  const limit = stated !== undefined && stated >= 0 ? stated : MAX_TEXT_BOX_WORK;
  if (ctx.boxWork >= limit) return false;
  let weight = ctx.boxWeights.get(blocks);
  if (weight === undefined) ctx.boxWeights.set(blocks, (weight = blocksWeight(blocks, 0)));
  ctx.boxWork += weight;
  return true;
}

function textBoxItems(
  ctx: LayoutContext,
  owner: object,
  boxes: Block[][],
  width: number,
  insets: Insets,
  story: StoryRef,
  path: number[],
  depth: number,
  fields: FieldValues | undefined,
  out: PageItem[],
): number {
  if (depth >= MAX_LAYOUT_DEPTH) return 0;
  let y = insets.top;
  for (let k = 0; k < boxes.length; k++) {
    if (!takesBox(ctx, boxes[k], owner)) continue;
    const bands = ctx.layoutBlocks(ctx, boxes[k], Math.max(1, width - insets.left - insets.right), story, path.concat([k]), undefined, depth + 1, fields);
    y += stackBands(bands, out, insets.left, y);
  }
  return y + insets.bottom;
}

/** What a DrawingML drawing shows. `path` is the path of the text boxes in it (see TextItem.source). */
export function drawingContent(
  ctx: LayoutContext,
  item: DrawingItem,
  story: StoryRef,
  path: number[],
  depth: number,
  fields: FieldValues | undefined,
): BoxContent {
  const width = size(emu(item.width));
  let height = size(emu(item.height));
  const items: PageItem[] = [];
  if (item.image) {
    const image: ImageItem = { type: 'image', x: 0, y: 0, width, height };
    if (!item.image.external && item.image.target !== undefined) image.part = item.image.target;
    if (item.crop) {
      const part = (value: number | undefined) => (value === undefined ? 0 : clamp(value / 100000, -10, 1));
      image.crop = { left: part(item.crop.left), top: part(item.crop.top), right: part(item.crop.right), bottom: part(item.crop.bottom) };
    }
    const transform = item.transform;
    if (transform) {
      if (transform.rotation) image.rotation = transform.rotation / 60000;
      if (transform.flipH) image.flipH = true;
      if (transform.flipV) image.flipV = true;
    }
    if (item.effectExtent) {
      const e = item.effectExtent;
      image.effect = { left: size(emu(e.left)), top: size(emu(e.top)), right: size(emu(e.right)), bottom: size(emu(e.bottom)) };
    }
    if (item.description !== undefined) image.alt = item.description;
    items.push(image);
  }
  // What the shape looks like: a doc or RTF file states it in the model, a docx in its tree.
  let view: ShapeView | undefined;
  let drawn = false;
  if (!item.image) {
    if (item.shape) {
      view = lookView(item.shape);
      // As for a docx: a shape whose fill cannot be shown is a placeholder.
      drawn = item.shape.geometry !== 'other' && !(item.shape.lost && item.shape.lost.indexOf('fill') >= 0);
    } else if (item.kind === 'shape') {
      // Only a drawing that is one shape: the first shape of a group or a canvas says nothing about the whole.
      const found = drawingMlShape(ctx, item.xml);
      if (found) {
        view = found.view;
        drawn = found.preset && !found.hiddenFill;
      }
    }
  }
  if (item.textBoxes) {
    const content: PageItem[] = [];
    const used = textBoxItems(ctx, item, item.textBoxes, width, bodyInsets(item.xml), story, path, depth, fields, content);
    // A box without a height of its own is as high as its text.
    if (height === 0) height = used;
    framed(view, item.transform, width, height, content, items);
  } else if (!item.image) {
    // A shape without text is drawn as it is when all of it can be; a placeholder stands for the others.
    if (view && drawn) framed(view, item.transform, width, height, [], items);
    else {
      const placeholder: ImageItem = { type: 'image', x: 0, y: 0, width, height, placeholder: item.kind };
      if (item.description !== undefined) placeholder.alt = item.description;
      items.push(placeholder);
    }
  }
  return { width, height, items };
}

const UNITS: Record<string, number> = { pt: 4 / 3, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, px: 1, pc: 16, emu: 1 / 9525 };

function cssLength(text: string | undefined): number | undefined {
  if (text === undefined) return undefined;
  const match = /^\s*(-?\d{1,7}(?:\.\d{1,6})?)\s*(pt|in|cm|mm|px|pc|emu)?\s*$/.exec(text);
  if (!match) return undefined;
  return Number(match[1]) * (match[2] === undefined ? 1 / 9525 : UNITS[match[2]]);
}

const VML_SHAPES: Record<string, true> = { shape: true, rect: true, roundrect: true, oval: true, line: true, group: true, image: true };

/** The first VML shape under the element: where a <w:pict> says how large it is. */
function vmlShape(el: XmlElement, depth: number): XmlElement | undefined {
  // The reader bounds the depth of every tree it keeps; the check here is for callers that build trees themselves.
  if (depth > 64) return undefined;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const colon = node.name.indexOf(':');
    const local = colon < 0 ? node.name : node.name.slice(colon + 1);
    if (VML_SHAPES[local] && node.attrs.has('style')) return node;
    const inner = vmlShape(node, depth + 1);
    if (inner) return inner;
  }
  return undefined;
}

function styleMap(style: string): Map<string, string> {
  const out = new Map<string, string>();
  for (const part of style.split(';')) {
    const colon = part.indexOf(':');
    if (colon > 0) out.set(part.slice(0, colon).trim().toLowerCase(), part.slice(colon + 1).trim());
  }
  return out;
}

export interface LegacyContent extends BoxContent {
  /** Set when the shape is positioned absolutely: it does not sit in the line. */
  anchor?: DrawingAnchor;
}

const VML_RELATIVE: Record<string, string> = { page: 'page', margin: 'margin', text: 'column', char: 'column', line: 'paragraph' };

/** What a <w:pict> or <w:object> shows: its picture, its text boxes, or a placeholder of its size. */
export function legacyContent(
  ctx: LayoutContext,
  item: LegacyGraphicItem,
  story: StoryRef,
  path: number[],
  depth: number,
  fields: FieldValues | undefined,
): LegacyContent {
  const shape = vmlShape(item.xml, 0);
  const style = shape ? styleMap(shape.attrs.get('style') || '') : new Map<string, string>();
  const width = size(cssLength(style.get('width')) || 0);
  let height = size(cssLength(style.get('height')) || 0);
  const items: PageItem[] = [];
  if (item.image) {
    const image: ImageItem = { type: 'image', x: 0, y: 0, width, height };
    if (!item.image.external && item.image.target !== undefined) image.part = item.image.target;
    items.push(image);
  } else if (!item.textBoxes && width > 0 && height > 0) {
    items.push({ type: 'image', x: 0, y: 0, width, height, placeholder: item.type === 'object' ? 'object' : 'vml' });
  }
  if (item.textBoxes) {
    const content: PageItem[] = [];
    const used = textBoxItems(ctx, item, item.textBoxes, width, DEFAULT_INSETS, story, path, depth, fields, content);
    if (height === 0) height = used;
    framed(item.image ? undefined : vmlStyle(shape), undefined, width, height, content, items);
  }
  const out: LegacyContent = { width, height, items };
  if (style.get('position') === 'absolute') {
    const horizontal = style.get('mso-position-horizontal');
    const vertical = style.get('mso-position-vertical');
    const anchor: DrawingAnchor = {
      horizontal: { relativeFrom: VML_RELATIVE[style.get('mso-position-horizontal-relative') || 'text'] || 'column' },
      vertical: { relativeFrom: VML_RELATIVE[style.get('mso-position-vertical-relative') || 'line'] || 'paragraph' },
      wrap: 'none',
    };
    if (horizontal === 'center' || horizontal === 'left' || horizontal === 'right') anchor.horizontal!.align = horizontal;
    else anchor.horizontal!.offset = Math.round((cssLength(style.get('margin-left')) || 0) * 9525);
    if (vertical === 'center' || vertical === 'top' || vertical === 'bottom') anchor.vertical!.align = vertical;
    else anchor.vertical!.offset = Math.round((cssLength(style.get('margin-top')) || 0) * 9525);
    const z = Number(style.get('z-index'));
    if (z < 0) anchor.behindDoc = true;
    out.anchor = anchor;
  }
  return out;
}

export function makeFloat(content: BoxContent, anchor: DrawingAnchor, columnWidth: number, inCell: boolean): FloatSpec {
  return { width: content.width, height: content.height, items: content.items, anchor, dx: 0, dy: 0, columnWidth, inCell };
}

/** Geometry a float is placed in; all in page coordinates, px. */
export interface FloatFrame {
  pageWidth: number;
  pageHeight: number;
  /** Text area of the page. */
  left: number;
  right: number;
  top: number;
  bottom: number;
  /** Text area of the anchoring paragraph. */
  columnLeft: number;
  columnWidth: number;
  /** Top of the anchoring paragraph. */
  paragraphTop: number;
}

export interface Rect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/** Where an anchored drawing goes: its top left corner. */
export function floatPosition(float: FloatSpec, frame: FloatFrame): { x: number; y: number } {
  const anchor = float.anchor;
  if (anchor.simplePos) return { x: emu(anchor.simplePos.x), y: emu(anchor.simplePos.y) };

  let base = frame.columnLeft;
  let extent = frame.columnWidth;
  const h = anchor.horizontal;
  switch (h && h.relativeFrom) {
    case 'page':
      base = 0;
      extent = frame.pageWidth;
      break;
    case 'margin':
      base = frame.left;
      extent = frame.right - frame.left;
      break;
    case 'leftMargin':
    case 'insideMargin':
      base = 0;
      extent = frame.left;
      break;
    case 'rightMargin':
    case 'outsideMargin':
      base = frame.right;
      extent = frame.pageWidth - frame.right;
      break;
  }
  let x = base;
  if (h) {
    if (h.align === 'center') x = base + (extent - float.width) / 2;
    else if (h.align === 'right' || h.align === 'outside') x = base + extent - float.width;
    else if (h.align !== undefined) x = base;
    else if (h.pctOffset !== undefined) x = base + (extent * h.pctOffset) / 100000;
    else x = base + emu(h.offset);
  }

  let top = frame.paragraphTop;
  let span = 0;
  const v = anchor.vertical;
  switch (v && v.relativeFrom) {
    case 'page':
      top = 0;
      span = frame.pageHeight;
      break;
    case 'margin':
      top = frame.top;
      span = frame.bottom - frame.top;
      break;
    case 'topMargin':
      top = 0;
      span = frame.top;
      break;
    case 'bottomMargin':
      top = frame.bottom;
      span = frame.pageHeight - frame.bottom;
      break;
  }
  let y = top;
  if (v) {
    if (v.align === 'center') y = top + (span - float.height) / 2;
    else if (v.align === 'bottom' || v.align === 'outside') y = top + span - float.height;
    else if (v.align !== undefined) y = top;
    else if (v.pctOffset !== undefined) y = top + (span * v.pctOffset) / 100000;
    else y = top + emu(v.offset);
  }
  return { x: finite(x), y: finite(y) };
}

/** Where text may be beside a drawing it wraps around: on both sides, on one, or on the wider one. */
export type WrapSide = 'both' | 'left' | 'right' | 'largest';

/** The area text keeps out of for a float at (x, y); undefined when text runs over or under it. */
export function floatExclusion(float: FloatSpec, x: number, y: number): (Rect & { whole: boolean; side: WrapSide }) | undefined {
  const anchor = float.anchor;
  const wrap = anchor.wrap;
  if (float.inCell || wrap === undefined || wrap === 'none' || float.width <= 0 || float.height <= 0) return undefined;
  const d = anchor.wrapDist || {};
  const pick = (own: number | undefined, general: number | undefined) => size(emu(own !== undefined ? own : general));
  return {
    left: x - pick(d.left, anchor.distL),
    right: x + float.width + pick(d.right, anchor.distR),
    top: y - pick(d.top, anchor.distT),
    bottom: y + float.height + pick(d.bottom, anchor.distB),
    // Text stays clear of the whole width, not just of the drawing.
    whole: wrap === 'topAndBottom',
    side: anchor.wrapText === 'left' || anchor.wrapText === 'right' || anchor.wrapText === 'largest' ? anchor.wrapText : 'both',
  };
}
