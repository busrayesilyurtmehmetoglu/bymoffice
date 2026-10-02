// Drawings: what an inline or anchored drawing shows (a picture, the content
// of its text boxes, or a placeholder) and where an anchored one goes.

import type { Block, DrawingAnchor, DrawingItem, LegacyGraphicItem } from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { stackBands, type FieldValues, type FloatSpec, type LayoutContext, type StoryRef } from './context';
import type { ImageItem, PageItem } from './types';
import { MAX_LAYOUT_DEPTH, MAX_PAGE_SIDE, clamp, emu, finite } from './units';

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

function textBoxItems(
  ctx: LayoutContext,
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
  } else if (!item.textBoxes) {
    const placeholder: ImageItem = { type: 'image', x: 0, y: 0, width, height, placeholder: item.kind };
    if (item.description !== undefined) placeholder.alt = item.description;
    items.push(placeholder);
  }
  if (item.textBoxes) {
    const used = textBoxItems(ctx, item.textBoxes, width, bodyInsets(item.xml), story, path, depth, fields, items);
    // A box without a height of its own is as high as its text.
    if (height === 0) height = used;
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
    const used = textBoxItems(ctx, item.textBoxes, width, DEFAULT_INSETS, story, path, depth, fields, items);
    if (height === 0) height = used;
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

/** The area text keeps out of for a float at (x, y); undefined when text runs over or under it. */
export function floatExclusion(float: FloatSpec, x: number, y: number): (Rect & { whole: boolean }) | undefined {
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
  };
}
