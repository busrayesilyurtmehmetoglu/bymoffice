// Reads a drawing part (xl/drawings/drawingN.xml) into the read-only view the
// grid draws from: pictures, simple shapes with their text, groups, and
// placeholders for what is not drawn (charts, diagrams, embedded objects).
//
// The part itself stays in the package and is written back from there; nothing
// here is ever serialized. One anchor that cannot be read is left out and
// noted, the rest of the drawing is still shown.

import type { OpcPackage, Relationship } from '../../opc';
import {
  MAX_COLS,
  MAX_ROWS,
  type DrawingAnchor,
  type DrawingMarker,
  type DrawingObject,
  type DrawingObjectBase,
  type DrawingParagraph,
  type DrawingShape,
  type DrawingText,
  type DrawingTextRun,
  type SheetDrawing,
} from '../../model/sheet/types';
import { localName, parseXml, type XmlElement } from '../../xml';

/** Anchors read from one drawing; a sheet with more shows the first ones. */
export const MAX_SHEET_DRAWINGS = 20000;
/** Objects (group members included) read from one drawing. */
export const MAX_DRAWING_OBJECTS = 50000;
/** Groups inside groups; deeper members are left out. */
export const MAX_GROUP_DEPTH = 16;
/** Wrappers (mc:AlternateContent) looked through around an anchor or an object. */
const MAX_WRAPPER_DEPTH = 8;
const MAX_SHAPE_PARAGRAPHS = 2000;
const MAX_SHAPE_TEXT = 32767;
/** EMU per pixel at 96 dpi. */
export const EMU_PER_PIXEL = 9525;
/** A sheet is at most this wide or tall, in EMU (16,384 columns of 255 characters and then some). */
const MAX_EMU = 2e12;

export interface DrawingReadContext {
  /** Relationships of the drawing part. */
  rels: Relationship[];
  /** Theme colours as "RRGGBB" in the order the style sheet keeps them (lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink). */
  themeColors: string[];
  /** Says whether a part exists; a picture whose part is missing is shown as a placeholder. */
  hasPart?(part: string): boolean;
  warnings: string[];
}

interface ReadState {
  context: DrawingReadContext;
  relById: Map<string, Relationship>;
  objects: number;
}

function child(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const node of parent.children) if (node.type === 'element' && localName(node.name) === local) return node;
  return undefined;
}

function textOf(el: XmlElement | undefined): string {
  if (!el) return '';
  let text = '';
  for (const node of el.children) if (node.type === 'text') text += node.value;
  return text;
}

/** Attribute by local name with a namespace prefix (r:embed). */
function prefixedAttr(el: XmlElement, local: string): string | undefined {
  for (const [name, value] of el.attrs) if (name.includes(':') && localName(name) === local) return value;
  return undefined;
}

function int(text: string | undefined): number | undefined {
  if (text === undefined) return undefined;
  const trimmed = text.trim();
  if (!/^-?\d{1,15}$/.test(trimmed)) return undefined;
  return Number(trimmed);
}

const isTrue = (value: string | undefined): boolean => value === '1' || value === 'true';

const clamp = (value: number, low: number, high: number): number => (value < low ? low : value > high ? high : value);

// ------------------------------------------------------------------ colours

/** DrawingML names of theme colours -> index in StyleSheet.themeColors. */
const SCHEME_INDEX: { [name: string]: number } = {
  bg1: 0,
  lt1: 0,
  tx1: 1,
  dk1: 1,
  bg2: 2,
  lt2: 2,
  tx2: 3,
  dk2: 3,
  accent1: 4,
  accent2: 5,
  accent3: 6,
  accent4: 7,
  accent5: 8,
  accent6: 9,
  hlink: 10,
  folHlink: 11,
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
 * and close for colours (the same reading as in documents).
 */
export function drawingColor(themeColors: string[], el: XmlElement | undefined): string | undefined {
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
    const index = value !== undefined && Object.prototype.hasOwnProperty.call(SCHEME_INDEX, value) ? SCHEME_INDEX[value] : undefined;
    const theme = index === undefined ? undefined : themeColors[index];
    // Without a theme part the two colours every file relies on are still known.
    rgb = typeof theme === 'string' && HEX6.test(theme) ? theme : index === 0 ? 'FFFFFF' : index === 1 ? '000000' : undefined;
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

function firstColor(themeColors: string[], holder: XmlElement | undefined): string | undefined {
  if (!holder) return undefined;
  for (const node of holder.children) if (node.type === 'element') return drawingColor(themeColors, node);
  return undefined;
}

/** A fill stated among the children: a colour, null for a fill that shows nothing here, undefined when none is stated. */
function statedFill(themeColors: string[], holder: XmlElement | undefined): string | null | undefined {
  if (!holder) return undefined;
  for (const node of holder.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (local === 'solidFill') return firstColor(themeColors, node) || null;
    if (local === 'noFill' || local === 'pattFill' || local === 'blipFill' || local === 'grpFill') return null;
    if (local === 'gradFill') {
      // A gradient is shown as its first colour: closer to what the file has than nothing.
      const stops = child(node, 'gsLst');
      const stop = child(stops, 'gs');
      return firstColor(themeColors, stop) || null;
    }
  }
  return undefined;
}

// --------------------------------------------------------------------- text

const ALIGN: { [value: string]: DrawingParagraph['align'] } = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' };
const ANCHOR: { [value: string]: DrawingText['anchor'] } = { t: 'top', ctr: 'middle', b: 'bottom', just: 'top', dist: 'top' };

/** Default space around the text of a shape, EMU: 0.1 inch at the sides, 0.05 inch above and below. */
const DEFAULT_INSETS = [91440, 45720, 91440, 45720];
const MAX_INSET = 9525 * 2000;

function runProps(themeColors: string[], props: XmlElement | undefined, run: DrawingTextRun): void {
  if (!props) return;
  const size = int(props.attrs.get('sz'));
  if (size !== undefined && size >= 100 && size <= 400000) run.size = size / 100;
  const bold = props.attrs.get('b');
  if (bold !== undefined) run.bold = isTrue(bold);
  const italic = props.attrs.get('i');
  if (italic !== undefined) run.italic = isTrue(italic);
  const underline = props.attrs.get('u');
  if (underline !== undefined) run.underline = underline !== 'none';
  const strike = props.attrs.get('strike');
  if (strike !== undefined) run.strike = strike !== 'noStrike';
  const fill = statedFill(themeColors, props);
  if (fill) run.color = fill;
  const latin = child(props, 'latin');
  const face = latin ? latin.attrs.get('typeface') : undefined;
  // "+mn-lt" and the like name a font of the theme, not a font.
  if (face !== undefined && face !== '' && face.charAt(0) !== '+') run.font = face;
}

function readText(state: ReadState, body: XmlElement | undefined, defaults: DrawingTextRun): DrawingText | undefined {
  if (!body) return undefined;
  const themeColors = state.context.themeColors;
  const paragraphs: DrawingParagraph[] = [];
  let length = 0;
  let any = false;
  // Level 1 of the body's own list style states what its paragraphs start from.
  const listStyle = child(body, 'lstStyle');
  const level = child(listStyle, 'lvl1pPr');
  const base: DrawingTextRun = { ...defaults };
  runProps(themeColors, child(level, 'defRPr'), base);
  const baseAlign = level ? level.attrs.get('algn') : undefined;
  for (const node of body.children) {
    if (node.type !== 'element' || localName(node.name) !== 'p') continue;
    if (paragraphs.length >= MAX_SHAPE_PARAGRAPHS || length >= MAX_SHAPE_TEXT) break;
    const paragraph: DrawingParagraph = { runs: [] };
    const props = child(node, 'pPr');
    const align = (props && props.attrs.get('algn')) || baseAlign;
    if (align !== undefined && Object.prototype.hasOwnProperty.call(ALIGN, align) && ALIGN[align] !== 'left') paragraph.align = ALIGN[align];
    const start: DrawingTextRun = { ...base };
    runProps(themeColors, child(props, 'defRPr'), start);
    for (const item of node.children) {
      if (item.type !== 'element') continue;
      const kind = localName(item.name);
      if (kind !== 'r' && kind !== 'br' && kind !== 'fld') continue;
      const run: DrawingTextRun = { ...start, text: '' };
      runProps(themeColors, child(item, 'rPr'), run);
      let text = kind === 'br' ? '\n' : textOf(child(item, 't'));
      if (length + text.length > MAX_SHAPE_TEXT) text = text.slice(0, MAX_SHAPE_TEXT - length);
      if (text === '') continue;
      length += text.length;
      run.text = text;
      if (kind !== 'br') any = true;
      paragraph.runs.push(run);
    }
    paragraphs.push(paragraph);
  }
  if (!any) return undefined;
  const out: DrawingText = { paragraphs };
  const bodyProps = child(body, 'bodyPr');
  if (bodyProps) {
    const anchor = bodyProps.attrs.get('anchor');
    if (anchor !== undefined && Object.prototype.hasOwnProperty.call(ANCHOR, anchor) && ANCHOR[anchor] !== 'top') out.anchor = ANCHOR[anchor];
    if (bodyProps.attrs.get('wrap') === 'none') out.wrap = false;
  }
  const names = ['lIns', 'tIns', 'rIns', 'bIns'];
  const insets: [number, number, number, number] = [0, 0, 0, 0];
  for (let i = 0; i < 4; i++) {
    const stated = bodyProps ? int(bodyProps.attrs.get(names[i])) : undefined;
    insets[i] = clamp(stated === undefined ? DEFAULT_INSETS[i] : stated, 0, MAX_INSET) / EMU_PER_PIXEL;
  }
  out.insets = insets;
  return out;
}

// ------------------------------------------------------------------ objects

interface Transform {
  x: number;
  y: number;
  width: number;
  height: number;
  /** Child coordinate space of a group. */
  childX?: number;
  childY?: number;
  childWidth?: number;
  childHeight?: number;
}

function readTransform(props: XmlElement | undefined, base: DrawingObjectBase): Transform | undefined {
  const xfrm = child(props, 'xfrm');
  if (!xfrm) return undefined;
  const rot = int(xfrm.attrs.get('rot'));
  if (rot !== undefined && rot !== 0) {
    // 60,000ths of a degree.
    const degrees = ((rot / 60000) % 360 + 360) % 360;
    if (degrees !== 0) base.rotation = degrees;
  }
  if (isTrue(xfrm.attrs.get('flipH'))) base.flipH = true;
  if (isTrue(xfrm.attrs.get('flipV'))) base.flipV = true;
  const off = child(xfrm, 'off');
  const ext = child(xfrm, 'ext');
  const x = off ? int(off.attrs.get('x')) : undefined;
  const y = off ? int(off.attrs.get('y')) : undefined;
  const width = ext ? int(ext.attrs.get('cx')) : undefined;
  const height = ext ? int(ext.attrs.get('cy')) : undefined;
  if (x === undefined || y === undefined || width === undefined || height === undefined) return undefined;
  const out: Transform = { x, y, width, height };
  const chOff = child(xfrm, 'chOff');
  const chExt = child(xfrm, 'chExt');
  if (chOff && chExt) {
    out.childX = int(chOff.attrs.get('x'));
    out.childY = int(chOff.attrs.get('y'));
    out.childWidth = int(chExt.attrs.get('cx'));
    out.childHeight = int(chExt.attrs.get('cy'));
  }
  return out;
}

function readNames(holder: XmlElement | undefined, base: DrawingObjectBase): void {
  const props = child(holder, 'cNvPr');
  if (!props) return;
  const name = props.attrs.get('name');
  if (name !== undefined && name !== '') base.name = name.slice(0, 255);
  const description = props.attrs.get('descr');
  if (description !== undefined && description !== '') base.description = description.slice(0, 2000);
  if (isTrue(props.attrs.get('hidden'))) base.hidden = true;
}

/** Outline widths of the three lines of the usual theme, by the index a shape style refers to them with; px. */
const STYLE_LINE_WIDTHS = [0, 1, 1.5, 2];
const MAX_OUTLINE = 32;

function shapeStyle(state: ReadState, shape: XmlElement, props: XmlElement | undefined, out: DrawingShape): void {
  const themeColors = state.context.themeColors;
  const style = child(shape, 'style');
  const styleRef = (name: string): { color: string | undefined; index: number } | undefined => {
    const ref = child(style, name);
    const idx = ref ? int(ref.attrs.get('idx')) : undefined;
    if (!ref || idx === undefined || idx === 0) return undefined;
    return { color: firstColor(themeColors, ref), index: idx };
  };
  const fill = statedFill(themeColors, props);
  const fillRef = fill === undefined ? styleRef('fillRef') : undefined;
  const fillColor = fill === undefined ? (fillRef ? fillRef.color : undefined) : fill;
  if (fillColor) out.fill = fillColor;
  const outline = child(props, 'ln');
  const stroke = statedFill(themeColors, outline);
  const lineRef = styleRef('lnRef');
  const strokeColor = stroke === undefined ? (lineRef ? lineRef.color : undefined) : stroke;
  if (strokeColor) {
    const stated = outline ? int(outline.attrs.get('w')) : undefined;
    const width = stated !== undefined && stated >= 0 ? stated / EMU_PER_PIXEL : lineRef ? STYLE_LINE_WIDTHS[Math.min(lineRef.index, 3)] || 1 : 1;
    if (width > 0) {
      out.line = { color: strokeColor, width: Math.min(Math.max(width, 0.75), MAX_OUTLINE) };
      const dash = child(outline, 'prstDash');
      const dashKind = dash ? dash.attrs.get('val') : undefined;
      if ((dashKind !== undefined && dashKind !== 'solid') || child(outline, 'custDash')) out.line.dash = true;
    }
  }
}

const OBJECT_NAMES: { [local: string]: true } = { pic: true, sp: true, grpSp: true, cxnSp: true, graphicFrame: true, contentPart: true };

/**
 * The object an element stands for, looking through mc:AlternateContent: the
 * first choice that holds something known is taken, otherwise the fallback.
 * Never both, so nothing is shown twice.
 */
function objectElements(parent: XmlElement, depth: number, out: XmlElement[]): void {
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(OBJECT_NAMES, local)) out.push(node);
    else if (local === 'AlternateContent' && depth < MAX_WRAPPER_DEPTH) {
      const before = out.length;
      for (const branch of node.children) {
        if (branch.type !== 'element' || localName(branch.name) !== 'Choice') continue;
        objectElements(branch, depth + 1, out);
        if (out.length > before) break;
      }
      if (out.length === before) {
        const fallback = child(node, 'Fallback');
        if (fallback) objectElements(fallback, depth + 1, out);
      }
    }
  }
}

function readObject(state: ReadState, el: XmlElement, depth: number, transformOut?: { value?: Transform }): DrawingObject | undefined {
  if (state.objects >= MAX_DRAWING_OBJECTS) return undefined;
  state.objects++;
  const themeColors = state.context.themeColors;
  const base: DrawingObjectBase = {};
  const kind = localName(el.name);
  if (kind === 'pic') {
    readNames(child(el, 'nvPicPr'), base);
    const transform = readTransform(child(el, 'spPr'), base);
    if (transformOut) transformOut.value = transform;
    const picture: DrawingObject = { ...base, type: 'picture' };
    const blipFill = child(el, 'blipFill');
    const blip = child(blipFill, 'blip');
    const embed = blip ? prefixedAttr(blip, 'embed') : undefined;
    const link = blip ? prefixedAttr(blip, 'link') : undefined;
    const rel = embed !== undefined && embed !== '' ? state.relById.get(embed) : undefined;
    if (rel && !rel.external && rel.part && (!state.context.hasPart || state.context.hasPart(rel.part))) picture.media = rel.part;
    // A picture kept outside the file is never fetched; neither is an embedded one whose relationship points outside.
    else if ((link !== undefined && link !== '') || (rel && rel.external)) picture.external = true;
    const crop = child(blipFill, 'srcRect');
    if (crop) {
      const side = (name: string): number => {
        const value = int(crop.attrs.get(name));
        return value === undefined ? 0 : clamp(value / 100000, 0, 1);
      };
      const cut = { left: side('l'), top: side('t'), right: side('r'), bottom: side('b') };
      // What is left has to have a size.
      if ((cut.left > 0 || cut.top > 0 || cut.right > 0 || cut.bottom > 0) && cut.left + cut.right < 1 && cut.top + cut.bottom < 1) picture.crop = cut;
    }
    return picture;
  }
  if (kind === 'sp' || kind === 'cxnSp') {
    readNames(child(el, kind === 'sp' ? 'nvSpPr' : 'nvCxnSpPr'), base);
    const props = child(el, 'spPr');
    const transform = readTransform(props, base);
    if (transformOut) transformOut.value = transform;
    const geometry = child(props, 'prstGeom');
    const preset = geometry ? geometry.attrs.get('prst') : undefined;
    const shape: DrawingShape = { ...base, type: 'shape', geometry: preset !== undefined && /^[A-Za-z0-9]{1,64}$/.test(preset) ? preset : 'rect' };
    if (kind === 'cxnSp' || shape.geometry === 'line' || /Connector\d*$/.test(shape.geometry)) shape.connector = true;
    shapeStyle(state, el, props, shape);
    if (kind === 'sp') {
      // The colour a shape style gives its text: usually white on a filled shape.
      const defaults: DrawingTextRun = { text: '' };
      const fontRef = child(child(el, 'style'), 'fontRef');
      const fontColor = firstColor(themeColors, fontRef);
      if (fontColor) defaults.color = fontColor;
      const text = readText(state, child(el, 'txBody'), defaults);
      if (text) shape.text = text;
    }
    return shape;
  }
  if (kind === 'grpSp') {
    readNames(child(el, 'nvGrpSpPr'), base);
    const transform = readTransform(child(el, 'grpSpPr'), base);
    if (transformOut) transformOut.value = transform;
    const group: DrawingObject = { ...base, type: 'group', children: [] };
    if (depth >= MAX_GROUP_DEPTH) return group;
    const members: XmlElement[] = [];
    objectElements(el, 0, members);
    const space =
      transform && transform.childWidth !== undefined && transform.childHeight !== undefined && transform.childX !== undefined && transform.childY !== undefined
        ? { x: transform.childX, y: transform.childY, width: transform.childWidth, height: transform.childHeight }
        : transform
          ? { x: transform.x, y: transform.y, width: transform.width, height: transform.height }
          : undefined;
    for (const member of members) {
      const holder: { value?: Transform } = {};
      const object = readObject(state, member, depth + 1, holder);
      if (!object) break;
      const own = holder.value;
      // A member is placed by its own transform in the coordinates of the group; without both it has no place.
      if (!own || !space || !(space.width > 0) || !(space.height > 0)) continue;
      const box = {
        x: (own.x - space.x) / space.width,
        y: (own.y - space.y) / space.height,
        width: own.width / space.width,
        height: own.height / space.height,
      };
      if (!(Number.isFinite(box.x) && Number.isFinite(box.y) && box.width >= 0 && box.height >= 0 && Math.abs(box.x) < 1e6 && Math.abs(box.y) < 1e6 && box.width < 1e6 && box.height < 1e6)) continue;
      object.box = box;
      group.children.push(object);
    }
    return group;
  }
  // A graphic frame holds a chart, a diagram or something else that is not drawn here.
  if (kind === 'graphicFrame') readNames(child(el, 'nvGraphicFramePr'), base);
  const transform = readTransform(el, base);
  if (transformOut) transformOut.value = transform;
  const data = child(child(el, 'graphic'), 'graphicData');
  const uri = (data && data.attrs.get('uri')) || '';
  const what = /\/chart(ex)?$|\/chart\b/.test(uri) ? 'chart' : /diagram/.test(uri) ? 'diagram' : 'object';
  return { ...base, type: 'placeholder', what };
}

// ------------------------------------------------------------------ anchors

function readMarker(el: XmlElement | undefined): DrawingMarker | undefined {
  if (!el) return undefined;
  const col = int(textOf(child(el, 'col')));
  const row = int(textOf(child(el, 'row')));
  if (col === undefined || row === undefined || col < 0 || row < 0 || col >= MAX_COLS || row >= MAX_ROWS) return undefined;
  const colOff = int(textOf(child(el, 'colOff')));
  const rowOff = int(textOf(child(el, 'rowOff')));
  return { col, colOff: clamp(colOff === undefined ? 0 : colOff, 0, MAX_EMU), row, rowOff: clamp(rowOff === undefined ? 0 : rowOff, 0, MAX_EMU) };
}

function readAnchor(el: XmlElement): DrawingAnchor | undefined {
  const kind = localName(el.name);
  if (kind === 'twoCellAnchor') {
    const from = readMarker(child(el, 'from'));
    const to = readMarker(child(el, 'to'));
    if (!from || !to) return undefined;
    const editAs = el.attrs.get('editAs');
    const anchor: DrawingAnchor = { kind: 'twoCell', from, to };
    if (editAs === 'oneCell' || editAs === 'absolute') anchor.editAs = editAs;
    return anchor;
  }
  const ext = child(el, 'ext');
  const width = ext ? int(ext.attrs.get('cx')) : undefined;
  const height = ext ? int(ext.attrs.get('cy')) : undefined;
  if (width === undefined || height === undefined || width < 0 || height < 0) return undefined;
  if (kind === 'oneCellAnchor') {
    const from = readMarker(child(el, 'from'));
    if (!from) return undefined;
    return { kind: 'oneCell', from, width: Math.min(width, MAX_EMU), height: Math.min(height, MAX_EMU) };
  }
  const pos = child(el, 'pos');
  const x = pos ? int(pos.attrs.get('x')) : undefined;
  const y = pos ? int(pos.attrs.get('y')) : undefined;
  if (x === undefined || y === undefined) return undefined;
  return { kind: 'absolute', x: clamp(x, 0, MAX_EMU), y: clamp(y, 0, MAX_EMU), width: Math.min(width, MAX_EMU), height: Math.min(height, MAX_EMU) };
}

const ANCHOR_NAMES: { [local: string]: true } = { twoCellAnchor: true, oneCellAnchor: true, absoluteAnchor: true };

function anchorElements(parent: XmlElement, depth: number, out: XmlElement[]): void {
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(ANCHOR_NAMES, local)) out.push(node);
    else if (local === 'AlternateContent' && depth < MAX_WRAPPER_DEPTH) {
      // The same object in a newer and an older form, each with its own anchor: one of them is shown.
      const before = out.length;
      for (const branch of node.children) {
        if (branch.type !== 'element' || localName(branch.name) !== 'Choice') continue;
        anchorElements(branch, depth + 1, out);
        if (out.length > before) break;
      }
      if (out.length === before) {
        const fallback = child(node, 'Fallback');
        if (fallback) anchorElements(fallback, depth + 1, out);
      }
    }
  }
}

/**
 * The anchors of a drawing part with what they hold, bottom one first. Never
 * throws for what the part states: an anchor that cannot be read is left out
 * with a note in `context.warnings`.
 */
export function parseDrawing(root: XmlElement, context: DrawingReadContext): SheetDrawing[] {
  const state: ReadState = { context, relById: new Map(context.rels.map((rel) => [rel.id, rel])), objects: 0 };
  const anchors: XmlElement[] = [];
  anchorElements(root, 0, anchors);
  const out: SheetDrawing[] = [];
  let skipped = 0;
  for (let i = 0; i < anchors.length; i++) {
    if (out.length >= MAX_SHEET_DRAWINGS || state.objects >= MAX_DRAWING_OBJECTS) {
      context.warnings.push(`Drawing: only the first ${out.length} of ${anchors.length} objects are shown`);
      break;
    }
    let drawing: SheetDrawing | undefined;
    try {
      const anchor = readAnchor(anchors[i]);
      const content: XmlElement[] = [];
      objectElements(anchors[i], 0, content);
      const object = anchor && content.length > 0 ? readObject(state, content[0], 0) : undefined;
      if (anchor && object) drawing = { anchor, object };
    } catch (error) {
      // One object that cannot be read does not take the others with it.
      drawing = undefined;
    }
    if (drawing) out.push(drawing);
    else skipped++;
  }
  if (skipped > 0) context.warnings.push(`Drawing: ${skipped} object(s) could not be read and are not shown`);
  return out;
}

/** Parts larger than this are not read; the sheet is shown without its drawing. */
const MAX_PART_BYTES = 64 * 1024 * 1024;

/**
 * The drawings of a sheet from the parts its relationships name; undefined
 * for a sheet without any. Parts of pictures that are shown are added to
 * `media`. Never throws for what the parts hold.
 */
export async function readSheetDrawings(
  pkg: OpcPackage,
  sheetRels: Relationship[],
  themeColors: string[],
  media: Set<string>,
  warnings: string[],
): Promise<SheetDrawing[] | undefined> {
  let drawings: SheetDrawing[] = [];
  for (const rel of sheetRels) {
    if (rel.type.slice(rel.type.lastIndexOf('/') + 1) !== 'drawing' || rel.external || !rel.part || !pkg.hasPart(rel.part)) continue;
    try {
      const bytes = await pkg.readPart(rel.part);
      if (bytes.length > MAX_PART_BYTES) throw new Error('the part is too large');
      const rels = await pkg.relationships(rel.part);
      const read = parseDrawing(parseXml(bytes).root, { rels, themeColors, hasPart: (part) => pkg.hasPart(part), warnings });
      drawings = drawings.concat(read.slice(0, MAX_SHEET_DRAWINGS - drawings.length));
    } catch (error) {
      warnings.push(`Drawing could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  if (drawings.length === 0) return undefined;
  collectMedia(drawings.map((drawing) => drawing.object), media, 0);
  return drawings;
}

function collectMedia(objects: DrawingObject[], media: Set<string>, depth: number): void {
  for (const object of objects) {
    if (object.type === 'picture' && object.media !== undefined) media.add(object.media);
    else if (object.type === 'group' && depth <= MAX_GROUP_DEPTH) collectMedia(object.children, media, depth + 1);
  }
}
