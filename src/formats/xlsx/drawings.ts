// Reads a drawing part (xl/drawings/drawingN.xml) into the read-only view the
// grid draws from: pictures, simple shapes with their text, groups, charts
// (their parts are read by chart.ts), and placeholders for what is not drawn
// (charts of other kinds, diagrams, embedded objects).
//
// The part itself stays in the package and is written back from there; nothing
// here is ever serialized. One anchor that cannot be read is left out and
// noted, the rest of the drawing is still shown.

import type { OpcPackage, Relationship } from '../../opc';
import { MAX_CHART_PART_BYTES, MAX_SHEET_CHARTS, MAX_WORKBOOK_CHART_BYTES, noteChartPart, type Chart } from '../../model/sheet/chart';
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
  type ShapeTextLook,
} from '../../model/sheet/types';
import { localName, parseXml, xmlEncodingRead, type XmlElement } from '../../xml';
import { PIVOT_SCAN_LIMITS } from '../../model/sheet/pivot';
import { MAX_DRAWING_TEXT, textUpTo } from '../../model/sheet/pictures';
import { MAX_PATCH_DEPTH, chartKnown, chartLoose, nestedBeyond, type ChartPartFacts } from './chart-write';
import { LINE_ENDS, sameShapeElement, shapeElementXml } from './shape-write';

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
/** Chart parts read for one drawing; further charts are shown as placeholders. */
export const MAX_DRAWING_CHARTS = MAX_SHEET_CHARTS;
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
  /** The charts of the drawing that could be read, by the name of their part; a chart that is not among them is a placeholder. */
  charts?: Map<string, Chart>;
  /** What is known of the parts of those charts beside their models (DrawingChart.known), by the same names. */
  chartsKnown?: Map<string, ChartPartFacts>;
  warnings: string[];
}

/** Reads the root element of a chart part; undefined for a chart that is not drawn. See parseChart(). */
export type ChartPartReader = (root: XmlElement, themeColors: string[], budget?: { points: number }) => Chart | undefined;

/**
 * How the chart parts of a workbook are read, and what reading them may still cost. One
 * object is handed to readSheetDrawings() for every sheet of a workbook, so that the limits
 * hold for the workbook and not for each sheet again.
 */
export interface ChartReading {
  read: ChartPartReader;
  /** Reads the root element of a part of the newer chart types (chartex: parseChartEx()); without it such a chart is a placeholder. */
  readEx?: ChartPartReader;
  /** Chart parts that may still be read; counted down. */
  charts: number;
  /** Cached values (<c:pt>) that may still be read; counted down by `read`. */
  points: number;
  /** Chart parts larger than this are not read. Default 32 MB. */
  maxPartBytes?: number;
  /** Bytes of chart parts that may still be read; counted down. MAX_WORKBOOK_CHART_BYTES when missing. */
  bytes?: number;
  /** Bytes of drawing parts that may still be read; counted down. MAX_WORKBOOK_DRAWING_BYTES when missing. */
  drawingBytes?: number;
  /**
   * What was made of each chart part so far, by the name the package has it under (null: not
   * shown). Filled by the reader: a part that several drawings name, or one drawing in several
   * spellings, is read and counted once.
   */
  known?: Map<string, Chart | null>;
  /** Why a chart part of `known` could not be read: said again for every drawing that names it. */
  failed?: Map<string, string>;
  /** What is known of each chart part of `known` beside its model (DrawingChart.known), by the same names. */
  written?: Map<string, ChartPartFacts>;
  /** The names of the parts of the package; made by the reader on first use. */
  names?: PartNames;
  /** The package `known`, `failed` and `names` belong to: they are made anew for another one. */
  of?: OpcPackage;
}

/** What `reading` knows of the parts of `pkg`; what it knew of another package is let go of. */
function partsKnown(pkg: OpcPackage, reading: ChartReading): { names: PartNames; known: Map<string, Chart | null>; failed: Map<string, string>; written: Map<string, ChartPartFacts> } {
  if (reading.of !== pkg || !reading.names || !reading.known || !reading.failed || !reading.written) {
    reading.of = pkg;
    reading.names = partNames(pkg);
    reading.known = new Map();
    reading.failed = new Map();
    reading.written = new Map();
  }
  return { names: reading.names, known: reading.known, failed: reading.failed, written: reading.written };
}

/** Gives the name the package has a part under; undefined for a part it does not have. */
export type PartNames = (part: string) => string | undefined;

/**
 * The names of the parts of a package as it spells them. Part names are compared without
 * regard to case (OPC), so one part can be named in many spellings; whoever keeps what it
 * read by part name has to keep it under one of them. A name that is spelt as an entry of
 * the package is that entry; otherwise it is the first entry that differs in case only, as
 * OpcPackage.readPart() reads it.
 */
export function partNames(pkg: Pick<OpcPackage, 'partNames'>): PartNames {
  let exact: Set<string> | undefined;
  let folded: Map<string, string> | undefined;
  return (part) => {
    if (!exact || !folded) {
      exact = new Set();
      folded = new Map();
      for (const name of pkg.partNames()) {
        exact.add(name);
        const lower = name.toLowerCase();
        if (!folded.has(lower)) folded.set(lower, name);
      }
    }
    return exact.has(part) ? part : folded.get(part.toLowerCase());
  };
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

/**
 * How text typed into a shape looks (DrawingShape.textLook): the mark that ends its first
 * paragraph, or its first run when that has none, with what the body and the shape's style give
 * every run; the alignment of the first paragraph and where the body puts its text. Undefined for
 * a body without a paragraph.
 */
function readTextLook(state: ReadState, body: XmlElement, defaults: DrawingTextRun): ShapeTextLook | undefined {
  const themeColors = state.context.themeColors;
  const first = child(body, 'p');
  if (!first) return undefined;
  const level = child(child(body, 'lstStyle'), 'lvl1pPr');
  const run: DrawingTextRun = { ...defaults };
  runProps(themeColors, child(level, 'defRPr'), run);
  const props = child(first, 'pPr');
  runProps(themeColors, child(props, 'defRPr'), run);
  let mark = child(first, 'endParaRPr');
  if (!mark) {
    for (const item of first.children) {
      if (item.type !== 'element' || (localName(item.name) !== 'r' && localName(item.name) !== 'br')) continue;
      mark = child(item, 'rPr');
      break;
    }
  }
  runProps(themeColors, mark, run);
  const { text: _text, ...ownLook } = run;
  void _text;
  const look: ShapeTextLook = { run: ownLook };
  const align = (props && props.attrs.get('algn')) || (level ? level.attrs.get('algn') : undefined);
  if (align !== undefined && Object.prototype.hasOwnProperty.call(ALIGN, align) && ALIGN[align] !== 'left') look.align = ALIGN[align];
  const bodyProps = child(body, 'bodyPr');
  const anchor = bodyProps ? bodyProps.attrs.get('anchor') : undefined;
  if (anchor !== undefined && Object.prototype.hasOwnProperty.call(ANCHOR, anchor) && ANCHOR[anchor] !== 'top') look.anchor = ANCHOR[anchor];
  if (bodyProps && bodyProps.attrs.get('wrap') === 'none') look.wrap = false;
  const names = ['lIns', 'tIns', 'rIns', 'bIns'];
  const insets: [number, number, number, number] = [0, 0, 0, 0];
  for (let i = 0; i < 4; i++) {
    const stated = bodyProps ? int(bodyProps.attrs.get(names[i])) : undefined;
    insets[i] = clamp(stated === undefined ? DEFAULT_INSETS[i] : stated, 0, MAX_INSET) / EMU_PER_PIXEL;
  }
  look.insets = insets;
  return look;
}

/** The text of an element as XML says it, attributes in order: what tells two looks apart. */
function elementKey(el: XmlElement | undefined, depth = 0): string {
  if (!el) return '';
  let out = `<${localName(el.name)}`;
  for (const [name, value] of el.attrs) out += ` ${name}="${value}"`;
  out += '>';
  if (depth < 16) for (const node of el.children) out += node.type === 'element' ? elementKey(node, depth + 1) : node.type === 'text' ? node.value.trim() : '?';
  return out;
}

/**
 * Whether the text body of a shape is plain (DrawingShape.plainText): its children are the body
 * properties, a list style and paragraphs; each paragraph holds its properties, runs and the mark
 * that ends it, and nothing else; every run is a look and a text; every paragraph has the same
 * properties and the same mark, and every run the same look; and the reader keeps all of it (no
 * more than MAX_SHAPE_PARAGRAPHS paragraphs and MAX_SHAPE_TEXT characters). A new text is written
 * in that look, paragraph by paragraph: a line break, another mark or text the model does not
 * hold would be lost to it.
 */
function plainBody(body: XmlElement): boolean {
  let paragraphs = 0;
  let length = 0;
  let props: string | undefined;
  let mark: string | undefined;
  let look: string | undefined;
  const same = (held: string | undefined, key: string): boolean => held === undefined || held === key;
  for (const node of body.children) {
    if (node.type === 'text') {
      if (node.value.trim() !== '') return false;
      continue;
    }
    if (node.type !== 'element') return false;
    const local = localName(node.name);
    if (local === 'bodyPr' || local === 'lstStyle') continue;
    if (local !== 'p') return false;
    // The reader stops at a paragraph that comes once the text is as long as it reads: the body is not all read.
    if (length >= MAX_SHAPE_TEXT) return false;
    if (++paragraphs > MAX_SHAPE_PARAGRAPHS) return false;
    const key = elementKey(child(node, 'pPr'));
    if (!same(props, key)) return false;
    props = key;
    const end = elementKey(child(node, 'endParaRPr'));
    if (!same(mark, end)) return false;
    mark = end;
    for (const item of node.children) {
      if (item.type === 'text') {
        if (item.value.trim() !== '') return false;
        continue;
      }
      if (item.type !== 'element') return false;
      const kind = localName(item.name);
      if (kind === 'pPr' || kind === 'endParaRPr') continue;
      if (kind !== 'r') return false;
      for (const part of item.children) {
        if (part.type === 'text' && part.value.trim() === '') continue;
        if (part.type !== 'element') return false;
        const name = localName(part.name);
        if (name !== 'rPr' && name !== 't') return false;
        if (name === 't' && (length += textOf(part).length) > MAX_SHAPE_TEXT) return false;
      }
      const runLook = elementKey(child(item, 'rPr'));
      if (!same(look, runLook)) return false;
      look = runLook;
    }
  }
  return paragraphs > 0;
}

const SHAPE_SPACES = 'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

/** The children of a shape, and of its properties, that the writer writes (see shape-write.ts). */
const WRITTEN_SHAPE_PARTS: { [local: string]: true } = { nvSpPr: true, nvCxnSpPr: true, spPr: true, txBody: true };
const WRITTEN_SHAPE_PROPS: { [local: string]: true } = { xfrm: true, prstGeom: true, solidFill: true, noFill: true, ln: true };

/** Whether the element of a shape is what the writer makes of its model (DrawingShape.known). */
function writtenAsModel(el: XmlElement, shape: DrawingShape): boolean {
  // What the writer never writes says no at once, before anything is made to compare: a style, effects, a geometry of its own.
  for (const node of el.children) if (node.type === 'element' && !Object.prototype.hasOwnProperty.call(WRITTEN_SHAPE_PARTS, localName(node.name))) return false;
  const spPr = child(el, 'spPr');
  for (const node of spPr ? spPr.children : []) if (node.type === 'element' && !Object.prototype.hasOwnProperty.call(WRITTEN_SHAPE_PROPS, localName(node.name))) return false;
  const props = child(child(el, localName(el.name) === 'sp' ? 'nvSpPr' : 'nvCxnSpPr'), 'cNvPr');
  const id = props ? int(props.attrs.get('id')) : undefined;
  if (id === undefined || id < 1 || id > 4294967294) return false;
  try {
    const made = parseXml(`<w ${SHAPE_SPACES}>${shapeElementXml(shape, id, { x: 0, y: 0, width: 0, height: 0 })}</w>`).root;
    const own = made.children.find((node): node is XmlElement => node.type === 'element');
    return !!own && sameShapeElement(el, own);
  } catch (error) {
    return false;
  }
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
  if (description !== undefined && description !== '') base.description = textUpTo(description, MAX_DRAWING_TEXT);
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
      // The ends of a line: an arrow head, say. One of a kind that is not known is drawn as a triangle.
      for (const [name, key] of [['headEnd', 'head'], ['tailEnd', 'tail']] as const) {
        const end = child(outline, name);
        const type = end ? end.attrs.get('type') : undefined;
        if (type !== undefined && type !== 'none') out.line[key] = LINE_ENDS.indexOf(type) >= 0 ? type : 'triangle';
      }
    }
  }
}

const OBJECT_NAMES: { [local: string]: true } = { pic: true, sp: true, grpSp: true, cxnSp: true, graphicFrame: true, contentPart: true };

/**
 * The frame of a chart of the newer types (chartex) whose part was not read into a chart: what
 * stands beside it as the fallback is shown in its place.
 */
function unreadChartEx(state: ReadState, el: XmlElement): boolean {
  if (localName(el.name) !== 'graphicFrame') return false;
  const data = child(child(el, 'graphic'), 'graphicData');
  if (!/\/chartex$/.test((data && data.attrs.get('uri')) || '')) return false;
  const reference = child(data, 'chart');
  const id = reference ? prefixedAttr(reference, 'id') : undefined;
  const rel = id !== undefined ? state.relById.get(id) : undefined;
  return !(rel && !rel.external && rel.part && state.context.charts && state.context.charts.has(rel.part));
}

/**
 * The object an element stands for, looking through mc:AlternateContent: the
 * first choice that holds something known is taken, otherwise the fallback.
 * Never both, so nothing is shown twice. With `state`, a choice that holds a
 * chart of the newer types that was not read is not taken.
 */
function objectElements(parent: XmlElement, depth: number, out: XmlElement[], state?: ReadState): void {
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(OBJECT_NAMES, local)) out.push(node);
    else if (local === 'AlternateContent' && depth < MAX_WRAPPER_DEPTH) {
      const before = out.length;
      for (const branch of node.children) {
        if (branch.type !== 'element' || localName(branch.name) !== 'Choice') continue;
        objectElements(branch, depth + 1, out, state);
        if (state && out.length > before && out.slice(before).some((found) => unreadChartEx(state, found))) out.length = before;
        if (out.length > before) break;
      }
      if (out.length === before) {
        const fallback = child(node, 'Fallback');
        if (fallback) objectElements(fallback, depth + 1, out, state);
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
      const flags = child(child(el, 'nvSpPr'), 'cNvSpPr');
      if (flags && isTrue(flags.attrs.get('txBox'))) shape.textBox = true;
      // The colour a shape style gives its text: usually white on a filled shape.
      const defaults: DrawingTextRun = { text: '' };
      const fontRef = child(child(el, 'style'), 'fontRef');
      const fontColor = firstColor(themeColors, fontRef);
      if (fontColor) defaults.color = fontColor;
      const body = child(el, 'txBody');
      const text = readText(state, body, defaults);
      if (text) shape.text = text;
      // What text typed into it would look like, and whether its text can be replaced: not of a shape that shows the text of a cell.
      if (body && depth === 0) {
        const look = readTextLook(state, body, defaults);
        if (look) shape.textLook = look;
        const link = el.attrs.get('textlink');
        if ((link === undefined || link === '') && plainBody(body)) shape.plainText = true;
      }
    }
    // A shape of the drawing itself (not of a group) that is just what the writer makes of its model can be changed in every way.
    if (depth === 0 && writtenAsModel(el, shape)) shape.known = true;
    return shape;
  }
  if (kind === 'grpSp') {
    readNames(child(el, 'nvGrpSpPr'), base);
    const transform = readTransform(child(el, 'grpSpPr'), base);
    if (transformOut) transformOut.value = transform;
    const group: DrawingObject = { ...base, type: 'group', children: [] };
    if (depth >= MAX_GROUP_DEPTH) return group;
    const members: XmlElement[] = [];
    objectElements(el, 0, members, state);
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
  // The macro a frame is assigned: never run, kept where it is, and not carried to a copy of the chart.
  const macro = kind === 'graphicFrame' ? el.attrs.get('macro') : undefined;
  const transform = readTransform(el, base);
  if (transformOut) transformOut.value = transform;
  const data = child(child(el, 'graphic'), 'graphicData');
  const uri = (data && data.attrs.get('uri')) || '';
  const what = /\/chart(ex)?$|\/chart\b/.test(uri) ? 'chart' : /diagram/.test(uri) ? 'diagram' : 'object';
  if (what === 'chart' && state.context.charts) {
    // <c:chart r:id>: the part that holds the chart. One that was not read (another kind, damaged, missing) stays a placeholder.
    const reference = child(data, 'chart');
    const id = reference ? prefixedAttr(reference, 'id') : undefined;
    const rel = id !== undefined ? state.relById.get(id) : undefined;
    const chart = rel && !rel.external && rel.part ? state.context.charts.get(rel.part) : undefined;
    if (chart && rel && rel.part) {
      const facts = state.context.chartsKnown ? state.context.chartsKnown.get(rel.part) : undefined;
      if (!facts) return macro !== undefined && macro !== '' ? { ...base, type: 'chart', chart, part: rel.part, macro: macro.slice(0, 2000) } : { ...base, type: 'chart', chart, part: rel.part };
      const made: DrawingObject = { ...base, type: 'chart', chart, part: rel.part };
      if (macro !== undefined && macro !== '') made.macro = macro.slice(0, 2000);
      if (facts.known !== undefined) made.known = facts.known;
      if (facts.seriesFixed) made.seriesFixed = true;
      if (facts.loose !== undefined) made.loose = facts.loose.slice();
      return made;
    }
  }
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

/**
 * The anchors of a drawing part in the order they are shown. With `tops`, the child of the root
 * each stands in is noted beside it: the anchor itself, or the wrapper that holds it.
 */
function anchorElements(parent: XmlElement, depth: number, out: XmlElement[], tops?: XmlElement[], top?: XmlElement): void {
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    const local = localName(node.name);
    if (Object.prototype.hasOwnProperty.call(ANCHOR_NAMES, local)) {
      out.push(node);
      if (tops) tops.push(top || node);
    } else if (local === 'AlternateContent' && depth < MAX_WRAPPER_DEPTH) {
      // The same object in a newer and an older form, each with its own anchor: one of them is shown.
      const before = out.length;
      for (const branch of node.children) {
        if (branch.type !== 'element' || localName(branch.name) !== 'Choice') continue;
        anchorElements(branch, depth + 1, out, tops, top || node);
        if (out.length > before) break;
      }
      if (out.length === before) {
        const fallback = child(node, 'Fallback');
        if (fallback) anchorElements(fallback, depth + 1, out, tops, top || node);
      }
    }
  }
}

/** Where each drawing that was read stands in its file: see drawingSource(). */
const SOURCES = new WeakMap<SheetDrawing, { part?: string; index: number }>();

/**
 * Where a drawing that this reader made stands in the file: its drawing part (as the package
 * spells it) and which of the anchors of that part it is, counted as the reader takes them.
 * What lets the writer find the anchor of a drawing that was moved or deleted. Undefined for a
 * drawing no reader made.
 */
export function drawingSource(drawing: SheetDrawing): { part?: string; index: number } | undefined {
  return SOURCES.get(drawing);
}

/**
 * The anchors of a drawing part with what they hold, bottom one first. Never
 * throws for what the part states: an anchor that cannot be read is left out
 * with a note in `context.warnings`.
 */
export function parseDrawing(root: XmlElement, context: DrawingReadContext): SheetDrawing[] {
  const state: ReadState = { context, relById: new Map(context.rels.map((rel) => [rel.id, rel])), objects: 0 };
  const anchors: XmlElement[] = [];
  const tops: XmlElement[] = [];
  anchorElements(root, 0, anchors, tops);
  // How many anchors each wrapper holds: one that holds several cannot go with one of them.
  const held = new Map<XmlElement, number>();
  for (const top of tops) held.set(top, (held.get(top) || 0) + 1);
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
      objectElements(anchors[i], 0, content, state);
      const object = anchor && content.length > 0 ? readObject(state, content[0], 0) : undefined;
      if (anchor && object) {
        drawing = (held.get(tops[i]) || 0) > 1 ? { anchor, object, wrapped: true } : { anchor, object };
        // The object in a wrapper of alternatives (inside its anchor, or its anchor inside one): the file holds another form of it.
        if (tops[i] !== anchors[i] || anchors[i].children.indexOf(content[0]) < 0) drawing.alternative = true;
      }
    } catch (error) {
      // One object that cannot be read does not take the others with it.
      drawing = undefined;
    }
    if (drawing) {
      out.push(drawing);
      SOURCES.set(drawing, { index: i });
    } else skipped++;
  }
  if (skipped > 0) context.warnings.push(`Drawing: ${skipped} object(s) could not be read and are not shown`);
  return out;
}

/** Parts larger than this are not read; the sheet is shown without its drawing. */
const MAX_PART_BYTES = 64 * 1024 * 1024;

/** Bytes of drawing parts read for one workbook, all sheets together; a drawing beyond them is not shown. */
export const MAX_WORKBOOK_DRAWING_BYTES = 256 * 1024 * 1024;
export { MAX_CHART_PART_BYTES, MAX_WORKBOOK_CHART_BYTES };

/**
 * The charts a drawing part refers to, as far as `reading` reads them. A chart that could
 * have been drawn but was not read (a part that is no XML or holds a DTD, one that is too
 * large, one beyond the limits of the sheet or the workbook) is said in `warnings`: its
 * frame shows a placeholder. Never throws.
 */
async function readCharts(
  pkg: OpcPackage,
  rels: Relationship[],
  themeColors: string[],
  reading: ChartReading,
  sheetBudget: { left: number },
  warnings: string[],
  kinds: Map<string, ChartPartFacts>,
): Promise<Map<string, Chart>> {
  const charts = new Map<string, Chart>();
  const limit = reading.maxPartBytes !== undefined && reading.maxPartBytes >= 0 ? reading.maxPartBytes : MAX_CHART_PART_BYTES;
  const { names, known, failed, written } = partsKnown(pkg, reading);
  if (!(typeof reading.bytes === 'number' && reading.bytes >= 0)) reading.bytes = MAX_WORKBOOK_CHART_BYTES;
  const warn = (text: string) => {
    if (warnings.length < MAX_CHART_WARNINGS) warnings.push(text);
  };
  let beyond = 0;
  for (const rel of rels) {
    const type = rel.type.slice(rel.type.lastIndexOf('/') + 1);
    if ((type !== 'chart' && type !== 'chartEx') || rel.external || !rel.part || charts.has(rel.part)) continue;
    // A part of the newer chart types is read by its own reader, and is only shown.
    const extended = type === 'chartEx';
    // The part as the package spells it: however a relationship spells it, it is read and counted once for the workbook.
    const part = names(rel.part);
    if (part === undefined) continue;
    const made = known.get(part);
    if (made !== undefined) {
      if (made) {
        charts.set(rel.part, made);
        const kind = written.get(part);
        if (kind) kinds.set(rel.part, kind);
      } else if (failed.has(part)) warn(failed.get(part) as string);
      continue;
    }
    if (sheetBudget.left <= 0 || reading.charts <= 0) {
      beyond++;
      continue;
    }
    sheetBudget.left--;
    reading.charts--;
    let chart: Chart | undefined;
    try {
      const bytes = await pkg.readPart(part);
      if (bytes.length > limit) throw new Error('the part is too large');
      if (bytes.length > reading.bytes) throw new Error('the charts of the workbook hold too many bytes');
      reading.bytes -= bytes.length;
      const root = parseXml(bytes).root;
      const pointsBefore = reading.points;
      chart = extended ? (reading.readEx ? reading.readEx(root, themeColors, reading) : undefined) : reading.read(root, themeColors, reading);
      if (chart && extended) noteChartPart(chart, { bytes: bytes.length, points: Math.max(0, pointsBefore - reading.points) });
      else if (chart) {
        const source = pivotSourceIn(root);
        if (source !== undefined) chart.pivotSource = source;
        // What the part holds beside the model decides how far the chart can be changed (DrawingChart.known).
        let related = true;
        let stray = false;
        try {
          const partRels = await pkg.relationships(part);
          related = partRels.length > 0;
          stray = partRels.some((other) => !other.external && !!other.part && /(^|\/)_rels\/[^/]*\.rels$/i.test(other.part));
        } catch (error) {
          // Relationships that cannot be read are relationships: the part is not written anew.
        }
        // What the budgets of the workbook were charged for it: a copy of the chart costs as much again.
        const points = Math.max(0, pointsBefore - reading.points);
        noteChartPart(chart, stray ? { bytes: bytes.length, points, stray: true } : { bytes: bytes.length, points });
        // A part in an encoding that is not read (see xmlEncodingRead()) is shown; nothing is written from its text.
        const kind = xmlEncodingRead(bytes) ? chartKnown(root, chart, related) || chartLoose(root, chart) : undefined;
        if (kind) {
          written.set(part, kind);
          kinds.set(rel.part, kind);
        }
      }
    } catch (error) {
      // A chart part that cannot be read (not XML, a DTD, too much of it) leaves its frame a placeholder.
      const text = `Chart could not be read: ${error instanceof Error ? error.message : String(error)}`;
      failed.set(part, text);
      warn(text);
    }
    known.set(part, chart || null);
    if (chart) charts.set(rel.part, chart);
  }
  if (beyond > 0) warnings.push(`Charts beyond the limit of a sheet (${MAX_DRAWING_CHARTS}) or a workbook are not shown: ${beyond}`);
  return charts;
}

/** The table a pivot chart names (c:pivotSource/c:name, wherever it stands in the part); undefined for a chart that is none. */
function pivotSourceIn(root: XmlElement): string | undefined {
  const stack: XmlElement[] = [root];
  let found: string | undefined;
  const most = PIVOT_SCAN_LIMITS.chartNodes;
  for (let seen = 0; stack.length > 0; seen++) {
    // A part too large to be looked through to the end may be a pivot chart: of no table that is known (every cache is refused).
    if (seen >= most) return '';
    const el = stack.pop() as XmlElement;
    if (localName(el.name) === 'pivotSource') {
      const name = el.children.find((node): node is XmlElement => node.type === 'element' && localName(node.name) === 'name');
      let text = '';
      if (name) for (const node of name.children) if (node.type === 'text') text += node.value;
      // Two sources that name different tables, or one that names none: '' (no table is known).
      found = found === undefined ? text : found === text ? found : '';
      continue;
    }
    for (const node of el.children) if (node.type === 'element') stack.push(node);
  }
  return found;
}

/** Warnings about charts of one drawing that are listed one by one. */
const MAX_CHART_WARNINGS = 20;

/** What reading the drawings of a sheet says of their parts beside the drawings themselves. */
export interface DrawingPartFacts {
  /** The drawing parts of the sheet, as the package spells them. */
  parts: string[];
  /** Why the drawing parts take no change (see Sheet.drawingsFixed), when one of them does not. */
  fixed?: string;
}

/**
 * The drawings of a sheet from the parts its relationships name; undefined
 * for a sheet without any. Parts of pictures that are shown are added to
 * `media`. With `charts` the chart parts of the drawings are read too;
 * without it every chart is a placeholder. Never throws for what the parts hold.
 */
export async function readSheetDrawings(
  pkg: OpcPackage,
  sheetRels: Relationship[],
  themeColors: string[],
  media: Set<string>,
  warnings: string[],
  charts?: ChartReading,
  facts?: DrawingPartFacts,
): Promise<SheetDrawing[] | undefined> {
  let drawings: SheetDrawing[] = [];
  const chartBudget = { left: MAX_DRAWING_CHARTS };
  const names = charts ? partsKnown(pkg, charts).names : partNames(pkg);
  if (charts && !(typeof charts.drawingBytes === 'number' && charts.drawingBytes >= 0)) charts.drawingBytes = MAX_WORKBOOK_DRAWING_BYTES;
  /** The drawing parts of this sheet that were read, as the package spells them: one that is named again, in whatever spelling, is not read again. */
  const seen = new Set<string>();
  for (const rel of sheetRels) {
    if (rel.type.slice(rel.type.lastIndexOf('/') + 1) !== 'drawing' || rel.external || !rel.part) continue;
    const part = names(rel.part);
    if (part === undefined || seen.has(part)) continue;
    seen.add(part);
    if (facts) facts.parts.push(part);
    try {
      const bytes = await pkg.readPart(part);
      if (bytes.length > MAX_PART_BYTES) throw new Error('the part is too large');
      // A drawing several sheets name is read for each of them: what that costs is bounded for the workbook.
      if (charts && charts.drawingBytes !== undefined) {
        if (bytes.length > charts.drawingBytes) throw new Error('the drawings of the workbook hold too many bytes');
        charts.drawingBytes -= bytes.length;
      }
      const rels = await pkg.relationships(rel.part);
      const kinds = new Map<string, ChartPartFacts>();
      const found = charts ? await readCharts(pkg, rels, themeColors, charts, chartBudget, warnings, kinds) : undefined;
      const root = parseXml(bytes).root;
      if (facts && facts.fixed === undefined && nestedBeyond(root, MAX_PATCH_DEPTH)) facts.fixed = 'its drawing is nested too deeply';
      const read = parseDrawing(root, { rels, themeColors, hasPart: (part) => pkg.hasPart(part), charts: found, chartsKnown: kinds, warnings });
      // Where each anchor is in the file: what the writer finds it by.
      for (const drawing of read) {
        const source = SOURCES.get(drawing);
        if (source) source.part = part;
      }
      drawings = drawings.concat(read.slice(0, MAX_SHEET_DRAWINGS - drawings.length));
    } catch (error) {
      warnings.push(`Drawing could not be read: ${error instanceof Error ? error.message : String(error)}`);
      // What cannot be read cannot be changed either.
      if (facts && facts.fixed === undefined) facts.fixed = 'its drawing cannot be read';
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
