// A canvas that writes a page of a PDF. The painters of the engine draw on a
// CanvasRenderingContext2D; this is the part of that interface they use,
// turned into the operators of a content stream: paths, fills and strokes
// as paths, text as text in an embedded font, pictures as pictures. The same
// painter that fills the pixels of a page then writes the page that stays
// sharp at every size.
//
// The promise of this canvas is that the page it writes looks like the page
// the real canvas paints, or that it says it cannot: whatever is asked of it
// that it does not do exactly (a gradient, a font without a file, a character
// its font lacks, text whose shaping moves the glyphs, a member of the
// interface it does not have) is remembered in `unsupported`, and the page is
// then written as a picture instead. It never throws for that, because a
// painter may catch what it throws and draw something else.
//
// Nothing a file says reaches the content stream as text: numbers are
// written by pdfNumber(), text as the hexadecimal glyph numbers of its font,
// names of resources are made here.
//
// Coordinates: the canvas is `width` x `height` pixels, as the real one would
// be; a point is transformed by the matrix of the canvas when it is added to
// the path, as a canvas does, and written in points of the page with the y
// axis turned upwards. The content stream holds no `cm` but for pictures.

import type { PdfImage } from './image';
import { applyLigatures, fontLigatures, type PdfFontFace } from './sfnt';
import { glyphUsable } from './subset';
import { PdfError, pdfNumber, type PdfRef } from './writer';

/** A picture a page draws. The document that writes it remembers where (`ref`), so that it is in the file once. */
export interface PdfPicture {
  /** Size of what drawImage() was given, in its own pixels. */
  width: number;
  height: number;
  /** What the picture is; let go of once it is written. */
  image: PdfImage | undefined;
  ref?: PdfRef;
}

export interface PdfCanvasFonts {
  /** The face text in a CSS font is painted with and its size in px, or why no font file stands for it. */
  resolve(font: string): { face: PdfFontFace; size: number } | { reason: string };
}

/** What a canvas says of a text without kerning: its advance in px, and its picture as a digest - two texts with the same `ink` are painted the same. */
export interface PdfInk {
  width: number;
  ink: string;
}

export interface PdfCanvasMeasure {
  /** The width the real canvas measures for a text in a CSS font, in px, without letter spacing. */
  width(font: string, text: string): number;
  /**
   * What the real canvas measures and paints for a text with kerning turned
   * off; undefined where it cannot turn kerning off or its pixels cannot be
   * read. Asked to find out which ligatures the browser draws; without it
   * text in which a font could join letters leaves its page a picture.
   */
  plain?(font: string, text: string): PdfInk | undefined;
  /**
   * How far the alphabetic baseline lies below the line text is placed on
   * with the given `textBaseline`, in px (negative: above it). Without it
   * only text on the alphabetic baseline can be written.
   */
  baseline?(font: string, baseline: string): number;
}

export interface PdfCanvasImages {
  /** The picture a source of drawImage() is, the same object for the same source; undefined when it cannot be read. */
  read(source: unknown): PdfPicture | undefined;
  /**
   * Told when a page was painted: `kept` says whether the page is written
   * with what was read for it. What was read for a page that is not, and
   * for no page before it, need not be held any longer.
   */
  pageEnded?(kept: boolean): void;
}

export interface PdfCanvasOptions {
  /** Size of the canvas in its pixels. */
  width: number;
  height: number;
  /** Size of the page in points. */
  pageWidth: number;
  pageHeight: number;
  fonts: PdfCanvasFonts;
  measure: PdfCanvasMeasure;
  images?: PdfCanvasImages;
}

/** What a canvas wrote: the content of a page and what it needs. Names in the content: /F<n> fonts, /I<n> pictures, /G<n> states. */
export interface PdfVectorPage {
  content: string;
  /** The glyphs used of each font, with the text each one stands for (one character; several for a ligature). */
  fonts: { face: PdfFontFace; glyphs: Map<number, string> }[];
  pictures: PdfPicture[];
  /** Opacities: of fills (and pictures) or of strokes. */
  alphas: { stroke: boolean; value: number }[];
}

/** Average room between what the font says a run is wide and what the canvas measured, per glyph, in font sizes: kerning fits, shaping does not. */
export const TEXT_FIT_SHARE = 0.06;
/** How far one glyph may stand from where its font alone would put it after the glyph before it, in font sizes: a kerning pair, or a ligature the browser drew. */
export const GLYPH_FIT_SHARE = 0.3;
/** Glyphs that are written as one piece of text, each placed where the canvas has it; a longer run is written in pieces. */
export const MAX_PLACED_GLYPHS = 256;

/** A glyph of a run, the text it stands for (several characters for a ligature) and where in the run that text ends. */
interface Cluster {
  glyph: number;
  text: string;
  end: number;
}

/** Letters of scripts other than Latin: the rules a font has for them are not read. */
function otherScript(code: number): boolean {
  return (code >= 0x370 && code <= 0x1dff) || (code >= 0x1f00 && code <= 0x1fff) || (code >= 0x2c00 && code <= 0xa71f) || (code >= 0xa800 && code <= 0xd7ff) || (code >= 0x10000 && code <= 0x1efff);
}
/** Saves that may be open at once: readers of PDF files do not nest deeper. */
const MAX_SAVES = 24;
/** Bytes of content one page may have. */
const MAX_CONTENT = 48 * 1024 * 1024;
const MAX_TEXT_UNITS = 100000;
const MAX_DASHES = 64;

type Matrix = [number, number, number, number, number, number];

interface Paint {
  r: number;
  g: number;
  b: number;
  a: number;
}

interface State {
  m: Matrix;
  fillText: string;
  strokeText: string;
  fill: Paint | undefined;
  stroke: Paint | undefined;
  lineWidth: number;
  lineCap: string;
  lineJoin: string;
  miterLimit: number;
  dash: number[];
  dashOffset: number;
  alpha: number;
  font: string;
  align: string;
  baseline: string;
  spacing: number;
  /** What the content stream has in force, as it was written: empty when not known. */
  out: { fill: string; stroke: string; width: string; cap: string; join: string; miter: string; dash: string; ca: string; CA: string; font: string; tc: string };
}

const NAMED: { [name: string]: [number, number, number, number] } = {
  black: [0, 0, 0, 1],
  white: [255, 255, 255, 1],
  red: [255, 0, 0, 1],
  green: [0, 128, 0, 1],
  blue: [0, 0, 255, 1],
  gray: [128, 128, 128, 1],
  grey: [128, 128, 128, 1],
  silver: [192, 192, 192, 1],
  yellow: [255, 255, 0, 1],
  orange: [255, 165, 0, 1],
  transparent: [0, 0, 0, 0],
};

const finite = (value: unknown): value is number => typeof value === 'number' && isFinite(value);

/** Characters a reason is shown without: controls, and the marks that turn the direction of text around. */
const UNSHOWN = /[\u0000-\u0008\u000b-\u001f\u007f-\u009f\u202a-\u202e\u2066-\u2069]/g;
/** Characters of a reason. */
export const MAX_REASON = 300;

/**
 * A reason as plain text: without control characters (a tab or a line end is
 * a space) and no longer than MAX_REASON characters. A reason may quote what
 * a file says (the name of a font): whoever shows it shows it as text.
 */
export function cleanReason(text: string): string {
  if (typeof text !== 'string') return '';
  return text
    .slice(0, MAX_REASON * 4)
    .replace(UNSHOWN, '')
    .replace(/[\t\n]/g, ' ')
    .slice(0, MAX_REASON);
}

/**
 * The text a glyph stands for when it stands for more than one: the one that
 * comes first by its code points - the space, not the space that does not
 * break; the Latin letter, not its look-alike; the letters of a ligature, not
 * the one character there is for it.
 */
export function plainerText(a: string, b: string): string {
  const length = Math.min(a.length, b.length);
  for (let i = 0; i < length; i++) {
    const first = a.charCodeAt(i);
    const second = b.charCodeAt(i);
    if (first !== second) return first < second ? a : b;
  }
  return a.length <= b.length ? a : b;
}

/** A CSS colour of the forms the painters use; undefined for anything else. */
export function parseCssColor(value: unknown): Paint | undefined {
  if (typeof value !== 'string' || value.length > 64) return undefined;
  const text = value.trim().toLowerCase();
  const hex = /^#([0-9a-f]{3,8})$/.exec(text);
  if (hex) {
    const digits = hex[1];
    const part = (at: number, size: number): number => (size === 1 ? parseInt(digits.charAt(at), 16) * 17 : parseInt(digits.substr(at, 2), 16));
    if (digits.length === 3 || digits.length === 4) return { r: part(0, 1), g: part(1, 1), b: part(2, 1), a: digits.length === 4 ? part(3, 1) / 255 : 1 };
    if (digits.length === 6 || digits.length === 8) return { r: part(0, 2), g: part(2, 2), b: part(4, 2), a: digits.length === 8 ? part(6, 2) / 255 : 1 };
    return undefined;
  }
  // The two ways CSS writes it, never mixed: with commas all through, or with spaces and a slash before the opacity.
  const number = '(\\d+(?:\\.\\d+)?|\\.\\d+)';
  const call =
    new RegExp('^rgba?\\(\\s*' + number + '\\s*,\\s*' + number + '\\s*,\\s*' + number + '\\s*(?:,\\s*' + number + '(%?)\\s*)?\\)$').exec(text) ||
    new RegExp('^rgba?\\(\\s*' + number + '\\s+' + number + '\\s+' + number + '\\s*(?:/\\s*' + number + '(%?)\\s*)?\\)$').exec(text);
  if (call) {
    const channel = (digits: string): number => Math.min(255, Math.round(Number(digits)));
    let alpha = call[4] === undefined ? 1 : Number(call[4]) / (call[5] === '%' ? 100 : 1);
    if (!(alpha >= 0)) return undefined;
    if (alpha > 1) alpha = 1;
    return { r: channel(call[1]), g: channel(call[2]), b: channel(call[3]), a: alpha };
  }
  if (Object.prototype.hasOwnProperty.call(NAMED, text)) {
    const named = NAMED[text];
    return { r: named[0], g: named[1], b: named[2], a: named[3] };
  }
  return undefined;
}

/** Characters whose glyphs a font places by rules this canvas does not know: text with one of them is not written as text. */
function needsShaping(code: number): boolean {
  if (code < 0x300) return false;
  return (
    (code >= 0x300 && code <= 0x36f) ||
    (code >= 0x483 && code <= 0x489) ||
    (code >= 0x591 && code <= 0x8ff) ||
    (code >= 0x900 && code <= 0x109f) ||
    (code >= 0x1100 && code <= 0x11ff) ||
    (code >= 0x135d && code <= 0x135f) ||
    (code >= 0x1700 && code <= 0x18af) ||
    (code >= 0x1900 && code <= 0x1dff) ||
    (code >= 0x200c && code <= 0x200f) ||
    (code >= 0x202a && code <= 0x202e) ||
    (code >= 0x2066 && code <= 0x2069) ||
    (code >= 0x20d0 && code <= 0x20ff) ||
    (code >= 0x302a && code <= 0x302f) ||
    (code >= 0x3099 && code <= 0x309a) ||
    (code >= 0xa800 && code <= 0xabff) ||
    (code >= 0xd800 && code <= 0xdfff) ||
    (code >= 0xfb1d && code <= 0xfdff) ||
    (code >= 0xfe00 && code <= 0xfe0f) ||
    (code >= 0xfe20 && code <= 0xfe2f) ||
    (code >= 0xfe70 && code <= 0xfeff) ||
    (code >= 0x10800 && code <= 0x10fff) ||
    (code >= 0x11000 && code <= 0x11fff) ||
    (code >= 0x1e800 && code <= 0x1efff) ||
    (code >= 0x1f3fb && code <= 0x1f3ff) ||
    (code >= 0xe0000 && code <= 0xe0fff)
  );
}

/** Characters a browser paints nothing for and gives no width. */
const isIgnored = (code: number): boolean => code === 0xad || code === 0x200b || code === 0x2060;

const HEX = '0123456789abcdef';
const hex4 = (value: number): string => HEX.charAt((value >> 12) & 15) + HEX.charAt((value >> 8) & 15) + HEX.charAt((value >> 4) & 15) + HEX.charAt(value & 15);

const CAPS: { [cap: string]: string } = { butt: '0', round: '1', square: '2' };
const JOINS: { [join: string]: string } = { miter: '0', round: '1', bevel: '2' };

/** What a gradient or a pattern of this canvas is: something that cannot be painted with. */
const NO_PAINT = { addColorStop: (): void => undefined };

/** See the head of the file. Painters get it through createPdfCanvas(), which also answers for the members it does not have. */
export class PdfCanvas {
  /** Why the page cannot be written as it was painted; undefined while it can. The first reason stays. */
  unsupported: string | undefined;

  private readonly options: PdfCanvasOptions;
  /** Points per pixel of the canvas. */
  private readonly kx: number;
  private readonly ky: number;
  private state: State;
  private readonly stack: State[] = [];
  private readonly parts: string[] = [];
  private size = 0;
  /** Whether anything was painted yet. Until then the page is paper: white, as a page of a PDF is without being painted. */
  private marked = false;
  /** The current path, in points of the page, and where it stands. */
  private path: string[] = [];
  private point: [number, number] | undefined;
  private start: [number, number] | undefined;
  private readonly fonts: { face: PdfFontFace; glyphs: Map<number, string> }[] = [];
  private readonly pictures: PdfPicture[] = [];
  private readonly alphas: { stroke: boolean; value: number }[] = [];
  private readonly paints = new Map<string, Paint | null>();

  constructor(options: PdfCanvasOptions) {
    this.options = options;
    const sound = options && options.width > 0 && options.height > 0 && options.pageWidth > 0 && options.pageHeight > 0 && isFinite(options.width + options.height + options.pageWidth + options.pageHeight);
    this.kx = sound ? options.pageWidth / options.width : 1;
    this.ky = sound ? options.pageHeight / options.height : 1;
    this.state = {
      m: [1, 0, 0, 1, 0, 0],
      fillText: '#000000',
      strokeText: '#000000',
      fill: { r: 0, g: 0, b: 0, a: 1 },
      stroke: { r: 0, g: 0, b: 0, a: 1 },
      lineWidth: 1,
      lineCap: 'butt',
      lineJoin: 'miter',
      miterLimit: 10,
      dash: [],
      dashOffset: 0,
      alpha: 1,
      font: '10px sans-serif',
      align: 'start',
      baseline: 'alphabetic',
      spacing: 0,
      out: { fill: '', stroke: '', width: '', cap: '', join: '', miter: '', dash: '', ca: '', CA: '', font: '', tc: '' },
    };
    if (!sound) this.refuse('the page has no size');
  }

  /** Says that the page cannot be written as it is painted. Painting goes on, and writes nothing more. */
  refuse(reason: string): void {
    if (this.unsupported === undefined) this.unsupported = cleanReason(reason) || 'the page cannot be written as it is painted';
  }

  private emit(part: string): void {
    if (this.unsupported !== undefined) return;
    this.size += part.length + 1;
    if (this.size > MAX_CONTENT) return this.refuse('the page holds too much to be written as drawings');
    this.parts.push(part);
  }

  /** A number of the content stream; refuses the page for one that cannot be written. */
  private num(value: number, digits = 3): string {
    if (!isFinite(value) || Math.abs(value) > 1e9) {
      this.refuse('something is drawn at a place no page has');
      return '0';
    }
    try {
      return pdfNumber(value, digits);
    } catch (error) {
      if (!(error instanceof PdfError)) throw error;
      this.refuse('something is drawn at a place no page has');
      return '0';
    }
  }

  /** A point of the canvas (in its units, under its matrix) in points of the page. */
  private at(x: number, y: number): [number, number] {
    const m = this.state.m;
    return [(m[0] * x + m[2] * y + m[4]) * this.kx, this.options.pageHeight - (m[1] * x + m[3] * y + m[5]) * this.ky];
  }

  private xy(point: [number, number]): string {
    return this.num(point[0]) + ' ' + this.num(point[1]);
  }

  // ------------------------------------------------------------------ state

  save(): void {
    if (this.stack.length >= MAX_SAVES) return this.refuse('the painter saves the state of the canvas too many times over');
    const state = this.state;
    this.stack.push({ ...state, m: state.m.slice() as Matrix, dash: state.dash.slice(), out: { ...state.out } });
    this.emit('q');
  }

  restore(): void {
    const state = this.stack.pop();
    if (!state) return;
    this.state = state;
    this.emit('Q');
  }

  // -------------------------------------------------------------- transform

  private multiply(a: number, b: number, c: number, d: number, e: number, f: number): void {
    if (!finite(a) || !finite(b) || !finite(c) || !finite(d) || !finite(e) || !finite(f)) return;
    const m = this.state.m;
    this.state.m = [m[0] * a + m[2] * b, m[1] * a + m[3] * b, m[0] * c + m[2] * d, m[1] * c + m[3] * d, m[0] * e + m[2] * f + m[4], m[1] * e + m[3] * f + m[5]];
  }

  translate(x: number, y: number): void {
    this.multiply(1, 0, 0, 1, x, y);
  }

  scale(x: number, y: number): void {
    this.multiply(x, 0, 0, y, 0, 0);
  }

  rotate(angle: number): void {
    if (!finite(angle)) return;
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    this.multiply(cos, sin, -sin, cos, 0, 0);
  }

  transform(a: number, b: number, c: number, d: number, e: number, f: number): void {
    this.multiply(a, b, c, d, e, f);
  }

  setTransform(a?: number | { a: number; b: number; c: number; d: number; e: number; f: number }, b?: number, c?: number, d?: number, e?: number, f?: number): void {
    if (a === undefined) {
      this.state.m = [1, 0, 0, 1, 0, 0];
      return;
    }
    if (typeof a === 'object' && a !== null) {
      const matrix = a;
      if (!finite(matrix.a) || !finite(matrix.b) || !finite(matrix.c) || !finite(matrix.d) || !finite(matrix.e) || !finite(matrix.f)) return this.refuse('the canvas is given a transform that is none');
      this.state.m = [matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f];
      return;
    }
    if (!finite(a) || !finite(b) || !finite(c) || !finite(d) || !finite(e) || !finite(f)) return;
    this.state.m = [a, b, c, d, e, f];
  }

  resetTransform(): void {
    this.state.m = [1, 0, 0, 1, 0, 0];
  }

  getTransform(): { a: number; b: number; c: number; d: number; e: number; f: number } {
    const m = this.state.m;
    return { a: m[0], b: m[1], c: m[2], d: m[3], e: m[4], f: m[5] };
  }

  /** How much longer a length is on the page than on the canvas; undefined under a matrix that stretches one way more than the other. */
  private lengthScale(): number | undefined {
    const m = this.state.m;
    const a = m[0] * this.kx;
    const b = m[1] * this.ky;
    const c = m[2] * this.kx;
    const d = m[3] * this.ky;
    const first = a * a + b * b;
    const second = c * c + d * d;
    const size = Math.max(first, second);
    // The pixels of a page are whole numbers, so the two axes of the page may differ by a hair: that is no stretch.
    if (Math.abs(first - second) > size * 0.02 || Math.abs(a * c + b * d) > size * 0.02) return undefined;
    return Math.sqrt(Math.abs(a * d - b * c));
  }

  // ----------------------------------------------------------------- styles

  private paintOf(value: unknown): Paint | undefined {
    if (typeof value !== 'string') return undefined;
    let known = this.paints.get(value);
    if (known === undefined) {
      known = parseCssColor(value) || null;
      if (this.paints.size > 4096) this.paints.clear();
      this.paints.set(value, known);
    }
    return known || undefined;
  }

  get fillStyle(): unknown {
    return this.state.fillText;
  }
  set fillStyle(value: unknown) {
    // A paint that cannot be written is only a reason once something is painted with it.
    this.state.fill = this.paintOf(value);
    this.state.fillText = typeof value === 'string' ? value : '';
  }

  get strokeStyle(): unknown {
    return this.state.strokeText;
  }
  set strokeStyle(value: unknown) {
    this.state.stroke = this.paintOf(value);
    this.state.strokeText = typeof value === 'string' ? value : '';
  }

  createLinearGradient(): unknown {
    return NO_PAINT;
  }
  createRadialGradient(): unknown {
    return NO_PAINT;
  }
  createConicGradient(): unknown {
    return NO_PAINT;
  }
  createPattern(): unknown {
    return NO_PAINT;
  }

  get globalAlpha(): number {
    return this.state.alpha;
  }
  set globalAlpha(value: number) {
    if (finite(value) && value >= 0 && value <= 1) this.state.alpha = value;
  }

  get lineWidth(): number {
    return this.state.lineWidth;
  }
  set lineWidth(value: number) {
    if (finite(value) && value > 0) this.state.lineWidth = value;
  }

  get lineCap(): string {
    return this.state.lineCap;
  }
  set lineCap(value: string) {
    if (Object.prototype.hasOwnProperty.call(CAPS, value)) this.state.lineCap = value;
  }

  get lineJoin(): string {
    return this.state.lineJoin;
  }
  set lineJoin(value: string) {
    if (Object.prototype.hasOwnProperty.call(JOINS, value)) this.state.lineJoin = value;
  }

  get miterLimit(): number {
    return this.state.miterLimit;
  }
  set miterLimit(value: number) {
    if (finite(value) && value > 0) this.state.miterLimit = value;
  }

  get lineDashOffset(): number {
    return this.state.dashOffset;
  }
  set lineDashOffset(value: number) {
    if (finite(value)) this.state.dashOffset = value;
  }

  setLineDash(segments: number[]): void {
    if (!Array.isArray(segments)) return this.refuse('the dashes of a line are given in a form that is not read');
    if (segments.length > MAX_DASHES) return this.refuse('a line has too many dashes');
    for (let i = 0; i < segments.length; i++) if (!finite(segments[i]) || segments[i] < 0) return;
    // An odd number of lengths is the list twice.
    this.state.dash = segments.length % 2 === 1 ? segments.concat(segments) : segments.slice();
  }

  getLineDash(): number[] {
    return this.state.dash.slice();
  }

  // What a canvas has and this one only takes in its plain form.
  get globalCompositeOperation(): string {
    return 'source-over';
  }
  set globalCompositeOperation(value: string) {
    if (value !== 'source-over') this.refuse(`drawing is combined in a way that is not written ("${String(value).slice(0, 32)}")`);
  }
  get filter(): string {
    return 'none';
  }
  set filter(value: string) {
    if (value !== 'none') this.refuse('a filter is drawn through');
  }
  get shadowBlur(): number {
    return 0;
  }
  set shadowBlur(value: number) {
    if (value !== 0) this.refuse('a shadow is drawn');
  }
  get shadowOffsetX(): number {
    return 0;
  }
  set shadowOffsetX(value: number) {
    if (value !== 0) this.refuse('a shadow is drawn');
  }
  get shadowOffsetY(): number {
    return 0;
  }
  set shadowOffsetY(value: number) {
    if (value !== 0) this.refuse('a shadow is drawn');
  }
  shadowColor = 'rgba(0, 0, 0, 0)';
  imageSmoothingEnabled = true;
  imageSmoothingQuality = 'low';
  get direction(): string {
    return 'ltr';
  }
  set direction(value: string) {
    if (value !== 'ltr' && value !== 'inherit') this.refuse('text is written from right to left');
  }
  get fontKerning(): string {
    return 'auto';
  }
  set fontKerning(value: string) {
    if (value !== 'auto' && value !== 'normal') this.refuse('the kerning of the text is set');
  }
  get wordSpacing(): string {
    return '0px';
  }
  set wordSpacing(value: string) {
    if (value !== '0px') this.refuse('the spacing of words is set');
  }

  // ------------------------------------------------------------------- path

  beginPath(): void {
    this.path = [];
    this.point = undefined;
    this.start = undefined;
  }

  private move(point: [number, number]): void {
    this.path.push(this.xy(point) + ' m');
    this.point = point;
    this.start = point;
  }

  private line(point: [number, number]): void {
    if (!this.point) return this.move(point);
    this.path.push(this.xy(point) + ' l');
    this.point = point;
  }

  private curve(first: [number, number], second: [number, number], end: [number, number]): void {
    this.path.push(this.xy(first) + ' ' + this.xy(second) + ' ' + this.xy(end) + ' c');
    this.point = end;
  }

  moveTo(x: number, y: number): void {
    if (!finite(x) || !finite(y)) return;
    this.move(this.at(x, y));
  }

  lineTo(x: number, y: number): void {
    if (!finite(x) || !finite(y)) return;
    this.line(this.at(x, y));
  }

  closePath(): void {
    if (!this.point || !this.start) return;
    this.path.push('h');
    // A new subpath starts where the closed one did.
    this.point = this.start;
  }

  bezierCurveTo(x1: number, y1: number, x2: number, y2: number, x: number, y: number): void {
    if (!finite(x1) || !finite(y1) || !finite(x2) || !finite(y2) || !finite(x) || !finite(y)) return;
    const first = this.at(x1, y1);
    if (!this.point) this.move(first);
    this.curve(first, this.at(x2, y2), this.at(x, y));
  }

  quadraticCurveTo(cx: number, cy: number, x: number, y: number): void {
    if (!finite(cx) || !finite(cy) || !finite(x) || !finite(y)) return;
    const control = this.at(cx, cy);
    if (!this.point) this.move(control);
    const from = this.point as [number, number];
    const end = this.at(x, y);
    // The same curve as a cubic one: the control points lie two thirds of the way to the one it has.
    this.curve([from[0] + ((control[0] - from[0]) * 2) / 3, from[1] + ((control[1] - from[1]) * 2) / 3], [end[0] + ((control[0] - end[0]) * 2) / 3, end[1] + ((control[1] - end[1]) * 2) / 3], end);
  }

  rect(x: number, y: number, width: number, height: number): void {
    if (!finite(x) || !finite(y) || !finite(width) || !finite(height)) return;
    this.move(this.at(x, y));
    this.line(this.at(x + width, y));
    this.line(this.at(x + width, y + height));
    this.line(this.at(x, y + height));
    this.closePath();
  }

  arc(x: number, y: number, radius: number, startAngle: number, endAngle: number, counterclockwise?: boolean): void {
    if (!finite(x) || !finite(y) || !finite(radius) || !finite(startAngle) || !finite(endAngle)) return;
    if (radius < 0) return this.refuse('an arc has a radius below nothing');
    const turn = Math.PI * 2;
    let sweep: number;
    if (!counterclockwise) {
      if (endAngle - startAngle >= turn) sweep = turn;
      else {
        sweep = (endAngle - startAngle) % turn;
        if (sweep < 0) sweep += turn;
      }
    } else if (startAngle - endAngle >= turn) sweep = -turn;
    else {
      sweep = (endAngle - startAngle) % turn;
      if (sweep > 0) sweep -= turn;
    }
    const on = (angle: number): [number, number] => this.at(x + radius * Math.cos(angle), y + radius * Math.sin(angle));
    this.line(on(startAngle));
    if (radius === 0 || sweep === 0) return;
    // Quarter turns at most, each as one curve.
    const pieces = Math.max(1, Math.ceil(Math.abs(sweep) / (Math.PI / 2) - 1e-9));
    const step = sweep / pieces;
    const reach = (4 / 3) * Math.tan(step / 4) * radius;
    for (let i = 0; i < pieces; i++) {
      const from = startAngle + step * i;
      const to = from + step;
      this.curve(
        this.at(x + radius * Math.cos(from) - reach * Math.sin(from), y + radius * Math.sin(from) + reach * Math.cos(from)),
        this.at(x + radius * Math.cos(to) + reach * Math.sin(to), y + radius * Math.sin(to) - reach * Math.cos(to)),
        on(to),
      );
    }
  }

  // --------------------------------------------------------------- painting

  private alphaName(stroke: boolean, value: number): string {
    for (let i = 0; i < this.alphas.length; i++) if (this.alphas[i].stroke === stroke && this.alphas[i].value === value) return 'G' + i;
    this.alphas.push({ stroke, value });
    return 'G' + (this.alphas.length - 1);
  }

  private colour(paint: Paint, operator: string, grey: string): string {
    if (paint.r === paint.g && paint.g === paint.b) return pdfNumber(paint.r / 255, 3) + ' ' + grey;
    return pdfNumber(paint.r / 255, 3) + ' ' + pdfNumber(paint.g / 255, 3) + ' ' + pdfNumber(paint.b / 255, 3) + ' ' + operator;
  }

  /** Makes the fill of the content stream that of the canvas. False when nothing would be seen, or the paint cannot be written. */
  private useFill(): boolean {
    const state = this.state;
    const paint = state.fill;
    if (!paint) {
      this.refuse('something is filled with a gradient, a pattern or a colour that is not read');
      return false;
    }
    const alpha = paint.a * state.alpha;
    if (!(alpha > 0)) return false;
    const colour = this.colour(paint, 'rg', 'g');
    if (state.out.fill !== colour) {
      this.emit(colour);
      state.out.fill = colour;
    }
    this.useAlpha(false, alpha);
    return true;
  }

  private useAlpha(stroke: boolean, alpha: number): void {
    const out = this.state.out;
    const current = stroke ? out.CA : out.ca;
    const rounded = Math.round(alpha * 1000) / 1000;
    // Nothing was said yet and nothing needs to be: a page starts opaque.
    if (current === '' && rounded >= 1) return;
    const name = this.alphaName(stroke, rounded);
    if (current === name) return;
    this.emit('/' + name + ' gs');
    if (stroke) out.CA = name;
    else out.ca = name;
  }

  private useStroke(): boolean {
    const state = this.state;
    const paint = state.stroke;
    if (!paint) {
      this.refuse('a line is drawn with a gradient, a pattern or a colour that is not read');
      return false;
    }
    const alpha = paint.a * state.alpha;
    if (!(alpha > 0)) return false;
    const scale = this.lengthScale();
    if (scale === undefined) {
      this.refuse('a line is drawn on a canvas that is stretched one way more than the other');
      return false;
    }
    if (!(scale > 0)) return false;
    const out = state.out;
    const colour = this.colour(paint, 'RG', 'G');
    if (out.stroke !== colour) {
      this.emit(colour);
      out.stroke = colour;
    }
    const width = this.num(state.lineWidth * scale, 4);
    if (out.width !== width) {
      this.emit(width + ' w');
      out.width = width;
    }
    const cap = CAPS[state.lineCap];
    if (out.cap !== cap && !(out.cap === '' && cap === '0')) {
      this.emit(cap + ' J');
      out.cap = cap;
    }
    const join = JOINS[state.lineJoin];
    if (out.join !== join && !(out.join === '' && join === '0')) {
      this.emit(join + ' j');
      out.join = join;
    }
    const miter = this.num(Math.max(1, state.miterLimit), 3);
    if (out.miter !== miter && !(out.miter === '' && miter === '10')) {
      this.emit(miter + ' M');
      out.miter = miter;
    }
    let dash = '[] 0 d';
    let any = false;
    for (let i = 0; i < state.dash.length; i++) if (state.dash[i] > 0) any = true;
    if (any) {
      const lengths: string[] = [];
      for (let i = 0; i < state.dash.length; i++) lengths.push(this.num(state.dash[i] * scale, 4));
      // Lengths that are written as nothing are no dashes.
      if (lengths.some((length) => length !== '0')) dash = '[' + lengths.join(' ') + '] ' + this.num(state.dashOffset * scale, 4) + ' d';
    }
    if (out.dash !== dash && !(out.dash === '' && dash === '[] 0 d')) {
      this.emit(dash);
      out.dash = dash;
    }
    this.useAlpha(true, alpha);
    return true;
  }

  private rule(args: IArguments | unknown[], what: string): string | undefined {
    const first = args[0];
    if (args.length > 1 || (first !== undefined && first !== 'nonzero' && first !== 'evenodd')) {
      this.refuse(`${what} is given a path of its own`);
      return undefined;
    }
    return first === 'evenodd' ? '*' : '';
  }

  fill(): void {
    const rule = this.rule(arguments, 'fill()');
    if (rule === undefined || this.path.length === 0) return;
    if (!this.useFill()) return;
    this.marked = true;
    this.emit(this.path.join(' ') + ' f' + rule);
  }

  stroke(): void {
    if (arguments.length > 0) return this.refuse('stroke() is given a path of its own');
    if (this.path.length === 0) return;
    if (!this.useStroke()) return;
    this.marked = true;
    this.emit(this.path.join(' ') + ' S');
  }

  clip(): void {
    const rule = this.rule(arguments, 'clip()');
    if (rule === undefined) return;
    // Without a path nothing is left to draw into.
    this.emit((this.path.length === 0 ? '0 0 0 0 re' : this.path.join(' ')) + ' W' + rule + ' n');
  }

  private box(x: number, y: number, width: number, height: number): string {
    const m = this.state.m;
    const a = this.at(x, y);
    const c = this.at(x + width, y + height);
    // Upright: a rectangle; turned: its four corners.
    if (m[1] === 0 && m[2] === 0) return this.num(Math.min(a[0], c[0])) + ' ' + this.num(Math.min(a[1], c[1])) + ' ' + this.num(Math.abs(c[0] - a[0])) + ' ' + this.num(Math.abs(c[1] - a[1])) + ' re';
    return this.xy(a) + ' m ' + this.xy(this.at(x + width, y)) + ' l ' + this.xy(c) + ' l ' + this.xy(this.at(x, y + height)) + ' l h';
  }

  fillRect(x: number, y: number, width: number, height: number): void {
    if (!finite(x) || !finite(y) || !finite(width) || !finite(height) || width === 0 || height === 0) return;
    // White on paper nothing was painted on is the paper: what painters of pages start with.
    const paint = this.state.fill;
    if (!this.marked && paint && paint.r === 255 && paint.g === 255 && paint.b === 255 && paint.a === 1 && this.state.alpha === 1) return;
    if (!this.useFill()) return;
    this.marked = true;
    this.emit(this.box(x, y, width, height) + ' f');
  }

  strokeRect(x: number, y: number, width: number, height: number): void {
    if (!finite(x) || !finite(y) || !finite(width) || !finite(height) || (width === 0 && height === 0)) return;
    if (!this.useStroke()) return;
    this.marked = true;
    this.emit(this.box(x, y, width, height) + ' S');
  }

  // ------------------------------------------------------------------- text

  get font(): string {
    return this.state.font;
  }
  set font(value: string) {
    // Whether it can be written is asked when text is painted with it.
    this.state.font = typeof value === 'string' ? value : '';
  }

  get textAlign(): string {
    return this.state.align;
  }
  set textAlign(value: string) {
    if (value === 'left' || value === 'right' || value === 'center' || value === 'start' || value === 'end') this.state.align = value;
  }

  get textBaseline(): string {
    return this.state.baseline;
  }
  set textBaseline(value: string) {
    if (value === 'alphabetic' || value === 'top' || value === 'hanging' || value === 'middle' || value === 'ideographic' || value === 'bottom') this.state.baseline = value;
  }

  get letterSpacing(): string {
    return this.state.spacing + 'px';
  }
  set letterSpacing(value: string) {
    const parsed = /^\s*(-?\d*\.?\d+(?:e-?\d+)?)px\s*$/i.exec(String(value));
    if (!parsed || !isFinite(Number(parsed[1]))) this.refuse('the spacing of letters is given in a form that is not read');
    else this.state.spacing = Number(parsed[1]);
  }

  measureText(text: string): { width: number } {
    const value = String(text);
    let width = 0;
    try {
      width = value === '' ? 0 : this.options.measure.width(this.state.font, value);
    } catch {
      width = NaN;
    }
    if (!finite(width)) {
      this.refuse('text could not be measured');
      width = 0;
    }
    if (this.state.spacing !== 0) width += this.state.spacing * value.length;
    return { width };
  }

  private fontName(face: PdfFontFace): { name: string; glyphs: Map<number, string> } {
    for (let i = 0; i < this.fonts.length; i++) if (this.fonts[i].face === face) return { name: 'F' + i, glyphs: this.fonts[i].glyphs };
    const glyphs = new Map<number, string>();
    this.fonts.push({ face, glyphs });
    return { name: 'F' + (this.fonts.length - 1), glyphs };
  }

  /**
   * The glyphs a browser draws for the glyphs of a text: the ligatures of the
   * font, as far as the browser forms them. Which it forms depends on the
   * language of the text and on its settings, which are not known here, so
   * wherever the language systems of the font disagree the browser is asked:
   * the stretch is painted without kerning as it is and with each way of
   * joining it held apart by joiners, and the way whose picture is that of
   * the stretch itself is the one that was drawn. Undefined when that cannot
   * be told: the page is refused.
   */
  private shape(face: PdfFontFace, value: string, singles: Cluster[]): Cluster[] | undefined {
    const state = this.state;
    const ligatures = fontLigatures(face);
    if (ligatures.state === 'unknown') {
      this.refuse(`the font ${face.name} replaces glyphs by rules that are not read here`);
      return undefined;
    }
    if (ligatures.otherScripts) {
      for (let i = 0; i < singles.length; i++) {
        if (otherScript(singles[i].text.charCodeAt(0))) {
          this.refuse(`a text is of a script the font ${face.name} has rules for that are not read here`);
          return undefined;
        }
      }
    }
    // Spaced letters are not joined: a browser turns ligatures off for them.
    if (ligatures.state !== 'known' || state.spacing !== 0 || singles.length < 2) return singles;
    const glyphs = singles.map((single) => single.glyph);
    // What each language system of the font makes of the glyphs, and what none makes of them.
    const ways: { glyph: number; count: number }[][] = [glyphs.map((glyph) => ({ glyph, count: 1 }))];
    const seen: { [signature: string]: true } = { [ways[0].map((entry) => entry.glyph).join()]: true };
    for (let v = 0; v < ligatures.variants.length; v++) {
      const way = applyLigatures(ligatures.variants[v], glyphs);
      const signature = way.map((entry) => entry.glyph).join();
      if (!seen[signature]) {
        seen[signature] = true;
        ways.push(way);
      }
    }
    if (ways.length === 1) return singles;
    // The stretches of the text any of them joins: where they overlap they are one stretch.
    const joined: boolean[] = [];
    const starts = ways.map((way) => {
      const at: number[] = [];
      let index = 0;
      for (let i = 0; i < way.length; i++) {
        at.push(index);
        if (way[i].count > 1) for (let k = 0; k + 1 < way[i].count; k++) joined[index + k] = true;
        index += way[i].count;
      }
      return at;
    });
    const out: Cluster[] = [];
    for (let from = 0; from < singles.length; ) {
      if (!joined[from]) {
        out.push(singles[from]);
        from++;
        continue;
      }
      let to = from + 1;
      while (joined[to - 1]) to++;
      // The ways to write the stretch, each once.
      const candidates: Cluster[][] = [];
      for (let w = 0; w < ways.length; w++) {
        const candidate: Cluster[] = [];
        for (let i = 0; i < ways[w].length; i++) {
          const start = starts[w][i];
          if (start < from || start >= to) continue;
          let text = '';
          for (let k = 0; k < ways[w][i].count; k++) text += singles[start + k].text;
          candidate.push({ glyph: ways[w][i].glyph, text, end: singles[start + ways[w][i].count - 1].end });
        }
        if (!candidates.some((other) => other.length === candidate.length && other.every((cluster, index) => cluster.glyph === candidate[index].glyph))) candidates.push(candidate);
      }
      const chosen = candidates.length === 1 ? candidates[0] : this.drawnWay(face, value.slice(from === 0 ? 0 : singles[from - 1].end, singles[to - 1].end), candidates);
      if (!chosen) {
        this.refuse(`it cannot be told which ligatures of the font ${face.name} the browser draws in a text`);
        return undefined;
      }
      for (let i = 0; i < chosen.length; i++) out.push(chosen[i]);
      from = to;
    }
    return out;
  }

  /** Which of several ways to write a stretch of text the browser drew: see shape(). */
  private drawnWay(face: PdfFontFace, text: string, candidates: Cluster[][]): Cluster[] | undefined {
    const measure = this.options.measure;
    const font = this.state.font;
    const resolved = this.options.fonts.resolve(font);
    if (!measure.plain || !('face' in resolved)) return undefined;
    const size = resolved.size;
    const slack = size * 0.002 + 0.01;
    const same = (a: number, b: number): boolean => Math.abs(a - b) <= slack;
    let natural: PdfInk | undefined;
    try {
      natural = measure.plain(font, text);
    } catch {
      natural = undefined;
    }
    if (!natural || !finite(natural.width) || typeof natural.ink !== 'string') return undefined;
    let best: Cluster[] | undefined;
    let tied = false;
    for (let c = 0; c < candidates.length; c++) {
      const candidate = candidates[c];
      // Held apart by joiners that keep letters from being joined, the stretch is drawn this way and no other.
      let apart = '';
      let advance = 0;
      for (let i = 0; i < candidate.length; i++) {
        apart += (i > 0 ? '\u200c' : '') + candidate[i].text;
        advance += face.advance(candidate[i].glyph);
      }
      let ink: PdfInk | undefined;
      try {
        ink = apart === text ? natural : measure.plain(font, apart);
      } catch {
        ink = undefined;
      }
      // As wide as these glyphs are, and painted as the stretch itself is.
      if (!ink || !same(ink.width, (advance * size) / face.unitsPerEm) || !same(ink.width, natural.width) || ink.ink !== natural.ink) continue;
      // Of two ways that are painted the same the one with fewer ligatures is taken: its glyphs, each at its
      // place, are that picture (a way that joins everything is the stretch itself, and proves nothing).
      if (!best || candidate.length > best.length) {
        best = candidate;
        tied = false;
      } else if (candidate.length === best.length) tied = true;
    }
    return tied ? undefined : best;
  }

  fillText(text: string, x: number, y: number, maxWidth?: number): void {
    if (this.unsupported !== undefined) return;
    if (maxWidth !== undefined) return this.refuse('text is squeezed into a width');
    if (!finite(x) || !finite(y)) return;
    const value = String(text);
    if (value === '') return;
    if (value.length > MAX_TEXT_UNITS) return this.refuse('a text is too long to be written');
    const state = this.state;
    const resolved = this.options.fonts.resolve(state.font);
    if (!('face' in resolved)) return this.refuse(resolved.reason);
    const face = resolved.face;
    const size = resolved.size;
    const unit = size / face.unitsPerEm;
    // The glyphs of the characters, one each, and where in the text each one ends.
    const singles: Cluster[] = [];
    for (let i = 0; i < value.length; i++) {
      let code = value.charCodeAt(i);
      if (code >= 0xd800 && code <= 0xdbff && i + 1 < value.length) {
        const next = value.charCodeAt(i + 1);
        if (next >= 0xdc00 && next <= 0xdfff) {
          code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
          i++;
        }
      }
      // A canvas paints a space for what breaks or tabs a line.
      if (code === 9 || code === 10 || code === 11 || code === 12 || code === 13) code = 0x20;
      if (isIgnored(code)) continue;
      if (code < 0x20 || (code >= 0x7f && code < 0xa0)) return this.refuse('a text holds a control character');
      if (needsShaping(code)) return this.refuse('a text holds characters whose glyphs are placed by rules that are not followed here (a script that joins or stacks its letters, or text from right to left)');
      let glyph = face.glyph(code);
      // A space that does not break is the space of a font that has no glyph of its own for it.
      if (glyph === 0 && code === 0xa0) {
        glyph = face.glyph(0x20);
        code = 0x20;
      }
      if (glyph === 0) return this.refuse(`the font ${face.name} has no glyph for U+${code.toString(16).toUpperCase()}: the browser paints it with another font`);
      singles.push({ glyph, text: code < 0x10000 ? String.fromCharCode(code) : String.fromCharCode(0xd800 + ((code - 0x10000) >> 10), 0xdc00 + ((code - 0x10000) & 0x3ff)), end: i + 1 });
    }
    if (singles.length === 0) return;
    const clusters = this.shape(face, value, singles);
    if (!clusters) return;
    let natural = 0;
    for (let i = 0; i < clusters.length; i++) {
      // A glyph is known to be sound before it is written: what is found broken here leaves the page a picture.
      if (!glyphUsable(face, clusters[i].glyph)) return this.refuse(`glyph ${clusters[i].glyph} of the font ${face.name} is not sound: it cannot be put into the file`);
      natural += face.advance(clusters[i].glyph);
    }
    natural *= unit;
    const widthOf = (from: number, to: number): number => {
      let width: number;
      try {
        width = this.options.measure.width(state.font, value.slice(from, to));
      } catch {
        width = NaN;
      }
      return finite(width) && width >= 0 ? width : NaN;
    };
    const measured = widthOf(0, value.length);
    if (!finite(measured)) return this.refuse('text could not be measured');
    // What the canvas adds between the glyphs is kerning; more than that is another text.
    if (Math.abs(measured - natural) > TEXT_FIT_SHARE * size * clusters.length + 0.01) return this.refuse('a text is not as wide as its glyphs are: the browser shapes it (ligatures, or another font)');
    const width = measured + state.spacing * clusters.length;
    let shift = 0;
    if (state.baseline !== 'alphabetic') {
      const measure = this.options.measure;
      let down = NaN;
      try {
        down = measure.baseline ? measure.baseline(state.font, state.baseline) : NaN;
      } catch {
        down = NaN;
      }
      if (!finite(down)) return this.refuse(`text is placed by its ${state.baseline} line, which could not be measured`);
      shift = down;
    }
    const left = state.align === 'center' ? x - width / 2 : state.align === 'right' || state.align === 'end' ? x - width : x;
    if (!this.useFill()) return;
    const named = this.fontName(face);
    this.marked = true;
    const out = state.out;
    const m = state.m;
    // The run in pieces of as many glyphs as are placed one by one, cut after a space where there is one: every
    // piece starts where the canvas has it, and so does every glyph in it.
    for (let first = 0; first < clusters.length; ) {
      let end = Math.min(clusters.length, first + MAX_PLACED_GLYPHS);
      if (end < clusters.length) {
        for (let k = end; k > first + 1; k--) {
          if (clusters[k - 1].text === ' ') {
            end = k;
            break;
          }
        }
      }
      const from = first === 0 ? 0 : clusters[first - 1].end;
      // Where the piece starts: the text up to and with its first glyph ends where that glyph ends.
      let origin = 0;
      if (first > 0) {
        origin = widthOf(0, clusters[first].end) - face.advance(clusters[first].glyph) * unit;
        if (!finite(origin)) return this.refuse('text could not be measured');
      }
      // Where the canvas has each glyph of the piece: the piece up to and with it, measured, ends where the glyph
      // ends, so the glyph starts its own advance before that. What that differs by from where the glyph before it
      // ends (kerning) is moved by in the text that is written.
      let shown = '';
      let hex = '';
      let start = 0;
      for (let i = first; i < end; i++) {
        const cluster = clusters[i];
        hex += hex4(cluster.glyph);
        // A glyph that stands for several texts is copied as the plainest of them: see plainerText().
        const known = named.glyphs.get(cluster.glyph);
        named.glyphs.set(cluster.glyph, known === undefined ? cluster.text : plainerText(known, cluster.text));
        if (i + 1 >= end) break;
        const reach = widthOf(from, clusters[i + 1].end);
        if (!finite(reach)) return this.refuse('text could not be measured');
        const next = reach - face.advance(clusters[i + 1].glyph) * unit;
        const move = next - (start + face.advance(cluster.glyph) * unit);
        if (Math.abs(move) > GLYPH_FIT_SHARE * size) return this.refuse('a text is not as wide as its glyphs are: the browser shapes it (ligatures, or another font)');
        start = next;
        // A move to the right is a negative number between two strings, in thousandths of the font size.
        const moved = this.num((-move * 1000) / size, 2);
        if (moved !== '0') {
          shown += '<' + hex + '> ' + moved + ' ';
          hex = '';
        }
      }
      shown = shown === '' ? '<' + hex + '> Tj' : '[' + shown + '<' + hex + '>] TJ';
      let line = 'BT';
      const font = '/' + named.name + ' ' + this.num(size, 3) + ' Tf';
      if (out.font !== font) {
        line += ' ' + font;
        out.font = font;
      }
      const tc = this.num(state.spacing, 4);
      if (out.tc !== tc && !(out.tc === '' && tc === '0')) {
        line += ' ' + tc + ' Tc';
        out.tc = tc;
      }
      // Text space has its y axis upwards; the canvas has it downwards.
      const place = this.at(left + origin + state.spacing * first, y + shift);
      line += ' ' + this.num(m[0] * this.kx, 6) + ' ' + this.num(-m[1] * this.ky, 6) + ' ' + this.num(-m[2] * this.kx, 6) + ' ' + this.num(m[3] * this.ky, 6) + ' ' + this.xy(place) + ' Tm ' + shown + ' ET';
      this.emit(line);
      first = end;
    }
  }

  // --------------------------------------------------------------- pictures

  drawImage(source: unknown, ...numbers: number[]): void {
    if (this.unsupported !== undefined) return;
    for (let i = 0; i < numbers.length; i++) if (!finite(numbers[i])) return;
    const images = this.options.images;
    let picture: PdfPicture | undefined;
    try {
      picture = images ? images.read(source) : undefined;
    } catch {
      picture = undefined;
    }
    if (!picture || !(picture.width > 0) || !(picture.height > 0)) return this.refuse('a picture is drawn whose pixels cannot be read');
    let sx = 0;
    let sy = 0;
    let sw = picture.width;
    let sh = picture.height;
    let dx: number;
    let dy: number;
    let dw = sw;
    let dh = sh;
    if (numbers.length === 2) [dx, dy] = numbers;
    else if (numbers.length === 4) [dx, dy, dw, dh] = numbers;
    else if (numbers.length === 8) [sx, sy, sw, sh, dx, dy, dw, dh] = numbers;
    else return this.refuse('drawImage() is called in a form that is not read');
    if (sw === 0 || sh === 0 || dw === 0 || dh === 0) return;
    const slack = 0.01;
    if (sw < 0 || sh < 0 || sx < -slack || sy < -slack || sx + sw > picture.width + slack || sy + sh > picture.height + slack) return this.refuse('a picture is drawn from beyond its edges');
    const alpha = this.state.alpha;
    if (!(alpha > 0)) return;
    let index = this.pictures.indexOf(picture);
    if (index < 0) {
      this.pictures.push(picture);
      index = this.pictures.length - 1;
    }
    // Where the whole picture would lie when the part that is shown lies in the box.
    const fw = (picture.width * dw) / sw;
    const fh = (picture.height * dh) / sh;
    const fx = dx - (sx * dw) / sw;
    const fy = dy - (sy * dh) / sh;
    const cropped = sx > slack || sy > slack || sx + sw < picture.width - slack || sy + sh < picture.height - slack;
    const m = this.state.m;
    const corner = this.at(fx, fy + fh);
    this.marked = true;
    this.emit('q');
    const kept = { ...this.state.out };
    this.useAlpha(false, alpha);
    if (cropped) this.emit(this.box(dx, dy, dw, dh) + ' W n');
    // A picture is one unit wide and high, its first row at the top.
    this.emit(
      this.num(m[0] * fw * this.kx, 5) + ' ' + this.num(-m[1] * fw * this.ky, 5) + ' ' + this.num(-m[2] * fh * this.kx, 5) + ' ' + this.num(m[3] * fh * this.ky, 5) + ' ' + this.xy(corner) + ' cm /I' + index + ' Do',
    );
    this.emit('Q');
    this.state.out = kept;
  }

  // ----------------------------------------------------------------- result

  /** The page as it was painted. Throws when it could not be written: ask `unsupported` first. */
  finish(): PdfVectorPage {
    if (this.unsupported !== undefined) throw new PdfError('invalid', `The page cannot be written as drawings: ${this.unsupported}`);
    // A page nothing was painted on has nothing to say, whatever was saved, clipped and given back on it.
    if (!this.marked) return { content: '', fonts: [], pictures: [], alphas: [] };
    // Saves the painter did not give back.
    let content = this.parts.join('\n');
    for (let i = 0; i < this.stack.length; i++) content += '\nQ';
    return { content, fonts: this.fonts.slice(), pictures: this.pictures.slice(), alphas: this.alphas.slice() };
  }
}

/** The members of the canvas interface the PDF canvas answers for. */
const MEMBERS: { [name: string]: true } = {};
for (const member of [
  'save',
  'restore',
  'translate',
  'scale',
  'rotate',
  'transform',
  'setTransform',
  'resetTransform',
  'getTransform',
  'fillStyle',
  'strokeStyle',
  'createLinearGradient',
  'createRadialGradient',
  'createConicGradient',
  'createPattern',
  'globalAlpha',
  'lineWidth',
  'lineCap',
  'lineJoin',
  'miterLimit',
  'lineDashOffset',
  'setLineDash',
  'getLineDash',
  'globalCompositeOperation',
  'filter',
  'shadowBlur',
  'shadowOffsetX',
  'shadowOffsetY',
  'shadowColor',
  'imageSmoothingEnabled',
  'imageSmoothingQuality',
  'direction',
  'fontKerning',
  'wordSpacing',
  'beginPath',
  'moveTo',
  'lineTo',
  'closePath',
  'bezierCurveTo',
  'quadraticCurveTo',
  'rect',
  'arc',
  'fill',
  'stroke',
  'clip',
  'fillRect',
  'strokeRect',
  'font',
  'textAlign',
  'textBaseline',
  'letterSpacing',
  'measureText',
  'fillText',
  'drawImage',
])
  MEMBERS[member] = true;

/** The members of the canvas interface a PDF canvas has, for whoever lists what is drawn and what is not. */
export const PDF_CANVAS_MEMBERS: readonly string[] = Object.keys(MEMBERS);

/**
 * A canvas for a page of a PDF, as the context a painter draws on. Whatever
 * the painter asks of it that a PDF canvas does not have - reading or setting
 * a member it lacks, calling a method it lacks - makes the page unsupported
 * and does nothing. `canvas` holds the result: `unsupported`, and finish().
 */
export function createPdfCanvas(options: PdfCanvasOptions): { context: CanvasRenderingContext2D; canvas: PdfCanvas } {
  const canvas = new PdfCanvas(options);
  const target = canvas as unknown as { [name: string]: unknown };
  const bound = new Map<string, unknown>();
  /** What painters hang on a canvas for each other (names that start with two underscores). */
  const notes: { [name: string]: unknown } = {};
  const own = (key: string | symbol): boolean => typeof key === 'string' && Object.prototype.hasOwnProperty.call(MEMBERS, key);
  const context = new Proxy(canvas, {
    get(_target, key) {
      if (typeof key !== 'string') return undefined;
      if (own(key)) {
        const value = target[key];
        if (typeof value !== 'function') return value;
        let method = bound.get(key);
        if (!method) {
          method = (value as (...args: unknown[]) => unknown).bind(canvas);
          bound.set(key, method);
        }
        return method;
      }
      if (key.slice(0, 2) === '__') return notes[key];
      // What a promise or a printout of the object asks for is no drawing.
      if (key === 'then' || key === 'toJSON' || key === 'constructor' || key === 'toString' || key === 'valueOf') return undefined;
      canvas.refuse(`the painter uses "${key.slice(0, 40)}" of the canvas, which is not written`);
      return undefined;
    },
    set(_target, key, value) {
      if (typeof key !== 'string') return true;
      if (own(key)) {
        if (typeof target[key] === 'function') canvas.refuse(`the painter replaces "${key}" of the canvas`);
        else target[key] = value;
      } else if (key.slice(0, 2) === '__') notes[key] = value;
      else canvas.refuse(`the painter sets "${key.slice(0, 40)}" of the canvas, which is not written`);
      return true;
    },
    has(_target, key) {
      return own(key);
    },
    deleteProperty(_target, key) {
      if (typeof key === 'string' && key.slice(0, 2) === '__') delete notes[key];
      return true;
    },
  });
  return { context: context as unknown as CanvasRenderingContext2D, canvas };
}
