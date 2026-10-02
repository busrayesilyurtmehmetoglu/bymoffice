// Text measuring for the document layout engine, over a canvas.
//
// The same CSS font string is used for measuring and for painting, so that
// what was measured is what gets drawn. Font names come from files and are
// never trusted: they reach a font string only through sanitizeFontFamily().
//
// Common Office fonts are followed by fonts with the same metrics (Calibri ->
// Carlito, ...), so that a document breaks its lines the same way on a device
// that has only the substitute; the last entry is always a generic family.

import type { FontMetrics, LayoutFont, TextMeasurer } from '../layout/doc/types';
import { sanitizeFontFamily } from './text';

/** What the measurer needs of a canvas context. */
export interface MeasureContext {
  font: string;
  measureText(text: string): {
    width: number;
    fontBoundingBoxAscent?: number;
    fontBoundingBoxDescent?: number;
    actualBoundingBoxAscent?: number;
    actualBoundingBoxDescent?: number;
  };
}

/** What is used of the document's font table (DocModel.fonts). */
export interface FontHint {
  altName?: string;
  /** "roman", "swiss", "modern", "script", "decorative", "auto". */
  family?: string;
  /** "fixed", "variable", "default". */
  pitch?: string;
}

export interface CanvasMeasurerOptions {
  /** Context to measure with; a canvas is created on first use when missing. */
  context?: MeasureContext;
  /** The font table of the document, by font name: picks the generic family of fonts the device may not have. */
  fonts?: { get(name: string): FontHint | undefined };
  /** Tells whether the device can show a family; replaces the built-in test. */
  hasFamily?(family: string): boolean;
}

/** Fonts with the same advance widths, by lower-case family name. */
const SUBSTITUTES: { [family: string]: string[] } = {
  calibri: ['Carlito'],
  'calibri light': ['Carlito'],
  cambria: ['Caladea'],
  'cambria math': ['Caladea'],
  arial: ['Liberation Sans', 'Arimo', 'Helvetica'],
  helvetica: ['Arial', 'Liberation Sans', 'Arimo'],
  'arial narrow': ['Liberation Sans Narrow'],
  'times new roman': ['Liberation Serif', 'Tinos', 'Times'],
  times: ['Times New Roman', 'Liberation Serif', 'Tinos'],
  'courier new': ['Liberation Mono', 'Cousine', 'Courier'],
  courier: ['Courier New', 'Liberation Mono', 'Cousine'],
  georgia: ['Gelasio'],
  carlito: ['Calibri'],
  caladea: ['Cambria'],
};

const SERIF = [
  'times new roman',
  'times',
  'cambria',
  'cambria math',
  'georgia',
  'garamond',
  'book antiqua',
  'palatino linotype',
  'bookman old style',
  'century',
  'century schoolbook',
  'constantia',
  'caladea',
  'liberation serif',
  'tinos',
  'gelasio',
];
const MONOSPACE = ['courier new', 'courier', 'consolas', 'lucida console', 'liberation mono', 'cousine'];

const MIN_FONT_PX = 0.5;
const MAX_FONT_PX = 4000;
/** Longer strings are measured in pieces: a huge string is never handed to the canvas at once. */
const MAX_MEASURE_CHARS = 4096;
const WIDTH_CACHE_LIMIT = 20000;
const WIDTH_CACHE_MAX_TEXT = 64;
const FAMILY_CACHE_LIMIT = 2000;

/** Line extent of a font the canvas tells nothing about, as shares of its size. */
const FALLBACK_ASCENT = 0.91;
const FALLBACK_DESCENT = 0.24;

const PROBE_TEXT = 'mmmmmmmmmmlliWW@#01';
const PROBE_SIZE = '72px ';

const finitePositive = (value: unknown): value is number => typeof value === 'number' && value > 0 && value < Infinity;

function fontSize(size: number): number {
  if (!(size >= MIN_FONT_PX)) return MIN_FONT_PX;
  return size > MAX_FONT_PX ? MAX_FONT_PX : Math.round(size * 100) / 100;
}

/** Measures text for DocLayout on a canvas and names the CSS fonts to paint it with. */
export class CanvasTextMeasurer implements TextMeasurer {
  private context: MeasureContext | undefined;
  private readonly hints: { get(name: string): FontHint | undefined } | undefined;
  private readonly hasFamily: ((family: string) => boolean) | undefined;
  private current: string | undefined;
  private families = new Map<string, string>();
  private available = new Map<string, boolean>();
  private byFont = new WeakMap<LayoutFont, string>();
  private widths = new Map<string, number>();
  private extents = new Map<string, FontMetrics>();

  constructor(options: CanvasMeasurerOptions = {}) {
    this.context = options.context;
    this.hints = options.fonts;
    this.hasFamily = options.hasFamily;
    // Handed around on their own (to renderPage, say).
    this.cssFont = this.cssFont.bind(this);
    this.metrics = this.metrics.bind(this);
  }

  /**
   * The CSS family list for a family of the document: the family itself, the
   * alternative the document names, fonts with the same metrics and a generic
   * family. Every name is sanitized and quoted.
   */
  cssFamilies(family: string): string {
    let list = this.families.get(family);
    if (list !== undefined) return list;
    const names = this.chain(family);
    list = '';
    for (let i = 0; i < names.length; i++) list += '"' + names[i] + '", ';
    list += this.generic(family, names);
    if (this.families.size >= FAMILY_CACHE_LIMIT) this.families.clear();
    this.families.set(family, list);
    return list;
  }

  /** The value for `ctx.font`. */
  cssFont(font: LayoutFont): string {
    let css = this.byFont.get(font);
    if (css === undefined) {
      css = (font.italic === true ? 'italic ' : '') + (font.bold === true ? 'bold ' : '') + fontSize(font.size) + 'px ' + this.cssFamilies(font.family);
      this.byFont.set(font, css);
    }
    return css;
  }

  width(font: LayoutFont, text: string): number {
    if (text.length === 0) return 0;
    const css = this.cssFont(font);
    if (text.length > WIDTH_CACHE_MAX_TEXT) return this.measure(css, text);
    const key = css + '\n' + text;
    let width = this.widths.get(key);
    if (width === undefined) {
      width = this.measure(css, text);
      if (this.widths.size >= WIDTH_CACHE_LIMIT) this.widths.clear();
      this.widths.set(key, width);
    }
    return width;
  }

  metrics(font: LayoutFont): FontMetrics {
    const css = this.cssFont(font);
    let metrics = this.extents.get(css);
    if (metrics) return metrics;
    const size = fontSize(font.size);
    const measured = this.use(css).measureText('Hg');
    if (finitePositive(measured.fontBoundingBoxAscent) && typeof measured.fontBoundingBoxDescent === 'number' && measured.fontBoundingBoxDescent >= 0 && measured.fontBoundingBoxDescent < Infinity) {
      metrics = { ascent: measured.fontBoundingBoxAscent, descent: measured.fontBoundingBoxDescent, lineGap: 0 };
    } else {
      // The ink of two letters says little about the line; it only keeps the estimate from being too small.
      const inkUp = finitePositive(measured.actualBoundingBoxAscent) ? measured.actualBoundingBoxAscent : 0;
      const inkDown = finitePositive(measured.actualBoundingBoxDescent) ? measured.actualBoundingBoxDescent : 0;
      metrics = { ascent: Math.max(size * FALLBACK_ASCENT, inkUp), descent: Math.max(size * FALLBACK_DESCENT, inkDown), lineGap: 0 };
    }
    if (this.extents.size >= FAMILY_CACHE_LIMIT) this.extents.clear();
    this.extents.set(css, metrics);
    return metrics;
  }

  /** Whether the device can show the family, or a font with the same metrics. */
  hasFont(family: string): boolean {
    const names = this.chain(family);
    for (let i = 0; i < names.length; i++) if (this.isAvailable(names[i])) return true;
    return false;
  }

  /** Forgets everything measured: call when a web font finished loading, then lay out again. */
  clear(): void {
    this.widths.clear();
    this.extents.clear();
    this.available.clear();
    this.current = undefined;
  }

  /** The family, its alternative and its substitutes, sanitized, without repeats. */
  private chain(family: string): string[] {
    const names: string[] = [];
    const add = (name: unknown) => {
      const clean = sanitizeFontFamily(name);
      if (!clean) return;
      const lower = clean.toLowerCase();
      for (let i = 0; i < names.length; i++) if (names[i].toLowerCase() === lower) return;
      names.push(clean);
    };
    add(family);
    const hint = this.hints ? this.hints.get(family) : undefined;
    if (hint) add(hint.altName);
    const first = names.length ? names[0].toLowerCase() : '';
    const substitutes = Object.prototype.hasOwnProperty.call(SUBSTITUTES, first) ? SUBSTITUTES[first] : undefined;
    if (substitutes) for (let i = 0; i < substitutes.length; i++) add(substitutes[i]);
    return names;
  }

  private generic(family: string, names: string[]): string {
    const hint = this.hints ? this.hints.get(family) : undefined;
    for (let i = 0; i < names.length; i++) {
      const lower = names[i].toLowerCase();
      if (MONOSPACE.indexOf(lower) >= 0) return 'monospace';
      if (SERIF.indexOf(lower) >= 0) return 'serif';
    }
    if (hint) {
      if (hint.pitch === 'fixed' || hint.family === 'modern') return 'monospace';
      if (hint.family === 'roman') return 'serif';
      if (hint.family === 'script') return 'cursive';
    }
    return 'sans-serif';
  }

  private isAvailable(name: string): boolean {
    let known = this.available.get(name);
    if (known === undefined) {
      known = this.hasFamily ? this.hasFamily(name) === true : this.probe(name);
      if (this.available.size >= FAMILY_CACHE_LIMIT) this.available.clear();
      this.available.set(name, known);
    }
    return known;
  }

  /** `name` is sanitized. A family the device lacks measures like the generic family behind it, whichever that is. */
  private probe(name: string): boolean {
    const fonts = typeof document !== 'undefined' ? (document as { fonts?: { check?(font: string): boolean } }).fonts : undefined;
    if (fonts && typeof fonts.check === 'function') {
      try {
        // False is reliable (a web font that is not loaded, or no such font); true is not in every browser.
        if (!fonts.check('16px "' + name + '"')) return false;
      } catch {
        // The test below decides.
      }
    }
    const quoted = '"' + name + '", ';
    const generics = ['monospace', 'serif'];
    for (let i = 0; i < generics.length; i++) {
      const plain = this.use(PROBE_SIZE + generics[i]).measureText(PROBE_TEXT).width;
      const named = this.use(PROBE_SIZE + quoted + generics[i]).measureText(PROBE_TEXT).width;
      if (plain !== named) return true;
    }
    return false;
  }

  private use(css: string): MeasureContext {
    let context = this.context;
    if (!context) {
      const canvas = typeof document !== 'undefined' ? document.createElement('canvas') : undefined;
      const made = canvas ? canvas.getContext('2d') : null;
      if (!made) throw new Error('Text cannot be measured here: no 2D canvas context');
      context = this.context = made;
    }
    if (this.current !== css) {
      context.font = css;
      this.current = css;
    }
    return context;
  }

  private measure(css: string, text: string): number {
    const context = this.use(css);
    if (text.length <= MAX_MEASURE_CHARS) return context.measureText(text).width;
    let width = 0;
    for (let start = 0; start < text.length; ) {
      let end = Math.min(start + MAX_MEASURE_CHARS, text.length);
      // Not between the halves of a surrogate pair.
      if (end < text.length) {
        const code = text.charCodeAt(end - 1);
        if (code >= 0xd800 && code <= 0xdbff) end--;
      }
      width += context.measureText(text.slice(start, end)).width;
      start = end;
    }
    return width;
  }
}

/** A measurer over a canvas of its own. */
export function createCanvasMeasurer(options?: CanvasMeasurerOptions): CanvasTextMeasurer {
  return new CanvasTextMeasurer(options);
}

let shared: CanvasTextMeasurer | undefined;

/** The CSS font for a layout font when no measurer with the document's font table is at hand. */
export function docCssFont(font: LayoutFont): string {
  return (shared || (shared = new CanvasTextMeasurer())).cssFont(font);
}
