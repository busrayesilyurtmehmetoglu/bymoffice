// Text measuring for the document layout engine, over a canvas.
//
// The same CSS font string is used for measuring and for painting, so that
// what was measured is what gets drawn. Font names come from files and are
// never trusted: they reach a font string only through sanitizeFontFamily().
//
// Common Office fonts are followed by fonts with the same metrics (Calibri ->
// Carlito, ...), so that a document breaks its lines the same way on a device
// that has only the substitute; the last entry is always a generic family.
//
// A name with a code page suffix ("Bookman Old Style Tur", "Arial CE") is the
// font in front of the suffix: it is followed by that family, and everything
// that is looked up by family (substitutes, line metrics, the generic family)
// is looked up by it. Fonts the application registered for a family
// (registerFont() with `as`) come right after it. The name in the model is
// never changed.

import type { FontMetrics, LayoutFont, TextMeasurer } from '../layout/doc/types';
import { fontsVersion, replacementFonts, sanitizeFontFamily, splitFontFamily } from './fonts';

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
  // The families below are missing on Apple and Android devices. Their substitutes are clones of the PostScript
  // fonts the Monotype families were drawn to match (ITC Bookman, Palatino, New Century Schoolbook, ITC Avant Garde
  // Gothic). UNSURE: that the Monotype files have exactly the advance widths of those PostScript fonts was not
  // checked against the files (none was at hand).
  'bookman old style': ['TeX Gyre Bonum', 'URW Bookman', 'URW Bookman L', 'Bookman'],
  'book antiqua': ['TeX Gyre Pagella', 'Palatino Linotype', 'Palatino', 'URW Palladio L'],
  'century schoolbook': ['TeX Gyre Schola', 'Century Schoolbook L'],
  'century gothic': ['TeX Gyre Adventor', 'URW Gothic'],
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

/**
 * Vertical metrics of common fonts as Word uses them to set lines, in units
 * of 1/2048 em: [ascent, descent, leading between lines]. Ascent and descent
 * are usWinAscent and usWinDescent of the OS/2 table; the leading is
 * max(0, hhea.lineGap - ((usWinAscent + usWinDescent) - (hhea.ascender - hhea.descender))),
 * the "external leading" of GDI. Only numbers are kept here, no font data.
 *
 * A canvas reports the box of whatever font it ends up drawing with, some
 * browsers in whole pixels, and none reports the leading: lines would be
 * 17 px in one browser, 19 in another and 18.4 in Word. With the table the
 * lines of these fonts are as high as in Word, in every browser. The family
 * the document asks for decides, not the font the device shows it with: that
 * is how Word paginated the document.
 *
 * "verified": read from the font file. "UNSURE": from memory, no file of the
 * font was at hand to check; the sum (the line pitch) is the better known part.
 */
const UNITS_PER_EM = 2048;
const ARIAL: [number, number, number] = [1854, 434, 67]; // verified
const TIMES_NEW_ROMAN: [number, number, number] = [1825, 443, 87]; // verified
const COURIER_NEW: [number, number, number] = [1705, 615, 0]; // verified
const CALIBRI: [number, number, number] = [1950, 550, 0]; // UNSURE (line 1.2207 em)
const CAMBRIA: [number, number, number] = [1946, 455, 0]; // UNSURE (line 1.1724 em)
const VERTICAL_METRICS: { [family: string]: [number, number, number] } = {
  arial: ARIAL,
  // Metric-compatible with Arial by design; Helvetica is shown as Arial by Word on Windows.
  'liberation sans': ARIAL,
  arimo: ARIAL,
  helvetica: ARIAL,
  'times new roman': TIMES_NEW_ROMAN,
  'liberation serif': TIMES_NEW_ROMAN,
  tinos: TIMES_NEW_ROMAN,
  times: TIMES_NEW_ROMAN,
  'courier new': COURIER_NEW,
  'liberation mono': COURIER_NEW,
  cousine: COURIER_NEW,
  courier: COURIER_NEW,
  'arial narrow': [1888, 431, 31], // verified
  'arial black': [2254, 634, 0], // verified
  'arial unicode ms': [2189, 555, 0], // verified
  verdana: [2059, 430, 0], // verified
  tahoma: [2049, 423, 0], // verified
  georgia: [1878, 449, 0], // verified
  'trebuchet ms': [1923, 455, 0], // verified
  'comic sans ms': [2257, 597, 0], // verified
  impact: [2066, 432, 0], // verified
  'microsoft sans serif': [1888, 430, 0], // verified
  'andale mono': [1858, 446, 0], // verified
  'arial rounded mt bold': [1938, 432, 0], // verified (macOS file)
  'brush script mt': [1820, 692, 0], // verified (macOS file)
  rockwell: [2099, 600, 0], // verified (macOS file)
  webdings: [1638, 410, 0], // verified
  wingdings: [1841, 432, 0], // verified
  'wingdings 2': [1727, 432, 0], // verified
  'wingdings 3': [1900, 432, 0], // verified
  calibri: CALIBRI,
  'calibri light': CALIBRI, // UNSURE
  carlito: CALIBRI, // UNSURE: made to match Calibri
  cambria: CAMBRIA,
  caladea: CAMBRIA, // UNSURE: made to match Cambria
  'segoe ui': [2210, 514, 0], // UNSURE (line 1.33 em)
  consolas: [1884, 514, 0], // UNSURE (line 1.1709 em)
  garamond: [1765, 536, 0], // UNSURE (line 1.1235 em)
  'palatino linotype': [1499, 582, 682], // UNSURE (line 1.349 em)
  'century gothic': [2060, 451, 0], // UNSURE (line 1.226 em)
  symbol: [2059, 450, 0], // UNSURE (line 1.225 em)
};

/** The row of the table for a family of a document: by its name, then by the name without its code page suffix. */
function tableMetrics(family: unknown): [number, number, number] | undefined {
  if (typeof family !== 'string') return undefined;
  const name = family.trim().toLowerCase();
  if (Object.prototype.hasOwnProperty.call(VERTICAL_METRICS, name)) return VERTICAL_METRICS[name];
  const base = splitFontFamily(family).base.toLowerCase();
  return Object.prototype.hasOwnProperty.call(VERTICAL_METRICS, base) ? VERTICAL_METRICS[base] : undefined;
}

/** Size at which the font box of a family that is not in the table is measured; the result is scaled to the size asked for. */
const METRICS_PROBE_PX = 1000;

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
  /** Font box per px of size, by face (family list and style), of the families that are not in the table. */
  private extents = new Map<string, { ascent: number; descent: number } | null>();
  /** fontsVersion() the family lists were built for. */
  private registered = fontsVersion();

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
    this.sync();
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
    this.sync();
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
    const size = fontSize(font.size);
    const units = tableMetrics(font.family);
    if (units) {
      // The same in every browser, and what Word used when it broke the document into pages.
      return { ascent: (units[0] * size) / UNITS_PER_EM, descent: (units[1] * size) / UNITS_PER_EM, lineGap: (units[2] * size) / UNITS_PER_EM };
    }
    // Measured once per face at a large size: a canvas that reports whole pixels is off by up to 4% at text sizes.
    const face = (font.italic === true ? 'italic ' : '') + (font.bold === true ? 'bold ' : '') + METRICS_PROBE_PX + 'px ' + this.cssFamilies(font.family);
    let box = this.extents.get(face);
    if (box === undefined) {
      const measured = this.use(face).measureText('Hg');
      const up = measured.fontBoundingBoxAscent;
      const down = measured.fontBoundingBoxDescent;
      box = finitePositive(up) && typeof down === 'number' && down >= 0 && down < Infinity ? { ascent: up / METRICS_PROBE_PX, descent: down / METRICS_PROBE_PX } : null;
      if (this.extents.size >= FAMILY_CACHE_LIMIT) this.extents.clear();
      this.extents.set(face, box);
    }
    if (box) return { ascent: box.ascent * size, descent: box.descent * size, lineGap: 0 };
    // No font box on this canvas. The ink of two letters says little about the line; it only keeps the estimate from being too small.
    const measured = this.use(this.cssFont(font)).measureText('Hg');
    const inkUp = finitePositive(measured.actualBoundingBoxAscent) ? measured.actualBoundingBoxAscent : 0;
    const inkDown = finitePositive(measured.actualBoundingBoxDescent) ? measured.actualBoundingBoxDescent : 0;
    return { ascent: Math.max(size * FALLBACK_ASCENT, inkUp), descent: Math.max(size * FALLBACK_DESCENT, inkDown), lineGap: 0 };
  }

  /** Whether the device can show the family, or a font with the same metrics. */
  hasFont(family: string): boolean {
    this.sync();
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

  /** A font was registered or removed since the family lists were built: they and everything measured with them are forgotten. */
  private sync(): void {
    const now = fontsVersion();
    if (now === this.registered) return;
    this.registered = now;
    this.families.clear();
    this.byFont = new WeakMap();
    this.clear();
  }

  /**
   * The names to try for a family of the document, sanitized, without
   * repeats: registered fonts preferred to it, the name itself, the family
   * without its code page suffix, the other registered fonts that stand in
   * for it, the alternative of the font table and the substitutes.
   */
  private chain(family: string): string[] {
    const names: string[] = [];
    const add = (name: unknown) => {
      const clean = sanitizeFontFamily(name);
      if (!clean) return;
      const lower = clean.toLowerCase();
      for (let i = 0; i < names.length; i++) if (names[i].toLowerCase() === lower) return;
      names.push(clean);
    };
    const { full, base } = splitFontFamily(family);
    const registered = full === '' ? undefined : replacementFonts(full, base);
    if (registered) for (let i = 0; i < registered.first.length; i++) add(registered.first[i]);
    add(full);
    add(base);
    if (registered) for (let i = 0; i < registered.then.length; i++) add(registered.then[i]);
    const hint = this.hints ? this.hints.get(family) : undefined;
    if (hint) add(hint.altName);
    const lower = base.toLowerCase();
    const substitutes = Object.prototype.hasOwnProperty.call(SUBSTITUTES, lower) ? SUBSTITUTES[lower] : undefined;
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
