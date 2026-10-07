// Text helpers for Canvas drawing: CSS font strings, cached measuring and
// line breaking.
//
// Font names come from files and are never trusted: a name only reaches
// `ctx.font` after everything but letters, digits, spaces and hyphens has been
// removed from it.

import type { Font } from '../model/sheet/types';
import { fontFamilyChain, fontsVersion, sanitizeFontFamily } from './fonts';

export { sanitizeFontFamily };

/** Fonts tried after the one a cell asks for. Carlito is metric-compatible with Calibri. */
export const FONT_FALLBACK = '"Carlito", "Calibri", Arial, sans-serif';

const DEFAULT_FONT_POINTS = 11;
/** Excel's largest font size, in points. */
const MAX_FONT_POINTS = 409;
/** Font size in CSS pixels at 96 dpi; sizes that are missing or not a positive number read as 11 pt. */
export function fontPixels(font: Pick<Font, 'size'> | undefined): number {
  const size = font ? font.size : undefined;
  const points = typeof size === 'number' && size > 0 && Number.isFinite(size) ? Math.min(Math.max(size, 1), MAX_FONT_POINTS) : DEFAULT_FONT_POINTS;
  return Math.round(((points * 96) / 72) * 100) / 100;
}

/**
 * The value for `ctx.font`: style, weight, size and the quoted family followed by the fallback stack. A name with a
 * code page suffix ("Arial Tur") is followed by its family, and a registered font that stands in for the family by that font.
 */
export function cssFont(font: Font | undefined): string {
  const names = fontFamilyChain(font ? font.name : undefined);
  let families = '';
  for (let i = 0; i < names.length; i++) families += '"' + names[i] + '", ';
  return (font && font.italic === true ? 'italic ' : '') + (font && font.bold === true ? 'bold ' : '') + fontPixels(font) + 'px ' + families + FONT_FALLBACK;
}

const MEASURE_CACHE_LIMIT = 5000;
/** Longer strings are measured every time: they are rare and would make the cache keys large. */
const MEASURE_CACHE_MAX_TEXT = 200;
const measureCache = new Map<string, number>();
/** fontsVersion() the cached widths were measured under: a font that was registered or taken out since makes the same font string measure differently. */
let measuredUnder = fontsVersion();

/**
 * Empties the width cache. Widths are cached by font string and text, so call
 * this when a web font finishes loading and the same font string starts to
 * measure differently.
 */
export function clearTextMeasureCache(): void {
  measureCache.clear();
}

/**
 * Tracks the font and fill of a context so that they are assigned only when
 * they change, and measures text through a bounded cache.
 */
export class TextPainter {
  readonly ctx: CanvasRenderingContext2D;
  private font: string | undefined;
  private fill: string | undefined;

  constructor(ctx: CanvasRenderingContext2D) {
    this.ctx = ctx;
  }

  setFont(font: string): void {
    if (this.font === font) return;
    this.ctx.font = font;
    this.font = font;
  }

  setFill(fill: string): void {
    if (this.fill === fill) return;
    this.ctx.fillStyle = fill;
    this.fill = fill;
  }

  /** Call after `ctx.restore()` or after assigning `font`/`fillStyle` directly: the tracked values are no longer known. */
  invalidate(): void {
    this.font = undefined;
    this.fill = undefined;
  }

  /** Width of `text` in `font`, in CSS pixels. */
  measure(font: string, text: string): number {
    if (text.length === 0) return 0;
    if (text.length > MEASURE_CACHE_MAX_TEXT) {
      this.setFont(font);
      return this.ctx.measureText(text).width;
    }
    const now = fontsVersion();
    if (now !== measuredUnder) {
      measuredUnder = now;
      measureCache.clear();
    }
    const key = font + '\n' + text;
    let width = measureCache.get(key);
    if (width === undefined) {
      this.setFont(font);
      width = this.ctx.measureText(text).width;
      if (measureCache.size >= MEASURE_CACHE_LIMIT) measureCache.clear();
      measureCache.set(key, width);
    }
    return width;
  }
}

/** Words longer than this are measured character by character when wrapping. */
const LONG_WORD = 256;

/** Length in UTF-16 units of the character that starts at `index` (2 for a surrogate pair). */
function charLength(text: string, index: number, end: number): number {
  const code = text.charCodeAt(index);
  if (code >= 0xd800 && code <= 0xdbff && index + 1 < end) {
    const next = text.charCodeAt(index + 1);
    if (next >= 0xdc00 && next <= 0xdfff) return 2;
  }
  return 1;
}

/**
 * Breaks `text` into lines no wider than `maxWidth`. "\n" (also "\r\n") always
 * ends a line; otherwise lines break at spaces, and a word wider than a whole
 * line is broken between characters. Every line holds at least one character,
 * so the result is bounded by the text length whatever the width is.
 *
 * Returns a flat list of [start, end) offsets into `text`, two numbers per
 * line. Widths are summed from words and spaces measured on their own, which
 * ignores kerning across a space.
 */
export function wrapLines(text: string, maxWidth: number, measure: (text: string) => number): number[] {
  const lines: number[] = [];
  const limit = maxWidth > 0 ? maxWidth : 0;
  const length = text.length;
  let spaceWidth = -1;

  for (let paragraph = 0; paragraph <= length; ) {
    let next = text.indexOf('\n', paragraph);
    if (next < 0) next = length;
    const end = next > paragraph && text.charCodeAt(next - 1) === 13 ? next - 1 : next;

    // The current line is [lineStart, lineEnd); it is empty while lineEnd === lineStart.
    let lineStart = paragraph;
    let lineEnd = paragraph;
    let width = 0;
    let i = paragraph;
    while (i < end) {
      let wordStart = i;
      while (wordStart < end && text.charCodeAt(wordStart) === 32) wordStart++;
      let wordEnd = wordStart;
      while (wordEnd < end && text.charCodeAt(wordEnd) !== 32) wordEnd++;
      // Spaces at the end of a paragraph are dropped.
      if (wordEnd === wordStart) break;

      if (spaceWidth < 0) spaceWidth = measure(' ');
      let wordWidth = 0;
      if (wordEnd - wordStart <= LONG_WORD) {
        wordWidth = measure(text.slice(wordStart, wordEnd));
      } else {
        // Never hand a huge string to the canvas in one piece: add it up character by character.
        for (let pos = wordStart; pos < wordEnd; ) {
          const size = charLength(text, pos, wordEnd);
          wordWidth += measure(size === 1 ? text.charAt(pos) : text.slice(pos, pos + size));
          pos += size;
        }
      }
      let gap = (wordStart - i) * spaceWidth;

      if (lineEnd > lineStart && width + gap + wordWidth > limit) {
        // The word goes to a new line; the spaces before it are dropped.
        lines.push(lineStart, lineEnd);
        lineStart = wordStart;
        lineEnd = wordStart;
        width = 0;
        gap = 0;
      }
      if (lineEnd === lineStart && gap + wordWidth > limit) {
        // Too wide even alone: break between characters. Spaces that start a paragraph stay with it.
        let pos = lineStart;
        width = 0;
        while (pos < wordEnd) {
          const size = charLength(text, pos, wordEnd);
          const charWidth = text.charCodeAt(pos) === 32 ? spaceWidth : measure(size === 1 ? text.charAt(pos) : text.slice(pos, pos + size));
          if (pos > lineStart && width + charWidth > limit) {
            lines.push(lineStart, pos);
            lineStart = pos;
            width = 0;
          }
          width += charWidth;
          pos += size;
        }
      } else {
        width += gap + wordWidth;
      }
      lineEnd = wordEnd;
      i = wordEnd;
    }
    lines.push(lineStart, lineEnd);
    paragraph = next + 1;
  }
  return lines;
}
