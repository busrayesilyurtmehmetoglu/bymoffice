// The invisible text of a PDF page. The page itself is a picture; over it
// lies its text in text rendering mode 3 (neither filled nor stroked), every
// run at the place and with the width it has in the picture, so that a
// reader of the file can find, select and copy it.
//
// Where the runs come from: the text items of a laid out page of a document
// (docPageRuns), or whatever a painter writes on a canvas with fillText()
// (captureText) - cells, headings, headers and footers, labels of charts.
//
// What is in the text of a page is what can be seen in its picture: a run
// that a clip cuts off (the next cell of a row, the margin of the page, the
// frame of a chart, the edge of the paper) is in the text only as far as it
// shows, and one that is hidden altogether is not in it. Text that is painted
// and cannot be told from its background (white on white) is in the text: it
// is in the picture too.
//
// Runs are written in reading order: lines from the top of the page down,
// within a line from the left; the lines of a block (a cell) one after the
// other. The characters of a run are written in the order they are seen from
// left to right, which is what readers of PDF files expect: a stretch of
// Hebrew or Arabic is turned around (see visualOrder), and the reader turns
// it back. UNSURE: this is not the bidirectional algorithm. In a line whose
// base direction is right to left the stretches come out in the order of a
// line from left to right, and the runs of such a line are still ordered from
// the left; every word is found, the order of the words may differ.

import type { LayoutFont, LayoutPage } from '../layout/doc/types';
import type { PdfTextFont } from './font';
import { pdfNumber } from './writer';

/** Text on one baseline, as wide as it was painted. */
export interface PdfTextRun {
  text: string;
  /** Start of the baseline, from the top left corner of the page, in px of the page. */
  x: number;
  y: number;
  /** Advance along the baseline. */
  width: number;
  /** Size of the font. */
  size: number;
  /** How far the baseline is turned, in degrees clockwise (90: the text reads downwards). */
  angle?: number;
  /** Runs with the same number are the lines of one block of text (a cell): they are read one after the other. */
  block?: number;
  /**
   * The text goes on, unseen, before the start or after the end of the run (a clip cut it off there): whatever
   * stands next to it on that side is other text, and is kept apart from it by a space.
   */
  cutBefore?: boolean;
  cutAfter?: boolean;
}

/** Characters (UTF-16 units) of one run; what comes after them is left out, and the width is cut with them. */
export const MAX_RUN_CHARS = 2000;
/** Runs of one page; a page with more keeps the first ones. */
export const MAX_PAGE_RUNS = 200000;

const finite = (value: unknown): value is number => typeof value === 'number' && isFinite(value);

/** No page is larger than this in the unit of its runs, and no text on it larger. A run beyond it is no text of the page. */
const MAX_RUN_PLACE = 1e6;
const MAX_RUN_SIZE = 1e5;

function usable(run: PdfTextRun | undefined | null): run is PdfTextRun {
  if (!run || typeof run.text !== 'string' || run.text.length === 0) return false;
  if (!finite(run.x) || !finite(run.y) || Math.abs(run.x) > MAX_RUN_PLACE || Math.abs(run.y) > MAX_RUN_PLACE) return false;
  if (!finite(run.size) || !(run.size > 0) || run.size > MAX_RUN_SIZE) return false;
  // A width that is no number is a run without a width; one that is beyond every page is a run that cannot be.
  return !(typeof run.width === 'number' && Math.abs(run.width) > MAX_RUN_PLACE);
}

const turned = (run: PdfTextRun): boolean => finite(run.angle) && run.angle % 360 !== 0;

/**
 * Things in rows: rows from the top down, by where each thing has its
 * baseline. The baseline of a row is that of its largest text; smaller text
 * may sit above or below it (a footnote mark, a subscript). A turned thing is
 * a row of its own. Within a row things are ordered from the left.
 */
function rowsOf<T>(things: readonly T[], of: (thing: T) => PdfTextRun): T[][] {
  const kept = things.map((thing, index) => ({ thing, run: of(thing), index }));
  kept.sort((a, b) => a.run.y - b.run.y || a.index - b.index);
  const rows: (typeof kept)[] = [];
  let row: typeof kept | undefined;
  let baseline = 0;
  let size = 0;
  for (let i = 0; i < kept.length; i++) {
    const entry = kept[i];
    const run = entry.run;
    if (turned(run)) {
      rows.push([entry]);
      row = undefined;
      continue;
    }
    if (row && Math.abs(run.y - baseline) <= Math.max(size, run.size) * 0.5) {
      row.push(entry);
      if (run.size > size) {
        size = run.size;
        baseline = run.y;
      }
      continue;
    }
    row = [entry];
    rows.push(row);
    baseline = run.y;
    size = run.size;
  }
  const out: T[][] = [];
  for (let i = 0; i < rows.length; i++) {
    rows[i].sort((a, b) => a.run.x - b.run.x || a.index - b.index);
    out.push(rows[i].map((entry) => entry.thing));
  }
  return out;
}

/**
 * The runs in the lines they are written in. Runs of one block (the lines of
 * a cell) are kept together: the blocks of a row are read one after the
 * other, each from its first line to its last. Runs without a block, and
 * blocks of one line, are read line by line across the page.
 */
function linesOf(runs: readonly PdfTextRun[]): PdfTextRun[][] {
  const kept: PdfTextRun[] = [];
  for (let i = 0; i < runs.length && kept.length < MAX_PAGE_RUNS; i++) if (usable(runs[i])) kept.push(runs[i]);
  const own = (run: PdfTextRun) => run;
  let blocks: Map<number, PdfTextRun[]> | undefined;
  for (let i = 0; i < kept.length; i++) {
    const block = kept[i].block;
    if (typeof block !== 'number') continue;
    if (!blocks) blocks = new Map();
    const of = blocks.get(block);
    if (of) of.push(kept[i]);
    else blocks.set(block, [kept[i]]);
  }
  if (!blocks) return rowsOf(kept, own);
  // What is placed on the page: a run of its own, or a block with its lines.
  const units: PdfTextRun[][][] = [];
  const placed = new Set<number>();
  for (let i = 0; i < kept.length; i++) {
    const block = kept[i].block;
    if (typeof block !== 'number') units.push([[kept[i]]]);
    else if (!placed.has(block)) {
      placed.add(block);
      units.push(rowsOf(blocks.get(block) as PdfTextRun[], own));
    }
  }
  const out: PdfTextRun[][] = [];
  const rows = rowsOf(units, (unit) => unit[0][0]);
  for (let i = 0; i < rows.length; i++) {
    let line: PdfTextRun[] = [];
    for (let k = 0; k < rows[i].length; k++) {
      const unit = rows[i][k];
      if (unit.length === 1) {
        for (let n = 0; n < unit[0].length; n++) line.push(unit[0][n]);
        continue;
      }
      if (line.length > 0) out.push(line);
      line = [];
      for (let n = 0; n < unit.length; n++) out.push(unit[n]);
    }
    if (line.length > 0) out.push(line);
  }
  return out;
}

/** The runs that can be written, in the order they are written in. */
export function orderRuns(runs: readonly PdfTextRun[]): PdfTextRun[] {
  const out: PdfTextRun[] = [];
  const lines = linesOf(runs);
  for (let i = 0; i < lines.length; i++) for (let k = 0; k < lines[i].length; k++) out.push(lines[i][k]);
  return out;
}

/** The first characters of a text that is too long for a run; a pair of surrogates is not cut in two. */
function cut(text: string): string {
  if (text.length <= MAX_RUN_CHARS) return text;
  const last = text.charCodeAt(MAX_RUN_CHARS - 1);
  return text.slice(0, last >= 0xd800 && last <= 0xdbff ? MAX_RUN_CHARS - 1 : MAX_RUN_CHARS);
}

const isDigit = (code: number): boolean => (code >= 0x30 && code <= 0x39) || (code >= 0x660 && code <= 0x669) || (code >= 0x6f0 && code <= 0x6f9);

/** Whether a character is one of a script that is written from right to left (Hebrew, Arabic, Syriac, Thaana, N'Ko and their presentation forms). */
function isRightToLeft(code: number): boolean {
  if (isDigit(code)) return false;
  return (code >= 0x590 && code <= 0x8ff) || (code >= 0xfb1d && code <= 0xfdff) || (code >= 0xfe70 && code <= 0xfefc) || (code >= 0x10800 && code <= 0x10fff) || (code >= 0x1e800 && code <= 0x1efff);
}

/** Whether a character is a letter of a script that is written from left to right: roughly, a letter that is not one of the above. */
function isLeftToRight(code: number): boolean {
  if ((code >= 0x41 && code <= 0x5a) || (code >= 0x61 && code <= 0x7a)) return true;
  if (code < 0xc0 || code === 0xd7 || code === 0xf7) return false;
  // Punctuation, symbols and the marks of the directions themselves have no direction of their own.
  if ((code >= 0x2000 && code <= 0x2bff) || (code >= 0x3000 && code <= 0x303f) || (code >= 0xfe00 && code <= 0xfe6f) || (code >= 0xfff0 && code <= 0xffff)) return false;
  return !isRightToLeft(code);
}

/** The code points of a text; half a surrogate pair is a code point of its own. */
function codePoints(text: string): number[] {
  const points: number[] = [];
  for (let i = 0; i < text.length; i++) {
    let code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff && i + 1 < text.length) {
      const next = text.charCodeAt(i + 1);
      if (next >= 0xdc00 && next <= 0xdfff) {
        code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
        i++;
      }
    }
    points.push(code);
  }
  return points;
}

/** Turns the code points from `from` to `to` (both included) around; numbers in them stay as they are read ("2024", "3.5", "12:30"). */
function turnAround(points: number[], from: number, to: number): void {
  for (let a = from, b = to; a < b; a++, b--) {
    const swap = points[a];
    points[a] = points[b];
    points[b] = swap;
  }
  for (let i = from; i <= to; i++) {
    if (!isDigit(points[i])) continue;
    let end = i;
    for (let k = i + 1; k <= to; k++) {
      if (isDigit(points[k])) end = k;
      else if (!(k === end + 1 && (points[k] === 0x2e || points[k] === 0x2c || points[k] === 0x3a || points[k] === 0x2f || points[k] === 0x2d))) break;
    }
    for (let a = i, b = end; a < b; a++, b--) {
      const swap = points[a];
      points[a] = points[b];
      points[b] = swap;
    }
    i = end;
  }
}

/** Whether a text has letters of both directions. */
function mixedDirections(text: string): boolean {
  const points = codePoints(text);
  let left = false;
  let right = false;
  for (let i = 0; i < points.length; i++) {
    if (isLeftToRight(points[i])) left = true;
    else if (isRightToLeft(points[i])) right = true;
  }
  return left && right;
}

const hasRightToLeft = (text: string): boolean => codePoints(text).some(isRightToLeft);

/**
 * The characters of a run in the order they are seen from left to right. A
 * run with letters from right to left and none from left to right is turned
 * around as a whole, its numbers staying as they are read. In a run with
 * letters of both directions every stretch from right to left is turned on
 * its own: from its first such letter to its last, with what stands between
 * them, and with a number that follows it. This is not the bidirectional
 * algorithm: it is what makes each word, and each phrase of one direction,
 * come back as it was written.
 */
export function visualOrder(text: string): string {
  const points = codePoints(text);
  let left = false;
  let right = false;
  for (let i = 0; i < points.length; i++) {
    if (isLeftToRight(points[i])) left = true;
    else if (isRightToLeft(points[i])) right = true;
  }
  if (!right) return text;
  if (!left) turnAround(points, 0, points.length - 1);
  else {
    for (let i = 0; i < points.length; i++) {
      if (!isRightToLeft(points[i])) continue;
      // The stretch goes on to the last letter from right to left, or the last digit, before a letter from left to right.
      let end = i;
      for (let k = i + 1; k < points.length && !isLeftToRight(points[k]); k++) if (isRightToLeft(points[k]) || isDigit(points[k])) end = k;
      turnAround(points, i, end);
      i = end;
    }
  }
  let out = '';
  for (let i = 0; i < points.length; i++) out += points[i] < 0x10000 ? String.fromCharCode(points[i]) : String.fromCharCode(0xd800 + ((points[i] - 0x10000) >> 10), 0xdc00 + ((points[i] - 0x10000) & 0x3ff));
  return out;
}

/** A box on the page, in the unit of the runs. */
export interface ClipBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * How far the letters of a line reach above and below its baseline, in font
 * sizes, for telling whether a line shows in a clip. Rather too little than
 * too much: a line of which only the tip of an accent shows is left out, and
 * no line that does not show is taken.
 */
const SEEN_ABOVE = 0.7;
const SEEN_BELOW = 0.2;

/**
 * The part of a run that shows in a box. `advance(count)` says how wide the
 * first `count` UTF-16 units of the text are, in the unit of the run (the
 * whole text: the width of the run). A character that shows in part is in
 * the result; one that does not show is not, and the result starts where its
 * first character does. Undefined when nothing shows. Without `advance` a
 * run that is cut is left out altogether: where its characters are is not
 * guessed. So is a cut run with letters of both directions, whose characters
 * do not stand in the order of the text.
 */
export function clipRun(run: PdfTextRun, box: ClipBox, advance?: (count: number) => number): PdfTextRun | undefined {
  if (!(box.right > box.left) || !(box.bottom > box.top)) return undefined;
  const width = finite(run.width) && run.width > 0 ? run.width : 0;
  const angle = finite(run.angle) ? run.angle % 360 : 0;
  const radians = (angle * Math.PI) / 180;
  const cos = angle === 0 ? 1 : Math.cos(radians);
  const sin = angle === 0 ? 0 : Math.sin(radians);
  // The letters stand on the baseline: from it towards "up" (which a turn turns with the text) and a little the other way.
  // A point of the baseline shows when the box, widened by that much against the direction, holds it.
  const upX = sin * run.size;
  const upY = -cos * run.size;
  const left = box.left - Math.max(upX * SEEN_ABOVE, -upX * SEEN_BELOW);
  const right = box.right - Math.min(upX * SEEN_ABOVE, -upX * SEEN_BELOW);
  const top = box.top - Math.max(upY * SEEN_ABOVE, -upY * SEEN_BELOW);
  const bottom = box.bottom - Math.min(upY * SEEN_ABOVE, -upY * SEEN_BELOW);
  // The stretch of the baseline inside that box: from `from` to `to` along it.
  let from = 0;
  let to = width;
  const within = (start: number, step: number, low: number, high: number): boolean => {
    if (step === 0) return start > low && start < high;
    const a = (low - start) / step;
    const b = (high - start) / step;
    from = Math.max(from, Math.min(a, b));
    to = Math.min(to, Math.max(a, b));
    return true;
  };
  if (!within(run.x, cos, left, right) || !within(run.y, sin, top, bottom)) return undefined;
  const slack = Math.max(width, 1) * 1e-9;
  if (width === 0) return from <= slack && to >= -slack ? run : undefined;
  if (!(to - from > slack)) return undefined;
  if (from <= slack && to >= width - slack) return run;
  if (!advance || mixedDirections(run.text)) return undefined;
  const text = run.text;
  const length = text.length;
  // Text from right to left starts at the right end of its run.
  const mirrored = hasRightToLeft(text);
  const low = mirrored ? width - to : from;
  const high = mirrored ? width - from : to;
  const at = (count: number): number => {
    const value = count <= 0 ? 0 : count >= length ? width : advance(count);
    return finite(value) ? value : width;
  };
  // The first character that ends after `low`, and the first that starts at or after `high`.
  let first = 0;
  let a = 0;
  let b = length;
  while (a < b) {
    const mid = (a + b) >> 1;
    if (at(mid + 1) > low + slack) b = mid;
    else a = mid + 1;
  }
  first = a;
  a = first;
  b = length;
  while (a < b) {
    const mid = (a + b) >> 1;
    if (at(mid) < high - slack) a = mid + 1;
    else b = mid;
  }
  let end = a;
  // Half a pair of surrogates is no character: a pair that is cut is left out.
  const low16 = (index: number): boolean => index > 0 && index < length && (text.charCodeAt(index) & 0xfc00) === 0xdc00 && (text.charCodeAt(index - 1) & 0xfc00) === 0xd800;
  if (low16(first)) first++;
  if (low16(end)) end--;
  if (end <= first) return undefined;
  // The run is what shows of these characters: one that shows in part is in the text, and the run ends where the box does.
  const start = Math.max(at(first), low);
  const stop = Math.min(at(end), high);
  if (!(stop > start)) return undefined;
  const offset = mirrored ? width - stop : start;
  const out: PdfTextRun = { text: text.slice(first, end), x: run.x + offset * cos, y: run.y + offset * sin, width: stop - start, size: run.size };
  if (run.angle !== undefined) out.angle = run.angle;
  if (run.block !== undefined) out.block = run.block;
  // Along the baseline: text that is left out before the first character seen, or after the last.
  const before = mirrored ? end < length : first > 0;
  const after = mirrored ? first > 0 : end < length;
  if (before || run.cutBefore) out.cutBefore = true;
  if (after || run.cutAfter) out.cutAfter = true;
  return out;
}

/** Room between two runs of a line, in font sizes, from which on they are two words. */
const WORD_GAP = 0.15;

export interface TextLayerOptions {
  /** Height of the page in the unit of the runs. */
  height: number;
  /** Points per unit of the runs. */
  scale: number;
  font: PdfTextFont;
  /** The name of the font in the resources of the page; "F1" when missing. */
  fontName?: string;
}

/**
 * The operators that write the runs over a page: one text object, rendering
 * mode 3. Every run has a text matrix of its own (its place, and its turn),
 * and a horizontal scaling that makes its characters, each one em wide in
 * the glyphless font, take the width the run was measured with. Runs that
 * stand apart get a space between them, and a line ends with one, so that
 * what is copied does not join the words of two cells or the last word of a
 * line to the first of the next. Empty when nothing is written.
 */
export function textLayer(runs: readonly PdfTextRun[], options: TextLayerOptions): string {
  const { height, scale, font } = options;
  if (!finite(height) || !finite(scale) || !(scale > 0)) return '';
  const fontName = options.fontName || 'F1';
  const out: string[] = [];
  let currentSize = '';
  let currentScaling = '';
  const write = (text: string, x: number, y: number, width: number, size: number, angle: number): boolean => {
    // A run that was let through with a place no number of a PDF can state is no text of the page.
    if (!finite(x * scale) || !finite((height - y) * scale) || Math.abs(x * scale) >= 1e9 || Math.abs((height - y) * scale) >= 1e9 || !finite(size * scale) || size * scale >= 1e9) return false;
    const points = pdfNumber(size * scale, 2);
    if (points === '0') return false;
    const encoded = font.encode(visualOrder(text));
    if (encoded.count === 0) return false;
    // One em per character, scaled to what the run takes. A run that has no width that can be gets half an em per character.
    const natural = encoded.count * Number(points);
    const wanted = finite(width) && width > 0 ? (width * scale * 100) / natural : 50;
    const scaling = pdfNumber(Math.min(10000, Math.max(1, wanted)), 3);
    let line = '';
    if (points !== currentSize) {
      line += '/' + fontName + ' ' + points + ' Tf ';
      currentSize = points;
    }
    if (scaling !== currentScaling) {
      line += scaling + ' Tz ';
      currentScaling = scaling;
    }
    // The page of a PDF has its origin at the bottom left and its y axis pointing up: a clockwise turn there is a turn the other way.
    const radians = (angle * Math.PI) / 180;
    const cos = angle === 0 ? 1 : Math.cos(radians);
    const sin = angle === 0 ? 0 : Math.sin(radians);
    line += pdfNumber(cos) + ' ' + pdfNumber(-sin) + ' ' + pdfNumber(sin) + ' ' + pdfNumber(cos) + ' ' + pdfNumber(x * scale, 3) + ' ' + pdfNumber((height - y) * scale, 3) + ' Tm <' + encoded.hex + '> Tj';
    out.push(line);
    return true;
  };
  const lines = linesOf(runs);
  for (let i = 0; i < lines.length; i++) {
    let last: { run: PdfTextRun; text: string; width: number } | undefined;
    for (let k = 0; k < lines[i].length; k++) {
      const run = lines[i][k];
      const text = cut(run.text);
      // What is cut off the text is cut off its width.
      const width = text.length < run.text.length && finite(run.width) ? (run.width * text.length) / run.text.length : run.width;
      const angle = finite(run.angle) ? run.angle % 360 : 0;
      // Text that stands apart from the text before it (the next cell of a row) and brings no space is given one, as wide as the room between the two.
      if (last && angle === 0 && !/\s$/.test(last.text) && !/^\s/.test(text)) {
        const end = last.run.x + (finite(last.width) && last.width > 0 ? last.width : 0);
        const room = run.x - end;
        const apart = last.run.cutAfter === true || run.cutBefore === true;
        // Text that a clip cut off is other text than what stands next to it, however close the two are.
        if (room > Math.min(last.run.size, run.size) * WORD_GAP || apart) write(' ', end, last.run.y, Math.max(room, last.run.size * 0.05), last.run.size, 0);
      }
      if (write(text, run.x, run.y, width, run.size, angle)) last = { run, text, width };
    }
    if (last && !/\s$/.test(last.text)) {
      const run = last.run;
      const angle = finite(run.angle) ? run.angle % 360 : 0;
      const advance = finite(last.width) && last.width > 0 ? last.width : 0;
      const radians = (angle * Math.PI) / 180;
      write(' ', run.x + advance * Math.cos(radians), run.y + advance * Math.sin(radians), run.size * 0.25, run.size, angle);
    }
  }
  if (out.length === 0) return '';
  return 'BT\n3 Tr\n' + out.join('\n') + '\nET\n';
}

/**
 * The text of a laid out page of a document as runs: every text item, where
 * it is painted, as far as it lies on the paper. Text that is shown is in the
 * runs, also where it is no part of the text of the document any more (struck
 * out by a tracked deletion) or stands a second time (the header row of a
 * table, repeated on a new page): what can be read in the picture can be
 * found in it. Hidden text is not laid out, so it is in no item.
 *
 * The pages of a document are painted without clips: an item is seen unless
 * it lies beyond the edge of the paper. `measure` says how wide a text is in
 * a font; with it an item that crosses the edge is cut where the paper ends,
 * without it such an item is left out.
 *
 * The text of a table cell and of a text box is a block: its lines are read
 * one after the other, not across the cells of the row.
 */
export function docPageRuns(page: LayoutPage, measure?: (font: LayoutFont, text: string) => number): PdfTextRun[] {
  const runs: PdfTextRun[] = [];
  const items = page && Array.isArray(page.items) ? page.items : [];
  const paper: ClipBox | undefined = page && finite(page.width) && finite(page.height) ? { left: 0, top: 0, right: page.width, bottom: page.height } : undefined;
  const blocks = new Map<string, number>();
  /** Runs made up by the layout (the number of a list item), waiting for the text they stand before. */
  let waiting: PdfTextRun[] = [];
  for (let i = 0; i < items.length; i++) {
    const item = items[i];
    if (!item || item.type !== 'text' || typeof item.text !== 'string' || item.text.length === 0 || !item.font) continue;
    if (!finite(item.x) || !finite(item.y) || !finite(item.font.size) || !(item.font.size > 0)) continue;
    const whole: PdfTextRun = { text: item.text, x: item.x, y: item.y, width: finite(item.width) ? item.width : 0, size: item.font.size };
    if (item.rotation === 90 || item.rotation === 270) whole.angle = item.rotation;
    let block: number | undefined;
    const source = item.source;
    if (source && Array.isArray(source.path) && source.path.length > 1) {
      // Everything but the number of the paragraph: the cell, or the text box, the paragraph is in.
      const key = String(source.story) + ':' + source.path.slice(0, source.path.length - 1).join('.');
      block = blocks.get(key);
      if (block === undefined) {
        block = blocks.size + 1;
        blocks.set(key, block);
      }
    }
    if (source) {
      // What the layout put before this text on its line belongs where the text does.
      for (let k = 0; k < waiting.length; k++) if (block !== undefined && Math.abs(waiting[k].y - whole.y) <= whole.size * 0.5) waiting[k].block = block;
      waiting = [];
    }
    let run: PdfTextRun | undefined = whole;
    if (paper) {
      const font = item.font;
      const text = item.text;
      let full = -1;
      run = clipRun(
        whole,
        paper,
        measure
          ? (count) => {
              // Spacing and stretching are in the width of the item: the measure is scaled to it.
              if (full < 0) full = measure(font, text);
              return full > 0 ? (measure(font, text.slice(0, count)) * whole.width) / full : 0;
            }
          : undefined,
      );
    }
    if (!run) continue;
    if (block !== undefined) run.block = block;
    else if (!source) waiting.push(run);
    runs.push(run);
  }
  return runs;
}

/** What captureText() needs of a canvas context. */
export interface TextCanvas {
  fillText(text: string, x: number, y: number, maxWidth?: number): void;
  measureText(text: string): { width: number };
  getTransform?(): { a: number; b: number; c: number; d: number; e: number; f: number };
  font: string;
  textAlign: string;
  textBaseline: string;
  direction?: string;
}

/**
 * How far below the given line of the canvas the alphabetic baseline lies, in font sizes.
 * UNSURE: these are the proportions of a common text font, not those of the font that is painted;
 * the text is invisible, and a line that lies a pixel or two off is still found and selected.
 */
const BASELINE_SHIFT: { [baseline: string]: number } = { top: 0.8, hanging: 0.8, middle: 0.3, bottom: -0.2, ideographic: -0.2 };

/** Where a painter tells the listener that the text it paints next belongs together. Only there while captureText() listens. */
const BLOCK_KEY = '__bymTextBlock';

/**
 * Said by a painter around the lines of one block of text (a cell with more
 * than one line): `true` before the first line, `false` after the last. The
 * runs caught in between are one block, and are read one after the other.
 * Does nothing on a canvas nobody listens to.
 */
export function textBlock(ctx: object, open: boolean): void {
  const told = (ctx as { [BLOCK_KEY]?: unknown })[BLOCK_KEY];
  if (typeof told === 'function') (told as (open: boolean) => void)(open);
}

/** The clip of a canvas as far as it is followed: a box in canvas pixels, or nothing that can be seen through. */
interface ClipState {
  left: number;
  top: number;
  right: number;
  bottom: number;
  /** The canvas was clipped by something that could not be followed: nothing under it is taken for seen. */
  blind: boolean;
}

/** The methods that add to the path of a canvas, with the indexes of the arguments that are points of it (x, y pairs), and of a radius around the first. */
const PATH_POINTS: { [method: string]: { points: number[]; radius?: number[]; box?: boolean } } = {
  moveTo: { points: [0] },
  lineTo: { points: [0] },
  bezierCurveTo: { points: [0, 2, 4] },
  quadraticCurveTo: { points: [0, 2] },
  arcTo: { points: [0, 2] },
  arc: { points: [0], radius: [2] },
  ellipse: { points: [0], radius: [2, 3] },
  rect: { points: [0], box: true },
  roundRect: { points: [0], box: true },
};

/**
 * Catches the text a painter writes on a canvas: from now on every
 * fillText() is also told to `add`, as a run in px of the page (`scale` is
 * the number of canvas pixels per such px), with the place, the size and the
 * turn the transform of the canvas gives it.
 *
 * Only what can be seen is told. The clip of the canvas is followed (save,
 * restore, the path and clip(); a clip that is no rectangle counts as the
 * box around it, and one that cannot be followed hides everything under it),
 * `bounds` is the size of the canvas in pixels, which clips as well; a run is
 * cut to what shows of it (see clipRun), and one that does not show is not
 * told at all.
 *
 * Nothing about the painting changes: every call goes through to the canvas
 * with the arguments it came with, and a canvas that cannot say what its
 * transform is paints without being listened to. Returns what ends it.
 */
export function captureText(ctx: TextCanvas, scale: number, add: (run: PdfTextRun) => void, bounds?: { width: number; height: number }): () => void {
  const target = ctx as unknown as { [name: string]: unknown };
  const patched: { name: string; own: boolean; original: unknown; wrapper: unknown }[] = [];
  /** Puts `before` in front of a method of the canvas; what it throws never reaches the painter. */
  const wrap = (name: string, before: (args: IArguments) => void): boolean => {
    const original = target[name];
    if (typeof original !== 'function') return false;
    const own = Object.prototype.hasOwnProperty.call(target, name);
    const wrapper = function (this: unknown): unknown {
      try {
        before(arguments);
      } catch {
        // Listening must never keep the canvas from being painted on.
        clip.blind = true;
      }
      // With the arguments as they came: one that is undefined is not the same as none.
      return (original as (...args: unknown[]) => unknown).apply(this, arguments as unknown as unknown[]);
    };
    try {
      target[name] = wrapper;
    } catch {
      return false;
    }
    patched.push({ name, own, original, wrapper });
    return true;
  };

  const whole = (): ClipState =>
    bounds && bounds.width > 0 && bounds.height > 0 ? { left: 0, top: 0, right: bounds.width, bottom: bounds.height, blind: false } : { left: -Infinity, top: -Infinity, right: Infinity, bottom: Infinity, blind: false };
  let clip = whole();
  const saved: ClipState[] = [];
  // The box around the path, in canvas pixels; `open` while every point of it could be followed.
  let path = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity, known: true };
  const matrixNow = (): { a: number; b: number; c: number; d: number; e: number; f: number } | undefined => {
    if (typeof ctx.getTransform !== 'function') return undefined;
    const m = ctx.getTransform();
    return m && finite(m.a) && finite(m.b) && finite(m.c) && finite(m.d) && finite(m.e) && finite(m.f) ? m : undefined;
  };
  const point = (x: unknown, y: unknown, reach = 0): void => {
    const m = matrixNow();
    if (!m || !finite(x) || !finite(y) || !finite(reach)) {
      path.known = false;
      return;
    }
    // A reach around the point (the radius of an arc) is the box around it.
    const corners = reach > 0 ? [[x - reach, y - reach], [x + reach, y - reach], [x - reach, y + reach], [x + reach, y + reach]] : [[x, y]];
    for (let i = 0; i < corners.length; i++) {
      const px = m.a * corners[i][0] + m.c * corners[i][1] + m.e;
      const py = m.b * corners[i][0] + m.d * corners[i][1] + m.f;
      if (px < path.left) path.left = px;
      if (px > path.right) path.right = px;
      if (py < path.top) path.top = py;
      if (py > path.bottom) path.bottom = py;
    }
  };

  // The text first: without it nothing is listened to at all.
  let block: number | undefined;
  let blocks = 0;
  const listen = (text: unknown, x: unknown, y: unknown, maxWidth?: unknown): void => {
    if (typeof text !== 'string' || text.length === 0 || !finite(x) || !finite(y) || !(scale > 0) || clip.blind) return;
    const size = /(\d+(?:\.\d+)?|\.\d+)px/.exec(String(ctx.font));
    if (!size) return;
    const px = Number(size[1]);
    const matrix = matrixNow();
    if (!matrix || !(px > 0)) return;
    let measured = ctx.measureText(text).width;
    if (!finite(measured) || measured < 0) measured = 0;
    // A width the text is squeezed into squeezes every character alike.
    const squeeze = finite(maxWidth) && maxWidth >= 0 && measured > maxWidth ? maxWidth / measured : 1;
    const width = measured * squeeze;
    const rtl = ctx.direction === 'rtl';
    const align = ctx.textAlign;
    const left = align === 'center' ? x - width / 2 : align === 'right' || (align === 'end' && !rtl) || (align === 'start' && rtl) ? x - width : x;
    const baseline = y + (BASELINE_SHIFT[ctx.textBaseline] || 0) * px;
    const across = Math.sqrt(matrix.a * matrix.a + matrix.b * matrix.b);
    const down = Math.sqrt(matrix.c * matrix.c + matrix.d * matrix.d);
    if (!(across > 0) || !(down > 0)) return;
    // In canvas pixels, where the clip is known; then in px of the page.
    const whole: PdfTextRun = {
      text,
      x: matrix.a * left + matrix.c * baseline + matrix.e,
      y: matrix.b * left + matrix.d * baseline + matrix.f,
      width: width * across,
      size: px * down,
    };
    const angle = Math.round(((Math.atan2(matrix.b, matrix.a) * 180) / Math.PI) * 1e6) / 1e6;
    if (angle !== 0) whole.angle = angle < 0 ? angle + 360 : angle;
    const seen = clipRun(whole, clip, (count) => ctx.measureText(text.slice(0, count)).width * squeeze * across);
    if (!seen) return;
    const shown = cut(seen.text);
    const run: PdfTextRun = {
      text: shown,
      x: seen.x / scale,
      y: seen.y / scale,
      // What is cut off a very long text is cut off its width.
      width: ((shown.length < seen.text.length ? (seen.width * shown.length) / seen.text.length : seen.width)) / scale,
      size: seen.size / scale,
    };
    if (whole.angle !== undefined) run.angle = whole.angle;
    if (block !== undefined) run.block = block;
    if (seen.cutBefore) run.cutBefore = true;
    if (seen.cutAfter || shown.length < seen.text.length) run.cutAfter = true;
    add(run);
  };
  if (
    !wrap('fillText', (args) => {
      try {
        listen(args[0], args[1], args[2], args[3]);
      } catch {
        // A text that could not be listened to is no reason to stop following the clip.
      }
    })
  )
    return () => undefined;

  wrap('save', () => {
    saved.push({ left: clip.left, top: clip.top, right: clip.right, bottom: clip.bottom, blind: clip.blind });
  });
  wrap('restore', () => {
    const state = saved.pop();
    if (state) clip = state;
  });
  wrap('reset', () => {
    saved.length = 0;
    clip = whole();
    path = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity, known: true };
  });
  wrap('beginPath', () => {
    path = { left: Infinity, top: Infinity, right: -Infinity, bottom: -Infinity, known: true };
  });
  for (const name of Object.keys(PATH_POINTS)) {
    const shape = PATH_POINTS[name];
    wrap(name, (args) => {
      if (shape.box) {
        // A rectangle: its four corners.
        const x = args[0];
        const y = args[1];
        const w = args[2];
        const h = args[3];
        if (!finite(x) || !finite(y) || !finite(w) || !finite(h)) {
          path.known = false;
          return;
        }
        point(x, y);
        point(x + w, y);
        point(x, y + h);
        point(x + w, y + h);
        return;
      }
      let reach = 0;
      if (shape.radius) for (let i = 0; i < shape.radius.length; i++) reach = Math.max(reach, finite(args[shape.radius[i]]) ? Math.abs(args[shape.radius[i]]) : NaN);
      for (let i = 0; i < shape.points.length; i++) point(args[shape.points[i]], args[shape.points[i] + 1], reach);
    });
  }
  wrap('clip', (args) => {
    // clip(path) and clip(path, rule): a path object, which is not followed.
    if ((args.length > 0 && typeof args[0] === 'object' && args[0] !== null) || !path.known) {
      clip.blind = true;
      return;
    }
    // The clip only ever narrows; an empty path leaves nothing.
    clip = { left: Math.max(clip.left, path.left), top: Math.max(clip.top, path.top), right: Math.min(clip.right, path.right), bottom: Math.min(clip.bottom, path.bottom), blind: clip.blind };
  });

  let told = false;
  try {
    target[BLOCK_KEY] = (open: boolean) => {
      block = open ? ++blocks : undefined;
    };
    told = true;
  } catch {
    // A canvas that takes no property: its text is read line by line.
  }

  return () => {
    for (let i = patched.length - 1; i >= 0; i--) {
      const entry = patched[i];
      if (target[entry.name] !== entry.wrapper) continue;
      if (entry.own) target[entry.name] = entry.original;
      else delete target[entry.name];
    }
    if (told) delete target[BLOCK_KEY];
  };
}
