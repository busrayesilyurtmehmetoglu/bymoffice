// Where places of a document are on its pages: the box of a caret, the place
// under a point, the line above and below a line. All of it is read from the
// laid out pages (their text items and line marks, see LineMark), so nothing
// here knows the model, and nothing touches the DOM.

import type { LayoutPage, LineMark, TextItem, TextMeasurer } from '../layout/doc/types';

/** A place in the text: the same as TextItem.source and as DocPosition of the operations. */
export interface Place {
  story: string;
  path: number[];
  offset: number;
  /** Before what shows without width at the offset (DocPosition.before). */
  before?: true;
}

/**
 * Which of two lines a place on the border between them belongs to: the
 * start of the second one (the default), or the end of the first.
 */
export type Affinity = 'start' | 'end';

/** One line of a paragraph on a page, with the text items that lie on it, in the order of their offsets. */
export interface PageLine {
  mark: LineMark;
  texts: TextItem[];
}

export interface LineAt {
  /** Index of the page. */
  page: number;
  line: PageLine;
}

/** The box of a caret, in px of its page. */
export interface CaretBox extends LineAt {
  x: number;
  top: number;
  height: number;
}

interface PageIndex {
  lines: PageLine[];
  byParagraph: Map<string, PageLine[]>;
  /** Lowest and highest block of the body the page has lines of; -1 when it has none. */
  minBlock: number;
  maxBlock: number;
}

/** Pages looked at around a hint for a place that is not in the body. */
const MAX_SCAN = 4096;

export const paragraphKey = (story: string, path: number[]): string => story + '\u0001' + path.join('.');

const distance = (value: number, low: number, high: number): number => (value < low ? low - value : value > high ? value - high : 0);

/** Offset just past the text of the line: where a caret at its end stands. */
export const lineEnd = (line: PageLine): number => line.mark.source.offset + line.mark.length;

export class DocLines {
  private readonly cache = new WeakMap<LayoutPage, PageIndex>();

  /** `pages` is read anew at every call: the list a layout engine fills. */
  constructor(
    private readonly pages: () => LayoutPage[],
    private readonly measurer: TextMeasurer,
  ) {}

  private index(page: LayoutPage): PageIndex {
    let index = this.cache.get(page);
    if (index) return index;
    index = { lines: [], byParagraph: new Map(), minBlock: -1, maxBlock: -1 };
    const items = page.items;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      // Copies of header rows and lines turned in a cell are not places of the text.
      if (item.type !== 'mark' || item.repeat || item.rotation !== undefined) continue;
      const line: PageLine = { mark: item, texts: [] };
      index.lines.push(line);
      const key = paragraphKey(item.source.story, item.source.path);
      const list = index.byParagraph.get(key);
      if (list) list.push(line);
      else index.byParagraph.set(key, [line]);
      if (item.source.story === 'body') {
        const block = item.source.path[0];
        if (index.minBlock < 0 || block < index.minBlock) index.minBlock = block;
        if (block > index.maxBlock) index.maxBlock = block;
      }
    }
    index.byParagraph.forEach((list) => list.sort((a, b) => a.mark.source.offset - b.mark.source.offset));
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type !== 'text' || !item.source || item.repeat || item.rotation !== undefined || item.text.length === 0) continue;
      const list = index.byParagraph.get(paragraphKey(item.source.story, item.source.path));
      if (!list) continue;
      // The last line that starts at or before the text.
      let at = list.length - 1;
      while (at > 0 && list[at].mark.source.offset > item.source.offset) at--;
      list[at].texts.push(item);
    }
    for (let i = 0; i < index.lines.length; i++) index.lines[i].texts.sort((a, b) => (a.source as Place).offset - (b.source as Place).offset || a.x - b.x);
    this.cache.set(page, index);
    return index;
  }

  /** The lines of a page, in the order the page holds them. */
  linesOf(page: number): PageLine[] {
    const found = this.pages()[page];
    return found ? this.index(found).lines : [];
  }

  /** Advance from the left edge of a text item to its character `offset`. */
  advance(item: TextItem, offset: number): number {
    if (offset <= 0) return 0;
    if (offset >= item.text.length) return item.width;
    const width = this.measurer.width(item.font, item.text.slice(0, offset)) * (item.scale || 1) + (item.letterSpacing || 0) * offset;
    return width < 0 ? 0 : width > item.width ? item.width : width;
  }

  /** A page that has lines of the block `block` of the body; -1 when none has. */
  private bodyPage(block: number): number {
    const pages = this.pages();
    let low = 0;
    let high = pages.length - 1;
    while (low <= high) {
      const mid = (low + high) >> 1;
      // Pages without text of the body (an empty page before a section) are passed over.
      let at = mid;
      while (at <= high && this.index(pages[at]).minBlock < 0) at++;
      if (at > high) {
        high = mid - 1;
        continue;
      }
      const index = this.index(pages[at]);
      if (block < index.minBlock) high = mid - 1;
      else if (block > index.maxBlock) low = at + 1;
      else return at;
    }
    return -1;
  }

  /**
   * The line a place is on. `hint` is the page to look at first: a header is
   * on every page, and the one meant is the one the hint names.
   */
  lineOf(place: Place, affinity: Affinity = 'start', hint = 0): LineAt | undefined {
    const pages = this.pages();
    const count = pages.length;
    if (count === 0) return undefined;
    const key = paragraphKey(place.story, place.path);
    let page = -1;
    if (place.story === 'body') {
      const hinted = pages[hint];
      page = hinted && this.index(hinted).byParagraph.has(key) ? hint : this.bodyPage(place.path[0]);
      if (page < 0) return undefined;
      // The block may spread over pages; the paragraph is on one of those around.
      if (!this.index(pages[page]).byParagraph.has(key)) {
        let found = -1;
        for (let step = 1; step < count && found < 0; step++) {
          const before = page - step;
          const after = page + step;
          if (before < 0 && after >= count) break;
          if (after < count && this.index(pages[after]).byParagraph.has(key)) found = after;
          else if (before >= 0 && this.index(pages[before]).byParagraph.has(key)) found = before;
          else if ((before < 0 || this.index(pages[before]).maxBlock < place.path[0]) && (after >= count || this.index(pages[after]).minBlock > place.path[0])) break;
        }
        if (found < 0) return undefined;
        page = found;
      }
    } else {
      const start = Math.max(0, Math.min(count - 1, hint));
      for (let step = 0; step < MAX_SCAN && page < 0; step++) {
        const after = start + step;
        const before = start - step;
        if (after >= count && before < 0) break;
        if (after < count && this.index(pages[after]).byParagraph.has(key)) page = after;
        else if (step > 0 && before >= 0 && this.index(pages[before]).byParagraph.has(key)) page = before;
      }
      if (page < 0) return undefined;
    }
    // From the page found to the one that holds the offset: a paragraph may go on over pages.
    for (let guard = 0; guard < count; guard++) {
      const lines = this.index(pages[page]).byParagraph.get(key) as PageLine[];
      const first = lines[0];
      const last = lines[lines.length - 1];
      const previous = page > 0 ? this.index(pages[page - 1]).byParagraph.get(key) : undefined;
      const next = page + 1 < count ? this.index(pages[page + 1]).byParagraph.get(key) : undefined;
      if (place.offset < first.mark.source.offset && previous && lineEnd(previous[previous.length - 1]) <= first.mark.source.offset) {
        page--;
        continue;
      }
      const end = lineEnd(last);
      if (!last.mark.last && next && next[0].mark.source.offset >= end && (place.offset > end || (place.offset === end && affinity !== 'end'))) {
        page++;
        continue;
      }
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const to = lineEnd(line);
        if (place.offset < to) return { page, line };
        if (place.offset === to && (line.mark.last || affinity === 'end' || i === lines.length - 1)) return { page, line };
      }
      return { page, line: last };
    }
    return undefined;
  }

  /** The box of a caret at a place. */
  caret(place: Place, affinity: Affinity = 'start', hint = 0): CaretBox | undefined {
    const at = this.lineOf(place, affinity, hint);
    if (!at) return undefined;
    const { mark, texts } = at.line;
    const offset = place.offset;
    // The text the place is in or at the end of, the text that starts at it, and the pictures at it.
    let left: TextItem | undefined;
    let right: TextItem | undefined;
    for (let i = 0; i < texts.length; i++) {
      const text = texts[i];
      const from = (text.source as Place).offset;
      if (from > offset) break;
      if (from === offset) {
        if (!right) right = text;
        continue;
      }
      if (offset <= from + text.text.length) left = text;
    }
    let boxLeft = Infinity;
    let boxRight = -Infinity;
    if (mark.boxes) {
      for (const box of mark.boxes) {
        if (box.offset !== offset) continue;
        if (mark.x + box.dx < boxLeft) boxLeft = mark.x + box.dx;
        if (mark.x + box.dx + box.width > boxRight) boxRight = mark.x + box.dx + box.width;
      }
    }
    const leftEdge = left ? left.x + this.advance(left, offset - (left.source as Place).offset) : undefined;
    const inside = left !== undefined && offset < (left.source as Place).offset + left.text.length;
    let x: number;
    let item = left || right;
    if (inside) x = leftEdge as number;
    else if (place.before === true) {
      // Before what shows without width there (a picture, a note mark): on its left.
      x = boxLeft !== Infinity ? boxLeft : leftEdge !== undefined ? leftEdge : right ? right.x : mark.x;
    } else if (boxRight !== -Infinity) x = boxRight;
    else if (leftEdge !== undefined && right && right.x > leftEdge + 0.5) {
      // Something without a place in the text (a note mark) lies between: the caret is after it.
      x = right.x;
      item = right;
    } else x = leftEdge !== undefined ? leftEdge : right ? right.x : mark.x;
    if (!item && boxRight !== -Infinity) x = place.before === true ? boxLeft : boxRight;
    if (item) {
      const box = this.measurer.metrics(item.font);
      return { page: at.page, line: at.line, x, top: item.y - box.ascent, height: box.ascent + box.descent };
    }
    return { page: at.page, line: at.line, x, top: mark.y + mark.baseline - mark.ascent, height: mark.ascent + mark.descent };
  }

  /** Left and right edge of what a line shows. */
  private extent(line: PageLine): [number, number] {
    let left = line.mark.x;
    let right = line.mark.x + line.mark.width;
    for (let i = 0; i < line.texts.length; i++) {
      const text = line.texts[i];
      if (text.x < left) left = text.x;
      if (text.x + text.width > right) right = text.x + text.width;
    }
    return [left, right];
  }

  /**
   * The place in a line nearest to x. On a picture: before it for its left
   * half, after it for its right half, and `box` says the picture is there.
   */
  placeAt(line: PageLine, x: number): { place: Place; affinity: Affinity; box?: boolean } {
    const source = line.mark.source;
    const texts = line.texts;
    const at = (offset: number, affinity: Affinity) => ({ place: { story: source.story, path: source.path, offset }, affinity });
    if (line.mark.boxes) {
      for (const box of line.mark.boxes) {
        const left = line.mark.x + box.dx;
        if (x < left || x > left + box.width) continue;
        const place: Place = { story: source.story, path: source.path, offset: box.offset };
        if (x < left + box.width / 2) place.before = true;
        return { place, affinity: 'start', box: true };
      }
    }
    if (texts.length === 0) return at(source.offset, 'start');
    let best = texts[0];
    let bestDistance = Infinity;
    for (let i = 0; i < texts.length; i++) {
      const d = distance(x, texts[i].x, texts[i].x + texts[i].width);
      if (d < bestDistance) {
        bestDistance = d;
        best = texts[i];
      }
    }
    const target = x - best.x;
    let low = 0;
    let high = best.text.length;
    // The last boundary at or before the point, then the nearer of it and the next.
    while (low < high) {
      const mid = (low + high + 1) >> 1;
      if (this.advance(best, mid) <= target) low = mid;
      else high = mid - 1;
    }
    if (low < best.text.length && target - this.advance(best, low) > this.advance(best, low + 1) - target) low++;
    // Not into the middle of a surrogate pair.
    if (low > 0 && low < best.text.length) {
      const code = best.text.charCodeAt(low);
      if (code >= 0xdc00 && code <= 0xdfff) low++;
    }
    const offset = (best.source as Place).offset + low;
    // At the end of a line that wraps the place is also the start of the next line: it is meant as the end of this one.
    return at(offset, offset >= lineEnd(line) && !line.mark.last && offset > source.offset ? 'end' : 'start');
  }

  /** The line of a page nearest to a point (px of the page) among those `accept` lets through. */
  lineAt(page: number, x: number, y: number, accept?: (mark: LineMark) => boolean): PageLine | undefined {
    const lines = this.linesOf(page);
    let best: PageLine | undefined;
    let bestDistance = Infinity;
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (accept && !accept(line.mark)) continue;
      const [left, right] = this.extent(line);
      // The line counts for more than the place in it.
      const d = distance(y, line.mark.y, line.mark.y + line.mark.height) * 1000 + distance(x, left, right);
      if (d < bestDistance) {
        bestDistance = d;
        best = line;
      }
    }
    return best;
  }

  /** The place nearest to a point of a page. */
  hit(page: number, x: number, y: number, accept?: (mark: LineMark) => boolean): { place: Place; affinity: Affinity; line: PageLine; box?: boolean } | undefined {
    const line = this.lineAt(page, x, y, accept);
    if (!line) return undefined;
    const found = this.placeAt(line, x);
    return { place: found.place, affinity: found.affinity, line, box: found.box };
  }

  /**
   * The line above (direction -1) or below (1) a line as the page shows
   * them, nearest to x; on the next page when the page has no more.
   */
  vertical(from: LineAt, direction: 1 | -1, x: number, accept?: (mark: LineMark) => boolean): LineAt | undefined {
    const pages = this.pages();
    const here = from.line.mark;
    for (let page = from.page, steps = 0; page >= 0 && page < pages.length && steps < MAX_SCAN; page += direction, steps++) {
      const lines = this.index(pages[page]).lines;
      let nearest = Infinity;
      for (let i = 0; i < lines.length; i++) {
        const mark = lines[i].mark;
        if (lines[i] === from.line || (accept && !accept(mark))) continue;
        let gap: number;
        if (page !== from.page) gap = direction === 1 ? mark.y : -(mark.y + mark.height);
        else if (direction === 1) gap = mark.y >= here.y + here.height - 0.5 ? mark.y - (here.y + here.height) : Infinity;
        else gap = mark.y + mark.height <= here.y + 0.5 ? here.y - (mark.y + mark.height) : Infinity;
        if (gap < nearest) nearest = gap;
      }
      if (nearest === Infinity) continue;
      // Of the lines at that height (the cells of a table row), the one nearest to x.
      let best: PageLine | undefined;
      let bestDistance = Infinity;
      for (let i = 0; i < lines.length; i++) {
        const mark = lines[i].mark;
        if (lines[i] === from.line || (accept && !accept(mark))) continue;
        let gap: number;
        if (page !== from.page) gap = direction === 1 ? mark.y : -(mark.y + mark.height);
        else if (direction === 1) gap = mark.y >= here.y + here.height - 0.5 ? mark.y - (here.y + here.height) : Infinity;
        else gap = mark.y + mark.height <= here.y + 0.5 ? here.y - (mark.y + mark.height) : Infinity;
        if (gap > nearest + mark.height / 2) continue;
        const [left, right] = this.extent(lines[i]);
        const d = distance(x, left, right);
        if (d < bestDistance) {
          bestDistance = d;
          best = lines[i];
        }
      }
      if (best) return { page, line: best };
    }
    return undefined;
  }
}
