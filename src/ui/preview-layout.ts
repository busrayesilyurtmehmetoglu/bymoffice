// What the print preview works out without a DOM: where each page lies at a
// zoom (pages flow in rows, as many to a row as the width holds, so that a
// small zoom shows them as thumbnails), which pages a scroll position shows,
// which page is "the current one", the steps of the zoom, how many pixels a
// page is painted with under the budget of the device, and the states the
// preview goes through. The view (sheet-view.ts) and the overlay
// (print-preview.ts) use these; the tests check them on their own.

/** The size of a page in px of the paper (96 to the inch). */
export interface PreviewPageSize {
  width: number;
  height: number;
}

/** Where a page lies in the scrolled content, in CSS px. */
export interface PreviewSlot {
  x: number;
  y: number;
  width: number;
  height: number;
  /** The row it is in. */
  row: number;
}

export interface PreviewRow {
  top: number;
  bottom: number;
  /** The pages of the row: first..last. */
  first: number;
  last: number;
}

export interface PreviewLayout {
  slots: PreviewSlot[];
  rows: PreviewRow[];
  /** Size of the content. */
  width: number;
  height: number;
  zoom: number;
}

/** Space around and between the pages, in CSS px. */
export const PREVIEW_GAP = 16;
/** The steps of the zoom buttons and keys; "fit" lies wherever it lies between them. */
export const PREVIEW_ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
export const MIN_PREVIEW_ZOOM = PREVIEW_ZOOM_STEPS[0];
export const MAX_PREVIEW_ZOOM = PREVIEW_ZOOM_STEPS[PREVIEW_ZOOM_STEPS.length - 1];

export const clampZoom = (zoom: number): number => (zoom > 0 && zoom < Infinity ? Math.min(MAX_PREVIEW_ZOOM, Math.max(MIN_PREVIEW_ZOOM, zoom)) : 1);

/** A usable size: anything else (from a damaged page setup) reads as A4. */
function sizeOf(page: PreviewPageSize | undefined): PreviewPageSize {
  const ok = (n: unknown) => typeof n === 'number' && n >= 1 && n < 1e6;
  return page && ok(page.width) && ok(page.height) ? page : { width: 794, height: 1123 };
}

/**
 * Lays the pages out at a zoom in a content as wide as `viewportWidth`: rows
 * of as many pages as fit side by side (one at least), each row as high as its
 * highest page, the pages centred in it and the rows centred across. O(pages).
 */
export function layoutPreview(pages: readonly PreviewPageSize[], zoom: number, viewportWidth: number, gap: number = PREVIEW_GAP): PreviewLayout {
  const z = clampZoom(zoom);
  const room = Math.max(1, (viewportWidth > 0 && viewportWidth < Infinity ? viewportWidth : 0) - 2 * gap);
  const slots: PreviewSlot[] = new Array(pages.length);
  const rows: PreviewRow[] = [];
  let y = gap;
  let widest = 0;
  let i = 0;
  while (i < pages.length) {
    // As many pages as the width holds.
    const first = i;
    let width = 0;
    let height = 0;
    do {
      const size = sizeOf(pages[i]);
      const w = Math.round(size.width * z);
      if (i > first && width + gap + w > room) break;
      width += (i > first ? gap : 0) + w;
      height = Math.max(height, Math.round(size.height * z));
      i++;
    } while (i < pages.length);
    const last = i - 1;
    const row = rows.length;
    let x = gap + Math.max(0, (room - width) / 2);
    for (let p = first; p <= last; p++) {
      const size = sizeOf(pages[p]);
      const w = Math.round(size.width * z);
      const h = Math.round(size.height * z);
      slots[p] = { x: Math.round(x), y: y + Math.round((height - h) / 2), width: w, height: h, row };
      x += w + gap;
    }
    rows.push({ top: y, bottom: y + height, first, last });
    widest = Math.max(widest, width);
    y += height + gap;
  }
  return { slots, rows, width: Math.max(widest + 2 * gap, viewportWidth > 0 && viewportWidth < Infinity ? viewportWidth : 0), height: rows.length > 0 ? y : 2 * gap, zoom: z };
}

/** The first row whose bottom lies below `y`; rows.length when none does. */
function rowBelow(rows: readonly PreviewRow[], y: number): number {
  let low = 0;
  let high = rows.length;
  while (low < high) {
    const mid = (low + high) >>> 1;
    if (rows[mid].bottom + PREVIEW_GAP <= y) low = mid + 1;
    else high = mid;
  }
  return low;
}

/**
 * The pages a scroll position shows, with `overscan` px more above and below:
 * those that are worth painting now. Undefined when none is (no pages).
 */
export function visiblePages(layout: PreviewLayout, scrollTop: number, viewportHeight: number, overscan = 0): { first: number; last: number } | undefined {
  const rows = layout.rows;
  if (rows.length === 0) return undefined;
  const top = Math.max(0, scrollTop - Math.max(0, overscan));
  const bottom = scrollTop + Math.max(0, viewportHeight) + Math.max(0, overscan);
  const first = Math.min(rowBelow(rows, top), rows.length - 1);
  let last = first;
  while (last + 1 < rows.length && rows[last + 1].top < bottom) last++;
  return { first: rows[first].first, last: rows[last].last };
}

/**
 * The page a scroll position is at: the first page of the row that lies under
 * a third of the height of the viewport; `current` when it is in that row
 * already (a row of several pages keeps the one that was chosen).
 */
export function pageAtScroll(layout: PreviewLayout, scrollTop: number, viewportHeight: number, current: number): number {
  const rows = layout.rows;
  if (rows.length === 0) return 0;
  const probe = scrollTop + Math.max(0, viewportHeight) / 3;
  const index = Math.min(rowBelow(rows, probe), rows.length - 1);
  const row = rows[index];
  return current >= row.first && current <= row.last ? current : row.first;
}

/** Where the scroll goes to show a page: its row at the top, less the gap. */
export function scrollForPage(layout: PreviewLayout, page: number): number {
  const slot = layout.slots[Math.max(0, Math.min(page, layout.slots.length - 1))];
  return slot ? Math.max(0, layout.rows[slot.row].top - PREVIEW_GAP) : 0;
}

/** The zoom that shows a whole page in the viewport. */
export function fitZoom(page: PreviewPageSize | undefined, viewportWidth: number, viewportHeight: number, gap: number = PREVIEW_GAP): number {
  const size = sizeOf(page);
  const w = Math.max(1, viewportWidth - 2 * gap);
  const h = Math.max(1, viewportHeight - 2 * gap);
  return clampZoom(Math.min(w / size.width, h / size.height));
}

/** The next step of the zoom up (+1) or down (-1) from where it is. */
export function stepZoom(zoom: number, direction: 1 | -1): number {
  const z = clampZoom(zoom);
  if (direction > 0) {
    for (let i = 0; i < PREVIEW_ZOOM_STEPS.length; i++) if (PREVIEW_ZOOM_STEPS[i] > z + 1e-6) return PREVIEW_ZOOM_STEPS[i];
    return MAX_PREVIEW_ZOOM;
  }
  for (let i = PREVIEW_ZOOM_STEPS.length - 1; i >= 0; i--) if (PREVIEW_ZOOM_STEPS[i] < z - 1e-6) return PREVIEW_ZOOM_STEPS[i];
  return MIN_PREVIEW_ZOOM;
}

/** How high the content may grow: browsers stop sizing an element somewhere past 17 M px (Firefox); the grid keeps to 8 M too. */
export const MAX_PREVIEW_HEIGHT = 8_000_000;

/**
 * The largest zoom at which the pages, one under the other, stay within MAX_PREVIEW_HEIGHT; never
 * below the smallest zoom (pages so tall that even that is too much are cut off at the end: a known limit).
 */
export function maxPreviewZoom(pages: readonly PreviewPageSize[], gap: number = PREVIEW_GAP): number {
  let height = 0;
  for (let i = 0; i < pages.length; i++) height += sizeOf(pages[i]).height;
  const room = MAX_PREVIEW_HEIGHT - gap * (pages.length + 1);
  if (height <= 0) return MAX_PREVIEW_ZOOM;
  const most = room / height;
  return most >= MAX_PREVIEW_ZOOM ? MAX_PREVIEW_ZOOM : Math.max(MIN_PREVIEW_ZOOM, Math.floor(most * 100) / 100);
}

/**
 * What the preview builds for a scroll position: paper (an element) for every page in view and
 * half a screen around it, and canvases for every page in view - however many: their pixels are
 * shared out under the budget, so more of them only makes each coarser - and for the pages around
 * it, nearest the current page first, while there are fewer than `canvases`.
 */
export function previewWindow(layout: PreviewLayout, scrollTop: number, viewportHeight: number, current: number, canvases: number): { first: number; last: number; visible: number; canvases: number[] } {
  const range = visiblePages(layout, scrollTop, viewportHeight, viewportHeight / 2);
  if (!range) return { first: 0, last: -1, visible: 0, canvases: [] };
  const inView = visiblePages(layout, scrollTop, viewportHeight, 0) || range;
  const out: number[] = [];
  for (let page = inView.first; page <= inView.last; page++) out.push(page);
  const around: number[] = [];
  for (let page = range.first; page <= range.last; page++) if (page < inView.first || page > inView.last) around.push(page);
  around.sort((a, b) => Math.abs(a - current) - Math.abs(b - current) || a - b);
  for (let i = 0; i < around.length && out.length < canvases; i++) out.push(around[i]);
  out.sort((a, b) => a - b);
  return { first: range.first, last: range.last, visible: inView.last - inView.first + 1, canvases: out };
}

/**
 * How long a page whose pictures or formats are still to come is painted again: a time per
 * layout, from its first painting in that layout. Painting it again does not start it over.
 */
export class SettleClock {
  private generation = -1;
  private since = 0;
  constructor(private readonly limit: number) {}

  /** Whether a page painted unsettled in `generation` at `now` (ms) is to be painted again. */
  again(generation: number, now: number): boolean {
    if (generation !== this.generation) {
      this.generation = generation;
      this.since = now;
    }
    return now - this.since <= this.limit;
  }
}

/** What one canvas may hold, and all canvases of the preview together, in pixels. */
export interface PreviewBudget {
  pagePixels: number;
  totalPixels: number;
  /** Canvases at most: pages beyond them stay blank until they are scrolled to. */
  canvases: number;
}

/** Phones and tablets get less. 4 bytes a pixel: 128 MB and 48 MB of canvases at most. */
export function previewBudget(coarse: boolean): PreviewBudget {
  return coarse ? { pagePixels: 4 * 1024 * 1024, totalPixels: 12 * 1024 * 1024, canvases: 24 } : { pagePixels: 16 * 1024 * 1024, totalPixels: 32 * 1024 * 1024, canvases: 64 };
}

/**
 * The pixel width each of the pages shown is painted with: its CSS width
 * times the device pixel ratio, made smaller alike for all of them when their
 * pixels together would pass the budget, and for one alone when it would pass
 * the limit of a canvas. At least 1.
 */
export function previewPixelWidths(slots: readonly PreviewSlot[], pages: readonly PreviewPageSize[], indexes: readonly number[], ratio: number, budget: PreviewBudget): number[] {
  const r = ratio > 0 && ratio <= 8 ? ratio : 1;
  let total = 0;
  for (let i = 0; i < indexes.length; i++) {
    const slot = slots[indexes[i]];
    if (slot) total += slot.width * r * slot.height * r;
  }
  const shared = total > budget.totalPixels ? Math.sqrt(budget.totalPixels / total) : 1;
  const out: number[] = [];
  for (let i = 0; i < indexes.length; i++) {
    const slot = slots[indexes[i]];
    const size = sizeOf(pages[indexes[i]]);
    if (!slot) {
      out.push(1);
      continue;
    }
    let width = slot.width * r * shared;
    const height = (width * size.height) / size.width;
    if (width * height > budget.pagePixels) width = Math.sqrt((budget.pagePixels * size.width) / size.height);
    out.push(Math.max(1, Math.floor(width)));
  }
  return out;
}

/**
 * The canvases of the pages shown: one per page that is in the window, given
 * back when its page leaves it, so that a preview of 2,000 pages holds no more
 * than the window does. A canvas given back is emptied (its pixels let go of)
 * and kept for the next page, up to `spare` of them.
 */
export class CanvasPool<C extends { width: number; height: number }> {
  private readonly free: C[] = [];
  private readonly used = new Map<number, C>();
  constructor(
    private readonly make: () => C,
    private readonly spare = 4,
  ) {}

  /** The canvas of a page: the one it has, or one taken for it. */
  take(page: number): C {
    const had = this.used.get(page);
    if (had) return had;
    const canvas = this.free.pop() || this.make();
    this.used.set(page, canvas);
    return canvas;
  }

  /** The canvas a page has, if any. */
  get(page: number): C | undefined {
    return this.used.get(page);
  }

  /** Gives back the canvases of the pages not in `keep`; returns them so that their elements can be taken out. */
  keepOnly(keep: (page: number) => boolean): { page: number; canvas: C }[] {
    const gone: { page: number; canvas: C }[] = [];
    this.used.forEach((canvas, page) => {
      if (!keep(page)) gone.push({ page, canvas });
    });
    for (let i = 0; i < gone.length; i++) this.release(gone[i].page);
    return gone;
  }

  release(page: number): void {
    const canvas = this.used.get(page);
    if (!canvas) return;
    this.used.delete(page);
    // Its pixels go at once; some browsers keep them until collection otherwise.
    canvas.width = 0;
    canvas.height = 0;
    if (this.free.length < this.spare) this.free.push(canvas);
  }

  /** Lets go of every canvas. */
  clear(): void {
    this.used.forEach((_canvas, page) => this.release(page));
    this.used.clear();
    for (let i = 0; i < this.free.length; i++) {
      this.free[i].width = 0;
      this.free[i].height = 0;
    }
    this.free.length = 0;
  }

  /** Canvases that pages hold now. */
  get size(): number {
    return this.used.size;
  }

  /** Pixels the canvases of the pages hold now. */
  get pixels(): number {
    let total = 0;
    this.used.forEach((canvas) => (total += canvas.width * canvas.height));
    return total;
  }

  pages(): number[] {
    return Array.from(this.used.keys()).sort((a, b) => a - b);
  }
}

/**
 * The states of a preview: 'planning' while the pages are being laid out,
 * 'ready' with pages, 'empty' when there is nothing to print, 'failed' when
 * the pages cannot be made (too many), 'closed' once it is gone.
 */
export type PreviewState = 'closed' | 'planning' | 'ready' | 'empty' | 'failed';
export type PreviewEvent = 'open' | 'planned' | 'nothing' | 'error' | 'changed' | 'close';

/** The state after an event. A closed preview stays closed but for 'open'; an open one ignores 'open'. */
export function previewTransition(state: PreviewState, event: PreviewEvent): PreviewState {
  if (event === 'close') return 'closed';
  if (state === 'closed') return event === 'open' ? 'planning' : 'closed';
  switch (event) {
    case 'open':
      return state;
    case 'changed':
      return 'planning';
    case 'planned':
      return state === 'planning' ? 'ready' : state;
    case 'nothing':
      return state === 'planning' ? 'empty' : state;
    case 'error':
      return state === 'planning' ? 'failed' : state;
    default:
      return state;
  }
}
