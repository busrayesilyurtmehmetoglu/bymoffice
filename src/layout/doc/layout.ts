// The document layout engine: DocModel -> pages.
//
// The body is turned into bands (lines of paragraphs, rows of tables) block by
// block, and the bands are placed into the columns of pages one at a time.
// All state is in the object, so layout() can stop after any band and go on
// later: a long document shows its first pages at once.
//
// A band that does not fit ends the column. Bands that asked to stay with
// what follows them (keep with next, widow and orphan lines, header rows)
// are taken back out of the column and go to the next one together, unless
// that would leave the column empty. A table row that may be divided is
// divided. What does not fit into an empty column is placed there anyway, so
// every column takes at least one band and layout always ends.

import type { Block, DocModel, DrawingAnchor, DrawingPosition, HeaderFooterKind, SectionProps, Story, Table } from '../../model/doc/types';
import { getPrep, layoutBlocks, paragraphEnv } from './blocks';
import { bandsHeight, LayoutContext, shifted, type Band, type FieldTextItem, type FieldValues, type FloatSpec, type PrepSession, type StateEvent, type StoryRef } from './context';
import { floatExclusion, floatPosition, type FloatFrame, type Rect, type WrapSide } from './drawing';
import { inheritedPart, findSections, pageBorderItems, pageGeometry, type PageGeometry, type SectionInfo } from './page';
import { paragraphBands, type WrapEnv } from './paragraph';
import { isHiddenParagraph, noteLabel } from './prep';
import { formatPageNumber } from './symbols';
import { tableBands, tableLayout, type TableLayout } from './table';
import type { DocLayoutOptions, LayoutBudget, LayoutPage, LayoutProgress, PageItem, TextMeasurer } from './types';
import { EPSILON, MAX_PAGES, MIN_TEXT_EXTENT, MIN_WRAP_WIDTH, twips } from './units';

/** Where the engine goes on to another section. */
interface SectionMark {
  mark: true;
  section: number;
}

type QueueItem = Band | SectionMark;

const isMark = (item: QueueItem): item is SectionMark => (item as SectionMark).mark === true;

interface Exclusion extends Rect {
  whole: boolean;
  side: WrapSide;
}

interface PlacedFloat {
  float: FloatSpec;
  x: number;
  y: number;
}

/** A band in a column, with what is needed to take it out again. */
interface Placed {
  band: Band;
  /** Top of the band (below its space before). */
  y: number;
  /** State of the column and page before the band went in. */
  columnY: number;
  notes: number;
  notesHeight: number;
  carried: number;
  exclusions: number;
  floats?: PlacedFloat[];
  /** A copy of a header row, put in by the engine. */
  repeated?: boolean;
}

interface Column {
  index: number;
  x: number;
  width: number;
  top: number;
  bottom: number;
  y: number;
  placed: Placed[];
  /** Bands in the column that are content: copies of header rows do not count. */
  real: number;
  /** Footnote bands at the foot of the column, and the room they take (separator included). */
  notes: Band[];
  notesHeight: number;
  /** The column was started by an automatic break: space before the first band is dropped. */
  auto: boolean;
  /** The column starts at the top of the text area of the page. */
  full: boolean;
}

interface HeaderFooter {
  bands: Band[];
  height: number;
  /** Shows the number of pages of the document or section. */
  totals: boolean;
}

interface PageBuild {
  page: LayoutPage;
  geometry: PageGeometry;
  /** Top and bottom of the room for body text. */
  top: number;
  bottom: number;
  /** Columns of the section being laid out and the top they share. */
  columns: { x: number; width: number }[];
  separator: boolean;
  regionTop: number;
  /** Lowest point body text reached in the current run of columns. */
  regionEnd: number;
  /** Top of the highest footnote area of the columns closed so far. */
  notesTop: number;
  exclusions: Exclusion[];
  borders: PageItem[];
  behind: PageItem[];
  header: PageItem[];
  body: PageItem[];
  footer: PageItem[];
  front: PageItem[];
  /** Anchored drawings of the header and footer. */
  headerBehind: PageItem[];
  headerFront: PageItem[];
  /** A band was put on the page. */
  used: boolean;
  /** The page has the margins of an even page of a document with mirrored margins. */
  mirrored: boolean;
  /** What the header and footer were laid out from, for when the totals are known. */
  headerKey?: { id: string; kind: 'headers' | 'footers' };
  footerKey?: { id: string; kind: 'headers' | 'footers' };
  totals: boolean;
  firstOfSection: boolean;
}

/** What is kept of a finished page whose header or footer shows a total. */
interface TotalsRecord {
  index: number;
  build: PageBuild;
}

/**
 * The state of the engine where a page begins: enough to lay the document
 * out again from that page on. Kept for every page that does not begin in the
 * middle of a table that is still being fed.
 */
interface PageMark {
  nextBlock: number;
  feedSection: number;
  endnotesQueued: boolean;
  section: number;
  afterBreak: boolean;
  emptyColumns: number;
  /** Bands waiting to be placed. */
  queue: QueueItem[];
  carry: Band[];
  sectionStarted: boolean[];
  sectionPages: number[];
  /** Lengths of the lists of fields that show a total. */
  totalItems: number;
  totalPages: number;
}

/** Everything that changes while the engine runs, as it was when an edit came. */
interface LiveState {
  done: boolean;
  queue: QueueItem[];
  head: number;
  nextBlock: number;
  table: TableLayout | undefined;
  tableSession: PrepSession | undefined;
  feedSection: number;
  endnotesQueued: boolean;
  section: number;
  build: PageBuild | undefined;
  column: Column | undefined;
  fields: FieldValues;
  carry: Band[];
  emptyColumns: number;
  afterBreak: boolean;
  sectionStarted: boolean[];
  sectionPages: number[];
  regionBands: Band[] | undefined;
  regionBodyStart: number;
  regionTotalsStart: number;
}

/**
 * The pages an edit made stale, kept while the document is laid out again
 * from the page of the edit: as soon as a page begins in the state it began
 * in before, the old pages from there on are the new ones.
 */
interface OldLayout {
  /** Index of the first page laid out again. */
  from: number;
  /** First block after the ones the edit changed. */
  end: number;
  /** Pages and marks from `from` on. */
  pages: LayoutPage[];
  marks: (PageMark | undefined)[];
  live: LiveState;
  /** Lengths of the lists of totals at `from`, and what the lists held after that. */
  itemsBase: number;
  pagesBase: number;
  totalItems: { page: number; item: FieldTextItem }[];
  totalPages: TotalsRecord[];
}

/** What DocLayout.edited() did. */
export interface EditProgress extends LayoutProgress {
  /** Index of the first page that was laid out again; the number of pages when none was. */
  from: number;
}

/**
 * What the layout of the body depends on besides the blocks of the body, as
 * a list of objects: the blocks of the headers and footers, the notes, the
 * list definitions. The operations of ops/doc put new objects where they
 * change one of these, so two lists that hold the same objects mean nothing
 * of it changed.
 */
function sideParts(doc: DocModel): unknown[] {
  const out: unknown[] = [doc.body.section, doc.styles, doc.settings];
  const stories = (map: Map<string, Story>) =>
    map.forEach((story, id) => {
      out.push(id, story.blocks.length);
      for (let i = 0; i < story.blocks.length; i++) out.push(story.blocks[i]);
    });
  stories(doc.headers);
  stories(doc.footers);
  if (doc.footnotes) doc.footnotes.items.forEach((note) => out.push(note));
  out.push(null);
  if (doc.endnotes) doc.endnotes.items.forEach((note) => out.push(note));
  out.push(null);
  doc.numbering.abstracts.forEach((value) => out.push(value));
  doc.numbering.nums.forEach((value) => out.push(value));
  return out;
}

const sameList = <T>(a: T[], b: T[]): boolean => {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
};

const sameCounts = (a: number[], b: number[]): boolean => {
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i++) if ((a[i] || 0) !== (b[i] || 0)) return false;
  return true;
};

const sameFlags = (a: boolean[], b: boolean[]): boolean => {
  const length = Math.max(a.length, b.length);
  for (let i = 0; i < length; i++) if (!!a[i] !== !!b[i]) return false;
  return true;
};

const BODY: StoryRef = { id: 'body', main: true };
/** Share of a column footnotes carried over from the page before may take. */
const CARRIED_NOTES_SHARE = 0.75;
/** Height of the default footnote separator, px. */
const SEPARATOR_HEIGHT = 10;
const SEPARATOR_WIDTH = 192;
/** Rows of a table of the body laid out in one step. */
const TABLE_STEP_ROWS = 32;

const now = (): number => Date.now();

export class DocLayout {
  /** The pages laid out so far, complete ones only. */
  readonly pages: LayoutPage[] = [];
  /** Layout stopped at the page limit before the end of the document. */
  truncated = false;

  private readonly doc: DocModel;
  private readonly ctx: LayoutContext;
  private readonly options: DocLayoutOptions;
  private sections: SectionInfo[];
  /** The blocks of the body as they were when they were last looked at: an edit is found by comparing. */
  private blocks: Block[];
  /** See sideParts(). */
  private sides: unknown[];
  /** State at the start of every page, by page index. */
  private marks: (PageMark | undefined)[] = [];
  private old: OldLayout | undefined;
  /** Blocks [from, to) take the place of blocks that gave these events. */
  private expect: { from: number; to: number; events: StateEvent[]; at: number } | undefined;
  private tableSession: PrepSession | undefined;

  private done = false;
  /** Bands waiting to be placed, from `head` on. */
  private queue: QueueItem[] = [];
  private head = 0;
  /** Next block of the body to turn into bands, and the section it is in. */
  private nextBlock = 0;
  /** The table of the body whose rows are being fed, a few at a time. */
  private table: TableLayout | undefined;
  private feedSection = 0;
  private endnotesQueued = false;

  /** Section being laid out. */
  private section = 0;
  private build: PageBuild | undefined;
  private column: Column | undefined;
  private fields: FieldValues = {};
  /** Footnote bands that did not fit where they were referred to. */
  private carry: Band[] = [];
  /** Columns closed in a row without a band placed. */
  private emptyColumns = 0;
  /** The next column starts after a break the document asked for. */
  private afterBreak = true;
  private sectionStarted: boolean[] = [];
  private sectionPages: number[] = [];
  /** Body items that show a total, by page index. */
  private totalItems: { page: number; item: FieldTextItem }[] = [];
  private totalPages: TotalsRecord[] = [];
  private headerCache = new Map<string, { key: string; value: HeaderFooter }>();
  private noteCache = new Map<string, Band[]>();
  private updated: number[] = [];
  /**
   * The bands of the current run of columns on the current page, for
   * balancing them when a continuous section break ends the run; undefined
   * when the run holds something balancing does not handle (footnotes,
   * anchored drawings, breaks).
   */
  private regionBands: Band[] | undefined = [];
  private regionBodyStart = 0;
  private regionTotalsStart = 0;
  private balancing = false;
  private balanceFailed = false;
  /** Pages the caller has seen: those that were there when the current call of layout() began. */
  private handed = 0;

  constructor(doc: DocModel, measurer: TextMeasurer, options: DocLayoutOptions = {}) {
    this.doc = doc;
    this.options = options;
    this.ctx = new LayoutContext(doc, measurer, options);
    this.ctx.layoutBlocks = layoutBlocks;
    this.sections = findSections(doc);
    this.blocks = doc.body.blocks.slice();
    this.sides = sideParts(doc);
    this.setGrid();
  }

  /**
   * The body of the document was edited (blocks replaced by new objects, as
   * the operations in ops/doc do): lays it out again from the page of the
   * first changed block. Pages before it stay as they are; paragraphs that
   * did not change keep their lines. Layout runs here as far as the changed
   * blocks reach; call layout() for the rest, as after the constructor. When
   * the edit left the number of blocks alone and a later page begins exactly
   * as it began before, the pages from there on are kept as well.
   *
   * Returns undefined when the engine cannot follow the edit: it changed what
   * the document counts (the items of a list, the notes), its sections, or
   * something besides the blocks of the body (a header, a note, a list
   * definition). Make a new DocLayout then; this one is of no use any more.
   */
  edited(): EditProgress | undefined {
    const ctx = this.ctx;
    if (this.truncated || ctx.broken || this.options.maxPages !== undefined) return undefined;
    if (!sameList(this.sides, sideParts(this.doc))) {
      ctx.broken = true;
      return undefined;
    }
    const before = this.blocks;
    const blocks = this.doc.body.blocks;
    const shared = Math.min(before.length, blocks.length);
    let first = 0;
    while (first < shared && before[first] === blocks[first]) first++;
    let tail = 0;
    while (tail < shared - first && before[before.length - 1 - tail] === blocks[blocks.length - 1 - tail]) tail++;
    const removed = before.length - first - tail;
    const inserted = blocks.length - first - tail;
    this.updated = [];
    if (removed === 0 && inserted === 0) return { from: this.pages.length, done: this.done, updated: [] };
    this.blocks = blocks.slice();

    const sections = findSections(this.doc);
    if (sections.length !== this.sections.length) return undefined;
    for (let i = 0; i < sections.length; i++) if (sections[i].props !== this.sections[i].props) return undefined;

    // The paragraph before the change looks at what follows it (its borders, its spacing): it is laid out again too.
    let dirty = first;
    for (let i = first - 1, tries = 0; i >= 0 && tries < 16; i--, tries++) {
      const type = blocks[i].type;
      if (type === 'paragraph' || type === 'table') {
        dirty = i;
        break;
      }
    }
    if (!this.done && dirty >= this.nextBlock) {
      // Layout has not come that far yet.
      this.sections = sections;
      return { from: this.pages.length, done: false, updated: [] };
    }
    // Every changed block must have been turned into bands in full: its events are what the new ones get.
    if (first + removed > this.nextBlock - (this.table ? 1 : 0)) return undefined;
    const events: StateEvent[] = [];
    for (let i = first; i < first + removed; i++) {
      const own = ctx.blockEvents.get(before[i]);
      if (!own) return undefined;
      for (let k = 0; k < own.length; k++) events.push(own[k]);
    }

    let from = -1;
    for (let i = Math.min(this.marks.length, this.pages.length + 1) - 1; i >= 0; i--) {
      const mark = this.marks[i];
      if (mark && mark.nextBlock <= dirty) {
        from = i;
        break;
      }
    }
    if (from < 0) return undefined;
    const mark = this.marks[from] as PageMark;

    this.old =
      inserted === removed && from < this.pages.length
        ? {
            from,
            end: first + inserted,
            pages: this.pages.slice(from),
            marks: this.marks.slice(from, this.pages.length),
            live: this.live(),
            itemsBase: mark.totalItems,
            pagesBase: mark.totalPages,
            totalItems: this.totalItems.slice(mark.totalItems),
            totalPages: this.totalPages.slice(mark.totalPages),
          }
        : undefined;
    this.pages.length = from;
    this.marks.length = from;
    this.totalItems.length = mark.totalItems;
    this.totalPages.length = mark.totalPages;
    this.done = false;
    this.queue = mark.queue.slice();
    this.head = 0;
    this.nextBlock = mark.nextBlock;
    this.table = undefined;
    this.tableSession = undefined;
    this.feedSection = mark.feedSection;
    this.endnotesQueued = mark.endnotesQueued;
    this.section = mark.section;
    this.build = undefined;
    this.column = undefined;
    this.carry = mark.carry.slice();
    this.emptyColumns = mark.emptyColumns;
    this.afterBreak = mark.afterBreak;
    this.sectionStarted = mark.sectionStarted.slice();
    this.sectionPages = mark.sectionPages.slice();
    this.balancing = false;
    this.sections = sections;
    this.setGrid();
    this.expect = { from: first, to: first + inserted, events, at: 0 };
    this.handed = from;
    this.openPage(true);
    const end = first + inserted;
    // Bounded like layout(): every step places a band or ends a page.
    while (!this.done && !ctx.broken && (this.expect || this.table || this.nextBlock < end) && this.pages.length < MAX_PAGES) {
      if (!this.step()) break;
    }
    if (ctx.broken || this.expect) return undefined;
    return { from, done: this.done, updated: this.updated };
  }

  private live(): LiveState {
    return {
      done: this.done,
      queue: this.queue,
      head: this.head,
      nextBlock: this.nextBlock,
      table: this.table,
      tableSession: this.tableSession,
      feedSection: this.feedSection,
      endnotesQueued: this.endnotesQueued,
      section: this.section,
      build: this.build,
      column: this.column,
      fields: this.fields,
      carry: this.carry,
      emptyColumns: this.emptyColumns,
      afterBreak: this.afterBreak,
      sectionStarted: this.sectionStarted,
      sectionPages: this.sectionPages,
      regionBands: this.regionBands,
      regionBodyStart: this.regionBodyStart,
      regionTotalsStart: this.regionTotalsStart,
    };
  }

  /**
   * At the start of a page while the document is laid out again after an
   * edit: when everything that decides the rest of the layout is as it was
   * when the old page of the same index began, the old pages from here on are
   * taken back and the engine goes on where it was before the edit.
   */
  private resume(): boolean {
    const old = this.old;
    if (!old) return false;
    const index = this.pages.length;
    const k = index - old.from;
    if (k >= old.pages.length) {
      // More pages than before: nothing of the old layout can come back.
      this.old = undefined;
      return false;
    }
    const mark = k > 0 ? old.marks[k] : undefined;
    if (!mark || this.expect || this.table || this.balancing || this.nextBlock < old.end) return false;
    if (
      mark.nextBlock !== this.nextBlock ||
      mark.feedSection !== this.feedSection ||
      mark.endnotesQueued !== this.endnotesQueued ||
      mark.section !== this.section ||
      mark.afterBreak !== this.afterBreak ||
      mark.emptyColumns !== this.emptyColumns ||
      this.pages[index - 1].number !== old.pages[k - 1].number ||
      mark.queue.length !== this.queue.length - this.head ||
      !sameList(mark.carry, this.carry) ||
      !sameFlags(mark.sectionStarted, this.sectionStarted) ||
      !sameCounts(mark.sectionPages, this.sectionPages)
    ) {
      return false;
    }
    for (let i = 0; i < mark.queue.length; i++) if (mark.queue[i] !== this.queue[this.head + i]) return false;
    // The block that comes next looks at the one before it, which must not be one of the changed ones.
    const blocks = this.doc.body.blocks;
    for (let i = this.nextBlock - 1, tries = 0; i >= 0 && tries < 16; i--, tries++) {
      if (i < old.end) return false;
      const type = blocks[i].type;
      if (type === 'paragraph' || type === 'table') break;
    }

    this.old = undefined;
    const items = this.totalItems.length;
    const pages = this.totalPages.length;
    let goneItems = 0;
    for (const entry of old.totalItems) {
      if (entry.page >= index) this.totalItems.push(entry);
      else goneItems++;
    }
    let gonePages = 0;
    for (const record of old.totalPages) {
      if (record.index >= index) this.totalPages.push(record);
      else gonePages++;
    }
    // The lists are shorter or longer before this page than they were: what counts from their start moves along.
    const shiftItems = items - (old.itemsBase + goneItems);
    const shiftPages = pages - (old.pagesBase + gonePages);
    for (let j = k; j < old.pages.length; j++) {
      this.pages.push(old.pages[j]);
      const kept = old.marks[j];
      if (kept && (shiftItems || shiftPages)) {
        kept.totalItems += shiftItems;
        kept.totalPages += shiftPages;
      }
      this.marks[index + j - k] = kept;
    }
    const live = old.live;
    this.done = live.done;
    this.queue = live.queue;
    this.head = live.head;
    this.nextBlock = live.nextBlock;
    this.table = live.table;
    this.tableSession = live.tableSession;
    this.feedSection = live.feedSection;
    this.endnotesQueued = live.endnotesQueued;
    this.section = live.section;
    this.build = live.build;
    this.column = live.column;
    this.fields = live.fields;
    this.carry = live.carry;
    this.emptyColumns = live.emptyColumns;
    this.afterBreak = live.afterBreak;
    this.sectionStarted = live.sectionStarted;
    this.sectionPages = live.sectionPages;
    this.regionBands = live.regionBands;
    this.regionBodyStart = live.regionBodyStart;
    this.regionTotalsStart = live.regionTotalsStart + shiftItems;
    this.setGrid();
    // The number of pages is what it was: only the pages laid out again have totals to fill in.
    if (this.done) this.fillTotals(old.itemsBase, items, old.pagesBase, pages);
    return true;
  }

  /**
   * Lays out more of the document: to the end without a budget, otherwise
   * until the budget is used up. Call again while `done` is false.
   */
  layout(budget?: LayoutBudget): LayoutProgress {
    this.updated = [];
    this.handed = this.pages.length;
    if (this.done) return { done: true, updated: [] };
    const started = budget && budget.milliseconds !== undefined ? now() : 0;
    const target = budget && budget.pages !== undefined ? this.pages.length + Math.max(1, budget.pages) : Infinity;
    const limit = this.options.maxPages === undefined ? Infinity : this.options.maxPages;
    let steps = 0;
    for (;;) {
      if (this.pages.length >= target || this.pages.length >= limit) break;
      if (this.pages.length >= MAX_PAGES) {
        // Not a document any more: stop for good with what there is.
        this.truncated = true;
        this.build = undefined;
        this.column = undefined;
        this.finish();
        break;
      }
      if (!this.step()) break;
      if (budget && budget.milliseconds !== undefined && (++steps & 15) === 0 && now() - started >= budget.milliseconds) break;
    }
    return { done: this.done, updated: this.updated };
  }

  // ------------------------------------------------------------------ feed

  private setGrid(): void {
    const grid = this.sections[this.section].props.docGrid;
    const snaps = grid && (grid.type === 'lines' || grid.type === 'linesAndChars' || grid.type === 'snapToChars') && grid.linePitch;
    this.ctx.grid = snaps ? twips(grid!.linePitch) : undefined;
  }

  /** Puts the bands of the next block into the queue; false when the document has no more. */
  private fill(): boolean {
    const blocks = this.doc.body.blocks;
    this.queue.length = 0;
    this.head = 0;
    const ctx = this.ctx;
    while (this.table) {
      // More rows of the table that is under way.
      ctx.session = this.tableSession;
      let bands: Band[];
      try {
        bands = this.table.next(TABLE_STEP_ROWS);
      } finally {
        ctx.session = undefined;
      }
      if (this.table.done) {
        this.table = undefined;
        this.tableSession = undefined;
      }
      for (let i = 0; i < bands.length; i++) this.queue.push(bands[i]);
      if (this.queue.length) return true;
    }
    for (;;) {
      const expect = this.expect;
      if (expect && this.nextBlock >= expect.to) {
        // The changed blocks are all laid out: they must have taken every value the old ones had.
        if (expect.at !== expect.events.length) ctx.broken = true;
        this.expect = undefined;
      }
      if (this.nextBlock >= blocks.length) break;
      // The block is counted as fed only once it is: a page that begins while it is prepared begins before it.
      const index = this.nextBlock;
      const block = blocks[index];
      // Headers and footers of a page that begins here are not part of the block.
      this.needColumn();
      const known = ctx.blockEvents.get(block);
      const cursor = expect && index >= expect.from && index < expect.to ? expect : undefined;
      let session: PrepSession | undefined;
      if (!known) session = { record: [], seen: new Set(), expect: cursor };
      else if (cursor) {
        // A block that was laid out before (an edit undone): its values must be the ones expected here.
        ctx.session = { record: [], seen: new Set(), expect: cursor };
        ctx.replay(known);
      }
      ctx.session = session;
      let bands: Band[];
      try {
        bands = this.blockBands(blocks, index);
      } finally {
        ctx.session = undefined;
      }
      if (session) {
        ctx.blockEvents.set(block, session.record);
        if (this.table) this.tableSession = session;
      }
      this.nextBlock = index + 1;
      for (let i = 0; i < bands.length; i++) this.queue.push(bands[i]);
      if (index + 1 === this.sections[this.feedSection].end && this.feedSection < this.sections.length - 1) {
        this.feedSection++;
        this.queue.push({ mark: true, section: this.feedSection });
      }
      if (this.queue.length) return true;
    }
    if (!this.endnotesQueued) {
      this.queueEndnotes();
      this.endnotesQueued = true;
      if (this.queue.length) return true;
    }
    return false;
  }

  private blockBands(blocks: Block[], index: number): Band[] {
    const block = blocks[index];
    const column = this.needColumn();
    if (block.type === 'paragraph') {
      const prep = getPrep(this.ctx, block, BODY, [index], undefined, 0, this.fields);
      if (isHiddenParagraph(prep)) return [];
      const env = paragraphEnv(this.ctx, blocks, index, prep, undefined, this.fields);
      const dropBefore = this.dropsBefore(column);
      const wrap = this.wrapEnv(column, prep.floats, dropBefore ? 0 : env.before, dropBefore);
      return paragraphBands(this.ctx, prep, column.width, env, wrap);
    }
    if (block.type === 'table') {
      const floating = this.floatingTable(block, index, column);
      if (floating) return floating;
      // A long table must not hold up the pages before it: its rows are laid out as they are placed.
      const table = tableLayout(this.ctx, block, column.width, BODY, [index], 0, this.fields);
      const bands = table.next(TABLE_STEP_ROWS);
      if (!table.done) this.table = table;
      return bands;
    }
    return [];
  }

  /**
   * A table positioned on the page (not relative to the text around it) that
   * fits on one page is placed like an anchored drawing: where it says, with
   * the text keeping clear of it. Returns the band that carries it, undefined
   * for every other table. Floating tables positioned relative to the text
   * stay in the flow, where they can go from page to page.
   */
  private floatingTable(table: Table, index: number, column: Column): Band[] | undefined {
    const position = this.ctx.tableProps(table).floating;
    if (!position || (position.vertAnchor !== 'page' && position.vertAnchor !== 'margin')) return undefined;
    const build = this.build!;
    const geometry = build.geometry;
    const box = { left: 0, width: 0 };
    const bands = tableBands(this.ctx, table, geometry.right - geometry.left, BODY, [index], 0, this.fields, box);
    const height = bandsHeight(bands);
    if (!bands.length || height > geometry.bottom - geometry.top) return undefined;
    const items: PageItem[] = [];
    let notes: number[] | undefined;
    let y = 0;
    for (const band of bands) {
      for (let i = 0; i < band.items.length; i++) items.push(shifted(band.items[i], -box.left, y));
      if (band.notes) for (const note of band.notes) (notes || (notes = [])).push(note);
      y += band.height;
    }
    const emu = (value: number | undefined) => Math.round((value || 0) * 635);
    const horizontal: DrawingPosition = { relativeFrom: position.horzAnchor === 'page' ? 'page' : position.horzAnchor === 'margin' ? 'margin' : 'column' };
    if (position.xAlign !== undefined) horizontal.align = position.xAlign;
    else horizontal.offset = emu(position.x);
    const vertical: DrawingPosition = { relativeFrom: position.vertAnchor };
    if (position.yAlign !== undefined && position.yAlign !== 'inline') vertical.align = position.yAlign;
    else vertical.offset = emu(position.y);
    const anchor: DrawingAnchor = {
      horizontal,
      vertical,
      wrap: 'square',
      distT: emu(position.topFromText),
      distB: emu(position.bottomFromText),
      distL: emu(position.leftFromText),
      distR: emu(position.rightFromText),
    };
    const float: FloatSpec = { width: box.width, height, items, anchor, dx: 0, dy: 0, columnWidth: column.width, inCell: false };
    // The table stays with the paragraph after it, which is what anchors it.
    const band: Band = { height: 0, before: 0, after: 0, items: [], keep: true, floats: [float] };
    if (notes) band.notes = notes;
    return [band];
  }

  /** Endnotes go after the last text of the document, under a separator. */
  private queueEndnotes(): void {
    const endnotes = this.doc.endnotes;
    const order = this.ctx.endnoteOrder;
    if (!endnotes || !order.length) return;
    const column = this.needColumn();
    const y = SEPARATOR_HEIGHT / 2;
    this.queue.push({
      height: SEPARATOR_HEIGHT,
      before: 0,
      after: 0,
      keep: true,
      items: [{ type: 'line', x1: 0, y1: y, x2: Math.min(SEPARATOR_WIDTH, column.width), y2: y, color: '#000000', width: 0.75, style: 'single' }],
    });
    for (const id of order) {
      const note = endnotes.items.get(id);
      if (!note) continue;
      const story: StoryRef = { id: 'endnote:' + id, main: false, noteLabel: noteLabel(this.ctx, 'endnote', this.ctx.endnoteNumbers.get(id) || 1) };
      const bands = layoutBlocks(this.ctx, note.blocks, column.width, story, [], undefined, 0, undefined);
      for (let i = 0; i < bands.length; i++) this.queue.push(bands[i]);
    }
  }

  // ------------------------------------------------------------------ steps

  /** Places one band or follows one section mark; false when the document is done. */
  private step(): boolean {
    if (this.done) return false;
    if (this.head >= this.queue.length && !this.fill()) {
      this.finish();
      return false;
    }
    const item = this.queue[this.head];
    if (isMark(item)) {
      this.head++;
      this.changeSection(item.section);
    } else this.place(item);
    return true;
  }

  private finish(): void {
    if (this.build) {
      this.closeColumn();
      this.closePage();
    } else if (!this.pages.length && !this.truncated) {
      // Not even an empty paragraph: still a page.
      this.needColumn();
      this.closeColumn();
      this.closePage();
    }
    // Lines of footnotes that found no room on the pages of the text take pages of their own: nothing of a note is dropped.
    while (this.carry.length && !this.truncated && this.pages.length < MAX_PAGES) {
      const left = this.carry.length;
      this.openPage(true);
      this.closeColumn();
      this.closePage();
      if (this.carry.length >= left) break;
    }
    this.done = true;
    this.fillTotals();
  }

  private needColumn(): Column {
    if (!this.build) this.openPage(true);
    return this.column!;
  }

  /** Is the space before the first band of the column dropped. */
  private dropsBefore(column: Column): boolean {
    return column.placed.length === 0 && column.auto;
  }

  private place(band: Band): void {
    const column = this.needColumn();
    const build = this.build!;
    if (band.para) band = this.revalidate(band, column);

    if (band.breakBefore && build.used) {
      this.afterBreak = true;
      this.closeColumn();
      this.closePage();
      this.openPage();
      return;
    }

    // A table that goes on in a new column starts with its header rows again.
    if (band.row && !band.row.header && column.placed.length === 0 && band.row.table.headers.length) {
      const headers = band.row.table.headers;
      if (bandsHeight(headers) < (column.bottom - column.y) / 2) {
        for (const header of headers) {
          const copy: Band = { height: header.height, before: 0, after: 0, keep: true, items: header.items.map(asRepeat) };
          this.put(column, copy, column.y, true);
        }
      }
    }

    const real = column.real;
    const before = real === 0 && column.auto ? 0 : band.before;
    let y = column.y + before;
    if (!band.para || !band.para.slot) y = this.clearOf(column, y, band.height);
    let need = 0;
    if (band.notes && band.notes.length) {
      const first = this.noteBands(band.notes[0], column.width)[0];
      need = (column.notes.length === 0 ? SEPARATOR_HEIGHT : 0) + (first ? first.before + first.height : 0);
    }
    const limit = column.bottom - column.notesHeight - need;
    if (y + band.height + (band.overhang || 0) <= limit + EPSILON) {
      this.put(column, band, y);
      this.head++;
      this.afterPlaced(band);
      return;
    }

    // It does not fit. An empty column at the top of a page takes it anyway, divided if it can be: no other
    // column would do better. One that footnotes from the page before take room in is given up once first.
    const stuck = real === 0 && column.full && (column.notes.length === 0 || this.emptyColumns >= 2);
    if (this.balancing && stuck) {
      // Columns being balanced are not made to hold what does not fit: they get more height instead.
      this.balanceFailed = true;
      return;
    }
    const parts = band.split ? band.split(limit - y) : undefined;
    if (parts) {
      this.put(column, parts[0], y);
      this.queue[this.head] = parts[1];
      this.nextColumn();
      return;
    }
    if (stuck) {
      // What kept it from fitting may be a drawing that leaves no room on the page: then the drawing is not kept clear of.
      if (band.para && band.para.slot) band = this.relayout(band, column, undefined);
      this.put(column, band, column.y + before);
      this.head++;
      this.afterPlaced(band);
      return;
    }
    if (real > 0) {
      // Take back the bands that want to stay with this one, as long as something stays in the column.
      const placed = column.placed;
      let from = placed.length;
      while (from > 0 && placed[from - 1].band.keep) from--;
      let kept = 0;
      for (let i = 0; i < from; i++) if (!placed[i].repeated) kept++;
      if (kept > 0 && from < placed.length) {
        const first = placed[from];
        const back: QueueItem[] = [];
        for (let i = from; i < placed.length; i++) if (!placed[i].repeated) back.push(placed[i].band);
        placed.length = from;
        column.real = kept;
        column.y = first.columnY;
        column.notes.length = first.notes;
        column.notesHeight = first.notesHeight;
        this.carry.length = first.carried;
        build.exclusions.length = first.exclusions;
        replaceIn(this.queue, this.head, 0, back);
      } else if (this.balancing && from < placed.length) {
        // Nor are bands that want to stay together parted for the balance.
        this.balanceFailed = true;
        return;
      }
    }
    this.nextColumn();
  }

  private afterPlaced(band: Band): void {
    if (band.breakAfter === 'page') {
      this.afterBreak = true;
      this.closeColumn();
      this.closePage();
      this.openPage();
    } else if (band.breakAfter === 'column') {
      this.afterBreak = true;
      this.nextColumn();
    }
  }

  /** Puts a band into the column at y (the top of the band, below its space before). */
  private put(column: Column, band: Band, y: number, repeated = false): void {
    const build = this.build!;
    build.used = true;
    const placed: Placed = {
      band,
      y,
      columnY: column.y,
      notes: column.notes.length,
      notesHeight: column.notesHeight,
      carried: this.carry.length,
      exclusions: build.exclusions.length,
    };
    if (repeated) placed.repeated = true;
    else column.real++;
    column.placed.push(placed);
    column.y = y + band.height + band.after;
    this.emptyColumns = 0;
    if (band.floats) {
      placed.floats = [];
      for (const float of band.floats) {
        const position = floatPosition(float, this.floatFrame(build.geometry, column.x + float.dx, float.columnWidth, y + float.dy));
        placed.floats.push({ float, x: position.x, y: position.y });
        const exclusion = floatExclusion(float, position.x, position.y);
        if (exclusion) build.exclusions.push(exclusion);
      }
    }
    if (band.notes) {
      for (let n = 0; n < band.notes.length; n++) {
        const bands = this.noteBands(band.notes[n], column.width);
        let k = 0;
        for (; k < bands.length; k++) {
          const note = bands[k];
          const separator = column.notes.length === 0 ? SEPARATOR_HEIGHT : 0;
          // The first line of the first note stays with its reference (room was kept for it); the rest goes where there is room.
          if (this.carry.length) break;
          if ((n > 0 || k > 0) && column.y + note.before + note.height > column.bottom - column.notesHeight - separator + EPSILON) break;
          column.notes.push(note);
          column.notesHeight += note.before + note.height + note.after + separator;
        }
        for (; k < bands.length; k++) this.carry.push(bands[k]);
      }
    }
  }

  private floatFrame(geometry: PageGeometry, columnLeft: number, columnWidth: number, paragraphTop: number): FloatFrame {
    return {
      pageWidth: geometry.width,
      pageHeight: geometry.height,
      left: geometry.left,
      right: geometry.right,
      top: geometry.top,
      bottom: geometry.bottom,
      columnLeft,
      columnWidth,
      paragraphTop,
    };
  }

  // ------------------------------------------------------- drawings in the way

  private blockers(column: Column, top: number, bottom: number, extra: Exclusion[] | undefined): Exclusion[] | undefined {
    let out: Exclusion[] | undefined;
    const right = column.x + column.width;
    const scan = (list: Exclusion[]) => {
      for (let i = 0; i < list.length; i++) {
        const e = list[i];
        if (e.top < bottom - EPSILON && e.bottom > top + EPSILON && e.left < right && e.right > column.x) (out || (out = [])).push(e);
      }
    };
    scan(this.build!.exclusions);
    if (extra) scan(extra);
    return out;
  }

  /** Lowest y at or below `y` where a band of the height is clear of every drawing text wraps around. */
  private clearOf(column: Column, y: number, height: number): number {
    if (!this.build!.exclusions.length) return y;
    for (let tries = 0; tries < 64; tries++) {
      const hit = this.blockers(column, y, y + Math.max(height, 1), undefined);
      if (!hit) break;
      let down = y;
      for (const e of hit) if (e.bottom > down) down = e.bottom;
      if (down <= y) break;
      y = down;
      // Past the foot of the column: the band goes to the next one.
      if (y > column.bottom) break;
    }
    return y;
  }

  /**
   * What a paragraph that starts at the current place of the column has to
   * keep clear of; undefined when nothing is in its way. `floats` are the
   * anchored drawings of the paragraph itself, which are not on the page yet.
   */
  private wrapEnv(column: Column, floats: FloatSpec[] | undefined, before: number, dropBefore: boolean): WrapEnv | undefined {
    const build = this.build!;
    let own: Exclusion[] | undefined;
    if (floats) {
      for (const float of floats) {
        const position = floatPosition(float, this.floatFrame(build.geometry, column.x, column.width, column.y + before));
        const exclusion = floatExclusion(float, position.x, position.y);
        if (exclusion) (own || (own = [])).push(exclusion);
      }
    }
    const y = column.y;
    let active = own !== undefined;
    if (!active) {
      const right = column.x + column.width;
      for (const e of build.exclusions) {
        if (e.bottom > y + EPSILON && e.left < right && e.right > column.x) {
          active = true;
          break;
        }
      }
    }
    if (!active) return undefined;
    return {
      page: this.pages.length,
      column: column.index,
      y,
      dropBefore,
      free: (top, bottom) => {
        const hit = this.blockers(column, top, bottom, own);
        if (!hit) return [{ left: 0, right: column.width }];
        let down = Infinity;
        for (const e of hit) if (e.bottom < down) down = e.bottom;
        for (const e of hit) if (e.whole) return down;
        // What the drawings leave of the column: each takes its own width away, and with it the side text may not be on.
        let gaps = [{ left: 0, right: column.width }];
        for (const e of hit) {
          const left = e.left - column.x;
          const right = e.right - column.x;
          const side = e.side === 'largest' ? (left >= column.width - right ? 'left' : 'right') : e.side;
          const from = side === 'right' ? -Infinity : left;
          const to = side === 'left' ? Infinity : right;
          const next: { left: number; right: number }[] = [];
          for (const gap of gaps) {
            if (to <= gap.left || from >= gap.right) next.push(gap);
            else {
              if (from > gap.left) next.push({ left: gap.left, right: from });
              if (to < gap.right) next.push({ left: to, right: gap.right });
            }
          }
          gaps = next;
        }
        // Text is not squeezed into a sliver.
        const narrowest = Math.min(MIN_WRAP_WIDTH, column.width / 2);
        gaps = gaps.filter((gap) => gap.right - gap.left >= narrowest);
        return gaps.length ? gaps : down;
      },
    };
  }

  /**
   * A band of a paragraph is only good for the width it was broken for and,
   * when it was broken around drawings, for the place it was meant for. When
   * the band at the head of the queue is not, the rest of its paragraph is
   * laid out again for where it is now.
   */
  private revalidate(band: Band, column: Column): Band {
    const ref = band.para!;
    const whole = ref.atom === 0 && ref.ch === 0;
    const dropBefore = this.dropsBefore(column);
    const wrap = this.wrapEnv(column, whole ? ref.prep.floats : undefined, dropBefore || !whole ? 0 : ref.env.before, dropBefore);
    let good = Math.abs(ref.width - column.width) < EPSILON;
    if (good) {
      if (ref.slot) good = ref.slot.page === this.pages.length && ref.slot.column === column.index && Math.abs(ref.slot.y - column.y) < EPSILON && (band.before === 0 || !dropBefore);
      else good = !wrap;
    }
    if (good) return band;
    return this.relayout(band, column, wrap);
  }

  /** Lays out the paragraph of the band at the head of the queue again from that band on, and returns the new head. */
  private relayout(band: Band, column: Column, wrap: WrapEnv | undefined): Band {
    const ref = band.para!;
    let count = 1;
    while (this.head + count < this.queue.length) {
      const next = this.queue[this.head + count];
      if (isMark(next) || !next.para || next.para.prep !== ref.prep) break;
      count++;
    }
    const env = ref.env.fields ? { ...ref.env, fields: this.fields } : ref.env;
    const bands = paragraphBands(this.ctx, ref.prep, column.width, env, wrap, { atom: ref.atom, ch: ref.ch });
    replaceIn(this.queue, this.head, count, bands);
    return bands[0];
  }

  // -------------------------------------------------------------- footnotes

  private noteBands(id: number, width: number): Band[] {
    const key = id + '|' + width;
    let bands = this.noteCache.get(key);
    if (!bands) {
      const note = this.doc.footnotes && this.doc.footnotes.items.get(id);
      const story: StoryRef = { id: 'footnote:' + id, main: false, noteLabel: noteLabel(this.ctx, 'footnote', this.ctx.footnoteNumbers.get(id) || 1) };
      bands = note ? layoutBlocks(this.ctx, note.blocks, width, story, [], undefined, 0, undefined) : [];
      if (this.noteCache.size > 2000) this.noteCache.clear();
      this.noteCache.set(key, bands);
    }
    return bands;
  }

  // ---------------------------------------------------------------- columns

  private openColumn(index: number): void {
    const build = this.build!;
    const spec = build.columns[index];
    const column: Column = {
      index,
      x: spec.x,
      width: spec.width,
      top: build.regionTop,
      bottom: build.bottom,
      y: build.regionTop,
      placed: [],
      real: 0,
      notes: [],
      notesHeight: 0,
      auto: !this.afterBreak,
      full: build.regionTop <= build.top + EPSILON,
    };
    this.afterBreak = false;
    this.column = column;
    if (this.carry.length) {
      // Notes left over from the column before come first.
      const room = (column.bottom - column.top) * CARRIED_NOTES_SHARE;
      let k = 0;
      for (; k < this.carry.length; k++) {
        const note = this.carry[k];
        const height = note.before + note.height + note.after;
        if (k > 0 && column.notesHeight + height > room) break;
        if (k === 0) column.notesHeight += SEPARATOR_HEIGHT;
        column.notes.push(note);
        column.notesHeight += height;
      }
      this.carry = this.carry.slice(k);
    }
  }

  /** Moves the bands of the column onto the page. */
  private closeColumn(): void {
    const column = this.column;
    const build = this.build;
    if (!column || !build) return;
    this.column = undefined;
    const placed = column.placed;
    const out = build.body;
    const pageIndex = this.pages.length;
    for (let i = 0; i < placed.length; i++) {
      const p = placed[i];
      if (this.regionBands && !p.repeated) {
        const band = p.band;
        if (band.notes || band.floats || band.breakAfter || band.breakBefore || (band.para && band.para.slot)) this.regionBands = undefined;
        else this.regionBands.push(band);
      }
      this.emit(p.band.items, column.x, p.y, out, pageIndex, build);
      if (p.floats) {
        for (const f of p.floats) this.emit(f.float.items, f.x, f.y, f.float.anchor.behindDoc ? build.behind : build.front, pageIndex, build);
      }
    }
    if (placed.length) {
      // A table that goes on in the next column is closed with a line here.
      const last = placed[placed.length - 1];
      const next = this.head < this.queue.length ? this.queue[this.head] : undefined;
      const row = last.band.row;
      if (row && row.close && next && !isMark(next) && next.row && next.row.table === row.table) this.emit(row.close, column.x, last.y, out, pageIndex, build);
    }
    if (column.notes.length) {
      this.regionBands = undefined;
      let y = column.bottom - column.notesHeight;
      if (y < build.notesTop) build.notesTop = y;
      const line = y + SEPARATOR_HEIGHT / 2;
      out.push({ type: 'line', x1: column.x, y1: line, x2: column.x + Math.min(SEPARATOR_WIDTH, column.width), y2: line, color: '#000000', width: 0.75, style: 'single' });
      y += SEPARATOR_HEIGHT;
      for (const note of column.notes) {
        y += note.before;
        this.emit(note.items, column.x, y, out, pageIndex, build);
        y += note.height + note.after;
      }
    }
    if (column.y > build.regionEnd) build.regionEnd = column.y;
    if (!column.real) this.emptyColumns++;
  }

  /** Copies items onto the page, moved to their place; fields that show the page number get it here. */
  private emit(items: PageItem[], dx: number, dy: number, out: PageItem[], pageIndex: number, build: PageBuild): void {
    for (let i = 0; i < items.length; i++) {
      const item = shifted(items[i], dx, dy);
      out.push(item);
      if (item.type !== 'text') continue;
      const field = (item as FieldTextItem).field;
      if (field === undefined) continue;
      const text = item as FieldTextItem;
      if (field === 'PAGE') {
        const value = formatPageNumber(build.page.number, text.fieldFormat !== undefined ? text.fieldFormat : this.fields.pageFormat);
        if (value !== text.text) this.setText(text, value);
      } else this.totalItems.push({ page: pageIndex, item: text });
    }
  }

  private setText(item: FieldTextItem, text: string): void {
    item.text = text;
    item.width = this.ctx.width(item.font, text) * (item.scale || 1) + (item.letterSpacing || 0) * text.length;
  }

  private nextColumn(): void {
    const build = this.build!;
    const index = this.column ? this.column.index : 0;
    this.closeColumn();
    if (index + 1 < build.columns.length) this.openColumn(index + 1);
    else if (this.balancing) this.balanceFailed = true;
    else {
      this.closePage();
      this.openPage();
    }
  }

  // ------------------------------------------------------------------ pages

  /**
   * Starts a page of the current section. `direct`: the caller goes on with
   * the page at once, so the old pages are not looked at here (see resume()).
   */
  private openPage(direct = false): void {
    if (!direct && this.old && this.resume()) return;
    this.marks[this.pages.length] =
      this.table || this.balancing
        ? undefined
        : {
            nextBlock: this.nextBlock,
            feedSection: this.feedSection,
            endnotesQueued: this.endnotesQueued,
            section: this.section,
            afterBreak: this.afterBreak,
            emptyColumns: this.emptyColumns,
            queue: this.queue.slice(this.head),
            carry: this.carry.slice(),
            sectionStarted: this.sectionStarted.slice(),
            sectionPages: this.sectionPages.slice(),
            totalItems: this.totalItems.length,
            totalPages: this.totalPages.length,
          };
    const index = this.section;
    const section = this.sections[index];
    const props = section.props;
    const firstOfSection = !this.sectionStarted[index];
    this.sectionStarted[index] = true;

    const previous = this.pages.length ? this.pages[this.pages.length - 1].number : 0;
    const numbering = props.pageNumbering;
    let number = previous + 1;
    if (firstOfSection && numbering && numbering.start !== undefined) number = numbering.start;
    const even = number % 2 === 0;
    const mirrored = this.doc.settings.mirrorMargins && even;
    const geometry = pageGeometry(props, mirrored);
    const page: LayoutPage = { width: geometry.width, height: geometry.height, section: index, number, items: [] };
    this.fields = { page: number };
    if (numbering && numbering.format !== undefined) this.fields.pageFormat = numbering.format;

    const build: PageBuild = {
      page,
      geometry,
      top: geometry.top,
      bottom: geometry.bottom,
      columns: geometry.columns,
      separator: geometry.separator,
      regionTop: geometry.top,
      regionEnd: geometry.top,
      notesTop: Infinity,
      exclusions: [],
      borders: [],
      behind: [],
      header: [],
      body: [],
      footer: [],
      front: [],
      headerBehind: [],
      headerFront: [],
      used: false,
      mirrored,
      totals: false,
      firstOfSection,
    };
    this.build = build;
    pageBorderItems(this.ctx, props, geometry, firstOfSection, build.borders);

    const kind: HeaderFooterKind = firstOfSection && props.titlePage ? 'first' : this.doc.settings.evenAndOddHeaders && even ? 'even' : 'default';
    const headerId = inheritedPart(this.sections, index, 'headers', kind);
    const footerId = inheritedPart(this.sections, index, 'footers', kind);
    if (headerId !== undefined) build.headerKey = { id: headerId, kind: 'headers' };
    if (footerId !== undefined) build.footerKey = { id: footerId, kind: 'footers' };
    this.layoutHeaderFooter(build, this.fields);

    build.regionTop = build.top;
    build.regionEnd = build.top;
    this.startRegion(build);
    this.openColumn(0);
  }

  private startRegion(build: PageBuild): void {
    this.regionBands = [];
    this.regionBodyStart = build.body.length;
    this.regionTotalsStart = this.totalItems.length;
  }

  /** Lays out the header and footer of a page and fits the room for body text between them. */
  private layoutHeaderFooter(build: PageBuild, values: FieldValues): void {
    const geometry = build.geometry;
    build.header = [];
    build.footer = [];
    const behind: PageItem[] = [];
    const front: PageItem[] = [];
    let top = geometry.top;
    let bottom = geometry.bottom;
    let totals = false;
    const header = build.headerKey ? this.headerFooter(build.headerKey.id, 'headers', geometry, values) : undefined;
    if (header) {
      this.stack(header.bands, geometry, geometry.headerTop, build.header, behind, front, values);
      if (!geometry.fixedTop) top = Math.max(top, geometry.headerTop + header.height);
      totals = totals || header.totals;
    }
    const footer = build.footerKey ? this.headerFooter(build.footerKey.id, 'footers', geometry, values) : undefined;
    if (footer) {
      const y = geometry.footerBottom - footer.height;
      this.stack(footer.bands, geometry, y, build.footer, behind, front, values);
      if (!geometry.fixedBottom) bottom = Math.min(bottom, y);
      totals = totals || footer.totals;
    }
    // A header and footer that leave no room for text do not get to: the margins decide then.
    if (bottom - top < MIN_TEXT_EXTENT) {
      top = geometry.top;
      bottom = geometry.bottom;
    }
    build.totals = totals;
    build.top = top;
    build.bottom = bottom;
    build.headerBehind = behind;
    build.headerFront = front;
  }

  /** Stacks the bands of a header or footer from y down; their anchored drawings go behind or in front. */
  private stack(bands: Band[], geometry: PageGeometry, y: number, out: PageItem[], behind: PageItem[], front: PageItem[], values: FieldValues): void {
    const x = geometry.left;
    // Text boxes are laid out once, with their paragraph: fields in them get their value here.
    const add = (item: PageItem, dx: number, dy: number, target: PageItem[]) => {
      const copy = shifted(item, dx, dy);
      target.push(copy);
      if (copy.type !== 'text') return;
      const text = copy as FieldTextItem;
      if (text.field === undefined) return;
      const value = text.field === 'PAGE' ? values.page : text.field === 'NUMPAGES' ? values.numPages : values.sectionPages;
      if (value === undefined) return;
      const shown = formatPageNumber(value, text.fieldFormat !== undefined ? text.fieldFormat : text.field === 'PAGE' ? values.pageFormat : undefined);
      if (shown !== text.text) this.setText(text, shown);
    };
    for (const band of bands) {
      y += band.before;
      for (let i = 0; i < band.items.length; i++) add(band.items[i], x, y, out);
      if (band.floats) {
        for (const float of band.floats) {
          const position = floatPosition(float, this.floatFrame(geometry, x + float.dx, float.columnWidth, y + float.dy));
          const target = float.anchor.behindDoc ? behind : front;
          for (let i = 0; i < float.items.length; i++) add(float.items[i], position.x, position.y, target);
        }
      }
      y += band.height + band.after;
    }
  }

  private headerFooter(id: string, kind: 'headers' | 'footers', geometry: PageGeometry, values: FieldValues): HeaderFooter | undefined {
    const story: Story | undefined = (kind === 'headers' ? this.doc.headers : this.doc.footers).get(id);
    if (!story) return undefined;
    const width = geometry.right - geometry.left;
    const slot = kind + '\u0001' + id + '\u0001' + width;
    const cached = this.headerCache.get(slot);
    // A header without fields is the same on every page; one with fields is the same for the same values.
    const key = (values.page === undefined ? '' : values.page) + '|' + (values.pageFormat || '') + '|' + (values.numPages === undefined ? '' : values.numPages) + '|' + (values.sectionPages === undefined ? '' : values.sectionPages);
    if (cached && (cached.key === '' || cached.key === key)) return cached.value;
    const ref: StoryRef = { id: (kind === 'headers' ? 'header:' : 'footer:') + id, main: false };
    const bands = layoutBlocks(this.ctx, story.blocks, width, ref, [], undefined, 0, values);
    let live = false;
    let totals = false;
    const look = (items: PageItem[]) => {
      for (const item of items) {
        const field = item.type === 'text' ? (item as FieldTextItem).field : undefined;
        if (field === undefined) continue;
        live = true;
        if (field !== 'PAGE') totals = true;
      }
    };
    for (const band of bands) {
      look(band.items);
      if (band.floats) for (const float of band.floats) look(float.items);
    }
    const value: HeaderFooter = { bands, height: bandsHeight(bands), totals };
    this.headerCache.set(slot, { key: live ? key : '', value });
    return value;
  }

  /** Puts the finished page together and hands it out. */
  private closePage(): void {
    const build = this.build;
    if (!build) return;
    this.build = undefined;
    this.column = undefined;
    this.separators(build, build.regionEnd);
    build.page.items = this.compose(build);
    const index = this.pages.length;
    this.pages.push(build.page);
    this.sectionPages[build.page.section] = (this.sectionPages[build.page.section] || 0) + 1;
    if (build.totals) this.totalPages.push({ index, build });
  }

  private compose(build: PageBuild): PageItem[] {
    return build.borders.concat(build.headerBehind, build.behind, build.header, build.body, build.footer, build.headerFront, build.front);
  }

  /** The totals are known: the fields that show them get their text. */
  private fillTotals(itemFrom = 0, itemTo = this.totalItems.length, pageFrom = 0, pageTo = this.totalPages.length): void {
    const total = this.pages.length;
    const changed = new Set<number>();
    for (let i = itemFrom; i < itemTo; i++) {
      const entry = this.totalItems[i];
      const page = this.pages[entry.page];
      if (!page) continue;
      const item = entry.item;
      const value = item.field === 'NUMPAGES' ? total : this.sectionPages[page.section] || 1;
      const text = formatPageNumber(value, item.fieldFormat);
      if (text === item.text) continue;
      this.setText(item, text);
      changed.add(entry.page);
    }
    for (let i = pageFrom; i < pageTo; i++) {
      const record = this.totalPages[i];
      const page = this.pages[record.index];
      const build = record.build;
      const values: FieldValues = { page: page.number, numPages: total, sectionPages: this.sectionPages[page.section] || 1 };
      const numbering = this.sections[page.section].props.pageNumbering;
      if (numbering && numbering.format !== undefined) values.pageFormat = numbering.format;
      // The room for body text stays as it was: the text is where it is.
      const top = build.top;
      const bottom = build.bottom;
      this.layoutHeaderFooter(build, values);
      build.top = top;
      build.bottom = bottom;
      page.items = this.compose(build);
      changed.add(record.index);
    }
    // The lists stay: after an edit the totals may be different, and the pages before it show them too.
    changed.forEach((index) => {
      if (index < this.handed) this.updated.push(index);
    });
    this.updated.sort((a, b) => a - b);
  }

  // --------------------------------------------------------------- sections

  private changeSection(next: number): void {
    const from = this.sections[this.section].props;
    const to = this.sections[next].props;
    const type = to.type || 'nextPage';
    const build = this.build;
    const samePaper =
      build !== undefined &&
      (to.page ? to.page.width : undefined) === (from.page ? from.page.width : undefined) &&
      (to.page ? to.page.height : undefined) === (from.page ? from.page.height : undefined);
    if (build && type === 'continuous' && samePaper) {
      // The new section goes on below what the old one put on the page, in its own columns.
      this.closeColumn();
      this.balance(build);
      this.section = next;
      this.setGrid();
      this.continueOnPage(this.build!, to);
      return;
    }
    this.section = next;
    this.setGrid();
    if (!build) return;
    this.afterBreak = true;
    this.closeColumn();
    this.closePage();
    if (type === 'evenPage' || type === 'oddPage') {
      // A page of the wrong parity is left empty.
      // UNSURE: a section that restarts its page numbers (pgNumType start) gets no empty page here, whatever the
      // parity of the number it starts with; whether Word inserts one then was not checked against a document.
      const numbering = to.pageNumbering;
      const previous = this.pages.length ? this.pages[this.pages.length - 1].number : 0;
      const number = numbering && numbering.start !== undefined ? numbering.start : previous + 1;
      if ((number % 2 === 0) !== (type === 'evenPage') && !(numbering && numbering.start !== undefined)) {
        const geometry = pageGeometry(to, false);
        this.pages.push({ width: geometry.width, height: geometry.height, section: next, number, items: [] });
        this.sectionPages[next] = (this.sectionPages[next] || 0) + 1;
      }
    }
    this.openPage();
  }

  /**
   * Columns that a continuous section break ends are made equally high, as
   * far as their content allows: the bands of the columns are placed again
   * into columns of a lower height, the lowest that holds them all.
   */
  private balance(build: PageBuild): void {
    const bands = this.regionBands;
    const count = build.columns.length;
    if (!bands || count < 2 || bands.length < 2 || this.carry.length || build.exclusions.length) return;
    let total = 0;
    for (const band of bands) total += band.before + band.height + band.after;
    const full = build.bottom - build.regionTop;
    const bottom = build.bottom;
    let height = total / count;
    if (!(height < full - 1)) return;
    // The step by which the height is raised until everything fits: a small part of a line.
    const step = Math.max(2, (full - height) / 64);
    const end: SectionMark = { mark: true, section: -1 };
    this.balancing = true;
    for (let tries = 0; ; tries++) {
      const last = height >= full || tries >= 80;
      // The last try is the layout there was before: columns of the full height, and pages after them if need be.
      if (last) this.balancing = false;
      build.body.length = this.regionBodyStart;
      this.totalItems.length = this.regionTotalsStart;
      build.regionEnd = build.regionTop;
      build.bottom = last ? bottom : build.regionTop + height;
      replaceIn(this.queue, this.head, 0, bands);
      this.queue.splice(this.head + bands.length, 0, end);
      this.balanceFailed = false;
      this.afterBreak = true;
      this.emptyColumns = 0;
      this.regionBands = undefined;
      this.openColumn(0);
      let guard = bands.length * 50 + 1000;
      while (this.queue[this.head] !== end && !this.balanceFailed && guard-- > 0) this.place(this.queue[this.head] as Band);
      if (this.queue[this.head] === end && !this.balanceFailed) {
        this.queue.splice(this.head, 1);
        this.closeColumn();
        break;
      }
      // Too low: forget this try.
      this.column = undefined;
      const at = this.queue.indexOf(end, this.head);
      this.queue.splice(this.head, at - this.head + 1);
      if (last) break;
      height += step;
    }
    this.balancing = false;
    if (this.build === build) build.bottom = bottom;
  }

  /** Starts the columns of a section below what is on the page already. */
  private continueOnPage(build: PageBuild, to: SectionProps): void {
    this.closeColumn();
    const top = build.regionEnd;
    // Footnotes of the columns above stay where they are: the new columns end over them.
    if (build.notesTop < build.bottom) build.bottom = build.notesTop;
    this.separators(build, top);
    this.afterBreak = true;
    if (top < build.bottom - EPSILON) {
      // The page keeps its paper, header and footer; the left and right margins and the columns are those of the new section.
      const own = pageGeometry(to, build.mirrored);
      build.geometry = { ...build.geometry, left: own.left, right: own.right, columns: own.columns, separator: own.separator };
      build.columns = own.columns;
      build.separator = own.separator;
      build.regionTop = top;
      build.regionEnd = top;
      this.startRegion(build);
      this.openColumn(0);
      return;
    }
    this.closePage();
    this.openPage();
  }

  /** The lines between the columns of the current run, down to `bottom`. */
  private separators(build: PageBuild, bottom: number): void {
    if (!build.separator || !(bottom > build.regionTop)) return;
    for (let i = 1; i < build.columns.length; i++) {
      const x = (build.columns[i - 1].x + build.columns[i - 1].width + build.columns[i].x) / 2;
      build.body.push({ type: 'line', x1: x, y1: build.regionTop, x2: x, y2: bottom, color: '#000000', width: 0.75, style: 'single' });
    }
  }
}

/** Replaces `remove` items of the queue at `at` by `items`. Not with splice and spread: a paragraph may have more lines than a call takes arguments. */
function replaceIn(queue: QueueItem[], at: number, remove: number, items: QueueItem[]): void {
  const tail = queue.slice(at + remove);
  queue.length = at;
  for (let i = 0; i < items.length; i++) queue.push(items[i]);
  for (let i = 0; i < tail.length; i++) queue.push(tail[i]);
}

function asRepeat(item: PageItem): PageItem {
  return item.type === 'text' || item.type === 'mark' ? { ...item, repeat: true } : item;
}
