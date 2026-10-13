// Tracked changes as a list, and accepting and rejecting them.
//
// A change in the list is what a reader takes for one edit: pieces that
// follow each other without other text in between, of one kind, by one
// author, with one date (or one id), are one change. An insertion that runs
// over a paragraph mark, or over the rows of a table, is one change with the
// mark and the rows in it. A deletion beside an insertion of the same author
// is two changes, as Word counts them. The two ends of a move are two
// changes that are resolved together.
//
// What accepting and rejecting do (Word's meaning):
//   - insertion: accepted, its content is plain content; rejected, it goes.
//   - deletion: accepted, its content goes; rejected, it is plain content
//     again (<w:delText> is text, <w:delInstrText> field code).
//   - inserted paragraph mark: accepted, a plain mark; rejected, the mark
//     goes and the paragraph is joined to the next one. A deleted mark the
//     other way round. The joined paragraph is the second one with the
//     content of the first before its own: the mark that stays is its mark,
//     and with it its formatting, its list, the record of a change of them
//     and the section it ends (Word's rule). A section break the first
//     paragraph carried goes with its mark, and is counted (sectionsRemoved).
//     A mark that has to go and has no paragraph after it to join (the last
//     one of a cell, one before a table) goes with its paragraph when that is
//     empty and the list of blocks can do without it; otherwise the mark
//     stays, as a plain mark.
//   - move: accepted, the content leaves where it came from and is plain
//     where it went; rejected, the other way round. Both ends that carry the
//     same name (their range marks) are resolved together and the range marks
//     go; an end without a name is resolved alone.
//   - formatting change (<w:rPrChange>, <w:pPrChange>, <w:sectPrChange>,
//     <w:tblPrChange>, <w:tblPrExChange>, <w:trPrChange>, <w:tcPrChange>):
//     accepted, the record goes; rejected, the properties the record holds
//     are read with the reader of the docx format and take the place of the
//     present ones.
//   - inserted or deleted table row: the row stays or goes; the table goes
//     with its last row.
//   - inserted or deleted table column (<w:cellIns>, <w:cellDel> in a cell
//     of every row, over the same grid columns, by one author at one date):
//     ONE change, resolved with all its cells. The cells stay or go together,
//     a row goes with its last cell, the grid columns of cells that go leave
//     the grid (those columns, wherever they are) and a table of a fixed
//     width becomes that much narrower. No row is ever left with another
//     number of columns than the others. The record of the grid as it was
//     (<w:tblGridChange>, which names nobody) is part of the first inserted
//     column of its table and goes with it.
//   - a record of the grid in a table without an inserted column: accepted,
//     the record goes; rejected, the grid is the one the record holds when
//     that has as many columns as the rows need, and the record goes either
//     way.
// Of content that goes, the marks of bookmarks, comments and permissions and
// the references to comments stay where it was: the rule every deletion
// follows. A note whose last reference goes leaves with it (see Draft).
//
// What the model only keeps as XML (merged cells, a record of the grid that
// holds more than columns, a changed list number, changes inside text boxes
// and other kept elements), and inserted or deleted cells that do not make a
// whole column (some rows only, a merged cell among them, two authors or
// dates), is listed as "other" and cannot be resolved; "all"
// passes over it and says how many changes it passed over. Nothing is ever
// resolved halfway: a change one part of which cannot be resolved (it would
// leave half a field) is not touched at all.
//
// Accepting and rejecting are never tracked themselves, and a document whose
// protection is enforced allows neither.

import { createContext, createReaderState, kindOf, type PartContext } from '../../formats/docx/common';
import { readCellProps, readParaProps, readRowProps, readRunProps, readSectionProps, readTableProps } from '../../formats/docx/props';
import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  enforcedProtection,
  type Block,
  type CellProps,
  type DocModel,
  type ExtraXml,
  type Inline,
  type ParaProps,
  type Paragraph,
  type RevisionInfo,
  type RowProps,
  type Run,
  type RunItem,
  type RunProps,
  type SectionProps,
  type Story,
  type Table,
  type TableCell,
  type TableProps,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { buildEdit, type DocEdit, type Draft } from './draft';
import { flatten, harvestBlocks, harvestRow, paragraphWith, rebuild, withoutWrapper, type Harvest, type Piece } from './pieces';
import { collapsed, itemsLength, paragraphIn, paragraphLength, resolvePosition, storyOf, type DocPosition, type DocRange, type StoryHandle } from './position';
import { cellWith, rowSpans, rowWith, type Leftover } from './structure';
import { neighbour } from './text';
import { forcesTracking, plainProps, revisionExtras } from './track';

/** What accepting or rejecting throws in a document whose protection is enforced: Word allows neither there. */
export const REVISIONS_PROTECTED = 'The protection of the document does not allow changes to be accepted or rejected';
/** What accepting or rejecting throws for a change that cannot be resolved (yet): nothing is changed. */
export const REVISION_UNSUPPORTED = 'This change cannot be accepted or rejected yet';
/** What accepting or rejecting throws for a change the document no longer has. */
export const REVISION_GONE = 'The change is no longer in the document';
/** What setTrackRevisions() throws where tracking cannot be turned off. */
export const TRACK_ENFORCED = 'Tracking cannot be turned off in this document';

export type RevisionKind = 'insert' | 'delete' | 'moveFrom' | 'moveTo' | 'format' | 'paragraphFormat' | 'tableFormat' | 'sectionFormat' | 'other';

/** One tracked change of a document, as revisionsOf() lists it. It describes the document as it was when it was listed. */
export interface DocRevision {
  /** Its place in the list. */
  index: number;
  /**
   * "insert" and "delete": text, paragraph marks, table rows and table
   * cells; "moveFrom"
   * and "moveTo": the two ends of a move; "format": character formatting (of
   * text or of a paragraph mark); "paragraphFormat"; "tableFormat": of a
   * table, a row or a cell, and the grid of a table as it was before its
   * columns changed (which names no author); "sectionFormat"; "other": a
   * change the engine
   * only keeps and cannot resolve.
   */
  kind: RevisionKind;
  author?: string;
  /** ISO 8601 text as the file has it. */
  date?: string;
  /** The story it is in: "body", "header:<rId>", "footnote:<id>", ... */
  story: string;
  /** Where it begins and ends. The two are the same place for what takes no room: a deletion, a paragraph mark. */
  start: DocPosition;
  end: DocPosition;
  /** The beginning of its text (at most 80 characters and "…"); "¶" stands for a paragraph mark. Empty for formatting. */
  text: string;
  /** Paragraph marks and table rows it holds. */
  marks: number;
  rows: number;
  /** 1 for a column of a table that was put in or taken out (a cell in every row, and for an inserted column the record of the grid); 0 otherwise. */
  columns: number;
  /**
   * For a column: where each of its cells begins and ends, row by row. `start` and `end` are those of the cell in
   * the first row: the cells of a column do not follow each other in the text, other cells stand between them.
   */
  cells?: { start: DocPosition; end: DocPosition }[];
  /** For "other": how many elements of tracked changes the kept XML holds; 1 otherwise. */
  count: number;
  /** The name of the move this is an end of, when the file names it: the other end is resolved with it. */
  move?: string;
  /** Whether acceptRevision() and rejectRevision() can resolve it. Rejecting can be impossible where accepting is not. */
  canAccept: boolean;
  canReject: boolean;
}

export interface DocRevisions {
  list: DocRevision[];
  /** The document has more changes than the list holds. */
  more: boolean;
}

export interface ListRevisionsOptions {
  /** Most changes listed; the default is REVISION_LIST_LIMIT. */
  limit?: number;
}

/** Most changes listRevisions() lists unless it is told otherwise. */
export const REVISION_LIST_LIMIT = 2000;
/** Most characters of the text of a change. */
const TEXT_LIMIT = 80;
/** Most nodes of kept XML one listing looks through for changes. */
const SCAN_BUDGET = 400000;
const MAX_DEPTH = 256;
const MAX_CHAIN = 4096;

/** What accepting or rejecting returns besides the edit. */
export interface RevisionEdit extends DocEdit {
  /**
   * Section breaks that go with the paragraph marks the edit removes: the
   * text before each is in the section after it from then on (page size,
   * margins, headers). Word does the same; the number is here to be told.
   */
  sectionsRemoved: number;
}

/** What "all" returns besides: how many changes it resolves and how many it passes over. */
export interface RevisionsEdit extends RevisionEdit {
  resolved: number;
  skipped: number;
}

type Verdict = 'accept' | 'reject';
/** What in its host a part is: a wrapper or an inline of a paragraph, or a name. */
type Slot = Wrapper | Inline | string;

interface Part {
  /** The paragraph, table, row, cell or body the part is in (an object no two places share). */
  host: object;
  slot: Slot;
  /** Tables around the host, outermost first: where a resolver has to go into. */
  tables: Table[];
  rejectable: boolean;
}

interface Entry {
  rev: DocRevision;
  parts: Part[];
  ids: Set<number>;
}

/** The cells of one kind (inserted, deleted) over the same grid columns of a table, by one author at one date. */
interface ColumnGroup {
  flag: 'inserted' | 'deleted';
  /** The first grid column they cover. */
  from: number;
  rows: Set<number>;
  /** There is one in every row: a column, which is one change. */
  whole: boolean;
  entry?: Entry;
}

interface Collected {
  entries: Entry[];
  more: boolean;
  /** The range marks of each named move. */
  ranges: Map<string, Part[]>;
}

type RevisionWrapper = Extract<Wrapper, { kind: 'ins' | 'del' | 'moveFrom' | 'moveTo' }>;

const isRevision = (w: Wrapper): w is RevisionWrapper => w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo';
const hides = (w: Wrapper): boolean => w.kind === 'del' || w.kind === 'moveFrom';

const WRAPPER_KIND: { [kind: string]: RevisionKind } = { ins: 'insert', del: 'delete', moveFrom: 'moveFrom', moveTo: 'moveTo' };
const MARK_FLAGS = ['inserted', 'deleted', 'moveFrom', 'moveTo'] as const;
const FLAG_KIND: { [flag: string]: RevisionKind } = { inserted: 'insert', deleted: 'delete', moveFrom: 'moveFrom', moveTo: 'moveTo' };

/** Does a verdict take the content (or the mark, or the row) of a revision of this kind out of the document. */
const removes = (kind: RevisionKind, verdict: Verdict): boolean => (verdict === 'accept' ? kind === 'delete' || kind === 'moveFrom' : kind === 'insert' || kind === 'moveTo');

const localName = (name: string): string => name.slice(name.indexOf(':') + 1);

/** Elements that are tracked changes, by local name. */
const REVISION_ELEMENTS = new Set<string>([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'rPrChange',
  'pPrChange',
  'sectPrChange',
  'tblPrChange',
  'tblPrExChange',
  'trPrChange',
  'tcPrChange',
  'tblGridChange',
  'numberingChange',
  'cellIns',
  'cellDel',
  'cellMerge',
  'customXmlInsRangeStart',
  'customXmlDelRangeStart',
  'customXmlMoveFromRangeStart',
  'customXmlMoveToRangeStart',
]);
const CELL_MARKUP = new Set<string>(['cellIns', 'cellDel', 'cellMerge']);

function hasAttr(el: XmlElement, local: string): boolean {
  let found = false;
  el.attrs.forEach((_value, name) => {
    if (localName(name) === local) found = true;
  });
  return found;
}

/** What kept XML holds of tracked changes: how many, and the names of the moves whose range marks are in it. */
interface Kept {
  count: number;
  moves: string[];
}

/**
 * What was found in each element kept as XML. Such an element is never
 * changed, so it is looked through once however often the changes are listed.
 */
const KEPT = new WeakMap<XmlElement, Kept>();

/** How many nodes of kept XML were looked at since the module was loaded. For tests: a second listing looks at none again. */
export const revisionCounters = { scanned: 0 };

function attrOf(el: XmlElement, local: string): string | undefined {
  let found: string | undefined;
  el.attrs.forEach((value, name) => {
    if (found === undefined && localName(name) === local) found = value;
  });
  return found;
}

function scanKept(el: XmlElement, into: Kept, budget: { left: number }, depth: number): void {
  if (depth > MAX_DEPTH || budget.left-- <= 0) return;
  revisionCounters.scanned++;
  const local = localName(el.name);
  // Elements of other vocabularies share some of the names: a change has an id.
  if (REVISION_ELEMENTS.has(local) && hasAttr(el, 'id')) into.count++;
  else if (local === 'moveFromRangeStart' || local === 'moveToRangeStart') {
    const name = attrOf(el, 'name');
    if (name !== undefined && into.moves.length < 4096) into.moves.push(name);
  }
  for (const node of el.children) if (node.type === 'element') scanKept(node, into, budget, depth + 1);
}

/** The tracked changes an element kept as XML holds. */
function keptChanges(el: XmlElement, budget: { left: number }): Kept {
  const known = KEPT.get(el);
  if (known) return known;
  const found: Kept = { count: 0, moves: [] };
  scanKept(el, found, budget, 0);
  // What the budget cut short is not the whole answer: it is not kept.
  if (budget.left > 0) KEPT.set(el, found);
  return found;
}

const countChanges = (el: XmlElement, budget: { left: number }): number => keptChanges(el, budget).count;

/** Who and when, of a formatting change kept as XML. */
function infoOf(change: XmlElement): RevisionInfo {
  const info: RevisionInfo = {};
  change.attrs.forEach((value, name) => {
    const local = localName(name);
    if (local === 'id' && /^\d{1,10}$/.test(value)) info.id = Number(value);
    else if (local === 'author') info.author = value;
    else if (local === 'date') info.date = value;
  });
  return info;
}

const RAW_TEXT = new Set<string>(['text', 'deletedText']);

function itemsPreview(items: RunItem[], code: true | boolean[] | undefined): string {
  let out = '';
  for (let i = 0; i < items.length && out.length <= TEXT_LIMIT; i++) {
    if (code !== undefined && (code === true || code[i])) continue;
    const item = items[i];
    if (RAW_TEXT.has(item.type)) out += (item as { text: string }).text;
    else if (item.type === 'tab' || item.type === 'break') out += ' ';
    else if (item.type === 'noBreakHyphen') out += '-';
    else if (item.type === 'symbol') out += item.char;
  }
  return out;
}

const inlinePreview = (inline: Inline): string => (inline.type === 'run' ? itemsPreview(inline.content, inline.fieldCode) : inline.type === 'math' ? inline.text : '');

const MOVE_RANGE = new Set<string>(['moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd']);
/** What stays of content that goes: the marks of ranges that are not content themselves. */
const SURVIVORS = new Set<string>(['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'permStart', 'permEnd']);
/** Items of a run that are no content: a run of nothing else does not part two changes. */
const HINTS = new Set<string>(['lastRenderedPageBreak', 'misc']);

/** Is the wrapper inside an <mc:AlternateContent>, which is written back whole: nothing in it can be resolved. */
function inAlternate(wrap: Wrapper | undefined): boolean {
  let steps = 0;
  for (let w = wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) if (w.kind === 'alternate') return true;
  return false;
}

/** The paragraph a path leads to inside a list of blocks, through the cells of tables. */
function paragraphAtPath(blocks: Block[], rel: number[]): Paragraph | undefined {
  let list = blocks;
  for (let k = 0; k < rel.length && k < MAX_CHAIN; k += 3) {
    const block = list[rel[k]];
    if (!block) return undefined;
    if (k === rel.length - 1) return block.type === 'paragraph' ? block : undefined;
    if (block.type !== 'table') return undefined;
    const row = block.rows[rel[k + 1]];
    const cell = row && row.cells[rel[k + 2]];
    if (!cell) return undefined;
    list = cell.blocks;
  }
  return undefined;
}

// ----------------------------------------------------------------------- list

class Lister {
  readonly entries: Entry[] = [];
  readonly ranges = new Map<string, Part[]>();
  more = false;
  private stopped = false;
  private current: Entry | undefined;
  /** Nothing but changes stood between the current change and here. */
  private adjacent = false;
  /** The column whose cell is being walked: what the cell holds of its author and date is part of it, whatever else stands in the cell. */
  private column: Entry | undefined;
  private story = 'body';
  private readonly tables: Table[] = [];
  private readonly budget = { left: SCAN_BUDGET };
  private readonly scanned = new Set<object>();
  private readonly locked: boolean;
  /** The move a range mark opened and has not closed yet, by kind; and the names of the open ranges by their ids. */
  private activeMove: { moveFrom?: string; moveTo?: string } = {};
  private readonly moveIds = new Map<string, string>();
  private last: DocPosition = { story: 'body', path: [0], offset: 0 };
  /** Names of the moves that have range marks in XML the model only keeps. */
  private readonly keptMoves = new Set<string>();

  constructor(
    private readonly doc: DocModel,
    private readonly limit: number,
  ) {
    this.locked = enforcedProtection(doc) !== undefined;
  }

  run(): void {
    const doc = this.doc;
    const seen = new Set<object>();
    const story = (id: string, blocks: Block[], host: object): void => {
      if (this.stopped || seen.has(host)) return;
      seen.add(host);
      this.story = id;
      this.current = undefined;
      this.adjacent = false;
      this.activeMove = {};
      this.last = { story: id, path: [0], offset: 0 };
      this.blocks(blocks, [], 0);
    };
    story('body', doc.body.blocks, doc.body);
    const section = doc.body.section;
    if (!this.stopped && section && section.change) {
      this.story = 'body';
      this.adjacent = false;
      this.part('sectionFormat', infoOf(section.change), doc.body, 'section', this.last, this.last, '', true);
    }
    doc.headers.forEach((item, id) => story('header:' + id, item.blocks, item));
    doc.footers.forEach((item, id) => story('footer:' + id, item.blocks, item));
    for (const [kind, collection] of [
      ['footnote', doc.footnotes],
      ['endnote', doc.endnotes],
    ] as const) {
      if (!collection) continue;
      collection.items.forEach((note, id) => {
        // The ids storyOf() can name; the separators of the notes area hold no text of the document.
        if (note.kind === undefined && /^-?\d{1,9}$/.test(String(id))) story(`${kind}:${id}`, note.blocks, note);
      });
    }
    // Comments are no story a position can name: what they hold is kept, and counted.
    if (doc.comments && !this.stopped) {
      let count = 0;
      doc.comments.items.forEach((comment) => (count += this.countInBlocks(comment.blocks, 0)));
      if (count) this.other(count, { story: 'body', path: [0], offset: 0 });
    }
    if (this.keptMoves.size) {
      for (const entry of this.entries) {
        if (entry.rev.move === undefined || !this.keptMoves.has(entry.rev.move)) continue;
        entry.rev.canAccept = false;
        entry.rev.canReject = false;
      }
    }
  }

  /** Changes in blocks that are not walked (a comment): every kind is counted alike. */
  private countInBlocks(blocks: Block[], depth: number): number {
    if (depth > MAX_DEPTH) return 0;
    let count = 0;
    for (const block of blocks) {
      if (this.budget.left-- <= 0) break;
      if (block.type === 'paragraph') {
        const mark = block.props.markProps;
        if (block.props.change) count++;
        if (mark) {
          if (mark.change) count++;
          for (const flag of MARK_FLAGS) if (mark[flag]) count++;
        }
        const seen = new Set<Wrapper>();
        for (const inline of block.children) {
          if ((inline.type === 'moveFromRangeStart' || inline.type === 'moveToRangeStart') && inline.name !== undefined) this.keptMoves.add(inline.name);
          let steps = 0;
          for (let w = inline.wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) {
            if (isRevision(w) && !seen.has(w)) {
              seen.add(w);
              count++;
            }
          }
          if (inline.type === 'run' && inline.props.change) count++;
        }
      } else if (block.type === 'table') {
        if (block.gridChange) count++;
        for (const row of block.rows) {
          if (row.props.inserted) count++;
          if (row.props.deleted) count++;
          for (const cell of row.cells) {
            if (cell.props.inserted) count++;
            if (cell.props.deleted) count++;
            count += this.countInBlocks(cell.blocks, depth + 1);
          }
        }
      } else if (block.type === 'unknown') count += countChanges(block.xml, this.budget);
    }
    return count;
  }

  private canStart(): boolean {
    if (this.entries.length < this.limit) return true;
    this.more = true;
    this.stopped = true;
    return false;
  }

  /** Adds a part of a change: to the change before it when the two are one, else as a change of its own. */
  private part(kind: RevisionKind, info: RevisionInfo, host: object, slot: Slot, start: DocPosition, end: DocPosition, text: string, rejectable: boolean, move?: string): Entry | undefined {
    if (this.stopped) return undefined;
    const part: Part = { host, slot, tables: this.tables.slice(), rejectable };
    const fits = (to: Entry | undefined): boolean =>
      to !== undefined && to.rev.kind === kind && to.rev.author === info.author && to.rev.move === move && (to.rev.date === info.date || (info.id !== undefined && to.ids.has(info.id)));
    let entry = this.current;
    let same = this.adjacent && fits(entry);
    // In a cell of a column that is a change, the column comes first.
    if (fits(this.column)) {
      entry = this.column;
      same = true;
    }
    if (same && entry) {
      entry.parts.push(part);
      // (A column is where its cells are, see DocRevision.cells: what a cell holds does not move its end.)
      if (entry.rev.columns === 0) entry.rev.end = end;
      if (!rejectable) entry.rev.canReject = false;
    } else {
      if (!this.canStart()) return undefined;
      const rev: DocRevision = { index: this.entries.length, kind, story: this.story, start, end, text: '', marks: 0, rows: 0, columns: 0, count: 1, canAccept: !this.locked, canReject: !this.locked && rejectable };
      if (info.author !== undefined) rev.author = info.author;
      if (info.date !== undefined) rev.date = info.date;
      if (move !== undefined) rev.move = move;
      entry = { rev, parts: [part], ids: new Set<number>() };
      this.entries.push(entry);
      this.current = entry;
    }
    if (info.id !== undefined) entry.ids.add(info.id);
    this.addText(entry, text);
    this.adjacent = true;
    return entry;
  }

  private addText(entry: Entry, text: string): void {
    const rev = entry.rev;
    if (text === '' || rev.text.length > TEXT_LIMIT) return;
    const joined = rev.text + text;
    rev.text = joined.length > TEXT_LIMIT ? joined.slice(0, TEXT_LIMIT) + '…' : joined;
  }

  /** A change the model only keeps: listed, never resolved. */
  private other(count: number, at: DocPosition): void {
    this.adjacent = false;
    this.current = undefined;
    if (count <= 0 || this.stopped || !this.canStart()) return;
    this.entries.push({
      rev: { index: this.entries.length, kind: 'other', story: at.story, start: at, end: at, text: '', marks: 0, rows: 0, columns: 0, count, canAccept: false, canReject: false },
      parts: [],
      ids: new Set<number>(),
    });
  }

  private scan(el: XmlElement | undefined, at: DocPosition): void {
    if (!el || this.scanned.has(el)) return;
    this.scanned.add(el);
    const kept = keptChanges(el, this.budget);
    if (this.budget.left <= 0) this.more = true;
    // A move one end of which is here cannot be resolved: this end would stay behind.
    for (const name of kept.moves) this.keptMoves.add(name);
    if (kept.count) this.other(kept.count, at);
  }

  private scanExtras(extras: ExtraXml[] | undefined, at: DocPosition): void {
    if (extras) for (const extra of extras) this.scan(extra.el, at);
  }

  private scanItems(items: RunItem[], at: DocPosition, depth: number): void {
    for (const item of items) {
      if (item.type === 'drawing' || item.type === 'pict' || item.type === 'object' || item.type === 'alternate' || item.type === 'unknown') this.scan(item.xml, at);
      else if (depth > MAX_DEPTH) return;
    }
  }

  private rangeMark(type: string, id: string, name: string | undefined, host: object, slot: Slot): void {
    const kind = type.indexOf('moveFrom') === 0 ? 'moveFrom' : 'moveTo';
    const key = kind + '\u0001' + id;
    let move: string | undefined;
    if (type.slice(-5) === 'Start') {
      if (name === undefined) return;
      move = name;
      this.moveIds.set(key, name);
      this.activeMove[kind] = name;
    } else {
      move = this.moveIds.get(key);
      if (move === undefined) return;
      if (this.activeMove[kind] === move) delete this.activeMove[kind];
    }
    let list = this.ranges.get(move);
    if (!list) this.ranges.set(move, (list = []));
    list.push({ host, slot, tables: this.tables.slice(), rejectable: true });
  }

  private blocks(blocks: Block[], base: number[], depth: number): void {
    if (depth > MAX_DEPTH) return;
    for (let i = 0; i < blocks.length && !this.stopped; i++) {
      const block = blocks[i];
      if (block.wrap && (block.wrap.kind === 'alternate' || inAlternate(block.wrap))) {
        // Written back whole: what it holds is counted from its XML, once.
        let steps = 0;
        for (let w: Wrapper | undefined = block.wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) if (w.kind === 'alternate') this.scan(w.xml, this.last);
        this.adjacent = false;
        continue;
      }
      if (block.type === 'paragraph') this.paragraph(block, base.concat([i]));
      else if (block.type === 'table') this.table(block, blocks, base, i, depth);
      else if (block.type === 'mark') {
        if (MOVE_RANGE.has(block.mark.type)) this.rangeMark(block.mark.type, block.mark.id, block.mark.name, block, 'self');
      } else if (block.type === 'unknown') this.scan(block.xml, this.last);
    }
  }

  private paragraph(paragraph: Paragraph, path: number[]): void {
    const story = this.story;
    const at = (offset: number): DocPosition => ({ story, path, offset });
    const props = paragraph.props;
    const length = paragraphLength(paragraph);
    this.last = at(0);
    if (props.change) this.part('paragraphFormat', infoOf(props.change), paragraph, 'paraFormat', at(0), at(length), '', true);
    this.scanExtras(props.extras, at(0));

    const seen = new Map<Wrapper, Entry>();
    /** The wrappers around the inline before: one that comes again after something else is another element of the file. */
    let around = new Set<Wrapper>();
    let offset = 0;
    for (const inline of paragraph.children) {
      if (this.stopped) return;
      // The chain, innermost first; what is inside an AlternateContent cannot be resolved.
      const chain: Wrapper[] = [];
      for (let w = inline.wrap; w && chain.length < MAX_CHAIN; w = w.parent) chain.push(w);
      let frozen = -1;
      let hidden = false;
      for (let i = 0; i < chain.length; i++) {
        if (chain[i].kind === 'alternate') frozen = i;
        if (hides(chain[i])) hidden = true;
      }
      let len = 0;
      if (!hidden) {
        if (inline.type === 'math') len = inline.text.length;
        else if (inline.type === 'run' && inline.fieldCode !== true) len = itemsLength(inline.content, inline.fieldCode, 0);
      }
      const here = at(offset);
      const after = at(offset + len);
      if (frozen >= 0) for (let i = 0; i <= frozen; i++) if (chain[i].kind === 'alternate') this.scan((chain[i] as Extract<Wrapper, { kind: 'alternate' }>).xml, here);
      let tracked = false;
      for (let i = chain.length - 1; i > frozen; i--) {
        const w = chain[i];
        if (!isRevision(w)) continue;
        tracked = true;
        const kind = WRAPPER_KIND[w.kind];
        let entry = around.has(w) ? seen.get(w) : undefined;
        // The text of a deletion is what it hides; of an insertion, what it shows.
        const shows = kind === 'delete' || kind === 'moveFrom' ? true : !hidden;
        const text = shows ? inlinePreview(inline) : '';
        if (entry) {
          if (entry.rev.columns === 0) entry.rev.end = after;
          this.addText(entry, text);
          continue;
        }
        const move = kind === 'moveFrom' || kind === 'moveTo' ? this.activeMove[kind] : undefined;
        entry = this.part(kind, w, paragraph, w, here, after, text, true, move);
        if (!entry) return;
        seen.set(w, entry);
      }
      around = new Set<Wrapper>(chain);
      if (frozen < 0 && MOVE_RANGE.has(inline.type)) {
        const mark = inline as Extract<Inline, { id: string }>;
        this.rangeMark(mark.type, mark.id, mark.name, paragraph, inline);
      }
      if (inline.type === 'run') {
        if (inline.props.change && frozen < 0) {
          tracked = true;
          if (!this.part('format', infoOf(inline.props.change), paragraph, inline, here, after, hidden ? '' : inlinePreview(inline), true)) return;
        }
        this.scanItems(inline.content, here, 0);
        if (!tracked && inline.content.some((item) => !HINTS.has(item.type))) this.adjacent = false;
      } else if (inline.type === 'math') {
        if (!tracked) this.adjacent = false;
      } else if (inline.type === 'unknown') this.scan(inline.xml, here);
      offset += len;
    }

    const end = at(length);
    const mark = props.markProps;
    let flagged = false;
    if (mark) {
      for (const flag of MARK_FLAGS) {
        const info = mark[flag];
        if (!info) continue;
        flagged = true;
        const kind = FLAG_KIND[flag];
        const entry = this.part(kind, info, paragraph, 'mark:' + flag, end, end, '¶', true, kind === 'moveFrom' || kind === 'moveTo' ? this.activeMove[kind] : undefined);
        if (!entry) return;
        entry.rev.marks++;
      }
      if (mark.change) this.part('format', infoOf(mark.change), paragraph, 'markFormat', end, end, '', true);
      this.scanExtras(mark.extras, end);
    }
    if (!flagged) this.adjacent = false;
    if (props.section) {
      if (props.section.change) this.part('sectionFormat', infoOf(props.section.change), paragraph, 'section', end, end, '', true);
      this.scanExtras(props.section.extras, end);
    }
  }

  /**
   * The columns of a table that were put in or taken out. The cells that say so (<w:cellIns>, <w:cellDel>) are one
   * change when there is one of them in every row, over the same grid columns, by one author at one date: a column
   * is resolved whole or not at all, so that no row is ever left with another number of columns than the others.
   * Cells that do not make such a column (some rows only, a merged cell among them, two authors) are in `loose`:
   * they are kept and listed, never resolved. `owner` is the inserted column the record of the grid goes with: the
   * first of them.
   */
  private columnsOf(table: Table): { of: Map<TableCell, { inserted?: ColumnGroup; deleted?: ColumnGroup }>; owner: ColumnGroup | undefined } {
    const of = new Map<TableCell, { inserted?: ColumnGroup; deleted?: ColumnGroup }>();
    let owner: ColumnGroup | undefined;
    let any = false;
    for (const row of table.rows) for (const cell of row.cells) if (cell.props.inserted || cell.props.deleted) any = true;
    if (!any) return { of, owner };
    // A row or a cell that is written back whole cannot lose or keep a cell with the others.
    let frozen = false;
    for (const row of table.rows) {
      if (row.wrap && inAlternate(row.wrap)) frozen = true;
      for (const cell of row.cells) if (cell.wrap && inAlternate(cell.wrap)) frozen = true;
    }
    for (const flag of ['inserted', 'deleted'] as const) {
      const groups = new Map<string, ColumnGroup>();
      for (let r = 0; r < table.rows.length; r++) {
        const row = table.rows[r];
        const spans = rowSpans(row);
        for (let c = 0; c < row.cells.length; c++) {
          const info = row.cells[c].props[flag];
          if (!info) continue;
          const key = [spans[c][0], spans[c][1], info.author === undefined ? '' : 'a' + info.author, info.date === undefined ? '' : 'd' + info.date].join('\u0001');
          let group = groups.get(key);
          if (!group) groups.set(key, (group = { flag, from: spans[c][0], rows: new Set<number>(), whole: false }));
          group.rows.add(r);
          let entry = of.get(row.cells[c]);
          if (!entry) of.set(row.cells[c], (entry = {}));
          entry[flag] = group;
        }
      }
      groups.forEach((group) => {
        group.whole = !frozen && group.rows.size === table.rows.length;
        if (flag === 'inserted' && group.whole && (!owner || group.from < owner.from)) owner = group;
      });
    }
    return { of, owner };
  }

  private table(table: Table, siblings: Block[], base: number[], index: number, depth: number): void {
    const story = this.story;
    const path = base.concat([index]);
    const firstIn = paragraphIn(siblings, base, index, 1);
    const start: DocPosition = firstIn ? { story, path: firstIn, offset: 0 } : this.last;
    // The table is among the tables a resolver goes into for its own formatting too.
    this.tables.push(table);
    if (table.props.change) this.part('tableFormat', infoOf(table.props.change), table, 'tableFormat', start, start, '', true);
    const columns = this.columnsOf(table);
    // The grid as it was goes with the column that was put in; without one it is a change of the table nobody signs.
    if (table.gridChange && !columns.owner) {
      this.adjacent = false;
      this.part('tableFormat', table.gridChange.id === undefined ? {} : { id: table.gridChange.id }, table, 'grid', start, start, '', true);
      this.adjacent = false;
    }
    this.scanExtras(table.props.extras, start);
    this.scanExtras(table.extras, start);
    for (let r = 0; r < table.rows.length && !this.stopped; r++) {
      const row = table.rows[r];
      const rowPath = path.concat([r]);
      if (row.wrap && (row.wrap.kind === 'alternate' || inAlternate(row.wrap))) {
        this.adjacent = false;
        continue;
      }
      const head = paragraphIn(row.cells.length ? row.cells[0].blocks : [], rowPath.concat([0]), 0, 1);
      const from: DocPosition = head ? { story, path: head, offset: 0 } : start;
      let to = from;
      for (let c = row.cells.length - 1; c >= 0; c--) {
        const blocks = row.cells[c].blocks;
        const tail = paragraphIn(blocks, rowPath.concat([c]), blocks.length - 1, -1);
        if (!tail) continue;
        const last = paragraphAtPath(blocks, tail.slice(rowPath.length + 1));
        to = { story, path: tail, offset: last ? paragraphLength(last) : 0 };
        break;
      }
      let flagged = false;
      for (const flag of ['inserted', 'deleted'] as const) {
        const info = row.props[flag];
        if (!info) continue;
        flagged = true;
        const entry = this.part(FLAG_KIND[flag], info, row, 'row:' + flag, from, to, '', true);
        if (!entry) break;
        entry.rev.rows++;
      }
      if (this.stopped) break;
      // A row that is no change itself parts what is before it from what is in it, unless its formatting is one.
      if (!flagged && !row.props.change && !(row.tableProps && row.tableProps.change)) this.adjacent = false;
      if (row.props.change) this.part('tableFormat', infoOf(row.props.change), row, 'rowFormat', from, from, '', true);
      if (row.tableProps && row.tableProps.change) this.part('tableFormat', infoOf(row.tableProps.change), row, 'rowTableFormat', from, from, '', true);
      this.scanExtras(row.props.extras, from);
      if (row.tableProps) this.scanExtras(row.tableProps.extras, from);
      this.scanExtras(row.extras, from);
      for (let c = 0; c < row.cells.length && !this.stopped; c++) {
        const cell = row.cells[c];
        const cellPath = rowPath.concat([c]);
        if (cell.wrap && (cell.wrap.kind === 'alternate' || inAlternate(cell.wrap))) {
          this.adjacent = false;
          continue;
        }
        const head2 = paragraphIn(cell.blocks, cellPath, 0, 1);
        const cellStart: DocPosition = head2 ? { story, path: head2, offset: 0 } : from;
        // A cell that was put in or taken out with its column: part of the change its column is.
        let column: Entry | undefined;
        const groups = columns.of.get(cell);
        for (const flag of ['inserted', 'deleted'] as const) {
          const info = cell.props[flag];
          const group = groups && groups[flag];
          if (!info || !group) continue;
          if (!group.whole) {
            // Not a whole column: kept, and counted.
            this.other(1, cellStart);
            continue;
          }
          let cellEnd = cellStart;
          const tail2 = paragraphIn(cell.blocks, cellPath, cell.blocks.length - 1, -1);
          if (tail2) {
            const last2 = paragraphAtPath(cell.blocks, tail2.slice(cellPath.length));
            cellEnd = { story, path: tail2, offset: last2 ? paragraphLength(last2) : 0 };
          }
          const part: Part = { host: cell, slot: 'cell:' + flag, tables: this.tables.slice(), rejectable: true };
          if (!group.entry) {
            if (!this.canStart()) break;
            const rev: DocRevision = { index: this.entries.length, kind: FLAG_KIND[flag], story, start: cellStart, end: cellEnd, text: '', marks: 0, rows: 0, columns: 1, cells: [], count: 1, canAccept: !this.locked, canReject: !this.locked };
            if (info.author !== undefined) rev.author = info.author;
            if (info.date !== undefined) rev.date = info.date;
            group.entry = { rev, parts: [], ids: new Set<number>() };
            this.entries.push(group.entry);
            // The record of the grid is resolved with the column it was made for.
            if (group === columns.owner && table.gridChange) group.entry.parts.push({ host: table, slot: 'grid', tables: this.tables.slice(), rejectable: true });
          }
          group.entry.parts.push(part);
          (group.entry.rev.cells as { start: DocPosition; end: DocPosition }[]).push({ start: cellStart, end: cellEnd });
          if (info.id !== undefined) group.entry.ids.add(info.id);
          column = group.entry;
        }
        if (this.stopped) break;
        if (cell.props.change) {
          const markup = (cell.props.extras || []).some((extra) => CELL_MARKUP.has(localName(extra.el.name)));
          this.part('tableFormat', infoOf(cell.props.change), cell, 'cellFormat', cellStart, cellStart, '', !markup);
        }
        this.scanExtras(cell.props.extras, cellStart);
        // What the cell holds of the same author and date (the marks of its paragraphs, its deleted text) is the column too.
        this.column = column;
        this.blocks(cell.blocks, cellPath, depth + 1);
        this.column = undefined;
      }
    }
    this.tables.pop();
  }
}

function collect(doc: DocModel, limit: number): Collected {
  const lister = new Lister(doc, limit);
  lister.run();
  return { entries: lister.entries, more: lister.more, ranges: lister.ranges };
}

/**
 * The tracked changes of a document in reading order: the body, then the
 * headers and footers, then the footnotes and endnotes. At most
 * `options.limit` of them (REVISION_LIST_LIMIT unless given); `more` says
 * that the document has others. Costs one pass over the document.
 */
export function listRevisions(doc: DocModel, options?: ListRevisionsOptions): DocRevisions {
  let limit = REVISION_LIST_LIMIT;
  if (options && options.limit !== undefined) {
    if (typeof options.limit !== 'number' || !(options.limit >= 0)) throw new RangeError('The limit of a list of changes must be a number that is not negative');
    limit = Math.floor(options.limit);
  }
  const found = remember(collect(doc, limit));
  return { list: found.entries.map((entry) => entry.rev), more: found.more };
}

/** The tracked changes of a document: the list of listRevisions(). */
export function revisionsOf(doc: DocModel): DocRevision[] {
  return listRevisions(doc).list;
}

// ------------------------------------------------------------------- resolver

/**
 * Properties a rejected record brings back, with the tracked changes the present properties keep as XML: rejecting one
 * change of formatting never takes another change out of the document. A kept child of the same name that came with
 * the record gives way to the one that carries the change.
 */
function withRevisionExtras<T extends { extras?: ExtraXml[] }>(old: T, present: ExtraXml[] | undefined, empty: T): T {
  const kept = revisionExtras(present);
  if (!kept.length) return Object.keys(old).length ? old : empty;
  const names = new Set<string>();
  for (const extra of kept) names.add(localName(extra.el.name) + (extra.index < 0 ? ' (self)' : ''));
  const extras = (old.extras || []).filter((extra) => !names.has(localName(extra.el.name) + (extra.index < 0 ? ' (self)' : '')));
  return { ...old, extras: extras.concat(kept) };
}

const without = <T extends object>(props: T, key: keyof T, empty: T): T => {
  const out = { ...props };
  delete out[key];
  return Object.keys(out).length ? out : empty;
};

function childOf(ctx: PartContext, el: XmlElement, kind: string): XmlElement | undefined {
  for (const node of el.children) if (node.type === 'element' && kindOf(ctx, node.name) === kind) return node;
  return undefined;
}

/** An item as it is outside a deletion: <w:delText> is text again, <w:delInstrText> field code. */
function revived(content: RunItem[]): RunItem[] {
  let out: RunItem[] | undefined;
  for (let i = 0; i < content.length; i++) {
    const item = content[i];
    let next: RunItem | undefined;
    if (item.type === 'deletedText') next = item.source === undefined ? { type: 'text', text: item.text } : { type: 'text', text: item.text, source: item.source };
    else if (item.type === 'instrText' && item.deleted) {
      next = { ...item };
      delete next.deleted;
    }
    if (!next) continue;
    if (!out) out = content.slice();
    out[i] = next;
  }
  return out || content;
}

/**
 * The inlines with the complex fields that begin and end among them assembled
 * anew: the code and the result the first mark of each carries (what says
 * whether the engine computes the field) are those of the content as it is
 * now. Code that was deleted or inserted is no longer in it, or is. The rule
 * is the reader's (linkFields), so a field nothing happened to reads the same.
 */
function relinked(children: Inline[]): Inline[] {
  interface Open {
    inline: number;
    item: number;
    instruction: string;
    result: string;
    inResult: boolean;
  }
  const stack: Open[] = [];
  let changes: { inline: number; item: number; instruction: string; result: string }[] | undefined;
  for (let i = 0; i < children.length; i++) {
    const inline = children[i];
    if (inline.type !== 'run') continue;
    for (let k = 0; k < inline.content.length; k++) {
      const item = inline.content[k];
      if (item.type === 'fieldChar') {
        if (item.charType === 'begin') stack.push({ inline: i, item: k, instruction: '', result: '', inResult: false });
        else if (!stack.length) continue;
        else if (item.charType === 'separate') stack[stack.length - 1].inResult = true;
        else {
          const done = stack.pop() as Open;
          const begin = (children[done.inline] as Run).content[done.item];
          const field = begin.type === 'fieldChar' ? begin.field : undefined;
          if (field && !field.unclosed && (field.instruction !== done.instruction || field.result !== done.result)) (changes || (changes = [])).push(done);
        }
      } else if (!stack.length) continue;
      else if (item.type === 'instrText') {
        const top = stack[stack.length - 1];
        if (!top.inResult) top.instruction += item.text;
      } else if ((item.type === 'text' || item.type === 'tab') && stack.every((open) => open.inResult)) {
        for (const open of stack) open.result += item.type === 'tab' ? '\t' : item.text;
      }
    }
  }
  if (!changes) return children;
  const out = children.slice();
  for (const change of changes) {
    const run = out[change.inline] as Run;
    const content = run.content.slice();
    const begin = content[change.item];
    if (begin.type !== 'fieldChar') continue;
    content[change.item] = { ...begin, field: { instruction: change.instruction, result: change.result } };
    out[change.inline] = { ...run, content };
  }
  return out;
}

interface ResolvedParagraph {
  paragraph: Paragraph;
  /** The mark of the paragraph goes: it is to be joined to the paragraph after it. */
  join: boolean;
}

class Resolver {
  /** Parts that cannot be resolved as asked: the changes they belong to are left alone. */
  readonly refused = new Map<object, Set<Slot>>();
  /**
   * Paragraph marks that have to go but have to wait: what becomes of their paragraph depends on a change beside it
   * that is not decided about yet. They stay changes, to be resolved once that one is.
   */
  readonly deferred = new Map<object, Set<Slot>>();
  /** Section breaks that went with the marks of their paragraphs. */
  sectionsRemoved = 0;
  private readonly touched = new Set<Table>();
  private reader: PartContext | undefined;
  private host: StoryHandle['host'] | undefined;

  constructor(
    private readonly decisions: Map<object, Map<Slot, Verdict>>,
    tables: Iterable<Table>,
    /** The one verdict of all the decisions. */
    private readonly verdict: Verdict,
  ) {
    for (const table of tables) this.touched.add(table);
  }

  /**
   * The story whose blocks are resolved next: properties kept as XML are read
   * as its part spells them. With the blocks, the fields of the story are
   * checked first (see checkFields).
   */
  enter(host: StoryHandle['host'], blocks?: Block[]): void {
    this.host = host;
    this.reader = undefined;
    if (blocks) this.checkFields(blocks, [], undefined, 0);
  }

  /** The part whose decision takes an inline out of the document: the outermost wrapper whose content goes. */
  private removerOf(paragraph: Paragraph, inline: Inline): { host: object; slot: Slot } | undefined {
    const slots = this.decisions.get(paragraph);
    if (!slots) return undefined;
    let found: Wrapper | undefined;
    let steps = 0;
    for (let w = inline.wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) {
      if (!isRevision(w)) continue;
      const decided = slots.get(w);
      if (decided && removes(WRAPPER_KIND[w.kind], decided) && !inAlternate(w)) found = w;
    }
    return found ? { host: paragraph, slot: found } : undefined;
  }

  /**
   * A field goes whole or not at all. The marks of every complex field of the
   * story are followed, from paragraph to paragraph as the reader follows
   * them; where the decisions would take some marks of a field and leave
   * others, the parts that would take them are refused. `gone` is the part
   * that takes everything the blocks hold (a row that goes).
   */
  private checkFields(blocks: Block[], stack: { removers: ({ host: object; slot: Slot } | undefined)[] }[], gone: { host: object; slot: Slot } | undefined, depth: number): void {
    if (depth > MAX_DEPTH) return;
    for (const block of blocks) {
      if (block.type === 'table') {
        for (const row of block.rows) {
          let rowGone = gone;
          const slots = this.decisions.get(row);
          if (slots && !rowGone) {
            for (const flag of ['inserted', 'deleted'] as const) {
              const decided = slots.get('row:' + flag);
              if (decided && row.props[flag] !== undefined && removes(FLAG_KIND[flag], decided)) rowGone = { host: row, slot: 'row:' + flag };
            }
          }
          for (const cell of row.cells) {
            let cellGone = rowGone;
            const cellSlots = cellGone ? undefined : this.decisions.get(cell);
            if (cellSlots) {
              for (const flag of ['inserted', 'deleted'] as const) {
                const decided = cellSlots.get('cell:' + flag);
                if (decided && cell.props[flag] !== undefined && removes(FLAG_KIND[flag], decided)) cellGone = { host: cell, slot: 'cell:' + flag };
              }
            }
            this.checkFields(cell.blocks, stack, cellGone, depth + 1);
          }
        }
        continue;
      }
      if (block.type !== 'paragraph') continue;
      for (const inline of block.children) {
        if (inline.type !== 'run') continue;
        let remover: { host: object; slot: Slot } | undefined;
        let known = false;
        for (const item of inline.content) {
          if (item.type !== 'fieldChar') continue;
          if (!known) {
            remover = gone || this.removerOf(block, inline);
            known = true;
          }
          if (item.charType === 'begin') {
            if (stack.length < MAX_CHAIN) stack.push({ removers: [remover] });
            continue;
          }
          if (!stack.length) continue;
          const top = stack[stack.length - 1];
          top.removers.push(remover);
          if (item.charType !== 'end') continue;
          stack.pop();
          let taken = 0;
          for (const one of top.removers) if (one) taken++;
          if (taken === 0 || taken === top.removers.length) continue;
          for (const one of top.removers) if (one) this.refuse(one.host, one.slot);
        }
      }
    }
  }

  private ctx(): PartContext {
    if (!this.reader) {
      const host = this.host as Story;
      this.reader = createContext(createReaderState(), host.part, host.rootAttrs, host.relationships);
    }
    return this.reader;
  }

  /**
   * The decisions that take the mark of a paragraph are left for later (see list()): the mark stays the change it
   * is, and everything else that was decided is done.
   */
  private deferMarks(paragraph: Paragraph): void {
    const slots = this.decisions.get(paragraph);
    const mark = paragraph.props.markProps;
    if (!slots || !mark) return;
    let set = this.deferred.get(paragraph);
    if (!set) this.deferred.set(paragraph, (set = new Set<Slot>()));
    for (const flag of MARK_FLAGS) {
      const decided = slots.get('mark:' + flag);
      if (decided && mark[flag] !== undefined && removes(FLAG_KIND[flag], decided)) set.add('mark:' + flag);
    }
  }

  private refuse(host: object, slot: Slot): void {
    let set = this.refused.get(host);
    if (!set) this.refused.set(host, (set = new Set<Slot>()));
    set.add(slot);
  }

  // ---------------------------------------------------------- properties

  private runProps(props: RunProps, verdict: Verdict): RunProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', EMPTY_RUN_PROPS);
    const ctx = this.ctx();
    // The record is read as formatting and nothing else: a file that puts a tracked change into it (an inserted mark,
    // a record of its own) does not get one made of it.
    const old = plainProps(readRunProps(ctx, childOf(ctx, change, 'rPr')));
    // What says that the mark itself is a change is no formatting: it stays.
    let out: RunProps | undefined;
    for (const flag of MARK_FLAGS) {
      if (props[flag] === undefined) continue;
      if (!out) out = { ...old };
      out[flag] = props[flag];
    }
    return withRevisionExtras(out || old, props.extras, EMPTY_RUN_PROPS);
  }

  private paraProps(props: ParaProps, verdict: Verdict): ParaProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', EMPTY_PARA_PROPS);
    const ctx = this.ctx();
    const read = readParaProps(ctx, childOf(ctx, change, 'pPr'));
    // The record holds neither the mark nor the section: those are the present ones. A file that puts them into the
    // record all the same (the schema has no place for them there), or a tracked change, gets none of it applied.
    let old = read;
    if (read.markProps !== undefined || read.section !== undefined) {
      old = { ...read };
      delete old.markProps;
      delete old.section;
    }
    // The tracked changes the properties keep as XML (a list number somebody inserted) are not this one: they stay.
    old = withRevisionExtras(plainProps(old), props.extras, EMPTY_PARA_PROPS);
    // A list number that carries such a change stays with it, also where the record says the paragraph was in no list:
    // the number is somebody's change, to be resolved as one (and the kept <w:numPr> would be written whole anyway).
    if (old.numbering === undefined && props.numbering !== undefined && old.extras && old.extras.some((extra) => extra.partial === true && localName(extra.el.name) === 'numPr')) {
      old = { ...old, numbering: props.numbering };
    }
    if (props.markProps === undefined && props.section === undefined) return Object.keys(old).length ? old : EMPTY_PARA_PROPS;
    const out: ParaProps = { ...old };
    if (props.markProps !== undefined) out.markProps = props.markProps;
    if (props.section !== undefined) out.section = props.section;
    return out;
  }

  sectionProps(props: SectionProps, verdict: Verdict): SectionProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', {});
    const ctx = this.ctx();
    const inner = childOf(ctx, change, 'sectPr');
    const out: SectionProps = inner ? { ...plainProps(readSectionProps(ctx, inner)) } : {};
    // The record does not name the headers and footers.
    if (props.headers !== undefined) out.headers = props.headers;
    if (props.footers !== undefined) out.footers = props.footers;
    return out;
  }

  private tableProps(props: TableProps, verdict: Verdict): TableProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', {});
    const ctx = this.ctx();
    return plainProps(readTableProps(ctx, childOf(ctx, change, 'tblPr') || childOf(ctx, change, 'tblPrEx')));
  }

  private rowProps(props: RowProps, verdict: Verdict): RowProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', {});
    const ctx = this.ctx();
    const out: RowProps = { ...plainProps(readRowProps(ctx, childOf(ctx, change, 'trPr'))) };
    if (props.inserted !== undefined) out.inserted = props.inserted;
    if (props.deleted !== undefined) out.deleted = props.deleted;
    // The grid columns the row leaves out are its place in the table, which other changes made what it is (a column
    // put in or taken out since the record was made): no formatting, and not what the record brings back.
    for (const key of ['gridBefore', 'gridAfter'] as const) {
      if (props[key] !== undefined) out[key] = props[key];
      else delete out[key];
    }
    return out;
  }

  private cellProps(props: CellProps, verdict: Verdict): CellProps {
    const change = props.change;
    if (!change) return props;
    if (verdict === 'accept') return without(props, 'change', {});
    const ctx = this.ctx();
    const out: CellProps = { ...plainProps(readCellProps(ctx, childOf(ctx, change, 'tcPr'))) };
    // What says that the cell itself is a change is no formatting: it stays.
    if (props.inserted !== undefined) out.inserted = props.inserted;
    if (props.deleted !== undefined) out.deleted = props.deleted;
    // Nor are the grid columns and the rows the cell covers: its place in the table, as it is now.
    for (const key of ['gridSpan', 'vMerge', 'hMerge'] as const) {
      if (props[key] !== undefined) (out as Record<string, unknown>)[key] = props[key];
      else delete out[key];
    }
    return out;
  }

  // ----------------------------------------------------------- paragraphs

  private paragraph(paragraph: Paragraph): ResolvedParagraph {
    const slots = this.decisions.get(paragraph);
    if (!slots) return { paragraph, join: false };
    const refused = this.refused.get(paragraph);
    const waiting = this.deferred.get(paragraph);
    const verdict = (slot: Slot): Verdict | undefined => ((refused && refused.has(slot)) || (waiting && waiting.has(slot)) ? undefined : slots.get(slot));

    // The properties: the formatting changes, and what the mark says of itself.
    let props = paragraph.props;
    let join = false;
    const format = verdict('paraFormat');
    if (format) props = this.paraProps(props, format);
    let mark = props.markProps;
    if (mark) {
      const before = mark;
      const markFormat = verdict('markFormat');
      if (markFormat) mark = this.runProps(mark, markFormat);
      for (const flag of MARK_FLAGS) {
        const decided = verdict('mark:' + flag);
        if (!decided || mark[flag] === undefined) continue;
        if (removes(FLAG_KIND[flag], decided)) join = true;
        mark = without(mark, flag, EMPTY_RUN_PROPS);
      }
      if (mark !== before) {
        props = { ...props };
        if (Object.keys(mark).length) props.markProps = mark;
        else delete props.markProps;
        if (!Object.keys(props).length) props = EMPTY_PARA_PROPS;
      }
    }
    const section = verdict('section');
    if (section && props.section) props = { ...props, section: this.sectionProps(props.section, section) };

    // The content. What each wrapper comes to: its content goes, or it goes and its content stays.
    const acts = new Map<Wrapper, 'remove' | 'unwrap' | undefined>();
    const act = (w: Wrapper): 'remove' | 'unwrap' | undefined => {
      if (acts.has(w)) return acts.get(w);
      let out: 'remove' | 'unwrap' | undefined;
      const decided = isRevision(w) ? verdict(w) : undefined;
      if (decided && !inAlternate(w)) out = removes(WRAPPER_KIND[w.kind], decided) ? 'remove' : 'unwrap';
      acts.set(w, out);
      return out;
    };
    /** The outermost wrapper around an inline whose content goes. */
    const removedBy = (inline: Inline): Wrapper | undefined => {
      let found: Wrapper | undefined;
      let steps = 0;
      for (let w = inline.wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) if (act(w) === 'remove') found = w;
      return found;
    };

    const stripMemo = new Map<Wrapper, Wrapper | undefined>();
    /** A wrapper chain without the wrappers that go. */
    const strip = (wrap: Wrapper | undefined, depth = 0): Wrapper | undefined => {
      if (!wrap || depth > MAX_CHAIN) return wrap;
      if (stripMemo.has(wrap)) return stripMemo.get(wrap);
      const parent = strip(wrap.parent, depth + 1);
      let out: Wrapper | undefined;
      if (act(wrap) === 'unwrap') out = parent;
      else if (parent === wrap.parent) out = wrap;
      else {
        out = { ...wrap } as Wrapper;
        if (parent) out.parent = parent;
        else delete out.parent;
      }
      stripMemo.set(wrap, out);
      return out;
    };
    const hasDeletion = (wrap: Wrapper | undefined): boolean => {
      let steps = 0;
      for (let w = wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) if (w.kind === 'del') return true;
      return false;
    };

    const out: Inline[] = [];
    /** Wrappers around content that went, with where it stood: one that is left with nothing gets a stand-in, or goes. */
    const emptied: { wrapper: Wrapper; at: number }[] = [];
    const noted = new Set<Wrapper>();
    let changed = props !== paragraph.props;
    const placed = (inline: Inline, wrap: Wrapper | undefined): Inline => {
      const next = { ...inline, wrap } as Inline;
      if (!wrap) delete next.wrap;
      return next;
    };
    for (const inline of paragraph.children) {
      const remover = removedBy(inline);
      if (remover) {
        changed = true;
        const outside = strip(remover.parent);
        if (SURVIVORS.has(inline.type)) out.push(placed(inline, outside));
        else if (inline.type === 'run') {
          const kept = inline.content.filter((item) => item.type === 'commentReference');
          if (kept.length) {
            const run: Run = { type: 'run', props: inline.props, content: kept };
            if (outside) run.wrap = outside;
            out.push(run);
          }
        }
        let steps = 0;
        for (let w = remover.parent; w && steps < MAX_CHAIN; w = w.parent, steps++) {
          if (noted.has(w)) continue;
          noted.add(w);
          emptied.push({ wrapper: w, at: out.length });
        }
        continue;
      }
      // The range marks of a move that is resolved.
      if (MOVE_RANGE.has(inline.type) && verdict(inline)) {
        changed = true;
        continue;
      }
      const wrap = strip(inline.wrap);
      let next = inline;
      if (wrap !== inline.wrap) {
        // A stand-in for an empty wrapper that is gone stands for nothing.
        if (inline.type === 'empty' && (!wrap || !inline.wrap || act(inline.wrap) === 'unwrap')) {
          changed = true;
          continue;
        }
        next = placed(inline, wrap);
        if (next.type === 'run' && !hasDeletion(wrap)) {
          const content = revived(next.content);
          if (content !== next.content) next = { ...next, content };
        }
      }
      if (next.type === 'run') {
        const decided = verdict(inline);
        if (decided && next.props.change) next = { ...next, props: this.runProps(next.props, decided) };
      }
      if (next !== inline) changed = true;
      out.push(next);
    }
    if (!changed) return { paragraph, join };

    // Wrappers that lost everything they held.
    const carries = (inline: Inline, wrapper: Wrapper): boolean => {
      let steps = 0;
      for (let w = inline.wrap; w && steps < MAX_CHAIN; w = w.parent, steps++) if (w === wrapper) return true;
      return false;
    };
    const standIns: { at: number; inline: Inline }[] = [];
    for (const entry of emptied) {
      const kind = entry.wrapper.kind;
      if (act(entry.wrapper) === 'unwrap') continue;
      const now = strip(entry.wrapper) as Wrapper;
      if (kind === 'hyperlink') {
        // A link without anything to show goes; what it still holds stays in its place.
        let used = false;
        for (const inline of out) {
          if (!carries(inline, now)) continue;
          if (inline.type === 'math' || (inline.type === 'run' && inline.content.some((item) => !HINTS.has(item.type) && item.type !== 'commentReference'))) used = true;
        }
        if (used) continue;
        const memo = new Map<Wrapper, Wrapper | undefined>();
        for (let i = 0; i < out.length; i++) if (carries(out[i], now)) out[i] = placed(out[i], withoutWrapper(out[i].wrap, now, memo));
        continue;
      }
      // An insertion or deletion nobody decided about that holds nothing any more is nothing.
      if (isRevision(entry.wrapper) || kind === 'alternate') continue;
      if (out.some((inline) => carries(inline, now)) || standIns.some((stand) => carries(stand.inline, now))) continue;
      standIns.push({ at: entry.at, inline: { type: 'empty', wrap: now } });
    }
    // Last first, so that the places counted before still hold and stand-ins for one place keep their order.
    for (let i = standIns.length - 1; i >= 0; i--) out.splice(standIns[i].at, 0, standIns[i].inline);

    const shaped = paragraphWith(paragraph, out, props);
    return { paragraph: paragraphWith(paragraph, relinked(rebuild(flatten(shaped))), props), join };
  }

  // --------------------------------------------------------------- tables

  /** What stays of a row, or of a cell, that goes: the marks of bookmarks, comments and permissions, without what wrapped them. */
  private leftOf(row: TableRow | undefined, cell?: TableCell): Leftover | undefined {
    const harvest: Harvest = { pieces: [], blocks: [] };
    try {
      if (row) harvestRow(row, harvest);
      else if (cell) {
        if (cell.wrap && inAlternate(cell.wrap)) return undefined;
        harvestBlocks(cell.blocks, harvest);
      }
    } catch (error) {
      if (error instanceof RangeError) return undefined;
      throw error;
    }
    return this.survivors(harvest);
  }

  private survivors(harvest: { pieces: Piece[]; blocks: Block[] }): Leftover {
    const inlines: Inline[] = [];
    for (const piece of harvest.pieces) {
      if (piece.inline) {
        if (SURVIVORS.has(piece.inline.type)) {
          const kept = { ...piece.inline } as Inline;
          delete kept.wrap;
          inlines.push(kept);
        } else if (piece.inline.type === 'run') {
          const kept = piece.inline.content.filter((item) => item.type === 'commentReference');
          if (kept.length) inlines.push({ type: 'run', props: EMPTY_RUN_PROPS, content: kept });
        }
      } else if (piece.item && piece.item.type === 'commentReference') inlines.push({ type: 'run', props: EMPTY_RUN_PROPS, content: [piece.item] });
    }
    const blocks: Block[] = [];
    for (const block of harvest.blocks) {
      if (block.type !== 'mark' || !SURVIVORS.has(block.mark.type)) continue;
      const kept = { ...block };
      delete kept.wrap;
      delete kept.srcStart;
      delete kept.srcEnd;
      blocks.push(kept);
    }
    return { inlines, blocks };
  }

  /** The table as the decisions leave it; undefined when its last row goes, with what it leaves behind in `left`. */
  private table(table: Table, depth: number, left: Leftover): Table | undefined {
    if (!this.touched.has(table)) return table;
    let props = table.props;
    const own = this.decisions.get(table);
    const format = own ? own.get('tableFormat') : undefined;
    if (format) props = this.tableProps(props, format);
    const gridVerdict = own && table.gridChange ? own.get('grid') : undefined;

    const rows: TableRow[] = [];
    /** For each row that stays: the grid columns of the cells it lost. */
    const lost: Set<number>[] = [];
    let anyLost = false;
    /** The cells that go, with the decision that takes each: refused together when they are not a whole column. */
    const taken: { cell: TableCell; slot: string }[] = [];
    /** A column that was put in was decided about: the record of the grid goes with it, and brings no grid back. */
    let columnDecided = false;
    let changed = props !== table.props || gridVerdict !== undefined;
    /** Grid columns in which a vertically merged cell began in a row that went. */
    const heads = new Set<number>();
    let pending: Leftover = { inlines: [], blocks: [] };
    const add = (target: Leftover, more: Leftover): void => {
      for (const inline of more.inlines) target.inlines.push(inline);
      for (const block of more.blocks) target.blocks.push(block);
    };
    for (const row of table.rows) {
      const slots = this.decisions.get(row);
      const refused = this.refused.get(row);
      const verdict = (slot: string): Verdict | undefined => (slots && !(refused && refused.has(slot)) ? slots.get(slot) : undefined);
      let rowProps = row.props;
      let goes = false;
      for (const flag of ['inserted', 'deleted'] as const) {
        const decided = verdict('row:' + flag);
        if (!decided || rowProps[flag] === undefined) continue;
        if (removes(FLAG_KIND[flag], decided)) goes = true;
        rowProps = without(rowProps, flag, {});
      }
      if (goes) {
        const leftover = this.leftOf(row);
        if (!leftover) {
          // Content that is written back whole cannot be taken apart: the row stays as it is.
          for (const flag of ['inserted', 'deleted'] as const) if (slots && slots.has('row:' + flag)) this.refuse(row, 'row:' + flag);
          rows.push(row);
          lost.push(new Set<number>());
          continue;
        }
        add(pending, leftover);
        const spans = rowSpans(row);
        for (let c = 0; c < row.cells.length; c++) {
          const merge = row.cells[c].props.vMerge;
          if (merge === 'restart') heads.add(spans[c][0]);
          else if (merge !== 'continue') heads.delete(spans[c][0]);
        }
        changed = true;
        continue;
      }
      const rowFormat = verdict('rowFormat');
      if (rowFormat) rowProps = this.rowProps(rowProps, rowFormat);
      let tableProps = row.tableProps;
      const exFormat = verdict('rowTableFormat');
      if (exFormat && tableProps) tableProps = this.tableProps(tableProps, exFormat);

      let cells: TableCell[] | undefined;
      const spans = heads.size ? rowSpans(row) : undefined;
      /** The cells of the row that go (with their column), and what each leaves behind. */
      let gone: Map<number, Leftover> | undefined;
      for (let c = 0; c < row.cells.length; c++) {
        const cell = row.cells[c];
        let cellProps = cell.props;
        const cellSlots = this.decisions.get(cell);
        const cellRefused = this.refused.get(cell);
        const cellFormat = cellSlots ? cellSlots.get('cellFormat') : undefined;
        if (cellFormat) cellProps = this.cellProps(cellProps, cellFormat);
        let cellGoes = false;
        for (const flag of ['inserted', 'deleted'] as const) {
          const decided = cellSlots && !(cellRefused && cellRefused.has('cell:' + flag)) ? cellSlots.get('cell:' + flag) : undefined;
          if (!decided || cellProps[flag] === undefined) continue;
          if (flag === 'inserted') columnDecided = true;
          if (removes(FLAG_KIND[flag], decided)) {
            cellGoes = true;
            taken.push({ cell, slot: 'cell:' + flag });
          }
          cellProps = without(cellProps, flag, {});
        }
        if (cellGoes) {
          const leftover = this.leftOf(undefined, cell);
          if (leftover) {
            (gone || (gone = new Map<number, Leftover>())).set(c, leftover);
            continue;
          }
          // Content that is written back whole cannot be taken apart: the cell stays as it is.
          for (const flag of ['inserted', 'deleted'] as const) if (cellSlots && cellSlots.has('cell:' + flag)) this.refuse(cell, 'cell:' + flag);
          continue;
        }
        if (spans && heads.has(spans[c][0]) && cellProps.vMerge === 'continue') cellProps = { ...cellProps, vMerge: 'restart' };
        const blocks = this.list(cell.blocks, depth + 1);
        if (cellProps === cell.props && blocks === cell.blocks) continue;
        if (!cells) cells = row.cells.slice();
        cells[c] = { ...cell, props: cellProps, blocks };
      }
      heads.clear();
      const columns = new Set<number>();
      if (gone) {
        const all = rowSpans(row);
        const kept: TableCell[] = [];
        const source = cells || row.cells;
        /** What cells that went left behind, for the next cell that stays. */
        let carry: Leftover = { inlines: [], blocks: [] };
        for (let c = 0; c < source.length; c++) {
          const leftover = gone.get(c);
          if (leftover) {
            add(carry, leftover);
            for (let g = all[c][0]; g < all[c][1]; g++) columns.add(g);
            continue;
          }
          kept.push(carry.inlines.length || carry.blocks.length ? cellWith(source[c], carry, false) : source[c]);
          carry = { inlines: [], blocks: [] };
        }
        if (kept.length === 0) {
          // A row goes with its last cell.
          add(pending, carry);
          changed = true;
          continue;
        }
        if (carry.inlines.length || carry.blocks.length) kept[kept.length - 1] = cellWith(kept[kept.length - 1], carry, true);
        cells = kept;
        anyLost = true;
      }
      lost.push(columns);
      let next = row;
      if (cells || rowProps !== row.props || tableProps !== row.tableProps) {
        next = { ...row, props: rowProps, cells: cells || row.cells };
        if (tableProps) next.tableProps = tableProps;
      }
      // What the rows that went before this one left behind goes to its start.
      if (pending.inlines.length || pending.blocks.length) {
        next = rowWith(next, pending, false);
        pending = { inlines: [], blocks: [] };
      }
      if (next !== row) changed = true;
      rows.push(next);
    }
    if (!changed) return table;
    if (rows.length === 0) {
      add(left, pending);
      return undefined;
    }
    if (pending.inlines.length || pending.blocks.length) rows[rows.length - 1] = rowWith(rows[rows.length - 1], pending, true);
    let grid = table.grid;
    if (anyLost) {
      // The cells that go are whole columns (see Lister.columnsOf): every row that stays lost the same grid columns.
      // Those columns leave the grid, and a table of a fixed width is that much narrower.
      const columns: number[] = [];
      lost[0].forEach((column) => columns.push(column));
      let whole = true;
      for (const other of lost) if (other.size !== columns.length || !columns.every((column) => other.has(column))) whole = false;
      if (!whole) {
        // Never a table whose rows cover different columns: decisions that would leave one are not taken at all.
        for (const one of taken) this.refuse(one.cell, one.slot);
        return table;
      }
      columns.sort((x, y) => y - x);
      grid = grid.slice();
      let width = 0;
      for (const column of columns) {
        if (column >= grid.length) continue;
        width += grid[column] || 0;
        grid.splice(column, 1);
      }
      if (width > 0 && props.width && props.width.type === 'dxa') props = { ...props, width: { ...props.width, value: Math.max(0, props.width.value - width) } };
    }
    const out: Table = { ...table, props, grid, rows };
    if (gridVerdict && table.gridChange) {
      // The record alone, rejected: the grid is what it holds, when that fits the rows. With the column it was made
      // for, the column leaving the grid is what brings the grid back, whatever order columns are resolved in.
      if (gridVerdict === 'reject' && !columnDecided && !anyLost && table.gridChange.grid.length === grid.length) out.grid = table.gridChange.grid.slice();
      delete out.gridChange;
    }
    delete out.srcStart;
    delete out.srcEnd;
    return out;
  }

  // ---------------------------------------------------------------- lists

  /** Does a paragraph hold nothing but content the verdict would take, were it given to the changes that content is in. */
  private empties(paragraph: Paragraph): boolean {
    for (const inline of paragraph.children) {
      if (SURVIVORS.has(inline.type)) return false;
      if (inline.type === 'run' && inline.content.some((item) => item.type === 'commentReference')) return false;
      let goes = false;
      let steps = 0;
      for (let w = inline.wrap; w && steps < MAX_CHAIN && !goes; w = w.parent, steps++) goes = isRevision(w) && !inAlternate(w) && removes(WRAPPER_KIND[w.kind], this.verdict);
      if (!goes) return false;
    }
    return true;
  }

  /**
   * A list of blocks as the decisions leave it; the list itself when nothing in it changes.
   *
   * Where a paragraph mark goes and what happens to its paragraph depends on a neighbour that is itself a change
   * of the same kind not decided about yet (a table all of which would go, a paragraph whose mark would), the
   * mark waits (see `deferred`): it stays a change, and the rest of what was decided is done. So resolving
   * changes one by one, in any order, leaves what resolving them together leaves.
   */
  list(blocks: Block[], depth = 0): Block[] {
    if (depth > MAX_DEPTH) return blocks;
    const out: Block[] = [];
    let changed = false;
    /** The paragraphs joined so far to the one that comes next, and the blocks without content between them. */
    let open: { paragraph: Paragraph; loose: Block[]; sections: number } | undefined;
    const lastContent = (): Block | undefined => {
      for (let i = out.length - 1, tries = 0; i >= 0 && tries < 64; i--, tries++) if (out[i].type === 'paragraph' || out[i].type === 'table') return out[i];
      return undefined;
    };
    const isFrozen = (block: Block): boolean => block.wrap !== undefined && (block.wrap.kind === 'alternate' || inAlternate(block.wrap));
    /** Has `other` a mark that the verdict which takes the mark of `paragraph` would take too. */
    const waits = (other: Paragraph, paragraph: Paragraph): boolean => {
      const slots = this.decisions.get(paragraph);
      const mark = paragraph.props.markProps;
      const pending = other.props.markProps;
      if (!slots || !mark || !pending) return false;
      for (const flag of MARK_FLAGS) {
        const decided = slots.get('mark:' + flag);
        if (!decided || mark[flag] === undefined || !removes(FLAG_KIND[flag], decided)) continue;
        for (const one of MARK_FLAGS) if (pending[one] !== undefined && removes(FLAG_KIND[one], decided)) return true;
      }
      return false;
    };
    /**
     * Would the verdict that takes the mark of a paragraph take all of the table too, were it given to the changes of
     * the table: every row is one it takes (inserted, for a rejection; deleted, for an acceptance), or a row all of
     * whose cells are (the columns of the table).
     */
    const goesWith = (table: Table, paragraph: Paragraph): boolean => {
      const slots = this.decisions.get(paragraph);
      const mark = paragraph.props.markProps;
      if (!slots || !mark || table.rows.length === 0) return false;
      for (const flag of ['inserted', 'deleted'] as const) {
        const decided = slots.get('mark:' + flag);
        if (!decided || mark[flag] === undefined || !removes(FLAG_KIND[flag], decided)) continue;
        const taken = (props: { inserted?: RevisionInfo; deleted?: RevisionInfo }): boolean =>
          (props.inserted !== undefined && removes('insert', decided)) || (props.deleted !== undefined && removes('delete', decided));
        if (table.rows.every((row) => taken(row.props) || (row.cells.length > 0 && row.cells.every((cell) => taken(cell.props))))) return true;
      }
      return false;
    };
    /** What the tables of the list come to, each resolved once: a paragraph whose mark goes looks ahead past a table that goes. */
    const tables = new Map<number, { table: Table | undefined; left: Leftover }>();
    const tableAt = (index: number): { table: Table | undefined; left: Leftover } => {
      let known = tables.get(index);
      if (!known) {
        const left: Leftover = { inlines: [], blocks: [] };
        known = { table: this.table(blocks[index] as Table, depth, left), left };
        tables.set(index, known);
      }
      return known;
    };
    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const frozen = block.wrap !== undefined && (block.wrap.kind === 'alternate' || inAlternate(block.wrap));
      if (block.type === 'paragraph' && !frozen) {
        const resolved = this.paragraph(block);
        let paragraph = resolved.paragraph;
        // (As things were, should the mark of this paragraph turn out to have to wait.)
        const held = open;
        const sections = this.sectionsRemoved;
        const wait = (): void => {
          this.deferMarks(block);
          open = held;
          this.sectionsRemoved = sections;
          i--;
        };
        if (open) {
          const head = flatten(open.paragraph);
          const tail = flatten(paragraph);
          for (const piece of tail) head.push(piece);
          // The mark that stays is the one of this paragraph, and with it everything the paragraph is: its formatting,
          // its list, the record of a change of them, the section it ends. What the paragraphs before it were went
          // with their marks (Word's rule, and one that does not depend on the order changes are resolved in).
          if (open.sections) this.sectionsRemoved += open.sections;
          paragraph = paragraphWith(paragraph, relinked(rebuild(head)), paragraph.props);
        }
        const loose = open ? open.loose : [];
        open = undefined;
        if (resolved.join) {
          // A table that goes whole (its rows were inserted with the paragraph, or deleted with it) is not in the way.
          let after = neighbour(blocks, i, 1);
          for (let tries = 0; after >= 0 && tries < 4096; tries++) {
            const ahead = blocks[after];
            if (ahead.type !== 'table' || isFrozen(ahead)) break;
            const came = tableAt(after);
            if (came.table !== undefined || came.left.inlines.length) break;
            after = neighbour(blocks, after, 1);
          }
          const next = after >= 0 ? blocks[after] : undefined;
          // A table that would go whole with the same verdict, and is not decided about yet, stands between the mark
          // and the paragraph it has to join. Alone, the mark would stay as a plain mark and the two paragraphs two for
          // good, whatever happens to the table later: the mark waits for the table (or goes with it, in "all").
          // (Not for a table that goes now and leaves marks behind: those get a paragraph of their own in its place.)
          if (next && next.type === 'table' && !isFrozen(next) && tableAt(after).table !== undefined && goesWith(next, block)) {
            wait();
            continue;
          }
          if (next && next.type === 'paragraph' && !(next.wrap !== undefined && (next.wrap.kind === 'alternate' || inAlternate(next.wrap)))) {
            // The section break this paragraph carries goes with its mark; it is counted when the join is made.
            open = { paragraph, loose, sections: paragraph.props.section !== undefined ? 1 : 0 };
            changed = true;
            continue;
          }
          // No paragraph to join: an empty one goes where the list can do without it; any other keeps its mark, plain.
          const before = lastContent();
          const tableBefore = before !== undefined && before.type === 'table';
          const tableAfter = next !== undefined && next.type === 'table';
          // And the table before it may be one that would go with the same verdict, and is not decided about yet: a
          // paragraph that has to stay only because of that table (a table does not end a list) waits for it. So
          // does one before which stands a paragraph whose own mark would go the same way: once that goes, the two
          // are one paragraph, and whether that one stays is decided then. (Whether the paragraph is empty by then
          // depends on changes in it that may be pending too: it waits either way.)
          if (before !== undefined && ((before.type === 'table' && goesWith(before, block)) || (before.type === 'paragraph' && waits(before, block)))) {
            wait();
            continue;
          }
          const droppable = next !== undefined ? !(tableAfter && tableBefore) : before !== undefined && !tableBefore;
          // A paragraph that would go were it empty, and holds nothing but what the same verdict would take (changes
          // not decided about yet), waits for those: it goes with the last of them, as it does when all are resolved
          // together.
          if (droppable && paragraph.children.length > 0 && this.empties(paragraph)) {
            wait();
            continue;
          }
          if (paragraph.children.length === 0 && droppable) {
            if (paragraph.props.section !== undefined) this.sectionsRemoved++;
            changed = true;
            for (const kept of loose) out.push(kept);
            continue;
          }
        }
        if (paragraph !== block) changed = true;
        out.push(paragraph);
        for (const kept of loose) out.push(kept);
        continue;
      }
      if (block.type === 'table' && !frozen) {
        const { table, left } = tableAt(i);
        if (table === block) {
          out.push(block);
          continue;
        }
        changed = true;
        if (table) {
          out.push(table);
          continue;
        }
        if (open) {
          // The paragraph before it is being joined to the one after it: the table leaves nothing but its marks.
          for (const kept of left.blocks) open.loose.push(kept);
          continue;
        }
        // The table is gone. A paragraph takes its place where the list would otherwise be empty, end with a table or
        // have two tables touch, and where there is something the table left behind.
        const before = lastContent();
        const after = neighbour(blocks, i, 1);
        const tableBefore = before !== undefined && before.type === 'table';
        const tableAfter = after >= 0 && blocks[after].type === 'table';
        const needed = after >= 0 ? tableAfter && tableBefore : before === undefined || tableBefore;
        for (const kept of left.blocks) out.push(kept);
        if (needed || left.inlines.length) {
          const filler: Paragraph = { type: 'paragraph', props: EMPTY_PARA_PROPS, children: left.inlines };
          if (block.wrap) filler.wrap = block.wrap;
          out.push(filler);
        }
        continue;
      }
      // The range mark of a move that is resolved, between two blocks.
      if (block.type === 'mark' && this.decisions.has(block)) {
        changed = true;
        continue;
      }
      if (open) open.loose.push(block);
      else out.push(block);
    }
    return changed ? out : blocks;
  }
}

// ------------------------------------------------------------------ commands

interface Plan {
  /** The new blocks of each story that changes. */
  stories: { id: string; blocks: Block[] }[];
  section?: SectionProps;
  chosen: Entry[];
  sectionsRemoved: number;
}

/** The stories of a document that hold changes that can be resolved: the ids the lister gave them. */
function plan(doc: DocModel, found: Collected, wanted: Entry[], verdict: Verdict): Plan {
  let chosen = wanted.slice();
  // Every pass leaves out the changes a part of which was refused; with none refused the result stands.
  for (let pass = 0; pass < 32; pass++) {
    const decisions = new Map<object, Map<Slot, Verdict>>();
    const tables = new Set<Table>();
    const decide = (part: Part): void => {
      let slots = decisions.get(part.host);
      if (!slots) decisions.set(part.host, (slots = new Map<Slot, Verdict>()));
      slots.set(part.slot, verdict);
      for (const table of part.tables) tables.add(table);
    };
    const stories = new Set<string>();
    const moves = new Set<string>();
    for (const entry of chosen) {
      for (const part of entry.parts) decide(part);
      stories.add(entry.rev.story);
      if (entry.rev.move !== undefined) moves.add(entry.rev.move);
    }
    // The range marks of the moves; they may stand in another story than the ends do, so every story is looked at.
    moves.forEach((name) => {
      for (const part of found.ranges.get(name) || []) decide(part);
    });
    const everywhere = moves.size > 0;
    const resolver = new Resolver(decisions, tables, verdict);
    const out: Plan = { stories: [], chosen, sectionsRemoved: 0 };
    const ids: string[] = ['body'];
    doc.headers.forEach((_story, id) => ids.push('header:' + id));
    doc.footers.forEach((_story, id) => ids.push('footer:' + id));
    for (const [kind, collection] of [
      ['footnote', doc.footnotes],
      ['endnote', doc.endnotes],
    ] as const) {
      if (collection) collection.items.forEach((note, id) => (note.kind === undefined && /^-?\d{1,9}$/.test(String(id)) ? ids.push(`${kind}:${id}`) : 0));
    }
    const done = new Set<object>();
    for (const id of ids) {
      if (!everywhere && !stories.has(id)) continue;
      const handle = storyOf(doc, id);
      const key = handle.note || handle.host;
      if (done.has(key)) continue;
      done.add(key);
      resolver.enter(handle.host, handle.blocks);
      const blocks = resolver.list(handle.blocks);
      if (blocks !== handle.blocks) out.stories.push({ id, blocks });
    }
    const bodySlots = decisions.get(doc.body);
    if (bodySlots && bodySlots.has('section') && doc.body.section) {
      resolver.enter(doc.body);
      out.section = resolver.sectionProps(doc.body.section, verdict);
    }
    out.sectionsRemoved = resolver.sectionsRemoved;
    // A change that is nothing but marks that wait was not resolved: it is not said to be.
    if (resolver.deferred.size) {
      out.chosen = chosen.filter((entry) => !entry.parts.every((part) => (resolver.deferred.get(part.host) || new Set<Slot>()).has(part.slot)));
    }
    if (resolver.refused.size === 0) return out;
    const refusedMoves = new Set<string>();
    const hit = (entry: Entry): boolean => entry.parts.some((part) => (resolver.refused.get(part.host) || new Set<Slot>()).has(part.slot));
    for (const entry of chosen) if (entry.rev.move !== undefined && hit(entry)) refusedMoves.add(entry.rev.move);
    const next = chosen.filter((entry) => !hit(entry) && !(entry.rev.move !== undefined && refusedMoves.has(entry.rev.move)));
    if (next.length === chosen.length) break;
    chosen = next;
  }
  return { stories: [], chosen: [], sectionsRemoved: 0 };
}

function apply(draft: Draft, made: Plan): void {
  const doc = draft.doc;
  for (const story of made.stories) {
    const old = storyOf(doc, story.id).blocks.slice();
    const blocks = story.blocks;
    // Only the stretches that differ are replaced: what stands between two changes (and may be content that cannot
    // be replaced at all) is not touched. A block of the old list that the new one has too, further on, ends a stretch.
    const place = new Map<Block, number>();
    for (let j = 0; j < blocks.length; j++) place.set(blocks[j], j);
    const stretches: { at: number; remove: number; insert: Block[] }[] = [];
    let i = 0;
    let j = 0;
    while (i < old.length || j < blocks.length) {
      if (i < old.length && j < blocks.length && old[i] === blocks[j]) {
        i++;
        j++;
        continue;
      }
      let end = i;
      let to = blocks.length;
      for (; end < old.length; end++) {
        const found = place.get(old[end]);
        if (found !== undefined && found >= j) {
          to = found;
          break;
        }
      }
      stretches.push({ at: i, remove: end - i, insert: blocks.slice(j, to) });
      i = end;
      j = to;
    }
    // Last first, so that the places of the earlier ones still hold.
    for (let k = stretches.length - 1; k >= 0; k--) draft.replace(story.id, [], stretches[k].at, stretches[k].remove, stretches[k].insert);
  }
  if (made.section) draft.apply({ type: 'setBodySection', section: made.section });
}

/** A position that names a place in the document as it is now: the one asked for, or the nearest there is. */
function settled(doc: DocModel, wanted: DocPosition): DocPosition {
  try {
    const at = resolvePosition(doc, { story: wanted.story, path: wanted.path, offset: 0 });
    const offset = Math.max(0, Math.min(wanted.offset, at.length));
    return wanted.before === true && offset === wanted.offset ? { story: wanted.story, path: wanted.path.slice(), offset, before: true } : { story: wanted.story, path: wanted.path.slice(), offset };
  } catch (error) {
    if (!(error instanceof RangeError)) throw error;
  }
  for (const story of [wanted.story, 'body']) {
    let blocks: Block[];
    try {
      blocks = storyOf(doc, story).blocks;
    } catch (error) {
      continue;
    }
    const index = Math.max(0, Math.min(typeof wanted.path[0] === 'number' && story === wanted.story ? wanted.path[0] : 0, blocks.length - 1));
    const path = paragraphIn(blocks, [], index, 1) || paragraphIn(blocks, [], index, -1);
    if (!path) continue;
    const after = path[0] < index;
    return { story, path, offset: after ? resolvePosition(doc, { story, path, offset: 0 }).length : 0 };
  }
  throw new RangeError('The story holds no paragraph');
}

function guard(doc: DocModel): void {
  if (enforcedProtection(doc) !== undefined) throw new RangeError(REVISIONS_PROTECTED);
}

const samePlace = (a: DocPosition, b: DocPosition | undefined): boolean =>
  b !== undefined && b !== null && a.story === b.story && a.offset === b.offset && Array.isArray(b.path) && a.path.length === b.path.length && a.path.every((value, i) => value === b.path[i]);

/** The entry of a fresh listing that is the change `revision` names. */
function locate(found: Collected, revision: DocRevision): Entry {
  if (!revision || typeof revision !== 'object' || typeof revision.kind !== 'string') throw new RangeError('A change of the list of changes is needed');
  const known = PARTS.get(revision);
  if (!known) {
    // A description that did not come from a listing (it was copied, or sent over a bridge): found by what it says.
    const same = found.entries.find(
      (entry) =>
        entry.rev.index === revision.index &&
        entry.rev.kind === revision.kind &&
        entry.rev.author === revision.author &&
        entry.rev.date === revision.date &&
        entry.rev.text === revision.text &&
        entry.rev.story === revision.story &&
        samePlace(entry.rev.start, revision.start) &&
        samePlace(entry.rev.end, revision.end),
    );
    if (!same) throw new RangeError(REVISION_GONE);
    return same;
  }
  if (known.length === 0) {
    const same = found.entries.find((entry) => entry.rev.kind === 'other' && entry.rev.index === revision.index);
    if (!same) throw new RangeError(REVISION_GONE);
    return same;
  }
  for (const entry of found.entries) {
    if (entry.parts.length !== known.length) continue;
    let same = true;
    for (let i = 0; i < known.length && same; i++) same = entry.parts[i].host === known[i].host && entry.parts[i].slot === known[i].slot;
    if (same) return entry;
  }
  throw new RangeError(REVISION_GONE);
}

/** The parts each listed change stands for, by the object that describes it. */
const PARTS = new WeakMap<DocRevision, Part[]>();

function remember(found: Collected): Collected {
  for (const entry of found.entries) PARTS.set(entry.rev, entry.parts);
  return found;
}

function resolveOne(doc: DocModel, revision: DocRevision, verdict: Verdict): RevisionEdit {
  guard(doc);
  const found = collect(doc, Infinity);
  const entry = locate(found, revision);
  if (entry.rev.kind === 'other' || !(verdict === 'reject' ? entry.rev.canReject : entry.rev.canAccept)) throw new RangeError(REVISION_UNSUPPORTED);
  // The other end of a move goes with it.
  const wanted = entry.rev.move === undefined ? [entry] : found.entries.filter((other) => other.rev.move === entry.rev.move && (other.rev.kind === 'moveFrom' || other.rev.kind === 'moveTo'));
  const made = plan(doc, found, wanted, verdict);
  // Nothing is ever said to be resolved that was not: a change that leaves the document as it is was not resolved.
  if (made.chosen.length !== wanted.length || (made.stories.length === 0 && made.section === undefined)) throw new RangeError(REVISION_UNSUPPORTED);
  const stays = !removes(entry.rev.kind, verdict);
  const edit = buildEdit(
    doc,
    (draft) => {
      apply(draft, made);
      // After what was accepted; where what went stood.
      return settled(doc, stays && entry.rev.kind === 'insert' ? entry.rev.end : entry.rev.start);
    },
    undefined,
    true,
  );
  return { ...edit, sectionsRemoved: made.sectionsRemoved };
}

/**
 * Accepts one change of the list: see the top of this file for what that
 * does to each kind. The caret is left at the end of an accepted insertion,
 * and where the change began otherwise. Throws RangeError: REVISION_GONE for
 * a change the document no longer has (the list is older than an edit),
 * REVISION_UNSUPPORTED for one that cannot be resolved, REVISIONS_PROTECTED
 * in a document whose protection is enforced. The document is then unchanged.
 */
export function acceptRevision(doc: DocModel, revision: DocRevision): RevisionEdit {
  return resolveOne(doc, revision, 'accept');
}

/** Rejects one change of the list; see acceptRevision(). */
export function rejectRevision(doc: DocModel, revision: DocRevision): RevisionEdit {
  return resolveOne(doc, revision, 'reject');
}

function resolveAll(doc: DocModel, verdict: Verdict, filter: ((revision: DocRevision) => boolean) | undefined, selection: DocRange | undefined): RevisionsEdit {
  guard(doc);
  if (filter !== undefined && typeof filter !== 'function') throw new RangeError('The filter of the changes must be a function');
  const found = collect(doc, Infinity);
  const asked = filter ? found.entries.filter((entry) => filter(entry.rev)) : found.entries;
  // A move is resolved with both its ends, also when the filter named one of them.
  const moves = new Set<string>();
  for (const entry of asked) if (entry.rev.move !== undefined) moves.add(entry.rev.move);
  const picked = new Set<Entry>(asked);
  const candidates = found.entries.filter((entry) => picked.has(entry) || (entry.rev.move !== undefined && moves.has(entry.rev.move) && (entry.rev.kind === 'moveFrom' || entry.rev.kind === 'moveTo')));
  const wanted = candidates.filter((entry) => entry.rev.kind !== 'other' && (verdict === 'accept' ? entry.rev.canAccept : entry.rev.canReject));
  const made: Plan = wanted.length ? plan(doc, found, wanted, verdict) : { stories: [], chosen: [], sectionsRemoved: 0 };
  const edit = buildEdit(
    doc,
    (draft) => {
      apply(draft, made);
      if (selection) return { anchor: settled(doc, selection.anchor), focus: settled(doc, selection.focus) };
      return settled(doc, { story: 'body', path: [0], offset: 0 });
    },
    undefined,
    true,
  );
  return { ...edit, resolved: made.chosen.length, skipped: candidates.length - made.chosen.length, sectionsRemoved: made.sectionsRemoved };
}

/**
 * Accepts every change of the document, or those `filter` picks. What cannot
 * be resolved is passed over: `skipped` says how many changes that were, and
 * `resolved` how many are resolved. `selection` is the selection to keep
 * (moved to the nearest place when its own is gone); without it the caret is
 * left at the start of the body.
 */
export function acceptAll(doc: DocModel, filter?: (revision: DocRevision) => boolean, selection?: DocRange): RevisionsEdit {
  return resolveAll(doc, 'accept', filter, selection);
}

/** Rejects every change of the document, or those `filter` picks; see acceptAll(). */
export function rejectAll(doc: DocModel, filter?: (revision: DocRevision) => boolean, selection?: DocRange): RevisionsEdit {
  return resolveAll(doc, 'reject', filter, selection);
}

// ------------------------------------------------------------------ tracking

/**
 * Turns the "track changes" setting of the document on or off, as an edit
 * that can be undone. Turning it off is refused with
 * RangeError(TRACK_ENFORCED) where a protection of the document enforces
 * tracked changes, and where the application asked for tracking
 * (setTracking with `force`): what an application turned on, its user does
 * not turn off; nor does turning it on there write the setting of the
 * document, which stays what the file says. No operations come back when
 * the setting is as asked.
 */
export function setTrackRevisions(doc: DocModel, on: boolean, selection?: DocRange): DocEdit {
  if (typeof on !== 'boolean') throw new RangeError('Tracking is either on or off');
  if (!on && (enforcedProtection(doc) === 'trackedChanges' || forcesTracking(doc))) throw new RangeError(TRACK_ENFORCED);
  const at = selection ? { anchor: settled(doc, selection.anchor), focus: settled(doc, selection.focus) } : collapsed(settled(doc, { story: 'body', path: [0], offset: 0 }));
  // Tracking the application asked for is on whatever the document says: its own setting is left as the file has it.
  if ((doc.settings.trackRevisions === true) === on || forcesTracking(doc)) return { ops: [], selection: at };
  return { ops: [{ type: 'setTrackRevisions', on }], selection: at };
}

// ----------------------------------------------- the selection and the changes

/** Two places of one story in reading order: negative when `a` comes first. Paths are compared number by number, which holds for any two paragraphs. */
export function placeOrder(a: DocPosition, b: DocPosition): number {
  const shared = Math.min(a.path.length, b.path.length);
  for (let i = 0; i < shared; i++) if (a.path[i] !== b.path[i]) return a.path[i] - b.path[i];
  if (a.path.length !== b.path.length) return a.path.length - b.path.length;
  return a.offset - b.offset;
}

/**
 * The places in a list of changes of those a selection touches. A selection
 * holds every change it runs over, and those without room (a deletion, a
 * paragraph mark) that stand inside it. A caret is in one change at most:
 * the one it stands inside of; on the border of changes (at the end of one,
 * the start of another, where a deletion stands) the first of them.
 */
export function revisionsTouching(list: DocRevision[], range: DocRange): number[] {
  const [start, end] = placeOrder(range.anchor, range.focus) <= 0 ? [range.anchor, range.focus] : [range.focus, range.anchor];
  const caret = placeOrder(start, end) === 0;
  const out: number[] = [];
  let border = -1;
  for (let i = 0; i < list.length; i++) {
    const revision = list[i];
    if (revision.story !== start.story) continue;
    // A column is where its cells are: one stretch for each row.
    const stretches = revision.cells && revision.cells.length ? revision.cells : [revision];
    let inside = false;
    let beside = false;
    for (const stretch of stretches) {
      const from = placeOrder(stretch.start, end);
      const to = placeOrder(stretch.end, start);
      if (from < 0 && to > 0) inside = true;
      if (from <= 0 && to >= 0) beside = true;
    }
    if (caret) {
      if (inside) return [i];
      if (beside && border < 0) border = i;
    } else if (inside) out.push(i);
  }
  return caret && border >= 0 ? [border] : out;
}

/**
 * The change a step to the next or the previous one arrives at: from the
 * change at `current` (the one stepped to before; -1 for none), else from
 * where the selection is. Around the ends of the list, which `wrapped` says.
 * Undefined for an empty list.
 */
export function revisionStep(list: DocRevision[], range: DocRange | undefined, direction: 'next' | 'previous', current: number): { index: number; wrapped?: 'first' | 'last' } | undefined {
  if (list.length === 0) return undefined;
  let index: number;
  // The changes of the story the selection is in, among those of the stories before and after it.
  let firstOwn = -1;
  let lastOwn = -1;
  let from: DocPosition | undefined;
  let to: DocPosition | undefined;
  if (range && !(current >= 0 && current < list.length)) {
    [from, to] = placeOrder(range.anchor, range.focus) <= 0 ? [range.anchor, range.focus] : [range.focus, range.anchor];
    for (let i = 0; i < list.length; i++) {
      if (list[i].story !== range.focus.story) continue;
      if (firstOwn < 0) firstOwn = i;
      lastOwn = i;
    }
  }
  if (direction === 'next') {
    if (current >= 0 && current < list.length) index = current + 1;
    else if (!to || firstOwn < 0) index = 0;
    else {
      // The first change of the story that begins at or after the end of the selection, else the first one of the stories after it.
      index = lastOwn + 1;
      for (let i = firstOwn; i <= lastOwn; i++) {
        if (list[i].story === to.story && placeOrder(list[i].start, to) >= 0) {
          index = i;
          break;
        }
      }
    }
    return index >= list.length ? { index: 0, wrapped: 'first' } : { index };
  }
  if (current >= 0 && current < list.length) index = current - 1;
  else if (!from || firstOwn < 0) index = -1;
  else {
    index = firstOwn - 1;
    for (let i = lastOwn; i >= firstOwn; i--) {
      if (list[i].story === from.story && placeOrder(list[i].end, from) <= 0) {
        index = i;
        break;
      }
    }
  }
  return index < 0 ? { index: list.length - 1, wrapped: 'last' } : { index };
}
