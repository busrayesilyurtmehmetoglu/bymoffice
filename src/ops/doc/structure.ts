// Structural commands: tables and their rows and columns, hyperlinks,
// pictures.

import {
  EMPTY_PARA_PROPS,
  type Block,
  type Border,
  type CellProps,
  type DocMedia,
  type DocModel,
  type DrawingItem,
  type Paragraph,
  type RowProps,
  type RunItem,
  type Table,
  type TableCell,
  type TableProps,
  type TableRow,
  type Wrapper,
} from '../../model/doc/types';
import type { XmlElement, XmlNode } from '../../xml';
import { buildEdit, type DocEdit, type DocEditOptions, type Draft } from './draft';
import { patchRunProps } from './format';
import { boundary, flatten, groupFor, inheritedProps, itemPiece, offsetOf, paragraphWith, rebuild, textPieces, wrapBetween, type Piece } from './pieces';
import { comparePositions, containerBlocks, inTextBox, paragraphIn, resolvePosition, storyOf, type DocPosition, type DocRange } from './position';
import { placeBlock } from './text';

const bad = (message: string): never => {
  throw new RangeError(message);
};

const whole = (value: unknown, min: number, max: number): value is number => typeof value === 'number' && value >= min && value <= max && Math.floor(value) === value;

/** Most columns a table of Word has. */
const MAX_COLUMNS = 63;
/** Most cells a table made in one step has. */
const MAX_NEW_CELLS = 100000;

const spanOf = (props: CellProps): number => (props.gridSpan !== undefined && props.gridSpan > 1 ? props.gridSpan : 1);

/** Grid columns of the cells of a row, as [start, end) each. */
function rowSpans(row: TableRow): [number, number][] {
  const out: [number, number][] = [];
  let at = row.props.gridBefore || 0;
  for (const cell of row.cells) {
    const span = spanOf(cell.props);
    out.push([at, at + span]);
    at += span;
  }
  return out;
}

/** An empty paragraph formatted like the first paragraph of a cell: what a new cell beside it holds. */
function emptyLike(cell: TableCell | undefined): Paragraph {
  if (cell) {
    for (const block of cell.blocks) {
      if (block.type !== 'paragraph') continue;
      let props = block.props;
      if (props.section) {
        props = { ...props };
        delete props.section;
      }
      return { type: 'paragraph', props, children: [] };
    }
  }
  return { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] };
}

interface CellSite {
  story: string;
  /** Path of the list of blocks the table is in. */
  container: number[];
  index: number;
  table: Table;
  row: number;
  cell: number;
}

/** The innermost table cell a position is in. Throws RangeError when it is in none. */
function cellSite(doc: DocModel, position: DocPosition): CellSite {
  resolvePosition(doc, position);
  const path = position.path;
  if (inTextBox(path)) bad('The content of a text box cannot be edited');
  if (path.length < 4) bad('The position is not in a table');
  const container = path.slice(0, -4);
  const index = path[path.length - 4];
  const table = containerBlocks(doc, position.story, container)[index];
  if (!table || table.type !== 'table') return bad('The position is not in a table');
  return { story: position.story, container, index, table, row: path[path.length - 3], cell: path[path.length - 2] };
}

function tableWith(table: Table, rows: TableRow[], grid = table.grid, props = table.props): Table {
  const out: Table = { ...table, rows, grid, props };
  delete out.srcStart;
  delete out.srcEnd;
  return out;
}

/** Start of the first paragraph of a cell of the table a site names, after the table was replaced. */
function cellStart(doc: DocModel, site: CellSite, row: number, cell: number): DocPosition {
  const base = site.container.concat([site.index, row, cell]);
  const path = paragraphIn(containerBlocks(doc, site.story, base), base, 0, 1);
  if (!path) return bad('The cell holds no paragraph');
  return { story: site.story, path, offset: 0 };
}

const ROW_REVISION = ['inserted', 'deleted', 'change'] as const;

/**
 * Inserts a row above or below the row a position is in. The new row has the
 * cells of that row, empty, with their formatting; where the row is crossed
 * by vertically merged cells, the new cells continue the merge.
 */
export function insertRow(doc: DocModel, position: DocPosition, where: 'above' | 'below', _options?: DocEditOptions): DocEdit {
  if (where !== 'above' && where !== 'below') bad('A row goes above or below');
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    const ref = table.rows[site.row];
    const spans = rowSpans(ref);
    const below = table.rows[site.row + 1];
    const belowSpans = below ? rowSpans(below) : [];
    const cells: TableCell[] = ref.cells.map((cell, c) => {
      const props: CellProps = { ...cell.props };
      delete props.change;
      let continues = false;
      if (where === 'above') continues = cell.props.vMerge === 'continue';
      else if (cell.props.vMerge !== undefined && below) {
        for (let k = 0; k < below.cells.length; k++) if (belowSpans[k][0] === spans[c][0] && below.cells[k].props.vMerge === 'continue') continues = true;
      }
      if (continues) props.vMerge = 'continue';
      else delete props.vMerge;
      return { props, blocks: [emptyLike(cell)] };
    });
    const props: RowProps = { ...ref.props };
    for (const key of ROW_REVISION) delete props[key];
    // Only the rows at the top of a table repeat on every page.
    if (where === 'below') delete props.header;
    const row: TableRow = { props, cells };
    if (ref.tableProps) row.tableProps = ref.tableProps;
    const at = where === 'above' ? site.row : site.row + 1;
    const wrap = wrapBetween(at > 0 ? table.rows[at - 1].wrap : undefined, at < table.rows.length ? table.rows[at].wrap : undefined);
    if (wrap) row.wrap = wrap;
    const rows = table.rows.slice();
    rows.splice(at, 0, row);
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows)]);
    return cellStart(doc, site, at, site.cell);
  });
}

/** Removes a table; an empty paragraph takes its place where the list of blocks would otherwise end with a table, be empty, or have two tables touch. */
function removeTable(draft: Draft, site: CellSite): DocPosition {
  const blocks = containerBlocks(draft.doc, site.story, site.container);
  const near = (step: 1 | -1): Block | undefined => {
    for (let i = site.index + step, tries = 0; i >= 0 && i < blocks.length && tries < 64; i += step, tries++) {
      if (blocks[i].type === 'paragraph' || blocks[i].type === 'table') return blocks[i];
    }
    return undefined;
  };
  const before = near(-1);
  const after = near(1);
  const needed = after ? after.type === 'table' && before !== undefined && before.type === 'table' : !before || before.type === 'table';
  const filler: Paragraph = { type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] };
  if (site.table.wrap) filler.wrap = site.table.wrap;
  draft.replace(site.story, site.container, site.index, 1, needed ? [filler] : []);
  const now = containerBlocks(draft.doc, site.story, site.container);
  const path = paragraphIn(now, site.container, site.index, 1) || paragraphIn(now, site.container, site.index - 1, -1);
  if (!path) return bad('The story holds no paragraph');
  const end = path[path.length - 1] < site.index;
  return { story: site.story, path, offset: end ? resolvePosition(draft.doc, { story: site.story, path, offset: 0 }).length : 0 };
}

/**
 * Deletes the row a position is in; the last row takes the table with it. A
 * vertically merged cell that began in the row begins in the next row
 * afterwards.
 */
export function deleteRow(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    if (table.rows.length <= 1) return removeTable(draft, site);
    const rows = table.rows.slice();
    const gone = rows.splice(site.row, 1)[0];
    const next = rows[site.row];
    if (next) {
      const spans = rowSpans(gone);
      const nextSpans = rowSpans(next);
      let cells: TableCell[] | undefined;
      for (let c = 0; c < gone.cells.length; c++) {
        if (gone.cells[c].props.vMerge !== 'restart') continue;
        for (let k = 0; k < next.cells.length; k++) {
          if (nextSpans[k][0] !== spans[c][0] || next.cells[k].props.vMerge !== 'continue') continue;
          if (!cells) cells = next.cells.slice();
          cells[k] = { ...next.cells[k], props: { ...next.cells[k].props, vMerge: 'restart' } };
        }
      }
      if (cells) rows[site.row] = { ...next, cells };
    }
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows)]);
    const row = Math.min(site.row, rows.length - 1);
    return cellStart(doc, site, row, Math.min(site.cell, rows[row].cells.length - 1));
  });
}

function widened(props: TableProps, by: number): TableProps {
  if (!props.width || props.width.type !== 'dxa') return props;
  return { ...props, width: { ...props.width, value: Math.max(0, props.width.value + by) } };
}

function cellWidened(props: CellProps, by: number): CellProps {
  if (!props.width || props.width.type !== 'dxa') return props;
  return { ...props, width: { ...props.width, value: Math.max(0, props.width.value + by) } };
}

/**
 * Inserts a column left or right of the cell a position is in, as wide as
 * the column beside it. Every row gets a cell at that place in the grid; a
 * cell that spans the place becomes one column wider instead, which is what
 * Word does with merged cells.
 */
export function insertColumn(doc: DocModel, position: DocPosition, side: 'left' | 'right', _options?: DocEditOptions): DocEdit {
  if (side !== 'left' && side !== 'right') bad('A column goes left or right');
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    if (table.grid.length >= MAX_COLUMNS) bad('The table has no room for another column');
    const refSpans = rowSpans(table.rows[site.row]);
    const [start, end] = refSpans[site.cell];
    const place = side === 'left' ? start : end;
    const beside = side === 'left' ? start : end - 1;
    const refWidth = table.rows[site.row].cells[site.cell].props.width;
    let width = table.grid[beside];
    if (!(width > 0)) width = refWidth && refWidth.type === 'dxa' && refWidth.value > 0 ? Math.round(refWidth.value / (end - start)) : 1440;
    const rows = table.rows.map((row): TableRow => {
      const before = row.props.gridBefore || 0;
      if (place < before) return { ...row, props: { ...row.props, gridBefore: before + 1 } };
      const spans = rowSpans(row);
      const fresh = (model: TableCell): TableCell => {
        const props: CellProps = { ...model.props };
        delete props.gridSpan;
        delete props.change;
        delete props.hMerge;
        if (props.width) props.width = { type: 'dxa', value: width };
        return { props, blocks: [emptyLike(model)] };
      };
      for (let c = 0; c < row.cells.length; c++) {
        const [from, to] = spans[c];
        if (from === place) {
          const cells = row.cells.slice();
          cells.splice(c, 0, fresh(side === 'right' && c > 0 ? row.cells[c - 1] : row.cells[c]));
          return { ...row, cells };
        }
        if (from < place && place < to) {
          const cells = row.cells.slice();
          cells[c] = { ...row.cells[c], props: { ...cellWidened(row.cells[c].props, width), gridSpan: to - from + 1 } };
          return { ...row, cells };
        }
      }
      const rowEnd = spans.length ? spans[spans.length - 1][1] : before;
      if (place === rowEnd && row.cells.length) return { ...row, cells: row.cells.concat([fresh(row.cells[row.cells.length - 1])]) };
      return { ...row, props: { ...row.props, gridAfter: (row.props.gridAfter || 0) + 1 } };
    });
    let grid = table.grid;
    if (place <= grid.length) {
      grid = grid.slice();
      grid.splice(place, 0, width);
    }
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows, grid, widened(table.props, width))]);
    return cellStart(doc, site, site.row, side === 'left' ? site.cell : site.cell + 1);
  });
}

/**
 * Deletes the grid columns the cell a position is in covers. Cells wholly in
 * those columns go; cells that also cover other columns become narrower. A
 * row left without cells goes, and the table with its last row.
 */
export function deleteColumn(doc: DocModel, position: DocPosition, _options?: DocEditOptions): DocEdit {
  return buildEdit(doc, (draft) => {
    const site = cellSite(doc, position);
    const table = site.table;
    const [start, end] = rowSpans(table.rows[site.row])[site.cell];
    const overlap = (from: number, to: number) => Math.max(0, Math.min(to, end) - Math.max(from, start));
    const widthOf = (from: number, to: number) => {
      let sum = 0;
      for (let g = Math.max(from, start); g < Math.min(to, end); g++) sum += table.grid[g] || 0;
      return sum;
    };
    const rows: TableRow[] = [];
    let caretRow = -1;
    for (let r = 0; r < table.rows.length; r++) {
      const row = table.rows[r];
      const spans = rowSpans(row);
      const before = row.props.gridBefore || 0;
      const after = row.props.gridAfter || 0;
      const rowEnd = spans.length ? spans[spans.length - 1][1] : before;
      const cells: TableCell[] = [];
      for (let c = 0; c < row.cells.length; c++) {
        const [from, to] = spans[c];
        const cut = overlap(from, to);
        if (cut === 0) cells.push(row.cells[c]);
        else if (cut < to - from) {
          const props: CellProps = { ...cellWidened(row.cells[c].props, -widthOf(from, to)), gridSpan: to - from - cut };
          if (props.gridSpan === 1) delete props.gridSpan;
          cells.push({ ...row.cells[c], props });
        }
      }
      if (cells.length === 0) continue;
      const cutBefore = overlap(0, before);
      const cutAfter = overlap(rowEnd, rowEnd + after);
      let props = row.props;
      if (cutBefore || cutAfter) {
        props = { ...props };
        if (cutBefore) props.gridBefore = before - cutBefore;
        if (cutAfter) props.gridAfter = after - cutAfter;
      }
      if (r === site.row) caretRow = rows.length;
      rows.push(cells.length === row.cells.length && props === row.props && cells.every((cell, i) => cell === row.cells[i]) ? row : { ...row, props, cells });
    }
    if (rows.length === 0) return removeTable(draft, site);
    const grid = table.grid.slice();
    if (start < grid.length) grid.splice(start, end - start);
    draft.replace(site.story, site.container, site.index, 1, [tableWith(table, rows, grid, widened(table.props, -widthOf(start, end)))]);
    const row = caretRow >= 0 ? caretRow : Math.min(site.row, rows.length - 1);
    return cellStart(doc, site, row, Math.min(site.cell, rows[row].cells.length - 1));
  });
}

const SINGLE: Border = Object.freeze({ style: 'single', size: 4, space: 0, color: Object.freeze({ auto: true }) });

/** Width of the text area at a position, twips: of the cell it is in, else of the page between its margins. */
function availableWidth(doc: DocModel, position: DocPosition): number {
  const path = position.path;
  if (path.length >= 4) {
    const table = containerBlocks(doc, position.story, path.slice(0, -4))[path[path.length - 4]];
    if (table && table.type === 'table') {
      const row = table.rows[path[path.length - 3]];
      const cell = row.cells[path[path.length - 2]];
      const [from, to] = rowSpans(row)[path[path.length - 2]];
      let width = cell.props.width && cell.props.width.type === 'dxa' ? cell.props.width.value : 0;
      if (!(width > 0)) for (let g = from; g < to; g++) width += table.grid[g] || 0;
      // Less the default cell margins of Word on both sides.
      if (width > 0) return Math.max(360, width - 216);
    }
  }
  const section = doc.body.section;
  const page = (section && section.page && section.page.width) || 12240;
  const margins = (section && section.margins) || {};
  const left = margins.left === undefined ? 1440 : margins.left;
  const right = margins.right === undefined ? 1440 : margins.right;
  return Math.max(360, page - left - right);
}

/**
 * Inserts a table of `rows` by `columns` empty cells of equal width at a
 * position (see insertBlock for where exactly). The table uses the table
 * style "TableGrid" when the document has it and single borders of its own
 * otherwise. The caret goes into the first cell.
 */
export function insertTable(doc: DocModel, position: DocPosition, rows: number, columns: number, _options?: DocEditOptions): DocEdit {
  if (!whole(rows, 1, MAX_NEW_CELLS) || !whole(columns, 1, MAX_COLUMNS) || rows * columns > MAX_NEW_CELLS) bad('The size of the table is not valid');
  return buildEdit(doc, (draft) => {
    resolvePosition(doc, position);
    const width = Math.max(1, Math.floor(availableWidth(doc, position) / columns));
    const style = doc.styles.styles.get('TableGrid');
    const props: TableProps = { width: { type: 'auto', value: 0 } };
    if (style && style.type === 'table') {
      props.style = 'TableGrid';
      props.look = { firstRow: true, lastRow: false, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, val: 0x04a0 };
    } else props.borders = { top: SINGLE, left: SINGLE, bottom: SINGLE, right: SINGLE, insideH: SINGLE, insideV: SINGLE };
    const grid: number[] = [];
    for (let c = 0; c < columns; c++) grid.push(width);
    const cellProps: CellProps = { width: { type: 'dxa', value: width } };
    const list: TableRow[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < columns; c++) cells.push({ props: cellProps, blocks: [{ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] }] });
      list.push({ props: {}, cells });
    }
    const index = placeBlock(draft, position, { type: 'table', props, grid, rows: list });
    return { story: position.story, path: position.path.slice(0, -1).concat([index, 0, 0, 0]), offset: 0 };
  });
}

// ------------------------------------------------------------------ hyperlinks

const REL_BASE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_BASE_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';

/** An id no relationship of the part has yet. */
function freeRelId(relationships: Map<string, unknown>): string {
  let next = 1;
  relationships.forEach((_value, id) => {
    const match = /^rId(\d{1,9})$/.exec(id);
    if (match && Number(match[1]) >= next) next = Number(match[1]) + 1;
  });
  return 'rId' + next;
}

const isRevision = (w: Wrapper) => w.kind === 'ins' || w.kind === 'del' || w.kind === 'moveFrom' || w.kind === 'moveTo';

/**
 * Makes the text of a range a hyperlink to `target`: an address ("https://…",
 * "mailto:…"), or "#name" for a bookmark of the document. The range must lie
 * in one paragraph; a collapsed range gets the target itself as its text. A
 * link the text already has is replaced. For an address the part of the story
 * gets a relationship, which undoing removes again.
 */
export function insertHyperlink(doc: DocModel, range: DocRange, target: string, _options?: DocEditOptions): DocEdit {
  if (typeof target !== 'string' || target === '' || target === '#' || target.length > 2083 || /[\u0000-\u001f]/.test(target)) bad('The target of the hyperlink is not valid');
  return buildEdit(doc, (draft) => {
    const ordered = comparePositions(range.anchor, range.focus) <= 0;
    const start = ordered ? range.anchor : range.focus;
    const end = ordered ? range.focus : range.anchor;
    const at = resolvePosition(doc, start);
    resolvePosition(doc, end);
    if (inTextBox(start.path)) bad('The content of a text box cannot be edited');
    if (start.path.length !== end.path.length || start.path.some((value, i) => value !== end.path[i])) bad('A hyperlink must lie in one paragraph');
    const story = start.story;
    let link: Wrapper;
    if (target.charAt(0) === '#') link = { kind: 'hyperlink', anchor: target.slice(1) };
    else {
      const host = storyOf(doc, story).host;
      const id = freeRelId(host.relationships);
      draft.apply({ type: 'setRelationship', story, id, relationship: { id, type: (doc.strict ? REL_BASE_STRICT : REL_BASE) + '/hyperlink', target, external: true, part: null } });
      link = { kind: 'hyperlink', rId: id, target };
    }
    const paragraph = at.paragraph;
    const pieces = flatten(paragraph);
    const from = boundary(pieces, start.offset, start.offset === end.offset ? 'caret' : 'start');
    let to: number;
    if (start.offset === end.offset) {
      const props = inheritedProps(pieces, from, paragraph.props.markProps);
      const wrap = wrapBetween(from > 0 ? pieces[from - 1].wrap : undefined, from < pieces.length ? pieces[from].wrap : undefined);
      const added = textPieces(target.replace(/[\t\n\u000b]/g, ' '), props, wrap, {});
      pieces.splice(from, 0, ...added);
      to = from + added.length;
    } else to = boundary(pieces, end.offset, 'caret');
    const styled = doc.styles.styles.get('Hyperlink');
    const useStyle = styled !== undefined && styled.type === 'character';
    // The link goes inside the wrappers the text is in, but around tracked changes (which may only hold runs), and
    // replaces a link that is there. Wrappers whose parent changes are copied, once each, so that pieces that
    // shared a wrapper still do.
    const stripped = new Map<Wrapper, Wrapper | undefined>();
    const strip = (w: Wrapper | undefined): Wrapper | undefined => {
      if (!w) return undefined;
      if (stripped.has(w)) return stripped.get(w);
      const parent = strip(w.parent);
      let out: Wrapper | undefined;
      if (isRevision(w) || w.kind === 'hyperlink') out = parent;
      else if (parent === w.parent) out = w;
      else {
        out = { ...w };
        if (parent) out.parent = parent;
        else delete out.parent;
      }
      stripped.set(w, out);
      return out;
    };
    const links = new Map<Wrapper | undefined, Wrapper>();
    const linkIn = (outer: Wrapper | undefined): Wrapper => {
      let mine = links.get(outer);
      if (!mine) {
        mine = { ...link };
        if (outer) mine.parent = outer;
        links.set(outer, mine);
      }
      return mine;
    };
    const revisions = new Map<Wrapper, Wrapper>();
    const linked = (w: Wrapper | undefined, base: Wrapper): Wrapper => {
      if (!w) return base;
      if (!isRevision(w)) return linked(w.parent, base);
      let copy = revisions.get(w);
      if (!copy) revisions.set(w, (copy = { ...w, parent: linked(w.parent, base) }));
      return copy;
    };
    for (let q = from; q < to; q++) {
      const piece = pieces[q];
      if (piece.frozen) continue;
      const next: Piece = { ...piece, wrap: linked(piece.wrap, linkIn(strip(piece.wrap))) };
      if (piece.item && !piece.code && piece.item.type !== 'fieldChar') {
        next.props = useStyle ? patchRunProps(piece.props, { style: 'Hyperlink' }) : patchRunProps(piece.props, { color: '0563C1', underline: true });
      }
      pieces[q] = next;
    }
    const first = offsetOf(pieces, from);
    const last = offsetOf(pieces, to);
    draft.setParagraph(story, start.path, paragraphWith(paragraph, rebuild(pieces)));
    return { anchor: { story, path: start.path.slice(), offset: first }, focus: { story, path: start.path.slice(), offset: last } };
  });
}

// -------------------------------------------------------------------- pictures

export interface NewImage {
  bytes: Uint8Array;
  /** Media type: "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff", "image/x-emf", "image/x-wmf". */
  type: string;
  /** Size as shown, in EMU (914400 to the inch, 9525 to the CSS pixel). */
  width: number;
  height: number;
  /** Alternative text. */
  description?: string;
}

const IMAGE_EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/tiff': 'tiff',
  'image/x-emf': 'emf',
  'image/x-wmf': 'wmf',
};

const DRAWING_NS = {
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  r: REL_BASE,
};
const DRAWING_NS_STRICT = {
  wp: 'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
  a: 'http://purl.oclc.org/ooxml/drawingml/main',
  pic: 'http://purl.oclc.org/ooxml/drawingml/picture',
  r: REL_BASE_STRICT,
};

const el = (name: string, attrs: [string, string][], children: XmlNode[] = []): XmlElement => ({ type: 'element', name, attrs: new Map(attrs), children });

/** Highest number N among the parts named /word/media/imageN.* the document knows of, and the highest drawing id. */
function mediaNumbers(doc: DocModel): { image: number; drawing: number } {
  let image = 0;
  let drawing = 0;
  const part = (name: string | null | undefined) => {
    const match = name ? /\/media\/image(\d{1,9})\.[^/]*$/.exec(name) : null;
    if (match && Number(match[1]) > image) image = Number(match[1]);
  };
  if (doc.media) doc.media.forEach((_media, name) => part(name));
  const hosts: { relationships: Map<string, { part: string | null }> }[] = [doc.body];
  doc.headers.forEach((story) => hosts.push(story));
  doc.footers.forEach((story) => hosts.push(story));
  for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) if (collection) hosts.push(collection);
  for (const host of hosts) host.relationships.forEach((relationship) => part(relationship.part));

  const items = (list: RunItem[], depth: number) => {
    for (const item of list) {
      if (item.type === 'alternate' && depth < 32) items(item.items, depth + 1);
      if (item.type !== 'drawing') continue;
      for (const frame of item.xml.children) {
        if (frame.type !== 'element') continue;
        for (const child of frame.children) {
          if (child.type !== 'element' || child.name.slice(child.name.indexOf(':') + 1) !== 'docPr') continue;
          const id = Number(child.attrs.get('id'));
          if (id > drawing && id < 0x7fffffff) drawing = id;
        }
      }
      if (item.textBoxes && depth < 32) for (const box of item.textBoxes) blocks(box, depth + 1);
    }
  };
  const blocks = (list: Block[], depth: number) => {
    if (depth > 64) return;
    for (const block of list) {
      if (block.type === 'paragraph') {
        for (const inline of block.children) if (inline.type === 'run') items(inline.content, depth);
      } else if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) blocks(cell.blocks, depth + 1);
      }
    }
  };
  blocks(doc.body.blocks, 0);
  doc.headers.forEach((story) => blocks(story.blocks, 0));
  doc.footers.forEach((story) => blocks(story.blocks, 0));
  for (const collection of [doc.footnotes, doc.endnotes]) if (collection) collection.items.forEach((note) => blocks(note.blocks, 0));
  return { image, drawing };
}

/**
 * Inserts a picture in line with the text at a position.
 *
 * The bytes go into DocModel.media under the name of a new package part,
 * "/word/media/imageN.ext", and the part of the story gets a relationship to
 * it; the drawing refers to both (ImageRef.rId, ImageRef.target). A writer
 * has to add every entry of DocModel.media that the package does not hold as
 * a part, with a content type for its extension. Undoing removes the entry
 * and the relationship again.
 */
export function insertImage(doc: DocModel, position: DocPosition, image: NewImage, _options?: DocEditOptions): DocEdit {
  const extension = image && typeof image.type === 'string' ? IMAGE_EXTENSIONS[image.type] : undefined;
  if (!image || !(image.bytes instanceof Uint8Array) || image.bytes.length === 0) bad('A picture needs its bytes');
  if (extension === undefined) bad('The type of the picture is not supported');
  if (!whole(image.width, 1, 0x7fffffff) || !whole(image.height, 1, 0x7fffffff)) bad('The size of the picture is not valid');
  return buildEdit(doc, (draft) => {
    const at = resolvePosition(doc, position);
    if (inTextBox(position.path)) bad('The content of a text box cannot be edited');
    const story = position.story;
    const host = storyOf(doc, story).host;
    const numbers = mediaNumbers(doc);
    const name = `/word/media/image${numbers.image + 1}.${extension}`;
    const media: DocMedia = { bytes: image.bytes, type: image.type };
    draft.apply({ type: 'setMedia', name, media });
    const rId = freeRelId(host.relationships);
    // Relative to the part of the story when that is beside the media folder, as Word writes it.
    const beside = /^\/word\/[^/]+$/.test(host.part);
    draft.apply({
      type: 'setRelationship',
      story,
      id: rId,
      relationship: { id: rId, type: (doc.strict ? REL_BASE_STRICT : REL_BASE) + '/image', target: beside ? name.slice('/word/'.length) : name, external: false, part: name },
    });

    const ns = doc.strict ? DRAWING_NS_STRICT : DRAWING_NS;
    let w = 'w';
    for (const [attr, value] of host.rootAttrs) if (attr.indexOf('xmlns:') === 0 && /wordprocessingml\/(2006\/)?main$/.test(value)) w = attr.slice(6);
    const id = String(numbers.drawing + 1);
    const title = `Picture ${id}`;
    const size: [string, string][] = [
      ['cx', String(image.width)],
      ['cy', String(image.height)],
    ];
    const docPr: [string, string][] = [
      ['id', id],
      ['name', title],
    ];
    if (image.description !== undefined) docPr.push(['descr', String(image.description)]);
    const xml = el(`${w}:drawing`, [], [
      el('wp:inline', [['xmlns:wp', ns.wp], ['distT', '0'], ['distB', '0'], ['distL', '0'], ['distR', '0']], [
        el('wp:extent', size),
        el('wp:docPr', docPr),
        el('wp:cNvGraphicFramePr', [], [el('a:graphicFrameLocks', [['xmlns:a', ns.a], ['noChangeAspect', '1']])]),
        el('a:graphic', [['xmlns:a', ns.a]], [
          el('a:graphicData', [['uri', ns.pic]], [
            el('pic:pic', [['xmlns:pic', ns.pic]], [
              el('pic:nvPicPr', [], [el('pic:cNvPr', [['id', id], ['name', title]]), el('pic:cNvPicPr', [])]),
              el('pic:blipFill', [], [el('a:blip', [['xmlns:r', ns.r], ['r:embed', rId]]), el('a:stretch', [], [el('a:fillRect', [])])]),
              el('pic:spPr', [], [
                el('a:xfrm', [], [el('a:off', [['x', '0'], ['y', '0']]), el('a:ext', size)]),
                el('a:prstGeom', [['prst', 'rect']], [el('a:avLst', [])]),
              ]),
            ]),
          ]),
        ]),
      ]),
    ]);
    const ref = { rId, target: name };
    const item: DrawingItem = { type: 'drawing', placement: 'inline', kind: 'picture', width: image.width, height: image.height, name: title, images: [{ image: ref }], image: ref, xml };
    if (image.description !== undefined) item.description = String(image.description);

    const paragraph = at.paragraph;
    const pieces = flatten(paragraph);
    const index = boundary(pieces, position.offset, 'caret');
    const props = inheritedProps(pieces, index, paragraph.props.markProps);
    const wrap = wrapBetween(index > 0 ? pieces[index - 1].wrap : undefined, index < pieces.length ? pieces[index].wrap : undefined);
    pieces.splice(index, 0, itemPiece(item, props, wrap, groupFor(pieces, index, props, wrap)));
    draft.setParagraph(story, position.path, paragraphWith(paragraph, rebuild(pieces)));
    return { story, path: position.path.slice(), offset: offsetOf(pieces, index) };
  });
}
