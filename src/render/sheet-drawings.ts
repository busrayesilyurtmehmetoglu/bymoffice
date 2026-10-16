// What floats above the cells of a sheet: pictures, simple shapes with their
// text, charts (drawn by render/chart.ts), placeholders for what is not
// drawn, and the boxes of comments that are always shown. Also says which
// cells carry a comment.
//
// The model holds drawings and comments as the file has them. Rows and columns
// inserted or deleted since (Sheet.structuralEdits) are applied here, by the
// rules the writer moves the anchors in the file with, so that what is on
// screen is where a saved and reopened file shows it.
//
// Only what is in sight is touched: the objects are kept sorted by their top
// edge, and a frame looks at those that can reach into the pane it draws.
// Nothing that comes from a file reaches `ctx.font`, `fillStyle` or
// `strokeStyle` unchecked.

import { adjustRange } from '../calc/refs';
import type { Rect } from '../layout/grid-geometry';
import type { SheetMetrics } from '../layout/sheet-metrics';
import {
  MAX_COLS,
  MAX_ROWS,
  type DrawingAnchor,
  type DrawingMarker,
  type DrawingObject,
  type DrawingShape,
  type DrawingText,
  type DrawingTextRun,
  type Font,
  type Sheet,
  type SheetComment,
  type SheetDrawing,
  type SheetStructuralEdit,
} from '../model/sheet/types';
import { ChartRenderer, type ChartRendererOptions } from './chart';
import type { DecodedImage } from './images';
import { fontsVersion } from './fonts';
import { cssFont, FONT_FALLBACK, TextPainter, wrapLines } from './text';

const EMU_PER_PIXEL = 9525;
/** Groups inside groups that are drawn; the reader stops at the same depth. */
const MAX_DEPTH = 16;
/** Objects drawn in one pane of one frame, group members included. */
const MAX_PAINTED = 20000;
const MAX_TEXT_LINES = 1000;
const MAX_TEXT_PIECES = 8000;
const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const DEFAULT_TEXT_POINTS = 11;
const LINE_HEIGHT = 1.2;
/** Size of the box of a comment the file gives no place for, in pixels. */
const NOTE_WIDTH = 144;
const NOTE_HEIGHT = 79;
const NOTE_FILL = '#FFFFE1';
const PLACEHOLDER_FILL = '#F1F3F4';
const PLACEHOLDER_BORDER = '#BDC1C6';
const PLACEHOLDER_TEXT = '#5F6368';
const LOADING_FILL = '#F6F7F8';
const LABEL_FONT = '12px ' + FONT_FALLBACK;

// ------------------------------------------------------------ structural edits

/** Where a row or column index ends up; `snapped` when it was deleted and now stands for the place of the deletion. */
function shiftIndex(value: number, edit: SheetStructuralEdit): { value: number; snapped: boolean } {
  const limit = edit.axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (edit.kind === 'insert') return { value: value >= edit.at ? Math.min(value + edit.count, limit - 1) : value, snapped: false };
  const end = edit.at + edit.count - 1;
  if (value > end) return { value: value - edit.count, snapped: false };
  if (value >= edit.at) return { value: edit.at, snapped: true };
  return { value, snapped: false };
}

/** Puts a corner at the start of its row or column. True when that moved it. */
function snap(marker: DrawingMarker, row: boolean): boolean {
  let changed = false;
  if (row) {
    if (marker.rowOff !== 0 || (marker.rowShare !== undefined && marker.rowShare !== 0)) changed = true;
    marker.rowOff = 0;
    if (marker.rowShare !== undefined) marker.rowShare = 0;
  } else {
    if (marker.colOff !== 0 || (marker.colShare !== undefined && marker.colShare !== 0)) changed = true;
    marker.colOff = 0;
    if (marker.colShare !== undefined) marker.colShare = 0;
  }
  return changed;
}

/**
 * An anchor after rows and columns were inserted or deleted, exactly as the
 * writer moves it in the file: an object that moves and sizes with cells has
 * both corners adjusted; one that only moves keeps its span of cells; an
 * absolute one stays. A corner whose row or column was deleted goes to where
 * the deletion was. The same object comes back when nothing moved.
 */
export function adjustDrawingAnchor(anchor: DrawingAnchor, edits: readonly SheetStructuralEdit[] | undefined): DrawingAnchor {
  if (!edits || edits.length === 0 || anchor.kind === 'absolute' || !anchor.from) return anchor;
  if (anchor.kind === 'twoCell' && anchor.editAs === 'absolute') return anchor;
  const from: DrawingMarker = { ...anchor.from };
  const to: DrawingMarker | undefined = anchor.kind === 'twoCell' && anchor.to ? { ...anchor.to } : undefined;
  let changed = false;
  for (const edit of edits) {
    const row = edit.axis === 'row';
    const before = row ? from.row : from.col;
    const moved = shiftIndex(before, edit);
    if (moved.value !== before) changed = true;
    if (row) from.row = moved.value;
    else from.col = moved.value;
    if (moved.snapped && snap(from, row)) changed = true;
    if (!to) continue;
    const toBefore = row ? to.row : to.col;
    if (anchor.editAs === 'oneCell') {
      // Moves with its top-left cell and keeps its size.
      const limit = row ? MAX_ROWS : MAX_COLS;
      const next = Math.max(0, Math.min(toBefore + moved.value - before, limit - 1));
      if (next !== toBefore) changed = true;
      if (row) to.row = next;
      else to.col = next;
    } else {
      const end = shiftIndex(toBefore, edit);
      if (end.value !== toBefore) changed = true;
      if (row) to.row = end.value;
      else to.col = end.value;
      if (end.snapped && snap(to, row)) changed = true;
    }
  }
  if (!changed) return anchor;
  return to ? { ...anchor, from, to } : { ...anchor, from };
}

/**
 * The structural edits that move a drawing: all of them for a drawing of the file, and for one
 * that was added, moved or changed since only those made after it was (SheetDrawing.since).
 */
function drawingEdits(drawing: Pick<SheetDrawing, 'since'>, edits: readonly SheetStructuralEdit[] | undefined): readonly SheetStructuralEdit[] | undefined {
  const since = drawing.since;
  if (!edits || since === undefined || !(since > 0)) return edits;
  // A count that is not one of this list (it can only come from outside) moves the drawing by nothing.
  return since >= edits.length || Math.floor(since) !== since ? [] : edits.slice(since);
}

/** Where a drawing of the model is now: the one rule by which the view, the writer and the commands place it. */
export function drawingAnchorNow(drawing: Pick<SheetDrawing, 'anchor' | 'since'>, edits: readonly SheetStructuralEdit[] | undefined): DrawingAnchor {
  return adjustDrawingAnchor(drawing.anchor, drawingEdits(drawing, edits));
}

/** Where the cell of a comment ends up; undefined when it was deleted (the writer removes the comment then). */
export function adjustCommentCell(row: number, col: number, edits: readonly SheetStructuralEdit[] | undefined): { row: number; col: number } | undefined {
  if (!edits) return { row, col };
  for (const edit of edits) {
    const moved = adjustRange({ top: row, bottom: row, left: col, right: col }, edit);
    if (!moved) return undefined;
    row = moved.top;
    col = moved.left;
  }
  return { row, col };
}

// --------------------------------------------------------------------- places

/** The rectangle of an anchor on the sheet, in pixels from the top-left corner of cell A1. */
export function anchorRect(anchor: DrawingAnchor, metrics: SheetMetrics): Rect {
  const px = (emu: number | undefined): number => (typeof emu === 'number' && emu > 0 && emu < Infinity ? emu / EMU_PER_PIXEL : 0);
  // An offset never reaches beyond its cell: a hidden row or column takes what was in it along.
  const share = (value: number): number => (value > 0 ? (value < 1 ? value : 1) : 0);
  const markerX = (marker: DrawingMarker): number => {
    const size = metrics.cols.size(marker.col);
    return metrics.cols.offset(marker.col) + (marker.colShare !== undefined ? share(marker.colShare) * size : Math.min(px(marker.colOff), size));
  };
  const markerY = (marker: DrawingMarker): number => {
    const size = metrics.rows.size(marker.row);
    return metrics.rows.offset(marker.row) + (marker.rowShare !== undefined ? share(marker.rowShare) * size : Math.min(px(marker.rowOff), size));
  };
  if (anchor.kind === 'absolute' || !anchor.from) return { x: px(anchor.x), y: px(anchor.y), width: px(anchor.width), height: px(anchor.height) };
  const x = markerX(anchor.from);
  const y = markerY(anchor.from);
  if (anchor.kind === 'oneCell' || !anchor.to) return { x, y, width: px(anchor.width), height: px(anchor.height) };
  return { x, y, width: Math.max(markerX(anchor.to) - x, 0), height: Math.max(markerY(anchor.to) - y, 0) };
}

/** Rotations near a quarter turn: the file gives the box the object takes up turned, not the box that is turned. */
export function quarterTurned(rotation: number | undefined): boolean {
  if (rotation === undefined) return false;
  const angle = ((rotation % 360) + 360) % 360;
  return (angle >= 45 && angle < 135) || (angle >= 225 && angle < 315);
}

interface Placed {
  /** Position in the order of the file: what comes later is drawn on top. */
  order: number;
  object: DrawingObject;
  /** The box the object is drawn in (before its own rotation), sheet pixels. */
  x: number;
  y: number;
  width: number;
  height: number;
  /** What it may cover on the sheet, rotation and outline included. */
  left: number;
  top: number;
  right: number;
  bottom: number;
}

function place(order: number, object: DrawingObject, rect: Rect): Placed | undefined {
  let { x, y, width, height } = rect;
  if (!(width >= 0 && height >= 0 && width < Infinity && height < Infinity && Number.isFinite(x) && Number.isFinite(y))) return undefined;
  const line = object.type === 'shape' && object.connector === true;
  // Squeezed to nothing by hidden rows or columns; a line may well be flat.
  if (line ? width === 0 && height === 0 : width === 0 || height === 0) return undefined;
  if (quarterTurned(object.rotation)) {
    x += (width - height) / 2;
    y += (height - width) / 2;
    const swap = width;
    width = height;
    height = swap;
  }
  let reach = object.type === 'shape' && object.line ? object.line.width / 2 + 1 : 1;
  let left = x;
  let top = y;
  let right = x + width;
  let bottom = y + height;
  if (object.rotation) {
    // The circle around the box holds it at any angle.
    const radius = Math.sqrt(width * width + height * height) / 2;
    const cx = x + width / 2;
    const cy = y + height / 2;
    left = cx - radius;
    right = cx + radius;
    top = cy - radius;
    bottom = cy + radius;
  }
  if (object.type === 'group') reach += 16;
  return { order, object, x, y, width, height, left: left - reach, top: top - reach, right: right + reach, bottom: bottom + reach };
}

/**
 * The structural edits that move a comment: all of them for a comment of the file, and for a
 * note added since only those made after it was (SheetComment.since).
 */
export function commentEdits(comment: Pick<SheetComment, 'since'>, edits: readonly SheetStructuralEdit[] | undefined): readonly SheetStructuralEdit[] | undefined {
  const since = comment.since;
  if (!edits || since === undefined || !(since > 0)) return edits;
  // A count that is not one of this list (it can only come from outside) moves the note by nothing.
  return since >= edits.length || Math.floor(since) !== since ? [] : edits.slice(since);
}

/**
 * The cell a comment is on now; undefined when its cell was deleted. The one rule by which
 * the view, the writer and the operations place a comment of the model.
 */
export function commentCellNow(comment: Pick<SheetComment, 'row' | 'col' | 'since'>, edits: readonly SheetStructuralEdit[] | undefined): { row: number; col: number } | undefined {
  return adjustCommentCell(comment.row, comment.col, commentEdits(comment, edits));
}

// -------------------------------------------------------------------- comments

/** A comment where it is now. */
export interface PlacedComment {
  row: number;
  col: number;
  comment: SheetComment;
}

function noteRuns(comment: SheetComment): DrawingTextRun[] {
  const runs: DrawingTextRun[] = [];
  const base = { size: 9, color: '#000000' };
  if (comment.threaded) {
    const message = (author: string | undefined, text: string, first: boolean) => {
      if (author !== undefined && author !== '') runs.push({ ...base, text: `${first ? '' : '\n'}${author}:\n`, bold: true });
      else if (!first) runs.push({ ...base, text: '\n' });
      runs.push({ ...base, text });
    };
    message(comment.author, comment.text, true);
    for (const reply of comment.replies || []) message(reply.author, reply.text, false);
    return runs;
  }
  if (comment.runs) {
    for (const run of comment.runs) {
      const font = run.font;
      runs.push({ ...base, text: run.text, bold: font ? font.bold : false, italic: font ? font.italic : false, size: font && font.size ? font.size : base.size, font: font ? font.name : undefined });
    }
    return runs;
  }
  runs.push({ ...base, text: comment.text });
  return runs;
}

/** The box of a comment as a shape: what Excel draws for a note that is always shown. */
function noteShape(comment: SheetComment): DrawingShape {
  return {
    type: 'shape',
    geometry: 'rect',
    fill: NOTE_FILL,
    line: { color: '#000000', width: 1 },
    text: { paragraphs: [{ runs: noteRuns(comment) }], insets: [4, 3, 4, 3] },
  };
}

// ---------------------------------------------------------------------- source

export interface SheetDrawingLabels {
  chart: string;
  diagram: string;
  object: string;
  picture: string;
}

const DEFAULT_LABELS: SheetDrawingLabels = { chart: 'Chart', diagram: 'Diagram', object: 'Object', picture: 'Picture' };

export interface SheetDrawingsOptions {
  /**
   * Decoded pictures by their key in Workbook.media; `width` and `height` are the canvas
   * pixels the whole picture will cover. Without it every picture is a placeholder.
   */
  images?: { get(part: string, width?: number, height?: number): DecodedImage | undefined; status?(part: string): 'pending' | 'ready' | 'failed' };
  /** Words shown inside placeholders. */
  labels?: Partial<SheetDrawingLabels>;
  /** False leaves the boxes of comments that are always shown out of the layer: on paper comments are not printed. Default true. */
  notes?: boolean;
  /**
   * Where charts get their numbers from: the workbook whose cells the series read, the locale
   * and a counter of changes. Without it charts are drawn from the values their files cached.
   * A renderer may be handed in instead: the layers of several sheets then share what it keeps,
   * and the limit of what it keeps.
   */
  charts?: ChartRendererOptions | ChartRenderer;
}

/** Where the layer is drawn: a clip rectangle on the canvas and the canvas position of the top-left corner of cell A1. */
export interface DrawingTarget {
  metrics: SheetMetrics;
  clip: Rect;
  originX: number;
  originY: number;
  /** Canvas backing store pixels per CSS pixel; pictures are decoded for it. */
  devicePixelRatio: number;
}

interface TextPiece {
  text: string;
  font: string;
  color: string;
  size: number;
  width: number;
  underline: boolean;
  strike: boolean;
  space: boolean;
}

interface TextLine {
  pieces: TextPiece[];
  width: number;
  height: number;
  align: 'left' | 'center' | 'right';
}

interface TextLayout {
  width: number;
  lines: TextLine[];
  height: number;
}

/** The side of the squares a selected drawing is sized by, in pixels. */
export const HANDLE_SIZE = 8;
/** How far above the middle of its top side the handle that turns a picture stands, in pixels. */
export const ROTATE_HANDLE_GAP = 20;
/**
 * The same on a touch screen, where a finger takes a handle within 16 pixels: far enough up that
 * the handle that turns and the handle of the top side never take the same press.
 */
export const ROTATE_HANDLE_GAP_COARSE = 40;
/** The handles of a selected drawing: its corners and the middles of its sides. */
export type DrawingHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w';
const HANDLES: DrawingHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const SELECTION_COLOR = '#1A73E8';

/** Where the middle of a handle of a box is. */
function handlePoint(rect: Rect, handle: DrawingHandle): { x: number; y: number } {
  const x = handle === 'nw' || handle === 'w' || handle === 'sw' ? rect.x : handle === 'n' || handle === 's' ? rect.x + rect.width / 2 : rect.x + rect.width;
  const y = handle === 'nw' || handle === 'n' || handle === 'ne' ? rect.y : handle === 'w' || handle === 'e' ? rect.y + rect.height / 2 : rect.y + rect.height;
  return { x, y };
}

/** The handle of a box that a point is on, within `reach` pixels of its middle; undefined when it is on none. */
export function drawingHandleAt(rect: Rect, x: number, y: number, reach: number = HANDLE_SIZE): DrawingHandle | undefined {
  for (const handle of HANDLES) {
    const at = handlePoint(rect, handle);
    if (Math.abs(x - at.x) <= reach && Math.abs(y - at.y) <= reach) return handle;
  }
  return undefined;
}

/**
 * A box after one of its handles was dragged by (dx, dy), or the whole box (`handle` "move").
 * It keeps at least `min` pixels each way, and the side that is not dragged stays where it is.
 */
export function resizedRect(rect: Rect, handle: DrawingHandle | 'move', dx: number, dy: number, min: number): Rect {
  if (handle === 'move') return { x: Math.max(rect.x + dx, 0), y: Math.max(rect.y + dy, 0), width: rect.width, height: rect.height };
  let left = rect.x;
  let top = rect.y;
  let right = rect.x + rect.width;
  let bottom = rect.y + rect.height;
  if (handle === 'nw' || handle === 'w' || handle === 'sw') left = Math.max(Math.min(left + dx, right - min), 0);
  if (handle === 'ne' || handle === 'e' || handle === 'se') right = Math.max(right + dx, left + min);
  if (handle === 'nw' || handle === 'n' || handle === 'ne') top = Math.max(Math.min(top + dy, bottom - min), 0);
  if (handle === 'sw' || handle === 's' || handle === 'se') bottom = Math.max(bottom + dy, top + min);
  return { x: left, y: top, width: right - left, height: bottom - top };
}

/**
 * A box sized by a corner handle (see resizedRect()) made to keep the proportions of the box it
 * was sized from: the larger of the two changes wins, and the corner across from the handle stays
 * where it is. A side handle and a move are left as they are.
 */
export function ratioRect(start: Rect, rect: Rect, handle: DrawingHandle | 'move', min: number): Rect {
  if (handle === 'move' || handle === 'n' || handle === 's' || handle === 'e' || handle === 'w' || !(start.width > 0) || !(start.height > 0)) return rect;
  const scale = Math.max(rect.width / start.width, rect.height / start.height, min / start.width, min / start.height);
  const width = start.width * scale;
  const height = start.height * scale;
  const left = handle === 'nw' || handle === 'sw' ? start.x + start.width - width : start.x;
  const top = handle === 'nw' || handle === 'ne' ? start.y + start.height - height : start.y;
  // Not past the corner of the sheet: the box is made smaller instead, in proportion.
  if (left < 0 || top < 0) {
    const room = Math.min(handle === 'nw' || handle === 'sw' ? (start.x + start.width) / width : 1, handle === 'nw' || handle === 'ne' ? (start.y + start.height) / height : 1);
    const w = width * room;
    const h = height * room;
    return { x: handle === 'nw' || handle === 'sw' ? start.x + start.width - w : start.x, y: handle === 'nw' || handle === 'ne' ? start.y + start.height - h : start.y, width: w, height: h };
  }
  return { x: left, y: top, width, height };
}

/**
 * The drawings and comments of one sheet as they are shown. Made once per
 * sheet and asked every frame; what it works out is kept until the sheet's
 * drawings, comments, structural edits or row and column sizes change.
 */
export class SheetDrawings {
  private readonly sheet: Sheet;
  private readonly images: SheetDrawingsOptions['images'];
  private labels: SheetDrawingLabels;
  private readonly notes: boolean;
  private readonly charts: ChartRenderer;

  private drawingsSeen: SheetDrawing[] | undefined;
  private commentsSeen: SheetComment[] | undefined;
  private editsSeen: SheetStructuralEdit[] | undefined;
  private metricsSeen: SheetMetrics | undefined;
  private first = true;

  /** Anchors with the structural edits applied, index for index with Sheet.drawings. */
  private anchors: DrawingAnchor[] = [];
  private placedComments: PlacedComment[] = [];
  private commentByCell = new Map<number, PlacedComment>();
  /** Sorted by `top`; `reach[i]` is the largest `bottom` of items 0..i. */
  private placed: Placed[] = [];
  private tops: number[] = [];
  private reach: number[] = [];
  private layouts = new WeakMap<DrawingText, TextLayout>();
  /** fontsVersion() the layouts were made under: they hold font strings and widths of that time. */
  private layoutFonts = fontsVersion();
  /** The drawing that is shown as selected, by its place in Sheet.drawings: a frame with handles is drawn around it. */
  selected: number | undefined;
  /** Where the selected drawing would stand if the drag that is under way ended now (sheet pixels): drawn as a dashed frame. */
  ghost: Rect | undefined;
  /** How the ghost is turned while the handle that turns a picture is dragged; otherwise it is turned as the picture is. */
  ghostTurn: number | undefined;
  /** The selected drawing is a picture that can be turned: a handle above its frame turns it. */
  turnable = false;
  /** How far above the frame that handle stands (see ROTATE_HANDLE_GAP and ROTATE_HANDLE_GAP_COARSE of render/picture-frame.ts). */
  turnGap = ROTATE_HANDLE_GAP;
  /**
   * The selected picture is being cropped: the box the whole picture would take up, in the frame of
   * its box (see uncroppedBox()), drawn dashed; its handles are those that crop it (dark).
   */
  cropping: Rect | undefined;

  constructor(sheet: Sheet, options: SheetDrawingsOptions = {}) {
    this.sheet = sheet;
    this.images = options.images && typeof options.images.get === 'function' ? options.images : undefined;
    this.labels = { ...DEFAULT_LABELS, ...(options.labels || {}) };
    this.notes = options.notes !== false;
    this.charts = options.charts instanceof ChartRenderer ? options.charts : new ChartRenderer(options.charts || {});
  }

  /** True when the sheet has nothing for this layer; a view then leaves it out of its frames. */
  static isEmpty(sheet: Sheet): boolean {
    return !(sheet.drawings && sheet.drawings.length > 0) && !(sheet.comments && sheet.comments.length > 0);
  }

  setLabels(labels: Partial<SheetDrawingLabels>): void {
    this.labels = { ...DEFAULT_LABELS, ...labels };
  }

  /** Number of objects placed on the sheet (comment boxes that are always shown included). */
  get count(): number {
    return this.placed.length;
  }

  /** The anchors as they are now, index for index with Sheet.drawings. */
  currentAnchors(): DrawingAnchor[] {
    this.syncModel();
    return this.anchors;
  }

  /** The comments as they are now: at the cells their cells moved to, without those whose cells were deleted. */
  currentComments(): PlacedComment[] {
    this.syncModel();
    return this.placedComments;
  }

  /** Whether a cell carries a comment. A plain lookup: the grid asks for every visible cell of every frame. */
  has(row: number, col: number): boolean {
    this.syncModel();
    return this.commentByCell.has(row * MAX_COLS + col);
  }

  commentAt(row: number, col: number): SheetComment | undefined {
    this.syncModel();
    const found = this.commentByCell.get(row * MAX_COLS + col);
    return found ? found.comment : undefined;
  }

  /** The box of the drawing at a place of Sheet.drawings, in sheet pixels; undefined for one that is not shown. */
  boxOf(metrics: SheetMetrics, index: number): Rect | undefined {
    this.syncPlaces(metrics);
    for (const item of this.placed) if (item.order === index) return { x: item.x, y: item.y, width: item.width, height: item.height };
    return undefined;
  }

  /**
   * The drawing under a point of the sheet (pixels from the corner of cell A1) that `wanted`
   * accepts, the one on top first; -1 when there is none. Comment boxes are not drawings.
   */
  hit(metrics: SheetMetrics, x: number, y: number, wanted: (object: DrawingObject) => boolean = () => true): number {
    this.syncPlaces(metrics);
    const count = (this.sheet.drawings || []).length;
    let found = -1;
    for (const item of this.placed) {
      if (item.order >= count || item.order < found) continue;
      // A turned picture is found where it is drawn: the point is taken into the frame of its box.
      let px = x;
      let py = y;
      const turn = item.object.type === 'picture' && typeof item.object.rotation === 'number' && Number.isFinite(item.object.rotation) ? item.object.rotation : 0;
      if (turn) {
        const cx = item.x + item.width / 2;
        const cy = item.y + item.height / 2;
        const angle = (-turn * Math.PI) / 180;
        px = cx + (x - cx) * Math.cos(angle) - (y - cy) * Math.sin(angle);
        py = cy + (x - cx) * Math.sin(angle) + (y - cy) * Math.cos(angle);
      }
      if (px >= item.x && px <= item.x + item.width && py >= item.y && py <= item.y + item.height && wanted(item.object)) found = item.order;
    }
    return found;
  }

  /** The box of each object on the sheet, in pixels from the top-left corner of cell A1, bottom one first. */
  rects(metrics: SheetMetrics): (Rect & { object: DrawingObject })[] {
    this.syncPlaces(metrics);
    return this.placed
      .slice()
      .sort((a, b) => a.order - b.order)
      .map((item) => ({ x: item.x, y: item.y, width: item.width, height: item.height, object: item.object }));
  }

  private syncModel(): boolean {
    const sheet = this.sheet;
    if (!this.first && sheet.drawings === this.drawingsSeen && sheet.comments === this.commentsSeen && sheet.structuralEdits === this.editsSeen) return false;
    this.first = false;
    this.drawingsSeen = sheet.drawings;
    this.commentsSeen = sheet.comments;
    this.editsSeen = sheet.structuralEdits;
    this.metricsSeen = undefined;
    const edits = sheet.structuralEdits;
    const anchors: DrawingAnchor[] = [];
    const drawings = sheet.drawings || [];
    for (let i = 0; i < drawings.length; i++) {
      const drawing = drawings[i];
      anchors.push(drawing && drawing.anchor ? drawingAnchorNow(drawing, edits) : { kind: 'absolute' });
    }
    this.anchors = anchors;
    const placed: PlacedComment[] = [];
    const byCell = new Map<number, PlacedComment>();
    const comments = sheet.comments || [];
    for (let i = 0; i < comments.length; i++) {
      const comment = comments[i];
      if (!comment || !(comment.row >= 0 && comment.row < MAX_ROWS && comment.col >= 0 && comment.col < MAX_COLS)) continue;
      const cell = commentCellNow(comment, edits);
      if (!cell) continue;
      const key = cell.row * MAX_COLS + cell.col;
      if (byCell.has(key)) continue;
      const item = { row: cell.row, col: cell.col, comment };
      placed.push(item);
      byCell.set(key, item);
    }
    this.placedComments = placed;
    this.commentByCell = byCell;
    return true;
  }

  private syncPlaces(metrics: SheetMetrics): void {
    this.syncModel();
    if (this.metricsSeen === metrics) return;
    this.metricsSeen = metrics;
    const items: Placed[] = [];
    const drawings = this.sheet.drawings || [];
    for (let i = 0; i < drawings.length; i++) {
      const drawing = drawings[i];
      if (!drawing || !drawing.object || drawing.object.hidden) continue;
      const item = place(i, drawing.object, anchorRect(this.anchors[i], metrics));
      if (item) items.push(item);
    }
    // The boxes of comments that are always shown lie above the drawings.
    const edits = this.sheet.structuralEdits;
    for (let i = 0; this.notes && i < this.placedComments.length; i++) {
      const { row, col, comment } = this.placedComments[i];
      if (!comment.visible) continue;
      let rect: Rect;
      if (comment.box) rect = anchorRect(adjustDrawingAnchor(comment.box, commentEdits(comment, edits)), metrics);
      else rect = { x: metrics.cols.offset(col + 1) + 11, y: Math.max(metrics.rows.offset(row) - 7, 0), width: NOTE_WIDTH, height: NOTE_HEIGHT };
      const item = place(drawings.length + i, noteShape(comment), rect);
      if (item) items.push(item);
    }
    items.sort((a, b) => a.top - b.top || a.order - b.order);
    const tops: number[] = [];
    const reach: number[] = [];
    let far = -Infinity;
    for (let i = 0; i < items.length; i++) {
      tops.push(items[i].top);
      if (items[i].bottom > far) far = items[i].bottom;
      reach.push(far);
    }
    this.placed = items;
    this.tops = tops;
    this.reach = reach;
  }

  /** The objects that may show inside a rectangle of the sheet, bottom one first. */
  private visible(left: number, top: number, right: number, bottom: number): Placed[] {
    const { placed, tops, reach } = this;
    // First item whose predecessors and itself reach down to `top`.
    let low = 0;
    let high = placed.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (reach[mid] < top) low = mid + 1;
      else high = mid;
    }
    const start = low;
    // First item that starts below `bottom`.
    high = placed.length;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (tops[mid] <= bottom) low = mid + 1;
      else high = mid;
    }
    const found: Placed[] = [];
    for (let i = start; i < low; i++) {
      const item = placed[i];
      if (item.bottom >= top && item.right >= left && item.left <= right) found.push(item);
    }
    if (found.length > 1) found.sort((a, b) => a.order - b.order);
    return found;
  }

  /** Draws what shows inside the clip rectangle of the target. The state of the context is as it was afterwards. */
  paint(ctx: CanvasRenderingContext2D, target: DrawingTarget): void {
    this.syncPlaces(target.metrics);
    if (this.placed.length === 0) return;
    const { clip, originX, originY } = target;
    if (!(clip.width > 0 && clip.height > 0)) return;
    const items = this.visible(clip.x - originX, clip.y - originY, clip.x + clip.width - originX, clip.y + clip.height - originY);
    const marked = this.selected !== undefined ? this.boxOf(target.metrics, this.selected) : undefined;
    if (items.length === 0 && !marked) return;
    const dpr = target.devicePixelRatio > 0 && target.devicePixelRatio < Infinity ? target.devicePixelRatio : 1;
    // A font was registered or taken out: the broken lines hold the font strings and the widths from before.
    const fonts = fontsVersion();
    if (fonts !== this.layoutFonts) {
      this.layoutFonts = fonts;
      this.layouts = new WeakMap();
    }
    const frame: PaintFrame = { ctx, painter: new TextPainter(ctx), dpr, images: this.images, labels: this.labels, layouts: this.layouts, charts: this.charts, painted: 0, depth: 0 };
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.width, clip.height);
    ctx.clip();
    for (let i = 0; i < items.length && frame.painted < MAX_PAINTED; i++) {
      const item = items[i];
      try {
        paintObject(frame, item.object, originX + item.x, originY + item.y, item.width, item.height, 0);
      } catch (error) {
        // One object the canvas refuses does not take the frame with it; what it had saved is given back.
        while (frame.depth > 0) popState(frame);
        frame.painter.invalidate();
      }
    }
    if (marked) {
      // The frame of a picture is turned with it; that of a chart is not (a chart is not drawn turned by its frame in Excel).
      const object = this.selected !== undefined ? ((this.sheet.drawings || [])[this.selected] || ({} as SheetDrawing)).object : undefined;
      const turn = object && object.type === 'picture' && typeof object.rotation === 'number' && Number.isFinite(object.rotation) ? object.rotation : 0;
      paintSelection(ctx, marked, this.ghost, originX, originY, {
        rotation: turn,
        ghostRotation: this.ghostTurn !== undefined && Number.isFinite(this.ghostTurn) ? this.ghostTurn : turn,
        turnable: this.turnable && !this.cropping,
        turnGap: this.turnGap > 0 && this.turnGap < 1000 ? this.turnGap : ROTATE_HANDLE_GAP,
        cropping: this.cropping,
      });
    }
    ctx.restore();
  }
}

/** How the frame of the selected drawing is drawn. */
interface SelectionLook {
  /** Degrees the frame is turned about the middle of its box. */
  rotation: number;
  ghostRotation: number;
  /** A handle above the frame turns it, so many pixels above it. */
  turnable: boolean;
  turnGap: number;
  /** The picture is being cropped: the box the whole picture takes up, in the frame of the box. */
  cropping: Rect | undefined;
}

const CROP_COLOR = '#202124';

/** Turns the context by `degrees` about the middle of a box drawn at (originX + box.x, originY + box.y). */
function turnAbout(ctx: CanvasRenderingContext2D, box: Rect, originX: number, originY: number, degrees: number): void {
  if (!degrees) return;
  const cx = originX + box.x + box.width / 2;
  const cy = originY + box.y + box.height / 2;
  ctx.translate(cx, cy);
  ctx.rotate((degrees * Math.PI) / 180);
  ctx.translate(-cx, -cy);
}

/** The frame and the handles of the selected drawing, and the dashed frame of where a drag would put it. */
function paintSelection(ctx: CanvasRenderingContext2D, box: Rect, ghost: Rect | undefined, originX: number, originY: number, look?: SelectionLook): void {
  const rotation = look ? look.rotation : 0;
  ctx.save();
  ctx.globalAlpha = 1;
  ctx.lineWidth = 1;
  ctx.strokeStyle = SELECTION_COLOR;
  ctx.setLineDash([]);
  ctx.save();
  turnAbout(ctx, box, originX, originY, rotation);
  const cropping = look && look.cropping && look.cropping.width > 0 && look.cropping.height > 0 ? look.cropping : undefined;
  if (cropping) {
    // The whole picture, of which the box shows a part.
    ctx.setLineDash([2, 2]);
    ctx.strokeRect(Math.round(originX + cropping.x) + 0.5, Math.round(originY + cropping.y) + 0.5, Math.round(cropping.width) - 1, Math.round(cropping.height) - 1);
    ctx.setLineDash([]);
  }
  ctx.strokeRect(Math.round(originX + box.x) + 0.5, Math.round(originY + box.y) + 0.5, Math.round(box.width) - 1, Math.round(box.height) - 1);
  const half = HANDLE_SIZE / 2;
  if (look && look.turnable) {
    // The handle that turns the picture: a round knob on a stalk above the middle of the top side.
    const mx = Math.round(originX + box.x + box.width / 2) + 0.5;
    const top = Math.round(originY + box.y);
    ctx.beginPath();
    ctx.moveTo(mx, top);
    ctx.lineTo(mx, top - look.turnGap + half);
    ctx.stroke();
    ctx.beginPath();
    ctx.arc(mx, top - look.turnGap, half, 0, Math.PI * 2);
    ctx.fillStyle = '#FFFFFF';
    ctx.fill();
    ctx.stroke();
  }
  for (const handle of HANDLES) {
    const at = handlePoint(box, handle);
    const x = Math.round(originX + at.x - half);
    const y = Math.round(originY + at.y - half);
    // The handles that crop are dark: they cut, they do not size.
    ctx.fillStyle = cropping ? CROP_COLOR : '#FFFFFF';
    ctx.fillRect(x, y, HANDLE_SIZE, HANDLE_SIZE);
    ctx.strokeRect(x + 0.5, y + 0.5, HANDLE_SIZE - 1, HANDLE_SIZE - 1);
  }
  ctx.restore();
  if (ghost && ghost.width > 0 && ghost.height > 0) {
    turnAbout(ctx, ghost, originX, originY, look ? look.ghostRotation : 0);
    ctx.setLineDash([4, 3]);
    ctx.strokeRect(Math.round(originX + ghost.x) + 0.5, Math.round(originY + ghost.y) + 0.5, Math.round(ghost.width) - 1, Math.round(ghost.height) - 1);
  }
  ctx.restore();
}

// -------------------------------------------------------------------- painting

interface PaintFrame {
  ctx: CanvasRenderingContext2D;
  painter: TextPainter;
  dpr: number;
  images: SheetDrawingsOptions['images'];
  labels: SheetDrawingLabels;
  layouts: WeakMap<DrawingText, TextLayout>;
  charts: ChartRenderer;
  painted: number;
  /** save() calls of the frame not yet matched by a restore(). */
  depth: number;
}

function pushState(frame: PaintFrame): void {
  frame.ctx.save();
  frame.depth++;
}

function popState(frame: PaintFrame): void {
  frame.ctx.restore();
  frame.depth--;
}

const safeColor = (color: string | undefined): string | undefined => (typeof color === 'string' && HEX_COLOR.test(color) ? color : undefined);

function paintObject(frame: PaintFrame, object: DrawingObject, x: number, y: number, width: number, height: number, depth: number): void {
  if (!object || object.hidden || frame.painted >= MAX_PAINTED) return;
  frame.painted++;
  const { ctx } = frame;
  const rotation = typeof object.rotation === 'number' && Number.isFinite(object.rotation) ? object.rotation : 0;
  const flipH = object.flipH === true;
  const flipV = object.flipV === true;
  const transformed = rotation !== 0 || flipH || flipV;
  if (transformed || object.type === 'group') {
    pushState(frame);
    frame.painter.invalidate();
  }
  if (rotation !== 0) {
    ctx.translate(x + width / 2, y + height / 2);
    ctx.rotate((rotation * Math.PI) / 180);
    ctx.translate(-(x + width / 2), -(y + height / 2));
  }
  if (object.type === 'group') {
    // A flipped group shows its members mirrored, each at the mirrored place.
    if (flipH || flipV) {
      ctx.translate(x + width / 2, y + height / 2);
      ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
      ctx.translate(-(x + width / 2), -(y + height / 2));
    }
    if (depth < MAX_DEPTH) {
      const children = object.children || [];
      for (let i = 0; i < children.length; i++) {
        const member = children[i];
        const box = member && member.box;
        if (!box) continue;
        const mw = box.width * width;
        const mh = box.height * height;
        if (!(mw >= 0 && mh >= 0 && mw < 1e7 && mh < 1e7)) continue;
        paintObject(frame, member, x + box.x * width, y + box.y * height, mw, mh, depth + 1);
      }
    }
  } else {
    const mirror = (flipH || flipV) && object.type !== 'placeholder' && object.type !== 'chart';
    if (mirror) {
      pushState(frame);
      ctx.translate(x + width / 2, y + height / 2);
      ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
      ctx.translate(-(x + width / 2), -(y + height / 2));
    }
    let drawn = true;
    if (object.type === 'picture') drawn = paintPicture(frame, object, x, y, width, height);
    else if (object.type === 'shape') paintShapeBody(frame, object, x, y, width, height);
    if (mirror) {
      popState(frame);
      frame.painter.invalidate();
    }
    // What stands for something that is not drawn is never mirrored: its words are to be read.
    if (object.type === 'chart') paintChart(frame, object, x, y, width, height);
    else if (object.type === 'placeholder') paintPlaceholder(frame, x, y, width, height, object.what === 'chart' ? frame.labels.chart : object.what === 'diagram' ? frame.labels.diagram : frame.labels.object);
    else if (!drawn && object.type === 'picture') paintPlaceholder(frame, x, y, width, height, object.description || object.name || frame.labels.picture);
    // Text is turned with its shape but never mirrored.
    if (object.type === 'shape' && object.text) paintText(frame, object.text, x, y, width, height);
  }
  if (transformed || object.type === 'group') {
    popState(frame);
    frame.painter.invalidate();
  }
}

/** Draws the picture, or its place while it is on its way. False for a picture that will not come: the caller draws a placeholder. */
function paintPicture(frame: PaintFrame, picture: Extract<DrawingObject, { type: 'picture' }>, x: number, y: number, width: number, height: number): boolean {
  const { ctx, images } = frame;
  const part = picture.media;
  const crop = picture.crop;
  const shownX = crop ? Math.max(1 - crop.left - crop.right, 0.001) : 1;
  const shownY = crop ? Math.max(1 - crop.top - crop.bottom, 0.001) : 1;
  let image: DecodedImage | undefined;
  if (part !== undefined && images) {
    try {
      // The whole picture has to have the pixels its shown part needs.
      image = images.get(part, Math.ceil((width * frame.dpr) / shownX), Math.ceil((height * frame.dpr) / shownY));
    } catch (error) {
      image = undefined;
    }
  }
  if (image && image.width > 0 && image.height > 0) {
    const sx = crop ? crop.left * image.width : 0;
    const sy = crop ? crop.top * image.height : 0;
    ctx.drawImage(image.source, sx, sy, image.width * shownX, image.height * shownY, x, y, width, height);
    return true;
  }
  const pending = part !== undefined && images !== undefined && (typeof images.status !== 'function' || images.status(part) === 'pending');
  if (pending) {
    // Until the picture is there: its place, lightly.
    frame.painter.setFill(LOADING_FILL);
    ctx.fillRect(x, y, width, height);
    return true;
  }
  return false;
}

/** Draws a chart in its frame; one that cannot be laid out or painted is shown as the placeholder of a chart. */
function paintChart(frame: PaintFrame, object: Extract<DrawingObject, { type: 'chart' }>, x: number, y: number, width: number, height: number): void {
  // Whole pixels: the lines of a chart are thin and many. The size is rounded by itself and not as the distance
  // between two rounded edges: that one changes by a pixel as the sheet is scrolled by parts of a pixel, and
  // every such change is a chart laid out again.
  const px = Math.round(x);
  const py = Math.round(y);
  const pw = Math.round(width);
  const ph = Math.round(height);
  let drawn = false;
  try {
    if (pw > 0 && ph > 0) frame.charts.paint(frame.ctx, object.chart, px, py, pw, ph);
    drawn = true;
  } catch (error) {
    // Whatever a chart states, it costs no more than its own picture.
  }
  // Measuring and painting set fonts and fills of their own.
  frame.painter.invalidate();
  if (!drawn) paintPlaceholder(frame, x, y, width, height, frame.labels.chart);
}

function paintPlaceholder(frame: PaintFrame, x: number, y: number, width: number, height: number, label: string): void {
  const { ctx } = frame;
  frame.painter.setFill(PLACEHOLDER_FILL);
  ctx.fillRect(x, y, width, height);
  ctx.strokeStyle = PLACEHOLDER_BORDER;
  ctx.lineWidth = 1;
  ctx.setLineDash([]);
  if (width > 1 && height > 1) ctx.strokeRect(x + 0.5, y + 0.5, width - 1, height - 1);
  if (width < 28 || height < 16 || typeof label !== 'string' || label === '') return;
  let text = label.length > 200 ? label.slice(0, 200) : label;
  const room = width - 8;
  if (frame.painter.measure(LABEL_FONT, text) > room) {
    while (text.length > 1 && frame.painter.measure(LABEL_FONT, text + '…') > room) text = text.slice(0, text.length - 1);
    text += '…';
    if (frame.painter.measure(LABEL_FONT, text) > room) return;
  }
  frame.painter.setFont(LABEL_FONT);
  frame.painter.setFill(PLACEHOLDER_TEXT);
  const textWidth = frame.painter.measure(LABEL_FONT, text);
  ctx.fillText(text, x + (width - textWidth) / 2, y + height / 2 + 4);
}

/** Presets drawn with rounded corners. */
const ROUNDED: { [prst: string]: true } = { roundRect: true, round1Rect: true, round2SameRect: true, round2DiagRect: true, flowChartAlternateProcess: true, flowChartTerminator: true };
const ELLIPSES: { [prst: string]: true } = { ellipse: true, flowChartConnector: true };
const KAPPA = 0.5522847498;

function shapePath(ctx: CanvasRenderingContext2D, geometry: string, x: number, y: number, width: number, height: number): void {
  ctx.beginPath();
  if (Object.prototype.hasOwnProperty.call(ELLIPSES, geometry)) {
    const rx = width / 2;
    const ry = height / 2;
    const cx = x + rx;
    const cy = y + ry;
    const ox = rx * KAPPA;
    const oy = ry * KAPPA;
    ctx.moveTo(cx - rx, cy);
    ctx.bezierCurveTo(cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry);
    ctx.bezierCurveTo(cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy);
    ctx.bezierCurveTo(cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry);
    ctx.bezierCurveTo(cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy);
    ctx.closePath();
    return;
  }
  if (Object.prototype.hasOwnProperty.call(ROUNDED, geometry)) {
    // A sixth of the shorter side, the corner Office gives a rounded rectangle to start with.
    const radius = Math.min(width, height) * (geometry === 'flowChartTerminator' ? 0.5 : 1 / 6);
    ctx.moveTo(x + radius, y);
    ctx.lineTo(x + width - radius, y);
    ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
    ctx.lineTo(x + width, y + height - radius);
    ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
    ctx.lineTo(x + radius, y + height);
    ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
    ctx.lineTo(x, y + radius);
    ctx.quadraticCurveTo(x, y, x + radius, y);
    ctx.closePath();
    return;
  }
  ctx.rect(x, y, width, height);
}

function paintShapeBody(frame: PaintFrame, shape: DrawingShape, x: number, y: number, width: number, height: number): void {
  const { ctx } = frame;
  const line = shape.line;
  const stroke = line ? safeColor(line.color) : undefined;
  const lineWidth = line && line.width > 0 && line.width < Infinity ? Math.min(line.width, 32) : 1;
  if (stroke) {
    ctx.strokeStyle = stroke;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(line && line.dash ? [lineWidth * 4, lineWidth * 3] : []);
  }
  if (shape.connector) {
    if (!stroke) return;
    // From corner to corner; a flip, applied by the caller, gives the other diagonal.
    ctx.beginPath();
    // A hairline that runs along the pixel grid is put on the pixel centres.
    const snap = lineWidth <= 1 ? 0.5 : 0;
    if (height === 0) {
      ctx.moveTo(x, Math.floor(y) + snap);
      ctx.lineTo(x + width, Math.floor(y) + snap);
    } else if (width === 0) {
      ctx.moveTo(Math.floor(x) + snap, y);
      ctx.lineTo(Math.floor(x) + snap, y + height);
    } else {
      ctx.moveTo(x, y);
      ctx.lineTo(x + width, y + height);
    }
    ctx.stroke();
    ctx.setLineDash([]);
    return;
  }
  const fill = safeColor(shape.fill);
  if (!fill && !stroke) return;
  // Shapes whose outline is not known are drawn as the box they fill.
  const px = Math.round(x);
  const py = Math.round(y);
  const pw = Math.max(Math.round(x + width) - px, 1);
  const ph = Math.max(Math.round(y + height) - py, 1);
  if (fill) {
    shapePath(ctx, shape.geometry, px, py, pw, ph);
    frame.painter.setFill(fill);
    ctx.fill();
  }
  if (stroke) {
    const inset = lineWidth <= 1 ? 0.5 : 0;
    shapePath(ctx, shape.geometry, px + inset, py + inset, Math.max(pw - inset * 2, 0), Math.max(ph - inset * 2, 0));
    ctx.stroke();
    ctx.setLineDash([]);
  }
}

// ------------------------------------------------------------------------ text

function runFont(run: DrawingTextRun): { font: string; size: number } {
  const points = typeof run.size === 'number' && run.size > 0 && run.size < Infinity ? run.size : DEFAULT_TEXT_POINTS;
  const font: Font = { name: run.font, size: points, bold: run.bold === true, italic: run.italic === true, strike: false, underline: 'none' };
  return { font: cssFont(font), size: Math.round(((Math.min(Math.max(points, 1), 409) * 96) / 72) * 100) / 100 };
}

function layoutText(frame: PaintFrame, text: DrawingText, width: number): TextLayout {
  const cached = frame.layouts.get(text);
  if (cached && cached.width === width) return cached;
  const wrap = text.wrap !== false;
  const lines: TextLine[] = [];
  let pieces = 0;
  const paragraphs = text.paragraphs || [];
  for (let p = 0; p < paragraphs.length && lines.length < MAX_TEXT_LINES && pieces < MAX_TEXT_PIECES; p++) {
    const paragraph = paragraphs[p];
    const align = paragraph.align === 'center' || paragraph.align === 'right' ? paragraph.align : 'left';
    let line: TextLine = { pieces: [], width: 0, height: 0, align };
    let lastSize = Math.round((DEFAULT_TEXT_POINTS * 96) / 72);
    /** The line was started by breaking the one before it. */
    let broken = false;
    const close = () => {
      // Spaces at the end of a line take no room.
      while (line.pieces.length > 0 && line.pieces[line.pieces.length - 1].space) line.width -= (line.pieces.pop() as TextPiece).width;
      if (line.height === 0) line.height = lastSize * LINE_HEIGHT;
      lines.push(line);
      line = { pieces: [], width: 0, height: 0, align };
    };
    const runs = paragraph.runs || [];
    for (let r = 0; r < runs.length && lines.length < MAX_TEXT_LINES && pieces < MAX_TEXT_PIECES; r++) {
      const run = runs[r];
      if (!run || typeof run.text !== 'string' || run.text === '') continue;
      const { font, size } = runFont(run);
      lastSize = size;
      const color = safeColor(run.color) || '#000000';
      const add = (part: string, space: boolean) => {
        const partWidth = frame.painter.measure(font, part);
        line.pieces.push({ text: part, font, color, size, width: partWidth, underline: run.underline === true, strike: run.strike === true, space });
        line.width += partWidth;
        if (size * LINE_HEIGHT > line.height) line.height = size * LINE_HEIGHT;
        pieces++;
      };
      const source = run.text.length > 32767 ? run.text.slice(0, 32767) : run.text;
      const length = source.length;
      for (let i = 0; i < length && lines.length < MAX_TEXT_LINES && pieces < MAX_TEXT_PIECES; ) {
        const code = source.charCodeAt(i);
        if (code === 10 || code === 13) {
          if (code === 13 && source.charCodeAt(i + 1) === 10) i++;
          i++;
          if (size * LINE_HEIGHT > line.height && line.pieces.length === 0) line.height = size * LINE_HEIGHT;
          close();
          broken = false;
          continue;
        }
        let end = i;
        const space = code === 32 || code === 9;
        while (end < length) {
          const c = source.charCodeAt(end);
          if (c === 10 || c === 13 || (c === 32 || c === 9) !== space) break;
          end++;
        }
        const part = space ? source.slice(i, end).replace(/\t/g, '    ') : source.slice(i, end);
        i = end;
        if (space) {
          // A line never starts with the spaces it was broken at.
          if (line.pieces.length > 0 || !broken) add(part, true);
          continue;
        }
        const partWidth = frame.painter.measure(font, part);
        if (wrap && line.pieces.length > 0 && line.width + partWidth > width) {
          close();
          broken = true;
        }
        if (wrap && partWidth > width && width > 0) {
          // Wider than a whole line: broken between characters.
          const breaks = wrapLines(part, width, (piece) => frame.painter.measure(font, piece));
          for (let b = 0; b < breaks.length && lines.length < MAX_TEXT_LINES; b += 2) {
            if (b > 0) {
              close();
              broken = true;
            }
            add(part.slice(breaks[b], breaks[b + 1]), false);
          }
        } else add(part, false);
      }
    }
    close();
  }
  let height = 0;
  for (let i = 0; i < lines.length; i++) height += lines[i].height;
  const layout = { width, lines, height };
  frame.layouts.set(text, layout);
  return layout;
}

function paintText(frame: PaintFrame, text: DrawingText, x: number, y: number, width: number, height: number): void {
  const { ctx } = frame;
  const inset = (index: number): number => {
    const value = text.insets ? text.insets[index] : 0;
    return typeof value === 'number' && value >= 0 && value < 4000 ? value : 0;
  };
  const left = x + inset(0);
  const top = y + inset(1);
  const innerWidth = Math.max(width - inset(0) - inset(2), 0);
  const innerHeight = Math.max(height - inset(1) - inset(3), 0);
  if (!(width > 2 && height > 2)) return;
  const layout = layoutText(frame, text, Math.floor(innerWidth));
  if (layout.lines.length === 0) return;
  pushState(frame);
  ctx.beginPath();
  ctx.rect(x, y, width, height);
  ctx.clip();
  frame.painter.invalidate();
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  let lineTop = top;
  // Text taller than its shape starts at the top, whatever it asks for: the beginning is what is read.
  if (layout.height < innerHeight) {
    if (text.anchor === 'middle') lineTop = top + (innerHeight - layout.height) / 2;
    else if (text.anchor === 'bottom') lineTop = top + innerHeight - layout.height;
  }
  for (let i = 0; i < layout.lines.length; i++) {
    const line = layout.lines[i];
    if (lineTop > y + height) break;
    if (lineTop + line.height >= y) {
      let penX = line.align === 'center' ? left + (innerWidth - line.width) / 2 : line.align === 'right' ? left + innerWidth - line.width : left;
      // The baseline of the tallest piece serves the whole line.
      const baseline = Math.round(lineTop + line.height * 0.8);
      for (let p = 0; p < line.pieces.length; p++) {
        const piece = line.pieces[p];
        if (!piece.space) {
          frame.painter.setFont(piece.font);
          frame.painter.setFill(piece.color);
          ctx.fillText(piece.text, penX, baseline);
        }
        if (piece.underline || piece.strike) {
          frame.painter.setFill(piece.color);
          const thickness = Math.max(1, Math.round(piece.size / 14));
          if (piece.underline) ctx.fillRect(penX, baseline + Math.max(1, Math.round(piece.size / 10)), piece.width, thickness);
          if (piece.strike) ctx.fillRect(penX, baseline - Math.round(piece.size * 0.3), piece.width, thickness);
        }
        penX += piece.width;
      }
    }
    lineTop += line.height;
  }
  popState(frame);
  frame.painter.invalidate();
}
