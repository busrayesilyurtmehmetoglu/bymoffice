// Pictures as the editor changes them: a picture added from its bytes, moved
// or resized, and deleted. Like the other commands these functions only build
// operations ("setDrawings", as for charts: see charts.ts); the caller applies
// them through a History so that each is one undo step.
//
// - A picture that is added is in no file: its bytes are in the model
//   (DrawingPicture.data) and the writer gives them a part of their own.
// - A picture of the file is moved and deleted; only its anchor changes in the
//   drawing part, everything else of the part stays as the file has it.
// - A picture keeps the kind of anchor it has: one that moves and sizes with
//   its cells, one that moves with its top-left cell, or one that stands where
//   it is whatever the cells do (DrawingAnchor.kind and editAs). Rows and
//   columns inserted or deleted later move it by those rules (see
//   drawingAnchorNow()).
//
// What kind of picture bytes are is read from the bytes, never from a name or
// from what a caller says they are; what is not a PNG, JPEG, GIF or BMP
// picture with a readable size is refused, as is one too large to be shown.

import type { SheetMetrics } from '../layout/sheet-metrics';
import { createSheetMetrics } from '../layout/sheet-metrics';
import { isAddedPicture, newPictureKey, PICTURE_EXTENSIONS } from '../model/sheet/pictures';
import { MAX_COLS, MAX_ROWS, type CellRange, type DrawingAnchor, type DrawingMarker, type DrawingPicture, type PictureData, type Sheet, type SheetDrawing, type Workbook } from '../model/sheet/types';
import { anchorRect, drawingAnchorNow } from '../render/sheet-drawings';
import { declaredImageSize, sniffImageType } from '../render/images';
import { checkedAnchor, MAX_SHEET_DRAWINGS } from './charts';
import { withCommand, type Op } from './ops';

/** The kinds of picture that are added, by their media type. */
export const PICTURE_TYPES = Object.keys(PICTURE_EXTENSIONS);
/** The largest picture that is added, in bytes. */
export const MAX_PICTURE_BYTES = 32 * 1024 * 1024;
/** The longest side, and the most pixels, a picture that is added may have: what is more cannot be shown (see render/images.ts). */
export const MAX_PICTURE_SIDE = 16384;
export const MAX_PICTURE_PIXELS = 50000000;
/** What the pictures added to one workbook hold together, in bytes. */
export const MAX_WORKBOOK_PICTURE_BYTES = 256 * 1024 * 1024;
/** Longest name and alternative text of a picture. */
export const MAX_PICTURE_TEXT = 1024;
/** The largest a picture is shown when it is added without a place, in pixels: a larger one is made smaller, its sides in proportion. UNSURE: Excel adds a picture at the size its own resolution gives it. */
export const DEFAULT_PICTURE_WIDTH = 800;
export const DEFAULT_PICTURE_HEIGHT = 600;
/** The smallest a picture is made by the commands, in pixels. */
export const MIN_PICTURE_SIZE = 4;
const EMU_PER_PIXEL = 9525;
const MAX_EMU = 2e12;

/** A picture of a sheet as the commands see it. */
export interface PictureInfo {
  /** The place of the picture in Sheet.drawings: what the commands name it by. It changes when a drawing before it is deleted. */
  id: number;
  /** Where the picture is now. */
  anchor: DrawingAnchor;
  name?: string;
  /** Alternative text. */
  description?: string;
  /** The picture is in the file the workbook was read from. */
  ofFile: boolean;
  /** Whether it can be deleted (a picture that shares its place in the drawing with another object cannot). */
  removable: boolean;
}

/** What is known of picture bytes once they are checked. */
export interface PictureFacts {
  type: string;
  width: number;
  height: number;
}

// ------------------------------------------------------------------ the bytes

const isMarker = (bytes: Uint8Array, at: number, text: string): boolean => {
  for (let i = 0; i < text.length; i++) if (bytes[at + i] !== text.charCodeAt(i)) return false;
  return true;
};

/**
 * What picture bytes are, or why they are not a picture that can be added: a message. `mime`,
 * when given, has to be what the bytes are ("image/jpg" is taken for "image/jpeg").
 * Only the header is read: a picture whose data is damaged behind a sound header is added, and
 * is shown as the placeholder of a picture.
 */
export function checkPicture(bytes: unknown, mime?: string): PictureFacts | string {
  if (!(bytes instanceof Uint8Array) || bytes.length === 0) return 'A picture is given as its bytes';
  if (bytes.length > MAX_PICTURE_BYTES) return `A picture cannot be larger than ${MAX_PICTURE_BYTES / 1024 / 1024} MB`;
  let wanted: string | undefined;
  if (mime !== undefined && mime !== null) {
    wanted = String(mime).trim().toLowerCase();
    if (wanted === 'image/jpg' || wanted === 'image/pjpeg') wanted = 'image/jpeg';
    if (wanted === 'image/x-ms-bmp' || wanted === 'image/x-bmp') wanted = 'image/bmp';
    if (PICTURE_TYPES.indexOf(wanted) < 0) return 'A picture is a PNG, JPEG, GIF or BMP picture';
  }
  const type = sniffImageType(bytes);
  if (type === undefined || PICTURE_TYPES.indexOf(type) < 0) return 'The bytes are not a PNG, JPEG, GIF or BMP picture';
  if (wanted !== undefined && wanted !== type) return `The bytes are a picture of another kind than ${wanted}`;
  // The headers the size is read from, as far as they can be told apart from noise.
  if (type === 'image/png' && !(bytes.length >= 33 && isMarker(bytes, 12, 'IHDR'))) return 'The picture cannot be read';
  if (type === 'image/gif' && !(isMarker(bytes, 3, '87a') || isMarker(bytes, 3, '89a'))) return 'The picture cannot be read';
  if (type === 'image/bmp') {
    const header = bytes.length >= 26 ? bytes[14] | (bytes[15] << 8) | (bytes[16] << 16) | (bytes[17] << 24) : 0;
    if (header !== 12 && header !== 40 && header !== 52 && header !== 56 && header !== 64 && header !== 108 && header !== 124) return 'The picture cannot be read';
  }
  const size = declaredImageSize(bytes, type);
  if (!size || !(size.width > 0) || !(size.height > 0)) return 'The picture cannot be read';
  if (size.width > MAX_PICTURE_SIDE || size.height > MAX_PICTURE_SIDE || size.width * size.height > MAX_PICTURE_PIXELS) {
    return `A picture cannot be larger than ${MAX_PICTURE_SIDE} pixels a side, nor have more than ${MAX_PICTURE_PIXELS / 1000000} million pixels`;
  }
  return { type, width: size.width, height: size.height };
}

// ------------------------------------------------------------------ helpers

function sheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change the pictures of a ${sheet.kind}`);
  return sheet;
}

/**
 * Whether the objects of a sheet are protected: the sheet is protected (sheetProtection
 * sheet="1") and its objects with it (objects="1"). UNSURE: the attributes are read as the
 * standard says; how Excel reads a sheet protected without `objects` was not checked.
 */
export function objectsProtected(sheet: Sheet): boolean {
  const attrs = sheet.protection && sheet.protection.attrs;
  if (!attrs) return false;
  const on = (name: string): boolean => attrs.some(([key, value]) => key === name && (value === '1' || value === 'true'));
  return on('sheet') && on('objects');
}

/** Why no picture can be added to a sheet, moved or deleted on it; undefined when that can be done. */
export function picturesProblem(workbook: Workbook, sheetIndex: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot change the pictures of a ${sheet.kind}`;
  if (sheet.drawingsFixed !== undefined && sheet.part !== '') return `Cannot change the pictures of this sheet: ${sheet.drawingsFixed}`;
  if (objectsProtected(sheet)) return 'Cannot change the pictures of this sheet: it is protected';
  return undefined;
}

function placedSheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = sheetOf(workbook, sheetIndex);
  const problem = picturesProblem(workbook, sheetIndex);
  if (problem) throw new RangeError(problem);
  return sheet;
}

/** The picture of a sheet a command is about. Throws RangeError when `id` names none. */
function pictureAt(sheet: Sheet, id: number): { drawing: SheetDrawing; object: DrawingPicture } {
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'picture') throw new RangeError(`No picture at index ${id}`);
  return { drawing, object: drawing.object };
}

/** A drawing of the file as it was read, for one that was moved since; the drawing itself for one of the file that was not; undefined for one that is in no file. */
function fileDrawing(drawing: SheetDrawing): SheetDrawing | undefined {
  if (drawing.from) return drawing.from;
  return drawing.since === undefined ? drawing : undefined;
}

function setDrawings(sheetIndex: number, list: SheetDrawing[]): Op[] {
  return [{ type: 'setDrawings', sheet: sheetIndex, drawings: list.length > 0 ? list : undefined }];
}

function checkedText(text: unknown, what: string): string | undefined {
  if (text === undefined || text === null) return undefined;
  const value = String(text).replace(/\r\n?/g, '\n');
  if (value.trim() === '') return undefined;
  if (value.length > MAX_PICTURE_TEXT) throw new RangeError(`The ${what} of a picture cannot be longer than ${MAX_PICTURE_TEXT} characters`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.test(value)) throw new RangeError(`The ${what} of a picture holds a character a file cannot hold`);
  return value;
}

/** What the pictures that were added to a workbook hold, each set of bytes counted once. */
function addedBytes(workbook: Workbook): number {
  const seen = new Set<Uint8Array>();
  let total = 0;
  for (const sheet of workbook.sheets) {
    for (const drawing of sheet.drawings || []) {
      const object = drawing ? drawing.object : undefined;
      if (!isAddedPicture(object) || seen.has(object.data.bytes)) continue;
      seen.add(object.data.bytes);
      total += object.data.bytes.length;
    }
  }
  return total;
}

/** The name a picture that is added to a sheet gets: "Picture" and one more than the pictures the sheet has. */
function nextName(sheet: Sheet): string {
  let number = 1;
  for (const drawing of sheet.drawings || []) if (drawing && drawing.object && drawing.object.type === 'picture') number++;
  return `Picture ${number}`;
}

// ------------------------------------------------------------------ anchors

const clampEmu = (value: number): number => Math.max(0, Math.min(Math.round(value * EMU_PER_PIXEL), MAX_EMU));

/** The corner of a cell a point of the sheet (pixels from the corner of cell A1) is in, with how far into it. */
function markerAt(metrics: SheetMetrics, x: number, y: number): DrawingMarker {
  const col = Math.min(metrics.cols.indexAt(Math.max(0, x)), MAX_COLS - 1);
  const row = Math.min(metrics.rows.indexAt(Math.max(0, y)), MAX_ROWS - 1);
  return { col, colOff: clampEmu(Math.max(0, x - metrics.cols.offset(col))), row, rowOff: clampEmu(Math.max(0, y - metrics.rows.offset(row))) };
}

/**
 * The anchor of the kind `like` is (with what it says of rows and columns that move: editAs)
 * that covers a rectangle of the sheet, in pixels from the corner of cell A1. The rectangle is
 * kept on the sheet and at least MIN_PICTURE_SIZE pixels each way.
 */
export function anchorLike(like: DrawingAnchor, metrics: SheetMetrics, rect: { x: number; y: number; width: number; height: number }): DrawingAnchor {
  const finite = (n: number, fallback: number): number => (typeof n === 'number' && isFinite(n) ? n : fallback);
  const limitX = metrics.cols.total;
  const limitY = metrics.rows.total;
  const width = Math.max(Math.min(finite(rect.width, MIN_PICTURE_SIZE), limitX), MIN_PICTURE_SIZE);
  const height = Math.max(Math.min(finite(rect.height, MIN_PICTURE_SIZE), limitY), MIN_PICTURE_SIZE);
  const x = Math.max(0, Math.min(finite(rect.x, 0), limitX - width));
  const y = Math.max(0, Math.min(finite(rect.y, 0), limitY - height));
  if (like.kind === 'absolute') return { kind: 'absolute', x: clampEmu(x), y: clampEmu(y), width: clampEmu(width), height: clampEmu(height) };
  if (like.kind === 'oneCell') return { kind: 'oneCell', from: markerAt(metrics, x, y), width: clampEmu(width), height: clampEmu(height) };
  const out: DrawingAnchor = { kind: 'twoCell', from: markerAt(metrics, x, y), to: markerAt(metrics, x + width, y + height) };
  if (like.editAs === 'oneCell' || like.editAs === 'absolute') out.editAs = like.editAs;
  return out;
}

/** Where a new picture of a size (pixels) stands at a cell: it moves with that cell and keeps its size, as Excel adds a picture. */
export function pictureAnchorAt(sheet: Sheet, row: number, col: number, width: number, height: number): DrawingAnchor {
  const metrics = createSheetMetrics(sheet);
  const r = Math.max(0, Math.min(Math.floor(row) || 0, MAX_ROWS - 1));
  const c = Math.max(0, Math.min(Math.floor(col) || 0, MAX_COLS - 1));
  return anchorLike({ kind: 'twoCell', editAs: 'oneCell' }, metrics, { x: metrics.cols.offset(c), y: metrics.rows.offset(r), width, height });
}

/** The size a picture of `width` by `height` pixels is shown at when it is added without a place. */
export function defaultPictureSize(width: number, height: number): { width: number; height: number } {
  const scale = Math.min(1, DEFAULT_PICTURE_WIDTH / width, DEFAULT_PICTURE_HEIGHT / height);
  return { width: Math.max(MIN_PICTURE_SIZE, Math.round(width * scale)), height: Math.max(MIN_PICTURE_SIZE, Math.round(height * scale)) };
}

const isRange = (value: unknown): value is CellRange => {
  const range = value as CellRange;
  return !!range && typeof range === 'object' && (range as { kind?: unknown }).kind === undefined && [range.top, range.left, range.bottom, range.right].every((n) => typeof n === 'number');
};

// ------------------------------------------------------------------ commands

/** Optional facts of a picture that is added. */
export interface PictureOptions {
  /** Alternative text. */
  description?: string;
  /** Default "Picture" and a number. */
  name?: string;
}

/**
 * Adds a picture to a sheet, above its other drawings. `mime` (optional) has to be what the bytes
 * are. `anchor` is where it stands (an anchor, or the cells it covers); without one it stands at
 * cell A1 in its own size (no larger than DEFAULT_PICTURE_WIDTH by DEFAULT_PICTURE_HEIGHT), and
 * moves with that cell. Returns the operations; the picture is the last entry of Sheet.drawings
 * once they are applied. Throws RangeError for bytes that are not a picture that can be added
 * (see checkPicture()), a sheet that takes no picture, and too many pictures.
 */
export function addPicture(workbook: Workbook, sheetIndex: number, bytes: Uint8Array, mime?: string, anchor?: DrawingAnchor | CellRange, options: PictureOptions = {}): Op[] {
  // The bytes are not what a collaborator would be sent: the command is named, the list is not shared (see collab/wire.ts).
  return withCommand(addPicturePlain(workbook, sheetIndex, bytes, mime, anchor, options), 'addPicture', []);
}

function addPicturePlain(workbook: Workbook, sheetIndex: number, bytes: Uint8Array, mime: string | undefined, anchor: DrawingAnchor | CellRange | undefined, options: PictureOptions): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const facts = checkPicture(bytes, mime);
  if (typeof facts === 'string') throw new RangeError(facts);
  const drawings = sheet.drawings || [];
  if (drawings.length >= MAX_SHEET_DRAWINGS) throw new RangeError('The sheet has too many drawings');
  if (addedBytes(workbook) + bytes.length > MAX_WORKBOOK_PICTURE_BYTES) throw new RangeError(`The pictures added to a workbook cannot hold more than ${MAX_WORKBOOK_PICTURE_BYTES / 1024 / 1024} MB`);
  const opts = options && typeof options === 'object' ? options : {};
  const description = checkedText(opts.description, 'alternative text');
  const name = checkedText(opts.name, 'name');
  let place: DrawingAnchor;
  if (anchor === undefined || anchor === null) {
    const size = defaultPictureSize(facts.width, facts.height);
    place = pictureAnchorAt(sheet, 0, 0, size.width, size.height);
  } else {
    place = checkedAnchor(anchor, 'a picture');
    // A range of cells: the picture moves and sizes with them.
  }
  // The bytes are copied: the caller's array may change after this.
  const data: PictureData = { bytes: bytes.slice(), type: facts.type };
  const object: DrawingPicture = { type: 'picture', media: newPictureKey(), data, name: name !== undefined ? name : nextName(sheet) };
  if (description !== undefined) object.description = description;
  const added: SheetDrawing = { anchor: place, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  return setDrawings(sheetIndex, drawings.concat([added]));
}

/**
 * Moves or resizes a picture: `anchor` is where it stands from now on - an anchor of the kind the
 * picture has (a picture keeps its kind: see anchorLike()), or the cells it covers (made into an
 * anchor of its kind). A twoCell anchor that says nothing of `editAs` keeps the picture's.
 * Returns no operation when the picture is there already. Throws RangeError for an anchor of
 * another kind.
 */
export function movePicture(workbook: Workbook, sheetIndex: number, id: number, anchor: DrawingAnchor | CellRange): Op[] {
  return withCommand(movePicturePlain(workbook, sheetIndex, id, anchor), 'movePicture', []);
}

function movePicturePlain(workbook: Workbook, sheetIndex: number, id: number, anchor: DrawingAnchor | CellRange): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = pictureAt(sheet, id);
  const now = drawingAnchorNow(drawing, sheet.structuralEdits);
  let wanted: DrawingAnchor;
  if (isRange(anchor)) {
    const cells = checkedAnchor(anchor, 'a picture');
    const metrics = createSheetMetrics(sheet);
    wanted = anchorLike(now, metrics, anchorRect(cells, metrics));
  } else {
    wanted = checkedAnchor(anchor, 'a picture');
    if (wanted.kind !== now.kind) throw new RangeError('A picture keeps the kind of its anchor');
    if (wanted.kind === 'twoCell' && (anchor as DrawingAnchor).editAs === undefined && (now.editAs === 'oneCell' || now.editAs === 'absolute')) wanted.editAs = now.editAs;
  }
  if (JSON.stringify(wanted) === JSON.stringify(now)) return [];
  const count = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
  const next: SheetDrawing = { anchor: wanted, object: drawing.object, since: count };
  const from = fileDrawing(drawing);
  if (from) next.from = from;
  const list = (sheet.drawings as SheetDrawing[]).slice();
  list[id] = next;
  return setDrawings(sheetIndex, list);
}

/**
 * Moves or resizes a picture to cover a rectangle of the sheet (pixels from the corner of cell A1,
 * as the view measures them), with the kind of anchor it has.
 */
export function movePictureTo(workbook: Workbook, sheetIndex: number, id: number, rect: { x: number; y: number; width: number; height: number }, metrics?: SheetMetrics): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = pictureAt(sheet, id);
  const now = drawingAnchorNow(drawing, sheet.structuralEdits);
  return movePicture(workbook, sheetIndex, id, anchorLike(now, metrics || createSheetMetrics(sheet), rect));
}

/**
 * Deletes a picture. Throws RangeError when `id` names none, and for a picture of the file that
 * shares a wrapper of alternatives with another drawing (SheetDrawing.wrapped).
 */
export function removePicture(workbook: Workbook, sheetIndex: number, id: number): Op[] {
  return withCommand(removePicturePlain(workbook, sheetIndex, id), 'removePicture', []);
}

function removePicturePlain(workbook: Workbook, sheetIndex: number, id: number): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = pictureAt(sheet, id);
  const inFile = fileDrawing(drawing);
  if (inFile && inFile.wrapped) throw new RangeError('Cannot delete this picture: it shares its place in the drawing with another object');
  const list = (sheet.drawings as SheetDrawing[]).slice();
  list.splice(id, 1);
  return setDrawings(sheetIndex, list);
}

/** The pictures of a sheet, the one on top last. Pictures inside groups are not listed (they move with their group). */
export function sheetPictures(workbook: Workbook, sheetIndex: number): PictureInfo[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet || sheet.kind !== 'worksheet' || !sheet.drawings) return [];
  const out: PictureInfo[] = [];
  sheet.drawings.forEach((drawing, id) => {
    if (!drawing || !drawing.object || drawing.object.type !== 'picture') return;
    const inFile = fileDrawing(drawing);
    const info: PictureInfo = { id, anchor: drawingAnchorNow(drawing, sheet.structuralEdits), ofFile: inFile !== undefined, removable: !(inFile && inFile.wrapped) };
    if (drawing.object.name !== undefined) info.name = drawing.object.name;
    if (drawing.object.description !== undefined) info.description = drawing.object.description;
    out.push(info);
  });
  return out;
}
