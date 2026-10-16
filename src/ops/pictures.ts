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
// - A picture is turned (DrawingObject.rotation), cropped (DrawingPicture.crop)
//   and given another alternative text; of a picture of the file the writer
//   changes only those (`a:xfrm rot`, `a:srcRect`, `descr`) besides its anchor.
//   Near a quarter turn the anchor of a picture is the box it takes up turned
//   (see turnedBox()): a turn across that line changes the anchor too.
// - A picture is copied with its bytes, its size, its turn, its crop and its
//   alternative text (copyPicture()), and pasted (pastePicture()) on any sheet
//   of any workbook: what a copy holds names nothing of the workbook it came from.
//
// What kind of picture bytes are is read from the bytes, never from a name or
// from what a caller says they are; what is not a PNG, JPEG, GIF or BMP
// picture with a readable size is refused, as is one too large to be shown.

import type { SheetMetrics } from '../layout/sheet-metrics';
import { createSheetMetrics } from '../layout/sheet-metrics';
import { isAddedPicture, MAX_DRAWING_TEXT, newPictureKey, pictureSource, PICTURE_EXTENSIONS, textUpTo } from '../model/sheet/pictures';
import { MAX_COLS, MAX_ROWS, type CellRange, type DrawingAnchor, type DrawingMarker, type DrawingObject, type DrawingPicture, type PictureData, type Sheet, type SheetDrawing, type Workbook } from '../model/sheet/types';
import { anchorRect, drawingAnchorNow } from '../render/sheet-drawings';
import { croppedBox, normalCrop, turnedBox, type PictureCrop } from '../render/picture-frame';
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
/** Longest name and alternative text of a picture: the alternative text a reader keeps of a file (see MAX_DRAWING_TEXT). */
export const MAX_PICTURE_TEXT = MAX_DRAWING_TEXT;
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
  /** Degrees it is turned, clockwise; absent when it is not. */
  rotation?: number;
  /** Shares cut off its sides; absent when nothing is. */
  crop?: PictureCrop;
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
export function fileDrawing(drawing: SheetDrawing): SheetDrawing | undefined {
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

/** What the pictures that were added to a workbook hold, each set of bytes counted once; `also`, when the workbook holds it already, is taken off (it would add nothing). */
function addedBytes(workbook: Workbook, also?: Uint8Array): number {
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
  return also !== undefined && seen.has(also) ? total - also.length : total;
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
 * kept on the sheet and at least `min` pixels each way (MIN_PICTURE_SIZE; a line may be flat: 0).
 */
export function anchorLike(like: DrawingAnchor, metrics: SheetMetrics, rect: { x: number; y: number; width: number; height: number }, min: number = MIN_PICTURE_SIZE): DrawingAnchor {
  const finite = (n: number, fallback: number): number => (typeof n === 'number' && isFinite(n) ? n : fallback);
  const limitX = metrics.cols.total;
  const limitY = metrics.rows.total;
  const least = typeof min === 'number' && min >= 0 && min <= MIN_PICTURE_SIZE ? min : MIN_PICTURE_SIZE;
  const width = Math.max(Math.min(finite(rect.width, least), limitX), least);
  const height = Math.max(Math.min(finite(rect.height, least), limitY), least);
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
  return movedDrawing(sheetIndex, sheet, id, drawing, anchor, 'a picture');
}

/**
 * The operations that put a drawing of a sheet (a picture, a shape) at another anchor: one of the
 * kind it has, or the cells it covers. What movePicture() does once the drawing is found; `what`
 * names it in a message. `min` is the least size the cells it covers are made into (see anchorLike()).
 */
export function movedDrawing(sheetIndex: number, sheet: Sheet, id: number, drawing: SheetDrawing, anchor: DrawingAnchor | CellRange, what: string, min: number = MIN_PICTURE_SIZE): Op[] {
  const now = drawingAnchorNow(drawing, sheet.structuralEdits);
  let wanted: DrawingAnchor;
  if (isRange(anchor)) {
    const cells = checkedAnchor(anchor, what);
    const metrics = createSheetMetrics(sheet);
    wanted = anchorLike(now, metrics, anchorRect(cells, metrics), min);
  } else {
    wanted = checkedAnchor(anchor, what);
    if (wanted.kind !== now.kind) throw new RangeError(`${what.charAt(0).toUpperCase()}${what.slice(1)} keeps the kind of its anchor`);
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
  return removedDrawing(sheetIndex, sheet, id, drawing, 'picture');
}

/** The operations that take a drawing off a sheet; one of the file that shares a wrapper of alternatives with another drawing is refused. `what` names it in the message. */
export function removedDrawing(sheetIndex: number, sheet: Sheet, id: number, drawing: SheetDrawing, what: string): Op[] {
  const inFile = fileDrawing(drawing);
  if (inFile && inFile.wrapped) throw new RangeError(`Cannot delete this ${what}: it shares its place in the drawing with another object`);
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
    if (drawing.object.rotation) info.rotation = drawing.object.rotation;
    if (drawing.object.crop) info.crop = { ...drawing.object.crop };
    out.push(info);
  });
  return out;
}

// ------------------------------------------------------------------ turn, crop, text

/**
 * Why a picture cannot be turned, cropped or given another alternative text, as a message;
 * undefined when it can. Besides what keeps every picture of the sheet as it is
 * (picturesProblem()): a picture the file holds in a wrapper of alternatives (inside its anchor, or
 * its anchor inside one: SheetDrawing.alternative), whose other form would be left as it was. It
 * is moved and deleted all the same.
 */
export function pictureEditProblem(workbook: Workbook, sheetIndex: number, id: number): string | undefined {
  const problem = picturesProblem(workbook, sheetIndex);
  if (problem) return problem;
  const sheet = workbook.sheets[sheetIndex];
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'picture') return `No picture at index ${id}`;
  const inFile = fileDrawing(drawing);
  if (inFile && inFile.alternative) return 'Cannot change this picture: the file holds it in another form as well, which would be left as it was';
  return undefined;
}

/** The picture of a sheet a command changes, with where it stands now and its drawn box (pixels; see turnedBox()). */
function pictureNow(workbook: Workbook, sheetIndex: number, id: number) {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing, object } = pictureAt(sheet, id);
  const problem = pictureEditProblem(workbook, sheetIndex, id);
  if (problem) throw new RangeError(problem);
  const now = drawingAnchorNow(drawing, sheet.structuralEdits);
  const metrics = createSheetMetrics(sheet);
  return { sheet, drawing, object, now, metrics, drawn: turnedBox(anchorRect(now, metrics), object.rotation) };
}

/** The list of drawings with the drawing at `id` holding another object and, when given, standing at another anchor. */
export function changedDrawing(sheetIndex: number, sheet: Sheet, id: number, drawing: SheetDrawing, object: DrawingObject, now: DrawingAnchor, anchor: DrawingAnchor | undefined): Op[] {
  const moved = anchor !== undefined && JSON.stringify(anchor) !== JSON.stringify(now);
  // A picture that moves is put where it is from the structural edits made so far on (as movePicture() does); one that stays keeps its place as it is.
  const next: SheetDrawing = moved ? { anchor: anchor as DrawingAnchor, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 } : { ...drawing, object };
  const from = fileDrawing(drawing);
  if (from) next.from = from;
  else delete next.from;
  const list = (sheet.drawings as SheetDrawing[]).slice();
  list[id] = next;
  return setDrawings(sheetIndex, list);
}

/** An angle in degrees as a picture holds it: 0 <= angle < 360, in the steps of the file (1/60,000 of a degree). Throws RangeError for what is not a number. */
export function pictureRotation(degrees: unknown): number {
  if (typeof degrees !== 'number' || !isFinite(degrees) || Math.abs(degrees) > 1e7) throw new RangeError('The rotation of a picture is a number of degrees');
  const steps = Math.round((((degrees % 360) + 360) % 360) * 60000);
  return steps >= 360 * 60000 ? 0 : steps / 60000;
}

/** Whether a box is given turned by a quarter at one rotation and not at the other (see turnedBox()). */
const crossesQuarter = (a: number, b: number): boolean => turnedBox({ x: 0, y: 0, width: 2, height: 1 }, a).width !== turnedBox({ x: 0, y: 0, width: 2, height: 1 }, b).width;

/**
 * Turns a picture to `degrees` (clockwise; 90 is a quarter turn to the right) about the middle
 * of its box. Its box keeps its size and its middle; near a quarter turn its anchor is the box
 * turned (as the file gives such a picture), so a turn across 45 degrees changes its anchor.
 * Returns no operation when the picture is turned so already. Throws RangeError for an angle that
 * is not a number.
 */
export function rotatePicture(workbook: Workbook, sheetIndex: number, id: number, degrees: number): Op[] {
  return withCommand(rotatePicturePlain(workbook, sheetIndex, id, degrees), 'rotatePicture', []);
}

function rotatePicturePlain(workbook: Workbook, sheetIndex: number, id: number, degrees: number): Op[] {
  const rotation = pictureRotation(degrees);
  const { sheet, drawing, object, now, metrics } = pictureNow(workbook, sheetIndex, id);
  return turnedDrawing(sheetIndex, sheet, id, drawing, object, now, metrics, rotation);
}

/**
 * The operations that turn a drawing of a sheet (a picture, a shape) to `rotation` degrees (as
 * pictureRotation() gives them) about the middle of its box: see rotatePicture(). `min` is the
 * least size its anchor is kept at (see anchorLike()).
 */
export function turnedDrawing(sheetIndex: number, sheet: Sheet, id: number, drawing: SheetDrawing, object: DrawingObject, now: DrawingAnchor, metrics: SheetMetrics, rotation: number, min: number = MIN_PICTURE_SIZE): Op[] {
  const before = object.rotation || 0;
  if (before === rotation) return [];
  const drawn = turnedBox(anchorRect(now, metrics), object.rotation);
  const next: DrawingObject = { ...object };
  if (rotation === 0) delete next.rotation;
  else next.rotation = rotation;
  // Unless a quarter turn is crossed the anchor is left as it is.
  if (!crossesQuarter(before, rotation)) return changedDrawing(sheetIndex, sheet, id, drawing, next, now, undefined);
  const edits = sheet.structuralEdits ? sheet.structuralEdits.length : 0;
  // Turned back across the line, where the last turn put it and with no row or column moved since: the anchor it had.
  const back = drawing.turned;
  if (back && back.edits === edits && JSON.stringify(back.to) === JSON.stringify(now)) {
    const ops = changedDrawing(sheetIndex, sheet, id, drawing, next, now, back.from);
    const list = (ops[0] as Extract<Op, { type: 'setDrawings' }>).drawings as SheetDrawing[];
    delete list[id].turned;
    return ops;
  }
  // The turned box is kept on the sheet, which may move it; the anchor it had is kept for the turn back.
  const anchor = anchorLike(now, metrics, turnedBox(drawn, rotation), min);
  const ops = changedDrawing(sheetIndex, sheet, id, drawing, next, now, anchor);
  const list = (ops[0] as Extract<Op, { type: 'setDrawings' }>).drawings as SheetDrawing[];
  list[id].turned = { from: now, to: drawingAnchorNow(list[id], sheet.structuralEdits), edits };
  return ops;
}

/** A crop as a command takes it: shares from 0 to 1 that leave something of the picture each way, in the steps of the file; undefined for none. Throws RangeError otherwise. */
export function checkedCrop(crop: unknown): PictureCrop | undefined {
  if (crop === undefined || crop === null) return undefined;
  const given = crop as PictureCrop;
  if (typeof given !== 'object' || Array.isArray(given)) throw new RangeError('The crop of a picture gives the shares cut off its sides');
  for (const side of ['left', 'top', 'right', 'bottom'] as const) {
    const value = given[side] === undefined ? 0 : given[side];
    if (typeof value !== 'number' || !isFinite(value) || value < 0 || value >= 1) throw new RangeError('The crop of a picture gives the shares cut off its sides, each from 0 to 1');
  }
  const made = normalCrop({ left: given.left || 0, top: given.top || 0, right: given.right || 0, bottom: given.bottom || 0 });
  if (made && (made.left + made.right > 0.999 || made.top + made.bottom > 0.999)) throw new RangeError('A crop cannot leave nothing of the picture');
  return made;
}

/**
 * Crops a picture: `crop` gives the shares (0..1) of the picture cut off each side, undefined
 * (or nothing each way) for none. The picture keeps its scale, and what is still shown of it
 * stays where it is: the sides of its box move (see croppedBox()). Shares are kept in the steps of
 * the file (1/100,000). Returns no operation when the picture is cropped so already. Throws
 * RangeError for a share that is not a number from 0 to 1, and for a crop that leaves nothing.
 */
export function cropPicture(workbook: Workbook, sheetIndex: number, id: number, crop: PictureCrop | undefined): Op[] {
  return withCommand(cropPicturePlain(workbook, sheetIndex, id, crop), 'cropPicture', []);
}

function cropPicturePlain(workbook: Workbook, sheetIndex: number, id: number, crop: PictureCrop | undefined): Op[] {
  const wanted = checkedCrop(crop);
  const { sheet, drawing, object, now, metrics, drawn } = pictureNow(workbook, sheetIndex, id);
  if (JSON.stringify(object.crop || null) === JSON.stringify(wanted || null)) return [];
  const next: DrawingPicture = { ...object };
  if (wanted) next.crop = wanted;
  else delete next.crop;
  const box = croppedBox(drawn, object.rotation, object.crop, wanted, object.flipH === true, object.flipV === true);
  // As the handles do (see draggedCrop()): no less than the least size of a picture is left of it each way.
  if (box.width < MIN_PICTURE_SIZE - 1e-6 || box.height < MIN_PICTURE_SIZE - 1e-6) throw new RangeError(`A crop cannot leave less than ${MIN_PICTURE_SIZE} pixels of the picture`);
  return changedDrawing(sheetIndex, sheet, id, drawing, next, now, anchorLike(now, metrics, turnedBox(box, object.rotation)));
}

/**
 * Gives a picture another alternative text (`descr` of the file): at most MAX_PICTURE_TEXT
 * characters, line breaks kept; empty or undefined takes it away. Returns no operation when the
 * picture has that text already. Throws RangeError for a text too long, or with a character a
 * file cannot hold.
 */
export function describePicture(workbook: Workbook, sheetIndex: number, id: number, text: string | undefined): Op[] {
  return withCommand(describePicturePlain(workbook, sheetIndex, id, text), 'describePicture', []);
}

function describePicturePlain(workbook: Workbook, sheetIndex: number, id: number, text: string | undefined): Op[] {
  const description = checkedText(text, 'alternative text');
  const { sheet, drawing, object, now } = pictureNow(workbook, sheetIndex, id);
  if (object.description === description) return [];
  const next: DrawingPicture = { ...object };
  if (description === undefined) delete next.description;
  else next.description = description;
  return changedDrawing(sheetIndex, sheet, id, drawing, next, now, undefined);
}

// ------------------------------------------------------------------ copies

/** How far a copy is put from a picture that stands where it would stand: so many pixels to the right and down. */
export const PICTURE_COPY_OFFSET = 24;

/**
 * A picture taken off a sheet to be put down again: what copyPicture() and cutPicture() give and
 * pastePicture() takes. It holds the bytes of the picture and nothing of the workbook it came
 * from: it is put down on any sheet of any workbook. Not to be altered.
 */
export interface PictureCopy {
  readonly bytes: Uint8Array;
  /** What the bytes are: "image/png", "image/jpeg", "image/gif" or "image/bmp" (read from them). */
  readonly type: string;
  /** Its box as its anchor gives it, pixels from the corner of cell A1: where it stood, and its size. */
  readonly rect: { x: number; y: number; width: number; height: number };
  /** The kind of anchor it had: the copy gets the same. */
  readonly anchor: { kind: DrawingAnchor['kind']; editAs?: DrawingAnchor['editAs'] };
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly crop?: PictureCrop;
  readonly description?: string;
}

/** What pastePicture() gives. */
export interface PastedPicture {
  /** The operations that add the picture: one undo step. */
  ops: Op[];
  /** The id the picture has on its sheet once they are applied. */
  id: number;
}

/** Copies this module made: their bytes were checked, and are no caller's. */
const MADE = new WeakSet<object>();

/** The parts a copy can be made of, by the extension of their names; the bytes say at last. */
const COPIED_EXTENSIONS = /\.(png|jpe?g|jfif|gif|bmp|dib)$/i;

/**
 * Why a picture cannot be copied (or cut), as a message; undefined when it can: no picture there,
 * one whose bytes are kept outside the file or are missing, one of a kind no picture is added as
 * (a metafile, SVG, TIFF, WebP: told by the name of its part; the bytes are checked once read).
 */
export function pictureCopyProblem(workbook: Workbook, sheetIndex: number, id: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot copy the pictures of a ${sheet.kind}`;
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'picture') return `No picture at index ${id}`;
  const object = drawing.object;
  if (isAddedPicture(object)) return undefined;
  if (object.external || object.media === undefined) return 'Cannot copy this picture: its bytes are not in the file';
  if (!COPIED_EXTENSIONS.test(object.media)) return 'Cannot copy this picture: only PNG, JPEG, GIF and BMP pictures are copied';
  return undefined;
}

/**
 * Takes a copy of a picture; nothing of the workbook is changed. What the copy holds is taken
 * now (its place, size, turn, crop and alternative text); its bytes are those of the model (a
 * picture that was added) or are read from the file. The promise is refused (RangeError) when the
 * bytes are not a picture that can be added (see checkPicture()). Throws RangeError at once for a
 * picture that cannot be copied: pictureCopyProblem() says why.
 */
export function copyPicture(workbook: Workbook, sheetIndex: number, id: number): Promise<PictureCopy> {
  const problem = pictureCopyProblem(workbook, sheetIndex, id);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[sheetIndex];
  const { drawing, object } = pictureAt(sheet, id);
  const anchor = drawingAnchorNow(drawing, sheet.structuralEdits);
  const box = anchorRect(anchor, createSheetMetrics(sheet));
  const held: { -readonly [key in keyof PictureCopy]?: PictureCopy[key] } = {
    rect: { x: box.x, y: box.y, width: box.width, height: box.height },
    anchor: anchor.kind === 'twoCell' && anchor.editAs !== undefined ? { kind: anchor.kind, editAs: anchor.editAs } : { kind: anchor.kind },
  };
  if (object.rotation) held.rotation = object.rotation;
  if (object.flipH) held.flipH = true;
  if (object.flipV) held.flipV = true;
  const crop = normalCrop(object.crop);
  if (crop) held.crop = crop;
  if (object.description !== undefined && object.description !== '') held.description = textUpTo(object.description, MAX_PICTURE_TEXT);
  // The bytes of a picture that was added are those of the model (never altered); those of the file are read, and kept as a copy of their own.
  const read: Promise<Uint8Array> = isAddedPicture(object)
    ? Promise.resolve(object.data.bytes)
    : pictureSource(workbook)
        .readPart(object.media as string)
        .then((bytes) => (bytes instanceof Uint8Array ? bytes.slice() : bytes));
  return read.then((bytes) => {
    const facts = checkPicture(bytes);
    if (typeof facts === 'string') throw new RangeError(`Cannot copy this picture: ${facts.charAt(0).toLowerCase()}${facts.slice(1)}`);
    held.bytes = bytes;
    held.type = facts.type;
    MADE.add(held);
    return held as PictureCopy;
  });
}

/**
 * Takes a copy of a picture and, once the copy is made, the operations that delete it (one undo
 * step; the paste is another). Nothing is deleted before the copy is there: the promise is refused
 * for a picture whose copy cannot be made, and for one that was changed or deleted while its bytes
 * were read. Throws RangeError at once for a picture that cannot be copied or cannot be deleted.
 */
export function cutPicture(workbook: Workbook, sheetIndex: number, id: number): Promise<{ copy: PictureCopy; ops: Op[] }> {
  // What refuses the deletion is said now, before anything is read.
  removePicture(workbook, sheetIndex, id);
  const copy = copyPicture(workbook, sheetIndex, id);
  const sheet = workbook.sheets[sheetIndex];
  const drawing = (sheet.drawings || [])[id];
  return copy.then((made) => {
    const home = workbook.sheets.indexOf(sheet);
    const at = home >= 0 ? (sheet.drawings || []).indexOf(drawing) : -1;
    if (at < 0) throw new RangeError('The picture was changed or deleted before it was cut');
    return { copy: made, ops: removePicture(workbook, home, at) };
  });
}

/**
 * Puts a copy of a picture on a sheet, above its other drawings, with the size, kind of anchor,
 * turn, flips, crop and alternative text of the picture it was taken from, and a name of its own.
 * With `at` its box has its top left corner at the corner of that cell; without, it stands where
 * the picture stood, moved right and down (PICTURE_COPY_OFFSET) for as long as a picture stands
 * there already. A copy that does not come from copyPicture() is checked as any picture that is
 * added. Throws RangeError for what is not a copy of a picture, and for a sheet or a workbook that
 * takes no more pictures.
 */
export function pastePicture(workbook: Workbook, sheetIndex: number, copy: PictureCopy, at?: { row: number; col: number }): PastedPicture {
  const sheet = placedSheetOf(workbook, sheetIndex);
  if (!copy || typeof copy !== 'object' || !(copy.bytes instanceof Uint8Array)) throw new RangeError('What is pasted is not a picture that was copied');
  const own = MADE.has(copy);
  const facts = own ? { type: copy.type } : checkPicture(copy.bytes);
  if (typeof facts === 'string') throw new RangeError(facts);
  // The bytes of a copy made elsewhere are taken as a copy of their own: its array may change afterwards.
  const bytes = own ? copy.bytes : copy.bytes.slice();
  const drawings = sheet.drawings || [];
  if (drawings.length >= MAX_SHEET_DRAWINGS) throw new RangeError('The sheet has too many drawings');
  if (addedBytes(workbook, bytes) + bytes.length > MAX_WORKBOOK_PICTURE_BYTES) throw new RangeError(`The pictures added to a workbook cannot hold more than ${MAX_WORKBOOK_PICTURE_BYTES / 1024 / 1024} MB`);
  const number = (n: unknown, fallback: number): number => (typeof n === 'number' && isFinite(n) ? n : fallback);
  const rect = copy.rect && typeof copy.rect === 'object' ? copy.rect : { x: 0, y: 0, width: DEFAULT_PICTURE_WIDTH, height: DEFAULT_PICTURE_HEIGHT };
  const width = Math.max(number(rect.width, DEFAULT_PICTURE_WIDTH), MIN_PICTURE_SIZE);
  const height = Math.max(number(rect.height, DEFAULT_PICTURE_HEIGHT), MIN_PICTURE_SIZE);
  const metrics = createSheetMetrics(sheet);
  let x = Math.max(0, number(rect.x, 0));
  let y = Math.max(0, number(rect.y, 0));
  if (at !== undefined) {
    const whole = (n: unknown, limit: number) => typeof n === 'number' && n >= 0 && n < limit && Math.floor(n) === n;
    if (!at || !whole(at.row, MAX_ROWS) || !whole(at.col, MAX_COLS)) throw new RangeError('The place of a picture is a cell of the sheet');
    x = metrics.cols.offset(at.col);
    y = metrics.rows.offset(at.row);
  } else {
    const taken: { x: number; y: number }[] = [];
    for (const drawing of drawings) {
      if (!drawing || !drawing.object || drawing.object.type !== 'picture') continue;
      const box = anchorRect(drawingAnchorNow(drawing, sheet.structuralEdits), metrics);
      taken.push({ x: box.x, y: box.y });
    }
    for (let n = 0; n < 64 && taken.some((box) => Math.abs(box.x - x) < 1 && Math.abs(box.y - y) < 1); n++) {
      x += PICTURE_COPY_OFFSET;
      y += PICTURE_COPY_OFFSET;
    }
  }
  const kind = copy.anchor && (copy.anchor.kind === 'oneCell' || copy.anchor.kind === 'absolute') ? copy.anchor.kind : 'twoCell';
  const like: DrawingAnchor = { kind };
  if (kind === 'twoCell' && copy.anchor && (copy.anchor.editAs === 'oneCell' || copy.anchor.editAs === 'absolute')) like.editAs = copy.anchor.editAs;
  const anchor = anchorLike(like, metrics, { x, y, width, height });
  const object: DrawingPicture = { type: 'picture', media: newPictureKey(), data: { bytes, type: facts.type }, name: nextName(sheet) };
  const description = checkedText(copy.description, 'alternative text');
  if (description !== undefined) object.description = description;
  if (copy.rotation !== undefined) {
    const rotation = pictureRotation(copy.rotation);
    if (rotation !== 0) object.rotation = rotation;
  }
  if (copy.flipH === true) object.flipH = true;
  if (copy.flipV === true) object.flipV = true;
  const crop = checkedCrop(copy.crop);
  if (crop) object.crop = crop;
  const added: SheetDrawing = { anchor, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  // The bytes are not what a collaborator would be sent: the command is named, the list is not shared (see collab/wire.ts).
  return { ops: withCommand(setDrawings(sheetIndex, drawings.concat([added])), 'pastePicture', []), id: drawings.length };
}
