// The frame of a picture as it is turned, sized and cropped by hand, in pixels of the sheet
// (from the corner of cell A1). A picture is drawn in its box (the "drawn box") and turned about
// the middle of it; the anchor of the file gives that box as it is, or, near a quarter turn, the
// box turned by 90 degrees (see quarterTurned()): turnedBox() goes from one to the other.
//
// Handles are those of the drawn box turned with the picture: a point is taken into the frame of
// the box (localPoint()), sized there, and the box is put back where its middle went
// (sizedFrame()). Cropping is done in the same frame: the sides of the box that are cut are those
// of the picture as it is shown (flipped pictures show their right side on the left).
// Nothing here changes a workbook: the commands (ops/pictures.ts) and the view use these.

import type { Rect } from '../layout/grid-geometry';
import { quarterTurned, ratioRect, resizedRect, ROTATE_HANDLE_GAP, ROTATE_HANDLE_GAP_COARSE, type DrawingHandle } from './sheet-drawings';

export { ROTATE_HANDLE_GAP, ROTATE_HANDLE_GAP_COARSE };

/** Shares (0..1) of a picture cut off each side, as DrawingPicture.crop holds them. */
export interface PictureCrop {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const finite = (n: number): number => (typeof n === 'number' && isFinite(n) ? n : 0);

/** The box of the anchor for a drawn box, and the drawn box for the box of an anchor: near a quarter turn the box turned about its middle. */
export function turnedBox(rect: Rect, rotation: number | undefined): Rect {
  if (!quarterTurned(rotation)) return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  return { x: rect.x + (rect.width - rect.height) / 2, y: rect.y + (rect.height - rect.width) / 2, width: rect.height, height: rect.width };
}

/** A vector turned by `degrees` (clockwise on the screen, as the canvas turns). */
export function turnVector(dx: number, dy: number, degrees: number): { x: number; y: number } {
  if (!degrees) return { x: dx, y: dy };
  const angle = (degrees * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  return { x: dx * cos - dy * sin, y: dx * sin + dy * cos };
}

/** A point of the sheet in the frame of a box turned by `rotation` about its middle: where it is on the box as if the box were not turned. */
export function localPoint(box: Rect, rotation: number | undefined, x: number, y: number): { x: number; y: number } {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const turned = turnVector(x - cx, y - cy, -finite(rotation || 0));
  return { x: cx + turned.x, y: cy + turned.y };
}

/** A point of a box turned by `rotation` about its middle, given in the frame of the box, as a point of the sheet. */
export function sheetPoint(box: Rect, rotation: number | undefined, x: number, y: number): { x: number; y: number } {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  const turned = turnVector(x - cx, y - cy, finite(rotation || 0));
  return { x: cx + turned.x, y: cy + turned.y };
}

/**
 * The drawn box that comes of `local` (a box in the frame of `start`, which is turned by
 * `rotation`): its size, with its middle where the middle of `local` is on the sheet.
 */
export function placedFrame(start: Rect, local: Rect, rotation: number | undefined): Rect {
  const middle = sheetPoint(start, rotation, local.x + local.width / 2, local.y + local.height / 2);
  return { x: middle.x - local.width / 2, y: middle.y - local.height / 2, width: local.width, height: local.height };
}

/** A box sized by a handle in its own frame, with nothing that keeps it on the sheet (a turned box can reach beyond its corner). */
function sizedFree(rect: Rect, handle: DrawingHandle, dx: number, dy: number, min: number): Rect {
  let left = rect.x;
  let top = rect.y;
  let right = rect.x + rect.width;
  let bottom = rect.y + rect.height;
  if (handle === 'nw' || handle === 'w' || handle === 'sw') left = Math.min(left + dx, right - min);
  if (handle === 'ne' || handle === 'e' || handle === 'se') right = Math.max(right + dx, left + min);
  if (handle === 'nw' || handle === 'n' || handle === 'ne') top = Math.min(top + dy, bottom - min);
  if (handle === 'sw' || handle === 's' || handle === 'se') bottom = Math.max(bottom + dy, top + min);
  return { x: left, y: top, width: right - left, height: bottom - top };
}

/**
 * The drawn box of a picture turned by `rotation` once a handle of its frame was dragged by
 * (dx, dy) pixels of the sheet, or the whole picture ("move"). With `keepRatio` a corner keeps
 * the proportions of the box. A picture that is not turned is sized as any drawing is
 * (resizedRect(), ratioRect()); a turned one in its own frame, the side across from the handle
 * staying where it is on the sheet.
 */
export function sizedFrame(start: Rect, rotation: number | undefined, handle: DrawingHandle | 'move', dx: number, dy: number, min: number, keepRatio = false): Rect {
  const turn = finite(rotation || 0) % 360;
  if (handle === 'move' || turn === 0) {
    const sized = resizedRect(start, handle, dx, dy, min);
    return keepRatio ? ratioRect(start, sized, handle, min) : sized;
  }
  const local = turnVector(dx, dy, -turn);
  let sized = sizedFree(start, handle, local.x, local.y, min);
  if (keepRatio && handle !== 'n' && handle !== 's' && handle !== 'e' && handle !== 'w' && start.width > 0 && start.height > 0) {
    const scale = Math.max(sized.width / start.width, sized.height / start.height, min / start.width, min / start.height);
    const width = start.width * scale;
    const height = start.height * scale;
    sized = { x: handle === 'nw' || handle === 'sw' ? start.x + start.width - width : start.x, y: handle === 'nw' || handle === 'ne' ? start.y + start.height - height : start.y, width, height };
  }
  return placedFrame(start, sized, turn);
}

/** Where the handle that turns a picture stands on the sheet: `gap` pixels above the middle of the top side of its frame. */
export function rotateHandlePoint(box: Rect, rotation: number | undefined, gap: number = ROTATE_HANDLE_GAP): { x: number; y: number } {
  return sheetPoint(box, rotation, box.x + box.width / 2, box.y - gap);
}

/**
 * The rotation a picture gets when the handle that turns it is dragged to a point: the angle of
 * the point about the middle of the box, 0 straight above it, clockwise, in whole degrees; with
 * `snap` in steps of 15 degrees. Always 0 <= rotation < 360.
 */
export function rotationAt(box: Rect, x: number, y: number, snap: boolean): number {
  const cx = box.x + box.width / 2;
  const cy = box.y + box.height / 2;
  if (!(isFinite(x) && isFinite(y)) || (x === cx && y === cy)) return 0;
  const degrees = (Math.atan2(y - cy, x - cx) * 180) / Math.PI + 90;
  const step = snap ? 15 : 1;
  const rounded = Math.round(degrees / step) * step;
  return ((rounded % 360) + 360) % 360;
}

/** The shares of a picture cut off the sides of its box as it is shown: those of the picture, mirrored where it is flipped. */
function boxSides(crop: PictureCrop | undefined, flipH: boolean, flipV: boolean): PictureCrop {
  const c = crop || { left: 0, top: 0, right: 0, bottom: 0 };
  return { left: flipH ? c.right : c.left, right: flipH ? c.left : c.right, top: flipV ? c.bottom : c.top, bottom: flipV ? c.top : c.bottom };
}

/**
 * The box the whole picture would take up (nothing cut off) in the frame of its drawn box `box`,
 * whose picture is cut by `crop`.
 */
export function uncroppedBox(box: Rect, crop: PictureCrop | undefined, flipH = false, flipV = false): Rect {
  const sides = boxSides(crop, flipH, flipV);
  const width = box.width / Math.max(1 - sides.left - sides.right, 0.001);
  const height = box.height / Math.max(1 - sides.top - sides.bottom, 0.001);
  return { x: box.x - sides.left * width, y: box.y - sides.top * height, width, height };
}

/**
 * The drawn box of a picture once it is cut by `next` in place of `crop`: the picture keeps its
 * scale and what is still shown of it stays where it is on the sheet; what is cut off or given
 * back moves the sides of the box.
 */
export function croppedBox(box: Rect, rotation: number | undefined, crop: PictureCrop | undefined, next: PictureCrop | undefined, flipH = false, flipV = false): Rect {
  const whole = uncroppedBox(box, crop, flipH, flipV);
  const sides = boxSides(next, flipH, flipV);
  const local = { x: whole.x + sides.left * whole.width, y: whole.y + sides.top * whole.height, width: whole.width * (1 - sides.left - sides.right), height: whole.height * (1 - sides.top - sides.bottom) };
  return placedFrame(box, local, rotation);
}

const clamp = (value: number, low: number, high: number): number => Math.max(low, Math.min(high, value));

/**
 * The crop of a picture once a handle of its frame was dragged by (dx, dy) pixels of the sheet
 * in cropping: the side the handle is on (both sides of a corner) cuts off more or less of the
 * picture, never less than nothing and never so much that less than `min` pixels are left.
 * Undefined for a crop that cuts nothing.
 */
export function draggedCrop(box: Rect, rotation: number | undefined, crop: PictureCrop | undefined, handle: DrawingHandle, dx: number, dy: number, min: number, flipH = false, flipV = false): PictureCrop | undefined {
  const whole = uncroppedBox(box, crop, flipH, flipV);
  const sides = boxSides(crop, flipH, flipV);
  const local = turnVector(dx, dy, -finite(rotation || 0));
  const minX = Math.min(min / Math.max(whole.width, 1e-9), 1);
  const minY = Math.min(min / Math.max(whole.height, 1e-9), 1);
  if (handle === 'nw' || handle === 'w' || handle === 'sw') sides.left = clamp(sides.left + local.x / whole.width, 0, Math.max(0, 1 - sides.right - minX));
  if (handle === 'ne' || handle === 'e' || handle === 'se') sides.right = clamp(sides.right - local.x / whole.width, 0, Math.max(0, 1 - sides.left - minX));
  if (handle === 'nw' || handle === 'n' || handle === 'ne') sides.top = clamp(sides.top + local.y / whole.height, 0, Math.max(0, 1 - sides.bottom - minY));
  if (handle === 'sw' || handle === 's' || handle === 'se') sides.bottom = clamp(sides.bottom - local.y / whole.height, 0, Math.max(0, 1 - sides.top - minY));
  return normalCrop(boxSides(sides, flipH, flipV));
}

/** A crop as the file holds it: each share in 1/100,000 (see `a:srcRect`); undefined for one that cuts nothing. */
export function normalCrop(crop: PictureCrop | undefined): PictureCrop | undefined {
  if (!crop) return undefined;
  const share = (value: number): number => Math.round(clamp(finite(value), 0, 1) * 100000) / 100000;
  const out = { left: share(crop.left), top: share(crop.top), right: share(crop.right), bottom: share(crop.bottom) };
  return out.left === 0 && out.top === 0 && out.right === 0 && out.bottom === 0 ? undefined : out;
}
