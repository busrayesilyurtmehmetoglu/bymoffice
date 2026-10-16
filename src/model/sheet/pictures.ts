// Pictures that were added to a workbook since it was read. Their bytes are in
// the model (DrawingPicture.data), not in a part of the file, and they are
// known by a key of their own (DrawingPicture.media) that names no part.
// Whoever reads the bytes of pictures (the view, printing, a workbook built
// anew) reads them through pictureSource(), which knows both kinds.

import type { DrawingObject, PictureData, SheetDrawing, Workbook, WorkbookMedia } from './types';

/** The start of the key of a picture that is in no file: no part name starts so (part names start with "/"). */
export const ADDED_PICTURE_KEY = 'added:';

/** File name extension of each kind of picture that can be added, by its media type. */
export const PICTURE_EXTENSIONS: { [type: string]: string } = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp' };

/**
 * The longest alternative text of a drawing, in UTF-16 code units: what a reader keeps of the file
 * (`descr`) and what a command gives a picture. One limit, so that a text read from a file is
 * always one a command takes again.
 */
export const MAX_DRAWING_TEXT = 2000;

/** A text cut to at most `limit` code units, never between the two halves of a character. */
export function textUpTo(text: string, limit: number): string {
  if (text.length <= limit) return text;
  const code = text.charCodeAt(limit - 1);
  return text.slice(0, code >= 0xd800 && code <= 0xdbff ? limit - 1 : limit);
}

let lastKey = 0;

/** A key for a picture that is added: one no other picture of any workbook has. */
export function newPictureKey(): string {
  lastKey++;
  return `${ADDED_PICTURE_KEY}${lastKey.toString(36)}`;
}

/** Whether an object is a picture that was added since the workbook was read. */
export function isAddedPicture(object: DrawingObject | undefined): object is Extract<DrawingObject, { type: 'picture' }> & { data: PictureData; media: string } {
  return !!object && object.type === 'picture' && object.data !== undefined && object.data.bytes instanceof Uint8Array && typeof object.media === 'string' && object.media.indexOf(ADDED_PICTURE_KEY) === 0;
}

/**
 * The added pictures of a workbook by their keys, worked out again only when a list of drawings
 * was replaced (an edit puts another list in place: see "setDrawings") or a sheet came or went.
 */
interface AddedIndex {
  lists: (SheetDrawing[] | undefined)[];
  bytes: Map<string, Uint8Array>;
}

function addedIndex(workbook: Workbook, held: AddedIndex | undefined): AddedIndex {
  const sheets = workbook.sheets;
  if (held && held.lists.length === sheets.length && sheets.every((sheet, i) => sheet.drawings === held.lists[i])) return held;
  const bytes = new Map<string, Uint8Array>();
  for (const sheet of sheets) {
    for (const drawing of sheet.drawings || []) {
      const object = drawing ? drawing.object : undefined;
      if (isAddedPicture(object) && !bytes.has(object.media)) bytes.set(object.media, object.data.bytes);
    }
  }
  return { lists: sheets.map((sheet) => sheet.drawings), bytes };
}

/**
 * Where the bytes of the pictures of a workbook come from: those of the file from the file (as
 * Workbook.media reads them), those that were added from the model (only while the model holds
 * them: a picture taken away is not read).
 */
export function pictureSource(workbook: Workbook): WorkbookMedia {
  let index: AddedIndex | undefined;
  return {
    readPart(key: string): Promise<Uint8Array> {
      if (typeof key === 'string' && key.indexOf(ADDED_PICTURE_KEY) === 0) {
        index = addedIndex(workbook, index);
        const bytes = index.bytes.get(key);
        return bytes ? Promise.resolve(bytes) : Promise.reject(new Error('Not a picture of this workbook'));
      }
      const media = workbook.media;
      return media && typeof media.readPart === 'function' ? media.readPart(key) : Promise.reject(new Error(`Not a picture of this workbook: ${key}`));
    },
  };
}
