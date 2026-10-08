// Notes on cells (the comments older applications call comments, and Excel
// 365 calls notes) as the editor changes them: a note added, its text edited,
// a note deleted. Like the other commands these functions only build
// operations; the caller applies them through a History so that each is one
// undo step.
//
// Sheet.comments holds the comments of the file where the file has them, and
// the rows and columns moved since (Sheet.structuralEdits) are applied where a
// comment is looked at. A note added here is at the cell it was put on, which
// is a cell of the sheet as it is now: it remembers how many structural edits
// had been made (SheetComment.since) and is moved by the later ones only. A
// comment of the file whose text is edited stays where it is in the model and
// gets another text. commentCellNow() is the one rule for where either is.
//
// Threaded comments (conversations) are not edited: a cell that has one
// refuses every change here, and the writer leaves its parts alone.

import { MAX_COLS, MAX_ROWS, cellName, type Sheet, type SheetComment, type Workbook } from '../model/sheet/types';
import { commentCellNow } from '../render/sheet-drawings';
import type { Op } from './ops';

/** Longest text of a note: the longest text a cell holds. */
export const MAX_NOTE_TEXT = 32767;
/** Who a new note is by when nobody is named. */
export const DEFAULT_NOTE_AUTHOR = 'Author';
/** Longest name of an author that is kept. */
const MAX_AUTHOR = 255;
/** Comments one sheet holds: as many as are read from a file. */
export const MAX_SHEET_NOTES = 100000;

export interface NoteOptions {
  /** Who a new note is by. Default "Author". An existing note keeps its author when its text is edited. */
  author?: string;
}

/** The merged range a cell is in. */
function mergeOf(sheet: Sheet, row: number, col: number): { top: number; left: number; bottom: number; right: number } | undefined {
  const merges = sheet.merges;
  for (let i = 0; i < merges.length; i++) {
    const merge = merges[i];
    if (row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right) return merge;
  }
  return undefined;
}

/** The cell a new note of `row`, `col` is put on: the top-left cell of the merged range the cell is in, else the cell. */
function noteCell(sheet: Sheet, row: number, col: number): { row: number; col: number } {
  const merge = mergeOf(sheet, row, col);
  return merge ? { row: merge.top, col: merge.left } : { row, col };
}

/** Where a comment of the model is now; undefined when its cell was deleted or is none. */
export function cellNow(sheet: Sheet, comment: SheetComment): { row: number; col: number } | undefined {
  if (!comment || !(comment.row >= 0 && comment.row < MAX_ROWS && comment.col >= 0 && comment.col < MAX_COLS)) return undefined;
  const edits = sheet.structuralEdits;
  // Without edits a comment is where the model says; nothing is worked out for it.
  return edits && edits.length > 0 ? commentCellNow(comment, edits) : comment;
}

/** Whether the sheet has a comment on a cell that is there (not one that went with its row or column). */
export function hasLiveComments(sheet: Sheet): boolean {
  const comments = sheet.comments;
  if (!comments) return false;
  for (let i = 0; i < comments.length; i++) if (cellNow(sheet, comments[i])) return true;
  return false;
}

/**
 * The comment that is on a cell now, with its index in Sheet.comments; undefined when the
 * cell has none. As everywhere, the first comment of a cell is the one that counts. For a
 * merged cell this is the comment of its top-left cell, or else the first one on any cell
 * inside it (cells merged over a note keep the note): a merged cell has one note.
 */
export function commentAtCell(sheet: Sheet, row: number, col: number): { comment: SheetComment; index: number } | undefined {
  const comments = sheet.comments;
  if (!comments || comments.length === 0) return undefined;
  const merge = mergeOf(sheet, row, col);
  const top = merge ? merge.top : row;
  const left = merge ? merge.left : col;
  let inside: { comment: SheetComment; index: number } | undefined;
  for (let i = 0; i < comments.length; i++) {
    const cell = cellNow(sheet, comments[i]);
    if (!cell) continue;
    if (cell.row === top && cell.col === left) return { comment: comments[i], index: i };
    if (merge && !inside && cell.row >= merge.top && cell.row <= merge.bottom && cell.col >= merge.left && cell.col <= merge.right) inside = { comment: comments[i], index: i };
  }
  return inside;
}

/** Why no note of a sheet can be changed, by what the reader found (Sheet.notesBlocked). */
const BLOCKED: Record<string, string> = {
  drawing: 'its legacy drawing cannot be read',
  shared: 'its comment parts are shared with another sheet',
  parts: 'it has more than one comments part',
};

function checkTarget(workbook: Workbook, sheetIndex: number, row: number, col: number): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (!(row >= 0 && row < MAX_ROWS && col >= 0 && col < MAX_COLS) || Math.floor(row) !== row || Math.floor(col) !== col) {
    throw new RangeError(`Cell (${row}, ${col}) is outside the sheet`);
  }
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change comments in a ${sheet.kind}`);
  return sheet;
}

/**
 * Why the note of a cell cannot be added, edited or deleted, as a message; undefined when it
 * can. Throws RangeError for a cell or a sheet that does not exist.
 */
export function noteProblem(workbook: Workbook, sheetIndex: number, row: number, col: number): string | undefined {
  const sheet = checkTarget(workbook, sheetIndex, row, col);
  if (sheet.notesBlocked !== undefined && Object.prototype.hasOwnProperty.call(BLOCKED, sheet.notesBlocked)) {
    return `Cannot change the comments of this sheet: ${BLOCKED[sheet.notesBlocked]}`;
  }
  const found = commentAtCell(sheet, row, col);
  if (found && found.comment.threaded) {
    const at = cellNow(sheet, found.comment) || noteCell(sheet, row, col);
    return `Cannot change the comment in ${cellName(at.row, at.col)}: it is a threaded comment`;
  }
  return undefined;
}

/**
 * Sets the note of a cell: `text` becomes its text, and a cell without a note gets one by
 * `options.author`. One rule for the text: `undefined` deletes the note; the text the note
 * has already (line breaks aside) changes nothing; any other text that is empty or nothing
 * but white space deletes the note. In a merged cell the note is the one of the merged cell
 * (see commentAtCell()), and a new one is put on its top-left cell.
 *
 * The text is plain: a note of the file whose text has several formats has one after it was
 * edited. Its author, its box and whether the box is always shown stay.
 *
 * Returns no operation when nothing would change. Throws RangeError for a text longer than
 * MAX_NOTE_TEXT, for a cell with a threaded comment, and for a cell or sheet that does not exist.
 */
export function setComment(workbook: Workbook, sheetIndex: number, row: number, col: number, text: string | undefined, options: NoteOptions = {}): Op[] {
  const sheet = checkTarget(workbook, sheetIndex, row, col);
  const problem = noteProblem(workbook, sheetIndex, row, col);
  if (problem) throw new RangeError(problem);
  const comments = sheet.comments || [];
  const found = commentAtCell(sheet, row, col);
  // A text field hands out "\n"; whatever else a caller passes, and what the file has, is compared as that.
  const lines = (value: string) => value.replace(/\r\n?/g, '\n');
  let value = text === undefined || text === null ? undefined : lines(String(text));
  // The text the note has is no change, whatever it is (an empty note of the file stays one).
  if (value !== undefined && found && lines(String(found.comment.text)) === value) return [];
  if (value !== undefined && value.trim() === '') value = undefined;
  if (value !== undefined && value.length > MAX_NOTE_TEXT) throw new RangeError(`The text of a note cannot be longer than ${MAX_NOTE_TEXT} characters`);

  let next: SheetComment[];
  if (value === undefined) {
    if (!found) return [];
    next = comments.slice();
    next.splice(found.index, 1);
  } else if (found) {
    const edited: SheetComment = { ...found.comment, text: value };
    delete edited.runs;
    next = comments.slice();
    next[found.index] = edited;
  } else {
    // Notes that went with their rows or columns are still in the list (undo brings them back) and do not count.
    let live = 0;
    if (comments.length >= MAX_SHEET_NOTES) for (let i = 0; i < comments.length; i++) if (cellNow(sheet, comments[i])) live++;
    if (live >= MAX_SHEET_NOTES) throw new RangeError('The sheet has too many comments');
    const at = noteCell(sheet, row, col);
    const named = typeof options.author === 'string' ? options.author.trim().slice(0, MAX_AUTHOR) : '';
    const added: SheetComment = { row: at.row, col: at.col, text: value, author: named !== '' ? named : DEFAULT_NOTE_AUTHOR, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
    next = comments.concat([added]);
  }
  return [{ type: 'setComments', sheet: sheetIndex, comments: next.length > 0 ? next : undefined }];
}

/** Whether the notes of a sheet can be changed at all (see Sheet.notesBlocked). */
export function notesBlocked(sheet: Sheet): boolean {
  return sheet.notesBlocked !== undefined && Object.prototype.hasOwnProperty.call(BLOCKED, sheet.notesBlocked);
}

/**
 * Deletes the notes on the cells of a range, as cutting the cells does: the notes go with
 * what was cut. `rows` limits it to those rows of the range (the rows a filter shows).
 * Threaded comments stay, and so does every note of a sheet whose notes cannot be changed:
 * nothing is refused, and no operation is returned when no note goes.
 */
export function clearNotes(workbook: Workbook, sheetIndex: number, range: { top: number; left: number; bottom: number; right: number }, rows?: readonly number[]): Op[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  const comments = sheet.comments;
  if (!comments || comments.length === 0 || sheet.kind !== 'worksheet' || notesBlocked(sheet)) return [];
  const wanted = rows ? new Set(rows) : undefined;
  const kept: SheetComment[] = [];
  for (let i = 0; i < comments.length; i++) {
    const comment = comments[i];
    const cell = comment && !comment.threaded ? cellNow(sheet, comment) : undefined;
    const inside = cell !== undefined && cell.row >= range.top && cell.row <= range.bottom && cell.col >= range.left && cell.col <= range.right && (!wanted || wanted.has(cell.row));
    if (!inside) kept.push(comment);
  }
  if (kept.length === comments.length) return [];
  return [{ type: 'setComments', sheet: sheetIndex, comments: kept.length > 0 ? kept : undefined }];
}
