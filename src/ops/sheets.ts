// Operations on the sheets of a workbook as the editor asks for them: a new
// sheet, a sheet deleted, a sheet renamed. Like the commands on cells these
// functions only build operations; the caller applies them through a History
// so that each is one undo step. What follows from them (formulas, defined
// names and links that name the sheet) is worked out when they are applied.

import { foldSheetName, type Workbook } from '../model/sheet/types';
import { deleteSheetProblem, insertSheetProblem, renameSheetProblem, type Op } from './ops';

export interface InsertSheetOptions {
  /** The name of the new sheet; the next free default name of the language without it. */
  name?: string;
  /** The language default names are in, as a tag ("tr", "tr-TR", "en"): Turkish gives "Sayfa1", every other "Sheet1". */
  locale?: string;
}

/** Highest number a default name is tried with; a workbook never has that many sheets. */
const MAX_DEFAULT_NUMBER = 1000000;

/**
 * The name a new sheet gets when none is given: "Sheet" (Turkish: "Sayfa") with the lowest
 * number from 1 that no sheet of the workbook has, whatever the case of its letters.
 */
export function defaultSheetName(workbook: Workbook, locale?: string): string {
  const base = typeof locale === 'string' && /^tr(?:[-_]|$)/i.test(locale.trim()) ? 'Sayfa' : 'Sheet';
  const taken = new Set<string>();
  for (const sheet of workbook.sheets) taken.add(foldSheetName(sheet.name));
  for (let n = 1; n <= MAX_DEFAULT_NUMBER; n++) {
    const name = base + n;
    if (!taken.has(foldSheetName(name))) return name;
  }
  throw new RangeError('The workbook has too many sheets');
}

/**
 * Why the sheet at `index` cannot be renamed, as a message; undefined when it can. With
 * `name` the name is checked as well (its characters and length, and that no other sheet
 * has it). Cheap: no cell is looked at, so a formula that cannot be read is only found
 * when the operation is applied.
 */
export function canRenameSheet(workbook: Workbook, index: number, name?: string): string | undefined {
  return renameSheetProblem(workbook, index, name);
}

/** Why the sheet at `index` cannot be deleted, as a message; undefined when it can. Cheap, like canRenameSheet(). */
export function canDeleteSheet(workbook: Workbook, index: number): string | undefined {
  return deleteSheetProblem(workbook, index);
}

/** Why no sheet can be inserted at `at` under `name` (or under the default name), as a message; undefined when one can. */
export function canInsertSheet(workbook: Workbook, at: number, name?: string): string | undefined {
  return insertSheetProblem(workbook, at, name === undefined ? defaultSheetName(workbook) : name);
}

/**
 * Renames the sheet at `index`. Formulas, defined names and links of the whole workbook that
 * name the sheet are written with the new name when the operation is applied. Nothing to do
 * (an empty list) when the sheet has that name already. Throws RangeError for a name a sheet
 * cannot have or another sheet has, and for a sheet that cannot be renamed.
 */
export function renameSheet(workbook: Workbook, index: number, name: string): Op[] {
  const problem = renameSheetProblem(workbook, index, name);
  if (problem) throw new RangeError(problem);
  if (workbook.sheets[index].name === name) return [];
  return [{ type: 'renameSheet', sheet: index, name }];
}

/** Deletes the sheet at `index`. Throws RangeError for a sheet that cannot be deleted (the last visible one, say). */
export function deleteSheet(workbook: Workbook, index: number): Op[] {
  const problem = deleteSheetProblem(workbook, index);
  if (problem) throw new RangeError(problem);
  return [{ type: 'deleteSheet', sheet: index }];
}

/**
 * Inserts an empty worksheet so that it is sheet number `at` (0: first; the number of
 * sheets: last). Throws RangeError for an index outside that and for a name that cannot be used.
 */
export function insertSheet(workbook: Workbook, at: number, options: InsertSheetOptions = {}): Op[] {
  const name = options.name !== undefined ? options.name : defaultSheetName(workbook, options.locale);
  const problem = insertSheetProblem(workbook, at, name);
  if (problem) throw new RangeError(problem);
  return [{ type: 'insertSheet', sheet: at, name }];
}
