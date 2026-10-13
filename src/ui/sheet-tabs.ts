// What the sheet tabs of the workbook view decide without touching the DOM:
// which sheet is shown once one is gone, and what a refused operation on a
// sheet says in the language of the view.

import type { Sheet } from '../model/sheet/types';
import type { UiStrings } from './strings';

/** What an undo step did to the list of sheets. */
export type SheetOpKind = 'insert' | 'delete' | 'rename' | 'move' | 'state' | 'copy' | 'color';

/**
 * The sheet to show when the one at `index` cannot be (it was deleted, or the index is
 * past the end): the next sheet with a tab from there, else the nearest one before, else
 * the nearest sheet there is. -1 for an empty list.
 */
export function sheetToShow(sheets: readonly Pick<Sheet, 'state'>[], index: number): number {
  const count = sheets.length;
  if (count === 0) return -1;
  const from = Number.isInteger(index) ? Math.min(Math.max(index, 0), count) : 0;
  for (let i = from; i < count; i++) if (sheets[i].state === 'visible') return i;
  for (let i = from - 1; i >= 0; i--) if (sheets[i].state === 'visible') return i;
  return Math.min(from, count - 1);
}

/** Every index of a list of sheets: what an operation on the list itself reports as changed. */
export function allSheets(count: number): number[] {
  const out: number[] = [];
  for (let i = 0; i < count; i++) out.push(i);
  return out;
}

/**
 * What a refused operation on a sheet says, from the message of the operations layer;
 * undefined for a message that is about something else.
 */
export function sheetErrorText(message: string, strings: UiStrings): string | undefined {
  let match: RegExpExecArray | null;
  if (message.indexOf('A sheet name has 1 to 31 characters') === 0) return strings.badSheetName;
  // The name is the user's own text and may hold anything, line breaks aside (those are refused before).
  if ((match = /^Cannot name the sheet "([\s\S]*)": another sheet has that name$/.exec(message))) return strings.sheetNameTaken(match[1]);
  if (message === 'Cannot delete the last visible sheet') return strings.cannotDeleteLastSheet;
  if (message === 'Cannot hide the last visible sheet') return strings.cannotHideLastSheet;
  if (/^Cannot copy a \S+: not supported yet$/.test(message)) return strings.sheetKindFixed('copy');
  if (/^Cannot colour the tab of a \S+: not supported yet$/.test(message)) return strings.sheetKindFixed('color');
  if ((match = /^Cannot (rename|delete) a sheet (?:in a workbook )?with (.+): not supported yet$/.exec(message))) {
    return strings.sheetBlockedBy(match[1] as 'rename' | 'delete', match[2]);
  }
  if ((match = /^Cannot (rename|delete) a \S+: not supported yet$/.exec(message))) return strings.sheetKindFixed(match[1] as 'rename' | 'delete');
  if ((match = /^Cannot (rename|delete) the sheet: the formula of ([\s\S]+) cannot be read$/.exec(message))) {
    return strings.sheetFormulaUnread(match[1] as 'rename' | 'delete', match[2]);
  }
  if ((match = /^Cannot change the sheets: "([\s\S]*)" and "([\s\S]*)" differ only by case$/.exec(message))) return strings.sheetNamesClash(match[1], match[2]);
  return undefined;
}

/** A cell by the index of its sheet. */
export interface SheetCellAddress {
  sheet: number;
  row: number;
  col: number;
}

/** What recalculateForSheets() needs of the calculation engine. */
export interface SheetCalcEngine {
  /** A calculation ran out of its time and has formulas left to calculate. */
  readonly incomplete: boolean;
  /** The dependency graph is built and nothing waits to be put into it. */
  readonly graphReady: boolean;
  invalidate(): void;
  recalculate(changed?: SheetCellAddress[], options?: { overwrite?: boolean; sheetsOfBefore?: boolean }): unknown;
}

/** What happened to the list of sheets, as far as formulas care. */
export interface SheetListChange {
  /**
   * "names": a sheet got another name, or an empty sheet came or went; no formula is worth
   * something else. "gone": a sheet went, and the formulas that read it were rewritten to
   * #REF!. "back": such a sheet came back, and those formulas with it.
   */
  kind: 'names' | 'gone' | 'back';
  /**
   * The cells whose formulas were rewritten. For "gone" by the indexes the sheets had before
   * the sheet went (the engine still has that list), for "back" by the indexes they have now.
   * Missing when they are not known or too many: everything is calculated then.
   */
  cells?: SheetCellAddress[];
  /** Something else than those cells may have changed what formulas are worth (defined names were rewritten). */
  unsure?: boolean;
}

/**
 * Brings a calculation engine in line with a list of sheets that changed, calculating as
 * little as can be: the values a file has for its formulas are worth more than this
 * engine's, and stay where nothing the formula reads has changed.
 *
 * - The engine is always built again (it knows sheets by index and by name).
 * - "names": nothing is calculated, not even formulas that give another value each time.
 * - "gone": the engine is told about the rewritten cells while it still has the graph of
 *   before: only so does it see that their formulas changed, and that their values and the
 *   values of what reads them give way. Then the graph is forgotten.
 * - "back": the cells come back with the values they had; they and what reads them are
 *   calculated with the new graph.
 * - A calculation that was under way would be forgotten with the graph, leaving its formulas
 *   with the values of before for good; then, and whenever the cells are not known,
 *   everything is calculated and the values of the file give way. Rare, and right.
 *
 * The engine may be left `incomplete`: the caller goes on with it as after any calculation.
 */
export function recalculateForSheets(engine: SheetCalcEngine, change: SheetListChange): void {
  const cells = change.cells;
  const all = engine.incomplete || change.unsure === true || (change.kind !== 'names' && !cells);
  if (all) {
    engine.invalidate();
    engine.recalculate(undefined, { overwrite: true });
    return;
  }
  if (change.kind === 'names' || !cells || cells.length === 0) {
    engine.invalidate();
    return;
  }
  if (change.kind === 'back') {
    engine.invalidate();
    engine.recalculate(cells);
    return;
  }
  if (!engine.graphReady) {
    // No graph of before to see the change with.
    engine.invalidate();
    engine.recalculate(undefined, { overwrite: true });
    return;
  }
  // By the list of before: the engine would otherwise see that the list changed and forget the graph by itself.
  engine.recalculate(cells, { sheetsOfBefore: true });
  const left = engine.incomplete;
  engine.invalidate();
  if (left) engine.recalculate(undefined, { overwrite: true });
}

/**
 * Where a tab that is dragged goes: the index moveSheet() is to be given for the sheet at
 * `from` so that it stands before the tab `gap` of the visible ones (`visible`: the indexes
 * of the sheets that have a tab, in order; a gap of their number is behind the last).
 * Undefined when the sheet would stay where it is, between the same tabs.
 */
export function dropIndex(visible: readonly number[], from: number, gap: number): number | undefined {
  const count = visible.length;
  if (count === 0 || !(gap >= 0 && gap <= count) || Math.floor(gap) !== gap) return undefined;
  // Beside itself on either side: nothing moves.
  if ((gap < count && visible[gap] === from) || (gap > 0 && visible[gap - 1] === from)) return undefined;
  if (gap === count) {
    const last = visible[count - 1];
    return (last > from ? last - 1 : last) + 1;
  }
  const before = visible[gap];
  return before > from ? before - 1 : before;
}
