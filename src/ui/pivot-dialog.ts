// The dialog that makes a pivot table: the cells it is made from (a range, on this sheet or
// another, or a defined name) and where it goes (a new sheet, or a cell of the shown sheet).
// What the fields are read as has no DOM in it and is tested on its own.

import { quoteSheetName } from '../calc/refs';
import { foldSheetName, parseRange, rangeName, type CellRange, type Workbook } from '../model/sheet/types';
import type { PivotSourceSpec } from '../ops/pivot-create';
import { Dialog, FieldError } from './rules-dialog';
import type { PivotStrings } from './strings';

/** A reference to cells as the dialog takes it: "A1:D20", "Data!A1:D20", "'My data'!$A$1:$D$20". */
const QUALIFIED = /^(?:'((?:[^']|'')+)'|([^!'"]+))!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)$/;

/** What the field of the cells says: a range of a sheet, or a defined name; undefined for what is neither. */
export function readPivotSource(text: string, workbook: Workbook, sheet: number): PivotSourceSpec | undefined {
  const typed = text.trim().replace(/^=\s*/, '');
  if (typed === '' || typed.length > 1024) return undefined;
  const qualified = QUALIFIED.exec(typed);
  if (qualified) {
    const name = qualified[1] !== undefined ? qualified[1].replace(/''/g, "'") : qualified[2];
    const index = workbook.sheets.findIndex((one) => foldSheetName(one.name) === foldSheetName(name));
    const range = parseRange(qualified[3].replace(/\$/g, ''));
    return index >= 0 && range ? { sheet: index, range } : undefined;
  }
  const range = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?$/.test(typed) ? parseRange(typed.replace(/\$/g, '')) : undefined;
  if (range) return { sheet, range };
  const named = workbook.definedNames.find((one) => one.name.toLowerCase() === typed.toLowerCase() && one.localSheet === undefined);
  return named ? { name: named.name } : undefined;
}

/** The text the field of the cells starts with: the range, on the shown sheet without its name. */
export function pivotSourceText(workbook: Workbook, sheet: number, range: CellRange, shown: number): string {
  const name = rangeName(range);
  return sheet === shown ? name : `${quoteSheetName(workbook.sheets[sheet].name)}!${name}`;
}

/** The cell of a field ("F2"); undefined for anything else. */
export function readPivotCell(text: string): { row: number; col: number } | undefined {
  const typed = text.trim().replace(/\$/g, '');
  if (!/^[A-Za-z]{1,3}\d{1,7}$/.test(typed)) return undefined;
  const range = parseRange(typed);
  return range ? { row: range.top, col: range.left } : undefined;
}

export interface PivotDialogOptions {
  parent: HTMLElement;
  strings: PivotStrings;
  /** What the fields start with. */
  source: string;
  cell: string;
  coarse?: boolean;
  /** Makes the table; what it throws is shown in the dialog, which then stays open. */
  onApply(source: string, place: { newSheet: true } | { cell: string }): void;
  explain(error: unknown): string;
  onClose(): void;
}

/** The dialog that makes a pivot table. */
export class PivotDialog {
  readonly dialog: Dialog;

  constructor(options: PivotDialogOptions) {
    const strings = options.strings;
    const dialog = new Dialog({ parent: options.parent, title: strings.createTitle, name: 'pivot', onClose: options.onClose, coarse: options.coarse });
    this.dialog = dialog;
    const body = dialog.body;
    const source = dialog.input(1024);
    source.value = options.source;
    dialog.addField(body, strings.source, source, 'source');
    const place = dialog.select([
      ['new', strings.newSheet],
      ['here', strings.thisSheet],
    ]);
    place.value = 'new';
    dialog.addField(body, strings.place, place, 'place');
    const cell = dialog.input(16);
    cell.value = options.cell;
    const row = dialog.addField(body, strings.cell, cell, 'cell');
    const sync = () => {
      row.style.display = place.value === 'here' ? '' : 'none';
    };
    place.addEventListener('change', sync);
    sync();
    dialog.addAction(strings.cancel, () => dialog.close()).setAttribute('data-bym-action', 'cancel');
    dialog
      .addAction(
        strings.ok,
        () => {
          dialog.showError('');
          try {
            if (place.value === 'here' && !readPivotCell(cell.value)) throw new FieldError(strings.badCell, 'cell');
            options.onApply(source.value, place.value === 'here' ? { cell: cell.value } : { newSheet: true });
          } catch (error) {
            const field = error instanceof FieldError ? (error.field === 'cell' ? cell : source) : undefined;
            dialog.showError(error instanceof FieldError ? error.message : options.explain(error), field);
            return;
          }
          dialog.close();
        },
        'primary',
      )
      .setAttribute('data-bym-action', 'ok');
    dialog.markPristine();
    dialog.focusFirst();
  }
}
