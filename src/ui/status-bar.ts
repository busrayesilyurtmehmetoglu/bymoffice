// The status bar: a thin strip under the grid. On the left a badge while the
// workbook is read-only and the reason a command failed; on the right the
// sum, average and count of the selected cells.

import type { FormatLocale } from '../format/locale';
import { formatValue } from '../format/number-format';
import type { Cell, CellRange, Sheet } from '../model/sheet/types';
import type { EditorStrings } from './editor-strings';

/** Selections with more cells than this inside the used area are not summed. */
export const MAX_SUMMARY_CELLS = 200000;

export interface SelectionSummary {
  sum: number;
  /** Cells that hold a number. */
  numbers: number;
  /** Cells that hold anything at all. */
  count: number;
}

/**
 * Adds up the numbers of a range. Only the part of the range that can hold
 * cells is visited; undefined when that part is larger than MAX_SUMMARY_CELLS.
 */
export function summarize(sheet: Sheet, range: CellRange): SelectionSummary | undefined {
  const summary: SelectionSummary = { sum: 0, numbers: 0, count: 0 };
  const top = Math.max(range.top, 0);
  const left = Math.max(range.left, 0);
  const bottom = Math.min(range.bottom, sheet.rowCount - 1);
  const right = Math.min(range.right, sheet.colCount - 1);
  if (!(bottom >= top && right >= left)) return summary;
  if ((bottom - top + 1) * (right - left + 1) > MAX_SUMMARY_CELLS) return undefined;
  for (let r = top; r <= bottom; r++) {
    const row = sheet.rows.get(r);
    if (!row) continue;
    const cells = row.cells;
    const last = Math.min(right, cells.length - 1);
    for (let c = left; c <= last; c++) {
      const cell: Cell | undefined = cells[c];
      if (!cell || cell.value === null || cell.value === undefined || cell.kind === 'blank') continue;
      summary.count++;
      if (cell.kind === 'number' && typeof cell.value === 'number' && isFinite(cell.value)) {
        summary.sum += cell.value;
        summary.numbers++;
      }
    }
  }
  return summary;
}

export class StatusBar {
  readonly element: HTMLDivElement;
  private readonly badge: HTMLSpanElement;
  private readonly message: HTMLSpanElement;
  private readonly summary: HTMLSpanElement;
  private strings: EditorStrings;

  constructor(strings: EditorStrings) {
    this.strings = strings;
    const element = document.createElement('div');
    element.className = 'bym-status';
    element.setAttribute('role', 'group');

    const badge = document.createElement('span');
    badge.className = 'bym-status-badge';
    badge.hidden = true;

    const message = document.createElement('span');
    message.className = 'bym-status-message';
    message.setAttribute('role', 'status');

    const summary = document.createElement('span');
    summary.className = 'bym-status-summary';

    element.appendChild(badge);
    element.appendChild(message);
    element.appendChild(summary);
    this.element = element;
    this.badge = badge;
    this.message = message;
    this.summary = summary;
    this.setStrings(strings);
  }

  setStrings(strings: EditorStrings): void {
    this.strings = strings;
    this.element.setAttribute('aria-label', strings.statusBar);
    this.badge.textContent = strings.readOnly;
  }

  setReadOnly(readOnly: boolean): void {
    this.badge.hidden = !readOnly;
  }

  /** Shows why a command failed; an empty text clears it. */
  showMessage(text: string): void {
    if (this.message.textContent !== text) this.message.textContent = text;
  }

  /**
   * Shows sum, average and count for a selection of more than one cell with at
   * least one number in it; nothing otherwise, and nothing for a selection too
   * large to add up.
   */
  showSummary(sheet: Sheet | undefined, range: CellRange, single: boolean, locale: FormatLocale): void {
    let text = '';
    if (sheet && sheet.kind === 'worksheet' && !single) {
      const summary = summarize(sheet, range);
      if (summary && summary.numbers > 0) {
        const show = (value: number) => formatValue(value, 'General', { locale }).text;
        text = this.strings.summary(show(summary.sum), show(summary.sum / summary.numbers), show(summary.count));
      }
    }
    if (this.summary.textContent !== text) this.summary.textContent = text;
  }

  destroy(): void {
    if (this.element.parentNode) this.element.parentNode.removeChild(this.element);
  }
}
