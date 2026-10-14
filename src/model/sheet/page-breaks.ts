// Where a manual page break is after rows and columns were inserted or deleted: one rule, used
// by the layout of the printed pages, by the commands of the page setup and by the writer of the
// file, so that what is printed, what the model says and what the file says agree.
//
// The rule is that of the file: <brk id="n"> breaks the page after row (or column) n, counted
// from 1, so a break belongs to the row ABOVE it (the last row of its page). In the model a break
// is `at` = n, the first row of the next page counted from 0: the row it belongs to is at - 1.
//  - Rows inserted at or above that row push the break down; rows inserted below it (at `at`,
//    right under the break) go onto the next page.
//  - When that row is deleted, the break goes to the row before the deleted ones (UNSURE: what
//    Excel does there is not known); a break that comes to stand above the first row is gone.
//  - Along the other axis a break spans columns (of a row break) from min to max; those follow the
//    inserted and deleted columns, a break that spans every column keeps doing so, and a break
//    whose columns are all deleted is gone.

import { MAX_COLS, MAX_ROWS, type SheetStructuralEdit } from './types';

/**
 * The band of the other axis a <brk> spans, as the attributes say it: what is missing or no
 * number is 0, the default of the format. The reader of the file and its writer both read it so.
 */
export function breakSpanOf(min: string | undefined, max: string | undefined): { min: number; max: number } {
  const read = (text: string | undefined) => (text !== undefined && /^[0-9]{1,9}$/.test(text) ? Number(text) : 0);
  return { min: read(min), max: read(max) };
}

/** The place of a break and the band of the other axis it spans. */
export interface BreakPlaceSpan {
  at: number;
  min: number;
  max: number;
}

/** Where a break is after one edit; undefined when it is gone. `wholeSpan`: min and max are not looked at (the break spans the whole other axis). */
function moveOnce(place: BreakPlaceSpan, edit: SheetStructuralEdit, axis: 'row' | 'col', wholeSpan: boolean): BreakPlaceSpan | undefined {
  if (!edit || !(edit.count > 0) || !(edit.at >= 0)) return place;
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  const crossLimit = axis === 'row' ? MAX_COLS : MAX_ROWS;
  const { at: start, count } = edit;
  if (edit.axis === axis) {
    let anchor = place.at - 1;
    if (edit.kind === 'insert') {
      if (anchor >= start) anchor += count;
    } else if (anchor >= start + count) anchor -= count;
    else if (anchor >= start) anchor = start - 1;
    const at = anchor + 1;
    if (at < 1 || at >= limit) return undefined;
    return at === place.at ? place : { ...place, at };
  }
  if (wholeSpan || (place.min === 0 && place.max >= crossLimit - 1)) return place;
  let { min, max } = place;
  if (edit.kind === 'insert') {
    if (min >= start) {
      min += count;
      max += count;
    } else if (max >= start) max += count;
    if (min >= crossLimit) return undefined;
    max = Math.min(max, crossLimit - 1);
  } else {
    const end = start + count;
    if (max < start) return place;
    const newMin = min < start ? min : min >= end ? min - count : start;
    const newMax = max >= end ? max - count : start - 1;
    if (newMax < newMin) return undefined;
    min = newMin;
    max = newMax;
  }
  return min === place.min && max === place.max ? place : { ...place, min, max };
}

/**
 * Where a break is after the edits, in order (see the rule above); undefined when it is gone.
 * The object comes back as it is when nothing moved it. With `wholeSpan` the band of the other
 * axis is left alone.
 */
export function movePageBreak<T extends BreakPlaceSpan>(place: T, edits: readonly SheetStructuralEdit[] | undefined, axis: 'row' | 'col', wholeSpan = false): T | undefined {
  if (!(place.at >= 1)) return undefined;
  let now: BreakPlaceSpan | undefined = place;
  for (let i = 0; edits && i < edits.length && now; i++) now = moveOnce(now, edits[i], axis, wholeSpan);
  return now as T | undefined;
}
