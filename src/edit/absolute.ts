// F4 while a formula is typed: the references at the caret, or in the selection,
// go from relative to absolute and back in the order Excel has: A1, $A$1, A$1,
// $A1, A1. Nothing in here touches the DOM; the workbook view writes the text.

import { formulaReferences, type FormulaReference } from '../calc/refs';
import { MAX_POINT_TEXT, isFormulaText } from './point';

/** The reference written in the state that follows the one it is in. */
function nextSpelling(ref: FormulaReference): string {
  // A whole column or row has one side only: with "$" or without.
  if (ref.kind === 'cols') return ref.write(!ref.absCol, false);
  if (ref.kind === 'rows') return ref.write(false, !ref.absRow);
  if (!ref.absCol && !ref.absRow) return ref.write(true, true);
  if (ref.absCol && ref.absRow) return ref.write(false, true);
  if (!ref.absCol && ref.absRow) return ref.write(true, false);
  return ref.write(false, false);
}

/**
 * Moves every reference that the selection from `start` to `end` touches (the reference the
 * caret stands in, before or right behind when they are the same) on to its next state:
 * a cell from relative to absolute, to an absolute row, to an absolute column and back; a
 * whole column or row from relative to absolute and back. Both corners of a range go
 * together, from the state the first corner is in. The name of a sheet before a reference
 * stays as written.
 *
 * Returns the new text and the selection that belongs to it: for a caret, a caret right
 * behind the reference, where typing goes on and where the next call finds the reference
 * again; for a selection, the selection grown to hold the references it touched. Undefined
 * when there is nothing to do: the text is not a formula, or no reference is touched (a
 * text in quotes holds none, and the name of a function is none).
 */
export function cycleAbsolute(text: string, start: number, end: number): { text: string; start: number; end: number } | undefined {
  if (!isFormulaText(text) || text.length > MAX_POINT_TEXT) return undefined;
  if (!Number.isInteger(start) || !Number.isInteger(end) || start < 0 || end < start || end > text.length) return undefined;
  const touched = formulaReferences(text).filter((ref) => ref.start <= end && ref.end >= start);
  if (touched.length === 0) return undefined;
  const first = touched[0];
  const last = touched[touched.length - 1];
  let out = '';
  let copied = 0;
  // How much longer the text is behind the last reference but one, and behind the last.
  let grownBefore = 0;
  let grown = 0;
  for (const ref of touched) {
    const spelling = nextSpelling(ref);
    out += text.substring(copied, ref.bodyStart) + spelling;
    copied = ref.end;
    grownBefore = grown;
    grown += spelling.length - (ref.end - ref.bodyStart);
  }
  out += text.substring(copied);
  const lastEnd = last.end + grown;
  // A caret stays a caret: typing goes on behind the reference, and the next call finds it there.
  if (start === end) return { text: out, start: lastEnd, end: lastEnd };
  // Nothing before the first reference itself has moved.
  const newStart = Math.min(start, first.bodyStart);
  let newEnd: number;
  if (end < last.bodyStart) newEnd = end + grownBefore;
  else newEnd = end <= last.end ? lastEnd : end + grown;
  return { text: out, start: newStart, end: newEnd };
}
