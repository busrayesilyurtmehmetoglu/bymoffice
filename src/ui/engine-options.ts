// What a view makes its calculation engine with. Kept apart from the view so that it can be
// tested without a document to draw into.

import type { CalcEngineOptions } from '../calc/engine';
import type { FormatLocale } from '../format/locale';

/** Longest slice of a recalculation before the view lets the page breathe. */
export const CALC_SLICE_MS = 250;

/**
 * Options of the engine of a view: the time one slice may take, and the locale of the view
 * as it is whenever a formula is calculated (UPPER, LOWER and PROPER follow it).
 */
export function viewEngineOptions(locale: () => FormatLocale): CalcEngineOptions {
  return { budget: { milliseconds: CALC_SLICE_MS }, locale };
}
