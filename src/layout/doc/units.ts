// Unit conversion and the limits the document layout engine works within.
//
// The model keeps the units of the file; layout works in CSS pixels at 96 dpi.

/** Twips (1/20 pt) -> px. */
export const twips = (value: number | undefined): number => (value === undefined ? 0 : value / 15);
/** Half-points -> px. */
export const halfPoints = (value: number): number => (value * 2) / 3;
/** Points -> px. */
export const points = (value: number): number => (value * 4) / 3;
/** Eighths of a point (border widths) -> px. */
export const eighths = (value: number): number => value / 6;
/** EMU -> px. */
export const emu = (value: number | undefined): number => (value === undefined ? 0 : value / 9525);

export const clamp = (value: number, min: number, max: number): number => (value < min ? min : value > max ? max : value);

/** A finite number, or the fallback: values from a measurer or from arithmetic on file values are not trusted. */
export const finite = (value: number, fallback = 0): number => (value === value && value !== Infinity && value !== -Infinity ? value : fallback);

/** US Letter, what Word assumes for a section that states no page size; twips. */
export const DEFAULT_PAGE_WIDTH = 12240;
export const DEFAULT_PAGE_HEIGHT = 15840;
/** One inch, Word's margin when a section states none; twips. */
export const DEFAULT_MARGIN = 1440;
export const DEFAULT_HEADER_DISTANCE = 720;
export const DEFAULT_COLUMN_SPACE = 720;
/** Left and right cell margin of a table that states none; twips. */
export const DEFAULT_CELL_MARGIN = 108;

/** Smallest paper side and text extent the engine lays out in, px: keeps degenerate sections from dividing by zero. */
export const MIN_PAGE_SIDE = 16;
export const MIN_TEXT_EXTENT = 8;
/** Largest paper side, px (Word's own limit is 22 inches). */
export const MAX_PAGE_SIDE = 20000;

/** Word allows 45 columns; a section that asks for more gets this many. */
export const MAX_COLUMNS = 45;
/** Pages after which layout stops for good. */
export const MAX_PAGES = 100000;
/** Lines one paragraph may break into; what is left after that goes into the last line. */
export const MAX_PARAGRAPH_LINES = 200000;
/** Depth of tables in tables and boxes in boxes that is laid out; deeper content is skipped. */
export const MAX_LAYOUT_DEPTH = 70;
/** Narrowest gap beside a wrapped drawing that text is put in, px (Word uses about an inch). */
export const MIN_WRAP_WIDTH = 96;

/** Tolerance for comparing positions that were added up in a different order. */
export const EPSILON = 0.01;
