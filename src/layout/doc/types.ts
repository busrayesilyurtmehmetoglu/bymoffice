// What the document layout engine produces and what it needs from its host.
// The engine itself never touches a canvas: text is measured through
// TextMeasurer, so the same code lays out in a browser and in Node (tests).
//
// All lengths here are CSS pixels at 96 dpi (1 pt = 4/3 px, 1 twip = 1/15 px,
// 1 EMU = 1/9525 px). The model keeps the units of the file; the engine
// converts once.

/** A font as text is measured and painted with it. */
export interface LayoutFont {
  /** Family name from the document, already resolved through styles and theme. */
  family: string;
  /** Size in px. */
  size: number;
  bold: boolean;
  italic: boolean;
}

/** Vertical extent of a font, in px, measured from the baseline. */
export interface FontMetrics {
  ascent: number;
  descent: number;
  /** Extra space Word adds between lines of this font ("external leading"); 0 when unknown. */
  lineGap: number;
}

/** How the engine measures text. Implemented over a canvas in the browser and by a fixed table in tests. */
export interface TextMeasurer {
  /** Advance width of the text in px. */
  width(font: LayoutFont, text: string): number;
  metrics(font: LayoutFont): FontMetrics;
  /**
   * Optional: false when the host cannot show the family. The engine then
   * measures with DocLayoutOptions.fallbackFamily and names that family in
   * the items.
   */
  hasFont?(family: string): boolean;
}

/** Text on one baseline in one font. */
export interface TextItem {
  type: 'text';
  /** Left edge and baseline. */
  x: number;
  y: number;
  width: number;
  text: string;
  font: LayoutFont;
  /** "#RRGGBB". */
  color: string;
  underline?: { color: string; style: string };
  strike?: boolean | 'double';
  /** Background behind the text (highlight or run shading), "#RRGGBB". */
  highlight?: string;
  /** Target of a hyperlink, already passed through safeHyperlinkTarget. */
  link?: string;
  /** Extra advance after every character, px (may be negative); `width` includes it. */
  letterSpacing?: number;
  /** Horizontal stretch of the glyphs, 1 = none; `width` includes it. */
  scale?: number;
  /**
   * The text is turned around (x, y), the start of its baseline: clockwise by
   * 90 degrees (it reads downwards) or by 270 (upwards). Text in table cells
   * with a vertical text direction. `width` is the advance along the baseline.
   */
  rotation?: 90 | 270;
  /**
   * The item repeats text that is laid out elsewhere (a table header row
   * repeated at the top of a page): it has no place of its own in the story.
   */
  repeat?: true;
  /**
   * Where the text comes from, for selection and hit testing later: story id
   * ("body", "header:<rId>", "footer:<rId>", "footnote:<id>", "endnote:<id>"),
   * the path of the paragraph (index into the blocks of the story; inside a
   * table: table index, row, cell, block index, ...; inside a text box: the
   * path of the anchoring paragraph, then -1, the number of the drawing in
   * the paragraph, the number of the box in the drawing and the block index)
   * and the offset of the first character in paragraphText() of that
   * paragraph. Text the engine makes up (list numbers, tab leaders, note
   * numbers) has no source.
   */
  source?: { story: string; path: number[]; offset: number };
}

/** A filled rectangle: paragraph or cell shading, page background. */
export interface RectItem {
  type: 'rect';
  x: number;
  y: number;
  width: number;
  height: number;
  fill: string;
}

/** A straight line: borders of paragraphs, tables and pages. */
export interface LineItem {
  type: 'line';
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color: string;
  /** Thickness in px. */
  width: number;
  /** Word's border style name ("single", "double", "dashed", "dotted", ...). */
  style: string;
}

/**
 * A shape: a geometry in the box (x, y, width, height), filled and outlined.
 * A "line" runs from the top left to the bottom right corner of the box
 * before it is flipped; a box of no height is a horizontal line.
 */
export interface ShapeItem {
  type: 'shape';
  x: number;
  y: number;
  width: number;
  height: number;
  geometry: 'rect' | 'roundRect' | 'ellipse' | 'line';
  /** For "roundRect": the radius of the corners, px. */
  radius?: number;
  /** "#RRGGBB"; not filled when missing. */
  fill?: string;
  /** The outline: thickness in px, and how it is broken up, by the name of a DrawingML dash ("dash", "sysDot", ...). */
  stroke?: { color: string; width: number; dash?: string };
  /** Clockwise rotation in degrees around the centre. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
}

/** A picture. The painter loads the part and draws it; until it is loaded it draws a placeholder. */
export interface ImageItem {
  type: 'image';
  x: number;
  y: number;
  width: number;
  height: number;
  /** Name of the package part that holds the picture ("/word/media/image1.png"); undefined for a linked or missing one. */
  part?: string;
  /** Fractions of the picture cut off at each side, 0..1. */
  crop?: { left: number; top: number; right: number; bottom: number };
  /** Clockwise rotation in degrees around the centre. */
  rotation?: number;
  /** Alternative text, for accessibility. */
  alt?: string;
  /**
   * What stands here when it is not a picture the painter can load: "chart",
   * "diagram", "shape", "vml", "object", ... The painter draws a placeholder
   * of the size of the item.
   */
  placeholder?: string;
  /** Room the effects of the picture (shadow, glow) take beyond its box, px. */
  effect?: { left: number; top: number; right: number; bottom: number };
  flipH?: boolean;
  flipV?: boolean;
}

/**
 * Where a line of a paragraph lies. Nothing is painted for it: an editor
 * finds the place of a caret and the line under a pointer with it, also where
 * a line holds no text. Only made with DocLayoutOptions.lineMarks.
 */
export interface LineMark {
  type: 'mark';
  /** Left edge of the content of the line (where its first character is, or would be) and top of the line. */
  x: number;
  y: number;
  /** Extent of the content, trailing spaces left out, and height of the line. */
  width: number;
  height: number;
  /** Baseline, from the top of the line. */
  baseline: number;
  /** Ascent and descent of the font a caret in an empty line has. */
  ascent: number;
  descent: number;
  /** The paragraph (as TextItem.source) and the offset of the first character of the line. */
  source: { story: string; path: number[]; offset: number };
  /** Offsets the line covers; the line that ends the paragraph also takes a caret after them. */
  length: number;
  last: boolean;
  /**
   * Pictures and drawings in the line: where they are (from x of the mark) and
   * the offset of the paragraph they sit at. They take no offset of their own:
   * a caret stands before or after them at that offset.
   */
  boxes?: { dx: number; width: number; offset: number }[];
  /** As for text items: a copy of a header row, and a line of text turned in a table cell. */
  repeat?: true;
  rotation?: number;
}

export type PageItem = TextItem | RectItem | LineItem | ImageItem | ShapeItem | LineMark;

/** One laid out page. Items are in painting order: what comes later is painted over what came before. */
export interface LayoutPage {
  /** Size of the paper in px. */
  width: number;
  height: number;
  /** Index of the section the page belongs to. */
  section: number;
  /** The page number as the document shows it (PAGE field), after restarts. */
  number: number;
  items: PageItem[];
}

export interface DocLayoutOptions {
  /** Font used when the document names none that the measurer knows. */
  fallbackFamily?: string;
  /** Pages laid out before `layout` stops with `done: false`; no limit when missing. */
  maxPages?: number;
  /**
   * What laying out the text of text boxes may cost in one document, in characters (a paragraph
   * or a table row counts as 16 more); 4,000,000 when missing. Boxes that come after the limit
   * is reached are drawn empty. A host with little memory may lower it.
   */
  maxTextBoxWork?: number;
  /** Adds a LineMark for every line of every paragraph to the items of the pages. */
  lineMarks?: boolean;
}

/** How much one call of DocLayout.layout() may do; without a budget it runs to the end. */
export interface LayoutBudget {
  /** Stop after this many more pages. */
  pages?: number;
  /** Stop when this much time has passed. */
  milliseconds?: number;
}

export interface LayoutProgress {
  /** The whole document is laid out. */
  done: boolean;
  /**
   * Indexes of pages handed out by an earlier call whose items changed in
   * this one: the pages that show the total number of pages, once it is known.
   */
  updated: number[];
}
