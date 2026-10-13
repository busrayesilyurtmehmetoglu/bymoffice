// Document model shared by every word processing format (docx now, doc and rtf later).
//
// The model states properties the way the file does: nothing is inherited or
// resolved here (see the resolve helpers of the reader for effective values).
//
// Units are the ones WordprocessingML uses, so that values go back to the file
// unchanged:
//   twips (1/20 pt)     lengths: indents, spacing, page size, widths
//   half-points         font size, raised/lowered position
//   eighths of a point  border width
//   EMU (1/914400 in)   drawings
//
// Lossless principle. Every element the reader does not understand is kept as
// an XML tree in the model:
//   - property bags (ParaProps, RunProps, ...) list such children in `extras`,
//     each with the position it had among its siblings. A child that was only
//     partly understood (an unknown attribute, an out-of-range value) is listed
//     too, marked `partial`: the fields hold what was read, the tree holds all.
//   - unknown content is an item of type "unknown" in its place in the stream.
//   - attributes of content elements (rsids, w14:paraId, ...) are kept in `attrs`.
//   - comments and processing instructions are items of type "misc" in their
//     place in the stream, or entries of a `misc` list where there is no
//     stream (between the rows of a table, the cells of a row).
// A story keeps the text of its part once, in `source`. Every block read
// directly from a story (not nested in a table or wrapper) says where in that
// text it came from (`srcStart`, `srcEnd`; see rawOf), so that a writer can
// copy an untouched block byte for byte, the way the xlsx writer copies
// untouched rows.
//
// Property bags are shared between the items that use the same properties and
// must never be changed in place; an edit replaces the reference.

import type { XmlComment, XmlElement, XmlPI } from '../../xml';
import type { Chart } from '../sheet/chart';

/** Attributes in document order, names as written (with prefix). */
export type AttrList = [name: string, value: string][];

/** A child element kept as read. */
export interface ExtraXml {
  /**
   * Position among the element children of the parent it was read from; -1
   * when `el` is the parent itself, kept for the attributes or comments it
   * carries. With `wrap`, the position among the content children of that
   * wrapper instead.
   */
  index: number;
  el: XmlElement;
  /** The model holds part of this element in its own fields; a writer starts from `el` and overrides those. */
  partial?: boolean;
  /** Rows and tables only: the row- or cell-level wrapper the element sat in. */
  wrap?: Wrapper;
}

/** A comment or processing instruction kept where there is no stream of items to put it in. */
export interface MiscXml {
  /** Number of element children before it in the parent it was read from; with `wrap`, of content children of that wrapper. */
  index: number;
  node: XmlComment | XmlPI;
  wrap?: Wrapper;
}

/** Where in the source text of its story (Story.source) something was read from. */
export interface SourceSpan {
  srcStart?: number;
  /** Offset just past the last character. */
  srcEnd?: number;
}

// ---------------------------------------------------------------- properties

export interface DocColor {
  /** "RRGGBB", upper case. With a theme colour this is the value the writing application computed. */
  rgb?: string;
  auto?: boolean;
  /** Theme colour name as in ST_ThemeColor: "accent1", "text1", "background2", "hyperlink", ... */
  theme?: string;
  /** 0..255; lightens the theme colour. */
  tint?: number;
  /** 0..255; darkens the theme colour. */
  shade?: number;
}

export interface Border {
  /** "single", "double", "dashed", "nil", "none", ... as in ST_Border. */
  style: string;
  /** Width in eighths of a point. */
  size?: number;
  /** Distance from the text in points. */
  space?: number;
  color?: DocColor;
  shadow?: boolean;
  frame?: boolean;
  /** A left or right edge that the file names by reading direction (<w:start>, <w:end>). */
  logical?: true;
}

export interface Shading {
  /** "clear", "solid", "pct10", ... as in ST_Shd; absent when the file states none. */
  pattern?: string;
  /** Background colour; the visible colour of a "clear" shading. */
  fill?: DocColor;
  /** Pattern colour. */
  color?: DocColor;
}

export interface RunFonts {
  ascii?: string;
  hAnsi?: string;
  eastAsia?: string;
  cs?: string;
  /** Theme font references ("majorHAnsi", "minorBidi", ...); each wins over the plain name of its slot. */
  asciiTheme?: string;
  hAnsiTheme?: string;
  eastAsiaTheme?: string;
  csTheme?: string;
  hint?: string;
}

export interface Language {
  val?: string;
  eastAsia?: string;
  bidi?: string;
}

/** Who made a tracked change and when. */
export interface RevisionInfo {
  id?: number;
  author?: string;
  /** ISO 8601 text as written. */
  date?: string;
  /** Every other attribute the element of the revision carries (w16du:dateUtc, an id that is no number): part of it, written back with it. */
  attrs?: AttrList;
}

/** Run properties that toggle between style levels instead of overriding (ECMA-376 Part 1, 17.7.3). */
export const TOGGLE_RUN_PROPS = [
  'bold',
  'boldCs',
  'italic',
  'italicCs',
  'caps',
  'smallCaps',
  'strike',
  'emboss',
  'imprint',
  'outline',
  'shadow',
  'vanish',
] as const;

export interface RunProps {
  /** Character style id. */
  style?: string;
  fonts?: RunFonts;
  /** Half-points. */
  size?: number;
  sizeCs?: number;
  bold?: boolean;
  boldCs?: boolean;
  italic?: boolean;
  italicCs?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  strike?: boolean;
  doubleStrike?: boolean;
  outline?: boolean;
  shadow?: boolean;
  emboss?: boolean;
  imprint?: boolean;
  /** Hidden text. */
  vanish?: boolean;
  webHidden?: boolean;
  specVanish?: boolean;
  noProof?: boolean;
  snapToGrid?: boolean;
  rtl?: boolean;
  /** Complex script formatting applies (use the "...Cs" properties). */
  complexScript?: boolean;
  /** `style` is "single", "double", "none", ... as in ST_Underline. */
  underline?: { style?: string; color?: DocColor };
  color?: DocColor;
  /** Highlight colour name ("yellow", "darkBlue", "none", ...). */
  highlight?: string;
  shading?: Shading;
  border?: Border;
  vertAlign?: 'baseline' | 'superscript' | 'subscript';
  /** Raised (positive) or lowered text, half-points. */
  position?: number;
  /** Character spacing adjustment, twips. */
  spacing?: number;
  /** Horizontal scale, percent. */
  scale?: number;
  /** Font size from which kerning applies, half-points. */
  kerning?: number;
  lang?: Language;
  /** Emphasis mark. */
  emphasis?: string;
  effect?: string;
  /** On a paragraph mark or a table row: the mark itself is a tracked insertion/deletion/move. */
  inserted?: RevisionInfo;
  deleted?: RevisionInfo;
  moveFrom?: RevisionInfo;
  moveTo?: RevisionInfo;
  /** <w:rPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

export interface TabStop {
  /** Twips from the leading edge. */
  pos: number;
  /** "left", "center", "right", "decimal", "bar", "num", "clear", "start", "end". */
  align: string;
  leader?: string;
}

export interface Indent {
  /** Leading edge (w:left or w:start). */
  left?: number;
  /** Trailing edge (w:right or w:end). */
  right?: number;
  firstLine?: number;
  hanging?: number;
  /** The same in hundredths of a character unit; these win over the twips values when present. */
  leftChars?: number;
  rightChars?: number;
  firstLineChars?: number;
  hangingChars?: number;
  /** The file names the edges by reading direction (w:start, w:end, w:startChars, w:endChars). */
  logical?: true;
}

export interface ParaSpacing {
  before?: number;
  after?: number;
  /** 240ths of a line when lineRule is "auto" (or absent), twips otherwise. */
  line?: number;
  lineRule?: 'auto' | 'exact' | 'atLeast';
  beforeAuto?: boolean;
  afterAuto?: boolean;
  /** Hundredths of a line. */
  beforeLines?: number;
  afterLines?: number;
}

export interface ParaBorders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  between?: Border;
  bar?: Border;
}

/** <w:framePr>: a text frame, or a dropped capital. */
export interface FrameProps {
  /** "drop" (in the text area) or "margin"; "none" for an ordinary frame. */
  dropCap?: string;
  /** Height of a dropped capital in lines. */
  lines?: number;
  /** Twips. */
  width?: number;
  height?: number;
  heightRule?: 'auto' | 'exact' | 'atLeast';
  /** Distance from the text around it, twips. */
  hSpace?: number;
  vSpace?: number;
  /** "around", "notBeside", "none", "tight", "through", "auto". */
  wrap?: string;
  /** "text", "margin", "page". */
  hAnchor?: string;
  vAnchor?: string;
  /** Position relative to the anchor, twips; the alignments win over these. */
  x?: number;
  y?: number;
  /** "left", "center", "right", "inside", "outside". */
  xAlign?: string;
  /** "inline", "top", "center", "bottom", "inside", "outside". */
  yAlign?: string;
  anchorLock?: boolean;
}

export interface ParaProps {
  /** Paragraph style id. */
  style?: string;
  /** "left", "center", "right", "both", "distribute", "start", "end", ... as in ST_Jc. */
  align?: string;
  indent?: Indent;
  spacing?: ParaSpacing;
  contextualSpacing?: boolean;
  tabs?: TabStop[];
  /** numId 0 means "no numbering", which removes a numbering inherited from the style. */
  numbering?: { numId?: number; level?: number };
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  pageBreakBefore?: boolean;
  suppressLineNumbers?: boolean;
  suppressAutoHyphens?: boolean;
  bidi?: boolean;
  mirrorIndents?: boolean;
  snapToGrid?: boolean;
  wordWrap?: boolean;
  adjustRightInd?: boolean;
  autoSpaceDE?: boolean;
  autoSpaceDN?: boolean;
  kinsoku?: boolean;
  overflowPunct?: boolean;
  topLinePunct?: boolean;
  borders?: ParaBorders;
  shading?: Shading;
  /** 0..8 for heading levels 1..9; 9 is body text. */
  outlineLevel?: number;
  textAlignment?: string;
  textDirection?: string;
  /** Table conditional formatting flags of the paragraph, the 12 digits of w:cnfStyle/@w:val. */
  cnfStyle?: string;
  /** The paragraph is (part of) a text frame or holds a dropped capital. */
  frame?: FrameProps;
  /** Properties of the paragraph mark (<w:rPr> inside <w:pPr>). */
  markProps?: RunProps;
  /** Section that ends with this paragraph. */
  section?: SectionProps;
  /** <w:pPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

/** A table measurement (CT_TblWidth). */
export interface TableWidth {
  /** "dxa": twips; "pct": fiftieths of a percent; "auto" and "nil" mean no width, whatever the value. */
  type: 'dxa' | 'pct' | 'auto' | 'nil';
  /** As written (0 when the file gives none), so also for "auto" and "nil". */
  value: number;
  /** Cell margins only: a left or right edge that the file names by reading direction (<w:start>, <w:end>). */
  logical?: true;
}

export interface TableBorders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  insideH?: Border;
  insideV?: Border;
  /** Cells only: diagonals. */
  tl2br?: Border;
  tr2bl?: Border;
}

export interface CellMargins {
  top?: TableWidth;
  left?: TableWidth;
  bottom?: TableWidth;
  right?: TableWidth;
}

export interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstColumn: boolean;
  lastColumn: boolean;
  noHBand: boolean;
  noVBand: boolean;
  /** w:val as written, a bit field (bits other than those of the six flags included); absent when the file has none. */
  val?: number;
  /** The flags the file states with an attribute of their own, as a bit field with the bits of w:val; absent when none. */
  stated?: number;
}

/** Bits of TableLook.val and TableLook.stated. */
export const TABLE_LOOK_BITS = Object.freeze({
  firstRow: 0x0020,
  lastRow: 0x0040,
  firstColumn: 0x0080,
  lastColumn: 0x0100,
  noHBand: 0x0200,
  noVBand: 0x0400,
});

/** <w:tblpPr>: where a floating table sits and how far the text around it keeps away. All lengths in twips. */
export interface TablePosition {
  leftFromText?: number;
  rightFromText?: number;
  topFromText?: number;
  bottomFromText?: number;
  /** "text", "margin", "page". */
  vertAnchor?: string;
  horzAnchor?: string;
  /** Position relative to the anchor; the alignments win over these. */
  x?: number;
  y?: number;
  /** "left", "center", "right", "inside", "outside". */
  xAlign?: string;
  /** "inline", "top", "center", "bottom", "inside", "outside". */
  yAlign?: string;
}

export interface TableProps {
  /** Table style id. */
  style?: string;
  width?: TableWidth;
  /** "left", "center", "right", "start", "end". */
  align?: string;
  indent?: TableWidth;
  layout?: 'fixed' | 'autofit';
  borders?: TableBorders;
  shading?: Shading;
  cellMargins?: CellMargins;
  cellSpacing?: TableWidth;
  look?: TableLook;
  /** Rows/columns per band of a banded table style. */
  rowBandSize?: number;
  colBandSize?: number;
  bidiVisual?: boolean;
  /** Present for a floating table. */
  floating?: TablePosition;
  /** "never" or "overlap": may this floating table overlap another one. */
  overlap?: string;
  caption?: string;
  description?: string;
  /** <w:tblPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

export interface RowProps {
  /** Twips. */
  height?: number;
  heightRule?: 'auto' | 'exact' | 'atLeast';
  /** Repeats at the top of each page. */
  header?: boolean;
  cantSplit?: boolean;
  align?: string;
  /** Grid columns skipped before the first and after the last cell. */
  gridBefore?: number;
  gridAfter?: number;
  widthBefore?: TableWidth;
  widthAfter?: TableWidth;
  cellSpacing?: TableWidth;
  hidden?: boolean;
  cnfStyle?: string;
  inserted?: RevisionInfo;
  deleted?: RevisionInfo;
  /** <w:trPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

export interface CellProps {
  width?: TableWidth;
  gridSpan?: number;
  /** "restart" starts a vertically merged cell, "continue" continues the one above. */
  vMerge?: 'restart' | 'continue';
  hMerge?: 'restart' | 'continue';
  borders?: TableBorders;
  shading?: Shading;
  margins?: CellMargins;
  vAlign?: 'top' | 'center' | 'bottom' | 'both';
  textDirection?: string;
  noWrap?: boolean;
  fitText?: boolean;
  hideMark?: boolean;
  cnfStyle?: string;
  /** The cell was put in, or taken out, with its column as a tracked change (<w:cellIns>, <w:cellDel>). */
  inserted?: RevisionInfo;
  deleted?: RevisionInfo;
  /** <w:tcPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

export type HeaderFooterKind = 'default' | 'first' | 'even';

export interface SectionColumns {
  count?: number;
  /** Twips between columns of equal width. */
  space?: number;
  equalWidth?: boolean;
  separator?: boolean;
  /** Individual columns when they are not of equal width. */
  columns?: { width?: number; space?: number }[];
}

/** <w:footnotePr> / <w:endnotePr> of a section. */
export interface NoteProps {
  /** "pageBottom", "beneathText", "sectEnd", "docEnd". */
  position?: string;
  /** Number format of the note marks, as in ST_NumberFormat. */
  format?: string;
  start?: number;
  /** "continuous", "eachSect", "eachPage". */
  restart?: string;
}

export interface PageBorders {
  top?: Border;
  left?: Border;
  bottom?: Border;
  right?: Border;
  /** "page" or "text": what the `space` of the borders is measured from. */
  offsetFrom?: string;
  /** "allPages", "firstPage", "notFirstPage". */
  display?: string;
  /** "front" or "back". */
  zOrder?: string;
}

export interface SectionProps {
  /** How the section starts; "nextPage" when absent. */
  type?: 'nextPage' | 'continuous' | 'evenPage' | 'oddPage' | 'nextColumn';
  page?: { width?: number; height?: number; orientation?: 'portrait' | 'landscape'; code?: number };
  margins?: {
    top?: number;
    right?: number;
    bottom?: number;
    left?: number;
    header?: number;
    footer?: number;
    gutter?: number;
  };
  columns?: SectionColumns;
  /** Relationship ids of the header/footer parts (keys of DocModel.headers / DocModel.footers). */
  headers?: Partial<Record<HeaderFooterKind, string>>;
  footers?: Partial<Record<HeaderFooterKind, string>>;
  /** The first page has its own header and footer. */
  titlePage?: boolean;
  pageNumbering?: { format?: string; start?: number; chapStyle?: number; chapSep?: string };
  vAlign?: 'top' | 'center' | 'both' | 'bottom';
  bidi?: boolean;
  rtlGutter?: boolean;
  textDirection?: string;
  docGrid?: { type?: string; linePitch?: number; charSpace?: number };
  pageBorders?: PageBorders;
  /** Line numbers in the margin: every `countBy`th line, `distance` twips from the text; `restart` is "newPage", "newSection" or "continuous". */
  lineNumbering?: { countBy?: number; start?: number; distance?: number; restart?: string };
  footnotes?: NoteProps;
  endnotes?: NoteProps;
  /** rsid attributes of <w:sectPr>. */
  attrs?: AttrList;
  /** <w:sectPrChange> as read. */
  change?: XmlElement;
  extras?: ExtraXml[];
}

// ------------------------------------------------------------------- content

/**
 * An element that wraps content without being content itself. The wrapped
 * items are listed flat in their stream and point to the wrapper; items that
 * share the same Wrapper object (by identity) and are adjacent belong to the
 * same source element. `parent` is the next wrapper outwards.
 */
export type Wrapper = (
  | {
      kind: 'hyperlink';
      /** Relationship id; `target` is filled in from the relationships of the part. */
      rId?: string;
      target?: string;
      /** Bookmark name inside the document. */
      anchor?: string;
      tooltip?: string;
      attrs?: AttrList;
    }
  | ({ kind: 'ins' | 'del' | 'moveFrom' | 'moveTo'; attrs?: AttrList } & RevisionInfo)
  | {
      kind: 'simpleField';
      /** Field code, e.g. " PAGE \\* MERGEFORMAT ". The wrapped items are the cached result. */
      instruction: string;
      dirty?: boolean;
      locked?: boolean;
      attrs?: AttrList;
    }
  | {
      /** Structured document tag (content control). */
      kind: 'sdt';
      pr?: XmlElement;
      endPr?: XmlElement;
    }
  | {
      kind: 'smartTag' | 'customXml';
      /** <w:smartTagPr> / <w:customXmlPr>. */
      pr?: XmlElement;
      attrs?: AttrList;
    }
  | { kind: 'dir' | 'bdo'; val?: string; attrs?: AttrList }
  | {
      /**
       * <mc:AlternateContent>. The wrapped items come from the selected
       * branch; `xml` is the whole element, which a writer puts back as long
       * as the wrapped items are untouched.
       */
      kind: 'alternate';
      selected: 'choice' | 'fallback' | 'none';
      xml: XmlElement;
    }
) &
  SourceSpan & {
    parent?: Wrapper;
    /** Children of the wrapper element that are neither content nor understood. */
    extras?: ExtraXml[];
    /**
     * Comments and processing instructions directly in a wrapper element whose
     * content sits one level further down (<w:sdt>); those among the content
     * are items of the stream.
     */
    misc?: MiscXml[];
  };

export interface TextItem {
  type: 'text';
  /** Text as it shows: outer whitespace is already removed unless the file asked to preserve it. */
  text: string;
  /** The text as written, when removing the outer whitespace changed it. */
  source?: string;
}

export interface DeletedTextItem {
  /** <w:delText>: text of a tracked deletion. */
  type: 'deletedText';
  text: string;
  source?: string;
}

export interface InstrTextItem {
  /** Field code text of a complex field; `deleted` for <w:delInstrText>. */
  type: 'instrText';
  text: string;
  source?: string;
  deleted?: boolean;
}

/** Field of the complex kind, assembled from the fieldChar/instrText items between "begin" and "end". */
export interface ComplexField {
  /** Field code text. */
  instruction: string;
  /** Plain text of the cached result. */
  result: string;
  /**
   * The story has no "end" for this field. It is taken to end with the
   * paragraph it begins in, so that it cannot hide the rest of the story.
   */
  unclosed?: true;
}

export interface FieldCharItem {
  type: 'fieldChar';
  charType: 'begin' | 'separate' | 'end';
  dirty?: boolean;
  locked?: boolean;
  /** On "begin" items: code and cached result of the field this item opens. */
  field?: ComplexField;
  /** Children as read (<w:ffData>, <w:fldData>, ...). */
  data?: XmlElement[];
  /** The element as read, when it says more than the fields above hold (other attributes, a comment). */
  xml?: XmlElement;
}

/** Position of an anchored drawing along one axis. One of `offset`, `pctOffset` and `align` is set. */
export interface DrawingPosition {
  /** "page", "margin", "column", "paragraph", "character", "line", ... */
  relativeFrom?: string;
  /** EMU. */
  offset?: number;
  /** Thousandths of a percent of the extent `relativeFrom` names. */
  pctOffset?: number;
  align?: string;
}

/** <a:xfrm>: place, size and orientation of a shape, picture or group. Lengths in EMU. */
export interface DrawingTransform {
  /** Offset in the coordinate space of the group around it (the drawing itself for the outermost). */
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  /** Clockwise, 60000ths of a degree. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /** Groups only: the coordinate space its children are placed in, mapped onto the offset and size above. */
  childX?: number;
  childY?: number;
  childWidth?: number;
  childHeight?: number;
}

/** Cropping of an image, thousandths of a percent of each side. */
export interface ImageCrop {
  left?: number;
  top?: number;
  right?: number;
  bottom?: number;
}

/** One picture of a drawing. */
export interface DrawingPicture {
  image: ImageRef;
  crop?: ImageCrop;
  name?: string;
  description?: string;
  /** The image is the fill of a shape, not a picture of its own. */
  fill?: true;
  /** Of the picture (of the shape, for a fill). */
  transform?: DrawingTransform;
  /** The groups the picture sits in, outermost first. */
  groups?: DrawingTransform[];
}

export interface DrawingAnchor {
  horizontal?: DrawingPosition;
  vertical?: DrawingPosition;
  wrap?: 'none' | 'square' | 'tight' | 'through' | 'topAndBottom';
  /** "bothSides", "left", "right", "largest". */
  wrapText?: string;
  /** Distances from the text stated on the wrapping element itself, EMU; they win over the ones below. */
  wrapDist?: { top?: number; bottom?: number; left?: number; right?: number };
  /** Outline the text wraps along (tight and through wrapping), in 21600ths of the size of the drawing. */
  wrapPolygon?: { edited?: boolean; points: [x: number, y: number][] };
  /** Size as thousandths of a percent of the page or margin `relativeFrom` names; 0 means "not relative". */
  sizeRelH?: { relativeFrom?: string; pct?: number };
  sizeRelV?: { relativeFrom?: string; pct?: number };
  behindDoc?: boolean;
  /** Distances from the text, EMU. */
  distT?: number;
  distB?: number;
  distL?: number;
  distR?: number;
  relativeHeight?: number;
  allowOverlap?: boolean;
  layoutInCell?: boolean;
  locked?: boolean;
  /** Absolute position when the anchor uses it instead of horizontal/vertical. */
  simplePos?: { x: number; y: number };
}

export interface ImageRef {
  /** Relationship id of the embedded (r:embed) or linked (r:link) image. */
  rId: string;
  linked?: boolean;
  /** Part name of the image, or the address of an external one; undefined when the relationship is missing. */
  target?: string;
  external?: boolean;
}

/** How the outline of a shape is broken up, by the names DrawingML gives its preset dashes. */
export type ShapeDash = 'dash' | 'dot' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot' | 'sysDashDot' | 'sysDashDotDot';

/**
 * What a shape looks like: its geometry, outline and fill. Rotation and
 * flipping are in DrawingItem.transform; a "line" runs from the top left to
 * the bottom right corner of the drawing before it is flipped.
 */
export interface ShapeLook {
  /** "other" is a geometry the reader does not know: the shape is shown, and written, as the rectangle of its size. */
  geometry: 'rect' | 'roundRect' | 'ellipse' | 'line' | 'other';
  /** For "roundRect": the radius of the corners in 100000ths of the shorter side; 16667 when missing. */
  corner?: number;
  /** "RRGGBB"; a shape without it is not filled. */
  fill?: string;
  /** A shape without it has no outline. `width` in EMU; 0 is the thinnest line a device draws. */
  line?: { color: string; width: number; dash?: ShapeDash };
  /**
   * What the file states of the shape and the look does not hold; each is
   * counted when the shape is written to another format. "fill": a fill that
   * is not a plain colour (a pattern, a texture, a picture, a gradient), which
   * is not shown. "opacity": a fill that is partly transparent, which is
   * shown opaque. "color": a colour that is an index into a palette, a scheme
   * or the colours of the system; such a fill is not shown, such an outline
   * is black.
   */
  lost?: ('fill' | 'opacity' | 'color')[];
}

export interface DrawingItem {
  type: 'drawing';
  placement: 'inline' | 'anchor';
  /** What the graphic holds, from the graphicData uri. */
  kind: 'picture' | 'shape' | 'group' | 'canvas' | 'chart' | 'diagram' | 'other';
  /** Size in EMU. */
  width?: number;
  height?: number;
  name?: string;
  description?: string;
  /** Present for anchored drawings. */
  anchor?: DrawingAnchor;
  /** Present for inline drawings that state them: distances from the text, EMU. */
  inline?: { distT?: number; distB?: number; distL?: number; distR?: number };
  /** Room the effects of the drawing (shadow, glow, rotation) take beyond its size, EMU. */
  effectExtent?: { left?: number; top?: number; right?: number; bottom?: number };
  /** Of the outermost object of the drawing: rotation, flipping, and for a group its child coordinate space. */
  transform?: DrawingTransform;
  /** Every picture in the drawing, in document order; pictures in text boxes belong to the drawings there. */
  images?: DrawingPicture[];
  /**
   * For a drawing of kind "chart" read from a docx: the chart part its frame
   * names ("/word/charts/chart1.xml"), as the package spells it, when the
   * relationship leads to a part of the package. The chart itself is DocModel.charts.get() of it; a part
   * that could not be read has none there.
   */
  chartPart?: string;
  /** The first of `images`, with its cropping. */
  image?: ImageRef;
  crop?: ImageCrop;
  /** Contents of the text boxes inside the drawing. */
  textBoxes?: Block[][];
  /**
   * Geometry, outline and fill of a shape read from a doc or RTF file, where
   * `xml` is only a stand-in. A drawing read from a docx has none: there
   * they are in `xml`.
   */
  shape?: ShapeLook;
  /** The whole <w:drawing>, the source of truth for writing. */
  xml: XmlElement;
}

export interface LegacyGraphicItem {
  /** <w:pict> (VML) or <w:object> (OLE), kept whole. */
  type: 'pict' | 'object';
  image?: ImageRef;
  textBoxes?: Block[][];
  xml: XmlElement;
}

export type RunItem =
  | TextItem
  | DeletedTextItem
  | InstrTextItem
  | { type: 'tab' }
  | { type: 'break'; kind: 'line' | 'page' | 'column'; clear?: string; cr?: boolean }
  | { type: 'noBreakHyphen' }
  | { type: 'softHyphen' }
  | { type: 'symbol'; font?: string; char: string }
  | FieldCharItem
  | { type: 'footnoteReference' | 'endnoteReference'; id: number; customMark?: boolean }
  | { type: 'commentReference'; id: number }
  /** The automatic number inside a note or comment, separators of the notes area, page number, layout hint. */
  | {
      type:
        | 'footnoteRef'
        | 'endnoteRef'
        | 'annotationRef'
        | 'separator'
        | 'continuationSeparator'
        | 'pageNumber'
        | 'lastRenderedPageBreak';
    }
  | { type: 'positionalTab'; alignment?: string; relativeTo?: string; leader?: string }
  | DrawingItem
  | LegacyGraphicItem
  | { type: 'alternate'; selected: 'choice' | 'fallback' | 'none'; items: RunItem[]; xml: XmlElement }
  /** A comment or processing instruction. */
  | { type: 'misc'; node: XmlComment | XmlPI }
  | { type: 'unknown'; xml: XmlElement };

interface InlineBase {
  /** Innermost wrapper (hyperlink, tracked change, content control, ...). */
  wrap?: Wrapper;
}

export interface Run extends InlineBase {
  type: 'run';
  props: RunProps;
  content: RunItem[];
  /**
   * Which items of `content` sit in the code of a complex field (between
   * "begin" and "separate", or "end" when there is no separator) and so are
   * not shown: `true` when all do, otherwise a flag per item. See inFieldCode.
   */
  fieldCode?: true | boolean[];
  attrs?: AttrList;
}

export interface RangeMark extends InlineBase {
  type:
    | 'bookmarkStart'
    | 'bookmarkEnd'
    | 'commentRangeStart'
    | 'commentRangeEnd'
    | 'moveFromRangeStart'
    | 'moveFromRangeEnd'
    | 'moveToRangeStart'
    | 'moveToRangeEnd'
    | 'permStart'
    | 'permEnd';
  /** Ids are numbers except for permissions, which may use any text. */
  id: string;
  /** Bookmark or move name. */
  name?: string;
  /** Every other attribute (author, date, colFirst, ...). */
  attrs?: AttrList;
}

export interface MathItem extends InlineBase {
  /** <m:oMath> or <m:oMathPara>, kept whole. */
  type: 'math';
  /** Plain text of the formula, for search and fallback display. */
  text: string;
  xml: XmlElement;
}

export type Inline =
  | Run
  | RangeMark
  | MathItem
  | ({ type: 'proofError'; kind: string } & InlineBase)
  /** Stands for a wrapper element that has no content. */
  | ({ type: 'empty' } & InlineBase)
  /** A comment or processing instruction. */
  | ({ type: 'misc'; node: XmlComment | XmlPI } & InlineBase)
  | ({ type: 'unknown'; xml: XmlElement } & InlineBase);

/** The source span is only on blocks read directly from a story. */
interface BlockBase extends SourceSpan {
  /** Innermost block-level wrapper (content control, customXml, AlternateContent). */
  wrap?: Wrapper;
}

export interface Paragraph extends BlockBase {
  type: 'paragraph';
  props: ParaProps;
  children: Inline[];
  attrs?: AttrList;
}

export interface TableCell {
  props: CellProps;
  blocks: Block[];
  /** Cell-level wrapper inside the row. */
  wrap?: Wrapper;
  attrs?: AttrList;
}

export interface TableRow {
  props: RowProps;
  /** Table property exceptions of this row (<w:tblPrEx>). */
  tableProps?: TableProps;
  cells: TableCell[];
  /** Row-level wrapper inside the table. */
  wrap?: Wrapper;
  attrs?: AttrList;
  /** Children of <w:tr>, or of a cell-level wrapper in it, that are neither cells nor properties (bookmarks, ...). */
  extras?: ExtraXml[];
  misc?: MiscXml[];
}

export interface Table extends BlockBase {
  type: 'table';
  props: TableProps;
  /** Grid column widths in twips. */
  grid: number[];
  /**
   * The grid as it was before a tracked change of the columns (<w:tblGridChange>, which names nobody: it has an id
   * and nothing else). Only a record that holds plain columns, in a grid of plain columns, is read into this; any other stays in `extras` with its grid.
   */
  gridChange?: { id?: number; grid: number[] };
  rows: TableRow[];
  attrs?: AttrList;
  /** Children of <w:tbl>, or of a row-level wrapper in it, that are not rows or properties; <w:tblGrid> too when it has more than plain columns. */
  extras?: ExtraXml[];
  misc?: MiscXml[];
}

export type Block =
  | Paragraph
  | Table
  /** A range mark between blocks. */
  | ({ type: 'mark'; mark: RangeMark } & BlockBase)
  | ({ type: 'empty' } & BlockBase)
  /** A comment or processing instruction between blocks. */
  | ({ type: 'misc'; node: XmlComment | XmlPI } & BlockBase)
  | ({ type: 'unknown'; xml: XmlElement } & BlockBase);

// --------------------------------------------------------------------- parts

/** Relationship of a part, as the package states it. */
export interface DocRelationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
  /** Resolved part name for internal targets, null for external ones. */
  part: string | null;
}

/** One flow of blocks: the body, a header, a footer. */
export interface Story {
  /** Part the story was read from, e.g. "/word/document.xml". */
  part: string;
  /** Root element of the part: the namespace declarations the source texts and kept trees depend on. */
  rootName: string;
  rootAttrs: AttrList;
  blocks: Block[];
  /** Relationships of the part by id. */
  relationships: Map<string, DocRelationship>;
  /**
   * Text of the part as it was read (decoded, without byte order mark, line
   * ends as "\n"): what the source spans of blocks and wrappers point into.
   */
  source: string;
  /** What the part has before its root element (the XML declaration, comments) and, if anything, after it. */
  prolog?: string;
  epilog?: string;
  /**
   * The story was edited. An edit never changes a block in place: it puts a
   * new block object, without a source span, where the old one was. So the
   * writer copies every block that still has its span from `source` and
   * writes the others from the model; a story that is not modified is not
   * written at all.
   */
  modified?: boolean;
}

export interface Body extends Story {
  /** Properties of the last section (<w:sectPr> at the end of the body). */
  section?: SectionProps;
  /** Where that <w:sectPr> is in the source. */
  sectionSpan?: SourceSpan;
  /** Number of blocks before that <w:sectPr>, when blocks follow it (it is last in a valid document). */
  sectionIndex?: number;
  /** Children of <w:document> other than the body (<w:background>, ...); `index` counts the body too. */
  extras?: ExtraXml[];
  /** Comments and processing instructions directly in <w:document>. */
  misc?: MiscXml[];
}

export interface Note extends SourceSpan {
  id: number;
  /** "separator", "continuationSeparator", "continuationNotice"; absent for ordinary notes. */
  kind?: string;
  blocks: Block[];
  attrs?: AttrList;
}

export interface Comment extends SourceSpan {
  id: number;
  author?: string;
  initials?: string;
  /** ISO 8601 text as written. */
  date?: string;
  blocks: Block[];
  attrs?: AttrList;
  /**
   * What the comment is known by in the parts beside the comments (the
   * replies and "resolved" of commentsExtended.xml, commentsIds.xml): the
   * w14:paraId of its last paragraph, 8 hex digits, upper case. The paragraph
   * of a comment read from a file carries the attribute too (in its `attrs`);
   * one made here does not, and the writer adds it.
   */
  paraId?: string;
  /** The comment answers another one: the id of that comment (w15:paraIdParent of commentsExtended.xml). */
  parent?: number;
  /** The comment is resolved (w15:done of commentsExtended.xml). */
  done?: boolean;
}

/** Footnotes, endnotes or comments: a part that holds many small stories. */
export interface StoryCollection<T> extends Omit<Story, 'blocks'> {
  /** By id, in document order. */
  items: Map<number, T>;
  /**
   * Children of the root that are not items, and items that could not get a
   * place in `items` (no usable id, or an id used before), kept whole.
   */
  extras?: ExtraXml[];
  /** Comments and processing instructions between the items. */
  misc?: MiscXml[];
}

export type StyleType = 'paragraph' | 'character' | 'table' | 'numbering';

/** One <w:tblStylePr>: formatting a table style applies to a part of the table. */
export interface TableStyleCondition {
  para?: ParaProps;
  run?: RunProps;
  table?: TableProps;
  row?: RowProps;
  cell?: CellProps;
}

export interface DocStyle extends TableStyleCondition {
  id: string;
  type: StyleType;
  name?: string;
  basedOn?: string;
  /** Style of the paragraph that follows. */
  next?: string;
  /** Linked character/paragraph style. */
  link?: string;
  isDefault: boolean;
  custom?: boolean;
  hidden?: boolean;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  quickFormat?: boolean;
  locked?: boolean;
  uiPriority?: number;
  /** Table styles: by <w:tblStylePr> type ("firstRow", "band1Horz", "nwCell", ...). */
  conditions?: Map<string, TableStyleCondition>;
  extras?: ExtraXml[];
}

export interface DocStyles {
  /** <w:docDefaults>. */
  defaults: { para: ParaProps; run: RunProps };
  /** By style id. */
  styles: Map<string, DocStyle>;
  /** Id of the default style of each type. */
  defaultIds: Partial<Record<StyleType, string>>;
}

export interface NumLevel {
  /** 0..8. */
  level: number;
  start?: number;
  /** "decimal", "bullet", "lowerRoman", ... as in ST_NumberFormat. */
  format?: string;
  /** Text of the number with "%1".."%9" standing for the counters of levels 0..8. */
  text?: string;
  suffix?: 'tab' | 'space' | 'nothing';
  /** Alignment of the number: "left", "center", "right", "start", "end". */
  align?: string;
  /** 1-based level whose use restarts this one; 0 never restarts; absent restarts after any higher level. */
  restart?: number;
  /** Show the numbers of the higher levels as decimals. */
  legal?: boolean;
  /** Paragraph style tied to this level. */
  paraStyle?: string;
  para?: ParaProps;
  run?: RunProps;
  extras?: ExtraXml[];
}

export interface AbstractNum {
  id: number;
  levels: Map<number, NumLevel>;
  multiLevelType?: string;
  /** Numbering style this definition takes its levels from. */
  numStyleLink?: string;
  /** Numbering style this definition is the body of. */
  styleLink?: string;
}

export interface NumInstance {
  id: number;
  abstractId: number;
  /** By level: a new start value and/or a replacement level definition. */
  overrides: Map<number, { start?: number; level?: NumLevel }>;
}

export interface Numbering {
  abstracts: Map<number, AbstractNum>;
  nums: Map<number, NumInstance>;
  /** An edit added definitions or instances (lists made in the editor): the numbering part has to be written. */
  modified?: boolean;
}

export interface ThemeFont {
  latin?: string;
  eastAsia?: string;
  complexScript?: string;
  /** Typeface by script tag ("Jpan", "Arab", ...). */
  scripts: Map<string, string>;
}

export interface DocTheme {
  /** Colour scheme: "dk1", "lt1", "dk2", "lt2", "accent1".."accent6", "hlink", "folHlink" -> "RRGGBB". */
  colors: Map<string, string>;
  majorFont: ThemeFont;
  minorFont: ThemeFont;
}

export interface DocSettings {
  /** Distance between automatic tab stops, twips. */
  defaultTabStop: number;
  evenAndOddHeaders: boolean;
  mirrorMargins: boolean;
  autoHyphenation: boolean;
  trackRevisions: boolean;
  /** <w:documentProtection>: what the document allows to be edited. Absent when the file states none. */
  protection?: DocProtection;
  /** <w:writeProtection> as read: a recommendation to open the document read-only. Kept, never applied. */
  writeProtection?: XmlElement;
  /** Theme colour name ("bg1", "t1", "accent1", ...) -> scheme slot ("lt1", "dk1", ...). */
  colorMap: Map<string, string>;
  themeFontLang?: Language;
  /** On/off children of <w:compat> by local name. */
  compat: Map<string, boolean>;
  /** <w:compatSetting> values by name ("compatibilityMode" -> "15"). */
  compatSettings: Map<string, string>;
}

/** <w:documentProtection>. The engine neither checks the password nor removes the protection. */
export interface DocProtection {
  /**
   * w:edit as written: "readOnly", "comments", "trackedChanges", "forms",
   * "none", or whatever else a file says; "" when the file gives none.
   */
  edit: string;
  /** w:enforcement: the protection applies. A protection that is not enforced is only kept. */
  enforced: boolean;
  /**
   * The element as read from a docx, with the attributes of the password
   * hash: what a writer puts back. A protection read from a doc or RTF file
   * has none, and is written without a password.
   */
  xml?: XmlElement;
}

export interface FontInfo {
  name: string;
  altName?: string;
  /** 20 hex digits. */
  panose?: string;
  charset?: string;
  /** "roman", "swiss", "modern", "script", "decorative", "auto". */
  family?: string;
  /** "fixed", "variable", "default". */
  pitch?: string;
  /** Relationship ids of embedded font files by face ("regular", "bold", "italic", "boldItalic"). */
  embedded?: Map<string, string>;
}

/** A picture that is held in the model itself, for formats without a package to read it from (doc). */
export interface DocMedia {
  bytes: Uint8Array;
  /** Media type: "image/png", "image/jpeg", "image/bmp", "image/x-emf", "image/x-wmf", "image/tiff", "image/pict". */
  type: string;
}

export interface DocModel {
  body: Body;
  styles: DocStyles;
  numbering: Numbering;
  theme: DocTheme;
  settings: DocSettings;
  /** By font name. */
  fonts: Map<string, FontInfo>;
  /** By the relationship id section properties refer to. Ids that share a part share the Story. */
  headers: Map<string, Story>;
  footers: Map<string, Story>;
  footnotes?: StoryCollection<Note>;
  endnotes?: StoryCollection<Note>;
  comments?: StoryCollection<Comment>;
  /**
   * Pictures by the part name image references give as their target
   * (ImageRef.target). Only set by readers of formats that are not packages:
   * a docx file keeps its pictures in the package.
   */
  media?: Map<string, DocMedia>;
  /**
   * The charts of the document by the name of their part
   * (DrawingItem.chartPart): what the parts state, with the values they
   * cached, as far as they could be read. A view for showing them; the parts
   * stay in the package and are written back from there. Only set by the
   * docx reader, and only when a chart was read.
   */
  charts?: Map<string, Chart>;
  /** The document uses the strict (ISO/IEC 29500) namespaces. */
  strict: boolean;
  /** Things the reader got past: missing parts, relationships without a target, ... */
  warnings: string[];
  /**
   * How often each kind of element was kept without being (fully) understood,
   * keyed "context/name", e.g. "rPr/w14:ligatures". Diagnostic only.
   */
  unknown: Map<string, number>;
}

const PART_NAMESPACES = new WeakMap<object, AttrList>();

/**
 * Remembers the namespace declarations of the root of the part a style sheet
 * or the numbering was read from ("xmlns:w14", ...): what the prefixes in the
 * kept trees of the definitions stand for. Kept beside the model, not in it:
 * it says nothing about the document, only about how one file spelt it.
 */
export function setPartNamespaces(of: DocStyles | Numbering, declarations: AttrList): void {
  PART_NAMESPACES.set(of, declarations);
}

/** The declarations setPartNamespaces() was given; undefined for definitions that were not read from a part. */
export function partNamespaces(of: DocStyles | Numbering): AttrList | undefined {
  return PART_NAMESPACES.get(of);
}

export const EMPTY_RUN_PROPS: RunProps = Object.freeze({});
export const EMPTY_PARA_PROPS: ParaProps = Object.freeze({});

export function emptyDocStyles(): DocStyles {
  return { defaults: { para: EMPTY_PARA_PROPS, run: EMPTY_RUN_PROPS }, styles: new Map(), defaultIds: {} };
}

export function emptyTheme(): DocTheme {
  return { colors: new Map(), majorFont: { scripts: new Map() }, minorFont: { scripts: new Map() } };
}

export function defaultSettings(): DocSettings {
  return {
    defaultTabStop: 720,
    evenAndOddHeaders: false,
    mirrorMargins: false,
    autoHyphenation: false,
    trackRevisions: false,
    // The mapping every document has unless <w:clrSchemeMapping> says otherwise.
    colorMap: new Map([
      ['bg1', 'lt1'],
      ['t1', 'dk1'],
      ['bg2', 'lt2'],
      ['t2', 'dk2'],
    ]),
    compat: new Map(),
    compatSettings: new Map(),
  };
}

/**
 * What an enforced protection of the document restricts editing to; undefined
 * when nothing is enforced, and for a protection that states no kind of
 * editing (one that only restricts formatting). An enforced protection of a
 * kind that is not known counts as "readOnly", the strictest.
 */
export function enforcedProtection(doc: { settings: DocSettings }): 'readOnly' | 'comments' | 'forms' | 'trackedChanges' | undefined {
  const protection = doc.settings ? doc.settings.protection : undefined;
  if (!protection || !protection.enforced) return undefined;
  const edit = protection.edit;
  // Without w:edit the protection restricts formatting only, which restricts nothing here.
  if (edit === 'none' || edit === '') return undefined;
  return edit === 'comments' || edit === 'forms' || edit === 'trackedChanges' ? edit : 'readOnly';
}

/** Does the document ask for its changes to be tracked: by its own setting, or by a protection that allows nothing else. */
export function tracksChanges(doc: { settings: DocSettings }): boolean {
  return (doc.settings !== undefined && doc.settings.trackRevisions === true) || enforcedProtection(doc) === 'trackedChanges';
}

function blocksHavePermissions(blocks: Block[], depth: number): boolean {
  if (depth > 64) return false;
  for (const block of blocks) {
    if (block.type === 'mark') {
      if (block.mark.type === 'permStart') return true;
    } else if (block.type === 'paragraph') {
      for (const inline of block.children) if (inline.type === 'permStart') return true;
    } else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) if (blocksHavePermissions(cell.blocks, depth + 1)) return true;
    }
  }
  return false;
}

/**
 * Does the document mark ranges that its protection leaves open for editing
 * (<w:permStart>). The source text of each part answers for a document read
 * from a docx; the model is searched for the others.
 */
export function hasPermissionRanges(doc: DocModel): boolean {
  const stories: { source: string; blocks?: Block[]; items?: Map<number, { blocks: Block[] }> }[] = [doc.body];
  doc.headers.forEach((story) => stories.push(story));
  doc.footers.forEach((story) => stories.push(story));
  for (const collection of [doc.footnotes, doc.endnotes, doc.comments]) if (collection) stories.push(collection);
  for (const story of stories) {
    if (story.source) {
      if (story.source.indexOf(':permStart') >= 0 || story.source.indexOf('<permStart') >= 0) return true;
      continue;
    }
    if (story.blocks && blocksHavePermissions(story.blocks, 0)) return true;
    if (story.items) for (const item of story.items.values()) if (blocksHavePermissions(item.blocks, 0)) return true;
  }
  return false;
}

/** The wrapper of the given kind around an item, innermost first; undefined when there is none. */
export function findWrapper<K extends Wrapper['kind']>(
  item: { wrap?: Wrapper },
  kind: K,
): Extract<Wrapper, { kind: K }> | undefined {
  // The chain is as deep as the source nesting, which the reader limits.
  for (let w = item.wrap; w; w = w.parent) {
    if (w.kind === kind) return w as Extract<Wrapper, { kind: K }>;
  }
  return undefined;
}

/** Source text of a block, wrapper, note or section mark of the story; undefined when it carries no span. */
export function rawOf(story: { source: string }, item: SourceSpan | undefined): string | undefined {
  if (!item || item.srcStart === undefined || item.srcEnd === undefined) return undefined;
  return story.source.slice(item.srcStart, item.srcEnd);
}

/** True when the item of the run at `index` sits in the code of a complex field and so is not shown. */
export function inFieldCode(run: Run, index: number): boolean {
  const code = run.fieldCode;
  return code === true || (code !== undefined && code[index] === true);
}

function itemsText(items: RunItem[], code: true | boolean[] | undefined): string {
  let out = '';
  for (let i = 0; i < items.length; i++) {
    if (code !== undefined && (code === true || code[i])) continue;
    const item = items[i];
    if (item.type === 'text') out += item.text;
    else if (item.type === 'tab') out += '\t';
    else if (item.type === 'break') out += '\n';
    else if (item.type === 'noBreakHyphen') out += '‑';
    else if (item.type === 'symbol') out += item.char;
    // Only the selected branch is in `items`; the reader limits how deep these nest.
    else if (item.type === 'alternate') out += itemsText(item.items, undefined);
  }
  return out;
}

/**
 * Plain text of a paragraph as it shows: field codes, deleted text and text
 * that was moved away are left out.
 */
export function paragraphText(paragraph: Paragraph): string {
  let out = '';
  for (const inline of paragraph.children) {
    if (inline.type !== 'run' && inline.type !== 'math') continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    if (inline.type === 'math') out += inline.text;
    else if (inline.fieldCode !== true) out += itemsText(inline.content, inline.fieldCode);
  }
  return out;
}

const HYPERLINK_SCHEMES = new Set<string>(['http', 'https', 'mailto', 'tel']);

/**
 * The target of a hyperlink when it is safe to hand to the host as a link,
 * undefined otherwise. Only web addresses (http, https), mailto:, tel: and
 * references into the document itself ("#bookmark") pass; scripts, data,
 * files and everything relative do not. The target is read the way a browser
 * reads it (spaces and control characters around it and tabs and line ends
 * anywhere in it do not count), and it is that reading which is returned.
 */
export function safeHyperlinkTarget(target: string | undefined): string | undefined {
  if (target === undefined) return undefined;
  let start = 0;
  let end = target.length;
  while (start < end && target.charCodeAt(start) <= 0x20) start++;
  while (end > start && target.charCodeAt(end - 1) <= 0x20) end--;
  let url = '';
  let from = start;
  for (let i = start; i < end; i++) {
    const c = target.charCodeAt(i);
    if (c !== 9 && c !== 10 && c !== 13) continue;
    url += target.slice(from, i);
    from = i + 1;
  }
  url += target.slice(from, end);
  if (url.charCodeAt(0) === 35) return url;

  // The scheme: a letter, then letters, digits, "+", "-" and "." up to the first colon.
  let colon = 0;
  for (; colon < url.length; colon++) {
    const c = url.charCodeAt(colon) | 0x20;
    const letter = c >= 97 && c <= 122;
    if (letter) continue;
    const other = url.charCodeAt(colon);
    if (colon > 0 && ((other >= 48 && other <= 57) || other === 43 || other === 45 || other === 46)) continue;
    break;
  }
  // No scheme the reader allows is longer than this, so a long run of letters is not copied.
  if (colon === 0 || colon > 8 || url.charCodeAt(colon) !== 58) return undefined;
  const scheme = url.slice(0, colon).toLowerCase();
  if (!HYPERLINK_SCHEMES.has(scheme)) return undefined;
  const rest = url.length - colon - 1;
  if (scheme === 'http' || scheme === 'https') {
    // An address needs a host: "https:path" and "https:/path" are relative to whatever page shows the link.
    return rest > 2 && url.charCodeAt(colon + 1) === 47 && url.charCodeAt(colon + 2) === 47 ? url : undefined;
  }
  return rest > 0 ? url : undefined;
}
