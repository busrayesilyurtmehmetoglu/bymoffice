// What the parts of the layout engine share: the document, the measurer with
// its caches, effective formatting (resolved once per property bag), and the
// "band", the unit pagination works with.

import {
  resolveCellProps,
  resolveColor,
  resolveParaProps,
  resolveRunProps,
  resolveTableProps,
  tableConditions,
  type TableContext,
} from '../../formats/docx/resolve';
import { NumberingCounters } from '../../formats/docx/numbering';
import type { Block, Border, CellProps, DocColor, DocModel, DrawingAnchor, ParaProps, RunProps, Shading, Table, TableProps } from '../../model/doc/types';
import type { ParaEnv } from './paragraph';
import type { ParaPrep } from './prep';
import { HIGHLIGHT, symbolFont, type LiveField } from './symbols';
import type { DocLayoutOptions, FontMetrics, LayoutFont, PageItem, TextItem, TextMeasurer } from './types';
import { eighths, finite, halfPoints, points, twips } from './units';

// --------------------------------------------------------------------- bands

/** A drawing that is placed relative to the page, a margin, a column or its paragraph instead of in a line. */
export interface FloatSpec {
  width: number;
  height: number;
  /** What the drawing shows, relative to its top left corner. */
  items: PageItem[];
  anchor: DrawingAnchor;
  /** Top left corner of the text area of the anchoring paragraph, relative to the origin of the band that carries the float. */
  dx: number;
  dy: number;
  /** Width of that text area (the column, or the cell the paragraph is in). */
  columnWidth: number;
  /** The paragraph is in a table cell: the float does not push text around. */
  inCell: boolean;
}

/** Where a band of a paragraph came from, so that the rest of the paragraph can be laid out again from there. */
export interface ParaRef {
  prep: ParaPrep;
  /** First atom and character of the line. */
  atom: number;
  ch: number;
  /** Width the lines were broken for. */
  width: number;
  /**
   * Set when the lines were broken around drawings: page and column the band
   * was meant for and the y it was meant to start at. Elsewhere the band does
   * not fit its surroundings and the paragraph is broken again from here.
   */
  slot?: { page: number; column: number; y: number };
  /** What the band generator was given, to do the same again. */
  env: ParaEnv;
}

export interface RowRef {
  /** Shared by the bands of one table. */
  table: { headers: Band[] };
  header: boolean;
  /** Lines that close the table under this row when the next row goes to another page. */
  close?: PageItem[];
}

/**
 * A horizontal strip of laid out content that pagination does not look into:
 * a line of a paragraph, a table row. Items are relative to the left edge of
 * the column and the top of the strip (below `before`).
 */
export interface Band {
  height: number;
  /** Space above (paragraph spacing); dropped at the top of a column after an automatic break. */
  before: number;
  /** Space below; may hang out of the column. */
  after: number;
  items: PageItem[];
  /** No page or column break is wanted between this band and the next. */
  keep: boolean;
  breakBefore?: 'page' | 'column';
  breakAfter?: 'page' | 'column';
  /** Ids of the footnotes referred to in the band. */
  notes?: number[];
  floats?: FloatSpec[];
  para?: ParaRef;
  row?: RowRef;
  /** Splits the band so that the first part is at most `available` high; undefined when it cannot be split there. */
  split?: (available: number) => [Band, Band] | undefined;
}

/** The story content is being laid out for. */
export interface StoryRef {
  /** "body", "header:<rId>", "footnote:<id>", ... */
  id: string;
  /** Text of the number of the note, inside a footnote or endnote. */
  noteLabel?: string;
  /** The body: numbered paragraphs and note references count through the document. */
  main: boolean;
}

/** Values of the fields the engine computes; a missing value shows the result cached in the file. */
export interface FieldValues {
  page?: number;
  /** Number format of the section. */
  pageFormat?: string;
  numPages?: number;
  sectionPages?: number;
}

/** A paragraph in a table cell: where, and a key that is equal for cells formatted alike. */
export interface CellWhere {
  context: TableContext;
  key: string;
}

/** Text item with what the engine needs to fill in a field value later. */
export interface FieldTextItem extends TextItem {
  field?: LiveField;
  fieldFormat?: string;
}

// -------------------------------------------------------------------- styles

export interface BorderLine {
  /** Thickness, px. */
  width: number;
  color: string;
  style: string;
  /** Distance from the text, px. */
  space: number;
}

/** Run formatting as layout uses it. */
export interface RunStyle {
  font: LayoutFont;
  /** Font for East Asian characters when the document names another one for them. */
  eastAsia?: LayoutFont;
  /** Fonts for the lower case letters of small capitals. */
  small?: LayoutFont;
  color: string;
  underline?: { color: string; style: string };
  strike?: boolean | 'double';
  highlight?: string;
  /** Baseline shift, px, positive upwards. */
  shift: number;
  /** Extra advance per character, px. */
  spacing: number;
  /** Horizontal scale, 1 = none. */
  scale: number;
  caps: boolean;
  smallCaps: boolean;
  hidden: boolean;
  /** Text is Turkish or Azeri: i/ı have their own capitals. */
  dotted: boolean;
  /** 1 Symbol, 2 Wingdings, 0 neither. */
  symbol: 0 | 1 | 2;
  superscript: boolean;
}

interface FontEntry {
  font: LayoutFont;
  widths: Map<string, number>;
  metrics?: FontMetrics;
}

/** Entries one font may cache before its cache is emptied. */
const MAX_WIDTHS_PER_FONT = 60000;

export type LayoutBlocksFn = (
  ctx: LayoutContext,
  blocks: Block[],
  width: number,
  story: StoryRef,
  path: number[],
  where: CellWhere | undefined,
  depth: number,
  fields: FieldValues | undefined,
) => Band[];

export class LayoutContext {
  readonly doc: DocModel;
  readonly measurer: TextMeasurer;
  readonly options: DocLayoutOptions;
  /** Default tab distance, px. */
  readonly defaultTab: number;
  /** compatibilityMode of the document (15 = Word 2013 and later); 12 when it states none. */
  readonly compatMode: number;
  /** Counters of the body and of everything else (headers, notes, text boxes). */
  readonly counters: NumberingCounters;
  readonly sideCounters: NumberingCounters;
  /** Number of each footnote/endnote, in the order the body refers to them. */
  readonly footnoteNumbers = new Map<number, number>();
  readonly endnoteNumbers = new Map<number, number>();
  /** Endnote ids in the order of their first reference. */
  readonly endnoteOrder: number[] = [];
  /** Line pitch of the document grid of the section being laid out, px; undefined when lines do not snap to it. */
  grid: number | undefined;
  /** Set by blocks.ts: lays out a list of blocks. A property to break the import cycle. */
  layoutBlocks!: LayoutBlocksFn;

  private fonts = new Map<string, FontEntry>();
  private byFont = new Map<LayoutFont, FontEntry>();
  private runStyles = new WeakMap<RunProps, Map<string, RunStyle>>();
  private paraResolved = new WeakMap<ParaProps, Map<string, ParaProps>>();
  private tableResolved = new WeakMap<Table, TableProps>();
  private families = new Map<string, string>();

  constructor(doc: DocModel, measurer: TextMeasurer, options: DocLayoutOptions) {
    this.doc = doc;
    this.measurer = measurer;
    this.options = options;
    const tab = twips(doc.settings.defaultTabStop);
    this.defaultTab = tab >= 1 ? tab : 48;
    const mode = Number(doc.settings.compatSettings.get('compatibilityMode'));
    this.compatMode = mode >= 11 && mode <= 99 ? mode : 12;
    this.counters = new NumberingCounters(doc.numbering, doc.styles);
    this.sideCounters = new NumberingCounters(doc.numbering, doc.styles);
  }

  /** The family text is measured and painted with: the fallback when the host does not have the one asked for. */
  family(name: string): string {
    let family = this.families.get(name);
    if (family === undefined) {
      family = name;
      const fallback = this.options.fallbackFamily;
      if (fallback !== undefined && this.measurer.hasFont && !this.measurer.hasFont(name)) family = fallback;
      this.families.set(name, family);
    }
    return family;
  }

  /** The one LayoutFont object for these values: fonts are compared by identity everywhere. */
  font(family: string, size: number, bold: boolean, italic: boolean): LayoutFont {
    const name = this.family(family);
    const key = name + '\u0001' + size + (bold ? 'b' : '') + (italic ? 'i' : '');
    let entry = this.fonts.get(key);
    if (!entry) {
      entry = { font: { family: name, size, bold, italic }, widths: new Map() };
      this.fonts.set(key, entry);
      this.byFont.set(entry.font, entry);
    }
    return entry.font;
  }

  private entry(font: LayoutFont): FontEntry {
    let entry = this.byFont.get(font);
    if (!entry) {
      entry = this.byFont.get(this.font(font.family, font.size, font.bold, font.italic))!;
      this.byFont.set(font, entry);
    }
    return entry;
  }

  /** Advance width of the text, cached by font and text. */
  width(font: LayoutFont, text: string): number {
    if (text === '') return 0;
    const entry = this.entry(font);
    let width = entry.widths.get(text);
    if (width === undefined) {
      width = finite(this.measurer.width(font, text));
      if (width < 0) width = 0;
      if (entry.widths.size >= MAX_WIDTHS_PER_FONT) entry.widths.clear();
      entry.widths.set(text, width);
    }
    return width;
  }

  metrics(font: LayoutFont): FontMetrics {
    const entry = this.entry(font);
    if (!entry.metrics) {
      const raw = this.measurer.metrics(font);
      let ascent = finite(raw.ascent);
      let descent = finite(raw.descent);
      const lineGap = Math.max(0, finite(raw.lineGap));
      if (!(ascent + descent > 0) || ascent < 0 || descent < 0) {
        ascent = font.size * 0.8;
        descent = font.size * 0.2;
      }
      entry.metrics = { ascent, descent, lineGap };
    }
    return entry.metrics;
  }

  color(color: DocColor | undefined, fallback: string): string {
    return resolveColor(this.doc, color) || fallback;
  }

  /** Visible colour of a shading, undefined when it shows nothing. */
  shade(shading: Shading | undefined): string | undefined {
    if (!shading) return undefined;
    const pattern = shading.pattern;
    if (pattern === 'nil') return undefined;
    if (pattern === 'solid') return resolveColor(this.doc, shading.color) || '#000000';
    const fill = resolveColor(this.doc, shading.fill);
    const match = pattern === undefined ? null : /^pct(\d{1,2})$/.exec(pattern);
    if (!match) return fill;
    // A percentage pattern: the pattern colour over the fill, mixed.
    const share = Number(match[1]) / 100;
    const front = resolveColor(this.doc, shading.color) || '#000000';
    const back = fill || '#FFFFFF';
    let out = '#';
    for (let i = 1; i < 7; i += 2) {
      const value = Math.round(parseInt(front.slice(i, i + 2), 16) * share + parseInt(back.slice(i, i + 2), 16) * (1 - share));
      out += (value < 16 ? '0' : '') + value.toString(16).toUpperCase();
    }
    return out;
  }

  /** A border as a line, undefined when it draws nothing. */
  border(border: Border | undefined): BorderLine | undefined {
    if (!border || border.style === 'nil' || border.style === 'none') return undefined;
    const width = eighths(border.size === undefined ? 4 : border.size);
    return {
      width: width > 0 ? Math.min(width, 16) : 0.5,
      color: this.color(border.color, '#000000'),
      style: border.style,
      space: points(border.space === undefined ? 0 : Math.min(Math.max(border.space, 0), 31)),
    };
  }

  /** Paragraph properties in effect. The result is shared: it must not be changed. */
  paraProps(props: ParaProps, where: CellWhere | undefined): ParaProps {
    const key = where ? where.key : '';
    let byKey = this.paraResolved.get(props);
    if (!byKey) this.paraResolved.set(props, (byKey = new Map()));
    let resolved = byKey.get(key);
    if (!resolved) byKey.set(key, (resolved = resolveParaProps(this.doc, props, where ? where.context : undefined)));
    return resolved;
  }

  tableProps(table: Table, row?: number): TableProps {
    if (row !== undefined && table.rows[row] && table.rows[row].tableProps) return resolveTableProps(this.doc, table, row);
    let props = this.tableResolved.get(table);
    if (!props) this.tableResolved.set(table, (props = resolveTableProps(this.doc, table)));
    return props;
  }

  /** Where a cell is, with the key that tells cells of different table formatting apart. */
  cellWhere(table: Table, row: number, cell: number): CellWhere {
    const context = { table, row, cell };
    const style = table.props.style;
    // Without a table style no cell of any table differs from another.
    const key = style === undefined && this.doc.styles.defaultIds.table === undefined ? 't' : 't' + (style || '') + '\u0001' + tableConditions(this.doc, context).join(',');
    return { context, key };
  }

  cellProps(where: CellWhere): CellProps {
    return resolveCellProps(this.doc, where.context);
  }

  /** Run formatting in effect for a run with the direct properties `run` in a paragraph with the direct properties `para`. */
  runStyle(para: ParaProps, run: RunProps, where: CellWhere | undefined): RunStyle {
    const key = (para.style === undefined ? '' : para.style) + '\u0001' + (where ? where.key : '');
    let byKey = this.runStyles.get(run);
    if (!byKey) this.runStyles.set(run, (byKey = new Map()));
    let style = byKey.get(key);
    if (!style) byKey.set(key, (style = this.makeStyle(resolveRunProps(this.doc, para, run, where ? where.context : undefined))));
    return style;
  }

  /** Layout formatting from resolved run properties. */
  makeStyle(props: RunProps): RunStyle {
    const fonts = props.fonts || {};
    const complex = props.complexScript === true || props.rtl === true;
    const sizeHalf = (complex ? props.sizeCs : props.size) || 20;
    let size = halfPoints(Math.min(Math.max(sizeHalf, 2), 3276));
    const bold = (complex ? props.boldCs : props.bold) === true;
    const italic = (complex ? props.italicCs : props.italic) === true;
    const family = (complex ? fonts.cs : fonts.ascii) || 'Times New Roman';

    let shift = props.position === undefined ? 0 : halfPoints(props.position);
    const raised = props.vertAlign === 'superscript';
    const lowered = props.vertAlign === 'subscript';
    if (raised || lowered) {
      // Word shows these at about two thirds of the size, a third of the size up or a seventh down.
      shift += raised ? size * 0.33 : -size * 0.14;
      size *= 0.65;
    }
    const font = this.font(family, size, bold, italic);
    const color = this.color(props.color, '#000000');
    const style: RunStyle = {
      font,
      color,
      shift,
      spacing: props.spacing === undefined ? 0 : twips(props.spacing),
      scale: props.scale !== undefined && props.scale > 0 && props.scale !== 100 ? props.scale / 100 : 1,
      caps: props.caps === true,
      smallCaps: props.smallCaps === true && props.caps !== true,
      hidden: props.vanish === true,
      dotted: props.lang !== undefined && props.lang.val !== undefined && /^(tr|az)\b/i.test(props.lang.val),
      symbol: symbolFont(family),
      superscript: raised,
    };
    const eastAsia = fonts.eastAsia;
    if (eastAsia !== undefined && eastAsia !== family && !complex) style.eastAsia = this.font(eastAsia, size, bold, italic);
    if (style.smallCaps) style.small = this.font(family, size * 0.8, bold, italic);
    const underline = props.underline;
    if (underline && underline.style !== undefined && underline.style !== 'none') {
      style.underline = { color: this.color(underline.color, color), style: underline.style };
    }
    if (props.doubleStrike) style.strike = 'double';
    else if (props.strike) style.strike = true;
    if (props.highlight !== undefined && HIGHLIGHT[props.highlight] !== undefined) style.highlight = HIGHLIGHT[props.highlight];
    else {
      const shade = this.shade(props.shading);
      if (shade !== undefined) style.highlight = shade;
    }
    return style;
  }
}

/** The style for a note mark: raised and smaller, unless the run already is. */
const RAISED = new WeakMap<RunStyle, RunStyle>();
export function raisedStyle(ctx: LayoutContext, style: RunStyle): RunStyle {
  if (style.superscript) return style;
  let raised = RAISED.get(style);
  if (!raised) {
    const font = ctx.font(style.font.family, style.font.size * 0.65, style.font.bold, style.font.italic);
    raised = { ...style, font, shift: style.shift + style.font.size * 0.33, superscript: true, eastAsia: undefined, small: undefined };
    RAISED.set(style, raised);
  }
  return raised;
}

/** Total height of bands stacked on each other. */
export function bandsHeight(bands: Band[]): number {
  let height = 0;
  for (let i = 0; i < bands.length; i++) height += bands[i].before + bands[i].height + bands[i].after;
  return height;
}

/** Appends the items of bands stacked from (x, y) downwards; returns the height they take. */
export function stackBands(bands: Band[], out: PageItem[], x: number, y: number): number {
  let top = y;
  for (let i = 0; i < bands.length; i++) {
    const band = bands[i];
    top += band.before;
    pushShifted(out, band.items, x, top);
    top += band.height + band.after;
  }
  return top - y;
}

/** Copy of an item moved by (dx, dy). */
export function shifted(item: PageItem, dx: number, dy: number): PageItem {
  if (item.type === 'line') return { ...item, x1: item.x1 + dx, y1: item.y1 + dy, x2: item.x2 + dx, y2: item.y2 + dy };
  return { ...item, x: item.x + dx, y: item.y + dy };
}

/** Appends the items of a band, moved to where the band is. */
export function pushShifted(out: PageItem[], items: PageItem[], dx: number, dy: number): void {
  for (let i = 0; i < items.length; i++) out.push(shifted(items[i], dx, dy));
}
