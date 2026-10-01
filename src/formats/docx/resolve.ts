// Effective formatting: follows styles, table styles, numbering and the theme
// so that layout gets plain values.
//
// Order of application, lowest priority first (ECMA-376 Part 1, 17.7.2):
//   what Word assumes when nothing is said (IMPLICIT_*) -> document defaults
//   -> table style (whole table, then the conditional formats that apply to
//   the cell) -> numbering level (paragraph properties only) -> paragraph
//   style chain -> character style chain (runs only) -> direct formatting.
// Within a chain a derived style overrides the style it is based on.
//
// The numbering level is the one exception to a fixed order. Its paragraph
// properties (the indents, mostly) belong to whatever brings the numbering:
//   - a style of the chain that states the numbering: the level goes right
//     below that style, so the style's own indents and those of the styles
//     derived from it win, and the level wins over the styles it is based on;
//   - the paragraph itself (<w:numPr> in its own properties): the level goes
//     above the whole style chain, below the rest of the direct formatting.
//
// The toggle run properties (bold, italic, caps, ...) are different: between
// the table, paragraph and character style levels a "true" flips the result
// instead of setting it (17.7.3); direct formatting is absolute.
//
// The functions only work on the model, so readers of other formats that
// produce the same model can use them as well. Results are new objects; the
// shared property bags of the model are never changed.

import {
  TOGGLE_RUN_PROPS,
  type CellMargins,
  type CellProps,
  type DocColor,
  type DocModel,
  type DocStyle,
  type DocStyles,
  type NumLevel,
  type ParaProps,
  type RunFonts,
  type RunProps,
  type StyleType,
  type TabStop,
  type Table,
  type TableBorders,
  type TableLook,
  type TableProps,
  type TableStyleCondition,
} from '../../model/doc/types';
import { abstractOf, levelOf } from './numbering';
import { schemeSlot } from './parts';

/** Longest "based on" chain that is followed. */
export const MAX_STYLE_CHAIN = 64;

/** Where a paragraph sits when it is inside a table: indexes into table.rows and row.cells. */
export interface TableContext {
  table: Table;
  row: number;
  cell: number;
}

const IMPLICIT_FONT = 'Times New Roman';

/**
 * Run formatting Word assumes where neither the document defaults nor any
 * style says otherwise (also when there is no styles part at all): 10 point
 * Times New Roman. This is the one place these are defined.
 */
export const IMPLICIT_RUN_PROPS: RunProps = Object.freeze({
  fonts: Object.freeze({ ascii: IMPLICIT_FONT, hAnsi: IMPLICIT_FONT, eastAsia: IMPLICIT_FONT, cs: IMPLICIT_FONT }),
  size: 20,
});

/** Paragraph formatting Word assumes likewise: left aligned, single line spacing, no space before or after. */
export const IMPLICIT_PARA_PROPS: ParaProps = Object.freeze({
  align: 'left',
  spacing: Object.freeze({ before: 0, after: 0, line: 240, lineRule: 'auto' as const }),
});

/** What Word gives a table that does not say: first row and first column special, row banding. */
const DEFAULT_LOOK: TableLook = {
  firstRow: true,
  lastRow: false,
  firstColumn: true,
  lastColumn: false,
  noHBand: false,
  noVBand: true,
};

/**
 * A style and the styles it is based on, base first. A missing id, an id of
 * another style type or a loop ends the chain. An id that does not name a
 * style of the type falls back to the default style of the type.
 */
export function styleChain(styles: DocStyles, id: string | undefined, type: StyleType): DocStyle[] {
  const chain: DocStyle[] = [];
  let style = id === undefined ? undefined : styles.styles.get(id);
  if (!style || style.type !== type) {
    const fallback = styles.defaultIds[type];
    style = fallback === undefined ? undefined : styles.styles.get(fallback);
  }
  const seen = new Set<DocStyle>();
  while (style && style.type === type && !seen.has(style) && chain.length < MAX_STYLE_CHAIN) {
    seen.add(style);
    chain.push(style);
    style = style.basedOn === undefined ? undefined : styles.styles.get(style.basedOn);
  }
  return chain.reverse();
}

const FONT_SLOTS: [plain: keyof RunFonts, theme: keyof RunFonts][] = [
  ['ascii', 'asciiTheme'],
  ['hAnsi', 'hAnsiTheme'],
  ['eastAsia', 'eastAsiaTheme'],
  ['cs', 'csTheme'],
];

function mergeFonts(base: RunFonts | undefined, over: RunFonts): RunFonts {
  const out: RunFonts = { ...base };
  for (const [plain, theme] of FONT_SLOTS) {
    // A slot is set as a whole: a plain name pushes out an inherited theme font and the other way round.
    if (over[plain] === undefined && over[theme] === undefined) continue;
    delete out[plain];
    delete out[theme];
    if (over[plain] !== undefined) out[plain] = over[plain];
    if (over[theme] !== undefined) out[theme] = over[theme];
  }
  if (over.hint !== undefined) out.hint = over.hint;
  return out;
}

const RUN_SKIP = new Set<string>(['extras', 'change', 'inserted', 'deleted', 'moveFrom', 'moveTo']);

/** Lays `over` on top of `target`, in place. Tracked change data and kept XML are not formatting and stay out. */
export function mergeRunProps(target: RunProps, over: RunProps | undefined): RunProps {
  if (!over) return target;
  const out = target as Record<string, unknown>;
  for (const key of Object.keys(over)) {
    if (RUN_SKIP.has(key)) continue;
    const value = (over as Record<string, unknown>)[key];
    if (value === undefined) continue;
    if (key === 'fonts') target.fonts = mergeFonts(target.fonts, over.fonts as RunFonts);
    else if (key === 'lang') target.lang = { ...target.lang, ...over.lang };
    else if (key === 'underline') target.underline = { ...target.underline, ...over.underline };
    else out[key] = value;
  }
  return target;
}

function mergeTabs(base: TabStop[] | undefined, over: TabStop[]): TabStop[] {
  const byPos = new Map<number, TabStop>();
  if (base) for (const tab of base) byPos.set(tab.pos, tab);
  for (const tab of over) {
    if (tab.align === 'clear') byPos.delete(tab.pos);
    else byPos.set(tab.pos, tab);
  }
  return Array.from(byPos.values()).sort((a, b) => a.pos - b.pos);
}

const PARA_SKIP = new Set<string>(['extras', 'change', 'markProps', 'section', 'style']);

/** Lays `over` on top of `target`, in place. The style id, the mark properties and the section are not inherited. */
export function mergeParaProps(target: ParaProps, over: ParaProps | undefined): ParaProps {
  if (!over) return target;
  const out = target as Record<string, unknown>;
  for (const key of Object.keys(over)) {
    if (PARA_SKIP.has(key)) continue;
    const value = (over as Record<string, unknown>)[key];
    if (value === undefined) continue;
    if (key === 'indent') {
      const indent = { ...target.indent, ...over.indent };
      // First line indent and hanging indent replace each other.
      if (over.indent!.firstLine !== undefined && over.indent!.hanging === undefined) delete indent.hanging;
      if (over.indent!.hanging !== undefined && over.indent!.firstLine === undefined) delete indent.firstLine;
      target.indent = indent;
    } else if (key === 'spacing') target.spacing = { ...target.spacing, ...over.spacing };
    else if (key === 'borders') target.borders = { ...target.borders, ...over.borders };
    else if (key === 'numbering') target.numbering = { ...target.numbering, ...over.numbering };
    else if (key === 'tabs') target.tabs = mergeTabs(target.tabs, over.tabs as TabStop[]);
    else out[key] = value;
  }
  return target;
}

function mergeEdges<T extends TableBorders | CellMargins>(base: T | undefined, over: T | undefined): T | undefined {
  if (!over) return base;
  return { ...base, ...over };
}

function mergeCellProps(target: CellProps, over: CellProps | undefined): CellProps {
  if (!over) return target;
  const out = target as Record<string, unknown>;
  for (const key of Object.keys(over)) {
    if (key === 'extras' || key === 'change') continue;
    const value = (over as Record<string, unknown>)[key];
    if (value === undefined) continue;
    if (key === 'borders') target.borders = mergeEdges(target.borders, over.borders);
    else if (key === 'margins') target.margins = mergeEdges(target.margins, over.margins);
    else out[key] = value;
  }
  return target;
}

function mergeTableProps(target: TableProps, over: TableProps | undefined): TableProps {
  if (!over) return target;
  const out = target as Record<string, unknown>;
  for (const key of Object.keys(over)) {
    if (key === 'extras' || key === 'change' || key === 'style') continue;
    const value = (over as Record<string, unknown>)[key];
    if (value === undefined) continue;
    if (key === 'borders') target.borders = mergeEdges(target.borders, over.borders);
    else if (key === 'cellMargins') target.cellMargins = mergeEdges(target.cellMargins, over.cellMargins);
    else out[key] = value;
  }
  return target;
}

/** Table properties in effect: table style chain, then the table's own, then the row's exceptions. */
export function resolveTableProps(doc: DocModel, table: Table, row?: number): TableProps {
  const out: TableProps = {};
  for (const style of styleChain(doc.styles, table.props.style, 'table')) mergeTableProps(out, style.table);
  mergeTableProps(out, table.props);
  if (row !== undefined && table.rows[row]) mergeTableProps(out, table.rows[row].tableProps);
  if (table.props.style !== undefined) out.style = table.props.style;
  return out;
}

/**
 * Names of the table style conditional formats (<w:tblStylePr> types) that
 * apply to a cell, in the order they are applied (a later one overrides an
 * earlier one): banded columns, banded rows, first column, last column,
 * first row, last row, corner cell.
 *
 * UNSURE: rows after columns, so that in a corner cell the header row wins
 * over the first column, which is how Word is reported to show such a cell.
 * ECMA-376 Part 1, 17.7.6 lists the rows before the columns, which would let
 * the column win. Neither was checked against Word here.
 */
export function tableConditions(doc: DocModel, where: TableContext): string[] {
  const { table } = where;
  const row = table.rows[where.row];
  if (!row || !row.cells[where.cell]) return [];
  const props = resolveTableProps(doc, table, where.row);
  const look = props.look || DEFAULT_LOOK;
  const rowBand = props.rowBandSize === undefined ? 1 : props.rowBandSize;
  const colBand = props.colBandSize === undefined ? 1 : props.colBandSize;

  let column = row.props.gridBefore || 0;
  for (let i = 0; i < where.cell; i++) column += row.cells[i].props.gridSpan || 1;
  const firstRow = look.firstRow && where.row === 0;
  const lastRow = look.lastRow && where.row === table.rows.length - 1 && !firstRow;
  const firstColumn = look.firstColumn && column === 0;
  const lastColumn = look.lastColumn && where.cell === row.cells.length - 1 && !firstColumn;

  const out: string[] = [];
  const bandColumn = column - (look.firstColumn ? 1 : 0);
  if (!look.noVBand && colBand > 0 && bandColumn >= 0) out.push(Math.floor(bandColumn / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert');
  const bandRow = where.row - (look.firstRow ? 1 : 0);
  if (!look.noHBand && rowBand > 0 && bandRow >= 0) out.push(Math.floor(bandRow / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz');
  if (firstColumn) out.push('firstCol');
  if (lastColumn) out.push('lastCol');
  if (firstRow) out.push('firstRow');
  if (lastRow) out.push('lastRow');
  // At most one of these applies: a row is not both first and last here, nor is a column.
  if (firstRow && lastColumn) out.push('neCell');
  if (firstRow && firstColumn) out.push('nwCell');
  if (lastRow && lastColumn) out.push('seCell');
  if (lastRow && firstColumn) out.push('swCell');
  return out;
}

/** What the table style gives a cell: whole-table formatting, then each applicable condition. */
function tableFormatting(doc: DocModel, where: TableContext): { para: ParaProps; run: RunProps; cell: CellProps } {
  const result = { para: {} as ParaProps, run: {} as RunProps, cell: {} as CellProps };
  const chain = styleChain(doc.styles, where.table.props.style, 'table');
  if (!chain.length) return result;
  const apply = (formatting: TableStyleCondition | undefined) => {
    if (!formatting) return;
    mergeParaProps(result.para, formatting.para);
    mergeRunProps(result.run, formatting.run);
    mergeCellProps(result.cell, formatting.cell);
  };
  for (const style of chain) apply(style);
  for (const condition of tableConditions(doc, where)) {
    for (const style of chain) apply(style.conditions && style.conditions.get(condition));
  }
  return result;
}

/** Cell properties in effect: table style (with its conditions), then the cell's own. */
export function resolveCellProps(doc: DocModel, where: TableContext): CellProps {
  const out = tableFormatting(doc, where).cell;
  const row = where.table.rows[where.row];
  return mergeCellProps(out, row && row.cells[where.cell] ? row.cells[where.cell].props : undefined);
}

/** The numbering level a paragraph with this numbering in effect uses, if it is numbered. */
function numberingLevel(doc: DocModel, numbering: NonNullable<ParaProps['numbering']>, chain: DocStyle[]): NumLevel | undefined {
  if (numbering.numId === undefined || numbering.numId === 0) return undefined;
  let level = numbering.level;
  if (level === undefined) {
    // Numbering that comes from a style: the level is the one tied to the style.
    const abstract = abstractOf(doc.numbering, doc.styles, numbering.numId);
    if (abstract) {
      for (const candidate of abstract.levels.values()) {
        if (candidate.paraStyle !== undefined && chain.some((style) => style.id === candidate.paraStyle)) level = candidate.level;
      }
    }
  }
  return levelOf(doc.numbering, doc.styles, numbering.numId, level === undefined ? 0 : level);
}

/**
 * Paragraph properties in effect. `props` are the paragraph's own; `where`
 * is given for a paragraph inside a table cell.
 */
export function resolveParaProps(doc: DocModel, props: ParaProps, where?: TableContext): ParaProps {
  const out: ParaProps = {};
  mergeParaProps(out, IMPLICIT_PARA_PROPS);
  mergeParaProps(out, doc.styles.defaults.para);
  if (where) mergeParaProps(out, tableFormatting(doc, where).para);
  const chain = styleChain(doc.styles, props.style, 'paragraph');

  // The numbering in effect, and what brings it: the style of the chain that
  // states which numbering (index into the chain), or the paragraph itself
  // (chain.length). UNSURE: a paragraph that states only the level of a
  // numbering its style brings counts as bringing the numbering itself.
  const numbering: NonNullable<ParaProps['numbering']> = {};
  let source = -1;
  for (let i = 0; i < chain.length; i++) {
    const stated = chain[i].para && chain[i].para!.numbering;
    if (!stated) continue;
    if (stated.level !== undefined) numbering.level = stated.level;
    if (stated.numId !== undefined) {
      numbering.numId = stated.numId;
      source = i;
    }
  }
  if (props.numbering && (props.numbering.numId !== undefined || props.numbering.level !== undefined)) {
    if (props.numbering.level !== undefined) numbering.level = props.numbering.level;
    if (props.numbering.numId !== undefined) numbering.numId = props.numbering.numId;
    source = chain.length;
  }
  const level = source < 0 ? undefined : numberingLevel(doc, numbering, chain);
  const applyLevel = () => {
    if (!level || !level.para) return;
    const kept = out.numbering;
    mergeParaProps(out, level.para);
    if (kept === undefined) delete out.numbering;
    else out.numbering = kept;
  };

  for (let i = 0; i < chain.length; i++) {
    if (i === source) applyLevel();
    mergeParaProps(out, chain[i].para);
  }
  if (source === chain.length) applyLevel();
  mergeParaProps(out, props);
  if (chain.length) out.style = chain[chain.length - 1].id;
  if (props.markProps) out.markProps = props.markProps;
  if (props.section) out.section = props.section;
  return out;
}

/** Run properties as stated by the levels below, toggles worked out, not yet made concrete. */
function mergedRunProps(doc: DocModel, para: ParaProps, run: RunProps, where?: TableContext): RunProps {
  const levels: RunProps[] = [];
  if (where) levels.push(tableFormatting(doc, where).run);
  const paraLevel: RunProps = {};
  for (const style of styleChain(doc.styles, para.style, 'paragraph')) mergeRunProps(paraLevel, style.run);
  levels.push(paraLevel);
  const charLevel: RunProps = {};
  for (const style of styleChain(doc.styles, run.style, 'character')) mergeRunProps(charLevel, style.run);
  levels.push(charLevel);

  const out: RunProps = {};
  mergeRunProps(out, IMPLICIT_RUN_PROPS);
  mergeRunProps(out, doc.styles.defaults.run);
  for (const level of levels) mergeRunProps(out, level);
  mergeRunProps(out, run);

  for (const key of TOGGLE_RUN_PROPS) {
    if (run[key] !== undefined) continue;
    let stated = false;
    let value = false;
    for (const level of levels) {
      if (level[key] === undefined) continue;
      stated = true;
      if (level[key]) value = !value;
    }
    if (stated) out[key] = value;
  }
  return out;
}

/**
 * Fills in what layout must not have to guess: every font slot gets a font
 * name (the theme font where one is referred to and the theme names it, the
 * stated name otherwise, the implied font failing both), and the complex
 * script size is the size where none was stated.
 *
 * UNSURE: Word itself falls back to 10 points for the complex script size, not
 * to the size; a document that states one and not the other is rare, and text
 * in two sizes would be the more surprising result.
 */
function concrete(doc: DocModel, out: RunProps): RunProps {
  const fonts: RunFonts = { ...out.fonts };
  for (const [plain] of FONT_SLOTS) {
    fonts[plain] = resolveFont(doc, out.fonts, plain as 'ascii' | 'hAnsi' | 'eastAsia' | 'cs') || IMPLICIT_FONT;
  }
  out.fonts = fonts;
  if (out.sizeCs === undefined) out.sizeCs = out.size;
  return out;
}

/**
 * Run properties in effect for a run with the direct properties `run` in a
 * paragraph with the direct properties `para`. The result always states a
 * size and a font name for each of the four font slots.
 */
export function resolveRunProps(doc: DocModel, para: ParaProps, run: RunProps, where?: TableContext): RunProps {
  return concrete(doc, mergedRunProps(doc, para, run, where));
}

/**
 * Run properties in effect for the number or bullet of a numbered paragraph
 * with the direct properties `para`: those of the paragraph mark, resolved
 * like a run of the paragraph, with the run properties of the numbering level
 * laid on top.
 */
export function resolveNumberRunProps(doc: DocModel, para: ParaProps, level: NumLevel, where?: TableContext): RunProps {
  const out = mergedRunProps(doc, para, para.markProps || {}, where);
  return concrete(doc, mergeRunProps(out, level.run));
}

function hueToChannel(p: number, q: number, t: number): number {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

/** A channel value in 0..1 as two hex digits; the fraction is cut, not rounded. */
const hex2 = (channel: number) => {
  const n = Math.max(0, Math.min(255, Math.floor(channel * 255 + 1e-9)));
  return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
};

/**
 * Scales the luminance of "RRGGBB" in HSL, hue and saturation kept, as
 * ECMA-376 Part 1 describes it (17.3.2.41, 17.3.2.40): towards white by `tint`,
 * then towards black by `shade` (both 0..255, 255 = unchanged).
 *
 * UNSURE: the arithmetic Word uses is not published, and the values it writes
 * next to its theme references are not the ones exact arithmetic gives: now
 * and then one channel is a step lower. The two roundings below (the hue to
 * fifths of a degree, the scaled luminance cut to thousandths) are a fit, not
 * Word's algorithm: of the variants tried they reproduce the most reference
 * values, namely every value pinned in the tests, with the rest of the palette
 * listed there within one step of one channel. A grey has no hue and is
 * scaled exactly.
 */
function tintShade(rgb: string, tint: number | undefined, shade: number | undefined): string {
  const r = parseInt(rgb.slice(0, 2), 16) / 255;
  const g = parseInt(rgb.slice(2, 4), 16) / 255;
  const b = parseInt(rgb.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  if (max === min) {
    if (tint !== undefined) l = (l * tint) / 255 + (1 - tint / 255);
    if (shade !== undefined) l = (l * shade) / 255;
    return hex2(l) + hex2(l) + hex2(l);
  }
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h: number;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  h = Math.floor((h / 6) * 1800 + 0.5) / 1800;
  const cut = (value: number) => Math.floor(value * 1000 + 1e-9) / 1000;
  if (tint !== undefined) l = cut((l * tint) / 255 + (1 - tint / 255));
  if (shade !== undefined) l = cut((l * shade) / 255);
  const q = l <= 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return hex2(hueToChannel(p, q, h + 1 / 3)) + hex2(hueToChannel(p, q, h)) + hex2(hueToChannel(p, q, h - 1 / 3));
}

/** Largest difference between the channels of two "RRGGBB" colours. */
function colorDistance(a: string, b: string): number {
  let distance = 0;
  for (let i = 0; i < 6; i += 2) {
    distance = Math.max(distance, Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)));
  }
  return distance;
}

/** How far the value in the file may be from the computed colour and still count as the same colour. */
const SAME_COLOR = 2;

/**
 * Resolves a colour to "#RRGGBB"; undefined when it is automatic or cannot be
 * resolved. A theme colour is computed from the theme; without a usable
 * theme the value the file carries next to it is used.
 *
 * With a tint or shade, the value in the file is what the writing application
 * got for the same reference. Applications round differently (see tintShade),
 * so when that value is within two steps of what the reader computes it is
 * taken to be the same colour and is returned as written: the document then
 * shows exactly the colour it was saved with. A value further away is stale
 * (the theme was changed since) or a placeholder, and the computed colour is
 * returned.
 */
export function resolveColor(doc: DocModel, color: DocColor | undefined): string | undefined {
  if (!color) return undefined;
  if (color.theme !== undefined) {
    const slot = schemeSlot(doc.settings, color.theme);
    const base = slot === undefined ? undefined : doc.theme.colors.get(slot);
    if (base !== undefined) {
      const tint = color.tint === 255 ? undefined : color.tint;
      const shade = color.shade === 255 ? undefined : color.shade;
      if (tint === undefined && shade === undefined) return '#' + base;
      const computed = tintShade(base, tint, shade);
      return '#' + (color.rgb !== undefined && colorDistance(computed, color.rgb) <= SAME_COLOR ? color.rgb : computed);
    }
  }
  return color.rgb === undefined ? undefined : '#' + color.rgb;
}

/**
 * Font name for one of the four font slots of a run, with theme fonts
 * followed ("minorHAnsi" -> the Latin font of the minor theme font).
 */
export function resolveFont(doc: DocModel, fonts: RunFonts | undefined, slot: 'ascii' | 'hAnsi' | 'eastAsia' | 'cs'): string | undefined {
  if (!fonts) return undefined;
  const themeRef = fonts[(slot + 'Theme') as keyof RunFonts];
  if (themeRef !== undefined) {
    const font = themeRef.startsWith('major') ? doc.theme.majorFont : doc.theme.minorFont;
    const name = themeRef.endsWith('EastAsia') ? font.eastAsia : themeRef.endsWith('Bidi') ? font.complexScript : font.latin;
    if (name) return name;
  }
  return fonts[slot];
}
