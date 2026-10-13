// Spreadsheet model shared by every spreadsheet format (xlsx now, xls later).
//
// Rows and columns are 0-based everywhere in the model; "A1" is row 0, col 0.
// The model holds what the engine understands. Everything else stays in the
// source package and is written back from there (see the lossless principle).

import { parseXml, type XmlElement } from '../../xml';
import type { Chart } from './chart';

export const MAX_ROWS = 1048576;
export const MAX_COLS = 16384;

/** A colour as the file states it; resolveColor() turns it into "#RRGGBB". */
export interface Color {
  /** "AARRGGBB" or "RRGGBB", upper case. */
  rgb?: string;
  /** Index into the theme colour list. */
  theme?: number;
  /** -1..1, lightens (positive) or darkens (negative) the base colour. */
  tint?: number;
  /** Index into the legacy indexed palette. */
  indexed?: number;
  auto?: boolean;
}

export type Underline = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';

export interface Font {
  name?: string;
  /** Size in points. */
  size?: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  underline: Underline;
  color?: Color;
  vertAlign?: 'superscript' | 'subscript';
  family?: number;
  charset?: number;
  scheme?: 'major' | 'minor' | 'none';
}

export type Fill =
  | {
      type: 'pattern';
      /** "none", "solid", "gray125", ... as in ST_PatternType. */
      pattern: string;
      /** Foreground colour; this is the visible colour of a solid fill. */
      fg?: Color;
      bg?: Color;
    }
  | {
      type: 'gradient';
      /** The <gradientFill> element as read; not interpreted yet. */
      raw: XmlElement;
    };

export interface BorderEdge {
  /** "thin", "medium", "dashed", ... as in ST_BorderStyle. */
  style: string;
  color?: Color;
}

export interface Border {
  left?: BorderEdge;
  right?: BorderEdge;
  top?: BorderEdge;
  bottom?: BorderEdge;
  diagonal?: BorderEdge;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

export interface Alignment {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
  wrapText?: boolean;
  shrinkToFit?: boolean;
  /** 0..180 degrees, or 255 for stacked vertical text. */
  textRotation?: number;
  indent?: number;
  readingOrder?: number;
}

export interface Protection {
  locked?: boolean;
  hidden?: boolean;
}

/** One <xf> record: indexes into the other style tables. */
export interface CellFormat {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  /** Index into cellStyleXfs; only present on cellXfs entries. */
  xfId?: number;
  alignment?: Alignment;
  protection?: Protection;
  quotePrefix: boolean;
}

export interface NamedCellStyle {
  name: string;
  xfId: number;
  builtinId?: number;
}

export interface StyleSheet {
  /** Custom number formats by id; built-in ids are answered by formatCodeOf(). */
  numFmts: Map<number, string>;
  fonts: Font[];
  fills: Fill[];
  borders: Border[];
  cellStyleXfs: CellFormat[];
  /** Cell.style is an index into this list. */
  cellXfs: CellFormat[];
  cellStyles: NamedCellStyle[];
  /** Differential formats (conditional formatting, tables); kept as read. */
  dxfs: XmlElement[];
  /** The same records as far as the engine draws with them; index for index with `dxfs`. */
  dxfFormats?: DifferentialFormat[];
  /** Custom indexed palette ("RRGGBB") when the file overrides the default one. */
  indexedColors?: string[];
  /** Theme colours as "RRGGBB", in the order theme indexes refer to them. */
  themeColors: string[];
  /** Set once a format was added after the styles were read. */
  modified?: boolean;
}

/**
 * A differential format (<dxf>): only what it states is applied on top of the
 * format of a cell. A view for reading; the record itself stays in the package.
 */
export interface DifferentialFormat {
  font?: {
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    underline?: Underline;
    color?: Color;
  };
  /**
   * As the file states it, with one difference from a cell fill: a <patternFill>
   * without a patternType is a solid fill, and the colour of a solid fill is `bg`
   * (see differentialFillColor()).
   */
  fill?: Fill;
  border?: {
    left?: BorderEdge;
    right?: BorderEdge;
    top?: BorderEdge;
    bottom?: BorderEdge;
  };
  numFmt?: { id: number; code: string };
}

/** A threshold of a colour scale, a data bar or an icon set (<cfvo>). */
export interface ConditionalValue {
  /** "min", "max", "num", "percent", "percentile", "formula"; "autoMin" and "autoMax" in the 2010 extension. */
  type: string;
  /** A number or a formula (without "="), as written. */
  value?: string;
  /** Icon sets: the threshold itself belongs to the icon above it. Default true. */
  gte: boolean;
}

export interface ConditionalColorScale {
  values: ConditionalValue[];
  colors: Color[];
}

export interface ConditionalDataBar {
  min: ConditionalValue;
  max: ConditionalValue;
  color?: Color;
  showValue: boolean;
  /** Shortest and longest bar, in percent of the cell width. */
  minLength: number;
  maxLength: number;
  /** The rest comes from the 2010 extension (<x14:dataBar>); `extended` says that one was found. */
  extended: boolean;
  gradient: boolean;
  borderColor?: Color;
  negativeFillColor?: Color;
  negativeBorderColor?: Color;
  negativeFillSameAsPositive: boolean;
  negativeBorderSameAsPositive: boolean;
  axisColor?: Color;
  axisPosition: 'automatic' | 'middle' | 'none';
  direction: 'context' | 'leftToRight' | 'rightToLeft';
}

export interface ConditionalIconSet {
  /** "3Arrows", "4Rating", "5Quarters", ...; "3TrafficLights1" when the file names none. */
  set: string;
  values: ConditionalValue[];
  showValue: boolean;
  reverse: boolean;
  /** Icons chosen one by one (2010 extension), lowest first; a set of "NoIcons" shows nothing. */
  custom?: { set: string; id: number }[];
}

/**
 * An attribute the model has no field for: the namespace it is in ('' for none, which is what
 * an attribute without a prefix is in), its local name and its value. The namespace is the
 * one the prefix of the file stood for, never the prefix: another part may bind it otherwise.
 */
export type KeptAttr = [ns: string, name: string, value: string];

/**
 * An element the model has no field for, as a small tree: its namespace (the one a name
 * without a prefix is in included), its local name, its attributes and what it holds - elements
 * and text, in the order of the file. Comments and processing instructions are not in it.
 */
export interface KeptNode {
  ns: string;
  name: string;
  attrs?: KeptAttr[];
  children?: (KeptNode | string)[];
}

/** What one element the model does have fields for (a threshold, a colour) holds beside them. */
export interface KeptItem {
  attrs?: KeptAttr[];
  children?: KeptNode[];
}

/**
 * What the <colorScale>, <dataBar> or <iconSet> of a rule holds beside the fields of the model.
 * It goes with the rule while the rule states an element of that kind.
 */
export interface KeptHolder {
  kind: 'colorScale' | 'dataBar' | 'iconSet';
  attrs?: KeptAttr[];
  /**
   * Elements among its thresholds that the model does not know: those before the first <cfvo>,
   * before the second, ..., and the last entry those behind the last one (written at the end of
   * the element).
   */
  between?: KeptNode[][];
  /** Per threshold, what its <cfvo> holds beside type, value (and, in an icon set, `gte`); null for one that holds nothing else. Used while the rule has as many thresholds. */
  cfvo?: (KeptItem | null)[];
  /** The same per <color> of a scale or a bar. */
  colors?: (KeptItem | null)[];
}

/**
 * What the element of a rule (<cfRule>, <dataValidation>) holds beside the fields of the model,
 * with every name resolved to its namespace, as a tree: what the reader collects, what a writer
 * states, and what the model holds packed into one text (see RuleKept, packKept()).
 * What names the element itself is never in it (an attribute `uid` of a revision namespace, the
 * <ext> that ties a rule to its counterpart in the extension list): no second element may state that.
 */
export interface KeptContent {
  /** Attributes of the element, in the order of the file. */
  attrs?: KeptAttr[];
  /** Children of the element, in the order of the file. */
  children?: KeptNode[];
  /** Of a conditional rule: what its scale, bar or icon set holds. */
  holder?: KeptHolder;
}

/**
 * What the element of a rule holds beside the fields of the model, as the model holds it: the
 * reader fills it, a rule that is copied takes it along, and the writer states it wherever the
 * rule is written without its own element at hand. A rule without any such content has no `kept`.
 */
export interface RuleKept {
  /**
   * The content (KeptContent) as one text: see packKept() and unpackKept(). A text takes the
   * memory of its characters, where a tree of objects takes many times that; two rules that
   * hold the same have the same text.
   */
  packed?: string;
  /**
   * The namespaces among those used in the content that the part of the sheet the rule is in
   * lets readers pass over (mc:Ignorable of its root), sorted. Content in one of them is written
   * only into a part whose root says the same of that namespace; elsewhere it is left out, and
   * whoever puts the rule there counts it (see keptFor() of the rule commands).
   */
  ignorable?: string[];
  /**
   * The element holds what a rule does not carry: more than MAX_KEPT_SIZE or deeper than
   * MAX_KEPT_DEPTH, more than the budget of its sheet or workbook has left (MAX_KEPT_SHEET,
   * MAX_KEPT_BOOK), content that names prefixes in its values (the markup-compatibility
   * namespace, xsi:type), a comment or a processing instruction, a prefix nothing declares, or
   * more of an element than the model has room for. Nothing of it is in the model. In the file
   * it stays byte for byte, and a rule that is changed where it is still has all of it; the rule
   * is not copied (it is kept as the file has it), which is counted where a copy is asked for.
   */
  over?: true;
}

/** The namespace kept content is said to be in when it is in the one of the part itself (transitional or strict): every part writes it as its own. */
export const KEPT_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Most of what a rule carries beside its fields: the characters of all names, namespaces and values together, and how deep its elements nest. */
export const MAX_KEPT_SIZE = 8192;
export const MAX_KEPT_DEPTH = 8;
/** Most of it the rules of one sheet, and those of one workbook, hold together (in characters); the content of a rule beyond that stays in the file alone, as one over MAX_KEPT_SIZE does. */
export const MAX_KEPT_SHEET = 4 * 1024 * 1024;
export const MAX_KEPT_BOOK = 16 * 1024 * 1024;

const NS_XML_SPACE = 'http://www.w3.org/XML/1998/namespace';
const packText = (text: string) => text.replace(/[&<>\r]/g, (ch) => (ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : '&#13;'));
const packValue = (text: string) => text.replace(/[&<"\t\n\r]/g, (ch) => (ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '"' ? '&quot;' : `&#${ch.charCodeAt(0)};`));

/**
 * Kept content as one text: a small XML document of its own, in which every namespace of the
 * content has a prefix the document declares itself ("n0", "n1", ... in the order they are first
 * used), so that the text says all there is to say. Equal content gives equal text. Undefined
 * for content that holds nothing.
 *
 *   <k xmlns:n0="..."> <a .../> <c>...</c> <h k="iconSet"> <a .../> <b>...</b>... <vs><v><a .../><c>...</c></v>...</vs> <ts>...</ts> </h> </k>
 *
 * <a> carries attributes, <c> children, <b> one list of what stands among the thresholds, <vs>
 * the thresholds and <ts> the colours (one <v> each). The elements of the document itself are in
 * no namespace and stand at fixed places; the content is what is inside <a>, <c> and <b>.
 */
export function packKept(content: KeptContent | undefined): string | undefined {
  if (!content || (!content.attrs && !content.children && !content.holder)) return undefined;
  const prefixes = new Map<string, string>();
  const named = (ns: string, name: string) => {
    if (ns === '') return name;
    if (ns === NS_XML_SPACE) return `xml:${name}`;
    let prefix = prefixes.get(ns);
    if (prefix === undefined) prefixes.set(ns, (prefix = `n${prefixes.size}`));
    return `${prefix}:${name}`;
  };
  const attrs = (list: KeptAttr[] | undefined) => {
    if (!list || list.length === 0) return '';
    let out = '<a';
    for (const attr of list) out += ` ${named(attr[0], attr[1])}="${packValue(attr[2])}"`;
    return `${out}/>`;
  };
  const node = (item: KeptNode): string => {
    let out = `<${named(item.ns, item.name)}`;
    for (const attr of item.attrs || []) out += ` ${named(attr[0], attr[1])}="${packValue(attr[2])}"`;
    if (!item.children || item.children.length === 0) return `${out}/>`;
    out += '>';
    for (const child of item.children) out += typeof child === 'string' ? packText(child) : node(child);
    return `${out}</${named(item.ns, item.name)}>`;
  };
  const nodes = (tag: string, list: (KeptNode | string)[] | undefined, always = false) => {
    if (!list || list.length === 0) return always ? `<${tag}/>` : '';
    let out = `<${tag}>`;
    for (const item of list) if (typeof item !== 'string') out += node(item);
    return `${out}</${tag}>`;
  };
  const items = (tag: string, list: (KeptItem | null)[] | undefined) => {
    if (!list) return '';
    let out = `<${tag}>`;
    for (const item of list) out += item ? `<v>${attrs(item.attrs)}${nodes('c', item.children)}</v>` : '<v/>';
    return `${out}</${tag}>`;
  };
  let body = attrs(content.attrs) + nodes('c', content.children);
  const holder = content.holder;
  if (holder) {
    body += `<h k="${holder.kind}">${attrs(holder.attrs)}`;
    for (const list of holder.between || []) body += nodes('b', list, true);
    body += `${items('vs', holder.cfvo)}${items('ts', holder.colors)}</h>`;
  }
  let root = '<k';
  for (const [ns, prefix] of prefixes) root += ` xmlns:${prefix}="${packValue(ns)}"`;
  // One text of its own: put together piece by piece, it would be held as the tree of its pieces, which takes many times its size.
  return [root, '>', body, '</k>'].join('');
}

/** The content a text made by packKept() holds; undefined for none, and for a text that is no such text. */
export function unpackKept(packed: string | undefined): KeptContent | undefined {
  if (typeof packed !== 'string' || packed === '') return undefined;
  let root: XmlElement;
  try {
    root = parseXml(packed).root;
  } catch (error) {
    return undefined;
  }
  if (root.name !== 'k') return undefined;
  const spaces = new Map<string, string>();
  for (const [name, value] of root.attrs) if (name.slice(0, 6) === 'xmlns:') spaces.set(name.slice(6), value);
  const split = (name: string): [string, string] => {
    const colon = name.indexOf(':');
    if (colon < 0) return ['', name];
    const prefix = name.slice(0, colon);
    return [prefix === 'xml' ? NS_XML_SPACE : spaces.get(prefix) || '', name.slice(colon + 1)];
  };
  const attrs = (el: XmlElement): KeptAttr[] | undefined => {
    let out: KeptAttr[] | undefined;
    for (const [name, value] of el.attrs) {
      const [ns, local] = split(name);
      (out || (out = [])).push([ns, local, value]);
    }
    return out;
  };
  const node = (el: XmlElement): KeptNode => {
    const [ns, name] = split(el.name);
    const out: KeptNode = { ns, name };
    const own = attrs(el);
    if (own) out.attrs = own;
    let children: (KeptNode | string)[] | undefined;
    for (const child of el.children) {
      if (child.type === 'element') (children || (children = [])).push(node(child));
      else if (child.type === 'text' && child.value !== '') {
        const last = children ? children[children.length - 1] : undefined;
        if (children && typeof last === 'string') children[children.length - 1] = last + child.value;
        else (children || (children = [])).push(child.value);
      }
    }
    if (children) out.children = children;
    return out;
  };
  const elements = (el: XmlElement) => el.children.filter((child): child is XmlElement => child.type === 'element');
  const nodes = (el: XmlElement) => elements(el).map(node);
  const items = (el: XmlElement) =>
    elements(el).map((item): KeptItem | null => {
      const out: KeptItem = {};
      for (const part of elements(item)) {
        if (part.name === 'a') {
          const own = attrs(part);
          if (own) out.attrs = own;
        } else if (part.name === 'c') out.children = nodes(part);
      }
      return out.attrs || out.children ? out : null;
    });
  const content: KeptContent = {};
  for (const part of elements(root)) {
    if (part.name === 'a') {
      const own = attrs(part);
      if (own) content.attrs = own;
    } else if (part.name === 'c') content.children = nodes(part);
    else if (part.name === 'h') {
      const kind = part.attrs.get('k');
      if (kind !== 'colorScale' && kind !== 'dataBar' && kind !== 'iconSet') continue;
      const holder: KeptHolder = { kind };
      for (const inner of elements(part)) {
        if (inner.name === 'a') {
          const own = attrs(inner);
          if (own) holder.attrs = own;
        } else if (inner.name === 'b') (holder.between || (holder.between = [])).push(nodes(inner));
        else if (inner.name === 'vs') holder.cfvo = items(inner);
        else if (inner.name === 'ts') holder.colors = items(inner);
      }
      content.holder = holder;
    }
  }
  return content.attrs || content.children || content.holder ? content : undefined;
}

/** What a rule that is put into a sheet has to know of the part of that sheet. */
export interface SheetSpaces {
  /** The namespaces the root of the part lets readers pass over (the prefixes of its mc:Ignorable, resolved), sorted. */
  ignorable: string[];
}

/** One <cfRule>. Attributes that only some rule types use are present when the file states them. */
export interface ConditionalRule {
  /** "cellIs", "expression", "colorScale", "dataBar", "iconSet", "top10", ... as in ST_CfType; unknown types are kept and never match. */
  type: string;
  /** Lower numbers are applied first; Infinity when the file states none. */
  priority: number;
  stopIfTrue: boolean;
  /** Index into StyleSheet.dxfFormats. */
  dxfId?: number;
  /** A format stated in the rule itself (rules of the 2010 extension). */
  dxf?: DifferentialFormat;
  /** cellIs: "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual". */
  operator?: string;
  /** Formula texts without "=", written for the top-left cell of the first range. */
  formulas: string[];
  /** containsText, notContainsText, beginsWith, endsWith. */
  text?: string;
  timePeriod?: string;
  /** top10. */
  rank?: number;
  percent: boolean;
  bottom: boolean;
  /** aboveAverage: false means "below average". */
  aboveAverage: boolean;
  equalAverage: boolean;
  stdDev?: number;
  colorScale?: ConditionalColorScale;
  dataBar?: ConditionalDataBar;
  iconSet?: ConditionalIconSet;
  /** Id that ties a rule to its counterpart in the 2010 extension list. */
  extId?: string;
  /** The rule exists only in the 2010 extension list. */
  extended?: boolean;
  /**
   * Set by the reader on a rule of the main part: which <conditionalFormatting> element of the
   * sheet part it is in (counted among the children of the worksheet) and which <cfRule> of
   * that element it is. A rule that was changed keeps it: the writer finds by it what the
   * source element holds beside the fields of the model. A rule added since has none.
   */
  src?: [format: number, rule: number];
  /** What the element of the rule holds beside these fields (see RuleKept); not read for a rule of the extension list, which is never written from the model. */
  kept?: RuleKept;
}

/**
 * One <conditionalFormatting>: rules and the cells they apply to. One that was read is kept
 * as the file has it, and the writer copies its source element. One that was added or changed
 * since is `live`.
 */
export interface ConditionalFormat {
  ranges: CellRange[];
  rules: ConditionalRule[];
  /** Set by the reader: the index of the <conditionalFormatting> element among the children of the worksheet. */
  src?: number;
  /**
   * Set on a format that was added or changed since the sheet was read: its ranges, and the
   * formulas of its rules, are those of the sheet as it is now, under the names the sheets
   * have now, and every operation that moves cells or names sheets keeps them so. A format
   * without it is as the file has it (see Sheet.conditionalFormats). conditionalFormatNow()
   * is the one rule for where either applies.
   */
  live?: boolean;
}

/** One condition of a custom filter ("greater than 5", "begins with a"). */
export interface FilterCondition {
  /** "equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual". */
  operator: string;
  /** As written; "*" and "?" are wildcards of "equal" and "notEqual". */
  value: string;
}

/** A date or a part of one chosen in a filter list (<dateGroupItem>): everything up to `grouping` has to match. */
export interface FilterDateGroup {
  /** "year", "month", "day", "hour", "minute", "second". */
  grouping: string;
  year: number;
  month?: number;
  day?: number;
  hour?: number;
  minute?: number;
  second?: number;
}

/**
 * What a filter asks of one of its columns (<filterColumn>). Exactly one kind
 * of criterion is stated; a column without any has only button settings.
 */
export interface FilterColumn {
  /** Offset of the column from the left edge of the filter range. */
  colId: number;
  /**
   * "values": the texts in `values` (compared with what the cells show), blanks and dates;
   * "custom": one or two conditions; "top10"; "dynamic" (above average, today, ...);
   * "color" and "icon" are read and kept but never applied; "none": no criterion;
   * "unknown": something this engine does not know, kept as it is.
   */
  kind: 'values' | 'custom' | 'top10' | 'dynamic' | 'color' | 'icon' | 'none' | 'unknown';
  /** values: the texts that are shown. */
  values?: string[];
  /** values: empty cells are shown as well. */
  blank?: boolean;
  dateGroups?: FilterDateGroup[];
  /** custom: both conditions have to hold (`and`) or one of them. */
  and?: boolean;
  conditions?: FilterCondition[];
  /** top10: the `value` largest (or smallest) items, or that percentage of them. */
  top?: boolean;
  percent?: boolean;
  value?: number;
  /** dynamic: "aboveAverage", "belowAverage", "today", ... as in ST_DynamicFilterType. */
  dynamicType?: string;
  /** The column shows no button (hiddenButton="1" or showButton="0"). */
  hiddenButton?: boolean;
  /**
   * The source element. While it is there the column is written back exactly as it
   * was read (attributes and extensions the model has no field for included);
   * a column changed by an edit has none.
   */
  raw?: string;
}

export interface SortCondition {
  /** The cells of the key column or row. */
  range: CellRange;
  descending: boolean;
  /** The source element, for what the model has no field for (sortBy, customList, dxfId, iconSet). */
  raw?: string;
}

/** How a range was last sorted (<sortState>); kept so that the application can sort again. */
export interface SortState {
  range: CellRange;
  conditions: SortCondition[];
  caseSensitive?: boolean;
  columnSort?: boolean;
  /** Attributes and children the model has no field for, as they go into the element. */
  attrs?: [name: string, value: string][];
  extra?: string;
}

/** Defined name, local to a sheet and hidden, under which applications keep the range of the sheet's filter. */
export const FILTER_NAME = '_xlnm._FilterDatabase';

/** The filter of a sheet (<autoFilter>): a header row with a button per column, and the rows under it. */
export interface AutoFilter {
  /** Header row and data rows. */
  range: CellRange;
  /** Columns that state something, by colId; at most one per column. */
  columns: FilterColumn[];
  sortState?: SortState;
  /** Attributes of the element the model has no field for (xr:uid, ...), kept to be written back. */
  attrs?: [name: string, value: string][];
  /** Child elements the model has no field for (<extLst>), as source text. */
  extra?: string;
  /**
   * Rows that a criterion this engine cannot apply (a colour, an icon) keeps hidden, noted
   * when another criterion is changed, so that they stay hidden whatever that one shows.
   * Not part of the file.
   */
  held?: number[];
}

/**
 * One <dataValidation>: what may be typed into its cells. A rule that was read is kept as the
 * file has it, and the writer copies its source element; one that was added or changed since
 * is `live`.
 */
export interface DataValidation {
  ranges: CellRange[];
  /** "none", "whole", "decimal", "list", "date", "time", "textLength", "custom". */
  type: string;
  /** "between" (the default), "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual". */
  operator: string;
  /** Formula texts without "=", written for the top-left cell of the first range. */
  formula1?: string;
  formula2?: string;
  allowBlank: boolean;
  /** A list shows no arrow in the cell. The file calls this showDropDown="1": the attribute means the opposite of its name. */
  hideDropDown: boolean;
  showInputMessage: boolean;
  promptTitle?: string;
  prompt?: string;
  showErrorMessage: boolean;
  errorStyle: 'stop' | 'warning' | 'information';
  errorTitle?: string;
  error?: string;
  /** The rule comes from the 2010 extension list (its list may be on another sheet). */
  extended?: boolean;
  /**
   * Set by the reader on a rule of the main part: which <dataValidation> element of the sheet
   * part it is, counted through all <dataValidations> of the worksheet. A rule that was
   * changed keeps it: the writer finds by it what the source element holds beside the fields
   * of the model. A rule added since has none.
   */
  src?: number;
  /**
   * Set on a rule that was added or changed since the sheet was read: its ranges and formulas
   * are those of the sheet as it is now, under the names the sheets have now, and every
   * operation that moves cells or names sheets keeps them so. A rule without it is as the
   * file has it (see Sheet.dataValidations). dataValidationNow() is the one rule for where
   * either applies.
   */
  live?: boolean;
  /** What the element of the rule holds beside these fields (see RuleKept). */
  kept?: RuleKept;
}

export interface RichRun {
  text: string;
  font?: Font;
}

export interface SharedString {
  /** Plain text of the string (all runs joined, phonetic text excluded). */
  text: string;
  /** Present only for rich text. */
  runs?: RichRun[];
}

export interface CellFormula {
  /** Formula text without the leading "=". Empty for a shared formula that only refers to its master. */
  text: string;
  type: 'normal' | 'array' | 'shared' | 'dataTable';
  /** Range the array or shared formula applies to, e.g. "A1:A10". */
  ref?: string;
  /** Shared formula group index. */
  si?: number;
  /** Attributes of <f> the model has no field for, kept to be written back. */
  attrs?: Map<string, string>;
}

export type CellKind = 'blank' | 'number' | 'string' | 'boolean' | 'error' | 'date';

export interface Cell {
  kind: CellKind;
  /**
   * number -> number, boolean -> boolean, string -> string, error -> the error
   * text ("#DIV/0!"), date -> ISO 8601 text (cells of type "d"), blank -> null.
   * Dates are normally plain numbers with a date number format.
   */
  value: number | string | boolean | null;
  /** Index into StyleSheet.cellXfs. */
  style: number;
  formula?: CellFormula;
  /** Rich text runs, when the text has mixed formatting. Shared between cells that use the same string. */
  runs?: RichRun[];
  /**
   * Index of the shared string this cell was read from. Only cells that come
   * from the file have it; the writer uses it to put the cell back exactly as
   * it was, phonetic text included.
   */
  si?: number;
  /** Cell metadata index (dynamic array formulas) and value metadata index (rich values), kept to be written back. */
  cm?: number;
  vm?: number;
  /** The cell shows phonetic text. */
  ph?: boolean;
}

export interface Row {
  /** Height in points when the row states one. */
  height?: number;
  customHeight: boolean;
  hidden: boolean;
  /** Row-level format (cellXfs index) applied to cells without their own. */
  style?: number;
  outlineLevel: number;
  collapsed: boolean;
  /** Sparse, indexed by column. */
  cells: Cell[];
  /**
   * Attributes of <row> the model has no field for (thickTop, ph, x14ac:dyDescent, ...),
   * kept to be written back. "spans" is not kept: it is a hint that is stale after any edit.
   */
  attrs?: [name: string, value: string][];
}

/** Properties shared by a run of columns, as in <col min max>. */
export interface ColumnRange {
  /** First and last column, inclusive. */
  min: number;
  max: number;
  /** Width in characters of the default font's widest digit. */
  width?: number;
  customWidth: boolean;
  hidden: boolean;
  style?: number;
  outlineLevel: number;
  collapsed: boolean;
}

/** Inclusive on all four sides. */
export interface CellRange {
  top: number;
  left: number;
  bottom: number;
  right: number;
}

export interface Hyperlink {
  range: CellRange;
  /** Relationship id of an external target; `target` is filled in from the relationships. */
  rId?: string;
  target?: string;
  /** Location inside the workbook, e.g. "Sheet2!A1". */
  location?: string;
  display?: string;
  tooltip?: string;
}

export interface SheetView {
  showGridLines: boolean;
  showRowColHeaders: boolean;
  showZeros: boolean;
  rightToLeft: boolean;
  tabSelected: boolean;
  /** Percent, 100 = normal. */
  zoomScale: number;
  /** Rows and columns kept in place by frozen panes; 0 when not frozen. */
  frozenRows: number;
  frozenCols: number;
  /** First cell of the scrollable area of frozen panes, e.g. "B2". */
  topLeftCell?: string;
  /**
   * Top-left cell of the view as a whole. With frozen panes this is where the
   * frozen rows and columns start: panes frozen while scrolled to C5 keep row 5
   * and column C in place, not row 1 and column A.
   */
  firstVisibleCell?: string;
  activeCell?: string;
}

/** A corner of a drawing: a cell and how far into it, in EMU (914,400 per inch; 9,525 per pixel). */
export interface DrawingMarker {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
  /**
   * How far into the cell as a share (0..1) of the column's width and the row's height, for
   * files that state it so (xls). Where present they take the place of colOff and rowOff.
   */
  colShare?: number;
  rowShare?: number;
}

/**
 * Where a drawing is on its sheet. "twoCell": from one cell to another; "oneCell": from a
 * cell, with a size; "absolute": a place and a size on the sheet, whatever its cells do.
 * Sizes and absolute positions are in EMU.
 */
export interface DrawingAnchor {
  kind: 'twoCell' | 'oneCell' | 'absolute';
  /** twoCell only: what rows and columns inserted or deleted do to it. "twoCell" when the file states nothing. */
  editAs?: 'twoCell' | 'oneCell' | 'absolute';
  from?: DrawingMarker;
  to?: DrawingMarker;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
}

export interface DrawingTextRun {
  text: string;
  /** Font name as the file states it; never trusted. */
  font?: string;
  /** Points. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** "#RRGGBB". */
  color?: string;
}

export interface DrawingParagraph {
  runs: DrawingTextRun[];
  align?: 'left' | 'center' | 'right' | 'justify';
}

/** The text of a shape. */
export interface DrawingText {
  paragraphs: DrawingParagraph[];
  /** Where the text stands in the height of the shape. Default "top". */
  anchor?: 'top' | 'middle' | 'bottom';
  /** Space between the shape and its text: left, top, right, bottom, in pixels. */
  insets?: [number, number, number, number];
  /** Lines are broken at the width of the shape. Default true. */
  wrap?: boolean;
}

/** What every object of a drawing has, whatever it shows. */
export interface DrawingObjectBase {
  name?: string;
  /** Alternative text. */
  description?: string;
  hidden?: boolean;
  /** Degrees, clockwise. */
  rotation?: number;
  flipH?: boolean;
  flipV?: boolean;
  /**
   * For an object inside a group: its place in the group, as shares (0..1) of the
   * group's width and height. The object of an anchor fills the anchor and has none.
   */
  box?: { x: number; y: number; width: number; height: number };
}

export interface DrawingPicture extends DrawingObjectBase {
  type: 'picture';
  /** Key of the picture in Workbook.media; missing for a picture that is not in the file. */
  media?: string;
  /** The picture is linked from outside the file. It is never fetched. */
  external?: boolean;
  /** Shares (0..1) cut off each side of the picture. */
  crop?: { left: number; top: number; right: number; bottom: number };
}

export interface DrawingShape extends DrawingObjectBase {
  type: 'shape';
  /** Preset geometry name ("rect", "ellipse", "line", ...); "rect" when the file states none. */
  geometry: string;
  /** A connector or a line: drawn from corner to corner. */
  connector?: boolean;
  /** "#RRGGBB"; missing for a shape without a fill (or with one that is not drawn: gradients, patterns). */
  fill?: string;
  /** Outline: colour ("#RRGGBB") and width in pixels. */
  line?: { color: string; width: number; dash?: boolean };
  text?: DrawingText;
}

export interface DrawingGroup extends DrawingObjectBase {
  type: 'group';
  children: DrawingObject[];
}

/** Something that is not drawn as it is: a chart of a kind that is not read (or that cannot be read), a diagram, an embedded object. */
export interface DrawingPlaceholder extends DrawingObjectBase {
  type: 'placeholder';
  what: 'chart' | 'diagram' | 'object';
}

/**
 * A chart whose part could be read. The frame is the anchor, as for any other object; what
 * the chart shows is worked out from the cells its series refer to whenever it is drawn.
 */
export interface DrawingChart extends DrawingObjectBase {
  type: 'chart';
  chart: Chart;
  /**
   * The chart part, e.g. "/xl/charts/chart1.xml"; missing for a chart that is in no file (one
   * added since the workbook was read: the writer gives it a part). A part that was not changed
   * is written back as it is; see ChartOrigin for one that was.
   */
  part?: string;
  /**
   * Set by the reader, for a chart of a file, by what its part holds beside the model:
   * "all" when the part is exactly what the writer makes of the model (the engine wrote it): such a
   * chart can be changed in every way, and is written anew from the model;
   * "refs" when every formula of the part is in the model (the references of the series, of their
   * names and of the titles): rows and columns inserted or deleted and sheets renamed can be followed,
   * and what the model holds of it can be changed in the part while everything else of the part stays.
   * Missing for a chart of which neither is known: it is shown, moved and deleted, not changed.
   */
  known?: 'all' | 'refs';
  /**
   * Set by the reader beside `known: 'refs'` when the part holds content that is bound to the
   * number and order of its series and that the model does not hold (a legend entry of a trend
   * line, the formats of a pivot chart): the series are changed, not added, removed or reordered.
   */
  seriesFixed?: boolean;
  /** Set once a chart of the file was changed: what it was read as. On a copy (`source`): what the part it is made from was read as. */
  origin?: ChartOrigin;
  /**
   * Set by the reader: the macro the frame of the chart is assigned in the file (the `macro` of its
   * <xdr:graphicFrame>), when it has one. It is never run; the frame keeps it, and a copy of the
   * chart is made without it (and says so).
   */
  macro?: string;
  /**
   * Set on a chart that was pasted as a copy of a chart of the file: the chart part of that
   * chart, e.g. "/xl/charts/chart1.xml". The writer makes the frame of the copy (its anchor in
   * the drawing) from the frame that part has in the file, so that what a frame says beside its
   * name and its place (whether it is published, locked, printed, hidden, its title) goes with
   * the copy. Without it, and when the file has no such frame, the frame is a plain one.
   */
  frame?: string;
  /**
   * Set by the reader on a chart of a file without `known` whose part would be "refs" but for
   * children of its root that lead out of the part: shapes drawn over the chart ("userShapes") and
   * data kept outside the workbook ("externalData"). The names of those. A copy of the chart is
   * made without them, and says so (see copyChart() of the operations); the chart itself is
   * moved and deleted like any other of which nothing is known.
   */
  loose?: string[];
  /**
   * Set on a chart that is in no file and was pasted as a copy of a chart of the file whose part
   * holds more than its model: where the writer takes the part of the copy from. Such a chart has
   * no `part`; it has `known: 'refs'`, `seriesFixed` as the chart it was copied from and always an
   * `origin`, and is changed the way a chart of the file with those is.
   */
  source?: ChartSource;
}

/**
 * The part a copy of a chart of the file is made from, as the file has it: its bytes with what
 * differs between `origin` and the chart changed in them, and with every part it leads to (styles,
 * colours, pictures of fills) copied beside it. The part is read when the workbook is written;
 * the file the workbook was read from still holds it then, whatever became of the chart it was
 * copied from.
 */
export interface ChartSource {
  /** The chart part of the file, e.g. "/xl/charts/chart1.xml". */
  part: string;
  /** Children of the root of the part that are left out of the copy, with the relationships only they use (DrawingChart.loose). */
  drops?: string[];
}

/**
 * What a chart of the file that was changed came from: the chart as it was read, and, for each
 * chart type (group) of the chart as it is now, the series it has by their places in the same
 * group of the chart as read (-1 for a series added since). The writer changes the part by the
 * differences between the two, and leaves the rest of it alone.
 */
export interface ChartOrigin {
  chart: Chart;
  series: number[][];
}

export type DrawingObject = DrawingPicture | DrawingShape | DrawingGroup | DrawingPlaceholder | DrawingChart;

/**
 * One anchor of a sheet's drawing with what it holds. A view for reading: the drawing
 * part stays in the package and is written back from there.
 */
export interface SheetDrawing {
  anchor: DrawingAnchor;
  object: DrawingObject;
  /**
   * Set on a drawing that was added, moved or changed since the sheet was read: how many of the
   * sheet's structuralEdits had been made by then. `anchor` is where the drawing was at that
   * time, and only the edits from this index on move it. A drawing of the file has none: it is
   * where the file has it, and every structural edit moves it. See drawingAnchorNow().
   */
  since?: number;
  /** Set on a drawing of the file that was moved or changed: the drawing as the sheet was read with it, by which the writer knows its anchor in the file. */
  from?: SheetDrawing;
  /**
   * Set by a reader on a drawing that shares a wrapper of alternatives (mc:AlternateContent) with
   * another drawing of the file: it is moved and changed like any other, and it is not deleted
   * (the wrapper cannot go with one of the drawings it holds).
   */
  wrapped?: boolean;
}

/** A comment (note) on a cell. A view for reading, like SheetDrawing. */
export interface SheetComment {
  row: number;
  col: number;
  author?: string;
  /** The whole text; for a legacy note that starts with the author's name, the name is part of it. */
  text: string;
  /** Present when the text has mixed formatting. */
  runs?: RichRun[];
  /** The comment is a threaded one (a conversation); `text` is then the first message and `replies` the rest. */
  threaded?: boolean;
  replies?: { author?: string; text: string }[];
  /** The box of the comment is shown all the time, not only when the cell is pointed at. */
  visible?: boolean;
  /** Where the box is, when the file says so. Offsets are in EMU, as in every anchor. */
  box?: DrawingAnchor;
  /**
   * Set on a note that was added since the sheet was read: how many of the sheet's
   * structuralEdits had been made by then. `row` and `col` are where the cell was at that
   * time, and only the edits from this index on move the note. A comment of the file has
   * none: it is where the file has it, and every structural edit moves it. (A note of the
   * file whose text is edited stays a comment of the file: only its text is another.)
   */
  since?: number;
  /**
   * Set, together with `since`, on a comment of the file that a sort moved with its row: the
   * cell the file has the comment on. `row`, `col` and `box` are then where the comment was put
   * (after `since` structural edits), like those of a note added since; the writer still knows
   * the comment of the file by this cell and writes it with its new one. A note added since
   * has none, however often it is sorted.
   */
  from?: { row: number; col: number };
}

/** The margins of the printed page, in inches (<pageMargins>). `header` and `footer` are the distances of the header and the footer from the edge of the paper. */
export interface PageMargins {
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
}

/**
 * Paper, orientation and scaling of the printed page (<pageSetup>). A field is present when the
 * file states it; a missing one has the default named with it.
 */
export interface PageSettings {
  /** Paper size code (1 Letter, 5 Legal, 8 A3, 9 A4, 11 A5, ...). Default 1; with `usePrinterDefaults` the printer's paper. */
  paperSize?: number;
  /** Size of a paper that has no code, as a measure with its unit ("210mm"); replaces `paperSize`. */
  paperWidth?: string;
  paperHeight?: string;
  /** "default" (the default) leaves the choice to the application, which prints upright. */
  orientation?: 'default' | 'portrait' | 'landscape';
  /** Percent, 10..400. Default 100. Not used while the sheet is fitted to pages (SheetPageSetup.fitToPage). */
  scale?: number;
  /** Pages the sheet is fitted to, across and down, while `fitToPage` is set; 0 is "as many as needed". Default 1. */
  fitToWidth?: number;
  fitToHeight?: number;
  /** Number of the first page; counts only with `useFirstPageNumber`. Default 1. */
  firstPageNumber?: number;
  /** Default false: pages are numbered from 1. */
  useFirstPageNumber?: boolean;
  /** The order pages are printed and numbered in. Default "downThenOver". */
  pageOrder?: 'downThenOver' | 'overThenDown';
  /** Default false. */
  blackAndWhite?: boolean;
  /** Draft quality, without graphics. Default false. */
  draft?: boolean;
  /** Whether comments are printed, and where. Default "none". */
  cellComments?: 'none' | 'asDisplayed' | 'atEnd';
  /** What is printed for a cell that holds an error. Default "displayed". */
  errors?: 'displayed' | 'blank' | 'dash' | 'NA';
  /** Print resolution in dots per inch; values near 4,294,967,295 are the printer driver's quality levels (draft, low, medium, high). Default 600. */
  horizontalDpi?: number;
  verticalDpi?: number;
  /** Default 1. */
  copies?: number;
  /** What the element does not state is the printer's own setting and not the default of the format. Default true. */
  usePrinterDefaults?: boolean;
}

/** What is printed beside the cells, and where on the page (<printOptions>). Present when the file states it, like PageSettings. */
export interface PrintOptions {
  /** The cells are centred between the left and the right margin. Default false. */
  horizontalCentered?: boolean;
  /** The cells are centred between the top and the bottom margin. Default false. */
  verticalCentered?: boolean;
  /** Row numbers and column letters are printed. Default false. */
  headings?: boolean;
  /** Grid lines are printed; only together with `gridLinesSet`. Default false. */
  gridLines?: boolean;
  /** Default true. */
  gridLinesSet?: boolean;
}

/**
 * Headers and footers of the printed page (<headerFooter>). A text is in the code of the
 * format: "&L", "&C" and "&R" start the left, centre and right section, "&P" is the page
 * number, "&N" the page count, "&D" the date, "&T" the time, "&F" the file name, "&A" the
 * sheet name, '&"font,style"' and "&12" set a font and a size, "&&" is an ampersand.
 */
export interface HeaderFooter {
  /** Header and footer of every page, or of the odd pages only while `differentOddEven` is set. */
  oddHeader?: string;
  oddFooter?: string;
  /** Used while `differentOddEven` is set. */
  evenHeader?: string;
  evenFooter?: string;
  /** Used while `differentFirst` is set. */
  firstHeader?: string;
  firstFooter?: string;
  /** Default false. */
  differentOddEven?: boolean;
  /** Default false. */
  differentFirst?: boolean;
  /** Headers and footers are scaled with the sheet. Default true. */
  scaleWithDoc?: boolean;
  /** Headers and footers start and end at the page margins. Default true. */
  alignWithMargins?: boolean;
}

/** A page break (<brk>). */
export interface PageBreak {
  /**
   * The first row (of rowBreaks) or column (of colBreaks) of the new page: the break is above
   * that row, left of that column. Never 0.
   */
  at: number;
  /** The columns (of a row break) or rows (of a column break) the break spans, first and last. */
  min: number;
  max: number;
  /** Set by hand, not by the application while fitting pages. */
  manual: boolean;
  /** Made by a pivot table. */
  pivot?: boolean;
}

/**
 * How a sheet is printed: one member per element of a worksheet part (<pageMargins>,
 * <pageSetup>, <printOptions>, <headerFooter>, <rowBreaks>, <colBreaks>, and <pageSetUpPr> of
 * <sheetPr>). A view for reading, like conditional formats: a save of the file it was read
 * from copies the source elements, and edits do not change it (structuralEdits are applied
 * to the breaks where a new file is written). Print areas and print titles are defined
 * names ("_xlnm.Print_Area", "_xlnm.Print_Titles"). Data of a printer driver is not held.
 */
export interface SheetPageSetup {
  margins?: PageMargins;
  page?: PageSettings;
  print?: PrintOptions;
  headerFooter?: HeaderFooter;
  /** The sheet is scaled to `page.fitToWidth` by `page.fitToHeight` pages, not by `page.scale`. Default false. */
  fitToPage?: boolean;
  /** Automatic page breaks are shown on the sheet. Default true. */
  autoPageBreaks?: boolean;
  /** In ascending order, one per row. */
  rowBreaks?: PageBreak[];
  /** In ascending order, one per column. */
  colBreaks?: PageBreak[];
}

/** Where the bytes of pictures come from; read only when a picture is first shown. */
export interface WorkbookMedia {
  readPart(key: string): Promise<Uint8Array>;
}

export type SheetState = 'visible' | 'hidden' | 'veryHidden';
export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface Sheet {
  name: string;
  sheetId: number;
  state: SheetState;
  kind: SheetKind;
  /**
   * Part this sheet was read from, e.g. "/xl/worksheets/sheet1.xml". Empty for a sheet that is
   * not in the file (one added since it was read): the writer gives such a sheet its part.
   */
  part: string;
  /**
   * Set by a reader whose file has no parts (xls) on every sheet it reads: the sheet is one of
   * the file although `part` is empty. Tells such a sheet from one added since the file was
   * read: see isFileSheet(). Never set on a sheet that was read from a part.
   */
  readFromFile?: boolean;
  /**
   * The name the sheet has in the file, set when a sheet that was read from the file is first
   * given another name. Content kept the way the file has it (conditional formats, data
   * validations, the source XML) names sheets as the file does: see fileSheetName() and
   * sheetNameMap(). Undoing the rename removes it again; renaming the sheet back to the name
   * of the file by another rename leaves it, equal to `name`.
   */
  nameInFile?: string;
  /** Sparse, keyed by row index. */
  rows: Map<number, Row>;
  /** Sorted by `min`, not overlapping. */
  columns: ColumnRange[];
  merges: CellRange[];
  hyperlinks: Hyperlink[];
  /** Used range as the file declares it; may be missing or wrong. */
  dimension?: CellRange;
  /** One past the last row and column that hold a cell (0 for an empty sheet). */
  rowCount: number;
  colCount: number;
  /** Default row height in points. */
  defaultRowHeight?: number;
  /** Default column width in characters. */
  defaultColWidth?: number;
  baseColWidth?: number;
  view: SheetView;
  tabColor?: Color;
  /** The colour of the tab was changed ("setTabColor"); the writer brings <tabColor> of the sheet part in line with the model. Set on its own, without `modified`. */
  tabColorModified?: boolean;
  /** The colour the tab had before it was first changed (none: `color` undefined). Set with `tabColorModified`. */
  tabColorRead?: { color: Color | undefined };
  /**
   * How the sheet is protected (<sheetProtection>), and the ranges that are open in a
   * protected sheet (<protectedRanges>), as the file states them. Not applied and not changed
   * by edits: a save of the file copies the source elements (and moves the ranges with the
   * rows and columns inserted and deleted since, see structuralEdits); a copy of the sheet
   * and a sheet written from the model get them from here with the ranges moved the same
   * way (protectionNow() of the xlsx builder).
   */
  protection?: SheetProtection;
  /**
   * Local names of what the sheet part holds beside what the model has a field for, one entry
   * per element (or attribute) found: "sortState", "customSheetViews", "ignoredErrors",
   * "cellWatches", "phoneticPr", "sparklineGroups" (of the extension list), "outlinePr" and
   * "codeName" (of <sheetPr>), "outlineLevels" (of <sheetFormatPr>), "legacyDrawingHF" (the
   * pictures of headers and footers), "picture" (the background picture), ... Names only: the content
   * stays in the part and is written back from there. A copy of the sheet does not get it,
   * which is what the list is for (see copySheetLosses()).
   */
  unmodelled?: string[];
  /**
   * Set by the reader for a part whose root lets readers pass over namespaces (mc:Ignorable):
   * content of such a namespace that a rule holds beside the fields of the model can be put into
   * this sheet only if its part says the same of that namespace (see RuleKept.ignorable and
   * keptFor()). A sheet that is in no file has none.
   */
  spaces?: SheetSpaces;
  /**
   * Conditional formats. Those of the file are as the file has them and are not changed by
   * edits: rows and columns inserted or deleted since (structuralEdits) are applied where they
   * are evaluated. Those added or changed since are `live` (see ConditionalFormat.live). The
   * list is replaced, never altered, by an edit ("setConditionalFormats").
   */
  conditionalFormats?: ConditionalFormat[];
  /**
   * Conditional formats were added, changed or deleted; the writer brings the elements of the
   * sheet part in line with the model. Set on its own, without `modified`: no cell changed.
   */
  conditionalModified?: boolean;
  /** The list the sheet had before its conditional formats were first changed. Set with `conditionalModified`. */
  conditionalRead?: { list: ConditionalFormat[] | undefined };
  /** The filter of the sheet, kept up to date by edits. Rows it hides are `hidden`, like rows hidden by hand. */
  autoFilter?: AutoFilter;
  /** The filter was changed (switched on or off, a criterion, a sort); the writer regenerates <autoFilter> from the model. */
  autoFilterModified?: boolean;
  /**
   * Data validations. Those of the file are as the file has them and, like conditional
   * formats, are not changed by edits: structuralEdits are applied where they are evaluated
   * (see DataValidations). Those added or changed since are `live` (see DataValidation.live).
   * The list is replaced, never altered, by an edit ("setDataValidations").
   */
  dataValidations?: DataValidation[];
  /**
   * Data validations were added, changed or deleted; the writer brings <dataValidations> of
   * the sheet part in line with the model. Set on its own, without `modified`: no cell changed.
   */
  validationsModified?: boolean;
  /** The list the sheet had before its data validations were first changed. Set with `validationsModified`. */
  validationsRead?: { list: DataValidation[] | undefined };
  /** Margins, paper, scaling, headers and footers and page breaks of the printed sheet, as the file has them. */
  pageSetup?: SheetPageSetup;
  /** Set once anything in the sheet was changed after it was read; the writer regenerates only such sheets. */
  modified?: boolean;
  /**
   * Rows in which a cell was changed since the sheet was read (indexes as they
   * were when the change happened). While there are no structural edits, every
   * other row is written back byte for byte from the original part.
   */
  changedRows?: Set<number>;
  /** Column widths, visibility or formats were changed; the writer regenerates <cols>. */
  columnsModified?: boolean;
  /** Cells were merged or unmerged; the writer regenerates <mergeCells>. */
  mergesModified?: boolean;
  /** Links changed without the sheet's own cells moving; the writer regenerates <hyperlinks>. */
  hyperlinksModified?: boolean;
  /**
   * Kinds of content anchored to cells but stored outside the sheet part
   * ("drawing", "comments", "table", "pivotTable", ...). Their positions are
   * not adjusted yet, so rows and columns cannot be inserted or deleted in a
   * sheet that has any.
   */
  anchored?: string[];
  /**
   * Rows and columns inserted or deleted since the sheet was read, in order.
   * The writer replays them on the parts of the sheet the model does not hold
   * (conditional formats, validations, ...), so that their ranges follow.
   */
  structuralEdits?: SheetStructuralEdit[];
  /**
   * Pictures, shapes and charts as the file has them, bottom one first. A chart sheet has the
   * one drawing that holds its chart. Like
   * conditional formats they are not changed by edits: structuralEdits are applied where they
   * are shown (see SheetDrawings), by the rules the writer moves the anchors in the file with.
   */
  drawings?: SheetDrawing[];
  /**
   * Charts were added, moved, changed or deleted ("setDrawings"), or a chart follows rows,
   * columns or sheets that changed; the writer brings the drawing and chart parts of the sheet in
   * line with the model. Set on its own, without `modified`: no cell changed.
   */
  drawingsModified?: boolean;
  /**
   * Set by a reader on a sheet whose drawing part takes no change: why ("its drawing is shared
   * with another sheet", "its drawing is nested too deeply", "its drawing cannot be read"). No
   * chart is added to such a sheet, moved or deleted on it; a chart of it whose own part can be
   * changed (DrawingChart.known) still is. Drawings and charts are shown all the same.
   */
  drawingsFixed?: string;
  /** The list the sheet had before its drawings were first changed. Set with `drawingsModified`. */
  drawingsRead?: { list: SheetDrawing[] | undefined };
  /**
   * Comments as the file has them; structuralEdits are applied where they are shown. Notes
   * added by editing stand where they were put (SheetComment.since says after how many of the
   * structural edits), so that the same rule places both: see commentCellNow(). The list is
   * replaced, never altered, by an edit ("setComments").
   */
  comments?: SheetComment[];
  /**
   * Notes were added, edited or deleted; the writer brings the comment parts of the sheet in
   * line with the model. Set on its own, without `modified`: the cells did not change.
   */
  commentsModified?: boolean;
  /**
   * Set by the reader for a sheet whose comment parts the writer could not bring in line with
   * notes that are added, edited or deleted: its legacy drawing cannot be read ("drawing"),
   * another sheet shows the same comments part or legacy drawing ("shared"), or the sheet has
   * more than one comments part ("parts"). Notes of such a sheet are shown and not changed.
   */
  notesBlocked?: 'drawing' | 'shared' | 'parts';
  /** What of the sheet's drawings and comments could not be read and is not shown (it is still written back). */
  warnings?: string[];
}

/** <sheetProtection> and <protectedRanges> of a sheet: attributes as the file has them, never interpreted. */
export interface SheetProtection {
  /** The attributes of <sheetProtection> (flags, and the hash of a password); missing for a sheet with ranges alone. */
  attrs?: [name: string, value: string][];
  /** One entry per <protectedRange>: its attributes, and the texts of its <securityDescriptor> children. */
  ranges?: { attrs: [name: string, value: string][]; descriptors?: string[] }[];
}

export interface SheetStructuralEdit {
  kind: 'insert' | 'delete';
  axis: 'row' | 'col';
  at: number;
  count: number;
}

export interface DefinedName {
  name: string;
  /** The formula or reference text. */
  value: string;
  /** Index of the sheet the name is local to; undefined for workbook scope. */
  localSheet?: number;
  hidden: boolean;
}

export interface Workbook {
  sheets: Sheet[];
  styles: StyleSheet;
  sharedStrings: SharedString[];
  definedNames: DefinedName[];
  /** Dates count from 1904-01-01 instead of 1899-12-30. */
  date1904: boolean;
  /** Index of the sheet shown when the workbook opens. */
  activeSheet: number;
  /** The pictures of the workbook's drawings, by DrawingPicture.media. */
  media?: WorkbookMedia;
  /**
   * Set by the reader of a file with chart parts when every one of them is shown as a chart of a
   * worksheet whose references the model holds (DrawingChart.known): the charts then follow rows
   * and columns that are inserted or deleted and sheets that are renamed or deleted (see
   * charts.ts). Without it a workbook with charts (Sheet.anchored names "chart") keeps its rows,
   * columns and sheet names as they are.
   */
  chartsFollow?: boolean;
  /**
   * Set once a sheet was inserted, deleted, renamed, moved, copied, hidden, shown or given
   * another tab colour after the workbook was read (and never cleared, not by undo either): the
   * list of sheets, their names, order and states or the defined names may differ from the
   * file, and the writer regenerates what holds them.
   */
  sheetsModified?: boolean;
  /**
   * The names, as the file has them, of sheets of the file that were deleted. Undoing the
   * deletion takes the name out again. A sheet added after the file was read and deleted
   * again is not listed. Missing while nothing was deleted.
   */
  removedSheetNames?: string[];
  /**
   * One entry per name of removedSheetNames, in the same order: the sheets that stood before
   * and after the deleted sheet. A 3D reference that ended at the deleted sheet ends at one of
   * them (see SheetRuns). Kept pointing at sheets of the workbook: when such a neighbour is
   * deleted in turn, its own neighbour takes its place here.
   */
  removedSheetNeighbours?: SheetNeighbours[];
}

/** The sheets on either side of a deleted sheet; undefined where it was the first or the last. */
export interface SheetNeighbours {
  previous?: Sheet;
  next?: Sheet;
}

/**
 * Sheet names compare without regard to case. The four Turkish letters I, İ, ı, i count as
 * one, without relying on locale data: "İ" lower-cases to "i" plus a combining dot (U+0307).
 */
export function foldSheetName(name: string): string {
  return String(name).toLowerCase().replace(/\u0307/g, '').replace(/\u0131/g, 'i');
}

/**
 * What a sheet may be called: 1 to 31 characters that are not all spaces,
 * none of []:*?/\ and no control character, no apostrophe at either end,
 * not "History" (which Excel keeps for itself), and only characters an XML
 * file can hold. Returns the name; throws RangeError for a name that is none.
 */
export function checkSheetName(name: string): string {
  const ok =
    typeof name === 'string' &&
    name.length >= 1 &&
    name.length <= 31 &&
    name.trim() !== '' &&
    !/[[\]:*?/\\\u0000-\u001f\ufffe\uffff]/.test(name) &&
    // A surrogate without its other half.
    !/[\ud800-\udbff](?![\udc00-\udfff])|(?:^|[^\ud800-\udbff])[\udc00-\udfff]/.test(name) &&
    name.charAt(0) !== "'" &&
    name.charAt(name.length - 1) !== "'" &&
    name.toLowerCase() !== 'history';
  if (!ok) throw new RangeError('A sheet name has 1 to 31 characters, none of [ ] : * ? / \\ and no control character, no apostrophe at its start or end, and is not "History"');
  return name;
}

/**
 * Whether a sheet was read from the file, as opposed to added since: it has a part, or its
 * reader marked it (a file without parts). Content kept the way the file has it can only name
 * such sheets. Which sheets a writer has to make a part for is another question: that is
 * `part === ''` alone.
 */
export function isFileSheet(sheet: Sheet): boolean {
  return sheet.part !== '' || sheet.readFromFile === true;
}

/** The name a sheet has in the file it was read from; undefined for a sheet that is not in the file. */
export function fileSheetName(sheet: Sheet): string | undefined {
  if (!isFileSheet(sheet)) return undefined;
  return sheet.nameInFile !== undefined ? sheet.nameInFile : sheet.name;
}

/**
 * What a sheet name of the file stands for now: the name the sheet has today, null when the
 * sheet was deleted, undefined when the name is to be left as it is (the sheet still has it,
 * or the file has no sheet of that name).
 */
export interface SheetNameMap {
  (name: string): string | null | undefined;
  /** How a 3D reference goes on when the sheet at one of its ends was deleted; without it such a reference is an error. */
  runs?: SheetRuns;
}

/**
 * What a 3D reference ("Sheet2:Sheet4!A1") needs to know when a sheet at one of its ends was
 * deleted: the run then ends at the sheet that stood next to the deleted one, inside the run.
 */
export interface SheetRuns {
  /**
   * The name, as it is now, of the sheet that stood after (`step` 1) or before (`step` -1) the
   * deleted sheet the map answers null for; undefined when there is none.
   */
  neighbour(name: string, step: 1 | -1): string | undefined;
  /** Where the sheet of this name (as it is now) stands among the sheets; -1 for no sheet. */
  index(name: string): number;
}

/**
 * The map from the sheet names of the file to the sheets as they are now, for content that
 * is kept the way the file has it. Undefined while every sheet of the file is there under
 * its own name, which is the usual case and needs no work. All names are looked up in the
 * state of the file at once, never one rename after another: with "Sheet2" deleted and
 * "Sheet3" renamed to "Sheet2", the file's "Sheet2" is the deleted sheet (null) and the
 * file's "Sheet3" is today's "Sheet2". The map reflects the workbook at the time of the
 * call; ask again after the sheets changed.
 */
export function sheetNameMap(workbook: Workbook): SheetNameMap | undefined {
  const removed = workbook.removedSheetNames;
  let names: Map<string, string | null> | undefined;
  if (removed && removed.length > 0) {
    names = new Map();
    for (const name of removed) names.set(foldSheetName(name), null);
  }
  for (const sheet of workbook.sheets || []) {
    if (!isFileSheet(sheet) || sheet.nameInFile === undefined || sheet.nameInFile === sheet.name) continue;
    if (!names) names = new Map();
    names.set(foldSheetName(sheet.nameInFile), sheet.name);
  }
  if (!names) return undefined;
  const known = names;
  const map: SheetNameMap = (name) => known.get(foldSheetName(name));
  const neighbours = workbook.removedSheetNeighbours;
  if (removed && removed.length > 0 && neighbours) {
    const sheets = workbook.sheets;
    const beside = new Map<string, SheetNeighbours>();
    removed.forEach((name, i) => {
      if (neighbours[i]) beside.set(foldSheetName(name), neighbours[i]);
    });
    map.runs = {
      neighbour: (name, step) => {
        const entry = beside.get(foldSheetName(name));
        const sheet = entry && (step > 0 ? entry.next : entry.previous);
        return sheet && sheets.indexOf(sheet) >= 0 ? sheet.name : undefined;
      },
      index: (name) => {
        const wanted = foldSheetName(name);
        for (let i = 0; i < sheets.length; i++) if (foldSheetName(sheets[i].name) === wanted) return i;
        return -1;
      },
    };
  }
  return map;
}

/** What one sheet name of the file stands for now; sheetNameMap() is the one to use for many names. */
export function currentSheetName(workbook: Workbook, nameInFile: string): string | null | undefined {
  const map = sheetNameMap(workbook);
  return map ? map(nameInFile) : undefined;
}

export function createSheet(init: Pick<Sheet, 'name' | 'sheetId' | 'state' | 'kind' | 'part'>): Sheet {
  return {
    ...init,
    rows: new Map(),
    columns: [],
    merges: [],
    hyperlinks: [],
    rowCount: 0,
    colCount: 0,
    view: {
      showGridLines: true,
      showRowColHeaders: true,
      showZeros: true,
      rightToLeft: false,
      tabSelected: false,
      zoomScale: 100,
      frozenRows: 0,
      frozenCols: 0,
    },
  };
}

export function emptyStyleSheet(): StyleSheet {
  const font: Font = { name: 'Calibri', size: 11, bold: false, italic: false, strike: false, underline: 'none' };
  const border: Border = { diagonalUp: false, diagonalDown: false };
  const format: CellFormat = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false };
  return {
    numFmts: new Map(),
    fonts: [font],
    fills: [{ type: 'pattern', pattern: 'none' }],
    borders: [border],
    cellStyleXfs: [{ ...format }],
    cellXfs: [{ ...format, xfId: 0 }],
    cellStyles: [],
    dxfs: [],
    themeColors: [],
  };
}

/** Column index -> letters: 0 -> "A", 25 -> "Z", 26 -> "AA". */
export function columnName(col: number): string {
  let name = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
}

export function cellName(row: number, col: number): string {
  return columnName(col) + (row + 1);
}

/**
 * Parses "B7" (also "$B$7") into 0-based indexes. Returns undefined for
 * anything that is not a cell reference inside the sheet limits.
 */
export function parseCellRef(ref: string): { row: number; col: number } | undefined {
  let i = 0;
  const n = ref.length;
  if (ref.charCodeAt(i) === 36) i++;
  let col = 0;
  const lettersStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i) & ~32;
    if (c < 65 || c > 90) break;
    col = col * 26 + (c - 64);
    if (col > MAX_COLS) return undefined;
  }
  if (i === lettersStart) return undefined;
  if (ref.charCodeAt(i) === 36) i++;
  let row = 0;
  const digitsStart = i;
  for (; i < n; i++) {
    const c = ref.charCodeAt(i);
    if (c < 48 || c > 57) return undefined;
    row = row * 10 + (c - 48);
    if (row > MAX_ROWS) return undefined;
  }
  if (i === digitsStart || row === 0) return undefined;
  return { row: row - 1, col: col - 1 };
}

/** Parses "A1:C3" or a single "A1"; corners may come in any order. */
export function parseRange(ref: string): CellRange | undefined {
  const colon = ref.indexOf(':');
  const first = parseCellRef(colon < 0 ? ref : ref.slice(0, colon));
  const second = colon < 0 ? first : parseCellRef(ref.slice(colon + 1));
  if (!first || !second) return undefined;
  return {
    top: Math.min(first.row, second.row),
    left: Math.min(first.col, second.col),
    bottom: Math.max(first.row, second.row),
    right: Math.max(first.col, second.col),
  };
}

export function rangeName(range: CellRange): string {
  const first = cellName(range.top, range.left);
  const last = cellName(range.bottom, range.right);
  return first === last ? first : `${first}:${last}`;
}

export function getCell(sheet: Sheet, row: number, col: number): Cell | undefined {
  return sheet.rows.get(row)?.cells[col];
}
