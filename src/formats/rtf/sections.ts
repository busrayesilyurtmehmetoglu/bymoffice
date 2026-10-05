// Document and section formatting of an RTF file -> the section properties
// and settings of the document model.
//
// The document states the paper and the margins once ("\paperw", "\margl",
// ...). "\sectd" sets the section back to those, and the "...sxn" words that
// follow change them for the one section. "\sect" ends a section: the
// paragraph it ends carries the properties, as in a docx file.

import type { Border, PageBorders, SectionProps } from '../../model/doc/types';
import { lookup, twipsOf } from './state';

/** What the document says for all of its sections. */
export interface Dop {
  paperWidth: number;
  paperHeight: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  gutter: number;
  landscape: boolean;
  facingPages: boolean;
  mirrorMargins: boolean;
  defaultTab: number;
  widowControl: boolean;
  autoHyphenation: boolean;
  trackRevisions: boolean;
  /** "\pgnstart": the number of the first page. */
  pageStart: number;
  /** What the document is protected for by the older words ("\formprot", "\revprot", "\annotprot", "\readprot"). */
  protection?: 'forms' | 'trackedChanges' | 'comments' | 'readOnly';
  /** "\protlevel": what an enforced protection is for. Alone it protects nothing. */
  protectionLevel?: 'forms' | 'trackedChanges' | 'comments' | 'readOnly';
  /** "\enforceprot": whether the protection applies; undefined when the file does not say. */
  enforceProtection?: boolean;
}

export function defaultDop(): Dop {
  return {
    paperWidth: 12240,
    paperHeight: 15840,
    marginLeft: 1800,
    marginRight: 1800,
    marginTop: 1440,
    marginBottom: 1440,
    gutter: 0,
    landscape: false,
    facingPages: false,
    mirrorMargins: false,
    defaultTab: 720,
    widowControl: false,
    autoHyphenation: false,
    trackRevisions: false,
    pageStart: 1,
  };
}

/** A control word of the document; false when it is not one. */
export function dopWord(dop: Dop, word: string, param: number, hasParam: boolean): boolean {
  const on = !hasParam || param !== 0;
  switch (word) {
    case 'paperw':
      dop.paperWidth = twipsOf(param);
      return true;
    case 'paperh':
      dop.paperHeight = twipsOf(param);
      return true;
    case 'margl':
      dop.marginLeft = twipsOf(param);
      return true;
    case 'margr':
      dop.marginRight = twipsOf(param);
      return true;
    case 'margt':
      dop.marginTop = twipsOf(param);
      return true;
    case 'margb':
      dop.marginBottom = twipsOf(param);
      return true;
    case 'gutter':
      dop.gutter = Math.max(0, twipsOf(param));
      return true;
    case 'landscape':
      dop.landscape = true;
      return true;
    case 'facingp':
      dop.facingPages = true;
      return true;
    case 'margmirror':
      dop.mirrorMargins = true;
      return true;
    case 'deftab':
      if (param > 0 && param <= 31680) dop.defaultTab = param;
      return true;
    case 'widowctrl':
      dop.widowControl = true;
      return true;
    case 'hyphauto':
      dop.autoHyphenation = on;
      return true;
    case 'revisions':
      dop.trackRevisions = true;
      return true;
    case 'pgnstart':
      dop.pageStart = param;
      return true;
    case 'formprot':
      dop.protection = 'forms';
      return true;
    case 'revprot':
      dop.protection = 'trackedChanges';
      return true;
    case 'annotprot':
      dop.protection = 'comments';
      return true;
    case 'readprot':
      dop.protection = 'readOnly';
      return true;
    case 'enforceprot':
      dop.enforceProtection = on;
      return true;
    case 'protlevel': {
      // The kind as newer writers state it, beside the older words above.
      const level = (['trackedChanges', 'comments', 'forms', 'readOnly'] as const)[param];
      if (hasParam && level !== undefined) dop.protectionLevel = level;
      return true;
    }
  }
  return false;
}

/** The formatting of the section being read. */
export interface Sep {
  /** Paper and margins of this section; undefined where it follows the document. */
  pageWidth: number | undefined;
  pageHeight: number | undefined;
  marginLeft: number | undefined;
  marginRight: number | undefined;
  marginTop: number | undefined;
  marginBottom: number | undefined;
  gutter: number | undefined;
  landscape: boolean;
  headerY: number;
  footerY: number;
  columns: number;
  columnSpace: number;
  columnLine: boolean;
  /** Columns of their own width: the one "\colno" names is being described. */
  columnList: { width?: number; space?: number }[];
  columnIndex: number;
  type: SectionProps['type'];
  pageStart: number;
  pageRestart: boolean;
  pageFormat: string | undefined;
  titlePage: boolean;
  vAlign: SectionProps['vAlign'];
  rtl: boolean;
  lineCount: number;
  lineDistance: number | undefined;
  lineStart: number | undefined;
  lineRestart: string | undefined;
  borders: PageBorders | undefined;
}

/** "\sectd": the section as the document describes it. */
export function defaultSep(): Sep {
  return {
    pageWidth: undefined,
    pageHeight: undefined,
    marginLeft: undefined,
    marginRight: undefined,
    marginTop: undefined,
    marginBottom: undefined,
    gutter: undefined,
    landscape: false,
    headerY: 720,
    footerY: 720,
    columns: 1,
    columnSpace: 720,
    columnLine: false,
    columnList: [],
    columnIndex: -1,
    type: undefined,
    pageStart: 1,
    pageRestart: false,
    pageFormat: undefined,
    titlePage: false,
    vAlign: undefined,
    rtl: false,
    lineCount: 0,
    lineDistance: undefined,
    lineStart: undefined,
    lineRestart: undefined,
    borders: undefined,
  };
}

/** Columns a section may have; Word stops at 45. */
const MAX_COLUMNS = 45;

const PAGE_BORDERS: { [word: string]: 'top' | 'left' | 'bottom' | 'right' } = lookup({ pgbrdrt: 'top', pgbrdrl: 'left', pgbrdrb: 'bottom', pgbrdrr: 'right' });

/** The side of the page a border word names; undefined when it names none. */
export const pageBorderSide = (word: string): 'top' | 'left' | 'bottom' | 'right' | undefined => PAGE_BORDERS[word];

export function setPageBorder(sep: Sep, side: 'top' | 'left' | 'bottom' | 'right', border: Border | undefined): void {
  const borders: PageBorders = { ...sep.borders };
  if (border) borders[side] = border;
  else delete borders[side];
  sep.borders = borders;
}

/** A control word of a section; false when it is not one. */
export function sepWord(sep: Sep, word: string, param: number, hasParam: boolean): boolean {
  switch (word) {
    case 'pgwsxn':
      sep.pageWidth = twipsOf(param);
      return true;
    case 'pghsxn':
      sep.pageHeight = twipsOf(param);
      return true;
    case 'marglsxn':
      sep.marginLeft = twipsOf(param);
      return true;
    case 'margrsxn':
      sep.marginRight = twipsOf(param);
      return true;
    case 'margtsxn':
      sep.marginTop = twipsOf(param);
      return true;
    case 'margbsxn':
      sep.marginBottom = twipsOf(param);
      return true;
    case 'guttersxn':
      sep.gutter = Math.max(0, twipsOf(param));
      return true;
    case 'lndscpsxn':
      sep.landscape = true;
      return true;
    case 'headery':
      sep.headerY = Math.max(0, twipsOf(param));
      return true;
    case 'footery':
      sep.footerY = Math.max(0, twipsOf(param));
      return true;
    case 'cols':
      sep.columns = Math.max(1, Math.min(param, MAX_COLUMNS));
      return true;
    case 'colsx':
      sep.columnSpace = Math.max(0, twipsOf(param));
      return true;
    case 'linebetcol':
      sep.columnLine = true;
      return true;
    case 'colno':
      sep.columnIndex = param >= 1 && param <= MAX_COLUMNS ? param - 1 : -1;
      return true;
    case 'colw':
    case 'colsr': {
      const index = sep.columnIndex;
      if (index < 0) return true;
      while (sep.columnList.length <= index) sep.columnList.push({});
      const column = { ...sep.columnList[index] };
      if (word === 'colw') column.width = Math.max(0, twipsOf(param));
      else column.space = Math.max(0, twipsOf(param));
      sep.columnList = sep.columnList.slice();
      sep.columnList[index] = column;
      return true;
    }
    case 'sbknone':
      sep.type = 'continuous';
      return true;
    case 'sbkcol':
      sep.type = 'nextColumn';
      return true;
    case 'sbkpage':
      sep.type = 'nextPage';
      return true;
    case 'sbkeven':
      sep.type = 'evenPage';
      return true;
    case 'sbkodd':
      sep.type = 'oddPage';
      return true;
    case 'pgnstarts':
      sep.pageStart = param;
      return true;
    case 'pgnrestart':
      sep.pageRestart = true;
      return true;
    case 'pgncont':
      sep.pageRestart = false;
      return true;
    case 'pgndec':
      sep.pageFormat = 'decimal';
      return true;
    case 'pgnucrm':
      sep.pageFormat = 'upperRoman';
      return true;
    case 'pgnlcrm':
      sep.pageFormat = 'lowerRoman';
      return true;
    case 'pgnucltr':
      sep.pageFormat = 'upperLetter';
      return true;
    case 'pgnlcltr':
      sep.pageFormat = 'lowerLetter';
      return true;
    case 'titlepg':
      sep.titlePage = !hasParam || param !== 0;
      return true;
    case 'vertalt':
      sep.vAlign = 'top';
      return true;
    case 'vertalc':
      sep.vAlign = 'center';
      return true;
    case 'vertalj':
      sep.vAlign = 'both';
      return true;
    case 'vertalb':
      sep.vAlign = 'bottom';
      return true;
    case 'rtlsect':
      sep.rtl = true;
      return true;
    case 'ltrsect':
      sep.rtl = false;
      return true;
    case 'linemod':
      sep.lineCount = Math.max(0, param);
      return true;
    case 'linex':
      sep.lineDistance = Math.max(0, twipsOf(param));
      return true;
    case 'linestarts':
      sep.lineStart = param;
      return true;
    case 'linerestart':
      sep.lineRestart = 'newSection';
      return true;
    case 'lineppage':
      sep.lineRestart = 'newPage';
      return true;
    case 'linecont':
      sep.lineRestart = 'continuous';
      return true;
  }
  return false;
}

/** The section properties of the model for a section; `first` says that it is the first of the document. */
export function sectionProps(sep: Sep, dop: Dop, first: boolean): SectionProps {
  const props: SectionProps = {};
  if (sep.type !== undefined) props.type = sep.type;
  const or = (value: number | undefined, fallback: number): number => (value === undefined ? fallback : value);
  props.page = { width: or(sep.pageWidth, dop.paperWidth), height: or(sep.pageHeight, dop.paperHeight) };
  if (sep.landscape || dop.landscape) props.page.orientation = 'landscape';
  props.margins = {
    top: or(sep.marginTop, dop.marginTop),
    right: or(sep.marginRight, dop.marginRight),
    bottom: or(sep.marginBottom, dop.marginBottom),
    left: or(sep.marginLeft, dop.marginLeft),
    header: sep.headerY,
    footer: sep.footerY,
    gutter: or(sep.gutter, dop.gutter),
  };
  if (sep.columns > 1 || sep.columnList.length) {
    props.columns = { count: sep.columns, space: sep.columnSpace };
    if (sep.columnLine) props.columns.separator = true;
    if (sep.columnList.length) {
      props.columns.equalWidth = false;
      props.columns.columns = sep.columnList.slice(0, sep.columns);
    }
  }
  if (sep.titlePage) props.titlePage = true;
  const numbering: NonNullable<SectionProps['pageNumbering']> = {};
  if (sep.pageFormat !== undefined && sep.pageFormat !== 'decimal') numbering.format = sep.pageFormat;
  if (sep.pageRestart) numbering.start = sep.pageStart;
  else if (first && dop.pageStart !== 1) numbering.start = dop.pageStart;
  if (numbering.format !== undefined || numbering.start !== undefined) props.pageNumbering = numbering;
  if (sep.vAlign !== undefined && sep.vAlign !== 'top') props.vAlign = sep.vAlign;
  if (sep.rtl) props.bidi = true;
  if (sep.lineCount > 0) {
    const lines: NonNullable<SectionProps['lineNumbering']> = { countBy: sep.lineCount };
    if (sep.lineDistance !== undefined) lines.distance = sep.lineDistance;
    if (sep.lineStart !== undefined) lines.start = sep.lineStart;
    if (sep.lineRestart !== undefined) lines.restart = sep.lineRestart;
    props.lineNumbering = lines;
  }
  if (sep.borders && Object.keys(sep.borders).length) props.pageBorders = sep.borders;
  return props;
}
