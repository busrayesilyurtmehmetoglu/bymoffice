// Property elements of WordprocessingML (rPr, pPr, tblPr, trPr, tcPr, sectPr)
// -> the property bags of the document model.
//
// Each reader walks the children of a property element. A child it does not
// know, a value it cannot accept (not a number, out of range) and an attribute
// it has no field for all leave the child in `extras`, so nothing the file
// says is dropped; so does a comment, a processing instruction or stray text
// inside a child. "start"/"end" spellings of edges are read as left/right,
// the names the model uses for the leading and trailing edge, and the model
// notes that the file used them (`logical`); a child that states an edge in
// both spellings is kept, as the model holds only one of the two values.

import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  type Border,
  type CellMargins,
  type CellProps,
  TABLE_LOOK_BITS,
  type DocColor,
  type ExtraXml,
  type FrameProps,
  type HeaderFooterKind,
  type Indent,
  type NoteProps,
  type PageBorders,
  type ParaBorders,
  type ParaProps,
  type ParaSpacing,
  type RevisionInfo,
  type RowProps,
  type RunFonts,
  type RunProps,
  type SectionProps,
  type Shading,
  type TabStop,
  type TableBorders,
  type TableLook,
  type TablePosition,
  type TableProps,
  type TableWidth,
} from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import {
  MAX_TWIPS,
  attr,
  attrList,
  boolAttr,
  elements,
  hasContent,
  hasLooseContent,
  intAttr,
  keep,
  kindOf,
  measureAttr,
  onOff,
  parseMeasure,
  pushExtra,
  relAttr,
  serializeElement,
  twipsAttr,
  utwipsAttr,
  type PartContext,
} from './common';

/** What a child handler found. */
const UNKNOWN = 0; // not understood: kept whole
const LEAF = 1; // understood, provided every attribute was read and the element is empty
const DONE = 2; // understood, child elements included
const PARTIAL = 3; // understood in part: kept whole and marked
const KEPT = 4; // held whole elsewhere in the model (a tree, or a property bag of its own)

type Handler = (kind: string, el: XmlElement) => number;

/** Runs the handler over the element children and collects what it did not fully understand. */
function readChildren(ctx: PartContext, parent: XmlElement, context: string, handle: Handler): ExtraXml[] | undefined {
  let extras: ExtraXml[] | undefined;
  let index = -1;
  for (const node of parent.children) {
    if (node.type !== 'element') continue;
    index++;
    ctx.hits = 0;
    const status = handle(kindOf(ctx, node.name), node);
    let partial = status === PARTIAL;
    if (status === LEAF && (ctx.hits !== node.attrs.size || hasContent(node))) partial = true;
    // Comments and the like between the children of a container the model spreads over its fields.
    else if (status === DONE && hasLooseContent(node)) partial = true;
    if (status === UNKNOWN || partial) extras = pushExtra(ctx, extras, index, node, context, partial);
  }
  return extras;
}

/** After a nested leaf was read with the counter reset: was all of it understood? */
const complete = (ctx: PartContext, el: XmlElement) => ctx.hits === el.attrs.size && !hasContent(el);

const HEX6 = /^[0-9A-Fa-f]{6}$/;
const HEX2 = /^[0-9A-Fa-f]{2}$/;
const TOKEN = /^[A-Za-z0-9_-]{1,40}$/;

/** An enumeration value: a short plain token. Anything else is treated as not understood. */
function tokenAttr(ctx: PartContext, el: XmlElement, local: string): string | undefined {
  const text = attr(ctx, el, local);
  if (text === undefined) return undefined;
  if (TOKEN.test(text)) return text;
  ctx.hits--;
  return undefined;
}

function hexByteAttr(ctx: PartContext, el: XmlElement, local: string): number | undefined {
  const text = attr(ctx, el, local);
  if (text === undefined) return undefined;
  if (HEX2.test(text)) return parseInt(text, 16);
  ctx.hits--;
  return undefined;
}

/** A colour given by up to four attributes: the value, the theme colour, its tint and its shade. */
function readColor(ctx: PartContext, el: XmlElement, val: string, theme: string, tint: string, shade: string): DocColor | undefined {
  const color: DocColor = {};
  let any = false;
  const text = attr(ctx, el, val);
  if (text !== undefined) {
    if (text === 'auto') color.auto = any = true;
    else if (HEX6.test(text)) {
      color.rgb = text.toUpperCase();
      any = true;
    } else ctx.hits--;
  }
  const themeName = tokenAttr(ctx, el, theme);
  if (themeName !== undefined) {
    color.theme = themeName;
    any = true;
  }
  const tintValue = hexByteAttr(ctx, el, tint);
  if (tintValue !== undefined) color.tint = tintValue;
  const shadeValue = hexByteAttr(ctx, el, shade);
  if (shadeValue !== undefined) color.shade = shadeValue;
  return any ? color : undefined;
}

/** CT_Border. Undefined when the element has no usable style. */
function readBorder(ctx: PartContext, el: XmlElement): Border | undefined {
  const style = tokenAttr(ctx, el, 'val');
  if (style === undefined) return undefined;
  const border: Border = { style };
  const size = measureAttr(ctx, el, 'sz', 8, 0, 1000);
  if (size !== undefined) border.size = size;
  const space = measureAttr(ctx, el, 'space', 1, 0, 1000);
  if (space !== undefined) border.space = space;
  const color = readColor(ctx, el, 'color', 'themeColor', 'themeTint', 'themeShade');
  if (color) border.color = color;
  const shadow = boolAttr(ctx, el, 'shadow');
  if (shadow !== undefined) border.shadow = shadow;
  const frame = boolAttr(ctx, el, 'frame');
  if (frame !== undefined) border.frame = frame;
  return border;
}

/** CT_Shd. */
function readShading(ctx: PartContext, el: XmlElement): Shading | undefined {
  const pattern = tokenAttr(ctx, el, 'val');
  const shading: Shading = {};
  if (pattern !== undefined) shading.pattern = pattern;
  const fill = readColor(ctx, el, 'fill', 'themeFill', 'themeFillTint', 'themeFillShade');
  if (fill) shading.fill = fill;
  const color = readColor(ctx, el, 'color', 'themeColor', 'themeTint', 'themeShade');
  if (color) shading.color = color;
  return pattern === undefined && !fill && !color ? undefined : shading;
}

const EDGE_NAMES = new Map<string, string>([
  ['top', 'top'],
  ['left', 'left'],
  ['start', 'left'],
  ['bottom', 'bottom'],
  ['right', 'right'],
  ['end', 'right'],
  ['insideH', 'insideH'],
  ['insideV', 'insideV'],
  ['between', 'between'],
  ['bar', 'bar'],
  ['tl2br', 'tl2br'],
  ['tr2bl', 'tr2bl'],
]);

/** Children of a border container into `target`; false when one of them was not read in full. */
function readEdges(ctx: PartContext, el: XmlElement, target: Record<string, Border | undefined>, allowed: readonly string[]): boolean {
  let whole = true;
  for (const child of elements(el)) {
    const name = kindOf(ctx, child.name);
    const key = EDGE_NAMES.get(name);
    ctx.hits = 0;
    const border = key !== undefined && allowed.indexOf(key) >= 0 ? readBorder(ctx, child) : undefined;
    // An edge stated twice (also as <w:left> and <w:start>): the first is read, the container is kept for the other.
    if (!border || key === undefined || target[key] !== undefined || !complete(ctx, child)) whole = false;
    if (border && key !== undefined && target[key] === undefined) {
      if (name !== key) border.logical = true;
      target[key] = border;
    }
  }
  return whole;
}

/** Children of a border container (pBdr, tblBorders, tcBorders) into `target`. */
function readBorders(ctx: PartContext, el: XmlElement, target: Record<string, Border | undefined>, allowed: readonly string[]): number {
  const whole = readEdges(ctx, el, target, allowed);
  return whole && !el.attrs.size ? DONE : PARTIAL;
}

const PAGE_EDGES = ['top', 'left', 'bottom', 'right'];
const PARA_EDGES = ['top', 'left', 'bottom', 'right', 'between', 'bar'];
const TABLE_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
const CELL_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];

/** CT_TblWidth: w + type. Strict documents may write a percentage as "50%". */
function readTableWidth(ctx: PartContext, el: XmlElement): TableWidth | undefined {
  const typeText = attr(ctx, el, 'type');
  const text = attr(ctx, el, 'w');
  if (typeText === 'auto' || typeText === 'nil') {
    // The value means nothing with these types, but Word writes one and it goes back as written.
    const value = text === undefined ? 0 : parseMeasure(text, 20, -MAX_TWIPS, MAX_TWIPS);
    if (value === undefined) ctx.hits--;
    return { type: typeText, value: value === undefined ? 0 : value };
  }
  if (text === undefined) {
    if (typeText !== undefined) ctx.hits--;
    return undefined;
  }
  if (/^-?\d{1,6}(\.\d{1,6})?%$/.test(text)) return { type: 'pct', value: Math.round(Number(text.slice(0, -1)) * 50) };
  if (typeText !== undefined && typeText !== 'dxa' && typeText !== 'pct') {
    ctx.hits -= 2;
    return undefined;
  }
  const value = parseMeasure(text, typeText === 'pct' ? 1 : 20, -MAX_TWIPS, MAX_TWIPS);
  if (value === undefined) {
    ctx.hits -= typeText === undefined ? 1 : 2;
    return undefined;
  }
  return { type: typeText === 'pct' ? 'pct' : 'dxa', value };
}

function readCellMargins(ctx: PartContext, el: XmlElement, target: CellMargins): number {
  let status = el.attrs.size ? PARTIAL : DONE;
  for (const child of elements(el)) {
    const name = kindOf(ctx, child.name);
    const key = EDGE_NAMES.get(name);
    ctx.hits = 0;
    const width =
      key === 'top' || key === 'left' || key === 'bottom' || key === 'right' ? readTableWidth(ctx, child) : undefined;
    const edge = key as keyof CellMargins;
    if (!width || target[edge] !== undefined || !complete(ctx, child)) status = PARTIAL;
    if (width && target[edge] === undefined) {
      if (name !== key) width.logical = true;
      target[edge] = width;
    }
  }
  return status;
}

/** Author, date and id of a tracked change element. */
export function readRevision(ctx: PartContext, el: XmlElement): RevisionInfo {
  const info: RevisionInfo = {};
  const id = intAttr(ctx, el, 'id', -0x80000000, 0x7fffffff);
  if (id !== undefined) info.id = id;
  const author = attr(ctx, el, 'author');
  if (author !== undefined) info.author = author;
  const date = attr(ctx, el, 'date');
  if (date !== undefined) info.date = date;
  return info;
}

const RUN_ON_OFF = new Map<string, keyof RunProps>([
  ['b', 'bold'],
  ['bCs', 'boldCs'],
  ['i', 'italic'],
  ['iCs', 'italicCs'],
  ['caps', 'caps'],
  ['smallCaps', 'smallCaps'],
  ['strike', 'strike'],
  ['dstrike', 'doubleStrike'],
  ['outline', 'outline'],
  ['shadow', 'shadow'],
  ['emboss', 'emboss'],
  ['imprint', 'imprint'],
  ['vanish', 'vanish'],
  ['webHidden', 'webHidden'],
  ['specVanish', 'specVanish'],
  ['noProof', 'noProof'],
  ['snapToGrid', 'snapToGrid'],
  ['rtl', 'rtl'],
  ['cs', 'complexScript'],
]);

const FONT_ATTRS: [attribute: string, field: keyof RunFonts][] = [
  ['ascii', 'ascii'],
  ['hAnsi', 'hAnsi'],
  ['eastAsia', 'eastAsia'],
  ['cs', 'cs'],
  ['asciiTheme', 'asciiTheme'],
  ['hAnsiTheme', 'hAnsiTheme'],
  ['eastAsiaTheme', 'eastAsiaTheme'],
  // The schema spells this one in lower case.
  ['cstheme', 'csTheme'],
  ['hint', 'hint'],
];

const REVISION_MARKS = new Map<string, 'inserted' | 'deleted' | 'moveFrom' | 'moveTo'>([
  ['ins', 'inserted'],
  ['del', 'deleted'],
  ['moveFrom', 'moveFrom'],
  ['moveTo', 'moveTo'],
]);

function parseRunProps(ctx: PartContext, el: XmlElement): RunProps {
  const p: RunProps = {};
  const flags = p as Record<string, boolean | undefined>;
  const extras = readChildren(ctx, el, 'rPr', (kind, child) => {
    const flag = RUN_ON_OFF.get(kind);
    if (flag !== undefined) {
      const value = onOff(ctx, child);
      if (value === undefined || flags[flag] !== undefined) return UNKNOWN;
      flags[flag] = value;
      return LEAF;
    }
    const mark = REVISION_MARKS.get(kind);
    if (mark !== undefined) {
      if (p[mark]) return UNKNOWN;
      p[mark] = readRevision(ctx, child);
      return LEAF;
    }
    switch (kind) {
      case 'rStyle': {
        const value = attr(ctx, child, 'val');
        if (value === undefined || p.style !== undefined) return UNKNOWN;
        p.style = value;
        return LEAF;
      }
      case 'rFonts': {
        if (p.fonts) return UNKNOWN;
        const fonts: RunFonts = {};
        for (const [name, field] of FONT_ATTRS) {
          const value = attr(ctx, child, name);
          if (value !== undefined) fonts[field] = value;
        }
        p.fonts = fonts;
        return LEAF;
      }
      case 'sz':
      case 'szCs':
      case 'kern': {
        const value = measureAttr(ctx, child, 'val', 2, 0, 7200);
        const field = kind === 'sz' ? 'size' : kind === 'szCs' ? 'sizeCs' : 'kerning';
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'u': {
        if (p.underline) return UNKNOWN;
        const underline: NonNullable<RunProps['underline']> = {};
        const style = tokenAttr(ctx, child, 'val');
        if (style !== undefined) underline.style = style;
        const color = readColor(ctx, child, 'color', 'themeColor', 'themeTint', 'themeShade');
        if (color) underline.color = color;
        p.underline = underline;
        return LEAF;
      }
      case 'color': {
        const color = readColor(ctx, child, 'val', 'themeColor', 'themeTint', 'themeShade');
        if (!color || p.color) return UNKNOWN;
        p.color = color;
        return LEAF;
      }
      case 'highlight': {
        const value = tokenAttr(ctx, child, 'val');
        if (value === undefined || p.highlight !== undefined) return UNKNOWN;
        p.highlight = value;
        return LEAF;
      }
      case 'shd': {
        const shading = readShading(ctx, child);
        if (!shading || p.shading) return UNKNOWN;
        p.shading = shading;
        return LEAF;
      }
      case 'bdr': {
        const border = readBorder(ctx, child);
        if (!border || p.border) return UNKNOWN;
        p.border = border;
        return LEAF;
      }
      case 'vertAlign': {
        const value = attr(ctx, child, 'val');
        if ((value !== 'baseline' && value !== 'superscript' && value !== 'subscript') || p.vertAlign) return UNKNOWN;
        p.vertAlign = value;
        return LEAF;
      }
      case 'position': {
        const value = measureAttr(ctx, child, 'val', 2, -7200, 7200);
        if (value === undefined || p.position !== undefined) return UNKNOWN;
        p.position = value;
        return LEAF;
      }
      case 'spacing': {
        const value = twipsAttr(ctx, child, 'val');
        if (value === undefined || p.spacing !== undefined) return UNKNOWN;
        p.spacing = value;
        return LEAF;
      }
      case 'w': {
        const text = attr(ctx, child, 'val');
        const m = text === undefined ? null : /^(\d{1,3})%?$/.exec(text);
        if (!m || Number(m[1]) < 1 || Number(m[1]) > 600 || p.scale !== undefined) return UNKNOWN;
        p.scale = Number(m[1]);
        return LEAF;
      }
      case 'lang': {
        if (p.lang) return UNKNOWN;
        p.lang = {};
        for (const name of ['val', 'eastAsia', 'bidi'] as const) {
          const value = attr(ctx, child, name);
          if (value !== undefined) p.lang[name] = value;
        }
        return LEAF;
      }
      case 'em':
      case 'effect': {
        const value = tokenAttr(ctx, child, 'val');
        const field = kind === 'em' ? 'emphasis' : 'effect';
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'rPrChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'rPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  return p;
}

/** <w:rPr> -> RunProps. Equal elements give the same object. */
export function readRunProps(ctx: PartContext, el: XmlElement | undefined): RunProps {
  if (!el || (el.children.length === 0 && el.attrs.size === 0)) return EMPTY_RUN_PROPS;
  const cache = ctx.state.runProps;
  // The part is in the key because relationship ids and prefixes mean something else in another part.
  const key = `${ctx.part}\n${serializeElement(el)}`;
  let props = cache.get(key);
  if (!props) {
    props = parseRunProps(ctx, el);
    if (el.attrs.size || hasLooseContent(el)) props.extras = pushExtra(ctx, props.extras, -1, el, 'rPr', true);
    cache.set(key, props);
  }
  return props;
}

const PARA_ON_OFF = new Map<string, keyof ParaProps>(
  (
    [
      'keepNext',
      'keepLines',
      'widowControl',
      'pageBreakBefore',
      'suppressLineNumbers',
      'suppressAutoHyphens',
      'bidi',
      'mirrorIndents',
      'snapToGrid',
      'wordWrap',
      'adjustRightInd',
      'autoSpaceDE',
      'autoSpaceDN',
      'kinsoku',
      'overflowPunct',
      'topLinePunct',
      'contextualSpacing',
    ] as const
  ).map((name) => [name, name]),
);

/**
 * <w:ind>. `whole` is false when the model cannot say what the element says:
 * an edge given in both spellings (the first is read), or one edge by side
 * and another by reading direction.
 */
function readIndent(ctx: PartContext, el: XmlElement): { indent: Indent; whole: boolean } {
  const indent: Indent = {};
  let bySide = false;
  let byDirection = false;
  let twice = false;
  const edge = (field: 'left' | 'right' | 'leftChars' | 'rightChars', side: number | undefined, direction: number | undefined) => {
    if (side !== undefined) {
      indent[field] = side;
      bySide = true;
      if (direction !== undefined) twice = true;
    } else if (direction !== undefined) {
      indent[field] = direction;
      byDirection = true;
    }
  };
  edge('left', twipsAttr(ctx, el, 'left'), twipsAttr(ctx, el, 'start'));
  edge('right', twipsAttr(ctx, el, 'right'), twipsAttr(ctx, el, 'end'));
  const firstLine = utwipsAttr(ctx, el, 'firstLine');
  if (firstLine !== undefined) indent.firstLine = firstLine;
  const hanging = utwipsAttr(ctx, el, 'hanging');
  if (hanging !== undefined) indent.hanging = hanging;
  edge('leftChars', intAttr(ctx, el, 'leftChars', -100000, 100000), intAttr(ctx, el, 'startChars', -100000, 100000));
  edge('rightChars', intAttr(ctx, el, 'rightChars', -100000, 100000), intAttr(ctx, el, 'endChars', -100000, 100000));
  const firstLineChars = intAttr(ctx, el, 'firstLineChars', 0, 100000);
  if (firstLineChars !== undefined) indent.firstLineChars = firstLineChars;
  const hangingChars = intAttr(ctx, el, 'hangingChars', 0, 100000);
  if (hangingChars !== undefined) indent.hangingChars = hangingChars;
  const whole = !twice && !(bySide && byDirection);
  if (whole && byDirection) indent.logical = true;
  return { indent, whole };
}

const HEIGHT_RULES = ['auto', 'exact', 'atLeast'];

function readFrame(ctx: PartContext, el: XmlElement): FrameProps {
  const frame: FrameProps = {};
  const dropCap = tokenAttr(ctx, el, 'dropCap');
  if (dropCap !== undefined) frame.dropCap = dropCap;
  const lines = intAttr(ctx, el, 'lines', 1, 1000);
  if (lines !== undefined) frame.lines = lines;
  const width = utwipsAttr(ctx, el, 'w');
  if (width !== undefined) frame.width = width;
  const height = utwipsAttr(ctx, el, 'h');
  if (height !== undefined) frame.height = height;
  const rule = attr(ctx, el, 'hRule');
  if (rule !== undefined && HEIGHT_RULES.indexOf(rule) >= 0) frame.heightRule = rule as FrameProps['heightRule'];
  else if (rule !== undefined) ctx.hits--;
  for (const name of ['hSpace', 'vSpace'] as const) {
    const value = utwipsAttr(ctx, el, name);
    if (value !== undefined) frame[name] = value;
  }
  for (const name of ['wrap', 'hAnchor', 'vAnchor', 'xAlign', 'yAlign'] as const) {
    const value = tokenAttr(ctx, el, name);
    if (value !== undefined) frame[name] = value;
  }
  for (const name of ['x', 'y'] as const) {
    const value = twipsAttr(ctx, el, name);
    if (value !== undefined) frame[name] = value;
  }
  const anchorLock = boolAttr(ctx, el, 'anchorLock');
  if (anchorLock !== undefined) frame.anchorLock = anchorLock;
  return frame;
}

function readSpacing(ctx: PartContext, el: XmlElement): ParaSpacing {
  const spacing: ParaSpacing = {};
  const before = utwipsAttr(ctx, el, 'before');
  if (before !== undefined) spacing.before = before;
  const after = utwipsAttr(ctx, el, 'after');
  if (after !== undefined) spacing.after = after;
  const line = twipsAttr(ctx, el, 'line');
  if (line !== undefined) spacing.line = line;
  const rule = attr(ctx, el, 'lineRule');
  if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') spacing.lineRule = rule;
  else if (rule !== undefined) ctx.hits--;
  const beforeAuto = boolAttr(ctx, el, 'beforeAutospacing');
  if (beforeAuto !== undefined) spacing.beforeAuto = beforeAuto;
  const afterAuto = boolAttr(ctx, el, 'afterAutospacing');
  if (afterAuto !== undefined) spacing.afterAuto = afterAuto;
  const beforeLines = intAttr(ctx, el, 'beforeLines', 0, 100000);
  if (beforeLines !== undefined) spacing.beforeLines = beforeLines;
  const afterLines = intAttr(ctx, el, 'afterLines', 0, 100000);
  if (afterLines !== undefined) spacing.afterLines = afterLines;
  return spacing;
}

function readTabs(ctx: PartContext, el: XmlElement, tabs: TabStop[]): number {
  let status = el.attrs.size ? PARTIAL : DONE;
  for (const child of elements(el)) {
    ctx.hits = 0;
    if (kindOf(ctx, child.name) !== 'tab') {
      status = PARTIAL;
      continue;
    }
    const align = tokenAttr(ctx, child, 'val');
    const pos = twipsAttr(ctx, child, 'pos');
    const leader = tokenAttr(ctx, child, 'leader');
    if (align === undefined || pos === undefined) {
      status = PARTIAL;
      continue;
    }
    const tab: TabStop = { pos, align };
    if (leader !== undefined) tab.leader = leader;
    tabs.push(tab);
    if (!complete(ctx, child)) status = PARTIAL;
  }
  return status;
}

function parseParaProps(ctx: PartContext, el: XmlElement): ParaProps {
  const p: ParaProps = {};
  const flags = p as Record<string, boolean | undefined>;
  const extras = readChildren(ctx, el, 'pPr', (kind, child) => {
    const flag = PARA_ON_OFF.get(kind);
    if (flag !== undefined) {
      const value = onOff(ctx, child);
      if (value === undefined || flags[flag] !== undefined) return UNKNOWN;
      flags[flag] = value;
      return LEAF;
    }
    switch (kind) {
      case 'pStyle': {
        const value = attr(ctx, child, 'val');
        if (value === undefined || p.style !== undefined) return UNKNOWN;
        p.style = value;
        return LEAF;
      }
      case 'jc':
      case 'textAlignment':
      case 'textDirection': {
        const value = tokenAttr(ctx, child, 'val');
        const field = kind === 'jc' ? 'align' : kind;
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'ind': {
        if (p.indent) return UNKNOWN;
        const { indent, whole } = readIndent(ctx, child);
        p.indent = indent;
        return whole ? LEAF : PARTIAL;
      }
      case 'framePr':
        if (p.frame) return UNKNOWN;
        p.frame = readFrame(ctx, child);
        return LEAF;
      case 'spacing':
        if (p.spacing) return UNKNOWN;
        p.spacing = readSpacing(ctx, child);
        return LEAF;
      case 'tabs': {
        if (p.tabs) return UNKNOWN;
        p.tabs = [];
        return readTabs(ctx, child, p.tabs);
      }
      case 'numPr': {
        if (p.numbering) return UNKNOWN;
        const numbering: NonNullable<ParaProps['numbering']> = {};
        let status = child.attrs.size ? PARTIAL : DONE;
        for (const item of elements(child)) {
          const name = kindOf(ctx, item.name);
          ctx.hits = 0;
          if (name === 'numId' && numbering.numId === undefined) {
            const numId = intAttr(ctx, item, 'val', 0, 0x7fffffff);
            if (numId !== undefined) numbering.numId = numId;
            if (numId === undefined || !complete(ctx, item)) status = PARTIAL;
          } else if (name === 'ilvl' && numbering.level === undefined) {
            const level = intAttr(ctx, item, 'val', 0, 8);
            if (level !== undefined) numbering.level = level;
            if (level === undefined || !complete(ctx, item)) status = PARTIAL;
          } else status = PARTIAL;
        }
        p.numbering = numbering;
        return status;
      }
      case 'pBdr': {
        if (p.borders) return UNKNOWN;
        const borders: ParaBorders = {};
        p.borders = borders;
        return readBorders(ctx, child, borders as Record<string, Border | undefined>, PARA_EDGES);
      }
      case 'shd': {
        const shading = readShading(ctx, child);
        if (!shading || p.shading) return UNKNOWN;
        p.shading = shading;
        return LEAF;
      }
      case 'outlineLvl': {
        const value = intAttr(ctx, child, 'val', 0, 9);
        if (value === undefined || p.outlineLevel !== undefined) return UNKNOWN;
        p.outlineLevel = value;
        return LEAF;
      }
      case 'cnfStyle': {
        const value = readCnfStyle(ctx, child);
        if (value === undefined || p.cnfStyle !== undefined) return UNKNOWN;
        p.cnfStyle = value;
        return LEAF;
      }
      case 'rPr':
        if (p.markProps) return UNKNOWN;
        p.markProps = readRunProps(ctx, child);
        return KEPT;
      case 'sectPr':
        if (p.section) return UNKNOWN;
        p.section = readSectionProps(ctx, child);
        return KEPT;
      case 'pPrChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'pPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  return p;
}

/** <w:pPr> -> ParaProps. Equal elements give the same object. */
export function readParaProps(ctx: PartContext, el: XmlElement | undefined): ParaProps {
  if (!el || (el.children.length === 0 && el.attrs.size === 0)) return EMPTY_PARA_PROPS;
  const cache = ctx.state.paraProps;
  const key = `${ctx.part}\n${serializeElement(el)}`;
  let props = cache.get(key);
  if (!props) {
    props = parseParaProps(ctx, el);
    if (el.attrs.size || hasLooseContent(el)) props.extras = pushExtra(ctx, props.extras, -1, el, 'pPr', true);
    cache.set(key, props);
  }
  return props;
}

const CNF_ATTRS = [
  'firstRow',
  'lastRow',
  'firstColumn',
  'lastColumn',
  'oddVBand',
  'evenVBand',
  'oddHBand',
  'evenHBand',
  'firstRowFirstColumn',
  'firstRowLastColumn',
  'lastRowFirstColumn',
  'lastRowLastColumn',
];

/** <w:cnfStyle>: the 12 flags as a string of digits, from w:val or from the individual attributes. */
function readCnfStyle(ctx: PartContext, el: XmlElement): string | undefined {
  const val = attr(ctx, el, 'val');
  let bits = val !== undefined && /^[01]{12}$/.test(val) ? val : undefined;
  if (val !== undefined && bits === undefined) ctx.hits--;
  let fromAttrs = '';
  let any = false;
  for (const name of CNF_ATTRS) {
    const flag = boolAttr(ctx, el, name);
    if (flag !== undefined) any = true;
    fromAttrs += flag ? '1' : '0';
  }
  if (bits === undefined && any) bits = fromAttrs;
  return bits;
}

/**
 * <w:tblLook>: a bit field in w:val, an attribute per flag, or both; an
 * attribute wins over its bit. Which of the two the file used, and the bits
 * of w:val that are not flags, are kept for writing.
 */
function readTableLook(ctx: PartContext, el: XmlElement): TableLook | undefined {
  const text = attr(ctx, el, 'val');
  let val: number | undefined;
  if (text !== undefined) {
    if (/^[0-9A-Fa-f]{4}$/.test(text)) val = parseInt(text, 16);
    else ctx.hits--;
  }
  const bits = val === undefined ? 0 : val;
  let stated = 0;
  const flag = (name: keyof typeof TABLE_LOOK_BITS) => {
    const value = boolAttr(ctx, el, name);
    if (value === undefined) return (bits & TABLE_LOOK_BITS[name]) !== 0;
    stated |= TABLE_LOOK_BITS[name];
    return value;
  };
  const look: TableLook = {
    firstRow: flag('firstRow'),
    lastRow: flag('lastRow'),
    firstColumn: flag('firstColumn'),
    lastColumn: flag('lastColumn'),
    noHBand: flag('noHBand'),
    noVBand: flag('noVBand'),
  };
  if (!ctx.hits) return undefined;
  if (val !== undefined) look.val = val;
  if (stated) look.stated = stated;
  return look;
}

function readTablePosition(ctx: PartContext, el: XmlElement): TablePosition {
  const position: TablePosition = {};
  for (const name of ['leftFromText', 'rightFromText', 'topFromText', 'bottomFromText'] as const) {
    const value = utwipsAttr(ctx, el, name);
    if (value !== undefined) position[name] = value;
  }
  for (const name of ['vertAnchor', 'horzAnchor'] as const) {
    const value = tokenAttr(ctx, el, name);
    if (value !== undefined) position[name] = value;
  }
  const x = twipsAttr(ctx, el, 'tblpX');
  if (x !== undefined) position.x = x;
  const y = twipsAttr(ctx, el, 'tblpY');
  if (y !== undefined) position.y = y;
  const xAlign = tokenAttr(ctx, el, 'tblpXSpec');
  if (xAlign !== undefined) position.xAlign = xAlign;
  const yAlign = tokenAttr(ctx, el, 'tblpYSpec');
  if (yAlign !== undefined) position.yAlign = yAlign;
  return position;
}

/** <w:tblPr> or <w:tblPrEx> -> TableProps. */
export function readTableProps(ctx: PartContext, el: XmlElement | undefined): TableProps {
  const p: TableProps = {};
  if (!el) return p;
  const extras = readChildren(ctx, el, 'tblPr', (kind, child) => {
    switch (kind) {
      case 'tblStyle': {
        const value = attr(ctx, child, 'val');
        if (value === undefined || p.style !== undefined) return UNKNOWN;
        p.style = value;
        return LEAF;
      }
      case 'tblW':
      case 'tblInd':
      case 'tblCellSpacing': {
        const field = kind === 'tblW' ? 'width' : kind === 'tblInd' ? 'indent' : 'cellSpacing';
        const width = readTableWidth(ctx, child);
        if (!width || p[field]) return UNKNOWN;
        p[field] = width;
        return LEAF;
      }
      case 'jc':
      case 'tblOverlap': {
        const field = kind === 'jc' ? 'align' : 'overlap';
        const value = tokenAttr(ctx, child, 'val');
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'tblpPr':
        if (p.floating) return UNKNOWN;
        p.floating = readTablePosition(ctx, child);
        return LEAF;
      case 'tblLayout': {
        const value = attr(ctx, child, 'type');
        if ((value !== 'fixed' && value !== 'autofit') || p.layout) return UNKNOWN;
        p.layout = value;
        return LEAF;
      }
      case 'tblBorders': {
        if (p.borders) return UNKNOWN;
        const borders: TableBorders = {};
        p.borders = borders;
        return readBorders(ctx, child, borders as Record<string, Border | undefined>, TABLE_EDGES);
      }
      case 'shd': {
        const shading = readShading(ctx, child);
        if (!shading || p.shading) return UNKNOWN;
        p.shading = shading;
        return LEAF;
      }
      case 'tblCellMar':
        if (p.cellMargins) return UNKNOWN;
        p.cellMargins = {};
        return readCellMargins(ctx, child, p.cellMargins);
      case 'tblLook': {
        const look = readTableLook(ctx, child);
        if (!look || p.look) return UNKNOWN;
        p.look = look;
        return LEAF;
      }
      case 'tblStyleRowBandSize':
      case 'tblStyleColBandSize': {
        const field = kind === 'tblStyleRowBandSize' ? 'rowBandSize' : 'colBandSize';
        const value = intAttr(ctx, child, 'val', 0, 1000);
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'bidiVisual': {
        const value = onOff(ctx, child);
        if (value === undefined || p.bidiVisual !== undefined) return UNKNOWN;
        p.bidiVisual = value;
        return LEAF;
      }
      case 'tblCaption':
      case 'tblDescription': {
        const field = kind === 'tblCaption' ? 'caption' : 'description';
        const value = attr(ctx, child, 'val');
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'tblPrChange':
      case 'tblPrExChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'tblPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  if (el.attrs.size || hasLooseContent(el)) p.extras = pushExtra(ctx, p.extras, -1, el, 'tblPr', true);
  return p;
}

/** <w:trPr> -> RowProps. */
export function readRowProps(ctx: PartContext, el: XmlElement | undefined): RowProps {
  const p: RowProps = {};
  if (!el) return p;
  const extras = readChildren(ctx, el, 'trPr', (kind, child) => {
    switch (kind) {
      case 'trHeight': {
        if (p.height !== undefined || p.heightRule !== undefined) return UNKNOWN;
        const height = utwipsAttr(ctx, child, 'val');
        if (height !== undefined) p.height = height;
        const rule = attr(ctx, child, 'hRule');
        if (rule === 'auto' || rule === 'exact' || rule === 'atLeast') p.heightRule = rule;
        else if (rule !== undefined) ctx.hits--;
        return LEAF;
      }
      case 'tblHeader':
      case 'cantSplit':
      case 'hidden': {
        const field = kind === 'tblHeader' ? 'header' : kind;
        const value = onOff(ctx, child);
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'jc': {
        const value = tokenAttr(ctx, child, 'val');
        if (value === undefined || p.align !== undefined) return UNKNOWN;
        p.align = value;
        return LEAF;
      }
      case 'gridBefore':
      case 'gridAfter': {
        const value = intAttr(ctx, child, 'val', 0, 1000);
        if (value === undefined || p[kind] !== undefined) return UNKNOWN;
        p[kind] = value;
        return LEAF;
      }
      case 'wBefore':
      case 'wAfter':
      case 'tblCellSpacing': {
        const field = kind === 'wBefore' ? 'widthBefore' : kind === 'wAfter' ? 'widthAfter' : 'cellSpacing';
        const width = readTableWidth(ctx, child);
        if (!width || p[field]) return UNKNOWN;
        p[field] = width;
        return LEAF;
      }
      case 'cnfStyle': {
        const value = readCnfStyle(ctx, child);
        if (value === undefined || p.cnfStyle !== undefined) return UNKNOWN;
        p.cnfStyle = value;
        return LEAF;
      }
      case 'ins':
      case 'del': {
        const field = kind === 'ins' ? 'inserted' : 'deleted';
        if (p[field]) return UNKNOWN;
        p[field] = readRevision(ctx, child);
        return LEAF;
      }
      case 'trPrChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'trPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  if (el.attrs.size || hasLooseContent(el)) p.extras = pushExtra(ctx, p.extras, -1, el, 'trPr', true);
  return p;
}

/** <w:tcPr> -> CellProps. */
export function readCellProps(ctx: PartContext, el: XmlElement | undefined): CellProps {
  const p: CellProps = {};
  if (!el) return p;
  const extras = readChildren(ctx, el, 'tcPr', (kind, child) => {
    switch (kind) {
      case 'tcW': {
        const width = readTableWidth(ctx, child);
        if (!width || p.width) return UNKNOWN;
        p.width = width;
        return LEAF;
      }
      case 'gridSpan': {
        const value = intAttr(ctx, child, 'val', 1, 1000);
        if (value === undefined || p.gridSpan !== undefined) return UNKNOWN;
        p.gridSpan = value;
        return LEAF;
      }
      case 'vMerge':
      case 'hMerge': {
        const value = attr(ctx, child, 'val');
        if ((value !== undefined && value !== 'restart' && value !== 'continue') || p[kind]) return UNKNOWN;
        p[kind] = value === 'restart' ? 'restart' : 'continue';
        return LEAF;
      }
      case 'tcBorders': {
        if (p.borders) return UNKNOWN;
        const borders: TableBorders = {};
        p.borders = borders;
        return readBorders(ctx, child, borders as Record<string, Border | undefined>, CELL_EDGES);
      }
      case 'shd': {
        const shading = readShading(ctx, child);
        if (!shading || p.shading) return UNKNOWN;
        p.shading = shading;
        return LEAF;
      }
      case 'tcMar':
        if (p.margins) return UNKNOWN;
        p.margins = {};
        return readCellMargins(ctx, child, p.margins);
      case 'vAlign': {
        const value = attr(ctx, child, 'val');
        if ((value !== 'top' && value !== 'center' && value !== 'bottom' && value !== 'both') || p.vAlign) return UNKNOWN;
        p.vAlign = value;
        return LEAF;
      }
      case 'textDirection': {
        const value = tokenAttr(ctx, child, 'val');
        if (value === undefined || p.textDirection !== undefined) return UNKNOWN;
        p.textDirection = value;
        return LEAF;
      }
      case 'noWrap':
      case 'tcFitText':
      case 'hideMark': {
        const field = kind === 'tcFitText' ? 'fitText' : kind;
        const value = onOff(ctx, child);
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'cnfStyle': {
        const value = readCnfStyle(ctx, child);
        if (value === undefined || p.cnfStyle !== undefined) return UNKNOWN;
        p.cnfStyle = value;
        return LEAF;
      }
      case 'tcPrChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'tcPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  if (el.attrs.size || hasLooseContent(el)) p.extras = pushExtra(ctx, p.extras, -1, el, 'tcPr', true);
  return p;
}

const SECTION_TYPES = ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'];

/** <w:footnotePr> / <w:endnotePr> of a section into `note`. */
function readNoteProps(ctx: PartContext, el: XmlElement, note: NoteProps): number {
  let status = el.attrs.size ? PARTIAL : DONE;
  for (const child of elements(el)) {
    const kind = kindOf(ctx, child.name);
    ctx.hits = 0;
    let understood = false;
    if (kind === 'numStart' && note.start === undefined) {
      const value = intAttr(ctx, child, 'val', 0, 0x7fffffff);
      if (value !== undefined) {
        note.start = value;
        understood = true;
      }
    } else {
      const field = kind === 'pos' ? 'position' : kind === 'numFmt' ? 'format' : kind === 'numRestart' ? 'restart' : undefined;
      const value = field === undefined || note[field] !== undefined ? undefined : tokenAttr(ctx, child, 'val');
      if (field !== undefined && value !== undefined) {
        note[field] = value;
        understood = true;
      }
    }
    if (!understood || !complete(ctx, child)) status = PARTIAL;
  }
  return status;
}

/** <w:sectPr> -> SectionProps. */
export function readSectionProps(ctx: PartContext, el: XmlElement): SectionProps {
  const p: SectionProps = {};
  const attrs = attrList(el);
  if (attrs) p.attrs = attrs;
  const extras = readChildren(ctx, el, 'sectPr', (kind, child) => {
    switch (kind) {
      case 'headerReference':
      case 'footerReference': {
        const type = attr(ctx, child, 'type');
        const id = relAttr(ctx, child, 'id');
        if ((type !== 'default' && type !== 'first' && type !== 'even') || id === undefined) return UNKNOWN;
        const field = kind === 'headerReference' ? 'headers' : 'footers';
        const refs: Partial<Record<HeaderFooterKind, string>> = p[field] || (p[field] = {});
        if (refs[type] !== undefined) return UNKNOWN;
        refs[type] = id;
        return LEAF;
      }
      case 'type': {
        const value = attr(ctx, child, 'val');
        if (value === undefined || SECTION_TYPES.indexOf(value) < 0 || p.type) return UNKNOWN;
        p.type = value as SectionProps['type'];
        return LEAF;
      }
      case 'pgSz': {
        if (p.page) return UNKNOWN;
        const page: NonNullable<SectionProps['page']> = {};
        const width = utwipsAttr(ctx, child, 'w');
        if (width !== undefined) page.width = width;
        const height = utwipsAttr(ctx, child, 'h');
        if (height !== undefined) page.height = height;
        const orientation = attr(ctx, child, 'orient');
        if (orientation === 'portrait' || orientation === 'landscape') page.orientation = orientation;
        else if (orientation !== undefined) ctx.hits--;
        const code = intAttr(ctx, child, 'code', 0, 100000);
        if (code !== undefined) page.code = code;
        p.page = page;
        return LEAF;
      }
      case 'pgMar': {
        if (p.margins) return UNKNOWN;
        const margins: NonNullable<SectionProps['margins']> = {};
        // Top and bottom may be negative: the margin is then fixed and does not grow with the header.
        for (const name of ['top', 'bottom'] as const) {
          const value = twipsAttr(ctx, child, name);
          if (value !== undefined) margins[name] = value;
        }
        for (const name of ['right', 'left', 'header', 'footer', 'gutter'] as const) {
          const value = utwipsAttr(ctx, child, name);
          if (value !== undefined) margins[name] = value;
        }
        p.margins = margins;
        return LEAF;
      }
      case 'cols': {
        if (p.columns) return UNKNOWN;
        const columns: NonNullable<SectionProps['columns']> = {};
        const count = intAttr(ctx, child, 'num', 1, 45);
        if (count !== undefined) columns.count = count;
        const space = utwipsAttr(ctx, child, 'space');
        if (space !== undefined) columns.space = space;
        const equalWidth = boolAttr(ctx, child, 'equalWidth');
        if (equalWidth !== undefined) columns.equalWidth = equalWidth;
        const separator = boolAttr(ctx, child, 'sep');
        if (separator !== undefined) columns.separator = separator;
        p.columns = columns;
        let status = ctx.hits === child.attrs.size ? DONE : PARTIAL;
        for (const col of elements(child)) {
          ctx.hits = 0;
          if (kindOf(ctx, col.name) !== 'col') {
            status = PARTIAL;
            continue;
          }
          const column: { width?: number; space?: number } = {};
          const width = utwipsAttr(ctx, col, 'w');
          if (width !== undefined) column.width = width;
          const gap = utwipsAttr(ctx, col, 'space');
          if (gap !== undefined) column.space = gap;
          (columns.columns || (columns.columns = [])).push(column);
          if (!complete(ctx, col)) status = PARTIAL;
        }
        return status;
      }
      case 'titlePg':
      case 'bidi':
      case 'rtlGutter': {
        const field = kind === 'titlePg' ? 'titlePage' : kind;
        const value = onOff(ctx, child);
        if (value === undefined || p[field] !== undefined) return UNKNOWN;
        p[field] = value;
        return LEAF;
      }
      case 'pgNumType': {
        if (p.pageNumbering) return UNKNOWN;
        const numbering: NonNullable<SectionProps['pageNumbering']> = {};
        const format = tokenAttr(ctx, child, 'fmt');
        if (format !== undefined) numbering.format = format;
        const start = intAttr(ctx, child, 'start', 0, 0x7fffffff);
        if (start !== undefined) numbering.start = start;
        const chapStyle = intAttr(ctx, child, 'chapStyle', 0, 9);
        if (chapStyle !== undefined) numbering.chapStyle = chapStyle;
        const chapSep = tokenAttr(ctx, child, 'chapSep');
        if (chapSep !== undefined) numbering.chapSep = chapSep;
        p.pageNumbering = numbering;
        return LEAF;
      }
      case 'vAlign': {
        const value = attr(ctx, child, 'val');
        if ((value !== 'top' && value !== 'center' && value !== 'both' && value !== 'bottom') || p.vAlign) return UNKNOWN;
        p.vAlign = value;
        return LEAF;
      }
      case 'textDirection': {
        const value = tokenAttr(ctx, child, 'val');
        if (value === undefined || p.textDirection !== undefined) return UNKNOWN;
        p.textDirection = value;
        return LEAF;
      }
      case 'docGrid': {
        if (p.docGrid) return UNKNOWN;
        const grid: NonNullable<SectionProps['docGrid']> = {};
        const type = tokenAttr(ctx, child, 'type');
        if (type !== undefined) grid.type = type;
        const linePitch = intAttr(ctx, child, 'linePitch', -MAX_TWIPS, MAX_TWIPS);
        if (linePitch !== undefined) grid.linePitch = linePitch;
        const charSpace = intAttr(ctx, child, 'charSpace', -0x7fffffff, 0x7fffffff);
        if (charSpace !== undefined) grid.charSpace = charSpace;
        p.docGrid = grid;
        return LEAF;
      }
      case 'pgBorders': {
        if (p.pageBorders) return UNKNOWN;
        const borders: PageBorders = {};
        for (const name of ['offsetFrom', 'display', 'zOrder'] as const) {
          const value = tokenAttr(ctx, child, name);
          if (value !== undefined) borders[name] = value;
        }
        const attrsRead = ctx.hits === child.attrs.size;
        p.pageBorders = borders;
        // An art border names its image with a relationship id, which the model does not hold: kept.
        const whole = readEdges(ctx, child, borders as Record<string, Border | undefined>, PAGE_EDGES);
        return attrsRead && whole ? DONE : PARTIAL;
      }
      case 'lnNumType': {
        if (p.lineNumbering) return UNKNOWN;
        const numbering: NonNullable<SectionProps['lineNumbering']> = {};
        const countBy = intAttr(ctx, child, 'countBy', 0, 100000);
        if (countBy !== undefined) numbering.countBy = countBy;
        const start = intAttr(ctx, child, 'start', 0, 0x7fffffff);
        if (start !== undefined) numbering.start = start;
        const distance = utwipsAttr(ctx, child, 'distance');
        if (distance !== undefined) numbering.distance = distance;
        const restart = tokenAttr(ctx, child, 'restart');
        if (restart !== undefined) numbering.restart = restart;
        p.lineNumbering = numbering;
        return LEAF;
      }
      case 'footnotePr':
      case 'endnotePr': {
        const field = kind === 'footnotePr' ? 'footnotes' : 'endnotes';
        if (p[field]) return UNKNOWN;
        const note: NoteProps = {};
        p[field] = note;
        return readNoteProps(ctx, child, note);
      }
      case 'sectPrChange':
        if (p.change) return UNKNOWN;
        p.change = keep(ctx, child, 'sectPr');
        return KEPT;
    }
    return UNKNOWN;
  });
  if (extras) p.extras = extras;
  // The attributes are in `attrs`; a comment between the children needs the element itself.
  if (hasLooseContent(el)) p.extras = pushExtra(ctx, p.extras, -1, el, 'sectPr', true);
  return p;
}
