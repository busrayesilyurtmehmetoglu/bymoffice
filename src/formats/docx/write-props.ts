// Property bags of the document model -> the property elements of
// WordprocessingML (rPr, pPr, tblPr, tblPrEx, trPr, tcPr, sectPr): the reverse
// of props.ts.
//
// Word is sensitive to the order of the children of these elements, so each
// is written in the order of its complex type in ECMA-376 Part 1 (the tables
// below). What the reader kept without understanding it (`extras`) goes back:
//   - a child the model holds in part (`partial`) is written from the kept
//     tree, with the fields of the model on top of it. As long as the fields
//     still say what the tree says, the tree goes back as it is;
//   - a kept child whose name the schema knows takes its place in the order;
//   - any other kept child (an extension element) goes where it was among its
//     siblings (`index`);
//   - the element itself, kept for its attributes or comments (index -1),
//     goes back whole while the model still says what it says, and gives its
//     attributes otherwise.

import type {
  AttrList,
  Border,
  CellMargins,
  CellProps,
  DocColor,
  ExtraXml,
  FrameProps,
  HeaderFooterKind,
  Indent,
  NoteProps,
  ParaProps,
  ParaSpacing,
  RevisionInfo,
  RowProps,
  RunFonts,
  RunProps,
  SectionProps,
  Shading,
  TableBorders,
  TableLook,
  TablePosition,
  TableProps,
  TableWidth,
} from '../../model/doc/types';
import { TABLE_LOOK_BITS } from '../../model/doc/types';
import type { XmlElement, XmlNode } from '../../xml';
import { kindOf, serializeElement, type PartContext } from './common';
import { readCellProps, readParaProps, readRowProps, readRunProps, readSectionProps, readTableProps } from './props';

/** What the property writers need to know about the part they write into. */
export interface PropsContext {
  /** Prefix (with the colon) of WordprocessingML elements in the part; empty when it is the default namespace. */
  w: string;
  /** Prefix of WordprocessingML attributes. */
  a: string;
  /** Prefix of relationship attributes. */
  r: string;
  /** The part uses the strict (ISO/IEC 29500) namespaces. */
  strict: boolean;
  /** A reader for the part: kept trees are read again to see whether the model still says what they say. */
  reader: PartContext;
  /** Header/footer reference id of the model -> the relationship id to write; undefined leaves the reference out. */
  headerRef?(list: 'headers' | 'footers', id: string): string | undefined;
  /** Rewrites a colour before it is written (a package without a theme states plain values). */
  color?(color: DocColor): DocColor;
  /** Rewrites fonts before they are written. */
  fonts?(fonts: RunFonts): RunFonts;
}

type AttrValue = string | number | undefined;

function ranks(names: string[]): Map<string, number> {
  const map = new Map<string, number>();
  names.forEach((name, i) => map.set(name, i));
  return map;
}

/** CT_RPr / CT_ParaRPr: tracked marks of a paragraph mark first, the change record last. */
export const RPR_ORDER = ranks([
  'ins',
  'del',
  'moveFrom',
  'moveTo',
  'rStyle',
  'rFonts',
  'b',
  'bCs',
  'i',
  'iCs',
  'caps',
  'smallCaps',
  'strike',
  'dstrike',
  'outline',
  'shadow',
  'emboss',
  'imprint',
  'noProof',
  'snapToGrid',
  'vanish',
  'webHidden',
  'color',
  'spacing',
  'w',
  'kern',
  'position',
  'sz',
  'szCs',
  'highlight',
  'u',
  'effect',
  'bdr',
  'shd',
  'fitText',
  'vertAlign',
  'rtl',
  'cs',
  'em',
  'lang',
  'eastAsianLayout',
  'specVanish',
  'oMath',
  'rPrChange',
]);

/** CT_PPr. */
export const PPR_ORDER = ranks([
  'pStyle',
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'framePr',
  'widowControl',
  'numPr',
  'suppressLineNumbers',
  'pBdr',
  'shd',
  'tabs',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'spacing',
  'ind',
  'contextualSpacing',
  'mirrorIndents',
  'suppressOverlap',
  'jc',
  'textDirection',
  'textAlignment',
  'textboxTightWrap',
  'outlineLvl',
  'divId',
  'cnfStyle',
  'rPr',
  'sectPr',
  'pPrChange',
]);

/** CT_TblPr and CT_TblPrEx (whose children are a subset, in the same order). */
export const TBLPR_ORDER = ranks([
  'tblStyle',
  'tblpPr',
  'tblOverlap',
  'bidiVisual',
  'tblStyleRowBandSize',
  'tblStyleColBandSize',
  'tblW',
  'jc',
  'tblCellSpacing',
  'tblInd',
  'tblBorders',
  'shd',
  'tblLayout',
  'tblCellMar',
  'tblLook',
  'tblCaption',
  'tblDescription',
  'tblPrChange',
  'tblPrExChange',
]);

/** CT_TrPr. */
export const TRPR_ORDER = ranks([
  'cnfStyle',
  'divId',
  'gridBefore',
  'gridAfter',
  'wBefore',
  'wAfter',
  'cantSplit',
  'trHeight',
  'tblHeader',
  'tblCellSpacing',
  'jc',
  'hidden',
  'ins',
  'del',
  'trPrChange',
]);

/** CT_TcPr. */
export const TCPR_ORDER = ranks([
  'cnfStyle',
  'tcW',
  'gridSpan',
  'hMerge',
  'vMerge',
  'tcBorders',
  'shd',
  'noWrap',
  'tcMar',
  'textDirection',
  'tcFitText',
  'vAlign',
  'hideMark',
  'headers',
  'cellIns',
  'cellDel',
  'cellMerge',
  'tcPrChange',
]);

/** CT_SectPr. */
export const SECTPR_ORDER = ranks([
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
]);
// Header and footer references may come in any order among themselves.
SECTPR_ORDER.set('footerReference', 0);

/** A child written from the fields of the model. */
interface Gen {
  /** Local name of the element. */
  kind: string;
  el: XmlElement;
}

/** A WordprocessingML element; attributes without a value are left out, names without a prefix get the one of the part. */
export function wEl(ctx: PropsContext, local: string, attrs?: [name: string, value: AttrValue][], children?: XmlNode[]): XmlElement {
  const map = new Map<string, string>();
  if (attrs) {
    for (const [name, value] of attrs) {
      if (value === undefined) continue;
      map.set(name.indexOf(':') < 0 ? ctx.a + name : name, typeof value === 'number' ? String(value) : value);
    }
  }
  return { type: 'element', name: ctx.w + local, attrs: map, children: children || [] };
}

const bool = (value: boolean | undefined): string | undefined => (value === undefined ? undefined : value ? '1' : '0');

/** An on/off element: true is the element alone. */
const onOffEl = (ctx: PropsContext, local: string, value: boolean): XmlElement => wEl(ctx, local, value ? undefined : [['val', '0']]);

const hexByte = (value: number | undefined): string | undefined =>
  value === undefined ? undefined : ((value & 0xff) < 16 ? '0' : '') + (value & 0xff).toString(16).toUpperCase();

function colorAttrs(
  ctx: PropsContext,
  out: [string, AttrValue][],
  color: DocColor | undefined,
  val: string,
  theme: string,
  tint: string,
  shade: string,
): void {
  if (!color) return;
  const c = ctx.color ? ctx.color(color) : color;
  out.push([val, c.rgb !== undefined ? c.rgb : c.auto ? 'auto' : undefined]);
  out.push([theme, c.theme], [tint, hexByte(c.tint)], [shade, hexByte(c.shade)]);
}

function borderEl(ctx: PropsContext, local: string, border: Border): XmlElement {
  const attrs: [string, AttrValue][] = [
    ['val', border.style],
    ['sz', border.size],
    ['space', border.space],
  ];
  colorAttrs(ctx, attrs, border.color, 'color', 'themeColor', 'themeTint', 'themeShade');
  attrs.push(['shadow', bool(border.shadow)], ['frame', bool(border.frame)]);
  return wEl(ctx, local, attrs);
}

function shadingEl(ctx: PropsContext, shading: Shading): XmlElement {
  const attrs: [string, AttrValue][] = [['val', shading.pattern]];
  colorAttrs(ctx, attrs, shading.color, 'color', 'themeColor', 'themeTint', 'themeShade');
  colorAttrs(ctx, attrs, shading.fill, 'fill', 'themeFill', 'themeFillTint', 'themeFillShade');
  return wEl(ctx, 'shd', attrs);
}

/** The edges of a border container in the order of the schema; left and right by reading direction when the model says so. */
function edgeEls(ctx: PropsContext, borders: Record<string, Border | undefined>, edges: readonly string[]): XmlElement[] {
  const out: XmlElement[] = [];
  for (const edge of edges) {
    const border = borders[edge];
    if (!border) continue;
    const name = border.logical ? (edge === 'left' ? 'start' : edge === 'right' ? 'end' : edge) : edge;
    out.push(borderEl(ctx, name, border));
  }
  return out;
}

const PARA_EDGES = ['top', 'left', 'bottom', 'right', 'between', 'bar'];
const TABLE_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
const CELL_EDGES = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl'];
const PAGE_EDGES = ['top', 'left', 'bottom', 'right'];

function widthEl(ctx: PropsContext, local: string, width: TableWidth): XmlElement {
  return wEl(ctx, local, [
    ['w', width.value],
    ['type', width.type],
  ]);
}

function marginsEl(ctx: PropsContext, local: string, margins: CellMargins): XmlElement {
  const children: XmlElement[] = [];
  for (const edge of ['top', 'left', 'bottom', 'right'] as const) {
    const width = margins[edge];
    if (!width) continue;
    const name = width.logical ? (edge === 'left' ? 'start' : edge === 'right' ? 'end' : edge) : edge;
    children.push(widthEl(ctx, name, width));
  }
  return wEl(ctx, local, undefined, children);
}

function revisionEl(ctx: PropsContext, local: string, info: RevisionInfo): XmlElement {
  return wEl(ctx, local, [
    ['id', info.id],
    ['author', info.author],
    ['date', info.date],
  ]);
}

/** <w:cnfStyle>: the bit string in transitional documents, an attribute per set flag in strict ones (which have no w:val). */
function cnfStyleEl(ctx: PropsContext, bits: string): XmlElement {
  if (!ctx.strict) return wEl(ctx, 'cnfStyle', [['val', bits]]);
  const attrs: [string, AttrValue][] = [];
  CNF_ATTRS.forEach((name, i) => {
    if (bits.charAt(i) === '1') attrs.push([name, '1']);
  });
  return wEl(ctx, 'cnfStyle', attrs);
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

// ----------------------------------------------------------------------- run

const RUN_FLAGS: [local: string, field: keyof RunProps][] = [
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
  ['noProof', 'noProof'],
  ['snapToGrid', 'snapToGrid'],
  ['vanish', 'vanish'],
  ['webHidden', 'webHidden'],
  ['rtl', 'rtl'],
  ['cs', 'complexScript'],
  ['specVanish', 'specVanish'],
];

const FONT_ATTRS: [attribute: string, field: keyof RunFonts][] = [
  ['hint', 'hint'],
  ['ascii', 'ascii'],
  ['hAnsi', 'hAnsi'],
  ['eastAsia', 'eastAsia'],
  ['cs', 'cs'],
  ['asciiTheme', 'asciiTheme'],
  ['hAnsiTheme', 'hAnsiTheme'],
  ['eastAsiaTheme', 'eastAsiaTheme'],
  ['cstheme', 'csTheme'],
];

function runGens(ctx: PropsContext, p: RunProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  if (p.inserted) add('ins', revisionEl(ctx, 'ins', p.inserted));
  if (p.deleted) add('del', revisionEl(ctx, 'del', p.deleted));
  if (p.moveFrom) add('moveFrom', revisionEl(ctx, 'moveFrom', p.moveFrom));
  if (p.moveTo) add('moveTo', revisionEl(ctx, 'moveTo', p.moveTo));
  if (p.style !== undefined) add('rStyle', wEl(ctx, 'rStyle', [['val', p.style]]));
  if (p.fonts) {
    const fonts = ctx.fonts ? ctx.fonts(p.fonts) : p.fonts;
    add('rFonts', wEl(ctx, 'rFonts', FONT_ATTRS.map(([name, field]) => [name, fonts[field]] as [string, AttrValue])));
  }
  const flags = p as Record<string, unknown>;
  for (const [local, field] of RUN_FLAGS) {
    const value = flags[field];
    if (typeof value === 'boolean') add(local, onOffEl(ctx, local, value));
  }
  if (p.color) {
    const attrs: [string, AttrValue][] = [];
    colorAttrs(ctx, attrs, p.color, 'val', 'themeColor', 'themeTint', 'themeShade');
    add('color', wEl(ctx, 'color', attrs));
  }
  if (p.spacing !== undefined) add('spacing', wEl(ctx, 'spacing', [['val', p.spacing]]));
  if (p.scale !== undefined) add('w', wEl(ctx, 'w', [['val', p.scale]]));
  if (p.kerning !== undefined) add('kern', wEl(ctx, 'kern', [['val', p.kerning]]));
  if (p.position !== undefined) add('position', wEl(ctx, 'position', [['val', p.position]]));
  if (p.size !== undefined) add('sz', wEl(ctx, 'sz', [['val', p.size]]));
  if (p.sizeCs !== undefined) add('szCs', wEl(ctx, 'szCs', [['val', p.sizeCs]]));
  if (p.highlight !== undefined) add('highlight', wEl(ctx, 'highlight', [['val', p.highlight]]));
  if (p.underline) {
    const attrs: [string, AttrValue][] = [['val', p.underline.style]];
    colorAttrs(ctx, attrs, p.underline.color, 'color', 'themeColor', 'themeTint', 'themeShade');
    add('u', wEl(ctx, 'u', attrs));
  }
  if (p.effect !== undefined) add('effect', wEl(ctx, 'effect', [['val', p.effect]]));
  if (p.border) add('bdr', borderEl(ctx, 'bdr', p.border));
  if (p.shading) add('shd', shadingEl(ctx, p.shading));
  if (p.vertAlign !== undefined) add('vertAlign', wEl(ctx, 'vertAlign', [['val', p.vertAlign]]));
  if (p.emphasis !== undefined) add('em', wEl(ctx, 'em', [['val', p.emphasis]]));
  if (p.lang) {
    add(
      'lang',
      wEl(ctx, 'lang', [
        ['val', p.lang.val],
        ['eastAsia', p.lang.eastAsia],
        ['bidi', p.lang.bidi],
      ]),
    );
  }
  if (p.change) add(localOf(p.change.name), restatedRecord(ctx, p.change, RUN_KIND));
  return out;
}

// ----------------------------------------------------------------- paragraph

const PARA_FLAGS = [
  'keepNext',
  'keepLines',
  'pageBreakBefore',
  'widowControl',
  'suppressLineNumbers',
  'suppressAutoHyphens',
  'kinsoku',
  'wordWrap',
  'overflowPunct',
  'topLinePunct',
  'autoSpaceDE',
  'autoSpaceDN',
  'bidi',
  'adjustRightInd',
  'snapToGrid',
  'contextualSpacing',
  'mirrorIndents',
] as const;

function indentEl(ctx: PropsContext, indent: Indent): XmlElement {
  const logical = indent.logical === true;
  return wEl(ctx, 'ind', [
    [logical ? 'start' : 'left', indent.left],
    [logical ? 'startChars' : 'leftChars', indent.leftChars],
    [logical ? 'end' : 'right', indent.right],
    [logical ? 'endChars' : 'rightChars', indent.rightChars],
    ['hanging', indent.hanging],
    ['hangingChars', indent.hangingChars],
    ['firstLine', indent.firstLine],
    ['firstLineChars', indent.firstLineChars],
  ]);
}

function spacingEl(ctx: PropsContext, spacing: ParaSpacing): XmlElement {
  return wEl(ctx, 'spacing', [
    ['before', spacing.before],
    ['beforeLines', spacing.beforeLines],
    ['beforeAutospacing', bool(spacing.beforeAuto)],
    ['after', spacing.after],
    ['afterLines', spacing.afterLines],
    ['afterAutospacing', bool(spacing.afterAuto)],
    ['line', spacing.line],
    ['lineRule', spacing.lineRule],
  ]);
}

function frameEl(ctx: PropsContext, frame: FrameProps): XmlElement {
  return wEl(ctx, 'framePr', [
    ['dropCap', frame.dropCap],
    ['lines', frame.lines],
    ['w', frame.width],
    ['h', frame.height],
    ['vSpace', frame.vSpace],
    ['hSpace', frame.hSpace],
    ['wrap', frame.wrap],
    ['hAnchor', frame.hAnchor],
    ['vAnchor', frame.vAnchor],
    ['x', frame.x],
    ['xAlign', frame.xAlign],
    ['y', frame.y],
    ['yAlign', frame.yAlign],
    ['hRule', frame.heightRule],
    ['anchorLock', bool(frame.anchorLock)],
  ]);
}

function paraGens(ctx: PropsContext, p: ParaProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  if (p.style !== undefined) add('pStyle', wEl(ctx, 'pStyle', [['val', p.style]]));
  for (const name of PARA_FLAGS) {
    const value = p[name];
    if (value !== undefined) add(name, onOffEl(ctx, name, value));
  }
  if (p.frame) add('framePr', frameEl(ctx, p.frame));
  if (p.numbering) {
    const children: XmlElement[] = [];
    if (p.numbering.level !== undefined) children.push(wEl(ctx, 'ilvl', [['val', p.numbering.level]]));
    if (p.numbering.numId !== undefined) children.push(wEl(ctx, 'numId', [['val', p.numbering.numId]]));
    add('numPr', wEl(ctx, 'numPr', undefined, children));
  }
  if (p.borders) add('pBdr', wEl(ctx, 'pBdr', undefined, edgeEls(ctx, p.borders as Record<string, Border | undefined>, PARA_EDGES)));
  if (p.shading) add('shd', shadingEl(ctx, p.shading));
  if (p.tabs) {
    add(
      'tabs',
      wEl(
        ctx,
        'tabs',
        undefined,
        p.tabs.map((tab) =>
          wEl(ctx, 'tab', [
            ['val', tab.align],
            ['leader', tab.leader],
            ['pos', tab.pos],
          ]),
        ),
      ),
    );
  }
  if (p.spacing) add('spacing', spacingEl(ctx, p.spacing));
  if (p.indent) add('ind', indentEl(ctx, p.indent));
  if (p.align !== undefined) add('jc', wEl(ctx, 'jc', [['val', p.align]]));
  if (p.textDirection !== undefined) add('textDirection', wEl(ctx, 'textDirection', [['val', p.textDirection]]));
  if (p.textAlignment !== undefined) add('textAlignment', wEl(ctx, 'textAlignment', [['val', p.textAlignment]]));
  if (p.outlineLevel !== undefined) add('outlineLvl', wEl(ctx, 'outlineLvl', [['val', p.outlineLevel]]));
  if (p.cnfStyle !== undefined) add('cnfStyle', cnfStyleEl(ctx, p.cnfStyle));
  if (p.markProps) add('rPr', runPropsElement(ctx, p.markProps) || wEl(ctx, 'rPr'));
  if (p.section) add('sectPr', sectionPropsElement(ctx, p.section));
  if (p.change) add(localOf(p.change.name), restatedRecord(ctx, p.change, PARA_KIND));
  return out;
}

// --------------------------------------------------------------------- table

function lookEl(ctx: PropsContext, look: TableLook): XmlElement {
  const attrs: [string, AttrValue][] = [];
  const names = ['firstRow', 'lastRow', 'firstColumn', 'lastColumn', 'noHBand', 'noVBand'] as const;
  if (look.val === undefined && look.stated === undefined) {
    // A model that does not come from a docx file: both spellings, as Word writes them.
    let bits = 0;
    for (const name of names) if (look[name]) bits |= TABLE_LOOK_BITS[name];
    if (!ctx.strict) attrs.push(['val', ('000' + bits.toString(16).toUpperCase()).slice(-4)]);
    for (const name of names) attrs.push([name, bool(look[name])]);
    return wEl(ctx, 'tblLook', attrs);
  }
  let val = look.val;
  const stated = look.stated || 0;
  if (val !== undefined) {
    // A flag that is not stated on its own is read from its bit, so the bit follows the model.
    for (const name of names) {
      if (stated & TABLE_LOOK_BITS[name]) continue;
      val = look[name] ? val | TABLE_LOOK_BITS[name] : val & ~TABLE_LOOK_BITS[name];
    }
    attrs.push(['val', ('000' + (val & 0xffff).toString(16).toUpperCase()).slice(-4)]);
  }
  for (const name of names) {
    // Without a bit field every flag that is set has to be stated.
    if (stated & TABLE_LOOK_BITS[name] || (look.val === undefined && look[name])) attrs.push([name, bool(look[name])]);
  }
  return wEl(ctx, 'tblLook', attrs);
}

function positionEl(ctx: PropsContext, position: TablePosition): XmlElement {
  return wEl(ctx, 'tblpPr', [
    ['leftFromText', position.leftFromText],
    ['rightFromText', position.rightFromText],
    ['topFromText', position.topFromText],
    ['bottomFromText', position.bottomFromText],
    ['vertAnchor', position.vertAnchor],
    ['horzAnchor', position.horzAnchor],
    ['tblpXSpec', position.xAlign],
    ['tblpX', position.x],
    ['tblpYSpec', position.yAlign],
    ['tblpY', position.y],
  ]);
}

function tableGens(ctx: PropsContext, p: TableProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  if (p.style !== undefined) add('tblStyle', wEl(ctx, 'tblStyle', [['val', p.style]]));
  if (p.floating) add('tblpPr', positionEl(ctx, p.floating));
  if (p.overlap !== undefined) add('tblOverlap', wEl(ctx, 'tblOverlap', [['val', p.overlap]]));
  if (p.bidiVisual !== undefined) add('bidiVisual', onOffEl(ctx, 'bidiVisual', p.bidiVisual));
  if (p.rowBandSize !== undefined) add('tblStyleRowBandSize', wEl(ctx, 'tblStyleRowBandSize', [['val', p.rowBandSize]]));
  if (p.colBandSize !== undefined) add('tblStyleColBandSize', wEl(ctx, 'tblStyleColBandSize', [['val', p.colBandSize]]));
  if (p.width) add('tblW', widthEl(ctx, 'tblW', p.width));
  if (p.align !== undefined) add('jc', wEl(ctx, 'jc', [['val', p.align]]));
  if (p.cellSpacing) add('tblCellSpacing', widthEl(ctx, 'tblCellSpacing', p.cellSpacing));
  if (p.indent) add('tblInd', widthEl(ctx, 'tblInd', p.indent));
  if (p.borders) add('tblBorders', wEl(ctx, 'tblBorders', undefined, edgeEls(ctx, p.borders as Record<string, Border | undefined>, TABLE_EDGES)));
  if (p.shading) add('shd', shadingEl(ctx, p.shading));
  if (p.layout !== undefined) add('tblLayout', wEl(ctx, 'tblLayout', [['type', p.layout]]));
  if (p.cellMargins) add('tblCellMar', marginsEl(ctx, 'tblCellMar', p.cellMargins));
  if (p.look) add('tblLook', lookEl(ctx, p.look));
  if (p.caption !== undefined) add('tblCaption', wEl(ctx, 'tblCaption', [['val', p.caption]]));
  if (p.description !== undefined) add('tblDescription', wEl(ctx, 'tblDescription', [['val', p.description]]));
  if (p.change) add(localOf(p.change.name), p.change);
  return out;
}

function rowGens(ctx: PropsContext, p: RowProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  if (p.cnfStyle !== undefined) add('cnfStyle', cnfStyleEl(ctx, p.cnfStyle));
  if (p.gridBefore !== undefined) add('gridBefore', wEl(ctx, 'gridBefore', [['val', p.gridBefore]]));
  if (p.gridAfter !== undefined) add('gridAfter', wEl(ctx, 'gridAfter', [['val', p.gridAfter]]));
  if (p.widthBefore) add('wBefore', widthEl(ctx, 'wBefore', p.widthBefore));
  if (p.widthAfter) add('wAfter', widthEl(ctx, 'wAfter', p.widthAfter));
  if (p.cantSplit !== undefined) add('cantSplit', onOffEl(ctx, 'cantSplit', p.cantSplit));
  if (p.height !== undefined || p.heightRule !== undefined) {
    add(
      'trHeight',
      wEl(ctx, 'trHeight', [
        ['val', p.height],
        ['hRule', p.heightRule],
      ]),
    );
  }
  if (p.header !== undefined) add('tblHeader', onOffEl(ctx, 'tblHeader', p.header));
  if (p.cellSpacing) add('tblCellSpacing', widthEl(ctx, 'tblCellSpacing', p.cellSpacing));
  if (p.align !== undefined) add('jc', wEl(ctx, 'jc', [['val', p.align]]));
  if (p.hidden !== undefined) add('hidden', onOffEl(ctx, 'hidden', p.hidden));
  if (p.inserted) add('ins', revisionEl(ctx, 'ins', p.inserted));
  if (p.deleted) add('del', revisionEl(ctx, 'del', p.deleted));
  if (p.change) add(localOf(p.change.name), p.change);
  return out;
}

function cellGens(ctx: PropsContext, p: CellProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  if (p.cnfStyle !== undefined) add('cnfStyle', cnfStyleEl(ctx, p.cnfStyle));
  if (p.width) add('tcW', widthEl(ctx, 'tcW', p.width));
  if (p.gridSpan !== undefined) add('gridSpan', wEl(ctx, 'gridSpan', [['val', p.gridSpan]]));
  // "continue" is what the element says without a value.
  if (p.hMerge !== undefined) add('hMerge', wEl(ctx, 'hMerge', [['val', p.hMerge === 'restart' ? 'restart' : undefined]]));
  if (p.vMerge !== undefined) add('vMerge', wEl(ctx, 'vMerge', [['val', p.vMerge === 'restart' ? 'restart' : undefined]]));
  if (p.borders) add('tcBorders', wEl(ctx, 'tcBorders', undefined, edgeEls(ctx, p.borders as Record<string, Border | undefined>, CELL_EDGES)));
  if (p.shading) add('shd', shadingEl(ctx, p.shading));
  if (p.noWrap !== undefined) add('noWrap', onOffEl(ctx, 'noWrap', p.noWrap));
  if (p.margins) add('tcMar', marginsEl(ctx, 'tcMar', p.margins));
  if (p.textDirection !== undefined) add('textDirection', wEl(ctx, 'textDirection', [['val', p.textDirection]]));
  if (p.fitText !== undefined) add('tcFitText', onOffEl(ctx, 'tcFitText', p.fitText));
  if (p.vAlign !== undefined) add('vAlign', wEl(ctx, 'vAlign', [['val', p.vAlign]]));
  if (p.hideMark !== undefined) add('hideMark', onOffEl(ctx, 'hideMark', p.hideMark));
  if (p.change) add(localOf(p.change.name), p.change);
  return out;
}

// ------------------------------------------------------------------- section

function notePropsEl(ctx: PropsContext, local: string, note: NoteProps): XmlElement {
  const children: XmlElement[] = [];
  if (note.position !== undefined) children.push(wEl(ctx, 'pos', [['val', note.position]]));
  if (note.format !== undefined) children.push(wEl(ctx, 'numFmt', [['val', note.format]]));
  if (note.start !== undefined) children.push(wEl(ctx, 'numStart', [['val', note.start]]));
  if (note.restart !== undefined) children.push(wEl(ctx, 'numRestart', [['val', note.restart]]));
  return wEl(ctx, local, undefined, children);
}

const HEADER_KINDS: HeaderFooterKind[] = ['even', 'default', 'first'];

function sectionGens(ctx: PropsContext, p: SectionProps): Gen[] {
  const out: Gen[] = [];
  const add = (kind: string, el: XmlElement) => out.push({ kind, el });
  for (const list of ['headers', 'footers'] as const) {
    const refs = p[list];
    if (!refs) continue;
    const local = list === 'headers' ? 'headerReference' : 'footerReference';
    // In the order the model lists them, which is the order of the file.
    const kinds = (Object.keys(refs) as HeaderFooterKind[]).filter((kind) => HEADER_KINDS.indexOf(kind) >= 0);
    for (const kind of kinds) {
      const id = refs[kind];
      if (id === undefined) continue;
      const rId = ctx.headerRef ? ctx.headerRef(list, id) : id;
      if (rId === undefined) continue;
      add(
        local,
        wEl(ctx, local, [
          ['type', kind],
          [ctx.r + 'id', rId],
        ]),
      );
    }
  }
  if (p.footnotes) add('footnotePr', notePropsEl(ctx, 'footnotePr', p.footnotes));
  if (p.endnotes) add('endnotePr', notePropsEl(ctx, 'endnotePr', p.endnotes));
  if (p.type !== undefined) add('type', wEl(ctx, 'type', [['val', p.type]]));
  if (p.page) {
    add(
      'pgSz',
      wEl(ctx, 'pgSz', [
        ['w', p.page.width],
        ['h', p.page.height],
        ['orient', p.page.orientation],
        ['code', p.page.code],
      ]),
    );
  }
  if (p.margins) {
    const m = p.margins;
    add(
      'pgMar',
      wEl(ctx, 'pgMar', [
        ['top', m.top],
        ['right', m.right],
        ['bottom', m.bottom],
        ['left', m.left],
        ['header', m.header],
        ['footer', m.footer],
        ['gutter', m.gutter],
      ]),
    );
  }
  if (p.pageBorders) {
    const b = p.pageBorders;
    const el = wEl(
      ctx,
      'pgBorders',
      [
        ['zOrder', b.zOrder],
        ['display', b.display],
        ['offsetFrom', b.offsetFrom],
      ],
      edgeEls(ctx, b as unknown as Record<string, Border | undefined>, PAGE_EDGES),
    );
    add('pgBorders', el);
  }
  if (p.lineNumbering) {
    const n = p.lineNumbering;
    add(
      'lnNumType',
      wEl(ctx, 'lnNumType', [
        ['countBy', n.countBy],
        ['start', n.start],
        ['distance', n.distance],
        ['restart', n.restart],
      ]),
    );
  }
  if (p.pageNumbering) {
    const n = p.pageNumbering;
    add(
      'pgNumType',
      wEl(ctx, 'pgNumType', [
        ['fmt', n.format],
        ['start', n.start],
        ['chapStyle', n.chapStyle],
        ['chapSep', n.chapSep],
      ]),
    );
  }
  if (p.columns) {
    const c = p.columns;
    const cols = (c.columns || []).map((column) =>
      wEl(ctx, 'col', [
        ['w', column.width],
        ['space', column.space],
      ]),
    );
    add(
      'cols',
      wEl(
        ctx,
        'cols',
        [
          ['num', c.count],
          ['sep', bool(c.separator)],
          ['space', c.space],
          ['equalWidth', bool(c.equalWidth)],
        ],
        cols,
      ),
    );
  }
  if (p.vAlign !== undefined) add('vAlign', wEl(ctx, 'vAlign', [['val', p.vAlign]]));
  if (p.titlePage !== undefined) add('titlePg', onOffEl(ctx, 'titlePg', p.titlePage));
  if (p.textDirection !== undefined) add('textDirection', wEl(ctx, 'textDirection', [['val', p.textDirection]]));
  if (p.bidi !== undefined) add('bidi', onOffEl(ctx, 'bidi', p.bidi));
  if (p.rtlGutter !== undefined) add('rtlGutter', onOffEl(ctx, 'rtlGutter', p.rtlGutter));
  if (p.docGrid) {
    add(
      'docGrid',
      wEl(ctx, 'docGrid', [
        ['type', p.docGrid.type],
        ['linePitch', p.docGrid.linePitch],
        ['charSpace', p.docGrid.charSpace],
      ]),
    );
  }
  if (p.change) add(localOf(p.change.name), p.change);
  return out;
}

// ------------------------------------------------------------------ assembly

const localOf = (name: string): string => name.slice(name.indexOf(':') + 1);

const elementsOf = (el: XmlElement): XmlElement[] => el.children.filter((node): node is XmlElement => node.type === 'element');

/**
 * A kept element with the fields of the model on top of it. `understood` is
 * what the writer makes of the kept element alone (undefined when the reader
 * understood nothing of it), `wanted` what it makes of the model. What the
 * kept element says beyond `understood` stays; what `understood` says and
 * `wanted` does not is gone; a value that did not change keeps its spelling.
 */
export function mergeKept(kept: XmlElement, understood: XmlElement | undefined, wanted: XmlElement, depth = 0): XmlElement {
  if (depth > 64) return wanted;
  const attrs = new Map<string, string>();
  for (const [name, value] of kept.attrs) {
    const next = wanted.attrs.get(name);
    if (understood && understood.attrs.has(name)) {
      if (next !== undefined) attrs.set(name, next === understood.attrs.get(name) ? value : next);
    } else attrs.set(name, next !== undefined ? next : value);
  }
  for (const [name, value] of wanted.attrs) if (!kept.attrs.has(name)) attrs.set(name, value);

  // Children pair up by name and order: the n-th <x> of each.
  const nth = (list: XmlElement[]) => {
    const seen = new Map<string, number>();
    const map = new Map<string, XmlElement>();
    for (const el of list) {
      const n = seen.get(el.name) || 0;
      seen.set(el.name, n + 1);
      map.set(el.name + '\n' + n, el);
    }
    return map;
  };
  const understoodChildren = nth(understood ? elementsOf(understood) : []);
  const wantedChildren = nth(elementsOf(wanted));
  // What became of each child of `wanted` that has a counterpart in the kept element.
  const placed = new Map<XmlElement, XmlNode>();
  const children: XmlNode[] = [];
  const seen = new Map<string, number>();
  for (const node of kept.children) {
    if (node.type !== 'element') {
      children.push(node);
      continue;
    }
    const n = seen.get(node.name) || 0;
    seen.set(node.name, n + 1);
    const key = node.name + '\n' + n;
    const was = understoodChildren.get(key);
    const now = wantedChildren.get(key);
    if (now) {
      const merged = mergeKept(node, was, now, depth + 1);
      placed.set(now, merged);
      children.push(merged);
    } else if (!was) children.push(node);
  }
  // The children the kept element lacks go after the one that precedes them in `wanted`, which is in the order of the schema.
  let after = -1;
  for (const el of elementsOf(wanted)) {
    const merged = placed.get(el);
    if (merged) after = children.indexOf(merged);
    else children.splice(++after, 0, el);
  }
  return { type: 'element', name: kept.name, attrs, children };
}

interface Bag {
  extras?: ExtraXml[];
}

interface Kind<P extends Bag> {
  local: string;
  order: Map<string, number>;
  gens(ctx: PropsContext, props: P): Gen[];
  read(ctx: PartContext, el: XmlElement): P;
}

/** The rank of a kept element in the order of its container; undefined for an element of another vocabulary. */
function rankOf(ctx: PropsContext, order: Map<string, number>, el: XmlElement): number | undefined {
  const kind = kindOf(ctx.reader, el.name);
  // Names of other namespaces come back with their prefix and so are not in the table.
  return order.get(kind);
}

/**
 * The children of a container in their order: those the schema knows by its
 * sequence (kept ones after written ones of the same name, in the order they
 * had), the others where they were among their siblings.
 */
function place(ctx: PropsContext, order: Map<string, number>, gens: Gen[], loose: ExtraXml[]): XmlNode[] {
  const ranked: { rank: number; at: number; el: XmlElement }[] = [];
  const foreign: ExtraXml[] = [];
  gens.forEach((gen, i) => {
    const rank = order.get(gen.kind);
    ranked.push({ rank: rank === undefined ? order.size : rank, at: i - gens.length, el: gen.el });
  });
  for (const extra of loose) {
    const rank = rankOf(ctx, order, extra.el);
    if (rank === undefined) foreign.push(extra);
    else ranked.push({ rank, at: extra.index, el: extra.el });
  }
  ranked.sort((x, y) => x.rank - y.rank || x.at - y.at);

  const total = ranked.length + foreign.length;
  const slots: (XmlElement | undefined)[] = new Array(total);
  const late: XmlElement[] = [];
  for (const extra of foreign) {
    if (extra.index < total && slots[extra.index] === undefined) slots[extra.index] = extra.el;
    else late.push(extra.el);
  }
  const children: XmlNode[] = [];
  let next = 0;
  for (let i = 0; i < total; i++) {
    const kept = slots[i];
    if (kept) children.push(kept);
    else if (next < ranked.length) children.push(ranked[next++].el);
  }
  while (next < ranked.length) children.push(ranked[next++].el);
  for (const el of late) children.push(el);
  return children;
}

/**
 * An element of a part the model holds in fields (a style, a numbering
 * level): the written children with what the reader kept of the element. A
 * kept child the model holds in part is written from the kept tree with the
 * fields on top; the element itself, kept for its attributes, gives those the
 * model has no field for.
 */
export function keptElement(
  ctx: PropsContext,
  local: string,
  order: Map<string, number>,
  gens: { kind: string; el: XmlElement }[],
  extras: ExtraXml[] | undefined,
  attrs: [name: string, value: string | number | undefined][],
): XmlElement {
  const el = wEl(ctx, local, attrs);
  const loose: ExtraXml[] = [];
  const taken = new Set<Gen>();
  for (const extra of extras || []) {
    if (extra.index < 0) {
      for (const [name, value] of extra.el.attrs) if (!el.attrs.has(name)) el.attrs.set(name, value);
      continue;
    }
    if (!extra.partial) {
      loose.push(extra);
      continue;
    }
    const name = kindOf(ctx.reader, extra.el.name);
    const free = gens.filter((candidate) => candidate.kind === name && !taken.has(candidate));
    // Of several written children of that name (the conditions of a table style), the one with the same attributes.
    const gen = free.find((candidate) => Array.from(candidate.el.attrs).every(([key, value]) => extra.el.attrs.get(key) === value)) || free[0];
    if (!gen) {
      loose.push(extra);
      continue;
    }
    taken.add(gen);
    gen.el = mergeKept(extra.el, undefined, gen.el);
  }
  el.children = place(ctx, order, gens, loose);
  return el;
}

function assemble<P extends Bag>(ctx: PropsContext, kind: Kind<P>, props: P, attrs: AttrList | undefined, local: string, top: boolean): XmlElement | undefined {
  const gens = kind.gens(ctx, props);
  let self: XmlElement | undefined;
  const loose: ExtraXml[] = [];
  if (props.extras) {
    const taken = new Set<Gen>();
    for (const extra of props.extras) {
      if (extra.index < 0) {
        if (!self) self = extra.el;
        continue;
      }
      if (!extra.partial) {
        loose.push(extra);
        continue;
      }
      // The child the model holds a part of: the first written child of that name not yet matched.
      const name = kindOf(ctx.reader, extra.el.name);
      // What the writer makes of the kept child alone.
      const alone: XmlElement = { type: 'element', name: ctx.w + kind.local, attrs: new Map(), children: [extra.el] };
      const understood = kind.gens(ctx, kind.read(ctx.reader, alone)).find((candidate) => candidate.kind === name);
      const same = understood ? serializeElement(understood.el) : undefined;
      const free = gens.filter((candidate) => candidate.kind === name && !taken.has(candidate));
      // Of several written children of that name (header references), the one that says the same.
      const unchanged = same === undefined ? undefined : free.find((candidate) => serializeElement(candidate.el) === same);
      const gen = unchanged || free[0];
      if (!gen) {
        loose.push(extra);
        continue;
      }
      taken.add(gen);
      if (unchanged) gen.el = extra.el;
      else gen.el = mergeKept(extra.el, understood ? understood.el : undefined, gen.el);
    }
  }

  const children = place(ctx, kind.order, gens, loose);

  const map = new Map<string, string>();
  if (self) for (const [name, value] of self.attrs) map.set(name, value);
  else if (attrs) for (const [name, value] of attrs) map.set(name, value);
  if (!children.length && !map.size && !self) return undefined;
  const el: XmlElement = { type: 'element', name: ctx.w + local, attrs: map, children };

  if (self && top) {
    // The element as it was read, comments and all, as long as the model still says what it says.
    const again = assemble(ctx, kind, kind.read(ctx.reader, self), undefined, local, false);
    if (again && serializeElement(again) === serializeElement(el)) return self;
  }
  return el;
}

/**
 * A record of changed formatting (<w:rPrChange>, <w:pPrChange>) as it is written. Where colours or fonts are
 * rewritten for the package (one made without a theme), the properties the record holds are rewritten like the
 * present ones: rejected, the change gives the colour and the font the text had. Everywhere else the record is
 * written as it is kept, byte for byte.
 */
function restatedRecord<P extends Bag>(ctx: PropsContext, change: XmlElement, kind: Kind<P>): XmlElement {
  if (!ctx.color && !ctx.fonts) return change;
  let children: XmlNode[] | undefined;
  for (let i = 0; i < change.children.length; i++) {
    const node = change.children[i];
    if (node.type !== 'element' || kindOf(ctx.reader, node.name) !== kind.local) continue;
    const stated = assemble(ctx, kind, kind.read(ctx.reader, node), undefined, kind.local, true) || wEl(ctx, kind.local);
    if (serializeElement(stated) !== serializeElement(node)) {
      children = change.children.slice();
      children[i] = stated;
    }
    break;
  }
  return children ? { ...change, children } : change;
}

const RUN_KIND: Kind<RunProps> = { local: 'rPr', order: RPR_ORDER, gens: runGens, read: readRunProps };
const PARA_KIND: Kind<ParaProps> = { local: 'pPr', order: PPR_ORDER, gens: paraGens, read: readParaProps };
const TABLE_KIND: Kind<TableProps> = { local: 'tblPr', order: TBLPR_ORDER, gens: tableGens, read: readTableProps };
const ROW_KIND: Kind<RowProps> = { local: 'trPr', order: TRPR_ORDER, gens: rowGens, read: readRowProps };
const CELL_KIND: Kind<CellProps> = { local: 'tcPr', order: TCPR_ORDER, gens: cellGens, read: readCellProps };
const SECTION_KIND: Kind<SectionProps> = { local: 'sectPr', order: SECTPR_ORDER, gens: sectionGens, read: readSectionProps };

/** <w:rPr> for the properties; undefined when they say nothing. */
export function runPropsElement(ctx: PropsContext, props: RunProps): XmlElement | undefined {
  return assemble(ctx, RUN_KIND, props, undefined, 'rPr', true);
}

/** <w:pPr> for the properties; undefined when they say nothing. */
export function paraPropsElement(ctx: PropsContext, props: ParaProps): XmlElement | undefined {
  return assemble(ctx, PARA_KIND, props, undefined, 'pPr', true);
}

/** <w:tblPr> (or <w:tblPrEx> for the exceptions of a row); undefined when the properties say nothing. */
export function tablePropsElement(ctx: PropsContext, props: TableProps, local: 'tblPr' | 'tblPrEx' = 'tblPr'): XmlElement | undefined {
  return assemble(ctx, TABLE_KIND, props, undefined, local, true);
}

/** <w:trPr>; undefined when the properties say nothing. */
export function rowPropsElement(ctx: PropsContext, props: RowProps): XmlElement | undefined {
  return assemble(ctx, ROW_KIND, props, undefined, 'trPr', true);
}

/** <w:tcPr>; undefined when the properties say nothing. */
export function cellPropsElement(ctx: PropsContext, props: CellProps): XmlElement | undefined {
  return assemble(ctx, CELL_KIND, props, undefined, 'tcPr', true);
}

/** <w:sectPr>: always an element, as a section mark means something without any property. */
export function sectionPropsElement(ctx: PropsContext, props: SectionProps): XmlElement {
  return assemble(ctx, SECTION_KIND, props, props.attrs, 'sectPr', true) || wEl(ctx, 'sectPr');
}
