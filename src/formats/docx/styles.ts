// Reader for word/styles.xml.
//
// The part itself stays in the package and is written back from there; the
// model built here is what layout resolves formatting with. Styles are keyed
// by id in a Map, so an id such as "__proto__" is just another key.

import {
  emptyDocStyles,
  setPartNamespaces,
  type DocStyle,
  type DocStyles,
  type ExtraXml,
  type StyleType,
  type TableStyleCondition,
} from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { attr, boolAttr, childOfKind, intAttr, kindOf, onOff, pushExtra, type PartContext } from './common';
import { readCellProps, readParaProps, readRowProps, readRunProps, readTableProps } from './props';

/** More styles than this in one document is not a genuine file. */
export const MAX_STYLES = 20000;

const STYLE_TYPES: readonly string[] = ['paragraph', 'character', 'table', 'numbering'];

const STYLE_FLAGS = new Map<string, 'hidden' | 'semiHidden' | 'unhideWhenUsed' | 'quickFormat' | 'locked'>([
  ['hidden', 'hidden'],
  ['semiHidden', 'semiHidden'],
  ['unhideWhenUsed', 'unhideWhenUsed'],
  ['qFormat', 'quickFormat'],
  ['locked', 'locked'],
]);

/** The property children shared by <w:style> and <w:tblStylePr>; false when the child is something else. */
function readFormatting(ctx: PartContext, kind: string, el: XmlElement, target: TableStyleCondition): boolean {
  if (kind === 'pPr' && !target.para) target.para = readParaProps(ctx, el);
  else if (kind === 'rPr' && !target.run) target.run = readRunProps(ctx, el);
  else if (kind === 'tblPr' && !target.table) target.table = readTableProps(ctx, el);
  else if (kind === 'trPr' && !target.row) target.row = readRowProps(ctx, el);
  else if (kind === 'tcPr' && !target.cell) target.cell = readCellProps(ctx, el);
  else return false;
  return true;
}

function readStyle(ctx: PartContext, el: XmlElement): DocStyle | undefined {
  const id = attr(ctx, el, 'styleId');
  if (id === undefined) return undefined;
  const type = attr(ctx, el, 'type');
  const style: DocStyle = {
    id,
    // A style without a type is a paragraph style.
    type: type !== undefined && STYLE_TYPES.indexOf(type) >= 0 ? (type as StyleType) : 'paragraph',
    isDefault: boolAttr(ctx, el, 'default') === true,
  };
  if (boolAttr(ctx, el, 'customStyle')) style.custom = true;
  let extras: ExtraXml[] | undefined;
  let index = -1;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    index++;
    const kind = kindOf(ctx, node.name);
    const flag = STYLE_FLAGS.get(kind);
    ctx.hits = 0;
    if (flag !== undefined) {
      const value = onOff(ctx, node);
      if (value !== undefined) style[flag] = value;
      else extras = pushExtra(ctx, extras, index, node, 'style', false);
    } else if (kind === 'name' || kind === 'basedOn' || kind === 'next' || kind === 'link') {
      const value = attr(ctx, node, 'val');
      if (value !== undefined && style[kind] === undefined) style[kind] = value;
      else extras = pushExtra(ctx, extras, index, node, 'style', false);
    } else if (kind === 'uiPriority') {
      const value = intAttr(ctx, node, 'val', 0, 0x7fffffff);
      if (value !== undefined) style.uiPriority = value;
      else extras = pushExtra(ctx, extras, index, node, 'style', false);
    } else if (kind === 'tblStylePr') {
      const conditionType = attr(ctx, node, 'type');
      const conditions = style.conditions || (style.conditions = new Map());
      // Only the first condition of a type counts; one without a type, or a repeated one, is not in the model at all.
      if (conditionType === undefined || conditions.has(conditionType)) {
        extras = pushExtra(ctx, extras, index, node, 'style', false);
        continue;
      }
      const condition: TableStyleCondition = {};
      let whole = node.attrs.size === 1;
      for (const inner of node.children) {
        if (inner.type === 'element' && !readFormatting(ctx, kindOf(ctx, inner.name), inner, condition)) whole = false;
      }
      conditions.set(conditionType, condition);
      if (!whole) extras = pushExtra(ctx, extras, index, node, 'style', true);
    } else if (!readFormatting(ctx, kind, node, style)) extras = pushExtra(ctx, extras, index, node, 'style', false);
  }
  if (extras) style.extras = extras;
  return style;
}

/** Reads the root element of the styles part; undefined gives the empty style sheet. */
export function parseDocStyles(ctx: PartContext, root: XmlElement | undefined): DocStyles {
  const sheet = emptyDocStyles();
  if (!root) return sheet;
  setPartNamespaces(sheet, Array.from(root.attrs).filter(([name]) => name === 'xmlns' || name.indexOf('xmlns:') === 0));
  let count = 0;
  for (const node of root.children) {
    if (node.type !== 'element') continue;
    const kind = kindOf(ctx, node.name);
    if (kind === 'docDefaults') {
      const run = childOfKind(ctx, node, 'rPrDefault');
      const para = childOfKind(ctx, node, 'pPrDefault');
      sheet.defaults = {
        run: readRunProps(ctx, run && childOfKind(ctx, run, 'rPr')),
        para: readParaProps(ctx, para && childOfKind(ctx, para, 'pPr')),
      };
    } else if (kind === 'style') {
      if (++count > MAX_STYLES) throw new Error(`Invalid docx: more than ${MAX_STYLES} styles`);
      const style = readStyle(ctx, node);
      if (!style) continue;
      // The first definition of an id is the one that counts.
      if (sheet.styles.has(style.id)) {
        ctx.state.warnings.push(`${ctx.part}: a style id is defined more than once`);
        continue;
      }
      sheet.styles.set(style.id, style);
      if (style.isDefault && sheet.defaultIds[style.type] === undefined) sheet.defaultIds[style.type] = style.id;
    }
  }
  return sheet;
}
