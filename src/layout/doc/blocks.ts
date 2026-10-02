// Lays out a list of blocks (the content of a cell, a header, a note, a text
// box) as bands, and the pieces of that which the body, laid out block by
// block, uses as well.

import type { Block, ParaProps, Paragraph } from '../../model/doc/types';
import type { Band, CellWhere, FieldValues, LayoutContext, StoryRef } from './context';
import { paragraphBands, type ParaEnv } from './paragraph';
import { isAutoSpace, paragraphSpace, prepareParagraph, type ParaPrep } from './prep';
import { tableBands } from './table';

const PREPS = new WeakMap<LayoutContext, WeakMap<Paragraph, ParaPrep>>();

/**
 * The prepared form of a paragraph, made on first use. Numbered paragraphs
 * take their number then, so paragraphs have to be asked for in document
 * order the first time.
 */
export function getPrep(
  ctx: LayoutContext,
  para: Paragraph,
  story: StoryRef,
  path: number[],
  where: CellWhere | undefined,
  depth: number,
  fields: FieldValues | undefined,
): ParaPrep {
  let preps = PREPS.get(ctx);
  if (!preps) PREPS.set(ctx, (preps = new WeakMap()));
  let prep = preps.get(para);
  if (!prep) preps.set(para, (prep = prepareParagraph(ctx, para, story, path, where, depth, fields)));
  return prep;
}

const BORDER_KEYS = new WeakMap<ParaProps, string>();

/** Identifies the borders and indents of a paragraph: neighbours with the same key share one box. Empty without borders. */
function borderKey(props: ParaProps): string {
  if (!props.borders) return '';
  let key = BORDER_KEYS.get(props);
  if (key === undefined) {
    const indent = props.indent || {};
    key = JSON.stringify([props.borders.top, props.borders.left, props.borders.bottom, props.borders.right, props.borders.between, indent.left, indent.right]);
    BORDER_KEYS.set(props, key);
  }
  return key;
}

/** The numbering a paragraph belongs to, 0 when it is not numbered. */
const listOf = (props: ParaProps): number => (props.numbering && props.numbering.numId !== undefined ? props.numbering.numId : 0);

/** The nearest paragraph or table before (step -1) or after (step 1) a block; marks between blocks do not count. */
function neighbour(blocks: Block[], index: number, step: number): Block | undefined {
  // Marks come in small groups; a long run of them is not looked through.
  for (let i = index + step, tries = 0; i >= 0 && i < blocks.length && tries < 16; i += step, tries++) {
    const block = blocks[i];
    if (block.type === 'paragraph' || block.type === 'table') return block;
  }
  return undefined;
}

/** What the neighbours of a paragraph mean for its spacing and borders. */
export function paragraphEnv(
  ctx: LayoutContext,
  blocks: Block[],
  index: number,
  prep: ParaPrep,
  where: CellWhere | undefined,
  fields: FieldValues | undefined,
): ParaEnv {
  const env: ParaEnv = { joinPrev: false, joinNext: false, before: prep.before, after: prep.after };
  if (fields) env.fields = fields;
  if (ctx.grid !== undefined) env.grid = ctx.grid;
  const key = borderKey(prep.props);
  if (key === '' && !prep.contextual && !prep.beforeAuto && !prep.afterAuto && env.before === 0) return env;
  const before = neighbour(blocks, index, -1);
  const after = neighbour(blocks, index, 1);
  const list = listOf(prep.props);
  if (before && before.type === 'paragraph') {
    const props = ctx.paraProps(before.props, where);
    // A paragraph that ends a section closes its box.
    if (key !== '' && borderKey(props) === key && !before.props.section) env.joinPrev = true;
    if (prep.contextual && props.style === prep.props.style) env.before = 0;
    else if (prep.beforeAuto || isAutoSpace(ctx, props, 'after')) {
      // Automatic spacing does not add up: the larger of the two spaces is what separates the paragraphs.
      // Between the items of one list there is none.
      const sameList = list !== 0 && listOf(props) === list;
      const above = sameList && isAutoSpace(ctx, props, 'after') ? 0 : paragraphSpace(ctx, props, 'after');
      env.before = sameList && prep.beforeAuto ? 0 : Math.max(0, env.before - above);
    }
  } else if (!before && where && prep.beforeAuto) env.before = 0;
  if (after && after.type === 'paragraph') {
    const props = ctx.paraProps(after.props, where);
    if (key !== '' && borderKey(props) === key && !prep.para.props.section) env.joinNext = true;
    if (prep.contextual && props.style === prep.props.style) env.after = 0;
    else if (prep.afterAuto && list !== 0 && listOf(props) === list) env.after = 0;
  } else if (!after && where && prep.afterAuto) env.after = 0;
  return env;
}

/**
 * The bands of a list of blocks set to `width`. `path` is the path of the
 * list (the blocks are path + index); `where` is given for the content of a
 * table cell.
 */
export function layoutBlocks(
  ctx: LayoutContext,
  blocks: Block[],
  width: number,
  story: StoryRef,
  path: number[],
  where: CellWhere | undefined,
  depth: number,
  fields: FieldValues | undefined,
): Band[] {
  let out: Band[] | undefined;
  let single: Band[] | undefined;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    let bands: Band[];
    if (block.type === 'paragraph') {
      const prep = getPrep(ctx, block, story, path.concat([i]), where, depth, fields);
      bands = paragraphBands(ctx, prep, width, paragraphEnv(ctx, blocks, i, prep, where, fields));
    } else if (block.type === 'table') bands = tableBands(ctx, block, width, story, path.concat([i]), depth, fields);
    else continue;
    // The usual cell holds one paragraph: its bands are handed on without a copy.
    if (!single && !out) single = bands;
    else {
      if (!out) out = single!.slice();
      for (let k = 0; k < bands.length; k++) out.push(bands[k]);
    }
  }
  return out || single || [];
}
