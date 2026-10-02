// Sections and the geometry of their pages: paper, margins, columns, which
// header and footer a page gets, page borders.

import type { DocModel, HeaderFooterKind, SectionProps } from '../../model/doc/types';
import type { LayoutContext } from './context';
import type { PageItem } from './types';
import {
  DEFAULT_COLUMN_SPACE,
  DEFAULT_HEADER_DISTANCE,
  DEFAULT_MARGIN,
  DEFAULT_PAGE_HEIGHT,
  DEFAULT_PAGE_WIDTH,
  MAX_COLUMNS,
  MAX_PAGE_SIDE,
  MIN_PAGE_SIDE,
  MIN_TEXT_EXTENT,
  clamp,
  twips,
} from './units';

export interface SectionInfo {
  props: SectionProps;
  /** Blocks of the body that belong to the section: [start, end). */
  start: number;
  end: number;
}

const NO_SECTION: SectionProps = Object.freeze({});

/** The sections of the body, in order. There is always at least one. */
export function findSections(doc: DocModel): SectionInfo[] {
  const blocks = doc.body.blocks;
  const out: SectionInfo[] = [];
  let start = 0;
  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    if (block.type !== 'paragraph' || !block.props.section) continue;
    out.push({ props: block.props.section, start, end: i + 1 });
    start = i + 1;
  }
  // The section mark of the body stands for everything after the last paragraph that ends a section.
  out.push({ props: doc.body.section || NO_SECTION, start, end: blocks.length });
  return out;
}

export interface PageGeometry {
  width: number;
  height: number;
  /** The text area: x of its left and right edge, y of its top and bottom. */
  left: number;
  right: number;
  top: number;
  bottom: number;
  /** The margin is exact: a tall header does not push the text down (a negative margin in the file). */
  fixedTop: boolean;
  fixedBottom: boolean;
  /** y where the header starts, and where the footer ends. */
  headerTop: number;
  footerBottom: number;
  columns: { x: number; width: number }[];
  separator: boolean;
}

/** Page geometry of a section. `mirrored`: the page is an even one of a document with mirrored margins. */
export function pageGeometry(props: SectionProps, mirrored: boolean): PageGeometry {
  const page = props.page || {};
  const margins = props.margins || {};
  const side = (value: number | undefined, fallback: number) => {
    const px = twips(value === undefined || !(value > 0) ? fallback : value);
    return clamp(px, MIN_PAGE_SIDE, MAX_PAGE_SIDE);
  };
  const width = side(page.width, DEFAULT_PAGE_WIDTH);
  const height = side(page.height, DEFAULT_PAGE_HEIGHT);
  const margin = (value: number | undefined, fallback: number) => Math.abs(twips(value === undefined ? fallback : value));

  let left = margin(margins.left, DEFAULT_MARGIN) + Math.max(0, twips(margins.gutter));
  let right = margin(margins.right, DEFAULT_MARGIN);
  if (mirrored) {
    const swap = left;
    left = right;
    right = swap;
  }
  // Margins that leave no room for text are scaled down until some is left.
  if (left + right > width - MIN_TEXT_EXTENT) {
    const factor = (width - MIN_TEXT_EXTENT) / (left + right);
    left *= factor;
    right *= factor;
  }
  let top = margin(margins.top, DEFAULT_MARGIN);
  let bottom = margin(margins.bottom, DEFAULT_MARGIN);
  if (top + bottom > height - MIN_TEXT_EXTENT) {
    const factor = (height - MIN_TEXT_EXTENT) / (top + bottom);
    top *= factor;
    bottom *= factor;
  }
  const geometry: PageGeometry = {
    width,
    height,
    left,
    right: width - right,
    top,
    bottom: height - bottom,
    fixedTop: margins.top !== undefined && margins.top < 0,
    fixedBottom: margins.bottom !== undefined && margins.bottom < 0,
    headerTop: clamp(margin(margins.header, DEFAULT_HEADER_DISTANCE), 0, height),
    footerBottom: height - clamp(margin(margins.footer, DEFAULT_HEADER_DISTANCE), 0, height),
    columns: [],
    separator: false,
  };
  geometry.columns = sectionColumns(props, left, width - right);
  geometry.separator = props.columns !== undefined && props.columns.separator === true && geometry.columns.length > 1;
  return geometry;
}

/** The columns of a section between two x positions. */
export function sectionColumns(props: SectionProps, left: number, right: number): { x: number; width: number }[] {
  const content = right - left;
  const spec = props.columns;
  const single = [{ x: left, width: content }];
  if (!spec) return single;
  const listed = spec.columns;
  if (spec.equalWidth === false && listed && listed.length > 0) {
    const count = Math.min(listed.length, MAX_COLUMNS);
    let total = 0;
    const widths: number[] = [];
    const spaces: number[] = [];
    for (let i = 0; i < count; i++) {
      widths.push(Math.max(0, twips(listed[i].width)));
      spaces.push(i < count - 1 ? Math.max(0, twips(listed[i].space)) : 0);
      total += widths[i] + spaces[i];
    }
    if (count === 1 || !(total > 0)) return single;
    const factor = total > content ? content / total : 1;
    const out: { x: number; width: number }[] = [];
    let x = left;
    for (let i = 0; i < count; i++) {
      // A column of no width still holds text, a character to the line.
      out.push({ x, width: Math.max(1, widths[i] * factor) });
      x += (widths[i] + spaces[i]) * factor;
    }
    return out;
  }
  let count = clamp(spec.count === undefined ? 1 : spec.count, 1, MAX_COLUMNS);
  if (count === 1) return single;
  let space = Math.max(0, twips(spec.space === undefined ? DEFAULT_COLUMN_SPACE : spec.space));
  let width = (content - (count - 1) * space) / count;
  if (width < MIN_TEXT_EXTENT) {
    // Too many columns for the page: first the space between them goes, then columns do.
    space = 0;
    count = Math.max(1, Math.min(count, Math.floor(content / MIN_TEXT_EXTENT)));
    width = content / count;
    if (count === 1) return single;
  }
  const out: { x: number; width: number }[] = [];
  for (let i = 0; i < count; i++) out.push({ x: left + i * (width + space), width });
  return out;
}

/**
 * The header or footer part a page of a section uses: its own for the kind,
 * or that of the nearest earlier section that has one. undefined when no
 * section up to this one names a part of the kind.
 */
export function inheritedPart(sections: SectionInfo[], index: number, list: 'headers' | 'footers', kind: HeaderFooterKind): string | undefined {
  for (let i = Math.min(index, sections.length - 1); i >= 0; i--) {
    const parts = sections[i].props[list];
    const id = parts && parts[kind];
    if (id !== undefined) return id;
  }
  return undefined;
}

/** The lines of the page border of a section on one page. */
export function pageBorderItems(ctx: LayoutContext, props: SectionProps, geometry: PageGeometry, firstOfSection: boolean, out: PageItem[]): void {
  const borders = props.pageBorders;
  if (!borders) return;
  if ((borders.display === 'firstPage' && !firstOfSection) || (borders.display === 'notFirstPage' && firstOfSection)) return;
  const fromPage = borders.offsetFrom === 'page';
  const top = ctx.border(borders.top);
  const left = ctx.border(borders.left);
  const bottom = ctx.border(borders.bottom);
  const right = ctx.border(borders.right);
  const x0 = fromPage ? (left ? left.space : 0) : geometry.left - (left ? left.space : 0);
  const x1 = fromPage ? geometry.width - (right ? right.space : 0) : geometry.right + (right ? right.space : 0);
  const y0 = fromPage ? (top ? top.space : 0) : geometry.top - (top ? top.space : 0);
  const y1 = fromPage ? geometry.height - (bottom ? bottom.space : 0) : geometry.bottom + (bottom ? bottom.space : 0);
  if (top) out.push({ type: 'line', x1: x0, y1: y0, x2: x1, y2: y0, color: top.color, width: top.width, style: top.style });
  if (bottom) out.push({ type: 'line', x1: x0, y1: y1, x2: x1, y2: y1, color: bottom.color, width: bottom.width, style: bottom.style });
  if (left) out.push({ type: 'line', x1: x0, y1: y0, x2: x0, y2: y1, color: left.color, width: left.width, style: left.style });
  if (right) out.push({ type: 'line', x1: x1, y1: y0, x2: x1, y2: y1, color: right.color, width: right.width, style: right.style });
}
