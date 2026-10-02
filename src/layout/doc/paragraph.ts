// Lays out a prepared paragraph: lines, their heights, alignment, borders and
// shading. The result is a list of bands, one per line, which pagination
// places without looking into them. It depends on the width and on a few
// facts about the neighbours only, so it is kept with the paragraph and used
// again when pagination runs again.

import type { Band, FieldTextItem, FieldValues, FloatSpec, LayoutContext, ParaRef } from './context';
import { pushShifted } from './context';
import { breakLine, type Line, type Piece } from './line-break';
import { borderRoom, type ParaPrep } from './prep';
import { LEADER } from './symbols';
import type { PageItem } from './types';
import { EPSILON } from './units';

/** What a paragraph is laid out in, beyond its width. */
export interface ParaEnv {
  /** The neighbour above/below has the same borders: the two share one box. */
  joinPrev: boolean;
  joinNext: boolean;
  /** Space before and after the paragraph, px: its own, less what the neighbours take away (see paragraphEnv). */
  before: number;
  after: number;
  fields?: FieldValues;
  /** Line pitch of the document grid, px, when lines snap to it. */
  grid?: number;
}

/** Drawings the lines have to keep clear of. */
export interface WrapEnv {
  page: number;
  column: number;
  /** y of the slot the first band goes into (before its space before). */
  y: number;
  /** The space before of the first band does not apply (top of a column). */
  dropBefore: boolean;
  /**
   * The part of the column a line from `top` to `bottom` may use, as offsets
   * from the left edge of the column; a number when the line has to move down
   * to that y instead.
   */
  free(top: number, bottom: number): { left: number; right: number } | number;
}

interface LineBox {
  height: number;
  baseline: number;
}

function lineBox(ctx: LayoutContext, prep: ParaPrep, line: Line, grid: number | undefined): LineBox {
  let ascent = 0;
  let descent = 0;
  let gap = 0;
  /** How far boxes (pictures) reach above the baseline. */
  let boxes = 0;
  const add = (font: Piece['font'], shift: number) => {
    const m = ctx.metrics(font);
    if (m.ascent + shift > ascent) ascent = m.ascent + shift;
    if (m.descent - shift > descent) descent = m.descent - shift;
    if (m.lineGap > gap) gap = m.lineGap;
  };
  const pieces = line.pieces;
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    if (piece.kind === 'box') {
      // A box stands on the baseline like a character of its run: the descent of the font stays below it.
      const top = (piece.height || 0) + piece.style.shift;
      if (top > boxes) boxes = top;
      const m = ctx.metrics(piece.font);
      if (m.descent - piece.style.shift > descent) descent = m.descent - piece.style.shift;
    } else add(piece.font, piece.style.shift);
  }
  // The paragraph mark sits on the last line; a line that only holds a break has the height of the break.
  if (line.last) add(prep.mark.font, 0);
  else if (pieces.length === 0) add(line.brkFont || prep.mark.font, 0);

  // Line spacing is a matter of the text: a picture makes the line as much taller as it is taller than the text.
  const text = ascent + descent + gap;
  if (boxes > ascent) ascent = boxes;
  const natural = ascent + descent + gap;
  let height: number;
  let baseline: number;
  if (prep.rule === 'exact' && prep.line > 0) {
    height = prep.line;
    baseline = height - descent;
  } else if (prep.rule === 'atLeast') {
    height = Math.max(natural, prep.line);
    baseline = height - descent - gap;
  } else {
    // A multiple of the natural height of the text. More room goes below the text, less is taken from above it.
    height = Math.max(natural + text * (prep.line - 1), boxes);
    baseline = ascent - Math.max(0, natural - height);
  }
  if (grid !== undefined && grid > 0 && prep.rule !== 'exact' && prep.props.snapToGrid !== false) {
    const snapped = Math.ceil(height / grid - EPSILON) * grid;
    baseline += (snapped - height) / 2;
    height = snapped;
  }
  return { height, baseline };
}

/** Moves and stretches the pieces of a line for the alignment of the paragraph. */
function alignLine(prep: ParaPrep, line: Line, x1: number, stretch: boolean): boolean {
  const extra = x1 - line.right;
  if (!(extra > EPSILON) || prep.align === 'left') return false;
  const pieces = line.pieces;
  if (prep.align === 'both') {
    if (!stretch) return false;
    let spaces = 0;
    for (let i = line.free; i < pieces.length - 1; i++) spaces += pieces[i].spaces;
    if (spaces === 0) return false;
    const each = extra / spaces;
    let moved = 0;
    for (let i = line.free; i < pieces.length; i++) {
      const piece = pieces[i];
      piece.x += moved;
      if (i < pieces.length - 1 && piece.spaces) {
        const more = each * piece.spaces;
        piece.width += more;
        piece.trail += more;
        moved += more;
      }
    }
    return true;
  }
  const shift = prep.align === 'center' ? extra / 2 : extra;
  for (let i = 0; i < pieces.length; i++) pieces[i].x += shift;
  return false;
}

/** The items of one line, relative to the left edge of the column and to `top`. */
function lineItems(ctx: LayoutContext, prep: ParaPrep, line: Line, top: number, baseline: number, justified: boolean, out: PageItem[]): void {
  const pieces = line.pieces;
  let open: FieldTextItem | undefined;
  let openPiece: Piece | undefined;
  /** Offset just past the text of the open item, -1 when it has no source. */
  let openEnd = -1;
  for (let i = 0; i < pieces.length; i++) {
    const piece = pieces[i];
    const style = piece.style;
    const y = top + baseline - style.shift;
    if (piece.kind === 'box') {
      open = undefined;
      if (piece.items) pushShifted(out, piece.items, piece.x, y - (piece.height || 0));
      continue;
    }
    let text = piece.text;
    let x = piece.x;
    let width = i === pieces.length - 1 ? piece.width - piece.trail : piece.width;
    if (piece.kind === 'tab') {
      open = undefined;
      const leader = piece.leader === undefined ? undefined : LEADER[piece.leader];
      if (leader !== undefined) {
        const each = ctx.width(piece.font, leader);
        const count = each > 0 ? Math.min(Math.floor((width + EPSILON) / each), 2000) : 0;
        if (count <= 0) continue;
        text = '';
        for (let k = 0; k < count; k++) text += leader;
        // The dots end at the tab stop.
        x += width - count * each;
        width = count * each;
      } else if (!style.underline && !style.strike && style.highlight === undefined) continue;
    }
    const mergeable = piece.kind === 'text' && !piece.label && piece.field === undefined;
    if (
      open &&
      openPiece &&
      mergeable &&
      openPiece.style === style &&
      openPiece.font === piece.font &&
      openPiece.link === piece.link &&
      (piece.offset < 0 ? openEnd < 0 : piece.offset === openEnd) &&
      !(justified && openPiece.spaces > 0) &&
      Math.abs(open.x + open.width - x) < EPSILON
    ) {
      open.text += text;
      open.width = x + width - open.x;
      openPiece = piece;
      if (openEnd >= 0) openEnd += text.length;
      continue;
    }
    const item: FieldTextItem = { type: 'text', x, y, width, text, font: piece.font, color: style.color };
    if (style.underline) item.underline = style.underline;
    if (style.strike) item.strike = style.strike;
    if (style.highlight !== undefined) item.highlight = style.highlight;
    if (piece.link !== undefined) item.link = piece.link;
    if (style.spacing !== 0) item.letterSpacing = style.spacing;
    if (style.scale !== 1) item.scale = style.scale;
    if (piece.field !== undefined) {
      item.field = piece.field;
      if (piece.fieldFormat !== undefined) item.fieldFormat = piece.fieldFormat;
    }
    if (piece.offset >= 0 && piece.kind === 'text') item.source = { story: prep.story, path: prep.path, offset: piece.offset };
    out.push(item);
    if (mergeable) {
      open = item;
      openPiece = piece;
      openEnd = piece.offset < 0 ? -1 : piece.offset + text.length;
    } else open = undefined;
  }
}

/**
 * The bands of a paragraph from the given atom on (the whole paragraph when
 * `from` is missing). With `wrap`, lines keep clear of drawings and the bands
 * are only good for the place `wrap` describes.
 */
export function paragraphBands(
  ctx: LayoutContext,
  prep: ParaPrep,
  width: number,
  env: ParaEnv,
  wrap?: WrapEnv,
  from?: { atom: number; ch: number },
): Band[] {
  const whole = !from || (from.atom === 0 && from.ch === 0);
  const cacheable = !wrap && whole && (!prep.live || !env.fields);
  const key = cacheable
    ? width + (env.joinPrev ? 'p' : '') + (env.joinNext ? 'n' : '') + 'b' + env.before + 'a' + env.after + (env.grid === undefined ? '' : 'g' + env.grid)
    : '';
  if (cacheable && prep.cache && prep.cache.key === key) return prep.cache.bands;

  const borders = prep.borders;
  const topLine = borders ? (env.joinPrev ? borders.between : borders.top) : undefined;
  const bottomLine = borders && !env.joinNext ? borders.bottom : undefined;
  const topRoom = whole ? borderRoom(topLine) : 0;
  const before = !whole || (wrap && wrap.dropBefore) ? 0 : env.before;
  const after = env.after;
  // A paragraph mark that follows a page break stays on the page of the break in documents of Word 2013 and later.
  const markStays = ctx.compatMode >= 15 && !ctx.doc.settings.compat.get('splitPgBreakAndParaMark');
  const stretchBreaks = !ctx.doc.settings.compat.get('doNotExpandShiftReturn');

  // Text area of the paragraph, never narrower than a sliver however large the indents are.
  let left = prep.left;
  let right = width - prep.right;
  if (right - left < 1) {
    left = Math.max(0, Math.min(left, width - 1));
    right = Math.max(left + 1, Math.min(width, right));
  }
  const boxLeft = Math.min(left, left + prep.first) - (borders && borders.left ? borders.left.space : 0);
  const boxRight = right + (borders && borders.right ? borders.right.space : 0);

  const bands: Band[] = [];
  let atom = from ? from.atom : 0;
  let ch = from ? from.ch : 0;
  let slot = wrap ? wrap.y : 0;
  // Every line takes at least one character or atom, so this many lines is more than the paragraph can have.
  let guard = prep.atoms.length * 2 + 8;
  for (let i = 0; i < prep.atoms.length; i++) {
    const a = prep.atoms[i];
    if (a.t === 'text') guard += a.text.length;
  }
  let first = true;
  for (;;) {
    const firstLine = atom === 0 && ch === 0;
    let x0 = left + (firstLine ? prep.first : 0);
    let x1 = right;
    const lead = first ? before + topRoom : 0;
    let gap = 0;
    let line: Line;
    let box: LineBox;
    if (wrap) {
      // Find room beside the drawings for a line of the height of the paragraph mark, then check with the real height.
      const m = ctx.metrics(prep.mark.font);
      let estimate = (m.ascent + m.descent + m.lineGap) * (prep.rule === 'auto' ? prep.line : 1);
      if (prep.rule === 'exact' && prep.line > 0) estimate = prep.line;
      let tries = 0;
      for (;;) {
        const top = slot + lead + gap;
        const room = wrap.free(top, top + estimate);
        if (typeof room === 'number') {
          if (room > top && tries++ < 64) {
            gap += room - top;
            continue;
          }
          line = breakLine(ctx, prep, atom, ch, x0, x1, env.fields, markStays);
          box = lineBox(ctx, prep, line, env.grid);
          break;
        }
        const l = Math.max(left, room.left) + (firstLine && prep.first > 0 ? prep.first : 0);
        const from0 = room.left > left ? l : x0;
        const to = Math.min(right, room.right);
        line = breakLine(ctx, prep, atom, ch, from0, Math.max(to, from0 + 1), env.fields, markStays);
        box = lineBox(ctx, prep, line, env.grid);
        if (box.height > estimate + 0.5 && tries++ < 64) {
          const again = wrap.free(top, top + box.height);
          if (typeof again === 'number' || again.left > room.left + EPSILON || again.right < room.right - EPSILON) {
            estimate = box.height;
            continue;
          }
        }
        x0 = from0;
        x1 = Math.max(to, from0 + 1);
        break;
      }
    } else {
      line = breakLine(ctx, prep, atom, ch, x0, x1, env.fields, markStays);
      box = lineBox(ctx, prep, line, env.grid);
    }
    const stretch = !line.last && (line.brk === undefined || (line.brk === 'line' && stretchBreaks));
    const justified = alignLine(prep, line, x1, stretch);

    const bottomRoom = line.last ? borderRoom(bottomLine) : 0;
    const top = gap + (first ? topRoom : 0);
    const height = top + box.height + bottomRoom;
    const items: PageItem[] = [];
    if (prep.shade !== undefined && boxRight > boxLeft) items.push({ type: 'rect', x: boxLeft, y: gap, width: boxRight - boxLeft, height: height - gap, fill: prep.shade });
    if (borders) {
      const up = first && env.joinPrev ? -before : gap;
      const down = line.last && env.joinNext ? height + after : height;
      if (borders.left) {
        const b = borders.left;
        const x = boxLeft - b.width / 2;
        items.push({ type: 'line', x1: x, y1: up, x2: x, y2: down, color: b.color, width: b.width, style: b.style });
      }
      if (borders.right) {
        const b = borders.right;
        const x = boxRight + b.width / 2;
        items.push({ type: 'line', x1: x, y1: up, x2: x, y2: down, color: b.color, width: b.width, style: b.style });
      }
      if (first && topLine) {
        const y = gap + topLine.width / 2;
        items.push({ type: 'line', x1: boxLeft, y1: y, x2: boxRight, y2: y, color: topLine.color, width: topLine.width, style: topLine.style });
      }
      if (line.last && bottomLine) {
        const y = height - bottomLine.width / 2;
        items.push({ type: 'line', x1: boxLeft, y1: y, x2: boxRight, y2: y, color: bottomLine.color, width: bottomLine.width, style: bottomLine.style });
      }
    }
    lineItems(ctx, prep, line, top, box.baseline, justified, items);

    const ref: ParaRef = { prep, atom, ch, width, env };
    if (wrap) ref.slot = { page: wrap.page, column: wrap.column, y: slot };
    const band: Band = { height, before: first ? before : 0, after: line.last ? after : 0, items, keep: false, para: ref };
    if (line.brk === 'page' || line.brk === 'column') band.breakAfter = line.brk;
    for (let i = 0; i < line.pieces.length; i++) {
      const note = line.pieces[i].note;
      if (note !== undefined) (band.notes || (band.notes = [])).push(note);
    }
    bands.push(band);
    slot += band.before + band.height + band.after;
    first = false;
    atom = line.atom;
    ch = line.ch;
    if (line.last || --guard <= 0) break;
  }

  // Where the paragraph may be divided between pages.
  const count = bands.length;
  for (let i = 0; i < count - 1; i++) {
    const band = bands[i];
    if (band.breakAfter) continue;
    if (prep.keepLines) band.keep = true;
    else if (prep.widow) {
      // Neither a single first line at the bottom of a page nor a single last line at the top of one.
      if ((whole && i === 0) || i === count - 2) band.keep = true;
    }
  }
  if (prep.keepNext && !bands[count - 1].breakAfter) bands[count - 1].keep = true;
  if (whole) {
    if (prep.pageBreakBefore) bands[0].breakBefore = 'page';
    if (prep.floats) {
      const floats: FloatSpec[] = [];
      for (const float of prep.floats) floats.push({ ...float, columnWidth: width });
      bands[0].floats = floats;
    }
  }
  if (cacheable) prep.cache = { key, bands };
  return bands;
}
