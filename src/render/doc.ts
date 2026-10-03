// Paints one laid out page of a document on a 2D canvas context.
//
// Everything on the page comes from a file: text only ever reaches the canvas
// through fillText, colours are checked before they are used, and numbers
// that are not finite paint nothing. Newer canvas features (letterSpacing)
// are used where they exist and replaced by plain drawing where they do not.

import type { ImageItem, LayoutFont, LayoutPage, LineItem, RectItem, ShapeItem, TextItem } from '../layout/doc/types';
import { docCssFont } from './doc-measure';

/** A decoded picture and its size in its own pixels (which may be fewer than the file has). */
export interface PageImage {
  source: CanvasImageSource;
  width: number;
  height: number;
}

export interface RenderPageOptions {
  /** Canvas pixels per px of the page: zoom times device pixel ratio. Default 1. */
  scale?: number;
  /**
   * The decoded picture of a part, undefined while it is not loaded (a
   * placeholder is drawn). `width` and `height` say how many canvas pixels the
   * whole picture covers where it is drawn, so that a store can decode it to
   * that size instead of the size it has in the file.
   */
  images?: { get(part: string, width?: number, height?: number): PageImage | undefined };
  /** The CSS font for a layout font; must be the one the page was measured with. */
  cssFont?(font: LayoutFont): string;
  /** Line extent of a font, for highlights; estimated from the size when missing. */
  metrics?(font: LayoutFont): { ascent: number; descent: number };
  /** Paper colour; null leaves the canvas as it is. Default white. */
  background?: string | null;
  /** Words shown inside a placeholder (what stands there, or its alternative text). */
  placeholderLabel?(item: ImageItem): string | undefined;
}

const COLOR = /^#[0-9a-fA-F]{6}$/;
const safeColor = (color: unknown, fallback: string): string => (typeof color === 'string' && COLOR.test(color) ? color : fallback);

const finite = (value: unknown): value is number => typeof value === 'number' && value === value && value !== Infinity && value !== -Infinity;

const PLACEHOLDER_FILL = '#f1f3f4';
const PLACEHOLDER_LINE = '#bdc1c6';
const PLACEHOLDER_TEXT = '#5f6368';
const PLACEHOLDER_FONT = '11px system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif';
/** Waves and dashes are cut off after this many segments, whatever a line claims to be long. */
const MAX_SEGMENTS = 4000;

type Context = CanvasRenderingContext2D;

/** Newer canvases space letters themselves. (A function, so that the type of the context is not narrowed away where they do not.) */
const hasLetterSpacing = (ctx: object): boolean => 'letterSpacing' in ctx;

interface Painter {
  ctx: Context;
  scale: number;
  options: RenderPageOptions;
  /** Font and fill last assigned, so that they are only assigned when they change. */
  font: string | undefined;
  fill: string | undefined;
}

function setFill(p: Painter, fill: string): void {
  if (p.fill === fill) return;
  p.ctx.fillStyle = fill;
  p.fill = fill;
}

function setFont(p: Painter, font: string): void {
  if (p.font === font) return;
  p.ctx.font = font;
  p.font = font;
}

/** After save/restore or a direct assignment the tracked values are no longer known. */
function forget(p: Painter): void {
  p.font = undefined;
  p.fill = undefined;
}

/**
 * A stroke of `thickness` centred on `position`, moved and widened so that it
 * covers whole canvas pixels: thin lines stay sharp at every zoom.
 */
function snap(position: number, thickness: number, scale: number): { start: number; size: number } {
  const size = Math.max(1, Math.round(thickness * scale));
  const start = Math.round(position * scale - size / 2);
  return { start: start / scale, size: size / scale };
}

const snapPoint = (value: number, scale: number): number => Math.round(value * scale) / scale;

function setDash(ctx: Context, dash: number[] | undefined): void {
  if (typeof ctx.setLineDash === 'function') ctx.setLineDash(dash || []);
}

/** One straight stroke; `dash` in px of the page. */
function stroke(p: Painter, x1: number, y1: number, x2: number, y2: number, thickness: number, color: string, dash?: number[]): void {
  const ctx = p.ctx;
  if (y1 === y2 || x1 === x2) {
    const horizontal = y1 === y2;
    const across = snap(horizontal ? y1 : x1, thickness, p.scale);
    const from = snapPoint(Math.min(horizontal ? x1 : y1, horizontal ? x2 : y2), p.scale);
    const to = snapPoint(Math.max(horizontal ? x1 : y1, horizontal ? x2 : y2), p.scale);
    if (!(to > from)) return;
    if (!dash || typeof ctx.setLineDash !== 'function') {
      setFill(p, color);
      if (horizontal) ctx.fillRect(from, across.start, to - from, across.size);
      else ctx.fillRect(across.start, from, across.size, to - from);
      return;
    }
    const centre = across.start + across.size / 2;
    ctx.strokeStyle = color;
    ctx.lineWidth = across.size;
    ctx.setLineDash(dash);
    ctx.beginPath();
    if (horizontal) {
      ctx.moveTo(from, centre);
      ctx.lineTo(to, centre);
    } else {
      ctx.moveTo(centre, from);
      ctx.lineTo(centre, to);
    }
    ctx.stroke();
    ctx.setLineDash([]);
    return;
  }
  ctx.strokeStyle = color;
  ctx.lineWidth = Math.max(thickness, 1 / p.scale);
  setDash(ctx, dash);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
  setDash(ctx, undefined);
}

/** A wavy stroke from (x1, y1) to (x2, y2). */
function wave(p: Painter, x1: number, y1: number, x2: number, y2: number, thickness: number, color: string): void {
  const ctx = p.ctx;
  const length = Math.sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
  if (!(length > 0)) return;
  const amplitude = Math.max(thickness * 1.25, 1);
  let half = Math.max(thickness * 3, 2);
  let count = Math.ceil(length / half);
  if (count > MAX_SEGMENTS) {
    count = MAX_SEGMENTS;
    half = length / count;
  }
  const ux = (x2 - x1) / length;
  const uy = (y2 - y1) / length;
  ctx.strokeStyle = color;
  ctx.lineWidth = Math.max(thickness, 1 / p.scale);
  setDash(ctx, undefined);
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  for (let i = 0; i < count; i++) {
    const end = Math.min((i + 1) * half, length);
    const mid = (i * half + end) / 2;
    const side = i % 2 === 0 ? -amplitude : amplitude;
    // The control point lies beside the line, on alternating sides.
    ctx.quadraticCurveTo(x1 + ux * mid - uy * side * 2, y1 + uy * mid + ux * side * 2, x1 + ux * end, y1 + uy * end);
  }
  ctx.stroke();
}

/** Dash patterns in units of the line thickness (at least 1 px). */
const DASHES: { [style: string]: number[] } = {
  dotted: [1, 1],
  dashed: [4, 3],
  dashSmallGap: [4, 1.5],
  dotDash: [4, 2, 1, 2],
  dotDotDash: [4, 2, 1, 2, 1, 2],
  dashDotStroked: [4, 2, 1, 2],
};

/**
 * Borders made of several parallel lines: thickness of each line and the gaps
 * between them, in units of the border width, from the top (or left) down.
 */
const COMPOUND: { [style: string]: number[] } = {
  double: [1, 1, 1],
  triple: [1, 1, 1, 1, 1],
  thinThickSmallGap: [1 / 3, 1 / 3, 1],
  thickThinSmallGap: [1, 1 / 3, 1 / 3],
  thinThickThinSmallGap: [1 / 3, 1 / 3, 1, 1 / 3, 1 / 3],
  thinThickMediumGap: [1 / 3, 2 / 3, 1],
  thickThinMediumGap: [1, 2 / 3, 1 / 3],
  thinThickThinMediumGap: [1 / 3, 2 / 3, 1, 2 / 3, 1 / 3],
  thinThickLargeGap: [1 / 3, 1, 1],
  thickThinLargeGap: [1, 1, 1 / 3],
  thinThickThinLargeGap: [1 / 3, 1, 1, 1, 1 / 3],
};

function drawLine(p: Painter, item: LineItem): void {
  const { x1, y1, x2, y2 } = item;
  if (!finite(x1) || !finite(y1) || !finite(x2) || !finite(y2)) return;
  const style = item.style;
  if (style === 'none' || style === 'nil') return;
  const width = finite(item.width) && item.width > 0 ? Math.min(item.width, 64) : 1;
  const color = safeColor(item.color, '#000000');
  const length = Math.sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1));
  if (!(length > 0)) return;

  if (style === 'wave' || style === 'doubleWave') {
    if (style === 'wave') return wave(p, x1, y1, x2, y2, width, color);
    const nx = (-(y2 - y1) / length) * width * 1.5;
    const ny = ((x2 - x1) / length) * width * 1.5;
    wave(p, x1 - nx, y1 - ny, x2 - nx, y2 - ny, width, color);
    wave(p, x1 + nx, y1 + ny, x2 + nx, y2 + ny, width, color);
    return;
  }

  const parts = Object.prototype.hasOwnProperty.call(COMPOUND, style) ? COMPOUND[style] : undefined;
  if (parts) {
    // The lines lie side by side across the border, centred on it.
    let total = 0;
    for (let i = 0; i < parts.length; i++) total += parts[i] * width;
    let nx = -(y2 - y1) / length;
    let ny = (x2 - x1) / length;
    // Downwards, or to the right for a vertical border, whichever way the line was given.
    if (ny < 0 || (ny === 0 && nx < 0)) {
      nx = -nx;
      ny = -ny;
    }
    let offset = -total / 2;
    for (let i = 0; i < parts.length; i++) {
      const size = parts[i] * width;
      if (i % 2 === 0) {
        const centre = offset + size / 2;
        stroke(p, x1 + nx * centre, y1 + ny * centre, x2 + nx * centre, y2 + ny * centre, size, color);
      }
      offset += size;
    }
    return;
  }

  const pattern = Object.prototype.hasOwnProperty.call(DASHES, style) ? DASHES[style] : undefined;
  if (pattern) {
    const unit = Math.max(width, 1 / p.scale);
    const dash: number[] = [];
    for (let i = 0; i < pattern.length; i++) dash.push(pattern[i] * unit);
    return stroke(p, x1, y1, x2, y2, width, color, dash);
  }
  // single, thick, the 3D styles and every style this painter does not know.
  stroke(p, x1, y1, x2, y2, width, color);
}

function drawRect(p: Painter, item: RectItem): void {
  if (!finite(item.x) || !finite(item.y) || !(item.width > 0) || !(item.height > 0) || !finite(item.width) || !finite(item.height)) return;
  setFill(p, safeColor(item.fill, '#ffffff'));
  // Whole canvas pixels: neighbouring shadings neither overlap nor leave a seam.
  const left = snapPoint(item.x, p.scale);
  const top = snapPoint(item.y, p.scale);
  p.ctx.fillRect(left, top, snapPoint(item.x + item.width, p.scale) - left, snapPoint(item.y + item.height, p.scale) - top);
}

/** The preset dashes of DrawingML, in units of the line thickness: a stroke, a gap, ... */
const SHAPE_DASHES: { [name: string]: number[] } = {
  sysDot: [1, 1],
  sysDash: [3, 1],
  sysDashDot: [3, 1, 1, 1],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
  dot: [1, 3],
  dash: [4, 3],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
};

/** How far the control points of a cubic curve lie from its ends for a quarter of a circle of radius 1. */
const QUARTER = 0.5522847498;

/** A shape: its outline as a path around the centre of its box, filled and then stroked. */
function drawShape(p: Painter, item: ShapeItem): void {
  const { x, y, width, height } = item;
  if (!finite(x) || !finite(y) || !finite(width) || !finite(height) || width < 0 || height < 0) return;
  const line = item.geometry === 'line';
  if (line ? !(width > 0 || height > 0) : !(width > 0 && height > 0)) return;
  const ctx = p.ctx;
  const fill = !line && typeof item.fill === 'string' && COLOR.test(item.fill) ? item.fill : undefined;
  const outline = item.stroke;
  const thickness = outline && finite(outline.width) && outline.width > 0 ? Math.min(outline.width, 64) : 0;
  if (fill === undefined && thickness === 0) return;
  const color = outline ? safeColor(outline.color, '#000000') : '#000000';
  let dash: number[] | undefined;
  if (outline && typeof outline.dash === 'string' && Object.prototype.hasOwnProperty.call(SHAPE_DASHES, outline.dash)) {
    const unit = Math.max(thickness, 1 / p.scale);
    dash = SHAPE_DASHES[outline.dash].map((part) => part * unit);
  }
  const rotation = finite(item.rotation) ? item.rotation % 360 : 0;
  const up = line && (item.flipH === true) !== (item.flipV === true);
  if (line && rotation === 0 && (width === 0 || height === 0)) {
    // A horizontal or vertical line is laid on whole canvas pixels, as borders are.
    if (thickness > 0) stroke(p, x, y, x + width, y + height, thickness, color, dash);
    return;
  }
  const w = width / 2;
  const h = height / 2;
  ctx.save();
  try {
    ctx.translate(x + w, y + h);
    if (rotation !== 0) ctx.rotate((rotation * Math.PI) / 180);
    ctx.beginPath();
    if (line) {
      ctx.moveTo(-w, up ? h : -h);
      ctx.lineTo(w, up ? -h : h);
    } else if (item.geometry === 'ellipse') {
      ctx.moveTo(w, 0);
      ctx.bezierCurveTo(w, h * QUARTER, w * QUARTER, h, 0, h);
      ctx.bezierCurveTo(-w * QUARTER, h, -w, h * QUARTER, -w, 0);
      ctx.bezierCurveTo(-w, -h * QUARTER, -w * QUARTER, -h, 0, -h);
      ctx.bezierCurveTo(w * QUARTER, -h, w, -h * QUARTER, w, 0);
      ctx.closePath();
    } else {
      const r = item.geometry === 'roundRect' && finite(item.radius) && item.radius > 0 ? Math.min(item.radius, w, h) : 0;
      if (r > 0) {
        const c = r * (1 - QUARTER);
        ctx.moveTo(-w + r, -h);
        ctx.lineTo(w - r, -h);
        ctx.bezierCurveTo(w - c, -h, w, -h + c, w, -h + r);
        ctx.lineTo(w, h - r);
        ctx.bezierCurveTo(w, h - c, w - c, h, w - r, h);
        ctx.lineTo(-w + r, h);
        ctx.bezierCurveTo(-w + c, h, -w, h - c, -w, h - r);
        ctx.lineTo(-w, -h + r);
        ctx.bezierCurveTo(-w, -h + c, -w + c, -h, -w + r, -h);
      } else {
        ctx.moveTo(-w, -h);
        ctx.lineTo(w, -h);
        ctx.lineTo(w, h);
        ctx.lineTo(-w, h);
      }
      ctx.closePath();
    }
    if (fill !== undefined) {
      ctx.fillStyle = fill;
      ctx.fill();
    }
    if (thickness > 0) {
      ctx.strokeStyle = color;
      ctx.lineWidth = Math.max(thickness, 1 / p.scale);
      setDash(ctx, dash);
      ctx.stroke();
      setDash(ctx, undefined);
    }
  } finally {
    ctx.restore();
    forget(p);
  }
}

function extent(p: Painter, font: LayoutFont): { ascent: number; descent: number } {
  if (p.options.metrics) {
    const metrics = p.options.metrics(font);
    if (metrics && finite(metrics.ascent) && finite(metrics.descent) && metrics.ascent + metrics.descent > 0) return metrics;
  }
  const size = finite(font.size) && font.size > 0 ? font.size : 12;
  return { ascent: size * 0.91, descent: size * 0.24 };
}

/** Underline patterns in units of the underline thickness; styles that are not here are solid. */
const UNDERLINE_DASHES: { [style: string]: number[] } = {
  dotted: [1.5, 1.5],
  dottedHeavy: [1.5, 1.5],
  dash: [5, 3],
  dashedHeavy: [5, 3],
  dashLong: [10, 4],
  dashLongHeavy: [10, 4],
  dotDash: [5, 2.5, 1.5, 2.5],
  dashDotHeavy: [5, 2.5, 1.5, 2.5],
  dotDotDash: [5, 2.5, 1.5, 2.5, 1.5, 2.5],
  dashDotDotHeavy: [5, 2.5, 1.5, 2.5, 1.5, 2.5],
};
const HEAVY_UNDERLINES = ['thick', 'dottedHeavy', 'dashedHeavy', 'dashLongHeavy', 'dashDotHeavy', 'dashDotDotHeavy', 'wavyHeavy'];

function drawUnderline(p: Painter, item: TextItem, size: number): void {
  const underline = item.underline;
  if (!underline || underline.style === 'none') return;
  const style = underline.style;
  const color = safeColor(underline.color, safeColor(item.color, '#000000'));
  let thickness = Math.max(size * 0.06, 0.5) * (item.font.bold ? 1.3 : 1);
  if (HEAVY_UNDERLINES.indexOf(style) >= 0) thickness *= 2;
  const y = item.y + size * 0.12 + thickness / 2;
  const x1 = item.x;
  const x2 = item.x + item.width;
  if (style === 'wave' || style === 'wavyHeavy') return wave(p, x1, y + thickness, x2, y + thickness, thickness, color);
  if (style === 'wavyDouble') {
    wave(p, x1, y, x2, y, thickness, color);
    return wave(p, x1, y + thickness * 3, x2, y + thickness * 3, thickness, color);
  }
  if (style === 'double') {
    const thin = thickness * 0.75;
    stroke(p, x1, y, x2, y, thin, color);
    // At least one clear canvas pixel between the two lines.
    return stroke(p, x1, y + Math.max(thin * 2, 2 / p.scale + thin), x2, y + Math.max(thin * 2, 2 / p.scale + thin), thin, color);
  }
  const pattern = Object.prototype.hasOwnProperty.call(UNDERLINE_DASHES, style) ? UNDERLINE_DASHES[style] : undefined;
  if (!pattern) return stroke(p, x1, y, x2, y, thickness, color);
  const unit = Math.max(thickness, 1 / p.scale);
  const dash: number[] = [];
  for (let i = 0; i < pattern.length; i++) dash.push(pattern[i] * unit);
  stroke(p, x1, y, x2, y, thickness, color, dash);
}

/** Length in UTF-16 units of the character at `index`. */
function charLength(text: string, index: number): number {
  const code = text.charCodeAt(index);
  if (code >= 0xd800 && code <= 0xdbff && index + 1 < text.length) {
    const next = text.charCodeAt(index + 1);
    if (next >= 0xdc00 && next <= 0xdfff) return 2;
  }
  return 1;
}

function drawText(p: Painter, item: TextItem): void {
  if (!finite(item.x) || !finite(item.y) || !item.font || typeof item.text !== 'string') return;
  const ctx = p.ctx;
  if (item.rotation === 90 || item.rotation === 270) {
    // Vertical text: the same drawing in a frame turned around the start of the baseline.
    ctx.save();
    ctx.translate(item.x, item.y);
    ctx.rotate(item.rotation === 90 ? Math.PI / 2 : -Math.PI / 2);
    const straight: TextItem = { ...item, x: 0, y: 0 };
    delete straight.rotation;
    try {
      drawText(p, straight);
    } finally {
      ctx.restore();
      forget(p);
    }
    return;
  }
  const size = finite(item.font.size) && item.font.size > 0 ? item.font.size : 12;
  const width = finite(item.width) && item.width > 0 ? item.width : 0;

  if (item.highlight !== undefined && width > 0) {
    const box = extent(p, item.font);
    setFill(p, safeColor(item.highlight, '#ffff00'));
    const top = snapPoint(item.y - box.ascent, p.scale);
    ctx.fillRect(item.x, top, width, snapPoint(item.y + box.descent, p.scale) - top);
  }

  const text = item.text;
  if (text.length > 0 && /[^\s]/.test(text)) {
    setFont(p, (p.options.cssFont || docCssFont)(item.font));
    setFill(p, safeColor(item.color, '#000000'));
    const spacing = finite(item.letterSpacing) ? item.letterSpacing : 0;
    const stretch = finite(item.scale) && item.scale > 0 && item.scale !== 1 ? item.scale : 1;
    if (spacing === 0 && stretch === 1) {
      ctx.fillText(text, item.x, item.y);
    } else if (spacing === 0 || hasLetterSpacing(ctx)) {
      ctx.save();
      ctx.translate(item.x, item.y);
      if (stretch !== 1) ctx.scale(stretch, 1);
      // Spacing is not stretched with the glyphs: the engine adds it after scaling.
      if (spacing !== 0) (ctx as Context & { letterSpacing: string }).letterSpacing = spacing / stretch + 'px';
      ctx.fillText(text, 0, 0);
      ctx.restore();
      forget(p);
    } else {
      // No letterSpacing on this canvas: one character after the other.
      let x = item.x;
      for (let i = 0; i < text.length; ) {
        const length = charLength(text, i);
        const char = length === 1 ? text.charAt(i) : text.substr(i, length);
        if (stretch === 1) ctx.fillText(char, x, item.y);
        else {
          ctx.save();
          ctx.translate(x, item.y);
          ctx.scale(stretch, 1);
          ctx.fillText(char, 0, 0);
          ctx.restore();
        }
        x += ctx.measureText(char).width * stretch + spacing * length;
        i += length;
      }
      if (stretch !== 1) forget(p);
    }
  }

  if (width > 0) {
    if (item.underline) drawUnderline(p, item, size);
    if (item.strike) {
      const color = safeColor(item.color, '#000000');
      const thickness = Math.max(size * 0.06, 0.5);
      const y = item.y - size * 0.28;
      if (item.strike === 'double') {
        const gap = Math.max(thickness, 1 / p.scale);
        stroke(p, item.x, y - gap, item.x + width, y - gap, thickness, color);
        stroke(p, item.x, y + gap, item.x + width, y + gap, thickness, color);
      } else stroke(p, item.x, y, item.x + width, y, thickness, color);
    }
  }
}

function drawPlaceholder(p: Painter, item: ImageItem): void {
  const ctx = p.ctx;
  const { x, y, width, height } = item;
  setFill(p, PLACEHOLDER_FILL);
  ctx.fillRect(x, y, width, height);
  const line = 1 / p.scale;
  setFill(p, PLACEHOLDER_LINE);
  ctx.fillRect(x, y, width, line);
  ctx.fillRect(x, y + height - line, width, line);
  ctx.fillRect(x, y, line, height);
  ctx.fillRect(x + width - line, y, line, height);
  const label = p.options.placeholderLabel ? p.options.placeholderLabel(item) : undefined;
  if (typeof label !== 'string' || label === '' || width < 40 || height < 16) return;
  setFont(p, PLACEHOLDER_FONT);
  setFill(p, PLACEHOLDER_TEXT);
  // Cut to what fits; the number of tries is bounded by the halving.
  let shown = label.length > 200 ? label.slice(0, 200) : label;
  const room = width - 8;
  while (shown.length > 1 && ctx.measureText(shown).width > room) shown = shown.slice(0, Math.max(1, shown.length >> 1));
  if (shown.length < label.length) shown = shown.length > 1 ? shown.slice(0, shown.length - 1) + '…' : '';
  if (shown === '' || ctx.measureText(shown).width > room) return;
  ctx.save();
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(shown, x + width / 2, y + height / 2);
  ctx.restore();
  forget(p);
}

const fraction = (value: unknown): number => (finite(value) && value > 0 ? Math.min(value, 1) : 0);

function drawImage(p: Painter, item: ImageItem): void {
  if (!finite(item.x) || !finite(item.y) || !finite(item.width) || !finite(item.height) || !(item.width > 0) || !(item.height > 0)) return;
  const ctx = p.ctx;
  const images = p.options.images;
  // The part of the picture that is shown, as shares of it; a crop that leaves nothing is not applied.
  const crop = item.crop;
  let cropLeft = 0;
  let cropTop = 0;
  let shownWidth = 1;
  let shownHeight = 1;
  if (crop) {
    const left = fraction(crop.left);
    const top = fraction(crop.top);
    const cw = 1 - left - fraction(crop.right);
    const ch = 1 - top - fraction(crop.bottom);
    if (cw > 0 && ch > 0) {
      cropLeft = left;
      cropTop = top;
      shownWidth = cw;
      shownHeight = ch;
    }
  }
  const loaded =
    item.part !== undefined && item.placeholder === undefined && images
      ? images.get(item.part, Math.ceil((item.width * p.scale) / shownWidth), Math.ceil((item.height * p.scale) / shownHeight))
      : undefined;
  const rotation = finite(item.rotation) ? item.rotation % 360 : 0;
  const transformed = rotation !== 0 || item.flipH === true || item.flipV === true;
  if (transformed) {
    ctx.save();
    ctx.translate(item.x + item.width / 2, item.y + item.height / 2);
    if (rotation !== 0) ctx.rotate((rotation * Math.PI) / 180);
    // A placeholder is turned with the picture but its words are not mirrored.
    if (loaded && (item.flipH || item.flipV)) ctx.scale(item.flipH ? -1 : 1, item.flipV ? -1 : 1);
    ctx.translate(-(item.x + item.width / 2), -(item.y + item.height / 2));
  }
  let drawn = false;
  if (loaded && loaded.width > 0 && loaded.height > 0) {
    const sx = cropLeft * loaded.width;
    const sy = cropTop * loaded.height;
    const sw = shownWidth * loaded.width;
    const sh = shownHeight * loaded.height;
    try {
      ctx.drawImage(loaded.source, sx, sy, sw, sh, item.x, item.y, item.width, item.height);
      drawn = true;
    } catch {
      // A picture the canvas refuses (broken, or closed in the meantime) is shown as a placeholder.
    }
  }
  if (!drawn) drawPlaceholder(p, item);
  if (transformed) {
    ctx.restore();
    forget(p);
  }
}

/**
 * Paints a page. The context is expected to map its origin to the top left
 * corner of the page; everything is drawn in px of the page times `scale`.
 * The state of the context is restored afterwards.
 */
export function renderPage(ctx: CanvasRenderingContext2D, page: LayoutPage, options: RenderPageOptions = {}): void {
  const scale = finite(options.scale) && options.scale > 0 ? options.scale : 1;
  const p: Painter = { ctx, scale, options, font: undefined, fill: undefined };
  ctx.save();
  try {
    if (scale !== 1) ctx.scale(scale, scale);
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
    if (hasLetterSpacing(ctx)) (ctx as Context & { letterSpacing: string }).letterSpacing = '0px';
    if (options.background !== null && finite(page.width) && finite(page.height)) {
      setFill(p, safeColor(options.background === undefined ? '#ffffff' : options.background, '#ffffff'));
      ctx.fillRect(0, 0, page.width, page.height);
    }
    const items = page.items;
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      if (item.type === 'text') drawText(p, item);
      else if (item.type === 'rect') drawRect(p, item);
      else if (item.type === 'line') drawLine(p, item);
      else if (item.type === 'image') drawImage(p, item);
      else if (item.type === 'shape') drawShape(p, item);
    }
  } finally {
    ctx.restore();
  }
}
