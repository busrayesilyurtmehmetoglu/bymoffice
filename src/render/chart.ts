// Paints a chart on a canvas: the items a layout lists (see layout/chart.ts),
// in the frame the chart has on its sheet. Also keeps what was worked out for
// a chart until the cells or the fonts change, per size of its frame, so that
// scrolling past a chart of many points does not read its cells every frame;
// what is kept for all charts together is limited.
//
// Nothing that comes from a file reaches `ctx.font`, `fillStyle` or
// `strokeStyle` unchecked, and text only ever reaches the canvas through
// fillText().

import type { FormatLocale } from '../format/locale';
import { chartLayoutWeight as layoutWeight, chartValues, layoutChart, type ChartFont, type ChartItem, type ChartLayout, type ChartValues } from '../layout/chart';
import type { Chart } from '../model/sheet/chart';
import type { Font, Workbook } from '../model/sheet/types';
import { fontsVersion } from './fonts';
import { cssFont, TextPainter } from './text';

const HEX_COLOR = /^#[0-9A-Fa-f]{6}$/;
const safeColor = (color: string | undefined): string | undefined => (typeof color === 'string' && HEX_COLOR.test(color) ? color : undefined);

/** The value for `ctx.font` of a font of a chart. */
export function chartCssFont(font: ChartFont): string {
  const model: Font = { name: font.family, size: (font.size * 72) / 96, bold: font.bold === true, italic: font.italic === true, strike: false, underline: 'none' };
  return cssFont(model);
}

function roundedRect(ctx: CanvasRenderingContext2D, x: number, y: number, width: number, height: number, radius: number): void {
  const r = Math.max(Math.min(radius, width / 2, height / 2), 0);
  ctx.beginPath();
  if (!(r > 0)) {
    ctx.rect(x, y, width, height);
    return;
  }
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + width - r, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + r);
  ctx.lineTo(x + width, y + height - r);
  ctx.quadraticCurveTo(x + width, y + height, x + width - r, y + height);
  ctx.lineTo(x + r, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

function setStroke(ctx: CanvasRenderingContext2D, color: string, width: number | undefined, dash: boolean | undefined): void {
  const lineWidth = typeof width === 'number' && width > 0 && width < 64 ? width : 1;
  ctx.strokeStyle = color;
  ctx.lineWidth = lineWidth;
  ctx.setLineDash(dash ? [lineWidth * 4, lineWidth * 3] : []);
}

/** The outline of a marker around (x, y), added to the current path. True for a shape that is a line and has no inside. */
function markerPath(ctx: CanvasRenderingContext2D, symbol: string, x: number, y: number, size: number): boolean {
  const half = size / 2;
  switch (symbol) {
    case 'square':
      ctx.rect(x - half, y - half, size, size);
      return false;
    case 'diamond':
      ctx.moveTo(x, y - half);
      ctx.lineTo(x + half, y);
      ctx.lineTo(x, y + half);
      ctx.lineTo(x - half, y);
      ctx.closePath();
      return false;
    case 'triangle':
      ctx.moveTo(x, y - half);
      ctx.lineTo(x + half, y + half);
      ctx.lineTo(x - half, y + half);
      ctx.closePath();
      return false;
    case 'x':
      ctx.moveTo(x - half, y - half);
      ctx.lineTo(x + half, y + half);
      ctx.moveTo(x + half, y - half);
      ctx.lineTo(x - half, y + half);
      return true;
    case 'star':
      ctx.moveTo(x - half, y - half);
      ctx.lineTo(x + half, y + half);
      ctx.moveTo(x + half, y - half);
      ctx.lineTo(x - half, y + half);
      ctx.moveTo(x, y - half);
      ctx.lineTo(x, y + half);
      return true;
    case 'plus':
      ctx.moveTo(x - half, y);
      ctx.lineTo(x + half, y);
      ctx.moveTo(x, y - half);
      ctx.lineTo(x, y + half);
      return true;
    case 'dash':
      ctx.moveTo(x - half, y);
      ctx.lineTo(x + half, y);
      return true;
    case 'dot':
      ctx.moveTo(x - half / 2, y);
      ctx.lineTo(x + half / 2, y);
      return true;
    default:
      // "circle" and whatever else: a round dot.
      ctx.moveTo(x + half, y);
      ctx.arc(x, y, half, 0, Math.PI * 2);
      return false;
  }
}

function paintItem(ctx: CanvasRenderingContext2D, painter: TextPainter, item: ChartItem, css: (font: ChartFont) => string): void {
  if (item.kind === 'rect') {
    const fill = safeColor(item.fill);
    const stroke = safeColor(item.stroke);
    if (!(item.width >= 0 && item.height >= 0)) return;
    if (fill) {
      painter.setFill(fill);
      if (item.radius) {
        roundedRect(ctx, item.x, item.y, item.width, item.height, item.radius);
        ctx.fill();
      } else ctx.fillRect(item.x, item.y, item.width, item.height);
    }
    if (stroke && item.width > 1 && item.height > 1) {
      const width = item.lineWidth || 1;
      setStroke(ctx, stroke, width, false);
      // The outline lies inside the box, so that the frame of a chart is not cut by its own clip.
      const inset = Math.min(width / 2, item.width / 2, item.height / 2);
      roundedRect(ctx, item.x + inset, item.y + inset, item.width - inset * 2, item.height - inset * 2, item.radius ? Math.max(item.radius - inset, 0) : 0);
      ctx.stroke();
    }
    return;
  }
  if (item.kind === 'path') {
    const points = item.points;
    if (points.length < 4) return;
    const fill = item.closed ? safeColor(item.fill) : undefined;
    const stroke = safeColor(item.stroke);
    if (!fill && !stroke) return;
    ctx.beginPath();
    if (!item.segments) ctx.moveTo(points[0], points[1]);
    if (item.curve) {
      // Pieces of a curve: two control points and an end point each.
      for (let i = 2; i + 5 < points.length; i += 6) ctx.bezierCurveTo(points[i], points[i + 1], points[i + 2], points[i + 3], points[i + 4], points[i + 5]);
    } else if (item.segments) {
      // Lines of their own: from every first point to every second.
      for (let i = 0; i + 3 < points.length; i += 4) {
        ctx.moveTo(points[i], points[i + 1]);
        ctx.lineTo(points[i + 2], points[i + 3]);
      }
    } else for (let i = 2; i + 1 < points.length; i += 2) ctx.lineTo(points[i], points[i + 1]);
    if (item.closed) ctx.closePath();
    if (fill) {
      painter.setFill(fill);
      ctx.fill();
    }
    if (stroke) {
      setStroke(ctx, stroke, item.lineWidth, item.dash);
      ctx.lineJoin = 'round';
      ctx.stroke();
      ctx.setLineDash([]);
    }
    return;
  }
  if (item.kind === 'markers') {
    const fill = safeColor(item.fill);
    const stroke = safeColor(item.stroke);
    const size = item.size > 0 && item.size < 200 ? item.size : 5;
    const points = item.points;
    ctx.beginPath();
    let open = false;
    for (let i = 0; i + 1 < points.length; i += 2) open = markerPath(ctx, item.symbol, points[i], points[i + 1], size);
    if (open) {
      // Crosses and dashes have no inside: they are drawn in the colour of the marker.
      const color = stroke || fill;
      if (!color) return;
      setStroke(ctx, color, Math.max(size / 5, 1), false);
      ctx.stroke();
      return;
    }
    if (fill) {
      painter.setFill(fill);
      ctx.fill();
    }
    if (stroke) {
      setStroke(ctx, stroke, 1, false);
      ctx.stroke();
    }
    return;
  }
  if (item.kind === 'circles') {
    const fill = safeColor(item.fill);
    const stroke = safeColor(item.stroke);
    const points = item.points;
    if (!fill && !stroke) return;
    ctx.beginPath();
    for (let i = 0; i + 2 < points.length; i += 3) {
      const radius = points[i + 2];
      if (!(radius > 0)) continue;
      ctx.moveTo(points[i] + radius, points[i + 1]);
      ctx.arc(points[i], points[i + 1], radius, 0, Math.PI * 2);
    }
    if (fill) {
      painter.setFill(fill);
      ctx.fill();
    }
    if (stroke) {
      setStroke(ctx, stroke, item.lineWidth, false);
      ctx.stroke();
    }
    return;
  }
  if (item.kind === 'slice') {
    const fill = safeColor(item.fill);
    if (!fill || !(item.radius > 0) || !(item.end > item.start)) return;
    ctx.beginPath();
    if (item.inner > 0) {
      ctx.arc(item.cx, item.cy, item.radius, item.start, item.end);
      ctx.arc(item.cx, item.cy, item.inner, item.end, item.start, true);
    } else {
      ctx.moveTo(item.cx, item.cy);
      ctx.arc(item.cx, item.cy, item.radius, item.start, item.end);
    }
    ctx.closePath();
    painter.setFill(fill);
    ctx.fill();
    const stroke = safeColor(item.stroke);
    if (stroke) {
      setStroke(ctx, stroke, item.lineWidth, false);
      ctx.lineJoin = 'round';
      ctx.stroke();
    }
    return;
  }
  // Text: only ever through fillText().
  if (typeof item.text !== 'string' || item.text === '') return;
  painter.setFont(css(item.font));
  painter.setFill(safeColor(item.color) || '#000000');
  ctx.textAlign = item.align;
  ctx.textBaseline = 'alphabetic';
  const angle = item.rotated ? -90 : typeof item.angle === 'number' && Number.isFinite(item.angle) ? item.angle : 0;
  if (angle !== 0) {
    // Turned around its own place: read from the bottom up for a turn against the clock.
    ctx.save();
    ctx.translate(item.x, item.y);
    ctx.rotate((angle * Math.PI) / 180);
    ctx.fillText(item.text, 0, 0);
    ctx.restore();
    painter.invalidate();
  } else ctx.fillText(item.text, item.x, item.y);
}

/**
 * Paints a layout with its top-left corner at (x, y), clipped to the frame of
 * the chart. The state of the context is as it was afterwards, also when
 * painting fails (the error is passed on). `css` names the value for
 * `ctx.font` of a font of the chart, for a host that measured the layout with
 * fonts of its own; default chartCssFont().
 */
export function paintChartLayout(ctx: CanvasRenderingContext2D, layout: ChartLayout, x: number, y: number, css: (font: ChartFont) => string = chartCssFont): void {
  if (!(layout.width > 0 && layout.height > 0)) return;
  const painter = new TextPainter(ctx);
  ctx.save();
  try {
    ctx.translate(x, y);
    roundedRect(ctx, 0, 0, layout.width, layout.height, layout.radius);
    ctx.clip();
    ctx.globalAlpha = 1;
    ctx.lineCap = 'butt';
    ctx.lineJoin = 'miter';
    ctx.setLineDash([]);
    const items = layout.items;
    for (let i = 0; i < items.length; i++) paintItem(ctx, painter, items[i], css);
  } finally {
    ctx.restore();
  }
}

export interface ChartRendererOptions {
  /** The workbook the series read their cells from; without it charts are drawn from the values their files cached. */
  workbook?: Workbook;
  /** Asked at every frame: number formats of axes and labels follow it. */
  locale?(): FormatLocale | undefined;
  /**
   * A number that changes whenever cells, hidden rows and columns, sheet names or the locale
   * may have changed. What was worked out for a chart is kept while it stays the same;
   * without it nothing is kept from one frame to the next.
   */
  version?(): number;
  /** The name of a series that has none. */
  seriesName?(n: number): string;
  /**
   * The value for `ctx.font` of a font of a chart, for measuring and for painting alike:
   * for a host that names its fonts another way than the grid does (a document, with the
   * font list of its own measurer). It must be safe for a font name that comes from a
   * file. Default chartCssFont().
   */
  cssFont?(font: ChartFont): string;
}

/** Numbers kept for the charts of one renderer, all charts together: values of points, cells of labels and the coordinates of what is drawn. */
export const CHART_KEEP_BUDGET = 2000000;
/** Sizes of the frame of one chart whose layouts are kept side by side. */
const MAX_KEPT_SIZES = 4;

interface Kept {
  version: number;
  fonts: number;
  values: ChartValues;
  /** The layouts by the size of the frame ("width x height"), the one used last at the end. */
  layouts: Map<string, ChartLayout>;
  /** What the entry weighs: see CHART_KEEP_BUDGET. */
  weight: number;
}

/** What a renderer did and holds: for tests, and for whoever wants to see what charts cost. */
export interface ChartRendererState {
  /** Charts something is kept for. */
  charts: number;
  /** What is kept, counted in numbers; see CHART_KEEP_BUDGET. */
  weight: number;
  /** How often the numbers of a chart were worked out from cells or from the cache. */
  values: number;
  /** How often a chart was laid out. */
  layouts: number;
}

/**
 * Draws charts, and remembers per chart what it worked out for the cells and for the sizes of
 * the frames it was last drawn in. What is remembered is limited for all charts together: the
 * charts that were drawn longest ago are forgotten first.
 */
export class ChartRenderer {
  private readonly options: ChartRendererOptions;
  /** In the order of use: the chart drawn last is the last one. */
  private kept = new Map<Chart, Kept>();
  private weight = 0;
  private valuesMade = 0;
  private layoutsMade = 0;
  private readonly budget: number;

  /** `budget`: the most that is kept, counted in numbers. Default CHART_KEEP_BUDGET. */
  constructor(options: ChartRendererOptions = {}, budget = CHART_KEEP_BUDGET) {
    this.options = options;
    this.budget = budget >= 0 ? budget : CHART_KEEP_BUDGET;
  }

  /** The layout of a chart in a frame of the given size, with the cells as they are now. */
  layout(ctx: CanvasRenderingContext2D, chart: Chart, width: number, height: number): ChartLayout {
    const options = this.options;
    const version = options.version ? options.version() : NaN;
    const fonts = fontsVersion();
    const locale = options.locale ? options.locale() : undefined;
    let kept = this.kept.get(chart);
    // NaN never equals itself: without a version everything is worked out anew.
    if (!kept || kept.version !== version || kept.fonts !== fonts) {
      if (kept) this.weight -= kept.weight;
      const values = chartValues(chart, { workbook: options.workbook, locale, seriesName: options.seriesName });
      this.valuesMade++;
      kept = { version, fonts, values, layouts: new Map(), weight: values.size + values.series.length * 8 };
      this.weight += kept.weight;
    }
    // Used last: moved to the end.
    this.kept.delete(chart);
    this.kept.set(chart, kept);
    // The same chart may stand in frames of several sizes (on its sheet and on paper, in two anchors): each keeps its own.
    const key = `${width}x${height}`;
    let layout = kept.layouts.get(key);
    if (layout) kept.layouts.delete(key);
    else {
      const painter = new TextPainter(ctx);
      const css = new Map<ChartFont, string>();
      const measure = (text: string, font: ChartFont): number => {
        let value = css.get(font);
        if (value === undefined) {
          value = options.cssFont ? options.cssFont(font) : chartCssFont(font);
          css.set(font, value);
        }
        return painter.measure(value, text);
      };
      layout = layoutChart(chart, kept.values, width, height, measure, { locale, date1904: options.workbook ? options.workbook.date1904 === true : chart.date1904 === true });
      this.layoutsMade++;
      const weight = layoutWeight(layout);
      kept.weight += weight;
      this.weight += weight;
      if (kept.layouts.size >= MAX_KEPT_SIZES) {
        const oldest = kept.layouts.keys().next().value as string;
        const dropped = layoutWeight(kept.layouts.get(oldest) as ChartLayout);
        kept.layouts.delete(oldest);
        kept.weight -= dropped;
        this.weight -= dropped;
      }
    }
    kept.layouts.set(key, layout);
    this.trim(chart);
    return layout;
  }

  /** Forgets the charts used longest ago until what is kept fits the budget; never the one in use. */
  private trim(current: Chart): void {
    if (this.weight <= this.budget) return;
    const charts = Array.from(this.kept.keys());
    for (let i = 0; i < charts.length && this.weight > this.budget; i++) {
      if (charts[i] === current) continue;
      this.weight -= (this.kept.get(charts[i]) as Kept).weight;
      this.kept.delete(charts[i]);
    }
  }

  /** Forgets everything that was worked out: for charts that are no longer in sight. */
  clear(): void {
    this.kept.clear();
    this.weight = 0;
  }

  /** What the renderer holds and how often it worked something out. */
  state(): ChartRendererState {
    return { charts: this.kept.size, weight: this.weight, values: this.valuesMade, layouts: this.layoutsMade };
  }

  /** The layouts that are kept, the one used last at the end: what was last drawn, for tests. */
  layouts(): ChartLayout[] {
    const out: ChartLayout[] = [];
    this.kept.forEach((kept) => kept.layouts.forEach((layout) => out.push(layout)));
    return out;
  }

  /** Paints a chart into the box (x, y, width, height). Throws when the chart cannot be laid out or painted; the context is then as it was. */
  paint(ctx: CanvasRenderingContext2D, chart: Chart, x: number, y: number, width: number, height: number): void {
    paintChartLayout(ctx, this.layout(ctx, chart, width, height), x, y, this.options.cssFont);
  }
}
