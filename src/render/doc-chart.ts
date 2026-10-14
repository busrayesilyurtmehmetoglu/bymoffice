// Draws the charts of a document on the canvas of a page: the charts the
// reader took from the chart parts (DocModel.charts), laid out and painted by
// the chart engine the grid uses, from the values the parts cached.
//
// The words of a chart are measured and painted with the font list of the
// document's measurer (the same substitutes, the fonts the application
// registered, the font of the package), so that a chart is set in the fonts
// the text around it is set in. A page is painted in px of the page under the
// scale of the view or of the printer: the layout of a chart depends on the
// size of its frame alone, so it is worked out once and painted at every zoom
// and on paper from what was kept.

import type { FormatLocale } from '../format/locale';
import type { ChartFont, ChartWord } from '../layout/chart';
import type { LayoutFont } from '../layout/doc/types';
import type { Chart } from '../model/sheet/chart';
import { ChartRenderer, type ChartRendererState } from './chart';
import { fontsVersion } from './fonts';

export interface DocChartPainterOptions {
  /** The CSS font of a layout font: that of the measurer the document is laid out with. */
  cssFont(font: LayoutFont): string;
  /** The font of words of a chart that name none: the body font of the theme of the document. Default "Calibri". */
  family?: string;
  /** Asked whenever a chart is laid out: number formats of axes and labels follow it. */
  locale?(): FormatLocale | undefined;
  /** The name of a series that has none. */
  seriesName?(n: number): string;
  /** The words a chart shows of its own (the legend of a waterfall). */
  chartWord?(word: ChartWord): string;
}

/** Paints the charts of the pages of one document; see RenderPageOptions.charts. */
export class DocChartPainter {
  private readonly renderer: ChartRenderer;
  /** The font strings the host gave, by face. */
  private readonly css = new Map<string, string>();

  constructor(options: DocChartPainterOptions) {
    const family = typeof options.family === 'string' && options.family !== '' ? options.family : 'Calibri';
    // The font strings by face; the list of a family changes when a font is registered or arrives.
    let fonts = fontsVersion();
    this.renderer = new ChartRenderer({
      locale: options.locale,
      seriesName: options.seriesName,
      chartWord: options.chartWord,
      // The cached values of a part never change: what was worked out is kept until the fonts change.
      version: () => 0,
      cssFont: (font: ChartFont): string => {
        const named = typeof font.family === 'string' && font.family !== '' ? font.family : family;
        const key = `${font.italic ? 'i' : ''}${font.bold ? 'b' : ''}${font.size}/${named}`;
        if (fonts !== fontsVersion() || this.css.size > 256) {
          this.css.clear();
          fonts = fontsVersion();
        }
        let value = this.css.get(key);
        if (value === undefined) {
          value = options.cssFont({ family: named, size: font.size, bold: font.bold === true, italic: font.italic === true });
          this.css.set(key, value);
        }
        return value;
      },
    });
  }

  /** Paints a chart into the box, in the units of the context. Throws when the chart cannot be laid out or painted; the context is then as it was. */
  paint(ctx: CanvasRenderingContext2D, chart: Chart, x: number, y: number, width: number, height: number): void {
    this.renderer.paint(ctx, chart, x, y, width, height);
  }

  /**
   * Forgets what was worked out for the charts, and the font strings of the host: for a
   * document that is no longer shown, and whenever the host may measure differently (it
   * lays its pages out anew, its font table changed, its language changed).
   */
  clear(): void {
    this.css.clear();
    this.renderer.clear();
  }

  /** What is kept and how often something was worked out: for tests. */
  state(): ChartRendererState {
    return this.renderer.state();
  }
}
