// The charts of a document: the chart parts its frames name
// (word/charts/chartN.xml), read with the reader of the charts of a workbook
// into the same model. A chart of a document is drawn from the values its
// part cached (<c:numCache>, <c:strCache>, literals): the workbook Word embeds
// beside the part (word/embeddings/*.xlsx) is never opened.
//
// Nothing here is ever serialized: the chart part, its relationships and the
// embedded workbook stay in the package and are written back from there. A
// part that cannot be read, that holds nothing that is drawn, or that lies
// beyond the limits of the document is said in the warnings of the document;
// its frame shows a placeholder.

import { MAX_WORKBOOK_CHARTS, MAX_WORKBOOK_CHART_POINTS, type Chart } from '../../model/sheet/chart';
import type { OpcPackage } from '../../opc';
import { parseXml } from '../../xml';
import { ChartBudgetError, parseChart } from '../xlsx/chart';
import { MAX_CHART_PART_BYTES, MAX_WORKBOOK_CHART_BYTES } from '../xlsx/drawings';
import { quote } from './common';

/** Chart parts read for one document, and cached values read for all of them together: the limits of a workbook. */
export const MAX_DOC_CHARTS = MAX_WORKBOOK_CHARTS;
export const MAX_DOC_CHART_POINTS = MAX_WORKBOOK_CHART_POINTS;
/** Bytes of chart parts read for one document, all charts together: the limit of a workbook. */
export const MAX_DOC_CHART_BYTES = MAX_WORKBOOK_CHART_BYTES;
/** Warnings about single charts of one document; what goes wrong beyond them is counted in one line. */
export const MAX_DOC_CHART_WARNINGS = 20;

const UNCACHED_NOTE = 'A series without cached values is drawn empty: the embedded workbook of a chart is not opened';

export interface DocChartLimits {
  /** Chart parts that are read. Default MAX_DOC_CHARTS. */
  maxCharts?: number;
  /** Cached values (<c:pt>) that are read, all charts together. Default MAX_DOC_CHART_POINTS. */
  maxChartPoints?: number;
  /** A chart part larger than this is not parsed. Default 32 MB. */
  maxChartPartBytes?: number;
  /** Bytes of chart parts that are read, all charts together. Default MAX_DOC_CHART_BYTES. */
  maxChartBytes?: number;
}

const limitOf = (value: number | undefined, fallback: number): number => (typeof value === 'number' && value >= 0 ? Math.min(value, fallback) : fallback);

/**
 * True for a chart with a series whose values have no cache in the part: no
 * value and no count of points. A cache that is there and holds only blanks
 * (a count, no point) is what the cells held, not a missing cache.
 */
function hasUncachedSeries(chart: Chart): boolean {
  for (const group of chart.groups) {
    for (const series of group.series) {
      const data = series.values;
      if (!data || !data.cache || !Array.isArray(data.cache.index)) return true;
      if (data.cache.index.length === 0 && !(data.count > 0)) return true;
    }
  }
  return false;
}

/**
 * Reads the chart parts (names of parts as the package spells them, each
 * once, in the order their frames stand in the document) into models, by
 * part name.
 * `themeColors` as parseChart() takes them. Undefined when no part gave a
 * chart. Never throws: what goes wrong is added to `warnings`.
 */
export async function readDocCharts(pkg: OpcPackage, parts: Iterable<string>, themeColors: string[], warnings: string[], limits: DocChartLimits = {}): Promise<Map<string, Chart> | undefined> {
  const maxCharts = limitOf(limits.maxCharts, MAX_DOC_CHARTS);
  const maxBytes = limitOf(limits.maxChartPartBytes, MAX_CHART_PART_BYTES);
  const budget = { points: limitOf(limits.maxChartPoints, MAX_DOC_CHART_POINTS) };
  let bytesLeft = limitOf(limits.maxChartBytes, MAX_DOC_CHART_BYTES);
  const charts = new Map<string, Chart>();
  let read = 0;
  let beyond = 0;
  let said = 0;
  let unsaid = 0;
  const warn = (text: string) => {
    if (said < MAX_DOC_CHART_WARNINGS) {
      warnings.push(text);
      said++;
    } else unsaid++;
  };
  for (const part of parts) {
    if (read >= maxCharts) {
      beyond++;
      continue;
    }
    read++;
    try {
      const bytes = await pkg.readPart(part);
      if (bytes.length > maxBytes) throw new Error('the part is too large');
      if (bytes.length > bytesLeft) {
        warn(`Chart ${quote(part)} is not shown: the charts of the document hold too many bytes`);
        continue;
      }
      bytesLeft -= bytes.length;
      const chart = parseChart(parseXml(bytes).root, themeColors, budget);
      if (!chart) {
        warn(`Chart ${quote(part)} holds nothing that is drawn (no series, or a chart type that is not drawn); a placeholder is shown`);
        continue;
      }
      if (hasUncachedSeries(chart) && chart.notes.indexOf(UNCACHED_NOTE) < 0) chart.notes.push(UNCACHED_NOTE);
      charts.set(part, chart);
    } catch (error) {
      // A chart part that cannot be read (not XML, a DTD, too much of it) leaves its frame a placeholder.
      if (error instanceof ChartBudgetError) warn(`Chart ${quote(part)} is not shown: the charts of the document hold too many cached values`);
      else warn(`Chart ${quote(part)} could not be read: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
  if (unsaid > 0) warnings.push(`More charts that are not shown: ${unsaid}`);
  if (beyond > 0) warnings.push(`Charts beyond the limit of a document (${maxCharts}) are not shown: ${beyond}`);
  return charts.size > 0 ? charts : undefined;
}
