// Charts made up at random, many seeds: every chart type with every option
// the layout knows of (data labels, a second scale, sides, turned labels,
// levels of categories, logarithmic and date axes, trend lines, error bars,
// lines of chart types, places set by hand, slices pulled out), in frames of
// every size, with numbers at the edge of what a number can be. Whatever
// comes in, the layout throws nothing, holds only finite numbers, keeps what
// it places inside the frame and stays within its budgets.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_AXIS_TICKS, MAX_DATA_LABELS, MAX_DRAWN_ITEMS, MAX_DRAWN_TRENDS, MAX_LAYOUT_WEIGHT, chartLayoutWeight, chartValues, layoutChart, type ChartItem, type ChartLayout, type ChartScale } from '../src/layout/chart';
import { chartXml, fakeMeasure, legend, numLit, readChart, ser, strLit, title, xySer } from './chart-fixtures';

function assertScale(scale: ChartScale, message: string): void {
  assert.ok(Number.isFinite(scale.min) && Number.isFinite(scale.max) && scale.max > scale.min, `${message}: ${scale.min}..${scale.max}`);
  assert.ok(Number.isFinite(scale.step) && scale.step > 0, `${message}: step ${scale.step}`);
  assert.ok(scale.ticks.length >= 1 && scale.ticks.length <= MAX_AXIS_TICKS + 1, `${message}: ${scale.ticks.length} ticks`);
  for (const tick of scale.ticks) assert.ok(Number.isFinite(tick), `${message}: tick ${tick}`);
  if (scale.log !== undefined) assert.ok(scale.min > 0 && scale.log >= 2 && scale.log <= 1000, `${message}: a logarithmic scale from ${scale.min} to base ${scale.log}`);
}

/** Every number a layout holds is finite, and what it places lies in its frame. */
function assertSound(layout: ChartLayout, message: string): void {
  const finite = (value: unknown, where: string): void => {
    if (typeof value === 'number') assert.ok(Number.isFinite(value), `${message}: ${where} is ${value}`);
    else if (Array.isArray(value)) for (let i = 0; i < value.length; i++) finite(value[i], `${where}[${i}]`);
    else if (value && typeof value === 'object') for (const key of Object.keys(value)) finite((value as Record<string, unknown>)[key], `${where}.${key}`);
  };
  finite([layout.width, layout.height, layout.radius], 'the frame');
  finite(layout.plot, 'plot');
  for (const scale of [layout.scale, layout.xScale, layout.scale2]) if (scale) assertScale(scale, message);
  const w = layout.width;
  const h = layout.height;
  // Half a pixel for lines that are put on the middle of a pixel.
  const slack = 0.5001;
  const inside = (x: number, y: number, where: string): void => assert.ok(x >= -slack && x <= w + slack && y >= -slack && y <= h + slack, `${message}: ${where} at ${x}, ${y} lies outside ${w} x ${h}`);
  let numbers = 0;
  let labels = 0;
  let trends = 0;
  layout.items.forEach((item: ChartItem, i) => {
    const where = `item ${i} (${item.kind}/${(item as { role?: string }).role || ''})`;
    finite(item, where);
    if (item.kind === 'rect') {
      inside(item.x, item.y, where);
      inside(item.x + item.width, item.y + item.height, `${where}, far corner`);
      assert.ok(item.width >= 0 && item.height >= 0, `${message}: ${where} has a size below nothing`);
    } else if (item.kind === 'path' || item.kind === 'markers') {
      for (let k = 0; k + 1 < item.points.length; k += 2) inside(item.points[k], item.points[k + 1], `${where}, point ${k / 2}`);
      numbers += item.points.length;
      if (item.kind === 'path' && item.role === 'trend') trends++;
    } else if (item.kind === 'circles') {
      for (let k = 0; k + 2 < item.points.length; k += 3) {
        inside(item.points[k], item.points[k + 1], `${where}, bubble ${k / 3}`);
        assert.ok(item.points[k + 2] >= 0 && item.points[k + 2] <= Math.max(w, h), `${message}: ${where} has a radius of ${item.points[k + 2]}`);
      }
      numbers += item.points.length;
    } else if (item.kind === 'slice') {
      inside(item.cx, item.cy, where);
      assert.ok(item.radius >= 0 && item.radius <= Math.max(w, h) && item.inner >= 0 && item.inner <= item.radius + 1e-9, `${message}: ${where} has radii ${item.inner}..${item.radius}`);
    } else {
      inside(item.x, item.y, `${where} "${item.text}"`);
      assert.ok(item.text.length <= 301, `${message}: ${where} holds ${item.text.length} characters`);
      assert.ok(item.angle === undefined || (item.angle >= -90 && item.angle <= 90), `${message}: ${where} is turned by ${item.angle}`);
      if (item.role === 'label') labels++;
    }
  });
  assert.ok(labels <= MAX_DATA_LABELS, `${message}: ${labels} data labels`);
  assert.ok(trends <= MAX_DRAWN_TRENDS * 64, `${message}: ${trends} pieces of trend lines`);
  // What a frame is worth painting: a bounded number of coordinates, whatever the chart holds.
  assert.ok(layout.items.length <= MAX_DRAWN_ITEMS * 2 && numbers <= MAX_DRAWN_ITEMS * 40, `${message}: ${layout.items.length} items, ${numbers} numbers`);
  assert.ok(chartLayoutWeight(layout) <= MAX_LAYOUT_WEIGHT, `${message}: the layout weighs ${chartLayoutWeight(layout)}`);
}

test('layout: every chart type with every option, numbers of every size, frames of every size (many seeds)', () => {
  const special: (() => (number | null)[])[] = [
    () => [0, 0, 0],
    () => [5],
    () => [-3, -1, -7],
    () => [1e308, 1e308],
    () => [1.7e308, -1.7e308],
    () => [1e-308, 2e-308],
    () => [null, null],
    () => [],
    () => [null, 3, null, 4, null],
    () => [1e300, -1e-300, 0],
    () => [5e-324, 0],
    () => [100, 100, 100],
    () => [99.5, 100, 99.9],
    () => [45292, 45293, 45300, 45658],
    () => [2958465, 1, 0, 2958466],
  ];
  const sizes: [number, number][] = [[480, 300], [0, 0], [-5, 100], [1, 1], [79, 59], [100, 100], [181, 121], [3000, 20], [20, 3000], [1e9, 1e9], [NaN, 100], [200.5, 133.3], [640, 400], [900, 250]];
  const kinds = ['bar', 'barH', 'line', 'area', 'pie', 'doughnut', 'ofPie', 'scatter', 'bubble', 'radar', 'stock', 'combo', 'second'];
  let checked = 0;
  const seen = new Set<string>();
  for (let seed = 1; seed <= 10; seed++) {
    let state = seed * 104729;
    const random = () => (state = (state * 1103515245 + 12345) & 0x7fffffff) / 0x7fffffff;
    const pick = <T>(list: T[]): T => list[Math.floor(random() * list.length) % list.length];
    const maybe = (text: string, chance = 0.3): string => (random() < chance ? text : '');
    const number = (): string => String(pick([0, 1, -1, 0.5, 5, 50, 1e-9, 1e9, 1e308, -1e308, 5e-324, 7.25]));
    const data = (): (number | null)[] => {
      if (random() < 0.35) return pick(special)();
      const scale = Math.pow(10, Math.floor(random() * 620 - 310));
      const sign = random();
      return Array.from({ length: Math.floor(random() * 9) }, () => (random() < 0.1 ? null : (random() - (sign < 0.3 ? 0.5 : sign < 0.4 ? 1 : 0)) * scale));
    };
    const flag = (name: string): string => `<c:${name} val="${pick([0, 1])}"/>`;
    const labels = (): string =>
      maybe(
        '<c:dLbls>' +
          maybe(`<c:dLbl><c:idx val="${pick([0, 1, 2, 99999])}"/>${pick(['<c:delete val="1"/>', '<c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Özel etiket</a:t></a:r></a:p></c:rich></c:tx><c:showVal val="1"/>', '<c:dLblPos val="ctr"/><c:showCatName val="1"/>'])}</c:dLbl>`) +
          maybe(`<c:numFmt formatCode="${pick(['0.0', '0%', 'General', '#,##0', 'd.m.yyyy', '[Red]0;;'])}" sourceLinked="${pick([0, 1])}"/>`) +
          maybe(`<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz="${pick([100, 900, 4000, 400000])}" b="1"/></a:pPr></a:p></c:txPr>`) +
          maybe(`<c:dLblPos val="${pick(['outEnd', 'inEnd', 'ctr', 'inBase', 'bestFit', 't', 'b', 'l', 'r'])}"/>`, 0.6) +
          ['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize'].map(flag).join('') +
          maybe(`<c:separator>${pick([' ', '; ', '\n', 'çok uzun bir ayraç'])}</c:separator>`) +
          '</c:dLbls>',
        0.45,
      );
    const trend = (): string =>
      maybe(
        `<c:trendline><c:trendlineType val="${pick(['linear', 'exp', 'log', 'poly', 'power', 'movingAvg'])}"/><c:order val="${pick([2, 3, 6])}"/><c:period val="${pick([2, 3, 50])}"/>${maybe(`<c:forward val="${number()}"/>`)}${maybe(`<c:backward val="${number()}"/>`)}</c:trendline>`,
        0.3,
      ).repeat(pick([1, 1, 3]));
    const errors = (xy: boolean): string =>
      maybe(
        `<c:errBars><c:errDir val="${xy ? pick(['x', 'y']) : 'y'}"/><c:errBarType val="${pick(['both', 'plus', 'minus'])}"/><c:errValType val="${pick(['fixedVal', 'percentage', 'stdDev', 'stdErr', 'cust'])}"/>${flag('noEndCap')}` +
          `<c:plus>${numLit(data())}</c:plus><c:minus>${numLit(data())}</c:minus><c:val val="${number()}"/></c:errBars>`,
        0.3,
      );
    const extras = (xy: boolean): string => labels() + trend() + errors(xy) + maybe(`<c:explosion val="${pick([0, 25, 400, 100000])}"/>`, 0.2);
    const manual = (): string =>
      maybe(
        `<c:layout><c:manualLayout>${maybe(`<c:layoutTarget val="${pick(['inner', 'outer'])}"/>`, 0.7)}${maybe('<c:xMode val="edge"/>', 0.8)}${maybe('<c:yMode val="edge"/>', 0.8)}` +
          `<c:x val="${pick([0, 0.1, 0.5, 0.95, 1, -1, 2, 1e308])}"/><c:y val="${pick([0, 0.1, 0.5, 0.95, 1, -0.5, 7])}"/>${maybe(`<c:w val="${pick([0.001, 0.3, 0.8, 1, 5, -1])}"/><c:h val="${pick([0.001, 0.3, 0.8, 1, 5, -1])}"/>`, 0.8)}</c:manualLayout></c:layout>`,
        0.3,
      ) || '<c:layout/>';
    const scaling = (): string =>
      `<c:scaling>${maybe(`<c:logBase val="${pick([2, 10, 1000, 2.5])}"/>`, 0.25)}<c:orientation val="${pick(['minMax', 'maxMin'])}"/>${maybe(`<c:max val="${number()}"/>`, 0.2)}${maybe(`<c:min val="${number()}"/>`, 0.2)}</c:scaling>`;
    const common = (): string =>
      `<c:delete val="${pick([0, 0, 0, 1])}"/><c:axPos val="${pick(['b', 'l', 't', 'r'])}"/>${maybe('<c:majorGridlines/>', 0.6)}${maybe('<c:minorGridlines/>', 0.3)}` +
      maybe('<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Eksen başlığı</a:t></a:r></a:p></c:rich></c:tx></c:title>') +
      maybe(`<c:numFmt formatCode="${pick(['0.0', 'd.m.yyyy', 'mmm', 'General'])}" sourceLinked="0"/>`) +
      `<c:tickLblPos val="${pick(['nextTo', 'nextTo', 'high', 'low', 'none'])}"/>` +
      maybe(`<c:txPr><a:bodyPr rot="${pick([-5400000, -2700000, 0, 2700000, 5400000, 600000, 99999999])}"/><a:p><a:pPr><a:defRPr/></a:pPr></a:p></c:txPr>`, 0.25);
    const crossing = (): string => maybe(pick(['<c:crosses val="max"/>', '<c:crosses val="min"/>', '<c:crosses val="autoZero"/>', `<c:crossesAt val="${number()}"/>`]), 0.4);
    const valueAxis = (id: number, cross: number): string =>
      `<c:valAx><c:axId val="${id}"/>${scaling()}${common()}<c:crossAx val="${cross}"/>${crossing()}${maybe('<c:crossBetween val="midCat"/>')}${maybe(`<c:majorUnit val="${number()}"/>`, 0.2)}</c:valAx>`;
    const categoryAxis = (id: number, cross: number): string => {
      const name = pick(['catAx', 'catAx', 'dateAx']);
      const time = name === 'dateAx' ? maybe(`<c:baseTimeUnit val="${pick(['days', 'months', 'years'])}"/>`, 0.6) + maybe(`<c:majorUnit val="${pick([1, 3, 1000000])}"/><c:majorTimeUnit val="${pick(['days', 'months', 'years'])}"/>`, 0.4) : '';
      return `<c:${name}><c:axId val="${id}"/><c:scaling><c:orientation val="${pick(['minMax', 'maxMin'])}"/></c:scaling>${common()}<c:crossAx val="${cross}"/>${crossing()}${time}</c:${name}>`;
    };
    const categories = (length: number): string => {
      const kind = random();
      if (kind < 0.25) return numLit(Array.from({ length }, (_value, i) => pick([45292 + i * pick([1, 31, 366]), i, -1, 2958465, null])));
      if (kind < 0.4) {
        const names = Array.from({ length }, (_value, i) => `<c:pt idx="${i}"><c:v>iç ${i} ${'x'.repeat(pick([0, 3, 40]))}</c:v></c:pt>`).join('');
        return `<c:multiLvlStrRef><c:f>S!$A$1:$B$9</c:f><c:multiLvlStrCache><c:ptCount val="${length}"/><c:lvl>${names}</c:lvl><c:lvl><c:pt idx="0"><c:v>dış</c:v></c:pt><c:pt idx="${pick([1, 2, 500])}"><c:v>öbür</c:v></c:pt></c:lvl></c:multiLvlStrCache></c:multiLvlStrRef>`;
      }
      return strLit(Array.from({ length }, (_value, i) => `Kategori ${i} ${'uzun '.repeat(pick([0, 0, 2, 30]))}`));
    };
    const ids = (a: number, b: number): string => `<c:axId val="${a}"/><c:axId val="${b}"/>`;
    const plain = (list: (number | null)[][], from: number, smooth = false): string =>
      list.map((values, i) => ser(from + i, { cat: maybe(categories(values.length), 0.7) || undefined, val: numLit(values), props: extras(false), tail: smooth ? maybe('<c:smooth val="1"/>', 0.5) : '' })).join('');
    const xy = (list: (number | null)[][], from: number, bubble: boolean): string =>
      list
        .map((values, i) => {
          const xs = random() < 0.2 ? undefined : numLit(values.map((_value, j) => (random() < 0.1 ? null : pick([j, j * 1e305, -j * 1e-310, j + random(), -j]))));
          const sizes = bubble ? maybe(`<c:bubbleSize>${numLit(values.map(() => pick([1, 4, 0, -3, null, 1e308, 5e-324])))}</c:bubbleSize>`, 0.8) : '';
          return xySer(from + i, { cat: xs, val: numLit(values), props: extras(true), tail: sizes + maybe('<c:smooth val="1"/>') });
        })
        .join('');
    const lines = (): string => maybe('<c:dropLines/>') + maybe('<c:hiLowLines/>') + maybe(`<c:upDownBars><c:gapWidth val="${pick([0, 150, 500])}"/><c:upBars/><c:downBars/></c:upDownBars>`);

    for (let round = 0; round < 260; round++) {
      const kind = pick(kinds);
      const grouping = pick(['standard', 'stacked', 'percentStacked']);
      const series: (number | null)[][] = Array.from({ length: Math.floor(random() * 4) + (round % 7 === 0 ? 0 : 1) }, data);
      const pair = categoryAxis(1, 2) + valueAxis(2, 1);
      let plot: string;
      if (kind === 'scatter' || kind === 'bubble') {
        const name = kind === 'bubble' ? 'bubbleChart' : 'scatterChart';
        const head = kind === 'bubble' ? flag('varyColors') : `<c:scatterStyle val="${pick(['lineMarker', 'smoothMarker', 'marker', 'line'])}"/>`;
        const tail = kind === 'bubble' ? `<c:bubbleScale val="${pick([0, 50, 100, 300])}"/>${flag('showNegBubbles')}<c:sizeRepresents val="${pick(['area', 'w'])}"/>` : '';
        const half = Math.ceil(series.length / 2);
        const twice = random() < 0.3;
        plot = `<c:${name}>${head}${xy(twice ? series.slice(0, half) : series, 0, kind === 'bubble')}${tail}${ids(1, 2)}</c:${name}>` + (twice ? `<c:${name}>${head}${xy(series.slice(half), half, kind === 'bubble')}${tail}${ids(3, 4)}</c:${name}>` : '') + valueAxis(1, 2) + valueAxis(2, 1) + (twice ? valueAxis(3, 4) + valueAxis(4, 3) : '');
      } else if (kind === 'pie' || kind === 'doughnut' || kind === 'ofPie') {
        const tail = kind === 'ofPie' ? `<c:splitType val="${pick(['auto', 'pos', 'val', 'percent', 'cust'])}"/><c:splitPos val="${number()}"/><c:custSplit><c:secondPiePt val="${pick([0, 1, 99])}"/></c:custSplit>` : `<c:firstSliceAng val="${pick([0, 90, 360])}"/>${kind === 'doughnut' ? `<c:holeSize val="${pick([1, 50, 90])}"/>` : ''}`;
        plot = `<c:${kind}Chart>${flag('varyColors')}${plain(series, 0)}${tail}</c:${kind}Chart>`;
      } else if (kind === 'radar') {
        plot = `<c:radarChart><c:radarStyle val="${pick(['standard', 'marker', 'filled'])}"/>${plain(series, 0)}${ids(1, 2)}</c:radarChart>${pair}`;
      } else if (kind === 'stock') {
        plot = `<c:stockChart>${plain(series, 0)}${lines()}${ids(1, 2)}</c:stockChart>${pair}`;
      } else if (kind === 'combo' || kind === 'second') {
        const half = Math.ceil(series.length / 2);
        const other = pick(['barChart', 'lineChart', 'areaChart']);
        const first = `<c:${other}><c:barDir val="${pick(['col', 'bar'])}"/><c:grouping val="${grouping}"/>${plain(series.slice(0, half), 0)}${ids(1, 2)}</c:${other}>`;
        if (kind === 'combo') plot = first + `<c:scatterChart><c:scatterStyle val="lineMarker"/>${xy(series.slice(half), half, false)}${ids(1, 2)}</c:scatterChart>` + pair;
        else plot = first + `<c:lineChart><c:grouping val="${pick(['standard', 'stacked'])}"/>${plain(series.slice(half), half, true)}${lines()}<c:marker val="1"/>${ids(3, 4)}</c:lineChart>` + pair + categoryAxis(3, 4) + valueAxis(4, 3);
      } else {
        const type = kind === 'barH' ? 'bar' : kind;
        const head = (type === 'bar' ? `<c:barDir val="${kind === 'barH' ? 'bar' : 'col'}"/>` : '') + `<c:grouping val="${type === 'bar' && grouping === 'standard' ? 'clustered' : grouping}"/>`;
        plot = `<c:${type}Chart>${head}${plain(series, 0, type === 'line')}${labels()}${type === 'bar' ? '' : lines()}${ids(1, 2)}</c:${type}Chart>${pair}`;
      }
      const head = title('Başlık').replace('<c:overlay', `${manual() === '<c:layout/>' ? '' : manual()}<c:overlay`);
      const xml = chartXml({ head, plot, tail: legend(pick(['r', 'b', 't', 'l', 'tr']), manual()) + `<c:plotVisOnly val="1"/><c:dispBlanksAs val="${pick(['gap', 'zero', 'span'])}"/>` }).replace('<c:plotArea><c:layout/>', `<c:plotArea>${manual()}`);
      const chart = readChart(xml);
      if (!chart) continue;
      const [width, height] = pick(sizes);
      const spec = `seed ${seed}, round ${round}: ${kind} ${grouping} ${width}x${height} ${JSON.stringify(series)} ${plot.replace(/<c:(cat|val|xVal|yVal)>.*?<\/c:\1>/g, '')}`;
      let layout: ChartLayout | undefined;
      assert.doesNotThrow(() => (layout = layoutChart(chart, chartValues(chart), width, height, fakeMeasure, { date1904: random() < 0.2 })), spec);
      assertSound(layout as ChartLayout, spec);
      for (const item of (layout as ChartLayout).items) seen.add(`${item.kind}/${(item as { role?: string }).role || ''}`);
      checked++;
    }
  }
  assert.ok(checked > 2400, `${checked} layouts were looked at`);
  // Every kind of thing a layout can hold came up: the charts above reach all of the layout.
  for (const wanted of ['rect/bar', 'rect/updown', 'path/line', 'path/area', 'path/trend', 'path/error', 'path/drop', 'path/hilow', 'path/grid', 'path/axis', 'markers/marker', 'circles/bubble', 'slice/', 'text/label', 'text/category', 'text/tick', 'text/axisTitle', 'text/legend', 'text/title']) {
    assert.ok(seen.has(wanted), `no layout held ${wanted}`);
  }
});
