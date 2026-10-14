// Charts of the newer types (chartex) in a Word document: Word writes the frame in alternate
// content, the Choice with the chart and a Fallback (a picture) beside it. The Choice is taken
// where the chart part can be drawn, the Fallback where it cannot, never both; the chart is
// drawn from the values its part cached; nothing of it is ever written anew, not even when the
// paragraph that holds it is edited.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { emptyWriteReport, writeDocx } from '../src/formats/docx/write';
import { open as openFile } from '../src/index';
import type { ImageItem, PageItem } from '../src/layout/doc/types';
import { chartValues, layoutChart } from '../src/layout/chart';
import type { Block, DrawingItem, Paragraph, RunItem } from '../src/model/doc/types';
import { DocHistory, insertText } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { fakeMeasure } from './chart-fixtures';
import { chartParts, columnsChart, inlineChart, partRels } from './docx-chart-fixtures';
import { lay, p } from './doc-layout-fixtures';
import { buildDocx, documentXml, THEME, type DocxSpec } from './docx-fixtures';
import { CHARTEX_REL, CHARTEX_TYPE, chartExRun, chartExXml, data, NS_CX1, numDim, series, strDim, treemapXml, waterfallXml } from './chartex-fixtures';

const PART = '/word/charts/chartEx1.xml';
/** The smallest PNG: one white pixel. */
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00,
  0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xff, 0xff, 0x3f, 0x00, 0x05, 0xfe, 0x02, 0xfe, 0xa7, 0x35, 0x81, 0x84, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

const regionMap = () => chartExXml({ data: data(0, strDim(undefined, [['Türkiye']]) + numDim(undefined, [1])), series: series({ layout: 'regionMap', dataId: 0 }) });

/** The parts of a document whose body shows the chart part `xml` through "rIdX1", with the picture "rIdP1" as its fallback. */
function parts(xml: string): Required<Pick<DocxSpec, 'rels' | 'parts'>> {
  return {
    rels: [
      ['rIdX1', CHARTEX_REL, 'charts/chartEx1.xml'],
      ['rIdP1', 'image', 'media/fallback.png'],
    ],
    parts: [
      ['word/charts/chartEx1.xml', xml],
      ['word/media/fallback.png', PNG],
    ],
  };
}

/** The run of a chart, with the namespace its Choice requires declared on the Choice (Word declares it on the root). */
const run = () => chartExRun('rIdX1', 300, 200, 'rIdP1').replace('<mc:Choice Requires="cx1">', `<mc:Choice xmlns:cx1="${NS_CX1}" Requires="cx1">`);

function alternates(blocks: Block[]): Extract<RunItem, { type: 'alternate' }>[] {
  const out: Extract<RunItem, { type: 'alternate' }>[] = [];
  for (const block of blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of (block as Paragraph).children) {
      if (inline.type !== 'run') continue;
      for (const item of inline.content) if (item.type === 'alternate') out.push(item);
    }
  }
  return out;
}

const images = (pages: { items: PageItem[] }[]): ImageItem[] => pages.reduce<PageItem[]>((out, page) => out.concat(page.items), []).filter((item): item is ImageItem => item.type === 'image');

test('a chart of a newer type is the Choice of its alternate content, drawn from its part', async () => {
  const laid = await lay({ body: p('önce') + p(run()) + p('sonra'), theme: THEME, ...parts(waterfallXml()) });
  const doc = laid.doc;
  const [alternate] = alternates(doc.body.blocks);
  assert.equal(alternate.selected, 'choice');
  const drawing = alternate.items.find((item): item is DrawingItem => item.type === 'drawing');
  assert.ok(drawing);
  assert.equal(drawing.kind, 'chart');
  assert.equal(drawing.chartPart, PART);
  const chart = doc.charts?.get(PART);
  assert.ok(chart && chart.extended);
  assert.equal(chart.groups[0].kind, 'waterfall');
  assert.deepEqual(chartValues(chart).series[0].values.toArray(), [100, 20, -50, 30, 100], 'from the cache of the part');
  const [image] = images(laid.pages);
  assert.equal(image.placeholder, 'chart');
  assert.equal(image.chart, chart);
  assert.ok(layoutChart(chart, chartValues(chart), 300, 200, fakeMeasure).items.some((item) => item.kind === 'rect' && item.role === 'bar'));
  assert.deepEqual(doc.warnings, []);
});

test('the namespace the Choice requires may be declared on the root of the document part, as Word does', async () => {
  const body = p(chartExRun('rIdX1', 300, 200, 'rIdP1'));
  const document = documentXml(body).replace('<w:document ', `<w:document xmlns:cx1="${NS_CX1}" `);
  const doc = await readDocx(await OpcPackage.open(await buildDocx({ document, theme: THEME, ...parts(treemapXml()) })));
  assert.equal(alternates(doc.body.blocks)[0].selected, 'choice');
  assert.equal(doc.charts?.get(PART)?.groups[0].kind, 'treemap');
});

test('a chart of a newer type that is not drawn shows the fallback of its alternate content, and no warning', async () => {
  for (const xml of [regionMap(), '<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart>']) {
    const laid = await lay({ body: p(run()), theme: THEME, ...parts(xml) });
    const [alternate] = alternates(laid.doc.body.blocks);
    assert.equal(alternate.selected, 'fallback');
    const [image] = images(laid.pages);
    assert.equal(image.part, '/word/media/fallback.png', 'the picture of the fallback');
    assert.equal(image.chart, undefined);
    assert.equal(laid.doc.charts, undefined);
    assert.deepEqual(laid.doc.warnings, []);
  }
  // A Choice that requires something else beside the chart is not taken.
  const other = await lay({ body: p(run().replace('Requires="cx1"', 'Requires="cx1 foo"')), theme: THEME, ...parts(waterfallXml()) });
  assert.equal(alternates(other.doc.body.blocks)[0].selected, 'fallback');
});

test('a frame of a newer type without alternate content: the chart, or a placeholder with a warning', async () => {
  const bare = (xml: string) => {
    const inner = run();
    const drawing = inner.slice(inner.indexOf('<w:drawing>'), inner.indexOf('</w:drawing>') + '</w:drawing>'.length);
    return lay({ body: p(`<w:r>${drawing}</w:r>`), theme: THEME, ...parts(xml) });
  };
  const shown = await bare(waterfallXml());
  assert.ok(images(shown.pages)[0].chart);
  const missing = await bare(regionMap());
  const [image] = images(missing.pages);
  assert.equal(image.placeholder, 'chart');
  assert.equal(image.chart, undefined);
  assert.ok(missing.doc.warnings.some((line) => /chartEx1\.xml.*nothing that is drawn/.test(line)), JSON.stringify(missing.doc.warnings));
});

test('lossless: the document is saved with its chart, its fallback and its alternate content as they were, also when the paragraph of the chart is edited', async () => {
  const content = await buildDocx({ body: p('metin') + p(`<w:r><w:t>Grafik: </w:t></w:r>${run()}`), theme: THEME, ...parts(waterfallXml()) });
  // A content type of its own, as Word gives the part.
  const zip = ZipArchive.read(content);
  const types = new TextDecoder().decode(await zip.read('[Content_Types].xml')).replace('</Types>', `<Override PartName="${PART}" ContentType="${CHARTEX_TYPE}"/></Types>`);
  zip.set('[Content_Types].xml', new TextEncoder().encode(types));
  const bytes = await zip.write();
  const file = await openFile(bytes);
  assert.ok((await file.document()).charts?.has(PART));
  assert.deepEqual(await file.save(), bytes, 'opened and saved: the bytes of the file');

  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  history.applyEdit(insertText(doc, { story: 'body', path: [1], offset: 0 }, 'Yeni '), undefined);
  history.applyEdit(insertText(doc, { story: 'body', path: [0], offset: 5 }, ' daha'), undefined);
  await writeDocx(pkg, doc, { report: emptyWriteReport() });
  const saved = await pkg.save();
  const xml = new TextDecoder().decode(await pkg.readPart('/word/document.xml'));
  const alternate = run().slice('<w:r>'.length, run().length - '</w:r>'.length);
  assert.ok(xml.indexOf(alternate) >= 0, 'the alternate content, both branches, character for character');
  for (const name of ['word/charts/chartEx1.xml', 'word/media/fallback.png', 'word/_rels/document.xml.rels']) {
    assert.deepEqual(await ZipArchive.read(saved).read(name), await ZipArchive.read(bytes).read(name), name);
  }
  const again = await readDocx(await OpcPackage.open(saved));
  assert.equal(alternates(again.body.blocks)[0].selected, 'choice');
  assert.ok(again.charts?.has(PART));
});

test('charts of newer types share the limits of the document with the other charts', async () => {
  const doc = await readDocx(await OpcPackage.open(await buildDocx({ body: p(run()), theme: THEME, ...parts(waterfallXml()) })), { maxChartPoints: 5 });
  // Ten cached values do not fit five: the chart is not read, the fallback is shown, and nothing is said of a chart the document does not show.
  assert.equal(alternates(doc.body.blocks)[0].selected, 'fallback');
  assert.equal(doc.charts, undefined);
  assert.deepEqual(doc.warnings, []);
  const none = await readDocx(await OpcPackage.open(await buildDocx({ body: p(run()), theme: THEME, ...parts(waterfallXml()) })), { maxCharts: 0 });
  assert.equal(alternates(none.body.blocks)[0].selected, 'fallback');
});

test('a part of a newer type no frame shows is not read: it takes nothing of the limits the shown charts share', async () => {
  const columns = await chartParts([{ rel: 'rIdC1', xml: columnsChart() }]);
  const spec = (orphan: boolean): DocxSpec => ({
    body: p(inlineChart('rIdC1', 300, 200)),
    theme: THEME,
    rels: orphan ? [...columns.rels, ['rIdX9', CHARTEX_REL, 'charts/chartEx9.xml']] : columns.rels,
    parts: orphan ? [...columns.parts, ['word/charts/chartEx9.xml', waterfallXml()]] : columns.parts,
  });
  // One chart of six cached values: the limits of a document that holds that chart alone.
  for (const limits of [{ maxCharts: 1 }, { maxChartPoints: 6 }]) {
    assert.deepEqual(Array.from((await readDocx(await OpcPackage.open(await buildDocx(spec(false))), limits)).charts?.keys() || []), ['/word/charts/chart1.xml'], 'without the part');
    const doc = await readDocx(await OpcPackage.open(await buildDocx(spec(true))), limits);
    assert.deepEqual(Array.from(doc.charts?.keys() || []), ['/word/charts/chart1.xml'], JSON.stringify(limits));
    assert.deepEqual(doc.warnings, [], `${JSON.stringify(limits)}: nothing is said of a part no frame shows`);
  }
  // A frame in a header shows its chart of a newer type.
  const header = await readDocx(
    await OpcPackage.open(
      await buildDocx({
        body: p('gövde'),
        theme: THEME,
        headers: { rIdH1: p(run()) },
        parts: [
          ['word/charts/chartEx1.xml', waterfallXml()],
          ['word/media/fallback.png', PNG],
          partRels('header1.xml', [
            ['rIdX1', CHARTEX_REL, 'charts/chartEx1.xml'],
            ['rIdP1', 'image', 'media/fallback.png'],
          ]),
        ],
      }),
    ),
  );
  assert.ok(header.charts?.has(PART), 'the chart of the header is read');
});

test('a part of many unclosed chart tags is looked through in one pass, not once from every tag', async () => {
  // Each "<c:chart " is never closed: a search that reads to the next ">" from each of them reads the rest of the part
  // again and again (a megabyte took minutes). The document is broken XML and is refused; what is timed is the refusal.
  const document = documentXml(p('x')).replace('<w:body>', `<w:body>${'<c:chart '.repeat(120_000)}`);
  const started = Date.now();
  await assert.rejects(async () => readDocx(await OpcPackage.open(await buildDocx({ document, theme: THEME, ...parts(waterfallXml()) }))));
  assert.ok(Date.now() - started < 10_000, `took ${Date.now() - started} ms`);
});
