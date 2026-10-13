// What the tests of copied charts share: workbooks whose charts lead to parts of their own
// (a style, colours, a picture), charts that name no sheet and a chart with shapes drawn over
// it. (The browser tests use them as well.)

import { ZipArchive } from '../src/index';
import { FOREIGN, salesBook, text } from './chart-edit-fixtures';
import { catAx, chartFrameOf, chartXml, group, numRef, ser, sheetData, strRef, valAx } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const encoder = new TextEncoder();
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const RELS = (inner: string) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${inner}</Relationships>`;

export const frameAt = (rel: string, name: string, row: number, id: number) => twoCell([4, row], [9, row + 11], chartFrameOf(rel, name).replace('id="2"', `id="${id}"`));

/**
 * A chart as another application writes one (FOREIGN) whose caches say what the cells say, and
 * that fills its area with a picture: unchanged, its part is written with the bytes it has.
 */
export const RICH = FOREIGN.split(numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]))
  .join(numRef('Sayfa1!$B$2:$B$4', [10, 20, 30]))
  .split(numRef('Sayfa1!$C$2:$C$4', [9, 9, 9]))
  .join(numRef('Sayfa1!$C$2:$C$4', [30, 20, 10]))
  .replace('</c:chart><c:txPr>', '</c:chart><c:spPr><a:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></c:spPr><c:txPr>');

const BAR = '<c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>';

/** A chart whose references name no sheet, with a gradient no model holds. */
export const SHEETLESS = chartXml({
  plot:
    group(
      'barChart',
      BAR,
      ser(0, { tx: strRef('$B$1', ['North']), cat: strRef('$A$2:$A$4', ['Jan', 'Feb', 'Mar']), val: numRef('$B$2:$B$4', [10, 20, 30]), props: '<c:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="112233"/></a:gs></a:gsLst></a:gradFill></c:spPr>' }),
      '<c:gapWidth val="77"/>',
    ) +
    catAx() +
    valAx(),
});

/** RICH with shapes drawn over it: a part that leads out of the chart, which no copy takes along. */
export const SHAPED = RICH.replace('<c:extLst><c:ext uri="{781A3756', '<c:userShapes r:id="rId4"/><c:extLst><c:ext uri="{781A3756');

/** A chart with a formula in an extension: a reference no model holds. */
export const UNKNOWN = RICH.replace(
  '<c14:pivotOptions xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"/>',
  '<c15:fullRef xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:sqref>Sayfa1!$B$2:$B$9</c15:sqref></c15:fullRef>',
);

/** A chart part in Latin-1 that says so, with a title that is not ASCII ("Café üö"): one byte for each letter. */
export const LATIN_TITLE = 'Café üö';
export const latinChart = (xml: string = RICH): Uint8Array =>
  Uint8Array.from(xml.replace('>Old<', `>${LATIN_TITLE}<`).replace(/^<\?xml[^?]*\?>/, '<?xml version="1.0" encoding="ISO-8859-1"?>'), (ch) => ch.charCodeAt(0));

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89]);
export const STYLE = '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201"><cs:axisTitle><cs:lnRef idx="0"/></cs:axisTitle></cs:chartStyle>';
export const COLORS = '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10"><a:schemeClr val="accent1"/></cs:colorStyle>';

export interface RichSpec {
  /** The first chart of the first sheet; RICH by default. */
  first?: string | Uint8Array;
  /** The first chart has shapes drawn over it (SHAPED): their part and its relationship are in the package. */
  shapes?: boolean;
}

/**
 * A workbook with the numbers of SALES on "Sayfa1" and some on "Other". "Sayfa1" has two
 * charts: one that leads to a style, colours and a picture (chart1), and one whose references
 * name no sheet (chart2).
 */
export async function richBook(spec: RichSpec = {}): Promise<Uint8Array> {
  const base = await salesBook({
    drawing: frameAt('rIdC1', 'Grafik 1', 1, 2) + frameAt('rIdC2', 'Grafik 2', 14, 3),
    charts: [spec.first !== undefined ? spec.first : spec.shapes ? SHAPED : RICH, SHEETLESS],
    sheets: [
      {
        name: 'Other',
        xml: sheetData([
          ['Ad', 'Deger'],
          ['x', 5],
          ['y', 7],
          ['z', 9],
        ]),
      },
    ],
  });
  const zip = ZipArchive.read(base);
  zip.set('xl/charts/style1.xml', encoder.encode(STYLE));
  zip.set('xl/charts/colors1.xml', encoder.encode(COLORS));
  zip.set('xl/media/image1.png', PNG);
  let rels =
    '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/>' +
    `<Relationship Id="rId3" Type="${REL}/image" Target="../media/image1.png"/>`;
  let types =
    '<Default Extension="png" ContentType="image/png"/><Override PartName="/xl/charts/style1.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/><Override PartName="/xl/charts/colors1.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>';
  if (spec.shapes) {
    rels += `<Relationship Id="rId4" Type="${REL}/chartUserShapes" Target="../drawings/drawing7.xml"/>`;
    types += '<Override PartName="/xl/drawings/drawing7.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml"/>';
    zip.set(
      'xl/drawings/drawing7.xml',
      encoder.encode('<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'),
    );
  }
  zip.set('xl/charts/_rels/chart1.xml.rels', encoder.encode(RELS(rels)));
  zip.set('[Content_Types].xml', encoder.encode((await text(base, '[Content_Types].xml')).replace('</Types>', `${types}</Types>`)));
  return zip.write();
}
