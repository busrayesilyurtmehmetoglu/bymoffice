import { ZipArchive } from '../src/zip';

const encoder = new TextEncoder();
export const utf8 = (text: string) => encoder.encode(text);

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Smallest package that is recognisable as an xlsx workbook. */
export async function minimalXlsx(): Promise<Uint8Array> {
  const zip = new ZipArchive();
  zip.set(
    '[Content_Types].xml',
    utf8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '</Types>',
    ),
  );
  zip.set(
    '_rels/.rels',
    utf8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>',
    ),
  );
  zip.set(
    'xl/workbook.xml',
    utf8(
      `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        '<sheets><sheet name="Sayfa1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
  );
  zip.set(
    'xl/_rels/workbook.xml.rels',
    utf8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
        '</Relationships>',
    ),
  );
  zip.set('xl/worksheets/sheet1.xml', utf8(`${XML}<worksheet><sheetData/></worksheet>`));
  // A part the engine knows nothing about; it must survive a save untouched.
  zip.set('customXml/unknown.bin', new Uint8Array([0, 1, 2, 3, 255, 254, 253]));
  return zip.write();
}

/**
 * Version 3 compound file with one stream in regular sectors and one in the
 * mini stream. Layout: sector 0 FAT, 1 directory, 2 mini FAT, 3 mini stream,
 * 4..12 the "Workbook" stream.
 */
export function minimalCfb(big: Uint8Array, small: Uint8Array): Uint8Array {
  const SECTOR = 512;
  const END = 0xfffffffe;
  const FREE = 0xffffffff;
  if (big.length !== 9 * SECTOR || small.length > SECTOR) throw new Error('fixture sizes are fixed');

  const out = new Uint8Array(SECTOR * 14);
  const view = new DataView(out.buffer);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true); // FAT sector count
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, 4096, true); // mini stream cutoff
  view.setUint32(60, 2, true); // first mini FAT sector
  view.setUint32(64, 1, true);
  view.setUint32(68, END, true); // no DIFAT sectors
  view.setUint32(76, 0, true);
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, FREE, true);

  const sector = (n: number) => (n + 1) * SECTOR;

  const fat = sector(0);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(fat + i * 4, FREE, true);
  view.setUint32(fat, 0xfffffffd, true);
  for (const s of [1, 2, 3, 12]) view.setUint32(fat + s * 4, END, true);
  for (let s = 4; s < 12; s++) view.setUint32(fat + s * 4, s + 1, true);

  const entry = (index: number, name: string, type: number, child: number, right: number, start: number, size: number) => {
    const p = sector(1) + index * 128;
    for (let i = 0; i < name.length; i++) view.setUint16(p + i * 2, name.charCodeAt(i), true);
    view.setUint16(p + 64, name.length ? (name.length + 1) * 2 : 0, true);
    view.setUint8(p + 66, type);
    view.setUint32(p + 68, FREE, true);
    view.setUint32(p + 72, right, true);
    view.setUint32(p + 76, child, true);
    view.setUint32(p + 116, start, true);
    view.setUint32(p + 120, size, true);
  };
  entry(0, 'Root Entry', 5, 1, FREE, 3, SECTOR);
  entry(1, 'Workbook', 2, FREE, 2, 4, big.length);
  entry(2, 'Small', 2, FREE, FREE, 0, small.length);
  entry(3, '', 0, FREE, FREE, 0, 0);

  const miniFat = sector(2);
  const miniSectors = Math.ceil(small.length / 64);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(miniFat + i * 4, FREE, true);
  for (let i = 0; i < miniSectors; i++) view.setUint32(miniFat + i * 4, i === miniSectors - 1 ? END : i + 1, true);

  out.set(small, sector(3));
  out.set(big, sector(4));
  return out;
}

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface XlsxSheetSpec {
  name: string;
  /** Complete worksheet (or chartsheet) XML. */
  xml: string;
  state?: 'hidden' | 'veryHidden';
  kind?: 'worksheet' | 'chartsheet';
  /** Relationships of the sheet part: id -> external target. */
  links?: Record<string, string>;
}

export interface XlsxSpec {
  sheets: XlsxSheetSpec[];
  /** Inner XML of <sst>. */
  sharedStrings?: string;
  /** Complete styles part. */
  styles?: string;
  /** Complete theme part. */
  theme?: string;
  /** Extra XML placed inside <workbook> before <sheets> (workbookPr, bookViews). */
  workbookHead?: string;
  /** Extra XML placed inside <workbook> after <sheets> (definedNames). */
  workbookTail?: string;
  /** Complete cell metadata part (xl/metadata.xml). */
  metadata?: string;
}

/** Wraps sheet content in a worksheet root element. */
export const worksheetXml = (inner: string) => `${XML}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${inner}</worksheet>`;

/** Builds a workbook package from XML fragments, the way Excel lays one out. */
export async function buildXlsx(spec: XlsxSpec): Promise<Uint8Array> {
  const zip = new ZipArchive();
  const overrides: string[] = [];
  const workbookRels: string[] = [];
  const sheetTags: string[] = [];
  const parts: [string, string][] = [];

  spec.sheets.forEach((sheet, i) => {
    const chart = sheet.kind === 'chartsheet';
    const path = chart ? `chartsheets/sheet${i + 1}.xml` : `worksheets/sheet${i + 1}.xml`;
    const type = chart ? 'chartsheet' : 'worksheet';
    overrides.push(`<Override PartName="/xl/${path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${type}+xml"/>`);
    workbookRels.push(`<Relationship Id="rId${i + 1}" Type="${NS_REL}/${type}" Target="${path}"/>`);
    sheetTags.push(`<sheet name="${sheet.name}" sheetId="${i + 1}"${sheet.state ? ` state="${sheet.state}"` : ''} r:id="rId${i + 1}"/>`);
    parts.push([`xl/${path}`, sheet.xml]);
    if (sheet.links) {
      const rels = Object.entries(sheet.links)
        .map(([id, target]) => `<Relationship Id="${id}" Type="${NS_REL}/hyperlink" Target="${target}" TargetMode="External"/>`)
        .join('');
      const slash = path.lastIndexOf('/');
      parts.push([`xl/${path.slice(0, slash)}/_rels/${path.slice(slash + 1)}.rels`, `${XML}<Relationships xmlns="${NS_PKG_REL}">${rels}</Relationships>`]);
    }
  });

  let next = spec.sheets.length + 1;
  const optional = (content: string | undefined, path: string, type: string, contentType: string) => {
    if (content === undefined) return;
    overrides.push(`<Override PartName="/xl/${path}" ContentType="${contentType}"/>`);
    workbookRels.push(`<Relationship Id="rId${next++}" Type="${NS_REL}/${type}" Target="${path}"/>`);
    parts.push([`xl/${path}`, content]);
  };
  optional(
    spec.sharedStrings === undefined ? undefined : `${XML}<sst xmlns="${NS_MAIN}">${spec.sharedStrings}</sst>`,
    'sharedStrings.xml',
    'sharedStrings',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  );
  optional(spec.styles, 'styles.xml', 'styles', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
  optional(spec.theme, 'theme/theme1.xml', 'theme', 'application/vnd.openxmlformats-officedocument.theme+xml');
  optional(spec.metadata, 'metadata.xml', 'sheetMetadata', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml');

  zip.set(
    '[Content_Types].xml',
    utf8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        `${overrides.join('')}</Types>`,
    ),
  );
  zip.set(
    '_rels/.rels',
    utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
  );
  zip.set(
    'xl/workbook.xml',
    utf8(
      `${XML}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}">${spec.workbookHead ?? ''}` +
        `<sheets>${sheetTags.join('')}</sheets>${spec.workbookTail ?? ''}</workbook>`,
    ),
  );
  zip.set('xl/_rels/workbook.xml.rels', utf8(`${XML}<Relationships xmlns="${NS_PKG_REL}">${workbookRels.join('')}</Relationships>`));
  for (const [path, content] of parts) zip.set(path, utf8(content));
  return zip.write();
}

/** A workbook that exercises most of what the xlsx reader understands. */
export function sampleWorkbook(): Promise<Uint8Array> {
  return buildXlsx({
    workbookHead: '<workbookPr date1904="1"/><bookViews><workbookView activeTab="1"/></bookViews>',
    workbookTail:
      '<definedNames><definedName name="Toplam">Veri!$B$2:$B$4</definedName>' +
      '<definedName name="_xlnm.Print_Area" localSheetId="0" hidden="1">Veri!$A$1:$C$4</definedName></definedNames>',
    sharedStrings:
      '<si><t>Ürün</t></si><si><t>Fiyat</t></si><si><t xml:space="preserve"> Çay </t></si>' +
      '<si><r><t>Kalın</t></r><r><rPr><b/><sz val="14"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t> şeker</t></r></si>',
    styles:
      `${XML}<styleSheet xmlns="${NS_MAIN}">` +
      '<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₺&quot;"/></numFmts>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="12"/><color theme="4"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border><left style="thin"><color auto="1"/></left><right/><top/><bottom style="medium"><color rgb="FF0000FF"/></bottom><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0"><alignment horizontal="center" wrapText="1"/></xf>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
      '</styleSheet>',
    theme:
      `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office">` +
      '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
      '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
      '</a:clrScheme></a:themeElements></a:theme>',
    sheets: [
      {
        name: 'Veri',
        links: { rId1: 'https://example.com/çay' },
        xml: worksheetXml(
          '<dimension ref="A1:C4"/>' +
            '<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
            '<selection pane="bottomRight" activeCell="B3" sqref="B3"/></sheetView></sheetViews>' +
            '<sheetFormatPr defaultRowHeight="15"/>' +
            '<cols><col min="1" max="1" width="18.5" customWidth="1"/><col min="3" max="4" width="9" hidden="1"/></cols>' +
            '<sheetData>' +
            '<row r="1" ht="24" customHeight="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" s="1" t="s"><v>1</v></c></row>' +
            '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="2"><v>12.5</v></c><c r="C2" t="b"><v>1</v></c></row>' +
            '<row r="3" hidden="1"><c r="A3" t="s"><v>3</v></c><c r="B3" s="2"><v>7</v></c><c r="C3" t="e"><v>#DIV/0!</v></c></row>' +
            '<row r="4"><c r="A4" t="inlineStr"><is><t>Toplam</t></is></c><c r="B4" s="2"><f>SUM(B2:B3)</f><v>19.5</v></c></row>' +
            '</sheetData>' +
            '<mergeCells count="1"><mergeCell ref="A5:C6"/></mergeCells>' +
            '<hyperlinks><hyperlink ref="A2" r:id="rId1" tooltip="Çay"/><hyperlink ref="A3" location="Gizli!A1"/></hyperlinks>',
        ),
      },
      { name: 'Gizli', state: 'hidden', xml: worksheetXml('<sheetData/>') },
      {
        name: 'Grafik',
        kind: 'chartsheet',
        xml: `${XML}<chartsheet xmlns="${NS_MAIN}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>`,
      },
    ],
  });
}
