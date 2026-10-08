// A workbook with conditional formats of every kind of place: several rules in one element,
// a rule of each family, a data bar with its counterpart in the extension list, an icon set,
// and (on request) a rule that only the extension list has.

import type { CellRange, ConditionalFormat, Workbook } from '../src/index';
import { rangeName } from '../src/index';
import { conditionalFormatNow, mainPartRule, ruleSpace } from '../src/ops/rules';
import { buildXlsx, worksheetXml } from './fixtures';
import { THEME } from './sheet-drawing-fixtures';

export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"';
const XM = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';

/** 0 red fill, 1 bold, 2 a fill with a border and an alignment the model has no field for. */
export const DXFS =
  '<dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>' +
  '<dxf><font><b/></font></dxf>' +
  '<dxf><font><i/><sz val="14"/></font><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill><alignment horizontal="center"/><border><left style="thin"><color rgb="FF000000"/></left></border></dxf>';
export const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS_MAIN}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><dxfs count="3">${DXFS}</dxfs></styleSheet>`;

export const cf = (sqref: string, rules: string, extra = '') => `<conditionalFormatting sqref="${sqref}"${extra}>${rules}</conditionalFormatting>`;

// Written the way no writer of ours would: odd spacing, an attribute and a comment the model has no field for.
export const FIRST_RULE = '<cfRule type="cellIs"  dxfId="0" priority="2" operator="greaterThan" xr:uid="{R1}"><formula>$H$1</formula><!-- kept --></cfRule>';
export const SECOND_RULE = "<cfRule type='expression' dxfId='1' priority='1' stopIfTrue='1'><formula>AND(A1&gt;B1,Lists!$A$1&lt;&gt;A1)</formula></cfRule>";
export const BLOCK_A = cf('A1:A10 F1:F3', `\n  ${FIRST_RULE}\n  ${SECOND_RULE}\n`, ' xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"');
export const BLOCK_B = cf('B1:B5', '<cfRule type="colorScale" priority="3"><colorScale><cfvo type="min"/><cfvo type="formula" val="$H$2"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule>');
export const BAR_ID = '{DA7ABA51-AAAA-BBBB-0001-000000000001}';
export const BLOCK_C = cf(
  'C1:C5',
  `<cfRule type="dataBar" priority="4"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" ${X14}><x14:id>${BAR_ID}</x14:id></ext></extLst></cfRule>`,
);
export const BLOCK_D = cf('D1:D5', '<cfRule type="iconSet" priority="5"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule><cfRule type="timePeriod" dxfId="2" priority="6" timePeriod="today"><formula>FLOOR(D1,1)=TODAY()</formula></cfRule>');
export const EXT_BAR =
  `<x14:conditionalFormatting ${XM}><x14:cfRule type="dataBar" id="${BAR_ID}"><x14:dataBar minLength="0" maxLength="100" border="1"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/></x14:dataBar></x14:cfRule><xm:sqref>C1:C5</xm:sqref></x14:conditionalFormatting>`;
export const EXT_ONLY =
  `<x14:conditionalFormatting ${XM}><x14:cfRule type="iconSet" priority="7" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>E1:E5</xm:sqref></x14:conditionalFormatting>`;
export const extLst = (inner: string) => `<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ${X14}><x14:conditionalFormattings>${inner}</x14:conditionalFormattings></ext></extLst>`;

export const CELLS = '<sheetData><row r="1"><c r="A1"><v>5</v></c><c r="B1"><v>1</v></c><c r="H1"><v>3</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="H2"><v>2</v></c></row><row r="3"><c r="A3"><v>9</v></c><c r="B3"><v>3</v></c></row></sheetData>';
export const TAIL = '<hyperlinks><hyperlink ref="H1" location="Data!A1"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';

/** "Data" with the formats above, "Lists" and "Other". `extensionOnly` adds the rule only the extension list has. */
/** Something else a sheet keeps in its extension list: not a rule, and nobody's to change. */
export const OTHER_EXT = '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Data!A1:A3</xm:f><xm:sqref>J1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>';

export function conditionalBook(options: { extensionOnly?: boolean; bare?: boolean; otherExt?: boolean; barPriority?: boolean } = {}): Promise<Uint8Array> {
  const formats = options.bare ? '' : BLOCK_A + BLOCK_B + BLOCK_C + BLOCK_D;
  const bar = options.barPriority ? EXT_BAR.replace('type="dataBar" id=', 'type="dataBar" priority="4" id=') : EXT_BAR;
  let ext = options.bare ? '' : extLst(bar + (options.extensionOnly ? EXT_ONLY : ''));
  if (options.otherExt) ext = ext.replace('</extLst>', `${OTHER_EXT}</extLst>`);
  return buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(CELLS + formats + TAIL + ext) },
      { name: 'Lists', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' + cf('A1:B3', '<cfRule type="cellIs" dxfId="1" priority="1" operator="equal"><formula>Data!$H$1</formula></cfRule>')) },
      { name: 'Other', xml: worksheetXml('<sheetData/>') },
    ],
    styles: STYLES,
  });
}

const refs = (ranges: CellRange[]) => ranges.map(rangeName).join(' ');

/** The rules of a sheet as they apply now, in order of priority: "ranges|type|priority|formulas". */
export function conditionalNow(wb: Workbook, sheetIndex = 0): string[] {
  const sheet = wb.sheets[sheetIndex];
  const space = ruleSpace(wb, sheet);
  const out: { priority: number; order: number; text: string }[] = [];
  (sheet.conditionalFormats || []).forEach((format) => {
    const now = conditionalFormatNow(format, space) as ConditionalFormat | undefined;
    if (!now) return;
    for (const rule of now.rules) {
      const more = rule.colorScale ? rule.colorScale.values.map((value) => value.value ?? value.type).join(';') : rule.operator ?? rule.text ?? '';
      out.push({ priority: rule.priority >= 0 ? rule.priority : Infinity, order: out.length, text: `${refs(now.ranges)}|${rule.type}|${rule.priority}|${rule.formulas.join(';')}${more === '' ? '' : `|${more}`}` });
    }
  });
  out.sort((a, b) => (a.priority === b.priority ? a.order - b.order : a.priority < b.priority ? -1 : 1));
  return out.map((entry) => entry.text);
}

/** The rules of every sheet as they apply now, in order of priority, with the format each gives: plain data to compare. */
export function conditionalContent(wb: Workbook, mainOnly = false): unknown[] {
  const dxfs = wb.styles.dxfFormats || [];
  return wb.sheets.map((sheet) => {
    const space = ruleSpace(wb, sheet);
    const out: { priority: number; order: number; data: unknown }[] = [];
    (sheet.conditionalFormats || []).forEach((format) => {
      const now = conditionalFormatNow(format, space) as ConditionalFormat | undefined;
      if (!now) return;
      for (const stated of now.rules) {
        // What a file built from the model says: the rules of the main part, without what only the extension list knows.
        const rule = mainOnly ? mainPartRule(stated) : stated;
        if (!rule) continue;
        const dxf = rule.dxf || (rule.dxfId !== undefined ? dxfs[rule.dxfId] : undefined);
        out.push({
          priority: rule.priority >= 0 ? rule.priority : Infinity,
          order: out.length,
          data: {
            ranges: refs(now.ranges),
            type: rule.type,
            priority: rule.priority === Infinity ? 'none' : rule.priority,
            stopIfTrue: rule.stopIfTrue,
            operator: rule.operator,
            formulas: rule.formulas,
            text: rule.text,
            timePeriod: rule.timePeriod,
            rank: rule.rank,
            percent: rule.percent,
            bottom: rule.bottom,
            aboveAverage: rule.aboveAverage,
            equalAverage: rule.equalAverage,
            colorScale: rule.colorScale,
            dataBar: rule.dataBar,
            iconSet: rule.iconSet,
            extended: rule.extended === true,
            dxf,
          },
        });
      }
    });
    out.sort((a, b) => (a.priority === b.priority ? a.order - b.order : a.priority < b.priority ? -1 : 1));
    return [sheet.name, out.map((entry) => entry.data)];
  });
}

/** Records for REVIEW_BOOK: 0 a border alone, 1 a number format alone, 2 a gradient alone, 3 bold stated as off with a size, a fill and a border. */
export const REVIEW_DXFS =
  '<dxf><border><left style="thin"><color rgb="FF000000"/></left></border></dxf>' +
  '<dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf>' +
  '<dxf><fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color rgb="FF4472C4"/></stop></gradientFill></fill></dxf>' +
  '<dxf><font><b val="0"/><sz val="14"/><color theme="1"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill><border><left style="thin"><color rgb="FF000000"/></left></border></dxf>';

/**
 * Rules that say more than the dialogs have fields for, each in an element of its own, in this
 * order of priority: a scale with thresholds of three kinds, a bar with lengths and no value
 * shown, an average with deviations, a text rule written for "$A1", and four "cell value"
 * rules with the records above.
 */
export const REVIEW_RULES = [
  cf('A1:A9', '<cfRule type="colorScale" priority="1"><colorScale><cfvo type="num" val="0"/><cfvo type="formula" val="$H$2"/><cfvo type="percent" val="90"/><color rgb="FFF8696B"/><color theme="4"/><color rgb="FF63BE7B"/></colorScale></cfRule>'),
  cf('B1:B9', '<cfRule type="dataBar" priority="2"><dataBar minLength="5" maxLength="80" showValue="0"><cfvo type="num" val="0"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>'),
  cf('C1:C9', '<cfRule type="aboveAverage" dxfId="3" priority="3" aboveAverage="0" stdDev="2"/>'),
  cf('A1:A9', '<cfRule type="containsText" dxfId="3" priority="4" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",$A1)))</formula></cfRule>'),
  cf('D1:D9', '<cfRule type="cellIs" dxfId="0" priority="5" operator="greaterThan"><formula>SUM( $H$1:$H$2 )</formula></cfRule>'),
  cf('E1:E9', '<cfRule type="cellIs" dxfId="1" priority="6" operator="between"><formula>1</formula><formula>SUM( $H$1:$H$2 )</formula></cfRule>'),
  cf('F1:F9', '<cfRule type="cellIs" dxfId="2" priority="7" operator="equal"><formula>"x"</formula></cfRule>'),
  cf('G1:G9', '<cfRule type="cellIs" dxfId="3" priority="8" operator="equal"><formula>5</formula></cfRule>'),
  cf('A1:A9', '<cfRule type="beginsWith" priority="9" operator="beginsWith" text="q"><formula>LEFT(A1,1)="Q"</formula></cfRule>'),
];

export function reviewBook(): Promise<Uint8Array> {
  return buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(CELLS + REVIEW_RULES.join('') + TAIL) }], styles: STYLES.replace(/<dxfs count="3">.*<\/dxfs>/, `<dxfs count="4">${REVIEW_DXFS}</dxfs>`), theme: THEME });
}

/** REVIEW_RULES for a style sheet in which REVIEW_DXFS follow `offset` other records. */
export const reviewRules = (offset: number): string => REVIEW_RULES.join('').replace(/dxfId="(\d+)"/g, (_, id) => `dxfId="${Number(id) + offset}"`);

/** Elements a reader does not take into the model: a format without a rule, one whose range is none, a validation whose range is none. */
export const UNMODELLED_FORMATS = '<conditionalFormatting sqref="J1:J5" pivot="1"></conditionalFormatting><conditionalFormatting sqref="XFE1:XFE5"><cfRule type="cellIs" priority="77" operator="equal"><formula>1</formula></cfRule></conditionalFormatting>';
export const UNMODELLED_VALIDATION = '<dataValidation type="whole" sqref="XFE9"><formula1>1</formula1><formula2>2</formula2></dataValidation>';
