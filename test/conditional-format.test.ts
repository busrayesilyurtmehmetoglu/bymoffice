import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CalcEngine } from '../src/calc/engine';
import { ConditionalFormats, ICON_SETS, differentialFillColor, iconGlyph, type CellConditional, type ConditionalOptions } from '../src/format/conditional';
import { parseStyles } from '../src/formats/xlsx/styles';
import { parseWorksheet } from '../src/formats/xlsx/worksheet';
import { createSheet, parseCellRef, type Cell, type ConditionalFormat, type ConditionalRule, type Sheet, type Workbook } from '../src/model/sheet/types';
import { parseXml } from '../src/xml';

// --- helpers ---------------------------------------------------------------

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"';
const utf8 = (text: string) => new TextEncoder().encode(text);

type Value = number | string | boolean | null | { error: string } | { formula: string; value?: number };

const colName = (col: number) => String.fromCharCode(65 + col);

/** <sheetData> for rows of values; null leaves the cell out. */
function sheetData(rows: Value[][]): string {
  let out = '<sheetData>';
  rows.forEach((row, r) => {
    out += `<row r="${r + 1}">`;
    row.forEach((value, c) => {
      const ref = `${colName(c)}${r + 1}`;
      if (value === null) return;
      if (typeof value === 'number') out += `<c r="${ref}"><v>${value}</v></c>`;
      else if (typeof value === 'boolean') out += `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
      else if (typeof value === 'string') out += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${value}</t></is></c>`;
      else if ('error' in value) out += `<c r="${ref}" t="e"><v>${value.error}</v></c>`;
      else out += `<c r="${ref}"><f>${value.formula}</f>${value.value === undefined ? '' : `<v>${value.value}</v>`}</c>`;
    });
    out += '</row>';
  });
  return out + '</sheetData>';
}

const column = (values: Value[]) => values.map((value) => [value]);

/** Differential formats used by most tests: 0 red fill, 1 bold, 2 green fill + italic, 3 blue font. */
const DXFS =
  '<dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf>' +
  '<dxf><font><b/></font></dxf>' +
  '<dxf><font><i/></font><fill><patternFill patternType="solid"><fgColor rgb="FF123456"/><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>' +
  '<dxf><font><color rgb="FF0000FF"/></font></dxf>';

function book(body: string, dxfs = DXFS): { workbook: Workbook; sheet: Sheet } {
  const sheet = createSheet({ name: 'Data', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(utf8(`<worksheet ${NS} ${X14}>${body}</worksheet>`), [], sheet);
  const styles = parseStyles(parseXml(`<styleSheet ${NS}><dxfs>${dxfs}</dxfs></styleSheet>`), undefined);
  return { workbook: { sheets: [sheet], styles, sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 }, sheet };
}

const cf = (sqref: string, rules: string) => `<conditionalFormatting sqref="${sqref}">${rules}</conditionalFormatting>`;

/** Serial 45000 is Wednesday 15 March 2023. */
const NOW = 45000.5;

function withEngine(workbook: Workbook): ConditionalOptions {
  const engine = new CalcEngine(workbook, { now: () => NOW });
  return { engine: () => engine };
}

function at(formats: ConditionalFormats, ref: string): CellConditional | undefined {
  const cell = parseCellRef(ref)!;
  return formats.resolve(cell.row, cell.col);
}

/** The cells of a list that get a fill from a format (not from a colour scale). */
function filled(formats: ConditionalFormats, refs: string[]): string[] {
  return refs.filter((ref) => {
    const found = at(formats, ref);
    return !!found && !!found.style && found.style.fill !== undefined;
  });
}

const refs = (col: string, count: number) => Array.from({ length: count }, (_, i) => `${col}${i + 1}`);

function setValue(sheet: Sheet, ref: string, value: number | string): void {
  const { row, col } = parseCellRef(ref)!;
  const cell: Cell = typeof value === 'number' ? { kind: 'number', value, style: 0 } : { kind: 'string', value, style: 0 };
  let found = sheet.rows.get(row);
  if (!found) {
    found = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] };
    sheet.rows.set(row, found);
  }
  found.cells[col] = cell;
}

// --- reading ---------------------------------------------------------------

test('conditional: every rule type is read with its attributes', () => {
  const { sheet } = book(
    sheetData([[1]]) +
      cf(
        'A1:A10 C1:C5',
        '<cfRule type="cellIs" dxfId="0" priority="3" operator="between" stopIfTrue="1"><formula>1</formula><formula>$B$2</formula></cfRule>' +
          '<cfRule type="expression" dxfId="1" priority="1"><formula>$A1&gt;5</formula></cfRule>',
      ) +
      cf(
        'B1:B10',
        '<cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
          '<color rgb="FFF8696B"/><color theme="4" tint="0.5"/><color indexed="3"/></colorScale></cfRule>' +
          '<cfRule type="dataBar" priority="4"><dataBar minLength="5" maxLength="95" showValue="0"><cfvo type="num" val="0"/><cfvo type="formula" val="$Z$1"/><color rgb="FF638EC6"/></dataBar></cfRule>' +
          '<cfRule type="iconSet" priority="5"><iconSet iconSet="4Arrows" showValue="0" reverse="1"><cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="num" val="7" gte="0"/><cfvo type="percent" val="75"/></iconSet></cfRule>' +
          '<cfRule type="top10" dxfId="0" priority="6" rank="3" percent="1" bottom="1"/>' +
          '<cfRule type="aboveAverage" dxfId="0" priority="7" aboveAverage="0" equalAverage="1" stdDev="2"/>' +
          '<cfRule type="uniqueValues" dxfId="0" priority="8"/>' +
          '<cfRule type="duplicateValues" dxfId="0" priority="9"/>' +
          '<cfRule type="containsText" dxfId="0" priority="10" operator="containsText" text="ab"><formula>NOT(ISERROR(SEARCH("ab",B1)))</formula></cfRule>' +
          '<cfRule type="notContainsText" dxfId="0" priority="11" operator="notContains" text="ab"><formula>ISERROR(SEARCH("ab",B1))</formula></cfRule>' +
          '<cfRule type="beginsWith" dxfId="0" priority="12" operator="beginsWith" text="ab"><formula>LEFT(B1,2)="ab"</formula></cfRule>' +
          '<cfRule type="endsWith" dxfId="0" priority="13" operator="endsWith" text="ab"><formula>RIGHT(B1,2)="ab"</formula></cfRule>' +
          '<cfRule type="containsBlanks" dxfId="0" priority="14"><formula>LEN(TRIM(B1))=0</formula></cfRule>' +
          '<cfRule type="notContainsBlanks" dxfId="0" priority="15"><formula>LEN(TRIM(B1))&gt;0</formula></cfRule>' +
          '<cfRule type="containsErrors" dxfId="0" priority="16"><formula>ISERROR(B1)</formula></cfRule>' +
          '<cfRule type="notContainsErrors" dxfId="0" priority="17"><formula>NOT(ISERROR(B1))</formula></cfRule>' +
          '<cfRule type="timePeriod" dxfId="0" priority="18" timePeriod="last7Days"><formula>AND(TODAY()-FLOOR(B1,1)&lt;=6,FLOOR(B1,1)&lt;=TODAY())</formula></cfRule>',
      ),
  );
  const formats = sheet.conditionalFormats!;
  assert.equal(formats.length, 2);
  assert.deepEqual(formats[0].ranges, [
    { top: 0, left: 0, bottom: 9, right: 0 },
    { top: 0, left: 2, bottom: 4, right: 2 },
  ]);
  const base = { percent: false, bottom: false, aboveAverage: true, equalAverage: false, stopIfTrue: false };
  assert.deepEqual(formats[0].rules, [
    { ...base, type: 'cellIs', priority: 3, stopIfTrue: true, dxfId: 0, operator: 'between', formulas: ['1', '$B$2'] },
    { ...base, type: 'expression', priority: 1, dxfId: 1, formulas: ['$A1>5'] },
  ]);

  const rules = formats[1].rules;
  assert.deepEqual(
    rules.map((rule) => rule.type),
    ['colorScale', 'dataBar', 'iconSet', 'top10', 'aboveAverage', 'uniqueValues', 'duplicateValues', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod'],
  );
  assert.deepEqual(
    rules.map((rule) => rule.priority),
    [2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18],
  );
  assert.deepEqual(rules[0].colorScale, {
    values: [
      { type: 'min', gte: true },
      { type: 'percentile', value: '50', gte: true },
      { type: 'max', gte: true },
    ],
    colors: [{ rgb: 'FFF8696B' }, { theme: 4, tint: 0.5 }, { indexed: 3 }],
  });
  assert.deepEqual(rules[1].dataBar, {
    min: { type: 'num', value: '0', gte: true },
    max: { type: 'formula', value: '$Z$1', gte: true },
    color: { rgb: 'FF638EC6' },
    showValue: false,
    minLength: 5,
    maxLength: 95,
    extended: false,
    gradient: true,
    negativeFillSameAsPositive: false,
    negativeBorderSameAsPositive: true,
    axisPosition: 'automatic',
    direction: 'context',
  });
  assert.deepEqual(rules[2].iconSet, {
    set: '4Arrows',
    showValue: false,
    reverse: true,
    values: [
      { type: 'percent', value: '0', gte: true },
      { type: 'percent', value: '25', gte: true },
      { type: 'num', value: '7', gte: false },
      { type: 'percent', value: '75', gte: true },
    ],
  });
  assert.deepEqual({ rank: rules[3].rank, percent: rules[3].percent, bottom: rules[3].bottom }, { rank: 3, percent: true, bottom: true });
  assert.deepEqual({ above: rules[4].aboveAverage, equal: rules[4].equalAverage, stdDev: rules[4].stdDev }, { above: false, equal: true, stdDev: 2 });
  assert.equal(rules[7].text, 'ab');
  assert.equal(rules[7].operator, 'containsText');
  assert.deepEqual(rules[7].formulas, ['NOT(ISERROR(SEARCH("ab",B1)))']);
  assert.equal(rules[15].timePeriod, 'last7Days');
});

test('conditional: defaults, a prefixed document, whole columns and rows, and what is left out', () => {
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });
  parseWorksheet(
    utf8(
      '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData/>' +
        '<x:conditionalFormatting sqref="B:C 2:3 $D$4 nonsense A0"><x:cfRule type="iconSet"><x:iconSet><x:cfvo type="percent" val="0"/></x:iconSet></x:cfRule>' +
        '<x:cfRule type="futureKind" priority="x" dxfId="-1"><x:formula>1</x:formula></x:cfRule></x:conditionalFormatting>' +
        // No usable range, and no rules: neither is kept.
        '<x:conditionalFormatting sqref="nothing"><x:cfRule type="cellIs" priority="1"/></x:conditionalFormatting>' +
        '<x:conditionalFormatting sqref="A1"/>' +
        '<x:conditionalFormatting><x:cfRule type="cellIs" priority="1"/></x:conditionalFormatting>' +
        '</x:worksheet>',
    ),
    [],
    sheet,
  );
  const formats = sheet.conditionalFormats!;
  assert.equal(formats.length, 1);
  assert.deepEqual(formats[0].ranges, [
    { top: 0, left: 1, bottom: 1048575, right: 2 },
    { top: 1, left: 0, bottom: 2, right: 16383 },
    { top: 3, left: 3, bottom: 3, right: 3 },
  ]);
  const [icons, unknown] = formats[0].rules;
  assert.equal(icons.priority, Infinity);
  assert.equal(icons.iconSet!.set, '3TrafficLights1');
  assert.equal(icons.iconSet!.showValue, true);
  assert.equal(icons.iconSet!.reverse, false);
  assert.equal(unknown.type, 'futureKind');
  assert.equal(unknown.priority, Infinity);
  assert.equal(unknown.dxfId, undefined);

  // A sheet without conditional formats has no list at all.
  assert.equal(book(sheetData([[1]])).sheet.conditionalFormats, undefined);
});

const EXT_BAR =
  '<x14:cfRule type="dataBar" id="{AAAA-1}"><x14:dataBar minLength="0" maxLength="100" border="1" gradient="0" negativeBarBorderColorSameAsPositive="0" axisPosition="middle" direction="rightToLeft">' +
  '<x14:cfvo type="autoMin"/><x14:cfvo type="num"><xm:f>$Z$9</xm:f></x14:cfvo>' +
  '<x14:borderColor rgb="FF111111"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FF222222"/><x14:axisColor rgb="FF333333"/>' +
  '</x14:dataBar></x14:cfRule>';

test('conditional: the 2010 extension of a data bar is joined to its rule by id', () => {
  const { sheet } = book(
    sheetData([[1]]) +
      cf(
        'A1:A5',
        '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF00B050"/></dataBar>' +
          '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{AAAA-1}</x14:id></ext></extLst></cfRule>' +
          '<cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF0000FF"/></dataBar>' +
          '<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{NO-SUCH}</x14:id></ext></extLst></cfRule>',
      ) +
      '<extLst><ext uri="{other}"><x14:slicerList/></ext>' +
      '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>' +
      `<x14:conditionalFormatting>${EXT_BAR}<xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>` +
      // Rules that exist only in the extension list: an icon set with custom icons, and a rule with its own format.
      '<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="7" id="{BBBB-2}"><x14:iconSet iconSet="3Stars" custom="1" showValue="0">' +
      '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>5</xm:f></x14:cfvo><x14:cfvo type="num" gte="0"><xm:f>Other!$A$1</xm:f></x14:cfvo>' +
      '<x14:cfIcon iconSet="3Flags" iconId="2"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="5Boxes" iconId="4"/>' +
      '</x14:iconSet></x14:cfRule>' +
      '<x14:cfRule type="expression" priority="8" id="{CCCC-3}"><xm:f>Other!$A$1&gt;0</xm:f><x14:dxf><font><b/></font><fill><patternFill><bgColor rgb="FFABCDEF"/></patternFill></fill></x14:dxf></x14:cfRule>' +
      '<xm:sqref>C1:C3 E5</xm:sqref></x14:conditionalFormatting>' +
      '</x14:conditionalFormattings></ext></extLst>',
  );
  const formats = sheet.conditionalFormats!;
  assert.equal(formats.length, 2);
  const [linked, alone] = formats[0].rules;
  assert.deepEqual(linked.dataBar, {
    min: { type: 'autoMin', gte: true },
    max: { type: 'num', value: '$Z$9', gte: true },
    color: { rgb: 'FF00B050' },
    showValue: true,
    minLength: 0,
    maxLength: 100,
    extended: true,
    gradient: false,
    borderColor: { rgb: 'FF111111' },
    negativeFillColor: { rgb: 'FFFF0000' },
    negativeBorderColor: { rgb: 'FF222222' },
    negativeFillSameAsPositive: false,
    negativeBorderSameAsPositive: false,
    axisColor: { rgb: 'FF333333' },
    axisPosition: 'middle',
    direction: 'rightToLeft',
  });
  assert.equal(linked.priority, 1);
  // An id nothing answers to leaves the rule as Excel 2007 reads it.
  assert.equal(alone.dataBar!.extended, false);
  assert.equal(alone.extId, '{NO-SUCH}');

  assert.deepEqual(formats[1].ranges, [
    { top: 0, left: 2, bottom: 2, right: 2 },
    { top: 4, left: 4, bottom: 4, right: 4 },
  ]);
  const [icons, expression] = formats[1].rules;
  assert.equal(icons.extended, true);
  assert.equal(icons.priority, 7);
  assert.deepEqual(icons.iconSet, {
    set: '3Stars',
    showValue: false,
    reverse: false,
    values: [
      { type: 'percent', value: '0', gte: true },
      { type: 'num', value: '5', gte: true },
      { type: 'num', value: 'Other!$A$1', gte: false },
    ],
    custom: [
      { set: '3Flags', id: 2 },
      { set: 'NoIcons', id: 0 },
      { set: '5Boxes', id: 4 },
    ],
  });
  assert.deepEqual(expression.formulas, ['Other!$A$1>0']);
  assert.deepEqual(expression.dxf, { font: { bold: true }, fill: { type: 'pattern', pattern: 'solid', bg: { rgb: 'FFABCDEF' } } });
});

// --- differential formats ----------------------------------------------------

test('conditional: the colour of a differential fill', () => {
  const styles = parseStyles(
    parseXml(
      `<styleSheet ${NS}><dxfs>` +
        // Solid without a pattern type: the colour is bgColor, whatever fgColor says.
        '<dxf><fill><patternFill><fgColor rgb="FF111111"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>' +
        '<dxf><fill><patternFill patternType="solid"><bgColor theme="1"/></patternFill></fill></dxf>' +
        // Only a foreground colour: better that than nothing.
        '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/></patternFill></fill></dxf>' +
        '<dxf><fill><patternFill patternType="none"/></fill></dxf>' +
        '<dxf><fill><patternFill patternType="gray125"><fgColor rgb="FF0000FF"/><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>' +
        '<dxf><fill><gradientFill degree="90"><stop position="0"><color rgb="FFAA0000"/></stop><stop position="1"><color rgb="FF0000AA"/></stop></gradientFill></fill></dxf>' +
        '<dxf><fill><patternFill><bgColor auto="1"/></patternFill></fill></dxf>' +
        '<dxf><font><b val="0"/><i/><strike/><u val="double"/><color rgb="FF9C0006"/><sz val="20"/><name val="Arial"/></font>' +
        '<numFmt numFmtId="164" formatCode="0.00"/><border><left style="thin"><color rgb="FFFF0000"/></left><right/><top style="none"/><bottom style="double"/></border></dxf>' +
        '<dxf/>' +
        '</dxfs></styleSheet>',
    ),
    undefined,
  );
  styles.themeColors = ['FFFFFF', '1F2937'];
  const formats = styles.dxfFormats!;
  assert.equal(formats.length, 9);
  assert.deepEqual(
    formats.slice(0, 7).map((format) => differentialFillColor(format.fill, styles)),
    ['#FFC7CE', '#1F2937', '#00FF00', null, '#FFFF00', '#AA0000', undefined],
  );
  assert.equal(formats[0].fill!.type === 'pattern' && formats[0].fill!.pattern, 'solid');
  assert.deepEqual(formats[7], {
    font: { bold: false, italic: true, strike: true, underline: 'double', color: { rgb: 'FF9C0006' } },
    numFmt: { id: 164, code: '0.00' },
    border: { left: { style: 'thin', color: { rgb: 'FFFF0000' } }, bottom: { style: 'double' } },
  });
  assert.deepEqual(formats[8], {});
  assert.equal(differentialFillColor(undefined, styles), undefined);
});

// --- cell value rules --------------------------------------------------------

test('conditional: cellIs with every operator', () => {
  const values: Value[] = [1, 5, 9, 'abc', true, null, { error: '#N/A' }];
  const cells = refs('A', 7);
  const expect: [string, string, string[]][] = [
    ['equal', '<formula>5</formula>', ['A2']],
    ['notEqual', '<formula>5</formula>', ['A1', 'A3', 'A4', 'A5', 'A6']],
    // Text and booleans rank above every number, as in a formula.
    ['greaterThan', '<formula>5</formula>', ['A3', 'A4', 'A5']],
    ['greaterThanOrEqual', '<formula>5</formula>', ['A2', 'A3', 'A4', 'A5']],
    // UNSURE: an empty cell counts as 0 here, as it does in "=A6<5"; Excel is believed to do the same.
    ['lessThan', '<formula>5</formula>', ['A1', 'A6']],
    ['lessThanOrEqual', '<formula>5</formula>', ['A1', 'A2', 'A6']],
    ['between', '<formula>2</formula><formula>9</formula>', ['A2', 'A3']],
    // UNSURE: limits in the wrong order are taken as the same interval.
    ['between', '<formula>9</formula><formula>2</formula>', ['A2', 'A3']],
    ['notBetween', '<formula>2</formula><formula>9</formula>', ['A1', 'A4', 'A5', 'A6']],
    ['equal', '<formula>"ABC"</formula>', ['A4']],
    ['equal', '<formula>TRUE</formula>', ['A5']],
    ['noSuchOperator', '<formula>5</formula>', []],
    ['between', '<formula>2</formula>', []],
    ['equal', '', []],
  ];
  for (const [operator, formulas, wanted] of expect) {
    const { workbook, sheet } = book(sheetData(column(values)) + cf('A1:A7', `<cfRule type="cellIs" dxfId="0" priority="1" operator="${operator}">${formulas}</cfRule>`));
    // Numbers and quoted text need no engine.
    assert.deepEqual(filled(new ConditionalFormats(workbook, sheet), cells), wanted, `${operator} ${formulas}`);
  }
});

test('conditional: formulas move with the cell; without an engine they do not apply', () => {
  const rows: Value[][] = [
    [1, 10, 0],
    [6, 2, 0],
    [7, 9, 0],
    [3, 1, 0],
  ];
  const body =
    sheetData(rows) +
    // "$A1>5": column fixed, row moves. Applied to A:C, every cell of rows 2 and 3 matches.
    cf('A1:C4', '<cfRule type="expression" dxfId="0" priority="1"><formula>$A1&gt;5</formula></cfRule>') +
    // "A1>B$1": compared with the cell to the right in row 1.
    // UNSURE: formulas are taken as written for the top-left cell of the FIRST range of the list (the writer of
    // this project assumes the same); for a list whose first range is not its top-left-most one Excel may differ.
    cf('E1:F4', '<cfRule type="expression" dxfId="1" priority="2"><formula>A1&gt;B$1</formula></cfRule>') +
    // A cellIs operand that is a formula.
    cf('A1:A4', '<cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>$B1</formula></cfRule>');
  const { workbook, sheet } = book(body);
  const formats = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  const all = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'A4', 'B4', 'C4'];
  assert.deepEqual(filled(formats, all), ['A2', 'B2', 'C2', 'A3', 'B3', 'C3', 'A4']);
  // A2 gets the red fill of the first rule; A4 only matches the third.
  assert.equal(at(formats, 'A2')!.style!.fill, '#FF0000');
  assert.deepEqual(at(formats, 'A4')!.style, { italic: true, fill: '#00FF00' });
  // E1 stands for A1 > B$1 (1 > 10), E2 for A2 > B$1 (6 > 10), F3 for B3 > C$1 (9 > 0).
  const bold = (ref: string) => !!at(formats, ref) && at(formats, ref)!.style!.bold === true;
  assert.deepEqual(
    ['E1', 'E2', 'E3', 'E4', 'F1', 'F2', 'F3', 'F4'].filter(bold),
    ['F1', 'F2', 'F3', 'F4'],
  );

  const plain = new ConditionalFormats(workbook, sheet);
  assert.deepEqual(filled(plain, all), []);
  assert.equal(at(plain, 'F1'), undefined);
  // An engine that cannot be had is the same as none.
  const broken = new ConditionalFormats(workbook, sheet, {
    engine: () => {
      throw new Error('no engine');
    },
  });
  assert.deepEqual(filled(broken, all), []);
});

test('conditional: formulas that cannot be parsed or calculated do not apply and do not throw', () => {
  const bad = ['1+', 'NOSUCHFUNCTION(A1)', '"text"&"more"', '#REF!&gt;1', 'A1/0', 'XFD1048576&gt;0', ''];
  const body = sheetData([[1, 2]]) + bad.map((text, i) => cf('A1:B1', `<cfRule type="expression" dxfId="0" priority="${i + 1}"><formula>${text}</formula></cfRule>`)).join('');
  const { workbook, sheet } = book(body);
  const formats = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  assert.equal(at(formats, 'A1'), undefined);
  // In B1 the last reference moves off the sheet.
  assert.equal(at(formats, 'B1'), undefined);
  // Asked again: the answer is kept, the broken formula is not parsed again.
  assert.equal(at(formats, 'A1'), undefined);
});

test('conditional: priority, the first rule to state a property, and stopIfTrue', () => {
  const body =
    sheetData(column([10, 20, 30])) +
    // Listed first but applied last.
    cf('A1:A3', '<cfRule type="cellIs" dxfId="0" priority="5" operator="greaterThan"><formula>0</formula></cfRule>') +
    cf(
      'A1:A3',
      '<cfRule type="cellIs" dxfId="3" priority="4" operator="greaterThan"><formula>0</formula></cfRule>' +
        '<cfRule type="cellIs" dxfId="2" priority="2" operator="greaterThan"><formula>15</formula></cfRule>' +
        '<cfRule type="cellIs" dxfId="1" priority="3" operator="greaterThan" stopIfTrue="1"><formula>25</formula></cfRule>',
    );
  const { workbook, sheet } = book(body);
  const formats = new ConditionalFormats(workbook, sheet);
  // 10: only the two rules for "> 0": blue font, then the red fill.
  assert.deepEqual(at(formats, 'A1')!.style, { color: '#0000FF', fill: '#FF0000' });
  // 20: the green fill comes first and keeps the fill; red is too late.
  assert.deepEqual(at(formats, 'A2')!.style, { italic: true, fill: '#00FF00', color: '#0000FF' });
  // 30: green, bold, and then the bold rule stops everything after it.
  assert.deepEqual(at(formats, 'A3')!.style, { italic: true, fill: '#00FF00', bold: true });
  // Cells with the same rules share one object, which is what the renderer keys its cache on.
  setValue(sheet, 'A1', 22);
  formats.invalidate();
  assert.equal(at(formats, 'A1')!.style, at(formats, 'A2')!.style);

  // A rule that matches without a format (its dxfId points nowhere) still stops.
  const stop = book(
    sheetData([[1]]) +
      cf('A1', '<cfRule type="cellIs" dxfId="99" priority="1" operator="equal" stopIfTrue="1"><formula>1</formula></cfRule><cfRule type="cellIs" dxfId="0" priority="2" operator="equal"><formula>1</formula></cfRule>'),
  );
  assert.equal(at(new ConditionalFormats(stop.workbook, stop.sheet), 'A1'), undefined);
});

// --- rules over all cells of the range ---------------------------------------

function rule(values: Value[], attributes: string, count = values.length): string[] {
  const { workbook, sheet } = book(sheetData(column(values)) + cf(`A1:A${count}`, `<cfRule dxfId="0" priority="1" ${attributes}/>`));
  return filled(new ConditionalFormats(workbook, sheet), refs('A', count));
}

test('conditional: top and bottom', () => {
  const values: Value[] = [5, 1, 9, 9, 3, 'x', null, 7, { error: '#DIV/0!' }, 2];
  assert.deepEqual(rule(values, 'type="top10" rank="2"'), ['A3', 'A4']);
  // Ties at the limit all count.
  assert.deepEqual(rule(values, 'type="top10" rank="1"'), ['A3', 'A4']);
  assert.deepEqual(rule(values, 'type="top10" rank="3"'), ['A3', 'A4', 'A8']);
  assert.deepEqual(rule(values, 'type="top10" rank="2" bottom="1"'), ['A2', 'A10']);
  // Seven numbers: 30% is two of them. UNSURE: Excel is believed to round the count down, with at least one.
  assert.deepEqual(rule(values, 'type="top10" rank="30" percent="1"'), ['A3', 'A4']);
  assert.deepEqual(rule(values, 'type="top10" rank="1" percent="1" bottom="1"'), ['A2']);
  // The default rank is 10: everything numeric.
  assert.deepEqual(rule(values, 'type="top10"'), ['A1', 'A2', 'A3', 'A4', 'A5', 'A8', 'A10']);
  assert.deepEqual(rule(['a', null], 'type="top10" rank="1"'), []);
});

test('conditional: above and below average', () => {
  // Mean 5; sample deviation of 2, 4, 4, 4, 5, 5, 7, 9 is 2.138...
  const values: Value[] = [2, 4, 4, 4, 5, 5, 7, 9, 'text'];
  assert.deepEqual(rule(values, 'type="aboveAverage"'), ['A7', 'A8']);
  assert.deepEqual(rule(values, 'type="aboveAverage" equalAverage="1"'), ['A5', 'A6', 'A7', 'A8']);
  assert.deepEqual(rule(values, 'type="aboveAverage" aboveAverage="0"'), ['A1', 'A2', 'A3', 'A4']);
  assert.deepEqual(rule(values, 'type="aboveAverage" aboveAverage="0" equalAverage="1"'), ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']);
  // UNSURE: the deviation is taken as the sample deviation (STDEV), as Apache POI does.
  assert.deepEqual(rule(values, 'type="aboveAverage" stdDev="1"'), ['A8']);
  assert.deepEqual(rule(values, 'type="aboveAverage" aboveAverage="0" stdDev="1"'), ['A1']);
  assert.deepEqual(rule(values, 'type="aboveAverage" stdDev="3"'), []);
});

test('conditional: unique and repeated values', () => {
  const values: Value[] = [1, 'Ali', 'ALI', 2, 1, '1', null, '', true, true, { error: '#N/A' }, { error: '#N/A' }, 3];
  // Text compares without case; the text "1" counts as the number 1 (see conditional-fixes); empty cells and errors take no part.
  // UNSURE: booleans are counted like any other value here.
  assert.deepEqual(rule(values, 'type="duplicateValues"'), ['A1', 'A2', 'A3', 'A5', 'A6', 'A9', 'A10']);
  assert.deepEqual(rule(values, 'type="uniqueValues"'), ['A4', 'A13']);
});

test('conditional: text rules', () => {
  const values: Value[] = ['Hello world', 'WORLD cup', 'low', 12345, null, { error: '#N/A' }, 'a*b', 'x'];
  const cells = refs('A', 8);
  const run = (type: string, text: string) => rule(values, `type="${type}" text="${text}"`);
  assert.deepEqual(run('containsText', 'world'), ['A1', 'A2']);
  assert.deepEqual(run('containsText', '234'), ['A4']);
  // Wildcards as in SEARCH.
  assert.deepEqual(run('containsText', 'w?r'), ['A1', 'A2']);
  assert.deepEqual(run('containsText', 'h*d'), ['A1']);
  assert.deepEqual(run('containsText', 'a~*b'), ['A7']);
  assert.deepEqual(run('containsText', ''), ['A1', 'A2', 'A3', 'A4', 'A5', 'A7', 'A8']);
  // An error cell does not contain the text, so "does not contain" applies to it.
  assert.deepEqual(run('notContainsText', 'o'), ['A4', 'A5', 'A6', 'A7', 'A8']);
  assert.deepEqual(run('beginsWith', 'WOR'), ['A2']);
  assert.deepEqual(run('beginsWith', '12'), ['A4']);
  assert.deepEqual(run('endsWith', 'LD'), ['A1']);
  assert.deepEqual(run('endsWith', 'b'), ['A7']);

  // Without the text attribute the rule is its formula, which needs the engine.
  const { workbook, sheet } = book(sheetData(column(values)) + cf('A1:A8', '<cfRule type="containsText" dxfId="0" priority="1"><formula>NOT(ISERROR(SEARCH("world",A1)))</formula></cfRule>'));
  assert.deepEqual(filled(new ConditionalFormats(workbook, sheet, withEngine(workbook)), cells), ['A1', 'A2']);
  assert.deepEqual(filled(new ConditionalFormats(workbook, sheet), cells), []);
});

test('conditional: blanks and errors', () => {
  const values: Value[] = [1, '', '   ', null, 'x', { error: '#REF!' }, 0, false];
  assert.deepEqual(rule(values, 'type="containsBlanks"'), ['A2', 'A3', 'A4']);
  assert.deepEqual(rule(values, 'type="notContainsBlanks"'), ['A1', 'A5', 'A7', 'A8']);
  assert.deepEqual(rule(values, 'type="containsErrors"'), ['A6']);
  assert.deepEqual(rule(values, 'type="notContainsErrors"'), ['A1', 'A2', 'A3', 'A4', 'A5', 'A7', 'A8']);
});

test('conditional: time periods', () => {
  // Today is Wednesday 15 March 2023 (45000); its week runs from Sunday 12 (44997) to Saturday 18 (45003).
  const days = [45000, 45000.99, 44999, 45001, 44994, 44995, 44996, 44997, 45003, 45004, 45010, 45011, 44990, 44989, 44985, 44986, 44958, 44957, 45016, 45017, 45046, 45047];
  const values: Value[] = [...days, 'x', null, true];
  const { workbook, sheet } = book(sheetData(column(values)));
  const run = (period: string, options: ConditionalOptions = { now: () => NOW }) => {
    sheet.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 30, right: 0 }], rules: [{ type: 'timePeriod', priority: 1, stopIfTrue: false, dxfId: 0, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false, timePeriod: period }] }];
    const formats = new ConditionalFormats(workbook, sheet, options);
    return days.filter((_, i) => !!formats.resolve(i, 0));
  };
  assert.deepEqual(run('today'), [45000, 45000.99]);
  assert.deepEqual(run('yesterday'), [44999]);
  assert.deepEqual(run('tomorrow'), [45001]);
  assert.deepEqual(run('last7Days'), [45000, 45000.99, 44999, 44994, 44995, 44996, 44997]);
  assert.deepEqual(run('thisWeek'), [45000, 45000.99, 44999, 45001, 44997, 45003]);
  assert.deepEqual(run('lastWeek'), [44994, 44995, 44996, 44990]);
  assert.deepEqual(run('nextWeek'), [45004, 45010]);
  // March 2023 is 44986..45016.
  assert.deepEqual(run('thisMonth'), [45000, 45000.99, 44999, 45001, 44994, 44995, 44996, 44997, 45003, 45004, 45010, 45011, 44990, 44989, 44986, 45016]);
  assert.deepEqual(run('lastMonth'), [44985, 44958]);
  assert.deepEqual(run('nextMonth'), [45017, 45046]);
  assert.deepEqual(run('someday'), []);
  // The date comes from the engine's clock when no other is given.
  assert.deepEqual(run('today', withEngine(workbook)), [45000, 45000.99]);
  // And from the system clock without either: whatever today is, it is not in March 2023.
  assert.deepEqual(run('today', {}), []);

  // In the 1904 date system the same days are 1462 lower.
  workbook.date1904 = true;
  assert.deepEqual(run('yesterday', { now: () => NOW - 1462 }), []);
  setValue(sheet, 'A1', 44999 - 1462);
  sheet.conditionalFormats![0].rules[0].timePeriod = 'yesterday';
  assert.ok(new ConditionalFormats(workbook, sheet, { now: () => NOW - 1462 }).resolve(0, 0));
});

// --- colour scales, data bars, icons -----------------------------------------

test('conditional: colour scales', () => {
  const values: Value[] = [0, 25, 50, 100, 'text', { error: '#N/A' }, null, true];
  const scale = (cfvos: string, colors: string) => `<cfRule type="colorScale" priority="1"><colorScale>${cfvos}${colors}</colorScale></cfRule>`;
  const two = book(sheetData(column(values)) + cf('A1:A8', scale('<cfvo type="min"/><cfvo type="max"/>', '<color rgb="FFFF0000"/><color rgb="FF0000FF"/>')));
  const formats = new ConditionalFormats(two.workbook, two.sheet);
  assert.deepEqual(
    refs('A', 8).map((ref) => at(formats, ref)?.fill),
    ['#FF0000', '#BF0040', '#800080', '#0000FF', undefined, undefined, undefined, undefined],
  );
  assert.equal(at(formats, 'A1')!.style, undefined);

  // Three colours; the middle one sits at the 50th percentile, which is 37.5 here.
  const three = book(
    sheetData(column([0, 25, 50, 100])) + cf('A1:A4', scale('<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>', '<color rgb="FF000000"/><color rgb="FF646464"/><color rgb="FFC8C8C8"/>')),
  );
  const grey = new ConditionalFormats(three.workbook, three.sheet);
  // 25 is two thirds of the way from 0 to 37.5; 50 is a fifth of the way from 37.5 to 100.
  assert.deepEqual(
    refs('A', 4).map((ref) => at(grey, ref)!.fill),
    ['#000000', '#434343', '#787878', '#C8C8C8'],
  );

  // num, percent and formula thresholds; values outside take the end colours.
  const fixed = book(
    sheetData([[0, 40], [25], [50], [100]]) + cf('A1:A4', scale('<cfvo type="num" val="20"/><cfvo type="percent" val="50"/><cfvo type="formula" val="$B$1*2"/>', '<color rgb="FF000000"/><color rgb="FF646464"/><color rgb="FFC8C8C8"/>')),
  );
  const mixed = new ConditionalFormats(fixed.workbook, fixed.sheet, withEngine(fixed.workbook));
  // Thresholds 20, 50 and 80.
  assert.deepEqual(
    refs('A', 4).map((ref) => at(mixed, ref)!.fill),
    ['#000000', '#111111', '#646464', '#C8C8C8'],
  );
  // The formula threshold cannot be had without an engine: the rule is off.
  assert.equal(at(new ConditionalFormats(fixed.workbook, fixed.sheet), 'A1'), undefined);
});

test('conditional: a format with a fill and a colour scale - the earlier one fills', () => {
  const scale = (priority: number) => `<cfRule type="colorScale" priority="${priority}"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>`;
  const red = (priority: number) => `<cfRule type="cellIs" dxfId="0" priority="${priority}" operator="greaterThan"><formula>5</formula></cfRule>`;
  const first = book(sheetData(column([0, 10])) + cf('A1:A2', scale(1) + red(2)));
  const scaleFirst = at(new ConditionalFormats(first.workbook, first.sheet), 'A2')!;
  assert.equal(scaleFirst.fill, '#FFFFFF');
  const second = book(sheetData(column([0, 10])) + cf('A1:A2', scale(2) + red(1)));
  const formatFirst = at(new ConditionalFormats(second.workbook, second.sheet), 'A2')!;
  assert.equal(formatFirst.fill, undefined);
  assert.equal(formatFirst.style!.fill, '#FF0000');
});

test('conditional: data bars as Excel 2007 draws them', () => {
  const bar = '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule>';
  const { workbook, sheet } = book(sheetData(column([-10, 0, 40, 90, 'x', null])) + cf('A1:A6', bar));
  const formats = new ConditionalFormats(workbook, sheet);
  const bars = refs('A', 6).map((ref) => at(formats, ref)?.bar);
  // From 10% of the cell for the smallest value to 90% for the largest.
  assert.deepEqual(
    bars.map((found) => found && Math.round(found.to * 1000) / 1000),
    [0.1, 0.18, 0.5, 0.9, undefined, undefined],
  );
  assert.deepEqual(bars[0], { from: 0, to: 0.1, color: '#638EC6', gradient: true, negative: false, rightToLeft: false });
  assert.equal(at(formats, 'A1')!.hideValue, undefined);

  // Fixed limits and no value shown.
  const fixed = book(sheetData(column([-5, 50, 500])) + cf('A1:A3', '<cfRule type="dataBar" priority="1"><dataBar showValue="0" minLength="0" maxLength="100"><cfvo type="num" val="0"/><cfvo type="num" val="100"/></dataBar></cfRule>'));
  const clipped = new ConditionalFormats(fixed.workbook, fixed.sheet);
  assert.deepEqual(
    refs('A', 3).map((ref) => at(clipped, ref)!.bar!.to),
    [0, 0.5, 1],
  );
  assert.equal(at(clipped, 'A2')!.hideValue, true);
  // No colour stated: the default blue.
  assert.equal(at(clipped, 'A2')!.bar!.color, '#638EC6');
});

test('conditional: data bars with the 2010 extension - negative values and the axis', () => {
  const main = '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF00B050"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{1}</x14:id></ext></extLst></cfRule>';
  const ext = (attributes: string, extra = '') =>
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting>' +
    `<x14:cfRule type="dataBar" id="{1}"><x14:dataBar minLength="0" maxLength="100" ${attributes}><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>` +
    `<x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FF990000"/><x14:axisColor rgb="FF123456"/>${extra}</x14:dataBar></x14:cfRule>` +
    '<xm:sqref>A1:A4</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>';
  const read = (values: number[], attributes: string, extra = '') => {
    const { workbook, sheet } = book(sheetData(column(values)) + cf('A1:A4', main) + ext(attributes, extra));
    const formats = new ConditionalFormats(workbook, sheet);
    return refs('A', values.length).map((ref) => at(formats, ref)!.bar!);
  };

  // -50..150: zero is a quarter of the way across.
  // UNSURE: "automatic" limits are taken as min(0, smallest) and max(0, largest); Excel's own choice of the axis
  // position (it weighs the lengths of the two sides) may differ by some pixels.
  const auto = read([-50, 0, 75, 150], 'gradient="0" border="1" negativeBarBorderColorSameAsPositive="0"', '<x14:borderColor rgb="FF005500"/>');
  assert.deepEqual(auto[0], { from: 0, to: 0.25, color: '#FF0000', gradient: false, negative: true, rightToLeft: false, border: '#990000', axis: 0.25, axisColor: '#123456' });
  assert.deepEqual(auto[1], { from: 0.25, to: 0.25, color: '#00B050', gradient: false, negative: false, rightToLeft: false, border: '#005500', axis: 0.25, axisColor: '#123456' });
  assert.deepEqual([auto[2].from, auto[2].to], [0.25, 0.625]);
  assert.deepEqual([auto[3].from, auto[3].to], [0.25, 1]);

  // Only positive values: the automatic minimum is 0, there is no axis.
  const positive = read([50, 100], '');
  assert.deepEqual(positive[0], { from: 0, to: 0.5, color: '#00B050', gradient: true, negative: false, rightToLeft: false });
  assert.equal(positive[1].to, 1);

  // Only negative values: the axis is at the far end.
  const negative = read([-50, -100], '');
  assert.deepEqual([negative[0].from, negative[0].to, negative[0].axis], [0.5, 1, 1]);
  assert.deepEqual([negative[1].from, negative[1].to], [0, 1]);

  const middle = read([-50, 100], 'axisPosition="middle" negativeBarColorSameAsPositive="1" direction="rightToLeft"');
  assert.deepEqual(middle[0], { from: 0.25, to: 0.5, color: '#00B050', gradient: true, negative: true, rightToLeft: true, axis: 0.5, axisColor: '#123456' });
  assert.deepEqual([middle[1].from, middle[1].to], [0.5, 1]);

  // No axis: negative values are placed between the limits like any other.
  const none = read([-100, 0, 100], 'axisPosition="none"');
  assert.deepEqual(
    none.map((found) => [found.from, found.to, found.axis, found.negative]),
    [
      [0, 0, undefined, false],
      [0, 0.5, undefined, false],
      [0, 1, undefined, false],
    ],
  );
});

test('conditional: icon sets', () => {
  const icons = (attributes: string, cfvos: string) => `<cfRule type="iconSet" priority="1"><iconSet ${attributes}>${cfvos}</iconSet></cfRule>`;
  const thirds = '<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/>';
  const values: Value[] = [0, 33, 50, 67, 100, 'x', null, { error: '#N/A' }];
  const read = (rule: string, options?: (workbook: Workbook) => ConditionalOptions) => {
    const { workbook, sheet } = book(sheetData(column(values)) + cf('A1:A8', rule));
    const formats = new ConditionalFormats(workbook, sheet, options ? options(workbook) : {});
    return refs('A', 8).map((ref) => at(formats, ref));
  };
  const indexes = (found: (CellConditional | undefined)[]) => found.map((cell) => (cell && cell.icon ? cell.icon.index : undefined));

  const arrows = read(icons('iconSet="3Arrows"', thirds));
  assert.deepEqual(indexes(arrows), [0, 1, 1, 2, 2, undefined, undefined, undefined]);
  assert.deepEqual(arrows[0]!.icon, { set: '3Arrows', index: 0 });
  assert.equal(arrows[0]!.hideValue, undefined);
  assert.equal(arrows[5], undefined);

  // The default set, reversed, without values.
  const lights = read(icons('reverse="1" showValue="0"', thirds));
  assert.deepEqual(indexes(lights), [2, 1, 1, 0, 0, undefined, undefined, undefined]);
  assert.deepEqual(lights[0], { icon: { set: '3TrafficLights1', index: 2 }, hideValue: true });

  // gte="0": the threshold itself still belongs to the icon below.
  const strict = read(icons('iconSet="3Flags"', '<cfvo type="percent" val="0"/><cfvo type="num" val="33" gte="0"/><cfvo type="num" val="67" gte="0"/>'));
  assert.deepEqual(indexes(strict), [0, 0, 1, 1, 2, undefined, undefined, undefined]);

  const five = read(icons('iconSet="5Quarters"', '<cfvo type="percent" val="0"/><cfvo type="percent" val="20"/><cfvo type="percent" val="40"/><cfvo type="percent" val="60"/><cfvo type="percent" val="80"/>'));
  assert.deepEqual(indexes(five), [0, 1, 2, 3, 4, undefined, undefined, undefined]);
  const four = read(icons('iconSet="4Rating"', '<cfvo type="percent" val="0"/><cfvo type="percentile" val="25"/><cfvo type="percentile" val="50"/><cfvo type="formula" val="60+7"/>'), withEngine);
  // Percentiles of 0, 33, 50, 67, 100: the 25th is 33, the 50th is 50.
  assert.deepEqual(indexes(four), [0, 1, 2, 3, 3, undefined, undefined, undefined]);

  // Too few thresholds for the set: the rule is off. An unknown set of three is drawn as traffic lights.
  assert.deepEqual(indexes(read(icons('iconSet="5Arrows"', thirds))), new Array(8).fill(undefined));
  assert.equal(read(icons('iconSet="3Unknown"', thirds))[4]!.icon!.set, '3TrafficLights1');
  assert.equal(read(icons('iconSet="nonsense"', thirds))[4]!.icon!.set, '3TrafficLights1');
});

test('conditional: custom icons of the 2010 extension', () => {
  const { workbook, sheet } = book(
    sheetData(column([1, 5, 9])) +
      '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings><x14:conditionalFormatting>' +
      '<x14:cfRule type="iconSet" priority="1" id="{2}"><x14:iconSet iconSet="3Arrows" custom="1" showValue="0">' +
      '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>4</xm:f></x14:cfvo><x14:cfvo type="num"><xm:f>8</xm:f></x14:cfvo>' +
      '<x14:cfIcon iconSet="3Stars" iconId="2"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="5Boxes" iconId="9"/>' +
      '</x14:iconSet></x14:cfRule><xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>',
  );
  const formats = new ConditionalFormats(workbook, sheet);
  assert.deepEqual(at(formats, 'A1'), { icon: { set: '3Stars', index: 2 }, hideValue: true });
  // "No icon", and an icon the set does not have: the value stays hidden, nothing is drawn.
  assert.deepEqual(at(formats, 'A2'), { hideValue: true });
  assert.deepEqual(at(formats, 'A3'), { hideValue: true });
});

test('conditional: every built-in icon set has the icons its name promises', () => {
  for (const [name, icons] of ICON_SETS) {
    assert.equal(icons.length, Number(name[0]), name);
    for (let i = 0; i < icons.length; i++) {
      const found = iconGlyph({ set: name, index: i })!;
      assert.match(found.color, /^#[0-9A-F]{6}$/);
      assert.equal(typeof found.shape, 'string');
    }
  }
  assert.equal(ICON_SETS.size, 20);
  assert.equal(iconGlyph({ set: 'toString', index: 0 }), undefined);
  assert.equal(iconGlyph({ set: '3Arrows', index: 3 }), undefined);
  assert.equal(iconGlyph({ set: '3Arrows', index: -1 }), undefined);
});

// --- caching -----------------------------------------------------------------

/** A Map that counts how often it is walked and read. */
class CountingRows<K, V> extends Map<K, V> {
  walks = 0;
  reads = 0;
  override forEach(callback: (value: V, key: K, map: Map<K, V>) => void): void {
    this.walks++;
    super.forEach(callback);
  }
  override get(key: K): V | undefined {
    this.reads++;
    return super.get(key);
  }
}

test('conditional: what a rule knows about its range is worked out once and again after invalidate()', () => {
  const count = 5000;
  const values = Array.from({ length: count }, (_, i) => i);
  const { workbook, sheet } = book(
    sheetData(column(values)) +
      // Whole column and whole sheet: only the rows that exist are looked at.
      cf('A:A', '<cfRule type="top10" dxfId="0" priority="1" rank="1"/><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FF000000"/><color rgb="FFFFFFFF"/></colorScale></cfRule>') +
      cf('A1:XFD1048576', '<cfRule type="duplicateValues" dxfId="1" priority="3"/>'),
  );
  const rows = new CountingRows(sheet.rows);
  sheet.rows = rows;
  const formats = new ConditionalFormats(workbook, sheet);

  for (let r = 0; r < count; r++) formats.resolve(r, 0);
  // One walk for the numbers of column A (shared by both of its rules is not required: at most one each), one for the counts.
  assert.ok(rows.walks >= 2 && rows.walks <= 3, `${rows.walks} walks`);
  const walks = rows.walks;
  const reads = rows.reads;
  assert.equal(at(formats, `A${count}`)!.style!.fill, '#FF0000');
  assert.equal(at(formats, `A${count}`)!.fill, undefined);
  assert.equal(at(formats, 'A1')!.fill, '#000000');

  // Asking again touches nothing: neither the range nor the cells.
  for (let r = 0; r < count; r++) formats.resolve(r, 0);
  assert.equal(rows.walks, walks);
  assert.equal(rows.reads, reads);

  // A change is not seen until invalidate() ...
  setValue(sheet, 'A1', 99999);
  setValue(sheet, 'A2', 4999);
  assert.equal(at(formats, 'A1')!.fill, '#000000');
  assert.equal(at(formats, `A${count}`)!.style!.fill, '#FF0000');
  // ... and then everything follows: A1 is the top value, the old top is no longer, 4999 is there twice.
  formats.invalidate();
  assert.equal(at(formats, 'A1')!.style!.fill, '#FF0000');
  assert.deepEqual(at(formats, `A${count}`)!.style, { bold: true });
  assert.deepEqual(at(formats, 'A2')!.style, { bold: true });
  assert.equal(at(formats, 'A2')!.fill, '#0D0D0D');
  assert.ok(rows.walks > walks);

  // Cells outside every range cost nothing and get nothing.
  const outside = new ConditionalFormats(book(sheetData([[1]]) + cf('B2:C3', '<cfRule type="containsBlanks" dxfId="0" priority="1"/>')).workbook, sheet);
  assert.equal(outside.resolve(500, 5), undefined);
});

test('conditional: 100,000 rows under whole-column rules stay fast', () => {
  const sheet = createSheet({ name: 'Big', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/x' });
  for (let r = 0; r < 100000; r++) sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: (r * 7919) % 1000, style: 0 }] });
  sheet.rowCount = 100000;
  sheet.colCount = 1;
  const whole = [{ top: 0, left: 0, bottom: 1048575, right: 0 }];
  const base = { stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false };
  sheet.conditionalFormats = [
    {
      ranges: whole,
      rules: [
        { ...base, type: 'colorScale', priority: 1, colorScale: { values: [{ type: 'min', gte: true }, { type: 'percentile', value: '50', gte: true }, { type: 'max', gte: true }], colors: [{ rgb: 'FF0000' }, { rgb: 'FFFF00' }, { rgb: '00FF00' }] } },
        { ...base, type: 'top10', priority: 2, rank: 10, percent: true, dxfId: 0 },
        { ...base, type: 'aboveAverage', priority: 3, dxfId: 0 },
        { ...base, type: 'duplicateValues', priority: 4, dxfId: 0 },
      ],
    },
  ];
  const { workbook } = book(sheetData([[1]]));
  workbook.sheets = [sheet];
  const formats = new ConditionalFormats(workbook, sheet);
  let started = Date.now();
  assert.equal(formats.resolve(0, 0)!.fill, '#FF0000');
  const first = Date.now() - started;
  started = Date.now();
  for (let r = 0; r < 100000; r++) formats.resolve(r, 0);
  const all = Date.now() - started;
  console.log(`conditional: 100,000 rows x 4 whole-column rules: first cell ${first} ms, every cell ${all} ms`);
  assert.ok(first < 2000, `first cell took ${first} ms`);
  assert.ok(all < 3000, `every cell took ${all} ms`);
  // An empty cell far below the data, inside the range: nothing to show, nothing to scan.
  assert.equal(formats.resolve(900000, 0), undefined);
});

// --- edits -------------------------------------------------------------------

test('conditional: ranges and formulas follow rows and columns inserted since the file was read', () => {
  const body =
    sheetData([
      [1, 5],
      [9, 5],
      [3, 5],
    ]) +
    cf('A1:A3', '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;$B$2</formula></cfRule>') +
    cf('A2', '<cfRule type="cellIs" dxfId="1" priority="2" operator="equal"><formula>9</formula></cfRule>');
  const { workbook, sheet } = book(body);
  const formats = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  assert.deepEqual(filled(formats, ['A1', 'A2', 'A3', 'A4']), ['A2']);
  assert.equal(at(formats, 'A2')!.style!.bold, true);

  // Two rows inserted above row 2, as the insert operation leaves the model: cells moved, the edit logged.
  const moved = new Map(sheet.rows);
  sheet.rows.clear();
  moved.forEach((row, r) => sheet.rows.set(r >= 1 ? r + 2 : r, row));
  sheet.structuralEdits = [{ kind: 'insert', axis: 'row', at: 1, count: 2 }];
  formats.invalidate();
  // The range is A1:A5 now, the formula A1>$B$4; the single cell rule moved to A4.
  assert.deepEqual(filled(formats, ['A1', 'A2', 'A3', 'A4', 'A5', 'A6']), ['A4']);
  assert.equal(at(formats, 'A4')!.style!.bold, true);
  assert.equal(at(formats, 'A2'), undefined);
  // The model itself is untouched: the writer works from the source element.
  assert.deepEqual(sheet.conditionalFormats![0].ranges, [{ top: 0, left: 0, bottom: 2, right: 0 }]);
  assert.deepEqual(sheet.conditionalFormats![0].rules[0].formulas, ['A1>$B$2']);

  // Undo: the log is the earlier one again.
  sheet.rows.clear();
  moved.forEach((row, r) => sheet.rows.set(r, row));
  sheet.structuralEdits = undefined;
  formats.invalidate();
  assert.deepEqual(filled(formats, ['A1', 'A2', 'A3', 'A4']), ['A2']);

  // Deleting every row of a range takes its rules away.
  sheet.structuralEdits = [{ kind: 'delete', axis: 'row', at: 0, count: 3 }];
  sheet.rows.clear();
  formats.invalidate();
  assert.equal(at(formats, 'A1'), undefined);
});

// --- formulas within a time budget -------------------------------------------

test('conditional: formulas beyond the time of a frame are left for the next one', () => {
  const { workbook, sheet } = book(sheetData(column([1, 2, 3, 4, 5, 6])) + cf('A1:A6', '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;0</formula></cfRule>'));
  let now = 0;
  // Every look at the clock takes 10 ms; a frame has 25.
  const formats = new ConditionalFormats(workbook, sheet, { ...withEngine(workbook), budget: 25, clock: () => (now += 10) });
  const cells = refs('A', 6);
  let frames = 0;
  let done: string[] = [];
  do {
    formats.begin();
    done = filled(formats, cells);
    frames++;
    assert.ok(frames < 20);
  } while (formats.pending);
  assert.deepEqual(done, cells);
  assert.ok(frames > 1, 'the work was spread over several frames');
  // Without begin() there is no limit.
  const direct = new ConditionalFormats(workbook, sheet, { ...withEngine(workbook), budget: 25, clock: () => (now += 1000) });
  assert.deepEqual(filled(direct, cells), cells);
  assert.equal(direct.pending, false);
});

// --- damaged rules -----------------------------------------------------------

test('conditional: damaged rules neither throw nor hang', () => {
  const base = { stopIfTrue: false, formulas: [], percent: false, bottom: false, aboveAverage: true, equalAverage: false };
  const junk = { top: 0, left: 0, bottom: 5, right: 0 };
  const rules = [
    { ...base, type: 'cellIs', priority: 1, dxfId: 12345, operator: 'equal', formulas: ['1'] },
    { ...base, type: 'cellIs', priority: NaN, dxfId: -3 },
    { ...base, type: 'colorScale', priority: 2 },
    { ...base, type: 'colorScale', priority: 2, colorScale: { values: [{ type: 'bogus', gte: true }, { type: 'num', value: 'x y z', gte: true }], colors: [{ rgb: 'ZZZ' }, {}] } },
    { ...base, type: 'colorScale', priority: 2, colorScale: { values: [], colors: [{ rgb: 'FF0000' }, { rgb: '00FF00' }] } },
    { ...base, type: 'colorScale', priority: 2, colorScale: { values: [{ type: 'num', value: '1e999', gte: true }, { type: 'percentile', value: '-5', gte: true }], colors: [{ rgb: 'FF0000' }, { rgb: '00FF00' }] } },
    { ...base, type: 'dataBar', priority: 3 },
    { ...base, type: 'dataBar', priority: 3, dataBar: { min: undefined, max: null, minLength: -5, maxLength: 1e9 } },
    { ...base, type: 'iconSet', priority: 4, iconSet: { set: 7, values: 'no', custom: [{}] } },
    { ...base, type: 'iconSet', priority: 4, iconSet: { set: '__proto__', values: [{ type: 'min' }, { type: 'max' }, null] } },
    { ...base, type: 'top10', priority: 5, rank: -1, dxfId: 0 },
    { ...base, type: 'top10', priority: 5, rank: 1e12, percent: true, dxfId: 0 },
    { ...base, type: 'aboveAverage', priority: 6, stdDev: 1e308, dxfId: 0 },
    { ...base, type: 'timePeriod', priority: 7, timePeriod: 'constructor', dxfId: 0 },
    { ...base, type: 'containsText', priority: 8, text: '*'.repeat(100000), dxfId: 0 },
    { ...base, type: 'expression', priority: 9, formulas: ['('.repeat(100000)], dxfId: 0 },
    { ...base, type: 'constructor', priority: 10 },
    null,
    undefined,
    { type: 'cellIs' },
  ] as unknown as ConditionalRule[];
  const { workbook, sheet } = book(sheetData(column([1, 2, 'a*b', null, { error: '#N/A' }, 3])));
  sheet.conditionalFormats = [
    { ranges: [junk, { top: 5, left: 0, bottom: 1, right: 0 }, null as never], rules },
    { ranges: [], rules },
    { ranges: [junk], rules: null as never },
    null as never,
  ] as ConditionalFormat[];
  const formats = new ConditionalFormats(workbook, sheet, withEngine(workbook));
  for (let r = 0; r < 8; r++) {
    for (let c = 0; c < 3; c++) assert.doesNotThrow(() => formats.resolve(r, c));
  }
  for (const position of [-1, NaN, Infinity, 1e12, 0.5]) assert.equal(formats.resolve(position, position), undefined);
  formats.invalidate();
  assert.doesNotThrow(() => formats.resolve(0, 0));

  // A list that is not a list, and styles without differential formats.
  sheet.conditionalFormats = 'broken' as never;
  assert.equal(new ConditionalFormats(workbook, sheet).resolve(0, 0), undefined);
  assert.equal(new ConditionalFormats(workbook, sheet).active, false);
  const bare = book(sheetData([[1]]) + cf('A1', '<cfRule type="cellIs" dxfId="0" priority="1" operator="equal"><formula>1</formula></cfRule>'), '');
  bare.workbook.styles.dxfFormats = undefined;
  assert.equal(new ConditionalFormats(bare.workbook, bare.sheet).resolve(0, 0), undefined);
});

test('conditional: 100,000 rules and a range as large as the sheet', () => {
  const rules: string[] = [];
  for (let i = 0; i < 100000; i++) rules.push(`<cfRule type="cellIs" dxfId="${i % 4}" priority="${i + 1}" operator="equal"><formula>${i}</formula></cfRule>`);
  let started = Date.now();
  const { workbook, sheet } = book(sheetData(column([3, 99999, 'x'])) + cf('A1:XFD1048576', rules.join('')) + cf('A1:XFD1048576', '<cfRule type="aboveAverage" dxfId="0" priority="200000"/><cfRule type="duplicateValues" dxfId="0" priority="200001"/>'));
  const readMs = Date.now() - started;
  assert.equal(sheet.conditionalFormats![0].rules.length, 100000);
  const formats = new ConditionalFormats(workbook, sheet);
  started = Date.now();
  assert.deepEqual(at(formats, 'A1')!.style, { color: '#0000FF' });
  // 99999 matches its own rule (blue font) and is above the average (red fill).
  assert.deepEqual(at(formats, 'A2')!.style, { color: '#0000FF', fill: '#FF0000' });
  assert.equal(at(formats, 'A3'), undefined);
  // UNSURE: an empty cell compares as 0 (see the cellIs test), so the rule "equal to 0" takes the far corner of the sheet.
  assert.deepEqual(formats.resolve(1048575, 16383)!.style, { fill: '#FF0000' });
  for (let r = 0; r < 10; r++) for (let c = 0; c < 5; c++) formats.resolve(r + 5000, c);
  const evalMs = Date.now() - started;
  console.log(`conditional: 100,000 rules read in ${readMs} ms, 54 cells resolved in ${evalMs} ms`);
  assert.ok(readMs < 10000, `reading took ${readMs} ms`);
  assert.ok(evalMs < 10000, `resolving took ${evalMs} ms`);

  // Within a frame the work stops when its time is up, however many rules there are; the next frames go on.
  let now = 0;
  const framed = new ConditionalFormats(workbook, sheet, { budget: 5, clock: () => (now += 1) });
  let frames = 0;
  let missing = 0;
  do {
    framed.begin();
    missing = 0;
    for (let r = 0; r < 30; r++) if (!framed.resolve(r + 100, 0)) missing++;
    frames++;
    assert.ok(frames < 100);
  } while (framed.pending);
  assert.equal(missing, 0);
  assert.ok(frames > 2, `${frames} frames`);
});
