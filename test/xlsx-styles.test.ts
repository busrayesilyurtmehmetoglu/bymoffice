import assert from 'node:assert/strict';
import { test } from 'node:test';
import { MAX_STYLE_RECORDS, parseColor, parseFont, parseStyles, parseThemeColors } from '../src/formats/xlsx/styles';
import { formatCodeOf, resolveCellStyle, resolveColor } from '../src/model/sheet/color';
import { emptyStyleSheet } from '../src/model/sheet/types';
import { parseXml } from '../src/xml';

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
const A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';

const sheetXml = (body: string) => parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet ${NS}>${body}</styleSheet>`);

// Shaped like what Excel writes; hand-made.
const FULL_STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet ${NS} xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <numFmts count="2">
    <numFmt numFmtId="164" formatCode="dd/mm/yyyy;@"/>
    <numFmt numFmtId="165" formatCode="#,##0.00\\ &quot;₺&quot;"/>
  </numFmts>
  <fonts count="4" x14ac:knownFonts="1">
    <font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><charset val="162"/><scheme val="minor"/></font>
    <font><b/><i val="1"/><strike/><u/><sz val="14.5"/><color rgb="ffff0000"/><name val="Arial"/><family val="2"/></font>
    <font><b val="0"/><i val="false"/><strike val="0"/><u val="double"/><vertAlign val="superscript"/><sz val="9"/><color indexed="10"/><name val="Times New Roman"/><scheme val="none"/></font>
    <font><u val="none"/><vertAlign val="subscript"/><color theme="4" tint="-0.249977111117893"/><name val="Cambria"/><scheme val="major"/></font>
  </fonts>
  <fills count="5">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
    <fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>
    <fill><patternFill><fgColor theme="4" tint="0.7999816888943144"/></patternFill></fill>
  </fills>
  <borders count="3">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border diagonalUp="1"><left style="thin"><color indexed="64"/></left><right style="medium"><color rgb="FF00B050"/></right><top style="none"/><bottom style="double"><color auto="1"/></bottom><diagonal style="dashed"><color theme="5"/></diagonal></border>
    <border diagonalDown="true"><start style="hair"/><end style="thick"/><top style="thin"/><bottom/></border>
  </borders>
  <cellStyleXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
    <xf numFmtId="9" fontId="0" fillId="0" borderId="0" applyFont="0" applyFill="0"/>
  </cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" quotePrefix="1">
      <alignment horizontal="center" vertical="top" wrapText="1" shrinkToFit="0" textRotation="45" indent="2" readingOrder="1"/>
      <protection locked="0" hidden="1"/>
    </xf>
    <xf numFmtId="9" fontId="2" fillId="3" borderId="2" xfId="1"><alignment textRotation="255"/></xf>
    <xf numFmtId="14" fontId="3" fillId="4" borderId="0" xfId="0"><protection/></xf>
  </cellXfs>
  <cellStyles count="2">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
    <cellStyle name="Yüzde" xfId="1" builtinId="5"/>
  </cellStyles>
  <dxfs count="2">
    <dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
    <dxf><border><left style="thin"/></border></dxf>
  </dxfs>
  <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
  <colors>
    <indexedColors><rgbColor rgb="00112233"/><rgbColor rgb="00aabbcc"/><rgbColor rgb="bogus"/></indexedColors>
    <mruColors><color rgb="FF00B050"/></mruColors>
  </colors>
  <extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"/></extLst>
</styleSheet>`;

const THEME = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme ${A_NS} name="Office Theme"><a:themeElements><a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
  <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="44546A"/></a:dk2>
  <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
  <a:accent1><a:srgbClr val="4472c4"/></a:accent1>
  <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
  <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
  <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
  <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme><a:fontScheme name="Office"/></a:themeElements></a:theme>`;

test('styles: missing styles part gives the default style sheet', () => {
  assert.deepEqual(parseStyles(undefined, undefined), emptyStyleSheet());
});

test('styles: empty or absent tables keep one default entry each', () => {
  const bodies = [
    '',
    '<fonts count="0"/><fills/><borders/><cellStyleXfs/><cellXfs/><cellStyles/><dxfs/><colors/>',
    '<numFmts/><colors><indexedColors/></colors>',
  ];
  for (const body of bodies) {
    const styles = parseStyles(sheetXml(body), undefined);
    assert.deepEqual(styles, emptyStyleSheet());
    assert.equal(styles.indexedColors, undefined);
  }
});

test('styles: full style sheet - number formats', () => {
  const styles = parseStyles(parseXml(FULL_STYLES), parseXml(THEME));
  assert.deepEqual([...styles.numFmts], [
    [164, 'dd/mm/yyyy;@'],
    [165, '#,##0.00\\ "₺"'],
  ]);
  assert.equal(formatCodeOf(styles, 164), 'dd/mm/yyyy;@');
  assert.equal(formatCodeOf(styles, 9), '0%');
});

test('styles: full style sheet - fonts', () => {
  const { fonts } = parseStyles(parseXml(FULL_STYLES), undefined);
  assert.equal(fonts.length, 4);
  assert.deepEqual(fonts[0], {
    bold: false, italic: false, strike: false, underline: 'none',
    size: 11, color: { theme: 1 }, name: 'Calibri', family: 2, charset: 162, scheme: 'minor',
  });
  assert.deepEqual(fonts[1], {
    bold: true, italic: true, strike: true, underline: 'single',
    size: 14.5, color: { rgb: 'FFFF0000' }, name: 'Arial', family: 2,
  });
  assert.deepEqual(fonts[2], {
    bold: false, italic: false, strike: false, underline: 'double', vertAlign: 'superscript',
    size: 9, color: { indexed: 10 }, name: 'Times New Roman', scheme: 'none',
  });
  assert.deepEqual(fonts[3], {
    bold: false, italic: false, strike: false, underline: 'none', vertAlign: 'subscript',
    color: { theme: 4, tint: -0.249977111117893 }, name: 'Cambria', scheme: 'major',
  });
});

test('styles: font details - boolean values, underline kinds, bad numbers', () => {
  const font = (inner: string) => parseFont(parseXml(`<font>${inner}</font>`).root);
  assert.equal(font('<b val="true"/>').bold, true);
  assert.equal(font('<b val="0"/>').bold, false);
  assert.equal(font('<b val="false"/>').bold, false);
  assert.equal(font('<i val="1"/>').italic, true);
  assert.equal(font('<strike val="false"/>').strike, false);
  assert.equal(font('<u val="singleAccounting"/>').underline, 'singleAccounting');
  assert.equal(font('<u val="doubleAccounting"/>').underline, 'doubleAccounting');
  assert.equal(font('<u val="wavy"/>').underline, 'single');
  assert.equal(font('<sz val="abc"/>').size, undefined);
  assert.equal(font('<sz val="-4"/>').size, undefined);
  assert.equal(font('<family val="x"/>').family, undefined);
  assert.equal(font('<scheme val="other"/>').scheme, undefined);
  assert.equal(font('<vertAlign val="baseline"/>').vertAlign, undefined);
  assert.equal(font('<color/>').color, undefined);
  // Rich text run properties name the font with <rFont>.
  assert.equal(font('<rFont val="Verdana"/><b/>').name, 'Verdana');
  // Unknown children do not disturb the rest.
  assert.deepEqual(font('<condense/><extend val="1"/><outline/><b/>'), {
    bold: true, italic: false, strike: false, underline: 'none',
  });
});

test('styles: colour attributes', () => {
  const color = (attrs: string) => parseColor(parseXml(`<color ${attrs}/>`).root);
  assert.deepEqual(color('rgb="ff00b050"'), { rgb: 'FF00B050' });
  assert.deepEqual(color('rgb="00b050"'), { rgb: '00B050' });
  assert.deepEqual(color('theme="3" tint="0.5"'), { theme: 3, tint: 0.5 });
  assert.deepEqual(color('indexed="64"'), { indexed: 64 });
  assert.deepEqual(color('auto="1"'), { auto: true });
  assert.deepEqual(color('auto="0" rgb="FF000000"'), { rgb: 'FF000000', auto: false });
  assert.deepEqual(color('theme="1" tint="-9"'), { theme: 1, tint: -1 });
  assert.equal(color('rgb="red"'), undefined);
  assert.equal(color('theme="-1"'), undefined);
  assert.equal(color('tint="NaN"'), undefined);
  assert.equal(color('unknown="1"'), undefined);
  assert.equal(parseColor(undefined), undefined);
});

test('styles: full style sheet - fills', () => {
  const { fills } = parseStyles(parseXml(FULL_STYLES), undefined);
  assert.equal(fills.length, 5);
  assert.deepEqual(fills[0], { type: 'pattern', pattern: 'none' });
  assert.deepEqual(fills[1], { type: 'pattern', pattern: 'gray125' });
  assert.deepEqual(fills[2], { type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFFFF00' }, bg: { indexed: 64 } });
  const gradient = fills[3];
  assert.equal(gradient.type, 'gradient');
  if (gradient.type === 'gradient') {
    assert.equal(gradient.raw.name, 'gradientFill');
    assert.equal(gradient.raw.attrs.get('degree'), '90');
    assert.equal(gradient.raw.children.length, 2);
  }
  // No patternType reads as "none".
  assert.deepEqual(fills[4], { type: 'pattern', pattern: 'none', fg: { theme: 4, tint: 0.7999816888943144 } });
});

test('styles: an empty <fill> is a "none" pattern', () => {
  const styles = parseStyles(sheetXml('<fills><fill/><fill><somethingNew/></fill></fills>'), undefined);
  assert.deepEqual(styles.fills, [
    { type: 'pattern', pattern: 'none' },
    { type: 'pattern', pattern: 'none' },
  ]);
});

test('styles: full style sheet - borders', () => {
  const { borders } = parseStyles(parseXml(FULL_STYLES), undefined);
  assert.equal(borders.length, 3);
  assert.deepEqual(borders[0], { diagonalUp: false, diagonalDown: false });
  assert.deepEqual(borders[1], {
    diagonalUp: true,
    diagonalDown: false,
    left: { style: 'thin', color: { indexed: 64 } },
    right: { style: 'medium', color: { rgb: 'FF00B050' } },
    bottom: { style: 'double', color: { auto: true } },
    diagonal: { style: 'dashed', color: { theme: 5 } },
  });
  // start/end are the older names of left/right.
  assert.deepEqual(borders[2], {
    diagonalUp: false,
    diagonalDown: true,
    left: { style: 'hair' },
    right: { style: 'thick' },
    top: { style: 'thin' },
  });
});

test('styles: full style sheet - cell formats', () => {
  const styles = parseStyles(parseXml(FULL_STYLES), undefined);
  assert.deepEqual(styles.cellStyleXfs, [
    { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false },
    { numFmtId: 9, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false },
  ]);
  assert.equal(styles.cellXfs.length, 4);
  assert.deepEqual(styles.cellXfs[0], { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false });
  assert.deepEqual(styles.cellXfs[1], {
    numFmtId: 164, fontId: 1, fillId: 2, borderId: 1, xfId: 0, quotePrefix: true,
    alignment: {
      horizontal: 'center', vertical: 'top', wrapText: true, shrinkToFit: false,
      textRotation: 45, indent: 2, readingOrder: 1,
    },
    protection: { locked: false, hidden: true },
  });
  assert.deepEqual(styles.cellXfs[2], {
    numFmtId: 9, fontId: 2, fillId: 3, borderId: 2, xfId: 1, quotePrefix: false,
    alignment: { textRotation: 255 },
  });
  assert.deepEqual(styles.cellXfs[3], {
    numFmtId: 14, fontId: 3, fillId: 4, borderId: 0, xfId: 0, quotePrefix: false, protection: {},
  });
});

test('styles: full style sheet - named styles, dxfs, palette, theme', () => {
  const styles = parseStyles(parseXml(FULL_STYLES), parseXml(THEME));
  assert.deepEqual(styles.cellStyles, [
    { name: 'Normal', xfId: 0, builtinId: 0 },
    { name: 'Yüzde', xfId: 1, builtinId: 5 },
  ]);
  assert.equal(styles.dxfs.length, 2);
  assert.equal(styles.dxfs[0].name, 'dxf');
  assert.equal(styles.dxfs[0].children.length, 2);
  assert.deepEqual(styles.indexedColors, ['112233', 'AABBCC', '000000']);
  assert.equal(styles.themeColors.length, 12);

  // The pieces work together through the resolvers.
  const resolved = resolveCellStyle(styles, 3);
  assert.equal(resolved.numFmt, 'mm-dd-yy');
  assert.equal(resolved.font.name, 'Cambria');
  assert.equal(resolveColor({ indexed: 1 }, styles), '#AABBCC');
  assert.equal(resolveColor(styles.fonts[0].color, styles), '#000000');
});

test('styles: alignment ignores values outside the schema', () => {
  const styles = parseStyles(
    sheetXml(
      '<cellXfs><xf><alignment horizontal="sideways" vertical="middle" wrapText="maybe" textRotation="200" indent="-1"/></xf></cellXfs>',
    ),
    undefined,
  );
  assert.deepEqual(styles.cellXfs[0].alignment, {});
});

test('styles: out-of-range and malformed indexes read as 0', () => {
  const styles = parseStyles(
    sheetXml(
      '<fonts><font/><font><b/></font></fonts>' +
        '<cellStyleXfs><xf fontId="9"/></cellStyleXfs>' +
        '<cellXfs>' +
        '<xf numFmtId="x" fontId="2" fillId="1" borderId="1" xfId="1"/>' +
        '<xf numFmtId="4" fontId="1" fillId="-1" borderId="4294967296" xfId="abc"/>' +
        '<xf fontId="1.5"/>' +
        '<xf fontId="99999999999999999999999999"/>' +
        '</cellXfs>' +
        '<cellStyles><cellStyle xfId="5"/></cellStyles>',
    ),
    undefined,
  );
  assert.equal(styles.cellStyleXfs[0].fontId, 0);
  assert.deepEqual(styles.cellXfs[0], { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false });
  assert.deepEqual(styles.cellXfs[1], { numFmtId: 4, fontId: 1, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false });
  assert.equal(styles.cellXfs[2].fontId, 0);
  assert.equal(styles.cellXfs[3].fontId, 0);
  assert.deepEqual(styles.cellStyles, [{ name: '', xfId: 0 }]);
  for (let i = 0; i < styles.cellXfs.length; i++) assert.doesNotThrow(() => resolveCellStyle(styles, i));
});

test('styles: number formats without an id or a code are skipped', () => {
  const styles = parseStyles(
    sheetXml('<numFmts><numFmt formatCode="0.0"/><numFmt numFmtId="170"/><numFmt numFmtId="171" formatCode=""/></numFmts>'),
    undefined,
  );
  assert.deepEqual([...styles.numFmts], [[171, '']]);
});

test('styles: prefixed document reads the same', () => {
  const prefixed = FULL_STYLES.replace(/<(\/?)([A-Za-z])/g, '<$1x:$2').replace(
    'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
  );
  const doc = parseXml(prefixed);
  assert.equal(doc.root.name, 'x:styleSheet');
  const plain = parseStyles(parseXml(FULL_STYLES), undefined);
  const read = parseStyles(doc, undefined);
  // Raw elements keep their own (prefixed) names, so compare everything else.
  assert.deepEqual(
    { ...read, dxfs: [], fills: read.fills.filter((f) => f.type === 'pattern') },
    { ...plain, dxfs: [], fills: plain.fills.filter((f) => f.type === 'pattern') },
  );
  assert.equal(read.dxfs.length, 2);
  assert.equal(read.fills[3].type, 'gradient');
});

test('styles: a root that is not a style sheet gives defaults', () => {
  assert.deepEqual(parseStyles(parseXml('<worksheet><sheetData/></worksheet>'), undefined), emptyStyleSheet());
});

test('styles: oversized tables are refused with a clear error', () => {
  const many = '<xf/>'.repeat(MAX_STYLE_RECORDS + 1);
  assert.throws(
    () => parseStyles(sheetXml(`<cellXfs>${many}</cellXfs>`), undefined),
    /Invalid styles: <cellXfs> has 65537 records, the limit is 65536/,
  );
  const fonts = '<font/>'.repeat(MAX_STYLE_RECORDS + 1);
  assert.throws(() => parseStyles(sheetXml(`<fonts>${fonts}</fonts>`), undefined), /<fonts> has 65537 records/);
  // Exactly at the limit is accepted.
  const atLimit = parseStyles(sheetXml(`<cellXfs>${'<xf/>'.repeat(MAX_STYLE_RECORDS)}</cellXfs>`), undefined);
  assert.equal(atLimit.cellXfs.length, MAX_STYLE_RECORDS);
});

test('theme: colours come out in spreadsheet index order', () => {
  assert.deepEqual(parseThemeColors(parseXml(THEME)), [
    'FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
    'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72',
  ]);
});

test('theme: order does not depend on document order or prefix', () => {
  const theme = parseXml(
    `<t:theme xmlns:t="http://schemas.openxmlformats.org/drawingml/2006/main"><t:themeElements><t:clrScheme name="x">` +
      `<t:lt1><t:srgbClr val="FEFEFE"/></t:lt1><t:dk1><t:srgbClr val="010101"/></t:dk1>` +
      `<t:lt2><t:srgbClr val="EEEEEE"/></t:lt2><t:dk2><t:srgbClr val="222222"/></t:dk2>` +
      `</t:clrScheme></t:themeElements></t:theme>`,
  );
  const colors = parseThemeColors(theme);
  assert.equal(colors.length, 12);
  assert.deepEqual(colors.slice(0, 4), ['FEFEFE', '010101', 'EEEEEE', '222222']);
  // Missing scheme colours hold their position as unresolvable entries.
  assert.deepEqual(colors.slice(4), ['', '', '', '', '', '', '', '']);
  const styles = parseStyles(undefined, theme);
  assert.equal(resolveColor({ theme: 1 }, styles), '#010101');
  assert.equal(resolveColor({ theme: 4 }, styles), undefined);
});

test('theme: system colours without lastClr, unsupported colour kinds, no scheme', () => {
  const theme = parseXml(
    `<a:theme ${A_NS}><a:themeElements><a:clrScheme name="x">` +
      `<a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>` +
      `<a:dk2><a:sysClr val="highlight"/></a:dk2><a:lt2><a:srgbClr val="nothex"/></a:lt2>` +
      `<a:accent1><a:hslClr hue="0" sat="0" lum="0"/></a:accent1><a:accent2/>` +
      `</a:clrScheme></a:themeElements></a:theme>`,
  );
  assert.deepEqual(parseThemeColors(theme).slice(0, 6), ['FFFFFF', '000000', '', '', '', '']);
  assert.deepEqual(parseThemeColors(parseXml(`<a:theme ${A_NS}><a:themeElements/></a:theme>`)), []);
  assert.deepEqual(parseThemeColors(parseXml('<other/>')), []);
  assert.deepEqual(parseStyles(undefined, parseXml('<other/>')), emptyStyleSheet());
});

test('theme + styles: tinted theme fill resolves to the colour Excel shows', () => {
  const styles = parseStyles(parseXml(FULL_STYLES), parseXml(THEME));
  const fill = styles.fills[4];
  assert.equal(fill.type, 'pattern');
  const hex = fill.type === 'pattern' ? resolveColor(fill.fg, styles) : undefined;
  assert.ok(hex);
  const expected = [0xda, 0xe3, 0xf3];
  for (let i = 0; i < 3; i++) {
    assert.ok(Math.abs(parseInt(hex.slice(1 + i * 2, 3 + i * 2), 16) - expected[i]) <= 1, hex);
  }
});
