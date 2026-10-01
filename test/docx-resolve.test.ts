// Regression tests for the findings of the independent review of the docx
// reader that concern effective formatting: theme colours, the order in which
// styles, numbering and table conditions apply, and the implied defaults.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { NumberingCounters, formatCounter, levelOf } from '../src/formats/docx/numbering';
import {
  IMPLICIT_PARA_PROPS,
  IMPLICIT_RUN_PROPS,
  resolveCellProps,
  resolveColor,
  resolveFont,
  resolveNumberRunProps,
  resolveParaProps,
  resolveRunProps,
  resolveTableProps,
  tableConditions,
} from '../src/formats/docx/resolve';
import type { Paragraph, Run, Table } from '../src/model/doc/types';
import { readSpec, THEME } from './docx-fixtures';

const style = (id: string, type: string, inner: string, attrs = '') => `<w:style w:type="${type}" w:styleId="${id}"${attrs}>${inner}</w:style>`;
const cell = (text: string, props = '') => `<w:tc>${props}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells: string, props = '') => `<w:tr>${props}${cells}</w:tr>`;

// ------------------------------------------------------------ theme colours

/** Largest difference between the channels of two "#RRGGBB" / "RRGGBB" colours. */
function distance(a: string, b: string): number {
  const channels = (text: string) => [0, 2, 4].map((i) => parseInt(text.replace('#', '').slice(i, i + 2), 16));
  const [x, y] = [channels(a), channels(b)];
  return Math.max(...x.map((value, i) => Math.abs(value - y[i])));
}

test('docx resolve: tint and shade give the colours Word writes', async () => {
  const doc = await readSpec({ theme: THEME, body: '<w:p/>' });
  const color = (theme: string, tint?: number, shade?: number) => resolveColor(doc, { theme, tint, shade });

  // Both together: the tint is applied first, then the shade.
  assert.equal(color('text1', 0x80, 0x95), '#4A4A4A');
  assert.equal(color('text1', 0x80), '#7F7F7F');
  // The same order on a colour that has a hue. The value is one found next to this reference in
  // documents; the other order gives a much lighter colour.
  assert.ok(distance(color('accent1', 0x9a, 0x95)!, '335AA1') <= 1, color('accent1', 0x9a, 0x95));
  // Values Word wrote next to these references in documents (theme: Office 2013).
  assert.equal(color('accent1', 0x99), '#8EAADB');
  assert.equal(color('accent1', 0x66), '#B4C6E7');
  assert.equal(color('accent1', 0x33), '#D9E2F3');
  assert.equal(color('accent1', undefined, 0xbf), '#2F5496');
  assert.equal(color('accent1', undefined, 0x80), '#1F3864');
  assert.equal(color('accent2', 0x66), '#F7CAAC');
  assert.equal(color('accent5', undefined, 0xbf), '#2E74B5');
  // 255 is "no change", for either.
  assert.equal(color('accent1', 255, 255), '#4472C4');
  assert.equal(color('accent1', 255, 0x80), '#1F3864');

  // UNSURE: the rest of the palette of this theme as recalled from documents
  // Word produced, not checked against Word here. The arithmetic Word uses is
  // not published; the reader's reproduces these to within one step of one
  // channel, and most of them exactly.
  const palette: [theme: string, values: [tint: number | undefined, shade: number | undefined, rgb: string][]][] = [
    ['accent2', [[0x33, , 'FBE4D5'], [0x99, , 'F4B083'], [, 0xbf, 'C45911'], [, 0x80, '833C0B']]],
    ['accent3', [[0x33, , 'EDEDED'], [0x66, , 'DBDBDB'], [0x99, , 'C9C9C9'], [, 0xbf, '7B7B7B'], [, 0x80, '525252']]],
    ['accent4', [[0x33, , 'FFF2CC'], [0x66, , 'FFE599'], [0x99, , 'FFD966'], [, 0xbf, 'BF8F00'], [, 0x80, '7F6000']]],
    ['accent5', [[0x33, , 'DEEAF6'], [0x66, , 'BDD6EE'], [0x99, , '9CC2E5'], [, 0x80, '1F4E79']]],
    ['accent6', [[0x33, , 'E2EFD9'], [0x66, , 'C5E0B3'], [0x99, , 'A8D08D'], [, 0xbf, '538135'], [, 0x80, '385623']]],
    ['text2', [[0x33, , 'D5DCE4'], [0x66, , 'ACB9CA'], [0x99, , '8496B0'], [, 0xbf, '323E4F'], [, 0x80, '222A35']]],
    ['background2', [[, 0xe6, 'D0CECE'], [, 0xbf, 'AEAAAA'], [, 0x80, '767171']]],
    ['background1', [[, 0xf2, 'F2F2F2'], [, 0xd9, 'D9D9D9'], [, 0xbf, 'BFBFBF'], [, 0xa6, 'A6A6A6'], [, 0x80, '808080']]],
    ['text1', [[0xa6, , '595959'], [0xbf, , '404040'], [0xd9, , '262626'], [0xf2, , '0D0D0D']]],
  ];
  let total = 0;
  const inexact: string[] = [];
  for (const [theme, values] of palette) {
    for (const [tint, shade, rgb] of values) {
      const got = color(theme, tint, shade)!;
      total++;
      assert.ok(distance(got, rgb) <= 1, `${theme} tint ${tint} shade ${shade}: ${got} is not within one step of ${rgb}`);
      if (got !== '#' + rgb) inexact.push(`${theme}/${tint === undefined ? 's' + shade!.toString(16) : 't' + tint.toString(16)}`);
    }
  }
  assert.equal(total, 40);
  // The ones that are one step off, so that a change of the arithmetic shows up here.
  assert.deepEqual(inexact, ['accent4/sbf', 'accent6/t66', 'text2/t99', 'background2/s80']);

  // The value a document carries next to the reference is the writing application's own result. When it
  // is the same colour to within its rounding, the document shows it as it was saved...
  assert.equal(color('accent4', undefined, 0xbf), '#BE8F00');
  assert.equal(resolveColor(doc, { rgb: 'BF8F00', theme: 'accent4', shade: 0xbf }), '#BF8F00');
  assert.equal(resolveColor(doc, { rgb: 'BC9102', theme: 'accent4', shade: 0xbf }), '#BC9102');
  // ...and when it is not (the theme was changed since, or it is a placeholder), the theme decides.
  assert.equal(resolveColor(doc, { rgb: 'BB8F00', theme: 'accent4', shade: 0xbf }), '#BE8F00');
  assert.equal(resolveColor(doc, { rgb: 'BE8F03', theme: 'accent4', shade: 0xbf }), '#BE8F00');
  assert.equal(resolveColor(doc, { rgb: '000000', theme: 'accent4', shade: 0xbf }), '#BE8F00');
  assert.equal(resolveColor(doc, { rgb: 'FFFFFF', theme: 'accent1', tint: 0x33 }), '#D9E2F3');
  // Without tint or shade there is nothing to round: the theme colour itself.
  assert.equal(resolveColor(doc, { rgb: 'FFC001', theme: 'accent4' }), '#FFC000');
  assert.equal(resolveColor(doc, { rgb: 'FFC001', theme: 'accent4', tint: 255, shade: 255 }), '#FFC000');
});

// ------------------------------------------------- numbering and paragraphs

const NUMBERING =
  '<w:abstractNum w:abstractNumId="0">' +
  '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:tabs><w:tab w:val="num" w:pos="432"/></w:tabs><w:ind w:left="432" w:hanging="432"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Symbol"/><w:b w:val="0"/><w:color w:val="00FF00"/></w:rPr></w:lvl>' +
  '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2."/><w:pPr><w:ind w:left="864" w:hanging="432"/></w:pPr></w:lvl>' +
  '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';

test('docx resolve: the indent of a numbering level sits right below the style that brings the numbering', async () => {
  const numbered = '<w:numPr><w:numId w:val="1"/></w:numPr>';
  const doc = await readSpec({
    numbering: NUMBERING,
    styles:
      style('Num', 'paragraph', `<w:pPr>${numbered}<w:ind w:left="0" w:firstLine="0"/></w:pPr>`) +
      style('NumChild', 'paragraph', '<w:basedOn w:val="Num"/><w:pPr><w:ind w:left="50"/></w:pPr>') +
      style('Base', 'paragraph', '<w:pPr><w:ind w:left="999" w:hanging="10"/><w:tabs><w:tab w:val="left" w:pos="100"/></w:tabs></w:pPr>') +
      style('NumOnBase', 'paragraph', `<w:basedOn w:val="Base"/><w:pPr>${numbered}</w:pPr>`) +
      style('Leaf', 'paragraph', '<w:basedOn w:val="NumOnBase"/><w:pPr><w:ind w:hanging="5"/></w:pPr>'),
    body:
      '<w:p><w:pPr><w:pStyle w:val="Num"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="NumChild"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="NumOnBase"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Leaf"/></w:pPr></w:p>' +
      `<w:p><w:pPr><w:pStyle w:val="Base"/>${numbered}</w:pPr></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Base"/>${numbered}<w:ind w:left="7"/></w:pPr></w:p>` +
      '<w:p><w:pPr><w:pStyle w:val="Num"/><w:numPr><w:ilvl w:val="1"/></w:numPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Num"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Base"/></w:pPr></w:p>',
  });
  const props = (i: number) => resolveParaProps(doc, (doc.body.blocks[i] as Paragraph).props);
  // The style states numbering and an indent: its own indent wins over the level's.
  assert.deepEqual(props(0).indent, { left: 0, firstLine: 0 });
  assert.deepEqual(props(0).tabs, [{ pos: 432, align: 'num' }]);
  // So does the indent of a style derived from it.
  assert.deepEqual(props(1).indent, { left: 50, firstLine: 0 });
  // The level wins over the styles below the one that brings the numbering...
  assert.deepEqual(props(2).indent, { left: 432, hanging: 432 });
  assert.deepEqual(props(2).tabs, [
    { pos: 100, align: 'left' },
    { pos: 432, align: 'num' },
  ]);
  // ...and loses to the ones above it.
  assert.deepEqual(props(3).indent, { left: 432, hanging: 5 });
  // Numbering given to the paragraph itself: the level wins over every style, direct indents over the level.
  assert.deepEqual(props(4).indent, { left: 432, hanging: 432 });
  assert.deepEqual(props(5).indent, { left: 7, hanging: 432 });
  // UNSURE: a paragraph that only names another level of the numbering of its style is read as numbered directly.
  assert.deepEqual(props(6).indent, { left: 864, hanging: 432 });
  assert.deepEqual(props(6).numbering, { numId: 1, level: 1 });
  // Numbering taken away: no level, no level indent.
  assert.deepEqual(props(7).indent, { left: 0, firstLine: 0 });
  assert.deepEqual(props(7).tabs, undefined);
  assert.deepEqual(props(8).indent, { left: 999, hanging: 10 });
});

test('docx resolve: legal numbering shows every level as a decimal, its own included', async () => {
  const lvl = (ilvl: number, fmt: string, text: string, extra = '') =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>${extra}<w:lvlText w:val="${text}"/></w:lvl>`;
  const doc = await readSpec({
    numbering:
      `<w:abstractNum w:abstractNumId="0">${lvl(0, 'upperRoman', '%1')}${lvl(1, 'lowerLetter', '%1.%2', '<w:isLgl/>')}${lvl(2, 'lowerLetter', '%1.%2.%3')}` +
      `${lvl(3, 'bullet', 'o', '<w:isLgl/>')}${lvl(4, 'none', '%1-%5', '<w:isLgl/>')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`,
    body: '<w:p/>',
  });
  const counters = new NumberingCounters(doc.numbering, doc.styles);
  assert.equal(counters.next(1, 0)!.text, 'I');
  assert.equal(counters.next(1, 1)!.text, '1.1');
  assert.equal(counters.next(1, 1)!.text, '1.2');
  // Without the flag each level keeps its own format.
  assert.equal(counters.next(1, 2)!.text, 'I.b.a');
  assert.equal(counters.next(1, 3)!.text, 'o');
  // A level without a number shows none, legal or not.
  assert.equal(counters.next(1, 4)!.text, '1-');
});

test('docx resolve: counter text takes an optional language and ignores it for now', () => {
  // UNSURE: Word may use the alphabet of the text language for letters (Turkish has ç, ğ, ı, ...); not verified, so not done.
  assert.equal(formatCounter(3, 'lowerLetter', 'tr-TR'), 'c');
  assert.equal(formatCounter(29, 'upperLetter', 'tr-TR'), formatCounter(29, 'upperLetter'));
  assert.equal(formatCounter(4, 'decimal', undefined), '4');
});

test('docx resolve: the label of a numbered paragraph is formatted by the paragraph mark and then the level', async () => {
  const doc = await readSpec({
    numbering: NUMBERING,
    theme: THEME,
    styles:
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      style('Body', 'paragraph', '<w:rPr><w:i/><w:sz w:val="28"/></w:rPr>'),
    body:
      '<w:p><w:pPr><w:pStyle w:val="Body"/><w:numPr><w:numId w:val="1"/></w:numPr><w:rPr><w:b/><w:color w:val="FF0000"/><w:sz w:val="30"/><w:u w:val="single"/></w:rPr></w:pPr>' +
      '<w:r><w:rPr><w:strike/></w:rPr><w:t>x</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>',
  });
  const first = doc.body.blocks[0] as Paragraph;
  const label = resolveNumberRunProps(doc, first.props, levelOf(doc.numbering, doc.styles, 1, 0)!);
  assert.deepEqual(label, {
    // From the level, on top of everything else.
    fonts: { ascii: 'Symbol', hAnsi: 'Calibri', hAnsiTheme: 'minorHAnsi', eastAsia: 'Times New Roman', cs: 'Times New Roman' },
    bold: false,
    color: { rgb: '00FF00' },
    // From the paragraph mark.
    size: 30,
    sizeCs: 30,
    underline: { style: 'single' },
    // From the paragraph style.
    italic: true,
  });
  // The formatting of the runs of the paragraph plays no part.
  assert.equal(label.strike, undefined);
  // A level without run properties: the paragraph mark alone, here with nothing but the defaults.
  const second = doc.body.blocks[1] as Paragraph;
  assert.deepEqual(resolveNumberRunProps(doc, second.props, levelOf(doc.numbering, doc.styles, 1, 1)!), {
    fonts: { ascii: 'Calibri', asciiTheme: 'minorHAnsi', hAnsi: 'Calibri', hAnsiTheme: 'minorHAnsi', eastAsia: 'Times New Roman', cs: 'Times New Roman' },
    size: 22,
    sizeCs: 22,
  });
});

// --------------------------------------------------------- implied defaults

test('docx resolve: what Word assumes when a document does not say is defined in one place', async () => {
  assert.deepEqual(IMPLICIT_RUN_PROPS, {
    fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', cs: 'Times New Roman' },
    size: 20,
  });
  assert.deepEqual(IMPLICIT_PARA_PROPS, { align: 'left', spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' } });
  assert.ok(Object.isFrozen(IMPLICIT_RUN_PROPS) && Object.isFrozen(IMPLICIT_RUN_PROPS.fonts) && Object.isFrozen(IMPLICIT_PARA_PROPS.spacing));

  // No styles part at all: 10 point Times New Roman, left aligned, single spaced.
  const bare = await readSpec({ body: '<w:p><w:r><w:t>x</w:t></w:r><w:r><w:rPr><w:szCs w:val="30"/><w:rFonts w:cs="Arial"/></w:rPr></w:r></w:p>' });
  const p = bare.body.blocks[0] as Paragraph;
  const plain = resolveRunProps(bare, p.props, (p.children[0] as Run).props);
  assert.deepEqual(plain, { fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', cs: 'Times New Roman' }, size: 20, sizeCs: 20 });
  assert.deepEqual(resolveParaProps(bare, p.props), { align: 'left', spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' } });
  // The results are fresh objects: changing one does not change the defaults.
  plain.fonts!.ascii = 'Changed';
  assert.equal(IMPLICIT_RUN_PROPS.fonts!.ascii, 'Times New Roman');
  assert.equal(resolveRunProps(bare, p.props, (p.children[0] as Run).props).fonts!.ascii, 'Times New Roman');
  // A complex script size of its own is kept; without one it is the size.
  const own = resolveRunProps(bare, p.props, (p.children[1] as Run).props);
  assert.deepEqual([own.size, own.sizeCs, own.fonts], [20, 30, { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', cs: 'Arial' }]);

  // Defaults that state some things leave the rest implied.
  const partial = await readSpec({
    theme: THEME,
    styles:
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cs="Arial"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      style('Small', 'character', '<w:rPr><w:sz w:val="16"/></w:rPr>'),
    body: '<w:p><w:r><w:t>x</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Small"/></w:rPr></w:r></w:p>',
  });
  const q = partial.body.blocks[0] as Paragraph;
  const resolved = resolveRunProps(partial, q.props, (q.children[0] as Run).props);
  assert.deepEqual(resolved, {
    // A theme font is followed; the theme names no East Asian font, so that slot has the implied one.
    fonts: { ascii: 'Calibri Light', asciiTheme: 'majorHAnsi', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', eastAsiaTheme: 'minorEastAsia', cs: 'Arial' },
    size: 24,
    sizeCs: 24,
  });
  for (const slot of ['ascii', 'hAnsi', 'eastAsia', 'cs'] as const) assert.equal(resolveFont(partial, resolved.fonts, slot), resolved.fonts![slot]);
  assert.deepEqual(resolveParaProps(partial, q.props), { align: 'left', spacing: { before: 0, after: 160, line: 240, lineRule: 'auto' } });
  const small = resolveRunProps(partial, q.props, (q.children[1] as Run).props);
  assert.deepEqual([small.size, small.sizeCs], [16, 16]);
});

// ------------------------------------------------------------------- tables

const TABLE_STYLE = style(
  'T',
  'table',
  '<w:rPr><w:color w:val="000001"/></w:rPr>' +
    ['band1Vert', 'band2Vert', 'band1Horz', 'band2Horz', 'firstCol', 'lastCol', 'firstRow', 'lastRow', 'neCell', 'nwCell', 'seCell', 'swCell']
      .map((type, i) => `<w:tblStylePr w:type="${type}"><w:rPr><w:color w:val="0000${(i + 16).toString(16).toUpperCase()}"/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="${type === 'firstRow' ? 'AAAAAA' : type === 'firstCol' ? 'BBBBBB' : 'CCCCCC'}"/></w:tcPr></w:tblStylePr>`)
      .join(''),
);
const look = (flags: string) => `<w:tblLook ${flags}/>`;
const ALL = 'w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"';

test('docx resolve: table style conditions apply in a fixed order, rows over columns over bands', async () => {
  const three = row(cell('a') + cell('b') + cell('c'));
  const doc = await readSpec({
    styles: TABLE_STYLE,
    body:
      `<w:tbl><w:tblPr><w:tblStyle w:val="T"/>${look(ALL)}</w:tblPr>${three}${three}${three}${three}</w:tbl>` +
      // One column: it is the first column and nothing else.
      `<w:tbl><w:tblPr><w:tblStyle w:val="T"/>${look(ALL)}</w:tblPr>${row(cell('a'))}${row(cell('b'))}</w:tbl>` +
      // One row: it is the first row and nothing else.
      `<w:tbl><w:tblPr><w:tblStyle w:val="T"/>${look(ALL)}</w:tblPr>${three}</w:tbl>` +
      `<w:tbl><w:tblPr><w:tblStyle w:val="T"/>${look('w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="1" w:noHBand="1" w:noVBand="1"')}</w:tblPr>${row(cell('a'))}</w:tbl>`,
  });
  const table = doc.body.blocks[0] as Table;
  const at = (r: number, c: number, t = table) => tableConditions(doc, { table: t, row: r, cell: c });
  // UNSURE: this is the order asked for by the review of the reader (a header row wins over the first
  // column in the corner cell, as Word shows it). ECMA-376 17.7.6 lists rows before columns, which would
  // let the column win. Not checked against Word here.
  assert.deepEqual(at(0, 0), ['firstCol', 'firstRow', 'nwCell']);
  assert.deepEqual(at(0, 1), ['band1Vert', 'firstRow']);
  assert.deepEqual(at(0, 2), ['band2Vert', 'lastCol', 'firstRow', 'neCell']);
  assert.deepEqual(at(1, 0), ['band1Horz', 'firstCol']);
  assert.deepEqual(at(1, 1), ['band1Vert', 'band1Horz']);
  assert.deepEqual(at(2, 2), ['band2Vert', 'band2Horz', 'lastCol']);
  assert.deepEqual(at(3, 0), ['band1Horz', 'firstCol', 'lastRow', 'swCell']);
  assert.deepEqual(at(3, 1), ['band1Vert', 'band1Horz', 'lastRow']);
  assert.deepEqual(at(3, 2), ['band2Vert', 'band1Horz', 'lastCol', 'lastRow', 'seCell']);

  // The later condition wins: in the corner the header row decides, not the first column.
  assert.deepEqual(resolveCellProps(doc, { table, row: 0, cell: 0 }).shading, { pattern: 'clear', fill: { rgb: 'CCCCCC' } });
  const first = table.rows[0].cells[0].blocks[0] as Paragraph;
  const rgb = (r: number, c: number) => resolveRunProps(doc, first.props, {}, { table, row: r, cell: c }).color?.rgb;
  assert.equal(rgb(0, 0), '000019'); // nwCell
  assert.equal(rgb(0, 1), '000016'); // firstRow over band1Vert
  assert.equal(rgb(1, 0), '000014'); // firstCol over band1Horz
  assert.equal(rgb(1, 1), '000012'); // band1Horz over band1Vert
  assert.equal(rgb(3, 2), '00001A'); // seCell
  const noCorner = await readSpec({
    styles: style('T', 'table', '<w:tblStylePr w:type="firstRow"><w:tcPr><w:shd w:val="clear" w:fill="AAAAAA"/></w:tcPr></w:tblStylePr><w:tblStylePr w:type="firstCol"><w:tcPr><w:shd w:val="clear" w:fill="BBBBBB"/></w:tcPr></w:tblStylePr>'),
    body: `<w:tbl><w:tblPr><w:tblStyle w:val="T"/>${look(ALL)}</w:tblPr>${three}${three}</w:tbl>`,
  });
  assert.deepEqual(resolveCellProps(noCorner, { table: noCorner.body.blocks[0] as Table, row: 0, cell: 0 }).shading, { pattern: 'clear', fill: { rgb: 'AAAAAA' } });

  // A column that is both first and last is the first column; likewise a row.
  const narrow = doc.body.blocks[1] as Table;
  assert.deepEqual(at(0, 0, narrow), ['firstCol', 'firstRow', 'nwCell']);
  assert.deepEqual(at(1, 0, narrow), ['band1Horz', 'firstCol', 'lastRow', 'swCell']);
  const flat = doc.body.blocks[2] as Table;
  assert.deepEqual(at(0, 0, flat), ['firstCol', 'firstRow', 'nwCell']);
  assert.deepEqual(at(0, 2, flat), ['band2Vert', 'lastCol', 'firstRow', 'neCell']);
  // Without the first column in play, the only column is the last one.
  assert.deepEqual(at(0, 0, doc.body.blocks[3] as Table), ['lastCol']);
});

test('docx resolve: the table property exceptions of a row take part', async () => {
  const doc = await readSpec({
    styles: TABLE_STYLE,
    body:
      `<w:tbl><w:tblPr><w:tblStyle w:val="T"/><w:jc w:val="center"/>${look('w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"')}</w:tblPr>` +
      row(cell('a') + cell('b')) +
      row(cell('c') + cell('d'), `<w:tblPrEx><w:jc w:val="right"/>${look('w:firstRow="0" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="1" w:noVBand="1"')}</w:tblPrEx>`) +
      '</w:tbl>',
  });
  const table = doc.body.blocks[0] as Table;
  assert.equal(resolveTableProps(doc, table).align, 'center');
  assert.equal(resolveTableProps(doc, table, 0).align, 'center');
  assert.equal(resolveTableProps(doc, table, 1).align, 'right');
  assert.equal(resolveTableProps(doc, table, 1).style, 'T');
  assert.equal(resolveTableProps(doc, table, 7).align, 'center');
  // The look of the exception row decides which conditions its cells get.
  assert.deepEqual(tableConditions(doc, { table, row: 0, cell: 0 }), []);
  assert.deepEqual(tableConditions(doc, { table, row: 1, cell: 0 }), ['firstCol']);
  assert.deepEqual(resolveCellProps(doc, { table, row: 1, cell: 0 }).shading, { pattern: 'clear', fill: { rgb: 'BBBBBB' } });
  assert.equal(resolveCellProps(doc, { table, row: 0, cell: 0 }).shading, undefined);
});
