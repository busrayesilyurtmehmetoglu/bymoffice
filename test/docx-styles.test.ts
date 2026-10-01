import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  IMPLICIT_PARA_PROPS,
  IMPLICIT_RUN_PROPS,
  resolveCellProps,
  resolveColor,
  resolveFont,
  resolveParaProps,
  resolveRunProps,
  resolveTableProps,
  styleChain,
  tableConditions,
} from '../src/formats/docx/resolve';
import type { Paragraph, Run, Table } from '../src/model/doc/types';
import { readSpec, THEME } from './docx-fixtures';

const style = (id: string, type: string, inner: string, attrs = '') => `<w:style w:type="${type}" w:styleId="${id}"${attrs}>${inner}</w:style>`;

const STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/><w:sz w:val="22"/><w:lang w:val="tr-TR"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
  '<w:latentStyles w:defLockedState="0" w:count="2"><w:lsdException w:name="Normal" w:uiPriority="0"/></w:latentStyles>' +
  style('Normal', 'paragraph', '<w:name w:val="Normal"/><w:qFormat/>', ' w:default="1"') +
  style(
    'Heading1',
    'paragraph',
    '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:link w:val="Heading1Char"/><w:uiPriority w:val="9"/><w:qFormat/><w:rsid w:val="00AB"/>' +
      '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:ind w:left="432" w:hanging="432"/><w:outlineLvl w:val="0"/>' +
      '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:hAnsi="Georgia"/><w:b/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/></w:rPr>',
  ) +
  style('Title2', 'paragraph', '<w:basedOn w:val="Heading1"/><w:pPr><w:ind w:firstLine="200"/><w:tabs><w:tab w:val="clear" w:pos="720"/><w:tab w:val="right" w:pos="9000"/></w:tabs></w:pPr><w:rPr><w:b w:val="0"/><w:i/></w:rPr>') +
  style('DefaultParagraphFont', 'character', '<w:name w:val="Default Paragraph Font"/><w:semiHidden/><w:unhideWhenUsed/>', ' w:default="1"') +
  style('Heading1Char', 'character', '<w:link w:val="Heading1"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr>', ' w:customStyle="1"') +
  style('Strong', 'character', '<w:rPr><w:b/><w:caps/></w:rPr>') +
  style('Emphasis', 'character', '<w:basedOn w:val="Strong"/><w:rPr><w:i/><w:caps w:val="0"/></w:rPr>') +
  style('Gizli', 'character', '<w:hidden/><w:locked/><w:rPr><w:vanish/></w:rPr>') +
  style('LoopA', 'paragraph', '<w:basedOn w:val="LoopB"/><w:pPr><w:jc w:val="center"/></w:pPr>') +
  style('LoopB', 'paragraph', '<w:basedOn w:val="LoopA"/><w:pPr><w:keepLines/></w:pPr>') +
  style('Self', 'paragraph', '<w:basedOn w:val="Self"/><w:rPr><w:strike/></w:rPr>') +
  style('__proto__', 'paragraph', '<w:basedOn w:val="constructor"/><w:pPr><w:jc w:val="right"/></w:pPr>') +
  style('TableNormal', 'table', '<w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>', ' w:default="1"') +
  style(
    'Grid',
    'table',
    '<w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>' +
      '<w:tblPr><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/></w:tblBorders>' +
      '<w:tblCellMar><w:left w:w="50" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tcPr><w:vAlign w:val="center"/></w:tcPr>' +
      '<w:tblStylePr w:type="firstRow"><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="firstCol"><w:rPr><w:b/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:fill="FFFFFF"/><w:tcBorders><w:bottom w:val="double"/></w:tcBorders></w:tcPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="band1Vert"><w:rPr><w:u w:val="single"/></w:rPr></w:tblStylePr>' +
      '<w:tblStylePr w:type="nwCell"><w:rPr><w:caps/></w:rPr></w:tblStylePr>',
  ) +
  style('NoList', 'numbering', '<w:name w:val="No List"/>', ' w:default="1"') +
  // A second definition of an id is ignored.
  style('Strong', 'character', '<w:rPr><w:vanish/></w:rPr>') +
  '<w:style w:type="paragraph"><w:name w:val="kimliksiz"/></w:style>';

const cell = (text: string, props = '') => `<w:tc>${props}<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells: string) => `<w:tr>${cells}</w:tr>`;

test('docx styles: the style sheet is read', async () => {
  const doc = await readSpec({ styles: STYLES, body: '<w:p/>' });
  const styles = doc.styles;
  assert.deepEqual(styles.defaults.run, {
    fonts: { asciiTheme: 'minorHAnsi', hAnsiTheme: 'minorHAnsi', csTheme: 'minorBidi' },
    size: 22,
    lang: { val: 'tr-TR' },
  });
  assert.deepEqual(styles.defaults.para, { spacing: { after: 160, line: 259, lineRule: 'auto' } });
  assert.deepEqual(styles.defaultIds, { paragraph: 'Normal', character: 'DefaultParagraphFont', table: 'TableNormal', numbering: 'NoList' });
  assert.equal(styles.styles.size, 15);
  assert.equal(doc.warnings.length, 1);

  const heading = styles.styles.get('Heading1')!;
  assert.deepEqual(
    [heading.type, heading.name, heading.basedOn, heading.next, heading.link, heading.uiPriority, heading.quickFormat, heading.isDefault],
    ['paragraph', 'heading 1', 'Normal', 'Normal', 'Heading1Char', 9, true, false],
  );
  assert.equal(heading.para?.outlineLevel, 0);
  assert.equal(heading.run?.size, 32);
  assert.deepEqual(heading.extras?.map((extra) => [extra.index, extra.el.name]), [[6, 'w:rsid']]);

  assert.equal(styles.styles.get('Normal')!.isDefault, true);
  assert.equal(styles.styles.get('Heading1Char')!.custom, true);
  const hidden = styles.styles.get('Gizli')!;
  assert.deepEqual([hidden.hidden, hidden.locked, hidden.semiHidden], [true, true, undefined]);
  const fallback = styles.styles.get('DefaultParagraphFont')!;
  assert.deepEqual([fallback.semiHidden, fallback.unhideWhenUsed], [true, true]);
  // The first definition won.
  assert.equal(styles.styles.get('Strong')!.run?.vanish, undefined);

  const grid = styles.styles.get('Grid')!;
  assert.equal(grid.type, 'table');
  assert.equal(grid.table?.rowBandSize, 2);
  assert.deepEqual(grid.cell, { vAlign: 'center' });
  assert.deepEqual([...grid.conditions!.keys()], ['firstRow', 'lastRow', 'firstCol', 'band1Horz', 'band2Horz', 'band1Vert', 'nwCell']);
  assert.deepEqual(grid.conditions!.get('firstRow'), {
    para: { align: 'center' },
    run: { bold: true, color: { rgb: 'FFFFFF' } },
    cell: { shading: { pattern: 'clear', fill: { rgb: '4472C4' } } },
  });

  // An id that is a property of every object is an ordinary key.
  assert.equal(styles.styles.get('__proto__')!.para?.align, 'right');
  assert.equal(styles.styles.get('constructor'), undefined);
});

test('docx styles: chains follow basedOn and stop at loops', async () => {
  const { styles } = await readSpec({ styles: STYLES, body: '<w:p/>' });
  const ids = (id: string | undefined, type: 'paragraph' | 'character' | 'table' = 'paragraph') => styleChain(styles, id, type).map((s) => s.id);
  assert.deepEqual(ids('Title2'), ['Normal', 'Heading1', 'Title2']);
  assert.deepEqual(ids('LoopA'), ['LoopB', 'LoopA']);
  assert.deepEqual(ids('Self'), ['Self']);
  assert.deepEqual(ids('__proto__'), ['__proto__']);
  // No style, an unknown one, or one of another type: the default style of the type.
  assert.deepEqual(ids(undefined), ['Normal']);
  assert.deepEqual(ids('Yok'), ['Normal']);
  assert.deepEqual(ids('Strong'), ['Normal']);
  assert.deepEqual(ids('Emphasis', 'character'), ['Strong', 'Emphasis']);
  assert.deepEqual(ids('Grid', 'table'), ['TableNormal', 'Grid']);

  // A very long chain is cut, not followed forever.
  let long = '';
  for (let i = 0; i < 500; i++) long += style(`S${i}`, 'paragraph', i ? `<w:basedOn w:val="S${i - 1}"/>` : '');
  const many = await readSpec({ styles: long, body: '<w:p/>' });
  assert.equal(styleChain(many.styles, 'S499', 'paragraph').length, 64);
});

test('docx styles: effective paragraph properties', async () => {
  const doc = await readSpec({
    styles: STYLES,
    body:
      '<w:p/>' +
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Title2"/><w:spacing w:line="480"/><w:ind w:left="100"/><w:rPr><w:b/></w:rPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="LoopA"/></w:pPr></w:p>',
  });
  const props = (i: number) => resolveParaProps(doc, (doc.body.blocks[i] as Paragraph).props);
  // What no style states is what Word implies: left aligned, no space before.
  assert.deepEqual(props(0), { style: 'Normal', align: 'left', spacing: { before: 0, after: 160, line: 259, lineRule: 'auto' } });
  assert.deepEqual(props(1), {
    style: 'Heading1',
    align: 'left',
    spacing: { before: 240, after: 0, line: 259, lineRule: 'auto' },
    keepNext: true,
    indent: { left: 432, hanging: 432 },
    outlineLevel: 0,
    tabs: [
      { pos: 720, align: 'left' },
      { pos: 1440, align: 'left' },
    ],
  });
  assert.deepEqual(props(2), {
    style: 'Title2',
    align: 'left',
    // Attributes merge one by one; a first line indent replaces the inherited hanging indent.
    spacing: { before: 240, after: 0, line: 480, lineRule: 'auto' },
    keepNext: true,
    indent: { left: 100, firstLine: 200 },
    outlineLevel: 0,
    // The derived style cleared one tab stop and added another.
    tabs: [
      { pos: 1440, align: 'left' },
      { pos: 9000, align: 'right' },
    ],
    markProps: { bold: true },
  });
  assert.deepEqual(props(3), { style: 'LoopA', spacing: { before: 0, after: 160, line: 259, lineRule: 'auto' }, align: 'center', keepLines: true });
  // The model itself is not touched by resolving.
  assert.deepEqual((doc.body.blocks[1] as Paragraph).props, { style: 'Heading1' });
});

test('docx styles: effective run properties and the toggle rule', async () => {
  const run = (rPr: string) => `<w:r><w:rPr>${rPr}</w:rPr><w:t>x</w:t></w:r>`;
  const doc = await readSpec({
    styles: STYLES,
    body:
      `<w:p>${run('')}${run('<w:rFonts w:ascii="Arial"/><w:sz w:val="18"/>')}</w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run('')}${run('<w:rStyle w:val="Strong"/>')}${run('<w:rStyle w:val="Strong"/><w:b/>')}` +
      `${run('<w:rStyle w:val="Strong"/><w:b w:val="0"/>')}${run('<w:rStyle w:val="Emphasis"/>')}${run('<w:rStyle w:val="Yok"/>')}</w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Title2"/></w:pPr>${run('')}${run('<w:rStyle w:val="Strong"/>')}</w:p>`,
  });
  const resolve = (p: number, r: number) => {
    const paragraph = doc.body.blocks[p] as Paragraph;
    return resolveRunProps(doc, paragraph.props, (paragraph.children[r] as Run).props);
  };
  // This document has no theme part, so its theme font references lead nowhere and
  // every slot ends up with the implied font; the references themselves stay.
  const implied = { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', cs: 'Times New Roman' };
  const defaults = { fonts: { ...implied, asciiTheme: 'minorHAnsi', hAnsiTheme: 'minorHAnsi', csTheme: 'minorBidi' }, size: 22, sizeCs: 22, lang: { val: 'tr-TR' } };
  assert.deepEqual(resolve(0, 0), defaults);
  // A plain font name replaces the inherited theme font of the same slot only.
  assert.deepEqual(resolve(0, 1), { ...defaults, fonts: { ...implied, ascii: 'Arial', hAnsiTheme: 'minorHAnsi', csTheme: 'minorBidi' }, size: 18, sizeCs: 18 });

  const heading = resolve(1, 0);
  assert.deepEqual(heading, {
    fonts: { ...implied, asciiTheme: 'majorHAnsi', hAnsi: 'Georgia', csTheme: 'minorBidi' },
    size: 32,
    sizeCs: 32,
    lang: { val: 'tr-TR' },
    bold: true,
    color: { rgb: '2F5496', theme: 'accent1', shade: 0xbf },
  });
  // Bold in the paragraph style and bold in the character style cancel out.
  assert.equal(resolve(1, 1).bold, false);
  assert.equal(resolve(1, 1).caps, true);
  assert.equal(resolve(1, 1).style, 'Strong');
  // Direct formatting is absolute.
  assert.equal(resolve(1, 2).bold, true);
  assert.equal(resolve(1, 3).bold, false);
  // Inside one chain the derived style overrides: caps off, bold still from the base, and it toggles.
  assert.deepEqual([resolve(1, 4).bold, resolve(1, 4).caps, resolve(1, 4).italic], [false, false, true]);
  // An unknown character style formats nothing.
  assert.equal(resolve(1, 5).bold, true);

  // Title2 turns bold off in its own chain, so the character style turns it on.
  assert.deepEqual([resolve(2, 0).bold, resolve(2, 0).italic], [false, true]);
  assert.equal(resolve(2, 1).bold, true);
});

test('docx styles: table style conditions reach cells, paragraphs and runs', async () => {
  const doc = await readSpec({
    styles: STYLES,
    body:
      '<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
      '<w:tblBorders><w:top w:val="nil"/></w:tblBorders></w:tblPr>' +
      row(cell('a', '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>') + cell('b')) +
      row(cell('c') + cell('d') + cell('e')) +
      row(cell('f') + cell('g', '<w:tcPr><w:shd w:val="clear" w:fill="FF0000"/></w:tcPr>') + cell('h')) +
      row(cell('i') + cell('j') + cell('k')) +
      row(cell('l') + cell('m') + cell('n')) +
      '</w:tbl>' +
      `<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/><w:tblLook w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="0"/></w:tblPr>${row(cell('x') + cell('y'))}</w:tbl>` +
      `<w:tbl>${row(cell('z'))}</w:tbl>`,
  });
  const table = doc.body.blocks[0] as Table;
  const at = (r: number, c: number) => ({ table, row: r, cell: c });
  assert.deepEqual(tableConditions(doc, at(0, 0)), ['firstCol', 'firstRow', 'nwCell']);
  assert.deepEqual(tableConditions(doc, at(0, 1)), ['firstRow']);
  // Bands are two rows high and start under the header row.
  assert.deepEqual(tableConditions(doc, at(1, 1)), ['band1Horz']);
  assert.deepEqual(tableConditions(doc, at(2, 2)), ['band1Horz']);
  assert.deepEqual(tableConditions(doc, at(3, 0)), ['band2Horz', 'firstCol']);
  assert.deepEqual(tableConditions(doc, at(4, 1)), ['band2Horz', 'lastRow']);
  assert.deepEqual(tableConditions(doc, at(9, 9)), []);

  assert.deepEqual(resolveCellProps(doc, at(0, 0)), { vAlign: 'center', shading: { pattern: 'clear', fill: { rgb: '4472C4' } }, gridSpan: 2 });
  assert.deepEqual(resolveCellProps(doc, at(1, 0)).shading, { pattern: 'clear', fill: { rgb: 'D9E2F3' } });
  // The cell's own shading wins over the band.
  assert.deepEqual(resolveCellProps(doc, at(2, 1)).shading, { pattern: 'clear', fill: { rgb: 'FF0000' } });
  assert.deepEqual(resolveCellProps(doc, at(3, 1)).borders, { bottom: { style: 'double' } });

  const paragraph = (r: number, c: number) => table.rows[r].cells[c].blocks[0] as Paragraph;
  const head = paragraph(0, 0);
  assert.deepEqual(resolveParaProps(doc, head.props, at(0, 0)), { style: 'Normal', spacing: { before: 0, after: 0, line: 259, lineRule: 'auto' }, align: 'center' });
  const headRun = resolveRunProps(doc, head.props, (head.children[0] as Run).props, at(0, 0));
  // firstRow and firstCol both make it bold inside the table style: one level, so it stays bold.
  assert.deepEqual([headRun.bold, headRun.caps, headRun.size, headRun.color], [true, true, 20, { rgb: 'FFFFFF' }]);
  const last = paragraph(4, 1);
  assert.equal(resolveRunProps(doc, last.props, (last.children[0] as Run).props, at(4, 1)).italic, true);
  assert.equal(resolveRunProps(doc, last.props, (last.children[0] as Run).props, at(4, 1)).bold, undefined);
  // Outside a table none of it applies.
  assert.equal(resolveRunProps(doc, last.props, (last.children[0] as Run).props).size, 22);

  assert.deepEqual(resolveTableProps(doc, table), {
    style: 'Grid',
    cellMargins: { left: { type: 'dxa', value: 50 }, right: { type: 'dxa', value: 108 } },
    rowBandSize: 2,
    colBandSize: 1,
    borders: { top: { style: 'nil' }, insideH: { style: 'single', size: 4 } },
    look: { firstRow: true, lastRow: true, firstColumn: true, lastColumn: false, noHBand: false, noVBand: true, val: 0x04a0, stated: 0x07e0 },
  });

  const banded = doc.body.blocks[1] as Table;
  assert.deepEqual(tableConditions(doc, { table: banded, row: 0, cell: 0 }), ['band1Vert']);
  assert.deepEqual(tableConditions(doc, { table: banded, row: 0, cell: 1 }), ['band2Vert']);
  // A table without a style takes the default table style.
  assert.deepEqual(resolveTableProps(doc, doc.body.blocks[2] as Table).cellMargins, { left: { type: 'dxa', value: 108 }, right: { type: 'dxa', value: 108 } });
});

test('docx styles: theme colours and fonts, settings and the font table', async () => {
  const doc = await readSpec({
    theme: THEME,
    styles: STYLES,
    settings:
      '<w:zoom w:percent="120"/><w:trackRevisions/><w:defaultTabStop w:val="708"/><w:autoHyphenation w:val="0"/><w:evenAndOddHeaders/><w:mirrorMargins/>' +
      '<w:compat><w:useFELayout/><w:doNotExpandShiftReturn w:val="0"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' +
      '<w:themeFontLang w:val="tr-TR" w:eastAsia="ja-JP"/>' +
      '<w:clrSchemeMapping w:bg1="dark1" w:t1="light1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>',
    fontTable:
      '<w:font w:name="Calibri"><w:panose1 w:val="020F0502020204030204"/><w:charset w:val="A2"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>' +
      '<w:sig w:usb0="E4002EFF" w:usb1="C200247B"/><w:embedRegular r:id="rIdFont1" w:fontKey="{0}"/><w:embedBold r:id="rIdFont2"/></w:font>' +
      '<w:font w:name="Yazı"><w:altName w:val="Arial"/><w:family w:val="auto"/></w:font><w:font w:name="Calibri"><w:family w:val="roman"/></w:font><w:font/>',
    body: '<w:p/>',
  });
  assert.deepEqual([...doc.theme.colors], [
    ['lt1', 'FFFFFF'],
    ['dk1', '000000'],
    ['lt2', 'E7E6E6'],
    ['dk2', '44546A'],
    ['accent1', '4472C4'],
    ['accent2', 'ED7D31'],
    ['accent3', 'A5A5A5'],
    ['accent4', 'FFC000'],
    ['accent5', '5B9BD5'],
    ['accent6', '70AD47'],
    ['hlink', '0563C1'],
    ['folHlink', '954F72'],
  ]);
  assert.equal(doc.theme.majorFont.latin, 'Calibri Light');
  assert.equal(doc.theme.majorFont.scripts.get('Jpan'), 'Yu Gothic Light');
  assert.equal(doc.theme.minorFont.complexScript, 'Arial');

  assert.equal(resolveColor(doc, { rgb: '123ABC' }), '#123ABC');
  assert.equal(resolveColor(doc, { auto: true }), undefined);
  assert.equal(resolveColor(doc, undefined), undefined);
  // The theme is the source; the value next to it is only what the writer computed.
  assert.equal(resolveColor(doc, { rgb: '000000', theme: 'accent2' }), '#ED7D31');
  assert.equal(resolveColor(doc, { theme: 'hyperlink' }), '#0563C1');
  // This document swaps text and background.
  assert.equal(resolveColor(doc, { theme: 'text1' }), '#FFFFFF');
  assert.equal(resolveColor(doc, { theme: 'background1' }), '#000000');
  assert.equal(resolveColor(doc, { theme: 'dark2' }), '#44546A');
  assert.equal(resolveColor(doc, { rgb: 'ABCDEF', theme: 'neYazikKi' }), '#ABCDEF');
  // Shade darkens, tint lightens; 255 changes nothing.
  // The expected values are the ones Word itself writes next to these theme references.
  assert.equal(resolveColor(doc, { theme: 'accent1', shade: 0xbf }), '#2F5496');
  assert.equal(resolveColor(doc, { theme: 'accent1', tint: 0x99 }), '#8EAADB');
  assert.equal(resolveColor(doc, { theme: 'accent1', tint: 0x66 }), '#B4C6E7');
  assert.equal(resolveColor(doc, { theme: 'accent1', tint: 0x33 }), '#D9E2F3');
  assert.equal(resolveColor(doc, { theme: 'accent1', shade: 0x80 }), '#1F3864');
  assert.equal(resolveColor(doc, { theme: 'accent1', tint: 255, shade: 255 }), '#4472C4');
  assert.equal(resolveColor(doc, { theme: 'dark1', tint: 0x80 }), '#7F7F7F');

  assert.equal(resolveFont(doc, { asciiTheme: 'majorHAnsi', ascii: 'Arial' }, 'ascii'), 'Calibri Light');
  assert.equal(resolveFont(doc, { hAnsiTheme: 'minorHAnsi' }, 'hAnsi'), 'Calibri');
  assert.equal(resolveFont(doc, { csTheme: 'minorBidi' }, 'cs'), 'Arial');
  // The theme names no East Asian font: the plain name is used.
  assert.equal(resolveFont(doc, { eastAsiaTheme: 'minorEastAsia', eastAsia: 'MS Mincho' }, 'eastAsia'), 'MS Mincho');
  assert.equal(resolveFont(doc, { ascii: 'Arial' }, 'cs'), undefined);
  assert.equal(resolveFont(doc, undefined, 'ascii'), undefined);

  const settings = doc.settings;
  assert.deepEqual(
    [settings.defaultTabStop, settings.evenAndOddHeaders, settings.mirrorMargins, settings.autoHyphenation, settings.trackRevisions],
    [708, true, true, false, true],
  );
  assert.deepEqual(settings.themeFontLang, { val: 'tr-TR', eastAsia: 'ja-JP' });
  assert.deepEqual([...settings.compat], [
    ['useFELayout', true],
    ['doNotExpandShiftReturn', false],
  ]);
  assert.deepEqual([...settings.compatSettings], [['compatibilityMode', '15']]);

  assert.deepEqual([...doc.fonts.keys()], ['Calibri', 'Yazı']);
  assert.deepEqual(doc.fonts.get('Calibri'), {
    name: 'Calibri',
    panose: '020F0502020204030204',
    charset: 'A2',
    family: 'swiss',
    pitch: 'variable',
    embedded: new Map([
      ['regular', 'rIdFont1'],
      ['bold', 'rIdFont2'],
    ]),
  });
  assert.deepEqual(doc.fonts.get('Yazı'), { name: 'Yazı', altName: 'Arial', family: 'auto' });

  // Without the parts everything has its default.
  const bare = await readSpec({ body: '<w:p><w:r><w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr></w:r></w:p>' });
  assert.equal(bare.settings.defaultTabStop, 720);
  assert.equal(bare.theme.colors.size, 0);
  assert.equal(bare.fonts.size, 0);
  assert.equal(bare.styles.styles.size, 0);
  const run = (bare.body.blocks[0] as Paragraph).children[0] as Run;
  assert.equal(resolveColor(bare, run.props.color), '#FF0000');
  assert.deepEqual(resolveRunProps(bare, {}, run.props), { ...IMPLICIT_RUN_PROPS, sizeCs: 20, color: { rgb: 'FF0000', theme: 'accent1' } });
  assert.deepEqual(resolveParaProps(bare, {}), IMPLICIT_PARA_PROPS);
});
