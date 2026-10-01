// Cases found by review: each one was a defect before it became a test.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseSharedStrings } from '../src/formats/xlsx/shared-strings';
import { parseStyles } from '../src/formats/xlsx/styles';
import { parseWorksheet } from '../src/formats/xlsx/worksheet';
import { XmlScanner } from '../src/xml-scan';
import { DEFAULT_INDEXED_COLORS, createSheet, emptyStyleSheet, formatCodeOf, getCell, open, parseCellRef, parseXml } from '../src/index';
import { buildXlsx, utf8, worksheetXml } from './fixtures';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sheetOf = (inner: string, strings = parseSharedStrings(utf8(`<sst xmlns="${NS}"/>`))) => {
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '/xl/worksheets/sheet1.xml' });
  parseWorksheet(utf8(worksheetXml(inner)), strings, sheet);
  return sheet;
};
const cellXml = (v: string, attrs = '') => `<sheetData><row r="1"><c r="A1"${attrs}>${v}</c></row></sheetData>`;

test('xml: entity names are not looked up on Object.prototype', () => {
  for (const name of ['constructor', 'toString', 'valueOf', 'hasOwnProperty']) {
    assert.throws(() => parseXml(`<a>&${name};</a>`), /Unknown XML entity/, name);
    assert.throws(() => parseXml(`<a b="&${name};"/>`), /Unknown XML entity/, name);
    assert.throws(() => {
      const scanner = new XmlScanner(`<a b="&${name};">&${name};</a>`);
      while (scanner.next()) if (scanner.kind === 'open') scanner.attr('b');
    }, /Unknown XML entity/);
  }
});

test('xml-scan: agrees with parseXml on empty processing instructions and text after the root', () => {
  const drain = (xml: string) => {
    const scanner = new XmlScanner(xml);
    while (scanner.next());
  };
  assert.throws(() => parseXml('<a><?>x</a>'), /empty processing instruction/);
  assert.throws(() => drain('<a><?>x</a>'), /empty processing instruction/);
  assert.doesNotThrow(() => parseXml('<a/> &#32; '));
  assert.doesNotThrow(() => drain('<a/> &#32; '));
  assert.throws(() => drain('<a/> &#65; '), /text outside the root/);
});

test('xlsx: underline values are not looked up on Object.prototype', () => {
  for (const value of ['constructor', 'toString', 'valueOf']) {
    const [item] = parseSharedStrings(utf8(`<sst xmlns="${NS}"><si><r><rPr><u val="${value}"/></rPr><t>a</t></r></si></sst>`));
    assert.equal(item.runs?.[0].font?.underline, 'single', value);
  }
});

test('xlsx: plain text after rich runs stays part of the string', () => {
  const [item] = parseSharedStrings(utf8(`<sst xmlns="${NS}"><si><r><rPr><b/></rPr><t>kalın</t></r><t> düz</t></si></sst>`));
  assert.equal(item.text, 'kalın düz');
  assert.deepEqual(item.runs?.map((run) => run.text), ['kalın', ' düz']);
});

test('xlsx: numbers follow xsd:double, not JavaScript literals', () => {
  for (const good of ['1', '-1.5', '+2', '.5', '5.', '1e3', '1.5E-3', ' 7 ']) {
    assert.equal(getCell(sheetOf(cellXml(`<v>${good}</v>`)), 0, 0)?.value, Number(good), good);
  }
  for (const bad of ['0x1F', '0b11', '0o7', 'NaN', 'INF', 'Infinity', '1e999', '1,5', 'abc', '1_000']) {
    assert.throws(() => sheetOf(cellXml(`<v>${bad}</v>`)), /is not a number/, bad);
  }
  // The same rule for sizes and splits: a malformed value is ignored, not reinterpreted.
  const sheet = sheetOf(
    '<sheetViews><sheetView workbookViewId="0" zoomScale="0x50"><pane ySplit="0x10" xSplit="2" state="frozen"/></sheetView></sheetViews>' +
      '<cols><col min="1" max="1" width="0x20"/></cols><sheetData><row r="1" ht="0b11"/></sheetData>',
  );
  assert.deepEqual([sheet.view.zoomScale, sheet.view.frozenRows, sheet.view.frozenCols], [100, 0, 2]);
  assert.equal(sheet.columns[0].width, undefined);
  assert.equal(sheet.rows.get(0)?.height, undefined);

  const styles = parseStyles(parseXml(`<styleSheet xmlns="${NS}"><fonts><font><sz val="0x10"/></font></fonts></styleSheet>`), undefined);
  assert.equal(styles.fonts[0].size, undefined);
});

test('xlsx: error messages quote file content only up to a short length', () => {
  const long = '9'.repeat(100000);
  for (const xml of [cellXml(`<v>${long}x</v>`), cellXml(`<v>${long}</v>`, ' t="s"'), `<sheetData><row r="${long}"/></sheetData>`]) {
    assert.throws(
      () => sheetOf(xml),
      (error: Error) => error.message.length < 200,
    );
  }
  assert.throws(() => sheetOf(`<sheetData><row r="1"><c r="${'A'.repeat(100000)}"/></row></sheetData>`), (error: Error) => error.message.length < 200);
});

test('xlsx: formula text is literal, cell text decodes _xHHHH_ escapes', () => {
  const sheet = sheetOf(
    '<sheetData><row r="1"><c r="A1" t="str"><f>"_x0041_"&amp;Tbl[_x0042_col]</f><v>satır_x000A_sonu _x005F_x0041_</v></c></row></sheetData>',
  );
  const cell = getCell(sheet, 0, 0);
  assert.equal(cell?.formula?.text, '"_x0041_"&Tbl[_x0042_col]');
  assert.equal(cell?.value, 'satır\nsonu _x0041_');
});

test('xlsx: boolean cells accept both spellings', () => {
  for (const [text, expected] of [['1', true], ['true', true], ['0', false], ['false', false]] as const) {
    assert.equal(getCell(sheetOf(cellXml(`<v>${text}</v>`, ' t="b"')), 0, 0)?.value, expected, text);
  }
});

test('xlsx: panes frozen while scrolled keep where the frozen area starts', () => {
  const sheet = sheetOf(
    '<sheetViews><sheetView workbookViewId="0" topLeftCell="C5"><pane xSplit="1" ySplit="1" topLeftCell="D6" state="frozen"/></sheetView></sheetViews><sheetData/>',
  );
  assert.deepEqual(
    [sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.firstVisibleCell, sheet.view.topLeftCell],
    [1, 1, 'C5', 'D6'],
  );
  const huge = sheetOf('<sheetViews><sheetView workbookViewId="0"><pane xSplit="99999" ySplit="9999999" state="frozen"/></sheetView></sheetViews><sheetData/>');
  assert.deepEqual([huge.view.frozenCols, huge.view.frozenRows], [16383, 1048575]);
});

test('xlsx: a part listed as two sheets is rejected instead of being read twice', async () => {
  const bytes = await buildXlsx({ sheets: [{ name: 'A', xml: worksheetXml('<sheetData/>') }] });
  const file = await open(bytes);
  const workbook = await file.package!.readXml('/xl/workbook.xml');
  const sheets = workbook.root.children.find((node) => node.type === 'element' && node.name === 'sheets');
  assert.ok(sheets && sheets.type === 'element');
  sheets.children.push({ ...(sheets.children[0] as object) } as (typeof sheets.children)[0]);
  file.package!.writePart('/xl/workbook.xml', workbook);
  const again = await open(await file.save());
  await assert.rejects(again.workbook(), /shares its part with another sheet/);
});

test('xlsx: activeTab and localSheetId must be plain non-negative integers', async () => {
  for (const value of ['1.5', '-1', 'abc', '1e0', '']) {
    const file = await open(
      await buildXlsx({
        workbookHead: `<bookViews><workbookView activeTab="${value}"/></bookViews>`,
        workbookTail: `<definedNames><definedName name="N" localSheetId="${value}">A!$A$1</definedName></definedNames>`,
        sheets: [
          { name: 'A', xml: worksheetXml('<sheetData/>') },
          { name: 'B', xml: worksheetXml('<sheetData/>') },
        ],
      }),
    );
    const workbook = await file.workbook();
    assert.equal(workbook.activeSheet, 0, value);
    assert.equal(workbook.definedNames[0].localSheet, undefined, value);
  }
});

test('xlsx: strict OOXML namespaces read the same as transitional ones', async () => {
  const inner = '<sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData>';
  const transitional = await open(await buildXlsx({ sheets: [{ name: 'A', xml: worksheetXml(inner) }] }));
  const strict = await open(await buildXlsx({ sheets: [{ name: 'A', xml: worksheetXml(inner) }] }));
  const swap = (text: string) =>
    text
      .replace(/http:\/\/schemas\.openxmlformats\.org\/spreadsheetml\/2006\/main/g, 'http://purl.oclc.org/ooxml/spreadsheetml/main')
      .replace(/http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships/g, 'http://purl.oclc.org/ooxml/officeDocument/relationships');
  for (const part of ['/_rels/.rels', '/xl/workbook.xml', '/xl/_rels/workbook.xml.rels', '/xl/worksheets/sheet1.xml']) {
    strict.package!.writePart(part, utf8(swap(new TextDecoder().decode(await strict.package!.readPart(part)))));
  }
  const reopened = await open(await strict.save());
  assert.deepEqual((await reopened.workbook()).sheets[0].rows, (await transitional.workbook()).sheets[0].rows);
});

test('model: reference parsing limits and the full default palette', () => {
  assert.equal(parseCellRef('A0'), undefined);
  assert.equal(parseCellRef('A'), undefined);
  assert.equal(parseCellRef('1'), undefined);
  assert.equal(parseCellRef('XFE1'), undefined);
  assert.equal(parseCellRef('A1048577'), undefined);
  assert.deepEqual(parseCellRef('XFD1048576'), { row: 1048575, col: 16383 });
  assert.deepEqual(parseCellRef('$b$7'), { row: 6, col: 1 });

  // ECMA-376 18.8.27, all 64 entries.
  assert.equal(
    DEFAULT_INDEXED_COLORS.join(' '),
    '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
      '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
      '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
      '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333',
  );
  const styles = emptyStyleSheet();
  assert.deepEqual(
    [37, 38, 39, 40].map((id) => formatCodeOf(styles, id)),
    ['#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)'],
  );
});
