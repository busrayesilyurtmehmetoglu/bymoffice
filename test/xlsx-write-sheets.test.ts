// Saving a workbook whose sheets were inserted, deleted or renamed. Every scenario applies the
// operations through a History, saves, opens the result with the reader and compares the two
// models; the package itself is checked part by part (relationships, content types, orphans).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { isDeepStrictEqual } from 'node:util';
import { writeXlsx } from '../src/formats/xlsx/write';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { CalcEngine, ConditionalFormats, DataValidations, ZipArchive, createSheet, emptyStyleSheet, open, parseCellRef, parseXml } from '../src/index';
import type { OfficeFile, Workbook, XmlElement } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { mergeRange, setColumnWidth } from '../src/ops/format';
import { History, type Op } from '../src/ops/ops';
import { deleteSheet, insertSheet, renameSheet } from '../src/ops/sheets';
import { buildXlsx, utf8, worksheetXml } from './fixtures';
import { comment, commentsXml, drawingXml, pic, sp, twoCell, vmlNote, vmlXml } from './sheet-drawing-fixtures';

// --- helpers ---------------------------------------------------------------

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_STRICT = 'http://purl.oclc.org/ooxml/spreadsheetml/main';
const NS_REL_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships';
const decoder = new TextDecoder();

interface Loaded {
  bytes: Uint8Array;
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { bytes, file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const reopen = async (bytes: Uint8Array) => (await open(bytes)).workbook();
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

function type(doc: Loaded, sheet: number, ref: string, text: string): void {
  const { row, col } = parseCellRef(ref)!;
  apply(doc, setCellInput(doc.wb, sheet, row, col, text));
}

const names = (wb: Workbook) => wb.sheets.map((sheet) => sheet.name);

function formulaOf(wb: Workbook, sheet: number, ref: string): string | undefined {
  const { row, col } = parseCellRef(ref)!;
  const cell = wb.sheets[sheet].rows.get(row)?.cells[col];
  return cell?.formula?.text;
}

function valueOf(wb: Workbook, sheet: number, ref: string): unknown {
  const { row, col } = parseCellRef(ref)!;
  return wb.sheets[sheet].rows.get(row)?.cells[col]?.value;
}

/** What a saved file has to bring back of the model. */
function picture(wb: Workbook): unknown {
  return {
    active: wb.activeSheet,
    names: wb.definedNames.map((name) => [name.name, name.value, name.localSheet === undefined ? null : name.localSheet, name.hidden]),
    sheets: wb.sheets.map((sheet) => ({
      name: sheet.name,
      state: sheet.state,
      kind: sheet.kind,
      merges: sheet.merges,
      columns: sheet.columns.map((range) => [range.min, range.max, range.width === undefined ? null : range.width]),
      links: sheet.hyperlinks.map((link) => [link.range, link.location === undefined ? null : link.location, link.target === undefined ? null : link.target]),
      rows: [...sheet.rows.entries()]
        .sort((a, b) => a[0] - b[0])
        .filter(([, row]) => row.cells.some(() => true))
        .map(([index, row]) => {
          const cells: unknown[] = [];
          row.cells.forEach((cell, col) => cells.push([col, cell.kind, cell.value === undefined ? null : cell.value, cell.formula ? cell.formula.text : null]));
          return [index, cells];
        }),
    })),
  };
}

function same(actual: unknown, expected: unknown, label: string): void {
  if (isDeepStrictEqual(actual, expected)) return;
  const a = JSON.stringify(actual);
  const b = JSON.stringify(expected);
  let at = 0;
  while (at < a.length && at < b.length && a.charCodeAt(at) === b.charCodeAt(at)) at++;
  assert.fail(`${label}: differs at ${at}\n  actual:   ${a.slice(Math.max(0, at - 80), at + 120)}\n  expected: ${b.slice(Math.max(0, at - 80), at + 120)}`);
}

async function content(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = ZipArchive.read(bytes);
  const out = new Map<string, Uint8Array>();
  for (const name of zip.names()) out.set(name, await zip.read(name));
  return out;
}

const text = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));

const sameBytes = (a: Uint8Array | undefined, b: Uint8Array | undefined) => !!a && !!b && a.length === b.length && a.every((value, i) => value === b[i]);

/** The names of the entries whose content is not what it was, added and removed ones included. */
async function changedParts(before: Uint8Array, after: Uint8Array): Promise<string[]> {
  const a = await content(before);
  const b = await content(after);
  const out: string[] = [];
  for (const name of new Set([...a.keys(), ...b.keys()])) if (!sameBytes(a.get(name), b.get(name))) out.push(name);
  return out.sort();
}

const all = (el: XmlElement, local: string, out: XmlElement[] = []): XmlElement[] => {
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    if (node.name.replace(/^.*:/, '') === local) out.push(node);
    all(node, local, out);
  }
  return out;
};

function resolve(source: string, target: string): string {
  const base = target.startsWith('/') ? target.slice(1) : `${source.slice(0, source.lastIndexOf('/') + 1)}${target}`;
  const out: string[] = [];
  for (const segment of base.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}

interface PackageFacts {
  parts: string[];
  /** Per sheet of workbook.xml, in order. */
  sheets: { name: string; sheetId: string; part: string; tabSelected: boolean; state: string }[];
  activeTab: number;
}

/**
 * Checks that the package holds together: every sheet has a relationship and every relationship
 * a part; every part has a content type and every override a part; no sheet, drawing, comment,
 * picture or relationships part is left that nothing leads to; one sheet is selected, the active one.
 */
async function checkPackage(bytes: Uint8Array): Promise<PackageFacts> {
  const parts = await content(bytes);
  const partNames = [...parts.keys()].filter((name) => name !== '[Content_Types].xml');
  const lower = new Set(partNames.map((name) => name.toLowerCase()));
  assert.equal(lower.size, partNames.length, 'part names differ by more than case');

  const types = parseXml(parts.get('[Content_Types].xml')!).root;
  const overrides = new Map<string, string>();
  for (const el of all(types, 'Override')) {
    const name = el.attrs.get('PartName')!;
    assert.ok(!overrides.has(name.toLowerCase()), `two content types for ${name}`);
    overrides.set(name.toLowerCase(), el.attrs.get('ContentType')!);
    assert.ok(lower.has(name.slice(1).toLowerCase()), `content type for a part that is not there: ${name}`);
  }
  const defaults = new Set(all(types, 'Default').map((el) => el.attrs.get('Extension')!.toLowerCase()));
  for (const name of partNames) {
    assert.ok(overrides.has(`/${name}`.toLowerCase()) || defaults.has(name.slice(name.lastIndexOf('.') + 1).toLowerCase()), `no content type for ${name}`);
  }

  // Relationships: each belongs to a part that is there, and leads to one that is.
  const targets = new Set<string>();
  const relsOf = new Map<string, Map<string, { type: string; part: string }>>();
  for (const name of partNames) {
    const match = /^(.*\/)?_rels\/([^/]*)\.rels$/.exec(name);
    if (!match) continue;
    const source = (match[1] || '') + match[2];
    if (source !== '') assert.ok(parts.has(source), `relationships of a part that is not there: ${name}`);
    const byId = new Map<string, { type: string; part: string }>();
    for (const rel of all(parseXml(parts.get(name)!).root, 'Relationship')) {
      const id = rel.attrs.get('Id')!;
      assert.ok(!byId.has(id), `${name}: relationship id ${id} twice`);
      if (rel.attrs.get('TargetMode') === 'External') continue;
      // A place in the workbook, not a part.
      if (rel.attrs.get('Target')!.startsWith('#')) continue;
      const part = resolve(source, rel.attrs.get('Target')!);
      assert.ok(parts.has(part), `${name}: ${id} leads to ${part}, which is not there`);
      byId.set(id, { type: rel.attrs.get('Type')!, part });
      targets.add(part);
    }
    relsOf.set(source, byId);
  }
  for (const name of partNames) {
    if (/^xl\/(worksheets|chartsheets|drawings|media|printerSettings|threadedComments)\/[^/]+$/.test(name) || /^xl\/comments\d*\.xml$/.test(name) || name === 'xl/calcChain.xml') {
      assert.ok(targets.has(name), `nothing leads to ${name}`);
    }
  }

  const book = parseXml(parts.get('xl/workbook.xml')!).root;
  const bookRels = relsOf.get('xl/workbook.xml')!;
  const sheets: PackageFacts['sheets'] = [];
  const seenParts = new Set<string>();
  const seenIds = new Set<string>();
  const seenNames = new Set<string>();
  for (const el of all(book, 'sheet')) {
    let id: string | undefined;
    for (const [name, value] of el.attrs) if (name.includes(':') && name.endsWith(':id')) id = value;
    const rel = bookRels.get(id ?? '');
    assert.ok(rel, `sheet "${el.attrs.get('name')}" has no relationship (${id})`);
    assert.ok(/\/(worksheet|chartsheet|dialogsheet|macrosheet)$/.test(rel.type), `sheet relationship of type ${rel.type}`);
    assert.ok(!seenParts.has(rel.part), `two sheets in ${rel.part}`);
    seenParts.add(rel.part);
    const sheetId = el.attrs.get('sheetId')!;
    assert.ok(/^[1-9]\d*$/.test(sheetId) && !seenIds.has(sheetId), `sheetId ${sheetId} is not a number of its own`);
    seenIds.add(sheetId);
    const name = el.attrs.get('name')!;
    assert.ok(!seenNames.has(name.toLowerCase()), `two sheets called ${name}`);
    seenNames.add(name.toLowerCase());
    const views = all(parseXml(parts.get(rel.part)!).root, 'sheetView');
    sheets.push({ name, sheetId, part: rel.part, tabSelected: views.some((view) => view.attrs.get('tabSelected') === '1'), state: el.attrs.get('state') ?? 'visible' });
  }
  assert.ok(sheets.length > 0, 'no sheets');
  // Every sheet part is a sheet of the workbook.
  for (const name of partNames) if (/^xl\/worksheets\/[^/]+$/.test(name)) assert.ok(seenParts.has(name), `${name} is no sheet of the workbook`);

  const view = all(book, 'workbookView')[0];
  const activeTab = Number(view?.attrs.get('activeTab') ?? '0');
  assert.ok(activeTab >= 0 && activeTab < sheets.length, `activeTab ${activeTab} of ${sheets.length} sheets`);
  assert.equal(sheets[activeTab].state, 'visible', 'the active sheet is hidden');
  const firstSheet = Number(view?.attrs.get('firstSheet') ?? '0');
  assert.ok(firstSheet >= 0 && firstSheet < sheets.length, `firstSheet ${firstSheet} of ${sheets.length} sheets`);
  for (const name of all(book, 'definedName')) {
    const local = name.attrs.get('localSheetId');
    if (local !== undefined) assert.ok(Number(local) < sheets.length, `localSheetId ${local} of ${sheets.length} sheets`);
  }
  return { parts: partNames.sort(), sheets, activeTab };
}

/** One sheet is marked as selected: the active one. */
function checkSelection(facts: PackageFacts): void {
  assert.deepEqual(
    facts.sheets.map((sheet) => sheet.tabSelected),
    facts.sheets.map((_sheet, index) => index === facts.activeTab),
    'tabSelected',
  );
}

/** Saves, checks the package, opens it again and compares the model that comes back. */
async function roundTrip(doc: Loaded, selection = true): Promise<{ bytes: Uint8Array; facts: PackageFacts; wb: Workbook }> {
  const bytes = await save(doc);
  const facts = await checkPackage(bytes);
  if (selection) checkSelection(facts);
  const wb = await reopen(bytes);
  same(picture(wb), picture(doc.wb), 'the model read back');
  assert.deepEqual(
    facts.sheets.map((sheet) => sheet.name),
    names(doc.wb),
  );
  return { bytes, facts, wb };
}

// --- the workbook of most tests ------------------------------------------------

const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
const formula = (r: string, f: string, v: number) => `<c r="${r}"><f>${f}</f><v>${v}</v></c>`;
const expression = (sqref: string, f: string, priority: number) =>
  `<conditionalFormatting sqref="${sqref}"><cfRule type="expression" dxfId="0" priority="${priority}"><formula>${f}</formula></cfRule></conditionalFormatting>`;

const STYLES =
  `${XML}<styleSheet xmlns="${NS}"><fonts count="1"><font/></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>` +
  '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>' +
  '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs></styleSheet>';

const X14 =
  '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>' +
  '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="9" id="{00000000-0000-0000-0000-000000000001}">' +
  '<xm:f>Sheet3!$A$2&gt;Sheet2!$A$1</xm:f><x14:dxf><fill><patternFill><bgColor rgb="FF00FF00"/></patternFill></fill></x14:dxf></x14:cfRule><xm:sqref>D1</xm:sqref></x14:conditionalFormatting>' +
  '</x14:conditionalFormattings></ext></extLst>';

const APP =
  `${XML}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
  '<Application>Test</Application><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant>' +
  '<vt:variant><vt:lpstr>Named Ranges</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
  '<TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr><vt:lpstr>Sheet2</vt:lpstr><vt:lpstr>Sheet3</vt:lpstr><vt:lpstr>Global</vt:lpstr></vt:vector></TitlesOfParts></Properties>';

const relsXml = (inner: string) => `${XML}<Relationships xmlns="${NS_PKG_REL}">${inner}</Relationships>`;
const rel = (id: string, kind: string, target: string, external = false) => `<Relationship Id="${id}" Type="${NS_REL}/${kind}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;

interface BookOptions {
  /** Leaves out the formulas and names that refer to Sheet2, so that deleting it changes no cell. */
  loneSheet2?: boolean;
  calcChain?: boolean;
}

/**
 * Three sheets. "Sheet1" (selected, with a picture and a shape that shows Sheet3!$A$2) has rules
 * and validations that read the other two. "Sheet2" has a comment, a link, printer settings and
 * a drawing with two pictures, one of which Sheet1 shows as well. "Sheet3" has plain cells.
 */
async function book(options: BookOptions = {}): Promise<Uint8Array> {
  const lone = !!options.loneSheet2;
  const bytes = await buildXlsx({
    styles: STYLES,
    workbookHead: '<bookViews><workbookView xWindow="10" activeTab="0" firstSheet="0"/></bookViews>',
    workbookTail:
      '<definedNames><definedName name="Global">Sheet3!$A$1:$A$2</definedName>' +
      '<definedName name="Local2" localSheetId="1">Sheet2!$A$1</definedName>' +
      '<definedName name="Local3" localSheetId="2" comment="kept">Sheet3!$A$2</definedName>' +
      (lone ? '' : '<definedName name="Two">Sheet2!$A$1</definedName>') +
      '</definedNames><calcPr calcId="191029"/>',
    sheets: [
      {
        name: 'Sheet1',
        xml: worksheetXml(
          '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>' +
            `<sheetData><row r="1">${num('A1', 1)}${lone ? num('B1', 12) : formula('B1', 'Sheet2!A1+Sheet3!A2', 12)}${num('C1', 1)}${num('D1', 9)}</row></sheetData>` +
            expression('A1', 'Sheet3!$A$2&gt;0', 1) +
            (lone ? '' : expression('B1', 'Sheet2!$A$1&gt;0', 2)) +
            '<dataValidations count="2">' +
            `<dataValidation type="whole" sqref="A5" showErrorMessage="1"><formula1>Sheet3!$A$2</formula1><formula2>${lone ? '99' : 'Sheet2!$A$1'}</formula2></dataValidation>` +
            `<dataValidation type="list" sqref="B5" showErrorMessage="1"><formula1>${lone ? 'Sheet3' : 'Sheet2'}!$A$1:$A$2</formula1></dataValidation>` +
            '</dataValidations><drawing r:id="rIdD"/>' +
            (lone ? '' : X14),
        ),
      },
      {
        name: 'Sheet2',
        xml: worksheetXml(
          '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
            `<sheetData><row r="1">${num('A1', 7)}${formula('B1', 'Sheet3!A1+1', 1)}</row><row r="2">${num('A2', 8)}</row></sheetData>` +
            '<hyperlinks><hyperlink ref="C1" r:id="rIdH"/></hyperlinks><pageSetup orientation="landscape" r:id="rIdP"/><drawing r:id="rIdD"/><legacyDrawing r:id="rIdV"/>',
        ),
      },
      {
        name: 'Sheet3',
        xml: worksheetXml(
          '<sheetViews><sheetView workbookViewId="0"/></sheetViews>' +
            `<sheetData><row r="1">${num('A1', 0)}${formula('B1', 'Sheet1!A1*2', 2)}${lone ? '' : formula('C1', 'SUM(Sheet2!A1:A2)', 15)}</row><row r="2">${num('A2', 5)}</row></sheetData>`,
        ),
      },
    ],
  });
  const zip = ZipArchive.read(bytes);
  const shape = sp({ body: '<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p/></xdr:txBody>' }).replace('<xdr:sp>', '<xdr:sp macro="" textlink="Sheet3!$A$2">');
  zip.set('xl/worksheets/_rels/sheet1.xml.rels', utf8(relsXml(rel('rIdD', 'drawing', '../drawings/drawing1.xml'))));
  zip.set('xl/drawings/drawing1.xml', utf8(drawingXml(twoCell([1, 1], [3, 4], pic('rId1')) + twoCell([4, 1], [6, 4], shape))));
  zip.set('xl/drawings/_rels/drawing1.xml.rels', utf8(relsXml(rel('rId1', 'image', '../media/image1.png'))));
  zip.set(
    'xl/worksheets/_rels/sheet2.xml.rels',
    utf8(
      relsXml(
        rel('rIdH', 'hyperlink', 'https://example.com/', true) +
          rel('rIdP', 'printerSettings', '../printerSettings/printerSettings1.bin') +
          rel('rIdD', 'drawing', '../drawings/drawing2.xml') +
          rel('rIdV', 'vmlDrawing', '../drawings/vmlDrawing1.vml') +
          rel('rIdC', 'comments', '../comments1.xml'),
      ),
    ),
  );
  zip.set('xl/drawings/drawing2.xml', utf8(drawingXml(twoCell([1, 1], [3, 4], pic('rId1')) + twoCell([4, 1], [6, 4], pic('rId2')))));
  zip.set('xl/drawings/_rels/drawing2.xml.rels', utf8(relsXml(rel('rId1', 'image', '../media/image1.png') + rel('rId2', 'image', '../media/image2.png'))));
  zip.set('xl/media/image1.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 1, 1]));
  zip.set('xl/media/image2.png', new Uint8Array([0x89, 0x50, 0x4e, 0x47, 2, 2, 2, 2]));
  zip.set('xl/drawings/vmlDrawing1.vml', utf8(vmlXml(vmlNote(1, 1))));
  zip.set('xl/comments1.xml', utf8(commentsXml(['A'], comment('B2', 0, '<t>note</t>'))));
  zip.set('xl/printerSettings/printerSettings1.bin', new Uint8Array([9, 8, 7, 6, 5]));
  zip.set('docProps/app.xml', utf8(APP));
  zip.set(
    '_rels/.rels',
    utf8(relsXml(rel('rId1', 'officeDocument', 'xl/workbook.xml') + '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>')),
  );
  let extraTypes =
    '<Default Extension="png" ContentType="image/png"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' +
    '<Default Extension="bin" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>' +
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
    '<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
    '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>' +
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
  if (options.calcChain) {
    zip.set('xl/calcChain.xml', utf8(`${XML}<calcChain xmlns="${NS}"><c r="B1" i="2"/><c r="B1" i="3"/></calcChain>`));
    const name = 'xl/_rels/workbook.xml.rels';
    zip.set(name, utf8(decoder.decode(await zip.read(name)).replace('</Relationships>', `${rel('rIdChain', 'calcChain', 'calcChain.xml')}</Relationships>`)));
    extraTypes += '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>';
  }
  zip.set('[Content_Types].xml', utf8(decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', `${extraTypes}</Types>`)));
  return zip.write();
}

/**
 * What the rules of the first sheet show, as a view would work it out: which of A..D of the row
 * are filled, and the validations of the two cells below (their limits and the items of the list).
 */
function rulesShown(wb: Workbook, row = 0): unknown {
  const host = wb.sheets[0];
  const engine = new CalcEngine(wb);
  engine.invalidate();
  engine.recalculate(undefined, { overwrite: true });
  const formats = new ConditionalFormats(wb, host, { engine: () => engine, budget: Infinity });
  formats.invalidate();
  formats.begin();
  const rules = new DataValidations(wb, host, { engine: () => engine });
  const filled = ['A', 'B', 'C', 'D'].filter((letter) => {
    const found = formats.resolve(row, letter.charCodeAt(0) - 65);
    return !!found && !!found.style;
  });
  const whole = rules.ruleAt(row + 4, 0);
  return { filled: filled.join(''), whole: whole ? [whole.formula1, whole.formula2] : null, list: rules.ruleAt(row + 4, 1)?.formula1 ?? null, items: rules.list(row + 4, 1)?.items ?? null };
}

/** Of rulesShown(): what does not depend on what the sheets are called. */
const results = (shown: unknown) => [(shown as { filled: string }).filled, (shown as { items: unknown }).items];

// --- renaming --------------------------------------------------------------------

test('sheets: renaming a sheet nothing refers to writes the workbook part and nothing else', async () => {
  const original = await buildXlsx({
    workbookHead: '<bookViews><workbookView activeTab="1"/></bookViews>',
    sheets: [
      { name: 'Bir', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}${formula('B1', 'A1*2', 2)}</row></sheetData>`) },
      { name: 'İki', xml: worksheetXml(`<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1">${num('A1', 2)}</row></sheetData>`) },
    ],
  });
  for (const name of ['Özet & <Yeni> "1"', 'Otuzbir karakterlik bir sayfa a', 'Ağustos Çizelgesi ığüşöç', "It's"]) {
    const doc = await load(original);
    apply(doc, renameSheet(doc.wb, 0, name));
    const saved = await roundTrip(doc);
    assert.deepEqual(await changedParts(original, saved.bytes), ['xl/workbook.xml'], name);
    // Entries that were not written keep the compressed bytes they had.
    const before = ZipArchive.read(original);
    const after = ZipArchive.read(saved.bytes);
    for (const entry of before.names()) if (entry !== 'xl/workbook.xml') assert.ok(sameBytes(after.entry(entry)!.raw, before.entry(entry)!.raw), entry);
    assert.deepEqual(names(saved.wb), [name, 'İki']);
    assert.equal(saved.wb.activeSheet, 1);
    const xml = await text(saved.bytes, 'xl/workbook.xml');
    assert.ok(!xml.includes('fullCalcOnLoad'), 'no formula changed: nothing to recalculate');
    assert.ok(xml.includes('activeTab="1"'));
  }
});

test('sheets: a name of 32 characters and a name another sheet has are refused before anything is written', async () => {
  const doc = await load(await book());
  assert.throws(() => renameSheet(doc.wb, 0, 'x'.repeat(32)), RangeError);
  assert.throws(() => renameSheet(doc.wb, 0, 'sheet2'), RangeError);
  assert.ok(sameBytes(await save(doc), doc.bytes), 'nothing was modified');
});

test('sheets: after a rename, formulas, names, rules, validations, extension rules and shapes name the sheet as it is called now', async () => {
  const original = await book();
  const doc = await load(original);
  const name = 'Özet & <Yeni> "1"';
  apply(doc, renameSheet(doc.wb, 2, name));
  const saved = await roundTrip(doc);
  const quoted = `'${name}'`;

  assert.equal(formulaOf(saved.wb, 0, 'B1'), `Sheet2!A1+${quoted}!A2`);
  assert.equal(formulaOf(saved.wb, 1, 'B1'), `${quoted}!A1+1`);
  assert.deepEqual(
    saved.wb.definedNames.map((defined) => [defined.name, defined.value, defined.localSheet]),
    [
      ['Global', `${quoted}!$A$1:$A$2`, undefined],
      ['Local2', 'Sheet2!$A$1', 1],
      ['Local3', `${quoted}!$A$2`, 2],
      ['Two', 'Sheet2!$A$1', undefined],
    ],
  );
  // What the model does not hold was rewritten in the source text of the parts.
  const host = saved.wb.sheets[0];
  assert.deepEqual(host.conditionalFormats!.slice(0, 2).map((format) => format.rules[0].formulas), [[`${quoted}!$A$2>0`], ['Sheet2!$A$1>0']]);
  assert.deepEqual(
    host.dataValidations!.map((rule) => [rule.formula1, rule.formula2]),
    [
      [`${quoted}!$A$2`, 'Sheet2!$A$1'],
      ['Sheet2!$A$1:$A$2', undefined],
    ],
  );
  const sheet1 = await text(saved.bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet1.includes("<xm:f>'Özet &amp; &lt;Yeni&gt; \"1\"'!$A$2&gt;Sheet2!$A$1</xm:f>"), sheet1);
  assert.ok(sheet1.includes('<xm:sqref>D1</xm:sqref>'));
  const drawing = parseXml(await text(saved.bytes, 'xl/drawings/drawing1.xml')).root;
  assert.equal(all(drawing, 'sp')[0].attrs.get('textlink'), `${quoted}!$A$2`);
  const app = parseXml(await text(saved.bytes, 'docProps/app.xml')).root;
  assert.deepEqual(
    all(app, 'lpstr').map((el) => (el.children[0] as { value: string }).value),
    ['Worksheets', 'Named Ranges', 'Sheet1', 'Sheet2', name, 'Global'],
  );
  // The rules give what they gave in the workbook that was saved.
  assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb));
  assert.deepEqual(results(rulesShown(doc.wb)), results(rulesShown(await reopen(original))), 'a rename changes no result');

  // The renamed sheet itself has no formula that names it: its part is as it was, like everything not named below.
  assert.deepEqual(await changedParts(original, saved.bytes), ['docProps/app.xml', 'xl/drawings/drawing1.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml']);
  assert.ok((await text(saved.bytes, 'xl/workbook.xml')).includes('comment="kept"'), 'unknown attributes of a defined name are kept');

  // Undone: the file names the sheets as it did, and the parts only the rename touched are the file's again.
  doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(names(back.wb), ['Sheet1', 'Sheet2', 'Sheet3']);
  same(picture(back.wb), picture(await reopen(original)), 'the model after undo');
  // The sheets whose formulas were rewritten stay marked as modified and are written from the model; nothing else is touched.
  const touched = await changedParts(original, back.bytes);
  assert.ok(touched.indexOf('xl/workbook.xml') >= 0 && touched.every((part) => ['xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'].indexOf(part) >= 0), touched.join());
});

test('sheets: a sheet that is only named by raw content of an unmodified sheet makes that sheet be written', async () => {
  const original = await buildXlsx({
    styles: STYLES,
    sheets: [
      { name: 'Kurallar', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row></sheetData>${expression('A1', "'Veri Sayfası'!$A$1&gt;0", 1)}`) },
      { name: 'Veri Sayfası', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 3)}</row></sheetData>`) },
      { name: 'Başka', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 3)}</row></sheetData>${expression('A1', '$A$1&gt;0', 1)}`) },
    ],
  });
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 1, 'VERİ'));
  assert.ok(!doc.wb.sheets.some((sheet) => sheet.modified), 'no cell changed');
  const saved = await roundTrip(doc, false);
  assert.deepEqual(await changedParts(original, saved.bytes), ['xl/workbook.xml', 'xl/worksheets/sheet1.xml']);
  assert.deepEqual(saved.wb.sheets[0].conditionalFormats![0].rules[0].formulas, ["'VERİ'!$A$1>0"]);
  assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb));

  // Renamed again: every save starts from the file, not from the last save.
  apply(doc, renameSheet(doc.wb, 1, 'Veri 2'));
  const again = await roundTrip(doc, false);
  assert.deepEqual(again.wb.sheets[0].conditionalFormats![0].rules[0].formulas, ["'Veri 2'!$A$1>0"]);

  // Both undone: every part is what the file had.
  doc.history.undo();
  doc.history.undo();
  const back = await save(doc);
  assert.deepEqual(await changedParts(original, back), []);
});

test('sheets: renamed and undone before the first save, nothing is written', async () => {
  const original = await buildXlsx({ sheets: [{ name: 'Bir', xml: worksheetXml('<sheetData/>') }] });
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 0, 'İki'));
  doc.history.undo();
  assert.equal(doc.wb.sheetsModified, true);
  assert.ok(sameBytes(await save(doc), original));
});

// --- deleting --------------------------------------------------------------------

test('sheets: deleting a sheet takes its part and what only it leads to out of the package', async () => {
  const original = await book();
  const before = await checkPackage(original);
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  const saved = await roundTrip(doc);

  // Saved again without a change: the same package.
  assert.deepEqual(await changedParts(saved.bytes, await save(doc)), []);

  const goneParts = before.parts.filter((name) => saved.facts.parts.indexOf(name) < 0);
  assert.deepEqual(goneParts, [
    'xl/comments1.xml',
    'xl/drawings/_rels/drawing2.xml.rels',
    'xl/drawings/drawing2.xml',
    'xl/drawings/vmlDrawing1.vml',
    'xl/media/image2.png',
    'xl/printerSettings/printerSettings1.bin',
    'xl/worksheets/_rels/sheet2.xml.rels',
    'xl/worksheets/sheet2.xml',
  ]);
  assert.deepEqual(saved.facts.parts.filter((name) => before.parts.indexOf(name) < 0), [], 'no part was added');
  assert.ok(saved.facts.parts.indexOf('xl/media/image1.png') >= 0, 'the picture the first sheet shows as well stays');
  const types = await text(saved.bytes, '[Content_Types].xml');
  assert.ok(!types.includes('sheet2.xml') && !types.includes('drawing2.xml') && !types.includes('comments1.xml'), types);

  assert.deepEqual(
    saved.facts.sheets.map((sheet) => [sheet.name, sheet.sheetId, sheet.part]),
    [
      ['Sheet1', '1', 'xl/worksheets/sheet1.xml'],
      ['Sheet3', '3', 'xl/worksheets/sheet3.xml'],
    ],
  );
  // Formulas that named the sheet say so; the names of the sheet are gone, those after it follow their sheet.
  assert.equal(formulaOf(saved.wb, 0, 'B1'), '#REF!+Sheet3!A2');
  assert.equal(formulaOf(saved.wb, 1, 'C1'), 'SUM(#REF!)');
  assert.deepEqual(
    saved.wb.definedNames.map((defined) => [defined.name, defined.value, defined.localSheet]),
    [
      ['Global', 'Sheet3!$A$1:$A$2', undefined],
      ['Local3', 'Sheet3!$A$2', 1],
      ['Two', '#REF!', undefined],
    ],
  );
  const host = saved.wb.sheets[0];
  assert.deepEqual(host.conditionalFormats!.slice(0, 2).map((format) => format.rules[0].formulas), [['Sheet3!$A$2>0'], ['#REF!>0']]);
  assert.deepEqual(
    host.dataValidations!.map((rule) => [rule.formula1, rule.formula2]),
    [
      ['Sheet3!$A$2', '#REF!'],
      ['#REF!', undefined],
    ],
  );
  assert.ok((await text(saved.bytes, 'xl/worksheets/sheet1.xml')).includes('<xm:f>Sheet3!$A$2&gt;#REF!</xm:f>'));
  assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb));
  const app = parseXml(await text(saved.bytes, 'docProps/app.xml')).root;
  assert.deepEqual(
    all(app, 'lpstr').map((el) => (el.children[0] as { value: string }).value),
    ['Worksheets', 'Named Ranges', 'Sheet1', 'Sheet3', 'Global'],
  );
  assert.equal((all(app, 'i4')[0].children[0] as { value: string }).value, '2');
  assert.equal(all(app, 'vector')[1].attrs.get('size'), '3');
  const workbookXml = await text(saved.bytes, 'xl/workbook.xml');
  assert.ok(workbookXml.includes('xWindow="10"') && workbookXml.includes('comment="kept"') && workbookXml.includes('localSheetId="1"') && !workbookXml.includes('Local2'), workbookXml);

  // Undone: the sheet is back with every part it had, byte for byte.
  doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, before.parts);
  // Sheets whose formulas were rewritten and put back are written from the model, and the workbook part with them.
  const touched = await changedParts(original, back.bytes);
  assert.ok(touched.every((part) => ['[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml'].indexOf(part) >= 0), touched.join());
  same(picture(back.wb), picture(await reopen(original)), 'the model after undo');
  assert.equal(back.wb.sheets[1].comments!.length, 1);
  assert.equal(back.wb.sheets[1].drawings!.length, 2);
  assert.deepEqual(rulesShown(back.wb), rulesShown(await reopen(original)));

  // And deleted again by redo.
  doc.history.redo();
  const redone = await roundTrip(doc);
  assert.deepEqual(redone.facts.parts, saved.facts.parts);
  // (The results of formulas were calculated above, for the rules: the two sheets with formulas hold other values.)
  assert.deepEqual((await changedParts(saved.bytes, redone.bytes)).filter((name) => !/sheet[13]\.xml$/.test(name)), []);
});

test('sheets: a deleted sheet nothing refers to comes back by undo with the package as it was, calculation chain included', async () => {
  const original = await book({ loneSheet2: true, calcChain: true });
  const before = await checkPackage(original);
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  assert.ok(!doc.wb.sheets.some((sheet) => sheet.modified), 'no cell changed');
  const saved = await roundTrip(doc);
  assert.ok(saved.facts.parts.indexOf('xl/calcChain.xml') < 0, 'the chain counts sheets that are gone');
  assert.ok(saved.facts.parts.indexOf('xl/worksheets/sheet2.xml') < 0);
  assert.ok((await text(saved.bytes, 'xl/workbook.xml')).includes('fullCalcOnLoad="1"'));
  // The third sheet was not selected and is not now; only what lists sheets was written.
  assert.deepEqual(
    (await changedParts(original, saved.bytes)).filter((name) => saved.facts.parts.indexOf(name) >= 0),
    ['docProps/app.xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml'],
  );

  doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, before.parts);
  // Every part is the file's, the list of content types and of titles included, and so is the archive.
  assert.deepEqual(await changedParts(original, back.bytes), []);
  assert.ok(sameBytes(back.bytes, original), 'the file that was opened, byte for byte');
});

test('sheets: deleting the active sheet, the first sheet and a hidden sheet', async () => {
  const original = await buildXlsx({
    workbookHead: '<bookViews><workbookView activeTab="2" firstSheet="2"/></bookViews>',
    workbookTail: '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="3">D!$A$1:$B$2</definedName></definedNames>',
    sheets: [
      { name: 'A', xml: worksheetXml('<sheetData/>') },
      { name: 'B', xml: worksheetXml('<sheetData/>'), state: 'hidden' },
      { name: 'C', xml: worksheetXml('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/>') },
      { name: 'D', xml: worksheetXml('<sheetData/>') },
    ],
  });
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 2));
  let saved = await roundTrip(doc);
  assert.deepEqual(names(saved.wb), ['A', 'B', 'D']);
  assert.equal(saved.facts.activeTab, 2, 'the next visible sheet is shown');
  assert.ok((await text(saved.bytes, 'xl/workbook.xml')).includes('firstSheet="2"'));
  assert.deepEqual(saved.wb.definedNames.map((defined) => [defined.name, defined.localSheet]), [['_xlnm.Print_Area', 2]]);
  // The sheet that is shown now had no view at all: it got one that says it is selected.
  assert.ok((await text(saved.bytes, 'xl/worksheets/sheet4.xml')).includes('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/>'));

  apply(doc, deleteSheet(doc.wb, 0));
  saved = await roundTrip(doc);
  assert.deepEqual(names(saved.wb), ['B', 'D']);
  assert.equal(saved.facts.activeTab, 1);
  assert.deepEqual(saved.wb.sheets.map((sheet) => sheet.state), ['hidden', 'visible']);

  assert.throws(() => deleteSheet(doc.wb, 1), /last visible/);
  apply(doc, deleteSheet(doc.wb, 0));
  saved = await roundTrip(doc);
  assert.deepEqual(names(saved.wb), ['D']);
  assert.ok(!(await text(saved.bytes, 'xl/workbook.xml')).includes('activeTab'));
  assert.deepEqual(saved.facts.parts.filter((name) => name.startsWith('xl/worksheets/')), ['xl/worksheets/sheet4.xml']);

  for (let i = 0; i < 3; i++) doc.history.undo();
  const back = await save(doc);
  assert.ok(sameBytes(back, original), 'the file that was opened, byte for byte');
});

test('sheets: a save that fails leaves the package as it was, and the next one succeeds', async () => {
  const original = await book();
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  apply(doc, insertSheet(doc.wb, 0, { name: 'Yeni' }));
  // A name the file does not have is a mistake of the caller: the writer refuses.
  const names = doc.wb.definedNames;
  doc.wb.definedNames = names.concat([{ name: 'Yok', value: 'Sheet1!$A$1', hidden: false }]);
  await assert.rejects(writeXlsx(doc.file.package!, doc.wb), /defined name "Yok"/);
  assert.deepEqual(await changedParts(original, await doc.file.package!.save()), []);
  doc.wb.definedNames = names;
  const saved = await roundTrip(doc);
  assert.deepEqual(saved.facts.sheets.map((sheet) => sheet.name), ['Yeni', 'Sheet1', 'Sheet3']);
});

test('sheets: a custom view that showed a deleted sheet shows the active one; a second window opens on it', async () => {
  const original = await buildXlsx({
    workbookHead: '<bookViews><workbookView activeTab="0"/><workbookView activeTab="2" firstSheet="1"/></bookViews>',
    workbookTail:
      '<customWorkbookViews><customWorkbookView name="a" guid="{00000000-0000-0000-0000-000000000001}" windowWidth="1" windowHeight="1" activeSheetId="2"/>' +
      '<customWorkbookView name="b" guid="{00000000-0000-0000-0000-000000000002}" windowWidth="1" windowHeight="1" activeSheetId="3"/></customWorkbookViews>',
    sheets: [
      { name: 'A', xml: worksheetXml('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/><sheetView workbookViewId="1"/></sheetViews><sheetData/>') },
      { name: 'B', xml: worksheetXml('<sheetData/>') },
      { name: 'C', xml: worksheetXml('<sheetViews><sheetView workbookViewId="0"/><sheetView tabSelected="1" workbookViewId="1"/></sheetViews><sheetData/>') },
    ],
  });
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  // The model knows the first view of a sheet only: what the second window has selected in a sheet that is not rewritten stays.
  const saved = await roundTrip(doc, false);
  const xml = await text(saved.bytes, 'xl/workbook.xml');
  assert.ok(xml.includes('<bookViews><workbookView activeTab="0"/><workbookView firstSheet="1"/></bookViews>'), xml);
  assert.ok(/name="a"[^>]*activeSheetId="1"/.test(xml) && /name="b"[^>]*activeSheetId="3"/.test(xml), xml);
});

// --- inserting -------------------------------------------------------------------

test('sheets: a sheet inserted first, in the middle and last gets a part, a relationship and a content type', async () => {
  const original = await book();
  const before = await checkPackage(original);
  for (const at of [0, 1, 3]) {
    const doc = await load(original);
    apply(doc, insertSheet(doc.wb, at, { name: 'Yeni & "Sayfa"' }));
    type(doc, at, 'A1', '41');
    type(doc, at, 'B2', 'metin');
    type(doc, at, 'C3', '=A1+Sheet3!A2');
    type(doc, doc.wb.sheets.findIndex((sheet) => sheet.name === 'Sheet3'), 'D1', `='Yeni & "Sayfa"'!A1*2`);
    apply(doc, setColumnWidth(doc.wb, at, 1, 2, 20));
    apply(doc, mergeRange(doc.wb, at, { top: 4, left: 0, bottom: 5, right: 1 }));
    assert.equal(doc.wb.sheets[at].part, '', 'the model is not told about the part');

    const saved = await roundTrip(doc);
    assert.deepEqual(saved.facts.parts.filter((name) => before.parts.indexOf(name) < 0), ['xl/worksheets/sheet4.xml'], `inserted at ${at}`);
    assert.equal(saved.facts.sheets[at].part, 'xl/worksheets/sheet4.xml');
    assert.equal(saved.facts.sheets[at].sheetId, '4');
    assert.equal(saved.facts.activeTab, at === 0 ? 1 : 0, 'the sheet that was shown is the one shown');
    assert.ok((await text(saved.bytes, '[Content_Types].xml')).includes('<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'));
    assert.equal(formulaOf(saved.wb, at, 'C3'), 'A1+Sheet3!A2');
    assert.equal(valueOf(saved.wb, at, 'A1'), 41);
    assert.equal(saved.wb.sheets[at].columns[0].width, 20);
    assert.deepEqual(saved.wb.sheets[at].merges, [{ top: 4, left: 0, bottom: 5, right: 1 }]);
    // Names of sheets count the sheets as they are now.
    const shift = (index: number) => (index >= at ? index + 1 : index);
    assert.deepEqual(
      saved.wb.definedNames.map((defined) => [defined.name, defined.localSheet]),
      [
        ['Global', undefined],
        ['Local2', shift(1)],
        ['Local3', shift(2)],
        ['Two', undefined],
      ],
    );
    assert.deepEqual(rulesShown({ ...saved.wb, sheets: saved.wb.sheets.filter((sheet) => sheet.name !== 'Yeni & "Sayfa"') }), rulesShown(await reopen(original)));
    const app = parseXml(await text(saved.bytes, 'docProps/app.xml')).root;
    assert.deepEqual(all(app, 'lpstr').slice(2).map((el) => (el.children[0] as { value: string }).value), [...names(doc.wb), 'Global']);
    // Nothing of the sheets around it changed, but the cell typed into Sheet3.
    assert.deepEqual(
      (await changedParts(original, saved.bytes)).filter((name) => name !== 'xl/worksheets/sheet4.xml'),
      at === 0
        ? ['[Content_Types].xml', 'docProps/app.xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet3.xml']
        : ['[Content_Types].xml', 'docProps/app.xml', 'xl/_rels/workbook.xml.rels', 'xl/workbook.xml', 'xl/worksheets/sheet3.xml'],
    );

    // A second save writes the sheet to the same part: nothing multiplies.
    type(doc, at, 'A2', '7');
    const again = await roundTrip(doc);
    assert.deepEqual(again.facts.parts, saved.facts.parts);
    assert.equal(valueOf(again.wb, at, 'A2'), 7);
    assert.deepEqual(await changedParts(saved.bytes, again.bytes), ['xl/worksheets/sheet4.xml']);
  }
});

test('sheets: insert, save, undo, save leaves no trace of the sheet; redo brings it back', async () => {
  const original = await book({ loneSheet2: true });
  const before = await checkPackage(original);
  const doc = await load(original);
  apply(doc, insertSheet(doc.wb, 1, { locale: 'tr' }));
  assert.deepEqual(names(doc.wb), ['Sheet1', 'Sayfa1', 'Sheet2', 'Sheet3']);
  const saved = await roundTrip(doc);
  assert.equal(saved.facts.parts.length, before.parts.length + 1);

  doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, before.parts);
  // Nothing is left of it: the file that was opened, byte for byte.
  assert.deepEqual(await changedParts(original, back.bytes), []);
  assert.ok(sameBytes(back.bytes, original));

  doc.history.redo();
  const redone = await roundTrip(doc);
  assert.deepEqual(redone.facts.parts, saved.facts.parts);
  assert.deepEqual(await changedParts(saved.bytes, redone.bytes), []);

  // Deleted by an operation of its own instead of undo: the part goes as well.
  apply(doc, deleteSheet(doc.wb, 1));
  const deleted = await roundTrip(doc);
  assert.deepEqual(deleted.facts.parts, before.parts);
  assert.deepEqual(await changedParts(original, deleted.bytes), []);
});

test('sheets: two new sheets, one deleted again, and a name of the file taken by a new sheet', async () => {
  const original = await book();
  const doc = await load(original);
  apply(doc, insertSheet(doc.wb, 3, { name: 'X' }));
  apply(doc, insertSheet(doc.wb, 0, { name: 'Y' }));
  type(doc, 0, 'A1', 'y');
  type(doc, 4, 'A1', 'x');
  let saved = await roundTrip(doc);
  assert.deepEqual(saved.facts.sheets.map((sheet) => [sheet.name, sheet.part, sheet.sheetId]), [
    ['Y', 'xl/worksheets/sheet4.xml', '5'],
    ['Sheet1', 'xl/worksheets/sheet1.xml', '1'],
    ['Sheet2', 'xl/worksheets/sheet2.xml', '2'],
    ['Sheet3', 'xl/worksheets/sheet3.xml', '3'],
    ['X', 'xl/worksheets/sheet5.xml', '4'],
  ]);
  // The first new sheet is deleted; the other keeps the part it was given.
  apply(doc, deleteSheet(doc.wb, 0));
  saved = await roundTrip(doc);
  assert.deepEqual(saved.facts.sheets.map((sheet) => [sheet.name, sheet.part]).slice(3), [['X', 'xl/worksheets/sheet5.xml']]);
  assert.ok(saved.facts.parts.indexOf('xl/worksheets/sheet4.xml') < 0);
  assert.equal(valueOf(saved.wb, 3, 'A1'), 'x');

  // A sheet of the file is deleted, and a new sheet takes its name and the place of its part.
  apply(doc, deleteSheet(doc.wb, 1));
  apply(doc, insertSheet(doc.wb, 1, { name: 'Sheet2' }));
  type(doc, 1, 'A1', '100');
  saved = await roundTrip(doc);
  assert.deepEqual(saved.facts.sheets.map((sheet) => [sheet.name, sheet.part, sheet.sheetId]), [
    ['Sheet1', 'xl/worksheets/sheet1.xml', '1'],
    ['Sheet2', 'xl/worksheets/sheet4.xml', '5'],
    ['Sheet3', 'xl/worksheets/sheet3.xml', '3'],
    ['X', 'xl/worksheets/sheet5.xml', '4'],
  ]);
  // What the file said about its Sheet2 is not said about the new one.
  assert.equal(formulaOf(saved.wb, 0, 'B1'), '#REF!+Sheet3!A2');
  assert.deepEqual(saved.wb.sheets[0].dataValidations!.map((rule) => rule.formula1), ['Sheet3!$A$2', '#REF!']);
  assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb));
  assert.equal(saved.wb.sheets[1].comments, undefined);

  // Everything undone: the package of the file.
  while (doc.history.canUndo) doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, (await checkPackage(original)).parts);
  same(picture(back.wb), picture(await reopen(original)), 'the model after undoing everything');
});

// --- everything at once --------------------------------------------------------

test('sheets: delete Sheet2, call Sheet3 "Sheet2", add a sheet, move rows: raw content follows the sheets of the file', async () => {
  const original = await book();
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  apply(doc, renameSheet(doc.wb, 1, 'Sheet2'));
  apply(doc, insertSheet(doc.wb, 2, { name: 'Sheet3' }));
  type(doc, 2, 'A1', '1000');
  type(doc, 2, 'A2', '2000');
  // Rows move in the first sheet (its rules move), in the sheet the file calls Sheet3 (what points
  // at it follows) and in the new sheet called Sheet3 (which nothing of the file points at).
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }]);
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 2 }]);
  apply(doc, [{ type: 'insert', sheet: 2, axis: 'row', at: 0, count: 5 }]);
  const saved = await roundTrip(doc);

  assert.deepEqual(saved.facts.sheets.map((sheet) => [sheet.name, sheet.part]), [
    ['Sheet1', 'xl/worksheets/sheet1.xml'],
    ['Sheet2', 'xl/worksheets/sheet3.xml'],
    // Not the part of the deleted sheet, which may come back.
    ['Sheet3', 'xl/worksheets/sheet4.xml'],
  ]);
  assert.equal(formulaOf(saved.wb, 0, 'B2'), '#REF!+Sheet2!A4');
  const host = saved.wb.sheets[0];
  assert.deepEqual(
    host.conditionalFormats!.map((format) => [format.ranges, format.rules[0].formulas]).slice(0, 2),
    [
      [[{ top: 1, left: 0, bottom: 1, right: 0 }], ['Sheet2!$A$4>0']],
      [[{ top: 1, left: 1, bottom: 1, right: 1 }], ['#REF!>0']],
    ],
  );
  assert.deepEqual(
    host.dataValidations!.map((rule) => [rule.ranges, rule.formula1, rule.formula2]),
    [
      [[{ top: 5, left: 0, bottom: 5, right: 0 }], 'Sheet2!$A$4', '#REF!'],
      [[{ top: 5, left: 1, bottom: 5, right: 1 }], '#REF!', undefined],
    ],
  );
  const sheet1 = await text(saved.bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet1.includes('<xm:f>Sheet2!$A$4&gt;#REF!</xm:f>') && sheet1.includes('<xm:sqref>D2</xm:sqref>'), sheet1);
  const drawing = parseXml(await text(saved.bytes, 'xl/drawings/drawing1.xml')).root;
  assert.equal(all(drawing, 'sp')[0].attrs.get('textlink'), 'Sheet2!$A$4');
  assert.deepEqual(
    saved.wb.definedNames.map((defined) => [defined.name, defined.value, defined.localSheet]),
    [
      ['Global', 'Sheet2!$A$3:$A$4', undefined],
      ['Local3', 'Sheet2!$A$4', 1],
      ['Two', '#REF!', undefined],
    ],
  );
  // The rules are evaluated on the saved file as they were on the workbook in memory.
  assert.deepEqual(rulesShown(saved.wb, 1), rulesShown(doc.wb, 1));
  // The new sheet has its own part with its own cells, five rows down.
  assert.equal(valueOf(saved.wb, 2, 'A6'), 1000);
  assert.equal(saved.wb.sheets[2].comments, undefined);
  assert.ok(saved.facts.parts.indexOf('xl/comments1.xml') < 0);

  // Saved once more after another edit, then everything undone.
  type(doc, 0, 'F1', '5');
  await roundTrip(doc);
  while (doc.history.canUndo) doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, (await checkPackage(original)).parts);
  assert.deepEqual(back.facts.sheets.map((sheet) => [sheet.name, sheet.part]), [
    ['Sheet1', 'xl/worksheets/sheet1.xml'],
    ['Sheet2', 'xl/worksheets/sheet2.xml'],
    ['Sheet3', 'xl/worksheets/sheet3.xml'],
  ]);
  same(picture(back.wb), picture(await reopen(original)), 'the model after undoing everything');
  assert.deepEqual(rulesShown(back.wb), rulesShown(await reopen(original)));
  // Parts no cell edit touched are the file's, the drawing whose shape and anchors moved included.
  const changed = await changedParts(original, back.bytes);
  assert.ok(changed.every((part) => ['[Content_Types].xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml'].indexOf(part) >= 0), changed.join());
});

test('sheets: a sheet that an earlier save wrote for its raw content alone is written from the model once it is edited', async () => {
  // The validation of the second sheet names the first; rows inserted there make the second be rewritten.
  const original = await buildXlsx({
    sheets: [
      { name: 'Bir', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row></sheetData>`) },
      {
        name: 'İki',
        xml: worksheetXml(
          `<sheetData><row r="1">${num('A1', 2)}</row></sheetData><dataValidations count="1"><dataValidation type="list" sqref="B1"><formula1>Bir!$A$1:$A$3</formula1></dataValidation></dataValidations>`,
        ),
      },
    ],
  });
  const doc = await load(original);
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 2 }]);
  let saved = await roundTrip(doc, false);
  assert.equal(saved.wb.sheets[1].dataValidations![0].formula1, 'Bir!$A$3:$A$5');
  type(doc, 1, 'C1', '77');
  saved = await roundTrip(doc, false);
  assert.equal(valueOf(saved.wb, 1, 'C1'), 77, 'the edit is in the file');
  assert.equal(saved.wb.sheets[1].dataValidations![0].formula1, 'Bir!$A$3:$A$5');
});

// --- flavours of the format ------------------------------------------------------

/** The same package in the strict flavour: other namespaces for the main schema and for relationships. */
async function strict(bytes: Uint8Array): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  for (const name of zip.names()) {
    if (!/\.(xml|rels)$/.test(name)) continue;
    zip.set(name, utf8(decoder.decode(await zip.read(name)).split(NS_REL).join(NS_REL_STRICT).split(NS).join(NS_STRICT)));
  }
  return zip.write();
}

test('sheets: a sheet inserted into a strict package is written in the strict flavour', async () => {
  const original = await strict(
    await buildXlsx({
      sheets: [
        { name: 'Bir', xml: worksheetXml(`<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData><row r="1">${num('A1', 1)}</row></sheetData>`) },
        { name: 'İki', xml: worksheetXml('<sheetData/>') },
      ],
    }),
  );
  const doc = await load(original);
  assert.deepEqual(names(doc.wb), ['Bir', 'İki']);
  apply(doc, insertSheet(doc.wb, 1, { name: 'Yeni' }));
  type(doc, 1, 'A1', '=Bir!A1+1');
  const saved = await roundTrip(doc);
  const rels = await text(saved.bytes, 'xl/_rels/workbook.xml.rels');
  assert.ok(rels.includes(`Type="${NS_REL_STRICT}/worksheet" Target="worksheets/sheet3.xml"`), rels);
  const sheet = await text(saved.bytes, 'xl/worksheets/sheet3.xml');
  assert.ok(sheet.includes(`<worksheet xmlns="${NS_STRICT}" xmlns:r="${NS_REL_STRICT}">`), sheet);
  assert.ok(!sheet.includes(NS) && !(await text(saved.bytes, 'xl/workbook.xml')).includes(NS));
  assert.equal(formulaOf(saved.wb, 1, 'A1'), 'Bir!A1+1');
});

test('sheets: the relationship id of a new sheet uses the prefix of the workbook part, or declares its own', async () => {
  const plain = await buildXlsx({ sheets: [{ name: 'Bir', xml: worksheetXml('<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData/>') }] });
  const variants: [label: string, edit: (xml: string) => string, expected: RegExp][] = [
    ['another prefix', (xml) => xml.replace('xmlns:r=', 'xmlns:rel=').replace(/ r:id=/g, ' rel:id='), /<sheet name="Yeni" sheetId="2" rel:id="rId2"\/>/],
    [
      'declared on the element',
      (xml) => xml.replace(` xmlns:r="${NS_REL}"`, '').replace(/ r:id=/g, ` xmlns:q="${NS_REL}" q:id=`),
      new RegExp(`<sheet name="Yeni" sheetId="2" xmlns:r="${NS_REL.replace(/[.\\/]/g, '\\$&')}" r:id="rId2"/>`),
    ],
    [
      'a prefix on every element',
      (xml) => xml.replace('<workbook xmlns=', '<x:workbook xmlns:x=').replace('</workbook>', '</x:workbook>').replace(/<(\/?)(sheets|sheet)\b/g, '<$1x:$2'),
      /<x:sheet name="Yeni" sheetId="2" r:id="rId2"\/>/,
    ],
  ];
  for (const [label, edit, expected] of variants) {
    const zip = ZipArchive.read(plain);
    zip.set('xl/workbook.xml', utf8(edit(decoder.decode(await zip.read('xl/workbook.xml')))));
    const doc = await load(await zip.write());
    apply(doc, insertSheet(doc.wb, 1, { name: 'Yeni' }));
    type(doc, 1, 'B2', 'x');
    const saved = await roundTrip(doc);
    const xml = await text(saved.bytes, 'xl/workbook.xml');
    assert.ok(expected.test(xml), `${label}: ${xml}`);
    assert.equal(valueOf(saved.wb, 1, 'B2'), 'x', label);
  }
});

test('sheets: a new sheet does not take the name of a part that is there, whatever its case, nor a sheetId in use', async () => {
  const plain = await buildXlsx({
    sheets: [
      { name: 'Bir', xml: worksheetXml('<sheetData/>') },
      { name: 'İki', xml: worksheetXml('<sheetData/>') },
    ],
  });
  const zip = ZipArchive.read(plain);
  // The second sheet lives in "Sheet3.XML" and has the id 7; relationships without a part lie under the name sheet2.
  zip.set('xl/worksheets/Sheet3.XML', await zip.read('xl/worksheets/sheet2.xml'));
  zip.delete('xl/worksheets/sheet2.xml');
  zip.set('xl/worksheets/_rels/sheet2.xml.rels', utf8(relsXml('')));
  zip.set('xl/_rels/workbook.xml.rels', utf8(decoder.decode(await zip.read('xl/_rels/workbook.xml.rels')).replace('worksheets/sheet2.xml', 'worksheets/Sheet3.XML')));
  zip.set('xl/workbook.xml', utf8(decoder.decode(await zip.read('xl/workbook.xml')).replace('sheetId="2"', 'sheetId="7"')));
  zip.set('[Content_Types].xml', utf8(decoder.decode(await zip.read('[Content_Types].xml')).replace('/xl/worksheets/sheet2.xml', '/xl/worksheets/Sheet3.XML')));
  const doc = await load(await zip.write());
  apply(doc, insertSheet(doc.wb, 2));
  apply(doc, insertSheet(doc.wb, 3));
  await writeXlsx(doc.file.package!, doc.wb);
  const bytes = await doc.file.package!.save();
  const book = await text(bytes, 'xl/workbook.xml');
  assert.ok(/<sheet name="Sheet1" sheetId="8" r:id="rId3"\/><sheet name="Sheet2" sheetId="9" r:id="rId4"\/>/.test(book), book);
  const rels = await text(bytes, 'xl/_rels/workbook.xml.rels');
  assert.ok(rels.includes('Target="worksheets/sheet4.xml"') && rels.includes('Target="worksheets/sheet5.xml"'), rels);
  assert.deepEqual(names(await reopen(bytes)), ['Bir', 'İki', 'Sheet1', 'Sheet2']);
});

// --- the other writer ---------------------------------------------------------------

test('sheets: a workbook that is in no file (as read from xls) is built with the sheets as they are', async () => {
  // Every sheet has part '': buildXlsx() writes the model as it is.
  const wb: Workbook = { sheets: [], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 1 };
  for (const [i, name] of ['Sheet1', 'Sheet2', 'Sheet3'].entries()) wb.sheets.push(createSheet({ name, sheetId: i + 1, state: 'visible', kind: 'worksheet', part: '' }));
  wb.definedNames.push({ name: 'Global', value: 'Sheet3!$A$1', hidden: false }, { name: 'Local2', value: 'Sheet2!$A$1', localSheet: 1, hidden: false }, { name: 'Local3', value: 'Sheet3!$A$2', localSheet: 2, hidden: false });
  const history = new History(wb);
  const input = (sheet: number, ref: string, value: string) => {
    const { row, col } = parseCellRef(ref)!;
    history.apply(setCellInput(wb, sheet, row, col, value), undefined, undefined);
  };
  input(0, 'A1', '=Sheet2!A1+Sheet3!A1');
  input(1, 'A1', '7');
  input(2, 'A1', '5');
  history.apply(deleteSheet(wb, 1), undefined, undefined);
  history.apply(renameSheet(wb, 1, 'Özet 1'), undefined, undefined);
  history.apply(insertSheet(wb, 0, { name: 'Baş' }), undefined, undefined);
  input(0, 'A1', "='Özet 1'!A1");
  const bytes = await buildFromModel(wb);
  const facts = await checkPackage(bytes);
  assert.deepEqual(facts.sheets.map((sheet) => sheet.name), ['Baş', 'Sheet1', 'Özet 1']);
  const back = await reopen(bytes);
  assert.equal(formulaOf(back, 1, 'A1'), "#REF!+'Özet 1'!A1");
  assert.equal(formulaOf(back, 0, 'A1'), "'Özet 1'!A1");
  assert.deepEqual(
    back.definedNames.map((defined) => [defined.name, defined.value, defined.localSheet]),
    [
      ['Global', "'Özet 1'!$A$1", undefined],
      ['Local3', "'Özet 1'!$A$2", 2],
    ],
  );
  assert.equal(back.activeSheet, wb.activeSheet);
});

// --- findings of the review ------------------------------------------------------

const TABLE_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml';
const inline = (r: string, value: string) => `<c r="${r}" t="inlineStr"><is><t>${value}</t></is></c>`;

/** "Sheet1" holds a table whose second column is calculated from Sheet2 and totalled from Sheet3. */
async function tableBook(): Promise<Uint8Array> {
  const bytes = await buildXlsx({
    sheets: [
      {
        name: 'Sheet1',
        xml: worksheetXml(
          `<sheetData><row r="1">${inline('A1', 'A')}${inline('B1', 'B')}</row><row r="2">${num('A2', 1)}${num('B2', 2)}</row><row r="3">${num('A3', 1)}${num('B3', 2)}</row></sheetData>` +
            '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>',
        ),
      },
      { name: 'Sheet2', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 7)}</row></sheetData>`) },
      { name: 'Sheet3', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 0)}</row><row r="2">${num('A2', 5)}</row></sheetData>`) },
    ],
  });
  const zip = ZipArchive.read(bytes);
  zip.set('xl/worksheets/_rels/sheet1.xml.rels', utf8(relsXml(rel('rIdT', 'table', '../tables/table1.xml'))));
  zip.set(
    'xl/tables/table1.xml',
    utf8(
      `${XML}<table xmlns="${NS}" id="1" name="T1" displayName="T1" ref="A1:B3" totalsRowCount="1"><tableColumns count="2"><tableColumn id="1" name="A"/>` +
        '<tableColumn id="2" name="B" totalsRowFunction="custom"><calculatedColumnFormula>Sheet2!$A$1*2</calculatedColumnFormula><totalsRowFormula>SUM(Sheet3!A1:A2)+Sheet2!A1</totalsRowFormula></tableColumn>' +
        '</tableColumns></table>',
    ),
  );
  zip.set('[Content_Types].xml', utf8(decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', `<Override PartName="/xl/tables/table1.xml" ContentType="${TABLE_TYPE}"/></Types>`)));
  return zip.write();
}

const tableFormulas = async (bytes: Uint8Array) => {
  const root = parseXml(await text(bytes, 'xl/tables/table1.xml')).root;
  return [...all(root, 'calculatedColumnFormula'), ...all(root, 'totalsRowFormula')].map((el) => (el.children[0] as { value: string }).value);
};

test('review: the formulas of tables name sheets as they are called now', async () => {
  const original = await tableBook();
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 1, 'Yeni Ad'));
  let saved = await roundTrip(doc, false);
  assert.deepEqual(await tableFormulas(saved.bytes), ["'Yeni Ad'!$A$1*2", "SUM(Sheet3!A1:A2)+'Yeni Ad'!A1"]);
  assert.deepEqual(await changedParts(original, saved.bytes), ['xl/tables/table1.xml', 'xl/workbook.xml']);

  // A header that is edited renames its column in the same part.
  type(doc, 0, 'B1', 'Tutar');
  saved = await roundTrip(doc, false);
  assert.deepEqual(await tableFormulas(saved.bytes), ["'Yeni Ad'!$A$1*2", "SUM(Sheet3!A1:A2)+'Yeni Ad'!A1"]);
  assert.ok((await text(saved.bytes, 'xl/tables/table1.xml')).includes('name="Tutar"'));
  doc.history.undo();

  // A sheet a table reads is deleted (the sheet of the table itself cannot be).
  apply(doc, deleteSheet(doc.wb, 2));
  saved = await roundTrip(doc, false);
  assert.deepEqual(await tableFormulas(saved.bytes), ["'Yeni Ad'!$A$1*2", "SUM(#REF!)+'Yeni Ad'!A1"]);
  assert.throws(() => deleteSheet(doc.wb, 0), /table/);

  // Rows that move in a sheet the table reads.
  doc.history.undo();
  apply(doc, [{ type: 'insert', sheet: 2, axis: 'row', at: 0, count: 3 }]);
  saved = await roundTrip(doc, false);
  assert.deepEqual(await tableFormulas(saved.bytes), ["'Yeni Ad'!$A$1*2", "SUM(Sheet3!A4:A5)+'Yeni Ad'!A1"]);

  while (doc.history.canUndo) doc.history.undo();
  const back = await roundTrip(doc, false);
  assert.ok(sameBytes((await content(back.bytes)).get('xl/tables/table1.xml'), (await content(original)).get('xl/tables/table1.xml')), 'the table part is the file\'s again');
});

const SPARK = (f: string) =>
  '<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
  `<x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>${f}</xm:f><xm:sqref>C1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext>`;
const X14_RULE = (f: string) =>
  '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>' +
  '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="expression" priority="9" id="{00000000-0000-0000-0000-000000000001}">' +
  `<xm:f>${f}</xm:f></x14:cfRule><xm:sqref>D1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>`;

/** Ways a file may spell the name of a sheet inside the text of an element. */
const SPELLINGS: [label: string, name: string, written: string][] = [
  ['plain', 'Veri', 'Veri'],
  ['an ampersand as an entity', 'R&D', "'R&amp;D'"],
  ['an ampersand as a character reference', 'R&D', "'R&#38;D'"],
  ['letters as decimal and hexadecimal character references', 'Bölüm', "'B&#246;l&#xFC;m'"],
  ['apostrophes as entities', "It's", '&apos;It&apos;&apos;s&apos;'],
  ['apostrophes as character references', "It's", "'It&#39;&#39;s'"],
  ['a space as a character reference', 'My Sheet', "'My&#32;Sheet'"],
  ['a less-than sign half escaped', 'a<b>c', "'a&lt;b>c'"],
];

async function spelledBook(name: string, written: string): Promise<Uint8Array> {
  return buildXlsx({
    styles: STYLES,
    sheets: [
      {
        name: 'Kurallar',
        xml: worksheetXml(
          `<sheetData><row r="1">${num('A1', 1)}</row></sheetData>${expression('A1', `${written}!$A$1&gt;0`, 1)}` +
            `<conditionalFormatting sqref="B1"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="num" val="0"/><cfvo type="formula" val="${written.replace(/'/g, '&#39;').replace(/&#39;&#39;/g, "''").replace(/^&#39;|&#39;$/g, "'")}!$A$2"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting>` +
            `<dataValidations count="1"><dataValidation type="whole" sqref="A5"><formula1>${written}!$A$1</formula1><formula2>${written}!$A$2+1</formula2></dataValidation></dataValidations>` +
            `<extLst>${X14_RULE(`${written}!$A$2&gt;1`)}${SPARK(`${written}!A1:A2`)}</extLst>`,
        ),
      },
      { name: name.replace(/&/g, '&amp;').replace(/</g, '&lt;'), xml: worksheetXml(`<sheetData><row r="1">${num('A1', 3)}</row><row r="2">${num('A2', 4)}</row></sheetData>`) },
      { name: 'Başka', xml: worksheetXml('<sheetData/>') },
    ],
  });
}

/** The formulas of the first sheet that the file keeps as text: rules, bars, validations, extension rules, sparklines. */
async function writtenFormulas(bytes: Uint8Array): Promise<string[]> {
  const wb = await reopen(bytes);
  const host = wb.sheets[0];
  const root = parseXml(await text(bytes, 'xl/worksheets/sheet1.xml')).root;
  return [
    ...host.conditionalFormats!.slice(0, 1).map((format) => format.rules[0].formulas[0]),
    ...all(root, 'cfvo').filter((el) => el.attrs.get('type') === 'formula').map((el) => el.attrs.get('val')!),
    host.dataValidations![0].formula1!,
    host.dataValidations![0].formula2!,
    ...all(root, 'f').map((el) => (el.children[0] as { value: string }).value),
  ];
}

test('review: however the file spells the name of a sheet, raw content follows a rename and a deletion', async () => {
  for (const [label, name, written] of SPELLINGS) {
    const original = await spelledBook(name, written);
    const quoted = /^[A-Za-z_]+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
    assert.deepEqual(await writtenFormulas(original), [`${quoted}!$A$1>0`, `${quoted}!$A$2`, `${quoted}!$A$1`, `${quoted}!$A$2+1`, `${quoted}!$A$2>1`, `${quoted}!A1:A2`], `${label}: the fixture`);

    const doc = await load(original);
    assert.deepEqual(names(doc.wb), ['Kurallar', name, 'Başka'], label);
    apply(doc, renameSheet(doc.wb, 1, 'Yeni'));
    assert.ok(!doc.wb.sheets[0].modified, 'no cell names the sheet: the rules are raw content of an unmodified sheet');
    let saved = await roundTrip(doc, false);
    assert.deepEqual(await writtenFormulas(saved.bytes), ['Yeni!$A$1>0', 'Yeni!$A$2', 'Yeni!$A$1', 'Yeni!$A$2+1', 'Yeni!$A$2>1', 'Yeni!A1:A2'], `${label}: renamed`);
    // What the view works out in memory is what the file says.
    const rules = new DataValidations(doc.wb, doc.wb.sheets[0], {});
    assert.deepEqual([rules.ruleAt(4, 0)?.formula1, rules.ruleAt(4, 0)?.formula2], ['Yeni!$A$1', 'Yeni!$A$2+1'], label);
    assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb), label);

    doc.history.undo();
    apply(doc, deleteSheet(doc.wb, 1));
    saved = await roundTrip(doc, false);
    assert.deepEqual(await writtenFormulas(saved.bytes), ['#REF!>0', '#REF!', '#REF!', '#REF!+1', '#REF!>1', '#REF!'], `${label}: deleted`);
    assert.deepEqual([rules.ruleAt(4, 0)?.formula1, rules.ruleAt(4, 0)?.formula2], ['#REF!', '#REF!+1'], label);

    doc.history.undo();
    assert.deepEqual((await changedParts(original, await save(doc))).filter((part) => part !== '[Content_Types].xml'), [], `${label}: undone`);
  }
});

test('review: relationships of a deleted sheet that point at the parts the package stands on delete none of them', async () => {
  const plain = await book({ loneSheet2: true });
  const zip = ZipArchive.read(plain);
  const name = 'xl/worksheets/_rels/sheet2.xml.rels';
  const hostile =
    rel('rIdX1', 'image', '/xl/_rels/workbook.xml.rels') +
    rel('rIdX2', 'image', '/_rels/.rels') +
    rel('rIdX3', 'image', '/xl/workbook.xml') +
    rel('rIdX4', 'image', '/xl/worksheets/sheet1.xml') +
    rel('rIdX5', 'image', '/xl/worksheets/_rels/sheet1.xml.rels') +
    rel('rIdX6', 'image', '/[Content_Types].xml') +
    rel('rIdX7', 'image', '../drawings/_rels/drawing1.xml.rels') +
    rel('rIdX8', 'image', '../worksheets/sheet3.xml') +
    rel('rIdX9', 'image', '../styles.xml') +
    rel('rIdX10', 'image', '_rels/sheet2.xml.rels');
  zip.set(name, utf8(decoder.decode(await zip.read(name)).replace('</Relationships>', `${hostile}</Relationships>`)));
  const original = await zip.write();
  const doc = await load(original);
  apply(doc, deleteSheet(doc.wb, 1));
  const saved = await roundTrip(doc);
  const before = [...(await content(original)).keys()].filter((part) => part !== '[Content_Types].xml').sort();
  assert.deepEqual(
    before.filter((part) => saved.facts.parts.indexOf(part) < 0),
    [
      'xl/comments1.xml',
      'xl/drawings/_rels/drawing2.xml.rels',
      'xl/drawings/drawing2.xml',
      'xl/drawings/vmlDrawing1.vml',
      'xl/media/image2.png',
      'xl/printerSettings/printerSettings1.bin',
      'xl/worksheets/_rels/sheet2.xml.rels',
      'xl/worksheets/sheet2.xml',
    ],
  );
  assert.deepEqual(names(saved.wb), ['Sheet1', 'Sheet3']);
  assert.equal(saved.wb.sheets[0].drawings!.length, 2);
  assert.equal(valueOf(saved.wb, 1, 'A2'), 5);
  doc.history.undo();
  const back = await save(doc);
  assert.ok(sameBytes(back, original), 'the file that was opened, byte for byte');
  assert.deepEqual(names(await reopen(back)), ['Sheet1', 'Sheet2', 'Sheet3']);
});

test('review: a sheet renamed and then deleted leaves errors in the raw rules of the file', async () => {
  const original = await book();
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 1, 'Geçici'));
  apply(doc, deleteSheet(doc.wb, 1));
  const saved = await roundTrip(doc);
  const host = saved.wb.sheets[0];
  assert.deepEqual(host.conditionalFormats!.slice(0, 2).map((format) => format.rules[0].formulas), [['Sheet3!$A$2>0'], ['#REF!>0']]);
  assert.deepEqual(host.dataValidations!.map((rule) => [rule.formula1, rule.formula2]), [
    ['Sheet3!$A$2', '#REF!'],
    ['#REF!', undefined],
  ]);
  assert.ok((await text(saved.bytes, 'xl/worksheets/sheet1.xml')).includes('<xm:f>Sheet3!$A$2&gt;#REF!</xm:f>'));
  assert.deepEqual(rulesShown(saved.wb), rulesShown(doc.wb));
});

test('review: consolidation sources and 3D references in raw content follow renamed and deleted sheets', async () => {
  const original = await buildXlsx({
    sheets: [
      {
        name: 'Sheet1',
        xml: worksheetXml(
          '<sheetData/><dataConsolidate><dataRefs count="3"><dataRef ref="A1:B2" sheet="Sheet2"/><dataRef ref="A1:B2" sheet="sheet3"/><dataRef ref="A1:B2" sheet="Sheet4"/></dataRefs></dataConsolidate>' +
            '<dataValidations count="1"><dataValidation type="whole" sqref="A5"><formula1>SUM(Sheet2:Sheet4!$A$1)</formula1><formula2>SUM(Sheet3:Sheet3!$A$1)</formula2></dataValidation></dataValidations>',
        ),
      },
      ...['Sheet2', 'Sheet3', 'Sheet4'].map((name) => ({ name, xml: worksheetXml(`<sheetData><row r="1">${num('A1', 3)}</row></sheetData>`) })),
    ],
  });
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 1, 'İlk & Son'));
  let saved = await roundTrip(doc, false);
  let sheet1 = await text(saved.bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet1.includes('<dataRefs count="3"><dataRef ref="A1:B2" sheet="İlk &amp; Son"/><dataRef ref="A1:B2" sheet="sheet3"/><dataRef ref="A1:B2" sheet="Sheet4"/></dataRefs>'), sheet1);
  assert.equal(saved.wb.sheets[0].dataValidations![0].formula1, "SUM('İlk & Son:Sheet4'!$A$1)");

  apply(doc, deleteSheet(doc.wb, 2));
  saved = await roundTrip(doc, false);
  sheet1 = await text(saved.bytes, 'xl/worksheets/sheet1.xml');
  assert.ok(sheet1.includes('<dataRefs count="2"><dataRef ref="A1:B2" sheet="İlk &amp; Son"/><dataRef ref="A1:B2" sheet="Sheet4"/></dataRefs>'), sheet1);
  assert.deepEqual([saved.wb.sheets[0].dataValidations![0].formula1, saved.wb.sheets[0].dataValidations![0].formula2], ["SUM('İlk & Son:Sheet4'!$A$1)", 'SUM(#REF!)']);

  // The first end of the run goes: the file says what the view shows.
  apply(doc, deleteSheet(doc.wb, 1));
  saved = await roundTrip(doc, false);
  const rules = new DataValidations(doc.wb, doc.wb.sheets[0], {});
  assert.equal(rules.ruleAt(4, 0)?.formula1, 'SUM(Sheet4!$A$1)');
  assert.equal(saved.wb.sheets[0].dataValidations![0].formula1, 'SUM(Sheet4!$A$1)');
  assert.ok((await text(saved.bytes, 'xl/worksheets/sheet1.xml')).includes('<dataRefs count="1"><dataRef ref="A1:B2" sheet="Sheet4"/></dataRefs>'));
});

test('review: parts around a sheet that an earlier save changed are the file\'s again when the sheet comes back', async () => {
  const original = await book();
  const zip = ZipArchive.read(original);
  // The drawing of the second sheet shows a cell of the third.
  const shape = sp({}).replace('<xdr:sp>', '<xdr:sp macro="" textlink="Sheet3!$A$1">');
  zip.set('xl/drawings/drawing2.xml', utf8(drawingXml(twoCell([1, 1], [3, 4], pic('rId1')) + twoCell([4, 1], [6, 4], pic('rId2')) + twoCell([7, 1], [8, 4], shape))));
  const start = await zip.write();
  const doc = await load(start);
  apply(doc, renameSheet(doc.wb, 2, 'Son'));
  let saved = await roundTrip(doc);
  assert.equal(all(parseXml(await text(saved.bytes, 'xl/drawings/drawing2.xml')).root, 'sp')[0].attrs.get('textlink'), 'Son!$A$1');
  // The sheet is deleted with the drawing the last save changed, saved, and brought back with the rename undone.
  apply(doc, deleteSheet(doc.wb, 1));
  saved = await roundTrip(doc);
  assert.ok(saved.facts.parts.indexOf('xl/drawings/drawing2.xml') < 0);
  doc.history.undo();
  saved = await roundTrip(doc);
  assert.equal(all(parseXml(await text(saved.bytes, 'xl/drawings/drawing2.xml')).root, 'sp')[0].attrs.get('textlink'), 'Son!$A$1');
  doc.history.undo();
  const back = await roundTrip(doc);
  const a = await content(start);
  const b = await content(back.bytes);
  for (const part of ['xl/drawings/drawing2.xml', 'xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing2.xml.rels', 'xl/comments1.xml', 'xl/drawings/vmlDrawing1.vml', 'docProps/app.xml']) {
    assert.ok(sameBytes(b.get(part), a.get(part)), part);
  }
});

test('review: a list of titles that does not hold the names of the sheets is left alone', async () => {
  const lpstrs = async (bytes: Uint8Array) => all(parseXml(await text(bytes, 'docProps/app.xml')).root, 'lpstr').map((el) => (el.children[0] as { value: string }).value);
  const variants: [label: string, edit: (xml: string) => string][] = [
    ['another name', (xml) => xml.replace('<vt:lpstr>Sheet2</vt:lpstr>', '<vt:lpstr>Other</vt:lpstr>')],
    ['another count', (xml) => xml.replace('<vt:i4>3</vt:i4>', '<vt:i4>2</vt:i4>')],
    ['counts that do not add up', (xml) => xml.replace('<vt:i4>1</vt:i4>', '<vt:i4>4</vt:i4>')],
    ['a count that is no number', (xml) => xml.replace('<vt:i4>1</vt:i4>', '<vt:i4>x</vt:i4>')],
    ['a title that is no text', (xml) => xml.replace('<vt:lpstr>Global</vt:lpstr>', '<vt:i4>1</vt:i4>')],
    ['not XML', () => '<Properties'],
  ];
  for (const [label, edit] of variants) {
    const zip = ZipArchive.read(await book());
    zip.set('docProps/app.xml', utf8(edit(APP)));
    const original = await zip.write();
    const doc = await load(original);
    apply(doc, renameSheet(doc.wb, 2, 'Yeni'));
    apply(doc, insertSheet(doc.wb, 0));
    const saved = await save(doc);
    assert.ok(sameBytes((await content(saved)).get('docProps/app.xml'), (await content(original)).get('docProps/app.xml')), label);
    assert.deepEqual(names(await reopen(saved)), ['Sheet3', 'Sheet1', 'Sheet2', 'Yeni'], label);
  }
  // The same operations on the list as the file of the other tests has it.
  const doc = await load(await book());
  apply(doc, renameSheet(doc.wb, 2, 'Yeni'));
  apply(doc, insertSheet(doc.wb, 0));
  assert.deepEqual(await lpstrs(await save(doc)), ['Worksheets', 'Named Ranges', 'Sheet3', 'Sheet1', 'Sheet2', 'Yeni', 'Global']);
});

test('review: a part of a deleted sheet that stays in the package is not left as an earlier save wrote it', async () => {
  const zip = ZipArchive.read(await book());
  // The drawing of the second sheet shows a cell of the third, and the workbook part leads to it as well.
  const shape = sp({}).replace('<xdr:sp>', '<xdr:sp macro="" textlink="Sheet3!$A$1">');
  zip.set('xl/drawings/drawing2.xml', utf8(drawingXml(twoCell([1, 1], [3, 4], pic('rId1')) + twoCell([4, 1], [6, 4], pic('rId2')) + twoCell([7, 1], [8, 4], shape))));
  const relsName = 'xl/_rels/workbook.xml.rels';
  zip.set(relsName, utf8(decoder.decode(await zip.read(relsName)).replace('</Relationships>', `${rel('rIdKeep', 'drawing', 'drawings/drawing2.xml')}</Relationships>`)));
  const start = await zip.write();
  const doc = await load(start);
  apply(doc, renameSheet(doc.wb, 2, 'Son'));
  let saved = await save(doc);
  assert.equal(all(parseXml(await text(saved, 'xl/drawings/drawing2.xml')).root, 'sp')[0].attrs.get('textlink'), 'Son!$A$1');
  doc.history.undo();
  apply(doc, deleteSheet(doc.wb, 1));
  saved = (await roundTrip(doc)).bytes;
  const after = await content(saved);
  assert.ok(!after.has('xl/worksheets/sheet2.xml') && !after.has('xl/comments1.xml'));
  assert.ok(sameBytes(after.get('xl/drawings/drawing2.xml'), (await content(start)).get('xl/drawings/drawing2.xml')), 'the drawing is the file\'s again');
  assert.ok(after.has('xl/media/image2.png'), 'what the drawing shows stays with it');
});

// --- second review ---------------------------------------------------------------

test('review 2: a sheet that is not in the file does not cost the file its calculation chain', async () => {
  const original = await book({ loneSheet2: true, calcChain: true });
  const before = await checkPackage(original);

  // (a) insert, save, undo, save: the archive that was opened.
  const doc = await load(original);
  apply(doc, insertSheet(doc.wb, 1, { locale: 'en' }));
  const saved = await roundTrip(doc);
  assert.ok(saved.facts.parts.indexOf('xl/calcChain.xml') < 0, 'the chain counts sheets by their place');
  doc.history.undo();
  const back = await roundTrip(doc);
  assert.deepEqual(back.facts.parts, before.parts);
  assert.deepEqual(await changedParts(original, back.bytes), []);
  assert.ok(sameBytes(back.bytes, original), 'the file that was opened, byte for byte');

  // (b) insert, save, delete that sheet, save: a package that holds together, with the chain of the file.
  const other = await load(original);
  apply(other, insertSheet(other.wb, 3, { locale: 'en' }));
  type(other, 3, 'A1', '=1+1');
  await roundTrip(other);
  apply(other, deleteSheet(other.wb, 3));
  const deleted = await roundTrip(other);
  assert.deepEqual(deleted.facts.parts, before.parts);
  assert.deepEqual(await changedParts(original, deleted.bytes), []);

  // (c) a cell of the file changes: the chain is gone for good, with its relationship and its content type.
  const third = await load(original);
  apply(third, insertSheet(third.wb, 1, { locale: 'en' }));
  await roundTrip(third);
  type(third, 0, 'A1', '5');
  const edited = await roundTrip(third);
  assert.ok(edited.facts.parts.indexOf('xl/calcChain.xml') < 0);
  assert.ok(!(await text(edited.bytes, 'xl/_rels/workbook.xml.rels')).includes('calcChain'));
  assert.ok(!(await text(edited.bytes, '[Content_Types].xml')).includes('calcChain'));
  third.history.undo();
  third.history.undo();
  const after = await roundTrip(third);
  assert.ok(after.facts.parts.indexOf('xl/calcChain.xml') < 0, 'a modified sheet stays modified');
  assert.ok(!(await text(after.bytes, 'xl/_rels/workbook.xml.rels')).includes('calcChain'));
});

/** The book of most tests with links to places in the workbook written as relationship targets: of a shape of Sheet1, and of a cell of Sheet3. */
async function placesBook(): Promise<Uint8Array> {
  const zip = ZipArchive.read(await book({ loneSheet2: true }));
  const drawingRels = 'xl/drawings/_rels/drawing1.xml.rels';
  zip.set(
    drawingRels,
    utf8(
      decoder
        .decode(await zip.read(drawingRels))
        .replace('</Relationships>', `${rel('rIdL', 'hyperlink', '#Sheet3!A2')}${rel('rIdM', 'hyperlink', "#'Sheet2'!$B$3:$C$4")}${rel('rIdN', 'hyperlink', '#Global')}${rel('rIdX', 'hyperlink', '#Sheet3!A2', true)}</Relationships>`),
    ),
  );
  const sheet3 = 'xl/worksheets/sheet3.xml';
  zip.set(sheet3, utf8(decoder.decode(await zip.read(sheet3)).replace('</sheetData>', '</sheetData><hyperlinks><hyperlink ref="D1" r:id="rIdIn"/></hyperlinks>')));
  zip.set('xl/worksheets/_rels/sheet3.xml.rels', utf8(relsXml(rel('rIdIn', 'hyperlink', '#Sheet2!A1'))));
  return zip.write();
}

const targetsIn = async (bytes: Uint8Array, name: string) => all(parseXml(await ZipArchive.read(bytes).read(name)).root, 'Relationship').map((el) => `${el.attrs.get('Id')}=${el.attrs.get('Target')}`);

test('review 2: relationship targets that are places in the workbook follow renamed and deleted sheets and moved rows', async () => {
  const original = await placesBook();
  const DRAWING = 'xl/drawings/_rels/drawing1.xml.rels';
  const SHEET = 'xl/worksheets/_rels/sheet3.xml.rels';
  const doc = await load(original);
  apply(doc, renameSheet(doc.wb, 2, 'Third one'));
  apply(doc, renameSheet(doc.wb, 1, 'Two'));
  let bytes = await save(doc);
  await checkPackage(bytes);
  // A name and a link to another file are not places of a sheet.
  assert.deepEqual(await targetsIn(bytes, DRAWING), ['rId1=../media/image1.png', "rIdL=#'Third one'!A2", 'rIdM=#Two!$B$3:$C$4', 'rIdN=#Global', 'rIdX=#Sheet3!A2']);
  assert.deepEqual(await targetsIn(bytes, SHEET), ['rIdIn=#Two!A1']);

  // Rows move in the sheet a target names.
  apply(doc, [{ type: 'insert', sheet: 1, axis: 'row', at: 0, count: 2 }]);
  bytes = await save(doc);
  await checkPackage(bytes);
  assert.deepEqual((await targetsIn(bytes, DRAWING)).slice(1, 3), ["rIdL=#'Third one'!A2", 'rIdM=#Two!$B$5:$C$6']);
  assert.deepEqual(await targetsIn(bytes, SHEET), ['rIdIn=#Two!A3']);
  doc.history.undo();

  // The sheet is deleted: the relationship stays (the drawing names it by its id) and leads nowhere.
  apply(doc, deleteSheet(doc.wb, 1));
  bytes = await save(doc);
  await checkPackage(bytes);
  assert.deepEqual((await targetsIn(bytes, DRAWING)).slice(1, 3), ["rIdL=#'Third one'!A2", 'rIdM=#REF!']);
  assert.deepEqual(await targetsIn(bytes, SHEET), ['rIdIn=#REF!']);

  // Undone to the start: the parts the file had.
  doc.history.undo();
  doc.history.undo();
  doc.history.undo();
  bytes = await save(doc);
  assert.deepEqual(names(doc.wb), ['Sheet1', 'Sheet2', 'Sheet3']);
  const changed = await changedParts(original, bytes);
  assert.ok(changed.indexOf(DRAWING) < 0 && changed.indexOf(SHEET) < 0, changed.join(' '));
});
