// Parts that declare a single-byte encoding (ISO-8859-1, windows-1252, ISO-8859-9,
// windows-1254, US-ASCII) are read as what they declare; a part that is not touched keeps its
// bytes, and one that is written anew is UTF-8 and says so.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, parseXml, ZipArchive } from '../src/index';
import { XmlScanner } from '../src/xml-scan';
import { decodeXmlBytes, xmlEncodingRead } from '../src/xml';
import { copyChart, duplicateChart, sheetCharts, updateChart } from '../src/ops/charts';
import { RICH, richBook } from './chart-copy-fixtures';
import { chartsOf, load, packageProblems, part, reopen, same, save, text } from './chart-edit-fixtures';
import { buildDocx } from './docx-fixtures';

const TURKISH = 'ğüşiöçİĞÜŞÖÇ ı';
const TR: Record<string, number> = { ğ: 0xf0, Ğ: 0xd0, ş: 0xfe, Ş: 0xde, ı: 0xfd, İ: 0xdd };

/** A text as bytes of windows-1254 (which ISO-8859-9 is, but for 0x80 to 0x9F) or of Latin-1. */
function single(textOf: string, turkish: boolean): Uint8Array {
  return Uint8Array.from(textOf, (ch) => {
    if (turkish && TR[ch] !== undefined) return TR[ch];
    const code = ch.charCodeAt(0);
    assert.ok(code < 256, `"${ch}" is a letter of the encoding`);
    return code;
  });
}

const declare = (xml: string, encoding: string) => (/^<\?xml/.test(xml) ? xml.replace(/^<\?xml[^?]*\?>/, `<?xml version="1.0" encoding="${encoding}" standalone="yes"?>`) : `<?xml version="1.0" encoding="${encoding}"?>${xml}`);

test('xml encoding: the single-byte encodings are decoded as the platform decodes them', () => {
  const all = new Uint8Array(256).map((_v, i) => i);
  // As browsers read them: ISO-8859-1 is windows-1252 and ISO-8859-9 is windows-1254 (0x80 to 0x9F hold letters and marks).
  for (const [name, turkish] of [
    ['windows-1252', false],
    ['ISO-8859-1', false],
    ['latin1', false],
    ['windows-1254', true],
    ['ISO-8859-9', true],
    ['cp1254', true],
  ] as const) {
    const head = new TextEncoder().encode(`<?xml version='1.0' encoding='${name}'?><a>`);
    const bytes = new Uint8Array(head.length + 256);
    bytes.set(head);
    bytes.set(all, head.length);
    const decoded = decodeXmlBytes(bytes).slice(head.length);
    assert.equal(decoded.length, 256, name);
    for (let i = 0; i < 0x80; i++) assert.equal(decoded.charCodeAt(i), i, `${name}: ASCII`);
    assert.equal(decoded.slice(0x80, 0xa0), `\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D${turkish ? '\u008E' : '\u017D'}\u008F\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D${turkish ? '\u009E' : '\u017E'}\u0178`, name);
    for (let i = 0xa0; i < 0x100; i++) {
      const special = turkish ? ({ 0xd0: 'Ğ', 0xdd: 'İ', 0xde: 'Ş', 0xf0: 'ğ', 0xfd: 'ı', 0xfe: 'ş' } as Record<number, string>)[i] : undefined;
      assert.equal(decoded.charAt(i), special !== undefined ? special : String.fromCharCode(i), `${name}: 0x${i.toString(16)}`);
    }
    assert.equal(xmlEncodingRead(bytes), true, name);
  }
  const ascii = new TextEncoder().encode('<?xml version="1.0" encoding="us-ascii"?><a>x</a>');
  assert.equal(decodeXmlBytes(ascii), '<?xml version="1.0" encoding="us-ascii"?><a>x</a>');
  assert.equal(decodeXmlBytes(Uint8Array.from([...ascii.slice(0, -4), 0xe9, ...ascii.slice(-4)])).slice(-6), 'x\uFFFD</a>');
  // UTF-8 as ever: with a declaration, without one, with a byte order mark (which wins over what is declared).
  const utf8 = new TextEncoder().encode(`<a>${TURKISH}</a>`);
  assert.equal(decodeXmlBytes(utf8), `<a>${TURKISH}</a>`);
  assert.equal(xmlEncodingRead(utf8), true);
  const marked = new Uint8Array([0xef, 0xbb, 0xbf, ...new TextEncoder().encode(declare(`<a>${TURKISH}</a>`, 'ISO-8859-1'))]);
  assert.match(decodeXmlBytes(marked), new RegExp(TURKISH));
  // An encoding that is not read here: taken as UTF-8, as it always was, and said not to be read.
  for (const name of ['windows-1251', 'Shift_JIS', 'ISO-8859-2', 'x-unknown']) {
    const other = new TextEncoder().encode(declare('<a>x</a>', name));
    assert.equal(xmlEncodingRead(other), false, name);
    assert.equal(parseXml(other).root.name, 'a');
  }
});

test('xml encoding: the parser and the scanner read a part in the encoding it declares', () => {
  for (const [encoding, turkish] of [
    ['windows-1254', true],
    ['ISO-8859-9', true],
    ['ISO-8859-1', false],
    ['Windows-1252', false],
  ] as const) {
    const said = turkish ? TURKISH : 'Café üö ÿ';
    const bytes = single(declare(`<a t="${said}">${said}</a>`, encoding), turkish);
    const root = parseXml(bytes).root;
    assert.equal(root.attrs.get('t'), said, encoding);
    assert.equal(root.children[0].type === 'text' && root.children[0].value, said, encoding);
    const scanner = new XmlScanner(bytes);
    assert.equal(scanner.next(), true);
    assert.equal(scanner.attr('t'), said, `${encoding}: the scanner`);
    assert.equal(scanner.readText(), said, `${encoding}: the scanner`);
  }
});

const SHEET = (rows: string) => `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData>${rows}</sheetData></worksheet>`;

test('xml encoding: a sheet and shared strings in windows-1254 are read, kept byte for byte, and written as UTF-8 once changed', async () => {
  const zip = ZipArchive.read(await richBook());
  const sheet = single(declare(SHEET(`<row r="1"><c r="A1" t="inlineStr"><is><t>${TURKISH}</t></is></c><c r="B1" t="s"><v>0</v></c><c r="C1"><v>7</v></c></row>`), 'windows-1254'), true);
  const strings = single(declare(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Şube ${TURKISH}</t></si></sst>`, 'ISO-8859-9'), true);
  zip.set('xl/worksheets/sheet2.xml', sheet);
  zip.set('xl/sharedStrings.xml', strings);
  zip.set('xl/_rels/workbook.xml.rels', new TextEncoder().encode((await text(await zip.write(), 'xl/_rels/workbook.xml.rels')).replace('</Relationships>', '<Relationship Id="rIdSst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')));
  zip.set('[Content_Types].xml', new TextEncoder().encode((await text(await zip.write(), '[Content_Types].xml')).replace('</Types>', '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')));
  const bytes = await zip.write();
  const doc = await load(bytes);
  const cells = (wb: typeof doc.wb) => (wb.sheets[1].rows.get(0)?.cells || []).map((cell) => cell.value);
  assert.deepEqual(cells(doc.wb), [TURKISH, `Şube ${TURKISH}`, 7]);
  same(await save(doc), bytes, 'nothing changed: the file as it was');
  // A cell of the sheet changes: the part is written as UTF-8 that says so, with every letter it had.
  doc.history.apply([{ type: 'setCells', sheet: 1, cells: [{ row: 0, col: 2, cell: { kind: 'number', value: 8, style: 0 } }] }], undefined, undefined);
  const out = await save(doc);
  // (The strings were not touched and are still in the encoding they declare.)
  assert.deepEqual((await packageProblems(out)).filter((line) => !/^xl\/sharedStrings\.xml: declares the encoding/.test(line)), []);
  const written = await text(out, 'xl/worksheets/sheet2.xml');
  assert.match(written, /^<\?xml version="1\.0" encoding="UTF-8"/);
  assert.ok(written.indexOf(TURKISH) > 0, 'the letters are in the part');
  same(await part(out, 'xl/sharedStrings.xml'), strings, 'the strings were not touched');
  assert.deepEqual(cells(await reopen(out)), [TURKISH, `Şube ${TURKISH}`, 8]);
});

test('xml encoding: the body of a document in windows-1254 is read, kept byte for byte, and written as UTF-8 once changed', async () => {
  const plain = await buildDocx({ body: `<w:p><w:r><w:t>${TURKISH}</w:t></w:r></w:p><w:p><w:r><w:t>iki</w:t></w:r></w:p>` });
  const zip = ZipArchive.read(plain);
  const body = single(declare(new TextDecoder().decode(await zip.read('word/document.xml')), 'windows-1254'), true);
  zip.set('word/document.xml', body);
  const bytes = await zip.write();
  const file = await open(bytes);
  const doc = await file.document();
  const texts = (model: typeof doc) => JSON.stringify(model.body).match(new RegExp(TURKISH, 'g'));
  assert.equal((texts(doc) || []).length >= 1, true, 'the letters are read');
  assert.equal(JSON.stringify(doc.body).indexOf('�'), -1);
  same(await file.save(), bytes, 'nothing changed: the file as it was');
});

test('xml encoding: a chart part in Latin-1 keeps its letters when it is changed and when it is copied', async () => {
  const first = single(declare(RICH.replace('>Old<', '>Café üö<'), 'ISO-8859-1'), false);
  const bytes = await richBook({ first });
  const doc = await load(bytes);
  assert.equal(sheetCharts(doc.wb, 0)[0].title, 'Café üö');
  doc.history.apply(updateChart(doc.wb, 0, 0, { legend: 'b' }), undefined, undefined);
  doc.history.apply(duplicateChart(doc.wb, 0, 0).ops, undefined, undefined);
  const out = await save(doc);
  assert.deepEqual(await packageProblems(out), []);
  for (const name of ['xl/charts/chart1.xml', 'xl/charts/chart3.xml']) {
    const written = await text(out, name);
    assert.match(written, /^<\?xml version="1\.0" encoding="UTF-8"/, name);
    assert.match(written, /<a:t>Café üö<\/a:t>/, name);
  }
  assert.deepEqual(sheetCharts(await reopen(out), 0).map((chart) => chart.title).filter((title) => title !== undefined), ['Café üö', 'Café üö']);
  // Not changed, the part keeps its bytes.
  const again = await load(bytes);
  again.history.apply(duplicateChart(again.wb, 0, 1).ops, undefined, undefined);
  same(await part(await save(again), 'xl/charts/chart1.xml'), first, 'the chart that was not touched');
});

test('xml encoding: a chart part in an encoding that is not read is neither changed nor copied', async () => {
  const first = new TextEncoder().encode(declare(RICH, 'windows-1251'));
  const doc = await load(await richBook({ first }));
  const chart = chartsOf(doc.wb.sheets[0])[0];
  assert.equal(chart.known, undefined);
  assert.equal(chart.loose, undefined);
  assert.equal(sheetCharts(doc.wb, 0)[0].editable, 'none');
  assert.throws(() => updateChart(doc.wb, 0, 0, { legend: 'b' }), /Cannot change this chart/);
  assert.throws(() => copyChart(doc.wb, 0, 0), /Cannot copy this chart/);
  assert.throws(() => duplicateChart(doc.wb, 0, 0), /Cannot copy this chart/);
});
