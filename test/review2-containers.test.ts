// Second review of the containers: a part read as UTF-16 and written again says the encoding it
// is written in, and the list of content types names a part once.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { writeXlsx } from '../src/formats/xlsx/write';
import { OpcPackage, ZipArchive, open, parseXml, serializeXml } from '../src/index';
import { DocHistory, insertHyperlink } from '../src/ops/doc/index';
import { History } from '../src/ops/ops';
import { insertSheet } from '../src/ops/sheets';
import { buildDocx } from './docx-fixtures';
import { buildXlsx, minimalXlsx, worksheetXml } from './fixtures';

const decoder = new TextDecoder('utf-8', { fatal: true });

/** The text as UTF-16 (little endian) with a byte order mark. */
function utf16(text: string): Uint8Array {
  const bytes = new Uint8Array(2 + text.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let i = 0; i < text.length; i++) {
    bytes[2 + i * 2] = text.charCodeAt(i) & 0xff;
    bytes[3 + i * 2] = text.charCodeAt(i) >> 8;
  }
  return bytes;
}

/** The entries of the archive given as UTF-8 text re-encoded as UTF-16, their declaration saying so. */
async function inUtf16(bytes: Uint8Array, names: string[]): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  for (const name of names) {
    const text = decoder.decode(await zip.read(name));
    assert.ok(/encoding="UTF-8"/.test(text), name);
    zip.set(name, utf16(text.replace('encoding="UTF-8"', 'encoding="UTF-16"')));
  }
  return zip.write();
}

/**
 * What a parser that believes the part does: no byte order mark means the bytes are in the
 * encoding the declaration names, and they have to be valid in it.
 */
function consistent(bytes: Uint8Array, name: string): string {
  const marked = (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff);
  if (marked) return new TextDecoder(bytes[0] === 0xff ? 'utf-16le' : 'utf-16be', { fatal: true }).decode(bytes);
  const text = decoder.decode(bytes);
  const declared = /^<\?xml[^>]*?encoding\s*=\s*["']([^"']*)["']/.exec(text);
  if (declared) assert.match(declared[1], /^utf-?8$/i, `${name} is UTF-8 and says ${declared[1]}`);
  return text;
}

test('review 2: serializeXml names UTF-8 in a declaration that named another encoding, and adds none', () => {
  const doc = parseXml(utf16('<?xml version="1.0" encoding="UTF-16" standalone="yes"?><a b="ğ">ş</a>'));
  assert.equal(serializeXml(doc), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a b="ğ">ş</a>');
  assert.equal(serializeXml(parseXml(utf16("<?xml version='1.0' encoding = 'utf-16'?><a/>"))), "<?xml version='1.0' encoding = 'UTF-8'?>\n<a/>");
  // Left as written: UTF-8 in whatever spelling, no encoding, no declaration.
  for (const text of ['<?xml version="1.0" encoding="utf-8"?>\n<a/>', '<?xml version="1.0" standalone="yes"?>\n<a/>', '<a/>']) {
    assert.equal(serializeXml(parseXml(new TextEncoder().encode(text))), text);
  }
});

test('review 2: workbook parts in UTF-16 that are written again are UTF-8 and say so', async () => {
  const names = ['[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'];
  const bytes = await inUtf16(await buildXlsx({ sheets: [{ name: 'Kasım', xml: worksheetXml('<sheetData/>') }] }), names);
  const file = await open(bytes);
  const wb = await file.workbook();
  new History(wb).apply(insertSheet(wb, 1, { locale: 'tr' }), undefined, undefined);
  await writeXlsx(file.package!, wb);
  const saved = await file.package!.save();
  const zip = ZipArchive.read(saved);
  for (const name of names) {
    const text = consistent(await zip.read(name), name);
    parseXml(await zip.read(name));
    if (name === 'xl/workbook.xml') assert.ok(text.includes('name="Kasım"'), text);
  }
  const again = await (await open(saved)).workbook();
  assert.deepEqual(
    again.sheets.map((sheet) => sheet.name),
    ['Kasım', 'Sayfa1'],
  );
});

test('review 2: document relationships in UTF-16 that are written again are UTF-8 and say so', async () => {
  const name = 'word/_rels/document.xml.rels';
  const bytes = await inUtf16(await buildDocx({ body: '<w:p><w:r><w:t>linked text</w:t></w:r></w:p>' }), [name, '[Content_Types].xml']);
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const history = new DocHistory(doc);
  history.applyEdit(insertHyperlink(doc, { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 6 } }, 'https://example.com/ğ'), undefined);
  await writeDocx(pkg, doc);
  const saved = await pkg.save();
  const zip = ZipArchive.read(saved);
  for (const entry of zip.names()) if (/\.(xml|rels)$/.test(entry)) consistent(await zip.read(entry), entry);
  assert.ok(consistent(await zip.read(name), name).includes('https://example.com/'));
  const rels = await (await OpcPackage.open(saved)).relationships('/word/document.xml');
  assert.ok(rels.some((rel) => rel.external && rel.type.endsWith('/hyperlink')));
});

const overridesOf = async (pkg: OpcPackage) => {
  const zip = ZipArchive.read(await pkg.save());
  const text = decoder.decode(await zip.read('[Content_Types].xml'));
  return text.match(/<Override[^>]*>/g) || [];
};

test('review 2: writePart with another content type changes the declaration of the part instead of adding a second one', async () => {
  const pkg = await OpcPackage.open(await minimalXlsx());
  const macro = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
  pkg.writePart('/xl/workbook.xml', await pkg.readPart('/xl/workbook.xml'), macro);
  assert.deepEqual(await overridesOf(pkg), [`<Override PartName="/xl/workbook.xml" ContentType="${macro}"/>`]);
  assert.equal(pkg.contentTypeOf('/xl/workbook.xml'), macro);
  // A part whose type comes from its extension gets a declaration of its own, once.
  pkg.writePart('/xl/worksheets/sheet1.xml', await pkg.readPart('/xl/worksheets/sheet1.xml'), 'application/x-one');
  pkg.writePart('/xl/Worksheets/Sheet1.xml', await pkg.readPart('/xl/worksheets/sheet1.xml'), 'application/x-two');
  assert.deepEqual((await overridesOf(pkg)).slice(1), ['<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/x-two"/>']);
});

test('review 2: a part deleted and written again has the declaration it had, attributes that are not known included', async () => {
  const zip = ZipArchive.read(await minimalXlsx());
  const types = decoder.decode(await zip.read('[Content_Types].xml')).replace('</Types>', '<Override PartName="/customXml/Item1.xml" ContentType="application/x-item" vendor:note="kept" xmlns:vendor="urn:vendor"/></Types>');
  zip.set('[Content_Types].xml', new TextEncoder().encode(types));
  zip.set('customXml/Item1.xml', new TextEncoder().encode('<item/>'));
  const original = await zip.write();
  const pkg = await OpcPackage.open(original);
  const bytes = await pkg.readPart('/customXml/Item1.xml');
  assert.ok(pkg.deletePart('/customXml/Item1.xml'));
  assert.equal((await overridesOf(pkg)).length, 1);
  pkg.writePart('/customXml/Item1.xml', bytes, 'application/x-item');
  const saved = await pkg.save();
  assert.equal(decoder.decode(await ZipArchive.read(saved).read('[Content_Types].xml')), types);
});
