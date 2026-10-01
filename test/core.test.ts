import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile, OpcPackage, UnsupportedFileError, ZipArchive, open, parseXml, serializeXml } from '../src/index';
import { minimalCfb, minimalXlsx, utf8 } from './fixtures';

test('zip: written archive reads back with the same content and order', async () => {
  const zip = new ZipArchive();
  const repetitive = utf8('abc'.repeat(1000));
  const incompressible = new Uint8Array([7, 200, 13, 99]);
  zip.set('b.txt', repetitive);
  zip.set('a/ğüş.bin', incompressible);

  const again = ZipArchive.read(await zip.write());
  assert.deepEqual(again.names(), ['b.txt', 'a/ğüş.bin']);
  assert.deepEqual(await again.read('b.txt'), repetitive);
  assert.deepEqual(await again.read('a/ğüş.bin'), incompressible);
  assert.equal(again.entry('b.txt')?.method, 8);
  assert.equal(again.entry('a/ğüş.bin')?.method, 0);
});

test('zip: untouched entries keep their original compressed bytes', async () => {
  const zip = new ZipArchive();
  zip.set('keep.txt', utf8('keep '.repeat(500)));
  zip.set('change.txt', utf8('old'));
  const first = ZipArchive.read(await zip.write());
  const before = first.entry('keep.txt')!.raw.slice();

  first.set('change.txt', utf8('new content'));
  const second = ZipArchive.read(await first.write());
  assert.deepEqual(second.entry('keep.txt')!.raw, before);
  assert.deepEqual(await second.read('change.txt'), utf8('new content'));
  assert.deepEqual(second.names(), ['keep.txt', 'change.txt']);
});

test('zip: corrupted content is detected', async () => {
  const zip = new ZipArchive();
  zip.set('x.bin', new Uint8Array([1, 2, 3, 4]));
  const bytes = await zip.write();
  bytes[30 + 'x.bin'.length] ^= 0xff;
  await assert.rejects(ZipArchive.read(bytes).read('x.bin'), /CRC/);
});

test('xml: parse then serialize keeps names, attribute order, entities and whitespace', () => {
  const source =
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<w:doc xmlns:w="urn:w" b="2" a="1 &amp; &quot;x&quot;&#10;y"><!-- note -->' +
    '<w:t xml:space="preserve">  a &lt; b &#x15F; </w:t><w:br/><![CDATA[<raw>]]></w:doc>';
  const doc = parseXml(source);
  assert.deepEqual([...doc.root.attrs.keys()], ['xmlns:w', 'b', 'a']);
  assert.equal(doc.root.attrs.get('a'), '1 & "x"\ny');
  assert.deepEqual(parseXml(serializeXml(doc)), doc);
  assert.equal(
    serializeXml(doc),
    '<?xml version="1.0" encoding="UTF-8"?>\n' +
      '<w:doc xmlns:w="urn:w" b="2" a="1 &amp; &quot;x&quot;&#10;y"><!-- note -->' +
      '<w:t xml:space="preserve">  a &lt; b ş </w:t><w:br/><![CDATA[<raw>]]></w:doc>',
  );
});

test('xml: malformed input and DTDs are rejected', () => {
  assert.throws(() => parseXml('<a><b></a>'), /closing tag/);
  assert.throws(() => parseXml('<a>'), /unclosed/);
  assert.throws(() => parseXml('<a/><b/>'), /more than one root/);
  assert.throws(() => parseXml('<!DOCTYPE a [<!ENTITY x "y">]><a/>'), /DTD/);
  assert.throws(() => parseXml('<a>&nope;</a>'), /Unknown XML entity/);
});

test('opc: content types, relationships and main part', async () => {
  const pkg = await OpcPackage.open(await minimalXlsx());
  assert.equal(await pkg.mainPart(), '/xl/workbook.xml');
  assert.match(pkg.contentTypeOf('/xl/workbook.xml')!, /spreadsheetml\.sheet\.main/);
  assert.equal(pkg.contentTypeOf('/xl/worksheets/sheet1.xml'), 'application/xml');
  assert.equal(pkg.contentTypeOf('/customXml/unknown.bin'), undefined);

  const rels = await pkg.relationships('/xl/workbook.xml');
  assert.equal(rels[0].part, '/xl/worksheets/sheet1.xml');
  assert.deepEqual({ external: rels[1].external, part: rels[1].part }, { external: true, part: null });
  assert.deepEqual(await pkg.relationships('/xl/worksheets/sheet1.xml'), []);
});

test('opc: editing one part leaves every other part byte-identical', async () => {
  const original = await OpcPackage.open(await minimalXlsx());
  const edited = await OpcPackage.open(await minimalXlsx());

  const sheet = await edited.readXml('/xl/worksheets/sheet1.xml');
  sheet.root.attrs.set('edited', 'yes');
  edited.writePart('/xl/worksheets/sheet1.xml', sheet);
  edited.writePart('/xl/media/image1.png', new Uint8Array([137, 80, 78, 71]), 'image/png');

  const saved = await OpcPackage.open(await edited.save());
  assert.deepEqual(saved.partNames(), [...original.partNames(), '/xl/media/image1.png']);
  assert.equal(saved.contentTypeOf('/xl/media/image1.png'), 'image/png');
  assert.equal((await saved.readXml('/xl/worksheets/sheet1.xml')).root.attrs.get('edited'), 'yes');
  for (const part of original.partNames()) {
    if (part === '/xl/worksheets/sheet1.xml') continue;
    assert.deepEqual(await saved.readPart(part), await original.readPart(part), part);
  }
});

test('cfb: streams in regular sectors and in the mini stream', () => {
  const big = Uint8Array.from({ length: 9 * 512 }, (_, i) => (i * 7) & 0xff);
  const small = Uint8Array.from({ length: 100 }, (_, i) => 255 - i);
  const cfb = CfbFile.read(minimalCfb(big, small));
  assert.deepEqual(
    cfb.entries.map((e) => [e.path, e.type, e.size]),
    [
      ['Workbook', 'stream', big.length],
      ['Small', 'stream', small.length],
    ],
  );
  assert.deepEqual(cfb.stream('Workbook'), big);
  assert.deepEqual(cfb.stream('Small'), small);
  assert.throws(() => cfb.stream('Missing'), /not found/);
});

test('open: format is detected from content', async () => {
  const xlsx = await open(await minimalXlsx());
  assert.deepEqual([xlsx.format, xlsx.kind, xlsx.container], ['xlsx', 'spreadsheet', 'ooxml']);

  const xls = await open(minimalCfb(new Uint8Array(9 * 512), new Uint8Array(10)));
  assert.deepEqual([xls.format, xls.kind, xls.container], ['xls', 'spreadsheet', 'cfb']);

  await assert.rejects(open(utf8('just some text')), UnsupportedFileError);
});
