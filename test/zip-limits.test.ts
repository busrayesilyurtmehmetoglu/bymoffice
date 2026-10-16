// What reading an archive may unpack: one part at most its limit (the bodies of sheets and
// documents and every file that is not XML 512 MB, any other XML part 64 MB), and the parts
// read from one archive together at most 256 times its size (64 MB at least). A small file
// that unpacks to gigabytes is refused before anything is inflated, with an error that does
// not name the file.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  MAX_UNPACKED_PART_BYTES,
  MAX_UNPACKED_XML_BYTES,
  MIN_UNPACK_BUDGET_BYTES,
  PartTooLargeError,
  ZipArchive,
  createBlank,
  open,
} from '../src/index';
import { unpackedLimitOf } from '../src/zip';
import { buildXlsx, utf8, worksheetXml } from './fixtures';

const decoder = new TextDecoder();
const MB = 1024 * 1024;

/** The archive with the size an entry unpacks to declared as `size` in its central directory header. */
function declareSize(bytes: Uint8Array, name: string, size: number): Uint8Array {
  const out = bytes.slice();
  const view = new DataView(out.buffer);
  const wanted = utf8(name);
  for (let p = 0; p + 46 <= out.length; p++) {
    if (view.getUint32(p, true) !== 0x02014b50) continue;
    const length = view.getUint16(p + 28, true);
    if (length !== wanted.length || !wanted.every((byte, i) => out[p + 46 + i] === byte)) continue;
    view.setUint32(p + 24, size, true);
    return out;
  }
  throw new Error(`no central header for ${name}`);
}

/** XML of a part with `mb` megabytes of a comment of spaces after its declaration: what packs a thousand times. */
function padded(xml: string, mb: number): Uint8Array {
  const at = xml.indexOf('?>') + 2;
  return utf8(`${xml.slice(0, at)}<!--${' '.repeat(mb * MB)}-->${xml.slice(at)}`);
}

const isTooLarge = (kind: 'part' | 'archive', ...unnamed: string[]) => (error: unknown) => {
  assert.ok(error instanceof PartTooLargeError, `a PartTooLargeError, not ${String(error)}`);
  assert.equal(error.name, 'PartTooLargeError');
  assert.equal(error.kind, kind);
  for (const word of unnamed) assert.ok(!error.message.includes(word), `the message does not name ${word}: ${error.message}`);
  return true;
};

test('unpackedLimitOf(): bodies, shared strings and files that are not XML get the large limit, other XML parts the small one', () => {
  for (const name of ['xl/worksheets/sheet1.xml', 'XL/Worksheets/Sheet12.XML', 'xl/sharedStrings.xml', 'xl/pivotCache/pivotCacheRecords1.xml', 'word/document.xml', 'word/document2.xml', 'xl/macrosheets/sheet1.xml', 'xl/media/image1.png', 'xl/vbaProject.bin', 'word/embeddings/Microsoft_Excel_Worksheet.xlsx']) {
    assert.equal(unpackedLimitOf(name), MAX_UNPACKED_PART_BYTES, name);
  }
  for (const name of ['xl/styles.xml', 'xl/workbook.xml', '[Content_Types].xml', '_rels/.rels', 'xl/drawings/vmlDrawing1.vml', 'xl/charts/chart1.xml', 'word/styles.xml', 'word/glossary/document.xml', 'docProps/core.xml']) {
    assert.equal(unpackedLimitOf(name), MAX_UNPACKED_XML_BYTES, name);
  }
});

test('a sheet that declares more than the limit of a body is refused before it is inflated', async () => {
  const bytes = declareSize(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }] }), 'xl/worksheets/sheet1.xml', MAX_UNPACKED_PART_BYTES + 1);
  await assert.rejects(ZipArchive.read(bytes).read('xl/worksheets/sheet1.xml'), isTooLarge('part', 'sheet1'));
  await assert.rejects(async () => (await open(bytes)).workbook(), isTooLarge('part', 'sheet1', 'worksheets'));
});

test('a document body that declares more than its limit is refused', async () => {
  const bytes = declareSize(await createBlank('docx'), 'word/document.xml', MAX_UNPACKED_PART_BYTES + 1);
  await assert.rejects(async () => {
    const file = await open(bytes);
    if (file.kind !== 'document') throw new Error('not a document');
    await file.document();
  }, isTooLarge('part', 'document.xml'));
});

test('a styles part that really unpacks to more than 64 MB is refused, from a file of a few hundred kilobytes', async () => {
  const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>';
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }], styles: decoder.decode(padded(styles, 65)) });
  assert.ok(bytes.length < 1 * MB, `the file has ${bytes.length} bytes`);
  await assert.rejects(async () => (await open(bytes)).workbook(), isTooLarge('part', 'styles'));
});

test('parts that together unpack to more than 256 times the file (and 64 MB) are refused', async () => {
  // Each sheet is under every limit of a part; the two together are not, for a file this small.
  const sheet = decoder.decode(padded(worksheetXml('<sheetData/>'), 40));
  const bytes = await buildXlsx({ sheets: [{ name: 'A', xml: sheet }, { name: 'B', xml: sheet }] });
  assert.ok(256 * bytes.length < MIN_UNPACK_BUDGET_BYTES, `the file has ${bytes.length} bytes`);
  const zip = ZipArchive.read(bytes);
  assert.equal((await zip.read('xl/worksheets/sheet1.xml')).length > 40 * MB, true);
  // The same part read again is not counted again.
  await zip.read('xl/worksheets/sheet1.xml');
  await assert.rejects(zip.read('xl/worksheets/sheet2.xml'), isTooLarge('archive', 'sheet2'));
  await assert.rejects(async () => (await open(bytes)).workbook(), isTooLarge('archive'));
});

test('what was written into an archive is not held to the limits of what the file brings', async () => {
  const zip = ZipArchive.read(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }] }));
  const large = padded('<?xml version="1.0"?><styleSheet/>', 70);
  zip.set('xl/styles.xml', large);
  await zip.write();
  // The write compressed it and keeps only the compressed bytes: reading it inflates them.
  assert.equal((await zip.read('xl/styles.xml')).length, large.length);
});

test('a workbook of usual size opens as before', async () => {
  const rows = Array.from({ length: 2000 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}"><v>${r}</v></c></row>`).join('');
  const wb = await (await open(await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`<sheetData>${rows}</sheetData>`) }] }))).workbook();
  assert.equal(wb.sheets[0].rows.size, 2000);
});
