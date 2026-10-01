import { ZipArchive } from '../src/zip';

const encoder = new TextEncoder();
export const utf8 = (text: string) => encoder.encode(text);

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Smallest package that is recognisable as an xlsx workbook. */
export async function minimalXlsx(): Promise<Uint8Array> {
  const zip = new ZipArchive();
  zip.set(
    '[Content_Types].xml',
    utf8(
      `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '</Types>',
    ),
  );
  zip.set(
    '_rels/.rels',
    utf8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>',
    ),
  );
  zip.set(
    'xl/workbook.xml',
    utf8(
      `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        '<sheets><sheet name="Sayfa1" sheetId="1" r:id="rId1"/></sheets></workbook>',
    ),
  );
  zip.set(
    'xl/_rels/workbook.xml.rels',
    utf8(
      `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>' +
        '</Relationships>',
    ),
  );
  zip.set('xl/worksheets/sheet1.xml', utf8(`${XML}<worksheet><sheetData/></worksheet>`));
  // A part the engine knows nothing about; it must survive a save untouched.
  zip.set('customXml/unknown.bin', new Uint8Array([0, 1, 2, 3, 255, 254, 253]));
  return zip.write();
}

/**
 * Version 3 compound file with one stream in regular sectors and one in the
 * mini stream. Layout: sector 0 FAT, 1 directory, 2 mini FAT, 3 mini stream,
 * 4..12 the "Workbook" stream.
 */
export function minimalCfb(big: Uint8Array, small: Uint8Array): Uint8Array {
  const SECTOR = 512;
  const END = 0xfffffffe;
  const FREE = 0xffffffff;
  if (big.length !== 9 * SECTOR || small.length > SECTOR) throw new Error('fixture sizes are fixed');

  const out = new Uint8Array(SECTOR * 14);
  const view = new DataView(out.buffer);
  out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x3e, true);
  view.setUint16(26, 3, true);
  view.setUint16(28, 0xfffe, true);
  view.setUint16(30, 9, true);
  view.setUint16(32, 6, true);
  view.setUint32(44, 1, true); // FAT sector count
  view.setUint32(48, 1, true); // first directory sector
  view.setUint32(56, 4096, true); // mini stream cutoff
  view.setUint32(60, 2, true); // first mini FAT sector
  view.setUint32(64, 1, true);
  view.setUint32(68, END, true); // no DIFAT sectors
  view.setUint32(76, 0, true);
  for (let i = 1; i < 109; i++) view.setUint32(76 + i * 4, FREE, true);

  const sector = (n: number) => (n + 1) * SECTOR;

  const fat = sector(0);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(fat + i * 4, FREE, true);
  view.setUint32(fat, 0xfffffffd, true);
  for (const s of [1, 2, 3, 12]) view.setUint32(fat + s * 4, END, true);
  for (let s = 4; s < 12; s++) view.setUint32(fat + s * 4, s + 1, true);

  const entry = (index: number, name: string, type: number, child: number, right: number, start: number, size: number) => {
    const p = sector(1) + index * 128;
    for (let i = 0; i < name.length; i++) view.setUint16(p + i * 2, name.charCodeAt(i), true);
    view.setUint16(p + 64, name.length ? (name.length + 1) * 2 : 0, true);
    view.setUint8(p + 66, type);
    view.setUint32(p + 68, FREE, true);
    view.setUint32(p + 72, right, true);
    view.setUint32(p + 76, child, true);
    view.setUint32(p + 116, start, true);
    view.setUint32(p + 120, size, true);
  };
  entry(0, 'Root Entry', 5, 1, FREE, 3, SECTOR);
  entry(1, 'Workbook', 2, FREE, 2, 4, big.length);
  entry(2, 'Small', 2, FREE, FREE, 0, small.length);
  entry(3, '', 0, FREE, FREE, 0, 0);

  const miniFat = sector(2);
  const miniSectors = Math.ceil(small.length / 64);
  for (let i = 0; i < SECTOR / 4; i++) view.setUint32(miniFat + i * 4, FREE, true);
  for (let i = 0; i < miniSectors; i++) view.setUint32(miniFat + i * 4, i === miniSectors - 1 ? END : i + 1, true);

  out.set(small, sector(3));
  out.set(big, sector(4));
  return out;
}
