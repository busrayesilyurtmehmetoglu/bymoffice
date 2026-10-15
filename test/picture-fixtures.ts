// Pictures made by code for the tests of adding, moving and deleting pictures in
// a workbook: PNG files that decode (a solid colour), and the headers of JPEG,
// GIF and BMP files with a size in them. Also a workbook with pictures of each
// kind of anchor, laid out as Excel lays them out.

import { deflateSync } from 'node:zlib';
import { ZipArchive } from '../src/zip';
import { buildXlsx, utf8, worksheetXml } from './fixtures';
import { absolute, drawingXml, oneCell, pic, twoCell } from './sheet-drawing-fixtures';

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  for (let i = 0; i < 4; i++) out[4 + i] = type.charCodeAt(i);
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

/** A PNG of one colour, `width` by `height` pixels, that any decoder reads. */
export function png(width: number, height: number, rgb: [number, number, number] = [200, 30, 30]): Uint8Array {
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header[8] = 8;
  header[9] = 2;
  const row = new Uint8Array(1 + width * 3);
  for (let x = 0; x < width; x++) row.set(rgb, 1 + x * 3);
  const raw = new Uint8Array(row.length * height);
  for (let y = 0; y < height; y++) raw.set(row, y * row.length);
  const parts = [new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', new Uint8Array(deflateSync(raw))), chunk('IEND', new Uint8Array(0))];
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let at = 0;
  for (const part of parts) {
    out.set(part, at);
    at += part.length;
  }
  return out;
}

/** Only the header of a PNG: what is read of it says `width` by `height`. */
export function pngHeader(width: number, height: number): Uint8Array {
  const out = new Uint8Array(64);
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(out.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  return out;
}

/** The start of a JPEG with a frame header that says `width` by `height`. */
export function jpegHeader(width: number, height: number): Uint8Array {
  return new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 0, 16, 0x4a, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0, 0xff, 0xc0, 0, 11, 8, height >> 8, height & 255, width >> 8, width & 255, 1, 1, 0x11, 0, 0xff, 0xd9]);
}

export function gifHeader(width: number, height: number): Uint8Array {
  return new Uint8Array([0x47, 0x49, 0x46, 0x38, 0x39, 0x61, width & 255, width >> 8, height & 255, height >> 8, 0, 0, 0, 0x3b]);
}

export function bmpHeader(width: number, height: number): Uint8Array {
  const out = new Uint8Array(58);
  const view = new DataView(out.buffer);
  out[0] = 0x42;
  out[1] = 0x4d;
  view.setUint32(2, out.length, true);
  view.setUint32(10, 54, true);
  view.setUint32(14, 40, true);
  view.setInt32(18, width, true);
  view.setInt32(22, height, true);
  view.setUint16(26, 1, true);
  view.setUint16(28, 24, true);
  return out;
}

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * A workbook with one sheet ("Sayfa1") and a drawing of three pictures of `image1.png`: one that
 * moves and sizes with its cells (B2:D6), one that moves with its cell (F2, 50 by 40 pixels) and
 * one that stands where it is (at 400, 300; 60 by 30 pixels). The first picture carries content
 * the engine has no field for. With `shared`, the picture of a second relationship shows the same part.
 */
export async function pictureWorkbook(options: { sheet?: string; editAs?: string; shared?: boolean; drawing?: string; protection?: string } = {}): Promise<Uint8Array> {
  const image = png(4, 3);
  const odd = '<a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>';
  const first = pic('rIdI1', { name: 'Logo', descr: 'Bir logo' }).replace(`r:embed="rIdI1"/>`, `r:embed="rIdI1">${odd}</a:blip>`);
  const drawing =
    options.drawing ??
    twoCell([1, 1], [4, 6], first, options.editAs) +
      oneCell([5, 1], 50, 40, pic(options.shared ? 'rIdI2' : 'rIdI1', { name: 'İkinci' }).replace('cNvPr id="2"', 'cNvPr id="3"')) +
      absolute(400, 300, 60, 30, pic('rIdI1', { name: 'Üçüncü' }).replace('cNvPr id="2"', 'cNvPr id="4"'));
  const refs = '<drawing r:id="rIdD"/>';
  const bytes = await buildXlsx({ sheets: [{ name: 'Sayfa1', xml: worksheetXml((options.sheet ?? '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') + (options.protection ?? '') + refs) }] });
  const zip = ZipArchive.read(bytes);
  zip.set('xl/worksheets/_rels/sheet1.xml.rels', utf8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdD" Type="${NS_REL}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`));
  zip.set('xl/drawings/drawing1.xml', utf8(drawingXml(drawing)));
  const rels = `<Relationship Id="rIdI1" Type="${NS_REL}/image" Target="../media/image1.png"/>` + (options.shared ? `<Relationship Id="rIdI2" Type="${NS_REL}/image" Target="../media/image1.png"/>` : '');
  zip.set('xl/drawings/_rels/drawing1.xml.rels', utf8(`${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`));
  zip.set('xl/media/image1.png', image);
  const types = new TextDecoder().decode(await zip.read('[Content_Types].xml'));
  zip.set(
    '[Content_Types].xml',
    utf8(types.replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"').replace('</Types>', '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')),
  );
  return zip.write();
}
