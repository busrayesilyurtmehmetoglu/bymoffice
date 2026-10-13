// The pictures of a PDF: a JPEG file as it is (the DCTDecode filter reads
// it), or samples (grey or RGB, eight bits each) that are compressed when
// they are written, with an optional mask of how opaque each pixel is.
// Samples are stored with the predictors of PNG: every row is written as its
// difference to the pixel before or the row above, whichever leaves least,
// which is what makes soft pictures small.

import { zlibDeflate } from './flate';
import { PdfError } from './writer';

export type PdfImage =
  | { kind: 'jpeg'; data: Uint8Array; width: number; height: number; components: 1 | 3 | 4 }
  | { kind: 'samples'; data: Uint8Array; width: number; height: number; colors: 1 | 3; alpha?: Uint8Array }
  /** Samples that are compressed already (zlib), with the PNG predictors when `predictor` says so. */
  | { kind: 'packed'; data: Uint8Array; width: number; height: number; colors: 1 | 3; predictor: boolean };

export type PdfSamples = PdfImage & { kind: 'samples' };
export type PdfPacked = PdfImage & { kind: 'packed' };

/** How many markers are looked at before a file counts as no JPEG. */
const MAX_JPEG_SEGMENTS = 4096;

/**
 * Size and number of components of a JPEG file, from its frame header, and
 * whether it carries the marker of Adobe's programs (which store the four
 * components of a CMYK picture the other way round). Throws a PdfError for
 * bytes that are no JPEG, or one a PDF cannot show (two components, or more
 * than four).
 */
export function jpegInfo(data: Uint8Array): { width: number; height: number; components: 1 | 3 | 4; adobe: boolean } {
  const bad = (why: string): never => {
    throw new PdfError('invalid', `The picture is no JPEG file a PDF can hold: ${why}`);
  };
  const length = data ? data.length : 0;
  if (length < 4 || data[0] !== 0xff || data[1] !== 0xd8) bad('it does not start as one');
  let at = 2;
  let adobe = false;
  for (let segment = 0; segment < MAX_JPEG_SEGMENTS; segment++) {
    if (at + 4 > length) return bad('it ends before its frame header');
    if (data[at] !== 0xff) return bad('a marker is missing');
    // Any number of 0xFF may stand before a marker.
    let fill = 0;
    while (at < length && data[at] === 0xff) {
      at++;
      if (++fill > 65536) return bad('it holds no frame header');
    }
    if (at + 3 > length) return bad('it ends before its frame header');
    const marker = data[at++];
    // Markers without a segment.
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) continue;
    if (marker === 0xd9 || marker === 0xda) return bad('it holds no frame header');
    const size = (data[at] << 8) | data[at + 1];
    if (size < 2 || at + size > length) return bad('a segment does not fit');
    // SOF0 to SOF15, without the three markers in that range that are no frame headers.
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      if (size < 8) return bad('its frame header is cut short');
      const height = (data[at + 3] << 8) | data[at + 4];
      const width = (data[at + 5] << 8) | data[at + 6];
      const components = data[at + 7];
      if (data[at + 2] !== 8) return bad('it has not eight bits per sample');
      if (width < 1 || height < 1) return bad('it has no size');
      if (components !== 1 && components !== 3 && components !== 4) return bad(`it has ${components} components`);
      if (size < 8 + components * 3) return bad('its frame header is cut short');
      return { width, height, components, adobe };
    }
    // APP14, "Adobe".
    if (marker === 0xee && size >= 7 && data[at + 2] === 0x41 && data[at + 3] === 0x64 && data[at + 4] === 0x6f && data[at + 5] === 0x62 && data[at + 6] === 0x65) adobe = true;
    at += size;
  }
  return bad('it holds no frame header');
}

/**
 * The pixels of a canvas (RGBA, the paper behind them opaque) as what a page
 * of a PDF needs of them: nothing but the colour when all pixels have the
 * same one, grey samples when every pixel is grey, RGB samples otherwise.
 */
export function imageOfPixels(rgba: ArrayLike<number>, width: number, height: number): { image?: PdfImage; fill?: [number, number, number] } {
  if (!(width >= 1) || !(height >= 1) || Math.floor(width) !== width || Math.floor(height) !== height) throw new PdfError('invalid', 'A picture needs a width and a height in whole pixels');
  const count = width * height;
  if (!rgba || rgba.length < count * 4) throw new PdfError('invalid', `A picture of ${width} x ${height} needs ${count * 4} bytes of pixels`);
  const r0 = rgba[0];
  const g0 = rgba[1];
  const b0 = rgba[2];
  let uniform = true;
  let grey = true;
  for (let i = 0, at = 0; i < count; i++, at += 4) {
    const r = rgba[at];
    const g = rgba[at + 1];
    const b = rgba[at + 2];
    if (uniform && (r !== r0 || g !== g0 || b !== b0)) uniform = false;
    if (grey && (r !== g || g !== b)) grey = false;
    if (!uniform && !grey) break;
  }
  if (uniform) return { fill: [r0, g0, b0] };
  if (grey) {
    const data = new Uint8Array(count);
    for (let i = 0, at = 0; i < count; i++, at += 4) data[i] = rgba[at];
    return { image: { kind: 'samples', data, width, height, colors: 1 } };
  }
  const data = new Uint8Array(count * 3);
  for (let i = 0, at = 0, to = 0; i < count; i++, at += 4) {
    data[to++] = rgba[at];
    data[to++] = rgba[at + 1];
    data[to++] = rgba[at + 2];
  }
  return { image: { kind: 'samples', data, width, height, colors: 3 } };
}

/**
 * Samples with the predictors of PNG (what /Predictor 15 of a PDF reads):
 * every row starts with the number of its filter and holds, for every byte,
 * its difference to what the filter predicts - nothing (0), the pixel before
 * (1, Sub), the row above (2, Up), the mean of the two (3) or whichever of
 * the three neighbours is nearest to their trend (4, Paeth). For each row the
 * filter that leaves the smallest sum is taken; the first one where several do.
 */
export function pngFilter(samples: Uint8Array, width: number, height: number, colors: number): Uint8Array {
  const stride = width * colors;
  if (!(width >= 1) || !(height >= 1) || (colors !== 1 && colors !== 3) || !samples || samples.length !== stride * height) throw new PdfError('invalid', `Samples of ${width} x ${height} with ${colors} components need ${stride * height} bytes`);
  const out = new Uint8Array((stride + 1) * height);
  const rows = [new Uint8Array(stride), new Uint8Array(stride), new Uint8Array(stride), new Uint8Array(stride), new Uint8Array(stride)];
  const sums = [0, 0, 0, 0, 0];
  for (let y = 0; y < height; y++) {
    const at = y * stride;
    const above = at - stride;
    sums[0] = sums[1] = sums[2] = sums[3] = sums[4] = 0;
    for (let i = 0; i < stride; i++) {
      const value = samples[at + i];
      const a = i >= colors ? samples[at + i - colors] : 0;
      const b = y > 0 ? samples[above + i] : 0;
      const c = y > 0 && i >= colors ? samples[above + i - colors] : 0;
      const guess = a + b - c;
      const pa = guess > a ? guess - a : a - guess;
      const pb = guess > b ? guess - b : b - guess;
      const pc = guess > c ? guess - c : c - guess;
      const paeth = pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      const none = value;
      const sub = (value - a) & 255;
      const up = (value - b) & 255;
      const mean = (value - ((a + b) >> 1)) & 255;
      const near = (value - paeth) & 255;
      rows[0][i] = none;
      rows[1][i] = sub;
      rows[2][i] = up;
      rows[3][i] = mean;
      rows[4][i] = near;
      // As signed bytes: 255 is as near to nothing as 1 is.
      sums[0] += none < 128 ? none : 256 - none;
      sums[1] += sub < 128 ? sub : 256 - sub;
      sums[2] += up < 128 ? up : 256 - up;
      sums[3] += mean < 128 ? mean : 256 - mean;
      sums[4] += near < 128 ? near : 256 - near;
    }
    let best = 0;
    for (let f = 1; f < 5; f++) if (sums[f] < sums[best]) best = f;
    const to = y * (stride + 1);
    out[to] = best;
    out.set(rows[best], to + 1);
  }
  return out;
}

/**
 * Below this share of pixels that differ from their neighbour a picture is
 * one of text: mostly paper. Above PHOTO_SHARE it is a photograph.
 */
export const TEXT_SHARE = 0.25;
export const PHOTO_SHARE = 0.6;

/**
 * Samples, compressed: what a lossless picture is in the file. With the
 * predictors where they help: a picture with soft changes of colour gets a
 * good deal smaller with them, a picture of text - long stretches of paper,
 * and letters that repeat - larger (about a tenth, as measured on pages of
 * text), so that one is compressed as it is. `predictor` says which, when the
 * caller knows; `native: false` keeps to the engine's own compression.
 */
export async function packSamples(image: PdfSamples, native = true, predictor?: boolean): Promise<PdfPacked> {
  const predicted = predictor === undefined ? changeShare(image) >= TEXT_SHARE : predictor;
  const data = await zlibDeflate(predicted ? pngFilter(image.data, image.width, image.height, image.colors) : image.data, native);
  return { kind: 'packed', data, width: image.width, height: image.height, colors: image.colors, predictor: predicted };
}

/** Rows that are looked at to tell what kind of picture a page is. */
const SHARE_ROWS = 512;

/**
 * The share of the pixels of a picture that differ from the pixel to their
 * left, between 0 and 1. A page of text is mostly paper, and its letters have
 * few edges for their area: a small share, and a picture that compresses
 * well as it is. In a photograph nearly every pixel differs from the one
 * before it: a large share, and a picture for a JPEG.
 */
export function changeShare(image: PdfSamples): number {
  const { width, height, colors, data } = image;
  if (!(width > 1) || !(height >= 1)) return 0;
  const step = Math.max(1, Math.floor(height / SHARE_ROWS));
  let changed = 0;
  let looked = 0;
  for (let y = 0; y < height; y += step) {
    const row = y * width * colors;
    for (let i = colors, end = width * colors; i < end; i += colors) {
      if (data[row + i] !== data[row + i - colors] || (colors === 3 && (data[row + i + 1] !== data[row + i - 2] || data[row + i + 2] !== data[row + i - 1]))) changed++;
      looked++;
    }
  }
  return looked > 0 ? changed / looked : 0;
}
