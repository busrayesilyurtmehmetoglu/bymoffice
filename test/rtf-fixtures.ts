// Helpers for the tests of the RTF reader. Every document is written out in
// the test that reads it; nothing here comes from a real file.

import { resolveParaProps, resolveRunProps } from '../src/formats/docx/resolve';
import { readRtf, type ReadRtfOptions } from '../src/formats/rtf/document';
import { paragraphText, type Block, type DocModel, type DrawingItem, type ParaProps, type Paragraph, type Run, type RunItem, type RunProps, type Table } from '../src/model/doc/types';

/** The bytes of RTF source: one byte for each character, which must be below 256. */
export function bytesOf(source: string): Uint8Array {
  const out = new Uint8Array(source.length);
  for (let i = 0; i < source.length; i++) {
    const code = source.charCodeAt(i);
    if (code > 255) throw new Error(`character ${code} at ${i} is not a byte`);
    out[i] = code;
  }
  return out;
}

export const read = (source: string, options?: ReadRtfOptions): DocModel => readRtf(bytesOf(source), options);

/** A document around body text, with a font table of a Western, a Turkish and a symbol font. */
export const doc = (body: string, head = ''): string =>
  '{\\rtf1\\ansi\\ansicpg1252\\deff0{\\fonttbl{\\f0\\froman\\fcharset0 Times New Roman;}{\\f1\\fswiss\\fcharset162 Arial;}{\\f2\\fnil\\fcharset2 Symbol;}}' +
  '{\\colortbl;\\red255\\green0\\blue0;\\red0\\green0\\blue255;\\red255\\green255\\blue0;}' +
  // An empty group between the two: a control word that ends the head must not run into the text.
  (head ? head + '{}' : '') +
  body +
  '}';

/** The paragraphs of blocks in document order, those in tables included. */
export function paragraphsOf(blocks: Block[], out: Paragraph[] = []): Paragraph[] {
  for (const block of blocks) {
    if (block.type === 'paragraph') out.push(block);
    else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) paragraphsOf(cell.blocks, out);
  }
  return out;
}

export const paragraphs = (model: DocModel): Paragraph[] => paragraphsOf(model.body.blocks);
export const texts = (model: DocModel): string[] => paragraphs(model).map(paragraphText);
export const tables = (model: DocModel): Table[] => model.body.blocks.filter((block): block is Table => block.type === 'table');
export const runsOf = (paragraph: Paragraph): Run[] => paragraph.children.filter((inline): inline is Run => inline.type === 'run');

/** The items of a paragraph, of all its runs. */
export function itemsOf(paragraph: Paragraph): RunItem[] {
  const out: RunItem[] = [];
  for (const run of runsOf(paragraph)) for (const item of run.content) out.push(item);
  return out;
}

export const drawingsOf = (paragraph: Paragraph): DrawingItem[] => itemsOf(paragraph).filter((item): item is DrawingItem => item.type === 'drawing');

/** The run of a paragraph that holds the text. */
export function runWith(paragraph: Paragraph, text: string): Run {
  for (const run of runsOf(paragraph)) {
    if (run.content.some((item) => (item.type === 'text' || item.type === 'instrText' || item.type === 'deletedText') && item.text.indexOf(text) >= 0)) return run;
  }
  throw new Error(`no run with "${text}"`);
}

/** Run properties in effect for the run that holds the text. */
export const effectiveRun = (model: DocModel, paragraph: Paragraph, text: string): RunProps => resolveRunProps(model, paragraph.props, runWith(paragraph, text).props);
export const effectivePara = (model: DocModel, paragraph: Paragraph): ParaProps => resolveParaProps(model, paragraph.props);

/** `count` hex digits pairs: the bytes 0, 1, 2, ... as hex. */
export function hexOf(bytes: Uint8Array): string {
  const digits = '0123456789abcdef';
  let out = '';
  for (let i = 0; i < bytes.length; i++) out += digits.charAt(bytes[i] >> 4) + digits.charAt(bytes[i] & 15);
  return out;
}

/** A PNG file header that states a size, followed by `extra` bytes of filler: enough for the reader, which never decodes pictures. */
export function fakePng(width: number, height: number, extra = 0): Uint8Array {
  const out = new Uint8Array(33 + extra);
  out.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  out[16] = width >>> 24;
  out[17] = (width >>> 16) & 255;
  out[18] = (width >>> 8) & 255;
  out[19] = width & 255;
  out[20] = height >>> 24;
  out[21] = (height >>> 16) & 255;
  out[22] = (height >>> 8) & 255;
  out[23] = height & 255;
  out[24] = 8;
  out[25] = 6;
  for (let i = 33; i < out.length; i++) out[i] = (i * 31) & 255;
  return out;
}
