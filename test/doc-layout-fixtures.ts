// Helpers for the document layout tests: a measurer with round numbers and
// builders for small synthetic documents.
//
// With the styles below a character is 8 px wide (bold: 8.8) and a line 16 px
// high; the standard section is a page of 400 x 400 px with margins of 40 px,
// so a line holds 40 characters and a page 20 lines.

import { DocLayout } from '../src/layout/doc/layout';
import type { DocLayoutOptions, FontMetrics, LayoutFont, LayoutPage, LineItem, PageItem, RectItem, TextItem, TextMeasurer } from '../src/layout/doc/types';
import type { DocModel } from '../src/model/doc/types';
import { readSpec, type DocxSpec } from './docx-fixtures';

/** Every character is half the font size wide (bold: a tenth more); ascent 0.8, descent 0.2 of the size. */
export const fakeMeasurer: TextMeasurer = {
  width(font: LayoutFont, text: string): number {
    return text.length * font.size * 0.5 * (font.bold ? 1.1 : 1);
  },
  metrics(font: LayoutFont): FontMetrics {
    return { ascent: font.size * 0.8, descent: font.size * 0.2, lineGap: 0 };
  },
};

/** 12 point text, no spacing: 8 px a character, 16 px a line. Widow control is off unless a test turns it on. */
export const STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Test" w:hAnsi="Test" w:eastAsia="Test" w:cs="Test"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>' +
  '<w:pPrDefault><w:pPr><w:widowControl w:val="0"/></w:pPr></w:pPrDefault></w:docDefaults>';

/** Page of 400 x 400 px, margins of 40 px, header and footer 10 px from the edge. */
export const sectPr = (inner = '', attrs = '') =>
  `<w:sectPr${attrs}><w:pgSz w:w="6000" w:h="6000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/>${inner}</w:sectPr>`;

const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A run of text. */
export const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escape(text)}</w:t></w:r>`;

/** A paragraph: `content` is text (one run) or run XML when it starts with "<". */
export const p = (content = '', pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.charAt(0) === '<' ? content : content ? r(content) : ''}</w:p>`;

/** `count` words of `length` letters, numbered so that every word differs. */
export function words(count: number, length = 4, start = 0): string {
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    let word = String(start + i);
    while (word.length < length) word = 'w' + word;
    out.push(word);
  }
  return out.join(' ');
}

export interface Laid {
  doc: DocModel;
  layout: DocLayout;
  pages: LayoutPage[];
}

/** Builds the document (with the standard styles and section unless the spec has its own) and lays it out. */
export async function lay(spec: DocxSpec | string, options?: DocLayoutOptions, measurer: TextMeasurer = fakeMeasurer): Promise<Laid> {
  const full: DocxSpec = typeof spec === 'string' ? { body: spec } : { ...spec };
  if (full.styles === undefined) full.styles = STYLES;
  if (full.body !== undefined && full.body.indexOf('<w:sectPr') < 0) full.body += sectPr();
  const doc = await readSpec(full);
  const layout = new DocLayout(doc, measurer, options);
  const result = layout.layout();
  if (!result.done) throw new Error('layout did not finish');
  return { doc, layout, pages: layout.pages };
}

export const texts = (page: LayoutPage): TextItem[] => page.items.filter((item): item is TextItem => item.type === 'text');
export const lines = (page: LayoutPage): LineItem[] => page.items.filter((item): item is LineItem => item.type === 'line');
export const rects = (page: LayoutPage): RectItem[] => page.items.filter((item): item is RectItem => item.type === 'rect');

/** The text items of a page as "text@x,y" (rounded to hundredths), for compact expectations. */
export const at = (page: LayoutPage): string[] => texts(page).map((item) => `${item.text}@${round(item.x)},${round(item.y)}`);

export const round = (value: number) => Math.round(value * 100) / 100;

/** Baselines of the text of a page, in order, without repeats. */
export function baselines(page: LayoutPage): number[] {
  const out: number[] = [];
  for (const item of texts(page)) {
    const y = round(item.y);
    if (out.indexOf(y) < 0) out.push(y);
  }
  return out;
}

/** All text of a page that has a source, in painting order. */
export const pageText = (page: LayoutPage): string => texts(page).filter((item) => item.source && !item.repeat).map((item) => item.text).join('');

export const allItems = (pages: LayoutPage[]): PageItem[] => pages.reduce<PageItem[]>((out, page) => out.concat(page.items), []);

/** Settings of a document of Word 2013 or later. */
export const COMPAT15 = '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>';

/** `count` paragraphs of one line each: "L0", "L1", ... */
export function filler(count: number, prefix = 'L'): string {
  let out = '';
  for (let i = 0; i < count; i++) out += p(prefix + i);
  return out;
}

/** The first text item whose text starts with `start`; fails when there is none. */
export function find(page: LayoutPage, start: string): TextItem {
  for (const item of texts(page)) if (item.text.indexOf(start) === 0) return item;
  throw new Error(`no text item starts with "${start}" on the page: ${at(page).join(' | ')}`);
}

export const tc = (content: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${content.charAt(0) === '<' ? content : p(content)}</w:tc>`;
export const tr = (cells: string, trPr = '') => `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells}</w:tr>`;
/** Cell margins of 10 px left and right, so that positions are round numbers. */
export const CELL_MARGINS = '<w:tblCellMar><w:left w:w="150" w:type="dxa"/><w:right w:w="150" w:type="dxa"/></w:tblCellMar>';
/** A table with the given grid (twips); cell margins of 10 px unless tblPr says otherwise. */
export const tbl = (rows: string, grid: number[], tblPr = CELL_MARGINS) =>
  `<w:tbl><w:tblPr>${tblPr}</w:tblPr>${grid.length ? `<w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` : ''}${rows}</w:tbl>`;
