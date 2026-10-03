// What a copy of a part of a document puts on the clipboard, and what the
// editor keeps of it to paste it back with its formatting.
//
// The HTML is built here, piece by piece, from text that is escaped: nothing
// of a document reaches the clipboard as markup of its own. HTML that comes
// from the clipboard is never parsed; it is only searched for the mark that
// says the copy was made by this editor.

import { findWrapper, paragraphText, type Block, type DocModel, type Paragraph } from '../model/doc/types';
import { applyDocOps, asRange, deleteRange, itemLength, paragraphsBetween, rangeFormat, resolvePosition, storyOf, type DocPosition, type DocRange, type RunFormatPatch } from '../ops/doc/index';

/** Text in one formatting. A line break inside a paragraph is "\v", as insertText() takes it. */
export interface SliceRun {
  text: string;
  format: RunFormatPatch;
}

export interface SliceParagraph {
  runs: SliceRun[];
  /** Alignment and style of the paragraph, for the paragraphs whose end is part of the copy. */
  align?: string;
  style?: string;
}

/** A part of a document as the editor pastes it back: paragraphs of formatted text. The last one is without its end. */
export interface DocSlice {
  paragraphs: SliceParagraph[];
  /** The cells of a table that were selected, by row: pasted as a table. */
  table?: string[][];
}

/** A copy larger than this is kept as plain text only. */
const MAX_SLICE_PARAGRAPHS = 5000;
const MAX_SLICE_RUNS = 20000;

const escapeHtml = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** A font name as it may stand in a style attribute: letters, digits, spaces and hyphens only. */
const safeFont = (name: string): string => name.replace(/[^A-Za-z0-9 \-À-￿]/g, '').slice(0, 64);

/** The runs of a paragraph between two offsets, each with the formatting in effect for it. */
function runsOf(doc: DocModel, story: string, path: number[], paragraph: Paragraph, from: number, to: number, budget: { runs: number }): SliceRun[] | undefined {
  const text = paragraphText(paragraph);
  const out: SliceRun[] = [];
  let at = 0;
  const add = (start: number, end: number): boolean => {
    const a = Math.max(start, from);
    const b = Math.min(end, to);
    if (b <= a) return true;
    if (--budget.runs < 0) return false;
    const values = rangeFormat(doc, { anchor: { story, path, offset: a }, focus: { story, path, offset: b } });
    const format: RunFormatPatch = {
      bold: values.bold === true,
      italic: values.italic === true,
      underline: values.underline === true,
      strike: values.strike === true,
    };
    if (values.font !== undefined) format.font = values.font;
    if (values.size !== undefined) format.size = values.size;
    if (values.color !== undefined) format.color = values.color;
    if (values.highlight !== undefined) format.highlight = values.highlight;
    if (values.vertAlign !== undefined) format.vertAlign = values.vertAlign;
    const piece = text.slice(a, b).replace(/\n/g, '\v');
    const last = out[out.length - 1];
    if (last && JSON.stringify(last.format) === JSON.stringify(format)) last.text += piece;
    else out.push({ text: piece, format });
    return true;
  };
  for (const inline of paragraph.children) {
    if (inline.type !== 'run' && inline.type !== 'math') continue;
    if (inline.wrap && (findWrapper(inline, 'del') || findWrapper(inline, 'moveFrom'))) continue;
    let length = 0;
    if (inline.type === 'math') length = inline.text.length;
    else if (inline.fieldCode !== true) {
      const code = inline.fieldCode;
      for (let k = 0; k < inline.content.length; k++) if (!(code !== undefined && code[k])) length += itemLength(inline.content[k]);
    }
    if (length > 0 && !add(at, at + length)) return undefined;
    at += length;
    if (at >= to) break;
  }
  return out;
}

/**
 * The part of a story between two positions (`start` before `end`) with its
 * formatting; undefined when it is too large to keep or cannot be read.
 */
export function sliceOf(doc: DocModel, start: DocPosition, end: DocPosition): DocSlice | undefined {
  try {
    const paths = paragraphsBetween(doc, start.story, start.path, end.path, MAX_SLICE_PARAGRAPHS);
    const budget = { runs: MAX_SLICE_RUNS };
    const paragraphs: SliceParagraph[] = [];
    for (let i = 0; i < paths.length; i++) {
      const at = resolvePosition(doc, { story: start.story, path: paths[i], offset: 0 });
      const from = i === 0 ? start.offset : 0;
      const to = i === paths.length - 1 ? end.offset : at.length;
      const runs = runsOf(doc, start.story, paths[i], at.paragraph, from, to, budget);
      if (!runs) return undefined;
      const paragraph: SliceParagraph = { runs };
      if (i < paths.length - 1) {
        const values = rangeFormat(doc, { story: start.story, path: paths[i], offset: 0 });
        if (values.align !== undefined) paragraph.align = values.align;
        if (values.paraStyle !== undefined && values.paraStyle !== '') paragraph.style = values.paraStyle;
      }
      paragraphs.push(paragraph);
    }
    return { paragraphs };
  } catch {
    return undefined;
  }
}

/** What a story holds, counted: the length of its text, and everything in it that is more than text in some formatting. */
interface Census {
  text: number;
  other: string;
}

/** Items a copy carries as characters of its text, and a hint of the layout that stands for nothing. */
const AS_TEXT = new Set<string>(['text', 'tab', 'noBreakHyphen', 'lastRenderedPageBreak']);
/** Inlines that are never removed with the text around them, or mean nothing by themselves. */
const NO_CONTENT = new Set<string>(['proofError', 'empty']);

function census(blocks: Block[]): Census {
  const counts = new Map<string, number>();
  const count = (key: string, by = 1): void => {
    counts.set(key, (counts.get(key) || 0) + by);
  };
  let text = 0;
  const walk = (list: Block[], depth: number): void => {
    for (const block of list) {
      if (block.type === 'table') {
        if (depth < 64) for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks, depth + 1);
        continue;
      }
      if (block.type !== 'paragraph') continue;
      text += paragraphText(block).length;
      for (const inline of block.children) {
        // Text inside a link, a field, a content control or a tracked change is more than its characters.
        let wrapped = '';
        for (let w = inline.wrap; w; w = w.parent) wrapped += w.kind + ' ';
        if (inline.type !== 'run') {
          if (!NO_CONTENT.has(inline.type)) count(inline.type);
          if (inline.type === 'math' && wrapped !== '') count('in ' + wrapped, inline.text.length);
          continue;
        }
        const code = inline.fieldCode;
        for (let k = 0; k < inline.content.length; k++) {
          const item = inline.content[k];
          if (code === true || (code !== undefined && code[k])) {
            count('code');
            continue;
          }
          if (item.type === 'break') {
            if (item.kind !== 'line') count('break ' + item.kind);
          } else if (!AS_TEXT.has(item.type)) count(item.type);
          if (wrapped !== '') count('in ' + wrapped, Math.max(1, itemLength(item)));
        }
      }
    }
  };
  walk(blocks, 0);
  const keys: string[] = [];
  counts.forEach((_, key) => keys.push(key));
  keys.sort();
  return { text, other: keys.map((key) => `${key}=${counts.get(key)}`).join(';') };
}

/**
 * Whether a cut of a range would take something out of the document that the
 * copy of it (`slice`; undefined when only plain text is on the clipboard)
 * does not hold: a picture, a note reference, a field, an equation, a link, a
 * content control, a tracked change, a page break. The deletion itself is the
 * judge: it is made, what the story holds is counted before and after, and it
 * is taken back, so that nothing here has to know what a deletion removes.
 * Throws the RangeError of a place that cannot be edited.
 */
export function cutLoses(doc: DocModel, range: DocRange, slice: DocSlice | undefined): boolean {
  const story = asRange(range).focus.story;
  const before = census(storyOf(doc, story).blocks);
  const edit = deleteRange(doc, range);
  if (edit.ops.length === 0) return false;
  const inverse = applyDocOps(doc, edit.ops);
  let after: Census;
  try {
    after = census(storyOf(doc, story).blocks);
  } finally {
    applyDocOps(doc, inverse);
  }
  if (after.other !== before.other) return true;
  if (slice && !slice.table) {
    // Every character that goes must be one the copy holds.
    let carried = 0;
    for (const paragraph of slice.paragraphs) for (const run of paragraph.runs) carried += run.text.length;
    if (before.text - after.text !== carried) return true;
  }
  return false;
}

const ALIGN: { [name: string]: string } = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

function runHtml(run: SliceRun): string {
  let html = escapeHtml(run.text).replace(/\t/g, '&#9;').replace(/\v/g, '<br>');
  if (html === '') return '';
  const format = run.format;
  let style = '';
  if (typeof format.color === 'string' && /^[0-9A-Fa-f]{6}$/.test(format.color)) style += `color:#${format.color};`;
  if (typeof format.size === 'number' && format.size > 0 && format.size < 3277) style += `font-size:${format.size / 2}pt;`;
  if (typeof format.font === 'string' && safeFont(format.font) !== '') style += `font-family:'${safeFont(format.font)}';`;
  if (format.vertAlign === 'superscript') html = `<sup>${html}</sup>`;
  else if (format.vertAlign === 'subscript') html = `<sub>${html}</sub>`;
  if (format.strike) html = `<s>${html}</s>`;
  if (format.underline) html = `<u>${html}</u>`;
  if (format.italic) html = `<i>${html}</i>`;
  if (format.bold) html = `<b>${html}</b>`;
  return style ? `<span style="${style}">${html}</span>` : html;
}

/** A slice as HTML: paragraphs with bold, italic, underline, colour and size; a selection of cells as a table. */
export function sliceHtml(slice: DocSlice, mark: string): string {
  let html = '';
  if (slice.table) {
    html += '<table>';
    for (const row of slice.table) {
      html += '<tr>';
      for (const cell of row) html += `<td>${escapeHtml(cell).replace(/\n/g, '<br>')}</td>`;
      html += '</tr>';
    }
    html += '</table>';
  } else {
    for (const paragraph of slice.paragraphs) {
      const align = paragraph.align !== undefined ? ALIGN[paragraph.align] : undefined;
      let inner = '';
      for (const run of paragraph.runs) inner += runHtml(run);
      html += `<p${align ? ` style="text-align:${align}"` : ''}>${inner === '' ? '<br>' : inner}</p>`;
    }
  }
  // The mark is a string of letters and digits made by the editor.
  return `<div data-bym-doc-copy="${escapeHtml(mark)}">${html}</div>`;
}

/** Whether HTML from the clipboard carries the mark of a copy. The HTML is searched, never parsed. */
export function hasCopyMark(html: string | undefined, mark: string): boolean {
  return typeof html === 'string' && mark !== '' && html.indexOf(`data-bym-doc-copy="${mark}"`) >= 0;
}
