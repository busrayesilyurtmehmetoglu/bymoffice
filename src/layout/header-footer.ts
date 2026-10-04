// The code a sheet's header and footer are written in (SheetPageSetup.headerFooter):
// "&L", "&C" and "&R" start the left, centre and right section, "&P" is the page
// number, '&"Arial,Bold"' a font, "&12" a size, "&&" an ampersand, and so on.
//
// parseHeaderFooter() turns the code into runs of text and fields with their
// formatting; headerFooterLines() puts the values of a page into the fields and
// cuts the result into lines. Nothing here draws or measures.

/** Longest code that is read; Excel stops at 255 characters. */
export const MAX_HEADER_FOOTER_CHARS = 8192;
const MAX_RUNS = 512;
const MAX_LINES = 64;
/** Font sizes in points a code can state. */
const MIN_SIZE = 1;
const MAX_SIZE = 409;

export type HeaderFooterField = 'page' | 'pages' | 'date' | 'time' | 'file' | 'path' | 'sheet';

export interface HeaderFooterFormat {
  /** Font name; undefined is the default font of the workbook. */
  font?: string;
  /** Points; undefined is the default size. */
  size?: number;
  bold?: boolean;
  italic?: boolean;
  /** Number of strokes: 1, or 2 for a double underline. */
  underline?: 1 | 2;
  strike?: boolean;
  /** Set above or below the line, in a smaller size. */
  vertical?: 'super' | 'sub';
  /** "#RRGGBB". Colours of the theme are not resolved: the text keeps the default colour. */
  color?: string;
}

export interface HeaderFooterRun {
  /** The text, or what a field adds to its number ("&P+2"). */
  text?: string;
  field?: HeaderFooterField;
  /** Added to the number of a page field. */
  offset?: number;
  format: HeaderFooterFormat;
}

export interface HeaderFooterSections {
  left: HeaderFooterRun[];
  center: HeaderFooterRun[];
  right: HeaderFooterRun[];
}

/** What the fields of one page stand for. */
export interface HeaderFooterValues {
  page: number;
  pages: number;
  date: string;
  time: string;
  /** Name of the file, empty when it is not known. */
  file: string;
  /** Path of the file, empty when it is not known. */
  path: string;
  sheet: string;
}

export interface HeaderFooterPiece {
  text: string;
  format: HeaderFooterFormat;
}

const isDigit = (code: number): boolean => code >= 48 && code <= 57;
const HEX6 = /^[0-9A-Fa-f]{6}$/;

/**
 * Reads the code of a header or footer. Text before the first section code
 * belongs to the centre. A code that is not known is left out with its
 * letter; nothing of a code ever becomes text.
 */
export function parseHeaderFooter(code: string | undefined): HeaderFooterSections {
  const sections: HeaderFooterSections = { left: [], center: [], right: [] };
  if (typeof code !== 'string' || code.length === 0) return sections;
  const source = code.length > MAX_HEADER_FOOTER_CHARS ? code.slice(0, MAX_HEADER_FOOTER_CHARS) : code;
  const length = source.length;
  let target = sections.center;
  let format: HeaderFooterFormat = {};
  let text = '';
  let runs = 0;

  const flush = () => {
    if (text.length === 0) return;
    if (runs < MAX_RUNS) {
      target.push({ text, format });
      runs++;
    }
    text = '';
  };
  const field = (name: HeaderFooterField, offset?: number) => {
    flush();
    if (runs >= MAX_RUNS) return;
    const run: HeaderFooterRun = { field: name, format };
    if (offset) run.offset = offset;
    target.push(run);
    runs++;
  };
  // A format is never changed in place: the runs made so far keep theirs.
  const change = (patch: HeaderFooterFormat) => {
    flush();
    const next: HeaderFooterFormat = { ...format, ...patch };
    for (const key of Object.keys(next) as (keyof HeaderFooterFormat)[]) if (next[key] === undefined || next[key] === false) delete next[key];
    format = next;
  };

  let i = 0;
  while (i < length) {
    const ch = source.charAt(i);
    if (ch !== '&') {
      text += ch;
      i++;
      continue;
    }
    i++;
    if (i >= length) break;
    const letter = source.charAt(i);
    const code0 = source.charCodeAt(i);
    if (letter === '&') {
      text += '&';
      i++;
      continue;
    }
    if (letter === '"') {
      // &"font,style": up to the closing quote; without one the rest of the code is the name and is dropped.
      const end = source.indexOf('"', i + 1);
      const inner = end < 0 ? '' : source.slice(i + 1, end);
      i = end < 0 ? length : end + 1;
      if (end < 0) continue;
      const comma = inner.indexOf(',');
      const name = (comma < 0 ? inner : inner.slice(0, comma)).trim();
      const style = comma < 0 ? '' : inner.slice(comma + 1).toLowerCase();
      // "-" keeps the default font.
      const patch: HeaderFooterFormat = { font: name === '' || name === '-' ? undefined : name.slice(0, 128) };
      if (comma >= 0) {
        patch.bold = style.indexOf('bold') >= 0;
        patch.italic = style.indexOf('italic') >= 0 || style.indexOf('oblique') >= 0;
      }
      change(patch);
      continue;
    }
    if (isDigit(code0)) {
      // A size: all the digits that follow belong to the code, so none of them becomes text.
      let end = i;
      while (end < length && isDigit(source.charCodeAt(end))) end++;
      const size = end - i > 4 ? Infinity : Number(source.slice(i, end));
      i = end;
      if (size >= MIN_SIZE && size <= MAX_SIZE) change({ size });
      continue;
    }
    i++;
    switch (letter) {
      case 'L':
      case 'l':
        flush();
        target = sections.left;
        format = {};
        break;
      case 'C':
      case 'c':
        flush();
        target = sections.center;
        format = {};
        break;
      case 'R':
      case 'r':
        flush();
        target = sections.right;
        format = {};
        break;
      case 'P':
      case 'p': {
        // "&P+3" and "&P-1": the number of the page plus or minus a number.
        let offset = 0;
        const sign = source.charAt(i);
        if ((sign === '+' || sign === '-') && isDigit(source.charCodeAt(i + 1))) {
          let end = i + 1;
          while (end < length && end - i <= 6 && isDigit(source.charCodeAt(end))) end++;
          offset = Number(source.slice(i + 1, end)) * (sign === '-' ? -1 : 1);
          i = end;
        }
        field('page', offset);
        break;
      }
      case 'N':
      case 'n':
        field('pages');
        break;
      case 'D':
      case 'd':
        field('date');
        break;
      case 'T':
      case 't':
        field('time');
        break;
      case 'F':
      case 'f':
        field('file');
        break;
      case 'Z':
      case 'z':
        field('path');
        break;
      case 'A':
      case 'a':
        field('sheet');
        break;
      case 'B':
      case 'b':
        change({ bold: !format.bold });
        break;
      case 'I':
      case 'i':
        change({ italic: !format.italic });
        break;
      case 'U':
      case 'u':
        change({ underline: format.underline === 1 ? undefined : 1 });
        break;
      case 'E':
      case 'e':
        change({ underline: format.underline === 2 ? undefined : 2 });
        break;
      case 'S':
      case 's':
        change({ strike: !format.strike });
        break;
      case 'X':
      case 'x':
        change({ vertical: format.vertical === 'super' ? undefined : 'super' });
        break;
      case 'Y':
      case 'y':
        change({ vertical: format.vertical === 'sub' ? undefined : 'sub' });
        break;
      case 'K':
      case 'k': {
        // A colour: six hex digits, or a colour of the theme ("04+000": two digits, a sign, three digits), which is skipped.
        const hex = source.slice(i, i + 6);
        if (HEX6.test(hex)) {
          i += 6;
          change({ color: '#' + hex.toUpperCase() });
        } else if (/^[0-9]{2}[+-][0-9]{3}$/.test(hex)) {
          i += 6;
          change({ color: undefined });
        }
        break;
      }
      // "&G" is a picture, "&O" outline and "&H" shadow: none of them is drawn, and none of them is text.
      default:
        break;
    }
  }
  flush();
  return sections;
}

/** Whether a section has anything to show. */
export function headerFooterIsEmpty(sections: HeaderFooterSections): boolean {
  return sections.left.length === 0 && sections.center.length === 0 && sections.right.length === 0;
}

function fieldText(run: HeaderFooterRun, values: HeaderFooterValues): string {
  switch (run.field) {
    case 'page':
      return String(values.page + (run.offset || 0));
    case 'pages':
      return String(values.pages);
    case 'date':
      return values.date;
    case 'time':
      return values.time;
    case 'file':
      return values.file;
    case 'path':
      return values.path;
    case 'sheet':
      return values.sheet;
    default:
      return run.text || '';
  }
}

/**
 * The lines of one section for one page: the fields replaced by their
 * values, the text cut at its line breaks. Neighbouring pieces of one format
 * are joined; a line without text is an empty list.
 */
export function headerFooterLines(runs: readonly HeaderFooterRun[], values: HeaderFooterValues): HeaderFooterPiece[][] {
  const lines: HeaderFooterPiece[][] = [[]];
  const add = (text: string, format: HeaderFooterFormat) => {
    if (text.length === 0) return;
    const line = lines[lines.length - 1];
    const last = line.length > 0 ? line[line.length - 1] : undefined;
    if (last && last.format === format) last.text += text;
    else line.push({ text, format });
  };
  for (let i = 0; i < runs.length; i++) {
    const run = runs[i];
    const text = fieldText(run, values);
    // A field's value is one piece of a line, whatever it holds.
    const parts = run.field ? [text.replace(/[\r\n]+/g, ' ')] : text.split(/\r\n|\r|\n/);
    for (let p = 0; p < parts.length; p++) {
      if (p > 0) {
        if (lines.length >= MAX_LINES) return lines;
        lines.push([]);
      }
      add(parts[p], run.format);
    }
  }
  // A section that ends with a line break has no line after it.
  while (lines.length > 1 && lines[lines.length - 1].length === 0) lines.pop();
  if (lines.length === 1 && lines[0].length === 0) return [];
  return lines;
}
