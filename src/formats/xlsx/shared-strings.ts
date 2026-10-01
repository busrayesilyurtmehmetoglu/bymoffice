// Shared string table (xl/sharedStrings.xml) and the string item syntax it
// shares with inline strings (<si> and <is> have the same content model).

import type { Color, Font, RichRun, SharedString, Underline } from '../../model/sheet/types';
import { XmlScanner } from '../../xml-scan';

/**
 * Decodes the "_xHHHH_" escapes SpreadsheetML uses for characters XML cannot
 * carry ("_x000D_" is a carriage return, "_x005F_" a literal underscore that
 * would otherwise start an escape).
 */
export function decodeXstring(text: string): string {
  if (text.indexOf('_x') < 0) return text;
  return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_match, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

const UNDERLINES: Record<string, Underline> = {
  none: 'none',
  single: 'single',
  double: 'double',
  singleAccounting: 'singleAccounting',
  doubleAccounting: 'doubleAccounting',
};

/** Plain non-negative integer, or undefined. */
function uint(text: string | undefined): number | undefined {
  return text !== undefined && /^[0-9]{1,9}$/.test(text) ? Number(text) : undefined;
}

function finite(text: string | undefined): number | undefined {
  // xsd:double as written in files; Number() alone would also take "0x10".
  if (text === undefined || !/^\s*[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?\s*$/.test(text)) return undefined;
  const value = Number(text);
  return isFinite(value) ? value : undefined;
}

/** A boolean property element such as <b/>: present without `val` means true. */
function flag(val: string | undefined): boolean {
  return !(val === '0' || val === 'false');
}

/**
 * With a colour element (<color>, <tabColor>, <fgColor>, ...) current: reads
 * its attributes and consumes it. Undefined when it states nothing usable.
 */
export function parseColor(scanner: XmlScanner): Color | undefined {
  const color: Color = {};
  let any = false;
  const rgb = scanner.attr('rgb');
  if (rgb !== undefined && /^([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})$/.test(rgb)) {
    color.rgb = rgb.toUpperCase();
    any = true;
  }
  const theme = uint(scanner.attr('theme'));
  if (theme !== undefined) {
    color.theme = theme;
    any = true;
  }
  const tint = finite(scanner.attr('tint'));
  if (tint !== undefined) {
    color.tint = Math.max(-1, Math.min(1, tint));
    any = true;
  }
  const indexed = uint(scanner.attr('indexed'));
  if (indexed !== undefined) {
    color.indexed = indexed;
    any = true;
  }
  const auto = scanner.attr('auto');
  if (auto !== undefined) {
    color.auto = flag(auto);
    any = true;
  }
  scanner.skip();
  return any ? color : undefined;
}

/**
 * With an <rPr> open tag current: reads the run properties into a Font and
 * consumes the element.
 */
export function parseRunFont(scanner: XmlScanner): Font {
  const font: Font = { bold: false, italic: false, strike: false, underline: 'none' };
  if (scanner.selfClosing) return font;
  while (scanner.next()) {
    if (scanner.kind === 'close') break;
    if (scanner.kind !== 'open') continue;
    const val = scanner.attr('val');
    switch (scanner.local) {
      case 'b':
        font.bold = flag(val);
        break;
      case 'i':
        font.italic = flag(val);
        break;
      case 'strike':
        font.strike = flag(val);
        break;
      case 'u':
        font.underline = val === undefined ? 'single' : (Object.prototype.hasOwnProperty.call(UNDERLINES, val) && UNDERLINES[val]) || 'single';
        break;
      case 'sz': {
        const size = finite(val);
        if (size !== undefined && size > 0) font.size = size;
        break;
      }
      case 'color': {
        const color = parseColor(scanner);
        if (color) font.color = color;
        continue; // already consumed
      }
      // Run properties name the font with <rFont>; style sheet fonts use <name>.
      case 'rFont':
      case 'name':
        if (val !== undefined) font.name = val;
        break;
      case 'family': {
        const family = uint(val);
        if (family !== undefined) font.family = family;
        break;
      }
      case 'charset': {
        const charset = uint(val);
        if (charset !== undefined) font.charset = charset;
        break;
      }
      case 'scheme':
        if (val === 'major' || val === 'minor' || val === 'none') font.scheme = val;
        break;
      case 'vertAlign':
        if (val === 'superscript' || val === 'subscript') font.vertAlign = val;
        break;
    }
    scanner.skip();
  }
  return font;
}

/** With an <r> open tag current: reads one rich text run. */
function parseRun(scanner: XmlScanner): RichRun {
  const run: RichRun = { text: '' };
  if (scanner.selfClosing) return run;
  while (scanner.next()) {
    if (scanner.kind === 'close') break;
    if (scanner.kind !== 'open') continue;
    if (scanner.local === 't') run.text += decodeXstring(scanner.readText());
    else if (scanner.local === 'rPr') run.font = parseRunFont(scanner);
    else scanner.skip();
  }
  return run;
}

/**
 * With an <si> or <is> open tag current: reads the string item and consumes
 * the element. Phonetic runs (<rPh>) and <phoneticPr> are not part of the text.
 */
export function parseStringItem(scanner: XmlScanner): SharedString {
  let text = '';
  let runs: RichRun[] | undefined;
  if (!scanner.selfClosing) {
    while (scanner.next()) {
      if (scanner.kind === 'close') break;
      if (scanner.kind !== 'open') continue;
      if (scanner.local === 't') {
        const plain = decodeXstring(scanner.readText());
        text += plain;
        if (runs) runs.push({ text: plain });
      } else if (scanner.local === 'r') {
        // A plain <t> before the first run is not valid, but keep its text in step with the runs.
        if (!runs) runs = text ? [{ text }] : [];
        const run = parseRun(scanner);
        runs.push(run);
        text += run.text;
      } else scanner.skip();
    }
  }
  return runs ? { text, runs } : { text };
}

/** Moves to the root element and checks its local name. */
function openRoot(scanner: XmlScanner): void {
  while (scanner.next()) {
    if (scanner.kind === 'open') break;
  }
  if (scanner.local !== 'sst') {
    throw new Error(`Invalid shared strings part: root element is <${scanner.name}>, expected <sst>`);
  }
}

export function parseSharedStrings(bytes: Uint8Array): SharedString[] {
  const scanner = new XmlScanner(bytes);
  const strings: SharedString[] = [];
  openRoot(scanner);
  if (!scanner.selfClosing) {
    while (scanner.next()) {
      if (scanner.kind === 'close') break;
      if (scanner.kind !== 'open') continue;
      if (scanner.local === 'si') strings.push(parseStringItem(scanner));
      else scanner.skip();
    }
  }
  // Anything after the root is still checked for well-formedness.
  while (scanner.next());
  return strings;
}
