// Colour, number format and cell style resolution for the spreadsheet model.
//
// The model keeps colours and formats as the file states them (theme index,
// tint, palette index, table indexes); the functions here follow those
// references so that drawing code gets plain values.

import { emptyStyleSheet } from './types';
import type { Alignment, Border, Color, Fill, Font, StyleSheet } from './types';

/**
 * The default legacy palette (ECMA-376 Part 1, 18.8.27), indexes 0..63 as
 * "RRGGBB". Indexes 64 (system foreground) and 65 (system background) are not
 * in the list; they resolve as automatic.
 */
export const DEFAULT_INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

/** Built-in number formats (ECMA-376 Part 1, 18.8.30), en-US codes. Ids not listed read as "General". */
const BUILTIN_NUM_FMTS: Readonly<Record<number, string>> = {
  0: 'General',
  1: '0',
  2: '0.00',
  3: '#,##0',
  4: '#,##0.00',
  5: '$#,##0_);($#,##0)',
  6: '$#,##0_);[Red]($#,##0)',
  7: '$#,##0.00_);($#,##0.00)',
  8: '$#,##0.00_);[Red]($#,##0.00)',
  9: '0%',
  10: '0.00%',
  11: '0.00E+00',
  12: '# ?/?',
  13: '# ??/??',
  14: 'mm-dd-yy',
  15: 'd-mmm-yy',
  16: 'd-mmm',
  17: 'mmm-yy',
  18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM',
  20: 'h:mm',
  21: 'h:mm:ss',
  22: 'm/d/yy h:mm',
  // 37-40 as Excel applies them; the standard prints them without the "_)" padding.
  37: '#,##0_);(#,##0)',
  38: '#,##0_);[Red](#,##0)',
  39: '#,##0.00_);(#,##0.00)',
  40: '#,##0.00_);[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  45: 'mm:ss',
  46: '[h]:mm:ss',
  47: 'mmss.0',
  48: '##0.0E+0',
  49: '@',
};

const HEX6 = /^[0-9A-Fa-f]{6}$/;
const HEX8 = /^[0-9A-Fa-f]{8}$/;

/** "RRGGBB" or "AARRGGBB" -> upper case "RRGGBB"; undefined for anything else. */
function toRgb6(value: string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (HEX6.test(value)) return value.toUpperCase();
  if (HEX8.test(value)) return value.slice(2).toUpperCase();
  return undefined;
}

function hueToChannel(p: number, q: number, t: number): number {
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

const hex2 = (channel: number) => {
  const n = Math.max(0, Math.min(255, Math.round(channel * 255)));
  return (n < 16 ? '0' : '') + n.toString(16).toUpperCase();
};

/**
 * Applies a tint to "RRGGBB" (ECMA-376 Part 1, 18.8.19): the colour goes to
 * HLS, luminance is scaled towards black (tint < 0) or white (tint > 0), and
 * the result goes back to RGB. Luminance is kept as a 0..1 fraction, which is
 * the same formula as the HLSMAX-scaled one in the standard.
 */
function applyTint(rgb: string, tint: number): string {
  const r = parseInt(rgb.slice(0, 2), 16) / 255;
  const g = parseInt(rgb.slice(2, 4), 16) / 255;
  const b = parseInt(rgb.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }

  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;

  if (s === 0) return hex2(l) + hex2(l) + hex2(l);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return hex2(hueToChannel(p, q, h + 1 / 3)) + hex2(hueToChannel(p, q, h)) + hex2(hueToChannel(p, q, h - 1 / 3));
}

/** Resolves a colour to "#RRGGBB"; undefined when it is automatic or cannot be resolved. */
export function resolveColor(color: Color | undefined, styles: StyleSheet): string | undefined {
  if (!color) return undefined;
  let base: string | undefined;
  if (color.rgb !== undefined) base = toRgb6(color.rgb);
  else if (color.theme !== undefined) base = toRgb6(styles.themeColors[color.theme]);
  else if (color.indexed !== undefined) {
    // 64 and 65 are the system foreground/background: automatic. They are
    // never part of a palette, custom or not.
    const palette = styles.indexedColors && styles.indexedColors.length ? styles.indexedColors : DEFAULT_INDEXED_COLORS;
    if (color.indexed < 64) base = toRgb6(palette[color.indexed] ?? DEFAULT_INDEXED_COLORS[color.indexed]);
  }
  if (base === undefined) return undefined;

  const tint = color.tint;
  if (tint !== undefined && tint !== 0 && Number.isFinite(tint)) base = applyTint(base, Math.max(-1, Math.min(1, tint)));
  return '#' + base;
}

/** Built-in number format code for an id, or the custom one from the style sheet; "General" when unknown. */
export function formatCodeOf(styles: StyleSheet, numFmtId: number): string {
  return styles.numFmts.get(numFmtId) ?? BUILTIN_NUM_FMTS[numFmtId] ?? 'General';
}

/** Everything needed to draw a cell, with indexes followed. */
export interface ResolvedCellStyle {
  font: Font;
  fill: Fill;
  border: Border;
  alignment: Alignment;
  numFmt: string;
  /** Id of the number format; ids below 164 that the file does not redefine are built-in formats. */
  numFmtId: number;
}

/**
 * Follows the indexes of cellXfs[index]. An index outside the table reads as
 * the first format, and a table that lacks the entry falls back to its first
 * entry and then to the defaults, so this never throws on a damaged style sheet.
 */
let fallbackStyles: StyleSheet | undefined;

export function resolveCellStyle(styles: StyleSheet, index: number): ResolvedCellStyle {
  // Only reached for a style sheet with empty tables; built once, never handed out for writing.
  const defaults = fallbackStyles || (fallbackStyles = emptyStyleSheet());
  const xf = styles.cellXfs[index] ?? styles.cellXfs[0] ?? defaults.cellXfs[0];
  return {
    font: styles.fonts[xf.fontId] ?? styles.fonts[0] ?? defaults.fonts[0],
    fill: styles.fills[xf.fillId] ?? styles.fills[0] ?? defaults.fills[0],
    border: styles.borders[xf.borderId] ?? styles.borders[0] ?? defaults.borders[0],
    alignment: xf.alignment ?? {},
    numFmt: formatCodeOf(styles, xf.numFmtId),
    numFmtId: xf.numFmtId,
  };
}
