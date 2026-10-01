// Text shown for a cell. Number formats (dates, currency, percent, custom
// codes) are not applied yet; every number is shown the way "General" shows it.

import type { Cell } from '../model/sheet/types';

const trimZeros = (text: string) => (text.indexOf('.') < 0 ? text : text.replace(/\.?0+$/, ''));

function scientific(value: number): string {
  const [mantissa, exponent] = value.toExponential(5).split('e');
  const power = Number(exponent);
  return `${trimZeros(mantissa)}E${power < 0 ? '-' : '+'}${String(Math.abs(power)).padStart(2, '0')}`;
}

/**
 * "General" for a number, as Excel shows it in a wide enough cell: at most 11
 * characters (12 with a minus sign), scientific notation when the digits do
 * not fit.
 */
export function formatGeneral(value: number): string {
  if (value === 0) return '0';
  if (!isFinite(value)) return '#NUM!';
  const abs = Math.abs(value);
  if (abs < 1e11 && Math.floor(value) === value) return String(value);
  const magnitude = Math.floor(Math.log10(abs));
  if (magnitude >= -4 && magnitude <= -1) return trimZeros(value.toPrecision(10 + magnitude));
  if (magnitude === 10) {
    // Eleven integer digits leave no room for decimals; rounding may add a twelfth digit.
    const rounded = Math.round(value);
    return Math.abs(rounded) < 1e11 ? String(rounded) : scientific(value);
  }
  if (magnitude >= -9 && magnitude <= 9) {
    const width = value < 0 ? 12 : 11;
    const fixed = trimZeros(value.toFixed(12));
    if (fixed.length <= width) return fixed;
    const precise = value.toPrecision(10);
    if (precise.indexOf('e') < 0 && trimZeros(precise).length <= width) return trimZeros(precise);
  }
  return scientific(value);
}

export function cellText(cell: Cell | undefined): string {
  if (!cell || cell.value === null) return '';
  switch (cell.kind) {
    case 'number':
      return formatGeneral(cell.value as number);
    case 'boolean':
      return cell.value ? 'TRUE' : 'FALSE';
    default:
      return String(cell.value);
  }
}
