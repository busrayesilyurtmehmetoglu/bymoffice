// Differential test against SheetJS SSF, an independent implementation of the
// same format codes (a dev dependency, never shipped). The grid below avoids
// the inputs where SSF itself departs from Excel: decimal half-way cases it
// rounds in binary (1.005), times within a second of midnight, negative dates
// (it returns ""), fraction and exponent layouts it pads differently, and
// negative numbers that round to zero, where it drops the sign in some codes
// and keeps it in others (Excel keeps it: -0.004 as "0.00" shows "-0.00").
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import { test } from 'node:test';
import { formatValue } from '../src/format/number-format';

const SSF = createRequire(import.meta.url)('ssf') as { format(code: string, value: number): string };

const NUMBER_CODES = [
  'General', '0', '0.00', '#,##0', '#,##0.00', '0%', '0.00%', '0.00E+00', '0.0', '0.000', '#,###', '000', '00000',
  '#,##0_);(#,##0)', '#,##0_);[Red](#,##0)', '#,##0.00_);(#,##0.00)', '#,##0.00_);[Red](#,##0.00)',
  '$#,##0_);($#,##0)', '$#,##0.00_);[Red]($#,##0.00)',
  '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)', '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  '0.00;-0.00;"zero"', '0;;', '0.0;[Red]-0.0', '0 "units"', '0.00 "TL"', '#,##0.00\\ "₺"', '[Red]0.00', '[Blue]General',
  '#,##0,', '#,##0.0,"K"',
];
const NUMBERS = [
  0, 1, -1, 0.5, 1.5, 12, 12.34, -12.34, 123.456, 1234.5678, -1234.5678, 1234567.891, 0.001, 0.0001234, 1e-7, 1e10, 1.23e15,
  2.25, 3.75, 0.046, 99.996, 999.9996, 0.3, 0.333333, 100, 1000, 1e6, 45000, 45000.5, 123456789012,
];

const DATE_CODES = [
  'yyyy-mm-dd', 'dd/mm/yyyy', 'd-mmm-yy', 'd-mmm', 'mmm-yy', 'mmmm d, yyyy', 'dddd, mmmm dd, yyyy', 'ddd', 'mmmmm', 'yy',
  'h:mm', 'h:mm:ss', 'hh:mm:ss', 'h:mm AM/PM', 'h:mm:ss AM/PM', 'mm:ss', '[h]:mm:ss', '[mm]:ss', '[s]', 'mmss.0',
  'h:mm:ss.000', 'm/d/yyyy h:mm', 'yyyy-mm-dd hh:mm:ss', 'd/m/yy', 'm', 'mm', 'mmm', 'd', 'dd', 'h', 'hh', 's', 'ss',
  '[$-409]d-mmm-yyyy;@', '[$-409]mmmm d, yyyy',
];
const SERIALS = [0, 1, 0.5, 1.5, 12, 59, 60, 61, 0.75, 0.25, 0.123456, 36526, 45000, 45000.5, 45292.573472, 43831.000011574, 2958465, 100.99];

test('number formats: numeric codes agree with an independent implementation', () => {
  let compared = 0;
  for (const code of NUMBER_CODES) {
    for (const value of NUMBERS) {
      assert.equal(formatValue(value, code).text, SSF.format(code, value), `${JSON.stringify(code)} with ${value}`);
      compared++;
    }
  }
  assert.equal(compared, NUMBER_CODES.length * NUMBERS.length);
});

test('number formats: date and time codes agree with an independent implementation', () => {
  for (const code of DATE_CODES) {
    for (const value of SERIALS) {
      assert.equal(formatValue(value, code).text, SSF.format(code, value), `${JSON.stringify(code)} with ${value}`);
    }
  }
});
