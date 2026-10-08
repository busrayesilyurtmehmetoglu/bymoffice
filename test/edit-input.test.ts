import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cellEditText, parseCellInput, type InputOptions, type ParsedInput } from '../src/edit/input';
import { LOCALE_EN, LOCALE_TR, normalizeLocale, type FormatLocale } from '../src/format/locale';
import type { Cell } from '../src/model/sheet/types';

/** A fixed "today" so that dates typed without a year do not depend on the clock. */
const TODAY = new Date(2024, 5, 10);
const EN: InputOptions = { locale: LOCALE_EN, today: TODAY };
const TR: InputOptions = { locale: LOCALE_TR, today: TODAY };

/** Serial of a date on or after 1 March 1900 in the 1900 system. */
function serial(year: number, month: number, day: number): number {
  return Date.UTC(year, month - 1, day) / 86400000 + 25569;
}

function time(hour: number, minute: number, second = 0): number {
  return (hour * 3600 + minute * 60 + second) / 86400;
}

const number = (value: number, numFmtId?: number, numFmt?: string): ParsedInput => {
  const result: ParsedInput = { kind: 'number', value };
  if (numFmtId !== undefined) result.numFmtId = numFmtId;
  if (numFmt !== undefined) result.numFmt = numFmt;
  return result;
};
const text = (value: string): ParsedInput => ({ kind: 'string', value });

type Row = [string, InputOptions, ParsedInput];

function check(rows: Row[]): void {
  for (const [input, options, expected] of rows) {
    const locale = options.locale === LOCALE_TR ? 'tr' : 'en';
    assert.deepEqual(parseCellInput(input, options), expected, `${JSON.stringify(input)} (${locale})`);
  }
}

test('input: blank, text, apostrophe, formula', () => {
  check([
    ['', EN, { kind: 'blank', value: null }],
    [' ', EN, text(' ')],
    ['   ', TR, text('   ')],
    ['hello', EN, text('hello')],
    ['  padded  ', EN, text('  padded  ')],
    ['İstanbul', TR, text('İstanbul')],
    ["'123", EN, text('123')],
    ["'", EN, text('')],
    ["''quoted'", EN, text("'quoted'")],
    ["'=A1", EN, text('=A1')],
    ["'TRUE", EN, text('TRUE')],
    ["' 15.03.2024", TR, text(' 15.03.2024')],
    ['=A1+B1', EN, { kind: 'blank', value: null, formula: 'A1+B1' }],
    ['= SUM(A1:A3) ', EN, { kind: 'blank', value: null, formula: ' SUM(A1:A3) ' }],
    ['==', EN, { kind: 'blank', value: null, formula: '=' }],
    ['=', EN, text('=')],
    // Formula-like input starting with a sign is kept as text (Excel would make it a formula).
    ['+A1', EN, text('+A1')],
    ['-A1', EN, text('-A1')],
    // Only a leading "=" makes a formula.
    [' =A1', EN, text(' =A1')],
  ]);
});

test('input: numbers in the locale notation', () => {
  check([
    ['0', EN, number(0)],
    ['123', EN, number(123)],
    ['  42  ', EN, number(42)],
    ['-12.5', EN, number(-12.5)],
    ['+7', EN, number(7)],
    ['.5', EN, number(0.5)],
    ['5.', EN, number(5)],
    ['-.5', EN, number(-0.5)],
    ['(12.5)', EN, number(-12.5)],
    ['( 3 )', EN, number(-3)],
    ['-0', EN, number(0)],
    ['1,234.56', EN, number(1234.56, 4)],
    ['1,234', EN, number(1234, 3)],
    ['1,234,567', EN, number(1234567, 3)],
    ['-1,234.5', EN, number(-1234.5, 4)],
    ['1.5e3', EN, number(1500)],
    ['1.5E-3', EN, number(0.0015)],
    ['1e+20', EN, number(1e20)],
    ['0.1', EN, number(0.1)],
    // Turkish: comma decimal, dot thousands.
    ['1234,56', TR, number(1234.56)],
    ['1.234,56', TR, number(1234.56, 4)],
    ['1.234', TR, number(1234, 3)],
    ['1.234.567,8', TR, number(1234567.8, 4)],
    [',5', TR, number(0.5)],
    ['-12,5', TR, number(-12.5)],
    ['(12,5)', TR, number(-12.5)],
    ['1,5E3', TR, number(1500)],
    ['1,5e-3', TR, number(0.0015)],
    // Not numbers.
    ['1,23,4', EN, text('1,23,4')],
    ['12,34', EN, text('12,34')],
    ['1,2345', EN, text('1,2345')],
    [',123', EN, text(',123')],
    ['1.234.56', TR, text('1.234.56')],
    ['1.23.4', TR, text('1.23.4')],
    ['1.2.3.4', EN, text('1.2.3.4')],
    ['1,234.56', TR, text('1,234.56')],
    ['--5', EN, text('--5')],
    ['(-5)', EN, text('(-5)')],
    ['(5', EN, text('(5')],
    ['- 5', EN, text('- 5')],
    ['5 5', EN, text('5 5')],
    ['1e', EN, text('1e')],
    ['e5', EN, text('e5')],
    ['.', EN, text('.')],
    ['-', EN, text('-')],
    ['1e999', EN, text('1e999')],
    ['0x10', EN, text('0x10')],
    ['١٢٣', EN, text('١٢٣')],
  ]);
});

test('input: percentages', () => {
  check([
    ['12%', EN, number(0.12, 9)],
    ['12.5%', EN, number(0.125, 10)],
    ['12,5 %', TR, number(0.125, 10)],
    ['7%', TR, number(0.07, 9)],
    ['-5%', EN, number(-0.05, 9)],
    ['100%', EN, number(1, 9)],
    ['0.1%', EN, number(0.001, 10)],
    // Turkish style: the sign in front.
    ['%12', TR, number(0.12, 9)],
    ['% 12,5', TR, number(0.125, 10)],
    ['%', EN, text('%')],
    ['12%%', EN, text('12%%')],
    ['a%', EN, text('a%')],
  ]);
});

test('input: currency', () => {
  check([
    ['$1,234.50', EN, number(1234.5, undefined, '"$"#,##0.00')],
    ['$5', EN, number(5, undefined, '"$"#,##0')],
    ['-$5', EN, number(-5, undefined, '"$"#,##0')],
    ['$-5', EN, number(-5, undefined, '"$"#,##0')],
    ['($5.25)', EN, number(-5.25, undefined, '"$"#,##0.00')],
    ['€ 10', EN, number(10, undefined, '"€"#,##0')],
    ['£3.5', EN, number(3.5, undefined, '"£"#,##0.00')],
    ['10 €', EN, number(10, undefined, '#,##0 "€"')],
    ['₺12,50', TR, number(12.5, undefined, '"₺"#,##0.00')],
    ['12,50 ₺', TR, number(12.5, undefined, '#,##0.00 "₺"')],
    ['12,50₺', TR, number(12.5, undefined, '#,##0.00 "₺"')],
    ['1.234,50 TL', TR, number(1234.5, undefined, '#,##0.00 "TL"')],
    ['1.234,50 tl', TR, number(1234.5, undefined, '#,##0.00 "TL"')],
    ['-1.234 TL', TR, number(-1234, undefined, '#,##0 "TL"')],
    ['$', EN, text('$')],
    ['TL', TR, text('TL')],
    ['$abc', EN, text('$abc')],
    ['$$5', EN, text('$$5')],
    ['$5%', EN, text('$5%')],
    ['5 USD', EN, text('5 USD')],
    ['$1e5', EN, text('$1e5')],
  ]);
});

test('input: dates', () => {
  const day = serial(2024, 3, 15);
  check([
    ['15.03.2024', TR, number(day, 14)],
    ['15/3/24', TR, number(day, 14)],
    ['15-03-2024', TR, number(day, 14)],
    ['15.3.2024', TR, number(day, 14)],
    ['3/15/2024', EN, number(day, 14)],
    ['03/15/24', EN, number(day, 14)],
    ['3-15-2024', EN, number(day, 14)],
    // ISO order works in every locale.
    ['2024-03-15', EN, number(day, 14)],
    ['2024-03-15', TR, number(day, 14)],
    ['2024/3/15', TR, number(day, 14)],
    ['2024.03.15', TR, number(day, 14)],
    // The same digits, different locales.
    ['3/4/2024', EN, number(serial(2024, 3, 4), 14)],
    ['3/4/2024', TR, number(serial(2024, 4, 3), 14)],
    // Month names, in the locale's language or in English.
    ['15 Mart 2024', TR, number(day, 15)],
    ['15 mart 2024', TR, number(day, 15)],
    ['15 MART 2024', TR, number(day, 15)],
    ['15 Mar 2024', TR, number(day, 15)],
    ['15-Mar-24', EN, number(day, 15)],
    ['15 March 2024', EN, number(day, 15)],
    ['March 15, 2024', EN, number(day, 15)],
    ['15 March 2024', TR, number(day, 15)],
    ['1 Ağustos 2024', TR, number(serial(2024, 8, 1), 15)],
    ['1 AĞUSTOS 2024', TR, number(serial(2024, 8, 1), 15)],
    ['5 Mayıs 2024', TR, number(serial(2024, 5, 5), 15)],
    ['5 MAYIS 2024', TR, number(serial(2024, 5, 5), 15)],
    ['5 mayis 2024', TR, number(serial(2024, 5, 5), 15)],
    ['7 Şub 2024', TR, number(serial(2024, 2, 7), 15)],
    ['3 Aralık 2024', TR, number(serial(2024, 12, 3), 15)],
    // Without a year: this year.
    ['15.03', TR, number(day, 16)],
    ['15/3', TR, number(day, 16)],
    ['3/15', EN, number(day, 16)],
    ['15 Mar', TR, number(day, 16)],
    ['15 Mart', TR, number(day, 16)],
    ['Mar 15', EN, number(day, 16)],
    // A month of a year: its first day.
    ['Mart 2024', TR, number(serial(2024, 3, 1), 17)],
    ['Mar 2024', EN, number(serial(2024, 3, 1), 17)],
    ['3/2024', EN, number(serial(2024, 3, 1), 17)],
    ['2024-03', TR, number(serial(2024, 3, 1), 17)],
    // Turkish "1.5" is no number (a dot is the thousands separator): it is 1 May.
    ['1.5', TR, number(serial(2024, 5, 1), 16)],
    ['1.5', EN, number(1.5)],
    ['12.5', TR, number(serial(2024, 5, 12), 16)],
    ['1.234', TR, number(1234, 3)],
    // Two-digit years: 00-29 are 20xx, 30-99 are 19xx.
    ['01.01.29', TR, number(serial(2029, 1, 1), 14)],
    ['01.01.30', TR, number(serial(1930, 1, 1), 14)],
    ['01.01.00', TR, number(serial(2000, 1, 1), 14)],
    ['31.12.99', TR, number(serial(1999, 12, 31), 14)],
    ['1/1/5', EN, number(serial(2005, 1, 1), 14)],
    // Leap years.
    ['29.02.2024', TR, number(serial(2024, 2, 29), 14)],
    ['29.02.2000', TR, number(serial(2000, 2, 29), 14)],
    ['29.02.2023', TR, text('29.02.2023')],
    ['29.02.2100', TR, text('29.02.2100')],
    // Excel's 29 February 1900 is accepted, as serial 60.
    ['29.02.1900', TR, number(60, 14)],
    ['28.02.1900', TR, number(59, 14)],
    ['01.03.1900', TR, number(61, 14)],
    ['01.01.1900', TR, number(1, 14)],
    ['31.12.9999', TR, number(2958465, 14)],
    // Not dates.
    ['31.02.2024', TR, text('31.02.2024')],
    ['31.04.2024', TR, text('31.04.2024')],
    ['00.01.2024', TR, text('00.01.2024')],
    ['15.13.2024', TR, text('15.13.2024')],
    ['15/03/2024', EN, text('15/03/2024')],
    ['31.12.1899', TR, text('31.12.1899')],
    ['01.01.10000', TR, text('01.01.10000')],
    ['01.01.202', TR, text('01.01.202')],
    ['15.03/2024', TR, text('15.03/2024')],
    ['15..03', TR, text('15..03')],
    ['15 Foo 2024', TR, text('15 Foo 2024')],
    ['Mart', TR, text('Mart')],
    ['Mart Mart', TR, text('Mart Mart')],
    ['32 Mart', TR, text('32 Mart')],
    ['1-2-3-4', EN, text('1-2-3-4')],
  ]);
});

test('input: times and date-times', () => {
  const day = serial(2024, 3, 15);
  check([
    ['13:45', EN, number(time(13, 45), 20)],
    ['13:45', TR, number(time(13, 45), 20)],
    ['9:05', EN, number(time(9, 5), 20)],
    ['13:45:30', EN, number(time(13, 45, 30), 21)],
    ['0:00', EN, number(0, 20)],
    ['00:00:00', EN, number(0, 21)],
    ['23:59:59', EN, number(time(23, 59, 59), 21)],
    ['1:30 PM', EN, number(time(13, 30), 18)],
    ['1:30 pm', EN, number(time(13, 30), 18)],
    ['1:30PM', EN, number(time(13, 30), 18)],
    ['1:30:15 AM', EN, number(time(1, 30, 15), 19)],
    ['12:00 AM', EN, number(0, 18)],
    ['12:00 PM', EN, number(0.5, 18)],
    ['1:30 ÖS', TR, number(time(13, 30), 18)],
    ['1:30 ös', TR, number(time(13, 30), 18)],
    ['1:30 ÖÖ', TR, number(time(1, 30), 18)],
    ['1:30 PM', TR, number(time(13, 30), 18)],
    ['1:30 p', EN, number(time(13, 30), 18)],
    ['1:30p', EN, number(time(13, 30), 18)],
    ['5 a', EN, number(time(5, 0), 18)],
    ['9 am', EN, number(time(9, 0), 18)],
    ['1 PM', EN, number(time(13, 0), 18)],
    // 24 hours and more: elapsed time.
    ['24:00', EN, number(1, 46)],
    ['25:00', EN, number(25 / 24, 46)],
    ['100:30:00', EN, number(time(100, 30), 46)],
    // Fractions of a second.
    ['13:45:30.5', EN, number((13 * 3600 + 45 * 60 + 30.5) / 86400, 47)],
    ['13:45:30,5', TR, number((13 * 3600 + 45 * 60 + 30.5) / 86400, 47)],
    // Date and time.
    ['15.03.2024 13:45', TR, number(day + time(13, 45), 22)],
    ['15.03.2024 13:45:30', TR, number(day + time(13, 45, 30), 22)],
    ['3/15/2024 1:45 PM', EN, number(day + time(13, 45), 22)],
    ['2024-03-15 13:45', EN, number(day + time(13, 45), 22)],
    ['2024-03-15T13:45:30', TR, number(day + time(13, 45, 30), 22)],
    ['15 Mart 2024 13:45', TR, number(day + time(13, 45), 22)],
    ['15.03.2024 00:00', TR, number(day, 22)],
    ['15.03 9:00', TR, number(day + time(9, 0), 22)],
    // Not times.
    ['13:60', EN, text('13:60')],
    ['13:45:60', EN, text('13:45:60')],
    ['13:', EN, text('13:')],
    [':30', EN, text(':30')],
    ['1:2:3:4', EN, text('1:2:3:4')],
    ['13:30 PM', EN, text('13:30 PM')],
    ['a:30', EN, text('a:30')],
    ['-1:30', EN, text('-1:30')],
    ['31.02.2024 13:45', TR, text('31.02.2024 13:45')],
    ['15.03.2024 25:00', TR, text('15.03.2024 25:00')],
    ['15.03.2024 13:60', TR, text('15.03.2024 13:60')],
    ['note: 5', EN, text('note: 5')],
    ['31.12.9999 23:59', TR, number(2958465 + time(23, 59), 22)],
  ]);
});

test('input: 1904 date system', () => {
  const options: InputOptions = { locale: LOCALE_TR, date1904: true, today: TODAY };
  check([
    ['01.01.1904', options, number(0, 14)],
    ['02.01.1904', options, number(1, 14)],
    ['15.03.2024', options, number(serial(2024, 3, 15) - 1462, 14)],
    ['31.12.9999', options, number(2957003, 14)],
    ['29.02.1904', options, number(59, 14)],
    // Before the first day of the system, and the day that only the 1900 system has.
    ['31.12.1903', options, text('31.12.1903')],
    ['29.02.1900', options, text('29.02.1900')],
    ['01.01.1904 12:00', options, number(0.5, 22)],
  ]);
});

test('input: dates without a year use today, or the clock', () => {
  const in2031: InputOptions = { locale: LOCALE_TR, today: new Date(2031, 0, 5) };
  assert.deepEqual(parseCellInput('15.03', in2031), number(serial(2031, 3, 15), 16));
  // 29 February only exists in some years.
  assert.deepEqual(parseCellInput('29.02', TR), number(serial(2024, 2, 29), 16));
  assert.deepEqual(parseCellInput('29.02', in2031), text('29.02'));
  const now = parseCellInput('15.03', { locale: LOCALE_TR });
  assert.equal(now.kind, 'number');
  assert.equal(now.value, serial(new Date().getFullYear(), 3, 15));
  // An unusable "today" falls back to the clock.
  const broken = { locale: LOCALE_TR, today: new Date(NaN) };
  assert.equal(parseCellInput('15.03', broken).value, now.value);
  assert.equal(parseCellInput('15.03', { locale: LOCALE_TR, today: 'x' as unknown as Date }).value, now.value);
});

test('input: booleans and errors', () => {
  check([
    ['TRUE', EN, { kind: 'boolean', value: true }],
    ['true', EN, { kind: 'boolean', value: true }],
    ['False', EN, { kind: 'boolean', value: false }],
    [' TRUE ', EN, { kind: 'boolean', value: true }],
    ['TRUE', TR, { kind: 'boolean', value: true }],
    ['false', TR, { kind: 'boolean', value: false }],
    ['DOĞRU', TR, { kind: 'boolean', value: true }],
    ['doğru', TR, { kind: 'boolean', value: true }],
    ['Doğru', TR, { kind: 'boolean', value: true }],
    ['YANLIŞ', TR, { kind: 'boolean', value: false }],
    ['yanlış', TR, { kind: 'boolean', value: false }],
    ['Yanlış', TR, { kind: 'boolean', value: false }],
    ['DOĞRU', EN, text('DOĞRU')],
    ['yes', EN, text('yes')],
    ['TRUE!', EN, text('TRUE!')],
    ['#N/A', EN, { kind: 'error', value: '#N/A' }],
    ['#DIV/0!', EN, { kind: 'error', value: '#DIV/0!' }],
    ['#REF!', TR, { kind: 'error', value: '#REF!' }],
    ['#VALUE!', EN, { kind: 'error', value: '#VALUE!' }],
    ['#NAME?', EN, { kind: 'error', value: '#NAME?' }],
    ['#NUM!', EN, { kind: 'error', value: '#NUM!' }],
    ['#NULL!', EN, { kind: 'error', value: '#NULL!' }],
    ['#n/a', EN, { kind: 'error', value: '#N/A' }],
    [' #REF! ', EN, { kind: 'error', value: '#REF!' }],
    ['#FOO!', EN, text('#FOO!')],
    ['#', EN, text('#')],
    ['#1', EN, text('#1')],
  ]);
});

test('input: hostile and huge input', () => {
  const huge = 'x'.repeat(100000);
  const started = Date.now();
  assert.equal((parseCellInput(huge).value as string).length, 32767);
  assert.equal((parseCellInput("'" + huge).value as string).length, 32767);
  assert.equal(parseCellInput('=' + huge).formula, huge);
  for (const piece of ['1', '1,', '1.', ' ', '1 ', ':', '1:', '%', '$', '(', '-', 'e', '1e', '.1', '/', '1/', 'Mar ']) {
    for (const length of [70, 300, 100000]) {
      const input = piece.repeat(Math.ceil(length / piece.length));
      for (const options of [EN, TR]) {
        const result = parseCellInput(input, options);
        assert.ok(result.kind === 'string' || result.kind === 'number', input.slice(0, 20));
      }
    }
  }
  assert.ok(Date.now() - started < 20000, `${Date.now() - started} ms`);

  // Wrong types and odd options never throw.
  assert.deepEqual(parseCellInput(undefined as unknown as string), { kind: 'blank', value: null });
  assert.deepEqual(parseCellInput(null as unknown as string), { kind: 'blank', value: null });
  assert.deepEqual(parseCellInput(5 as unknown as string), { kind: 'blank', value: null });
  assert.deepEqual(parseCellInput('1.5', null as unknown as InputOptions), number(1.5));
  assert.deepEqual(parseCellInput('1.5', { locale: {} as FormatLocale }), number(1.5));
  assert.deepEqual(parseCellInput('1.5', { locale: 'tr' as unknown as FormatLocale }), number(1.5));
  for (const input of ['\u0000', '\ud83d', '1\u0000', '\n1\n', '\t5', '١', '１２', '1 ', 'İ', 'ß PM', '1:30 İ']) {
    for (const options of [EN, TR]) assert.equal(typeof parseCellInput(input, options).kind, 'string');
  }
  assert.deepEqual(parseCellInput('\t5'), number(5));
  assert.deepEqual(parseCellInput('１２'), text('１２'));

  // A locale whose separators collide or are not single characters still parses plain digits.
  const odd = normalizeLocale({ decimal: '.', group: '.', am: '', pm: '' });
  assert.deepEqual(parseCellInput('1.5', { locale: odd }), number(1.5));
  assert.deepEqual(parseCellInput('12', { locale: normalizeLocale({ decimal: '', group: '' }) }), number(12));
  assert.deepEqual(parseCellInput('1:30', { locale: odd }), number(time(1, 30), 20));
});

test('input: custom date order', () => {
  const ymd: InputOptions = { locale: { ...LOCALE_EN, dateOrder: 'ymd', dateSeparator: '-' }, today: TODAY };
  assert.deepEqual(parseCellInput('24-03-15', ymd), number(serial(2024, 3, 15), 14));
  assert.deepEqual(parseCellInput('3-15', ymd), number(serial(2024, 3, 15), 16));
  assert.equal(cellEditText({ kind: 'number', value: serial(2024, 3, 15), style: 0 }, 'yyyy', ymd), '2024-03-15');
  assert.equal(normalizeLocale({ dateOrder: 'xyz' as 'dmy' }).dateOrder, 'mdy');
  assert.equal(normalizeLocale({ dateOrder: 'dmy' }).dateOrder, 'dmy');
  assert.equal(LOCALE_TR.dateOrder, 'dmy');
  assert.equal(LOCALE_EN.dateOrder, 'mdy');
});

const cell = (kind: Cell['kind'], value: Cell['value']): Cell => ({ kind, value, style: 0 });

test('edit text: by cell kind', () => {
  const day = serial(2024, 3, 15);
  const rows: [Cell | undefined, string, InputOptions, string][] = [
    [undefined, 'General', EN, ''],
    [cell('blank', null), 'General', EN, ''],
    [cell('number', 1234.5), 'General', EN, '1234.5'],
    [cell('number', 1234.5), 'General', TR, '1234,5'],
    [cell('number', 1234567.891), '#,##0.00', TR, '1234567,891'],
    [cell('number', 1234567.891), '#,##0.00 "₺"', TR, '1234567,891'],
    [cell('number', -0.5), '0', EN, '-0.5'],
    [cell('number', 0), '0.00', EN, '0'],
    [cell('number', 0.1 + 0.2), 'General', EN, '0.30000000000000004'],
    [cell('number', 0.1 + 0.2), '0%', TR, '30,000000000000004%'],
    [cell('number', 1234567890.1234567), 'General', TR, '1234567890,1234567'],
    [cell('number', 123456789012345), 'General', EN, '123456789012345'],
    [cell('number', 1e15), 'General', EN, '1E+15'],
    [cell('number', 1e20), 'General', EN, '1E+20'],
    [cell('number', 1.5e20), 'General', TR, '1,5E+20'],
    [cell('number', 0.000000001), 'General', EN, '0.000000001'],
    [cell('number', 1.25e-10), 'General', TR, '1,25E-10'],
    [cell('number', -1e-300), 'General', EN, '-1E-300'],
    [cell('number', 1 / 3), 'General', EN, '0.3333333333333333'],
    [cell('number', 12345678901234567890), 'General', EN, '1.2345678901234567E+19'],
    [cell('number', NaN), 'General', EN, '#NUM!'],
    // Percent formats.
    [cell('number', 0.125), '0%', TR, '12,5%'],
    [cell('number', 0.125), '0.00%', EN, '12.5%'],
    [cell('number', 0.07), '0%', EN, '7%'],
    [cell('number', 1), '0%', EN, '100%'],
    [cell('number', 0), '0%', EN, '0%'],
    [cell('number', 0.125), '0 "%"', EN, '0.125'],
    [cell('number', 0.125), '0\\%', EN, '0.125'],
    [cell('number', 0.125), '[Red]0%', EN, '12.5%'],
    // Date formats.
    [cell('number', day), 'yyyy-mm-dd', EN, '3/15/2024'],
    [cell('number', day), 'd-mmm-yy', TR, '15.03.2024'],
    [cell('number', day + time(13, 45)), 'yyyy-mm-dd', TR, '15.03.2024 13:45:00'],
    [cell('number', day + time(13, 45, 30)), 'h:mm', EN, '3/15/2024 13:45:30'],
    [cell('number', time(13, 45)), 'h:mm', EN, '13:45:00'],
    [cell('number', time(9, 5, 7)), 'h:mm:ss', TR, '9:05:07'],
    [cell('number', 0), 'h:mm', EN, '0:00:00'],
    [cell('number', 60), 'dd/mm/yyyy', TR, '29.02.1900'],
    [cell('number', 1), 'dd/mm/yyyy', TR, '01.01.1900'],
    [cell('number', 1.5), '[h]:mm', EN, '1/1/1900 12:00:00'],
    // A serial that is no date is shown as the number it is.
    [cell('number', -1), 'yyyy-mm-dd', EN, '-1'],
    [cell('number', 3000000), 'yyyy-mm-dd', EN, '3000000'],
    // Other kinds.
    [cell('boolean', true), 'General', EN, 'TRUE'],
    [cell('boolean', false), 'General', TR, 'YANLIŞ'],
    [cell('error', '#DIV/0!'), 'General', TR, '#DIV/0!'],
    [cell('string', 'hello'), 'General', EN, 'hello'],
    [cell('string', ' padded '), 'General', EN, ' padded '],
    [cell('date', '2024-03-15'), 'General', TR, '15.03.2024'],
    [cell('date', '2024-03-15T13:45:00'), 'General', EN, '3/15/2024 13:45:00'],
    [cell('date', 'not a date'), 'General', EN, 'not a date'],
    // Text that would be read back as something else gets the apostrophe.
    [cell('string', '123'), 'General', EN, "'123"],
    [cell('string', '=x'), 'General', EN, "'=x"],
    [cell('string', 'TRUE'), 'General', EN, "'TRUE"],
    [cell('string', 'doğru'), 'General', TR, "'doğru"],
    [cell('string', 'doğru'), 'General', EN, 'doğru'],
    [cell('string', '15.03.2024'), 'General', TR, "'15.03.2024"],
    [cell('string', '15.03.2024'), 'General', EN, '15.03.2024'],
    [cell('string', '1.5'), 'General', TR, "'1.5"],
    [cell('string', '12%'), 'General', EN, "'12%"],
    [cell('string', '#N/A'), 'General', EN, "'#N/A"],
    [cell('string', "'quoted"), 'General', EN, "''quoted"],
    [cell('string', ''), 'General', EN, "'"],
    [cell('string', '='), 'General', EN, '='],
  ];
  for (const [input, numFmt, options, expected] of rows) {
    assert.equal(cellEditText(input, numFmt, options), expected, `${JSON.stringify(input)} ${numFmt}`);
  }
  // 1904 system.
  const options: InputOptions = { locale: LOCALE_TR, date1904: true };
  assert.equal(cellEditText(cell('number', 0), 'yyyy', options), '01.01.1904');
  assert.equal(cellEditText(cell('number', 0.5), 'yyyy', options), '01.01.1904 12:00:00');
});

test('edit text: formulas', () => {
  const formula: Cell = { kind: 'number', value: 3, style: 0, formula: { text: 'A1+B1', type: 'normal' } };
  assert.equal(cellEditText(formula, 'General', EN), '=A1+B1');
  assert.deepEqual(parseCellInput(cellEditText(formula, 'General', EN)), { kind: 'blank', value: null, formula: 'A1+B1' });
  // A shared formula that only refers to its master has no text of its own: the value is shown.
  const shared: Cell = { kind: 'number', value: 3, style: 0, formula: { text: '', type: 'shared', si: 0 } };
  assert.equal(cellEditText(shared, 'General', EN), '3');
  // Malformed arguments.
  assert.equal(cellEditText(cell('number', 1.5), undefined as unknown as string, EN), '1.5');
  assert.equal(cellEditText(cell('number', 1.5), '[', EN), '1.5');
  assert.equal(cellEditText(null as unknown as Cell, 'General'), '');
  assert.equal(cellEditText(cell('number', 1.5), 'General', { locale: {} as FormatLocale }), '1.5');
});

/** Deterministic pseudo-random numbers (linear congruential), so failures can be reproduced. */
function generator(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state / 4294967296;
  };
}

test('round trip: parseCellInput(cellEditText(cell)) gives the cell back', () => {
  const random = generator(20240315);
  const integer = (limit: number) => Math.floor(random() * limit);
  const texts = [
    'hello', '123', '12345678901234567890', '+905551234567890123', '12A', '5a', '9am', '1.2.3', '01.03.2024', '1,5', '1.5', '1.234,56', '1,234.56', '=A1', '=', "'x", "''", 'TRUE', 'false', 'doğru',
    'YANLIŞ', '15.03.2024', '3/15/2024', '2024-03-15', '13:45', '1:30 PM', '12%', '%5', '$5', '5 TL', '#N/A',
    '#REF!', ' 42 ', '', ' ', '1e5', '(5)', '-', '+A1', 'Mart 2024', '15 Mar', '1 p', 'İI ıi', '0', '-0',
  ];
  for (const options of [EN, TR, { ...EN, date1904: true }, { ...TR, date1904: true }]) {
    const date1904 = options.date1904 === true;
    const lastDay = date1904 ? 2957003 : 2958465;
    const roundTrip = (input: Cell, numFmt: string): ParsedInput => {
      const shown = cellEditText(input, numFmt, options);
      const parsed = parseCellInput(shown, options);
      const label = `${JSON.stringify(input.value)} as ${numFmt} -> ${JSON.stringify(shown)}`;
      assert.equal(parsed.kind, input.kind, label);
      assert.equal(parsed.value, input.value, label);
      assert.equal(parsed.formula, undefined, label);
      return parsed;
    };
    for (let i = 0; i < 2000; i++) {
      // Numbers of up to 15 significant digits over a wide range of magnitudes.
      const digits = 1 + integer(15);
      let mantissa = '';
      for (let d = 0; d < digits; d++) mantissa += String(d === 0 ? 1 + integer(9) : integer(10));
      const value = Number(`${random() < 0.5 ? '-' : ''}${mantissa}e${integer(60) - 40}`);
      roundTrip(cell('number', value), 'General');
      roundTrip(cell('number', value), '#,##0.00');
      // Arbitrary doubles: every bit comes back.
      const full = (random() - 0.5) * Math.pow(10, integer(40) - 20);
      roundTrip(cell('number', full), 'General');
      roundTrip(cell('number', full), '0.0%');

      // Percentages.
      const percent = Number(`${random() < 0.3 ? '-' : ''}${1 + integer(99999)}e${-integer(7)}`);
      const asPercent = roundTrip(cell('number', percent), '0.00%');
      assert.ok(asPercent.numFmtId === 9 || asPercent.numFmtId === 10, String(percent));

      // Dates, with and without a time, and times alone.
      const day = (date1904 ? 0 : 1) + integer(lastDay);
      assert.equal(roundTrip(cell('number', day), 'yyyy-mm-dd').numFmtId, 14, String(day));
      const seconds = 1 + integer(86399);
      assert.equal(roundTrip(cell('number', day + seconds / 86400), 'dd/mm/yyyy hh:mm').numFmtId, 22);
      if (!date1904) assert.equal(roundTrip(cell('number', seconds / 86400), 'h:mm').numFmtId, 21);

      roundTrip(cell('boolean', random() < 0.5), 'General');
      roundTrip(cell('string', texts[integer(texts.length)]), 'General');
      // Random short strings from characters that matter to the parser.
      const alphabet = "0123456789.,-/: %$'=()eEaApPmM₺TLİıtrue#";
      let noise = '';
      for (let n = integer(8); n >= 0; n--) noise += alphabet.charAt(integer(alphabet.length));
      roundTrip(cell('string', noise), 'General');
    }
    for (const fixed of [0, 1, -1, 0.5, 60, 61, 1e15, 1e-9, 999999999999999, 0.07, 1e20, 1e-300, 1.7976931348623157e308]) {
      roundTrip(cell('number', fixed), 'General');
    }
    const errors = [
      '#N/A', '#DIV/0!', '#REF!', '#VALUE!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#GETTING_DATA',
      '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!', '#BUSY!',
    ];
    for (const error of errors) {
      roundTrip(cell('error', error), 'General');
    }
    for (const value of texts) roundTrip(cell('string', value), 'General');
  }
});

test('input: codes that only look like times stay text', () => {
  check([
    ['12A', EN, text('12A')],
    ['5a', EN, text('5a')],
    ['3p', EN, text('3p')],
    ['7P', TR, text('7P')],
    ['0p', EN, text('0p')],
    ['9am', EN, text('9am')],
    ['11PM', EN, text('11PM')],
    ['5öö', TR, text('5öö')],
    ['12a', TR, text('12a')],
    // With a space, or after a time with a colon, the marker is a marker.
    ['5 a', EN, number(time(5, 0), 18)],
    ['9 AM', EN, number(time(9, 0), 18)],
    ['3 ÖS', TR, number(time(15, 0), 18)],
    ['1:30PM', EN, number(time(13, 30), 18)],
    ['1:30a', EN, number(time(1, 30), 18)],
    ['15.03.2024 1:30PM', TR, number(serial(2024, 3, 15) + time(13, 30), 22)],
  ]);
});

test('input: whole numbers a double cannot hold stay text', () => {
  check([
    ['12345678901234567890', EN, text('12345678901234567890')],
    ['1234567890123456', EN, number(1234567890123456)],
    ['9007199254740993', EN, text('9007199254740993')],
    ['9007199254740992', EN, number(9007199254740992)],
    ['4111111111111111', EN, number(4111111111111111)],
    // Sixteen digits below 2^53 are exact, so they are numbers.
    ['4111111111111117', EN, number(4111111111111117)],
    ['99999999999999999', EN, text('99999999999999999')],
    ['+905551234567890123', TR, text('+905551234567890123')],
    ['-12345678901234567890', EN, text('-12345678901234567890')],
    ['(12345678901234567890)', EN, text('(12345678901234567890)')],
    ['12,345,678,901,234,567,890', EN, text('12,345,678,901,234,567,890')],
    ['0012345678901234567890', EN, text('0012345678901234567890')],
    ['$12345678901234567890', EN, text('$12345678901234567890')],
    ['1' + '0'.repeat(25), EN, text('1' + '0'.repeat(25))],
    // Up to 15 digits always fit; so do longer ones that happen to be exact.
    ['999999999999999', EN, number(999999999999999)],
    ['10000000000000000', EN, number(1e16)],
    ['100000000000000000000', EN, number(1e20)],
    // Leading zeros are dropped, as Excel does.
    ['0123', EN, number(123)],
    ['000000000000000000123', EN, number(123)],
    // With a decimal part or an exponent the input is a measurement: rounded like any number.
    ['0.1234567890123456789', EN, number(0.12345678901234568)],
    ['0,1234567890123456789', TR, number(0.12345678901234568)],
    ['12345678901234567890.5', EN, number(12345678901234567000)],
    ['1e30', EN, number(1e30)],
    ['12345678901234567890e0', EN, number(12345678901234567000)],
  ]);
  // Such text needs no apostrophe in the editor: it is read back as text anyway.
  for (const value of ['12345678901234567890', '+905551234567890123', '9007199254740993']) {
    assert.equal(cellEditText(cell('string', value), 'General', EN), value);
    assert.equal(cellEditText(cell('string', value), 'General', TR), value);
  }
  assert.equal(cellEditText(cell('string', '1234567890123456'), 'General', EN), "'1234567890123456");
});

test('input: all error constants', () => {
  for (const error of ['#SPILL!', '#GETTING_DATA', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!', '#BUSY!']) {
    assert.deepEqual(parseCellInput(error, EN), { kind: 'error', value: error });
    assert.deepEqual(parseCellInput(error.toLowerCase(), TR), { kind: 'error', value: error });
    assert.equal(cellEditText(cell('error', error), 'General', TR), error);
  }
});

test('input: dotted dates only where the locale writes dates with dots', () => {
  check([
    ['01.03.2024', EN, text('01.03.2024')],
    ['3.15.2024', EN, text('3.15.2024')],
    ['29.02.2024', EN, text('29.02.2024')],
    ['1.2.3', EN, text('1.2.3')],
    ['2024.03.15', EN, text('2024.03.15')],
    ['1.2024', EN, number(1.2024)],
    ['01.03.2024', TR, number(serial(2024, 3, 1), 14)],
    ['1.2.3', TR, number(serial(2003, 2, 1), 14)],
    // Slashes, dashes and ISO dates in every locale.
    ['3/1/2024', EN, number(serial(2024, 3, 1), 14)],
    ['3-1-2024', EN, number(serial(2024, 3, 1), 14)],
    ['2024-03-01', EN, number(serial(2024, 3, 1), 14)],
    ['1/3/2024', TR, number(serial(2024, 3, 1), 14)],
    ['1-3-2024', TR, number(serial(2024, 3, 1), 14)],
  ]);
  assert.equal(cellEditText(cell('string', '01.03.2024'), 'General', EN), '01.03.2024');
  assert.equal(cellEditText(cell('string', '01.03.2024'), 'General', TR), "'01.03.2024");
});

test('input: scientific notation suggests no format', () => {
  for (const input of ['1E+20', '1.5e3', '-2E-7']) {
    const parsed = parseCellInput(input, EN);
    assert.equal(parsed.kind, 'number', input);
    assert.equal(parsed.numFmtId, undefined, input);
    assert.equal(parsed.numFmt, undefined, input);
  }
  assert.deepEqual(parseCellInput(cellEditText(cell('number', 1e20), 'General', EN), EN), number(1e20));
});
