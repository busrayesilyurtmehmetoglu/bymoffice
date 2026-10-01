import assert from 'node:assert/strict';
import { test } from 'node:test';
import { cellText, formatGeneral } from '../src/render/cell-text';

test('cell text: General number display', () => {
  const cases: [number, string][] = [
    [0, '0'],
    [-0, '0'],
    [1, '1'],
    [-12.5, '-12.5'],
    [0.1 + 0.2, '0.3'],
    [1234567890.123, '1234567890'],
    [123456.789012345, '123456.789'],
    [0.000123456789012, '0.000123457'],
    [99999999999, '99999999999'],
    [100000000000, '1E+11'],
    [1.23456789e15, '1.23457E+15'],
    [-1.5e-12, '-1.5E-12'],
    [1 / 3, '0.333333333'],
    // Small numbers: plain while the digits fit in 11 characters, scientific otherwise.
    [1e-7, '0.0000001'],
    [1e-9, '0.000000001'],
    [1.2345e-7, '1.2345E-07'],
    [0.0000123456, '1.23456E-05'],
    [1.2e-9, '1.2E-09'],
    [1e-10, '1E-10'],
    [-0.000001, '-0.000001'],
    [0.00012345678, '0.000123457'],
    // The 11 digit boundary.
    [99999999999.9, '1E+11'],
    [12345678901.4, '12345678901'],
    [123456789012, '1.23457E+11'],
    [-1234567.891234, '-1234567.891'],
    [1234567.891234, '1234567.891'],
    [Infinity, '#NUM!'],
    [-Infinity, '#NUM!'],
    [NaN, '#NUM!'],
  ];
  for (const [value, text] of cases) assert.equal(formatGeneral(value), text, String(value));
});

test('cell text: by cell kind', () => {
  assert.equal(cellText(undefined), '');
  assert.equal(cellText({ kind: 'blank', value: null, style: 0 }), '');
  assert.equal(cellText({ kind: 'boolean', value: true, style: 0 }), 'TRUE');
  assert.equal(cellText({ kind: 'boolean', value: false, style: 0 }), 'FALSE');
  assert.equal(cellText({ kind: 'error', value: '#DIV/0!', style: 0 }), '#DIV/0!');
  assert.equal(cellText({ kind: 'string', value: ' Çay ', style: 0 }), ' Çay ');
  assert.equal(cellText({ kind: 'number', value: 19.5, style: 0 }), '19.5');
});
