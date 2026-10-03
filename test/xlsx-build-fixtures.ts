// Models made in code for the buildXlsx() tests, and the way a test builds one
// and reads the file back.

import assert from 'node:assert/strict';
import { buildXlsx, createSheet, emptyStyleSheet, localName, open, type Cell, type Sheet, type Workbook, type XlsxBuildReport, type XmlElement } from '../src/index';
import type { OpcPackage } from '../src/opc';

export const text = (bytes: Uint8Array) => new TextDecoder().decode(bytes);

export function sheetOf(name: string, id = 1): Sheet {
  return createSheet({ name, sheetId: id, state: 'visible', kind: 'worksheet', part: '' });
}

export function bookOf(...sheets: Sheet[]): Workbook {
  return { sheets, styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
}

export function put(sheet: Sheet, row: number, col: number, cell: Partial<Cell> & Pick<Cell, 'kind' | 'value'>): void {
  let target = sheet.rows.get(row);
  if (!target) sheet.rows.set(row, (target = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [] }));
  target.cells[col] = { style: 0, ...cell };
  sheet.rowCount = Math.max(sheet.rowCount, row + 1);
  sheet.colCount = Math.max(sheet.colCount, col + 1);
}

export interface Built {
  bytes: Uint8Array;
  after: Workbook;
  pkg: OpcPackage;
  warnings: string[];
}

export async function build(workbook: Workbook): Promise<Built> {
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const bytes = await buildXlsx(workbook, { report });
  const file = await open(bytes);
  assert.equal(file.format, 'xlsx');
  return { bytes, after: await file.workbook(), pkg: file.package!, warnings: report.warnings };
}

export function allNamed(el: XmlElement, local: string, out: XmlElement[] = []): XmlElement[] {
  if (localName(el.name) === local) out.push(el);
  for (const child of el.children) if (child.type === 'element') allNamed(child, local, out);
  return out;
}
