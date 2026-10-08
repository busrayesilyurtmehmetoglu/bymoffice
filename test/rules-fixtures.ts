// Shared by the tests of rules that are edited (data validations, conditional formats):
// loading and saving a workbook, comparing archives, and what makes a package a package.

import assert from 'node:assert/strict';
import { writeXlsx } from '../src/formats/xlsx/write';
import { open, parseXml, ZipArchive } from '../src/index';
import type { OfficeFile, Workbook } from '../src/index';
import { History } from '../src/ops/ops';

const decoder = new TextDecoder();

export interface Loaded {
  bytes: Uint8Array;
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

export async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { bytes, file, wb, history: new History(wb) };
}

export async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

export const reopen = async (bytes: Uint8Array) => (await open(bytes)).workbook();
export const partNames = (bytes: Uint8Array) => ZipArchive.read(bytes).names().sort();
export const partText = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));

export function sameBytes(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

/** Every part is named by the content types, every relationship leads to a part, no id is used twice, every part is well-formed. */
export async function checkPackage(bytes: Uint8Array): Promise<void> {
  const zip = ZipArchive.read(bytes);
  const all = new Set(zip.names());
  const types = parseXml(await zip.read('[Content_Types].xml')).root;
  const defaults = new Set<string>();
  const overrides = new Set<string>();
  for (const node of types.children) {
    if (node.type !== 'element') continue;
    if (node.name === 'Default') defaults.add((node.attrs.get('Extension') || '').toLowerCase());
    if (node.name === 'Override') {
      const part = (node.attrs.get('PartName') || '').slice(1);
      assert.ok(all.has(part), `an override names ${part}, which is not in the package`);
      assert.ok(!overrides.has(part), `two overrides for ${part}`);
      overrides.add(part);
    }
  }
  for (const name of all) {
    if (name === '[Content_Types].xml') continue;
    assert.ok(overrides.has(name) || defaults.has(name.slice(name.lastIndexOf('.') + 1).toLowerCase()), `${name} has no content type`);
    if (/\.(xml|rels|vml)$/.test(name)) parseXml(await zip.read(name));
    if (!/\.rels$/.test(name)) continue;
    const source = name.replace(/_rels\/([^/]*)\.rels$/, '$1');
    const ids = new Set<string>();
    for (const rel of parseXml(await zip.read(name)).root.children) {
      if (rel.type !== 'element') continue;
      const id = rel.attrs.get('Id') || '';
      assert.ok(!ids.has(id), `${name}: the id ${id} is used twice`);
      ids.add(id);
      if (rel.attrs.get('TargetMode') === 'External') continue;
      const target = rel.attrs.get('Target') || '';
      const from = source.split('/').slice(0, -1);
      const path = target.charAt(0) === '/' ? target.slice(1).split('/') : from.concat(target.split('/'));
      const resolved: string[] = [];
      for (const step of path) {
        if (step === '..') resolved.pop();
        else if (step !== '.' && step !== '') resolved.push(step);
      }
      assert.ok(all.has(resolved.join('/')), `${name}: ${id} leads to ${resolved.join('/')}, which is not in the package`);
    }
  }
}

const CT_WORKSHEET = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

/** The children of a worksheet come in the order of the schema. */
export function checkSheetOrder(xml: string, label = 'sheet'): void {
  let last = -1;
  for (const node of parseXml(xml).root.children) {
    if (node.type !== 'element') continue;
    const rank = CT_WORKSHEET.indexOf(node.name.slice(node.name.indexOf(':') + 1));
    assert.ok(rank >= 0, `${label}: <${node.name}> is no child of a worksheet`);
    assert.ok(rank >= last, `${label}: <${node.name}> comes after an element the schema puts behind it`);
    last = rank;
  }
}
