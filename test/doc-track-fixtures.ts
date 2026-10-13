// Helpers for the tests of tracked changes: a small document that tracks, the
// file it is written as, and what every such file has to be.

import assert from 'node:assert/strict';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import { DocHistory, revisionsOf, type DocEditOptions, type DocRange } from '../src/ops/doc/index';
import { OpcPackage } from '../src/opc';
import { parseXml } from '../src/xml';
import { STYLES } from './doc-ops-fixtures';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { modelPicture, schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();
export const ME: DocEditOptions = { author: 'Me', date: '2026-10-05T10:00:00Z' };
export const YOU: DocEditOptions = { author: 'You', date: '2026-10-06T10:00:00Z' };
const TRACK = '<w:trackRevisions/>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/><w:pgMar w:top="600" w:right="600" w:bottom="600" w:left="600" w:header="150" w:footer="150" w:gutter="0"/></w:sectPr>';
export const ATTRS = ' w:author="Me" w:date="2026-10-05T10:00:00Z"';

export interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
  history: DocHistory<DocRange | undefined>;
  bytes: Uint8Array;
}

export async function open(body: string, extra: Partial<DocxSpec> = {}): Promise<Opened> {
  const bytes = await buildSpec({ body: body + SECT, styles: STYLES, settings: TRACK, ...extra });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  return { pkg, doc, history: new DocHistory<DocRange | undefined>(doc), bytes };
}

export async function written(o: Opened, part = '/word/document.xml'): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  return decoder.decode(await o.pkg.readPart(part));
}

export async function body(o: Opened): Promise<string> {
  const text = await written(o);
  return text.slice(text.indexOf('<w:body>') + 8, text.lastIndexOf('<w:sectPr>'));
}

/** The body as plain data, without where it was read from: equal for two documents that hold the same. */
export const shape = (doc: DocModel): string => JSON.stringify(modelPicture(doc.body.blocks));

/** Every w:id of a tracked change in the XML. */
export function revisionIds(xml: string): string[] {
  const out: string[] = [];
  for (const match of xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|rPrChange|pPrChange|tblPrChange|trPrChange|tcPrChange|tblGridChange|sectPrChange)\b[^>]*? w:id="(\d+)"/g)) out.push(match[1]);
  return out;
}

/** What every written file has to be: well formed, in schema order, with no id of a tracked change twice. */
export async function checkFile(o: Opened, what = ''): Promise<string> {
  const xml = await written(o);
  parseXml(xml);
  assert.deepEqual(schemaOrderViolations(xml), [], what);
  const ids = revisionIds(xml);
  assert.equal(new Set(ids).size, ids.length, `${what}: an id is written twice: ${ids.join(',')}`);
  return xml;
}

/** The tracked changes of a document, as text to compare. */
export const changes = (doc: DocModel): string =>
  revisionsOf(doc)
    .map((revision) => [revision.kind, revision.author, revision.date, revision.text, revision.marks, revision.rows, revision.story, JSON.stringify(revision.start), JSON.stringify(revision.end)].join('|'))
    .join('\n');

/**
 * The model a saved file is read into again says what the model that was written says. With
 * `whole` false only its tracked changes are compared: of a picture that was just put in the
 * reader knows more than the command wrote into the model (its place in the drawing).
 */
export async function checkReopened(o: Opened, what = '', whole = true): Promise<DocModel> {
  await writeDocx(o.pkg, o.doc);
  const again = await readDocx(await OpcPackage.open(await o.pkg.save()));
  // Not the two pictures themselves: a difference would print both in full.
  if (whole) assert.ok(shape(again) === shape(o.doc), `${what}: read again, the model is another`);
  assert.equal(changes(again), changes(o.doc), `${what}: read again, the changes are others`);
  return again;
}

/** What has to hold of every table: each row covers the columns of the grid, and a fixed width is the sum of the grid. */
export function checkTables(doc: DocModel, what = ''): void {
  const walk = (blocks: DocModel['body']['blocks'], depth: number): void => {
    for (const block of blocks) {
      if (block.type !== 'table' || depth > 16) continue;
      for (let r = 0; r < block.rows.length; r++) {
        const row = block.rows[r];
        let covered = (row.props.gridBefore || 0) + (row.props.gridAfter || 0);
        for (const one of row.cells) {
          covered += one.props.gridSpan && one.props.gridSpan > 1 ? one.props.gridSpan : 1;
          walk(one.blocks, depth + 1);
        }
        assert.equal(covered, block.grid.length, `${what}: row ${r} covers ${covered} of ${block.grid.length} grid columns`);
      }
      const width = block.props.width;
      if (width && width.type === 'dxa' && width.value > 0) {
        assert.equal(
          width.value,
          block.grid.reduce((sum, w) => sum + w, 0),
          `${what}: the width of the table is not the sum of its grid ${JSON.stringify(block.grid)}`,
        );
      }
    }
  };
  walk(doc.body.blocks, 0);
}
