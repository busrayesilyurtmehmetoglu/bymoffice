// Changes of a document of every kind that can be shared, made from random numbers alone: what
// the tests of documents that are edited together make their participants do.

import { traced } from '../src/collab/doc/trace';
import type { Block, DocModel } from '../src/model/doc/types';
import * as raw from '../src/ops/doc/index';
import type { DocEdit, DocPosition, DocRange } from '../src/ops/doc/index';
import { docFile, int, pick } from './collab-doc-fixtures';

export const ops = traced(raw);

const TABLE =
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  '<w:tr><w:tc><w:p><w:r><w:t>cell one</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>cell two</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>three</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>';

/** A file with plain and formatted paragraphs, a table and a header. */
export function aloneFile(): Promise<Uint8Array> {
  return docFile(['The first paragraph of the file.', 'A second one, a little longer than the first.', ''], {
    body:
      '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Bold words </w:t></w:r><w:r><w:t>and plain ones.</w:t></w:r></w:p>' +
      TABLE +
      '<w:p><w:r><w:t>After the table.</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
    headers: { rIdH: '<w:p><w:r><w:t>Header text</w:t></w:r></w:p>' },
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
  });
}

/** Every paragraph of a list of blocks, by its path. */
function paragraphs(blocks: Block[], base: number[], out: { path: number[]; length: number }[]): void {
  blocks.forEach((block, index) => {
    if (block.type === 'paragraph') out.push({ path: base.concat([index]), length: raw.paragraphLength(block) });
    else if (block.type === 'table') block.rows.forEach((row, r) => row.cells.forEach((cell, c) => paragraphs(cell.blocks, base.concat([index, r, c]), out)));
  });
}

/** One change, as what makes it on a document; made from random numbers alone, so that it is the same on two documents that hold the same. */
export function make(doc: DocModel, random: () => number, undoable = false): (target: DocModel) => DocEdit {
  const story = random() < 0.12 ? `header:${Array.from(doc.headers.keys())[0]}` : 'body';
  const list: { path: number[]; length: number }[] = [];
  paragraphs(raw.storyOf(doc, story).blocks, [], list);
  const first = pick(random, list);
  const place = (item: { path: number[]; length: number }, offset = int(random, 0, item.length)): DocPosition => ({ story, path: item.path.slice(), offset });
  const at = place(first);
  let near = list[Math.min(list.length - 1, list.indexOf(first) + (random() < 0.3 ? 1 : 0))];
  // Where every step is to be undone, a range stays in one list of blocks, with no table between its ends.
  if (undoable && (near.path.length !== first.path.length || near.path.slice(0, -1).join() !== first.path.slice(0, -1).join() || near.path[near.path.length - 1] - first.path[first.path.length - 1] > 1)) near = first;
  const range: DocRange = { anchor: at, focus: place(near) };
  const inTable = at.path.length > 1;
  const text = pick(random, ['a', 'word', ' ', 'two words', 'x\ny', '\tz', 'İşık', '.']);
  const kinds: [number, (target: DocModel) => DocEdit][] = [
    [34, (target) => ops.insertText(target, at, text)],
    [5, (target) => ops.insertText(target, range, text, { bold: true })],
    [8, (target) => ops.deleteRange(target, range)],
    [8, (target) => ops.deleteBackward(target, at, random() < 0.3 ? 'word' : 'char')],
    [6, (target) => ops.deleteForward(target, at, 'char')],
    [8, (target) => ops.splitParagraph(target, at)],
    [5, (target) => ops.setRunFormat(target, range, { italic: true, color: 'C00000' })],
    [5, (target) => ops.toggleRunFormat(target, range, 'bold')],
    [5, (target) => ops.setParaFormat(target, range, { align: pick(random, ['left', 'center', 'right']) })],
    [3, (target) => ops.toggleList(target, range, random() < 0.5 ? 'bullet' : 'number')],
    [2, (target) => ops.indentList(target, range, 1)],
    [2, (target) => ops.insertPageBreak(target, at)],
    [2, (target) => ops.insertTable(target, at, 2, 2)],
    [2, (target) => ops.insertHyperlink(target, range, 'https://example.com/a')],
  ];
  // What cannot be taken back by commands while a document is edited together (a list that is made, a link that gets
  // its address, a table that changes its rows): left out where every step is to be undone.
  if (undoable) for (let i = kinds.length - 1; i >= 0; i--) if (/toggleList|indentList|insertHyperlink|insertTable/.test(String(kinds[i][1]))) kinds.splice(i, 1);
  if (inTable && !undoable) {
    kinds.push([4, (target) => ops.insertRow(target, at, 'below')]);
    kinds.push([2, (target) => ops.deleteRow(target, at)]);
    kinds.push([3, (target) => ops.insertColumn(target, at, 'right')]);
    kinds.push([2, (target) => ops.deleteColumn(target, at)]);
  }
  let total = 0;
  for (const kind of kinds) total += kind[0];
  let roll = random() * total;
  for (const kind of kinds) {
    roll -= kind[0];
    if (roll < 0) return kind[1];
  }
  return kinds[0][1];
}

