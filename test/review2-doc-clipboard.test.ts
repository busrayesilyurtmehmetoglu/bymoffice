// Second review of the copy and paste of documents: a copy carries no tracked change, and a
// copy that names more styles or lists than it can carry is not pasted into another document.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type DocModel, type Paragraph } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { DocHistory, fragmentOf, insertFragment, type DocFragment, type DocPosition, type DocRange } from '../src/ops/doc/index';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';
const ANN = ' w:author="Ann" w:date="2020-01-01T00:00:00Z"';
const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const decoder = new TextDecoder();

async function open(spec: DocxSpec): Promise<{ pkg: OpcPackage; doc: DocModel }> {
  const full: DocxSpec = { styles: STYLES, ...spec };
  if (full.body !== undefined && full.body.indexOf('<w:sectPr') < 0) full.body += SECT;
  const pkg = await OpcPackage.open(await buildDocx(full));
  return { pkg, doc: await readDocx(pkg) };
}

const at = (path: number[] | number, offset: number): DocPosition => ({ story: 'body', path: typeof path === 'number' ? [path] : path, offset });
const range = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });

function paste(doc: DocModel, fragment: DocFragment, where: DocPosition | DocRange): void {
  const history = new DocHistory(doc);
  history.applyEdit(insertFragment(doc, where, fragment), undefined);
  history.seal();
}

async function written(o: { pkg: OpcPackage; doc: DocModel }): Promise<string> {
  await writeDocx(o.pkg, o.doc);
  const saved = await OpcPackage.open(await o.pkg.save());
  await readDocx(saved);
  return decoder.decode(await saved.readPart('/word/document.xml'));
}

/** The ids of everything in the text that records a tracked change, in order. */
const revisionIds = (xml: string) => [...xml.matchAll(/<w:(ins|del|moveFrom|moveTo|cellIns|cellDel|cellMerge|[A-Za-z]+Change)\b[^>]*? w:id="(\d+)"/g)].map((match) => `${match[1]} ${match[2]}`);

/**
 * P0 "keep" + deleted "DEAD" and a deleted tab + inserted "INS", its mark a tracked deletion.
 * P1 "second": its mark a tracked insertion, its formatting and that of its run changed.
 * A table whose row was inserted, with a cell that was inserted and one whose properties changed, and a changed grid.
 * P3 "tail".
 */
const SOURCE =
  `<w:p><w:pPr><w:rPr><w:del w:id="7"${ANN}/></w:rPr></w:pPr>${r('keep ')}<w:del w:id="5"${ANN}><w:r><w:delText>DEAD</w:delText></w:r><w:r><w:tab/></w:r></w:del><w:ins w:id="6"${ANN}>${r('INS')}</w:ins></w:p>` +
  `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:ins w:id="8"${ANN}/><w:b/><w:rPrChange w:id="11"${ANN}><w:rPr/></w:rPrChange></w:rPr><w:pPrChange w:id="9"${ANN}><w:pPr/></w:pPrChange></w:pPr>` +
  `${r('second', `<w:i/><w:rPrChange w:id="10"${ANN}><w:rPr/></w:rPrChange>`)}</w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblPrChange w:id="12"${ANN}><w:tblPr/></w:tblPrChange></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:tblGridChange w:id="13"><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid></w:tblGridChange></w:tblGrid>` +
  `<w:tr><w:trPr><w:ins w:id="14"${ANN}/><w:trPrChange w:id="15"${ANN}><w:trPr/></w:trPrChange></w:trPr>` +
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:cellIns w:id="16"${ANN}/></w:tcPr><w:p>${r('one')}</w:p></w:tc>` +
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcPrChange w:id="17"${ANN}><w:tcPr/></w:tcPrChange></w:tcPr><w:p>${r('two')}</w:p></w:tc></w:tr></w:tbl>` +
  `<w:p>${r('tail')}</w:p>`;

/** A document that tracks nothing itself and has revisions under the ids the source uses. */
const TARGET = `<w:p>${r('target ')}<w:ins w:id="7" w:author="Bob" w:date="2021-01-01T00:00:00Z">${r('theirs')}</w:ins></w:p><w:p><w:del w:id="9" w:author="Bob" w:date="2021-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del>${r('end')}</w:p>`;

test('review 2: a copy carries no tracked change into another document: no revision of the source, no id twice, no author', async () => {
  const source = await open({ body: SOURCE });
  // The reader took every record the test is about: without them it would prove nothing.
  const before = revisionIds(await written(await open({ body: SOURCE })));
  assert.deepEqual(before.map((entry) => entry.split(' ')[1]).sort((a, b) => Number(a) - Number(b)), ['5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15', '16', '17']);
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(3, 2)));

  const target = await open({ body: TARGET });
  paste(target.doc, fragment, at(1, 3));
  const xml = await written(target);
  // What the target had, and nothing else that records a change.
  assert.deepEqual(revisionIds(xml), ['ins 7', 'del 9']);
  assert.ok(xml.indexOf('Ann') < 0, 'the author of the source came along');
  assert.ok(xml.indexOf('DEAD') < 0, 'deleted text is not part of a copy');
  const back = await readDocx(await OpcPackage.open(await target.pkg.save()));
  const texts = back.body.blocks.filter((block): block is Paragraph => block.type === 'paragraph').map(paragraphText);
  // The mark that was a tracked deletion is a plain mark: the two paragraphs stay two.
  assert.deepEqual(texts, ['target theirs', 'endkeep INS', 'second', 'ta']);
  assert.equal(back.body.blocks.filter((block) => block.type === 'table').length, 1);
  // The formatting as it is now came along.
  assert.ok(/<w:jc w:val="center"\/>/.test(xml) && /<w:i\/>/.test(xml), xml);
});

test('review 2: pasted into the document it was taken from, a copy brings no revision a second time', async () => {
  const source = await open({ body: SOURCE });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(3, 2)));
  paste(source.doc, fragment, at(3, 4));
  const ids = revisionIds(await written(source));
  assert.equal(new Set(ids.map((entry) => entry.split(' ')[1])).size, ids.length, `an id twice: ${ids.join(', ')}`);
  assert.equal(ids.length, 13, 'the revisions of the document, once');
  // A range inside one paragraph, and whole cells of the table.
  const inside = fragmentOf(source.doc, range(at(1, 1), at(1, 4)));
  paste(source.doc, inside, at(0, 0));
  const cells = fragmentOf(source.doc, range(at([2, 0, 0, 0], 0), at([2, 0, 1, 0], 3)));
  paste(source.doc, cells, at(0, 0));
  const after = revisionIds(await written(source));
  assert.equal(after.length, 13, after.join(', '));
});

const style = (id: string, basedOn?: string) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${id}"/>${basedOn ? `<w:basedOn w:val="${basedOn}"/>` : ''}</w:style>`;

test('review 2: a copy that names more styles than it can carry is not pasted into another document, and is into its own', async () => {
  // A chain of 300 styles, each based on the next; the paragraph has the first.
  let styles = STYLES;
  for (let i = 0; i < 300; i++) styles += style(`S${i}`, i < 299 ? `S${i + 1}` : 'Normal');
  const source = await open({ styles, body: `<w:p><w:pPr><w:pStyle w:val="S0"/></w:pPr>${r('styled')}</w:p><w:p>${r('next')}</w:p>` });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(1, 2)));
  const target = await open({ styles: STYLES + style('S299'), body: `<w:p>${r('target')}</w:p>` });
  assert.throws(() => insertFragment(target.doc, at(0, 0), fragment), RangeError);
  assert.equal(paragraphText(target.doc.body.blocks[0] as Paragraph), 'target');
  assert.equal(target.doc.styles.styles.has('S0'), false);
  // In its own document the styles are there already.
  paste(source.doc, fragment, at(1, 4));
  assert.equal(paragraphText(source.doc.body.blocks[1] as Paragraph), 'nextstyled');

  // A copy within the limit goes through as before.
  let few = STYLES;
  for (let i = 0; i < 20; i++) few += style(`S${i}`, i < 19 ? `S${i + 1}` : 'Normal');
  const small = await open({ styles: few, body: `<w:p><w:pPr><w:pStyle w:val="S0"/></w:pPr>${r('styled')}</w:p><w:p>${r('next')}</w:p>` });
  const other = await open({ body: `<w:p>${r('target')}</w:p>` });
  paste(other.doc, fragmentOf(small.doc, range(at(0, 0), at(1, 2))), at(0, 0));
  assert.equal(other.doc.styles.styles.has('S19'), true);
});
