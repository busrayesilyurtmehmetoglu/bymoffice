import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { MAX_NESTING, MAX_XML_DEPTH } from '../src/formats/docx/common';
import { MAX_STYLES } from '../src/formats/docx/styles';
import { IMPLICIT_PARA_PROPS, IMPLICIT_RUN_PROPS, resolveParaProps, resolveRunProps } from '../src/formats/docx/resolve';
import { NumberingCounters } from '../src/formats/docx/numbering';
import { paragraphText, rawOf, type FieldCharItem, type Paragraph, type Run, type Table } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { ZipArchive } from '../src/zip';
import { buildDocx, documentXml, readSpec, type DocxSpec } from './docx-fixtures';

const fails = (spec: DocxSpec, pattern: RegExp) => assert.rejects(() => readSpec(spec), (error: Error) => error instanceof Error && pattern.test(error.message));

test('docx robustness: damaged XML is an error, not a hang or a partial result', async () => {
  const good = documentXml('<w:p><w:r><w:t>metin</w:t></w:r></w:p>');
  for (const cut of [20, 150, good.length - 30, good.length - 5]) {
    await fails({ document: good.slice(0, cut) }, /Invalid XML/);
  }
  await fails({ document: '' }, /Invalid XML/);
  await fails({ document: 'düz metin' }, /Invalid XML/);
  await fails({ document: documentXml('<w:p><w:r></w:p>') }, /Invalid XML.*closing tag/);
  await fails({ document: documentXml('<w:p><w:r><w:t>&bilinmeyen;</w:t></w:r></w:p>') }, /Unknown XML entity/);
  await fails({ document: documentXml('<w:p w:a="1></w:p>') }, /Invalid XML/);
  await fails({ document: good + '<w:fazla/>' }, /more than one root/);
  await fails({ document: '<!DOCTYPE w:document [<!ENTITY x "y">]>' + good }, /DTD declarations are not allowed/);
  await fails({ body: '<w:p/>', styles: '<w:style></w:styles>' }, /Invalid XML/);
  await fails({ body: '<w:p/>', footnotes: '<w:footnote w:id="1"><w:p></w:footnote>' }, /Invalid XML/);
  await fails({ body: '<w:p/>', headers: { rId1: '<w:p><!DOCTYPE x></w:p>' } }, /DTD/);
  await fails({ body: '<w:p/>', parts: [['word/_rels/document.xml.rels', '<Relationships><Relationship']] }, /Invalid XML/);

  // The right container with the wrong content.
  await fails({ document: '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>' }, /not a WordprocessingML document/);
  await fails({ document: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' }, /no body/);
  await fails({ document: '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:document>' }, /no body/);

  const zip = new ZipArchive();
  zip.set('[Content_Types].xml', new TextEncoder().encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'));
  await assert.rejects(async () => readDocx(await OpcPackage.open(await zip.write())), /no main document part/);
  // The package names a document part that is not there.
  zip.set('_rels/.rels', new TextEncoder().encode('<Relationships><Relationship Id="rId1" Type="x/officeDocument" Target="word/document.xml"/></Relationships>'));
  await assert.rejects(async () => readDocx(await OpcPackage.open(await zip.write())), /no main document part/);
});

test('docx robustness: deep nesting ends in an error, never in a stack overflow', async () => {
  const deep = (open: string, close: string, depth: number, core = '') => open.repeat(depth) + core + close.repeat(depth);
  const limit = /nested deeper than \d+ levels/;
  const depth = 10000;
  await fails({ body: deep('<w:tbl><w:tr><w:tc>', '</w:tc></w:tr></w:tbl>', depth, '<w:p/>') }, limit);
  await fails({ body: deep('<w:sdt><w:sdtContent>', '</w:sdtContent></w:sdt>', depth, '<w:p/>') }, limit);
  await fails({ body: `<w:p>${deep('<w:ins w:id="1">', '</w:ins>', depth, '<w:r/>')}</w:p>` }, limit);
  await fails({ body: `<w:p>${deep('<w:smartTag>', '</w:smartTag>', depth)}</w:p>` }, limit);
  await fails({ body: `<w:p><w:r>${deep('<x>', '</x>', depth)}</w:r></w:p>` }, limit);
  await fails({ body: `<w:p><w:pPr>${deep('<w:pPrChange>', '</w:pPrChange>', depth)}</w:pPr></w:p>` }, limit);
  await fails({ body: `<w:p><w:r><w:rPr>${deep('<w:x>', '</w:x>', depth)}</w:rPr></w:r></w:p>` }, limit);
  await fails({ body: deep('<w:unknown>', '</w:unknown>', depth) }, limit);
  await fails({ body: '<w:p/>', styles: deep('<w:style w:styleId="a">', '</w:style>', depth) }, limit);
  await fails({ body: '<w:p/>', numbering: deep('<w:abstractNum w:abstractNumId="1">', '</w:abstractNum>', depth) }, limit);
  await fails({ body: '<w:p/>', footnotes: `<w:footnote w:id="1">${deep('<w:tbl><w:tr><w:tc>', '</w:tc></w:tr></w:tbl>', depth)}</w:footnote>` }, limit);
  const box = (inner: string) => `<w:p><w:r><w:pict><v:textbox><w:txbxContent>${inner}</w:txbxContent></v:textbox></w:pict></w:r></w:p>`;
  let boxes = '<w:p/>';
  for (let i = 0; i < MAX_NESTING + 2; i++) boxes = box(boxes);
  await fails({ body: boxes }, limit);
  await fails({ body: `<w:p><w:r>${deep('<mc:AlternateContent><mc:Fallback>', '</mc:Fallback></mc:AlternateContent>', 100)}</w:r></w:p>` }, limit);

  // Nesting that real documents have is fine.
  const ok = await readSpec({ body: deep('<w:tbl><w:tr><w:tc>', '<w:p/></w:tc></w:tr></w:tbl>', 20, '<w:p><w:r><w:t>dip</w:t></w:r></w:p>') });
  let block = ok.body.blocks[0];
  let levels = 0;
  while (block.type === 'table') {
    block = block.rows[0].cells[0].blocks[0];
    levels++;
  }
  assert.equal(levels, 20);
  assert.equal(paragraphText(block as Paragraph), 'dip');
  assert.ok(MAX_XML_DEPTH > MAX_NESTING);

  // Many unclosed fields: assembled up to a limit, the rest left as plain items.
  const fields = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'.repeat(500) + '<w:r><w:instrText>X</w:instrText></w:r>' + '<w:r><w:fldChar w:fldCharType="end"/></w:r>'.repeat(500);
  const many = await readSpec({ body: `<w:p>${fields}<w:r><w:t>son</w:t></w:r></w:p>` });
  const p = many.body.blocks[0] as Paragraph;
  const assembled = p.children.filter((inline) => ((inline as Run).content[0] as FieldCharItem).field !== undefined).length;
  assert.equal(assembled, 32);
  assert.equal((p.children[1001] as Run).fieldCode, undefined);
  assert.equal(paragraphText(p), 'son');
});

test('docx robustness: missing parts and broken relationships do not stop the reading', async () => {
  const doc = await readSpec({
    missing: ['styles', 'numbering', 'settings', 'fontTable', 'footnotes', 'endnotes', 'comments', 'theme'],
    rels: [
      ['rIdH', 'header', 'header9.xml'],
      ['rIdF', 'footer', 'https://example.com/footer.xml', true],
      ['rIdImg', 'image', 'media/yok.png'],
    ],
    body:
      '<w:p><w:pPr><w:pStyle w:val="Yok"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="5"/></w:numPr></w:pPr>' +
      '<w:r><w:rPr><w:rStyle w:val="Yok"/><w:b/></w:rPr><w:t>metin</w:t><w:footnoteReference w:id="2"/></w:r>' +
      '<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData uri="x/picture"><pic:pic><pic:blipFill><a:blip r:embed="rIdYok"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
      '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>',
  });
  assert.equal(doc.warnings.length, 9);
  assert.ok(doc.warnings.every((warning) => /is missing/.test(warning)));
  assert.equal(doc.styles.styles.size, 0);
  assert.equal(doc.numbering.nums.size, 0);
  assert.equal(doc.footnotes, undefined);
  assert.equal(doc.headers.size, 0);
  assert.equal(doc.footers.size, 0);
  assert.equal(doc.body.section?.headers?.default, 'rIdH');

  const p = doc.body.blocks[0] as Paragraph;
  assert.equal(paragraphText(p), 'metin');
  const drawing = (p.children[1] as Run).content[0];
  assert.ok(drawing.type === 'drawing');
  assert.deepEqual(drawing.image, { rId: 'rIdYok' });
  // Resolving against the empty tables works and gives the direct formatting on top of what Word implies.
  assert.deepEqual(resolveParaProps(doc, p.props), { ...IMPLICIT_PARA_PROPS, numbering: { level: 0, numId: 5 } });
  assert.deepEqual(resolveRunProps(doc, p.props, (p.children[0] as Run).props), { ...IMPLICIT_RUN_PROPS, sizeCs: 20, style: 'Yok', bold: true });
  assert.equal(new NumberingCounters(doc.numbering, doc.styles).next(5, 0), undefined);

  // No relationships part at all.
  const bytes = await buildDocx({ body: '<w:p><w:hyperlink r:id="rId1"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p>' });
  const pkg = await OpcPackage.open(bytes);
  pkg.zip.delete('word/_rels/document.xml.rels');
  const bare = await readDocx(pkg);
  assert.equal(bare.body.relationships.size, 0);
  assert.equal(paragraphText(bare.body.blocks[0] as Paragraph), 'x');
  assert.equal(bare.warnings.length, 1);
});

test('docx robustness: numbers are validated', async () => {
  const bad = ['NaN', 'Infinity', '-1', '1e5', '0x10', '99999999999999999999', '', '1.5.2', '12px', '2 4'];
  const runs = bad.map((value) => `<w:r><w:rPr><w:sz w:val="${value}"/></w:rPr></w:r>`).join('');
  const doc = await readSpec({
    body:
      `<w:p>${runs}</w:p>` +
      '<w:p><w:pPr><w:ind w:left="1000001" w:right="-1000001" w:firstLine="-5" w:hanging="1e3"/><w:spacing w:before="-1" w:after="NaN" w:line="abc" w:lineRule="bogus"/>' +
      '<w:numPr><w:ilvl w:val="9"/><w:numId w:val="-3"/></w:numPr><w:outlineLvl w:val="10"/><w:tabs><w:tab w:val="left" w:pos="x"/><w:tab w:val="left" w:pos="500"/></w:tabs></w:pPr></w:p>' +
      '<w:tbl><w:tblGrid><w:gridCol w:w="-5"/><w:gridCol w:w="99999999"/><w:gridCol/><w:gridCol w:w="2000"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:trHeight w:val="-10"/><w:gridBefore w:val="5000"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="100000"/><w:tcW w:w="1e9" w:type="dxa"/><w:vMerge w:val="sideways"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:p><w:r><w:rPr><w:sz w:val="7200"/><w:position w:val="-7201"/><w:w w:val="601"/><w:spacing w:val="1.5"/><w:b w:val="maybe"/><w:color w:val="12345"/><w:u w:val="single" w:color="red"/></w:rPr>' +
      '<w:footnoteReference w:id="x"/><w:sym w:font="A" w:char="12345"/><w:sym w:char="zz"/></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="-1" w:h="1e9"/><w:pgMar w:top="-500" w:left="-500"/><w:cols w:num="0"/><w:pgNumType w:start="-1"/></w:sectPr>',
  });
  const first = doc.body.blocks[0] as Paragraph;
  for (const inline of first.children) {
    const props = (inline as Run).props;
    assert.equal(props.size, undefined);
    // The element is kept so that the file does not lose what it said.
    assert.equal(props.extras?.length, 1);
  }

  const second = (doc.body.blocks[1] as Paragraph).props;
  assert.deepEqual(second.indent, {});
  assert.deepEqual(second.spacing, {});
  assert.deepEqual(second.numbering, {});
  assert.equal(second.outlineLevel, undefined);
  assert.deepEqual(second.tabs, [{ pos: 500, align: 'left' }]);
  assert.deepEqual(second.extras!.map((extra) => extra.el.name).sort(), ['w:ind', 'w:numPr', 'w:outlineLvl', 'w:spacing', 'w:tabs']);

  const table = doc.body.blocks[2] as Table;
  assert.deepEqual(table.grid, [0, 0, 0, 2000]);
  assert.equal(table.extras?.length, 1);
  assert.deepEqual({ ...table.rows[0].props, extras: undefined }, { extras: undefined });
  assert.deepEqual({ ...table.rows[0].cells[0].props, extras: undefined }, { extras: undefined });
  assert.equal(table.rows[0].cells[0].props.extras?.length, 3);

  const run = (doc.body.blocks[3] as Paragraph).children[0] as Run;
  assert.equal(run.props.size, 7200);
  assert.deepEqual({ ...run.props, extras: undefined }, { size: 7200, spacing: 1.5, underline: { style: 'single' }, extras: undefined });
  assert.deepEqual(run.content.map((item) => item.type), ['unknown', 'unknown', 'unknown']);

  const section = doc.body.section!;
  assert.deepEqual(section.page, {});
  assert.deepEqual(section.margins, { top: -500 });
  assert.deepEqual(section.columns, {});
  assert.deepEqual(section.pageNumbering, {});
  assert.equal(section.extras?.length, 4);
});

test('docx robustness: tables that are too large to be genuine are refused', async () => {
  let styles = '';
  for (let i = 0; i <= MAX_STYLES; i++) styles += `<w:style w:styleId="s${i}"/>`;
  await fails({ body: '<w:p/>', styles }, /more than \d+ styles/);
});

test('docx performance: a document of 50,000 paragraphs', async () => {
  const count = 50000;
  const parts: string[] = [];
  for (let i = 0; i < count; i++) {
    parts.push(
      `<w:p w:rsidR="00A1B2C3" w:rsidRDefault="00D4E5F6"><w:pPr><w:pStyle w:val="${i % 50 === 0 ? 'Heading1' : 'Normal'}"/><w:spacing w:after="${(i % 4) * 60}"/></w:pPr>` +
        `<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Paragraf ${i} </w:t></w:r>` +
        `<w:r><w:t>sıradan bir metin parçası, biraz daha uzun olsun diye &amp; devam ediyor</w:t></w:r>` +
        (i % 10 === 0 ? '<w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:tab/><w:t>ek</w:t></w:r>' : '') +
        '</w:p>',
    );
  }
  const bytes = await buildDocx({ body: parts.join('') });
  const pkg = await OpcPackage.open(bytes);
  const before = process.memoryUsage().heapUsed;
  const start = performance.now();
  const doc = await readDocx(pkg);
  const elapsed = performance.now() - start;
  const heap = process.memoryUsage().heapUsed - before;
  console.log(`docx: ${count} paragraphs read in ${elapsed.toFixed(0)} ms, heap +${(heap / 1048576).toFixed(0)} MB`);

  assert.equal(doc.body.blocks.length, count);
  const last = doc.body.blocks[count - 1] as Paragraph;
  assert.equal(paragraphText(last), `Paragraf ${count - 1} sıradan bir metin parçası, biraz daha uzun olsun diye & devam ediyor`);
  assert.ok(rawOf(doc.body, last)!.startsWith('<w:p w:rsidR'));
  // Property bags are shared: a handful of objects for the whole document.
  const paraProps = new Set(doc.body.blocks.map((block) => (block as Paragraph).props));
  const runProps = new Set<object>();
  for (const block of doc.body.blocks) for (const inline of (block as Paragraph).children) runProps.add((inline as Run).props);
  assert.equal(paraProps.size, 6);
  assert.equal(runProps.size, 3);
  // Generous: the usual time is a small fraction of this.
  assert.ok(elapsed < 100000, `reading took ${elapsed.toFixed(0)} ms`);
});
