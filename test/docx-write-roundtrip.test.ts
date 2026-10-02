// The docx writer, block by block: a synthetic document is read, every block
// is written from the model (the `regenerate` option), and the result is read
// again. The model must come out the same, and for sources that are spelled
// the way the writer spells them, so must the XML.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { writeDocx } from '../src/formats/docx/write';
import type { DocModel } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import { buildDocx as buildSpec, type DocxSpec } from './docx-fixtures';
import { canonicalText, documentPicture, firstDifference, schemaOrderViolations } from './docx-write-helpers.test';

const decoder = new TextDecoder();

interface Regenerated {
  source: DocModel;
  result: DocModel;
  /** The document part before and after. */
  before: string;
  after: string;
  pkg: OpcPackage;
}

async function regenerate(spec: DocxSpec): Promise<Regenerated> {
  const bytes = await buildSpec(spec);
  const pkg = await OpcPackage.open(bytes);
  const before = decoder.decode(await pkg.readPart('/word/document.xml'));
  const source = await readDocx(pkg);
  const picture = documentPicture(source);
  await writeDocx(pkg, source, { regenerate: true });
  const saved = await OpcPackage.open(await pkg.save());
  const result = await readDocx(saved);
  assert.equal(firstDifference(picture, documentPicture(result)), undefined, 'the model read back differs');
  const after = decoder.decode(await saved.readPart('/word/document.xml'));
  return { source, result, before, after, pkg: saved };
}

/** Reads, regenerates and checks: same model, same XML (in canonical form), children in schema order. */
async function exact(body: string, spec: DocxSpec = {}): Promise<Regenerated> {
  const out = await regenerate({ ...spec, body });
  assert.equal(canonicalText(out.after), canonicalText(out.before));
  assert.deepEqual(schemaOrderViolations(out.after), []);
  return out;
}

const PPR =
  '<w:pPr><w:pStyle w:val="Heading1"/><w:keepNext/><w:keepLines w:val="0"/><w:pageBreakBefore/>' +
  '<w:framePr w:dropCap="drop" w:lines="3" w:w="2000" w:h="1000" w:vSpace="10" w:hSpace="20" w:wrap="around" w:hAnchor="text" w:vAnchor="page" w:x="100" w:xAlign="left" w:y="200" w:yAlign="top" w:hRule="exact" w:anchorLock="1"/>' +
  '<w:widowControl w:val="0"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr><w:suppressLineNumbers/>' +
  '<w:pBdr><w:top w:val="single" w:sz="4" w:space="1" w:color="FF0000"/><w:left w:val="double" w:sz="6" w:space="4" w:color="auto"/>' +
  '<w:bottom w:val="dashed" w:sz="8" w:space="1" w:color="0000FF" w:themeColor="accent1" w:themeTint="99" w:themeShade="0F"/><w:right w:val="nil"/>' +
  '<w:between w:val="single" w:sz="4" w:space="1" w:color="000000" w:shadow="1" w:frame="0"/><w:bar w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:pBdr>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="FFFF00" w:themeFill="accent2" w:themeFillTint="33" w:themeFillShade="80"/>' +
  '<w:tabs><w:tab w:val="left" w:pos="720"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/><w:tab w:val="clear" w:pos="1440"/></w:tabs>' +
  '<w:suppressAutoHyphens/><w:kinsoku w:val="0"/><w:wordWrap w:val="0"/><w:overflowPunct w:val="0"/><w:topLinePunct/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/>' +
  '<w:bidi/><w:adjustRightInd w:val="0"/><w:snapToGrid w:val="0"/>' +
  '<w:spacing w:before="120" w:beforeLines="50" w:beforeAutospacing="1" w:after="240" w:afterLines="100" w:afterAutospacing="0" w:line="360" w:lineRule="auto"/>' +
  '<w:ind w:left="720" w:leftChars="100" w:right="360" w:rightChars="50" w:hanging="360" w:hangingChars="150"/>' +
  '<w:contextualSpacing/><w:mirrorIndents/><w:jc w:val="both"/><w:textDirection w:val="lrTb"/><w:textAlignment w:val="center"/><w:outlineLvl w:val="2"/>' +
  '<w:cnfStyle w:val="100000000000"/><w:rPr><w:ins w:id="5" w:author="A" w:date="2020-01-01T00:00:00Z"/><w:b/></w:rPr>' +
  '<w:pPrChange w:id="8" w:author="A"><w:pPr><w:jc w:val="left"/></w:pPr></w:pPrChange></w:pPr>';

const RPR =
  '<w:rPr><w:rStyle w:val="Strong"/>' +
  '<w:rFonts w:hint="eastAsia" w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun" w:cs="Arial" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>' +
  '<w:b/><w:bCs w:val="0"/><w:i/><w:iCs/><w:caps/><w:smallCaps w:val="0"/><w:strike/><w:dstrike w:val="0"/><w:outline/><w:shadow/><w:emboss/><w:imprint/>' +
  '<w:noProof/><w:snapToGrid w:val="0"/><w:vanish/><w:webHidden/>' +
  '<w:color w:val="FF0000" w:themeColor="accent1" w:themeTint="80" w:themeShade="0F"/><w:spacing w:val="20"/><w:w w:val="150"/><w:kern w:val="28"/><w:position w:val="-6"/>' +
  '<w:sz w:val="24"/><w:szCs w:val="26"/><w:highlight w:val="yellow"/><w:u w:val="single" w:color="00FF00"/><w:effect w:val="shimmer"/>' +
  '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:shd w:val="pct10" w:color="auto" w:fill="CCCCCC"/><w:vertAlign w:val="superscript"/>' +
  '<w:rtl/><w:cs/><w:em w:val="dot"/><w:lang w:val="tr-TR" w:eastAsia="ja-JP" w:bidi="ar-SA"/><w:specVanish/>' +
  '<w:rPrChange w:id="9" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr>';

test('paragraph properties: every field goes back, in the order of the schema', async () => {
  const out = await exact(`<w:p w:rsidR="00AB12CD" w14:paraId="1A2B3C4D">${PPR}<w:r><w:t>x</w:t></w:r></w:p>`);
  assert.ok(out.after.indexOf('w14:paraId="1A2B3C4D"') > 0);
});

test('run properties: every field goes back, in the order of the schema', async () => {
  await exact(`<w:p><w:r w:rsidRPr="00112233">${RPR}<w:t>x</w:t></w:r></w:p>`);
});

test('property children written out of order come back in the order of the schema', async () => {
  const out = await regenerate({
    body:
      '<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:sz w:val="20"/><w:b/></w:rPr><w:spacing w:after="0"/><w:ind w:left="10"/><w:pStyle w:val="A"/><w:keepNext/></w:pPr>' +
      '<w:r><w:rPr><w:lang w:val="en-US"/><w:sz w:val="20"/><w:color w:val="FF0000"/><w:i/><w:rFonts w:ascii="Arial"/></w:rPr><w:t>x</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblLook w:val="04A0"/><w:tblW w:w="0" w:type="auto"/><w:tblStyle w:val="T"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid>' +
      '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:vAlign w:val="center"/><w:shd w:val="clear" w:fill="EEEEEE"/><w:tcW w:w="100" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>' +
      '<w:sectPr><w:docGrid w:linePitch="360"/><w:pgMar w:top="1" w:right="2" w:bottom="3" w:left="4" w:header="5" w:footer="6" w:gutter="0"/><w:pgSz w:w="100" w:h="200"/></w:sectPr>',
  });
  assert.notDeepEqual(schemaOrderViolations(out.before), []);
  assert.deepEqual(schemaOrderViolations(out.after), []);
  assert.ok(out.after.indexOf('<w:pStyle w:val="A"/><w:keepNext/><w:spacing w:after="0"/><w:ind w:left="10"/><w:jc w:val="center"/><w:rPr><w:b/><w:sz w:val="20"/></w:rPr>') > 0);
  assert.ok(out.after.indexOf('<w:rFonts w:ascii="Arial"/><w:i/><w:color w:val="FF0000"/><w:sz w:val="20"/><w:lang w:val="en-US"/>') > 0);
  assert.ok(out.after.indexOf('<w:pgSz w:w="100" w:h="200"/><w:pgMar') > 0);
});

test('logical edges (start/end) are written as the file had them', async () => {
  const out = await exact(
    '<w:p><w:pPr><w:ind w:start="720" w:startChars="10" w:end="360" w:endChars="5" w:firstLine="200" w:firstLineChars="20"/></w:pPr></w:p>' +
      '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single"/><w:start w:val="double"/><w:bottom w:val="single"/><w:end w:val="dotted"/></w:tblBorders>' +
      '<w:tblCellMar><w:top w:w="10" w:type="dxa"/><w:start w:w="20" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:end w:w="40" w:type="dxa"/></w:tblCellMar></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcBorders><w:start w:val="single"/><w:end w:val="single"/></w:tcBorders>' +
      '<w:tcMar><w:start w:w="5" w:type="dxa"/><w:end w:w="6" w:type="dxa"/></w:tcMar></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
  );
  assert.ok(out.after.indexOf('w:start="720"') > 0 && out.after.indexOf('w:left=') < 0);
});

test('kept children and partly understood children of property elements survive', async () => {
  const out = await exact(
    // An extension element, an unknown attribute, a value out of range, a duplicate, a comment in the element.
    '<w:p><w:pPr w:custom="1"><w:pStyle w:val="A"/><w:spacing w:before="120" w:mystery="7"/><w:ind w:left="oops" w:right="20"/><w:jc w:val="center"/><w:jc w:val="left"/><w14:extra w14:val="1"/></w:pPr>' +
      '<w:r><w:rPr><w:b/><w:sz w:val="24" w:extra="x"/><w14:ligatures w14:val="standard"/><w:lang w:val="tr-TR"/><w14:cntxtAlts/></w:rPr><w:t>x</w:t></w:r>' +
      '<w:r><w:rPr><!--why--><w:i/></w:rPr><w:t>y</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct" w:odd="1"/><w:tblBorders w:x="1"><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:top w:val="double"/></w:tblBorders><w15:thing/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="100"/><w:tblGridChange w:id="1"><w:tblGrid><w:gridCol w:w="50"/></w:tblGrid></w:tblGridChange></w:tblGrid>' +
      '<w:tr><w:trPr><w:divId w:val="5"/><w:trHeight w:val="300" w:hRule="exact" w:z="1"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/><w:headers><w:header w:val="h"/></w:headers><w:cellMerge w:id="1" w:author="A"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>',
  );
  for (const piece of ['w:custom="1"', 'w:mystery="7"', 'w:left="oops"', '<w14:ligatures w14:val="standard"/>', '<!--why-->', 'w:odd="1"', '<w:tblGridChange', 'w:z="1"', '<w:cellMerge']) {
    assert.ok(out.after.indexOf(piece) > 0, piece);
  }
});

test('a partly understood child takes the edits of the model and keeps what the model does not hold', async () => {
  const bytes = await buildSpec({ body: '<w:p><w:pPr><w:spacing w:before="120" w:after="60" w:mystery="7"/><w:ind w:left="oops" w:right="20"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>' });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph');
  // An edit: a new block with new properties; the kept children stay with the bag.
  doc.body.blocks[0] = {
    type: 'paragraph',
    props: { ...paragraph.props, spacing: { before: 240 }, indent: { right: 20, hanging: 5 } },
    children: paragraph.children,
  };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  // "after" is gone with the model, "before" changed, the unknown attribute stays; the unread left indent stays too.
  assert.ok(xml.indexOf('<w:spacing w:before="240" w:mystery="7"/>') > 0, xml);
  assert.ok(xml.indexOf('<w:ind w:left="oops" w:right="20" w:hanging="5"/>') > 0, xml);
});

test('an edge added to a partly understood border container takes its place in the order of the schema', async () => {
  const bytes = await buildSpec({ body: '<w:p><w:pPr><w:pBdr w:odd="1"><w:top w:val="single"/><w:bottom w:val="single"/><w:mystery/></w:pBdr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>' });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph' && paragraph.props.borders);
  doc.body.blocks[0] = {
    type: 'paragraph',
    props: { ...paragraph.props, borders: { ...paragraph.props.borders, left: { style: 'double' }, right: { style: 'dotted' }, bottom: undefined } },
    children: paragraph.children,
  };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:pBdr w:odd="1"><w:top w:val="single"/><w:left w:val="double"/><w:right w:val="dotted"/><w:mystery/></w:pBdr>') > 0, xml);
  assert.deepEqual(schemaOrderViolations(xml), []);
});

test('run content: every kind of item', async () => {
  await exact(
    '<w:p><w:r><w:t>plain</w:t><w:t xml:space="preserve"> lead </w:t><w:t> trimmed </w:t><w:t/><w:tab/><w:br/><w:br w:type="page"/><w:br w:type="column"/>' +
      '<w:br w:type="textWrapping" w:clear="all"/><w:cr/><w:noBreakHyphen/><w:softHyphen/><w:sym w:font="Wingdings" w:char="F0FC"/><w:sym w:char="00B7"/>' +
      '<w:footnoteReference w:id="2"/><w:footnoteReference w:customMarkFollows="1" w:id="3"/><w:endnoteReference w:id="1"/><w:commentReference w:id="0"/>' +
      '<w:footnoteRef/><w:endnoteRef/><w:annotationRef/><w:separator/><w:continuationSeparator/><w:pgNum/><w:lastRenderedPageBreak/>' +
      '<w:ptab w:relativeTo="margin" w:alignment="center" w:leader="dot"/><w:delText xml:space="preserve"> gone</w:delText>' +
      '<w:instrText xml:space="preserve"> PAGE </w:instrText><w:delInstrText>X</w:delInstrText><!--c--><?pi data?><w:unknownThing w:a="1"><w:x/></w:unknownThing></w:r></w:p>',
  );
});

test('text: outer white space, entities and the text as written', async () => {
  const out = await exact('<w:p><w:r><w:t>a &amp; b &lt; c &gt; d "q" \'s\'</w:t><w:t xml:space="preserve">\ttab\n</w:t><w:t>  two  spaces  </w:t><w:t>]]&gt;</w:t></w:r></w:p>');
  const paragraph = out.result.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph' && paragraph.children[0].type === 'run');
  const items = paragraph.children[0].content;
  assert.deepEqual(items[2], { type: 'text', text: 'two  spaces', source: '  two  spaces  ' });
  assert.deepEqual(items[3], { type: 'text', text: ']]>' });
});

test('complex fields: characters, form data and what the element says beyond the model', async () => {
  await exact(
    '<w:p><w:r><w:fldChar w:fldCharType="begin" w:dirty="1" w:fldLock="0"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin" w:other="x" w:dirty="maybe"><w:ffData><w:name w:val="Check1"/></w:ffData></w:fldChar></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"><w:fldData xml:space="preserve">QUJD</w:fldData></w:fldChar></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  );
});

const INLINE_WRAPPERS =
  '<w:p><w:hyperlink r:id="rIdL1" w:anchor="top" w:tooltip="tip" w:history="1"><w:r><w:t>link</w:t></w:r></w:hyperlink>' +
  '<w:fldSimple w:instr=" PAGE " w:fldLock="1" w:dirty="0"><w:fldData xml:space="preserve">AAAA</w:fldData><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
  '<w:ins w:id="1" w:author="A" w:date="2020-01-01T00:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins><w:del w:id="2" w:author="A"><w:r><w:delText>old</w:delText></w:r></w:del>' +
  '<w:moveFrom w:id="3" w:author="B"><w:r><w:t>m</w:t></w:r></w:moveFrom><w:moveTo w:id="4" w:author="B"><w:r><w:t>m</w:t></w:r></w:moveTo>' +
  '<w:sdt><w:sdtPr><w:alias w:val="Name"/><w:id w:val="5"/></w:sdtPr><w:sdtEndPr><w:rPr><w:b/></w:rPr></w:sdtEndPr><w:sdtContent><w:r><w:t>in</w:t></w:r><!--inside--></w:sdtContent></w:sdt>' +
  '<w:smartTag w:uri="urn:x" w:element="place"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr><w:r><w:t>Paris</w:t></w:r></w:smartTag>' +
  '<w:customXml w:uri="urn:y" w:element="e"><w:customXmlPr><w:placeholder w:val="p"/></w:customXmlPr><w:r><w:t>c</w:t></w:r></w:customXml>' +
  '<w:dir w:val="rtl"><w:r><w:t>d</w:t></w:r></w:dir><w:bdo w:val="ltr"><w:r><w:t>b</w:t></w:r></w:bdo>' +
  '<w:hyperlink w:anchor="a"><w:ins w:id="9" w:author="C"><w:r><w:t>nested</w:t></w:r><w:bookmarkStart w:id="7" w:name="bm"/></w:ins><w:r><w:t>after</w:t></w:r></w:hyperlink>' +
  '<w:hyperlink w:anchor="one"><w:r><w:t>1</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="two"><w:r><w:t>2</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="empty"/>' +
  '<w:sdt><w:sdtContent/></w:sdt><w:bookmarkEnd w:id="7"/><w:commentRangeStart w:id="0"/><w:commentRangeEnd w:id="0"/><w:proofErr w:type="spellStart"/>' +
  '<w:permStart w:id="abc" w:edGrp="everyone"/><w:permEnd w:id="abc"/><w:moveFromRangeStart w:id="11" w:name="mv" w:author="A"/><w:moveFromRangeEnd w:id="11"/>' +
  '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:mystery w:a="1"/></w:p>';

test('inline wrappers, range marks and math', async () => {
  const out = await exact(INLINE_WRAPPERS, { rels: [['rIdL1', 'hyperlink', 'https://example.com/', true]] });
  // Two adjacent links stay two elements.
  assert.ok(out.after.indexOf('<w:hyperlink w:anchor="one"><w:r><w:t>1</w:t></w:r></w:hyperlink><w:hyperlink w:anchor="two">') > 0);
});

test('alternate content: kept whole while its branch is unchanged, at every level', async () => {
  const out = await exact(
    '<w:p><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent>' +
      '<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>a</w:t></mc:Choice><mc:Fallback><w:t>b</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>' +
      '<mc:AlternateContent><mc:Choice Requires="unknownns"><w:p><w:r><w:t>no</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>yes</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent>' +
      '<mc:AlternateContent><mc:Choice Requires="nothing"><w:p/></mc:Choice></mc:AlternateContent>' +
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="10"/></w:tblGrid><mc:AlternateContent><mc:Choice Requires="w14"><w:tr><w:tc><w:p/></w:tc></w:tr></mc:Choice><mc:Fallback><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></mc:Fallback></mc:AlternateContent>' +
      '<w:tr><mc:AlternateContent><mc:Choice Requires="w14"><w:tc><w:p/></w:tc></mc:Choice><mc:Fallback><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></mc:Fallback></mc:AlternateContent></w:tr></w:tbl>',
  );
  assert.ok(out.after.indexOf('fallback') > 0);
});

test('alternate content whose branch was edited is written as that branch alone', async () => {
  const bytes = await buildSpec({
    body:
      '<w:p><mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t>choice</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>fallback</w:t></w:r></mc:Fallback></mc:AlternateContent>' +
      '<w:r><mc:AlternateContent><mc:Choice Requires="w14"><w:t>a</w:t></mc:Choice><mc:Fallback><w:t>b</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>',
  });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph');
  const first = paragraph.children[0];
  const second = paragraph.children[1];
  assert.ok(first.type === 'run' && second.type === 'run' && second.content[0].type === 'alternate');
  doc.body.blocks[0] = {
    type: 'paragraph',
    props: paragraph.props,
    children: [
      { ...first, content: [{ type: 'text', text: 'edited' }] },
      { ...second, content: [{ ...second.content[0], items: [{ type: 'text', text: 'z' }] }] },
    ],
  };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:body><w:p><w:r><w:t>edited</w:t></w:r><w:r><w:t>z</w:t></w:r></w:p></w:body>') > 0, xml);
});

test('block-level wrappers, marks, comments and unknown elements between blocks', async () => {
  await exact(
    '<!--top--><w:bookmarkStart w:id="1" w:name="a"/><w:p><w:r><w:t>1</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>' +
      '<w:sdt w:odd="1"><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr><!--loose--><w:sdtEndPr/><w:sdtContent><w:p><w:r><w:t>toc</w:t></w:r></w:p>' +
      '<w:sdt><w:sdtContent><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/></w:sdtContent><w:extra/></w:sdt>' +
      '<w:customXml w:uri="u" w:element="e"><w:customXmlPr><w:attr w:name="n" w:val="v"/></w:customXmlPr><w:p/><w:tbl><w:tblPr/><w:tblGrid/><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:customXml>' +
      '<w:sdt><w:sdtContent/></w:sdt><w:altChunk r:id="rIdA"/><?break here?><w:p/>',
  );
});

const TBLPR =
  '<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblpPr w:leftFromText="141" w:rightFromText="142" w:topFromText="1" w:bottomFromText="2" w:vertAnchor="page" w:horzAnchor="margin" w:tblpXSpec="center" w:tblpX="10" w:tblpYSpec="top" w:tblpY="20"/>' +
  '<w:tblOverlap w:val="never"/><w:bidiVisual/><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="2"/><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>' +
  '<w:tblCellSpacing w:w="15" w:type="dxa"/><w:tblInd w:w="-108" w:type="dxa"/>' +
  '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="none"/><w:insideV w:val="nil"/></w:tblBorders>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/><w:tblLayout w:type="fixed"/>' +
  '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>' +
  '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/><w:tblCaption w:val="Cap"/><w:tblDescription w:val="Desc"/>' +
  '<w:tblPrChange w:id="3" w:author="A"><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr></w:tblPrChange></w:tblPr>';

const TABLE =
  `<w:tbl>${TBLPR}<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  '<w:tr w:rsidR="00112233"><w:tblPrEx><w:tblW w:w="100" w:type="dxa"/><w:jc w:val="left"/><w:tblBorders><w:top w:val="double"/></w:tblBorders><w:tblCellMar><w:left w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPrEx>' +
  '<w:trPr><w:cnfStyle w:val="100000000000"/><w:gridBefore w:val="1"/><w:gridAfter w:val="1"/><w:wBefore w:w="10" w:type="dxa"/><w:wAfter w:w="20" w:type="dxa"/><w:cantSplit/>' +
  '<w:trHeight w:val="400" w:hRule="atLeast"/><w:tblHeader/><w:tblCellSpacing w:w="5" w:type="dxa"/><w:jc w:val="center"/><w:hidden w:val="0"/>' +
  '<w:ins w:id="20" w:author="A"/><w:trPrChange w:id="21" w:author="A"><w:trPr><w:cantSplit w:val="0"/></w:trPr></w:trPrChange></w:trPr>' +
  '<w:tc><w:tcPr><w:cnfStyle w:val="001000000000"/><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/><w:hMerge w:val="restart"/><w:vMerge w:val="restart"/>' +
  '<w:tcBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000"/><w:left w:val="nil"/><w:bottom w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/><w:tl2br w:val="single"/><w:tr2bl w:val="dashed"/></w:tcBorders>' +
  '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/><w:noWrap/><w:tcMar><w:top w:w="10" w:type="dxa"/><w:left w:w="20" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tcMar>' +
  '<w:textDirection w:val="btLr"/><w:tcFitText/><w:vAlign w:val="center"/><w:hideMark/><w:tcPrChange w:id="22" w:author="A"><w:tcPr><w:tcW w:w="1" w:type="dxa"/></w:tcPr></w:tcPrChange></w:tcPr>' +
  '<w:p><w:r><w:t>a</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="500"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/></w:tc>' +
  '<w:tc w:id="c2"><w:tcPr><w:tcW w:w="0" w:type="auto"/><w:hMerge/><w:vMerge/><w:vAlign w:val="bottom"/></w:tcPr><!--cell--><w:p/></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="50" w:type="nil"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>';

test('tables: properties of the table, its rows and cells, exceptions, nesting', async () => {
  await exact(TABLE);
});

test('tables: wrappers around rows and cells, kept children and comments in their places', async () => {
  await exact(
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid><!--before rows--><w:bookmarkStart w:id="1" w:name="t"/>' +
      '<w:tr><w:tc><w:p/></w:tc><w:bookmarkEnd w:id="1"/><!--between cells--><w:tc><w:p/></w:tc></w:tr>' +
      '<w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:tr><w:tc><w:p/></w:tc><w:sdt><w:sdtPr><w:id w:val="10"/></w:sdtPr><w:sdtContent><w:tc><w:p><w:r><w:t>in sdt</w:t></w:r></w:p></w:tc><w:permStart w:id="p"/></w:sdtContent></w:sdt></w:tr>' +
      '<w:bookmarkStart w:id="2" w:name="u"/><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr><!--in row sdt--></w:sdtContent></w:sdt>' +
      '<w:customXml w:element="row"><w:tr><w:customXml w:element="cell"><w:tc><w:p/></w:tc></w:customXml><w:tc><w:p/></w:tc></w:tr></w:customXml>' +
      '<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr><w:bookmarkEnd w:id="2"/></w:tbl>',
  );
});

const SECTPR =
  '<w:sectPr w:rsidR="00AA11BB" w:rsidSect="00CC22DD"><w:headerReference w:type="even" r:id="rIdH2"/><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/><w:footerReference w:type="first" r:id="rIdF1"/>' +
  '<w:footnotePr><w:pos w:val="beneathText"/><w:numFmt w:val="lowerRoman"/><w:numStart w:val="3"/><w:numRestart w:val="eachPage"/></w:footnotePr><w:endnotePr><w:pos w:val="sectEnd"/><w:numFmt w:val="decimal"/></w:endnotePr>' +
  '<w:type w:val="continuous"/><w:pgSz w:w="16838" w:h="11906" w:orient="landscape" w:code="9"/><w:pgMar w:top="-1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>' +
  '<w:paperSrc w:first="1"/><w:pgBorders w:zOrder="back" w:display="firstPage" w:offsetFrom="page"><w:top w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="24" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="24" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="24" w:color="auto"/></w:pgBorders>' +
  '<w:lnNumType w:countBy="5" w:start="1" w:distance="100" w:restart="newSection"/><w:pgNumType w:fmt="upperRoman" w:start="4" w:chapStyle="1" w:chapSep="hyphen"/>' +
  '<w:cols w:num="2" w:sep="1" w:space="708" w:equalWidth="0"><w:col w:w="4000" w:space="500"/><w:col w:w="5000"/></w:cols><w:formProt w:val="0"/><w:vAlign w:val="center"/><w:noEndnote/><w:titlePg/>' +
  '<w:textDirection w:val="tbRl"/><w:bidi/><w:rtlGutter/><w:docGrid w:type="lines" w:linePitch="360" w:charSpace="-2049"/><w:printerSettings r:id="rIdP"/>' +
  '<w:sectPrChange w:id="30" w:author="A"><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:sectPrChange></w:sectPr>';

test('section properties: in a paragraph and at the end of the body', async () => {
  const spec: DocxSpec = { headers: { rIdH1: '<w:p><w:r><w:t>head</w:t></w:r></w:p>', rIdH2: '<w:p/>' }, footers: { rIdF1: '<w:p><w:r><w:pgNum/></w:r></w:p>' } };
  const out = await exact(`<w:p><w:pPr>${SECTPR}</w:pPr></w:p><w:p/>${SECTPR}`, spec);
  assert.equal(out.result.body.section?.margins?.top, -1417);
});

test('a section mark that is not the last child of the body stays where it was', async () => {
  const out = await exact('<w:p><w:r><w:t>1</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="100" w:h="200"/></w:sectPr><w:p><w:r><w:t>2</w:t></w:r></w:p>');
  assert.equal(out.result.body.sectionIndex, 1);
});

const PICTURE =
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="p.png" descr="alt"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdI1"/><a:srcRect l="1000" t="2000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>';

const INLINE_DRAWING =
  '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="1" b="2"/><wp:docPr id="1" name="Picture 1" descr="alt"/>' +
  `<wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">${PICTURE}</a:graphicData></a:graphic></wp:inline></w:drawing>`;

const TEXT_BOX = (inner: string) =>
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>100</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:align>top</wp:align></wp:positionV>' +
  '<wp:extent cx="1828800" cy="914400"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="2" name="Text Box 2"/><wp:cNvGraphicFramePr/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>' +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);

const DRAWING_SPEC: DocxSpec = { rels: [['rIdI1', 'image', 'media/image1.png']], parts: [['word/media/image1.png', PNG]] };

test('drawings, VML pictures and objects go back as the trees they are', async () => {
  await exact(
    `<w:p><w:r>${INLINE_DRAWING}</w:r><w:r>${TEXT_BOX('<w:p><w:r><w:t>boxed</w:t></w:r></w:p>')}</w:r>` +
      '<w:r><w:pict><v:shape id="s1" style="width:100pt;height:50pt"><v:imagedata r:id="rIdI1"/><v:textbox><w:txbxContent><w:p><w:r><w:t>vml</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>' +
      '<w:r><w:object w:dxaOrig="100" w:dyaOrig="100"><v:shape id="s2"/><o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ProgID="Excel.Sheet.12" ShapeID="s2" DrawAspect="Content" ObjectID="_1" r:id="rIdO"/></w:object></w:r></w:p>',
    DRAWING_SPEC,
  );
});

test('an edited text box is written into the kept tree of its drawing; an edited size too', async () => {
  const bytes = await buildSpec({ ...DRAWING_SPEC, body: `<w:p><w:r>${TEXT_BOX('<w:p><w:r><w:t>boxed</w:t></w:r></w:p>')}</w:r><w:r>${INLINE_DRAWING}</w:r></w:p>` });
  const pkg = await OpcPackage.open(bytes);
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph');
  const [first, second] = paragraph.children;
  assert.ok(first.type === 'run' && second.type === 'run');
  const box = first.content[0];
  const picture = second.content[0];
  assert.ok(box.type === 'drawing' && picture.type === 'drawing' && box.textBoxes);
  doc.body.blocks[0] = {
    type: 'paragraph',
    props: paragraph.props,
    children: [
      { ...first, content: [{ ...box, textBoxes: [[{ type: 'paragraph', props: {}, children: [{ type: 'run', props: { bold: true }, content: [{ type: 'text', text: 'changed' }] }] }]] }] },
      { ...second, content: [{ ...picture, width: 1000, height: 500 }] },
    ],
  };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>changed</w:t></w:r></w:p></w:txbxContent>') > 0, xml);
  assert.ok(xml.indexOf('boxed') < 0);
  // The rest of the shape is as it was.
  assert.ok(xml.indexOf('<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/>') > 0);
  assert.ok(xml.indexOf('<wp:extent cx="1000" cy="500"/>') > 0 && xml.indexOf('<a:ext cx="1000" cy="500"/>') > 0, xml);
  const again = await readDocx(await OpcPackage.open(await pkg.save()));
  const block = again.body.blocks[0];
  assert.ok(block.type === 'paragraph' && block.children[1].type === 'run');
  const read = block.children[1].content[0];
  assert.ok(read.type === 'drawing');
  assert.deepEqual([read.width, read.height, read.image?.target], [1000, 500, '/word/media/image1.png']);
});

test('headers, footers, notes and comments are written with the same writer', async () => {
  const spec: DocxSpec = {
    body: '<w:p><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:commentReference w:id="4"/></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/></w:sectPr>',
    headers: { rIdH1: `<!--h--><w:p>${PPR}<w:r>${RPR}<w:t>head</w:t></w:r></w:p>${TABLE}` },
    footers: { rIdF1: '<w:sdt><w:sdtPr><w:id w:val="1"/></w:sdtPr><w:sdtContent><w:p><w:r><w:pgNum/></w:r></w:p></w:sdtContent></w:sdt>' },
    footnotes:
      '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><!--between--><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      '<w:footnote w:id="1" w:custom="x"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> note</w:t></w:r></w:p><w:p/></w:footnote><w:footnote w:id="1"><w:p/></w:footnote><w:other/>',
    endnotes: '<w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r></w:p></w:endnote>',
    comments: '<w:comment w:id="4" w:author="Someone" w:date="2021-02-03T04:05:06Z" w:initials="S" w:x="1"><w:p><w:r><w:annotationRef/></w:r><w:r><w:t>hm</w:t></w:r></w:p></w:comment>',
  };
  const out = await regenerate(spec);
  for (const part of ['header1.xml', 'footer1.xml', 'footnotes.xml', 'endnotes.xml', 'comments.xml']) {
    const original = await (await OpcPackage.open(await buildSpec(spec))).readPart('/word/' + part);
    const written = await out.pkg.readPart('/word/' + part);
    assert.equal(canonicalText(written), canonicalText(original), part);
    assert.deepEqual(schemaOrderViolations(written), [], part);
  }
});

test('strict documents are written with the names and spellings of the strict vocabulary', async () => {
  const out = await regenerate({
    strict: true,
    body:
      '<w:p><w:pPr><w:cnfStyle w:firstRow="1" w:lastRowLastColumn="true"/><w:ind w:start="2.5cm" w:end="12pt"/><w:jc w:val="start"/></w:pPr><w:r><w:rPr><w:b w:val="true"/><w:sz w:val="12pt"/></w:rPr><w:t>x</w:t></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="50%"/></w:tblPr><w:tblGrid><w:gridCol w:w="1in"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl><w:sectPr><w:pgSz w:w="210mm" w:h="297mm"/></w:sectPr>',
  });
  assert.equal(out.result.strict, true);
  assert.ok(out.after.indexOf('xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"') > 0);
  // The flags of a strict cnfStyle are attributes: it has no w:val.
  assert.ok(out.after.indexOf('<w:cnfStyle w:firstRow="1" w:lastRowLastColumn="1"/>') > 0, out.after);
  assert.ok(out.after.indexOf('w:start="1417"') > 0 && out.after.indexOf('<w:tblW w:w="2500" w:type="pct"/>') > 0);
});

test('a document whose WordprocessingML prefix is not "w" is written with its own prefix', async () => {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const document =
    `<?xml version="1.0" encoding="UTF-8"?><x:document xmlns:x="${ns}" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    '<x:body><x:p><x:pPr><x:jc x:val="center"/></x:pPr><x:hyperlink rel:id="rIdL1"><x:r><x:rPr><x:b/></x:rPr><x:t xml:space="preserve"> a </x:t></x:r></x:hyperlink></x:p>' +
    '<x:tbl><x:tblPr><x:tblW x:w="0" x:type="auto"/></x:tblPr><x:tblGrid><x:gridCol x:w="100"/></x:tblGrid><x:tr><x:tc><x:p/></x:tc></x:tr></x:tbl><x:sectPr><x:pgSz x:w="1" x:h="2"/></x:sectPr></x:body></x:document>';
  const out = await regenerate({ document, rels: [['rIdL1', 'hyperlink', 'https://example.com/', true]] });
  assert.equal(canonicalText(out.after), canonicalText(document));
  assert.ok(out.after.indexOf('<w:') < 0);
});

test('a document with WordprocessingML as its default namespace is written without prefixes', async () => {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const document = `<document xmlns="${ns}"><body><p><pPr><jc val="center"/></pPr><r><rPr><b/></rPr><t>a</t></r></p><sectPr><pgSz w="1" h="2"/></sectPr></body></document>`;
  const out = await regenerate({ document });
  assert.equal(canonicalText(out.after), canonicalText(document));
});

test('what the document element holds besides the body, and what stands around it, goes back', async () => {
  const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const document =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!--lead--><w:document xmlns:w="${ns}" w:conformance="transitional"><!--a--><w:background w:color="FFFFFF"/><?pi x?>` +
    '<w:body><w:p/></w:body><!--b--><w:extra/></w:document><!--tail-->';
  const out = await regenerate({ document });
  assert.equal(out.after, document);
});

test('deliberate normalizations: spellings the model does not keep', async () => {
  const out = await regenerate({
    body:
      '<w:p><w:pPr><w:keepNext w:val="true"/><w:keepLines w:val="off"/><w:pBdr><w:bottom w:val="single" w:color="ff00aa"/><w:top w:val="single"/></w:pBdr><w:cnfStyle w:val="100000000000" w:firstRow="1"/></w:pPr>' +
      '<w:r><w:rPr/><w:t xml:space="preserve">no outer space</w:t><w:br w:type="textWrapping"/><w:sym w:font="Symbol" w:char="f0b7"/></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:type="auto"/></w:tblPr><w:tr><w:trPr/><w:tc><w:tcPr><w:vMerge w:val="continue"/><w:tcW w:w="100"/></w:tcPr><w:p><w:pPr/></w:p></w:tc></w:tr></w:tbl>',
  });
  const body = out.after.slice(out.after.indexOf('<w:body>'));
  assert.equal(
    body,
    '<w:body><w:p><w:pPr><w:keepNext/><w:keepLines w:val="0"/><w:pBdr><w:top w:val="single"/><w:bottom w:val="single" w:color="FF00AA"/></w:pBdr><w:cnfStyle w:val="100000000000"/></w:pPr>' +
      '<w:r><w:t>no outer space</w:t><w:br/><w:sym w:font="Symbol" w:char="F0B7"/></w:r></w:p>' +
      '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid/><w:tr><w:tc><w:tcPr><w:tcW w:w="100" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>',
  );
});

test('a property element kept for its attributes and comments: whole while unchanged, its attributes after an edit', async () => {
  const body = '<w:p><w:pPr w:custom="1"><!--why--><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>';
  // Unchanged: the element as it was, in the order it had and with its comment.
  const unchanged = await regenerate({ body });
  assert.ok(unchanged.after.indexOf('<w:pPr w:custom="1"><!--why--><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr>') > 0, unchanged.after);

  const pkg = await OpcPackage.open(await buildSpec({ body }));
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph');
  doc.body.blocks[0] = { type: 'paragraph', props: { ...paragraph.props, align: 'right' }, children: paragraph.children };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  // Written from the model, in the order of the schema, with the attribute; the comment had no place in the model.
  assert.ok(xml.indexOf('<w:pPr w:custom="1"><w:spacing w:after="0"/><w:jc w:val="right"/></w:pPr>') > 0, xml);
});

test('a table grid kept for its change record takes the columns of the model', async () => {
  const body =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="200"/><w:tblGridChange w:id="1"><w:tblGrid><w:gridCol w:w="50"/></w:tblGrid></w:tblGridChange></w:tblGrid><w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>';
  const pkg = await OpcPackage.open(await buildSpec({ body }));
  const doc = await readDocx(pkg);
  const table = doc.body.blocks[0];
  assert.ok(table.type === 'table');
  doc.body.blocks[0] = { ...table, grid: [150, 250, 300], srcStart: undefined, srcEnd: undefined };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<w:tblGrid><w:gridCol w:w="150"/><w:gridCol w:w="250"/><w:gridCol w:w="300"/><w:tblGridChange w:id="1"><w:tblGrid><w:gridCol w:w="50"/></w:tblGrid></w:tblGridChange></w:tblGrid>') > 0, xml);
});

test('what a table keeps inside a wrapper that holds no row is written with its wrapper', async () => {
  const out = await regenerate({
    body:
      '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:sdt><w:sdtPr><w:id w:val="3"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="1" w:name="x"/><!--note--></w:sdtContent></w:sdt>' +
      '<w:tr><w:tc><w:p/></w:tc><w:customXml w:element="e"><w:bookmarkEnd w:id="1"/></w:customXml></w:tr></w:tbl>',
  });
  // The model does not say where such a wrapper stood among the rows: it goes to the end of the table (of the row).
  assert.ok(out.after.indexOf('<w:sdt><w:sdtPr><w:id w:val="3"/></w:sdtPr><w:sdtContent><w:bookmarkStart w:id="1" w:name="x"/><!--note--></w:sdtContent></w:sdt></w:tbl>') > 0, out.after);
  assert.ok(out.after.indexOf('<w:customXml w:element="e"><w:bookmarkEnd w:id="1"/></w:customXml></w:tr>') > 0, out.after);
});

test('a VML text box that was edited is written into the kept picture', async () => {
  const body =
    '<w:p><w:r><w:pict><v:shape id="s1" style="width:100pt;height:50pt"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>vml</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>';
  const pkg = await OpcPackage.open(await buildSpec({ body }));
  const doc = await readDocx(pkg);
  const paragraph = doc.body.blocks[0];
  assert.ok(paragraph.type === 'paragraph' && paragraph.children[0].type === 'run');
  const run = paragraph.children[0];
  const pict = run.content[0];
  assert.ok(pict.type === 'pict');
  doc.body.blocks[0] = {
    type: 'paragraph',
    props: paragraph.props,
    children: [{ ...run, content: [{ ...pict, textBoxes: [[{ type: 'paragraph', props: {}, children: [{ type: 'run', props: {}, content: [{ type: 'text', text: 'new text' }] }] }]] }] }],
  };
  doc.body.modified = true;
  await writeDocx(pkg, doc);
  const xml = decoder.decode(await pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>new text</w:t></w:r></w:p></w:txbxContent></v:textbox>') > 0, xml);
});
