// The copy of a part of a document (fragmentOf), the paste of it
// (insertFragment), in the document it came from and in another one, and the
// copies the editors of a page keep (doc-clipboard).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { readDocx } from '../src/formats/docx/document';
import { readRtf } from '../src/formats/rtf/document';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type Block, type DocModel, type DrawingItem, type Paragraph, type Run, type Table } from '../src/model/doc/types';
import { OpcPackage } from '../src/opc';
import {
  DocHistory,
  MAX_TABLE_DEPTH,
  applyDocOps,
  deleteRange,
  expandRangeForTables,
  fragmentNeeds,
  fragmentOf,
  insertFragment,
  insertImage,
  orderRange,
  toggleList,
  type DocFragment,
  type DocPosition,
  type DocRange,
} from '../src/ops/doc/index';
import {
  MAX_STORED_BYTES,
  MAX_STORED_COPIES,
  clearStoredCopies,
  copyMarkOf,
  cutLoses,
  setCopyMedia,
  sliceHtml,
  sliceOf,
  storeCopy,
  storedCopies,
  storedCopy,
} from '../src/ui/doc-clipboard';
import { buildDocx, type DocxSpec } from './docx-fixtures';

const STYLES =
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Heading1"/><w:rPr><w:i/></w:rPr></w:style>' +
  '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>';
const PLAIN_STYLES = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const NUMBERING =
  '<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>' +
  '<w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>';
const SECT = '<w:sectPr><w:pgSz w:w="9000" w:h="9000"/></w:sectPr>';

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 2, 0, 0, 0, 1, 8, 2, 0, 0, 0]);
const PNG2 = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0, 3, 0, 0, 0, 3, 8, 2, 0, 0, 0]);

const esc = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;');
const r = (text: string, rPr = '') => `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const p = (content = '', pPr = '') => `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${content.charAt(0) === '<' ? content : content ? r(content) : ''}</w:p>`;
const tc = (content: string, tcPr = '') => `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${content.charAt(0) === '<' ? content : p(content)}</w:tc>`;
const tr = (cells: string) => `<w:tr>${cells}</w:tr>`;
const tbl = (rows: string, grid: number[]) => `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${grid.map((w) => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
const picture = (rId: string, id = 1, frame = 'inline') =>
  `<w:r><w:drawing><wp:${frame}${frame === 'anchor' ? ' simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV' : ''}>` +
  `<wp:extent cx="381000" cy="152400"/>${frame === 'anchor' ? '<wp:wrapSquare wrapText="bothSides"/>' : ''}<wp:docPr id="${id}" name="P${id}"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="x"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${rId}"/></pic:blipFill><pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:${frame}></w:drawing></w:r>`;
const FIELD = (code: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">' + code + '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r(result) + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

interface Opened {
  pkg: OpcPackage;
  doc: DocModel;
}

async function open(spec: DocxSpec | string): Promise<Opened> {
  const full: DocxSpec = typeof spec === 'string' ? { body: spec } : { ...spec };
  if (full.styles === undefined) full.styles = STYLES;
  if (full.body !== undefined && full.body.indexOf('<w:sectPr') < 0) full.body += SECT;
  const pkg = await OpcPackage.open(await buildDocx(full));
  return { pkg, doc: await readDocx(pkg) };
}

const WITH_PICTURE: Partial<DocxSpec> = { rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', PNG]] };

const at = (path: number[] | number, offset: number, before?: boolean): DocPosition =>
  before ? { story: 'body', path: typeof path === 'number' ? [path] : path, offset, before: true } : { story: 'body', path: typeof path === 'number' ? [path] : path, offset };
const range = (a: DocPosition, b: DocPosition): DocRange => ({ anchor: a, focus: b });

/** The content of blocks as text: runs item by item, wrappers and field code marked, tables with their rows and cells. */
function shape(blocks: Block[]): string {
  const out: string[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') {
      let text = '';
      for (const inline of block.children) {
        if (inline.type !== 'run') {
          text += `{${inline.type}}`;
          continue;
        }
        let wrap = '';
        for (let w = inline.wrap; w; w = w.parent) wrap = w.kind + ':' + wrap;
        inline.content.forEach((item, k) => {
          const code = inline.fieldCode === true || (inline.fieldCode !== undefined && inline.fieldCode[k]);
          let piece = item.type === 'text' || item.type === 'instrText' ? item.text : item.type === 'break' ? `<break ${item.kind}>` : item.type === 'fieldChar' ? `<${item.charType}>` : `<${item.type}>`;
          if (code) piece = `(${piece})`;
          text += wrap ? `[${wrap}${piece}]` : piece;
        });
      }
      out.push(text);
    } else if (block.type === 'table') {
      out.push('TABLE(' + block.rows.map((row) => row.cells.map((cell) => shape(cell.blocks)).join(' | ')).join(' / ') + ')');
    }
  }
  return out.join('¶');
}

const body = (doc: DocModel): string => shape(doc.body.blocks);

/** Copies a range the way the editor does, pastes it at a place, as one step of a history. */
function paste(doc: DocModel, history: DocHistory, fragment: DocFragment, where: DocPosition | DocRange, taken?: (part: string) => boolean): DocRange {
  const edit = insertFragment(doc, where, fragment, taken ? { taken } : undefined);
  history.applyEdit(edit, undefined);
  history.seal();
  return edit.selection as DocRange;
}

function cut(doc: DocModel, history: DocHistory, selection: DocRange): DocFragment {
  const fragment = fragmentOf(doc, expandRangeForTables(doc, selection));
  history.applyEdit(deleteRange(doc, selection), undefined);
  history.seal();
  return fragment;
}

const drawings = (doc: DocModel): DrawingItem[] => {
  const out: DrawingItem[] = [];
  const walk = (blocks: Block[]) => {
    for (const block of blocks) {
      if (block.type === 'paragraph') {
        for (const inline of block.children) if (inline.type === 'run') for (const item of inline.content) if (item.type === 'drawing') out.push(item);
      } else if (block.type === 'table') for (const row of block.rows) for (const cell of row.cells) walk(cell.blocks);
    }
  };
  walk(doc.body.blocks);
  return out;
};

const docPrId = (item: DrawingItem): string => {
  const frame = item.xml.children.find((node) => node.type === 'element') as { children: unknown[] };
  const docPr = (frame.children as { type: string; name: string; attrs: Map<string, string> }[]).find((node) => node.type === 'element' && /docPr$/.test(node.name))!;
  return docPr.attrs.get('id')!;
};

const embedOf = (item: DrawingItem): string => {
  let found = '';
  const walk = (el: { attrs: Map<string, string>; children: { type: string }[] }) => {
    const id = el.attrs.get('r:embed');
    if (id !== undefined) found = id;
    for (const node of el.children) if (node.type === 'element') walk(node as never);
  };
  walk(item.xml as never);
  return found;
};

async function saved(opened: Opened): Promise<Opened> {
  await writeDocx(opened.pkg, opened.doc);
  const pkg = await OpcPackage.open(await opened.pkg.save());
  return { pkg, doc: await readDocx(pkg) };
}

const decoder = new TextDecoder();

// ------------------------------------------------------------------ pictures

test('a picture copied and pasted in its document: the same part and relationship, a drawing id of its own, one undo step', async () => {
  const opened = await open({ body: p(r('ab') + picture('rIdImg', 5) + r('cd')) + p('next'), ...WITH_PICTURE });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const before = body(doc);
  const fragment = fragmentOf(doc, range(at(0, 1), at(0, 3)));
  assert.equal(shape(fragment.blocks), 'b<drawing>c');
  assert.deepEqual([fragment.openStart, fragment.openEnd], [true, true]);
  assert.deepEqual(Array.from(fragment.media.keys()), ['/word/media/image1.png']);
  // The bytes are in the package, not in the model; the document itself does not need them.
  assert.deepEqual(fragmentNeeds(doc, fragment), []);
  const relsBefore = doc.body.relationships;
  const caret = paste(doc, history, fragment, at(1, 4));
  assert.equal(body(doc), 'ab<drawing>cd¶nextb<drawing>c');
  assert.deepEqual(caret.focus, at(1, 6));
  const [first, second] = drawings(doc);
  assert.equal(embedOf(second), 'rIdImg');
  assert.equal(doc.body.relationships, relsBefore, 'no relationship was added');
  assert.equal(doc.media, undefined, 'no picture was added');
  assert.equal(docPrId(first), '5');
  assert.equal(docPrId(second), '6');
  assert.equal(second.image!.target, '/word/media/image1.png');
  // Twice more: every drawing has an id of its own.
  paste(doc, history, fragment, at(1, 0));
  assert.deepEqual(drawings(doc).map(docPrId).sort(), ['5', '6', '7']);
  history.undo();
  history.undo();
  assert.equal(body(doc), before);
  history.redo();
  assert.equal(body(doc), 'ab<drawing>cd¶nextb<drawing>c');

  const again = await saved(opened);
  assert.equal(body(again.doc), 'ab<drawing>cd¶nextb<drawing>c');
  const xml = decoder.decode(await again.pkg.readPart('/word/document.xml'));
  assert.equal((xml.match(/r:embed="rIdImg"/g) || []).length, 2);
  assert.deepEqual(
    again.pkg.partNames().filter((name) => name.indexOf('/media/') >= 0),
    ['/word/media/image1.png'],
  );
});

test('a cut and a paste are a move: a picture alone, a picture among text, a floating picture', async () => {
  const opened = await open({ body: p(r('ab') + picture('rIdImg', 5) + r('cd')) + p(r('x') + picture('rIdImg', 9, 'anchor') + r('y')) + p('end'), ...WITH_PICTURE });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const start = body(doc);
  // The picture alone, as a press on it selects it.
  let selection = range(at(0, 2, true), at(0, 2));
  assert.equal(cutLoses(doc, selection, sliceOf(doc, selection.anchor, selection.focus)), false);
  let fragment = cut(doc, history, selection);
  assert.equal(body(doc), 'abcd¶x<drawing>y¶end');
  paste(doc, history, fragment, at(0, 2));
  assert.equal(body(doc), start);
  // To another place.
  selection = range(at(0, 1), at(0, 3));
  fragment = cut(doc, history, selection);
  paste(doc, history, fragment, at(2, 3));
  assert.equal(body(doc), 'ad¶x<drawing>y¶endb<drawing>c');
  // The floating one with its paragraph end.
  selection = range(at(1, 0), at(2, 0));
  assert.equal(cutLoses(doc, selection, sliceOf(doc, selection.anchor, selection.focus)), false);
  fragment = cut(doc, history, selection);
  assert.equal(body(doc), 'ad¶endb<drawing>c');
  paste(doc, history, fragment, at(0, 0));
  assert.equal(body(doc), 'x<drawing>y¶ad¶endb<drawing>c');
  assert.equal(drawings(doc)[0].placement, 'anchor');
  const again = await saved(opened);
  assert.equal(body(again.doc), 'x<drawing>y¶ad¶endb<drawing>c');
  assert.deepEqual(
    drawings(again.doc).map((item) => [item.placement, item.image && item.image.target]),
    [
      ['anchor', '/word/media/image1.png'],
      ['inline', '/word/media/image1.png'],
    ],
  );
  // Every step taken back: the document as it was read.
  while (history.canUndo) history.undo();
  assert.equal(body(doc), start);
});

test('a picture pasted into another document brings its part, relationship and content type; undo takes them out again', async () => {
  const source = await open({ body: p(r('ab') + picture('rIdImg', 5) + r('cd')), ...WITH_PICTURE });
  const target = await open({ body: p('target') + p(picture('rId7', 3)), rels: [['rId7', 'image', 'media/image1.png']], parts: [['word/media/image1.png', PNG2]] });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 4)));
  // The bytes are still in the package of the source.
  assert.deepEqual(fragmentNeeds(target.doc, fragment), ['/word/media/image1.png']);
  assert.throws(() => insertFragment(target.doc, at(0, 6), fragment), RangeError);
  fragment.media.get('/word/media/image1.png')!.bytes = await source.pkg.readPart('/word/media/image1.png');
  assert.deepEqual(fragmentNeeds(target.doc, fragment), []);

  const history = new DocHistory(target.doc);
  const relsBefore = target.doc.body.relationships;
  paste(target.doc, history, fragment, at(0, 6), (part) => target.pkg.hasPart(part));
  assert.equal(body(target.doc), 'targetab<drawing>cd¶<drawing>');
  const pasted = drawings(target.doc)[0];
  // A name of its own (image1 is another picture here), a relationship of its own, an id no drawing has.
  assert.equal(pasted.image!.target, '/word/media/image2.png');
  assert.deepEqual(Array.from(target.doc.media!.keys()), ['/word/media/image2.png']);
  const rId = embedOf(pasted);
  assert.notEqual(rId, 'rId7');
  assert.equal(target.doc.body.relationships.get(rId)!.part, '/word/media/image2.png');
  assert.equal(pasted.image!.rId, rId);
  assert.equal(docPrId(pasted), '4');
  // A second paste of the same picture uses the part and the relationship of the first.
  paste(target.doc, history, fragment, at(0, 0));
  assert.deepEqual(Array.from(target.doc.media!.keys()), ['/word/media/image2.png']);
  assert.equal(embedOf(drawings(target.doc)[0]), rId);
  history.undo();

  const again = await saved(target);
  assert.equal(body(again.doc), 'targetab<drawing>cd¶<drawing>');
  assert.deepEqual(Array.from(await again.pkg.readPart('/word/media/image2.png')), Array.from(PNG));
  assert.deepEqual(Array.from(await again.pkg.readPart('/word/media/image1.png')), Array.from(PNG2));
  assert.equal(again.pkg.contentTypeOf('/word/media/image2.png'), 'image/png');
  assert.deepEqual(
    drawings(again.doc).map((item) => item.image && item.image.target),
    ['/word/media/image2.png', '/word/media/image1.png'],
  );
  assert.equal(new Set(drawings(again.doc).map(docPrId)).size, 2);

  history.undo();
  assert.equal(body(target.doc), 'target¶<drawing>');
  assert.equal(target.doc.media, undefined);
  assert.equal(target.doc.body.relationships, relsBefore);
  // The source is as it was.
  assert.equal(body(source.doc), 'ab<drawing>cd');
  assert.equal(source.doc.body.modified, undefined);
});

test('a picture put in while editing travels with its bytes; the copy does not follow later edits of its document', async () => {
  const source = await open(p('abcd') + p('more'));
  const history = new DocHistory(source.doc);
  history.applyEdit(insertImage(source.doc, at(0, 2), { bytes: PNG, type: 'image/png', width: 9525, height: 9525 }), undefined);
  history.seal();
  const fragment = fragmentOf(source.doc, range(at(0, 1), at(1, 2)));
  const kept = shape(fragment.blocks);
  assert.equal(kept, 'b<drawing>cd¶mo');
  assert.ok(fragment.media.get('/word/media/image1.png')!.bytes, 'the bytes are in the model and so in the copy');
  // The source changes, and is undone past the picture: the copy stays what it was, and can still be pasted there.
  history.applyEdit(deleteRange(source.doc, range(at(0, 0), at(1, 4))), undefined);
  assert.equal(shape(fragment.blocks), kept);
  while (history.canUndo) history.undo();
  assert.equal(source.doc.media, undefined);
  assert.equal(shape(fragment.blocks), kept);
  // Changing the blocks of the source in place (which no command does) does not reach the copy either.
  (source.doc.body.blocks[0] as Paragraph).children.length = 0;
  assert.equal(shape(fragment.blocks), kept);
  const target = await open(p('x'));
  paste(target.doc, new DocHistory(target.doc), fragment, at(0, 1));
  assert.equal(body(target.doc), 'xb<drawing>cd¶mo');
  const again = await saved(target);
  assert.deepEqual(Array.from(await again.pkg.readPart('/word/media/image1.png')), Array.from(PNG));
});

// --------------------------------------------------------------------- links

test('links: with their text, whole or in part; in another document the relationship is made, or one for the same address is used', async () => {
  const LINKS = p(r('ab') + '<w:hyperlink r:id="rIdLink" w:tooltip="tip">' + r('link') + '</w:hyperlink>' + r('cd') + '<w:hyperlink w:anchor="mark">' + r('inner') + '</w:hyperlink>');
  const source = await open({ body: LINKS, rels: [['rIdLink', 'hyperlink', 'https://example.com/a?b=1&c=2', true]] });
  const history = new DocHistory(source.doc);
  const selection = range(at(0, 1), at(0, 11));
  assert.equal(cutLoses(source.doc, selection, sliceOf(source.doc, selection.anchor, selection.focus)), false);
  const part = range(at(0, 3), at(0, 5));
  assert.equal(cutLoses(source.doc, part, sliceOf(source.doc, part.anchor, part.focus)), false);
  assert.equal(shape(fragmentOf(source.doc, part).blocks), '[hyperlink:in]');
  const fragment = cut(source.doc, history, selection);
  assert.equal(shape(fragment.blocks), 'b[hyperlink:link]cd[hyperlink:inn]');
  assert.equal(body(source.doc), 'a[hyperlink:er]');
  paste(source.doc, history, fragment, at(0, 1));
  assert.equal(body(source.doc), 'ab[hyperlink:link]cd[hyperlink:inn][hyperlink:er]');
  const link = ((source.doc.body.blocks[0] as Paragraph).children[1] as Run).wrap!;
  assert.deepEqual([link.kind, (link as { rId?: string }).rId, (link as { tooltip?: string }).tooltip], ['hyperlink', 'rIdLink', 'tip']);

  for (const rels of [[], [['rId3', 'hyperlink', 'https://example.com/a?b=1&c=2', true]]] as DocxSpec['rels'][]) {
    const target = await open({ body: p('target'), rels });
    const targetHistory = new DocHistory(target.doc);
    const relsBefore = target.doc.body.relationships;
    paste(target.doc, targetHistory, fragment, at(0, 3));
    assert.equal(body(target.doc), 'tarb[hyperlink:link]cd[hyperlink:inn]get');
    const pasted = ((target.doc.body.blocks[0] as Paragraph).children[1] as Run).wrap as { rId?: string; target?: string };
    const rel = target.doc.body.relationships.get(pasted.rId!)!;
    assert.deepEqual([rel.target, rel.external, /\/hyperlink$/.test(rel.type)], ['https://example.com/a?b=1&c=2', true, true]);
    if (rels!.length) {
      assert.equal(pasted.rId, 'rId3');
      assert.equal(target.doc.body.relationships, relsBefore);
    }
    const again = await saved(target);
    assert.equal(body(again.doc), 'tarb[hyperlink:link]cd[hyperlink:inn]get');
    const reread = ((again.doc.body.blocks[0] as Paragraph).children[1] as Run).wrap as { target?: string; tooltip?: string };
    assert.deepEqual([reread.target, reread.tooltip], ['https://example.com/a?b=1&c=2', 'tip']);
    const inner = ((again.doc.body.blocks[0] as Paragraph).children[3] as Run).wrap as { anchor?: string };
    assert.equal(inner.anchor, 'mark');
    targetHistory.undo();
    assert.equal(target.doc.body.relationships, relsBefore);
  }
});

// -------------------------------------------------------------------- fields

test('fields: a complex field with its code and result, a simple field; part of a result is text', async () => {
  const BODY = p(r('page ') + FIELD(' PAGE ', '7') + r(' of ') + '<w:fldSimple w:instr=" AUTHOR ">' + r('name') + '</w:fldSimple>' + r('.') + FIELD(' DATE \\@ "d MMMM" ', '4 October'));
  const opened = await open(BODY);
  const { doc } = opened;
  const history = new DocHistory(doc);
  const whole = range(at(0, 2), at(0, 16));
  assert.equal(cutLoses(doc, whole, sliceOf(doc, whole.anchor, whole.focus)), false);
  const fragment = fragmentOf(doc, whole);
  // The last field is only partly in the range: of it the text comes, as text. What is in the code of a field is flagged as that.
  assert.equal(shape(fragment.blocks), 'ge <begin>( PAGE )(<separate>)7<end> of [simpleField:name].4');
  const runs = (fragment.blocks[0] as Paragraph).children as Run[];
  assert.ok(runs.some((run) => run.fieldCode === true));
  const neat = fragmentOf(doc, range(at(0, 2), at(0, 15)));
  assert.equal(shape(neat.blocks), 'ge <begin>( PAGE )(<separate>)7<end> of [simpleField:name].');
  const target = await open(p('target'));
  paste(target.doc, new DocHistory(target.doc), neat, at(0, 6));
  assert.equal(body(target.doc), 'targetge <begin>( PAGE )(<separate>)7<end> of [simpleField:name].');
  const again = await saved(target);
  assert.equal(body(again.doc), 'targetge <begin>( PAGE )(<separate>)7<end> of [simpleField:name].');
  const begin = ((again.doc.body.blocks[0] as Paragraph).children as Run[]).map((run) => run.content[0]).find((item) => item.type === 'fieldChar' && item.charType === 'begin') as { field?: { instruction: string; result: string } };
  assert.deepEqual(begin.field, { instruction: ' PAGE ', result: '7' });
  const xml = decoder.decode(await again.pkg.readPart('/word/document.xml'));
  assert.ok(xml.indexOf('w:instr=" AUTHOR "') > 0);

  // A cut of the first two fields and a paste at the same place: a move.
  const start = body(doc);
  const cutRange = range(at(0, 5), at(0, 14));
  const moved = cut(doc, history, cutRange);
  assert.equal(body(doc), 'page .<begin>( DATE \\@ "d MMMM" )(<separate>)4 October<end>');
  paste(doc, history, moved, at(0, 5));
  assert.equal(body(doc), start);
  // Part of the result of a field that is not computed: the characters, and the field stays.
  const inside = range(at(0, 16), at(0, 20));
  assert.equal(shape(fragmentOf(doc, inside).blocks), ' Oct');
  assert.equal(cutLoses(doc, inside, sliceOf(doc, inside.anchor, inside.focus)), false);
});

// -------------------------------------------------- breaks, tabs, hyphens, symbols

test('breaks of pages, columns and lines, tabs, soft and non-breaking hyphens, symbols and paragraph ends are carried', async () => {
  const BODY =
    p('<w:r><w:t>ab</w:t><w:softHyphen/><w:t>c</w:t><w:br w:type="page"/><w:t>d</w:t><w:tab/><w:t>e</w:t><w:br/><w:t>f</w:t><w:noBreakHyphen/><w:t>g</w:t><w:br w:type="column"/><w:sym w:font="Wingdings" w:char="F04A"/><w:t xml:space="preserve">h i</w:t></w:r>') +
    p('second', '<w:jc w:val="center"/>') +
    p('third');
  const opened = await open(BODY);
  const { doc } = opened;
  const history = new DocHistory(doc);
  const start = body(doc);
  assert.equal(start, 'ab<softHyphen>c<break page>d<tab>e<break line>f<noBreakHyphen>g<break column><symbol>h i¶second¶third');
  const selection = range(at(0, 1), at(1, 3));
  assert.equal(cutLoses(doc, selection, sliceOf(doc, selection.anchor, selection.focus)), false);
  const fragment = cut(doc, history, selection);
  assert.equal(shape(fragment.blocks), 'b<softHyphen>c<break page>d<tab>e<break line>f<noBreakHyphen>g<break column><symbol>h i¶sec');
  assert.equal(body(doc), 'aond¶third');
  paste(doc, history, fragment, at(0, 1));
  assert.equal(body(doc), start);
  // Into another document, saved and read again.
  const target = await open(p('xy'));
  paste(target.doc, new DocHistory(target.doc), fragment, at(0, 1));
  const expected = 'xb<softHyphen>c<break page>d<tab>e<break line>f<noBreakHyphen>g<break column><symbol>h i¶secy';
  assert.equal(body(target.doc), expected);
  const again = await saved(target);
  assert.equal(body(again.doc), expected);
  const symbol = ((again.doc.body.blocks[0] as Paragraph).children as Run[]).flatMap((run) => run.content).find((item) => item.type === 'symbol') as { font?: string; char: string };
  assert.deepEqual([symbol.font, symbol.char.charCodeAt(0)], ['Wingdings', 0xf04a]);
});

// --------------------------------------------------------------------- tables

const TABLE = tbl(
  tr(tc('A1', '<w:tcW w:w="1000" w:type="dxa"/><w:shd w:val="clear" w:fill="FFFF00"/>') + tc('B1', '<w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/>')) +
    tr(tc('A2', '<w:tcW w:w="1000" w:type="dxa"/><w:vMerge w:val="restart"/>') + tc('B2', '<w:tcW w:w="1000" w:type="dxa"/>') + tc(p('C2') + tbl(tr(tc('in')), [500]) + p(''), '<w:tcW w:w="1000" w:type="dxa"/>')) +
    tr(tc('', '<w:tcW w:w="1000" w:type="dxa"/><w:vMerge/>') + tc('B3', '<w:tcW w:w="1000" w:type="dxa"/>') + tc('C3', '<w:tcW w:w="1000" w:type="dxa"/>')),
  [1000, 1000, 1000],
);

test('a table in the range comes as a table: grid, cell properties, merges, what its cells hold', async () => {
  const opened = await open(p('before') + TABLE + p('after'));
  const { doc } = opened;
  const history = new DocHistory(doc);
  const start = body(doc);
  assert.equal(start, 'before¶TABLE(A1 | B1 / A2 | B2 | C2¶TABLE(in)¶ /  | B3 | C3)¶after');
  const selection = range(at(0, 3), at(2, 2));
  assert.equal(cutLoses(doc, selection, sliceOf(doc, selection.anchor, selection.focus)), false);
  const fragment = cut(doc, history, selection);
  assert.equal(fragment.tableDepth, 2);
  assert.equal(body(doc), 'befter');
  paste(doc, history, fragment, at(0, 3));
  assert.equal(body(doc), start);
  const table = doc.body.blocks[1] as Table;
  assert.deepEqual(table.grid, [1000, 1000, 1000]);
  assert.equal(table.rows[0].cells[1].props.gridSpan, 2);
  assert.deepEqual(table.rows[0].cells[0].props.shading, { pattern: 'clear', fill: { rgb: 'FFFF00' } });
  assert.deepEqual([table.rows[1].cells[0].props.vMerge, table.rows[2].cells[0].props.vMerge], ['restart', 'continue']);
  const again = await saved(opened);
  assert.equal(body(again.doc), start);
  const reread = again.doc.body.blocks[1] as Table;
  assert.deepEqual(reread.grid, [1000, 1000, 1000]);
  assert.deepEqual([reread.rows[0].cells[1].props.gridSpan, reread.rows[1].cells[0].props.vMerge, reread.rows[2].cells[0].props.vMerge], [2, 'restart', 'continue']);

  // A copy of the table pasted at the end of a paragraph that a paragraph follows: no empty paragraph is made.
  const copy = fragmentOf(doc, range(at(0, 6), at(2, 0)));
  assert.equal(shape(copy.blocks), '¶TABLE(A1 | B1 / A2 | B2 | C2¶TABLE(in)¶ /  | B3 | C3)¶');
  // In a cell a table becomes a table inside the cell.
  paste(doc, history, copy, at([1, 0, 0, 0], 2));
  const cell = (doc.body.blocks[1] as Table).rows[0].cells[0];
  assert.equal(shape(cell.blocks), 'A1¶TABLE(A1 | B1 / A2 | B2 | C2¶TABLE(in)¶ /  | B3 | C3)¶');
  history.undo();
  assert.equal(body(doc), start);
});

test('rows and cells of a table: whole rows as rows, a rectangle of cells as a table of those columns, a range that leaves the table', async () => {
  const opened = await open(p('before') + TABLE + p('after'));
  const { doc } = opened;
  const history = new DocHistory(doc);
  // Two whole rows.
  let fragment = fragmentOf(doc, expandRangeForTables(doc, range(at([1, 1, 0, 0], 0), at([1, 2, 2, 0], 2))));
  assert.equal(shape(fragment.blocks), 'TABLE(A2 | B2 | C2¶TABLE(in)¶ /  | B3 | C3)');
  let table = fragment.blocks[0] as Table;
  assert.deepEqual(table.grid, [1000, 1000, 1000]);
  assert.deepEqual([fragment.openStart, fragment.openEnd], [false, false]);
  // A rectangle: the last two columns of the last two rows.
  fragment = fragmentOf(doc, expandRangeForTables(doc, range(at([1, 1, 1, 0], 1), at([1, 2, 2, 0], 1))));
  assert.equal(shape(fragment.blocks), 'TABLE(B2 | C2¶TABLE(in)¶ / B3 | C3)');
  table = fragment.blocks[0] as Table;
  assert.deepEqual(table.grid, [1000, 1000]);
  assert.deepEqual(table.props.width, { type: 'auto', value: 0 });
  // A row that continued a merged cell begins it in the copy.
  fragment = fragmentOf(doc, expandRangeForTables(doc, range(at([1, 2, 0, 0], 0), at([1, 2, 2, 0], 2))));
  assert.equal((fragment.blocks[0] as Table).rows[0].cells[0].props.vMerge, 'restart');
  // Pasted between paragraphs, after a table: a paragraph keeps the two tables apart.
  paste(doc, history, fragment, at(2, 0));
  assert.equal(body(doc), 'before¶TABLE(A1 | B1 / A2 | B2 | C2¶TABLE(in)¶ /  | B3 | C3)¶¶TABLE( | B3 | C3)¶after');
  history.undo();
  // A range from the text into the table: the cells a deletion empties, as rows; a cut of it is pasted back whole.
  const selection = range(at(0, 3), at([1, 1, 1, 0], 1));
  const shown = orderRange(expandRangeForTables(doc, selection));
  const slice = sliceOf(doc, shown.start, shown.end)!;
  assert.equal(cutLoses(doc, selection, slice), false);
  assert.equal(shape(slice.fragment!.blocks), 'ore¶TABLE(A1 | B1 / A2 | B2 | )');
  const cutOut = cut(doc, history, selection);
  assert.equal(body(doc), 'bef¶TABLE( |  /  |  | C2¶TABLE(in)¶ /  | B3 | C3)¶after');
  paste(doc, history, cutOut, at(2, 0));
  assert.equal(body(doc), 'bef¶TABLE( |  /  |  | C2¶TABLE(in)¶ /  | B3 | C3)¶ore¶TABLE(A1 | B1 / A2 | B2 | )¶after');
  const again = await saved(opened);
  assert.equal(body(again.doc), body(doc));
});

test('tables are not nested deeper than the limit by pasting', async () => {
  const opened = await open(p('before') + tbl(tr(tc('A1') + tc('B1')), [1000, 1000]) + p('after'));
  const { doc } = opened;
  const history = new DocHistory(doc);
  // The whole table into its own first cell, again and again: the depth doubles each time.
  let failed = false;
  for (let i = 0; i < 12 && !failed; i++) {
    const fragment = fragmentOf(doc, range(at(0, 6), at(2, 0)));
    const path = [1];
    for (let table = doc.body.blocks[1] as Table; ; ) {
      path.push(0, 0);
      const inner = table.rows[0].cells[0].blocks.findIndex((block) => block.type === 'table');
      if (inner < 0) {
        path.push(0);
        break;
      }
      path.push(inner);
      table = table.rows[0].cells[0].blocks[inner] as Table;
    }
    const before = body(doc);
    try {
      paste(doc, history, fragment, at(path, 0));
    } catch (error) {
      assert.ok(error instanceof RangeError);
      assert.equal(body(doc), before, 'a refused paste leaves the document as it was');
      failed = true;
    }
  }
  assert.ok(failed, 'the paste was refused before the tables got too deep');
  let depth = 0;
  for (let block: Block | undefined = doc.body.blocks[1]; block && block.type === 'table'; block = block.rows[0].cells[0].blocks.find((inner) => inner.type === 'table')) depth++;
  assert.ok(depth <= MAX_TABLE_DEPTH, String(depth));
});

// ---------------------------------------------- paragraph properties, styles, lists

test('paragraph properties, styles and lists: the ids stay in the same document', async () => {
  const BODY =
    p('Title', '<w:pStyle w:val="Heading1"/><w:jc w:val="center"/><w:ind w:left="300"/>') +
    p(r('item ') + r('strong', '<w:rStyle w:val="Strong"/>'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>') +
    p('tail');
  const opened = await open({ body: BODY, numbering: NUMBERING });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const stylesBefore = doc.styles;
  const fragment = fragmentOf(doc, range(at(0, 0), at(2, 0)));
  assert.deepEqual(Array.from(fragment.styles.keys()).sort(), ['Heading1', 'Normal', 'Strong']);
  assert.deepEqual([Array.from(fragment.nums.keys()), Array.from(fragment.abstracts.keys())], [[7], [4]]);
  paste(doc, history, fragment, at(2, 0));
  assert.equal(body(doc), 'Title¶item strong¶Title¶item strong¶tail');
  const [first, second] = [doc.body.blocks[2] as Paragraph, doc.body.blocks[3] as Paragraph];
  assert.deepEqual([first.props.style, first.props.align, first.props.indent], ['Heading1', 'center', { left: 300 }]);
  assert.deepEqual(second.props.numbering, { numId: 7, level: 0 });
  assert.equal((second.children[1] as Run).props.style, 'Strong');
  // The property bags are the ones of the document: nothing was copied.
  assert.equal(first.props, (doc.body.blocks[0] as Paragraph).props);
  assert.equal(doc.styles, stylesBefore);
  assert.equal(doc.numbering.nums.size, 1);
  assert.equal((doc.body.blocks[4] as Paragraph).props.style, undefined, 'the paragraph the caret was in keeps its own formatting');
});

test('another document: a style it lacks is added with what it is based on, one it has is its own; a list comes as a new list; undo takes both out', async () => {
  const BODY =
    p('Quoted', '<w:pStyle w:val="Quote"/>') + p(r('item ') + r('strong', '<w:rStyle w:val="Strong"/>'), '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>') + p('list made here') + p('tail');
  const source = await open({ body: BODY, numbering: NUMBERING });
  // A list made in the editor, not in the numbering part.
  applyDocOps(source.doc, toggleList(source.doc, at(2, 0), 'bullet').ops);
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(3, 0)));
  assert.equal(fragment.nums.size, 2);

  const target = await open({
    body: p('first', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr>') + p('last'),
    styles: PLAIN_STYLES + '<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:i/></w:rPr></w:style>',
    numbering: '<w:abstractNum w:abstractNumId="4"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="4"/></w:num>',
  });
  const history = new DocHistory(target.doc);
  const stylesBefore = target.doc.styles;
  const strong = target.doc.styles.styles.get('Strong');
  paste(target.doc, history, fragment, at(1, 0));
  assert.equal(body(target.doc), 'first¶Quoted¶item strong¶list made here¶last');
  // Quote and the Heading1 it is based on came along; Normal and Strong are the document's own.
  assert.notEqual(target.doc.styles, stylesBefore, 'a new table of styles: what depends on the styles sees the change');
  assert.deepEqual(Array.from(target.doc.styles.styles.keys()).sort(), ['Heading1', 'Normal', 'Quote', 'Strong']);
  assert.equal(target.doc.styles.styles.get('Strong'), strong);
  assert.equal(target.doc.styles.styles.get('Quote')!.basedOn, 'Heading1');
  // The list of the copy is not list 7 of this document.
  const item = target.doc.body.blocks[2] as Paragraph;
  const numId = item.props.numbering!.numId!;
  assert.notEqual(numId, 7);
  const num = target.doc.numbering.nums.get(numId)!;
  assert.equal(target.doc.numbering.abstracts.get(num.abstractId)!.levels.get(0)!.format, 'decimal');
  assert.equal(target.doc.numbering.abstracts.get(4)!.levels.get(0)!.format, 'bullet');
  const made = (target.doc.body.blocks[3] as Paragraph).props.numbering!.numId!;
  assert.ok(made !== numId && made !== 7 && target.doc.numbering.nums.has(made));
  assert.equal(target.doc.numbering.nums.size, 3);

  const again = await saved(target);
  assert.equal(body(again.doc), 'first¶Quoted¶item strong¶list made here¶last');
  assert.deepEqual(Array.from(again.doc.styles.styles.keys()).sort(), ['Heading1', 'Normal', 'Quote', 'Strong']);
  assert.equal(again.doc.styles.styles.get('Quote')!.run!.italic, true);
  assert.equal((again.doc.body.blocks[1] as Paragraph).props.style, 'Quote');
  const reread = (again.doc.body.blocks[2] as Paragraph).props.numbering!.numId!;
  assert.equal(again.doc.numbering.abstracts.get(again.doc.numbering.nums.get(reread)!.abstractId)!.levels.get(0)!.format, 'decimal');
  assert.equal(again.doc.numbering.nums.size, 3);

  history.undo();
  assert.equal(body(target.doc), 'first¶last');
  assert.equal(target.doc.styles, stylesBefore);
  assert.deepEqual(Array.from(target.doc.numbering.nums.keys()), [7]);
  assert.deepEqual(Array.from(target.doc.numbering.abstracts.keys()), [4]);
});

test('a section break goes with its paragraph into the body, without the headers of another document, and not into a table cell', async () => {
  const SECTION = '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="5000" w:h="7000"/></w:sectPr>';
  const source = await open({ body: p('one', SECTION) + p('two') + tbl(tr(tc('cell')), [1000]) + p('end'), headers: { rIdH: p('header') } });
  const { doc } = source;
  const history = new DocHistory(doc);
  const selection = range(at(0, 1), at(1, 1));
  assert.equal(cutLoses(doc, selection, sliceOf(doc, selection.anchor, selection.focus)), false);
  const fragment = fragmentOf(doc, selection);
  paste(doc, history, fragment, at(3, 0));
  const pasted = doc.body.blocks[3] as Paragraph;
  assert.deepEqual(pasted.props.section!.headers, { default: 'rIdH' });
  history.undo();
  paste(doc, history, fragment, at([2, 0, 0, 0], 0));
  const inCell = (doc.body.blocks[2] as Table).rows[0].cells[0].blocks[0] as Paragraph;
  assert.equal(paragraphText(inCell), 'ne');
  assert.equal(inCell.props.section, undefined);
  history.undo();
  const target = await open(p('x'));
  paste(target.doc, new DocHistory(target.doc), fragment, at(0, 0));
  const other = target.doc.body.blocks[0] as Paragraph;
  assert.deepEqual(other.props.section!.page, { width: 5000, height: 7000 });
  assert.equal(other.props.section!.headers, undefined);
});

test('pasted inside a link or a content control the copy stays inside of it; a link of the copy joins the link around it', async () => {
  const BODY =
    p(r('ab') + '<w:hyperlink r:id="rIdLink">' + r('link') + '</w:hyperlink>' + r('cd')) +
    p(r('ab') + '<w:sdt><w:sdtPr><w:id w:val="5"/></w:sdtPr><w:sdtContent>' + r('control') + '</w:sdtContent></w:sdt>' + r('cd')) +
    p(r('x') + '<w:hyperlink w:anchor="mark">' + r('other') + '</w:hyperlink>' + FIELD(' DATE ', 'today'));
  const opened = await open({ body: BODY, rels: [['rIdLink', 'hyperlink', 'https://example.com/', true]] });
  const { doc } = opened;
  const history = new DocHistory(doc);
  const fragment = fragmentOf(doc, range(at(2, 3), at(2, 11)));
  assert.equal(shape(fragment.blocks), '[hyperlink:her]<begin>( DATE )(<separate>)today<end>');
  paste(doc, history, fragment, at(0, 4));
  assert.equal(shape([doc.body.blocks[0]]), 'ab[hyperlink:liher][hyperlink:<begin>][hyperlink:( DATE )][hyperlink:(<separate>)][hyperlink:today][hyperlink:<end>][hyperlink:nk]cd');
  // One link still: every piece of it is in the same wrapper.
  const links = new Set((doc.body.blocks[0] as Paragraph).children.map((inline) => inline.wrap).filter((wrap) => wrap));
  assert.equal(links.size, 1);
  paste(doc, history, fragment, at(1, 5));
  assert.equal(shape([doc.body.blocks[1]]), 'ab[sdt:con][sdt:hyperlink:her][sdt:<begin>][sdt:( DATE )][sdt:(<separate>)][sdt:today][sdt:<end>][sdt:trol]cd');
  const again = await saved(opened);
  assert.equal(shape([again.doc.body.blocks[0]]), shape([doc.body.blocks[0]]));
  assert.equal(shape([again.doc.body.blocks[1]]), shape([doc.body.blocks[1]]));
  // The control was not split: the document has one, with its id.
  const xml = decoder.decode(await again.pkg.readPart('/word/document.xml'));
  assert.equal((xml.match(/<w:sdt>/g) || []).length, 1);
  assert.equal((xml.match(/<w:hyperlink /g) || []).length, 3);
});

test('a copy from a file that is no package (RTF): its picture, its field and its style go into a docx and are saved there', async () => {
  const hex = Array.from(PNG, (byte) => ('0' + byte.toString(16)).slice(-2)).join('');
  const rtf = `{\\rtf1\\ansi{\\stylesheet{\\s1\\b Heading;}}\\pard\\s1 Title\\par\\pard before {\\pict\\pngblip\\picw2\\pich1\\picwgoal300\\pichgoal150 ${hex}} after {\\field{\\*\\fldinst HYPERLINK "https://example.com/"}{\\fldrslt link}} end\\par second\\par}`;
  const source = readRtf(new TextEncoder().encode(rtf));
  const fragment = fragmentOf(source, range(at(0, 0), at(2, 3)));
  assert.equal(fragment.media.size, 1);
  fragment.media.forEach((media) => assert.equal(media.bytes!.length, PNG.length));
  const target = await open({ body: p('xy'), styles: PLAIN_STYLES });
  const history = new DocHistory(target.doc);
  paste(target.doc, history, fragment, at(0, 1));
  assert.deepEqual(Array.from(target.doc.media!.keys()), ['/word/media/image1.png']);
  assert.ok(target.doc.styles.styles.has('Heading'));
  const again = await saved(target);
  assert.equal(body(again.doc).replace(/\(?<separate>\)?/, '<separate>'), 'xTitle¶before <drawing> after <begin>(HYPERLINK "https://example.com/")<separate>link<end> end¶secy');
  assert.equal(drawings(again.doc)[0].image!.target, '/word/media/image1.png');
  assert.deepEqual(Array.from(await again.pkg.readPart('/word/media/image1.png')), Array.from(PNG));
  assert.equal((again.doc.body.blocks[0] as Paragraph).props.style, 'Heading');
  history.undo();
  assert.equal(target.doc.media, undefined);
  assert.ok(!target.doc.styles.styles.has('Heading'));
});

// ------------------------------------------------- what a document cannot take

test('kept XML brings the namespaces the other document does not declare; a document of the other kind is refused', async () => {
  const BODY = p(r('fancy', '<w14:ligatures w14:val="all"/>') + picture('rIdImg', 2));
  const source = await open({ body: BODY, ...WITH_PICTURE });
  const fragment = fragmentOf(source.doc, range(at(0, 0), at(0, 5)));
  fragment.media.get('/word/media/image1.png')!.bytes = PNG;
  // A document whose root declares neither w14 nor the drawing namespaces.
  const bare =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:body>${p('x')}${SECT}</w:body></w:document>`;
  const target = await open({ document: bare });
  paste(target.doc, new DocHistory(target.doc), fragment, at(0, 1));
  assert.equal(body(target.doc), 'xfancy<drawing>');
  const again = await saved(target);
  assert.equal(body(again.doc), 'xfancy<drawing>');
  const xml = decoder.decode(await again.pkg.readPart('/word/document.xml'));
  assert.ok(/<w14:ligatures xmlns:w14="http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordml" w14:val="all"\/>|<w14:ligatures w14:val="all" xmlns:w14="http:\/\/schemas.microsoft.com\/office\/word\/2010\/wordml"\/>/.test(xml), xml);
  assert.ok(/<w:drawing [^>]*xmlns:wp="http:\/\/schemas.openxmlformats.org\/drawingml\/2006\/wordprocessingDrawing"/.test(xml), xml);
  assert.equal(drawings(again.doc)[0].image!.target, '/word/media/image1.png');
  assert.equal(drawings(again.doc)[0].kind, 'picture');

  const strict = await open({ body: p('x'), strict: true });
  const before = body(strict.doc);
  assert.throws(() => insertFragment(strict.doc, at(0, 0), fragment), RangeError);
  assert.equal(body(strict.doc), before);
  assert.equal(strict.doc.body.modified, undefined);
});

test('what a copy does not carry is left out of it, and a cut of it is refused', async () => {
  const TEXT_BOX =
    '<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="8" name="box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>' +
    p('boxed') +
    '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const CHART =
    '<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="9" name="chart"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const opened = await open({
    body:
      p(r('note') + '<w:r><w:footnoteReference w:id="2"/></w:r>' + r(' after')) + // 0
      p(r('ab') + '<w:sdt><w:sdtPr/><w:sdtContent>' + r('ctl') + '</w:sdtContent></w:sdt>' + r('cd')) + // 1
      p(r('ab') + '<w:ins w:id="1" w:author="a">' + r('ins') + '</w:ins>' + r('cd')) + // 2
      p(r('ab') + '<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>' + r('cd')) + // 3
      p(r('ab') + TEXT_BOX + r('cd')) + // 4
      p(r('ab') + CHART + r('cd')) + // 5
      p(r('ab') + picture('rIdGone', 4) + r('cd')) + // 6
      p(r('abc') + '<w:r><w:commentReference w:id="9"/></w:r>' + r('def')) + // 7
      p('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' + r('one')) + // 8
      p(r('two') + '<w:r><w:fldChar w:fldCharType="end"/></w:r>') + // 9
      '<w:sdt><w:sdtPr/><w:sdtContent>' +
      p('in a control') +
      '</w:sdtContent></w:sdt>' + // 10
      p('end'), // 11
    footnotes: '<w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> the note</w:t></w:r></w:p></w:footnote>',
    rels: [['rIdChart', 'chart', 'charts/chart1.xml']],
    parts: [['word/charts/chart1.xml', '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>']],
  });
  const { doc } = opened;
  const picture0 = JSON.stringify(doc.body.blocks, (key, value) => (value instanceof Map ? Array.from(value) : value));
  const check = (start: DocPosition, end: DocPosition, copied: string, loses: boolean, label: string) => {
    const slice = sliceOf(doc, start, end)!;
    assert.equal(shape(slice.fragment!.blocks), copied, label);
    assert.equal(cutLoses(doc, range(start, end), slice), loses, label);
    assert.equal(JSON.stringify(doc.body.blocks, (key, value) => (value instanceof Map ? Array.from(value) : value)), picture0, label);
  };
  check(at(0, 2), at(0, 6), 'te a', true, 'a note reference');
  check(at(1, 1), at(1, 6), 'bctlc', true, 'a content control: its text comes without it');
  check(at(2, 1), at(2, 6), 'binsc', true, 'a tracked insertion');
  check(at(3, 1), at(3, 4), 'bc', true, 'an equation');
  check(at(4, 1), at(4, 3), 'bc', true, 'a text box');
  check(at(5, 1), at(5, 3), 'bc', true, 'a chart');
  check(at(6, 1), at(6, 3), 'bc', true, 'a picture whose relationship is missing');
  check(at(7, 1), at(7, 5), 'bcde', false, 'a comment reference stays in the document: nothing is lost');
  check(at(8, 0), at(9, 3), 'one¶two', true, 'a field over two paragraphs');
  check(at(9, 3), at(11, 1), '¶in a control¶e', true, 'a paragraph in a block content control');
  check(at(0, 0), at(0, 3), 'not', false, 'text before a note reference');
  // Without a copy the clipboard holds the text only: a table cannot go then.
  const tableDoc = (await open(p('a') + tbl(tr(tc('A1')), [1000]) + p('b'))).doc;
  assert.equal(cutLoses(tableDoc, range(at(0, 0), at(2, 1)), undefined), true);
  assert.equal(cutLoses(tableDoc, range(at(0, 0), at(2, 1)), sliceOf(tableDoc, at(0, 0), at(2, 1))), false);
  // A copy that does not hold what the deletion removes (the ends of the two do not agree).
  assert.equal(cutLoses(tableDoc, range(at(0, 0), at(2, 1)), sliceOf(tableDoc, at(0, 0), at(0, 1))), true);
});

test('text inside a text box can be copied (not edited): the copy is taken from the box', async () => {
  const BOX =
    '<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/><wp:docPr id="8" name="box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
    '<wps:wsp><wps:spPr/><wps:txbx><w:txbxContent>' +
    p('boxed one') +
    tbl(tr(tc('in a cell')), [800]) +
    p(r('boxed ') + r('two', '<w:b/>')) +
    '</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const { doc } = await open(p(r('ab') + BOX + r('cd')) + p('end'));
  const inBox = (path: number[], offset: number): DocPosition => ({ story: 'body', path: [0, -1, 0, 0].concat(path), offset });
  assert.equal(shape(fragmentOf(doc, range(inBox([0], 2), inBox([0], 7))).blocks), 'xed o');
  const fragment = fragmentOf(doc, range(inBox([0], 6), inBox([2], 8)));
  assert.equal(shape(fragment.blocks), 'one¶TABLE(in a cell)¶boxed tw');
  assert.equal(sliceOf(doc, inBox([0], 6), inBox([2], 8))!.fragment!.tableDepth, 1);
  // Pasted into the text of the document; into the box itself nothing can be pasted.
  const history = new DocHistory(doc);
  paste(doc, history, fragment, at(1, 3));
  assert.equal(shape(doc.body.blocks.slice(1)), 'endone¶TABLE(in a cell)¶boxed tw');
  assert.throws(() => insertFragment(doc, inBox([0], 2), fragment), RangeError);
  // The two ends of a range are in one flow of text.
  assert.throws(() => fragmentOf(doc, range(at(0, 1), inBox([0], 2))), RangeError);
});

// ------------------------------------------------- the copies a page keeps

test('the copies of the page: found by mark and text, the last few only, up to a number of bytes', async () => {
  clearStoredCopies();
  const { doc } = await open(p('abcdef'));
  const slice = sliceOf(doc, at(0, 1), at(0, 4))!;
  const copy = storeCopy('bcd', slice);
  assert.ok(/^[a-z0-9]+$/.test(copy.mark));
  const html = sliceHtml(slice, copy.mark);
  assert.equal(copyMarkOf(html), copy.mark);
  assert.equal(storedCopy(html, 'bcd'), copy);
  // Line ends are compared as the clipboard may hand them back.
  const lines = storeCopy('a\nb', slice);
  assert.equal(storedCopy(sliceHtml(slice, lines.mark), 'a\r\nb'), lines);
  // Another text, another mark, no HTML, a mark that is no mark: not a copy of the page.
  assert.equal(storedCopy(html, 'bcd '), undefined);
  assert.equal(storedCopy(sliceHtml(slice, 'zzz'), 'bcd'), undefined);
  assert.equal(storedCopy(undefined, 'bcd'), undefined);
  assert.equal(storedCopy('<div data-bym-doc-copy="<script>">', 'bcd'), undefined);
  assert.equal(copyMarkOf('<div data-bym-doc-copy="a b">'), undefined);

  for (let i = 0; i < MAX_STORED_COPIES + 5; i++) storeCopy('text ' + i, slice);
  assert.equal(storedCopies().count, MAX_STORED_COPIES);
  assert.equal(storedCopy(html, 'bcd'), undefined, 'the oldest copies are forgotten');

  // Pictures count: copies that hold too much together push the older ones out.
  clearStoredCopies();
  const big = Math.floor(MAX_STORED_BYTES / 2) - 4096;
  const withPicture = async (): Promise<ReturnType<typeof storeCopy>> => {
    const opened = await open({ body: p(r('ab') + picture('rIdImg', 5) + r('cd')), ...WITH_PICTURE });
    const made = storeCopy('bc', sliceOf(opened.doc, at(0, 1), at(0, 3))!);
    setCopyMedia(made, '/word/media/image1.png', new Uint8Array(big));
    return made;
  };
  const one = await withPicture();
  const two = await withPicture();
  assert.equal(storedCopies().count, 2);
  assert.ok(storedCopies().bytes >= 2 * big && storedCopies().bytes <= MAX_STORED_BYTES);
  const three = await withPicture();
  assert.equal(storedCopies().count, 2);
  assert.equal(storedCopy(sliceHtml(one.slice, one.mark), 'bc'), undefined);
  assert.equal(storedCopy(sliceHtml(two.slice, two.mark), 'bc'), two);
  assert.equal(storedCopy(sliceHtml(three.slice, three.mark), 'bc'), three);
  // A picture handed to a copy that was forgotten is not kept.
  setCopyMedia(one, '/word/media/image1.png', new Uint8Array(16));
  assert.ok(storedCopies().bytes <= MAX_STORED_BYTES);
  // One copy that is too large alone keeps its content without the bytes of its pictures.
  clearStoredCopies();
  const huge = await withPicture();
  setCopyMedia(huge, '/word/media/image1.png', new Uint8Array(1));
  const opened = await open({ body: p(r('ab') + picture('rIdImg', 5) + r('cd')), ...WITH_PICTURE });
  const alone = storeCopy('bc', sliceOf(opened.doc, at(0, 1), at(0, 3))!);
  clearStoredCopies();
  const kept = storeCopy('bc', alone.slice);
  setCopyMedia(kept, '/word/media/image1.png', new Uint8Array(MAX_STORED_BYTES + 1));
  assert.equal(storedCopies().count, 1);
  assert.ok(storedCopies().bytes <= MAX_STORED_BYTES);
  assert.equal(kept.slice.fragment!.media.get('/word/media/image1.png')!.bytes, undefined);
  assert.equal(shape(kept.slice.fragment!.blocks), 'b<drawing>c');
  clearStoredCopies();
  assert.equal(storedCopies().count, 0);
});

test('a copy that is too large to keep is plain text; a fragment that is not one is refused', async () => {
  const many = Array.from({ length: 5200 }, (_, i) => p('p' + i)).join('');
  const { doc } = await open(many);
  assert.throws(() => fragmentOf(doc, range(at(0, 0), at(5199, 1))), RangeError);
  assert.equal(sliceOf(doc, at(0, 0), at(5199, 1)), undefined);
  assert.throws(() => insertFragment(doc, at(0, 0), {} as DocFragment), RangeError);
  assert.throws(() => insertFragment(doc, at(0, 0), { ...fragmentOf(doc, range(at(0, 0), at(0, 1))), blocks: [{ type: 'unknown' } as never] }), RangeError);
  assert.throws(() => fragmentOf(doc, range(at(0, 0), at(9000, 1))), RangeError);
  // An empty range is an empty copy; pasting it changes nothing.
  const empty = fragmentOf(doc, range(at(0, 1), at(0, 1)));
  assert.equal(empty.blocks.length, 0);
  assert.equal(insertFragment(doc, at(0, 0), empty).ops.length, 0);
});
