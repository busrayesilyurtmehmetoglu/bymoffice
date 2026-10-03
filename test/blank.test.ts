// New, empty files (createBlank): they open, they are what an editor expects
// to start from, they survive a save, and edits made in them are written.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { applyOp, createBlank, docOps, open, paragraphText, parseXml, resolveParaProps, resolveRunProps, setCellInput } from '../src/index';
import { getCell } from '../src/model/sheet/types';

const text = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

test('a blank docx: one empty paragraph, the styles of an editor, the paper and the language of the locale', async () => {
  const bytes = await createBlank('docx', { locale: 'tr' });
  const file = await open(bytes);
  assert.equal(file.format, 'docx');
  const doc = await file.document!();
  assert.equal(doc.body.blocks.length, 1);
  assert.equal(doc.body.blocks[0].type, 'paragraph');
  assert.equal(paragraphText(doc.body.blocks[0] as never), '');
  for (const id of ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Title', 'Hyperlink', 'TableNormal']) assert.ok(doc.styles.styles.has(id), id);
  // A4 with margins of 2.5 cm, Calibri 11.
  const section = doc.body.section as { page?: { width?: number; height?: number } };
  assert.deepEqual([section.page!.width, section.page!.height], [11906, 16838]);
  const para = resolveParaProps(doc, (doc.body.blocks[0] as { props: never }).props);
  const run = resolveRunProps(doc, para, {} as never) as { size?: number; fonts?: { ascii?: string } };
  assert.equal(run.size, 22);
  assert.equal(run.fonts!.ascii, 'Calibri');
  const styles = text(await file.package!.readPart('/word/styles.xml'));
  assert.ok(styles.indexOf('w:lang w:val="tr-TR"') > 0);
  // Every part is XML that parses, and the file is the same after a save that changed nothing.
  for (const part of ['/word/document.xml', '/word/styles.xml', '/word/settings.xml']) parseXml(text(await file.package!.readPart(part)));
  const again = await open(await file.save());
  assert.equal((await again.document!()).body.blocks.length, 1);
});

test('a blank docx in English is on Letter paper for the United States and on A4 elsewhere', async () => {
  const size = async (locale?: string) => {
    const doc = await (await open(await createBlank('docx', { locale }))).document!();
    const section = doc.body.section as { page?: { width?: number; height?: number } };
    return [section.page!.width, section.page!.height];
  };
  assert.deepEqual(await size('en'), [12240, 15840]);
  assert.deepEqual(await size('en-US'), [12240, 15840]);
  assert.deepEqual(await size('en-GB'), [11906, 16838]);
  assert.deepEqual(await size(), [11906, 16838]);
  assert.deepEqual(await size('de-DE'), [11906, 16838]);
  // A locale that is not a language tag never reaches the XML.
  const odd = new TextDecoder().decode(await (await open(await createBlank('docx', { locale: '"/><x' }))).package!.readPart('/word/styles.xml'));
  assert.ok(odd.indexOf('w:lang w:val="en-GB"') > 0);
});

test('a blank docx can be written in: text, a heading, a list, a table, and it is saved with them', async () => {
  const file = await open(await createBlank('docx'));
  const doc = await file.document!();
  const history = new docOps.DocHistory(doc);
  const at = (path: number[], offset: number) => ({ story: 'body', path, offset });
  history.applyEdit(docOps.insertText(doc, at([0], 0), 'Title line\nFirst item\nSecond item\nAfter'), undefined);
  history.applyEdit(docOps.setParaFormat(doc, at([0], 0), { style: 'Heading1' }), undefined);
  history.applyEdit(docOps.toggleList(doc, { anchor: at([1], 0), focus: at([2], 3) }, 'number'), undefined);
  history.applyEdit(docOps.insertTable(doc, at([3], 5), 2, 2), undefined);
  const saved = await open(await file.save());
  const read = await saved.document!();
  const paragraphs = read.body.blocks.filter((block) => block.type === 'paragraph').map((block) => paragraphText(block as never));
  assert.deepEqual(paragraphs.slice(0, 4), ['Title line', 'First item', 'Second item', 'After']);
  assert.equal((read.body.blocks[0] as { props: { style?: string } }).props.style, 'Heading1');
  assert.ok(read.body.blocks.some((block) => block.type === 'table'));
  assert.ok((read.body.blocks[1] as { props: { numbering?: unknown } }).props.numbering, 'the list item is numbered');
});

test('a blank xlsx: one empty sheet named by the locale, which takes input and is saved with it', async () => {
  const bytes = await createBlank('xlsx', { locale: 'tr' });
  const file = await open(bytes);
  assert.equal(file.format, 'xlsx');
  const workbook = await file.workbook();
  assert.deepEqual(workbook.sheets.map((sheet) => [sheet.name, sheet.state, sheet.rows.size]), [['Sayfa1', 'visible', 0]]);
  assert.equal(workbook.styles.fonts[0].name, 'Calibri');
  assert.equal((await (await open(await createBlank('xlsx'))).workbook()).sheets[0].name, 'Sheet1');
  assert.equal((await (await open(await createBlank('xlsx', { sheetName: 'Bütçe 2026' }))).workbook()).sheets[0].name, 'Bütçe 2026');
  for (const op of setCellInput(workbook, 0, 0, 0, '12.5')) applyOp(workbook, op);
  for (const op of setCellInput(workbook, 0, 1, 0, 'Toplam')) applyOp(workbook, op);
  const again = await (await open(await file.save())).workbook();
  assert.equal(getCell(again.sheets[0], 0, 0)!.value, 12.5);
  assert.equal(getCell(again.sheets[0], 1, 0)!.value, 'Toplam');
});

test('the same options give the same bytes; what cannot be made is refused', async () => {
  for (const kind of ['docx', 'xlsx'] as const) {
    const a = await createBlank(kind, { locale: 'tr' });
    const b = await createBlank(kind, { locale: 'tr' });
    assert.deepEqual(Array.from(a), Array.from(b), kind);
  }
  await assert.rejects(() => createBlank('pptx' as never), RangeError);
  for (const name of ['', 'a'.repeat(32), 'a/b', 'a:b', '[x]', "'x", "x'", 'a*b', 'a?b', 'a\\b', ' ', '   ', 'History', 'history', 'a\tb', 'a\uffffb', 'a\ufffeb', 'a\ud800b', '\udc00b']) await assert.rejects(() => createBlank('xlsx', { sheetName: name }), RangeError, name);
});

test('review: the locale is a language tag in any case, with a hyphen or an underscore', async () => {
  const made = async (locale: string) => {
    const file = await open(await createBlank('docx', { locale }));
    const read = async (part: string) => new TextDecoder().decode(await file.package!.readPart(part));
    const styles = await read('/word/styles.xml');
    const settings = await read('/word/settings.xml');
    const section = (await file.document!()).body.section as { page?: { width?: number } };
    const sheet = (await (await open(await createBlank('xlsx', { locale }))).workbook()).sheets[0].name;
    return [/w:lang w:val="([^"]*)"/.exec(styles)![1], section.page!.width === 12240 ? 'Letter' : 'A4', /w:decimalSymbol w:val="(.)"/.exec(settings)![1], /w:listSeparator w:val="(.)"/.exec(settings)![1], sheet];
  };
  for (const locale of ['tr', 'TR', 'tr-TR', 'tr-tr', 'tr_TR', 'TR_tr']) assert.deepEqual(await made(locale), ['tr-TR', 'A4', ',', ';', 'Sayfa1'], locale);
  assert.deepEqual(await made('tr-CY'), ['tr-CY', 'A4', ',', ';', 'Sayfa1']);
  for (const locale of ['en', 'EN', 'en-US', 'en-us', 'EN_us']) assert.deepEqual(await made(locale), ['en-US', 'Letter', '.', ',', 'Sheet1'], locale);
  assert.deepEqual(await made('en-ca'), ['en-CA', 'Letter', '.', ',', 'Sheet1']);
  assert.deepEqual(await made('en-gb'), ['en-GB', 'A4', '.', ',', 'Sheet1']);
  assert.deepEqual(await made('zh-hant-tw'), ['zh-Hant-TW', 'A4', '.', ',', 'Sheet1']);
  assert.deepEqual(await made('not a tag'), ['en-GB', 'A4', '.', ',', 'Sheet1']);
  // A pair of surrogates is a character like any other in a sheet name.
  assert.equal((await (await open(await createBlank('xlsx', { sheetName: 'Rapor 📊' }))).workbook()).sheets[0].name, 'Rapor 📊');
});
