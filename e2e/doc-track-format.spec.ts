import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES } from './fixtures/docx';
import { ZipArchive } from '../src/zip';
import { expect, test } from './helpers';

// Formatting in a document that tracks its changes: bold, alignment, lists
// and styles are tracked changes like typing is (the saved file records what
// the formatting was), they are listed, rejected and accepted like the
// others. (Tables are tracked changes as well: e2e/doc-track-tables.spec.ts.)

const STYLES =
  DOC_STYLES +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/></w:rPr></w:style>';
const TRACK = '<w:trackRevisions/>';
const NOW = '2026-10-05T09:08:07.654Z';
const WHO = ' w:author="Me" w:date="2026-10-05T09:08:07Z"';

const sample = (body: string, settings = TRACK) => docx(body, { styles: STYLES, settings });

async function mountEdit(page: Page, bytes: Uint8Array | number[], options: { author?: string; trackChanges?: boolean } = {}): Promise<void> {
  if (!Array.isArray(bytes)) {
    await page.goto('/e2e/pages/doc-edit.html');
    await expect(page.locator('#state')).toHaveText('ready');
  }
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: 'Me', now: NOW, locale: 'en', ...options }] as [number[], object]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
const notices = (page: Page): Promise<{ message: string; kind: string }[]> => page.evaluate(() => (window as any).__notices.slice());
const button = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);
const commandKey = (page: Page) => page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));

/** The changes as the view lists them: [kind, author, text]. */
const revisions = (page: Page): Promise<[string, string, string][]> =>
  page.evaluate(() => ((window as any).__view.revisions() as any[]).map((revision) => [revision.kind, revision.author, revision.text] as [string, string, string]));

/** What the toolbar would show for the selection. */
const format = (page: Page): Promise<{ bold?: boolean; align?: string; list?: string; paraStyle?: string }> =>
  page.evaluate(() => {
    const now = (window as any).__view.selectionFormat;
    return { bold: now.bold, align: now.align, list: now.list, paraStyle: now.paraStyle };
  });

async function select(page: Page, path: number[], offset: number, to?: number): Promise<void> {
  await page.evaluate(
    ([where, from, end]) => {
      const view = (window as any).__view;
      view.setSelection({ anchor: { story: 'body', path: where, offset: from }, focus: { story: 'body', path: where, offset: end === null ? from : end } });
      view.focus();
    },
    [path, offset, to === undefined ? null : to] as [number[], number, number | null],
  );
  await expect(page.locator('#host .bym-doc-input')).toBeFocused();
  await settled(page);
}

const savedBytes = (page: Page): Promise<number[]> => page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));

async function savedXml(page: Page): Promise<string> {
  return new TextDecoder().decode(await ZipArchive.read(new Uint8Array(await savedBytes(page))).read('word/document.xml'));
}

test.describe('tracked changes of formatting', () => {
  test('bold over a selection is a tracked change: the text is bold, the list has one change, the file says what it was; rejecting brings it back', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    const before = await savedXml(page);
    await expect(button(page, 'bold')).toBeEnabled();
    await select(page, [0], 0, 5);
    await page.keyboard.press(`${await commandKey(page)}+b`);
    await settled(page);
    expect(await notices(page)).toEqual([]);
    expect((await format(page)).bold).toBe(true);
    await expect(button(page, 'bold')).toHaveAttribute('aria-pressed', 'true');
    expect(await revisions(page)).toEqual([['format', 'Me', 'First']]);
    expect(await page.evaluate(() => (window as any).__view.revisions().length)).toBe(1);
    const xml = await savedXml(page);
    expect(xml).toContain(`<w:r><w:rPr><w:b/><w:bCs/><w:rPrChange w:id="1"${WHO}><w:rPr/></w:rPrChange></w:rPr><w:t>First</w:t></w:r>`);
    // The change is the one at the caret: the button rejects it.
    await expect(button(page, 'rejectChange')).toBeEnabled();
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    await select(page, [0], 0, 5);
    expect((await format(page)).bold).toBe(false);
    expect(await savedXml(page)).not.toContain('rPrChange');
    // One step each: undo brings the change back, undo again takes the bold away, and the file is the one opened.
    const undo = `${await commandKey(page)}+z`;
    await page.keyboard.press(undo);
    await settled(page);
    expect(await revisions(page)).toEqual([['format', 'Me', 'First']]);
    await page.keyboard.press(undo);
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    expect(await savedXml(page)).toBe(before);
  });

  test('the toolbar offers alignment, lists and styles while changes are tracked, and each is a tracked change of the paragraph', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.') + p('Third.')));
    for (const command of ['alignLeft', 'alignCenter', 'alignRight', 'alignJustify', 'bulletList', 'numberedList', 'decreaseIndent', 'increaseIndent', 'bold', 'italic']) await expect(button(page, command)).toBeEnabled();
    await expect(page.locator('#host .bym-doc-toolbar select.bym-tb-style')).toBeEnabled();
    await select(page, [0], 3);
    await button(page, 'alignCenter').click();
    await settled(page);
    expect((await format(page)).align).toBe('center');
    expect(await revisions(page)).toEqual([['paragraphFormat', 'Me', '']]);
    await select(page, [1], 2);
    await button(page, 'bulletList').click();
    await settled(page);
    expect((await format(page)).list).toBe('bullet');
    await select(page, [2], 1);
    await page.locator('#host .bym-doc-toolbar select.bym-tb-style').selectOption('Heading1');
    await settled(page);
    expect((await format(page)).paraStyle).toBe('Heading1');
    expect(await notices(page)).toEqual([]);
    expect(await revisions(page)).toEqual([
      ['paragraphFormat', 'Me', ''],
      ['paragraphFormat', 'Me', ''],
      ['paragraphFormat', 'Me', ''],
    ]);
    const xml = await savedXml(page);
    expect(xml).toContain(`<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1"${WHO}><w:pPr/></w:pPrChange></w:pPr>`);
    expect(xml).toMatch(new RegExp(`<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="\\d+"/></w:numPr><w:pPrChange w:id="2"${WHO}><w:pPr/></w:pPrChange></w:pPr>`));
    expect(xml).toContain(`<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pPrChange w:id="3"${WHO}><w:pPr/></w:pPrChange></w:pPr>`);
    // Rejecting the one at the caret takes the style away again; the others stay.
    expect(await page.evaluate(() => (window as any).__view.rejectRevision(2))).toBe(true);
    await settled(page);
    expect((await format(page)).paraStyle).toBe('');
    expect((await revisions(page)).length).toBe(2);
    // All accepted: the formatting stays and no record does.
    expect(await page.evaluate(() => (window as any).__view.acceptAllRevisions())).toBe(true);
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    const accepted = await savedXml(page);
    expect(accepted).not.toContain('pPrChange');
    expect(accepted).toContain('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>');
  });

  test('a table is a tracked change too where changes are tracked, and the toolbar offers it', async ({ page }) => {
    // (What it is in the file, and accepting and rejecting it: e2e/doc-track-tables.spec.ts.)
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    await select(page, [0], 3);
    for (const command of ['insertTable', 'insertPicture', 'insertLink']) await expect(button(page, command)).toBeEnabled();
    // Rows and columns are those of a table: the caret is in none.
    await expect(button(page, 'rowsColumns')).toBeDisabled();
    expect(await page.evaluate(() => (window as any).__view.insertTable(2, 2))).toBe(true);
    expect(await notices(page)).toEqual([]);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(true);
    expect((await savedXml(page)).split('<w:trPr><w:ins ').length).toBe(3);
  });

  test('saved and opened again, the changes of formatting are still there, and rejecting all of them gives the formatting back', async ({ page }) => {
    await mountEdit(page, await sample(p(r('First', '<w:i/>') + r(' paragraph.')) + p('Second one.', '<w:jc w:val="right"/>')));
    await select(page, [0], 2, 9);
    await page.keyboard.press(`${await commandKey(page)}+b`);
    await settled(page);
    await select(page, [1], 0);
    await button(page, 'alignCenter').click();
    await settled(page);
    const made = await revisions(page);
    expect(made).toEqual([
      ['format', 'Me', 'rst par'],
      ['paragraphFormat', 'Me', ''],
    ]);
    await mountEdit(page, await savedBytes(page));
    expect(await revisions(page)).toEqual(made);
    await select(page, [0], 2, 9);
    expect((await format(page)).bold).toBe(true);
    const xml = await savedXml(page);
    expect(xml).toContain(`<w:rPrChange w:id="1"${WHO}><w:rPr><w:i/></w:rPr></w:rPrChange>`);
    expect(xml).toContain(`<w:pPrChange w:id="3"${WHO}><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange>`);
    expect(await page.evaluate(() => (window as any).__view.rejectAllRevisions())).toBe(true);
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    expect((await format(page)).bold).toBe(false);
    await select(page, [1], 0);
    expect((await format(page)).align).toBe('right');
    const rejected = await savedXml(page);
    expect(rejected).not.toContain('Change');
    expect(rejected).toContain('<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>First</w:t></w:r><w:r><w:t xml:space="preserve"> paragraph.</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="right"/></w:pPr>');
  });

  test('at a caret in an empty paragraph: bold turned on and off again leaves no change; turned on and typed into, the mark has one', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('') + p('Last.')));
    const before = await savedXml(page);
    const key = `${await commandKey(page)}+b`;
    await select(page, [1], 0);
    await page.keyboard.press(key);
    await settled(page);
    expect((await format(page)).bold).toBe(true);
    expect(await revisions(page)).toEqual([['format', 'Me', '']]);
    await page.keyboard.press(key);
    await settled(page);
    expect((await format(page)).bold).toBe(false);
    expect(await revisions(page)).toEqual([]);
    const after = await savedXml(page);
    expect(after).not.toContain('rPrChange');
    expect(after).not.toContain('w:val="0"');
    expect(after.replace('<w:p></w:p>', '<w:p/>')).toBe(before.replace('<w:p></w:p>', '<w:p/>'));
    // With the toolbar the same.
    await button(page, 'bold').click();
    await settled(page);
    await button(page, 'bold').click();
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    // Turned on and typed into: the text is the author's insertion, bold; the mark says what it was.
    await page.keyboard.press(key);
    await page.keyboard.type('new');
    await settled(page);
    expect((await revisions(page)).map((revision) => revision[0]).sort()).toEqual(['format', 'insert']);
    const typed = await savedXml(page);
    expect(typed).toMatch(new RegExp(`<w:p><w:pPr><w:rPr><w:b/><w:bCs/><w:rPrChange w:id="\\d+"${WHO}><w:rPr/></w:rPrChange></w:rPr></w:pPr><w:ins w:id="\\d+"${WHO}><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>new</w:t></w:r></w:ins></w:p>`));
  });

  test('at a caret in text, and in a document that does not track: bold on and off again leaves no trace in the file', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p(''), ''));
    const before = await savedXml(page);
    const key = `${await commandKey(page)}+b`;
    for (const [path, offset] of [
      [[1], 0],
      [[0], 5],
    ] as [number[], number][]) {
      await select(page, path, offset);
      await page.keyboard.press(key);
      await settled(page);
      expect((await format(page)).bold).toBe(true);
      await page.keyboard.press(key);
      await settled(page);
      expect((await format(page)).bold).toBe(false);
    }
    await page.keyboard.type('x');
    await settled(page);
    const xml = await savedXml(page);
    expect(xml).not.toContain('w:val="0"');
    expect(xml).not.toMatch(/<w:b[ /]/);
    expect(xml).toContain('<w:t>Firstx paragraph.</w:t>');
    expect(before).not.toMatch(/<w:b[ /]/);
  });

  test('Tab and Shift+Tab in a list item change its level as a tracked change, which can be rejected', async ({ page }) => {
    const numbering =
      '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    const numPr = (level: number) => `<w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="1"/></w:numPr>`;
    await mountEdit(page, await docx(p('Item one.', numPr(0)) + p('Item two.', numPr(0)), { styles: STYLES, settings: TRACK, numbering }));
    await select(page, [1], 0);
    await page.keyboard.press('Tab');
    await settled(page);
    expect(await notices(page)).toEqual([]);
    expect(await page.evaluate(() => (window as any).__view.selectionFormat.level)).toBe(1);
    expect(await revisions(page)).toEqual([['paragraphFormat', 'Me', '']]);
    expect(await savedXml(page)).toContain(`<w:p><w:pPr>${numPr(1)}<w:pPrChange w:id="1"${WHO}><w:pPr>${numPr(0)}</w:pPr></w:pPrChange></w:pPr>`);
    // Back again: the author's own record of this session goes.
    await page.keyboard.press('Shift+Tab');
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    expect(await savedXml(page)).not.toContain('pPrChange');
    // Deeper once more, and rejected.
    await page.keyboard.press('Tab');
    await settled(page);
    expect((await revisions(page)).length).toBe(1);
    await expect(button(page, 'rejectChange')).toBeEnabled();
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    await select(page, [1], 0);
    expect(await page.evaluate(() => (window as any).__view.selectionFormat.level)).toBe(0);
    const xml = await savedXml(page);
    expect(xml).not.toContain('pPrChange');
    expect(xml).toContain(`<w:p><w:pPr>${numPr(0)}</w:pPr><w:r><w:t>Item two.</w:t></w:r></w:p>`);
  });

  test('a document that does not track is formatted without a record', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.'), ''));
    await select(page, [0], 0, 5);
    await page.keyboard.press(`${await commandKey(page)}+b`);
    await settled(page);
    await button(page, 'alignCenter').click();
    await settled(page);
    expect(await revisions(page)).toEqual([]);
    expect(await savedXml(page)).not.toContain('Change');
  });
});
