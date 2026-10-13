import type { Page } from '@playwright/test';
import { p, r, tbl, tc, tr } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES, png, RED, BLUE } from './fixtures/docx';
import { ZipArchive } from '../src/zip';
import { expect, test } from './helpers';

// Tables, pictures, links and pasted content in a document that tracks its
// changes: rows, columns and tables that are put in or taken out, a picture,
// a link and a rich paste are tracked changes in the saved file, which the
// editor accepts and rejects; the toolbar offers all of them; and what still
// cannot be tracked is refused, and said.

const STYLES = DOC_STYLES + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const TRACK = '<w:trackRevisions/>';
const NOW = '2026-10-05T09:08:07.654Z';
const AUTHOR = 'Me';
const REFUSED = 'Changes are tracked in this document, and this kind of change cannot be tracked yet; nothing was changed';

const TABLE = tbl(tr(tc('a1') + tc('b1')) + tr(tc('a2') + tc('b2')), [2000, 2000]);
const sample = (body: string) => docx(body, { styles: STYLES, settings: TRACK });

async function mountEdit(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: AUTHOR, now: NOW }] as [number[], { author: string; now: string }]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

/** Lets the view bring the screen up to date: its updates run in animation frames. */
const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

const paragraphs = (page: Page) => page.evaluate(() => ((window as any).BymOffice.documentText((window as any).__view.document) as string).split('\n'));
const notices = (page: Page): Promise<{ message: string; kind: string }[]> => page.evaluate(() => (window as any).__notices.slice());
const button = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);

/** Puts a selection (a caret when the two ends are one) into the document and the focus into the editor. */
async function select(page: Page, path: number[], offset: number, toPath: number[] = path, toOffset: number = offset): Promise<void> {
  await page.evaluate(
    ([a, from, b, to]) => {
      const view = (window as any).__view;
      view.setSelection({ anchor: { story: 'body', path: a, offset: from }, focus: { story: 'body', path: b, offset: to } });
      view.focus();
    },
    [path, offset, toPath, toOffset] as [number[], number, number[], number],
  );
  await expect(page.locator('#host .bym-doc-input')).toBeFocused();
  await settled(page);
}

/** The document part of the file the editor saves. */
async function savedXml(page: Page): Promise<string> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  return new TextDecoder().decode(await ZipArchive.read(new Uint8Array(bytes)).read('word/document.xml'));
}

/** The changes as the view lists them: [kind, author, rows it holds]. */
const revisions = (page: Page): Promise<[string, string | undefined, number][]> =>
  page.evaluate(() => ((window as any).__view.revisions() as any[]).map((revision) => [revision.kind, revision.author, revision.rows] as [string, string | undefined, number]));

/** How many rows the first table of the body has, and how many cells each of them. */
const tableShape = (page: Page): Promise<number[]> =>
  page.evaluate(() => {
    const table = ((window as any).__view.document.body.blocks as any[]).find((block) => block.type === 'table');
    return table ? (table.rows as any[]).map((row) => row.cells.length) : [];
  });

const count = (text: string, part: string): number => text.split(part).length - 1;

/** Fires a copy or a paste at the field for typing, the way the browser does for the keys. */
const clipboard = (page: Page, kind: 'copy' | 'paste', give: { text?: string; html?: string } = {}) =>
  page.evaluate(
    ([type, content]) => {
      const target = document.querySelector('#host .bym-doc-input')!;
      const data = new DataTransfer();
      if (content.text !== undefined) data.setData('text/plain', content.text);
      if (content.html !== undefined) data.setData('text/html', content.html);
      let event: Event | undefined;
      try {
        event = new ClipboardEvent(type, { clipboardData: data, bubbles: true, cancelable: true });
      } catch {
        event = undefined;
      }
      if (!event || (event as ClipboardEvent).clipboardData !== data) {
        event = new Event(type, { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
      }
      target.dispatchEvent(event);
      return { text: data.getData('text/plain'), html: data.getData('text/html') };
    },
    [kind, give] as ['copy' | 'paste', { text?: string; html?: string }],
  );

test.describe('tracked changes of tables, pictures, links and pasted content', () => {
  test('the toolbar offers tables, pictures, links, rows and columns where changes are tracked', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + TABLE + p('Last one.')));
    await select(page, [0], 3);
    for (const command of ['insertTable', 'insertPicture', 'insertLink', 'pageBreak', 'bold']) await expect(button(page, command)).toBeEnabled();
    // Rows and columns are those of the table the caret is in.
    await expect(button(page, 'rowsColumns')).toBeDisabled();
    await select(page, [1, 0, 0, 0], 1);
    await expect(button(page, 'rowsColumns')).toBeEnabled();
    expect(await notices(page)).toEqual([]);
  });

  test('an inserted row is a tracked change; rejected, the table is what it was', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + TABLE + p('Last one.')));
    const before = await paragraphs(page);
    await select(page, [1, 0, 0, 0], 1);
    expect(await page.evaluate(() => (window as any).__view.insertRow('below'))).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([2, 2, 2]);
    expect(await revisions(page)).toEqual([['insert', AUTHOR, 1]]);
    const xml = await savedXml(page);
    expect(count(xml, '<w:trPr><w:ins ')).toBe(1);
    expect(xml).toContain(`<w:trPr><w:ins w:id="1" w:author="${AUTHOR}" w:date="2026-10-05T09:08:07Z"/></w:trPr>`);

    expect(await page.evaluate(() => (window as any).__view.rejectRevision(0))).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([2, 2]);
    expect(await revisions(page)).toEqual([]);
    expect(await paragraphs(page)).toEqual(before);
    expect(await savedXml(page)).not.toContain('<w:ins ');
    expect(await notices(page)).not.toContainEqual({ message: REFUSED, kind: 'info' });
  });

  test('a deleted column stays, struck, until the change is accepted', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + TABLE + p('Last one.')));
    await select(page, [1, 0, 1, 0], 0);
    expect(await page.evaluate(() => (window as any).__view.deleteColumn())).toBe(true);
    await settled(page);
    // The cells are still there, and so is their text, as deleted text.
    expect(await tableShape(page)).toEqual([2, 2]);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'a1', '', 'a2', '', 'Last one.']);
    const xml = await savedXml(page);
    expect(count(xml, '<w:cellDel ')).toBe(2);
    expect(xml).toContain('<w:delText>b1</w:delText>');
    // The column is one change, whatever the number of rows, and is said to be a column.
    expect(await revisions(page)).toEqual([['delete', AUTHOR, 0]]);
    expect(await page.evaluate(() => (window as any).__view.revisions()[0].columns)).toBe(1);
    expect(await page.evaluate(() => (window as any).__view.describeRevision((window as any).__view.revisions()[0]))).toContain('Column deleted by Me');

    expect(await page.evaluate(() => (window as any).__view.acceptAllRevisions())).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([1, 1]);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'a1', 'a2', 'Last one.']);
    expect(await revisions(page)).toEqual([]);
    const accepted = await savedXml(page);
    expect(accepted).not.toContain('cellDel');
    expect(count(accepted, '<w:gridCol ')).toBe(1);
  });

  test('an inserted column is one change with the record of the grid; rejected, the table is what it was', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + TABLE + p('Last one.')));
    const before = await savedXml(page);
    await select(page, [1, 0, 0, 0], 0);
    expect(await page.evaluate(() => (window as any).__view.insertColumn('right'))).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([3, 3]);
    expect(await revisions(page)).toEqual([['insert', AUTHOR, 0]]);
    expect(await page.evaluate(() => (window as any).__view.describeRevision((window as any).__view.revisions()[0]))).toContain('Column inserted by Me');
    const xml = await savedXml(page);
    expect(count(xml, '<w:cellIns ')).toBe(2);
    expect(count(xml, '<w:tblGridChange ')).toBe(1);

    expect(await page.evaluate(() => (window as any).__view.rejectRevision(0))).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([2, 2]);
    expect(await revisions(page)).toEqual([]);
    const rejected = await savedXml(page);
    expect(rejected).not.toContain('cellIns');
    expect(rejected).not.toContain('tblGridChange');
    expect(count(rejected, '<w:gridCol ')).toBe(count(before, '<w:gridCol '));
  });

  test('an inserted table is a table of inserted rows, and one change', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Last one.')));
    await select(page, [0], 16);
    expect(await page.evaluate(() => (window as any).__view.insertTable(2, 3))).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([3, 3]);
    expect((await revisions(page)).map((revision) => revision[0])).toEqual(['insert']);
    const xml = await savedXml(page);
    expect(count(xml, '<w:trPr><w:ins ')).toBe(2);
    expect(await notices(page)).toEqual([]);

    expect(await page.evaluate(() => (window as any).__view.rejectAllRevisions())).toBe(true);
    await settled(page);
    expect(await tableShape(page)).toEqual([]);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Last one.']);
  });

  test('an inserted picture is a tracked insertion; rejected, it is gone', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Last one.')));
    await select(page, [0], 5);
    const bytes = Array.from(png(40, 20, RED, BLUE));
    expect(await page.evaluate((input) => (window as any).__view.insertPicture(new Uint8Array(input)), bytes)).toBe(true);
    await settled(page);
    const xml = await savedXml(page);
    expect(xml).toMatch(/<w:ins w:id="1" w:author="Me"[^>]*><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:drawing>/);
    expect((await revisions(page)).map((revision) => revision[0])).toEqual(['insert']);

    expect(await page.evaluate(() => (window as any).__view.rejectRevision(0))).toBe(true);
    await settled(page);
    expect(await savedXml(page)).not.toContain('<w:drawing');
    expect(await revisions(page)).toEqual([]);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Last one.']);
  });

  test('a link over text is the text deleted and inserted inside the link; rejected, the text is plain again', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Last one.')));
    await select(page, [0], 0, [0], 5);
    expect(await page.evaluate(() => (window as any).__view.insertLink('https://example.org/'))).toBe(true);
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Last one.']);
    const xml = await savedXml(page);
    expect(xml).toMatch(/<w:del [^>]*><w:r><w:delText>First<\/w:delText><\/w:r><\/w:del><w:hyperlink r:id="rId\d+"><w:ins [^>]*><w:r>/);
    expect((await revisions(page)).map((revision) => revision[0])).toEqual(['delete', 'insert']);

    expect(await page.evaluate(() => (window as any).__view.rejectAllRevisions())).toBe(true);
    await settled(page);
    const rejected = await savedXml(page);
    expect(rejected).not.toContain('<w:hyperlink');
    expect(rejected).not.toContain('<w:del ');
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Last one.']);
  });

  test('a copy is pasted whole as a tracked insertion, formatting and table included; all rejected, nothing of it is left', async ({ page }) => {
    await mountEdit(page, await sample(p(r('Bold', '<w:b/>') + r(' and plain.')) + TABLE + p('Last one.')));
    const before = await paragraphs(page);
    // From the start of the first paragraph into the last one: the bold text, the table and a word.
    await select(page, [0], 0, [2], 4);
    const copied = await clipboard(page, 'copy');
    expect(copied.text).toContain('Bold and plain.');
    await select(page, [2], 9);
    await clipboard(page, 'paste', copied);
    await settled(page);
    // No notice: nothing stayed behind.
    expect(await notices(page)).toEqual([]);
    const xml = await savedXml(page);
    expect(count(xml, '<w:tbl>')).toBe(2);
    expect(count(xml, '<w:trPr><w:ins ')).toBe(2);
    expect(xml).toMatch(/<w:ins [^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>Bold<\/w:t><\/w:r>/);
    expect((await revisions(page)).every((revision) => revision[0] === 'insert' && revision[1] === AUTHOR)).toBe(true);

    expect(await page.evaluate(() => (window as any).__view.rejectAllRevisions())).toBe(true);
    await settled(page);
    expect(await paragraphs(page)).toEqual(before);
    expect(await revisions(page)).toEqual([]);
    const rejected = await savedXml(page);
    expect(count(rejected, '<w:tbl>')).toBe(1);
    expect(rejected).not.toContain('<w:ins ');
  });

  test('what still cannot be tracked is refused with a notice: a column through a merged cell', async ({ page }) => {
    const merged = tbl(tr(tc('wide', '<w:gridSpan w:val="2"/>')) + tr(tc('a2') + tc('b2')), [2000, 2000]);
    await mountEdit(page, await sample(p('First paragraph.') + merged + p('Last one.')));
    const before = await savedXml(page);
    await select(page, [1, 1, 0, 0], 0);
    expect(await page.evaluate(() => (window as any).__view.insertColumn('right'))).toBe(false);
    await settled(page);
    expect(await notices(page)).toEqual([{ message: REFUSED, kind: 'info' }]);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    expect(await savedXml(page)).toBe(before);
  });
});
