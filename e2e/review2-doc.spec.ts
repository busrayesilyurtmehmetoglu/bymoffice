// Second review of documents that track their changes: a copy that holds no text (a picture
// alone). It used to be refused, with a notice, while a paste was tracked as text only; a paste
// is tracked whole now, so the picture goes in as a tracked insertion, which rejecting takes out.
import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES, picture, png, RED, BLUE } from './fixtures/docx';
import { ZipArchive } from '../src/zip';
import { expect, test } from './helpers';

const STYLES = DOC_STYLES + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';

const sample = () =>
  docx(p(r('ab') + picture(120, 60) + r('cd')) + p('Second one.'), {
    styles: STYLES,
    settings: '<w:trackRevisions/>',
    rels: [['rIdImg', 'image', 'media/image1.png']],
    parts: [['word/media/image1.png', png(40, 20, RED, BLUE)]],
  });

async function mountEdit(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate((input) => (window as any).__mountEdit(input, { author: 'Me', now: '2026-10-05T09:08:07.654Z', locale: 'en' }), Array.from(bytes));
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

/** Selects from one offset of a paragraph to another; `picture`: from before what stands at the first offset without a length of its own. */
async function select(page: Page, path: number[], offset: number, to: number, picture = false): Promise<void> {
  await page.evaluate(
    ([where, from, end, before]) => {
      const view = (window as any).__view;
      const anchor: Record<string, unknown> = { story: 'body', path: where, offset: from };
      if (before) anchor.before = true;
      view.setSelection({ anchor, focus: { story: 'body', path: where, offset: end } });
      view.focus();
    },
    [path, offset, to, picture] as [number[], number, number, boolean],
  );
  await expect(page.locator('#host .bym-doc-input')).toBeFocused();
  await settled(page);
}

function clipboard(page: Page, kind: 'copy' | 'paste', give: { text?: string; html?: string } = {}): Promise<{ text: string; html: string }> {
  return page.evaluate(
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
}

/** The document part of the file the editor saves. */
async function savedXml(page: Page): Promise<string> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  return new TextDecoder().decode(await ZipArchive.read(new Uint8Array(bytes)).read('word/document.xml'));
}

const count = (text: string, part: string): number => text.split(part).length - 1;

test('review 2: in a document that tracks changes, a copy of a picture alone is pasted as a tracked insertion, which rejecting takes out again', async ({ page }) => {
  await mountEdit(page, await sample());
  // The picture and nothing else.
  await select(page, [0], 2, 2, true);
  const copied = await clipboard(page, 'copy');
  expect(copied.text).toBe('');
  expect(copied.html).not.toBe('');
  await select(page, [1], 0, 0);
  const before = await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__view.document));
  expect(count(await savedXml(page), '<w:drawing>')).toBe(1);
  await clipboard(page, 'paste', copied);
  await settled(page);
  // A picture is no text: the text is what it was, and nothing had to be said.
  expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__view.document))).toBe(before);
  expect(await page.evaluate(() => (window as any).__notices.slice())).toEqual([]);
  expect(await page.evaluate(() => (window as any).__view.canUndo)).toBe(true);
  // The second picture is in the second paragraph, inside an insertion of the editor, and it is the one change.
  const xml = await savedXml(page);
  expect(count(xml, '<w:drawing>')).toBe(2);
  expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Me" w:date="2026-10-05T09:08:07Z"><w:r>(<w:rPr>.*?<\/w:rPr>)?<w:drawing>/);
  expect(count(xml, '<w:ins ')).toBe(1);
  expect(await page.evaluate(() => ((window as any).__view.revisions() as any[]).map((revision) => [revision.kind, revision.author]))).toEqual([['insert', 'Me']]);
  // Rejected, the document has the one picture it had.
  expect(await page.evaluate(() => (window as any).__view.rejectAllRevisions())).toBe(true);
  await settled(page);
  const rejected = await savedXml(page);
  expect(count(rejected, '<w:drawing>')).toBe(1);
  expect(rejected).not.toContain('<w:ins ');
  expect(await page.evaluate(() => (window as any).__view.revisions().length)).toBe(0);
});
