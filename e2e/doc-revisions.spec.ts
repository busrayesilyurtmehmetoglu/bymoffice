import type { Page } from '@playwright/test';
import { p } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES } from './fixtures/docx';
import { ZipArchive } from '../src/zip';
import { expect, test } from './helpers';

// Accepting and rejecting tracked changes, stepping through them, and turning
// the tracking on and off, from the toolbar of the document editor and from
// the methods of the view.

interface MountOptions {
  readOnly?: boolean;
  locale?: string;
  trackChanges?: boolean;
  author?: string;
  now?: string;
}

const STYLES = DOC_STYLES + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const NOW = '2026-10-08T09:08:07.654Z';
/** A name that would be markup if it were ever put into the page as HTML. */
const EVIL = 'Eve <img src=x onerror="window.__pwned=1">';
const EVIL_XML = 'Eve &lt;img src=x onerror=&quot;window.__pwned=1&quot;>';

const who = (id: number, author: string, date = '2026-10-01T09:00:00Z') => ` w:id="${id}" w:author="${author}" w:date="${date}"`;
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const ins = (id: number, text: string, author = 'Ann') => `<w:ins${who(id, author)}>${t(text)}</w:ins>`;
const del = (id: number, text: string, author = 'Bob') => `<w:del${who(id, author)}><w:r><w:delText xml:space="preserve">${text}</w:delText></w:r></w:del>`;

/** Three changes: an insertion of Ann, a deletion of Bob, an insertion of a third author in the second paragraph. */
const BODY = (third = 'Cid') => p(t('One ') + ins(1, 'two ') + del(2, 'three ') + t('four.')) + p(t('Five ') + ins(3, 'six ', third) + t('seven.')) + p(t('Last paragraph.'));
const sample = (body: string, settings: string | null = null) => docx(body, settings === null ? { styles: STYLES } : { styles: STYLES, settings });
const protect = (edit: string) => `<w:documentProtection w:edit="${edit}" w:enforcement="1"/>`;

async function mountEdit(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: 'Me', now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
const paragraphs = (page: Page) => page.evaluate(() => ((window as any).BymOffice.documentText((window as any).__view.document) as string).split('\n'));
const notices = (page: Page): Promise<{ message: string; kind: string }[]> => page.evaluate(() => (window as any).__notices.slice());
const lastNotice = async (page: Page): Promise<string> => {
  const all = await notices(page);
  return all.length ? all[all.length - 1].message : '';
};
const button = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);
const input = (page: Page) => page.locator('#host .bym-doc-input');
const live = (page: Page) => page.locator('#host .bym-doc-live');

/** The changes as the view lists them: [kind, author, text]. */
const revisions = (page: Page): Promise<[string, string, string][]> =>
  page.evaluate(() => ((window as any).__view.revisions() as any[]).map((revision) => [revision.kind, revision.author, revision.text] as [string, string, string]));

const selection = (page: Page): Promise<[number[], number, number[], number] | undefined> =>
  page.evaluate(() => {
    const range = (window as any).__view.getSelection();
    return range ? [range.anchor.path, range.anchor.offset, range.focus.path, range.focus.offset] : undefined;
  });

async function select(page: Page, path: number[], offset: number, to?: [number[], number]): Promise<void> {
  await page.evaluate(
    ([where, from, end]) => {
      const view = (window as any).__view;
      view.setSelection({ anchor: { story: 'body', path: where, offset: from }, focus: end ? { story: 'body', path: end[0], offset: end[1] } : { story: 'body', path: where, offset: from } });
      view.focus();
    },
    [path, offset, to || null] as [number[], number, [number[], number] | null],
  );
  await expect(input(page)).toBeFocused();
  await settled(page);
}

async function savedPart(page: Page, name: string): Promise<string | undefined> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  const zip = ZipArchive.read(new Uint8Array(bytes));
  return zip.has(name) ? new TextDecoder().decode(await zip.read(name)) : undefined;
}

async function reopen(page: Page, options: MountOptions = {}): Promise<void> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  await page.evaluate(([data, opts]) => (window as any).__mountEdit(data, opts), [bytes, { author: 'Me', now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

const commandKey = (page: Page) => page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));

test.describe('tracked changes: accepting, rejecting, stepping', () => {
  test('the toolbar has the group of the changes, labelled in English and Turkish, enabled by what the document and the caret hold', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    for (const [command, label] of [
      ['trackChanges', 'Track changes'],
      ['previousChange', 'Previous change'],
      ['nextChange', 'Next change'],
      ['acceptChange', 'Accept change'],
      ['rejectChange', 'Reject change'],
      ['changesMenu', 'All changes'],
    ]) {
      await expect(button(page, command)).toHaveAttribute('aria-label', label);
      await expect(button(page, command)).toHaveAttribute('title', label);
    }
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-pressed', 'false');
    await expect(button(page, 'changesMenu')).toHaveAttribute('aria-haspopup', 'menu');
    expect(await revisions(page)).toEqual([
      ['insert', 'Ann', 'two '],
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
    // The caret is in plain text: there is a change to step to, none to accept.
    await select(page, [2], 3);
    await expect(button(page, 'nextChange')).toBeEnabled();
    await expect(button(page, 'previousChange')).toBeEnabled();
    await expect(button(page, 'changesMenu')).toBeEnabled();
    await expect(button(page, 'acceptChange')).toBeDisabled();
    await expect(button(page, 'rejectChange')).toBeDisabled();
    // Inside the insertion they can be pressed.
    await select(page, [0], 6);
    await expect(button(page, 'acceptChange')).toBeEnabled();
    await expect(button(page, 'rejectChange')).toBeEnabled();

    await page.evaluate(() => (window as any).__mounted.editor.setLocale('tr'));
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-label', 'Değişiklikleri izle');
    await expect(button(page, 'acceptChange')).toHaveAttribute('aria-label', 'Değişikliği kabul et');
    await expect(button(page, 'rejectChange')).toHaveAttribute('aria-label', 'Değişikliği reddet');
    await expect(button(page, 'nextChange')).toHaveAttribute('aria-label', 'Sonraki değişiklik');
    await expect(button(page, 'previousChange')).toHaveAttribute('aria-label', 'Önceki değişiklik');
    await expect(button(page, 'changesMenu')).toHaveAttribute('aria-label', 'Tüm değişiklikler');
  });

  test('a document without changes: nothing to step to or resolve, tracking can still be turned on', async ({ page }) => {
    await mountEdit(page, await sample(p('Plain text.') + p('More.')));
    await select(page, [0], 2);
    for (const command of ['previousChange', 'nextChange', 'acceptChange', 'rejectChange', 'changesMenu']) await expect(button(page, command)).toBeDisabled();
    await expect(button(page, 'trackChanges')).toBeEnabled();
    expect(await page.evaluate(() => [(window as any).__view.revisions().length, (window as any).__view.goToRevision('next'), (window as any).__view.acceptAllRevisions()])).toEqual([0, false, false]);
    expect(await lastNotice(page)).toBe('The document has no tracked changes');
  });

  test('next and previous select the changes in turn, say who changed what, and wrap around', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    await select(page, [0], 0);
    await button(page, 'nextChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[0], 4, [0], 8]);
    await expect(live(page)).toHaveText('Change 1 of 3. Inserted by Ann, 2026-10-01 09:00: two ');
    // The focus is back in the document.
    await expect(input(page)).toBeFocused();
    await button(page, 'nextChange').click();
    await settled(page);
    // A deletion takes no room: the caret stands where it is.
    expect(await selection(page)).toEqual([[0], 8, [0], 8]);
    await expect(live(page)).toHaveText('Change 2 of 3. Deleted by Bob, 2026-10-01 09:00: three ');
    await button(page, 'nextChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[1], 5, [1], 9]);
    await button(page, 'nextChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[0], 4, [0], 8]);
    await expect(live(page)).toHaveText('Back at the first change. Change 1 of 3. Inserted by Ann, 2026-10-01 09:00: two ');
    await button(page, 'previousChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[1], 5, [1], 9]);
    await expect(live(page)).toHaveText('At the last change. Change 3 of 3. Inserted by Cid, 2026-10-01 09:00: six ');
    await button(page, 'previousChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[0], 8, [0], 8]);
    // By number, from the methods of the view.
    expect(await page.evaluate(() => (window as any).__view.goToRevision(2))).toBe(true);
    expect(await selection(page)).toEqual([[1], 5, [1], 9]);
    expect(await page.evaluate(() => [(window as any).__view.goToRevision(7), (window as any).__view.goToRevision(-1)])).toEqual([false, false]);
  });

  test('accept and reject resolve the change at the caret and go on to the next one; undo and redo take it back and make it again', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    await select(page, [0], 6);
    await button(page, 'acceptChange').click();
    await settled(page);
    expect(await revisions(page)).toEqual([
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
    expect(await paragraphs(page)).toEqual(['One two four.', 'Five six seven.', 'Last paragraph.']);
    expect((await notices(page)).map((notice) => notice.message)).toContain('Accepted 1 change. 2 changes left.');
    // On to the next change: the deletion, right there.
    expect(await selection(page)).toEqual([[0], 8, [0], 8]);
    await expect(input(page)).toBeFocused();
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'Five six seven.', 'Last paragraph.']);
    expect((await notices(page)).map((notice) => notice.message)).toContain('Rejected 1 change. 1 change left.');
    expect(await selection(page)).toEqual([[1], 5, [1], 9]);
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'Five seven.', 'Last paragraph.']);
    expect(await revisions(page)).toEqual([]);
    expect(await lastNotice(page)).toBe('Rejected 1 change. No changes left.');
    for (const command of ['previousChange', 'nextChange', 'acceptChange', 'rejectChange', 'changesMenu']) await expect(button(page, command)).toBeDisabled();
    const xml = (await savedPart(page, 'word/document.xml')) as string;
    expect(xml).not.toMatch(/<w:(ins|del|delText)\b/);
    expect(xml).toContain('<w:t>One two three four.</w:t>');

    // Three steps back, one at a time; and forward again.
    const key = await commandKey(page);
    for (let i = 0; i < 3; i++) await page.keyboard.press(`${key}+z`);
    await settled(page);
    expect(await revisions(page)).toEqual([
      ['insert', 'Ann', 'two '],
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
    expect(await paragraphs(page)).toEqual(['One two four.', 'Five six seven.', 'Last paragraph.']);
    await button(page, 'redo').click();
    await settled(page);
    expect(await revisions(page)).toEqual([
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
  });

  test('a selection over several changes resolves all of them at once, as one step', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    await select(page, [0], 5, [[1], 7]);
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await paragraphs(page)).toEqual(['One three four.', 'Five seven.', 'Last paragraph.']);
    expect(await revisions(page)).toEqual([]);
    expect(await lastNotice(page)).toBe('Rejected 3 changes. No changes left.');
    await button(page, 'undo').click();
    await settled(page);
    expect((await revisions(page)).length).toBe(3);
  });

  test('"accept all" and "reject all" from the menu; saved and opened again, no change is left', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    await select(page, [2], 4);
    await button(page, 'changesMenu').click();
    const menu = page.locator('#host .bym-popover');
    await expect(menu).toBeVisible();
    await expect(menu.locator('.bym-pop-item')).toHaveText(['Accept all changes', 'Reject all changes']);
    await menu.locator('.bym-pop-item', { hasText: 'Accept all changes' }).click();
    await settled(page);
    await expect(menu).toHaveCount(0);
    expect(await paragraphs(page)).toEqual(['One two four.', 'Five six seven.', 'Last paragraph.']);
    expect(await lastNotice(page)).toBe('Accepted 3 changes. No changes left.');
    // The caret stays where it was.
    expect(await selection(page)).toEqual([[2], 4, [2], 4]);
    await expect(input(page)).toBeFocused();
    expect((await savedPart(page, 'word/document.xml')) as string).not.toMatch(/<w:(ins|del|delText)\b/);
    await reopen(page);
    expect(await revisions(page)).toEqual([]);
    expect(await paragraphs(page)).toEqual(['One two four.', 'Five six seven.', 'Last paragraph.']);

    await mountEdit(page, await sample(BODY()));
    await select(page, [2], 4);
    await button(page, 'changesMenu').click();
    await page.locator('#host .bym-popover .bym-pop-item', { hasText: 'Reject all changes' }).click();
    await settled(page);
    expect(await paragraphs(page)).toEqual(['One three four.', 'Five seven.', 'Last paragraph.']);
    await button(page, 'undo').click();
    await settled(page);
    expect((await revisions(page)).length).toBe(3);
    // Undone, the file is the one that was opened.
    expect((await savedPart(page, 'word/document.xml')) as string).toContain('<w:ins w:id="1" w:author="Ann"');
  });

  test('the methods of the view: by number, by the change itself, and all; a change that is gone is refused', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      const before = (window as any).__changes;
      const list = view.revisions();
      const out: unknown[] = [];
      out.push(view.rejectRevision(list[2]));
      out.push(view.acceptRevision(0));
      // The description is older than the edits: the change it names is no longer there.
      out.push(view.acceptRevision(list[0]));
      out.push(view.acceptRevision(5));
      out.push(view.revisions().map((revision: any) => revision.kind));
      out.push(view.rejectAllRevisions());
      out.push(view.revisions().length);
      out.push((window as any).__changes - before);
      return out;
    });
    expect(result).toEqual([true, true, false, false, ['delete'], true, 0, 3]);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'Five seven.', 'Last paragraph.']);
    expect((await notices(page)).map((notice) => notice.message)).toContain('The change is no longer in the document');
  });

  test('the name of an author is text wherever it is shown', async ({ page }) => {
    await mountEdit(page, await sample(p(t('One ') + `<w:ins${who(1, EVIL_XML)}>${t('two ')}</w:ins>` + t('three.'))));
    expect(await revisions(page)).toEqual([['insert', EVIL, 'two ']]);
    await select(page, [0], 0);
    await button(page, 'nextChange').click();
    await settled(page);
    await expect(live(page)).toHaveText(`Change 1 of 1. Inserted by ${EVIL}, 2026-10-01 09:00: two `);
    await expect(button(page, 'acceptChange')).toHaveAttribute('title', `Accept change: Inserted by ${EVIL}, 2026-10-01 09:00: two `);
    expect(await page.evaluate(() => [document.querySelectorAll('#host img').length, (window as any).__pwned])).toEqual([0, undefined]);
  });

  test('the pages say what the change under the pointer is', async ({ page, isMobile }) => {
    test.skip(isMobile, 'a tool tip needs a pointer that hovers');
    await mountEdit(page, await sample(BODY(EVIL_XML)));
    await select(page, [2], 0);
    const pointAt = (path: number[], offset: number) =>
      page.evaluate(
        ([where, at]) => {
          const view = (window as any).__view;
          const box = view.lines.caret({ story: 'body', path: where, offset: at }, 'start', view.currentPage - 1);
          const element = document.querySelector(`#host .bym-doc-page[data-page="${box.page + 1}"]`)!;
          const rect = element.getBoundingClientRect();
          const laid = view.pages[box.page];
          return { x: rect.left + (box.x * rect.width) / laid.width, y: rect.top + ((box.top + box.height / 2) * rect.height) / laid.height };
        },
        [path, offset] as [number[], number],
      );
    const strip = page.locator('#host .bym-doc-strip');
    // In the middle of the inserted word.
    const inserted = await pointAt([0], 5);
    const next = await pointAt([0], 6);
    await page.mouse.move((inserted.x + next.x) / 2, inserted.y);
    await expect(strip).toHaveAttribute('title', 'Inserted by Ann, 2026-10-01 09:00: two ');
    // Over plain text there is nothing to say.
    const plain = await pointAt([0], 1);
    await page.mouse.move(plain.x + 1, plain.y);
    await expect(strip).toHaveAttribute('title', '');
    // The name of an author is the text of the tool tip, whatever it holds.
    const other = await pointAt([1], 6);
    const after = await pointAt([1], 7);
    await page.mouse.move((other.x + after.x) / 2, other.y);
    await expect(strip).toHaveAttribute('title', `Inserted by ${EVIL}, 2026-10-01 09:00: six `);
    expect(await page.evaluate(() => [document.querySelectorAll('#host img').length, (window as any).__pwned])).toEqual([0, undefined]);
  });

  test('a caret on the border of two changes resolves one of them: the one the button names', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    // Offset 8: the insertion of Ann ends here and the deletion of Bob stands here.
    await select(page, [0], 8);
    expect(await page.evaluate(() => ((window as any).__view.selectedRevisions() as any[]).map((revision) => revision.kind))).toEqual(['insert']);
    await expect(button(page, 'acceptChange')).toHaveAttribute('title', 'Accept change: Inserted by Ann, 2026-10-01 09:00: two ');
    await button(page, 'acceptChange').click();
    await settled(page);
    expect(await revisions(page)).toEqual([
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
    expect((await notices(page)).map((notice) => notice.message)).toContain('Accepted 1 change. 2 changes left.');
  });

  test('more changes than the list holds: those beyond it are found at the caret and stepped to, and the count says "2000+"', async ({ page }) => {
    let content = '';
    for (let k = 0; k < 41; k++) {
      let inner = '';
      for (let i = 0; i < 50; i++) inner += ins(k * 50 + i + 1, 'x', 'A' + ((k * 50 + i) % 7)) + t('.');
      content += p(inner);
    }
    await mountEdit(page, await sample(content + p(t('end'))));
    const total = () => page.evaluate(() => (window as any).BymOffice.docOps.listRevisions((window as any).__view.document, { limit: Infinity }).list.length);
    expect(await page.evaluate(() => (window as any).__view.revisions().length)).toBe(2000);
    expect(await total()).toBe(2050);
    // The caret in change 2031.
    await select(page, [40], 61);
    await expect(button(page, 'rejectChange')).toBeEnabled();
    await expect(button(page, 'rejectChange')).toHaveAttribute('title', 'Reject change: Inserted by A0, 2026-10-01 09:00: x');
    await button(page, 'nextChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[40], 62, [40], 63]);
    await expect(live(page)).toHaveText('Change 2032 of 2000+. Inserted by A1, 2026-10-01 09:00: x');
    await button(page, 'rejectChange').click();
    await settled(page);
    expect(await total()).toBe(2049);
    expect((await notices(page)).map((notice) => notice.message)).toContain('Rejected 1 change. 2000+ changes left.');
    // From the last change of the bounded list the step goes on.
    expect(await page.evaluate(() => (window as any).__view.goToRevision(1999))).toBe(true);
    await button(page, 'nextChange').click();
    await settled(page);
    await expect(live(page)).toHaveText(/^Change 2001 of 2000\+\. /);
    await button(page, 'previousChange').click();
    await button(page, 'previousChange').click();
    await settled(page);
    await expect(live(page)).toHaveText(/^Change 1999 of 2000\+\. /);
  });

  test('a section break that goes with an accepted deletion of its paragraph mark is said', async ({ page }) => {
    const section = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
    const first = `<w:p><w:pPr><w:rPr><w:del${who(1, 'Bob')}/></w:rPr>${section}</w:pPr>${t('One')}</w:p>`;
    await mountEdit(page, await sample(first + p(t('Two.')) + p(t('Last.'))));
    await select(page, [1], 2);
    expect(await page.evaluate(() => (window as any).__view.acceptAllRevisions())).toBe(true);
    expect(await paragraphs(page)).toEqual(['OneTwo.', 'Last.']);
    expect(await lastNotice(page)).toBe('Accepted 1 change. No changes left. 1 section break was removed with its paragraph mark.');
  });

  test('the toolbar is used with the keyboard alone', async ({ page }) => {
    await mountEdit(page, await sample(BODY()));
    await select(page, [0], 0);
    await button(page, 'bold').focus();
    await page.keyboard.press('End');
    await expect(button(page, 'changesMenu')).toBeFocused();
    await page.keyboard.press('ArrowLeft');
    // Accept and reject have nothing to work on at the caret: the arrow passes them.
    await expect(button(page, 'nextChange')).toBeFocused();
    await page.keyboard.press('Enter');
    await settled(page);
    expect(await selection(page)).toEqual([[0], 4, [0], 8]);
    await expect(input(page)).toBeFocused();
    await button(page, 'nextChange').focus();
    await page.keyboard.press('ArrowRight');
    await expect(button(page, 'acceptChange')).toBeFocused();
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowRight');
    await expect(button(page, 'changesMenu')).toBeFocused();
    await page.keyboard.press('Enter');
    const menu = page.locator('#host .bym-popover');
    await expect(menu).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(menu).toHaveCount(0);
    await expect(button(page, 'changesMenu')).toBeFocused();
    await expect(button(page, 'changesMenu')).toHaveAttribute('aria-expanded', 'false');
  });

  test('on a narrow screen the toolbar stays one row that scrolls, and its last buttons can be used', async ({ page }) => {
    await page.setViewportSize({ width: 360, height: 640 });
    await mountEdit(page, await sample(BODY()));
    const sizes = await page.evaluate(() => {
      const bar = document.querySelector('#host .bym-doc-toolbar') as HTMLElement;
      return { height: bar.offsetHeight, scroll: bar.scrollWidth, width: bar.clientWidth, page: document.documentElement.scrollWidth, window: window.innerWidth };
    });
    expect(sizes.height).toBeLessThan(60);
    expect(sizes.scroll).toBeGreaterThan(sizes.width);
    // The page itself does not get wider than the screen.
    expect(sizes.page).toBeLessThanOrEqual(sizes.window);
    await select(page, [0], 6);
    await button(page, 'acceptChange').scrollIntoViewIfNeeded();
    await button(page, 'acceptChange').click();
    await settled(page);
    expect((await revisions(page)).length).toBe(2);
    await button(page, 'changesMenu').click();
    const item = page.locator('#host .bym-popover .bym-pop-item', { hasText: 'Reject all changes' });
    await expect(item).toBeVisible();
    // The menu is inside the screen.
    const box = (await page.locator('#host .bym-popover').boundingBox())!;
    expect(box.x).toBeGreaterThanOrEqual(0);
    expect(box.x + box.width).toBeLessThanOrEqual(360);
    await item.click();
    await settled(page);
    expect(await revisions(page)).toEqual([]);
  });
});

test.describe('tracked changes: turning tracking on and off', () => {
  test('the button turns tracking on and off: what is typed while it is on is a tracked insertion, what is typed after is not', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.'), '<w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/>'));
    const track = button(page, 'trackChanges');
    await expect(track).toHaveAttribute('aria-pressed', 'false');
    expect(await page.evaluate(() => (window as any).__view.trackingChanges)).toBe(false);
    await select(page, [0], 5);
    await track.click();
    await settled(page);
    await expect(track).toHaveAttribute('aria-pressed', 'true');
    expect(await lastNotice(page)).toBe('Changes are tracked from now on');
    await expect(input(page)).toBeFocused();
    await page.keyboard.type(' new');
    await settled(page);
    expect(await revisions(page)).toEqual([['insert', 'Me', ' new']]);
    // What cannot be tracked is not offered while tracking is on.
    await expect(button(page, 'alignCenter')).toBeDisabled();
    expect(await savedPart(page, 'word/settings.xml')).toContain('<w:zoom w:percent="100"/><w:trackRevisions/><w:defaultTabStop w:val="708"/>');

    await track.click();
    await settled(page);
    await expect(track).toHaveAttribute('aria-pressed', 'false');
    expect(await lastNotice(page)).toBe('Changes are no longer tracked');
    await expect(button(page, 'alignCenter')).toBeEnabled();
    await page.keyboard.type(' plain');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First new plain paragraph.', 'Second one.']);
    expect(await revisions(page)).toEqual([['insert', 'Me', ' new']]);
    const settings = (await savedPart(page, 'word/settings.xml')) as string;
    expect(settings).toContain('<w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/>');
    expect(settings).not.toContain('trackRevisions');
    const xml = (await savedPart(page, 'word/document.xml')) as string;
    expect(xml.split('<w:ins ').length).toBe(2);
    expect(xml).toContain('<w:ins w:id="1" w:author="Me" w:date="2026-10-08T09:08:07Z"><w:r><w:t xml:space="preserve"> new</w:t></w:r></w:ins>');

    await reopen(page);
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-pressed', 'false');
    expect(await revisions(page)).toEqual([['insert', 'Me', ' new']]);
    // Turning it on is a step that can be undone like any other.
    expect(await page.evaluate(() => [(window as any).__view.setTrackChanges(true), (window as any).__view.trackingChanges, (window as any).__view.canUndo])).toEqual([true, true, true]);
    await settled(page);
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-pressed', 'true');
    await button(page, 'undo').click();
    await settled(page);
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-pressed', 'false');
  });

  test('a document without a settings part gets one when tracking is turned on', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.')));
    await select(page, [0], 5);
    await button(page, 'trackChanges').click();
    await settled(page);
    expect(await savedPart(page, 'word/settings.xml')).toContain('<w:trackRevisions/>');
    await reopen(page);
    await expect(button(page, 'trackChanges')).toHaveAttribute('aria-pressed', 'true');
    expect(await page.evaluate(() => (window as any).__view.document.settings.trackRevisions)).toBe(true);
  });

  test('where a protection enforces tracked changes, tracking stays on and no change can be accepted or rejected', async ({ page }) => {
    await mountEdit(page, await sample(BODY(), protect('trackedChanges')));
    const track = button(page, 'trackChanges');
    await expect(track).toHaveAttribute('aria-pressed', 'true');
    await expect(track).toBeDisabled();
    await expect(track).toHaveAttribute('title', 'Track changes (cannot be turned off in this document)');
    await select(page, [0], 6);
    await expect(button(page, 'nextChange')).toBeEnabled();
    for (const command of ['acceptChange', 'rejectChange', 'changesMenu']) await expect(button(page, command)).toBeDisabled();
    await expect(button(page, 'acceptChange')).toHaveAttribute('title', 'Accept change (the protection of this document does not allow it)');
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      return [view.setTrackChanges(false), view.acceptRevision(0), view.rejectRevision(0), view.acceptAllRevisions(), view.rejectAllRevisions(), view.revisions().length, view.trackingChanges, view.canUndo];
    });
    expect(result).toEqual([false, false, false, false, false, 3, true, false]);
    const said = (await notices(page)).map((notice) => notice.message);
    expect(said).toContain('Tracking cannot be turned off in this document');
    expect(said).toContain('The protection of this document does not allow changes to be accepted or rejected');
    // Stepping through the changes is reading: it works.
    await button(page, 'nextChange').click();
    await settled(page);
    expect(await selection(page)).toEqual([[0], 8, [0], 8]);
  });

  test('tracking the application asked for cannot be turned off by its user; the changes can still be resolved', async ({ page }) => {
    await mountEdit(page, await sample(BODY()), { trackChanges: true });
    const track = button(page, 'trackChanges');
    await expect(track).toHaveAttribute('aria-pressed', 'true');
    await expect(track).toBeDisabled();
    expect(await page.evaluate(() => [(window as any).__view.setTrackChanges(false), (window as any).__view.trackingChanges])).toEqual([false, true]);
    await select(page, [0], 6);
    await button(page, 'acceptChange').click();
    await settled(page);
    // Accepting is not tracked itself: no change comes of it.
    expect(await revisions(page)).toEqual([
      ['delete', 'Bob', 'three '],
      ['insert', 'Cid', 'six '],
    ]);
    // The setting of the file is not touched by what the application asked for.
    expect(await savedPart(page, 'word/settings.xml')).toBeUndefined();
  });

  test('a document shown read-only lists its changes and lets none be resolved', async ({ page }) => {
    await mountEdit(page, await sample(BODY()), { readOnly: true });
    await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      return [view.revisions().length, view.acceptRevision(0), view.rejectRevision(0), view.acceptAllRevisions(), view.rejectAllRevisions(), view.setTrackChanges(true), view.revisions().length, view.trackingChanges];
    });
    expect(result).toEqual([3, false, false, false, false, false, 3, false]);
    // Editing turned on, they can be.
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await page.waitForFunction(() => (window as any).__view.layoutDone === true);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.acceptAllRevisions())).toBe(true);
    expect(await revisions(page)).toEqual([]);
  });
});
