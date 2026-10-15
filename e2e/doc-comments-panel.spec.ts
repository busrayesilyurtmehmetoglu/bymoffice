import type { Page } from '@playwright/test';
import { cEnd, cRef, cStart, commentXml, extendedXml, EXTENDED_REL } from '../test/doc-comments-fixtures';
import { p } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES } from './fixtures/docx';
import { expect, test } from './helpers';

// The panel that lists the comments of a document beside the pages: opened
// from the bar, the keyboard, mount() and the methods of the view; what it
// lists and in what order; a thread picked there shown in the document and a
// thread the caret goes into picked there; answering, editing, resolving and
// deleting from it, with undo; the keys; read-only and comments-only views;
// a narrow screen. Nothing is compared by screenshot.

interface MountOptions {
  readOnly?: boolean;
  locale?: string;
  author?: string;
  now?: string;
  commentsPanel?: 'auto' | 'open' | 'closed';
}

const STYLES =
  DOC_STYLES +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/></w:style>' +
  '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>';
const NOW = '2026-10-12T08:09:10.000Z';
/** A name that would be markup if it were ever put into the page as HTML. */
const EVIL = 'Eve <img src=x onerror="window.__pwned=1">';
const t = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A thread of Ann and Bob about "two", a resolved comment of Cid about "second", one of Eve about "last". */
const BODY =
  p(t('One ') + cStart(1) + cStart(2) + t('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + t(' three four.')) +
  p(t('A ') + cStart(3) + t('second') + cEnd(3) + cRef(3) + t(' paragraph of plain words.')) +
  p(t('The ') + cStart(4) + t('last') + cEnd(4) + cRef(4) + t(' paragraph.'));
const COMMENTS = [
  commentXml({ id: 1, author: 'Ann', text: 'Why two?', paraId: '0000000A', date: '2026-10-01T09:30:00Z' }),
  commentXml({ id: 2, author: 'Bob', text: 'Because.', paraId: '0000000B', date: '2026-10-02T10:00:00Z' }),
  commentXml({ id: 3, author: 'Cid', text: 'Settled long ago.', paraId: '0000000C' }),
  commentXml({ id: 4, author: EVIL, text: '<b>bold?</b>', paraId: '0000000D' }),
].join('');
const EXTENDED = extendedXml([['0000000A'], ['0000000B', '0000000A'], ['0000000C', undefined, true], ['0000000D']]);

const commented = (settings?: string) =>
  docx(BODY, { styles: STYLES, comments: COMMENTS, settings, rels: [['rIdCex', EXTENDED_REL, 'commentsExtended.xml']], parts: [['word/commentsExtended.xml', EXTENDED]] });
const plain = () => docx(p(t('Nothing to say about this.')), { styles: STYLES });
/** Forty paragraphs, each with a comment of its own: a list longer than the panel, and pages longer than the view. */
const long = () => {
  let body = '';
  const comments: string[] = [];
  for (let i = 1; i <= 40; i++) {
    body += p(t(`Paragraph ${i} has `) + cStart(i) + t(`words ${i}`) + cEnd(i) + cRef(i) + t(' and more words after them, enough for a line.'));
    comments.push(commentXml({ id: i, author: 'Ann', text: `Comment number ${i}` }));
  }
  return docx(body, { styles: STYLES, comments: comments.join('') });
};
const protect = (edit: string) => `<w:documentProtection w:edit="${edit}" w:enforcement="1"/>`;

async function mountEdit(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: 'Me', now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
  await settled(page);
}

const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
const panel = (page: Page) => page.locator('#host .bym-doc-cpanel');
const entry = (page: Page, id: number) => panel(page).locator(`.bym-doc-cpanel-item[data-thread="${id}"]`);
const entries = (page: Page) => panel(page).locator('.bym-doc-cpanel-item');
const toggle = (page: Page) => page.locator('#host .bym-doc-bar .bym-doc-comments');
const balloon = (page: Page) => page.locator('#host .bym-doc-comment-balloon');
const input = (page: Page) => page.locator('#host .bym-doc-input');
const action = (page: Page, id: number, name: string, comment?: number) =>
  entry(page, id).locator(`button[data-action="${name}"]${comment === undefined ? '' : `[data-comment="${comment}"]`}`);
const threadIds = (page: Page) => entries(page).evaluateAll((all) => all.map((element) => Number(element.getAttribute('data-thread'))));
const comments = (page: Page): Promise<[string, string, number | null, boolean][]> =>
  page.evaluate(() => ((window as any).__view.comments() as any[]).map((c) => [c.author, c.text, c.parent === undefined ? null : c.parent, c.done] as [string, string, number | null, boolean]));
const selection = (page: Page) => page.evaluate(() => (window as any).__view.getSelection());
const shown = (page: Page) => page.evaluate(() => (window as any).__view.shownComment);
// The shortcut is the editor's: it is heard where the focus is in the view. Right after mounting the focus is on the
// page, so it is put in the document first (as a click into the text would).
async function shortcut(page: Page): Promise<void> {
  await page.evaluate(() => {
    if (!document.querySelector('#host')!.contains(document.activeElement)) (window as any).__view.focus();
  });
  await page.keyboard.press('ControlOrMeta+Alt+Shift+KeyM');
}

async function select(page: Page, path: number[], offset: number): Promise<void> {
  await page.evaluate(
    ([where, at]) => {
      const view = (window as any).__view;
      view.setSelection({ story: 'body', path: where, offset: at });
      view.focus();
    },
    [path, offset] as [number[], number],
  );
  await settled(page);
}

test.describe('the panel of comments', () => {
  test('the button of the bar shows the panel: every thread in the order of its text, with its answers, as text', async ({ page, isMobile }) => {
    await mountEdit(page, await commented(), { locale: 'en' });
    await expect(toggle(page)).toBeVisible();
    await expect(toggle(page)).toHaveAttribute('aria-pressed', 'false');
    await expect(panel(page)).toHaveCount(0);
    await toggle(page).click();
    await expect(panel(page)).toBeVisible();
    await expect(toggle(page)).toHaveAttribute('aria-pressed', 'true');
    await expect(panel(page)).toHaveAttribute('role', 'complementary');
    await expect(panel(page)).toHaveAttribute('aria-label', 'Comments');
    expect(await threadIds(page)).toEqual([1, 3, 4]);
    await expect(entry(page, 1)).toContainText('Ann');
    await expect(entry(page, 1)).toContainText('10/1/2026 9:30');
    await expect(entry(page, 1)).toContainText('“two”');
    await expect(entry(page, 1)).toContainText('Why two?');
    await expect(entry(page, 1)).toContainText('1 reply');
    await expect(entry(page, 3)).toHaveAttribute('data-done', 'true');
    // What the file says reaches the page as text.
    await expect(entry(page, 4)).toContainText(EVIL);
    await expect(entry(page, 4)).toContainText('<b>bold?</b>');
    expect(await page.evaluate(() => (window as any).__pwned)).toBeUndefined();
    expect(await entry(page, 4).locator('img, b').count()).toBe(0);
    // Beside the pages on a wide screen; the pages give it the room.
    if (!isMobile) {
      const boxes = await page.evaluate(() => {
        const scroller = document.querySelector('#host .bym-doc-scroller')!.getBoundingClientRect();
        const aside = document.querySelector('#host .bym-doc-cpanel')!.getBoundingClientRect();
        return { right: scroller.right, left: aside.left, width: aside.width };
      });
      expect(boxes.width).toBeGreaterThan(200);
      expect(boxes.right).toBeLessThanOrEqual(boxes.left + 1);
    }
    // The button hides it again.
    await toggle(page).click();
    await expect(panel(page)).toBeHidden();
    await expect(toggle(page)).toHaveAttribute('aria-pressed', 'false');
  });

  test('a thread picked in the panel is selected and marked in the document; a thread the caret goes into is picked in the panel, without a balloon', async ({ page }) => {
    await mountEdit(page, await commented());
    await page.evaluate(() => (window as any).__view.setCommentsPanel(true));
    await entry(page, 3).locator('.bym-doc-cpanel-text').click();
    await expect(entry(page, 3)).toHaveAttribute('aria-current', 'true');
    expect(await shown(page)).toBe(3);
    expect(await selection(page)).toEqual({ anchor: { story: 'body', path: [1], offset: 2 }, focus: { story: 'body', path: [1], offset: 8 } });
    await expect(page.locator('#host .bym-doc-comment-mark[data-comment="3"]')).toHaveAttribute('aria-expanded', 'true');
    await expect(entry(page, 3)).toBeFocused();
    // The caret goes into the text of the thread of Ann.
    await select(page, [0], 5);
    await expect(entry(page, 1)).toHaveAttribute('aria-current', 'true');
    await expect(balloon(page)).toBeHidden();
    // And a press on the mark of a thread picks it there, with the focus.
    await page.locator('#host .bym-doc-comment-mark[data-comment="4"]').click();
    await expect(entry(page, 4)).toHaveAttribute('aria-current', 'true');
    await expect(entry(page, 4)).toBeFocused();
    await expect(balloon(page)).toBeHidden();
  });

  test('answering, editing, resolving and deleting from the panel are steps of the history; the panel follows undo and redo', async ({ page }) => {
    await mountEdit(page, await commented());
    await page.evaluate(() => (window as any).__view.setCommentsPanel(true));
    await entry(page, 1).click();
    await action(page, 1, 'reply').click();
    const field = entry(page, 1).locator('textarea');
    await expect(field).toBeFocused();
    await field.fill('Agreed.');
    await action(page, 1, 'save').click();
    expect((await comments(page)).filter((c) => c[2] === 1).map((c) => c[1])).toEqual(['Because.', 'Agreed.']);
    await expect(entry(page, 1)).toContainText('Agreed.');
    // Editing the answer of Bob.
    await action(page, 1, 'edit', 2).click();
    await expect(entry(page, 1).locator('textarea')).toHaveValue('Because.');
    await entry(page, 1).locator('textarea').fill('Because it is.');
    await entry(page, 1).locator('textarea').press('ControlOrMeta+Enter');
    await expect(entry(page, 1)).toContainText('Because it is.');
    // Resolving: Resolve becomes Reopen.
    await action(page, 1, 'resolve').click();
    await expect(entry(page, 1)).toHaveAttribute('data-done', 'true');
    await expect(action(page, 1, 'reopen')).toBeVisible();
    // Undo, from the document, and the panel follows.
    await page.evaluate(() => (window as any).__view.undo());
    await expect(entry(page, 1)).toHaveAttribute('data-done', 'false');
    await page.evaluate(() => (window as any).__view.undo());
    await expect(entry(page, 1)).toContainText('Because.');
    await page.evaluate(() => (window as any).__view.redo());
    await expect(entry(page, 1)).toContainText('Because it is.');
    // Deleting the thread takes it out of the list.
    await action(page, 1, 'delete', 1).click();
    await expect(entry(page, 1)).toHaveCount(0);
    expect(await threadIds(page)).toEqual([3, 4]);
    await page.evaluate(() => (window as any).__view.undo());
    expect(await threadIds(page)).toEqual([1, 3, 4]);
  });

  test('the keys: the shortcut opens the panel with the focus in it, arrows step, Enter goes to the document, Escape closes', async ({ page }) => {
    await mountEdit(page, await commented());
    await select(page, [2], 0);
    await shortcut(page);
    await expect(panel(page)).toBeVisible();
    await expect(entry(page, 1)).toBeFocused();
    await page.keyboard.press('ArrowDown');
    await expect(entry(page, 3)).toBeFocused();
    await expect(entry(page, 3)).toHaveAttribute('aria-current', 'true');
    // Stepping shows the thread in the document, and leaves the focus in the panel.
    expect((await selection(page)).anchor).toEqual({ story: 'body', path: [1], offset: 2 });
    await page.keyboard.press('End');
    await expect(entry(page, 4)).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(input(page)).toBeFocused();
    expect((await selection(page)).anchor).toEqual({ story: 'body', path: [2], offset: 4 });
    // Back into the panel, and Escape closes it with the focus back in the document.
    await shortcut(page);
    await expect(panel(page)).toBeHidden();
    await shortcut(page);
    await expect(entry(page, 4)).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(panel(page)).toBeHidden();
    await expect(input(page)).toBeFocused();
  });

  test('the focus is not trapped: Tab leaves the panel', async ({ page }) => {
    await mountEdit(page, await commented());
    await shortcut(page);
    await expect(entry(page, 1)).toBeFocused();
    let left = false;
    for (let i = 0; i < 12 && !left; i++) {
      await page.keyboard.press('Tab');
      left = await page.evaluate(() => !document.querySelector('#host .bym-doc-cpanel')!.contains(document.activeElement));
    }
    expect(left).toBe(true);
  });

  test('the filter hides the resolved threads and says how many', async ({ page }) => {
    await mountEdit(page, await commented(), { locale: 'tr' });
    await page.evaluate(() => (window as any).__view.setCommentsPanel(true));
    await expect(panel(page)).toHaveAttribute('aria-label', 'Yorumlar');
    await expect(entry(page, 1)).toContainText('01.10.2026 09:30');
    await panel(page).locator('input[type="checkbox"]').uncheck();
    expect(await threadIds(page)).toEqual([1, 4]);
    await expect(panel(page).locator('.bym-doc-cpanel-count')).toHaveText('2 yorum, 1 çözülmüş gizli');
    await panel(page).locator('input[type="checkbox"]').check();
    expect(await threadIds(page)).toEqual([1, 3, 4]);
  });

  test('a read-only document lists its comments and offers no change; a document protected for comments offers them', async ({ page }) => {
    await mountEdit(page, await commented(), { readOnly: true });
    await toggle(page).click();
    await entry(page, 1).click();
    await expect(entry(page, 1)).toHaveAttribute('aria-current', 'true');
    await expect(entry(page, 1).locator('button')).toHaveCount(0);
    expect(await shown(page)).toBe(1);

    await mountEdit(page, await commented(protect('comments')));
    await toggle(page).click();
    await entry(page, 4).click();
    await action(page, 4, 'reply').click();
    await entry(page, 4).locator('textarea').fill('Allowed');
    await action(page, 4, 'save').click();
    expect((await comments(page)).map((c) => c[1])).toContain('Allowed');
    expect(await page.evaluate(() => (window as any).__view.readOnly)).toBe(true);
  });

  test('a document without comments has no button unless comments can be written; mount() opens the panel when asked', async ({ page }) => {
    await mountEdit(page, await plain(), { readOnly: true });
    await expect(toggle(page)).toBeHidden();
    await mountEdit(page, await plain());
    await expect(toggle(page)).toBeVisible();
    await toggle(page).click();
    await expect(panel(page).locator('.bym-doc-cpanel-empty')).toBeVisible();
    await mountEdit(page, await commented(), { commentsPanel: 'open' });
    await expect(panel(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.commentsPanelOpen)).toBe(true);
  });

  test('on a narrow screen the panel lies over the bottom of the pages', async ({ page }) => {
    await page.setViewportSize({ width: 380, height: 640 });
    await mountEdit(page, await commented());
    const width = await page.evaluate(() => (document.querySelector('#host .bym-doc-scroller') as HTMLElement).clientWidth);
    await toggle(page).click();
    await expect(panel(page)).toHaveAttribute('data-narrow', 'true');
    const box = await page.evaluate(() => {
      const aside = document.querySelector('#host .bym-doc-cpanel')!.getBoundingClientRect();
      const body = document.querySelector('#host .bym-doc-body')!.getBoundingClientRect();
      return { bottom: Math.round(aside.bottom - body.bottom), width: Math.round(aside.width - body.width), top: aside.top > body.top };
    });
    expect(box).toEqual({ bottom: 0, width: 0, top: true });
    // The pages keep their width.
    expect(await page.evaluate(() => (document.querySelector('#host .bym-doc-scroller') as HTMLElement).clientWidth)).toBe(width);
  });

  test('stepping through a long list keeps the entry with the focus in sight in the panel', async ({ page }) => {
    await mountEdit(page, await long());
    await shortcut(page);
    await expect(entry(page, 1)).toBeFocused();
    const inSight = () =>
      page.evaluate(() => {
        const list = document.querySelector('#host .bym-doc-cpanel-scroll')!.getBoundingClientRect();
        const focused = (document.activeElement as HTMLElement).getBoundingClientRect();
        return focused.top >= list.top - 1 && focused.bottom <= list.bottom + 1;
      });
    // The first entry is in sight: selecting it does not scroll the list.
    await page.keyboard.press('Space');
    expect(await page.evaluate(() => document.querySelector('#host .bym-doc-cpanel-scroll')!.scrollTop)).toBe(0);
    for (let i = 0; i < 30; i++) await page.keyboard.press('ArrowDown');
    await expect(entry(page, 31)).toBeFocused();
    expect(await inSight()).toBe(true);
    await page.keyboard.press('Home');
    await expect(entry(page, 1)).toBeFocused();
    expect(await inSight()).toBe(true);
  });

  test('on a narrow screen the text of a thread picked in the panel is brought into sight above the drawer', async ({ page }) => {
    await page.setViewportSize({ width: 380, height: 640 });
    for (const readOnly of [false, true]) {
      await mountEdit(page, await long(), { readOnly });
      await shortcut(page);
      await expect(panel(page)).toHaveAttribute('data-narrow', 'true');
      await page.keyboard.press('End');
      await expect(entry(page, 40)).toHaveAttribute('aria-current', 'true');
      await settled(page);
      const box = await page.evaluate(() => {
        const mark = document.querySelector('#host .bym-doc-comment-mark[data-comment="40"]');
        const drawer = document.querySelector('#host .bym-doc-cpanel')!.getBoundingClientRect();
        const pages = document.querySelector('#host .bym-doc-scroller')!.getBoundingClientRect();
        if (!mark) return null;
        const rect = mark.getBoundingClientRect();
        return { top: rect.top, bottom: rect.bottom, drawer: drawer.top, pages: pages.top };
      });
      expect(box, `readOnly ${readOnly}: the thread has no mark on the pages that are shown`).not.toBeNull();
      // The mark in the margin is taller than the line it stands for; its middle is the line.
      expect((box!.top + box!.bottom) / 2, `readOnly ${readOnly}`).toBeLessThanOrEqual(box!.drawer);
      expect(box!.top, `readOnly ${readOnly}`).toBeGreaterThanOrEqual(box!.pages);
    }
  });
});
