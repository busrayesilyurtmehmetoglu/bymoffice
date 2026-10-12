import type { Page } from '@playwright/test';
import { COMMENT_ALPHA, commentColor } from '../src/ui/doc-comments';
import { ZipArchive } from '../src/zip';
import { cEnd, cRef, cStart, commentXml, extendedXml, EXTENDED_REL } from '../test/doc-comments-fixtures';
import { p } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES } from './fixtures/docx';
import { expect, test } from './helpers';

// Comments of a document in the editor: the tint over the text they are
// about and the mark in the margin, the balloon with the thread, and adding,
// editing, answering, resolving and removing from the toolbar, the balloon,
// the keyboard and the methods of the view. Nothing is compared by
// screenshot: the pixels of the page canvas are read.

interface MountOptions {
  readOnly?: boolean;
  locale?: string;
  trackChanges?: boolean;
  author?: string;
  now?: string;
  toolbar?: boolean;
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

const PLAIN = p(t('One two three four.')) + p(t('A second paragraph of plain words.')) + p(t('The last paragraph.'));
/** A thread of Ann and Bob about "two", and a resolved comment of Cid about "second". */
const BODY =
  p(t('One ') + cStart(1) + cStart(2) + t('two') + cEnd(1) + cRef(1) + cEnd(2) + cRef(2) + t(' three four.')) +
  p(t('A ') + cStart(3) + t('second') + cEnd(3) + cRef(3) + t(' paragraph of plain words.')) +
  p(t('The last paragraph.'));
const COMMENTS = [
  commentXml({ id: 1, author: 'Ann', text: 'Why two?', paraId: '0000000A', date: '2026-10-01T09:30:00Z' }),
  commentXml({ id: 2, author: 'Bob', text: 'Because.', paraId: '0000000B', date: '2026-10-02T10:00:00Z' }),
  commentXml({ id: 3, author: 'Cid', text: 'Settled long ago.', paraId: '0000000C' }),
].join('');
const EXTENDED = extendedXml([['0000000A'], ['0000000B', '0000000A'], ['0000000C', undefined, true]]);

const commented = (settings?: string) =>
  docx(BODY, { styles: STYLES, comments: COMMENTS, settings, rels: [['rIdCex', EXTENDED_REL, 'commentsExtended.xml']], parts: [['word/commentsExtended.xml', EXTENDED]] });
const plain = (settings?: string) => docx(PLAIN, { styles: STYLES, settings });
const protect = (edit: string) => `<w:documentProtection w:edit="${edit}" w:enforcement="1"/>`;

async function mountEdit(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: 'Me', now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
  await settled(page);
}

const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));
const paragraphs = (page: Page) => page.evaluate(() => ((window as any).BymOffice.documentText((window as any).__view.document) as string).split('\n'));
const notices = (page: Page): Promise<string[]> => page.evaluate(() => ((window as any).__notices as { message: string }[]).map((notice) => notice.message));
const button = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);
const input = (page: Page) => page.locator('#host .bym-doc-input');
const balloon = (page: Page) => page.locator('#host .bym-doc-comment-balloon');
const mark = (page: Page, id: number) => page.locator(`#host .bym-doc-comment-mark[data-comment="${id}"]`);
const marks = (page: Page) => page.locator('#host .bym-doc-comment-mark');
const action = (page: Page, name: string) => balloon(page).locator(`[data-action="${name}"]`);
const field = (page: Page) => balloon(page).locator('textarea');

/** The comments as the view lists them: [author, text, id of the comment it answers or null, resolved]. */
const comments = (page: Page): Promise<[string, string, number | null, boolean][]> =>
  page.evaluate(() => ((window as any).__view.comments() as any[]).map((c) => [c.author, c.text, c.parent === undefined ? null : c.parent, c.done] as [string, string, number | null, boolean]));
const ids = (page: Page): Promise<number[]> => page.evaluate(() => ((window as any).__view.comments() as any[]).map((c) => c.id));

async function select(page: Page, path: number[], offset: number, to?: [number[], number]): Promise<void> {
  await page.evaluate(
    ([where, from, end]) => {
      const view = (window as any).__view;
      view.setSelection({ anchor: { story: 'body', path: where, offset: from }, focus: end ? { story: 'body', path: end[0], offset: end[1] } : { story: 'body', path: where, offset: from } });
      view.focus();
    },
    [path, offset, to || null] as [number[], number, [number[], number] | null],
  );
  await settled(page);
}

/**
 * How many pixels of the first page have the tint an open (or resolved)
 * comment of an author lays over white paper; `shown` adds what the thread
 * whose balloon is open gets.
 */
async function tinted(page: Page, author: string, done = false, shown = false): Promise<number> {
  const [r, g, b] = commentColor(author);
  const alpha = (done ? COMMENT_ALPHA.resolved : COMMENT_ALPHA.open) + (shown ? COMMENT_ALPHA.shown : 0);
  const want = [r, g, b].map((channel) => Math.round(255 * (1 - alpha) + channel * alpha));
  await settled(page);
  return page.evaluate((colour) => {
    const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let i = 0; i < data.length; i += 4) {
      if (Math.abs(data[i] - colour[0]) <= 6 && Math.abs(data[i + 1] - colour[1]) <= 6 && Math.abs(data[i + 2] - colour[2]) <= 6) count++;
    }
    return count;
  }, want);
}

/** Every pixel of the first page that is neither white nor grey: none when nothing is tinted (the text is black). */
async function coloured(page: Page): Promise<number> {
  await settled(page);
  return page.evaluate(() => {
    const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    let count = 0;
    for (let i = 0; i < data.length; i += 4) if (Math.abs(data[i] - data[i + 1]) > 12 || Math.abs(data[i + 1] - data[i + 2]) > 12) count++;
    return count;
  });
}

async function savedPart(page: Page, name: string): Promise<string | undefined> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  const zip = ZipArchive.read(new Uint8Array(bytes));
  return zip.has(name) ? new TextDecoder().decode(await zip.read(name)) : undefined;
}

/** Is the balloon inside the editor, all of it. */
async function insideHost(page: Page): Promise<void> {
  const host = (await page.locator('#host').boundingBox())!;
  const box = (await balloon(page).boundingBox())!;
  expect(box.x).toBeGreaterThanOrEqual(host.x - 0.5);
  expect(box.y).toBeGreaterThanOrEqual(host.y - 0.5);
  expect(box.x + box.width).toBeLessThanOrEqual(host.x + host.width + 0.5);
  expect(box.y + box.height).toBeLessThanOrEqual(host.y + host.height + 0.5);
}

test.describe('comments of a document', () => {
  test('the text of a comment is tinted in the colour of its author, paler when resolved, and has a mark in the margin', async ({ page }) => {
    await mountEdit(page, await plain());
    expect(await coloured(page)).toBe(0);
    await expect(marks(page)).toHaveCount(0);

    await mountEdit(page, await commented());
    expect(await tinted(page, 'Ann')).toBeGreaterThan(40);
    expect(await tinted(page, 'Cid', true)).toBeGreaterThan(40);
    // One mark for each thread, not for each comment; named by who wrote it, and saying what else there is.
    await expect(marks(page)).toHaveCount(2);
    await expect(mark(page, 1)).toHaveAttribute('aria-label', 'Comment by Ann, 1 reply');
    await expect(mark(page, 1)).toHaveText('2');
    await expect(mark(page, 3)).toHaveAttribute('aria-label', 'Comment by Cid, Resolved');
    await expect(mark(page, 3)).toHaveAttribute('data-done', 'true');
    // In the right margin, beside the line of the comment and inside the page.
    const sheet = (await page.locator('#host .bym-doc-page[data-page="1"]').boundingBox())!;
    const first = (await mark(page, 1).boundingBox())!;
    const second = (await mark(page, 3).boundingBox())!;
    expect(first.x).toBeGreaterThan(sheet.x + sheet.width * 0.8);
    expect(first.x + first.width).toBeLessThanOrEqual(sheet.x + sheet.width + 0.5);
    expect(second.y).toBeGreaterThan(first.y);
    expect(await comments(page)).toEqual([
      ['Ann', 'Why two?', null, false],
      ['Bob', 'Because.', 1, false],
      ['Cid', 'Settled long ago.', null, true],
    ]);
  });

  test('a press on the mark opens the balloon with the thread; Escape closes it and the focus goes back to the mark', async ({ page }) => {
    await mountEdit(page, await commented());
    await mark(page, 1).click();
    await expect(balloon(page)).toBeVisible();
    await expect(balloon(page)).toHaveAttribute('role', 'dialog');
    await expect(balloon(page)).toHaveAttribute('aria-label', 'Comment by Ann');
    await expect(balloon(page).locator('.bym-doc-comment-item')).toHaveCount(2);
    await expect(balloon(page).locator('.bym-doc-comment-author')).toHaveText(['Ann', 'Bob']);
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText(['Why two?', 'Because.']);
    await expect(balloon(page).locator('.bym-doc-comment-date').first()).toContainText('2026-10-01 09:30');
    await expect(mark(page, 1)).toHaveAttribute('aria-expanded', 'true');
    // The focus is in the balloon, and the text of the thread that is shown is marked more strongly.
    expect(await page.evaluate(() => document.querySelector('#host .bym-doc-comment-balloon')!.contains(document.activeElement))).toBe(true);
    expect(await tinted(page, 'Ann', false, true)).toBeGreaterThan(40);
    await insideHost(page);
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(1);
    await page.keyboard.press('Escape');
    await expect(balloon(page)).toBeHidden();
    await expect(mark(page, 1)).toBeFocused();
    await expect(mark(page, 1)).toHaveAttribute('aria-expanded', 'false');
    // With the keyboard: Enter on the mark, the Close button of the balloon.
    await page.keyboard.press('Enter');
    await expect(balloon(page)).toBeVisible();
    await action(page, 'close').click();
    await expect(balloon(page)).toBeHidden();
  });

  test('names and text of a comment only ever reach the page as text', async ({ page }) => {
    const evil = commentXml({ id: 1, author: EVIL.replace(/&/g, '&amp;'), text: '<b>bold</b> <script>window.__pwned=2</script>' });
    await mountEdit(page, await docx(p(t('One ') + cStart(1) + t('two') + cEnd(1) + cRef(1) + t(' three.')), { styles: STYLES, comments: evil }));
    await expect(mark(page, 1)).toHaveAttribute('aria-label', `Comment by ${EVIL}`);
    await mark(page, 1).click();
    await expect(balloon(page).locator('.bym-doc-comment-author')).toHaveText(EVIL);
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText('<b>bold</b> <script>window.__pwned=2</script>');
    expect(await page.evaluate(() => [document.querySelectorAll('#host img, #host b, #host script').length, (window as any).__pwned])).toEqual([0, undefined]);
  });

  test('select, press Comment, write, save: the comment is in the document, on the page and in the balloon; undo takes it away', async ({ page }) => {
    await mountEdit(page, await plain());
    await select(page, [0], 4, [[0], 7]);
    await expect(button(page, 'newComment')).toBeEnabled();
    await expect(button(page, 'previousComment')).toBeDisabled();
    await expect(button(page, 'nextComment')).toBeDisabled();
    await button(page, 'newComment').click();
    await expect(balloon(page)).toBeVisible();
    await expect(balloon(page)).toHaveAttribute('aria-label', 'New comment');
    await expect(field(page)).toBeFocused();
    await expect(field(page)).toHaveAttribute('aria-label', 'Comment text');
    await insideHost(page);
    // Nothing is added before the field is saved, and an empty field adds nothing.
    await action(page, 'save').click();
    expect(await comments(page)).toEqual([]);
    expect((await notices(page)).pop()).toContain('needs text');
    await field(page).fill('Why two?\nReally.');
    await action(page, 'save').click();
    expect(await comments(page)).toEqual([['Me', 'Why two?\nReally.', null, false]]);
    expect((await notices(page)).pop()).toBe('Comment added');
    // The keys are the document's again, and the balloon shows the new thread.
    await expect(input(page)).toBeFocused();
    await expect(balloon(page)).toBeVisible();
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText('Why two?\nReally.');
    await expect(balloon(page).locator('.bym-doc-comment-date')).toContainText('2026-10-12 08:09');
    await expect(marks(page)).toHaveCount(1);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'A second paragraph of plain words.', 'The last paragraph.']);
    expect(await page.evaluate(() => [(window as any).__changes, (window as any).__view.isModified, (window as any).__view.canUndo])).toEqual([1, true, true]);
    // The selection is still the text of the comment: without it, the tint of the thread that is shown.
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('two');
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [2], offset: 0 }));
    await page.evaluate(() => (window as any).__view.showComment((window as any).__view.comments()[0].id));
    expect(await tinted(page, 'Me', false, true)).toBeGreaterThan(40);
    await expect(button(page, 'nextComment')).toBeEnabled();

    await page.evaluate(() => (window as any).__view.undo());
    expect(await comments(page)).toEqual([]);
    await expect(marks(page)).toHaveCount(0);
    await expect(balloon(page)).toBeHidden();
    // Undo brought the selection back; without it nothing on the page has a colour.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [2], offset: 0 }));
    expect(await coloured(page)).toBe(0);
    await page.evaluate(() => (window as any).__view.redo());
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [2], offset: 0 }));
    expect(await comments(page)).toEqual([['Me', 'Why two?\nReally.', null, false]]);
    await expect(marks(page)).toHaveCount(1);
    expect(await tinted(page, 'Me')).toBeGreaterThan(40);
  });

  test('Ctrl+Alt+M opens the field for a comment about the word at the caret; Cancel and Escape give it up', async ({ page }) => {
    await mountEdit(page, await plain());
    await select(page, [0], 5);
    const command = await page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));
    await page.keyboard.press(`${command}+Alt+KeyM`);
    await expect(field(page)).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(balloon(page)).toBeHidden();
    await expect(input(page)).toBeFocused();
    expect(await comments(page)).toEqual([]);
    await page.keyboard.press(`${command}+Alt+KeyM`);
    await field(page).fill('given up');
    await action(page, 'cancel').click();
    await expect(balloon(page)).toBeHidden();
    expect(await comments(page)).toEqual([]);
    // Saved with the keyboard: the comment is about the word the caret was in.
    await page.keyboard.press(`${command}+Alt+KeyM`);
    await field(page).fill('About the word');
    await page.keyboard.press(`${command}+Enter`);
    expect(await comments(page)).toEqual([['Me', 'About the word', null, false]]);
    const range = await page.evaluate(() => {
      const comment = (window as any).__view.comments()[0];
      return [comment.start.offset, comment.end.offset];
    });
    expect(range).toEqual([4, 7]);
    // Typing in the field never reached the document, and the space bar did not scroll the pages.
    expect(await paragraphs(page)).toEqual(['One two three four.', 'A second paragraph of plain words.', 'The last paragraph.']);
  });

  test('edit, reply, resolve, reopen and delete from the balloon, each one step that can be undone', async ({ page }) => {
    await mountEdit(page, await commented());
    await mark(page, 1).click();
    // Edit the answer of Bob.
    await balloon(page).locator('[data-action="edit"][data-comment="2"]').click();
    await expect(field(page)).toBeFocused();
    await expect(field(page)).toHaveValue('Because.');
    await field(page).fill('Because of one.');
    await action(page, 'save').click();
    expect((await comments(page))[1]).toEqual(['Bob', 'Because of one.', 1, false]);
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText(['Why two?', 'Because of one.']);
    await expect(input(page)).toBeFocused();
    // Reply.
    await action(page, 'reply').click();
    await expect(field(page)).toHaveAttribute('aria-label', 'Reply text');
    await field(page).fill('Fine.');
    await action(page, 'save').click();
    expect(await comments(page)).toEqual([
      ['Ann', 'Why two?', null, false],
      ['Bob', 'Because of one.', 1, false],
      ['Me', 'Fine.', 1, false],
      ['Cid', 'Settled long ago.', null, true],
    ]);
    await expect(mark(page, 1)).toHaveText('3');
    await expect(balloon(page).locator('.bym-doc-comment-item')).toHaveCount(3);
    // Resolve: the thread says so, its text is paler, and the button in the same place opens it again.
    await action(page, 'resolve').click();
    await expect(balloon(page).locator('.bym-doc-comment-state')).toHaveText('Resolved');
    expect((await comments(page)).slice(0, 3).map((c) => c[3])).toEqual([true, true, true]);
    await expect(mark(page, 1)).toHaveAttribute('data-done', 'true');
    await expect(action(page, 'reopen')).toBeFocused();
    await action(page, 'reopen').click();
    expect((await comments(page)).slice(0, 3).map((c) => c[3])).toEqual([false, false, false]);
    await expect(action(page, 'resolve')).toBeFocused();
    // Delete the answer of Bob, then the whole thread.
    await balloon(page).locator('[data-action="delete"][data-comment="2"]').click();
    expect((await comments(page)).map((c) => c[0])).toEqual(['Ann', 'Me', 'Cid']);
    await expect(balloon(page)).toBeVisible();
    await balloon(page).locator('[data-action="delete"][data-comment="1"]').click();
    expect((await comments(page)).map((c) => c[0])).toEqual(['Cid']);
    await expect(balloon(page)).toBeHidden();
    await expect(marks(page)).toHaveCount(1);
    await expect(input(page)).toBeFocused();
    expect(await tinted(page, 'Ann')).toBe(0);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'A second paragraph of plain words.', 'The last paragraph.']);
    // Six steps, each undone alone: the document is as it was opened.
    for (let i = 0; i < 6; i++) expect(await page.evaluate(() => (window as any).__view.undo())).toBe(true);
    expect(await page.evaluate(() => (window as any).__view.canUndo)).toBe(false);
    expect(await comments(page)).toEqual([
      ['Ann', 'Why two?', null, false],
      ['Bob', 'Because.', 1, false],
      ['Cid', 'Settled long ago.', null, true],
    ]);
    await expect(marks(page)).toHaveCount(2);
  });

  test('saved and opened again, the comments, the answers and what is resolved are there', async ({ page }) => {
    await mountEdit(page, await commented());
    const made = await page.evaluate(() => {
      const view = (window as any).__view;
      const added = view.addComment('A new one', { anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } });
      const reply = view.replyToComment(added, 'And its answer');
      view.resolveComment(1, true);
      view.editComment(3, 'Settled again.');
      return [added, reply];
    });
    expect(made.every((id) => typeof id === 'number')).toBe(true);
    const expected = await comments(page);
    expect(expected).toEqual([
      ['Ann', 'Why two?', null, true],
      ['Bob', 'Because.', 1, true],
      ['Cid', 'Settled again.', null, true],
      ['Me', 'A new one', null, false],
      ['Me', 'And its answer', made[0], false],
    ]);
    const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    const zip = ZipArchive.read(new Uint8Array(bytes));
    const extended = new TextDecoder().decode(await zip.read('word/commentsExtended.xml'));
    expect(extended).toContain('w15:paraId="0000000A" w15:done="1"');
    expect((extended.match(/<w15:commentEx /g) || []).length).toBe(5);
    await mountEdit(page, new Uint8Array(bytes));
    expect(await comments(page)).toEqual(expected);
    await expect(marks(page)).toHaveCount(3);
    expect(await tinted(page, 'Me')).toBeGreaterThan(40);
  });

  test('add a comment, undo, save: the bytes of the file that was opened', async ({ page }) => {
    const bytes = await commented();
    await mountEdit(page, bytes);
    await page.evaluate(() => (window as any).__view.addComment('Soon gone', { anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } }));
    expect((await comments(page)).length).toBe(4);
    await page.evaluate(() => (window as any).__view.undo());
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    expect(saved).toEqual(Array.from(bytes));
  });

  test('the comment buttons step through the threads, select their text and say what they are', async ({ page }) => {
    await mountEdit(page, await commented());
    await select(page, [0], 0);
    await button(page, 'nextComment').click();
    await expect(balloon(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(1);
    expect((await notices(page)).pop()).toBe('Comment 1 of 2, by Ann: Why two?');
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('two');
    await expect(input(page)).toBeFocused();
    await button(page, 'nextComment').click();
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(3);
    expect((await notices(page)).pop()).toBe('Comment 2 of 2, by Cid: Settled long ago.');
    await button(page, 'nextComment').click();
    expect((await notices(page)).pop()).toBe('Back at the first comment. Comment 1 of 2, by Ann: Why two?');
    await button(page, 'previousComment').click();
    expect((await notices(page)).pop()).toBe('At the last comment. Comment 2 of 2, by Cid: Settled long ago.');
    await expect(page.locator('#host .bym-doc-live')).toHaveText('At the last comment. Comment 2 of 2, by Cid: Settled long ago.');
    // In Turkish.
    await page.evaluate(() => (window as any).__mounted.editor.setLocale('tr'));
    await settled(page);
    await expect(button(page, 'previousComment')).toHaveAttribute('aria-label', 'Önceki yorum');
    await expect(button(page, 'nextComment')).toHaveAttribute('aria-label', 'Sonraki yorum');
    expect(await button(page, 'newComment').getAttribute('aria-label')).toMatch(/^Yorum \(/);
    await expect(balloon(page).locator('.bym-doc-comment-state')).toHaveText('Çözüldü');
    await expect(mark(page, 3)).toHaveAttribute('aria-label', 'Yorum: Cid, Çözüldü');
  });

  test('the caret in the text of a comment brings its balloon; outside of it the balloon goes', async ({ page, isMobile }) => {
    test.skip(isMobile, 'on a narrow screen the balloon lies along the bottom edge and comes only when asked for');
    await mountEdit(page, await commented());
    await select(page, [2], 3);
    await expect(balloon(page)).toBeHidden();
    await select(page, [0], 5);
    await expect(balloon(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(1);
    // It did not take the focus: typing goes on in the document, and the comment grows with its text.
    await expect(input(page)).toBeFocused();
    await page.keyboard.type('XY');
    expect((await paragraphs(page))[0]).toBe('One tXYwo three four.');
    await expect(balloon(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.comments()[0].end.offset)).toBe(9);
    // Escape sends it away, and it stays away while the caret stays in that comment.
    await page.keyboard.press('Escape');
    await expect(balloon(page)).toBeHidden();
    await page.keyboard.press('ArrowLeft');
    await expect(balloon(page)).toBeHidden();
    await select(page, [2], 3);
    await select(page, [1], 4);
    await expect(balloon(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(3);
    await select(page, [2], 3);
    await expect(balloon(page)).toBeHidden();
  });

  test('deleting the text of a comment leaves the comment, about a place, with its mark', async ({ page }) => {
    await mountEdit(page, await commented());
    await select(page, [1], 2, [[1], 8]);
    await page.keyboard.press('Delete');
    expect((await paragraphs(page))[1]).toBe('A  paragraph of plain words.');
    expect((await comments(page)).map((c) => c[0])).toEqual(['Ann', 'Bob', 'Cid']);
    await expect(mark(page, 3)).toBeVisible();
    expect(await page.evaluate(() => {
      const comment = (window as any).__view.comments()[2];
      return [comment.start.offset, comment.end.offset];
    })).toEqual([2, 2]);
  });

  test('a read-only view shows the comments and their balloon, and offers nothing that changes them', async ({ page }) => {
    await mountEdit(page, await commented(), { readOnly: true });
    await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
    await expect(marks(page)).toHaveCount(2);
    expect(await tinted(page, 'Ann')).toBeGreaterThan(40);
    await mark(page, 1).click();
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText(['Why two?', 'Because.']);
    await expect(balloon(page).locator('[data-action="edit"], [data-action="delete"], [data-action="reply"], [data-action="resolve"]')).toHaveCount(0);
    await expect(action(page, 'close')).toBeVisible();
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      return [view.canComment, view.addComment('x', { story: 'body', path: [2], offset: 5 }), view.editComment(1, 'x'), view.replyToComment(1, 'x'), view.resolveComment(1, true), view.removeComment(1), view.startComment()];
    });
    expect(result).toEqual([false, undefined, false, undefined, false, false, false]);
    expect((await comments(page)).length).toBe(3);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    // A press on the commented text itself brings the balloon too; a press beside it takes it away.
    await action(page, 'close').click();
    await expect(balloon(page)).toBeHidden();
    const word = await page.evaluate(() => {
      const view = (window as any).__view;
      const sheet = document.querySelector('#host .bym-doc-page[data-page="1"]') as HTMLElement;
      const rect = sheet.getBoundingClientRect();
      const mark = document.querySelector('#host .bym-doc-comment-mark[data-comment="3"]') as HTMLElement;
      const line = mark.getBoundingClientRect();
      void view;
      return { left: rect.left, width: rect.width, y: line.top + line.height / 2 };
    });
    // "second" begins after "A " at the left margin (40 of 600 px): a point well inside the word.
    await page.mouse.click(word.left + (word.width * 62) / 600, word.y);
    await expect(balloon(page)).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.shownComment)).toBe(3);
    // Beside the text, and beside the balloon (which lies under the line, along the right edge of the page).
    await page.mouse.click(word.left + (word.width * 50) / 600, word.y + 90);
    await expect(balloon(page)).toBeHidden();
  });

  test('a document protected for comments takes comments and nothing else, and is saved with them', async ({ page }) => {
    await mountEdit(page, await commented(protect('comments')));
    expect(await page.evaluate(() => [(window as any).__mounted.protection, (window as any).__view.readOnly, (window as any).__view.canComment])).toEqual([{ edit: 'comments', exceptions: false }, true, true]);
    const said = await notices(page);
    expect(said.length).toBe(1);
    expect(said[0]).toContain('only comments');
    // The toolbar is there for the comments; what changes the text is not offered.
    await expect(page.locator('#host .bym-doc-toolbar')).toBeVisible();
    for (const command of ['bold', 'italic', 'alignCenter', 'insertTable', 'pageBreak', 'trackChanges']) await expect(button(page, command)).toBeDisabled();
    await expect(button(page, 'newComment')).toBeEnabled();
    await expect(button(page, 'nextComment')).toBeEnabled();
    // Typing changes nothing.
    await page.evaluate(() => (window as any).__view.focus());
    await page.keyboard.type('X');
    await page.keyboard.press('Delete');
    expect(await page.evaluate(() => (window as any).__view.insertText('X'))).toBe(false);
    expect(await paragraphs(page)).toEqual(['One two three four.', 'A second paragraph of plain words.', 'The last paragraph.']);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    // A comment about text that is selected (a selection of the read-only view), from the toolbar.
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } }));
    await button(page, 'newComment').click();
    await field(page).fill('Allowed');
    await action(page, 'save').click();
    expect((await comments(page)).map((c) => c[1])).toEqual(['Why two?', 'Because.', 'Settled long ago.', 'Allowed']);
    await expect(marks(page)).toHaveCount(3);
    // The balloon offers what the protection allows.
    await mark(page, 1).click();
    await action(page, 'resolve').click();
    expect((await comments(page))[0][3]).toBe(true);
    expect(await page.evaluate(() => [(window as any).__view.isModified, (window as any).__view.canUndo])).toEqual([true, true]);
    await expect(button(page, 'undo')).toBeEnabled();
    const saved = (await savedPart(page, 'word/comments.xml')) as string;
    expect(saved).toContain('Allowed');
    expect(await savedPart(page, 'word/settings.xml')).toContain(protect('comments'));
    expect(await savedPart(page, 'word/commentsExtended.xml')).toContain('w15:paraId="0000000A" w15:done="1"');
    // Undo works for the comments, and only for them.
    await button(page, 'undo').click();
    expect((await comments(page))[0][3]).toBe(false);
    // Turned off and on again by the application.
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(true));
    await settled(page);
    expect(await page.evaluate(() => [(window as any).__view.canComment, (window as any).__view.addComment('no', { story: 'body', path: [0], offset: 1 })])).toEqual([false, undefined]);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await settled(page);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeVisible();
    expect((await notices(page)).filter((message) => message.indexOf('protected') >= 0).length).toBe(1);
  });

  test('with tracking on, comments are no tracked changes', async ({ page }) => {
    await mountEdit(page, await commented('<w:trackRevisions/>'));
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      const added = view.addComment('Tracked?', { anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } });
      view.replyToComment(1, 'No.');
      view.removeComment(3);
      return [typeof added, view.trackingChanges, view.revisions().length];
    });
    expect(result).toEqual(['number', true, 0]);
    expect((await comments(page)).map((c) => c[1])).toEqual(['Why two?', 'Because.', 'No.', 'Tracked?']);
    const body = (await savedPart(page, 'word/document.xml')) as string;
    expect(body).not.toMatch(/<w:ins |<w:del /);
  });

  test('in a small editor the balloon lies along the bottom edge, inside the editor, and works by touch', async ({ page }) => {
    await page.setViewportSize({ width: 340, height: 520 });
    await mountEdit(page, await commented());
    const touch = test.info().project.use.hasTouch === true;
    const press = (locator: ReturnType<typeof mark>) => (touch ? locator.tap() : locator.click());
    await mark(page, 1).scrollIntoViewIfNeeded();
    await press(mark(page, 1));
    await expect(balloon(page)).toBeVisible();
    await insideHost(page);
    const host = (await page.locator('#host').boundingBox())!;
    const box = (await balloon(page).boundingBox())!;
    expect(box.width).toBeGreaterThan(host.width - 40);
    expect(box.y + box.height).toBeGreaterThan(host.y + host.height - 24);
    // The page itself does not get wider than the screen.
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
    await press(action(page, 'reply'));
    await expect(field(page)).toBeFocused();
    await insideHost(page);
    await field(page).fill('From a small screen');
    await press(action(page, 'save'));
    expect((await comments(page)).map((c) => c[1])).toEqual(['Why two?', 'Because.', 'From a small screen', 'Settled long ago.']);
    await expect(balloon(page).locator('.bym-doc-comment-item')).toHaveCount(3);
    await insideHost(page);
    await press(action(page, 'close'));
    await expect(balloon(page)).toBeHidden();
  });

  // ------------------------------------------------- who may change comments

  /** Every way the view has of changing a comment, tried: [canComment, add, edit, reply, resolve, remove, the field of a new one]. */
  const tryAll = (page: Page) =>
    page.evaluate(() => {
      const view = (window as any).__view;
      const at = { anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 3 } };
      return [view.canComment, view.addComment('x', at), view.editComment(1, 'x'), view.replyToComment(1, 'x'), view.resolveComment(1, true), view.removeComment(1), view.startComment()];
    });
  const NOTHING = [false, undefined, false, undefined, false, false, false];

  for (const edit of ['forms', 'readOnly']) {
    test(`a document protected for "${edit}" takes no comments, whatever is asked of it, and is saved as it was opened`, async ({ page }) => {
      const bytes = await commented(protect(edit));
      await mountEdit(page, bytes);
      expect(await tryAll(page)).toEqual(NOTHING);
      // Asking for editing again opens nothing.
      await page.evaluate(() => {
        (window as any).__mounted.editor.setReadOnly(false);
        (window as any).__view.setReadOnly(false);
      });
      await settled(page);
      expect(await page.evaluate(() => [(window as any).__view.readOnly, (window as any).__view.canComment])).toEqual([true, false]);
      expect(await tryAll(page)).toEqual(NOTHING);
      await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
      // The shortcut opens no field, and the balloon offers nothing that changes a comment.
      await page.evaluate(() => (window as any).__view.focus());
      const command = await page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));
      await page.keyboard.press(`${command}+Alt+KeyM`);
      await expect(balloon(page)).toBeHidden();
      await mark(page, 1).click();
      await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText(['Why two?', 'Because.']);
      await expect(balloon(page).locator('[data-action="edit"], [data-action="delete"], [data-action="reply"], [data-action="resolve"], textarea')).toHaveCount(0);
      expect((await comments(page)).length).toBe(3);
      expect(await page.evaluate(() => [(window as any).__view.isModified, (window as any).__view.canUndo, (window as any).__changes])).toEqual([false, false, 0]);
      const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
      expect(saved).toEqual(Array.from(bytes));
    });
  }

  test('a view that was mounted read-only takes no comments until editing is turned on', async ({ page }) => {
    await mountEdit(page, await commented(), { readOnly: true });
    expect(await tryAll(page)).toEqual(NOTHING);
    expect((await comments(page)).length).toBe(3);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await settled(page);
    expect(await page.evaluate(() => (window as any).__view.canComment)).toBe(true);
    expect(await page.evaluate(() => (window as any).__view.resolveComment(1, true))).toBe(true);
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(true));
    await settled(page);
    expect(await tryAll(page)).toEqual(NOTHING);
  });

  test('an RTF file that was not converted shows its comments and takes none: nothing turns that on', async ({ page }) => {
    const rtf = '{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\pard One two {\\*\\atnid AB}{\\*\\atnauthor Ann Bee}\\chatn{\\*\\annotation \\pard\\plain {\\chatn }A remark.\\par}three four.\\par Second paragraph.\\par}';
    const bytes = new TextEncoder().encode(rtf);
    await mountEdit(page, bytes);
    expect(await page.evaluate(() => [(window as any).__mounted.file.format, (window as any).__view.readOnly])).toEqual(['rtf', true]);
    expect(await comments(page)).toEqual([['Ann Bee', 'A remark.', null, false]]);
    await expect(marks(page)).toHaveCount(1);
    const id = (await ids(page))[0];
    const attempt = () =>
      page.evaluate((comment) => {
        const view = (window as any).__view;
        return [view.canComment, view.addComment('x', { story: 'body', path: [0], offset: 1 }), view.editComment(comment, 'x'), view.replyToComment(comment, 'x'), view.resolveComment(comment, true), view.removeComment(comment), view.startComment()];
      }, id);
    expect(await attempt()).toEqual(NOTHING);
    await page.evaluate(() => {
      (window as any).__mounted.editor.setReadOnly(false);
      (window as any).__view.setReadOnly(false);
    });
    await settled(page);
    expect(await attempt()).toEqual(NOTHING);
    await mark(page, id).click();
    await expect(balloon(page).locator('.bym-doc-comment-body')).toHaveText('A remark.');
    await expect(balloon(page).locator('[data-action="edit"], [data-action="delete"], [data-action="reply"], [data-action="resolve"]')).toHaveCount(0);
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    expect(saved).toEqual(Array.from(bytes));
  });

  test('in a document protected for comments nothing but a comment changes anything: keys, the clipboard, an input method, formatting, tracked changes, a long press', async ({ page }) => {
    const who = ' w:id="40" w:author="Bob" w:date="2026-10-01T09:00:00Z"';
    const body = BODY + p(t('Tracked: ') + `<w:ins${who}>${t('inserted')}</w:ins>` + t(' text.'));
    const bytes = await docx(body, { styles: STYLES, comments: COMMENTS, settings: protect('comments'), rels: [['rIdCex', EXTENDED_REL, 'commentsExtended.xml']], parts: [['word/commentsExtended.xml', EXTENDED]] });
    await mountEdit(page, bytes);
    const text = ['One two three four.', 'A second paragraph of plain words.', 'The last paragraph.', 'Tracked: inserted text.'];
    const unchanged = async (what: string) => {
      await settled(page);
      expect(await paragraphs(page), what).toEqual(text);
      expect(await page.evaluate(() => [(window as any).__view.isModified, (window as any).__view.canUndo, (window as any).__changes]), what).toEqual([false, false, 0]);
    };
    const command = await page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));
    // A selection of the text, the focus in the view: what an editor would type over.
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 7 } });
      view.focus();
    });
    await expect(input(page)).toBeFocused();
    for (const key of ['x', 'Delete', 'Backspace', 'Enter', 'Tab', `${command}+b`, `${command}+i`, `${command}+z`, `${command}+y`, `${command}+Enter`, `${command}+x`]) await page.keyboard.press(key);
    await unchanged('keys');
    // The clipboard, a drop and an input method, as their events arrive at the field of the view.
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const data = new DataTransfer();
      data.setData('text/plain', 'pasted text');
      data.setData('text/html', '<b>pasted</b>');
      field.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
      field.dispatchEvent(new ClipboardEvent('cut', { clipboardData: new DataTransfer(), bubbles: true, cancelable: true }));
      const sheet = document.querySelector('#host .bym-doc-page[data-page="1"]') as HTMLElement;
      for (const type of ['dragenter', 'dragover', 'drop']) sheet.dispatchEvent(new DragEvent(type, { dataTransfer: data, bubbles: true, cancelable: true, clientX: 100, clientY: 100 }));
      field.dispatchEvent(new CompositionEvent('compositionstart', { data: '', bubbles: true }));
      field.value = 'にほん';
      field.dispatchEvent(new InputEvent('input', { data: 'にほん', inputType: 'insertCompositionText', isComposing: true, bubbles: true }));
      field.dispatchEvent(new CompositionEvent('compositionend', { data: 'にほん', bubbles: true }));
      field.dispatchEvent(new InputEvent('beforeinput', { data: 'typed', inputType: 'insertText', bubbles: true, cancelable: true }));
      field.value = 'typed';
      field.dispatchEvent(new InputEvent('input', { data: 'typed', inputType: 'insertText', bubbles: true }));
    });
    await unchanged('the clipboard, a drop and an input method');
    // Everything the view offers an editor.
    const results = await page.evaluate(() => {
      const view = (window as any).__view;
      view.setReadOnly(false);
      (window as any).__mounted.editor.setReadOnly(false);
      const out: unknown[] = [view.readOnly];
      out.push(view.insertText('X'), view.format({ bold: true }), view.toggleFormat('bold'), view.formatParagraph({ align: 'center' }), view.toggleList('bullet'), view.indent(1));
      out.push(view.insertTable(2, 2), view.insertPageBreak(), view.insertLink('https://example.org/'), view.insertPicture(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])));
      out.push(view.setTrackChanges(true), view.acceptRevision(0), view.rejectRevision(0), view.acceptAllRevisions(), view.rejectAllRevisions(), view.undo(), view.redo());
      out.push(view.revisions().length);
      return out;
    });
    expect(results).toEqual([true, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, 1]);
    await unchanged('the methods of the view');
    // A finger that stays on a word selects nothing that could be cut or typed over.
    if (test.info().project.use.hasTouch === true) {
      const sheet = (await page.locator('#host .bym-doc-page[data-page="1"]').boundingBox())!;
      const client = await page.context().newCDPSession(page).catch(() => undefined);
      if (client) {
        const point = { x: Math.round(sheet.x + sheet.width * 0.2), y: Math.round(sheet.y + sheet.height * 0.13) };
        await client.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [point] });
        await page.waitForTimeout(800);
        await client.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
        await expect(page.locator('#host .bym-doc-bubble, #host .bym-doc-handle')).toHaveCount(0);
        await unchanged('a long press');
      }
    }
    // Not one byte of the file is different.
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    expect(saved).toEqual(Array.from(bytes));

    // A comment is the one thing that is taken, and Ctrl+Z takes it back and nothing else.
    await page.evaluate(() => (window as any).__view.addComment('Allowed', { anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } }));
    expect((await comments(page)).length).toBe(4);
    const withComment = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    const before = ZipArchive.read(new Uint8Array(bytes));
    const after = ZipArchive.read(new Uint8Array(withComment));
    const differing: string[] = [];
    for (const name of after.names()) {
      const now = Array.from(await after.read(name));
      const was = before.has(name) ? Array.from(await before.read(name)) : [];
      if (now.length !== was.length || now.some((byte, i) => byte !== was[i])) differing.push(name);
    }
    expect(differing.sort()).toEqual(['word/comments.xml', 'word/commentsExtended.xml', 'word/document.xml']);
    const changedBody = new TextDecoder().decode(await after.read('word/document.xml'));
    const originalBody = new TextDecoder().decode(await before.read('word/document.xml'));
    // The paragraphs the comment is not about are the bytes of the file; the one it is about gained its three marks and nothing else.
    const paragraphsOf = (xml: string) => xml.match(/<w:p>.*?<\/w:p>/g) || [];
    const [was, now] = [paragraphsOf(originalBody), paragraphsOf(changedBody)];
    expect(now.length).toBe(4);
    for (const index of [0, 1, 3]) expect(now[index]).toBe(was[index]);
    expect(now[2].replace(/<[^>]+>/g, '')).toBe('The last paragraph.');
    expect((now[2].match(/<w:comment(RangeStart|RangeEnd|Reference) w:id="\d+"\/>/g) || []).length).toBe(3);
    expect(changedBody.replace(/<w:p>.*<\/w:p>/, '')).toBe(originalBody.replace(/<w:p>.*<\/w:p>/, ''));
    await page.evaluate(() => (window as any).__view.focus());
    await page.keyboard.press(`${command}+z`);
    expect((await comments(page)).length).toBe(3);
    await page.keyboard.press(`${command}+z`);
    await page.keyboard.press(`${command}+z`);
    expect(await paragraphs(page)).toEqual(text);
    expect(await page.evaluate(() => (window as any).__view.canUndo)).toBe(false);
    const back = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    expect(back).toEqual(Array.from(bytes));
    await page.keyboard.press(`${command}+y`);
    expect((await comments(page)).length).toBe(4);
  });

  test('a document protected for comments is saved to where it came from with its comments', async ({ page, request }) => {
    const info = test.info();
    const url = `/e2e/saved/${`${info.project.name}-${info.testId}`.replace(/[^A-Za-z0-9_-]/g, '_')}.docx`;
    const seeded = await request.put(url, { data: Buffer.from(await commented(protect('comments'))), headers: { 'content-type': 'application/octet-stream' } });
    expect(seeded.status()).toBe(204);
    await page.goto('/e2e/pages/mount.html');
    await page.evaluate(async (file) => {
      (window as any).__mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file, saveTo: true, locale: 'en', author: 'Me' });
    }, url);
    const save = page.locator('#editor .bym-toolbar [data-command="save"]');
    await expect(save).toBeVisible();
    await expect(save).toBeDisabled();
    expect(await page.evaluate(() => (window as any).__mounted.view.addComment('Kept on the server', { anchor: { story: 'body', path: [2], offset: 4 }, focus: { story: 'body', path: [2], offset: 8 } }))).not.toBeUndefined();
    await expect(save).toBeEnabled();
    expect(await page.evaluate(() => (window as any).__mounted.hasUnsavedChanges)).toBe(true);
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
    await expect(save).toBeDisabled();
    const answer = await request.get(url);
    const zip = ZipArchive.read(new Uint8Array(await answer.body()));
    expect(new TextDecoder().decode(await zip.read('word/comments.xml'))).toContain('Kept on the server');
    // A protection that allows nothing is not written to at all.
    const locked = url.replace('.docx', '-forms.docx');
    await request.put(locked, { data: Buffer.from(await commented(protect('forms'))), headers: { 'content-type': 'application/octet-stream' } });
    await page.evaluate(async (file) => {
      (window as any).__mounted.destroy();
      (window as any).__mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file, saveTo: true, locale: 'en' });
    }, locked);
    expect(await page.evaluate(() => [(window as any).__mounted.view.canComment, (window as any).__mounted.view.addComment('no', { story: 'body', path: [0], offset: 1 })])).toEqual([false, undefined]);
    await expect(page.locator('#editor .bym-toolbar [data-command="save"]')).toBeHidden();
  });

  test('typing paints the page that is typed on and no other; a change of the comments paints them all', async ({ page }) => {
    // Three pages; the comment is on the first, the typing on the third (the layout lays out and paints again
    // the page before the one that is typed on and every page after it, with or without comments).
    const body = p(t('First page: ') + cStart(1) + t('commented') + cEnd(1) + cRef(1) + t(' text.')) + p(t('Second page.'), '<w:pageBreakBefore/>') + p(t('Third page.'), '<w:pageBreakBefore/>') + p(t('Third page, where the typing is.'));
    await mountEdit(page, await docx(body, { styles: STYLES, comments: commentXml({ id: 1, author: 'Ann', text: 'Here' }) }));
    await settled(page);
    /** A dot the view did not paint, on the first page: it is gone as soon as that page is painted again. */
    const dot = () =>
      page.evaluate(() => {
        const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
        const context = canvas.getContext('2d')!;
        context.save();
        context.setTransform(1, 0, 0, 1, 0, 0);
        context.fillStyle = 'rgb(1, 200, 3)';
        context.fillRect(2, 2, 4, 4);
        context.restore();
      });
    const dotThere = () =>
      page.evaluate(() => {
        const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
        const pixel = canvas.getContext('2d')!.getImageData(3, 3, 1, 1).data;
        return pixel[0] === 1 && pixel[1] === 200 && pixel[2] === 3;
      });
    await select(page, [3], 5);
    await page.keyboard.type('a');
    await settled(page);
    // The first page is still shown (the pages are small), though the caret is on the third.
    await expect(page.locator('#host .bym-doc-page[data-page="1"] canvas')).toHaveCount(1);
    await dot();
    await page.keyboard.type('bcd');
    await settled(page);
    expect((await paragraphs(page))[3]).toBe('Thirdabcd page, where the typing is.');
    await expect(page.locator('#host .bym-doc-page[data-page="1"] canvas')).toHaveCount(1);
    expect(await dotThere(), 'the page with the comment was painted again although nothing on it changed').toBe(true);
    expect(await page.evaluate(() => (window as any).__view.resolveComment(1, true))).toBe(true);
    await settled(page);
    expect(await dotThere(), 'the page with the comment was not painted again when the comment was resolved').toBe(false);
    expect(await page.evaluate(() => (window as any).__view.commentsRevision)).toBeGreaterThan(0);
  });

  test('nothing of the comments stays behind when the editor is destroyed', async ({ page }) => {
    // What stays after an editor of a document without comments was destroyed is what may stay here.
    await mountEdit(page, await plain());
    const empty = await page.evaluate(() => {
      (window as any).__mounted.destroy();
      (window as any).__mounted = null;
      return { ...(window as any).__counts };
    });
    await page.evaluate(([bytes]) => (window as any).__mountEdit(bytes, { author: 'Me' }), [Array.from(await commented())]);
    await page.waitForFunction(() => (window as any).__view.layoutDone === true);
    await mark(page, 1).click();
    await action(page, 'reply').click();
    await field(page).fill('never saved');
    const after = await page.evaluate(() => {
      (window as any).__mounted.destroy();
      (window as any).__mounted = null;
      return { counts: { ...(window as any).__counts }, nodes: document.getElementById('host')!.childNodes.length };
    });
    expect(after.counts).toEqual(empty);
    expect(after.nodes).toBe(0);
    expect(await page.evaluate(() => document.querySelectorAll('.bym-doc-comment-balloon, .bym-doc-comment-mark').length)).toBe(0);
    await page.waitForTimeout(200);
    expect(await page.evaluate(() => ({ ...(window as any).__counts }))).toEqual(empty);
  });
});
