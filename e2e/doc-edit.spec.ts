import type { Page } from '@playwright/test';
import { p, r, tbl, tc, tr } from '../test/doc-layout-fixtures';
import { docx, docSection, DOC_STYLES, field, picture, png, RED, BLUE, textBoxRun } from './fixtures/docx';
import { buildDocx } from '../test/docx-fixtures';
import { expect, test } from './helpers';

// The pages of the samples are 600 x 400 px with margins of 40 px, in Arial 12 pt.

interface MountOptions {
  readOnly?: boolean;
  locale?: string;
  zoom?: number | 'fit-width' | 'fit-page';
  toolbar?: boolean;
  shadow?: 'open' | 'closed';
  transform?: string;
}

interface Position {
  story: string;
  path: number[];
  offset: number;
}

const STYLES =
  DOC_STYLES +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>';

const sample = (body: string, extra: Parameters<typeof docx>[1] = {}) => docx(body, { styles: STYLES, ...extra });

async function mountEdit(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), options] as [number[], MountOptions]);
}

const laidOut = (page: Page) => page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });

/** Lets the view bring the screen up to date: its updates run in animation frames. */
const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

/** The paragraphs of the body, as text. */
const paragraphs = (page: Page) => page.evaluate(() => ((window as any).BymOffice.documentText((window as any).__view.document) as string).split('\n'));

const selection = (page: Page): Promise<{ anchor: Position; focus: Position } | undefined> => page.evaluate(() => (window as any).__view.getSelection());

const caretAt = async (page: Page): Promise<[number[], number]> => {
  const range = await selection(page);
  if (!range) throw new Error('no caret');
  expect(range.anchor).toEqual(range.focus);
  return [range.focus.path, range.focus.offset];
};

/** Where a place of the document is on screen: the box a caret has there. */
function boxOf(page: Page, path: number[], offset: number, story = 'body'): Promise<{ x: number; y: number; top: number; height: number; page: number }> {
  return page.evaluate(
    ([where, at, id]) => {
      const view = (window as any).__view;
      const box = view.lines.caret({ story: id, path: where, offset: at }, 'start', view.currentPage - 1);
      if (!box) throw new Error('the place is not on a page that is laid out');
      const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
      const element = root.querySelector(`.bym-doc-page[data-page="${box.page + 1}"]`);
      if (!element) throw new Error(`page ${box.page + 1} is not mounted`);
      const rect = element.getBoundingClientRect();
      const laid = view.pages[box.page];
      const sx = rect.width / laid.width;
      const sy = rect.height / laid.height;
      return { x: rect.left + box.x * sx, top: rect.top + box.top * sy, height: box.height * sy, y: rect.top + (box.top + box.height / 2) * sy, page: box.page + 1 };
    },
    [path, offset, story] as [number[], number, string],
  );
}

/** Clicks with the mouse where a place of the document is. */
async function clickAt(page: Page, path: number[], offset: number, options: { shift?: boolean; count?: number; story?: string } = {}): Promise<void> {
  const box = await boxOf(page, path, offset, options.story);
  if (options.shift) await page.keyboard.down('Shift');
  await page.mouse.click(box.x, box.y, { clickCount: options.count || 1 });
  if (options.shift) await page.keyboard.up('Shift');
  await settled(page);
}

/** The caret as the page shows it: the rectangle of its element, or null when it is not shown. */
const caretRect = async (page: Page) => {
  // The caret is placed in the next frame.
  await settled(page);
  return page.evaluate(() => {
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    const caret = root.querySelector('.bym-doc-caret') as HTMLElement | null;
    if (!caret || !caret.isConnected) return null;
    const rect = caret.getBoundingClientRect();
    return { x: rect.left + rect.width / 2, top: rect.top, height: rect.height, width: rect.width };
  });
};

const pageText = (page: Page, number: number) =>
  page.evaluate((n) => {
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    const element = root.querySelector(`.bym-doc-page[data-page="${n}"] .bym-doc-text`);
    return element ? element.textContent : null;
  }, number);

/** Formatting of the runs of a paragraph of the body: [text, "b" "i" "u" flags] for each. */
const runsOf = (page: Page, index: number) =>
  page.evaluate((i) => {
    const paragraph = (window as any).__view.document.body.blocks[i];
    return paragraph.children
      .filter((inline: any) => inline.type === 'run')
      .map((run: any) => [
        run.content.map((item: any) => (item.type === 'text' ? item.text : `<${item.type}>`)).join(''),
        (run.props.bold === true ? 'b' : '') + (run.props.italic === true ? 'i' : '') + (run.props.underline && run.props.underline.style !== 'none' ? 'u' : ''),
      ]);
  }, index);

const input = (page: Page) => page.locator('#host .bym-doc-input');
const toolbarButton = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);
const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
/** The key that moves by words: Option on Apple systems, Control elsewhere. */
const wordKey = (page: Page) => page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Alt' : 'Control'));
const commandKey = (page: Page) => page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));

/** Fires a clipboard event at the field for typing, the way the browser does for the keys, and returns what is on the clipboard afterwards. */
function clipboardEvent(page: Page, type: 'copy' | 'cut' | 'paste', give: { text?: string; html?: string; image?: number[] } = {}): Promise<{ prevented: boolean; text: string; html: string }> {
  return page.evaluate(
    ([kind, content]) => {
      const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
      const target = root.querySelector('.bym-doc-input')!;
      let data: any = new DataTransfer();
      if (content.text !== undefined) data.setData('text/plain', content.text);
      if (content.html !== undefined) data.setData('text/html', content.html);
      if (content.image) {
        const file = new File([new Uint8Array(content.image)], 'picture.png', { type: 'image/png' });
        let added = false;
        try {
          data.items.add(file);
          added = data.files.length > 0;
        } catch {
          added = false;
        }
        // Where a test cannot put a file on a DataTransfer, an object with the same members stands in for it.
        if (!added) data = { getData: () => '', setData: () => undefined, items: [{ kind: 'file', type: 'image/png', getAsFile: () => file }], files: [file] };
      }
      let event: Event | undefined;
      try {
        event = new ClipboardEvent(kind, { clipboardData: data instanceof DataTransfer ? data : undefined, bubbles: true, cancelable: true });
      } catch {
        event = undefined;
      }
      if (!event || (event as ClipboardEvent).clipboardData !== data) {
        event = new Event(kind, { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
      }
      const allowed = target.dispatchEvent(event);
      return { prevented: !allowed, text: data.getData('text/plain'), html: data.getData('text/html') };
    },
    [type, give] as ['copy' | 'cut' | 'paste', { text?: string; html?: string; image?: number[] }],
  );
}

/**
 * What an input method does to the field for typing, step by step: the text
 * it is composing takes the place of what it composed before, with the events
 * a browser sends for it. `android`: the events of a keyboard that composes
 * every word (no compositionupdate, a delete that cannot be cancelled).
 */
function compose(page: Page, steps: string[], end: string | null, android = false): Promise<void> {
  return page.evaluate(
    async ([texts, final, mobile]) => {
      const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
      const field = root.querySelector('.bym-doc-input') as HTMLTextAreaElement;
      const frame = () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
      const fire = (type: string, inputType: string, data: string | null) =>
        field.dispatchEvent(new InputEvent(type, { inputType, data, isComposing: true, bubbles: true, cancelable: false }));
      field.dispatchEvent(new CompositionEvent('compositionstart', { data: '', bubbles: true }));
      const start = field.selectionStart;
      let length = field.selectionEnd - field.selectionStart;
      const put = (text: string, inputType: string) => {
        fire('beforeinput', inputType, text);
        if (!mobile) field.dispatchEvent(new CompositionEvent('compositionupdate', { data: text, bubbles: true }));
        field.setRangeText(text, start, start + length, 'end');
        length = text.length;
        fire('input', inputType, text);
      };
      for (const text of texts) {
        // A shorter text on Android arrives as a deletion.
        put(text, mobile && text.length < length ? 'deleteContentBackward' : 'insertCompositionText');
        await frame();
      }
      if (final !== null && final !== texts[texts.length - 1]) put(final, 'insertCompositionText');
      field.dispatchEvent(new CompositionEvent('compositionend', { data: final === null ? '' : final, bubbles: true }));
      await frame();
    },
    [steps, end, android] as [string[], string | null, boolean],
  );
}

const SENTENCE = 'The quick brown fox jumps over the lazy dog and keeps running along the river bank. ';

function threeParagraphs(): Promise<Uint8Array> {
  return sample(p('First paragraph of the document.') + p('Second one, a bit longer than the first paragraph is.') + p('Third and last.'));
}

test.describe('document editor: caret and typing', () => {
  test('a click puts the caret where the text is; typing goes into the model and onto the page', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeVisible();
    await expect(input(page)).toHaveAttribute('role', 'textbox');
    await expect(input(page)).toHaveAttribute('aria-multiline', 'true');
    await expect(input(page)).toHaveAttribute('aria-label', 'Document editor');

    // No caret before the view has the focus.
    expect(await caretRect(page)).toBeNull();
    const wanted = await boxOf(page, [1], 7);
    await clickAt(page, [1], 7);
    expect(await caretAt(page)).toEqual([[1], 7]);
    await expect(input(page)).toBeFocused();
    const caret = await caretRect(page);
    expect(caret).not.toBeNull();
    expect(Math.abs(caret!.x - wanted.x)).toBeLessThan(1.5);
    expect(Math.abs(caret!.top - wanted.top)).toBeLessThan(1.5);
    // As high as the text of the line: Arial 12 pt is 16 px, and more than that with its descent.
    const zoom = await page.evaluate(() => (window as any).__view.zoom);
    expect(caret!.height).toBeGreaterThan(14 * zoom);
    expect(caret!.height).toBeLessThan(24 * zoom);

    await page.keyboard.type('NEW ');
    await settled(page);
    expect((await paragraphs(page))[1]).toBe('Second NEW one, a bit longer than the first paragraph is.');
    expect(await caretAt(page)).toEqual([[1], 11]);
    expect(await pageText(page, 1)).toContain('Second NEW one, a bit longer');
    // The caret moved right by what was typed.
    const after = await caretRect(page);
    expect(after!.x).toBeGreaterThan(caret!.x + 20 * zoom);
    expect(await page.evaluate(() => [(window as any).__view.isModified, (window as any).__changes > 0])).toEqual([true, true]);
  });

  test('Turkish characters, capitals and symbols are typed as they are', async ({ page }) => {
    await mountEdit(page, await sample(p('')));
    await laidOut(page);
    await clickAt(page, [0], 0);
    await page.keyboard.type('ığüşöç İĞÜŞÖÇ “â” €');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['ığüşöç İĞÜŞÖÇ “â” €']);
    expect(await pageText(page, 1)).toContain('ığüşöç İĞÜŞÖÇ');
  });

  test('a click in an empty paragraph, beside a line and in the margin finds the nearest place', async ({ page }) => {
    await mountEdit(page, await sample(p('One') + p('') + p('Three')));
    await laidOut(page);
    const empty = await boxOf(page, [1], 0);
    await page.mouse.click(empty.x + 150, empty.y);
    expect(await caretAt(page)).toEqual([[1], 0]);
    const caret = await caretRect(page);
    expect(Math.abs(caret!.x - empty.x)).toBeLessThan(1.5);
    // Far to the right of a short line: its end.
    const end = await boxOf(page, [0], 3);
    await page.mouse.click(end.x + 200, end.y);
    expect(await caretAt(page)).toEqual([[0], 3]);
    // In the left margin: the start of the line.
    const start = await boxOf(page, [2], 0);
    await page.mouse.click(start.x - 25, start.y);
    expect(await caretAt(page)).toEqual([[2], 0]);
    // Below the last line of the page: the last line.
    await page.mouse.click(start.x + 5, start.y + 120);
    expect((await caretAt(page))[0]).toEqual([2]);
  });

  test('Enter, Shift+Enter, Backspace and Delete, by characters and by words', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 5);
    await page.keyboard.press('Enter');
    expect((await paragraphs(page)).slice(0, 2)).toEqual(['First', ' paragraph of the document.']);
    expect(await caretAt(page)).toEqual([[1], 0]);
    // Backspace at the start of a paragraph joins it with the one before.
    await page.keyboard.press('Backspace');
    expect((await paragraphs(page))[0]).toBe('First paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 5]);
    await page.keyboard.press('Backspace');
    expect((await paragraphs(page))[0]).toBe('Firs paragraph of the document.');
    await page.keyboard.press('Delete');
    expect((await paragraphs(page))[0]).toBe('Firsparagraph of the document.');
    const word = await wordKey(page);
    await page.keyboard.press(`${word}+Delete`);
    expect((await paragraphs(page))[0]).toBe('Firsof the document.');
    await page.keyboard.press(`${word}+Backspace`);
    expect((await paragraphs(page))[0]).toBe('of the document.');
    // Delete at the end of a paragraph brings the next one up.
    await page.keyboard.press('End');
    await page.keyboard.press('Delete');
    expect(await paragraphs(page)).toEqual(['of the document.Second one, a bit longer than the first paragraph is.', 'Third and last.']);
    // Shift+Enter is a line break inside the paragraph.
    await page.keyboard.press('Shift+Enter');
    const model = await page.evaluate(() => (window as any).BymOffice.paragraphText((window as any).__view.document.body.blocks[0]));
    expect(model).toBe('of the document.\nSecond one, a bit longer than the first paragraph is.');
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks.length)).toBe(2);
    // Ctrl+Enter is a page break: what follows starts the next page.
    await page.keyboard.press('End');
    await page.keyboard.press(`${await commandKey(page)}+Enter`);
    await laidOut(page);
    expect(await page.evaluate(() => (window as any).__view.pageCount)).toBe(2);
  });

  test('arrow keys: characters, words, lines with the x kept, Home and End, the ends of the document', async ({ page }) => {
    const long = `${SENTENCE}${SENTENCE}${SENTENCE}`;
    await mountEdit(page, await sample(p(long) + p('Short.') + p(long)));
    await laidOut(page);
    await clickAt(page, [0], 10);
    await page.keyboard.press('ArrowRight');
    expect(await caretAt(page)).toEqual([[0], 11]);
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('ArrowLeft');
    expect(await caretAt(page)).toEqual([[0], 9]);
    const word = await wordKey(page);
    await page.keyboard.press(`${word}+ArrowRight`);
    // "The quick |brown": the start of the next word.
    expect(await caretAt(page)).toEqual([[0], 10]);
    await page.keyboard.press(`${word}+ArrowLeft`);
    expect(await caretAt(page)).toEqual([[0], 4]);

    // Down a line: the caret keeps its x as well as the text allows, and gets it back after a short line.
    await clickAt(page, [0], 30);
    const first = (await caretRect(page))!;
    await page.keyboard.press('ArrowDown');
    const second = (await caretRect(page))!;
    expect(second.top).toBeGreaterThan(first.top + first.height * 0.8);
    expect(Math.abs(second.x - first.x)).toBeLessThan(8);
    expect((await caretAt(page))[1]).toBeGreaterThan(40);
    // To the last line of the paragraph, then onto the short one, then into the third paragraph.
    for (let i = 0; i < 12; i++) {
      const [path] = await caretAt(page);
      if (path[0] === 1) break;
      await page.keyboard.press('ArrowDown');
    }
    expect(await caretAt(page)).toEqual([[1], 6]);
    await page.keyboard.press('ArrowDown');
    const third = (await caretRect(page))!;
    expect((await caretAt(page))[0]).toEqual([2]);
    expect(Math.abs(third.x - first.x)).toBeLessThan(8);
    await page.keyboard.press('ArrowUp');
    expect(await caretAt(page)).toEqual([[1], 6]);

    // Home and End are the ends of the line the caret is on, not of the paragraph.
    await clickAt(page, [0], 30);
    await page.keyboard.press('End');
    const [, lineEnd] = await caretAt(page);
    expect(lineEnd).toBeGreaterThan(40);
    expect(lineEnd).toBeLessThan(long.length);
    const atEnd = (await caretRect(page))!;
    expect(Math.abs(atEnd.top - first.top)).toBeLessThan(1.5);
    expect(atEnd.x).toBeGreaterThan(first.x + 100 * (await page.evaluate(() => (window as any).__view.zoom)));
    await page.keyboard.press('Home');
    expect(await caretAt(page)).toEqual([[0], 0]);
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+End`);
    expect(await caretAt(page)).toEqual([[2], long.length]);
    await page.keyboard.press(`${command}+Home`);
    expect(await caretAt(page)).toEqual([[0], 0]);
  });

  test('Shift with the arrows and with a click extends the selection; typing replaces it', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 6);
    for (let i = 0; i < 9; i++) await page.keyboard.press('Shift+ArrowRight');
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('paragraph');
    // No caret while text is selected.
    await settled(page);
    expect(await caretRect(page)).toBeNull();
    await clickAt(page, [1], 6, { shift: true });
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('paragraph of the document.\nSecond');
    await page.keyboard.type('X');
    expect(await paragraphs(page)).toEqual(['First X one, a bit longer than the first paragraph is.', 'Third and last.']);
    expect(await caretAt(page)).toEqual([[0], 7]);
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+a`);
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('First X one, a bit longer than the first paragraph is.\nThird and last.');
    await page.keyboard.press('Backspace');
    expect(await paragraphs(page)).toEqual(['']);
  });

  test('a double click selects a word, a triple click the paragraph, a drag what it passes', async ({ page, isMobile }) => {
    test.skip(isMobile, 'selecting by dragging is for the mouse');
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [1], 9, { count: 2 });
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('one');
    const box = await boxOf(page, [1], 9);
    await page.mouse.click(box.x, box.y, { clickCount: 3 });
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('Second one, a bit longer than the first paragraph is.');
    // Let the clicks be over before the drag starts.
    await page.waitForTimeout(600);
    const from = await boxOf(page, [0], 6);
    const to = await boxOf(page, [2], 5);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move((from.x + to.x) / 2, (from.y + to.y) / 2, { steps: 3 });
    await page.mouse.move(to.x, to.y, { steps: 3 });
    await page.mouse.up();
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('paragraph of the document.\nSecond one, a bit longer than the first paragraph is.\nThird');
    const range = await selection(page);
    expect(range!.anchor).toEqual({ story: 'body', path: [0], offset: 6 });
    expect(range!.focus).toEqual({ story: 'body', path: [2], offset: 5 });
  });

  test('a drag held at the bottom edge scrolls the pages and the selection follows', async ({ page, isMobile }) => {
    test.skip(isMobile, 'selecting by dragging is for the mouse');
    let body = '';
    for (let i = 1; i <= 60; i++) body += p(`Paragraph number ${i} of the long document.`);
    await mountEdit(page, await sample(body));
    await laidOut(page);
    const from = await boxOf(page, [0], 3);
    const host = (await page.locator('#host .bym-doc-scroller').boundingBox())!;
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(from.x + 40, host.y + host.height - 4, { steps: 4 });
    await page.waitForFunction(() => (window as any).__view.getViewState().scrollY > 300);
    await page.mouse.up();
    const range = await selection(page);
    expect(range!.anchor).toEqual({ story: 'body', path: [0], offset: 3 });
    expect(range!.focus.path[0]).toBeGreaterThan(12);
    expect(await page.evaluate(() => (window as any).__counts.intervals)).toBe(0);
  });

  test('text typed at the foot of a page flows onto the next page, and a new page appears when it is needed', async ({ page }) => {
    let body = '';
    for (let i = 1; i <= 17; i++) body += p(`Line ${i}`);
    await mountEdit(page, await sample(body));
    await laidOut(page);
    const before = await page.evaluate(() => (window as any).__view.pageCount);
    expect(before).toBe(1);
    await clickAt(page, [16], 7);
    // Enter adds lines until one does not fit.
    for (let i = 0; i < 4; i++) {
      await page.keyboard.press('Enter');
      await page.keyboard.type(`More ${i}`);
    }
    await laidOut(page);
    await settled(page);
    expect(await page.evaluate(() => (window as any).__view.pageCount)).toBe(2);
    const text = await pageText(page, 2);
    expect(text).toContain('More 3');
    // The caret went with its text, and the view scrolled to it.
    const caret = await caretRect(page);
    expect(caret).not.toBeNull();
    const scroller = (await page.locator('#host .bym-doc-scroller').boundingBox())!;
    expect(caret!.top).toBeGreaterThanOrEqual(scroller.y - 1);
    expect(caret!.top + caret!.height).toBeLessThanOrEqual(scroller.y + scroller.height + 1);
    expect((await boxOf(page, (await caretAt(page))[0], (await caretAt(page))[1])).page).toBe(2);
  });
});

test.describe('document editor: formatting', () => {
  test('bold and italic on a selection, and at a caret for what is typed next', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 6);
    for (let i = 0; i < 9; i++) await page.keyboard.press('Shift+ArrowRight');
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+b`);
    expect(await runsOf(page, 0)).toEqual([
      ['First ', ''],
      ['paragraph', 'b'],
      [' of the document.', ''],
    ]);
    await settled(page);
    await expect(toolbarButton(page, 'bold')).toHaveAttribute('aria-pressed', 'true');
    // A selection that is only partly bold: neither on nor off.
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Shift+ArrowRight');
    await settled(page);
    await expect(toolbarButton(page, 'bold')).toHaveAttribute('aria-pressed', 'mixed');

    // At a caret: the next text is italic, the text before it is not changed.
    await clickAt(page, [2], 5);
    await page.keyboard.press(`${command}+i`);
    await settled(page);
    await expect(toolbarButton(page, 'italic')).toHaveAttribute('aria-pressed', 'true');
    await page.keyboard.type('XY');
    expect(await runsOf(page, 2)).toEqual([
      ['Third', ''],
      ['XY', 'i'],
      [' and last.', ''],
    ]);
    // A format chosen at a caret is dropped when the caret moves away before anything is typed.
    await page.keyboard.press(`${command}+u`);
    await page.keyboard.press('ArrowRight');
    await page.keyboard.type('Z');
    expect((await runsOf(page, 2)).some((run: string[]) => run[1].indexOf('u') >= 0)).toBe(false);

    // The buttons do the same.
    await clickAt(page, [1], 0);
    for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight');
    await toolbarButton(page, 'underline').click();
    expect((await runsOf(page, 1))[0]).toEqual(['Second', 'u']);
    await expect(input(page)).toBeFocused();
  });

  test('paragraph style, font, size, colour, alignment and lists from the toolbar', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [1], 3);
    const style = page.locator('#host .bym-doc-toolbar .bym-tb-style');
    await settled(page);
    await expect(style).toHaveValue('Normal');
    expect(await style.locator('option:not([hidden])').allTextContents()).toEqual(['Normal', 'heading 1', 'heading 2']);
    await style.selectOption('Heading1');
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks[1].props.style)).toBe('Heading1');
    await settled(page);
    await expect(style).toHaveValue('Heading1');
    // The heading is taller: the paragraph after it moved down.
    await laidOut(page);

    await toolbarButton(page, 'alignCenter').click();
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks[1].props.align)).toBe('center');
    await settled(page);
    await expect(toolbarButton(page, 'alignCenter')).toHaveAttribute('aria-pressed', 'true');
    await expect(toolbarButton(page, 'alignLeft')).toHaveAttribute('aria-pressed', 'false');
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+j`);
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks[1].props.align)).toBe('both');

    // A list, and its level.
    await clickAt(page, [2], 0);
    await toolbarButton(page, 'numberedList').click();
    await laidOut(page);
    await settled(page);
    await expect(toolbarButton(page, 'numberedList')).toHaveAttribute('aria-pressed', 'true');
    expect(await pageText(page, 1)).toContain('1.');
    await toolbarButton(page, 'increaseIndent').click();
    expect(await page.evaluate(() => (window as any).__view.selectionFormat.level)).toBe(1);
    await toolbarButton(page, 'bulletList').click();
    await settled(page);
    await expect(toolbarButton(page, 'bulletList')).toHaveAttribute('aria-pressed', 'true');
    await toolbarButton(page, 'bulletList').click();
    expect(await page.evaluate(() => (window as any).__view.selectionFormat.list)).toBe('none');

    // Font, size and colour of a selection.
    await clickAt(page, [0], 0);
    for (let i = 0; i < 5; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.locator('#host .bym-doc-toolbar .bym-tb-size').selectOption('18');
    await page.locator('#host .bym-doc-toolbar .bym-tb-font').selectOption('Georgia');
    await toolbarButton(page, 'textColor').click();
    await page.locator('#host .bym-popover [data-color="#FF0000"]').click();
    const props = await page.evaluate(() => (window as any).__view.document.body.blocks[0].children[0].props);
    expect(props.size).toBe(36);
    expect(props.fonts.ascii).toBe('Georgia');
    expect(props.color.rgb || props.color.value || JSON.stringify(props.color)).toContain('FF0000');
    await settled(page);
    await expect(page.locator('#host .bym-doc-toolbar .bym-tb-size')).toHaveValue('18');
    await expect(toolbarButton(page, 'textColor').locator('.bym-icon-bar')).toHaveAttribute('data-color', '#FF0000');
    await toolbarButton(page, 'highlight').click();
    await page.locator('#host .bym-popover [data-color="#FFFF00"]').click();
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks[0].children[0].props.highlight)).toBe('yellow');
  });

  test('a table from the grid; Tab goes from cell to cell and adds a row after the last one; rows and columns', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await settled(page);
    // Outside a table there are no rows and columns to change.
    await clickAt(page, [0], 32);
    await settled(page);
    await expect(toolbarButton(page, 'rowsColumns')).toBeDisabled();
    await toolbarButton(page, 'insertTable').click();
    await page.locator('#host .bym-popover [data-color="2x3"]').click();
    await laidOut(page);
    const table = () =>
      page.evaluate(() => {
        const block = (window as any).__view.document.body.blocks.find((b: any) => b.type === 'table');
        return block.rows.map((row: any) => row.cells.map((cell: any) => (window as any).BymOffice.paragraphText(cell.blocks[0])));
      });
    expect(await table()).toEqual([
      ['', '', ''],
      ['', '', ''],
    ]);
    const tableIndex = await page.evaluate(() => (window as any).__view.document.body.blocks.findIndex((b: any) => b.type === 'table'));
    expect(await caretAt(page)).toEqual([[tableIndex, 0, 0, 0], 0]);
    await page.keyboard.type('a');
    await page.keyboard.press('Tab');
    await page.keyboard.type('b');
    await page.keyboard.press('Tab');
    await page.keyboard.type('c');
    await page.keyboard.press('Tab');
    await page.keyboard.type('d');
    expect((await selection(page))!.focus.path).toEqual([tableIndex, 1, 0, 0]);
    await page.keyboard.press('Shift+Tab');
    // The content of the cell gone back to is selected: typing replaces it.
    await page.keyboard.type('C');
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await page.keyboard.press('Tab');
    await page.keyboard.type('f');
    // Tab in the last cell makes a new row.
    await page.keyboard.press('Tab');
    await page.keyboard.type('g');
    expect(await table()).toEqual([
      ['a', 'b', 'C'],
      ['d', '', 'f'],
      ['g', '', ''],
    ]);
    await laidOut(page);
    await settled(page);
    await expect(toolbarButton(page, 'rowsColumns')).toBeEnabled();
    await toolbarButton(page, 'rowsColumns').click();
    await page.locator('#host .bym-popover [role="menuitem"]', { hasText: 'Insert column right' }).click();
    expect((await table())[0]).toEqual(['a', '', 'b', 'C']);
    await toolbarButton(page, 'rowsColumns').click();
    await page.locator('#host .bym-popover [role="menuitem"]', { hasText: 'Delete row' }).click();
    expect((await table()).length).toBe(2);
    // The text of the cells is on the page.
    await laidOut(page);
    await settled(page);
    expect(await pageText(page, 1)).toContain('a');
    // Outside a table Tab types a tab.
    await clickAt(page, [0], 5);
    await page.keyboard.press('Tab');
    expect(await page.evaluate(() => (window as any).BymOffice.paragraphText((window as any).__view.document.body.blocks[0]))).toBe('First\t paragraph of the document.');
  });

  test('a link from the field of the toolbar and from Ctrl+K; an address that is not safe is refused', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 6);
    for (let i = 0; i < 9; i++) await page.keyboard.press('Shift+ArrowRight');
    await toolbarButton(page, 'insertLink').click();
    const field = page.locator('#host .bym-doc-linkinput');
    await expect(field).toBeFocused();
    await field.fill('javascript:alert(1)');
    await page.keyboard.press('Enter');
    await expect(page.locator('#host .bym-doc-linkfield')).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__notices.map((n: any) => n.message))).toContain('This address cannot be used as a link');
    expect(await page.evaluate(() => (window as any).__changes)).toBe(0);
    const linked = () =>
      page.evaluate(() => {
        const paragraph = (window as any).__view.document.body.blocks[0];
        return paragraph.children.filter((inline: any) => inline.wrap && inline.wrap.kind === 'hyperlink').map((inline: any) => inline.content.map((item: any) => item.text).join(''));
      });
    expect(await linked()).toEqual([]);
    await expect(input(page)).toBeFocused();

    await page.keyboard.press(`${await commandKey(page)}+k`);
    await expect(field).toBeFocused();
    await field.fill('https://example.com/page');
    await page.locator('#host .bym-doc-linkapply').click();
    expect(await linked()).toEqual(['paragraph']);
    await laidOut(page);
    await settled(page);
    // The link is on the page; in an editor a plain click puts the caret, Ctrl+click follows it.
    const link = page.locator('#host .bym-doc-link');
    await expect(link).toHaveCount(1);
    await expect(input(page)).toBeFocused();
    const box = await boxOf(page, [0], 9);
    await page.mouse.click(box.x, box.y);
    expect(await page.evaluate(() => (window as any).__opened.length)).toBe(0);
    await page.keyboard.down(modifier === 'Meta' ? 'Meta' : 'Control');
    await page.mouse.click(box.x, box.y);
    await page.keyboard.up(modifier === 'Meta' ? 'Meta' : 'Control');
    expect(await page.evaluate(() => (window as any).__opened.map((o: any) => o.url))).toEqual(['https://example.com/page']);
    // Escape closes the field and gives the focus back to its button.
    await toolbarButton(page, 'insertLink').click();
    await page.keyboard.press('Escape');
    await expect(page.locator('#host .bym-doc-linkfield')).toHaveCount(0);
    await expect(toolbarButton(page, 'insertLink')).toBeFocused();
  });

  test('the toolbar is one Tab stop with the arrow keys inside, labelled in Turkish and English', async ({ page }) => {
    await mountEdit(page, await threeParagraphs(), { locale: 'tr' });
    await laidOut(page);
    await expect(page.locator('#host .bym-doc-toolbar')).toHaveAttribute('aria-label', 'Biçimlendirme');
    await expect(toolbarButton(page, 'bulletList')).toHaveAttribute('aria-label', 'Madde imli liste');
    await expect(page.locator('#host .bym-doc-toolbar .bym-tb-style')).toHaveAttribute('aria-label', 'Paragraf stili');
    await expect(input(page)).toHaveAttribute('aria-label', 'Belge düzenleyici');
    const stops = await page.evaluate(() => Array.from(document.querySelectorAll('#host .bym-doc-toolbar button')).filter((b) => (b as HTMLElement).tabIndex === 0).length);
    expect(stops).toBe(1);
    // Nothing is undone yet: the first enabled button is the stop.
    await toolbarButton(page, 'bold').focus();
    await page.keyboard.press('ArrowRight');
    await expect(toolbarButton(page, 'italic')).toBeFocused();
    await page.keyboard.press('End');
    await expect(toolbarButton(page, 'pageBreak')).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(input(page)).toBeFocused();
    await page.evaluate(() => (window as any).__mounted.editor.setLocale('en'));
    await expect(toolbarButton(page, 'bulletList')).toHaveAttribute('aria-label', 'Bulleted list');
    // Narrower than its buttons, the toolbar scrolls sideways instead of wrapping.
    const sizes = await page.evaluate(() => {
      const bar = document.querySelector('#host .bym-doc-toolbar') as HTMLElement;
      return { height: bar.offsetHeight, scroll: bar.scrollWidth, width: bar.clientWidth };
    });
    expect(sizes.height).toBeLessThan(60);
    expect(sizes.scroll).toBeGreaterThanOrEqual(sizes.width);
  });
});

test.describe('document editor: history', () => {
  test('undo and redo bring the text and the caret back; typing is one step per word', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await settled(page);
    await expect(toolbarButton(page, 'undo')).toBeDisabled();
    await clickAt(page, [0], 5);
    await page.keyboard.type('abc');
    await page.keyboard.type(' def');
    expect((await paragraphs(page))[0]).toBe('Firstabc def paragraph of the document.');
    await settled(page);
    await expect(toolbarButton(page, 'undo')).toBeEnabled();
    const command = await commandKey(page);
    // The word typed last goes first, then the one before it with its space.
    await page.keyboard.press(`${command}+z`);
    expect((await paragraphs(page))[0]).toBe('Firstabc  paragraph of the document.');
    await page.keyboard.press(`${command}+z`);
    expect((await paragraphs(page))[0]).toBe('First paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 5]);
    expect(await page.evaluate(() => [(window as any).__view.canUndo, (window as any).__view.canRedo])).toEqual([false, true]);
    await page.keyboard.press(`${command}+y`);
    await page.keyboard.press(`${command}+Shift+z`);
    expect((await paragraphs(page))[0]).toBe('Firstabc def paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 12]);
    expect(await pageText(page, 1)).toContain('Firstabc def paragraph');

    // Moving the caret ends a step: the two letters are undone one by one.
    await page.keyboard.type('x');
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.type('y');
    await page.keyboard.press(`${command}+z`);
    expect((await paragraphs(page))[0]).toBe('Firstabc defx paragraph of the document.');
    // A structural change and its undo: the selection that was replaced comes back as it was.
    await clickAt(page, [1], 0);
    for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Enter');
    expect((await paragraphs(page)).length).toBe(4);
    await toolbarButton(page, 'undo').click();
    expect((await paragraphs(page)).length).toBe(3);
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('Second');
    await toolbarButton(page, 'redo').click();
    expect((await paragraphs(page)).length).toBe(4);
  });

  test('read-only: nothing changes the document, there is no caret and no toolbar; the selection can still be copied', async ({ page }) => {
    await mountEdit(page, await threeParagraphs(), { readOnly: true });
    await laidOut(page);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
    await expect(input(page)).toHaveAttribute('inputmode', 'none');
    const before = await paragraphs(page);
    const point = await page.evaluate(() => {
      const element = document.querySelector('#host .bym-doc-page[data-page="1"]')!;
      const rect = element.getBoundingClientRect();
      return { x: rect.left + rect.width * 0.3, y: rect.top + rect.height * 0.12 };
    });
    await page.mouse.click(point.x, point.y);
    await page.keyboard.type('typed');
    await page.keyboard.press('Enter');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('Delete');
    await clipboardEvent(page, 'paste', { text: 'pasted' });
    await clipboardEvent(page, 'cut');
    expect(
      await page.evaluate(() => {
        const view = (window as any).__view;
        return [view.insertText('x'), view.format({ bold: true }), view.insertTable(2, 2), view.undo(), view.insertLink('https://example.com'), view.toggleList('bullet')];
      }),
    ).toEqual([false, false, false, false, false, false]);
    await settled(page);
    expect(await paragraphs(page)).toEqual(before);
    expect(await caretRect(page)).toBeNull();
    expect(await page.evaluate(() => [(window as any).__view.isModified, (window as any).__changes, (window as any).__view.readOnly])).toEqual([false, 0, true]);
    await page.keyboard.press(`${await commandKey(page)}+a`);
    expect((await clipboardEvent(page, 'copy')).text).toBe(before.join('\n'));
    // Editing can be turned on later.
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await laidOut(page);
    await expect(page.locator('#host .bym-doc-toolbar')).toBeVisible();
    await clickAt(page, [0], 5);
    await page.keyboard.type('!');
    expect((await paragraphs(page))[0]).toBe('First! paragraph of the document.');
  });

  test('the public API: selection, insertText, format, focus, commitEditing, onSelectionChange', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      const out: any = {};
      view.setSelection({ story: 'body', path: [2], offset: 5 });
      out.caret = view.getSelection();
      out.inserted = view.insertText(' part');
      view.setSelection({ anchor: { story: 'body', path: [2], offset: 0 }, focus: { story: 'body', path: [2], offset: 5 } });
      out.formatted = view.format({ bold: true, size: 30 });
      out.format = view.selectionFormat;
      out.text = view.getSelectedText();
      try {
        view.setSelection({ story: 'body', path: [9], offset: 0 });
      } catch (error) {
        out.error = (error as Error).name;
      }
      view.commitEditing();
      out.modified = view.isModified;
      view.markSaved();
      out.saved = view.isModified;
      return out;
    });
    expect(result.caret).toEqual({ anchor: { story: 'body', path: [2], offset: 5 }, focus: { story: 'body', path: [2], offset: 5 } });
    expect(result.inserted).toBe(true);
    expect(result.formatted).toBe(true);
    expect(result.format.bold).toBe(true);
    expect(result.format.size).toBe(30);
    expect(result.text).toBe('Third');
    expect(result.error).toBe('RangeError');
    expect([result.modified, result.saved]).toEqual([true, false]);
    expect((await paragraphs(page))[2]).toBe('Third part and last.');
    // Ctrl+S saves through mount(): the bytes hold the change.
    await page.evaluate(() => (window as any).__view.focus());
    await expect(input(page)).toBeFocused();
    await page.keyboard.press(`${await commandKey(page)}+s`);
    await page.waitForFunction(() => (window as any).__saved.length === 1);
    const reopened = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const bytes = await (window as any).__mounted.save();
      const file = await api.open(bytes);
      return api.documentText(await file.document());
    });
    expect(reopened.split('\n')[2]).toBe('Third part and last.');
  });
});

test.describe('document editor: input methods', () => {
  test('a composition is typed into the document as it goes, underlined, and is one step to undo', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 5);
    // Before the composition ends the text is in the document already, with a line under it.
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      field.dispatchEvent(new CompositionEvent('compositionstart', { data: '', bubbles: true }));
      const start = field.selectionStart;
      field.dispatchEvent(new CompositionEvent('compositionupdate', { data: 'にほ', bubbles: true }));
      field.setRangeText('にほ', start, start, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'insertCompositionText', data: 'にほ', isComposing: true, bubbles: true }));
    });
    await settled(page);
    expect((await paragraphs(page))[0]).toBe('Firstにほ paragraph of the document.');
    const underline = await page.evaluate(() => {
      const element = document.querySelector('#host .bym-doc-compose') as HTMLElement | null;
      return element && element.isConnected ? element.getBoundingClientRect().width : 0;
    });
    expect(underline).toBeGreaterThan(10);
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const end = field.selectionStart;
      field.dispatchEvent(new CompositionEvent('compositionupdate', { data: '日本', bubbles: true }));
      field.setRangeText('日本', end - 2, end, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'insertCompositionText', data: '日本', isComposing: true, bubbles: true }));
      field.dispatchEvent(new CompositionEvent('compositionend', { data: '日本', bubbles: true }));
    });
    await settled(page);
    expect((await paragraphs(page))[0]).toBe('First日本 paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 7]);
    expect(await page.evaluate(() => document.querySelectorAll('#host .bym-doc-compose').length)).toBe(0);
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+z`);
    expect((await paragraphs(page))[0]).toBe('First paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 5]);

    // A composition in several steps, the helper's way; then one that is cancelled.
    await compose(page, ['k', 'ka', 'か', 'かn', 'かん'], '漢');
    expect((await paragraphs(page))[0]).toBe('First漢 paragraph of the document.');
    await compose(page, ['x', 'xy'], null);
    await page.evaluate(() => {
      // The browser takes the composed text out of the field when the composition is cancelled.
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      void field;
    });
    await page.keyboard.type('!');
    expect((await paragraphs(page))[0]).toContain('First漢');
  });

  test('a cancelled composition takes its text back', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 5);
    await compose(page, ['a', 'ab', 'abc', ''], '');
    expect((await paragraphs(page))[0]).toBe('First paragraph of the document.');
    expect(await caretAt(page)).toEqual([[0], 5]);
    await page.keyboard.type('ok');
    expect((await paragraphs(page))[0]).toBe('Firstok paragraph of the document.');
  });

  test('an Android keyboard that composes every word, deletes in it and replaces it', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [2], 15);
    // "hello" typed letter by letter, a letter deleted, then the suggestion taken.
    await compose(page, [' ', ' h', ' he', ' hel', ' helo', ' hel', ' hell'], ' hello', true);
    expect((await paragraphs(page))[2]).toBe('Third and last. hello');
    expect(await caretAt(page)).toEqual([[2], 21]);
    // A Backspace that cannot be cancelled, outside a composition: the field loses a character and so does the document.
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.dispatchEvent(new InputEvent('beforeinput', { inputType: 'deleteContentBackward', bubbles: true, cancelable: false }));
      field.setRangeText('', at - 1, at, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentBackward', bubbles: true }));
    });
    expect((await paragraphs(page))[2]).toBe('Third and last. hell');
    // The correction a phone makes: a word before the caret is replaced as a whole.
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertReplacementText', data: 'help', bubbles: true, cancelable: false }));
      field.setRangeText('help', at - 4, at, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'insertReplacementText', data: 'help', bubbles: true }));
    });
    expect((await paragraphs(page))[2]).toBe('Third and last. help');
    expect(await caretAt(page)).toEqual([[2], 20]);
    // At the start of a paragraph the deletion that cannot be cancelled joins it with the one before.
    await clickAt(page, [1], 0);
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.setRangeText('', at - 1, at, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentBackward', bubbles: true }));
    });
    expect((await paragraphs(page)).length).toBe(2);
    expect(await caretAt(page)).toEqual([[0], 32]);
  });

  test('the field for typing sits at the caret and mirrors the text around it', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [1], 10);
    await settled(page);
    const state = await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const rect = field.getBoundingClientRect();
      return { value: field.value, start: field.selectionStart, end: field.selectionEnd, x: rect.left, y: rect.top, readOnly: field.readOnly, mode: field.getAttribute('inputmode') };
    });
    // The paragraph starts in the mirror (a line end stands for its start), and 40 characters after the caret are in it.
    expect(state.value).toBe('\nSecond one, a bit longer than the first paragraph ');
    expect([state.start, state.end]).toEqual([11, 11]);
    expect([state.readOnly, state.mode]).toEqual([false, 'text']);
    const caret = (await caretRect(page))!;
    expect(Math.abs(state.x - caret.x)).toBeLessThan(3);
    expect(Math.abs(state.y - caret.top)).toBeLessThan(3);
  });
});

test.describe('document editor: clipboard', () => {
  test('copy gives text and HTML made from escaped text; a paste of the own copy keeps the formatting', async ({ page }) => {
    await mountEdit(page, await sample(p(r('Plain ') + r('bold & <b>', '<w:b/>') + r(' and ') + r('italic', '<w:i/><w:color w:val="FF0000"/>') + r(' end.')) + p('Second.', '<w:jc w:val="center"/>') + p('')));
    await laidOut(page);
    await clickAt(page, [0], 6);
    await clickAt(page, [1], 3, { shift: true });
    const copied = await clipboardEvent(page, 'copy');
    expect(copied.prevented).toBe(true);
    expect(copied.text).toBe('bold & <b> and italic end.\nSec');
    expect(copied.html).toContain('<b>bold &amp; &lt;b&gt;</b>');
    expect(copied.html).toContain('<i>italic</i>');
    expect(copied.html).toContain('color:#FF0000');
    expect(copied.html).not.toContain('<b>bold & <b>');
    expect(copied.html).toMatch(/data-bym-doc-copy="[a-z0-9]+"/);

    await clickAt(page, [2], 0);
    const pasted = await clipboardEvent(page, 'paste', { text: copied.text, html: copied.html });
    expect(pasted.prevented).toBe(true);
    const all = await paragraphs(page);
    expect(all.slice(2)).toEqual(['bold & <b> and italic end.', 'Sec']);
    expect(await runsOf(page, 2)).toEqual([
      ['bold & <b>', 'b'],
      [' and ', ''],
      ['italic', 'i'],
      [' end.', ''],
    ]);
    // One step: a single undo takes the whole paste back.
    await page.keyboard.press(`${await commandKey(page)}+z`);
    expect(await paragraphs(page)).toEqual(['Plain bold & <b> and italic end.', 'Second.', '']);

    // The same text from elsewhere (no mark of this editor) is plain text, whatever its HTML says.
    await clickAt(page, [2], 0);
    await clipboardEvent(page, 'paste', { text: 'bold & <b>', html: '<b onclick="window.__hacked=1">bold</b><img src=x onerror="window.__hacked=1">' });
    expect((await paragraphs(page))[2]).toBe('bold & <b>');
    expect(await runsOf(page, 2)).toEqual([['bold & <b>', '']]);
    expect(await page.evaluate(() => (window as any).__hacked)).toBeUndefined();
  });

  test('cut removes the selection; text of several lines is pasted as paragraphs; control characters are dropped', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 0);
    for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight');
    const cut = await clipboardEvent(page, 'cut');
    expect(cut.text).toBe('First ');
    expect((await paragraphs(page))[0]).toBe('paragraph of the document.');
    await clipboardEvent(page, 'paste', { text: 'one\r\ntwo\u0000\u0007\n\nfour\tend ' });
    expect((await paragraphs(page)).slice(0, 4)).toEqual(['one', 'two', '', 'four\tend paragraph of the document.']);
    expect(await caretAt(page)).toEqual([[3], 9]);
    // Over a selection: the selection goes.
    await page.keyboard.press(`${await commandKey(page)}+a`);
    await clipboardEvent(page, 'paste', { text: 'all new' });
    expect(await paragraphs(page)).toEqual(['all new']);
  });

  test('a picture on the clipboard is put into the text, no wider than the text is', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [1], 0);
    // 1200 px wide: the text area is 520 px.
    await clipboardEvent(page, 'paste', { image: Array.from(png(1200, 300, RED, BLUE)) });
    await page.waitForFunction(() => (window as any).__changes > 0);
    const drawing = await page.evaluate(() => {
      const doc = (window as any).__view.document;
      const paragraph = doc.body.blocks[1];
      let found: any;
      for (const inline of paragraph.children) for (const item of inline.content || []) if (item.type === 'drawing') found = item;
      return { has: !!found, media: doc.media ? doc.media.size : 0 };
    });
    expect(drawing).toEqual({ has: true, media: 1 });
    await laidOut(page);
    // It is painted: red on the left, blue on the right, 520 px wide and 130 px high at 100%.
    await page.waitForFunction(() => {
      const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
      const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      let red = 0;
      let blue = 0;
      for (let i = 0; i < data.length; i += 4) {
        if (data[i] > 180 && data[i + 1] < 60 && data[i + 2] < 60) red++;
        else if (data[i + 2] > 180 && data[i] < 60 && data[i + 1] < 60) blue++;
      }
      const share = (520 * 130) / (600 * 400);
      const total = canvas.width * canvas.height;
      return red > total * share * 0.4 && blue > total * share * 0.4 && red + blue < total * share * 1.1;
    });
    // Bytes that are not a picture are refused, whatever the clipboard calls them.
    await clipboardEvent(page, 'paste', { image: Array.from(new TextEncoder().encode('<svg onload="alert(1)"></svg>             ')) });
    await page.waitForFunction(() => (window as any).__notices.some((n: any) => n.message === 'This picture cannot be inserted'));
    expect(await page.evaluate(() => (window as any).__view.document.media.size)).toBe(1);
  });

  test('a paste that is too long is refused', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 0);
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input')!;
      const event = new Event('paste', { bubbles: true, cancelable: true });
      const huge = new Array(5000002).join('x');
      Object.defineProperty(event, 'clipboardData', { value: { getData: (type: string) => (type === 'text/plain' ? huge : ''), items: [], files: [] } });
      field.dispatchEvent(event);
    });
    expect((await paragraphs(page))[0]).toBe('First paragraph of the document.');
    expect(await page.evaluate(() => (window as any).__notices.map((n: any) => n.message))).toContain('The text on the clipboard is too long to paste');
  });
});

test.describe('document editor: pictures and tables in the selection', () => {
  test('a press on a picture selects it; the caret stands before or after it; Delete removes it', async ({ page, isMobile }) => {
    test.skip(isMobile, 'pressing with the mouse');
    const bytes = await sample(p(r('ab') + picture(120, 60) + r('cd')), { rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', png(40, 20, RED, BLUE)]] });
    await mountEdit(page, bytes);
    await laidOut(page);
    const before = await boxOf(page, [0], 2);
    // The caret after "ab" is after the picture, which takes no offset of its own; with `before` it is on its left.
    const shape = await page.evaluate(() => {
      const view = (window as any).__view;
      const at = (before: boolean) => view.lines.caret({ story: 'body', path: [0], offset: 2, before: before || undefined });
      return { left: at(true).x, right: at(false).x };
    });
    expect(shape.right - shape.left).toBeGreaterThan(100);
    const zoom = await page.evaluate(() => (window as any).__view.zoom);
    const middle = { x: before.x - (shape.right - shape.left) * 0.6 * zoom, y: before.y - 20 * zoom };
    await page.mouse.click(middle.x, middle.y);
    const range = (await selection(page))!;
    expect(range.anchor).toEqual({ story: 'body', path: [0], offset: 2, before: true });
    expect(range.focus).toEqual({ story: 'body', path: [0], offset: 2 });
    // No caret: the picture is selected, and painted so.
    expect(await caretRect(page)).toBeNull();
    const blue = await page.evaluate(() => {
      const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
      const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
      let tinted = 0;
      for (let i = 0; i < data.length; i += 4) if (data[i + 2] > 180 && data[i] < 200 && data[i] > 40 && data[i + 1] > 40) tinted++;
      return tinted;
    });
    expect(blue).toBeGreaterThan(100);
    await page.keyboard.press('ArrowLeft');
    expect((await selection(page))!.focus).toEqual({ story: 'body', path: [0], offset: 2, before: true });
    // Typing before the picture puts the text before it.
    await page.keyboard.type('X');
    const items = () =>
      page.evaluate(() =>
        (window as any).__view.document.body.blocks[0].children.flatMap((inline: any) => (inline.content || []).map((item: any) => (item.type === 'text' ? item.text : `<${item.type}>`))),
      );
    expect((await items()).join('')).toBe('abX<drawing>cd');
    // Still before the picture: one step with Shift takes the picture, which Delete removes.
    expect((await selection(page))!.focus).toEqual({ story: 'body', path: [0], offset: 3, before: true });
    await page.keyboard.press('Shift+ArrowRight');
    await page.keyboard.press('Delete');
    expect((await items()).join('')).toBe('abXcd');
    expect(await page.evaluate(() => (window as any).__view.getSelection().focus.offset)).toBe(3);
  });

  test('a selection that goes into a table covers whole cells, as deleting it would', async ({ page }) => {
    await mountEdit(page, await sample(p('Before the table') + tbl(tr(tc('Cell one') + tc('Cell two')) + tr(tc('Cell three') + tc('Cell four')), [3000, 3000]) + p('After')));
    await laidOut(page);
    await clickAt(page, [0], 7);
    await clickAt(page, [1, 0, 0, 0], 4, { shift: true });
    const shown = await page.evaluate(() => (window as any).__view.getSelectedText());
    expect(shown).toBe('the table\nCell one');
    // Delete removes what was shown.
    await page.keyboard.press('Delete');
    const cells = await page.evaluate(() => {
      const table = (window as any).__view.document.body.blocks[1];
      return table.rows[0].cells.map((cell: any) => (window as any).BymOffice.paragraphText(cell.blocks[0]));
    });
    expect(cells).toEqual(['', 'Cell two']);
    expect((await paragraphs(page))[0]).toBe('Before ');
  });
});

test.describe('review: what a cut removes is what a paste brings back', () => {
  const items = (page: Page) =>
    page.evaluate(() =>
      (window as any).__view.document.body.blocks[0].children
        .flatMap((inline: any) => (inline.content || []).map((item: any) => (item.type === 'text' ? item.text : `<${item.type}>`)))
        .join(''),
    );
  const notices = (page: Page) => page.evaluate(() => (window as any).__notices.map((n: any) => n.message) as string[]);
  const withPicture = () => sample(p(r('ab') + picture(120, 60) + r('cd')) + p('next'), { rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', png(40, 20, RED, BLUE)]] });

  test('a cut of a selection that goes into a table is pasted back whole', async ({ page }) => {
    await mountEdit(page, await sample(p('Before') + tbl(tr(tc('A1') + tc('B1')) + tr(tc('A2') + tc('B2')), [3000, 3000]) + p('After')));
    await laidOut(page);
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: 3 }, focus: { story: 'body', path: [1, 1, 0, 0], offset: 1 } }));
    const cut = await clipboardEvent(page, 'cut');
    expect(cut.text).toBe('ore\nA1\tB1\nA2');
    expect(await paragraphs(page)).toEqual(['Bef', '\t', '\tB2', 'After']);
    // The copy the editor kept is what went out of the document: no letter of the cells is missing from it.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [2], offset: 0 }));
    await clipboardEvent(page, 'paste', { text: cut.text, html: cut.html });
    // The cells come back as the rows of a table, after the text that was before them.
    expect(await paragraphs(page)).toEqual(['Bef', '\t', '\tB2', 'ore', 'A1\tB1', 'A2\t', 'After']);
    // One step back and the cut is undone as well.
    await page.evaluate(() => {
      (window as any).__view.undo();
      (window as any).__view.undo();
    });
    expect(await paragraphs(page)).toEqual(['Before', 'A1\tB1', 'A2\tB2', 'After']);
  });

  test('a cut of a picture, alone or among text, takes it out and a paste puts it back; one the copy cannot carry stays, with a word about it', async ({ page }) => {
    await mountEdit(page, await withPicture());
    await laidOut(page);
    // The picture alone, as a press on it selects it.
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: 2, before: true }, focus: { story: 'body', path: [0], offset: 2 } }));
    let cut = await clipboardEvent(page, 'cut', { text: 'what was there' });
    expect(cut.prevented).toBe(true);
    expect(cut.text).toBe('');
    expect(await items(page)).toBe('abcd');
    expect(await notices(page)).toEqual([]);
    await clipboardEvent(page, 'paste', { text: cut.text, html: cut.html });
    expect(await items(page)).toBe('ab<drawing>cd');
    // Text with the picture in it, to the next paragraph.
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: 1 }, focus: { story: 'body', path: [1], offset: 2 } }));
    cut = await clipboardEvent(page, 'cut');
    expect(cut.text).toBe('bcd\nne');
    expect(await items(page)).toBe('axt');
    await clipboardEvent(page, 'paste', { text: cut.text, html: cut.html });
    expect(await items(page)).toBe('ab<drawing>cd');
    expect(await paragraphs(page)).toEqual(['abcd', 'next']);

    // A picture whose relationship is missing cannot be made again elsewhere: it is left in place.
    await mountEdit(page, await sample(p(r('ab') + picture(120, 60, 'rIdGone') + r('cd')) + p('next')));
    await laidOut(page);
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: 1 }, focus: { story: 'body', path: [0], offset: 3 } }));
    cut = await clipboardEvent(page, 'cut', { text: 'what was there' });
    expect(cut.text).toBe('what was there');
    expect(await items(page)).toBe('ab<drawing>cd');
    expect(await notices(page)).toContain('The selection holds something that cannot be cut (a note, an equation, a text box, a content control, a tracked change); it was left in place');
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    // Text beside it is cut as ever.
    await page.evaluate(() => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: 2 }, focus: { story: 'body', path: [0], offset: 4 } }));
    cut = await clipboardEvent(page, 'cut');
    expect(cut.text).toBe('cd');
    expect(await items(page)).toBe('ab<drawing>');
  });

  test('a cut without a clipboard to write removes nothing', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 0);
    for (let i = 0; i < 6; i++) await page.keyboard.press('Shift+ArrowRight');
    const before = await paragraphs(page);
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input')!;
      field.dispatchEvent(new Event('cut', { bubbles: true, cancelable: true }));
    });
    expect(await paragraphs(page)).toEqual(before);
    expect(await notices(page)).toContain('The clipboard cannot be written here; nothing was cut');
  });

  test('a cut of a link, whole or in part, takes the text with its link; a paste brings both back', async ({ page }) => {
    await mountEdit(page, await sample(p(r('ab') + '<w:hyperlink r:id="rIdLink">' + r('link') + '</w:hyperlink>' + r('cd')) + p('next'), { rels: [['rIdLink', 'hyperlink', 'https://example.com/', 'External']] }));
    await laidOut(page);
    const linked = () =>
      page.evaluate(() =>
        (window as any).__view.document.body.blocks[0].children
          .filter((inline: any) => inline.type === 'run' && inline.wrap && inline.wrap.kind === 'hyperlink')
          .map((inline: any) => inline.content.map((item: any) => item.text).join('') + '>' + inline.wrap.target)
          .join('|'),
      );
    for (const [from, to, text, rest] of [[1, 7, 'blinkc', ''], [3, 5, 'in', 'lk>https://example.com/']] as [number, number, string, string][]) {
      await page.evaluate(([a, b]) => (window as any).__view.setSelection({ anchor: { story: 'body', path: [0], offset: a }, focus: { story: 'body', path: [0], offset: b } }), [from, to]);
      const cut = await clipboardEvent(page, 'cut');
      expect(cut.text).toBe(text);
      expect(await linked()).toBe(rest);
      await clipboardEvent(page, 'paste', { text: cut.text, html: cut.html });
      expect((await paragraphs(page))[0]).toBe('ablinkcd');
    }
    // Every piece of the text is a link to the same address still.
    expect(await linked()).toMatch(/^(l|li|lin|link|in|nk|k|i|n)(>https:\/\/example\.com\/\|?)/);
    expect((await linked()).replace(/>https:\/\/example\.com\//g, '').replace(/\|/g, '')).toBe('link');
    expect(await notices(page)).toEqual([]);
  });

  test('a picture at the start or the end of a paragraph: the deletion that comes through the field takes the picture and joins nothing', async ({ page }) => {
    const rels = { rels: [['rIdImg', 'image', 'media/image1.png']] as [string, string, string][], parts: [['word/media/image1.png', png(40, 20, RED, BLUE)]] as [string, Uint8Array][] };
    await mountEdit(page, await sample(p('prev') + p(picture(120, 60) + r('cd')) + p(r('ab') + picture(120, 60)) + p('next'), rels));
    await laidOut(page);
    await clickAt(page, [0], 0);
    // After the picture at the very start: the Backspace takes the line end out of the field.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [1], offset: 0 }));
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.setRangeText('', at - 1, at, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentBackward', bubbles: true }));
    });
    expect(await paragraphs(page)).toEqual(['prev', 'cd', 'ab', 'next']);
    // Before the picture at the very end: a Delete of the same kind.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [2], offset: 2, before: true }));
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.setRangeText('', at, at + 1, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentForward', bubbles: true }));
    });
    expect(await paragraphs(page)).toEqual(['prev', 'cd', 'ab', 'next']);
    const drawings = await page.evaluate(() => JSON.stringify((window as any).__view.document.body.blocks).split('"type":"drawing"').length - 1);
    expect(drawings).toBe(0);
  });

  test('beside a picture, a deletion and a composition that come through the field do what the keys do', async ({ page }) => {
    await mountEdit(page, await withPicture());
    await laidOut(page);
    // The caret after the picture: a Backspace that cannot be cancelled removes the picture, not the letter before it.
    await clickAt(page, [1], 0);
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [0], offset: 2 }));
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.setRangeText('', at - 1, at, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentBackward', bubbles: true }));
    });
    expect(await items(page)).toBe('abcd');
    await page.evaluate(() => (window as any).__view.undo());
    expect(await items(page)).toBe('ab<drawing>cd');
    // Before the picture, a Delete of the same kind.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [0], offset: 2, before: true }));
    await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const at = field.selectionStart;
      field.setRangeText('', at, at + 1, 'end');
      field.dispatchEvent(new InputEvent('input', { inputType: 'deleteContentForward', bubbles: true }));
    });
    expect(await items(page)).toBe('abcd');
    await page.evaluate(() => (window as any).__view.undo());
    // A composition before the picture stays before it when its text is replaced.
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [0], offset: 2, before: true }));
    await compose(page, ['k', 'か'], 'か');
    expect(await items(page)).toBe('abか<drawing>cd');
    expect((await selection(page))!.focus).toEqual({ story: 'body', path: [0], offset: 3, before: true });
  });

  test('a paragraph with a carriage return in its text is not split by a keyboard that composes', async ({ page }) => {
    await mountEdit(page, await sample('<w:p><w:r><w:t>ab&#13;cd</w:t></w:r></w:p>' + p('next')));
    await laidOut(page);
    const count = (await paragraphs(page)).length;
    await clickAt(page, [1], 0);
    await page.evaluate(() => (window as any).__view.setSelection({ story: 'body', path: [0], offset: 5 }));
    await compose(page, ['k'], 'k', true);
    const text = await page.evaluate(() => (window as any).BymOffice.paragraphText((window as any).__view.document.body.blocks[0]));
    expect(text).toBe('ab\rcdk');
    expect((await paragraphs(page)).length).toBe(count);
  });

  test('making the whole of a long document bold does not lay it all out in one go, and ends with the right pages', async ({ page }) => {
    let body = '';
    for (let i = 0; i < 1500; i++) body += p(`${i} ${SENTENCE}${SENTENCE}`);
    await mountEdit(page, await sample(body));
    await laidOut(page);
    const pages = await page.evaluate(() => (window as any).__view.pageCount as number);
    await clickAt(page, [0], 0);
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      view.selectAll();
      // What the view asks of the layout engine: a limit of time, so that the pages come in slices as when the document was opened.
      const engine = view.engine;
      const edited = engine.edited.bind(engine);
      const budgets: unknown[] = [];
      engine.edited = (budget: unknown) => {
        budgets.push(budget);
        return edited(budget);
      };
      const started = performance.now();
      view.format({ bold: true });
      return { ms: performance.now() - started, budgets };
    });
    expect(result.budgets.length).toBe(1);
    expect((result.budgets[0] as { milliseconds: number }).milliseconds).toBeGreaterThan(0);
    expect((result.budgets[0] as { milliseconds: number }).milliseconds).toBeLessThanOrEqual(200);
    expect(result.ms).toBeLessThan(1500);
    await laidOut(page);
    expect(await page.evaluate(() => (window as any).__view.pageCount)).toBeGreaterThanOrEqual(pages);
    expect(await page.evaluate(() => (window as any).__view.document.body.blocks.every((block: any) => block.type !== 'paragraph' || block.children.every((inline: any) => inline.type !== 'run' || (inline.props && inline.props.bold === true))))).toBe(true);
  });
});

test.describe('document editor: focus, stories and places that cannot be edited', () => {
  test('Tab types in the editor; after Escape it leaves it, and that is announced', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await page.locator('#before').focus();
    // Into the toolbar first, then its lists, then the document.
    for (let i = 0; i < 16; i++) {
      if (await input(page).evaluate((element) => element === document.activeElement)) break;
      await page.keyboard.press('Tab');
    }
    await expect(input(page)).toBeFocused();
    await settled(page);
    // The caret is at the start of the text and shown.
    expect(await caretAt(page)).toEqual([[0], 0]);
    expect(await caretRect(page)).not.toBeNull();
    await page.keyboard.press('Tab');
    await expect(input(page)).toBeFocused();
    expect(await page.evaluate(() => (window as any).BymOffice.paragraphText((window as any).__view.document.body.blocks[0]))).toBe('\tFirst paragraph of the document.');
    await page.keyboard.press('Escape');
    expect(await page.locator('#host .bym-doc-live').textContent()).toBe('Press Tab to leave the editor');
    await page.keyboard.press('Tab');
    await expect(input(page)).not.toBeFocused();
    expect(await page.evaluate(() => (window as any).BymOffice.paragraphText((window as any).__view.document.body.blocks[0]))).toBe('\tFirst paragraph of the document.');
    await settled(page);
    // Without the focus there is no caret.
    expect(await caretRect(page)).toBeNull();
  });

  test('the caret says where it is: the character, the word, the line', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 0);
    const live = page.locator('#host .bym-doc-live');
    await page.keyboard.press('ArrowRight');
    await expect(live).toHaveText('i');
    await page.keyboard.press(`${await wordKey(page)}+ArrowRight`);
    await expect(live).toHaveText('paragraph');
    await page.keyboard.press('ArrowDown');
    await expect(live).toHaveText('Second one, a bit longer than the first paragraph is.');
    await page.keyboard.press('Shift+ArrowRight');
    await expect(live).toHaveText(/^\S selected$/);
    const said = await live.textContent();
    // Typing says nothing.
    await page.keyboard.type('abc');
    await expect(live).toHaveText(said as string);
  });

  test('a double click enters the header, typing changes it on every page, a double click in the text leaves it', async ({ page, isMobile }) => {
    test.skip(isMobile, 'double clicks are for the mouse');
    let body = '';
    for (let i = 1; i <= 30; i++) body += p(`Body line ${i}`);
    const bytes = await buildDocx({
      styles: STYLES,
      body: body + docSection('<w:headerReference w:type="default" r:id="h1"/><w:footerReference w:type="default" r:id="f1"/>'),
      headers: { h1: p('Header text') },
      footers: { f1: p(r('Page ') + field(' PAGE ', '1')) },
    });
    await mountEdit(page, bytes);
    await laidOut(page);
    // A single click on the header stays in the text.
    const header = await boxOf(page, [0], 3, 'header:h1');
    await page.mouse.click(header.x, header.y);
    expect((await selection(page))!.focus.story).toBe('body');
    await page.waitForTimeout(600);
    await page.mouse.dblclick(header.x, header.y);
    expect((await selection(page))!.focus.story).toBe('header:h1');
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('Header');
    await expect(page.locator('#host .bym-doc-live')).toHaveText('Editing the header');
    await page.keyboard.type('Title');
    await laidOut(page);
    await settled(page);
    expect(await pageText(page, 1)).toContain('Title text');
    expect(await pageText(page, 2)).toContain('Title text');
    // Now a single click in the header moves the caret there, and one in the text does not leave the header.
    const again = await boxOf(page, [0], 2, 'header:h1');
    await page.mouse.click(again.x, again.y);
    expect((await selection(page))!.focus).toEqual({ story: 'header:h1', path: [0], offset: 2 });
    await page.waitForTimeout(600);
    const text = await boxOf(page, [3], 4);
    await page.mouse.dblclick(text.x, text.y);
    expect((await selection(page))!.focus.story).toBe('body');
    await expect(page.locator('#host .bym-doc-live')).toHaveText('Editing the text');
    // The undo of the header edit is followed too: the word typed after the first letter, then the letter that replaced the selection.
    const command = await commandKey(page);
    await page.keyboard.press(`${command}+z`);
    await page.keyboard.press(`${command}+z`);
    await laidOut(page);
    await settled(page);
    expect(await pageText(page, 1)).toContain('Header text');
  });

  test('the text of a text box is not a place for the caret: a click goes to the text beside it', async ({ page }) => {
    await mountEdit(page, await sample(p(r('Anchor paragraph') + textBoxRun(300, 40, 160, 60, p('Boxed words'))) + p('After the box.')));
    await laidOut(page);
    const inBox = await page.evaluate(() => {
      const view = (window as any).__view;
      const laid = view.pages[0];
      const item = laid.items.find((entry: any) => entry.type === 'text' && entry.text.indexOf('Boxed') === 0);
      const rect = document.querySelector('#host .bym-doc-page[data-page="1"]')!.getBoundingClientRect();
      return { x: rect.left + ((item.x + 20) / laid.width) * rect.width, y: rect.top + ((item.y - 4) / laid.height) * rect.height };
    });
    await page.mouse.click(inBox.x, inBox.y);
    const range = (await selection(page))!;
    expect(range.focus.story).toBe('body');
    expect(range.focus.path.indexOf(-1)).toBe(-1);
    await page.keyboard.type('Q');
    await laidOut(page);
    expect(await pageText(page, 1)).toContain('Boxed words');
    expect((await paragraphs(page)).join('|')).toContain('Q');
    // The API refuses a place in the box without changing anything.
    const changes = await page.evaluate(() => (window as any).__changes);
    const refused = await page.evaluate(() => {
      const view = (window as any).__view;
      view.setSelection({ story: 'body', path: [0, -1, 0, 0, 0], offset: 2 });
      return view.insertText('nope');
    });
    expect(refused).toBe(false);
    expect(await page.evaluate(() => (window as any).__changes)).toBe(changes);
    expect(await page.evaluate(() => (window as any).__notices.map((n: any) => n.message))).toContain('This part of the document cannot be edited');
  });

  test('the caret does not blink when the user asked for less motion', async ({ page }) => {
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    await clickAt(page, [0], 3);
    await settled(page);
    expect(await page.evaluate(() => getComputedStyle(document.querySelector('#host .bym-doc-caret')!).animationName)).toBe('none');
    await page.emulateMedia({ reducedMotion: 'no-preference' });
    expect(await page.evaluate(() => getComputedStyle(document.querySelector('#host .bym-doc-caret')!).animationName)).toMatch(/^bym-doc-blink-/);
  });
});

test.describe('document editor: embedding', () => {
  for (const mode of ['open', 'closed'] as const) {
    test(`works inside a ${mode} shadow root`, async ({ page }) => {
      await mountEdit(page, await threeParagraphs(), { shadow: mode });
      await laidOut(page);
      await clickAt(page, [1], 6);
      expect(await caretAt(page)).toEqual([[1], 6]);
      await page.keyboard.type(' typed');
      expect((await paragraphs(page))[1]).toBe('Second typed one, a bit longer than the first paragraph is.');
      await settled(page);
      expect(await caretRect(page)).not.toBeNull();
      // The toolbar and what drops down from it are in the shadow root as well.
      const state = await page.evaluate(() => {
        const shadow = (window as any).__shadow as ShadowRoot;
        (shadow.querySelector('[data-command="textColor"]') as HTMLElement).click();
        const popover = shadow.querySelector('.bym-popover');
        (shadow.querySelector('[data-command="bold"]') as HTMLElement).click();
        return { popover: !!popover, closed: !shadow.querySelector('.bym-popover'), outside: document.querySelectorAll('body > .bym-popover').length, styles: shadow.querySelectorAll('style').length };
      });
      expect(state.popover).toBe(true);
      expect(state.closed).toBe(true);
      expect(state.outside).toBe(0);
      expect(state.styles).toBeGreaterThanOrEqual(2);
      const copied = await clipboardEvent(page, 'copy');
      expect(copied.text).toBe('');
      await page.keyboard.press('Shift+ArrowLeft');
      expect((await clipboardEvent(page, 'copy')).text).toBe('d');
    });
  }

  test('hit testing and the caret are right when the host is scaled with a CSS transform', async ({ page, isMobile }) => {
    test.skip(isMobile, 'the scaled host does not fit the screen of a phone');
    await mountEdit(page, await threeParagraphs(), { transform: 'scale(0.75)' });
    await laidOut(page);
    const wanted = await boxOf(page, [1], 12);
    await page.mouse.click(wanted.x, wanted.y);
    expect(await caretAt(page)).toEqual([[1], 12]);
    await settled(page);
    const caret = (await caretRect(page))!;
    expect(Math.abs(caret.x - wanted.x)).toBeLessThan(1.5);
    expect(Math.abs(caret.top - wanted.top)).toBeLessThan(1.5);
    await page.keyboard.type('ok ');
    expect((await paragraphs(page))[1]).toBe('Second one, ok a bit longer than the first paragraph is.');
    // The field for typing is at the caret on screen, transform and all.
    await settled(page);
    const field = (await input(page).boundingBox())!;
    const now = (await caretRect(page))!;
    expect(Math.abs(field.x - now.x)).toBeLessThan(4);
    expect(Math.abs(field.y - now.top)).toBeLessThan(4);
  });

  test('destroy leaves no listener, timer or frame behind, also in the middle of work', async ({ page }) => {
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    const empty = await page.evaluate(() => {
      (window as any).__mounted.destroy();
      (window as any).__mounted = null;
      return { ...(window as any).__counts };
    });
    let body = '';
    for (let i = 1; i <= 400; i++) body += p(`${i}. ${SENTENCE}${SENTENCE}`);
    await page.evaluate(([input]) => (window as any).__mountEdit(input, {}), [Array.from(await sample(body))]);
    await clickAt(page, [1], 4);
    await page.keyboard.type('abc');
    await page.keyboard.press('Enter');
    await toolbarButton(page, 'textColor').click();
    // A drag that is still going on, a composition, a popover that is open, layout that has not finished.
    const from = await boxOf(page, [0], 3);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(from.x + 30, from.y + 30);
    const after = await page.evaluate(() => {
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      field.dispatchEvent(new CompositionEvent('compositionstart', { data: '', bubbles: true }));
      const counts = (window as any).__counts;
      const busy = { ...counts };
      (window as any).__mounted.destroy();
      (window as any).__mounted = null;
      return { busy, left: { ...counts }, nodes: document.getElementById('host')!.childNodes.length, styles: document.querySelectorAll('style[data-bymoffice-doc], style[data-bymoffice-doc-editor]').length };
    });
    await page.mouse.up();
    expect(after.busy.listeners).toBeGreaterThan(empty.listeners);
    expect(after.left).toEqual(empty);
    expect(after.nodes).toBe(0);
    expect(after.styles).toBe(0);
    // Nothing of it runs afterwards.
    await page.waitForTimeout(300);
    expect(await page.evaluate(() => ({ ...(window as any).__counts }))).toEqual(empty);
  });

  test('taps: a tap puts the caret, a long press selects a word and brings the bubble', async ({ page, browserName, isMobile }) => {
    test.skip(!isMobile, 'touch is for the phone profiles');
    await mountEdit(page, await threeParagraphs());
    await laidOut(page);
    const at = await boxOf(page, [1], 9);
    await page.touchscreen.tap(at.x, at.y);
    await page.waitForFunction(() => (window as any).__view.getSelection() !== undefined);
    expect(await caretAt(page)).toEqual([[1], 9]);
    await expect(input(page)).toBeFocused();
    await expect(input(page)).toHaveAttribute('inputmode', 'text');
    await page.keyboard.type('x');
    expect((await paragraphs(page))[1]).toBe('Second onxe, a bit longer than the first paragraph is.');
    test.skip(browserName !== 'chromium', 'touch events can only be made up in Chromium');
    const word = await boxOf(page, [0], 9);
    await page.evaluate(
      ([x, y]) => {
        const target = document.elementFromPoint(x, y)!;
        const touch = new Touch({ identifier: 1, target, clientX: x, clientY: y });
        target.dispatchEvent(new TouchEvent('touchstart', { touches: [touch], targetTouches: [touch], changedTouches: [touch], bubbles: true, cancelable: true }));
        (window as any).__touch = { target, touch };
      },
      [word.x, word.y],
    );
    await page.waitForFunction(() => (window as any).__view.getSelectedText() === 'paragraph', undefined, { timeout: 3000 });
    await page.evaluate(() => {
      const { target, touch } = (window as any).__touch;
      target.dispatchEvent(new TouchEvent('touchend', { touches: [], targetTouches: [], changedTouches: [touch], bubbles: true, cancelable: true }));
    });
    await settled(page);
    await expect(page.locator('#host .bym-doc-handle')).toHaveCount(2);
    await expect(page.locator('#host .bym-doc-bubble button')).toHaveText(['Cut', 'Copy', 'Paste', 'Select all']);
    await page.locator('#host .bym-doc-bubble-all').tap();
    expect(await page.evaluate(() => (window as any).__view.getSelectedText().split('\n').length)).toBe(3);
    // Typing ends the selection, and with it the handles and the bubble.
    await page.keyboard.type('Z');
    await settled(page);
    await expect(page.locator('#host .bym-doc-handle')).toHaveCount(0);
    await expect(page.locator('#host .bym-doc-bubble')).toHaveCount(0);
  });
});

test.describe('document editor: speed', () => {
  /** Types `count` letters at the caret, each as the browser would, and brings the screen up to date after each. Returns the time each took, ms. */
  function typeTimed(page: Page, count: number): Promise<number[]> {
    return page.evaluate((n) => {
      const view = (window as any).__view;
      const field = document.querySelector('#host .bym-doc-input') as HTMLTextAreaElement;
      const times: number[] = [];
      const letters = 'etaoin shrdlu ';
      for (let i = 0; i < n; i++) {
        const started = performance.now();
        field.dispatchEvent(new InputEvent('beforeinput', { inputType: 'insertText', data: letters.charAt(i % letters.length), bubbles: true, cancelable: true }));
        // Model, layout and paint: what a frame after the key would do.
        view.update();
        times.push(performance.now() - started);
      }
      return times;
    }, count);
  }

  const summary = (times: number[]) => {
    const sorted = times.slice().sort((a, b) => a - b);
    return { median: sorted[Math.floor(sorted.length / 2)], p95: sorted[Math.floor(sorted.length * 0.95)], worst: sorted[sorted.length - 1] };
  };

  test('typing 200 characters into a document of 50 pages', async ({ page }, info) => {
    let body = '';
    for (let i = 1; i <= 290; i++) body += p(`${i}. ${SENTENCE}${SENTENCE}`);
    await mountEdit(page, await sample(body));
    await laidOut(page);
    const pages = await page.evaluate(() => (window as any).__view.pageCount);
    expect(pages).toBeGreaterThanOrEqual(45);
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.focus();
      view.setSelection({ story: 'body', path: [140], offset: 20 });
    });
    await settled(page);
    const before = (await paragraphs(page))[140];
    const result = summary(await typeTimed(page, 200));
    info.annotations.push({ type: 'typing-50-pages', description: JSON.stringify(result) });
    console.log(`[doc-edit speed] ${info.project.name} 50 pages (${pages}): median ${result.median.toFixed(2)} ms, p95 ${result.p95.toFixed(2)} ms, worst ${result.worst.toFixed(2)} ms`);
    expect((await paragraphs(page))[140].length).toBe(before.length + 200);
    // Generous: the machines that run this differ. A frame is 16 ms.
    expect(result.median).toBeLessThan(40);
    expect(result.worst).toBeLessThan(600);
    await laidOut(page);
    // What is on screen is what a new layout of the edited document gives.
    const same = await page.evaluate(() => {
      const api = (window as any).BymOffice;
      const view = (window as any).__view;
      const layout = new api.DocLayout(view.document, view.measurer, { lineMarks: true });
      layout.layout();
      return JSON.stringify(layout.pages) === JSON.stringify(view.pages);
    });
    expect(same).toBe(true);
  });

  test('typing in the middle of a document of 1000 pages', async ({ page }, info) => {
    test.setTimeout(240_000);
    let body = '';
    for (let i = 1; i <= 5800; i++) body += p(`${i}. ${SENTENCE}${SENTENCE}`);
    await mountEdit(page, await sample(body));
    await laidOut(page);
    const pages = await page.evaluate(() => (window as any).__view.pageCount);
    expect(pages).toBeGreaterThanOrEqual(950);
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.focus();
      view.setSelection({ story: 'body', path: [2900], offset: 20 });
    });
    await settled(page);
    expect(await page.evaluate(() => (window as any).__view.currentPage)).toBeGreaterThan(400);
    const result = summary(await typeTimed(page, 100));
    info.annotations.push({ type: 'typing-1000-pages', description: JSON.stringify(result) });
    console.log(`[doc-edit speed] ${info.project.name} 1000 pages (${pages}): median ${result.median.toFixed(2)} ms, p95 ${result.p95.toFixed(2)} ms, worst ${result.worst.toFixed(2)} ms`);
    expect(result.median).toBeLessThan(80);
    expect(result.worst).toBeLessThan(1500);
    // Enter moves every block after it; the page of the caret is there at once all the same.
    const enter = await page.evaluate(() => {
      const view = (window as any).__view;
      const started = performance.now();
      view.insertText('\n');
      view.update();
      const time = performance.now() - started;
      return { time, caret: !!view.lines.caret(view.getSelection().focus), done: view.layoutDone };
    });
    console.log(`[doc-edit speed] ${info.project.name} 1000 pages: Enter ${enter.time.toFixed(2)} ms`);
    expect(enter.caret).toBe(true);
    expect(enter.time).toBeLessThan(1500);
    await laidOut(page);
    expect(await page.evaluate(() => (window as any).__view.pageCount)).toBeGreaterThanOrEqual(pages);
  });
});

test.describe('saving an edited docx', () => {
  test('mount() saves the typed text into the file, and the file opens with it', async ({ page }) => {
    await page.goto('/e2e/pages/doc-edit.html');
    const result = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      // A docx written by the library itself from a one-paragraph RTF.
      const source = await api.open(new TextEncoder().encode('{\\rtf1\\ansi First paragraph\\par}'));
      const docx = await source.saveAsDocx();
      const host = document.createElement('div');
      host.style.height = '400px';
      document.body.appendChild(host);
      const mounted = await api.mount(host, { file: docx });
      mounted.view.setSelection({ story: 'body', path: [0], offset: 0 });
      mounted.view.insertText('Hello ');
      const modified = mounted.view.isModified;
      const bytes = await mounted.save();
      const afterSave = mounted.view.isModified;
      mounted.destroy();
      host.remove();
      const again = await api.open(bytes);
      const doc = await again.document();
      return { modified, afterSave, text: api.paragraphText(doc.body.blocks[0]) };
    });
    expect(result).toEqual({ modified: true, afterSave: false, text: 'Hello First paragraph' });
  });

  test('review: a change made while the file is being written is not taken as saved', async ({ page }) => {
    await page.goto('/e2e/pages/doc-edit.html');
    const result = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const source = await api.open(new TextEncoder().encode('{\\rtf1\\ansi First paragraph\\par}'));
      const host = document.createElement('div');
      host.style.height = '400px';
      document.body.appendChild(host);
      const mounted = await api.mount(host, { file: await source.saveAsDocx() });
      mounted.view.setSelection({ story: 'body', path: [0], offset: 0 });
      mounted.view.insertText('Hello ');
      const writing = mounted.save();
      mounted.view.insertText('again ');
      await writing;
      const duringSave = mounted.view.isModified;
      await mounted.save();
      const afterSave = mounted.view.isModified;
      mounted.destroy();
      host.remove();
      return { duringSave, afterSave };
    });
    expect(result).toEqual({ duringSave: true, afterSave: false });
  });
});
