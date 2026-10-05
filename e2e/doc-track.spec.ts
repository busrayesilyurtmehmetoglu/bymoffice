import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { docx, DOC_STYLES } from './fixtures/docx';
import { ZipArchive } from '../src/zip';
import { expect, test } from './helpers';

// Documents that track their changes, and documents that are protected: what
// is typed and deleted becomes a tracked change in the saved file and shows as
// one on the page; what cannot be tracked is refused; a protected document
// stays read-only.

interface MountOptions {
  readOnly?: boolean;
  locale?: string;
  trackChanges?: boolean;
  author?: string;
  now?: string;
}

const STYLES = DOC_STYLES + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>';
const TRACK = '<w:trackRevisions/>';
const NOW = '2026-10-05T09:08:07.654Z';
const DATE = '2026-10-05T09:08:07Z';
const AUTHOR = 'Ayşe <b>&"Yılmaz"';
/** The name as an attribute value holds it: nothing of it is markup. */
const AUTHOR_XML = 'Ayşe &lt;b>&amp;&quot;Yılmaz&quot;';

/** A document with the given settings; null for one without a settings part. */
const sample = (body: string, settings: string | null = TRACK) => docx(body, settings === null ? { styles: STYLES } : { styles: STYLES, settings });
const protect = (edit: string, enforcement = '1') => `<w:documentProtection w:edit="${edit}" w:enforcement="${enforcement}" w:cryptProviderType="rsaAES" w:cryptAlgorithmSid="14" w:cryptSpinCount="100000" w:hash="q83vEjRWeJA=" w:salt="AAECAwQFBgc="/>`;

async function mountEdit(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { author: AUTHOR, now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

/** Lets the view bring the screen up to date: its updates run in animation frames. */
const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

const paragraphs = (page: Page) => page.evaluate(() => ((window as any).BymOffice.documentText((window as any).__view.document) as string).split('\n'));

const notices = (page: Page): Promise<{ message: string; kind: string }[]> => page.evaluate(() => (window as any).__notices.slice());

/** Puts the caret (or a selection) into the document and the focus into the editor. */
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

/** The text items of the first page as the layout made them: text, struck, underlined, part of the text. */
const shown = (page: Page): Promise<[string, boolean, boolean, boolean][]> =>
  page.evaluate(() =>
    ((window as any).__view.pages[0].items as any[])
      .filter((item) => item.type === 'text' && item.text !== '')
      .map((item) => [item.text, item.strike === true, item.underline !== undefined, item.source !== undefined] as [string, boolean, boolean, boolean]),
  );

/** The document part of the file the editor saves. */
async function savedXml(page: Page): Promise<string> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  return new TextDecoder().decode(await ZipArchive.read(new Uint8Array(bytes)).read('word/document.xml'));
}

/** Saves, and opens what was saved in a new editor. */
async function reopen(page: Page, options: MountOptions = {}): Promise<void> {
  const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [bytes, { author: AUTHOR, now: NOW, ...options }] as [number[], MountOptions]);
  await page.waitForFunction(() => (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
}

/** The tracked changes of the body as the model holds them: [kind, author, date, text] for every run in one. */
const revisions = (page: Page): Promise<[string, string, string, string][]> =>
  page.evaluate(() => {
    const out: [string, string, string, string][] = [];
    for (const block of (window as any).__view.document.body.blocks) {
      if (block.type !== 'paragraph') continue;
      for (const inline of block.children) {
        if (inline.type !== 'run' || !inline.wrap) continue;
        const text = inline.content.map((item: any) => (item.type === 'text' || item.type === 'deletedText' ? item.text : `<${item.type}>`)).join('');
        for (let w = inline.wrap; w; w = w.parent) if (w.kind === 'ins' || w.kind === 'del') out.push([w.kind, w.author, w.date, text]);
      }
    }
    return out;
  });

const commandKey = (page: Page) => page.evaluate(() => (/Mac|iPhone|iPad|iPod/.test(navigator.platform) ? 'Meta' : 'Control'));
const toolbarButton = (page: Page, command: string) => page.locator(`#host .bym-doc-toolbar [data-command="${command}"]`);

test.describe('tracked changes', () => {
  test('what is typed is a tracked insertion: in the model, on the page and in the saved file, with the author and the date', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    expect(await page.evaluate(() => [(window as any).__mounted.trackChanges === true, (window as any).__view.trackingChanges, (window as any).__mounted.protection])).toEqual([true, true, undefined]);
    await select(page, [0], 5);
    await page.keyboard.type(' new');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First new paragraph.', 'Second one.']);
    expect(await revisions(page)).toEqual([['ins', AUTHOR, DATE, ' new']]);
    // Underlined on the page, and part of the text.
    expect(await shown(page)).toContainEqual([' new', false, true, true]);

    const xml = await savedXml(page);
    expect(xml).toContain(`<w:ins w:id="1" w:author="${AUTHOR_XML}" w:date="${DATE}"><w:r><w:t xml:space="preserve"> new</w:t></w:r></w:ins>`);
    expect(xml.split('<w:ins ').length).toBe(2);
    expect(xml).not.toContain('<w:del ');

    await reopen(page);
    expect(await paragraphs(page)).toEqual(['First new paragraph.', 'Second one.']);
    expect(await revisions(page)).toEqual([['ins', AUTHOR, DATE, ' new']]);
    expect(await page.evaluate(() => (window as any).__view.document.settings.trackRevisions)).toBe(true);
  });

  test('what is deleted stays on the page struck through, and is a tracked deletion in the saved file', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    await select(page, [0], 5);
    for (let i = 0; i < 5; i++) await page.keyboard.press('Backspace');
    await settled(page);
    expect(await paragraphs(page)).toEqual([' paragraph.', 'Second one.']);
    expect(await revisions(page)).toEqual([['del', AUTHOR, DATE, 'First']]);
    const items = await shown(page);
    // Struck through, shown, and no part of the text any more.
    expect(items).toContainEqual(['First', true, false, false]);
    expect(await page.evaluate(() => (window as any).__view.getSelection().focus.offset)).toBe(0);
    // The page as text, for assistive technology, is the text that is there.
    const spoken = await page.evaluate(() => document.querySelector('#host .bym-doc-page[data-page="1"] .bym-doc-text')!.textContent);
    expect(spoken).toContain('paragraph.');
    expect(spoken).not.toContain('First');

    // Delete goes on from there, and a selection typed over is deleted and replaced.
    await page.keyboard.press('Delete');
    await select(page, [1], 0, 6);
    await page.keyboard.type('X');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['paragraph.', 'X one.']);

    const xml = await savedXml(page);
    expect(xml).toContain('<w:delText xml:space="preserve">First </w:delText>');
    expect(xml).toContain('<w:delText>Second</w:delText>');
    expect(xml.split('<w:del ').length).toBe(3);
    expect(xml.split('<w:ins ').length).toBe(2);
    // Deleted text is never ordinary text in the file.
    expect(xml).not.toMatch(/<w:del [^>]*>(?:(?!<\/w:del>).)*<w:t[ >]/);

    await reopen(page);
    expect(await paragraphs(page)).toEqual(['paragraph.', 'X one.']);
    expect((await revisions(page)).map((entry) => [entry[0], entry[3]])).toEqual([
      ['del', 'First '],
      ['del', 'Second'],
      ['ins', 'X'],
    ]);
    expect(await shown(page)).toContainEqual(['Second', true, false, false]);
  });

  test('deleting what the same author typed leaves no trace; undo and redo keep the tracked changes', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.')));
    const before = await savedXml(page);
    await select(page, [0], 5);
    await page.keyboard.type('abc');
    for (let i = 0; i < 3; i++) await page.keyboard.press('Backspace');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First paragraph.']);
    expect(await revisions(page)).toEqual([]);
    const xml = await savedXml(page);
    expect(xml).not.toContain('<w:ins ');
    expect(xml).not.toContain('<w:del ');

    // Everything undone: the file is the one that was opened.
    await page.evaluate(() => {
      const view = (window as any).__view;
      while (view.canUndo) view.undo();
    });
    expect(await savedXml(page)).toBe(before);
    await page.evaluate(() => (window as any).__view.redo());
    expect((await revisions(page)).length).toBeGreaterThan(0);
  });

  test('Enter and Backspace over a paragraph mark are tracked: the paragraphs stay two until the change is accepted', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    await select(page, [1], 0);
    await page.keyboard.press('Backspace');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Second one.']);
    // The caret went to the end of the paragraph before.
    expect(await page.evaluate(() => (window as any).__view.getSelection().focus)).toMatchObject({ path: [0], offset: 16 });
    await page.keyboard.press('Enter');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['First paragraph.', '', 'Second one.']);
    const xml = await savedXml(page);
    expect(xml).toContain(`<w:rPr><w:del w:id="1" w:author="${AUTHOR_XML}" w:date="${DATE}"/></w:rPr>`);
    // The new mark is the one of the first half; the deleted one stays the end of what it ended.
    expect(xml.indexOf('<w:rPr><w:ins w:id="2" ')).toBeGreaterThan(0);
    expect(xml.indexOf('<w:rPr><w:ins w:id="2" ')).toBeLessThan(xml.indexOf('<w:rPr><w:del w:id="1" '));
  });

  test('a change that cannot be tracked is refused with a notice, and the toolbar does not offer what it cannot track', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    const before = await savedXml(page);
    await select(page, [0], 0, 5);
    await page.keyboard.press(`${await commandKey(page)}+b`);
    await settled(page);
    expect(await notices(page)).toEqual([{ message: 'Changes are tracked in this document, and this kind of change cannot be tracked yet; nothing was changed', kind: 'info' }]);
    expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
    expect(await savedXml(page)).toBe(before);
    for (const command of ['alignCenter', 'bulletList', 'insertTable', 'insertPicture', 'insertLink']) await expect(toolbarButton(page, command)).toBeDisabled();
    await expect(toolbarButton(page, 'bold')).toBeEnabled();
    // From the program too.
    expect(await page.evaluate(() => (window as any).__view.insertTable(2, 2))).toBe(false);
    expect((await notices(page)).length).toBe(2);
    expect(await paragraphs(page)).toEqual(['First paragraph.', 'Second one.']);
  });

  test('in Turkish the refusal is said in Turkish', async ({ page }) => {
    await mountEdit(page, await sample(p('Birinci paragraf.')), { locale: 'tr' });
    await select(page, [0], 0, 7);
    await page.keyboard.press(`${await commandKey(page)}+b`);
    await settled(page);
    expect((await notices(page))[0].message).toBe('Bu belgede değişiklikler izleniyor ve bu tür değişiklik henüz izlenemiyor; hiçbir şey değiştirilmedi');
  });

  test('a copy is pasted as tracked text, and the user is told that its formatting stayed behind', async ({ page }) => {
    await mountEdit(page, await sample(p(r('Bold', '<w:b/>') + r(' and plain.')) + p('Second one.')));
    await select(page, [0], 0, 4);
    const clipboard = (kind: 'copy' | 'paste', give: { text?: string; html?: string } = {}) =>
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
    const copied = await clipboard('copy');
    expect(copied.text).toBe('Bold');
    await select(page, [1], 0);
    await clipboard('paste', copied);
    await settled(page);
    expect(await paragraphs(page)).toEqual(['Bold and plain.', 'BoldSecond one.']);
    expect(await revisions(page)).toEqual([['ins', AUTHOR, DATE, 'Bold']]);
    expect(await notices(page)).toEqual([{ message: 'Changes are tracked in this document: the copy was pasted as text, without its formatting and pictures', kind: 'info' }]);
  });

  test('a cut puts the text on the clipboard and leaves it in the document as a tracked deletion', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.') + p('Second one.')));
    await select(page, [0], 0, 6);
    const cut = await page.evaluate(() => {
      const target = document.querySelector('#host .bym-doc-input')!;
      const data = new DataTransfer();
      let event: Event | undefined;
      try {
        event = new ClipboardEvent('cut', { clipboardData: data, bubbles: true, cancelable: true });
      } catch {
        event = undefined;
      }
      if (!event || (event as ClipboardEvent).clipboardData !== data) {
        event = new Event('cut', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
      }
      target.dispatchEvent(event);
      return data.getData('text/plain');
    });
    await settled(page);
    expect(cut).toBe('First ');
    expect(await paragraphs(page)).toEqual(['paragraph.', 'Second one.']);
    expect(await revisions(page)).toEqual([['del', AUTHOR, DATE, 'First ']]);
    expect(await notices(page)).toEqual([]);
    // What is copied now does not hold the deleted text.
    await page.evaluate(() => (window as any).__view.selectAll());
    expect(await page.evaluate(() => (window as any).__view.getSelectedText())).toBe('paragraph.\nSecond one.');
  });

  test('a cut over what the same author inserted takes out the typed text only: the note and the field stay, as deletions', async ({ page }) => {
    const field =
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>today</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const own = `<w:ins w:id="7" w:author="Ada" w:date="2020-01-01T00:00:00Z">${r('x')}<w:r><w:footnoteReference w:id="2"/></w:r>${field}${r('y')}</w:ins>`;
    const bytes = await docx(p(r('a') + own + r('b')), {
      styles: STYLES,
      settings: TRACK,
      footnotes: '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>the note</w:t></w:r></w:p></w:footnote>',
    });
    await mountEdit(page, bytes, { author: 'Ada' });
    const before = await savedXml(page);
    await page.evaluate(() => (window as any).__view.selectAll());
    await page.evaluate(() => (window as any).__view.focus());
    const cut = await page.evaluate(() => {
      const target = document.querySelector('#host .bym-doc-input')!;
      const data = new DataTransfer();
      let event: Event | undefined;
      try {
        event = new ClipboardEvent('cut', { clipboardData: data, bubbles: true, cancelable: true });
      } catch {
        event = undefined;
      }
      if (!event || (event as ClipboardEvent).clipboardData !== data) {
        event = new Event('cut', { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'clipboardData', { value: data });
      }
      target.dispatchEvent(event);
      return data.getData('text/plain');
    });
    await settled(page);
    expect(cut).toContain('today');
    expect((await paragraphs(page))[0]).toBe('');
    expect(await notices(page)).toEqual([]);
    expect(await page.evaluate(() => (window as any).__view.document.footnotes.items.has(2))).toBe(true);
    const xml = await savedXml(page);
    expect(xml).toContain('<w:footnoteReference w:id="2"/>');
    expect(xml).toContain('<w:delInstrText xml:space="preserve"> DATE </w:delInstrText>');
    expect(xml).toContain('<w:delText>today</w:delText>');
    expect(xml).not.toContain('>x<');
    expect(xml).not.toContain('>y<');
    // The note is still in its part.
    const saved = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    expect(new TextDecoder().decode(await ZipArchive.read(new Uint8Array(saved)).read('word/footnotes.xml'))).toContain('the note');
    await page.evaluate(() => (window as any).__view.undo());
    expect(await savedXml(page)).toBe(before);
  });

  test('a document that does not track is edited without tracked changes, unless the application asks for them', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.'), null));
    expect(await page.evaluate(() => [(window as any).__mounted.trackChanges === true, (window as any).__view.trackingChanges])).toEqual([false, false]);
    await select(page, [0], 5);
    await page.keyboard.type('X');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('Backspace');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['Firs paragraph.']);
    expect(await revisions(page)).toEqual([]);
    expect(await savedXml(page)).not.toMatch(/<w:(ins|del) /);

    await mountEdit(page, await sample(p('First paragraph.'), null), { trackChanges: true });
    expect(await page.evaluate(() => (window as any).__mounted.trackChanges)).toBe(true);
    await select(page, [0], 5);
    await page.keyboard.type('X');
    await settled(page);
    expect(await revisions(page)).toEqual([['ins', AUTHOR, DATE, 'X']]);
  });

  test('without a name the changes are those of "Author"', async ({ page }) => {
    await page.goto('/e2e/pages/doc-edit.html');
    await expect(page.locator('#state')).toHaveText('ready');
    const bytes = await sample(p('First paragraph.'));
    await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { now: NOW }] as [number[], MountOptions]);
    await page.waitForFunction(() => (window as any).__view.layoutDone === true);
    await select(page, [0], 0);
    await page.keyboard.type('X');
    await settled(page);
    expect(await revisions(page)).toEqual([['ins', 'Author', DATE, 'X']]);
  });
});

test.describe('protected documents', () => {
  for (const edit of ['readOnly', 'comments', 'forms']) {
    test(`a document protected for "${edit}" is read-only and stays so; the reason is said once`, async ({ page }) => {
      await mountEdit(page, await sample(p('First paragraph.') + p('Second one.'), protect(edit)));
      expect(await page.evaluate(() => [(window as any).__mounted.protection, (window as any).__view.readOnly, (window as any).__view.protection])).toEqual([{ edit, exceptions: false }, true, edit]);
      await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
      const said = await notices(page);
      expect(said.length).toBe(1);
      expect(said[0].kind).toBe('info');
      expect(said[0].message).toContain('protected');

      // Nothing turns editing on, neither the editor nor the view, and nothing more is said.
      await page.evaluate(() => {
        (window as any).__mounted.editor.setReadOnly(false);
        (window as any).__view.setReadOnly(false);
      });
      await settled(page);
      expect(await page.evaluate(() => (window as any).__view.readOnly)).toBe(true);
      await expect(page.locator('#host .bym-doc-toolbar')).toBeHidden();
      expect((await notices(page)).length).toBe(1);

      // Typing changes nothing, and the file is saved as it was opened.
      await page.evaluate(() => {
        const view = (window as any).__view;
        view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 0 } });
        view.focus();
      });
      await page.keyboard.type('X');
      await page.keyboard.press('Delete');
      expect(await page.evaluate(() => (window as any).__view.insertText('X'))).toBe(false);
      await settled(page);
      expect(await paragraphs(page)).toEqual(['First paragraph.', 'Second one.']);
      expect(await page.evaluate(() => (window as any).__view.isModified)).toBe(false);
      const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
      const settings = new TextDecoder().decode(await ZipArchive.read(new Uint8Array(bytes)).read('word/settings.xml'));
      expect(settings).toContain(protect(edit));
    });
  }

  test('a protection that is not enforced restricts nothing', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.'), protect('readOnly', '0')));
    expect(await page.evaluate(() => [(window as any).__mounted.protection, (window as any).__view.readOnly, (window as any).__mounted.trackChanges === true])).toEqual([undefined, false, false]);
    expect(await notices(page)).toEqual([]);
    await select(page, [0], 0);
    await page.keyboard.type('X');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['XFirst paragraph.']);
  });

  test('ranges exempt from the protection are reported; the document is read-only all the same', async ({ page }) => {
    const body = `<w:p><w:permStart w:id="1" w:edGrp="everyone"/>${r('free')}<w:permEnd w:id="1"/></w:p>`;
    await mountEdit(page, await sample(body, protect('readOnly')));
    expect(await page.evaluate(() => [(window as any).__mounted.protection, (window as any).__view.readOnly])).toEqual([{ edit: 'readOnly', exceptions: true }, true]);
  });

  test('a document protected for tracked changes is edited, and every change is tracked whatever the application asks', async ({ page }) => {
    await mountEdit(page, await sample(p('First paragraph.'), protect('trackedChanges')), { trackChanges: false });
    expect(await page.evaluate(() => [(window as any).__mounted.protection, (window as any).__view.readOnly, (window as any).__mounted.trackChanges === true])).toEqual([{ edit: 'trackedChanges', exceptions: false }, false, true]);
    expect(await notices(page)).toEqual([]);
    await select(page, [0], 5);
    await page.keyboard.type('X');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('Backspace');
    await settled(page);
    expect(await paragraphs(page)).toEqual(['FirstXparagraph.']);
    expect((await revisions(page)).map((entry) => [entry[0], entry[3]])).toEqual([
      ['ins', 'X'],
      ['del', ' '],
    ]);
    const bytes = await page.evaluate(async () => Array.from((await (window as any).__mounted.save()) as Uint8Array));
    const zip = ZipArchive.read(new Uint8Array(bytes));
    expect(new TextDecoder().decode(await zip.read('word/settings.xml'))).toContain(protect('trackedChanges'));
    expect(new TextDecoder().decode(await zip.read('word/document.xml'))).toContain('<w:delText xml:space="preserve"> </w:delText>');
  });
});

test.describe('the WebView bridge', () => {
  const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');

  async function bridge(page: Page) {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto('/dist/bymoffice.webview.html');
    const messages = () => page.evaluate(() => ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)));
    const send = (message: object) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
    const waitFor = async (type: string, count = 1) => {
      let found: any[] = [];
      await expect.poll(async () => (found = (await messages()).filter((message) => message.type === type)).length, { message: `${count} x ${type}` }).toBeGreaterThanOrEqual(count);
      return found[count - 1];
    };
    await waitFor('ready');
    return { send, waitFor, messages };
  }

  test('loaded says that a document is protected, and it stays read-only', async ({ page }) => {
    const app = await bridge(page);
    await app.send({ type: 'load', id: 'a', data: base64(await sample(p('First paragraph.'), protect('forms'))), readOnly: false, toolbar: true });
    expect(await app.waitFor('loaded')).toEqual({ type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, readOnly: true, protection: 'forms' });
    await app.send({ type: 'setReadOnly', id: 'b', readOnly: false });
    expect(await app.waitFor('done')).toEqual({ type: 'done', id: 'b', command: 'setReadOnly', readOnly: true });
    await expect(page.locator('.bym-doc-toolbar')).toBeHidden();
  });

  test('loaded says that changes are tracked; the author of the load names them in the saved file', async ({ page }) => {
    const app = await bridge(page);
    await app.send({ type: 'load', id: 'a', data: base64(await sample(p('First paragraph.'))), author: 'Köprü Yazarı', toolbar: true });
    expect(await app.waitFor('loaded')).toEqual({ type: 'loaded', id: 'a', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, trackChanges: true });
    const input = page.locator('.bym-doc-input');
    await input.focus();
    await expect(input).toBeFocused();
    await page.keyboard.type('Z');
    await app.waitFor('changed');
    await app.send({ type: 'save', id: 's' });
    const saved = await app.waitFor('saved');
    const xml = new TextDecoder().decode(await ZipArchive.read(new Uint8Array(Buffer.from(saved.data, 'base64'))).read('word/document.xml'));
    expect(xml).toMatch(/<w:ins w:id="1" w:author="Köprü Yazarı" w:date="\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ"><w:r><w:t>Z<\/w:t><\/w:r><\/w:ins>/);
  });
});
