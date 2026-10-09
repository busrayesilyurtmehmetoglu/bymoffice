import type { APIRequestContext, Locator, Page } from '@playwright/test';
import { createBlank, documentText, getCell, open } from '../src/index';
import { cat, number, workbookStream, xlsFile } from '../test/xls-fixtures';
import { expect, test } from './helpers';

// MountOptions.saveTo: a file that is opened from an address is edited and
// saved back to it by the editor itself: the Save button of the toolbar,
// Ctrl+S / Cmd+S and saveNow(). The server of these tests keeps what it is
// sent at /e2e/saved/<name> (in memory) and gives it back with an ETag.

const PAGE = '/e2e/pages/mount.html';
const SAVE = '#editor .bym-toolbar [data-command="save"]';
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const DOCX_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** An address of the server that no other test uses. */
function address(extension: string): string {
  const info = test.info();
  return `/e2e/saved/${`${info.project.name}-${info.testId}`.replace(/[^A-Za-z0-9_-]/g, '_')}.${extension}`;
}

async function seed(request: APIRequestContext, url: string, bytes: Uint8Array): Promise<void> {
  const answer = await request.put(url, { data: Buffer.from(bytes), headers: { 'content-type': 'application/octet-stream' } });
  expect(answer.status(), `seeding ${url}`).toBe(204);
}

/** What the server holds at an address: the file's first cell or text, and what the request said it was. */
async function onServer(request: APIRequestContext, url: string): Promise<{ format: string; text: string; type: string | undefined; said: string | undefined; from: string | undefined }> {
  const answer = await request.get(url);
  expect(answer.status(), `reading ${url}`).toBe(200);
  const file = await open(new Uint8Array(await answer.body()));
  let text: string;
  if (file.kind === 'document') text = documentText(await file.document!());
  else {
    const workbook = await file.workbook();
    const cells = [getCell(workbook.sheets[0], 0, 0), getCell(workbook.sheets[0], 1, 0)];
    text = cells.map((cell) => (cell ? String(cell.value) : '')).join('|');
  }
  const headers = answer.headers();
  return { format: file.format, text, type: headers['x-saved-content-type'], said: headers['x-saved-format'], from: headers['x-saved-converted-from'] };
}

interface Spec {
  file: string | number[];
  /** What is given as `saveTo` when it can be sent to the page as it is. */
  saveTo?: unknown;
  /** A function as `saveTo`: one that keeps what it is given, or one that rejects. */
  saver?: 'keeps' | 'rejects';
  locale?: string;
  readOnly?: boolean;
  toolbar?: boolean;
  convert?: boolean;
  warnOnLeave?: boolean;
  onSave?: boolean;
}

/** Mounts in the page; resolves to "ok" or to what mount() rejected with. */
function mount(page: Page, spec: Spec): Promise<string> {
  return page.evaluate(async (given) => {
    const w = window as any;
    w.__states = [];
    w.__notices = [];
    w.__errors = [];
    w.__onSave = [];
    w.__kept = [];
    w.__keys = [];
    if (!w.__watchingKeys) {
      w.__watchingKeys = true;
      // What is left of Ctrl+S once the editor has seen it: whether the browser still gets it. Looked at after
      // the event has gone all its way, and heard on its way down: a dialog inside the editor lets no key out.
      window.addEventListener(
        'keydown',
        (event) => {
          if (event.key === 's' || event.key === 'S') setTimeout(() => w.__keys.push(event.defaultPrevented), 0);
        },
        true,
      );
    }
    const options: any = {
      file: typeof given.file === 'string' ? given.file : new Uint8Array(given.file),
      locale: given.locale || 'en',
      onSaveState: (state: unknown) => w.__states.push(state),
      onNotice: (notice: { kind: string; message: string }) => w.__notices.push(`${notice.kind}: ${notice.message}`),
      onError: (error: { code?: string; status?: number; name?: string }) => w.__errors.push(`${error.name} ${error.code} ${error.status}`),
    };
    if ('saveTo' in given) options.saveTo = given.saveTo;
    if (given.saver === 'keeps') options.saveTo = (bytes: Uint8Array, info: unknown) => void w.__kept.push({ bytes: Array.from(bytes), info });
    if (given.saver === 'rejects') options.saveTo = () => Promise.reject(new Error('The disk is <b>full</b>'));
    for (const name of ['readOnly', 'toolbar', 'convert', 'warnOnLeave'] as const) if (name in given) options[name] = given[name];
    if (given.onSave) options.onSave = (bytes: Uint8Array) => w.__onSave.push(`${bytes.length} bytes, dirty ${w.__mounted.hasUnsavedChanges}`);
    try {
      w.__mounted = await w.BymOffice.mount(document.getElementById('editor'), options);
      return 'ok';
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  }, spec);
}

const lastState = (page: Page) => page.evaluate(() => (window as any).__states[(window as any).__states.length - 1]);
const notices = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__notices);
const unsaved = (page: Page): Promise<boolean> => page.evaluate(() => (window as any).__mounted.hasUnsavedChanges);
/** Whether the page would ask before it is left: what the browser asks its `beforeunload` listeners. */
const asksBeforeLeaving = (page: Page): Promise<boolean> =>
  page.evaluate(() => {
    const event = new Event('beforeunload', { cancelable: true });
    window.dispatchEvent(event);
    return event.defaultPrevented;
  });

/** A press with whatever the project has: a finger where there is a touch screen, the mouse elsewhere. */
const press = (target: Locator, hasTouch: boolean | undefined) => (hasTouch ? target.tap() : target.click());

/** Types into the grid with real keys, from the cell that is selected. */
async function typeInGrid(page: Page, text: string, enter = true): Promise<void> {
  await page.evaluate(() => (window as any).__mounted.view.focus());
  await expect(page.locator('#editor .bym-input')).toBeFocused();
  await page.keyboard.type(text);
  if (enter) await page.keyboard.press('Enter');
}

/** Types at the start of the document with real keys. */
async function typeInDocument(page: Page, text: string): Promise<void> {
  await page.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.focus();
    view.setSelection({ story: 'body', path: [0], offset: 0 });
  });
  await page.keyboard.type(text);
}

test.describe('saveTo: a file is saved to where it came from', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(PAGE);
  });

  test('workbook: Save is dimmed until something is typed; pressing it sends the file back, and the next save follows the new ETag', async ({ page, request, hasTouch }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    const save = page.locator(SAVE);
    await expect(save).toBeVisible();
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-label', /^Save \((Ctrl\+S|⌘S)\)$/);
    await expect(save).toHaveAttribute('aria-busy', 'false');
    // The first control of the toolbar, before Undo.
    expect(await page.locator('#editor .bym-toolbar button').first().getAttribute('data-command')).toBe('save');
    expect(await unsaved(page)).toBe(false);
    expect(await asksBeforeLeaving(page)).toBe(false);

    await typeInGrid(page, '42');
    await expect(save).toBeEnabled();
    expect(await unsaved(page)).toBe(true);
    expect(await lastState(page)).toEqual({ dirty: true, saving: false });
    expect(await asksBeforeLeaving(page)).toBe(true);

    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');
    expect(await onServer(request, url)).toEqual({ format: 'xlsx', text: '42|', type: XLSX_TYPE, said: 'xlsx', from: undefined });
    expect(await unsaved(page)).toBe(false);
    expect(await asksBeforeLeaving(page)).toBe(false);
    const state = await lastState(page);
    expect([state.dirty, state.saving, typeof state.lastSavedAt, 'error' in state]).toEqual([false, false, 'number', false]);
    expect(await notices(page)).toContain('info: Saved');
    await expect(page.locator('#editor .bym-status-message')).toHaveText('Saved');

    // The server answered with a new ETag: the second save is sent with that one, and is taken.
    await typeInGrid(page, '7');
    await expect(save).toBeEnabled();
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect((await onServer(request, url)).text).toBe('42|7');
    expect(await page.evaluate(() => (window as any).__errors)).toEqual([]);
  });

  test('workbook: what is still being typed in a cell is applied and saved; Undo makes the file unsaved again', async ({ page, request, hasTouch }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    const save = page.locator(SAVE);
    await typeInGrid(page, 'draft', false);
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing)).toBe(true);
    // Nothing is written to the workbook yet, but there is something to save.
    await expect(save).toBeEnabled();
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing)).toBe(false);
    expect((await onServer(request, url)).text).toBe('draft|');

    await page.evaluate(() => (window as any).__mounted.view.focus());
    await page.keyboard.press('ControlOrMeta+z');
    await expect(save).toBeEnabled();
    expect(await unsaved(page)).toBe(true);
    await page.keyboard.press('ControlOrMeta+s');
    await expect(save).toBeDisabled();
    expect((await onServer(request, url)).text).toBe('|');
  });

  test('document: text is typed, Save sends the docx back; Ctrl/Cmd+S does the same and the browser does not get the key', async ({ page, request, hasTouch }) => {
    const url = address('docx');
    await seed(request, url, await createBlank('docx'));
    expect(await mount(page, { file: url, saveTo: true, locale: 'tr', onSave: true })).toBe('ok');
    const save = page.locator(SAVE);
    await expect(save).toBeVisible();
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-label', /^Kaydet \(/);

    await typeInDocument(page, 'Merhaba ');
    await expect(save).toBeEnabled();
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect(await onServer(request, url)).toEqual({ format: 'docx', text: 'Merhaba ', type: DOCX_TYPE, said: 'docx', from: undefined });
    expect(await notices(page)).toContain('info: Kaydedildi');
    expect(await page.evaluate(() => (window as any).__mounted.view.isModified)).toBe(false);

    await typeInDocument(page, 'dünya, ');
    await expect(save).toBeEnabled();
    await page.keyboard.press('ControlOrMeta+s');
    await expect(save).toBeDisabled();
    expect((await onServer(request, url)).text).toBe('dünya, Merhaba ');
    await expect.poll(() => page.evaluate(() => (window as any).__keys)).toEqual([true]);
    // onSave comes after the save, with the bytes that were sent.
    const called: string[] = await page.evaluate(() => (window as any).__onSave);
    expect(called.length).toBe(2);
    for (const entry of called) expect(entry).toMatch(/^\d+ bytes, dirty false$/);

    // Nothing to save: the key is still the editor's, and nothing is sent.
    let puts = 0;
    page.on('request', (sent) => {
      if (sent.method() === 'PUT') puts++;
    });
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => page.evaluate(() => (window as any).__keys)).toEqual([true, true]);
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
    expect(puts).toBe(0);
  });

  test('a file that was changed on the server since it was opened is not overwritten: 412 is told, and the file stays unsaved', async ({ page, request, hasTouch, pageProblems }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    // Someone else saves the file.
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Theirs' }));
    await typeInGrid(page, 'mine');
    const save = page.locator(SAVE);
    await press(save, hasTouch);
    await expect(page.locator('#editor .bym-status-message')).toHaveText('The file was not saved: the file on the server is not the one you opened (someone else changed it, or an earlier save of yours was written although its answer did not arrive)');
    await expect(save).toBeEnabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');
    expect(await lastState(page)).toEqual({ dirty: true, saving: false, error: 'The file was not saved: the file on the server is not the one you opened (someone else changed it, or an earlier save of yours was written although its answer did not arrive)' });
    expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError conflict 412']);
    expect((await onServer(request, url)).text).toBe('|');
    expect(await asksBeforeLeaving(page)).toBe(true);
    // Pressing Save again does not overwrite either; saying so does, and from then on saves are taken again.
    await press(save, hasTouch);
    await expect.poll(() => page.evaluate(() => (window as any).__errors.length)).toBe(2);
    expect((await onServer(request, url)).text).toBe('|');
    expect(await page.evaluate(() => (window as any).__mounted.saveNow({ overwrite: true }))).toBe(true);
    await expect(save).toBeDisabled();
    expect((await onServer(request, url)).text).toBe('mine|');
    await typeInGrid(page, 'more');
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect((await onServer(request, url)).text).toBe('mine|more');
    expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError conflict 412', 'SaveError conflict 412']);
    // The browser notes the refused request in its console; nothing else is there.
    expect(pageProblems.filter((problem) => !problem.includes(url))).toEqual([]);
    pageProblems.length = 0;
  });

  test('a save the server refuses, or that does not get through, is told in words of its own; a later one succeeds', async ({ page, request, hasTouch, pageProblems }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    let answer: number | 'abort' | 'pass' = 403;
    const bodies: string[] = [];
    await page.route(`**${url}`, async (route) => {
      if (route.request().method() !== 'PUT' || answer === 'pass') return route.fallback();
      if (answer === 'abort') return route.abort('connectionrefused');
      // A long answer that nobody reads.
      bodies.push('sent');
      return route.fulfill({ status: answer, contentType: 'text/html', body: `<script>window.__ran = true</script>${'x'.repeat(100000)}` });
    });
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    const save = page.locator(SAVE);
    const message = page.locator('#editor .bym-status-message');
    await typeInGrid(page, '1');

    await press(save, hasTouch);
    await expect(message).toHaveText('The file was not saved: you are not allowed to save it');
    answer = 423;
    await press(save, hasTouch);
    await expect(message).toHaveText('The file was not saved: it is locked');
    answer = 500;
    await press(save, hasTouch);
    await expect(message).toHaveText('The file could not be saved (HTTP 500)');
    answer = 'abort';
    await press(save, hasTouch);
    await expect(message).toHaveText('The file could not be saved: the server could not be reached');
    await expect(save).toBeEnabled();
    expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError denied 403', 'SaveError locked 423', 'SaveError http 500', 'SaveError network undefined']);
    expect(await page.evaluate(() => (window as any).__ran)).toBeUndefined();
    expect(await page.locator('#editor script').count()).toBe(0);
    expect((await onServer(request, url)).text).toBe('|');

    answer = 'pass';
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    await expect(message).toHaveText('Saved');
    expect((await onServer(request, url)).text).toBe('1|');
    expect(await lastState(page)).not.toHaveProperty('error');
    // The browser notes the requests that failed in its console; nothing else is there.
    expect(pageProblems.filter((problem) => !problem.includes(url))).toEqual([]);
    pageProblems.length = 0;
  });

  test('while a save is under way the button is busy; what is typed meanwhile stays unsaved, and a second press sends once more afterwards', async ({ page, request, hasTouch }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    let release: () => void = () => undefined;
    const held = new Promise<void>((resolve) => (release = resolve));
    let puts = 0;
    await page.route(`**${url}`, async (route) => {
      if (route.request().method() !== 'PUT') return route.fallback();
      if (++puts === 1) await held;
      return route.fallback();
    });
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    const save = page.locator(SAVE);
    await typeInGrid(page, 'first');
    await press(save, hasTouch);
    await expect(save).toHaveAttribute('aria-busy', 'true');
    await expect.poll(() => puts).toBe(1);
    expect(await lastState(page)).toEqual({ dirty: true, saving: true });
    // A second press with nothing new to send is the same save.
    await press(save, hasTouch);
    await typeInGrid(page, 'second');
    // And one with something new waits for the first.
    await press(save, hasTouch);
    expect(puts).toBe(1);
    release();
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');
    expect(puts).toBe(2);
    expect((await onServer(request, url)).text).toBe('first|second');
    expect(await unsaved(page)).toBe(false);
  });

  test('leaving the page with unsaved changes asks first; not when the option turns it off, and not after the view is destroyed', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    expect(await asksBeforeLeaving(page)).toBe(false);
    await typeInGrid(page, 'x');
    expect(await asksBeforeLeaving(page)).toBe(true);
    await page.evaluate(() => (window as any).__mounted.destroy());
    expect(await asksBeforeLeaving(page)).toBe(false);
    await expect(page.locator('#editor')).toBeEmpty();

    expect(await mount(page, { file: url, saveTo: true, warnOnLeave: false })).toBe('ok');
    await typeInGrid(page, 'y');
    expect(await unsaved(page)).toBe(true);
    expect(await asksBeforeLeaving(page)).toBe(false);
    await page.evaluate(() => (window as any).__mounted.destroy());

    // Without saveTo the page is never asked for: nobody said the editor saves.
    expect(await mount(page, { file: url })).toBe('ok');
    await typeInGrid(page, 'z');
    expect(await unsaved(page)).toBe(true);
    expect(await asksBeforeLeaving(page)).toBe(false);
    expect(await page.locator(SAVE).count()).toBe(0);
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(false);
  });

  test('read-only: no Save button and nothing is sent; a file that cannot be written has none either; converted, it is saved as the new format', async ({ page, request, hasTouch }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    let puts = 0;
    page.on('request', (sent) => {
      if (sent.method() === 'PUT') puts++;
    });
    expect(await mount(page, { file: url, saveTo: true, readOnly: true })).toBe('ok');
    // The toolbar of a read-only workbook is there (it prints), without the button and its separator.
    await expect(page.locator('#editor .bym-toolbar [data-command="print"]')).toBeVisible();
    await expect(page.locator(SAVE)).toBeHidden();
    expect(await page.locator('#editor .bym-toolbar .bym-tb-sep').first().isHidden()).toBe(true);
    await page.evaluate(() => (window as any).__mounted.view.focus());
    await page.keyboard.type('no');
    await page.keyboard.press('ControlOrMeta+s');
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
    // Made editable, the button appears.
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await expect(page.locator(SAVE)).toBeVisible();
    await expect(page.locator(SAVE)).toBeDisabled();
    await page.evaluate(() => (window as any).__mounted.destroy());

    const docUrl = address('docx');
    await seed(request, docUrl, await createBlank('docx'));
    expect(await mount(page, { file: docUrl, saveTo: true, readOnly: true })).toBe('ok');
    await expect(page.locator(SAVE)).toBeHidden();
    await page.evaluate(() => (window as any).__mounted.destroy());
    expect(puts).toBe(0);

    // An xls as it is cannot be written: there is no Save, and saveNow() says so.
    const xlsUrl = address('xls');
    await seed(request, xlsUrl, xlsFile(workbookStream({ sheets: [{ name: 'Old', records: cat(number(0, 0, 41)) }] })));
    puts = 0;
    expect(await mount(page, { file: xlsUrl, saveTo: true })).toBe('ok');
    await expect(page.locator(SAVE)).toBeHidden();
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(false);
    await page.evaluate(() => (window as any).__mounted.destroy());
    expect(puts).toBe(0);

    // Converted to be edited, it would replace the xls with an xlsx: not unless the page allows that. It is
    // said when the file is opened, and again when the user saves; nothing is sent.
    const refusal = 'error: Changes to this file cannot be saved here: it was converted to xlsx, and this page has not allowed saving that in place of the original file';
    for (const target of [true, xlsUrl]) {
      expect(await mount(page, { file: xlsUrl, saveTo: target, convert: true })).toBe('ok');
      expect((await notices(page)).filter((notice) => notice === refusal).length).toBe(1);
      await expect(page.locator(SAVE)).toBeVisible();
      await page.evaluate(() => (window as any).__mounted.view.select({ top: 1, left: 0, bottom: 1, right: 0 }));
      await typeInGrid(page, '9');
      await press(page.locator(SAVE), hasTouch);
      await expect.poll(() => page.evaluate(() => (window as any).__errors)).toEqual(['SaveError converted undefined']);
      expect((await notices(page)).filter((notice) => notice === refusal).length).toBe(2);
      expect(await unsaved(page)).toBe(true);
      await page.evaluate(() => (window as any).__mounted.destroy());
    }
    expect(puts).toBe(0);
    expect((await onServer(request, xlsUrl)).format).toBe('xls');

    // Allowed, it is saved to the same address as an xlsx, and the request says what it was.
    expect(await mount(page, { file: xlsUrl, saveTo: { allowConverted: true }, convert: true })).toBe('ok');
    expect((await notices(page)).filter((notice) => notice === refusal)).toEqual([]);
    const save = page.locator(SAVE);
    await expect(save).toBeDisabled();
    await page.evaluate(() => (window as any).__mounted.view.select({ top: 1, left: 0, bottom: 1, right: 0 }));
    await typeInGrid(page, '9');
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect(await onServer(request, xlsUrl)).toEqual({ format: 'xlsx', text: '41|9', type: XLSX_TYPE, said: 'xlsx', from: 'xls' });
  });

  test('review: a save the server redirects (to a login page) is not called saved, and the redirect is not followed', async ({ page, request, hasTouch, pageProblems }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    // Where the server redirects to: a client that followed would leave a file there (the method and the body are kept).
    const followed = url.replace('/e2e/saved/', '/e2e/saved/followed-');
    for (const status of [302, 307]) {
      expect(await mount(page, { file: url, saveTo: { url: `${url}?status=${status}` } })).toBe('ok');
      const save = page.locator(SAVE);
      await typeInGrid(page, 'lost?');
      await press(save, hasTouch);
      await expect(page.locator('#editor .bym-status-message')).toHaveText('The file was not saved: the server sent the request to another address; you may have been signed out');
      await expect(save).toBeEnabled();
      expect(await lastState(page)).toMatchObject({ dirty: true, saving: false });
      expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError redirected undefined']);
      expect(await notices(page)).not.toContain('info: Saved');
      expect(await asksBeforeLeaving(page)).toBe(true);
      // The address the server points to was never sent the file.
      expect((await request.get(followed)).status(), `followed ${status}`).toBe(404);
      await page.evaluate(() => (window as any).__mounted.destroy());
    }
    expect((await onServer(request, url)).text).toBe('|');
    expect(pageProblems.filter((problem) => !problem.includes(url))).toEqual([]);
    pageProblems.length = 0;
  });

  test('review: Save pressed with the keyboard leaves the focus in the editor; the key saves from a cell that is being typed in and under a rule dialog with changes', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
    const save = page.locator(SAVE);
    const focused = () => page.evaluate(() => (document.activeElement ? `${document.activeElement.tagName} ${document.activeElement.className}` : 'none'));
    await typeInGrid(page, 'a');
    // The toolbar follows in the next frame: a button that is still dimmed takes no focus.
    await expect(save).toBeEnabled();
    await save.focus();
    expect(await focused()).toBe('BUTTON bym-tb-btn');
    await page.keyboard.press('Enter');
    await expect(save).toBeDisabled();
    // The button is dimmed now: the focus is not left on it (nor on the page), it is back in the grid.
    expect(await focused()).toBe('TEXTAREA bym-input');
    expect((await onServer(request, url)).text).toBe('a|');

    // From the cell editor: the browser does not get the key, and what was typed is in the file.
    await typeInGrid(page, 'typed', false);
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing)).toBe(true);
    await page.keyboard.press('ControlOrMeta+s');
    await expect(save).toBeDisabled();
    await expect.poll(() => unsaved(page)).toBe(false);
    expect((await onServer(request, url)).text).toBe('a|typed');
    await expect.poll(() => page.evaluate(() => (window as any).__keys)).toEqual([true]);

    // Under a dialog of rules with changes: the key is still the editor's, nothing is saved, and it is said why.
    await typeInGrid(page, 'held');
    expect(await page.evaluate(() => (window as any).__mounted.view.openValidationDialog())).toBe(true);
    const dialog = page.locator('#editor [data-bym-dialog="validation"]');
    await dialog.locator('[data-bym-field="type"]').selectOption('whole');
    await dialog.locator('[data-bym-field="value1"]').fill('4');
    await page.keyboard.press('ControlOrMeta+s');
    await expect(dialog.locator('[role="alert"]')).toHaveText('Confirm or close the open rule dialog first');
    await expect.poll(() => page.evaluate(() => (window as any).__keys)).toEqual([true, true]);
    expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError pending undefined']);
    expect(await lastState(page)).toEqual(expect.objectContaining({ dirty: true, saving: false, error: 'The file was not saved: what is being typed could not be applied' }));
    expect((await onServer(request, url)).text).toBe('a|typed');
    await dialog.locator('[data-bym-action="cancel"]').click();
    await page.evaluate(() => (window as any).__mounted.view.focus());
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => unsaved(page)).toBe(false);
  });

  test('review: destroying the editor itself, not the mounted result, also ends the question before leaving and the save under way', async ({ page, request }) => {
    for (const kind of ['xlsx', 'docx'] as const) {
      const url = address(kind);
      await seed(request, url, await createBlank(kind));
      expect(await mount(page, { file: url, saveTo: true })).toBe('ok');
      if (kind === 'xlsx') await typeInGrid(page, 'x');
      else await typeInDocument(page, 'x');
      expect(await asksBeforeLeaving(page)).toBe(true);
      await page.evaluate(() => (window as any).__mounted.editor.destroy());
      expect(await asksBeforeLeaving(page), kind).toBe(false);
      await expect(page.locator('#editor')).toBeEmpty();
      expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(false);
      // Twice, and through the result, does no harm.
      await page.evaluate(() => {
        (window as any).__mounted.destroy();
        (window as any).__mounted.editor.destroy();
      });
    }
  });

  test('toolbar: false leaves no button, and the key and saveNow() save all the same', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx'));
    expect(await mount(page, { file: url, saveTo: { method: 'POST', headers: { 'X-Token': 'abc' } }, toolbar: false })).toBe('ok');
    expect(await page.locator('#editor .bym-toolbar').count()).toBe(0);
    const sent: (string | undefined)[] = [];
    page.on('request', (made) => {
      if (made.method() === 'POST') sent.push(made.headers()['x-token'], made.headers()['if-match'] === undefined ? 'no If-Match' : 'If-Match');
    });
    await typeInGrid(page, 'a');
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => unsaved(page)).toBe(false);
    expect((await onServer(request, url)).text).toBe('a|');
    await typeInGrid(page, 'b');
    expect(await page.evaluate(() => (window as any).__mounted.saveNow())).toBe(true);
    expect((await onServer(request, url)).text).toBe('a|b');
    expect(sent).toEqual(['abc', 'If-Match', 'abc', 'If-Match']);
  });

  test('a function saves by itself: it is given the bytes and what they are; one that rejects is a save that failed, shown as text', async ({ page, hasTouch }) => {
    const bytes = Array.from(await createBlank('xlsx'));
    expect(await mount(page, { file: bytes, saver: 'keeps' })).toBe('ok');
    const save = page.locator(SAVE);
    await typeInGrid(page, 'kept');
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    const kept = await page.evaluate(async () => {
      const w = window as any;
      const file = await w.BymOffice.open(new Uint8Array(w.__kept[0].bytes));
      const workbook = await file.workbook();
      return { count: w.__kept.length, info: w.__kept[0].info, a1: w.BymOffice.getCell(workbook.sheets[0], 0, 0).value };
    });
    expect(kept).toEqual({ count: 1, info: { format: 'xlsx' }, a1: 'kept' });
    await page.evaluate(() => (window as any).__mounted.destroy());

    expect(await mount(page, { file: bytes, saver: 'rejects' })).toBe('ok');
    await typeInGrid(page, 'x');
    await press(page.locator(SAVE), hasTouch);
    const message = page.locator('#editor .bym-status-message');
    await expect(message).toHaveText('The file could not be saved (The disk is <b>full</b>)');
    expect(await message.locator('b').count()).toBe(0);
    await expect(page.locator(SAVE)).toBeEnabled();
    expect(await page.evaluate(() => (window as any).__errors)).toEqual(['SaveError refused undefined']);
  });

  test('a saveTo that cannot be saved to makes mount() reject before anything is fetched or shown', async ({ page }) => {
    const bytes = Array.from(await createBlank('xlsx'));
    const fetched: string[] = [];
    page.on('request', (made) => fetched.push(new URL(made.url()).pathname));
    expect(await mount(page, { file: bytes, saveTo: true })).toMatch(/^TypeError: "saveTo" needs an address/);
    for (const bad of ['javascript:alert(1)', 'data:text/plain,x', 'file:///etc/passwd', ' JavaScript:alert(1)', '']) {
      expect(await mount(page, { file: '/e2e/saved/never-fetched.xlsx', saveTo: bad }), JSON.stringify(bad)).toMatch(/^TypeError: "saveTo" must be an http\(s\) address/);
      expect(await mount(page, { file: bytes, saveTo: { url: bad } }), JSON.stringify(bad)).toMatch(/^TypeError: "saveTo" must be an http\(s\) address/);
    }
    expect(await mount(page, { file: bytes, saveTo: { url: '/e2e/saved/x.xlsx', method: 'DELETE' } })).toMatch(/^TypeError: "saveTo.method"/);
    expect(await mount(page, { file: bytes, saveTo: 7 })).toMatch(/^TypeError: "saveTo" must be true/);
    await expect(page.locator('#editor')).toBeEmpty();
    expect(fetched.filter((path) => path.includes('/e2e/saved/'))).toEqual([]);
    // false is "no saveTo".
    expect(await mount(page, { file: bytes, saveTo: false })).toBe('ok');
    expect(await page.locator(SAVE).count()).toBe(0);
  });
});

test.describe('saveTo in the test page (demo)', () => {
  test('a blank workbook is put on the server, edited and saved back; a copy changed by someone else is not overwritten', async ({ page, hasTouch, pageProblems }) => {
    await page.goto('/demo/index.html');
    const state = page.locator('#saveto-state');
    const notice = page.locator('#saveto-notice');
    const save = page.locator('#saveto-editor .bym-toolbar [data-command="save"]');
    await press(page.locator('#saveto-xlsx'), hasTouch);
    await expect(state).toHaveText('Değişiklik yok. — dirty: false, saving: false');
    await expect(save).toBeDisabled();
    const type = async (text: string) => {
      await page.evaluate(() => (window as any).__saveTo.view.focus());
      await page.keyboard.type(text);
      await page.keyboard.press('Enter');
    };
    await type('deneme');
    await expect(state).toHaveText('Kaydedilmemiş değişiklik var. — dirty: true, saving: false');
    await press(save, hasTouch);
    await expect(state).toHaveText(/^Kaydedildi \(.+\)\. — dirty: false, saving: false$/);
    await press(page.locator('#saveto-check'), hasTouch);
    await expect(notice).toHaveText(/^Sunucudaki kopya: .+ bayt, xlsx — Sayfa1!A1 = deneme$/);

    await press(page.locator('#saveto-conflict'), hasTouch);
    await expect(notice).toHaveText('Sunucudaki kopya değiştirildi: bir değişiklik yapıp kaydedin.');
    await type('ikinci');
    await press(save, hasTouch);
    await expect(state).toHaveText(/^Kaydedilemedi: Dosya kaydedilmedi: sunucudaki dosya sizin açtığınızdan farklı .+ — dirty: true, saving: false$/);
    await press(page.locator('#saveto-check'), hasTouch);
    await expect(notice).toHaveText(/^Sunucudaki kopya: .+ bayt, xlsx — Sheet1!A1 = \(boş\)$/);
    // The user keeps their own: the copy on the server is overwritten, because they said so.
    await press(page.locator('#saveto-overwrite'), hasTouch);
    await expect(state).toHaveText(/^Kaydedildi \(.+\)\. — dirty: false, saving: false$/);
    await press(page.locator('#saveto-check'), hasTouch);
    await expect(notice).toHaveText(/^Sunucudaki kopya: .+ bayt, xlsx — Sayfa1!A1 = deneme$/);
    // The browser notes the refused request in its console; nothing else is there.
    expect(pageProblems.filter((problem) => !problem.includes('/e2e/saved/demo-'))).toEqual([]);
    pageProblems.length = 0;
  });
});

test.describe('saveTo on the bridge page: load with save', () => {
  const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');
  const messages = (page: Page): Promise<any[]> => page.evaluate(() => ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)));
  const send = (page: Page, message: unknown) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
  const waitFor = async (page: Page, type: string, count = 1): Promise<any> => {
    let found: any[] = [];
    await expect.poll(async () => (found = (await messages(page)).filter((message) => message.type === type)).length, { message: `${count} x ${type}` }).toBeGreaterThanOrEqual(count);
    return found[count - 1];
  };
  const firstCells = async (data: string): Promise<string> => {
    const workbook = await (await open(new Uint8Array(Buffer.from(data, 'base64')))).workbook();
    return [getCell(workbook.sheets[0], 0, 0), getCell(workbook.sheets[0], 1, 0)].map((cell) => (cell ? String(cell.value) : '')).join('|');
  };

  test.beforeEach(async ({ page }) => {
    // What a WebView's JavaScript channel looks like to the page.
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto('/dist/bymoffice.webview.html');
  });

  test('the user saves with the button and with the key: the application is handed the file, and its saveResult says what became of it', async ({ page, hasTouch }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), name: 'liste.xlsx', save: true, toolbar: true });
    const loaded = await waitFor(page, 'loaded');
    expect([loaded.format, loaded.save]).toEqual(['xlsx', true]);
    const save = page.locator('#editor .bym-toolbar [data-command="save"]');
    await expect(save).toBeVisible();
    await expect(save).toBeDisabled();

    const input = page.locator('#editor .bym-input');
    await input.focus();
    await page.keyboard.type('12');
    await page.keyboard.press('Enter');
    await expect(save).toBeEnabled();
    await press(save, hasTouch);
    const saved = await waitFor(page, 'saved');
    expect([saved.by, saved.format, saved.name, 'id' in saved, 'convertedFrom' in saved]).toEqual(['user', 'xlsx', 'liste.xlsx', false, false]);
    expect(await firstCells(saved.data)).toBe('12|');
    // The application is storing the file.
    await expect(save).toHaveAttribute('aria-busy', 'true');
    expect(saved.saveId).toBe('u1');
    // A result that names no save, or another one, answers nothing.
    await send(page, { type: 'saveResult', ok: true });
    await send(page, { type: 'saveResult', saveId: 'u0', ok: true });
    await waitFor(page, 'error', 2);
    await expect(save).toHaveAttribute('aria-busy', 'true');
    await send(page, { type: 'saveResult', saveId: saved.saveId, ok: true });
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');
    await expect(page.locator('#editor .bym-status-message')).toHaveText('Saved');

    // With the key, and this time the application could not store it.
    await input.focus();
    await page.keyboard.type('34');
    await page.keyboard.press('Enter');
    await page.keyboard.press('ControlOrMeta+s');
    const again = await waitFor(page, 'saved', 2);
    expect(again.by).toBe('user');
    expect(await firstCells(again.data)).toBe('12|34');
    await send(page, { type: 'saveResult', saveId: again.saveId, ok: false, message: 'No <i>room</i> left' });
    await expect(page.locator('#editor .bym-status-message')).toHaveText('The file could not be saved (No <i>room</i> left)');
    expect(await page.locator('#editor .bym-status-message i').count()).toBe(0);
    await expect(save).toBeEnabled();

    // A save the application asks for itself is answered as before.
    await send(page, { type: 'save', id: 's' });
    const own = await waitFor(page, 'saved', 3);
    expect([own.id, 'by' in own]).toEqual(['s', false]);
  });

  test('without save in the load there is no button, and the key is not taken', async ({ page }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), toolbar: true });
    const loaded = await waitFor(page, 'loaded');
    expect('save' in loaded).toBe(false);
    await expect(page.locator('#editor .bym-toolbar [data-command="undo"]')).toBeVisible();
    expect(await page.locator('#editor .bym-toolbar [data-command="save"]').count()).toBe(0);
    await page.evaluate(() => {
      (window as any).__keys = [];
      window.addEventListener('keydown', (event) => {
        if (event.key === 's' || event.key === 'S') (window as any).__keys.push(event.defaultPrevented);
      });
    });
    const input = page.locator('#editor .bym-input');
    await input.focus();
    await page.keyboard.type('1');
    await page.keyboard.press('Enter');
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => page.evaluate(() => (window as any).__keys)).toEqual([false]);
    expect((await messages(page)).filter((message) => message.type === 'saved')).toEqual([]);
  });

  test('a document: the Save button of its toolbar hands the docx to the application', async ({ page, hasTouch }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('docx')), save: true, toolbar: true, locale: 'tr' });
    const loaded = await waitFor(page, 'loaded');
    expect([loaded.kind, loaded.save]).toEqual(['document', true]);
    const save = page.locator('#editor .bym-toolbar [data-command="save"]');
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-label', /^Kaydet \(/);
    await page.locator('#editor .bym-doc-page').first().click({ position: { x: 150, y: 150 } });
    await page.keyboard.type('Köprü');
    await expect(save).toBeEnabled();
    await press(save, hasTouch);
    const saved = await waitFor(page, 'saved');
    expect([saved.by, saved.format]).toEqual(['user', 'docx']);
    const file = await open(new Uint8Array(Buffer.from(saved.data, 'base64')));
    expect(documentText(await file.document!())).toBe('Köprü');
    await send(page, { type: 'saveResult', saveId: saved.saveId, ok: true });
    await expect(save).toBeDisabled();
  });
});
