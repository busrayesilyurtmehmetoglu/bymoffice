import type { APIRequestContext, Locator, Page } from '@playwright/test';
import { createBlank, documentText, getCell, open } from '../src/index';
import { expect, test } from './helpers';

// MountOptions.autoSave: the editor saves by itself, a delay after the last change, without
// cutting short what is being typed and without asking about a conflict on its own. The time of
// the page is the test's (page.clock): nothing here waits for real seconds to pass. The server
// of these tests keeps what it is sent at /e2e/saved/<name> and honours If-Match.

const PAGE = '/e2e/pages/mount.html';
const SAVE = '#editor .bym-toolbar [data-command="save"]';
const CONFIRM = '#editor [data-bym-confirm="save-conflict"]';

/** An address of the server that no other test uses. */
function address(extension: string): string {
  const info = test.info();
  return `/e2e/saved/auto-${`${info.project.name}-${info.testId}`.replace(/[^A-Za-z0-9_-]/g, '_')}.${extension}`;
}

async function seed(request: APIRequestContext, url: string, bytes: Uint8Array): Promise<void> {
  const answer = await request.put(url, { data: Buffer.from(bytes), headers: { 'content-type': 'application/octet-stream' } });
  expect(answer.status(), `seeding ${url}`).toBe(204);
}

/** What the server holds at an address: the first two cells of a workbook (and the name of its sheet), or the text of a document. */
async function onServer(request: APIRequestContext, url: string): Promise<string> {
  const answer = await request.get(url);
  expect(answer.status(), `reading ${url}`).toBe(200);
  const file = await open(new Uint8Array(await answer.body()));
  if (file.kind === 'document') return documentText(await file.document!());
  const workbook = await file.workbook();
  return `${workbook.sheets[0].name}:${[getCell(workbook.sheets[0], 0, 0), getCell(workbook.sheets[0], 1, 0)].map((cell) => (cell ? String(cell.value) : '')).join('|')}`;
}

interface Spec {
  file: string | number[];
  saveTo?: unknown;
  autoSave?: unknown;
}

/** Mounts in the page; resolves to "ok" or to what mount() rejected with. */
function mount(page: Page, spec: Spec): Promise<string> {
  return page.evaluate(async (given) => {
    const w = window as any;
    w.__states = [];
    w.__notices = [];
    w.__errors = [];
    const options: any = {
      file: typeof given.file === 'string' ? given.file : new Uint8Array(given.file),
      locale: 'en',
      onSaveState: (state: unknown) => w.__states.push(state),
      onNotice: (notice: { kind: string; message: string }) => w.__notices.push(`${notice.kind}: ${notice.message}`),
      onError: (error: { code?: string }) => w.__errors.push(String(error.code)),
    };
    if ('saveTo' in given) options.saveTo = given.saveTo;
    if ('autoSave' in given) options.autoSave = given.autoSave;
    try {
      w.__mounted = await w.BymOffice.mount(document.getElementById('editor'), options);
      return 'ok';
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  }, spec);
}

const lastState = (page: Page) => page.evaluate(() => (window as any).__states[(window as any).__states.length - 1]);
const errors = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__errors);
const notices = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__notices);
const press = (target: Locator, hasTouch: boolean | undefined) => (hasTouch ? target.tap() : target.click());

/** Stops the time of the page (page.clock.install() came first): from here on it passes only when the test says. */
const stopTime = async (page: Page): Promise<void> => page.clock.pauseAt((await page.evaluate(() => Date.now())) + 1000);
/**
 * Lets a few frames pass. With the time stopped requestAnimationFrame stands still as well, and the
 * toolbar is drawn from it: what a button shows is looked at only after this.
 */
const frames = (page: Page): Promise<void> => page.clock.runFor(100);
/** Whether the page would ask before it is left: what the browser asks its `beforeunload` listeners. */
const asksBeforeLeaving = (page: Page): Promise<boolean> =>
  page.evaluate(() => {
    const event = new Event('beforeunload', { cancelable: true });
    window.dispatchEvent(event);
    return event.defaultPrevented;
  });
/** Takes the ETag off every answer to a PUT of an address: a server that names no version of what it holds. */
const withoutETag = async (page: Page, url: string): Promise<void> => {
  await page.route(`**${url}`, async (route) => {
    if (route.request().method() !== 'PUT') return route.continue();
    const response = await route.fetch();
    const headers = { ...response.headers() };
    delete headers.etag;
    await route.fulfill({ response, headers });
  });
};
const NO_VERSION = 'error: Auto-save is not working: the server did not say which version of the file it holds. Save with the Save button.';

/** Counts the PUT requests the page sends to an address. */
function countPuts(page: Page, url: string): { count: number } {
  const seen = { count: 0 };
  page.on('request', (request) => {
    if (request.method() === 'PUT' && new URL(request.url()).pathname === url) seen.count++;
  });
  return seen;
}

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

test.describe('autoSave: the editor saves by itself', () => {
  test.beforeEach(async ({ page }) => {
    await page.clock.install();
    await page.goto(PAGE);
  });

  test('workbook: what is typed is on the server two seconds after the last change, and nobody pressed Save', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);
    const save = page.locator(SAVE);
    await frames(page);
    await expect(save).toBeDisabled();

    await typeInGrid(page, '42');
    expect((await lastState(page)).auto, 'the state says that auto-save is on').toBe(true);
    await page.clock.runFor(1500);
    await typeInGrid(page, '7');
    // The second change started the two seconds again.
    await page.clock.runFor(1900);
    expect(puts.count, 'nothing is sent before the delay after the last change is over').toBe(0);
    expect(await onServer(request, url)).toBe('Data:|');
    await page.clock.runFor(200);
    await expect.poll(() => onServer(request, url), { message: 'the file is on the server' }).toBe('Data:42|7');
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
    expect(puts.count).toBe(1);
    await frames(page);
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');
    expect(await errors(page)).toEqual([]);
    // A save nobody asked for is not announced: the button and the state say it.
    expect(await notices(page)).toEqual([]);
    expect('autoPaused' in (await lastState(page))).toBe(false);

    // A file that is saved is not sent again.
    await page.clock.runFor(60000);
    expect(puts.count).toBe(1);
  });

  test('workbook: nothing is saved while a cell is being typed, and the cell stays open; after Enter it is saved', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);

    await typeInGrid(page, 'first');
    // The second cell is being typed when the time of the first comes.
    await typeInGrid(page, 'half', false);
    await page.clock.runFor(60000);
    expect(puts.count, 'no save falls into a cell that is being typed').toBe(0);
    expect(await page.evaluate(() => [(window as any).__mounted.view.isEditing, (window as any).__mounted.view.editText])).toEqual([true, 'half']);
    expect(await onServer(request, url)).toBe('Data:|');

    await page.keyboard.press('Enter');
    await page.clock.runFor(2000);
    await expect.poll(() => onServer(request, url)).toBe('Data:first|half');
    expect(puts.count).toBe(1);
  });

  test('workbook: who never stops typing is saved the longest wait after the first change', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: { delayMs: 2000, maxWaitMs: 5000 } })).toBe('ok');
    await stopTime(page);

    // A change every second: the two seconds of quiet never come.
    for (let i = 0; i < 4; i++) {
      await typeInGrid(page, String(i));
      await page.clock.runFor(1000);
      expect(puts.count, `after ${i + 1} s`).toBe(0);
    }
    await typeInGrid(page, '4');
    await page.clock.runFor(1000);
    await expect.poll(() => puts.count, { message: 'five seconds after the first change' }).toBe(1);
    await expect.poll(async () => (await lastState(page)).saving).toBe(false);
    expect(await onServer(request, url)).toBe('Data:0|1');
  });

  test('workbook: a file changed on the server is not overwritten and not asked about by itself; Save is what asks', async ({ page, request, hasTouch, pageProblems }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);
    // Someone else saves the file meanwhile.
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Theirs' }));

    await typeInGrid(page, 'mine');
    await page.clock.runFor(2000);
    await expect.poll(() => errors(page)).toEqual(['conflict']);
    expect(puts.count).toBe(1);
    await expect(page.locator(CONFIRM), 'no question opens under the hands of someone who is typing').toHaveCount(0);
    const state = await lastState(page);
    expect([state.dirty, state.saving, typeof state.error]).toEqual([true, false, 'string']);

    // The user types on: Enter lands in the grid, and nothing is sent by itself any more.
    await typeInGrid(page, 'more');
    await page.clock.runFor(120000);
    expect(puts.count, 'auto-save holds still after a conflict').toBe(1);
    await expect(page.locator(CONFIRM)).toHaveCount(0);
    expect(await onServer(request, url)).toBe('Theirs:|');

    expect((await lastState(page)).autoPaused, 'the state says that auto-save waits for the user').toBe(true);

    // The Save button: the same guarded request, and now the question.
    await frames(page);
    await press(page.locator(SAVE), hasTouch);
    await expect(page.locator(CONFIRM)).toBeVisible();
    expect(puts.count).toBe(2);
    expect(await onServer(request, url), 'nothing was overwritten unasked').toBe('Theirs:|');
    // The browser notes the refused requests in its console; nothing else is there.
    expect(pageProblems.filter((problem) => !problem.includes('/e2e/saved/'))).toEqual([]);
    pageProblems.length = 0;
  });

  test('workbook: a save that does not get through is tried again five seconds later, and told once', async ({ page, request, pageProblems }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    // The first PUT never reaches the server.
    let broken = 0;
    await page.route(`**${url}`, async (route) => {
      if (route.request().method() === 'PUT' && broken === 0) {
        broken++;
        return route.abort('failed');
      }
      await route.continue();
    });
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);

    await typeInGrid(page, 'kept');
    await page.clock.runFor(2000);
    await expect.poll(() => errors(page)).toEqual(['network']);
    expect(puts.count).toBe(1);
    await page.clock.runFor(4900);
    expect(puts.count, 'not before five seconds are over').toBe(1);
    await page.clock.runFor(200);
    await expect.poll(() => onServer(request, url)).toBe('Data:kept|');
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
    expect(puts.count).toBe(2);
    expect('error' in (await lastState(page))).toBe(false);
    expect(await errors(page)).toEqual(['network']);
    // Told once that it failed, and once that it is over.
    expect((await notices(page)).filter((notice) => notice.startsWith('error: '))).toHaveLength(1);
    expect((await notices(page)).filter((notice) => notice === 'info: Saved')).toHaveLength(1);
    pageProblems.length = 0;
  });

  test('workbook: a server that names no version of the file is not saved to by itself: said once, and the Save button works', async ({ page, request, hasTouch }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    await withoutETag(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);

    // The file came with an ETag: the first save by itself is guarded with it.
    await typeInGrid(page, 'one');
    await page.clock.runFor(2000);
    await expect.poll(() => onServer(request, url)).toBe('Data:one|');
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
    expect((await lastState(page)).autoPaused, 'the answer named no version: auto-save has nothing to guard the next save with').toBe(true);

    // Someone else saves meanwhile. Nothing goes over it unasked.
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Theirs' }));
    await typeInGrid(page, 'two');
    await page.clock.runFor(120000);
    expect(puts.count, 'no request without a condition').toBe(1);
    expect(await onServer(request, url)).toBe('Theirs:|');
    expect(await notices(page)).toEqual([NO_VERSION]);
    expect(await errors(page)).toEqual([]);
    await typeInGrid(page, 'three');
    await page.clock.runFor(120000);
    expect(await notices(page), 'said once').toEqual([NO_VERSION]);
    expect(puts.count).toBe(1);

    // The user saves, as ever (a save without an ETag carries no condition: that is the user's doing).
    await frames(page);
    await press(page.locator(SAVE), hasTouch);
    await expect.poll(() => puts.count).toBe(2);
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
  });

  test('a file given as bytes has no version to guard a save with: auto-save waits, unless the page says { unconditional: true }', async ({ page, request }) => {
    const url = address('xlsx');
    const bytes = Array.from(await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: bytes, saveTo: url, autoSave: true })).toBe('ok');
    await stopTime(page);
    await typeInGrid(page, 'kept');
    await page.clock.runFor(60000);
    expect(puts.count).toBe(0);
    expect(await notices(page)).toEqual([NO_VERSION]);
    expect((await lastState(page)).autoPaused).toBe(true);

    // The page takes the risk on itself.
    await page.evaluate(() => (window as any).__mounted.setAutoSave({ unconditional: true }));
    await page.clock.runFor(2000);
    await expect.poll(() => puts.count).toBe(1);
    await expect.poll(() => onServer(request, url)).toBe('Data:kept|');
  });

  test('workbook: half a cell typed after a save by itself is unsaved work: hasUnsavedChanges and the question before leaving say so, the state of the file does not', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);
    const unsaved = () => page.evaluate(() => (window as any).__mounted.hasUnsavedChanges);

    await typeInGrid(page, 'done');
    await page.clock.runFor(2000);
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
    expect(await unsaved()).toBe(false);
    expect(await asksBeforeLeaving(page)).toBe(false);

    await typeInGrid(page, 'half', false);
    expect((await lastState(page)).dirty, 'the file itself is as the server holds it').toBe(false);
    expect(await unsaved()).toBe(true);
    expect(await asksBeforeLeaving(page)).toBe(true);
    // Given up with Escape: nothing is left to lose.
    await page.keyboard.press('Escape');
    expect(await unsaved()).toBe(false);
    expect(await asksBeforeLeaving(page)).toBe(false);
    expect(await onServer(request, url)).toBe('Data:done|');
  });

  test('document: what is typed is saved by itself', async ({ page, request }) => {
    const url = address('docx');
    await seed(request, url, await createBlank('docx'));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);

    await typeInDocument(page, 'Minutes');
    await page.clock.runFor(1900);
    expect(puts.count).toBe(0);
    await page.clock.runFor(200);
    await expect.poll(() => onServer(request, url)).toBe('Minutes');
    await expect.poll(async () => (await lastState(page)).dirty).toBe(false);
    expect(puts.count).toBe(1);
    // The caret is where it was: typing goes on in the same paragraph.
    await page.keyboard.type(' of May');
    await page.clock.runFor(2000);
    await expect.poll(() => onServer(request, url)).toBe('Minutes of May');
  });

  test('a page that is hidden saves what waits at once; setAutoSave turns it off and on', async ({ page, request }) => {
    const url = address('xlsx');
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
    const puts = countPuts(page, url);
    expect(await mount(page, { file: url, saveTo: true, autoSave: true })).toBe('ok');
    await stopTime(page);

    await typeInGrid(page, 'leaving');
    // What a browser tells a page whose tab is left (the event is made here: a test cannot hide its own tab).
    await page.evaluate(() => {
      Object.defineProperty(document, 'visibilityState', { configurable: true, get: () => 'hidden' });
      document.dispatchEvent(new Event('visibilitychange'));
    });
    await expect.poll(() => onServer(request, url), { message: 'saved without any time passing' }).toBe('Data:leaving|');
    expect(puts.count).toBe(1);

    await page.evaluate(() => (window as any).__mounted.setAutoSave(false));
    await typeInGrid(page, 'off');
    await page.clock.runFor(60000);
    expect(puts.count, 'turned off, nothing is sent').toBe(1);
    await page.evaluate(() => (window as any).__mounted.setAutoSave({ delayMs: 500 }));
    await page.clock.runFor(500);
    await expect.poll(() => onServer(request, url)).toBe('Data:leaving|off');
    expect(await page.evaluate(() => {
      try {
        (window as any).__mounted.setAutoSave({ delayMs: 1 });
        return 'taken';
      } catch (error) {
        return (error as Error).name;
      }
    })).toBe('TypeError');
  });

  test('mount() refuses auto-save without saveTo and times that cannot be', async ({ page }) => {
    const bytes = Array.from(await createBlank('xlsx'));
    expect(await mount(page, { file: bytes, autoSave: true })).toMatch(/^TypeError: "autoSave" needs "saveTo"/);
    expect(await mount(page, { file: bytes, saveTo: '/e2e/saved/never.xlsx', autoSave: { delayMs: 100 } })).toMatch(/^TypeError: "autoSave.delayMs"/);
    expect(await mount(page, { file: bytes, saveTo: '/e2e/saved/never.xlsx', autoSave: { delayMs: 1000, maxWaitMs: 500 } })).toMatch(/^TypeError: "autoSave.maxWaitMs"/);
    expect(await mount(page, { file: bytes, autoSave: false })).toBe('ok');
  });
});

test.describe('autoSave in the test page (demo)', () => {
  test('with "Kendiliğinden kaydet" ticked, what is typed is saved without the button', async ({ page, hasTouch }) => {
    await page.clock.install();
    await page.goto('/demo/index.html');
    const state = page.locator('#saveto-state');
    await press(page.locator('#saveto-auto'), hasTouch);
    await press(page.locator('#saveto-xlsx'), hasTouch);
    await expect(state).toHaveText('Değişiklik yok. — dirty: false, saving: false, auto: true');
    await stopTime(page);
    await page.evaluate(() => (window as any).__saveTo.view.focus());
    await page.keyboard.type('deneme');
    await page.keyboard.press('Enter');
    await expect(state).toHaveText('Kaydedilmemiş değişiklik var. — dirty: true, saving: false, auto: true');
    await page.clock.runFor(2000);
    await expect(state).toHaveText(/^Kaydedildi \(.+\)\. — dirty: false, saving: false, auto: true$/);
    await press(page.locator('#saveto-check'), hasTouch);
    await expect(page.locator('#saveto-notice')).toHaveText(/^Sunucudaki kopya: .+ bayt, xlsx — Sayfa1!A1 = deneme$/);
  });
});

test.describe('autoSave on the bridge page: load with save: { auto }', () => {
  const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');
  const messages = (page: Page): Promise<any[]> => page.evaluate(() => ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)));
  const send = (page: Page, message: unknown) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
  const waitFor = async (page: Page, type: string, count = 1): Promise<any> => {
    let found: any[] = [];
    await expect.poll(async () => (found = (await messages(page)).filter((message) => message.type === type)).length, { message: `${count} x ${type}` }).toBeGreaterThanOrEqual(count);
    return found[count - 1];
  };
  const savedCount = async (page: Page): Promise<number> => (await messages(page)).filter((message) => message.type === 'saved').length;
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
    await page.clock.install();
    await page.goto('/dist/bymoffice.webview.html');
  });

  test('the page hands the file over by itself as "saved" by "auto"; a conflict stops it, and the Save of the user is what asks', async ({ page, hasTouch }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), name: 'liste.xlsx', toolbar: true, save: { auto: { delayMs: 1000, maxWaitMs: 5000 } } });
    const loaded = await waitFor(page, 'loaded');
    expect([loaded.save, loaded.autoSave]).toEqual([true, true]);
    await stopTime(page);
    const save = page.locator('#editor .bym-toolbar [data-command="save"]');
    const input = page.locator('#editor .bym-input');

    await input.focus();
    await page.keyboard.type('12');
    await page.keyboard.press('Enter');
    await page.clock.runFor(900);
    expect(await savedCount(page)).toBe(0);
    await page.clock.runFor(200);
    const saved = await waitFor(page, 'saved');
    expect([saved.by, saved.saveId, saved.format, saved.name, 'id' in saved]).toEqual(['auto', 'a1', 'xlsx', 'liste.xlsx', false]);
    expect(await firstCells(saved.data)).toBe('12|');
    await frames(page);
    await expect(save).toHaveAttribute('aria-busy', 'true');
    await send(page, { type: 'saveResult', saveId: 'a1', ok: true });
    await frames(page);
    await expect(save).toBeDisabled();
    await expect(save).toHaveAttribute('aria-busy', 'false');

    // Nothing is handed over while a cell is being typed.
    await input.focus();
    await page.keyboard.type('3');
    await page.keyboard.press('Enter');
    await page.keyboard.type('half');
    await page.clock.runFor(30000);
    expect(await savedCount(page)).toBe(1);
    await page.keyboard.press('Enter');
    await page.clock.runFor(1000);
    const second = await waitFor(page, 'saved', 2);
    expect([second.by, second.saveId]).toEqual(['auto', 'a2']);

    // The application says that it holds another file: no question by itself, and nothing more is handed over.
    await send(page, { type: 'saveResult', saveId: 'a2', ok: false, conflict: true });
    await frames(page);
    await expect(save).toHaveAttribute('aria-busy', 'false');
    await expect(page.locator('#editor [data-bym-confirm="save-conflict"]')).toHaveCount(0);
    await input.focus();
    await page.keyboard.type('5');
    await page.keyboard.press('Enter');
    await page.clock.runFor(60000);
    expect(await savedCount(page)).toBe(2);

    // The user saves: the same file by "user", and the conflict of that save is asked about.
    await frames(page);
    await press(save, hasTouch);
    const third = await waitFor(page, 'saved', 3);
    expect([third.by, third.saveId, 'overwrite' in third]).toEqual(['user', 'u3', false]);
    await send(page, { type: 'saveResult', saveId: 'u3', ok: false, conflict: true });
    await expect(page.locator('#editor [data-bym-confirm="save-conflict"]')).toBeVisible();
  });
});
