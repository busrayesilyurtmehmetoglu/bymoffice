import type { APIRequestContext, Locator, Page } from '@playwright/test';
import { createBlank, documentText, getCell, open } from '../src/index';
import { expect, test } from './helpers';

// A save the server refuses because the file it holds was changed since it was opened (HTTP 412):
// the editor asks, inside its container, whether to overwrite, and does so only when the user
// says so. The server of these tests keeps what it is sent at /e2e/saved/<name> and honours If-Match.

const PAGE = '/e2e/pages/mount.html';
const SAVE = '#editor .bym-toolbar [data-command="save"]';
const CONFIRM = '#editor [data-bym-confirm="save-conflict"]';
const OVERWRITE = `${CONFIRM} [data-bym-action="overwrite"]`;
const CANCEL = `${CONFIRM} [data-bym-action="cancel"]`;
/** "Overwrite" once it answers: for a moment after the question came it does not, and says so with aria-disabled. */
const ARMED = `${OVERWRITE}:not([aria-disabled])`;
const QUESTION_EN =
  'The file was changed on the server after you opened it (someone else changed it, or an earlier save of yours was written although its answer did not arrive). If you save, the changes there are overwritten.';
const QUESTION_TR =
  'Dosya, siz açtıktan sonra sunucuda değişmiş (başkası değiştirmiş ya da önceki kaydınız yanıtı gelmeden yazılmış olabilir). Kaydederseniz oradaki değişikliklerin üzerine yazılır.';

/** An address of the server that no other test uses. */
function address(extension: string): string {
  const info = test.info();
  return `/e2e/saved/conflict-${`${info.project.name}-${info.testId}`.replace(/[^A-Za-z0-9_-]/g, '_')}.${extension}`;
}

async function seed(request: APIRequestContext, url: string, bytes: Uint8Array): Promise<void> {
  const answer = await request.put(url, { data: Buffer.from(bytes), headers: { 'content-type': 'application/octet-stream' } });
  expect(answer.status(), `seeding ${url}`).toBe(204);
}

/** What the server holds at an address: the first two cells of a workbook, or the text of a document. */
async function onServer(request: APIRequestContext, url: string): Promise<string> {
  const answer = await request.get(url);
  expect(answer.status(), `reading ${url}`).toBe(200);
  const file = await open(new Uint8Array(await answer.body()));
  if (file.kind === 'document') return documentText(await file.document!());
  const workbook = await file.workbook();
  return [getCell(workbook.sheets[0], 0, 0), getCell(workbook.sheets[0], 1, 0)].map((cell) => (cell ? String(cell.value) : '')).join('|');
}

interface Spec {
  file: string;
  saveTo?: unknown;
  locale?: string;
  toolbar?: boolean;
  /** The page answers a conflict by itself, from inside onError. */
  overwriteInOnError?: boolean;
}

function mount(page: Page, spec: Spec): Promise<string> {
  return page.evaluate(async (given) => {
    const w = window as any;
    w.__states = [];
    w.__notices = [];
    w.__errors = [];
    w.__answers = [];
    const options: any = {
      file: given.file,
      saveTo: 'saveTo' in given ? given.saveTo : true,
      locale: given.locale || 'en',
      onSaveState: (state: unknown) => w.__states.push(state),
      onNotice: (notice: { kind: string; message: string }) => w.__notices.push(`${notice.kind}: ${notice.message}`),
      onError: (error: { code?: string; status?: number; name?: string }) => {
        // Whether the question was already there when the page heard of the conflict.
        w.__errors.push(`${error.name} ${error.code} ${error.status} asked:${document.querySelector('[data-bym-confirm]') !== null}`);
        if (given.overwriteInOnError && error.code === 'conflict') w.__mounted.saveNow({ overwrite: true }).then((ok: boolean) => w.__answers.push(`own ${ok}`));
      },
    };
    if ('toolbar' in given) options.toolbar = given.toolbar;
    try {
      w.__mounted = await w.BymOffice.mount(document.getElementById('editor'), options);
      return 'ok';
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  }, spec);
}

/** A press with whatever the project has: a finger where there is a touch screen, the mouse elsewhere. */
const press = (target: Locator, hasTouch: boolean | undefined) => (hasTouch ? target.tap() : target.click());
/** Presses "overwrite" once it answers. */
async function overwrite(page: Page, hasTouch: boolean | undefined): Promise<void> {
  await expect(page.locator(ARMED), '"overwrite" answers after a moment').toBeVisible();
  await press(page.locator(OVERWRITE), hasTouch);
}
/** A press at a point of the page, whatever is there by the time it arrives. */
const pressAt = (page: Page, hasTouch: boolean | undefined, at: [number, number]) => (hasTouch ? page.touchscreen.tap(at[0], at[1]) : page.mouse.click(at[0], at[1]));
const centre = async (target: Locator): Promise<[number, number]> => {
  const box = (await target.boundingBox())!;
  return [box.x + box.width / 2, box.y + box.height / 2];
};
/**
 * Holds every PUT to an address back until the returned function is called: the answer of the server comes
 * when the test says. Called with `true`, the function holds the PUTs that come from then on back again.
 */
async function holdPuts(page: Page, url: string): Promise<(again?: boolean) => void> {
  let open: () => void = () => undefined;
  let gate: Promise<void> = Promise.resolve();
  const hold = (): void => {
    gate = new Promise<void>((resolve) => {
      open = resolve;
    });
  };
  hold();
  await page.route(`**${url}`, async (route) => {
    if (route.request().method() === 'PUT') await gate;
    await route.continue();
  });
  return (again) => (again === true ? hold() : open());
}
/** Stops the time of the page (page.clock.install() came first): "a moment" is then as long as the test says. */
const stopTime = async (page: Page): Promise<void> => page.clock.pauseAt((await page.evaluate(() => Date.now())) + 1000);
/** What is at a point of the page: the answer of the question it belongs to, or where it is. */
const topmostAt = (page: Page, at: [number, number]): Promise<string> =>
  page.evaluate(([x, y]) => {
    const element = document.elementFromPoint(x, y);
    if (!element) return 'nothing';
    return element.getAttribute('data-bym-action') || (element.closest('[data-bym-confirm]') ? 'question' : element.closest('.bym-confirm-backdrop') ? 'backdrop' : element.className || element.tagName);
  }, at);
const errors = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__errors);
const unsaved = (page: Page): Promise<boolean> => page.evaluate(() => (window as any).__mounted.hasUnsavedChanges);
/** Asks for a save from the page and notes what its promise resolves to, without waiting for it. */
const saveLater = (page: Page, name: string, options?: { overwrite: boolean }) =>
  page.evaluate(
    ([label, given]) => {
      const w = window as any;
      w.__mounted.saveNow(given || undefined).then((ok: boolean) => w.__answers.push(`${label} ${ok}`));
    },
    [name, options] as const,
  );
const answers = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__answers);

async function typeInGrid(page: Page, text: string): Promise<void> {
  await page.evaluate(() => (window as any).__mounted.view.focus());
  await expect(page.locator('#editor .bym-input')).toBeFocused();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

async function typeInDocument(page: Page, text: string): Promise<void> {
  await page.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.focus();
    view.setSelection({ story: 'body', path: [0], offset: 0 });
  });
  await page.keyboard.type(text);
}

/** The PUT requests the page sends to an address: whether each carried If-Match. */
function watchPuts(page: Page, url: string): string[] {
  const puts: string[] = [];
  page.on('request', (sent) => {
    if (sent.method() === 'PUT' && new URL(sent.url()).pathname === url) puts.push(sent.headers()['if-match'] === undefined ? 'unconditional' : 'if-match');
  });
  return puts;
}

/** The browser notes a refused request in its console; nothing else may be there. */
function onlyRefusals(pageProblems: string[], url: string): void {
  expect(pageProblems.filter((problem) => !problem.includes(url))).toEqual([]);
  pageProblems.length = 0;
}

/** A workbook from the server, changed there by someone else, with a cell typed here. */
async function conflictedWorkbook(page: Page, request: APIRequestContext, spec: Partial<Spec> = {}): Promise<string> {
  const url = address('xlsx');
  await seed(request, url, await createBlank('xlsx'));
  expect(await mount(page, { file: url, ...spec })).toBe('ok');
  await seed(request, url, await createBlank('xlsx', { sheetName: 'Theirs' }));
  await typeInGrid(page, 'mine');
  return url;
}

test.describe('saveTo: a conflict asks before it overwrites', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto(PAGE);
  });

  test('workbook: 412 opens the question with the focus on Cancel; Escape cancels, Save asks again, Overwrite saves over the server', async ({ page, request, hasTouch, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const save = page.locator(SAVE);
    await press(save, hasTouch);

    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await expect(confirm).toHaveAttribute('role', 'alertdialog');
    await expect(confirm).toHaveAccessibleName('The file was changed on the server');
    await expect(confirm).toHaveAccessibleDescription(QUESTION_EN);
    await expect(page.locator(CANCEL)).toBeFocused();
    await expect(page.locator(CANCEL)).toHaveText('Cancel');
    await expect(page.locator(OVERWRITE)).toHaveText('Overwrite');
    expect(await errors(page), 'onError heard of the conflict before the question was there').toEqual(['SaveError conflict 412 asked:false']);
    expect(puts, 'one conditional save, and nothing by itself').toEqual(['if-match']);
    expect(await onServer(request, url)).toBe('|');

    // What is typed while the question is open goes to no cell, and a press on the cells behind it selects nothing.
    await page.keyboard.type('x');
    const box = (await page.locator('#editor').boundingBox())!;
    if (hasTouch) await page.touchscreen.tap(box.x + 20, box.y + box.height - 60);
    else await page.mouse.click(box.x + 20, box.y + box.height - 60);
    await page.keyboard.type('y');
    await expect(confirm).toBeVisible();
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing), 'no cell is being edited').toBe(false);
    // A save asked for meanwhile asks nothing more and sends nothing.
    await page.keyboard.press('ControlOrMeta+s');
    await saveLater(page, 'second');
    await expect(confirm).toHaveCount(1);
    expect(await answers(page)).toEqual([]);

    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    await expect(page.locator('#editor .bym-input'), 'the focus is back in the grid').toBeFocused();
    await expect.poll(() => answers(page), { message: 'the save that waited for the answer has failed' }).toEqual(['second false']);
    expect(puts).toEqual(['if-match']);
    expect(await unsaved(page)).toBe(true);
    await expect(save).toBeEnabled();
    expect(await onServer(request, url), 'nothing was overwritten').toBe('|');
    expect(await page.evaluate(() => (window as any).__mounted.view.workbook.sheets[0].rows.size), 'only the cell that was typed before the question has a value').toBe(1);

    // The next save is sent, refused again, and asks again.
    await press(save, hasTouch);
    await expect(confirm).toBeVisible();
    await expect(page.locator(CANCEL)).toBeFocused();
    expect(puts).toEqual(['if-match', 'if-match']);
    await overwrite(page, hasTouch);
    await expect(confirm).toHaveCount(0);
    await expect(save).toBeDisabled();
    await expect(page.locator('#editor .bym-status-message')).toHaveText('Saved');
    expect(puts, 'the overwrite is one request without a condition').toEqual(['if-match', 'if-match', 'unconditional']);
    expect(await onServer(request, url)).toBe('mine|');
    expect(await unsaved(page)).toBe(false);

    // From here on saves are guarded and taken again.
    await typeInGrid(page, 'more');
    await press(save, hasTouch);
    await expect(save).toBeDisabled();
    expect(await onServer(request, url)).toBe('mine|more');
    await expect(confirm).toHaveCount(0);
    onlyRefusals(pageProblems, url);
  });

  test('the two answers are reached with Tab and not with the arrow keys, Enter on Cancel cancels, and Tab leaves the question for the page', async ({ page, request, browserName, pageProblems }) => {
    await page.evaluate(() => {
      for (const id of ['before', 'after']) {
        const button = document.createElement('button');
        button.id = id;
        button.textContent = `${id} the editor`;
        if (id === 'before') document.body.insertBefore(button, document.body.firstChild);
        else document.body.appendChild(button);
      }
    });
    // Safari leaves buttons out of the Tab order unless Option is held (a setting of the system, not of the page).
    const tab = browserName === 'webkit' ? 'Alt+Tab' : 'Tab';
    const back = browserName === 'webkit' ? 'Alt+Shift+Tab' : 'Shift+Tab';
    const url = await conflictedWorkbook(page, request);
    await page.keyboard.press('ControlOrMeta+s');
    await expect(page.locator(CANCEL)).toBeFocused();
    // The keys that move through cells move nothing here: a user who was on the way through the grid answers nothing.
    for (const arrow of ['ArrowLeft', 'ArrowUp', 'ArrowRight', 'ArrowDown']) {
      await page.keyboard.press(arrow);
      await expect(page.locator(CANCEL), `${arrow} leaves the focus on Cancel`).toBeFocused();
    }
    await page.keyboard.press(back);
    await expect(page.locator(OVERWRITE)).toBeFocused();
    // Nothing of the editor behind the question takes the focus: before the first answer comes the page.
    await page.keyboard.press(back);
    await expect(page.locator('#before'), 'the editor behind the question is passed over').toBeFocused();
    await page.keyboard.press(tab);
    await expect(page.locator(OVERWRITE)).toBeFocused();
    await page.keyboard.press(tab);
    await expect(page.locator(CANCEL)).toBeFocused();
    await page.keyboard.press(tab);
    await expect(page.locator('#after'), 'the focus is not trapped').toBeFocused();
    await expect(page.locator(CONFIRM)).toBeVisible();
    await page.locator(CANCEL).focus();
    await page.keyboard.press('Enter');
    await expect(page.locator(CONFIRM)).toHaveCount(0);
    expect(await unsaved(page)).toBe(true);
    expect(await onServer(request, url)).toBe('|');
    onlyRefusals(pageProblems, url);
  });

  test('in a small container the question stays inside it with both answers in reach, in Turkish', async ({ page, request, hasTouch, pageProblems }) => {
    await page.evaluate(() => {
      const editor = document.getElementById('editor')!;
      editor.style.width = '320px';
      editor.style.height = '200px';
    });
    const url = await conflictedWorkbook(page, request, { locale: 'tr' });
    await page.keyboard.press('ControlOrMeta+s');
    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await expect(confirm).toHaveAccessibleName('Dosya sunucuda değişmiş');
    await expect(confirm).toHaveAccessibleDescription(QUESTION_TR);
    await expect(page.locator(OVERWRITE)).toHaveText('Üzerine yaz');
    await expect(page.locator(CANCEL)).toHaveText('Vazgeç');
    const outer = (await page.locator('#editor').boundingBox())!;
    for (const part of [confirm, page.locator(OVERWRITE), page.locator(CANCEL)]) {
      const box = (await part.boundingBox())!;
      expect(box.x, 'left edge').toBeGreaterThanOrEqual(outer.x);
      expect(box.y, 'top edge').toBeGreaterThanOrEqual(outer.y);
      expect(box.x + box.width, 'right edge').toBeLessThanOrEqual(outer.x + outer.width + 0.5);
      expect(box.y + box.height, 'bottom edge').toBeLessThanOrEqual(outer.y + outer.height + 0.5);
    }
    // Large enough for a finger where there is one.
    const cancel = (await page.locator(CANCEL).boundingBox())!;
    expect(cancel.height).toBeGreaterThanOrEqual(hasTouch ? 40 : 32);
    await overwrite(page, hasTouch);
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => unsaved(page)).toBe(false);
    expect(await onServer(request, url)).toBe('mine|');
    onlyRefusals(pageProblems, url);
  });

  test('without a toolbar the question comes with Ctrl+S, and the page can answer it with saveNow({ overwrite: true })', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request, { toolbar: false });
    const puts = watchPuts(page, url);
    expect(await page.locator(SAVE).count()).toBe(0);
    await page.keyboard.press('ControlOrMeta+s');
    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await expect(page.locator(CANCEL)).toBeFocused();
    await saveLater(page, 'waiting');
    await saveLater(page, 'own', { overwrite: true });
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => answers(page)).toEqual(['waiting true', 'own true']);
    expect(puts, 'the page\'s overwrite is the one request after the refused save').toEqual(['if-match', 'unconditional']);
    expect(await onServer(request, url)).toBe('mine|');
    await expect(page.locator('#editor .bym-input')).toBeFocused();
    onlyRefusals(pageProblems, url);
  });

  test('a page that overwrites from inside onError is not asked as well', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request, { overwriteInOnError: true });
    const puts = watchPuts(page, url);
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => answers(page)).toEqual(['own true']);
    await expect(page.locator(CONFIRM)).toHaveCount(0);
    expect(puts).toEqual(['if-match', 'unconditional']);
    expect(await onServer(request, url)).toBe('mine|');
    onlyRefusals(pageProblems, url);
  });

  test('onConflict "fail": the conflict is only told, and nothing is asked', async ({ page, request, hasTouch, pageProblems }) => {
    const url = await conflictedWorkbook(page, request, { saveTo: { onConflict: 'fail' } });
    const save = page.locator(SAVE);
    await press(save, hasTouch);
    await expect(page.locator('#editor .bym-status-message')).toHaveText(/^The file was not saved: the file on the server is not the one you opened/);
    expect(await errors(page)).toEqual(['SaveError conflict 412 asked:false']);
    await expect(page.locator(CONFIRM)).toHaveCount(0);
    await expect(save).toBeEnabled();
    expect(await onServer(request, url)).toBe('|');
    expect(await mount(page, { file: url, saveTo: { onConflict: 'always' } })).toBe('TypeError: "saveTo.onConflict" must be "ask" or "fail"');
    onlyRefusals(pageProblems, url);
  });

  test('the question goes away when the view is made read-only and when it is destroyed', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    await page.keyboard.press('ControlOrMeta+s');
    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await saveLater(page, 'waiting');
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(true));
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => answers(page)).toEqual(['waiting false']);
    expect(await page.locator('#editor [inert]').count(), 'the editor takes input again').toBe(0);

    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    await page.evaluate(() => (window as any).__mounted.destroy());
    await expect(confirm).toHaveCount(0);
    await expect(page.locator('#editor')).toBeEmpty();
    expect(await onServer(request, url)).toBe('|');
    onlyRefusals(pageProblems, url);
  });

  test('document: the same question; what is typed while it is open is not in the document, and Overwrite saves it', async ({ page, request, hasTouch, pageProblems }) => {
    const url = address('docx');
    await seed(request, url, await createBlank('docx'));
    expect(await mount(page, { file: url, locale: 'tr' })).toBe('ok');
    await seed(request, url, await createBlank('docx'));
    await typeInDocument(page, 'benim');
    const save = page.locator(SAVE);
    await press(save, hasTouch);
    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await expect(confirm).toHaveAttribute('role', 'alertdialog');
    await expect(confirm).toHaveAccessibleDescription(QUESTION_TR);
    await expect(page.locator(CANCEL)).toBeFocused();
    await page.keyboard.type('zz');
    const box = (await page.locator('#editor').boundingBox())!;
    if (hasTouch) await page.touchscreen.tap(box.x + box.width / 2, box.y + box.height - 40);
    else await page.mouse.click(box.x + box.width / 2, box.y + box.height - 40);
    await page.keyboard.type('qq');
    await expect(confirm).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document)), 'nothing was typed into the document').toBe('benim');
    expect(await onServer(request, url)).toBe('');
    // The focus is back in the document: what is typed now is.
    await page.keyboard.type('!');
    expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document))).toContain('!');
    const text = await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__mounted.document));

    await press(save, hasTouch);
    await expect(confirm).toBeVisible();
    await overwrite(page, hasTouch);
    await expect(confirm).toHaveCount(0);
    await expect(save).toBeDisabled();
    expect(await onServer(request, url)).toBe(text);
    onlyRefusals(pageProblems, url);
  });

  test('keys that were on their way to the grid when the question came do not overwrite: arrows and Enter, pressed and held', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const release = await holdPuts(page, url);
    const confirm = page.locator(CONFIRM);
    await page.keyboard.press('ControlOrMeta+s');
    // The user goes on through the cells while the server has not answered.
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('ArrowDown');
    await expect(confirm).toHaveCount(0);
    const before = await page.evaluate(() => (window as any).__mounted.view.selection.active.row);
    release();
    await expect(confirm).toBeVisible();
    await page.keyboard.press('ArrowDown');
    await page.keyboard.press('Enter');
    await expect(confirm, 'Enter was pressed on Cancel').toHaveCount(0);
    expect(puts, 'nothing was sent over the file on the server').toEqual(['if-match']);
    expect(await onServer(request, url)).toBe('|');
    expect(await unsaved(page)).toBe(true);
    expect(await page.evaluate(() => (window as any).__mounted.view.selection.active.row), 'the keys moved no cell either').toBe(before);

    // An arrow that is held down goes back and forth between nothing, and Enter after it cancels.
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    for (const arrow of ['ArrowDown', 'ArrowLeft']) {
      for (let i = 0; i < 4; i++) await page.keyboard.down(arrow);
      await page.keyboard.up(arrow);
      await expect(page.locator(CANCEL), `${arrow} held down leaves the focus on Cancel`).toBeFocused();
    }
    // Enter held down on "overwrite", from before it answers: no answer, however long it is held.
    await page.locator(OVERWRITE).focus();
    for (let i = 0; i < 3; i++) await page.keyboard.down('Enter');
    await expect(page.locator(ARMED)).toBeVisible();
    for (let i = 0; i < 3; i++) await page.keyboard.down('Enter');
    await page.keyboard.up('Enter');
    await expect(confirm, 'a key that is held down is no answer').toBeVisible();
    expect(puts).toEqual(['if-match', 'if-match']);
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    expect(await onServer(request, url)).toBe('|');
    onlyRefusals(pageProblems, url);
  });

  test('"overwrite" answers nothing until a moment after the question came, whatever presses it; Cancel answers at once', async ({ page, request, hasTouch, pageProblems }) => {
    // The time of the page stands still when the test says so: what "a moment" is does not depend on how fast the machine is.
    await page.clock.install();
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const confirm = page.locator(CONFIRM);
    await stopTime(page);
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    const button = page.locator(OVERWRITE);
    await expect(button).toHaveAttribute('aria-disabled', 'true');
    expect(await button.evaluate((element) => (element as HTMLButtonElement).disabled), 'not disabled: it is there for a screen reader and for Tab').toBe(false);
    await pressAt(page, hasTouch, await centre(button));
    await pressAt(page, hasTouch, await centre(button));
    await button.focus();
    await expect(button).toBeFocused();
    await page.keyboard.press('Enter');
    await page.keyboard.press('Space');
    await page.clock.runFor(399);
    await pressAt(page, hasTouch, await centre(button));
    await expect(confirm, 'no press was an answer').toBeVisible();
    await expect(button).toHaveAttribute('aria-disabled', 'true');
    expect(puts).toEqual(['if-match']);
    await page.clock.runFor(1);
    await expect(page.locator(ARMED)).toBeVisible();
    await page.keyboard.press('Enter');
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => puts).toEqual(['if-match', 'unconditional']);
    await page.clock.resume();
    await expect.poll(() => unsaved(page)).toBe(false);
    expect(await onServer(request, url)).toBe('mine|');

    // Cancel does not wait.
    await seed(request, url, await createBlank('xlsx', { sheetName: 'Theirs again' }));
    await typeInGrid(page, 'more');
    await stopTime(page);
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    await press(page.locator(CANCEL), hasTouch);
    await expect(confirm).toHaveCount(0);
    await page.clock.resume();
    expect(await unsaved(page)).toBe(true);
    onlyRefusals(pageProblems, url);
  });

  test('a second press right after "save" is no answer, on "save" itself and where "overwrite" comes up, in a small container', async ({ page, request, hasTouch, pageProblems }) => {
    await page.clock.install();
    await page.evaluate(() => {
      const editor = document.getElementById('editor')!;
      editor.style.width = '320px';
      editor.style.height = '200px';
    });
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const confirm = page.locator(CONFIRM);
    await expect(page.locator(SAVE), 'the button knows of the change').toBeEnabled();
    const save = await centre(page.locator(SAVE));
    // Where the answer comes up.
    await pressAt(page, hasTouch, save);
    await expect(confirm).toBeVisible();
    const over = await centre(page.locator(OVERWRITE));
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);

    for (const second of [save, over]) {
      const sent = puts.length;
      // The server answers at once, and no time passes for the page between the presses: a double press.
      await stopTime(page);
      await pressAt(page, hasTouch, save);
      await page.locator(OVERWRITE).waitFor({ state: 'attached' });
      await pressAt(page, hasTouch, second);
      await pressAt(page, hasTouch, second);
      await expect(confirm, 'the second press was no answer').toBeVisible();
      await page.clock.runFor(400);
      await expect(page.locator(ARMED), 'the question is open, and from now on "overwrite" answers').toBeVisible();
      expect(puts.length, 'one save was sent').toBe(sent + 1);
      expect(puts.filter((put) => put !== 'if-match'), 'nothing was sent without its condition').toEqual([]);
      await page.keyboard.press('Escape');
      await expect(confirm).toHaveCount(0);
      await page.clock.resume();
    }
    expect(await onServer(request, url), 'the file on the server is the one someone else put there').toBe('|');
    onlyRefusals(pageProblems, url);
  });

  test('the question covers the editor whatever the page put into the container, and a root the page made static', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    // What an application adds to the container after mount(): a notice of its own in a corner.
    await page.evaluate(() => {
      const toast = document.createElement('div');
      toast.id = 'toast';
      toast.textContent = 'x';
      toast.style.cssText = 'position:absolute;right:0;bottom:0;width:12px;height:12px;overflow:hidden';
      const editor = document.getElementById('editor')!;
      editor.style.position = 'relative';
      editor.appendChild(toast);
    });
    const confirm = page.locator(CONFIRM);
    const places = () =>
      page.evaluate(() => {
        const backdrop = document.querySelector('.bym-confirm-backdrop') as HTMLElement;
        const root = document.querySelector('#editor > .bym-editor') as HTMLElement;
        const box = (element: Element) => {
          const rect = element.getBoundingClientRect();
          return [rect.left, rect.top, rect.width, rect.height].map(Math.round).join(' ');
        };
        return {
          parent: backdrop.parentElement === root ? 'the root of the editor' : (backdrop.parentElement as HTMLElement).id || (backdrop.parentElement as HTMLElement).className,
          covers: box(backdrop) === box(root),
          silenced: Array.from(root.children).filter((child) => child !== backdrop && !child.hasAttribute('inert')).length,
          toastInert: (document.getElementById('toast') as HTMLElement).closest('[inert]') !== null,
        };
      });
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    expect(await places()).toEqual({ parent: 'the root of the editor', covers: true, silenced: 0, toastInert: false });
    await expect(page.locator(CANCEL)).toBeFocused();
    expect(await topmostAt(page, await centre(page.locator(CANCEL))), 'the answer can be pressed').toBe('cancel');
    // The editor behind the question takes nothing.
    await page.keyboard.type('zz');
    const box = (await page.locator('#editor').boundingBox())!;
    await page.mouse.click(box.x + 30, box.y + box.height - 60);
    await page.keyboard.type('q');
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing), 'no cell is being edited').toBe(false);
    await expect(confirm).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    expect(await page.locator('#editor [inert]').count(), 'the editor takes input again').toBe(0);

    // A rule of the page that undoes the position the editor gave its root.
    await page.addStyleTag({ content: '#editor > .bym-editor { position: static !important; }' });
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    expect(await places()).toEqual({ parent: 'the root of the editor', covers: true, silenced: 0, toastInert: false });
    expect(await topmostAt(page, await centre(page.locator(CANCEL)))).toBe('cancel');
    const outer = (await page.locator('#editor').boundingBox())!;
    const inner = (await confirm.boundingBox())!;
    expect(inner.x >= outer.x && inner.y >= outer.y && inner.x + inner.width <= outer.x + outer.width + 0.5 && inner.y + inner.height <= outer.y + outer.height + 0.5, 'the question is inside the editor').toBe(true);
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    expect(await page.evaluate(() => (document.querySelector('#editor > .bym-editor') as HTMLElement).style.getPropertyPriority('position')), 'what the question set on the root is taken back').toBe('');
    expect(puts).toEqual(['if-match', 'if-match']);
    expect(await onServer(request, url)).toBe('|');
    onlyRefusals(pageProblems, url);
  });

  test('a view that is only read is not asked: Ctrl+S and saveNow() tell the conflict and nothing more', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(true));
    await page.evaluate(() => (window as any).__mounted.view.focus());
    await page.keyboard.press('ControlOrMeta+s');
    await expect.poll(() => errors(page), { message: 'the save was sent and refused' }).toEqual(['SaveError conflict 412 asked:false']);
    await saveLater(page, 'own');
    await expect.poll(() => answers(page)).toEqual(['own false']);
    await expect(page.locator(CONFIRM), 'nothing is asked over a view that is only read').toHaveCount(0);
    expect(await errors(page)).toEqual(['SaveError conflict 412 asked:false', 'SaveError conflict 412 asked:false']);
    expect(puts, 'the saves themselves went, as they always did').toEqual(['if-match', 'if-match']);
    expect(await page.evaluate(() => (window as any).__notices.filter((notice: string) => notice.startsWith('error: The file was not saved: the file on the server is not the one you opened')).length), 'the conflict is told').toBe(2);
    expect(await page.locator('#editor [inert]').count()).toBe(0);
    expect(await unsaved(page)).toBe(true);
    expect(await onServer(request, url)).toBe('|');

    // Editable again: the next conflict is asked about.
    await page.evaluate(() => (window as any).__mounted.editor.setReadOnly(false));
    await page.keyboard.press('ControlOrMeta+s');
    await expect(page.locator(CONFIRM)).toBeVisible();
    await page.keyboard.press('Escape');
    onlyRefusals(pageProblems, url);
  });

  test('a cell that is being typed when the question comes: Cancel leaves it open as it was, Overwrite writes it and saves it', async ({ page, request, hasTouch, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const release = await holdPuts(page, url);
    const confirm = page.locator(CONFIRM);
    const editing = () => page.evaluate(() => (window as any).__mounted.view.isEditing);
    await page.keyboard.press('ControlOrMeta+s');
    // Typed while the server has not answered: half a text, in a cell that is still open.
    await page.keyboard.type('half');
    expect(await editing()).toBe(true);
    release();
    await expect(confirm).toBeVisible();
    expect(await editing(), 'the cell stays open behind the question').toBe(true);
    await page.keyboard.type('XX');
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    expect(await editing(), 'Cancel leaves the cell as it was: open').toBe(true);
    expect(await onServer(request, url)).toBe('|');
    // The focus is back in the cell: what is typed now goes on where the text stopped.
    await page.keyboard.type(' text');
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    expect(puts).toEqual(['if-match', 'if-match']);
    await overwrite(page, hasTouch);
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => unsaved(page)).toBe(false);
    expect(puts).toEqual(['if-match', 'if-match', 'unconditional']);
    expect(await onServer(request, url), 'what was being typed was written to its cell and is in the file that went over the server').toBe('mine|half text');
    onlyRefusals(pageProblems, url);
  });

  test('a cell that is half typed when Overwrite is pressed is written first, as any save does', async ({ page, request, hasTouch, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const release = await holdPuts(page, url);
    const confirm = page.locator(CONFIRM);
    await page.keyboard.press('ControlOrMeta+s');
    await page.keyboard.type('half');
    release();
    await expect(confirm).toBeVisible();
    await overwrite(page, hasTouch);
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => unsaved(page)).toBe(false);
    expect(await page.evaluate(() => (window as any).__mounted.view.isEditing), 'the cell was closed by the save').toBe(false);
    expect(await onServer(request, url)).toBe('mine|half');
    onlyRefusals(pageProblems, url);
  });

  test('with the menu of a cell or a dialog of rules open the question is on top, and the editor is whole again after it', async ({ page, request, pageProblems }) => {
    const url = await conflictedWorkbook(page, request);
    const puts = watchPuts(page, url);
    const release = await holdPuts(page, url);
    const confirm = page.locator(CONFIRM);
    const menu = page.locator('#editor .bym-menu');
    await page.keyboard.press('ControlOrMeta+s');
    // The menu of a cell is opened while the server has not answered.
    const grid = (await page.locator('#editor .bym-canvas').boundingBox())!;
    await page.evaluate(
      ([x, y]) => document.querySelector('#editor .bym-canvas')!.dispatchEvent(new MouseEvent('contextmenu', { clientX: x, clientY: y, button: 2, bubbles: true, cancelable: true })),
      [grid.x + 80, grid.y + 60],
    );
    await expect(menu).toBeVisible();
    release();
    await expect(confirm).toBeVisible();
    await expect(page.locator(CANCEL)).toBeFocused();
    for (const answer of [OVERWRITE, CANCEL]) expect(await topmostAt(page, await centre(page.locator(answer))), 'the answers are on top of the menu').toBe(answer === OVERWRITE ? 'overwrite' : 'cancel');
    expect(await menu.evaluate((element) => element.closest('[inert]') !== null), 'the menu takes no press meanwhile').toBe(true);
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    await expect(menu, 'Escape answered the question, and only that').toBeVisible();
    expect(await page.locator('#editor [inert]').count()).toBe(0);
    await page.keyboard.press('Escape');
    await expect(menu).toHaveCount(0);
    expect(puts).toEqual(['if-match']);

    // A dialog of rules that is opened while the server has not answered, with nothing changed in it.
    release(true);
    await page.keyboard.press('ControlOrMeta+s');
    expect(await page.evaluate(() => (window as any).__mounted.view.openValidationDialog())).toBe(true);
    const dialog = page.locator('#editor [data-bym-dialog="validation"]');
    await expect(dialog).toBeVisible();
    release();
    await expect(confirm).toBeVisible();
    for (const answer of [OVERWRITE, CANCEL]) expect(await topmostAt(page, await centre(page.locator(answer))), 'the answers are on top of the dialog').toBe(answer === OVERWRITE ? 'overwrite' : 'cancel');
    expect(await dialog.evaluate((element) => element.closest('[inert]') !== null), 'the dialog takes no press meanwhile').toBe(true);
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);
    await expect(dialog, 'the dialog is as it was').toBeVisible();
    expect(await page.locator('#editor [inert]').count()).toBe(0);
    expect(puts).toEqual(['if-match', 'if-match']);

    // A save closes a dialog in which nothing was changed, as it always did; the question comes over the grid then.
    await page.keyboard.press('ControlOrMeta+s');
    await expect(confirm).toBeVisible();
    await expect(dialog).toHaveCount(0);
    await expect(page.locator(ARMED)).toBeVisible();
    await page.locator(OVERWRITE).click();
    await expect(confirm).toHaveCount(0);
    await expect.poll(() => unsaved(page)).toBe(false);
    expect(puts).toEqual(['if-match', 'if-match', 'if-match', 'unconditional']);
    expect(await onServer(request, url)).toBe('mine|');
    expect(await page.locator('#editor [inert]').count(), 'nothing of the editor is left silenced').toBe(0);
    await typeInGrid(page, 'after');
    expect(await unsaved(page), 'the grid takes what is typed').toBe(true);
    onlyRefusals(pageProblems, url);
  });
});

test.describe('the bridge page: saveResult with conflict asks before it overwrites', () => {
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
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto('/dist/bymoffice.webview.html');
  });

  for (const toolbar of [true, false]) {
    test(`a workbook ${toolbar ? 'with' : 'without'} its bars: conflict asks, Cancel sends nothing, Overwrite hands the file over with overwrite: true`, async ({ page, hasTouch }) => {
      await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), name: 'liste.xlsx', save: true, toolbar, locale: 'tr' });
      await waitFor(page, 'loaded');
      const input = page.locator('#editor .bym-input');
      await input.focus();
      await page.keyboard.type('12');
      await page.keyboard.press('Enter');
      await page.keyboard.press('ControlOrMeta+s');
      const saved = await waitFor(page, 'saved');
      expect('overwrite' in saved).toBe(false);
      await send(page, { type: 'saveResult', saveId: saved.saveId, ok: false, conflict: true, message: 'Ada <b>değiştirdi</b>' });

      const confirm = page.locator(CONFIRM);
      await expect(confirm).toBeVisible();
      await expect(confirm).toHaveAttribute('role', 'alertdialog');
      await expect(confirm).toHaveAccessibleName('Dosya sunucuda değişmiş');
      await expect(page.locator(CANCEL)).toBeFocused();
      // What the application said is text, never markup.
      await expect(page.locator(`${CONFIRM} .bym-confirm-detail`)).toHaveText('Ada <b>değiştirdi</b>');
      expect(await page.locator(`${CONFIRM} b`).count()).toBe(0);
      await page.keyboard.type('9');
      // The key asks for nothing more while the question is open.
      await page.keyboard.press('ControlOrMeta+s');
      const before = (await messages(page)).length;
      await press(page.locator(CANCEL), hasTouch);
      await expect(confirm).toHaveCount(0);
      await expect(input).toBeFocused();
      expect((await messages(page)).length, 'Cancel sends no message').toBe(before);
      expect((await messages(page)).filter((message) => message.type === 'saved').length).toBe(1);

      await page.keyboard.press('ControlOrMeta+s');
      const second = await waitFor(page, 'saved', 2);
      expect('overwrite' in second, 'the next save decides nothing for the application').toBe(false);
      await send(page, { type: 'saveResult', saveId: second.saveId, ok: false, conflict: true });
      await expect(confirm).toBeVisible();
      expect(await page.locator(`${CONFIRM} .bym-confirm-detail`).count(), 'nothing was said this time').toBe(0);
      await overwrite(page, hasTouch);
      await expect(confirm).toHaveCount(0);
      const third = await waitFor(page, 'saved', 3);
      expect([third.by, third.overwrite, third.format, third.name]).toEqual(['user', true, 'xlsx', 'liste.xlsx']);
      expect(new Set([saved.saveId, second.saveId, third.saveId]).size, 'every save has an id of its own').toBe(3);
      expect(await firstCells(third.data), 'what was typed while the question was open is not in the file').toBe('12|');
      await send(page, { type: 'saveResult', saveId: third.saveId, ok: true });
      const states = async () => (await messages(page)).filter((message) => message.type === 'saveState').map((message) => `${message.dirty} ${message.saving}`);
      await expect.poll(async () => (await states()).pop()).toBe('false false');
      expect((await messages(page)).filter((message) => message.type === 'error')).toEqual([]);
    });
  }

  test('a document: conflict asks; setReadOnly and a new load take the question away', async ({ page, hasTouch }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('docx')), save: true, toolbar: true });
    await waitFor(page, 'loaded');
    await page.locator('#editor .bym-doc-page').first().click({ position: { x: 150, y: 150 } });
    await page.keyboard.type('Mine');
    await page.keyboard.press('ControlOrMeta+s');
    const saved = await waitFor(page, 'saved');
    await send(page, { type: 'saveResult', saveId: saved.saveId, ok: false, conflict: true });
    const confirm = page.locator(CONFIRM);
    await expect(confirm).toBeVisible();
    await expect(page.locator(CANCEL)).toBeFocused();
    await send(page, { type: 'setReadOnly', readOnly: true });
    await expect(confirm).toHaveCount(0);
    await send(page, { type: 'setReadOnly', readOnly: false });

    await page.keyboard.press('ControlOrMeta+s');
    const second = await waitFor(page, 'saved', 2);
    await send(page, { type: 'saveResult', saveId: second.saveId, ok: false, conflict: true });
    await expect(confirm).toBeVisible();
    await overwrite(page, hasTouch);
    const third = await waitFor(page, 'saved', 3);
    expect([third.overwrite, third.format]).toEqual([true, 'docx']);
    const file = await open(new Uint8Array(Buffer.from(third.data, 'base64')));
    expect(documentText(await file.document!())).toBe('Mine');
    await send(page, { type: 'saveResult', saveId: third.saveId, ok: false, conflict: true });
    await expect(page.locator('#editor .bym-toolbar [data-command="save"]')).toBeEnabled();
    await expect(confirm, 'an overwrite that is refused is not asked about').toHaveCount(0);

    await page.locator('#editor .bym-doc-page').first().click({ position: { x: 150, y: 150 } });
    await page.keyboard.press('ControlOrMeta+s');
    const fourth = await waitFor(page, 'saved', 4);
    await send(page, { type: 'saveResult', saveId: fourth.saveId, ok: false, conflict: true });
    await expect(confirm).toBeVisible();
    await send(page, { type: 'load', id: 'm', data: base64(await createBlank('docx')), save: true, toolbar: true });
    await waitFor(page, 'loaded', 2);
    await expect(confirm).toHaveCount(0);
  });

  test('an application that answers at once: a second press after "save" is no answer, in a small container', async ({ page, hasTouch }) => {
    await page.clock.install();
    await page.evaluate(() => {
      const editor = document.getElementById('editor')!;
      editor.style.cssText += ';position:absolute;left:0;top:0;right:auto;bottom:auto;width:320px;height:200px';
      // The application says "conflict" the moment it is handed a file.
      const w = window as any;
      const host = w.BymOfficeHost;
      w.BymOfficeHost = {
        postMessage: (json: string) => {
          host.postMessage(json);
          const message = JSON.parse(json);
          if (message.type === 'saved' && message.by === 'user') {
            void Promise.resolve().then(() => w.BymOfficeBridge.receive(JSON.stringify({ type: 'saveResult', saveId: message.saveId, ok: false, conflict: true })));
          }
        },
      };
    });
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), save: true, toolbar: true });
    await waitFor(page, 'loaded');
    await page.locator('#editor .bym-input').focus();
    await page.keyboard.type('12');
    await page.keyboard.press('Enter');
    const confirm = page.locator(CONFIRM);
    const userSaves = async () => (await messages(page)).filter((message) => message.type === 'saved' && message.by === 'user');
    await expect(page.locator(SAVE), 'the button knows of the change').toBeEnabled();
    const save = await centre(page.locator(SAVE));
    await pressAt(page, hasTouch, save);
    await expect(confirm).toBeVisible();
    const outer = (await page.locator('#editor').boundingBox())!;
    const inner = (await confirm.boundingBox())!;
    expect(inner.x >= outer.x && inner.y >= outer.y && inner.x + inner.width <= outer.x + outer.width + 0.5 && inner.y + inner.height <= outer.y + outer.height + 0.5, 'the question is inside the container').toBe(true);
    const over = await centre(page.locator(OVERWRITE));
    await page.keyboard.press('Escape');
    await expect(confirm).toHaveCount(0);

    for (const second of [save, over]) {
      const sent = (await userSaves()).length;
      // No time passes for the page between the presses: a double press.
      await stopTime(page);
      await pressAt(page, hasTouch, save);
      await page.locator(OVERWRITE).waitFor({ state: 'attached' });
      await pressAt(page, hasTouch, second);
      await pressAt(page, hasTouch, second);
      await expect(confirm, 'the second press was no answer').toBeVisible();
      await page.clock.runFor(400);
      await expect(page.locator(ARMED), 'the question is open, and from now on "overwrite" answers').toBeVisible();
      expect((await userSaves()).filter((message) => 'overwrite' in message), 'no file was handed over to overwrite').toEqual([]);
      expect((await userSaves()).length, 'and none after the one that was refused').toBe(sent + 1);
      await page.keyboard.press('Escape');
      await expect(confirm).toHaveCount(0);
      await page.clock.resume();
    }
  });

  test('a file that is only read is not asked about: the conflict is told', async ({ page }) => {
    await send(page, { type: 'load', id: 'l', data: base64(await createBlank('xlsx')), save: true, toolbar: true });
    await waitFor(page, 'loaded');
    const input = page.locator('#editor .bym-input');
    await input.focus();
    await page.keyboard.type('12');
    await page.keyboard.press('Enter');
    await send(page, { type: 'setReadOnly', id: 'r', readOnly: true });
    await waitFor(page, 'done');
    await input.focus();
    await page.keyboard.press('ControlOrMeta+s');
    const saved = await waitFor(page, 'saved');
    await send(page, { type: 'saveResult', id: 's', saveId: saved.saveId, ok: false, conflict: true });
    await waitFor(page, 'done', 2);
    await expect(page.locator('#editor .bym-status-message')).toHaveText(/^The file was not saved: the file on the server is not the one you opened/);
    await expect(page.locator(CONFIRM), 'nothing is asked over a view that is only read').toHaveCount(0);
    expect(await page.locator('#editor [inert]').count()).toBe(0);
    expect((await messages(page)).filter((message) => message.type === 'saved').length).toBe(1);
  });
});
