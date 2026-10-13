import type { APIRequestContext, BrowserContext, Page } from '@playwright/test';
import { createBlank, getCell, open } from '../src/index';
import { expect, test } from './helpers';

// MountOptions.collab: two pages of one browser open the same file and edit it together. The
// room is held by the first page (a relay of the page, reached over a BroadcastChannel) or by
// the relay of the test server (a WebSocket). No screenshots: what is drawn is read from the
// canvas with getImageData().

const PAGE = '/e2e/pages/mount.html';
const RED = '#D93025';
const BLUE = '#1A73E8';

/** A name no other test uses: for the room and for the file. */
function unique(): string {
  const info = test.info();
  return `${info.project.name}-${info.testId}-${info.repeatEachIndex}`.replace(/[^A-Za-z0-9_-]/g, '_');
}

async function seed(request: APIRequestContext, url: string, bytes: Uint8Array): Promise<void> {
  const answer = await request.put(url, { data: Buffer.from(bytes), headers: { 'content-type': 'application/octet-stream' } });
  expect(answer.status(), `seeding ${url}`).toBe(204);
}

interface Spec {
  file: string | number[];
  room: string;
  /** "host": this page holds the room; "tab": it reaches the room of another page; "socket": the relay of the server. */
  via: 'host' | 'tab' | 'socket';
  name?: string;
  color?: string;
  readOnly?: boolean;
  saveTo?: boolean;
  autoSave?: boolean;
}

function mount(page: Page, spec: Spec): Promise<string> {
  return page.evaluate(async (given) => {
    const w = window as any;
    w.__notices = [];
    w.__errors = [];
    w.__states = [];
    const api = w.BymOffice;
    let collab: any;
    if (given.via === 'socket') collab = { url: `ws://${location.host}/e2e/collab/${given.room}` };
    else {
      if (given.via === 'host') {
        w.__relay = api.collab.memoryServer();
        w.__relay.serve(given.room);
      }
      w.__transport = api.collab.broadcastTransport(given.room);
      collab = { transport: w.__transport };
    }
    collab.user = { name: given.name, color: given.color };
    collab.onState = (state: string) => w.__states.push(state);
    const options: any = {
      file: typeof given.file === 'string' ? given.file : new Uint8Array(given.file),
      locale: 'en',
      collab,
      onNotice: (notice: { kind: string; message: string }) => w.__notices.push(`${notice.kind}: ${notice.message}`),
      onError: (error: { code?: string; status?: number; name?: string }) => w.__errors.push(`${error.name} ${error.code} ${error.status}`),
    };
    if (given.readOnly) options.readOnly = true;
    if (given.saveTo) options.saveTo = true;
    if (given.autoSave) options.autoSave = { delayMs: 300, maxWaitMs: 2000 };
    try {
      w.__mounted = await api.mount(document.getElementById('editor'), options);
      return 'ok';
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  }, spec);
}

const online = (page: Page) => expect.poll(() => page.evaluate(() => (window as any).__mounted.collab.state), { message: 'the page is online' }).toBe('online');
const peers = (page: Page): Promise<number> => page.evaluate(() => (window as any).__mounted.collab.peers.length);
const cell = (page: Page, row: number, col: number, sheet = 0): Promise<unknown> =>
  page.evaluate(([s, r, c]) => {
    const w = window as any;
    const found = w.BymOffice.getCell(w.__mounted.workbook.sheets[s], r, c);
    return found ? found.value : null;
  }, [sheet, row, col]);
const type = (page: Page, row: number, col: number, text: string): Promise<void> => page.evaluate(([r, c, t]) => (window as any).__mounted.view.setCellInput(r, c, t), [row, col, text] as [number, number, string]);
const notices = (page: Page): Promise<string[]> => page.evaluate(() => (window as any).__notices);
const settled = (page: Page) => expect.poll(() => page.evaluate(() => (window as any).__mounted.collab.pending), { message: 'nothing waits for the relay' }).toBe(0);

/** How many pixels of the canvas have exactly this colour, after a forced redraw. */
function pixelsOf(page: Page, colour: string): Promise<number> {
  return page.evaluate((hex) => {
    const view = (window as any).__mounted.view;
    view.render();
    const canvas = document.querySelector('#editor .bym-canvas') as HTMLCanvasElement;
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    let n = 0;
    for (let i = 0; i < data.length; i += 4) if (data[i] === r && data[i + 1] === g && data[i + 2] === b) n++;
    return n;
  }, colour);
}

/** Two pages on one file and one room; the first one holds the room unless the relay of the server does. */
async function pair(context: BrowserContext, page: Page, request: APIRequestContext, options: { socket?: boolean; saveTo?: boolean; autoSave?: boolean; second?: Partial<Spec> } = {}): Promise<{ a: Page; b: Page; url: string; room: string }> {
  const room = unique();
  const url = `/e2e/saved/${room}.xlsx`;
  await seed(request, url, await createBlank('xlsx', { sheetName: 'Data' }));
  const b = await context.newPage();
  const problems: string[] = [];
  b.on('console', (message) => {
    if (message.type() === 'error') problems.push(message.text());
  });
  b.on('pageerror', (error) => problems.push(`${error.name}: ${error.message}`));
  (b as any).__problems = problems;
  await page.goto(PAGE);
  await b.goto(PAGE);
  expect(await mount(page, { file: url, room, via: options.socket ? 'socket' : 'host', name: 'Ada', color: RED, saveTo: options.saveTo, autoSave: options.autoSave })).toBe('ok');
  expect(await mount(b, { file: url, room, via: options.socket ? 'socket' : 'tab', name: 'Ben', color: BLUE, saveTo: options.saveTo, autoSave: options.autoSave, ...options.second })).toBe('ok');
  await online(page);
  await online(b);
  await expect.poll(() => peers(page), { message: 'the first page sees the second' }).toBe(1);
  await expect.poll(() => peers(b), { message: 'the second page sees the first' }).toBe(1);
  return { a: page, b, url, room };
}

const clean = (b: Page) => expect((b as any).__problems, 'console errors and uncaught exceptions of the second page').toEqual([]);

for (const socket of [false, true]) {
  test(`what one page types, the other one shows (${socket ? 'the relay of the server' : 'a room in the first page'})`, async ({ page, context, request }) => {
    const { a, b } = await pair(context, page, request, { socket });
    await type(a, 0, 0, '42');
    await expect.poll(() => cell(b, 0, 0), { message: 'the second page has the value' }).toBe(42);
    await type(b, 1, 0, '=A1*2');
    await expect.poll(() => cell(a, 1, 0), { message: 'the first page calculated the formula of the second' }).toBe(84);
    await settled(a);
    await settled(b);
    expect(await notices(a), 'nothing was said to the first page').toEqual([]);
    clean(b);
  });
}

test('keys typed into the grid of one page arrive in the other', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await a.evaluate(() => (window as any).__mounted.view.focus());
  await expect(a.locator('#editor .bym-input')).toBeFocused();
  await a.keyboard.type('hello');
  await a.keyboard.press('Enter');
  await expect.poll(() => cell(b, 0, 0), { message: 'the second page has the text' }).toBe('hello');
  clean(b);
});

test('two pages that write at once both keep what the other wrote; one cell written by both is the same on both', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await Promise.all([type(a, 0, 0, 'from a'), type(b, 0, 1, 'from b'), type(a, 2, 2, 'a again'), type(b, 2, 2, 'b again')]);
  await settled(a);
  await settled(b);
  await expect.poll(() => cell(a, 0, 1)).toBe('from b');
  await expect.poll(() => cell(b, 0, 0)).toBe('from a');
  await expect.poll(async () => (await cell(a, 2, 2)) === (await cell(b, 2, 2)), { message: 'the cell both wrote is the same on both pages' }).toBe(true);
  expect(String(await cell(a, 2, 2)), 'one of the two texts').toMatch(/^[ab] again$/);
  clean(b);
});

test('a change in a row someone else deleted meanwhile is dropped, and its author is told', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await type(a, 3, 0, 'row four');
  await expect.poll(() => cell(b, 3, 0)).toBe('row four');
  // The second page goes offline, writes into the row, and comes back after the row is gone.
  await b.evaluate(() => (window as any).__transport.suspend());
  await expect.poll(() => b.evaluate(() => (window as any).__mounted.collab.state)).toBe('offline');
  await type(b, 3, 1, 'too late');
  await a.evaluate(() => (window as any).__mounted.view.deleteRows(3, 1));
  await settled(a);
  await b.evaluate(() => (window as any).__transport.resume());
  await online(b);
  await settled(b);
  await expect.poll(() => notices(b), { message: 'the second page was told' }).toContainEqual(expect.stringMatching(/^error: One of your changes could not be applied/));
  expect(await cell(b, 3, 1), 'nothing of it is left').toBe(null);
  expect(await cell(a, 3, 1), 'and nothing arrived').toBe(null);
  clean(b);
});

test('the selection of the other page is drawn in its colour with its name, and its sheet has a dot on the tab', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await b.evaluate(() => (window as any).__mounted.view.select({ top: 4, left: 2, bottom: 4, right: 2 }));
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.collab.peers[0].presence), { message: 'the first page knows where the second is' }).toEqual({ sheet: 0, row: 4, col: 2, range: { top: 4, left: 2, bottom: 4, right: 2 } });
  // A frame of two pixels around a cell and a label: hundreds of pixels of exactly that colour.
  await expect.poll(() => pixelsOf(a, BLUE), { message: 'the frame and the label in the colour of the second page' }).toBeGreaterThan(200);
  expect(await pixelsOf(b, BLUE), 'nobody draws their own frame').toBe(0);
  // What the others select arrives in its own time (presence is gathered before it is drawn): asked for until it is there.
  await expect.poll(() => pixelsOf(b, RED), { message: 'the second page draws the first one' }).toBeGreaterThan(200);
  await expect(a.locator('#editor .bym-collab-peer')).toHaveText('Ben');
  await expect(a.locator('#editor .bym-tab .bym-tab-peer')).toHaveCount(1);
  await expect(a.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'online');
  // The page that leaves takes its frame along.
  await b.evaluate(() => (window as any).__mounted.destroy());
  await expect.poll(() => peers(a), { message: 'the first page is alone again' }).toBe(0);
  await expect.poll(() => pixelsOf(a, BLUE)).toBe(0);
  await expect(a.locator('#editor .bym-tab .bym-tab-peer')).toHaveCount(0);
  clean(b);
});

test('undo takes back the own change and leaves the change of the other page', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await type(a, 0, 0, 'mine');
  await expect.poll(() => cell(b, 0, 0)).toBe('mine');
  await type(b, 0, 1, 'theirs');
  await expect.poll(() => cell(a, 0, 1)).toBe('theirs');
  expect(await a.evaluate(() => (window as any).__mounted.view.undo()), 'the first page had a step to undo').toBe(true);
  await expect.poll(() => cell(b, 0, 0), { message: 'the change of the first page is gone on both' }).toBe(null);
  expect(await cell(a, 0, 0)).toBe(null);
  expect(await cell(a, 0, 1), 'the change of the second page stays').toBe('theirs');
  expect(await cell(b, 0, 1)).toBe('theirs');
  expect(await a.evaluate(() => (window as any).__mounted.view.undo()), 'there is nothing of its own left to undo').toBe(false);
  expect(await a.evaluate(() => (window as any).__mounted.view.redo()), 'and the step can be made again').toBe(true);
  await expect.poll(() => cell(b, 0, 0)).toBe('mine');
  clean(b);
});

test('when one page saves, the other one has nothing left to save and its next save is no conflict', async ({ page, context, request }) => {
  const { a, b, url, room } = await pair(context, page, request, { saveTo: true });
  await type(a, 0, 0, 'first');
  await expect.poll(() => cell(b, 0, 0)).toBe('first');
  await expect.poll(() => b.evaluate(() => (window as any).__mounted.hasUnsavedChanges), { message: 'what the first page wrote is unsaved on the second as well' }).toBe(true);
  expect(await a.evaluate(() => (window as any).__mounted.saveNow()), 'the first page saved').toBe(true);
  await expect.poll(() => b.evaluate(() => (window as any).__mounted.hasUnsavedChanges), { message: 'the second page heard of the save' }).toBe(false);
  await type(b, 1, 0, 'second');
  expect(await b.evaluate(() => (window as any).__mounted.saveNow()), 'the second page saved over the file the first one saved').toBe(true);
  expect(await b.evaluate(() => (window as any).__errors), 'no conflict').toEqual([]);
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.hasUnsavedChanges), { message: 'and the first page heard of that save' }).toBe(false);
  const file = await open(new Uint8Array(await (await request.get(url)).body()));
  const sheet = (await file.workbook()).sheets[0];
  expect([getCell(sheet, 0, 0)?.value, getCell(sheet, 1, 0)?.value], 'the server holds both changes').toEqual(['first', 'second']);
  // A page that opens the file now joins at the step the file holds.
  const c = await context.newPage();
  await c.goto(PAGE);
  expect(await mount(c, { file: url, room, via: 'tab', name: 'Cem' })).toBe('ok');
  await online(c);
  expect(await cell(c, 1, 0), 'the page that came later read the saved file').toBe('second');
  await type(a, 2, 0, 'third');
  await expect.poll(() => cell(c, 2, 0), { message: 'and gets what is changed from then on' }).toBe('third');
  clean(b);
});

test('a page that is offline says so, goes on editing, and catches up when it is back', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await b.evaluate(() => (window as any).__transport.suspend());
  await expect(b.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'offline');
  await expect(b.locator('#editor .bym-collab-state')).toHaveText('Offline: changes are sent when the connection is back');
  await type(b, 0, 0, 'offline one');
  await type(b, 1, 0, 'offline two');
  await type(a, 0, 1, 'meanwhile');
  expect(await b.evaluate(() => (window as any).__mounted.collab.pending), 'two changes wait').toBe(2);
  expect(await cell(a, 0, 0), 'nothing arrived yet').toBe(null);
  await b.evaluate(() => (window as any).__transport.resume());
  await expect(b.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'online');
  await settled(b);
  await expect.poll(() => cell(a, 1, 0), { message: 'the first page got what was written offline' }).toBe('offline two');
  expect(await cell(a, 0, 0)).toBe('offline one');
  await expect.poll(() => cell(b, 0, 1), { message: 'and the second page what it missed' }).toBe('meanwhile');
  clean(b);
});

test('a page whose file is not the one of the room loads it again in place and joins', async ({ page, context, request }) => {
  const { a, b, url, room } = await pair(context, page, request, { saveTo: true });
  await type(a, 0, 0, 'saved');
  expect(await a.evaluate(() => (window as any).__mounted.saveNow()), 'the first page saved').toBe(true);
  await type(a, 1, 0, 'after the save');
  await settled(a);
  // A third page is first handed a file nobody in the room knows (a stale copy from somewhere), then the real one.
  const c = await context.newPage();
  const other = Buffer.from(await createBlank('xlsx', { sheetName: 'Stale' }));
  let asked = 0;
  await c.route(`**${url}`, (route) => (route.request().method() === 'GET' && asked++ === 0 ? route.fulfill({ body: other, contentType: 'application/octet-stream' }) : route.continue()));
  await c.goto(PAGE);
  expect(await mount(c, { file: url, room, via: 'tab', name: 'Cem', saveTo: true })).toBe('ok');
  expect(await c.evaluate(() => (window as any).__mounted.workbook.sheets[0].name), 'the stale file was opened first').toBe('Stale');
  await online(c);
  expect(asked, 'the file was fetched a second time').toBe(2);
  expect(await c.evaluate(() => (window as any).__mounted.workbook.sheets[0].name), 'the result of mount() holds the file of the room now').toBe('Data');
  await expect.poll(() => cell(c, 1, 0), { message: 'with what was changed after it was saved' }).toBe('after the save');
  await expect(c.locator('#editor .bym-editor')).toHaveCount(1);
  // The room was asked to save because someone could not come in: the one who saves for the others did, and everybody has nothing left to save.
  await expect.poll(() => c.evaluate(() => (window as any).__mounted.hasUnsavedChanges), { message: 'the file that is stored holds everything, and the page that came in knows' }).toBe(false);
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.hasUnsavedChanges), { message: 'also for the page that saved it' }).toBe(false);
  // The page that was loaded again is a participant like any other.
  await type(c, 2, 0, 'from the third');
  await expect.poll(() => cell(b, 2, 0), { message: 'what it writes arrives' }).toBe('from the third');
  await type(b, 3, 0, 'to the third');
  await expect.poll(() => cell(c, 3, 0), { message: 'and what the others write reaches it' }).toBe('to the third');
  expect(await c.evaluate(() => (window as any).__mounted.saveNow()), 'and it saves without a conflict').toBe(true);
  expect(await c.evaluate(() => (window as any).__errors)).toEqual([]);
  clean(b);
});

test('a read-only page follows and changes nothing', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { second: { readOnly: true } });
  await type(a, 0, 0, 'shown');
  await expect.poll(() => cell(b, 0, 0), { message: 'the page that only reads has the value' }).toBe('shown');
  await type(b, 1, 0, 'never');
  expect(await cell(b, 1, 0), 'a read-only view writes nothing').toBe(null);
  expect(await b.evaluate(() => (window as any).__mounted.collab.pending)).toBe(0);
  expect(await a.evaluate(() => (window as any).__mounted.collab.peers[0].readOnly), 'the others know that it only reads').toBe(true);
  clean(b);
});

test('a document is not edited together: it is shown, and a notice says so', async ({ page }) => {
  await page.goto(PAGE);
  const bytes = await createBlank('docx');
  expect(await mount(page, { file: Array.from(bytes), room: unique(), via: 'host' })).toBe('ok');
  expect(await notices(page)).toContain('error: Editing together is not available for documents yet.');
  expect(await page.evaluate(() => (window as any).__mounted.collab), 'no session').toBe(undefined);
});

test('"collab" that is no relay is refused before anything is fetched', async ({ page }) => {
  await page.goto(PAGE);
  const answer = await page.evaluate(async () => {
    try {
      await (window as any).BymOffice.mount(document.getElementById('editor'), { file: '/e2e/saved/never-asked-for.xlsx', collab: 'https://example.com/room' });
      return 'mounted';
    } catch (error) {
      return `${(error as Error).name}: ${(error as Error).message}`;
    }
  });
  expect(answer).toBe('TypeError: "collab" is the address of a relay: "ws://…" or "wss://…"');
});

// The bridge page: `load` with `collab`, and what it tells the application with `collabState`.
test('the bridge page joins a room for a load with "collab" and tells the application who is there', async ({ page, context }) => {
  const room = unique();
  const bytes = Buffer.from(await createBlank('xlsx')).toString('base64');
  const second = await context.newPage();
  for (const one of [page, second]) {
    await one.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await one.goto('/dist/bymoffice.webview.html');
    await one.waitForFunction(() => (window as any).__messages.some((message: { type: string }) => message.type === 'ready'));
  }
  const send = (to: Page, message: unknown) => to.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
  const last = (of: Page, type: string) => of.evaluate((wanted) => (window as any).__messages.filter((message: { type: string }) => message.type === wanted).pop(), type);
  const address = `ws://${new URL(test.info().project.use.baseURL!).host}/e2e/collab/${room}`;
  await send(page, { type: 'load', data: bytes, name: 'book.xlsx', locale: 'en', collab: { url: address, user: { name: 'Ada', color: RED } } });
  await send(second, { type: 'load', data: bytes, name: 'book.xlsx', locale: 'en', collab: address });
  await expect.poll(async () => ((await last(page, 'collabState')) || {}).state, { message: 'the first page is online' }).toBe('online');
  await expect.poll(async () => ((await last(second, 'collabState')) || {}).peers, { message: 'the second page is told of the first' }).toEqual([{ id: expect.any(String), name: 'Ada', color: RED, readOnly: false }]);
  // The first page types; the second one hears of a change and writes it into the file it hands out.
  await page.locator('.bym-input').focus();
  await page.keyboard.type('77');
  await page.keyboard.press('Enter');
  await expect.poll(async () => ((await last(second, 'changed')) || {}).sheets, { message: 'the second page changed' }).toEqual([0]);
  await send(second, { type: 'save', id: 's1' });
  await expect.poll(async () => typeof ((await last(second, 'saved')) || {}).data, { message: 'the second page handed its file out' }).toBe('string');
  const saved = await last(second, 'saved');
  const sheet = (await (await open(new Uint8Array(Buffer.from(saved.data, 'base64')))).workbook()).sheets[0];
  expect(getCell(sheet, 0, 0)?.value, 'what the first page typed is in the file of the second').toBe(77);
  // A load that is no relay is refused as a bad message.
  await send(second, { type: 'load', id: 'bad', data: bytes, collab: 'https://example.com/x' });
  await expect.poll(async () => ((await last(second, 'error')) || {}).code).toBe('bad-message');
});

test('the demo page opens two editors that edit together, in a room of the page and on the relay of the server', async ({ page }) => {
  await page.goto('/demo/index.html');
  for (const socket of [false, true]) {
    await page.locator('#collab-socket').setChecked(socket);
    await page.locator('#collab-start').click();
    await expect(page.locator('#collab-state')).toContainText('2. düzenleyici: online', { timeout: 15000 });
    await expect(page.locator('#collab-a .bym-collab-peer')).toHaveText('Deniz');
    await expect(page.locator('#collab-b .bym-collab-peer')).toHaveText('Ayşe');
    await page.locator('#collab-a .bym-input').focus();
    await page.keyboard.type('12');
    await page.keyboard.press('Enter');
    await expect(page.locator('#collab-state')).toContainText('2. düzenleyici: online, adım 1, bekleyen 0, öbürleri 1');
    await page.locator('#collab-cut').click();
    await expect(page.locator('#collab-b .bym-collab')).toHaveAttribute('data-state', 'offline');
    await page.locator('#collab-cut').click();
    await expect(page.locator('#collab-b .bym-collab')).toHaveAttribute('data-state', 'online');
  }
});

test('what one page is typing stays open while the other one writes, moves with its row, and a formula that cells moved under is given up and said', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request);
  await a.evaluate(() => (window as any).__mounted.view.select({ top: 5, left: 1, bottom: 5, right: 1 }));
  await a.evaluate(() => (window as any).__mounted.view.focus());
  await expect(a.locator('#editor .bym-input')).toBeFocused();
  await a.keyboard.type('half');
  // Someone else writes a cell: nothing of the edit is touched.
  await type(b, 0, 0, 'elsewhere');
  await expect.poll(() => cell(a, 0, 0)).toBe('elsewhere');
  expect(await a.evaluate(() => (window as any).__mounted.view.editText), 'the text that is being typed is still there').toBe('half');
  // Someone else inserts two rows above: the cell that is typed in is two rows further down.
  await b.evaluate(() => (window as any).__mounted.view.insertRows(1, 2));
  await expect.poll(() => a.evaluate(() => (window as any).__mounted.collab.rev), { message: 'the rows arrived' }).toBe(2);
  expect(await a.evaluate(() => (window as any).__mounted.view.editText), 'still being typed').toBe('half');
  await a.keyboard.type(' done');
  await a.keyboard.press('Enter');
  await expect.poll(() => cell(b, 7, 1), { message: 'what was typed is in the row that moved' }).toBe('half done');
  expect(await cell(b, 5, 1), 'and not where the row was').toBe(null);
  // A formula that is being typed while rows move is not applied to cells it did not mean.
  await a.evaluate(() => (window as any).__mounted.view.select({ top: 9, left: 3, bottom: 9, right: 3 }));
  await a.keyboard.type('=B8');
  await b.evaluate(() => (window as any).__mounted.view.insertRows(0, 1));
  await expect.poll(() => notices(a), { message: 'the first page was told' }).toContainEqual(expect.stringMatching(/^error: Someone else moved cells or renamed a sheet while you were typing a formula/));
  expect(await a.evaluate(() => (window as any).__mounted.view.editText), 'the edit is over').toBe('');
  await settled(a);
  expect(await cell(b, 10, 3), 'nothing was written').toBe(null);
  clean(b);
});

test('a room that started anew detaches a page with unsaved work: the work stays, it says so, and after a save it joins again', async ({ page, context, request }) => {
  const { a, b } = await pair(context, page, request, { saveTo: true });
  await type(b, 0, 0, 'unsaved work');
  await settled(b);
  // The page goes offline; the relay forgets the room (as one that was started anew does); the page comes back.
  await b.evaluate(() => (window as any).__transport.suspend());
  await expect(b.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'offline');
  await a.evaluate(() => (window as any).__relay.roomsByKey.clear());
  await b.evaluate(() => (window as any).__transport.resume());
  await expect(b.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'detached');
  await expect(b.locator('#editor .bym-collab-state')).toHaveText('Detached: your changes are here only. Save the file');
  expect(await cell(b, 0, 0), 'the work is still there').toBe('unsaved work');
  expect(await notices(b)).toContainEqual(expect.stringMatching(/^error: The file was changed elsewhere, or the room was started anew.*Nothing was thrown away/));
  await type(b, 1, 0, 'still editing');
  expect(await b.evaluate(() => (window as any).__mounted.collab.pending), 'what is changed now is told to nobody').toBe(0);
  // Saved: the file holds everything, and the page joins a room with it.
  expect(await b.evaluate(() => (window as any).__mounted.saveNow()), 'the page saved').toBe(true);
  await expect(b.locator('#editor .bym-collab')).toHaveAttribute('data-state', 'online');
  expect(await cell(b, 1, 0)).toBe('still editing');
  clean(b);
});

test('while the file is loaded again nothing can be typed into the workbook that is about to go', async ({ page, context, request }) => {
  const { a, b, url } = await pair(context, page, request, { saveTo: true });
  let release: () => void = () => undefined;
  const held = new Promise<void>((resolve) => (release = resolve));
  let asked = 0;
  await b.route(`**${url}`, async (route) => {
    if (route.request().method() !== 'GET') return route.continue();
    asked++;
    await held;
    await route.continue();
  });
  await b.evaluate(() => (window as any).__mounted.collab.reload());
  await expect.poll(() => asked, { message: 'the file is asked for' }).toBe(1);
  expect(await b.evaluate(() => (window as any).__mounted.view.readOnly), 'the view is closed to editing meanwhile').toBe(true);
  expect(await notices(b)).toContainEqual(expect.stringMatching(/loaded again; it cannot be edited until that is done/));
  await b.evaluate(() => {
    try {
      (window as any).__mounted.view.setCellInput(3, 0, 'typed into a workbook that goes');
    } catch {
      // Refusing with an error is as good as refusing without one.
    }
  });
  expect(await cell(b, 3, 0), 'nothing was written').toBe(null);
  release();
  await expect.poll(() => b.evaluate(() => (window as any).__mounted.view.readOnly), { message: 'the file was loaded and the view can be edited again' }).toBe(false);
  await online(b);
  await type(b, 3, 0, 'typed afterwards');
  await expect.poll(() => cell(a, 3, 0), { message: 'what is typed afterwards reaches the others' }).toBe('typed afterwards');
  clean(b);
});

test('auto-save on both pages of a room: one page saves, nobody is told of a conflict, and the stored file holds what was typed', async ({ page, context, request }) => {
  const { a, b, url } = await pair(context, page, request, { saveTo: true, autoSave: true });
  const puts: string[] = [];
  for (const [name, each] of [['a', a], ['b', b]] as [string, Page][]) {
    each.on('response', (response) => {
      if (response.request().method() === 'PUT' && response.url().endsWith(url)) puts.push(`${name} ${response.status()}`);
    });
  }
  const stored = async (): Promise<string> => {
    const answer = await request.get(url);
    const book = await (await open(new Uint8Array(await answer.body()))).workbook();
    return JSON.stringify(Array.from(book.sheets[0].rows.entries()).map(([row, value]) => [row, value.cells.map((cell) => (cell ? cell.value : null))]));
  };
  // Typed on one page, then on the other: saved each time, by one page.
  for (const [who, text, row] of [[b, 'typed by the second', 2], [a, 'typed by the first', 3], [b, 'typed by the second again', 4]] as [Page, string, number][]) {
    const before = puts.length;
    await type(who, row, 0, text);
    await expect.poll(stored, { message: `"${text}" is in the stored file` }).toContain(text);
    await expect.poll(() => Promise.all([a, b].map((each) => each.evaluate(() => (window as any).__mounted.hasUnsavedChanges))), { message: 'both pages have nothing left to save' }).toEqual([false, false]);
    expect(puts.slice(before).filter((put) => put.endsWith(' 412')), 'no save met another').toEqual([]);
    expect(new Set(puts.slice(before).map((put) => put[0])).size, `one page saved: ${puts.slice(before).join()}`).toBe(1);
  }
  expect(new Set(puts.map((put) => put[0])).size, `the same page every time: ${puts.join()}`).toBe(1);
  for (const each of [a, b]) {
    expect(await each.evaluate(() => (window as any).__errors), 'no error of a save').toEqual([]);
    expect((await notices(each)).filter((notice) => notice.startsWith('error')), 'nothing is said of a conflict').toEqual([]);
  }
  clean(b);
});

test('what is being typed when the file is loaded again is said back, not lost without a word', async ({ page, context, request }) => {
  const { b } = await pair(context, page, request, { saveTo: true });
  await b.evaluate(() => {
    const view = (window as any).__mounted.view;
    view.select({ top: 5, left: 0, bottom: 5, right: 0 });
    view.startEditing('half a sentence');
  });
  expect(await b.evaluate(() => (window as any).__mounted.view.hasUnappliedText()), 'a cell is being typed in').toBe(true);
  await b.evaluate(() => (window as any).__mounted.collab.reload());
  await expect.poll(() => notices(b), { message: 'the text is said back' }).toContainEqual(expect.stringMatching(/what you were typing was not applied\. You typed: half a sentence/));
  await online(b);
  expect(await cell(b, 5, 0), 'and it is in no cell').toBe(null);
  clean(b);
});
