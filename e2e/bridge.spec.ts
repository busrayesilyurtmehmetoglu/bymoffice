import { readFile } from 'node:fs/promises';
import type { FrameLocator, Page } from '@playwright/test';
import { ZipArchive, documentText, getCell, open } from '../src/index';
import { buildXlsx, sampleWorkbook, worksheetXml } from '../test/fixtures';
import { cat, number, workbookStream, xlsFile } from '../test/xls-fixtures';
import { docx, sampleDocx } from './fixtures/docx';
import { textFile, truncatedXlsx, xls } from './fixtures/files';
import { expect, expectResponsive, test } from './helpers';

const PAGE = '/dist/bymoffice.webview.html';
/** The application that embeds the bridge page in an iframe: another origin than the one the page is served from. */
const APP = 'http://app.example';
const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');

/** One sheet with a number in A1 and nothing else, so that typing lands in A1. */
const smallWorkbook = () =>
  buildXlsx({ sheets: [{ name: 'Sayfa1', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') }] });

type Mode = 'channel' | 'iframe';

/** The application's side of the bridge, for both ways of hosting the page. */
interface App {
  /** Sends a message; anything that is not a string is sent as JSON text. */
  send(message: unknown): Promise<void>;
  /** Everything the page sent so far. The `data` of a `saved` message is replaced by its length. */
  messages(): Promise<any[]>;
  /** Waits until `count` messages of a type have arrived and returns the last of them. */
  waitFor(type: string, count?: number): Promise<any>;
  /** The base64 `data` of the n-th `saved` message. */
  savedData(index?: number): Promise<string>;
  /** Where the editor's elements are. */
  editor: Page | FrameLocator;
  /** URLs the browser requested. */
  requests: string[];
}

interface FrameSetup {
  /** Origin the embedding page is served from. Default: APP. */
  parent?: string;
  /** What the bridge page lists as origins that may embed it. Default: the parent. */
  allow?: string;
}

/**
 * Serves the embedding page from origins that do not exist (so that it is
 * cross-origin to the bridge page), and the bridge page with its list of
 * allowed parents filled in when the address carries "?allow=".
 */
async function routes(page: Page): Promise<void> {
  const hostPage = await readFile(new URL('./pages/bridge-host.html', import.meta.url));
  await page.route(/^http:\/\/[a-z]+\.example\/host\.html/, (route) => route.fulfill({ body: hostPage, contentType: 'text/html; charset=utf-8' }));
  const bridgePage = await readFile(new URL('../dist/bymoffice.webview.html', import.meta.url), 'utf8');
  await page.route(/\/dist\/bymoffice\.webview\.html\?allow=/, (route) => {
    const allow = decodeURIComponent(new URL(route.request().url()).searchParams.get('allow') ?? '');
    const marker = '<meta name="bymoffice-parent-origins" content="" />';
    if (!bridgePage.includes(marker)) throw new Error('the bridge page has no empty list of parent origins');
    return route.fulfill({ body: bridgePage.replace(marker, () => `<meta name="bymoffice-parent-origins" content="${allow}" />`), contentType: 'text/html; charset=utf-8' });
  });
}

async function start(page: Page, mode: Mode, setup: FrameSetup = {}): Promise<App> {
  const requests: string[] = [];
  page.on('request', (request) => requests.push(request.url()));
  if (mode === 'channel') {
    // What a WebView's JavaScript channel looks like to the page.
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto(PAGE);
  } else {
    await routes(page);
    const parent = setup.parent ?? APP;
    const allow = setup.allow ?? parent;
    const origin = new URL(test.info().project.use.baseURL!).origin;
    // Always through the route (an empty list leaves the page as built): Chromium does not let a
    // page of a public origin frame one from a loopback address that it has to fetch itself.
    const src = `${origin}${PAGE}?allow=${encodeURIComponent(allow)}`;
    // The real server has the embedding page under its path in the repository.
    await page.goto(`${parent}${parent === origin ? '/e2e/pages/bridge-host.html' : '/host.html'}?src=${encodeURIComponent(src)}`);
    await expect(page.frameLocator('#bridge').locator('#status')).toBeVisible();
  }
  return app(page, mode, requests);
}

function app(page: Page, mode: Mode, requests: string[]): App {
  const messages = () =>
    page.evaluate(() =>
      ((window as any).__messages as unknown[]).map((entry) => {
        if (typeof entry !== 'string') return { notAString: typeof entry };
        const message = JSON.parse(entry);
        if (message.type === 'saved') message.data = message.data.length;
        return message;
      }),
    );
  return {
    requests,
    messages,
    editor: mode === 'channel' ? page : page.frameLocator('#bridge'),
    async send(message) {
      const text = typeof message === 'string' ? message : JSON.stringify(message);
      if (mode === 'channel') await page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), text);
      else await page.evaluate((json) => (window as any).__send(json), text);
    },
    async waitFor(type, count = 1) {
      let found: any[] = [];
      await expect
        .poll(async () => (found = (await messages()).filter((message) => message.type === type)).length, { message: `${count} x ${type}` })
        .toBeGreaterThanOrEqual(count);
      return found[count - 1];
    },
    savedData: (index = 0) =>
      page.evaluate((n) => {
        const saved = ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)).filter((message) => message.type === 'saved');
        return saved[n].data as string;
      }, index),
  };
}

/** The page only ever talks to the server it came from (and, framed, the browser fetched the embedding page). */
function expectNoOutsideRequests(requests: string[]): void {
  const origin = new URL(test.info().project.use.baseURL!).origin;
  const outside = requests.filter((url) => !url.startsWith(`${origin}/`) && !url.startsWith(`${APP}/host.html`) && !url.startsWith('data:') && !url.startsWith('blob:'));
  expect(outside, 'requests to other origins').toEqual([]);
}

async function typeIntoGrid(page: Page, editor: Page | FrameLocator, text: string): Promise<void> {
  const input = editor.locator('.bym-input');
  await input.focus();
  await expect(input).toBeFocused();
  await page.keyboard.type(text);
  await page.keyboard.press('Enter');
}

const p = (text: string) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** Puts the caret at the start of the document that is shown and types there. */
async function typeIntoDocument(page: Page, editor: Page | FrameLocator, text: string, isMobile: boolean): Promise<void> {
  const sheet = editor.locator('#editor .bym-doc-page[data-page="1"]');
  await expect(sheet).toBeVisible();
  // Near the top left corner of the page: the caret goes to the start of the first line.
  if (isMobile) await sheet.tap({ position: { x: 6, y: 6 } });
  else await sheet.click({ position: { x: 6, y: 6 } });
  const input = editor.locator('#editor .bym-doc-input');
  await expect(input).toBeFocused();
  await page.keyboard.type(text);
}

/** The format and the text of the document in the n-th `saved` message. */
async function textOfSaved(bridge: App, index: number): Promise<{ format: string; text: string }> {
  const file = await open(new Uint8Array(Buffer.from(await bridge.savedData(index), 'base64')));
  return { format: file.format, text: documentText(await file.document!()) };
}

/** Gives messages that should not come the time to come, then checks that none did. */
async function expectSilence(page: Page, bridge: App, expected: unknown[] = []): Promise<void> {
  await expectResponsive(page);
  await page.waitForTimeout(600);
  expect(await bridge.messages()).toEqual(expected);
}

for (const mode of ['channel', 'iframe'] as Mode[]) {
  test.describe(`bridge page (${mode})`, () => {
    test('ready, load, edit, changed, selection, save: the edit is in the saved file', async ({ page }) => {
      const bridge = await start(page, mode);
      const ready = await bridge.waitFor('ready');
      expect(ready).toEqual({ type: 'ready', version: expect.stringMatching(/^\d+\.\d+\.\d+/) });
      await expect(bridge.editor.locator('#status')).toHaveText('No file is open');

      await bridge.send({ type: 'load', id: 'l1', data: base64(await smallWorkbook()), name: 'küçük.xlsx', locale: 'tr' });
      expect(await bridge.waitFor('loaded')).toEqual({
        type: 'loaded',
        id: 'l1',
        format: 'xlsx',
        sheets: [{ name: 'Sayfa1', state: 'visible', kind: 'worksheet' }],
        activeSheet: 0,
      });
      await expect(bridge.editor.locator('#editor canvas')).toBeVisible();
      await expect(bridge.editor.locator('#editor .bym-tabs button')).toHaveText(['Sayfa1']);
      await expect(bridge.editor.locator('#status')).toBeHidden();
      // The editor fills the page.
      const size = await bridge.editor.locator('#editor').evaluate((element) => {
        const box = element.getBoundingClientRect();
        return { width: box.width, height: box.height, pageWidth: window.innerWidth, pageHeight: window.innerHeight };
      });
      expect(size.width).toBe(size.pageWidth);
      expect(size.height).toBe(size.pageHeight);

      await typeIntoGrid(page, bridge.editor, '42');
      expect(await bridge.waitFor('changed')).toEqual({ type: 'changed', sheets: [0], canUndo: true, canRedo: false });
      // Enter moved the selection down.
      expect(await bridge.waitFor('selection')).toEqual({ type: 'selection', sheet: 0, range: 'A2', active: 'A2' });

      await bridge.send({ type: 'save', id: 's1' });
      expect(await bridge.waitFor('saved')).toEqual({ type: 'saved', id: 's1', data: expect.any(Number), name: 'küçük.xlsx' });
      const saved = new Uint8Array(Buffer.from(await bridge.savedData(), 'base64'));
      const workbook = await (await open(saved)).workbook();
      expect(getCell(workbook.sheets[0], 0, 0)).toMatchObject({ kind: 'number', value: 42 });

      // Every message was a string, and none of them an error.
      const all = await bridge.messages();
      expect(all.filter((message) => message.notAString || message.type === 'error')).toEqual([]);
      expectNoOutsideRequests(bridge.requests);
    });

    test('undo, redo, setReadOnly, setLocale and close, confirmed with done when they carry an id', async ({ page }) => {
      const bridge = await start(page, mode);
      await bridge.send({ type: 'load', data: base64(await smallWorkbook()), locale: 'tr' });
      await bridge.waitFor('loaded');
      await expect(bridge.editor.locator('.bym-view')).toHaveAttribute('aria-label', 'Hesap tablosu');
      await expect(bridge.editor.locator('html')).toHaveAttribute('lang', 'tr');

      await typeIntoGrid(page, bridge.editor, '7');
      await bridge.waitFor('changed');
      await bridge.send({ type: 'undo', id: 'u' });
      expect(await bridge.waitFor('changed', 2)).toEqual({ type: 'changed', sheets: [0], canUndo: false, canRedo: true });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'u', command: 'undo' });
      await bridge.send({ type: 'redo' });
      expect(await bridge.waitFor('changed', 3)).toEqual({ type: 'changed', sheets: [0], canUndo: true, canRedo: false });

      await bridge.send({ type: 'setLocale', locale: 'en', id: 'loc' });
      expect(await bridge.waitFor('done', 2)).toEqual({ type: 'done', id: 'loc', command: 'setLocale' });
      await expect(bridge.editor.locator('.bym-view')).toHaveAttribute('aria-label', 'Spreadsheet');
      await expect(bridge.editor.locator('html')).toHaveAttribute('lang', 'en');

      await bridge.send({ type: 'setReadOnly', readOnly: true, id: 'ro' });
      expect(await bridge.waitFor('done', 3)).toEqual({ type: 'done', id: 'ro', command: 'setReadOnly' });
      await expect(bridge.editor.locator('.bym-input')).toHaveJSProperty('readOnly', true);
      await typeIntoGrid(page, bridge.editor, '9');
      // A read-only view does not undo either.
      await bridge.send({ type: 'undo' });
      await bridge.send({ type: 'save', id: 'ro' });
      await bridge.waitFor('saved');
      expect((await bridge.messages()).filter((message) => message.type === 'changed')).toHaveLength(3);
      const workbook = await (await open(new Uint8Array(Buffer.from(await bridge.savedData(), 'base64')))).workbook();
      expect(getCell(workbook.sheets[0], 0, 0)).toMatchObject({ value: 7 });

      await bridge.send({ type: 'setReadOnly', readOnly: false });
      await expect(bridge.editor.locator('.bym-input')).toHaveJSProperty('readOnly', false);

      await bridge.send({ type: 'close', id: 'c' });
      expect(await bridge.waitFor('done', 4)).toEqual({ type: 'done', id: 'c', command: 'close' });
      await expect(bridge.editor.locator('#editor')).toBeEmpty();
      await expect(bridge.editor.locator('#status')).toHaveText('No file is open');
      await bridge.send({ type: 'save', id: 'gone' });
      expect(await bridge.waitFor('error')).toEqual({ type: 'error', id: 'gone', code: 'not-loaded', message: expect.any(String) });
      // Commands without an id were not confirmed.
      expect((await bridge.messages()).filter((message) => message.type === 'done')).toHaveLength(4);
      expectNoOutsideRequests(bridge.requests);
    });

    test('load can show a hidden sheet; changing the tab is reported; hello repeats ready', async ({ page }) => {
      const bridge = await start(page, mode);
      await bridge.send({ type: 'load', data: base64(await sampleWorkbook()), sheet: 1 });
      expect(await bridge.waitFor('loaded')).toMatchObject({
        sheets: [
          { name: 'Veri', state: 'visible', kind: 'worksheet' },
          { name: 'Gizli', state: 'hidden', kind: 'worksheet' },
          { name: 'Grafik', state: 'visible', kind: 'chartsheet' },
        ],
        activeSheet: 1,
      });
      // The hidden sheet is shown but has no tab.
      await expect(bridge.editor.locator('#editor .bym-tabs button')).toHaveText(['Veri', 'Grafik']);
      await bridge.editor.locator('#editor .bym-tabs button', { hasText: 'Grafik' }).click();
      expect(await bridge.waitFor('sheet')).toEqual({ type: 'sheet', index: 2, name: 'Grafik' });
      await bridge.editor.locator('#editor .bym-tabs button', { hasText: 'Veri' }).click();
      expect(await bridge.waitFor('sheet', 2)).toEqual({ type: 'sheet', index: 0, name: 'Veri' });

      await bridge.send({ type: 'hello' });
      expect(await bridge.waitFor('ready', 2)).toEqual(await bridge.waitFor('ready'));
    });

    test('broken messages and files are answered with errors; a failed load keeps the open file', async ({ page }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      const cases: [unknown, string][] = [
        ['{not json', 'bad-message'],
        [{ type: 'explode', id: 'c1' }, 'bad-message'],
        [{ id: 'c2' }, 'bad-message'],
        [{ type: 'save', id: 'c3' }, 'not-loaded'],
        [{ type: 'undo', id: 'c4' }, 'not-loaded'],
        [{ type: 'load', id: 'c5' }, 'bad-message'],
        [{ type: 'load', id: 'c6', data: '!!not base64!!' }, 'bad-message'],
        [{ type: 'load', id: 'c7', url: 'javascript:alert(1)' }, 'bad-message'],
        [{ type: 'load', id: 'c8', data: base64(await truncatedXlsx()) }, 'corrupt'],
        [{ type: 'load', id: 'c9', data: base64(textFile()) }, 'unsupported'],
        // A compound file whose "Workbook" stream is not a workbook: an xls that cannot be read.
        [{ type: 'load', id: 'c10', data: base64(xls()) }, 'corrupt'],
        [{ type: 'load', id: 'c11', data: '' }, 'unsupported'],
        [{ type: 'load', id: 'c12', data: base64(await smallWorkbook()), sheet: 5 }, 'bad-message'],
        [{ type: 'setReadOnly', id: 'c13' }, 'bad-message'],
        [{ type: 'load', id: 'c14', url: 'https://files.example/a.xlsx' }, 'bad-message'],
        [{ type: 'load', id: 'c15', url: '//files.example/a.xlsx' }, 'bad-message'],
      ];
      // One at a time: a load sent while another is under way would cancel it.
      for (let i = 0; i < cases.length; i++) {
        await bridge.send(cases[i][0]);
        await bridge.waitFor('error', i + 1);
      }
      const errors = (await bridge.messages()).filter((message) => message.type === 'error');
      expect(errors.map((error) => error.code)).toEqual(cases.map(([, code]) => code));
      expect(errors.map((error) => error.id)).toEqual([undefined, ...cases.slice(1).map((_, index) => `c${index + 1}`)]);
      for (const error of errors) expect(typeof error.message).toBe('string');
      await expect(bridge.editor.locator('#editor')).toBeEmpty();
      await expectResponsive(page);

      // A file that could not be opened is said so on the page in the page's own words,
      // in the language of the load; the engine's message goes to the application only.
      await bridge.send({ type: 'load', data: base64(textFile()), locale: 'en' });
      expect((await bridge.waitFor('error', cases.length + 1)).message).toBe('Not a Word or Excel file');
      await expect(bridge.editor.locator('#status')).toHaveText('This type of file is not supported');
      await bridge.send({ type: 'load', data: base64(await truncatedXlsx()) });
      await bridge.waitFor('error', cases.length + 2);
      await expect(bridge.editor.locator('#status')).toHaveText('The file is damaged and could not be opened');

      // The bridge still works afterwards.
      await bridge.send({ type: 'load', id: 'ok', data: base64(await smallWorkbook()), name: 'kept.xlsx', locale: 'tr' });
      expect((await bridge.waitFor('loaded')).id).toBe('ok');
      await expect(bridge.editor.locator('#editor canvas')).toBeVisible();
      await typeIntoGrid(page, bridge.editor, '5');
      await bridge.waitFor('changed');

      // Loads that fail now leave that file open, on screen, with its unsaved edit.
      await bridge.send({ type: 'load', id: 'f1', data: base64(await truncatedXlsx()) });
      expect(await bridge.waitFor('error', cases.length + 3)).toMatchObject({ id: 'f1', code: 'corrupt' });
      await bridge.send({ type: 'load', id: 'f2', data: base64(textFile()) });
      expect(await bridge.waitFor('error', cases.length + 4)).toMatchObject({ id: 'f2', code: 'unsupported' });
      await expect(bridge.editor.locator('#editor canvas')).toBeVisible();
      await expect(bridge.editor.locator('#status')).toBeHidden();
      await bridge.send({ type: 'save', id: 'kept' });
      expect(await bridge.waitFor('saved')).toMatchObject({ id: 'kept', name: 'kept.xlsx' });
      const workbook = await (await open(new Uint8Array(Buffer.from(await bridge.savedData(), 'base64')))).workbook();
      expect(getCell(workbook.sheets[0], 0, 0)).toMatchObject({ value: 5 });

      // The words of the page follow the language of the last successful load.
      await bridge.send({ type: 'close' });
      await expect(bridge.editor.locator('#status')).toHaveText('Açık dosya yok');
      await expect(bridge.editor.locator('html')).toHaveAttribute('lang', 'tr');
      expectNoOutsideRequests(bridge.requests);
    });

    test('a 5 MB file goes in and comes back out as base64', async ({ page }) => {
      test.setTimeout(120_000);
      const zip = ZipArchive.read(await sampleWorkbook());
      const noise = new Uint8Array(5 * 1024 * 1024);
      let seed = 0x9e3779b9;
      for (let i = 0; i < noise.length; i++) {
        seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
        noise[i] = seed >>> 24;
      }
      zip.set('xl/media/noise.bin', noise);
      const file = await zip.write();
      expect(file.length).toBeGreaterThan(5 * 1024 * 1024);

      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      const message = JSON.stringify({ type: 'load', data: base64(file) });
      let started = Date.now();
      await bridge.send(message);
      await bridge.waitFor('loaded');
      const loadMs = Date.now() - started;

      started = Date.now();
      await bridge.send({ type: 'save' });
      const saved = await bridge.waitFor('saved');
      const saveMs = Date.now() - started;
      console.log(`bridge 5 MB (${mode}, ${test.info().project.name}): load ${loadMs} ms, save ${saveMs} ms`);
      expect(saved.data).toBeGreaterThan((5 * 1024 * 1024 * 4) / 3);
      expect(loadMs).toBeLessThan(15_000);
      expect(saveMs).toBeLessThan(15_000);

      const back = ZipArchive.read(new Uint8Array(Buffer.from(await bridge.savedData(), 'base64')));
      expect(Buffer.compare(Buffer.from(await back.read('xl/media/noise.bin')), Buffer.from(noise))).toBe(0);
      await expectResponsive(page);
    });

    test('a docx is shown read-only as pages: loaded, pictures, link, save, commands', async ({ page, isMobile }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      const input = await sampleDocx();
      // Read-only: in a document that is edited a plain click puts the caret, and a link needs Ctrl.
      await bridge.send({ type: 'load', id: 'd1', data: base64(input), name: 'belge.docx', locale: 'tr', readOnly: true });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 'd1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
      await expect(bridge.editor.locator('#status')).toBeHidden();
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveAttribute('aria-label', 'Belge');
      await expect(bridge.editor.locator('#editor .bym-doc-pages')).toHaveText('Sayfa 1 / 3');
      await expect(bridge.editor.locator('#editor .bym-view')).toHaveCount(0);

      // The page's policy lets the picture of the document be decoded from its package: red pixels appear.
      const red = () =>
        bridge.editor.locator('#editor .bym-doc-page[data-page="1"] canvas').evaluate((element) => {
          const canvas = element as HTMLCanvasElement;
          const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
          let count = 0;
          for (let i = 0; i < data.length; i += 4) if (data[i] > 180 && data[i + 1] < 60 && data[i + 2] < 60) count++;
          return count;
        });
      await expect.poll(red).toBeGreaterThan(1000);

      // A link is handed to the application; the page itself goes nowhere and opens nothing.
      const address = page.url();
      const link = bridge.editor.locator('#editor .bym-doc-link');
      await expect(link).toHaveCount(1);
      if (isMobile) await link.tap();
      else await link.click();
      expect(await bridge.waitFor('link')).toEqual({ type: 'link', url: 'https://example.com/docs?a=1' });
      expect(page.url()).toBe(address);
      expect(page.context().pages().length).toBe(1);

      // Saving returns the file as it came.
      await bridge.send({ type: 'save', id: 's1' });
      expect(await bridge.waitFor('saved')).toEqual({ type: 'saved', id: 's1', data: expect.any(Number), name: 'belge.docx' });
      expect(Buffer.from(await bridge.savedData(), 'base64').equals(Buffer.from(input))).toBe(true);

      await bridge.send({ type: 'setLocale', id: 'c1', locale: 'en' });
      await bridge.send({ type: 'setReadOnly', id: 'c2', readOnly: false });
      await bridge.send({ type: 'undo', id: 'c3' });
      await bridge.send({ type: 'redo', id: 'c4' });
      expect(await bridge.waitFor('done', 4)).toEqual({ type: 'done', id: 'c4', command: 'redo' });
      await expect(bridge.editor.locator('#editor .bym-doc-pages')).toHaveText('Page 1 / 3');
      await expect(bridge.editor.locator('html')).toHaveAttribute('lang', 'en');

      // A workbook takes the place of the document, and a document that of the workbook.
      await bridge.send({ type: 'load', id: 'x1', data: base64(await smallWorkbook()) });
      await bridge.waitFor('loaded', 2);
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor .bym-view')).toHaveCount(1);
      await bridge.send({ type: 'load', id: 'd2', data: base64(input) });
      await bridge.waitFor('loaded', 3);
      await expect(bridge.editor.locator('#editor .bym-view')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveCount(1);

      // A damaged document is refused and the open one stays.
      await bridge.send({ type: 'load', id: 'bad', data: base64(input.slice(0, input.length >> 1)) });
      expect(await bridge.waitFor('error')).toMatchObject({ type: 'error', id: 'bad', code: 'corrupt' });
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveCount(1);

      await bridge.send({ type: 'close', id: 'z' });
      expect(await bridge.waitFor('done', 5)).toEqual({ type: 'done', id: 'z', command: 'close' });
      await expect(bridge.editor.locator('#editor')).toBeEmpty();
      await expect(bridge.editor.locator('#status')).toHaveText('No file is open');

      // A document that is not edited reports no changes, and none reports a selection.
      const all = await bridge.messages();
      expect(all.filter((message) => message.notAString || message.type === 'changed' || message.type === 'selection')).toEqual([]);
      expect(all.filter((message) => message.type === 'error').length).toBe(1);
      expectNoOutsideRequests(bridge.requests);
    });

    test('the tracked changes of a docx: their number in loaded, accept all, tracking on and off, and what is typed while it is on', async ({ page, isMobile }) => {
      const who = (id: number, author: string) => ` w:id="${id}" w:author="${author}" w:date="2026-10-01T09:00:00Z"`;
      const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
      const body = `<w:p>${run('One ')}<w:ins${who(1, 'Ann')}>${run('two ')}</w:ins><w:del${who(2, 'Bob')}><w:r><w:delText xml:space="preserve">three </w:delText></w:r></w:del>${run('four.')}</w:p>${p('Second paragraph.')}`;
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      await bridge.send({ type: 'load', id: 'd1', data: base64(await docx(body)), author: 'Me' });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 'd1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, revisions: 2 });

      await bridge.send({ type: 'rejectAllRevisions', id: 'r' });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'r', command: 'rejectAllRevisions', revisions: 0, trackChanges: false });
      // The count is told before the change it comes with.
      const told = (await bridge.messages()).filter((message) => message.type === 'revisions' || message.type === 'changed');
      expect(told).toEqual([
        { type: 'revisions', count: 0, trackChanges: false },
        { type: 'changed', sheets: [], canUndo: true, canRedo: false },
      ]);
      await bridge.send({ type: 'save', id: 's1' });
      await bridge.waitFor('saved');
      expect((await textOfSaved(bridge, 0)).text).toBe('One three four.\nSecond paragraph.');

      // Undone, the two changes are back; accepted, the text is the other one.
      await bridge.send({ type: 'undo', id: 'u' });
      expect(await bridge.waitFor('revisions', 2)).toEqual({ type: 'revisions', count: 2, trackChanges: false });
      await bridge.send({ type: 'acceptAllRevisions', id: 'a' });
      expect(await bridge.waitFor('done', 3)).toEqual({ type: 'done', id: 'a', command: 'acceptAllRevisions', revisions: 0, trackChanges: false });
      await bridge.send({ type: 'save', id: 's2' });
      await bridge.waitFor('saved', 2);
      expect((await textOfSaved(bridge, 1)).text).toBe('One two four.\nSecond paragraph.');

      // Tracking on: what is typed is a tracked change, and the application hears of it.
      await bridge.send({ type: 'setTrackChanges', id: 't', on: true });
      expect(await bridge.waitFor('done', 4)).toEqual({ type: 'done', id: 't', command: 'setTrackChanges', revisions: 0, trackChanges: true });
      await typeIntoDocument(page, bridge.editor, 'X', isMobile);
      expect(await bridge.waitFor('revisions', 5)).toEqual({ type: 'revisions', count: 1, trackChanges: true });
      await bridge.send({ type: 'setTrackChanges', id: 't2', on: false });
      expect(await bridge.waitFor('done', 5)).toEqual({ type: 'done', id: 't2', command: 'setTrackChanges', revisions: 1, trackChanges: false });
      await bridge.send({ type: 'save', id: 's3' });
      await bridge.waitFor('saved', 3);
      const saved = Buffer.from(await bridge.savedData(2), 'base64');
      const xml = new TextDecoder().decode(await ZipArchive.read(new Uint8Array(saved)).read('word/document.xml'));
      expect(xml).toMatch(/<w:ins w:id="\d+" w:author="Me" w:date="[^"]+"><w:r><w:t>X<\/w:t><\/w:r><\/w:ins>/);
      expect(xml.split('<w:ins ').length).toBe(2);

      // A workbook has no tracked changes to resolve.
      await bridge.send({ type: 'load', id: 'x', data: base64(await smallWorkbook()) });
      await bridge.waitFor('loaded', 2);
      await bridge.send({ type: 'acceptAllRevisions', id: 'n' });
      expect(await bridge.waitFor('error')).toMatchObject({ type: 'error', id: 'n', code: 'no-document' });
      const all = await bridge.messages();
      expect(all.filter((message) => message.notAString || message.type === 'error')).toHaveLength(1);
      expectNoOutsideRequests(bridge.requests);
    });

    test('a docx is edited: typing sends changed, undo and redo work, the saved file holds the text', async ({ page, isMobile }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      await bridge.send({ type: 'load', id: 'd1', data: base64(await docx(p('First paragraph.') + p('Second paragraph.'))), name: 'yazı.docx' });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 'd1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
      // The pages alone: the formatting toolbar comes with `toolbar`.
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveCount(1);
      await expect(bridge.editor.locator('#editor .bym-toolbar')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor .bym-doc-input')).toHaveJSProperty('readOnly', false);

      await typeIntoDocument(page, bridge.editor, 'Typed ', isMobile);
      expect(await bridge.waitFor('changed')).toEqual({ type: 'changed', sheets: [], canUndo: true, canRedo: false });
      await bridge.send({ type: 'save', id: 's1' });
      expect(await bridge.waitFor('saved')).toEqual({ type: 'saved', id: 's1', data: expect.any(Number), name: 'yazı.docx' });
      expect(await textOfSaved(bridge, 0)).toEqual({ format: 'docx', text: expect.stringContaining('Typed ') });
      expect((await textOfSaved(bridge, 0)).text).toContain('Second paragraph.');

      // Undo takes the typing back, redo brings it again; both are reported.
      const before = (await bridge.messages()).filter((message) => message.type === 'changed').length;
      await bridge.send({ type: 'undo', id: 'u' });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'u', command: 'undo' });
      const undone = await bridge.waitFor('changed', before + 1);
      expect(undone).toEqual({ type: 'changed', sheets: [], canUndo: false, canRedo: true });
      await bridge.send({ type: 'save', id: 's2' });
      await bridge.waitFor('saved', 2);
      expect((await textOfSaved(bridge, 1)).text).toBe('First paragraph.\nSecond paragraph.');
      await bridge.send({ type: 'redo' });
      expect(await bridge.waitFor('changed', before + 2)).toEqual({ type: 'changed', sheets: [], canUndo: true, canRedo: false });

      // Read-only: typing changes nothing and nothing is reported.
      await bridge.send({ type: 'setReadOnly', id: 'ro', readOnly: true });
      expect(await bridge.waitFor('done', 2)).toEqual({ type: 'done', id: 'ro', command: 'setReadOnly' });
      await expect(bridge.editor.locator('#editor .bym-doc-input')).toHaveJSProperty('readOnly', true);
      const count = (await bridge.messages()).filter((message) => message.type === 'changed').length;
      await bridge.editor.locator('#editor .bym-doc-input').focus();
      await page.keyboard.type('no');
      await bridge.send({ type: 'save', id: 's3' });
      await bridge.waitFor('saved', 3);
      expect((await textOfSaved(bridge, 2)).text).toContain('Typed ');
      expect((await textOfSaved(bridge, 2)).text).not.toContain('no');
      expect((await bridge.messages()).filter((message) => message.type === 'changed')).toHaveLength(count);

      const all = await bridge.messages();
      expect(all.filter((message) => message.notAString || message.type === 'error' || message.type === 'selection')).toEqual([]);
      expectNoOutsideRequests(bridge.requests);
    });

    test('convert: an RTF file is edited and saved as a docx; without convert it stays read-only', async ({ page, isMobile }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      const rtf = new TextEncoder().encode('{\\rtf1\\ansi Old text of the file.\\par}');

      await bridge.send({ type: 'load', id: 'r1', data: base64(rtf), name: 'eski.rtf' });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 'r1', format: 'rtf', kind: 'document', sheets: [], activeSheet: 0, readOnly: true });
      await expect(bridge.editor.locator('#editor .bym-doc-input')).toHaveJSProperty('readOnly', true);
      await bridge.send({ type: 'setReadOnly', id: 'open', readOnly: false });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'open', command: 'setReadOnly', readOnly: true });
      await expect(bridge.editor.locator('#editor .bym-doc-input')).toHaveJSProperty('readOnly', true);
      await bridge.send({ type: 'save', id: 's0' });
      await bridge.waitFor('saved');
      expect(Buffer.from(await bridge.savedData(0), 'base64').equals(Buffer.from(rtf))).toBe(true);

      await bridge.send({ type: 'load', id: 'r2', data: base64(rtf), name: 'eski.rtf', convert: true });
      expect(await bridge.waitFor('loaded', 2)).toEqual({ type: 'loaded', id: 'r2', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, convertedFrom: 'rtf', warnings: [] });
      await expect(bridge.editor.locator('#editor .bym-doc-input')).toHaveJSProperty('readOnly', false);
      await typeIntoDocument(page, bridge.editor, 'New ', isMobile);
      expect(await bridge.waitFor('changed')).toEqual({ type: 'changed', sheets: [], canUndo: true, canRedo: false });
      await bridge.send({ type: 'save', id: 's1' });
      expect(await bridge.waitFor('saved', 2)).toEqual({ type: 'saved', id: 's1', data: expect.any(Number), name: 'eski.rtf', format: 'docx', convertedFrom: 'rtf' });
      const saved = await textOfSaved(bridge, 1);
      expect(saved.format).toBe('docx');
      expect(saved.text).toContain('New ');
      expect(saved.text).toContain('text of the file.');

      // An xls file with a chart sheet: converted, and the answer says what was left behind.
      const xlsBytes = xlsFile(workbookStream({ sheets: [{ name: 'Old', records: cat(number(0, 0, 41)) }, { name: 'Chart1', type: 2 }] }));
      await bridge.send({ type: 'load', id: 'x1', data: base64(xlsBytes), convert: true });
      const loaded = await bridge.waitFor('loaded', 3);
      expect(loaded).toMatchObject({ id: 'x1', format: 'xlsx', convertedFrom: 'xls' });
      expect(loaded.readOnly).toBeUndefined();
      expect(loaded.warnings).toHaveLength(1);
      await typeIntoGrid(page, bridge.editor, '7');
      await bridge.send({ type: 'save', id: 's2' });
      expect(await bridge.waitFor('saved', 3)).toEqual({ type: 'saved', id: 's2', data: expect.any(Number), format: 'xlsx', convertedFrom: 'xls' });
      const again = await open(new Uint8Array(Buffer.from(await bridge.savedData(2), 'base64')));
      expect(again.format).toBe('xlsx');
      expect(getCell((await again.workbook()).sheets[0], 0, 0)).toMatchObject({ value: 7 });

      expect((await bridge.messages()).filter((message) => message.notAString || message.type === 'error')).toEqual([]);
      expectNoOutsideRequests(bridge.requests);
    });

    test('blank: a new workbook and a new document are started, edited and saved', async ({ page, isMobile }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      await bridge.send({ type: 'load', id: 'b1', blank: 'xlsx', locale: 'tr', name: 'yeni.xlsx' });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 'b1', format: 'xlsx', sheets: [{ name: 'Sayfa1', state: 'visible', kind: 'worksheet' }], activeSheet: 0 });
      await expect(bridge.editor.locator('#editor .bym-tabs button')).toHaveText(['Sayfa1']);
      await typeIntoGrid(page, bridge.editor, '12');
      await bridge.waitFor('changed');
      await bridge.send({ type: 'save', id: 's1' });
      expect(await bridge.waitFor('saved')).toEqual({ type: 'saved', id: 's1', data: expect.any(Number), name: 'yeni.xlsx' });
      const workbook = await (await open(new Uint8Array(Buffer.from(await bridge.savedData(0), 'base64')))).workbook();
      expect(getCell(workbook.sheets[0], 0, 0)).toMatchObject({ kind: 'number', value: 12 });

      await bridge.send({ type: 'load', id: 'b2', blank: 'docx' });
      expect(await bridge.waitFor('loaded', 2)).toEqual({ type: 'loaded', id: 'b2', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
      await typeIntoDocument(page, bridge.editor, 'Hello', isMobile);
      await bridge.waitFor('changed', 2);
      await bridge.send({ type: 'save', id: 's2' });
      await bridge.waitFor('saved', 2);
      expect(await textOfSaved(bridge, 1)).toEqual({ format: 'docx', text: 'Hello' });

      // Exactly one source, and only the two kinds.
      await bridge.send({ type: 'load', id: 'e1', blank: 'pptx' });
      await bridge.send({ type: 'load', id: 'e2', blank: 'docx', data: 'AAAA' });
      expect(await bridge.waitFor('error', 2)).toMatchObject({ id: 'e2', code: 'bad-message' });
      expect((await bridge.messages()).filter((message) => message.type === 'error').map((message) => [message.id, message.code])).toEqual([
        ['e1', 'bad-message'],
        ['e2', 'bad-message'],
      ]);
      // The document that was open is still there.
      await expect(bridge.editor.locator('#editor .bym-doc')).toHaveCount(1);
      expectNoOutsideRequests(bridge.requests);
    });

    test('toolbar: the editors with their bars are shown, and the commands and messages stay the same', async ({ page, isMobile }) => {
      const bridge = await start(page, mode);
      await bridge.waitFor('ready');
      await bridge.send({ type: 'load', id: 't1', blank: 'docx', toolbar: true, locale: 'tr' });
      expect(await bridge.waitFor('loaded')).toEqual({ type: 'loaded', id: 't1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
      const toolbar = bridge.editor.locator('#editor .bym-doc-editor .bym-toolbar');
      await expect(toolbar).toBeVisible();
      await expect(bridge.editor.locator('#editor .bym-doc-editor .bym-doc')).toHaveCount(1);

      await typeIntoDocument(page, bridge.editor, 'Bold', isMobile);
      expect(await bridge.waitFor('changed')).toEqual({ type: 'changed', sheets: [], canUndo: true, canRedo: false });
      await bridge.send({ type: 'undo', id: 'u' });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'u', command: 'undo' });
      await bridge.send({ type: 'redo', id: 'r' });
      expect(await bridge.waitFor('done', 2)).toEqual({ type: 'done', id: 'r', command: 'redo' });
      await bridge.send({ type: 'save', id: 's' });
      await bridge.waitFor('saved');
      expect(await textOfSaved(bridge, 0)).toEqual({ format: 'docx', text: 'Bold' });

      // The shell follows the commands: no toolbar while the document is read-only, and its words in the language asked for.
      await bridge.send({ type: 'setReadOnly', id: 'ro', readOnly: true });
      await bridge.waitFor('done', 3);
      await expect(toolbar).toBeHidden();
      await bridge.send({ type: 'setReadOnly', id: 'rw', readOnly: false });
      await bridge.waitFor('done', 4);
      await expect(toolbar).toBeVisible();
      const label = () => toolbar.locator('button').first().getAttribute('aria-label');
      const turkish = await label();
      await bridge.send({ type: 'setLocale', id: 'en', locale: 'en' });
      await bridge.waitFor('done', 5);
      await expect.poll(label).not.toBe(turkish);

      // A workbook with `toolbar`: the spreadsheet editor, as the page shows it without it too.
      await bridge.send({ type: 'load', id: 't2', blank: 'xlsx', toolbar: true });
      await bridge.waitFor('loaded', 2);
      await expect(bridge.editor.locator('#editor .bym-doc-editor')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor .bym-editor .bym-toolbar')).toBeVisible();
      await expect(bridge.editor.locator('#editor .bym-editor .bym-formula')).toBeVisible();
      await bridge.send({ type: 'load', id: 't3', blank: 'xlsx' });
      await bridge.waitFor('loaded', 3);
      await expect(bridge.editor.locator('#editor .bym-editor .bym-toolbar')).toBeVisible();

      // A workbook with `toolbar: false`: a viewer. The grid, its tabs and the status bar; no toolbar, no formula bar.
      await bridge.send({ type: 'load', id: 't4', blank: 'xlsx', toolbar: false, readOnly: true });
      await bridge.waitFor('loaded', 4);
      await expect(bridge.editor.locator('#editor canvas').first()).toBeVisible();
      await expect(bridge.editor.locator('#editor .bym-toolbar')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor .bym-formula')).toHaveCount(0);
      await expect(bridge.editor.locator('#editor [role="tab"]').first()).toBeVisible();

      // Closing removes the shell with the view.
      await bridge.send({ type: 'close', id: 'c' });
      await bridge.waitFor('done', 6);
      await expect(bridge.editor.locator('#editor')).toBeEmpty();
      expect((await bridge.messages()).filter((message) => message.notAString || message.type === 'error')).toEqual([]);
      expectNoOutsideRequests(bridge.requests);
    });

    test('sheets: inserted, renamed and deleted by command and by the user, the list is sent each time; refusals carry a code', async ({ page }) => {
      const bridge = await start(page, mode);
      await bridge.send({ type: 'load', id: 'l', data: base64(await smallWorkbook()), locale: 'tr' });
      await bridge.waitFor('loaded');
      const one = { name: 'Sayfa1', state: 'visible', kind: 'worksheet' };
      const sheet = (name: string) => ({ ...one, name });

      // A new sheet after the last one, with the default name of the view's language; it is the one shown.
      await bridge.send({ type: 'insertSheet', id: 'i1' });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'i1', command: 'insertSheet', index: 1, name: 'Sayfa2' });
      expect(await bridge.waitFor('sheets')).toEqual({ type: 'sheets', sheets: [one, sheet('Sayfa2')], activeSheet: 1 });
      expect(await bridge.waitFor('sheet')).toEqual({ type: 'sheet', index: 1, name: 'Sayfa2' });
      await expect(bridge.editor.locator('#editor .bym-tabs button')).toHaveText(['Sayfa1', 'Sayfa2']);
      // The list comes before what else the command caused.
      const order = (await bridge.messages()).map((message) => message.type).filter((type) => type === 'sheets' || type === 'sheet' || type === 'changed' || type === 'done');
      expect(order).toEqual(['sheets', 'sheet', 'changed', 'done']);

      await bridge.send({ type: 'insertSheet', id: 'i2', at: 0, name: 'Özet' });
      expect(await bridge.waitFor('done', 2)).toEqual({ type: 'done', id: 'i2', command: 'insertSheet', index: 0, name: 'Özet' });
      expect(await bridge.waitFor('sheets', 2)).toEqual({ type: 'sheets', sheets: [sheet('Özet'), one, sheet('Sayfa2')], activeSheet: 0 });

      await bridge.send({ type: 'renameSheet', id: 'r', index: 2, name: 'Notlar' });
      expect(await bridge.waitFor('done', 3)).toEqual({ type: 'done', id: 'r', command: 'renameSheet' });
      expect(await bridge.waitFor('sheets', 3)).toEqual({ type: 'sheets', sheets: [sheet('Özet'), one, sheet('Notlar')], activeSheet: 0 });
      await expect(bridge.editor.locator('#editor .bym-tabs button')).toHaveText(['Özet', 'Sayfa1', 'Notlar']);

      await bridge.send({ type: 'deleteSheet', id: 'd', index: 0 });
      expect(await bridge.waitFor('done', 4)).toEqual({ type: 'done', id: 'd', command: 'deleteSheet' });
      expect(await bridge.waitFor('sheets', 4)).toEqual({ type: 'sheets', sheets: [one, sheet('Notlar')], activeSheet: 0 });

      // Undo brings the sheet back, and the list says so.
      await bridge.send({ type: 'undo', id: 'u' });
      await bridge.waitFor('done', 5);
      expect(await bridge.waitFor('sheets', 5)).toEqual({ type: 'sheets', sheets: [sheet('Özet'), one, sheet('Notlar')], activeSheet: 0 });

      // The user adds a sheet with the button of the view: no command, the same message. The name is the first one that is free.
      await bridge.editor.locator('#editor .bym-tab-add').click();
      expect(await bridge.waitFor('sheets', 6)).toEqual({ type: 'sheets', sheets: [sheet('Özet'), one, sheet('Notlar'), sheet('Sayfa2')], activeSheet: 3 });

      // Refusals: the reason in the language of the view, the code the same for all of them.
      await bridge.send({ type: 'renameSheet', id: 'e1', index: 0, name: 'notlar' });
      expect(await bridge.waitFor('error')).toEqual({ type: 'error', id: 'e1', code: 'sheet-refused', message: expect.stringContaining('notlar') });
      await bridge.send({ type: 'renameSheet', id: 'e2', index: 0, name: 'a/b' });
      expect(await bridge.waitFor('error', 2)).toMatchObject({ id: 'e2', code: 'sheet-refused' });
      await bridge.send({ type: 'deleteSheet', id: 'e3', index: 9 });
      expect(await bridge.waitFor('error', 3)).toEqual({ type: 'error', id: 'e3', code: 'no-sheet', message: 'No sheet with index 9' });
      await bridge.send({ type: 'insertSheet', id: 'e4', name: 7 });
      expect(await bridge.waitFor('error', 4)).toMatchObject({ id: 'e4', code: 'bad-message' });
      await bridge.send({ type: 'setReadOnly', readOnly: true });
      await bridge.send({ type: 'insertSheet', id: 'e5' });
      expect(await bridge.waitFor('error', 5)).toEqual({ type: 'error', id: 'e5', code: 'read-only', message: 'The file is read-only' });
      // Nothing of what was refused changed the list.
      expect((await bridge.messages()).filter((message) => message.type === 'sheets')).toHaveLength(6);

      // The saved file has the sheets, in order.
      await bridge.send({ type: 'save', id: 's' });
      await bridge.waitFor('saved');
      const workbook = await (await open(new Uint8Array(Buffer.from(await bridge.savedData(), 'base64')))).workbook();
      expect(workbook.sheets.map((entry) => entry.name)).toEqual(['Özet', 'Sayfa1', 'Notlar', 'Sayfa2']);
      expect(workbook.activeSheet).toBe(3);

      // A document has no sheets.
      await bridge.send({ type: 'load', id: 'doc', blank: 'docx' });
      await bridge.waitFor('loaded', 2);
      await bridge.send({ type: 'insertSheet', id: 'e6' });
      expect(await bridge.waitFor('error', 6)).toEqual({ type: 'error', id: 'e6', code: 'no-sheet', message: 'The open file is a document: it has no sheets' });
      expect((await bridge.messages()).filter((message) => message.notAString)).toEqual([]);
      expectNoOutsideRequests(bridge.requests);
    });

    test('sheets: moved, hidden and shown, copied and coloured by command; the list is sent each time and the saved file has it', async ({ page }) => {
      const bridge = await start(page, mode);
      const three = await buildXlsx({
        sheets: ['Bir', 'İki', 'Üç'].map((name, index) => ({ name, xml: worksheetXml(`<sheetData><row r="1"><c r="A1"><v>${index + 1}</v></c></row></sheetData>`) })),
      });
      await bridge.send({ type: 'load', id: 'l', data: base64(three), locale: 'tr' });
      await bridge.waitFor('loaded');
      const sheet = (name: string, more: Record<string, unknown> = {}) => ({ name, state: 'visible', kind: 'worksheet', ...more });

      await bridge.send({ type: 'moveSheet', id: 'm', from: 0, to: 2 });
      expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'm', command: 'moveSheet' });
      // The sheet that was shown is still the one shown, at the place it has now.
      expect(await bridge.waitFor('sheets')).toEqual({ type: 'sheets', sheets: [sheet('İki'), sheet('Üç'), sheet('Bir')], activeSheet: 2 });
      await expect(bridge.editor.locator('#editor .bym-tabs button.bym-tab')).toHaveText(['İki', 'Üç', 'Bir']);

      await bridge.send({ type: 'setSheetState', id: 'h', index: 2, state: 'hidden' });
      expect(await bridge.waitFor('done', 2)).toEqual({ type: 'done', id: 'h', command: 'setSheetState' });
      expect(await bridge.waitFor('sheets', 2)).toEqual({ type: 'sheets', sheets: [sheet('İki'), sheet('Üç'), sheet('Bir', { state: 'hidden' })], activeSheet: 1 });
      await expect(bridge.editor.locator('#editor .bym-tabs button.bym-tab')).toHaveText(['İki', 'Üç']);

      await bridge.send({ type: 'copySheet', id: 'c', index: 0 });
      expect(await bridge.waitFor('done', 3)).toEqual({ type: 'done', id: 'c', command: 'copySheet', index: 1, name: 'İki (2)' });
      expect(await bridge.waitFor('sheets', 3)).toEqual({ type: 'sheets', sheets: [sheet('İki'), sheet('İki (2)'), sheet('Üç'), sheet('Bir', { state: 'hidden' })], activeSheet: 1 });

      await bridge.send({ type: 'setTabColor', id: 't', index: 2, color: '#ff8800' });
      expect(await bridge.waitFor('done', 4)).toEqual({ type: 'done', id: 't', command: 'setTabColor' });
      expect(await bridge.waitFor('sheets', 4)).toEqual({
        type: 'sheets',
        sheets: [sheet('İki'), sheet('İki (2)'), sheet('Üç', { tabColor: '#FF8800' }), sheet('Bir', { state: 'hidden' })],
        activeSheet: 1,
      });

      await bridge.send({ type: 'save', id: 's' });
      await bridge.waitFor('saved');
      const workbook = await (await open(new Uint8Array(Buffer.from(await bridge.savedData(), 'base64')))).workbook();
      expect(workbook.sheets.map((entry) => `${entry.name}:${entry.state}:${entry.tabColor ? entry.tabColor.rgb : '-'}:${getCell(entry, 0, 0)!.value}`)).toEqual([
        'İki:visible:-:2',
        'İki (2):visible:-:2',
        'Üç:visible:FFFF8800:3',
        'Bir:hidden:-:1',
      ]);
      expect(workbook.activeSheet).toBe(1);

      // Refusals carry the code and the reason in the language of the view; fields that are not in the contract are bad messages.
      await bridge.send({ type: 'setSheetState', id: 'e1', index: 0, state: 'hidden' });
      await bridge.send({ type: 'setSheetState', id: 'e2', index: 1, state: 'hidden' });
      await bridge.waitFor('done', 6);
      await bridge.send({ type: 'setSheetState', id: 'e3', index: 2, state: 'hidden' });
      expect(await bridge.waitFor('error')).toEqual({ type: 'error', id: 'e3', code: 'sheet-refused', message: 'Görünen son sayfa gizlenemez' });
      await bridge.send({ type: 'setTabColor', id: 'e4', index: 0, color: 'red' });
      expect(await bridge.waitFor('error', 2)).toMatchObject({ id: 'e4', code: 'bad-message' });
      await bridge.send({ type: 'moveSheet', id: 'e5', from: 9, to: 0 });
      expect(await bridge.waitFor('error', 3)).toMatchObject({ id: 'e5', code: 'no-sheet' });
      // Undo by command takes the last change of the list back, and the list says so.
      await bridge.send({ type: 'undo', id: 'u' });
      await bridge.waitFor('done', 7);
      expect(((await bridge.messages()).filter((message) => message.type === 'sheets').pop() as any).sheets.map((entry: any) => entry.state)).toEqual(['hidden', 'visible', 'visible', 'hidden']);
      expectNoOutsideRequests(bridge.requests);
    });
  });
}

test.describe('bridge page: who may embed it', () => {
  test('as shipped, with no parent origins listed, a framed page neither answers nor speaks', async ({ page }) => {
    // From another origin and from its own: the list is what counts, not where the parent is.
    for (const parent of [APP, new URL(test.info().project.use.baseURL!).origin]) {
      const bridge = await start(page, 'iframe', { parent, allow: '' });
      await bridge.send({ type: 'hello' });
      await bridge.send({ type: 'save', id: 'x' });
      await bridge.send({ type: 'load', id: 'l', data: base64(await smallWorkbook()) });
      await expectSilence(page, bridge);
      await expect(bridge.editor.locator('#editor')).toBeEmpty();
    }
  });

  test('an origin that is not listed gets nothing, whatever it sends', async ({ page, browserName, pageProblems }) => {
    const bridge = await start(page, 'iframe', { parent: 'http://evil.example', allow: APP });
    await bridge.send({ type: 'hello' });
    await bridge.send({ type: 'load', id: 'l', data: base64(await smallWorkbook()) });
    await bridge.send({ type: 'save', id: 's' });
    await expectSilence(page, bridge);
    await expect(bridge.editor.locator('#editor')).toBeEmpty();
    // Firefox does not tell a frame who its parent is, so `ready` is addressed to the one
    // listed origin and the browser refuses to deliver it; it may say so on the console.
    if (browserName === 'firefox') {
      expect(pageProblems.filter((problem) => !/postMessage.*http:\/\/app\.example/.test(problem))).toEqual([]);
      pageProblems.length = 0;
    }
  });

  test('listed origins are compared in normal form; with several, the parent is recognised', async ({ page }) => {
    const bridge = await start(page, 'iframe', { allow: 'https://other.example HTTP://APP.EXAMPLE:80/ null *' });
    // Where the browser does not tell the frame who its parent is, the first message from it does.
    await bridge.send({ type: 'hello' });
    await bridge.waitFor('ready');
    await bridge.send({ type: 'save', id: 'a' });
    expect((await bridge.waitFor('error')).id).toBe('a');
  });

  test('a framed page loads from data only, never from a URL', async ({ page }) => {
    const bytes = await smallWorkbook();
    await page.route('**/files/secret.xlsx', (route) => route.fulfill({ body: Buffer.from(bytes), contentType: 'application/octet-stream' }));
    const bridge = await start(page, 'iframe');
    await bridge.waitFor('ready');
    const urls = ['/files/secret.xlsx', `${new URL(test.info().project.use.baseURL!).origin}/files/secret.xlsx`, `data:application/octet-stream;base64,${base64(bytes)}`];
    for (let i = 0; i < urls.length; i++) {
      await bridge.send({ type: 'load', id: `u${i}`, url: urls[i] });
      expect(await bridge.waitFor('error', i + 1)).toMatchObject({ id: `u${i}`, code: 'bad-message' });
    }
    await expect(bridge.editor.locator('#editor')).toBeEmpty();
    expect(bridge.requests.filter((url) => url.includes('secret'))).toEqual([]);
  });

  test('messages from a sibling frame and from the frame itself are ignored', async ({ page }) => {
    const bridge = await start(page, 'iframe');
    await bridge.waitFor('ready');
    // A sibling of the editor, of the allowed origin even, is not its parent.
    await page.evaluate(
      () =>
        new Promise<void>((resolve) => {
          const sibling = document.createElement('iframe');
          sibling.onload = () => {
            const script = sibling.contentDocument!.createElement('script');
            script.textContent = 'parent.frames[0].postMessage(\'{"type":"save","id":"sibling"}\', "*")';
            sibling.contentDocument!.body.appendChild(script);
            resolve();
          };
          sibling.src = '/host.html?src=about%3Ablank';
          document.body.appendChild(sibling);
        }),
    );
    await page.frameLocator('#bridge').locator('body').evaluate(() => window.postMessage('{"type":"save","id":"self"}', '*'));
    await bridge.send({ type: 'save', id: 'marker' });
    await bridge.waitFor('error');
    expect((await bridge.messages()).map((message) => message.id)).toEqual([undefined, 'marker']);
  });

  test('a top-level page takes no window messages: not its own, not its opener\'s', async ({ page, context }) => {
    await context.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto('/e2e/pages/script.html');
    const [popup] = await Promise.all([page.waitForEvent('popup'), page.evaluate((url) => void ((window as any).__popup = window.open(url)), PAGE)]);
    const opened = app(popup, 'channel', []);
    await opened.waitFor('ready');
    await page.evaluate(() => (window as any).__popup.postMessage('{"type":"save","id":"opener"}', '*'));
    await popup.evaluate(() => window.postMessage('{"type":"save","id":"self"}', '*'));
    await opened.send({ type: 'save', id: 'marker' });
    await opened.waitFor('error');
    expect((await opened.messages()).map((message) => message.id)).toEqual([undefined, 'marker']);
  });
});

test.describe('bridge page: channels and security', () => {
  test('ready still arrives when the channel is injected after the page has loaded', async ({ page }) => {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      setTimeout(() => {
        (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
      }, 700);
    });
    await page.goto(PAGE);
    expect(await page.evaluate(() => typeof (window as any).BymOfficeBridge.receive)).toBe('function');
    expect(await page.evaluate(() => (window as any).__messages.length)).toBe(0);
    const bridge = app(page, 'channel', []);
    await bridge.waitFor('ready');
    expect((await bridge.messages()).map((message) => message.type)).toEqual(['ready']);
  });

  test('receive() before any channel exists keeps the answers until there is one', async ({ page }) => {
    await page.goto(PAGE);
    // Returns nothing: some WebViews fail on a result they cannot pass on.
    expect(await page.evaluate(() => (window as any).BymOfficeBridge.receive('{"type":"save","id":"early"}'))).toBeUndefined();
    await page.evaluate(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    const bridge = app(page, 'channel', []);
    await bridge.waitFor('error');
    expect((await bridge.messages()).map((message) => message.type)).toEqual(['ready', 'error']);
  });

  test('the other channels are used in order of preference', async ({ page }) => {
    await page.addInitScript(() => {
      const log: string[] = ((window as any).__log = []);
      const note = (name: string) => (json: string) => log.push(`${name} ${JSON.parse(json).type}`);
      (window as any).webkit = { messageHandlers: { bymoffice: { postMessage: note('webkit') } } };
      (window as any).ReactNativeWebView = { postMessage: note('rn') };
      (window as any).flutter_inappwebview = { callHandler: (name: string, json: string) => note(name)(json) };
      const chrome = (window as any).chrome || ((window as any).chrome = {});
      chrome.webview = { postMessage: note('webview2') };
    });
    await page.goto(PAGE);
    const log = () => page.evaluate(() => (window as any).__log);
    await expect.poll(log).toEqual(['webkit ready']);
    for (const [gone, expected] of [
      ['webkit', 'bymoffice error'],
      // The in-app plugin runs on top of WebView2 on Windows, so it comes before WebView2's own channel.
      ['flutter_inappwebview', 'webview2 error'],
      ['chrome', 'rn error'],
    ]) {
      await page.evaluate((name) => {
        if (name === 'chrome') delete (window as any).chrome.webview;
        else delete (window as any)[name];
        (window as any).BymOfficeBridge.receive('{"type":"undo"}');
      }, gone);
      await expect.poll(async () => (await log()).slice(-1)[0]).toBe(expected);
    }
    expect(await log()).toEqual(['webkit ready', 'bymoffice error', 'webview2 error', 'rn error']);
  });

  test('a WebView2-style channel sends and receives on window.chrome.webview', async ({ page }) => {
    await page.addInitScript(() => {
      const log: string[] = ((window as any).__log = []);
      const listeners: ((event: { data: unknown }) => void)[] = [];
      const chrome = (window as any).chrome || ((window as any).chrome = {});
      chrome.webview = {
        postMessage: (json: string) => log.push(JSON.parse(json).code || JSON.parse(json).type),
        addEventListener: (type: string, listener: (event: { data: unknown }) => void) => type === 'message' && listeners.push(listener),
      };
      (window as any).__fromApp = (data: unknown) => listeners.forEach((listener) => listener({ data }));
    });
    await page.goto(PAGE);
    await expect.poll(() => page.evaluate(() => (window as any).__log)).toEqual(['ready']);
    // PostWebMessageAsString and PostWebMessageAsJson: text and an object.
    await page.evaluate(() => {
      (window as any).__fromApp('{"type":"save"}');
      (window as any).__fromApp({ type: 'explode' });
    });
    await expect.poll(() => page.evaluate(() => (window as any).__log)).toEqual(['ready', 'not-loaded', 'bad-message']);
  });

  test('installBridge().destroy() removes the global, the listeners and the view', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    const bytes = base64(await smallWorkbook());
    const result = await page.evaluate(async (data) => {
      const api = (window as any).BymOffice;
      const sent: string[] = [];
      const listeners = new Set<unknown>();
      const chrome = (window as any).chrome || ((window as any).chrome = {});
      chrome.webview = {
        postMessage: (json: string) => sent.push(JSON.parse(json).type),
        addEventListener: (_type: string, listener: unknown) => listeners.add(listener),
        removeEventListener: (_type: string, listener: unknown) => listeners.delete(listener),
      };
      const bridge = api.installBridge({
        window,
        container: document.getElementById('editor'),
        open: api.open,
        createView: (container: HTMLElement, workbook: unknown, options: unknown) => new api.WorkbookView(container, workbook, options),
        version: '1.2.3',
      });
      await bridge.receive(JSON.stringify({ type: 'load', data }));
      const before = { global: typeof (window as any).BymOfficeBridge, listeners: listeners.size, canvas: document.querySelectorAll('#editor canvas').length };
      bridge.destroy();
      await bridge.receive('{"type":"save"}');
      return { before, sent, global: typeof (window as any).BymOfficeBridge, listeners: listeners.size, children: document.getElementById('editor')!.childNodes.length };
    }, bytes);
    expect(result).toEqual({ before: { global: 'object', listeners: 1, canvas: 1 }, sent: ['ready', 'loaded'], global: 'undefined', listeners: 0, children: 0 });
  });

  test('load by URL: the page\'s own origin and data: work; a missing file is load-failed', async ({ page, pageProblems }) => {
    const bytes = await smallWorkbook();
    await page.route('**/files/sample.xlsx', (route) => route.fulfill({ body: Buffer.from(bytes), contentType: 'application/octet-stream' }));
    await page.route('**/files/missing.xlsx', (route) => route.fulfill({ status: 404, body: 'nope' }));
    const bridge = await start(page, 'channel');

    await bridge.send({ type: 'load', id: 'u1', url: '/files/sample.xlsx' });
    expect((await bridge.waitFor('loaded')).id).toBe('u1');
    await bridge.send({ type: 'load', id: 'u2', url: `data:application/octet-stream;base64,${base64(bytes)}` });
    expect((await bridge.waitFor('loaded', 2)).id).toBe('u2');
    expect(pageProblems).toEqual([]);

    await bridge.send({ type: 'load', id: 'u3', url: '/files/missing.xlsx' });
    expect(await bridge.waitFor('error')).toEqual({ type: 'error', id: 'u3', code: 'load-failed', message: 'Could not load the file: HTTP 404' });
    // The file that was open is still there.
    await expect(bridge.editor.locator('#editor canvas')).toBeVisible();
    // The browser reports the 404 on the console itself; nothing else may be there.
    expect(pageProblems.filter((problem) => !problem.includes('/files/missing.xlsx') && !/404/.test(problem))).toEqual([]);
    pageProblems.length = 0;
    expectNoOutsideRequests(bridge.requests);
  });

  test('a load that hangs does not block the bridge: close and a new load cancel it', async ({ page }) => {
    const held: (() => void)[] = [];
    await page.route('**/files/slow.xlsx', (route) => {
      // Never answered while the test runs.
      held.push(() => void route.abort().catch(() => undefined));
    });
    const bridge = await start(page, 'channel');
    await bridge.waitFor('ready');

    await bridge.send({ type: 'load', id: 'h1', url: '/files/slow.xlsx' });
    await expect(bridge.editor.locator('#status')).toHaveText('Opening the file…');
    await expect.poll(() => held.length).toBe(1);
    await bridge.send({ type: 'save', id: 'queued' });
    await bridge.send({ type: 'close', id: 'c' });
    expect(await bridge.waitFor('done')).toEqual({ type: 'done', id: 'c', command: 'close' });
    expect(await bridge.waitFor('error', 2)).toBeTruthy();
    expect((await bridge.messages()).slice(1).map((message) => [message.id, message.code ?? message.type])).toEqual([
      ['c', 'done'],
      ['h1', 'load-failed'],
      ['queued', 'not-loaded'],
    ]);
    await expect(bridge.editor.locator('#status')).toHaveText('No file is open');

    // A second load overtakes a hanging one.
    await bridge.send({ type: 'load', id: 'h2', url: '/files/slow.xlsx' });
    await expect.poll(() => held.length).toBe(2);
    await bridge.send({ type: 'load', id: 'next', data: base64(await smallWorkbook()) });
    expect((await bridge.waitFor('loaded')).id).toBe('next');
    expect(await bridge.waitFor('error', 3)).toMatchObject({ id: 'h2', code: 'load-failed', message: 'The load was cancelled' });
    await expect(bridge.editor.locator('#editor canvas')).toBeVisible();
    for (const release of held) release();
  });

  test('the page carries its own policy and the engine, and loads nothing else', async ({ page, pageProblems }) => {
    const bridge = await start(page, 'channel');
    await bridge.waitFor('ready');
    const facts = await page.evaluate(() => ({
      policy: document.querySelector('meta[http-equiv="Content-Security-Policy"]')!.getAttribute('content'),
      viewport: document.querySelector('meta[name="viewport"]')!.getAttribute('content'),
      parents: document.querySelector('meta[name="bymoffice-parent-origins"]')!.getAttribute('content'),
      external: document.querySelectorAll('script[src], link[rel="stylesheet"], img, iframe').length,
      engine: typeof (window as any).BymOffice.open,
    }));
    expect(facts.policy).toMatch(/^default-src 'none'; script-src 'sha256-[A-Za-z0-9+/=]+' 'sha256-[A-Za-z0-9+/=]+'; /);
    expect(facts.policy).not.toMatch(/script-src[^;]*unsafe|https?:|\*/);
    expect(facts.viewport).toContain('viewport-fit=cover');
    expect(facts.parents).toBe('');
    expect(facts.external).toBe(0);
    expect(facts.engine).toBe('function');
    expect(bridge.requests).toEqual([new URL(PAGE, page.url()).href]);

    // The policy is in force: a script added later does not run.
    const ran = await page.evaluate(
      () =>
        new Promise<boolean>((resolve) => {
          const script = document.createElement('script');
          script.textContent = 'window.__injected = true';
          document.addEventListener('securitypolicyviolation', () => resolve(false), { once: true });
          document.body.appendChild(script);
          setTimeout(() => resolve((window as any).__injected === true), 300);
        }),
    );
    expect(ran).toBe(false);
    // The browser reports the script it refused; that is the only thing on the console.
    expect(pageProblems.filter((problem) => !/Content Security Policy|Content-Security-Policy/i.test(problem))).toEqual([]);
    pageProblems.length = 0;
  });
});
