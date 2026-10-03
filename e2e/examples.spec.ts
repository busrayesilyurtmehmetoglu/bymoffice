import type { Page } from '@playwright/test';
import { createBlank } from '../src/index';
import { expect, test } from './helpers';

// examples/html/index.html: the page an application opens (in a tab or an
// iframe) to show a file in the editor and to get it back when it is saved.

const PAGE = '/examples/html/index.html';

/** Serves `bytes` at /files/sample and records what the page sends to /files/save. */
async function server(page: Page, bytes?: Uint8Array): Promise<{ saved: Buffer[] }> {
  const saved: Buffer[] = [];
  // By path: the address of the page itself ends in "save=/files/save".
  await page.route((url) => url.pathname === '/files/sample', (route) => route.fulfill({ status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes || new Uint8Array()) }));
  await page.route((url) => url.pathname === '/files/save', (route) => {
    const request = route.request();
    if (request.method() === 'PUT') saved.push(request.postDataBuffer() || Buffer.alloc(0));
    return route.fulfill({ status: 204, body: '' });
  });
  return { saved };
}

const ready = (page: Page) => page.waitForFunction(() => (window as any).bymEditor !== undefined, undefined, { timeout: 60_000 });

/** Opens bytes with the engine of the page and says what is in them. */
function contentOf(page: Page, bytes: Buffer): Promise<{ format: string; text: string }> {
  return page.evaluate(async (data) => {
    const api = (window as any).BymOffice;
    const file = await api.open(new Uint8Array(data));
    if (file.kind === 'document') return { format: file.format, text: api.documentText(await file.document()) };
    const workbook = await file.workbook();
    const cell = api.getCell(workbook.sheets[0], 0, 0);
    return { format: file.format, text: `${workbook.sheets[0].name}!A1=${cell ? cell.value : ''}` };
  }, Array.from(bytes));
}

test.describe('example page: a file of the site in the editor', () => {
  test('a new document is written in, saved with the button and arrives at the server as a docx', async ({ page }) => {
    const { saved } = await server(page);
    await page.goto(`${PAGE}?new=docx&save=/files/save&name=letter.docx&lang=tr`);
    await ready(page);
    await expect(page.locator('#name')).toHaveText('letter.docx');
    await expect(page.locator('#save')).toHaveText('Kaydet');
    await page.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('Merhaba dünya');
    });
    await expect(page.locator('#status')).toHaveText('Kaydedilmedi');
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Kaydedildi');
    expect(saved.length).toBe(1);
    expect(await contentOf(page, saved[0])).toEqual({ format: 'docx', text: 'Merhaba dünya' });
    expect(await page.evaluate(() => (window as any).bymEditor.view.isModified)).toBe(false);
  });

  test('a new workbook takes a value and is saved as an xlsx with a sheet named in the language', async ({ page }) => {
    const { saved } = await server(page);
    await page.goto(`${PAGE}?new=xlsx&save=/files/save&lang=tr`);
    await ready(page);
    await page.evaluate(() => (window as any).bymEditor.view.setCellInput(0, 0, '42'));
    await expect(page.locator('#status')).toHaveText('Kaydedilmedi');
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Kaydedildi');
    expect(saved.length).toBe(1);
    expect(await contentOf(page, saved[0])).toEqual({ format: 'xlsx', text: 'Sayfa1!A1=42' });
  });

  test('a file of the site is fetched and shown; in view mode nothing can be changed and there is no Save', async ({ page }) => {
    const { saved } = await server(page, await createBlank('docx'));
    await page.goto(`${PAGE}?file=/files/sample&save=/files/save&name=read.docx&mode=view`);
    await ready(page);
    await expect(page.locator('#save')).toBeHidden();
    const changed = await page.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      return [view.readOnly, view.insertText('x'), view.isModified];
    });
    expect(changed).toEqual([true, false, false]);
    expect(saved.length).toBe(0);
  });

  test('a file fetched for editing is saved back with what was typed', async ({ page }) => {
    const { saved } = await server(page, await createBlank('xlsx', { sheetName: 'Data' }));
    await page.goto(`${PAGE}?file=/files/sample&save=/files/save&name=data.xlsx`);
    await ready(page);
    await expect(page.locator('#save')).toHaveText('Save');
    await page.evaluate(() => (window as any).bymEditor.view.setCellInput(0, 0, 'hello'));
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Saved');
    expect(await contentOf(page, saved[0])).toEqual({ format: 'xlsx', text: 'Data!A1=hello' });
  });

  test('review: what is typed while the file is on its way is not called saved, and the next save holds it', async ({ page }) => {
    const saved: Buffer[] = [];
    let release: () => void = () => undefined;
    let held = new Promise<void>((resolve) => (release = resolve));
    await page.route((url) => url.pathname === '/files/save', async (route) => {
      saved.push(route.request().postDataBuffer() || Buffer.alloc(0));
      await held;
      await route.fulfill({ status: 204, body: '' });
    });
    await page.goto(`${PAGE}?new=docx&save=/files/save&name=race.docx`);
    await ready(page);
    const type = (text: string) =>
      page.evaluate((value) => {
        const view = (window as any).bymEditor.view;
        view.setSelection({ story: 'body', path: [0], offset: 0 });
        view.insertText(value);
      }, text);
    await type('first ');
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Saving…');
    await expect.poll(() => saved.length).toBe(1);
    // The request is under way: a second press sends nothing, and more is typed.
    await page.locator('#save').click();
    await type('second ');
    release();
    await expect(page.locator('#status')).toHaveText('Not saved');
    expect(saved.length).toBe(1);
    expect((await contentOf(page, saved[0])).text).toBe('first ');
    held = Promise.resolve();
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Saved');
    expect(saved.length).toBe(2);
    expect((await contentOf(page, saved[1])).text).toBe('second first ');
  });

  test('review: Ctrl+S in the editor saves like the button; without a name the file is called by its path', async ({ page }) => {
    const { saved } = await server(page, await createBlank('docx'));
    await page.goto(`${PAGE}?file=/files/sample?v=2&save=/files/save&new=pptx`);
    await ready(page);
    // The path, not the query; and a kind that is not one is no new file.
    await expect(page.locator('#name')).toHaveText('sample');
    await page.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.focus();
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('by the key');
    });
    await page.keyboard.press(process.platform === 'darwin' ? 'Meta+s' : 'Control+s');
    await expect(page.locator('#status')).toHaveText('Saved');
    expect(saved.length).toBe(1);
    expect((await contentOf(page, saved[0])).text).toBe('by the key');
  });

  test('convert=1: an RTF file is edited as a docx; the name, the header of the request and the message to the parent say so', async ({ page }) => {
    const formats: (string | undefined)[] = [];
    const saved: Buffer[] = [];
    await page.route((url) => url.pathname === '/files/sample', (route) => route.fulfill({ status: 200, contentType: 'application/octet-stream', body: Buffer.from('{\\rtf1\\ansi Old text\\par}') }));
    await page.route((url) => url.pathname === '/files/save', (route) => {
      formats.push(route.request().headers()['x-bym-format']);
      saved.push(route.request().postDataBuffer() || Buffer.alloc(0));
      return route.fulfill({ status: 204, body: '' });
    });
    // Without it the file is shown as it is, read-only.
    await page.goto(`${PAGE}?file=/files/sample&save=/files/save&name=old.rtf`);
    await ready(page);
    await expect(page.locator('#save')).toBeHidden();
    await expect(page.locator('#name')).toHaveText('old.rtf');

    await page.goto(`${PAGE}?file=/files/sample&save=/files/save&name=old.rtf&convert=1&lang=tr`);
    await ready(page);
    await expect(page.locator('#name')).toHaveText('old.docx');
    await expect(page.locator('#save')).toBeVisible();
    await page.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('New and ');
    });
    await page.locator('#save').click();
    await expect(page.locator('#status')).toHaveText('Kaydedildi');
    expect(formats).toEqual(['docx']);
    expect(await contentOf(page, saved[0])).toEqual({ format: 'docx', text: 'New and Old text' });
  });

  test('addresses of another site are refused, and so is a page without a file', async ({ page }) => {
    await page.goto(`${PAGE}?file=https://example.com/a.docx`);
    await expect(page.locator('#status')).toHaveText('Only files of this site can be opened');
    await page.goto(`${PAGE}?new=docx&save=//example.com/save`);
    await expect(page.locator('#status')).toHaveText('Only files of this site can be opened');
    await page.goto(`${PAGE}?lang=tr`);
    await expect(page.locator('#status')).toHaveText('Dosya verilmedi');
    expect(await page.evaluate(() => (window as any).bymEditor)).toBeUndefined();
  });

  test('inside an iframe the page tells its parent when it is ready, changed and saved', async ({ page }) => {
    const { saved } = await server(page);
    await page.goto('/e2e/pages/blank.html');
    await page.evaluate((src) => {
      (window as any).__messages = [];
      window.addEventListener('message', (event) => {
        if (event.origin === location.origin && event.data && event.data.source === 'bymoffice') (window as any).__messages.push(event.data.type);
      });
      const frame = document.createElement('iframe');
      frame.src = src;
      frame.style.cssText = 'width:800px;height:500px;border:0';
      document.body.appendChild(frame);
    }, `${PAGE}?new=docx&save=/files/save`);
    await page.waitForFunction(() => (window as any).__messages.includes('ready'), undefined, { timeout: 60_000 });
    const frame = page.frames().find((candidate) => candidate.url().includes('/examples/html/'))!;
    await frame.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('From the frame');
    });
    await frame.locator('#save').click();
    await page.waitForFunction(() => (window as any).__messages.includes('saved'));
    expect(await page.evaluate(() => (window as any).__messages)).toEqual(['ready', 'changed', 'saved']);
    expect(saved.length).toBe(1);
  });
});
