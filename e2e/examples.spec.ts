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

  test('convert=1: an RTF file is edited as a docx; the name and the header of the request say so, and the notice stays in the status line', async ({ page }) => {
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
    // The notice of the conversion is still there once the page is ready.
    await expect(page.locator('#status')).toHaveText('Bu rtf dosyası düzenlenebilmesi için docx biçimine çevrildi; docx olarak kaydedilir');
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

  test('convert=1: only the extension of the old format is replaced in the name; any other end of a name is kept', async ({ page }) => {
    await page.route((url) => url.pathname === '/files/sample', (route) => route.fulfill({ status: 200, contentType: 'application/octet-stream', body: Buffer.from('{\\rtf1\\ansi Old text\\par}') }));
    for (const [given, shown] of [
      ['Rapor 2024.10', 'Rapor 2024.10.docx'],
      ['teklif.v2', 'teklif.v2.docx'],
      ['OLD.RTF', 'OLD.docx'],
      ['letter.doc', 'letter.docx'],
      ['noextension', 'noextension.docx'],
      ['notes.docx', 'notes.docx.docx'],
    ]) {
      await page.goto(`${PAGE}?file=/files/sample&name=${encodeURIComponent(given)}&convert=1`);
      await ready(page);
      await expect(page.locator('#name')).toHaveText(shown);
      expect(await page.title()).toBe(shown);
    }
  });

  test('convert=1 inside an iframe: the messages to the parent carry the new format and the one it was converted from', async ({ page }) => {
    await page.route((url) => url.pathname === '/files/sample', (route) => route.fulfill({ status: 200, contentType: 'application/octet-stream', body: Buffer.from('{\\rtf1\\ansi Old text\\par}') }));
    await page.route((url) => url.pathname === '/files/save', (route) => route.fulfill({ status: 204, body: '' }));
    await page.goto('/e2e/pages/blank.html');
    const embed = (src: string) =>
      page.evaluate((address) => {
        (window as any).__messages = [];
        if (!(window as any).__listening) {
          (window as any).__listening = true;
          window.addEventListener('message', (event) => {
            if (event.origin === location.origin && event.data && event.data.source === 'bymoffice') (window as any).__messages.push(event.data);
          });
        }
        for (const old of Array.from(document.querySelectorAll('iframe'))) old.remove();
        const frame = document.createElement('iframe');
        frame.src = address;
        frame.style.cssText = 'width:800px;height:500px;border:0';
        document.body.appendChild(frame);
      }, src);
    const messages = () => page.evaluate(() => (window as any).__messages as { type: string; format?: string; convertedFrom?: string }[]);
    const waitFor = (type: string) => page.waitForFunction((wanted) => (window as any).__messages.some((message: { type: string }) => message.type === wanted), type, { timeout: 60_000 });

    await embed(`${PAGE}?file=/files/sample&save=/files/save&name=old.rtf&convert=1`);
    await waitFor('ready');
    expect((await messages())[0]).toMatchObject({ source: 'bymoffice', type: 'ready', format: 'docx', convertedFrom: 'rtf' });
    const frame = page.frames().find((candidate) => candidate.url().includes('/examples/html/'))!;
    await frame.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('New ');
    });
    await frame.locator('#save').click();
    await waitFor('saved');
    expect((await messages()).map((message) => [message.type, message.format, message.convertedFrom])).toEqual([
      ['ready', 'docx', 'rtf'],
      ['changed', 'docx', 'rtf'],
      ['saved', 'docx', 'rtf'],
    ]);

    // Without convert the file is what it was, and nothing says it was converted.
    await embed(`${PAGE}?file=/files/sample&name=old.rtf`);
    await waitFor('ready');
    const plain = (await messages())[0];
    expect([plain.type, plain.format, 'convertedFrom' in plain]).toEqual(['ready', 'rtf', false]);
  });

  test('Print: a document is printed with the button, in view mode too; under the print medium only its pages show; a workbook is printed as well', async ({ page }) => {
    for (const query of ['new=docx&lang=tr', 'new=docx&mode=view']) {
      await page.goto(`${PAGE}?${query}`);
      await ready(page);
      await expect(page.locator('#print')).toBeVisible();
      await expect(page.locator('#print')).toHaveText(query.includes('tr') ? 'Yazdır' : 'Print');
      // No dialog is opened from a test: the call is recorded, and the "dialog" stays open.
      await page.evaluate(() => {
        (window as any).__prints = 0;
        window.print = () => void (window as any).__prints++;
      });
      await page.locator('#print').click();
      await expect.poll(() => page.evaluate(() => (window as any).__prints)).toBe(1);
      await expect(page.locator('body > .bym-print img')).toHaveCount(1);
      // The page lays its body out as a column of full height: none of that reaches the paper.
      await page.emulateMedia({ media: 'print' });
      await expect(page.locator('.bym-print img')).toBeVisible();
      await expect(page.locator('#bar')).toBeHidden();
      await expect(page.locator('#editor')).toBeHidden();
      const box = await page.evaluate(() => {
        const rect = document.querySelector('.bym-print-sheet')!.getBoundingClientRect();
        return [rect.left, rect.top, getComputedStyle(document.body).display];
      });
      expect(box).toEqual([0, 0, 'block']);
      await page.emulateMedia({ media: 'screen' });
      await expect(page.locator('#bar')).toBeVisible();
      await page.evaluate(() => void window.dispatchEvent(new Event('afterprint')));
      await expect(page.locator('.bym-print')).toHaveCount(0);
      expect(await page.evaluate(() => (window as any).bymEditor.view.printing)).toBe(false);
    }
    // A workbook: an empty one has nothing to print and says so; with a cell in use it is printed, in view mode too.
    for (const query of ['new=xlsx', 'new=xlsx&mode=view']) {
      await page.goto(`${PAGE}?${query}`);
      await ready(page);
      await expect(page.locator('#print')).toBeVisible();
      await page.evaluate(() => {
        (window as any).__prints = 0;
        window.print = () => void (window as any).__prints++;
      });
      await page.locator('#print').click();
      await expect(page.locator('#status')).toHaveText('There is nothing to print');
      expect(await page.evaluate(() => (window as any).__prints)).toBe(0);
      await page.evaluate(() => {
        const workbook = (window as any).bymEditor.workbook;
        const row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 42, style: 0 }] };
        workbook.sheets[0].rows.set(0, row);
      });
      await page.locator('#print').click();
      await expect.poll(() => page.evaluate(() => (window as any).__prints)).toBe(1);
      await expect(page.locator('body > .bym-print img')).toHaveCount(1);
      await page.emulateMedia({ media: 'print' });
      await expect(page.locator('.bym-print img')).toBeVisible();
      await expect(page.locator('#editor')).toBeHidden();
      await page.emulateMedia({ media: 'screen' });
      await page.evaluate(() => void window.dispatchEvent(new Event('afterprint')));
      await expect(page.locator('.bym-print')).toHaveCount(0);
      expect(await page.evaluate(() => (window as any).bymEditor.view.printing)).toBe(false);
    }
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
