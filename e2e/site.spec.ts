import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import type { Download, Frame, Page } from '@playwright/test';
import { REPOSITORY_URL } from '../scripts/site-config.mjs';
import { expect, test } from './helpers';

// The public demo site (site/ -> site-dist/ by scripts/build-site.mjs; `pnpm e2e` builds it),
// served under /bymoffice/ as GitHub Pages serves it: every page in both languages, the live
// editor, the samples, and each usage example as it runs.

const SITE = '/bymoffice/';
const OUT = new URL('../site-dist/', import.meta.url);

/** The texts of the site, as its pages read them. */
async function dictionary(): Promise<Record<'tr' | 'en', Record<string, string>>> {
  const source = await readFile(new URL('../site/assets/i18n.js', import.meta.url), 'utf8');
  return new Function(`const window = {}; ${source}; return window.BYM_SITE_TEXT;`)();
}

const PAGES = [
  { path: '', title: 'home.title', heading: 'home.tagline' },
  { path: 'editor.html', title: 'editor.title', heading: 'editor.heading' },
  { path: 'examples.html', title: 'examples.title', heading: 'examples.heading' },
];
const EXAMPLES = ['script', 'module', 'viewer', 'save', 'autosave', 'collab', 'pdf', 'bridge', 'locale'];

test.beforeAll(() => {
  if (!existsSync(new URL('index.html', OUT))) throw new Error('site-dist/ is missing: run "node scripts/build-site.mjs" after "pnpm build" (pnpm e2e does both)');
});

/** The bytes of a download. */
async function bytesOf(download: Download): Promise<Buffer> {
  const path = await download.path();
  if (!path) throw new Error('the download has no file');
  return readFile(path);
}

/** Opens bytes with the engine of the page and says what is in them. */
function contentOf(page: Page | Frame, bytes: Buffer): Promise<{ format: string; text: string }> {
  return page.evaluate(async (data) => {
    const api = (window as any).BymOffice;
    const file = await api.open(new Uint8Array(data));
    if (file.kind === 'document') return { format: file.format, text: api.documentText(await file.document()) };
    const workbook = await file.workbook();
    const cell = api.getCell(workbook.sheets[0], 0, 0);
    return { format: file.format, text: `${workbook.sheets[0].name}!A1=${cell ? cell.value : ''}` };
  }, Array.from(bytes));
}

/** In an example's page: it has opened its file (the editors are mounted; the bridge page said "loaded"). */
function exampleReady(): boolean {
  const mounted = (window as any).mounted;
  if (mounted) return !Array.isArray(mounted) || mounted.length === 2;
  const log = document.getElementById('log');
  return !!log && (log.textContent || '').includes('"type":"loaded"');
}

/**
 * Waits until every example that runs in an iframe of the page is in `lang` and has opened its
 * file. (A frame taken away while its file is being read leaves an error in WebKit's console.)
 */
async function examplesSettled(page: Page, lang: string): Promise<void> {
  await expect
    .poll(
      async () => {
        const frames = page.frames().filter((frame) => /\/examples\/\w+\.html/.test(frame.url()));
        if (frames.length === 0 || frames.some((frame) => !frame.url().endsWith(`lang=${lang}`))) return false;
        const ready = await Promise.all(frames.map((frame) => frame.evaluate(exampleReady).catch(() => false)));
        return ready.every(Boolean);
      },
      { message: `the examples run in ${lang}`, timeout: 60_000 },
    )
    .toBe(true);
}

const editorReady = (page: Page) => page.waitForFunction(() => (window as any).bymEditor !== undefined, undefined, { timeout: 60_000 });

/** The page is not wider than the screen. */
async function expectNoSideScroll(page: Page, where: string): Promise<void> {
  const sizes = await page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, view: document.documentElement.clientWidth }));
  expect(sizes.scroll, `${where}: the page scrolls sideways (${sizes.scroll} > ${sizes.view})`).toBeLessThanOrEqual(sizes.view);
}

test.describe('demo site: pages and languages', () => {
  test('every page shows in Turkish and in English: lang, title, heading; the choice is remembered', async ({ page }) => {
    const text = await dictionary();
    for (const entry of PAGES) {
      await page.goto(SITE + entry.path);
      const examples = entry.path === 'examples.html';
      for (const lang of ['tr', 'en'] as const) {
        if (examples) await examplesSettled(page, await page.locator('#lang').inputValue());
        await page.selectOption('#lang', lang);
        await expect(page.locator('html')).toHaveAttribute('lang', lang);
        await expect(page).toHaveTitle(text[lang][entry.title]);
        await expect(page.locator('h1')).toHaveText(text[lang][entry.heading]);
        // Every text of the page is in the dictionary of that language.
        const unknown = await page.evaluate((all) => {
          const missing: string[] = [];
          document.querySelectorAll('[data-i18n]').forEach((node) => {
            const key = node.getAttribute('data-i18n')!;
            if (node.textContent !== all[key]) missing.push(key);
          });
          return missing;
        }, text[lang]);
        expect(unknown, `${entry.path || 'index'} in ${lang}`).toEqual([]);
      }
      // Turkish again, then a reload: the page comes back in Turkish.
      if (examples) await examplesSettled(page, 'en');
      await page.selectOption('#lang', 'tr');
      if (examples) await examplesSettled(page, 'tr');
      await page.reload();
      await expect(page.locator('html')).toHaveAttribute('lang', 'tr');
      if (examples) await examplesSettled(page, 'tr');
      await page.selectOption('#lang', 'en');
      if (examples) await examplesSettled(page, 'en');
    }
  });

  test('the home page gives the one-line installation and links to the demos and the repository', async ({ page }) => {
    await page.goto(SITE);
    await expect(page.locator('#install-cdn')).toHaveText('<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>');
    await expect(page.locator('#install-npm')).toHaveText('npm i bymoffice');
    await expect(page.locator('main a[href="editor.html"]')).toHaveCount(1);
    await expect(page.locator('main a[href="examples.html"]')).toHaveCount(1);
    await expect(page.locator(`nav a[href="${REPOSITORY_URL}"]`)).toHaveCount(1);
    // No product other than this one is named on the page.
    const words = (await page.locator('body').innerText()).toLowerCase();
    for (const name of ['onlyoffice', 'sheetjs', 'google', 'microsoft', 'libreoffice', 'collabora', 'syncfusion', 'univer']) expect(words).not.toContain(name);
  });

  test('nothing is fetched from another host', async ({ page, baseURL }) => {
    const own = new URL(baseURL!).host;
    const others: string[] = [];
    page.on('request', (request) => {
      const url = new URL(request.url());
      if ((url.protocol === 'http:' || url.protocol === 'https:') && url.host !== own) others.push(request.url());
    });
    for (const entry of PAGES) await page.goto(SITE + entry.path);
    // The examples page with its first examples running, and the editor with a sample open.
    await page.goto(`${SITE}examples.html`);
    await expect(page.locator('body')).toHaveAttribute('data-code', 'ready');
    await examplesSettled(page, 'en');
    await page.goto(`${SITE}editor.html`);
    await page.click('#sample-doc');
    await editorReady(page);
    expect(others).toEqual([]);
  });
});

test.describe('demo site: English first, even in a Turkish browser', () => {
  test.use({ locale: 'tr-TR' });
  test('the first visit is in English; ?lang=tr opens Turkish and is remembered', async ({ page }) => {
    await page.goto(SITE);
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
    await expect(page.locator('#lang')).toHaveValue('en');

    await page.goto(`${SITE}?lang=tr`);
    await expect(page.locator('html')).toHaveAttribute('lang', 'tr');
    await expect(page.locator('#lang')).toHaveValue('tr');
    await page.goto(`${SITE}examples.html`);
    await expect(page.locator('html')).toHaveAttribute('lang', 'tr');

    await page.goto(`${SITE}?lang=en`);
    await expect(page.locator('html')).toHaveAttribute('lang', 'en');
  });
});

test.describe('demo site: the live editor', () => {
  test('says that the file stays in the browser', async ({ page }) => {
    await page.goto(`${SITE}editor.html`);
    await page.selectOption('#lang', 'tr');
    await expect(page.locator('#privacy')).toContainText('hiçbir yere gönderilmez');
    await page.selectOption('#lang', 'en');
    await expect(page.locator('#privacy')).toContainText('not sent anywhere');
  });

  test('a blank workbook is written in and downloaded; the bytes open with the value', async ({ page }) => {
    await page.goto(`${SITE}editor.html`);
    await page.selectOption('#lang', 'en');
    await page.click('#blank-sheet');
    await editorReady(page);
    await page.evaluate(() => (window as any).bymEditor.view.setCellInput(0, 0, '42'));
    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
    expect(download.suggestedFilename()).toBe('workbook.xlsx');
    expect(await contentOf(page, await bytesOf(download))).toEqual({ format: 'xlsx', text: 'Sheet1!A1=42' });
    await expect(page.locator('#status')).toContainText('workbook.xlsx');
  });

  test('a blank document is written in and downloaded; the bytes open with the text', async ({ page }) => {
    await page.goto(`${SITE}editor.html`);
    await page.selectOption('#lang', 'tr');
    await page.click('#blank-doc');
    await editorReady(page);
    await page.evaluate(() => {
      const view = (window as any).bymEditor.view;
      view.setSelection({ story: 'body', path: [0], offset: 0 });
      view.insertText('Merhaba dünya');
    });
    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#download')]);
    expect(download.suggestedFilename()).toBe('belge.docx');
    expect(await contentOf(page, await bytesOf(download))).toEqual({ format: 'docx', text: 'Merhaba dünya' });
  });

  test('the sample workbook has formulas, formats and charts; it becomes a PDF; read-only turns editing off', async ({ page }) => {
    await page.goto(`${SITE}editor.html`);
    await page.selectOption('#lang', 'tr');
    await page.click('#sample-sheet');
    await editorReady(page);
    const facts = await page.evaluate(() => {
      const w = window as any;
      const sheet = w.bymEditor.workbook.sheets[0];
      const cell = (row: number, col: number) => w.BymOffice.getCell(sheet, row, col);
      return { name: sheet.name, profit: cell(7, 4).value, formula: cell(7, 4).formula.text, charts: w.bymEditor.view.charts().length, shown: w.BymOffice.cellDisplayText(w.bymEditor.workbook, cell(1, 2), w.BymOffice.LOCALE_TR) };
    });
    expect(facts).toEqual({ name: 'Satışlar', profit: 3900, formula: 'SUM(E2:E7)', charts: 2, shown: '1.200,00' });
    const [pdf] = await Promise.all([page.waitForEvent('download'), page.click('#pdf')]);
    expect(pdf.suggestedFilename()).toBe('ornek-tablo.pdf');
    expect((await bytesOf(pdf)).subarray(0, 5).toString('latin1')).toBe('%PDF-');
    await page.check('#read-only');
    expect(await page.evaluate(() => (window as any).bymEditor.view.readOnly)).toBe(true);
    await page.uncheck('#read-only');
    expect(await page.evaluate(() => (window as any).bymEditor.view.readOnly)).toBe(false);
  });

  test('the sample document opens, from the button and as a chosen file', async ({ page }) => {
    await page.goto(`${SITE}editor.html`);
    await page.selectOption('#lang', 'en');
    await page.click('#sample-doc');
    await editorReady(page);
    const table = () => page.evaluate(() => (window as any).bymEditor.document.body.blocks.some((block: { type: string }) => block.type === 'table'));
    expect(await table()).toBe(true);
    // The same file through the file field: what is chosen is opened in the page.
    await page.evaluate(() => ((window as any).bymEditor = undefined));
    await page.setInputFiles('#pick', { name: 'chosen.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', buffer: await readFile(new URL('samples/sample-doc-en.docx', OUT)) });
    await editorReady(page);
    await expect(page.locator('#status')).toHaveText('chosen.docx is open.');
    expect(await table()).toBe(true);
    expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).bymEditor.document))).toContain('Sample document');
  });
});

test.describe('demo site: the usage examples', () => {
  test('the code shown is the code that runs, with the CDN addresses', async ({ page }) => {
    await page.goto(`${SITE}examples.html`);
    await page.selectOption('#lang', 'en');
    await expect(page.locator('body')).toHaveAttribute('data-code', 'ready');
    const code = JSON.parse(await readFile(new URL('examples/code.json', OUT), 'utf8')) as Record<string, string>;
    expect(Object.keys(code).sort()).toEqual([...EXAMPLES].sort());
    for (const id of EXAMPLES) {
      await expect(page.locator(`#code-${id}`)).toHaveText(code[id]);
      // The file the iframe runs is that code, but for the addresses of the package and the policy line.
      const running = await (await page.request.get(`${SITE}examples/${id}.html`)).text();
      const shown = running
        .replace(/^[ \t]*<meta http-equiv="Content-Security-Policy"[^\n]*\n/m, '')
        .split('../bymoffice.min.js').join('https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js')
        .split('../bymoffice.esm.js').join('https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.esm.js')
        .split("'../fonts/'").join("'https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/fonts/'");
      expect(shown, id).toBe(code[id]);
      await expect(page.locator(`#ex-${id} iframe`)).toHaveAttribute('data-src', `examples/${id}.html?lang=en`);
    }
    // The examples start when they come near the screen.
    await expect(page.locator('#ex-script iframe')).toHaveAttribute('src', 'examples/script.html?lang=en');
    await page.locator('#ex-locale iframe').scrollIntoViewIfNeeded();
    await expect(page.locator('#ex-locale iframe')).toHaveAttribute('src', 'examples/locale.html?lang=en');
    await examplesSettled(page, 'en');
    expect(code.script).toContain('<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>');
    expect(code.module).toContain("from 'https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.esm.js'");
    // The examples follow the language of the site.
    await page.selectOption('#lang', 'tr');
    await expect(page.locator('#ex-script iframe')).toHaveAttribute('src', 'examples/script.html?lang=tr');
    await examplesSettled(page, 'tr');
  });

  const mounted = (page: Page) => page.waitForFunction(() => (window as any).mounted !== undefined && (!Array.isArray((window as any).mounted) || (window as any).mounted.length === 2), undefined, { timeout: 60_000 });

  test('one script tag, an ES module and a read-only viewer', async ({ page }) => {
    await page.goto(`${SITE}examples/script.html?lang=tr`);
    await mounted(page);
    expect(await page.evaluate(() => [(window as any).mounted.kind, (window as any).mounted.workbook.sheets[0].name])).toEqual(['spreadsheet', 'Satışlar']);
    await page.goto(`${SITE}examples/module.html?lang=en`);
    await mounted(page);
    expect(await page.evaluate(() => (window as any).mounted.kind)).toBe('document');
    await page.goto(`${SITE}examples/viewer.html?lang=en`);
    await mounted(page);
    expect(await page.evaluate(() => (window as any).mounted.view.readOnly)).toBe(true);
    expect(await page.evaluate(() => (window as any).mounted.view.insertText('x'))).toBe(false);
  });

  test('saving to a function gives a link to the saved file; saving by itself saves after a change', async ({ page }) => {
    await page.goto(`${SITE}examples/save.html?lang=en`);
    await mounted(page);
    await page.evaluate(() => (window as any).mounted.view.setCellInput(0, 0, 'Saved here'));
    await expect(page.locator('#state')).toHaveText(/Not saved/);
    expect(await page.evaluate(() => (window as any).mounted.saveNow())).toBe(true);
    await expect(page.locator('#state')).toHaveText('Saved');
    await expect(page.locator('#link')).toBeVisible();
    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#link')]);
    expect(download.suggestedFilename()).toBe('saved.xlsx');
    expect(await contentOf(page, await bytesOf(download))).toEqual({ format: 'xlsx', text: 'Sales!A1=Saved here' });

    await page.goto(`${SITE}examples/autosave.html?lang=en`);
    await mounted(page);
    await page.evaluate(() => (window as any).mounted.view.setCellInput(1, 1, '7'));
    await expect.poll(() => page.evaluate(() => (window as any).saves.filter((save: { auto: boolean }) => save.auto).length), { timeout: 20_000 }).toBeGreaterThan(0);
    await expect(page.locator('#state')).toHaveText(/^Saved /);
  });

  test('editing together: what is typed in one editor appears in the other', async ({ page }) => {
    await page.goto(`${SITE}examples/collab.html?lang=en`);
    await mounted(page);
    await expect.poll(() => page.evaluate(() => (window as any).mounted.map((m: any) => m.collab.state).join())).toBe('online,online');
    await page.evaluate(() => (window as any).mounted[0].view.setCellInput(2, 1, 'from Ayşe'));
    await expect
      .poll(() => page.evaluate(() => {
        const w = window as any;
        const cell = w.BymOffice.getCell(w.mounted[1].workbook.sheets[0], 2, 1);
        return cell ? cell.value : null;
      }))
      .toBe('from Ayşe');
    expect(await page.evaluate(() => (window as any).mounted[1].collab.people.map((p: { name: string }) => p.name).sort())).toEqual(['Ayşe', 'Deniz']);
  });

  test('a PDF is made and downloaded', async ({ page }) => {
    await page.goto(`${SITE}examples/pdf.html?lang=en`);
    await mounted(page);
    await expect(page.locator('#pdf')).toBeEnabled();
    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#pdf')]);
    expect(download.suggestedFilename()).toBe('document.pdf');
    expect((await bytesOf(download)).subarray(0, 5).toString('latin1')).toBe('%PDF-');
  });

  test('the bridge page in an iframe answers ready and loaded, and saves', async ({ page }) => {
    await page.goto(`${SITE}examples/bridge.html?lang=en`);
    await expect(page.locator('#log')).toContainText('← {"type":"ready"', { timeout: 60_000 });
    await expect(page.locator('#log')).toContainText('← {"type":"loaded","id":"1","format":"xlsx"', { timeout: 60_000 });
    await page.click('#doc');
    await expect(page.locator('#log')).toContainText('"type":"loaded","id":"2","format":"docx"', { timeout: 60_000 });
    await page.click('#save');
    await expect(page.locator('#log')).toContainText('← {"type":"saved","id":"3"');
    await expect(page.locator('#link')).toBeVisible();
    const [download] = await Promise.all([page.waitForEvent('download'), page.click('#link')]);
    expect(download.suggestedFilename()).toBe('document.docx');
    // The page of the app has no engine; the bridge page has one.
    const bridge = page.frame({ url: /bymoffice\.webview\.html$/ });
    expect(bridge).not.toBeNull();
    expect((await contentOf(bridge!, await bytesOf(download))).format).toBe('docx');
  });

  test('the language of an editor changes with setLocale', async ({ page }) => {
    await page.goto(`${SITE}examples/locale.html?lang=en`);
    await mounted(page);
    const labels = () => page.evaluate(() => Array.from(document.querySelectorAll('#editor [aria-label]'), (node) => node.getAttribute('aria-label')).join('|'));
    await expect(page.locator('[data-locale="en"]')).toHaveAttribute('aria-pressed', 'true');
    const english = await labels();
    await page.click('[data-locale="tr"]');
    await expect(page.locator('[data-locale="tr"]')).toHaveAttribute('aria-pressed', 'true');
    await expect.poll(labels).not.toBe(english);
    await page.click('[data-locale="en"]');
    await expect.poll(labels).toBe(english);
  });
});

test.describe('demo site: phone width', () => {
  test('no page scrolls sideways', async ({ page, isMobile }) => {
    test.skip(!isMobile, 'phone profiles only');
    for (const entry of PAGES) {
      await page.goto(SITE + entry.path);
      await expectNoSideScroll(page, entry.path || 'index');
      if (entry.path === 'examples.html') await examplesSettled(page, 'en');
    }
    await page.goto(`${SITE}editor.html`);
    await page.click('#sample-sheet');
    await editorReady(page);
    await expectNoSideScroll(page, 'editor with a workbook');
    for (const id of EXAMPLES) {
      await page.goto(`${SITE}examples/${id}.html?lang=en`);
      await expect.poll(() => page.evaluate(exampleReady), { message: `example ${id} runs`, timeout: 60_000 }).toBe(true);
      await expectNoSideScroll(page, `example ${id}`);
    }
  });
});
