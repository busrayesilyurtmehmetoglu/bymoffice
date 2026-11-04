import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cpSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { join, relative, resolve, sep } from 'node:path';
import type { Page, Request } from '@playwright/test';
import { documentText, getCell, open } from '../src/index';
import { linesDocx } from './fixtures/docx';
import { expect, test } from './helpers';

// examples/vite, examples/react-vite and examples/vue-vite, as somebody who ran `npm i bymoffice` has them.
//
// Each example is copied to tmp-framework-examples/<key>/<name>/, with the package as it is published (`npm pack`
// of this repository, unpacked) in its own node_modules/bymoffice: that is what `import 'bymoffice'` and
// vite.config.js find there, before anything further up. Vite, React and Vue are found further up, in the
// node_modules of the repository. Then a production build of each, through Vite's own API, under a base that is
// where the test server serves it; the build says which modules went into it, and none may come from src/.
//
// The builds are made once for all the workers of a run: the first worker takes a lock (a folder) and builds,
// the others wait for it. The key is a hash of everything that goes in, so a run after a change builds again.

const ROOT = resolve('.');
const WORK = join(ROOT, 'tmp-framework-examples');
const EXAMPLES = ['vite', 'react-vite', 'vue-vite'] as const;
type Example = (typeof EXAMPLES)[number];
/** What each example opens when it is given no file. */
const BLANK: Record<Example, 'xlsx' | 'docx'> = { vite: 'xlsx', 'react-vite': 'docx', 'vue-vite': 'xlsx' };
const FONT_FILES = ['texgyrebonum-regular.otf', 'texgyrebonum-bold.otf', 'texgyrebonum-italic.otf', 'texgyrebonum-bolditalic.otf'];

/** The files of an example as they are checked in (what `npm install` / a build add is left out). */
function sourcesOf(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir).sort()) {
    if (name === 'node_modules' || name === 'dist' || name === 'bymoffice-fonts') continue;
    const path = join(dir, name);
    if (statSync(path).isDirectory()) sourcesOf(path, out);
    else out.push(path);
  }
  return out;
}

function keyOfInputs(): string {
  const hash = createHash('sha256');
  const files = [join(ROOT, 'package.json'), ...sourcesOf(join(ROOT, 'dist')), ...EXAMPLES.flatMap((name) => sourcesOf(join(ROOT, 'examples', name)))];
  for (const file of files) hash.update(relative(ROOT, file)).update('\0').update(readFileSync(file)).update('\0');
  hash.update(readFileSync(join(ROOT, 'node_modules', 'vite', 'package.json')));
  return hash.digest('hex').slice(0, 16);
}

const sleep = (ms: number) => new Promise((done) => setTimeout(done, ms));

/** Packs the package, gives every example a copy of it and builds them; resolves to the folder that holds them. */
async function prepare(): Promise<string> {
  if (!existsSync(join(ROOT, 'dist', 'bymoffice.esm.js'))) throw new Error('dist/ is missing: run pnpm build first');
  const key = keyOfInputs();
  const work = join(WORK, key);
  const done = join(work, 'done');
  const lock = join(WORK, `${key}-lock`);
  mkdirSync(WORK, { recursive: true });
  for (let waited = 0; !existsSync(done); waited += 250) {
    try {
      mkdirSync(lock);
    } catch {
      // Another worker builds. A lock older than five minutes is left from a run that was stopped.
      if (Date.now() - (statSync(lock, { throwIfNoEntry: false })?.mtimeMs ?? Date.now()) > 300_000) rmSync(lock, { recursive: true, force: true });
      if (waited > 300_000) throw new Error('the examples were not built in five minutes');
      await sleep(250);
      continue;
    }
    try {
      if (!existsSync(done)) await buildAll(work);
      writeFileSync(done, new Date().toISOString());
    } finally {
      rmSync(lock, { recursive: true, force: true });
    }
  }
  return work;
}

async function buildAll(work: string): Promise<void> {
  // What an earlier run built, for these inputs or others (only the lock this worker holds stays).
  for (const name of readdirSync(WORK)) if (!name.endsWith('-lock')) rmSync(join(WORK, name), { recursive: true, force: true });
  mkdirSync(join(work, 'package'), { recursive: true });
  const packed = JSON.parse(execFileSync('npm', ['pack', '--json', '--ignore-scripts', '--pack-destination', work], { cwd: ROOT, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] }))[0];
  execFileSync('tar', ['-xzf', join(work, packed.filename), '-C', join(work, 'package')]);
  const unpacked = join(work, 'package', 'package');
  const { build } = await import('vite');
  for (const name of EXAMPLES) {
    const app = join(work, name);
    cpSync(join(ROOT, 'examples', name), app, { recursive: true, filter: (from) => !/[\\/](node_modules|dist|bymoffice-fonts)$/.test(from) });
    // As `npm install` leaves it: the package in node_modules of the app.
    cpSync(unpacked, join(app, 'node_modules', 'bymoffice'), { recursive: true });
    const result = await build({
      root: app,
      configFile: join(app, 'vite.config.js'),
      base: `/${relative(ROOT, app).split(sep).join('/')}/dist/`,
      logLevel: 'warn',
      // The engine is one module of about 3 MB: more than Vite's warning limit, which says so in every run.
      build: { outDir: join(app, 'dist'), emptyOutDir: true, chunkSizeWarningLimit: 4096 },
    });
    const outputs = (Array.isArray(result) ? result : [result]).flatMap((one: any) => one.output ?? []);
    const modules: string[] = outputs.flatMap((chunk: any) => (chunk.type === 'chunk' ? chunk.moduleIds ?? Object.keys(chunk.modules ?? {}) : []));
    const engine = join(app, 'node_modules', 'bymoffice', 'dist', 'bymoffice.esm.js');
    if (!modules.includes(engine)) throw new Error(`${name}: the build does not hold the packed package (${engine})`);
    const fromSource = modules.filter((id) => id.startsWith(join(ROOT, 'src') + sep) || id.startsWith(join(ROOT, 'dist') + sep));
    if (fromSource.length) throw new Error(`${name}: the build holds modules of the repository: ${fromSource.slice(0, 3).join(', ')}`);
    for (const file of FONT_FILES) {
      if (!existsSync(join(app, 'dist', 'bymoffice-fonts', 'tex-gyre-bonum', file))) throw new Error(`${name}: the build lacks bymoffice-fonts/tex-gyre-bonum/${file}`);
    }
  }
}

let work = '';

test.beforeAll(async ({}, info) => {
  info.setTimeout(360_000);
  work = await prepare();
});

/** Where the test server serves the build of an example (it gives no index.html for a folder: pages are named). */
const builtAt = (name: Example) => `/${relative(ROOT, join(work, name)).split(sep).join('/')}/dist/`;

/** Every request of the page, and every answer it got to one. */
function traffic(page: Page): { requests: Request[]; answers: Map<string, number> } {
  const requests: Request[] = [];
  const answers = new Map<string, number>();
  page.on('request', (request) => requests.push(request));
  page.on('response', (response) => answers.set(response.url(), response.status()));
  return { requests, answers };
}

/** No request went anywhere but to where the page came from (a CDN would be another origin). */
function expectOneOrigin(page: Page, requests: Request[]): void {
  const origin = new URL(page.url()).origin;
  const outside = requests.map((request) => request.url()).filter((url) => !url.startsWith(`${origin}/`) && !url.startsWith('data:') && !url.startsWith('blob:'));
  expect(outside, 'requests to other origins').toEqual([]);
}

const editors = (page: Page) => page.locator('.bym-editor').count();

/**
 * Records every editor the page makes. An editor is a `.bym-editor` element that mount() appends to its container
 * and destroy() takes out again; one that is never destroyed keeps its place, also in a container that has left
 * the page (a component that was unmounted while mount() was under way), where counting the page cannot see it.
 */
async function watchEditors(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const made: Element[] = ((window as any).__editorsMade = []);
    const append = Node.prototype.appendChild;
    Node.prototype.appendChild = function <T extends Node>(this: Node, node: T): T {
      if (node instanceof Element && node.classList.contains('bym-editor') && made.indexOf(node) < 0) made.push(node);
      return append.call(this, node) as T;
    };
  });
}

/** How many editors were made, and how many of them were not destroyed. */
const editorsMade = (page: Page) =>
  page.evaluate(() => {
    const made: Element[] = (window as any).__editorsMade;
    return { made: made.length, alive: made.filter((root) => root.parentNode !== null).length };
  });

/** The example shows exactly one editor and holds no other, and keeps it so for a while. */
async function expectOneEditor(page: Page): Promise<void> {
  await page.waitForFunction(() => (window as any).bymEditor !== undefined, undefined, { timeout: 30_000 });
  await expect.poll(() => editors(page)).toBe(1);
  // An editor whose mount() was overtaken arrives late, and is to be destroyed then: give it the time to arrive.
  await page.waitForTimeout(1000);
  expect(await editors(page)).toBe(1);
  expect((await editorsMade(page)).alive, 'editors that were made and not destroyed').toBe(1);
}

/** Types where the editor of the page is: into cell A1 of a workbook, at the start of a document. */
async function typeInto(page: Page, kind: 'xlsx' | 'docx', text: string, isMobile: boolean): Promise<void> {
  if (kind === 'xlsx') {
    const input = page.locator('.bym-input');
    await input.focus();
    await expect(input).toBeFocused();
    await page.keyboard.type(text);
    await page.keyboard.press('Enter');
    return;
  }
  const sheet = page.locator('.bym-doc-page[data-page="1"]');
  await expect(sheet).toBeVisible();
  if (isMobile) await sheet.tap({ position: { x: 6, y: 6 } });
  else await sheet.click({ position: { x: 6, y: 6 } });
  await expect(page.locator('.bym-doc-input')).toBeFocused();
  await page.keyboard.type(text);
}

/** What the editor of the page saves, read again here: the text of a document, A1 of a workbook. */
async function savedText(page: Page): Promise<{ format: string; text: string }> {
  const bytes = await page.evaluate(async () => Array.from(await (window as any).bymEditor.save()));
  const file = await open(new Uint8Array(bytes as number[]));
  if (file.kind === 'document') return { format: file.format, text: documentText(await file.document!()) };
  const cell = getCell((await file.workbook()).sheets[0], 0, 0);
  return { format: file.format, text: cell ? String(cell.value) : '' };
}

/** Whether the device has the family itself: then nothing is fetched for it. */
function deviceHas(page: Page, family: string): Promise<boolean> {
  return page.evaluate((name) => {
    const context = document.createElement('canvas').getContext('2d')!;
    for (const generic of ['monospace', 'serif']) {
      context.font = `72px ${generic}`;
      const plain = context.measureText('mmmmmmmmmmlliWW@#01').width;
      context.font = `72px "${name}", ${generic}`;
      if (context.measureText('mmmmmmmmmmlliWW@#01').width !== plain) return true;
    }
    return false;
  }, family);
}

for (const name of EXAMPLES) {
  test.describe(`examples/${name}, built with Vite from the packed package`, () => {
    test('the editor mounts, takes what is typed and saves it', async ({ page, isMobile }) => {
      const { requests } = traffic(page);
      await watchEditors(page);
      await page.goto(`${builtAt(name)}index.html`);
      await expectOneEditor(page);
      await typeInto(page, BLANK[name], 'Typed 42', isMobile);
      expect(await savedText(page)).toEqual({ format: BLANK[name], text: 'Typed 42' });
      expectOneOrigin(page, requests);
    });

    test('a document in Bookman Old Style gets the font of the package from the folder of the example', async ({ page }) => {
      const { requests, answers } = traffic(page);
      const base = builtAt(name);
      const bytes = await linesDocx('Bookman Old Style', 24, 3);
      await page.route((url) => url.pathname === `${base}files/bookman.docx`, (route) => route.fulfill({ status: 200, contentType: 'application/octet-stream', body: Buffer.from(bytes) }));
      await watchEditors(page);
      await page.goto(`${base}index.html?file=${encodeURIComponent(`${base}files/bookman.docx`)}`);
      await expectOneEditor(page);
      test.skip(await deviceHas(page, 'Bookman Old Style'), 'this device has Bookman Old Style: nothing is fetched for it');
      const origin = new URL(page.url()).origin;
      const wanted = FONT_FILES.map((file) => `${origin}${base}bymoffice-fonts/tex-gyre-bonum/${file}`);
      await expect.poll(() => wanted.map((url) => answers.get(url) ?? 0), { timeout: 15_000 }).toEqual([200, 200, 200, 200]);
      // The font is used once all four faces are in the page.
      await expect
        .poll(() => page.evaluate(() => Array.from(document.fonts as unknown as Iterable<FontFace>).filter((face) => face.family.replace(/["']/g, '') === 'TeX Gyre Bonum' && face.status === 'loaded').length))
        .toBe(4);
      expect(requests.filter((request) => /\.otf(\?|$)/.test(request.url())).map((request) => request.url()).sort()).toEqual([...wanted].sort());
      expectOneOrigin(page, requests);
    });

    test('closing and opening the editor again leaves one editor, also when it is done quickly', async ({ page, isMobile }) => {
      await watchEditors(page);
      await page.goto(`${builtAt(name)}index.html`);
      await expectOneEditor(page);
      const toggle = page.getByRole('button', { name: /editor$/ });
      await toggle.click();
      await expect(toggle).toHaveText('Open editor');
      await expect.poll(() => editors(page)).toBe(0);
      expect(await editorsMade(page)).toEqual({ made: 1, alive: 0 });
      expect(await page.evaluate(() => (window as any).bymEditor)).toBeUndefined();
      await toggle.click();
      await expectOneEditor(page);
      // Closed and opened again twice before the editors that were asked for have arrived: one turn of the event
      // loop between the clicks, for the framework to apply each of them (Vue updates the page in a microtask).
      await page.evaluate(async () => {
        const button = document.querySelector('header button') as HTMLButtonElement;
        for (let i = 0; i < 4; i++) {
          button.click();
          await new Promise((done) => setTimeout(done, 0));
        }
      });
      // Three editors were asked for in all after the first: all made, two destroyed as they arrived.
      await expect.poll(async () => (await editorsMade(page)).made).toBeGreaterThanOrEqual(3);
      await expect(toggle).toHaveText('Close editor');
      await expectOneEditor(page);
      await typeInto(page, BLANK[name], 'After', isMobile);
      expect((await savedText(page)).text).toBe('After');
    });
  });
}

test.describe('examples/react-vite in the Vite dev server', () => {
  test('StrictMode mounts the component twice, and one editor stays and works', async ({ page, isMobile }, info) => {
    info.setTimeout(120_000);
    const { createServer } = await import('vite');
    // A copy of its own: every dev server copies the font folder into public/ as it starts (vite.config.js), and two
    // of them copying into the same folder at once trip over each other's files.
    const app = join(work, `dev-${info.project.name}-${info.workerIndex}-${info.repeatEachIndex}`);
    rmSync(app, { recursive: true, force: true });
    cpSync(join(work, 'react-vite'), app, { recursive: true, filter: (from) => from !== join(work, 'react-vite', 'dist') });
    const server = await createServer({
      root: app,
      configFile: join(app, 'vite.config.js'),
      cacheDir: join(app, 'vite-cache'),
      logLevel: 'warn',
      server: { host: '127.0.0.1', port: 0, strictPort: false, watch: null, hmr: false },
    });
    try {
      await server.listen();
      const url = server.resolvedUrls!.local[0];
      const { requests } = traffic(page);
      await watchEditors(page);
      await page.goto(url);
      await expectOneEditor(page);
      // The development build of React with StrictMode: the effect ran twice, both runs made an editor, one is left.
      expect(await editorsMade(page)).toEqual({ made: 2, alive: 1 });
      await typeInto(page, 'docx', 'Dev', isMobile);
      expect(await savedText(page)).toEqual({ format: 'docx', text: 'Dev' });
      expectOneOrigin(page, requests);
    } finally {
      await server.close();
      rmSync(app, { recursive: true, force: true });
    }
  });
});
