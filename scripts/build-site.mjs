// Builds the public demo site: site/ (the source) -> site-dist/ (what is published).
//
// The site is static and works under any path (GitHub Pages serves it under /bymoffice/): every
// address in it is relative, and every file it uses comes from its own origin. Nothing is fetched
// from another host: the engine, its fonts and the bridge page are copied from dist/ (run
// `pnpm build` first; `pnpm site` does both), and the sample files are made here, in Node, with
// the engine's own public API (synthetic content only).
//
// The usage examples are pages of their own (site/examples/*.html). Each runs as it is, in an
// iframe of the examples page, and the code shown beside it is made from the same file, with the
// local addresses of the package replaced by those of the CDN: what is shown is what runs.

import { createHash } from 'node:crypto';
import { cp, mkdir, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { dirname, join, relative, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { REPOSITORY_URL } from './site-config.mjs';
import { makeSamples } from './site-samples.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const SOURCE = join(ROOT, 'site');
const DIST = join(ROOT, 'dist');
const OUT = join(ROOT, 'site-dist');


const { version } = JSON.parse(await readFile(join(ROOT, 'package.json'), 'utf8'));
const CDN = `https://cdn.jsdelivr.net/npm/bymoffice@${version}/dist/`;

/** What the engine needs beside a page, as `pnpm build` makes it. */
const FROM_DIST = ['bymoffice.min.js', 'bymoffice.esm.js', 'bymoffice.webview.html', 'fonts'];

/**
 * The policy of every page of the site. The engine adds its own style element (and style
 * attributes), hence 'unsafe-inline' for styles; pictures of a document are blob: addresses it
 * makes from the bytes of the file; the examples page shows the examples in iframes of the site.
 */
const policy = (scripts) =>
  [
    "default-src 'self'",
    `script-src 'self'${scripts.map((hash) => ` '${hash}'`).join('')}`,
    "style-src 'self' 'unsafe-inline'",
    "img-src 'self' data: blob:",
    "font-src 'self' data:",
    "connect-src 'self' blob: data:",
    "frame-src 'self' blob:",
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; ');

const sha256 = (text) => `sha256-${createHash('sha256').update(text, 'utf8').digest('base64')}`;

/** Replaces each marker of the source, which must be there exactly once (`all`: at least once). */
function fill(text, marker, value, where, all = false) {
  const count = text.split(marker).length - 1;
  if (all ? count < 1 : count !== 1) throw new Error(`${where} must contain ${marker} ${all ? 'at least' : 'exactly'} once (found ${count})`);
  return text.split(marker).join(value);
}

/** The inline scripts of a page, as the browser hashes them (the text between the tags). */
const inlineScripts = (page) => [...page.matchAll(/<script(?:\s+type="module")?>([\s\S]*?)<\/script>/g)].map((match) => match[1]);

/**
 * An example page as it is shown: the addresses of the package on this site become those of the
 * CDN, the policy line (which the site adds) is left out.
 */
function shownCode(page) {
  return page
    .replace(/^[ \t]*<meta http-equiv="Content-Security-Policy"[^\n]*\n/m, '')
    .split('../bymoffice.min.js').join(`${CDN}bymoffice.min.js`)
    .split('../bymoffice.esm.js').join(`${CDN}bymoffice.esm.js`)
    .split("'../fonts/'").join(`'${CDN}fonts/'`);
}

/** Fails when a page asks for anything of another host (the CDN appears only as text in code). */
function checkOwnOrigin(source, where) {
  // Code shown on a page is text (escaped), not markup.
  const page = source.replace(/<pre\b[\s\S]*?<\/pre>/g, '');
  for (const match of page.matchAll(/\b(?:src|href|action|data|poster)\s*=\s*["']([^"']*)["']/gi)) {
    const value = match[1];
    if (value === REPOSITORY_URL || value.startsWith(`${REPOSITORY_URL}/`)) continue;
    if (/^[a-z][a-z0-9+.-]*:/i.test(value) || value.startsWith('//')) throw new Error(`${where}: "${value}" is not an address of the site itself`);
    if (value.startsWith('/')) throw new Error(`${where}: "${value}" is not relative; the site is served under a path of its own`);
  }
  if (/@import|url\(\s*["']?(?:https?:)?\/\//i.test(page)) throw new Error(`${where}: a style loads something from another host`);
}

for (const name of FROM_DIST) {
  try {
    await stat(join(DIST, name));
  } catch {
    throw new Error(`dist/${name} is missing: run "pnpm build" first (or "pnpm site", which does both)`);
  }
}

await rm(OUT, { recursive: true, force: true });
await mkdir(OUT, { recursive: true });

// The engine, its fonts and the bridge page, unchanged.
for (const name of FROM_DIST) await cp(join(DIST, name), join(OUT, name), { recursive: true });

// The pages and their files.
const text = await readFile(join(SOURCE, 'assets', 'i18n.js'), 'utf8');
const dictionary = new Function(`const window = {}; ${text}; return window.BYM_SITE_TEXT;`)();
const languages = Object.keys(dictionary);
if (languages.join() !== 'tr,en') throw new Error('site/assets/i18n.js must give the texts in Turkish and English');
const keys = (lang) => Object.keys(dictionary[lang]).sort();
const missing = keys('tr').filter((key) => !(key in dictionary.en)).concat(keys('en').filter((key) => !(key in dictionary.tr)));
if (missing.length) throw new Error(`site/assets/i18n.js: these texts are not in both languages: ${missing.join(', ')}`);

/** Every text key a page names exists in the dictionary. */
function checkKeys(page, where) {
  const named = [...page.matchAll(/data-i18n="([^"]+)"/g)].map((match) => match[1]);
  for (const match of page.matchAll(/data-i18n-attr="([^"]+)"/g)) for (const pair of match[1].split(';')) named.push(pair.split(':')[1]);
  for (const match of page.matchAll(/data-title="([^"]+)"/g)) named.push(match[1]);
  for (const key of named) if (!(key in dictionary.en)) throw new Error(`${where}: the text "${key}" is not in site/assets/i18n.js`);
}

const placeholders = (page, where) => {
  let out = page.split('__BYM_VERSION__').join(version).split('__BYM_REPOSITORY_URL__').join(REPOSITORY_URL).split('__BYM_CDN__').join(CDN);
  if (/__BYM_[A-Z_]+__/.test(out.replace('__BYM_CSP__', ''))) throw new Error(`${where}: an unknown placeholder ${/__BYM_[A-Z_]+__/.exec(out)[0]}`);
  return out;
};

async function copyTree(from, to) {
  await mkdir(to, { recursive: true });
  for (const entry of await readdir(from, { withFileTypes: true })) {
    const source = join(from, entry.name);
    const target = join(to, entry.name);
    const where = relative(ROOT, source);
    if (entry.isDirectory()) {
      await copyTree(source, target);
      continue;
    }
    if (entry.name.startsWith('.')) continue;
    if (!entry.name.endsWith('.html')) {
      await cp(source, target);
      continue;
    }
    const isExample = relative(SOURCE, source).startsWith('examples');
    const page = placeholders(await readFile(source, 'utf8'), where);
    checkOwnOrigin(page, where);
    if (!isExample) checkKeys(page, where);
    // Only the scripts of the page itself run: the examples' inline scripts by their hashes.
    const scripts = inlineScripts(page);
    if (!isExample && scripts.length) throw new Error(`${where}: the pages of the site keep their scripts in files`);
    await writeFile(target, fill(page, '__BYM_CSP__', policy(scripts.map(sha256)), where));
  }
}
await copyTree(SOURCE, OUT);

// The code of the examples, as the examples page shows it.
const examples = {};
for (const name of (await readdir(join(SOURCE, 'examples'))).filter((file) => file.endsWith('.html')).sort()) {
  const page = placeholders(await readFile(join(SOURCE, 'examples', name), 'utf8'), `site/examples/${name}`);
  examples[name.replace(/\.html$/, '')] = shownCode(page);
}
await writeFile(join(OUT, 'examples', 'code.json'), JSON.stringify(examples, null, 2));

// The sample files: made with the engine itself, from synthetic content.
const engine = await import(pathToFileURL(join(DIST, 'bymoffice.esm.js')).href);
await mkdir(join(OUT, 'samples'), { recursive: true });
for (const [name, bytes] of Object.entries(await makeSamples(engine))) await writeFile(join(OUT, 'samples', name), bytes);

// Nothing private and no address of this machine goes out with the site.
async function checkTree(dir) {
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) await checkTree(path);
    else if (/\.(html|js|css|json|svg|txt)$/.test(entry.name)) {
      const content = await readFile(path, 'utf8');
      if (content.includes(ROOT) || /\/Users\//.test(content)) throw new Error(`${relative(ROOT, path)} contains a local path`);
    }
  }
}
await checkTree(OUT);

console.log(`site: ${relative(ROOT, OUT)}/ (version ${version}, ${Object.keys(examples).length} examples)`);
