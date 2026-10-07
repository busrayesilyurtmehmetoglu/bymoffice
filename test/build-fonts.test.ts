import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { EMBEDDED_FONT_FACES } from '../src/render/embedded-fonts';

// What the build does with the font of the package: the files are copied
// next to the scripts as they are, and the bridge page carries them inside
// itself. Built into folders of its own (tmp-build-fonts/), so that dist/ is
// left alone.

const SOURCE = 'assets/fonts/tex-gyre-bonum';
const LICENSE = 'GUST-FONT-LICENSE.txt';
const OUT = 'tmp-build-fonts';

function build(name: string, env: { [key: string]: string }): string {
  const out = join(OUT, name);
  rmSync(out, { recursive: true, force: true });
  execFileSync(process.execPath, ['scripts/build.mjs'], { env: { ...process.env, BYM_NO_FONTS: '', ...env, BYM_OUT_DIR: out }, stdio: ['ignore', 'pipe', 'pipe'] });
  return out;
}

const sha = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

/** The font files a bridge page carries, by file name. */
function blocks(page: string): Map<string, Buffer> {
  const out = new Map<string, Buffer>();
  const pattern = /<script type="text\/plain" data-bym-font="([^"]+)">([A-Za-z0-9+/=]*)<\/script>/g;
  for (let match = pattern.exec(page); match; match = pattern.exec(page)) out.set(match[1], Buffer.from(match[2], 'base64'));
  return out;
}

/** The inline scripts the policy of a page allows, against the ones it has. */
function policyHolds(page: string): boolean {
  const policy = (/http-equiv="Content-Security-Policy" content="([^"]*)"/.exec(page) || [])[1] || '';
  const scripts = [...page.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((match) => `'sha256-${createHash('sha256').update(match[1], 'utf8').digest('base64')}'`);
  return scripts.length === 2 && scripts.every((hash) => policy.indexOf(hash) >= 0);
}

test('build: the font files are next to the scripts, byte for byte, and inside the bridge page; BYM_NO_FONTS leaves them out of the page', () => {
  const out = build('with', {});
  const files = EMBEDDED_FONT_FACES.map((face) => face.file);

  // dist/fonts: the four files and the licence, nothing else, nothing changed.
  const copied = join(out, 'fonts', 'tex-gyre-bonum');
  assert.deepEqual(readdirSync(copied).sort(), [...files, LICENSE].sort());
  assert.deepEqual(readdirSync(SOURCE).sort(), [...files, LICENSE].sort());
  for (const name of [...files, LICENSE]) {
    assert.ok(readFileSync(join(copied, name)).equals(readFileSync(join(SOURCE, name))), `${name} was changed on its way`);
  }

  // The scripts do not carry the fonts.
  const script = readFileSync(join(out, 'bymoffice.min.js'), 'utf8');
  const module = readFileSync(join(out, 'bymoffice.esm.js'), 'utf8');
  const regular = readFileSync(join(SOURCE, files[0]));
  const piece = regular.toString('base64').slice(1000, 1200);
  assert.equal(script.indexOf(piece), -1);
  assert.equal(module.indexOf(piece), -1);
  // Neither output names import.meta: both must parse as ES2019 (old bundlers, test runners that turn modules into
  // CommonJS). The script finds its folder by document.currentScript; a module is told where it is.
  for (const text of [script, module]) {
    assert.equal(text.indexOf('import.meta'), -1);
    assert.equal(text.indexOf('import_meta'), -1);
  }
  assert.ok(script.indexOf('currentScript') >= 0);

  // The bridge page: the four files as base64, the same bytes; the licence and where the font is from.
  const page = readFileSync(join(out, 'bymoffice.webview.html'), 'utf8');
  const inPage = blocks(page);
  assert.deepEqual([...inPage.keys()].sort(), files.slice().sort());
  for (const name of files) assert.equal(sha(inPage.get(name) as Buffer), sha(readFileSync(join(SOURCE, name))), name);
  assert.ok(page.indexOf('TeX Gyre Bonum') >= 0);
  assert.ok(page.indexOf('GUST Font License') >= 0);
  assert.ok(page.indexOf('https://www.gust.org.pl/projects/e-foundry/tex-gyre') >= 0);
  assert.ok(page.indexOf(readFileSync(join(SOURCE, LICENSE), 'utf8').trim().split('\n')[9]) >= 0, 'the text of the licence is in the page');
  assert.equal(page.indexOf('__BYM_FONTS__'), -1);
  assert.ok(policyHolds(page));

  // Without the fonts: a page that is smaller by what the fonts took, and still whole.
  const bare = build('without', { BYM_NO_FONTS: '1' });
  const barePage = readFileSync(join(bare, 'bymoffice.webview.html'), 'utf8');
  assert.equal(blocks(barePage).size, 0);
  assert.equal(barePage.indexOf('data-bym-font="'), -1);
  assert.equal(barePage.indexOf('__BYM_FONTS__'), -1);
  assert.ok(policyHolds(barePage));
  const total = files.reduce((sum, name) => sum + statSync(join(SOURCE, name)).size, 0);
  const grown = page.length - barePage.length;
  assert.ok(grown > (total * 4) / 3 && grown < (total * 4) / 3 + 8192, `the fonts add ${grown} characters to the page`);
  // The package itself keeps the files either way.
  assert.ok(existsSync(join(bare, 'fonts', 'tex-gyre-bonum', files[0])));
  // The same engine in both.
  assert.equal(readFileSync(join(bare, 'bymoffice.min.js'), 'utf8'), script);
  rmSync(OUT, { recursive: true, force: true });
});
