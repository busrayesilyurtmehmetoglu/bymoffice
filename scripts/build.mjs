import { createHash } from 'node:crypto';
import { readFile, writeFile } from 'node:fs/promises';
import { Script } from 'node:vm';
import { build } from 'esbuild';

const shared = {
  entryPoints: ['src/index.ts'],
  bundle: true,
  // Old WebViews (iOS before 16.4, Android 8-era) must parse the bundle; newer syntax is lowered.
  target: 'es2019',
  sourcemap: true,
};

await Promise.all([
  // <script src=".../bymoffice.min.js"> exposes a global `BymOffice`.
  build({ ...shared, format: 'iife', globalName: 'BymOffice', minify: true, outfile: 'dist/bymoffice.min.js' }),
  build({ ...shared, format: 'esm', outfile: 'dist/bymoffice.esm.js' }),
]);

// dist/bymoffice.webview.html: the bridge page for WebViews, with the engine
// inside it, so that an application ships one file and needs no network.
{
  const template = await readFile('src/embed/host.html', 'utf8');
  const { version } = JSON.parse(await readFile('package.json', 'utf8'));
  const engine = (await readFile('dist/bymoffice.min.js', 'utf8'))
    // The source map is not shipped with the page.
    .replace(/\n\/\/# sourceMappingURL=.*\s*$/, '\n')
    // Inside an inline script these would end the element or open an HTML comment.
    .replace(/<\/script/gi, '<\\/script')
    .replace(/<!--/g, '<\\!--');
  // The escapes are harmless inside strings; anywhere else they would break the code, and this notices.
  new Script(engine, { filename: 'bymoffice.webview.html (engine)' });

  // Replacer functions: the engine's text contains "$&" and the like.
  const fill = (text, marker, value) => {
    if (text.split(marker).length !== 2) throw new Error(`src/embed/host.html must contain ${marker} exactly once`);
    return text.replace(marker, () => value);
  };
  let page = fill(fill(template, '__BYM_ENGINE__', engine), '__BYM_VERSION__', version);

  // BYM_PARENT_ORIGINS="https://app.example.com https://admin.example.com": the pages that may
  // embed this one in an iframe. Not set: none, which is right for a page shipped inside an application.
  const origins = (process.env.BYM_PARENT_ORIGINS ?? '').split(/\s+/).filter(Boolean).map((text) => {
    let url;
    try {
      url = new URL(text);
    } catch {
      throw new Error(`BYM_PARENT_ORIGINS: "${text}" is not an origin`);
    }
    if ((url.protocol !== 'http:' && url.protocol !== 'https:') || url.origin === 'null') throw new Error(`BYM_PARENT_ORIGINS: "${text}" is not an http(s) origin`);
    return url.origin;
  });
  page = fill(page, '__BYM_PARENT_ORIGINS__', origins.join(' '));

  // Only the two scripts of the page itself may run: each is allowed by the hash of its text.
  const hashes = [];
  for (const match of page.matchAll(/<script>([\s\S]*?)<\/script>/g)) {
    hashes.push(`'sha256-${createHash('sha256').update(match[1], 'utf8').digest('base64')}'`);
  }
  if (hashes.length !== 2) throw new Error(`Expected 2 inline scripts in the bridge page, found ${hashes.length}`);
  const policy = [
    "default-src 'none'",
    `script-src ${hashes.join(' ')}`,
    // The view adds its own style element, whose text is not known here.
    "style-src 'unsafe-inline'",
    'img-src data:',
    // No remote hosts: a `load` with a URL reaches the page's own origin, blob: and data: only.
    "connect-src 'self' blob: data:",
    "base-uri 'none'",
    "form-action 'none'",
  ].join('; ');
  page = fill(page, '__BYM_CSP__', policy);
  await writeFile('dist/bymoffice.webview.html', page);
}
