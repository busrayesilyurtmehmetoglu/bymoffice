// The declarations tsc writes name their neighbours without an extension
// ("./zip"), which a bundler resolves and Node's own resolution ("node16",
// "nodenext") does not. After tsc, every relative specifier in dist/types
// gets the extension of the file it means, so that the types of the package
// resolve whichever way the project that uses it is set up.

import { existsSync } from 'node:fs';
import { readdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';

const ROOT = 'dist/types';

async function files(dir) {
  const out = [];
  for (const item of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, item.name);
    if (item.isDirectory()) out.push(...(await files(path)));
    else if (item.name.endsWith('.d.ts')) out.push(path);
  }
  return out;
}

let changed = 0;
for (const file of await files(ROOT)) {
  const text = await readFile(file, 'utf8');
  const fixed = text.replace(/(\bfrom\s+|\bimport\s*\(\s*|\bimport\s+)(['"])(\.{1,2}\/[^'"]*)\2/g, (whole, lead, quote, specifier) => {
    if (/\.(js|mjs|cjs|json)$/.test(specifier)) return whole;
    const target = resolve(dirname(file), specifier);
    if (existsSync(`${target}.d.ts`)) return `${lead}${quote}${specifier}.js${quote}`;
    if (existsSync(join(target, 'index.d.ts'))) return `${lead}${quote}${specifier.replace(/\/$/, '')}/index.js${quote}`;
    throw new Error(`${file}: "${specifier}" names no declaration file`);
  });
  if (fixed !== text) {
    await writeFile(file, fixed);
    changed++;
  }
}
console.log(`dist/types: relative imports given their extensions in ${changed} files`);
