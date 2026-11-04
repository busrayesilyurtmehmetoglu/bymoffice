// A file server for `saveTo`, with Node's own `http` module and no dependencies.
//
//   GET /files/<name>   the file, with a strong ETag
//   PUT /files/<name>   the new content as the raw body; with If-Match, 412 when the file changed
//   GET /               a page with the editor (the script from the CDN)
//
// Run: node examples/server/node.mjs [folder]   (default folder: ./files, port 8080 or $PORT)
// This is an example: put your own sign-in and permission checks where `allowed()` is.

import { createHash, randomUUID } from 'node:crypto';
import { readFile, rename, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const TYPES = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};

const etagOf = (bytes) => `"${createHash('sha1').update(bytes).digest('hex')}"`;

/** Only plain names of docx and xlsx files in the folder (letters of any language): no paths, no dot files. */
function allowed(name) {
  return /^[\p{L}\p{N}_-][\p{L}\p{N}_ .-]*\.(docx|xlsx)$/iu.test(name) && !name.includes('..');
}

/** The name in the address, or '' when it is not one (a broken %-escape too). */
function nameOf(pathname) {
  const match = /^\/files\/([^/]+)$/.exec(pathname);
  if (!match) return '';
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return '';
  }
}

async function current(path) {
  try {
    return await readFile(path);
  } catch (error) {
    if (error.code === 'ENOENT') return undefined;
    throw error;
  }
}

/** The body, or undefined when it is larger than `limit` bytes. */
function bodyOf(req, limit) {
  return new Promise((done, fail) => {
    const chunks = [];
    let size = 0;
    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        req.removeAllListeners('data');
        req.resume();
        done(undefined);
      } else chunks.push(chunk);
    });
    req.on('end', () => size <= limit && done(Buffer.concat(chunks)));
    req.on('error', fail);
  });
}

const PAGE = `<!doctype html>
<html><head><meta charset="utf-8"><title>Editor</title>
<style>html, body, #editor { height: 100%; margin: 0 }</style></head>
<body><div id="editor"></div>
<script src="https://cdn.jsdelivr.net/npm/bymoffice/dist/bymoffice.min.js"></script>
<script>
  const file = new URLSearchParams(location.search).get('file') || '/files/report.xlsx';
  BymOffice.mount(document.getElementById('editor'), { file, saveTo: true });
</script></body></html>`;

/**
 * A server for the files of `folder`. `maxBytes` is the largest body a save may send
 * (default 50 MB); a larger one is answered with 413.
 */
export function createFileServer(folder, { maxBytes = 50 * 1024 * 1024 } = {}) {
  const root = resolve(folder);
  // Saves of one file run one after another: the If-Match check and the write are then one step,
  // and two saves at the same instant can neither mix their bytes nor both pass the check.
  const queues = new Map();
  const inTurn = (path, work) => {
    const run = (queues.get(path) ?? Promise.resolve()).then(work, work);
    const tail = run.catch(() => undefined);
    queues.set(path, tail);
    tail.then(() => queues.get(path) === tail && queues.delete(path));
    return run;
  };
  return createServer(async (req, res) => {
    try {
      const url = new URL(req.url, 'http://localhost');
      if (url.pathname === '/' && req.method === 'GET') {
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        return res.end(PAGE);
      }
      const name = nameOf(url.pathname);
      if (!allowed(name)) {
        res.writeHead(404);
        return res.end();
      }
      const path = join(root, name);

      if (req.method === 'GET') {
        const bytes = await current(path);
        if (!bytes) {
          res.writeHead(404);
          return res.end();
        }
        res.writeHead(200, {
          'Content-Type': TYPES[name.split('.').pop().toLowerCase()],
          'ETag': etagOf(bytes),
          'Cache-Control': 'no-store',
        });
        return res.end(bytes);
      }

      if (req.method === 'PUT') {
        const body = await bodyOf(req, maxBytes);
        if (!body) {
          res.writeHead(413);
          return res.end();
        }
        const want = req.headers['if-match'];
        const saved = await inTurn(path, async () => {
          const before = await current(path);
          if (want && before && want !== etagOf(before)) return false;
          // Written beside it under a name of its own and renamed: a reader never gets half a file.
          const part = `${path}.${randomUUID()}.part`;
          await writeFile(part, body);
          await rename(part, path);
          return true;
        });
        res.writeHead(saved ? 200 : 412, saved ? { 'ETag': etagOf(body) } : {});
        return res.end();
      }

      res.writeHead(405, { 'Allow': 'GET, PUT' });
      res.end();
    } catch {
      if (!res.headersSent) res.writeHead(500);
      res.end();
    }
  });
}

if (import.meta.url === pathToFileURL(process.argv[1] ?? '').href) {
  const port = Number(process.env.PORT) || 8080;
  createFileServer(process.argv[2] ?? 'files').listen(port, '127.0.0.1', () => {
    console.log(`http://127.0.0.1:${port}/?file=/files/report.xlsx`);
  });
}
