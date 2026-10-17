// Dependency-free static server for the Playwright tests. Serves the repository
// root so that /dist, /demo and /e2e/pages are reachable exactly as files on disk.
//
// One place takes what a page sends: /e2e/saved/<name>, for the tests of saving a
// file to where it came from and for the demo page. What is sent there is kept in
// memory only (nothing is written to disk), under a plain name, up to a size and
// a number of files; it is gone when the server stops.
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { stat } from 'node:fs/promises';
import { dirname, extname, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createCollabServer } from '../server/collab-server.mjs';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const PORT = Number(process.env.E2E_PORT ?? 4317);

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.map': 'application/json; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.woff2': 'font/woff2',
  '.otf': 'font/otf',
  '.txt': 'text/plain; charset=utf-8',
};

const send = (res, status, text) => {
  res.writeHead(status, { 'content-type': 'text/plain; charset=utf-8' });
  res.end(text);
};

/** Where the demo site is served, as on GitHub Pages. */
const SITE = '/bymoffice/';

/** Where saved files are sent to and fetched from again. */
const SAVED = '/e2e/saved/';
/** A name is letters, digits, dots, dashes and underscores, and does not start with a dot: it is never a path. */
const SAVED_NAME = /^[A-Za-z0-9_-][A-Za-z0-9._-]{0,79}$/;
const MAX_SAVED_BYTES = 8 * 1024 * 1024;
const MAX_SAVED_FILES = 64;
const MAX_SAVED_TOTAL = 64 * 1024 * 1024;
/** name -> { bytes, etag, type, format, convertedFrom }, oldest first. */
const saved = new Map();
let savedTotal = 0;
let savedSerial = 0;
/** Uploads that are being received: each holds up to MAX_SAVED_BYTES in memory until it ends. */
let receiving = 0;
const MAX_RECEIVING = 8;

/**
 * Whether a request that writes comes from a page of this server. A page of another site can make
 * a browser send a "simple" POST here; the browser says where it comes from, and that is refused.
 * (A client that is no browser sends neither header: the tests seed files that way.)
 */
const fromThisSite = (req) => {
  const site = req.headers['sec-fetch-site'];
  if (typeof site === 'string' && site !== 'same-origin' && site !== 'none') return false;
  const origin = req.headers.origin;
  return typeof origin !== 'string' || origin === `http://${req.headers.host}`;
};

const forget = (name) => {
  const entry = saved.get(name);
  if (!entry) return;
  savedTotal -= entry.bytes.length;
  saved.delete(name);
};

/** GET and HEAD give a saved file back with its ETag; PUT and POST take one, and honour If-Match as a file server would. */
const savedFile = (req, res, name) => {
  if (!SAVED_NAME.test(name)) return send(res, 403, 'Forbidden');
  const entry = saved.get(name);
  if (req.method === 'GET' || req.method === 'HEAD') {
    if (!entry) return send(res, 404, 'Not found');
    const headers = { 'content-type': 'application/octet-stream', 'content-length': entry.bytes.length, 'cache-control': 'no-store', etag: entry.etag };
    // What the request said the file is, for a test to read back.
    if (entry.type) headers['x-saved-content-type'] = entry.type;
    if (entry.format) headers['x-saved-format'] = entry.format;
    if (entry.convertedFrom) headers['x-saved-converted-from'] = entry.convertedFrom;
    res.writeHead(200, headers);
    return req.method === 'HEAD' ? res.end() : res.end(entry.bytes);
  }
  if (req.method !== 'PUT' && req.method !== 'POST') return send(res, 405, 'Method not allowed');
  if (!fromThisSite(req)) return send(res, 403, 'Forbidden');
  // "?status=503" answers with that status and keeps nothing: how a page shows a save that failed.
  // "?status=302" (any 3xx) redirects instead, as a server whose session ran out does: to "followed-<name>"
  // here, so that a client that follows the redirect leaves a file under that name for a test to find.
  const wanted = Number(new URL(req.url ?? '/', 'http://localhost').searchParams.get('status'));
  const refuse = Number.isInteger(wanted) && wanted >= 300 && wanted <= 599 ? wanted : 0;
  if (Number(req.headers['content-length']) > MAX_SAVED_BYTES) {
    send(res, 413, 'Too large');
    return void req.destroy();
  }
  if (receiving >= MAX_RECEIVING) {
    send(res, 503, 'Too many uploads at once');
    return void req.destroy();
  }
  receiving++;
  let counted = true;
  const received = () => {
    if (counted) receiving--;
    counted = false;
  };
  req.on('close', received);
  const chunks = [];
  let size = 0;
  let over = false;
  req.on('data', (chunk) => {
    if (over) return;
    size += chunk.length;
    if (size > MAX_SAVED_BYTES) {
      over = true;
      send(res, 413, 'Too large');
      return void req.destroy();
    }
    chunks.push(chunk);
  });
  req.on('error', () => undefined);
  req.on('end', () => {
    received();
    if (over) return;
    if (refuse >= 300 && refuse < 400) {
      res.writeHead(refuse, { location: `${SAVED}followed-${name}`.slice(0, SAVED.length + 80), 'content-type': 'text/plain; charset=utf-8' });
      return res.end('Redirected as asked');
    }
    if (refuse) return send(res, refuse, 'Refused as asked');
    const now = saved.get(name);
    const match = req.headers['if-match'];
    // The file was changed by someone else since the page fetched it: it is not overwritten.
    // "*" asks for a file that exists, whatever its ETag.
    if (typeof match === 'string' && (match === '*' ? !now : !now || now.etag !== match)) {
      res.writeHead(412, { 'content-type': 'text/plain; charset=utf-8', ...(now ? { etag: now.etag } : {}) });
      return res.end('Precondition failed');
    }
    const bytes = Buffer.concat(chunks);
    forget(name);
    while (saved.size >= MAX_SAVED_FILES || (saved.size > 0 && savedTotal + bytes.length > MAX_SAVED_TOTAL)) forget(saved.keys().next().value);
    const header = (key) => (typeof req.headers[key] === 'string' ? req.headers[key].slice(0, 200) : '');
    const etag = `"${++savedSerial}-${createHash('sha1').update(bytes).digest('hex').slice(0, 16)}"`;
    saved.set(name, { bytes, etag, type: header('content-type'), format: header('x-bym-format'), convertedFrom: header('x-bym-converted-from') });
    savedTotal += bytes.length;
    res.writeHead(204, { etag });
    res.end();
  });
};

const server = createServer(async (req, res) => {
  let pathname;
  try {
    pathname = decodeURIComponent(new URL(req.url ?? '/', 'http://localhost').pathname);
  } catch {
    return send(res, 400, 'Bad request');
  }
  // The one place that takes what a page sends; its names are checked there, and never reach the disk.
  if (pathname.startsWith(SAVED)) return savedFile(req, res, pathname.slice(SAVED.length));
  if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'Method not allowed');
  // The demo site (site-dist/, made by scripts/build-site.mjs) under the path it is published at,
  // with a folder's index.html as GitHub Pages gives it. The checks below apply to it as to any path.
  if (pathname.startsWith(SITE)) pathname = `/site-dist/${pathname.slice(SITE.length)}${pathname.endsWith('/') ? 'index.html' : ''}`;
  // Reject NUL bytes, backslashes and any dot-segment or hidden file (.git, .env, ..).
  if (pathname.includes('\0') || pathname.includes('\\') || pathname.split('/').some((part) => part.startsWith('.'))) {
    return send(res, 403, 'Forbidden');
  }
  const file = resolve(ROOT, `.${pathname}`);
  if (file !== ROOT && !file.startsWith(ROOT + sep)) return send(res, 403, 'Forbidden');

  let info;
  try {
    info = await stat(file);
  } catch {
    return send(res, 404, 'Not found');
  }
  if (!info.isFile()) return send(res, 404, 'Not found');

  res.writeHead(200, {
    'content-type': TYPES[extname(file).toLowerCase()] ?? 'application/octet-stream',
    'content-length': info.size,
    'cache-control': 'no-store',
  });
  if (req.method === 'HEAD') return res.end();
  createReadStream(file).on('error', () => res.destroy()).pipe(res);
});

// The relay for the tests of editing together and for the demo page: rooms under /e2e/collab/, in
// memory, for pages of this server only.
createCollabServer({ path: '/e2e/collab/', origins: [`http://127.0.0.1:${PORT}`, `http://localhost:${PORT}`] }).attach(server);

server.listen(PORT, '127.0.0.1', () => {
  console.log(`e2e static server: http://127.0.0.1:${PORT}/ (root ${ROOT})`);
});
