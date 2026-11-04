// The Node file server of examples/server/node.mjs, which the server guide (docs/server.md)
// shows: what `saveTo` sends is what it answers, and a file opened from it and saved back is
// the file that was written.

import assert from 'node:assert/strict';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import type { AddressInfo } from 'node:net';
import { join } from 'node:path';
import { test } from 'node:test';
// @ts-expect-error: a plain JavaScript example without types.
import { createFileServer } from '../examples/server/node.mjs';
import { createBlank, open } from '../src/index';

async function serving<T>(maxBytes: number | undefined, body: (base: string, dir: string) => Promise<T>): Promise<T> {
  const dir = mkdtempSync(join(process.cwd(), 'test', '.node-server-'));
  const server = createFileServer(dir, maxBytes === undefined ? undefined : { maxBytes });
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done));
  try {
    return await body(`http://127.0.0.1:${(server.address() as AddressInfo).port}`, dir);
  } finally {
    await new Promise((done) => server.close(done));
    rmSync(dir, { recursive: true, force: true });
  }
}

const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

test('example server: GET gives the file and a strong ETag; PUT with that ETag saves and gives the next one', async () => {
  await serving(undefined, async (base, dir) => {
    const blank = await createBlank('xlsx');
    writeFileSync(join(dir, 'report.xlsx'), blank);

    const got = await fetch(`${base}/files/report.xlsx`);
    assert.equal(got.status, 200);
    assert.equal(got.headers.get('content-type'), XLSX);
    const etag = got.headers.get('etag') ?? '';
    assert.match(etag, /^"[0-9a-f]{40}"$/);
    const bytes = new Uint8Array(await got.arrayBuffer());
    assert.deepEqual(bytes, new Uint8Array(blank));

    // What saveTo sends: the file as the body, its type, and If-Match.
    const file = await open(bytes);
    const saved = await file.save();
    const put = await fetch(`${base}/files/report.xlsx`, {
      method: 'PUT',
      headers: { 'Content-Type': XLSX, 'X-Bym-Format': 'xlsx', 'If-Match': etag },
      body: saved,
    });
    assert.equal(put.status, 200);
    const next = put.headers.get('etag');
    assert.equal(next, etag, 'an untouched file is saved byte for byte, so its ETag stays');
    assert.deepEqual(new Uint8Array(readFileSync(join(dir, 'report.xlsx'))), saved);

    // Someone else writes it: the old ETag no longer matches.
    const other = await createBlank('xlsx', { sheetName: 'Other' });
    assert.equal((await fetch(`${base}/files/report.xlsx`, { method: 'PUT', body: other })).status, 200);
    const stale = await fetch(`${base}/files/report.xlsx`, { method: 'PUT', headers: { 'If-Match': etag }, body: saved });
    assert.equal(stale.status, 412);
    assert.deepEqual(new Uint8Array(readFileSync(join(dir, 'report.xlsx'))), new Uint8Array(other));
  });
});

test('example server: a body over the limit is refused with 413 and the file stays', async () => {
  await serving(1024, async (base, dir) => {
    const blank = await createBlank('docx');
    writeFileSync(join(dir, 'a.docx'), blank);
    const big = await fetch(`${base}/files/a.docx`, { method: 'PUT', body: new Uint8Array(4096) });
    assert.equal(big.status, 413);
    assert.deepEqual(new Uint8Array(readFileSync(join(dir, 'a.docx'))), new Uint8Array(blank));
  });
});

test('example server: names that are not plain docx/xlsx names in the folder are not served', async () => {
  await serving(undefined, async (base) => {
    for (const path of ['/files/%E0%A4%A.xlsx', '/files/..%2Fpackage.json', '/files/notes.txt', '/files/.secret.xlsx', '/files/a%2Fb.xlsx', '/files/missing.xlsx', '/other']) {
      assert.equal((await fetch(`${base}${path}`)).status, 404, path);
    }
    assert.equal((await fetch(`${base}/files/x.xlsx`, { method: 'DELETE' })).status, 405);
    const page = await fetch(`${base}/`);
    assert.equal(page.status, 200);
    assert.match(await page.text(), /saveTo: true/);
  });
});

test('example server: names in any language and an upper-case extension are files like any other', async () => {
  await serving(undefined, async (base, dir) => {
    const blank = await createBlank('xlsx');
    for (const name of ['Bütçe-2026 ğüşıöç.xlsx', 'RAPOR.XLSX']) {
      writeFileSync(join(dir, name), blank);
      const got = await fetch(`${base}/files/${encodeURIComponent(name)}`);
      assert.equal(got.status, 200, name);
      assert.equal(got.headers.get('content-type'), XLSX, name);
    }
  });
});

test('example server: saves of one file at the same instant follow each other: one whole body is stored, its ETag is the answer, the other gets 412', async () => {
  await serving(undefined, async (base, dir) => {
    const path = join(dir, 'race.docx');
    for (let round = 0; round < 10; round++) {
      writeFileSync(path, await createBlank('docx'));
      const etag = (await fetch(`${base}/files/race.docx`)).headers.get('etag') ?? '';
      const bodies = [new Uint8Array(4 * 1024 * 1024).fill(0x41), new Uint8Array(4 * 1024 * 1024).fill(0x42)];
      const answers = await Promise.all(
        bodies.map((body) => fetch(`${base}/files/race.docx`, { method: 'PUT', headers: { 'If-Match': etag }, body })),
      );
      const statuses = answers.map((answer) => answer.status).sort();
      assert.deepEqual(statuses, [200, 412], `round ${round}`);
      const winner = answers.find((answer) => answer.status === 200)!;
      const stored = readFileSync(path);
      assert.ok(bodies.some((body) => Buffer.compare(Buffer.from(body), stored) === 0), 'the stored file is one of the bodies, whole');
      assert.equal(winner.headers.get('etag'), (await fetch(`${base}/files/race.docx`)).headers.get('etag'));
    }
    // Without If-Match both are saved, one after the other; the stored file is still one of them.
    const plain = await Promise.all([0x43, 0x44].map((byte) => fetch(`${base}/files/race.docx`, { method: 'PUT', body: new Uint8Array(1 << 20).fill(byte) })));
    assert.deepEqual(plain.map((answer) => answer.status), [200, 200]);
    const stored = readFileSync(path);
    assert.ok(stored.every((byte) => byte === stored[0]) && (stored[0] === 0x43 || stored[0] === 0x44));
  });
});
