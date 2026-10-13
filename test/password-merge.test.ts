// Password-protected files where they meet what was built next to them: editing together (the
// versions a file is known by, the file that is loaded again), saving to an address and by
// itself, and the commands of the bridge that arrive while a password is being asked for.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { toHex } from '../src/crypto/bytes';
import { contentVersion } from '../src/collab/index';
import { base64Decode, base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import { PasswordError, encryptionOf, getCell, open, setCellInput, applyOp, type OfficeFile } from '../src/index';
import { minimalXlsx } from './fixtures';

const PASSWORD = 'Gizli-Parola-42';
const isCfb = (bytes: Uint8Array): boolean => toHex(bytes.subarray(0, 8)) === 'd0cf11e0a1b11ae1';

let cached: Uint8Array | undefined;
const encrypted = async (): Promise<Uint8Array> => (cached ??= await (await open(await minimalXlsx())).save({ password: PASSWORD }));

// ------------------------------------------------------------------ versions and reloading

test('editing together: a version stands for the bytes at the address, which every participant reads alike; two saves of one content are two versions', async () => {
  const stored = await encrypted();
  // Two participants fetch the same file: the same version, whoever decrypts it.
  assert.equal(contentVersion(stored), contentVersion(stored.slice()));
  const first = await open(stored, { password: PASSWORD });
  const second = await open(stored.slice(), { password: PASSWORD });
  // Each of them saving the same workbook writes other bytes (new salts and keys): the version of a save is the
  // version of the bytes that were stored, never one that a participant could work out from the content.
  const a = await first.save();
  const b = await second.save();
  assert.notEqual(contentVersion(a), contentVersion(b));
  // Whoever loads what was stored again finds the version the saver announced.
  assert.equal(contentVersion(a), contentVersion(a.slice()));
  assert.equal((await open(a, { passwordOf: second })).format, 'xlsx');
});

test('editing together: the next version of a protected file is opened with the password the open one is kept with', async () => {
  const file = await open(await encrypted(), { password: PASSWORD });
  const workbook = await file.workbook();
  for (const op of setCellInput(workbook, 0, 0, 0, '77')) applyOp(workbook, op);
  const next = await file.save();
  // What mount() does when the room is at a newer version: the application is not asked for the password again.
  const reloaded = await open(next, { passwordOf: file });
  assert.deepEqual([reloaded.format, reloaded.passwordProtected], ['xlsx', true]);
  const cell = getCell((await reloaded.workbook()).sheets[0], 0, 0);
  assert.equal(cell && cell.value, 77);
  // It has a password of its own: closing the first file does not take it away.
  file.close();
  assert.ok(isCfb(await reloaded.save()), 'the reloaded file is saved encrypted');
  // A file that was closed lends nothing; a file with another password is wrong; a plain file needs none.
  await assert.rejects(open(next, { passwordOf: file }), (error: unknown) => error instanceof PasswordError && error.code === 'required');
  const other = await open(await (await open(await minimalXlsx())).save({ password: 'başka' }), { password: 'başka' });
  await assert.rejects(open(next, { passwordOf: other }), (error: unknown) => error instanceof PasswordError && error.code === 'wrong');
  assert.equal((await open(await minimalXlsx(), { passwordOf: reloaded })).passwordProtected, false);
  // A password that is given counts before one that is lent.
  assert.equal((await open(next, { password: PASSWORD, passwordOf: other })).format, 'xlsx');
});

// ------------------------------------------------------------------ the bridge

interface Fake {
  host: BridgeHost;
  sent: any[];
  joins: any[];
  views: number;
  pdfs: number;
  asked: number;
}

function fake(): Fake {
  const state: Fake = { host: undefined as unknown as BridgeHost, sent: [], joins: [], views: 0, pdfs: 0, asked: 0 };
  state.host = {
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: (bytes, options) => open(bytes, options),
    createView: () => {
      state.views++;
      return {
        history: { canUndo: false, canRedo: false },
        activeSheet: 0,
        isEditing: false,
        readOnly: false,
        commitEditing: () => undefined,
        commitPending: () => true,
        hasPendingInput: () => false,
        localizeError: (error: Error) => error.message,
        showNotice: () => undefined,
        setReadOnly: () => undefined,
        setLocale: () => undefined,
        undo: () => undefined,
        redo: () => undefined,
        attachCollab: () => undefined,
        collabChanged: () => undefined,
        collabNotice: () => undefined,
        collabPeersChanged: () => undefined,
        exportPdf: async () => {
          state.pdfs++;
          return new Uint8Array([0x25, 0x50, 0x44, 0x46]);
        },
        destroy: () => undefined,
      } as any;
    },
    createTransport: () => ({
      send: (text: string) => void state.joins.push(JSON.parse(text)),
      onMessage: () => undefined,
      close: () => undefined,
    }),
    passwordPrompt: () => ({
      ask: () => {
        state.asked++;
        return new Promise<string | null>(() => undefined);
      },
      close: () => undefined,
    }),
    timer: () => undefined,
    frame: () => undefined,
  };
  return state;
}

async function until(f: Fake, type: string, count = 1): Promise<any> {
  for (let i = 0; i < 3000; i++) {
    const found = f.sent.filter((message) => message.type === type);
    if (found.length >= count) return found[count - 1];
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error(`no "${type}" message (${count}) among: ${f.sent.map((message) => message.type).join(', ')}`);
}

const leaks = (f: Fake): boolean => JSON.stringify(f.sent).includes(PASSWORD) || JSON.stringify(f.joins).includes(PASSWORD);

test('bridge: a protected workbook that is edited together joins its room as the bytes that were loaded, and says it is encrypted', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const bytes = await encrypted();
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(bytes), password: PASSWORD, collab: { url: 'wss://relay.example/room', room: 'oda' } });
  const loaded = f.sent.find((message) => message.type === 'loaded');
  assert.deepEqual([loaded.format, loaded.encrypted], ['xlsx', true]);
  assert.equal(f.sent.some((message) => message.type === 'collabState'), true, 'the session was started');
  const join = f.joins.find((message) => message.t === 'join');
  assert.ok(join, `a join was sent: ${JSON.stringify(f.joins).slice(0, 200)}`);
  // The version of the encrypted file as it was handed over: what every other participant computes from the same file.
  assert.deepEqual(join.v, [contentVersion(bytes)]);
  // What `save` hands back is still encrypted.
  await bridge.receive({ type: 'save', id: 's' });
  const saved = f.sent.find((message) => message.type === 'saved');
  assert.equal(saved.encrypted, true);
  assert.ok(isCfb(base64Decode(saved.data)), 'the saved bytes are a compound file');
  assert.equal(leaks(f), false, 'the password is in no message, to the application or to the room');
  bridge.destroy();
});

test('bridge: commands that arrive while a password is asked for wait for the file, and are answered once it is open', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), collab: 'wss://relay.example/room' });
  await until(f, 'passwordRequired');
  assert.equal(f.asked, 1, 'the page asks as well');
  // The application goes on as if the file were open: a PDF, an undo, a save.
  const pdf = bridge.receive({ type: 'exportPdf', id: 'pdf' });
  const undo = bridge.receive({ type: 'undo', id: 'undo' });
  await new Promise((resolve) => setTimeout(resolve, 30));
  assert.deepEqual(
    f.sent.filter((message) => message.id === 'pdf' || message.id === 'undo'),
    [],
    'nothing is answered, and nothing is refused, before the file is there',
  );
  assert.equal(f.sent.some((message) => message.type === 'collabState'), false, 'no room is joined for a file that is not open');
  assert.equal(f.joins.length, 0);
  await bridge.receive({ type: 'password', value: PASSWORD });
  await Promise.all([loading, pdf, undo]);
  assert.equal(f.sent.some((message) => message.type === 'loaded'), true);
  assert.equal(f.pdfs, 1, 'the PDF was made of the file that was opened');
  assert.equal(f.sent.filter((message) => message.id === 'pdf' && message.type !== 'error').length >= 1, true, 'and answered');
  assert.deepEqual(
    f.sent.filter((message) => message.id === 'undo').map((message) => message.type),
    ['done'],
  );
  assert.equal(f.joins.some((message) => message.t === 'join'), true, 'the room is joined once the file is open');
  assert.equal(leaks(f), false);
  bridge.destroy();
});

test('bridge: when the password is given up, what waited behind the load is told that no file is open', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const loading = bridge.receive({ type: 'load', id: 'a', data: base64Encode(await encrypted()), collab: 'wss://relay.example/room' });
  await until(f, 'passwordRequired');
  const pdf = bridge.receive({ type: 'exportPdf', id: 'pdf' });
  await bridge.receive({ type: 'password', value: null });
  await Promise.all([loading, pdf]);
  assert.equal(f.sent.find((message) => message.id === 'a' && message.type === 'error').code, 'password-cancelled');
  assert.equal(f.sent.find((message) => message.id === 'pdf').code, 'not-loaded');
  assert.equal(f.pdfs, 0);
  assert.equal(f.joins.length, 0, 'no room was joined');
  assert.equal(f.sent.some((message) => message.type === 'collabState'), false);
  bridge.destroy();
});

test('saving a protected file that is edited: what is written is encrypted, whatever else the save goes through', async () => {
  // file.save() is what every way of saving ends in (saveTo, auto-save, the bridge, the bytes a room is told about).
  const file: OfficeFile = await open(await encrypted(), { password: PASSWORD });
  const saved = await file.save();
  assert.deepEqual(encryptionOf(saved) && encryptionOf(saved)!.encryption!.scheme, 'agile');
  assert.match(contentVersion(saved), /^h:\d+:[0-9a-f]{16}$/);
});
