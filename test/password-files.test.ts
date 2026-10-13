// Password-protected files through open(): docx and xlsx (opened, and saved encrypted again),
// xls and doc (opened read-only), and what each of them does when it is saved. The encrypted
// files come from test/tools/offcrypto_ref.py, and what the engine encrypts goes back to it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';
import { utf16le } from '../src/crypto/bytes';
import { PasswordSession } from '../src/crypto/session';
import { decryptDocStreams, xorArray, xorKey, xorPasswordBytes, xorVerifier } from '../src/crypto/legacy';
import { toHex } from '../src/crypto/bytes';
import { convertForEditing, openConverted } from '../src/formats/convert';
import { DEFAULT_PASSWORD, PasswordError, encryptionOf, open, type OfficeFile } from '../src/index';
import type { Paragraph } from '../src/model/doc/types';
import { getCell, type Sheet } from '../src/model/sheet/types';
import { setCellInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';
import { paragraphText } from '../src/model/doc/types';
import { DocBuilder, simpleDoc } from './doc-binary-fixtures';
import { buildDocx as docxFixture } from './docx-fixtures';
import { minimalXlsx } from './fixtures';
import { REFTOOLS, passwordArg, ref, refOk, refTest } from './reftools';
import { cat, number, workbookStream } from './xls-fixtures';

const PASSWORD = 'Password1234_';

async function rejectsWith(work: Promise<unknown>, code: string, label: string): Promise<PasswordError> {
  try {
    await work;
  } catch (error) {
    assert.ok(error instanceof PasswordError, `${label}: a PasswordError is thrown, got ${String(error)}`);
    assert.equal(error.code, code, `${label}: ${error.message}`);
    return error;
  }
  throw new assert.AssertionError({ message: `${label}: nothing was thrown` });
}

const isCfb = (bytes: Uint8Array): boolean => toHex(bytes.subarray(0, 8)) === 'd0cf11e0a1b11ae1';
const isZip = (bytes: Uint8Array): boolean => bytes[0] === 0x50 && bytes[1] === 0x4b;

/** What the reference makes of an encrypted file: the payload, checked as a ZIP file. */
function refDecrypt(file: Uint8Array, password: string): Uint8Array {
  const done = ref('decrypt', { password: passwordArg(password), zip: 'check' }, file);
  assert.equal(done.status, 0, `the reference decrypts the file: ${done.err}`);
  assert.match(done.err, /kind=agile hmac=True/, 'agile encryption with an HMAC that holds');
  return done.out;
}

/**
 * The package inside an encrypted file. Where the reference tools are there, the reference decrypts it (and checks
 * the HMAC and the ZIP file); where they are not, the engine does: the rules of saving are tested either way.
 */
async function decrypted(file: Uint8Array, password: string): Promise<Uint8Array> {
  assert.ok(isCfb(file), 'an encrypted file is a compound file');
  if (REFTOOLS) return refDecrypt(file, password);
  return (await open(file, { password, requireIntegrity: true })).save({ password: null });
}

/** That a password does not open a file: for the engine, and for the reference where it is there. */
async function refuses(file: Uint8Array, password: string, label: string): Promise<void> {
  await rejectsWith(open(file, { password }), 'wrong', label);
  if (REFTOOLS) assert.equal(ref('decrypt', { password: passwordArg(password) }, file).status, 2, `${label}: the reference refuses it as well`);
}

let ownOnce: Uint8Array | undefined;
/** An xlsx the engine itself encrypted with PASSWORD: what the rules of saving are tested on, with or without the reference tools. */
const ownXlsx = async (): Promise<Uint8Array> => (ownOnce ??= await (await open(await minimalXlsx())).save({ password: PASSWORD }));

const encryptedXlsx = async (password = PASSWORD, options: Record<string, string | number> = {}): Promise<Uint8Array> =>
  refOk('agile-encrypt', { password: passwordArg(password), spin: 1000, ...options }, await minimalXlsx());

const valueAt = (sheet: Sheet, row: number, col: number): unknown => {
  const cell = getCell(sheet, row, col);
  return cell ? cell.value : undefined;
};

const firstSheetValue = async (file: OfficeFile): Promise<unknown> => valueAt((await file.workbook()).sheets[0], 0, 0);

// ------------------------------------------------------------------ xlsx and docx

refTest('open: an encrypted xlsx is opened with its password, and says how it was encrypted', async () => {
  const bytes = await encryptedXlsx();
  assert.deepEqual(encryptionOf(bytes), { format: 'ooxml', encryption: { scheme: 'agile', cipher: 'AES', keyBits: 256, hash: 'SHA512', spinCount: 1000, integrity: true } });
  const file = await open(bytes, { password: PASSWORD });
  assert.deepEqual([file.format, file.kind, file.container, file.passwordProtected], ['xlsx', 'spreadsheet', 'ooxml', true]);
  assert.deepEqual(file.encryption, { scheme: 'agile', cipher: 'AES', keyBits: 256, hash: 'SHA512', spinCount: 1000, integrity: true });
  assert.equal((await file.workbook()).sheets.length, 1);
});

refTest('open: an encrypted file without a password is "required", with a wrong one "wrong", and both say what kind of file it is', async () => {
  const bytes = await encryptedXlsx();
  const required = await rejectsWith(open(bytes), 'required', 'no password');
  assert.equal(required.format, 'ooxml');
  const wrong = await rejectsWith(open(bytes, { password: 'password1234_' }), 'wrong', 'a wrong password');
  assert.equal(wrong.format, 'ooxml');
  await rejectsWith(open(bytes, { password: '' }), 'wrong', 'the empty password');
  assert.ok(!wrong.message.includes('password1234_'), 'the message does not repeat the password');
});

refTest('open: an encrypted docx (standard encryption, as Office 2007 writes) is opened, and saved with agile encryption', async () => {
  const docx = await docxFixture({ body: '<w:p><w:r><w:t>Gizli belge</w:t></w:r></w:p>' });
  const bytes = refOk('standard-encrypt', { password: passwordArg(PASSWORD), bits: 128 }, docx);
  assert.deepEqual(encryptionOf(bytes), { format: 'ooxml', encryption: { scheme: 'standard', cipher: 'AES', keyBits: 128, hash: 'SHA1', spinCount: 50000, integrity: false } });
  const file = await open(bytes, { password: PASSWORD });
  assert.deepEqual([file.format, file.passwordProtected, file.encryption && file.encryption.scheme], ['docx', true, 'standard']);
  const model = await file.document!();
  assert.equal(paragraphText(model.body.blocks[0] as Paragraph), 'Gizli belge');
  const saved = await file.save();
  assert.ok(isCfb(saved), 'what is saved is a compound file');
  const again = await open(refDecrypt(saved, PASSWORD));
  assert.equal(paragraphText((await again.document!()).body.blocks[0] as Paragraph), 'Gizli belge');
});

refTest('open: a file encrypted with the password Office uses when the user gave none opens without a password, and is saved with it again', async () => {
  const bytes = await encryptedXlsx(DEFAULT_PASSWORD);
  const file = await open(bytes);
  assert.deepEqual([file.format, file.passwordProtected], ['xlsx', true]);
  const saved = await file.save();
  assert.ok(isCfb(saved), 'saved encrypted');
  refDecrypt(saved, DEFAULT_PASSWORD);
  assert.equal(file.encryption && file.encryption.defaultPassword, true, 'the file says that anyone can open it');
  // A password that was given and is wrong is wrong: the one everyone knows is not tried behind it.
  await rejectsWith(open(bytes, { password: 'something else' }), 'wrong', 'a password that is not needed, and not right');
});

test('save: a file that was opened with a password is saved with the same password, changes included', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  const workbook = await file.workbook();
  for (const op of setCellInput(workbook, 0, 0, 0, '4711')) applyOp(workbook, op);
  const saved = await file.save();
  assert.ok(isCfb(saved) && !isZip(saved), 'what is saved is a compound file, not a ZIP file');
  assert.deepEqual(encryptionOf(saved), { format: 'ooxml', encryption: { scheme: 'agile', cipher: 'AES', keyBits: 256, hash: 'SHA512', spinCount: 100000, integrity: true } });
  // It opens with the password, and with no other (for the reference too, where it is there).
  const plain = await decrypted(saved, PASSWORD);
  assert.ok(isZip(plain), 'the package inside');
  assert.equal(await firstSheetValue(await open(plain)), 4711);
  await refuses(saved, 'Password1234', 'another password');
  // And so does the engine.
  assert.equal(await firstSheetValue(await open(saved, { password: PASSWORD })), 4711);
  await rejectsWith(open(saved), 'required', 'the saved file without a password');
  // The streams Office expects next to the two of the encryption.
  const cfb = CfbFile.read(saved);
  for (const path of ['EncryptionInfo', 'EncryptedPackage', '\u0006DataSpaces/Version', '\u0006DataSpaces/DataSpaceMap', '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary']) {
    assert.ok(cfb.has(path), `the saved file has ${JSON.stringify(path)}`);
  }
  assert.equal(toHex(cfb.stream('\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary').subarray(0, 12)), '58000000010000004c000000');
  assert.equal(cfb.stream('\u0006DataSpaces/Version').length, 76);
  assert.equal(cfb.stream('\u0006DataSpaces/DataSpaceMap').length, 112);
  assert.equal(cfb.stream('\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace').length, 64);
  assert.equal(cfb.stream('\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary').length, 200);
});

test('save: every save has salts and keys of its own', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  const a = CfbFile.read(await file.save());
  const b = CfbFile.read(await file.save());
  assert.notEqual(toHex(a.stream('EncryptedPackage').subarray(8, 72)), toHex(b.stream('EncryptedPackage').subarray(8, 72)), 'the same package is another cipher text');
  const info = (cfb: CfbFile): string => new TextDecoder().decode(cfb.stream('EncryptionInfo').subarray(8));
  const keyValue = (text: string): string => (/encryptedKeyValue="([^"]+)"/.exec(text) || [])[1];
  const dataSalt = (text: string): string => (/<keyData[^>]*saltValue="([^"]+)"/.exec(text) || [])[1];
  assert.notEqual(keyValue(info(a)), keyValue(info(b)), 'another key');
  assert.notEqual(dataSalt(info(a)), dataSalt(info(b)), 'another salt');
  // The salt of the password, and the rounds over it, are done once for an open file: every later save is quick.
  const passwordSalt = (text: string): string => (/<p:encryptedKey[^>]*saltValue="([^"]+)"/.exec(text) || [])[1];
  assert.equal(passwordSalt(info(a)), passwordSalt(info(b)), 'the salt of the password is kept while the file is open');
  const other = await open(await ownXlsx(), { password: PASSWORD });
  assert.notEqual(passwordSalt(info(CfbFile.read(await other.save()))), passwordSalt(info(a)), 'and is another one for another open file');
});

test('save: saves that run at the same time all give files that open; a file that is closed while it is saved is not saved', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  // Three at once on a file that has derived nothing yet, and two more behind them.
  const first = await Promise.all([file.save(), file.save(), file.save()]);
  const later = await Promise.all([file.save(), file.save()]);
  for (const saved of [...first, ...later]) {
    assert.ok(isZip(await decrypted(saved, PASSWORD)), 'the package inside');
    assert.equal((await open(saved, { password: PASSWORD })).format, 'xlsx');
  }
  const closing = file.save();
  file.close();
  await assert.rejects(closing, /password of this file was forgotten/);
});

test('session: the password is kept as bytes, zeroed when it is forgotten, and shows in no printout', () => {
  const session = PasswordSession.of('Sızmaz-1');
  const bytes = session.password;
  assert.equal(toHex(bytes), toHex(utf16le('Sızmaz-1')));
  const derived = { salt: new Uint8Array(16).fill(1), base: new Uint8Array(64).fill(7), rounds: 1 };
  session.agileDerived = derived;
  const copy = session.clone();
  for (const text of [JSON.stringify(session), JSON.stringify({ session }), String(session), `${session}`]) {
    assert.ok(!text.includes('Sızmaz') && !text.includes(toHex(bytes)), `nothing of the password in ${text}`);
  }
  session.forget();
  assert.equal(session.forgotten, true);
  assert.ok(bytes.every((byte) => byte === 0), 'the bytes of the password are zero');
  assert.ok(derived.base.every((byte) => byte === 0), 'and so is what was derived from it');
  assert.throws(() => session.password, /forgotten/);
  assert.equal(toHex(copy.password), toHex(utf16le('Sızmaz-1')), 'a copy has bytes of its own');
  session.forget();
});

test('save: the password is taken away or changed only when that is asked for', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  // For one save.
  const plain = await file.save({ password: null });
  assert.ok(isZip(plain), 'save({ password: null }) gives the package itself');
  assert.equal((await open(plain)).passwordProtected, false);
  const other = await file.save({ password: 'Yeni-Şifre-2' });
  assert.ok(isZip(await decrypted(other, 'Yeni-Şifre-2')), 'encrypted with the password of this one save');
  await refuses(other, PASSWORD, 'the old password');
  // The file itself is still kept with the password it was opened with.
  assert.ok(isZip(await decrypted(await file.save(), PASSWORD)), 'the next save has the old password again');
  // From now on.
  file.setPassword('Kalıcı');
  assert.equal(file.passwordProtected, true);
  const kept = await file.save();
  assert.ok(isZip(await decrypted(kept, 'Kalıcı')), 'encrypted with the new password');
  await refuses(kept, PASSWORD, 'the password the file was opened with');
  file.setPassword(null);
  assert.equal(file.passwordProtected, false);
  assert.ok(isZip(await file.save()), 'after setPassword(null) the file is saved without a password');
});

test('save: a file without a password is encrypted when a save asks for it, and only then', async () => {
  const file = await open(await minimalXlsx());
  assert.deepEqual([file.passwordProtected, file.encryption], [false, undefined]);
  assert.ok(isZip(await file.save()), 'saved as it was opened');
  assert.ok(isZip(await decrypted(await file.save({ password: 'Yeni' }), 'Yeni')), 'encrypted for the one save');
  assert.ok(isZip(await file.save()), 'the save after it is plain again');
});

test('save: an empty password, one that is too long and one that is no string are refused', async () => {
  const file = await open(await minimalXlsx());
  await assert.rejects(file.save({ password: '' }), /empty password/);
  await assert.rejects(file.save({ password: 'x'.repeat(256) }), /at most 255 characters/);
  await assert.rejects(file.save({ password: 5 as unknown as string }), /A password is a string/);
  assert.throws(() => file.setPassword(''), /empty password/);
  assert.throws(() => file.setPassword('x'.repeat(256)), /at most 255 characters/);
  assert.equal(file.passwordProtected, false, 'a refused password changes nothing');
});

test('save: where no secure random numbers are to be had, a password-protected file is not saved at all', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  const real = Object.getOwnPropertyDescriptor(globalThis, 'crypto');
  Object.defineProperty(globalThis, 'crypto', { value: undefined, configurable: true, writable: true });
  let saved: unknown;
  let failure: unknown;
  try {
    saved = await file.save();
  } catch (error) {
    failure = error;
  } finally {
    if (real) Object.defineProperty(globalThis, 'crypto', real);
  }
  assert.equal(saved, undefined, 'nothing was saved');
  assert.match(String(failure), /no secure source of random numbers/);
  // The plain file can still be asked for by name.
  assert.ok(isCfb(await file.save()), 'with random numbers back the file is saved encrypted');
});

test('close: a file that was closed forgets its password and is not saved without it', async () => {
  const file = await open(await ownXlsx(), { password: PASSWORD });
  file.close();
  assert.equal(file.passwordProtected, true, 'still a protected file');
  await assert.rejects(file.save(), /password of this file was forgotten/);
  assert.ok(isZip(await file.save({ password: null })), 'asked for by name, the plain file is given');
});

refTest('open: a damaged encrypted file is "damaged", an encryption that is not read "unsupported"', async () => {
  const bytes = await encryptedXlsx();
  const cfb = CfbFile.read(bytes);
  const pkg = cfb.stream('EncryptedPackage').slice();
  pkg[100] ^= 1;
  const tampered = writeCfb([
    { path: 'EncryptionInfo', data: cfb.stream('EncryptionInfo') },
    { path: 'EncryptedPackage', data: pkg },
  ]);
  const damaged = await rejectsWith(open(tampered, { password: PASSWORD }), 'damaged', 'a changed package');
  assert.equal(damaged.format, 'ooxml');
  const cut = writeCfb([
    { path: 'EncryptionInfo', data: cfb.stream('EncryptionInfo') },
    { path: 'EncryptedPackage', data: pkg.subarray(0, 40) },
  ]);
  await rejectsWith(open(cut, { password: PASSWORD }), 'damaged', 'a package cut short');
  const extensible = writeCfb([
    { path: 'EncryptionInfo', data: new Uint8Array([4, 0, 3, 0, 0x10, 0, 0, 0, 1, 2, 3, 4]) },
    { path: 'EncryptedPackage', data: pkg },
  ]);
  await rejectsWith(open(extensible, { password: PASSWORD }), 'unsupported', 'the extensible scheme');
  assert.deepEqual(encryptionOf(extensible), { format: 'ooxml' });
  // The password opens the file, and what comes out is no package.
  const notZip = refOk('agile-encrypt', { password: passwordArg(PASSWORD), spin: 10 }, new TextEncoder().encode('This is not a ZIP file at all, only some text.'));
  await rejectsWith(open(notZip, { password: PASSWORD }), 'damaged', 'an encrypted file that holds no package');
});

test('open: the opening of an encrypted file can be called off', async () => {
  const bytes = await ownXlsx();
  await rejectsWith(open(bytes, { password: PASSWORD, signal: { aborted: true } }), 'cancelled', 'called off before it started');
});

test('encryptionOf: files that are not encrypted, and files that are no files, are undefined', async () => {
  assert.equal(encryptionOf(await minimalXlsx()), undefined);
  assert.equal(encryptionOf(new Uint8Array(0)), undefined);
  assert.equal(encryptionOf(new Uint8Array(600)), undefined);
  assert.equal(encryptionOf(simpleDoc('Text\r').build()), undefined);
  assert.equal(encryptionOf(writeCfb([{ path: 'Workbook', data: workbookStream({ sheets: [{ name: 'S' }] }) }])), undefined);
});

// ------------------------------------------------------------------ XOR keys

refTest('xor: key, verifier and both arrays are what the reference computes another way, for passwords of every length', () => {
  const passwords = ['a', 'ab', 'Abc', 'test', 'Pass1', 'secret', 'Sekiz888', 'ninechars', '0123456789', 'elevenchars', 'twelve_chars', 'thirteenchars', 'fourteen-chars', 'fifteen-chars..', PASSWORD, 'ÿ~\u0001', 'ĞÜŞ'];
  for (const password of passwords) {
    const expected = JSON.parse(new TextDecoder().decode(refOk('xor-key', { password: passwordArg(password) }))) as { key: number; verifier: number; xls: string; doc: string };
    const chars = xorPasswordBytes(utf16le(password));
    assert.deepEqual(
      { key: xorKey(chars), verifier: xorVerifier(chars), xls: toHex(xorArray(chars, 2)), doc: toHex(xorArray(chars, 7)) },
      expected,
      `password of ${password.length} characters`,
    );
  }
});

test('xor: the verifier of "test" is the one every description of the scheme names', () => {
  assert.equal(xorVerifier(xorPasswordBytes(utf16le('test'))), 0xcbeb);
});

// ------------------------------------------------------------------ xls

/** A workbook whose stream is longer than two blocks of key stream, with a second sheet behind it. */
function xlsStream(): Uint8Array {
  const records = [];
  for (let row = 0; row < 120; row++) records.push(number(row, 0, row * 1.5 + 0.25), number(row, 1, 1000000 - row));
  return workbookStream({
    sheets: [
      { name: 'Gizli', records: cat(...records) },
      { name: 'İkinci', records: number(0, 0, 42) },
    ],
  });
}

const cellValues = async (file: OfficeFile): Promise<unknown[]> => {
  const workbook = await file.workbook();
  const out: unknown[] = workbook.sheets.map((sheet) => sheet.name);
  for (const row of [0, 1, 63, 119]) out.push(valueAt(workbook.sheets[0], row, 0), valueAt(workbook.sheets[0], row, 1));
  out.push(valueAt(workbook.sheets[1], 0, 0));
  return out;
};

const XLS_METHODS: [method: string, bits: number, scheme: string, keyBits: number][] = [
  ['xor', 0, 'xor', 16],
  ['rc4', 0, 'rc4', 40],
  ['capi', 128, 'rc4-cryptoapi', 128],
  ['capi', 40, 'rc4-cryptoapi', 40],
  ['capi', 56, 'rc4-cryptoapi', 56],
];

for (const [method, bits, scheme, keyBits] of XLS_METHODS) {
  refTest(`xls: a workbook the reference protected with ${method}${bits ? ` (${bits} bits)` : ''} is opened with its password, read-only`, async () => {
    const stream = xlsStream();
    assert.ok(stream.length > 3000, 'the stream is longer than two blocks of 1024 bytes');
    const plainFile = writeCfb([{ path: 'Workbook', data: stream }]);
    const expected = await cellValues(await open(plainFile));
    assert.deepEqual(expected.slice(0, 4), ['Gizli', 'İkinci', 0.25, 1000000]);

    const password = method === 'xor' ? 'Parola12' : PASSWORD;
    const encrypted = refOk('xls-encrypt', { password: passwordArg(password), method, bits: bits || 128, seed: `${method}-${bits}` }, stream);
    assert.ok(encrypted.length > stream.length, 'the reference put a FILEPASS record in');
    const bytes = writeCfb([{ path: 'Workbook', data: encrypted }]);
    assert.deepEqual(encryptionOf(bytes), { format: 'xls' });

    const required = await rejectsWith(open(bytes), 'required', 'no password');
    assert.equal(required.format, 'xls');
    await rejectsWith(open(bytes, { password: `${password}x` }), 'wrong', 'a wrong password');
    await rejectsWith(open(bytes, { password: password.toLowerCase() }), 'wrong', 'the password in lower case');

    const file = await open(bytes, { password });
    assert.deepEqual([file.format, file.passwordProtected, file.encryption && file.encryption.scheme, file.encryption && file.encryption.keyBits], ['xls', true, scheme, keyBits]);
    assert.deepEqual(await cellValues(file), expected, 'every cell is what it was before the encryption');

    // An xls is not written: saving gives the bytes it was opened from, still encrypted.
    assert.equal(await file.save(), bytes);
    await assert.rejects(file.save({ password: null }), /cannot be changed, set or taken away/);
    await assert.rejects(file.save({ password: 'x' }), /cannot be changed, set or taken away/);
  });
}

refTest('xls: a protected workbook saved as xlsx, or converted to be edited, is encrypted with the same password', async () => {
  const stream = xlsStream();
  const bytes = writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'capi', bits: 128 }, stream) }]);
  const file = await open(bytes, { password: PASSWORD });
  const expected = await cellValues(file);

  const xlsx = await file.saveAsXlsx!();
  assert.ok(isCfb(xlsx), 'saveAsXlsx() gives an encrypted file');
  assert.deepEqual(await cellValues(await open(refDecrypt(xlsx, PASSWORD))), expected);
  assert.ok(isZip(await file.saveAsXlsx!({ password: null })), 'without a password only when asked');

  // What mount({ convert: true }) and the bridge do.
  const opened = await openConverted(file, (converted) => open(converted));
  assert.deepEqual([opened.file.format, opened.file.passwordProtected, opened.conversion && opened.conversion.from], ['xlsx', true, 'xls']);
  assert.equal(opened.file.encryption && opened.file.encryption.scheme, 'rc4-cryptoapi', 'the converted file says how its original was encrypted');
  const saved = await opened.file.save();
  assert.ok(isCfb(saved), 'the converted file is saved encrypted');
  assert.deepEqual(await cellValues(await open(refDecrypt(saved, PASSWORD))), expected);
  assert.deepEqual(await cellValues(await open(saved, { password: PASSWORD })), expected);
  // The bytes of the conversion are encrypted as well: nothing made from a protected file is given out plain.
  const conversion = await convertForEditing(file);
  assert.ok(conversion && isCfb(conversion.bytes) && conversion.encrypted === true, 'convertForEditing() gives an encrypted file');
  assert.deepEqual(await cellValues(await open(conversion.bytes, { password: PASSWORD })), expected);
});

refTest('xls: a workbook of Excel 95 (BIFF5) with XOR obfuscation is opened', async () => {
  const stream = workbookStream({ version: 0x0500, sheets: [{ name: 'Old', records: cat(number(0, 0, 7.5), number(1, 0, -3)) }] });
  // BIFF5 has no type in its FILEPASS: the record is the key and the verifier alone.
  const encrypted = refOk('xls-encrypt', { password: passwordArg('abc'), method: 'xor' }, stream);
  const bof = 4 + (encrypted[2] | (encrypted[3] << 8));
  assert.deepEqual([...encrypted.subarray(bof, bof + 4)], [0x2f, 0, 4, 0], 'the reference wrote a FILEPASS of four bytes');
  const bytes = writeCfb([{ path: 'Book', data: encrypted }]);
  const file = await open(bytes, { password: 'abc' });
  const workbook = await file.workbook();
  assert.deepEqual([workbook.sheets[0].name, valueAt(workbook.sheets[0], 0, 0), valueAt(workbook.sheets[0], 1, 0)], ['Old', 7.5, -3]);
});

test('xls: a FILEPASS record of a kind that is not known, or cut short, is told and nothing loops', async () => {
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 1) }] });
  const bofEnd = 4 + (stream[2] | (stream[3] << 8));
  const withPass = (body: number[]): Uint8Array => {
    const out = new Uint8Array(stream.length + 4 + body.length);
    out.set(stream.subarray(0, bofEnd));
    out.set([0x2f, 0, body.length, 0, ...body], bofEnd);
    out.set(stream.subarray(bofEnd), bofEnd + 4 + body.length);
    return writeCfb([{ path: 'Workbook', data: out }]);
  };
  await rejectsWith(open(withPass([7, 0, 1, 2, 3, 4]), { password: 'x' }), 'unsupported', 'type 7');
  await rejectsWith(open(withPass([1, 0, 1, 0, 1, 0, 9, 9]), { password: 'x' }), 'damaged', 'an RC4 header of eight bytes');
  await rejectsWith(open(withPass([1, 0, 9, 0, 9, 0, 0, 0]), { password: 'x' }), 'unsupported', 'a header of version 9.9');
  await rejectsWith(open(withPass([0, 0, 1]), { password: 'x' }), 'damaged', 'an XOR record of three bytes');
  await rejectsWith(open(withPass([1, 0, 2, 0, 2, 0, 4, 0, 0, 0, 0xff, 0xff, 0xff, 0x7f]), { password: 'x' }), 'damaged', 'a CryptoAPI header that says it is 2 GB long');
  // An XOR password of sixteen characters cannot be: Excel takes fifteen.
  await rejectsWith(open(withPass([0, 0, 1, 2, 3, 4]), { password: 'x'.repeat(16) }), 'wrong', 'sixteen characters');
});

// ------------------------------------------------------------------ doc

const DOC_PARAGRAPHS = Array.from({ length: 120 }, (_, i) => `Paragraf ${i + 1}: gizli metin, şifreli belge.`);
/** A document whose text is in a piece for every paragraph: its table stream is longer than two blocks of 512 bytes as well. */
function longDoc(): Uint8Array {
  const builder = new DocBuilder();
  for (const paragraph of DOC_PARAGRAPHS) builder.text(`${paragraph}\r`);
  builder.papxAll(`${DOC_PARAGRAPHS.join('\r')}\r`);
  return builder.build();
}
const docTexts = async (file: OfficeFile): Promise<string[]> => (await file.document!()).body.blocks.map((block) => paragraphText(block as Paragraph));

const DOC_METHODS: [method: string, bits: number, scheme: string, keyBits: number][] = [
  ['xor', 0, 'xor', 16],
  ['rc4', 0, 'rc4', 40],
  ['capi', 128, 'rc4-cryptoapi', 128],
  ['capi', 40, 'rc4-cryptoapi', 40],
];

for (const [method, bits, scheme, keyBits] of DOC_METHODS) {
  refTest(`doc: a document the reference protected with ${method}${bits ? ` (${bits} bits)` : ''} is opened with its password, read-only`, async () => {
    const plainFile = longDoc();
    const expected = await docTexts(await open(plainFile));
    assert.equal(expected.length, 120, 'a hundred and twenty paragraphs');
    assert.equal(expected[119], 'Paragraf 120: gizli metin, şifreli belge.');
    assert.ok(CfbFile.read(plainFile).stream('WordDocument').length > 1600, 'the text is longer than three blocks of 512 bytes');
    assert.ok(CfbFile.read(plainFile).stream('1Table').length > 1100, 'the table stream is longer than two blocks of 512 bytes');

    const password = method === 'xor' ? 'Parola12' : PASSWORD;
    const bytes = refOk('doc-encrypt', { password: passwordArg(password), method, bits: bits || 128, seed: `doc-${method}-${bits}` }, plainFile);
    assert.deepEqual(encryptionOf(bytes), { format: 'doc' });
    const required = await rejectsWith(open(bytes), 'required', 'no password');
    assert.equal(required.format, 'doc');
    await rejectsWith(open(bytes, { password: `${password}!` }), 'wrong', 'a wrong password');

    const file = await open(bytes, { password });
    assert.deepEqual([file.format, file.passwordProtected, file.encryption && file.encryption.scheme, file.encryption && file.encryption.keyBits], ['doc', true, scheme, keyBits]);
    assert.deepEqual(await docTexts(file), expected, 'every paragraph is what it was before the encryption');
    assert.equal(await file.save(), bytes, 'saving gives the bytes that were opened');
    await assert.rejects(file.save({ password: null }), /cannot be changed, set or taken away/);

    // As a docx it keeps its password.
    const docx = await file.saveAsDocx!();
    assert.ok(isCfb(docx), 'saveAsDocx() gives an encrypted file');
    assert.deepEqual(await docTexts(await open(refDecrypt(docx, password))), expected);
    const opened = await openConverted(file, (converted) => open(converted));
    assert.deepEqual([opened.file.format, opened.file.passwordProtected], ['docx', true]);
    assert.ok(isCfb(await opened.file.save()), 'the converted document is saved encrypted');
  });
}

for (const method of ['rc4', 'capi', 'xor']) {
  refTest(`doc: the streams of a document protected with ${method} come out byte for byte as they went in`, () => {
    const plainFile = longDoc();
    const password = 'Parola12';
    const plain = CfbFile.read(plainFile);
    const sealed = CfbFile.read(refOk('doc-encrypt', { password: passwordArg(password), method, bits: 128, seed: `bytes-${method}` }, plainFile));
    const word = plain.stream('WordDocument');
    const table = plain.stream('1Table');
    assert.notEqual(toHex(sealed.stream('WordDocument').subarray(word.length - 400)), toHex(word.subarray(word.length - 400)), 'the text is not readable in the protected file');
    const out = decryptDocStreams({ word: sealed.stream('WordDocument'), table: sealed.stream('1Table'), data: new Uint8Array(0) }, utf16le(password));
    // The start of the header was never encrypted; what follows it, from byte 68 on, was.
    assert.equal(toHex(sealed.stream('WordDocument').subarray(18, 68)), toHex(word.subarray(18, 68)), 'bytes 18 to 67 are plain in the file');
    assert.equal(toHex(out.word.subarray(18, 152)), toHex(word.subarray(18, 152)), 'the header up to the list of offsets');
    assert.equal(toHex(out.word.subarray(word.length - 1200)), toHex(word.subarray(word.length - 1200)), 'the text');
    assert.equal(out.word[11] & 0x81, 0, 'the document that comes out does not say it is encrypted');
    // With RC4 the header of the encryption sits in front of the table stream.
    const header = out.table.length - table.length;
    assert.equal(header > 0, method !== 'xor');
    assert.equal(toHex(out.table.subarray(header)), toHex(table), 'the table stream');
    assert.equal(toHex(out.table.subarray(0, header)), toHex(sealed.stream('1Table').subarray(0, header)), 'the header of the encryption is left as it is in the file');
  });
}

test('doc: an encryption header that is not inside the table stream, or of a version that is not known, is told', async () => {
  const build = (lKey: number, table: number[]) =>
    simpleDoc('Text\r').build({
      flags: 0x0100,
      patch: (streams) => {
        new DataView(streams.word.buffer, streams.word.byteOffset).setUint32(14, lKey, true);
        streams.table.set(table);
      },
    });
  await rejectsWith(open(build(0x7fffffff, []), { password: 'x' }), 'damaged', 'a header of 2 GB');
  await rejectsWith(open(build(2, []), { password: 'x' }), 'damaged', 'a header of two bytes');
  await rejectsWith(open(build(8, [5, 0, 5, 0]), { password: 'x' }), 'unsupported', 'a header of version 5.5');
  await rejectsWith(open(build(8, [1, 0, 1, 0]), { password: 'x' }), 'damaged', 'an RC4 header of eight bytes');
});

test('rtf and plain doc: the bytes they were opened from cannot be asked for with a password, their docx can', async () => {
  const rtf = await open(new TextEncoder().encode('{\\rtf1\\ansi Hello\\par}'));
  assert.equal(rtf.passwordProtected, false);
  await assert.rejects(rtf.save({ password: 'x' }), /A RTF file is saved as the bytes it was opened from/);
  assert.ok(isCfb(await rtf.saveAsDocx!({ password: 'Yeni' })), 'an RTF document can be saved as an encrypted docx');
  assert.ok(isZip(await rtf.saveAsDocx!()), 'and is not, unless asked');
});
