// Password-protected files, after the independent review: what a file without an integrity check
// says of itself, that nothing derived from an encrypted file is ever written without its
// password, that large packages are worked on in slices, that a file without a password costs
// nothing extra to open, the password Office uses when the user gave none, passwords of XOR
// obfuscation in a code page, and the bytes of the streams nobody could check against Office.
//
// The encrypted files here are made by the engine itself (the tests run without the reference
// tools); what needs the reference script says so (refTest).

import assert from 'node:assert/strict';
import { test } from 'node:test';
import v8 from 'node:v8';
import vm from 'node:vm';
import { CfbFile } from '../src/cfb';
import { writeCfb } from '../src/cfb-write';
import { AgileEncryptor, agileDecryptPackage, agilePackageSize, agileUnlock, parseAgileInfo } from '../src/crypto/agile';
import { toHex, utf16le } from '../src/crypto/bytes';
import { PasswordError } from '../src/crypto/errors';
import { hmac, type HashName } from '../src/crypto/hash';
import { decryptDocStreams, decryptXlsStream, xorPasswordCandidates, xorVerifier } from '../src/crypto/legacy';
import { decryptPackage, encryptPackage } from '../src/crypto/ooxml';
import { PasswordSession } from '../src/crypto/session';
import { parseStandardInfo } from '../src/crypto/standard';
import { SaveController } from '../src/embed/save-to';
import { convertForEditing, openConverted } from '../src/formats/convert';
import { DEFAULT_PASSWORD, encryptionOf, open, type OfficeFile } from '../src/index';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { DocBuilder, simpleDoc } from './doc-binary-fixtures';
import { buildDocx as docxFixture } from './docx-fixtures';
import { minimalXlsx } from './fixtures';
import { passwordArg, refOk, refTest } from './reftools';
import { cat, number, rec, u16, workbookStream } from './xls-fixtures';

const PASSWORD = 'Gizli-Parola-42';
const isCfb = (bytes: Uint8Array): boolean => toHex(bytes.subarray(0, 8)) === 'd0cf11e0a1b11ae1';
const isZip = (bytes: Uint8Array): boolean => bytes[0] === 0x50 && bytes[1] === 0x4b;

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

/** A package encrypted quickly: the rounds of key derivation are not what these tests are about. */
async function sealedFile(plain: Uint8Array, password = PASSWORD, spinCount = 100): Promise<Uint8Array> {
  const session = PasswordSession.of(password);
  try {
    return await encryptPackage(plain, session, { spinCount });
  } finally {
    session.forget();
  }
}

let xlsxOnce: Uint8Array | undefined;
const encryptedXlsx = async (): Promise<Uint8Array> => (xlsxOnce ??= await sealedFile(await minimalXlsx()));

/** The two streams of an encrypted file, to change them. */
const streamsOf = (file: Uint8Array): { info: Uint8Array; pkg: Uint8Array } => {
  const cfb = CfbFile.read(file);
  return { info: cfb.stream('EncryptionInfo'), pkg: cfb.stream('EncryptedPackage') };
};
const fileOf = (info: Uint8Array, pkg: Uint8Array): Uint8Array =>
  writeCfb([
    { path: 'EncryptionInfo', data: info },
    { path: 'EncryptedPackage', data: pkg },
  ]);
const infoText = (info: Uint8Array): string => new TextDecoder().decode(info.subarray(8));
const infoOf = (version: Uint8Array, text: string): Uint8Array => {
  const body = new TextEncoder().encode(text);
  const out = new Uint8Array(8 + body.length);
  out.set(version.subarray(0, 8));
  out.set(body, 8);
  return out;
};

/** An encrypted file whose integrity check was taken out, by someone who does not know the password. */
async function stripped(change?: (pkg: Uint8Array) => void): Promise<Uint8Array> {
  const { info, pkg } = streamsOf(await encryptedXlsx());
  const text = infoText(info);
  assert.match(text, /<dataIntegrity [^>]*\/>/, 'the engine writes an integrity check');
  const changed = pkg.slice();
  if (change) change(changed);
  return fileOf(infoOf(info, text.replace(/<dataIntegrity [^>]*\/>/, '')), changed);
}

// ------------------------------------------------------------------ B1: no integrity check

test('integrity: a file whose integrity check was taken out opens, and says that it has none', async () => {
  const file = await open(await stripped(), { password: PASSWORD });
  assert.equal(file.format, 'xlsx');
  assert.equal(file.encryption && file.encryption.integrity, false);
  assert.deepEqual(
    (file.warnings || []).map((warning) => warning.code),
    ['no-integrity'],
  );
  assert.match((file.warnings || [])[0].message, /no integrity check/);
  // A file that has its check says nothing.
  const whole = await open(await encryptedXlsx(), { password: PASSWORD });
  assert.deepEqual([whole.encryption && whole.encryption.integrity, whole.warnings], [true, undefined]);
  // And a file that was never encrypted has nothing to say either.
  assert.equal((await open(await minimalXlsx())).warnings, undefined);
});

test('integrity: with requireIntegrity a file without a check is damaged; one with a check opens', async () => {
  const error = await rejectsWith(open(await stripped(), { password: PASSWORD, requireIntegrity: true }), 'damaged', 'no check, and one is required');
  assert.match(error.message, /no integrity check/);
  assert.equal(error.format, 'ooxml');
  assert.equal((await open(await encryptedXlsx(), { password: PASSWORD, requireIntegrity: true })).format, 'xlsx');
  // The wrong password is still told as that, not as damage.
  await rejectsWith(open(await stripped(), { password: 'x', requireIntegrity: true }), 'wrong', 'a wrong password');
  // At the level of the package as well.
  await rejectsWith(decryptPackage(CfbFile.read(await stripped()), utf16le(PASSWORD), { requireIntegrity: true }), 'damaged', 'decryptPackage');
});

test('integrity: content that was changed in a file without a check, and no longer reads as a package, is told as damage of that kind', async () => {
  // The end of the package holds the directory of the ZIP file: with its last block changed nothing can be read.
  const changed = await stripped((pkg) => {
    pkg[pkg.length - 5] ^= 0xff;
  });
  const error = await rejectsWith(open(changed, { password: PASSWORD }), 'damaged', 'a package that was changed');
  assert.match(error.message, /no integrity check/);
  assert.equal(error.format, 'ooxml');
});

test('integrity: what the engine writes always carries its check, also for a file that came without one', async () => {
  const file = await open(await stripped(), { password: PASSWORD });
  const saved = await file.save();
  assert.deepEqual(encryptionOf(saved) && encryptionOf(saved)!.encryption!.integrity, true);
  const again = await open(saved, { password: PASSWORD, requireIntegrity: true });
  assert.equal(again.warnings, undefined);
});

test('integrity: the warning has words in both languages', () => {
  assert.match(STRINGS_EN.noIntegrityCheck, /integrity check/);
  assert.match(STRINGS_TR.noIntegrityCheck, /bütünlük denetimi/);
});

// ------------------------------------------------------------------ B2: never a plain file from an encrypted one

type Outcome = 'encrypted' | 'refused' | 'original' | 'PLAIN';

/** What a way of saving gives: an encrypted compound file, a refusal, the bytes that were opened - or a plain package. */
async function outcome(work: () => Promise<Uint8Array | undefined>, original?: Uint8Array): Promise<Outcome> {
  let bytes: Uint8Array | undefined;
  try {
    bytes = await work();
  } catch {
    return 'refused';
  }
  if (!bytes) return 'refused';
  if (original && bytes === original) return 'original';
  if (isZip(bytes)) return 'PLAIN';
  assert.ok(isCfb(bytes) && encryptionOf(bytes) !== undefined, 'what was saved is an encrypted compound file');
  return 'encrypted';
}

/** Every way a file leaves, for a file as open() gave it. */
async function everyWay(file: OfficeFile, original: Uint8Array): Promise<Record<string, Outcome>> {
  const out: Record<string, Outcome> = {};
  out.save = await outcome(() => file.save(), original);
  if (file.saveAsXlsx) out.saveAsXlsx = await outcome(() => file.saveAsXlsx!(), original);
  if (file.saveAsDocx) out.saveAsDocx = await outcome(() => file.saveAsDocx!(), original);
  out.convertForEditing = await outcome(async () => {
    const conversion = await convertForEditing(file);
    return conversion ? conversion.bytes : undefined;
  });
  out.openConverted = await outcome(async () => (await openConverted(file, (bytes: Uint8Array) => open(bytes))).file.save(), original);
  // The Save button and auto-save of a view: both go through the controller.
  const saves: Uint8Array[] = [];
  const controller = new SaveController({
    target: { kind: 'function', save: async (bytes: Uint8Array) => void saves.push(bytes) },
    writable: () => true,
    commit: () => true,
    bytes: () => file.save(),
    bytesAsIs: () => file.save(),
    pending: () => false,
    readOnly: () => false,
    autoSave: { delayMs: 250, maxWaitMs: 250, unconditional: true },
    info: () => ({ format: file.format }),
    strings: () => STRINGS_EN,
    notice: () => undefined,
    etag: null,
  } as unknown as ConstructorParameters<typeof SaveController>[0]);
  controller.changed();
  await controller.saveNow();
  out.saveNow = await outcome(async () => saves[0], original);
  saves.length = 0;
  controller.changed();
  for (let i = 0; i < 60 && saves.length === 0; i++) await new Promise((resolve) => setTimeout(resolve, 10));
  out.autoSave = await outcome(async () => saves[0], original);
  controller.destroy();
  return out;
}

const SOURCES: [name: string, make: () => Promise<{ bytes: Uint8Array; password: string }>, needsReference: boolean][] = [
  ['an encrypted xlsx', async () => ({ bytes: await encryptedXlsx(), password: PASSWORD }), false],
  ['an encrypted docx', async () => ({ bytes: await sealedFile(await docxFixture({ body: '<w:p><w:r><w:t>Gizli</w:t></w:r></w:p>' })), password: PASSWORD }), false],
  [
    'a protected xls',
    async () => ({
      bytes: writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'capi', bits: 128 }, workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 5) }] })) }]),
      password: PASSWORD,
    }),
    true,
  ],
  ['a protected doc', async () => ({ bytes: refOk('doc-encrypt', { password: passwordArg(PASSWORD), method: 'rc4' }, simpleDoc('Gizli metin\r').build()), password: PASSWORD }), true],
];

for (const [name, make, needsReference] of SOURCES) {
  (needsReference ? refTest : test)(`never plain: ${name} leaves encrypted or not at all, by every way of saving, before and after it is closed`, async () => {
    const { bytes, password } = await make();
    const file = await open(bytes, { password });
    const before = await everyWay(file, bytes);
    for (const way of Object.keys(before)) assert.notEqual(before[way], 'PLAIN', `${way}, while the file is open`);
    // doc and xls are saved as the bytes they were opened from, which are encrypted; everything else is encrypted anew.
    const legacy = file.format === 'doc' || file.format === 'xls';
    assert.deepEqual([before.save, before.saveNow, before.autoSave], legacy ? ['original', 'original', 'original'] : ['encrypted', 'encrypted', 'encrypted']);
    if (legacy) assert.deepEqual([before.convertForEditing, before.openConverted], ['encrypted', 'encrypted']);

    file.close();
    const after = await everyWay(file, bytes);
    for (const way of Object.keys(after)) assert.notEqual(after[way], 'PLAIN', `${way}, after the file was closed`);
    if (legacy) {
      // The file could not be converted (its password is gone): what comes back is the file as it was, which saves its own bytes.
      assert.deepEqual([after.save, after.convertForEditing, after.openConverted], ['original', 'refused', 'original']);
      assert.deepEqual([after.saveAsXlsx || after.saveAsDocx], ['refused']);
    } else assert.deepEqual([after.save, after.saveNow, after.autoSave], ['refused', 'refused', 'refused']);
  });
}

refTest('never plain: a converted file whose original was closed afterwards keeps the password; converting a closed original is refused', async () => {
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 5) }] });
  const bytes = writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'rc4' }, stream) }]);
  const file = await open(bytes, { password: PASSWORD });
  const first = await openConverted(file, (converted: Uint8Array) => open(converted));
  file.close();
  assert.ok(isCfb(await first.file.save()), 'converted before the original was closed: saved encrypted');
  const second = await openConverted(file, (converted: Uint8Array) => open(converted));
  assert.equal(second.conversion, undefined, 'nothing was converted');
  assert.match(second.failed || '', /password of this file was forgotten/);
  assert.equal(second.file, file, 'the file that comes back is the original');
  assert.equal(second.file.passwordProtected, true);
});

test('never plain: a file that was given a password and then closed is not converted or saved without it either', async () => {
  const rtf = await open(new TextEncoder().encode('{\\rtf1\\ansi Hello\\par}'));
  rtf.setPassword(PASSWORD);
  const conversion = await convertForEditing(rtf);
  assert.ok(conversion && isCfb(conversion.bytes), 'the bytes of a conversion are encrypted when the file has a password');
  assert.equal(conversion && conversion.encrypted, true);
  rtf.close();
  await assert.rejects(convertForEditing(rtf), /password of this file was forgotten/);
  await assert.rejects(rtf.saveAsDocx!(), /password of this file was forgotten/);
  const opened = await openConverted(rtf, (bytes: Uint8Array) => open(bytes));
  assert.equal(opened.file, rtf);
  // Taking the password away is still something that can be asked for.
  rtf.setPassword(null);
  const plain = await convertForEditing(rtf);
  assert.ok(plain && isZip(plain.bytes) && plain.encrypted === undefined, 'without a password the conversion is a plain package');
});

// ------------------------------------------------------------------ B3: slices

/** Options that count the pauses and make every look at the clock end a slice. */
function counting(): { options: { spinCount: number; sliceMs: number; now(): number; pause(): Promise<void>; signal: { aborted: boolean } }; pauses(): number; stopAt(count: number): void } {
  let clock = 0;
  let pauses = 0;
  let stop = Infinity;
  const signal = { aborted: false };
  return {
    options: {
      spinCount: 0,
      sliceMs: 1,
      now: () => (clock += 2),
      pause: async () => {
        if (++pauses >= stop) signal.aborted = true;
      },
      signal,
    },
    pauses: () => pauses,
    stopAt: (count) => void (stop = count),
  };
}

test('slices: the HMAC of a large package is computed in slices, when it is written and when it is read', async () => {
  // 6 MB: 1465 segments of cipher (a turn every 16 of them), and an HMAC over all of it.
  const plain = new Uint8Array(6 * 1024 * 1024).map((_, i) => i & 0xff);
  const segments = Math.ceil(plain.length / 4096 / 16);
  const writing = counting();
  const session = PasswordSession.of('pw');
  const file = await encryptPackage(plain, session, writing.options);
  // More turns than the cipher alone takes: the HMAC gives the page its turn as well, at least every megabyte.
  assert.ok(writing.pauses() >= segments + 6, `${writing.pauses()} pauses while writing; the cipher alone has ${segments}`);
  const reading = counting();
  const out = await decryptPackage(CfbFile.read(file), utf16le('pw'), reading.options);
  assert.equal(out.bytes.length, plain.length);
  assert.ok(reading.pauses() >= segments + 6, `${reading.pauses()} pauses while reading; the cipher alone has ${segments}`);
});

test('slices: reading and writing a large package can be called off in the middle of the HMAC', async () => {
  const plain = new Uint8Array(6 * 1024 * 1024);
  const file = await encryptPackage(plain, PasswordSession.of('pw'), { spinCount: 0 });
  // Reading starts with the HMAC: called off at the third pause, long before the cipher.
  const reading = counting();
  reading.stopAt(3);
  await rejectsWith(decryptPackage(CfbFile.read(file), utf16le('pw'), reading.options), 'cancelled', 'reading');
  assert.ok(reading.pauses() <= 4, `reading stopped after ${reading.pauses()} pauses`);
  // Writing ends with it: called off after the cipher is through.
  const writing = counting();
  const segments = Math.ceil(plain.length / 4096 / 16);
  writing.stopAt(segments + 3);
  await rejectsWith(encryptPackage(plain, PasswordSession.of('pw'), writing.options), 'cancelled', 'writing');
  assert.ok(writing.pauses() <= segments + 4, `writing stopped after ${writing.pauses()} pauses`);
});

test('slices: no stretch of work on a large package is long without a pause', async () => {
  const plain = new Uint8Array(24 * 1024 * 1024);
  let last = Date.now();
  let worst = 0;
  const work = {
    spinCount: 0,
    pause: async () => {
      worst = Math.max(worst, Date.now() - last);
      await new Promise((resolve) => setTimeout(resolve, 0));
      last = Date.now();
    },
  };
  const file = await encryptPackage(plain, PasswordSession.of('pw'), work);
  await decryptPackage(CfbFile.read(file), utf16le('pw'), work);
  // A slice is 12 ms; unsliced, the HMAC of 24 MB took some 450 ms on its own. Generous: this guards against a missing pause, it does not measure speed.
  assert.ok(worst < 250, `the longest stretch without a pause was ${worst} ms`);
});

// ------------------------------------------------------------------ B4: files without a password

/** Counts the reads of whole streams, by path, while `work` runs. */
async function streamReads(work: () => Promise<void>): Promise<Record<string, number>> {
  const reads: Record<string, number> = {};
  const real = CfbFile.prototype.stream;
  CfbFile.prototype.stream = function (this: CfbFile, path: string): Uint8Array {
    reads[path] = (reads[path] || 0) + 1;
    return real.call(this, path);
  };
  try {
    await work();
  } finally {
    CfbFile.prototype.stream = real;
  }
  return reads;
}

function bigXls(megabytes: number): Uint8Array {
  const one = number(0, 0, 1);
  const count = Math.floor((megabytes * 1024 * 1024) / one.length);
  const records = new Uint8Array(one.length * count);
  for (let i = 0; i < count; i++) {
    records.set(one, i * one.length);
    const row = Math.floor(i / 30);
    records[i * one.length + 4] = row & 0xff;
    records[i * one.length + 5] = row >> 8;
    records[i * one.length + 6] = i % 30;
  }
  return writeCfb([{ path: 'Workbook', data: workbookStream({ sheets: [{ name: 'S', records }] }) }]);
}

test('unprotected files: opening an xls or a doc without a password reads no whole stream, and reading its model reads each once', async () => {
  const xls = writeCfb([{ path: 'Workbook', data: workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 5) }] }) }]);
  let file: OfficeFile | undefined;
  assert.deepEqual(await streamReads(async () => void (file = await open(xls))), {}, 'open() of an xls');
  assert.deepEqual(await streamReads(async () => void (await file!.workbook())), { Workbook: 1 }, 'workbook() of an xls');
  const doc = simpleDoc('Text\r').build();
  assert.deepEqual(await streamReads(async () => void (file = await open(doc))), {}, 'open() of a doc');
  assert.deepEqual(await streamReads(async () => void (await file!.document!())), { WordDocument: 1, '1Table': 1 }, 'document() of a doc');
  assert.deepEqual(await streamReads(async () => void encryptionOf(xls)), {}, 'encryptionOf() of an xls');
  assert.deepEqual(await streamReads(async () => void encryptionOf(doc)), {}, 'encryptionOf() of a doc');
});

test('unprotected files: an open xls of 12 MB holds nothing but the bytes it was given', async () => {
  v8.setFlagsFromString('--expose-gc');
  const collect = vm.runInNewContext('gc') as () => void;
  const bytes = bigXls(12);
  collect();
  const before = process.memoryUsage().arrayBuffers;
  const file = await open(bytes);
  collect();
  const held = process.memoryUsage().arrayBuffers - before;
  assert.equal(file.format, 'xls');
  // A copy of the stream would be 12 MB.
  assert.ok(held < 2 * 1024 * 1024, `${(held / 1048576).toFixed(1)} MB are held after open()`);
});

// ------------------------------------------------------------------ B5: the default password

test('default password: a file encrypted with it opens without one and says that anyone can open it', async () => {
  const bytes = await sealedFile(await minimalXlsx(), DEFAULT_PASSWORD);
  const file = await open(bytes);
  assert.deepEqual([file.passwordProtected, file.encryption && file.encryption.defaultPassword], [true, true]);
  // Named, it is the password as well.
  assert.equal((await open(bytes, { password: DEFAULT_PASSWORD })).encryption!.defaultPassword, true);
  // A file with a password of its own says nothing of the kind.
  const own = await open(await encryptedXlsx(), { password: PASSWORD });
  assert.equal('defaultPassword' in own.encryption!, false);
});

test('default password: a password that was given and is wrong is wrong; the default one is not tried behind it', async () => {
  const bytes = await sealedFile(await minimalXlsx(), DEFAULT_PASSWORD);
  await rejectsWith(open(bytes, { password: 'something else' }), 'wrong', 'a wrong password for a file anyone can open');
  // The same for a doc or an xls, where trying costs nothing.
  assert.equal((await open(bytes)).format, 'xlsx', 'without a password it still opens');
});

// ------------------------------------------------------------------ B6: XOR passwords in a code page

test('xor: a password is tried as Excel would have written it: low bytes, Windows-1254, Windows-1252', () => {
  const forms = (password: string): string[] => xorPasswordCandidates(utf16le(password)).map(toHex);
  assert.deepEqual(forms('abc'), ['616263'], 'ASCII is the same in all three');
  // ş U+015F, ğ U+011F, ı U+0131: FE, F0, FD in Windows-1254; not in Windows-1252 at all.
  assert.deepEqual(forms('şğı'), ['5f1f31', 'fef0fd']);
  // é U+00E9 is E9 everywhere; € U+20AC is 80 in both code pages, AC by its low byte.
  assert.deepEqual(forms('é€'), ['e9ac', 'e980']);
  // Ð U+00D0 is D0 in Windows-1252; Windows-1254 has Ğ there, and no Ð.
  assert.deepEqual(forms('Ð'), ['d0']);
  // Ž U+017D: 8E in Windows-1252 only.
  assert.deepEqual(forms('Ž'), ['7d', '8e']);
  // A character none of the code pages has leaves the first rule alone.
  assert.deepEqual(forms('密'), ['c6']);
});

test('xor: a workbook whose Turkish password was hashed in Windows-1254 opens (a stream the reviewer made)', () => {
  // BOF, FILEPASS with the key and the verifier of "şifreğı" in Windows-1254, one record, EOF.
  const stream = new Uint8Array(Buffer.from('CQgQAAAGBQABAM0HAAAAAAYAAAAvAAYAAABh/CSEMQAPAGhlbGxvIHdvcmxkIDEyMwoAAAA=', 'base64'));
  const out = decryptXlsStream(stream, utf16le('şifreğı'));
  assert.equal(out.encryption.scheme, 'xor');
  assert.throws(() => decryptXlsStream(stream, utf16le('sifregi')), (error: unknown) => error instanceof PasswordError && error.code === 'wrong');
});

refTest('xor: xls and doc files with passwords in Windows-1254 and Windows-1252 open with the password as it is typed', async () => {
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 7.5) }] });
  for (const [password, codepage] of [
    ['şifreğıİĞŞ', 'cp1254'],
    ['Pässwörd-Ðý', 'cp1252'],
    ['Çöğüş€', 'cp1254'],
  ]) {
    const xls = writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(password), method: 'xor', codepage }, stream) }]);
    const file = await open(xls, { password });
    assert.equal(file.encryption && file.encryption.scheme, 'xor', `xls, ${codepage}`);
    assert.equal((await file.workbook()).sheets[0].name, 'S');
    await rejectsWith(open(xls, { password: `${password}x` }), 'wrong', `xls, ${codepage}, a wrong password`);
    const doc = refOk('doc-encrypt', { password: passwordArg(password), method: 'xor', codepage }, simpleDoc('Gizli metin\r').build());
    assert.equal((await open(doc, { password })).encryption!.scheme, 'xor', `doc, ${codepage}`);
  }
});

// ------------------------------------------------------------------ B8e: long passwords

test('long passwords: a file that was opened with more than 255 characters is saved with them; only a new password has the limit', async () => {
  const long = 'u'.repeat(300);
  const file = await open(await sealedFile(await minimalXlsx(), long), { password: long });
  const saved = await file.save();
  assert.ok(isCfb(saved), 'saved encrypted');
  assert.equal((await open(saved, { password: long })).format, 'xlsx');
  // UNSURE: 255 is the limit of Office's own dialog. A password that is set here keeps to it.
  assert.throws(() => file.setPassword('x'.repeat(256)), /at most 255 characters/);
  await assert.rejects(file.save({ password: 'x'.repeat(256) }), /at most 255 characters/);
});

// ------------------------------------------------------------------ B7: what the first tests left open

test('data spaces: the four streams next to an encrypted package are these bytes, exactly', async () => {
  // Nobody could check these against Office: they are written from MS-OFFCRYPTO 2.1 and from what other writers
  // produce, and must not change by accident. A text is its length in bytes (4 bytes), UTF-16LE, filled up to 4 bytes.
  const cfb = CfbFile.read(await encryptedXlsx());
  const hex = (path: string): string => toHex(cfb.stream(path));
  const text = (value: string): string => toHex(utf16le(value));
  // DataSpaceVersionInfo (76 bytes).
  assert.equal(
    hex('\u0006DataSpaces/Version'),
    '3c000000' + // the name is 60 bytes
      '4d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300' + // "Microsoft.Container.DataSpaces" in UTF-16LE; 60 is a multiple of 4, nothing fills it up
      '01000000' + // reader 1.0: two 16-bit numbers
      '01000000' + // updater 1.0
      '01000000', // writer 1.0
  );
  assert.ok(hex('\u0006DataSpaces/Version').includes(text('Microsoft.Container.DataSpaces')), 'the name of the feature');

  // DataSpaceMap (112 bytes).
  assert.equal(
    hex('\u0006DataSpaces/DataSpaceMap'),
    '08000000' + // the header is 8 bytes
      '01000000' + // one entry
      '68000000' + // the entry is 0x68 bytes, this number included
      '01000000' + // it names one component
      '00000000' + // which is a stream (0), not a storage
      '20000000' + // whose name is 32 bytes
      '45006e0063007200790070007400650064005000610063006b00610067006500' + // "EncryptedPackage"
      '32000000' + // the data space it is in: a name of 50 bytes
      '5300740072006f006e00670045006e006300720079007000740069006f006e00440061007400610053007000610063006500' + // "StrongEncryptionDataSpace"
      '0000', // filled up to 4 bytes
  );

  // DataSpaceDefinition (64 bytes).
  assert.equal(
    hex('\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace'),
    '08000000' + // the header is 8 bytes
      '01000000' + // one transform
      '32000000' + // a name of 50 bytes
      '5300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00' + // "StrongEncryptionTransform"
      '0000', // filled up to 4 bytes
  );

  // The transform (200 bytes).
  assert.equal(
    hex('\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary'),
    '58000000' + // TransformInfoHeader: up to the name of the transform it is 0x58 bytes
      '01000000' + // the type of the transform: 1
      '4c000000' + // its identifier: 76 bytes
      '7b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d00' + // "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}"
      '4e000000' + // its name: 78 bytes
      '4d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00' + // "Microsoft.Container.EncryptionTransform"
      '0000' + // filled up to 4 bytes
      '01000000' + // reader 1.0: two 16-bit numbers
      '01000000' + // updater 1.0
      '01000000' + // writer 1.0
      '00000000' + // EncryptionTransformInfo: no name of an algorithm
      '00000000' + // no block size
      '00000000' + // no cipher mode
      '04000000', // and 4, which every writer puts last
  );
  // The texts are the texts they are said to be.
  for (const [path, value] of [
    ['\u0006DataSpaces/DataSpaceMap', 'EncryptedPackage'],
    ['\u0006DataSpaces/DataSpaceMap', 'StrongEncryptionDataSpace'],
    ['\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', 'StrongEncryptionTransform'],
    ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', '{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}'],
    ['\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', 'Microsoft.Container.EncryptionTransform'],
  ]) {
    assert.ok(hex(path).includes(text(value)), `${JSON.stringify(path)} holds ${JSON.stringify(value)}`);
  }
  // And the EncryptionInfo stream starts with version 4.4 and the flag of agile encryption.
  assert.equal(toHex(cfb.stream('EncryptionInfo').subarray(0, 8)), '0400040040000000');
  // Six streams, no more: the two of the encryption and these four.
  assert.equal(cfb.entries.filter((entry) => entry.type === 'stream').length, 6);
});

refTest('xls: every record that is never encrypted comes out as it went in, and everything else is decrypted (RC4 and XOR)', () => {
  // BOF and FILEPASS are in every file; UsrExcl, FileLock, InterfaceHdr, RRDInfo and RRDHead are the other five.
  const filler = (id: number, size: number): Uint8Array => rec(id, new Uint8Array(size).map((_, i) => (i * 5 + id) & 0xff));
  const head = cat(filler(0x00e1, 2), filler(0x0194, 6), filler(0x0195, 10), filler(0x0196, 14), filler(0x0138, 18), filler(0x0042, 2), filler(0x005c, 112));
  const plain = workbookStream({ head, sheets: [{ name: 'Gizli', records: cat(number(0, 0, 1), number(1, 0, 2)) }] });
  for (const method of ['rc4', 'capi', 'xor']) {
    const sealed = refOk('xls-encrypt', { password: passwordArg('Parola12'), method, bits: 128 }, plain);
    const out = decryptXlsStream(sealed, utf16le('Parola12')).bytes;
    // What went in, with the FILEPASS record the reference put in after BOF and the positions of the sheets moved by its size.
    const bofEnd = 4 + (plain[2] | (plain[3] << 8));
    const pass = sealed.subarray(bofEnd, bofEnd + 4 + (sealed[bofEnd + 2] | (sealed[bofEnd + 3] << 8)));
    const expected = cat(plain.subarray(0, bofEnd), pass, plain.subarray(bofEnd));
    for (let p = 0; p + 4 <= expected.length; p += 4 + (expected[p + 2] | (expected[p + 3] << 8))) {
      if ((expected[p] | (expected[p + 1] << 8)) === 0x0085 && p > bofEnd) new DataView(expected.buffer).setUint32(p + 4, new DataView(expected.buffer).getUint32(p + 4, true) + pass.length, true);
    }
    assert.equal(toHex(out), toHex(expected), `${method}: the decrypted stream`);
    for (const id of [0x00e1, 0x0194, 0x0195, 0x0196, 0x0138]) {
      const at = toHex(sealed).indexOf(toHex(filler(id, id === 0x00e1 ? 2 : id === 0x0194 ? 6 : id === 0x0195 ? 10 : id === 0x0196 ? 14 : 18)));
      assert.ok(at >= 0, `${method}: record 0x${id.toString(16)} is plain in the file`);
    }
    assert.equal(toHex(sealed).indexOf(toHex(filler(0x005c, 112))), -1, `${method}: a record of another kind is not`);
  }
});

refTest('doc: a protected document with a Data stream of several blocks comes out byte for byte', () => {
  const builder = new DocBuilder();
  builder.text('Görselli belge\r');
  builder.papxAll('Görselli belge\r');
  const data = new Uint8Array(1800).map((_, i) => (i * 11 + 3) & 0xff);
  builder.dataBlob(data);
  const plainFile = builder.build();
  for (const method of ['rc4', 'capi', 'xor']) {
    const sealed = CfbFile.read(refOk('doc-encrypt', { password: passwordArg('Parola12'), method, bits: 128 }, plainFile));
    assert.notEqual(toHex(sealed.stream('Data').subarray(600, 1800)), toHex(data.subarray(600, 1800)), `${method}: the Data stream is not readable in the file`);
    const out = decryptDocStreams({ word: sealed.stream('WordDocument'), table: sealed.stream('1Table'), data: sealed.stream('Data') }, utf16le('Parola12'));
    assert.equal(toHex(out.data), toHex(data), `${method}: the Data stream`);
  }
});

refTest('rc4: what the keys of the blocks are made from is zeroed when a stream is decrypted, and when the password is wrong', () => {
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 7.5) }] });
  const sealed = refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'rc4' }, stream);
  // Office 97 RC4 keeps five bytes of a digest of the password and the salt for the keys of the blocks: the only
  // array of five bytes that is ever zeroed here.
  const zeroedFives = (work: () => void): number => {
    let count = 0;
    const real = Uint8Array.prototype.fill;
    Uint8Array.prototype.fill = function (this: Uint8Array, ...args: [number, number?, number?]) {
      if (this.length === 5 && args[0] === 0 && args.length === 1) count++;
      return real.apply(this, args);
    } as typeof real;
    try {
      work();
    } finally {
      Uint8Array.prototype.fill = real;
    }
    return count;
  };
  assert.equal(zeroedFives(() => void decryptXlsStream(sealed, utf16le(PASSWORD))), 1, 'after the stream was decrypted');
  assert.equal(zeroedFives(() => assert.throws(() => decryptXlsStream(sealed, utf16le('wrong')), /not the password/)), 1, 'after a wrong password');
});

refTest('rc4 cryptoapi: a key size of 0 in the header stands for 40 bits', async () => {
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 7.5) }] });
  const xls = writeCfb([{ path: 'Workbook', data: refOk('xls-encrypt', { password: passwordArg(PASSWORD), method: 'capi', bits: 0 }, stream) }]);
  const file = await open(xls, { password: PASSWORD });
  assert.deepEqual([file.encryption!.scheme, file.encryption!.keyBits], ['rc4-cryptoapi', 40]);
  assert.equal((await file.workbook()).sheets[0].name, 'S');
});

test('xor: a password of sixteen characters is wrong even where a key of 0 and its verifier would fit', () => {
  // XOR obfuscation takes fifteen characters: for more there is no key, and a record that says "key 0" must not be opened by them.
  const password = 'sixteen-chars-16';
  const verifier = xorVerifier(new TextEncoder().encode(password));
  const stream = workbookStream({ sheets: [{ name: 'S', records: number(0, 0, 1) }] });
  const bofEnd = 4 + (stream[2] | (stream[3] << 8));
  const sealed = cat(stream.subarray(0, bofEnd), rec(0x002f, u16(0), u16(0), u16(verifier)), stream.subarray(bofEnd));
  assert.throws(() => decryptXlsStream(sealed, utf16le(password)), (error: unknown) => error instanceof PasswordError && error.code === 'wrong');
});

const standardHeader = (flags: number): Uint8Array => {
  const out = new Uint8Array(12 + 32 + 4 + 16 + 16 + 4 + 32);
  const view = new DataView(out.buffer);
  view.setUint16(0, 4, true);
  view.setUint16(2, 2, true);
  view.setUint32(4, flags, true);
  view.setUint32(8, 32, true);
  view.setUint32(12, flags, true);
  view.setUint32(20, 0x660e, true);
  view.setUint32(24, 0x8004, true);
  view.setUint32(28, 128, true);
  view.setUint32(44, 16, true);
  view.setUint32(80, 20, true);
  return out;
};

test('standard: a header that names an outside module (fExternal) is not read, whatever version it says', () => {
  assert.equal(parseStandardInfo(standardHeader(0x24)).keyBits, 128);
  assert.throws(() => parseStandardInfo(standardHeader(0x24 | 0x10)), (error: unknown) => error instanceof PasswordError && error.code === 'unsupported' && /extensible/.test(error.message));
});

/** The descriptor of a real file, with attributes replaced. */
async function withAttributes(change: (text: string) => string): Promise<{ info: Uint8Array; pkg: Uint8Array }> {
  const { info, pkg } = streamsOf(await encryptedXlsx());
  return { info: infoOf(info, change(infoText(info))), pkg };
}

test('agile: an HMAC value that is not as long as a digest is damage, shorter or longer', async () => {
  for (const blocks of [3, 5]) {
    const changed = await withAttributes((text) => text.replace(/encryptedHmacValue="[^"]+"/, `encryptedHmacValue="${Buffer.alloc(blocks * 16, 7).toString('base64')}"`));
    assert.throws(() => parseAgileInfo(changed.info), (error: unknown) => error instanceof PasswordError && error.code === 'damaged', `${blocks} blocks`);
  }
});

test('agile: an HMAC that is right in its first 20 bytes and wrong after them is wrong', async () => {
  const plain = new Uint8Array(100).fill(9);
  const encryptor = await AgileEncryptor.create(utf16le(PASSWORD), { spinCount: 10 });
  const pkg = new Uint8Array(agilePackageSize(plain.length));
  await encryptor.encryptInto(plain, pkg);
  const descriptor = (encryptor as unknown as { descriptor(mac: Uint8Array): Uint8Array }).descriptor.bind(encryptor);
  const key = (encryptor as unknown as { hmacKey: Uint8Array }).hmacKey;
  const mac = hmac('SHA512', key, pkg);
  const good = parseAgileInfo(descriptor(mac));
  assert.equal((await agileDecryptPackage(good, await agileUnlock(good, utf16le(PASSWORD)), pkg)).length, 100);
  for (const at of [20, 40, 63]) {
    const bad = mac.slice();
    bad[at] ^= 1;
    const info = parseAgileInfo(descriptor(bad));
    await rejectsWith(agileDecryptPackage(info, await agileUnlock(info, utf16le(PASSWORD)), pkg), 'damaged', `byte ${at} of the HMAC`);
  }
});

test('agile: a package whose last block is not whole is damage, also when its HMAC was made to fit', async () => {
  const plain = new Uint8Array(20).fill(9);
  const encryptor = await AgileEncryptor.create(utf16le(PASSWORD), { spinCount: 10 });
  const pkg = new Uint8Array(agilePackageSize(plain.length));
  await encryptor.encryptInto(plain, pkg);
  assert.equal(pkg.length, 8 + 32);
  // 20 bytes of content need 32 of cipher: a stream that ends after 20, 24 or 31 does not hold them.
  for (const length of [8 + 20, 8 + 24, 8 + 31]) {
    const cut = pkg.slice(0, length);
    const info = parseAgileInfo(await encryptor.info(cut));
    await rejectsWith(agileDecryptPackage(info, await agileUnlock(info, utf16le(PASSWORD)), cut), 'damaged', `a stream of ${length} bytes`);
  }
  const info = parseAgileInfo(await encryptor.info(pkg));
  assert.equal((await agileDecryptPackage(info, await agileUnlock(info, utf16le(PASSWORD)), pkg)).length, 20);
});

for (const [packageBits, passwordBits] of [
  [256, 128],
  [128, 256],
  [192, 128],
]) {
  refTest(`agile: a package key of ${packageBits} bits that is encrypted with a password key of ${passwordBits} bits is opened`, async () => {
    const plain = new Uint8Array(5000).map((_, i) => (i * 3) & 0xff);
    const file = refOk('agile-encrypt', { password: passwordArg(PASSWORD), hash: 'SHA256', bits: packageBits, 'password-bits': passwordBits, spin: 100 }, plain);
    const cfb = CfbFile.read(file);
    const info = parseAgileInfo(cfb.stream('EncryptionInfo'));
    assert.deepEqual([info.keyData.keyBits, info.password.keyBits], [packageBits, passwordBits]);
    assert.equal(toHex((await decryptPackage(cfb, utf16le(PASSWORD))).bytes), toHex(plain));
  });
}

test('hmac: a key that is exactly one block long is used as it is, not hashed first', () => {
  // python3: hmac.new(bytes(range(n)), b'what do ya want for nothing?', h).hexdigest(), n the block size of h.
  const expected: Record<HashName, string> = {
    MD5: '1febc4e155fc69ff7ca35fcbed89172c',
    SHA1: '879dc82ca50b25b0f7bb9057993904f55678535b',
    SHA256: '5431cc41830bee7889a6b5d04b33877387ea9b8170759f4dca4323cfb5725508',
    SHA384: 'c9c29fca7348d7199b42f78b7556c31325552abfb330e72ed240145ba451a69254ef2a967136ecea77803262f2d00494',
    SHA512: '45a2353553c24eb6dc843fa22df01bec0a487ca3c7fe017d2d7bec8e7714686d2d9ab5a2817902eac0a6a50bcc8265f00308b8258c903c2ec7f7e4305d546cf4',
  };
  const message = new TextEncoder().encode('what do ya want for nothing?');
  for (const name of Object.keys(expected) as HashName[]) {
    const size = name === 'SHA384' || name === 'SHA512' ? 128 : 64;
    assert.equal(toHex(hmac(name, new Uint8Array(size).map((_, i) => i), message)), expected[name], name);
  }
});

test('session: what was derived for one number of rounds is not used for another', async () => {
  const plain = await minimalXlsx();
  const session = PasswordSession.of(PASSWORD);
  const first = await encryptPackage(plain, session, { spinCount: 10 });
  const second = await encryptPackage(plain, session, { spinCount: 20 });
  const third = await encryptPackage(plain, session, { spinCount: 20 });
  for (const [file, rounds] of [
    [first, 10],
    [second, 20],
    [third, 20],
  ] as [Uint8Array, number][]) {
    const opened = await decryptPackage(CfbFile.read(file), utf16le(PASSWORD));
    assert.equal(opened.encryption.spinCount, rounds);
    assert.equal(toHex(opened.bytes), toHex(plain), `the file of ${rounds} rounds opens`);
  }
});
