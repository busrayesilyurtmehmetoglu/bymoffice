// ECMA-376 document encryption (MS-OFFCRYPTO): the agile and the standard scheme, key derivation
// against known answers, and both directions against the reference implementation of
// test/tools/offcrypto_ref.py.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { CfbFile } from '../src/cfb';
import { AgileEncryptor, agileDecryptPackage, agileKey, agilePackageSize, agileUnlock, parseAgileInfo } from '../src/crypto/agile';
import { fromHex, toHex, utf16le } from '../src/crypto/bytes';
import { PasswordError } from '../src/crypto/errors';
import type { HashName } from '../src/crypto/hash';
import { parseStandardInfo, standardDecryptPackage, standardKey, standardUnlock } from '../src/crypto/standard';
import { MAX_SPIN_COUNT, spinHash } from '../src/crypto/work';
import { minimalXlsx } from './fixtures';
import { passwordArg, ref, refOk, refTest } from './reftools';

const PASSWORD = 'Password1234_';
const SALT = fromHex('000102030405060708090a0b0c0d0e0f');

/** Rejects with a PasswordError of the code; gives the error back. */
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

// offcrypto_ref.py agile-key --hash H --salt 00..0f --spin 1000 --block fea7d2763b4b9e79 --bits 256, password "Password1234_".
const AGILE_KEYS: Record<string, string> = {
  SHA1: 'b89b55e9c16e251ae089deba9a60a146daa8c7cc363636363636363636363636',
  SHA256: '65b07f3a871cca083417422f068564ab815bfd5c604b8b61d082e41118d9bc25',
  SHA384: '4f24ee86843f0e41b478b29726a215359f5e3dc8bc2dbc47bbbce7441a04bc59',
  SHA512: '17a33d5d69308ad043c371cdfd7f2887e060da16f9be90c4bdee43f2d512eb2e',
};

test('agile: the key derived from a password, a salt and a block key is the known one, for every digest and key size', async () => {
  const block = fromHex('fea7d2763b4b9e79');
  for (const hash of Object.keys(AGILE_KEYS) as HashName[]) {
    const base = await spinHash(hash, SALT, utf16le(PASSWORD), 1000);
    for (const bits of [128, 192, 256]) {
      // A SHA-1 digest is 20 bytes: a longer key is filled up with 0x36.
      assert.equal(toHex(agileKey(hash, base, block, bits / 8)), AGILE_KEYS[hash].slice(0, bits / 4), `${hash}, ${bits} bits`);
    }
  }
});

test('agile: a hundred thousand rounds of SHA-512 give the known key', async () => {
  const base = await spinHash('SHA512', SALT, utf16le(PASSWORD), 100000);
  assert.equal(toHex(agileKey('SHA512', base, fromHex('146e0be7abacd0d6'), 32)), '7a8b2091cd76dd40577bbc7b165de0985a9de0e0aded58ce94fc4b35294c0d0e');
});

test('standard: the key derived from a password and a salt is the known one, for every key size', async () => {
  // offcrypto_ref.py standard-key --salt 00..0f --bits N
  const expected = '340fc93a8e2154db8304d055df81cebcef5e89fb8e5eb520efdca76f3330e032';
  for (const bits of [128, 192, 256]) assert.equal(toHex(await standardKey(SALT, utf16le(PASSWORD), bits / 8)), expected.slice(0, bits / 4), `${bits} bits`);
});

test('key derivation: more rounds than Office allows are refused at once, and so are counts that are no counts', async () => {
  const started = Date.now();
  for (const count of [MAX_SPIN_COUNT + 1, 4294967295, -1, 1.5, NaN, Infinity]) await rejectsWith(spinHash('SHA512', SALT, utf16le('x'), count), 'unsupported', `spin count ${count}`);
  assert.ok(Date.now() - started < 2000, 'refusing took no time to speak of');
});

test('key derivation: the work is done in slices with pauses between them, and gives the same result', async () => {
  let pauses = 0;
  let clock = 0;
  // A clock that jumps by 5 ms each time it is read: every third look at it ends a slice of 12 ms.
  const sliced = await spinHash('SHA256', SALT, utf16le(PASSWORD), 20000, { now: () => (clock += 5), pause: async () => void pauses++, sliceMs: 12 });
  const atOnce = await spinHash('SHA256', SALT, utf16le(PASSWORD), 20000, { now: () => 0, pause: async () => assert.fail('no pause without time passing') });
  assert.equal(toHex(sliced), toHex(atOnce));
  assert.ok(pauses >= 20 && pauses <= 80, `paused ${pauses} times in 20 000 rounds`);
});

test('key derivation: it can be called off, before it starts and while it runs', async () => {
  await rejectsWith(spinHash('SHA512', SALT, utf16le(PASSWORD), 100000, { signal: { aborted: true } }), 'cancelled', 'called off before the start');
  const signal = { aborted: false };
  let pauses = 0;
  const started = Date.now();
  const work = spinHash('SHA512', SALT, utf16le(PASSWORD), MAX_SPIN_COUNT, {
    signal,
    sliceMs: 1,
    pause: async () => {
      if (++pauses === 3) signal.aborted = true;
    },
  });
  await rejectsWith(work, 'cancelled', 'called off at the third pause');
  // Ten million rounds would take the better part of a minute.
  assert.ok(Date.now() - started < 5000, `stopped after ${Date.now() - started} ms`);
});

test('key derivation: the page gets its turn while a hundred thousand rounds run', async () => {
  let ticks = 0;
  const timer = setInterval(() => ticks++, 5);
  try {
    const started = Date.now();
    await spinHash('SHA512', SALT, utf16le(PASSWORD), 100000);
    const elapsed = Date.now() - started;
    // A generous bound: this guards against a derivation that got stuck, it does not measure speed.
    assert.ok(elapsed < 30000, `a hundred thousand rounds took ${elapsed} ms`);
    assert.ok(elapsed < 60 || ticks >= 1, `a timer of the page fired ${ticks} times in ${elapsed} ms`);
  } finally {
    clearInterval(timer);
  }
});

// ------------------------------------------------------------------ reading what the reference wrote

interface Streams {
  info: Uint8Array;
  pkg: Uint8Array;
}

const streamsOf = (file: Uint8Array): Streams => {
  const cfb = CfbFile.read(file);
  return { info: cfb.stream('EncryptionInfo'), pkg: cfb.stream('EncryptedPackage') };
};

async function openAgile(streams: Streams, password: string): Promise<Uint8Array> {
  const info = parseAgileInfo(streams.info);
  const key = await agileUnlock(info, utf16le(password));
  return agileDecryptPackage(info, key, streams.pkg);
}

async function openStandard(streams: Streams, password: string): Promise<Uint8Array> {
  const info = parseStandardInfo(streams.info);
  const key = await standardUnlock(info, utf16le(password));
  return standardDecryptPackage(info, key, streams.pkg);
}

/** A payload that is not a multiple of anything: 3 segments and a bit. */
const payload = (size: number): Uint8Array => new Uint8Array(size).map((_, i) => (i * 31 + (i >> 8)) & 0xff);

for (const hash of ['SHA1', 'SHA256', 'SHA384', 'SHA512']) {
  for (const bits of [128, 192, 256]) {
    refTest(`agile: a file the reference encrypted with ${hash} and AES-${bits} is opened`, async () => {
      const plain = payload(4096 * 2 + 1234);
      const file = refOk('agile-encrypt', { password: passwordArg(PASSWORD), hash, bits, spin: 1000, seed: `${hash}-${bits}` }, plain);
      const streams = streamsOf(file);
      const info = parseAgileInfo(streams.info);
      assert.deepEqual([info.password.hash, info.password.keyBits, info.password.spinCount, info.keyData.hash, info.keyData.keyBits], [hash, bits, 1000, hash, bits]);
      assert.equal(toHex(await openAgile(streams, PASSWORD)), toHex(plain));
      await rejectsWith(openAgile(streams, 'Password1234'), 'wrong', 'a password that is one character short');
      await rejectsWith(openAgile(streams, ''), 'wrong', 'the empty password');
    });
  }
}

for (const bits of [128, 192, 256]) {
  refTest(`standard: a file the reference encrypted with AES-${bits} is opened`, async () => {
    const plain = payload(5000);
    const streams = streamsOf(refOk('standard-encrypt', { password: passwordArg(PASSWORD), bits, seed: `standard-${bits}` }, plain));
    const info = parseStandardInfo(streams.info);
    assert.deepEqual([info.keyBits, info.cipher, info.hash], [bits, 'AES', 'SHA1']);
    assert.equal(toHex(await openStandard(streams, PASSWORD)), toHex(plain));
    await rejectsWith(openStandard(streams, 'password1234_'), 'wrong', 'a password in another case');
  });
}

refTest('agile: passwords that are empty, Turkish, outside the basic plane or longer than 255 characters open their files', async () => {
  const plain = payload(100);
  // UNSURE: Office does not let a password of more than 255 characters be typed; a file that has one is opened all the same.
  for (const password of ['', 'ŞifreİıĞğÜüÖöÇç', 'p\u{1F512}\u{10FFFF}w', 'x'.repeat(300)]) {
    const streams = streamsOf(refOk('agile-encrypt', { password: passwordArg(password), spin: 50, seed: 'unicode' }, plain));
    assert.equal(toHex(await openAgile(streams, password)), toHex(plain), `password of ${password.length} units`);
    await rejectsWith(openAgile(streams, `${password} `), 'wrong', 'the password with a space after it');
  }
});

refTest('agile: a file without an integrity check is decrypted and says it has none; with requireIntegrity it is damage', async () => {
  // The element is optional in the schema: there is nothing to check, and nothing is claimed.
  const plain = payload(3000);
  const streams = streamsOf(refOk('agile-encrypt', { password: passwordArg(PASSWORD), spin: 100, seed: 'no-hmac' }, plain));
  const text = new TextDecoder().decode(streams.info.subarray(8));
  assert.match(text, /<dataIntegrity [^>]*\/>/);
  const body = new TextEncoder().encode(text.replace(/<dataIntegrity [^>]*\/>/, ''));
  const info = new Uint8Array(8 + body.length);
  info.set(streams.info.subarray(0, 8));
  info.set(body, 8);
  assert.equal(parseAgileInfo(info).integrity, undefined);
  assert.equal(toHex(await openAgile({ info, pkg: streams.pkg }, PASSWORD)), toHex(plain));
  // Nothing protects such a file: a bit that was changed comes out as other content, and nobody notices here. Whoever
  // opens the file is told by open() (`warnings`), and can ask for such files to be refused.
  const parsed = parseAgileInfo(info);
  const key = await agileUnlock(parsed, utf16le(PASSWORD));
  const changed = streams.pkg.slice();
  changed[100] ^= 1;
  assert.notEqual(toHex(await agileDecryptPackage(parsed, key, changed)), toHex(plain), 'changed content is read as it is');
  await rejectsWith(agileDecryptPackage(parsed, key, streams.pkg, undefined, { requireIntegrity: true }), 'damaged', 'an integrity check is required');
});

// ------------------------------------------------------------------ writing what the reference reads

/** An encrypted package as two streams, the way AgileEncryptor is used. */
async function sealed(plain: Uint8Array, password: string, options: { spinCount?: number } = {}): Promise<Streams> {
  const encryptor = await AgileEncryptor.create(utf16le(password), options);
  const pkg = new Uint8Array(agilePackageSize(plain.length));
  await encryptor.encryptInto(plain, pkg);
  const info = await encryptor.info(pkg);
  assert.equal(info.length, encryptor.infoSize, 'the size of the EncryptionInfo stream is known before it is written');
  encryptor.wipe();
  return { info, pkg };
}

const refFile = (streams: Streams): Uint8Array => refOk('cfb-write', {}, new TextEncoder().encode(JSON.stringify({ EncryptionInfo: toHex(streams.info), EncryptedPackage: toHex(streams.pkg) })));

refTest('agile: what the engine encrypts, the reference decrypts: the payload is the same, the HMAC holds, the ZIP inside is sound', async () => {
  const xlsx = await minimalXlsx();
  const streams = await sealed(xlsx, PASSWORD);
  const info = parseAgileInfo(streams.info);
  assert.deepEqual(
    [info.password.hash, info.password.keyBits, info.password.spinCount, info.password.chaining, info.keyData.hash, info.keyData.keyBits, info.keyData.salt.length],
    ['SHA512', 256, 100000, 'CBC', 'SHA512', 256, 16],
    'what a file is encrypted with by default',
  );
  const done = ref('decrypt', { password: passwordArg(PASSWORD), zip: 'check' }, refFile(streams));
  assert.equal(done.status, 0, done.err);
  assert.match(done.err, /kind=agile hmac=True/);
  assert.equal(toHex(done.out), toHex(xlsx));
  assert.equal(ref('decrypt', { password: passwordArg('Password1234-') }, refFile(streams)).status, 2, 'the reference refuses a wrong password');
});

refTest('agile: payloads at the edges of a segment and of a block come back whole from the reference', async () => {
  for (const size of [0, 1, 15, 16, 17, 4095, 4096, 4097, 8192]) {
    const plain = payload(size);
    const streams = await sealed(plain, 'ş\u{1F511}', { spinCount: 10 });
    assert.equal(streams.pkg.length, 8 + Math.ceil(size / 16) * 16, `the package of ${size} bytes`);
    const done = ref('decrypt', { password: passwordArg('ş\u{1F511}') }, refFile(streams));
    assert.equal(done.status, 0, `${size} bytes: ${done.err}`);
    assert.equal(toHex(done.out), toHex(plain), `${size} bytes`);
    assert.equal(toHex(await openAgile(streams, 'ş\u{1F511}')), toHex(plain), `${size} bytes, read back by the engine`);
  }
});

refTest('agile: the reference notices a package the engine wrote and someone changed', async () => {
  const streams = await sealed(payload(5000), PASSWORD, { spinCount: 10 });
  streams.pkg[streams.pkg.length - 1] ^= 1;
  assert.equal(ref('decrypt', { password: passwordArg(PASSWORD) }, refFile(streams)).status, 3);
});

test('agile: two encryptions of the same file with the same password share no salt, no key and no cipher text', async () => {
  const plain = payload(64);
  const a = await sealed(plain, PASSWORD, { spinCount: 10 });
  const b = await sealed(plain, PASSWORD, { spinCount: 10 });
  const ia = parseAgileInfo(a.info);
  const ib = parseAgileInfo(b.info);
  assert.notEqual(toHex(ia.password.salt), toHex(ib.password.salt), 'the salt of the password');
  assert.notEqual(toHex(ia.keyData.salt), toHex(ib.keyData.salt), 'the salt of the data');
  assert.notEqual(toHex(ia.password.encryptedKeyValue), toHex(ib.password.encryptedKeyValue), 'the key');
  assert.notEqual(toHex(a.pkg), toHex(b.pkg), 'the cipher text');
});

test('agile: without a secure source of random numbers nothing is encrypted', async () => {
  await assert.rejects(AgileEncryptor.create(utf16le(PASSWORD), { spinCount: 10, random: {} }), /no secure source of random numbers/);
});

test('agile: a password of any length is encrypted with (the limit is on a password that is set: see MAX_PASSWORD_CHARS)', async () => {
  const long = 'x'.repeat(300);
  const streams = await sealed(payload(40), long, { spinCount: 10 });
  assert.equal(toHex(await openAgile(streams, long)), toHex(payload(40)));
});

// ------------------------------------------------------------------ damage and hostility

test('agile: a package that was changed, cut short or lies about its size fails its check and is not opened', async () => {
  const plain = payload(9000);
  const whole = await sealed(plain, PASSWORD, { spinCount: 10 });
  assert.equal(toHex(await openAgile(whole, PASSWORD)), toHex(plain));

  for (const at of [0, 7, 8, 5000, whole.pkg.length - 1]) {
    const pkg = whole.pkg.slice();
    pkg[at] ^= 0x80;
    const error = await rejectsWith(openAgile({ info: whole.info, pkg }, PASSWORD), 'damaged', `byte ${at} changed`);
    // A change in the size field may make the size impossible, which is noticed first.
    if (at >= 8) assert.match(error.message, /integrity check/);
  }
  await rejectsWith(openAgile({ info: whole.info, pkg: whole.pkg.subarray(0, whole.pkg.length - 16) }, PASSWORD), 'damaged', 'the last block missing');
  await rejectsWith(openAgile({ info: whole.info, pkg: whole.pkg.subarray(0, 4) }, PASSWORD), 'damaged', 'four bytes of package');
  await rejectsWith(openAgile({ info: whole.info, pkg: new Uint8Array(0) }, PASSWORD), 'damaged', 'no package at all');
});

test('agile: a package whose size field says more than there is is refused even when its HMAC was made to fit', async () => {
  // Someone who knows the password can make the HMAC fit anything: the size is still checked against the stream.
  const encryptor = await AgileEncryptor.create(utf16le(PASSWORD), { spinCount: 10 });
  const pkg = new Uint8Array(agilePackageSize(100));
  await encryptor.encryptInto(payload(100), pkg);
  new DataView(pkg.buffer).setUint32(0, 1 << 30, true);
  const lying = { info: await encryptor.info(pkg), pkg };
  await rejectsWith(openAgile(lying, PASSWORD), 'damaged', 'a size of a gigabyte in 120 bytes');
  new DataView(pkg.buffer).setUint32(0, 100, true);
  new DataView(pkg.buffer).setUint32(4, 0x7fffffff, true);
  await rejectsWith(openAgile({ info: await encryptor.info(pkg), pkg }, PASSWORD), 'damaged', 'a size beyond 2^53');
});

const AGILE_XML = (attrs: Record<string, string | number> = {}, extra = ''): Uint8Array => {
  const a: Record<string, string | number> = {
    spinCount: 100000,
    saltSize: 16,
    blockSize: 16,
    keyBits: 256,
    hashSize: 64,
    cipherAlgorithm: 'AES',
    cipherChaining: 'ChainingModeCBC',
    hashAlgorithm: 'SHA512',
    saltValue: 'AAAAAAAAAAAAAAAAAAAAAA==',
    encryptedVerifierHashInput: 'AAAAAAAAAAAAAAAAAAAAAA==',
    encryptedVerifierHashValue: `${'A'.repeat(84)}AA==`,
    encryptedKeyValue: `${'A'.repeat(42)}A=`,
    ...attrs,
  };
  const password = Object.keys(a)
    .map((name) => `${name}="${a[name]}"`)
    .join(' ');
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${extra}<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">` +
    `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="AAAAAAAAAAAAAAAAAAAAAA=="/>` +
    `<dataIntegrity encryptedHmacKey="${'A'.repeat(84)}AA==" encryptedHmacValue="${'A'.repeat(84)}AA=="/>` +
    `<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey ${password}/></keyEncryptor></keyEncryptors></encryption>`;
  const body = new TextEncoder().encode(xml);
  const out = new Uint8Array(8 + body.length);
  out.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
  out.set(body, 8);
  return out;
};

const parseFails = (stream: Uint8Array, code: string, label: string): PasswordError => {
  try {
    parseAgileInfo(stream);
  } catch (error) {
    assert.ok(error instanceof PasswordError, `${label}: a PasswordError is thrown, got ${String(error)}`);
    assert.equal(error.code, code, `${label}: ${error.message}`);
    return error;
  }
  throw new assert.AssertionError({ message: `${label}: nothing was thrown` });
};

test('agile: the descriptor the tests build is read', () => {
  const info = parseAgileInfo(AGILE_XML());
  assert.deepEqual([info.password.spinCount, info.password.hash, info.password.keyBits, info.integrity !== undefined], [100000, 'SHA512', 256, true]);
});

test('agile: a descriptor that asks for too much, or for what is not read, is refused before any work is done', async () => {
  parseFails(AGILE_XML({ spinCount: 10000001 }), 'unsupported', 'a spin count over ten million');
  parseFails(AGILE_XML({ spinCount: 99999999999 }), 'unsupported', 'a huge spin count');
  parseFails(AGILE_XML({ spinCount: '-1' }), 'damaged', 'a negative spin count');
  parseFails(AGILE_XML({ spinCount: '1e5' }), 'damaged', 'a spin count that is no integer');
  parseFails(AGILE_XML({ cipherChaining: 'ChainingModeCFB' }), 'unsupported', 'CFB chaining');
  parseFails(AGILE_XML({ cipherAlgorithm: '3DES' }), 'unsupported', 'triple DES');
  parseFails(AGILE_XML({ hashAlgorithm: 'WHIRLPOOL' }), 'unsupported', 'another digest');
  parseFails(AGILE_XML({ hashAlgorithm: 'MD5', hashSize: 16 }), 'unsupported', 'MD5');
  parseFails(AGILE_XML({ keyBits: 512 }), 'unsupported', 'a key of 512 bits');
  parseFails(AGILE_XML({ blockSize: 32 }), 'unsupported', 'a block of 32 bytes');
  parseFails(AGILE_XML({ hashSize: 20 }), 'damaged', 'a digest size that is not the digest\'s');
  parseFails(AGILE_XML({ saltSize: 32 }), 'damaged', 'a salt that is not as long as it says');
  parseFails(AGILE_XML({ saltSize: 0, saltValue: '' }), 'damaged', 'no salt');
  parseFails(AGILE_XML({ saltValue: '!!!!' }), 'damaged', 'a salt that is not base64');
  parseFails(AGILE_XML({ encryptedKeyValue: 'AAAA' }), 'damaged', 'a key that is no whole block');
});

test('agile: a descriptor that is not what it should be is damaged, and a DTD is never read', () => {
  parseFails(new Uint8Array(4), 'damaged', 'four bytes');
  parseFails(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0, 0x3c]), 'damaged', 'XML that ends at once');
  const dtd = parseFails(AGILE_XML({}, '<!DOCTYPE encryption [<!ENTITY a "aaaaaaaaaa">]>'), 'damaged', 'a document type');
  assert.match(dtd.message, /DTD|document type/i);
  const noPassword = AGILE_XML().slice();
  const text = new TextDecoder().decode(noPassword.subarray(8)).replace(/<p:encryptedKey[^>]*\/>/, '');
  const stream = new Uint8Array(8 + text.length);
  stream.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
  stream.set(new TextEncoder().encode(text), 8);
  parseFails(stream, 'unsupported', 'a file that only a certificate opens');
  const big = new Uint8Array(8 + (1 << 21)).fill(0x20);
  big.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
  parseFails(big, 'damaged', 'two megabytes of descriptor');
});

test('standard: headers that are cut short, of another cipher or with sizes that cannot be are refused', () => {
  const fails = (stream: Uint8Array, code: string, label: string): void => {
    try {
      parseStandardInfo(stream);
    } catch (error) {
      assert.ok(error instanceof PasswordError, `${label}: a PasswordError is thrown, got ${String(error)}`);
      assert.equal(error.code, code, `${label}: ${error.message}`);
      return;
    }
    throw new assert.AssertionError({ message: `${label}: nothing was thrown` });
  };
  const header = (alg: number, hash: number, bits: number, saltSize = 16, headerSize = 32): Uint8Array => {
    const out = new Uint8Array(12 + 32 + 4 + 16 + 16 + 4 + 32);
    const view = new DataView(out.buffer);
    view.setUint16(0, 4, true);
    view.setUint16(2, 2, true);
    view.setUint32(4, 0x24, true);
    view.setUint32(8, headerSize, true);
    view.setUint32(12, 0x24, true);
    view.setUint32(20, alg, true);
    view.setUint32(24, hash, true);
    view.setUint32(28, bits, true);
    view.setUint32(44, saltSize, true);
    view.setUint32(80, 20, true);
    return out;
  };
  assert.equal(parseStandardInfo(header(0x660e, 0x8004, 128)).keyBits, 128);
  fails(header(0x6801, 0x8004, 128), 'unsupported', 'RC4 in a package');
  fails(header(0x660e, 0x8003, 128), 'unsupported', 'MD5');
  fails(header(0x660e, 0x8004, 192), 'damaged', 'AES-128 with a key of 192 bits');
  fails(header(0x660e, 0x8004, 128, 4096), 'damaged', 'a salt of 4096 bytes');
  fails(header(0x660e, 0x8004, 128, 16, 0x7fffffff), 'damaged', 'a header longer than the stream');
  fails(header(0x660e, 0x8004, 128).subarray(0, 60), 'damaged', 'a stream cut short');
  fails(new Uint8Array(3), 'damaged', 'three bytes');
  const extensible = header(0x660e, 0x8004, 128);
  extensible[2] = 3;
  fails(extensible, 'unsupported', 'the extensible scheme (version 4.3)');
});

test('standard: a package that is cut short or lies about its size is damaged', async () => {
  const info = parseStandardInfo(
    (() => {
      const out = new Uint8Array(12 + 32 + 4 + 16 + 16 + 4 + 32);
      const view = new DataView(out.buffer);
      view.setUint16(0, 4, true);
      view.setUint16(2, 2, true);
      view.setUint32(4, 0x24, true);
      view.setUint32(8, 32, true);
      view.setUint32(12, 0x24, true);
      view.setUint32(20, 0x660e, true);
      view.setUint32(24, 0x8004, true);
      view.setUint32(28, 128, true);
      view.setUint32(44, 16, true);
      view.setUint32(80, 20, true);
      return out;
    })(),
  );
  const key = new Uint8Array(16);
  const pkg = new Uint8Array(8 + 32);
  new DataView(pkg.buffer).setUint32(0, 33, true);
  await rejectsWith(standardDecryptPackage(info, key, pkg), 'damaged', '33 bytes in two blocks');
  await rejectsWith(standardDecryptPackage(info, key, pkg.subarray(0, 6)), 'damaged', 'six bytes of package');
  new DataView(pkg.buffer).setUint32(0, 20, true);
  assert.equal((await standardDecryptPackage(info, key, pkg)).length, 20);
});

test('encryption: no error of a wrong password or of a damaged file holds the password or a key', async () => {
  const password = 'S3cret-Parola-ĞÜŞ';
  const whole = await sealed(payload(200), password, { spinCount: 10 });
  const seen: string[] = [];
  const keep = async (work: Promise<unknown>): Promise<void> => {
    try {
      await work;
    } catch (error) {
      seen.push(`${String(error)} ${(error as Error).stack || ''} ${JSON.stringify(error)}`);
    }
  };
  await keep(openAgile(whole, `${password}x`));
  const broken = whole.pkg.slice();
  broken[20] ^= 1;
  await keep(openAgile({ info: whole.info, pkg: broken }, password));
  await keep(openAgile({ info: whole.info, pkg: whole.pkg.subarray(0, 9) }, password));
  assert.equal(seen.length, 3, 'three failures');
  const info = parseAgileInfo(whole.info);
  const key = await agileUnlock(info, utf16le(password));
  for (const text of seen) {
    assert.ok(!text.includes(password) && !text.includes(`${password}x`), `the password is not in: ${text.slice(0, 80)}`);
    assert.ok(!text.includes(toHex(key)) && !text.includes(toHex(utf16le(password))), `no key is in: ${text.slice(0, 80)}`);
  }
});
