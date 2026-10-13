// The cryptographic primitives the engine writes by hand, against known answers: the digests of
// FIPS 180 and RFC 1321, AES of FIPS 197, RC4 of RFC 6229, HMAC of RFC 2202/4231. The other
// expected values were produced once with Python's hashlib/hmac and the openssl command line and
// are constants here: nothing is called while the tests run.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { Aes, aesCbcDecrypt, aesCbcEncrypt, aesEcbDecrypt, aesEcbEncrypt } from '../src/crypto/aes';
import { constantTimeEqual, fromHex, randomBytes, toHex, utf16le, wipe } from '../src/crypto/bytes';
import { HASH_NAMES, createHash, digest, hmac, type HashName } from '../src/crypto/hash';
import { Rc4 } from '../src/crypto/rc4';

const ascii = (text: string): Uint8Array => new TextEncoder().encode(text);
const LONG = new Uint8Array(256 * 5).map((_, i) => i & 0xff);

const DIGESTS: Record<string, Record<HashName, string>> = {
  empty: {
    MD5: 'd41d8cd98f00b204e9800998ecf8427e',
    SHA1: 'da39a3ee5e6b4b0d3255bfef95601890afd80709',
    SHA256: 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855',
    SHA384: '38b060a751ac96384cd9327eb1b1e36a21fdb71114be07434c0cc7bf63f6e1da274edebfe76f65fbd51ad2f14898b95b',
    SHA512: 'cf83e1357eefb8bdf1542850d66d8007d620e4050b5715dc83f4a921d36ce9ce47d0d13c5d85f2b0ff8318d2877eec2f63b931bd47417a81a538327af927da3e',
  },
  abc: {
    MD5: '900150983cd24fb0d6963f7d28e17f72',
    SHA1: 'a9993e364706816aba3e25717850c26c9cd0d89d',
    SHA256: 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad',
    SHA384: 'cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7',
    SHA512: 'ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f',
  },
  two: {
    MD5: '8215ef0796a20bcaaae116d3876c664a',
    SHA1: '84983e441c3bd26ebaae4aa1f95129e5e54670f1',
    SHA256: '248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1',
    SHA384: '3391fdddfc8dc7393707a65b1b4709397cf8b1d162af05abfe8f450de5f36bc6b0455a8520bc4e6f5fe95b1fe3c8452b',
    SHA512: '204a8fc6dda82f0a0ced7beb8e08a41657c16ef468b228a8279be331a703c33596fd15c13b1b07f9aa1d3bea57789ca031ad85c7a71dd70354ec631238ca3445',
  },
  long: {
    MD5: '82829f1f3f2bb0f18b25f278e5bba8bd',
    SHA1: 'e37a04cb2353309f5cff4ee036cfb91a5e31cefd',
    SHA256: 'd414b085826eb06778483ba35564dc849e643359f69ed9747878ba6e54985bed',
    SHA384: 'f5943b1fe9e4b5bd3b45c83ef1e748f5e0f095f405a67d11a37016e5e908a4c5033bf084186d1e352db6fdda24dda58c',
    SHA512: 'c93f55ccf2fa8c82699ff9b58afe3591242b135d908a6d865e17e38adb41c21d1d5359e51273036373d54d20b5659cc87e6e7b381ff027d33f971416cc590f90',
  },
};
const MESSAGES: Record<string, Uint8Array> = {
  empty: new Uint8Array(0),
  abc: ascii('abc'),
  two: ascii('abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq'),
  long: LONG,
};

for (const name of HASH_NAMES) {
  test(`crypto: ${name} gives the known digests`, () => {
    for (const key of Object.keys(MESSAGES)) assert.equal(toHex(digest(name, MESSAGES[key])), DIGESTS[key][name], `${name} of "${key}"`);
  });

  test(`crypto: ${name} gives the same digest however the message is cut into updates, and after reset()`, () => {
    const hash = createHash(name);
    for (const step of [1, 3, 63, 64, 65, 127, 128, 129, 500]) {
      hash.reset();
      for (let at = 0; at < LONG.length; at += step) hash.update(LONG.subarray(at, Math.min(LONG.length, at + step)));
      assert.equal(toHex(hash.digest()), DIGESTS.long[name], `${name} in updates of ${step}`);
    }
  });

  test(`crypto: ${name} pads right at every length around its block size`, () => {
    // A message of n bytes hashed whole and hashed byte by byte give the same digest; lengths 55, 56, 63, 64, 111, 112, 119, 120, 127, 128 are where the padding wraps.
    for (let n = 0; n <= 260; n++) {
      const whole = toHex(digest(name, LONG.subarray(0, n)));
      const hash = createHash(name);
      for (let i = 0; i < n; i++) hash.update(LONG.subarray(i, i + 1));
      assert.equal(toHex(hash.digest()), whole, `${name} at ${n} bytes`);
    }
  });
}

// python3: the digest of the digests of the first n bytes of LONG, n = 0..260, laid end to end.
const EVERY_LENGTH: Record<HashName, string> = {
  MD5: '917aa9eb3c5503ede3c7fbf37749c51e',
  SHA1: '57d7fea532c9b15118a920d62b0f88a9b297f2fc',
  SHA256: 'b462003d0a60fe61db68b56c40eec21e993531dd139094e48ded151822bfadf9',
  SHA384: 'f9d6df1d872ef4e7e2d16e0757021afcd81361d3cacf074f9e07527e6e6e12b3917babb6ab7bd3e942bbcc97295a47f0',
  SHA512: '0956c25aa2cba477972641067ef9118e2d1faca127309d5283c7b954faf6d75e6ae9a6545d0c6a6c9b824b0ce9ccc1a37313c06ad19d7ff430b8c01b27fd9a05',
};

test('crypto: the digests of every length from 0 to 260 bytes are the ones hashlib gives', () => {
  for (const name of HASH_NAMES) {
    const all = createHash(name);
    for (let n = 0; n <= 260; n++) all.update(digest(name, LONG.subarray(0, n)));
    assert.equal(toHex(all.digest()), EVERY_LENGTH[name], name);
  }
});

const HMACS: Record<HashName, [short: string, long: string]> = {
  MD5: ['750c783e6ab0b503eaa86e310a5db738', 'f047deaf39e3cd3215a8669bca6f680f'],
  SHA1: ['effcdf6ae5eb2fa2d27416d5f184df9c259a7c79', '31b89321774b40a4c1358505fe11a2852ef457e4'],
  SHA256: ['5bdcc146bf60754e6a042426089575c75a003f089d2739839dec58b964ec3843', 'cc4b6725da8b7b5d1ed9c85fd1ecbbfe427fe67161c41edf20e3a86051b825f6'],
  SHA384: [
    'af45d2e376484031617f78d2b58a6b1b9c7ef464f5a01b47e42ec3736322445e8e2240ca5e69e2c78b3239ecfab21649',
    '83239e6fe735adcbbc8888ba7db42ccc4cdc7b0c1eb4d269185e78e742fba16c39e14c8ab1424be3844a8123a80bfff9',
  ],
  SHA512: [
    '164b7a7bfcf819e2e395fbe73b56e0a387bd64222e831fd610270cd7ea2505549758bf75c05a994a6d034f65f8f0e6fdcaeab1a34d4a6b4b636e070a38bce737',
    '6fc7235ebdbda4f25b0054df9a97c4062f0aec3de928dd2b3ed12d2f1670e349fd4baf8885b952bad99cdcf445bf7e521e76ef59bbf02f3ab0bda9b8c342196a',
  ],
};

test('crypto: HMAC gives the answers of RFC 2202 and RFC 4231 (test case 2), and with a key longer than a block', () => {
  for (const name of HASH_NAMES) {
    assert.equal(toHex(hmac(name, ascii('Jefe'), ascii('what do ya want for nothing?'))), HMACS[name][0], `HMAC-${name}, short key`);
    // The message in two parts: hmac() takes any number of them.
    assert.equal(toHex(hmac(name, new Uint8Array(200).fill(0xaa), LONG.subarray(0, 100), LONG.subarray(100, 256))), HMACS[name][1], `HMAC-${name}, key of 200 bytes`);
  }
});

const PLAIN = new Uint8Array(48).map((_, i) => i);
const IV = new Uint8Array(16).map((_, i) => 16 + i);
const AES: Record<number, { fips: string; ecb: string; cbc: string }> = {
  128: {
    fips: '69c4e0d86a7b0430d8cdb78070b4c55a',
    ecb: '0a940bb5416ef045f1c39458c653ea5a07feef74e1d5036e900eee118e9492935be87e2e5b447c944b21c9af7756c0d8',
    cbc: '954f64f2e4e86e9eee82d20216684899a93b9ddb22e8ab104c61e728831d6d5a1f51afc19343249ef3b4495be12dc4f3',
  },
  192: {
    fips: 'dda97ca4864cdfe06eaf70a0ec0d7191',
    ecb: '0060bffe46834bb8da5cf9a61ff220ae93ae3b7f9fc2e8159d05a6a9f5e24f2d779fe44f29fffa2c25d3464fc3db3c49',
    cbc: '3fe7286abde5f03943d5777020259626f7a43626455133f50664fa760ae36984d9e18b4ab13f579c71b4b43baf8bed76',
  },
  256: {
    fips: '8ea2b7ca516745bfeafc49904b496089',
    ecb: '5a6e045708fb7196f02e553d02c3a692e9c3ef8ab23453e6f0749cd636e7a88e61a6936e4e8f101c1cc1f993b542a0d4',
    cbc: '9f3b7504926f8bd36e3118e903a4cd4a25c183f70fdb4812cc2453fa00b3d390fbd621f919e2d9a0cb10aa5647bee3a7',
  },
};

for (const bits of [128, 192, 256]) {
  const key = new Uint8Array(bits / 8).map((_, i) => i);
  test(`crypto: AES-${bits} encrypts the block of FIPS 197 appendix C and decrypts it again`, () => {
    const block = fromHex('00112233445566778899aabbccddeeff');
    const aes = new Aes(key);
    const out = new Uint8Array(16);
    aes.encryptBlock(block, 0, out, 0);
    assert.equal(toHex(out), AES[bits].fips);
    const back = new Uint8Array(16);
    aes.decryptBlock(out, 0, back, 0);
    assert.equal(toHex(back), toHex(block));
  });

  test(`crypto: AES-${bits} in ECB and CBC gives what openssl gives, both ways`, () => {
    assert.equal(toHex(aesEcbEncrypt(key, PLAIN)), AES[bits].ecb);
    assert.equal(toHex(aesEcbDecrypt(key, fromHex(AES[bits].ecb))), toHex(PLAIN));
    assert.equal(toHex(aesCbcEncrypt(key, IV, PLAIN)), AES[bits].cbc);
    assert.equal(toHex(aesCbcDecrypt(key, IV, fromHex(AES[bits].cbc))), toHex(PLAIN));
  });
}

test('crypto: AES refuses keys, vectors and data of a size it cannot work with', () => {
  assert.throws(() => new Aes(new Uint8Array(15)), /AES key must be 16, 24 or 32 bytes/);
  assert.throws(() => new Aes(new Uint8Array(0)), /AES key must be 16, 24 or 32 bytes/);
  const key = new Uint8Array(16);
  assert.throws(() => aesEcbDecrypt(key, new Uint8Array(17)), /multiple of 16/);
  assert.throws(() => aesCbcEncrypt(key, new Uint8Array(16), new Uint8Array(5)), /multiple of 16/);
  assert.throws(() => aesCbcDecrypt(key, new Uint8Array(8), new Uint8Array(16)), /vector must be 16 bytes/);
  assert.equal(aesCbcDecrypt(key, new Uint8Array(16), new Uint8Array(0)).length, 0);
});

test('crypto: AES does not change its input', () => {
  const data = PLAIN.slice();
  aesCbcEncrypt(new Uint8Array(32), IV, data);
  aesCbcDecrypt(new Uint8Array(32), IV, data);
  assert.deepEqual(data, PLAIN);
});

test('crypto: RC4 gives the known key streams (RFC 6229) and cipher texts', () => {
  const run = (key: Uint8Array, data: Uint8Array): string => {
    const out = data.slice();
    new Rc4(key).apply(out);
    return toHex(out);
  };
  assert.equal(run(ascii('Key'), ascii('Plaintext')), 'bbf316e8d940af0ad3');
  assert.equal(run(ascii('Wiki'), ascii('pedia')), '1021bf0420');
  assert.equal(run(ascii('Secret'), ascii('Attack at dawn')), '45a01f645fc35b383552544b9bf5');
  assert.equal(run(new Uint8Array([1, 2, 3, 4, 5]), new Uint8Array(32)), 'b2396305f03dc027ccc3524a0a1118a86982944f18fc82d589c403a47a0d0919');
  assert.equal(run(new Uint8Array(16).map((_, i) => i + 1), new Uint8Array(32)), '9ac7cc9a609d1ef7b2932899cde41b975248c4959014126a6e8a84f11d1a9e1c');
});

test('crypto: RC4 goes on where it stopped, skips key stream, and works on a part of a buffer', () => {
  const key = new Uint8Array([1, 2, 3, 4, 5]);
  const whole = new Uint8Array(64);
  new Rc4(key).apply(whole);
  const parts = new Uint8Array(64);
  const rc4 = new Rc4(key);
  rc4.apply(parts, 0, 10);
  rc4.skip(20);
  rc4.apply(parts, 30, 64);
  assert.equal(toHex(parts.subarray(0, 10)), toHex(whole.subarray(0, 10)));
  assert.equal(toHex(parts.subarray(10, 30)), '00'.repeat(20), 'skipped bytes are not touched');
  assert.equal(toHex(parts.subarray(30)), toHex(whole.subarray(30)));
  assert.throws(() => new Rc4(new Uint8Array(0)), /RC4 key must be 1 to 256 bytes/);
});

test('crypto: constantTimeEqual compares content, and lengths', () => {
  assert.equal(constantTimeEqual(fromHex('0102ff'), fromHex('0102ff')), true);
  assert.equal(constantTimeEqual(fromHex('0102ff'), fromHex('0102fe')), false);
  assert.equal(constantTimeEqual(fromHex('0102ff'), fromHex('0102')), false);
  assert.equal(constantTimeEqual(new Uint8Array(0), new Uint8Array(0)), true);
});

test('crypto: utf16le writes code units little-endian, astral characters as two units', () => {
  assert.equal(toHex(utf16le('Aş')), '41005f01');
  assert.equal(toHex(utf16le('\u{1F600}')), '3dd800de');
  assert.equal(utf16le('').length, 0);
});

test('crypto: wipe zeroes buffers, and takes missing ones', () => {
  const a = fromHex('01020304');
  const b = new Uint32Array([1, 2]);
  wipe(a, undefined, b);
  assert.deepEqual([...a, ...b], [0, 0, 0, 0, 0, 0]);
});

test('crypto: randomBytes comes from the system and refuses to work without it', () => {
  const a = randomBytes(32);
  const b = randomBytes(32);
  assert.equal(a.length, 32);
  assert.notEqual(toHex(a), toHex(b));
  // More than getRandomValues() gives in one call.
  assert.equal(randomBytes(70000).length, 70000);
  assert.throws(() => randomBytes(16, {}), /no secure source of random numbers/);
  assert.throws(() => randomBytes(16, { crypto: {} }), /no secure source of random numbers/);
  // A source that fills nothing in is not believed either.
  assert.throws(() => randomBytes(32, { crypto: { getRandomValues: (array: Uint8Array) => array } }), /no secure source of random numbers/);
  assert.throws(() => randomBytes(8, { crypto: { getRandomValues: () => undefined } }), /no secure source of random numbers/);
});
