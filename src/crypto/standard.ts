// ECMA-376 standard encryption (MS-OFFCRYPTO 2.3.4.5 and after): what Office
// 2007 writes for a docx or xlsx with a password. The EncryptionInfo stream
// is a binary header (version 2.2, 3.2 or 4.2) that names a CryptoAPI cipher;
// the package is AES in ECB mode with a key from 50 000 rounds of SHA-1.
// There is no integrity check. Read only: nothing is written this way.
//
// The same header, with RC4 as its cipher, is what password-protected xls
// and doc files of Office 2002 and later carry (see legacy.ts).

import { Aes } from './aes';
import { constantTimeEqual, le32, wipe } from './bytes';
import { PasswordError, damagedEncryption, unsupportedEncryption } from './errors';
import { digest } from './hash';
import { Slicer, slicerOf, spinHash, type WorkOptions } from './work';

export const ALG_RC4 = 0x6801;
export const ALG_AES128 = 0x660e;
export const ALG_AES192 = 0x660f;
export const ALG_AES256 = 0x6610;
const ALG_SHA1 = 0x8004;

const FLAG_AES = 0x20;
const FLAG_EXTERNAL = 0x10;

/** An EncryptionHeader and its EncryptionVerifier, as the file has them. */
export interface CryptoApiInfo {
  major: number;
  minor: number;
  flags: number;
  algId: number;
  algIdHash: number;
  /** As the header says; 0 stands for 40 with RC4. */
  keyBits: number;
  salt: Uint8Array;
  encryptedVerifier: Uint8Array;
  verifierHashSize: number;
  /** What follows the size of the verifier's digest, to the end of the structure. */
  encryptedVerifierHash: Uint8Array;
  /** Bytes the structure takes in the stream it was read from. */
  size: number;
}

/**
 * Reads the version, the EncryptionHeader and the EncryptionVerifier at the
 * start of `stream`. `hashBytes` is how long the encrypted digest of the
 * verifier is: 32 with AES (a SHA-1 digest filled up to two blocks), 20 with RC4.
 */
export function parseCryptoApiInfo(stream: Uint8Array, hashBytes: (algId: number) => number): CryptoApiInfo {
  if (stream.length < 12) throw damagedEncryption('the encryption header is cut short');
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const major = view.getUint16(0, true);
  const minor = view.getUint16(2, true);
  const flags = view.getUint32(4, true);
  const headerSize = view.getUint32(8, true);
  // Eight fields of four bytes, then the name of the provider.
  if (headerSize < 32 || headerSize > stream.length - 12) throw damagedEncryption('the encryption header is longer than its stream');
  const algId = view.getUint32(12 + 8, true);
  const algIdHash = view.getUint32(12 + 12, true);
  const keyBits = view.getUint32(12 + 16, true);
  let at = 12 + headerSize;
  if (at + 4 > stream.length) throw damagedEncryption('the encryption header has no verifier');
  const saltSize = view.getUint32(at, true);
  if (saltSize !== 16) throw damagedEncryption('the salt of the encryption header is not 16 bytes');
  at += 4;
  const hashSize = hashBytes(algId);
  if (at + 16 + 16 + 4 + hashSize > stream.length) throw damagedEncryption('the verifier of the encryption header is cut short');
  const salt = stream.slice(at, at + 16);
  const encryptedVerifier = stream.slice(at + 16, at + 32);
  const verifierHashSize = view.getUint32(at + 32, true);
  const encryptedVerifierHash = stream.slice(at + 36, at + 36 + hashSize);
  return { major, minor, flags, algId, algIdHash, keyBits, salt, encryptedVerifier, verifierHashSize, encryptedVerifierHash, size: at + 36 + hashSize };
}

export interface StandardInfo {
  cipher: 'AES';
  hash: 'SHA1';
  keyBits: 128 | 192 | 256;
  salt: Uint8Array;
  encryptedVerifier: Uint8Array;
  encryptedVerifierHash: Uint8Array;
}

/** Reads an EncryptionInfo stream of the standard scheme. Nothing is computed yet. */
export function parseStandardInfo(stream: Uint8Array): StandardInfo {
  if (stream.length < 8) throw damagedEncryption('the EncryptionInfo stream is cut short');
  const minor = stream[2] | (stream[3] << 8);
  if (minor === 3) throw unsupportedEncryption('the extensible scheme (a third-party encryption module)');
  if (minor !== 2) throw unsupportedEncryption(`an EncryptionInfo of version ${stream[0] | (stream[1] << 8)}.${minor}`);
  const raw = parseCryptoApiInfo(stream, () => 32);
  if (raw.flags & FLAG_EXTERNAL) throw unsupportedEncryption('the extensible scheme (a third-party encryption module)');
  // A header that names no cipher leaves it to the flags and the key size.
  const algId = raw.algId === 0 && raw.flags & FLAG_AES ? (raw.keyBits === 192 ? ALG_AES192 : raw.keyBits === 256 ? ALG_AES256 : ALG_AES128) : raw.algId;
  const keyBits = algId === ALG_AES128 ? 128 : algId === ALG_AES192 ? 192 : algId === ALG_AES256 ? 256 : 0;
  if (keyBits === 0) throw unsupportedEncryption(algId === ALG_RC4 ? 'RC4 in a docx or xlsx file' : `the cipher 0x${algId.toString(16)}`);
  if (raw.algIdHash !== 0 && raw.algIdHash !== ALG_SHA1) throw unsupportedEncryption(`the digest 0x${raw.algIdHash.toString(16)}`);
  if (raw.keyBits !== keyBits) throw damagedEncryption(`an AES-${keyBits} key is not ${raw.keyBits} bits`);
  if (raw.verifierHashSize !== 20) throw damagedEncryption('the digest of the password verifier is not as long as a SHA-1 digest');
  return { cipher: 'AES', hash: 'SHA1', keyBits, salt: raw.salt, encryptedVerifier: raw.encryptedVerifier, encryptedVerifierHash: raw.encryptedVerifierHash };
}

const STANDARD_SPIN = 50000;

/** The AES key of the standard scheme for a salt and a password (UTF-16LE), `keyBytes` long. */
export async function standardKey(salt: Uint8Array, password: Uint8Array, keyBytes: number, work?: WorkOptions | Slicer): Promise<Uint8Array> {
  const base = await spinHash('SHA1', salt, password, STANDARD_SPIN, work);
  const final = digest('SHA1', base, le32(0));
  // CryptDeriveKey: the digest is spread over two pads, and the key is cut from the digests of both.
  const inner = new Uint8Array(64).fill(0x36);
  const outer = new Uint8Array(64).fill(0x5c);
  for (let i = 0; i < final.length; i++) {
    inner[i] ^= final[i];
    outer[i] ^= final[i];
  }
  const x1 = digest('SHA1', inner);
  const x2 = digest('SHA1', outer);
  const key = new Uint8Array(keyBytes);
  key.set(x1.subarray(0, Math.min(20, keyBytes)));
  if (keyBytes > 20) key.set(x2.subarray(0, keyBytes - 20), 20);
  wipe(base, final, inner, outer, x1, x2);
  return key;
}

/** Checks the password and gives the key of the package; rejects with PasswordError "wrong". */
export async function standardUnlock(info: StandardInfo, password: Uint8Array, work?: WorkOptions | Slicer): Promise<Uint8Array> {
  const key = await standardKey(info.salt, password, info.keyBits / 8, work);
  const aes = new Aes(key);
  const verifier = new Uint8Array(16);
  const hash = new Uint8Array(32);
  aes.ecb(false, info.encryptedVerifier, verifier);
  aes.ecb(false, info.encryptedVerifierHash, hash);
  aes.wipe();
  const expected = digest('SHA1', verifier);
  const same = constantTimeEqual(expected, hash.subarray(0, 20));
  wipe(verifier, hash, expected);
  if (!same) {
    wipe(key);
    throw new PasswordError('wrong', 'The password is not the password of this file');
  }
  return key;
}

/** Decrypts an EncryptedPackage stream with the key standardUnlock() gave. */
export async function standardDecryptPackage(_info: StandardInfo, key: Uint8Array, stream: Uint8Array, work?: WorkOptions | Slicer): Promise<Uint8Array> {
  const slicer = slicerOf(work);
  if (stream.length < 8) throw damagedEncryption('the encrypted package is cut short');
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const high = view.getUint32(4, true);
  const size = view.getUint32(0, true) + high * 4294967296;
  const blocks = Math.ceil(size / 16) * 16;
  if (high > 0x1fffff || blocks > stream.length - 8) throw damagedEncryption('the encrypted package is shorter than it says it is');
  const out = new Uint8Array(blocks);
  const aes = new Aes(key);
  try {
    const step = 1 << 16;
    for (let at = 0; at < blocks; at += step) {
      const end = Math.min(blocks, at + step);
      aes.ecb(false, stream.subarray(8 + at, 8 + end), out.subarray(at, end));
      await slicer.turn();
    }
  } finally {
    aes.wipe();
  }
  return out.subarray(0, size);
}
