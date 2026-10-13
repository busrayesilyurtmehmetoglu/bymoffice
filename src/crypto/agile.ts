// ECMA-376 agile encryption (MS-OFFCRYPTO 2.3.4.10 and after): what Office
// 2010 and later write for a docx or xlsx with a password. The
// EncryptionInfo stream is version 4.4 and an XML descriptor; the package is
// encrypted in segments of 4096 bytes, each with its own initialisation
// vector, and an HMAC over the encrypted stream says whether it is whole.
//
// Reading takes AES (128, 192, 256) in CBC mode with SHA-1, SHA-256, SHA-384
// or SHA-512. Writing is always AES-256, SHA-512 and 100 000 rounds, which
// is what Office writes.

import { Aes } from './aes';
import { constantTimeEqual, fromBase64, randomBytes, toBase64, wipe, type RandomScope } from './bytes';
import { PasswordError, damagedEncryption, unsupportedEncryption } from './errors';
import { Hmac, createHash, digest, digestSizeOf, type HashName } from './hash';
import { MAX_SPIN_COUNT, Slicer, slicerOf, spinHash, type WorkOptions } from './work';
import { childElements, localName, parseXml, type XmlElement } from '../xml';

/** The attributes keyData and p:encryptedKey share. */
export interface AgileCipher {
  saltSize: number;
  blockSize: number;
  keyBits: number;
  hashSize: number;
  cipher: 'AES';
  chaining: 'CBC';
  hash: HashName;
  salt: Uint8Array;
}

export interface AgilePasswordKey extends AgileCipher {
  spinCount: number;
  encryptedVerifierHashInput: Uint8Array;
  encryptedVerifierHashValue: Uint8Array;
  encryptedKeyValue: Uint8Array;
}

export interface AgileInfo {
  /** How the package is encrypted. */
  keyData: AgileCipher;
  /** The HMAC of the encrypted package, and its key, both encrypted. A file may leave it out. */
  integrity?: { encryptedHmacKey: Uint8Array; encryptedHmacValue: Uint8Array };
  /** How the key of the package is encrypted with the password. */
  password: AgilePasswordKey;
}

const BLOCK_VERIFIER_INPUT = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
const BLOCK_VERIFIER_VALUE = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
const BLOCK_KEY_VALUE = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
const BLOCK_HMAC_KEY = new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]);
const BLOCK_HMAC_VALUE = new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]);

const SEGMENT = 4096;
/** A descriptor is a few hundred bytes; one with certificates a few thousand. */
const MAX_DESCRIPTOR = 1 << 20;
const PASSWORD_URI = 'http://schemas.microsoft.com/office/2006/keyEncryptor/password';

const HASHES: Record<string, HashName> = { SHA1: 'SHA1', 'SHA-1': 'SHA1', SHA256: 'SHA256', SHA384: 'SHA384', SHA512: 'SHA512' };

const roundUp = (size: number): number => Math.ceil(size / 16) * 16;

/** `size` bytes of `data`: cut, or filled up with 0x36 (MS-OFFCRYPTO 2.3.4.11). */
function fit(data: Uint8Array, size: number): Uint8Array {
  if (data.length >= size) return data.slice(0, size);
  const out = new Uint8Array(size).fill(0x36);
  out.set(data);
  return out;
}

/** The key for one purpose (`blockKey`) from the iterated digest of the password. */
export function agileKey(hash: HashName, base: Uint8Array, blockKey: Uint8Array, keyBytes: number): Uint8Array {
  const full = digest(hash, base, blockKey);
  const key = fit(full, keyBytes);
  wipe(full);
  return key;
}

function attribute(element: XmlElement, name: string): string {
  const value = element.attrs.get(name);
  if (value === undefined) throw damagedEncryption(`the encryption descriptor has no "${name}"`);
  return value;
}

function integer(element: XmlElement, name: string): number {
  const text = attribute(element, name);
  if (!/^[0-9]{1,15}$/.test(text)) throw damagedEncryption(`"${name}" of the encryption descriptor is not a number`);
  return Number(text);
}

function binary(element: XmlElement, name: string): Uint8Array {
  const text = attribute(element, name);
  const bytes = text.length <= MAX_DESCRIPTOR ? fromBase64(text) : undefined;
  if (!bytes) throw damagedEncryption(`"${name}" of the encryption descriptor is not base64`);
  return bytes;
}

function cipherOf(element: XmlElement): AgileCipher {
  const algorithm = attribute(element, 'cipherAlgorithm');
  if (algorithm !== 'AES') throw unsupportedEncryption(`the cipher ${JSON.stringify(algorithm.slice(0, 24))} (only AES is read)`);
  const chaining = attribute(element, 'cipherChaining');
  if (chaining !== 'ChainingModeCBC') throw unsupportedEncryption(`${JSON.stringify(chaining.slice(0, 24))} (only cipher block chaining, CBC, is read)`);
  const hashName = attribute(element, 'hashAlgorithm');
  const hash = Object.prototype.hasOwnProperty.call(HASHES, hashName) ? HASHES[hashName] : undefined;
  if (!hash) throw unsupportedEncryption(`the digest ${JSON.stringify(hashName.slice(0, 24))} (SHA-1, SHA-256, SHA-384 and SHA-512 are read)`);
  const keyBits = integer(element, 'keyBits');
  if (keyBits !== 128 && keyBits !== 192 && keyBits !== 256) throw unsupportedEncryption(`an AES key of ${keyBits} bits`);
  const blockSize = integer(element, 'blockSize');
  if (blockSize !== 16) throw unsupportedEncryption(`an AES block of ${blockSize} bytes`);
  const hashSize = integer(element, 'hashSize');
  if (hashSize !== digestSizeOf(hash)) throw damagedEncryption(`a ${hash} digest is not ${hashSize} bytes`);
  const saltSize = integer(element, 'saltSize');
  const salt = binary(element, 'saltValue');
  if (saltSize < 1 || saltSize > 65536 || salt.length !== saltSize) throw damagedEncryption('the salt is not as long as the descriptor says');
  return { saltSize, blockSize, keyBits, hashSize, cipher: 'AES', chaining: 'CBC', hash, salt };
}

/** Finds the first descendant with the local name; the descriptor is three levels deep. */
function find(parent: XmlElement, name: string, depth = 0): XmlElement | undefined {
  const children = childElements(parent);
  for (let i = 0; i < children.length; i++) {
    if (localName(children[i].name) === name) return children[i];
    const inner = depth < 4 ? find(children[i], name, depth + 1) : undefined;
    if (inner) return inner;
  }
  return undefined;
}

/** Reads an EncryptionInfo stream of version 4.4. Nothing is computed yet. */
export function parseAgileInfo(stream: Uint8Array): AgileInfo {
  if (stream.length < 9) throw damagedEncryption('the EncryptionInfo stream is cut short');
  if (stream[0] !== 4 || stream[1] !== 0 || stream[2] !== 4 || stream[3] !== 0) throw damagedEncryption('the EncryptionInfo stream is not of version 4.4');
  if (stream.length - 8 > MAX_DESCRIPTOR) throw damagedEncryption('the encryption descriptor is larger than any descriptor is');
  let root: XmlElement;
  try {
    // The parser of the packages: it refuses a DTD.
    root = parseXml(stream.subarray(8)).root;
  } catch (error) {
    throw damagedEncryption(`the encryption descriptor cannot be read (${error instanceof Error ? error.message.slice(0, 120) : 'not XML'})`);
  }
  if (localName(root.name) !== 'encryption') throw damagedEncryption('the encryption descriptor is not one');
  const keyDataElement = childElements(root, 'keyData')[0];
  if (!keyDataElement) throw damagedEncryption('the encryption descriptor has no "keyData"');
  const keyData = cipherOf(keyDataElement);

  let passwordElement: XmlElement | undefined;
  const encryptors = find(root, 'keyEncryptors');
  const list = encryptors ? childElements(encryptors, 'keyEncryptor') : [];
  for (let i = 0; i < list.length && i < 64 && !passwordElement; i++) {
    if (list[i].attrs.get('uri') === PASSWORD_URI) passwordElement = childElements(list[i], 'encryptedKey')[0];
  }
  if (!passwordElement) throw unsupportedEncryption('the file is not opened by a password (a certificate, perhaps)');
  const cipher = cipherOf(passwordElement);
  const spinText = attribute(passwordElement, 'spinCount');
  if (!/^[0-9]+$/.test(spinText)) throw damagedEncryption('"spinCount" of the encryption descriptor is not a number');
  const spinCount = spinText.length > 15 ? Infinity : Number(spinText);
  if (spinCount > MAX_SPIN_COUNT) throw new PasswordError('unsupported', `The file asks for more than ${MAX_SPIN_COUNT} rounds of key derivation: it is not opened`);
  const password: AgilePasswordKey = {
    ...cipher,
    spinCount,
    encryptedVerifierHashInput: binary(passwordElement, 'encryptedVerifierHashInput'),
    encryptedVerifierHashValue: binary(passwordElement, 'encryptedVerifierHashValue'),
    encryptedKeyValue: binary(passwordElement, 'encryptedKeyValue'),
  };
  if (password.encryptedVerifierHashInput.length !== roundUp(cipher.saltSize)) throw damagedEncryption('the password verifier is not as long as its salt');
  if (password.encryptedVerifierHashValue.length !== roundUp(cipher.hashSize)) throw damagedEncryption('the digest of the password verifier is not as long as a digest');
  if (password.encryptedKeyValue.length !== roundUp(keyData.keyBits / 8)) throw damagedEncryption('the encrypted key is not as long as a key');

  const info: AgileInfo = { keyData, password };
  const integrity = childElements(root, 'dataIntegrity')[0];
  if (integrity) {
    const encryptedHmacKey = binary(integrity, 'encryptedHmacKey');
    const encryptedHmacValue = binary(integrity, 'encryptedHmacValue');
    const size = roundUp(keyData.hashSize);
    // Some writers give the key one block more than a digest needs.
    if (encryptedHmacKey.length < size || encryptedHmacKey.length > size + 64 || encryptedHmacKey.length % 16 !== 0 || encryptedHmacValue.length !== size) {
      throw damagedEncryption('the integrity check of the descriptor is not as long as a digest');
    }
    info.integrity = { encryptedHmacKey, encryptedHmacValue };
  }
  return info;
}

function cbc(encrypt: boolean, key: Uint8Array, iv: Uint8Array, data: Uint8Array): Uint8Array {
  const aes = new Aes(key);
  const out = new Uint8Array(data.length);
  try {
    aes.cbc(encrypt, iv, data, out);
  } finally {
    aes.wipe();
  }
  return out;
}

/**
 * Checks the password and gives the key of the package. Rejects with
 * PasswordError "wrong" for a password that is not the file's. `password` is
 * the UTF-16LE form of the password; the key is the caller's to wipe.
 */
export async function agileUnlock(info: AgileInfo, password: Uint8Array, work?: WorkOptions | Slicer): Promise<Uint8Array> {
  const p = info.password;
  const base = await spinHash(p.hash, p.salt, password, p.spinCount, work);
  const iv = fit(p.salt, p.blockSize);
  const keyBytes = p.keyBits / 8;
  const inputKey = agileKey(p.hash, base, BLOCK_VERIFIER_INPUT, keyBytes);
  const valueKey = agileKey(p.hash, base, BLOCK_VERIFIER_VALUE, keyBytes);
  const secretKey = agileKey(p.hash, base, BLOCK_KEY_VALUE, keyBytes);
  const input = cbc(false, inputKey, iv, p.encryptedVerifierHashInput);
  const value = cbc(false, valueKey, iv, p.encryptedVerifierHashValue);
  const expected = digest(p.hash, input.subarray(0, p.saltSize));
  try {
    if (!constantTimeEqual(expected, value.subarray(0, p.hashSize))) throw new PasswordError('wrong', 'The password is not the password of this file');
    const secret = cbc(false, secretKey, iv, p.encryptedKeyValue);
    const key = secret.slice(0, info.keyData.keyBits / 8);
    wipe(secret);
    return key;
  } finally {
    wipe(base, inputKey, valueKey, secretKey, input, value, expected);
  }
}

/** The initialisation vector of a purpose: the digest of the data salt and a block key. */
function vector(keyData: AgileCipher, blockKey: Uint8Array): Uint8Array {
  return fit(digest(keyData.hash, keyData.salt, blockKey), keyData.blockSize);
}

/** The size the package stream says its content has; throws for a stream that cannot hold it. */
function packageSize(stream: Uint8Array): number {
  if (stream.length < 8) throw damagedEncryption('the encrypted package is cut short');
  const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
  const high = view.getUint32(4, true);
  const size = view.getUint32(0, true) + high * 4294967296;
  if (high > 0x1fffff || roundUp(size) > stream.length - 8) throw damagedEncryption('the encrypted package is shorter than it says it is');
  return size;
}

/** Bytes of a stream that are hashed between two turns of the page: some 4 ms of SHA-512. */
const HMAC_STEP = 1 << 18;

/** The HMAC of a whole stream, computed in slices: a package of tens of megabytes takes a second. */
async function streamHmac(hash: HashName, key: Uint8Array, stream: Uint8Array, slicer: Slicer): Promise<Uint8Array> {
  const mac = new Hmac(hash, key);
  try {
    for (let at = 0; at < stream.length; at += HMAC_STEP) {
      mac.update(stream.subarray(at, Math.min(stream.length, at + HMAC_STEP)));
      await slicer.turn();
    }
    return mac.digest();
  } catch (error) {
    mac.wipe();
    throw error;
  }
}

export interface AgileDecryptOptions {
  /** A file without an integrity check (no `dataIntegrity` in its descriptor) is refused as damaged instead of being read unchecked. */
  requireIntegrity?: boolean;
}

/**
 * Decrypts an EncryptedPackage stream with the key agileUnlock() gave. The
 * HMAC of the stream is checked first: a stream that fails it is not
 * decrypted, and the call rejects with PasswordError "damaged". A file
 * that has no HMAC is decrypted unchecked (`info.integrity` says so, and the
 * caller has to tell whoever opens the file) unless `requireIntegrity` is set.
 */
export async function agileDecryptPackage(
  info: AgileInfo,
  key: Uint8Array,
  stream: Uint8Array,
  work?: WorkOptions | Slicer,
  options: AgileDecryptOptions = {},
): Promise<Uint8Array> {
  const slicer = slicerOf(work);
  const keyData = info.keyData;
  const size = packageSize(stream);
  if (!info.integrity && options.requireIntegrity === true) throw damagedEncryption('it has no integrity check (its content could have been changed by someone who does not know the password)');
  if (info.integrity) {
    const hmacKeyFull = cbc(false, key, vector(keyData, BLOCK_HMAC_KEY), info.integrity.encryptedHmacKey);
    const expected = cbc(false, key, vector(keyData, BLOCK_HMAC_VALUE), info.integrity.encryptedHmacValue);
    await slicer.turn();
    let actual: Uint8Array;
    try {
      actual = await streamHmac(keyData.hash, hmacKeyFull.subarray(0, keyData.hashSize), stream, slicer);
    } catch (error) {
      wipe(hmacKeyFull, expected);
      throw error;
    }
    const same = constantTimeEqual(actual, expected.subarray(0, keyData.hashSize));
    wipe(hmacKeyFull, expected, actual);
    if (!same) throw damagedEncryption('it fails its integrity check (the file was changed or is incomplete)');
    await slicer.turn();
  }
  const out = new Uint8Array(roundUp(size));
  const aes = new Aes(key);
  const hash = createHash(keyData.hash);
  const counter = new Uint8Array(4);
  const full = new Uint8Array(hash.digestSize);
  try {
    for (let at = 0, segment = 0; at < out.length; at += SEGMENT, segment++) {
      const end = Math.min(out.length, at + SEGMENT);
      new DataView(counter.buffer).setUint32(0, segment, true);
      hash.update(keyData.salt).update(counter).digest(full);
      aes.cbc(false, fit(full, 16), stream.subarray(8 + at, 8 + end), out.subarray(at, end));
      if ((segment & 15) === 15) await slicer.turn();
    }
  } catch (error) {
    wipe(out);
    throw error;
  } finally {
    aes.wipe();
    hash.wipe();
  }
  // The padding of the last block is not part of the file.
  return out.subarray(0, size);
}

/** Bytes in the EncryptedPackage stream of a package of `plainSize` bytes. */
export function agilePackageSize(plainSize: number): number {
  return 8 + roundUp(plainSize);
}

export interface AgileEncryptOptions extends WorkOptions {
  /** Rounds of key derivation. Default 100 000, as Office. */
  spinCount?: number;
  /** Where random numbers come from (tests). Default: the system's. */
  random?: RandomScope;
  /**
   * The iterated digest of the password from an earlier encryption with the
   * same password, and its salt: saves the rounds when the same file is
   * saved again and again. Everything else is made anew.
   */
  derived?: { salt: Uint8Array; base: Uint8Array };
}

/**
 * The longest password Office takes, in characters. UNSURE: the limit of its dialog. It is the limit for a password
 * that is set (setPassword(), save({ password })); a file that came with a longer one is read, and saved with it again.
 */
export const MAX_PASSWORD_CHARS = 255;

const WRITE_HASH: HashName = 'SHA512';
const WRITE_KEY_BYTES = 32;
const WRITE_HASH_SIZE = 64;

/**
 * Encrypts one package: AES-256 in CBC mode, SHA-512, with new salts, a new
 * key and a new HMAC key from the random number generator of the system.
 */
export class AgileEncryptor {
  private constructor(
    private readonly spinCount: number,
    /** The salt of the password and the iterated digest it gave (see AgileEncryptOptions.derived); zeroed by wipe(). */
    readonly derived: { salt: Uint8Array; base: Uint8Array },
    private readonly dataSalt: Uint8Array,
    private readonly secret: Uint8Array,
    private readonly verifier: Uint8Array,
    private readonly hmacKey: Uint8Array,
  ) {}

  static async create(password: Uint8Array, options: AgileEncryptOptions = {}): Promise<AgileEncryptor> {
    const spinCount = options.spinCount === undefined ? 100000 : options.spinCount;
    // Before any work: nothing is derived for a file that cannot be encrypted.
    const random = (count: number): Uint8Array => (options.random ? randomBytes(count, options.random) : randomBytes(count));
    const dataSalt = random(16);
    const secret = random(WRITE_KEY_BYTES);
    const verifier = random(16);
    const hmacKey = random(WRITE_HASH_SIZE);
    // Copies of its own: whoever keeps the digest for the next save may zero theirs while this file is still being written.
    let derived: { salt: Uint8Array; base: Uint8Array } | undefined = options.derived ? { salt: options.derived.salt.slice(), base: options.derived.base.slice() } : undefined;
    if (!derived) {
      const salt = random(16);
      derived = { salt, base: await spinHash(WRITE_HASH, salt, password, spinCount, options) };
    }
    return new AgileEncryptor(spinCount, derived, dataSalt, secret, verifier, hmacKey);
  }

  /** Bytes in the EncryptionInfo stream this encryptor writes: known before anything is encrypted. */
  get infoSize(): number {
    return this.descriptor(new Uint8Array(WRITE_HASH_SIZE)).length;
  }

  /** Writes the EncryptedPackage stream of `plain` into `target`, which is agilePackageSize(plain.length) bytes. */
  async encryptInto(plain: Uint8Array, target: Uint8Array, work?: WorkOptions | Slicer): Promise<void> {
    if (target.length !== agilePackageSize(plain.length)) throw new Error('The target is not the size of the encrypted package');
    const slicer = slicerOf(work);
    const view = new DataView(target.buffer, target.byteOffset, target.byteLength);
    view.setUint32(0, plain.length >>> 0, true);
    view.setUint32(4, Math.floor(plain.length / 4294967296), true);
    const aes = new Aes(this.secret);
    const hash = createHash(WRITE_HASH);
    const counter = new Uint8Array(4);
    const full = new Uint8Array(hash.digestSize);
    const last = new Uint8Array(16);
    try {
      for (let at = 0, segment = 0; at < plain.length; at += SEGMENT, segment++) {
        const end = Math.min(plain.length, at + SEGMENT);
        new DataView(counter.buffer).setUint32(0, segment, true);
        hash.update(this.dataSalt).update(counter).digest(full);
        const whole = at + Math.floor((end - at) / 16) * 16;
        const iv = full.subarray(0, 16);
        aes.cbc(true, iv, plain.subarray(at, whole), target.subarray(8 + at, 8 + whole));
        if (whole < end) {
          // The last block of the file: filled up with zeros, and chained to the block before it.
          last.fill(0);
          last.set(plain.subarray(whole, end));
          aes.cbc(true, whole > at ? target.subarray(8 + whole - 16, 8 + whole) : iv, last, target.subarray(8 + whole, 8 + whole + 16));
        }
        if ((segment & 15) === 15) await slicer.turn();
      }
    } finally {
      aes.wipe();
      hash.wipe();
      wipe(last, full);
    }
  }

  /** The EncryptionInfo stream for the encrypted package that encryptInto() wrote: its HMAC is computed here, in slices. */
  async info(packageStream: Uint8Array, work?: WorkOptions | Slicer): Promise<Uint8Array> {
    const mac = await streamHmac(WRITE_HASH, this.hmacKey, packageStream, slicerOf(work));
    const out = this.descriptor(mac);
    wipe(mac);
    return out;
  }

  private descriptor(mac: Uint8Array): Uint8Array {
    const { salt, base } = this.derived;
    const keyData: AgileCipher = { saltSize: 16, blockSize: 16, keyBits: WRITE_KEY_BYTES * 8, hashSize: WRITE_HASH_SIZE, cipher: 'AES', chaining: 'CBC', hash: WRITE_HASH, salt: this.dataSalt };
    const sealed = (blockKey: Uint8Array, data: Uint8Array): string => {
      const key = agileKey(WRITE_HASH, base, blockKey, WRITE_KEY_BYTES);
      const padded = new Uint8Array(roundUp(data.length));
      padded.set(data);
      const out = cbc(true, key, salt, padded);
      wipe(key, padded);
      return toBase64(out);
    };
    const verifierHash = digest(WRITE_HASH, this.verifier);
    const common = `saltSize="16" blockSize="16" keyBits="${WRITE_KEY_BYTES * 8}" hashSize="${WRITE_HASH_SIZE}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="${WRITE_HASH}"`;
    const xml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
      `<keyData ${common} saltValue="${toBase64(this.dataSalt)}"/>` +
      `<dataIntegrity encryptedHmacKey="${toBase64(cbc(true, this.secret, vector(keyData, BLOCK_HMAC_KEY), this.hmacKey))}" ` +
      `encryptedHmacValue="${toBase64(cbc(true, this.secret, vector(keyData, BLOCK_HMAC_VALUE), mac))}"/>` +
      `<keyEncryptors><keyEncryptor uri="${PASSWORD_URI}">` +
      `<p:encryptedKey spinCount="${this.spinCount}" ${common} saltValue="${toBase64(salt)}" ` +
      `encryptedVerifierHashInput="${sealed(BLOCK_VERIFIER_INPUT, this.verifier)}" ` +
      `encryptedVerifierHashValue="${sealed(BLOCK_VERIFIER_VALUE, verifierHash)}" ` +
      `encryptedKeyValue="${sealed(BLOCK_KEY_VALUE, this.secret)}"/>` +
      '</keyEncryptor></keyEncryptors></encryption>';
    const body = new TextEncoder().encode(xml);
    const out = new Uint8Array(8 + body.length);
    // Version 4.4, and the flag that says "agile".
    out.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
    out.set(body, 8);
    return out;
  }

  /** Zeroes the keys of this file and the digest of the password: whoever wants `derived` for the next save copies it first. */
  wipe(): void {
    wipe(this.secret, this.verifier, this.hmacKey, this.derived.base);
  }
}
