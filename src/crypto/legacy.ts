// The protection of doc and xls files (MS-OFFCRYPTO 2.3.5 to 2.3.7, MS-XLS
// 2.2.10, MS-DOC 2.2.6): XOR obfuscation (Office 95 and before), RC4 with a
// 40-bit key (Office 97 and 2000) and RC4 through the CryptoAPI (Office 2002
// and 2003). All three are weak and none is ever written: they are read so
// that old files open.
//
// In every scheme a byte keeps its place: the stream is as long encrypted as
// plain, some parts of it are left plain (record headers, the first bytes of
// a Word document) and the key stream runs on underneath them.

import { constantTimeEqual, le32, wipe } from './bytes';
import { PasswordError, damagedEncryption, unsupportedEncryption } from './errors';
import { createHash, digest } from './hash';
import type { FileEncryption } from './ooxml';
import { Rc4 } from './rc4';
import { ALG_RC4, parseCryptoApiInfo } from './standard';
import { decodeBytes } from '../formats/xls/codepage';

const wrong = (): PasswordError => new PasswordError('wrong', 'The password is not the password of this file');

// ----------------------------------------------------------------- XOR

/** XOR obfuscation takes a password of at most 15 single-byte characters. */
const XOR_MAX = 15;

/**
 * One byte for each UTF-16 unit of the password: the low byte, or the high
 * one where the low one is zero (MS-OFFCRYPTO 2.3.7.4). For characters
 * outside ASCII this is one of several ways the password may have been
 * written: see xorPasswordCandidates().
 */
export function xorPasswordBytes(password: Uint8Array): Uint8Array {
  const out = new Uint8Array(password.length >> 1);
  for (let i = 0; i < out.length; i++) out[i] = password[i * 2] || password[i * 2 + 1];
  return out;
}

/** The byte of each character in a Windows code page, for the characters above ASCII; made on first use. */
const reverse = new Map<number, Map<number, number>>();

function codepageByte(codepage: number, code: number): number | undefined {
  if (code < 0x80) return code;
  let table = reverse.get(codepage);
  if (!table) {
    table = new Map();
    const high = new Uint8Array(128);
    for (let i = 0; i < 128; i++) high[i] = 0x80 + i;
    const text = decodeBytes(high, 0, 128, codepage);
    for (let i = 0; i < 128 && i < text.length; i++) {
      const char = text.charCodeAt(i);
      // A place the code page leaves empty reads as a replacement character, or as the control character of its own number.
      if (char !== 0xfffd && !(char >= 0x80 && char < 0xa0) && !table.has(char)) table.set(char, 0x80 + i);
    }
    reverse.set(codepage, table);
  }
  return table.get(code);
}

/** The code pages a password of XOR obfuscation is tried in, after the rule of the specification: Turkish, then Western. */
const XOR_CODEPAGES = [1254, 1252];

/**
 * The single-byte forms a password may have been hashed in, the likeliest
 * first: the rule of the specification (xorPasswordBytes), then the
 * password written in Windows-1254 and in Windows-1252. Excel and Word
 * convert the password to the code page of the system it is typed on, which
 * the file does not name; a code page that lacks one of the characters is
 * left out, and so is a form that is the same as an earlier one. For a
 * password of ASCII characters there is one form.
 *
 * Whichever form fits the 16-bit key and the 16-bit verifier of the file
 * opens it. A form that is not the one the file was written with fits both
 * by chance once in 2^32 times, the verifier alone once in 65 536.
 */
export function xorPasswordCandidates(password: Uint8Array): Uint8Array[] {
  const out = [xorPasswordBytes(password)];
  const units = password.length >> 1;
  for (let c = 0; c < XOR_CODEPAGES.length; c++) {
    const form = new Uint8Array(units);
    let whole = true;
    for (let i = 0; i < units && whole; i++) {
      const byte = codepageByte(XOR_CODEPAGES[c], password[i * 2] | (password[i * 2 + 1] << 8));
      if (byte === undefined) whole = false;
      else form[i] = byte;
    }
    let known = !whole;
    for (let k = 0; k < out.length && !known; k++) known = out[k].length === form.length && out[k].every((byte, i) => byte === form[i]);
    if (known) wipe(form);
    else out.push(form);
  }
  return out;
}

// MS-OFFCRYPTO 2.3.7.2: the tables of the 16-bit key.
// prettier-ignore
const INITIAL_CODE = [0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9, 0x280c, 0xa96a, 0x4ec3];
// prettier-ignore
const XOR_MATRIX = [
  0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09, 0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf,
  0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0, 0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40,
  0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5, 0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a,
  0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9, 0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0,
  0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec, 0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10,
  0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168, 0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c,
  0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad, 0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc,
  0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4,
];
const PAD = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00];

/** The 16-bit key of a password of 1 to 15 single-byte characters (CreateXorKey_Method1). */
export function xorKey(chars: Uint8Array): number {
  if (chars.length === 0 || chars.length > XOR_MAX) return 0;
  let key = INITIAL_CODE[chars.length - 1];
  let element = 0x68;
  for (let i = chars.length - 1; i >= 0; i--) {
    let char = chars[i];
    for (let bit = 0; bit < 7; bit++) {
      if (char & 0x40) key ^= XOR_MATRIX[element];
      char <<= 1;
      element--;
    }
  }
  return key;
}

/** The 16-bit verifier of a password (CreatePasswordVerifier_Method1): also the hash of sheet protection. */
export function xorVerifier(chars: Uint8Array): number {
  let verifier = 0;
  for (let i = chars.length; i >= 0; i--) {
    // The length of the password goes in front of its characters.
    const byte = i === 0 ? chars.length : chars[i - 1];
    verifier = (((verifier >> 14) & 1) | ((verifier << 1) & 0x7fff)) ^ byte;
  }
  return verifier ^ 0xce4b;
}

/**
 * The sixteen bytes a stream is XORed with: the password, filled up with a
 * fixed pad, each byte XORed with a byte of the key and rotated left by
 * `rotate` bits (2 in a workbook, 7 in a Word document).
 */
export function xorArray(chars: Uint8Array, rotate: number): Uint8Array {
  const key = xorKey(chars);
  const out = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    const byte = (i < chars.length ? chars[i] : PAD[i - chars.length]) ^ (i & 1 ? key >> 8 : key & 0xff);
    out[i] = ((byte << rotate) | (byte >> (8 - rotate))) & 0xff;
  }
  return out;
}

/** The characters of an XOR password, checked against the key and the verifier of the file; throws "wrong". */
function xorUnlock(password: Uint8Array, key: number, verifier: number): Uint8Array {
  const forms = xorPasswordCandidates(password);
  let found: Uint8Array | undefined;
  for (let i = 0; i < forms.length; i++) {
    const chars = forms[i];
    // Both numbers are compared whatever the first comparison says. A password of more than fifteen characters has no key.
    const same = chars.length > 0 && chars.length <= XOR_MAX && ((xorKey(chars) ^ key) | (xorVerifier(chars) ^ verifier)) === 0;
    if (same && !found) found = chars;
    else wipe(chars);
  }
  if (!found) throw wrong();
  return found;
}

// ----------------------------------------------------------------- RC4

/** The key of a block of the stream (the caller wipes it), and what zeroes the digest of the password the keys are made from. */
type BlockKey = ((block: number) => Uint8Array) & { wipe(): void };

/** Office 97 RC4 (MS-OFFCRYPTO 2.3.6.2): 40 bits of the password's MD5, stretched with the salt, then MD5 with the block number. */
function rc4Keys(password: Uint8Array, salt: Uint8Array): BlockKey {
  const first = digest('MD5', password);
  const buffer = new Uint8Array(21 * 16);
  for (let i = 0; i < 16; i++) {
    buffer.set(first.subarray(0, 5), i * 21);
    buffer.set(salt, i * 21 + 5);
  }
  const second = digest('MD5', buffer).slice(0, 5);
  wipe(first, buffer);
  return Object.assign((block: number) => digest('MD5', second, le32(block)), { wipe: () => wipe(second) });
}

/** RC4 CryptoAPI (MS-OFFCRYPTO 2.3.5.2): SHA-1 of salt and password, then SHA-1 with the block number, cut to the key size. */
function capiKeys(password: Uint8Array, salt: Uint8Array, keyBits: number): BlockKey {
  const base = digest('SHA1', salt, password);
  return Object.assign(
    (block: number) => {
      const full = digest('SHA1', base, le32(block));
      // A 40-bit key is five bytes of digest and eleven zeros.
      const key = keyBits === 40 ? new Uint8Array(16) : new Uint8Array(keyBits / 8);
      key.set(full.subarray(0, keyBits / 8));
      wipe(full);
      return key;
    },
    { wipe: () => wipe(base) },
  );
}

/** Decrypts the verifier and its digest with the key of block 0 and compares; throws "wrong". */
function rc4Verify(keys: BlockKey, sealed: Uint8Array, hash: 'MD5' | 'SHA1'): void {
  const key = keys(0);
  const rc4 = new Rc4(key);
  const open = sealed.slice();
  rc4.apply(open);
  const expected = createHash(hash).update(open.subarray(0, 16)).digest();
  const same = constantTimeEqual(expected, open.subarray(16));
  rc4.wipe();
  wipe(key, open, expected);
  if (!same) throw wrong();
}

/** Decrypts `data` in place, block by block: every block has its own key. */
function rc4Blocks(data: Uint8Array, keys: BlockKey, blockSize: number): void {
  for (let at = 0, block = 0; at < data.length; at += blockSize, block++) {
    const key = keys(block);
    const rc4 = new Rc4(key);
    rc4.apply(data, at, Math.min(data.length, at + blockSize));
    rc4.wipe();
    wipe(key);
  }
}

/** rc4Verify(), and nothing of the password is left behind when it is wrong. */
function verified(keys: BlockKey, sealed: Uint8Array, hash: 'MD5' | 'SHA1'): void {
  try {
    rc4Verify(keys, sealed, hash);
  } catch (error) {
    keys.wipe();
    throw error;
  }
}

interface Rc4Scheme {
  encryption: FileEncryption;
  keys: BlockKey;
}

/**
 * Reads the encryption header of an RC4 scheme (its version first: 1.1 is
 * Office 97, x.2 the CryptoAPI) and checks the password against it.
 */
function rc4Unlock(header: Uint8Array, password: Uint8Array): Rc4Scheme {
  if (header.length < 4) throw damagedEncryption('the encryption header is cut short');
  const major = header[0] | (header[1] << 8);
  const minor = header[2] | (header[3] << 8);
  if (major === 1 && minor === 1) {
    if (header.length < 52) throw damagedEncryption('the encryption header is cut short');
    const keys = rc4Keys(password, header.subarray(4, 20));
    verified(keys, header.subarray(20, 52), 'MD5');
    return { keys, encryption: { scheme: 'rc4', cipher: 'RC4', keyBits: 40, hash: 'MD5', integrity: false } };
  }
  if (minor === 2 && major >= 2 && major <= 4) {
    const info = parseCryptoApiInfo(header, () => 20);
    if (info.algId !== ALG_RC4) throw unsupportedEncryption(`the cipher 0x${info.algId.toString(16)} in a file of this format`);
    if (info.algIdHash !== 0 && info.algIdHash !== 0x8004) throw unsupportedEncryption(`the digest 0x${info.algIdHash.toString(16)}`);
    const keyBits = info.keyBits === 0 ? 40 : info.keyBits;
    if (keyBits < 40 || keyBits > 128 || keyBits % 8 !== 0) throw damagedEncryption(`an RC4 key is not ${keyBits} bits`);
    if (info.verifierHashSize !== 20) throw damagedEncryption('the digest of the password verifier is not as long as a SHA-1 digest');
    const keys = capiKeys(password, info.salt, keyBits);
    const sealed = new Uint8Array(36);
    sealed.set(info.encryptedVerifier);
    sealed.set(info.encryptedVerifierHash, 16);
    verified(keys, sealed, 'SHA1');
    return { keys, encryption: { scheme: 'rc4-cryptoapi', cipher: 'RC4', keyBits, hash: 'SHA1', integrity: false } };
  }
  throw unsupportedEncryption(`an encryption header of version ${major}.${minor}`);
}

// ----------------------------------------------------------------- xls

const REC_BOF = 0x0809;
const REC_EOF = 0x000a;
const REC_FILEPASS = 0x002f;
const REC_BOUNDSHEET = 0x0085;
/** BOF, FILEPASS, UsrExcl, FileLock, InterfaceHdr, RRDInfo and RRDHead are never encrypted (MS-XLS 2.2.10). */
const XLS_PLAIN = [REC_BOF, REC_FILEPASS, 0x0194, 0x0195, 0x00e1, 0x0196, 0x0138];
/** How many records of the workbook globals are looked at for FILEPASS: it is the second or the third. */
const FILEPASS_WITHIN = 16;

/**
 * Where the FILEPASS record of a workbook stream is (the offset of its
 * content, and its length), or undefined for a stream that is not encrypted.
 * Record headers are never encrypted, so this needs no password.
 */
export function findFilePass(stream: Uint8Array): { offset: number; length: number } | undefined {
  let p = 0;
  for (let i = 0; i < FILEPASS_WITHIN && p + 4 <= stream.length; i++) {
    const id = stream[p] | (stream[p + 1] << 8);
    const length = stream[p + 2] | (stream[p + 3] << 8);
    if (i === 0 && id !== REC_BOF) return undefined;
    if (id === REC_EOF) return undefined;
    if (id === REC_FILEPASS) return p + 4 + length <= stream.length ? { offset: p + 4, length } : undefined;
    p += 4 + length;
  }
  return undefined;
}

/**
 * Decrypts a Workbook stream that has a FILEPASS record. The result is as
 * long as the stream and has every record where it was, FILEPASS included.
 * Throws PasswordError: "wrong", "unsupported", "damaged".
 */
export function decryptXlsStream(stream: Uint8Array, password: Uint8Array): { bytes: Uint8Array; encryption: FileEncryption } {
  const found = findFilePass(stream);
  if (!found) throw damagedEncryption('the workbook has no FILEPASS record');
  const body = stream.subarray(found.offset, found.offset + found.length);
  const biff8 = (stream[4] | (stream[5] << 8)) === 0x0600 || found.length !== 4;
  // BIFF8 says which scheme first; before it there is only XOR, and the record is the key and the verifier.
  const type = biff8 ? body[0] | (body[1] << 8) : 0;
  const rest = biff8 ? body.subarray(2) : body;
  const out = stream.slice();

  if (type === 0) {
    if (rest.length < 4) throw damagedEncryption('the FILEPASS record is cut short');
    const chars = xorUnlock(password, rest[0] | (rest[1] << 8), rest[2] | (rest[3] << 8));
    const array = xorArray(chars, 2);
    wipe(chars);
    for (let p = 0; p + 4 <= stream.length; ) {
      const id = stream[p] | (stream[p + 1] << 8);
      const length = Math.min(stream[p + 2] | (stream[p + 3] << 8), stream.length - p - 4);
      const start = p + 4;
      if (XLS_PLAIN.indexOf(id) < 0) {
        // The place in the array starts from where the record ends (MS-XLS 2.2.10; UNSURE: read from other implementations).
        for (let i = id === REC_BOUNDSHEET ? 4 : 0; i < length; i++) {
          const byte = stream[start + i];
          out[start + i] = (((byte << 3) | (byte >> 5)) & 0xff) ^ array[(start + length + i) & 15];
        }
      }
      p = start + length;
    }
    wipe(array);
    return { bytes: out, encryption: { scheme: 'xor', cipher: 'XOR', keyBits: 16, integrity: false } };
  }
  if (type !== 1) throw unsupportedEncryption(`a FILEPASS record of type ${type}`);

  const scheme = rc4Unlock(rest, password);
  // The key stream runs over the whole stream in blocks of 1024 bytes; what is not encrypted is put back.
  try {
    rc4Blocks(out, scheme.keys, 1024);
  } finally {
    // The keys of the blocks are not needed again: what they are made from is zeroed.
    scheme.keys.wipe();
  }
  for (let p = 0; p + 4 <= stream.length; ) {
    const id = stream[p] | (stream[p + 1] << 8);
    const length = Math.min(stream[p + 2] | (stream[p + 3] << 8), stream.length - p - 4);
    const plain = XLS_PLAIN.indexOf(id) >= 0 ? length : id === REC_BOUNDSHEET ? Math.min(4, length) : 0;
    out.set(stream.subarray(p, p + 4 + plain), p);
    p += 4 + length;
  }
  return { bytes: out, encryption: scheme.encryption };
}

// ----------------------------------------------------------------- doc

const FIB_ENCRYPTED = 0x0100;
const FIB_TABLE_1 = 0x0200;
const FIB_OBFUSCATED = 0x8000;
/** The start of a WordDocument stream that is never encrypted: the FibBase and the word that follows it. */
const DOC_PLAIN = 68;

/** What the header of a WordDocument stream says of its encryption; undefined for a document that has none. */
export function docEncryptionFlags(word: Uint8Array): { obfuscated: boolean; lKey: number; table: '0Table' | '1Table' } | undefined {
  // The first 18 bytes are enough: the flags and lKey.
  if (word.length < 18 || (word[0] | (word[1] << 8)) !== 0xa5ec) return undefined;
  const flags = word[10] | (word[11] << 8);
  if (!(flags & FIB_ENCRYPTED)) return undefined;
  const lKey = (word[14] | (word[15] << 8) | (word[16] << 16)) + word[17] * 0x1000000;
  return { obfuscated: (flags & FIB_OBFUSCATED) !== 0, lKey, table: flags & FIB_TABLE_1 ? '1Table' : '0Table' };
}

export interface DocStreams {
  word: Uint8Array;
  table: Uint8Array;
  data: Uint8Array;
}

/**
 * Decrypts the three streams of a Word document. The WordDocument stream
 * that comes back says it is not encrypted, so that the reader reads it.
 * Throws PasswordError: "wrong", "unsupported", "damaged".
 */
export function decryptDocStreams(streams: DocStreams, password: Uint8Array): DocStreams & { encryption: FileEncryption } {
  const flags = docEncryptionFlags(streams.word);
  if (!flags) throw damagedEncryption('the document does not say it is encrypted');
  const word = streams.word.slice();
  const table = streams.table.slice();
  const data = streams.data.slice();
  let encryption: FileEncryption;
  if (flags.obfuscated) {
    // lKey is the verifier (low half) and the key (high half). UNSURE: the order of the halves is read from other implementations.
    const chars = xorUnlock(password, flags.lKey >>> 16, flags.lKey & 0xffff);
    const array = xorArray(chars, 7);
    wipe(chars);
    const clear = (bytes: Uint8Array, from: number): void => {
      for (let i = from; i < bytes.length; i++) {
        const byte = bytes[i];
        const open = byte ^ array[i & 15];
        // A zero byte, and a byte that would become zero, were left as they are.
        if (byte !== 0 && open !== 0) bytes[i] = open;
      }
    };
    clear(word, DOC_PLAIN);
    clear(table, 0);
    clear(data, 0);
    wipe(array);
    encryption = { scheme: 'xor', cipher: 'XOR', keyBits: 16, integrity: false };
  } else {
    // The header of the encryption is the start of the table stream, lKey bytes long, and is not encrypted itself.
    if (flags.lKey < 4 || flags.lKey > table.length || flags.lKey > 4096) throw damagedEncryption('the encryption header of the document is not inside its table stream');
    const scheme = rc4Unlock(streams.table.subarray(0, flags.lKey), password);
    try {
      rc4Blocks(word, scheme.keys, 512);
      word.set(streams.word.subarray(0, DOC_PLAIN));
      rc4Blocks(table, scheme.keys, 512);
      table.set(streams.table.subarray(0, flags.lKey));
      rc4Blocks(data, scheme.keys, 512);
    } finally {
      scheme.keys.wipe();
    }
    encryption = scheme.encryption;
  }
  word[11] &= ~((FIB_ENCRYPTED | FIB_OBFUSCATED) >> 8) & 0xff;
  return { word, table, data, encryption };
}
