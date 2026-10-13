// A password-protected docx or xlsx: a compound file with the streams
// EncryptionInfo (how) and EncryptedPackage (the ZIP package, encrypted).
// This reads both schemes of ECMA-376 and writes the agile one, with the
// \x06DataSpaces storage Office expects next to the two streams.

import type { CfbFile } from '../cfb';
import { allocateCfb } from '../cfb-write';
import { AgileEncryptor, agileDecryptPackage, agilePackageSize, agileUnlock, parseAgileInfo, type AgileDecryptOptions, type AgileEncryptOptions, type AgileInfo } from './agile';
import { wipe } from './bytes';
import { damagedEncryption, unsupportedEncryption } from './errors';
import type { HashName } from './hash';
import type { PasswordSession } from './session';
import { parseStandardInfo, standardDecryptPackage, standardUnlock, type StandardInfo } from './standard';
import { slicerOf, type WorkOptions } from './work';

/**
 * The password Office encrypts with when a file has to be encrypted and the
 * user gave none (a workbook whose structure is protected, for one). Such a
 * file opens without a question, in Office and here.
 */
export const DEFAULT_PASSWORD = 'VelvetSweatshop';

/** How an encrypted file is, or was, encrypted. */
export interface FileEncryption {
  /**
   * "agile" and "standard": a docx or xlsx (ECMA-376); "rc4-cryptoapi", "rc4"
   * and "xor": a doc or xls of Office 2002-2003, Office 97-2000 and before.
   */
  scheme: 'agile' | 'standard' | 'rc4-cryptoapi' | 'rc4' | 'xor';
  cipher: 'AES' | 'RC4' | 'XOR';
  /** Bits of the key; 16 for the XOR scheme, which is no encryption. */
  keyBits: number;
  hash?: HashName;
  /** Rounds of key derivation. */
  spinCount?: number;
  /**
   * Whether the file carries an integrity check (which was verified when it was opened). Only agile encryption has
   * one. An agile file without it (false) could have been changed by someone who does not know the password: open()
   * says so in `warnings`, and refuses the file with `requireIntegrity`.
   */
  integrity: boolean;
  /**
   * Set (true) when the file opens with the password Office uses when the user gave none (DEFAULT_PASSWORD): anyone
   * can open it. Such a file is encrypted, and it keeps nothing secret.
   */
  defaultPassword?: true;
}

export function isEncryptedPackage(cfb: Pick<CfbFile, 'has'>): boolean {
  return cfb.has('EncryptionInfo') && cfb.has('EncryptedPackage');
}

/** The largest EncryptionInfo stream that is looked at. */
const MAX_INFO = 1 << 21;

type ParsedInfo = { scheme: 'agile'; info: AgileInfo } | { scheme: 'standard'; info: StandardInfo };

function parseInfo(cfb: Pick<CfbFile, 'has' | 'stream' | 'entries'>): ParsedInfo {
  const entry = cfb.entries.find((candidate) => candidate.path === 'EncryptionInfo');
  if (!entry || entry.type !== 'stream') throw damagedEncryption('there is no EncryptionInfo stream');
  if (entry.size > MAX_INFO) throw damagedEncryption('the EncryptionInfo stream is larger than any is');
  const stream = cfb.stream('EncryptionInfo');
  if (stream.length < 8) throw damagedEncryption('the EncryptionInfo stream is cut short');
  const major = stream[0] | (stream[1] << 8);
  const minor = stream[2] | (stream[3] << 8);
  if (major === 4 && minor === 4) return { scheme: 'agile', info: parseAgileInfo(stream) };
  if ((major === 3 || major === 4) && minor === 3) throw unsupportedEncryption('the extensible scheme (a third-party encryption module)');
  if ((major === 2 || major === 3 || major === 4) && minor === 2) return { scheme: 'standard', info: parseStandardInfo(stream) };
  throw unsupportedEncryption(`an EncryptionInfo of version ${major}.${minor}`);
}

function describe(parsed: ParsedInfo): FileEncryption {
  if (parsed.scheme === 'agile') {
    const { keyData, password, integrity } = parsed.info;
    return { scheme: 'agile', cipher: 'AES', keyBits: keyData.keyBits, hash: password.hash, spinCount: password.spinCount, integrity: integrity !== undefined };
  }
  return { scheme: 'standard', cipher: 'AES', keyBits: parsed.info.keyBits, hash: 'SHA1', spinCount: 50000, integrity: false };
}

/** How the package is encrypted, without a password. Throws PasswordError for what is not read. */
export function packageEncryption(cfb: Pick<CfbFile, 'has' | 'stream' | 'entries'>): FileEncryption {
  return describe(parseInfo(cfb));
}

/**
 * Decrypts the package of an encrypted docx or xlsx. Rejects with
 * PasswordError: "wrong" for a password that is not the file's, "damaged"
 * for a file that fails its integrity check or is cut short, "unsupported"
 * for a scheme that is not read.
 */
export async function decryptPackage(
  cfb: Pick<CfbFile, 'has' | 'stream' | 'entries'>,
  password: Uint8Array,
  work?: WorkOptions & AgileDecryptOptions,
): Promise<{ bytes: Uint8Array; encryption: FileEncryption }> {
  const parsed = parseInfo(cfb);
  const slicer = slicerOf(work);
  let key: Uint8Array | undefined;
  try {
    if (parsed.scheme === 'agile') {
      key = await agileUnlock(parsed.info, password, slicer);
      return { bytes: await agileDecryptPackage(parsed.info, key, cfb.stream('EncryptedPackage'), slicer, { requireIntegrity: work ? work.requireIntegrity : undefined }), encryption: describe(parsed) };
    }
    key = await standardUnlock(parsed.info, password, slicer);
    return { bytes: await standardDecryptPackage(parsed.info, key, cfb.stream('EncryptedPackage'), slicer), encryption: describe(parsed) };
  } finally {
    wipe(key);
  }
}

// ----------------------------------------------------------------- writing

/** A text as Office's data spaces store it: its length in bytes, UTF-16LE, filled up to four bytes. */
function lengthPrefixed(text: string): number[] {
  const out = u32(text.length * 2);
  for (let i = 0; i < text.length; i++) out.push(text.charCodeAt(i) & 0xff, text.charCodeAt(i) >> 8);
  while (out.length % 4 !== 0) out.push(0);
  return out;
}

const u32 = (value: number): number[] => [value & 0xff, (value >> 8) & 0xff, (value >> 16) & 0xff, (value >>> 24) & 0xff];
/** Reader, updater and writer version, each 1.0. */
const VERSIONS = [1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0];

/**
 * The streams of the \x06DataSpaces storage (MS-OFFCRYPTO 2.1 and 2.3.4.1 to
 * 2.3.4.3): they say that the stream EncryptedPackage went through the
 * transform "StrongEncryptionTransform". The same bytes for every file.
 * UNSURE: written from the specification and from what other writers
 * produce; not opened in Office.
 */
function dataSpaces(): { path: string; data: Uint8Array }[] {
  const component = lengthPrefixed('EncryptedPackage');
  const space = lengthPrefixed('StrongEncryptionDataSpace');
  const transform = lengthPrefixed('StrongEncryptionTransform');
  const id = lengthPrefixed('{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}');
  const map = [...u32(8), ...u32(1), ...u32(4 + 4 + 4 + component.length + space.length), ...u32(1), ...u32(0), ...component, ...space];
  const primary = [
    // The header of the transform: its length up to the name, its type, its identifier and its name.
    ...u32(8 + id.length),
    ...u32(1),
    ...id,
    ...lengthPrefixed('Microsoft.Container.EncryptionTransform'),
    ...VERSIONS,
    // EncryptionTransformInfo: no name, no block size, no cipher mode; four is what every writer puts last.
    ...u32(0),
    ...u32(0),
    ...u32(0),
    ...u32(4),
  ];
  return [
    { path: '\u0006DataSpaces/Version', data: new Uint8Array([...lengthPrefixed('Microsoft.Container.DataSpaces'), ...VERSIONS]) },
    { path: '\u0006DataSpaces/DataSpaceMap', data: new Uint8Array(map) },
    { path: '\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace', data: new Uint8Array([...u32(8), ...u32(1), ...transform]) },
    { path: '\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary', data: new Uint8Array(primary) },
  ];
}

export type EncryptPackageOptions = Omit<AgileEncryptOptions, 'derived'>;

/**
 * Encrypts a package (the bytes of a docx or an xlsx) with the password of
 * the session and gives the compound file Office opens: agile encryption,
 * AES-256 in CBC mode, SHA-512, 100 000 rounds, an HMAC over the encrypted
 * package. Salts and keys are new for every call and come from the random
 * number generator of the system; where there is none the call rejects, and
 * nothing is written. Only the iterated digest of the password is kept in
 * the session, so that the next save of the same file takes no time.
 */
export async function encryptPackage(plain: Uint8Array, session: PasswordSession, options: EncryptPackageOptions = {}): Promise<Uint8Array> {
  const slicer = slicerOf(options);
  const spinCount = options.spinCount === undefined ? 100000 : options.spinCount;
  const kept = session.agileDerived;
  const encryptor = await AgileEncryptor.create(session.password, { ...options, spinCount, derived: kept && kept.rounds === spinCount ? kept : undefined });
  try {
    // A copy: the encryptor zeroes its own when this file is written. Not kept for a session that was closed meanwhile.
    if ((!kept || kept.rounds !== spinCount) && !session.forgotten) {
      session.agileDerived = { salt: encryptor.derived.salt.slice(), base: encryptor.derived.base.slice(), rounds: spinCount };
    }
    const spaces = dataSpaces();
    const layout = allocateCfb([
      ...spaces.map((stream) => ({ path: stream.path, size: stream.data.length })),
      { path: 'EncryptionInfo', size: encryptor.infoSize },
      { path: 'EncryptedPackage', size: agilePackageSize(plain.length) },
    ]);
    for (let i = 0; i < spaces.length; i++) layout.stream(spaces[i].path).set(spaces[i].data);
    // Straight into its place in the file: the package is never held twice.
    const target = layout.stream('EncryptedPackage');
    await encryptor.encryptInto(plain, target, slicer);
    await slicer.turn();
    layout.stream('EncryptionInfo').set(await encryptor.info(target, slicer));
    return layout.bytes;
  } finally {
    encryptor.wipe();
  }
}
