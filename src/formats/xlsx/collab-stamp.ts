// The step of a room a file was written at, kept in the file: one custom document property.
// Only a workbook that is edited together writes it; whatever other properties the file has
// stay as they are.

import type { OpcPackage } from '../../opc';
import { turnOfPackage } from '../../package-turn';
import { ZipArchive } from '../../zip';
import { childElements, localName, parseXml, type XmlElement } from '../../xml';

/** The room (as its relay names it: a new name every time the room is started) and the step of its order a file holds. */
export interface CollabStamp {
  ep: string;
  rev: number;
}

const NAME = 'BymOfficeCollab';
const REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
const CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties';
const NS_VT = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes';
const FMTID = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const DEFAULT_PART = '/docProps/custom.xml';
/** A file with more custom properties than this is left alone: the stamp is not worth reading megabytes for. */
const MAX_PART_BYTES = 1024 * 1024;

async function partOf(pkg: OpcPackage): Promise<string | undefined> {
  const rels = await pkg.relationships('/');
  for (const rel of rels) if (rel.type === REL_TYPE && rel.part && pkg.hasPart(rel.part)) return rel.part;
  return undefined;
}

function parseStamp(text: string): CollabStamp | undefined {
  const match = /^([A-Za-z0-9_-]{1,64}):(0|[1-9][0-9]{0,15})$/.exec(text);
  if (!match) return undefined;
  const rev = Number(match[2]);
  return Number.isSafeInteger(rev) ? { ep: match[1], rev } : undefined;
}

const ours = (root: XmlElement): XmlElement[] => childElements(root, 'property').filter((property) => property.attrs.get('name') === NAME);

/** The part of custom properties with the stamp alone. */
function fresh(text: string): Uint8Array {
  return new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${NS}" xmlns:vt="${NS_VT}"><property fmtid="${FMTID}" pid="2" name="${NAME}"><vt:lpwstr>${text}</vt:lpwstr></property></Properties>`);
}

/**
 * The stamp a file carries; undefined when it has none, or one that is not as this module writes
 * it. What a CollabSession is given as `stamp` when it is made for the file (`file.package`).
 */
export async function readCollabStamp(pkg: OpcPackage): Promise<CollabStamp | undefined> {
  try {
    const part = await partOf(pkg);
    if (!part) return undefined;
    const bytes = await pkg.readPart(part);
    if (bytes.length > MAX_PART_BYTES) return undefined;
    const all = ours(parseXml(bytes).root);
    // Two of the name: there is no saying which is meant, and the file says nothing.
    if (all.length !== 1) return undefined;
    const property = all[0];
    let text = '';
    for (const value of childElements(property)) for (const node of value.children) if (node.type === 'text') text += node.value;
    return parseStamp(text);
  } catch {
    // A part that cannot be read holds no stamp; the file is opened by what it is known by.
    return undefined;
  }
}

/**
 * Writes the stamp into the package, or takes it out (`undefined`). The part and its
 * relationship are made when the file has none; a file without a stamp is not touched when
 * there is none to write. False when the file could not be made to say it: its part of custom
 * properties is too large or cannot be read (it is left exactly as it is, and no stamp is read
 * from such a part either).
 */
export async function writeCollabStamp(pkg: OpcPackage, stamp: CollabStamp | undefined): Promise<boolean> {
  const text = stamp ? `${stamp.ep}:${stamp.rev}` : undefined;
  if (text !== undefined && !parseStamp(text)) throw new RangeError('The stamp of the room is not one');
  const part = await partOf(pkg);
  if (part) {
    let doc;
    try {
      const bytes = await pkg.readPart(part);
      if (bytes.length > MAX_PART_BYTES) return text === undefined;
      doc = parseXml(bytes);
    } catch {
      // Not XML this engine reads (or with a DTD, which it never reads): not a byte of it is touched.
      return text === undefined;
    }
    const root = doc.root;
    const had = ours(root);
    if (text === undefined) {
      if (had.length === 0) return true;
      root.children = root.children.filter((node) => had.indexOf(node as XmlElement) < 0);
      pkg.writePart(part, doc);
      return true;
    }
    // A part that holds the stamp and nothing else is written as it is when it is made: the same workbook at the
    // same step is the same file, whoever writes it and however often it was stamped before.
    if (had.length === 1 && root.children.filter((node) => node.type === 'element').length === 1 && root.name === 'Properties' && root.attrs.get('xmlns') === NS) {
      pkg.writePart(part, fresh(text));
      return true;
    }
    if (had.length === 1) {
      const only = had[0];
      const values = childElements(only);
      const now = values.map((child) => (child.children.length === 1 && child.children[0].type === 'text' ? child.children[0].value : '')).join('');
      if (now === text && values.length === 1 && localName(values[0].name) === 'lpwstr') return true;
      only.attrs.set('xmlns:vt', NS_VT);
      only.children = [{ type: 'element', name: 'vt:lpwstr', attrs: new Map(), children: [{ type: 'text', value: text }] }];
      pkg.writePart(part, doc);
      return true;
    }
    // None, or several of the name (then none of them is kept).
    if (had.length > 0) root.children = root.children.filter((node) => had.indexOf(node as XmlElement) < 0);
    const prefix = root.name.indexOf(':') >= 0 ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
    // The least number from 2 that no property has (they are 32-bit numbers: "one more than the largest" may be none).
    const taken = new Set<number>();
    for (const property of childElements(root, 'property')) taken.add(Number(property.attrs.get('pid')));
    let pid = 2;
    while (taken.has(pid)) pid++;
    root.children.push({
      type: 'element',
      name: `${prefix}property`,
      attrs: new Map([
        ['xmlns:vt', NS_VT],
        ['fmtid', FMTID],
        ['pid', String(pid)],
        ['name', NAME],
      ]),
      children: [{ type: 'element', name: 'vt:lpwstr', attrs: new Map(), children: [{ type: 'text', value: text }] }],
    });
    pkg.writePart(part, doc);
    return true;
  }
  if (text === undefined) return true;
  // No custom properties yet: the part, what declares it, and the relationship that leads to it.
  const relsPart = '/_rels/.rels';
  if (!pkg.hasPart(relsPart) || pkg.hasPart(DEFAULT_PART)) return false;
  let rels;
  try {
    rels = await pkg.readXml(relsPart);
  } catch {
    return false;
  }
  const ids = new Set(childElements(rels.root, 'Relationship').map((rel) => rel.attrs.get('Id') || ''));
  let n = ids.size + 1;
  while (ids.has(`rId${n}`)) n++;
  const relPrefix = rels.root.name.indexOf(':') >= 0 ? rels.root.name.slice(0, rels.root.name.indexOf(':') + 1) : '';
  rels.root.children.push({
    type: 'element',
    name: `${relPrefix}Relationship`,
    attrs: new Map([
      ['Id', `rId${n}`],
      ['Type', REL_TYPE],
      ['Target', DEFAULT_PART.slice(1)],
    ]),
    children: [],
  });
  pkg.writePart(DEFAULT_PART, fresh(text), CONTENT_TYPE);
  pkg.writePart(relsPart, rels);
  return true;
}

/** What writing a file needs of a session a workbook is edited together in. */
export interface StampingSession {
  readonly state: string;
  readonly savableRev: number | undefined;
  readonly stampNow: CollabStamp | undefined;
  whenSettled(ms: number): Promise<unknown>;
  pause(): void;
  resume(): boolean;
}

/** A file written by writeSharedFile(). */
export interface SharedFile {
  bytes: Uint8Array;
  /** The step of the room the file holds; undefined when it is the file of no one step. */
  rev: number | undefined;
  /** The file could not be made to say its step (its custom properties cannot be read or are too large). */
  unstamped?: true;
}

/**
 * Writes the file of a workbook that is edited together: once what was changed here has its
 * place in the common order (or after a short wait), with what arrives held back until the file
 * is written, so that it is the workbook of one moment; and with the room and the step written
 * into it, so that whoever opens it later joins at that step whatever anybody says of the file.
 * `rev`: the step the file holds, undefined when it is the file of no one step (something of
 * this participant waited, or was typed while the file was written): such a file says nothing
 * of itself.
 *
 * `pkg` is the package of the file (`file.package`) and `take` writes it (`() => file.save()`);
 * mount() saves this way, and a program that runs a CollabSession without mount() does the same,
 * then tells the room with `session.announceSaved(versions, rev)`.
 */
export async function writeSharedFile(
  session: StampingSession,
  pkg: OpcPackage | undefined,
  take: () => Promise<Uint8Array>,
  /**
   * Runs what it is given when nothing else writes the package: the stamp and the save are one step,
   * which no other save of the file may cut into. Without it the turn of the file the package belongs
   * to is taken (and `take` may be `() => file.save()`: called from within the turn, it does not wait
   * for it).
   */
  lock?: <T>(run: () => Promise<T>) => Promise<T>,
): Promise<SharedFile> {
  const turn = lock ? undefined : turnOfPackage(pkg);
  const locked = lock || (turn ? turn.locked : <T>(run: () => Promise<T>): Promise<T> => run());
  const taken = turn ? (): Promise<Uint8Array> => turn.inside(take) : take;
  if (session.state === 'closed') {
    // On its own for good: the file is of no room any more.
    return locked(async () => {
      if (pkg) await writeCollabStamp(pkg, undefined);
      return { bytes: await taken(), rev: undefined };
    });
  }
  if (session.state === 'online') await session.whenSettled(3000);
  for (let attempt = 0; ; attempt++) {
    const done = await locked(async (): Promise<SharedFile | undefined> => {
      const rev = session.savableRev;
      // Never a step the file may not be of: after two tries into which something was typed, it says nothing.
      const stamp = attempt < 2 ? session.stampNow : undefined;
      session.pause();
      let bytes: Uint8Array;
      let touched = false;
      try {
        if (pkg && stamp !== undefined && !(await writeCollabStamp(pkg, stamp))) {
          // The file cannot be made to say its step (its custom properties cannot be read): it is saved all the same,
          // as the file of no step, and whoever opens it joins as with any file the room has not saved.
          bytes = await taken();
          return { bytes, rev: undefined, unstamped: true };
        }
        if (pkg && stamp === undefined) await writeCollabStamp(pkg, undefined);
        bytes = await taken();
      } finally {
        touched = session.resume();
      }
      if (touched && stamp !== undefined) return undefined;
      return { bytes, rev: touched ? undefined : rev };
    });
    if (done) return done;
  }
}

/** Most bytes, unpacked, of two packages that are compared part by part. */
const MAX_COMPARED_BYTES = 512 * 1024 * 1024;

/**
 * Whether two packages are the same file whatever they were packed with: the same parts by name, each with
 * the same content. (The order of the parts, how each was compressed and what the archive says beside them do
 * not count: two browsers write the same workbook as different bytes.) False for what is no archive.
 */
export async function samePackage(a: Uint8Array, b: Uint8Array): Promise<boolean> {
  try {
    const first = ZipArchive.read(a);
    const second = ZipArchive.read(b);
    const names = first.names().slice().sort();
    const others = second.names().slice().sort();
    if (names.length !== others.length) return false;
    // Before anything is unpacked: the same names, each part as long as its counterpart says it is, and no more in
    // all than is looked at. (A part that claims to unpack to more than its counterpart is another part: a file
    // that is small and unpacks to gigabytes is turned away here, by what its directory says.)
    let total = 0;
    for (let i = 0; i < names.length; i++) {
      if (names[i] !== others[i]) return false;
      const x = first.entry(names[i]);
      const y = second.entry(names[i]);
      if (!x || !y || x.uncompressedSize !== y.uncompressedSize) return false;
      total += x.uncompressedSize * 2;
      if (total > MAX_COMPARED_BYTES) return false;
    }
    // Unpacking stops at the length the directory gave (the reader of archives sees to that).
    for (const name of names) {
      const x = await first.read(name);
      const y = await second.read(name);
      if (x.length !== y.length) return false;
      for (let i = 0; i < x.length; i++) if (x[i] !== y[i]) return false;
    }
    return true;
  } catch {
    return false;
  }
}

/**
 * The file this workbook would be stored as now, to compare with what someone else stored: written from
 * a COPY of the package, which gets the stamp, so that the package of the file itself is read and never
 * changed for it. `save`: the bytes of the file as it is (called with the turn of the package held);
 * `open`: a package of bytes. `rev` undefined: the file of no one step (something of this participant
 * waits, was typed meanwhile, or the bytes are no package - a file that is kept with a password).
 */
export async function proofFile(
  session: StampingSession,
  save: () => Promise<Uint8Array>,
  open: (bytes: Uint8Array) => Promise<OpcPackage>,
  locked: <T>(run: () => Promise<T>) => Promise<T> = (run) => run(),
): Promise<{ bytes: Uint8Array; rev: number | undefined }> {
  if (session.state !== 'online') return { bytes: new Uint8Array(0), rev: undefined };
  await session.whenSettled(3000);
  const taken = await locked(async () => {
    const rev = session.savableRev;
    const stamp = session.stampNow;
    session.pause();
    let bytes: Uint8Array;
    let touched = false;
    try {
      bytes = await save();
    } finally {
      touched = session.resume();
    }
    return { bytes, rev: touched ? undefined : rev, stamp: touched ? undefined : stamp };
  });
  if (taken.rev === undefined || taken.stamp === undefined) return { bytes: taken.bytes, rev: undefined };
  try {
    const copy = await open(taken.bytes);
    if (!(await writeCollabStamp(copy, taken.stamp))) return { bytes: taken.bytes, rev: undefined };
    return { bytes: await copy.save(), rev: taken.rev };
  } catch {
    return { bytes: taken.bytes, rev: undefined };
  }
}
