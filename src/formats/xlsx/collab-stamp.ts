// The step of a room a file was written at, kept in the file: one custom document property.
// Only a workbook that is edited together writes it; whatever other properties the file has
// stay as they are.

import type { OpcPackage } from '../../opc';
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

const ours = (root: XmlElement): XmlElement | undefined => childElements(root, 'property').find((property) => property.attrs.get('name') === NAME);

/** The stamp a file carries; undefined when it has none, or one that is not as this module writes it. */
export async function readCollabStamp(pkg: OpcPackage): Promise<CollabStamp | undefined> {
  try {
    const part = await partOf(pkg);
    if (!part) return undefined;
    const bytes = await pkg.readPart(part);
    if (bytes.length > MAX_PART_BYTES) return undefined;
    const property = ours(parseXml(bytes).root);
    if (!property) return undefined;
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
 * there is none to write.
 */
export async function writeCollabStamp(pkg: OpcPackage, stamp: CollabStamp | undefined): Promise<void> {
  const text = stamp ? `${stamp.ep}:${stamp.rev}` : undefined;
  if (text !== undefined && !parseStamp(text)) throw new RangeError('The stamp of the room is not one');
  const part = await partOf(pkg);
  if (part) {
    const bytes = await pkg.readPart(part);
    if (bytes.length > MAX_PART_BYTES) return;
    const doc = parseXml(bytes);
    const root = doc.root;
    const had = ours(root);
    if (text === undefined) {
      if (!had) return;
      root.children = root.children.filter((node) => node !== had);
      pkg.writePart(part, doc);
      return;
    }
    const prefix = root.name.indexOf(':') >= 0 ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
    let pid = 1;
    for (const property of childElements(root, 'property')) {
      const n = Number(property.attrs.get('pid'));
      if (Number.isSafeInteger(n) && n > pid) pid = n;
    }
    const value: XmlElement = { type: 'element', name: 'vt:lpwstr', attrs: new Map(), children: [{ type: 'text', value: text }] };
    if (had) {
      const now = childElements(had).map((child) => (child.children.length === 1 && child.children[0].type === 'text' ? child.children[0].value : '')).join('');
      if (now === text && localName(childElements(had)[0]?.name || '') === 'lpwstr') return;
      had.attrs.set('xmlns:vt', NS_VT);
      had.children = [value];
    } else {
      root.children.push({
        type: 'element',
        name: `${prefix}property`,
        attrs: new Map([
          ['xmlns:vt', NS_VT],
          ['fmtid', FMTID],
          ['pid', String(pid + 1)],
          ['name', NAME],
        ]),
        children: [value],
      });
    }
    pkg.writePart(part, doc);
    return;
  }
  if (text === undefined) return;
  // No custom properties yet: the part, what declares it, and the relationship that leads to it.
  const relsPart = '/_rels/.rels';
  if (!pkg.hasPart(relsPart) || pkg.hasPart(DEFAULT_PART)) return;
  const rels = await pkg.readXml(relsPart);
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
  const encoder = new TextEncoder();
  pkg.writePart(DEFAULT_PART, encoder.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="${NS}" xmlns:vt="${NS_VT}"><property fmtid="${FMTID}" pid="2" name="${NAME}"><vt:lpwstr>${text}</vt:lpwstr></property></Properties>`), CONTENT_TYPE);
  pkg.writePart(relsPart, rels);
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

/**
 * Writes the file of a workbook that is edited together: once what was changed here has its
 * place in the common order (or after a short wait), with what arrives held back until the file
 * is written, so that it is the workbook of one moment; and with the room and the step written
 * into it, so that whoever opens it later joins at that step whatever anybody says of the file.
 * `rev`: the step the file holds, undefined when it is the file of no one step (something of
 * this participant waited, or was typed while the file was written): such a file says nothing
 * of itself.
 */
export async function writeSharedFile(session: StampingSession, pkg: OpcPackage | undefined, take: () => Promise<Uint8Array>): Promise<{ bytes: Uint8Array; rev: number | undefined }> {
  if (session.state === 'closed') {
    // On its own for good: the file is of no room any more.
    if (pkg) await writeCollabStamp(pkg, undefined);
    return { bytes: await take(), rev: undefined };
  }
  if (session.state === 'online') await session.whenSettled(3000);
  for (let attempt = 0; ; attempt++) {
    const rev = session.savableRev;
    // Never a step the file may not be of: after two tries into which something was typed, it says nothing.
    const stamp = attempt < 2 ? session.stampNow : undefined;
    session.pause();
    let bytes: Uint8Array;
    let touched = false;
    try {
      if (pkg) await writeCollabStamp(pkg, stamp);
      bytes = await take();
    } finally {
      touched = session.resume();
    }
    if (touched && stamp !== undefined) continue;
    return { bytes, rev: touched ? undefined : rev };
  }
}
