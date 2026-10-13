// Compound File Binary writer: the smallest one that writes what Office
// reads. Version 3, sectors of 512 bytes, a mini stream for streams under
// 4096 bytes, as many FAT and DIFAT sectors as the content needs. It writes
// the container of a password-protected docx or xlsx (see crypto/ooxml.ts);
// doc and xls files are not written.
//
// The file is laid out first and filled in afterwards: allocateCfb() takes
// the sizes of the streams and gives a view of each stream inside the file,
// so that a large stream is written straight to its place and never copied.
// Every stream lies in one run of sectors.

const SECTOR = 512;
const MINI_SECTOR = 64;
const MINI_CUTOFF = 4096;
const ENTRY = 128;
const PER_SECTOR = SECTOR / 4;

const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT_SECTOR = 0xfffffffd;
const DIFAT_SECTOR = 0xfffffffc;
/** Sector numbers stop here (MAXREGSECT). */
const MAX_SECTORS = 0xfffffffa;
/** A stream of a version 3 file is smaller than 2 GB. */
const MAX_STREAM = 0x80000000;

const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;
const TYPE_ROOT = 5;
const RED = 0;
const BLACK = 1;

export interface CfbStreamSize {
  /** Path from the root, "/"-separated; the storages on the way are made. */
  path: string;
  size: number;
}

export interface CfbLayout {
  /** The whole file. Its streams are zeros until they are written through stream(). */
  bytes: Uint8Array;
  /** The bytes of a stream inside `bytes`: writing to them writes the file. */
  stream(path: string): Uint8Array;
}

interface Node {
  name: string;
  type: number;
  children: Node[];
  size: number;
  /** Index in the directory. */
  id: number;
  left: number;
  right: number;
  child: number;
  colour: number;
  start: number;
  /** Where the content is in the file. */
  offset: number;
}

const node = (name: string, type: number, size = 0): Node => ({ name, type, children: [], size, id: 0, left: FREE, right: FREE, child: FREE, colour: BLACK, start: 0, offset: 0 });

function checkName(name: string, path: string): void {
  if (name.length === 0) throw new Error(`Compound file: empty name in ${JSON.stringify(path)}`);
  if (name.length > 31) throw new Error(`Compound file: the name ${JSON.stringify(name)} is longer than 31 characters`);
  if (/[/\\:!]/.test(name)) throw new Error(`Compound file: "/", "\\", ":" and "!" cannot be in a name (${JSON.stringify(name)})`);
}

/** How siblings are ordered: the shorter name first, then by the upper-case code units (MS-CFB 2.6.4). */
function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  // toUpperCase() may change the length ("ß"); such names are compared as they are.
  if (x.length !== a.length || y.length !== b.length) return a < b ? -1 : a > b ? 1 : 0;
  return x < y ? -1 : x > y ? 1 : 0;
}

/**
 * Links sorted siblings into a red-black tree and returns its root. The
 * middle one of a run is its root, so the two sides never differ by more
 * than one entry: every level but the last is full. The entries of the last
 * level are red where it is not full, everything else black, which gives
 * every path the same number of black entries.
 */
function link(siblings: Node[]): number {
  const count = siblings.length;
  if (count === 0) return FREE;
  // The number of entries on the shortest way from the root to a missing child.
  const full = Math.floor(Math.log2(count + 1) + 1e-9);
  const build = (low: number, high: number, depth: number): number => {
    if (low > high) return FREE;
    const middle = (low + high) >> 1;
    const entry = siblings[middle];
    entry.colour = depth > full ? RED : BLACK;
    entry.left = build(low, middle - 1, depth + 1);
    entry.right = build(middle + 1, high, depth + 1);
    return entry.id;
  };
  return build(0, count - 1, 1);
}

/** Lays out a compound file for streams of the given sizes. */
export function allocateCfb(streams: CfbStreamSize[]): CfbLayout {
  const root = node('Root Entry', TYPE_ROOT);
  const byPath = new Map<string, Node>();
  for (let i = 0; i < streams.length; i++) {
    const { path, size } = streams[i];
    if (!(size >= 0 && size < MAX_STREAM) || Math.floor(size) !== size) throw new Error(`Compound file: a stream of ${size} bytes cannot be written`);
    const parts = path.split('/');
    let parent = root;
    for (let k = 0; k < parts.length; k++) {
      const name = parts[k];
      checkName(name, path);
      const last = k === parts.length - 1;
      let found: Node | undefined;
      for (let c = 0; c < parent.children.length && !found; c++) if (compareNames(parent.children[c].name, name) === 0) found = parent.children[c];
      if (last) {
        if (found) throw new Error(found.type === TYPE_STORAGE ? `Compound file: ${JSON.stringify(path)} is both a stream and a storage` : `Compound file: the stream ${JSON.stringify(path)} is there twice`);
        const stream = node(name, TYPE_STREAM, size);
        parent.children.push(stream);
        byPath.set(path, stream);
      } else {
        if (found && found.type !== TYPE_STORAGE) throw new Error(`Compound file: ${JSON.stringify(parts.slice(0, k + 1).join('/'))} is both a stream and a storage`);
        if (!found) parent.children.push((found = node(name, TYPE_STORAGE)));
        parent = found;
      }
    }
  }

  // The directory: the root, then every storage's children after the storage itself.
  const entries: Node[] = [root];
  for (let i = 0; i < entries.length; i++) {
    const storage = entries[i];
    storage.children.sort((a, b) => compareNames(a.name, b.name));
    for (let c = 0; c < storage.children.length; c++) {
      storage.children[c].id = entries.length;
      entries.push(storage.children[c]);
    }
  }
  for (let i = 0; i < entries.length; i++) entries[i].child = link(entries[i].children);

  // The mini stream: the small streams one after another, each in whole mini sectors.
  let miniSectors = 0;
  let bigSectors = 0;
  for (let i = 1; i < entries.length; i++) {
    const entry = entries[i];
    if (entry.type !== TYPE_STREAM || entry.size === 0) continue;
    if (entry.size < MINI_CUTOFF) {
      entry.start = miniSectors;
      miniSectors += Math.ceil(entry.size / MINI_SECTOR);
    } else bigSectors += Math.ceil(entry.size / SECTOR);
  }
  const directorySectors = Math.ceil((entries.length * ENTRY) / SECTOR);
  const miniFatSectors = Math.ceil(miniSectors / PER_SECTOR);
  const miniStreamSectors = Math.ceil((miniSectors * MINI_SECTOR) / SECTOR);
  const content = directorySectors + miniFatSectors + miniStreamSectors + bigSectors;
  // The FAT describes every sector, its own and those of the DIFAT included; 109 FAT sectors are named in the header.
  let fatSectors = 0;
  let difatSectors = 0;
  for (;;) {
    const needed = Math.ceil((content + fatSectors + difatSectors) / PER_SECTOR);
    const more = needed > 109 ? Math.ceil((needed - 109) / (PER_SECTOR - 1)) : 0;
    if (needed === fatSectors && more === difatSectors) break;
    fatSectors = needed;
    difatSectors = more;
  }
  const total = difatSectors + fatSectors + content;
  if (total >= MAX_SECTORS || (total + 1) * SECTOR >= 4294967296) throw new Error('Compound file: the content is too large for a file of this kind (4 GB)');

  const bytes = new Uint8Array((total + 1) * SECTOR);
  const view = new DataView(bytes.buffer);
  const at = (sector: number): number => (sector + 1) * SECTOR;

  // Sectors, in this order: DIFAT, FAT, directory, mini FAT, mini stream, the large streams.
  const firstFat = difatSectors;
  const firstDirectory = firstFat + fatSectors;
  const firstMiniFat = firstDirectory + directorySectors;
  const firstMiniStream = firstMiniFat + miniFatSectors;
  let next = firstMiniStream + miniStreamSectors;

  const fat = (sector: number, value: number): void => view.setUint32(at(firstFat) + sector * 4, value, true);
  const chain = (first: number, count: number): void => {
    for (let i = 0; i < count; i++) fat(first + i, i === count - 1 ? END : first + i + 1);
  };
  for (let i = 0; i < fatSectors * PER_SECTOR; i++) fat(i, FREE);
  for (let i = 0; i < difatSectors; i++) fat(i, DIFAT_SECTOR);
  for (let i = 0; i < fatSectors; i++) fat(firstFat + i, FAT_SECTOR);
  chain(firstDirectory, directorySectors);
  chain(firstMiniFat, miniFatSectors);
  chain(firstMiniStream, miniStreamSectors);

  // The mini FAT: every small stream is one run.
  for (let i = 0; i < miniFatSectors * PER_SECTOR; i++) view.setUint32(at(firstMiniFat) + i * 4, FREE, true);
  for (let i = 1; i < entries.length; i++) {
    const entry = entries[i];
    if (entry.type !== TYPE_STREAM) continue;
    if (entry.size === 0) {
      entry.start = END;
    } else if (entry.size < MINI_CUTOFF) {
      const count = Math.ceil(entry.size / MINI_SECTOR);
      for (let k = 0; k < count; k++) view.setUint32(at(firstMiniFat) + (entry.start + k) * 4, k === count - 1 ? END : entry.start + k + 1, true);
      entry.offset = at(firstMiniStream) + entry.start * MINI_SECTOR;
    } else {
      const count = Math.ceil(entry.size / SECTOR);
      entry.start = next;
      entry.offset = at(next);
      chain(next, count);
      next += count;
    }
  }
  root.start = miniSectors > 0 ? firstMiniStream : END;
  root.size = miniSectors * MINI_SECTOR;

  // The header.
  bytes.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
  view.setUint16(24, 0x3e, true); // minor version
  view.setUint16(26, 3, true); // major version
  view.setUint16(28, 0xfffe, true); // little-endian
  view.setUint16(30, 9, true); // sectors of 512 bytes
  view.setUint16(32, 6, true); // mini sectors of 64 bytes
  view.setUint32(40, 0, true); // directory sectors: not counted in version 3
  view.setUint32(44, fatSectors, true);
  view.setUint32(48, firstDirectory, true);
  view.setUint32(56, MINI_CUTOFF, true);
  view.setUint32(60, miniFatSectors > 0 ? firstMiniFat : END, true);
  view.setUint32(64, miniFatSectors, true);
  view.setUint32(68, difatSectors > 0 ? 0 : END, true);
  view.setUint32(72, difatSectors, true);
  for (let i = 0; i < 109; i++) view.setUint32(76 + i * 4, i < fatSectors ? firstFat + i : FREE, true);
  // The FAT sectors the header has no room for, 127 to a DIFAT sector, each of which names the next.
  for (let d = 0; d < difatSectors; d++) {
    const base = at(d);
    for (let k = 0; k < PER_SECTOR - 1; k++) {
      const index = 109 + d * (PER_SECTOR - 1) + k;
      view.setUint32(base + k * 4, index < fatSectors ? firstFat + index : FREE, true);
    }
    view.setUint32(base + (PER_SECTOR - 1) * 4, d === difatSectors - 1 ? END : d + 1, true);
  }

  // The directory. Entries that are not used name no neighbours.
  const directory = at(firstDirectory);
  for (let i = 0; i < directorySectors * (SECTOR / ENTRY); i++) {
    const p = directory + i * ENTRY;
    const entry = i < entries.length ? entries[i] : undefined;
    view.setUint32(p + 68, entry ? entry.left : FREE, true);
    view.setUint32(p + 72, entry ? entry.right : FREE, true);
    view.setUint32(p + 76, entry ? entry.child : FREE, true);
    if (!entry) continue;
    for (let k = 0; k < entry.name.length; k++) view.setUint16(p + k * 2, entry.name.charCodeAt(k), true);
    view.setUint16(p + 64, (entry.name.length + 1) * 2, true);
    view.setUint8(p + 66, entry.type);
    view.setUint8(p + 67, entry.colour);
    view.setUint32(p + 116, entry.type === TYPE_STORAGE ? 0 : entry.start, true);
    view.setUint32(p + 120, entry.size, true);
  }

  return {
    bytes,
    stream(path: string): Uint8Array {
      const entry = byPath.get(path);
      if (!entry) throw new Error(`No such stream: ${path}`);
      return bytes.subarray(entry.offset, entry.offset + entry.size);
    },
  };
}

/** A compound file of the given streams. */
export function writeCfb(streams: { path: string; data: Uint8Array }[]): Uint8Array {
  const layout = allocateCfb(streams.map((stream) => ({ path: stream.path, size: stream.data.length })));
  for (let i = 0; i < streams.length; i++) layout.stream(streams[i].path).set(streams[i].data);
  return layout.bytes;
}
