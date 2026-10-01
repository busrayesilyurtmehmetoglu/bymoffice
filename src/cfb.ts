// Compound File Binary reader: the container behind legacy .doc and .xls.
// Read-only for now; the writer comes with binary-format saving.

const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const MAX_REGULAR_SECTOR = 0xfffffffa;
const NO_STREAM = 0xffffffff;
const DIRECTORY_ENTRY_SIZE = 128;
const MINI_SECTOR_SIZE = 64;

const TYPE_STORAGE = 1;
const TYPE_STREAM = 2;

export interface CfbEntry {
  /** Path from the root, "/"-separated: "WordDocument", "ObjectPool/_1/Ole". */
  path: string;
  name: string;
  type: 'storage' | 'stream';
  size: number;
}

interface RawEntry {
  name: string;
  type: number;
  left: number;
  right: number;
  child: number;
  start: number;
  size: number;
}

const utf16 = new TextDecoder('utf-16le');

export class CfbFile {
  readonly entries: CfbEntry[] = [];
  private readonly raw = new Map<string, RawEntry>();
  private readonly view: DataView;
  private readonly sectorSize: number;
  private readonly miniCutoff: number;
  private readonly fat: number[] = [];
  private readonly miniFat: number[] = [];
  private readonly miniStream: Uint8Array = new Uint8Array(0);

  static isCfb(bytes: Uint8Array): boolean {
    return bytes.length >= 512 && SIGNATURE.every((byte, i) => bytes[i] === byte);
  }

  static read(bytes: Uint8Array): CfbFile {
    return new CfbFile(bytes);
  }

  private constructor(private readonly bytes: Uint8Array) {
    if (!CfbFile.isCfb(bytes)) throw new Error('Not a compound file');
    const view = (this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength));
    const majorVersion = view.getUint16(26, true);
    this.sectorSize = 1 << view.getUint16(30, true);
    if (this.sectorSize !== 512 && this.sectorSize !== 4096) throw new Error('Corrupt compound file: bad sector size');
    const fatSectorCount = view.getUint32(44, true);
    const directoryStart = view.getUint32(48, true);
    this.miniCutoff = view.getUint32(56, true);
    const miniFatStart = view.getUint32(60, true);
    let difatSector = view.getUint32(68, true);

    const perSector = this.sectorSize / 4;
    const maxSectors = Math.ceil(bytes.length / this.sectorSize);
    const fatSectors: number[] = [];
    for (let i = 0; i < 109 && fatSectors.length < fatSectorCount; i++) {
      const sector = view.getUint32(76 + i * 4, true);
      if (sector <= MAX_REGULAR_SECTOR) fatSectors.push(sector);
    }
    for (let hops = 0; difatSector <= MAX_REGULAR_SECTOR && fatSectors.length < fatSectorCount; hops++) {
      if (hops > maxSectors) throw new Error('Corrupt compound file: DIFAT chain loop');
      const offset = this.offsetOf(difatSector);
      for (let i = 0; i < perSector - 1 && fatSectors.length < fatSectorCount; i++) {
        const sector = view.getUint32(offset + i * 4, true);
        if (sector <= MAX_REGULAR_SECTOR) fatSectors.push(sector);
      }
      difatSector = view.getUint32(offset + (perSector - 1) * 4, true);
    }
    for (const sector of fatSectors) {
      const offset = this.offsetOf(sector);
      for (let i = 0; i < perSector; i++) this.fat.push(view.getUint32(offset + i * 4, true));
    }

    const miniFatBytes = this.readChain(miniFatStart);
    const miniFatView = new DataView(miniFatBytes.buffer);
    for (let i = 0; i + 4 <= miniFatBytes.length; i += 4) this.miniFat.push(miniFatView.getUint32(i, true));

    const directory = this.readChain(directoryStart);
    const dirView = new DataView(directory.buffer);
    const rawEntries: RawEntry[] = [];
    for (let p = 0; p + DIRECTORY_ENTRY_SIZE <= directory.length; p += DIRECTORY_ENTRY_SIZE) {
      const nameBytes = Math.min(64, Math.max(0, dirView.getUint16(p + 64, true) - 2));
      const low = dirView.getUint32(p + 120, true);
      rawEntries.push({
        name: utf16.decode(directory.subarray(p, p + nameBytes)),
        type: dirView.getUint8(p + 66),
        left: dirView.getUint32(p + 68, true),
        right: dirView.getUint32(p + 72, true),
        child: dirView.getUint32(p + 76, true),
        start: dirView.getUint32(p + 116, true),
        // Version 3 files may leave garbage in the high half of the size.
        size: majorVersion === 3 ? low : low + dirView.getUint32(p + 124, true) * 2 ** 32,
      });
    }
    const root = rawEntries[0];
    if (!root) throw new Error('Corrupt compound file: empty directory');
    this.miniStream = this.readChain(root.start).subarray(0, root.size);

    // Children of a storage form a binary tree linked through left/right.
    const seen = new Set<number>();
    const walk = (id: number, prefix: string): void => {
      if (id === NO_STREAM || id >= rawEntries.length || seen.has(id)) return;
      seen.add(id);
      const entry = rawEntries[id];
      walk(entry.left, prefix);
      const path = prefix + entry.name;
      if (entry.type === TYPE_STORAGE || entry.type === TYPE_STREAM) {
        const isStorage = entry.type === TYPE_STORAGE;
        this.entries.push({ path, name: entry.name, type: isStorage ? 'storage' : 'stream', size: isStorage ? 0 : entry.size });
        this.raw.set(path, entry);
        if (isStorage) walk(entry.child, `${path}/`);
      }
      walk(entry.right, prefix);
    };
    walk(root.child, '');
  }

  private offsetOf(sector: number): number {
    return (sector + 1) * this.sectorSize;
  }

  private chain(start: number, table: number[]): number[] {
    const sectors: number[] = [];
    for (let sector = start; sector <= MAX_REGULAR_SECTOR; sector = table[sector] ?? NO_STREAM) {
      if (sectors.length > table.length) throw new Error('Corrupt compound file: sector chain loop');
      sectors.push(sector);
    }
    return sectors;
  }

  private readChain(start: number): Uint8Array {
    const sectors = this.chain(start, this.fat);
    const out = new Uint8Array(sectors.length * this.sectorSize);
    sectors.forEach((sector, i) => {
      const offset = this.offsetOf(sector);
      out.set(this.bytes.subarray(offset, offset + this.sectorSize), i * this.sectorSize);
    });
    return out;
  }

  has(path: string): boolean {
    return this.raw.has(path);
  }

  /** Returns the content of a stream. */
  stream(path: string): Uint8Array {
    const entry = this.raw.get(path);
    if (!entry || entry.type !== TYPE_STREAM) throw new Error(`Stream not found: ${path}`);
    if (entry.size >= this.miniCutoff) return this.readChain(entry.start).subarray(0, entry.size);

    const sectors = this.chain(entry.start, this.miniFat);
    const out = new Uint8Array(sectors.length * MINI_SECTOR_SIZE);
    sectors.forEach((sector, i) => {
      const offset = sector * MINI_SECTOR_SIZE;
      out.set(this.miniStream.subarray(offset, offset + MINI_SECTOR_SIZE), i * MINI_SECTOR_SIZE);
    });
    return out.subarray(0, entry.size);
  }
}
