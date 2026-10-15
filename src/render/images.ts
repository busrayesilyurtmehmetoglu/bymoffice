// Decoding of the pictures of a file, shared by the document view and the
// sheet view: the kind of a picture is told from its bytes, never from a name
// or a content type; what is not a raster picture a browser decodes safely
// (EMF, WMF, TIFF, SVG, ...) is never handed to it; and the decoded pictures
// held at any one time stay within a budget.
//
// Nothing in this module touches the DOM at import time.

/** A decoded picture and its size in its own pixels (which may be fewer than the file has). */
export interface DecodedImage {
  source: CanvasImageSource;
  width: number;
  height: number;
}

/** Decoded pictures held at any one time, in bytes (four per pixel). The limit is strict. */
export const IMAGE_BUDGET_BYTES = 96 * 1024 * 1024;
export const IMAGE_BUDGET_BYTES_COARSE = 48 * 1024 * 1024;
/** A picture with more pixels than this is not decoded at all. */
const MAX_IMAGE_PIXELS = 50000000;
/** One decoded picture takes at most this share of the budget; larger ones are kept at a lower resolution. */
const IMAGE_SHARE = 1 / 8;
/** Pictures decoded at the same time. */
const MAX_IMAGE_LOADS = 2;
/** Smallest side a picture is decoded to, canvas px: zooming in a little does not start a new decode. */
const MIN_DECODED_SIDE = 64;

/** The kind of raster picture the bytes are, by their first bytes; undefined for everything else (EMF, WMF, TIFF, SVG, ...). */
export function sniffImageType(bytes: Uint8Array): string | undefined {
  if (bytes.length < 12) return undefined;
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return 'image/png';
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg';
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) return 'image/gif';
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return 'image/bmp';
  if (bytes[0] === 0x52 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x46 && bytes[8] === 0x57 && bytes[9] === 0x45 && bytes[10] === 0x42 && bytes[11] === 0x50) {
    return 'image/webp';
  }
  return undefined;
}

/**
 * Width and height the picture claims to have, read from its header without
 * decoding it; undefined when the header does not say (the decoder decides then).
 */
export function declaredImageSize(bytes: Uint8Array, type: string): { width: number; height: number } | undefined {
  const u16be = (at: number) => (bytes[at] << 8) | bytes[at + 1];
  const u16le = (at: number) => bytes[at] | (bytes[at + 1] << 8);
  const u32be = (at: number) => bytes[at] * 0x1000000 + ((bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]);
  const i32le = (at: number) => bytes[at] | (bytes[at + 1] << 8) | (bytes[at + 2] << 16) | (bytes[at + 3] << 24);
  const size = (width: number, height: number) => ({ width, height });
  if (type === 'image/png') return bytes.length >= 24 ? size(u32be(16), u32be(20)) : undefined;
  if (type === 'image/gif') return size(u16le(6), u16le(8));
  if (type === 'image/bmp') {
    // An OS/2 header (BITMAPCOREHEADER, 12 bytes) gives the size in two 16-bit words; the others in two 32-bit numbers, the height negative for a picture stored top down.
    if (bytes.length >= 22 && i32le(14) === 12) return size(u16le(18), u16le(20));
    return bytes.length >= 26 ? size(Math.abs(i32le(18)), Math.abs(i32le(22))) : undefined;
  }
  if (type === 'image/jpeg') {
    // Segments until the frame header; every step moves forward, so the loop ends with the data.
    let at = 2;
    while (at + 9 < bytes.length) {
      if (bytes[at] !== 0xff) return undefined;
      const marker = bytes[at + 1];
      if (marker === 0xff) {
        at++;
        continue;
      }
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return size(u16be(at + 7), u16be(at + 5));
      if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
        at += 2;
        continue;
      }
      const length = u16be(at + 2);
      if (length < 2) return undefined;
      at += 2 + length;
    }
    return undefined;
  }
  if (type === 'image/webp' && bytes.length >= 30) {
    const kind = String.fromCharCode(bytes[12], bytes[13], bytes[14], bytes[15]);
    if (kind === 'VP8X') return size(1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16)), 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16)));
    if (kind === 'VP8 ') return size(u16le(26) & 0x3fff, u16le(28) & 0x3fff);
    if (kind === 'VP8L') {
      const bits = bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24);
      return size(1 + (bits & 0x3fff), 1 + ((bits >>> 14) & 0x3fff));
    }
  }
  return undefined;
}

/** Number of pixels the picture claims to have; undefined when its header does not say. */
export function declaredImagePixels(bytes: Uint8Array, type: string): number | undefined {
  const size = declaredImageSize(bytes, type);
  return size ? size.width * size.height : undefined;
}

interface ImageEntry {
  state: 'waiting' | 'loading' | 'ready' | 'failed';
  image?: DecodedImage;
  /** Bytes of the decoded picture held. */
  bytes: number;
  /** update() pass in which the picture was last asked for. */
  used: number;
  close?(): void;
  /** Largest size the picture was asked for, canvas px; 0 when no size was given. */
  wantWidth: number;
  wantHeight: number;
  /** Size of the picture in the file, once it is known. */
  naturalWidth: number;
  naturalHeight: number;
  /** In the queue or being decoded (a picture that is ready is there again when more of its pixels are wanted). */
  pending: boolean;
  /** Pass in which there was no room for the picture: it is not tried again before the next. */
  deferred: number;
}

interface Decoded {
  image: DecodedImage;
  close?(): void;
}

/** The size a picture is decoded to: what the page needs of it, never more than it has or than `maxPixels`. */
export function decodedImageSize(
  naturalWidth: number,
  naturalHeight: number,
  wantWidth: number,
  wantHeight: number,
  maxPixels: number,
): { width: number; height: number } {
  let scale = 1;
  if (wantWidth > 0 && wantHeight > 0) {
    const wanted = Math.max(Math.max(wantWidth, MIN_DECODED_SIDE) / naturalWidth, Math.max(wantHeight, MIN_DECODED_SIDE) / naturalHeight);
    // In steps of a half: a slightly larger zoom finds the pixels already there.
    while (scale / 2 >= wanted) scale /= 2;
  }
  const limit = Math.sqrt(maxPixels / (naturalWidth * naturalHeight));
  if (scale > limit) scale = limit;
  return { width: Math.max(1, Math.floor(naturalWidth * scale)), height: Math.max(1, Math.floor(naturalHeight * scale)) };
}

/**
 * Decodes the pictures of a package on demand and never holds more decoded
 * bytes than its budget. A picture is decoded to the size the page shows it
 * at, not to the size it has in the file, and to no more than a share of the
 * budget; when the budget is used up by pictures of the pages being painted,
 * one more stays a placeholder until there is room.
 */
export class ImageStore {
  /** In order of use: the first entry is the one used longest ago. */
  private entries = new Map<string, ImageEntry>();
  private queue: string[] = [];
  private loading = 0;
  /** Bytes held, and bytes promised to pictures being decoded. */
  private total = 0;
  private reserved = 0;
  /** Decodes of large pictures run one after the other. */
  private heavy: Promise<unknown> = Promise.resolve();
  private destroyed = false;
  /** Counts the passes of the view; a picture used in the current pass is never dropped. */
  pass = 0;

  constructor(
    private readonly source: { readPart(part: string): Promise<Uint8Array> } | undefined,
    readonly budget: number,
    private readonly onLoad: (part: string) => void,
  ) {}

  /** Bytes of decoded pictures held now, those being decoded included. */
  get bytes(): number {
    return this.total + this.reserved;
  }

  /** The decoded picture of a part, undefined while there is none. `width` and `height`: the canvas pixels it will cover. */
  get(part: string, width?: number, height?: number): DecodedImage | undefined {
    if (this.destroyed || !this.source) return undefined;
    const wantWidth = typeof width === 'number' && width > 0 && width < Infinity ? width : 0;
    const wantHeight = typeof height === 'number' && height > 0 && height < Infinity ? height : 0;
    let entry = this.entries.get(part);
    if (!entry) {
      entry = { state: 'waiting', bytes: 0, used: this.pass, wantWidth, wantHeight, naturalWidth: 0, naturalHeight: 0, pending: false, deferred: -1 };
      this.entries.set(part, entry);
      this.enqueue(part, entry);
      return undefined;
    }
    entry.used = this.pass;
    if (wantWidth > entry.wantWidth) entry.wantWidth = wantWidth;
    if (wantHeight > entry.wantHeight) entry.wantHeight = wantHeight;
    if (entry.state === 'waiting') {
      // There was no room for it; a later pass may have some.
      if (entry.deferred !== this.pass) this.enqueue(part, entry);
      return undefined;
    }
    if (entry.state !== 'ready') return undefined;
    // Most recently used last.
    this.entries.delete(part);
    this.entries.set(part, entry);
    const image = entry.image as DecodedImage;
    if (entry.deferred !== this.pass && entry.naturalWidth > 0) {
      // Shown larger than it was decoded for (the zoom grew): decoded again when it has more pixels to give.
      const better = decodedImageSize(entry.naturalWidth, entry.naturalHeight, entry.wantWidth, entry.wantHeight, this.maxPixels());
      if (better.width > image.width * 1.25 || better.height > image.height * 1.25) this.enqueue(part, entry);
    }
    return image;
  }

  /** Whether a picture is being decoded or waits its turn: until then a 'pending' status may still change. */
  get busy(): boolean {
    return this.loading > 0 || this.queue.length > 0;
  }

  /**
   * Whether a picture that get() gave nothing for will ever come: 'failed' for one that
   * cannot be shown (not a raster picture, too large, damaged), 'pending' otherwise.
   */
  status(part: string): 'pending' | 'ready' | 'failed' {
    if (this.destroyed || !this.source) return 'failed';
    const entry = this.entries.get(part);
    return !entry ? 'pending' : entry.state === 'ready' ? 'ready' : entry.state === 'failed' ? 'failed' : 'pending';
  }

  destroy(): void {
    this.destroyed = true;
    this.entries.forEach((entry) => {
      if (entry.close) entry.close();
    });
    this.entries.clear();
    this.queue.length = 0;
    // A picture that is still being decoded keeps its address until the decode ends (see decode): taking it
    // away under the image element is a failed load, which browsers report on the console.
    this.total = 0;
  }

  private maxPixels(): number {
    return Math.floor((this.budget * IMAGE_SHARE) / 4);
  }

  private enqueue(part: string, entry: ImageEntry): void {
    if (entry.pending) return;
    entry.pending = true;
    this.queue.push(part);
    this.pump();
  }

  private pump(): void {
    while (this.loading < MAX_IMAGE_LOADS && this.queue.length > 0) {
      const part = this.queue.shift() as string;
      const entry = this.entries.get(part);
      if (!entry || !entry.pending) continue;
      if (entry.state === 'waiting') entry.state = 'loading';
      this.loading++;
      this.load(part, entry).then(
        (loaded) => this.settle(part, entry, loaded),
        () => this.settle(part, entry, null),
      );
    }
  }

  /** `loaded`: the decoded picture, null when it cannot be shown, undefined when there was no room for it. */
  private settle(part: string, entry: ImageEntry, loaded: Decoded | null | undefined): void {
    this.loading--;
    entry.pending = false;
    if (this.destroyed || this.entries.get(part) !== entry) {
      if (loaded && loaded.close) loaded.close();
      this.pump();
      return;
    }
    if (loaded && loaded.image.width > 0 && loaded.image.height > 0) {
      // A picture decoded again replaces its smaller self.
      if (entry.close) entry.close();
      this.total -= entry.bytes;
      entry.state = 'ready';
      entry.image = loaded.image;
      entry.close = loaded.close;
      entry.bytes = loaded.image.width * loaded.image.height * 4;
      this.total += entry.bytes;
      this.onLoad(part);
    } else {
      if (loaded && loaded.close) loaded.close();
      if (loaded === undefined) entry.deferred = this.pass;
      // What it had before stays: a picture that could not be decoded larger is still shown as it was.
      if (entry.state === 'loading') entry.state = loaded === undefined ? 'waiting' : 'failed';
    }
    this.pump();
  }

  /**
   * Makes room for `bytes` more by dropping pictures that no page painted in
   * the current pass uses, the one used longest ago first. False when the
   * rest is all in use: the budget is not exceeded for anything.
   */
  private makeRoom(bytes: number, keep: ImageEntry): boolean {
    if (this.total + this.reserved + bytes <= this.budget) return true;
    const drop: string[] = [];
    let free = this.budget - this.total - this.reserved;
    this.entries.forEach((entry, part) => {
      if (free >= bytes || entry === keep || entry.state !== 'ready' || entry.pending || entry.used >= this.pass) return;
      free += entry.bytes;
      drop.push(part);
    });
    if (free < bytes) return false;
    for (let i = 0; i < drop.length; i++) {
      const entry = this.entries.get(drop[i]) as ImageEntry;
      if (entry.close) entry.close();
      this.total -= entry.bytes;
      this.entries.delete(drop[i]);
    }
    return true;
  }

  private async load(part: string, entry: ImageEntry): Promise<Decoded | null | undefined> {
    const bytes = await (this.source as { readPart(part: string): Promise<Uint8Array> }).readPart(part);
    const type = sniffImageType(bytes);
    // Vector and other formats are not handed to the browser: an SVG, say, could ask for outside resources.
    if (type === undefined) return null;
    const size = declaredImageSize(bytes, type);
    const pixels = size ? size.width * size.height : undefined;
    if (pixels !== undefined && !(pixels > 0 && pixels <= MAX_IMAGE_PIXELS)) return null;
    const limit = this.maxPixels();
    // Without a size in the header the decoder decides; room is kept for the most a picture may take.
    const target = size ? decodedImageSize(size.width, size.height, entry.wantWidth, entry.wantHeight, limit) : undefined;
    if (size) {
      entry.naturalWidth = size.width;
      entry.naturalHeight = size.height;
      // Asked for again, and what is there is as good as it gets within the budget.
      if (entry.image && target && target.width <= entry.image.width && target.height <= entry.image.height) return undefined;
    }
    const room = target ? target.width * target.height * 4 : limit * 4;
    if (!this.makeRoom(room, entry)) return undefined;
    this.reserved += room;
    try {
      const blob = new Blob([bytes as unknown as BlobPart], { type });
      const decode = () => this.decode(blob, size, target, limit);
      // The decoder holds the whole picture for a moment, whatever size is asked for: large ones go one at a time.
      if (pixels === undefined || pixels * 4 > this.budget / 4) {
        const run = this.heavy.then(decode, decode);
        this.heavy = run.then(
          () => undefined,
          () => undefined,
        );
        return await run;
      }
      return await decode();
    } finally {
      this.reserved -= room;
    }
  }

  private async decode(
    blob: Blob,
    size: { width: number; height: number } | undefined,
    target: { width: number; height: number } | undefined,
    limit: number,
  ): Promise<Decoded | null> {
    if (this.destroyed) return null;
    if (typeof createImageBitmap === 'function') {
      try {
        const smaller = size !== undefined && target !== undefined && (target.width < size.width || target.height < size.height);
        const bitmap = smaller
          ? await createImageBitmap(blob, { resizeWidth: (target as { width: number }).width, resizeHeight: (target as { height: number }).height, resizeQuality: 'medium' })
          : await createImageBitmap(blob);
        let closed = false;
        const close = () => {
          if (closed) return;
          closed = true;
          bitmap.close();
        };
        // A store that was destroyed while the decoder worked has no use for the bitmap, and draws nothing with it.
        if (this.destroyed) {
          close();
          return null;
        }
        try {
          return this.fitted(bitmap, bitmap.width, bitmap.height, close, target, limit);
        } catch (error) {
          // Whatever went wrong while it was drawn smaller, the bitmap is given back; the image element is tried next.
          close();
          throw error;
        }
      } catch {
        // Some browsers decode fewer formats here than in an image element.
      }
    }
    if (this.destroyed || typeof Image !== 'function' || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') return null;
    return new Promise((resolve) => {
      const url = URL.createObjectURL(blob);
      const element = new Image();
      const finish = (ok: boolean) => {
        element.onload = element.onerror = null;
        // The decoded picture stays usable after its address is given back; a store that was destroyed
        // meanwhile gives it back here too.
        URL.revokeObjectURL(url);
        const release = () => {
          element.removeAttribute('src');
        };
        if (this.destroyed) {
          release();
          resolve(null);
          return;
        }
        resolve(ok ? this.fitted(element, element.naturalWidth, element.naturalHeight, release, target, limit) : null);
      };
      element.onload = () => finish(true);
      element.onerror = () => finish(false);
      element.src = url;
    });
  }

  /**
   * The decoded picture as it is kept: itself when it is no larger than was
   * asked for, otherwise drawn down to that size on a canvas of its own, with
   * the large original given back at once (a browser that ignores the resize
   * options of createImageBitmap, or a picture decoded in an image element).
   */
  private fitted(
    source: CanvasImageSource,
    width: number,
    height: number,
    close: () => void,
    target: { width: number; height: number } | undefined,
    limit: number,
  ): Decoded | null {
    if (!(width > 0) || !(height > 0)) {
      close();
      return null;
    }
    const wanted = target || decodedImageSize(width, height, 0, 0, limit);
    // Not one pixel more than room was kept for: the budget is strict.
    if (width * height <= wanted.width * wanted.height) return { image: { source, width, height }, close };
    const canvas = document.createElement('canvas');
    canvas.width = wanted.width;
    canvas.height = wanted.height;
    const context = canvas.getContext('2d');
    let drawn = false;
    if (context) {
      try {
        context.drawImage(source, 0, 0, wanted.width, wanted.height);
        drawn = true;
      } catch {
        // A picture the canvas refuses is not shown.
      }
    }
    close();
    const release = () => {
      canvas.width = 0;
      canvas.height = 0;
    };
    if (!drawn) {
      release();
      return null;
    }
    return { image: { source: canvas, width: wanted.width, height: wanted.height }, close: release };
  }
}
