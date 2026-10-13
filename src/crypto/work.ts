// Long cryptographic work without freezing the page: a hundred thousand
// rounds of a digest, or a large file run through AES, are done in slices
// with a pause between them, and can be called off.

import { wipe } from './bytes';
import { PasswordError } from './errors';
import { createHash, type HashName } from './hash';

export interface WorkOptions {
  /** Calls the work off: it rejects with PasswordError "cancelled" at the next slice. An AbortSignal fits. */
  signal?: { readonly aborted: boolean };
  /** How long a slice of work may take before the page gets a turn, in milliseconds. Default 12. */
  sliceMs?: number;
  /** What a pause is. Default: a timer of 0 ms. */
  pause?(): Promise<void>;
  /** The clock. Default: Date.now. */
  now?(): number;
}

/** The most rounds a file may ask for: what Office itself allows. More is refused, not computed. */
export const MAX_SPIN_COUNT = 10000000;

const defaultPause = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

export class Slicer {
  private readonly now: () => number;
  private readonly pause: () => Promise<void>;
  private readonly sliceMs: number;
  private readonly signal: { readonly aborted: boolean } | undefined;
  private since: number;

  constructor(options: WorkOptions = {}) {
    this.now = options.now || Date.now;
    this.pause = options.pause || defaultPause;
    this.sliceMs = options.sliceMs !== undefined && options.sliceMs >= 0 ? options.sliceMs : 12;
    this.signal = options.signal;
    this.since = this.now();
  }

  /** Throws when the work was called off. */
  check(): void {
    if (this.signal && this.signal.aborted) throw new PasswordError('cancelled', 'The work on the encrypted file was cancelled');
  }

  /** Called between pieces of work: pauses when the slice is used up. */
  async turn(): Promise<void> {
    this.check();
    if (this.now() - this.since < this.sliceMs) return;
    await this.pause();
    this.check();
    this.since = this.now();
  }
}

export const slicerOf = (work?: WorkOptions | Slicer): Slicer => (work instanceof Slicer ? work : new Slicer(work));

/** Rounds between two looks at the clock. */
const BATCH = 256;

/**
 * The iterated digest of ECMA-376 key derivation: H(salt + password), then
 * `count` times H(round number + previous digest), the round number as four
 * little-endian bytes. Takes its time in slices (see WorkOptions).
 */
export async function spinHash(name: HashName, salt: Uint8Array, password: Uint8Array, count: number, work?: WorkOptions | Slicer): Promise<Uint8Array> {
  if (!(count >= 0 && count <= MAX_SPIN_COUNT) || Math.floor(count) !== count) {
    throw new PasswordError('unsupported', `The file asks for ${String(count).slice(0, 24)} rounds of key derivation; at most ${MAX_SPIN_COUNT} are done`);
  }
  const slicer = slicerOf(work);
  const hash = createHash(name);
  const buffer = new Uint8Array(4 + hash.digestSize);
  const value = buffer.subarray(4);
  try {
    hash.update(salt).update(password).digest(value);
    for (let i = 0; i < count; ) {
      const end = Math.min(count, i + BATCH);
      for (; i < end; i++) {
        buffer[0] = i & 0xff;
        buffer[1] = (i >>> 8) & 0xff;
        buffer[2] = (i >>> 16) & 0xff;
        buffer[3] = i >>> 24;
        hash.update(buffer).digest(value);
      }
      await slicer.turn();
    }
    return value.slice();
  } finally {
    wipe(buffer);
    hash.wipe();
  }
}
