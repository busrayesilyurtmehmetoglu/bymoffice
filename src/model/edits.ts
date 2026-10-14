// How often a model was changed: counted where changes are applied, so that whoever writes a file
// from a model over several moments can tell whether the model stayed the one it began with.

const counts = new WeakMap<object, number>();

/** The models that are being read for a file right now. */
const reading = new Set<object>();
/** Those of them that were changed while they were read. */
const broken = new Set<object>();
/** How often a model was changed when it was last read for a file, and how often it was read for one. */
const readAt = new WeakMap<object, number>();
const readings = new WeakMap<object, number>();

/** A change was applied to the model. */
export function touchModel(model: object): void {
  counts.set(model, (counts.get(model) || 0) + 1);
  if (reading.size > 0 && reading.has(model)) broken.add(model);
}

/** A number that is another one after every change that was applied to the model. */
export function editsOf(model: object): number {
  return counts.get(model) || 0;
}

/**
 * Reads a model for a file, in one go: `read` runs without waiting for anything, so that what it
 * makes is the model of one moment - the one editsOf() names when it starts, which readMomentOf()
 * answers afterwards. Throws when `read` did wait (it answered a promise) or the model was changed
 * while it ran: a writer that does either writes the file of no moment.
 */
export function readModel<T>(model: object, read: () => T): T {
  // Within a reading of the same model: the same moment.
  if (reading.has(model)) return read();
  const at = editsOf(model);
  reading.add(model);
  broken.delete(model);
  let out: T;
  try {
    out = read();
  } finally {
    reading.delete(model);
    readings.set(model, (readings.get(model) || 0) + 1);
  }
  const changed = broken.delete(model) || editsOf(model) !== at;
  if (out !== null && (typeof out === 'object' || typeof out === 'function') && typeof (out as { then?: unknown }).then === 'function') {
    throw new Error('A model is read for a file in one go: what reads it may not wait for anything');
  }
  if (changed) throw new Error('A model was changed while it was read for a file');
  readAt.set(model, at);
  return out;
}

/** What editsOf() answered when the model was last read for a file: the moment that file is of. Undefined for a model no file was written from. */
export function readMomentOf(model: object): number | undefined {
  return readAt.get(model);
}

/** How often the model was read for a file (readModel), whatever came of it. */
export function readingsOf(model: object): number {
  return readings.get(model) || 0;
}

/**
 * Sets the count back to what it was: for what applies changes and takes every one of them back
 * in one go, without anything else running in between (a command that is tried on the model to see
 * what it would do). The model is what it was, and a file that is being written from it is whole.
 */
export function restoreEdits(model: object, count: number): void {
  counts.set(model, count);
}

/** A file could not be written because its model kept changing while it was written. Nothing was written half. */
export class SaveInterruptedError extends Error {
  constructor() {
    super('The file was being edited the whole time it was written; it was not saved. It is saved when the editing pauses');
    this.name = 'SaveInterruptedError';
  }
}

/** How long nothing may change before a file whose model changed while it was written is written again. */
export const QUIET_MS = 150;
/** How long a save waits in all for the editing to pause before it gives up. */
export const MAX_STEADY_MS = 10000;

export interface SteadyOptions {
  quietMs?: number;
  maxMs?: number;
  sleep?(ms: number): Promise<void>;
  now?(): number;
}

/**
 * Writes a file from a model: `write` is run, and its file is answered. A writer reads the model
 * in one go (readModel), and its file is the model of that moment whatever was changed before and
 * after: nothing is written again for it. What is left here is for a writer that does not: its
 * file is answered only when the model was not changed while it ran - a file that was written
 * while the model changed may hold some parts of before the change and some of after it, and is
 * the file of no moment. (So may what `write` throws: a writer that read half a change may trip
 * over it.) It is written again then, once the changes have paused for a moment - not at once,
 * into the next change. When they do not pause for as long as a save may wait,
 * SaveInterruptedError is thrown.
 */
export async function writeSteady<T>(model: object | undefined, write: () => Promise<T>, options: SteadyOptions = {}): Promise<T> {
  if (!model) return write();
  const quiet = options.quietMs === undefined ? QUIET_MS : options.quietMs;
  const most = options.maxMs === undefined ? MAX_STEADY_MS : options.maxMs;
  const sleep = options.sleep || ((ms: number) => new Promise<void>((resolve) => setTimeout(resolve, ms)));
  const now = options.now || (() => Date.now());
  const started = now();
  for (;;) {
    const before = editsOf(model);
    const told = readingsOf(model);
    let out: T | undefined;
    let failed: unknown;
    let threw = false;
    try {
      out = await write();
    } catch (error) {
      threw = true;
      failed = error;
    }
    // Read in one go (the writer said so), or not changed meanwhile: the file, or the failure, of one moment.
    if (readingsOf(model) !== told || editsOf(model) === before) {
      if (threw) throw failed;
      return out as T;
    }
    // Changed while it was written: again, when nothing has changed for a moment.
    for (;;) {
      if (now() - started >= most) throw new SaveInterruptedError();
      const seen = editsOf(model);
      await sleep(quiet);
      if (editsOf(model) === seen) break;
    }
  }
}
