// How often a model was changed: counted where changes are applied, so that whoever writes a file
// from a model over several moments can tell whether the model stayed the one it began with.

const counts = new WeakMap<object, number>();

/** A change was applied to the model. */
export function touchModel(model: object): void {
  counts.set(model, (counts.get(model) || 0) + 1);
}

/** A number that is another one after every change that was applied to the model. */
export function editsOf(model: object): number {
  return counts.get(model) || 0;
}

/** A file could not be written because its model kept changing while it was written. Nothing was written half. */
export class SaveInterruptedError extends Error {
  constructor() {
    super('The file was changed again and again while it was written; it was not saved. Save again');
    this.name = 'SaveInterruptedError';
  }
}

/** How often a file is written anew when its model changed while it was written. */
export const MAX_WRITE_TRIES = 5;

/**
 * Writes a file from a model: `write` is run, and its file is answered only when the model was
 * not changed while it ran - a file that was written while the model changed may hold some parts
 * of before the change and some of after it, and is the file of no moment. It is written again
 * then; when the model changes every time, SaveInterruptedError is thrown.
 */
export async function writeSteady<T>(model: object | undefined, write: () => Promise<T>): Promise<T> {
  if (!model) return write();
  for (let attempt = 0; attempt < MAX_WRITE_TRIES; attempt++) {
    const before = editsOf(model);
    const out = await write();
    if (editsOf(model) === before) return out;
  }
  throw new SaveInterruptedError();
}
