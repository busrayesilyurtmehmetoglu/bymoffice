// Long work done a slice at a time, so that the page keeps answering: the
// first slice runs at once (small work is done before the call returns), the
// others in later turns of the page.

/** How long one slice may take, milliseconds. */
export const SLICE_MS = 12;
/** How many slices are run at most: a bound for work that would never end. */
const MAX_SLICES = 1000000;

/**
 * Calls `step` until it returns true, in slices of about `sliceMs` with a turn of the page
 * between them. Resolves to true when the work is done, to false when `alive` said no before
 * a slice. What `step` throws rejects the promise.
 */
export function runSliced(step: () => boolean, alive: () => boolean, sliceMs = SLICE_MS): Promise<boolean> {
  return new Promise<boolean>((resolve, reject) => {
    let slices = 0;
    const slice = (): void => {
      try {
        if (!alive()) return resolve(false);
        const start = Date.now();
        do {
          if (step()) return resolve(true);
        } while (Date.now() - start < sliceMs);
        if (++slices > MAX_SLICES) return reject(new RangeError('The work took too long'));
        setTimeout(slice, 0);
      } catch (error) {
        reject(error);
      }
    };
    slice();
  });
}
