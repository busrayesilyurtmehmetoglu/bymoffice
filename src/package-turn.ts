// The turn of a file's package: whatever writes the package does so alone.

/** Where a file keeps the turn of whoever writes its package, and the save that is called by who has the turn. */
export const PACKAGE_TURN = '__bymPackageTurn';

export interface PackageTurn<Options> {
  /** Runs `run` when nothing else writes the package. */
  locked<T>(run: () => Promise<T>): Promise<T>;
  /** The save for whoever has the turn (the file's own save() would wait for the turn it is called in). */
  save(options?: Options): Promise<Uint8Array>;
}

/**
 * The turn of a file's package, for what has to prepare the package and save it as one step.
 * A file that was made elsewhere (a stand-in of a test) has no turn to wait for.
 */
export function packageTurn<Options>(file: { save(options?: Options): Promise<Uint8Array> }): PackageTurn<Options> {
  const turn = (file as unknown as { [name: string]: unknown })[PACKAGE_TURN] as PackageTurn<Options> | undefined;
  return turn || { locked: (run) => run(), save: (options) => file.save(options) };
}
