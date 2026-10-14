// The turn of a file's package: whatever writes the package does so alone.

/** Where a file keeps the turn of whoever writes its package. */
export const PACKAGE_TURN = '__bymPackageTurn';

export interface PackageTurn<Options> {
  /** Runs `run` when nothing else writes the package. A save of the file that `run` starts at once (before it waits for anything) is part of the turn. */
  locked<T>(run: () => Promise<T>): Promise<T>;
  /** The save for whoever has the turn (the file's own save() would wait for the turn it is called in). */
  save(options?: Options): Promise<Uint8Array>;
  /** Calls `run` as part of the turn that is held: a save of the file it starts at once does not wait for the turn. */
  inside<T>(run: () => T): T;
}

/** The turn of each package, by the package: whoever has the package finds the turn of its file. */
const turns = new WeakMap<object, PackageTurn<unknown>>();

/** Makes the turn of a file. `raw`: the save that writes the package, without waiting. */
export function makeTurn<Options>(raw: (options?: Options) => Promise<Uint8Array>): PackageTurn<Options> & { held(): boolean } {
  let last: Promise<unknown> = Promise.resolve();
  /** Set for the moment in which what has the turn is called: what it starts in that moment is the turn's own. */
  let own = false;
  const inside = <T>(run: () => T): T => {
    const was = own;
    own = true;
    try {
      return run();
    } finally {
      own = was;
    }
  };
  const locked = <T>(run: () => Promise<T>): Promise<T> => {
    // From within the turn (called at once by what holds it): it is the turn, and waits for nothing.
    if (own) return run();
    const go = (): Promise<T> => inside(run);
    const next = last.then(go, go);
    last = next.catch(() => undefined);
    return next;
  };
  return { locked, save: raw, inside, held: () => own };
}

export function registerTurn(pkg: object, turn: PackageTurn<unknown>): void {
  turns.set(pkg, turn);
}

/** The turn of the file a package belongs to; undefined for a package that is no file's. */
export function turnOfPackage(pkg: object | undefined): PackageTurn<unknown> | undefined {
  return pkg ? turns.get(pkg) : undefined;
}

/**
 * The turn of a file's package, for what has to prepare the package and save it as one step.
 * A file that was made elsewhere (a stand-in of a test) has no turn to wait for.
 */
export function packageTurn<Options>(file: { save(options?: Options): Promise<Uint8Array> }): PackageTurn<Options> {
  const turn = (file as unknown as { [name: string]: unknown })[PACKAGE_TURN] as PackageTurn<Options> | undefined;
  return turn || { locked: (run) => run(), save: (options) => file.save(options), inside: (run) => run() };
}
