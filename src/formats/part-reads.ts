// A file is written from its model in one go: everything the writer needs of the package is read
// first (which takes moments), then the model is read and every part produced without waiting for
// anything, so that the file is the model of one moment whatever is changed around the write.
//
// What the writer needs of the package depends on the model, and is not known before the model is
// read. So the reading is tried: a part it asks for that was not read ends the try, the part is
// read, and the reading starts over from the model as it is then. Each try reads at least one part
// more of a package that has a number of them, so that this ends however long the editing goes on.

import { readModel } from '../model/edits';
import { relationshipsOf, relsPartFor, type OpcPackage, type Relationship } from '../opc';
import { parseXml } from '../xml';

/** A reading asked for parts that were not read yet: it starts over once they were. */
class PartsWanted {
  constructor(readonly parts: string[]) {}
}

/** What a part leads to is read along with a part that was not foreseen, up to this many parts. */
const MAX_ALONG = 64;

/** The parts of a package as they are while a file is written, for what cannot wait for them. */
export class PartReads {
  /** By the name the package has the part under. null: a part that could not be read (see `failed`). */
  private readonly loaded = new Map<string, Uint8Array | null>();
  /** The names the package has its parts under: as they are spelled, and the first of each without regard to case (as the package itself finds a part). */
  private readonly exact = new Set<string>();
  private readonly folded = new Map<string, string>();
  private readonly failed = new Map<string, unknown>();
  private readonly related = new Map<string, Relationship[]>();
  /** Set while what a reading may need is gathered: the parts asked for and not read. */
  private wishes: Set<string> | undefined;
  /** How often the parts were read from the package: a number for the tests. */
  rounds = 0;

  constructor(readonly pkg: OpcPackage) {
    for (const name of pkg.partNames()) {
      this.exact.add(name);
      const lower = name.toLowerCase();
      if (!this.folded.has(lower)) this.folded.set(lower, name);
    }
  }

  /** The name the package has a part under; undefined for a part it does not have. (The package is not changed while a file is written from it.) */
  private nameOf(part: string): string | undefined {
    return this.exact.has(part) ? part : this.folded.get(part.toLowerCase());
  }

  /** Whether asking for the part gives an answer at once. */
  ready(part: string): boolean {
    const name = this.nameOf(part);
    return name === undefined || this.loaded.has(name);
  }

  /**
   * The part as the package has it; undefined for a part the package does not have. A part that
   * was not read yet ends the reading (or, while what a reading may need is gathered, is noted
   * and answered as if it were not there).
   */
  bytes(part: string): Uint8Array | undefined {
    const name = this.nameOf(part);
    if (name === undefined) return undefined;
    const kept = this.loaded.get(name);
    if (kept !== undefined) {
      if (kept === null) throw this.failed.get(name);
      return kept;
    }
    if (this.wishes) {
      this.wishes.add(name);
      return undefined;
    }
    throw new PartsWanted([name]);
  }

  /** Ends the reading when any of the parts was not read yet, with all of them asked for at once. */
  need(parts: Iterable<string>): void {
    const missing: string[] = [];
    for (const part of parts) if (!this.ready(part)) missing.push(part);
    if (missing.length === 0) return;
    if (this.wishes) for (const part of missing) this.wishes.add(part);
    else throw new PartsWanted(missing);
  }

  /** The relationships of a part ("/": those of the package), as OpcPackage.relationships() gives them. */
  relationships(source = '/'): Relationship[] {
    const relsPart = relsPartFor(source);
    if (this.nameOf(relsPart) === undefined) return [];
    let list = this.related.get(source);
    if (!list) {
      const bytes = this.bytes(relsPart);
      if (!bytes) return [];
      list = relationshipsOf(source, parseXml(bytes));
      this.related.set(source, list);
    }
    return list;
  }

  /** The main document part, as OpcPackage.mainPart() gives it. */
  mainPart(): string | undefined {
    const found = this.relationships('/').find((rel) => rel.type.endsWith('/officeDocument'));
    return (found && found.part) || undefined;
  }

  /** Runs `gather`, which asks for what a reading may need; answers the parts it asked for that were not read. */
  wished(gather: () => void): string[] {
    const wishes = (this.wishes = new Set());
    try {
      gather();
    } catch (error) {
      // What cannot be foreseen is found out by the reading itself, which says what is wrong.
    } finally {
      this.wishes = undefined;
    }
    return [...wishes];
  }

  /** Reads parts from the package. One that cannot be read is an error for whoever asks for it, not here. */
  async load(parts: Iterable<string>): Promise<void> {
    let any = false;
    for (const part of parts) {
      const name = this.nameOf(part);
      if (name === undefined || this.loaded.has(name)) continue;
      any = true;
      try {
        this.loaded.set(name, await this.pkg.readPart(name));
      } catch (error) {
        this.loaded.set(name, null);
        this.failed.set(name, error);
      }
    }
    if (any) this.rounds++;
  }

  /** Reads parts that were asked for and not foreseen, and with each what it leads to: who needs a part often needs those as well. */
  async loadWithNeighbours(parts: string[]): Promise<void> {
    await this.load(parts);
    const along: string[] = [];
    for (const part of parts) {
      if (along.length >= MAX_ALONG) break;
      const relsPart = relsPartFor(part);
      const relsName = this.nameOf(relsPart);
      if (/\.rels$/i.test(part) || relsName === undefined) continue;
      await this.load([relsName]);
      const bytes = this.loaded.get(relsName);
      if (!bytes) continue;
      try {
        for (const rel of relationshipsOf(part, parseXml(bytes))) if (rel.part !== null && along.length < MAX_ALONG) along.push(rel.part, relsPartFor(rel.part));
      } catch (error) {
        // Relationships that cannot be read lead nowhere here.
      }
    }
    await this.load(along);
  }
}

/** For the tests: with `foresee` off nothing is read ahead, and every part is asked for by a reading that starts over for it. */
export const oneGo = { foresee: true };

/** What the tests ask of the last write of a package: how often its parts were read, and how often the model. */
export interface OneGoTrace {
  /** How often the model was read: once, and once more for every time a part was found missing. */
  readings: number;
  /** How often parts were read from the package. */
  rounds: number;
  /** How long the reading that made the file took, in milliseconds: for that long nothing else ran. */
  readMs: number;
}
const traces = new WeakMap<object, OneGoTrace>();

/** How the last file was written from the model: for the tests. */
export function lastOneGo(model: object): OneGoTrace | undefined {
  return traces.get(model);
}

/**
 * Writes from a model in one go. `foresee` asks (reads.bytes, reads.relationships) for the parts
 * the writer may need, looking at the model no more than it takes to tell; `produce` reads the
 * model and writes, and waits for nothing. `produce` runs again from the start whenever it asked
 * for a part that was not read - so it changes nothing that lasts before it has everything it needs.
 */
export async function writeInOneGo<T>(pkg: OpcPackage, model: object, foresee: (reads: PartReads) => void, produce: (reads: PartReads) => T): Promise<T> {
  const reads = new PartReads(pkg);
  // Each try that fails reads a part more, and there are no more parts than the package has.
  const most = pkg.partNames().length + 8;
  let readings = 0;
  for (let attempt = 0; ; attempt++) {
    // What can be foreseen, in passes: what a part leads to is known once the part was read.
    for (let pass = 0; pass < 8 && oneGo.foresee; pass++) {
      const wishes = reads.wished(() => foresee(reads));
      if (wishes.length === 0) break;
      await reads.load(wishes);
    }
    readings++;
    const began = Date.now();
    const tried = readModel(model, (): { out: T } | PartsWanted => {
      try {
        return { out: produce(reads) };
      } catch (error) {
        if (error instanceof PartsWanted) return error;
        throw error;
      }
    });
    if (!(tried instanceof PartsWanted)) {
      traces.set(model, { readings, rounds: reads.rounds, readMs: Date.now() - began });
      return tried.out;
    }
    if (attempt >= most) throw new Error('Cannot write the file: its parts could not be gathered');
    if (oneGo.foresee) await reads.loadWithNeighbours(tried.parts);
    else await reads.load(tried.parts);
  }
}
