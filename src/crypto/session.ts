// The password of an open file, for as long as the file is open: saving a
// file that was opened with a password encrypts it with the same password
// again, so the password has to be there when the user saves.
//
// It is kept as bytes (its UTF-16LE form, which is what every scheme hashes)
// and not as a string, because bytes can be zeroed when the file is closed
// and a string cannot. Nothing here prints or serialises it.

import { utf16le, wipe } from './bytes';

export class PasswordSession {
  private bytes: Uint8Array | undefined;
  /** The iterated digest of the password for one salt: saves the rounds of key derivation on every save but the first. */
  private derived: { salt: Uint8Array; base: Uint8Array; rounds: number } | undefined;

  private constructor(bytes: Uint8Array) {
    this.bytes = bytes;
  }

  static of(password: string): PasswordSession {
    return new PasswordSession(utf16le(password));
  }

  /** Whether forget() was called. */
  get forgotten(): boolean {
    return this.bytes === undefined;
  }

  /** The UTF-16LE bytes of the password. Not a copy: do not change or keep them. */
  get password(): Uint8Array {
    if (!this.bytes) throw new Error('The password of this file was forgotten: the file was closed');
    return this.bytes;
  }

  get agileDerived(): { salt: Uint8Array; base: Uint8Array; rounds: number } | undefined {
    return this.derived;
  }

  set agileDerived(value: { salt: Uint8Array; base: Uint8Array; rounds: number } | undefined) {
    if (this.derived && this.derived !== value) wipe(this.derived.base);
    this.derived = value;
  }

  /** A session of its own with the same password: for a second file that is saved with it. */
  clone(): PasswordSession {
    return new PasswordSession(this.password.slice());
  }

  /** Zeroes the password and what was derived from it. */
  forget(): void {
    wipe(this.bytes, this.derived && this.derived.base);
    this.bytes = undefined;
    this.derived = undefined;
  }

  // What a debugger, a log or JSON.stringify() shows of a session.
  toJSON(): string {
    return '[password]';
  }

  toString(): string {
    return '[password]';
  }
}

/** What an open file is saved with. */
export interface Protection {
  session?: PasswordSession;
}

const holders = new WeakMap<object, Protection>();

/** Says what an open file is saved with. */
export function bindProtection(file: object, protection: Protection): void {
  holders.set(file, protection);
}

/** The protection of an open file; made on first use. */
export function protectionOf(file: object): Protection {
  let holder = holders.get(file);
  if (!holder) holders.set(file, (holder = {}));
  return holder;
}

/**
 * Throws for a file that has a password and has forgotten it (it was closed): nothing may be derived from such a
 * file, because what is derived could only be written without the password.
 */
export function assertProtectionAlive(file: object): void {
  const holder = holders.get(file);
  // The getter throws for a session that was forgotten.
  if (holder && holder.session) void holder.session.password;
}

/** The session a file is saved with, if it has one. */
export function sessionOf(file: object): PasswordSession | undefined {
  const holder = holders.get(file);
  return holder ? holder.session : undefined;
}

/**
 * Makes `to` be saved with the password `from` was opened with: a file that
 * was converted to be edited keeps the password of its original.
 */
export function inheritProtection(from: object, to: object): void {
  const source = holders.get(from);
  if (!source || !source.session) return;
  // A file that was closed: its password is gone, and what is made from it cannot be written with it. Throws.
  assertProtectionAlive(from);
  const target = protectionOf(to);
  if (target.session) target.session.forget();
  target.session = source.session.clone();
}
