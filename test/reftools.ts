// The second implementation the tests of encryption are checked against:
// test/tools/offcrypto_ref.py (Python's hashlib and hmac, AES by the openssl
// command line). Where python3 or openssl is missing the checks are skipped,
// and say so; BYM_REQUIRE_REFTOOLS=1 makes them fail instead. The rules of
// saving (never without the password, never with weak random numbers) are
// tested on files the engine encrypts itself, and do not need the tools.

import { spawnSync } from 'node:child_process';
import { test } from 'node:test';

const SCRIPT = new URL('./tools/offcrypto_ref.py', import.meta.url).pathname;

function found(command: string, args: string[]): boolean {
  try {
    return spawnSync(command, args, { stdio: 'ignore' }).status === 0;
  } catch {
    return false;
  }
}

/** Whether both tools are there. BYM_NO_REFTOOLS=1 says they are not: how the tests are run as a machine without them runs them. */
export const REFTOOLS = process.env.BYM_NO_REFTOOLS !== '1' && found('python3', ['-c', 'import hashlib, hmac, struct, zipfile']) && found('openssl', ['version']);

export const REFTOOLS_SKIPPED = 'reference tools not found: independent crypto checks skipped';

if (!REFTOOLS) console.log(`# ${REFTOOLS_SKIPPED}`);

export interface RefResult {
  /** 0 done, 2 wrong password, 3 integrity check failed, 4 not a sound ZIP file, 5 anything else. */
  status: number;
  out: Uint8Array;
  err: string;
}

/** The password as the script takes it: its UTF-16LE bytes in hexadecimal. */
export function passwordArg(password: string): string {
  let out = '';
  for (let i = 0; i < password.length; i++) {
    const code = password.charCodeAt(i);
    out += (code & 0xff).toString(16).padStart(2, '0') + (code >>> 8).toString(16).padStart(2, '0');
  }
  return out;
}

/** Runs one command of the script; `options` are its "--name value" pairs. */
export function ref(command: string, options: Record<string, string | number>, input: Uint8Array = new Uint8Array(0)): RefResult {
  const args = [SCRIPT, command];
  for (const name of Object.keys(options)) args.push(`--${name}`, String(options[name]));
  const done = spawnSync('python3', args, { input, maxBuffer: 1 << 28 });
  return { status: done.status === null ? -1 : done.status, out: new Uint8Array(done.stdout), err: String(done.stderr) };
}

/** As ref(), for a command that has to succeed. */
export function refOk(command: string, options: Record<string, string | number>, input?: Uint8Array): Uint8Array {
  const done = ref(command, options, input);
  if (done.status !== 0) throw new Error(`offcrypto_ref.py ${command} failed (${done.status}): ${done.err}`);
  return done.out;
}

/** A test that needs the reference tools: skipped, visibly, without them. */
export function refTest(name: string, body: () => void | Promise<void>): void {
  if (REFTOOLS) test(name, body);
  else if (process.env.BYM_REQUIRE_REFTOOLS === '1') {
    test(name, () => {
      throw new Error(`${REFTOOLS_SKIPPED} (BYM_REQUIRE_REFTOOLS=1)`);
    });
  } else test(name, { skip: REFTOOLS_SKIPPED }, body);
}
