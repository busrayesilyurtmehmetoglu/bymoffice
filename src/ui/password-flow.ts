// Opening a file that may need a password: the password that was given is
// tried, and for as long as it is missing or wrong somebody is asked, the
// application's `onPassword` or the question of the editor (password-dialog.ts).
// mount() opens its file through this. Nothing here touches the DOM.

import { PasswordError } from '../crypto/errors';
import type { PasswordRequest } from './password-dialog';

export interface PasswordFlowOptions {
  /** The password to try first. */
  password?: string | null;
  /** The application's own question (see MountOptions.onPassword). */
  onPassword?(request: PasswordRequest): string | null | undefined | Promise<string | null | undefined>;
}

/** The question of the editor, as openAsking() uses it. */
export interface PasswordAsker {
  ask(request: PasswordRequest): Promise<string | null>;
  close(): void;
}

/**
 * Opens a file with `openOnce`, which is called with the password to try
 * (undefined the first time, when none was given) and with a signal that
 * says when the user gave up.
 *
 * When `openOnce` rejects with PasswordError "required" or "wrong",
 * `onPassword` is asked, or the question `makeAsker` makes; with neither the
 * error is the caller's. An answer of null gives PasswordError "cancelled".
 * Every other error is passed on as it is.
 */
export async function openAsking<F>(
  openOnce: (password: string | undefined, signal: { readonly aborted: boolean }) => Promise<F>,
  options: PasswordFlowOptions,
  makeAsker?: (onCancel: () => void) => PasswordAsker | undefined,
): Promise<F> {
  let password = options.password === null ? undefined : options.password;
  let attempts = 0;
  let asker: PasswordAsker | undefined;
  // The user giving up while a password is being checked ends the check.
  const signal = { aborted: false };
  try {
    for (;;) {
      try {
        return await openOnce(password, signal);
      } catch (error) {
        if (!(error instanceof PasswordError) || (error.code !== 'required' && error.code !== 'wrong')) throw error;
        if (error.code === 'wrong') attempts++;
        const request: PasswordRequest = { attempts, wrong: error.code === 'wrong' };
        let answer: string | null | undefined;
        if (typeof options.onPassword === 'function') answer = await options.onPassword(request);
        else {
          if (!asker && makeAsker) {
            asker = makeAsker(() => {
              signal.aborted = true;
            });
          }
          if (!asker) throw error;
          answer = await asker.ask(request);
        }
        if (answer === null || answer === undefined) {
          const cancelled = new PasswordError('cancelled', 'The file was not opened: no password was given for it');
          cancelled.format = error.format;
          throw cancelled;
        }
        if (typeof answer !== 'string') throw new TypeError('onPassword() gives the password as a string, or null to give up');
        password = answer;
      }
    }
  } finally {
    if (asker) asker.close();
  }
}
