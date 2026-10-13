// What goes wrong with a password-protected file.

/**
 * - `required`: the file is encrypted and no password was given;
 * - `wrong`: the password does not open the file;
 * - `cancelled`: the user (or the application) gave up being asked, or the work was called off;
 * - `unsupported`: the file is encrypted in a way this engine does not read;
 * - `damaged`: the encrypted file is cut short, contradicts itself or fails its integrity check.
 */
export type PasswordErrorCode = 'required' | 'wrong' | 'cancelled' | 'unsupported' | 'damaged';

/** What kind of file the encrypted one is, as far as that can be told without its password. */
export type EncryptedFormat = 'ooxml' | 'doc' | 'xls';

/**
 * A password-protected file could not be opened. The message never holds the
 * password, a key or anything derived from them.
 */
export class PasswordError extends Error {
  override name = 'PasswordError';
  /** Set when the kind of file is known: "ooxml" is a docx or an xlsx (which of the two only the password tells). */
  format?: EncryptedFormat;

  constructor(
    readonly code: PasswordErrorCode,
    message: string,
  ) {
    super(message);
  }
}

export const unsupportedEncryption = (what: string): PasswordError => new PasswordError('unsupported', `This kind of encryption is not supported: ${what}`);
export const damagedEncryption = (what: string): PasswordError => new PasswordError('damaged', `The encrypted file is damaged: ${what}`);
