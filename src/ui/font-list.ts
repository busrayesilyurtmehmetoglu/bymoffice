// The names a list of fonts of a toolbar offers: the common fonts of Office, the
// fonts the application registered, the names the font of the package stands
// for, the fonts the file itself names, and the font of the selection. Names
// come from the file: they only ever reach the DOM as text.

import { registeredFonts } from '../render/fonts';
import { embeddedFontNames } from '../render/embedded-fonts';

/** Fonts offered besides those the file names. */
export const COMMON_FONTS = ['Arial', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Tahoma', 'Times New Roman', 'Verdana'];
/** Most names a list holds. */
export const MAX_FONT_ENTRIES = 120;
/** Longer names, and names with control characters, are not offered (the font of the selection is still shown). */
const MAX_NAME = 128;

const usable = (name: string): boolean => name !== '' && name.length <= MAX_NAME && !/[\u0000-\u001f\u007f]/.test(name);

/**
 * The names in alphabetical order (without regard to case), each once, and the entry of
 * `current` among them (undefined when there is none). `fromFile` are the names the file uses.
 */
export function fontChoices(fromFile: Iterable<string>, current: string | undefined): { names: string[]; shown: string | undefined } {
  const names = new Map<string, string>();
  const add = (name: string, always = false) => {
    if ((always ? name !== '' : usable(name)) && (names.size < MAX_FONT_ENTRIES || always) && !names.has(name.toLowerCase())) names.set(name.toLowerCase(), name);
  };
  for (let i = 0; i < COMMON_FONTS.length; i++) add(COMMON_FONTS[i]);
  // The fonts the application registered: they are there for certain.
  const registered = registeredFonts();
  for (let i = 0; i < registered.length; i++) if (registered[i].status === 'loaded') add(registered[i].family);
  // The names the font of the package is shown for (Bookman Old Style): the file gets that name, and where
  // the device lacks the font the text is drawn with the one of the package.
  const embedded = embeddedFontNames();
  for (let i = 0; i < embedded.length; i++) add(embedded[i]);
  for (const name of fromFile) add(name);
  if (current !== undefined) add(current, true);
  const list: string[] = [];
  names.forEach((name) => list.push(name));
  list.sort((a, b) => (a.toLowerCase() < b.toLowerCase() ? -1 : 1));
  return { names: list, shown: current === undefined ? undefined : names.get(current.toLowerCase()) };
}
