// Links and the outline of a PDF: what a page of the file holds besides what
// is painted on it (a link annotation over the text of a hyperlink), and the
// bookmarks a reader shows beside the pages (the headings of a document, the
// sheets of a workbook).
//
// Nothing that comes from a file is written as an address before it was
// looked at here: only http and https addresses with a host and mailto:
// become links that leave the file (pdfLinkUri); a link into the file itself
// points at a page and a place on it. Everything else is left out and
// counted, so that whoever made the file can say what was not carried over.

/** A rectangle on a page, in points from the top left corner of the page. */
export interface PdfLinkRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** A place in the file: a page (its index, from 0) and a point on it, in points from the top left corner. */
export interface PdfDestination {
  page: number;
  x?: number;
  y?: number;
}

/** What a link of a page goes to, and where on the page it can be pressed. */
export interface PdfLink {
  /** Where the link lies on its page: one rectangle per piece (a link broken over two lines has two). */
  rects: readonly PdfLinkRect[];
  /** An address outside the file; it is checked again (pdfLinkUri) and the link is left out when it does not pass. */
  uri?: string;
  /** A place in the file. */
  dest?: PdfDestination;
}

/** One entry of the outline: `level` 0 is at the top, every deeper level is under the entry before it that is less deep. */
export interface PdfOutlineEntry {
  title: string;
  level: number;
  dest: PdfDestination;
}

/**
 * What became of the links and the outline of a file. `external` and
 * `internal` count the links that were written (one per link, however many
 * pieces it has); `refused` the links whose address is of a kind a PDF does
 * not get (javascript:, file:, data:, tel:, relative ones); `unresolved` the
 * links into the file whose target was not found (a bookmark that is not
 * there, a cell that is not printed); `dropped` what was left out beyond the
 * limits (links of a page, of the file, entries of the outline); `rotated`
 * the links in turned text (a cell of a table with vertical text), which get
 * no link; `outlineSkipped` the headings that are not in the outline because
 * their title is empty (or hidden text only) or they are on no page;
 * `unread` the blocks (paragraphs, tables) of a document that were not looked
 * at for links, bookmarks and headings because the document is beyond the
 * limit of what is walked.
 */
export interface PdfLinkReport {
  external: number;
  internal: number;
  refused: number;
  unresolved: number;
  dropped: number;
  rotated: number;
  /** Entries of the outline that were written. */
  outline: number;
  outlineSkipped: number;
  unread: number;
}

export const emptyLinkReport = (): PdfLinkReport => ({ external: 0, internal: 0, refused: 0, unresolved: 0, dropped: 0, rotated: 0, outline: 0, outlineSkipped: 0, unread: 0 });

/** Links of one page; a page with more keeps the first ones. */
export const MAX_PAGE_LINKS = 2000;
/** Links of the whole file. */
export const MAX_FILE_LINKS = 50000;
/** Rectangles of one link: what a link of a few lines needs, not one that runs over a page of pieces. */
export const MAX_LINK_RECTS = 64;
/** Entries of the outline. */
export const MAX_OUTLINE_ENTRIES = 10000;
/** Levels of the outline: Word has nine levels of headings. */
export const MAX_OUTLINE_LEVELS = 9;
/** Characters of the title of an entry of the outline. */
export const MAX_OUTLINE_TITLE = 256;
/** Bytes of an address, once it is written as ASCII. */
export const MAX_URI_BYTES = 8192;

const HEX = '0123456789ABCDEF';

/**
 * The address a link of a PDF may go to: an http or https address with a
 * host, or a mailto: address, as ASCII - every byte of the UTF-8 of a
 * character beyond ASCII, of a space, of a control and of the characters an
 * address does not hold as they are ("<", ">", '"', "\", "^", "`", "{", "|",
 * "}") written as %XX, a "%" that starts no %XX too. In an http or https
 * address a backslash before the query is a slash, as readers of addresses
 * take it, and the host is what they find in it (after the user, before the
 * port; see hostOf). Undefined for anything else: javascript:, file:, data:,
 * tel:, an address without a scheme (it would be read relative to wherever
 * the file lies), one without a host or with a host that holds what a host
 * cannot, one of more than MAX_URI_BYTES. Spaces and controls around the address and tabs
 * and line ends in it are not part of it (as a browser reads it).
 */
export function pdfLinkUri(target: string | undefined): string | undefined {
  const read = pdfLinkTarget(target);
  return read ? read.uri : undefined;
}

/** Characters a host may hold as they are; anything beyond ASCII is a letter of a name the reader turns into ASCII (IDNA). */
const HOST_CHAR = /^[A-Za-z0-9.\-_~]$/;

/**
 * The host of an http or https address after "//" (backslashes before the
 * query already read as slashes), as a reader of addresses (the URL
 * standard) finds it: the authority runs to the first "/", "?" or "#"; what
 * is before its last "@" is the user; a port after the last ":" (outside
 * "[...]") is digits or nothing. Undefined for an authority with no host, or
 * whose host holds a character a host cannot (it would be read as another
 * host, or not at all).
 */
function hostOf(rest: string): string | undefined {
  let end = 2;
  while (end < rest.length && '/?#'.indexOf(rest.charAt(end)) < 0) end++;
  const authority = rest.slice(2, end);
  let host = authority.slice(authority.lastIndexOf('@') + 1);
  // A port is digits, at most 65535; nothing after the colon is no port.
  const portOk = (port: string): boolean => port === '' || (/^\d{1,16}$/.test(port) && Number(port) <= 65535);
  if (host.charAt(0) === '[') {
    const close = host.indexOf(']');
    if (close < 0) return undefined;
    const port = host.slice(close + 1);
    if (port !== '' && (port.charAt(0) !== ':' || !portOk(port.slice(1)))) return undefined;
    host = host.slice(0, close + 1);
    return isIpv6(host.slice(1, -1)) ? host : undefined;
  }
  const colon = host.lastIndexOf(':');
  if (colon >= 0) {
    if (!portOk(host.slice(colon + 1))) return undefined;
    host = host.slice(0, colon);
  }
  if (host === '') return undefined;
  let wide = false;
  for (let i = 0; i < host.length; i++) {
    if (host.charCodeAt(i) >= 0x80) wide = true;
    else if (!HOST_CHAR.test(host.charAt(i))) return undefined;
  }
  return domainOk(host, wide) ? host : undefined;
}

/** A letter, a mark or a decimal digit beyond ASCII; undefined where the runtime has no Unicode properties in its regular expressions. */
const WIDE_LETTER = ((): RegExp | undefined => {
  try {
    return new RegExp('^[\\p{L}\\p{M}\\p{Nd}]$', 'u');
  } catch {
    return undefined;
  }
})();
const MARK = ((): RegExp | undefined => {
  try {
    return new RegExp('^\\p{M}', 'u');
  } catch {
    return undefined;
  }
})();

/**
 * Whether a reader of addresses reads the host as the name it is: a name of
 * letters, digits and "-._~" (beyond ASCII only letters, marks and decimal
 * digits, which the reader turns into ASCII: what they are turned into must
 * be the same), whose labels written as punycode ("xn--") are punycode of
 * such letters, and that does not end in a number unless it is an IPv4
 * address of four decimal numbers. Anything else the reader refuses or reads
 * as an address of another kind: such a host is not written.
 */
function domainOk(host: string, wide: boolean): boolean {
  let name = host;
  if (wide) {
    // Beyond ASCII without a way to tell letters, nothing is taken.
    if (!WIDE_LETTER || !MARK || typeof name.normalize !== 'function') return false;
    for (const char of name) if (char.charCodeAt(0) >= 0x80 && (!WIDE_LETTER.test(char) || rightToLeft(char.codePointAt(0) as number))) return false;
    // As the reader maps it (compatibility forms, case): what comes out must be letters still (the full-width at sign is "@", a digit with a full stop is "1.").
    name = name.normalize('NFKC').toLowerCase();
    for (const char of name) if (char.charCodeAt(0) < 0x80 ? !HOST_CHAR.test(char) : !WIDE_LETTER.test(char)) return false;
  }
  const labels = name.split('.');
  // The last label decides whether the host is a number (an IPv4 address), as the reader decides it; a dot at the end is no label.
  const last = labels[labels.length - 1] === '' && labels.length > 1 ? labels[labels.length - 2] : labels[labels.length - 1];
  if (/^(\d+|0[xX][0-9A-Fa-f]*)$/.test(last)) return /^(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}$/.test(host);
  for (let i = 0; i < labels.length; i++) {
    const label = labels[i].toLowerCase();
    if (wide && MARK && MARK.test(label)) return false;
    if (label.length >= 4 && label.slice(0, 4).toLowerCase() === 'xn--') {
      const decoded = punycodeDecode(label.slice(4));
      if (decoded === undefined || !/[^\x00-\x7f]/.test(decoded) || !WIDE_LETTER || !MARK || MARK.test(decoded)) return false;
      for (const char of decoded) if (char.charCodeAt(0) < 0x80 ? !/^[a-z0-9-]$/.test(char) : !WIDE_LETTER.test(char) || rightToLeft(char.codePointAt(0) as number)) return false;
      if (typeof decoded.normalize !== 'function' || decoded.normalize('NFKC').toLowerCase() !== decoded) return false;
    }
  }
  return true;
}

/**
 * A letter of a script written from right to left (Hebrew, Arabic, Syriac,
 * Thaana, NKo and their kin, with their presentation forms). A name that has
 * one is under the rules of direction of names (RFC 5893), which readers
 * apply differently: such a host is not written at all.
 */
function rightToLeft(code: number): boolean {
  return (code >= 0x0590 && code <= 0x08ff) || (code >= 0xfb1d && code <= 0xfdff) || (code >= 0xfe70 && code <= 0xfeff) || (code >= 0x10800 && code <= 0x10fff) || (code >= 0x1e800 && code <= 0x1efff);
}

/** The characters punycode (RFC 3492) stands for; undefined for text that is no punycode. */
function punycodeDecode(input: string): string | undefined {
  const BASE = 36;
  const T_MIN = 1;
  const T_MAX = 26;
  if (input.length === 0 || input.length > 63) return undefined;
  const delimiter = input.lastIndexOf('-');
  const output: number[] = [];
  for (let i = 0; i < Math.max(delimiter, 0); i++) {
    const code = input.charCodeAt(i);
    if (code >= 0x80) return undefined;
    output.push(code);
  }
  let n = 128;
  let bias = 72;
  let i = 0;
  let first = true;
  for (let at = delimiter > 0 ? delimiter + 1 : 0; at < input.length; ) {
    const old = i;
    let w = 1;
    for (let k = BASE; ; k += BASE) {
      if (at >= input.length) return undefined;
      const c = input.charCodeAt(at++);
      const digit = c >= 48 && c <= 57 ? c - 22 : c >= 65 && c <= 90 ? c - 65 : c >= 97 && c <= 122 ? c - 97 : -1;
      if (digit < 0) return undefined;
      i += digit * w;
      if (i > 0x10ffff * 64) return undefined;
      const t = k <= bias ? T_MIN : k >= bias + T_MAX ? T_MAX : k - bias;
      if (digit < t) break;
      w *= BASE - t;
      if (w > 0x10ffff * 64) return undefined;
    }
    const length = output.length + 1;
    // Adapting the bias (RFC 3492, 6.1).
    let delta = first ? Math.floor((i - old) / 700) : Math.floor((i - old) / 2);
    first = false;
    delta += Math.floor(delta / length);
    let k = 0;
    while (delta > ((BASE - T_MIN) * T_MAX) >> 1) {
      delta = Math.floor(delta / (BASE - T_MIN));
      k += BASE;
    }
    bias = k + Math.floor(((BASE - T_MIN + 1) * delta) / (delta + 38));
    n += Math.floor(i / length);
    i %= length;
    if (n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) return undefined;
    output.splice(i, 0, n);
    i++;
  }
  let out = '';
  for (let k = 0; k < output.length; k++) out += String.fromCodePoint(output[k]);
  return out;
}

/** An IPv6 address as the reader reads one in "[...]": eight groups of up to four hex digits, one "::" for any run of zero groups, the last two groups possibly a dotted IPv4 address. */
function isIpv6(text: string): boolean {
  if (text === '' || text.length > 45) return false;
  const double = text.indexOf('::');
  if (double >= 0 && text.indexOf('::', double + 1) >= 0) return false;
  const parts = (side: string): string[] | undefined => (side === '' ? [] : side.split(':'));
  const head = double >= 0 ? parts(text.slice(0, double)) : parts(text);
  const tail = double >= 0 ? parts(text.slice(double + 2)) : [];
  if (!head || !tail) return false;
  const groups = head.concat(tail);
  let count = 0;
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    if (/^[0-9A-Fa-f]{1,4}$/.test(group)) count++;
    else if (g === groups.length - 1 && /^(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)(\.(25[0-5]|2[0-4]\d|1\d\d|[1-9]?\d)){3}$/.test(group)) count += 2;
    else return false;
  }
  return double >= 0 ? count <= 7 : count === 8;
}

/**
 * As pdfLinkUri(), with the host the address was checked for (http and
 * https): a reader of the written address finds that host in it.
 */
export function pdfLinkTarget(target: string | undefined): { uri: string; host?: string } | undefined {
  if (typeof target !== 'string') return undefined;
  let start = 0;
  let end = target.length;
  while (start < end && target.charCodeAt(start) <= 0x20) start++;
  while (end > start && target.charCodeAt(end - 1) <= 0x20) end--;
  // An address longer than this cannot be one that is written: it is not read on.
  if (end - start > MAX_URI_BYTES * 2) return undefined;
  let url = '';
  for (let i = start; i < end; i++) {
    const c = target.charCodeAt(i);
    if (c !== 9 && c !== 10 && c !== 13) url += target.charAt(i);
  }
  const colon = url.indexOf(':');
  if (colon <= 0 || colon > 6) return undefined;
  const scheme = url.slice(0, colon).toLowerCase();
  let rest = url.slice(colon + 1);
  let host: string | undefined;
  if (scheme === 'http' || scheme === 'https') {
    // Readers of addresses take a backslash before the query for a slash (the URL standard): so does this, before the
    // host is looked for, so that the host that is checked is the one that is read ("http://a.com\@b.com/" goes to a.com).
    const query = rest.search(/[?#]/);
    const head = query < 0 ? rest : rest.slice(0, query);
    rest = head.replace(/\\/g, '/') + (query < 0 ? '' : rest.slice(query));
    // A host is needed: "https:path" and "https:/path" are relative to whatever shows the link; "https:///x" is
    // read with "x" for its host by some readers and not by others.
    if (rest.slice(0, 2) !== '//' || rest.charAt(2) === '/') return undefined;
    host = hostOf(rest);
    if (host === undefined) return undefined;
  } else if (scheme === 'mailto') {
    if (rest === '') return undefined;
  } else return undefined;
  let out = scheme + ':';
  for (let i = 0; i < rest.length; i++) {
    let code = rest.charCodeAt(i);
    const char = rest.charAt(i);
    if (code > 0x20 && code < 0x7f && '<>"\\^`{|}'.indexOf(char) < 0) {
      // A "%" that is not the start of %XX is written as the character it is.
      if (char === '%' && !/^[0-9A-Fa-f]{2}$/.test(rest.substr(i + 1, 2))) out += '%25';
      else out += char;
    } else {
      if (code >= 0xd800 && code <= 0xdfff) {
        const next = i + 1 < rest.length ? rest.charCodeAt(i + 1) : 0;
        if (code <= 0xdbff && next >= 0xdc00 && next <= 0xdfff) {
          code = 0x10000 + ((code - 0xd800) << 10) + (next - 0xdc00);
          i++;
        } else code = 0xfffd;
      }
      const bytes = code < 0x80 ? [code] : code < 0x800 ? [0xc0 | (code >> 6), 0x80 | (code & 63)] : code < 0x10000 ? [0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63)] : [0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63)];
      for (let b = 0; b < bytes.length; b++) out += '%' + HEX.charAt(bytes[b] >> 4) + HEX.charAt(bytes[b] & 15);
    }
    if (out.length > MAX_URI_BYTES) return undefined;
  }
  return host === undefined ? { uri: out } : { uri: out, host };
}

/**
 * The title of an entry of the outline: white space of any kind as one
 * space, controls left out, at most MAX_OUTLINE_TITLE characters (never cut
 * inside a pair of surrogates). Empty when nothing is left.
 */
export function outlineTitle(text: string): string {
  if (typeof text !== 'string') return '';
  let out = '';
  let space = false;
  for (let i = 0; i < text.length && out.length < MAX_OUTLINE_TITLE; i++) {
    const code = text.charCodeAt(i);
    // Controls are left out; those that separate words (tab, line end) are white space.
    if (code < 0x20 && (code < 9 || code > 13)) continue;
    if (code <= 0x20 || code === 0xa0 || (code >= 0x2000 && code <= 0x200a) || code === 0x2028 || code === 0x2029 || code === 0x3000) {
      space = out.length > 0;
      continue;
    }
    if ((code >= 0x7f && code < 0xa0) || code === 0xfeff || code === 0xad) continue;
    if (space) {
      out += ' ';
      space = false;
    }
    if (code >= 0xd800 && code <= 0xdbff) {
      if (out.length + 2 > MAX_OUTLINE_TITLE) break;
      out += text.substr(i, 2);
      i++;
      continue;
    }
    out += text.charAt(i);
  }
  return out;
}

/** One node of the outline as it is written: the entry, its parent and its children (indexes into the list). */
export interface PdfOutlineNode {
  entry: PdfOutlineEntry;
  parent: number;
  children: number[];
  /** Its place among the children of its parent (among the entries at the top for those). */
  index: number;
  /** Entries under it, at every depth. */
  descendants: number;
}

const levelOf = (entry: PdfOutlineEntry): number => Math.min(MAX_OUTLINE_LEVELS - 1, Math.max(0, Math.floor(entry.level) || 0));

/**
 * The outline as a tree: every entry under the nearest entry before it with
 * a lower level (a level that skips one, as Heading 1 then Heading 3, is
 * simply one deeper). Parents are -1 at the top. Made without recursion.
 */
export function outlineTree(entries: readonly PdfOutlineEntry[]): { nodes: PdfOutlineNode[]; top: number[] } {
  const nodes: PdfOutlineNode[] = [];
  const top: number[] = [];
  const stack: number[] = [];
  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const level = levelOf(entry);
    while (stack.length > 0 && levelOf(nodes[stack[stack.length - 1]].entry) >= level) stack.pop();
    const parent = stack.length > 0 ? stack[stack.length - 1] : -1;
    const index = nodes.length;
    const siblings = parent >= 0 ? nodes[parent].children : top;
    nodes.push({ entry, parent, children: [], index: siblings.length, descendants: 0 });
    siblings.push(index);
    for (let k = 0; k < stack.length; k++) nodes[stack[k]].descendants++;
    stack.push(index);
  }
  return { nodes, top };
}

/** True for a number that is a place on a page. */
export const finitePlace = (value: unknown): value is number => typeof value === 'number' && isFinite(value) && Math.abs(value) < 1e6;
