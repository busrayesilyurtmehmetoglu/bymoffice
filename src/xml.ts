// Minimal XML parser/serializer for OOXML parts.
//
// Qualified names are kept exactly as written (prefixes are not resolved) and
// attribute order is preserved, so parse -> serialize keeps everything a
// consumer can observe. DTDs are rejected: OOXML forbids them and they are an
// attack surface.

export interface XmlElement {
  type: 'element';
  name: string;
  attrs: Map<string, string>;
  children: XmlNode[];
}

export interface XmlText {
  type: 'text';
  value: string;
  cdata?: boolean;
}

export interface XmlComment {
  type: 'comment';
  value: string;
}

export interface XmlPI {
  type: 'pi';
  target: string;
  value: string;
}

export type XmlNode = XmlElement | XmlText | XmlComment | XmlPI;

export interface XmlDocument {
  /** The `<?xml ...?>` declaration exactly as written, if present. */
  declaration: string | null;
  leading: XmlNode[];
  root: XmlElement;
  trailing: XmlNode[];
}

const NAMED_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

function decodeEntities(text: string): string {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z]+);/g, (match, ref: string) => {
    if (ref[0] === '#') {
      return String.fromCodePoint(ref[1] === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10));
    }
    // Own properties only: "&constructor;" must not find Object.prototype members.
    const value = Object.prototype.hasOwnProperty.call(NAMED_ENTITIES, ref) ? NAMED_ENTITIES[ref] : undefined;
    if (value === undefined) throw new Error(`Unknown XML entity: ${match}`);
    return value;
  });
}

// What the bytes 0x80 to 0x9F are in windows-1252; windows-1254 has two of them as controls.
const CP1252_HIGH = [
  0x20ac, 0x81, 0x201a, 0x192, 0x201e, 0x2026, 0x2020, 0x2021, 0x2c6, 0x2030, 0x160, 0x2039, 0x152, 0x8d, 0x17d, 0x8f, 0x90, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x2dc, 0x2122, 0x161, 0x203a, 0x153, 0x9d,
  0x17e, 0x178,
];
// The six letters Turkish has where Latin-1 has others.
const TURKISH: [byte: number, code: number][] = [
  [0xd0, 0x11e],
  [0xdd, 0x130],
  [0xde, 0x15e],
  [0xf0, 0x11f],
  [0xfd, 0x131],
  [0xfe, 0x15f],
];

const SINGLE_BYTE: { [name: string]: 'latin' | 'turkish' | 'ascii' } = {};
for (const name of ['iso-8859-1', 'iso8859-1', 'iso_8859-1', 'latin1', 'l1', 'cp1252', 'windows-1252', 'x-cp1252', 'cp819', 'ibm819']) SINGLE_BYTE[name] = 'latin';
for (const name of ['iso-8859-9', 'iso8859-9', 'iso_8859-9', 'latin5', 'l5', 'cp1254', 'windows-1254', 'x-cp1254']) SINGLE_BYTE[name] = 'turkish';
for (const name of ['us-ascii', 'ascii', 'ansi_x3.4-1968', 'iso646-us']) SINGLE_BYTE[name] = 'ascii';

const tables: { [kind: string]: Uint16Array } = {};

/** The characters the 256 bytes of a single-byte encoding stand for. As browsers do, ISO-8859-1 is read as windows-1252 and ISO-8859-9 as windows-1254. */
function tableOf(kind: 'latin' | 'turkish' | 'ascii'): Uint16Array {
  let table = tables[kind];
  if (table) return table;
  table = new Uint16Array(256);
  for (let i = 0; i < 256; i++) table[i] = kind === 'ascii' && i >= 0x80 ? 0xfffd : i;
  if (kind !== 'ascii') for (let i = 0; i < 32; i++) table[0x80 + i] = CP1252_HIGH[i];
  if (kind === 'turkish') {
    table[0x8e] = 0x8e;
    table[0x9e] = 0x9e;
    for (const [byte, code] of TURKISH) table[byte] = code;
  }
  tables[kind] = table;
  return table;
}

/**
 * The encoding the XML declaration at the start of a part names, in lower case; undefined for
 * a part without one, one that says nothing of its encoding, and one that is not written in
 * bytes of ASCII up to there (UTF-16).
 */
function declaredEncoding(bytes: Uint8Array): string | undefined {
  const start = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0;
  if (bytes[start] !== 60 || bytes[start + 1] !== 63) return undefined;
  let head = '';
  for (let i = start; i < bytes.length && i < start + 200; i++) {
    const code = bytes[i];
    head += String.fromCharCode(code);
    if (code === 62) break;
  }
  const match = /^<\?xml\s[^?>]*?\bencoding\s*=\s*["']([^"']*)["']/.exec(head);
  return match ? match[1].trim().toLowerCase() : undefined;
}

/** The single-byte encoding a part is read in; undefined for a part that is read as UTF-8 or UTF-16. A byte order mark wins over what is declared. */
function singleByteOf(bytes: Uint8Array): 'latin' | 'turkish' | 'ascii' | undefined {
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)) return undefined;
  const name = declaredEncoding(bytes);
  return name !== undefined && Object.prototype.hasOwnProperty.call(SINGLE_BYTE, name) ? SINGLE_BYTE[name] : undefined;
}

/**
 * Whether a part is in an encoding that is read here: UTF-8, UTF-16 with a byte order mark, or
 * one of the single-byte encodings decodeXmlBytes() knows. A part that declares another one is
 * taken as UTF-8 all the same (as it always was); what it holds beyond ASCII is then not what
 * was written, so nothing is made of its text.
 */
export function xmlEncodingRead(bytes: Uint8Array): boolean {
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf)) return true;
  const name = declaredEncoding(bytes);
  return name === undefined || /^utf-?8$/.test(name) || Object.prototype.hasOwnProperty.call(SINGLE_BYTE, name);
}

/**
 * A part as text: UTF-16 by its byte order mark; ISO-8859-1, windows-1252, ISO-8859-9,
 * windows-1254 and US-ASCII where the declaration of the part names one of them; UTF-8
 * otherwise. Line breaks stay as they are. A byte order mark is left out, unless `keepMark`.
 */
export function decodeXmlBytes(bytes: Uint8Array, keepMark = false): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le', { ignoreBOM: keepMark }).decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be', { ignoreBOM: keepMark }).decode(bytes);
  const kind = singleByteOf(bytes);
  if (kind === undefined) return new TextDecoder('utf-8', { ignoreBOM: keepMark }).decode(bytes);
  const table = tableOf(kind);
  const n = bytes.length;
  const parts: string[] = [];
  const chunk = new Uint16Array(Math.min(n, 8192));
  for (let at = 0; at < n; ) {
    const size = Math.min(n - at, chunk.length);
    for (let i = 0; i < size; i++) chunk[i] = table[bytes[at + i]];
    parts.push(String.fromCharCode.apply(null, (size === chunk.length ? chunk : chunk.subarray(0, size)) as unknown as number[]));
    at += size;
  }
  return parts.join('');
}

const decodeBytes = (bytes: Uint8Array): string => decodeXmlBytes(bytes);

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;
// space, '/', '>', '='
const isNameEnd = (c: number) => isSpace(c) || c === 47 || c === 62 || c === 61;

export function parseXml(input: string | Uint8Array): XmlDocument {
  let s = typeof input === 'string' ? input : decodeBytes(input);
  if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
  if (s.includes('\r')) s = s.replace(/\r\n?/g, '\n');

  const n = s.length;
  const fail = (message: string, at: number): never => {
    throw new Error(`Invalid XML at offset ${at}: ${message}`);
  };
  const find = (needle: string, from: number): number => {
    const j = s.indexOf(needle, from);
    return j < 0 ? fail(`missing "${needle}"`, from) : j;
  };

  let declaration: string | null = null;
  let root: XmlElement | undefined;
  const leading: XmlNode[] = [];
  const trailing: XmlNode[] = [];
  const stack: XmlElement[] = [];

  const push = (node: XmlNode, at: number) => {
    if (stack.length) stack[stack.length - 1].children.push(node);
    else if (node.type === 'text') {
      if (node.cdata || node.value.trim()) fail('text outside the root element', at);
    } else (root ? trailing : leading).push(node);
  };

  let i = 0;
  while (i < n) {
    if (s.charCodeAt(i) !== 60) {
      let j = s.indexOf('<', i);
      if (j < 0) j = n;
      push({ type: 'text', value: decodeEntities(s.slice(i, j)) }, i);
      i = j;
      continue;
    }

    const c = s.charCodeAt(i + 1);
    if (c === 47) {
      const j = find('>', i);
      const name = s.slice(i + 2, j).trimEnd();
      const open = stack.pop();
      if (!open || open.name !== name) fail(`unexpected closing tag </${name}>`, i);
      i = j + 1;
      continue;
    }
    if (c === 63) {
      const j = find('?>', i);
      const body = s.slice(i + 2, j);
      if (i === 0 && /^xml\s/.test(body)) declaration = s.slice(i, j + 2);
      else {
        const m = /^(\S+)\s*([\s\S]*)$/.exec(body);
        if (!m) fail('empty processing instruction', i);
        else push({ type: 'pi', target: m[1], value: m[2] }, i);
      }
      i = j + 2;
      continue;
    }
    if (c === 33) {
      if (s.startsWith('<!--', i)) {
        const j = find('-->', i + 4);
        push({ type: 'comment', value: s.slice(i + 4, j) }, i);
        i = j + 3;
        continue;
      }
      if (s.startsWith('<![CDATA[', i)) {
        const j = find(']]>', i + 9);
        push({ type: 'text', value: s.slice(i + 9, j), cdata: true }, i);
        i = j + 3;
        continue;
      }
      fail('DTD declarations are not allowed', i);
    }

    let j = i + 1;
    while (j < n && !isNameEnd(s.charCodeAt(j))) j++;
    const element: XmlElement = { type: 'element', name: s.slice(i + 1, j), attrs: new Map(), children: [] };
    if (!element.name) fail('missing element name', i);

    let selfClosing = false;
    for (;;) {
      while (j < n && isSpace(s.charCodeAt(j))) j++;
      if (j >= n) fail('unterminated start tag', i);
      const ch = s.charCodeAt(j);
      if (ch === 62) {
        j++;
        break;
      }
      if (ch === 47) {
        if (s.charCodeAt(j + 1) !== 62) fail('expected "/>"', j);
        selfClosing = true;
        j += 2;
        break;
      }
      let k = j;
      while (k < n && !isNameEnd(s.charCodeAt(k))) k++;
      const name = s.slice(j, k);
      if (!name) fail('missing attribute name', j);
      while (k < n && isSpace(s.charCodeAt(k))) k++;
      if (s.charCodeAt(k) !== 61) fail(`attribute "${name}" has no value`, k);
      k++;
      while (k < n && isSpace(s.charCodeAt(k))) k++;
      const quote = s[k];
      if (quote !== '"' && quote !== "'") fail(`attribute "${name}" value is not quoted`, k);
      const end = find(quote, k + 1);
      // Attribute-value normalization: literal tabs/newlines read as spaces.
      element.attrs.set(name, decodeEntities(s.slice(k + 1, end).replace(/[\t\n]/g, ' ')));
      j = end + 1;
    }

    if (stack.length) stack[stack.length - 1].children.push(element);
    else if (root) fail('more than one root element', i);
    else root = element;
    if (!selfClosing) stack.push(element);
    i = j;
  }

  if (stack.length) fail(`unclosed element <${stack[stack.length - 1].name}>`, n);
  if (!root) throw new Error('Invalid XML: no root element');
  return { declaration, leading, root, trailing };
}

const TEXT_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '\r': '&#13;' };
const ATTR_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '"': '&quot;',
  '\t': '&#9;',
  '\n': '&#10;',
  '\r': '&#13;',
};

const escapeText = (value: string) => value.replace(/[&<>\r]/g, (ch) => TEXT_ESCAPES[ch]);
const escapeAttr = (value: string) => value.replace(/[&<"\t\n\r]/g, (ch) => ATTR_ESCAPES[ch]);

function writeNode(node: XmlNode, out: string[]): void {
  switch (node.type) {
    case 'text':
      out.push(node.cdata && !node.value.includes(']]>') ? `<![CDATA[${node.value}]]>` : escapeText(node.value));
      return;
    case 'comment':
      out.push('<!--', node.value, '-->');
      return;
    case 'pi':
      out.push('<?', node.target, node.value ? ` ${node.value}` : '', '?>');
      return;
    case 'element':
      out.push('<', node.name);
      for (const [name, value] of node.attrs) out.push(' ', name, '="', escapeAttr(value), '"');
      if (node.children.length === 0) {
        out.push('/>');
        return;
      }
      out.push('>');
      for (const child of node.children) writeNode(child, out);
      out.push('</', node.name, '>');
  }
}

/**
 * The declaration as it has to be written: what serializeXml returns is encoded as UTF-8 by
 * every caller, so a declaration that names another encoding (the part was read from UTF-16)
 * names UTF-8. Everything else in it stays as written; a declaration without an encoding gets none.
 */
function utf8Declaration(declaration: string): string {
  return declaration.replace(/(\sencoding\s*=\s*)(["'])([^"']*)\2/, (all, head: string, quote: string, name: string) =>
    /^utf-?8$/i.test(name) ? all : `${head}${quote}UTF-8${quote}`,
  );
}

export function serializeXml(doc: XmlDocument): string {
  const out: string[] = [];
  if (doc.declaration) out.push(utf8Declaration(doc.declaration), '\n');
  for (const node of doc.leading) writeNode(node, out);
  writeNode(doc.root, out);
  for (const node of doc.trailing) writeNode(node, out);
  return out.join('');
}

/** Name without its namespace prefix. */
export const localName = (name: string) => name.slice(name.indexOf(':') + 1);

export function childElements(parent: XmlElement, local?: string): XmlElement[] {
  return parent.children.filter(
    (node): node is XmlElement => node.type === 'element' && (local === undefined || localName(node.name) === local),
  );
}
