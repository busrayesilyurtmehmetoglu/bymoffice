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

function decodeBytes(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  return new TextDecoder('utf-8').decode(bytes);
}

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
