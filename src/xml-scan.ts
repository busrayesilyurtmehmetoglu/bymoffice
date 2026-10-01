// Pull scanner for large XML parts (sheet data, shared strings).
//
// Reads the same language as parseXml in xml.ts, with the same entity
// decoding, attribute-value normalization and errors, but hands out one token
// at a time instead of building a tree. Nothing is allocated per token apart
// from the strings a consumer asks for: attributes are only validated while
// scanning and are decoded on demand.

const NAMED_ENTITIES: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

function decodeEntities(text: string): string {
  if (text.indexOf('&') < 0) return text;
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[A-Za-z]+);/g, (match, ref: string) => {
    if (ref[0] === '#') {
      const code = ref[1] === 'x' ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
      if (!(code <= 0x10ffff)) throw new Error(`Invalid XML character reference: ${match}`);
      return String.fromCodePoint(code);
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

/** Attribute-value normalization (literal tabs/newlines read as spaces), then entities. */
function decodeAttrValue(raw: string): string {
  if (raw.indexOf('\n') >= 0 || raw.indexOf('\t') >= 0) raw = raw.replace(/[\t\n]/g, ' ');
  return decodeEntities(raw);
}

export class XmlScanner {
  kind: 'open' | 'close' | 'text' = 'text';
  /** Qualified name as written (open/close). */
  name = '';
  /** Name without prefix (open/close). */
  local = '';
  /** Set on open tokens. A self-closing element produces no separate close token. */
  selfClosing = false;
  /**
   * Number of attributes of the current open tag. Lets a consumer that knows
   * how many of them it has found skip the lookups for rare ones.
   */
  attrCount = 0;
  /** Decoded text of a text token (CDATA included). */
  text = '';

  private readonly s: string;
  private readonly n: number;
  private pos = 0;
  /** Offset of the "<" of the current open tag. */
  private tagStart = 0;
  /** Offset just past the name of the current open tag. */
  private attrStart = 0;
  /** Open elements, innermost last. */
  private names: string[] = [];
  private locals: string[] = [];
  private depth = 0;
  /** Last element name seen at each depth; siblings usually repeat it, which saves a slice per tag. */
  private nameCache: string[] = [];
  private localCache: string[] = [];
  private rootSeen = false;

  constructor(input: string | Uint8Array) {
    let s = typeof input === 'string' ? input : decodeBytes(input);
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
    if (s.indexOf('\r') >= 0) s = s.replace(/\r\n?/g, '\n');
    this.s = s;
    this.n = s.length;
  }

  private fail(message: string, at: number): never {
    throw new Error(`Invalid XML at offset ${at}: ${message}`);
  }

  private find(needle: string, from: number): number {
    const j = this.s.indexOf(needle, from);
    return j < 0 ? this.fail(`missing "${needle}"`, from) : j;
  }

  /** Advances to the next token; false at the end of input. */
  next(): boolean {
    const s = this.s;
    const n = this.n;
    let i = this.pos;
    for (;;) {
      if (i >= n) {
        this.pos = n;
        if (this.depth) this.fail(`unclosed element <${this.names[this.depth - 1]}>`, n);
        if (!this.rootSeen) throw new Error('Invalid XML: no root element');
        return false;
      }

      if (s.charCodeAt(i) !== 60) {
        let j = s.indexOf('<', i);
        if (j < 0) j = n;
        if (this.depth === 0) {
          if (decodeEntities(s.slice(i, j)).trim()) this.fail('text outside the root element', i);
          i = j;
          continue;
        }
        this.kind = 'text';
        this.text = decodeEntities(s.slice(i, j));
        this.pos = j;
        return true;
      }

      const c = s.charCodeAt(i + 1);
      if (c === 47) {
        const j = this.find('>', i);
        const depth = this.depth;
        if (depth === 0) this.fail(`unexpected closing tag </${s.slice(i + 2, j).trimEnd()}>`, i);
        const open = this.names[depth - 1];
        // Usual case: "</name>" with nothing after the name.
        if (!(j === i + 2 + open.length && s.startsWith(open, i + 2))) {
          const name = s.slice(i + 2, j).trimEnd();
          if (name !== open) this.fail(`unexpected closing tag </${name}>`, i);
        }
        this.kind = 'close';
        this.name = open;
        this.local = this.locals[depth - 1];
        this.depth = depth - 1;
        this.names.length = depth - 1;
        this.locals.length = depth - 1;
        this.pos = j + 1;
        return true;
      }
      if (c === 63) {
        const j = this.find('?>', i);
        // The XML declaration and processing instructions are skipped; a target is still required.
        if (j <= i + 2 || /\s/.test(s[i + 2])) this.fail('empty processing instruction', i);
        i = j + 2;
        continue;
      }
      if (c === 33) {
        if (s.startsWith('<!--', i)) {
          i = this.find('-->', i + 4) + 3;
          continue;
        }
        if (s.startsWith('<![CDATA[', i)) {
          const j = this.find(']]>', i + 9);
          if (this.depth === 0) this.fail('text outside the root element', i);
          this.kind = 'text';
          this.text = s.slice(i + 9, j);
          this.pos = j + 3;
          return true;
        }
        this.fail('DTD declarations are not allowed', i);
      }

      // Start tag. Attributes are checked here and decoded only when asked for.
      const depth = this.depth;
      let name = this.nameCache[depth];
      let local: string;
      let j: number;
      if (name !== undefined && s.startsWith(name, i + 1) && isNameEnd(s.charCodeAt(i + 1 + name.length))) {
        j = i + 1 + name.length;
        local = this.localCache[depth];
      } else {
        j = i + 1;
        while (j < n && !isNameEnd(s.charCodeAt(j))) j++;
        if (j === i + 1) this.fail('missing element name', i);
        name = s.slice(i + 1, j);
        local = name.slice(name.indexOf(':') + 1);
        this.nameCache[depth] = name;
        this.localCache[depth] = local;
      }
      const attrStart = j;

      let selfClosing = false;
      let attrCount = 0;
      for (;;) {
        while (j < n && isSpace(s.charCodeAt(j))) j++;
        if (j >= n) this.fail('unterminated start tag', i);
        const ch = s.charCodeAt(j);
        if (ch === 62) {
          j++;
          break;
        }
        if (ch === 47) {
          if (s.charCodeAt(j + 1) !== 62) this.fail('expected "/>"', j);
          selfClosing = true;
          j += 2;
          break;
        }
        let k = j;
        while (k < n && !isNameEnd(s.charCodeAt(k))) k++;
        if (k === j) this.fail('missing attribute name', j);
        const nameEnd = k;
        while (k < n && isSpace(s.charCodeAt(k))) k++;
        if (s.charCodeAt(k) !== 61) this.fail(`attribute "${s.slice(j, nameEnd)}" has no value`, k);
        k++;
        while (k < n && isSpace(s.charCodeAt(k))) k++;
        const quote = s.charCodeAt(k);
        if (quote !== 34 && quote !== 39) this.fail(`attribute "${s.slice(j, nameEnd)}" value is not quoted`, k);
        j = this.find(quote === 34 ? '"' : "'", k + 1) + 1;
        attrCount++;
      }

      if (depth === 0) {
        if (this.rootSeen) this.fail('more than one root element', i);
        this.rootSeen = true;
      }
      if (!selfClosing) {
        this.names.push(name);
        this.locals.push(local);
        this.depth = depth + 1;
      }
      this.kind = 'open';
      this.name = name;
      this.local = local;
      this.selfClosing = selfClosing;
      this.attrCount = attrCount;
      this.tagStart = i;
      this.attrStart = attrStart;
      this.pos = j;
      return true;
    }
  }

  /**
   * Attribute of the current open tag by qualified name, entity-decoded.
   * Undefined when absent or when the current token is not an open tag. If the
   * name is repeated the last one wins, as in parseXml.
   */
  attr(name: string): string | undefined {
    if (this.kind !== 'open') return undefined;
    const s = this.s;
    const length = name.length;
    let j = this.attrStart;
    let found: string | undefined;
    // The tag was validated by next(), so no bounds or syntax checks here.
    for (;;) {
      while (isSpace(s.charCodeAt(j))) j++;
      const ch = s.charCodeAt(j);
      if (ch === 62 || ch === 47) return found;
      let k = j;
      while (!isNameEnd(s.charCodeAt(k))) k++;
      const match = k - j === length && s.startsWith(name, j);
      while (s.charCodeAt(k) !== 61) k++;
      k++;
      while (isSpace(s.charCodeAt(k))) k++;
      const end = s.indexOf(s.charCodeAt(k) === 34 ? '"' : "'", k + 1);
      if (match) found = decodeAttrValue(s.slice(k + 1, end));
      j = end + 1;
    }
  }

  /** All attributes of the current open tag in document order (empty for other tokens). */
  attrs(): [string, string][] {
    const out: [string, string][] = [];
    if (this.kind !== 'open') return out;
    const s = this.s;
    let j = this.attrStart;
    for (;;) {
      while (isSpace(s.charCodeAt(j))) j++;
      const ch = s.charCodeAt(j);
      if (ch === 62 || ch === 47) return out;
      let k = j;
      while (!isNameEnd(s.charCodeAt(k))) k++;
      const name = s.slice(j, k);
      while (s.charCodeAt(k) !== 61) k++;
      k++;
      while (isSpace(s.charCodeAt(k))) k++;
      const end = s.indexOf(s.charCodeAt(k) === 34 ? '"' : "'", k + 1);
      out.push([name, decodeAttrValue(s.slice(k + 1, end))]);
      j = end + 1;
    }
  }

  private requireOpen(method: string): void {
    if (this.kind !== 'open') throw new Error(`XmlScanner.${method}() needs an open tag as the current token`);
  }

  /**
   * With an open tag current: skips to its matching close tag, which becomes
   * the current token (no-op when self-closing). The skipped content is still
   * checked for well-formedness.
   */
  skip(): void {
    this.requireOpen('skip');
    if (this.selfClosing) return;
    const target = this.depth - 1;
    while (this.next()) {
      if (this.kind === 'close' && this.depth === target) return;
    }
  }

  /**
   * With an open tag current: returns the concatenated text content of the
   * element (descendants included) and consumes it; the next call to next()
   * returns what follows the element.
   */
  readText(): string {
    this.requireOpen('readText');
    if (this.selfClosing) return '';
    const target = this.depth - 1;
    let out = '';
    while (this.next()) {
      if (this.kind === 'text') out = out === '' ? this.text : out + this.text;
      else if (this.kind === 'close' && this.depth === target) break;
    }
    return out;
  }

  /**
   * With an open tag current: returns the element's source text (start tag
   * through end tag) and consumes it like readText(). The text is exact except
   * that line endings were normalized to "\n" when the input was loaded.
   */
  readRaw(): string {
    this.requireOpen('readRaw');
    const start = this.tagStart;
    this.skip();
    return this.s.slice(start, this.pos);
  }
}
