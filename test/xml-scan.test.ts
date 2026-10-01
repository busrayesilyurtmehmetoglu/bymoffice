import assert from 'node:assert/strict';
import { test } from 'node:test';
import { parseXml, type XmlElement } from '../src/xml';
import { XmlScanner } from '../src/xml-scan';

// --- helpers ---------------------------------------------------------------

interface Tree {
  name: string;
  attrs: [string, string][];
  children: (Tree | string)[];
}

function pushText(children: (Tree | string)[], text: string): void {
  const last = children.length - 1;
  if (typeof children[last] === 'string') children[last] = (children[last] as string) + text;
  else children.push(text);
}

/** Tree built from scanner tokens; adjacent text tokens are joined. */
function scanTree(input: string | Uint8Array): Tree {
  const sc = new XmlScanner(input);
  let root: Tree | undefined;
  const stack: Tree[] = [];
  while (sc.next()) {
    if (sc.kind === 'open') {
      const node: Tree = { name: sc.name, attrs: [...new Map(sc.attrs())], children: [] };
      assert.equal(sc.local, sc.name.slice(sc.name.indexOf(':') + 1));
      for (const [name, value] of node.attrs) assert.equal(sc.attr(name), value);
      if (stack.length) stack[stack.length - 1].children.push(node);
      else root = node;
      if (!sc.selfClosing) stack.push(node);
    } else if (sc.kind === 'close') {
      const open = stack.pop();
      assert.equal(sc.name, open?.name);
    } else pushText(stack[stack.length - 1].children, sc.text);
  }
  assert.equal(sc.next(), false, 'next() stays false at the end');
  assert.ok(root);
  return root;
}

/** The parseXml tree in the same shape, without comments and PIs. */
function simplify(element: XmlElement): Tree {
  const out: Tree = { name: element.name, attrs: [...element.attrs], children: [] };
  // Iterative: generated documents can nest deeply.
  const work: [XmlElement, Tree][] = [[element, out]];
  while (work.length) {
    const [source, target] = work.pop()!;
    for (const child of source.children) {
      if (child.type === 'text') pushText(target.children, child.value);
      else if (child.type === 'element') {
        const node: Tree = { name: child.name, attrs: [...child.attrs], children: [] };
        target.children.push(node);
        work.push([child, node]);
      }
    }
  }
  return out;
}

function assertSameAsParseXml(input: string | Uint8Array): void {
  assert.deepEqual(scanTree(input), simplify(parseXml(input).root));
}

/** Deterministic PRNG (mulberry32). */
function prng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomDocument(seed: number): string {
  const rand = prng(seed);
  const pick = <T>(items: T[]): T => items[Math.floor(rand() * items.length)];
  const names = ['a', 'row', 'c', 'x:v', 'ns1:öğe', 'T', 'is'];
  const texts = ['', 'düz metin', ' ', '&amp;&lt;&gt;', '&#x15F;&#305;&#128512;', 'a\nb\tc', '😀 emoji', '&quot;x&apos;'];
  const values = ['', '1', 'İstanbul', 'a&amp;b', 'tab\there', 'line\nbreak', "it's", '&#10;', '<raw'];
  const element = (depth: number): string => {
    const name = pick(names);
    let out = '<' + name;
    const count = Math.floor(rand() * 4);
    for (let i = 0; i < count; i++) {
      const value = pick(values);
      const quote = value.includes("'") || rand() < 0.7 ? '"' : "'";
      out += pick([' ', '  ', '\n\t']) + `k${i}${pick(['=', ' = '])}${quote}${value.replace('<', '&lt;')}${quote}`;
    }
    if (rand() < 0.3) return out + pick(['/>', ' />']);
    out += '>';
    const children = depth > 5 ? 0 : Math.floor(rand() * 5);
    for (let i = 0; i < children; i++) {
      const roll = rand();
      if (roll < 0.45) out += element(depth + 1);
      else if (roll < 0.75) out += pick(texts);
      else if (roll < 0.85) out += '<!-- yorum <b> & -->';
      else if (roll < 0.93) out += '<![CDATA[ham <&> metin]]>';
      else out += '<?pi hedef?>';
    }
    return out + '</' + name + pick(['>', ' >']);
  };
  return pick(['', '<?xml version="1.0" encoding="UTF-8"?>\n']) + pick(['', '<!-- baş -->\n']) + element(0) + pick(['', '\n', '<!-- son -->']);
}

const failsBoth = (input: string, pattern: RegExp) => {
  assert.throws(() => parseXml(input), pattern, `parseXml: ${input}`);
  assert.throws(() => scanTree(input), pattern, `XmlScanner: ${input}`);
};

// --- tokens ----------------------------------------------------------------

test('xml-scan: yields open, text and close tokens with names and attributes', () => {
  const sc = new XmlScanner('<?xml version="1.0"?><x:a p="1" q:r=\'iki\'><b/>metin</x:a>');
  assert.equal(sc.next(), true);
  assert.deepEqual([sc.kind, sc.name, sc.local, sc.selfClosing], ['open', 'x:a', 'a', false]);
  assert.equal(sc.attr('p'), '1');
  assert.equal(sc.attr('q:r'), 'iki');
  assert.equal(sc.attr('r'), undefined);
  assert.equal(sc.attr('missing'), undefined);
  assert.deepEqual(sc.attrs(), [
    ['p', '1'],
    ['q:r', 'iki'],
  ]);
  assert.equal(sc.next(), true);
  assert.deepEqual([sc.kind, sc.name, sc.local, sc.selfClosing], ['open', 'b', 'b', true]);
  assert.deepEqual(sc.attrs(), []);
  // No close token for the self-closing element.
  assert.equal(sc.next(), true);
  assert.deepEqual([sc.kind, sc.text], ['text', 'metin']);
  assert.equal(sc.attr('p'), undefined);
  assert.equal(sc.next(), true);
  assert.deepEqual([sc.kind, sc.name, sc.local], ['close', 'x:a', 'a']);
  assert.equal(sc.next(), false);
  assert.equal(sc.next(), false);
});

test('xml-scan: decodes entities and normalizes attribute values like parseXml', () => {
  const sc = new XmlScanner('<a v="x&amp;y&#10;z\tq\nw &#x130;" w=\'&quot;&apos;&lt;&gt;\'>&lt;İ&#x1F600;&gt;<![CDATA[<&amp;>]]></a>');
  sc.next();
  assert.equal(sc.attr('v'), 'x&y\nz q w İ');
  assert.equal(sc.attr('w'), '"\'<>');
  sc.next();
  assert.equal(sc.text, '<İ😀>');
  sc.next();
  assert.equal(sc.text, '<&amp;>');
});

test('xml-scan: a repeated attribute reads as its last value, as in parseXml', () => {
  const xml = '<a k="1" k="2"/>';
  const sc = new XmlScanner(xml);
  sc.next();
  assert.equal(sc.attr('k'), '2');
  assert.equal(parseXml(xml).root.attrs.get('k'), '2');
});

test('xml-scan: attribute names are matched whole', () => {
  const sc = new XmlScanner('<c rr="1" r="A1" xr="3" s = "7"/>');
  sc.next();
  assert.equal(sc.attr('r'), 'A1');
  assert.equal(sc.attr('s'), '7');
  assert.equal(sc.attr('x'), undefined);
});

test('xml-scan: comments, processing instructions and the declaration are skipped', () => {
  const sc = new XmlScanner('<?xml version="1.0"?>\n<!-- a --><?pi x?><r><!-- <b> --><?p?>t</r><!-- z -->\n');
  const kinds: string[] = [];
  while (sc.next()) kinds.push(sc.kind);
  assert.deepEqual(kinds, ['open', 'text', 'close']);
});

test('xml-scan: byte input honours UTF-8 and UTF-16 byte order marks and line endings', () => {
  const xml = '<a t="ğ">Şırnak\r\nİzmir\r😀</a>';
  const utf8 = new TextEncoder().encode('﻿' + xml);
  const le = new Uint8Array(2 + xml.length * 2);
  const be = new Uint8Array(2 + xml.length * 2);
  le.set([0xff, 0xfe]);
  be.set([0xfe, 0xff]);
  for (let i = 0; i < xml.length; i++) {
    const code = xml.charCodeAt(i);
    le[2 + i * 2] = code & 0xff;
    le[3 + i * 2] = code >> 8;
    be[2 + i * 2] = code >> 8;
    be[3 + i * 2] = code & 0xff;
  }
  for (const input of [xml, utf8, le, be]) {
    const tree = scanTree(input);
    assert.deepEqual(tree, { name: 'a', attrs: [['t', 'ğ']], children: ['Şırnak\nİzmir\n😀'] });
    assertSameAsParseXml(input);
  }
});

// --- skip / readText / readRaw -----------------------------------------------

test('xml-scan: skip() moves to the matching close tag', () => {
  const sc = new XmlScanner('<r><a><a>x</a><b/></a><e/><z/></r>');
  sc.next(); // r
  sc.next(); // outer a
  sc.skip();
  assert.deepEqual([sc.kind, sc.name], ['close', 'a']);
  sc.next();
  assert.deepEqual([sc.kind, sc.name], ['open', 'e']);
  sc.skip(); // self-closing: no-op
  assert.deepEqual([sc.kind, sc.name], ['open', 'e']);
  sc.next();
  assert.equal(sc.name, 'z');
  sc.next();
  assert.deepEqual([sc.kind, sc.name], ['close', 'r']);
  assert.throws(() => sc.skip(), /open tag/);
});

test('xml-scan: readText() joins all text of the element and consumes it', () => {
  const sc = new XmlScanner('<r><t>a&amp;<b>b</b><!-- no --><![CDATA[<c>]]> d</t><e/><t></t><n>1</n></r>');
  sc.next();
  sc.next();
  assert.equal(sc.readText(), 'a&b<c> d');
  sc.next();
  assert.equal(sc.name, 'e');
  assert.equal(sc.readText(), '');
  sc.next();
  assert.equal(sc.readText(), '');
  sc.next();
  assert.equal(sc.name, 'n');
  assert.equal(sc.readText(), '1');
  sc.next();
  assert.deepEqual([sc.kind, sc.name], ['close', 'r']);
  assert.throws(() => sc.readText(), /open tag/);
});

test('xml-scan: readRaw() returns the exact source of the element', () => {
  const inner = '<x:a  k = "1&amp;2"><!-- c --><b/>t&lt;<x:a>iç</x:a></x:a >';
  const sc = new XmlScanner(`<r>${inner}<s a='1' /><t/></r>`);
  sc.next();
  sc.next();
  assert.equal(sc.readRaw(), inner);
  sc.next();
  assert.equal(sc.readRaw(), "<s a='1' />");
  sc.next();
  assert.equal(sc.name, 't');
  sc.next();
  assert.equal(sc.kind, 'close');
  assert.equal(sc.next(), false);
});

// --- differential tests against parseXml ------------------------------------------

test('xml-scan: hand-written documents read the same as with parseXml', () => {
  const documents = [
    '<a/>',
    '<a></a>',
    '<a> </a>',
    '\n<a/>\n',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<worksheet xmlns="urn:x" xmlns:r="urn:r"><sheetData><row r="1" spans="1:2"><c r="A1" t="s"><v>0</v></c><c r="B1"><f>SUM(A1:A2)&amp;""</f><v>3</v></c></row><row r="2"/></sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1"/></hyperlinks></worksheet>',
    '<a b="1"\n   c=\'2\'\td = "3"  />',
    '<a><b>bir</b>iki<c/>üç<!-- x -->dört<![CDATA[beş]]>altı</a>',
    '<a x="&lt;&gt;&amp;&quot;&apos;&#65;&#x41;">&#x1F600;&#128512; &amp;amp; &notanentity</a>',
    '<x:a xmlns:x="u"><x:b x:c="d"></x:b ></x:a\n>',
    '<a k="1" k="2" j="3"/>',
    '<a><?xml-stylesheet href="x"?><b/></a>',
    '<a>]]></a>',
    '<a b="x>y" c=\'/>\'>t</a>',
  ];
  for (const xml of documents) assertSameAsParseXml(xml);
});

test('xml-scan: generated documents read the same as with parseXml', () => {
  for (let seed = 1; seed <= 300; seed++) {
    const xml = randomDocument(seed);
    assertSameAsParseXml(xml);
    assertSameAsParseXml(new TextEncoder().encode(xml));
  }
});

test('xml-scan: readRaw() of every generated root is the whole element', () => {
  for (let seed = 1; seed <= 50; seed++) {
    const xml = randomDocument(seed);
    const sc = new XmlScanner(xml);
    sc.next();
    const raw = sc.readRaw();
    assert.ok(xml.includes(raw));
    assert.deepEqual(scanTree(raw), simplify(parseXml(xml).root));
    assert.equal(sc.next(), false);
  }
});

// --- malformed input -----------------------------------------------------------

test('xml-scan: malformed input fails with the same errors as parseXml', () => {
  failsBoth('<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>', /DTD declarations are not allowed/);
  failsBoth('<a><!ELEMENT b></a>', /DTD declarations are not allowed/);
  failsBoth('<a>&nbsp;</a>', /Unknown XML entity: &nbsp;/);
  failsBoth('<a b="&bogus;"/>', /Unknown XML entity: &bogus;/);
  failsBoth('', /no root element/);
  failsBoth('  \n ', /no root element/);
  failsBoth('<!-- only -->', /no root element/);
  failsBoth('<a>', /unclosed element <a>/);
  failsBoth('<a><b></b>', /unclosed element <a>/);
  failsBoth('<a>text', /unclosed element <a>/);
  failsBoth('<a><b></a></b>', /unexpected closing tag <\/a>/);
  failsBoth('</a>', /unexpected closing tag <\/a>/);
  failsBoth('<a></a></a>', /unexpected closing tag <\/a>/);
  failsBoth('<a></a', /missing ">"/);
  failsBoth('<a', /unterminated start tag/);
  failsBoth('<a b="1"', /unterminated start tag/);
  failsBoth('<a b="1', /missing """/);
  failsBoth("<a b='1", /missing "'"/);
  failsBoth('<a b>', /attribute "b" has no value/);
  failsBoth('<a b=1>', /attribute "b" value is not quoted/);
  failsBoth('<a / >', /expected "\/>"/);
  failsBoth('<a ="1"/>', /missing attribute name/);
  failsBoth('< a/>', /missing element name/);
  failsBoth('<a/><b/>', /more than one root element/);
  failsBoth('<a/>x', /text outside the root element/);
  failsBoth('x<a/>', /text outside the root element/);
  failsBoth('<a/><![CDATA[x]]>', /text outside the root element/);
  failsBoth('<a><!-- never closed</a>', /missing "-->"/);
  failsBoth('<a><![CDATA[never closed</a>', /missing "]]>"/);
  failsBoth('<a><?pi never closed</a>', /missing "\?>"/);
  failsBoth('<a><??></a>', /empty processing instruction/);
  failsBoth('<a><? x?></a>', /empty processing instruction/);
});

test('xml-scan: an out-of-range character reference is an error', () => {
  assert.throws(() => scanTree('<a>&#x110000;</a>'));
  assert.throws(() => scanTree('<a b="&#99999999;"/>'));
});

test('xml-scan: every truncation of a document ends in an error, never a hang', () => {
  const xml = '<?xml version="1.0"?><!-- c --><a k="v&amp;" j=\'2\'><b>t&lt;</b><![CDATA[x]]><c/><?p q?></a>';
  for (let length = 0; length < xml.length; length++) {
    const cut = xml.slice(0, length);
    let scanError: unknown;
    let parseError: unknown;
    try {
      scanTree(cut);
    } catch (error) {
      scanError = error;
    }
    try {
      parseXml(cut);
    } catch (error) {
      parseError = error;
    }
    assert.ok(scanError instanceof Error, `scanner accepted: ${cut}`);
    assert.ok(parseError instanceof Error, `parseXml accepted: ${cut}`);
    assert.equal(scanError.message, parseError.message, cut);
  }
});

test('xml-scan: skip(), readText() and readRaw() fail on truncated input', () => {
  for (const method of ['skip', 'readText', 'readRaw'] as const) {
    const sc = new XmlScanner('<a><b><c>x</c>');
    sc.next();
    sc.next();
    assert.throws(() => sc[method](), /unclosed element <b>/);
  }
});

test('xml-scan: deep nesting does not overflow the stack', () => {
  const depth = 200000;
  const xml = '<a>'.repeat(depth) + 'x' + '</a>'.repeat(depth);
  const sc = new XmlScanner(xml);
  let opens = 0;
  let closes = 0;
  while (sc.next()) {
    if (sc.kind === 'open') opens++;
    else if (sc.kind === 'close') closes++;
  }
  assert.equal(opens, depth);
  assert.equal(closes, depth);

  const again = new XmlScanner(xml);
  again.next();
  assert.equal(again.readText(), 'x');
  assert.equal(again.next(), false);
});

// --- speed ------------------------------------------------------------------

test('xml-scan: scans a million cells quickly', () => {
  const parts: string[] = ['<worksheet><sheetData>'];
  for (let row = 1; row <= 100000; row++) {
    let line = `<row r="${row}" spans="1:10">`;
    for (let col = 0; col < 10; col++) line += `<c r="${String.fromCharCode(65 + col)}${row}" s="3"><v>${row * 10 + col}</v></c>`;
    parts.push(line + '</row>');
  }
  parts.push('</sheetData></worksheet>');
  const xml = parts.join('');

  const started = Date.now();
  const sc = new XmlScanner(xml);
  let cells = 0;
  let sum = 0;
  let lastRef: string | undefined;
  while (sc.next()) {
    if (sc.kind === 'open' && sc.local === 'c') {
      lastRef = sc.attr('r');
      cells++;
    } else if (sc.kind === 'text') sum += Number(sc.text);
  }
  const elapsed = Date.now() - started;
  console.log(`xml-scan: ${(xml.length / 1e6).toFixed(1)} MB, ${cells} cells in ${elapsed} ms`);
  assert.equal(cells, 1000000);
  assert.equal(lastRef, 'J100000');
  assert.equal(sum, 10 * ((100000 * 100001) / 2) * 10 + 45 * 100000);
  assert.ok(elapsed < 10000, `took ${elapsed} ms`);
});
