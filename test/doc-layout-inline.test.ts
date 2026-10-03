import assert from 'node:assert/strict';
import { test } from 'node:test';
import { paragraphText, type Paragraph } from '../src/model/doc/types';
import { at, baselines, find, lay, p, pageText, r, round, sectPr, texts, words } from './doc-layout-fixtures';

// A character is 8 px, a line 16 px; a line holds 40 characters.

const a = (count: number, char = 'a') => new Array(count + 1).join(char);

test('doc layout: lines break at spaces and after hyphens, not at non-breaking characters', async () => {
  const hyphen = await lay(p(a(35) + '-' + a(10, 'b')));
  assert.deepEqual(at(hyphen.pages[0]), [a(35) + '-@40,52.8', a(10, 'b') + '@40,68.8']);
  // A hyphen before a digit, or one that does not follow a word, is not a place to break.
  const minus = await lay(p(a(30) + ' ' + a(6) + ' -' + a(5, '5')));
  assert.deepEqual(at(minus.pages[0]), [a(30) + ' ' + a(6) + ' @40,52.8', '-55555@40,68.8']);
  const range = await lay(p(a(35) + '-' + a(10, '5')));
  assert.deepEqual(
    texts(range.pages[0]).map((item) => item.text.length),
    [40, 6],
  );

  // A non-breaking space and a non-breaking hyphen hold the word together: it is cut where the line is full.
  const nbsp = await lay(p(a(30) + '\u00a0' + a(15, 'b')));
  assert.deepEqual(
    texts(nbsp.pages[0]).map((item) => item.text.length),
    [40, 6],
  );
  const hard = await lay(p(r(a(30)) + '<w:r><w:noBreakHyphen/></w:r>' + r(a(15, 'b'))));
  assert.deepEqual(
    texts(hard.pages[0]).map((item) => item.text),
    [a(30) + '\u2011' + a(9, 'b'), a(6, 'b')],
  );
});

test('doc layout: a word longer than the line is cut between characters', async () => {
  const { pages } = await lay(p(a(100)) + p('x ' + a(50, 'b')));
  assert.deepEqual(at(pages[0]), [
    a(40) + '@40,52.8',
    a(40) + '@40,68.8',
    a(20) + '@40,84.8',
    // A word that does not fit after other words starts a line of its own first.
    'x @40,100.8',
    a(40, 'b') + '@40,116.8',
    a(10, 'b') + '@40,132.8',
  ]);
  assert.deepEqual(
    texts(pages[0]).map((item) => item.source!.offset),
    [0, 40, 80, 0, 2, 42],
  );
});

test('doc layout: a word is not broken where its formatting changes', async () => {
  const { pages } = await lay(p(r(words(7) + ' abc') + r('def', '<w:b/>') + r('ghi more')));
  // 35 characters, then "abcdefghi" in three runs: it goes to the next line as one word.
  assert.deepEqual(at(pages[0]), ['www0 www1 www2 www3 www4 www5 www6 @40,52.8', 'abc@40,68.8', 'def@64,68.8', 'ghi more@90.4,68.8']);
});

test('doc layout: East Asian text breaks between any two characters, but not before closing punctuation', async () => {
  const { pages } = await lay(p(a(39, '\u6f22') + '\u5b57\u3002' + a(9, '\u6f22')));
  // 40 characters fit, but the full stop may not start a line: the character before it goes down with it.
  assert.deepEqual(
    texts(pages[0]).map((item) => item.text.length),
    [39, 11],
  );
  const mixed = await lay(p(a(38) + ' ' + a(5, '\u6f22')));
  assert.deepEqual(
    texts(mixed.pages[0]).map((item) => item.text),
    [a(38) + ' \u6f22', a(4, '\u6f22')],
  );
});

test('doc layout: a soft hyphen shows only where the line breaks at it', async () => {
  const broken = await lay(p(a(30) + '\u00ad' + a(15, 'b')));
  // The hyphen is an item of its own, without a source: it is not in the text of the paragraph.
  assert.deepEqual(at(broken.pages[0]), [a(30) + '@40,52.8', '-@280,52.8', a(15, 'b') + '@40,68.8']);
  assert.equal(texts(broken.pages[0])[1].source, undefined);
  assert.deepEqual(texts(broken.pages[0])[2].source, { story: 'body', path: [0], offset: 31 });
  const whole = await lay(p('aaaa\u00adbbbb') + p(r('cccc') + '<w:r><w:softHyphen/></w:r>' + r('dddd')));
  assert.equal(
    texts(whole.pages[0])
      .map((item) => item.text)
      .join(''),
    'aaaabbbbccccdddd',
  );
  assert.deepEqual(baselines(whole.pages[0]), [52.8, 68.8]);
});

test('doc layout: line, page and column breaks', async () => {
  const br = (type?: string) => `<w:r><w:br${type ? ` w:type="${type}"` : ''}/></w:r>`;
  const { pages } = await lay({ body: p(r('one') + br() + r('two') + br('page') + r('three')) + p('four'), settings: '' });
  assert.equal(pages.length, 2);
  assert.deepEqual(at(pages[0]), ['one@40,52.8', 'two@40,68.8']);
  assert.deepEqual(at(pages[1]), ['three@40,52.8', 'four@40,68.8']);

  // A line break at the end of a paragraph leaves an empty line for the paragraph mark.
  const tail = await lay(p(r('one') + br()) + p('two'));
  assert.deepEqual(at(tail.pages[0]), ['one@40,52.8', 'two@40,84.8']);

  // A column break in a section of one column is a page break.
  const column = await lay(p(r('one') + br('column') + r('two')));
  assert.equal(column.pages.length, 2);
});

test('doc layout: the paragraph mark after a page break stays on the page of the break in new documents only', async () => {
  const body = p('one') + p('<w:r><w:br w:type="page"/></w:r>') + p('two');
  const modern = await lay({ body, settings: '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' });
  assert.deepEqual(at(modern.pages[1]), ['two@40,52.8']);
  // Word 2010 and earlier: the mark goes to the new page, as an empty line above what follows.
  const old = await lay(body);
  assert.deepEqual(at(old.pages[1]), ['two@40,68.8']);
});

test('doc layout: the tallest run sets the height of the line; all runs share the baseline', async () => {
  const { pages } = await lay(p(r('small ') + r('BIG', '<w:sz w:val="48"/>') + r(' small')) + p('next'));
  // 24 pt: 32 px, ascent 25.6. The line is 32 high.
  assert.deepEqual(at(pages[0]), ['small @40,65.6', 'BIG@88,65.6', ' small@136,65.6', 'next@40,84.8']);
  assert.equal(find(pages[0], 'BIG').font.size, 32);
  // An empty paragraph has the height of its paragraph mark.
  const empty = await lay(p('one') + '<w:p><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr></w:p>' + p('two'));
  assert.deepEqual(baselines(empty.pages[0]), [52.8, 100.8]);
});

test('doc layout: superscript, subscript, raised text', async () => {
  const { pages } = await lay(p(r('x') + r('2', '<w:vertAlign w:val="superscript"/>') + r('y') + r('i', '<w:vertAlign w:val="subscript"/>') + r('up', '<w:position w:val="6"/>')));
  const [x, two, y, i, up] = texts(pages[0]);
  // The raised text needs 4 px more above the baseline: the line grows by that much.
  assert.equal(x.y, 56.8);
  // 65 % of the size, a third of the size up or 14 % down.
  assert.equal(round(two.font.size), 10.4);
  assert.equal(round(two.y), round(56.8 - 16 * 0.33));
  assert.equal(round(two.width), 5.2);
  assert.equal(y.x, 40 + 8 + 5.2);
  assert.equal(round(i.y), round(56.8 + 16 * 0.14));
  // 3 pt up.
  assert.equal(up.y, 56.8 - 4);
});

test('doc layout: capitals, small capitals, character spacing and scale', async () => {
  const { pages } = await lay(
    p(r('Caps ok', '<w:caps/>')) + p(r('Small Caps', '<w:smallCaps/>')) + p(r('wide', '<w:spacing w:val="30"/>') + r('x')) + p(r('half', '<w:w w:val="50"/>') + r('x')) + p(r('ilk', '<w:caps/><w:lang w:val="tr-TR"/>')),
  );
  const items = texts(pages[0]);
  assert.equal(items[0].text, 'CAPS OK');
  // Small capitals: the lower case letters as capitals of 80 % of the size.
  const small = items.filter((item) => round(item.y) === 68.8);
  assert.deepEqual(
    small.map((item) => [item.text, round(item.font.size), round(item.x)]),
    [
      ['S', 16, 40],
      ['MALL ', 12.8, 48],
      ['C', 16, 80],
      ['APS', 12.8, 88],
    ],
  );
  // 1.5 pt = 2 px after every character.
  const wide = find(pages[0], 'wide');
  assert.equal(wide.letterSpacing, 2);
  assert.equal(wide.width, 40);
  assert.equal(items[items.indexOf(wide) + 1].x, 80);
  const half = find(pages[0], 'half');
  assert.equal(half.scale, 0.5);
  assert.equal(half.width, 16);
  // Turkish text: the capital of i is dotted.
  assert.equal(items[items.length - 1].text, '\u0130LK');
});

test('doc layout: hidden text is left out; decorations and colours come through', async () => {
  const { pages, doc } = await lay(
    p(r('seen ') + r('hidden ', '<w:vanish/>') + r('again')) +
      p(
        r('u', '<w:u w:val="double" w:color="FF0000"/>') +
          r('s', '<w:strike/>') +
          r('d', '<w:dstrike/>') +
          r('h', '<w:highlight w:val="yellow"/>') +
          r('c', '<w:color w:val="00FF00"/><w:shd w:val="clear" w:fill="0000FF"/>') +
          r('n', '<w:u w:val="none"/>'),
      ),
  );
  const items = texts(pages[0]);
  assert.deepEqual(
    items.slice(0, 2).map((item) => [item.text, item.x, item.source!.offset]),
    [
      ['seen ', 40, 0],
      // The offset counts the hidden text: it is the place in the text of the paragraph.
      ['again', 80, 12],
    ],
  );
  assert.equal(paragraphText(doc.body.blocks[0] as Paragraph), 'seen hidden again');
  const [u, s, d, h, c, n] = items.slice(2);
  assert.deepEqual(u.underline, { style: 'double', color: '#FF0000' });
  assert.equal(s.strike, true);
  assert.equal(d.strike, 'double');
  assert.equal(h.highlight, '#FFFF00');
  assert.equal(c.color, '#00FF00');
  assert.equal(c.highlight, '#0000FF');
  assert.equal(n.underline, undefined);
  assert.equal(n.color, '#000000');
});

test('doc layout: hyperlinks carry their target only when it is safe', async () => {
  const link = (id: string, text: string) => `<w:hyperlink r:id="${id}">${r(text)}</w:hyperlink>`;
  const { pages } = await lay({
    body: p(link('rId1', 'good') + r(' ') + link('rId2', 'bad') + r(' ') + `<w:hyperlink w:anchor="mark">${r('inside')}</w:hyperlink>`),
    rels: [
      ['rId1', 'hyperlink', 'https://example.com/a', true],
      ['rId2', 'hyperlink', 'javascript:alert(1)', true],
    ],
  });
  assert.equal(find(pages[0], 'good').link, 'https://example.com/a');
  assert.equal(find(pages[0], ' bad').link, undefined);
  assert.equal(find(pages[0], 'inside').link, '#mark');
});

const field = (code: string, result: string) =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve">${code}</w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  r(result) +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

test('doc layout: fields show their result, never their code; PAGE and NUMPAGES are computed', async () => {
  const body =
    p(r('Page ') + field(' PAGE ', '99') + r(' of ') + field(' NUMPAGES ', '99') + r('.')) +
    p(field(' DATE \\@ "d.M.yyyy" ', '1.2.2020') + r(' ') + field(' REF mark \\h ', 'cached')) +
    p(field(' HYPERLINK "https://example.com/x" \\o "tip" ', 'link') + r(' ') + field(' HYPERLINK \\l "mark" ', 'jump') + r(' ') + field(' HYPERLINK "file:///c:/x" ', 'file')) +
    p(`<w:fldSimple w:instr=" PAGE \\* ROMAN ">${r('7')}</w:fldSimple>` + r(' ') + `<w:fldSimple w:instr=" AUTHOR ">${r('Someone')}</w:fldSimple>`) +
    p('<w:r><w:br w:type="page"/></w:r>') +
    p(r('p') + field(' PAGE ', '1') + r('s') + field(' SECTIONPAGES ', '5'));
  const { pages } = await lay(body);
  assert.equal(pages.length, 2);
  assert.deepEqual(at(pages[0]), [
    'Page @40,52.8',
    '1@80,52.8',
    ' of @88,52.8',
    // The total is filled in when it is known; what follows on the line stays where the cached "99" put it.
    '2@120,52.8',
    '.@136,52.8',
    '1.2.2020 cached@40,68.8',
    'link@40,84.8',
    ' @72,84.8',
    'jump@80,84.8',
    ' file@112,84.8',
    'I@40,100.8',
    ' Someone@48,100.8',
  ]);
  assert.equal(find(pages[0], 'link').link, 'https://example.com/x');
  assert.equal(find(pages[0], 'jump').link, '#mark');
  assert.equal(find(pages[0], ' file').link, undefined);
  assert.deepEqual(at(pages[1]).slice(-4), ['p@40,68.8', '2@48,68.8', 's@56,68.8', '2@64,68.8']);
  // Nothing of a field code is ever shown.
  for (const page of pages) for (const item of texts(page)) assert.ok(!/PAGE|DATE|REF|HYPERLINK|AUTHOR/.test(item.text), item.text);
});

test('doc layout: tracked changes show the final text', async () => {
  const { pages, doc } = await lay(
    p(
      r('keep ') +
        '<w:del w:id="1" w:author="a"><w:r><w:delText>gone </w:delText></w:r></w:del>' +
        `<w:ins w:id="2" w:author="a">${r('new ')}</w:ins>` +
        `<w:moveFrom w:id="3" w:author="a">${r('from ')}</w:moveFrom>` +
        `<w:moveTo w:id="4" w:author="a">${r('to')}</w:moveTo>`,
    ),
  );
  assert.equal(pageText(pages[0]), 'keep new to');
  assert.equal(pageText(pages[0]), paragraphText(doc.body.blocks[0] as Paragraph));
});

test('doc layout: symbols and formulas', async () => {
  const { pages } = await lay(
    p(r('a') + '<w:r><w:sym w:font="Wingdings" w:char="F0FC"/></w:r>' + '<w:r><w:sym w:font="Symbol" w:char="F061"/></w:r>' + '<w:r><w:sym w:font="Webdings" w:char="F041"/></w:r>') +
      p(r('f: ') + '<m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath>'),
  );
  const items = texts(pages[0]);
  assert.equal(items[0].text + items[1].text + items[2].text, 'a\u2713\u03b1');
  // A character the tables do not know stays as it is, in its font.
  assert.equal(items[3].text, '\uf041');
  assert.equal(items[3].font.family, 'Webdings');
  assert.equal(find(pages[0], 'f: ').text, 'f: x+1');
});

test('doc layout: text items say where their text comes from', async () => {
  const body = p(r('one ') + r('two', '<w:b/>') + '<w:r><w:tab/></w:r>' + r('three')) + p(words(12));
  const { pages, doc } = await lay(body + sectPr());
  for (const item of texts(pages[0])) {
    const source = item.source!;
    assert.equal(source.story, 'body');
    const text = paragraphText(doc.body.blocks[source.path[0]] as Paragraph);
    assert.equal(text.substr(source.offset, item.text.length), item.text);
  }
});
