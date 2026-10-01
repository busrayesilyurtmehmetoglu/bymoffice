import assert from 'node:assert/strict';
import { test } from 'node:test';
import { NumberingCounters, abstractOf, formatCounter, levelOf } from '../src/formats/docx/numbering';
import { IMPLICIT_PARA_PROPS, resolveParaProps } from '../src/formats/docx/resolve';
import type { Paragraph } from '../src/model/doc/types';
import { readSpec } from './docx-fixtures';

const lvl = (ilvl: number, fmt: string, text: string, extra = '', start = 1) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="${start}"/><w:numFmt w:val="${fmt}"/>${extra}<w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/></w:lvl>`;
const abstract = (id: number, inner: string) => `<w:abstractNum w:abstractNumId="${id}">${inner}</w:abstractNum>`;
const num = (id: number, abstractId: number, inner = '') => `<w:num w:numId="${id}"><w:abstractNumId w:val="${abstractId}"/>${inner}</w:num>`;

const NUMBERING =
  abstract(
    0,
    '<w:nsid w:val="12345678"/><w:multiLevelType w:val="multilevel"/>' +
      '<w:lvl w:ilvl="0" w:tplc="041F000F"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:suff w:val="space"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>' +
      '<w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:b/><w:rFonts w:ascii="Arial"/></w:rPr></w:lvl>' +
      lvl(1, 'lowerLetter', '%1.%2)', '<w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>') +
      lvl(2, 'upperRoman', '%1.%2.%3', '<w:lvlRestart w:val="0"/><w:legacy w:legacy="1"/>') +
      lvl(3, 'decimalZero', '%4:', '<w:lvlRestart w:val="1"/>'),
  ) +
  abstract(1, lvl(0, 'bullet', '', '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>') + lvl(1, 'bullet', 'o')) +
  abstract(2, lvl(0, 'upperRoman', 'Madde %1', '', 4) + lvl(1, 'decimal', '%1.%2', '<w:isLgl/>') + lvl(2, 'none', '(%3)') + lvl(3, 'chicago', '%4.')) +
  // Takes its levels from a numbering style; 4 and the style point at each other.
  abstract(3, '<w:numStyleLink w:val="Liste"/>') +
  abstract(4, '<w:styleLink w:val="Liste"/>' + lvl(0, 'ordinal', '%1')) +
  abstract(5, '<w:numStyleLink w:val="Dongu"/>') +
  abstract(6, '<w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl><w:lvl w:ilvl="12"><w:start w:val="1"/></w:lvl><w:lvl><w:start w:val="1"/></w:lvl>') +
  abstract(0, lvl(0, 'bullet', 'tekrar')) +
  num(1, 0) +
  num(2, 0) +
  num(3, 0, '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="3"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="[%2]"/></w:lvl></w:lvlOverride>') +
  num(4, 1) +
  num(5, 2) +
  num(6, 3) +
  num(7, 4) +
  num(8, 5) +
  num(9, 6) +
  num(10, 99) +
  '<w:num w:numId="11"/><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>';

const STYLES =
  '<w:style w:type="numbering" w:styleId="Liste"><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="numbering" w:styleId="Dongu"><w:pPr><w:numPr><w:numId w:val="8"/></w:numPr></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="100"/></w:pPr></w:style>';

test('docx numbering: counter text in each number format', () => {
  const sample = (format: string | undefined, values: number[]) => values.map((value) => formatCounter(value, format));
  assert.deepEqual(sample('decimal', [0, 1, 10, 123]), ['0', '1', '10', '123']);
  assert.deepEqual(sample('decimalZero', [1, 9, 10, 100]), ['01', '09', '10', '100']);
  assert.deepEqual(sample('lowerLetter', [1, 2, 26, 27, 28, 52, 53, 0, 781]), ['a', 'b', 'z', 'aa', 'bb', 'zz', 'aaa', '0', '781']);
  assert.deepEqual(sample('upperLetter', [1, 27]), ['A', 'AA']);
  assert.deepEqual(sample('lowerRoman', [1, 4, 9, 14, 40, 90, 400, 1994, 2024, 3999, 0]), [
    'i',
    'iv',
    'ix',
    'xiv',
    'xl',
    'xc',
    'cd',
    'mcmxciv',
    'mmxxiv',
    'mmmcmxcix',
    '0',
  ]);
  assert.deepEqual(sample('upperRoman', [3, 8]), ['III', 'VIII']);
  assert.deepEqual(sample('ordinal', [1, 2, 3, 4, 11, 12, 13, 21, 22, 23, 101, 111]), [
    '1st',
    '2nd',
    '3rd',
    '4th',
    '11th',
    '12th',
    '13th',
    '21st',
    '22nd',
    '23rd',
    '101st',
    '111th',
  ]);
  assert.deepEqual(sample('bullet', [1]), ['']);
  assert.deepEqual(sample('none', [5]), ['']);
  // Formats without an implementation, and no format at all, are decimal.
  assert.deepEqual(sample('chicago', [3]), ['3']);
  assert.deepEqual(sample('cardinalText', [3]), ['3']);
  assert.deepEqual(sample(undefined, [7]), ['7']);
  assert.deepEqual(sample('constructor', [7]), ['7']);
  // Absurd values do not produce absurdly long text.
  assert.equal(formatCounter(2147483647, 'lowerRoman'), '2147483647');
  assert.equal(formatCounter(2147483647, 'upperLetter'), '2147483647');
});

test('docx numbering: definitions, instances and overrides are read', async () => {
  const { numbering, styles, unknown } = await readSpec({ numbering: NUMBERING, styles: STYLES, body: '<w:p/>' });
  assert.deepEqual([...numbering.abstracts.keys()], [0, 1, 2, 3, 4, 5, 6]);
  // An instance without a definition reference and a repeated id are left out; a dangling reference is kept.
  assert.deepEqual([...numbering.nums.keys()], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);

  const first = numbering.abstracts.get(0)!;
  assert.equal(first.multiLevelType, 'multilevel');
  assert.deepEqual([...first.levels.keys()], [0, 1, 2, 3]);
  const level0 = first.levels.get(0)!;
  assert.deepEqual(
    { ...level0, extras: undefined },
    {
      level: 0,
      start: 1,
      format: 'decimal',
      paraStyle: 'Heading1',
      suffix: 'space',
      text: '%1.',
      align: 'left',
      para: { tabs: [{ pos: 720, align: 'num' }], indent: { left: 720, hanging: 360 } },
      run: { bold: true, fonts: { ascii: 'Arial' } },
      extras: undefined,
    },
  );
  // w:tplc is not interpreted: the level is kept for it.
  assert.deepEqual(level0.extras!.map((extra) => [extra.index, extra.partial]), [[-1, true]]);
  assert.equal(first.levels.get(2)!.restart, 0);
  assert.deepEqual(first.levels.get(2)!.extras!.map((extra) => extra.el.name), ['w:legacy']);
  assert.equal(unknown.get('lvl/w:legacy'), 1);
  assert.equal(numbering.abstracts.get(2)!.levels.get(1)!.legal, true);
  // Levels outside 0..8 and levels without an index are not levels.
  assert.deepEqual([...numbering.abstracts.get(6)!.levels.keys()], [0]);

  const third = numbering.nums.get(3)!;
  assert.equal(third.abstractId, 0);
  assert.deepEqual([...third.overrides.keys()], [0, 1]);
  assert.deepEqual(third.overrides.get(0), { start: 10 });
  assert.equal(third.overrides.get(1)!.level!.format, 'upperLetter');

  assert.equal(levelOf(numbering, styles, 3, 1)!.text, '[%2]');
  assert.equal(levelOf(numbering, styles, 3, 0)!.text, '%1.');
  assert.equal(levelOf(numbering, styles, 1, 7), undefined);
  assert.equal(levelOf(numbering, styles, 10, 0), undefined);
  assert.equal(levelOf(numbering, styles, 404, 0), undefined);

  // The style link leads to the definition that holds the levels.
  assert.equal(abstractOf(numbering, styles, 6)!.id, 4);
  assert.equal(levelOf(numbering, styles, 6, 0)!.format, 'ordinal');
  // A link that leads back to itself ends where it started.
  assert.equal(abstractOf(numbering, styles, 8)!.id, 5);
  assert.equal(levelOf(numbering, styles, 8, 0), undefined);
  // Without the styles the link cannot be followed.
  assert.equal(abstractOf(numbering, undefined, 6)!.id, 3);
});

test('docx numbering: counters run through the document', async () => {
  const { numbering, styles } = await readSpec({ numbering: NUMBERING, styles: STYLES, body: '<w:p/>' });
  const counters = new NumberingCounters(numbering, styles);
  const next = (numId: number, level: number) => counters.next(numId, level)?.text;

  assert.equal(next(1, 0), '1.');
  assert.equal(next(1, 0), '2.');
  assert.equal(next(1, 1), '2.a)');
  assert.equal(next(1, 1), '2.b)');
  assert.equal(next(1, 2), '2.b.I');
  assert.equal(next(1, 3), '01:');
  // A higher level restarts the deeper ones...
  assert.equal(next(1, 0), '3.');
  assert.equal(next(1, 1), '3.a)');
  // ...except a level that never restarts (lvlRestart 0)...
  assert.equal(next(1, 2), '3.a.II');
  // ...and level 4 here only restarts after level 1 (lvlRestart 1), which was used since.
  assert.equal(next(1, 3), '01:');
  assert.equal(next(1, 2), '3.a.III');
  assert.equal(next(1, 3), '02:');
  assert.equal(next(1, 1), '3.b)');
  assert.equal(next(1, 3), '03:');

  // Another instance of the same definition continues the same counters.
  assert.equal(next(2, 0), '4.');
  assert.equal(next(2, 1), '4.a)');
  // An instance with a start override restarts the first time it is used, then counts on.
  assert.equal(next(3, 0), '10.');
  assert.equal(next(3, 0), '11.');
  // Its level 1 is replaced as a whole: other text, other format, other start.
  assert.equal(next(3, 1), '[C]');
  assert.equal(next(3, 1), '[D]');
  assert.equal(next(1, 0), '12.');
  assert.equal(next(3, 0), '13.');

  // Bullets show their text as it is.
  assert.equal(next(4, 0), '');
  assert.equal(next(4, 1), 'o');
  assert.equal(next(4, 0), '');

  // Start values, a level used before its parent, legal numbering and "none".
  const label = counters.next(5, 1)!;
  assert.equal(label.text, '4.1');
  assert.deepEqual(label.values, [4, 1]);
  assert.equal(label.level.legal, true);
  assert.equal(next(5, 0), 'Madde IV');
  assert.equal(next(5, 0), 'Madde V');
  assert.equal(next(5, 1), '5.1');
  assert.equal(next(5, 2), '()');
  assert.equal(next(5, 3), '1.');

  assert.equal(next(6, 0), '1st');
  assert.equal(next(7, 0), '2nd');
  // No start element: counting starts at 1.
  assert.equal(next(9, 0), '1');
  // A level outside 0..8 is read as the first level.
  assert.equal(next(9, 42), '2');
  assert.equal(next(9, -1), '3');

  // Nothing to number with: no label, no error.
  assert.equal(counters.next(8, 0), undefined);
  assert.equal(counters.next(10, 0), undefined);
  assert.equal(counters.next(404, 0), undefined);
  assert.equal(counters.next(1, 8), undefined);

  // A second pass starts from the beginning.
  assert.equal(new NumberingCounters(numbering, styles).next(1, 0)!.text, '1.');
  // Text with "%" that is not a placeholder is left alone.
  const odd = await readSpec({ numbering: abstract(0, lvl(0, 'decimal', '%1 %% %0 %a %9')) + num(1, 0), body: '<w:p/>' });
  assert.equal(new NumberingCounters(odd.numbering).next(1, 0)!.text, '1 %% %0 %a ');
});

test('docx numbering: the level gives the paragraph its indents, under direct formatting and under the style that brings the numbering', async () => {
  const doc = await readSpec({
    numbering: NUMBERING,
    styles: STYLES,
    body:
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="2000"/></w:pPr></w:p>' +
      // Numbered through its style, which names no level: the level tied to the style is used.
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>' +
      // numId 0 takes the numbering of the style away.
      '<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr></w:p>' +
      '<w:p><w:pPr><w:numPr><w:numId w:val="404"/></w:numPr></w:pPr></w:p>',
  });
  const props = (i: number) => resolveParaProps(doc, (doc.body.blocks[i] as Paragraph).props);
  assert.deepEqual(props(0), { ...IMPLICIT_PARA_PROPS, numbering: { level: 1, numId: 1 }, indent: { left: 1440, hanging: 360 } });
  assert.deepEqual(props(1), {
    ...IMPLICIT_PARA_PROPS,
    numbering: { level: 0, numId: 1 },
    tabs: [{ pos: 720, align: 'num' }],
    indent: { left: 2000, hanging: 360 },
  });
  // The style states the numbering and an indent of its own, which wins over the one of the level.
  assert.deepEqual(props(2), {
    ...IMPLICIT_PARA_PROPS,
    style: 'Heading1',
    numbering: { numId: 1 },
    tabs: [{ pos: 720, align: 'num' }],
    indent: { left: 100, hanging: 360 },
  });
  assert.deepEqual(props(3), { ...IMPLICIT_PARA_PROPS, style: 'Heading1', numbering: { numId: 0 }, indent: { left: 100 } });
  assert.deepEqual(props(4), { ...IMPLICIT_PARA_PROPS, numbering: { numId: 404 } });
});
