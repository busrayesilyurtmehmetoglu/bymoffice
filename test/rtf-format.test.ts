// Formatting of an RTF file: character and paragraph formatting as groups
// save and restore it, styles and what the resolvers make of them, lists,
// sections and the settings of the document.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { NumberingCounters } from '../src/formats/docx/numbering';
import { resolveColor, resolveParaProps } from '../src/formats/docx/resolve';
import { findSections } from '../src/layout/doc/page';
import { paragraphText, type DocModel, type Paragraph } from '../src/model/doc/types';
import { doc, effectivePara, effectiveRun, paragraphs, read, runWith, runsOf, texts } from './rtf-fixtures';

const first = (model: DocModel): Paragraph => paragraphs(model)[0];

test('character formatting switches on with a word and off with the word and a zero', () => {
  const model = read(doc(String.raw`\b a\b0 b\i c\i0\ul d\ulnone\strike e\strike0\caps f\caps0\scaps g\scaps0\v h\v0 i\par`));
  const p = first(model);
  assert.equal(runWith(p, 'a').props.bold, true);
  // Nothing states bold for it any more.
  assert.equal(effectiveRun(model, p, 'b').bold, undefined);
  assert.equal(effectiveRun(model, p, 'c').italic, true);
  assert.equal(effectiveRun(model, p, 'd').underline!.style, 'single');
  assert.equal(effectiveRun(model, p, 'e').strike, true);
  assert.equal(effectiveRun(model, p, 'f').caps, true);
  assert.equal(effectiveRun(model, p, 'g').smallCaps, true);
  assert.equal(effectiveRun(model, p, 'h').vanish, true);
  assert.deepEqual(runWith(p, 'i').props, {});
});

test('a group restores the formatting it found', () => {
  const model = read(doc(String.raw`a{\b\i\fs40 b{\b0 c}d}e\par`));
  const p = first(model);
  assert.deepEqual(runWith(p, 'a').props, {});
  assert.deepEqual(runWith(p, 'b').props, { bold: true, italic: true, size: 40, sizeCs: 40 });
  assert.deepEqual(runWith(p, 'c').props, { italic: true, size: 40, sizeCs: 40 });
  assert.equal(runWith(p, 'd').props, runWith(p, 'b').props);
  assert.equal(runWith(p, 'e').props, runWith(p, 'a').props);
});

test('"\\plain" sets the character formatting back: 12 points, the default font, nothing else', () => {
  const model = read(doc(String.raw`\b\i\fs40\f1\cf1 a\plain b\par`));
  const p = first(model);
  assert.deepEqual(runWith(p, 'b').props, {});
  const effective = effectiveRun(model, p, 'b');
  assert.equal(effective.size, 24);
  assert.equal(effective.fonts!.ascii, 'Times New Roman');
  assert.equal(effective.bold, undefined);
});

test('font, size and colours', () => {
  const model = read(doc(String.raw`\f1\fs18\cf1 a\cf0\highlight3 b\highlight0\chcbpat2 c\chcbpat0\cb1 d\par`));
  const p = first(model);
  const a = effectiveRun(model, p, 'a');
  assert.equal(a.fonts!.ascii, 'Arial');
  assert.equal(a.fonts!.hAnsi, 'Arial');
  assert.equal(a.size, 18);
  assert.equal(resolveColor(model, a.color), '#FF0000');
  const b = effectiveRun(model, p, 'b');
  assert.equal(resolveColor(model, b.color), undefined);
  assert.equal(b.highlight, 'yellow');
  const c = effectiveRun(model, p, 'c');
  assert.equal(c.highlight, undefined);
  assert.deepEqual(c.shading, { pattern: 'clear', fill: { rgb: '0000FF' } });
  assert.deepEqual(effectiveRun(model, p, 'd').shading, { pattern: 'clear', fill: { rgb: 'FF0000' } });
});

test('a colour number the table does not have is the automatic colour', () => {
  const model = read(doc(String.raw`\cf99 a\cf-4 b\par`));
  assert.deepEqual(runWith(first(model), 'a').props, {});
});

test('underline kinds and the colour of the underline', () => {
  const model = read(doc(String.raw`\uld a\uldb b\ulw c\ulwave d\ulth e\uldash f\ul\ulc1 g\ul0 h\par`));
  const p = first(model);
  const style = (text: string) => effectiveRun(model, p, text).underline!.style;
  assert.deepEqual(['a', 'b', 'c', 'd', 'e', 'f'].map(style), ['dotted', 'double', 'words', 'wave', 'thick', 'dash']);
  assert.deepEqual(effectiveRun(model, p, 'g').underline, { style: 'single', color: { rgb: 'FF0000' } });
  assert.equal(effectiveRun(model, p, 'h').underline, undefined);
});

test('raised and lowered text, spacing, scale and the rest', () => {
  const model = read(doc(String.raw`\super a\nosupersub\sub b\nosupersub\up8 c\up0\dn4 d\dn0\expndtw20 e\expndtw0\charscalex150 f\charscalex100\striked1 g\striked0\kerning28 h\kerning0\lang1055 i\par`));
  const p = first(model);
  assert.equal(runWith(p, 'a').props.vertAlign, 'superscript');
  assert.equal(runWith(p, 'b').props.vertAlign, 'subscript');
  assert.equal(runWith(p, 'c').props.position, 8);
  assert.equal(runWith(p, 'd').props.position, -4);
  assert.equal(runWith(p, 'e').props.spacing, 20);
  assert.equal(runWith(p, 'f').props.scale, 150);
  assert.equal(runWith(p, 'g').props.doubleStrike, true);
  assert.equal(runWith(p, 'h').props.kerning, 28);
  assert.deepEqual(runWith(p, 'i').props, { lang: { val: 'tr-TR' } });
});

test('the words Word writes for complex scripts do not change the formatting of other text', () => {
  const model = read(doc(String.raw`{\rtlch\fcs1 \ab\ai\af1\afs40\alang1025 \ltrch\fcs0 \f0\fs24\lang1055 a}\par`));
  const props = runWith(first(model), 'a').props;
  assert.equal(props.bold, undefined);
  assert.equal(props.italic, undefined);
  assert.equal(props.size, undefined);
  assert.equal(props.boldCs, true);
  assert.equal(props.italicCs, true);
  assert.equal(props.sizeCs, 40);
  assert.equal(props.fonts!.cs, 'Arial');
  assert.equal(props.fonts!.ascii, undefined);
  assert.equal(props.rtl, undefined);
});

test('paragraph formatting: alignment, indents, spacing', () => {
  const model = read(doc(String.raw`\qc\li720\ri360\fi-360\sb120\sa240 a\par\pard\qj\fi283\sl360\slmult1 b\par\pard\qr\sl-240 c\par\pard\sl300\slmult0 d\par\pard e\par`));
  const [a, b, c, d, e] = paragraphs(model);
  assert.deepEqual(a.props, { align: 'center', indent: { left: 720, right: 360, hanging: 360 }, spacing: { before: 120, after: 240 }, widowControl: false });
  assert.equal(b.props.align, 'both');
  assert.deepEqual(b.props.indent, { firstLine: 283 });
  assert.deepEqual(b.props.spacing, { line: 360, lineRule: 'auto' });
  assert.equal(c.props.align, 'right');
  assert.deepEqual(c.props.spacing, { line: 240, lineRule: 'exact' });
  assert.deepEqual(d.props.spacing, { line: 300, lineRule: 'atLeast' });
  assert.deepEqual(e.props, { widowControl: false });
  const effective = effectivePara(model, e);
  assert.equal(effective.align, 'left');
  assert.deepEqual(effective.spacing, { before: 0, after: 0, line: 240, lineRule: 'auto' });
});

test('widow control is what the paragraph says, else what the document says', () => {
  const off = paragraphs(read(doc(String.raw`a\par\widctlpar b\par\pard\nowidctlpar c\par`)));
  assert.deepEqual([off[0].props.widowControl, off[1].props.widowControl, off[2].props.widowControl], [false, undefined, false]);
  const on = paragraphs(read(doc(String.raw`a\par\widctlpar b\par\pard\nowidctlpar c\par`, '\\widowctrl')));
  assert.deepEqual([on[0].props.widowControl, on[1].props.widowControl, on[2].props.widowControl], [undefined, undefined, false]);
});

test('tab stops with their kinds and leaders', () => {
  const model = read(doc(String.raw`\tx720\tqr\tldot\tx9000\tqc\tx4500\tqdec\tlhyph\tx6000\tb1000 a\par`));
  assert.deepEqual(first(model).props.tabs, [
    { pos: 720, align: 'left' },
    { pos: 1000, align: 'bar' },
    { pos: 4500, align: 'center' },
    { pos: 6000, align: 'decimal', leader: 'hyphen' },
    { pos: 9000, align: 'right', leader: 'dot' },
  ]);
});

test('keeping together, page break before, outline level and the other switches', () => {
  const model = read(doc(String.raw`\keep\keepn\pagebb\outlinelevel2\contextualspace\rtlpar\noline\hyphpar0 a\par\pard b\par`));
  const [a, b] = paragraphs(model);
  assert.equal(a.props.keepLines, true);
  assert.equal(a.props.keepNext, true);
  assert.equal(a.props.pageBreakBefore, true);
  assert.equal(a.props.outlineLevel, 2);
  assert.equal(a.props.contextualSpacing, true);
  assert.equal(a.props.bidi, true);
  assert.equal(a.props.suppressLineNumbers, true);
  assert.equal(a.props.suppressAutoHyphens, true);
  assert.deepEqual(b.props, { widowControl: false });
});

test('paragraph borders and shading', () => {
  const model = read(doc(String.raw`\brdrt\brdrs\brdrw15\brsp20\brdrcf1 \brdrb\brdrdb\brdrw30 \brdrbtw\brdrdot \shading2000\cfpat1\cbpat2 a\par\pard\box\brdrs\brdrw10 b\par\pard\cbpat3 c\par`));
  const [a, b, c] = paragraphs(model);
  assert.deepEqual(a.props.borders, {
    top: { style: 'single', size: 6, space: 1, color: { rgb: 'FF0000' } },
    bottom: { style: 'double', size: 12 },
    between: { style: 'dotted' },
  });
  assert.deepEqual(a.props.shading, { pattern: 'pct20', fill: { rgb: '0000FF' }, color: { rgb: 'FF0000' } });
  const box = { style: 'single', size: 4 };
  assert.deepEqual(b.props.borders, { top: box, left: box, bottom: box, right: box });
  assert.deepEqual(c.props.shading, { pattern: 'clear', fill: { rgb: 'FFFF00' } });
});

test('"\\pard" sets the paragraph formatting back and leaves the character formatting', () => {
  const model = read(doc(String.raw`\qc\li720\b a\par\pard b\par`));
  const [, b] = paragraphs(model);
  assert.deepEqual(b.props, { widowControl: false, markProps: { bold: true } });
  assert.equal(runWith(b, 'b').props.bold, true);
});

test('the paragraph has the formatting in effect at its mark, and the mark the character formatting', () => {
  const model = read(doc(String.raw`a\qr\b\par`));
  const p = first(model);
  assert.equal(p.props.align, 'right');
  assert.deepEqual(p.props.markProps, { bold: true });
  assert.deepEqual(runWith(p, 'a').props, {});
});

// A style sheet as Word writes it: every style states its formatting in full.
const SHEET =
  String.raw`{\stylesheet{\ql\sa200\f0\fs22 \snext0 \sqformat \spriority0 Normal;}` +
  String.raw`{\s1\ql\sb240\sa60\keepn\b\f1\fs32 \sbasedon0 \snext0 \slink15 \sqformat \spriority9 heading 1;}` +
  String.raw`{\s2\qc\sb240\sa60\keepn\b\i\f1\fs28 \sbasedon1 \snext0 heading 2;}` +
  String.raw`{\*\cs10 \additive \ssemihidden Default Paragraph Font;}` +
  String.raw`{\*\cs15 \additive \b\fs32 \sbasedon10 \slink1 Heading 1 Char;}` +
  String.raw`{\*\cs16 \additive \i\cf1 My Emphasis;}` +
  String.raw`{\*\ts11\tsrowd\trftsWidthB3\trpaddl108\trpaddr108\trpaddfl3\trpaddfr3 \snext11 \ssemihidden Normal Table;}` +
  String.raw`{\s20\ql\li720\sa200\f0\fs22 \sbasedon0 My Own Style;}}`;

test('styles get the ids and names a docx file gives them', () => {
  const model = read(doc('a\\par', SHEET));
  const styles = model.styles.styles;
  assert.deepEqual(Array.from(styles.keys()), ['Normal', 'Heading1', 'Heading2', 'DefaultParagraphFont', 'Heading1Char', 'MyEmphasis', 'TableNormal', 'MyOwnStyle']);
  const heading = styles.get('Heading1')!;
  assert.equal(heading.name, 'heading 1');
  assert.equal(heading.type, 'paragraph');
  assert.equal(heading.basedOn, 'Normal');
  assert.equal(heading.next, 'Normal');
  assert.equal(heading.link, 'Heading1Char');
  assert.equal(heading.quickFormat, true);
  assert.equal(heading.uiPriority, 9);
  assert.equal(styles.get('Heading2')!.basedOn, 'Heading1');
  assert.equal(styles.get('Heading1Char')!.type, 'character');
  assert.equal(styles.get('Heading1Char')!.basedOn, 'DefaultParagraphFont');
  assert.equal(styles.get('TableNormal')!.type, 'table');
  assert.equal(styles.get('DefaultParagraphFont')!.semiHidden, true);
  assert.deepEqual(model.styles.defaultIds, { paragraph: 'Normal', character: 'DefaultParagraphFont' });
  assert.equal(styles.get('Normal')!.isDefault, true);
});

test('a style states what it adds to the style it is based on', () => {
  const model = read(doc('a\\par', SHEET));
  const styles = model.styles.styles;
  assert.deepEqual(styles.get('Normal')!.para, { widowControl: false, spacing: { after: 200 } });
  assert.deepEqual(styles.get('Normal')!.run, { size: 22, sizeCs: 22 });
  assert.deepEqual(styles.get('Heading1')!.para, { keepNext: true, spacing: { before: 240, after: 60 } });
  assert.deepEqual(styles.get('Heading1')!.run, { bold: true, size: 32, sizeCs: 32, fonts: { ascii: 'Arial', hAnsi: 'Arial', eastAsia: 'Arial', cs: 'Arial' } });
  assert.deepEqual(styles.get('Heading2')!.para, { align: 'center' });
  assert.deepEqual(styles.get('Heading2')!.run, { italic: true, size: 28, sizeCs: 28 });
  assert.deepEqual(styles.get('MyEmphasis')!.run, { italic: true, color: { rgb: 'FF0000' } });
  assert.deepEqual(model.styles.defaults.run, { size: 24, sizeCs: 24, fonts: { ascii: 'Times New Roman', hAnsi: 'Times New Roman', eastAsia: 'Times New Roman', cs: 'Times New Roman' } });
});

test('a paragraph names its style and states only what differs from it', () => {
  const body = String.raw`\pard\plain\s1\ql\sb240\sa60\keepn\b\f1\fs32 Title\par\pard\plain\ql\sa200\f0\fs22 Text {\b bold}\par\pard\plain\s2\qc\sb240\sa60\keepn\b\i\f1\fs28 Sub\par`;
  const model = read(doc(body, SHEET));
  const [title, text, sub] = paragraphs(model);
  assert.deepEqual(title.props, { style: 'Heading1' });
  assert.deepEqual(runWith(title, 'Title').props, {});
  assert.deepEqual(text.props, {});
  assert.deepEqual(runWith(text, 'bold').props, { bold: true });
  assert.deepEqual(sub.props, { style: 'Heading2' });

  const run = effectiveRun(model, sub, 'Sub');
  assert.equal(run.bold, true);
  assert.equal(run.italic, true);
  assert.equal(run.size, 28);
  assert.equal(run.fonts!.ascii, 'Arial');
  const para = resolveParaProps(model, sub.props);
  assert.equal(para.align, 'center');
  assert.equal(para.keepNext, true);
  assert.deepEqual(para.spacing, { before: 240, after: 60, line: 240, lineRule: 'auto' });
  assert.equal(effectiveRun(model, text, 'Text').size, 22);
});

test('formatting the style has and the paragraph does not is switched off', () => {
  // A heading whose text is not bold, not kept with the next paragraph and has no space after it.
  const body = String.raw`\pard\plain\s1\ql\sb240\f1\fs32 Light\par`;
  const model = read(doc(body, SHEET));
  const p = first(model);
  assert.equal(p.props.style, 'Heading1');
  assert.equal(p.props.keepNext, false);
  assert.deepEqual(p.props.spacing, { after: 0 });
  assert.deepEqual(runWith(p, 'Light').props, { bold: false });
  assert.equal(effectiveRun(model, p, 'Light').bold, false);
  assert.equal(effectivePara(model, p).keepNext, false);
  assert.equal(effectivePara(model, p).spacing!.after, 0);
  assert.equal(effectivePara(model, p).spacing!.before, 240);
});

test('a character style is named by the run; what it brings is not stated again', () => {
  const body = String.raw`\pard\plain\ql\sa200\f0\fs22 a{\cs16\i\cf1 b}{\cs16\i\cf2 c}\par`;
  const model = read(doc(body, SHEET));
  const p = first(model);
  assert.deepEqual(runWith(p, 'b').props, { style: 'MyEmphasis' });
  assert.deepEqual(runWith(p, 'c').props, { style: 'MyEmphasis', color: { rgb: '0000FF' } });
  assert.equal(effectiveRun(model, p, 'b').italic, true);
  assert.equal(resolveColor(model, effectiveRun(model, p, 'c').color), '#0000FF');
});

test('the style of a paragraph may be stated after its text', () => {
  const body = String.raw`\pard\plain\ql\sa200\b\f1\fs32 Late\s1\sb240\sa60\keepn\par`;
  const model = read(doc(body, SHEET));
  const p = first(model);
  assert.equal(p.props.style, 'Heading1');
  // The run was made before the style was known; it states nothing the style gives.
  assert.deepEqual(runWith(p, 'Late').props, {});
  assert.equal(effectiveRun(model, p, 'Late').bold, true);
});

test('a document without a style sheet states everything directly', () => {
  const model = read(doc(String.raw`\pard\plain\qc\b\fs28 a\par`));
  const p = first(model);
  assert.equal(model.styles.styles.size, 0);
  assert.equal(effectiveRun(model, p, 'a').bold, true);
  assert.equal(effectiveRun(model, p, 'a').size, 28);
  assert.equal(effectivePara(model, p).align, 'center');
});

test('styles based on each other in a circle, on themselves or on nothing do not hang the reader', () => {
  const sheet = String.raw`{\stylesheet{\s1\b\sbasedon2 A;}{\s2\i\sbasedon1 B;}{\s3\sbasedon3 C;}{\s4\sbasedon99 D;}{\s1 Twice;}}`;
  const model = read(doc(String.raw`\pard\plain\s1\b a\par\pard\plain\s2\i b\par\pard\plain\s4 c\par`, sheet));
  assert.equal(model.styles.styles.size, 4);
  const [a, b] = paragraphs(model);
  assert.equal(effectiveRun(model, a, 'a').bold, true);
  assert.ok(!effectiveRun(model, a, 'a').italic);
  assert.equal(effectiveRun(model, b, 'b').italic, true);
  assert.ok(!effectiveRun(model, b, 'b').bold);
});

test('the digits Word writes after a byte in the name of a style are not part of the name', () => {
  const sheet = '{\\stylesheet{\\s3 \\sbasedon0 \\' + "'dcdcst Bilgi \\'fdFDx;}{\\s4 Plain dc;}}";
  const model = read('{\\rtf1\\ansi\\ansicpg1254' + sheet + 'a\\par}');
  const names = Array.from(model.styles.styles.values()).map((style) => style.name);
  assert.deepEqual(names, [String.fromCharCode(0xdc) + 'st Bilgi ' + String.fromCharCode(0x131) + 'x', 'Plain dc']);
});

test('style names that are names of object members are styles like any other', () => {
  const sheet = String.raw`{\stylesheet{\s1\b __proto__;}{\s2\i constructor;}{\s3 toString;}}`;
  const model = read(doc(String.raw`\pard\plain\s1\b a\par\pard\plain\s2\i b\par`, sheet));
  assert.deepEqual(Array.from(model.styles.styles.keys()), ['Proto', 'Constructor', 'toString']);
  assert.equal(paragraphs(model)[1].props.style, 'Constructor');
  assert.equal(({} as { bold?: boolean }).bold, undefined);
});

// A list table as Word writes it: a numbered list of two levels and a list of bullets in a font of symbols.
const LISTS =
  String.raw`{\*\listtable{\list\listtemplateid1\listhybrid` +
  '{\\listlevel\\levelnfc0\\levelnfcn0\\leveljc0\\leveljcn0\\levelfollow0\\levelstartat1{\\leveltext\\' + "'02\\'00.;}{\\levelnumbers\\'01;}\\fi-360\\li720\\jclisttab\\tx720 }" +
  '{\\listlevel\\levelnfc4\\levelnfcn4\\leveljc0\\leveljcn0\\levelfollow1\\levelstartat3{\\leveltext\\' + "'04\\'00.\\'01);}{\\levelnumbers\\'01\\'03;}\\b\\fi-360\\li1440 }" +
  String.raw`{\listname ;}\listid100}` +
  '{\\list\\listtemplateid2\\listsimple{\\listlevel\\levelnfc23\\levelnfcn23\\leveljc0\\levelfollow0\\levelstartat1{\\leveltext\\' + "'01\\u-3913 ?;}{\\levelnumbers;}\\f2\\fi-360\\li720 }" +
  String.raw`{\listname ;}\listid200}` +
  '{\\list\\listsimple{\\listlevel\\levelnfc23\\levelfollow2{\\leveltext\\' + "'01\\'b7;}{\\levelnumbers;}\\f2 }\\listid300}}" +
  String.raw`{\*\listoverridetable{\listoverride\listid100\listoverridecount0\ls1}{\listoverride\listid200\listoverridecount0\ls2}` +
  String.raw`{\listoverride\listid100\listoverridecount1{\lfolevel\listoverridestartat\levelstartat7}\ls3}{\listoverride\listid300\ls4}{\listoverride\listid999\ls5}}`;

test('lists: definitions, levels and the instances paragraphs refer to', () => {
  const model = read(doc('a\\par', LISTS));
  const { abstracts, nums } = model.numbering;
  assert.equal(abstracts.size, 3);
  const levels = abstracts.get(0)!.levels;
  assert.equal(abstracts.get(0)!.multiLevelType, 'hybridMultilevel');
  assert.deepEqual(levels.get(0), {
    level: 0,
    start: 1,
    format: 'decimal',
    text: '%1.',
    suffix: 'tab',
    align: 'left',
    para: { indent: { left: 720, hanging: 360 }, tabs: [{ pos: 720, align: 'left' }] },
  });
  assert.deepEqual(levels.get(1), { level: 1, start: 3, format: 'lowerLetter', text: '%1.%2)', suffix: 'space', align: 'left', para: { indent: { left: 1440, hanging: 360 } }, run: { bold: true } });
  const bullet = abstracts.get(1)!.levels.get(0)!;
  assert.equal(abstracts.get(1)!.multiLevelType, 'singleLevel');
  assert.equal(bullet.format, 'bullet');
  assert.equal(bullet.text, String.fromCharCode(0xf0b7));
  assert.deepEqual(bullet.run, { fonts: { ascii: 'Symbol', hAnsi: 'Symbol', eastAsia: 'Symbol', cs: 'Symbol' } });
  // A bullet given as a byte is a character of the font of the level.
  assert.equal(abstracts.get(2)!.levels.get(0)!.text, String.fromCharCode(0xf0b7));
  assert.equal(abstracts.get(2)!.levels.get(0)!.suffix, 'nothing');
  // An instance of a list the file lacks is left out.
  assert.deepEqual(Array.from(nums.keys()), [1, 2, 3, 4]);
  assert.equal(nums.get(2)!.abstractId, 1);
  assert.deepEqual(nums.get(3)!.overrides.get(0), { start: 7 });
});

test('numbered paragraphs: the text Word writes for old readers is left out and the numbers count', () => {
  const body =
    String.raw`{\listtext\pard\plain 1.\tab}\pard\plain\fi-360\li720\ls1 One\par` +
    String.raw`{\listtext\pard\plain 2.\tab}\pard\plain\fi-360\li720\ls1 Two\par` +
    String.raw`{\listtext\pard\plain 1.c)\tab}\pard\plain\fi-360\li1440\ls1\ilvl1 Deep\par` +
    String.raw`{\listtext\pard\plain\f2 \'b7\tab}\pard\plain\fi-360\li720\ls2 Dot\par` +
    String.raw`{\listtext\pard\plain 7.\tab}\pard\plain\fi-360\li720\ls3 Seven\par`;
  const model = read(doc(body, LISTS));
  assert.deepEqual(texts(model), ['One', 'Two', 'Deep', 'Dot', 'Seven']);
  const ps = paragraphs(model);
  assert.deepEqual(ps[0].props.numbering, { numId: 1, level: 0 });
  assert.deepEqual(ps[2].props.numbering, { numId: 1, level: 1 });
  const counters = new NumberingCounters(model.numbering, model.styles);
  const labels = ps.map((p) => counters.next(p.props.numbering!.numId!, p.props.numbering!.level!)!.text);
  assert.deepEqual(labels, ['1.', '2.', '2.c)', String.fromCharCode(0xf0b7), '7.']);
  // The indents of the paragraph are its own.
  assert.deepEqual(effectivePara(model, ps[2]).indent, { left: 1440, hanging: 360 });
});

test('the number text shows where the file has no list for the paragraph', () => {
  const body =
    String.raw`{\listtext\pard\plain 4.\tab}\pard\plain\ls9 Unknown list\par` +
    String.raw`{\pntext\pard\plain\b a)\tab}\pard\plain{\*\pn\pnlvlbody\pnlcltr\pnstart1{\pntxta )}}Old numbering\par` +
    String.raw`{\listtext\pard\plain x\tab}\pard\plain\ls5 No such list\par`;
  const model = read(doc(body, LISTS));
  assert.deepEqual(texts(model), ['4.\tUnknown list', 'a)\tOld numbering', 'x\tNo such list']);
  assert.equal(runsOf(paragraphs(model)[1])[0].props.bold, true);
});

test('paper, margins and the settings of the document', () => {
  const head = String.raw`\paperw11906\paperh16838\margl1417\margr1418\margt1134\margb850\gutter200\deftab708\facingp\margmirror\hyphauto\widowctrl\revisions`;
  const model = read(doc('a\\par', head));
  assert.deepEqual(model.body.section, {
    page: { width: 11906, height: 16838 },
    margins: { top: 1134, right: 1418, bottom: 850, left: 1417, header: 720, footer: 720, gutter: 200 },
  });
  assert.equal(model.settings.defaultTabStop, 708);
  assert.equal(model.settings.evenAndOddHeaders, true);
  assert.equal(model.settings.mirrorMargins, true);
  assert.equal(model.settings.autoHyphenation, true);
  assert.equal(model.settings.trackRevisions, true);
  assert.equal(model.settings.compatSettings.get('compatibilityMode'), '11');
});

test('a document that states nothing has the paper and margins RTF assumes', () => {
  assert.deepEqual(read('{\\rtf1 a}').body.section, {
    page: { width: 12240, height: 15840 },
    margins: { top: 1440, right: 1800, bottom: 1440, left: 1800, header: 720, footer: 720, gutter: 0 },
  });
});

test('"\\sect" ends a section: its last paragraph carries the properties', () => {
  const head = String.raw`\paperw12240\paperh15840\margl1000\margr1000\margt1000\margb1000`;
  const body =
    String.raw`\sectd\headery500\footery600\titlepg One\par Two\sect ` +
    String.raw`\sectd\sbknone\lndscpsxn\pgwsxn15840\pghsxn12240\marglsxn500\cols2\colsx400\linebetcol\pgnrestart\pgnstarts5\pgnlcrm Three\par\sect ` +
    String.raw`\sectd\sbkodd\vertalc Four\par`;
  const model = read(doc(body, head));
  const sections = findSections(model);
  assert.equal(sections.length, 3);
  const ps = paragraphs(model);
  assert.deepEqual(ps.map(paragraphText), ['One', 'Two', 'Three', '', 'Four']);
  assert.equal(ps[0].props.section, undefined);
  assert.deepEqual(ps[1].props.section, {
    page: { width: 12240, height: 15840 },
    margins: { top: 1000, right: 1000, bottom: 1000, left: 1000, header: 500, footer: 600, gutter: 0 },
    titlePage: true,
  });
  assert.deepEqual(ps[3].props.section, {
    type: 'continuous',
    page: { width: 15840, height: 12240, orientation: 'landscape' },
    margins: { top: 1000, right: 1000, bottom: 1000, left: 500, header: 720, footer: 720, gutter: 0 },
    columns: { count: 2, space: 400, separator: true },
    pageNumbering: { format: 'lowerRoman', start: 5 },
  });
  // The last section is that of the body; "\sectd" has set it back to what the document says.
  assert.deepEqual(model.body.section, {
    type: 'oddPage',
    page: { width: 12240, height: 15840 },
    margins: { top: 1000, right: 1000, bottom: 1000, left: 1000, header: 720, footer: 720, gutter: 0 },
    vAlign: 'center',
  });
});

test('columns of their own widths', () => {
  const model = read(doc(String.raw`\sectd\cols2\colno1\colw3000\colsr500\colno2\colw5000 a\par`));
  assert.deepEqual(model.body.section!.columns, { count: 2, space: 720, equalWidth: false, columns: [{ width: 3000, space: 500 }, { width: 5000 }] });
});

test('the font table: names, alternative names, families, and entries without groups', () => {
  const source =
    String.raw`{\rtf1\ansi\deff1{\fonttbl{\f0\froman\fcharset162\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times};}` +
    String.raw`{\f1\fswiss\fprq2 Arial;}{\f5\fmodern\fprq1 Courier New}{\flomajor\f31500\fbidi \froman\fcharset0 Cambria;}}\f5 a\plain b\par}`;
  const model = read(source);
  assert.deepEqual(model.fonts.get('Times New Roman'), { name: 'Times New Roman', altName: 'Times', panose: '02020603050405020304', charset: 'A2', family: 'roman', pitch: 'variable' });
  assert.deepEqual(model.fonts.get('Arial'), { name: 'Arial', family: 'swiss', pitch: 'variable' });
  // An entry whose semicolon is missing ends with its group.
  assert.equal(model.fonts.get('Courier New')!.pitch, 'fixed');
  assert.equal(model.fonts.get('Cambria')!.family, 'roman');
  const p = first(model);
  assert.equal(effectiveRun(model, p, 'a').fonts!.ascii, 'Courier New');
  // "\deff1": the default font is Arial.
  assert.equal(effectiveRun(model, p, 'b').fonts!.ascii, 'Arial');
  const flat = read(String.raw`{\rtf1{\fonttbl\f0\fswiss Helvetica;\f1\fmodern Courier;}\f1 a}`);
  assert.deepEqual(Array.from(flat.fonts.keys()), ['Helvetica', 'Courier']);
  assert.equal(effectiveRun(flat, first(flat), 'a').fonts!.ascii, 'Courier');
});

test('a paragraph placed by itself: a frame', () => {
  const model = read(doc(String.raw`\pard\phmrg\pvpara\posxr\posnegy-300\absw2000\absh-400\dxfrtext180\dfrmtxtx141\wraparound Framed\par\pard\phpg\pvpg\posx720\posyb\absh500\dropcapli3\dropcapt1 D\par\pard x\par`));
  const [a, b, c] = paragraphs(model);
  assert.deepEqual(a.props.frame, { hAnchor: 'margin', vAnchor: 'text', xAlign: 'right', y: -300, width: 2000, height: 400, heightRule: 'exact', hSpace: 141, vSpace: 180, wrap: 'around' });
  assert.deepEqual(b.props.frame, { hAnchor: 'page', vAnchor: 'page', x: 720, yAlign: 'bottom', height: 500, heightRule: 'atLeast', lines: 3, dropCap: 'drop' });
  assert.equal(c.props.frame, undefined);
});

test('borders of the page and of characters', () => {
  const model = read(doc(String.raw`\sectd\pgbrdrt\brdrs\brdrw20\brsp480\brdrcf1 \pgbrdrb\brdrdb\brdrw10 {\chbrdr\brdrs\brdrw10\brdrcf2 boxed} plain\par`));
  assert.deepEqual(model.body.section!.pageBorders, {
    top: { style: 'single', size: 8, space: 24, color: { rgb: 'FF0000' } },
    bottom: { style: 'double', size: 4 },
  });
  const p = first(model);
  assert.deepEqual(runWith(p, 'boxed').props.border, { style: 'single', size: 4, color: { rgb: '0000FF' } });
  assert.equal(runWith(p, 'plain').props.border, undefined);
});

test('a style sheet after the first text still gives the paragraphs before it their styles', () => {
  const source = String.raw`{\rtf1\ansi\deff0{\fonttbl{\f0 Times New Roman;}}\pard\plain\s1\b early\par{\stylesheet{\s1\b Strong Para;}}\pard\plain\s1\b late\par}`;
  const model = read(source);
  const [early, late] = paragraphs(model);
  assert.equal(late.props.style, 'StrongPara');
  assert.deepEqual(runWith(late, 'late').props, {});
  // The first paragraph was read before the style existed: it states its formatting itself.
  assert.equal(early.props.style, undefined);
  assert.equal(effectiveRun(model, early, 'early').bold, true);
});
