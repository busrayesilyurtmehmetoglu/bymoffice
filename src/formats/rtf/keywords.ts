// The control words of an RTF file that change formatting or stand for a
// character, each with what it does to the reader.
//
// A control word with a number takes the number as its value; one that
// switches something on does so without a number or with any number but 0
// ("\b" and "\b1" are bold, "\b0" is not).
//
// The words in IGNORED are known and deliberately passed over: revision
// numbers, hints for other programs, what the reader has no model for. They
// are listed so that the count of words the reader does not know (see
// `unknown` of the model) says something.

import type { RtfReader } from './document';
import { dopWord, sepWord, defaultSep } from './sections';
import { BORDER_STYLES, UNDERLINES, lookup, twipsOf, type Chp, type Pap } from './state';

export type WordHandler = (reader: RtfReader, param: number, hasParam: boolean, word: string) => void;

export const WORDS = new Map<string, WordHandler>();

type Flags<T> = { [K in keyof T]: T[K] extends boolean ? K : never }[keyof T];
type Numbers<T> = { [K in keyof T]: T[K] extends number ? K : never }[keyof T];

const isOn = (param: number, hasParam: boolean): boolean => !hasParam || param !== 0;

// ----------------------------------------------------------------- characters

/** A switch of the character formatting; `complex` names the one for text of complex scripts, which "\fcs1" selects. */
function charFlag(word: string, field: Flags<Chp>, complex?: Flags<Chp>): void {
  WORDS.set(word, (r, param, hasParam) => {
    const key = complex !== undefined && r.st.complex ? complex : field;
    const value = isOn(param, hasParam);
    if (r.st.chp[key] !== value) r.chpW()[key] = value;
  });
}

function charNumber(word: string, field: Numbers<Chp>, convert?: (param: number) => number): void {
  WORDS.set(word, (r, param) => {
    const value = convert ? convert(param) : param;
    if (r.st.chp[field] !== value) r.chpW()[field] = value;
  });
}

charFlag('b', 'bold', 'boldCs');
charFlag('ab', 'bold', 'boldCs');
charFlag('i', 'italic', 'italicCs');
charFlag('ai', 'italic', 'italicCs');
charFlag('caps', 'caps');
charFlag('scaps', 'smallCaps');
charFlag('strike', 'strike');
charFlag('striked', 'doubleStrike');
charFlag('outl', 'outline');
charFlag('shad', 'shadow');
charFlag('embo', 'emboss');
charFlag('impr', 'imprint');
charFlag('v', 'hidden');
charFlag('noproof', 'noProof');
charFlag('deleted', 'deleted');
charFlag('revised', 'inserted');
// Who made a tracked change and when. These say nothing about how text looks, and the reader keeps them apart from
// the formatting: see RtfReader.wrapper().
charNumber('revauth', 'revAuth');
charNumber('revauthdel', 'revAuthDel');
charNumber('revdttm', 'revDttm');
charNumber('revdttmdel', 'revDttmDel');

const halfPoints = (param: number): number => (param < 1 ? 1 : param > 3276 ? 3276 : param);

function sizeWord(r: RtfReader, param: number): void {
  const field = r.st.complex ? 'sizeCs' : 'size';
  const value = halfPoints(param);
  if (r.st.chp[field] !== value) r.chpW()[field] = value;
}
WORDS.set('fs', sizeWord);
WORDS.set('afs', sizeWord);

/**
 * "\fN" and "\afN". "\fN" is the font of the text: it replaces every font stated before it, the associated ones
 * included. "\loch", "\hich" and "\dbch" name a kind of text (low and high half of the code page, double-byte);
 * the font word right after one of them is about that kind alone: "\afN" is its associated font, "\fN" (which
 * Word writes for the kind the text at hand is of, after the associated fonts of the others) the font of the text,
 * which that kind follows from then on.
 */
function fontWord(r: RtfReader, param: number, _hasParam: boolean, word: string): void {
  const st = r.st;
  const value = param < 0 ? -1 : param;
  const slot = st.slot;
  st.slot = 0;
  if (word === 'af' && slot) {
    const field = slot === 1 ? 'fontLo' : slot === 2 ? 'fontHi' : 'fontDb';
    if (st.chp[field] !== value) r.chpW()[field] = value;
    return;
  }
  if (st.complex) {
    if (st.chp.fontCs !== value) r.chpW().fontCs = value;
    return;
  }
  const chp = st.chp;
  if (word === 'af' || slot) {
    // The font of the text; the kind named follows it.
    const field = slot === 1 ? 'fontLo' : slot === 2 ? 'fontHi' : slot === 3 ? 'fontDb' : undefined;
    if (chp.font !== value || (field !== undefined && chp[field] !== -1)) {
      const next = r.chpW();
      next.font = value;
      if (field !== undefined) next[field] = -1;
    }
    return;
  }
  if (chp.font !== value || chp.fontLo !== -1 || chp.fontHi !== -1 || chp.fontDb !== -1) {
    const next = r.chpW();
    next.font = value;
    next.fontLo = next.fontHi = next.fontDb = -1;
  }
}
WORDS.set('f', fontWord);
WORDS.set('af', fontWord);

function langWord(r: RtfReader, param: number): void {
  const field = r.st.complex ? 'langCs' : 'lang';
  // 1024 is "no language": the text is not checked, which says nothing about what it is written in.
  if (param === 1024) return;
  if (r.st.chp[field] !== param) r.chpW()[field] = param;
}
WORDS.set('lang', langWord);
WORDS.set('alang', langWord);
WORDS.set('langnp', langWord);
WORDS.set('langfe', (r, param) => {
  if (param !== 1024 && r.st.chp.langFe !== param) r.chpW().langFe = param;
});
WORDS.set('langfenp', WORDS.get('langfe')!);

for (const word of Object.keys(UNDERLINES)) {
  WORDS.set(word, (r, param, hasParam) => {
    const value = isOn(param, hasParam) ? UNDERLINES[word] : 'none';
    if (r.st.chp.underline !== value) r.chpW().underline = value;
  });
}
WORDS.set('ulnone', (r) => {
  if (r.st.chp.underline !== 'none') r.chpW().underline = 'none';
});

const index = (param: number): number => (param < 0 ? 0 : param);
charNumber('ulc', 'underlineColor', index);
charNumber('cf', 'color', index);
charNumber('cb', 'background', index);
charNumber('chcbpat', 'background', index);
charNumber('chcfpat', 'patternColor', index);
charNumber('chshdng', 'shading', (param) => (param < 0 ? 0 : param > 10000 ? 10000 : param));
charNumber('highlight', 'highlight', index);
charNumber('expndtw', 'spacing', twipsOf);
// Quarter points.
charNumber('expnd', 'spacing', (param) => twipsOf(param * 5));
charNumber('charscalex', 'scale', (param) => (param < 1 ? 100 : param > 600 ? 600 : param));
charNumber('kerning', 'kerning', (param) => (param < 0 ? 0 : param > 3276 ? 3276 : param));
charNumber('cs', 'style');

WORDS.set('super', (r, param, hasParam) => {
  const value = isOn(param, hasParam) ? 1 : 0;
  if (r.st.chp.vertical !== value) r.chpW().vertical = value;
});
WORDS.set('sub', (r, param, hasParam) => {
  const value = isOn(param, hasParam) ? 2 : 0;
  if (r.st.chp.vertical !== value) r.chpW().vertical = value;
});
WORDS.set('nosupersub', (r) => {
  if (r.st.chp.vertical !== 0) r.chpW().vertical = 0;
});
// Raised and lowered text; three points when the word gives no number.
WORDS.set('up', (r, param, hasParam) => {
  const value = hasParam ? Math.min(Math.max(param, 0), 3276) : 6;
  if (r.st.chp.position !== value) r.chpW().position = value;
});
WORDS.set('dn', (r, param, hasParam) => {
  const value = -(hasParam ? Math.min(Math.max(param, 0), 3276) : 6);
  if (r.st.chp.position !== value) r.chpW().position = value;
});

WORDS.set('plain', (r) => r.plain());
WORDS.set('rtlch', (r) => {
  r.st.rtl = true;
});
WORDS.set('ltrch', (r) => {
  r.st.rtl = false;
});
WORDS.set('fcs', (r, param) => {
  r.st.complex = param === 1;
  r.st.assoc = 0;
  r.st.slot = 0;
});
WORDS.set('loch', (r) => {
  r.st.assoc = r.st.slot = 1;
});
WORDS.set('hich', (r) => {
  r.st.assoc = r.st.slot = 2;
});
WORDS.set('dbch', (r) => {
  r.st.assoc = r.st.slot = 3;
});
WORDS.set('uc', (r, param) => r.setUc(param));
WORDS.set('u', (r, param) => r.unicode(param));

// ----------------------------------------------------------------- paragraphs

function paraFlag(word: string, field: Flags<Pap>, value?: boolean): void {
  WORDS.set(word, (r, param, hasParam) => {
    const next = value === undefined ? isOn(param, hasParam) : value;
    if (r.st.pap[field] !== next) r.papW()[field] = next;
  });
}

function paraNumber(word: string, field: Numbers<Pap>, convert?: (param: number) => number): void {
  WORDS.set(word, (r, param) => {
    const value = convert ? convert(param) : param;
    if (r.st.pap[field] !== value) r.papW()[field] = value;
  });
}

function paraAlign(word: string, align: string): void {
  WORDS.set(word, (r) => {
    if (r.st.pap.align !== align) r.papW().align = align;
  });
}

WORDS.set('pard', (r) => r.pard());
paraNumber('s', 'style', index);
paraAlign('ql', 'left');
paraAlign('qr', 'right');
paraAlign('qc', 'center');
paraAlign('qj', 'both');
paraAlign('qd', 'distribute');
paraAlign('qt', 'thaiDistribute');
paraNumber('li', 'left', twipsOf);
paraNumber('lin', 'left', twipsOf);
paraNumber('ri', 'right', twipsOf);
paraNumber('rin', 'right', twipsOf);
paraNumber('fi', 'first', twipsOf);
paraNumber('sb', 'before', (param) => Math.max(0, twipsOf(param)));
paraNumber('sa', 'after', (param) => Math.max(0, twipsOf(param)));
paraFlag('sbauto', 'beforeAuto');
paraFlag('saauto', 'afterAuto');
paraNumber('sl', 'line', (param) => Math.max(-31680, Math.min(param, 31680)));
paraFlag('slmult', 'lineMultiple');
paraFlag('keep', 'keep');
paraFlag('keepn', 'keepNext');
paraFlag('pagebb', 'pageBreakBefore');
paraFlag('noline', 'noLineNumbers');
paraFlag('contextualspace', 'contextual');
paraFlag('rtlpar', 'rtl', true);
paraFlag('ltrpar', 'rtl', false);
paraFlag('intbl', 'inTable');
paraNumber('itap', 'depth', (param) => (param < 0 ? 0 : param > 1000 ? 1000 : param));
paraNumber('ls', 'list', (param) => (param < 0 ? 0 : param));
paraNumber('ilvl', 'level', (param) => (param < 0 ? 0 : param > 8 ? 8 : param));
paraNumber('outlinelevel', 'outline', (param) => (param >= 0 && param <= 8 ? param : -1));
paraNumber('shading', 'shading', (param) => (param < 0 ? 0 : param > 10000 ? 10000 : param));
paraNumber('cfpat', 'patternColor', index);
paraNumber('cbpat', 'background', index);
WORDS.set('widctlpar', (r) => {
  if (r.st.pap.widowControl !== 1) r.papW().widowControl = 1;
});
WORDS.set('nowidctlpar', (r) => {
  if (r.st.pap.widowControl !== 0) r.papW().widowControl = 0;
});
// "\hyphpar0" takes the paragraph out of automatic hyphenation.
WORDS.set('hyphpar', (r, param, hasParam) => {
  const value = hasParam && param === 0;
  if (r.st.pap.noHyphens !== value) r.papW().noHyphens = value;
});

for (const word of ['tx', 'tb', 'tqr', 'tqc', 'tqdec', 'tldot', 'tlhyph', 'tlul', 'tlth', 'tleq', 'tlmdot']) {
  WORDS.set(word, (r, param, _hasParam, name) => {
    r.tabWord(name, param);
  });
}

for (const word of ['brdrt', 'brdrb', 'brdrl', 'brdrr', 'brdrbtw', 'brdrbar', 'box', 'chbrdr', 'pgbrdrt', 'pgbrdrb', 'pgbrdrl', 'pgbrdrr']) {
  WORDS.set(word, (r, _param, _hasParam, name) => {
    r.borderStart(name);
  });
}
for (const word of Object.keys(BORDER_STYLES).concat(['brdrw', 'brsp', 'brdrcf', 'brdrframe', 'brdrtbl'])) {
  WORDS.set(word, (r, param, _hasParam, name) => r.border(name, param));
}

// A paragraph that is placed by itself (a frame).
const frame = (word: string, set: (r: RtfReader, param: number) => void): void => {
  WORDS.set(word, (r, param) => set(r, param));
};
frame('absw', (r, param) => {
  r.frameW().width = Math.abs(twipsOf(param));
});
frame('absh', (r, param) => {
  const f = r.frameW();
  if (param === 0) {
    delete f.height;
    f.heightRule = 'auto';
  } else {
    f.height = Math.abs(twipsOf(param));
    f.heightRule = param < 0 ? 'exact' : 'atLeast';
  }
});
frame('phmrg', (r) => {
  r.frameW().hAnchor = 'margin';
});
frame('phpg', (r) => {
  r.frameW().hAnchor = 'page';
});
frame('phcol', (r) => {
  r.frameW().hAnchor = 'text';
});
frame('pvmrg', (r) => {
  r.frameW().vAnchor = 'margin';
});
frame('pvpg', (r) => {
  r.frameW().vAnchor = 'page';
});
frame('pvpara', (r) => {
  r.frameW().vAnchor = 'text';
});
frame('posx', (r, param) => {
  r.frameW().x = twipsOf(param);
});
frame('posnegx', (r, param) => {
  r.frameW().x = -Math.abs(twipsOf(param));
});
frame('posy', (r, param) => {
  r.frameW().y = twipsOf(param);
});
frame('posnegy', (r, param) => {
  r.frameW().y = -Math.abs(twipsOf(param));
});
const X_ALIGN: { [word: string]: string } = lookup({ posxc: 'center', posxi: 'inside', posxo: 'outside', posxl: 'left', posxr: 'right' });
for (const word of Object.keys(X_ALIGN)) {
  frame(word, (r) => {
    r.frameW().xAlign = X_ALIGN[word];
  });
}
const Y_ALIGN: { [word: string]: string } = lookup({ posyt: 'top', posyc: 'center', posyb: 'bottom', posyil: 'inline', posyin: 'inside', posyout: 'outside' });
for (const word of Object.keys(Y_ALIGN)) {
  frame(word, (r) => {
    r.frameW().yAlign = Y_ALIGN[word];
  });
}
frame('dxfrtext', (r, param) => {
  const f = r.frameW();
  f.hSpace = f.vSpace = Math.max(0, twipsOf(param));
});
frame('dfrmtxtx', (r, param) => {
  r.frameW().hSpace = Math.max(0, twipsOf(param));
});
frame('dfrmtxty', (r, param) => {
  r.frameW().vSpace = Math.max(0, twipsOf(param));
});
frame('wraparound', (r) => {
  r.frameW().wrap = 'around';
});
frame('nowrap', (r) => {
  r.frameW().wrap = 'notBeside';
});
frame('wraptight', (r) => {
  r.frameW().wrap = 'tight';
});
frame('wrapthrough', (r) => {
  r.frameW().wrap = 'through';
});
frame('abslock', (r) => {
  r.frameW().anchorLock = true;
});
frame('dropcapli', (r, param) => {
  r.frameW().lines = Math.max(1, Math.min(param, 10));
});
frame('dropcapt', (r, param) => {
  r.frameW().dropCap = param === 2 ? 'margin' : 'drop';
});

// --------------------------------------------------------------------- tables

const TABLE_WORDS = [
  'trowd',
  'cellx',
  'trleft',
  'trgaph',
  'trrh',
  'trhdr',
  'trkeep',
  'trql',
  'trqc',
  'trqr',
  'rtlrow',
  'ltrrow',
  'trautofit',
  'trftsWidth',
  'trwWidth',
  'trpaddl',
  'trpaddt',
  'trpaddr',
  'trpaddb',
  'trpaddfl',
  'trpaddft',
  'trpaddfr',
  'trpaddfb',
  'trspdl',
  'trspdt',
  'trspdr',
  'trspdb',
  'trspdfl',
  'trspdft',
  'trspdfr',
  'trspdfb',
  'trshdng',
  'trcfpat',
  'trcbpat',
  'trbrdrt',
  'trbrdrl',
  'trbrdrb',
  'trbrdrr',
  'trbrdrh',
  'trbrdrv',
  'tphmrg',
  'tphpg',
  'tphcol',
  'tpvmrg',
  'tpvpg',
  'tpvpara',
  'tposx',
  'tposnegx',
  'tposxc',
  'tposxl',
  'tposxr',
  'tposxi',
  'tposxo',
  'tposy',
  'tposnegy',
  'tposyt',
  'tposyc',
  'tposyb',
  'tposyil',
  'tposyin',
  'tposyout',
  'tdfrmtxtLeft',
  'tdfrmtxtRight',
  'tdfrmtxtTop',
  'tdfrmtxtBottom',
  'tabsnoovrlp',
  'clmgf',
  'clmrg',
  'clvmgf',
  'clvmrg',
  'clvertalt',
  'clvertalc',
  'clvertalb',
  'clcbpat',
  'clcfpat',
  'clshdng',
  'cltxlrtb',
  'cltxtbrl',
  'cltxbtlr',
  'cltxlrtbv',
  'cltxtbrlv',
  'clFitText',
  'clNoWrap',
  'clhidemark',
  'clpadl',
  'clpadt',
  'clpadr',
  'clpadb',
  'clpadfl',
  'clpadft',
  'clpadfr',
  'clpadfb',
  'clbrdrt',
  'clbrdrl',
  'clbrdrb',
  'clbrdrr',
  'cldglu',
  'cldgll',
];
for (const word of TABLE_WORDS) WORDS.set(word, (r, param, hasParam, name) => r.tableWord(name, param, hasParam));

WORDS.set('cell', (r) => r.cell(false));
WORDS.set('nestcell', (r) => r.cell(true));
WORDS.set('row', (r) => r.row(false));
WORDS.set('nestrow', (r) => r.row(true));

// ------------------------------------------------------- document and sections

const DOCUMENT_WORDS = [
  'paperw',
  'paperh',
  'margl',
  'margr',
  'margt',
  'margb',
  'gutter',
  'landscape',
  'facingp',
  'margmirror',
  'deftab',
  'widowctrl',
  'hyphauto',
  'revisions',
  'pgnstart',
  'formprot',
  'revprot',
  'annotprot',
  'readprot',
  'enforceprot',
  'protlevel',
];
for (const word of DOCUMENT_WORDS) {
  WORDS.set(word, (r, param, hasParam, name) => {
    dopWord(r.dop, name, param, hasParam);
  });
}

const SECTION_WORDS = [
  'pgwsxn',
  'pghsxn',
  'marglsxn',
  'margrsxn',
  'margtsxn',
  'margbsxn',
  'guttersxn',
  'lndscpsxn',
  'headery',
  'footery',
  'cols',
  'colsx',
  'linebetcol',
  'colno',
  'colw',
  'colsr',
  'sbknone',
  'sbkcol',
  'sbkpage',
  'sbkeven',
  'sbkodd',
  'pgnstarts',
  'pgnrestart',
  'pgncont',
  'pgndec',
  'pgnucrm',
  'pgnlcrm',
  'pgnucltr',
  'pgnlcltr',
  'titlepg',
  'vertalt',
  'vertalc',
  'vertalj',
  'vertalb',
  'rtlsect',
  'ltrsect',
  'linemod',
  'linex',
  'linestarts',
  'linerestart',
  'lineppage',
  'linecont',
];
for (const word of SECTION_WORDS) {
  WORDS.set(word, (r, param, hasParam, name) => {
    sepWord(r.sep, name, param, hasParam);
  });
}
WORDS.set('sectd', (r) => {
  r.sep = defaultSep();
});

WORDS.set('ansi', (r) => r.setCodepage(1252, false));
WORDS.set('mac', (r) => r.setCodepage(10000, false));
WORDS.set('pc', (r) => r.setCodepage(437, false));
WORDS.set('pca', (r) => r.setCodepage(850, false));
WORDS.set('ansicpg', (r, param) => r.setCodepage(param, true));
WORDS.set('deff', (r, param) => r.setDefaultFont(param));
WORDS.set('deflang', (r, param) => r.setDefaultLanguage(param));

// ------------------------------------------------------- characters and breaks

WORDS.set('par', (r) => r.par());
WORDS.set('sect', (r) => r.sect());
WORDS.set('line', (r) => r.lineBreak(undefined));
WORDS.set('lbr', (r, param) => r.lineBreak(param));
for (const word of ['tab', 'page', 'column', 'chpgn', 'chftn', 'chatn', 'chftnsep', 'chftnsepc', 'ftnalt']) {
  WORDS.set(word, (r, _param, _hasParam, name) => r.special(name));
}

const CHARACTERS: { [word: string]: number } = lookup({
  emdash: 0x2014,
  endash: 0x2013,
  lquote: 0x2018,
  rquote: 0x2019,
  ldblquote: 0x201c,
  rdblquote: 0x201d,
  bullet: 0x2022,
  emspace: 0x2003,
  enspace: 0x2002,
  qmspace: 0x2005,
  zwj: 0x200d,
  zwnj: 0x200c,
  ltrmark: 0x200e,
  rtlmark: 0x200f,
  zwbo: 0x200b,
  zwnbo: 0x2060,
});
for (const word of Object.keys(CHARACTERS)) WORDS.set(word, (r) => r.char(String.fromCharCode(CHARACTERS[word])));

/** Is the control word a character: after a "\u" it counts as one of the characters that stand in for it. */
export const isCharacterWord = (word: string): boolean => CHARACTERS[word] !== undefined;

// -------------------------------------------------------------------- ignored

const nothing: WordHandler = () => undefined;

/** Control words the reader knows and has no use for. */
export const IGNORED = [
  // The file itself.
  'rtf',
  'adeflang',
  'adeff',
  'deflangfe',
  'stshfdbch',
  'stshfloch',
  'stshfhich',
  'stshfbi',
  'themelang',
  'themelangfe',
  'themelangcs',
  'vern',
  'version',
  'edmins',
  'nofpages',
  'nofwords',
  'nofchars',
  'nofcharsws',
  'id',
  'yr',
  'mo',
  'dy',
  'hr',
  'min',
  'sec',
  // Revision numbers.
  'insrsid',
  'charrsid',
  'pararsid',
  'sectrsid',
  'tblrsid',
  'delrsid',
  'rsid',
  'rsidroot',
  'styrsid',
  'crauth',
  'crdate',
  // How the document is shown and checked, and what other versions of Word should do with it.
  'viewkind',
  'viewscale',
  'viewzk',
  'viewbksp',
  'uncompressed',
  'nouicompat',
  'splytwnine',
  'htmautsp',
  'noqfpromote',
  'trackmoves',
  'trackformatting',
  'donotembedsysfont',
  'relyonvml',
  'donotembedlingdata',
  'grfdocevents',
  'validatexml',
  'showplaceholdtext',
  'ignoremixedcontent',
  'saveinvalidxml',
  'showxmlerrors',
  'nofeaturethrottle',
  'ilfomacatclnup',
  'horzdoc',
  'vertdoc',
  'jcompress',
  'jexpand',
  'lnongrid',
  'dghspace',
  'dgvspace',
  'dghorigin',
  'dgvorigin',
  'dghshow',
  'dgvshow',
  'dgmargin',
  'dgsnap',
  'formshade',
  'allprot',
  'annotprot',
  'formprot',
  'readprot',
  'enforceprot',
  'protlevel',
  'saveprevpict',
  'hyphhotz',
  'hyphconsec',
  'hyphcaps',
  'noxlattoyen',
  'expshrtn',
  'noultrlspc',
  'dntblnsbdb',
  'nospaceforul',
  'lytprtmet',
  'lyttblrtgr',
  'lytexcttp',
  'cvmme',
  'sprstsm',
  'sprsspbf',
  'sprstsp',
  'sprsbsp',
  'sprslnsp',
  'wrppunct',
  'asianbrkrule',
  'newtblstyruls',
  'nogrowautofit',
  'usenormstyforlist',
  'noindnmbrts',
  'felnbrelev',
  'nocxsptable',
  'indrlsweleven',
  'noafcnsttbl',
  'afelev',
  'utinl',
  'hwelev',
  'spltpgpar',
  'notcvasp',
  'notbrkcnstfrctbl',
  'notvatxbx',
  'krnprsnet',
  'cachedcolbal',
  'nobrkwrptbl',
  'snaptogridincell',
  'allowfieldendsel',
  'nolnhtadjtbl',
  'useltbaln',
  'alntblind',
  'rempersonalinfo',
  'bdrrlswsix',
  'oldas',
  'truncex',
  'nolead',
  'msmcap',
  'truncatefontheight',
  'subfontbysize',
  'ApplyBrkRules',
  'pgbrdrhead',
  'pgbrdrfoot',
  'pgbrdropt',
  'pgbrdrsnap',
  'pgnhn',
  'pgnhnsh',
  'pgnhnsp',
  'pgnhnsc',
  'pgnhnsm',
  'pgnhnsn',
  'pgnx',
  'pgny',
  'pgndbnum',
  'pgndbnumd',
  // Notes: where they go and how they are numbered.
  'fet',
  'ftnbj',
  'ftntj',
  'aenddoc',
  'aendnotes',
  'endnotes',
  'enddoc',
  'ftnstart',
  'aftnstart',
  'ftnrstpg',
  'ftnrestart',
  'ftnrstcont',
  'aftnrestart',
  'aftnrstcont',
  'ftnnar',
  'ftnnalc',
  'ftnnauc',
  'ftnnrlc',
  'ftnnruc',
  'ftnnchi',
  'aftnnar',
  'aftnnalc',
  'aftnnauc',
  'aftnnrlc',
  'aftnnruc',
  'aftnnchi',
  'sftnbj',
  'sftntj',
  'saftnnar',
  'saftnnalc',
  'saftnnauc',
  'saftnnrlc',
  'saftnnruc',
  'saftnnchi',
  'sftnnar',
  'sftnnalc',
  'sftnnauc',
  'sftnnrlc',
  'sftnnruc',
  'sftnnchi',
  'sftnstart',
  'saftnstart',
  'sftnrstpg',
  'sftnrestart',
  'sftnrstcont',
  'saftnrestart',
  'saftnrstcont',
  'endnhere',
  // Sections.
  'sectdefaultcl',
  'sectlinegrid',
  'sectspecifyl',
  'sectspecifycl',
  'sectspecifygen',
  'sectexpand',
  'sectunlocked',
  'stextflow',
  'binfsxn',
  'binsxn',
  'psz',
  // Paragraphs: what the model has no place for or what says the same as another word.
  'wrapdefault',
  'aspalpha',
  'aspnum',
  'faauto',
  'fahang',
  'facenter',
  'faroman',
  'favar',
  'fafixed',
  'adjustright',
  'cgrid',
  'nosnaplinegrid',
  'nooverflow',
  'nocwrap',
  'nowwrap',
  'jclisttab',
  'tscellwidthfts',
  'yts',
  'irow',
  'irowband',
  'lastrow',
  'tsrowd',
  'tblind',
  'tblindtype',
  'tsvertalt',
  'tsvertalc',
  'tsvertalb',
  'tsbrdrt',
  'tsbrdrl',
  'tsbrdrb',
  'tsbrdrr',
  'tsbrdrh',
  'tsbrdrv',
  'tsbrdrdgl',
  'tsbrdrdgr',
  'tscellpaddl',
  'tscellpaddt',
  'tscellpaddr',
  'tscellpaddb',
  'tscellpaddfl',
  'tscellpaddft',
  'tscellpaddfr',
  'tscellpaddfb',
  'tscbandsh',
  'tscbandsv',
  'tbllkhdrrows',
  'tbllkhdrcols',
  'tbllklastrow',
  'tbllklastcol',
  'tbllknorowband',
  'tbllknocolband',
  'tbllkborder',
  'tbllkshading',
  'tbllkfont',
  'tbllkcolor',
  'tbllkbestfit',
  'trftsWidthB',
  'trftsWidthA',
  'trwWidthB',
  'trwWidthA',
  'clftsWidth',
  'clwWidth',
  'clshdrawnil',
  'clcbpatraw',
  'clcfpatraw',
  'clshdngraw',
  'trshdrawnil',
  'tcelld',
  'pnrnot',
  'collapsed',
  'level',
  'subdocument',
  'sbys',
  'overlay',
  // The colour table: how a colour follows from the theme.
  'ctint',
  'cshade',
  'cmaindarkone',
  'cmainlightone',
  'cmaindarktwo',
  'cmainlighttwo',
  'caccentone',
  'caccenttwo',
  'caccentthree',
  'caccentfour',
  'caccentfive',
  'caccentsix',
  'chyperlink',
  'cfollowedhyperlink',
  'cbackgroundone',
  'ctextone',
  'cbackgroundtwo',
  'ctexttwo',
  // Styles of tables and lists named where the file states their formatting anyway.
  'ts',
  'leveltemplateid',
  'listtemplateid',
  'listid',
  // Characters.
  'additive',
  'cchs',
  'fbias',
  'webhidden',
  'nosectexpand',
  'g',
  'gcw',
  'gridtbl',
  'fittext',
  'accnone',
  'accdot',
  'acccomma',
  'acccircle',
  'accunderdot',
  'chbgcross',
  // Fields and marks.
  'fldedit',
  'flddirty',
  'fldlock',
  'fldpriv',
  'bkmkcolf',
  'bkmkcoll',
  'softpage',
  'softcol',
  'softline',
  'softlheight',
  'sectnum',
  'chdate',
  'chtime',
  'chdpl',
  'chdpa',
  // Lists outside of the list table.
  'pnlvl',
  'pnlvlblt',
  'pnlvlbody',
  'pnlvlcont',
  'pndec',
  'pnucrm',
  'pnlcrm',
  'pnucltr',
  'pnlcltr',
  'pnstart',
  'pnindent',
  'pnhang',
  'listoverridecount',
  'lsdstimax',
  'lsdlockeddef',
  'lsdsemihiddendef',
  'lsdunhideuseddef',
  'lsdqformatdef',
  'lsdprioritydef',
  // Objects.
  'objemb',
  'objlink',
  'objautlink',
  'objsub',
  'objpub',
  'objicemb',
  'objhtml',
  'objocx',
  'objw',
  'objh',
  'objscalex',
  'objscaley',
  'objcropt',
  'objcropb',
  'objcropl',
  'objcropr',
  'objsetsize',
  'objupdate',
  'rsltpict',
  'rsltbmp',
  'rsltmerge',
  'rsltrtf',
  'rslttxt',
  'objattph',
  // Shapes and drawing objects outside of their groups.
  'shpbxpage',
  'shpbxmargin',
  'shpbxcolumn',
  'shpbxignore',
  'shpbypage',
  'shpbymargin',
  'shpbypara',
  'shpbyignore',
  'shpwr',
  'shpwrk',
  'shpfblwtxt',
  'shpz',
  'shplid',
  'shpfhdr',
  'shpleft',
  'shptop',
  'shpright',
  'shpbottom',
  'shplockanchor',
  'picw',
  'pich',
  'picwgoal',
  'pichgoal',
  'picscalex',
  'picscaley',
  'piccropl',
  'piccropr',
  'piccropt',
  'piccropb',
  'bliptag',
  'blipupi',
];
for (const word of IGNORED) if (!WORDS.has(word)) WORDS.set(word, nothing);

/** Control words of the drawing objects of Word 95 ("\dpline", "\dpfillbgcr", ...): lines and fills the reader does not draw. */
export const isLegacyDrawingWord = (word: string): boolean => word.length > 2 && word.charCodeAt(0) === 0x64 && (word.charCodeAt(1) === 0x70 || word.charCodeAt(1) === 0x6f);
