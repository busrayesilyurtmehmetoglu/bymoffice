// Reads an RTF file (Rich Text Format, up to version 1.9.1) into the document
// model that the docx and doc readers fill too.
//
// The file is one stream of groups, control words and text. The reader goes
// through it once, token by token (tokenizer.ts), with a stack of states:
// a group saves the formatting in effect and restores it when it closes. A
// control word either changes the formatting (keywords.ts), stands for a
// character or a break, or says what its group is (destinations.ts): the font
// table, the style sheet, a picture, a header, a note, a field. Text that
// arrives while a story is being read becomes runs and paragraphs (story.ts),
// paragraphs in tables are gathered into table blocks (tables.ts).
//
// How the result differs from a model read from docx:
//   - There is no XML. `Story.source` is empty and no block carries a source
//     span; `Story.part` names the story ("rtf:body", "rtf:header", ...);
//     `rootName` is empty; drawings carry a stand-in element.
//   - Header and footer ids are made up: "hdr:<section>:<kind>" and
//     "ftr:<section>:<kind>", with the section counted from 0.
//   - Pictures are in `DocModel.media`, by the part name their references
//     give as target.
//   - Note ids are the number of the note counted from 1, comment ids from 0.
//   - Fields are complex fields: a "begin" mark, the code as instrText, a
//     "separate" mark, the result, an "end" mark, as from a binary file. The
//     marks have the formatting of the code and of the result next to them.
//   - Text in a font of symbols ("\fcharset2") is kept in the private use
//     area (U+F021..U+F0FF), as Word keeps it, however the file writes it.
//   - Control words the reader passes over are counted in `unknown`
//     ("rtf/<word>"). The file itself is never changed: it is saved as read.
//
// Nothing in the file is trusted. Groups are kept on a stack of the reader,
// not of the machine, and groups nested deeper than the reader follows share
// the state of the last one it does; numbers are cut to what fits; a file
// that is cut short ends where it ends, with what was read up to there; a
// file that does not start with "{\rtf" is an error. What the model costs
// in memory is bounded by the size of the file: about forty times its bytes
// at the most, and up to sixty for a file of nothing but sections that each
// bring a header of their own (see ReadRtfOptions for the limits a host can
// set).

import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  defaultSettings,
  emptyTheme,
  type Block,
  type Body,
  type Border,
  type Comment,
  type DocColor,
  type DocModel,
  type DocRelationship,
  type FontInfo,
  type FrameProps,
  type HeaderFooterKind,
  type Inline,
  type Note,
  type ParaBorders,
  type ParaProps,
  type Paragraph,
  type Run,
  type RunItem,
  type RunProps,
  type SectionProps,
  type Story as ModelStory,
  type StoryCollection,
  type TabStop,
  type Wrapper,
} from '../../model/doc/types';
import { MediaStore } from '../doc/drawings';
import { languageTag } from '../doc/sprm';
import { linkFields } from '../docx/body';
import { resolveParaProps, resolveRunProps } from '../docx/resolve';
import { charsetCodepage, decodeByte, decodePair, decodeRange, isDoubleByte, isKnownRtfCodepage, isLeadByte, SYMBOL_CODEPAGE } from './codepage';
import * as D from './destinations';
import { WORDS, isCharacterWord, isLegacyDrawingWord, type WordHandler } from './keywords';
import { MAX_LEVEL_TEXT, MAX_LISTS, buildNumbering, levelWord, newLevel, type LevelState, type ListState, type OverrideState } from './lists';
import { newPict, newShape, pictBin, pictHex, pictItem, pictWord, shapeItem, shapeWord, type PictState, type ShapeState } from './pictures';
import { defaultDop, defaultSep, pageBorderSide, sectionProps, setPageBorder, type Dop, type Sep } from './sections';
import {
  borderWith,
  cloneChp,
  clonePap,
  defaultChp,
  defaultPap,
  diffParaProps,
  diffRunProps,
  fullParaProps,
  fullRunProps,
  lookup,
  type Chp,
  type FormatContext,
  type Pap,
} from './state';
import { Story, TEXT_CODE, TEXT_DELETED, TEXT_SHOWN, type StoryKind } from './story';
import type { TableBudget } from './tables';
import { MAX_STYLES, buildStyles, newRawStyle, styleWord, type RawStyle, type StyleSheet } from './styles';
import { RtfTokenizer, hexValue, T_BIN, T_CLOSE, T_EOF, T_HEX, T_OPEN, T_SYMBOL, T_TEXT, T_WORD } from './tokenizer';

const invalid = (message: string): Error => new Error(`Invalid rtf: ${message}`);

/** Paragraphs one document may have; see the doc reader. */
const MAX_PARAGRAPHS = 2000000;
/** Groups in groups whose state is kept apart; deeper ones share the state of the last. */
const MAX_DEPTH = 1000;
/** Text boxes in text boxes that are read. */
const MAX_BOX_DEPTH = 4;
/** Fields in fields, notes in notes: stories and marks the reader keeps open at once. */
const MAX_OPEN = 64;
/** Characters a name or value gathered from the file may have. */
const MAX_COLLECT = 4096;
const MAX_FONTS = 10000;
const MAX_COLORS = 10000;
/** Pictures and text boxes one shape shows, those of the shapes grouped in it included. */
const MAX_SHAPE_PARTS = 256;
/** "\ucN": bytes after a "\u" that stand for it in old readers. */
const MAX_UC = 32;

/** Bytes a file may have: 512 MiB. */
const MAX_BYTES = 512 * 1024 * 1024;
/**
 * Tables the reader builds for a file, as a share of its size: one for every
 * 32 bytes and a thousand more. A table of a genuine file takes hundreds of
 * bytes; a file that opens tables faster than this only asks for memory
 * ("\itap32 a\nestcell\itap1 b\cell" is 31 tables in 31 bytes).
 */
const TABLE_BYTES = 32;
const TABLE_FLOOR = 1024;
const MAX_TABLES = 4000000;

export interface ReadRtfOptions {
  /** Paragraphs the document may have before it is refused; 2,000,000 when missing. A host with little memory may lower it. */
  maxParagraphs?: number;
  /**
   * Size of the file in bytes beyond which it is refused; 512 MiB when missing. The model of a file that is made
   * to be expensive takes about forty times the size of the file in memory (text with a new formatting every few
   * characters, cells and paragraphs without text; sixty for sections that each have a header), a genuine one far
   * less, so a host with little memory may lower it.
   */
  maxBytes?: number;
  /**
   * Tables (those in cells included) the reader builds; one for every 32 bytes of the file and 1,024 more when
   * missing, never more than 4,000,000. When they are used up no further table is opened: what follows is read as
   * paragraphs of the innermost table still open, or of the story when none is, and the model says so in `warnings`.
   */
  maxTables?: number;
}

const NO_RELATIONSHIPS = new Map<string, DocRelationship>();

// Items without any state are shared.
const TAB: RunItem = Object.freeze({ type: 'tab' as const });
const NO_BREAK_HYPHEN: RunItem = Object.freeze({ type: 'noBreakHyphen' as const });
const SOFT_HYPHEN: RunItem = Object.freeze({ type: 'softHyphen' as const });
const LINE_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'line' as const });
const PAGE_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'page' as const });
const COLUMN_BREAK: RunItem = Object.freeze({ type: 'break' as const, kind: 'column' as const });
const PAGE_NUMBER: RunItem = Object.freeze({ type: 'pageNumber' as const });
const SEPARATOR: RunItem = Object.freeze({ type: 'separator' as const });
const CONTINUATION: RunItem = Object.freeze({ type: 'continuationSeparator' as const });
const FOOTNOTE_REF: RunItem = Object.freeze({ type: 'footnoteRef' as const });
const ENDNOTE_REF: RunItem = Object.freeze({ type: 'endnoteRef' as const });
const ANNOTATION_REF: RunItem = Object.freeze({ type: 'annotationRef' as const });
const NOTHING: RunItem = Object.freeze({ type: 'text' as const, text: '' });

function story(part: string, blocks: Block[]): ModelStory {
  return { part, rootName: '', rootAttrs: [], blocks, relationships: NO_RELATIONSHIPS, source: '' };
}

function collection<T>(part: string): StoryCollection<T> {
  return { part, rootName: '', rootAttrs: [], relationships: NO_RELATIONSHIPS, source: '', items: new Map() };
}

// What is done when a group closes.
const C_NONE = 0;
const C_STYLE = 1;
const C_STYLESHEET = 2;
const C_LIST = 3;
const C_LEVEL = 4;
const C_OVERRIDE = 5;
const C_LFOLEVEL = 6;
const C_PICT = 7;
const C_SHAPE = 8;
const C_SHAPE_NAME = 9;
const C_SHAPE_VALUE = 10;
const C_BOX_TEXT = 11;
const C_FIELD = 12;
const C_NOTE = 13;
const C_COMMENT = 14;
const C_HEADER = 15;
const C_SEPARATOR = 16;
const C_BOOKMARK_START = 17;
const C_BOOKMARK_END = 18;
const C_COMMENT_INITIALS = 19;
const C_COMMENT_AUTHOR = 20;
const C_RANGE_START = 21;
const C_RANGE_END = 22;
const C_COMMENT_REF = 23;
const C_LIST_TEXT = 24;
const C_ROW_PROPS = 25;
const C_FONT_ALT = 26;
const C_PANOSE = 27;
const C_DISCARD = 28;

/** What a group saves and restores. */
interface State {
  dest: number;
  chp: Chp;
  /** `chp` belongs to this group alone and may be changed in place. */
  chpOwn: boolean;
  pap: Pap;
  papOwn: boolean;
  uc: number;
  /** "\fcs1": the control words that follow describe text of complex scripts. */
  complex: boolean;
  /** Which part of the code page the text that follows is from: 0 not said, 1 low, 2 high, 3 double-byte. */
  assoc: number;
  /** The kind of text ("\loch" 1, "\hich" 2, "\dbch" 3) the next font word is about; 0 once a font word was read, and in a new group. */
  slot: number;
  rtl: boolean;
  /** Text is the code of a field. */
  code: boolean;
  /** "\*" was read and the next control word says what the group is. */
  star: boolean;
  closer: number;
  payload: unknown;
}

interface FontEntry {
  name: string;
  altName: string | undefined;
  charset: number | undefined;
  codepage: number | undefined;
  family: string | undefined;
  pitch: number | undefined;
  panose: string | undefined;
}

const FAMILIES: { [word: string]: string } = lookup({
  fnil: 'auto',
  froman: 'roman',
  fswiss: 'swiss',
  fmodern: 'modern',
  fscript: 'script',
  fdecor: 'decorative',
  ftech: 'auto',
  fbidi: 'auto',
});
const PITCHES = ['default', 'fixed', 'variable'];
const HEX = '0123456789ABCDEF';
const hex2 = (value: number): string => HEX.charAt((value >> 4) & 15) + HEX.charAt(value & 15);

const HEADERS: { [word: string]: { list: 'headers' | 'footers'; kind: HeaderFooterKind } } = lookup({
  header: { list: 'headers', kind: 'default' },
  headerr: { list: 'headers', kind: 'default' },
  headerl: { list: 'headers', kind: 'even' },
  headerf: { list: 'headers', kind: 'first' },
  footer: { list: 'footers', kind: 'default' },
  footerr: { list: 'footers', kind: 'default' },
  footerl: { list: 'footers', kind: 'even' },
  footerf: { list: 'footers', kind: 'first' },
});

/** The separators of the notes are notes of a kind of their own, as in a docx file. */
const SEPARATORS: { [word: string]: { list: 'footnotes' | 'endnotes'; id: number; kind: string } } = lookup({
  ftnsep: { list: 'footnotes', id: -1, kind: 'separator' },
  ftnsepc: { list: 'footnotes', id: 0, kind: 'continuationSeparator' },
  ftncn: { list: 'footnotes', id: -2, kind: 'continuationNotice' },
  aftnsep: { list: 'endnotes', id: -1, kind: 'separator' },
  aftnsepc: { list: 'endnotes', id: 0, kind: 'continuationSeparator' },
  aftncn: { list: 'endnotes', id: -2, kind: 'continuationNotice' },
});

const SKIP = new Set<string>(D.SKIPPED);

/** Every control word destination() looks at. */
const DESTINATIONS = new Set<string>(
  D.SKIPPED.concat(Object.keys(HEADERS), Object.keys(SEPARATORS), [
    'fonttbl',
    'colortbl',
    'stylesheet',
    'listtable',
    'list',
    'listlevel',
    'leveltext',
    'listoverridetable',
    'listoverride',
    'lfolevel',
    'pict',
    'shppict',
    'mmathPict',
    'shp',
    'shpgrp',
    'do',
    'shpinst',
    'sp',
    'sn',
    'sv',
    'shptxt',
    'dptxbxtext',
    'field',
    'fldinst',
    'fldrslt',
    'footnote',
    'annotation',
    'atnid',
    'atnauthor',
    'atrfstart',
    'atrfend',
    'atnref',
    'bkmkstart',
    'bkmkend',
    'falt',
    'panose',
    'listtext',
    'pntext',
    'nesttableprops',
    'object',
    'result',
    'upr',
    'ud',
  ]),
);

const TAB_LEADERS: { [word: string]: string } = lookup({ tldot: 'dot', tlhyph: 'hyphen', tlul: 'underscore', tlth: 'heavy', tleq: 'underscore', tlmdot: 'middleDot' });
const TAB_ALIGNS: { [word: string]: string } = lookup({ tqr: 'right', tqc: 'center', tqdec: 'decimal' });
const PARA_BORDERS: { [word: string]: keyof ParaBorders | 'box' } = lookup({
  brdrt: 'top',
  brdrb: 'bottom',
  brdrl: 'left',
  brdrr: 'right',
  brdrbtw: 'between',
  brdrbar: 'bar',
  box: 'box',
});
const BREAK_CLEAR = ['none', 'left', 'right', 'all'];

// Which border the border words describe.
const B_NONE = 0;
const B_PARA = 1;
const B_TABLE = 2;
const B_CHAR = 3;
const B_PAGE = 4;

interface Place {
  run: Run;
  index: number;
}

/** A field mark and where it is, for as long as its paragraph is being read. */
interface MarkRef {
  run: Run;
  story: Story;
  /** Paragraphs the story had ended when the mark was made, and the index of its run in the paragraph. */
  serial: number;
  index: number;
}

/** A field whose group is open. */
interface OpenField {
  separated: boolean;
  begin: MarkRef | undefined;
  separate: MarkRef | undefined;
  /** The formatting in full of the first text of the code and of the result. */
  code: RunProps | undefined;
  result: RunProps | undefined;
}

export class RtfReader {
  readonly warnings: string[] = [];
  readonly unknown = new Map<string, number>();
  readonly dop: Dop = defaultDop();
  sep: Sep = defaultSep();
  /** The state of the group being read. */
  st: State;
  story: Story;

  private readonly tk: RtfTokenizer;
  private readonly stack: State[] = [];
  private depth = 0;
  /** Groups open beyond the depth the reader follows. */
  private overflow = 0;
  /** Bytes after a "\u" that are still to be passed over. */
  private skip = 0;
  /** The first byte of a two-byte character whose second byte is still to come; -1 for none. */
  private lead = -1;
  /** In the name of a style: the byte just read as "\'hh", which Word writes once more as two plain digits; -1 for none. */
  private echo = -1;
  private readonly maxParagraphs: number;
  private paragraphs = 0;

  private codepage = 1252;
  private codepageStated = false;
  private defaultFont = 0;
  private defaultLanguage = 0;
  private readonly fonts = new Map<number, FontEntry>();
  private font: (FontEntry & { index: number }) | undefined;
  private fontDepth = -1;
  private readonly colors: (DocColor | undefined)[] = [];
  private red = 0;
  private green = 0;
  private blue = 0;
  private colorStated = false;

  private readonly rawStyles: RawStyle[] = [];
  /** The style whose group is being read. */
  private rawStyle: RawStyle | undefined;
  private sheet: StyleSheet | undefined;
  private readonly lists: ListState[] = [];
  private readonly overrides: OverrideState[] = [];
  private list: ListState | undefined;
  private level: LevelState | undefined;
  private override: OverrideState | undefined;
  private overrideLevel: OverrideState['levels'][number] | undefined;
  /** "\ls" numbers that name a list of the file. */
  private readonly knownLists = new Set<number>();

  private readonly buffers: string[] = [];
  private readonly store = new MediaStore();
  private pict: PictState | undefined;
  private readonly shapes: ShapeState[] = [];
  private boxDepth = 0;
  private openStories = 0;
  private readonly fields: OpenField[] = [];
  private readonly tableBudget: TableBudget;
  private lastSection: { key: string; props: SectionProps } | undefined;
  private lastSectionPara: { from: ParaProps; section: SectionProps; props: ParaProps } | undefined;

  private readonly model: DocModel;
  private readonly body: Story;
  private sectionIndex = 0;
  private sectionRefs: { headers: Partial<Record<HeaderFooterKind, string>>; footers: Partial<Record<HeaderFooterKind, string>> } = { headers: {}, footers: {} };
  private footnoteId = 0;
  private endnoteId = 0;
  private commentId = 0;
  private commentAuthor: string | undefined;
  private commentInitials: string | undefined;
  private readonly separators: { list: 'footnotes' | 'endnotes'; note: Note }[] = [];
  private readonly bookmarks = new Map<string, string>();
  private nextBookmark = 0;
  /** Marks of comment ranges by the text that ties them to their comment. */
  private readonly ranges = new Map<string, { id: string }[]>();

  private borderTarget = B_NONE;
  private borderSide: keyof ParaBorders | 'box' = 'top';
  private pageSide: 'top' | 'left' | 'bottom' | 'right' = 'top';
  private tabAlign = 'left';
  private tabLeader: string | undefined;

  private readonly plainChp = defaultChp();
  private readonly plainPap = defaultPap();
  private readonly format: FormatContext;
  private readonly formatBare: FormatContext;
  private readonly runBases = new Map<string, RunProps>();
  private readonly paraBases = new Map<string, ParaProps>();
  private readonly runPool = new Map<string, RunProps>();
  private readonly paraPool = new Map<string, ParaProps>();
  private readonly paraFinal = new WeakMap<ParaProps, Map<RunProps, ParaProps>>();
  private readonly deletion: Wrapper = { kind: 'del' };
  private readonly insertion: Wrapper = { kind: 'ins' };

  constructor(bytes: Uint8Array, options: ReadRtfOptions) {
    const limit = options.maxParagraphs;
    this.maxParagraphs = limit !== undefined && limit >= 1 && limit < MAX_PARAGRAPHS ? Math.floor(limit) : MAX_PARAGRAPHS;
    const maxBytes = options.maxBytes !== undefined && options.maxBytes >= 0 && options.maxBytes < MAX_BYTES ? Math.floor(options.maxBytes) : MAX_BYTES;
    if (bytes.length > maxBytes) throw invalid(`the file is larger than ${maxBytes} bytes`);
    const tables = options.maxTables !== undefined && options.maxTables >= 0 ? Math.floor(options.maxTables) : Math.floor(bytes.length / TABLE_BYTES) + TABLE_FLOOR;
    this.tableBudget = { left: Math.min(tables, MAX_TABLES), exceeded: false };

    // The file starts with "{\rtf"; a byte order mark and white space in front of it are let pass.
    let start = 0;
    if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) start = 3;
    while (start < bytes.length && (bytes[start] === 0x20 || bytes[start] === 0x09 || bytes[start] === 0x0a || bytes[start] === 0x0d)) start++;
    const head = [0x7b, 0x5c, 0x72, 0x74, 0x66];
    for (let i = 0; i < head.length; i++) if (bytes[start + i] !== head[i]) throw invalid('the file does not start with "{\\rtf"');
    this.tk = new RtfTokenizer(bytes, start);

    const fontName = (index: number): string | undefined => {
      const font = this.fonts.get(index < 0 ? this.defaultFont : index) || this.fonts.get(this.defaultFont);
      return font ? font.name : undefined;
    };
    const color = (index: number): DocColor | undefined => (index > 0 && index < this.colors.length ? this.colors[index] : undefined);
    const self = this;
    this.formatBare = {
      fontName,
      color,
      paraStyleId: () => undefined,
      charStyleId: () => undefined,
      get widowControl() {
        return self.dop.widowControl;
      },
    };
    this.format = {
      fontName,
      color,
      paraStyleId: (number) => {
        const sheet = this.styles();
        const id = sheet.paraIds.get(number);
        return id === sheet.styles.defaultIds.paragraph ? undefined : id;
      },
      charStyleId: (number) => {
        const sheet = this.styles();
        const id = sheet.charIds.get(number);
        return id === sheet.styles.defaultIds.character ? undefined : id;
      },
      get widowControl() {
        return self.dop.widowControl;
      },
    };

    const settings = defaultSettings();
    // What these files were made with: layout treats tables and spacing as Word 2003 does.
    settings.compatSettings.set('compatibilityMode', '11');
    this.model = {
      body: story('rtf:body', []) as Body,
      styles: { defaults: { para: EMPTY_PARA_PROPS, run: EMPTY_RUN_PROPS }, styles: new Map(), defaultIds: {} },
      numbering: { abstracts: new Map(), nums: new Map() },
      theme: emptyTheme(),
      settings,
      fonts: new Map(),
      headers: new Map(),
      footers: new Map(),
      strict: false,
      warnings: this.warnings,
      unknown: this.unknown,
    };
    this.body = this.story = new Story('body', this.format, this.tableBudget);
    this.st = this.newState();
  }

  private newState(): State {
    return {
      dest: D.D_NORMAL,
      chp: this.plainChp,
      chpOwn: false,
      pap: this.plainPap,
      papOwn: false,
      uc: 1,
      complex: false,
      assoc: 0,
      slot: 0,
      rtl: false,
      code: false,
      star: false,
      closer: C_NONE,
      payload: undefined,
    };
  }

  // -------------------------------------------------------------------- state

  /** The character formatting of the group, to change. */
  chpW(): Chp {
    const st = this.st;
    if (!st.chpOwn) {
      st.chp = cloneChp(st.chp);
      st.chpOwn = true;
    } else st.chp.cache = undefined;
    return st.chp;
  }

  /** The paragraph formatting of the group, to change. */
  papW(): Pap {
    const st = this.st;
    if (!st.papOwn) {
      st.pap = clonePap(st.pap);
      st.papOwn = true;
    } else st.pap.cache = undefined;
    return st.pap;
  }

  /** "\plain". */
  plain(): void {
    const st = this.st;
    st.chp = this.plainChp;
    st.chpOwn = false;
    st.complex = false;
    st.assoc = 0;
    st.slot = 0;
    st.rtl = false;
  }

  /** "\pard". */
  pard(): void {
    const st = this.st;
    st.pap = this.plainPap;
    st.papOwn = false;
    this.borderTarget = B_NONE;
    this.tabAlign = 'left';
    this.tabLeader = undefined;
  }

  count(key: string): void {
    this.unknown.set(key, (this.unknown.get(key) || 0) + 1);
  }

  setCodepage(codepage: number, stated: boolean): void {
    if (this.codepageStated && !stated) return;
    if (!isKnownRtfCodepage(codepage)) {
      this.warnings.push(`Text in code page ${codepage} is read as Windows-1252`);
      codepage = 1252;
    }
    this.codepage = codepage;
    if (stated) this.codepageStated = true;
  }

  setDefaultFont(index: number): void {
    this.defaultFont = index;
  }

  setDefaultLanguage(id: number): void {
    this.defaultLanguage = id;
  }

  /** The code page the bytes of the text being read are in. */
  private textCodepage(): number {
    const st = this.st;
    if (st.dest === D.D_FONTTBL && this.font && this.font.codepage !== undefined) return this.font.codepage;
    const chp = st.chp;
    let index = chp.font;
    if (st.assoc === 2 && chp.fontHi >= 0) index = chp.fontHi;
    else if (st.assoc === 3 && chp.fontDb >= 0) index = chp.fontDb;
    else if (st.assoc === 1 && chp.fontLo >= 0) index = chp.fontLo;
    else if (st.assoc === 0 && st.rtl && chp.fontCs >= 0) index = chp.fontCs;
    const font = this.fonts.get(index < 0 ? this.defaultFont : index);
    return font && font.codepage !== undefined ? font.codepage : this.codepage;
  }

  // ------------------------------------------------------------------- styles

  private styles(): StyleSheet {
    let sheet = this.sheet;
    if (!sheet) {
      sheet = this.sheet = buildStyles(this.rawStyles, this.formatBare, languageTag(this.defaultLanguage));
      this.model.styles = sheet.styles;
      // What was made for the styles there were before is no longer right.
      this.runBases.clear();
      this.paraBases.clear();
    }
    return sheet;
  }

  private paraBase(style: string | undefined): ParaProps {
    const key = style === undefined ? '' : style;
    let base = this.paraBases.get(key);
    if (!base) {
      this.styles();
      base = resolveParaProps(this.model, style === undefined ? {} : { style });
      this.paraBases.set(key, base);
    }
    return base;
  }

  private runBase(paraStyle: string, charStyle: string | undefined): RunProps {
    const key = paraStyle + '\u0000' + (charStyle === undefined ? '' : charStyle);
    let base = this.runBases.get(key);
    if (!base) {
      this.styles();
      base = resolveRunProps(this.model, paraStyle ? { style: paraStyle } : {}, charStyle === undefined ? {} : { style: charStyle });
      this.runBases.set(key, base);
    }
    return base;
  }

  /** One object for all properties that say the same. */
  private pooled<T extends object>(pool: Map<string, T>, props: T, empty: T): T {
    const key = JSON.stringify(props);
    if (key === '{}') return empty;
    const known = pool.get(key);
    if (known) return known;
    pool.set(key, props);
    return props;
  }

  /** The run properties in full for the character formatting in effect. */
  private fullRun(chp: Chp, rtl: boolean): RunProps {
    const cache = chp.cache || (chp.cache = new Map());
    const key = rtl ? '\u0000r' : '\u0000l';
    let full = cache.get(key);
    if (!full) cache.set(key, (full = fullRunProps(this.format, chp, rtl)));
    return full;
  }

  /** The direct run properties for text in a paragraph of a style. */
  private runPropsFor(chp: Chp, rtl: boolean, paraStyle: string): RunProps {
    const cache = chp.cache || (chp.cache = new Map());
    const key = (rtl ? 'r' : 'l') + paraStyle;
    let props = cache.get(key);
    if (!props) {
      const full = this.fullRun(chp, rtl);
      props = this.pooled(this.runPool, diffRunProps(full, this.runBase(paraStyle, full.style)), EMPTY_RUN_PROPS);
      cache.set(key, props);
    }
    return props;
  }

  /** The direct paragraph properties for the paragraph formatting in effect. */
  private paraProps(pap: Pap): ParaProps {
    let props = pap.cache;
    if (!props) {
      const full = fullParaProps(this.format, pap);
      props = pap.cache = this.pooled(this.paraPool, diffParaProps(full, this.paraBase(full.style)), EMPTY_PARA_PROPS);
    }
    return props;
  }

  /** Id of the style the paragraph being read has, "" for the default one. */
  private paraStyle(): string {
    const story = this.story;
    if (story.styleUsed === undefined) {
      const id = this.format.paraStyleId(this.st.pap.style);
      story.styleUsed = id === undefined ? '' : id;
    }
    return story.styleUsed;
  }

  // --------------------------------------------------------------------- text

  private wrapper(chp: Chp): Wrapper | undefined {
    return chp.deleted ? this.deletion : chp.inserted ? this.insertion : undefined;
  }

  /** Adds text to the story being read. */
  private addText(text: string): void {
    if (!text) return;
    const st = this.st;
    const chp = st.chp;
    const style = this.paraStyle();
    const kind = st.code ? TEXT_CODE : chp.deleted ? TEXT_DELETED : TEXT_SHOWN;
    const full = this.fullRun(chp, st.rtl);
    if (this.fields.length) this.fieldText(full);
    this.story.addText(text, kind, this.runPropsFor(chp, st.rtl, style), full, this.wrapper(chp));
  }

  /**
   * Text of the innermost open field arrives in the formatting `full`. The marks of a field are made before its
   * text, in the formatting around the field; Word gives them that of the text they stand next to: the mark that
   * begins the field that of its code, the one that separates code and result that of the result.
   */
  private fieldText(full: RunProps): void {
    const field = this.fields[this.fields.length - 1];
    // Text of another story inside the field (a note in its result) is not text of the field.
    if (!field.begin || field.begin.story !== this.story) return;
    if (this.st.code) {
      if (field.code || field.separated) return;
      field.code = full;
      if (field.begin) this.restyle(field.begin, full);
    } else if (field.separated && !field.result) {
      field.result = full;
      if (field.separate) this.restyle(field.separate, full);
    }
  }

  /** Gives a field mark another formatting, when the paragraph it is in is still being read. */
  private restyle(mark: MarkRef, full: RunProps): void {
    const story = mark.story;
    if (story !== this.story || story.paragraphs !== mark.serial || story.runs[mark.index] !== mark.run || story.styleUsed === undefined) return;
    story.fulls[mark.index] = full;
    mark.run.props = this.pooled(this.runPool, diffRunProps(full, this.runBase(story.styleUsed, full.style)), EMPTY_RUN_PROPS);
  }

  /** A field mark in a run of its own, in the formatting in effect. */
  private fieldMark(charType: 'begin' | 'separate' | 'end', chp: Chp, rtl: boolean, full: RunProps | undefined): MarkRef {
    const story = this.story;
    const style = this.paraStyle();
    const used = full || this.fullRun(chp, rtl);
    const props = full ? this.pooled(this.runPool, diffRunProps(full, this.runBase(style, full.style)), EMPTY_RUN_PROPS) : this.runPropsFor(chp, rtl, style);
    const run = story.addAlone({ type: 'fieldChar', charType }, props, used);
    return { run, story, serial: story.paragraphs, index: story.runs.length - 1 };
  }

  /** Adds an item to the story being read, if the group being read is text of it. */
  addItem(item: RunItem): Place | undefined {
    const st = this.st;
    if (st.dest !== D.D_NORMAL) return undefined;
    const chp = st.chp;
    const full = this.fullRun(chp, st.rtl);
    if (this.fields.length) this.fieldText(full);
    return this.story.addItem(item, this.runPropsFor(chp, st.rtl, this.paraStyle()), full, this.wrapper(chp));
  }

  /** Text as decoded: where it goes depends on what the group is. */
  private route(text: string): void {
    const st = this.st;
    switch (st.dest) {
      case D.D_NORMAL: {
        // A tab character in the text is a tab, as "\tab" is.
        let from = 0;
        for (let at = text.indexOf('\t'); at >= 0; at = text.indexOf('\t', from)) {
          this.addText(text.slice(from, at));
          this.addItem(TAB);
          from = at + 1;
        }
        this.addText(from ? text.slice(from) : text);
        return;
      }
      case D.D_FONTTBL: {
        const font = this.font;
        for (let i = 0; i < text.length; i++) {
          if (text.charCodeAt(i) === 0x3b) this.finishFont();
          else if (font && this.font === font && font.name.length < 256) font.name += text.charAt(i);
        }
        return;
      }
      case D.D_COLORTBL:
        for (let i = 0; i < text.length; i++) {
          if (text.charCodeAt(i) !== 0x3b) continue;
          if (this.colors.length < MAX_COLORS) this.colors.push(this.colorStated ? { rgb: hex2(this.red) + hex2(this.green) + hex2(this.blue) } : undefined);
          this.red = this.green = this.blue = 0;
          this.colorStated = false;
        }
        return;
      case D.D_LEVELTEXT: {
        const level = this.level;
        if (level) for (let i = 0; i < text.length && level.text.length < MAX_LEVEL_TEXT + 2; i++) level.text.push(text.charCodeAt(i));
        return;
      }
    }
    if (D.isCollector(st.dest) && this.buffers.length) {
      const top = this.buffers.length - 1;
      if (this.buffers[top].length < MAX_COLLECT) this.buffers[top] += text;
    }
  }

  /** bytes[start..stop) of the file are text. */
  private text(start: number, stop: number): void {
    const st = this.st;
    st.star = false;
    const bytes = this.tk.bytes;
    if (st.dest === D.D_PICT) {
      if (this.pict) pictHex(this.pict, bytes, start, stop);
      return;
    }
    // Line ends are not text. What is passed over after a "\u" is.
    let p = start;
    while (this.skip > 0 && p < stop) {
      const c = bytes[p++];
      if (c !== 0x0a && c !== 0x0d) this.skip--;
    }
    if (this.lead >= 0) {
      // The second byte of a two-byte character, written as it is.
      while (p < stop && (bytes[p] === 0x0a || bytes[p] === 0x0d)) p++;
      if (p < stop) {
        const lead = this.lead;
        this.lead = -1;
        this.route(decodePair(this.textCodepage(), lead, bytes[p++]));
      }
    }
    if (this.echo >= 0) {
      const echo = this.echo;
      this.echo = -1;
      if (p + 1 < stop && hexValue(bytes[p]) === echo >> 4 && hexValue(bytes[p + 1]) === (echo & 15)) p += 2;
    }
    if (p >= stop) return;
    if (st.dest !== D.D_NORMAL && st.dest !== D.D_FONTTBL && st.dest !== D.D_COLORTBL && st.dest !== D.D_LEVELTEXT && !D.isCollector(st.dest)) return;

    let out = '';
    let from = p;
    let plain = true;
    // In a font of symbols every byte is a glyph of the font, the letters and digits too.
    const symbols = this.symbolText();
    const flush = (to: number) => {
      if (to <= from) return;
      if (plain && !symbols) {
        // Characters of seven bits: the bytes are the text.
        for (let i = from; i < to; i += 8192) out += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(to, i + 8192)) as unknown as number[]);
      } else out += decodeRange(bytes, from, to, this.textCodepage());
      plain = true;
    };
    for (; p < stop; p++) {
      const c = bytes[p];
      if (c >= 0x80) plain = false;
      else if (c < 0x20 && c !== 0x09) {
        // Line ends and other control characters are not text; a tab is.
        flush(p);
        from = p + 1;
      }
    }
    flush(stop);
    this.route(out);
  }

  /**
   * The text being read is shown text in a font of symbols ("\fcharset2"): its characters are glyphs of the font,
   * which the model keeps in the private use area (U+F021..U+F0FF), however the file writes them: as bytes, as
   * "\'hh" or as "\uN".
   */
  private symbolText(): boolean {
    const st = this.st;
    return st.dest === D.D_NORMAL && !st.code && this.textCodepage() === SYMBOL_CODEPAGE;
  }

  /** "\'hh". */
  private hex(byte: number): void {
    const st = this.st;
    st.star = false;
    if (this.skip > 0) {
      this.skip--;
      return;
    }
    if (st.dest === D.D_PICT) return;
    if (st.dest === D.D_LEVELTEXT) {
      // Read in the code page of the font of the level, which the level states after its text.
      if (this.level && this.level.text.length < MAX_LEVEL_TEXT + 2) this.level.text.push(byte - 256);
      return;
    }
    const codepage = this.textCodepage();
    if (this.lead >= 0) {
      const lead = this.lead;
      this.lead = -1;
      this.route(decodePair(codepage, lead, byte));
      return;
    }
    if (isDoubleByte(codepage) && isLeadByte(codepage, byte)) {
      this.lead = byte;
      return;
    }
    // A line end or another control character given as a byte is not text; a tab is.
    if (byte < 0x20 && byte !== 0x09) return;
    this.route(decodeByte(codepage, byte));
    if (st.dest === D.D_STYLE && byte >= 0x80) this.echo = byte;
  }

  /** "\uN": a character by its number, followed by what old readers show instead. */
  unicode(value: number): void {
    const st = this.st;
    this.lead = -1;
    this.skip = st.uc;
    let code = value < 0 ? value + 65536 : value;
    if (code < 0 || code > 0xffff) code = 0xfffd;
    if (st.dest === D.D_LEVELTEXT) {
      if (this.level && this.level.text.length < MAX_LEVEL_TEXT + 2) this.level.text.push(code);
      return;
    }
    if (code < 0x20 && code !== 0x09) return;
    // Word writes a glyph of a font of symbols as "\u-3913" (U+F0B7); other writers give the code in the font.
    if (code > 0x20 && code < 0x100 && this.symbolText()) code += 0xf000;
    this.route(String.fromCharCode(code));
  }

  setUc(value: number): void {
    this.st.uc = value < 0 ? 0 : value > MAX_UC ? MAX_UC : value;
  }

  /** A character a control word or control symbol stands for. One that stands in for a "\u" is passed over: word() and symbol() see to that. */
  char(text: string): void {
    this.route(text);
  }

  // --------------------------------------------------------------- paragraphs

  private finalPara(props: ParaProps, mark: RunProps, section: SectionProps | undefined): ParaProps {
    let out = props;
    if (mark !== EMPTY_RUN_PROPS) {
      let byMark = this.paraFinal.get(props);
      if (!byMark) this.paraFinal.set(props, (byMark = new Map()));
      let withMark = byMark.get(mark);
      if (!withMark) byMark.set(mark, (withMark = { ...props, markProps: mark }));
      out = withMark;
    }
    if (!section) return out;
    // Paragraphs that end sections described alike share their properties.
    const last = this.lastSectionPara;
    if (last && last.from === out && last.section === section) return last.props;
    const ended = { ...out, section };
    this.lastSectionPara = { from: out, section, props: ended };
    return ended;
  }

  /** Ends the paragraph being read with the formatting in effect. */
  private endParagraph(section?: SectionProps): Paragraph {
    if (++this.paragraphs > this.maxParagraphs) throw invalid(`the document has more than ${this.maxParagraphs} paragraphs`);
    const st = this.st;
    const current = this.story;
    const pap = st.pap;
    const props = this.paraProps(pap);
    const style = props.style === undefined ? '' : props.style;
    if (current.hasRuns && current.styleUsed !== undefined && current.styleUsed !== style) {
      // The runs were made for another style than the paragraph ends up with: they are made again.
      for (let i = 0; i < current.runs.length; i++) {
        const full = current.fulls[i];
        current.runs[i].props = this.pooled(this.runPool, diffRunProps(full, this.runBase(style, full.style)), EMPTY_RUN_PROPS);
      }
    }
    const mark = this.runPropsFor(st.chp, st.rtl, style);
    this.dropPending(current);
    // The number a writer puts in front of the text shows only where the file has no list for the paragraph.
    const numbered = pap.list > 0 && this.knownLists.has(pap.list);
    return current.takeParagraph(this.finalPara(props, mark, section), !numbered);
  }

  /** Marks that waited for a note or comment that did not come are taken out. */
  private dropPending(current: Story): void {
    if (current.pendingNote) {
      current.pendingNote.run.content[current.pendingNote.index] = NOTHING;
      current.pendingNote = undefined;
    }
    if (current.pendingComment) {
      current.pendingComment.run.content[current.pendingComment.index] = NOTHING;
      current.pendingComment = undefined;
    }
  }

  private paraDepth(): number {
    const pap = this.st.pap;
    return pap.depth > 0 ? pap.depth : pap.inTable ? 1 : 0;
  }

  /** "\par". */
  par(): void {
    if (this.st.dest !== D.D_NORMAL) return;
    const depth = this.paraDepth();
    this.story.assembler.add(this.endParagraph(), depth);
  }

  /** "\sect": ends the paragraph and, in the body, the section. */
  sect(): void {
    if (this.st.dest !== D.D_NORMAL) return;
    if (this.story !== this.body) {
      this.par();
      return;
    }
    const depth = this.paraDepth();
    this.story.assembler.add(this.endParagraph(this.section()), depth);
    this.sectionIndex++;
    this.sectionRefs = { headers: {}, footers: {} };
  }

  private section(): SectionProps {
    const props = sectionProps(this.sep, this.dop, this.sectionIndex === 0);
    if (Object.keys(this.sectionRefs.headers).length) props.headers = this.sectionRefs.headers;
    if (Object.keys(this.sectionRefs.footers).length) props.footers = this.sectionRefs.footers;
    // Sections described alike share one description.
    const key = JSON.stringify(props);
    const last = this.lastSection;
    if (last) {
      if (last.key === key) return last.props;
      // What two sections that follow each other say alike (the paper, the margins) they share.
      const mine = props as unknown as { [name: string]: unknown };
      const theirs = last.props as unknown as { [name: string]: unknown };
      for (const name of Object.keys(mine)) {
        const value = mine[name];
        const other = theirs[name];
        if (value && other && typeof value === 'object' && value !== other && JSON.stringify(value) === JSON.stringify(other)) mine[name] = other;
      }
    }
    this.lastSection = { key, props };
    return props;
  }

  /** "\cell" and "\nestcell". */
  cell(nested: boolean): void {
    if (this.st.dest !== D.D_NORMAL) return;
    const assembler = this.story.assembler;
    const depth = nested ? Math.max(2, this.paraDepth()) : 1;
    assembler.add(this.endParagraph(), depth);
    assembler.endCell(depth);
  }

  /** "\row" and "\nestrow". */
  row(nested: boolean): void {
    if (this.st.dest !== D.D_NORMAL) return;
    const current = this.story;
    const assembler = current.assembler;
    const depth = nested ? Math.max(2, assembler.depth) : 1;
    // Text without a cell mark after it is a cell all the same.
    if (current.children.length) assembler.add(this.endParagraph(), depth);
    // Word writes the description of a row of a table in a cell after its cells; a writer that puts it in front of
    // them has the row mark there too, with no cell before it.
    assembler.endRow(depth, current.row.def(), nested && this.st.closer === C_ROW_PROPS);
  }

  // ------------------------------------------------------------------ borders

  /** A control word that says which border the border words after it describe; false when it is not one. */
  borderStart(word: string): boolean {
    const side = PARA_BORDERS[word];
    if (side !== undefined) {
      this.borderTarget = B_PARA;
      this.borderSide = side;
      this.story.row.release();
      return true;
    }
    if (word === 'chbrdr') {
      this.borderTarget = B_CHAR;
      this.story.row.release();
      return true;
    }
    const pageSide = pageBorderSide(word);
    if (pageSide !== undefined) {
      this.borderTarget = B_PAGE;
      this.pageSide = pageSide;
      this.story.row.release();
      return true;
    }
    return false;
  }

  /** A control word of a table row: borders that follow belong to the row. */
  tableWord(word: string, param: number, hasParam: boolean): void {
    if (word === 'trowd') {
      this.story.row.reset();
      this.borderTarget = B_NONE;
      return;
    }
    if (this.story.row.word(word, param, hasParam)) this.borderTarget = B_TABLE;
  }

  /** A control word that describes a border ("\brdrs", "\brdrw15", ...). */
  border(word: string, param: number): void {
    switch (this.borderTarget) {
      case B_TABLE:
        this.story.row.border(word, param);
        return;
      case B_PARA: {
        if (word === 'brdrtbl') return;
        const pap = this.papW();
        const borders: ParaBorders = { ...pap.borders };
        const sides: (keyof ParaBorders)[] = this.borderSide === 'box' ? ['top', 'left', 'bottom', 'right'] : [this.borderSide];
        for (const side of sides) {
          const next = borderWith(this.format, borders[side], word, param);
          if (next) borders[side] = next;
        }
        pap.borders = borders;
        return;
      }
      case B_CHAR: {
        const next = borderWith(this.format, this.st.chp.border, word, param);
        if (next) this.chpW().border = next;
        return;
      }
      case B_PAGE: {
        const current: Border | undefined = this.sep.borders ? this.sep.borders[this.pageSide] : undefined;
        setPageBorder(this.sep, this.pageSide, borderWith(this.format, current, word, param));
        return;
      }
    }
  }

  // --------------------------------------------------------------------- tabs

  /** A control word of a tab stop; false when it is not one. */
  tabWord(word: string, param: number): boolean {
    const align = TAB_ALIGNS[word];
    if (align !== undefined) {
      this.tabAlign = align;
      return true;
    }
    const leader = TAB_LEADERS[word];
    if (leader !== undefined) {
      this.tabLeader = leader;
      return true;
    }
    if (word !== 'tx' && word !== 'tb') return false;
    const tab: TabStop = { pos: Math.max(-31680, Math.min(param, 31680)), align: word === 'tb' ? 'bar' : this.tabAlign };
    if (this.tabLeader !== undefined) tab.leader = this.tabLeader;
    this.tabAlign = 'left';
    this.tabLeader = undefined;
    const pap = this.papW();
    // A paragraph has no more tab stops than Word allows; one at a position replaces the one that was there.
    const tabs = (pap.tabs || []).filter((other) => other.pos !== tab.pos);
    if (tabs.length >= 64) return true;
    tabs.push(tab);
    tabs.sort((a, b) => a.pos - b.pos);
    pap.tabs = tabs;
    return true;
  }

  /** The frame of the paragraph, to change. */
  frameW(): FrameProps {
    const pap = this.papW();
    return (pap.frame = { ...pap.frame });
  }

  // ------------------------------------------------------------- special marks

  lineBreak(clear: number | undefined): void {
    if (clear === undefined) this.addItem(LINE_BREAK);
    else this.addItem({ type: 'break', kind: 'line', clear: BREAK_CLEAR[clear] || 'none' });
  }

  special(word: string): void {
    switch (word) {
      case 'tab':
        this.addItem(TAB);
        return;
      case 'page':
        this.addItem(PAGE_BREAK);
        return;
      case 'column':
        this.addItem(COLUMN_BREAK);
        return;
      case 'chpgn':
        this.addItem(PAGE_NUMBER);
        return;
      case 'chftnsep':
        this.addItem(SEPARATOR);
        return;
      case 'chftnsepc':
        this.addItem(CONTINUATION);
        return;
      case 'chftn': {
        const current = this.story;
        if (current.kind === 'footnote') {
          const place = this.addItem(FOOTNOTE_REF);
          if (place) current.noteRefs.push(place);
          return;
        }
        if (current.pendingNote) {
          current.pendingNote.run.content[current.pendingNote.index] = NOTHING;
          current.pendingNote = undefined;
        }
        const place = this.addItem({ type: 'footnoteReference', id: 0 });
        if (place) current.pendingNote = place;
        return;
      }
      case 'chatn': {
        const current = this.story;
        if (current.kind === 'comment') {
          this.addItem(ANNOTATION_REF);
          return;
        }
        if (current.pendingComment) {
          current.pendingComment.run.content[current.pendingComment.index] = NOTHING;
          current.pendingComment = undefined;
        }
        const place = this.addItem({ type: 'commentReference', id: 0 });
        if (place) current.pendingComment = place;
        return;
      }
      case 'ftnalt':
        this.story.endnote = true;
        return;
    }
  }

  softHyphen(): void {
    this.addItem(SOFT_HYPHEN);
  }

  noBreakHyphen(): void {
    this.addItem(NO_BREAK_HYPHEN);
  }

  // ------------------------------------------------------------------- groups

  private open(): void {
    this.skip = 0;
    this.lead = -1;
    if (this.depth >= MAX_DEPTH) {
      if (this.overflow === 0) this.warnings.push('Groups are nested deeper than the reader follows');
      this.overflow++;
      return;
    }
    const parent = this.st;
    parent.star = false;
    this.stack[this.depth++] = parent;
    const st: State = {
      dest: parent.dest,
      chp: parent.chp,
      chpOwn: false,
      pap: parent.pap,
      papOwn: false,
      uc: parent.uc,
      complex: parent.complex,
      assoc: parent.assoc,
      slot: 0,
      rtl: parent.rtl,
      code: parent.code,
      star: false,
      closer: C_NONE,
      payload: undefined,
    };
    this.st = st;
    // What a group is by where it stands.
    if (parent.dest === D.D_STYLESHEET) {
      st.dest = D.D_STYLE;
      st.closer = C_STYLE;
      st.payload = this.rawStyle = newRawStyle();
      this.buffers.push('');
    } else if (parent.dest === D.D_UPR) st.dest = D.D_UPR_CHILD;
    else if (parent.dest === D.D_SKIP) this.tk.skipGroup();
  }

  /** Returns false when the group that closed was the outermost one. */
  private close(): boolean {
    this.skip = 0;
    this.lead = -1;
    if (this.overflow > 0) {
      this.overflow--;
      return true;
    }
    const st = this.st;
    if (st.closer !== C_NONE) this.runCloser(st);
    if (st.dest === D.D_FONTTBL && this.depth <= this.fontDepth + 1) {
      // An entry of the font table ends with its group even when the semicolon is missing.
      this.finishFont();
      if (this.depth === this.fontDepth) this.fontDepth = -1;
    }
    if (this.depth === 0) return false;
    this.st = this.stack[--this.depth];
    return true;
  }

  private setCloser(closer: number, payload: unknown): void {
    const st = this.st;
    // A group is one thing; what it was before ends here.
    if (st.closer !== C_NONE) this.runCloser(st);
    st.closer = closer;
    st.payload = payload;
  }

  /** The rest of the group is passed over. */
  private skipRest(): void {
    this.st.dest = D.D_SKIP;
    this.tk.skipGroup();
  }

  private collect(dest: number, closer: number): void {
    this.setCloser(closer, undefined);
    this.st.dest = dest;
    this.buffers.push('');
  }

  private collected(): string {
    const text = this.buffers.pop();
    return text === undefined ? '' : text;
  }

  private pushStory(kind: StoryKind, closer: number, payload: { [key: string]: unknown }): boolean {
    if (this.openStories >= MAX_OPEN) {
      this.skipRest();
      return false;
    }
    this.openStories++;
    payload.outer = this.story;
    payload.borderTarget = this.borderTarget;
    this.setCloser(closer, payload);
    this.story = new Story(kind, this.format, this.tableBudget);
    const st = this.st;
    st.dest = D.D_NORMAL;
    st.code = false;
    this.plain();
    this.pard();
    return true;
  }

  /** Ends the story being read and goes back to the one around it; returns the blocks of the story that ended. */
  private popStory(payload: { [key: string]: unknown }): Block[] {
    const ended = this.story;
    // Text without a paragraph mark after it is a paragraph all the same.
    if (ended.children.length || ended.listText) ended.assembler.add(this.endParagraph(), this.paraDepth());
    const blocks = ended.finish();
    if (!blocks.length) blocks.push({ type: 'paragraph', props: EMPTY_PARA_PROPS, children: [] });
    linkFields(blocks);
    this.story = payload.outer as Story;
    this.borderTarget = payload.borderTarget as number;
    this.openStories--;
    return blocks;
  }

  // ------------------------------------------------------------- destinations

  /** A control word that says what its group is; false when it is not one. */
  private destination(word: string): boolean {
    const st = this.st;
    switch (word) {
      case 'fonttbl':
        st.dest = D.D_FONTTBL;
        this.fontDepth = this.depth;
        return true;
      case 'colortbl':
        st.dest = D.D_COLORTBL;
        return true;
      case 'stylesheet':
        this.setCloser(C_STYLESHEET, undefined);
        st.dest = D.D_STYLESHEET;
        return true;
      case 'listtable':
        st.dest = D.D_LISTTABLE;
        return true;
      case 'list':
        if (st.dest !== D.D_LISTTABLE) return false;
        this.list = { listId: 0, simple: false, hybrid: false, levels: [] };
        this.setCloser(C_LIST, this.list);
        st.dest = D.D_LIST;
        return true;
      case 'listlevel': {
        if (st.dest !== D.D_LIST && st.dest !== D.D_LFOLEVEL) return false;
        const level = newLevel();
        this.setCloser(C_LEVEL, { level, outer: this.level, override: st.dest === D.D_LFOLEVEL });
        this.level = level;
        st.dest = D.D_LISTLEVEL;
        return true;
      }
      case 'leveltext':
        if (st.dest !== D.D_LISTLEVEL) return false;
        st.dest = D.D_LEVELTEXT;
        return true;
      case 'listoverridetable':
        st.dest = D.D_OVERRIDETABLE;
        return true;
      case 'listoverride':
        if (st.dest !== D.D_OVERRIDETABLE) return false;
        this.override = { listId: 0, number: 0, levels: [] };
        this.setCloser(C_OVERRIDE, this.override);
        st.dest = D.D_OVERRIDE;
        return true;
      case 'lfolevel':
        if (st.dest !== D.D_OVERRIDE) return false;
        this.overrideLevel = { start: undefined, level: undefined, startOverride: false, formatOverride: false };
        this.setCloser(C_LFOLEVEL, this.overrideLevel);
        st.dest = D.D_LFOLEVEL;
        return true;

      case 'pict': {
        // The picture of a shape is the value of one of its properties.
        const shape = st.dest === D.D_SV && this.shapes.length ? this.shapes[this.shapes.length - 1] : undefined;
        if (st.dest !== D.D_NORMAL && !shape) {
          this.skipRest();
          return true;
        }
        const pict = newPict();
        this.setCloser(C_PICT, { pict, outer: this.pict, shape, text: shape !== undefined });
        this.pict = pict;
        st.dest = D.D_PICT;
        return true;
      }
      case 'shppict':
      case 'mmathPict':
        // The picture inside is read as any other.
        return true;
      case 'shp':
      case 'shpgrp':
      case 'do': {
        if ((st.dest !== D.D_NORMAL && st.dest !== D.D_SHPINST && st.dest !== D.D_SHP) || this.shapes.length >= MAX_OPEN) {
          this.skipRest();
          return true;
        }
        const shape = newShape(word === 'shpgrp');
        // A shape among the properties of another one is part of that group; one in text (that of a text box
        // included) stands in the text.
        this.setCloser(C_SHAPE, { shape, legacy: word === 'do', grouped: st.dest !== D.D_NORMAL });
        this.shapes.push(shape);
        st.dest = word === 'do' ? D.D_DO : D.D_SHP;
        return true;
      }
      case 'shpinst':
        if (st.dest !== D.D_SHP) return false;
        st.dest = D.D_SHPINST;
        return true;
      case 'sp':
        if (st.dest !== D.D_SHPINST) return false;
        st.dest = D.D_SP;
        return true;
      case 'sn':
        if (st.dest !== D.D_SP) return false;
        this.collect(D.D_SN, C_SHAPE_NAME);
        return true;
      case 'sv':
        if (st.dest !== D.D_SP) return false;
        this.collect(D.D_SV, C_SHAPE_VALUE);
        return true;
      case 'shptxt':
      case 'dptxbxtext': {
        if ((st.dest !== D.D_SHPINST && st.dest !== D.D_DO) || !this.shapes.length || this.boxDepth >= MAX_BOX_DEPTH) {
          this.skipRest();
          return true;
        }
        if (this.pushStory('textbox', C_BOX_TEXT, { shape: this.shapes[this.shapes.length - 1] })) this.boxDepth++;
        return true;
      }

      case 'field': {
        if (st.dest !== D.D_NORMAL || this.fields.length >= MAX_OPEN) {
          this.skipRest();
          return true;
        }
        const begin = this.fieldMark('begin', st.chp, st.rtl, undefined);
        this.setCloser(C_FIELD, undefined);
        this.fields.push({ separated: false, begin, separate: undefined, code: undefined, result: undefined });
        st.dest = D.D_FIELD;
        return true;
      }
      case 'fldinst':
        if (st.dest !== D.D_FIELD) return false;
        st.dest = D.D_NORMAL;
        st.code = true;
        return true;
      case 'fldrslt': {
        if (st.dest !== D.D_FIELD) return false;
        st.dest = D.D_NORMAL;
        st.code = false;
        const field = this.fields[this.fields.length - 1];
        if (field && !field.separated) {
          field.separated = true;
          field.separate = this.fieldMark('separate', st.chp, st.rtl, undefined);
        }
        return true;
      }

      case 'footnote': {
        if (st.dest !== D.D_NORMAL) return false;
        const outer = this.story;
        let place = outer.pendingNote;
        outer.pendingNote = undefined;
        if (!place) place = this.addItem({ type: 'footnoteReference', id: 0, customMark: true });
        this.pushStory('footnote', C_NOTE, { place });
        return true;
      }
      case 'annotation': {
        if (st.dest !== D.D_NORMAL) return false;
        const outer = this.story;
        let place = outer.pendingComment;
        outer.pendingComment = undefined;
        if (!place) place = this.addItem({ type: 'commentReference', id: 0 });
        this.pushStory('comment', C_COMMENT, { place, ref: undefined });
        return true;
      }
      case 'atnid':
        this.collect(D.D_ATNID, C_COMMENT_INITIALS);
        return true;
      case 'atnauthor':
        this.collect(D.D_ATNAUTHOR, C_COMMENT_AUTHOR);
        return true;
      case 'atrfstart':
        this.collect(D.D_ATRFSTART, C_RANGE_START);
        return true;
      case 'atrfend':
        this.collect(D.D_ATRFEND, C_RANGE_END);
        return true;
      case 'atnref':
        this.collect(D.D_ATNREF, C_COMMENT_REF);
        return true;
      case 'bkmkstart':
        this.collect(D.D_BKMKSTART, C_BOOKMARK_START);
        return true;
      case 'bkmkend':
        this.collect(D.D_BKMKEND, C_BOOKMARK_END);
        return true;
      case 'falt':
        if (st.dest !== D.D_FONTTBL) return false;
        this.collect(D.D_FALT, C_FONT_ALT);
        return true;
      case 'panose':
        if (st.dest !== D.D_FONTTBL) return false;
        this.collect(D.D_PANOSE, C_PANOSE);
        return true;

      case 'listtext':
      case 'pntext':
        if (st.dest !== D.D_NORMAL) return false;
        this.setCloser(C_LIST_TEXT, undefined);
        this.story.beginCapture();
        return true;
      case 'nesttableprops':
        // The description of a row of a table in a cell; that of the row around it is put back afterwards.
        this.setCloser(C_ROW_PROPS, this.story.row.save());
        return true;
      case 'object':
        if (st.dest !== D.D_NORMAL) return false;
        st.dest = D.D_OBJECT;
        return true;
      case 'result':
        if (st.dest !== D.D_OBJECT) return false;
        st.dest = D.D_NORMAL;
        return true;
      case 'upr':
        this.setCloser(C_DISCARD, undefined);
        st.payload = st.dest;
        st.dest = D.D_UPR;
        return true;
      case 'ud': {
        if (st.dest !== D.D_UPR_CHILD) return false;
        // The Unicode form of what the group before it holds for old readers: this is the one that is read.
        const parent = this.stack[this.depth - 1];
        st.dest = parent && typeof parent.payload === 'number' ? parent.payload : D.D_NORMAL;
        return true;
      }
    }

    const header = HEADERS[word];
    if (header) {
      if (this.story !== this.body || st.dest !== D.D_NORMAL) {
        this.skipRest();
        return true;
      }
      this.pushStory('header', C_HEADER, { list: header.list, kind: header.kind });
      return true;
    }
    const separator = SEPARATORS[word];
    if (separator) {
      this.pushStory('footnote', C_SEPARATOR, { separator });
      return true;
    }
    if (SKIP.has(word)) {
      this.skipRest();
      return true;
    }
    return false;
  }

  private runCloser(st: State): void {
    const closer = st.closer;
    const payload = st.payload as { [key: string]: unknown };
    st.closer = C_NONE;
    st.payload = undefined;
    switch (closer) {
      case C_STYLE: {
        const raw = payload as unknown as RawStyle;
        let name = this.collected();
        const end = name.lastIndexOf(';');
        if (end >= 0) name = name.slice(0, end);
        raw.name = name.trim().slice(0, 255);
        raw.chp = st.chp;
        raw.pap = st.pap;
        raw.rtl = st.rtl;
        if (this.rawStyles.length < MAX_STYLES) this.rawStyles.push(raw);
        this.rawStyle = undefined;
        return;
      }
      case C_STYLESHEET:
        // The styles are made when the first paragraph needs them, or at the end: by then the fonts are known.
        // What was made from formatting before the style sheet (in a file that has text in front of it) is made again.
        this.sheet = undefined;
        this.plainChp.cache = undefined;
        this.plainPap.cache = undefined;
        return;
      case C_LIST: {
        const list = payload as unknown as ListState;
        if (this.lists.length < MAX_LISTS) this.lists.push(list);
        this.list = undefined;
        return;
      }
      case C_LEVEL: {
        const level = payload.level as LevelState;
        this.level = payload.outer as LevelState | undefined;
        // What the group says about text and paragraph is the formatting of the number.
        const run = diffRunProps(fullRunProps(this.formatBare, st.chp, false), fullRunProps(this.formatBare, this.plainChp, false));
        if (Object.keys(run).length) level.run = run;
        const pap = st.pap;
        const para: ParaProps = {};
        if (pap.left !== 0 || pap.first !== 0) para.indent = pap.first < 0 ? { left: pap.left, hanging: -pap.first } : { left: pap.left, firstLine: pap.first };
        if (pap.tabs && pap.tabs.length) para.tabs = pap.tabs;
        if (Object.keys(para).length) level.para = para;
        const font = this.fonts.get(st.chp.font < 0 ? this.defaultFont : st.chp.font);
        level.codepage = font && font.codepage !== undefined ? font.codepage : this.codepage;
        if (payload.override) {
          if (this.overrideLevel) this.overrideLevel.level = level;
        } else if (this.list && this.list.levels.length < 9) this.list.levels.push(level);
        return;
      }
      case C_OVERRIDE: {
        const override = payload as unknown as OverrideState;
        if (this.overrides.length < MAX_LISTS) {
          this.overrides.push(override);
          if (this.lists.some((list) => list.listId === override.listId)) this.knownLists.add(override.number);
        }
        this.override = undefined;
        return;
      }
      case C_LFOLEVEL: {
        const entry = payload as unknown as OverrideState['levels'][number];
        if (this.override && this.override.levels.length < 9) this.override.levels.push(entry);
        this.overrideLevel = undefined;
        return;
      }
      case C_PICT: {
        const pict = payload.pict as PictState;
        this.pict = payload.outer as PictState | undefined;
        const item = pictItem(pict, this.store);
        if (pict.tooLarge) this.warnings.push('A picture is larger than the reader keeps and was left out');
        if (!item) return;
        const shape = payload.shape as ShapeState | undefined;
        if (shape) {
          if (item.image && shape.images.length < MAX_SHAPE_PARTS) {
            shape.images.push(item.image);
            shape.crops.push(item.crop);
          }
          return;
        }
        // The state of the group around the picture: the picture is text of it.
        const parent = this.stack[this.depth - 1];
        if (!parent || parent.dest !== D.D_NORMAL) return;
        const chp = parent.chp;
        this.story.addItem(item, this.runPropsFor(chp, parent.rtl, this.paraStyle()), this.fullRun(chp, parent.rtl), this.wrapper(chp));
        return;
      }
      case C_SHAPE: {
        const shape = payload.shape as ShapeState;
        const at = this.shapes.lastIndexOf(shape);
        if (at >= 0) this.shapes.length = at;
        const outer = this.shapes[this.shapes.length - 1];
        if (payload.grouped) {
          if (!outer) return;
          // A shape of a group: what it shows is shown by the group.
          for (let i = 0; i < shape.images.length && outer.images.length < MAX_SHAPE_PARTS; i++) {
            outer.images.push(shape.images[i]);
            outer.crops.push(shape.crops[i]);
          }
          for (const box of shape.boxes) if (outer.boxes.length < MAX_SHAPE_PARTS) outer.boxes.push(box);
          return;
        }
        // A drawing object of Word 95 that is not a text box is lines and curves the reader does not draw.
        if (payload.legacy && !shape.boxes.length) return;
        const parent = this.stack[this.depth - 1];
        if (!parent || parent.dest !== D.D_NORMAL) return;
        const chp = parent.chp;
        this.story.addItem(shapeItem(shape, payload.legacy === true), this.runPropsFor(chp, parent.rtl, this.paraStyle()), this.fullRun(chp, parent.rtl), undefined);
        return;
      }
      case C_SHAPE_NAME: {
        const shape = this.shapes[this.shapes.length - 1];
        const name = this.collected().trim();
        if (shape) shape.name = name;
        return;
      }
      case C_SHAPE_VALUE: {
        const shape = this.shapes[this.shapes.length - 1];
        const value = this.collected();
        if (shape && shape.name && shape.props.size < 1000) shape.props.set(shape.name, value);
        return;
      }
      case C_BOX_TEXT: {
        const blocks = this.popStory(payload);
        this.boxDepth--;
        const shape = payload.shape as ShapeState;
        if (shape.boxes.length < MAX_SHAPE_PARTS) shape.boxes.push(blocks);
        return;
      }
      case C_FIELD: {
        const field = this.fields.pop();
        // The mark that ends a field has the formatting of its result, or of its code when it has none; that of the
        // text around the field when it has neither.
        const parent = this.stack[this.depth - 1] || st;
        this.fieldMark('end', parent.chp, parent.rtl, field ? field.result || field.code : undefined);
        return;
      }
      case C_NOTE: {
        const ended = this.story;
        const blocks = this.popStory(payload);
        const endnote = ended.endnote;
        const id = endnote ? ++this.endnoteId : ++this.footnoteId;
        const place = payload.place as Place | undefined;
        if (place) {
          const before = place.run.content[place.index] as { customMark?: boolean };
          const ref: RunItem = { type: endnote ? 'endnoteReference' : 'footnoteReference', id };
          if (before.customMark) ref.customMark = true;
          place.run.content[place.index] = ref;
        }
        if (endnote) for (const mark of ended.noteRefs) mark.run.content[mark.index] = ENDNOTE_REF;
        const model = this.model;
        const notes = endnote ? model.endnotes || (model.endnotes = collection<Note>('rtf:endnotes')) : model.footnotes || (model.footnotes = collection<Note>('rtf:footnotes'));
        notes.items.set(id, { id, blocks });
        return;
      }
      case C_COMMENT: {
        const ended = this.story;
        const blocks = this.popStory(payload);
        const id = this.commentId++;
        const place = payload.place as Place | undefined;
        if (place) place.run.content[place.index] = { type: 'commentReference', id };
        const comment: Comment = { id, blocks };
        if (this.commentAuthor) comment.author = this.commentAuthor;
        if (this.commentInitials) comment.initials = this.commentInitials;
        this.commentAuthor = this.commentInitials = undefined;
        const model = this.model;
        (model.comments || (model.comments = collection<Comment>('rtf:comments'))).items.set(id, comment);
        // The marks of the range the comment is about get its id.
        const key = ended.commentRef;
        const marks = key === undefined ? undefined : this.ranges.get(key);
        if (marks) {
          for (const mark of marks) mark.id = String(id);
          this.ranges.delete(key!);
        }
        return;
      }
      case C_COMMENT_REF:
        this.story.commentRef = this.collected().trim();
        return;
      case C_HEADER: {
        const blocks = this.popStory(payload);
        const list = payload.list as 'headers' | 'footers';
        const kind = payload.kind as HeaderFooterKind;
        const id = (list === 'headers' ? 'hdr:' : 'ftr:') + this.sectionIndex + ':' + kind;
        this.model[list].set(id, story(list === 'headers' ? 'rtf:header' : 'rtf:footer', blocks));
        this.sectionRefs[list][kind] = id;
        return;
      }
      case C_SEPARATOR: {
        const blocks = this.popStory(payload);
        const separator = payload.separator as { list: 'footnotes' | 'endnotes'; id: number; kind: string };
        this.separators.push({ list: separator.list, note: { id: separator.id, kind: separator.kind, blocks } });
        return;
      }
      case C_BOOKMARK_START: {
        const name = this.collected().trim();
        if (!name) return;
        const id = String(this.nextBookmark++);
        this.bookmarks.set(name, id);
        this.addMark({ type: 'bookmarkStart', id, name });
        return;
      }
      case C_BOOKMARK_END: {
        const name = this.collected().trim();
        const id = this.bookmarks.get(name);
        // An end without a start is the end of nothing.
        if (id === undefined) return;
        this.bookmarks.delete(name);
        this.addMark({ type: 'bookmarkEnd', id });
        return;
      }
      case C_COMMENT_INITIALS:
        this.commentInitials = this.collected().trim();
        return;
      case C_COMMENT_AUTHOR:
        this.commentAuthor = this.collected().trim();
        return;
      case C_RANGE_START:
      case C_RANGE_END: {
        const key = this.collected().trim();
        if (!key) return;
        const mark = { type: closer === C_RANGE_START ? ('commentRangeStart' as const) : ('commentRangeEnd' as const), id: 'atrf:' + key };
        let marks = this.ranges.get(key);
        if (!marks) this.ranges.set(key, (marks = []));
        if (marks.length < 4) marks.push(mark);
        this.addMark(mark);
        return;
      }
      case C_LIST_TEXT:
        this.story.endCapture();
        return;
      case C_ROW_PROPS:
        this.story.row.restore(payload);
        return;
      case C_FONT_ALT: {
        const name = this.collected().trim();
        if (this.font && name) this.font.altName = name.slice(0, 255);
        return;
      }
      case C_PANOSE: {
        const text = this.collected().trim();
        if (this.font && text.length === 20) this.font.panose = text.toUpperCase();
        return;
      }
    }
  }

  /** A mark between the characters of the text being read. */
  private addMark(mark: Inline): void {
    // The group of the mark has closed or is closing: the text it stands in is the story being read.
    this.story.addInline(mark);
  }

  // -------------------------------------------------------------------- fonts

  private fontWord(word: string, param: number): boolean {
    if (word === 'f') {
      this.finishFont();
      this.font = { index: param, name: '', altName: undefined, charset: undefined, codepage: undefined, family: undefined, pitch: undefined, panose: undefined };
      return true;
    }
    const font = this.font;
    switch (word) {
      case 'fcharset':
        if (font) {
          font.charset = param;
          if (font.codepage === undefined) font.codepage = charsetCodepage(param);
        }
        return true;
      case 'cpg':
        if (font && (isKnownRtfCodepage(param) || param === SYMBOL_CODEPAGE)) font.codepage = param;
        return true;
      case 'fprq':
        if (font) font.pitch = param;
        return true;
      case 'fbias':
      case 'flomajor':
      case 'fhimajor':
      case 'fdbmajor':
      case 'fbimajor':
      case 'flominor':
      case 'fhiminor':
      case 'fdbminor':
      case 'fbiminor':
        return true;
    }
    const family = FAMILIES[word];
    if (family !== undefined) {
      if (font && (font.family === undefined || word !== 'fbidi')) font.family = family;
      return true;
    }
    return false;
  }

  private finishFont(): void {
    const font = this.font;
    this.font = undefined;
    if (!font) return;
    font.name = font.name.trim();
    if (!font.name || this.fonts.has(font.index) || this.fonts.size >= MAX_FONTS) return;
    this.fonts.set(font.index, font);
  }

  // -------------------------------------------------------------------- words

  private word(word: string, param: number, hasParam: boolean): void {
    if (this.skip > 0) {
      // What stands in for a "\u" is characters, and a control word that is a character is one of them. Any other
      // control word ends it and is read as it is: a paragraph or cell mark is never passed over.
      if (isCharacterWord(word)) {
        this.skip--;
        return;
      }
      this.skip = 0;
    }
    this.lead = -1;
    this.echo = -1;
    const st = this.st;
    const starred = st.star;
    st.star = false;

    // Words that mean something of their own inside a table of the file.
    switch (st.dest) {
      case D.D_SKIP:
        return;
      case D.D_UPR_CHILD:
        if (word !== 'ud') {
          this.skipRest();
          return;
        }
        break;
      case D.D_FONTTBL:
        if (this.fontWord(word, param)) return;
        break;
      case D.D_COLORTBL:
        if (word === 'red' || word === 'green' || word === 'blue') {
          const value = param < 0 ? 0 : param > 255 ? 255 : param;
          if (word === 'red') this.red = value;
          else if (word === 'green') this.green = value;
          else this.blue = value;
          this.colorStated = true;
          return;
        }
        break;
      case D.D_STYLE:
        if (this.rawStyle && styleWord(this.rawStyle, word, param)) return;
        break;
      case D.D_PICT:
        if (this.pict && pictWord(this.pict, word, param)) return;
        break;
      case D.D_SHP:
      case D.D_SHPINST:
        if (this.shapes.length && shapeWord(this.shapes[this.shapes.length - 1], word, param)) return;
        break;
      case D.D_DO:
        if (this.shapes.length && shapeWord(this.shapes[this.shapes.length - 1], word, param)) return;
        if (isLegacyDrawingWord(word) && word !== 'dptxbxtext') return;
        break;
      case D.D_LIST: {
        const list = this.list;
        if (list) {
          if (word === 'listid') {
            list.listId = param;
            return;
          }
          if (word === 'listsimple') {
            list.simple = !hasParam || param !== 0;
            return;
          }
          if (word === 'listhybrid') {
            list.hybrid = true;
            return;
          }
          if (word === 'listtemplateid' || word === 'listrestarthdn' || word === 'liststyleid') return;
        }
        break;
      }
      case D.D_LISTLEVEL:
        if (this.level && levelWord(this.level, word, param)) return;
        break;
      case D.D_OVERRIDE: {
        const override = this.override;
        if (override) {
          if (word === 'listid') {
            override.listId = param;
            return;
          }
          if (word === 'ls') {
            override.number = param;
            return;
          }
          if (word === 'listoverridecount') return;
        }
        break;
      }
      case D.D_LFOLEVEL: {
        const entry = this.overrideLevel;
        if (entry) {
          if (word === 'listoverridestartat') {
            entry.startOverride = true;
            return;
          }
          if (word === 'listoverrideformat') {
            entry.formatOverride = true;
            return;
          }
          if (word === 'levelstartat') {
            entry.start = param;
            return;
          }
        }
        break;
      }
      case D.D_FIELD:
        if (word === 'flddirty' || word === 'fldedit' || word === 'fldlock' || word === 'fldpriv') return;
        break;
    }

    if (DESTINATIONS.has(word) && this.destination(word)) return;
    const handler: WordHandler | undefined = WORDS.get(word);
    if (handler) {
      handler(this, param, hasParam, word);
      return;
    }
    if (starred) {
      // A group marked as one that may be passed over, of a kind the reader does not know.
      this.skipRest();
      return;
    }
    this.count('rtf/' + word);
  }

  private symbol(code: number): void {
    const st = this.st;
    switch (code) {
      case 0x2a:
        st.star = true;
        return;
      case 0x7e:
        if (this.skip > 0) this.skip--;
        else this.char(String.fromCharCode(0xa0));
        return;
      case 0x2d:
        if (this.skip > 0) this.skip--;
        else this.softHyphen();
        return;
      case 0x5f:
        if (this.skip > 0) this.skip--;
        else this.noBreakHyphen();
        return;
      case 0x5c:
      case 0x7b:
      case 0x7d:
        if (this.skip > 0) this.skip--;
        else this.char(String.fromCharCode(code));
        return;
      case 0x0a:
      case 0x0d:
        // A backslash at the end of a line is a paragraph mark.
        this.skip = 0;
        this.par();
        return;
      case 0x09:
        this.skip = 0;
        this.addItem(TAB);
        return;
    }
    // "\|", "\:" and the rest stand for nothing that shows.
  }

  // --------------------------------------------------------------------- read

  read(): DocModel {
    const tk = this.tk;
    // The file starts with a group; its "{" opens the state the reader starts with.
    tk.next();
    let open = true;
    while (open) {
      const token = tk.next();
      switch (token) {
        case T_OPEN:
          this.open();
          break;
        case T_CLOSE:
          open = this.close();
          break;
        case T_WORD:
          this.word(tk.word, tk.param, tk.hasParam);
          break;
        case T_TEXT:
          this.text(tk.start, tk.stop);
          break;
        case T_HEX:
          this.hex(tk.code);
          break;
        case T_SYMBOL:
          this.symbol(tk.code);
          break;
        case T_BIN:
          this.skip = 0;
          if (this.st.dest === D.D_PICT && this.pict) pictBin(this.pict, tk.bytes, tk.start, tk.stop);
          break;
        case T_EOF:
          // A file that is cut short: every group that is open ends here.
          this.overflow = 0;
          while (this.close()) {
            // Each round closes one group.
          }
          open = false;
          break;
      }
    }
    return this.finish();
  }

  private finish(): DocModel {
    const model = this.model;
    const body = this.body;
    this.story = body;
    this.styles();
    if (body.children.length || body.listText) body.assembler.add(this.endParagraph(), this.paraDepth());
    const blocks = body.finish();
    linkFields(blocks);
    model.body.blocks = blocks;
    model.body.section = this.section();

    for (const { list, note } of this.separators) {
      const notes = model[list];
      if (notes && !notes.items.has(note.id)) notes.items.set(note.id, note);
    }

    model.numbering = buildNumbering(this.lists, this.overrides, decodeByte);
    for (const font of this.fonts.values()) {
      if (model.fonts.has(font.name)) continue;
      const info: FontInfo = { name: font.name };
      if (font.altName) info.altName = font.altName;
      if (font.panose) info.panose = font.panose;
      if (font.charset !== undefined) info.charset = hex2(font.charset);
      if (font.family !== undefined) info.family = font.family;
      if (font.pitch !== undefined && PITCHES[font.pitch] !== undefined) info.pitch = PITCHES[font.pitch];
      model.fonts.set(font.name, info);
    }

    if (this.tableBudget.exceeded) this.warnings.push('The file opens more tables than the reader builds for a file of its size; the rest is read without them');

    const settings = model.settings;
    const dop = this.dop;
    settings.defaultTabStop = dop.defaultTab;
    settings.evenAndOddHeaders = dop.facingPages;
    settings.mirrorMargins = dop.mirrorMargins;
    settings.autoHyphenation = dop.autoHyphenation;
    settings.trackRevisions = dop.trackRevisions;
    // "\enforceprot1" applies the protection "\protlevel" (or an older word) names; "\enforceprot0" keeps it without
    // applying it. A file that does not say either way is protected when it uses one of the older words, as files
    // written before "\enforceprot" existed do; a level alone protects nothing. UNSURE: from the RTF specification
    // (1.9.1), not compared with files Word protected.
    const kind = dop.enforceProtection === undefined ? dop.protection : dop.protectionLevel || dop.protection;
    if (kind !== undefined) settings.protection = { edit: kind, enforced: dop.enforceProtection !== false };
    if (this.store.media.size) model.media = this.store.media;
    return model;
  }
}

/**
 * Reads an RTF document. Throws an Error that starts with "Invalid rtf:" for
 * bytes that do not start with "{\rtf" and for a document with more
 * paragraphs than the reader takes; a file that is cut short is read as far
 * as it goes.
 */
export function readRtf(bytes: Uint8Array, options: ReadRtfOptions = {}): DocModel {
  return new RtfReader(bytes, options).read();
}
