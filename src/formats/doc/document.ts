// Reads a Word binary file (.doc, Word 97 to 2003; MS-DOC) into the document
// model that the docx reader fills too.
//
// The file is a compound file with three streams that matter: WordDocument
// (the header and the text), a table stream (where everything else is listed)
// and Data (pictures, long property lists). The reader goes through them in
// this order: header (fib.ts), text (piece-table.ts), fonts and styles
// (styles.ts), lists (lists.ts), then the stories: the main text, the notes,
// the headers and footers, the comments and the text boxes. A story is cut
// into paragraphs at the paragraph marks; each paragraph gets its properties
// from the place of its mark in the file (fkp.ts, props.ts) and its runs from
// the places of its characters; paragraphs in tables are gathered into table
// blocks (tables.ts).
//
// How the result differs from a model read from docx:
//   - There is no XML. `Story.source` is empty and no block carries a source
//     span; `Story.part` names the story ("doc:body", "doc:header", ...);
//     `rootName` is empty; drawings carry a stand-in element.
//   - Header and footer ids are made up: "hdr:<section>:<kind>" and
//     "ftr:<section>:<kind>", with the section counted from 0.
//   - Pictures are in `DocModel.media`, by the part name their references
//     give as target ("/doc/media/image1.png").
//   - Note ids are the number of the note counted from 1, comment ids from 0.
//   - What the reader passes over is counted in `unknown` ("sprm/0x<code>").
//     The file itself is never changed: it is saved as it was read.
//
// Nothing in the file is trusted: offsets are checked against the streams,
// counts are bounded, and loops advance by at least one unit of what they
// read. A file without usable text is an error; a damaged structure beside
// the text (a list, a drawing, a header) is left out with a warning.

import type { CfbFile } from '../../cfb';
import {
  EMPTY_PARA_PROPS,
  EMPTY_RUN_PROPS,
  TOGGLE_RUN_PROPS,
  defaultSettings,
  emptyTheme,
  type Block,
  type Body,
  type Comment,
  type DocModel,
  type DocRelationship,
  type DocSettings,
  type Inline,
  type Note,
  type ParaProps,
  type Paragraph,
  type RangeMark,
  type RevisionInfo,
  type Run,
  type RunItem,
  type RunProps,
  type SectionProps,
  type Story,
  type StoryCollection,
  type StyleType,
  type Table,
  type Wrapper,
} from '../../model/doc/types';
import { styleChain } from '../docx/resolve';
import { bytesKey, invalid, readPlc, readSttb, u16, u32, u8, utf16, type Plc } from './bytes';
import { MediaStore, floatingDrawing, inlinePicture, readDrawings, type Drawings } from './drawings';
import { FC, readFib, type Fib } from './fib';
import { linkDocFields } from './fields';
import { readChpxTable, readPapxTable, type RunTable } from './fkp';
import { readNumbering } from './lists';
import { readPieceTable, type Piece, type PieceTable } from './piece-table';
import { HEADER_ORDER, charProps, mirrored, paraProps, sectionProps, type CharProps, type ParaInfo, type PropContext } from './props';
import { finishStyles, readFonts, readStyleSheet, type StyleSheet } from './styles';
import { BlockAssembler } from './tables';

export { UnsupportedDocError } from './bytes';

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

const PARAGRAPH_MARK = 0x0d;
const CELL_MARK = 0x07;
const SECTION_MARK = 0x0c;

/**
 * Paragraphs one document may have. Word itself stops far below this; a file
 * of nothing but paragraph marks would otherwise take a hundred times its
 * size in memory.
 */
const MAX_PARAGRAPHS = 2000000;
/** Bytes of the document properties up to Word 2002, and with what Word 2003 added (MS-DOC 2.7.7, 2.7.8). */
const DOP2002_SIZE = 594;
const DOP2003_SIZE = 616;

export interface ReadDocOptions {
  /** Paragraphs the document may have before it is refused; 2,000,000 when missing. A host with little memory may lower it. */
  maxParagraphs?: number;
}

/** Text boxes in text boxes that are read; a genuine file has none. */
const MAX_BOX_DEPTH = 4;

type StoryKind = 'main' | 'footnote' | 'endnote' | 'comment' | 'header' | 'textbox';

/** What a level of styles says about a toggle: on, off, or the opposite of the levels below (character styles only). */
type ToggleState = 'on' | 'off' | 'flip';
type ToggleKey = (typeof TOGGLE_RUN_PROPS)[number];
type ToggleLevel = Map<ToggleKey, ToggleState>;

/** Something that sits between two characters: a bookmark end, a note with a mark of its own. */
interface Mark {
  cp: number;
  /** Order among the marks at the same position: ends before starts. */
  order: number;
  inline?: RangeMark;
  note?: RunItem;
  /** Already put into a paragraph: a mark at the border of two stories goes into the first only. */
  used?: boolean;
}

/** The items of a run as the contents of one or more runs: each field mark alone, the items between them together. */
function splitAtFieldMarks(items: RunItem[]): RunItem[][] {
  if (!items.length) return [];
  let marks = 0;
  for (const item of items) if (item.type === 'fieldChar') marks++;
  if (marks === 0 || items.length === 1) return [items];
  const out: RunItem[][] = [];
  let current: RunItem[] | undefined;
  for (const item of items) {
    if (item.type === 'fieldChar') {
      out.push([item]);
      current = undefined;
    } else {
      if (!current) out.push((current = []));
      current.push(item);
    }
  }
  return out;
}

function story(part: string, blocks: Block[]): Story {
  return { part, rootName: '', rootAttrs: [], blocks, relationships: NO_RELATIONSHIPS, source: '' };
}

function collection<T>(part: string): StoryCollection<T> {
  return { part, rootName: '', rootAttrs: [], relationships: NO_RELATIONSHIPS, source: '', items: new Map() };
}

class DocReader {
  readonly warnings: string[] = [];
  readonly unknown = new Map<string, number>();
  private readonly word: Uint8Array;
  private readonly table: Uint8Array;
  private readonly data: Uint8Array;
  private readonly fib: Fib;
  private readonly pieces: PieceTable;
  private readonly text: string;
  private readonly chpx: RunTable;
  private readonly papx: RunTable;
  private readonly sheet: StyleSheet;
  private readonly ctx: PropContext;
  private readonly model: DocModel;
  private readonly store = new MediaStore();
  private drawings: Drawings | undefined;

  /** Starts of the parts of the text, in characters: main text, footnotes, headers, macros, comments, endnotes, text boxes, header text boxes, end. */
  private readonly starts: number[];

  private readonly paraCache = new Map<string, ParaInfo>();
  private readonly paraFinal = new WeakMap<ParaProps, Map<RunProps, ParaProps>>();
  private readonly charCache = new Map<number, CharProps>();
  private readonly charByKey = new Map<string, CharProps>();
  private readonly inverted = new WeakMap<CharProps, Map<string, RunProps>>();
  /** What the reader made each set of run properties of the model from, to make them again inside a table with a style. */
  private readonly runSource = new WeakMap<RunProps, CharProps>();
  /** Run properties of a paragraph mark that was deleted or inserted, by the character properties and the properties without that. */
  private readonly revisedMarks = new WeakMap<CharProps, Map<RunProps, RunProps>>();
  private readonly toggleLevels = new Map<string, ToggleLevel | null>();
  private readonly tableParas = new WeakMap<ParaProps, Map<string, ParaProps>>();

  /** Character positions that end a section (all but the last section), in order, and whether a position is one. */
  private readonly sectionEnds: number[] = [];
  private readonly sectionMarks = new Set<number>();
  private readonly footnoteRefs = new Map<number, number>();
  private readonly endnoteRefs = new Map<number, number>();
  private readonly commentRefs = new Map<number, number>();
  /** Place of the Spa of the drawing object at a character position. */
  private readonly shapeAt = new Map<number, number>();
  private marks: Mark[] = [];
  /** Text boxes of the main text and of the headers: the list of their texts and the index of each by the id of its shape. */
  private readonly boxes: { plc: Plc; start: number; byShape: Map<number, number> }[] = [];
  /** Texts of text boxes that a shape has been given: each is read once and belongs to that shape alone. */
  private readonly boxesTaken = new Set<string>();
  private paragraphs = 0;
  /** Bytes of piece modifiers that may still be read (see modifierOf). */
  private modifierBudget = 64 * 1024 * 1024;
  private sections: SectionProps[] = [];
  private readonly maxParagraphs: number;

  constructor(cfb: CfbFile, options: ReadDocOptions) {
    const limit = options.maxParagraphs;
    this.maxParagraphs = limit !== undefined && limit >= 1 && limit < MAX_PARAGRAPHS ? Math.floor(limit) : MAX_PARAGRAPHS;
    if (!cfb.has('WordDocument')) throw invalid('the file has no WordDocument stream');
    const word = (this.word = cfb.stream('WordDocument'));
    const fib = (this.fib = readFib(word));
    // The header names its table stream; a file that has only the other one is read from that.
    const other = fib.tableStream === '1Table' ? '0Table' : '1Table';
    const name = cfb.has(fib.tableStream) ? fib.tableStream : cfb.has(other) ? other : undefined;
    if (name === undefined) throw invalid('the file has no table stream');
    const table = (this.table = cfb.stream(name));
    this.data = cfb.has('Data') ? cfb.stream('Data') : new Uint8Array(0);

    const [clxAt, clxSize] = fib.fcLcb(FC.Clx);
    this.pieces = readPieceTable(table, clxAt, clxSize, word);
    this.text = this.pieces.text;

    // The parts of the text follow each other; none reaches beyond the text there is.
    const length = this.text.length;
    const starts = [0];
    for (const count of [fib.ccpText, fib.ccpFtn, fib.ccpHdd, fib.ccpMcr, fib.ccpAtn, fib.ccpEdn, fib.ccpTxbx, fib.ccpHdrTxbx]) {
      starts.push(Math.min(length, starts[starts.length - 1] + count));
    }
    if (fib.ccpText > length) this.warnings.push('The file header states more text than the file holds');
    this.starts = starts;

    const [fontsAt, fontsSize] = fib.fcLcb(FC.SttbfFfn);
    const fonts = readFonts(table, fontsAt, fontsSize);
    const [stshAt, stshSize] = fib.fcLcb(FC.Stshf);
    const { sheet, raw } = readStyleSheet(table, stshAt, stshSize);
    this.sheet = sheet;
    const [authorsAt, authorsSize] = fib.fcLcb(FC.SttbfRMark);
    this.ctx = {
      fontName: (ftc) => fonts.names[ftc] || undefined,
      styleId: (istd) => sheet.ids[istd],
      authors: readSttb(table, authorsAt, authorsSize).strings,
      data: this.data,
      unknown: this.unknown,
      warnings: this.warnings,
    };
    finishStyles(this.ctx, table, sheet, raw, fonts.names);

    const [chpxAt, chpxSize] = fib.fcLcb(FC.PlcfBteChpx);
    this.chpx = readChpxTable(word, table, chpxAt, chpxSize);
    const [papxAt, papxSize] = fib.fcLcb(FC.PlcfBtePapx);
    this.papx = readPapxTable(word, table, papxAt, papxSize);
    if (this.chpx.broken || this.papx.broken) this.warnings.push('Some formatting of the document could not be read');

    this.model = {
      body: story('doc:body', []) as Body,
      styles: sheet.styles,
      numbering: readNumbering(this.ctx, table, fib.fcLcb(FC.PlfLst), fib.fcLcb(FC.PlfLfo)),
      theme: emptyTheme(),
      settings: this.settings(),
      fonts: fonts.fonts,
      headers: new Map(),
      footers: new Map(),
      strict: false,
      warnings: this.warnings,
      unknown: this.unknown,
    };
  }

  /** The left cell margin a table style gives, following the styles it is based on. */
  private readonly styleMargin = (id: string | undefined): number | undefined => {
    let left: number | undefined;
    for (const style of styleChain(this.sheet.styles, id, 'table')) {
      const margin = style.table && style.table.cellMargins && style.table.cellMargins.left;
      if (margin && margin.type === 'dxa') left = margin.value;
    }
    return left;
  };

  // ------------------------------------------------------------------ settings

  /** Document properties (Dop, MS-DOC 2.7): the few the model has a place for. */
  private settings(): DocSettings {
    const settings = defaultSettings();
    // What these files were made with: layout treats tables and spacing as Word 2003 does.
    settings.compatSettings.set('compatibilityMode', '11');
    const [at, size] = this.fib.fcLcb(FC.Dop);
    const b = this.table;
    if (size < 12 || at + size > b.length) return settings;
    settings.evenAndOddHeaders = (u8(b, at) & 0x01) !== 0;
    settings.autoHyphenation = (u8(b, at + 5) & 0x10) !== 0;
    settings.trackRevisions = (u8(b, at + 5) & 0x80) !== 0;
    settings.mirrorMargins = (u8(b, at + 6) & 0x20) !== 0;
    // DopBase (MS-DOC 2.7.2): three bits, each a protection by itself. fLockRev: only tracked changes may be made;
    // fLockAtn: only comments; fProtEnabled: only form fields. Dop2003.fTreatLockAtnAsReadOnly says that fLockAtn
    // stands for a read-only protection. UNSURE: when more than one bit is set, the order here (tracked changes,
    // then comments, then forms) decides; not compared with files Word protected. UNSURE: the place of
    // fTreatLockAtnAsReadOnly (the first bit after the 594 bytes of a Dop2002) is from the specification as
    // remembered; a wrong place would name a comment protection read-only or the other way round, and both
    // open locked here.
    const lockRev = (u8(b, at + 7) & 0x40) !== 0;
    const lockAtn = (u8(b, at + 6) & 0x10) !== 0;
    const protEnabled = (u8(b, at + 7) & 0x02) !== 0;
    const atnReadOnly = size >= DOP2003_SIZE && (u8(b, at + DOP2002_SIZE) & 0x01) !== 0;
    const lock = lockRev ? 'trackedChanges' : lockAtn ? (atnReadOnly ? 'readOnly' : 'comments') : protEnabled ? 'forms' : undefined;
    if (lock !== undefined) settings.protection = { edit: lock, enforced: true };
    const tab = u16(b, at + 10);
    if (tab > 0 && tab <= 31680) settings.defaultTabStop = tab;
    return settings;
  }

  // ---------------------------------------------------------------- properties

  private paraInfo(markCp: number): ParaInfo {
    const fc = this.pieces.fcOf(markCp);
    const index = fc < 0 ? -1 : this.papx.find(fc);
    const pieceIndex = this.pieces.pieceAt(markCp);
    const piece = pieceIndex < 0 ? undefined : this.pieces.pieces[pieceIndex];
    const modifier = piece ? this.modifierOf(piece) : undefined;
    const istd = index < 0 ? 0 : this.papx.istd[index];
    const at = index < 0 ? 0 : this.papx.at[index];
    const length = index < 0 ? 0 : this.papx.length[index];
    const key = istd + '|' + (modifier ? piece!.prm : 0) + '|' + bytesKey(this.word, at, at + length);
    let info = this.paraCache.get(key);
    if (!info) {
      const lists: [Uint8Array, number, number][] = [];
      if (length) lists.push([this.word, at, at + length]);
      if (modifier) {
        lists.push([modifier, 0, modifier.length]);
        this.spend(modifier);
      }
      info = paraProps(this.ctx, istd, lists);
      const style = this.sheet.ids[info.istd];
      if (info.physicalAlign && info.props.align !== undefined) {
        // The alignment names a side of the page, and whether the paragraph runs from the right is up to its style.
        let bidi = false;
        for (const entry of styleChain(this.sheet.styles, style, 'paragraph')) if (entry.para && entry.para.bidi !== undefined) bidi = entry.para.bidi;
        if (bidi) info.props.align = mirrored(info.props.align);
      }
      // A paragraph of the default style states no style, as in a docx file.
      if (style !== undefined && style !== this.sheet.styles.defaultIds.paragraph) info.props.style = style;
      if (!Object.keys(info.props).length) info.props = EMPTY_PARA_PROPS;
      this.paraCache.set(key, info);
    }
    return info;
  }

  /**
   * The property modifier of a piece. A long one is read again for every run
   * of the piece that is formatted differently, so the bytes spent that way
   * are counted; once the count is used up, modifiers are no longer applied.
   */
  private modifierOf(piece: Piece): Uint8Array | undefined {
    const modifier = this.pieces.modifier(piece);
    if (!modifier) return undefined;
    if (this.modifierBudget <= 0) return undefined;
    return modifier;
  }

  private spend(modifier: Uint8Array): void {
    this.modifierBudget -= modifier.length;
    if (this.modifierBudget <= 0) this.warnings.push('The formatting that pieces of the text add was not applied to all of the text');
  }

  /** Properties of the characters at a place of the WordDocument stream, for text of the piece with the given modifier. */
  private charInfo(run: number, prm: number, modifier: Uint8Array | undefined): CharProps {
    const cacheKey = (run + 1) * 65536 + (modifier ? prm : 0);
    let info = this.charCache.get(cacheKey);
    if (info) return info;
    const at = run < 0 ? 0 : this.chpx.at[run];
    const length = run < 0 ? 0 : this.chpx.length[run];
    // Runs that say the same share one set of properties.
    const key = (modifier ? prm : 0) + '|' + bytesKey(this.word, at, at + length);
    info = this.charByKey.get(key);
    if (!info) {
      const lists: [Uint8Array, number, number][] = [];
      if (length) lists.push([this.word, at, at + length]);
      if (modifier) {
        lists.push([modifier, 0, modifier.length]);
        this.spend(modifier);
      }
      info = charProps(this.ctx, lists);
      if (!Object.keys(info.props).length) info.props = EMPTY_RUN_PROPS;
      this.charByKey.set(key, info);
    }
    this.charCache.set(cacheKey, info);
    return info;
  }

  /**
   * What the styles of one level say about the toggles (bold, italic, ...): the style and those it is based
   * on, a derived style over its base. Undefined when they say nothing.
   */
  private toggleLevel(type: StyleType, id: string | undefined): ToggleLevel | undefined {
    const key = type + '\u0001' + (id === undefined ? '' : id);
    let level = this.toggleLevels.get(key);
    if (level === undefined) {
      const found: ToggleLevel = new Map();
      const flips = this.sheet.flips;
      for (const style of styleChain(this.sheet.styles, id, type)) {
        const run = style.run;
        if (!run) continue;
        const flipped = flips ? flips.get(style.id) : undefined;
        for (const name of TOGGLE_RUN_PROPS) {
          const value = run[name];
          if (value !== undefined) found.set(name, flipped && flipped.indexOf(name) >= 0 ? 'flip' : value ? 'on' : 'off');
        }
      }
      level = found.size ? found : null;
      this.toggleLevels.set(key, level);
    }
    return level === null ? undefined : level;
  }

  /**
   * The run properties for the model.
   *
   * A toggle (bold, italic, ...) works differently in this format than in the model. Here a
   * style states it as on or off, and the innermost statement holds: the character style
   * over the paragraph style over the table style. Only "the opposite" (operand 0x81) looks
   * at what is below. The model follows docx, where "on" in a style flips what the levels
   * below say, and only the run's own formatting is absolute. So where the two rules give
   * different answers for the styles of a run, and for every toggle the run itself states
   * as "the opposite of the style", the value this format gives is written into the run's
   * own formatting, which holds whatever the styles say.
   *
   * `tableStyle` is given for a run in a table: the id of the style of the table, null when
   * the table names none.
   */
  private runProps(info: CharProps, para: ParaProps, tableStyle?: string | null): RunProps {
    const table = tableStyle === undefined ? undefined : this.toggleLevel('table', tableStyle === null ? undefined : tableStyle);
    const paragraph = this.toggleLevel('paragraph', para.style);
    const character = this.toggleLevel('character', info.props.style);
    let props = info.props;
    // One level of styles at most and nothing relative: both rules agree.
    if (info.invert || (table ? 1 : 0) + (paragraph ? 1 : 0) + (character ? 1 : 0) > 1) {
      const key = (para.style === undefined ? '' : para.style) + '\u0001' + (tableStyle === undefined ? '' : tableStyle === null ? '\u0002' : '\u0003' + tableStyle);
      let byStyle = this.inverted.get(info);
      if (!byStyle) this.inverted.set(info, (byStyle = new Map()));
      const known = byStyle.get(key);
      if (known) props = known;
      else {
        let changed: RunProps | undefined;
        for (const name of TOGGLE_RUN_PROPS) {
          const inverted = info.invert !== undefined && (info.invert as string[]).indexOf(name) >= 0;
          if (!inverted && info.props[name] !== undefined) continue;
          const t = table && table.get(name);
          const p = paragraph && paragraph.get(name);
          const c = character && character.get(name);
          // This format: the innermost statement holds.
          let styled = false;
          if (t) styled = t === 'on';
          if (p) styled = p === 'on';
          if (c) styled = c === 'flip' ? !styled : c === 'on';
          // The model: every "on" flips.
          const flipped = (t === 'on') !== (p === 'on') !== (c === 'on' || c === 'flip');
          if (inverted) (changed || (changed = { ...info.props }))[name] = !styled;
          else if (styled !== flipped) (changed || (changed = { ...info.props }))[name] = styled;
        }
        if (changed) props = changed;
        byStyle.set(key, props);
      }
    }
    this.runSource.set(props, info);
    return props;
  }

  /**
   * A table with a style whose text has toggles: its runs were read before the table was known, so
   * their properties are made again with the style. Tables inside its cells have had their turn.
   */
  private readonly restyleTable = (table: Table): void => {
    const style = table.props.style === undefined ? null : table.props.style;
    if (!this.toggleLevel('table', style === null ? undefined : style)) return;
    const key = style === null ? '\u0002' : '\u0003' + style;
    for (const row of table.rows) {
      for (const cell of row.cells) {
        for (const block of cell.blocks) {
          if (block.type !== 'paragraph') continue;
          for (const inline of block.children) {
            if (inline.type !== 'run') continue;
            const source = this.runSource.get(inline.props);
            if (source) inline.props = this.runProps(source, block.props, style);
          }
          const mark = block.props.markProps;
          const source = mark ? this.runSource.get(mark) : undefined;
          if (!mark || !source) continue;
          const again = mark.deleted || mark.inserted ? this.revisedMark(this.runProps(source, block.props, style), source) : this.runProps(source, block.props, style);
          if (again === mark) continue;
          // Paragraph properties are shared between paragraphs: a changed copy is shared the same way.
          let byKey = this.tableParas.get(block.props);
          if (!byKey) this.tableParas.set(block.props, (byKey = new Map()));
          let changed = byKey.get(key);
          if (!changed) byKey.set(key, (changed = { ...block.props, markProps: again }));
          block.props = changed;
        }
      }
    }
  };

  /**
   * The run properties of a paragraph mark that the character properties say was deleted or inserted: the
   * properties with that revision. Shared, like every bag of properties, between the marks that are alike.
   */
  private revisedMark(mark: RunProps, source: CharProps): RunProps {
    if (!source.deleted && !source.inserted) return mark;
    let known = this.revisedMarks.get(source);
    if (!known) this.revisedMarks.set(source, (known = new Map()));
    let out = known.get(mark);
    if (!out) {
      out = { ...mark };
      if (source.deleted) out.deleted = source.deleted;
      if (source.inserted) out.inserted = source.inserted;
      known.set(mark, out);
      this.runSource.set(out, source);
    }
    return out;
  }

  private finalPara(info: ParaInfo, mark: RunProps, section: SectionProps | undefined): ParaProps {
    let props = info.props;
    if (mark !== EMPTY_RUN_PROPS) {
      let byMark = this.paraFinal.get(info.props);
      if (!byMark) this.paraFinal.set(info.props, (byMark = new Map()));
      let withMark = byMark.get(mark);
      if (!withMark) byMark.set(mark, (withMark = { ...info.props, markProps: mark }));
      props = withMark;
    }
    return section ? { ...props, section } : props;
  }

  // -------------------------------------------------------------------- stories

  /**
   * The blocks of the text [from, to). `kind` says what the special
   * characters in it mean; `depth` counts text boxes around it.
   */
  private blocks(from: number, to: number, kind: StoryKind, depth: number): Block[] {
    const text = this.text;
    const assembler = new BlockAssembler(this.ctx, this.styleMargin, this.restyleTable);
    const main = kind === 'main';
    // The marks of this stretch of text, in order.
    const marks = this.marks;
    let mark = 0;
    if (marks.length) {
      let lo = 0;
      let hi = marks.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (marks[mid].cp < from) lo = mid + 1;
        else hi = mid;
      }
      mark = lo;
    }
    let start = from;
    let last: Paragraph | undefined;
    /** Index of the next section that ends in this text. */
    let ending = 0;
    for (let cp = from; cp < to; cp++) {
      const code = text.charCodeAt(cp);
      const ends = code === PARAGRAPH_MARK || code === CELL_MARK || (code === SECTION_MARK && main && this.sectionMarks.has(cp));
      if (!ends && cp + 1 < to) continue;
      // A story that does not end with a mark: its last characters are a paragraph all the same.
      const markCp = cp;
      const end = ends ? cp : cp + 1;
      if (++this.paragraphs > this.maxParagraphs) throw invalid(`the document has more than ${this.maxParagraphs} paragraphs`);
      const info = this.paraInfo(markCp);
      const children: Inline[] = [];
      mark = this.runs(start, end, info.props, kind, depth, children, mark, markCp);
      const markRun = this.charAt(markCp);
      // The paragraph ends the sections whose last character it holds: one, unless the file has sections without text.
      let section: SectionProps | undefined;
      if (main) while (ending < this.sectionEnds.length && this.sectionEnds[ending] <= markCp) section = this.sections[ending++];
      // A paragraph mark is a character: one that was deleted or inserted as a tracked change says so. UNSURE: the
      // mark that ends a cell is left as it is; a deleted one belongs to a deleted row, which the reader has no
      // model for.
      let markProps = this.runProps(markRun, info.props);
      if (ends && code !== CELL_MARK) markProps = this.revisedMark(markProps, markRun);
      const paragraph: Paragraph = { type: 'paragraph', props: this.finalPara(info, markProps, section), children };
      assembler.add(paragraph, info, ends ? code : 0);
      last = paragraph;
      start = cp + 1;
    }
    // Marks at the very end of the text belong to its last paragraph.
    if (last) {
      while (mark < marks.length && marks[mark].cp <= to) {
        const item = marks[mark++];
        if (item.inline && !item.used) {
          item.used = true;
          last.children.push(item.inline);
        }
      }
    }
    const blocks = assembler.finish();
    linkDocFields(blocks);
    return blocks;
  }

  private charAt(cp: number): CharProps {
    const index = this.pieces.pieceAt(cp);
    if (index < 0) return this.charInfo(-1, 0, undefined);
    const piece = this.pieces.pieces[index];
    const fc = piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
    return this.charInfo(this.chpx.find(fc), piece.prm, this.modifierOf(piece));
  }

  /**
   * Appends the runs of the characters [from, to) of a paragraph to `out`,
   * with the marks up to `markCp` (the position of the paragraph mark) in
   * their places. Returns the index of the next mark.
   */
  private runs(from: number, to: number, para: ParaProps, kind: StoryKind, depth: number, out: Inline[], mark: number, markCp: number): number {
    const marks = this.marks;
    const pieces = this.pieces.pieces;
    let cp = from;
    let pieceIndex = from < to ? this.pieces.pieceAt(from) : -1;
    /** The tracked changes the run before was in: runs of one change that follow each other share its wrapper, as the runs in one <w:ins> do. */
    let open: { insKey?: string; ins?: Wrapper; delKey?: string; del?: Wrapper } | undefined;
    const keyOf = (info: RevisionInfo): string => (info.author === undefined ? '' : 'a' + info.author) + '\u0001' + (info.date === undefined ? '' : info.date);
    while (cp < to && pieceIndex >= 0 && pieceIndex < pieces.length) {
      const piece = pieces[pieceIndex];
      const width = piece.compressed ? 1 : 2;
      const pieceEnd = Math.min(to, piece.cpEnd);
      const modifier = this.modifierOf(piece);
      while (cp < pieceEnd) {
        // The stretch of characters with the same properties, as far as it lies in the piece.
        const fc = piece.fc + (cp - piece.cpStart) * width;
        const run = this.chpx.find(fc);
        const until = run >= 0 ? this.chpx.ends[run] : this.chpx.after(fc);
        let end = until === Infinity ? pieceEnd : Math.min(pieceEnd, cp + Math.max(1, Math.ceil((until - fc) / width)));
        // A mark inside the stretch cuts it.
        while (mark < marks.length && marks[mark].cp <= cp) mark = this.putMark(marks[mark], mark, out, para, cp);
        if (mark < marks.length && marks[mark].cp < end) end = marks[mark].cp;
        const info = this.charInfo(run, piece.prm, modifier);
        const all = this.items(cp, end, info, kind, depth);
        // A field mark gets a run of its own, as in a docx file: what reads the model looks at the fields
        // around a run, not around a part of one.
        for (const content of splitAtFieldMarks(all)) {
          const item: Run = { type: 'run', props: this.runProps(info, para), content };
          if (info.deleted || info.inserted) {
            // Text that was inserted and then deleted is in both changes: the deletion inside the insertion, each
            // with its own author and date, as a docx file has it.
            let ins: Wrapper | undefined;
            const insKey = info.inserted ? keyOf(info.inserted) : undefined;
            if (info.inserted) ins = open && open.ins && open.insKey === insKey ? open.ins : { kind: 'ins', ...info.inserted };
            let del: Wrapper | undefined;
            const delKey = info.deleted ? keyOf(info.deleted) : undefined;
            if (info.deleted) {
              if (open && open.del && open.delKey === delKey && open.del.parent === ins) del = open.del;
              else {
                del = { kind: 'del', ...info.deleted };
                if (ins) del.parent = ins;
              }
              for (let i = 0; i < content.length; i++) {
                const entry = content[i];
                if (entry.type === 'text') content[i] = { type: 'deletedText', text: entry.text };
              }
            }
            open = { insKey, ins, delKey, del };
            item.wrap = del || ins;
          } else open = undefined;
          out.push(item);
        }
        cp = end;
      }
      pieceIndex++;
    }
    while (mark < marks.length && marks[mark].cp <= markCp) mark = this.putMark(marks[mark], mark, out, para, markCp);
    return mark;
  }

  private putMark(mark: Mark, index: number, out: Inline[], para: ParaProps, cp: number): number {
    if (mark.used) return index + 1;
    mark.used = true;
    if (mark.inline) out.push(mark.inline);
    else if (mark.note) out.push({ type: 'run', props: this.runProps(this.charAt(cp), para), content: [mark.note] });
    return index + 1;
  }

  /** The items of the characters [from, to), which all have the properties `info`. */
  private items(from: number, to: number, info: CharProps, kind: StoryKind, depth: number): RunItem[] {
    const text = this.text;
    const out: RunItem[] = [];
    let plain = from;
    const flush = (end: number) => {
      if (end > plain) out.push({ type: 'text', text: text.slice(plain, end) });
    };
    for (let cp = from; cp < to; cp++) {
      const code = text.charCodeAt(cp);
      if (code >= 0x20 && !(info.special && code === 0x28)) continue;
      let item: RunItem | undefined;
      switch (code) {
        case 0x09:
          item = TAB;
          break;
        case 0x0b:
          item = LINE_BREAK;
          break;
        case 0x0c:
          item = PAGE_BREAK;
          break;
        case 0x0e:
          item = COLUMN_BREAK;
          break;
        case 0x1e:
          item = NO_BREAK_HYPHEN;
          break;
        case 0x1f:
          item = SOFT_HYPHEN;
          break;
        // The field marks are read as such whether the run says "special" or not: they are never text.
        case 0x13:
          item = { type: 'fieldChar', charType: 'begin' };
          break;
        case 0x14:
          item = { type: 'fieldChar', charType: 'separate' };
          break;
        case 0x15:
          item = { type: 'fieldChar', charType: 'end' };
          break;
        default:
          if (info.special) item = this.special(code, cp, info, kind, depth);
      }
      // A control character that stands for nothing here is dropped.
      flush(cp);
      plain = cp + 1;
      if (item) out.push(item);
    }
    flush(to);
    return out;
  }

  /** What a character of a "special" run stands for. */
  private special(code: number, cp: number, info: CharProps, kind: StoryKind, depth: number): RunItem | undefined {
    switch (code) {
      case 0x00:
        return PAGE_NUMBER;
      case 0x01: {
        // A picture, unless the place in the Data stream holds field data or an embedded object.
        if (info.picture === undefined || info.data || info.ole2 || info.object) return undefined;
        const item = inlinePicture(this.data, info.picture, this.store);
        if (!item) this.count('picture/unreadable');
        return item;
      }
      case 0x02: {
        if (kind === 'footnote') return FOOTNOTE_REF;
        if (kind === 'endnote') return ENDNOTE_REF;
        const footnote = this.footnoteRefs.get(cp);
        if (footnote !== undefined) return { type: 'footnoteReference', id: footnote };
        const endnote = this.endnoteRefs.get(cp);
        if (endnote !== undefined) return { type: 'endnoteReference', id: endnote };
        return undefined;
      }
      case 0x03:
        return SEPARATOR;
      case 0x04:
        return CONTINUATION;
      case 0x05: {
        if (kind === 'comment') return ANNOTATION_REF;
        const comment = this.commentRefs.get(cp);
        return comment === undefined ? undefined : { type: 'commentReference', id: comment };
      }
      case 0x08: {
        const spa = this.shapeAt.get(cp);
        if (spa === undefined) return undefined;
        if (!this.drawings) {
          const [at, size] = this.fib.fcLcb(FC.DggInfo);
          this.drawings = readDrawings(this.table, this.word, at, size);
        }
        const header = cp >= this.starts[2] && cp < this.starts[3];
        return floatingDrawing(this.table, this.word, spa, this.drawings, this.store, (shapeId, textId) => this.boxText(header ? 1 : 0, shapeId, textId, depth));
      }
      case 0x28: {
        if (!info.symbol) return { type: 'text', text: '(' };
        const item: RunItem = { type: 'symbol', char: info.symbol.char };
        if (info.symbol.font !== undefined) item.font = info.symbol.font;
        return item;
      }
    }
    return undefined;
  }

  private count(key: string): void {
    this.unknown.set(key, (this.unknown.get(key) || 0) + 1);
  }

  /**
   * The text of the text box of a shape; undefined when the shape has none.
   *
   * In a genuine file the text of a box belongs to one shape. Here it goes to
   * the first shape that asks for it and to no other, whatever the depth: a
   * second shape that names the same text gets none, and so does a shape
   * inside the text it is the box of (which would never end). Every text is
   * therefore read once at most, and the texts do not overlap (the list of
   * them is read in order), so all boxes together hold no more than the text
   * of the file that is set aside for them.
   */
  private boxText(set: number, shapeId: number, textId: number, depth: number): Block[] | undefined {
    const boxes = this.boxes[set];
    if (!boxes || depth >= MAX_BOX_DEPTH) return undefined;
    let index = boxes.byShape.get(shapeId);
    if (index === undefined && textId > 0 && textId <= boxes.plc.count) index = textId - 1;
    if (index === undefined) return undefined;
    const key = set + ':' + index;
    if (this.boxesTaken.has(key)) return undefined;
    const limit = this.starts[set === 0 ? 7 : 8];
    const from = Math.min(limit, boxes.start + boxes.plc.cps[index]);
    // The last character of the text of a box is a paragraph mark that separates it from the next one.
    const to = Math.min(limit, boxes.start + boxes.plc.cps[index + 1]) - 1;
    if (to <= from) return undefined;
    this.boxesTaken.add(key);
    return this.blocks(from, to, 'textbox', depth + 1);
  }

  // ------------------------------------------------------------------- lookups

  private plc(index: number, size: number, ordered = false): Plc {
    const [at, length] = this.fib.fcLcb(index);
    return readPlc(this.table, at, length, size, ordered);
  }

  /** Reads the structures that say what sits where in the text. */
  private prepare(): void {
    const table = this.table;
    const starts = this.starts;

    // Sections: all but the last end with a section mark.
    const sections = this.plc(FC.PlcfSed, 12, true);
    const props: SectionProps[] = [];
    const parsed = new Map<number, SectionProps>();
    for (let i = 0; i < sections.count; i++) {
      let at = u32(table, sections.data(i) + 2);
      if (at === 0xffffffff || at + 2 > this.word.length) at = -1;
      // Sections that point to the same properties are read once; each gets an object of its own, to which its
      // headers are added later.
      let read = parsed.get(at);
      if (!read) {
        if (at < 0) read = sectionProps(this.ctx, this.word, 0, 0);
        else read = sectionProps(this.ctx, this.word, at + 2, Math.min(this.word.length, at + 2 + u16(this.word, at)));
        parsed.set(at, read);
      }
      const section: SectionProps = { ...read };
      props.push(section);
      const end = sections.cps[i + 1];
      // The last section ends with the text; its properties are those of the body.
      if (i < sections.count - 1) {
        const mark = Math.max(0, Math.min(end, starts[1]) - 1);
        this.sectionEnds.push(mark);
        this.sectionMarks.add(mark);
      }
    }
    this.sections = props.length ? props : [sectionProps(this.ctx, this.word, 0, 0)];

    // Notes and comments: where the references are in the main text.
    const notes = (index: number, size: number, target: Map<number, number>, first: number) => {
      const plc = this.plc(index, size);
      for (let i = 0; i < plc.count; i++) if (!target.has(plc.cps[i])) target.set(plc.cps[i], i + first);
      return plc;
    };
    const marks: Mark[] = [];
    const footnotes = notes(FC.PlcffndRef, 2, this.footnoteRefs, 1);
    const endnotes = notes(FC.PlcfendRef, 2, this.endnoteRefs, 1);
    notes(FC.PlcfandRef, 30, this.commentRefs, 0);
    // A note whose mark is text of its own instead of the automatic number: the reference goes in front of that text.
    const custom = (plc: Plc, type: 'footnoteReference' | 'endnoteReference') => {
      for (let i = 0; i < plc.count; i++) {
        const cp = plc.cps[i];
        if (cp < starts[1] && this.text.charCodeAt(cp) !== 0x02) marks.push({ cp, order: 2, note: { type, id: i + 1, customMark: true } });
      }
    };
    custom(footnotes, 'footnoteReference');
    custom(endnotes, 'endnoteReference');

    // Bookmarks.
    const [namesAt, namesSize] = this.fib.fcLcb(FC.SttbfBkmk);
    const names = readSttb(table, namesAt, namesSize).strings;
    const first = this.plc(FC.PlcfBkf, 4);
    const [limAt, limSize] = this.fib.fcLcb(FC.PlcfBkl);
    // The ends have no data in files of Word 97 to 2003.
    const lim = readPlc(table, limAt, limSize, 0);
    const total = this.text.length;
    for (let i = 0; i < first.count && i < names.length; i++) {
      const start = first.cps[i];
      const endIndex = u16(table, first.data(i));
      if (start > total || endIndex >= lim.cps.length) continue;
      const end = lim.cps[endIndex];
      if (end < start || end > total) continue;
      const id = String(i);
      marks.push({ cp: start, order: 1, inline: { type: 'bookmarkStart', id, name: names[i] } });
      marks.push({ cp: end, order: 0, inline: { type: 'bookmarkEnd', id } });
    }
    marks.sort((a, b) => a.cp - b.cp || a.order - b.order);
    this.marks = marks;

    // Drawing objects: of the main text, and of the headers (positions counted from the start of the header text).
    const shapes = (index: number, base: number) => {
      const plc = this.plc(index, 26);
      for (let i = 0; i < plc.count; i++) this.shapeAt.set(base + plc.cps[i], plc.data(i));
    };
    shapes(FC.PlcSpaMom, 0);
    shapes(FC.PlcSpaHdr, starts[2]);

    // Text boxes: each entry names the shape it is the text of.
    const boxes = (index: number, start: number) => {
      const plc = this.plc(index, 22, true);
      const byShape = new Map<number, number>();
      for (let i = 0; i < plc.count; i++) {
        const shape = u32(table, plc.data(i) + 14);
        if (shape !== 0 && !byShape.has(shape)) byShape.set(shape, i);
      }
      this.boxes.push({ plc, start, byShape });
    };
    boxes(FC.PlcftxbxTxt, starts[6]);
    boxes(FC.PlcfHdrtxbxTxt, starts[7]);
  }

  // ---------------------------------------------------------------------- read

  read(): DocModel {
    const model = this.model;
    const starts = this.starts;
    this.prepare();

    model.body.blocks = this.blocks(0, starts[1], 'main', 0);
    model.body.section = this.sections[this.sections.length - 1];

    this.notes('footnote', FC.PlcffndRef, FC.PlcffndTxt, starts[1], starts[2]);
    this.notes('endnote', FC.PlcfendRef, FC.PlcfendTxt, starts[5], starts[6]);
    this.comments(starts[4], starts[5]);
    this.headers(starts[2], starts[3]);

    if (this.store.media.size) model.media = this.store.media;
    return model;
  }

  private notes(kind: 'footnote' | 'endnote', refIndex: number, textIndex: number, from: number, to: number): void {
    const refs = this.plc(refIndex, 2);
    const texts = this.plc(textIndex, 0, true);
    if (!refs.count || to <= from) return;
    const notes = collection<Note>(kind === 'footnote' ? 'doc:footnotes' : 'doc:endnotes');
    for (let i = 0; i < refs.count && i < texts.count; i++) {
      const start = Math.min(to, from + texts.cps[i]);
      const end = Math.min(to, from + texts.cps[i + 1]);
      notes.items.set(i + 1, { id: i + 1, blocks: end > start ? this.blocks(start, end, kind, 0) : [] });
    }
    if (kind === 'footnote') this.model.footnotes = notes;
    else this.model.endnotes = notes;
  }

  private comments(from: number, to: number): void {
    const refs = this.plc(FC.PlcfandRef, 30);
    const texts = this.plc(FC.PlcfandTxt, 0, true);
    if (!refs.count || to <= from) return;
    // The authors: strings one after the other, each with its length in front.
    const authors: string[] = [];
    const [ownersAt, ownersSize] = this.fib.fcLcb(FC.GrpXstAtnOwners);
    if (ownersAt + ownersSize <= this.table.length) {
      let p = ownersAt;
      const end = ownersAt + ownersSize;
      while (p + 2 <= end && authors.length < 10000) {
        const length = u16(this.table, p);
        if (p + 2 + length * 2 > end) break;
        authors.push(utf16(this.table, p + 2, length));
        p += 2 + length * 2;
      }
    }
    const comments = collection<Comment>('doc:comments');
    for (let i = 0; i < refs.count && i < texts.count; i++) {
      const start = Math.min(to, from + texts.cps[i]);
      const end = Math.min(to, from + texts.cps[i + 1]);
      const comment: Comment = { id: i, blocks: end > start ? this.blocks(start, end, 'comment', 0) : [] };
      const at = refs.data(i);
      const initials = utf16(this.table, at + 2, Math.min(9, u16(this.table, at)));
      if (initials) comment.initials = initials;
      const author = authors[u16(this.table, at + 20)];
      if (author !== undefined) comment.author = author;
      comments.items.set(i, comment);
    }
    this.model.comments = comments;
  }

  /**
   * Headers and footers. The text of the header part is a list of stories:
   * six for the separators of the notes, then six for each section (header
   * and footer of even pages, of odd pages, of the first page). An empty
   * story means "as in the section before", which is what a section of the
   * model says by not naming a header.
   */
  private headers(from: number, to: number): void {
    const plc = this.plc(FC.PlcfHdd, 0, true);
    if (!plc.count || to <= from) return;
    const model = this.model;
    const read = (index: number, kind: StoryKind): Block[] | undefined => {
      if (index + 1 >= plc.cps.length) return undefined;
      const start = Math.min(to, from + plc.cps[index]);
      // The last character of a story is a paragraph mark that separates it from the next one.
      const end = Math.min(to, from + plc.cps[index + 1]) - 1;
      return end > start ? this.blocks(start, end, kind, 0) : undefined;
    };

    // The separators of the notes are notes of a kind of their own, as in a docx file.
    const separators: [index: number, list: 'footnotes' | 'endnotes', id: number, kind: string][] = [
      [0, 'footnotes', -1, 'separator'],
      [1, 'footnotes', 0, 'continuationSeparator'],
      [2, 'footnotes', -2, 'continuationNotice'],
      [3, 'endnotes', -1, 'separator'],
      [4, 'endnotes', 0, 'continuationSeparator'],
      [5, 'endnotes', -2, 'continuationNotice'],
    ];
    for (const [index, list, id, kind] of separators) {
      const notes = model[list];
      if (!notes) continue;
      const blocks = read(index, list === 'footnotes' ? 'footnote' : 'endnote');
      if (blocks) notes.items.set(id, { id, kind, blocks });
    }

    for (let s = 0; s < this.sections.length; s++) {
      for (let k = 0; k < 6; k++) {
        const blocks = read(6 + s * 6 + k, 'header');
        if (!blocks) continue;
        const { list, kind } = HEADER_ORDER[k];
        const id = (list === 'headers' ? 'hdr:' : 'ftr:') + s + ':' + kind;
        model[list].set(id, story(list === 'headers' ? 'doc:header' : 'doc:footer', blocks));
        // The properties of a section are shared with the paragraph that ends it; they are filled in here once.
        const section = this.sections[s];
        (section[list] || (section[list] = {}))[kind] = id;
      }
    }
  }
}

/**
 * Reads a Word binary document from its compound file. Throws
 * UnsupportedDocError (named "UnsupportedFileError") for documents of Word 95
 * and earlier and for encrypted documents, and an Error that starts with
 * "Invalid doc:" for a file whose text cannot be read or that has more
 * paragraphs than the reader takes.
 */
export function readDoc(cfb: CfbFile, options: ReadDocOptions = {}): DocModel {
  return new DocReader(cfb, options).read();
}
