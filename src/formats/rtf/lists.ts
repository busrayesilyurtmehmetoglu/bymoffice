// Lists of an RTF file ("\listtable" and "\listoverridetable") -> the
// numbering of the document model.
//
// A "\list" with its one or nine "\listlevel" groups is what a docx file
// calls an abstract numbering; a "\listoverride", which a paragraph refers to
// by its "\ls" number, is a numbering instance. The id of an abstract
// numbering here is the index of its list in the table, the id of an
// instance the "\ls" number, as in the model read from a binary file.
//
// The text of a level ("\leveltext") is a counted string: its first
// character is the length, and a character below 9 stands for the counter of
// the level with that index.

import type { AbstractNum, NumInstance, NumLevel, Numbering, ParaProps, RunProps } from '../../model/doc/types';
import { numberFormat } from '../doc/sprm';
import { SYMBOL_CODEPAGE } from './codepage';

/** More lists than this in one document is not a genuine file. */
export const MAX_LISTS = 20000;
/** Characters the text of a level may have. */
export const MAX_LEVEL_TEXT = 64;

const ALIGNMENTS = ['left', 'center', 'right'];
const SUFFIXES: NonNullable<NumLevel['suffix']>[] = ['tab', 'space', 'nothing'];

/** A level while its group is read. */
export interface LevelState {
  format: number;
  align: number;
  follow: number;
  start: number;
  legal: boolean;
  noRestart: boolean;
  /**
   * The characters of the level text: a character code, or for a byte given
   * as "\'hh" the byte less 256 (so negative), which is read in the code
   * page of the font of the level once that is known.
   */
  text: number[];
  run: RunProps | undefined;
  para: ParaProps | undefined;
  /** Code page of the font of the level; text bytes are characters of it. */
  codepage: number;
}

export function newLevel(): LevelState {
  return { format: 0, align: 0, follow: 0, start: 1, legal: false, noRestart: false, text: [], run: undefined, para: undefined, codepage: 1252 };
}

/** A control word of a level; false when it is not one. */
export function levelWord(level: LevelState, word: string, param: number): boolean {
  switch (word) {
    case 'levelnfc':
    case 'levelnfcn':
      level.format = param;
      return true;
    case 'leveljc':
    case 'leveljcn':
      level.align = param;
      return true;
    case 'levelfollow':
      level.follow = param;
      return true;
    case 'levelstartat':
      level.start = param;
      return true;
    case 'levellegal':
      level.legal = param !== 0;
      return true;
    case 'levelnorestart':
      level.noRestart = param !== 0;
      return true;
    case 'levelspace':
    case 'levelindent':
    case 'levelold':
    case 'levelprev':
    case 'levelprevspace':
    case 'levelpicture':
    case 'levelpicturenosize':
    case 'lvltentative':
      return true;
  }
  return false;
}

export interface ListState {
  /** "\listid": what an override names its list by. */
  listId: number;
  simple: boolean;
  hybrid: boolean;
  levels: LevelState[];
}

export interface OverrideState {
  listId: number;
  /** "\ls": what a paragraph names the override by. */
  number: number;
  levels: { start: number | undefined; level: LevelState | undefined; startOverride: boolean; formatOverride: boolean }[];
}

function buildLevel(state: LevelState, index: number, decode: (codepage: number, byte: number) => string): NumLevel {
  const format = numberFormat(state.format);
  const level: NumLevel = { level: index, start: Math.max(0, state.start), format };
  const codes = state.text;
  // The first character counts the ones after it; a text without a usable count is taken up to its last semicolon.
  let from = 0;
  let to = codes.length;
  const first = codes.length ? (codes[0] < 0 ? codes[0] + 256 : codes[0]) : -1;
  if (first >= 0 && first <= MAX_LEVEL_TEXT && first < codes.length) {
    from = 1;
    to = 1 + first;
  } else {
    for (let i = codes.length - 1; i >= 0; i--) {
      if (codes[i] === 0x3b) {
        to = i;
        break;
      }
    }
  }
  let text = '';
  for (let i = from; i < to; i++) {
    const raw = codes[i] < 0;
    const code = raw ? codes[i] + 256 : codes[i];
    if (code < 9 && format !== 'bullet') text += '%' + (code + 1);
    else if (raw) text += decode(state.codepage, code);
    // A character of a font of symbols is a glyph of the font however it is written (see RtfReader.symbolText).
    else text += String.fromCharCode(state.codepage === SYMBOL_CODEPAGE && code > 0x20 && code < 0x100 ? 0xf000 + code : code);
  }
  level.text = text;
  const suffix = SUFFIXES[state.follow];
  if (suffix !== undefined) level.suffix = suffix;
  const align = ALIGNMENTS[state.align];
  if (align !== undefined) level.align = align;
  if (state.legal) level.legal = true;
  // The level only restarts when the list itself does, not after any higher level.
  if (state.noRestart) level.restart = 0;
  if (state.para) level.para = state.para;
  if (state.run) level.run = state.run;
  return level;
}

/** The numbering of the model from the lists and overrides of the file; overrides of a list the file lacks are left out. */
export function buildNumbering(lists: ListState[], overrides: OverrideState[], decode: (codepage: number, byte: number) => string): Numbering {
  const numbering: Numbering = { abstracts: new Map(), nums: new Map() };
  /** Abstract id by list id. */
  const byListId = new Map<number, number>();
  lists.forEach((list, i) => {
    const abstract: AbstractNum = { id: i, levels: new Map() };
    abstract.multiLevelType = list.simple ? 'singleLevel' : list.hybrid ? 'hybridMultilevel' : 'multilevel';
    const count = Math.min(list.levels.length, 9);
    for (let k = 0; k < count; k++) abstract.levels.set(k, buildLevel(list.levels[k], k, decode));
    if (!byListId.has(list.listId)) byListId.set(list.listId, i);
    numbering.abstracts.set(i, abstract);
  });
  for (const override of overrides) {
    const abstractId = byListId.get(override.listId);
    if (abstractId === undefined || override.number < 1 || numbering.nums.has(override.number)) continue;
    const num: NumInstance = { id: override.number, abstractId, overrides: new Map() };
    const count = Math.min(override.levels.length, 9);
    for (let k = 0; k < count; k++) {
      const entry = override.levels[k];
      const out: { start?: number; level?: NumLevel } = {};
      if (entry.formatOverride && entry.level) out.level = buildLevel(entry.level, k, decode);
      if (entry.startOverride) {
        const start = entry.start !== undefined ? entry.start : entry.level ? entry.level.start : undefined;
        if (start !== undefined) out.start = Math.max(0, start);
      }
      if (out.level || out.start !== undefined) num.overrides.set(k, out);
    }
    numbering.nums.set(num.id, num);
  }
  return numbering;
}
