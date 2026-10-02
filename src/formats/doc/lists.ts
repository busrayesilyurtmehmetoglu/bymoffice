// Lists (PlfLst and PlfLfo, MS-DOC 2.9.191 and 2.9.190) -> the numbering of
// the document model.
//
// A list definition (LSTF) with its one or nine levels (LVL) is what a docx
// file calls an abstract numbering; a list format override (LFO), which a
// paragraph refers to by its 1-based index, is a numbering instance. The id
// of an abstract numbering here is the index of its definition, the id of an
// instance the index paragraphs use.

import type { AbstractNum, NumInstance, NumLevel, Numbering } from '../../model/doc/types';
import { i32, u16, u32, u8, utf16 } from './bytes';
import { charProps, paraProps, type PropContext } from './props';
import { numberFormat } from './sprm';

/** More definitions than this in one document is not a genuine file. */
const MAX_LISTS = 20000;
const LSTF_SIZE = 28;
const LVLF_SIZE = 28;
const LFO_SIZE = 16;
const ALIGNMENTS = ['left', 'center', 'right'];
const SUFFIXES: NonNullable<NumLevel['suffix']>[] = ['tab', 'space', 'nothing'];
const ISTD_NIL = 0x0fff;

/** Reads the level at `p`; returns the level and where the next structure starts, undefined when it is cut short. */
function readLevel(ctx: PropContext, b: Uint8Array, p: number, end: number, index: number): { level: NumLevel; next: number } | undefined {
  if (p + LVLF_SIZE > end) return undefined;
  const flags = u8(b, p + 5);
  const chpxSize = u8(b, p + 24);
  const papxSize = u8(b, p + 25);
  const papx = p + LVLF_SIZE;
  const chpx = papx + papxSize;
  const xst = chpx + chpxSize;
  if (xst + 2 > end) return undefined;
  const length = u16(b, xst);
  const next = xst + 2 + length * 2;
  if (next > end) return undefined;

  const format = numberFormat(u8(b, p + 4));
  const level: NumLevel = { level: index, start: Math.max(0, i32(b, p)), format };
  // The text of the number: a character below 9 stands for the counter of the level with that index.
  const raw = utf16(b, xst + 2, length);
  let text = '';
  for (let i = 0; i < raw.length; i++) {
    const code = raw.charCodeAt(i);
    text += code < 9 && format !== 'bullet' ? '%' + (code + 1) : raw.charAt(i);
  }
  level.text = text;
  const suffix = SUFFIXES[u8(b, p + 15)];
  if (suffix !== undefined) level.suffix = suffix;
  const align = ALIGNMENTS[flags & 3];
  if (align !== undefined) level.align = align;
  if (flags & 0x04) level.legal = true;
  // fNoRestart: the level only restarts after the level named, not after any higher one.
  if (flags & 0x08) level.restart = Math.min(9, u8(b, p + 26));
  if (papxSize) {
    const info = paraProps(ctx, 0, [[b, papx, chpx]], true);
    if (Object.keys(info.props).length) level.para = info.props;
  }
  if (chpxSize) {
    const info = charProps(ctx, [[b, chpx, xst]]);
    if (Object.keys(info.props).length) level.run = info.props;
  }
  return { level, next };
}

/** Reads the list definitions and overrides; structures that are cut short or point nowhere are left out. */
export function readNumbering(ctx: PropContext, table: Uint8Array, lst: [number, number], lfo: [number, number]): Numbering {
  const numbering: Numbering = { abstracts: new Map(), nums: new Map() };
  const [lstAt, lstSize] = lst;
  if (lstSize < 2 || lstAt + lstSize > table.length) return numbering;
  const count = Math.min(u16(table, lstAt), MAX_LISTS, Math.floor((table.length - lstAt - 2) / LSTF_SIZE));
  /** Abstract id by list id (lsid). */
  const byListId = new Map<number, number>();
  const simple: boolean[] = [];
  for (let i = 0; i < count; i++) {
    const p = lstAt + 2 + i * LSTF_SIZE;
    const flags = u8(table, p + 26);
    const abstract: AbstractNum = { id: i, levels: new Map() };
    simple.push((flags & 1) !== 0);
    abstract.multiLevelType = flags & 1 ? 'singleLevel' : flags & 0x10 ? 'hybridMultilevel' : 'multilevel';
    const listId = i32(table, p);
    if (!byListId.has(listId)) byListId.set(listId, i);
    numbering.abstracts.set(i, abstract);
    // Styles tied to the levels are read with the levels below.
  }
  // The levels follow the array of definitions, one or nine for each definition, in order. Some writers count them
  // into the size of the PlfLst and some do not, so the size says nothing about where they are.
  let p = lstAt + 2 + count * LSTF_SIZE;
  const end = table.length;
  let complete = true;
  for (let i = 0; i < count && complete; i++) {
    const abstract = numbering.abstracts.get(i)!;
    const levels = simple[i] ? 1 : 9;
    for (let k = 0; k < levels; k++) {
      const read = readLevel(ctx, table, p, end, k);
      if (!read) {
        complete = false;
        break;
      }
      const styleIndex = u16(table, lstAt + 2 + i * LSTF_SIZE + 8 + k * 2);
      const style = styleIndex === ISTD_NIL ? undefined : ctx.styleId(styleIndex);
      if (style !== undefined) read.level.paraStyle = style;
      abstract.levels.set(k, read.level);
      p = read.next;
    }
  }
  if (!complete) ctx.warnings.push('The list definitions are cut short');

  const [lfoAt, lfoSize] = lfo;
  if (lfoSize < 4 || lfoAt + lfoSize > table.length) return numbering;
  const instances = Math.min(u32(table, lfoAt), MAX_LISTS, Math.floor((lfoSize - 4) / LFO_SIZE));
  const overrides: number[] = [];
  for (let i = 0; i < instances; i++) {
    const at = lfoAt + 4 + i * LFO_SIZE;
    const abstractId = byListId.get(i32(table, at));
    overrides.push(u8(table, at + 12));
    if (abstractId === undefined) continue;
    const num: NumInstance = { id: i + 1, abstractId, overrides: new Map() };
    numbering.nums.set(num.id, num);
  }
  // The data of the overrides follows the array: for each, four bytes and then its levels.
  let q = lfoAt + 4 + instances * LFO_SIZE;
  for (let i = 0; i < instances; i++) {
    q += 4;
    const num = numbering.nums.get(i + 1);
    const levels = Math.min(overrides[i], 9);
    for (let k = 0; k < levels; k++) {
      if (q + 8 > end) return numbering;
      const start = i32(table, q);
      const flags = u32(table, q + 4);
      q += 8;
      const index = flags & 15;
      const override: { start?: number; level?: NumLevel } = {};
      if (flags & 0x20) {
        // fFormatting: a whole level replaces the one of the definition.
        const read = readLevel(ctx, table, q, end, Math.min(index, 8));
        if (!read) return numbering;
        q = read.next;
        override.level = read.level;
      }
      if (flags & 0x10) override.start = Math.max(0, start);
      if (num && index <= 8 && !num.overrides.has(index)) num.overrides.set(index, override);
    }
  }
  return numbering;
}
