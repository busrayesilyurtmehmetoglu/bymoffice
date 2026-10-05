// Reader for word/numbering.xml, number text formatting and the counters
// that run through a document.
//
// A paragraph refers to a numbering instance (<w:num>), which refers to an
// abstract definition (<w:abstractNum>) with up to nine levels and may
// override levels or their start values. An abstract definition may instead
// point to a numbering style, whose own numbering holds the levels.

import { setPartNamespaces, type AbstractNum, type DocStyles, type ExtraXml, type NumInstance, type NumLevel, type Numbering } from '../../model/doc/types';
import type { XmlElement } from '../../xml';
import { attr, childOfKind, intAttr, kindOf, onOff, pushExtra, type PartContext } from './common';
import { readParaProps, readRunProps } from './props';

/** More definitions than this in one document is not a genuine file. */
export const MAX_NUMBERING_RECORDS = 100000;
const MAX_LEVEL = 8;
/** Longest chain of numbering style links that is followed. */
const MAX_LINKS = 8;

export function emptyNumbering(): Numbering {
  return { abstracts: new Map(), nums: new Map() };
}

function readLevel(ctx: PartContext, el: XmlElement): NumLevel | undefined {
  ctx.hits = 0;
  const index = intAttr(ctx, el, 'ilvl', 0, MAX_LEVEL);
  if (index === undefined) return undefined;
  const level: NumLevel = { level: index };
  let extras: ExtraXml[] | undefined;
  // tplc, tentative: not interpreted.
  if (ctx.hits !== el.attrs.size) extras = pushExtra(ctx, extras, -1, el, 'lvl', true);
  let position = -1;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    position++;
    const kind = kindOf(ctx, node.name);
    ctx.hits = 0;
    let understood = true;
    switch (kind) {
      case 'start': {
        const value = intAttr(ctx, node, 'val', 0, 0x7fffffff);
        if (value === undefined) understood = false;
        else level.start = value;
        break;
      }
      case 'numFmt': {
        const value = attr(ctx, node, 'val');
        if (value === undefined) understood = false;
        else level.format = value;
        break;
      }
      case 'lvlText': {
        const value = attr(ctx, node, 'val');
        if (value === undefined) understood = false;
        else level.text = value;
        break;
      }
      case 'suff': {
        const value = attr(ctx, node, 'val');
        if (value === 'tab' || value === 'space' || value === 'nothing') level.suffix = value;
        else understood = false;
        break;
      }
      case 'lvlJc': {
        const value = attr(ctx, node, 'val');
        if (value === undefined) understood = false;
        else level.align = value;
        break;
      }
      case 'lvlRestart': {
        const value = intAttr(ctx, node, 'val', 0, MAX_LEVEL + 1);
        if (value === undefined) understood = false;
        else level.restart = value;
        break;
      }
      case 'isLgl': {
        const value = onOff(ctx, node);
        if (value === undefined) understood = false;
        else level.legal = value;
        break;
      }
      case 'pStyle': {
        const value = attr(ctx, node, 'val');
        if (value === undefined) understood = false;
        else level.paraStyle = value;
        break;
      }
      case 'pPr':
        level.para = readParaProps(ctx, node);
        continue;
      case 'rPr':
        level.run = readRunProps(ctx, node);
        continue;
      default:
        understood = false;
    }
    if (!understood) extras = pushExtra(ctx, extras, position, node, 'lvl', false);
    else if (ctx.hits !== node.attrs.size) extras = pushExtra(ctx, extras, position, node, 'lvl', true);
  }
  if (extras) level.extras = extras;
  return level;
}

/** Reads the root element of the numbering part; undefined gives empty tables. */
export function parseNumbering(ctx: PartContext, root: XmlElement | undefined): Numbering {
  const numbering = emptyNumbering();
  if (!root) return numbering;
  setPartNamespaces(numbering, Array.from(root.attrs).filter(([name]) => name === 'xmlns' || name.indexOf('xmlns:') === 0));
  let count = 0;
  for (const node of root.children) {
    if (node.type !== 'element') continue;
    const kind = kindOf(ctx, node.name);
    if (kind !== 'abstractNum' && kind !== 'num') continue;
    if (++count > MAX_NUMBERING_RECORDS) throw new Error(`Invalid docx: more than ${MAX_NUMBERING_RECORDS} numbering definitions`);
    if (kind === 'abstractNum') {
      const id = intAttr(ctx, node, 'abstractNumId', 0, 0x7fffffff);
      if (id === undefined || numbering.abstracts.has(id)) continue;
      const abstract: AbstractNum = { id, levels: new Map() };
      for (const inner of node.children) {
        if (inner.type !== 'element') continue;
        const innerKind = kindOf(ctx, inner.name);
        if (innerKind === 'lvl') {
          const level = readLevel(ctx, inner);
          if (level && !abstract.levels.has(level.level)) abstract.levels.set(level.level, level);
        } else if (innerKind === 'multiLevelType' || innerKind === 'numStyleLink' || innerKind === 'styleLink') {
          const value = attr(ctx, inner, 'val');
          if (value !== undefined) abstract[innerKind] = value;
        }
      }
      numbering.abstracts.set(id, abstract);
    } else {
      const id = intAttr(ctx, node, 'numId', 0, 0x7fffffff);
      const reference = childOfKind(ctx, node, 'abstractNumId');
      const abstractId = reference ? intAttr(ctx, reference, 'val', 0, 0x7fffffff) : undefined;
      if (id === undefined || abstractId === undefined || numbering.nums.has(id)) continue;
      const num: NumInstance = { id, abstractId, overrides: new Map() };
      for (const inner of node.children) {
        if (inner.type !== 'element' || kindOf(ctx, inner.name) !== 'lvlOverride') continue;
        const index = intAttr(ctx, inner, 'ilvl', 0, MAX_LEVEL);
        if (index === undefined || num.overrides.has(index)) continue;
        const override: { start?: number; level?: NumLevel } = {};
        const startElement = childOfKind(ctx, inner, 'startOverride');
        const start = startElement ? intAttr(ctx, startElement, 'val', 0, 0x7fffffff) : undefined;
        if (start !== undefined) override.start = start;
        const levelElement = childOfKind(ctx, inner, 'lvl');
        const level = levelElement ? readLevel(ctx, levelElement) : undefined;
        if (level) override.level = level;
        num.overrides.set(index, override);
      }
      numbering.nums.set(id, num);
    }
  }
  return numbering;
}

/**
 * The abstract definition a numbering instance uses, following numbering
 * style links (a definition that only names a style takes its levels from the
 * numbering that style refers to). Links that loop or run too long end at the
 * definition reached so far.
 */
export function abstractOf(numbering: Numbering, styles: DocStyles | undefined, numId: number): AbstractNum | undefined {
  let num = numbering.nums.get(numId);
  let abstract = num && numbering.abstracts.get(num.abstractId);
  for (let hops = 0; abstract && abstract.numStyleLink !== undefined && styles && hops < MAX_LINKS; hops++) {
    const style = styles.styles.get(abstract.numStyleLink);
    const linkedId = style && style.para && style.para.numbering ? style.para.numbering.numId : undefined;
    num = linkedId === undefined ? undefined : numbering.nums.get(linkedId);
    const next = num && numbering.abstracts.get(num.abstractId);
    if (!next || next === abstract) break;
    abstract = next;
  }
  return abstract;
}

/** The level definition in effect for a numbering instance: its own override, else the abstract one. */
export function levelOf(numbering: Numbering, styles: DocStyles | undefined, numId: number, level: number): NumLevel | undefined {
  const num = numbering.nums.get(numId);
  if (!num) return undefined;
  const override = num.overrides.get(level);
  if (override && override.level) return override.level;
  const abstract = abstractOf(numbering, styles, numId);
  return abstract && abstract.levels.get(level);
}

const ROMAN: [value: number, text: string][] = [
  [1000, 'm'],
  [900, 'cm'],
  [500, 'd'],
  [400, 'cd'],
  [100, 'c'],
  [90, 'xc'],
  [50, 'l'],
  [40, 'xl'],
  [10, 'x'],
  [9, 'ix'],
  [5, 'v'],
  [4, 'iv'],
  [1, 'i'],
];

function roman(value: number): string {
  // Beyond what Roman numerals are used for, the plain number is shown.
  if (value < 1 || value > 32767) return String(value);
  let out = '';
  for (const [step, text] of ROMAN) {
    while (value >= step) {
      out += text;
      value -= step;
    }
  }
  return out;
}

/**
 * a..z, then aa, bb, ..., zz, aaa, ...: the letter repeats, as Word does it.
 *
 * The alphabet is the English one whatever the language of the text: Word
 * numbers a Turkish list a, b, c, d (no ç), which the numbers Word writes
 * into RTF files next to its lists show. What does follow the language is how
 * the letters are made upper case: see upperLetters().
 */
function letters(value: number): string {
  if (value < 1 || value > 780) return String(value);
  const letter = String.fromCharCode(97 + ((value - 1) % 26));
  let out = '';
  for (let i = Math.floor((value - 1) / 26); i >= 0; i--) out += letter;
  return out;
}

/** Is the language one whose capital of "i" has a dot: Turkish, Azerbaijani. */
function dottedCapital(lang: string | undefined): boolean {
  if (lang === undefined || lang.length < 2) return false;
  const code = lang.slice(0, 2).toLowerCase();
  if (code !== 'tr' && code !== 'az') return false;
  // "tr", "tr-TR", "az-Latn-AZ"; not a longer code that only starts the same.
  return lang.length === 2 || lang.charCodeAt(2) === 0x2d || lang.charCodeAt(2) === 0x5f;
}

/**
 * letters() in upper case, by the rules of the language of the text: in
 * Turkish the ninth item of a list is "İ.", which is what Word shows.
 *
 * UNSURE: upper case Roman numerals are left alone ("I", "IX" in every
 * language); there is no evidence that Word writes them with a dot.
 */
function upperLetters(value: number, lang: string | undefined): string {
  const text = letters(value);
  return (dottedCapital(lang) ? text.replace(/i/g, '\u0130') : text).toUpperCase();
}

function ordinal(value: number): string {
  const tens = value % 100;
  const suffix = tens >= 11 && tens <= 13 ? 'th' : value % 10 === 1 ? 'st' : value % 10 === 2 ? 'nd' : value % 10 === 3 ? 'rd' : 'th';
  return value + suffix;
}

/**
 * Text of a counter value in a number format. "bullet" and "none" have no
 * counter text; formats without an implementation fall back to decimal.
 * `lang` is the language of the numbered text ("tr-TR"): upper case letters
 * follow it (see upperLetters()).
 */
export function formatCounter(value: number, format: string | undefined, lang?: string): string {
  switch (format) {
    case 'bullet':
    case 'none':
      return '';
    case 'decimalZero':
      return value >= 0 && value < 10 ? '0' + value : String(value);
    case 'lowerLetter':
      return letters(value);
    case 'upperLetter':
      return upperLetters(value, lang);
    case 'lowerRoman':
      return roman(value);
    case 'upperRoman':
      return roman(value).toUpperCase();
    case 'ordinal':
      return ordinal(value);
  }
  return String(value);
}

export interface NumberLabel {
  /** Text of the number or bullet, e.g. "2.1." or "•". */
  text: string;
  /** The level definition used (format, suffix, alignment, run and paragraph properties). */
  level: NumLevel;
  /** Counter values of levels 0..level. */
  values: number[];
}

/**
 * Counter state for one pass over a document. Call next() for every numbered
 * paragraph in document order; it advances the counters and returns the label.
 *
 * Counters belong to the abstract definition, so instances that share one
 * continue each other's numbering. An instance with a start override restarts
 * the level the first time it is used.
 */
export class NumberingCounters {
  private counters = new Map<number, (number | undefined)[]>();
  private overridden = new Set<string>();

  constructor(
    private readonly numbering: Numbering,
    private readonly styles?: DocStyles,
  ) {}

  private startOf(numId: number, level: number): number {
    const override = this.numbering.nums.get(numId)?.overrides.get(level);
    if (override && override.start !== undefined) return override.start;
    const definition = levelOf(this.numbering, this.styles, numId, level);
    // The schema says an absent start is 0; writers always state it, and lists start at 1.
    return definition && definition.start !== undefined ? definition.start : 1;
  }

  /** `lang` is the language of the numbered text, which some number formats follow (see formatCounter). */
  next(numId: number, level: number, lang?: string): NumberLabel | undefined {
    if (!(level >= 0 && level <= MAX_LEVEL)) level = 0;
    const abstract = abstractOf(this.numbering, this.styles, numId);
    const definition = levelOf(this.numbering, this.styles, numId, level);
    if (!abstract || !definition) return undefined;

    let counters = this.counters.get(abstract.id);
    if (!counters) this.counters.set(abstract.id, (counters = []));

    const override = this.numbering.nums.get(numId)?.overrides.get(level);
    const key = `${numId}:${level}`;
    const current = counters[level];
    if (override && override.start !== undefined && !this.overridden.has(key)) {
      this.overridden.add(key);
      counters[level] = override.start;
    } else counters[level] = current === undefined ? this.startOf(numId, level) : current + 1;

    // Using a level restarts the deeper ones, unless they say otherwise: a
    // deeper level may name the (1-based) level whose use restarts it, and 0
    // there means "never", which the comparison below gives without a case of its own.
    for (let deeper = level + 1; deeper <= MAX_LEVEL; deeper++) {
      if (counters[deeper] === undefined) continue;
      const restart = levelOf(this.numbering, this.styles, numId, deeper)?.restart;
      if (restart === undefined || level < restart) counters[deeper] = undefined;
    }

    const values: number[] = [];
    for (let i = 0; i <= level; i++) {
      const value = counters[i];
      // A higher level that was never used shows its start value.
      values.push(value === undefined ? this.startOf(numId, i) : value);
    }

    let text = definition.text === undefined ? '' : definition.text;
    if (definition.format !== 'bullet') {
      text = text.replace(/%([1-9])/g, (_match, digit: string) => {
        const index = Number(digit) - 1;
        if (index > level) return '';
        let format = index === level ? definition.format : levelOf(this.numbering, this.styles, numId, index)?.format;
        // Legal numbering shows every level it displays as a decimal, its own included (ECMA-376 Part 1, 17.9.4).
        if (definition.legal && format !== 'none') format = 'decimal';
        return formatCounter(values[index], format, lang);
      });
    }
    return { text, level: definition, values };
  }
}
