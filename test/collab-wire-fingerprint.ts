// The fingerprint of the wire of editing together: which kinds of steps the decoders of
// workbooks and documents take, and which keys they (and the rules that interpret what they
// copy: patchRunProps, patchParaProps, paragraphOf) read by name, along sample steps of every
// kind. Nothing of it is taken from the source text: the decoders are run over the samples
// through proxies that write down every key that is asked for, whether it is there or not (an
// optional field is asked for with "has" even where a sample leaves it out).
//
// What it does not see: what a command is called with (`cmd.a` is copied as plain data and
// read by the operations of the model), the values a field takes (the members of an enum),
// fields of objects no sample holds, fields checked against a table rather than asked for by
// name (of an indent, of spacing: there the keys of the sample are what is seen, so the
// samples hold all of them), and the messages of the relay (join, hello, ...). A change of
// the samples alone changes the fingerprint too: then PINNED is updated without a new version.

import { createBlank } from '../src/formats/blank';
import { open } from '../src/index';
import type { StyleSheet } from '../src/model/sheet/types';
import { commandNames, decodeStep as decodeSheetStep, WIRE_KINDS } from '../src/collab/wire';
import { COMMANDS as DOC_COMMANDS, DOC_WIRE_KINDS, decodeStep as decodeDocStep, paragraphOf } from '../src/collab/doc/wire';
import { patchParaProps, patchRunProps, type ParaFormatPatch, type RunFormatPatch } from '../src/ops/doc/format';
import { EMPTY_PARA_PROPS, EMPTY_RUN_PROPS } from '../src/model/doc/types';
import type { PutParagraph } from '../src/collab/doc/rebase';

/** Keys asked for, by the path of the object they were asked of. */
type Seen = Map<string, Set<string>>;

/**
 * Paths whose objects a decoder copies as plain data, key by key as they come: what it reads of
 * them says what a sample holds, not what the decoder knows. (What a patch holds is read by name
 * by the rule that interprets it, which is probed on its own.)
 */
const COPIED = ['sheet:cmd.a', 'doc:cmd.a', 'doc:text.f', 'doc:run.f', 'doc:para.f', 'doc:put.x[].p', 'doc:put.x[].m', 'doc:put.x[].r[].p'];

const copied = (path: string): boolean => COPIED.some((prefix) => path === prefix || path.startsWith(prefix + '.') || path.startsWith(prefix + '['));

function note(seen: Seen, path: string, key: string): void {
  let keys = seen.get(path);
  if (!keys) seen.set(path, (keys = new Set()));
  keys.add(key);
}

/** `value` behind a proxy that writes down every key asked of it and of what it holds. */
export function probe<T>(value: T, path: string, seen: Seen): T {
  if (value === null || typeof value !== 'object' || copied(path)) return value;
  const list = Array.isArray(value);
  return new Proxy(value as unknown as object, {
    get(target, key) {
      const held = Reflect.get(target, key);
      if (typeof key !== 'string') return held;
      if (list) return /^(0|[1-9][0-9]*)$/.test(key) ? probe(held, `${path}[]`, seen) : held;
      note(seen, path, key);
      return probe(held, `${path}.${key}`, seen);
    },
    has(target, key) {
      if (!list && typeof key === 'string') note(seen, path, key);
      return Reflect.has(target, key);
    },
    getOwnPropertyDescriptor(target, key) {
      if (!list && typeof key === 'string') note(seen, path, key);
      return Reflect.getOwnPropertyDescriptor(target, key);
    },
  }) as unknown as T;
}

const FONT = { bold: false, italic: false, strike: false, underline: 'none' };
const COLOR = { rgb: 'FF0000' };
const RANGE = { top: 0, left: 0, bottom: 3, right: 1 };

/** A step of a workbook with every kind of operation, and in it every object an operation can hold. */
export function sheetSample(): unknown[] {
  return [
    {
      t: 'st',
      f: [
        {
          n: '0.00',
          f: { ...FONT, name: 'Arial', size: 11, color: COLOR },
          l: { type: 'pattern', pattern: 'solid', fg: { theme: 1, tint: 0.5 }, bg: { indexed: 64 } },
          b: { diagonalUp: false, diagonalDown: false, left: { style: 'thin', color: { auto: true } } },
          x: 0,
          a: { horizontal: 'left', wrapText: true },
          p: { locked: true },
          q: true,
        },
      ],
    },
    { t: 'cells', s: 0, c: [[0, 0, { v: { k: 'string', v: 'x', f: { x: 'A1', a: [['ca', '1']] }, r: [{ text: 'x', font: { ...FONT, color: COLOR } }], ph: true }, s: 0 }], [1, 1, 0]] },
    { t: 'rows', s: 0, r: [[0, { h: 20, ch: true, hid: false, s: 0 }]] },
    { t: 'cols', s: 0, p: [[0, 1, { w: 10, cw: true, hid: false, s: 0, ol: 0, col: false }], [2, 2, 0]] },
    { t: 'merges', s: 0, add: [[0, 0, 1, 1]], del: [[2, 0, 2, 1]] },
    {
      t: 'filter',
      s: 0,
      f: {
        range: RANGE,
        columns: [{ colId: 0, kind: 'values', values: ['a'], dateGroups: [{ grouping: 'year', year: 2020 }], conditions: [{ operator: 'equal', value: '1' }] }],
        sortState: { range: RANGE, conditions: [{ range: RANGE, descending: false }] },
        held: [1],
      },
      show: [1],
      hide: [2],
    },
    { t: 'sort', s: 0, g: [0, 0, 3, 1], o: [0, 1, 2, 3] },
    { t: 'ins', s: 0, x: 'row', at: 1, n: 1 },
    { t: 'del', s: 0, x: 'col', at: 1, n: 1 },
    { t: 'notes', s: 0, n: [[0, 0, { x: 't', a: 'Someone', r: [{ text: 't', font: FONT }], vis: true }], [1, 1, 0]] },
    { t: 'sheet+', s: 1, name: 'New' },
    { t: 'sheet-', s: 1 },
    { t: 'name', s: 0, name: 'Renamed' },
    { t: 'move', s: 0, to: 1 },
    { t: 'state', s: 0, state: 'hidden' },
    { t: 'tab', s: 0, color: COLOR },
    { t: 'pane', s: 0, r: 1, c: 1, tl: 'B2', fv: 'A1' },
    { t: 'cmd', s: 0, name: 'removeChart', a: [0] },
  ];
}

const place = (offset: number): unknown[] => ['body', [0], offset];
const TABS = [{ pos: 720, align: 'left', leader: 'dot' }];
// The fields of an indent and of spacing are checked against tables of the rules, not asked for by name: the samples hold every one of them.
const INDENT = { left: null, right: null, firstLine: null, hanging: null, leftChars: null, rightChars: null, firstLineChars: null, hangingChars: null };
const SPACING = { before: null, after: null, line: null, beforeLines: null, afterLines: null, lineRule: null, beforeAuto: null, afterAuto: null };
const PARA: ParaFormatPatch = { align: 'left', indent: INDENT, spacing: SPACING, numbering: { numId: 1, level: 0 }, tabs: { set: TABS, clear: [1440], clearAll: false } };
const RUN: RunFormatPatch = { bold: true, color: 'FF0000', font: 'Arial', size: 24 };
const PUT_FORMAT = { f: RUN, cs: { bold: true }, fonts: { ascii: 'Arial', hint: 'default' }, lang: { val: 'tr-TR' } };

/** A step of a document with every kind of command. */
export function docSample(): unknown[] {
  return [
    { t: 'text', at: place(0), to: place(1), x: 'hi', f: RUN },
    { t: 'del', at: place(0), to: place(1) },
    { t: 'split', at: place(0) },
    { t: 'run', at: place(0), to: place(1), f: RUN },
    { t: 'unrun', at: place(0), to: place(1), k: 'bold' },
    { t: 'para', at: place(0), to: place(1), f: PARA },
    { t: 'cmd', name: 'insertPageBreak', at: place(0), a: [] },
    {
      t: 'put',
      at: place(0),
      n: 1,
      h: '00000000',
      x: [
        {
          p: { align: 'left', indent: INDENT, spacing: SPACING, tabs: TABS },
          m: PUT_FORMAT,
          a: [['w:rsidR', '00000001']],
          r: [{ c: ['x', 0, ['y', 'z']], p: PUT_FORMAT, a: [['w:rsidR', '00000001']] }, { e: 'spellStart' }, { b: [0, '1', 'name'] }, { b: [1, '1'] }],
        },
      ],
    },
  ];
}

export interface Decoders {
  sheet(step: unknown[], styles: StyleSheet): { t: string }[];
  doc(step: unknown[]): { t: string }[];
}

const DECODERS: Decoders = { sheet: decodeSheetStep, doc: decodeDocStep };

let stylesOnce: Promise<StyleSheet> | undefined;
async function styles(): Promise<StyleSheet> {
  if (!stylesOnce) stylesOnce = createBlank('xlsx').then(async (bytes) => (await (await open(bytes)).workbook()).styles);
  return stylesOnce;
}

/** Each operation of a step behind its own proxy, named by its kind (as a step arrives: JSON). */
const probed = (side: string, step: unknown[], seen: Seen): unknown[] => (JSON.parse(JSON.stringify(step)) as { t: string }[]).map((op) => probe(op, `${side}:${op.t}`, seen));

/**
 * The fingerprint: one line for the kinds and the commands of each side, and one for every
 * path that was read, with the keys read of it. Sorted, so that it reads as a list.
 */
export async function wireFingerprint(decoders: Decoders = DECODERS): Promise<string[]> {
  const seen: Seen = new Map();
  const lines: string[] = [];

  const sheetKinds = Object.keys(WIRE_KINDS).sort();
  const sheetOps = decoders.sheet(probed('sheet', sheetSample(), seen), await styles());
  const decodedSheet = Array.from(new Set(sheetOps.map((op) => op.t))).sort();
  if (decodedSheet.join() !== sheetKinds.join()) throw new Error(`The sample of a workbook step has to hold every kind of operation: ${sheetKinds.join(' ')}, it holds ${decodedSheet.join(' ')}`);
  lines.push(`sheet kinds: ${sheetKinds.join(' ')}`);
  lines.push(`sheet commands: ${commandNames().sort().join(' ')}`);

  const docKinds = Object.keys(DOC_WIRE_KINDS).sort();
  const wires = decoders.doc(probed('doc', docSample(), seen));
  const decodedDoc = Array.from(new Set(wires.map((wire) => wire.t))).sort();
  if (decodedDoc.join() !== docKinds.join()) throw new Error(`The sample of a document step has to hold every kind of command: ${docKinds.join(' ')}, it holds ${decodedDoc.join(' ')}`);
  lines.push(`doc kinds: ${docKinds.join(' ')}`);
  lines.push(`doc commands: ${Object.keys(DOC_COMMANDS).sort().join(' ')}`);

  // What the decoder copies is read by the rules that make the model of it: those are probed with what was decoded.
  for (const wire of wires as { t: string; f?: unknown; x?: unknown }[]) {
    if (wire.t === 'text' || wire.t === 'run') patchRunProps(EMPTY_RUN_PROPS, probe(wire.f as RunFormatPatch, `doc:${wire.t}~f`, seen));
    else if (wire.t === 'para') patchParaProps(EMPTY_PARA_PROPS, probe(wire.f as ParaFormatPatch, 'doc:para~f', seen));
    else if (wire.t === 'put') for (const paragraph of wire.x as PutParagraph[]) paragraphOf(probe(paragraph, 'doc:put~x[]', seen));
  }

  for (const [path, keys] of seen) lines.push(`${path}: ${Array.from(keys).sort().join(' ')}`);
  return lines.sort();
}

/**
 * What is wrong with the protocol version, if anything: the wire is pinned at version `pinned.pv`
 * with the fingerprint `pinned.lines`, and this side speaks `pv`.
 */
export function versionProblem(lines: readonly string[], pv: number, pinned: { pv: number; lines: readonly string[] }): string | undefined {
  const same = lines.length === pinned.lines.length && lines.every((line, i) => line === pinned.lines[i]);
  if (same && pv === pinned.pv) return undefined;
  const added = lines.filter((line) => pinned.lines.indexOf(line) < 0);
  const removed = pinned.lines.filter((line) => lines.indexOf(line) < 0);
  const diff = [...removed.map((line) => `- ${line}`), ...added.map((line) => `+ ${line}`)].join('\n');
  if (!same && pv === pinned.pv) return `The wire of editing together changed while PROTOCOL_VERSION stayed ${pv}: raise PROTOCOL_VERSION (src/collab/room.ts) and update the fingerprint (PINNED in test/collab-protocol-version.test.ts).\n${diff}`;
  if (!same) return `PROTOCOL_VERSION is ${pv}: update the fingerprint (PINNED in test/collab-protocol-version.test.ts) to the wire of version ${pv}.\n${diff}`;
  return `PROTOCOL_VERSION is ${pv} but the fingerprint is pinned at version ${pinned.pv} with the same wire: set PINNED.pv to ${pv} (or take the change of version back).`;
}
