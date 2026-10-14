// What the tests of documents that are edited together share: a starting file, participants on
// a relay of the page whose messages the test delivers in whatever order it likes (each
// connection keeps its own order, as a socket does), and texts that say what a document holds.

import assert from 'node:assert/strict';
import { open, type OfficeFile } from '../src/index';
import { DocCollabSession, type DocChange, type DocSessionOptions } from '../src/collab/doc/session';
import type { DocWire, WirePos } from '../src/collab/doc/rebase';
import { contentVersion } from '../src/collab/protocol';
import type { CollabNotice } from '../src/collab/room';
import { MemoryServer, type MemoryTransport } from '../src/collab/server';
import { stable } from '../src/collab/wire';
import { readCollabStamp, writeCollabStamp } from '../src/formats/xlsx/collab-stamp';
import { writeDocx } from '../src/formats/docx/write';
import { paragraphText, type Block, type DocModel } from '../src/model/doc/types';
import { buildDocx, type DocxSpec } from './docx-fixtures';

export { seeded, pick, int } from './collab-fixtures';

const escape = (text: string): string => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** The XML of plain paragraphs. */
export const paragraphsXml = (paragraphs: readonly string[]): string => paragraphs.map((text) => (text === '' ? '<w:p/>' : `<w:p><w:r><w:t xml:space="preserve">${escape(text)}</w:t></w:r></w:p>`)).join('');

/** A document of plain paragraphs. */
export function docFile(paragraphs: readonly string[], spec: Partial<DocxSpec> = {}): Promise<Uint8Array> {
  return buildDocx({ ...spec, body: paragraphsXml(paragraphs) + (spec.body || '') });
}

/** The messages on their way, by connection and direction; the connections that were made with `instant` get theirs at once. */
export class DocNet {
  readonly server: MemoryServer;
  private lanes = new Map<string, (() => void)[]>();
  private instant = new Set<string>();
  private made = 0;

  constructor(options: { maxLogEntries?: number } = {}) {
    this.server = new MemoryServer({
      ...options,
      deliver: (run, lane) => {
        if (this.instant.has(lane.slice(0, lane.indexOf(':')))) return run();
        let waiting = this.lanes.get(lane);
        if (!waiting) this.lanes.set(lane, (waiting = []));
        waiting.push(run);
      },
    });
  }

  connect(instant = false): MemoryTransport {
    // The relay numbers its connections as they are made.
    if (instant) this.instant.add(String(this.made + 1));
    this.made++;
    return this.server.connect();
  }

  get waiting(): number {
    let n = 0;
    this.lanes.forEach((lane) => (n += lane.length));
    return n;
  }

  /** Delivers one message of a lane picked by `random`; false when nothing waits. */
  step(random: () => number): boolean {
    const busy = Array.from(this.lanes.values()).filter((lane) => lane.length > 0);
    if (busy.length === 0) return false;
    const run = busy[Math.floor(random() * busy.length)].shift()!;
    run();
    return true;
  }

  drain(random: () => number = () => 0, limit = 1000000): void {
    let n = 0;
    while (this.step(random)) if (++n > limit) assert.fail('The messages never ran out');
  }
}

export interface DocMember {
  name: string;
  file: OfficeFile;
  doc: DocModel;
  session: DocCollabSession;
  transport: MemoryTransport;
  notices: CollabNotice[];
  changes: DocChange[];
  loading?: Promise<void>;
}

export interface JoinOptions extends Partial<DocSessionOptions> {
  /** What arrives is delivered at once, not when the test says. */
  instant?: boolean;
  /** Sees every message that arrives, before the session does and after it did. */
  before?(message: { t: string; rev?: number; from?: string; id?: string }): void;
  after?(message: { t: string; rev?: number; from?: string; id?: string }): void;
  reload?(): Promise<Uint8Array>;
  unsaved?(): boolean;
}

/** A participant: reads `bytes` and joins. */
export async function joinDoc(net: DocNet, name: string, bytes: Uint8Array, options: JoinOptions = {}): Promise<DocMember> {
  const file = await open(bytes);
  const doc = await file.document!();
  const inner = net.connect(options.instant === true);
  const { instant, before, after, reload, unsaved, ...rest } = options;
  void instant;
  const transport: MemoryTransport = {
    ...inner,
    onMessage: (listener) =>
      inner.onMessage((text) => {
        const message = before || after ? JSON.parse(text) : undefined;
        if (before) before(message);
        listener(text);
        if (after) after(message);
      }),
  };
  const member: DocMember = { name, file, doc, session: undefined as unknown as DocCollabSession, transport, notices: [], changes: [] };
  member.session = new DocCollabSession(
    {
      document: () => member.doc,
      changed: (change) => member.changes.push(change),
      notice: (notice) => member.notices.push(notice),
      unsaved,
      reload: reload
        ? () => {
            const load = (async () => {
              const next = await reload();
              member.file = await open(next);
              member.doc = await member.file.document!();
              return { versions: [contentVersion(next)], stamp: await readCollabStamp(member.file.package!) };
            })();
            member.loading = load.then(
              () => new Promise<void>((resolve) => setImmediate(resolve)),
              () => undefined,
            );
            return load;
          }
        : undefined,
    },
    { transport, versions: [contentVersion(bytes)], stamp: await readCollabStamp(file.package!), user: { name }, timer: () => () => undefined, ...rest },
  );
  return member;
}

/** The file of a participant as mount() saves a document that is edited together: the room and the step go into it. */
export async function saveDoc(member: DocMember): Promise<Uint8Array> {
  await writeCollabStamp(member.file.package!, member.session.stampNow);
  await writeDocx(member.file.package!, member.doc);
  return member.file.package!.save();
}

/** A place of the body. */
export const at = (row: number, offset: number, story = 'body'): WirePos => ({ s: story, p: [row], o: offset });

/** Runs commands as a step of a participant; false when the document refused them. */
export function actDoc(member: DocMember, wires: DocWire[], meta?: { before?: unknown; after?: unknown; key?: string }): boolean {
  try {
    member.session.submit(wires, meta);
    return true;
  } catch (error) {
    if (error instanceof RangeError) return false;
    throw error;
  }
}

/** The text of the body: its paragraphs, a line each. Blocks that are no paragraphs are a line "[type]". */
export function bodyText(doc: DocModel): string {
  return doc.body.blocks.map((block) => (block.type === 'paragraph' ? paragraphText(block) : `[${block.type}]`)).join('\n');
}

/** The colour of every character of the body ("" where the text states none), paragraph ends left out. */
export function charColors(doc: DocModel): string[] {
  const out: string[] = [];
  for (const block of doc.body.blocks) {
    if (block.type !== 'paragraph') continue;
    for (const inline of block.children) {
      if (inline.type !== 'run') continue;
      const color = inline.props.color && inline.props.color.rgb ? inline.props.color.rgb : '';
      for (const item of inline.content) if (item.type === 'text') for (let i = 0; i < Array.from(item.text).length; i++) out.push(color);
    }
  }
  return out;
}

function plain(value: unknown, depth: number): unknown {
  if (depth > 200) return '…';
  if (value === null || typeof value !== 'object') return value;
  if (value instanceof Map) return Array.from(value.entries()).map(([key, item]) => [key, plain(item, depth + 1)]);
  if (value instanceof Uint8Array) return `bytes:${value.length}`;
  if (Array.isArray(value)) return value.map((item) => plain(item, depth + 1));
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(value as object)) {
    // Where a block stands in the file it was read from, and whether its part was edited, say nothing of what it holds.
    if (key === 'srcStart' || key === 'srcEnd' || key === 'modified' || key === 'source') continue;
    const item = (value as Record<string, unknown>)[key];
    if (item !== undefined) out[key] = plain(item, depth + 1);
  }
  return out;
}

const blocksText = (blocks: Block[]): unknown => plain(blocks, 0);

/** What a document holds, as a text that is the same for the same content. */
export function canonDoc(doc: DocModel): string {
  const notes = (collection: DocModel['footnotes']): unknown => (collection ? Array.from(collection.items.entries()).map(([id, note]) => [id, blocksText(note.blocks)]) : null);
  return stable({
    body: blocksText(doc.body.blocks),
    section: plain(doc.body.section, 0),
    headers: Array.from(doc.headers.entries()).map(([id, story]) => [id, blocksText(story.blocks)]),
    footers: Array.from(doc.footers.entries()).map(([id, story]) => [id, blocksText(story.blocks)]),
    footnotes: notes(doc.footnotes),
    endnotes: notes(doc.endnotes),
    numbering: plain(doc.numbering, 0),
  });
}

/** The text of a list of blocks with how it looks, whatever runs it is cut into: paragraphs as their properties and stretches of text that look alike. */
function looks(blocks: Block[], depth = 0): unknown {
  return blocks.map((block) => {
    if (block.type === 'table') return depth > 20 ? 'table' : { table: block.rows.map((row) => row.cells.map((cell) => looks(cell.blocks, depth + 1))) };
    if (block.type !== 'paragraph') return block.type;
    const parts: [string, string][] = [];
    for (const inline of block.children) {
      if (inline.type !== 'run') continue;
      // What says where a tracked change or a field begins is not how the text looks.
      const props = stable(plain(inline.props, 0));
      let text = '';
      for (const item of inline.content) text += item.type === 'text' ? item.text : `[${item.type}]`;
      if (text === '') continue;
      const last = parts[parts.length - 1];
      if (last && last[1] === props) last[0] += text;
      else parts.push([text, props]);
    }
    return { props: plain(block.props, 0), parts };
  });
}

/** What a document says and how it looks, the same for a document that was edited and for the same document read from its file. */
export function lookOfDoc(doc: DocModel): string {
  return stable({ body: looks(doc.body.blocks), headers: Array.from(doc.headers.entries()).map(([id, story]) => [id, looks(story.blocks)]) });
}

/** Where two texts first differ, with what stands around it: for a failure that can be read. */
export function firstDifference(a: string, b: string): string {
  let i = 0;
  while (i < a.length && i < b.length && a[i] === b[i]) i++;
  return `at ${i}: …${a.slice(Math.max(0, i - 80), i + 80)}… against …${b.slice(Math.max(0, i - 80), i + 80)}…`;
}

export function assertSameDoc(members: DocMember[], label: string): void {
  const texts = members.map((member) => canonDoc(member.doc));
  for (let i = 1; i < texts.length; i++) {
    if (texts[i] !== texts[0]) assert.fail(`${label}: ${members[i].name} differs from ${members[0].name} ${firstDifference(texts[0], texts[i])}`);
  }
}
