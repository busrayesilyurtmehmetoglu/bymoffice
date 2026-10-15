// What a command of a document was, kept beside the operations it returned: the view of a
// document calls the commands itself and applies what they return, and where the document is
// edited together the change has to travel as the command, not as its operations (which hold
// blocks of this participant's document by reference).
//
// traced() wraps the commands that can be shared. Each call notes, under the list of
// operations the command returned, the command as it travels; wiresOf() finds it again. A
// change whose operations no shared command returned cannot be shared, and is refused where
// the document is edited together (see history.ts). Outside of that nothing reads the notes.

import { paragraphText, type DocModel } from '../../model/doc/types';
import type { DocEdit } from '../../ops/doc/draft';
import { rangeFormat, type ParaFormatPatch, type RunFormatPatch, type ToggleKey } from '../../ops/doc/format';
import type { DocOp } from '../../ops/doc/ops';
import { asRange, comparePositions, nextCharOffset, nextWordOffset, resolvePosition, samePosition, type DocPosition, type DocRange } from '../../ops/doc/position';
import { deleteRange } from '../../ops/doc/text';
import type { DocWire, WirePos } from './rebase';
import { runWire, toWirePos } from './wire';
import { COMMENT_COMMANDS, commentWireOf } from './comments';

const MADE = new WeakMap<DocOp[], DocWire[]>();

/** How many views are in a room: the commands of a view are noted only while one is. */
let watching = 0;

/** Called by a view when it enters a room, and when it is destroyed: outside of rooms the commands are run as they are. */
export function tracing(on: boolean): void {
  watching = Math.max(0, watching + (on ? 1 : -1));
}

/** The commands a list of operations came from, as they travel; undefined for operations no shared command returned. */
/** An id of a relationship nobody holds: random, so that no participant has it from before (eight hex digits and a count). */
let relationships = 0;
function freshRelationshipId(): string {
  relationships = (relationships + 1) % 0x10000;
  // (A letter between the two: no run of digits is longer than nine, which the wire does not take.)
  return 'rIdC' + ('0000000' + Math.floor(Math.random() * 0x100000000).toString(16)).slice(-8) + 'n' + relationships.toString(16);
}

export function wiresOf(ops: DocOp[]): DocWire[] | undefined {
  return MADE.get(ops);
}

const same = (a: DocPosition, b: DocPosition): boolean => samePosition(a, b) && (a.before === true) === (b.before === true);

/** The two ends of a range as they travel; `to` is left out for a caret. */
function ends(at: DocPosition | DocRange): { at: WirePos; to?: WirePos } {
  const range = asRange(at);
  return same(range.anchor, range.focus) ? { at: toWirePos(range.anchor) } : { at: toWirePos(range.anchor), to: toWirePos(range.focus) };
}

function both(at: DocPosition | DocRange): { at: WirePos; to: WirePos } {
  const range = asRange(at);
  return { at: toWirePos(range.anchor), to: toWirePos(range.focus) };
}

/** Whether two values hold the same, whatever objects they are made of. */
function alike(a: unknown, b: unknown, depth: number): boolean {
  if (a === b) return true;
  if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null || depth > 96) return false;
  if (Array.isArray(a) || Array.isArray(b)) {
    if (!Array.isArray(a) || !Array.isArray(b) || a.length !== b.length) return false;
    for (let i = 0; i < a.length; i++) if (!alike(a[i], b[i], depth + 1)) return false;
    return true;
  }
  if (a instanceof Map || b instanceof Map) {
    if (!(a instanceof Map) || !(b instanceof Map) || a.size !== b.size) return false;
    for (const [key, value] of a) if (!b.has(key) || !alike(value, b.get(key), depth + 1)) return false;
    return true;
  }
  if (a instanceof Uint8Array || b instanceof Uint8Array) return false;
  const left = a as Record<string, unknown>;
  const right = b as Record<string, unknown>;
  const keys = Object.keys(left).filter((key) => left[key] !== undefined);
  if (keys.length !== Object.keys(right).filter((key) => right[key] !== undefined).length) return false;
  for (const key of keys) if (!alike(left[key], right[key], depth + 1)) return false;
  return true;
}

/**
 * Backspace and Delete as the deletion of a range, where that is all they did: the range can
 * be read through what others do, a key cannot (the character before the caret may be
 * another one by then). Where the key did something else - took a paragraph out of a list,
 * took an indent back - it travels as the key.
 */
function deletion(doc: DocModel, position: DocPosition, unit: 'char' | 'word', forward: boolean, edit: DocEdit): DocWire {
  const key: DocWire = { t: 'cmd', name: forward ? 'deleteForward' : 'deleteBackward', at: toWirePos(position), a: [unit] };
  try {
    let from: DocPosition;
    let to: DocPosition;
    if (forward) {
      const at = resolvePosition(doc, position);
      from = position;
      if (position.offset < at.length) {
        const text = paragraphText(at.paragraph);
        to = { story: position.story, path: position.path.slice(), offset: unit === 'word' ? nextWordOffset(text, position.offset) : nextCharOffset(text, position.offset) };
      } else {
        const next = at.blocks[at.index + 1];
        if (!next || next.type !== 'paragraph') return key;
        to = { story: position.story, path: at.container.concat([at.index + 1]), offset: 0 };
      }
    } else {
      from = edit.selection.focus;
      to = position;
      if (from.story !== to.story || comparePositions(from, to) >= 0) return key;
    }
    const range = deleteRange(doc, { anchor: from, focus: to });
    if (alike(range.ops, edit.ops, 0)) return { t: 'del', at: toWirePos(from), to: toWirePos(to) };
  } catch {
    // Not a range that can be deleted: the key it is.
  }
  return key;
}

type Maker = (doc: DocModel, args: unknown[], edit: DocEdit) => DocWire | undefined;

/** What each shared command travels as. `args`: what it was called with after the document. Undefined: this call cannot be shared (it asks for a tracked change, say). */
const MAKERS: { [name: string]: Maker } = {
  insertText: (_doc, args) => {
    if (args[3] !== undefined || typeof args[1] !== 'string') return undefined;
    const wire: DocWire = { t: 'text', ...ends(args[0] as DocPosition | DocRange), x: args[1] };
    if (args[2] !== undefined && args[2] !== null) wire.f = args[2] as RunFormatPatch;
    return wire;
  },
  deleteRange: (_doc, args) => (args[1] !== undefined ? undefined : { t: 'del', ...both(args[0] as DocRange) }),
  splitParagraph: (_doc, args) => (args[1] !== undefined ? undefined : { t: 'split', at: toWirePos(args[0] as DocPosition) }),
  deleteBackward: (doc, args, edit) => (args[2] !== undefined ? undefined : deletion(doc, args[0] as DocPosition, args[1] === 'word' ? 'word' : 'char', false, edit)),
  deleteForward: (doc, args, edit) => (args[2] !== undefined ? undefined : deletion(doc, args[0] as DocPosition, args[1] === 'word' ? 'word' : 'char', true, edit)),
  setRunFormat: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'run', ...both(args[0] as DocRange), f: args[1] as RunFormatPatch }),
  toggleRunFormat: (doc, args) => {
    if (args[2] !== undefined) return undefined;
    const key = args[1] as ToggleKey;
    // Turning it on is setting it; turning it off is said as that, so that it never turns anything on.
    return rangeFormat(doc, args[0] as DocRange)[key] === true ? { t: 'unrun', ...both(args[0] as DocRange), k: key } : { t: 'run', ...both(args[0] as DocRange), f: { [key]: true } };
  },
  setParaFormat: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'para', ...both(args[0] as DocPosition | DocRange), f: args[1] as ParaFormatPatch }),
  toggleList: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'cmd', name: 'toggleList', ...both(args[0] as DocPosition | DocRange), a: [args[1]] }),
  indentList: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'cmd', name: 'indentList', ...both(args[0] as DocPosition | DocRange), a: [args[1]] }),
  insertPageBreak: (_doc, args) => (args[1] !== undefined ? undefined : { t: 'cmd', name: 'insertPageBreak', at: toWirePos(args[0] as DocPosition), a: [] }),
  insertTable: (_doc, args) => (args[3] !== undefined ? undefined : { t: 'cmd', name: 'insertTable', at: toWirePos(args[0] as DocPosition), a: [args[1], args[2]] }),
  insertRow: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'cmd', name: 'insertRow', at: toWirePos(args[0] as DocPosition), a: [args[1]] }),
  deleteRow: (_doc, args) => (args[1] !== undefined ? undefined : { t: 'cmd', name: 'deleteRow', at: toWirePos(args[0] as DocPosition), a: [] }),
  insertColumn: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'cmd', name: 'insertColumn', at: toWirePos(args[0] as DocPosition), a: [args[1]] }),
  deleteColumn: (_doc, args) => (args[1] !== undefined ? undefined : { t: 'cmd', name: 'deleteColumn', at: toWirePos(args[0] as DocPosition), a: [] }),
  // The link gets a relationship id of its own, which travels: everybody makes the same (see wire.ts).
  insertHyperlink: (_doc, args) => (args[2] !== undefined ? undefined : { t: 'cmd', name: 'insertHyperlink', ...both(args[0] as DocRange), a: [args[1], freshRelationshipId()] }),
};

/**
 * The commands of the document with the ones that can be shared wrapped: what they return
 * is what they returned before, and the list of its operations is known by the command.
 * Without `always` they are noted only while a view is in a room (see tracing()).
 */
export function traced<T extends object>(api: T, always = true): T {
  const out = { ...api } as Record<string, unknown>;
  for (const name of Object.keys(MAKERS)) {
    const command = (api as Record<string, unknown>)[name];
    if (typeof command !== 'function') continue;
    const make = MAKERS[name];
    out[name] = (doc: DocModel, ...args: unknown[]): DocEdit => {
      if (!always && watching === 0) return (command as (doc: DocModel, ...rest: unknown[]) => DocEdit)(doc, ...args);
      // What the command is said to have been is read off the document as the command found it.
      let before: DocWire | undefined;
      const early = name === 'toggleRunFormat';
      if (early) {
        try {
          before = make(doc, args, undefined as unknown as DocEdit);
        } catch {
          before = undefined;
        }
      }
      const edit = (command as (doc: DocModel, ...rest: unknown[]) => DocEdit)(doc, ...args);
      if (edit && Array.isArray(edit.ops) && edit.ops.length > 0) {
        let wire: DocWire | undefined;
        try {
          wire = early ? before : make(doc, args, edit);
        } catch {
          wire = undefined;
        }
        if (wire) MADE.set(edit.ops, [wire]);
      }
      return edit;
    };
  }
  // A command of comments travels with what its author's side decided (who, when) and the ids it gives; what it does
  // here is what it does at everybody: the edit handed back is the one the command makes of the document (its new
  // comment has the number everybody gives it), with the selection the call left.
  for (const name of COMMENT_COMMANDS) {
    const command = (api as Record<string, unknown>)[name];
    if (typeof command !== 'function') continue;
    out[name] = (doc: DocModel, ...args: unknown[]): DocEdit => {
      const edit = (command as (doc: DocModel, ...rest: unknown[]) => DocEdit & { id?: number })(doc, ...args);
      if (!always && watching === 0) return edit;
      if (!edit || !Array.isArray(edit.ops) || edit.ops.length === 0) return edit;
      const wire = commentWireOf(doc, name, args, edit);
      if (!wire) return edit;
      const shared = runWire(doc, wire) as DocEdit & { id?: number };
      const made: DocEdit & { id?: number } = { ops: shared.ops, selection: edit.selection };
      if (shared.id !== undefined) made.id = shared.id;
      MADE.set(made.ops, [wire]);
      return made;
    };
  }
  return out as T;
}
