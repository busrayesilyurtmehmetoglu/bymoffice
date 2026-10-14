// A second model of a text that is edited together, written by the independent review of the
// first one and without looking at it (test/collab-doc-model.ts). It knows nothing of how the
// session works: characters are unique, and what a change MEANT is read off the document of its
// author before and after it (P -> Q); what it DID is read off the placed document before and
// after the step (B -> A) at a witness whom every message reaches at once. Undo and redo, tables,
// headers, a footnote, colours and indents are in it. Kept as it was written (one statement of a
// check per line, however long), with one refinement that is marked where it is made.
import { DocNet, joinDoc, canonDoc, seeded, firstDifference, type DocMember } from './collab-doc-fixtures';
import { buildDocx } from './docx-fixtures';
import { paragraphText, type Block, type DocModel } from '../src/model/doc/types';
import type { DocWire, WirePos } from '../src/collab/doc/rebase';
import { writeDocx } from '../src/formats/docx/write';
import { movedSelection } from '../src/ui/doc-collab';
import { containerBlocks } from '../src/ops/doc/position';

const TBL = '☃';
export interface Snap { text: Map<string, string>; color: Map<string, string>; indent: Map<string, number[]> }

function contOf(doc: DocModel): { key: string; story: string; prefix: number[]; blocks: Block[] }[] {
  const out: { key: string; story: string; prefix: number[]; blocks: Block[] }[] = [];
  out.push({ key: 'body', story: 'body', prefix: [], blocks: doc.body.blocks });
  doc.body.blocks.forEach((b, i) => {
    if (b.type === 'table') b.rows.forEach((row, r) => row.cells.forEach((cell, c) => out.push({ key: `cell:${r},${c}`, story: 'body', prefix: [i, r, c], blocks: cell.blocks })));
  });
  doc.headers.forEach((story, id) => out.push({ key: `header:${id}`, story: `header:${id}`, prefix: [], blocks: story.blocks }));
  if (doc.footnotes) doc.footnotes.items.forEach((note, id) => { if (id > 0) out.push({ key: `footnote:${id}`, story: `footnote:${id}`, prefix: [], blocks: note.blocks }); });
  return out;
}
export function snap(doc: DocModel): Snap {
  const s: Snap = { text: new Map(), color: new Map(), indent: new Map() };
  for (const c of contOf(doc)) {
    s.text.set(c.key, c.blocks.map((b) => (b.type === 'paragraph' ? paragraphText(b) : TBL)).join('\n'));
    s.indent.set(c.key, c.blocks.map((b) => (b.type === 'paragraph' && b.props.indent && typeof b.props.indent.left === 'number' ? b.props.indent.left : 0)));
    for (const b of c.blocks) {
      if (b.type !== 'paragraph') continue;
      for (const inl of b.children) {
        if (inl.type !== 'run') continue;
        const col = inl.props.color && inl.props.color.rgb ? inl.props.color.rgb : '';
        for (const item of inl.content) if (item.type === 'text') for (const ch of item.text) s.color.set(ch, col);
      }
    }
  }
  return s;
}
function posOf(doc: DocModel, key: string, index: number): WirePos {
  const c = contOf(doc).find((x) => x.key === key)!;
  const text = c.blocks.map((b) => (b.type === 'paragraph' ? paragraphText(b) : TBL)).join('\n');
  let row = 0, start = 0;
  for (let i = 0; i < index; i++) if (text[i] === '\n') { row++; start = i + 1; }
  return { s: c.story, p: c.prefix.concat([row]), o: index - start };
}

type Intent =
  | { kind: 'text'; key: string; P: Snap; Q: Snap; what: string }
  | { kind: 'color'; key: string; P: Snap; V: string[]; c: string; what: string }
  | { kind: 'indent'; key: string; P: Snap; gap: number; para: string[]; left?: string; right?: string; v: number; what: string }
  | { kind: 'undo'; P: Snap; Q: Snap; what: string };

const isCh = (c: string): boolean => c !== '\n';
const chars = (s: string): string => s.split('\n').join('');
const count = (s: string): number => s.length - chars(s).length;

/** Text check of one container. Returns [error | '', exact]. */
function textCheck(P: string, Q: string, B: string, A: string, loose = false, revived: Set<string> = new Set()): [string, boolean] {
  if (P === Q) return [A === B ? '' : 'text changed by a step that changed no text for its author', true];
  let a = 0;
  while (a < P.length && a < Q.length && P[a] === Q[a]) a++;
  let z = 0;
  while (z < P.length - a && z < Q.length - a && P[P.length - 1 - z] === Q[Q.length - 1 - z]) z++;
  const D = P.slice(a, P.length - z), N = Q.slice(a, Q.length - z);
  // anchors
  let li = -1, lp = -1;
  for (let i = a - 1; i >= 0; i--) if (isCh(P[i]) && B.indexOf(P[i]) >= 0) { lp = i; li = B.indexOf(P[i]); break; }
  let ri = B.length, rp = P.length;
  for (let i = P.length - z; i < P.length; i++) if (isCh(P[i]) && B.indexOf(P[i]) >= 0) { rp = i; ri = B.indexOf(P[i]); break; }
  if (li >= ri) return [`anchors out of order in the placed document`, false];
  const la = lp >= 0 ? A.indexOf(P[lp]) : -1;
  const ra = rp < P.length ? A.indexOf(P[rp]) : A.length;
  if ((lp >= 0 && la < 0) || (rp < P.length && ra < 0)) return ['an anchor the author did not delete went', false];
  if (A.slice(0, la + 1) !== B.slice(0, li + 1)) return ['text BEFORE the left anchor changed', false];
  if (A.slice(ra) !== B.slice(ri)) return ['text AFTER the right anchor changed', false];
  const Pseg = P.slice(lp + 1, rp), Bseg = B.slice(li + 1, ri), Aseg = A.slice(la + 1, ra);
  const Qseg = Q.slice(lp + 1, rp + (Q.length - P.length));
  if (Bseg === Pseg && !loose) return [Aseg === Qseg ? '' : `no concurrency near: expected ${JSON.stringify(Qseg)} got ${JSON.stringify(Aseg)}`, true];
  // weak: chars
  // (Refined: P -> Q need not be ONE replacement - an undoing may delete two stretches with unchanged characters
  // between them. What is in both D and N was neither deleted nor inserted.)
  const dAll = new Set(chars(D).split('')), nAll = new Set(chars(N).split(''));
  const dset = new Set(Array.from(dAll).filter((c) => !nAll.has(c)));
  const nstr = chars(N).split('').filter((c) => !dAll.has(c)).join('');
  const nset = new Set(nstr.split(''));
  // (Refined: a character the change meant to delete that someone else deleted first and then brought back by
  // undoing stays - the deletion went with the other one's, and was told so: see collab-doc-model.ts.)
  const expectOld = chars(Bseg).split('').filter((c) => !dset.has(c) || revived.has(c)).join('');
  const ach = chars(Aseg);
  const gotOld = ach.split('').filter((c) => !nset.has(c)).join('');
  if (gotOld !== expectOld) return [`surviving characters: expected ${JSON.stringify(expectOld)} got ${JSON.stringify(gotOld)} (B ${JSON.stringify(Bseg)} A ${JSON.stringify(Aseg)} D ${JSON.stringify(D)} N ${JSON.stringify(N)})`, false];
  // (Refined as well: where the change kept characters between what it put in - an undoing that puts back two stretches
  // around a character that stayed - what it put in has to stand in its order, not in one piece.)
  const kept = Array.from(dAll).some((c) => nAll.has(c));
  const inOrder = ach.split('').filter((c) => nset.has(c)).join('') === nstr;
  if (nstr.length > 0 && (kept ? !inOrder : ach.indexOf(nstr) < 0)) return [`inserted characters not contiguous/in order: ${JSON.stringify(Aseg)} for ${JSON.stringify(N)}`, false];
  const m = count(Aseg), mb = count(Bseg);
  const sl = (z > 0 && P[P.length - z] === '\n' ? 1 : 0) + (a > 0 && P[a - 1] === '\n' ? 1 : 0);
  if (m < mb - count(D) - sl || m > mb + count(N) + sl) return [`paragraph ends: ${m}, before ${mb}, deleted ${count(D)}, inserted ${count(N)}`, false];
  return ['', false];
}

export interface Cfg { seed: number; members: number; ticks: number; deliver: number; cuts: boolean; undo: number; quiesce?: boolean; fmt: boolean; dump?: string; trace?: boolean }
export interface Res { bad: string[]; made: number; placed: number; dropped: number; quiet: number; exact: number; weak: number; droppedKinds: Record<string, number>; notices: Record<string, number>; undoAsked: number; undoDone: number; unknown: number; caretMoves: number; caretOk: number; caretLost: number }

const FOOT = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';

export async function run(cfg: Cfg): Promise<Res> {
  const random = seeded(cfg.seed);
  const int = (lo: number, hi: number): number => lo + Math.floor(random() * (hi - lo + 1));
  let next = 0x4e00;
  const fresh = (n: number): string => { let s = ''; for (let i = 0; i < n; i++) s += String.fromCharCode(next++); return s; };
  const p = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
  // (Added: paragraphs with ids and rsids, a bookmark in one, one that lies over two: what a put has to bring back,
  // and may not bring twice.)
  let ids = 0x10000000;
  const named = (t: string, open = '', close = ''): string => `<w:p w14:paraId="${(ids++).toString(16).toUpperCase()}" w14:textId="77777777" w:rsidR="00AB12CD">${open}<w:r w:rsidRPr="00CD34EF"><w:t>${t}</w:t></w:r>${close}</w:p>`;
  const body = named(fresh(6), '<w:bookmarkStart w:id="1" w:name="one"/>', '<w:bookmarkEnd w:id="1"/>') + named(fresh(8), '<w:bookmarkStart w:id="2" w:name="over"/>') + named(fresh(5), '', '<w:bookmarkEnd w:id="2"/>') +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${p(fresh(4))}</w:tc><w:tc>${p(fresh(4))}</w:tc></w:tr><w:tr><w:tc>${p(fresh(3))}${p(fresh(3))}</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
    `<w:p><w:r><w:t>${fresh(3)}</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t>${fresh(2)}</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>`;
  const bytes = await buildDocx({ body, headers: { rIdH: p(fresh(6)) + p(fresh(4)) }, footnotes: FOOT + `<w:footnote w:id="1">${p(fresh(7))}${p(fresh(4))}</w:footnote>`,
    styles: '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' });
  const net = new DocNet();
  const bad: string[] = [];
  const fail = (t: string): void => { if (bad.length < 3) bad.push(t); };
  const intents = new Map<string, Intent>();
  const res: Res = { bad, made: 0, placed: 0, dropped: 0, quiet: 0, exact: 0, weak: 0, droppedKinds: {}, notices: {}, undoAsked: 0, undoDone: 0, unknown: 0, caretMoves: 0, caretOk: 0, caretLost: 0 };
  let witness: DocMember;
  let B: Snap | undefined;
  let rev = 0;
  const markRevs: { rev: number; key: string }[] = [];
  /** How often the author of a step had been told `partial` when it made the step: a later one is about it (or one after it). */
  const partialAt = new Map<string, number>();
  /** The steps a character went at (at the witness). */
  const wentAt = new Map<string, number[]>();
  const baseOf = new Map<string, number>();
  const placedIds = new Set<string>();
  witness = await joinDoc(net, 'witness', bytes, { instant: true, readOnly: true,
    before: (m) => { if (m.t === 'op' && typeof m.rev === 'number' && m.rev === rev + 1) B = snap(witness.doc); },
    after: (m) => {
      if (m.t !== 'op' || typeof m.rev !== 'number' || m.rev !== rev + 1 || !B) return;
      rev = m.rev;
      const A = snap(witness.doc);
      {
        const before = new Set(Array.from(B.text.values()).join('').split(''));
        const after = new Set(Array.from(A.text.values()).join('').split(''));
        before.forEach((ch) => { if (!after.has(ch) && isCh(ch)) wentAt.set(ch, (wentAt.get(ch) || []).concat([rev])); });
      }
      const it = intents.get(String(m.id));
      if (cfg.trace) { const ks = Array.from(A.text.keys()).filter((k) => A.text.get(k) !== B.text.get(k)); console.log(`rev ${rev} ${it ? it.what : '?'} ops ${JSON.stringify((m as unknown as { ops: unknown }).ops)}` + ks.map((k) => `\n     ${k}: ${JSON.stringify(B!.text.get(k))} -> ${JSON.stringify(A.text.get(k))}`).join('') + `\n     indents ${JSON.stringify(A.indent.get('body'))}`); }
      if (!it) { res.unknown++; B = undefined; return; }
      placedIds.add(String(m.id));
      const au = members.find((x) => it.what.indexOf(x.name + ' ') === 0);
      const where = `rev ${rev} [${it.kind}] AUTHORDROPS=${au ? au.session.stats.dropped : '?'} ${it.what}`;
      let changedKey: string | undefined;
      if (it.kind === 'text') changedKey = it.key;
      if (it.kind === 'undo') {
        const keys = Array.from(it.P.text.keys()).filter((k) => it.P.text.get(k) !== it.Q.text.get(k));
        if (keys.length > 1) return fail(`${where}: an undo changed ${keys.length} lists for its author`);
        changedKey = keys[0];
      }
      for (const key of Array.from(new Set([...B.text.keys(), ...A.text.keys()]))) {
        const b = B.text.get(key), a = A.text.get(key);
        if (b === undefined || a === undefined) { fail(`${where}: list ${key} came or went`); continue; }
        if (key !== changedKey) { if (a !== b) fail(`${where}: list ${key} changed: ${JSON.stringify(b)} -> ${JSON.stringify(a)}`); continue; }
        const P = it.P.text.get(key)!, Q = (it as { Q: Snap }).Q.text.get(key)!;
        const base = baseOf.get(String(m.id)) || 0;
        const loose = markRevs.some((x) => x.rev > base && x.key === key);
        { let x = 0; while (x < P.length && x < Q.length && P[x] === Q[x]) x++; let y = 0; while (y < P.length - x && y < Q.length - x && P[P.length - 1 - y] === Q[Q.length - 1 - y]) y++; if (count(b) !== count(a) || P.slice(x, P.length - y).indexOf('\n') >= 0 || Q.slice(x, Q.length - y).indexOf('\n') >= 0) markRevs.push({ rev, key }); }
        // What went after the author's step was made, and is here again: brought back by an undoing.
        const revived = new Set<string>();
        wentAt.forEach((revs, ch) => { if (revs.some((r) => r > base && r < rev) && b.indexOf(ch) >= 0) revived.add(ch); });
        const told = au ? au.notices.filter((n) => n.code === 'partial').length > (partialAt.has(String(m.id)) ? (partialAt.get(String(m.id)) as number) : 1e9) : false;
        const [err, exact] = textCheck(P, Q, b, a, loose, told ? revived : new Set());
        if (exact) res.exact++; else res.weak++;
        if (err) fail(`${where}: ${err}\n   P ${JSON.stringify(P)}\n   Q ${JSON.stringify(Q)}\n   B ${JSON.stringify(b)}\n   A ${JSON.stringify(a)}`);
      }
      // colours
      if (it.kind === 'color' || it.kind === 'text' || it.kind === 'indent') {
        const V = it.kind === 'color' ? new Set(it.V) : new Set<string>();
        A.color.forEach((col, ch) => {
          if (V.has(ch)) { if (col !== (it as { c: string }).c) fail(`${where}: character ${ch} was to be ${(it as { c: string }).c}, is ${col}`); }
          else if (B!.color.has(ch) && B!.color.get(ch) !== col) fail(`${where}: colour of ${ch} changed ${B!.color.get(ch)} -> ${col}`);
          else if (it.kind === 'color' && col === it.c) fail(`${where}: character ${ch} the author did not see got its colour`);
        });
      }
      if (it.kind === 'indent') {
        const text = A.text.get(it.key)!;
        const ind = A.indent.get(it.key)!;
        const hits = ind.map((v, i) => (v === it.v ? i : -1)).filter((i) => i >= 0);
        if (hits.length !== 1) fail(`${where}: ${hits.length} paragraphs got the indent`);
        else {
          const paras = text.split('\n');
          const Pt = it.P.text.get(it.key)!;
          let lo = 0, hi = paras.length - 1;
          const paraOf = (c: string): number => count(text.slice(0, text.indexOf(c)));
          for (let i = it.gap - 1; i >= 0; i--) if (isCh(Pt[i]) && text.indexOf(Pt[i]) >= 0) { lo = paraOf(Pt[i]); break; }
          for (let i = it.gap; i < Pt.length; i++) if (isCh(Pt[i]) && text.indexOf(Pt[i]) >= 0) { hi = paraOf(Pt[i]); break; }
          if (hits[0] < lo || hits[0] > hi) fail(`${where}: indent went to paragraph ${hits[0]}, meant one of ${lo}..${hi}\n   P ${JSON.stringify(Pt)} gap ${it.gap}\n   A ${JSON.stringify(text)}`);
          const bi = B.indent.get(it.key)!;
          for (let i = 0; i < ind.length; i++) if (i !== hits[0] && ind[i] !== bi[i]) fail(`${where}: indent of another paragraph changed`);
        }
      } else if (it.kind !== 'undo') {
        A.indent.forEach((ind, key) => { const bi = B!.indent.get(key); if (bi && A.text.get(key) === B!.text.get(key) && bi.join() !== ind.join()) fail(`${where}: paragraph indents changed in ${key}`); });
      }
      B = undefined;
    } });
  const members: DocMember[] = [];
  for (let i = 0; i < cfg.members; i++) members.push(await joinDoc(net, `m${i}`, bytes, {}));
  if (cfg.trace) for (const m of members) {
    const s = m.session as unknown as { unlink(step: { id: string; wires: unknown; effects: unknown }): void; stats: { quiet: number; dropped: number }; confirmed: number };
    const inner = s.unlink.bind(s);
    const s2 = m.session as unknown as { remote(raw: unknown[]): boolean; pendingSteps: { id: string; wires: unknown }[]; confirmed: number };
    const rem = s2.remote.bind(s2);
    const want = (process.env.TRACE_ID || '').split(',');
    if (process.env.TRACE_ID && m.name === process.env.TRACE_M) s2.remote = (raw) => { const out = rem(raw); console.log(`  REMOTE ${m.name} rev ${s2.confirmed + 1} ${JSON.stringify(raw)}: ` + s2.pendingSteps.filter((x) => want.indexOf(x.id.split(':')[1]) >= 0).map((x) => `${x.id.split(':')[1]}=${JSON.stringify(x.wires)}`).join(' ') + ` body ${JSON.stringify(snap(m.doc).text.get('body'))}`); return out; };
    s.unlink = (step) => { console.log(`  UNLINK ${m.name} step ${step.id} at confirmed ${s.confirmed} quiet ${s.stats.quiet} dropped ${s.stats.dropped}: wires ${JSON.stringify(step.wires)} effects ${JSON.stringify(step.effects)}`); inner(step); };
  }
  const lastId = (m: DocMember): string | undefined => { const q = (m.session as unknown as { pendingSteps: { id: string }[] }).pendingSteps; return q.length ? q[q.length - 1].id : undefined; };
  let unique = 0;
  const kindOf = new Map<string, string>();
  const act = (m: DocMember): void => {
    const roll = random();
    if (roll < cfg.undo && cfg.quiesce) { if (process.env.CARET) followCarets(); for (let i = 0; i < 20 && m.session.pending > 0; i++) net.drain(random); if (process.env.CARET) followCarets(); }
    const P = snap(m.doc);
    if (roll < cfg.undo && (m.session.pending === 0 || random() < 0.15)) {
      const redo = random() < 0.3;
      res.undoAsked++;
      const before = lastId(m);
      const sync = m.session.pending === 0;
      const done = redo ? m.session.redo() : m.session.undo();
      const id = lastId(m);
      if (cfg.trace) console.log(`  ACT ${m.name} ${redo ? 'redo' : 'undo'} sync ${sync} done ${!!done} id ${id} rev ${m.session.rev} pending ${m.session.pending}: ${JSON.stringify(P.text.get('body'))} -> ${JSON.stringify(snap(m.doc).text.get('body'))}`);
      if (done && id && id !== before) { res.undoDone++; intents.set(id, { kind: 'undo', P, Q: snap(m.doc), what: `${m.name} ${redo ? 'redo' : 'undo'}` }); kindOf.set(id, 'undo'); baseOf.set(id, m.session.rev); partialAt.set(id, m.notices.filter((n) => n.code === 'partial').length); }
      else if (sync) { const Q = snap(m.doc); P.text.forEach((t, k) => { if (Q.text.get(k) !== t) fail(`${m.name}: an undo that was not done changed ${k}`); }); }
      return;
    }
    const keys = Array.from(P.text.keys());
    const key = random() < 0.55 ? 'body' : keys[int(0, keys.length - 1)];
    const text = P.text.get(key)!;
    const limit = key === 'body' ? Math.max(0, text.indexOf(TBL) - 1) : text.length;
    const gap = int(0, limit);
    const range = (): [number, number] | undefined => { if (limit === 0) return undefined; const from = int(0, limit - 1); return [from, Math.min(limit, from + int(1, 4))]; };
    const typed = (): string => { let t = fresh(int(1, 3)); if (random() < 0.15) { const at = int(0, t.length); t = t.slice(0, at) + '\n' + t.slice(at); } return t; };
    let wires: DocWire[]; let intent: Intent; let kind: string;
    const r2 = cfg.fmt ? random() : random() * 0.78;
    if (r2 < 0.36) { const t = typed(); wires = [{ t: 'text', at: posOf(m.doc, key, gap), x: t }]; kind = 'ins'; }
    else if (r2 < 0.46) { wires = [{ t: 'split', at: posOf(m.doc, key, gap) }]; kind = 'split'; }
    else if (r2 < 0.7) { const r = range(); if (!r) return; wires = [{ t: 'del', at: posOf(m.doc, key, r[0]), to: posOf(m.doc, key, r[1]) }]; kind = 'del'; }
    else if (r2 < 0.78) { const r = range(); if (!r) return; wires = [{ t: 'text', at: posOf(m.doc, key, r[0]), to: posOf(m.doc, key, r[1]), x: typed() }]; kind = 'repl'; }
    else if (r2 < 0.9) {
      const r = range(); if (!r) return;
      const c = (0x100000 + ++unique * 7).toString(16).toUpperCase();
      wires = [{ t: 'run', at: posOf(m.doc, key, r[0]), to: posOf(m.doc, key, r[1]), f: { color: c } }];
      kind = 'color';
      intent = { kind: 'color', key, P, V: chars(text.slice(r[0], r[1])).split('').filter((x) => x), c, what: '' };
    } else {
      const v = 1000 + ++unique;
      const at = posOf(m.doc, key, gap);
      wires = [{ t: 'para', at, to: at, f: { indent: { left: v } } }];
      kind = 'indent';
      let s = gap; while (s > 0 && text[s - 1] !== '\n') s--;
      let e = gap; while (e < text.length && text[e] !== '\n') e++;
      intent = { kind: 'indent', key, P, gap, para: text.slice(s, e).split(''), v, what: '' };
    }
    const before = lastId(m);
    try { m.session.submit(wires); } catch (error) { return fail(`${m.name} could not make ${JSON.stringify(wires)} on ${JSON.stringify(text)}: ${String(error)}`); }
    const id = lastId(m);
    if (!id || id === before) { if (kind !== 'color' && kind !== 'indent') fail(`${m.name}: ${JSON.stringify(wires)} made no step`); return; }
    const what = `${m.name} ${JSON.stringify(wires)}`;
    if (kind === 'color' || kind === 'indent') { intent!.what = what; intents.set(id, intent!); }
    else {
      const Q = snap(m.doc);
      P.text.forEach((t, k) => { if (k !== key && Q.text.get(k) !== t) fail(`${what}: changed list ${k} locally`); });
      intents.set(id, { kind: 'text', key, P, Q, what });
    }
    kindOf.set(id, kind);
    baseOf.set(id, m.session.rev);
    partialAt.set(id, m.notices.filter((n) => n.code === 'partial').length);
    if (cfg.trace) console.log(`  ACT ${what} id ${id} at rev ${m.session.rev} pending ${m.session.pending}: ${JSON.stringify(P.text.get(key))} -> ${JSON.stringify(snap(m.doc).text.get(key))}`);
  };
  net.drain(random);
  const seenN = new Map<string, number>();
  const carets = new Map<string, { pos: { story: string; path: number[]; offset: number }; left: string; seen: number } | undefined>();
  const textAt = (m: DocMember, pos: { story: string; path: number[]; offset: number }): string | undefined => { try { const b = containerBlocks(m.doc, pos.story, pos.path.slice(0, -1))[pos.path[pos.path.length - 1]]; return b && b.type === 'paragraph' ? paragraphText(b) : undefined; } catch { return undefined; } };
  const pickCaret = (m: DocMember): void => {
    const c = contOf(m.doc); const one = c[int(0, c.length - 1)]; const rows = one.blocks.map((b, i) => (b.type === 'paragraph' && paragraphText(b).length > 0 ? i : -1)).filter((i) => i >= 0);
    if (rows.length === 0) { carets.set(m.name, undefined); return; }
    const row = rows[int(0, rows.length - 1)]; const t = paragraphText(one.blocks[row] as any); const offset = int(1, t.length);
    carets.set(m.name, { pos: { story: one.story, path: one.prefix.concat([row]), offset }, left: t[offset - 1], seen: m.changes.length });
  };
  const followCarets = (): void => {
    for (const m of members) {
      const c = carets.get(m.name);
      if (!c) { pickCaret(m); continue; }
      if (m.changes.length === c.seen) continue;
      let pos: { story: string; path: number[]; offset: number } | undefined = c.pos;
      for (let i = c.seen; i < m.changes.length && pos; i++) {
        const ch = m.changes[i];
        if (ch.reset || ch.local) { pos = undefined; break; }
        const next = movedSelection({ anchor: pos, focus: pos }, ch.marks);
        pos = next ? next.focus : undefined;
      }
      res.caretMoves++;
      if (!pos) { res.caretLost++; pickCaret(m); continue; }
      const all = Array.from(snap(m.doc).text.values()).join('');
      const t = textAt(m, pos);
      if (all.indexOf(c.left) >= 0) {
        if (t === undefined || t[pos.offset - 1] !== c.left) { fail(`${m.name}: the caret stood behind ${c.left} and is now at ${JSON.stringify(pos)} in ${JSON.stringify(t)} (marks ${JSON.stringify(m.changes.slice(c.seen).map((x) => x.marks)).slice(0, 400)})`); }
        else res.caretOk++;
      } else if (t === undefined || pos.offset > t.length) fail(`${m.name}: the caret is at no place: ${JSON.stringify(pos)}`);
      c.pos = pos; c.seen = m.changes.length;
      if (all.indexOf(c.left) < 0) pickCaret(m);
    }
  };
  for (let tick = 0; tick < cfg.ticks && bad.length === 0; tick++) {
    if (cfg.trace) for (const m of members) { const n = seenN.get(m.name) || 0; if (m.notices.length > n) { console.log(`  NOTICE ${m.name} at rev ${m.session.rev} pending ${m.session.pending}: ${JSON.stringify(m.notices.slice(n))} ids ${(m.session as unknown as { pendingSteps: { id: string }[] }).pendingSteps.map((x) => x.id.split(':')[1]).join(',')}`); seenN.set(m.name, m.notices.length); } }
    const roll = random();
    if (roll < cfg.deliver) { net.step(random); if (process.env.CARET) followCarets(); }
    else if (cfg.cuts && roll < cfg.deliver + 0.03) members[int(0, members.length - 1)].transport.suspend();
    else if (cfg.cuts && roll < cfg.deliver + 0.07) members[int(0, members.length - 1)].transport.resume();
    else { const m = members[int(0, members.length - 1)]; const burst = random() < 0.2 ? int(2, 6) : 1; for (let i = 0; i < burst && bad.length === 0; i++) { act(m); if (process.env.CARET) pickCaret(m); } if (process.env.CARET) followCarets(); }
  }
  if (bad.length === 0) {
    for (const m of members) m.transport.resume();
    for (let round = 0; round < 80; round++) { net.drain(random); await new Promise<void>((r) => setImmediate(r)); if (net.waiting === 0 && members.every((m) => m.session.pending === 0)) break; }
    const all = members.concat([witness]);
    const canon = all.map((m) => canonDoc(m.doc));
    for (let i = 0; i < all.length; i++) {
      if (i < members.length && all[i].session.pending !== 0) fail(`${all[i].name} still has ${all[i].session.pending} waiting (state ${all[i].session.state})`);
      if (canon[i] !== canon[all.length - 1]) fail(`${all[i].name} differs from the witness ${firstDifference(canon[all.length - 1], canon[i])}`);
    }
    // files
    const files: Uint8Array[] = [];
    for (const m of members.slice(0, 2)) { await writeDocx(m.file.package!, m.doc); files.push(await m.file.package!.save()); }
    if (files.length === 2 && Buffer.compare(Buffer.from(files[0]), Buffer.from(files[1])) !== 0) fail('two members write different files');
  }
  for (const m of members) {
    const st = m.session.stats;
    res.made += st.made; res.placed += st.placed; res.dropped += st.dropped; res.quiet += st.quiet;
    if (bad.length === 0 && st.made !== st.placed + st.dropped + st.quiet) fail(`${m.name}: made ${st.made} placed ${st.placed} dropped ${st.dropped} quiet ${st.quiet}`);
    for (const n of m.notices) res.notices[n.code] = (res.notices[n.code] || 0) + (n.count || 1);
    m.session.leave();
  }
  witness.session.leave();
  if (bad.length === 0) {
    // what never got a place
    const final = snap(witness.doc);
    const all = Array.from(final.text.values()).join('');
    intents.forEach((it, id) => {
      if (placedIds.has(id)) return;
      const k = kindOf.get(id) || '?';
      res.droppedKinds[k] = (res.droppedKinds[k] || 0) + 1;
      // (Refined: a colouring that went with someone else's deletion is said as `partial`, which counts as said.)
      if (it.kind === 'color' && it.V.some((c) => all.indexOf(c) >= 0) && res.dropped === 0 && !res.notices.partial) fail(`an unplaced colouring's text is still in the document and nothing was dropped: ${it.what}`);
      if (it.kind === 'text') {
        const key = it.key; const P = it.P.text.get(key)!, Q = it.Q.text.get(key)!;
        const added = chars(Q).split('').filter((c) => P.indexOf(c) < 0);
        if (added.some((c) => all.indexOf(c) >= 0)) fail(`an unplaced step's text is in the document: ${it.what}`);
        const gone = chars(P).split('').filter((c) => Q.indexOf(c) < 0);
        if (cfg.undo === 0 && k === 'del' && gone.some((c) => all.indexOf(c) >= 0)) fail(`an unplaced deletion's text is still in the document: ${it.what}`);
        if (cfg.undo === 0 && (k === 'ins' || k === 'split' || k === 'repl')) fail(`a step of kind ${k} never got a place: ${it.what}`);
      }
    });
    const unplaced = intents.size - placedIds.size;
    if (unplaced > res.dropped + res.quiet) fail(`${unplaced} steps never placed; dropped ${res.dropped} + quiet ${res.quiet}`);
    const said = (res.notices['dropped'] || 0) + (res.notices['undo-gone-pending'] || 0);
    void said;
  }
  return res;
}
