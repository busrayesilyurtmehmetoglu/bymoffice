// Where the text of a part names a relationship by its id: the elements with an attribute of a
// namespace whose local name is "id" (r:id) and whose value is the id. Read from the text without
// a tree, so that a sheet of any size is looked through in one pass for each id (an id that does
// not occur costs one search).

/** An element of a part that names a relationship: its qualified name, and where its start tag begins and ends. */
export interface NamingElement {
  name: string;
  start: number;
  /** Just past the ">" of its start tag. */
  end: number;
  selfClosing: boolean;
}

/** Elements naming one id that are looked at at most (a part that names it more often is hostile, or broken). */
const MAX_NAMING = 10000;

const isSpace = (c: string) => c === ' ' || c === '\t' || c === '\n' || c === '\r';

/** The name of the attribute whose quoted value starts at `at`, read backwards from it (in linear time); undefined when none is there. */
function attributeBefore(text: string, at: number, lt: number): string | undefined {
  let i = at - 1;
  while (i > lt && isSpace(text.charAt(i))) i--;
  if (text.charAt(i) !== '=') return undefined;
  i--;
  while (i > lt && isSpace(text.charAt(i))) i--;
  const end = i + 1;
  while (i > lt && !isSpace(text.charAt(i)) && '=<>"\'/'.indexOf(text.charAt(i)) < 0) i--;
  return end > i + 1 ? text.slice(i + 1, end) : undefined;
}

/** The elements of `text` (a part) whose prefixed "id" attribute is `id`; an id inside text, a comment or another attribute does not count. */
export function elementsNaming(text: string, id: string): NamingElement[] {
  const out: NamingElement[] = [];
  if (id === '') return out;
  for (const quote of ['"', "'"]) {
    const needle = `${quote}${id}${quote}`;
    // The last "<" and ">" before each place, found by pointers that only move forward: one pass over the text.
    let lt = -1;
    let gt = -1;
    // The first "<" and ">" after the last ones found: searched for again only once passed.
    let nextLt = text.indexOf('<');
    let nextGt = text.indexOf('>');
    let tag = -1;
    let tagClose = -1;
    for (let at = text.indexOf(needle); at >= 0 && out.length < MAX_NAMING; at = text.indexOf(needle, at + needle.length)) {
      while (nextLt >= 0 && nextLt < at) {
        lt = nextLt;
        nextLt = text.indexOf('<', nextLt + 1);
      }
      while (nextGt >= 0 && nextGt < at) {
        gt = nextGt;
        nextGt = text.indexOf('>', nextGt + 1);
      }
      // Not inside a tag: text of an element.
      if (lt < 0 || gt > lt) continue;
      const head = text.charAt(lt + 1);
      if (head === '!' || head === '?' || head === '/') continue;
      const attribute = attributeBefore(text, at, lt);
      if (attribute === undefined) continue;
      const colon = attribute.indexOf(':');
      if (colon <= 0 || attribute.slice(colon + 1) !== 'id' || attribute.slice(0, colon) === 'xmlns') continue;
      if (tag !== lt) {
        tag = lt;
        tagClose = text.indexOf('>', at + needle.length);
      }
      if (tagClose < 0) break;
      let n = lt + 1;
      while (n < text.length && !isSpace(text.charAt(n)) && text.charAt(n) !== '/' && text.charAt(n) !== '>') n++;
      out.push({ name: text.slice(lt + 1, n), start: lt, end: tagClose + 1, selfClosing: text.charAt(tagClose - 1) === '/' });
    }
  }
  return out.sort((a, b) => a.start - b.start);
}

/** The local name of a qualified name. */
export const localOf = (name: string): string => name.slice(name.indexOf(':') + 1);

/**
 * The text without the <pivotSelection> elements that name one of `ids` (the relationships of
 * pivot tables that go); undefined when there is none. Such an element does not hold another one.
 */
export function withoutPivotSelections(text: string, ids: Iterable<string>): string | undefined {
  const cuts: [number, number][] = [];
  for (const id of ids) {
    for (const el of elementsNaming(text, id)) {
      if (localOf(el.name) !== 'pivotSelection') continue;
      let end = el.end;
      if (!el.selfClosing) {
        const closing = text.indexOf(`</${el.name}`, el.end);
        if (closing < 0) continue;
        const gt = text.indexOf('>', closing);
        if (gt < 0) continue;
        end = gt + 1;
      }
      cuts.push([el.start, end]);
    }
  }
  if (cuts.length === 0) return undefined;
  cuts.sort((a, b) => a[0] - b[0]);
  let out = '';
  let from = 0;
  for (const [start, end] of cuts) {
    if (start < from) continue;
    out += text.slice(from, start);
    from = end;
  }
  return out + text.slice(from);
}
