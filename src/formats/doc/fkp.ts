// Formatted disk pages (FKP, MS-DOC 2.9.33 and 2.9.174): which stretch of the
// WordDocument stream has which character or paragraph properties.
//
// A bin table (PlcBteChpx, PlcBtePapx) lists the pages; a page is 512 bytes
// with file offsets at its start, property lists packed from its end and the
// number of runs in its last byte. All pages are read once into flat arrays
// sorted by offset, so that finding the properties of a character is one
// binary search.

import { readPlc, u16, u32, u8 } from './bytes';

const PAGE = 512;

export interface RunTable {
  /** Number of runs. */
  count: number;
  /** Offsets in the WordDocument stream: run i covers [starts[i], ends[i]). */
  starts: number[];
  ends: number[];
  /** Where the property list (grpprl) of run i is in the WordDocument stream, and its length; length 0 when it has none. */
  at: number[];
  length: number[];
  /** Paragraph runs only: the style of the paragraph. */
  istd: number[];
  /** Pages or runs that pointed outside their page or the stream and were left out. */
  broken: number;
  /** Index of the run that holds the stream offset; -1 when none does. */
  find(fc: number): number;
  /** Where the first run that starts after the stream offset starts; Infinity when there is none. */
  after(fc: number): number;
}

function finish(table: Omit<RunTable, 'find' | 'after' | 'count'>): RunTable {
  const { starts } = table;
  let sorted = true;
  for (let i = 1; i < starts.length; i++) {
    if (starts[i] < starts[i - 1]) {
      sorted = false;
      break;
    }
  }
  if (!sorted) {
    const order = starts.map((_, i) => i).sort((a, b) => starts[a] - starts[b]);
    const pick = (values: number[]) => order.map((i) => values[i]);
    table = { starts: pick(starts), ends: pick(table.ends), at: pick(table.at), length: pick(table.length), istd: table.istd.length ? pick(table.istd) : [], broken: table.broken };
  }
  const s = table.starts;
  const e = table.ends;
  return {
    ...table,
    count: s.length,
    find(fc: number): number {
      let lo = 0;
      let hi = s.length - 1;
      let found = -1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        if (s[mid] <= fc) {
          found = mid;
          lo = mid + 1;
        } else hi = mid - 1;
      }
      // Runs may overlap in a damaged file; the one that starts last is taken, as long as it reaches the offset.
      return found >= 0 && fc < e[found] ? found : -1;
    },
    after(fc: number): number {
      let lo = 0;
      let hi = s.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (s[mid] <= fc) lo = mid + 1;
        else hi = mid;
      }
      return lo < s.length ? s[lo] : Infinity;
    },
  };
}

/**
 * Calls `each` for every page of a bin table with the stretch of the stream
 * the table gives the page: [from, to). A page may hold runs outside of its
 * stretch (what an earlier save left there); those belong to other pages.
 */
function pages(word: Uint8Array, table: Uint8Array, fc: number, lcb: number, each: (page: number, from: number, to: number) => void): number {
  const plc = readPlc(table, fc, lcb, 4);
  let broken = 0;
  // A page cannot be listed more often than the stream has pages without the file lying; twice that is allowed for.
  const limit = Math.min(plc.count, Math.ceil(word.length / PAGE) * 2 + 2);
  for (let i = 0; i < limit; i++) {
    const page = (u32(table, plc.data(i)) & 0x3fffff) * PAGE;
    if (page + PAGE > word.length) broken++;
    else each(page, plc.cps[i], plc.cps[i + 1]);
  }
  return broken;
}

/** The character property runs of the file. */
export function readChpxTable(word: Uint8Array, table: Uint8Array, fc: number, lcb: number): RunTable {
  const out = { starts: [] as number[], ends: [] as number[], at: [] as number[], length: [] as number[], istd: [] as number[], broken: 0 };
  out.broken = pages(word, table, fc, lcb, (page, from, to) => {
    const runs = word[page + PAGE - 1];
    // The offsets and the one-byte places of the lists must fit into the page.
    if ((runs + 1) * 4 + runs > PAGE - 1) {
      out.broken++;
      return;
    }
    const places = page + (runs + 1) * 4;
    for (let i = 0; i < runs; i++) {
      const start = Math.max(from, u32(word, page + i * 4));
      const end = Math.min(to, u32(word, page + (i + 1) * 4));
      if (end <= start) continue;
      const place = word[places + i] * 2;
      let at = 0;
      let length = 0;
      if (place !== 0) {
        length = word[page + place];
        at = page + place + 1;
        if (place + 1 + length > PAGE) {
          out.broken++;
          length = 0;
        }
      }
      out.starts.push(start);
      out.ends.push(end);
      out.at.push(at);
      out.length.push(length);
    }
  });
  return finish(out);
}

/** The paragraph property runs of the file. */
export function readPapxTable(word: Uint8Array, table: Uint8Array, fc: number, lcb: number): RunTable {
  const out = { starts: [] as number[], ends: [] as number[], at: [] as number[], length: [] as number[], istd: [] as number[], broken: 0 };
  out.broken = pages(word, table, fc, lcb, (page, from, to) => {
    const runs = word[page + PAGE - 1];
    // Offsets, then a BxPap of 13 bytes for each run.
    if ((runs + 1) * 4 + runs * 13 > PAGE - 1) {
      out.broken++;
      return;
    }
    const places = page + (runs + 1) * 4;
    for (let i = 0; i < runs; i++) {
      const start = Math.max(from, u32(word, page + i * 4));
      const end = Math.min(to, u32(word, page + (i + 1) * 4));
      if (end <= start) continue;
      const place = word[places + i * 13] * 2;
      let at = 0;
      let length = 0;
      let istd = 0;
      if (place !== 0) {
        // PapxInFkp: a size in words, or 0 and then the size in the next byte.
        let size = u8(word, page + place) * 2 - 1;
        let first = place + 1;
        if (size < 0) {
          size = u8(word, page + place + 1) * 2;
          first = place + 2;
        }
        if (size >= 2 && first + size <= PAGE) {
          istd = u16(word, page + first);
          at = page + first + 2;
          length = size - 2;
        } else if (size !== 0) out.broken++;
      }
      out.starts.push(start);
      out.ends.push(end);
      out.at.push(at);
      out.length.push(length);
      out.istd.push(istd);
    }
  });
  return finish(out);
}
