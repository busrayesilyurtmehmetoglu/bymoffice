// A chart on the clipboard. The chart itself is not put there: it is held by the view that
// copied it, and the clipboard gets a word in its place and HTML that carries a mark. A paste is
// the chart that is held only when its HTML carries that very mark: text is never taken for a
// chart, whatever it says (the same word comes from anywhere).

const MARK = /<[^<>]*\sdata-bym-chart\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

/** What the clipboard gets for a chart: `word` stands for it, `mark` names the copy. Nothing of a file is in either. */
export function chartClip(word: string, mark: string): { text: string; html: string } {
  return { text: word, html: `<span data-bym-chart="${mark.replace(/[^A-Za-z0-9]/g, '')}">${word.replace(/[&<>"']/g, '')}</span>` };
}

/**
 * What the HTML of a paste is about: "chart" when it carries the mark of the chart that is held
 * (`held`); "stale" when it carries the mark of another copy of a chart (one of another view, or
 * of this page before it was loaded again): no chart can be made of that, and the word that
 * stands for it is no text to paste; "other" for anything else, HTML without a mark and no HTML
 * at all included.
 */
export function chartPasteKind(html: string | undefined, held: string | undefined): 'chart' | 'stale' | 'other' {
  if (typeof html !== 'string' || html.indexOf('data-bym-chart') < 0) return 'other';
  let found = false;
  MARK.lastIndex = 0;
  for (let match = MARK.exec(html), n = 0; match && n < 64; match = MARK.exec(html), n++) {
    found = true;
    const mark = match[1] !== undefined ? match[1] : match[2];
    if (held !== undefined && held !== '' && mark === held) return 'chart';
  }
  return found ? 'stale' : 'other';
}
