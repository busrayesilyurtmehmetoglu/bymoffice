// Clipboard text: rows separated by line breaks, cells by tabs, the way
// spreadsheets exchange ranges. A field with a tab, a line break or a quote is
// wrapped in quotes, with quotes inside it doubled.

/** Cells beyond this are dropped from pasted text, so a huge clipboard cannot stall the page. */
export const MAX_PASTE_CELLS = 1000000;

export function toTsv(grid: string[][]): string {
  return grid
    .map((row) => row.map((field) => (/[\t\n\r"]/.test(field) ? `"${field.replace(/"/g, '""')}"` : field)).join('\t'))
    .join('\n');
}

/** Parses clipboard text into rows of fields. Rows may differ in length. Text beyond MAX_PASTE_CELLS is dropped. */
export function parseTsv(text: string): string[][] {
  return parseTsvLimited(text).rows;
}

/** As parseTsv(), and says whether the text held more cells than are read. */
export function parseTsvLimited(text: string): { rows: string[][]; truncated: boolean } {
  const rows: string[][] = [];
  let row: string[] = [];
  let cells = 0;
  const n = text.length;
  let i = 0;
  const endField = (field: string) => {
    row.push(field);
    cells++;
  };
  while (i <= n && cells < MAX_PASTE_CELLS) {
    let field: string;
    if (text.charCodeAt(i) === 34) {
      // Quoted: runs to the closing quote; "" inside is one quote.
      let j = i + 1;
      let value = '';
      for (;;) {
        const quote = text.indexOf('"', j);
        if (quote < 0) {
          value += text.slice(j);
          j = n;
          break;
        }
        value += text.slice(j, quote);
        if (text.charCodeAt(quote + 1) === 34) {
          value += '"';
          j = quote + 2;
        } else {
          j = quote + 1;
          break;
        }
      }
      // Anything between the closing quote and the separator stays part of the field.
      let end = j;
      while (end < n && text.charCodeAt(end) !== 9 && text.charCodeAt(end) !== 10 && text.charCodeAt(end) !== 13) end++;
      field = value + text.slice(j, end);
      i = end;
    } else {
      let end = i;
      while (end < n && text.charCodeAt(end) !== 9 && text.charCodeAt(end) !== 10 && text.charCodeAt(end) !== 13) end++;
      field = text.slice(i, end);
      i = end;
    }
    endField(field);
    if (i >= n) break;
    const separator = text.charCodeAt(i);
    i++;
    if (separator === 9) continue;
    if (separator === 13 && text.charCodeAt(i) === 10) i++;
    rows.push(row);
    row = [];
    // A line break at the very end closes the last row instead of starting an empty one.
    if (i >= n) return { rows, truncated: false };
  }
  rows.push(row);
  return { rows, truncated: i < n };
}

const HTML_ESCAPES: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' };

/** The same grid as an HTML table, which is what other spreadsheets prefer to paste. */
export function toHtmlTable(grid: string[][]): string {
  const cell = (text: string) => `<td>${text.replace(/[&<>"]/g, (ch) => HTML_ESCAPES[ch]).replace(/\r?\n/g, '<br>')}</td>`;
  return `<table>${grid.map((row) => `<tr>${row.map(cell).join('')}</tr>`).join('')}</table>`;
}
