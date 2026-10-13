// Source text of a part and a cursor that follows the scanner through it.
//
// The scanner (XmlScanner) hands out tokens without saying where in the text
// they are, and passes over comments and processing instructions silently.
// The reader needs both: the place of every block, so that the model can point
// into the source instead of holding a copy of each block, and the comments
// and processing instructions between blocks, which the model keeps.
//
// SourceCursor walks the same text in step with the scanner. The scanner has
// already checked everything the cursor looks at, so the cursor only has to
// find its way, not to validate. Should the two ever disagree, the cursor
// gives up (`lost`) and the blocks read after that carry no source span; a
// writer then rebuilds them from the model.

import { decodeXmlBytes, type XmlComment, type XmlPI } from '../../xml';

/**
 * Decodes a part exactly as XmlScanner does (byte order mark dropped, line
 * ends as "\n"), so that offsets into the result are offsets into the text
 * the scanner reads. The scanner takes the string as it is.
 */
export function decodePart(bytes: Uint8Array): string {
  let text = decodeXmlBytes(bytes);
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  if (text.indexOf('\r') >= 0) text = text.replace(/\r\n?/g, '\n');
  return text;
}

const isSpace = (c: number) => c === 32 || c === 9 || c === 10 || c === 13;

export class SourceCursor {
  /** Offset the cursor is at. */
  pos = 0;
  /** The cursor could not follow the scanner; its position means nothing any more. */
  lost = false;

  constructor(readonly text: string) {}

  /**
   * Moves to the "<" of the next start or end tag. Comments and processing
   * instructions passed on the way are reported with their span; text and
   * CDATA sections are passed over.
   */
  toTag(onMisc?: (node: XmlComment | XmlPI, start: number, end: number) => void): void {
    if (this.lost) return;
    const text = this.text;
    let i = this.pos;
    for (;;) {
      i = text.indexOf('<', i);
      if (i < 0) break;
      const c = text.charCodeAt(i + 1);
      if (c === 33) {
        if (text.startsWith('<!--', i)) {
          const end = text.indexOf('-->', i + 4);
          if (end < 0) break;
          if (onMisc) onMisc({ type: 'comment', value: text.slice(i + 4, end) }, i, end + 3);
          i = end + 3;
          continue;
        }
        if (text.startsWith('<![CDATA[', i)) {
          const end = text.indexOf(']]>', i + 9);
          if (end < 0) break;
          i = end + 3;
          continue;
        }
        break;
      }
      if (c === 63) {
        const end = text.indexOf('?>', i + 2);
        if (end < 0) break;
        if (onMisc) {
          // Target and value as the tree parser splits them.
          let split = i + 2;
          while (split < end && !isSpace(text.charCodeAt(split))) split++;
          let value = split;
          while (value < end && isSpace(text.charCodeAt(value))) value++;
          onMisc({ type: 'pi', target: text.slice(i + 2, split), value: text.slice(value, end) }, i, end + 2);
        }
        i = end + 2;
        continue;
      }
      this.pos = i;
      return;
    }
    this.lost = true;
  }

  /** With the cursor on the "<" of a start tag: moves past the ">" that ends the tag. */
  enter(): void {
    if (this.lost) return;
    const text = this.text;
    const n = text.length;
    let i = this.pos + 1;
    while (i < n) {
      const c = text.charCodeAt(i);
      if (c === 62) {
        this.pos = i + 1;
        return;
      }
      if (c === 34 || c === 39) {
        const end = text.indexOf(c === 34 ? '"' : "'", i + 1);
        if (end < 0) break;
        i = end + 1;
      } else i++;
    }
    this.lost = true;
  }

  /** With the cursor on the "<" of an end tag: moves past it. */
  leave(): void {
    if (this.lost) return;
    const end = this.text.indexOf('>', this.pos);
    if (end < 0 || this.text.charCodeAt(this.pos + 1) !== 47) this.lost = true;
    else this.pos = end + 1;
  }

  /**
   * With the cursor on the "<" of an element whose source text the scanner
   * returned: moves past the element. False when the text is not what is at
   * the cursor.
   */
  skip(raw: string): boolean {
    if (this.lost) return false;
    const end = this.pos + raw.length;
    // The first and last character are compared, not the whole text: this runs for every block.
    if (raw.length < 2 || this.text.charCodeAt(this.pos) !== 60 || this.text.charCodeAt(end - 1) !== 62 || end > this.text.length) {
      this.lost = true;
      return false;
    }
    this.pos = end;
    return true;
  }
}
