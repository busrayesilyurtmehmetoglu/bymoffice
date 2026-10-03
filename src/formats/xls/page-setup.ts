// Page setup records of a sheet -> SheetPageSetup of the model.
//
// A sheet states how it is printed in a run of small records: margins, SETUP
// (paper, scaling, orientation), headers and footers, what is printed beside
// the cells, and the page breaks. Each record counts once: the first one that
// can be read. A record that is too short or holds a value that means nothing
// is left out, never a reason to refuse the sheet. The data of the printer
// driver (PLS) is not read.

import { MAX_COLS, MAX_ROWS, type HeaderFooter, type PageBreak, type PageSettings, type PrintOptions, type SheetPageSetup } from '../../model/sheet/types';
import {
  REC_BOTTOMMARGIN,
  REC_FOOTER,
  REC_GRIDSET,
  REC_HCENTER,
  REC_HEADER,
  REC_HEADERFOOTER,
  REC_HORIZONTALPAGEBREAKS,
  REC_LEFTMARGIN,
  REC_PRINTGRIDLINES,
  REC_PRINTHEADERS,
  REC_RIGHTMARGIN,
  REC_SETUP,
  REC_TOPMARGIN,
  REC_VCENTER,
  REC_VERTICALPAGEBREAKS,
  readerAt,
  type RecordStream,
} from './records';
import { readByteString, readString16 } from './strings';

/** Last row and column of a sheet of the format, which a break across the whole sheet ends at. */
const XLS_LAST_ROW = 65535;
const XLS_LAST_COL = 255;

/** The margins a sheet has where it states none: those of Excel 97-2003. */
const DEFAULT_MARGINS = { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 };

const SETUP_OVER_THEN_DOWN = 0x0001;
const SETUP_PORTRAIT = 0x0002;
/** Paper size, scale, resolution, copies and orientation were never set: they are undefined. */
const SETUP_NO_PRINTER_SETTINGS = 0x0004;
const SETUP_BLACK_AND_WHITE = 0x0008;
const SETUP_DRAFT = 0x0010;
const SETUP_COMMENTS = 0x0020;
/** The orientation was never set. */
const SETUP_NO_ORIENTATION = 0x0040;
const SETUP_USE_FIRST_PAGE = 0x0080;
const SETUP_COMMENTS_AT_END = 0x0200;

const ERRORS: readonly PageSettings['errors'][] = [undefined, 'blank', 'dash', 'NA'];

const WSBOOL_AUTO_BREAKS = 0x0001;
const WSBOOL_FIT_TO_PAGE = 0x0100;

/** A margin a page can have, in inches. */
const isMargin = (value: number) => value >= 0 && value < Infinity;

/**
 * A resolution as an xlsx file states it. The record holds 16 bits; a negative value is a
 * quality level of the printer driver (-1 draft .. -4 high), which the newer format writes as
 * the same number in 32 bits.
 */
const resolution = (value: number) => (value >= 0x8000 ? value + 0xffff0000 : value);

function sortBreaks(breaks: PageBreak[]): PageBreak[] {
  breaks.sort((a, b) => a.at - b.at);
  // One break per row or column: the first one stated.
  return breaks.filter((item, index) => index === 0 || item.at !== breaks[index - 1].at);
}

/** Collects the page setup records of one sheet. */
export class XlsPageSetup {
  /** Records that were read; a second one of the same kind is not. */
  private readonly seen = new Set<number>();
  private left: number | undefined;
  private right: number | undefined;
  private top: number | undefined;
  private bottom: number | undefined;
  private header: number | undefined;
  private footer: number | undefined;
  private page: PageSettings | undefined;
  private readonly print: PrintOptions = {};
  private readonly texts: HeaderFooter = {};
  private fitToPage = false;
  private autoPageBreaks = true;
  private rowBreaks: PageBreak[] | undefined;
  private colBreaks: PageBreak[] | undefined;

  constructor(
    private readonly stream: RecordStream,
    private readonly biff8: boolean,
    private readonly codepage: number,
  ) {}

  /**
   * Reads the record that starts at `offset`. False when nothing was taken from it: it is
   * not a page setup record, one of its kind was read before, or it cannot be read.
   */
  private optionsSeen = false;

  read(id: number, offset: number, length: number): boolean {
    if (this.seen.has(id)) return false;
    let taken: boolean;
    try {
      taken = this.record(id, offset, length);
    } catch {
      // A text that runs past its record.
      taken = false;
    }
    if (taken) this.seen.add(id);
    return taken;
  }

  /** The option flags of the sheet (WSBOOL), two of which are about printing. */
  options(flags: number): void {
    // The first record counts, as for the others.
    if (this.optionsSeen) return;
    this.optionsSeen = true;
    this.fitToPage = (flags & WSBOOL_FIT_TO_PAGE) !== 0;
    this.autoPageBreaks = (flags & WSBOOL_AUTO_BREAKS) !== 0;
  }

  private record(id: number, offset: number, length: number): boolean {
    const view = this.stream.view;
    const d = offset + 4;
    switch (id) {
      case REC_LEFTMARGIN:
      case REC_RIGHTMARGIN:
      case REC_TOPMARGIN:
      case REC_BOTTOMMARGIN: {
        if (length < 8) return false;
        const value = view.getFloat64(d, true);
        if (!isMargin(value)) return false;
        if (id === REC_LEFTMARGIN) this.left = value;
        else if (id === REC_RIGHTMARGIN) this.right = value;
        else if (id === REC_TOPMARGIN) this.top = value;
        else this.bottom = value;
        return true;
      }
      case REC_HEADER:
      case REC_FOOTER: {
        // An empty record: the sheet has none.
        const text = length === 0 ? '' : this.text(offset);
        if (text !== '') {
          if (id === REC_HEADER) this.texts.oddHeader = text;
          else this.texts.oddFooter = text;
        }
        return true;
      }
      case REC_HCENTER:
      case REC_VCENTER:
      case REC_PRINTHEADERS:
      case REC_PRINTGRIDLINES:
      case REC_GRIDSET: {
        if (length < 2) return false;
        const set = view.getUint16(d, true) !== 0;
        // Only what differs from the default is a statement.
        if (id === REC_GRIDSET) {
          // Says whether the setting was ever changed; whether the lines are printed is PRINTGRIDLINES alone.
          // In an xlsx the lines are printed only when both attributes are true, so this is not carried over:
          // a file that prints its lines and says 0 here would stop printing them.
        } else if (set) {
          if (id === REC_HCENTER) this.print.horizontalCentered = true;
          else if (id === REC_VCENTER) this.print.verticalCentered = true;
          else if (id === REC_PRINTHEADERS) this.print.headings = true;
          else this.print.gridLines = true;
        }
        return true;
      }
      case REC_SETUP:
        return this.setup(d, length);
      case REC_HORIZONTALPAGEBREAKS:
      case REC_VERTICALPAGEBREAKS:
        return this.breaks(id === REC_HORIZONTALPAGEBREAKS, d, length);
      case REC_HEADERFOOTER:
        return this.biff8 && this.headerFooter(offset, length);
      default:
        return false;
    }
  }

  /** The text of a HEADER or FOOTER record: a 16-bit string in BIFF8, a byte string before. */
  private text(offset: number): string {
    const r = readerAt(this.stream, offset);
    return this.biff8 ? readString16(r).text : readByteString(r, r.u8(), this.codepage);
  }

  private setup(d: number, length: number): boolean {
    // The flags say which of the other fields count.
    if (length < 12) return false;
    const view = this.stream.view;
    const flags = view.getUint16(d + 10, true);
    const page: PageSettings = {};
    if (!(flags & SETUP_NO_PRINTER_SETTINGS)) {
      const paper = view.getUint16(d, true);
      if (paper > 0) page.paperSize = paper;
      const scale = view.getUint16(d + 2, true);
      if (scale >= 10 && scale <= 400) page.scale = scale;
      if (!(flags & SETUP_NO_ORIENTATION)) page.orientation = flags & SETUP_PORTRAIT ? 'portrait' : 'landscape';
      if (length >= 16) {
        page.horizontalDpi = resolution(view.getUint16(d + 12, true));
        page.verticalDpi = resolution(view.getUint16(d + 14, true));
      }
      if (length >= 34) {
        const copies = view.getUint16(d + 32, true);
        if (copies > 0) page.copies = copies;
      }
    }
    // 1 by 1 is the default; 0 is "as many pages as needed".
    const fitWidth = view.getUint16(d + 6, true);
    const fitHeight = view.getUint16(d + 8, true);
    if (fitWidth !== 1 && fitWidth <= 32767) page.fitToWidth = fitWidth;
    if (fitHeight !== 1 && fitHeight <= 32767) page.fitToHeight = fitHeight;
    if (flags & SETUP_USE_FIRST_PAGE) {
      // The number means nothing without the flag.
      page.firstPageNumber = view.getInt16(d + 4, true);
      page.useFirstPageNumber = true;
    }
    if (flags & SETUP_OVER_THEN_DOWN) page.pageOrder = 'overThenDown';
    if (flags & SETUP_BLACK_AND_WHITE) page.blackAndWhite = true;
    if (flags & SETUP_DRAFT) page.draft = true;
    // BIFF5 has neither the place of the comments nor the way errors are printed.
    if (flags & SETUP_COMMENTS) page.cellComments = this.biff8 && flags & SETUP_COMMENTS_AT_END ? 'atEnd' : 'asDisplayed';
    // UNSURE: by the specification 1 is "errors are printed blank", yet every file of a large set of real
    // workbooks (three hundred sheets from several templates, written by Excel 2007 and later) has exactly
    // that value, and the xlsx files of the same templates state no such setting. Until it is checked in
    // Excel, 1 is taken as nothing said: wrongly printing errors is the smaller harm than wrongly hiding them.
    const errors = this.biff8 ? ERRORS[(flags >> 10) & 3] : undefined;
    if (errors && errors !== 'blank') page.errors = errors;
    this.page = page;

    if (length >= 24) {
      const header = view.getFloat64(d + 16, true);
      if (isMargin(header)) this.header = header;
    }
    if (length >= 32) {
      const footer = view.getFloat64(d + 24, true);
      if (isMargin(footer)) this.footer = footer;
    }
    return true;
  }

  /**
   * HORIZONTALPAGEBREAKS (between rows) and VERTICALPAGEBREAKS (between columns). BIFF8 states
   * the columns or rows each break spans; BIFF5 states only where it is. The count in the
   * record is not trusted beyond what the record holds.
   */
  private breaks(rows: boolean, d: number, length: number): boolean {
    if (length < 2) return false;
    const view = this.stream.view;
    const size = this.biff8 ? 6 : 2;
    const count = view.getUint16(d, true);
    const end = d + length;
    const last = rows ? XLS_LAST_ROW : XLS_LAST_COL;
    const crossLast = rows ? XLS_LAST_COL : XLS_LAST_ROW;
    const whole = (rows ? MAX_COLS : MAX_ROWS) - 1;
    const list: PageBreak[] = [];
    for (let i = 0, q = d + 2; i < count && q + size <= end; i++, q += size) {
      const at = view.getUint16(q, true);
      // A break above the first row or left of the first column is none.
      if (at < 1 || at > last) continue;
      let min = 0;
      let max = whole;
      if (this.biff8) {
        min = view.getUint16(q + 2, true);
        const stated = view.getUint16(q + 4, true);
        if (min > stated || min > crossLast) continue;
        // A break across the whole of the old sheet is one across the whole sheet.
        max = stated >= crossLast ? whole : stated;
      }
      list.push({ at, min, max, manual: true });
    }
    if (rows) this.rowBreaks = sortBreaks(list);
    else this.colBreaks = sortBreaks(list);
    return true;
  }

  /**
   * HEADERFOOTER (BIFF8, Excel 2007 and later): the headers and footers of even pages and of
   * the first page, and four flags. One that belongs to a custom view (it names the view) is
   * not the sheet's.
   */
  private headerFooter(offset: number, length: number): boolean {
    if (length < 38) return false;
    const { bytes, view } = this.stream;
    const d = offset + 4;
    for (let i = 12; i < 28; i++) if (bytes[d + i] !== 0) return false;
    const flags = view.getUint16(d + 28, true);
    const counts = [view.getUint16(d + 30, true), view.getUint16(d + 32, true), view.getUint16(d + 34, true), view.getUint16(d + 36, true)];
    const r = readerAt(this.stream, offset);
    r.skip(38);
    // Read before anything is kept: a text that is cut off leaves the record out as a whole.
    const texts = counts.map((count) => (count > 0 ? readString16(r).text : ''));
    const out = this.texts;
    if (texts[0] !== '') out.evenHeader = texts[0];
    if (texts[1] !== '') out.evenFooter = texts[1];
    if (texts[2] !== '') out.firstHeader = texts[2];
    if (texts[3] !== '') out.firstFooter = texts[3];
    if (flags & 0x01) out.differentOddEven = true;
    if (flags & 0x02) out.differentFirst = true;
    if (!(flags & 0x04)) out.scaleWithDoc = false;
    if (!(flags & 0x08)) out.alignWithMargins = false;
    return true;
  }

  /** What the records said; undefined for a sheet that says nothing about printing. */
  finish(): SheetPageSetup | undefined {
    const out: SheetPageSetup = {};
    let any = false;
    if (this.left !== undefined || this.right !== undefined || this.top !== undefined || this.bottom !== undefined || this.header !== undefined || this.footer !== undefined) {
      // A margin without a record is the default one.
      out.margins = {
        left: this.left !== undefined ? this.left : DEFAULT_MARGINS.left,
        right: this.right !== undefined ? this.right : DEFAULT_MARGINS.right,
        top: this.top !== undefined ? this.top : DEFAULT_MARGINS.top,
        bottom: this.bottom !== undefined ? this.bottom : DEFAULT_MARGINS.bottom,
        header: this.header !== undefined ? this.header : DEFAULT_MARGINS.header,
        footer: this.footer !== undefined ? this.footer : DEFAULT_MARGINS.footer,
      };
      any = true;
    }
    if (this.page) {
      out.page = this.page;
      any = true;
    }
    if (Object.keys(this.print).length > 0) {
      out.print = this.print;
      any = true;
    }
    if (Object.keys(this.texts).length > 0) {
      out.headerFooter = this.texts;
      any = true;
    }
    if (this.fitToPage) {
      out.fitToPage = true;
      any = true;
    }
    if (!this.autoPageBreaks) {
      out.autoPageBreaks = false;
      any = true;
    }
    if (this.rowBreaks && this.rowBreaks.length > 0) {
      out.rowBreaks = this.rowBreaks;
      any = true;
    }
    if (this.colBreaks && this.colBreaks.length > 0) {
      out.colBreaks = this.colBreaks;
      any = true;
    }
    return any ? out : undefined;
  }
}
