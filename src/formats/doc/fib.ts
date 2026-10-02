// The File Information Block: the header of the WordDocument stream
// (MS-DOC 2.5). It says which version wrote the file, which table stream
// belongs to it, how long each part of the text is, and where every other
// structure lies in the table stream (the fc/lcb pairs).

import { UnsupportedDocError, invalid, u16, u32 } from './bytes';

/** Indexes into the fc/lcb pairs of the FIB (FibRgFcLcb97 and later, MS-DOC 2.5.6). */
export const FC = {
  Stshf: 1,
  PlcffndRef: 2,
  PlcffndTxt: 3,
  PlcfandRef: 4,
  PlcfandTxt: 5,
  PlcfSed: 6,
  PlcfHdd: 11,
  PlcfBteChpx: 12,
  PlcfBtePapx: 13,
  SttbfFfn: 15,
  PlcfFldMom: 16,
  PlcfFldHdr: 17,
  PlcfFldFtn: 18,
  PlcfFldAtn: 19,
  SttbfBkmk: 21,
  PlcfBkf: 22,
  PlcfBkl: 23,
  Dop: 31,
  Clx: 33,
  GrpXstAtnOwners: 36,
  PlcSpaMom: 40,
  PlcSpaHdr: 41,
  PlcfendRef: 46,
  PlcfendTxt: 47,
  PlcfFldEdn: 48,
  DggInfo: 50,
  SttbfRMark: 51,
  PlcftxbxTxt: 56,
  PlcfFldTxbx: 57,
  PlcfHdrtxbxTxt: 58,
  PlcffldHdrTxbx: 59,
  PlfLst: 73,
  PlfLfo: 74,
} as const;

export interface Fib {
  nFib: number;
  /** Version from the FibRgCswNew, when the file has one (Word 2000 and later). */
  nFibNew: number;
  /** Language the application that wrote the file was installed in. */
  lid: number;
  /** The file is a template. */
  dot: boolean;
  /** The text was saved incrementally: pieces carry property modifiers. */
  complex: boolean;
  /** "0Table" or "1Table". */
  tableStream: string;
  /** Lengths of the parts of the text, in characters, in the order they follow each other. */
  ccpText: number;
  ccpFtn: number;
  ccpHdd: number;
  ccpMcr: number;
  ccpAtn: number;
  ccpEdn: number;
  ccpTxbx: number;
  ccpHdrTxbx: number;
  /** Offset and length of structure `index` (see FC) in the table stream; [0, 0] when the file has none. */
  fcLcb(index: number): [fc: number, lcb: number];
}

const WORD_IDENT = 0xa5ec;
/** nFib of Word 97; everything before it has another FIB layout and 8-bit text in a code page. */
const NFIB_WORD97 = 0x00c1;
/** Where the lengths of the text parts start (FibRgLw97.ccpText). */
const CCP = 76;
const FC_LCB = 154;

/** Reads the FIB. Throws UnsupportedDocError for files of Word 6/95 and for encrypted files. */
export function readFib(word: Uint8Array): Fib {
  if (word.length < 34) throw invalid('the WordDocument stream is too short');
  const wIdent = u16(word, 0);
  const nFib = u16(word, 2);
  if (wIdent !== WORD_IDENT) {
    // 0xA5DC is Word 6/95, 0xA59B/0xA5DB Word 2 and earlier.
    if (wIdent === 0xa5dc || wIdent === 0xa59b || wIdent === 0xa5db || wIdent === 0xa59c) {
      throw new UnsupportedDocError('Documents of Word 95 and earlier are not supported');
    }
    throw invalid('the WordDocument stream does not start with a Word header');
  }
  if (nFib < NFIB_WORD97) throw new UnsupportedDocError('Documents of Word 95 and earlier are not supported');
  const flags = u16(word, 10);
  if (flags & 0x0100) {
    // fEncrypted; with fObfuscated the "encryption" is the XOR scheme, which is not read either.
    throw new UnsupportedDocError('Password-protected documents are not supported yet');
  }
  if (word.length < FC_LCB) throw invalid('the file header is cut short');
  const pairs = u16(word, FC_LCB - 2);
  // The pairs the file says it has, as far as they are in the stream.
  const available = Math.max(0, Math.min(pairs, Math.floor((word.length - FC_LCB) / 8)));
  let nFibNew = 0;
  const cswNewAt = FC_LCB + pairs * 8;
  if (cswNewAt + 4 <= word.length && u16(word, cswNewAt) > 0) nFibNew = u16(word, cswNewAt + 2);
  return {
    nFib,
    nFibNew,
    lid: u16(word, 6),
    dot: (flags & 0x0001) !== 0,
    complex: (flags & 0x0004) !== 0,
    tableStream: flags & 0x0200 ? '1Table' : '0Table',
    ccpText: u32(word, CCP),
    ccpFtn: u32(word, CCP + 4),
    ccpHdd: u32(word, CCP + 8),
    ccpMcr: u32(word, CCP + 12),
    ccpAtn: u32(word, CCP + 16),
    ccpEdn: u32(word, CCP + 20),
    ccpTxbx: u32(word, CCP + 24),
    ccpHdrTxbx: u32(word, CCP + 28),
    fcLcb(index: number): [number, number] {
      if (index < 0 || index >= available) return [0, 0];
      return [u32(word, FC_LCB + index * 8), u32(word, FC_LCB + index * 8 + 4)];
    },
  };
}
