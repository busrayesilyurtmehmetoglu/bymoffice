// The destinations of an RTF file: the control words that say what the group
// they start is (the font table, a picture, a header, ...) rather than how
// its text looks.
//
// A group that starts with "\*" and a control word the reader does not know
// is passed over; so is a destination the reader knows and has no use for
// (the list below), which matters where the writer left out the "\*".
// A control word that is not known and not marked that way is passed over by
// itself: what its group holds is read as text, as the specification says.

/** Text and paragraphs of the story being read. */
export const D_NORMAL = 0;
export const D_SKIP = 1;
export const D_FONTTBL = 2;
export const D_FALT = 3;
export const D_PANOSE = 4;
export const D_COLORTBL = 5;
export const D_STYLESHEET = 6;
export const D_STYLE = 7;
export const D_LISTTABLE = 8;
export const D_LIST = 9;
export const D_LISTLEVEL = 10;
export const D_LEVELTEXT = 11;
export const D_OVERRIDETABLE = 12;
export const D_OVERRIDE = 13;
export const D_LFOLEVEL = 14;
export const D_PICT = 15;
export const D_SHP = 16;
export const D_SHPINST = 17;
export const D_SP = 18;
export const D_SN = 19;
export const D_SV = 20;
export const D_FIELD = 21;
export const D_BKMKSTART = 22;
export const D_BKMKEND = 23;
export const D_ATNID = 24;
export const D_ATNAUTHOR = 25;
export const D_ATRFSTART = 26;
export const D_ATRFEND = 27;
export const D_ATNREF = 28;
export const D_UPR = 29;
/** A group directly in "\upr": passed over unless it turns out to be "\ud". */
export const D_UPR_CHILD = 30;
export const D_OBJECT = 31;
/** A drawing object of Word 95 ("\do"). */
export const D_DO = 32;

/** Destinations that only gather text (a name, a value). */
export const isCollector = (dest: number): boolean =>
  dest === D_FALT ||
  dest === D_PANOSE ||
  dest === D_STYLE ||
  dest === D_SN ||
  dest === D_SV ||
  dest === D_BKMKSTART ||
  dest === D_BKMKEND ||
  dest === D_ATNID ||
  dest === D_ATNAUTHOR ||
  dest === D_ATRFSTART ||
  dest === D_ATRFEND ||
  dest === D_ATNREF;

/** Destinations the reader knows and passes over, with or without "\*" in front. */
export const SKIPPED = [
  // The document: who wrote it and when, what it was made with.
  'info',
  'title',
  'subject',
  'author',
  'manager',
  'company',
  'operator',
  'category',
  'keywords',
  'comment',
  'doccomm',
  'hlinkbase',
  'creatim',
  'revtim',
  'printim',
  'buptim',
  'userprops',
  'docvar',
  'template',
  'generator',
  'filetbl',
  'revtbl',
  'rsidtbl',
  'protusertbl',
  'passwordhash',
  'xmlnstbl',
  'mmathPr',
  'wgrffmtfilter',
  'pgdsctbl',
  'pgptbl',
  // Formatting for other programs, or stated again in the text.
  'themedata',
  'colorschememapping',
  'latentstyles',
  'datastore',
  'defchp',
  'defpap',
  'listpicture',
  'liststylename',
  'listname',
  'levelnumbers',
  'fontemb',
  'fontfile',
  'fname',
  'fchars',
  'lchars',
  'background',
  // Numbering of Word 95: the number itself is in the text ("\pntext").
  'pn',
  'pnseclvl',
  'pntxta',
  'pntxtb',
  // What stands in for something the reader reads itself.
  'nonshppict',
  'shprslt',
  'nonesttables',
  'picprop',
  'blipuid',
  // Fields and objects: data the reader does not look into, and never runs.
  'datafield',
  'formfield',
  'fldtype',
  'objdata',
  'objclass',
  'objname',
  'objalias',
  'objsect',
  'objtime',
  'oleclsid',
  // Index and table of contents entries: text that does not show.
  'xe',
  'tc',
  'tcn',
  'rxe',
  'txe',
  'bxe',
  'ixe',
  // Comments: what the reader has no place for.
  'atndate',
  'atntime',
  'atnicn',
  'atnparent',
  // Mail merge, form letters, private data.
  'mailmerge',
  'private',
  'xform',
  'ffname',
  'ffdeftext',
  'ffformat',
  'ffhelptext',
  'ffstattext',
  'ffentrymcr',
  'ffexitmcr',
  'ffl',
  // Mathematics: the reader has no model for it.
  'mmath',
];
