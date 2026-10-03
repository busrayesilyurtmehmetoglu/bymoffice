// New, empty files: a document with one empty paragraph and the styles an
// editor offers (headings, title), a workbook with one empty sheet. What an
// application gives its users when they ask for a new file instead of
// opening one.

import { createSheet, emptyStyleSheet, type Workbook } from '../model/sheet/types';
import { ZipArchive } from '../zip';
import { buildXlsx } from './xlsx/build';

export interface BlankOptions {
  /**
   * The language of the file, as a tag ("tr", "tr-TR", "en-GB"; the case and
   * an underscore for the hyphen do not matter): Turkish gives a sheet named
   * "Sayfa1" and Turkish separators; "en", "en-US" and "en-CA" a document on
   * Letter paper, every other language one on A4. What is not a language tag
   * counts as the default, "en-GB".
   */
  locale?: string;
  /** xlsx: the name of the sheet instead of the one the language gives. */
  sheetName?: string;
}

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WORD_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml';

/**
 * A language tag as a file may carry it: letters, digits and hyphens, the
 * language in small letters and the region in capitals whichever way it was
 * given ("TR", "tr_tr" and "tr-TR" are the same).
 */
function languageOf(locale: string | undefined): string {
  const tag = typeof locale === 'string' ? locale.trim().replace(/_/g, '-') : '';
  if (!/^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8}){0,3}$/.test(tag)) return 'en-GB';
  const parts = tag.split('-');
  parts[0] = parts[0].toLowerCase();
  for (let i = 1; i < parts.length; i++) {
    // A region is two letters or three digits; a script has four letters and a capital first.
    if (/^([A-Za-z]{2}|[0-9]{3})$/.test(parts[i])) parts[i] = parts[i].toUpperCase();
    else if (/^[A-Za-z]{4}$/.test(parts[i])) parts[i] = parts[i].charAt(0).toUpperCase() + parts[i].slice(1).toLowerCase();
  }
  if (parts.length === 1) {
    if (parts[0] === 'tr') return 'tr-TR';
    if (parts[0] === 'en') return 'en-US';
  }
  return parts.join('-');
}

/** Turkish, whatever the region: the names and the separators of the language. */
const isTurkish = (language: string): boolean => language === 'tr' || language.slice(0, 3) === 'tr-';

/**
 * What a sheet may be called: 1 to 31 characters that are not all spaces,
 * none of []:*?/\ and no control character, no apostrophe at either end,
 * not "History" (which Excel keeps for itself), and only characters an XML
 * file can hold.
 */
function checkSheetName(name: string): string {
  const ok =
    typeof name === 'string' &&
    name.length >= 1 &&
    name.length <= 31 &&
    name.trim() !== '' &&
    !/[[\]:*?/\\\u0000-\u001f\ufffe\uffff]/.test(name) &&
    // A surrogate without its other half.
    !/[\ud800-\udbff](?![\udc00-\udfff])|(?:^|[^\ud800-\udbff])[\udc00-\udfff]/.test(name) &&
    name.charAt(0) !== "'" &&
    name.charAt(name.length - 1) !== "'" &&
    name.toLowerCase() !== 'history';
  if (!ok) throw new RangeError('A sheet name has 1 to 31 characters, none of [ ] : * ? / \\ and no control character, no apostrophe at its start or end, and is not "History"');
  return name;
}

const heading = (level: number, size: number, before: number): string =>
  `<w:style w:type="paragraph" w:styleId="Heading${level}"><w:name w:val="heading ${level}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>` +
  `<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="80"/><w:outlineLvl w:val="${level - 1}"/></w:pPr>` +
  `<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:b/><w:bCs/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

function documentParts(language: string): [name: string, text: string][] {
  // Letter where it is the paper of the country, A4 everywhere else; margins of 2.5 cm (one inch on Letter).
  const letter = language === 'en-US' || language === 'en-CA';
  const page = letter ? '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' : '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>';
  const styles =
    `${XML_DECLARATION}<w:styles xmlns:w="${W}">` +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/>' +
    `<w:lang w:val="${language}" w:eastAsia="${language}" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>` +
    '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    heading(1, 32, 360) +
    heading(2, 26, 240) +
    heading(3, 24, 200) +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>' +
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' +
    '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>' +
    '</w:styles>';
  const settings = `${XML_DECLARATION}<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="${letter ? 720 : 708}"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:decimalSymbol w:val="${isTurkish(language) ? ',' : '.'}"/><w:listSeparator w:val="${isTurkish(language) ? ';' : ','}"/></w:settings>`;
  const document = `${XML_DECLARATION}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p/><w:sectPr>${page}<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`;
  const types =
    `${XML_DECLARATION}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/word/document.xml" ContentType="${WORD_TYPE}.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="${WORD_TYPE}.styles+xml"/>` +
    `<Override PartName="/word/settings.xml" ContentType="${WORD_TYPE}.settings+xml"/></Types>`;
  return [
    ['[Content_Types].xml', types],
    ['_rels/.rels', `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', document],
    ['word/_rels/document.xml.rels', `${XML_DECLARATION}<Relationships xmlns="${PACKAGE_RELS}"><Relationship Id="rId1" Type="${R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${R}/settings" Target="settings.xml"/></Relationships>`],
    ['word/styles.xml', styles],
    ['word/settings.xml', settings],
  ];
}

/**
 * A new, empty file as bytes: `docx`, a document with one empty paragraph,
 * or `xlsx`, a workbook with one empty sheet. The same options give the same
 * bytes. Throws RangeError for a kind that is neither, and for a sheet name
 * a workbook cannot have.
 */
export async function createBlank(kind: 'docx' | 'xlsx', options: BlankOptions = {}): Promise<Uint8Array> {
  const language = languageOf(options.locale);
  if (kind === 'xlsx') {
    const name = checkSheetName(options.sheetName !== undefined ? options.sheetName : isTurkish(language) ? 'Sayfa1' : 'Sheet1');
    const sheet = createSheet({ name, sheetId: 1, state: 'visible', kind: 'worksheet', part: '' });
    sheet.view.tabSelected = true;
    const workbook: Workbook = { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
    return buildXlsx(workbook);
  }
  if (kind !== 'docx') throw new RangeError('A new file is a "docx" or an "xlsx"');
  const encoder = new TextEncoder();
  const zip = new ZipArchive();
  for (const [name, text] of documentParts(language)) zip.set(name, encoder.encode(text));
  return zip.write();
}
