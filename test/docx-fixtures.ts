// Synthetic docx packages for the docx reader tests. Everything is built from
// XML fragments in code; no real document is involved.

import { OpcPackage } from '../src/opc';
import { readDocx } from '../src/formats/docx/document';
import type { DocModel } from '../src/model/doc/types';
import { ZipArchive } from '../src/zip';

const encoder = new TextEncoder();
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS_PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

export interface Namespaces {
  w: string;
  r: string;
  rel: string;
}

export const TRANSITIONAL: Namespaces = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
};

export const STRICT: Namespaces = {
  w: 'http://purl.oclc.org/ooxml/wordprocessingml/main',
  r: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
  rel: 'http://purl.oclc.org/ooxml/officeDocument/relationships',
};

export const NS_DECLARATIONS =
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

export interface DocxSpec {
  /** Inner XML of <w:body>. */
  body?: string;
  /** The whole document part, instead of `body`. */
  document?: string | Uint8Array;
  /** Inner XML of the root element of each supporting part. */
  styles?: string;
  numbering?: string;
  settings?: string;
  fontTable?: string;
  footnotes?: string;
  endnotes?: string;
  comments?: string;
  /** The whole theme part. */
  theme?: string;
  /** Inner XML of <w:hdr> / <w:ftr> by relationship id. */
  headers?: Record<string, string>;
  footers?: Record<string, string>;
  /** More relationships of the document part: [id, type suffix, target, external]. */
  rels?: [id: string, type: string, target: string, external?: boolean][];
  /** Raw parts to add: [zip entry name, content]. */
  parts?: [name: string, content: string | Uint8Array][];
  /** Supporting parts to relate but not write ("styles", "numbering", ...). */
  missing?: string[];
  strict?: boolean;
}

export async function buildDocx(spec: DocxSpec): Promise<Uint8Array> {
  const ns = spec.strict ? STRICT : TRANSITIONAL;
  const root = (name: string, inner: string) => `${XML}<w:${name} xmlns:w="${ns.w}" xmlns:r="${ns.r}" ${NS_DECLARATIONS}>${inner}</w:${name}>`;
  const zip = new ZipArchive();
  const rels: string[] = [];
  const put = (name: string, content: string | Uint8Array) => zip.set(name, typeof content === 'string' ? encoder.encode(content) : content);

  let next = 1;
  const supporting = (type: string, target: string, content: string | undefined, id?: string) => {
    const missing = spec.missing !== undefined && spec.missing.indexOf(type) >= 0;
    if (content === undefined && !missing) return;
    rels.push(`<Relationship Id="${id === undefined ? `rIdS${next++}` : id}" Type="${ns.rel}/${type}" Target="${target}"/>`);
    if (content !== undefined && !missing) put(`word/${target}`, content);
  };
  const wrapped = (name: string, inner: string | undefined) => (inner === undefined ? undefined : root(name, inner));
  supporting('styles', 'styles.xml', wrapped('styles', spec.styles));
  supporting('numbering', 'numbering.xml', wrapped('numbering', spec.numbering));
  supporting('settings', 'settings.xml', wrapped('settings', spec.settings));
  supporting('fontTable', 'fontTable.xml', wrapped('fonts', spec.fontTable));
  supporting('footnotes', 'footnotes.xml', wrapped('footnotes', spec.footnotes));
  supporting('endnotes', 'endnotes.xml', wrapped('endnotes', spec.endnotes));
  supporting('comments', 'comments.xml', wrapped('comments', spec.comments));
  supporting('theme', 'theme/theme1.xml', spec.theme);
  Object.keys(spec.headers || {}).forEach((id, i) => supporting('header', `header${i + 1}.xml`, wrapped('hdr', spec.headers![id]), id));
  Object.keys(spec.footers || {}).forEach((id, i) => supporting('footer', `footer${i + 1}.xml`, wrapped('ftr', spec.footers![id]), id));
  for (const [id, type, target, external] of spec.rels || []) {
    rels.push(`<Relationship Id="${id}" Type="${ns.rel}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`);
  }

  put(
    '[Content_Types].xml',
    `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  );
  put('_rels/.rels', `${XML}<Relationships xmlns="${NS_PKG_REL}"><Relationship Id="rId1" Type="${ns.rel}/officeDocument" Target="word/document.xml"/></Relationships>`);
  put('word/document.xml', spec.document !== undefined ? spec.document : root('document', `<w:body>${spec.body || ''}</w:body>`));
  put('word/_rels/document.xml.rels', `${XML}<Relationships xmlns="${NS_PKG_REL}">${rels.join('')}</Relationships>`);
  for (const [name, content] of spec.parts || []) put(name, content);
  return zip.write();
}

/** Builds the package and reads it into the model. */
export async function readSpec(spec: DocxSpec): Promise<DocModel> {
  return readDocx(await OpcPackage.open(await buildDocx(spec)));
}

/** A complete document part around the given body XML. */
export const documentXml = (body: string, ns: Namespaces = TRANSITIONAL) =>
  `${XML}<w:document xmlns:w="${ns.w}" xmlns:r="${ns.r}" ${NS_DECLARATIONS}><w:body>${body}</w:body></w:document>`;

export const THEME =
  `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test"><a:themeElements><a:clrScheme name="Test">` +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="Test"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface="Times New Roman"/>' +
  '<a:font script="Jpan" typeface="Yu Gothic Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme>' +
  '</a:themeElements></a:theme>';
