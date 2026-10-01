// Reads an xlsx package into the spreadsheet model.
//
// The workbook part, styles and theme are small and go through the tree
// parser; shared strings and sheet data are scanned as a stream.

import type { OpcPackage, Relationship } from '../../opc';
import { createSheet, type DefinedName, type Sheet, type SheetKind, type SheetState, type Workbook } from '../../model/sheet/types';
import { childElements, localName, type XmlDocument, type XmlElement } from '../../xml';
import { parseSharedStrings } from './shared-strings';
import { parseStyles } from './styles';
import { parseWorksheet } from './worksheet';

// Relationship types end the same way in the transitional and the strict namespace.
const SHEET_KINDS: [suffix: string, kind: SheetKind][] = [
  ['/worksheet', 'worksheet'],
  ['/chartsheet', 'chartsheet'],
  ['/dialogsheet', 'dialogsheet'],
  ['/xlMacrosheet', 'macrosheet'],
  ['/xlIntlMacrosheet', 'macrosheet'],
];

const child = (parent: XmlElement, local: string): XmlElement | undefined => childElements(parent, local)[0];

/** Attribute by local name, for attributes that carry a namespace prefix such as r:id. */
function prefixedAttr(element: XmlElement, local: string): string | undefined {
  for (const [name, value] of element.attrs) {
    if (name.includes(':') && localName(name) === local) return value;
  }
  return undefined;
}

function textOf(element: XmlElement): string {
  let text = '';
  for (const node of element.children) if (node.type === 'text') text += node.value;
  return text;
}

const isTrue = (value: string | undefined) => value === '1' || value === 'true';

async function optionalXml(pkg: OpcPackage, rels: Relationship[], suffix: string): Promise<XmlDocument | undefined> {
  const part = rels.find((rel) => rel.type.endsWith(suffix))?.part;
  return part && pkg.hasPart(part) ? pkg.readXml(part) : undefined;
}

export async function readXlsx(pkg: OpcPackage): Promise<Workbook> {
  const main = await pkg.mainPart();
  if (!main) throw new Error('Invalid xlsx: the package has no workbook part');
  const root = (await pkg.readXml(main)).root;
  const rels = await pkg.relationships(main);
  const relById = new Map(rels.map((rel) => [rel.id, rel]));

  const stringsPart = rels.find((rel) => rel.type.endsWith('/sharedStrings'))?.part;
  const sharedStrings = stringsPart && pkg.hasPart(stringsPart) ? parseSharedStrings(await pkg.readPart(stringsPart)) : [];
  const styles = parseStyles(await optionalXml(pkg, rels, '/styles'), await optionalXml(pkg, rels, '/theme'));

  const sheets: Sheet[] = [];
  const seenParts = new Set<string>();
  const sheetsElement = child(root, 'sheets');
  for (const element of sheetsElement ? childElements(sheetsElement, 'sheet') : []) {
    const name = element.attrs.get('name') ?? '';
    const rel = relById.get(prefixedAttr(element, 'id') ?? '');
    if (!rel || !rel.part) throw new Error(`Invalid xlsx: sheet "${name}" has no part`);
    // Each sheet has its own part; a part listed twice would be read into memory twice.
    if (seenParts.has(rel.part)) throw new Error(`Invalid xlsx: sheet "${name}" shares its part with another sheet`);
    seenParts.add(rel.part);
    const kind = SHEET_KINDS.find(([suffix]) => rel.type.endsWith(suffix))?.[1];
    if (!kind) throw new Error(`Invalid xlsx: sheet "${name}" has an unknown type`);
    const state = element.attrs.get('state');
    const sheet = createSheet({
      name,
      sheetId: Number(element.attrs.get('sheetId')) || 0,
      state: (state === 'hidden' || state === 'veryHidden' ? state : 'visible') as SheetState,
      kind,
      part: rel.part,
    });
    // A chart sheet holds a drawing, not cells; it stays in the package as it is.
    if (kind !== 'chartsheet') {
      parseWorksheet(await pkg.readPart(rel.part), sharedStrings, sheet);
      if (sheet.hyperlinks.some((link) => link.rId !== undefined)) {
        const targets = new Map((await pkg.relationships(rel.part)).map((r) => [r.id, r.target]));
        for (const link of sheet.hyperlinks) {
          if (link.rId !== undefined) link.target = targets.get(link.rId);
        }
      }
    }
    sheets.push(sheet);
  }
  if (sheets.length === 0) throw new Error('Invalid xlsx: the workbook has no sheets');

  const definedNames: DefinedName[] = [];
  const namesElement = child(root, 'definedNames');
  for (const element of namesElement ? childElements(namesElement, 'definedName') : []) {
    const local = element.attrs.get('localSheetId');
    definedNames.push({
      name: element.attrs.get('name') ?? '',
      value: textOf(element),
      localSheet: local !== undefined && /^\d{1,6}$/.test(local) ? Number(local) : undefined,
      hidden: isTrue(element.attrs.get('hidden')),
    });
  }

  const properties = child(root, 'workbookPr');
  const views = child(root, 'bookViews');
  const view = views && child(views, 'workbookView');
  const activeAttr = view?.attrs.get('activeTab') ?? '';
  const activeTab = /^\d{1,6}$/.test(activeAttr) ? Number(activeAttr) : 0;
  return {
    sheets,
    styles,
    sharedStrings,
    definedNames,
    date1904: isTrue(properties?.attrs.get('date1904')),
    activeSheet: activeTab < sheets.length ? activeTab : 0,
  };
}
