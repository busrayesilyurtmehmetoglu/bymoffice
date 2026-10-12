// Reads an xlsx package into the spreadsheet model.
//
// The workbook part, styles and theme are small and go through the tree
// parser; shared strings and sheet data are scanned as a stream.

import type { OpcPackage, Relationship } from '../../opc';
import { MAX_WORKBOOK_CHARTS, MAX_WORKBOOK_CHART_POINTS } from '../../model/sheet/chart';
import { createSheet, type DefinedName, type Sheet, type SheetKind, type SheetState, type Workbook } from '../../model/sheet/types';
import { childElements, localName, type XmlDocument, type XmlElement } from '../../xml';
import { parseChart } from './chart';
import { parsePersons, readSheetComments, vmlTakesNotes } from './comments';
import { readSheetDrawings, type ChartReading } from './drawings';
import { parseSharedStrings } from './shared-strings';
import { parseStyles } from './styles';
import { parseChartSheetSetup, parseWorksheet } from './worksheet';

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

  // A chart or a pivot table reads cells of some sheet from a part that says nothing about which.
  // While a workbook has any, rows and columns stay where they are in all of its sheets.
  const global: string[] = [];
  const allParts = pkg.partNames().map((part) => part.toLowerCase());
  if (allParts.some((part) => part.startsWith('/xl/charts/'))) global.push('chart');
  if (allParts.some((part) => part.startsWith('/xl/pivottables/') || part.startsWith('/xl/pivotcache/'))) global.push('pivotTable');

  // The people of threaded comments are kept once for the workbook.
  let persons = new Map<string, string>();
  const personsPart = rels.find((rel) => rel.type.endsWith('/person'))?.part;
  if (personsPart && pkg.hasPart(personsPart)) {
    try {
      persons = parsePersons((await pkg.readXml(personsPart)).root);
    } catch (error) {
      // Comments are then shown without the names of who wrote them.
    }
  }
  /** Parts of the pictures the drawings show: the only ones the view may ask for. */
  const mediaParts = new Set<string>();
  /** What the charts of the whole workbook may cost: the limits hold for all sheets together. */
  const chartReading: ChartReading = { read: parseChart, charts: MAX_WORKBOOK_CHARTS, points: MAX_WORKBOOK_CHART_POINTS };

  const sheets: Sheet[] = [];
  const seenParts = new Set<string>();
  /** The sheet that shows each comments part and legacy drawing: a part two sheets show cannot be changed for one of them. */
  const noteParts = new Map<string, Sheet>();
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
      const sheetRels = await pkg.relationships(rel.part);
      if (sheet.hyperlinks.some((link) => link.rId !== undefined)) {
        const targets = new Map(sheetRels.map((r) => [r.id, r.target]));
        for (const link of sheet.hyperlinks) {
          if (link.rId !== undefined) link.target = targets.get(link.rId);
        }
      }
      // Everything related to the sheet except links and printer settings is placed by cell position.
      const anchored = new Set<string>();
      let commentParts = 0;
      for (const r of sheetRels) {
        const type = r.type.slice(r.type.lastIndexOf('/') + 1);
        if ((type === 'comments' || type === 'vmlDrawing') && r.part && pkg.hasPart(r.part)) {
          if (type === 'comments' && ++commentParts > 1 && sheet.notesBlocked === undefined) sheet.notesBlocked = 'parts';
          const key = r.part.toLowerCase();
          const other = noteParts.get(key);
          if (other && other !== sheet) {
            other.notesBlocked = 'shared';
            sheet.notesBlocked = 'shared';
          } else noteParts.set(key, sheet);
        }
        if (type !== 'hyperlink' && type !== 'printerSettings') anchored.add(type);
        // A drawing may hold charts, whose series refer to cells of this sheet from yet another part.
        if (type === 'drawing' && r.part && pkg.hasPart(r.part)) {
          for (const inner of await pkg.relationships(r.part)) {
            const innerType = inner.type.slice(inner.type.lastIndexOf('/') + 1);
            if (innerType !== 'image' && innerType !== 'hyperlink') anchored.add(innerType);
          }
        }
        // Comment boxes are moved by the writer; form controls in the same part are not.
        if (type === 'vmlDrawing' && r.part && pkg.hasPart(r.part)) {
          const vmlBytes = await pkg.readPart(r.part);
          // A drawing without so much as a root takes no boxes: the notes of the sheet are shown, not changed.
          if (sheet.notesBlocked === undefined && !vmlTakesNotes(vmlBytes)) sheet.notesBlocked = 'drawing';
          const vml = new TextDecoder().decode(vmlBytes);
          const kinds = vml.match(/ObjectType\s*=\s*"[^"]*"/g) || [];
          if (kinds.some((kind) => !/"Note"$/.test(kind))) anchored.add('formControl');
        }
      }
      for (const kind of global) anchored.add(kind);
      if (anchored.size > 0) sheet.anchored = [...anchored].sort();
      // What is shown of them; the parts themselves are written back as they are.
      const warnings: string[] = [];
      const drawings = await readSheetDrawings(pkg, sheetRels, styles.themeColors, mediaParts, warnings, chartReading);
      if (drawings) sheet.drawings = drawings;
      const comments = await readSheetComments(pkg, sheetRels, persons, warnings);
      if (comments) sheet.comments = comments;
      if (warnings.length > 0) sheet.warnings = warnings;
    } else {
      // The chart of a chart sheet is found through the drawing its relationships name. Whatever
      // those parts hold, the sheet is still listed: without a chart that can be read it shows a note.
      try {
        const warnings: string[] = [];
        const drawings = await readSheetDrawings(pkg, await pkg.relationships(rel.part), styles.themeColors, mediaParts, warnings, chartReading);
        if (drawings) sheet.drawings = drawings;
        // A chart that could not be read is said, as on a worksheet: the sheet then shows its note.
        if (warnings.length > 0) sheet.warnings = warnings;
      } catch (error) {
        // The sheet stays as it was before charts were read: kept, not shown.
      }
      // How the sheet is printed. A part that cannot be read prints the way a new chart sheet does.
      try {
        const setup = parseChartSheetSetup(await pkg.readPart(rel.part));
        if (setup) sheet.pageSetup = setup;
      } catch (error) {
        // Nothing of the part is needed to show the sheet.
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
  const workbook: Workbook = {
    sheets,
    styles,
    sharedStrings,
    definedNames,
    date1904: isTrue(properties?.attrs.get('date1904')),
    activeSheet: activeTab < sheets.length ? activeTab : 0,
  };
  if (mediaParts.size > 0) {
    // Read from the package when a picture is first shown, and only what a drawing of this workbook shows.
    workbook.media = {
      readPart: (part: string) => (mediaParts.has(part) ? pkg.readPart(part) : Promise.reject(new Error(`Not a picture of this workbook: ${part}`))),
    };
  }
  return workbook;
}
