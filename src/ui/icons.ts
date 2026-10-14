// The editor's icons: line drawings on a 24 x 24 grid that take the colour of
// the text around them. They are built element by element (never from markup),
// so nothing here can become HTML.

import type { BorderPreset } from '../ops/format';

export type IconName =
  | 'save'
  | 'undo'
  | 'redo'
  | 'print'
  | 'pdf'
  | 'pageSetup'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'textColor'
  | 'fillColor'
  | 'borders'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'alignTop'
  | 'alignMiddle'
  | 'alignBottom'
  | 'wrapText'
  | 'merge'
  | 'mergeOptions'
  | 'sortAscending'
  | 'sortDescending'
  | 'filter'
  | 'validation'
  | 'conditionalFormat'
  | 'chart'
  | 'rowsColumns'
  | 'find';

type Shape = [tag: string, attrs: { [name: string]: string }];

const SVG_NS = 'http://www.w3.org/2000/svg';
/** Class of the strip under the colour icons; the toolbar paints it with the current colour. */
export const ICON_BAR_CLASS = 'bym-icon-bar';

const path = (d: string, extra?: { [name: string]: string }): Shape => ['path', { d, ...(extra || {}) }];
const bar: Shape = ['rect', { class: ICON_BAR_CLASS, x: '3.5', y: '18', width: '17', height: '4', rx: '1', stroke: 'none', fill: 'currentColor' }];

const ICONS: { [name in IconName]: Shape[] } = {
  // A disk: its label, its shutter, the notch of its corner.
  save: [path('M5 4h11l3 3v13H5z'), path('M8 4v5h7V4'), path('M8 20v-6h8v6')],
  undo: [path('M9 14 4 9l5-5'), path('M4 9h10a6 6 0 0 1 0 12h-3')],
  redo: [path('m15 14 5-5-5-5'), path('M20 9H10a6 6 0 0 0 0 12h3')],
  // The printer of the document bar.
  print: [path('M7 8V4h10v4'), path('M7 17H5a1 1 0 0 1-1-1v-6a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v6a1 1 0 0 1-1 1h-2'), path('M7 14h10v6H7z')],
  // The file of the document bar: a sheet of paper with its corner folded, and an arrow down.
  pdf: [path('M14 3H7a1 1 0 0 0-1 1v16a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1V7z'), path('M14 3v4h4'), path('M12 10v7'), path('m9 14 3 3 3-3')],
  // A page with its margins drawn in, and a corner turned: the page setup.
  pageSetup: [path('M6 3h12v18H6z'), path('M6 7h12 M6 17h12 M9 3v18 M15 3v18', { 'stroke-width': '1', 'stroke-dasharray': '1.5 1.5' })],
  bold: [path('M7 5h6a3.5 3.5 0 0 1 0 7H7z M7 12h7a3.5 3.5 0 0 1 0 7H7z', { 'stroke-width': '2.25' })],
  italic: [path('M10 5h7 M7 19h7 M14 5l-4 14')],
  underline: [path('M7 4v7a5 5 0 0 0 10 0V4 M5 20h14')],
  strikethrough: [
    path('M16.5 7.6C16 5.6 14.3 4.5 12 4.5c-2.6 0-4.5 1.3-4.5 3.4 0 1.2.6 2.1 1.7 2.7'),
    path('M7.3 16.2c.5 2 2.3 3.3 4.8 3.3 2.7 0 4.6-1.4 4.6-3.5 0-.7-.2-1.3-.5-1.8'),
    path('M4 12h16'),
  ],
  textColor: [path('M7 15 12 3.5 17 15 M9 11h6'), bar],
  fillColor: [
    path('M10.5 3 4.6 8.9a1.5 1.5 0 0 0 0 2.1l3.9 3.9a1.5 1.5 0 0 0 2.1 0L16.5 9z'),
    path('M4.6 9.6h11.3'),
    path('M19.3 11.5s1.7 2.1 1.7 3.2a1.7 1.7 0 0 1-3.4 0c0-1.1 1.7-3.2 1.7-3.2z'),
    bar,
  ],
  borders: [path('M4 4h16v16H4z M12 4v16 M4 12h16')],
  alignLeft: [path('M4 6h16 M4 10h10 M4 14h16 M4 18h10')],
  alignCenter: [path('M4 6h16 M7 10h10 M4 14h16 M7 18h10')],
  alignRight: [path('M4 6h16 M10 10h10 M4 14h16 M10 18h10')],
  // A line of text at the top, in the middle, at the bottom of a cell.
  alignTop: [path('M4 4h16'), path('M8 8h8 M8 11.5h8')],
  alignMiddle: [path('M4 4h16 M4 20h16'), path('M8 10.25h8 M8 13.75h8')],
  alignBottom: [path('M4 20h16'), path('M8 12.5h8 M8 16h8')],
  wrapText: [path('M4 6h16 M4 18h5 M4 12h13a3 3 0 0 1 0 6h-4'), path('m15 16-2 2 2 2')],
  merge: [path('M3 6h18v12H3z'), path('M6 12h4 M8 10l2 2-2 2 M18 12h-4 M16 10l-2 2 2 2')],
  // The other ways of merging: a narrow arrow down beside the merge button.
  mergeOptions: [path('m8 10 4 4 4-4')],
  // Bars that grow (or shrink) beside an arrow that points down.
  sortAscending: [path('M17 5v14 M14 16l3 3 3-3'), path('M4 7h4 M4 12h6 M4 17h8')],
  sortDescending: [path('M17 5v14 M14 16l3 3 3-3'), path('M4 7h8 M4 12h6 M4 17h4')],
  filter: [path('M4 5h16l-6.2 7.4V18l-3.6 1.8v-7.4z')],
  // Cells of which some stand out: a format that depends on the value.
  conditionalFormat: [path('M4 5h16v14H4z M4 9.7h16 M4 14.3h16'), path('M6.4 12h6.2', { 'stroke-width': '2.6' })],
  // A cell with a tick: what may be typed into it.
  validation: [path('M4 6h16v12H4z'), path('m8.5 12.2 2.4 2.4 4.6-5')],
  // Three columns on a base line.
  chart: [path('M4 20h16'), path('M6.5 20v-6h3v6 M10.5 20V7h3v13 M14.5 20v-9h3v9')],
  rowsColumns: [path('M4 5h16v14H4z M4 10h16 M10 5v14')],
  // A magnifying glass: find and replace.
  find: [path('M10.5 4a6.5 6.5 0 1 1 0 13 6.5 6.5 0 0 1 0-13z'), path('m15.3 15.3 4.7 4.7')],
};

/** Icons only the document editor has; the others it shares with the spreadsheet editor. */
export type DocIconName =
  | 'highlight'
  | 'alignJustify'
  | 'bulletList'
  | 'numberedList'
  | 'decreaseIndent'
  | 'increaseIndent'
  | 'tabStops'
  | 'insertTable'
  | 'insertPicture'
  | 'insertLink'
  | 'pageBreak'
  | 'trackChanges'
  | 'previousChange'
  | 'nextChange'
  | 'acceptChange'
  | 'rejectChange'
  | 'changesMenu';

const DOC_ICONS: { [name in DocIconName]: Shape[] } = {
  highlight: [path('m14.5 4 4.5 4.5-7.5 7.5H7v-4.5z'), path('M12.5 6 17 10.5'), bar],
  alignJustify: [path('M4 6h16 M4 10h16 M4 14h16 M4 18h16')],
  bulletList: [path('M9 6h11 M9 12h11 M9 18h11'), path('M4.5 6h.01 M4.5 12h.01 M4.5 18h.01', { 'stroke-width': '2.75' })],
  numberedList: [path('M10 6h10 M10 12h10 M10 18h10'), path('M4 5l1.5-1v4 M4 8h3', { 'stroke-width': '1.25' }), path('M4 14.5c0-1.5 2.8-1.5 2.8 0 0 1-2.8 2-2.8 3.5h3', { 'stroke-width': '1.25' })],
  decreaseIndent: [path('M4 5h16 M11 10h9 M11 14h9 M4 19h16'), path('m7.5 9.5-3 2.5 3 2.5')],
  increaseIndent: [path('M4 5h16 M11 10h9 M11 14h9 M4 19h16'), path('m4.5 9.5 3 2.5-3 2.5')],
  // A ruler with the mark of a left tab stop on it.
  tabStops: [path('M3 15h18 M6 15v-2 M10 15v-2 M14 15v-2 M18 15v-2'), path('M8 5v5h4', { 'stroke-width': '2.25' })],
  insertTable: [path('M4 5h16v14H4z M4 10h16 M4 14.5h16 M9.5 5v14 M14.5 5v14')],
  insertPicture: [path('M4 5h16v14H4z'), path('m4 16 4.5-4.5 4 4 2.5-2.5 5 5'), path('M15.5 9.5h.01', { 'stroke-width': '2.75' })],
  insertLink: [path('M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1'), path('M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1')],
  pageBreak: [path('M6 3v5h12V3 M6 21v-5h12v5'), path('M3 12h2 M8 12h2 M14 12h2 M19 12h2')],
  // A page with a pencil: what is written is marked.
  trackChanges: [path('M13 4H6v16h12v-9'), path('m19.5 3.5 1.5 1.5-7 7-2.5 1 1-2.5z'), path('M9 16h6')],
  // A page with an arrow up, and down: the change before, the change after.
  previousChange: [path('M6 20V4h8l4 4v12z'), path('M12 17v-7 M9 13l3-3 3 3')],
  nextChange: [path('M6 20V4h8l4 4v12z'), path('M12 10v7 M9 14l3 3 3-3')],
  acceptChange: [path('M6 20V4h8l4 4v12z'), path('m9 13.5 2.2 2.2 4-4.4')],
  rejectChange: [path('M6 20V4h8l4 4v12z'), path('m9.5 11.5 5 5 M14.5 11.5l-5 5')],
  // Two pages and a tick: all of the changes.
  changesMenu: [path('M8 20V7h7l3 3v10z'), path('M5 17V4h8'), path('m10.5 14.5 1.8 1.8 3-3.4')],
};

export function createDocIcon(name: DocIconName): SVGSVGElement {
  return build(DOC_ICONS[name]);
}

function build(shapes: Shape[]): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg') as SVGSVGElement;
  const attrs: { [name: string]: string } = {
    viewBox: '0 0 24 24',
    width: '20',
    height: '20',
    fill: 'none',
    stroke: 'currentColor',
    'stroke-width': '1.75',
    'stroke-linecap': 'round',
    'stroke-linejoin': 'round',
    'aria-hidden': 'true',
    focusable: 'false',
  };
  for (const name of Object.keys(attrs)) svg.setAttribute(name, attrs[name]);
  svg.setAttribute('class', 'bym-icon');
  for (let i = 0; i < shapes.length; i++) {
    const element = document.createElementNS(SVG_NS, shapes[i][0]);
    const own = shapes[i][1];
    for (const name of Object.keys(own)) element.setAttribute(name, own[name]);
    svg.appendChild(element);
  }
  return svg;
}

export function createIcon(name: IconName): SVGSVGElement {
  return build(ICONS[name]);
}

/** A cell block with the edges a border preset draws in full and the others faint and dashed. */
export function createBorderIcon(preset: BorderPreset): SVGSVGElement {
  const faint = { 'stroke-width': '1', 'stroke-dasharray': '1.5 2.5', opacity: '0.55' };
  const strong = { 'stroke-width': '2' };
  const edges: { [edge: string]: string } = { top: 'M4 4h16', bottom: 'M4 20h16', left: 'M4 4v16', right: 'M20 4v16' };
  const inner = 'M12 4v16 M4 12h16';
  const outline = preset === 'all' || preset === 'outline';
  const shapes: Shape[] = [];
  let weak = '';
  let solid = '';
  for (const edge of Object.keys(edges)) {
    if (outline || preset === edge) solid += ` ${edges[edge]}`;
    else weak += ` ${edges[edge]}`;
  }
  if (preset === 'all') solid += ` ${inner}`;
  else weak += ` ${inner}`;
  if (weak) shapes.push(path(weak.trim(), faint));
  if (solid) shapes.push(path(solid.trim(), strong));
  return build(shapes);
}
