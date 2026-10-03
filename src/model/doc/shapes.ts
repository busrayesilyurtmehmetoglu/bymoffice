// What a drawing object of Word 97 to 2003 looks like. A binary doc states it
// as OfficeArt properties (MS-ODRAW 2.3), an RTF file as the same properties
// by name ("{\sp{\sn lineWidth}{\sv 12700}}"); both readers hand them to
// officeArtLook(), which applies the defaults of the specification: a shape
// is filled in white and outlined in black, 0.75 pt wide, unless it says
// otherwise. What a shape states and the look cannot hold (a gradient, a
// transparent fill, a colour of a palette) is named in ShapeLook.lost.

import type { ShapeDash, ShapeLook } from './types';

/** The properties of a shape officeArtLook() looks at, as the file states them; undefined for what it does not state. */
export interface OfficeArtShape {
  /** The kind of shape (MSOSPT). */
  type?: number;
  fFilled?: boolean;
  /** MSOFILLTYPE: 0 a solid colour, then patterns, pictures and gradients. */
  fillType?: number;
  /** Colours as OfficeArtCOLORREF: red in the lowest byte, flags in the highest. */
  fillColor?: number;
  /** A 16.16 fixed point number: 65536 is opaque, 0 lets everything through. */
  fillOpacity?: number;
  fLine?: boolean;
  lineColor?: number;
  /** EMU. */
  lineWidth?: number;
  /** MSOLINEDASHING. */
  lineDashing?: number;
  /** The first adjust value: for a rounded rectangle the radius of its corners in 21600ths of the shorter side. */
  adjust?: number;
}

/** Width of the outline of a shape that states none: 0.75 pt, EMU. */
export const DEFAULT_LINE_WIDTH = 9525;
/** Widest outline the model holds: 1584 pt, the limit of Word; EMU. */
const MAX_LINE_WIDTH = 20116800;

const MSOSPT_RECTANGLE = 1;
const MSOSPT_ROUND_RECTANGLE = 2;
const MSOSPT_ELLIPSE = 3;
const MSOSPT_LINE = 20;
const MSOSPT_STRAIGHT_CONNECTOR = 32;
const MSOSPT_PICTURE_FRAME = 75;
const MSOSPT_TEXT_BOX = 202;

/** MSOLINEDASHING -> the dash of DrawingML that Word shows for it; 0 is a solid line. */
const DASHES: (ShapeDash | undefined)[] = [undefined, 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot', 'dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot'];

const HEX = '0123456789ABCDEF';

/**
 * An OfficeArtCOLORREF as "RRGGBB"; undefined when it does not state a colour
 * of its own (an index into a palette or a scheme, a system colour, a colour
 * derived from another property).
 */
export function officeArtColor(value: number | undefined): string | undefined {
  if (value === undefined || value !== value) return undefined;
  const flags = (value >>> 24) & 0xff;
  // fPaletteRGB still states the colour; every other flag makes the low bytes an index.
  if (flags !== 0 && flags !== 0x02) return undefined;
  let out = '';
  for (let i = 0; i < 3; i++) {
    const channel = (value >>> (i * 8)) & 0xff;
    out += HEX.charAt(channel >> 4) + HEX.charAt(channel & 15);
  }
  return out;
}

/** The look of a shape with the OfficeArt properties. */
export function officeArtLook(shape: OfficeArtShape): ShapeLook {
  const type = shape.type;
  let geometry: ShapeLook['geometry'] = 'other';
  // A file that does not say what the shape is means a box: that is what holds text and pictures.
  if (type === undefined || type === MSOSPT_RECTANGLE || type === MSOSPT_TEXT_BOX || type === MSOSPT_PICTURE_FRAME) geometry = 'rect';
  else if (type === MSOSPT_ROUND_RECTANGLE) geometry = 'roundRect';
  else if (type === MSOSPT_ELLIPSE) geometry = 'ellipse';
  else if (type === MSOSPT_LINE || type === MSOSPT_STRAIGHT_CONNECTOR) geometry = 'line';
  const look: ShapeLook = { geometry };
  if (geometry === 'roundRect' && shape.adjust !== undefined && shape.adjust >= 0 && shape.adjust <= 10800) look.corner = Math.round((shape.adjust / 21600) * 100000);
  const lost: NonNullable<ShapeLook['lost']> = [];
  const opacity = shape.fillOpacity;
  // A line has nothing to fill, and a fill that lets everything through is none.
  if (shape.fFilled !== false && geometry !== 'line' && !(opacity !== undefined && opacity <= 0)) {
    if (shape.fillType !== undefined && shape.fillType !== 0) lost.push('fill');
    else if (shape.fillColor === undefined) look.fill = 'FFFFFF';
    else {
      const color = officeArtColor(shape.fillColor);
      if (color === undefined) lost.push('color');
      else look.fill = color;
    }
    if (look.fill !== undefined && opacity !== undefined && opacity < 65536) lost.push('opacity');
  }
  if (shape.fLine !== false) {
    const stated = shape.lineWidth;
    const width = stated === undefined || stated !== stated || stated < 0 ? DEFAULT_LINE_WIDTH : Math.min(Math.round(stated), MAX_LINE_WIDTH);
    const color = shape.lineColor === undefined ? '000000' : officeArtColor(shape.lineColor);
    if (color === undefined && lost.indexOf('color') < 0) lost.push('color');
    look.line = { color: color || '000000', width };
    const dash = shape.lineDashing === undefined ? undefined : DASHES[shape.lineDashing];
    if (dash !== undefined) look.line.dash = dash;
  }
  if (lost.length) look.lost = lost;
  return look;
}
