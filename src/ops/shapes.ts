// Shapes of a sheet as the editor changes them: a rectangle, a rounded rectangle, an oval, a line,
// an arrow and a text box added; moved, sized, turned and deleted; filled, outlined and given text;
// copied and pasted. Like the other commands these functions only build operations
// ("setDrawings", as for pictures: see pictures.ts, whose moving, turning and deleting they share);
// the caller applies them through a History so that each is one undo step.
//
// - A shape that is added is in no file: the writer writes it from its model (shape-write.ts).
// - A shape of the file the reader knew in full (DrawingShape.known: one the engine wrote, say) is
//   changed like one that was added, and written anew from its model.
// - Any other shape of the file is moved, sized, turned and deleted, and its text is replaced when
//   it is plain (DrawingShape.plainText); its fill, outline and the look of its text are what the
//   file says, and stay so: what the model does not hold of it (a style, effects, a gradient) is
//   never lost to a change. Its turn and its text are refused when the file holds it in another
//   form as well (SheetDrawing.alternative), whose other form would be left as it was.
// - A shape is copied when it is one the writer writes in full (added, or known); a copy names
//   nothing of the workbook it came from and is pasted on any sheet of any workbook.

import type { SheetMetrics } from '../layout/sheet-metrics';
import { createSheetMetrics } from '../layout/sheet-metrics';
import {
  MAX_COLS,
  MAX_ROWS,
  type CellRange,
  type DrawingAnchor,
  type DrawingParagraph,
  type DrawingShape,
  type DrawingText,
  type DrawingTextRun,
  type ShapeTextLook,
  type Sheet,
  type SheetDrawing,
  type Workbook,
} from '../model/sheet/types';
import { MAX_DRAWING_TEXT, textUpTo } from '../model/sheet/pictures';
import { anchorRect, drawingAnchorNow } from '../render/sheet-drawings';
import { checkedAnchor, MAX_SHEET_DRAWINGS } from './charts';
import { withCommand, type Op } from './ops';
import { anchorLike, changedDrawing, fileDrawing, MIN_PICTURE_SIZE, movedDrawing, objectsProtected, pictureRotation, removedDrawing, turnedDrawing } from './pictures';

/** The kinds of shape that are added. */
export const SHAPE_KINDS = ['rect', 'roundRect', 'ellipse', 'line', 'arrow', 'doubleArrow', 'textBox'] as const;
export type ShapeKind = (typeof SHAPE_KINDS)[number];

/** The size a shape is added at without a place, in pixels: a line is as wide and flat. UNSURE: Excel adds one of about an inch by a click. */
export const DEFAULT_SHAPE_WIDTH = 160;
export const DEFAULT_SHAPE_HEIGHT = 80;
/** The longest text of a shape, in UTF-16 code units: what the reader keeps of one (see drawings.ts). */
export const MAX_SHAPE_TEXT = 32767;
/** The most lines the text of a shape is given in. */
export const MAX_SHAPE_LINES = 2000;
/** The widths an outline is given, in points: from a hairline to a thick line. */
export const MIN_LINE_POINTS = 0.25;
export const MAX_LINE_POINTS = 24;
/** The sizes text of a shape is given, in points. */
export const MIN_TEXT_POINTS = 1;
export const MAX_TEXT_POINTS = 409;
/** How far a copy is put from a shape that stands where it would stand. */
export const SHAPE_COPY_OFFSET = 24;

const HEX = /^#[0-9A-Fa-f]{6}$/;
const POINT_PIXELS = 96 / 72;

/** A shape of a sheet as the commands see it. */
export interface ShapeInfo {
  /** The place of the shape in Sheet.drawings: what the commands name it by. It changes when a drawing before it is deleted. */
  id: number;
  anchor: DrawingAnchor;
  name?: string;
  description?: string;
  /** The preset geometry ("rect", "roundRect", "ellipse", "line", ...). */
  geometry: string;
  /** A line: it has two ends and no text. */
  line: boolean;
  rotation?: number;
  /** "#RRGGBB", when it has a fill that is drawn. */
  fill?: string;
  /** "#RRGGBB" and points, when it has an outline. */
  outline?: { color: string; width: number; head?: string; tail?: string };
  /** Its text, paragraphs one to a line. */
  text: string;
  /** The colour ("#RRGGBB") and the size (points) its text has, or text typed into it gets, when that is known. */
  textColor?: string;
  textSize?: number;
  /** The shape is in the file the workbook was read from. */
  ofFile: boolean;
  /** Whether it can be deleted (one that shares its place in the drawing with another object cannot). */
  removable: boolean;
  /** Whether its fill, its outline and the look of its text can be changed, and it can be copied: one that was added, or one of the file the engine knows in full. */
  formattable: boolean;
  /** Whether it can be turned. */
  turnable: boolean;
  /** Whether its text can be replaced. */
  textEditable: boolean;
}

/** What a shape is given when it is added; whatever is left out is what its kind starts with. */
export interface ShapeOptions {
  /** Its text, lines broken by "\n"; not of a line. */
  text?: string;
  /** "#RRGGBB", or null for none. */
  fill?: string | null;
  /** The colour of the outline ("#RRGGBB"), or null for none. */
  lineColor?: string | null;
  /** Points. */
  lineWidth?: number;
  /** "#RRGGBB". */
  textColor?: string;
  /** Points. */
  textSize?: number;
  name?: string;
  description?: string;
}

/** What formatShape() changes: each field that is there is set (null takes a fill or an outline away). */
export interface ShapeFormat {
  fill?: string | null;
  lineColor?: string | null;
  /** Points. */
  lineWidth?: number;
  textColor?: string;
  /** Points. */
  textSize?: number;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
}

// ------------------------------------------------------------------ helpers

function sheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) throw new RangeError(`No sheet at index ${sheetIndex}`);
  if (sheet.kind !== 'worksheet') throw new RangeError(`Cannot change the shapes of a ${sheet.kind}`);
  return sheet;
}

/** Why no shape can be added to a sheet, moved or deleted on it; undefined when that can be done. */
export function shapesProblem(workbook: Workbook, sheetIndex: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot change the shapes of a ${sheet.kind}`;
  if (sheet.drawingsFixed !== undefined && sheet.part !== '') return `Cannot change the shapes of this sheet: ${sheet.drawingsFixed}`;
  if (objectsProtected(sheet)) return 'Cannot change the shapes of this sheet: it is protected';
  return undefined;
}

function placedSheetOf(workbook: Workbook, sheetIndex: number): Sheet {
  const sheet = sheetOf(workbook, sheetIndex);
  const problem = shapesProblem(workbook, sheetIndex);
  if (problem) throw new RangeError(problem);
  return sheet;
}

function shapeAt(sheet: Sheet, id: number): { drawing: SheetDrawing; object: DrawingShape } {
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'shape') throw new RangeError(`No shape at index ${id}`);
  return { drawing, object: drawing.object };
}

/** A shape that the writer writes from its model: one that was added, or one of the file it knows in full. */
const ownShape = (drawing: SheetDrawing, object: DrawingShape): boolean => fileDrawing(drawing) === undefined || object.known === true;

/** The least size of a shape each way: a line may be flat. */
const leastOf = (object: DrawingShape): number => (object.connector ? 0 : MIN_PICTURE_SIZE);

function setDrawings(sheetIndex: number, list: SheetDrawing[]): Op[] {
  return [{ type: 'setDrawings', sheet: sheetIndex, drawings: list.length > 0 ? list : undefined }];
}

/** A colour as a command takes it: "#RRGGBB" in capitals. Throws RangeError otherwise. */
function checkedColor(color: unknown, what: string): string {
  if (typeof color !== 'string' || !HEX.test(color)) throw new RangeError(`The ${what} of a shape is a colour such as #FF0000`);
  return color.toUpperCase();
}

/** A size in points as a command takes it, in pixels; in the steps of the file (1/12,700 of a point for a line, 1/100 for text). */
function checkedPoints(points: unknown, low: number, high: number, what: string): number {
  if (typeof points !== 'number' || !isFinite(points) || points < low || points > high) throw new RangeError(`The ${what} of a shape is a number of points from ${low} to ${high}`);
  return points;
}

const lineWidthOf = (points: number): number => Math.round(points * 12700) / 9525;

function checkedName(text: unknown, what: string, limit: number): string | undefined {
  if (text === undefined || text === null) return undefined;
  const value = String(text).replace(/\r\n?/g, '\n');
  if (value.trim() === '') return undefined;
  if (value.length > limit) throw new RangeError(`The ${what} of a shape cannot be longer than ${limit} characters`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/.test(value)) throw new RangeError(`The ${what} of a shape holds a character a file cannot hold`);
  return value;
}

/** A text as a command takes it: line breaks as "\n", tabs kept; at most MAX_SHAPE_TEXT code units in MAX_SHAPE_LINES lines. */
export function checkedShapeText(text: unknown): string {
  if (text === undefined || text === null) return '';
  const value = String(text).replace(/\r\n?/g, '\n');
  if (value.length > MAX_SHAPE_TEXT) throw new RangeError(`The text of a shape cannot be longer than ${MAX_SHAPE_TEXT} characters`);
  let lines = 1;
  for (let i = value.indexOf('\n'); i >= 0 && lines <= MAX_SHAPE_LINES; i = value.indexOf('\n', i + 1)) lines++;
  if (lines > MAX_SHAPE_LINES) throw new RangeError(`The text of a shape cannot have more than ${MAX_SHAPE_LINES} lines`);
  if (/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value)) throw new RangeError('The text of a shape holds a character a file cannot hold');
  return value;
}

/** The look of the first run of a text; undefined for a text without one. */
function firstRunLook(text: DrawingText | undefined): Omit<DrawingTextRun, 'text'> | undefined {
  for (const paragraph of text ? text.paragraphs : []) {
    for (const run of paragraph.runs) {
      if (run.text === '\n') continue;
      const { text: _text, ...look } = run;
      void _text;
      return look;
    }
  }
  return undefined;
}

/** Whether every run of a text looks the same, and every paragraph is aligned the same: what a new text keeps of it (see textIn()). */
function oneLook(text: DrawingText | undefined): boolean {
  if (!text) return true;
  let look: string | undefined;
  let align: string | undefined;
  for (let p = 0; p < text.paragraphs.length; p++) {
    const paragraph = text.paragraphs[p];
    const own = paragraph.align === undefined ? '' : paragraph.align;
    if (align !== undefined && own !== align) return false;
    align = own;
    for (const run of paragraph.runs) {
      if (run.text === '\n') continue;
      const { text: _text, ...rest } = run;
      void _text;
      const key = JSON.stringify(rest, Object.keys(rest).sort());
      if (look !== undefined && key !== look) return false;
      look = key;
    }
  }
  return true;
}

/** What text typed into a shape gets: the look of its first run, of its text look, or of nothing. */
function lookOf(object: DrawingShape): ShapeTextLook {
  const base: ShapeTextLook = object.textLook ? { ...object.textLook, run: { ...object.textLook.run } } : { run: {} };
  const run = firstRunLook(object.text);
  if (run) base.run = { ...run };
  const text = object.text;
  if (text) {
    if (text.anchor !== undefined) base.anchor = text.anchor;
    else delete base.anchor;
    if (text.insets) base.insets = [text.insets[0], text.insets[1], text.insets[2], text.insets[3]];
    if (text.wrap === false) base.wrap = false;
    else delete base.wrap;
    const first = text.paragraphs[0];
    if (first && first.align !== undefined) base.align = first.align;
    else if (first) delete base.align;
  }
  return base;
}

/** A text of one look, lines of `value` its paragraphs; undefined for no text. */
function textIn(value: string, look: ShapeTextLook): DrawingText | undefined {
  if (value === '') return undefined;
  const paragraphs: DrawingParagraph[] = value.split('\n').map((line) => {
    const paragraph: DrawingParagraph = { runs: line === '' ? [] : [{ ...look.run, text: line }] };
    if (look.align !== undefined) paragraph.align = look.align;
    return paragraph;
  });
  const text: DrawingText = { paragraphs };
  if (look.anchor !== undefined) text.anchor = look.anchor;
  if (look.insets) text.insets = [look.insets[0], look.insets[1], look.insets[2], look.insets[3]];
  if (look.wrap === false) text.wrap = false;
  return text;
}

/** The look a shape of a kind starts with: what Excel gives a shape it adds, near enough (UNSURE: the colours of its default theme, without the style that refers to them). */
function startLook(kind: ShapeKind): { fill?: string; line?: { color: string; width: number }; text: ShapeTextLook } {
  const insets: [number, number, number, number] = [9.6, 4.8, 9.6, 4.8];
  if (kind === 'textBox') return { fill: '#FFFFFF', line: { color: '#000000', width: lineWidthOf(0.75) }, text: { run: { size: 11, color: '#000000' }, insets } };
  if (kind === 'line' || kind === 'arrow' || kind === 'doubleArrow') return { line: { color: '#4472C4', width: lineWidthOf(1) }, text: { run: { size: 11 } } };
  return { fill: '#4472C4', line: { color: '#2F528F', width: lineWidthOf(1) }, text: { run: { size: 11, color: '#FFFFFF' }, align: 'center', anchor: 'middle', insets } };
}

const NAMES: { [kind in ShapeKind]: string } = {
  rect: 'Rectangle',
  roundRect: 'Rectangle: Rounded Corners',
  ellipse: 'Oval',
  line: 'Straight Connector',
  arrow: 'Straight Arrow Connector',
  doubleArrow: 'Straight Arrow Connector',
  textBox: 'TextBox',
};

/** The name a shape that is added gets: the name of its kind and one more than the shapes the sheet has. */
function nextName(sheet: Sheet, base: string): string {
  let number = 1;
  for (const drawing of sheet.drawings || []) if (drawing && drawing.object && drawing.object.type === 'shape') number++;
  return `${base} ${number}`;
}

/** The model of a shape of a kind as it is added. */
function shapeOfKind(kind: ShapeKind): DrawingShape {
  const look = startLook(kind);
  const line = kind === 'line' || kind === 'arrow' || kind === 'doubleArrow';
  const shape: DrawingShape = { type: 'shape', geometry: kind === 'textBox' ? 'rect' : line ? (kind === 'line' ? 'line' : 'straightConnector1') : kind };
  if (line) shape.connector = true;
  if (kind === 'textBox') shape.textBox = true;
  if (look.fill) shape.fill = look.fill;
  if (look.line) {
    shape.line = { ...look.line };
    if (kind === 'arrow' || kind === 'doubleArrow') shape.line.tail = 'triangle';
    if (kind === 'doubleArrow') shape.line.head = 'triangle';
  }
  if (!line) shape.textLook = look.text;
  return shape;
}

/** Where a new shape of a size (pixels) stands at a cell: it moves and sizes with the cells under it, as Excel adds a shape. */
export function shapeAnchorAt(sheet: Sheet, row: number, col: number, width: number, height: number, min: number = MIN_PICTURE_SIZE): DrawingAnchor {
  const metrics = createSheetMetrics(sheet);
  const r = Math.max(0, Math.min(Math.floor(row) || 0, MAX_ROWS - 1));
  const c = Math.max(0, Math.min(Math.floor(col) || 0, MAX_COLS - 1));
  return anchorLike({ kind: 'twoCell' }, metrics, { x: metrics.cols.offset(c), y: metrics.rows.offset(r), width, height }, min);
}

/** The size a shape of a kind is added at without a place, in pixels. */
export function defaultShapeSize(kind: ShapeKind): { width: number; height: number } {
  return kind === 'line' || kind === 'arrow' || kind === 'doubleArrow' ? { width: DEFAULT_SHAPE_WIDTH, height: 0 } : { width: DEFAULT_SHAPE_WIDTH, height: DEFAULT_SHAPE_HEIGHT };
}

/** The look of a run with a format patch applied. */
function runWith(run: Omit<DrawingTextRun, 'text'>, format: ShapeFormat): Omit<DrawingTextRun, 'text'> {
  const out = { ...run };
  if (format.textColor !== undefined) out.color = format.textColor;
  if (format.textSize !== undefined) out.size = Math.round(format.textSize * 100) / 100;
  if (format.bold !== undefined) out.bold = format.bold;
  if (format.italic !== undefined) out.italic = format.italic;
  return out;
}

/** A shape with a format patch applied (see ShapeFormat), every run of its text included. */
function formatted(object: DrawingShape, format: ShapeFormat): DrawingShape {
  const next: DrawingShape = { ...object };
  if (format.fill !== undefined) {
    if (format.fill === null) delete next.fill;
    else next.fill = format.fill;
  }
  if (format.lineColor !== undefined || format.lineWidth !== undefined) {
    if (format.lineColor === null) delete next.line;
    else {
      const was = object.line;
      const color = format.lineColor !== undefined ? format.lineColor : was ? was.color : '#000000';
      const width = format.lineWidth !== undefined ? lineWidthOf(format.lineWidth) : was ? was.width : lineWidthOf(0.75);
      next.line = was ? { ...was, color, width } : { color, width };
    }
  }
  const touchesText = format.textColor !== undefined || format.textSize !== undefined || format.bold !== undefined || format.italic !== undefined || format.align !== undefined;
  if (touchesText && !object.connector) {
    const look = lookOf(object);
    look.run = runWith(look.run, format);
    if (format.align !== undefined) {
      if (format.align === 'left') delete look.align;
      else look.align = format.align;
    }
    next.textLook = look;
    if (object.text) {
      next.text = {
        ...object.text,
        paragraphs: object.text.paragraphs.map((paragraph) => {
          const made: DrawingParagraph = { ...paragraph, runs: paragraph.runs.map((run) => ({ ...runWith(run, format), text: run.text })) };
          if (format.align !== undefined) {
            if (format.align === 'left') delete made.align;
            else made.align = format.align;
          }
          return made;
        }),
      };
    }
  }
  return next;
}

/** A format patch as a command takes it. Throws RangeError for what it cannot be. */
function checkedFormat(format: unknown): ShapeFormat {
  if (!format || typeof format !== 'object' || Array.isArray(format)) throw new RangeError('The format of a shape is an object');
  const given = format as ShapeFormat;
  const out: ShapeFormat = {};
  if (given.fill !== undefined) out.fill = given.fill === null ? null : checkedColor(given.fill, 'fill');
  if (given.lineColor !== undefined) out.lineColor = given.lineColor === null ? null : checkedColor(given.lineColor, 'outline');
  if (given.lineWidth !== undefined) out.lineWidth = checkedPoints(given.lineWidth, MIN_LINE_POINTS, MAX_LINE_POINTS, 'outline width');
  if (given.textColor !== undefined) out.textColor = checkedColor(given.textColor, 'text colour');
  if (given.textSize !== undefined) out.textSize = checkedPoints(given.textSize, MIN_TEXT_POINTS, MAX_TEXT_POINTS, 'text size');
  if (given.bold !== undefined) {
    if (typeof given.bold !== 'boolean') throw new RangeError('Bold is true or false');
    out.bold = given.bold;
  }
  if (given.italic !== undefined) {
    if (typeof given.italic !== 'boolean') throw new RangeError('Italic is true or false');
    out.italic = given.italic;
  }
  if (given.align !== undefined) {
    if (given.align !== 'left' && given.align !== 'center' && given.align !== 'right') throw new RangeError('The text of a shape is aligned left, center or right');
    out.align = given.align;
  }
  return out;
}

// ------------------------------------------------------------------ problems

/**
 * Why a shape cannot be turned or given another text, as a message; undefined when it can. Besides
 * what keeps every shape of the sheet as it is (shapesProblem()): one the file holds in a wrapper
 * of alternatives, whose other form would be left as it was, and (for its text) a line, and a shape
 * of the file whose text is not plain (see DrawingShape.plainText).
 */
export function shapeEditProblem(workbook: Workbook, sheetIndex: number, id: number, what: 'turn' | 'text' = 'turn'): string | undefined {
  const problem = shapesProblem(workbook, sheetIndex);
  if (problem) return problem;
  const sheet = workbook.sheets[sheetIndex];
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'shape') return `No shape at index ${id}`;
  const object = drawing.object;
  const inFile = fileDrawing(drawing);
  if (inFile && inFile.alternative) return 'Cannot change this shape: the file holds it in another form as well, which would be left as it was';
  if (what === 'text') {
    if (object.connector) return 'A line has no text';
    if (inFile && object.known !== true && object.plainText !== true) return 'Cannot change the text of this shape: the file gives its text more than one look, a field or the text of a cell';
    // A new text is written in one look (lookOf()): a text of more looks, a known shape's or one built anew, would lose the others.
    if (!oneLook(object.text)) return 'Cannot change the text of this shape: its text has more than one look';
  }
  return undefined;
}

/**
 * Why the fill, the outline and the look of the text of a shape cannot be changed, and why it
 * cannot be copied, as a message; undefined when they can. Only a shape the writer writes from its
 * model is: one that was added, or one of the file the engine knows in full (DrawingShape.known).
 */
export function shapeFormatProblem(workbook: Workbook, sheetIndex: number, id: number): string | undefined {
  const problem = shapeEditProblem(workbook, sheetIndex, id);
  if (problem) return problem;
  const { drawing, object } = shapeAt(workbook.sheets[sheetIndex], id);
  if (!ownShape(drawing, object)) return 'Cannot change the look of this shape: the file holds more of it than can be kept';
  return undefined;
}

/** Why a shape cannot be copied (or cut), as a message; undefined when it can. See shapeFormatProblem(). */
export function shapeCopyProblem(workbook: Workbook, sheetIndex: number, id: number): string | undefined {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet) return `No sheet at index ${sheetIndex}`;
  if (sheet.kind !== 'worksheet') return `Cannot copy the shapes of a ${sheet.kind}`;
  const drawing = sheet.drawings && typeof id === 'number' && Math.floor(id) === id ? sheet.drawings[id] : undefined;
  if (!drawing || !drawing.object || drawing.object.type !== 'shape') return `No shape at index ${id}`;
  if (!ownShape(drawing, drawing.object)) return 'Cannot copy this shape: the file holds more of it than a copy keeps';
  return undefined;
}

/** The shape a command changes, with where it stands now; throws RangeError for what refuses `what` (see shapeEditProblem()). */
function shapeNow(workbook: Workbook, sheetIndex: number, id: number, what: 'turn' | 'text' | 'format') {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing, object } = shapeAt(sheet, id);
  const problem = what === 'format' ? shapeFormatProblem(workbook, sheetIndex, id) : shapeEditProblem(workbook, sheetIndex, id, what);
  if (problem) throw new RangeError(problem);
  return { sheet, drawing, object, now: drawingAnchorNow(drawing, sheet.structuralEdits) };
}

// ------------------------------------------------------------------ commands

/**
 * Adds a shape of a kind (see SHAPE_KINDS) to a sheet, above its other drawings. `anchor` is where
 * it stands (an anchor, or the cells it covers); without one it stands at cell A1 in the default
 * size of its kind (see defaultShapeSize()). It moves and sizes with its cells. Returns the
 * operations; the shape is the last entry of Sheet.drawings once they are applied. Throws
 * RangeError for a kind that is not added, options it cannot take, a sheet that takes no shape and
 * too many drawings.
 */
export function addShape(workbook: Workbook, sheetIndex: number, kind: ShapeKind, anchor?: DrawingAnchor | CellRange, options: ShapeOptions = {}): Op[] {
  return withCommand(addShapePlain(workbook, sheetIndex, kind, anchor, options), 'addShape', []);
}

function addShapePlain(workbook: Workbook, sheetIndex: number, kind: ShapeKind, anchor: DrawingAnchor | CellRange | undefined, options: ShapeOptions): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  if ((SHAPE_KINDS as readonly string[]).indexOf(kind) < 0) throw new RangeError('A shape is a rectangle, a rounded rectangle, an oval, a line, an arrow or a text box');
  const drawings = sheet.drawings || [];
  if (drawings.length >= MAX_SHEET_DRAWINGS) throw new RangeError('The sheet has too many drawings');
  const opts = options && typeof options === 'object' ? options : {};
  let object = shapeOfKind(kind);
  const format = checkedFormat({ fill: opts.fill, lineColor: opts.lineColor, lineWidth: opts.lineWidth, textColor: opts.textColor, textSize: opts.textSize });
  if (object.connector && format.fill !== undefined) throw new RangeError('A line has no fill');
  if (object.connector && format.lineColor === null) throw new RangeError('A line has an outline');
  object = formatted(object, format);
  const name = checkedName(opts.name, 'name', 255);
  object.name = name !== undefined ? name : nextName(sheet, NAMES[kind]);
  const description = checkedName(opts.description, 'alternative text', MAX_DRAWING_TEXT);
  if (description !== undefined) object.description = description;
  const text = checkedShapeText(opts.text);
  if (text !== '') {
    if (object.connector) throw new RangeError('A line has no text');
    object.text = textIn(text, lookOf(object));
  }
  let place: DrawingAnchor;
  if (anchor === undefined || anchor === null) {
    const size = defaultShapeSize(kind);
    place = shapeAnchorAt(sheet, 0, 0, size.width, size.height, leastOf(object));
  } else place = checkedAnchor(anchor, 'a shape');
  const added: SheetDrawing = { anchor: place, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  return setDrawings(sheetIndex, drawings.concat([added]));
}

/** Moves or sizes a shape: `anchor` is an anchor of the kind it has, or the cells it covers (see movePicture()). */
export function moveShape(workbook: Workbook, sheetIndex: number, id: number, anchor: DrawingAnchor | CellRange): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing, object } = shapeAt(sheet, id);
  return withCommand(movedDrawing(sheetIndex, sheet, id, drawing, anchor, 'a shape', leastOf(object)), 'moveShape', []);
}

/** Moves or sizes a shape to cover a rectangle of the sheet (pixels from the corner of cell A1), with the kind of anchor it has. */
export function moveShapeTo(workbook: Workbook, sheetIndex: number, id: number, rect: { x: number; y: number; width: number; height: number }, metrics?: SheetMetrics): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing, object } = shapeAt(sheet, id);
  const now = drawingAnchorNow(drawing, sheet.structuralEdits);
  return moveShape(workbook, sheetIndex, id, anchorLike(now, metrics || createSheetMetrics(sheet), rect, leastOf(object)));
}

/** Deletes a shape. Throws RangeError for one of the file that shares a wrapper of alternatives with another drawing. */
export function removeShape(workbook: Workbook, sheetIndex: number, id: number): Op[] {
  const sheet = placedSheetOf(workbook, sheetIndex);
  const { drawing } = shapeAt(sheet, id);
  return withCommand(removedDrawing(sheetIndex, sheet, id, drawing, 'shape'), 'removeShape', []);
}

/** Turns a shape to `degrees` (clockwise) about the middle of its box, as a picture is turned (see rotatePicture()). */
export function rotateShape(workbook: Workbook, sheetIndex: number, id: number, degrees: number): Op[] {
  const rotation = pictureRotation(degrees);
  const { sheet, drawing, object, now } = shapeNow(workbook, sheetIndex, id, 'turn');
  return withCommand(turnedDrawing(sheetIndex, sheet, id, drawing, object, now, createSheetMetrics(sheet), rotation, leastOf(object)), 'rotateShape', []);
}

/**
 * Gives a shape another text: lines broken by "\n" are its paragraphs, each in the look text typed
 * into the shape gets (see DrawingShape.textLook); empty takes the text away. Returns no operation
 * when the shape has that text already. Throws RangeError for a line, a shape of the file whose
 * text is not plain, and a text that is too long or holds what a file cannot hold.
 */
export function setShapeText(workbook: Workbook, sheetIndex: number, id: number, text: string | undefined): Op[] {
  const value = checkedShapeText(text);
  const { sheet, drawing, object, now } = shapeNow(workbook, sheetIndex, id, 'text');
  const before = object.text ? object.text.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n') : '';
  if (before === value) return [];
  const next: DrawingShape = { ...object };
  const made = textIn(value, lookOf(object));
  if (made) next.text = made;
  else delete next.text;
  return withCommand(changedDrawing(sheetIndex, sheet, id, drawing, next, now, undefined), 'setShapeText', []);
}

/**
 * Changes the fill, the outline and the look of the text of a shape (see ShapeFormat); null takes a
 * fill or an outline away. Returns no operation when nothing changes. Throws RangeError for a shape
 * whose look is not changed (see shapeFormatProblem()), a line given a fill or no outline, and a
 * value it cannot take.
 */
export function formatShape(workbook: Workbook, sheetIndex: number, id: number, format: ShapeFormat): Op[] {
  const patch = checkedFormat(format);
  const { sheet, drawing, object, now } = shapeNow(workbook, sheetIndex, id, 'format');
  if (object.connector && patch.fill !== undefined && patch.fill !== null) throw new RangeError('A line has no fill');
  if (object.connector && patch.lineColor === null) throw new RangeError('A line has an outline');
  const next = formatted(object, patch);
  if (JSON.stringify(next) === JSON.stringify(object)) return [];
  return withCommand(changedDrawing(sheetIndex, sheet, id, drawing, next, now, undefined), 'formatShape', []);
}

/** The shapes of a sheet, the one on top last. Shapes inside groups are not listed (they move with their group). */
export function sheetShapes(workbook: Workbook, sheetIndex: number): ShapeInfo[] {
  const sheet = workbook.sheets[sheetIndex];
  if (!sheet || sheet.kind !== 'worksheet' || !sheet.drawings) return [];
  const out: ShapeInfo[] = [];
  const locked = shapesProblem(workbook, sheetIndex) !== undefined;
  sheet.drawings.forEach((drawing, id) => {
    if (!drawing || !drawing.object || drawing.object.type !== 'shape') return;
    const object = drawing.object;
    const inFile = fileDrawing(drawing);
    const info: ShapeInfo = {
      id,
      anchor: drawingAnchorNow(drawing, sheet.structuralEdits),
      geometry: object.geometry,
      line: object.connector === true,
      text: object.text ? object.text.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join('')).join('\n') : '',
      ofFile: inFile !== undefined,
      removable: !locked && !(inFile && inFile.wrapped),
      formattable: !locked && shapeFormatProblem(workbook, sheetIndex, id) === undefined,
      turnable: !locked && shapeEditProblem(workbook, sheetIndex, id) === undefined,
      textEditable: !locked && shapeEditProblem(workbook, sheetIndex, id, 'text') === undefined,
    };
    if (object.name !== undefined) info.name = object.name;
    if (object.description !== undefined) info.description = object.description;
    if (object.rotation) info.rotation = object.rotation;
    if (object.fill) info.fill = object.fill;
    if (!object.connector) {
      const look = lookOf(object).run;
      if (typeof look.color === 'string' && HEX.test(look.color)) info.textColor = look.color;
      if (typeof look.size === 'number' && look.size > 0) info.textSize = look.size;
    }
    if (object.line) {
      info.outline = { color: object.line.color, width: Math.round(((object.line.width * 9525) / 12700) * 100) / 100 };
      if (object.line.head) info.outline.head = object.line.head;
      if (object.line.tail) info.outline.tail = object.line.tail;
    }
    out.push(info);
  });
  return out;
}

// ------------------------------------------------------------------ copies

/**
 * A shape taken off a sheet to be put down again: what copyShape() gives and pasteShape() takes.
 * It holds the model of the shape and nothing of the workbook it came from. Not to be altered.
 */
export interface ShapeCopy {
  readonly shape: Readonly<DrawingShape>;
  /** Its box as its anchor gives it, pixels from the corner of cell A1. */
  readonly rect: { x: number; y: number; width: number; height: number };
  readonly anchor: { kind: DrawingAnchor['kind']; editAs?: DrawingAnchor['editAs'] };
}

/** A deep copy of the model of a shape, without what the reader said of the file (DrawingShape.known, plainText). */
function copiedShape(object: DrawingShape): DrawingShape {
  const made = JSON.parse(JSON.stringify(object)) as DrawingShape;
  delete made.known;
  delete made.plainText;
  delete made.box;
  return made;
}

/** Takes a copy of a shape; nothing of the workbook is changed. Throws RangeError for a shape that cannot be copied (see shapeCopyProblem()). */
export function copyShape(workbook: Workbook, sheetIndex: number, id: number): ShapeCopy {
  const problem = shapeCopyProblem(workbook, sheetIndex, id);
  if (problem) throw new RangeError(problem);
  const sheet = workbook.sheets[sheetIndex];
  const { drawing, object } = shapeAt(sheet, id);
  const anchor = drawingAnchorNow(drawing, sheet.structuralEdits);
  const box = anchorRect(anchor, createSheetMetrics(sheet));
  return {
    shape: copiedShape(object),
    rect: { x: box.x, y: box.y, width: box.width, height: box.height },
    anchor: anchor.kind === 'twoCell' && anchor.editAs !== undefined ? { kind: anchor.kind, editAs: anchor.editAs } : { kind: anchor.kind },
  };
}

/** A copy of a shape and the operations that delete it (one undo step; the paste is another). Throws RangeError for a shape that cannot be copied or deleted. */
export function cutShape(workbook: Workbook, sheetIndex: number, id: number): { copy: ShapeCopy; ops: Op[] } {
  const ops = removeShape(workbook, sheetIndex, id);
  return { copy: copyShape(workbook, sheetIndex, id), ops };
}

/** What a copy of a shape that comes from anywhere is checked to be before it is put down: a shape the writer can write. */
function checkedCopy(value: unknown): DrawingShape {
  const given = value as DrawingShape;
  if (!given || typeof given !== 'object' || given.type !== 'shape' || typeof given.geometry !== 'string' || !/^[A-Za-z0-9]{1,64}$/.test(given.geometry)) throw new RangeError('What is pasted is not a shape that was copied');
  const made = copiedShape(given);
  const color = (c: unknown) => typeof c === 'string' && HEX.test(c);
  if (made.fill !== undefined && !color(made.fill)) throw new RangeError('What is pasted is not a shape that was copied');
  if (made.line !== undefined && !(made.line && color(made.line.color) && typeof made.line.width === 'number' && made.line.width > 0 && made.line.width <= 64)) throw new RangeError('What is pasted is not a shape that was copied');
  if (made.text !== undefined) {
    const value_ = made.text && Array.isArray(made.text.paragraphs) ? made.text.paragraphs.map((p) => (Array.isArray(p.runs) ? p.runs.map((r) => String(r.text)).join('') : '')).join('\n') : '';
    checkedShapeText(value_);
  }
  if (made.name !== undefined) checkedName(made.name, 'name', 255);
  if (made.description !== undefined) made.description = textUpTo(String(made.description), MAX_DRAWING_TEXT);
  if (made.rotation !== undefined) {
    const rotation = pictureRotation(made.rotation);
    if (rotation === 0) delete made.rotation;
    else made.rotation = rotation;
  }
  return made;
}

/**
 * Puts a copy of a shape on a sheet, above its other drawings, with the size, kind of anchor, turn,
 * fill, outline and text of the shape it was taken from, and a name of its own. With `at` its box
 * has its top left corner at the corner of that cell; without, it stands where the shape stood,
 * moved right and down (SHAPE_COPY_OFFSET) for as long as a shape stands there already. Throws
 * RangeError for what is not a copy of a shape, and a sheet that takes no more drawings.
 */
export function pasteShape(workbook: Workbook, sheetIndex: number, copy: ShapeCopy, at?: { row: number; col: number }): { ops: Op[]; id: number } {
  const sheet = placedSheetOf(workbook, sheetIndex);
  if (!copy || typeof copy !== 'object') throw new RangeError('What is pasted is not a shape that was copied');
  const object = checkedCopy(copy.shape);
  const drawings = sheet.drawings || [];
  if (drawings.length >= MAX_SHEET_DRAWINGS) throw new RangeError('The sheet has too many drawings');
  const number = (n: unknown, fallback: number): number => (typeof n === 'number' && isFinite(n) ? n : fallback);
  const rect = copy.rect && typeof copy.rect === 'object' ? copy.rect : { x: 0, y: 0, width: DEFAULT_SHAPE_WIDTH, height: DEFAULT_SHAPE_HEIGHT };
  const least = leastOf(object);
  const width = Math.max(number(rect.width, DEFAULT_SHAPE_WIDTH), least);
  const height = Math.max(number(rect.height, DEFAULT_SHAPE_HEIGHT), least);
  const metrics = createSheetMetrics(sheet);
  let x = Math.max(0, number(rect.x, 0));
  let y = Math.max(0, number(rect.y, 0));
  if (at !== undefined) {
    const whole = (n: unknown, limit: number) => typeof n === 'number' && n >= 0 && n < limit && Math.floor(n) === n;
    if (!at || !whole(at.row, MAX_ROWS) || !whole(at.col, MAX_COLS)) throw new RangeError('The place of a shape is a cell of the sheet');
    x = metrics.cols.offset(at.col);
    y = metrics.rows.offset(at.row);
  } else {
    const taken: { x: number; y: number }[] = [];
    for (const drawing of drawings) {
      if (!drawing || !drawing.object || drawing.object.type !== 'shape') continue;
      const box = anchorRect(drawingAnchorNow(drawing, sheet.structuralEdits), metrics);
      taken.push({ x: box.x, y: box.y });
    }
    for (let n = 0; n < 64 && taken.some((box) => Math.abs(box.x - x) < 1 && Math.abs(box.y - y) < 1); n++) {
      x += SHAPE_COPY_OFFSET;
      y += SHAPE_COPY_OFFSET;
    }
  }
  const kind = copy.anchor && (copy.anchor.kind === 'oneCell' || copy.anchor.kind === 'absolute') ? copy.anchor.kind : 'twoCell';
  const like: DrawingAnchor = { kind };
  if (kind === 'twoCell' && copy.anchor && (copy.anchor.editAs === 'oneCell' || copy.anchor.editAs === 'absolute')) like.editAs = copy.anchor.editAs;
  const anchor = anchorLike(like, metrics, { x, y, width, height }, least);
  const base = (object.name || 'Shape').replace(/\s+\d+$/, '') || 'Shape';
  object.name = nextName(sheet, base.slice(0, 240));
  const added: SheetDrawing = { anchor, object, since: sheet.structuralEdits ? sheet.structuralEdits.length : 0 };
  return { ops: withCommand(setDrawings(sheetIndex, drawings.concat([added])), 'pasteShape', []), id: drawings.length };
}
