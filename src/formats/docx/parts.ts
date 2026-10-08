// The small supporting parts of a document: theme, settings and font table.
//
// These parts stay in the package and are written back from there; what is
// read here is only what layout needs from them.

import {
  defaultSettings,
  emptyTheme,
  type DocSettings,
  type DocTheme,
  type FontInfo,
  type ThemeFont,
} from '../../model/doc/types';
import { localName, type XmlDocument, type XmlElement } from '../../xml';
import { parseThemeColors } from '../xlsx/styles';
import { anyAttr, attr, kindOf, onOff, parseOnOff, relAttr, utwipsAttr, type PartContext } from './common';

/** More fonts than this in a font table is not a genuine file. */
export const MAX_FONTS = 20000;

/** Order of the list parseThemeColors() returns. */
const SCHEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

/** ST_WmlColorSchemeIndex / ST_ThemeColor names -> colour scheme slots. */
const SCHEME_NAMES = new Map<string, string>([
  ['dark1', 'dk1'],
  ['light1', 'lt1'],
  ['dark2', 'dk2'],
  ['light2', 'lt2'],
  ['hyperlink', 'hlink'],
  ['followedHyperlink', 'folHlink'],
]);
for (let i = 1; i <= 6; i++) SCHEME_NAMES.set(`accent${i}`, `accent${i}`);

/** Theme colour names that go through the document's colour mapping. */
const MAPPED_NAMES = new Map<string, string>([
  ['text1', 't1'],
  ['background1', 'bg1'],
  ['text2', 't2'],
  ['background2', 'bg2'],
]);

function firstChild(parent: XmlElement | undefined, local: string): XmlElement | undefined {
  if (!parent) return undefined;
  for (const node of parent.children) {
    if (node.type === 'element' && localName(node.name) === local) return node;
  }
  return undefined;
}

function readThemeFont(el: XmlElement | undefined): ThemeFont {
  const font: ThemeFont = { scripts: new Map() };
  if (!el) return font;
  for (const node of el.children) {
    if (node.type !== 'element') continue;
    const typeface = node.attrs.get('typeface');
    if (!typeface) continue;
    const local = localName(node.name);
    if (local === 'latin') font.latin = typeface;
    else if (local === 'ea') font.eastAsia = typeface;
    else if (local === 'cs') font.complexScript = typeface;
    else if (local === 'font') {
      const script = node.attrs.get('script');
      if (script !== undefined && font.scripts.size < 1000) font.scripts.set(script, typeface);
    }
  }
  return font;
}

/** Colour and font scheme of the theme part. */
export function parseTheme(theme: XmlDocument | undefined): DocTheme {
  const result = emptyTheme();
  if (!theme) return result;
  parseThemeColors(theme).forEach((rgb, i) => {
    if (rgb) result.colors.set(SCHEME_SLOTS[i], rgb);
  });
  const fonts = firstChild(firstChild(theme.root, 'themeElements'), 'fontScheme');
  result.majorFont = readThemeFont(firstChild(fonts, 'majorFont'));
  result.minorFont = readThemeFont(firstChild(fonts, 'minorFont'));
  return result;
}

/**
 * The colour scheme slot ("dk1", "accent3", ...) a theme colour name of a
 * document refers to; undefined for names that are not theme colours.
 */
export function schemeSlot(settings: DocSettings, name: string): string | undefined {
  const mapped = MAPPED_NAMES.get(name);
  if (mapped !== undefined) return settings.colorMap.get(mapped);
  return SCHEME_NAMES.get(name);
}

/** Reads the root element of the settings part. */
const PROTECTION_KINDS: { [lower: string]: string } = { none: 'none', readonly: 'readOnly', comments: 'comments', trackedchanges: 'trackedChanges', forms: 'forms' };

/** w:edit of a protection as the schema spells it, whatever case the file wrote it in; a kind that is not known stays as written. */
function protectionKind(edit: string): string {
  const known = PROTECTION_KINDS[edit.trim().toLowerCase()];
  return known !== undefined ? known : edit;
}

export function parseSettings(ctx: PartContext, root: XmlElement | undefined): DocSettings {
  const settings = defaultSettings();
  if (!root) return settings;
  for (const node of root.children) {
    if (node.type !== 'element') continue;
    const kind = kindOf(ctx, node.name);
    switch (kind) {
      case 'defaultTabStop': {
        const value = utwipsAttr(ctx, node, 'val');
        if (value !== undefined) settings.defaultTabStop = value;
        break;
      }
      case 'evenAndOddHeaders':
      case 'mirrorMargins':
      case 'autoHyphenation':
      case 'trackRevisions': {
        const value = onOff(ctx, node);
        if (value !== undefined) settings[kind] = value;
        break;
      }
      case 'documentProtection': {
        // The first one that is enforced counts, or the first one when none is (an empty element in front must not
        // take the lock off); the part itself is never rewritten, so nothing of the others is lost.
        if (settings.protection && settings.protection.enforced) break;
        const edit = attr(ctx, node, 'edit');
        const enforcement = attr(ctx, node, 'enforcement');
        // Read as loosely as files write them: "TRUE", "On", " 1 ", "TrackedChanges".
        const enforced = enforcement !== undefined && parseOnOff(enforcement.trim().toLowerCase()) === true;
        if (settings.protection && !enforced) break;
        settings.protection = { edit: edit === undefined ? '' : protectionKind(edit), enforced, xml: node };
        break;
      }
      case 'writeProtection':
        if (!settings.writeProtection) settings.writeProtection = node;
        break;
      case 'themeFontLang': {
        settings.themeFontLang = {};
        for (const name of ['val', 'eastAsia', 'bidi'] as const) {
          const value = attr(ctx, node, name);
          if (value !== undefined) settings.themeFontLang[name] = value;
        }
        break;
      }
      case 'clrSchemeMapping':
        for (const name of ['bg1', 't1', 'bg2', 't2']) {
          const slot = SCHEME_NAMES.get(attr(ctx, node, name) || '');
          if (slot !== undefined) settings.colorMap.set(name, slot);
        }
        break;
      case 'compat':
        for (const inner of node.children) {
          if (inner.type !== 'element' || settings.compat.size + settings.compatSettings.size >= 1000) continue;
          const name = kindOf(ctx, inner.name);
          if (name === 'compatSetting') {
            const key = attr(ctx, inner, 'name');
            const value = attr(ctx, inner, 'val');
            if (key !== undefined && value !== undefined) settings.compatSettings.set(key, value);
          } else {
            const value = onOff(ctx, inner);
            if (value !== undefined) settings.compat.set(name, value);
          }
        }
        break;
    }
  }
  return settings;
}

const EMBED_FACES = new Map<string, string>([
  ['embedRegular', 'regular'],
  ['embedBold', 'bold'],
  ['embedItalic', 'italic'],
  ['embedBoldItalic', 'boldItalic'],
]);

/** Reads the root element of the font table part. */
export function parseFontTable(ctx: PartContext, root: XmlElement | undefined): Map<string, FontInfo> {
  const fonts = new Map<string, FontInfo>();
  if (!root) return fonts;
  for (const node of root.children) {
    if (node.type !== 'element' || kindOf(ctx, node.name) !== 'font') continue;
    const name = attr(ctx, node, 'name');
    if (name === undefined || fonts.has(name)) continue;
    if (fonts.size >= MAX_FONTS) throw new Error(`Invalid docx: more than ${MAX_FONTS} fonts in the font table`);
    const font: FontInfo = { name };
    for (const inner of node.children) {
      if (inner.type !== 'element') continue;
      const kind = kindOf(ctx, inner.name);
      const face = EMBED_FACES.get(kind);
      if (face !== undefined) {
        const rId = relAttr(ctx, inner, 'id') || anyAttr(inner, 'id');
        if (rId !== undefined) (font.embedded || (font.embedded = new Map())).set(face, rId);
        continue;
      }
      const value = attr(ctx, inner, 'val');
      if (value === undefined) continue;
      if (kind === 'altName') font.altName = value;
      else if (kind === 'panose1') font.panose = value;
      else if (kind === 'charset') font.charset = value;
      else if (kind === 'family') font.family = value;
      else if (kind === 'pitch') font.pitch = value;
    }
    fonts.set(name, font);
  }
  return fonts;
}
