// Open Packaging Conventions: the ZIP-based container behind docx and xlsx.
//
// Parts are addressed by absolute part names ("/xl/workbook.xml"). Parts the
// engine does not understand are never parsed; they stay as raw ZIP entries and
// are written back untouched.

import { ZipArchive } from './zip';
import { childElements, parseXml, serializeXml, type XmlDocument, type XmlElement, type XmlNode } from './xml';

const CONTENT_TYPES = '[Content_Types].xml';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  external: boolean;
  /** Resolved part name for internal targets, null for external ones. */
  part: string | null;
}

const encoder = new TextEncoder();

function resolveTarget(source: string, target: string): string {
  const base = target.startsWith('/') ? target : `${source.slice(0, source.lastIndexOf('/'))}/${target}`;
  const segments: string[] = [];
  for (const segment of base.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') segments.pop();
    else segments.push(segment);
  }
  return `/${segments.join('/')}`;
}

/** The part that holds the relationships of a part ("/": those of the package). */
export function relsPartFor(source: string): string {
  const slash = source.lastIndexOf('/');
  return `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

/** The relationships a relationships part states, for the part they belong to. */
export function relationshipsOf(source: string, doc: XmlDocument): Relationship[] {
  return childElements(doc.root, 'Relationship').map((rel) => {
    const target = rel.attrs.get('Target') ?? '';
    const external = rel.attrs.get('TargetMode') === 'External';
    return {
      id: rel.attrs.get('Id') ?? '',
      type: rel.attrs.get('Type') ?? '',
      target,
      external,
      part: external ? null : resolveTarget(source, target),
    };
  });
}

/**
 * What a list of content types says, whatever order it says it in: every child of the root
 * with its attributes. Two lists with the same signature declare the same types.
 */
function typesSignature(doc: XmlDocument): string {
  const lines: string[] = [];
  for (const node of doc.root.children) {
    if (node.type !== 'element') continue;
    const attrs = [...node.attrs].map(([name, value]) => `${name}=${JSON.stringify(value)}`).sort();
    // Anything inside a declaration (there should be nothing) makes it one of its own.
    lines.push(`${node.name} ${attrs.join(' ')}${node.children.length > 0 ? ` ${serializeXml({ ...doc, root: node })}` : ''}`);
  }
  return lines.sort().join('\n');
}

export class OpcPackage {
  private typesDirty = false;
  /** The content types part as the file had it, and what it declares: see save(). */
  private readonly typesBytes: Uint8Array;
  private readonly typesWere: string;
  /** The override of each deleted part, by its name in lower case: see writePart(). */
  private readonly dropped = new Map<string, XmlElement>();

  private constructor(
    readonly zip: ZipArchive,
    private types: XmlDocument,
    typesBytes: Uint8Array,
  ) {
    this.typesBytes = typesBytes;
    this.typesWere = typesSignature(types);
  }

  static async open(bytes: Uint8Array): Promise<OpcPackage> {
    const zip = ZipArchive.read(bytes);
    if (!zip.has(CONTENT_TYPES)) throw new Error('Not an OPC package: [Content_Types].xml is missing');
    const typesBytes = await zip.read(CONTENT_TYPES);
    return new OpcPackage(zip, parseXml(typesBytes), typesBytes);
  }

  // Part names are case-insensitive in OPC, ZIP entry names are not.
  private entryName(part: string): string | undefined {
    const name = part.slice(1);
    if (this.zip.has(name)) return name;
    const lower = name.toLowerCase();
    return this.zip.names().find((candidate) => candidate.toLowerCase() === lower);
  }

  partNames(): string[] {
    return this.zip
      .names()
      .filter((name) => name !== CONTENT_TYPES && !name.endsWith('/'))
      .map((name) => `/${name}`);
  }

  hasPart(part: string): boolean {
    return this.entryName(part) !== undefined;
  }

  contentTypeOf(part: string): string | undefined {
    const lower = part.toLowerCase();
    for (const override of childElements(this.types.root, 'Override')) {
      if (override.attrs.get('PartName')?.toLowerCase() === lower) return override.attrs.get('ContentType');
    }
    const extension = lower.slice(lower.lastIndexOf('.') + 1);
    for (const fallback of childElements(this.types.root, 'Default')) {
      if (fallback.attrs.get('Extension')?.toLowerCase() === extension) return fallback.attrs.get('ContentType');
    }
    return undefined;
  }

  /** Relationships of a part; pass "/" for the package-level ones. */
  async relationships(source = '/'): Promise<Relationship[]> {
    const relsPart = relsPartFor(source);
    if (!this.hasPart(relsPart)) return [];
    return relationshipsOf(source, await this.readXml(relsPart));
  }

  /** The main document part (word/document.xml, xl/workbook.xml, ...). */
  async mainPart(): Promise<string | undefined> {
    const rels = await this.relationships('/');
    return rels.find((rel) => rel.type.endsWith('/officeDocument'))?.part ?? undefined;
  }

  async readPart(part: string): Promise<Uint8Array> {
    const name = this.entryName(part);
    if (name === undefined) throw new Error(`Part not found: ${part}`);
    return this.zip.read(name);
  }

  async readXml(part: string): Promise<XmlDocument> {
    return parseXml(await this.readPart(part));
  }

  /**
   * Adds or replaces a part. contentType is only needed for parts not yet declared. A part that
   * is declared with another type keeps its one declaration, which gets the type given.
   */
  writePart(part: string, data: Uint8Array | XmlDocument, contentType?: string): void {
    const bytes = data instanceof Uint8Array ? data : encoder.encode(serializeXml(data));
    this.zip.set(this.entryName(part) ?? part.slice(1), bytes);
    if (contentType && this.contentTypeOf(part) !== contentType) {
      const root = this.types.root;
      const lower = part.toLowerCase();
      const declared = childElements(root, 'Override').find((override) => (override.attrs.get('PartName') || '').toLowerCase() === lower);
      // The declaration the part had before it was deleted comes back with everything it said.
      const override = declared || this.dropped.get(lower);
      if (override) {
        override.attrs.set('ContentType', contentType);
        if (!declared) root.children.push(override);
      } else {
        const prefix = root.name.includes(':') ? root.name.slice(0, root.name.indexOf(':') + 1) : '';
        root.children.push({
          type: 'element',
          name: `${prefix}Override`,
          attrs: new Map([
            ['PartName', part],
            ['ContentType', contentType],
          ]),
          children: [],
        });
      }
      this.typesDirty = true;
    }
    this.dropped.delete(part.toLowerCase());
  }

  /**
   * Removes a part, its own relationships part and its content type override.
   * Relationships that point to the part are the caller's to remove.
   */
  deletePart(part: string): boolean {
    const name = this.entryName(part);
    if (name === undefined) return false;
    this.zip.delete(name);
    const slash = part.lastIndexOf('/');
    const rels = this.entryName(`${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`);
    if (rels !== undefined) this.zip.delete(rels);
    const lower = part.toLowerCase();
    const root = this.types.root;
    const kept: XmlNode[] = [];
    for (const node of root.children) {
      if (node.type === 'element' && node.name.replace(/^.*:/, '') === 'Override' && (node.attrs.get('PartName') || '').toLowerCase() === lower) {
        if (!this.dropped.has(lower)) this.dropped.set(lower, node);
      } else kept.push(node);
    }
    if (kept.length !== root.children.length) {
      root.children = kept;
      this.typesDirty = true;
    }
    return true;
  }

  /**
   * The package as a file. Content types that declare what the file declared (parts were
   * added and removed again, in whatever order) are written with the bytes the file had,
   * so that a package that was changed and changed back is the file that was opened.
   */
  async save(): Promise<Uint8Array> {
    if (this.typesDirty) {
      const asInFile = typesSignature(this.types) === this.typesWere;
      this.zip.set(CONTENT_TYPES, asInFile ? this.typesBytes : encoder.encode(serializeXml(this.types)));
      this.typesDirty = false;
    }
    return this.zip.write();
  }
}
