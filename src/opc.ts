// Open Packaging Conventions: the ZIP-based container behind docx and xlsx.
//
// Parts are addressed by absolute part names ("/xl/workbook.xml"). Parts the
// engine does not understand are never parsed; they stay as raw ZIP entries and
// are written back untouched.

import { ZipArchive } from './zip';
import { childElements, parseXml, serializeXml, type XmlDocument } from './xml';

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

export class OpcPackage {
  private typesDirty = false;

  private constructor(
    readonly zip: ZipArchive,
    private types: XmlDocument,
  ) {}

  static async open(bytes: Uint8Array): Promise<OpcPackage> {
    const zip = ZipArchive.read(bytes);
    if (!zip.has(CONTENT_TYPES)) throw new Error('Not an OPC package: [Content_Types].xml is missing');
    return new OpcPackage(zip, parseXml(await zip.read(CONTENT_TYPES)));
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
    const slash = source.lastIndexOf('/');
    const relsPart = `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
    if (!this.hasPart(relsPart)) return [];
    const doc = await this.readXml(relsPart);
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

  /** Adds or replaces a part. contentType is only needed for parts not yet declared. */
  writePart(part: string, data: Uint8Array | XmlDocument, contentType?: string): void {
    const bytes = data instanceof Uint8Array ? data : encoder.encode(serializeXml(data));
    this.zip.set(this.entryName(part) ?? part.slice(1), bytes);
    if (contentType && this.contentTypeOf(part) !== contentType) {
      const root = this.types.root;
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
      this.typesDirty = true;
    }
  }

  async save(): Promise<Uint8Array> {
    if (this.typesDirty) {
      this.zip.set(CONTENT_TYPES, encoder.encode(serializeXml(this.types)));
      this.typesDirty = false;
    }
    return this.zip.write();
  }
}
