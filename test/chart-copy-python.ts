// Saved files read by a reader that is not this engine (Python: zipfile, ElementTree and, where
// it is installed, openpyxl), for the tests of copied charts.

import { spawnSync } from 'node:child_process';
import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

// Reads packages without this engine. For each file: the archive opens and no two entries have
// one name; every part has a content type and every override a part; every relationship has an
// id of its own in its part and leads to a part that is there; every part is led to; every XML
// part is read by a parser that knows namespaces (an unbound prefix is an error there); every
// attribute that names a relationship names one the part has. With openpyxl: the workbook opens
// and each sheet has the charts that are expected.
const PYTHON = String.raw`
import json, sys, re, zipfile, posixpath, warnings
import xml.etree.ElementTree as ET
warnings.simplefilter('ignore')
try:
    import openpyxl
except Exception:
    openpyxl = None
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
RNS = ('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}', '{http://purl.oclc.org/ooxml/officeDocument/relationships}')

def rels_of(name):
    if name == '':
        return '_rels/.rels'
    return posixpath.join(posixpath.dirname(name), '_rels', posixpath.basename(name) + '.rels')

def check(path, expected):
    problems = []
    z = zipfile.ZipFile(path)
    names = z.namelist()
    folded = [n.lower() for n in names]
    if len(set(folded)) != len(folded):
        problems.append('two entries with one name')
    held = set(names)
    types = ET.fromstring(z.read('[Content_Types].xml'))
    defaults = set(); overrides = set()
    for el in types:
        if el.tag == CT + 'Default':
            defaults.add(el.get('Extension', '').lower())
        elif el.tag == CT + 'Override':
            part = el.get('PartName', '')[1:]
            if part in overrides: problems.append('two overrides for ' + part)
            overrides.add(part)
            if part not in held: problems.append('an override for ' + part + ', which is not in the package')
    reached = set(); queue = ['']; i = 0
    ids_of = {}
    while i < len(queue):
        source = queue[i]; i += 1
        rels = rels_of(source)
        if rels not in held: continue
        ids = set()
        for rel in ET.fromstring(z.read(rels)).findall(PKG + 'Relationship'):
            rid = rel.get('Id', '')
            if rid in ids: problems.append('%s: the id %s twice' % (rels, rid))
            ids.add(rid)
            if rel.get('TargetMode') == 'External': continue
            target = rel.get('Target', '')
            if target.startswith('#'): continue
            name = target[1:] if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
            if name not in held:
                problems.append('%s: %s leads to %s, which is not in the package' % (rels, rid, name))
            elif name not in reached:
                reached.add(name); queue.append(name)
        ids_of[source] = ids
    charts = 0
    for name in names:
        if name == '[Content_Types].xml': continue
        is_rels = re.search(r'(^|/)_rels/[^/]*\.rels$', name) is not None
        if is_rels:
            owner = re.sub(r'_rels/([^/]*)\.rels$', r'\1', name)
            if owner != '' and owner not in held: problems.append(name + ': the relationships of a part that is not in the package')
        elif name not in reached:
            problems.append(name + ': no relationship leads to it')
        ext = name.rsplit('.', 1)[-1].lower()
        if name not in overrides and ext not in defaults: problems.append(name + ': no content type')
        if not re.search(r'\.(xml|rels|vml)$', name): continue
        raw = z.read(name)
        # The parser honours the encoding a part declares: what is not well-formed in it is said below.
        try:
            root = ET.fromstring(raw)
        except ET.ParseError as e:
            problems.append('%s: %s' % (name, e)); continue
        if is_rels: continue
        if re.search(r'drawingml/(2006/)?chart\}chartSpace$', root.tag): charts += 1
        ids = ids_of.get(name, set())
        for el in root.iter():
            for key, value in el.attrib.items():
                if key.startswith(RNS) and key.split('}')[1] in ('id', 'embed', 'link', 'pict'):
                    if value not in ids: problems.append('%s: names the relationship %s, which it does not have' % (name, value))
    found = None
    if openpyxl is not None:
        try:
            wb = openpyxl.load_workbook(path)
            found = [len(ws._charts) for ws in wb.worksheets]
        except Exception as e:
            problems.append('openpyxl does not open it: %r' % (e,))
    if expected is not None:
        if charts != sum(expected): problems.append('%d chart parts, %d expected' % (charts, sum(expected)))
        if found is not None and found != expected: problems.append('openpyxl finds the charts %r, expected %r' % (found, expected))
    return problems

out = {'openpyxl': openpyxl is not None, 'files': 0, 'problems': []}
for path, expected in json.load(open(sys.argv[1])):
    out['files'] += 1
    try:
        for line in check(path, expected): out['problems'].append('%s: %s' % (posixpath.basename(path), line))
    except Exception as e:
        out['problems'].append('%s: %r' % (posixpath.basename(path), e))
print(json.dumps(out))
`;

/** Whether python3 runs here. */
export function pythonHere(): boolean {
  return spawnSync('python3', ['-c', 'import zipfile, xml.etree.ElementTree'], { encoding: 'utf8' }).status === 0;
}

/**
 * Has Python read the files. `charts` of a file: how many charts each of its worksheets has,
 * in the order of the workbook. Returns what it found wrong (one line each), how many files it
 * read and whether openpyxl took part.
 */
export function readByPython(files: { name: string; bytes: Uint8Array; charts?: number[] }[], label: string): { problems: string[]; files: number; openpyxl: boolean } {
  const folder = join(dirname(fileURLToPath(import.meta.url)), '..', `tmp-chart-copy-${label}-${process.pid}`);
  mkdirSync(folder, { recursive: true });
  try {
    const list: [string, number[] | null][] = [];
    files.forEach((file, i) => {
      const path = join(folder, `${i}-${file.name.replace(/[^A-Za-z0-9]+/g, '-').slice(0, 60)}.xlsx`);
      writeFileSync(path, file.bytes);
      list.push([path, file.charts || null]);
    });
    const listed = join(folder, 'files.json');
    writeFileSync(listed, JSON.stringify(list));
    const run = spawnSync('python3', ['-c', PYTHON, listed], { encoding: 'utf8', maxBuffer: 64 * 1024 * 1024 });
    if (run.status !== 0) return { problems: [`python failed: ${run.stderr.slice(-2000)}`], files: 0, openpyxl: false };
    return JSON.parse(run.stdout) as { problems: string[]; files: number; openpyxl: boolean };
  } finally {
    rmSync(folder, { recursive: true, force: true });
  }
}
