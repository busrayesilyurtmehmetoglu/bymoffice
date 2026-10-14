# Looks at docx and xlsx files with nothing of the engine: Python's zipfile and ElementTree.
#
#   python3 -I package_check.py <folder>
#
# Every file of the folder is checked, and a JSON list of what is wrong is printed (empty: all is
# well). What is checked is what a file written half from one moment of its model and half from
# another gets wrong: the archive (its directory, its local headers and what it says of itself
# agree), every XML part parses, every relationship leads to a part, no part is left that nothing
# leads to, every part has a content type; and the indexes parts name each other by: styles, shared
# strings, formats of rules, relationship ids, comments, numbering, styles of paragraphs.

import json
import os
import posixpath
import re
import struct
import sys
import zipfile
import xml.etree.ElementTree as E

PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'
CT = '{http://schemas.openxmlformats.org/package/2006/content-types}'
R_IDS = ('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}', '{http://purl.oclc.org/ooxml/officeDocument/relationships}')


def local(tag):
    return tag.rsplit('}', 1)[-1]


def rels_name(part):
    head, tail = posixpath.split(part)
    return posixpath.join(head, '_rels', tail + '.rels')


def resolve(source, target):
    if target.startswith('/'):
        return posixpath.normpath(target[1:])
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def attr(el, name):
    for key, value in el.attrib.items():
        if local(key) == name:
            return value
    return None


def check(path, bad):
    name = os.path.basename(path)

    def say(text):
        bad.append(name + ': ' + text)

    raw = open(path, 'rb').read()
    z = zipfile.ZipFile(path)
    if z.testzip() is not None:
        say('a part does not unpack')
    names = z.namelist()
    if len(set(n.lower() for n in names)) != len(names):
        say('two entries have the same name')
    e = raw.rfind(b'PK\x05\x06')
    said = struct.unpack('<H', raw[e + 10:e + 12])[0]
    central = raw.count(b'PK\x01\x02')
    locals_ = raw.count(b'PK\x03\x04')
    if not (said == len(names) == central == locals_):
        say('the archive says %d entries, its directory has %d, %d are in it, %d are listed' % (said, central, locals_, len(names)))
    if '[Content_Types].xml' not in names or '_rels/.rels' not in names:
        say('no content types or relationships')
        return
    # A file named "zip-..." is looked at as an archive alone (the tests of the archive take entries out of a package).
    if name.startswith('zip-'):
        for n in names:
            if n.endswith('.xml') or n.endswith('.rels'):
                data = z.read(n)
                if len(data) == 0:
                    say(n + ' is empty')
                else:
                    E.fromstring(data)
        return
    lower = dict((n.lower(), n) for n in names)
    trees = {}
    for n in names:
        if n.endswith('.xml') or n.endswith('.rels') or n.endswith('.vml'):
            data = z.read(n)
            if len(data) == 0:
                say(n + ' is empty')
                continue
            try:
                trees[n] = E.fromstring(data)
            except Exception as error:
                say(n + ' does not parse: ' + str(error))
    # Content types: every part has one.
    types = trees.get('[Content_Types].xml')
    defaults = set()
    overrides = {}
    if types is not None:
        for child in types:
            if local(child.tag) == 'Default':
                defaults.add((child.get('Extension') or '').lower())
            elif local(child.tag) == 'Override':
                part = (child.get('PartName') or '').lower()
                if part in overrides:
                    say('two content types for ' + part)
                overrides[part] = child.get('ContentType')
        for n in names:
            if n == '[Content_Types].xml':
                continue
            if ('/' + n).lower() not in overrides and n.rsplit('.', 1)[-1].lower() not in defaults:
                say(n + ' has no content type')
        for part in overrides:
            if part[1:] not in lower:
                say('a content type for ' + part + ', which is not in the file')
    # Relationships: each leads to a part; ids are told apart; every part is led to.
    led = set(['[content_types].xml', '_rels/.rels'])
    rels = {}
    for n in names:
        if not n.endswith('.rels'):
            continue
        head, tail = posixpath.split(n)
        source = posixpath.join(posixpath.dirname(head), tail[:-5])
        if source != '' and source.lower() not in lower:
            say(n + ' holds the relationships of a part that is not in the file')
        tree = trees.get(n)
        if tree is None:
            continue
        ids = {}
        for rel in tree:
            if local(rel.tag) != 'Relationship':
                continue
            rid = rel.get('Id')
            if rid in ids:
                say(n + ' has the id ' + str(rid) + ' twice')
            if rel.get('TargetMode') == 'External':
                ids[rid] = (rel.get('Type') or '', None)
                continue
            target = resolve(source, rel.get('Target') or '')
            ids[rid] = (rel.get('Type') or '', target)
            if target.lower() not in lower:
                # A link to a place of the file itself ("#Sheet!A1") is no part.
                if not (rel.get('Target') or '').startswith('#'):
                    say(n + ' leads to ' + target + ', which is not in the file')
            else:
                led.add(target.lower())
                led.add(rels_name(target).lower())
        rels[source.lower()] = ids
    for n in names:
        if n.lower() not in led:
            say(n + ' is a part nothing leads to')

    def ids_of(part):
        return rels.get(part.lower(), {})

    def refs(part, tree, what):
        known = ids_of(part)
        for el in tree.iter():
            for key, value in el.attrib.items():
                if key.startswith(R_IDS) and value != '' and value not in known:
                    say(part + ' names the relationship ' + value + ' (' + local(el.tag) + '), which ' + what + ' does not have')

    main = None
    for rid, (kind, target) in ids_of('').items():
        if kind.endswith('/officeDocument'):
            main = target
    if main is None or main.lower() not in lower:
        say('no main part')
        return
    main = lower[main.lower()]
    main_type = overrides.get(('/' + main).lower(), '')
    if 'spreadsheetml' in main_type:
        book(z, main, trees, ids_of, refs, lower, say)
    elif 'wordprocessingml' in main_type:
        document(z, main, trees, ids_of, refs, lower, say)


def book(z, main, trees, ids_of, refs, lower, say):
    wb = trees.get(main)
    if wb is None:
        return
    refs(main, wb, 'the workbook')
    related = ids_of(main)
    sheets = [el for el in wb.iter() if local(el.tag) == 'sheet' and local(wb.tag) == 'workbook']
    sheet_parts = []
    seen_names = set()
    seen_ids = set()
    for sheet in sheets:
        rid = None
        for key, value in sheet.attrib.items():
            if key.startswith(R_IDS):
                rid = value
        entry = related.get(rid)
        if entry is None or entry[1] is None:
            say('the sheet ' + str(sheet.get('name')) + ' has no part')
            continue
        sheet_parts.append(entry[1])
        if (sheet.get('name') or '').lower() in seen_names:
            say('two sheets are named ' + str(sheet.get('name')))
        seen_names.add((sheet.get('name') or '').lower())
        if sheet.get('sheetId') in seen_ids:
            say('two sheets have the id ' + str(sheet.get('sheetId')))
        seen_ids.add(sheet.get('sheetId'))
    held = [target for (kind, target) in related.values() if target is not None and (kind.endswith('/worksheet') or kind.endswith('/chartsheet'))]
    if sorted(p.lower() for p in held) != sorted(p.lower() for p in sheet_parts):
        say('the workbook names %d sheets and relates to %d sheet parts' % (len(sheet_parts), len(held)))
    for el in wb.iter():
        if local(el.tag) == 'definedName' and el.get('localSheetId') is not None and int(el.get('localSheetId')) >= len(sheets):
            say('a defined name is local to sheet %s of %d' % (el.get('localSheetId'), len(sheets)))
        if local(el.tag) == 'workbookView' and el.get('activeTab') is not None and int(el.get('activeTab')) >= max(len(sheets), 1):
            say('the active sheet is %s of %d' % (el.get('activeTab'), len(sheets)))
    xfs = None
    dxfs = 0
    strings = None
    for (kind, target) in related.values():
        if target is None or target.lower() not in lower:
            continue
        tree = trees.get(lower[target.lower()])
        if tree is None:
            continue
        if kind.endswith('/styles'):
            fonts = fills = borders = fmts = None
            for child in tree:
                tag = local(child.tag)
                count = len([c for c in child if isinstance(c.tag, str)])
                if tag == 'cellXfs':
                    xfs = count
                    for xf in child:
                        for key, most in (('fontId', fonts), ('fillId', fills), ('borderId', borders)):
                            if xf.get(key) is not None and most is not None and int(xf.get(key)) >= most:
                                say('a cell format names %s %s of %d' % (key, xf.get(key), most))
                elif tag == 'dxfs':
                    dxfs = count
                elif tag == 'fonts':
                    fonts = count
                elif tag == 'fills':
                    fills = count
                elif tag == 'borders':
                    borders = count
        elif kind.endswith('/sharedStrings'):
            strings = len([c for c in tree if local(c.tag) == 'si'])
    if xfs is None:
        xfs = 1
    for part in sheet_parts:
        if part.lower() not in lower:
            continue
        part = lower[part.lower()]
        tree = trees.get(part)
        if tree is None or local(tree.tag) != 'worksheet':
            continue
        refs(part, tree, 'the sheet')
        cells = set()
        for el in tree.iter():
            tag = local(el.tag)
            if tag == 'c':
                if el.get('s') is not None and int(el.get('s')) >= xfs:
                    say('%s names style %s of %d' % (part, el.get('s'), xfs))
                if el.get('t') == 's':
                    v = [c for c in el if local(c.tag) == 'v']
                    if v and (strings is None or int(v[0].text) >= strings):
                        say('%s names shared string %s of %s' % (part, v[0].text, strings))
                if el.get('r') is not None:
                    if el.get('r') in cells:
                        say('%s has the cell %s twice' % (part, el.get('r')))
                    cells.add(el.get('r'))
            elif tag in ('row', 'col') and el.get('s') is not None and int(el.get('s')) >= xfs:
                say('%s names style %s of %d in a %s' % (part, el.get('s'), xfs, tag))
            elif tag == 'cfRule' and el.get('dxfId') is not None and int(el.get('dxfId')) >= dxfs:
                say('%s names the format %s of %d of a rule' % (part, el.get('dxfId'), dxfs))
        # The notes of the sheet and their boxes go together.
        kinds = [kind.rsplit('/', 1)[-1] for (kind, target) in ids_of(part).values()]
        legacy = [el for el in tree if local(el.tag) == 'legacyDrawing']
        if kinds.count('comments') > 0 and not legacy:
            say(part + ' has comments and names no drawing for their boxes')
        for (kind, target) in ids_of(part).values():
            if target is None or target.lower() not in lower:
                continue
            inner = trees.get(lower[target.lower()])
            if inner is None:
                continue
            if kind.endswith('/drawing'):
                refs(lower[target.lower()], inner, 'the drawing')
            if kind.endswith('/comments'):
                authors = len([a for a in inner.iter() if local(a.tag) == 'author'])
                for comment in inner.iter():
                    if local(comment.tag) == 'comment' and comment.get('authorId') is not None and int(comment.get('authorId')) >= authors:
                        say('%s names author %s of %d' % (target, comment.get('authorId'), authors))


def document(z, main, trees, ids_of, refs, lower, say):
    doc = trees.get(main)
    if doc is None:
        return
    refs(main, doc, 'the document')
    related = ids_of(main)
    by_kind = {}
    for (kind, target) in related.values():
        if target is not None and target.lower() in lower:
            by_kind.setdefault(kind.rsplit('/', 1)[-1], []).append(lower[target.lower()])
    for kind in ('header', 'footer', 'footnotes', 'endnotes', 'comments'):
        for part in by_kind.get(kind, []):
            if trees.get(part) is not None:
                refs(part, trees[part], 'the part')
    stories = [(main, doc)] + [(p, trees[p]) for kind in ('header', 'footer', 'footnotes', 'endnotes', 'comments') for p in by_kind.get(kind, []) if trees.get(p) is not None]
    # Comments: what the text names is there, once.
    comment_ids = set()
    para_ids = set()
    for part in by_kind.get('comments', []):
        tree = trees.get(part)
        if tree is None:
            continue
        for comment in tree:
            if local(comment.tag) != 'comment':
                continue
            cid = attr(comment, 'id')
            if cid in comment_ids:
                say('two comments have the id ' + str(cid))
            comment_ids.add(cid)
            for el in comment.iter():
                if local(el.tag) == 'p' and attr(el, 'paraId') is not None:
                    para_ids.add(attr(el, 'paraId').upper())
    for el in doc.iter():
        if local(el.tag) in ('commentReference', 'commentRangeStart', 'commentRangeEnd') and attr(el, 'id') not in comment_ids:
            say('the text names the comment %s (%s), which is not there' % (attr(el, 'id'), local(el.tag)))
    for part in by_kind.get('commentsExtended', []):
        tree = trees.get(part)
        if tree is None:
            continue
        for el in tree:
            if local(el.tag) != 'commentEx':
                continue
            for key in ('paraId', 'paraIdParent'):
                value = attr(el, key)
                if value is not None and value.upper() not in para_ids:
                    say('%s names the comment paragraph %s (%s), which is not there' % (part, value, key))
    # Notes.
    for kind, ref in (('footnotes', 'footnoteReference'), ('endnotes', 'endnoteReference')):
        ids = set()
        for part in by_kind.get(kind, []):
            if trees.get(part) is not None:
                for note in trees[part]:
                    ids.add(attr(note, 'id'))
        for el in doc.iter():
            if local(el.tag) == ref and attr(el, 'id') not in ids:
                say('the text names the note %s, which is not there' % attr(el, 'id'))
    # Numbering and styles.
    nums = None
    for part in by_kind.get('numbering', []):
        tree = trees.get(part)
        if tree is None:
            continue
        abstracts = set(attr(el, 'abstractNumId') for el in tree if local(el.tag) == 'abstractNum')
        nums = set()
        for el in tree:
            if local(el.tag) != 'num':
                continue
            if attr(el, 'numId') in nums:
                say('two lists have the id ' + str(attr(el, 'numId')))
            nums.add(attr(el, 'numId'))
            for child in el:
                if local(child.tag) == 'abstractNumId' and attr(child, 'val') not in abstracts:
                    say('the list %s names the definition %s, which is not there' % (attr(el, 'numId'), attr(child, 'val')))
    styles = None
    for part in by_kind.get('styles', []):
        tree = trees.get(part)
        if tree is not None:
            styles = set(attr(el, 'styleId') for el in tree if local(el.tag) == 'style')
    for part, tree in stories:
        for el in tree.iter():
            tag = local(el.tag)
            if tag == 'numId' and attr(el, 'val') not in (None, '0') and nums is not None and attr(el, 'val') not in nums:
                say('%s names the list %s, which is not there' % (part, attr(el, 'val')))
            elif tag == 'numId' and attr(el, 'val') not in (None, '0') and nums is None:
                say('%s names the list %s, and the file has no numbering' % (part, attr(el, 'val')))
            elif tag in ('pStyle', 'rStyle', 'tblStyle') and styles is not None and attr(el, 'val') not in styles:
                say('%s names the style %s, which is not there' % (part, attr(el, 'val')))
    # The ids of tracked changes are told apart within the document.
    marks = {}
    for el in doc.iter():
        if local(el.tag) in ('ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'pPrChange', 'tblPrChange', 'trPrChange', 'tcPrChange', 'sectPrChange', 'cellIns', 'cellDel', 'cellMerge', 'numberingChange', 'tblGridChange'):
            wid = attr(el, 'id')
            if wid is None:
                say('a tracked change (%s) has no id' % local(el.tag))
            marks[wid] = marks.get(wid, 0) + 1
    twice = [wid for wid, count in marks.items() if count > 1 and wid is not None]
    if twice:
        say('tracked changes share the ids ' + ', '.join(sorted(twice)[:5]))


def main():
    bad = []
    folder = sys.argv[1]
    for f in sorted(os.listdir(folder)):
        try:
            check(os.path.join(folder, f), bad)
        except Exception as error:
            bad.append(f + ': ' + str(error))
    print(json.dumps(bad))


main()
