# Checks the pivot tables of xlsx files with nothing of the engine: Python's zipfile and ElementTree.
#
#   python3 pivot_check.py <folder> [--require-openpyxl]
#
# Every .xlsx of the folder is checked, and a JSON list of what is wrong is printed (empty: all is
# well). For each pivot table: its records are the rows of its source as the cells have them; the
# items of each field stand for every value of the cache once; the size of the table is what its
# lines of rows and columns say; and every cell of values is what the records of its row and column
# add up to, worked out here from the records by their own sums, counts, averages, ... When
# openpyxl is installed, the file is also opened with it and has to show as many pivot tables.

import json
import math
import os
import posixpath
import re
import statistics
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as E
from datetime import datetime, timedelta

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
PKG = '{http://schemas.openxmlformats.org/package/2006/relationships}'


def local(tag):
    return tag.rsplit('}', 1)[-1]


def kids(el, name):
    return [k for k in el if local(k.tag) == name]


def kid(el, name):
    found = kids(el, name)
    return found[0] if found else None


def rels_of(z, part):
    head, tail = posixpath.split(part)
    name = posixpath.join(head, '_rels', tail + '.rels')
    if name not in z.namelist():
        return {}
    out = {}
    for rel in E.fromstring(z.read(name)):
        target = rel.get('Target')
        if rel.get('TargetMode') == 'External':
            continue
        path = target[1:] if target.startswith('/') else posixpath.normpath(posixpath.join(head, target))
        out[rel.get('Id')] = (rel.get('Type').rsplit('/', 1)[-1], path)
    return out


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def parse_ref(ref):
    m = re.match(r'^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$', ref)
    top, left = int(m.group(2)) - 1, col_index(m.group(1))
    if m.group(3):
        return top, left, int(m.group(4)) - 1, col_index(m.group(3))
    return top, left, top, left


def serial_of_iso(text):
    d = datetime.strptime(text[:19], '%Y-%m-%dT%H:%M:%S')
    delta = d - datetime(1899, 12, 30)
    return delta.days + delta.seconds / 86400.0


def read_sheet(z, part, strings):
    cells = {}
    root = E.fromstring(z.read(part))
    data = kid(root, 'sheetData')
    for row in kids(data, 'row'):
        for c in kids(row, 'c'):
            m = re.match(r'^([A-Z]+)(\d+)$', c.get('r'))
            r, col = int(m.group(2)) - 1, col_index(m.group(1))
            t = c.get('t', 'n')
            v = kid(c, 'v')
            if t == 'inlineStr':
                value = ('s', ''.join(x.text or '' for x in c.iter() if local(x.tag) == 't'))
            elif v is None:
                continue
            elif t == 's':
                value = ('s', strings[int(v.text)])
            elif t == 'str':
                value = ('s', v.text or '')
            elif t == 'b':
                value = ('b', v.text == '1')
            elif t == 'e':
                value = ('e', v.text)
            else:
                value = ('n', float(v.text))
            cells[(r, col)] = (value, int(c.get('s', '0')))
    return cells


def cache_value(el):
    t = local(el.tag)
    v = el.get('v')
    if t == 'n':
        return ('n', float(v))
    if t == 'd':
        return ('n', serial_of_iso(v))
    if t == 'b':
        return ('b', v in ('1', 'true'))
    if t == 'm':
        return ('m', None)
    return (t, v)


def key(value):
    t, v = value
    if t == 's':
        return ('s', v.lower())
    return value


def aggregate(fn, values):
    # A result past the largest number is the error #NUM!, as Excel shows it.
    try:
        out = aggregate_plain(fn, values)
    except (OverflowError, ValueError):
        return '#NUM!'
    if isinstance(out, float) and not math.isfinite(out):
        return '#NUM!'
    return out


def aggregate_plain(fn, values):
    filled = [v for v in values if v[0] != 'm']
    if fn == 'count':
        return len(filled) if filled else None
    errors = [v for v in filled if v[0] == 'e']
    if errors:
        return errors[0][1]
    nums = [v[1] for v in filled if v[0] == 'n']
    if fn == 'sum':
        return math.fsum(nums)
    if fn == 'countNums':
        return len(nums)
    if fn == 'average':
        return math.fsum(nums) / len(nums) if nums else '#DIV/0!'
    if fn == 'max':
        return max(nums) if nums else 0
    if fn == 'min':
        return min(nums) if nums else 0
    if fn == 'product':
        out = 1.0
        for n in nums:
            out *= n
        return out if nums else 0
    if fn in ('var', 'stdDev'):
        if len(nums) < 2:
            return '#DIV/0!'
        var = statistics.variance(nums)
        return var if fn == 'var' else math.sqrt(var)
    if fn in ('varp', 'stdDevp'):
        if len(nums) < 1:
            return '#DIV/0!'
        var = statistics.pvariance(nums)
        return var if fn == 'varp' else math.sqrt(var)
    raise ValueError(fn)


# The order of text, worked out here on its own from the decomposed text (NFD) and a table of the
# Turkish alphabet written out by hand: a letter is a base character, followed by its marks. The
# marks that make the Turkish letters Ç Ğ İ Ö Ş Ü out of C G I O S U make a letter of its own; every
# other mark (category M) is an accent of the letter before it. Levels: the letters (case folded the
# Turkish way: I is ı, İ is i; Latin letters by the table, ASCII signs and numerals before them,
# other characters after them by their lower case), then the accents, then upper case after lower.
TURKISH = ['a', 'b', 'c', 'ç', 'd', 'e', 'f', 'g', 'ğ', 'h', 'ı', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'ö', 'p', 'q', 'r', 's', 'ş', 't', 'u', 'ü', 'v', 'w', 'x', 'y', 'z']
RANK = {letter: n for n, letter in enumerate(TURKISH)}
# Base letter (either case) and the mark that makes a Turkish letter of it.
MAKES = {('c', '\u0327'): 'ç', ('g', '\u0306'): 'ğ', ('o', '\u0308'): 'ö', ('s', '\u0327'): 'ş', ('u', '\u0308'): 'ü'}


def letters(text):
    """The letters of a text: [base character, marks]."""
    out = []
    for ch in unicodedata.normalize('NFD', text):
        if unicodedata.category(ch).startswith('M') or ch == '͏':
            if out:
                out[-1][1].append(ch)
            else:
                out.append(['', [ch]])
        else:
            out.append([ch, []])
    return out


def text_key(text):
    first, second, third = [], [], []
    for base, marks in letters(text):
        upper = base != '' and (base.isupper() or base in 'Iİ')
        # A bare I is the capital of ı; with an accent (Í, Ï) it is the capital of the accented i, as lower case has it.
        folded = 'ı' if base == 'I' and not [m for m in marks if m != '\u0307'] else base.lower()
        accents = list(marks)
        if base == 'I' and '\u0307' in accents:
            # The capital I with a dot above is İ, the capital of i: the dot is part of the letter. (A dot on i or ı is an accent.)
            folded = 'i'
            accents.remove('\u0307')
        else:
            for mark in accents:
                if (folded, mark) in MAKES:
                    folded = MAKES[(folded, mark)]
                    accents.remove(mark)
                    break
        if folded in RANK:
            first.append((1, RANK[folded]))
        elif base == '':
            pass
        elif ord(base) < 0x80:
            first.append((0, ord(base)))
        else:
            first.append((2, ord(unicodedata.normalize('NFD', base.lower())[0])))
        second.append(tuple(ord(m) for m in accents))
        third.append(1 if upper else 0)
    return (first, second, third, text)


def sort_key(value):
    t, v = value
    if t == 'n':
        return (0, v)
    if t == 's':
        return (1, text_key(v))
    if t == 'b':
        return (2, v)
    if t == 'e':
        return (3, v)
    return (4, 0)


def close(a, b):
    if isinstance(a, str) or isinstance(b, str):
        return a == b
    return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))


def check_table(z, name, part, sheet_cells, sheets, strings, bad):
    root = E.fromstring(z.read(part))
    where = '%s %s' % (name, part)
    rels = rels_of(z, part)
    cache_part = [p for (kind, p) in rels.values() if kind == 'pivotCacheDefinition'][0]
    cache = E.fromstring(z.read(cache_part))
    fields = kids(kid(cache, 'cacheFields'), 'cacheField')
    shared = []
    for f in fields:
        items = kid(f, 'sharedItems')
        shared.append([cache_value(x) for x in items] if items is not None and len(items) > 0 else None)
    crels = rels_of(z, cache_part)
    records_parts = [p for (kind, p) in crels.values() if kind == 'pivotCacheRecords']
    if not records_parts:
        return
    records = []
    for r in E.fromstring(z.read(records_parts[0])):
        if local(r.tag) != 'r':
            continue
        row = []
        for f, x in enumerate(r):
            if local(x.tag) == 'x':
                i = int(x.get('v', '0'))
                if shared[f] is None or i >= len(shared[f]):
                    bad.append('%s: a record names item %d of field %d, which does not list it' % (where, i, f))
                    return
                row.append(shared[f][i])
            else:
                row.append(cache_value(x))
        records.append(row)
    rec_root = E.fromstring(z.read(records_parts[0]))
    if int(rec_root.get('count', '-1')) != len(records):
        bad.append('%s: the records say there are %s, there are %d' % (where, rec_root.get('count'), len(records)))
    if cache.get('recordCount') is not None and int(cache.get('recordCount')) != len(records):
        bad.append('%s: the cache says it has %s records, there are %d' % (where, cache.get('recordCount'), len(records)))
    # The records are the rows of the source.
    src = kid(kid(cache, 'cacheSource'), 'worksheetSource')
    if src.get('ref') and src.get('sheet') in sheets:
        top, left, bottom, right = parse_ref(src.get('ref'))
        cells = sheets[src.get('sheet')]
        if bottom - top != len(records):
            bad.append('%s: the source has %d rows, the cache %d records' % (where, bottom - top, len(records)))
        for i, rec in enumerate(records):
            for f, value in enumerate(rec):
                got = cells.get((top + 1 + i, left + f))
                want = ('m', None) if got is None else got[0]
                if key(want) != key(value) and not (want[0] == 'n' and value[0] == 'n' and close(want[1], value[1])):
                    bad.append('%s: record %d field %d is %r, the source has %r' % (where, i, f, value, want))
                    return
    # The items of each field.
    pfields = kids(kid(root, 'pivotFields'), 'pivotField')
    items = []
    for f, pf in enumerate(pfields):
        el = kid(pf, 'items')
        if el is None:
            items.append(None)
            continue
        lst = kids(el, 'item')
        if int(el.get('count', '-1')) != len(lst):
            bad.append('%s: field %d says it has %s items, it has %d' % (where, f, el.get('count'), len(lst)))
        xs = [int(i.get('x')) for i in lst if i.get('t') in (None, 'data')]
        if shared[f] is not None and sorted(xs) != list(range(len(shared[f]))):
            bad.append('%s: the items of field %d do not stand for every value of the cache once (%r of %d)' % (where, f, xs, len(shared[f])))
        items.append(lst)
    row_fields = [int(x.get('x')) for x in kids(kid(root, 'rowFields'), 'field')] if kid(root, 'rowFields') is not None else []
    col_fields = [int(x.get('x')) for x in kids(kid(root, 'colFields'), 'field')] if kid(root, 'colFields') is not None else []
    data_fields = [(int(d.get('fld')), d.get('subtotal', 'sum')) for d in kids(kid(root, 'dataFields'), 'dataField')]
    loc = kid(root, 'location')
    top, left, bottom, right = parse_ref(loc.get('ref'))
    fdr, fdc = int(loc.get('firstDataRow')), int(loc.get('firstDataCol'))

    def lines(el):
        out, prev = [], []
        for i in kids(el, 'i'):
            r = int(i.get('r', '0'))
            path = prev[:r] + [int(x.get('v', '0')) for x in kids(i, 'x')]
            out.append((i.get('t', 'data'), path, int(i.get('i', '0'))))
            prev = path
        return out

    row_lines = lines(kid(root, 'rowItems'))
    col_lines = lines(kid(root, 'colItems'))
    if bottom - top + 1 != fdr + len(row_lines):
        bad.append('%s: %d rows in %s, %d lines of rows below row %d' % (where, bottom - top + 1, loc.get('ref'), len(row_lines), fdr))
    if right - left + 1 != fdc + len(col_lines):
        bad.append('%s: %d columns in %s, %d lines of columns after column %d' % (where, right - left + 1, loc.get('ref'), len(col_lines), fdc))
    values_in_cols = -2 in col_fields
    col_real = [f for f in col_fields if f != -2]

    def value_of_item(f, i):
        it = items[f][i]
        return shared[f][int(it.get('x'))]

    hidden = {}
    for f, lst in enumerate(items):
        if lst is None:
            continue
        hidden[f] = set(key(shared[f][int(i.get('x'))]) for i in lst if i.get('h') in ('1', 'true') and i.get('x') is not None)
    pages = []
    if kid(root, 'pageFields') is not None:
        for p in kids(kid(root, 'pageFields'), 'pageField'):
            f = int(p.get('fld'))
            pages.append((f, key(value_of_item(f, int(p.get('item')))) if p.get('item') is not None else None))

    def shown(rec):
        for f in row_fields + col_real:
            if key(rec[f]) in hidden.get(f, set()):
                return False
        for f, only in pages:
            if only is not None and key(rec[f]) != only:
                return False
            if only is None and key(rec[f]) in hidden.get(f, set()):
                return False
        return True

    visible = [rec for rec in records if shown(rec)]

    def match(rec, fields_, path):
        for k, i in enumerate(path):
            if key(rec[fields_[k]]) != key(value_of_item(fields_[k], i)):
                return False
        return True

    show_missing = root.get('showMissing', '1') not in ('0', 'false')
    missing_caption = root.get('missingCaption') if show_missing else None
    error_caption = (root.get('errorCaption') or '') if root.get('showError', '0') in ('1', 'true') else None
    # Sorted fields: their items in the order of their values.
    for f, pf in enumerate(pfields):
        sort = pf.get('sortType', 'manual')
        if sort not in ('ascending', 'descending') or items[f] is None or shared[f] is None:
            continue
        values = [shared[f][int(i.get('x'))] for i in items[f] if i.get('t') in (None, 'data')]
        wanted = sorted(values, key=sort_key, reverse=(sort == 'descending'))
        if [sort_key(v) for v in values] != [sort_key(v) for v in wanted]:
            bad.append('%s: the items of field %d are not in the order of %s: %r' % (where, f, sort, [v[1] for v in values]))
    for ri, (rt, rpath, _) in enumerate(row_lines):
        for ci, (ct, cpath, cdata) in enumerate(col_lines):
            r, c = top + fdr + ri, left + fdc + ci
            got = sheet_cells.get((r, c))
            d = cdata if values_in_cols else 0
            fld, fn = data_fields[d]
            rp = [] if rt == 'grand' else rpath
            cp = [] if (ct == 'grand' or values_in_cols) else cpath
            group = [rec[fld] for rec in visible if match(rec, row_fields, rp) and match(rec, col_real, cp)]
            n_rows = len([rec for rec in visible if match(rec, row_fields, rp) and match(rec, col_real, cp)])
            want = aggregate(fn, group) if n_rows > 0 else None
            if got is None:
                # A line above others of the rows may leave its values empty; a crossing without records is empty.
                if want is not None and not (rt == 'data' and len(rpath) < len(row_fields)):
                    bad.append('%s: the cell at row %d column %d is empty, %r was expected' % (where, r + 1, c + 1, want))
                continue
            value = got[0]
            # What the table shows in place of a missing value or of an error, when it says so.
            if want is None and missing_caption is not None and value == ('s', missing_caption):
                continue
            if isinstance(want, str) and error_caption is not None:
                if value != ('s', error_caption):
                    bad.append('%s: the cell at row %d column %d holds %r, the caption of errors %r was expected' % (where, r + 1, c + 1, value, error_caption))
                continue
            if want is None:
                bad.append('%s: the cell at row %d column %d holds %r, nothing was expected' % (where, r + 1, c + 1, value))
            elif value[0] == 'e' and want != value[1]:
                bad.append('%s: the cell at row %d column %d holds %r, %r was expected' % (where, r + 1, c + 1, value, want))
            elif value[0] == 'n' and (isinstance(want, str) or not close(value[1], want)):
                bad.append('%s: the cell at row %d column %d holds %r, %r was expected' % (where, r + 1, c + 1, value[1], want))
            elif value[0] not in ('n', 'e'):
                bad.append('%s: the cell at row %d column %d holds %r, a value was expected' % (where, r + 1, c + 1, value))


REQUIRE_OPENPYXL = '--require-openpyxl' in sys.argv


def check(path, bad):
    name = os.path.basename(path)
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        strings = []
        if 'xl/sharedStrings.xml' in names:
            for si in E.fromstring(z.read('xl/sharedStrings.xml')):
                strings.append(''.join(t.text or '' for t in si.iter() if local(t.tag) == 't'))
        book = E.fromstring(z.read('xl/workbook.xml'))
        brels = rels_of(z, 'xl/workbook.xml')
        sheets = {}
        parts = {}
        for s in kids(kid(book, 'sheets'), 'sheet'):
            rid = [v for k, v in s.attrib.items() if local(k) == 'id'][0]
            part = brels[rid][1]
            parts[s.get('name')] = part
            sheets[s.get('name')] = read_sheet(z, part, strings)
        tables = 0
        for sheet_name, part in parts.items():
            for kind, target in rels_of(z, part).values():
                if kind == 'pivotTable':
                    tables += 1
                    try:
                        check_table(z, name, target, sheets[sheet_name], sheets, strings, bad)
                    except Exception as error:
                        bad.append('%s %s: %s: %s' % (name, target, type(error).__name__, error))
    try:
        import openpyxl
    except ImportError:
        if REQUIRE_OPENPYXL:
            bad.append('%s: openpyxl is not installed' % name)
        return
    try:
        wb = openpyxl.load_workbook(path)
        found = sum(len(ws._pivots) for ws in wb.worksheets)
        if found != tables:
            bad.append('%s: openpyxl finds %d pivot tables, the file has %d' % (name, found, tables))
    except Exception as error:
        bad.append('%s: openpyxl cannot open it: %s: %s' % (name, type(error).__name__, error))


def main():
    folder = sys.argv[1]
    bad = []
    for entry in sorted(os.listdir(folder)):
        if entry.endswith('.xlsx'):
            check(os.path.join(folder, entry), bad)
    print(json.dumps(bad))


if __name__ == '__main__':
    main()
