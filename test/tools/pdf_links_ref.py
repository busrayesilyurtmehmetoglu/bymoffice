"""Reads the link annotations and the outline of a PDF file, for the tests.

Written from the PDF specification (ISO 32000-1, 7.3 objects, 7.5 file
structure, 12.3.3 outlines, 12.5.6.5 link annotations), sharing no code with
the engine: a check that the engine's files say what the engine thinks they
say. Only what the engine writes is read: a classic cross reference table,
objects that are not in object streams.

Prints JSON: {"pages": [{"annots": [{"rect": [...], "uri": "..."} |
{"rect": [...], "dest": [page index, left, top]} | {"rect": [...]}]}],
"outline": [{"title": "...", "dest": [...], "count": n, "kids": [...]}],
"outlineCount": n, "pageMode": "..."}

Usage: python3 -I pdf_links_ref.py file.pdf
"""

import json
import re
import sys

WHITE = b"\x00\t\n\x0c\r "
DELIMS = b"()<>[]{}/%"


class Ref:
    def __init__(self, num):
        self.num = num


class Name(str):
    pass


class Parser:
    def __init__(self, data, at):
        self.data = data
        self.at = at

    def skip(self):
        d = self.data
        while self.at < len(d):
            c = d[self.at]
            if c in WHITE:
                self.at += 1
            elif c == 0x25:  # % comment
                while self.at < len(d) and d[self.at] not in b"\r\n":
                    self.at += 1
            else:
                break

    def token(self):
        self.skip()
        d = self.data
        start = self.at
        while self.at < len(d) and d[self.at] not in WHITE and d[self.at] not in DELIMS:
            self.at += 1
        return d[start:self.at].decode("latin-1")

    def value(self):
        self.skip()
        d = self.data
        c = d[self.at]
        if d.startswith(b"<<", self.at):
            self.at += 2
            out = {}
            while True:
                self.skip()
                if d.startswith(b">>", self.at):
                    self.at += 2
                    return out
                key = self.value()
                out[str(key)] = self.value()
        if c == 0x5B:  # [
            self.at += 1
            out = []
            while True:
                self.skip()
                if d[self.at] == 0x5D:
                    self.at += 1
                    return out
                out.append(self.value())
        if c == 0x2F:  # /
            self.at += 1
            raw = self.token()
            return Name(re.sub(r"#([0-9A-Fa-f]{2})", lambda m: chr(int(m.group(1), 16)), raw))
        if c == 0x28:  # (
            return self.literal()
        if c == 0x3C:  # <hex>
            end = d.index(b">", self.at)
            hexa = re.sub(rb"\s", b"", d[self.at + 1:end])
            self.at = end + 1
            if len(hexa) % 2:
                hexa += b"0"
            return bytes.fromhex(hexa.decode("ascii"))
        word = self.token()
        if word in ("true", "false"):
            return word == "true"
        if word == "null":
            return None
        # A number, or the first of "n 0 R".
        save = self.at
        second = self.token()
        third = self.token()
        if re.fullmatch(r"\d+", word) and re.fullmatch(r"\d+", second) and third == "R":
            return Ref(int(word))
        self.at = save
        return float(word) if ("." in word) else int(word)

    def literal(self):
        d = self.data
        self.at += 1
        depth = 1
        out = bytearray()
        escapes = {ord("n"): 10, ord("r"): 13, ord("t"): 9, ord("b"): 8, ord("f"): 12}
        while True:
            c = d[self.at]
            self.at += 1
            if c == 0x5C:
                n = d[self.at]
                self.at += 1
                if n in escapes:
                    out.append(escapes[n])
                elif 0x30 <= n <= 0x37:
                    digits = chr(n)
                    while len(digits) < 3 and 0x30 <= d[self.at] <= 0x37:
                        digits += chr(d[self.at])
                        self.at += 1
                    out.append(int(digits, 8) & 255)
                elif n in b"\r\n":
                    if n == 0x0D and d[self.at] == 0x0A:
                        self.at += 1
                else:
                    out.append(n)
            elif c == 0x28:
                depth += 1
                out.append(c)
            elif c == 0x29:
                depth -= 1
                if depth == 0:
                    return bytes(out)
                out.append(c)
            else:
                out.append(c)


def text_string(raw):
    if raw[:2] == b"\xfe\xff":
        return raw[2:].decode("utf-16-be")
    return raw.decode("latin-1")


def main(path):
    data = open(path, "rb").read()
    start = int(re.search(rb"startxref\s+(\d+)\s+%%EOF\s*$", data).group(1))
    head = re.match(rb"xref\s+0 (\d+)\s+", data[start:])
    count = int(head.group(1))
    at = start + head.end()
    offsets = {}
    for i in range(count):
        entry = data[at + i * 20: at + i * 20 + 20]
        if entry[17:18] == b"n":
            offsets[i] = int(entry[:10])
    trailer = Parser(data, data.index(b"trailer", at) + 7).value()
    cache = {}

    def obj(ref):
        if not isinstance(ref, Ref):
            return ref
        if ref.num not in cache:
            p = Parser(data, offsets[ref.num])
            if p.token() != str(ref.num) or p.token() != "0" or p.token() != "obj":
                raise ValueError("object %d is not where the table says" % ref.num)
            cache[ref.num] = p.value()
        return cache[ref.num]

    root = obj(trailer["Root"])
    kids = []
    seen = set()
    stack = [root["Pages"]]
    while stack:
        node_ref = stack.pop()
        node = obj(node_ref)
        if node.get("Type") == "Pages":
            stack.extend(reversed(node["Kids"]))
        else:
            if node_ref.num in seen:
                raise ValueError("a page is in the tree twice")
            seen.add(node_ref.num)
            kids.append(node_ref.num)
    page_index = {num: i for i, num in enumerate(kids)}

    def dest(value):
        value = obj(value)
        if not isinstance(value, list) or not isinstance(value[0], Ref):
            raise ValueError("a destination is not [page /XYZ ...]")
        if value[1] != "XYZ":
            raise ValueError("a destination is not /XYZ")
        if value[0].num not in page_index:
            raise ValueError("a destination goes to an object that is no page")
        return [page_index[value[0].num], value[2], value[3]]

    pages = []
    for num in kids:
        page = obj(Ref(num))
        annots = []
        for ref in obj(page.get("Annots", [])):
            annot = obj(ref)
            if annot.get("Subtype") != "Link":
                raise ValueError("an annotation is not a link")
            entry = {"rect": annot["Rect"], "flags": annot.get("F"), "border": annot.get("Border")}
            if "A" in annot:
                action = obj(annot["A"])
                if action.get("S") != "URI":
                    raise ValueError("an action is not a URI")
                raw = action["URI"]
                if any(b > 126 or b < 33 for b in raw):
                    raise ValueError("a URI holds a byte that is not printable ASCII")
                entry["uri"] = raw.decode("ascii")
            if "Dest" in annot:
                entry["dest"] = dest(annot["Dest"])
            annots.append(entry)
        pages.append({"annots": annots})

    out = {"pages": pages, "outline": [], "outlineCount": None, "pageMode": root.get("PageMode")}
    if "Outlines" in root:
        top_ref = root["Outlines"]
        top = obj(top_ref)
        out["outlineCount"] = top.get("Count")
        visited = set()

        def items(parent_ref, parent):
            result = []
            ref = parent.get("First")
            previous = None
            while ref is not None:
                if ref.num in visited:
                    raise ValueError("an entry of the outline is reached twice")
                visited.add(ref.num)
                node = obj(ref)
                if node["Parent"].num != parent_ref.num:
                    raise ValueError("/Parent of an entry is not the entry it is under")
                prev = node.get("Prev")
                if (prev.num if prev else None) != (previous.num if previous else None):
                    raise ValueError("/Prev of an entry is not the entry before it")
                entry = {"title": text_string(node["Title"]), "dest": dest(node["Dest"]) if "Dest" in node else None, "count": node.get("Count"), "kids": items(ref, node)}
                if node.get("First") is not None:
                    last = node["Last"]
                    if last.num != max_last(node):
                        raise ValueError("/Last of an entry is not its last child")
                result.append(entry)
                previous = ref
                ref = node.get("Next")
            if result and parent.get("Last").num != previous.num:
                raise ValueError("/Last is not the last entry")
            return result

        def max_last(node):
            ref = node["First"]
            while obj(ref).get("Next") is not None:
                ref = obj(ref)["Next"]
            return ref.num

        out["outline"] = items(top_ref, top)
    print(json.dumps(out, ensure_ascii=False))


if __name__ == "__main__":
    main(sys.argv[1])
