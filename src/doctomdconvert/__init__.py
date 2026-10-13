"""Convert .docx / .pages to Markdown. Stdlib only."""
import re
import struct
import sys
import zipfile
from itertools import groupby
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def _on(el, tag):
    """True if run property <w:tag/> is present and not explicitly off."""
    p = el.find(f"{W}rPr/{W}{tag}") if el is not None else None
    return p is not None and p.get(f"{W}val", "true") not in ("0", "false", "none")


def _esc(s):
    s = re.sub(r"([\\`*_\[\]<>|])", r"\\\1", s)
    return re.sub(r"^(\s*)([#+-]|\d+[.)])(?=\s)", r"\1\\\2", s)  # would-be heading/list


def render(runs):
    """runs: (bold, italic, strike, url, code, text) -> one line of Markdown."""
    out = []
    for (b, i, s, url, code), grp in groupby(runs, key=lambda r: r[:5]):
        raw = "".join(r[5] for r in grp)
        text = raw.replace("\n", "  \n") if code else _esc(raw).replace("\n", "  \n")
        core = text.strip()
        if not core:
            out.append(text)
            continue
        lead, trail = text[: len(text) - len(text.lstrip())], text[len(text.rstrip()):]
        if code: core = f"`{core}`"
        if s: core = f"~~{core}~~"
        if i: core = f"*{core}*"
        if b: core = f"**{core}**"
        if url: core = f"[{core}]({url})"
        out.append(lead + core + trail)
    return "".join(out).strip()


class Docx:
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.root = ET.fromstring(z.read("word/document.xml"))
        self.styles = {}  # styleId -> lowercase name
        self.links = {}  # rId -> url
        self.bullets = {}  # (numId, ilvl) -> True if bullet
        names = z.namelist()
        if "word/styles.xml" in names:
            for s in ET.fromstring(z.read("word/styles.xml")).iter(f"{W}style"):
                n = s.find(f"{W}name")
                if n is not None:
                    self.styles[s.get(f"{W}styleId")] = n.get(f"{W}val").lower()
        if "word/_rels/document.xml.rels" in names:
            for r in ET.fromstring(z.read("word/_rels/document.xml.rels")).iter(f"{REL}Relationship"):
                if r.get("TargetMode") == "External":
                    self.links[r.get("Id")] = r.get("Target")
        if "word/numbering.xml" in names:
            num = ET.fromstring(z.read("word/numbering.xml"))
            absfmt = {}
            for a in num.iter(f"{W}abstractNum"):
                for l in a.iter(f"{W}lvl"):
                    f = l.find(f"{W}numFmt")
                    absfmt[(a.get(f"{W}abstractNumId"), l.get(f"{W}ilvl"))] = (
                        f is not None and f.get(f"{W}val") == "bullet")
            for n in num.iter(f"{W}num"):
                aid = n.find(f"{W}abstractNumId").get(f"{W}val")
                for (a, lvl), b in absfmt.items():
                    if a == aid:
                        self.bullets[(n.get(f"{W}numId"), lvl)] = b

    def runs(self, el, url=None):
        """Yield (bold, italic, strike, url, text) for every run under el."""
        for c in el:
            if c.tag == f"{W}hyperlink":
                yield from self.runs(c, self.links.get(c.get(f"{R}id")))
            elif c.tag == f"{W}r":
                t = "".join(
                    x.text or "" if x.tag == f"{W}t" else "\n" if x.tag in (f"{W}br", f"{W}cr") else " "
                    for x in c if x.tag in (f"{W}t", f"{W}br", f"{W}cr", f"{W}tab"))
                if t:
                    yield _on(c, "b"), _on(c, "i"), _on(c, "strike"), url, t
            elif c.tag in (f"{W}ins", f"{W}smartTag", f"{W}sdt", f"{W}sdtContent"):
                yield from self.runs(c, url)

    def inline(self, p):
        return render((b, i, st, url, False, t) for b, i, st, url, t in self.runs(p))

    def paragraph(self, p):
        text = self.inline(p)
        ppr = p.find(f"{W}pPr")
        sid = ppr.find(f"{W}pStyle") if ppr is not None else None
        style = self.styles.get(sid.get(f"{W}val"), "") if sid is not None else ""
        if not text:
            return None
        if style == "title":
            return f"# {text}"
        if m := re.match(r"heading (\d)", style):
            return f"{'#' * min(int(m[1]), 6)} {text}"
        num = ppr.find(f"{W}numPr") if ppr is not None else None
        if num is not None or style.startswith("list"):
            lvl = num.find(f"{W}ilvl").get(f"{W}val") if num is not None and num.find(f"{W}ilvl") is not None else "0"
            nid = num.find(f"{W}numId").get(f"{W}val") if num is not None and num.find(f"{W}numId") is not None else None
            bullet = self.bullets.get((nid, lvl), "number" not in style)
            return f"{'  ' * int(lvl)}{'-' if bullet else '1.'} {text}"
        return text

    def table(self, tbl):
        rows = []
        for tr in tbl.findall(f"{W}tr"):
            rows.append([" <br> ".join(filter(None, (self.inline(p) for p in tc.iter(f"{W}p")))).replace("\n", " ")
                         for tc in tr.findall(f"{W}tc")])
        if not rows:
            return None
        n = max(map(len, rows))
        rows = [r + [""] * (n - len(r)) for r in rows]
        line = lambda r: "| " + " | ".join(r) + " |"
        return "\n".join([line(rows[0]), line(["---"] * n), *map(line, rows[1:])])

    def markdown(self):
        blocks = []  # consecutive list items are joined by \n, everything else by blank line
        prev_list = False
        for el in self.root.find(f"{W}body"):
            md = self.paragraph(el) if el.tag == f"{W}p" else self.table(el) if el.tag == f"{W}tbl" else None
            if md is None:
                continue
            is_list = re.match(r"\s*(-|1\.) ", md) is not None and el.tag == f"{W}p"
            if is_list and prev_list:
                blocks[-1] += "\n" + md
            else:
                blocks.append(md)
            prev_list = is_list
        return "\n\n".join(blocks) + "\n"


def _snappy(b):
    """Raw snappy block decompression (the IWA chunk format)."""
    i = 0
    n = s = 0
    while True:
        c = b[i]; i += 1; n |= (c & 127) << s; s += 7
        if c < 128: break
    out = bytearray()
    while i < len(b):
        t = b[i]; i += 1; k = t & 3
        if k == 0:
            l = t >> 2
            if l >= 60:
                l = int.from_bytes(b[i:i + l - 59], "little"); i += (t >> 2) - 59
            l += 1; out += b[i:i + l]; i += l
        else:
            if k == 1: l = ((t >> 2) & 7) + 4; o = ((t >> 5) << 8) | b[i]; i += 1
            elif k == 2: l = (t >> 2) + 1; o = int.from_bytes(b[i:i + 2], "little"); i += 2
            else: l = (t >> 2) + 1; o = int.from_bytes(b[i:i + 4], "little"); i += 4
            for _ in range(l): out.append(out[-o])
    return bytes(out)


def _varint(b, i):
    n = s = 0
    while True:
        c = b[i]; i += 1; n |= (c & 127) << s; s += 7
        if c < 128: return n, i


def _pb(b):
    """Schema-less protobuf decode -> [(field, value)]; wire type 2 stays bytes, 5 -> float."""
    i, r = 0, []
    while i < len(b):
        t, i = _varint(b, i); f, w = t >> 3, t & 7
        if w == 0: v, i = _varint(b, i)
        elif w == 2: l, i = _varint(b, i); v = b[i:i + l]; i += l
        elif w == 1: v = b[i:i + 8]; i += 8
        elif w == 5: v = struct.unpack("<f", b[i:i + 4])[0]; i += 4
        else: raise ValueError("bad protobuf")
        r.append((f, v))
    return r


def _ref(b):
    """Object reference message {1: id} -> id."""
    return _pb(b)[0][1]


class Pages:
    # Apple's .pages is zip of .iwa files = snappy-chunked protobuf. No schemas are bundled:
    # field numbers below were reverse-engineered (TSWP.StorageArchive etc.).
    # ponytail: no images/tables/hyperlinks/nested lists; headings inferred from bold+size.
    def __init__(self, path):
        z = zipfile.ZipFile(path)
        self.obj = {}  # id -> (message type, payload)
        for n in z.namelist():
            if n.startswith("Index/") and n.endswith(".iwa"):
                self._load(z.read(n))
        self.storage = None
        for t, p in self.obj.values():
            if t == 2001:
                d = _pb(p)
                if dict(d).get(1) == 0 and (self.storage is None or len(dict(d).get(3, b"")) > len(self.storage[3])):
                    self.storage = {}
                    for f, v in d:
                        self.storage.setdefault(f, v) if f in (1, 3) else self.storage.setdefault(f, []).append(v)
        if self.storage is None or not isinstance(self.storage.get(3), bytes):
            sys.exit("No body text found in .pages file")

    def _load(self, data):
        raw, i = b"", 0
        while i < len(data):
            l = int.from_bytes(data[i + 1:i + 4], "little"); i += 4
            raw += _snappy(data[i:i + l]); i += l
        i = 0
        while i < len(raw):
            l, i = _varint(raw, i)
            info = _pb(raw[i:i + l]); i += l
            msgs = [dict(_pb(v)) for f, v in info if f == 2]
            oid = dict(info)[1]
            for m in msgs:
                self.obj.setdefault(oid, (m[1], raw[i:i + m[3]])); i += m[3]  # first message wins

    def style(self, oid):
        """-> (nearest name, merged char props {bold, italic, size, font}) walking the parent chain."""
        chain = []
        while oid in self.obj and len(chain) < 20:
            sup = dict(_pb(dict(_pb(self.obj[oid][1]))[1]))
            chain.append((self.obj[oid][1], sup))
            oid = _ref(sup[3]) if 3 in sup else None
        name, props = "", {}
        for payload, sup in reversed(chain):
            name = sup[1].decode() if 1 in sup else name
            cp = dict(_pb(payload)).get(11)
            for f, v in _pb(cp) if isinstance(cp, bytes) else ():  # list styles keep other data here
                props[f] = v
        return name, props

    def table(self, key):
        """Per-offset tables (field 5 para style, 7 list style, 8 char style) -> [(offset, style id|None)]."""
        out = []
        for e in self.storage.get(key, [b""]):
            for f, v in _pb(e):
                d = dict(_pb(v))
                out.append((d[1], _ref(d[2]) if 2 in d else None))
        return out

    def markdown(self):
        text = self.storage[3].decode("utf-8")
        u = text.encode("utf-16-le")  # Pages offsets count UTF-16 units
        cut = lambda a, b: u[2 * a:2 * b].decode("utf-16-le")
        paras = self.table(5)
        lists, chars = self.table(7), self.table(8)
        at = lambda tbl, pos: next((s for o, s in reversed(tbl) if o <= pos), None)
        starts, pos = [], 0
        for line in text.split("\n"):
            starts.append((pos, pos + len(line.encode("utf-16-le")) // 2)); pos = starts[-1][1] + 1
        blocks, prev_list = [], False
        for a, b in starts:
            if a == b:
                continue
            name, pr = self.style(at(paras, a)) if at(paras, a) else ("", {})
            heading = 0
            if name == "Title": heading = 1
            elif m := re.fullmatch(r"Heading ?(\d?)", name): heading = int(m[1] or 1) + 1
            elif pr.get(1) == 1 and pr.get(3, 12) >= 14: heading = 1 if pr[3] >= 24 else 2 if pr[3] >= 18 else 3
            cuts = sorted({a, b, *(o for o, _ in chars if a < o < b)})
            runs = []
            for x, y in zip(cuts, cuts[1:]):
                cs = at(chars, x)
                _, cp = self.style(cs) if cs else ("", {})
                runs.append((bool(cp.get(1, pr.get(1) if not heading else 0)), bool(cp.get(2, pr.get(2, 0))), False, None,
                             b"Courier" in cp.get(5, b"") or b"Mono" in cp.get(5, b""),
                             cut(x, y).replace("\u2028", "\n").replace("\ufffc", "")))
            md = render(runs)
            if not md:
                continue
            ls = at(lists, a)
            lname = self.style(ls)[0] if ls else "None"
            if heading:
                md, is_list = f"{'#' * heading} {md}", False
            elif lname not in ("None", ""):
                bullet = lname in ("Bullet", "Bullet Big", "Dash", "Image")
                md, is_list = f"{'-' if bullet else '1.'} {md}", True
            else:
                is_list = False
            if is_list and prev_list: blocks[-1] += "\n" + md
            else: blocks.append(md)
            prev_list = is_list
        return "\n\n".join(blocks) + "\n"


def convert(src, dst=None):
    src = Path(src).resolve()
    dst = Path(dst) if dst else src.with_suffix(".md")
    ext = src.suffix.lower()
    if ext == ".docx":
        dst.write_text(Docx(src).markdown(), encoding="utf-8")
    elif ext == ".pages":
        dst.write_text(Pages(src).markdown(), encoding="utf-8")
    else:
        sys.exit(f"Unsupported file type: {ext} (want .docx or .pages)")
    return dst


def main() -> None:
    if not 2 <= len(sys.argv) <= 3:
        sys.exit("usage: doctomdconvert INPUT.{docx,pages} [OUTPUT.md]")
    print(convert(*sys.argv[1:]))
