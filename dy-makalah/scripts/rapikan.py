#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YZMakalah V2 - perapi makalah .docx yang sudah jadi (YZ PROJECT - @yzproject24)
Pemakaian : python3 yzmakalah2.py periksa [MAKALAH.docx] [--penuh]
            python3 yzmakalah2.py konteks KATA [MAKALAH.docx]
            python3 yzmakalah2.py rapikan config.json [MAKALAH.docx] [--out FOLDER]
Kebutuhan : Python 3.8+ (hanya pustaka standar). LibreOffice + pdftotext bersifat
            opsional; jika ada, nomor halaman Daftar Isi diisi otomatis & diverifikasi.
Prinsip   : tata letak dibangun ulang, kata-kata naskah tidak disentuh. Satu-satunya
            perubahan teks adalah daftar "koreksi" di config, dan hasilnya diaudit
            kata per kata terhadap file asli.
JANGAN MENGUBAH SKRIP INI. Semua pengaturan dilakukan lewat config.json.
"""
import sys, os, re, json, zipfile, shutil, subprocess, tempfile, unicodedata, copy, difflib
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape as _xesc

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
for _p, _u in (("w", W), ("r", R), ("wp", WP), ("a", A), ("pic", PIC), ("mc", MC), ("m", M),
               ("v", "urn:schemas-microsoft-com:vml"), ("o", "urn:schemas-microsoft-com:office:office"),
               ("w10", "urn:schemas-microsoft-com:office:word"),
               ("w14", "http://schemas.microsoft.com/office/word/2010/wordml"),
               ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
               ("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
               ("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"),
               ("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"),
               ("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"),
               ("a14", "http://schemas.microsoft.com/office/drawing/2010/main"),
               ("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")):
    ET.register_namespace(_p, _u)

def q(t): return "{%s}%s" % (W, t)
def ln(e): return e.tag.split("}")[-1] if isinstance(e.tag, str) else ""
def esc(s): return _xesc(str(s), {'"': "&quot;"})

CM = 567
PAGE_W, PAGE_H = 11906, 16838
WARN = []
def warn(m):
    if m not in WARN: WARN.append(m)

def roman(n, lower=False):
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"),
            (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    s = ""
    for v, r in vals:
        while n >= v:
            s += r; n -= v
    return s.lower() if lower else s

def letter(n, upper=True):
    s = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        s = chr(65 + m) + s
    return s if upper else s.lower()

def norm(s):
    return " ".join(unicodedata.normalize("NFKC", s).split())

# =====================================================================
# 1. MODEL TEKS: Seg (potongan teks / objek yang tidak boleh disentuh)
# =====================================================================
class Seg:
    __slots__ = ("t", "b", "i", "u", "sup", "sub", "strike", "link", "raw", "note")
    def __init__(self, t="", b=False, i=False, u=False, sup=False, sub=False, strike=False, link=None, raw=None, note=False):
        self.t, self.b, self.i, self.u, self.sup, self.sub, self.strike = t, b, i, u, sup, sub, strike
        self.link, self.raw, self.note = link, raw, note
    def clone(self, t=None):
        c = Seg(self.t if t is None else t, self.b, self.i, self.u, self.sup, self.sub, self.strike, self.link, self.raw, self.note)
        return c

OBJ = "\uFFFC"
def plain(segs):
    return "".join(OBJ if s.raw is not None else s.t for s in segs)
def text_only(segs):
    return "".join("" if s.raw is not None else s.t for s in segs)

def split_seg_at(segs, pos):
    """Pastikan ada batas Seg tepat di posisi pos (indeks pada plain()). Mengembalikan indeks Seg."""
    n = 0
    for k, s in enumerate(segs):
        L = 1 if s.raw is not None else len(s.t)
        if pos == n: return k
        if n < pos < n + L:
            a, b = s.clone(s.t[:pos - n]), s.clone(s.t[pos - n:])
            segs[k:k + 1] = [a, b]
            return k + 1
        n += L
    return len(segs)

def replace_range(segs, start, end, new):
    a = split_seg_at(segs, start)
    b = split_seg_at(segs, end)
    a = split_seg_at(segs, start)
    fmt = segs[a].clone(new) if a < len(segs) and segs[a].raw is None else Seg(new)
    segs[a:b] = [fmt] if new else []

def apply_replace(segs, cari, ganti):
    """Ganti semua kemunculan cari -> ganti. Mengembalikan jumlah penggantian."""
    n, pos = 0, 0
    while True:
        p = plain(segs)
        i = p.find(cari, pos)
        if i < 0: return n
        replace_range(segs, i, i + len(cari), ganti)
        pos = i + len(ganti); n += 1

def mark_italic(segs, term):
    rx = re.compile(r"(?<![\w\-])" + re.escape(term) + r"(?![\w\-])")
    p = plain(segs)
    for m in reversed(list(rx.finditer(p))):
        a = split_seg_at(segs, m.start())
        b = split_seg_at(segs, m.end())
        a = split_seg_at(segs, m.start())
        for s in segs[a:b]:
            if s.raw is None: s.i = True

def tidy_ws(segs):
    """Rapikan spasi ganda & spasi di tepi paragraf. Kata tidak berubah."""
    out = [s for s in segs if s.raw is not None or s.t != ""]
    prev_space = True
    for s in out:
        if s.raw is not None:
            prev_space = False; continue
        t = re.sub(r"[ \u00a0]{2,}", " ", s.t)
        if prev_space: t = t.lstrip(" \u00a0\t")
        s.t = t
        if t: prev_space = t[-1] in " \u00a0\t\n"
    out = [s for s in out if s.raw is not None or s.t != ""]
    while out and out[-1].raw is None:
        out[-1].t = out[-1].t.rstrip(" \u00a0\t\n")
        if out[-1].t: break
        out.pop()
    while out and out[0].raw is None:
        out[0].t = out[0].t.lstrip(" \u00a0\t\n")
        if out[0].t: break
        out.pop(0)
    return out

def split_lines(segs):
    """Pecah daftar Seg pada karakter baris baru (w:br)."""
    res, cur = [], []
    for s in segs:
        if s.raw is not None or "\n" not in s.t:
            cur.append(s); continue
        parts = s.t.split("\n")
        for k, part in enumerate(parts):
            if part: cur.append(s.clone(part))
            if k < len(parts) - 1:
                res.append(cur); cur = []
    res.append(cur)
    return [tidy_ws(x) for x in res]

# =====================================================================
# 2. MEMBACA DOCX SUMBER
# =====================================================================
class Num:
    """Menghitung label penomoran otomatis Word (untuk judul bernomor otomatis)."""
    def __init__(self, xml):
        self.abs, self.nums, self.cnt, self.seen = {}, {}, {}, set()
        if not xml: return
        root = ET.fromstring(xml)
        for a in root.findall(q("abstractNum")):
            lv = {}
            for l in a.findall(q("lvl")):
                g = lambda t, d=None: (l.find(q(t)).get(q("val")) if l.find(q(t)) is not None else d)
                lv[int(l.get(q("ilvl")))] = {"fmt": g("numFmt", "decimal"), "text": g("lvlText", ""), "start": int(g("start", "1") or 1)}
            self.abs[a.get(q("abstractNumId"))] = lv
        for n in root.findall(q("num")):
            ov = {}
            for o in n.findall(q("lvlOverride")):
                so = o.find(q("startOverride"))
                if so is not None: ov[int(o.get(q("ilvl")))] = int(so.get(q("val")))
            self.nums[n.get(q("numId"))] = (n.find(q("abstractNumId")).get(q("val")), ov)

    def fmt(self, v, f):
        return {"decimal": str(v), "upperLetter": letter(v), "lowerLetter": letter(v, False), "upperRoman": roman(v),
                "lowerRoman": roman(v, True), "decimalZero": "%02d" % v}.get(f, str(v))

    def next(self, numId, ilvl):
        numId = str(numId)
        if numId not in self.nums: return "", "bullet"
        aid, ov = self.nums[numId]
        lv = self.abs.get(aid, {})
        if ilvl not in lv: return "", "bullet"
        c = self.cnt.setdefault(aid, {})
        if (numId, ilvl) not in self.seen and ilvl in ov:
            c[ilvl] = ov[ilvl] - 1
        self.seen.add((numId, ilvl))
        c[ilvl] = c.get(ilvl, lv[ilvl]["start"] - 1) + 1
        for k in list(c):
            if k > ilvl: del c[k]
        info = lv[ilvl]
        if info["fmt"] in ("bullet", "none"): return "", info["fmt"]
        def rep(m):
            k = int(m.group(1)) - 1
            if k not in lv: return ""
            return self.fmt(c.get(k, lv[k]["start"]), lv[k]["fmt"])
        return re.sub(r"%(\d)", rep, info["text"]), info["fmt"]

class Styles:
    def __init__(self, xml):
        self.p, self.c = {}, {}
        self.default = None
        if not xml: return
        root = ET.fromstring(xml)
        for s in root.findall(q("style")):
            sid = s.get(q("styleId")); typ = s.get(q("type"))
            nm = s.find(q("name")); base = s.find(q("basedOn"))
            d = {"name": (nm.get(q("val")) if nm is not None else sid).lower(), "base": base.get(q("val")) if base is not None else None,
                 "outline": None, "num": None, "jc": None, "b": None, "i": None}
            pp = s.find(q("pPr"))
            if pp is not None:
                o = pp.find(q("outlineLvl"))
                if o is not None: d["outline"] = int(o.get(q("val")))
                n = pp.find(q("numPr"))
                if n is not None:
                    ni, il = n.find(q("numId")), n.find(q("ilvl"))
                    d["num"] = (ni.get(q("val")) if ni is not None else None, int(il.get(q("val"))) if il is not None else 0)
                j = pp.find(q("jc"))
                if j is not None: d["jc"] = j.get(q("val"))
            rp = s.find(q("rPr"))
            if rp is not None:
                for t in ("b", "i"):
                    e = rp.find(q(t))
                    if e is not None: d[t] = e.get(q("val"), "1") not in ("0", "false")
            if typ == "paragraph":
                self.p[sid] = d
                if s.get(q("default")) == "1": self.default = sid
            elif typ == "character":
                self.c[sid] = d

    def chain(self, sid, key, table=None):
        table = self.p if table is None else table
        seen = set()
        while sid and sid in table and sid not in seen:
            seen.add(sid)
            if table[sid].get(key) is not None: return table[sid][key]
            sid = table[sid]["base"]
        return None

    def name(self, sid):
        return self.p[sid]["name"] if sid in self.p else (sid or "").lower()

class Block:
    def __init__(self, idx):
        self.id = str(idx); self.kind = "p"; self.segs = []; self.el = None
        self.style = ""; self.outline = None; self.num = None; self.numlabel = ""; self.numfmt = ""
        self.jc = None; self.left = 0; self.first = 0; self.hanging = 0
        self.in_toc = False; self.role = None; self.forced = False; self.note = ""; self.asli = ""
        self.has_img = False; self.prefix = ""; self.section = None; self.keep_next = False; self.src = idx
    @property
    def text(self): return norm(text_only(self.segs))

class Reader:
    def __init__(self, parts, scale_tw=None):
        self.styles = Styles(parts.get("word/styles.xml"))
        self.num = Num(parts.get("word/numbering.xml"))
        self.tw = scale_tw
        self.depth = 0; self.toc_depth = None
        self.pending_hidden = 0

    def fmt_of(self, rp):
        d = dict(b=False, i=False, u=False, sup=False, sub=False, strike=False)
        hidden = False
        if rp is not None:
            on = lambda e: e is not None and e.get(q("val"), "1") not in ("0", "false", "none")
            rs = rp.find(q("rStyle"))
            if rs is not None:
                sid = rs.get(q("val"))
                if self.styles.chain(sid, "b", self.styles.c): d["b"] = True
                if self.styles.chain(sid, "i", self.styles.c): d["i"] = True
            for t in ("b", "i"):
                e = rp.find(q(t))
                if e is not None: d[t] = on(e)
            d["u"] = on(rp.find(q("u")))
            d["strike"] = on(rp.find(q("strike")))
            va = rp.find(q("vertAlign"))
            if va is not None:
                d["sup"] = va.get(q("val")) == "superscript"; d["sub"] = va.get(q("val")) == "subscript"
            hidden = on(rp.find(q("vanish")))
        return d, hidden

    def scale_drawing(self, el):
        if not self.tw: return
        maxcx = int(self.tw * 635)
        for ext in el.iter("{%s}extent" % WP):
            cx, cy = int(ext.get("cx", 0)), int(ext.get("cy", 0))
            if cx > maxcx and cx > 0:
                k = maxcx / cx
                ext.set("cx", str(maxcx)); ext.set("cy", str(int(cy * k)))
                for e in el.iter("{%s}ext" % A):
                    if e.get("cx") and e.get("cy"):
                        e.set("cx", str(int(int(e.get("cx")) * k))); e.set("cy", str(int(int(e.get("cy")) * k)))
        if any(True for _ in el.iter("{%s}anchor" % WP)):
            warn("Ada gambar melayang (bukan sebaris teks). Posisinya tidak diubah, mohon dicek di Word.")

    def run(self, r, out, link=None, blk=None):
        d, hidden = self.fmt_of(r.find(q("rPr")))
        if hidden:
            if any(ln(c) == "t" and (c.text or "").strip() for c in r):
                warn("Ada teks tersembunyi (hidden text) di naskah. Teks itu tidak ditampilkan pada hasil.")
            return
        for c in r:
            n = ln(c)
            if n == "rPr" or n == "lastRenderedPageBreak": continue
            if n == "t":
                if c.text: out.append(Seg(c.text, link=link, **d))
            elif n == "tab": out.append(Seg("\t", link=link, **d))
            elif n in ("br", "cr"):
                if c.get(q("type")) in ("page", "column"): continue
                out.append(Seg("\n", link=link, **d))
            elif n == "noBreakHyphen": out.append(Seg("\u2011", link=link, **d))
            elif n == "softHyphen": continue
            else:
                if n == "fldChar":
                    ty = c.get(q("fldCharType"))
                    if ty == "begin": self.depth += 1
                    elif ty == "end":
                        self.depth -= 1
                        if self.toc_depth is not None and self.depth < self.toc_depth: self.toc_depth = None
                elif n == "instrText":
                    if re.match(r"\s*TOC\b", c.text or "") and self.toc_depth is None:
                        self.toc_depth = self.depth
                        if blk is not None: blk.in_toc = True
                if n in ("drawing", "pict", "object", "AlternateContent"):
                    self.scale_drawing(c)
                    if blk is not None: blk.has_img = True
                    if n == "AlternateContent" or any(ln(x) == "txbx" or ln(x) == "txbxContent" for x in c.iter()):
                        warn("Ada kotak teks/objek di naskah. Objek itu dipertahankan apa adanya, mohon dicek di Word.")
                note = n in ("footnoteReference", "endnoteReference")
                out.append(Seg(raw=ET.tostring(c, encoding="unicode"), link=link, note=note, sup=d["sup"]))

    def inline(self, node, out, link=None, blk=None):
        for c in node:
            n = ln(c)
            if n == "r": self.run(c, out, link, blk)
            elif n == "hyperlink":
                rid = c.get("{%s}id" % R)
                self.inline(c, out, rid or link, blk)
            elif n in ("ins", "smartTag", "customXml", "fldSimple", "moveTo", "dir", "bdo"):
                self.inline(c, out, link, blk)
            elif n == "sdt":
                sc = c.find(q("sdtContent"))
                if sc is not None: self.inline(sc, out, link, blk)
            elif n in ("oMath", "oMathPara"):
                out.append(Seg(raw=ET.tostring(c, encoding="unicode")))

    def para(self, p, idx):
        b = Block(idx); b.kind = "p"
        if self.toc_depth is not None: b.in_toc = True
        pp = p.find(q("pPr")); sid = None; direct_num = None
        if pp is not None:
            ps = pp.find(q("pStyle"))
            if ps is not None: sid = ps.get(q("val"))
            np = pp.find(q("numPr"))
            if np is not None:
                ni, il = np.find(q("numId")), np.find(q("ilvl"))
                if ni is not None: direct_num = (ni.get(q("val")), int(il.get(q("val"))) if il is not None else 0)
            j = pp.find(q("jc"))
            if j is not None: b.jc = j.get(q("val"))
            o = pp.find(q("outlineLvl"))
            if o is not None: b.outline = int(o.get(q("val")))
            ind = pp.find(q("ind"))
            if ind is not None:
                g = lambda *k: next((int(ind.get(q(x))) for x in k if ind.get(q(x)) not in (None, "")), 0)
                b.left, b.first, b.hanging = g("left", "start"), g("firstLine"), g("hanging")
            if pp.find(q("sectPr")) is not None:
                ss = pp.find(q("sectPr")).find(q("pgSz"))
                if ss is not None and ss.get(q("orient")) == "landscape":
                    warn("Ada halaman mendatar (landscape) di naskah asli; hasilnya dibuat tegak semua.")
        sid = sid or self.styles.default
        b.style = self.styles.name(sid)
        if b.outline is None: b.outline = self.styles.chain(sid, "outline")
        if b.jc is None: b.jc = self.styles.chain(sid, "jc")
        num = direct_num or self.styles.chain(sid, "num")
        if num and num[0] not in (None, "0"):
            b.num = num
            b.numlabel, b.numfmt = self.num.next(num[0], num[1])
        self.inline(p, b.segs, None, b)
        if self.toc_depth is not None: b.in_toc = True
        b.asli = text_only(b.segs)
        return b

    def table(self, t, idx):
        b = Block(idx); b.kind = "tbl"; b.el = t
        tx = []
        for tp in t.iter(q("p")):
            tmp = []
            self.inline(tp, tmp)
            tx.append(text_only(tmp))
        b.asli = " ".join(tx)
        return b

    def body(self, root):
        body = root.find(q("body"))
        out = []
        def walk(node):
            for c in node:
                n = ln(c)
                if n == "p": out.append(self.para(c, len(out)))
                elif n == "tbl": out.append(self.table(c, len(out)))
                elif n == "sdt":
                    sc = c.find(q("sdtContent"))
                    if sc is not None: walk(sc)
                elif n in ("customXml", "ins", "moveTo"): walk(c)
        walk(body)
        return out

def read_parts(path):
    parts = {}
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith("/"): continue
            parts[n] = z.read(n)
    return parts

# =====================================================================
# 3. MENGENALI PERAN TIAP PARAGRAF
# =====================================================================
FRONT_RX = re.compile(r"^(KATA PENGANTAR|PRAKATA|ABSTRAK|ABSTRACT|DAFTAR ISI|DAFTAR TABEL|DAFTAR GAMBAR|DAFTAR LAMPIRAN|DAFTAR SINGKATAN|DAFTAR LAMBANG)\s*:?$", re.I)
LISTS = {"DAFTAR ISI": "MAIN", "DAFTAR TABEL": "Tabel", "DAFTAR GAMBAR": "Gambar", "DAFTAR LAMPIRAN": "Lampiran"}
BACK_RX = re.compile(r"^(DAFTAR PUSTAKA|DAFTAR REFERENSI|REFERENSI|BIBLIOGRAFI|LAMPIRAN(\s+[\dIVX]+)?(-LAMPIRAN)?)\s*:?$", re.I)
BODYH_RX = re.compile(r"^(PENDAHULUAN|PEMBAHASAN|PENUTUP|KESIMPULAN)$")
BAB_RX = re.compile(r"^BAB\s+([IVXLCDM]+|\d+)\b")
CAP_RX = re.compile(r"^(Tabel|Gambar|Grafik|Bagan|Diagram|TABEL|GAMBAR)\s+(\d+(?:[.\-]\d+)*)\s*[.:\-\u2013]?\s*")
LAMP_RX = re.compile(r"^Lampiran\s+\d+", re.I)
SUMBER_RX = re.compile(r"^(sumber|source)\s*:", re.I)
TOCLINE_RX = re.compile(r"(\t\s*[\divxlcIVXLC]+\s*$)|([.\u2026]{2,}\s*[\divxlcIVXLC]+\s*$)|(\s\d{1,3}\s*$)")

def is_toc_line(b):
    if b.kind != "p": return False
    if b.in_toc or b.style.startswith("toc") or b.style == "table of figures": return True
    t = text_only(b.segs)
    return bool(TOCLINE_RX.search(t)) and len(t) < 200

def h1_kind(t, b):
    if len(t) > 140: return None
    if FRONT_RX.match(t): return "front"
    if BACK_RX.match(t): return "back"
    if BAB_RX.match(t): return "bab"
    if BODYH_RX.match(t) and (b.outline == 0 or b.style.startswith("heading")): return "bab"
    return None

def headish(t, b, nxt):
    if len(t) > 150 or re.search(r"[.;,?!:]$", t): return False
    txt = [s for s in b.segs if s.raw is None and s.t.strip()]
    allbold = bool(txt) and all(s.b for s in txt)
    if b.outline is not None or b.style.startswith("heading") or allbold: return True
    return nxt is not None and len(nxt.text) > 150

def classify(blocks):
    meta = {"toc_lama": [], "ada": set()}
    n = len(blocks)
    # gabungkan judul bab yang terpisah dua paragraf (BAB I / PENDAHULUAN)
    i = 0
    while i < n - 1:
        b, c = blocks[i], blocks[i + 1]
        if b.kind == "p" and c.kind == "p" and not b.in_toc and re.match(r"^BAB\s+([IVXLCDM]+|\d+)$", b.text) \
                and c.text and len(c.text) < 100 and c.text.upper() == c.text and not c.forced:
            b.segs += [Seg(" ")] + c.segs
            b.asli += " " + c.asli
            del blocks[i + 1]; n -= 1
        i += 1
    # bab pertama / halaman awal -> batas sampul
    first = None
    for k, b in enumerate(blocks):
        if b.kind == "p" and not is_toc_line(b) and (h1_kind(b.text, b) or b.role in ("h1",)):
            first = k; break
    if first is None:
        tocs = [k for k, b in enumerate(blocks) if is_toc_line(b)]
        first = tocs[0] if tocs else 0
    for k in range(first):
        b = blocks[k]
        if not b.forced: b.role = "sampul"
        elif b.role is None: b.role = "sampul"
    has_letter = any(b.kind == "p" and re.match(r"^[A-Z]\.\s", b.text) and len(b.text) < 150 for b in blocks[first:])
    mode, section = None, "front"
    if not any(b.kind == "p" and not is_toc_line(b) and (h1_kind(b.text, b) == "front" or b.role == "h1") for b in blocks[first:first + 1]):
        section = "isi"
    k = first
    while k < n:
        b = blocks[k]
        t = b.text
        nxt = blocks[k + 1] if k + 1 < n else None
        if b.kind == "tbl":
            b.section = section; k += 1; continue
        hk = h1_kind(t, b) if not is_toc_line(b) else None
        if mode == "list" and (is_toc_line(b) or not t) and not hk and not b.forced:
            b.role = "buang"; meta["toc_lama"].append(text_only(b.segs)); k += 1; continue
        mode = None
        if not b.forced and is_toc_line(b):
            b.role = "buang"; meta["toc_lama"].append(text_only(b.segs)); b.section = section; k += 1; continue
        if not t and not b.has_img and not any(s.raw is not None for s in b.segs):
            b.role = "buang"; k += 1; continue
        if hk or (b.forced and b.role == "h1"):
            up = t.upper().rstrip(":").strip()
            if hk == "front":
                meta["ada"].add(up)
                if up in LISTS:
                    b.role = "judul_daftar"; mode = "list"; b.section = "front"; k += 1; continue
                section = "front"
            else:
                section = "isi"
            if b.num and b.numlabel and not b.forced: b.prefix = b.numlabel
            if not b.forced: b.role = "h1"
            b.h1kind = hk
            b.section = section
            b.pustaka = bool(re.match(r"^(DAFTAR PUSTAKA|DAFTAR REFERENSI|REFERENSI|BIBLIOGRAFI)", up))
            mode_p = b.pustaka
            k += 1
            meta["_pustaka"] = mode_p
            continue
        b.section = section
        if b.forced:
            if b.role == "pustaka": pass
            k += 1; continue
        full = t
        if b.numlabel and b.outline is not None: full = b.numlabel + " " + t
        if meta.get("_pustaka"):
            b.role = "gambar" if b.has_img and not t else "pustaka"; k += 1; continue
        if b.has_img and not t:
            b.role = "gambar"; k += 1; continue
        if SUMBER_RX.match(t) and len(t) < 250:
            b.role = "sumber"; k += 1; continue
        if LAMP_RX.match(t) and len(t) < 200:
            b.role = "cap_lampiran"; k += 1; continue
        m = CAP_RX.match(t)
        if m and not re.match(r"^\S+\s+\S+\s*[.:\-\u2013]", t) and t[m.end():m.end() + 1].islower():
            m = None  # kalimat biasa ("Tabel 1 menunjukkan ..."), bukan judul tabel
        if m and len(t) < 300:
            prev = blocks[k - 1] if k else None
            near = (nxt is not None and (nxt.kind == "tbl" or (nxt.has_img and not nxt.text))) or \
                   (prev is not None and (prev.kind == "tbl" or (prev.has_img and not prev.text)))
            if near or b.style == "caption":
                b.role = "cap_tabel" if m.group(1).lower() == "tabel" else "cap_gambar"
                k += 1; continue
        if b.num and b.outline is None:
            b.role = "daftar"; k += 1; continue
        role = None
        if re.match(r"^\d+\.\d+\.\d+\.?\s", full) and headish(full, b, nxt): role = "h3"
        elif re.match(r"^\d+\.\d+\.?\s", full) and headish(full, b, nxt): role = "h2"
        elif re.match(r"^[A-Z]\.\s", full) and headish(full, b, nxt): role = "h2"
        elif re.match(r"^\d+\.\s", full) and headish(full, b, nxt): role = "h3" if has_letter else "h2"
        elif re.match(r"^[a-z]\.\s", full) and headish(full, b, nxt) and (b.outline is not None): role = "h4" if has_letter else "h3"
        elif b.outline in (0, 1, 2, 3) and len(t) <= 200:
            role = "h%d" % max(2, b.outline + 1)
        if role:
            b.role = role
            if b.num and b.numlabel: b.prefix = b.numlabel
            k += 1; continue
        if b.outline is not None and len(t) > 200:
            b.note = "CEK_GABUNG"
        if re.match(r"^(\d+[.)]|[a-z][.)]|[-\u2022\u2013])\s", t) and not re.search(r":$", t):
            b.role = "daftar_manual"; k += 1; continue
        if "quote" in b.style or "kutipan" in b.style:
            b.role = "kutipan"; k += 1; continue
        b.role = "badan"
        k += 1
    # catatan pemeriksaan untuk AI
    for k, b in enumerate(blocks):
        if b.kind != "p" or b.role in ("buang", "sampul"): continue
        t = b.text
        if b.role in ("h2", "h3", "h4") and re.search(r"\s\d+\.\s+[A-Z]", t):
            b.note = "CEK_GABUNG"
        if (b.outline is not None) and b.role == "badan" and not b.note:
            b.note = "CEK_GABUNG"
        if b.role == "badan":
            txt = [s for s in b.segs if s.raw is None and s.t.strip()]
            if txt and all(s.b for s in txt) and len(t) < 120 and not b.note:
                b.note = "CEK_JUDUL"
            nx = blocks[k + 1] if k + 1 < len(blocks) else None
            if nx is not None and nx.kind == "p" and nx.role == "badan" and nx.text and t and not re.search(r"[.?!:;\"\u201d)]$", t):
                if nx.text[0].islower(): b.note = b.note or "CEK_TERPOTONG"
    return meta

# =====================================================================
# 4. PEMBANGUN XML
# =====================================================================
NS = ('xmlns:w="%s" xmlns:r="%s" xmlns:wp="%s" xmlns:a="%s" xmlns:pic="%s" xmlns:mc="%s" xmlns:m="%s" '
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
      'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
      'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
      'mc:Ignorable="w14 w15 wp14"' % (W, R, WP, A, PIC, MC, M))

def rpr(b=False, i=False, u=False, sup=False, sub=False, strike=False, sz=None, caps=False, style=None, vanish=False):
    x = ('<w:rStyle w:val="%s"/>' % style) if style else ""
    x += "<w:b/><w:bCs/>" if b else ""
    x += "<w:i/><w:iCs/>" if i else ""
    x += "<w:caps/>" if caps else ""
    x += "<w:strike/>" if strike else ""
    x += "<w:vanish/>" if vanish else ""
    x += ('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (sz, sz)) if sz else ""
    x += '<w:u w:val="single"/>' if u else ""
    x += '<w:vertAlign w:val="superscript"/>' if sup else ('<w:vertAlign w:val="subscript"/>' if sub else "")
    return "<w:rPr>%s</w:rPr>" % x if x else ""

def text_runs(t, rp):
    out = ""
    for part in re.split(r"(\t|\n)", t):
        if part == "\t": out += "<w:r>%s<w:tab/></w:r>" % rp
        elif part == "\n": out += "<w:r>%s<w:br/></w:r>" % rp
        elif part: out += '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rp, esc(part))
    return out

def segs_xml(segs, plainbold=None, sz=None, caps=False, heading=False, force_bold=None):
    out, cur_link, buf = "", None, ""
    def flush():
        nonlocal out, buf, cur_link
        if cur_link: out += '<w:hyperlink r:id="%s" w:history="1">%s</w:hyperlink>' % (cur_link, buf)
        else: out += buf
        buf = ""
    for s in segs:
        if s.link != cur_link:
            flush(); cur_link = s.link
        if s.raw is not None:
            rp = rpr(style="FootnoteReference") if s.note else rpr(sup=s.sup)
            buf += "<w:r>%s%s</w:r>" % (rp, s.raw)
            continue
        b = s.b
        if heading: b = False
        if force_bold is not None: b = force_bold
        rp = rpr(b=b, i=s.i, u=(s.u and not heading), sup=s.sup, sub=s.sub, strike=s.strike, sz=sz, caps=caps)
        buf += text_runs(s.t, rp)
    flush()
    return out

def ppr(style=None, keep=False, keeplines=False, brk=False, num=None, tabs=None, before=None, after=None, line=None,
        left=None, first=None, hanging=None, right=None, jc=None, sect=""):
    x = ('<w:pStyle w:val="%s"/>' % style) if style else ""
    x += "<w:keepNext/>" if keep else ""
    x += "<w:keepLines/>" if keeplines else ""
    x += "<w:pageBreakBefore/>" if brk else ""
    x += ('<w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%s"/></w:numPr>' % (num[1], num[0])) if num else ""
    x += tabs or ""
    if before is not None or after is not None or line is not None:
        a = (' w:before="%d"' % before) if before is not None else ""
        a += (' w:after="%d"' % after) if after is not None else ""
        a += (' w:line="%d" w:lineRule="auto"' % line) if line is not None else ""
        x += "<w:spacing%s/>" % a
    if left is not None or first is not None or hanging is not None or right is not None:
        a = (' w:left="%d"' % left) if left is not None else ""
        a += (' w:right="%d"' % right) if right is not None else ""
        a += (' w:hanging="%d"' % hanging) if hanging is not None else ""
        a += (' w:firstLine="%d"' % first) if first is not None and hanging is None else ""
        x += "<w:ind%s/>" % a
    x += ('<w:jc w:val="%s"/>' % jc) if jc else ""
    x += sect
    return "<w:pPr>%s</w:pPr>" % x if x else ""

def P(inner, attr="", **k):
    return "<w:p%s>%s%s</w:p>" % (attr, ppr(**k), inner)

MARK_NEW = ' w:rsidR="00A20000"'  # penanda paragraf sisipan; dilewati saat audit

def tc_field(text, ident, lvl=1):
    txt = text.replace('"', '\\"')
    r = lambda inner: "<w:r>%s%s</w:r>" % (rpr(vanish=True), inner)
    return (r('<w:fldChar w:fldCharType="begin"/>')
            + r('<w:instrText xml:space="preserve"> TC "%s" \\f %s \\l %d </w:instrText>' % (esc(txt), ident, lvl))
            + r('<w:fldChar w:fldCharType="end"/>'))

class Builder:
    def __init__(self, C, blocks, meta, reader, labels, fit):
        self.C, self.blocks, self.meta, self.reader, self.labels, self.fit = C, blocks, meta, reader, labels, fit
        m = C["margin"]
        self.ml, self.mt, self.mr, self.mb = (int(round(m[k] * CM)) for k in ("kiri", "atas", "kanan", "bawah"))
        self.TW = PAGE_W - self.ml - self.mr
        self.toc = {"MAIN": [], "Gambar": [], "Tabel": [], "Lampiran": []}
        self.order_front, self.order_body = [], []
        self.bm = 0
        self.count = {"h1": 0, "h2": 0, "h3": 0, "tabel": 0, "gambar": 0}
        self.has_cover = False

    def sect(self, hdr, ftr, fmt=None, start=None):
        pg = ""
        if fmt or start:
            pg = "<w:pgNumType%s%s/>" % ((' w:fmt="%s"' % fmt) if fmt else "", (' w:start="%d"' % start) if start else "")
        return ('<w:sectPr><w:headerReference w:type="default" r:id="%s"/><w:footerReference w:type="default" r:id="%s"/>'
                '<w:type w:val="nextPage"/><w:pgSz w:w="%d" w:h="%d"/><w:pgMar w:top="%d" w:right="%d" w:bottom="%d" '
                'w:left="%d" w:header="851" w:footer="851" w:gutter="0"/>%s<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>'
                % (hdr, ftr, PAGE_W, PAGE_H, self.mt, self.mr, self.mb, self.ml, pg))

    @staticmethod
    def attach(par, sect):
        i = par.rfind("</w:pPr>")
        if i < 0:
            j = par.find(">")
            return par[:j + 1] + "<w:pPr>" + sect + "</w:pPr>" + par[j + 1:]
        return par[:i] + sect + par[i:]

    # ---------- sampul ----------
    def cover(self):
        items = []
        for b in self.blocks:
            if b.role != "sampul": continue
            if b.kind == "tbl":
                items.append(("tabel", b)); continue
            if b.has_img:
                imgs = [s for s in b.segs if s.raw is not None]
                if imgs: items.append(("img", imgs))
            for ln_ in split_lines([s for s in b.segs if s.raw is None]):
                ln_ = tidy_ws([s.clone(s.t.replace("\t", " ")) for s in ln_])
                if text_only(ln_).strip(): items.append(("teks", ln_))
        if not items: return []
        self.has_cover = True
        kinds = []
        LABEL = re.compile(r"^(disusun\s+oleh|oleh|penyusun|dosen|pengampu|pembimbing|mata\s*kuliah|kelompok|kelas|nama|nim)\b.*:?$", re.I)
        INST = re.compile(r"^(program\s+studi|prodi|jurusan|departemen|fakultas|universitas|institut|sekolah\s+tinggi|politeknik|akademi|stai|stie|stmik|stkip|iain|uin)\b", re.I)
        DOCTYPE = re.compile(r"^(makalah|proposal|paper|tugas|laporan|skripsi|artikel)\b.{0,30}$", re.I)
        seen_title, seen_label, seen_inst = False, False, False
        for kind, v in items:
            if kind != "teks":
                kinds.append(kind); continue
            t = norm(text_only(v))
            if INST.match(t) or (re.match(r"^\d{4}$", t)) or (seen_inst and len(t) < 40):
                kinds.append("inst"); seen_inst = True
            elif LABEL.match(t) and t.endswith(":"):
                kinds.append("label"); seen_label = True
            elif LABEL.match(t) and len(t) < 90:
                kinds.append("info"); seen_label = True
            elif re.match(r"^(disusun\s+untuk|diajukan|untuk\s+memenuhi|sebagai\s+(salah\s+satu\s+)?syarat|tugas\s+mata)", t, re.I):
                kinds.append("teks"); seen_label = seen_label or seen_title
            elif not seen_title and not seen_label and DOCTYPE.match(t) and len(items) > 1 and "jenis" not in kinds:
                kinds.append("jenis")
            elif not seen_label and not seen_inst:
                kinds.append("judul"); seen_title = True
            else:
                kinds.append("teks")
        # perkiraan tinggi untuk membagi jarak antargrup
        tw_chars = max(20, int(self.TW / 190))
        H = {"judul": 400, "jenis": 340, "teks": 340, "label": 340, "info": 340, "inst": 320}
        used, gaps = 0, []
        for j, ((kind, v), kd) in enumerate(zip(items, kinds)):
            if kd in H:
                ch = len(text_only(v))
                used += H[kd] * max(1, -(-ch // (tw_chars if kd != "judul" else int(tw_chars * 0.72))))
            elif kd == "img": used += 2600
            elif kd == "tabel": used += 340 * max(1, len(v.el.findall(q("tr"))))
            prev = kinds[j - 1] if j else None
            if j and ((kd in ("label", "info") and prev not in ("label", "info")) or (kd == "label") or (kd == "inst" and prev != "inst") or kd in ("img", "tabel") or (kd == "teks" and prev == "judul")
                      or (prev == "judul" and kd != "judul") or prev == "img"):
                if not gaps or gaps[-1] != j: gaps.append(j)
        avail = PAGE_H - self.mt - self.mb
        free = max(0, avail - used - 200)
        wt = {j: (1.6 if kinds[j] == "inst" else 1.0) for j in gaps}
        tot = sum(wt.values()) or 1
        out = []
        for j, ((kind, v), kd) in enumerate(zip(items, kinds)):
            bf = 0
            if j in wt: bf = int(min(4800, free * 0.94 * wt[j] / tot) * self.fit)
            if kd == "img":
                inner = segs_xml(v)
                out.append(P(inner, jc="center", left=0, first=0, before=bf, after=0, line=240))
            elif kd == "tabel":
                if bf: out.append(P("", jc="center", left=0, first=0, before=bf, after=0, line=240))
                out.append(self.table_xml(v.el, border=False))
            elif kd == "judul":
                out.append(P(segs_xml(v, sz=32, caps=True, force_bold=True), jc="center", left=0, first=0, before=bf, after=0, line=276))
            elif kd == "jenis":
                out.append(P(segs_xml(v, sz=28, caps=True, force_bold=True), jc="center", left=0, first=0, before=bf, after=int(240 * self.fit), line=276))
            elif kd == "inst":
                out.append(P(segs_xml(v, sz=26, caps=True, force_bold=True), jc="center", left=0, first=0, before=bf, after=0, line=276))
            elif kd == "label":
                out.append(P(segs_xml(v, force_bold=True), jc="center", left=0, first=0, before=bf, after=60, line=276))
            else:
                out.append(P(segs_xml(v, force_bold=False), jc="center", left=0, first=0, before=bf, after=0, line=276))
        if self.C.get("_halaman_judul"):
            out[0] = out[0].replace("</w:pPr>", "</w:pPr>" + tc_field("HALAMAN JUDUL", "H"), 1)
        out[-1] = self.attach(out[-1], self.sect("rIdYZH1", "rIdYZF1", "lowerRoman", 1))
        return out

    # ---------- judul ----------
    def heading(self, level, segs, brk=False, prefix="", front=False, entry=True):
        segs = tidy_ws(split_flat(segs))
        pre = [Seg(prefix + " ")] if prefix else []
        allsegs = pre + segs
        title = norm(text_only(allsegs))
        self.bm += 1
        name = "_Toc%08d" % (70000000 + self.bm)
        if entry and level <= 3:
            self.toc["MAIN"].append((level, title, allsegs, name))
            (self.order_front if front else self.order_body).append(title)
        inner = '<w:bookmarkStart w:id="%d" w:name="%s"/>%s<w:bookmarkEnd w:id="%d"/>' % (self.bm, name, segs_xml(allsegs, heading=True), self.bm)
        self.count["h%d" % level] = self.count.get("h%d" % level, 0) + 1
        return P(inner, style="Heading%d" % level, brk=brk)

    def toc_block(self, key):
        items = self.toc[key]
        if key == "MAIN":
            instr = 'TOC \\o "1-3" \\h \\z \\u' + (' \\f H' if self.C.get("_halaman_judul") else "")
        else:
            instr = 'TOC \\f %s \\h \\z' % key[0]
        out = []
        for i, it in enumerate(items):
            lvl, title, segs, name = it
            style = "TOC%d" % lvl if key == "MAIN" else "TableofFigures"
            pre = ""
            if i == 0:
                pre = ('<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> %s </w:instrText></w:r>'
                       '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' % esc(instr))
            post = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if i == len(items) - 1 else ""
            label = self.labels.get(title, "")
            if name:
                pg = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF %s \\h </w:instrText></w:r>'
                      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>%s<w:r><w:fldChar w:fldCharType="end"/></w:r>' % (name, text_runs(label, "")))
                ent = '<w:hyperlink w:anchor="%s" w:history="1">%s<w:r><w:tab/></w:r>%s</w:hyperlink>' % (name, segs_toc(segs), pg)
            else:
                ent = segs_toc(segs) + "<w:r><w:tab/></w:r>" + text_runs(label, "")
            out.append(P(pre + ent + post, style=style))
        return out

    # ---------- tabel ----------
    def cell_par(self, p, head):
        tmp = []
        self.reader.inline(p, tmp)
        tmp = tidy_ws(tmp)
        pp = p.find(q("pPr")); jc = None; num = None
        if pp is not None:
            j = pp.find(q("jc"))
            if j is not None and j.get(q("val")) in ("center", "right"): jc = j.get(q("val"))
            np = pp.find(q("numPr"))
            if np is not None and np.find(q("numId")) is not None:
                il = np.find(q("ilvl"))
                num = (np.find(q("numId")).get(q("val")), int(il.get(q("val"))) if il is not None else 0)
        if head: jc = "center"
        if num:
            return P(segs_xml(tmp, force_bold=True if head else None), num=num, jc=jc or "left", left=284, hanging=284, before=20, after=20, line=240)
        return P(segs_xml(tmp, force_bold=True if head else None), jc=jc or "left", left=0, first=0, before=40, after=40, line=240)

    def table_xml(self, t, border=True):
        grid = [int(g.get(q("w"), "0") or 0) for g in t.findall(q("tblGrid") + "/" + q("gridCol"))]
        rows = t.findall(q("tr"))
        if not grid or sum(grid) <= 0:
            nc = max((sum(int((c.find(q("tcPr")).find(q("gridSpan")).get(q("val"))) if c.find(q("tcPr")) is not None and c.find(q("tcPr")).find(q("gridSpan")) is not None else 1) for c in r.findall(q("tc"))) for r in rows), default=1)
            grid = [self.TW // max(nc, 1)] * max(nc, 1)
        tot = sum(grid)
        k = self.TW / tot if tot > self.TW else 1.0
        tp = t.find(q("tblPr"))
        if tp is not None and tp.find(q("tblW")) is not None and tp.find(q("tblW")).get(q("type")) == "pct":
            k = self.TW / tot
        grid2 = [max(200, int(g * k)) for g in grid]
        gt = sum(grid2)
        bd = '<w:%s w:val="single" w:sz="4" w:space="0" w:color="000000"/>' if border else '<w:%s w:val="nil"/>'
        x = ('<w:tbl><w:tblPr><w:tblW w:w="%d" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>%s</w:tblBorders>'
             '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
             '</w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>%s</w:tblGrid>'
             % (gt, "".join(bd % s for s in ("top", "left", "bottom", "right", "insideH", "insideV")),
                "".join('<w:gridCol w:w="%d"/>' % g for g in grid2)))
        for ri, r in enumerate(rows):
            head = border and ri == 0 and len(rows) > 1
            x += "<w:tr><w:trPr><w:cantSplit/>%s</w:trPr>" % ("<w:tblHeader/>" if head else "")
            col = 0
            for c in r:
                if ln(c) == "sdt":
                    sc = c.find(q("sdtContent"))
                    cells = sc.findall(q("tc")) if sc is not None else []
                elif ln(c) == "tc":
                    cells = [c]
                else:
                    continue
                for tc in cells:
                    cp = tc.find(q("tcPr"))
                    span, vm, va = 1, "", ""
                    if cp is not None:
                        gs = cp.find(q("gridSpan"))
                        if gs is not None: span = int(gs.get(q("val")))
                        v = cp.find(q("vMerge"))
                        if v is not None:
                            vm = '<w:vMerge w:val="restart"/>' if v.get(q("val")) == "restart" else "<w:vMerge/>"
                    w_ = sum(grid2[col:col + span]) or grid2[-1]
                    col += span
                    shd = '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' if head else ""
                    gsx = ('<w:gridSpan w:val="%d"/>' % span) if span > 1 else ""
                    inner = ""
                    for e in tc:
                        if ln(e) == "p": inner += self.cell_par(e, head)
                        elif ln(e) == "tbl": inner += self.table_xml(e, border=border) + P("", left=0, first=0, line=240)
                    if not inner: inner = P("", left=0, first=0, line=240)
                    x += '<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s%s%s<w:vAlign w:val="center"/></w:tcPr>%s</w:tc>' % (w_, gsx, vm, shd, inner)
            x += "</w:tr>"
        return x + "</w:tbl>"

    # ---------- urutan keterangan tabel/gambar ----------
    def arrange(self, seq):
        """Judul tabel di atas tabel; judul gambar di bawah gambar; sisipkan judul kosong bila diminta."""
        out = list(seq)
        i = 0
        while i < len(out):
            b = out[i]
            nx = out[i + 1] if i + 1 < len(out) else None
            pv = out[i - 1] if i else None
            if b.kind == "tbl" and b.role != "sampul" and nx is not None and nx.role == "cap_tabel" and (pv is None or pv.role != "cap_tabel"):
                out[i], out[i + 1] = nx, b
                self.moved = True
                i += 1
            if b.role == "gambar" and pv is not None and pv.role == "cap_gambar" and (nx is None or nx.role != "cap_gambar"):
                out[i - 1], out[i] = b, pv
                self.moved = True
            i += 1
        if self.C["daftar_tabel_gambar"]:
            final, nt, ng = [], 0, 0
            for j, b in enumerate(out):
                nx = out[j + 1] if j + 1 < len(out) else None
                pv = out[j - 1] if j else None
                if b.role == "cap_tabel":
                    m = CAP_RX.match(b.text)
                    if m and m.group(2).isdigit(): nt = int(m.group(2))
                if b.role == "cap_gambar":
                    m = CAP_RX.match(b.text)
                    if m and m.group(2).isdigit(): ng = int(m.group(2))
                if b.kind == "tbl" and b.role != "sampul" and not (pv is not None and pv.role == "cap_tabel"):
                    nt += 1
                    c = Block("baru"); c.role = "cap_tabel"; c.section = b.section; c.segs = [Seg("Tabel %d. [Judul Tabel]" % nt)]; c.new = True
                    final.append(c); self.inserted.append("Tabel %d" % nt)
                final.append(b)
                if b.role == "gambar" and not (nx is not None and nx.role == "cap_gambar"):
                    ng += 1
                    c = Block("baru"); c.role = "cap_gambar"; c.section = b.section; c.segs = [Seg("Gambar %d. [Judul Gambar]" % ng)]; c.new = True
                    final.append(c); self.inserted.append("Gambar %d" % ng)
            out = final
        return out

    # ---------- blok isi ----------
    def block_xml(self, b, brk, front):
        role = b.role
        if b.kind == "tbl":
            self.count["tabel"] += 1
            gap = [] if getattr(b, "next_role", None) == "sumber" else [P("", left=0, first=0, before=0, after=0, line=240)]
            return [self.table_xml(b.el)] + gap
        segs = b.segs
        if role == "h1":
            return [self.heading(1, segs, brk=brk, prefix=b.prefix, front=front)]
        if role in ("h2", "h3", "h4"):
            return [self.heading(int(role[1]), segs, prefix=b.prefix, front=front, entry=role != "h4")]
        if role == "cap_tabel" or role == "cap_gambar" or role == "cap_lampiran":
            segs = tidy_ws(split_flat(segs))
            key = {"cap_tabel": "Tabel", "cap_gambar": "Gambar", "cap_lampiran": "Lampiran"}[role]
            title = norm(text_only(segs))
            tc = ""
            if self.C["daftar_tabel_gambar"] or (key == "Lampiran" and "DAFTAR LAMPIRAN" in self.meta["ada"]):
                self.toc[key].append((1, title, segs, None))
                tc = tc_field(title, key[0])
                self.order_body.append(title)
            attr = MARK_NEW if getattr(b, "new", False) else ""
            if role == "cap_gambar":
                self.count["gambar"] = self.count.get("gambar", 0)
                return [P(segs_xml(segs) + tc, attr=attr, style="Caption", before=120, after=120, line=240)]
            if role == "cap_tabel":
                return [P(segs_xml(segs) + tc, attr=attr, style="Caption", keep=True, before=240, after=120, line=240)]
            return [P(segs_xml(segs, force_bold=True) + tc, style="Caption", jc="left", keep=True, before=0, after=120, line=360)]
        if role == "gambar":
            self.count["gambar"] += 1
            return [P(segs_xml(tidy_ws(segs)), jc="center", left=0, first=0, before=240, after=120, line=240, keep=True)]
        if role == "sumber":
            return [P(segs_xml(tidy_ws(split_flat(segs)), sz=20), jc="center", left=0, first=0, before=60, after=240, line=240)]
        if role == "pustaka":
            return [P(segs_xml(tidy_ws(split_flat(segs))), jc="left", left=720, hanging=720, before=0, after=240, line=240, keeplines=True)]
        if role == "kutipan":
            return [P(segs_xml(tidy_ws(split_flat(segs))), jc="both", left=720, right=720, first=0, before=120, after=120, line=240)]
        if role == "daftar":
            lvl = b.num[1] if b.num else 0
            return [P(segs_xml(tidy_ws(split_flat(segs))), num=b.num, left=720 + 360 * lvl, hanging=360, jc="both")]
        if role == "daftar_manual":
            segs = tidy_ws(split_flat(segs))
            if segs and segs[0].raw is None:
                segs[0].t = re.sub(r"^(\S+)\s+", lambda m: m.group(1) + "\t", segs[0].t, count=1)
            return [P(segs_xml(segs), left=720, hanging=360, jc="both")]
        # badan: baris baru di tengah paragraf
        lines = split_lines(segs)
        parts, cur = [], []
        for ln_ in lines:
            t = text_only(ln_).strip()
            if not t and not any(s.raw is not None for s in ln_): continue
            cur = cur + ([Seg(" ")] if cur else []) + ln_
            if re.search(r"[.?!:;\"\u201d)]$", t) or ln_ is lines[-1]:
                parts.append(cur); cur = []
        if cur: parts.append(cur)
        out = []
        for seg_ in parts:
            seg_ = tidy_ws([s.clone(s.t.replace("\t", " ")) if s.raw is None else s for s in seg_])
            t = norm(text_only(seg_))
            if b.jc == "right" and len(t) < 120:
                out.append(P(segs_xml(seg_), jc="right", left=0, first=0))
            elif b.jc == "center" and len(t) < 120:
                out.append(P(segs_xml(seg_), jc="center", left=0, first=0))
            else:
                out.append(P(segs_xml(seg_)))
        return out

    def build(self):
        self.moved = False; self.inserted = []
        C = self.C
        out = self.cover()
        seq = [b for b in self.blocks if b.role not in ("sampul", "buang", "judul_daftar")]
        seq = self.arrange(seq)
        self.final_seq = seq
        for j, b in enumerate(seq):
            b.next_role = seq[j + 1].role if j + 1 < len(seq) else None
        front_seq = [b for b in seq if b.section == "front"]
        body_seq = [b for b in seq if b.section != "front"]
        body, first = [], True
        for b in body_seq:
            brk = b.role == "h1" and not first
            body += self.block_xml(b, brk, False)
            if b.role == "h1": first = False
        main_body = self.toc["MAIN"]; self.toc["MAIN"] = []
        front, first = [], True
        for b in front_seq:
            brk = b.role == "h1" and not first
            front += self.block_xml(b, brk, True)
            if b.role == "h1": first = False
        lists = [("DAFTAR ISI", "MAIN")] if (main_body or front) else []
        if C["daftar_tabel_gambar"]:
            if self.toc["Tabel"]: lists.append(("DAFTAR TABEL", "Tabel"))
            if self.toc["Gambar"]: lists.append(("DAFTAR GAMBAR", "Gambar"))
        if self.toc["Lampiran"]: lists.append(("DAFTAR LAMPIRAN", "Lampiran"))
        head = []
        if C.get("_halaman_judul") and self.has_cover:
            head.append((1, "HALAMAN JUDUL", [Seg("HALAMAN JUDUL")], None))
            self.labels["HALAMAN JUDUL"] = "i"
        front_main = self.toc["MAIN"]; self.toc["MAIN"] = []
        for title, key in lists:
            self.bm += 1
            name = "_Toc%08d" % (70000000 + self.bm)
            front_main.append((1, title, [Seg(title)], name))
            self.order_front.append(title)
            self._bm = getattr(self, "_bm", {}); self._bm[title] = (self.bm, name)
        self.toc["MAIN"] = head + front_main + main_body
        for title, key in lists:
            idn, name = self._bm[title]
            hx = '<w:bookmarkStart w:id="%d" w:name="%s"/>%s<w:bookmarkEnd w:id="%d"/>' % (idn, name, text_runs(title, ""), idn)
            title_p = P(hx, style="Heading1", brk=bool(front))
            if key == "MAIN":
                front.append('<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr>'
                             '<w:sdtContent>%s%s</w:sdtContent></w:sdt>' % (title_p, "".join(self.toc_block(key))))
            else:
                front.append(title_p)
                front += self.toc_block(key)
        cover_fmt = "lowerRoman"
        sect_front = self.sect("rIdYZH2", "rIdYZF2", cover_fmt, None if self.has_cover else 1)
        if front:
            # paragraf mungil pembawa batas bagian, supaya Daftar Isi tetap utuh sebagai satu blok
            front.append('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:ind w:firstLine="0"/>'
                         '<w:rPr><w:sz w:val="2"/></w:rPr>%s</w:pPr></w:p>' % sect_front)
        final = self.sect("rIdYZH3", "rIdYZF3", "decimal", 1)
        return "".join(out + front + body) + final

    def hf(self, kind, with_num):
        tag = "hdr" if kind == "header" else "ftr"
        jc = {"kiri": "left", "tengah": "center", "kanan": "right"}[self.C["pos_h"]]
        fld = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
               '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:%s %s>%s</w:%s>'
                % (tag, NS, P(fld if with_num else "", jc=jc, left=0, first=0, before=0, after=0, line=240), tag))

def split_flat(segs):
    """Baris baru di dalam judul/daftar menjadi spasi (tata letak, bukan isi)."""
    return [s if s.raw is not None else s.clone(s.t.replace("\n", " ").replace("\t", " ")) for s in segs]

def segs_toc(segs):
    out = ""
    for s in segs:
        if s.raw is not None: continue
        out += text_runs(s.t, rpr(i=s.i))
    return out

def styles_xml(tw):
    F = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'
    def ps(sid, name, ppr_, rpr_="", based="Normal", q_=True):
        return ('<w:style w:type="paragraph" w:styleId="%s"><w:name w:val="%s"/>%s<w:next w:val="Normal"/>%s'
                '<w:pPr>%s</w:pPr><w:rPr>%s</w:rPr></w:style>'
                % (sid, name, ('<w:basedOn w:val="%s"/>' % based) if based else "", "<w:qFormat/>" if q_ else "", ppr_, rpr_))
    tab = '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="%d"/></w:tabs>' % tw
    sp = lambda b, a, l: '<w:spacing w:before="%d" w:after="%d" w:line="%d" w:lineRule="auto"/>' % (b, a, l)
    black = '<w:color w:val="000000"/>'
    toc = lambda sid, name, left, hang: ps(sid, name, tab + sp(0, 0, 360) + '<w:ind w:left="%d" w:right="567" w:hanging="%d"/><w:jc w:val="left"/>' % (left + hang, hang), "", q_=False)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="%s">' % W
            + '<w:docDefaults><w:rPrDefault><w:rPr>%s<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="id-ID" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
              '<w:pPrDefault><w:pPr>%s</w:pPr></w:pPrDefault></w:docDefaults>' % (F, sp(0, 0, 360))
            + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:widowControl/>%s<w:ind w:firstLine="720"/><w:jc w:val="both"/></w:pPr><w:rPr>%s<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' % (sp(0, 0, 360), F)
            + ps("Heading1", "heading 1", '<w:keepNext/><w:keepLines/>' + sp(0, 480, 360) + '<w:ind w:firstLine="0"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/>', "<w:b/><w:bCs/>" + black + '<w:sz w:val="28"/><w:szCs w:val="28"/>')
            + ps("Heading2", "heading 2", '<w:keepNext/><w:keepLines/>' + sp(240, 0, 360) + '<w:ind w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="1"/>', "<w:b/><w:bCs/>" + black)
            + ps("Heading3", "heading 3", '<w:keepNext/><w:keepLines/>' + sp(120, 0, 360) + '<w:ind w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="2"/>', "<w:b/><w:bCs/><w:i/><w:iCs/>" + black)
            + ps("Heading4", "heading 4", '<w:keepNext/><w:keepLines/>' + sp(120, 0, 360) + '<w:ind w:firstLine="0"/><w:jc w:val="left"/><w:outlineLvl w:val="3"/>', "<w:i/><w:iCs/>" + black)
            + ps("Caption", "caption", sp(120, 0, 240) + '<w:ind w:firstLine="0"/><w:jc w:val="center"/>', black)
            + toc("TOC1", "toc 1", 0, 0)
            + toc("TOC2", "toc 2", 567, 284)
            + toc("TOC3", "toc 3", 1134, 284)
            + ps("TableofFigures", "table of figures", tab + sp(0, 0, 360) + '<w:ind w:left="0" w:right="567" w:firstLine="0"/><w:jc w:val="left"/>', "", q_=False)
            + ps("FootnoteText", "footnote text", sp(0, 0, 240) + '<w:ind w:firstLine="0"/><w:jc w:val="both"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', q_=False)
            + '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>'
            + '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
            + '</w:styles>')

def clean_notes(xml, reader):
    """Rapikan catatan kaki: font & ukuran seragam, kata tidak berubah."""
    root = ET.fromstring(xml)
    root.attrib.pop("{%s}Ignorable" % MC, None)  # awalan yang tidak dipakai ikut hilang saat ditulis ulang
    for fn in root:
        if fn.get(q("type")) in ("separator", "continuationSeparator", "continuationNotice"): continue
        for p in fn.findall(q("p")):
            pp = p.find(q("pPr"))
            if pp is not None: p.remove(pp)
            pp = ET.Element(q("pPr"))
            st = ET.SubElement(pp, q("pStyle")); st.set(q("val"), "FootnoteText")
            p.insert(0, pp)
            for r in p.iter(q("r")):
                old = r.find(q("rPr"))
                d, _ = reader.fmt_of(old)
                isref = any(ln(c) in ("footnoteRef", "endnoteRef") for c in r)
                if old is not None: r.remove(old)
                new = ET.fromstring("<x xmlns:w=\"%s\">%s</x>" % (W, rpr(style="FootnoteReference") if isref else rpr(b=d["b"], i=d["i"], u=d["u"], sup=d["sup"], sub=d["sub"])))
                if len(new): r.insert(0, new[0])
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ET.tostring(root, encoding="unicode")).encode("utf-8")

# =====================================================================
# 5. PAKET DOCX
# =====================================================================
REPLACED = ("word/document.xml", "word/styles.xml", "word/settings.xml", "word/stylesWithEffects.xml")
DROP_TYPES = ("/header", "/footer", "/comments", "/commentsExtended", "/commentsIds", "/commentsExtensible", "/people", "/stylesWithEffects")

def package(D, src_parts, out):
    docxml = D.build()
    C = D.C
    document = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document %s><w:body>%s</w:body></w:document>' % (NS, docxml)
    parts = dict(src_parts)
    relname = "word/_rels/document.xml.rels"
    rel_root = ET.fromstring(parts.get(relname, b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'))
    RN = "http://schemas.openxmlformats.org/package/2006/relationships"
    drop = []
    used = document
    keep_rels = []
    for r in list(rel_root):
        ty, tg = r.get("Type"), r.get("Target")
        if any(ty.endswith(x) for x in DROP_TYPES):
            drop.append("word/" + tg.lstrip("/").replace("word/", ""))
            continue
        keep_rels.append(r)
    for r in rel_root.findall("{%s}Relationship" % RN):
        pass
    have = {r.get("Type").split("/")[-1] for r in keep_rels}
    top = C["pos_v"] == "atas"
    newparts = {
        "word/yzheader1.xml": D.hf("header", False), "word/yzfooter1.xml": D.hf("footer", False),
        "word/yzheader2.xml": D.hf("header", top), "word/yzfooter2.xml": D.hf("footer", not top),
        "word/yzheader3.xml": D.hf("header", top), "word/yzfooter3.xml": D.hf("footer", not top),
    }
    relx = ""
    BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    for r in keep_rels:
        mode = (' TargetMode="%s"' % r.get("TargetMode")) if r.get("TargetMode") else ""
        relx += '<Relationship Id="%s" Type="%s" Target="%s"%s/>' % (esc(r.get("Id")), esc(r.get("Type")), esc(r.get("Target")), mode)
    if "styles" not in have: relx += '<Relationship Id="rIdYZS" Type="%sstyles" Target="styles.xml"/>' % BASE
    if "settings" not in have: relx += '<Relationship Id="rIdYZT" Type="%ssettings" Target="settings.xml"/>' % BASE
    for i in (1, 2, 3):
        relx += '<Relationship Id="rIdYZH%d" Type="%sheader" Target="yzheader%d.xml"/>' % (i, BASE, i)
        relx += '<Relationship Id="rIdYZF%d" Type="%sfooter" Target="yzfooter%d.xml"/>' % (i, BASE, i)
    parts[relname] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">%s</Relationships>' % (RN, relx)).encode("utf-8")
    for d in drop:
        parts.pop(d, None)
        parts.pop(d.replace("word/", "word/_rels/") + ".rels", None)
    for k, v in newparts.items(): parts[k] = v.encode("utf-8")
    parts["word/document.xml"] = document.encode("utf-8")
    parts["word/styles.xml"] = styles_xml(D.TW).encode("utf-8")
    has_fn = "word/footnotes.xml" in parts; has_en = "word/endnotes.xml" in parts
    fnpr = ""
    if has_fn: fnpr += '<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>'
    if has_en: fnpr += '<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>'
    parts["word/settings.xml"] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="%s">'
                                  '<w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:updateFields w:val="true"/>%s'
                                  '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
                                  '</w:settings>' % (W, fnpr)).encode("utf-8")
    if "word/numbering.xml" in parts:
        n = parts["word/numbering.xml"].decode("utf-8", "ignore")
        parts["word/numbering.xml"] = re.sub(r"<w:pStyle\b[^>]*/>", "", n).encode("utf-8")
    for nm in ("word/footnotes.xml", "word/endnotes.xml"):
        if nm in parts: parts[nm] = clean_notes(parts[nm], D.reader)
    # content types
    ct = parts["[Content_Types].xml"].decode("utf-8")
    for d in drop:
        ct = re.sub(r'<Override[^>]*PartName="/%s"[^>]*/>' % re.escape(d), "", ct)
    ct = re.sub(r'<Override[^>]*PartName="/word/stylesWithEffects.xml"[^>]*/>', "", ct)
    ins = ""
    for k in ("header", "footer"):
        for i in (1, 2, 3):
            ins += '<Override PartName="/word/yz%s%d.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.%s+xml"/>' % (k, i, k)
    ct = ct.replace("</Types>", ins + "</Types>")
    parts["[Content_Types].xml"] = ct.encode("utf-8")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for k, v in parts.items():
            z.writestr(k, v)

# =====================================================================
# 6. RENDER -> NOMOR HALAMAN
# =====================================================================
def render_pdf(docx, outdir):
    pdf = os.path.join(outdir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    args = ["--headless", "--convert-to", "pdf", "--outdir", outdir, docx]
    for skdir in ("/mnt/skills/public/docx/scripts",):
        if os.path.isdir(os.path.join(skdir, "office")):
            try:
                sys.path.insert(0, skdir)
                from office.soffice import run_soffice
                run_soffice(args, capture_output=True, timeout=300)
                if os.path.exists(pdf): return pdf
            except Exception:
                pass
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe: return None
    prof = tempfile.mkdtemp(prefix="yz_lo_")
    try:
        subprocess.run([exe, "-env:UserInstallation=file://" + prof] + args, capture_output=True, timeout=300,
                       env=dict(os.environ, SAL_USE_VCLPLUGIN="svp"))
    except Exception:
        return None
    return pdf if os.path.exists(pdf) else None

def analyze(pdf, D):
    if not shutil.which("pdftotext"): return None
    t = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
    pages = t.split("\f")
    if pages and not pages[-1].strip(): pages = pages[:-1]
    lines = [[norm(l) for l in p.split("\n") if norm(l)] for p in pages]
    joined = [" ".join(l) for l in lines]
    loc, cur = {}, 0
    for e in D.order_front:
        k = norm(e)
        for p in range(cur, len(pages)):
            if k in lines[p] or k.upper() in lines[p]:
                loc[e] = p; cur = p; break
    first_front = min(loc.values()) if loc else None
    if D.order_body:
        for e in D.order_body:
            k = norm(e)
            found = None
            for p in range(cur, len(pages)):
                if any(l.startswith(k[:45]) and not re.search(r"[.\u2026]{3,}", l) for l in lines[p]):
                    found = p; break
                s = joined[p]
                for m in re.finditer(re.escape(k), s):
                    if not re.match(r"\s*[.\u2026]{3,}", s[m.end():]):
                        found = p; break
                if found is not None: break
            if found is not None:
                loc[e] = found; cur = found
    if not D.order_body or D.order_body[0] not in loc: return None
    b1 = loc[D.order_body[0]]
    lab = {e: (roman(p + 1, True) if p < b1 else str(p - b1 + 1)) for e, p in loc.items()}
    if D.C.get("_halaman_judul"): lab["HALAMAN JUDUL"] = "i"
    missing = [e for e in D.order_front + D.order_body if e not in loc]
    cover_ok = (first_front == 1) if D.has_cover else True
    return {"labels": lab, "cover_ok": cover_ok, "pages": len(pages), "missing": missing}

# =====================================================================
# 7. AUDIT KATA PER KATA
# =====================================================================
def audit(src_blocks_final, koreksi, out_path):
    expected = []
    for sid, text in src_blocks_final:
        expected.append(text)
    parts = read_parts(out_path)
    rd = Reader(parts)
    root = ET.fromstring(parts["word/document.xml"])
    got = []
    body = root.find(q("body"))
    for c in body:
        if ln(c) == "tbl":
            for tp in c.iter(q("p")):
                tmp = []; rd.inline(tp, tmp); got.append(text_only(tmp))
            continue
        if ln(c) != "p": continue
        if c.get(q("rsidR")) == "00A20000": continue
        pp = c.find(q("pPr")); st = ""
        if pp is not None and pp.find(q("pStyle")) is not None: st = pp.find(q("pStyle")).get(q("val"))
        if st.startswith("TOC") or st == "TableofFigures": continue
        tmp = []; rd.inline(c, tmp); t = text_only(tmp)
        if st == "Heading1" and norm(t) in LISTS: continue
        got.append(t)
    a, b = " ".join(expected).split(), " ".join(got).split()
    if a == b: return {"cocok": True, "jumlah_kata": len(b), "selisih": []}
    diff = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag != "equal":
            diff.append({"harusnya": " ".join(a[i1:i2])[:200], "tertulis": " ".join(b[j1:j2])[:200]})
    return {"cocok": False, "jumlah_kata": len(b), "selisih": diff[:30]}

# =====================================================================
# 8. CONFIG & PERINTAH
# =====================================================================
def parse_margin(m):
    d = {"kiri": 4.0, "atas": 3.0, "kanan": 3.0, "bawah": 3.0}
    vals = None
    if isinstance(m, dict):
        try: vals = [float(str(m.get(k, d[k])).replace(",", ".")) for k in ("kiri", "atas", "kanan", "bawah")]
        except ValueError: vals = None
    elif isinstance(m, str):
        nums = re.findall(r"\d+(?:\.\d+)?", m.replace(" ", ""))
        nums = re.findall(r"\d+", m) if len(nums) != 4 else nums
        if len(nums) == 4: vals = [float(x) for x in nums]
    if vals is None or not all(1.0 <= v <= 6.0 for v in vals):
        if m not in (None, "", {}): warn("Format margin tidak dikenali, dipakai 4-3-3-3.")
        return d
    return dict(zip(("kiri", "atas", "kanan", "bawah"), vals))

POSISI = {"1": "kiri atas", "2": "tengah atas", "3": "kanan atas", "4": "kiri bawah", "5": "tengah bawah", "6": "kanan bawah"}

def parse_posisi(p):
    s = str(p or "").lower().replace("-", " ").replace("_", " ").strip()
    s = POSISI.get(s, s)
    h = "kiri" if "kiri" in s else "kanan" if "kanan" in s else "tengah"
    v = "atas" if "atas" in s else "bawah"
    return h, v

def load_blocks(parts, C, tw=None):
    rd = Reader(parts, tw)
    root = ET.fromstring(parts["word/document.xml"])
    blocks = rd.body(root)
    for f in [x for x in parts if re.match(r"word/(header|footer)\d*\.xml$", x)]:
        tx = "".join(re.findall(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", parts[f].decode("utf-8", "ignore"))).strip()
        if re.search(r"[A-Za-z]{3,}", tx):
            warn("Header/footer lama berisi teks (\"%s\") dan diganti dengan nomor halaman saja." % tx[:60])
    return rd, blocks

def apply_ops(blocks, C):
    """Urutan: gabung -> koreksi -> pecah -> peran paksa."""
    byid = {b.id: b for b in blocks}
    for pair in C.get("gabung", []):
        ids = [str(x) for x in pair]
        if not all(i in byid for i in ids): warn("Blok gabung tidak ditemukan: %s" % ids); continue
        head = byid[ids[0]]
        for i in ids[1:]:
            o = byid[i]
            head.segs += [Seg(" ")] + o.segs
            head.asli += " " + o.asli
            blocks.remove(o); del byid[i]
    applied, missed = [], []
    C["_ditolak"] = []
    for k in C.get("koreksi", []):
        cari, ganti = k.get("cari", ""), k.get("ganti", "")
        if not cari: continue
        rasio = difflib.SequenceMatcher(None, cari.lower(), ganti.lower()).ratio()
        if rasio < 0.55 or abs(len(cari.split()) - len(ganti.split())) > 1:
            C["_ditolak"].append(cari); continue
        kb = str(k.get("blok", "")).split(".")[0]
        targets = [byid[kb]] if "blok" in k and kb in byid else ([] if "blok" in k else blocks)
        hit = 0
        L = r"(?<!\w)" if cari[0].isalnum() else ""
        Rr = r"(?!\w)" if cari[-1].isalnum() else ""
        rx = re.compile(L + re.escape(cari) + Rr)
        for b in targets:
            if b.kind != "p": continue
            n = 0
            for m in reversed(list(rx.finditer(plain(b.segs)))):
                replace_range(b.segs, m.start(), m.end(), ganti); n += 1
            if n:
                hit += n
                b.asli = rx.sub(lambda _m: ganti, b.asli)
        (applied if hit else missed).append({"cari": cari, "ganti": ganti, "jumlah": hit, "alasan": k.get("alasan", "")})
    for s in C.get("pecah", []):
        bid = str(s.get("blok")).split(".")[0]
        if bid not in byid: warn("Blok pecah tidak ditemukan: %s" % bid); continue
        b = byid[bid]
        p = plain(b.segs)
        cut = s.get("setelah", "")
        i = p.find(cut)
        if i < 0: warn("Teks pecah tidak ditemukan pada blok %s" % bid); continue
        pos = i + len(cut)
        k = split_seg_at(b.segs, pos)
        left, right = b.segs[:k], b.segs[k:]
        nb = Block(bid + ".1")
        nb.segs = tidy_ws(right); nb.style = "normal"; nb.outline = None; nb.jc = None; nb.src = b.src
        b.segs = tidy_ws(left)
        pr = s.get("peran", [])
        if len(pr) > 0 and pr[0]: b.role = pr[0]; b.forced = True
        if len(pr) > 1 and pr[1]: nb.role = pr[1]; nb.forced = True
        nb.asli = ""; 
        blocks.insert(blocks.index(b) + 1, nb)
        byid[nb.id] = nb
    return applied, missed

def apply_roles(blocks, C):
    byid = {b.id: b for b in blocks}
    for bid, role in (C.get("peran") or {}).items():
        bid = str(bid)
        if bid in byid:
            byid[bid].role = role; byid[bid].forced = True
        else:
            warn("Blok peran tidak ditemukan: %s" % bid)

def toc_entries(blocks):
    out = set()
    for b in blocks:
        if b.kind == "p" and is_toc_line(b):
            e = re.split(r"\t|[.\u2026]{2,}", text_only(b.segs))[0]
            e = norm(re.sub(r"\s+\d{1,3}\s*$", "", e))
            if len(e) >= 4: out.add(e)
    return out

def auto_split(blocks):
    entries = sorted(toc_entries(blocks), key=len, reverse=True)
    done = []
    for b in list(blocks):
        if b.kind != "p" or b.forced or b.outline is None or is_toc_line(b): continue
        p = plain(b.segs).replace("\n", " ")
        for e in entries:
            if p.startswith(e + " ") and len(p) > len(e) + 1:
                k = split_seg_at(b.segs, len(e))
                left, right = b.segs[:k], b.segs[k:]
                nb = Block(b.id + ".1")
                nb.segs = tidy_ws(right); nb.style = "normal"; nb.src = b.src
                b.segs = tidy_ws(left)
                blocks.insert(blocks.index(b) + 1, nb)
                done.append(e)
                break
    return done

def auto_italic_terms(blocks):
    """Istilah yang penulis sendiri miringkan, diseragamkan ke seluruh badan teks."""
    terms = set()
    for b in blocks:
        if b.kind != "p": continue
        buf = ""
        for s in b.segs + [Seg("")]:
            if s.raw is None and s.i and s.t: buf += s.t
            else:
                t = norm(buf); buf = ""
                wc = len(t.split())
                if 3 <= len(t) and wc <= 5 and t[-1:].isalnum() and (wc >= 2 or b.outline is not None or is_toc_line(b)):
                    terms.add(t)
    return sorted(terms, key=len)

def mark_italic_ci(segs, term):
    rx = re.compile(r"(?<![\w\-])" + re.escape(term) + r"(?![\w\-])", re.I)
    p = plain(segs)
    for m in reversed(list(rx.finditer(p))):
        a = split_seg_at(segs, m.start())
        b = split_seg_at(segs, m.end())
        a = split_seg_at(segs, m.start())
        for s in segs[a:b]:
            if s.raw is None: s.i = True

def sort_pustaka(blocks):
    idx = [k for k, b in enumerate(blocks) if b.role == "pustaka" and not b.num]
    if len(idx) < 2 or idx != list(range(idx[0], idx[-1] + 1)): return False
    key = lambda b: (re.findall(r"[^\W\d_]+", b.text.lower()) or [""])[0]
    items = [blocks[k] for k in idx]
    srt = sorted(items, key=key)
    if [id(x) for x in srt] == [id(x) for x in items]: return False
    blocks[idx[0]:idx[-1] + 1] = srt
    return True

def cmd_baca(path, naskah):
    parts = read_parts(path)
    C = {}
    rd, blocks = load_blocks(parts, C)
    auto_split(blocks)
    meta = classify(blocks)
    lines = []
    for b in blocks:
        if b.kind == "tbl":
            lines.append("[%s|tabel] %s" % (b.id, b.asli[:300])); continue
        tag = b.role or "?"
        if b.has_img: tag += "+objek"
        if b.note: tag += "|" + b.note
        tx = text_only(b.segs).replace("\n", " \u23ce ").replace("\t", " \u2192 ")
        if b.role == "buang" and not tx.strip(): continue
        lines.append("[%s|%s] %s" % (b.id, tag, tx))
    open(naskah, "w", encoding="utf-8").write("\n".join(lines))
    print("NASKAH=" + naskah)
    print("BLOK=%d TABEL=%d GAMBAR=%d CATATAN_KAKI=%s" % (len(blocks), sum(b.kind == "tbl" for b in blocks),
          sum(b.role == "gambar" for b in blocks), "word/footnotes.xml" in parts))
    for w in WARN: print("PERINGATAN:", w)

def cmd_rapikan(path, cfgpath, outdir):
    try:
        raw = json.load(open(cfgpath, encoding="utf-8-sig"))
    except Exception as e:
        print("CONFIG_ERROR: config.json tidak valid -> %s" % e); sys.exit(2)
    C = dict(raw)
    C["margin"] = parse_margin(raw.get("margin_cm", raw.get("margin")))
    C["pos_h"], C["pos_v"] = parse_posisi(raw.get("posisi_nomor"))
    C["daftar_tabel_gambar"] = str(raw.get("daftar_tabel_gambar", False)).strip().lower() in ("true", "ya", "y", "yes", "1")
    stem = re.sub(r"[^A-Za-z0-9_\-]+", "_", os.path.splitext(os.path.basename(path))[0]).strip("_") or "Makalah"
    name = re.sub(r"[^A-Za-z0-9_\-]+", "_", re.sub(r"\.docx?$", "", str(raw.get("nama_file") or stem + "_Rapi"), flags=re.I)).strip("_") + ".docx"
    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, name)
    tmp = tempfile.mkdtemp(prefix="yz2_")
    parts = read_parts(path)
    TW = PAGE_W - int(round(C["margin"]["kiri"] * CM)) - int(round(C["margin"]["kanan"] * CM))
    labels, fit, res, status = {}, 1.0, None, "TANPA_RENDER"
    miring = sorted({x for x in raw.get("miring", []) if x}, key=len)
    miring_p = sorted({x for x in raw.get("miring_pustaka", []) if x}, key=len)
    applied, missed, sorted_flag, D, meta = [], [], False, None, None
    work = os.path.join(tmp, name)
    for attempt in range(7):
        WARN_before = list(WARN)
        rd, blocks = load_blocks(parts, C, TW)
        applied, missed = apply_ops(blocks, C)
        terms_auto = auto_italic_terms(blocks)
        split_auto = auto_split(blocks)
        apply_roles(blocks, C)
        meta = classify(blocks)
        toc_old = [norm(re.sub(r"[.\u2026]{2,}", " ", t)) for t in meta["toc_lama"]]
        C["_halaman_judul"] = any(re.match(r"^(halaman\s+judul|halaman\s+sampul|sampul)\b", t, re.I) for t in toc_old)
        sorted_flag = sort_pustaka(blocks)
        for b in blocks:
            if b.kind != "p" or b.role in ("sampul", "buang", "judul_daftar"): continue
            if b.role != "pustaka":
                for term in terms_auto: mark_italic_ci(b.segs, term)
            for term in (miring_p if b.role == "pustaka" else miring):
                mark_italic(b.segs, term)
        D = Builder(C, blocks, meta, rd, dict(labels), fit)
        package(D, parts, work)
        pdf = render_pdf(work, tmp)
        res = analyze(pdf, D) if pdf else None
        if res is None: break
        if not res["cover_ok"] and fit > 0.3:
            fit = round(fit * 0.75, 3); continue
        if res["labels"] == labels:
            status = "OK"; break
        labels = res["labels"]; status = "BELUM_STABIL"
    if status == "BELUM_STABIL":
        rd, blocks = load_blocks(parts, C, TW)
        apply_ops(blocks, C); terms_auto = auto_italic_terms(blocks); auto_split(blocks); apply_roles(blocks, C); meta = classify(blocks); sort_pustaka(blocks)
        for b in blocks:
            if b.kind != "p" or b.role in ("sampul", "buang", "judul_daftar"): continue
            if b.role != "pustaka":
                for term in terms_auto: mark_italic_ci(b.segs, term)
            for term in (miring_p if b.role == "pustaka" else miring):
                mark_italic(b.segs, term)
        D = Builder(C, blocks, meta, rd, dict(labels), fit)
        package(D, parts, work)
    # audit
    order = []
    for b in blocks:
        if b.role == "sampul": order.append((b.id, b.asli))
    for sec in ("front", "isi"):
        for b in D.final_seq:
            if getattr(b, "new", False): continue
            if (b.section == "front") == (sec == "front") and not b.id.endswith(".1"):
                order.append((b.id, b.asli))
    aud = audit(order, C.get("koreksi", []), work)
    if aud["cocok"] and not C.get("_ditolak") and not missed: shutil.copy(work, out)
    heads = {norm(x[1]).lower() for k in D.toc for x in D.toc[k]}
    no_body = []
    for t in meta["toc_lama"]:
        e = re.sub(r"[.\u2026]{2,}.*$", "", t)
        e = re.sub(r"\t.*$", "", e)
        e = norm(re.sub(r"\s+[\divxlc]+\s*$", "", e))
        if e and e.lower() not in heads and not re.match(r"^(halaman\s+judul|halaman\s+sampul|sampul)\b", e, re.I) \
                and not re.match(r"^daftar\s+(isi|tabel|gambar|lampiran)", e, re.I) \
                and not re.match(r"^(tabel|gambar|lampiran)\s+\d", e, re.I):
            no_body.append(e)
    hasil = {
        "file": out, "status_nomor_halaman": status, "jumlah_halaman": res["pages"] if res else None,
        "sampul_satu_halaman": res["cover_ok"] if res else None, "judul_tidak_terdeteksi": res["missing"] if res else [],
        "margin_cm": C["margin"], "posisi_nomor": C["pos_v"] + "_" + C["pos_h"],
        "audit_isi": aud, "koreksi_diterapkan": applied, "koreksi_tidak_ditemukan": missed,
        "entri_daftar_isi_lama_tanpa_halaman": no_body, "halaman_judul_di_daftar_isi": bool(C.get("_halaman_judul")),
        "jumlah": {"bab_dll": D.count.get("h1", 0), "subbab": D.count.get("h2", 0), "subsubbab": D.count.get("h3", 0),
                   "tabel": D.count.get("tabel", 0), "gambar": D.count.get("gambar", 0)},
        "daftar_tabel": len(D.toc["Tabel"]) if C["daftar_tabel_gambar"] else 0,
        "daftar_gambar": len(D.toc["Gambar"]) if C["daftar_tabel_gambar"] else 0,
        "judul_kosong_disisipkan": D.inserted, "pustaka_diurutkan_ulang": sorted_flag,
        "keterangan_dipindah": D.moved, "judul_dipisah_otomatis": split_auto, "peringatan": WARN,
    }
    json.dump(hasil, open("hasil.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if not aud["cocok"]:
        print("STATUS=JANGAN_KIRIM")
        print("SELISIH=" + json.dumps(aud["selisih"][:8], ensure_ascii=False))
        print("Isi berbeda dari naskah asli. File tidak disimpan. Periksa koreksi di config lalu jalankan ulang.")
        return
    if C.get("_ditolak"):
        print("STATUS=PERBAIKI_CONFIG")
        print("KOREKSI_DITOLAK=" + json.dumps(C["_ditolak"], ensure_ascii=False))
        print("Perubahan ini bukan salah ketik dan mengubah isi. Hapus dari config, lalu jalankan ulang.")
        return
    if missed:
        print("STATUS=PERBAIKI_CONFIG")
        print("KOREKSI_TIDAK_DITEMUKAN=" + json.dumps([k["cari"] for k in missed], ensure_ascii=False))
        print("Tulis 'cari' persis seperti di naskah, atau hapus koreksinya, lalu jalankan ulang.")
        return
    print("STATUS=SIAP_KIRIM")
    print("FILE=" + out)
    print("---BALASAN---")
    print(balasan(hasil, name, C, applied))
    print("---AKHIR---")

def balasan(h, name, C, applied):
    f = lambda v: ("%g" % v)
    m = C["margin"]
    hal = ("ya" if C["daftar_tabel_gambar"] else "tidak")
    L = ["\u2705 Makalah kamu sudah rapi: **%s**" % name, ""]
    L.append("Pengaturan: margin %s cm \u00b7 halaman tabel/gambar %s \u00b7 nomor halaman %s %s%s." % (
        "-".join(f(m[k]) for k in ("kiri", "atas", "kanan", "bawah")), hal, C["pos_h"].capitalize(), C["pos_v"].capitalize(),
        (" \u00b7 %s halaman" % h["jumlah_halaman"]) if h["jumlah_halaman"] else ""))
    L.append("Isi naskah sudah dicek kata per kata (%s kata) dan sama dengan aslinya, kecuali koreksi berikut." % h["audit_isi"]["jumlah_kata"])
    L.append("")
    if applied:
        L.append("**Salah ketik yang dibetulkan:**")
        L += ['- "%s" \u2192 "%s"' % (k["cari"], k["ganti"]) for k in applied]
    else:
        L.append("Tidak ada salah ketik yang perlu dibetulkan.")
    cek = []
    for e in h["entri_daftar_isi_lama_tanpa_halaman"]:
        cek.append('"%s" ada di daftar isi lama, tetapi halamannya tidak ada di naskah, jadi entrinya dihapus.' % e)
    for e in h["judul_kosong_disisipkan"]:
        cek.append("%s belum punya judul; penandanya perlu kamu isi." % e)
    if C["daftar_tabel_gambar"] and not h["jumlah"]["tabel"] and not h["jumlah"]["gambar"]:
        cek.append("Naskah tidak punya tabel atau gambar, jadi halaman daftarnya tidak dibuat.")
    if h["pustaka_diurutkan_ulang"]: cek.append("Daftar Pustaka diurutkan ulang sesuai abjad.")
    if h["keterangan_dipindah"]: cek.append("Judul tabel/gambar dipindah ke posisi yang benar.")
    if h["sampul_satu_halaman"] is False: cek.append("Sampul terlalu padat sehingga lebih dari satu halaman.")
    if h["status_nomor_halaman"] != "OK" or h["judul_tidak_terdeteksi"]:
        cek.append("Nomor Daftar Isi baru pasti setelah diperbarui di Word.")
    cek += h["peringatan"]
    if cek:
        L += ["", "**Perlu kamu cek:**"] + ["- " + c for c in cek]
    L += ["", "Saat file dibuka di Word, klik **Yes/Ya** jika ditawari memperbarui field agar Daftar Isi menyesuaikan."]
    return "\n".join(L)

def cari_docx():
    fs = []
    for root_, _, names in os.walk("/mnt/user-data/uploads"):
        fs += [os.path.join(root_, n) for n in names if n.lower().endswith(".docx")]
    if not fs:
        print("FILE_TIDAK_ADA: belum ada makalah .docx yang dikirim."); sys.exit(3)
    return max(fs, key=os.path.getmtime)

WORD_RX = re.compile(r"[^\W\d_]+(?:-[^\W\d_]+)*")
TIDAK_BAKU = re.compile(r"\b(disini|disana|dimana|kemana|kesini|kesana|darimana|dikarenakan|diatas|dibawah|didalam|diluar|kedalam|keluar\s+negeri)\b", re.I)

def edit1(a, b):
    """Jarak sunting 1 (termasuk dua huruf tertukar)."""
    la, lb = len(a), len(b)
    if abs(la - lb) > 1 or a == b: return False
    if la == lb:
        d = [i for i in range(la) if a[i] != b[i]]
        return len(d) == 1 or (len(d) == 2 and d[1] == d[0] + 1 and a[d[0]] == b[d[1]] and a[d[1]] == b[d[0]])
    if la > lb: a, b = b, a
    return any(b[:i] + b[i + 1:] == a for i in range(len(b)))

def stem_known(w, vocab):
    for p in ("", "di", "ter", "ber", "se", "ke", "me", "pe"):
        if not w.startswith(p): continue
        x = w[len(p):]
        for suf in ("", "nya", "kan", "lah", "an", "i", "annya", "kannya"):
            if suf and not x.endswith(suf): continue
            y = x[:-len(suf)] if suf else x
            if y != w and len(y) >= 3 and y in vocab: return True
    return False

def konteks_of(t, a, b, n=5):
    kiri = t[:a].split()[-n:]; kanan = t[b:].split()[:n]
    return "%s [%s] %s" % (" ".join(kiri), t[a:b], " ".join(kanan))

def cmd_periksa(path, penuh=False):
    parts = read_parts(path)
    rd, blocks = load_blocks(parts, {})
    auto_split(blocks)
    classify(blocks)
    lines = []
    for b in blocks:
        tx = b.asli if b.kind == "tbl" else text_only(b.segs).replace("\n", " / ").replace("\t", " ")
        if b.role == "buang" or not tx.strip(): continue
        lines.append("[%s|%s] %s" % (b.id, b.role or "tabel", tx))
    open("naskah.txt", "w", encoding="utf-8").write("\n".join(lines))
    print("FILE=" + path)
    print("TABEL=%d GAMBAR=%d" % (sum(b.kind == "tbl" and b.role != "sampul" for b in blocks), sum(b.role == "gambar" for b in blocks)))
    for w in WARN: print("PERINGATAN: " + w)
    if penuh:
        print("\n".join(lines)); return
    teks = [b for b in blocks if b.kind == "p" and b.role not in ("buang", "pustaka", "judul_daftar")]
    print("== SAMPUL ==")
    for b in teks:
        if b.role == "sampul" and b.text: print("%s| %s" % (b.id, text_only(b.segs).replace("\n", " / ").strip()))
    print("== JUDUL ==")
    for b in teks:
        if b.role in ("h1", "h2", "h3", "h4", "cap_tabel", "cap_gambar", "cap_lampiran"): print("%s|%s| %s" % (b.id, b.role, b.text[:90]))
    cek = [b for b in blocks if b.kind == "p" and b.note]
    if cek:
        print("== PERLU DILIHAT ==")
        for b in cek: print("%s|%s|%s| %s" % (b.id, b.role, b.note, b.text[:110]))
    # kosakata naskah
    isi = [b for b in teks if b.role != "sampul"]
    cnt = {}
    for b in isi:
        for w in WORD_RX.findall(text_only(b.segs)):
            cnt[w.lower()] = cnt.get(w.lower(), 0) + 1
    vocab = set(cnt)
    sering = [w for w, c in cnt.items() if c >= 2 and len(w) >= 3]
    kand, jarang, seen = [], [], set()
    for b in teks:
        t = text_only(b.segs).replace("\n", " ")
        for m in re.finditer(r"\b(\w{2,}) \1\b", t):
            kand.append("%s| kata ganda | %s" % (b.id, konteks_of(t, m.start(), m.end())))
        for m in re.finditer(r"\b(?:[A-Z]{2,}[a-z]+[A-Za-z]*|[a-z]+[A-Z]+[a-z]*)\b", t):
            kand.append("%s| ketikan acak? | %s" % (b.id, konteks_of(t, m.start(), m.end())))
        for m in TIDAK_BAKU.finditer(t):
            kand.append("%s| ejaan? | %s" % (b.id, konteks_of(t, m.start(), m.end())))
        for m in re.finditer(r"\s+[,.;:?!](?=\s|$)|[,;:](?=[A-Za-z])", t):
            kand.append("%s| spasi/tanda baca? | %s" % (b.id, konteks_of(t, m.start(), m.end())))
        if b.role == "badan" and t.strip() and not re.search(r"[.?!:;\"\u201d)]\s*$", t):
            kand.append("%s| tanpa titik akhir | ...%s" % (b.id, t.strip()[-45:]))
        if b.role == "sampul": continue
        for m in WORD_RX.finditer(t):
            w, lw = m.group(0), m.group(0).lower()
            if lw in seen or cnt.get(lw, 0) != 1 or len(lw) < 2: continue
            seen.add(lw)
            at_start = m.start() == 0 or t[:m.start()].rstrip()[-1:] in ".?!:\"("
            if w[0].isupper() and not at_start: continue  # nama diri tidak disentuh
            if stem_known(lw, vocab): continue
            mirip = next((x for x in sering if cnt[x] >= 3 and edit1(lw, x)), None) if len(lw) >= 6 else None
            if mirip:
                kand.append("%s| dekat '%s' | %s" % (b.id, mirip, konteks_of(t, m.start(), m.end())))
            else:
                jarang.append(w)
    print("== KANDIDAT SALAH KETIK (nilai satu per satu) ==")
    print("\n".join(kand[:60]) if kand else "(tidak ada)")
    print("== KATA YANG MUNCUL SEKALI (cari yang salah eja) ==")
    print(" ".join(jarang))
    kur = []
    for b in teks:
        for m in re.finditer(r"\(([A-Za-z]{4,}(?: [A-Za-z]+)?)\)", text_only(b.segs)):
            if m.group(1) not in kur: kur.append(m.group(1))
    if kur: print("== DALAM KURUNG (istilah asing?) ==\n" + ", ".join(kur))
    pus = [b for b in blocks if b.role == "pustaka" and not any(s.i for s in b.segs)]
    if pus:
        print("== PUSTAKA TANPA HURUF MIRING ==")
        for b in pus: print(b.text)

def cmd_konteks(word, path):
    parts = read_parts(path)
    rd, blocks = load_blocks(parts, {})
    auto_split(blocks); classify(blocks)
    n = 0
    for b in blocks:
        if b.kind != "p" or b.role == "buang": continue
        t = text_only(b.segs).replace("\n", " ")
        for m in re.finditer(re.escape(word), t):
            print("%s| %s" % (b.id, konteks_of(t, m.start(), m.end(), 10))); n += 1
            if n >= 10: return
    if not n: print("TIDAK_ADA")

def main():
    a = sys.argv[1:]
    doc = next((x for x in a if x.lower().endswith(".docx")), None)
    if a and a[0] == "periksa":
        cmd_periksa(doc or cari_docx(), "--penuh" in a); return
    if len(a) >= 2 and a[0] == "konteks":
        cmd_konteks(a[1], doc or cari_docx()); return
    if len(a) >= 2 and a[0] == "rapikan" and not a[1].lower().endswith(".docx"):
        od = a[a.index("--out") + 1] if "--out" in a else ("/mnt/user-data/outputs" if os.path.isdir("/mnt/user-data") else os.getcwd())
        cmd_rapikan(doc or cari_docx(), a[1], od); return
    if len(a) >= 2 and a[0] == "baca":
        nk = a[a.index("--naskah") + 1] if "--naskah" in a else "naskah.txt"
        cmd_baca(a[1], nk)
    elif len(a) >= 3 and a[0] == "rapikan":
        od = a[a.index("--out") + 1] if "--out" in a else ("/mnt/user-data/outputs" if os.path.isdir("/mnt/user-data") else os.getcwd())
        cmd_rapikan(a[1], a[2], od)
    else:
        print("Pemakaian: yzmakalah2.py baca MAKALAH.docx [--naskah F] | rapikan MAKALAH.docx config.json [--out D]")
        sys.exit(1)

if __name__ == "__main__":
    main()
