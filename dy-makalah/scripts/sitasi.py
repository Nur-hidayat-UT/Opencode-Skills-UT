#!/usr/bin/env python3
"""Buat footnote nyata di DOCX dari placeholder [[FN:nomor]].

Pemakaian:
    python3 sitasi.py footnote "makalah.docx" catatan.json --out "folder"

Format catatan.json:
    {"catatan": {"1": "Teks catatan kaki 1", "2": "Teks catatan kaki 2"}}

Placeholder [[FN:1]] di dokumen akan diganti menjadi nomor footnote superscript,
dan teks catatan dimasukkan ke word/footnotes.xml. File hasil diberi akhiran
"_Sitasi.docx" agar file asli tidak berubah.
"""

import argparse
import json
import os
import re
import shutil
import sys
import zipfile
from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W, "r": R, "ct": CT, "rel": REL}

FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)
FOOTNOTES_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
)
PLACEHOLDER = re.compile(r"\[\[FN:([^\]]+)\]\]")


def qn(ns, tag):
    return f"{{{ns}}}{tag}"


def read_zip(path):
    with zipfile.ZipFile(path, "r") as z:
        return {name: z.read(name) for name in z.namelist()}


def write_zip(path, entries):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in entries.items():
            z.writestr(name, data)
    os.replace(tmp, path)


def ensure_content_type(entries):
    name = "[Content_Types].xml"
    if name not in entries:
        return
    root = etree.fromstring(entries[name])
    for ov in root.findall(qn(CT, "Override")):
        if ov.get("PartName") == "/word/footnotes.xml":
            return
    ov = etree.SubElement(root, qn(CT, "Override"))
    ov.set("PartName", "/word/footnotes.xml")
    ov.set("ContentType", FOOTNOTES_CT)
    entries[name] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def ensure_relationship(entries):
    name = "word/_rels/document.xml.rels"
    if name not in entries:
        root = etree.Element(qn(REL, "Relationships"), nsmap={None: REL})
    else:
        root = etree.fromstring(entries[name])
    for rel in root.findall(qn(REL, "Relationship")):
        if rel.get("Type") == FOOTNOTES_REL:
            return
    used = {rel.get("Id", "") for rel in root.findall(qn(REL, "Relationship"))}
    n = 1
    while f"rId{n}" in used:
        n += 1
    rel = etree.SubElement(root, qn(REL, "Relationship"))
    rel.set("Id", f"rId{n}")
    rel.set("Type", FOOTNOTES_REL)
    rel.set("Target", "footnotes.xml")
    entries[name] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def ensure_styles(entries):
    name = "word/styles.xml"
    if name not in entries:
        return
    root = etree.fromstring(entries[name])
    ids = {
        st.get(qn(W, "styleId"))
        for st in root.findall(qn(W, "style"))
    }
    changed = False
    if "FootnoteText" not in ids:
        st = etree.SubElement(root, qn(W, "style"))
        st.set(qn(W, "type"), "paragraph")
        st.set(qn(W, "styleId"), "FootnoteText")
        nm = etree.SubElement(st, qn(W, "name"))
        nm.set(qn(W, "val"), "footnote text")
        based = etree.SubElement(st, qn(W, "basedOn"))
        based.set(qn(W, "val"), "Normal")
        ppr = etree.SubElement(st, qn(W, "pPr"))
        sp = etree.SubElement(ppr, qn(W, "spacing"))
        sp.set(qn(W, "after"), "0")
        sp.set(qn(W, "line"), "240")
        sp.set(qn(W, "lineRule"), "auto")
        ind = etree.SubElement(ppr, qn(W, "ind"))
        ind.set(qn(W, "firstLine"), "0")
        jc = etree.SubElement(ppr, qn(W, "jc"))
        jc.set(qn(W, "val"), "both")
        rpr = etree.SubElement(st, qn(W, "rPr"))
        sz = etree.SubElement(rpr, qn(W, "sz"))
        sz.set(qn(W, "val"), "20")
        szcs = etree.SubElement(rpr, qn(W, "szCs"))
        szcs.set(qn(W, "val"), "20")
        changed = True
    if "FootnoteReference" not in ids:
        st = etree.SubElement(root, qn(W, "style"))
        st.set(qn(W, "type"), "character")
        st.set(qn(W, "styleId"), "FootnoteReference")
        nm = etree.SubElement(st, qn(W, "name"))
        nm.set(qn(W, "val"), "footnote reference")
        rpr = etree.SubElement(st, qn(W, "rPr"))
        va = etree.SubElement(rpr, qn(W, "vertAlign"))
        va.set(qn(W, "val"), "superscript")
        changed = True
    if changed:
        entries[name] = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )


def ensure_settings(entries):
    name = "word/settings.xml"
    if name not in entries:
        return
    root = etree.fromstring(entries[name])
    if root.find(qn(W, "footnotePr")) is not None:
        return
    fp = etree.Element(qn(W, "footnotePr"))
    for fid in ("-1", "0"):
        fn = etree.SubElement(fp, qn(W, "footnote"))
        fn.set(qn(W, "id"), fid)
    root.insert(0, fp)
    entries[name] = etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def load_footnotes(entries):
    name = "word/footnotes.xml"
    if name in entries:
        root = etree.fromstring(entries[name])
    else:
        root = etree.Element(qn(W, "footnotes"), nsmap={"w": W, "r": R})
        sep = etree.SubElement(root, qn(W, "footnote"))
        sep.set(qn(W, "type"), "separator")
        sep.set(qn(W, "id"), "-1")
        p = etree.SubElement(sep, qn(W, "p"))
        ppr = etree.SubElement(p, qn(W, "pPr"))
        sp = etree.SubElement(ppr, qn(W, "spacing"))
        sp.set(qn(W, "after"), "0")
        sp.set(qn(W, "line"), "240")
        sp.set(qn(W, "lineRule"), "auto")
        r = etree.SubElement(p, qn(W, "r"))
        etree.SubElement(r, qn(W, "separator"))
        cont = etree.SubElement(root, qn(W, "footnote"))
        cont.set(qn(W, "type"), "continuationSeparator")
        cont.set(qn(W, "id"), "0")
        p = etree.SubElement(cont, qn(W, "p"))
        ppr = etree.SubElement(p, qn(W, "pPr"))
        sp = etree.SubElement(ppr, qn(W, "spacing"))
        sp.set(qn(W, "after"), "0")
        sp.set(qn(W, "line"), "240")
        sp.set(qn(W, "lineRule"), "auto")
        r = etree.SubElement(p, qn(W, "r"))
        etree.SubElement(r, qn(W, "continuationSeparator"))
    return root


def next_note_id(root):
    ids = []
    for fn in root.findall(qn(W, "footnote")):
        try:
            ids.append(int(fn.get(qn(W, "id"))))
        except (TypeError, ValueError):
            pass
    return max([0] + ids) + 1


def make_ref_run(note_id):
    run = etree.Element(qn(W, "r"))
    rpr = etree.SubElement(run, qn(W, "rPr"))
    style = etree.SubElement(rpr, qn(W, "rStyle"))
    style.set(qn(W, "val"), "FootnoteReference")
    ref = etree.SubElement(run, qn(W, "footnoteReference"))
    ref.set(qn(W, "id"), str(note_id))
    return run


def add_footnote(root, note_id, text):
    fn = etree.SubElement(root, qn(W, "footnote"))
    fn.set(qn(W, "id"), str(note_id))
    p = etree.SubElement(fn, qn(W, "p"))
    ppr = etree.SubElement(p, qn(W, "pPr"))
    style = etree.SubElement(ppr, qn(W, "pStyle"))
    style.set(qn(W, "val"), "FootnoteText")
    sp = etree.SubElement(ppr, qn(W, "spacing"))
    sp.set(qn(W, "after"), "0")
    sp.set(qn(W, "line"), "240")
    sp.set(qn(W, "lineRule"), "auto")
    r1 = etree.SubElement(p, qn(W, "r"))
    rpr = etree.SubElement(r1, qn(W, "rPr"))
    rs = etree.SubElement(rpr, qn(W, "rStyle"))
    rs.set(qn(W, "val"), "FootnoteReference")
    etree.SubElement(r1, qn(W, "footnoteRef"))
    r2 = etree.SubElement(p, qn(W, "r"))
    t = etree.SubElement(r2, qn(W, "t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = " " + text


def clone_run(run, text):
    new = etree.Element(qn(W, "r"))
    rpr = run.find(qn(W, "rPr"))
    if rpr is not None:
        new.append(etree.fromstring(etree.tostring(rpr)))
    t = etree.SubElement(new, qn(W, "t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return new


def replace_placeholders(root, notes):
    inserted = []
    for run in list(root.iter(qn(W, "r"))):
        texts = run.findall(qn(W, "t"))
        if not texts:
            continue
        full = "".join(t.text or "" for t in texts)
        matches = list(PLACEHOLDER.finditer(full))
        if not matches:
            continue
        parent = run.getparent()
        idx = parent.index(run)
        pos = 0
        pieces = []
        for m in matches:
            key = m.group(1)
            if key not in notes:
                raise ValueError(f"Catatan untuk placeholder {key} tidak ditemukan")
            before = full[pos:m.start()]
            if before:
                pieces.append(clone_run(run, before))
            note_id = notes[key]
            pieces.append(make_ref_run(note_id))
            inserted.append((key, note_id))
            pos = m.end()
        after = full[pos:]
        if after:
            pieces.append(clone_run(run, after))
        parent.remove(run)
        for offset, piece in enumerate(pieces):
            parent.insert(idx + offset, piece)
    return inserted


def main():
    ap = argparse.ArgumentParser(description="Buat footnote DOCX dari placeholder [[FN:n]]")
    sub = ap.add_subparsers(dest="cmd", required=True)
    fn = sub.add_parser("footnote", help="ganti placeholder dengan footnote nyata")
    fn.add_argument("docx")
    fn.add_argument("spec", help="JSON berisi {\"catatan\": {\"1\": \"...\"}}")
    fn.add_argument("--out", default=None, help="folder keluaran (default: folder asli)")
    args = ap.parse_args()

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    raw_notes = spec.get("catatan", spec)
    if not isinstance(raw_notes, dict):
        sys.exit("Format spec salah: 'catatan' harus objek nomor -> teks")

    entries = read_zip(args.docx)
    doc_root = etree.fromstring(entries["word/document.xml"])
    fn_root = load_footnotes(entries)

    notes = {}
    for key, text in raw_notes.items():
        note_id = next_note_id(fn_root)
        add_footnote(fn_root, note_id, str(text))
        notes[str(key)] = note_id

    inserted = replace_placeholders(doc_root, notes)
    if not inserted:
        sys.exit("Tidak ada placeholder [[FN:...]] yang ditemukan di dokumen")

    entries["word/document.xml"] = etree.tostring(
        doc_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    entries["word/footnotes.xml"] = etree.tostring(
        fn_root, xml_declaration=True, encoding="UTF-8", standalone=True
    )
    ensure_content_type(entries)
    ensure_relationship(entries)
    ensure_styles(entries)
    ensure_settings(entries)

    out_dir = args.out or os.path.dirname(os.path.abspath(args.docx))
    os.makedirs(out_dir, exist_ok=True)
    stem = os.path.splitext(os.path.basename(args.docx))[0]
    out_path = os.path.join(out_dir, stem + "_Sitasi.docx")
    write_zip(out_path, entries)

    print("OK", out_path)
    for key, note_id in inserted:
        print(f"  [[FN:{key}]] -> footnote id {note_id}")


if __name__ == "__main__":
    main()
