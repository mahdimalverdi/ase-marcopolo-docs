#!/usr/bin/env python3
from __future__ import annotations

import argparse
import zipfile
import xml.etree.ElementTree as ET


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", help="Path to .docx")
    args = parser.parse_args()

    with zipfile.ZipFile(args.docx) as z:
        names = set(z.namelist())

        rels_path = "word/_rels/document.xml.rels"
        doc_path = "word/document.xml"
        if rels_path not in names:
            raise SystemExit(f"Missing: {rels_path}")
        if doc_path not in names:
            raise SystemExit(f"Missing: {doc_path}")

        rels = ET.fromstring(z.read(rels_path))
        img_rels: dict[str, str] = {}
        for rel in rels.findall(f"{{{REL_NS}}}Relationship"):
            if rel.attrib.get("Type", "").endswith("/image"):
                img_rels[rel.attrib["Id"]] = rel.attrib.get("Target", "")

        doc = ET.fromstring(z.read(doc_path))
        embeds: list[str] = []
        for el in doc.iter():
            rid = el.attrib.get(f"{{{R_NS}}}embed")
            if rid:
                embeds.append(rid)

        missing_rel = [rid for rid in embeds if rid not in img_rels]
        missing_media = []
        for rid in embeds:
            tgt = img_rels.get(rid)
            if not tgt:
                continue
            part = f"word/{tgt}"
            if part not in names:
                missing_media.append((rid, part))

        print(f"docx: {args.docx}")
        print(f"embedded refs (r:embed): {len(embeds)}")
        print(f"image relationships: {len(img_rels)}")
        print(f"missing relationship for embed: {len(missing_rel)}")
        print(f"missing media parts: {len(missing_media)}")
        if missing_rel:
            print("missing rel ids:", ", ".join(missing_rel[:20]))
        if missing_media:
            for rid, part in missing_media[:20]:
                print(f"missing media: {rid} -> {part}")

        # Non-zero exit if broken
        if missing_rel or missing_media:
            return 2
        return 0


if __name__ == "__main__":
    raise SystemExit(main())

