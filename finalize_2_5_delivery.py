#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Finalize section 2.5 delivery and write a non-circular verification manifest."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

OUT_DIR = Path("generated")
DOCX = OUT_DIR / "section_2_5_korean_translation.docx"
VERIFY = OUT_DIR / "section_2_5_verification.txt"
DELIVERY = OUT_DIR / "section_2_5_korean_translation.zip"
MANIFEST = OUT_DIR / "section_2_5_delivery_manifest.txt"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main() -> None:
    if not DOCX.is_file() or DOCX.stat().st_size <= 1000:
        raise RuntimeError("DOCX is missing or unexpectedly small")

    with zipfile.ZipFile(DOCX, "r") as zf:
        if zf.testzip() is not None:
            raise RuntimeError("DOCX package failed ZIP integrity test")
        required = {
            "[Content_Types].xml",
            "_rels/.rels",
            "word/document.xml",
            "word/styles.xml",
            "word/settings.xml",
            "word/fontTable.xml",
            "word/footnotes.xml",
            "word/_rels/document.xml.rels",
        }
        missing = required - set(zf.namelist())
        if missing:
            raise RuntimeError(f"Missing DOCX members: {sorted(missing)}")
        document_xml = zf.read("word/document.xml")
        footnotes_xml = zf.read("word/footnotes.xml")
        settings_xml = zf.read("word/settings.xml")

    document = ET.fromstring(document_xml)
    footnotes = ET.fromstring(footnotes_xml)
    ET.fromstring(settings_xml)
    ns = {"w": W_NS}

    paragraphs = document.findall(".//w:body/w:p", ns)
    refs = [int(x.attrib[f"{{{W_NS}}}id"]) for x in document.findall(".//w:footnoteReference", ns)]
    note_ids = [
        int(x.attrib[f"{{{W_NS}}}id"])
        for x in footnotes.findall("w:footnote", ns)
        if int(x.attrib[f"{{{W_NS}}}id"]) > 0
    ]
    text = "".join(document.itertext())

    checks = {
        "DOCX_ZIP_INTEGRITY": True,
        "REQUIRED_OOXML_PARTS": True,
        "DOCUMENT_XML_PARSE": True,
        "FOOTNOTES_XML_PARSE": True,
        "SETTINGS_XML_PARSE": True,
        "PARAGRAPH_COUNT_7": len(paragraphs) == 7,
        "BODY_PARAGRAPH_COUNT_6": len(paragraphs) - 1 == 6,
        "FOOTNOTE_REFERENCE_COUNT_11": len(refs) == 11,
        "FOOTNOTE_REFERENCE_IDS_1_TO_11": refs == list(range(1, 12)),
        "FOOTNOTE_NODE_IDS_1_TO_11": note_ids == list(range(1, 12)),
        "DISPLAY_NUMBER_RANGE_61_TO_71": b'numStart' in settings_xml and b'val="61"' in settings_xml,
        "TITLE_PRESENT": "2.5 협정 형성의 가능성 조건" in text,
        "NO_UNRESOLVED_MARKERS": "[[" not in text and "]]" not in text,
        "A4_PAGE_SIZE": b'w:w="11906"' in document_xml and b'w:h="16838"' in document_xml,
        "BATANG_FONT": "바탕".encode("utf-8") in zf_read(DOCX, "word/styles.xml"),
        "BODY_10_5_PT": b'w:sz w:val="21"' in zf_read(DOCX, "word/styles.xml"),
        "FIRST_LINE_INDENT": b'w:firstLine="420"' in document_xml,
        "JUSTIFIED_BODY": b'w:jc w:val="both"' in document_xml,
        "LINE_SPACING_1_6": b'w:line="384"' in document_xml,
    }
    failed = [name for name, ok in checks.items() if not ok]
    if failed:
        raise RuntimeError("Verification failed: " + ", ".join(failed))

    verify_lines = [
        "RESULT=PASS",
        f"DOCX={DOCX.name}",
        f"DOCX_BYTES={DOCX.stat().st_size}",
        f"DOCX_SHA256={sha256(DOCX)}",
        "TITLE=2.5 협정 형성의 가능성 조건",
        "STRUCTURE=1_TITLE_PLUS_6_BODY_PARAGRAPHS",
        "FOOTNOTES=TRUE_WORD_FOOTNOTES_61_TO_71",
        "FORMAT=A4; Batang; body 10.5pt; justified; first-line indent; 1.6 line spacing",
    ]
    verify_lines.extend(f"CHECK_{name}=PASS" for name in checks)
    VERIFY.write_text("\n".join(verify_lines) + "\n", encoding="utf-8")

    with zipfile.ZipFile(DELIVERY, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.write(DOCX, arcname=DOCX.name)
        zf.write(VERIFY, arcname=VERIFY.name)
    with zipfile.ZipFile(DELIVERY, "r") as zf:
        if zf.testzip() is not None:
            raise RuntimeError("Delivery ZIP failed integrity test")
        if set(zf.namelist()) != {DOCX.name, VERIFY.name}:
            raise RuntimeError("Delivery ZIP contains an unexpected member set")

    manifest_lines = [
        "DELIVERY_RESULT=PASS",
        f"ZIP={DELIVERY.name}",
        f"ZIP_BYTES={DELIVERY.stat().st_size}",
        f"ZIP_SHA256={sha256(DELIVERY)}",
        f"DOCX={DOCX.name}",
        f"DOCX_BYTES={DOCX.stat().st_size}",
        f"DOCX_SHA256={sha256(DOCX)}",
        "ZIP_MEMBER_COUNT=2",
        f"ZIP_MEMBER_1={DOCX.name}",
        f"ZIP_MEMBER_2={VERIFY.name}",
        "ZIP_INTEGRITY=PASS",
    ]
    MANIFEST.write_text("\n".join(manifest_lines) + "\n", encoding="utf-8")
    print(VERIFY.read_text(encoding="utf-8"), end="")
    print(MANIFEST.read_text(encoding="utf-8"), end="")


def zf_read(path: Path, member: str) -> bytes:
    with zipfile.ZipFile(path, "r") as zf:
        return zf.read(member)


if __name__ == "__main__":
    main()
