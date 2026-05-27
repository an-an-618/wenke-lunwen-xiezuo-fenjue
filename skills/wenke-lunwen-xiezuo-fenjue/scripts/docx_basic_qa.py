#!/usr/bin/env python3
"""Basic QA for humanities-paper DOCX files."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


def load_docx():
    try:
        from docx import Document
    except Exception as exc:  # pragma: no cover
        print(json.dumps({"error": f"python-docx is unavailable: {exc}"}, ensure_ascii=False))
        sys.exit(2)
    return Document


def iter_header_texts(document):
    for section in document.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            text = "\n".join(p.text for p in header.paragraphs).strip()
            if text:
                yield text


def count_blue_runs(document):
    count = 0
    samples = []
    for para in document.paragraphs:
        for run in para.runs:
            color = run.font.color.rgb
            if color is None:
                continue
            rgb = str(color).upper()
            if rgb.endswith("FF") or rgb in {"0000FF", "0563C1", "1F4E79", "2F5496"}:
                count += 1
                if len(samples) < 5 and run.text.strip():
                    samples.append(run.text.strip()[:40])
    return count, samples


def main() -> int:
    parser = argparse.ArgumentParser(description="Run basic structural QA on a DOCX file.")
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()

    if not args.docx.exists():
        print(json.dumps({"error": f"file not found: {args.docx}"}, ensure_ascii=False))
        return 1

    Document = load_docx()
    document = Document(str(args.docx))
    text = "\n".join(p.text for p in document.paragraphs)
    headers = list(iter_header_texts(document))
    blue_runs, blue_samples = count_blue_runs(document)
    result = {
        "file": str(args.docx),
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "headers_nonempty": len(headers),
        "header_samples": headers[:5],
        "chinese_chars": len(re.findall(r"[\u4e00-\u9fff]", text)),
        "nonspace_chars": len(re.sub(r"\s+", "", text)),
        "blue_runs": blue_runs,
        "blue_samples": blue_samples,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
