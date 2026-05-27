#!/usr/bin/env python3
"""Extract likely in-text citations from DOCX and create an audit checklist."""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path


CITATION_RE = re.compile(
    r"(?P<cite>[（(][^（）()]{1,40}?(?:19|20)\d{2}[a-z]?[^（）()]{0,20}[）)])"
)


def read_docx_text(path: Path) -> str:
    try:
        from docx import Document
    except Exception as exc:  # pragma: no cover
        print(f"python-docx is unavailable: {exc}", file=sys.stderr)
        raise SystemExit(2)
    document = Document(str(path))
    return "\n".join(p.text for p in document.paragraphs)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a citation audit checklist from a DOCX file.")
    parser.add_argument("docx", type=Path)
    parser.add_argument("--out", type=Path, default=None)
    args = parser.parse_args()

    if not args.docx.exists():
        print(f"file not found: {args.docx}", file=sys.stderr)
        return 1

    text = read_docx_text(args.docx)
    citations = [m.group("cite") for m in CITATION_RE.finditer(text)]
    counts = Counter(citations)

    lines = [
        "# 引文核验清单",
        "",
        f"- 来源文档：`{args.docx}`",
        f"- 候选文内引用数量：{sum(counts.values())}",
        f"- 去重后数量：{len(counts)}",
        "",
        "## 待核验条目",
        "",
    ]
    if not counts:
        lines.append("未抽取到明显的作者-年份式文内引用。")
    else:
        for idx, (cite, count) in enumerate(counts.most_common(), start=1):
            lines.extend(
                [
                    f"### {idx}. {cite}",
                    "",
                    f"- 出现次数：{count}",
                    "- 正文句子：",
                    "- Zotero 条目：",
                    "- 文献原文依据：",
                    "- 支持程度：完全支持 / 部分支持 / 不支持 / 待确认",
                    "- 修订动作：保留 / 弱化 / 删除 / 改为本文观点 / 另找文献",
                    "",
                ]
            )

    output = "\n".join(lines)
    if args.out:
        args.out.write_text(output, encoding="utf-8")
    else:
        print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
