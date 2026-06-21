#!/usr/bin/env python3
"""Diagnose a DOCX using the trusted official DOCX engine."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from official_docx_engine.diagnostics import diagnose_snapshot
from official_docx_engine.docx_reader import read_docx_snapshot
from official_docx_engine.reporting import structure_to_dict
from official_docx_engine.structure import analyze_structure


def _issue_to_dict(issue: Any) -> dict[str, Any]:
    return {
        "code": issue.code,
        "severity": issue.severity,
        "message": issue.message,
        "paragraph_index": issue.paragraph_index,
        "details": issue.details,
    }


def diagnose_docx(path: str | Path) -> dict[str, Any]:
    input_path = Path(path)
    snapshot = read_docx_snapshot(input_path)
    structure = analyze_structure(snapshot)
    issues = diagnose_snapshot(snapshot, structure)
    non_empty_by_index = {paragraph.index: paragraph.text.strip() for paragraph in snapshot.non_empty_paragraphs}

    title = "\n".join(
        non_empty_by_index[index]
        for index in structure.title_indices
        if index in non_empty_by_index
    )
    recipient = ""
    if structure.recipient_index is not None:
        recipient = non_empty_by_index.get(structure.recipient_index, "")

    return {
        "input": str(input_path),
        "title": title,
        "recipient": recipient,
        "body_paragraphs": len(structure.body_indices),
        "issuer_index": structure.issuer_index,
        "date_index": structure.date_index,
        "structure": structure_to_dict(structure),
        "issues": [_issue_to_dict(issue) for issue in issues],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Diagnose Chinese official DOCX structure")
    parser.add_argument("input", help="Input .docx path")
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    args = parser.parse_args()

    result = diagnose_docx(args.input)
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"input={result['input']}")
        print(f"title={result['title']}")
        print(f"recipient={result['recipient']}")
        print(f"body_paragraphs={result['body_paragraphs']}")
        print(f"issues={len(result['issues'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
