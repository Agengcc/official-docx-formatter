"""JSON reporting helpers for official DOCX formatting runs."""

from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass
from pathlib import Path
from typing import Any, Iterable

from .models import DiagnosticIssue, FormatOperation, StructureAnalysis


def _jsonable(value: Any) -> Any:
    if isinstance(value, Path):
        return str(value)
    if is_dataclass(value):
        return _jsonable(asdict(value))
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    return value


def structure_to_dict(structure: StructureAnalysis) -> dict[str, Any]:
    return {
        "title_indices": list(structure.title_indices),
        "recipient_index": structure.recipient_index,
        "body_indices": list(structure.body_indices),
        "attachment_indices": list(structure.attachment_indices),
        "issuer_index": structure.issuer_index,
        "date_index": structure.date_index,
        "findings": [
            {
                "paragraph_index": finding.paragraph_index,
                "role": finding.role,
                "confidence": finding.confidence,
                "reason": finding.reason,
            }
            for finding in structure.findings
        ],
    }


def write_report_json(
    report_path: str | Path,
    *,
    input_path: str | Path,
    output_path: str | Path,
    profile_id: str,
    doc_type: str,
    structure: StructureAnalysis,
    issues: Iterable[DiagnosticIssue],
    operations: Iterable[FormatOperation],
) -> Path:
    """Write the formatting report JSON and return its path."""

    path = Path(report_path)
    payload = {
        "input": str(input_path),
        "output": str(output_path),
        "profile_id": profile_id,
        "doc_type": doc_type,
        "structure": structure_to_dict(structure),
        "issues": _jsonable(tuple(issues)),
        "operations": _jsonable(tuple(operations)),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path
