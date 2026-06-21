"""Neutral data models for the trusted official DOCX engine."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional


@dataclass(frozen=True)
class RunSnapshot:
    text: str
    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None


@dataclass(frozen=True)
class ParagraphSnapshot:
    index: int
    text: str
    style_name: Optional[str] = None
    alignment: Optional[str] = None
    left_indent_pt: Optional[float] = None
    first_line_indent_pt: Optional[float] = None
    line_spacing: Optional[Any] = None
    line_spacing_pt: Optional[float] = None
    runs: tuple[RunSnapshot, ...] = field(default_factory=tuple)

    @property
    def font_names(self) -> tuple[str, ...]:
        return tuple(run.font_name for run in self.runs if run.font_name)

    @property
    def font_sizes_pt(self) -> tuple[float, ...]:
        return tuple(run.font_size_pt for run in self.runs if run.font_size_pt)


@dataclass(frozen=True)
class DocumentSnapshot:
    path: Optional[Path]
    paragraphs: tuple[ParagraphSnapshot, ...]
    table_count: int = 0

    @property
    def non_empty_paragraphs(self) -> tuple[ParagraphSnapshot, ...]:
        return tuple(p for p in self.paragraphs if p.text.strip())


@dataclass(frozen=True)
class ParagraphFinding:
    paragraph_index: int
    role: str
    text: str
    confidence: float = 1.0
    reason: str = ""


@dataclass(frozen=True)
class StructureAnalysis:
    findings: tuple[ParagraphFinding, ...]
    title_indices: tuple[int, ...] = field(default_factory=tuple)
    recipient_index: Optional[int] = None
    body_indices: tuple[int, ...] = field(default_factory=tuple)
    attachment_indices: tuple[int, ...] = field(default_factory=tuple)
    issuer_index: Optional[int] = None
    date_index: Optional[int] = None

    def role_for(self, paragraph_index: int) -> Optional[str]:
        for finding in self.findings:
            if finding.paragraph_index == paragraph_index:
                return finding.role
        return None


@dataclass(frozen=True)
class DiagnosticIssue:
    code: str
    severity: str
    message: str
    paragraph_index: Optional[int] = None
    details: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class FormatOperation:
    kind: str
    target: str
    params: dict[str, Any] = field(default_factory=dict)
    paragraph_index: Optional[int] = None
    reason: str = ""


@dataclass(frozen=True)
class FormatPlan:
    profile_id: str
    doc_type: str
    operations: tuple[FormatOperation, ...]

    def operations_of_kind(self, kind: str) -> tuple[FormatOperation, ...]:
        return tuple(operation for operation in self.operations if operation.kind == kind)
