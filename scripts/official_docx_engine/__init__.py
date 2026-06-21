"""Trusted engine foundation for Chinese official DOCX formatting."""

from .diagnostics import diagnose_snapshot
from .docx_reader import read_docx_snapshot
from .format_plan import build_format_plan
from .models import (
    DiagnosticIssue,
    DocumentSnapshot,
    FormatOperation,
    FormatPlan,
    ParagraphFinding,
    ParagraphSnapshot,
    RunSnapshot,
    StructureAnalysis,
)
from .structure import analyze_structure, paragraph_role

__all__ = [
    "DiagnosticIssue",
    "DocumentSnapshot",
    "FormatOperation",
    "FormatPlan",
    "ParagraphFinding",
    "ParagraphSnapshot",
    "RunSnapshot",
    "StructureAnalysis",
    "analyze_structure",
    "build_format_plan",
    "diagnose_snapshot",
    "paragraph_role",
    "read_docx_snapshot",
]
