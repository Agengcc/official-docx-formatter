"""Diagnostic checks for trusted official document snapshots."""

from __future__ import annotations

from collections import Counter

from .models import DiagnosticIssue, DocumentSnapshot, StructureAnalysis


def _mixed_values(values: list[object]) -> bool:
    return len({value for value in values if value is not None}) > 1


def diagnose_snapshot(snapshot: DocumentSnapshot, structure: StructureAnalysis) -> tuple[DiagnosticIssue, ...]:
    issues: list[DiagnosticIssue] = []

    if not structure.title_indices:
        issues.append(DiagnosticIssue("missing_title", "error", "未识别到标题。"))
    if structure.recipient_index is None:
        issues.append(DiagnosticIssue("missing_recipient", "warning", "未识别到主送机关。"))
    if structure.date_index is None:
        issues.append(DiagnosticIssue("missing_date", "warning", "未识别到成文日期。"))
    if snapshot.table_count:
        issues.append(
            DiagnosticIssue(
                "tables_present",
                "info",
                "文档包含表格，当前可信引擎只对段落生成格式计划。",
                details={"table_count": snapshot.table_count},
            )
        )

    for paragraph in snapshot.paragraphs:
        if not paragraph.text.strip():
            continue
        fonts = [font for font in paragraph.font_names if font]
        sizes = [size for size in paragraph.font_sizes_pt if size]
        if _mixed_values(fonts):
            counts = Counter(fonts)
            issues.append(
                DiagnosticIssue(
                    "mixed_fonts",
                    "warning",
                    "同一段落存在多种字体。",
                    paragraph_index=paragraph.index,
                    details={"fonts": dict(counts)},
                )
            )
        if _mixed_values(sizes):
            counts = Counter(sizes)
            issues.append(
                DiagnosticIssue(
                    "mixed_font_sizes",
                    "warning",
                    "同一段落存在多种字号。",
                    paragraph_index=paragraph.index,
                    details={"font_sizes_pt": dict(counts)},
                )
            )

    return tuple(issues)
