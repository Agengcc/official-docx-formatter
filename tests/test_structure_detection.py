from __future__ import annotations

from official_docx_engine.diagnostics import diagnose_snapshot
from official_docx_engine.models import DocumentSnapshot
from official_docx_engine.structure import analyze_structure, paragraph_role

from fixtures.text_cases import MIXED_RUN_PARAGRAPHS, MULTILINE_NOTICE, MISSING_RECIPIENT_REPORT, snapshot_from_lines


def test_detects_multiline_title_and_core_roles() -> None:
    snapshot = snapshot_from_lines(MULTILINE_NOTICE)
    structure = analyze_structure(snapshot)

    assert structure.title_indices == (0, 1)
    assert structure.recipient_index == 2
    assert 3 in structure.body_indices
    assert 4 in structure.body_indices
    assert structure.attachment_indices == (5,)
    assert structure.issuer_index == 6
    assert structure.date_index == 7
    assert structure.role_for(0) == "title"
    assert structure.role_for(5) == "attachment_note"


def test_dot_separated_date_sentence_is_body_not_title() -> None:
    text = "2026.04.20，项目完成现场检查。"

    assert paragraph_role(text, index=0, total=1) == "body"

    snapshot = snapshot_from_lines([text])
    structure = analyze_structure(snapshot)

    assert structure.title_indices == ()
    assert structure.body_indices == (0,)
    assert structure.date_index is None


def test_diagnostics_report_missing_recipient_and_mixed_run_styles() -> None:
    snapshot = snapshot_from_lines(MISSING_RECIPIENT_REPORT)
    structure = analyze_structure(snapshot)
    codes = {issue.code for issue in diagnose_snapshot(snapshot, structure)}

    assert "missing_recipient" in codes
    assert "missing_title" not in codes
    assert "missing_date" not in codes

    mixed_snapshot = DocumentSnapshot(path=None, paragraphs=MIXED_RUN_PARAGRAPHS, table_count=1)
    mixed_structure = analyze_structure(mixed_snapshot)
    mixed_codes = {issue.code for issue in diagnose_snapshot(mixed_snapshot, mixed_structure)}

    assert {"tables_present", "mixed_fonts", "mixed_font_sizes"} <= mixed_codes
