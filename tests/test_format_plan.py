from __future__ import annotations

from official_docx_engine.format_plan import build_format_plan
from official_docx_engine.structure import analyze_structure

from fixtures.text_cases import MULTILINE_NOTICE, snapshot_from_lines


def test_build_format_plan_includes_core_operations() -> None:
    snapshot = snapshot_from_lines(MULTILINE_NOTICE)
    structure = analyze_structure(snapshot)

    plan = build_format_plan(
        snapshot,
        structure,
        profile_id="standard-party-government",
        doc_type="通知",
        normalize_text=True,
        space_mode="keep_en_boundary",
    )

    kinds = [operation.kind for operation in plan.operations]
    paragraph_ops = plan.operations_of_kind("paragraph_style")

    assert kinds[0] == "page_setup"
    assert "signature_layout" in kinds
    assert "text_normalization" in kinds
    assert len(paragraph_ops) == len([line for line in MULTILINE_NOTICE if line.strip()])
    assert paragraph_ops[0].params["role"] == "title"
    assert paragraph_ops[2].params["role"] == "recipient"
    assert plan.profile_id == "standard-party-government"
    assert plan.doc_type == "通知"


def test_build_format_plan_respects_normalize_text_flag() -> None:
    snapshot = snapshot_from_lines(MULTILINE_NOTICE)
    structure = analyze_structure(snapshot)

    plan = build_format_plan(
        snapshot,
        structure,
        profile_id="standard-party-government",
        doc_type="通知",
        normalize_text=False,
        space_mode="keep_en_boundary",
    )

    assert plan.operations_of_kind("page_setup")
    assert plan.operations_of_kind("paragraph_style")
    assert not plan.operations_of_kind("text_normalization")


def test_build_format_plan_uses_requested_space_mode() -> None:
    snapshot = snapshot_from_lines(MULTILINE_NOTICE)
    structure = analyze_structure(snapshot)

    plan = build_format_plan(
        snapshot,
        structure,
        profile_id="standard-party-government",
        doc_type="通知",
        normalize_text=True,
        space_mode="remove_all",
    )

    text_ops = plan.operations_of_kind("text_normalization")
    assert text_ops
    assert {operation.params["space_mode"] for operation in text_ops} == {"remove_all"}
