"""Build declarative formatting plans without mutating documents."""

from __future__ import annotations

from .models import DocumentSnapshot, FormatOperation, FormatPlan, StructureAnalysis


ROLE_STYLE = {
    "title": "title",
    "recipient": "body",
    "body": "body",
    "attachment_note": "body",
    "issuer": "body",
    "date": "body",
}


def build_format_plan(
    snapshot: DocumentSnapshot,
    structure: StructureAnalysis,
    profile_id: str,
    doc_type: str,
    normalize_text: bool,
    space_mode: str = "keep_en_boundary",
) -> FormatPlan:
    operations: list[FormatOperation] = [
        FormatOperation(
            kind="page_setup",
            target="document",
            params={"profile_id": profile_id, "paper": "A4"},
            reason="apply profile page settings",
        )
    ]

    role_by_index = {finding.paragraph_index: finding.role for finding in structure.findings}
    for paragraph in snapshot.paragraphs:
        if not paragraph.text.strip():
            continue
        role = role_by_index.get(paragraph.index, "body")
        operations.append(
            FormatOperation(
                kind="paragraph_style",
                target="paragraph",
                paragraph_index=paragraph.index,
                params={
                    "role": role,
                    "style": ROLE_STYLE.get(role, "body"),
                    "profile_id": profile_id,
                },
                reason=f"format {role} paragraph",
            )
        )

    if structure.issuer_index is not None or structure.date_index is not None:
        operations.append(
            FormatOperation(
                kind="signature_layout",
                target="signature",
                params={
                    "issuer_index": structure.issuer_index,
                    "date_index": structure.date_index,
                    "profile_id": profile_id,
                },
                reason="align issuer and date according to profile",
            )
        )

    if normalize_text:
        for paragraph in snapshot.paragraphs:
            if paragraph.text.strip():
                operations.append(
                    FormatOperation(
                        kind="text_normalization",
                        target="paragraph",
                        paragraph_index=paragraph.index,
                        params={"space_mode": space_mode},
                        reason="normalize punctuation and spacing conservatively",
                    )
                )

    return FormatPlan(profile_id=profile_id, doc_type=doc_type, operations=tuple(operations))
