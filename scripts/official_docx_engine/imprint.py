"""Conservative formatting for existing official-document imprint lines."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt


IMPRINT_RE = re.compile(r"\d{4}年\d{1,2}月\d{1,2}日印发$")


@dataclass(frozen=True)
class ImprintFormatResult:
    imprint_count: int
    action: str
    paragraph_indices: tuple[int, ...] = ()


def format_existing_imprint(document, profile: dict[str, Any]) -> ImprintFormatResult:
    """Format existing imprint paragraphs near the end of the document."""

    indexed_paragraphs = list(enumerate(document.paragraphs))
    candidates = [
        (index, paragraph)
        for index, paragraph in indexed_paragraphs[-8:]
        if _is_imprint_text(paragraph.text)
    ]
    if not candidates:
        return ImprintFormatResult(imprint_count=0, action="no_imprint")

    for _, paragraph in candidates:
        _format_imprint_paragraph(paragraph, profile)

    return ImprintFormatResult(
        imprint_count=len(candidates),
        action="formatted",
        paragraph_indices=tuple(index for index, _ in candidates),
    )


def _is_imprint_text(text: str) -> bool:
    return bool(IMPRINT_RE.search(" ".join(text.strip().split())))


def _format_imprint_paragraph(paragraph, profile: dict[str, Any]) -> None:
    body_font = _preferred_font(profile, "body")
    body_size = _font_size(profile, "body", 16)
    layout = profile.get("layout", {})
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    paragraph.paragraph_format.first_line_indent = Pt(0)
    paragraph.paragraph_format.left_indent = Pt(0)
    paragraph.paragraph_format.right_indent = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(float(layout.get("line_spacing_pt", 28)))
    paragraph.paragraph_format.space_before = Pt(float(layout.get("space_before_pt", 0)))
    paragraph.paragraph_format.space_after = Pt(float(layout.get("space_after_pt", 0)))
    if not paragraph.runs:
        paragraph.add_run("")
    for run in paragraph.runs:
        run.font.name = body_font
        run.font.size = Pt(body_size)
        run.font.bold = False
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), body_font)


def _preferred_font(profile: dict[str, Any], key: str) -> str:
    fonts = profile.get("fonts", {}).get(key, {})
    fallbacks = fonts.get("fallbacks") or []
    return fallbacks[0] if fallbacks else "仿宋"


def _font_size(profile: dict[str, Any], key: str, default: float = 16) -> float:
    return float(profile.get("fonts", {}).get(key, {}).get("size_pt", default))
