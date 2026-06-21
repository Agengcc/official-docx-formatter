"""Conservative structure detection for official document snapshots."""

from __future__ import annotations

import re
from typing import Iterable, Optional

from .models import DocumentSnapshot, ParagraphFinding, StructureAnalysis

CN_DATE_RE = re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日$")
DOT_DATE_RE = re.compile(r"^\d{4}[.．]\d{1,2}[.．]\d{1,2}$")
DOT_DATE_SENTENCE_RE = re.compile(r"^\d{4}[.．]\d{1,2}[.．]\d{1,2}[，,。；;：:].+")
HEADING_RE = re.compile(
    r"^([一二三四五六七八九十]+[、，,]|（[一二三四五六七八九十]+）|\([一二三四五六七八九十]+\)|\d+[.．、])"
)
ORG_HINT_RE = re.compile(r"(公司|集团|局|厅|部|委|办|处|科|院|中心|人民政府|办公室|党组|党委)$")
TITLE_KEYWORD_RE = re.compile(r"(通知|通报|报告|请示|批复|函|纪要|意见|决定|决议|公告|通告|方案|总结|汇报)")
SENTENCE_PUNCT_RE = re.compile(r"[。；;！？!?]")


def _clean(text: str) -> str:
    return re.sub(r"\s+", "", text.strip())


def _has_sentence_punctuation(text: str) -> bool:
    return bool(SENTENCE_PUNCT_RE.search(text))


def _is_recipient(text: str) -> bool:
    stripped = text.strip()
    return len(stripped) <= 50 and stripped.endswith(("：", ":")) and not _has_sentence_punctuation(stripped)


def _is_attachment_note(text: str) -> bool:
    return text.strip().startswith(("附件：", "附件:", "附件 "))


def _is_date(text: str) -> bool:
    stripped = _clean(text)
    return bool(CN_DATE_RE.match(stripped) or DOT_DATE_RE.match(stripped))


def _is_signature_candidate(text: str, index: int, total: int) -> bool:
    stripped = _clean(text)
    if not stripped or len(stripped) > 40:
        return False
    if index < max(total - 4, 0):
        return False
    if _is_date(stripped) or _has_sentence_punctuation(stripped) or stripped.endswith(("：", ":")):
        return False
    return bool(ORG_HINT_RE.search(stripped))


def _is_title_candidate(text: str, index: int, total: int) -> bool:
    stripped = _clean(text)
    if not stripped or index > 3 or len(stripped) > 60:
        return False
    if _is_recipient(stripped) or _is_date(stripped) or _is_attachment_note(stripped):
        return False
    if DOT_DATE_SENTENCE_RE.match(stripped):
        return False
    if _has_sentence_punctuation(stripped):
        return False
    if HEADING_RE.match(stripped):
        return False
    return bool(TITLE_KEYWORD_RE.search(stripped) or index == 0)


def paragraph_role(text: str, index: int, total: int) -> str:
    """Classify one paragraph with context-light official document roles."""

    stripped = text.strip()
    if not stripped:
        return "empty"
    compact = _clean(stripped)
    if _is_recipient(stripped):
        return "recipient"
    if _is_attachment_note(stripped):
        return "attachment_note"
    if _is_date(stripped):
        return "date"
    if DOT_DATE_SENTENCE_RE.match(compact):
        return "body"
    if _is_signature_candidate(stripped, index, total):
        return "issuer"
    if _is_title_candidate(stripped, index, total):
        return "title"
    return "body"


def _non_empty_items(snapshot: DocumentSnapshot) -> list[tuple[int, str]]:
    return [(paragraph.index, paragraph.text.strip()) for paragraph in snapshot.paragraphs if paragraph.text.strip()]


def _find_date(items: Iterable[tuple[int, str]]) -> Optional[int]:
    for index, text in reversed(list(items)):
        if _is_date(text):
            return index
    return None


def analyze_structure(snapshot: DocumentSnapshot) -> StructureAnalysis:
    """Detect the document's main official-document structure."""

    items = _non_empty_items(snapshot)
    total = len(snapshot.paragraphs)
    if not items:
        return StructureAnalysis(findings=())

    item_by_index = dict(items)
    findings_by_index: dict[int, ParagraphFinding] = {}
    date_index = _find_date(items)
    issuer_index: Optional[int] = None

    if date_index is not None:
        preceding = [index for index, _ in items if index < date_index]
        if preceding:
            candidate = preceding[-1]
            if paragraph_role(item_by_index[candidate], candidate, total) in {"issuer", "body"}:
                text = item_by_index[candidate]
                if _is_signature_candidate(text, candidate, total) or len(_clean(text)) <= 40:
                    issuer_index = candidate

    title_indices: list[int] = []
    for index, text in items[:4]:
        if index in {issuer_index, date_index}:
            break
        if paragraph_role(text, index, total) == "title":
            title_indices.append(index)
            continue
        if title_indices and not _is_recipient(text):
            compact = _clean(text)
            if len(compact) <= 40 and not _has_sentence_punctuation(compact) and not HEADING_RE.match(compact):
                title_indices.append(index)
                continue
        break

    recipient_index: Optional[int] = None
    for index, text in items:
        if index in title_indices:
            continue
        if index in {issuer_index, date_index}:
            break
        if _is_recipient(text):
            recipient_index = index
            break
        if title_indices and index > title_indices[-1] + 3:
            break

    attachment_indices: list[int] = []
    body_indices: list[int] = []
    for index, text in items:
        if index in title_indices:
            role = "title"
            reason = "top short official-document heading"
        elif index == recipient_index:
            role = "recipient"
            reason = "short paragraph ending with colon"
        elif index == issuer_index:
            role = "issuer"
            reason = "signature paragraph before date"
        elif index == date_index:
            role = "date"
            reason = "standalone date"
        elif _is_attachment_note(text):
            role = "attachment_note"
            attachment_indices.append(index)
            reason = "attachment note prefix"
        else:
            role = "body"
            body_indices.append(index)
            reason = "default body paragraph"
        findings_by_index[index] = ParagraphFinding(index, role, text, reason=reason)

    return StructureAnalysis(
        findings=tuple(findings_by_index[index] for index, _ in items),
        title_indices=tuple(title_indices),
        recipient_index=recipient_index,
        body_indices=tuple(body_indices),
        attachment_indices=tuple(attachment_indices),
        issuer_index=issuer_index,
        date_index=date_index,
    )
