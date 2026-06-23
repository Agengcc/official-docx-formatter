#!/usr/bin/env python3
"""Classify Chinese official document types from DOCX or text."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any, Dict, List

try:
    from docx import Document
except ImportError:  # pragma: no cover - text mode still works
    Document = None

SKILL_DIR = Path(__file__).resolve().parents[1]
DOCUMENT_TYPES_FILE = SKILL_DIR / "references" / "document_types.json"
GENERIC_FORMAL_TEXT = "通用正式文本"


def load_document_types() -> Dict[str, Any]:
    return json.loads(DOCUMENT_TYPES_FILE.read_text(encoding="utf-8"))


def read_text(path: Path) -> List[str]:
    if path.suffix.lower() == ".docx":
        if Document is None:
            raise SystemExit("python-docx is required to read .docx files")
        doc = Document(str(path))
        return [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text)


def contains_any(text: str, values: List[str]) -> bool:
    return any(value and value in text for value in values)


def has_type_signal(info: Dict[str, Any], title: str, body_text: str) -> Dict[str, bool]:
    normalized_title = normalize(title)
    normalized_body = normalize(body_text)
    title_values = info.get("title_keywords", []) + info.get("aliases", [])
    body_values = info.get("content_signals", []) + info.get("ending_signals", [])
    return {
        "title": contains_any(normalized_title, title_values),
        "body": contains_any(normalized_body, body_values),
    }


def score_type(doc_type: str, info: Dict[str, Any], title: str, body_text: str) -> Dict[str, Any]:
    score = 0
    evidence: list[str] = []
    normalized_title = normalize(title)
    normalized_body = normalize(body_text)

    for keyword in info.get("title_keywords", []):
        if keyword and keyword in normalized_title:
            score += 5
            evidence.append(f"标题包含“{keyword}”")

    for alias in info.get("aliases", []):
        if alias and alias in normalized_title and alias not in info.get("title_keywords", []):
            score += 4
            evidence.append(f"标题包含别名“{alias}”")

    for signal in info.get("content_signals", []):
        if signal and signal in normalized_body:
            score += 2
            evidence.append(f"正文出现“{signal}”")

    for ending in info.get("ending_signals", []):
        if ending and ending in normalized_body:
            score += 3
            evidence.append(f"出现结尾语“{ending}”")

    # Common confusion guard: a document may mention another type in the body.
    # Title evidence should dominate; body-only evidence stays lower confidence.
    confidence = min(0.98, score / 10) if score else 0.0
    return {
        "doc_type": doc_type,
        "score": score,
        "confidence": round(confidence, 2),
        "evidence": evidence,
    }


def has_report_request_conflict(
    document_types: Dict[str, Any],
    title: str,
    body_text: str,
    candidates: List[Dict[str, Any]],
) -> bool:
    report_info = document_types.get("报告")
    request_info = document_types.get("请示")
    if not report_info or not request_info:
        return False

    report_signal = has_type_signal(report_info, title, body_text)
    request_signal = has_type_signal(request_info, title, body_text)
    if not (report_signal["title"] or report_signal["body"]):
        return False
    if not (request_signal["title"] or request_signal["body"]):
        return False

    scores = {item["doc_type"]: item["score"] for item in candidates}
    report_score = scores.get("报告", 0)
    request_score = scores.get("请示", 0)
    cross_type_signal = (report_signal["title"] and request_signal["body"]) or (
        request_signal["title"] and report_signal["body"]
    )
    body_only_close = (
        report_signal["body"]
        and request_signal["body"]
        and not report_signal["title"]
        and not request_signal["title"]
        and abs(report_score - request_score) <= 3
    )
    return cross_type_signal or body_only_close


def should_ask_user(
    top: Dict[str, Any],
    second: Dict[str, Any],
    document_types: Dict[str, Any],
    title: str,
    body_text: str,
    candidates: List[Dict[str, Any]],
) -> bool:
    if top["score"] < 5:
        return True
    if top["score"] - second["score"] <= 1 and second["score"] > 0:
        return True
    return has_report_request_conflict(document_types, title, body_text, candidates)


def classify_lines(lines: List[str]) -> Dict[str, Any]:
    document_types = load_document_types()
    title = lines[0] if lines else ""
    body_text = "\n".join(lines[1:] if len(lines) > 1 else lines)
    candidates = [
        score_type(doc_type, info, title, body_text)
        for doc_type, info in document_types.items()
    ]
    candidates.sort(key=lambda item: item["score"], reverse=True)
    top = candidates[0] if candidates else {"score": 0, "confidence": 0}
    second = candidates[1] if len(candidates) > 1 else {"score": 0}

    ask_user = should_ask_user(top, second, document_types, title, body_text, candidates)
    return {
        "title": title,
        "top": top,
        "candidates": [item for item in candidates[:5] if item["score"] > 0],
        "ask_user": ask_user,
        "question": build_question(candidates) if ask_user else "",
    }


def build_question(candidates: List[Dict[str, Any]]) -> str:
    positive = [item["doc_type"] for item in candidates if item["score"] > 0][:3]
    if len(positive) >= 2:
        return (
            f"我看这篇更像是【{positive[0]}】或【{positive[1]}】。你希望怎么处理："
            f"{positive[0]} / {positive[1]} / {GENERIC_FORMAL_TEXT} / 其他？"
        )
    if len(positive) == 1:
        return (
            f"我初步判断像【{positive[0]}】，但证据不够强。你希望怎么处理："
            f"{positive[0]} / {GENERIC_FORMAL_TEXT} / 其他？"
        )
    return f"我还不能可靠判断文种。你希望怎么处理：{GENERIC_FORMAL_TEXT} / 请示 / 报告 / 通知 / 函 / 纪要 / 其他？"


def main() -> int:
    parser = argparse.ArgumentParser(description="Classify Chinese official document type")
    parser.add_argument("input", help="Input .docx or UTF-8 text file")
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    args = parser.parse_args()

    result = classify_lines(read_text(Path(args.input)))
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        top = result["top"]
        print(f"top={top.get('doc_type', '未知')} confidence={top.get('confidence', 0)} score={top.get('score', 0)}")
        for item in result["candidates"]:
            evidence = "；".join(item["evidence"]) or "无"
            print(f"- {item['doc_type']}: score={item['score']} evidence={evidence}")
        if result["ask_user"]:
            print(result["question"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
