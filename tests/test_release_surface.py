from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from docx import Document


PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEXT_SUFFIXES = {".md", ".py", ".json", ".yaml", ".yml", ".toml", ".txt"}
FORBIDDEN_PATTERNS = (
    "/Us" + "ers",
    "liuzi" + "geng",
    "\u534e" + "\u80fd",
    "Hua" + "neng",
    "xiao" + "hongshu",
    "\u5bf9" + "\u6bd4\u56fe",
    "\u8d85" + "\u50a8\u7269\u8d44",
    "\u95f2" + "\u7f6e\u7269\u8d44",
    "\u8054" + "\u50a8\u8054\u5907",
    "mark" + "liuzz",
)


def _tracked_files() -> list[Path]:
    result = subprocess.run(
        ["git", "ls-files"],
        cwd=PROJECT_ROOT,
        text=True,
        capture_output=True,
        check=True,
    )
    files = [PROJECT_ROOT / line for line in result.stdout.splitlines() if line.strip()]
    fixture_files = sorted((PROJECT_ROOT / "tests" / "fixtures" / "docx").glob("*.docx"))
    return sorted({*files, *fixture_files})


def _extract_text(path: Path) -> str:
    if path.suffix == ".docx":
        document = Document(str(path))
        return "\n".join(paragraph.text for paragraph in document.paragraphs)
    if path.suffix in TEXT_SUFFIXES:
        return path.read_text(encoding="utf-8")
    return ""


def test_release_surface_has_no_internal_paths_or_business_terms() -> None:
    offenders: list[str] = []
    for path in _tracked_files():
        text = _extract_text(path)
        if not text:
            continue
        relative = path.relative_to(PROJECT_ROOT)
        for pattern in FORBIDDEN_PATTERNS:
            if pattern in text:
                offenders.append(f"{relative}: {pattern}")

    assert offenders == []
