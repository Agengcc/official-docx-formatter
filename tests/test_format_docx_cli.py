from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]
FIXTURE_DIR = Path("/Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test")
FORMAT_CLI = PROJECT_ROOT / "scripts" / "format_docx.py"


def run_format_cli(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [sys.executable, str(FORMAT_CLI), *args],
        cwd=PROJECT_ROOT,
        text=True,
        capture_output=True,
        check=False,
    )


def test_confident_notice_fixture_generates_docx_and_report(tmp_path: Path) -> None:
    input_path = FIXTURE_DIR / "01_格式混乱_通知.docx"
    output_path = tmp_path / "notice.docx"
    report_path = output_path.with_suffix(".report.json")

    result = run_format_cli(str(input_path), "-o", str(output_path), "--report")

    assert result.returncode == 0, result.stderr
    assert output_path.exists()
    assert report_path.exists()
    assert "saved=" in result.stdout
    assert "report=" in result.stdout

    report = json.loads(report_path.read_text(encoding="utf-8"))
    assert report["input"] == str(input_path)
    assert report["output"] == str(output_path)
    assert report["profile_id"] == "standard-party-government"
    assert report["doc_type"] == "通知"
    assert report["structure"]["title_indices"]
    assert report["operations"]


def test_ambiguous_fixture_returns_2_and_does_not_generate_docx(tmp_path: Path) -> None:
    input_path = FIXTURE_DIR / "06_文种暧昧_材料_应先询问.docx"
    output_path = tmp_path / "ambiguous.docx"
    report_path = output_path.with_suffix(".report.json")

    result = run_format_cli(str(input_path), "-o", str(output_path), "--report")

    assert result.returncode == 2
    assert not output_path.exists()
    assert not report_path.exists()
    assert "你希望按哪一种" in result.stdout or "不能可靠判断文种" in result.stdout
