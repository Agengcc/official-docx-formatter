from pathlib import Path

from docx import Document

from scripts.classify_document import classify_lines, read_text


def make_docx(path: Path, lines: list[str]) -> Path:
    document = Document()
    for line in lines:
        document.add_paragraph(line)
    document.save(str(path))
    return path


def classify_fixture(tmp_path: Path, name: str, lines: list[str]):
    return classify_lines(read_text(make_docx(tmp_path / name, lines)))


def test_external_fixtures_classify_confident_document_types(tmp_path: Path):
    cases = [
        ("notice.docx", ["关于开展安全生产检查的通知", "现将有关事项通知如下。"], "通知"),
        ("report.docx", ["关于年度重点工作推进情况的报告", "现将有关情况报告如下。", "特此报告。"], "报告"),
        ("request.docx", ["关于申请调整采购计划的请示", "请予批复。"], "请示"),
        ("letter.docx", ["关于商请协助提供有关资料的函", "请函复。"], "函"),
        ("minutes.docx", ["专题会议纪要", "会议认为", "会议要求"], "纪要"),
    ]

    for filename, lines, expected_type in cases:
        result = classify_fixture(tmp_path, filename, lines)

        assert result["top"]["doc_type"] == expected_type
        assert result["ask_user"] is False


def test_ambiguous_material_fixture_asks_user(tmp_path: Path):
    result = classify_fixture(
        tmp_path,
        "ambiguous.docx",
        ["关于数据治理有关事项的材料", "现将有关情况报告如下。", "请贵单位协助支持。"],
    )

    assert result["ask_user"] is True
    assert result["question"]


def test_low_score_document_asks_user():
    result = classify_lines(["关于推进专项工作的材料", "请有关单位认真研究。"])

    assert result["ask_user"] is True
    assert "通用正式文本" in result["question"]


def test_close_score_document_asks_user():
    result = classify_lines(["关于印发安全检查通报的通知", "请有关单位认真落实。"])

    assert result["ask_user"] is True


def test_report_request_strong_signal_conflict_asks_user():
    result = classify_lines([
        "关于采购设备有关情况的报告",
        "上级单位：",
        "现将有关情况报告如下。",
        "拟申请调整采购方式，请予批复。",
    ])

    assert result["top"]["doc_type"] == "报告"
    assert result["ask_user"] is True


def test_ambiguous_question_offers_generic_formal_text_choice():
    result = classify_lines(["专项工作交流材料", "一、基本情况", "有关工作正在稳步推进。"])

    assert result["ask_user"] is True
    assert "通用正式文本" in result["question"]
