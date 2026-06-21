from pathlib import Path

from scripts.classify_document import classify_lines, read_text


FIXTURE_DIR = Path("/Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test")


def classify_fixture(name: str):
    return classify_lines(read_text(FIXTURE_DIR / name))


def test_external_fixtures_classify_confident_document_types():
    cases = [
        ("01_格式混乱_通知.docx", "通知"),
        ("02_格式混乱_报告_含敏感信息.docx", "报告"),
        ("03_格式混乱_请示.docx", "请示"),
        ("04_格式混乱_公函.docx", "函"),
        ("05_格式混乱_会议纪要.docx", "纪要"),
    ]

    for filename, expected_type in cases:
        result = classify_fixture(filename)

        assert result["top"]["doc_type"] == expected_type
        assert result["ask_user"] is False


def test_ambiguous_material_fixture_asks_user():
    result = classify_fixture("06_文种暧昧_材料_应先询问.docx")

    assert result["ask_user"] is True
    assert result["question"]


def test_low_score_document_asks_user():
    result = classify_lines(["关于推进专项工作的材料", "请有关单位认真研究。"])

    assert result["ask_user"] is True


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
