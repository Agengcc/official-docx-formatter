import pytest

from scripts.normalize_text import normalize_text


@pytest.mark.parametrize(
    "text",
    [
        "详情见https://example.com/a:b,c...",
        "请发送至test@example.com。",
        "会议时间09:30。",
        "质量体系按ISO 9001:2015执行。",
        r"附件路径C:\Users\demo\file.txt",
        "合同编号HT-2026-0619。",
    ],
)
def test_special_patterns_are_not_changed(text):
    assert normalize_text(text) == text


def test_normalizes_chinese_punctuation_and_ellipsis():
    assert normalize_text("请各部门:按要求报送,不得拖延...") == "请各部门：按要求报送，不得拖延……"
