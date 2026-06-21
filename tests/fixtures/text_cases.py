from __future__ import annotations

from official_docx_engine.models import DocumentSnapshot, ParagraphSnapshot, RunSnapshot


def snapshot_from_lines(lines: list[str], table_count: int = 0) -> DocumentSnapshot:
    return DocumentSnapshot(
        path=None,
        table_count=table_count,
        paragraphs=tuple(
            ParagraphSnapshot(
                index=index,
                text=text,
                style_name="Normal",
                runs=(RunSnapshot(text=text, font_name="仿宋", font_size_pt=16.0, bold=False),) if text else (),
            )
            for index, text in enumerate(lines)
        ),
    )


MULTILINE_NOTICE = [
    "关于开展项目档案专项检查",
    "有关事项的通知",
    "各部门：",
    "为进一步规范项目档案管理，现将有关事项通知如下。",
    "2026.04.20，项目完成现场检查。",
    "附件：项目档案检查清单",
    "南京示例科技有限公司",
    "2026年4月21日",
]


MISSING_RECIPIENT_REPORT = [
    "项目推进情况报告",
    "今年以来，项目组按计划完成现场调研和资料汇总。",
    "南京示例科技有限公司",
    "2026年4月21日",
]


MIXED_RUN_PARAGRAPHS = (
    ParagraphSnapshot(
        index=0,
        text="项目推进情况报告",
        runs=(
            RunSnapshot("项目推进", font_name="宋体", font_size_pt=22.0),
            RunSnapshot("情况报告", font_name="黑体", font_size_pt=18.0),
        ),
    ),
)
