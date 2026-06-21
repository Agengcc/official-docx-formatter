# official-docx-formatter

中文公文 DOCX 格式整理工具。目标是把格式混乱的 Word 公文整理成更接近党政机关公文 / 国央企常用公文习惯的标准格式。

这个项目的定位是“公文格式可信引擎”，不是正文写作助手。

## 能做什么

- 本地读取 `.docx`。
- 识别常见文种，如通知、报告、请示、函、纪要等。
- 文种不确定时先询问，不冒险直接格式化。
- 按 profile 整理字体、字号、页边距、行距、标题层级、落款日期。
- 默认做保守的标点和空格规范化。
- 输出格式化后的 `.docx`。
- 重构后会输出本地 `.report.json`，记录识别结果、格式计划、实际操作和警告。

## 不做什么

- 不润色正文。
- 不重写正文。
- 不总结、扩写、缩写正文。
- 不改变公司名、金额、合同号、项目名等事实信息。
- 不做脱敏模式。
- 不自动补写正文、落款或日期。
- 不自动安装字体。

## 基本用法

```bash
python scripts/classify_document.py input.docx
python scripts/format_docx.py input.docx -o output.docx --profile standard-party-government
```

如果只想改 Word 样式，一个字符都不要动：

```bash
python scripts/format_docx.py input.docx -o output.docx --profile standard-party-government --no-normalize-text
```

## 默认配置

默认 profile 是：

```text
profiles/standard-party-government.json
```

展示名为：

```text
党政机关公文标准配置
```

它以 `GB/T 9704-2012《党政机关公文格式》`、《党政机关公文处理工作条例》和 `GB/T 33476.2-2016《党政机关电子公文格式规范 第2部分：显现》` 为主要参考。

## 开发

安装依赖：

```bash
python -m pip install -e ".[dev]"
```

运行测试：

```bash
python -m pytest tests -v
```

## 文档

- [可信引擎重构计划](docs/superpowers/plans/2026-06-21-official-docx-trusted-engine-refactor.md)
- [后续功能路线图](docs/roadmap-advanced-official-document-features.md)

## 设计边界

这个项目优先追求可信、可解释、可回归测试。宁可在文种和结构不确定时多问一次，也不要把文档“看起来格式化了”，但实际排错了。
