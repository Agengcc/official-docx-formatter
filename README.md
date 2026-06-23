# official-docx-formatter

中文公文 `.docx` 本地格式整理工具。

它适合已经有 Word 初稿、但格式不稳定的人：行政、综合、办公室、运营、国企/事业单位材料岗，以及需要把通知、报告、请示、函、纪要等正式材料整理成统一版式的团队。

这个项目只做格式整理，不做正文写作。

## 适合用来做什么

- 把格式混乱的 `.docx` 整理成接近党政机关公文习惯的版式。
- 识别常见文种，例如通知、报告、请示、函、纪要等。
- 文种不确定时先停止并提示选择，避免直接排错。
- 按默认党政机关公文标准配置整理字体、字号、页边距、行距、标题层级、落款日期。
- 保留 Word/WPS 自动编号中可见的 `一、`、`二、` 等层级前缀。
- 可选规范中文标点、空格和中英文混排间距。
- 可选插入页码、整理表格、生成目录、整理已有版记。
- 可用“通用正式文本”模式整理普通正式材料，不强行补公文结构。
- 自动识别带 `目次` / `前    言` / `1  范围` 结构的标准规范文本草案，并按规范文本习惯整理。
- 生成本地 `.report.json`，记录识别结果、格式动作和警告。

## 不适合做什么

- 不润色、扩写、缩写、总结正文。
- 不改公司名、金额、合同号、项目名等事实信息。
- 不做脱敏或内容重写。
- 不自动补写正文、主送单位、落款或日期。
- 不自动安装字体。
- 不提供自定义单位格式入口；当前版本只保留默认党政机关公文标准配置。

## 安装

```bash
git clone https://github.com/markliuzz/-official-docx-formatter.git
cd -official-docx-formatter
python -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"
```

## 基本使用

先判断文种：

```bash
python scripts/classify_document.py input.docx
```

格式化 Word：

```bash
python scripts/format_docx.py input.docx -o output.docx --report
```

如果只想改样式，不想做标点和空格规范化：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --no-normalize-text
```

## 常用选项

加入普通页码：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --page-numbers
```

保留并整理表格：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --format-tables
```

在标题层级清晰时生成 Word 目录字段：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --generate-toc
```

整理文末已有版记：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --format-imprint
```

把不确定文种的普通材料整理成正式文本版式：

```bash
python scripts/format_docx.py input.docx -o output.docx --report --generic-formal-text
```

显式按标准规范文本处理：

```bash
python scripts/format_docx.py standard-spec.docx -o output.docx --report --standard-spec-text
```

## 默认格式

当前版本只使用一套默认配置：

```text
profiles/standard-party-government.json
```

展示名：

```text
党政机关公文标准配置
```

主要参考：

- `GB/T 9704-2012《党政机关公文格式》`
- 《党政机关公文处理工作条例》
- `GB/T 33476.2-2016《党政机关电子公文格式规范 第2部分：显现》`

## 测试

```bash
python -m pytest tests -v
```

## 设计边界

这个工具优先追求可信、可解释、可回归测试。宁可在文种和结构不确定时多问一次，也不要把文档“看起来格式化了”，但实际排错了。
