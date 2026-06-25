# official-docx-formatter

一套给 AI Agent 和本地命令行使用的中文公文 Word 格式整理工具。

把已经写好的 `.docx` 初稿整理成更接近党政机关公文习惯的标准版式：字体、字号、页边距、行距、标题层级、落款日期、页码、表格、目录和版记。

核心原则很简单：只整理格式，不替你改正文。

---

## 适合谁

- 行政、综合、办公室、运营、材料岗。
- 经常处理通知、报告、请示、函、纪要等正式 Word 材料的人。
- 已经有正文初稿，但 Word 格式混乱、不统一、不确定是否符合公文习惯的团队。
- 想把“公文排版”交给 AI Agent 或脚本自动完成，但又不希望正文被乱改的人。

---

## 能做什么

- 识别常见文种：通知、报告、请示、函、纪要等。
- 文种不确定时先停下来提示选择，不冒险直接格式化。
- 使用默认党政机关公文标准配置整理版式。
- 保留 Word/WPS 自动编号里可见的 `一、`、`二、` 等层级前缀。
- 可选规范中文标点、空格和中英文混排间距。
- 可选插入页码、整理表格、生成目录、整理已有版记。
- 支持“通用正式文本”模式，用来整理不确定文种的普通材料。
- 自动识别带 `目次` / `前    言` / `1  范围` 结构的标准规范文本草案。
- 输出 `.report.json`，记录识别结果、格式动作和警告。

---

## 不做什么

- 不润色正文。
- 不重写、扩写、缩写、总结正文。
- 不改公司名、金额、合同号、项目名等事实信息。
- 不做脱敏或内容替换。
- 不自动补写正文、主送单位、落款或日期。
- 不自动安装字体。
- 不提供自定义单位格式入口；当前版本只保留默认党政机关公文标准配置。

---

## 核心逻辑

```text
读取 Word
  ↓
识别文种和结构
  ↓
不确定时停止并询问
  ↓
按默认公文配置整理格式
  ↓
输出 .docx + .report.json
```

工具优先追求可信、可解释、可回归测试。宁可多问一次，也不要把文档“看起来格式化了”，但实际排错了。

---

## 文件结构

```text
official-docx-formatter/
├── SKILL.md                    AI Agent 使用说明
├── README.md                   项目说明
├── profiles/                   默认格式配置
│   └── standard-party-government.json
├── references/                 文种、字体、标准依据说明
├── scripts/                    本地命令行工具
│   ├── classify_document.py    判断文种
│   ├── format_docx.py          格式化 Word
│   └── official_docx_engine/   结构识别、诊断、页码、表格、目录、版记等能力
└── tests/                      回归测试
```

---

## 安装

```bash
git clone https://github.com/Agengcc/official-docx-formatter.git
cd official-docx-formatter
python -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"
```

---

## 怎么用

### 1. 先判断文种

```bash
python scripts/classify_document.py input.docx
```

### 2. 格式化 Word

```bash
python scripts/format_docx.py input.docx -o output.docx --report
```

执行后会生成：

```text
output.docx
output.report.json
```

### 3. 如果只想改样式，不想改标点和空格

```bash
python scripts/format_docx.py input.docx -o output.docx --report --no-normalize-text
```

---

## 常用模式

### 加页码

```bash
python scripts/format_docx.py input.docx -o output.docx --report --page-numbers
```

### 整理表格

```bash
python scripts/format_docx.py input.docx -o output.docx --report --format-tables
```

### 生成目录

```bash
python scripts/format_docx.py input.docx -o output.docx --report --generate-toc
```

### 整理已有版记

```bash
python scripts/format_docx.py input.docx -o output.docx --report --format-imprint
```

### 普通正式材料

适合文种不明确、但想把材料整理得正式、整齐的 Word。

```bash
python scripts/format_docx.py input.docx -o output.docx --report --generic-formal-text
```

### 标准规范文本

适合带 `目次`、`前    言`、`1  范围` 等结构的标准草案。

```bash
python scripts/format_docx.py standard-spec.docx -o output.docx --report --standard-spec-text
```

---

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

---

## 测试

```bash
python -m pytest tests -v
```

---

## 反馈

如果在使用过程中有问题，请提交 issue，以确保我们会及时更正。
