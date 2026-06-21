# official-docx-formatter 可信引擎重构计划

> **给执行 agent 的要求：** 真正实施时必须使用 `superpowers:subagent-driven-development`（推荐）或 `superpowers:executing-plans`，按任务逐项执行。本文用 checkbox 作为进度追踪。

**目标：** 把 `official-docx-formatter` 重构成一个本地优先、可解释、可回归测试的“公文格式可信引擎”。

**架构方向：** 不再让 `format_docx.py` 一个脚本同时承担读取、识别、判断、修改和输出。改成稳定流水线：读取 DOCX → 文种分类 → 结构识别 → 格式诊断 → 生成格式计划 → 执行格式计划 → 输出改动报告。

**技术栈：** Python 3、`python-docx`、JSON profile、`pytest`、现有测试样例目录。

---

## 一、工作理解

项目根目录：

```text
/Users/liuzigeng/Ageng的自媒体/公文写作项目
```

当前实际运行的 skill 在三处：

```text
/Users/liuzigeng/.codex/skills/official-docx-formatter
/Users/liuzigeng/.claude/skills/official-docx-formatter
/Users/liuzigeng/.cc-switch/skills/official-docx-formatter
```

执行时先改 Codex 版本，测试通过后再同步到 Claude 和 cc-switch。

测试样例目录保持为：

```text
/Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test
```

---

## 二、边界和原则

### 要做

- 识别文种。
- 识别标题、主送机关、正文、附件说明、落款、日期等结构。
- 诊断格式风险。
- 根据 profile 套用字体、字号、行距、页边距、标题层级、落款日期位置。
- 默认做保守的标点和空格规范化。
- 输出 `.docx` 和一份本地 `.report.json` 改动报告。
- 用测试样例锁住行为，避免越改越玄学。

### 不做

- 不做通用文档平台。
- 不做学术、法律、营销等多预设体系。
- 不做正文润色、改写、总结、扩写、缩写。
- 不做脱敏模式。
- 不做自动补写落款、日期、正文内容。
- 不做 `.doc` / `.wps` 转换。
- 不自动安装字体。
- 不复制外部仓库代码，只借鉴工程思路。

---

## 三、成功标准

- 6 个现有 DOCX 样例中，通知、报告、请示、函、纪要能高置信度识别。
- `06_文种暧昧_材料_应先询问.docx` 必须先询问用户，不直接格式化。
- 格式化普通样例后，同时生成：
  - 格式化后的 `.docx`
  - 对应 `.report.json`
- 报告里至少包含：
  - 输入文件
  - 输出文件
  - 使用的 profile
  - 文种
  - 识别出的结构
  - 应用的格式操作
  - 标点空格规范化状态
  - 警告和风险提示
- `--no-normalize-text` 能保证正文字符不变，只改 Word 样式。
- 标点规范化必须保护：
  - URL
  - 邮箱
  - 时间，如 `9:30`
  - 标准编号，如 `ISO 9001:2015`
  - Windows 路径
  - 合同号、编号类内容
- 结构识别至少覆盖：
  - 多行标题
  - 标题前空段
  - 点分日期不误判为三级标题
  - 长落款
  - 联合落款
  - 附件说明
  - 落款日期识别

---

## 四、总体文件结构

主要修改目录：

```text
/Users/liuzigeng/.codex/skills/official-docx-formatter
```

新增核心包：

```text
scripts/official_docx_engine/
  __init__.py
  models.py              # 中立数据模型
  docx_reader.py         # DOCX 读取，不做修改
  structure.py           # 结构识别
  diagnostics.py         # 格式诊断
  format_plan.py         # 生成格式计划
  apply_docx.py          # 执行格式计划
  reporting.py           # 输出报告
```

新增命令：

```text
scripts/diagnose_docx.py
```

保留并改薄：

```text
scripts/classify_document.py
scripts/format_docx.py
scripts/normalize_text.py
scripts/profile_manager.py
```

新增测试：

```text
tests/conftest.py
tests/test_classify_fixtures.py
tests/test_normalize_text.py
tests/test_structure_detection.py
tests/test_format_plan.py
tests/test_format_docx_cli.py
tests/fixtures/text_cases.py
```

同步更新文档：

```text
SKILL.md
evals/evals.json
references/standards.md
references/profiles.md
/Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test/README.md
```

明确不碰：

```text
/Users/liuzigeng/华能/hn-skills/hn-docx-formatter
```

---

## 五、任务拆解

### Task 1：建立测试基线

**目标：** 先用现有 6 个 Word 样例锁住当前最核心行为。

**文件：**

- 新增：`tests/conftest.py`
- 新增：`tests/test_classify_fixtures.py`
- 修改：`official-docx-formatter-test/README.md`

**步骤：**

- [ ] 新增 fixture 路径配置，统一指向：

```text
/Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test
```

- [ ] 为 5 个明确文种样例写分类测试：

```text
01_格式混乱_通知.docx              => 通知
02_格式混乱_报告_含敏感信息.docx    => 报告
03_格式混乱_请示.docx              => 请示
04_格式混乱_公函.docx              => 函
05_格式混乱_会议纪要.docx          => 纪要
```

- [ ] 为暧昧样例写测试：

```text
06_文种暧昧_材料_应先询问.docx
```

预期：`ask_user = true`，不能直接格式化。

- [ ] 运行：

```bash
cd /Users/liuzigeng/.codex/skills/official-docx-formatter
python -m pytest tests/test_classify_fixtures.py -v
```

- [ ] 修改测试目录 README，删除已经废弃的 `redact_document.py` 示例，替换成：

```bash
python /Users/liuzigeng/.codex/skills/official-docx-formatter/scripts/format_docx.py 02_格式混乱_报告_含敏感信息.docx -o /tmp/报告_格式化.docx --no-normalize-text
```

---

### Task 2：建立中立 DOCX 数据模型

**目标：** 先读取文档状态，不急着改文档。让后续所有判断都有共同输入。

**文件：**

- 新增：`scripts/official_docx_engine/models.py`
- 新增：`scripts/official_docx_engine/docx_reader.py`

**核心模型：**

- `ParagraphSnapshot`
  - 段落序号
  - 段落文字
  - 样式名
  - 对齐方式
  - 缩进
  - 行距
  - 字体集合
  - 字号集合
  - 是否在表格内

- `DocumentSnapshot`
  - 文件路径
  - 段落列表
  - 表格数量

- `StructureAnalysis`
  - 标题
  - 主送机关
  - 正文段落序号
  - 落款段落序号
  - 日期段落序号
  - 识别证据
  - 警告

- `DiagnosticIssue`
  - 问题编码
  - 严重程度
  - 说明
  - 相关段落

- `FormatOperation`
  - 操作 ID
  - 操作类型
  - 影响段落
  - 操作说明
  - 修改前摘要
  - 修改后摘要

- `FormatPlan`
  - profile
  - 文种
  - 操作列表
  - 诊断问题

**验收：**

运行一个 fixture，能打印出段落数量和表格数量。

---

### Task 3：重构结构识别

**目标：** 把现在散在 `format_docx.py` 里的标题、主送、落款、日期判断，独立成上下文感知模块。

**文件：**

- 新增：`scripts/official_docx_engine/structure.py`
- 新增：`tests/fixtures/text_cases.py`
- 新增：`tests/test_structure_detection.py`

**优先覆盖场景：**

- 多行标题：

```text
关于进一步加强
安全生产管理工作的通知
```

应合并识别为：

```text
关于进一步加强安全生产管理工作的通知
```

- 点分日期：

```text
2026.04.20，项目完成现场检查。
```

应识别为正文，不是三级标题。

- 附件说明：

```text
附件：检查清单
```

应识别为附件说明。

- 联合落款：

```text
甲单位  乙单位
2026年6月21日
```

应识别为落款和日期。

**验收：**

```bash
cd /Users/liuzigeng/.codex/skills/official-docx-formatter
python -m pytest tests/test_structure_detection.py -v
```

---

### Task 4：新增诊断能力

**目标：** 格式化之前先能回答“这篇文档目前有哪些格式问题和结构风险”。

**文件：**

- 新增：`scripts/official_docx_engine/diagnostics.py`
- 新增：`scripts/diagnose_docx.py`

**诊断内容第一版：**

- 是否识别到标题。
- 是否识别到主送机关。
- 是否识别到日期。
- 是否存在表格。
- 字体是否混杂。
- 字号是否混杂。
- 段落缩进是否不统一。
- 行距是否不统一。
- 是否有可能误判的结构。

**命令：**

```bash
python scripts/diagnose_docx.py input.docx --json
```

**输出示例：**

```json
{
  "title": "关于进一步加强安全生产管理工作的通知",
  "recipient": "各部门",
  "body_paragraphs": 6,
  "issuer_indexes": [8],
  "date_index": 9,
  "issues": [
    {
      "code": "mixed_fonts",
      "severity": "warning",
      "message": "发现多种字体"
    }
  ]
}
```

---

### Task 5：引入“格式计划”

**目标：** 执行前先生成一份计划，说明准备改什么，而不是直接改。

**文件：**

- 新增：`scripts/official_docx_engine/format_plan.py`
- 新增：`tests/test_format_plan.py`

**格式计划包含：**

- 页面设置操作：
  - A4
  - 上 3.7cm
  - 下 3.5cm
  - 左 2.8cm
  - 右 2.6cm

- 标题操作：
  - 居中
  - 二号标题字体

- 主送机关操作：
  - 正文字体
  - 保留冒号

- 正文操作：
  - 三号仿宋类字体
  - 固定行距
  - 首行缩进

- 层级标题操作：
  - `一、` 黑体
  - `（一）` 楷体
  - `1.` 正文字体
  - `（1）` 正文字体

- 落款日期操作：
  - 按 profile 右对齐
  - 不自动补不存在的落款或日期

- 文本规范化操作：
  - 只在默认开启时出现
  - `--no-normalize-text` 时不出现

**验收：**

测试能确认 plan 中存在：

```text
page_setup
paragraph_style
signature_layout
text_normalization
```

并且关闭规范化时没有 `text_normalization`。

---

### Task 6：加固标点空格规范化边界

**目标：** 保留低风险字符级修复，但用测试证明它不碰危险内容。

**文件：**

- 新增：`tests/test_normalize_text.py`
- 修改：`scripts/normalize_text.py`

**必须通过的保护测试：**

输入：

```text
请访问 https://example.com:8080/path，联系 report@example.com，会议 9:30 开始，执行 ISO 9001:2015，路径 C:\Users\demo，合同 HT-2026-0619。
```

输出必须仍然包含：

```text
https://example.com:8080/path
report@example.com
9:30
ISO 9001:2015
C:\Users\demo
HT-2026-0619
```

**允许规范化的例子：**

```text
请各部门:按要求报送,不得拖延...
```

变成：

```text
请各部门：按要求报送，不得拖延……
```

**验收：**

```bash
cd /Users/liuzigeng/.codex/skills/official-docx-formatter
python -m pytest tests/test_normalize_text.py -v
```

---

### Task 7：把 formatter 改成“计划后执行”

**目标：** `format_docx.py` 变成薄入口，核心逻辑移到 engine 包里。

**文件：**

- 新增：`scripts/official_docx_engine/apply_docx.py`
- 新增：`scripts/official_docx_engine/reporting.py`
- 修改：`scripts/format_docx.py`
- 新增：`tests/test_format_docx_cli.py`

**新的执行流程：**

1. 读取 DOCX snapshot。
2. 文种分类。
3. 如果文种不确定，退出并询问用户。
4. 结构识别。
5. 格式诊断。
6. 生成格式计划。
7. 执行格式计划。
8. 保存 `.docx`。
9. 输出 `.report.json`。

**新增参数：**

```bash
--report /path/to/report.json
```

如果用户不指定，默认生成：

```text
输出文件名.report.json
```

**命令输出示例：**

```text
saved=/tmp/通知_格式化.docx
report=/tmp/通知_格式化.report.json
profile=standard-party-government
text_normalization=keep_en_boundary
paragraphs=12
```

**关键测试：**

- 明确文种样例能生成 `.docx` 和 `.report.json`。
- 暧昧文种样例返回退出码 `2`，不生成输出文件。

---

### Task 8：加强分类置信度规则

**目标：** 让分类更适合“可信引擎”，宁可多问一次，也不要误排。

**文件：**

- 修改：`scripts/classify_document.py`
- 修改：`references/document_types.json`
- 修改：`tests/test_classify_fixtures.py`

**规则：**

- 标题证据权重大于正文证据。
- 正文提到另一个文种时，不轻易覆盖标题文种。
- 报告和请示同时有强信号时，要问用户。
- 函和请示同时有强信号时，也要谨慎。
- top score 太低时必须问。
- top 和 second 分差太小时必须问。

**验收：**

```bash
cd /Users/liuzigeng/.codex/skills/official-docx-formatter
python -m pytest tests/test_classify_fixtures.py -v
```

---

### Task 9：更新 Skill 文档和 eval

**目标：** 让 skill 使用说明和新架构一致。

**文件：**

- 修改：`SKILL.md`
- 修改：`evals/evals.json`
- 修改：`references/standards.md`
- 修改：`references/profiles.md`

**SKILL.md 需要明确：**

- 普通格式化会自动执行分类、结构识别、诊断、格式计划、报告输出。
- 需要单独审查时，可以先运行：

```bash
python scripts/diagnose_docx.py input.docx --json
```

- 每次完成后要告诉用户：
  - 输出 docx 路径
  - report json 路径
  - 使用的 profile
  - 文种
  - 是否做了标点空格规范化
  - 是否有警告

- 不要把全文粘贴到对话里。

**eval 增加两个重点：**

1. 文种暧昧时必须先问。
2. 格式化后必须生成 report。

---

### Task 10：全量验证和同步

**目标：** 确认 Codex 版本稳定后，同步到 Claude 和 cc-switch。

**步骤：**

- [ ] 运行全部测试：

```bash
cd /Users/liuzigeng/.codex/skills/official-docx-formatter
python -m pytest tests -v
```

- [ ] 跑诊断 smoke test：

```bash
python scripts/diagnose_docx.py /Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test/01_格式混乱_通知.docx --json
```

- [ ] 跑格式化 smoke test：

```bash
python scripts/format_docx.py /Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test/01_格式混乱_通知.docx -o /tmp/通知_可信引擎_格式化.docx
```

预期生成：

```text
/tmp/通知_可信引擎_格式化.docx
/tmp/通知_可信引擎_格式化.report.json
```

- [ ] 跑暧昧文种阻断测试：

```bash
python scripts/format_docx.py /Users/liuzigeng/Ageng的自媒体/公文写作项目/official-docx-formatter-test/06_文种暧昧_材料_应先询问.docx -o /tmp/暧昧_不应生成.docx
```

预期：

```text
退出码 2
不生成 /tmp/暧昧_不应生成.docx
提示用户选择文种
```

- [ ] 同步三处 skill：

```bash
cd /Users/liuzigeng/Ageng的自媒体/公文写作项目
./sync-official-docx-formatter-to-claude.sh
```

- [ ] 验证同步结果：

```bash
test -f /Users/liuzigeng/.claude/skills/official-docx-formatter/scripts/official_docx_engine/models.py
test -f /Users/liuzigeng/.cc-switch/skills/official-docx-formatter/scripts/official_docx_engine/models.py
```

---

## 六、阶段优先级

### P0：先做可信底座

必须先完成：

```text
Task 1 测试基线
Task 2 中立模型
Task 3 结构识别
Task 4 诊断能力
Task 5 格式计划
Task 7 计划后执行
```

原因：这些决定“可信引擎”的骨架。

### P1：加固安全边界

接着完成：

```text
Task 6 标点空格规范化边界
Task 8 分类置信度
Task 9 文档和 eval
```

原因：这些决定工具不会越界。

### P2：同步发布

最后完成：

```text
Task 10 全量验证和同步
```

---

## 七、暂缓事项

这些以后可以单独开计划，不放进本次重构：

- 页码模块。
- 表格内部格式整理。
- 红头、版记、印章、页脚域。
- `.doc` / `.wps` 转换。
- GUI。
- 单位模板 `.dotx` 支持。
- 更多企业 profile。

原因：它们都有真实价值，但会显著扩大本次重构的风险面。当前目标是先把普通正文型公文的可信流水线立住。

---

## 八、风险控制

- 保持现有命令兼容：

```bash
python scripts/format_docx.py input.docx -o output.docx --profile standard-party-government
```

- 保持 `--no-normalize-text` 兼容。
- 新增 report，但不要求用户手动指定路径。
- 表格第一阶段只诊断，不强行重排。
- 缺字体只告警，不安装。
- 文种不确定时停止，不冒险格式化。
- 不把正文全文输出到对话中。

---

## 九、review 时建议重点看

你 review 时可以重点看这几件事：

1. 这个边界是否对：只做格式可信引擎，不做正文写作助手。
2. 是否同意“诊断和计划先行”，而不是继续增强一个大脚本。
3. 是否接受第一阶段暂缓页码、表格、红头、`.doc/.wps`。
4. report 是否是你想要的可信凭证。
5. 分类不确定时“宁可问，不直接排”的策略是否符合你的使用习惯。

---

## 十、执行方式

计划已保存到：

```text
/Users/liuzigeng/Ageng的自媒体/公文写作项目/docs/superpowers/plans/2026-06-21-official-docx-trusted-engine-refactor.md
```

确认后有两种执行方式：

1. **Subagent-Driven，推荐。** 每个任务派一个 fresh subagent 做，我负责 review 和合并判断。
2. **Inline Execution。** 我在当前会话按任务顺序直接做，中间给你 checkpoint。
