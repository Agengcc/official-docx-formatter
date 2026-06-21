---
name: official-docx-formatter
description: >-
  Classify and format Chinese official .docx documents with a local-first format-and-normalization safety model. Use this skill whenever the user asks to format, standardize, clean up, create, or inspect Chinese official documents, including 请示、报告、通知、批复、通报、函/公函、纪要、决定 and other party-government document types. First read/classify the document locally; if uncertain, ask the user to choose before changing styles. This skill changes Word styling, layout, page setup, hierarchy fonts, signature/date placement, and conservative punctuation/spacing normalization; it does not rewrite, polish, redact, summarize, or optimize body text. Prefer this skill over ad hoc DOCX edits whenever the user mentions 公文、红头文件、国标格式、Word 格式混乱、文种、函、报告、通知、请示、字体/行距/页边距, 标点/空格规范, or organization-specific formatting.
---

# Official DOCX Formatter

## Purpose

Use this skill to turn an unformatted or inconsistently formatted `.docx` into a clean Chinese official-document style with a local trusted-engine workflow. For ordinary formatting, execute locally: classify the document, identify structure, diagnose formatting issues, build a formatting plan, apply formatting, and output a local report. For existing documents, classify the document type before changing the file. Different document types have different structure rules; formatting is only the second step.

Do not modify legacy enterprise-specific skills unless the user explicitly asks. This skill is the general-purpose formatter.

## Safety Boundary

- Read and classify the document through local scripts.
- Do not ask the model to rewrite, polish, summarize, or optimize body text.
- Do not print or paste the full body text into the conversation for ordinary formatting tasks.
- Change Word styles, margins, line spacing, hierarchy fonts, and page setup.
- Apply only conservative text normalization: Chinese punctuation, ellipsis/dash, paired quotes, and spacing around Chinese/English/digits.
- Preserve substantive wording, company names, amounts, contract numbers, project names, facts, and paragraph order.
- Protect URLs, email addresses, times, standards, and code-like identifiers during punctuation normalization.
- Do not perform any redaction/refill cycle and do not desensitize user documents.
- Do not automatically add missing body text, issuer/signature, or date.
- Do not automatically install fonts; report missing fonts or follow fallback/profile guidance.
- If the user asks for content rewriting, explain that this formatter is format-only and suggest doing that as a separate task or separate skill.

## Default Question

When no profile is specified, ask:

> 请选择格式配置：  
> 1. 党政机关公文标准配置：参考 GB/T 9704-2012，推荐用于通用正式公文  
> 2. 自定义单位配置：使用你所在单位指定的字体、行距、页边距等要求

If the user chooses custom configuration, collect only the fields they know. Missing fields must inherit from `profiles/standard-party-government.json`.

## Workflow

1. Identify the task:
   - Convert an existing `.docx`
   - Create a new formatted `.docx`
   - Inspect or define a formatting profile
   - Answer a format requirement question

2. Read and classify before formatting:
   - For an existing `.docx`, run or follow `scripts/classify_document.py` before making style changes.
   - Classify from title, body signals, ending phrases, recipient relationship, and obvious intent.
   - Use `references/document_types.json` as the catalog. It includes the 15 official document types from `党政机关公文处理工作条例`; common enterprise use is usually 请示、报告、通知、批复、通报、函/公函、纪要、决定.
   - If confidence is low or the top candidates are close, stop and ask the user to choose. Do not format first and ask later.
   - Use this short question: `我看这篇更像是【A】或【B】。你希望按哪一种文种来排：A / B / 其他？`

3. Confirm or choose the document type:
   - For conversion, proceed directly only when the type is obvious from title or fixed ending phrases.
   - For creation, ask the user for the document type before drafting the title/body skeleton.
   - Read `references/document_types.md` when the type choice affects structure or wording.

4. Select a formatting profile:
   - Use `profiles/standard-party-government.json` by default.
   - Use another profile from `profiles/` when the user names one.
   - Create a new profile only when the user gives organization-specific settings.

5. For conversion:
   - Preserve original text order.
   - Reuse existing title, recipient, body, attachment notes, issuer, and date when detectable.
   - Do not add placeholder issuer/date if the source document has no clear issuer/date and the user did not provide them.
   - Apply the confirmed document type's structural rules and the selected profile's typography, margins, line spacing, paragraph spacing, and hierarchy rules.
   - Normalize punctuation and spacing unless the user explicitly asks to preserve characters exactly. Use `--no-normalize-text` for exact-character preservation.
   - For issuer/date placement, follow the selected profile. The default profile uses no-seal single-issuer placement: one blank line after the body or attachment note, then issuer/date on the right. Read `references/standards.md` before changing this behavior.
   - For font availability, follow `references/fonts.md`; do not auto-install fonts.
   - Output both the formatted `.docx` and the corresponding local `.report.json`.

6. Verify:
   - Confirm the output `.docx` path and report JSON path.
   - Report the profile, document type, text-normalization status, and any warnings.
   - Run a smoke check when possible by opening the generated `.docx` with `python-docx` and reporting a compact status.
   - Do not paste the full document text or full report into the conversation; summarize the key status fields only.
   - If exact visual validation is needed, tell the user that Word/WPS rendering may need manual inspection.

## Scripts

- `scripts/format_docx.py`: Convert an existing `.docx` or create a simple formatted document from text.
- `scripts/classify_document.py`: Read `.docx` or text and return likely document type candidates.
- `scripts/diagnose_docx.py`: Inspect/audit a `.docx` locally and emit document type, structure, diagnostics, planned operations, and warnings. Use `--json` for machine-readable output.
- `scripts/normalize_text.py`: Conservative punctuation and spacing normalization without rewriting substantive content.
- `scripts/profile_manager.py`: List, show, and create formatting profiles.
- `evals/evals.json`: Realistic test prompts and expectations for future iterations.
- `references/document_types.json`: Machine-readable document-type structures used for new document skeletons.

Run examples:

```bash
python scripts/profile_manager.py list
python scripts/classify_document.py input.docx
python scripts/diagnose_docx.py input.docx --json
python scripts/profile_manager.py create my-company --title-font "方正小标宋简体" --body-font "仿宋_GB2312"
python scripts/format_docx.py input.docx -o output.docx --profile standard-party-government
python scripts/format_docx.py input.docx -o output.docx --profile standard-party-government --no-normalize-text
python scripts/format_docx.py -o report.docx --doc-type 报告 --title "关于××工作的报告" --recipient "上级单位" --issuer "某单位" --date 2026年6月19日 --create-skeleton
```

## References

Read these only when needed:

- `references/standards.md`: Standard basis and what is hard standard vs Word implementation.
- `references/profiles.md`: How organization profiles override the default standard.
- `references/document_types.md`: Difference between formatting profiles and document-type structures.
- `references/fonts.md`: Font fallback and missing-font policy.

## Packaging Hygiene

Before packaging or sharing this skill, exclude generated cache files such as `__pycache__/`, `.pyc`, temporary `.docx` outputs, and local evaluation workspaces. The skill should contain source instructions, profiles, references, scripts, and eval definitions only.

## Formatting Rules

Default hierarchy:

| Element | Default font | Size | Notes |
| --- | --- | --- | --- |
| Title | 小标宋体 fallback list | 2 hao / 22 pt | Centered |
| Body | 仿宋体 fallback list | 3 hao / 16 pt | Two-character first-line indent |
| Level 1 heading `一、` | 黑体 fallback list | 3 hao / 16 pt | Bold |
| Level 2 heading `（一）` | 楷体 fallback list | 3 hao / 16 pt | Not bold |
| Level 3 heading `1.` | Body font | 3 hao / 16 pt | Not bold |
| Level 4 heading `（1）` | Body font | 3 hao / 16 pt | Not bold |

Only apply heading fonts to short standalone heading paragraphs. If a paragraph mixes a numbered prefix with sentence content, format the whole paragraph as body text.
