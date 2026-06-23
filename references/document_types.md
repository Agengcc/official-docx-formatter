# Document Types

Keep document type separate from formatting configuration:

- Formatting configuration answers: which fonts, margins, line spacing, and page rules should be used?
- Document type answers: what structure, title habit, main constraints, and ending phrase should the document follow?

This lets the skill combine document type with the default standard configuration, for example:

- `报告` + `standard-party-government`
- `通用正式文本` + `standard-party-government` for ambiguous existing materials that only need clean formal typography

## Common Types

| Type | Direction | Use When | Key Notes |
| --- | --- | --- | --- |
| 请示 | Upward | Request instructions or approval from a superior | Usually one matter per document; normally one main recipient; common ending: `妥否，请示。` |
| 报告 | Upward | Report work, reflect situations, or answer superior inquiries | Do not include approval-request matters; common ending: `特此报告。` |
| 通知 | Downward or broad circulation | Arrange work, publish or forward matters, ask recipients to know or execute | For forwarded documents, add concrete implementation requirements. |
| 函 | Parallel | Communicate between organizations without direct subordination | Use for consultation, inquiry, reply, request for support; common endings vary, such as `请函复。` |
| 纪要 | Circulation | Record meeting situation and agreed matters | Focus on meeting information and agreed items. |
| 批复 | Downward | Reply to a lower-level request for instructions | Usually references the received request and states the approval opinion. |
| 通报 | Downward or broad circulation | Commend, criticize, communicate important spirit or situations | Often includes facts, evaluation, and requirements. |
| 决定 | Downward | Make decisions and deployments on important matters | Title usually should be complete: issuer + matter + document type. |

## Creation Flow

When creating a new document, ask for document type first if the user has not named it. Then ask only for the missing essentials:

- Title or matter
- Recipient
- Main points/body
- Issuer and date if the user wants a complete output

Use `references/document_types.json` for aliases, classification signals, skeleton sections, constraints, and default ending phrases. Treat skeleton text as placeholders to be replaced, not final copy.

## Classification Rule

For an existing document, classify before formatting:

1. Read the title and first meaningful paragraphs.
2. Check whether the title contains an official document type or alias, for example `报告`, `请示`, `通知`, `函`, `公函`, `纪要`.
3. Check ending phrases and body signals, for example `特此报告。`, `妥否，请示。`, `请函复。`, `现将有关事项通知如下`.
4. If the top candidate is weak or close to another candidate, ask the user to choose before editing the Word file. Offer `通用正式文本` when the user only wants generic formal typography and does not want to force an official document type.

The 15 official document types are: 决议、决定、命令、公报、公告、通告、意见、通知、通报、报告、请示、批复、议案、函、纪要. In enterprise use, the high-frequency set is usually 请示、报告、通知、批复、通报、函/公函、纪要、决定, which is also the set emphasized by the legacy HN formatter.

`通用正式文本` is not one of the 15 official document types. Use it only for formatting an existing ambiguous material. It should preserve paragraph order, treat the first non-empty paragraph as the title, format numbered headings when clear, and avoid adding or extracting recipient, issuer, date, red-head, imprint, or seal-related structure.

## Layout Notes by Type

These notes are defaults for ordinary Word/WPS drafts. Red-head templates, seal templates, OA export metadata, and organization-specific templates are outside the current default formatter scope unless implemented as dedicated features.

| Type | Ending / Signature Habit |
| --- | --- |
| 报告 | Usually ends after the reporting conclusion, for example `特此报告。`; issuer and date stay at the end of the document. For no-seal single-issuer drafts, leave one blank line after the body, then place issuer/date on the right. Do not place issuer/date at the top. |
| 请示 | Usually ends with an approval/request phrase such as `妥否，请示。`; issuer/date stay at the end. The same no-seal signature placement applies unless a sealed template is supplied. |
| 函 | Ending depends on purpose: asking for reply, replying, requesting support, etc. Issuer/date stay at the end; no-seal drafts use the same right-side signature area. |
| 通知 | Often ends after listed matters or implementation requirements. Issuer/date stay at the end unless the source is an internal notice template with a different fixed layout. |
| 纪要 | Often uses meeting metadata and agreed items rather than a normal upward/downward ending phrase; signature/date may be absent in some internal meeting-minute templates. Preserve source structure unless the user asks for a formal issued version. |

For signature/date details, read `references/standards.md` before changing code or the default configuration.
