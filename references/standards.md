# Standards Basis

Use the default profile as a practical Word/WPS implementation of these sources:

- `GB/T 9704-2012《党政机关公文格式》`: the primary format basis.
- `《党政机关公文处理工作条例》`: document types, direction, and handling principles.
- `GB/T 33476.2-2016《党政机关电子公文格式规范 第2部分：显现》`: electronic document display consistency.

## Source Pointers

- National standards platform: `GB/T 9704-2012《党政机关公文格式》` is current and continued effective after the 2025-05-30 review.
- `《党政机关公文处理工作条例》` defines the 15 official document types and says official-document layout follows the national standard.
- The legacy HN formatter covers the high-frequency enterprise subset: 请示、报告、决定、通知、批复、通报、函、纪要.

Use these sources in that order: national rule first, then enterprise/high-frequency practice, then user-provided organization profile.

The formatter is an implementation aid, not an authoritative source. If a unit template or profile conflicts with the default formatter behavior, the unit template/profile takes priority.

## Hard Standards

Treat these as the base rules unless the user names an organization profile:

- A4 paper.
- Page margins: top 37 mm, bottom 35 mm, left 28 mm, right 26 mm.
- Body area target: 22 lines per page, 28 characters per line.
- Main text generally uses 3 hao fang-song style.
- Title generally uses 2 hao xiao-biao-song style.
- Hierarchy order: `一、`, `（一）`, `1.`, `（1）`.
- The official 15 document types are: 决议、决定、命令、公报、公告、通告、意见、通知、通报、报告、请示、批复、议案、函、纪要.
- For existing documents, classify the document type before applying formatting. The document type controls structure; the profile controls typography and page style.

## Signature, Date, and Spacing

This skill treats the common enterprise Word output as a no-seal single-issuer draft unless the user provides a red-head/seal template or an organization profile says otherwise.

For a no-seal single-issuer document, use this default:

- Put the issuer/signature area after the body or attachment note, not at the top of the document.
- Leave one blank line after the body or attachment note before the issuer.
- Put the issuer on the right side with about two Chinese characters of right space.
- Put the date on the next line under the issuer, also on the right side.
- Write the date with Arabic numerals in the form `2026年6月19日`; do not write `2026年06月09日`.

For a sealed single-issuer document, the national format is more stamp-aware:

- The date is generally arranged with four Chinese characters of right space.
- The issuer is placed above the date and centered relative to the date.
- The red seal should press down over the issuer/date area, and the seal top should stay within one line from the body or attachment note.

For a signed-stamp document, the standard has a separate layout: the signature stamp is generally placed two blank lines below the body or attachment note, and the date is placed one blank line below the signature stamp with four-character right space.

Because this skill does not insert or position real seals by default, the profile implements the no-seal single-issuer rule. If a unit uses sealed templates, represent that as an organization profile or a reusable `.docx` template rather than guessing from plain text.

## Word/WPS Implementation Choices

These are practical approximations, not separate legal standards:

- Use 22 pt for 2 hao title text.
- Use 16 pt for 3 hao body text.
- Use fixed 28 pt line spacing by default to approximate 22 lines per page.
- Use two-character first-line indent for body paragraphs.
- Use font fallback lists because systems may not have the exact same Chinese fonts installed.

## Enterprise Extensions

Large organizations often keep the national standard as the base and add internal habits:

- Designated font files.
- Slightly different line spacing due to Word/WPS templates.
- Red-head or letterhead templates.
- Internal forms such as 签报、呈批件、工作联系单.
- Approval, issuing, archive, or OA metadata requirements.

Represent those differences as profiles. Do not bake an enterprise name into the formatter code.
