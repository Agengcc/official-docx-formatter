# Font Policy

This skill should not automatically install fonts on the user's computer.

Reasons:

- Official-document fonts may be subject to organization, operating-system, or commercial font licensing.
- Installing fonts changes the user's system environment and may require administrator approval.
- Word/WPS rendering still depends on the fonts available on the machine that opens the document.

## Default Behavior

Use the selected profile's font list. The first font is the preferred official-document font; later fonts are acceptable fallbacks.

Default profile intent:

- Title: prefer `方正小标宋简体` or another xiao-biao-song style; fall back to `小标宋体`, `华文中宋`, `Songti SC`, or `宋体`.
- Body: prefer `仿宋_GB2312` or `仿宋`; fall back to `STFangsong`, `FangSong`, or similar fang-song fonts.
- Level 1 heading: prefer `黑体`; fall back to `SimHei`, `Heiti SC`, or similar hei-ti fonts.
- Level 2 heading: prefer `楷体_GB2312` or `楷体`; fall back to `Kaiti SC`, `KaiTi`, or similar kai-ti fonts.

## When Fonts Are Missing

For normal review drafts:

- Do not stop the formatting task only because an exact font is missing.
- Generate the Word file with the profile's preferred font names and fallback font names where supported.
- Tell the user that Word/WPS may substitute fonts visually if the machine lacks the exact font.

For strict printed or officially submitted documents:

- Ask the user for the unit's required font package or template.
- Do not install fonts automatically.
- If the user explicitly provides a font file and asks to install it, confirm licensing/permission first and treat installation as a separate system operation, not part of this formatter.

## Profile Guidance

If an organization mandates specific fonts, create a custom profile. Keep the standard profile generic and do not bake an enterprise font package into the formatter.
