# Profile Guide

Profiles live in `profiles/*.json`.

## Rule

`standard-party-government.json` is the base profile. Every custom organization profile should inherit from it unless the user explicitly needs a fully independent format.

A profile may override formatting only. It must not enable body rewriting, content completion, desensitization, or automatic addition of missing issuer/signature/date content.

## Minimal Custom Profile

Ask only for known differences. Missing values inherit from the base profile.

```json
{
  "profile_id": "my-company",
  "display_name": "某单位配置",
  "inherits": "standard-party-government",
  "fonts": {
    "title": {
      "fallbacks": ["方正小标宋简体", "小标宋体", "宋体"]
    },
    "body": {
      "fallbacks": ["仿宋_GB2312", "仿宋"]
    }
  }
}
```

## First-Use Questions

Use this wording:

> 请选择格式配置：  
> 1. 党政机关公文标准配置：参考 GB/T 9704-2012，推荐用于通用正式公文  
> 2. 自定义单位配置：使用你所在单位指定的字体、行距、页边距等要求

If custom:

- Ask profile name.
- Ask title font only if the user knows it.
- Ask body font only if the user knows it.
- Ask level-1 and level-2 heading fonts only if the user mentions a unit rule.
- Ask line spacing and margins only if the user says their unit differs.

Do not pressure the user to fill every field. Unknown is fine.
