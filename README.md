# Gongwen Writing Skill

This repository contains two Codex skills for Chinese official-document work:

- `official-document-writing-polisher`: text-level polishing for official documents and formal public-sector writing
- `docx-official-format`: Word `.doc` / `.docx` official-document printing format conversion

## What It Covers

It supports formal notices, reports, speeches, summaries, campus/government-style drafts, and Word layout conversion for official-document printing requirements.

## Recommended Inputs

For stronger drafts, provide the document type, purpose, background, time, location, audience, meeting materials, leadership talking points, required wording, and forbidden content. If key facts are missing, the writing skill should ask for the most important missing items before drafting, or mark assumptions when the user asks for a first draft immediately.

## Included Files

```text
gongwen-writing-skill/
├── README.md
├── README.zh-CN.md
├── LICENSE
├── 安装与调用示例.md
├── SKILL.md
├── docx-official-format/
│   ├── SKILL.md
│   ├── agents/
│   ├── references/
│   ├── scripts/
│   └── tests/
├── examples/
│   ├── README.md
│   ├── 新闻稿示例.md
│   ├── 讲话稿示例.md
│   ├── 通知示例.md
│   └── 校园新闻稿完整示例.md
└── references/
    └── 公文写作审校清单.md
```

## Install

If your Codex environment supports skill installation by repo URL, install from:

- `https://github.com/JulianZhang001/gongwen-writing-skill`

## Usage

```text
Use $official-document-writing-polisher to polish this draft into a formal Chinese public document while preserving meaning.
```

```text
Use $docx-official-format to format this Word file according to the local Chinese official-document printing requirements.
```
