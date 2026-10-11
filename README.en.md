# Official Document Writing Skill

> Give your AI the craft of Chinese official-document writing — polish and format in one pipeline.

A set of Codex / Claude Code / WorkBuddy skills for Chinese official-document (公文) workflows.
Once installed, your AI drafts and formats according to government-document conventions —
instead of merely making the prose sound nicer.

[中文说明](./README.md) · [Install & Usage](./安装与调用示例.md) · [Examples](./examples/README.md)

## Two Skills Included

| Skill | What it does | Problem it solves |
| --- | --- | --- |
| **`official-document-writing-polisher`** | Official-document polishing | Colloquialisms, AI-flavored phrasing, empty slogans, unclear structure, inconsistent terminology |
| **`docx-official-format`** | Word printing-format conversion | Margins, line spacing, fonts, attachment notes, issuer signature, page numbers — all automated |

Use either on its own, or chain them: **polish the text first, then apply the official format.**

## Key Capabilities

**Writing polish** — benchmarked against the shared traits of authoritative formal Chinese writing
(stable stance, precise judgment, correct wording, simple syntax, steady rhythm, clear hierarchy):

- Removes zero-information filler such as "在……背景下" or "进一步推动……走深走实"
- Turns verb-stacking like "进行研究部署" into direct verbs
- Rebuilds structure as background → objective → measures → requirements, not just word swaps
- Unifies organization names, policy wording, dates, numbers and punctuation
- Strictly preserves meaning: **no new facts, no altered intent**

**Format conversion** — applies the local "文件印制格式基本要求" standard:

- Page: A4, margins top 35mm / bottom 29mm / left-right 25.5mm; header 15mm / footer 22mm
- Paragraph: fixed 30pt line spacing
- Title: 方正小标宋_GBK 22pt; Body: 方正仿宋_GBK 16pt
- Level-1 heading "一、": 方正黑体_GBK; Level-2 "（一）": 方正楷体_GBK; Level-3 "1.": bold 仿宋
- Attachment notes, right-aligned issuer/date pair, page numbers (`— 2 —`, none on first page)
- Mixed script: Latin letters and digits in Times New Roman, CJK in the 方正 family, split across OOXML font slots

## Quick Start

### Install

Place the whole `gongwen-writing-skill` directory into your local skills folder, keeping the structure intact:

```bash
# WorkBuddy user level
cp -r gongwen-writing-skill ~/.workbuddy/skills/

# or project level
cp -r gongwen-writing-skill .workbuddy/skills/
```

Repository: `https://github.com/JulianZhang001/gongwen-writing-skill`

### Invoke

```text
Use $official-document-writing-polisher to polish this draft into a formal Chinese public document while preserving meaning.
```

```text
Use $docx-official-format to format this Word file according to the local Chinese official-document printing requirements.
```

### Command-line formatting

```bash
python3 docx-official-format/scripts/format_docx.py input.docx output_公文格式.docx
```

Legacy `.doc` input is also supported, and fonts can be overridden:

```bash
python3 docx-official-format/scripts/format_docx.py input.docx output.docx \
  --title-font 方正小标宋_GBK --body-font 方正仿宋_GBK \
  --heading-font 方正黑体_GBK --secondary-heading-font 方正楷体_GBK
```

## When to Use

- Notices, requests for instructions, reports, letters, plans, summaries, speeches, minutes
- Campus news releases and formal publicity drafts
- Turning colloquial drafts or AI-generated text into submission-ready versions
- Normalizing and auditing Word official-document layout

**Not suitable for**: content requiring fabricated facts or policy citations, marketing copy,
social-media copy, or texts governed by strict legal wording or academic citation formats.

## Recommended Inputs

Better input produces a stronger draft. Try to provide:

- **Document type and purpose** — speech? notice? report? plan?
- **Context** — what meeting, what event, why it is being written
- **Time and place** — date, venue, online or offline
- **Audience** — who it addresses, who speaks, who hosts
- **Core facts** — confirmed background, figures, outcomes, issues, deployments
- **Wording constraints** — phrasing to keep, content to avoid, leadership preferences
- **Reference material** — existing outlines, minutes, original notices, past drafts

If key facts are missing, the writing skill asks for the most important items first.
If you need an immediate first draft, it marks which parts are inferred from available material.

## Repository Layout

```text
gongwen-writing-skill/
├── README.md                    Chinese guide (primary)
├── README.en.md                 This file
├── SKILL.md                     Writing-polish skill
├── 安装与调用示例.md             Install & invoke quick reference
├── LICENSE                      MIT
├── docx-official-format/        Word formatting sub-skill
│   ├── SKILL.md
│   ├── references/              Detailed format rules
│   ├── scripts/                 Formatting and validation scripts
│   ├── agents/                  Platform config
│   └── tests/                   Tests
├── examples/                    Before/after examples
└── references/                  Review checklists
```

## The Four Hard Rules

1. **Style benchmark** — formal, steady, restrained; prefer the concrete over the vague, the short over the long
2. **Sentence polish** — clear subject-verb-object, short clauses, no AI filler or verb stacking
3. **Structure** — identify document type, audience and purpose first; front-load conclusions and requirements
4. **Tone** — cut empty slogans and repetition, unify terminology, raise information density

## Principles

- Preserve meaning, add no facts
- Fix structure before polishing sentences
- Steady, precise, concise, submission-ready

## Related Projects

Other tools in the same official-document pipeline:

- [Official Document Auto-Formatter](https://github.com/JulianZhang001/official-doc-formatter) — GB/T 9704 automatic formatting
- [Blank Official Template Generator](https://github.com/JulianZhang001/gongwen-blank-template)
- [GB-Standard Enhanced Formatting Skill](https://github.com/JulianZhang001/gb-official-doc-enhanced)

## License

MIT © 2026 居老师讲AI
