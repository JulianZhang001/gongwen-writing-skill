# 公文技能仓库

这是一个面向 Codex 的公文技能仓库，包含两个方向：

- `official-document-writing-polisher`：公文写作润色
- `docx-official-format`：Word 公文印制格式排版

适合正式公文、讲话稿、汇报稿、通知、总结、校园新闻稿，以及 Word 公文版式排版与规范化处理。

## 仓库内容

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

## 适用场景

- 正式公文、办公材料、送审稿
- 讲话稿、汇报稿、通知、方案、总结、纪要
- 校园新闻稿、正式宣传稿
- AI 初稿去模板味、去空话套话、提逻辑层次
- Word 公文格式排版、附件说明、落款、页码和版式规范

## 写作前建议提供的材料

为了让成稿更扎实，建议尽量提供：

- 文种和用途
- 背景、时间、地点
- 参会对象或发文对象
- 会议材料、通知原文、提纲、领导口径
- 必须保留或不能写的内容

如果材料不全，`official-document-writing-polisher` 会先追问关键项；如果需要先出一版，会标明基于现有信息的假设。

## 不适用场景

- 虚构事实、补政策依据、补领导表态
- 营销文案、社交媒体文案、轻口语内容
- 不需要正式文本语体的随意写作

## 安装方式

如果你的 Codex 环境支持通过仓库地址安装 skill，可直接使用：

- `https://gitee.com/julaoshi/gongwen-writing-skill`

如果你是手动安装，请把整个目录放到本地 skills 目录，并保持目录结构不变。

## 使用示例

```text
Use $official-document-writing-polisher 帮我把这份新闻稿改成正式公文风格，保留原意，不新增事实。
```

```text
Use $docx-official-format 帮我把这份通知排成公文格式，按本地印制要求处理页边距、字体、附件、落款和页码。
```

如果你想先看前后对比示例，可以直接查看 [`examples/`](./examples/README.md)。

如果你想直接复制安装和调用口令，可以查看 [`安装与调用示例.md`](./安装与调用示例.md)。

如果你只想检查公文印制格式子技能是否完整，可以运行 [`docx-official-format/scripts/validate_skill.py`](./docx-official-format/scripts/validate_skill.py)。

## 使用原则

- 保留原意
- 不新增事实
- 先调结构，再改句子
- 文风稳、准、简、可送审

## License

MIT
