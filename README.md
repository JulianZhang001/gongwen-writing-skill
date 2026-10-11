# 公文写作 Skill

> 给 AI 装上写公文的手艺——从**润色**到**排版**，一条链路走完。

面向中文公文场景的 Codex / Claude Code / WorkBuddy Skill 集合。装进 AI 之后，
它能按机关公文的规范帮你改稿、排版，而不是只把话说得更漂亮。

[English](./README.en.md) · [安装与调用示例](./安装与调用示例.md) · [示例对照](./examples/README.md)

## 包含两个 Skill

| Skill | 做什么 | 解决什么问题 |
| --- | --- | --- |
| **`official-document-writing-polisher`** | 公文写作润色 | 口语化、AI 味、空话套话、层次不清、术语不准 |
| **`docx-official-format`** | Word 公文印制格式排版 | 页边距、行距、字体、附件说明、落款、页码全自动 |

两个可以单独用，也可以串起来用：**先润色文字，再排成规范版式。**

## 它能做到什么

**写作润色**——对标权威正式文稿的共同特征（立场稳、判断准、措辞正、句法简、节奏稳、层次清）：

- 删掉「在……背景下」「进一步推动……走深走实」这类无信息增量的空转句
- 把「进行研究部署」「开展推进落实」这类虚词堆叠改成直接动词
- 按「背景 → 目标 → 措施 → 要求」重排结构，而不是只改字
- 统一机构名称、政策表述、时间数字与标点写法
- 严格守住「不新增事实、不改变原意」的底线

**格式排版**——按《文件印制格式基本要求》自动处理：

- 页面：A4，上 35mm / 下 29mm / 左右 25.5mm，页眉 15mm / 页脚 22mm
- 行距：固定 30 磅
- 标题：方正小标宋_GBK 二号；正文：方正仿宋_GBK 三号
- 一级标题「一、」方正黑体_GBK；二级「（一）」方正楷体_GBK；三级「1.」仿宋加粗
- 附件说明、落款与日期对齐、页码（`— 2 —`，首页无码）
- 中英混排：西文与数字用 Times New Roman，中文用方正系列，自动分槽位

## 快速开始

### 安装

把整个 `gongwen-writing-skill` 目录放进本地 skills 目录，保持结构不变：

```bash
# WorkBuddy 用户级
cp -r gongwen-writing-skill ~/.workbuddy/skills/

# 或项目级
cp -r gongwen-writing-skill .workbuddy/skills/
```

仓库地址：`https://github.com/JulianZhang001/gongwen-writing-skill`

### 调用

```text
Use $official-document-writing-polisher 帮我把这份新闻稿改成正式公文风格，保留原意，不新增事实。
```

```text
Use $docx-official-format 帮我把这份通知排成公文格式，按印制要求处理页边距、字体、附件、落款和页码。
```

### 命令行排版

```bash
python3 docx-official-format/scripts/format_docx.py 输入.docx 输出_公文格式.docx
```

支持 `.doc` 旧格式输入，也支持自定义字体：

```bash
python3 docx-official-format/scripts/format_docx.py 输入.docx 输出.docx \
  --title-font 方正小标宋_GBK --body-font 方正仿宋_GBK \
  --heading-font 方正黑体_GBK --secondary-heading-font 方正楷体_GBK
```

## 适用场景

- 通知、请示、报告、函、方案、总结、汇报稿、讲话稿、纪要
- 校园新闻稿与正式宣传稿
- 把口语化初稿、AI 草稿改成可送审、可上会的版本
- Word 公文的版式规范化与格式核查

**不适用**：需要虚构事实或政策依据、营销文案、社交媒体文案、
依赖法条逐字准确或学术引用格式的专门文本。

## 写作前建议提供

材料给得全，成稿就扎实。建议尽量提供：

- **文种和用途** —— 讲话稿？通知？报告？方案？
- **场景背景** —— 什么会议、什么活动、为什么写
- **时间地点** —— 日期、地点、线上线下
- **参与对象** —— 面向谁、谁发言、谁主办
- **核心事实** —— 已确定的背景、数据、成果、问题、部署
- **口径要求** —— 必须保留的表述、不能写的内容、领导偏好说法
- **参考材料** —— 已有提纲、会议纪要、通知原文、历史稿件

材料不全时，Skill 会先追问关键项；如果你要求先出一版，
它会明确标出哪些内容是基于现有材料的推定。

## 目录结构

```text
gongwen-writing-skill/
├── README.md                    中文说明（本文件）
├── README.en.md                 英文说明
├── SKILL.md                     写作润色 Skill 主文件
├── 安装与调用示例.md             装机与调用速查
├── LICENSE                      MIT
├── docx-official-format/        Word 排版子技能
│   ├── SKILL.md
│   ├── references/              格式规范详解
│   ├── scripts/                 排版与校验脚本
│   ├── agents/                  平台配置
│   └── tests/                   测试
├── examples/                    前后对照示例
│   ├── 新闻稿示例.md
│   ├── 讲话稿示例.md
│   ├── 通知示例.md
│   └── 校园新闻稿完整示例.md
└── references/                  审校清单
    ├── 公文写作审校清单.md
    └── 写作前信息清单.md
```

## 四条硬规则

1. **文风对标规范** —— 正式、稳健、克制，能写实的不写虚，能写短的不写长
2. **语句打磨规范** —— 主谓宾清楚，多用短句，删掉 AI 套句与虚词堆叠
3. **逻辑架构规范** —— 先判断文种对象目的，再决定结构；结论要求前置
4. **内容格调规范** —— 删空话套话重复，统一术语口径，提升信息密度

## 使用原则

- 保留原意，不新增事实
- 先调结构，再改句子
- 文风稳、准、简、可送审

## 相关项目

同一套公文工作流里的另外三个工具：

- [公文自动排版工具](https://github.com/JulianZhang001/official-doc-formatter) —— 国标 GB/T 9704 自动校排
- [公文空白模板生成](https://github.com/JulianZhang001/gongwen-blank-template) —— 一键生成规范空白模板
- [国标公文增强排版 Skill](https://github.com/JulianZhang001/gb-official-doc-enhanced)

## License

MIT © 2026 居老师讲AI
