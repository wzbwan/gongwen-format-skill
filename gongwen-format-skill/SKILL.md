---
name: gongwen-format-skill
description: Generate Chinese official document (公文) Word files from controlled Markdown or JSON. Use when producing or updating .docx 公文 that must follow fixed fonts, line spacing, title/recipient layout, attachments, and footer page numbers, especially when the input is a controlled Markdown protocol.
---

# Gongwen Markdown Docx

## Overview

Create standard-format 公文 .docx with a deterministic script. It supports controlled Markdown (with front matter) or JSON input and uses bundled fonts for consistent rendering.

## Quick Start

Use JSON input or controlled Markdown to generate a docx.

```bash
python scripts/gongwen_doc.py --input data.json -o 输出公文.docx
python scripts/gongwen_doc.py --md input.md -o 输出公文.docx
```

## Controlled Markdown

- Read `references/受控 Markdown 公文解析与渲染规范（v1.0）.md` before generating Markdown input.
- Read `references/公文格式要求.md` for font sizes, line spacing, and layout constraints.
- Front matter supports: `recipients`, `signer`, `date`, `attachments`.
- ASCII double quotes in content are normalized to Chinese quotes (“…”).

Example (controlled Markdown):

```markdown
---
recipients: 各相关单位
signer: XX单位
date: 2026年1月30日
attachments:
  - 年度工作总结模板
---
# 关于开展年度工作总结的通知
## 一、总体要求
### （一）突出重点。
这是第一自然段。
这是第二自然段。
```

## Fonts

Use fonts in `assets/` for consistent rendering. Install them on the OS or ensure Word can locate them.
