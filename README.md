# gongwen-format-skill

`gongwen-format-skill` 是一个面向 **Claude Code、Codex等AI Agent的 Skills/自动化场景** 的“确定性公文（公文格式）排版工具”：输入 **受控 Markdown**（含 YAML Front Matter）或 **JSON**，通过 `python-docx` 稳定生成符合常见公文排版约束的 Word `.docx`（字体/字号/行距/标题与层级/主送机关/附件/页码等）。

## 功能特性

- 受控输入协议：不做通用 Markdown 语义推断，靠“行首标记”做确定性渲染
- 标题与层级：`#`~`#####` 映射到公文标题/一二三级标题的固定字体与字号
- 正文排版：段前段后 0、固定行距、首行缩进（2 字）等
- 主送机关、附件说明、落款与日期（来自 Front Matter）
- 页码：页脚居中 `— 1 —`（Word 字段 `PAGE`）
- 依赖轻：核心仅依赖 `python-docx`，Front Matter 解析内置实现（不强依赖 PyYAML）

## 目录结构（建议开源仓库内保持一致）

> 典型布局如下（以 Skill 目录为根）：

```
gongwen-format-skill/
  SKILL.md
  scripts/
    gongwen_doc.py
  references/
    受控 Markdown 公文解析与渲染规范（v1.0）.md
    公文格式要求.md
  assets/
    （可选）字体文件 .ttf
```

## 安装到 Codex Skills（举例，其他AI Agent类似）

1. 将本仓库中的 `gongwen-format-skill/` 目录复制到你的 Codex Skills 目录下（例如 `~/.codex/skills/`）。
2. 确保 `SKILL.md` 位于 `gongwen-format-skill/` 目录根部。
3. 重启 Codex CLI（或让其重新加载 skills）。

> 你的环境里具体的 Skills 根目录可能不同；以你当前 Codex 配置为准。

## 作为脚本直接使用

### 依赖

- Python 3.9+
- `python-docx`

安装依赖：

```bash
pip install python-docx
```

### 生成 docx

JSON 输入：

```bash
python scripts/gongwen_doc.py --input data.json -o 输出公文.docx
```

受控 Markdown 输入：

```bash
python scripts/gongwen_doc.py --md input.md -o 输出公文.docx
```

也支持从标准输入读取（将 `--md -` 或 `--input -`）：

```bash
cat input.md | python scripts/gongwen_doc.py --md - -o 输出公文.docx
```

## 受控 Markdown（协议要点）

完整规范请阅读 `references/受控 Markdown 公文解析与渲染规范（v1.0）.md`。核心规则简述：

- **自然换行即自然段**：每行非空文本 = 一个 Word 段落
- **只识别行首 `#`**：根据 `#` 的数量判断段落类型
- **不做通用 Markdown**：列表、加粗、表格、引用等语法不解析，尽量不要输出相关符号
- **编号由你负责写在文本里**：脚本不补全、不校验编号连续性（如“一、”“（一）”“1.”“（1）”）

### Front Matter（可选）

文件开头可写 YAML Front Matter（仅支持受控子集）：

```markdown
---
recipients: 各相关单位
signer: XX单位
date: 2026年1月30日
attachments:
  - 年度工作总结模板
---
```

- `recipients`：主送机关（字符串或字符串数组）
- `signer`：落款单位
- `date`：成文日期（不做格式解析）
- `attachments`：附件列表

### 标题层级映射

- `# ...`：公文标题（居中，小标宋，2 号）
- `## ...`：一级标题（黑体，3 号）
- `### ...`：二级标题（楷体，3 号）
- `#### ...` / `##### ...`：更深层级（仿宋，3 号）
- 其他非空行：正文（仿宋，3 号，首行缩进）

## 示例

`input.md`：

```markdown
---
recipients: 各相关单位
signer: XX单位
date: 2026年1月30日
attachments:
  - 年度工作总结模板
---
# 关于开展年度工作总结的通知
为全面总结年度工作成果，梳理经验做法，现就有关事项通知如下。
## 一、总体要求
### （一）突出重点。
这是第一自然段。
这是第二自然段。
```

## 字体与合规提示（重要）

本项目通常会使用以下字体名进行渲染（由脚本常量控制），以满足常见公文排版习惯：

- `方正小标宋简体`（标题）
- `仿宋_GB2312`（正文）
- `楷体_GB2312`（二级标题）
- `黑体`（一级标题）

注意：

- **字体文件可能受版权/授权约束**。如你计划开源发布仓库，建议：
  - 不直接随仓库分发字体文件；或
  - 明确字体来源与授权条款，并在仓库中提供合规的获取方式。
- 若目标机器未安装对应字体，Word 的显示效果可能会回退到其他字体。

## 设计边界

- 面向“受控输入 + 固定版式”的公文生成；不是通用 Markdown 转 docx 工具
- 不做语义推断、自动编号、格式纠错与一致性校验

## 贡献与维护建议

- 将输入协议的变更放到 `references/` 并做版本号管理（例如 v1.1、v1.2）
- 对外开放时建议提供：
  - `LICENSE`（例如 MIT/Apache-2.0）
  - `CHANGELOG`（或 Releases 记录）
  - 最小可复现示例（`examples/`）

