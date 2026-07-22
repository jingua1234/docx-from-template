# docx-from-template

通用 Agent Skill：根据参考 `.docx` / `.xlsx` 模板结构，结合代码库证据生成交付文档。适用于 **Claude Code**、**Codex**、**Cursor** 等支持 Agent Skills（`SKILL.md`）的工具。

## 功能

- 解析模板标题层级、表格列头与版式
- 按文档类型从代码库取证（部署、提交记录、架构说明、Bug 表等）
- 对齐模板栏目生成 Word / Excel，缺证据处标注「待确认」而非编造

## 依赖

```bash
python -m pip install python-docx openpyxl
```

## 安装

将本仓库克隆或复制到对应工具的 Skills 目录（目录名保持 `docx-from-template`）：

| 工具 | 个人 Skills 目录 | 项目级（可选） |
|------|------------------|----------------|
| Claude Code | `~/.claude/skills/docx-from-template/` | `.claude/skills/docx-from-template/` |
| Codex | `~/.codex/skills/docx-from-template/` | `.codex/skills/docx-from-template/` |
| Cursor | `~/.cursor/skills/docx-from-template/` | `.cursor/skills/docx-from-template/` |

```bash
# 示例：安装到 Claude Code（macOS / Linux）
git clone https://github.com/jingua1234/docx-from-template.git ~/.claude/skills/docx-from-template

# 示例：安装到 Codex
git clone https://github.com/jingua1234/docx-from-template.git ~/.codex/skills/docx-from-template

# 示例：安装到 Cursor
git clone https://github.com/jingua1234/docx-from-template.git ~/.cursor/skills/docx-from-template
```

Windows（PowerShell）示例：

```powershell
git clone https://github.com/jingua1234/docx-from-template.git "$env:USERPROFILE\.claude\skills\docx-from-template"
# 或 Codex / Cursor：把路径中的 .claude 换成 .codex / .cursor
```

也可只复制 `SKILL.md`、`reference.md`、`scripts/` 到上述目录。

## 用法

在对话中附上模板并说明，例如：

> `@某模板.docx 按照这个文档的格式，帮我写当前项目的产品部署文档`

Agent 读取本 Skill 后，会先用 `scripts/inspect_docx.py` 解析模板结构，再仿写并自检。

手动检查模板：

```bash
python scripts/inspect_docx.py path/to/template.docx
python scripts/inspect_docx.py path/to/template.xlsx
```

## 目录

- `SKILL.md` — 工作流与触发说明（各工具共用的 Skill 入口）
- `reference.md` — 常见交付文档类型与取证要点
- `scripts/inspect_docx.py` — 模板结构检查脚本

## License

MIT
