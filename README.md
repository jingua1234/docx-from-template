# docx-from-template

Cursor Agent Skill：根据参考 `.docx` / `.xlsx` 模板结构，结合代码库证据生成交付文档。

## 安装

复制到个人 Skills 目录：

```bash
# Windows
xcopy /E /I docx-from-template %USERPROFILE%\.cursor\skills\docx-from-template

# macOS / Linux
cp -R docx-from-template ~/.cursor/skills/docx-from-template
```

依赖：

```bash
python -m pip install python-docx openpyxl
```

## 用法

在 Cursor 对话中附上模板并说明，例如：

> `@某模板.docx 按照这个文档的格式，帮我写当前项目的产品部署文档`

Agent 会读取本 Skill，先用 `scripts/inspect_docx.py` 解析模板结构，再仿写并自检。

## 目录

- `SKILL.md` — 工作流与触发说明
- `reference.md` — 常见交付文档类型与取证要点
- `scripts/inspect_docx.py` — 模板结构检查脚本

## License

MIT
