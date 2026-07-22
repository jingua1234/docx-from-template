---
name: docx-from-template
description: Generates Word (.docx) or Excel (.xlsx) delivery documents by mirroring a reference template's structure (headings, tables, columns) and filling content from the current codebase. Use when the user says 按照/参考格式, attaches a .docx/.xlsx template, or asks for 交付文档, 产品部署文档, 代码提交记录,核心技术说明, Bug跟踪表, 环境审查, 二进制文件清单, 开源组件清单, or 源代码架构说明.  
---

# 按模板生成交付文档

## 何时使用

用户给出参考 `.docx` / `.xlsx`，要求「按照这个格式」写新产品部署、提交记录、国产化/测评类交付物等时，必须使用本 Skill。PDF 任务改用 `pdf` Skill。

## 依赖

```bash
python -m pip install python-docx openpyxl
```

## 工作流（必须按序）

复制清单并跟踪进度：

```
任务进度:
- [ ] 1. 解析模板结构
- [ ] 2. 按文档类型收集代码库证据
- [ ] 3. 按模板版式生成文档
- [ ] 4. 回读自检
```

### 1. 解析模板结构

先跑检查脚本，**禁止只凭文件名猜栏目**：

```bash
python scripts/inspect_docx.py "path/to/template.docx"
# Excel:
python scripts/inspect_docx.py "path/to/template.xlsx"
```

记录：标题层级与样式名、表格列头、示例行模式、列表/编号习惯。复杂结构可再读 [reference.md](reference.md)。

### 2. 收集证据

按文档类型从代码库取证（见 [reference.md](reference.md)）。缺证据处写「待确认」或留空并在回复中列出，**禁止编造**提交哈希、版本号、证书编号等。

### 3. 对齐版式并写出

- 标题层级、表头列名、章节顺序与模板一致
- Word 用 `python-docx`；Excel 用 `openpyxl`
- 默认输出到项目 `doc/`；若无该目录则用用户指定路径，或询问后再写
- 文件名：模板语义 + 项目/范围后缀（例：`代码提交记录-四项目.docx`）

### 4. 自检

对输出再跑 `inspect_docx.py`，核对：

- 章节数与标题文案大致对齐模板
- 表头列名一致、无整表空白未说明
- 无工具残留占位符（如 `TODO`、`lorem`、未替换的 `{{...}}`）

## 输出约定

- 回复中给出：输出路径、相对模板的主要差异、仍待确认的条目
- 中间提取结果可放 `doc/_tmp_*.txt`，交付完成后可删

## 反模式

- 不改模板源文件（除非用户明确要求改模板）
- 不把 PDF 流程塞进本 Skill
- 不在无 git 历史时伪造真实 commit；可按用户要求生成「合理示意记录」并标明为示意
