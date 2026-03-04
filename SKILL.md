---
name: vibe-docx
description: |
  用自然语言操作 Word 文档 - 分析、编辑、转换 DOCX 文档。
  
  触发场景：
  - 用户提到 Word 文档、DOCX 文件操作
  - 需要修复格式、整理样式、调整页面设置
  - 需要修改表格数据、替换文本内容
  - 需要合并/拆分文档、操作章节
  - 需要转换文档风格（商务报告、学术论文等）
  - 需要基于模板创建新文档
---

# vibe-docx

**核心原则：先分析，后修改。** 所有修改操作需要会话，支持回滚。

## 导入说明

```python
# 格式操作（推荐）
import vibe_docx
vibe_docx.analyze("doc.docx")
vibe_docx.begin_session("doc.docx", backup=True)

# 内容操作（需完整 API）
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")
from scripts.builder import table_read, table_update, replace_text, delete_paragraphs
```

## 工作流索引

| 工作流 | 触发关键词 | 说明 | 示例 |
|--------|-----------|------|------|
| format-fix | 格式乱、修复格式、整理格式 | 修复字体、行距、空段落 | [examples.md#format-fix](references/examples.md#format-fix) |
| content-edit | 修改内容、编辑表格、替换文本、基于模板 | 修改表格/段落内容 | [examples.md#content-edit](references/examples.md#content-edit) |
| style-convert | 变成风格、专业报告、风格转换 | 应用预置模板 | [examples.md#style-convert](references/examples.md#style-convert) |
| analyze | 分析文档、检测问题、文档诊断 | 全面检测格式问题 | [examples.md#analyze](references/examples.md#analyze) |
| table-organize | 表格样式、表格边框 | 统一表格样式 | [examples.md#table-organize](references/examples.md#table-organize) |
| document-merge | 合并文档、批量合并 | 合并多个文档 | [examples.md#document-merge](references/examples.md#document-merge) |
| textbox-extract | 文本框、提取文本框、简历分析 | 检测/提取文本框 | [examples.md#textbox-extract](references/examples.md#textbox-extract) |

## 能力矩阵

| 需求 | 工具组合 | 示例 |
|------|---------|------|
| 修改表格单元格 | `table_read` → `table_update` | [examples.md#content-edit](references/examples.md#content-edit) |
| 替换段落文本 | `read_text` → `replace_text` | [examples.md#content-edit](references/examples.md#content-edit) |
| 删除段落范围 | `delete_paragraphs` | [examples.md#content-edit](references/examples.md#content-edit) |
| 基于模板创建 | `table_update` + `replace_text` | [examples.md#content-edit](references/examples.md#content-edit) |
| 修复文档格式 | `fix_formatting` | [examples.md#format-fix](references/examples.md#format-fix) |
| 应用专业风格 | `apply_style_template` | [examples.md#style-convert](references/examples.md#style-convert) |

## API 速查

### 只读工具（无需会话）

| 工具 | 说明 |
|------|------|
| `analyze(file_path)` | 全面分析文档 |
| `detect_textboxes(file_path)` | 检测文本框 |
| `get_section_outline(file_path)` | 获取章节大纲 |

### 修改工具（需要会话）

| 工具 | 说明 |
|------|------|
| `begin_session(file_path, backup=True)` | 开始会话 |
| `commit(session_id)` | 提交修改 |
| `rollback(session_id)` | 回滚修改 |
| `fix_formatting(session_id, options)` | 修复格式 |
| `apply_style_template(session_id, template_name)` | 应用模板 |
| `table_read(session_id, table_index)` | 读取表格 |
| `table_update(session_id, table_index, cells)` | 更新表格 |
| `replace_text(session_id, paragraph_index, text)` | 替换文本 |
| `delete_paragraphs(session_id, start, end, options)` | 删除段落 |

### 批量工具

| 工具 | 说明 |
|------|------|
| `merge_documents(paths, output)` | 合并文档 |

## 预置模板

| 模板 | 适用场景 |
|------|---------|
| `business_report` | 商务报告 |
| `academic_paper` | 学术论文 |
| `internal_simple` | 内部简报 |

## 错误码

| 错误码 | 说明 |
|-------|------|
| DOC001 | 文档不存在 |
| DOC002 | 不支持的格式 |
| SES001 | 会话无效 |
| SES002 | 会话已过期 |
| SES003 | 会话冲突 |

## 详细参考

| 文档 | 内容 |
|------|------|
| [examples.md](references/examples.md) | 完整代码示例 |
| [api.md](references/api.md) | API 参数详解 |
| [workflows.md](references/workflows.md) | 工作流定义 |
| [templates.md](references/templates.md) | 模板配置 |
| [intent-mapping.csv](references/intent-mapping.csv) | 意图映射 |

## 依赖

- Python >= 3.10
- python-docx >= 1.1.0