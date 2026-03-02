---
name: vibe-docx
description: |
  用自然语言操作 Word 文档 - 分析、编辑、转换 DOCX 文档。
  
  触发场景：
  - 用户提到 Word 文档、DOCX 文件操作
  - 需要修复格式、整理样式、调整页面设置
  - 需要合并/拆分文档、操作章节
  - 需要处理表格、文本框、图片
  - 需要转换文档风格（商务报告、学术论文等）
  - 需要在文档中插入 Markdown 内容
---

# vibe-docx - Word 文档处理

用自然语言操作 Word 文档，支持分析、编辑、格式转换、批量处理。

## 快速开始

```python
import vibe_docx

# 1. 分析文档（只读）
result = vibe_docx.analyze("document.docx")
print(result["document_info"])

# 2. 开始编辑会话
session = vibe_docx.begin_session("document.docx", backup=True)
session_id = session["session_id"]

# 3. 执行修改
vibe_docx.fix_formatting(session_id, options={"default_font": "宋体"})

# 4. 提交或回滚
vibe_docx.commit(session_id)  # 或 vibe_docx.rollback(session_id)
```

## 核心工作流

### 1. 格式修复 (format-fix)

修复标题样式、段落格式、空段落、行距等问题。

```
触发：格式乱、修复格式、整理格式、标题样式
```

```python
session = vibe_docx.begin_session("report.docx", backup=True)

result = vibe_docx.fix_formatting(session["session_id"], options={
    "default_font": "宋体",        # 正文字体
    "title_font": "黑体",          # 标题字体
    "remove_empty_paragraphs": True,  # 移除空段落
    "convert_markdown": True,      # 转换 Markdown 语法
    "line_spacing": 1.5,           # 统一行距
    "first_line_indent": 2,        # 首行缩进（字符数）
})

print(f"修改统计: {result['changes']}")
vibe_docx.commit(session["session_id"])
```

### 2. 文档分析 (analyze)

全面分析文档问题，检测格式、结构、内容问题。

```
触发：分析文档、检测问题、文档诊断
```

```python
result = vibe_docx.analyze("report.docx")

# 检测结果
print(result["document_info"])  # 段落数、表格数、图片数、文本框数
print(result["issues"])         # 发现的问题列表

# 检测能力：
# - 标题级别跳跃 (headings)
# - 表格边框缺失 (tables)
# - 空段落过多 (empty_paragraphs)
# - Markdown 未转换 (markdown)
# - 文本框检测 (textboxes)
# - 字体不一致 (fonts)
# - 页边距异常 (margins)
```

### 3. 风格转换 (style-convert)

将文档转换为专业风格：商务报告、简洁风格、学术论文。

```
触发：变成风格、专业报告、风格转换、应用模板
```

```python
session = vibe_docx.begin_session("report.docx", backup=True)

# 方式一：应用预置模板（推荐）
result = vibe_docx.apply_style_template(
    session["session_id"],
    template_name="business_report"  # 或 academic_paper, internal_simple
)
print(f"应用模板: {result['template']}")
print(f"修改内容: {result['changes']}")

# 方式二：手动设置
vibe_docx.fix_page_setup(session["session_id"], margins={
    "top": "2.54cm", "bottom": "2.54cm", 
    "left": "3.17cm", "right": "3.17cm"
})
vibe_docx.fix_formatting(session["session_id"], options={
    "default_font": "宋体",
    "title_font": "黑体",
    "line_spacing": 1.5
})

# 方式三：模板 + 自定义覆盖
vibe_docx.apply_style_template(
    session["session_id"],
    template_name="academic_paper",
    custom_options={"default_font": "楷体", "line_spacing": 2.0}
)

vibe_docx.commit(session["session_id"])
```

### 4. 表格整理 (table-organize)

统一表格样式，添加边框。

```
触发：表格样式、表格整理、表格边框、添加边框
```

```python
session = vibe_docx.begin_session("report.docx", backup=True)
vibe_docx.fix_table_borders(session["session_id"], border_style="single")
vibe_docx.commit(session["session_id"])
```

### 5. 章节操作 (section-operate)

提取、删除、移动章节。

```
触发：章节大纲、拆分章节、提取章节、删除章节
```

```python
session = vibe_docx.begin_session("report.docx", backup=True)
vibe_docx.remove_section(session["session_id"], "附录")
vibe_docx.add_section(session["session_id"], "新章节", "章节内容")
vibe_docx.commit(session["session_id"])
```

### 6. 文档合并 (document-merge)

合并多个文档。

```
触发：合并文档、批量合并、多个文档合并
```

```python
vibe_docx.merge_documents(
    ["part1.docx", "part2.docx", "part3.docx"],
    "merged.docx",
    options={"add_page_break": True}
)
```

### 7. 文本框处理 (textbox-extract)

检测、提取、转换文本框内容。

```
触发：文本框、提取文本框、简历分析
```

```python
# 检测文本框
textboxes = vibe_docx.detect_textboxes("resume.docx")
print(f"发现 {textboxes['stats']['total_count']} 个文本框")

# 提取内容
session = vibe_docx.begin_session("resume.docx", backup=True)
vibe_docx.extract_textbox_content(session["session_id"])
vibe_docx.commit(session["session_id"])
```

## 工具分类

### 只读工具（无需会话）

| 工具 | 说明 |
|------|------|
| `analyze(file_path)` | 全面分析文档，检测格式/结构/内容问题 |
| `detect_textboxes(file_path)` | 检测文本框 |
| `get_section_outline(file_path)` | 获取章节大纲 |
| `get_document_structure(file_path)` | 获取文档结构 |
| `validate_xml(doc_path)` | XML 底层验证 |

### 修改工具（需要会话）

| 工具 | 说明 |
|------|------|
| `begin_session(file_path, backup=True)` | 开始会话，返回 session_id |
| `commit(session_id)` | 提交修改，清理备份 |
| `rollback(session_id)` | 回滚修改，恢复原文档 |
| `fix_formatting(session_id, options)` | 修复格式（字体、行距、空段落等） |
| `fix_page_setup(session_id, margins, orientation)` | 页面设置 |
| `fix_table_borders(session_id, table_indices, style)` | 表格边框 |
| `fix_list_formatting(session_id, options)` | 列表格式修复 |
| `apply_style_template(session_id, template_name, custom_options)` | 应用预置样式模板 |
| `add_section(session_id, title, content)` | 添加章节 |
| `remove_section(session_id, section_title)` | 删除章节 |
| `extract_textbox_content(session_id)` | 提取文本框 |
| `image_list(session_id)` / `image_insert(session_id, path)` | 图片操作 |
| `table_list(session_id)` / `table_update(session_id, ...)` | 表格操作 |

### 批量工具（无需会话）

| 工具 | 说明 |
|------|------|
| `merge_documents(file_paths, output_path)` | 合并文档 |
| `split_document(file_path, split_points, output_dir)` | 拆分文档 |

## 会话管理

会话 ID 格式：`ses_{uuid_hex}`（如 `ses_ff6df5a2b3c4`）

会话特性：
- 自动备份（存放在 `.vibe-backups/` 目录）
- 文件锁（防止同一文件被多个会话修改）
- 自动过期（1 小时无活动自动清理）

```python
# 完整会话流程
session = vibe_docx.begin_session("document.docx", backup=True)

try:
    vibe_docx.fix_formatting(session["session_id"])
    vibe_docx.fix_table_borders(session["session_id"])
    vibe_docx.commit(session["session_id"])
except Exception as e:
    vibe_docx.rollback(session["session_id"])
    print(f"操作失败，已回滚: {e}")
```

## 错误码

| 错误码 | 说明 | 恢复建议 |
|-------|------|---------|
| DOC001 | 文档不存在 | 请提供有效的文档路径 |
| DOC002 | 不支持的文档格式 | 仅支持 .docx 格式 |
| DOC003 | 文档已损坏 | 尝试修复或使用备份 |
| DOC004 | 文档被锁定 | 请关闭其他程序后重试 |
| SES001 | 会话无效 | 请创建新会话 |
| SES002 | 会话已过期 | 会话超过 1 小时未活动 |
| SES003 | 会话冲突 | 文档正被其他会话使用 |
| SES004 | 备份失败 | 检查磁盘空间和权限 |
| VAL001 | 分析失败 | 检查文件是否损坏 |
| BLD001 | 执行失败 | 检查操作参数 |

## 响应格式

所有工具返回统一格式：

```python
# 成功
{"success": True, "data": {...}, "changes": {...}}

# 失败
{
    "success": False,
    "error": {
        "code": "DOC001",
        "message": "文档不存在",
        "detail": "/path/to/file.docx",
        "recovery": "请提供有效的文档路径",
        "can_retry": True
    }
}
```

## 预置模板

| 模板 | 适用场景 |
|------|---------|
| `business_report` | 商务报告 - 黑体标题 + 宋体正文 + 1.5倍行距 |
| `internal_simple` | 简洁风格 - 适合内部通知 |
| `academic_paper` | 学术论文 - 首行缩进 + 1.5倍行距 |

### 模板配置详情

```python
# business_report - 商务报告
{
    "margins": {"top": "2.54cm", "bottom": "2.54cm", "left": "3.17cm", "right": "3.17cm"},
    "default_font": "宋体",
    "title_font": "黑体",
    "line_spacing": 1.5,
    "first_line_indent": 0
}

# academic_paper - 学术论文
{
    "margins": {"top": "2.54cm", "bottom": "2.54cm", "left": "3cm", "right": "3cm"},
    "default_font": "宋体",
    "title_font": "黑体",
    "line_spacing": 1.5,
    "first_line_indent": 2  # 首行缩进 2 字符
}

# internal_simple - 内部简报
{
    "margins": {"top": "2cm", "bottom": "2cm", "left": "2.5cm", "right": "2.5cm"},
    "default_font": "宋体",
    "title_font": "黑体",
    "line_spacing": 1.15,
    "first_line_indent": 0
}
```

## LLM 使用流程

### 1. 理解用户需求

用户用自然语言描述需求，如"这个文档格式很乱"、"帮我转换成专业报告风格"。

### 2. 匹配意图

根据用户输入匹配对应的工作流或工具。参考 [references/intent-mapping.csv](references/intent-mapping.csv)。

### 3. 执行操作

- 只读操作：直接调用 validator 工具
- 修改操作：开始会话 → 执行修改 → 提交/回滚

### 4. 返回结果

返回统一格式的响应，包含成功状态和详细信息。

## 详细参考

- **[references/api.md](references/api.md)** - 完整 API 参考和参数说明
- **[references/tech-spec-v1.2.md](references/tech-spec-v1.2.md)** - V1.2 新增接口技术规格
- **[references/workflows.md](references/workflows.md)** - 工作流详细定义和执行流程
- **[references/templates.md](references/templates.md)** - 模板配置和自定义指南
- **[references/intent-mapping.csv](references/intent-mapping.csv)** - 意图映射表

## 依赖

- Python >= 3.10
- python-docx >= 1.1.0
