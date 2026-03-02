# API 参考

## 目录

1. [会话管理](#会话管理)
2. [文档分析（Validator）](#文档分析validator)
3. [格式操作](#格式操作)
4. [章节操作](#章节操作)
5. [表格操作](#表格操作)
6. [图片操作](#图片操作)
7. [文本框操作](#文本框操作)
8. [文本操作](#文本操作)
9. [批量操作](#批量操作)
10. [Markdown 支持](#markdown-支持)
11. [模板工具](#模板工具)
12. [V1.2 新增接口](#v12-新增接口)
13. [错误处理](#错误处理)

---

## 会话管理

### begin_session(file_path, backup=True)

开始编辑会话，创建备份。

**参数：**
- `file_path` (str, 必填): DOCX 文件路径
- `backup` (bool, 可选): 是否创建备份，默认 True

**返回：**
```python
{
    "success": True,
    "session_id": "abc12345",
    "backup_path": "/path/to/document.backup_abc12345.docx"
}
```

### commit(session_id, output_path=None)

提交修改。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `output_path` (str, 可选): 输出路径，默认覆盖原文件

**返回：**
```python
{
    "success": True,
    "changes_count": 5,
    "output_path": "/path/to/document.docx"
}
```

### rollback(session_id)

回滚修改，恢复备份。

**参数：**
- `session_id` (str, 必填): 会话 ID

**返回：**
```python
{
    "success": True,
    "message": "已恢复到原始文件: /path/to/document.docx"
}
```

---

## 文档分析（Validator）

### analyze(file_path, focus_areas=None)

全面分析 DOCX 文档结构和问题。

**参数：**
- `file_path` (str, 必填): DOCX 文件路径
- `focus_areas` (list, 可选): 关注领域，值: `["format", "structure", "reference", "content"]`

**返回：**
```python
{
    "success": True,
    "document_info": {
        "paragraphs_count": 150,
        "tables_count": 5,
        "sections_count": 8,
        "file_size": 102400
    },
    "issues": [
        {
            "id": "table_borders_missing",
            "type": "table_no_borders",
            "category": "format",
            "severity": "info",
            "detail": "发现 3 个表格缺少边框",
            "auto_fixable": True
        }
    ],
    "risk_factors": []
}
```

### detect_textboxes(file_path)

检测文档中的所有文本框。

**参数：**
- `file_path` (str, 必填): DOCX 文件路径

**返回：**
```python
{
    "success": True,
    "textboxes": [
        {
            "index": 0,
            "has_content": True,
            "paragraph_count": 3,
            "content_preview": "文本内容预览...",
            "paragraphs": ["段落1", "段落2", "段落3"]
        }
    ],
    "stats": {
        "total_count": 5,
        "has_content_count": 4,
        "empty_count": 1
    }
}
```

### get_section_outline(file_path)

获取章节大纲。

**参数：**
- `file_path` (str, 必填): DOCX 文件路径

**返回：**
```python
{
    "success": True,
    "sections": [
        {"index": 0, "level": 1, "title": "第一章 概述", "paragraph_index": 5},
        {"index": 1, "level": 2, "title": "1.1 背景", "paragraph_index": 12}
    ]
}
```

### validate_xml(doc_path)

XML 直接验证（底层）。

**参数：**
- `doc_path` (str, 必填): DOCX 文件路径

**返回：**
```python
{
    "success": True,
    "detected": {
        "bold_elements": 25,
        "italic_elements": 10,
        "tables": [{"index": 0, "rows": 5, "cols": 3, "has_borders": False}],
        "images": [{"rid": "rId1", "width": 500000, "height": 300000}],
        "headings": [{"level": 1, "text": "标题", "style": "Heading1"}],
        "page_settings": {"width": 11906, "height": 16838, "margins": {...}}
    },
    "potential_issues": [
        {"type": "markdown_unconverted", "detail": "可能的 Markdown 语法未转换"}
    ]
}
```

---

## 格式操作

### fix_formatting(session_id, options=None)

修复格式问题。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `options` (dict, 可选): 修复选项
  - `default_font` (str): 默认字体，如 "宋体"
  - `font_size` (int): 字号，如 12
  - `title_font` (str): 标题字体

**返回：**
```python
{"success": True, "fixed_count": 50}
```

### fix_page_setup(session_id, margins=None, orientation=None, page_size=None)

设置页面。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `margins` (dict, 可选): 页边距
  - `top`, `bottom`, `left`, `right` (str): 如 "2.54cm"
- `orientation` (str, 可选): 方向，`"portrait"` 或 `"landscape"`
- `page_size` (dict, 可选): 纸张大小
  - `width`, `height` (str): 如 "21cm"

**返回：**
```python
{"success": True, "changes": ["top_margin: 2.54cm", "orientation: portrait"]}
```

### fix_table_borders(session_id, table_indices=None, border_style="single", border_width="0.5pt", border_color="000000")

添加表格边框。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `table_indices` (list, 可选): 表格索引列表，如 `[0, 1]`，默认全部
- `border_style` (str, 可选): 边框样式，默认 "single"
- `border_width` (str, 可选): 边框宽度，如 "0.5pt"
- `border_color` (str, 可选): 边框颜色（十六进制），如 "000000"

**返回：**
```python
{"success": True, "fixed_count": 3}
```

---

## 章节操作

### add_section(session_id, title, content="", position="end", level=1)

添加章节。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `title` (str, 必填): 章节标题
- `content` (str, 可选): 章节内容
- `position` (str, 可选): 位置，`"end"` | `"after:标题"` | `"before:标题"`
- `level` (int, 可选): 标题级别，默认 1

**返回：**
```python
{"success": True, "position": 25}
```

### remove_section(session_id, section_title)

删除章节。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `section_title` (str, 必填): 章节标题

**返回：**
```python
{"success": True, "removed_count": 10}
```

### move_section(session_id, section_title, new_position)

移动章节。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `section_title` (str, 必填): 章节标题
- `new_position` (str, 必填): 新位置，如 `"after:第二章"`

---

## 表格操作

### table_list(session_id)

列出文档中的所有表格。

**返回：**
```python
{
    "success": True,
    "tables": [
        {"index": 0, "rows": 5, "cols": 3, "first_cell": "表头内容..."}
    ],
    "count": 1
}
```

### table_read(session_id, table_index)

读取表格内容为二维数组。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `table_index` (int, 必填): 表格索引

**返回：**
```python
{
    "success": True,
    "content": [
        ["姓名", "年龄", "城市"],
        ["张三", "25", "北京"]
    ]
}
```

### table_update(session_id, table_index, cells)

更新表格单元格。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `table_index` (int, 必填): 表格索引
- `cells` (list): 单元格更新列表
  - `[{"row": 0, "col": 0, "text": "新值"}, ...]`

**返回：**
```python
{"success": True, "updated_count": 2}
```

### table_create(session_id, rows, cols, data=None, position=None)

创建新表格。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `rows` (int): 行数
- `cols` (int): 列数
- `data` (list, 可选): 初始数据，`[["A1", "B1"], ["A2", "B2"]]`
- `position` (dict, 可选): 插入位置

**返回：**
```python
{"success": True, "table_index": 2}
```

---

## 图片操作

### image_list(session_id, section=None)

列出文档中的所有图片。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `section` (str, 可选): 只列出指定章节内的图片

**返回：**
```python
{
    "success": True,
    "images": [
        {"index": 0, "rid": "rId1", "filename": "image_0", "size_kb": 0}
    ],
    "count": 1
}
```

### image_insert(session_id, image_path, position=None, width_inches=5.5)

插入图片。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `image_path` (str): 图片文件路径
- `position` (dict, 可选): 插入位置，如 `{"after_paragraph": 10}`
- `width_inches` (float): 图片宽度（英寸）

**返回：**
```python
{"success": True, "message": "已插入图片: /path/to/image.png"}
```

### image_export(session_id, rids, output_dir)

导出图片到指定目录。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `rids` (list): 要导出的图片 rId 列表
- `output_dir` (str): 输出目录

---

## 文本框操作

### extract_textbox_content(session_id, textbox_indices=None, mode="append")

提取文本框内容。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `textbox_indices` (list, 可选): 文本框索引，默认全部
- `mode` (str, 可选): 模式，`"append"` | `"prepend"` | `"replace"`

**返回：**
```python
{
    "success": True,
    "extracted_count": 3,
    "content": ["文本1", "文本2", "文本3"]
}
```

### textbox_to_paragraph(session_id, textbox_indices=None, position="end")

将文本框转换为段落。

### remove_textbox(session_id, textbox_indices, preserve_content=False)

删除文本框。

---

## 文本操作

### read_section(session_id, section_title)

读取指定章节的内容。

**返回：**
```python
{
    "success": True,
    "text": "章节内容...",
    "paragraph_count": 15
}
```

### read_text(session_id, paragraph_index, context=0)

读取指定位置的文本。

**参数：**
- `context` (int): 上下文段落数量

### replace_text(session_id, paragraph_index, content)

替换指定位置的文本。

### splice_section(session_id, section_title, content, preserve_images=True)

替换章节内容。

---

## 批量操作

### merge_documents(file_paths, output_path, options=None)

合并多个文档。

**参数：**
- `file_paths` (list): 文件路径列表
- `output_path` (str): 输出文件路径
- `options` (dict, 可选): 选项
  - `add_page_break` (bool): 是否添加分页符
  - `unify_styles` (bool): 是否统一样式

**返回：**
```python
{
    "success": True,
    "merged_path": "/path/to/merged.docx",
    "stats": {
        "files_merged": 3,
        "total_paragraphs": 150,
        "total_tables": 5
    }
}
```

### split_document(file_path, split_points, output_dir)

拆分文档。

**参数：**
- `file_path` (str): 源文件路径
- `split_points` (list): 拆分点列表，如 `["第一章", "第二章"]`
- `output_dir` (str): 输出目录

---

## Markdown 支持

### insert_markdown(session_id, markdown_content, position=None)

将 Markdown 内容插入文档。

**支持的格式：**
- 标题：`# H1` ~ `###### H6`
- 加粗：`**text**` 或 `__text__`
- 斜体：`*text*` 或 `_text_`
- 删除线：`~~text~~`
- 无序列表：`- item` 或 `* item`
- 有序列表：`1. item`
- 表格：`| col1 | col2 |`

**返回：**
```python
{"success": True, "paragraphs_added": 10}
```

### markdown_to_document(markdown_content, output_path, template=None)

将 Markdown 转换为新文档。

---

## 模板工具

### get_template(template_name)

获取模板配置。

**可用模板：**
- `business_report` - 商务报告
- `internal_simple` - 简洁风格
- `academic_paper` - 学术论文

**返回：**
```python
{
    "success": True,
    "template": {
        "title_styles": {
            "Heading 1": {"font_name": "黑体", "font_size": 18, "bold": True}
        },
        "paragraph_style": {"line_spacing": 1.5, "font_name": "宋体", "font_size": 12},
        "page_setup": {"margins": {"top": "2.54cm", ...}}
    }
}
```

---

## V1.2 新增接口

> 以下接口用于解决章节替换、Section 保护、精确插入等核心问题。详见 [tech-spec-v1.2.md](tech-spec-v1.2.md)。

### detect_section_boundaries(file_path)

检测文档中所有 section 边界位置。

**参数：**
- `file_path` (str, 必填): DOCX 文件路径

**返回：**
```python
{
    "success": True,
    "sections": [
        {
            "index": 0,
            "start_paragraph": 0,
            "end_paragraph": 240,
            "sectPr_location": "in_paragraph",
            "page_settings": {
                "width": "11906",
                "height": "16838",
                "margins": {"top": "1440", "bottom": "1440", "left": "1800", "right": "1800"},
                "orientation": "portrait"
            }
        }
    ],
    "total_count": 1
}
```

### delete_paragraphs(session_id, start_index, end_index, options)

安全删除段落范围，保护 section 结构。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `start_index` (int, 必填): 起始段落索引
- `end_index` (int, 必填): 结束段落索引（不包含）
- `options` (dict, 可选):
  - `preserve_sections` (bool): 保护 section 设置，默认 True
  - `preserve_images` (bool): 保留图片，默认 True

**返回：**
```python
{
    "success": True,
    "deleted_count": 10,
    "preserved_images": [{"original_paragraph": 5, "new_location": 20, "image_id": "rId1"}],
    "section_preserved": True
}
```

### insert_after_paragraph(session_id, anchor, content, options)

在指定段落后插入内容。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `anchor` (dict, 必填): 锚点定位
  - `type` (str): "index" | "title" | "text"
  - `value` (str | int): 锚点值
- `content` (dict, 必填): 插入内容
  - `type` (str): "text" | "markdown"
  - `text` (str): 内容文本
  - `style` (dict, 可选): 样式选项
    - `inherit` (bool): 是否继承周围样式，默认 True
- `options` (dict, 可选):
  - `add_paragraph` (bool): 是否创建新段落，默认 True

**返回：**
```python
{
    "success": True,
    "inserted_at": 6,
    "style_applied": {"font_name": "宋体", "font_size": 12}
}
```

### get_body_style(session_id)

提取文档正文样式。

**参数：**
- `session_id` (str, 必填): 会话 ID

**返回：**
```python
{
    "success": True,
    "body_style": {
        "font_name": "宋体",
        "font_size": 12,
        "line_spacing": 1.5,
        "first_line_indent": 24
    },
    "heading_styles": [
        {"level": 1, "font_name": "黑体", "font_size": 16, "bold": True}
    ]
}
```

### replace_section(session_id, section_title, new_content, options)

替换章节内容，保留图片和 section 结构。

**参数：**
- `session_id` (str, 必填): 会话 ID
- `section_title` (str, 必填): 要替换的章节标题
- `new_content` (dict, 必填): 新内容
  - `type` (str): "markdown" | "text"
  - `text` (str): 内容文本
- `options` (dict, 可选):
  - `preserve_images` (bool): 保留图片，默认 True
  - `image_position` (str): 图片定位策略，"smart" | "end" | "start"，默认 "smart"
  - `style_inherit` (bool): 继承文档样式，默认 True

**返回：**
```python
{
    "success": True,
    "old_section": {"start_index": 263, "end_index": 298, "image_count": 4},
    "new_section": {"start_index": 263, "end_index": 358, "paragraph_count": 95},
    "images_relocated": [
        {"original_paragraph": 272, "new_location": 326, "method": "smart"}
    ],
    "section_preserved": True
}
```

---

## 错误处理

所有工具返回统一格式：

```python
# 成功
{"success": True, "data": {...}}

# 失败
{
    "success": False,
    "error": {
        "code": "VAL001",
        "message": "文档不存在",
        "detail": "/path/to/file.docx",
        "recovery": "请提供有效的文档路径",
        "can_retry": True,
        "error_type": "file_error"
    }
}
```

**错误码：**
| 代码 | 说明 | 恢复建议 |
|------|------|---------|
| VAL001 | 文档不存在 | 提供有效路径 |
| VAL002 | 不支持的格式 | 仅支持 .docx |
| VAL003 | 文档已损坏 | 尝试修复或使用备份 |
