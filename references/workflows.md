# 工作流定义

本文档详细说明各工作流的执行流程和步骤。

## 目录

1. [format-fix - 格式修复](#format-fix---格式修复)
2. [style-convert - 风格转换](#style-convert---风格转换)
3. [table-organize - 表格整理](#table-organize---表格整理)
4. [document-merge - 文档合并](#document-merge---文档合并)
5. [textbox-extract - 文本框提取](#textbox-extract---文本框提取)

---

## format-fix - 格式修复

分析和修复文档格式问题，包括标题样式、段落格式、列表编号、页边距等。

### 触发模式

```
格式.*乱|修复.*格式|整理.*格式|格式化.*文档|标题.*样式|段落.*格式
```

### 执行流程

```
┌─────────────┐
│   analyze   │ 分析文档结构和问题
└──────┬──────┘
       ▼
┌─────────────┐
│begin_session│ 开始编辑会话（创建备份）
└──────┬──────┘
       ▼
┌─────────────┐
│fix_formatting│ 修复格式
└──────┬──────┘
       ▼
┌─────────────┐
│   verify    │ 验证结果
└──────┬──────┘
       ▼
┌─────────────┐
│   commit    │ 提交修改
└─────────────┘
```

### 使用示例

```python
# 完整流程
from scripts.validator import analyze
from scripts.builder import begin_session, fix_formatting, commit

# 1. 分析文档
result = analyze("report.docx")
print(f"发现 {len(result['issues'])} 个问题")

# 2. 开始会话
session = begin_session("report.docx", backup=True)
session_id = session["session_id"]

# 3. 修复格式
fix_result = fix_formatting(session_id, options={
    "default_font": "宋体",
    "title_font": "黑体"
})
print(f"修复了 {fix_result['fixed_count']} 处格式")

# 4. 提交
commit(session_id)
```

---

## style-convert - 风格转换

将文档转换为专业风格（商务报告、简洁风格、学术论文）。

### 触发模式

```
变成.*风格|专业.*报告|风格.*转换|模板.*应用
```

### 预置风格

| 风格 | 标题字体 | 正文字体 | 行距 | 适用场景 |
|------|---------|---------|------|---------|
| `business_report` | 黑体 18pt | 宋体 12pt | 1.5倍 | 商业提案、项目报告 |
| `internal_simple` | 黑体 16pt | 宋体 12pt | 1.25倍 | 内部通知、简洁记录 |
| `academic_paper` | 黑体 16pt | 宋体 12pt | 1.5倍 + 首行缩进 | 研究报告、学术论文 |

### 执行流程

```
┌─────────────┐
│   analyze   │ 分析文档
└──────┬──────┘
       ▼
┌─────────────┐
│check_template│ 检查模板选择
└──────┬──────┘
       ▼ (如无模板)
┌─────────────┐
│suggest_tmpl │ 推荐模板
└──────┬──────┘
       ▼
┌─────────────┐
│begin_session│ 开始会话
└──────┬──────┘
       ▼
┌─────────────┐
│apply_titles │ 应用标题样式
└──────┬──────┘
       ▼
┌─────────────┐
│apply_page   │ 应用页面设置
└──────┬──────┘
       ▼
┌─────────────┐
│   commit    │ 提交修改
└─────────────┘
```

### 使用示例

```python
from scripts.builder import begin_session, fix_formatting, fix_page_setup, commit

session = begin_session("report.docx", backup=True)
session_id = session["session_id"]

# 应用商务报告风格
fix_page_setup(session_id, margins={
    "top": "2.54cm", "bottom": "2.54cm",
    "left": "3.17cm", "right": "3.17cm"
})
fix_formatting(session_id, options={
    "default_font": "宋体",
    "title_font": "黑体"
})

commit(session_id)
```

---

## table-organize - 表格整理

整理文档中的表格，添加边框、统一样式。

### 触发模式

```
表格.*样式|表格.*整理|表格.*边框|添加.*边框
```

### 执行流程

```
┌─────────────┐
│   analyze   │ 分析文档
└──────┬──────┘
       ▼
┌─────────────┐
│validate_xml │ 检测表格
└──────┬──────┘
       ▼
┌─────────────┐
│check_count  │ 检查表格数量
└──────┬──────┘
       ▼
┌─────────────┐
│begin_session│ 开始会话
└──────┬──────┘
       ▼
┌─────────────┐
│apply_borders│ 应用边框
└──────┬──────┘
       ▼
┌─────────────┐
│   verify    │ 验证结果
└──────┬──────┘
       ▼
┌─────────────┐
│   commit    │ 提交修改
└─────────────┘
```

### 使用示例

```python
from scripts.validator import validate_xml
from scripts.builder import begin_session, fix_table_borders, commit

# 检测表格
result = validate_xml("report.docx")
tables = result["detected"]["tables"]
print(f"发现 {len(tables)} 个表格")

# 找出无边框的表格
no_border_tables = [t["index"] for t in tables if not t["has_borders"]]

session = begin_session("report.docx", backup=True)
fix_table_borders(
    session["session_id"],
    table_indices=no_border_tables,  # 只修复无边框的表格
    border_style="single"
)
commit(session["session_id"])
```

---

## document-merge - 文档合并

合并多个文档为一个，支持添加分页符、统一样式。

### 触发模式

```
合并.*文档|批量.*合并|多个.*文档.*合并
```

### 执行流程

```
┌─────────────┐
│validate_file│ 验证文件
└──────┬──────┘
       ▼
┌─────────────┐
│   merge     │ 合并文档
└──────┬──────┘
       ▼
┌─────────────┐
│   report    │ 生成报告
└─────────────┘
```

### 使用示例

```python
from scripts.builder import merge_documents

result = merge_documents(
    ["part1.docx", "part2.docx", "part3.docx"],
    "merged.docx",
    options={
        "add_page_break": True,   # 文档间添加分页符
        "unify_styles": True      # 统一样式
    }
)

print(f"合并了 {result['stats']['files_merged']} 个文档")
print(f"总段落数: {result['stats']['total_paragraphs']}")
```

---

## textbox-extract - 文本框提取

检测、提取、转换文本框内容。

### 触发模式

```
文本框|提取.*文本框|检测.*文本框|简历.*分析
```

### 执行流程

```
┌─────────────┐
│   detect    │ 检测文本框
└──────┬──────┘
       ▼
┌─────────────┐
│ check_count │ 检查数量
└──────┬──────┘
       ▼ (有文本框)
┌─────────────┐
│begin_session│ 开始会话
└──────┬──────┘
       ▼
┌─────────────┐
│   extract   │ 提取内容
└──────┬──────┘
       ▼
┌─────────────┐
│   commit    │ 提交修改
└─────────────┘
```

### 使用示例

```python
from scripts.validator import detect_textboxes
from scripts.builder import begin_session, extract_textbox_content, commit

# 1. 检测文本框
result = detect_textboxes("resume.docx")
print(f"发现 {result['stats']['total_count']} 个文本框")

for tb in result["textboxes"]:
    print(f"  [{tb['index']}] {tb['content_preview'][:50]}...")

# 2. 提取内容
session = begin_session("resume.docx", backup=True)
extract_result = extract_textbox_content(session["session_id"])
print(f"提取了 {extract_result['extracted_count']} 个文本框的内容")

commit(session["session_id"])
```

---

## 意图映射

用户自然语言 → 工具映射：

| 用户输入模式 | 推荐工具 | 类型 |
|-------------|---------|------|
| 格式.*乱, 修复.*格式 | `analyze` + `fix_formatting` | 分析+修改 |
| 标题.*样式, 统一.*标题 | `fix_formatting` | 修改 |
| 变成.*风格, 专业.*报告 | `fix_page_setup` + `fix_formatting` | 修改 |
| 表格.*边框, 添加.*边框 | `fix_table_borders` | 修改 |
| 章节.*大纲, 文档.*结构 | `get_section_outline` | 只读 |
| 拆分.*章节, 提取.*章节 | `split_document` | 批量 |
| 合并.*文档 | `merge_documents` | 批量 |
| 文本框, 简历.*分析 | `detect_textboxes` | 只读 |
| 提取.*文本框 | `extract_textbox_content` | 修改 |
