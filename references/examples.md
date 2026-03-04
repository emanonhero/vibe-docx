# vibe-docx 示例代码

> 完整代码示例，按工作流分类。SKILL.md 只保留速查索引。

## 导入说明

```python
# 方式一：主模块 API（格式操作）
import vibe_docx

vibe_docx.analyze("doc.docx")
vibe_docx.begin_session("doc.docx", backup=True)
vibe_docx.fix_formatting(session["session_id"])
vibe_docx.commit(session["session_id"])

# 方式二：完整 API（内容操作必需）
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")

from scripts.validator import analyze, detect_textboxes
from scripts.builder import (
    begin_session, commit, rollback,
    table_list, table_read, table_update,
    read_text, replace_text, delete_paragraphs,
    fix_formatting, apply_style_template
)
```

---

## format-fix - 格式修复

**触发：** 格式乱、修复格式、整理格式、标题样式

```python
import vibe_docx

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

---

## analyze - 文档分析

**触发：** 分析文档、检测问题、文档诊断

```python
import vibe_docx

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

---

## style-convert - 风格转换

**触发：** 变成风格、专业报告、风格转换、应用模板

```python
import vibe_docx

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

---

## content-edit - 内容编辑

**触发：** 修改内容、编辑表格、替换文本、基于模板创建

### 修改表格数据

```python
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")
from scripts.validator import analyze
from scripts.builder import begin_session, table_list, table_read, table_update, commit

# 1. 分析文档结构
result = analyze("report.docx")
print(f"表格数: {result['document_info']['tables']}")

# 2. 开始会话
session = begin_session("report.docx", backup=True)

# 3. 列出表格
tables = table_list(session["session_id"])
for t in tables["tables"]:
    print(f"表格 {t['index']}: {t['rows']}行 x {t['cols']}列")

# 4. 读取表格内容
content = table_read(session["session_id"], table_index=0)
print(content["content"])  # [[cell1, cell2, ...], ...]

# 5. 更新单元格
table_update(session["session_id"], table_index=0, cells=[
    {"row": 0, "col": 0, "text": "新标题"},
    {"row": 1, "col": 2, "text": "更新数据"},
])

# 6. 提交
commit(session["session_id"])
```

### 替换段落文本

```python
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")
from scripts.builder import begin_session, read_text, replace_text, commit

session = begin_session("contract.docx", backup=True)

# 读取段落
para = read_text(session["session_id"], paragraph_index=5)
print(f"原文: {para['text']}")

# 替换文本
replace_text(session["session_id"], paragraph_index=5, text="新的段落内容")

commit(session["session_id"])
```

### 删除段落范围

```python
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")
from scripts.builder import begin_session, delete_paragraphs, commit

session = begin_session("report.docx", backup=True)

# 删除第 10-15 段（保留图片）
delete_paragraphs(
    session["session_id"],
    start_index=10,
    end_index=15,
    options={"preserve_images": True}
)

commit(session["session_id"])
```

### 基于模板创建文档

```python
import sys
sys.path.insert(0, "path/to/vibe_docx/skill_assets")
from scripts.builder import begin_session, table_update, replace_text, commit

# 1. 打开模板
session = begin_session("template_resume.docx", backup=False)

# 2. 填充表格（如个人信息表格）
table_update(session["session_id"], table_index=0, cells=[
    {"row": 0, "col": 1, "text": "张三"},
    {"row": 1, "col": 1, "text": "高级工程师"},
    {"row": 2, "col": 1, "text": "zhangsan@example.com"},
])

# 3. 替换占位符文本
replace_text(session["session_id"], paragraph_index=10, text="个人简介内容...")

# 4. 提交为新文件
commit(session["session_id"], output_path="简历_张三.docx")
```

---

## table-organize - 表格整理

**触发：** 表格样式、表格整理、表格边框、添加边框

```python
import vibe_docx

session = vibe_docx.begin_session("report.docx", backup=True)
vibe_docx.fix_table_borders(session["session_id"], border_style="single")
vibe_docx.commit(session["session_id"])
```

---

## section-operate - 章节操作

**触发：** 章节大纲、拆分章节、提取章节、删除章节

```python
import vibe_docx

session = vibe_docx.begin_session("report.docx", backup=True)
vibe_docx.remove_section(session["session_id"], "附录")
vibe_docx.add_section(session["session_id"], "新章节", "章节内容")
vibe_docx.commit(session["session_id"])
```

---

## document-merge - 文档合并

**触发：** 合并文档、批量合并、多个文档合并

```python
import vibe_docx

vibe_docx.merge_documents(
    ["part1.docx", "part2.docx", "part3.docx"],
    "merged.docx",
    options={"add_page_break": True}
)
```

---

## textbox-extract - 文本框处理

**触发：** 文本框、提取文本框、简历分析

```python
import vibe_docx

# 检测文本框
textboxes = vibe_docx.detect_textboxes("resume.docx")
print(f"发现 {textboxes['stats']['total_count']} 个文本框")

# 提取内容
session = vibe_docx.begin_session("resume.docx", backup=True)
vibe_docx.extract_textbox_content(session["session_id"])
vibe_docx.commit(session["session_id"])
```

---

## 会话管理

```python
import vibe_docx

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

---

## 错误处理示例

```python
import vibe_docx

result = vibe_docx.analyze("report.docx")

if not result["success"]:
    error = result["error"]
    print(f"错误码: {error['code']}")
    print(f"错误信息: {error['message']}")
    print(f"恢复建议: {error['recovery']}")
    
    if error["can_retry"]:
        # 重试逻辑
        pass
```

---

## 预置模板配置

| 模板 | 字体 | 行距 | 首行缩进 | 页边距 |
|------|------|------|---------|--------|
| `business_report` | 宋体/黑体 | 1.5 | 无 | 上下2.54cm 左右3.17cm |
| `academic_paper` | 宋体/黑体 | 1.5 | 2字符 | 上下2.54cm 左右3cm |
| `internal_simple` | 宋体/黑体 | 1.15 | 无 | 上下2cm 左右2.5cm |

详见 [templates.md](templates.md)
