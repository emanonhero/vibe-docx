# vibe-docx

> AI-First Word 文档处理库 - 让 LLM 用自然语言操作 DOCX

[![PyPI version](https://badge.fury.io/py/vibe-docx.svg)](https://badge.fury.io/py/vibe-docx)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 为什么 AI-First？

传统 DOCX 库为人类设计，返回复杂对象需二次处理。**vibe-docx 为 LLM 优化**：

| 传统库 | vibe-docx |
|--------|-----------|
| 返回复杂对象 | 返回 JSON 字典 |
| 需要链式调用 | 一句话完成操作 |
| 错误需 try-catch | 统一错误码 + 恢复建议 |
| 无会话概念 | 内置会话管理 + 自动回滚 |

```python
# 传统方式 - 对 AI 不友好
from docx import Document
doc = Document("report.docx")
for para in doc.paragraphs:  # AI 需要遍历理解结构
    for run in para.runs:
        run.font.name = "宋体"  # 需要知道对象模型
doc.save("report.docx")

# vibe-docx - AI 友好
import vibe_docx
session = vibe_docx.begin_session("report.docx")
vibe_docx.fix_formatting(session["session_id"], {"default_font": "宋体"})
vibe_docx.commit(session["session_id"])
```

## 安装

```bash
pip install vibe-docx
```

## 快速开始

```python
import vibe_docx

# 分析文档（只读）
result = vibe_docx.analyze("document.docx")
print(result["document_info"])  # {"paragraphs_count": 150, ...}
print(result["issues"])          # [{"id": "table_borders_missing", ...}]

# 编辑文档
session = vibe_docx.begin_session("document.docx", backup=True)
vibe_docx.fix_formatting(session["session_id"], options={
    "default_font": "宋体",
    "title_font": "黑体",
    "line_spacing": 1.5
})
vibe_docx.commit(session["session_id"])  # 或 rollback() 回滚
```

## 核心能力

### 只读工具（无需会话）

| 工具 | 说明 |
|------|------|
| `analyze(file_path)` | 全面分析：格式、结构、内容问题 |
| `detect_textboxes(file_path)` | 检测文本框 |
| `get_section_outline(file_path)` | 获取章节大纲 |
| `validate_xml(doc_path)` | XML 底层验证 |

### 修改工具（需要会话）

| 工具 | 说明 |
|------|------|
| `begin_session(file_path, backup=True)` | 开始会话，自动备份 |
| `commit(session_id)` | 提交修改 |
| `rollback(session_id)` | 回滚到原始文件 |
| `fix_formatting(session_id, options)` | 修复字体、行距、空段落 |
| `fix_page_setup(session_id, margins)` | 页边距、纸张方向 |
| `fix_table_borders(session_id)` | 添加表格边框 |
| `apply_style_template(session_id, template)` | 应用预置模板 |

### 预置模板

| 模板 | 适用场景 |
|------|---------|
| `business_report` | 商务报告 - 黑体标题 + 宋体正文 + 1.5倍行距 |
| `academic_paper` | 学术论文 - 首行缩进 + 1.5倍行距 |
| `internal_simple` | 简洁风格 - 内部通知 |

## AI 友好设计

### 1. 统一返回格式

所有工具返回一致的结构，LLM 无需处理多种格式：

```python
# 成功
{"success": True, "data": {...}, "changes": {...}}

# 失败
{
    "success": False,
    "error": {
        "code": "DOC001",
        "message": "文档不存在",
        "recovery": "请提供有效的文档路径",
        "can_retry": True
    }
}
```

### 2. 语义化错误码

错误包含恢复建议，LLM 可自主处理：

| 错误码 | 说明 | 恢复建议 |
|--------|------|---------|
| DOC001 | 文档不存在 | 请提供有效的文档路径 |
| DOC002 | 不支持的格式 | 仅支持 .docx 格式 |
| SES001 | 会话无效 | 请创建新会话 |
| SES003 | 会话冲突 | 文档正被其他会话使用 |

### 3. 会话 + 自动备份

```python
session = vibe_docx.begin_session("report.docx", backup=True)
# 自动创建备份: .vibe-backups/report.backup_1234567890.docx

try:
    vibe_docx.fix_formatting(session["session_id"])
    vibe_docx.commit(session["session_id"])
except:
    vibe_docx.rollback(session["session_id"])  # 自动恢复
```

### 4. 一句话完成复杂操作

```python
# 应用完整的商务报告风格
vibe_docx.apply_style_template(session_id, "business_report")

# 合并多个文档
vibe_docx.merge_documents(["part1.docx", "part2.docx"], "merged.docx")
```

## 作为 Claude Skill 使用

vibe-docx 包含 SKILL.md，可直接作为 Claude/iFlow 的 skill 使用：

```
# 安装 skill
cp -r vibe-docx ~/.iflow/skills/

# Claude 自动识别
用户: 帮我修复 report.docx 的格式
Claude: [触发 vibe-docx skill] 正在分析文档...
```

## API 参考

完整 API 文档见 [references/api.md](references/api.md)

- [工作流定义](references/workflows.md)
- [模板配置](references/templates.md)
- [意图映射](references/intent-mapping.csv)
- [V1.2 技术规格](references/tech-spec-v1.2.md)

## 依赖

- Python >= 3.10
- python-docx >= 0.8.11
- lxml >= 4.9.0

## License

MIT License - 自由使用，欢迎贡献！

---

**Made for AI, by AI** 🤖
