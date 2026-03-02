"""
docx SKILL - Word 文档处理工具集

Usage:
    from scripts import analyze, begin_session, commit, fix_formatting
    
    # 分析文档
    result = analyze("document.docx")
    
    # 编辑文档
    session = begin_session("document.docx", backup=True)
    fix_formatting(session["session_id"])
    commit(session["session_id"])
"""

# Validator 工具（只读，无需会话）
from .validator import (
    analyze,
    get_document_structure,
    get_section_outline,
    detect_textboxes,
    validate_xml,
)

# Builder 工具（修改，需要会话）
from .builder import (
    # 会话管理
    begin_session,
    commit,
    rollback,
    # 格式操作
    fix_formatting,
    fix_page_setup,
    fix_table_borders,
    fix_list_formatting,
    # 章节操作
    add_section,
    remove_section,
    move_section,
    # 批量操作
    merge_documents,
    split_document,
    # 文本框操作
    extract_textbox_content,
    textbox_to_paragraph,
    remove_textbox,
    # 模板
    get_template,
    # 图片操作
    image_list,
    image_insert,
    image_export,
    # 表格操作
    table_list,
    table_read,
    table_update,
    table_create,
    # 文本操作
    read_section,
    read_text,
    replace_text,
    splice_section,
    # Markdown
    insert_markdown,
    markdown_to_document,
)

# Markdown 解析器
from .markdown import MarkdownParser, parse_markdown_to_xml

__all__ = [
    # Validator tools
    "analyze",
    "get_document_structure",
    "get_section_outline",
    "detect_textboxes",
    "validate_xml",
    # Builder tools - 会话管理
    "begin_session",
    "commit",
    "rollback",
    # Builder tools - 格式操作
    "fix_formatting",
    "fix_page_setup",
    "fix_table_borders",
    "fix_list_formatting",
    # Builder tools - 章节操作
    "add_section",
    "remove_section",
    "move_section",
    "merge_documents",
    "split_document",
    # Builder tools - 文本框操作
    "extract_textbox_content",
    "textbox_to_paragraph",
    "remove_textbox",
    # Builder tools - 模板
    "get_template",
    # Builder tools - 图片操作
    "image_list",
    "image_insert",
    "image_export",
    # Builder tools - 表格操作
    "table_list",
    "table_read",
    "table_update",
    "table_create",
    # Builder tools - 文本操作
    "read_section",
    "read_text",
    "replace_text",
    "splice_section",
    # Builder tools - Markdown
    "insert_markdown",
    "markdown_to_document",
    # Markdown 解析器
    "MarkdownParser",
    "parse_markdown_to_xml",
]
