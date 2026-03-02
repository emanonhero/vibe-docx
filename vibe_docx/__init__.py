# -*- coding: utf-8 -*-
"""
vibe-docx - 用自然语言操作 Word 文档

支持分析、编辑、格式转换、批量处理 DOCX 文档。

Examples:
    # 分析文档（只读）
    from vibe_docx import analyze
    result = analyze("document.docx")
    print(result["document_info"])
    
    # 开始编辑会话
    from vibe_docx import begin_session, commit, fix_formatting
    session = begin_session("document.docx", backup=True)
    fix_formatting(session["session_id"])
    commit(session["session_id"])
"""

__version__ = "1.0.0"

# 从 scripts 导入现有功能（保持向后兼容）
from scripts.validator import (
    analyze,
    detect_textboxes,
    get_section_outline,
    get_document_structure,
    validate_xml,
    ValidatorError,
)

from scripts.builder import (
    begin_session,
    commit,
    rollback,
    fix_formatting,
    fix_page_setup,
    fix_table_borders,
    fix_list_formatting,
    apply_style_template,
    add_section,
    remove_section,
    move_section,
    merge_documents,
    split_document,
    extract_textbox_content,
    get_template,
    image_list,
    image_insert,
    image_export,
)

# 新模块导出
from vibe_docx.core import Result, Error, error_response, success_response
from vibe_docx.models import (
    ERROR_DEFINITIONS,
    get_error_definition,
    get_error_say,
    is_retryable,
)

__all__ = [
    # 版本
    "__version__",
    
    # 只读工具（无需会话）
    "analyze",
    "detect_textboxes",
    "get_section_outline",
    "get_document_structure",
    "validate_xml",
    
    # 会话管理
    "begin_session",
    "commit",
    "rollback",
    
    # 格式修复
    "fix_formatting",
    "fix_page_setup",
    "fix_table_borders",
    "fix_list_formatting",
    "apply_style_template",
    
    # 章节操作
    "add_section",
    "remove_section",
    "move_section",
    
    # 批量操作
    "merge_documents",
    "split_document",
    
    # 文本框
    "extract_textbox_content",
    
    # 模板
    "get_template",
    
    # 图片操作
    "image_list",
    "image_insert",
    "image_export",
    
    # 新类型
    "Result",
    "Error",
    "error_response",
    "success_response",
    "ERROR_DEFINITIONS",
    "get_error_definition",
    "get_error_say",
    "is_retryable",
    
    # 异常
    "ValidatorError",
]
