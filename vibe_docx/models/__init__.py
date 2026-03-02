# -*- coding: utf-8 -*-
"""
vibe_docx.models - 数据模型

包含文档、章节、文本框、错误等数据模型。
"""

from .error import (
    ErrorCategory,
    ErrorDefinition,
    ERROR_DEFINITIONS,
    get_error_definition,
    get_error_say,
    get_error_then,
    is_retryable,
    is_session_error,
    is_document_error,
)

__all__ = [
    "ErrorCategory",
    "ErrorDefinition",
    "ERROR_DEFINITIONS",
    "get_error_definition",
    "get_error_say",
    "get_error_then",
    "is_retryable",
    "is_session_error",
    "is_document_error",
]
