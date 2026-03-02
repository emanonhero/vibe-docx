# -*- coding: utf-8 -*-
"""
vibe_docx.core - 核心组件

包含 Validator, Builder, Session, Result 等核心类。
"""

from .result import Result, Error, error_response, success_response

__all__ = [
    "Result",
    "Error",
    "error_response", 
    "success_response",
]
