# -*- coding: utf-8 -*-
"""
vibe_docx.models.error - 错误码定义

定义所有错误码及其对话模板，确保 LLM 可以正确处理错误。
"""

from dataclasses import dataclass
from typing import Dict, Optional, Any
from enum import Enum


class ErrorCategory(Enum):
    """错误类别"""
    DOC = "document"      # 文档相关
    SES = "session"       # 会话相关
    VAL = "validator"     # 验证相关
    BLD = "builder"       # 构建相关
    SYS = "system"        # 系统相关


@dataclass
class ErrorDefinition:
    """错误定义"""
    code: str
    message: str
    say: str              # LLM 对用户说的话
    then: str             # LLM 接下来应该做什么
    recovery: Optional[str] = None
    example: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        result = {
            "code": self.code,
            "message": self.message,
            "say": self.say,
            "then": self.then,
        }
        if self.recovery:
            result["recovery"] = self.recovery
        if self.example:
            result["example"] = self.example
        return result


# 错误码定义
ERROR_DEFINITIONS: Dict[str, ErrorDefinition] = {
    # 文档相关 DOC001-DOC010
    "DOC001": ErrorDefinition(
        code="DOC001",
        message="文档不存在",
        say="找不到文档 {path}",
        then="等待用户提供正确路径",
        recovery="请提供有效的文档路径",
        example="""
用户: 帮我修 /wrong/path.docx
AI: 找不到文档，请检查路径是否正确
用户: /correct/path.docx
AI: 好的，正在分析..."""
    ),
    "DOC002": ErrorDefinition(
        code="DOC002",
        message="不支持的文档格式",
        say="不支持此文档格式，仅支持 .docx",
        then="提示用户提供 .docx 文件",
        recovery="请将文档另存为 .docx 格式"
    ),
    "DOC003": ErrorDefinition(
        code="DOC003",
        message="文档已损坏",
        say="文档可能已损坏，无法读取",
        then="建议用户尝试修复或使用备份",
        recovery="尝试用 Word 打开并重新保存，或使用备份文件"
    ),
    "DOC004": ErrorDefinition(
        code="DOC004",
        message="文档被锁定",
        say="文档正被其他程序使用",
        then="提示用户关闭后重试",
        recovery="请关闭 Word 或其他程序后重试",
        example="""
AI: 文档被锁定，请关闭 Word 后重试
用户: 已关闭
AI: [重试] 正在分析..."""
    ),
    "DOC005": ErrorDefinition(
        code="DOC005",
        message="文档加密",
        say="文档已加密，无法处理",
        then="提示用户解密后重试",
        recovery="请先移除文档密码保护"
    ),
    
    # 会话相关 SES001-SES010
    "SES001": ErrorDefinition(
        code="SES001",
        message="会话无效",
        say="会话已失效或不存在",
        then="创建新会话",
        recovery="将创建新会话继续操作"
    ),
    "SES002": ErrorDefinition(
        code="SES002",
        message="会话已过期",
        say="会话已超过有效期",
        then="创建新会话",
        recovery="将创建新会话继续操作"
    ),
    "SES003": ErrorDefinition(
        code="SES003",
        message="会话冲突",
        say="文档已被其他会话锁定",
        then="提示用户关闭其他会话",
        recovery="请先完成或关闭其他编辑会话"
    ),
    "SES004": ErrorDefinition(
        code="SES004",
        message="备份失败",
        say="无法创建备份文件",
        then="提示用户检查磁盘空间和权限",
        recovery="请检查磁盘空间和文件权限"
    ),
    
    # 验证相关 VAL001-VAL010
    "VAL001": ErrorDefinition(
        code="VAL001",
        message="分析失败",
        say="文档分析过程中出现错误",
        then="提供错误详情，建议用户检查文档",
        recovery="检查文档是否损坏，尝试用 Word 重新保存"
    ),
    "VAL002": ErrorDefinition(
        code="VAL002",
        message="XML 解析失败",
        say="文档 XML 结构异常",
        then="建议用户用 Word 修复文档",
        recovery="用 Word 打开文档并保存，让 Word 自动修复"
    ),
    
    # 构建相关 BLD001-BLD010
    "BLD001": ErrorDefinition(
        code="BLD001",
        message="执行失败",
        say="操作执行失败",
        then="提供错误详情，自动回滚",
        recovery="已自动回滚，请重试或报告问题"
    ),
    "BLD002": ErrorDefinition(
        code="BLD002",
        message="回滚失败",
        say="无法回滚到原始状态",
        then="建议用户使用备份文件",
        recovery="请使用备份文件恢复原始文档"
    ),
    "BLD003": ErrorDefinition(
        code="BLD003",
        message="保存失败",
        say="无法保存文档",
        then="提示用户检查磁盘空间和权限",
        recovery="请检查磁盘空间和文件写入权限"
    ),
    
    # 系统相关 SYS001-SYS010
    "SYS001": ErrorDefinition(
        code="SYS001",
        message="未知错误",
        say="操作失败，请重试",
        then="建议用户稍后重试",
        recovery="请稍后重试，如问题持续请报告"
    ),
    "SYS002": ErrorDefinition(
        code="SYS002",
        message="依赖缺失",
        say="缺少必要的依赖库",
        then="提示用户安装依赖",
        recovery="请运行 pip install python-docx lxml"
    ),
}


def get_error_definition(code: str) -> Optional[ErrorDefinition]:
    """
    获取错误定义。
    
    Args:
        code: 错误码
        
    Returns:
        ErrorDefinition 或 None
    """
    return ERROR_DEFINITIONS.get(code)


def get_error_say(code: str, **kwargs) -> str:
    """
    获取错误消息模板。
    
    Args:
        code: 错误码
        **kwargs: 模板变量
        
    Returns:
        格式化后的消息
    """
    definition = ERROR_DEFINITIONS.get(code)
    if definition:
        return definition.say.format(**kwargs)
    return f"错误 {code}"


def get_error_then(code: str) -> str:
    """
    获取错误后续动作。
    
    Args:
        code: 错误码
        
    Returns:
        后续动作描述
    """
    definition = ERROR_DEFINITIONS.get(code)
    if definition:
        return definition.then
    return "报告错误"


def is_retryable(code: str) -> bool:
    """
    判断错误是否可重试。
    
    Args:
        code: 错误码
        
    Returns:
        是否可重试
    """
    non_retryable = {"DOC002", "DOC005", "SYS002"}  # 格式不支持、加密、依赖缺失
    return code not in non_retryable


def is_session_error(code: str) -> bool:
    """
    判断是否为会话错误。
    
    Args:
        code: 错误码
        
    Returns:
        是否为会话错误
    """
    return code.startswith("SES")


def is_document_error(code: str) -> bool:
    """
    判断是否为文档错误。
    
    Args:
        code: 错误码
        
    Returns:
        是否为文档错误
    """
    return code.startswith("DOC")


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
