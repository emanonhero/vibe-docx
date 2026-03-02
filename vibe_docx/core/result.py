# -*- coding: utf-8 -*-
"""
vibe_docx.core.result - 统一返回值格式

定义 Result 数据类，确保所有工具返回统一格式。
"""

from dataclasses import dataclass, field
from typing import Optional, Any, Dict


@dataclass
class Error:
    """错误信息"""
    code: str
    message: str
    detail: Optional[str] = None
    recovery: Optional[str] = None
    can_retry: bool = True
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        result = {
            "code": self.code,
            "message": self.message,
        }
        if self.detail:
            result["detail"] = self.detail
        if self.recovery:
            result["recovery"] = self.recovery
        result["can_retry"] = self.can_retry
        return result


@dataclass
class Result:
    """
    统一返回值格式。
    
    所有工具都应返回此格式的对象，确保 LLM 可以一致地解读返回值。
    
    Examples:
        # 成功响应
        Result(success=True, data={"issues": [...], "stats": {...}})
        
        # 错误响应
        Result(
            success=False, 
            error=Error(code="DOC001", message="文档不存在", detail=path)
        )
    """
    success: bool
    data: Optional[Any] = None
    error: Optional[Error] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """转换为字典"""
        result: Dict[str, Any] = {"success": self.success}
        
        if self.data is not None:
            result["data"] = self.data
        
        if self.error is not None:
            result["error"] = self.error.to_dict()
        
        if self.metadata:
            result["metadata"] = self.metadata
        
        return result
    
    @classmethod
    def ok(cls, data: Any = None, **metadata) -> "Result":
        """
        创建成功结果。
        
        Args:
            data: 返回数据
            **metadata: 可选元数据
            
        Returns:
            Result 对象
        """
        return cls(success=True, data=data, metadata=metadata if metadata else {})
    
    @classmethod
    def fail(
        cls, 
        code: str, 
        message: str, 
        detail: Optional[str] = None,
        recovery: Optional[str] = None,
        can_retry: bool = True
    ) -> "Result":
        """
        创建失败结果。
        
        Args:
            code: 错误码
            message: 错误消息
            detail: 详细信息
            recovery: 恢复建议
            can_retry: 是否可重试
            
        Returns:
            Result 对象
        """
        return cls(
            success=False,
            error=Error(
                code=code,
                message=message,
                detail=detail,
                recovery=recovery,
                can_retry=can_retry
            )
        )
    
    def with_metadata(self, **kwargs) -> "Result":
        """添加元数据"""
        self.metadata.update(kwargs)
        return self


# 类型别名
ResultDict = Dict[str, Any]


def error_response(code: str, message: str, detail: str = "") -> ResultDict:
    """
    快捷创建错误响应字典。
    
    Args:
        code: 错误码
        message: 错误消息
        detail: 详细信息
        
    Returns:
        错误响应字典
    """
    return Result.fail(code, message, detail).to_dict()


def success_response(data: Any = None, **metadata) -> ResultDict:
    """
    快捷创建成功响应字典。
    
    Args:
        data: 返回数据
        **metadata: 可选元数据
        
    Returns:
        成功响应字典
    """
    return Result.ok(data, **metadata).to_dict()


__all__ = [
    "Error",
    "Result", 
    "ResultDict",
    "error_response",
    "success_response",
]
