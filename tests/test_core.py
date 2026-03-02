# -*- coding: utf-8 -*-
"""
Tests for vibe_docx core module.
"""

import pytest
import tempfile
import os
import shutil
from vibe_docx import (
    Result,
    Error,
    analyze,
    detect_textboxes,
    get_error_say,
    get_error_definition,
    ERROR_DEFINITIONS,
)
from scripts.builder import (
    begin_session,
    commit,
    rollback,
    SessionManager,
    SessionError,
    fix_formatting,
    apply_style_template,
    merge_documents,
)


class TestResult:
    """Test Result dataclass."""
    
    def test_ok_returns_success(self):
        result = Result.ok({"key": "value"})
        assert result.success is True
        assert result.data == {"key": "value"}
        assert result.error is None
    
    def test_fail_returns_error(self):
        result = Result.fail("DOC001", "文档不存在")
        assert result.success is False
        assert result.error is not None
        assert result.error.code == "DOC001"
    
    def test_to_dict(self):
        result = Result.ok({"test": 1})
        d = result.to_dict()
        assert d["success"] is True
        assert d["data"] == {"test": 1}


class TestError:
    """Test Error dataclass."""
    
    def test_error_creation(self):
        error = Error("DOC001", "测试错误")
        assert error.code == "DOC001"
        assert error.message == "测试错误"
    
    def test_error_to_dict(self):
        error = Error("DOC001", "测试错误", detail="/path/to/file")
        d = error.to_dict()
        assert d["code"] == "DOC001"
        assert d["message"] == "测试错误"
        assert d["detail"] == "/path/to/file"
        assert d["can_retry"] is True


class TestErrorCodes:
    """Test error code definitions."""
    
    def test_error_definitions_count(self):
        assert len(ERROR_DEFINITIONS) >= 16
    
    def test_get_error_definition(self):
        error = get_error_definition("DOC001")
        assert error is not None
        assert error.message == "文档不存在"
    
    def test_get_error_say(self):
        say = get_error_say("DOC001", path="/test.docx")
        assert "找不到文档" in say
        assert "/test.docx" in say
    
    def test_is_retryable(self):
        from vibe_docx import is_retryable
        assert is_retryable("DOC001") is True


class TestValidator:
    """Test validator functions."""
    
    def test_analyze_missing_file(self):
        result = analyze("/nonexistent/path.docx")
        assert result["success"] is False
        assert result["error"]["code"] == "DOC001"
    
    def test_analyze_invalid_format(self):
        import tempfile
        import os
        
        # 创建一个临时的非 docx 文件
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"test content")
            temp_path = f.name
        
        try:
            result = analyze(temp_path)
            assert result["success"] is False
            assert result["error"]["code"] == "DOC002"
        finally:
            os.unlink(temp_path)


class TestSession:
    """Test session management."""
    
    def test_begin_session_missing_file(self):
        result = begin_session("/nonexistent/file.docx")
        assert result["success"] is False
        assert result["error"]["code"] == "DOC001"
    
    def test_begin_session_invalid_format(self):
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"test")
            temp_path = f.name
        
        try:
            result = begin_session(temp_path)
            assert result["success"] is False
            assert result["error"]["code"] == "DOC002"
        finally:
            os.unlink(temp_path)
    
    def test_session_id_format(self):
        """Session ID should be ses_ prefix + hex"""
        # 使用测试文档
        test_doc = r"C:\GithubCodes\aer-bmad-method\测试用原始WORD.docx"
        if os.path.exists(test_doc):
            result = begin_session(test_doc)
            assert result["success"] is True
            session_id = result["session_id"]
            assert session_id.startswith("ses_")
            assert len(session_id) == 16  # ses_ + 12 hex chars
            
            # 清理
            if result["success"]:
                rollback(result["session_id"])
    
    def test_invalid_session_id(self):
        """Invalid session ID should return error"""
        result = rollback("invalid_id")
        assert result["success"] is False
        assert result["error"]["code"] == "SES001"
    
    def test_session_manager_stats(self):
        manager = SessionManager.get_instance()
        stats = manager.get_stats()
        assert "active_sessions" in stats
        assert "backups" in stats
        assert "locked_files" in stats


class TestFixFormatting:
    """Test fix_formatting function."""
    
    @pytest.fixture
    def temp_docx(self):
        """Create a temporary docx file for testing."""
        import shutil
        test_doc = r"C:\GithubCodes\aer-bmad-method\测试用原始WORD.docx"
        if not os.path.exists(test_doc):
            pytest.skip("Test document not found")
        
        temp_dir = tempfile.mkdtemp()
        temp_doc = os.path.join(temp_dir, "test.docx")
        shutil.copy(test_doc, temp_doc)
        yield temp_doc
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    def test_fix_formatting_basic(self, temp_docx):
        """Test basic fix_formatting functionality."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = fix_formatting(session["session_id"], options={
            "default_font": "宋体",
            "title_font": "黑体",
        })
        
        assert result["success"] is True
        assert "changes" in result
        assert result["changes"]["fonts_unified"] >= 0
        
        # Cleanup
        commit(session["session_id"])
    
    def test_fix_formatting_remove_empty_paragraphs(self, temp_docx):
        """Test removing empty paragraphs."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = fix_formatting(session["session_id"], options={
            "remove_empty_paragraphs": True,
        })
        
        assert result["success"] is True
        assert result["changes"]["empty_paragraphs_removed"] >= 0
        
        # Cleanup
        commit(session["session_id"])
    
    def test_fix_formatting_line_spacing(self, temp_docx):
        """Test fixing line spacing."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = fix_formatting(session["session_id"], options={
            "line_spacing": 1.5,
        })
        
        assert result["success"] is True
        assert result["changes"]["line_spacing_fixed"] >= 0
        
        # Cleanup
        commit(session["session_id"])


class TestApplyStyleTemplate:
    """测试 apply_style_template 函数"""
    
    @pytest.fixture
    def temp_docx(self):
        """创建临时测试文档"""
        test_doc = r"C:\GithubCodes\aer-bmad-method\测试用原始WORD.docx"
        if not os.path.exists(test_doc):
            pytest.skip("Test document not found")
        
        temp_dir = tempfile.mkdtemp()
        temp_doc = os.path.join(temp_dir, "test.docx")
        shutil.copy(test_doc, temp_doc)
        yield temp_doc
        shutil.rmtree(temp_dir, ignore_errors=True)
    
    def test_apply_business_report_template(self, temp_docx):
        """Test applying business_report template."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = apply_style_template(session["session_id"], "business_report")
        
        assert result["success"] is True
        assert "template" in result
        assert result["template"] == "business_report"
        
        # Cleanup
        commit(session["session_id"])
    
    def test_apply_academic_paper_template(self, temp_docx):
        """Test applying academic_paper template."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = apply_style_template(session["session_id"], "academic_paper")
        
        assert result["success"] is True
        assert result["template"] == "academic_paper"
        
        # Cleanup
        commit(session["session_id"])
    
    def test_apply_template_with_custom_options(self, temp_docx):
        """Test applying template with custom options override."""
        session = begin_session(temp_docx)
        assert session["success"] is True
        
        result = apply_style_template(
            session["session_id"],
            "business_report",
            custom_options={"default_font": "楷体", "line_spacing": 2.0}
        )
        
        assert result["success"] is True
        assert "changes" in result
        
        # Cleanup
        commit(session["session_id"])
    
    def test_apply_template_invalid_session(self):
        """Test applying template with invalid session."""
        result = apply_style_template("invalid_session_id", "business_report")
        
        assert result["success"] is False
        assert result["error"]["code"] == "SES001"


class TestMergeDocuments:
    """测试 merge_documents 函数"""
    
    def test_merge_empty_file_list(self):
        """Test merge with empty file list."""
        result = merge_documents([], "output.docx")
        
        assert result["success"] is False
        assert result["error"]["code"] == "DOC001"
    
    def test_merge_missing_file(self):
        """Test merge with missing file."""
        result = merge_documents(["nonexistent.docx"], "output.docx")
        
        assert result["success"] is False
        assert result["error"]["code"] == "DOC001"
