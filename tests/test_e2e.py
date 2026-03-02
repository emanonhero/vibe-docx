# -*- coding: utf-8 -*-
"""
End-to-End tests for vibe_docx workflows.
"""

import pytest
import tempfile
import os
import shutil
from vibe_docx import analyze, detect_textboxes
from scripts.builder import (
    begin_session,
    commit,
    rollback,
    fix_formatting,
    fix_table_borders,
    fix_page_setup,
    apply_style_template,
    merge_documents,
    split_document,
)


@pytest.fixture
def test_docx():
    """创建临时测试文档"""
    source = r"C:\GithubCodes\aer-bmad-method\测试用原始WORD.docx"
    if not os.path.exists(source):
        pytest.skip("Test document not found")
    
    temp_dir = tempfile.mkdtemp()
    temp_doc = os.path.join(temp_dir, "test.docx")
    shutil.copy(source, temp_doc)
    yield temp_doc
    shutil.rmtree(temp_dir, ignore_errors=True)


@pytest.fixture
def temp_dir():
    """创建临时目录"""
    temp = tempfile.mkdtemp()
    yield temp
    shutil.rmtree(temp, ignore_errors=True)


class TestFormatFixE2E:
    """E2E测试：格式修复工作流"""
    
    def test_analyze_and_fix_workflow(self, test_docx):
        """测试完整的分析和修复流程"""
        # 1. 分析文档
        analysis = analyze(test_docx)
        assert analysis["success"] is True
        
        # 获取问题统计
        issues = analysis.get("issues", {})
        
        # 2. 开始会话
        session = begin_session(test_docx, backup=True)
        assert session["success"] is True
        session_id = session["session_id"]
        
        # 3. 根据分析结果修复
        result = fix_formatting(session_id, options={
            "default_font": "宋体",
            "title_font": "黑体",
            "remove_empty_paragraphs": True,
            "line_spacing": 1.5,
        })
        assert result["success"] is True
        
        # 4. 提交
        commit_result = commit(session_id)
        assert commit_result["success"] is True
        
        # 5. 重新分析验证
        re_analysis = analyze(test_docx)
        assert re_analysis["success"] is True
    
    def test_font_unification_e2e(self, test_docx):
        """测试字体统一 E2E"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        # 统一字体
        result = fix_formatting(session_id, options={
            "default_font": "宋体",
            "title_font": "黑体",
        })
        
        assert result["success"] is True
        assert result["changes"]["fonts_unified"] >= 0
        
        commit(session_id)
    
    def test_empty_paragraph_removal_e2e(self, test_docx):
        """测试空段落删除 E2E"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = fix_formatting(session_id, options={
            "remove_empty_paragraphs": True,
        })
        
        assert result["success"] is True
        assert "empty_paragraphs_removed" in result["changes"]
        
        commit(session_id)


class TestStyleConvertE2E:
    """E2E测试：风格转换工作流"""
    
    def test_to_business_report(self, test_docx):
        """测试转换为商务报告风格"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        # 应用商务报告模板
        result = apply_style_template(session_id, "business_report")
        
        assert result["success"] is True
        assert "margins" in result["changes"]
        assert "formatting" in result["changes"]
        
        commit(session_id)
    
    def test_to_academic_paper(self, test_docx):
        """测试转换为学术论文风格"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = apply_style_template(session_id, "academic_paper")
        
        assert result["success"] is True
        
        commit(session_id)


class TestTableOrganizeE2E:
    """E2E测试：表格整理工作流"""
    
    def test_fix_table_borders_e2e(self, test_docx):
        """测试表格边框修复 E2E"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = fix_table_borders(session_id, border_style="single")
        
        assert result["success"] is True
        assert "tables_processed" in result
        
        commit(session_id)


class TestDocumentMergeE2E:
    """E2E测试：文档合并工作流"""
    
    def test_merge_documents_e2e(self, test_docx, temp_dir):
        """测试文档合并 E2E"""
        output_path = os.path.join(temp_dir, "merged.docx")
        
        result = merge_documents(
            [test_docx, test_docx],
            output_path,
            options={"add_page_break": True}
        )
        
        assert result["success"] is True
        assert os.path.exists(output_path)
        assert result["stats"]["files_merged"] == 2


class TestTextboxExtractE2E:
    """E2E测试：文本框提取工作流"""
    
    def test_detect_textboxes_e2e(self, test_docx):
        """测试文本框检测 E2E"""
        result = detect_textboxes(test_docx)
        
        assert result["success"] is True
        assert "stats" in result
        assert "total_count" in result["stats"]


class TestErrorHandlingE2E:
    """E2E测试：错误处理"""
    
    def test_missing_file_error(self):
        """测试缺失文件错误"""
        result = analyze("nonexistent.docx")
        
        assert result["success"] is False
        assert result["error"]["code"] == "DOC001"
    
    def test_invalid_format_error(self, temp_dir):
        """测试无效格式错误"""
        # 创建非 docx 文件
        invalid_file = os.path.join(temp_dir, "invalid.txt")
        with open(invalid_file, "w") as f:
            f.write("not a docx file")
        
        result = analyze(invalid_file)
        
        assert result["success"] is False
        assert result["error"]["code"] == "DOC002"
    
    def test_invalid_session_error(self):
        """测试无效会话错误"""
        result = fix_formatting("invalid_session", options={})
        
        assert result["success"] is False
        assert result["error"]["code"] == "SES001"
    
    def test_rollback_on_error(self, test_docx):
        """测试错误时回滚"""
        original_size = os.path.getsize(test_docx)
        
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        # 执行修改
        fix_formatting(session_id, options={"default_font": "宋体"})
        
        # 回滚
        rollback(session_id)
        
        # 验证文件大小保持不变
        assert os.path.getsize(test_docx) == original_size
