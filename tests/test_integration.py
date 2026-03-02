# -*- coding: utf-8 -*-
"""
Integration tests for vibe_docx session flow.
"""

import pytest
import tempfile
import os
import shutil
from scripts.builder import (
    begin_session,
    commit,
    rollback,
    fix_formatting,
    fix_table_borders,
    fix_page_setup,
    apply_style_template,
    SessionManager,
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


class TestSessionLifecycle:
    """测试会话生命周期"""
    
    def test_full_session_flow(self, test_docx):
        """测试完整会话流程：开始 -> 修改 -> 提交"""
        # 开始会话
        session = begin_session(test_docx, backup=True)
        assert session["success"] is True
        session_id = session["session_id"]
        
        # 执行修改
        result = fix_formatting(session_id, options={
            "default_font": "宋体",
            "remove_empty_paragraphs": True,
        })
        assert result["success"] is True
        
        # 提交
        commit_result = commit(session_id)
        assert commit_result["success"] is True
    
    def test_session_with_rollback(self, test_docx):
        """测试会话回滚流程"""
        original_size = os.path.getsize(test_docx)
        
        # 开始会话
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        # 执行修改
        fix_formatting(session_id, options={"default_font": "楷体"})
        
        # 回滚
        rollback_result = rollback(session_id)
        assert rollback_result["success"] is True
        
        # 验证文件未改变
        assert os.path.getsize(test_docx) == original_size
    
    def test_multiple_operations_in_session(self, test_docx):
        """测试同一会话中多个操作"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        # 执行多个修改
        r1 = fix_formatting(session_id, options={"default_font": "宋体"})
        assert r1["success"] is True
        
        r2 = fix_table_borders(session_id)
        assert r2["success"] is True
        
        r3 = fix_page_setup(session_id)
        assert r3["success"] is True
        
        # 提交
        commit_result = commit(session_id)
        assert commit_result["success"] is True
    
    def test_session_without_backup(self, test_docx):
        """测试无备份会话"""
        session = begin_session(test_docx, backup=False)
        assert session["success"] is True
        assert session["backup_path"] is None
        
        # 清理
        commit(session["session_id"])


class TestSessionManager:
    """测试会话管理器"""
    
    def test_manager_stats(self, test_docx):
        """测试会话统计"""
        manager = SessionManager()
        
        # 创建一个会话
        s1 = manager.create(test_docx, backup=True)
        
        stats = manager.get_stats()
        assert stats["active_sessions"] == 1
        assert stats["backups"] == 1
        
        # 清理
        manager.close(s1)
    
    def test_duplicate_session_blocked(self, test_docx):
        """测试同一文件重复会话被阻止"""
        manager = SessionManager()
        
        # 创建第一个会话
        s1 = manager.create(test_docx, backup=True)
        
        # 尝试创建第二个会话应该失败
        with pytest.raises(Exception):  # SessionError
            manager.create(test_docx, backup=True)
        
        # 清理
        manager.close(s1)
    
    def test_invalid_session_operation(self):
        """测试无效会话操作"""
        result = fix_formatting("invalid_session_id", options={})
        assert result["success"] is False
        assert result["error"]["code"] == "SES001"


class TestStyleTemplateIntegration:
    """测试样式模板集成"""
    
    def test_business_report_template(self, test_docx):
        """测试商务报告模板应用"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = apply_style_template(session_id, "business_report")
        assert result["success"] is True
        assert result["template"] == "business_report"
        
        commit(session_id)
    
    def test_academic_paper_template(self, test_docx):
        """测试学术论文模板应用"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = apply_style_template(session_id, "academic_paper")
        assert result["success"] is True
        
        commit(session_id)
    
    def test_template_with_custom_override(self, test_docx):
        """测试模板自定义覆盖"""
        session = begin_session(test_docx, backup=True)
        session_id = session["session_id"]
        
        result = apply_style_template(
            session_id,
            "business_report",
            custom_options={"line_spacing": 2.0}
        )
        assert result["success"] is True
        
        commit(session_id)
