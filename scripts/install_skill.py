# -*- coding: utf-8 -*-
"""
Vibe Docx - SKILL Installation Script

兼容入口：保留原有 `python scripts/install_skill.py ...` 用法，
内部实现迁移到 `vibe_docx.cli`，以支持安装后全局命令。
"""

from pathlib import Path
import sys

# 确保可导入项目内的 vibe_docx 包（支持从任意 cwd 以绝对路径执行此脚本）
project_root = Path(__file__).resolve().parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from vibe_docx.cli import (  # noqa: F401
    INSTALL_FILES,
    TOOL_CONFIGS,
    install_skill,
    list_supported_tools,
    skill_cli_main,
    verify_install,
)


if __name__ == "__main__":
    raise SystemExit(skill_cli_main())
