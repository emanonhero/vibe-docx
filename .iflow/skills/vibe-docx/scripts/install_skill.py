# -*- coding: utf-8 -*-
"""
Vibe Docx - SKILL Installation Script

安装 SKILL 到各种 LLM 工具的命令目录。

Usage:
    # 命令行使用
    python scripts/install_skill.py --target local --tools iflow
    python scripts/install_skill.py --target global --tools iflow,cursor,claude
    
    # 编程使用
    from scripts.install_skill import install_skill, verify_install
    
    result = install_skill("vibe-docx", {
        "target": "global",
        "tools": ["iflow", "cursor"]
    })
"""

import os
import sys
import shutil
import argparse
from pathlib import Path
from typing import Dict, List, Any, Optional


# 工具目录配置
TOOL_CONFIGS = {
    "iflow": {
        "global_dir": "~/.iflow/skills",
        "local_dir": "./.iflow/skills",
        "file_pattern": "{skill_name}/SKILL.md",  # 目录模式
        "description": "iFlow CLI"
    },
    "cursor": {
        "global_dir": "~/.cursor/commands",
        "local_dir": "./.cursor/commands",
        "file_pattern": "{skill_name}.md",  # 单文件模式
        "description": "Cursor IDE"
    },
    "claude": {
        "global_dir": "~/.claude/commands",
        "local_dir": "./.claude/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Claude Code"
    },
    "cline": {
        "global_dir": "~/.cline/commands",
        "local_dir": "./.cline/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Cline VS Code Extension"
    },
    "copilot": {
        "global_dir": "~/.github",
        "local_dir": "./.github",
        "file_pattern": "copilot-instructions.md",  # 固定文件名
        "description": "GitHub Copilot"
    },
    "windsurf": {
        "global_dir": "~/.windsurf/commands",
        "local_dir": "./.windsurf/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Windsurf IDE"
    }
}

# 需要安装的文件
INSTALL_FILES = ["SKILL.md", "references/", "scripts/"]


def get_skill_source_dir() -> Path:
    """获取 SKILL 源目录"""
    # 脚本所在目录的上级目录
    script_dir = Path(__file__).parent
    return script_dir.parent


def expand_path(path: str) -> Path:
    """展开路径（支持 ~ 和环境变量）"""
    return Path(os.path.expanduser(os.path.expandvars(path)))


def install_skill(skill_name: str, options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    安装 SKILL 到指定工具
    
    Args:
        skill_name: SKILL 名称 (如 "vibe-docx")
        options: 安装选项
            - target: "local" | "global" (默认 "local")
            - tools: List[str], 目标工具列表 (默认 ["iflow"])
            - files: List[str], 要安装的文件 (默认 INSTALL_FILES)
            - overwrite: bool, 是否覆盖已有 (默认 False)
    
    Returns:
        {
            "success": bool,
            "installed": list,  # 成功安装的工具
            "failed": list,     # 安装失败的工具
            "details": dict     # 详细信息
        }
    """
    options = options or {}
    target = options.get("target", "local")
    tools = options.get("tools", ["iflow"])
    files = options.get("files", INSTALL_FILES)
    overwrite = options.get("overwrite", False)
    
    source_dir = get_skill_source_dir()
    results = {
        "success": True,
        "installed": [],
        "failed": [],
        "details": {}
    }
    
    for tool in tools:
        if tool not in TOOL_CONFIGS:
            results["failed"].append(tool)
            results["details"][tool] = {"error": f"未知工具: {tool}"}
            continue
        
        config = TOOL_CONFIGS[tool]
        
        # 确定目标目录
        dir_key = "global_dir" if target == "global" else "local_dir"
        dest_base = expand_path(config[dir_key])
        
        # 确定目标路径
        file_pattern = config["file_pattern"].format(skill_name=skill_name)
        if "/" in file_pattern:
            # 目录模式
            dest_dir = dest_base / file_pattern.rsplit("/", 1)[0]
            dest_file = dest_base / file_pattern
        else:
            # 单文件模式
            dest_dir = dest_base
            dest_file = dest_base / file_pattern
        
        try:
            # 创建目录
            dest_dir.mkdir(parents=True, exist_ok=True)
            
            # 复制文件
            for f in files:
                src = source_dir / f
                if not src.exists():
                    continue
                
                if src.is_dir():
                    dst = dest_dir / src.name
                    if dst.exists():
                        if overwrite:
                            shutil.rmtree(dst)
                        else:
                            continue
                    shutil.copytree(src, dst)
                else:
                    dst = dest_dir / src.name
                    if dst.exists() and not overwrite:
                        continue
                    shutil.copy2(src, dst)
            
            results["installed"].append(tool)
            results["details"][tool] = {
                "path": str(dest_file),
                "files_copied": files
            }
            
        except Exception as e:
            results["failed"].append(tool)
            results["details"][tool] = {"error": str(e)}
            results["success"] = False
    
    return results


def verify_install(skill_name: str, tool: str, target: str = "local") -> Dict[str, Any]:
    """
    验证 SKILL 安装是否成功
    
    Args:
        skill_name: SKILL 名称
        tool: 工具名称
        target: "local" | "global"
    
    Returns:
        {
            "installed": bool,
            "skill_file": str,
            "missing_files": list
        }
    """
    if tool not in TOOL_CONFIGS:
        return {
            "installed": False,
            "error": f"未知工具: {tool}"
        }
    
    config = TOOL_CONFIGS[tool]
    dir_key = "global_dir" if target == "global" else "local_dir"
    dest_base = expand_path(config[dir_key])
    file_pattern = config["file_pattern"].format(skill_name=skill_name)
    dest_file = dest_base / file_pattern
    
    # 检查文件是否存在
    missing_files = []
    
    if not dest_file.exists():
        missing_files.append(str(dest_file))
    
    # 检查 references 目录
    if "/" in file_pattern:
        ref_dir = dest_file.parent / "references"
    else:
        ref_dir = dest_base / "references"
    
    # 不强制要求 references 目录
    
    return {
        "installed": len(missing_files) == 0,
        "skill_file": str(dest_file),
        "missing_files": missing_files
    }


def list_supported_tools() -> List[Dict[str, str]]:
    """列出支持的工具"""
    return [
        {"name": name, "description": config["description"]}
        for name, config in TOOL_CONFIGS.items()
    ]


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(
        description="安装 Vibe Docx SKILL 到各种 LLM 工具"
    )
    
    parser.add_argument(
        "--target", "-t",
        choices=["local", "global"],
        default="local",
        help="安装范围: local (项目) 或 global (全局)"
    )
    
    parser.add_argument(
        "--tools", "-T",
        default="iflow",
        help="目标工具，逗号分隔 (如: iflow,cursor,claude)"
    )
    
    parser.add_argument(
        "--verify", "-v",
        action="store_true",
        help="仅验证安装状态"
    )
    
    parser.add_argument(
        "--list-tools", "-l",
        action="store_true",
        help="列出支持的工具"
    )
    
    parser.add_argument(
        "--overwrite", "-o",
        action="store_true",
        help="覆盖已有文件"
    )
    
    args = parser.parse_args()
    
    # 列出工具
    if args.list_tools:
        print("\n支持的工具:")
        print("-" * 40)
        for tool in list_supported_tools():
            print(f"  {tool['name']:<10} - {tool['description']}")
        print()
        return 0
    
    # 解析工具列表
    tools = [t.strip() for t in args.tools.split(",")]
    skill_name = "vibe-docx"
    
    # 验证模式
    if args.verify:
        print(f"\n验证 {skill_name} 安装状态...\n")
        all_installed = True
        for tool in tools:
            result = verify_install(skill_name, tool, args.target)
            status = "✓" if result["installed"] else "✗"
            print(f"  {status} {tool}: {result.get('skill_file', 'N/A')}")
            if not result["installed"]:
                all_installed = False
                if result.get("missing_files"):
                    print(f"    缺少: {result['missing_files']}")
        print()
        return 0 if all_installed else 1
    
    # 安装模式
    print(f"\n安装 {skill_name} SKILL...")
    print(f"目标: {args.target}")
    print(f"工具: {', '.join(tools)}\n")
    
    result = install_skill(skill_name, {
        "target": args.target,
        "tools": tools,
        "overwrite": args.overwrite
    })
    
    # 显示结果
    if result["installed"]:
        print("成功安装到:")
        for tool in result["installed"]:
            details = result["details"][tool]
            print(f"  ✓ {tool}: {details.get('path', 'N/A')}")
    
    if result["failed"]:
        print("\n安装失败:")
        for tool in result["failed"]:
            details = result["details"][tool]
            print(f"  ✗ {tool}: {details.get('error', '未知错误')}")
    
    print()
    return 0 if result["success"] else 1


if __name__ == "__main__":
    sys.exit(main())
