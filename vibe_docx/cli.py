# -*- coding: utf-8 -*-
"""vibe-docx 命令行入口。"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from importlib import resources
from pathlib import Path
from typing import Any, Dict, List, Optional

from vibe_docx.version import __version__


TOOL_CONFIGS = {
    "iflow": {
        "global_dir": "~/.iflow/skills",
        "local_dir": "./.iflow/skills",
        "file_pattern": "{skill_name}/SKILL.md",
        "description": "iFlow CLI",
    },
    "cursor": {
        "global_dir": "~/.cursor/commands",
        "local_dir": "./.cursor/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Cursor IDE",
    },
    "claude": {
        "global_dir": "~/.claude/commands",
        "local_dir": "./.claude/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Claude Code",
    },
    "cline": {
        "global_dir": "~/.cline/commands",
        "local_dir": "./.cline/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Cline VS Code Extension",
    },
    "copilot": {
        "global_dir": "~/.github",
        "local_dir": "./.github",
        "file_pattern": "copilot-instructions.md",
        "description": "GitHub Copilot",
    },
    "windsurf": {
        "global_dir": "~/.windsurf/commands",
        "local_dir": "./.windsurf/commands",
        "file_pattern": "{skill_name}.md",
        "description": "Windsurf IDE",
    },
}

INSTALL_FILES = ["SKILL.md", "references/", "scripts/"]


def expand_path(path: str) -> Path:
    return Path(os.path.expanduser(os.path.expandvars(path)))


def _candidate_source_dirs() -> List[Path]:
    candidates: List[Path] = []

    env_source = os.environ.get("VIBE_DOCX_SKILL_SOURCE")
    if env_source:
        candidates.append(Path(env_source))

    package_dir = Path(__file__).resolve().parent
    candidates.append(package_dir.parent)

    try:
        package_root = Path(resources.files("vibe_docx"))
        candidates.append(package_root)
        candidates.append(package_root.parent)
    except Exception:
        pass

    return candidates


def get_skill_source_dir() -> Path:
    required = ["SKILL.md"]
    for candidate in _candidate_source_dirs():
        if all((candidate / item).exists() for item in required):
            return candidate

    searched = "\n  - ".join(str(path) for path in _candidate_source_dirs())
    raise FileNotFoundError(
        "未找到 SKILL 源目录（缺少 SKILL.md）。请设置环境变量 VIBE_DOCX_SKILL_SOURCE 指向源码目录。\n"
        f"已搜索路径:\n  - {searched}"
    )


def install_skill(skill_name: str, options: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    options = options or {}
    target = options.get("target", "local")
    tools = options.get("tools", ["iflow"])
    files = options.get("files", INSTALL_FILES)
    overwrite = options.get("overwrite", False)

    source_dir = get_skill_source_dir()
    results: Dict[str, Any] = {
        "success": True,
        "installed": [],
        "failed": [],
        "details": {},
    }

    for tool in tools:
        if tool not in TOOL_CONFIGS:
            results["failed"].append(tool)
            results["details"][tool] = {"error": f"未知工具: {tool}"}
            results["success"] = False
            continue

        config = TOOL_CONFIGS[tool]
        dir_key = "global_dir" if target == "global" else "local_dir"
        dest_base = expand_path(config[dir_key])

        file_pattern = config["file_pattern"].format(skill_name=skill_name)
        if "/" in file_pattern:
            dest_dir = dest_base / file_pattern.rsplit("/", 1)[0]
            dest_file = dest_base / file_pattern
        else:
            dest_dir = dest_base
            dest_file = dest_base / file_pattern

        copied_files: List[str] = []

        try:
            dest_dir.mkdir(parents=True, exist_ok=True)

            for item in files:
                src = source_dir / item
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
                copied_files.append(item)

            results["installed"].append(tool)
            results["details"][tool] = {
                "path": str(dest_file),
                "files_copied": copied_files,
            }
        except Exception as exc:
            results["failed"].append(tool)
            results["details"][tool] = {"error": str(exc)}
            results["success"] = False

    return results


def verify_install(skill_name: str, tool: str, target: str = "local") -> Dict[str, Any]:
    if tool not in TOOL_CONFIGS:
        return {"installed": False, "error": f"未知工具: {tool}"}

    config = TOOL_CONFIGS[tool]
    dir_key = "global_dir" if target == "global" else "local_dir"
    dest_base = expand_path(config[dir_key])
    file_pattern = config["file_pattern"].format(skill_name=skill_name)
    dest_file = dest_base / file_pattern

    missing_files = []
    if not dest_file.exists():
        missing_files.append(str(dest_file))

    return {
        "installed": len(missing_files) == 0,
        "skill_file": str(dest_file),
        "missing_files": missing_files,
    }


def list_supported_tools() -> List[Dict[str, str]]:
    return [
        {"name": name, "description": config["description"]}
        for name, config in TOOL_CONFIGS.items()
    ]


def skill_cli_main() -> int:
    parser = argparse.ArgumentParser(description="安装 Vibe Docx SKILL 到各种 LLM 工具")
    parser.add_argument("--target", "-t", choices=["local", "global"], default="local")
    parser.add_argument("--tools", "-T", default="iflow")
    parser.add_argument("--verify", "-v", action="store_true")
    parser.add_argument("--list-tools", "-l", action="store_true")
    parser.add_argument("--overwrite", "-o", action="store_true")

    args = parser.parse_args()

    if args.list_tools:
        print("\n支持的工具:")
        print("-" * 40)
        for tool in list_supported_tools():
            print(f"  {tool['name']:<10} - {tool['description']}")
        print()
        return 0

    tools = [item.strip() for item in args.tools.split(",") if item.strip()]
    skill_name = "vibe-docx"

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

    print(f"\n安装 {skill_name} SKILL...")
    print(f"目标: {args.target}")
    print(f"工具: {', '.join(tools)}\n")

    result = install_skill(
        skill_name,
        {
            "target": args.target,
            "tools": tools,
            "overwrite": args.overwrite,
        },
    )

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


def version_cli_main() -> int:
    print(__version__)
    return 0


if __name__ == "__main__":
    sys.exit(skill_cli_main())
