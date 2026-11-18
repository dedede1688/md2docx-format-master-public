"""将完整技能复制到用户明确指定的本地 AI 技能目录。"""
import argparse
from pathlib import Path
import shutil
import sys


def main():
    parser = argparse.ArgumentParser(description='安装 001-md2docx 到指定的本地 AI 技能目录')
    parser.add_argument('--target', required=True, help='AI 客户端的技能父目录，例如 D:/项目/.claude/skills')
    args = parser.parse_args()
    source = Path(__file__).resolve().parents[1]
    target = Path(args.target).expanduser().resolve() / source.name
    if source == target:
        print('该目录已是技能目录，无需复制。')
        return 0
    if source in target.parents or target in source.parents:
        print('安装目标不能与来源目录嵌套。', file=sys.stderr)
        return 1
    if target.exists():
        print('目标已存在，请检查后选择其他目录：' + str(target), file=sys.stderr)
        return 1
    try:
        shutil.copytree(source, target, ignore=shutil.ignore_patterns('__pycache__', '*.pyc', '.runtime', '.git'))
    except OSError as exc:
        print('安装失败：' + str(exc), file=sys.stderr)
        return 1
    print('已复制技能：' + str(target))
    print('环境检查：' + str(target / '检查环境.cmd'))
    print('运行环境、依赖和字体均已包含，无需另行安装。')
    print('请让本地 AI 加载该目录中的 SKILL.md；自动触发取决于客户端支持。')
    return 0


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
