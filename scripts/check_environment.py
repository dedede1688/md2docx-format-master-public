"""验证便携目录的运行路径、依赖、字体和文件清单。"""
from pathlib import Path
import hashlib
import importlib
import json
import platform
import re
import struct
import sys

ROOT = Path(__file__).resolve().parents[1]


def verify_manifest(root):
    root = root.resolve()
    manifest_path = root / 'verification/manifest.json'
    try:
        manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
        files = manifest['files']
        if not isinstance(files, dict) or not files:
            raise ValueError('files 必须是非空的路径与哈希映射')
    except FileNotFoundError:
        return 0, ['缺少完整性清单 verification/manifest.json']
    except (OSError, TypeError, KeyError, ValueError) as exc:
        return 0, [f'无法读取完整性清单 verification/manifest.json：{exc}']

    errors = []
    for relative, expected in files.items():
        if not isinstance(expected, str) or not re.fullmatch(r'[0-9a-f]{64}', expected):
            errors.append('文件校验值无效：' + relative)
            continue
        try:
            path = (root / relative).resolve()
            if Path(relative).is_absolute() or not path.is_relative_to(root):
                errors.append('清单路径超出技能目录：' + relative)
                continue
            with path.open('rb') as stream:
                if hashlib.file_digest(stream, 'sha256').hexdigest() != expected:
                    errors.append('文件校验不一致：' + relative)
        except FileNotFoundError:
            errors.append('缺少包内文件：' + relative)
        except (OSError, ValueError) as exc:
            errors.append(f'无法校验包内文件 {relative}：{exc}')
    return len(files), errors


def inspect_environment():
    errors = []
    modules = {}
    if sys.platform != 'win32' or struct.calcsize('P') != 8:
        errors.append('此便携包仅支持 Windows 64 位。')
    if not Path(sys.executable).resolve().is_relative_to(ROOT / 'runtime'):
        errors.append('请通过包内启动转换.cmd 或 runtime/windows-x64/python.exe 启动。')
    isolated = bool(sys.flags.isolated)
    external_paths = [str(path) for path in sys.path if not Path(path).resolve().is_relative_to(ROOT)]
    if not isolated:
        errors.append('包内 Python 未启用隔离模式。')
    if external_paths:
        errors.append('Python 搜索路径包含技能目录以外的位置。')
    for module in ('docx', 'lxml.etree', 'requests', 'urllib3', 'certifi', 'charset_normalizer', 'idna', 'typing_extensions', 'ssl', 'ctypes'):
        try:
            loaded = importlib.import_module(module)
            path = Path(loaded.__file__).resolve()
            modules[module] = str(path)
            if not path.is_relative_to(ROOT):
                errors.append('依赖来自技能目录外：' + module)
        except Exception as exc:
            errors.append('包内依赖不可用：' + module + '（' + str(exc) + '）')
    try:
        from font_manager import catalog
        fonts = catalog()
        for name in ('仿宋', '思源宋体 CN Heavy', '方正清刻本悦宋简体', 'Consolas', 'Times New Roman'):
            if not any(name in font['aliases'] for font in fonts):
                errors.append('包内缺少字体：' + name)
    except Exception as exc:
        fonts = []
        errors.append('字体检查失败：' + str(exc))
    checked, manifest_errors = verify_manifest(ROOT)
    errors.extend(manifest_errors)
    return {'ok': not errors, 'platform': platform.platform(), 'python': sys.version.split()[0],
            'executable': sys.executable, 'isolated': isolated, 'external_paths': external_paths,
            'modules': modules, 'fonts': len(fonts), 'files_verified': checked, 'errors': errors}


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    result = inspect_environment()
    print(json.dumps(result, ensure_ascii=False))
    raise SystemExit(0 if result['ok'] else 1)
