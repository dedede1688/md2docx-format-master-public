"""格式大师的独立命令行入口。"""
import argparse
import codecs
import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import zipfile

sys.dont_write_bytecode = True
skill_root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(skill_root / 'dependencies'))


def emit(data):
    print(json.dumps(data, ensure_ascii=False))


def read_source(path, encoding):
    data = path.read_bytes()
    if encoding.lower().replace('_', '-') in ('utf-8', 'utf8', 'utf-8-sig'):
        if data.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
            encoding = 'utf-32'
        elif data.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
            encoding = 'utf-16'
        else:
            encoding = 'utf-8-sig'
    try:
        return data.decode(encoding)
    except UnicodeDecodeError as exc:
        raise ValueError(f'无法以 {encoding} 读取 {path.name}。请通过 --encoding 指定源文件编码。') from exc


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        emit({'ok': False, 'error': message})
        raise SystemExit(2)


def main():
    parser = JsonArgumentParser(description='001-md2docx · 格式大师')
    source = parser.add_mutually_exclusive_group()
    source.add_argument('--input')
    source.add_argument('--text')
    destination = parser.add_mutually_exclusive_group()
    destination.add_argument('--output')
    destination.add_argument('--output-dir')
    parser.add_argument('--check', action='store_true')
    parser.add_argument('--base-dir')
    parser.add_argument('--config')
    parser.add_argument('--encoding', help='源文件编码，例如 gb18030')
    parser.add_argument('--template')
    parser.add_argument('--template-mode', choices=('append', 'styles', 'placeholders'))
    parser.add_argument('--no-cover', action='store_true')
    parser.add_argument('--toc', action='store_true')
    parser.add_argument('--no-h1-page-break', action='store_true', help='一级标题连续排版')
    parser.add_argument('--title')
    parser.add_argument('--header')
    parser.add_argument('--body-font-size', type=float)
    parser.add_argument('--overwrite', action='store_true')
    parser.add_argument('--no-embed-fonts', action='store_true', help='不嵌入字体，仅用于用户明确要求减小文件时')
    parser.add_argument('--offline', action='store_true', help='不请求网络图片')
    parser.add_argument('--strict', action='store_true', help='检测到内容或格式未能完整转换时不交付输出')
    args = parser.parse_args()
    missing = [package for package, module in [('python-docx', 'docx'), ('requests', 'requests')]
               if importlib.util.find_spec(module) is None]
    if args.check:
        from check_environment import inspect_environment
        result = inspect_environment()
        result['missing'] = missing
        emit(result)
        return 0 if result['ok'] else 1
    temporary = None
    generator = None
    diagnostics = []
    try:
        if missing:
            raise RuntimeError('技能包缺少依赖：' + '、'.join(missing) + '。请重新复制完整技能文件夹，无需下载安装。')
        from config import ConversionConfig
        from markdown_parser import MarkdownParser
        from word_generator import WordGenerator
        if args.input is None and args.text is None:
            raise ValueError('请提供 --input 文件或 --text 文本')
        input_path = Path(args.input).expanduser().resolve() if args.input else None
        if input_path and input_path.suffix.lower() not in ('.md', '.markdown', '.txt'):
            raise ValueError('输入仅支持 .md、.markdown、.txt 文件')
        config_data = json.loads((skill_root / 'config/default.json').read_text(encoding='utf-8'))
        if args.config:
            config_data.update(json.loads(Path(args.config).read_text(encoding='utf-8-sig')))
        config_data.setdefault('header_text', '')
        config_data.setdefault('margin_left', 2.5)
        config_data.setdefault('margin_right', 2.5)
        config = ConversionConfig.from_dict(config_data)
        if args.header is not None:
            config.header_text = args.header
        if args.no_cover:
            config.generate_cover = False
        if args.toc:
            config.generate_toc = True
        if args.no_h1_page_break:
            config.h1_page_break = False
        if args.title is not None:
            config.cover_title = args.title
        if args.body_font_size is not None:
            if not 6 <= args.body_font_size <= 72:
                raise ValueError('正文字号应在 6 至 72 磅之间')
            config.body_font_size = args.body_font_size
        if not 0 < config.body_line_spacing <= 5:
            raise ValueError('正文行距应大于 0 且不超过 5 倍')
        if min(config.margin_left, config.margin_right, config.margin_top, config.margin_bottom) < 0:
            raise ValueError('页边距不可为负数')
        template = args.template or config.template_path
        template_mode = args.template_mode or config.template_mode
        if template_mode not in ('append', 'styles', 'placeholders'):
            raise ValueError('模板模式须为 append、styles 或 placeholders')
        if template:
            template = Path(template).expanduser().resolve()
            if not template.is_file():
                raise FileNotFoundError('模板文件不存在：' + str(template))
        elif template_mode != 'append':
            raise ValueError('请为模板模式指定 --template 文件')
        if input_path:
            content = read_source(input_path, args.encoding or config.encoding)
            base_dir = input_path.parent
            default_name = input_path.stem + '.docx'
        else:
            content = args.text
            base_dir = Path(args.base_dir or os.getcwd()).resolve()
            default_name = '转换文档.docx'
        if not content.strip():
            raise ValueError('输入内容为空')
        output = (Path(args.output).expanduser() if args.output else
                  Path(args.output_dir).expanduser() / default_name if args.output_dir else
                  base_dir / default_name).resolve()
        if output.suffix.lower() != '.docx':
            raise ValueError('输出文件必须使用 .docx 扩展名')
        if template == output or (template and output.is_file() and template.samefile(output)):
            raise ValueError('输出文件与模板相同，请另存为不同的文件。')
        if output.exists() and not args.overwrite:
            raise FileExistsError('输出文件已存在；请改名或明确使用 --overwrite：' + str(output))
        markdown_parser = MarkdownParser()
        elements = markdown_parser.parse(content)
        diagnostics.extend(markdown_parser.diagnostics)
        if not elements:
            raise ValueError('输入没有可转换的正文。')
        if args.strict and diagnostics:
            raise ValueError('严格模式检测到转换问题，未写入输出文件。')
        generator = WordGenerator(template_path=template, config=config, template_mode=template_mode)
        generator.base_dir = str(base_dir)
        generator.offline = args.offline
        output.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(suffix='.docx', dir=output.parent)
        os.close(fd)
        generator.generate(elements, temporary)
        diagnostics.extend(generator.diagnostics)
        if args.strict and diagnostics:
            raise ValueError('严格模式检测到转换问题，未写入输出文件。')
        embedded = []
        if not args.no_embed_fonts:
            from font_embedding import embed_fonts
            embedded, font_warnings = embed_fonts(temporary)
            diagnostics.extend({'code': 'font_embedding', 'message': message, 'line': None}
                               for message in font_warnings)
            if args.strict and font_warnings:
                raise ValueError('严格模式检测到字体嵌入问题，未写入输出文件。')
        with zipfile.ZipFile(temporary) as archive:
            if archive.testzip() is not None or 'word/document.xml' not in archive.namelist():
                raise RuntimeError('生成的 Word 文件完整性检查失败')
        from docx import Document
        checked = Document(temporary)
        if args.overwrite:
            os.replace(temporary, output)
        elif os.name == 'nt':
            os.rename(temporary, output)
        else:
            os.link(temporary, output)
            os.unlink(temporary)
        temporary = None
        warnings = [f"第 {item['line']} 行：{item['message']}" if item['line'] else item['message']
                    for item in diagnostics]
        if config.generate_toc:
            warnings.append('请在 Word 中更新目录域。')
        emit({'ok': True, 'output': str(output), 'bytes': output.stat().st_size,
              'elements': len(elements), 'tables': len(checked.tables),
              'images': len(checked.inline_shapes), 'warnings': warnings,
              'diagnostics': diagnostics,
              'embedded_fonts': embedded,
              'font_note': '已嵌入包内所用字体；实际显示仍需由支持嵌入字体的阅读器验证。' if embedded else '未嵌入字体，实际显示取决于阅读机器字体。'})
        return 0
    except Exception as exc:
        if generator:
            diagnostics.extend(item for item in generator.diagnostics if item not in diagnostics)
        emit({'ok': False, 'error': str(exc), 'diagnostics': diagnostics})
        return 1
    finally:
        if temporary and Path(temporary).exists():
            Path(temporary).unlink()
        if generator:
            for path in generator.downloaded_images:
                Path(path).unlink(missing_ok=True)


if __name__ == '__main__':
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
    raise SystemExit(main())
