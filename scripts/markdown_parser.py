"""
Markdown 解析器模块
负责解析 Markdown 内容并转换为结构化数据
"""

import re
from dataclasses import dataclass, field
from typing import List, Optional, Union
from enum import Enum

FENCE = re.compile(r'^ {0,3}(`{3,}|~{3,})(.*)$')
LIST_MARKER = re.compile(r'^([-*+]|[0-9]{1,9}[.)])\s+(.*)$')
SETEXT = re.compile(r'^ {0,3}(=+|-+)\s*$')
FRONT_MATTER_KEY = re.compile(r"""(?:[\w][\w .-]*|"[^"]+"|'[^']+'):(?:[ \t]+(.*))?""")
LINK_LABEL = r'\[((?:\\.|[^\]\\\n])*)\]'
LINK_TARGET = r'\((<(?:\\.|[^<>\\\n])*>|(?:\\.|[^\s()\\]|\((?:\\.|[^()\\\n])*\))+)(?:[ \t]+"([^"\n]*)")?[ \t]*\)'
INLINE_LINK = LINK_LABEL + LINK_TARGET
INLINE_IMAGE = '!' + INLINE_LINK
# 兼容旧版允许的独立图片路径，其中空格直接写在括号内。
LEGACY_IMAGE = re.compile(r'!\[([^\]]*)\]\(([^)]+?)(?:\s+"([^"]+)")?\)')
ESCAPED_PUNCTUATION = re.compile(r'\\([\x21-\x2f\x3a-\x40\x5b-\x60\x7b-\x7e])')


class ElementType(Enum):
    """元素类型枚举"""
    HEADING = "heading"
    PARAGRAPH = "paragraph"
    LIST_ITEM = "list_item"
    CODE_BLOCK = "code_block"
    TABLE = "table"
    IMAGE = "image"
    HORIZONTAL_RULE = "horizontal_rule"
    BLOCKQUOTE = "blockquote"


@dataclass
class TableData:
    """表格数据结构"""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    alignments: List[str] = field(default_factory=list)


@dataclass
class CodeBlock:
    """代码块数据结构"""
    language: str = ""
    code: str = ""


@dataclass
class ImageData:
    """图片数据结构"""
    alt_text: str = ""
    url: str = ""


@dataclass
class DocumentElement:
    """文档元素基类"""
    type: ElementType
    content: Union[str, TableData, CodeBlock, ImageData]
    level: int = 0
    extra_data: dict = field(default_factory=dict)
    line: int = 0


class MarkdownParser:
    """Markdown 解析器"""
    
    def __init__(self):
        self._list_group_id = 0
        self.diagnostics = []
    
    def parse(self, markdown_text: str) -> List[DocumentElement]:
        """解析 Markdown 文本"""
        elements = []
        self._list_group_id = 0
        self.diagnostics.clear()
        lines = markdown_text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
        i = self._body_start(lines)
        
        while i < len(lines):
            line = lines[i]
            
            if not line.strip():
                i += 1
                continue
            
            # 检查标题
            if re.match(r'^#{1,6}\s+', line.strip()):
                heading = self._parse_heading(line)
                heading.line = i + 1
                self._check_unsupported_inline(heading.content, i + 1)
                elements.append(heading)
                i += 1
                continue
            
            # 检查代码块
            if FENCE.match(line):
                block, consumed = self._parse_code_block(lines, i)
                elements.append(DocumentElement(type=ElementType.CODE_BLOCK, content=block, line=i + 1))
                i += consumed
                continue
            
            # 检查表格
            if self._table_start(lines, i):
                table, consumed = self._parse_table(lines, i)
                elements.append(DocumentElement(type=ElementType.TABLE, content=table, line=i + 1))
                i += consumed
                continue
            
            # 检查水平线
            if self._is_horizontal_rule(line):
                elements.append(DocumentElement(type=ElementType.HORIZONTAL_RULE, content="", line=i + 1))
                i += 1
                continue
            
            # 检查引用块
            if line.strip().startswith('>'):
                blockquote, consumed = self._parse_blockquote(lines, i)
                self._check_unsupported_inline(blockquote, i + 1)
                elements.append(DocumentElement(type=ElementType.BLOCKQUOTE, content=blockquote, line=i + 1))
                i += consumed
                continue
            
            # 检查列表项
            if self._is_list_item(line):
                list_items, consumed = self._parse_list(lines, i)
                elements.extend(list_items)
                i += consumed
                continue
            
            image = self._standalone_image(line.strip())
            if image:
                data = ImageData(
                    alt_text=ESCAPED_PUNCTUATION.sub(r'\1', image.group(1)),
                    url=ESCAPED_PUNCTUATION.sub(r'\1', image.group(2)),
                )
                elements.append(DocumentElement(type=ElementType.IMAGE, content=data, line=i + 1))
                i += 1
                continue
            
            # 默认为段落
            paragraph, consumed, level = self._parse_paragraph(lines, i)
            kind = ElementType.HEADING if level else ElementType.PARAGRAPH
            self._check_unsupported_inline(paragraph, i + 1)
            elements.append(DocumentElement(type=kind, content=paragraph, level=level, line=i + 1))
            i += consumed
        
        return elements

    @staticmethod
    def _body_start(lines: List[str]) -> int:
        """跳过文首形如 YAML 键值表的笔记属性，保留原文件行号。"""
        if lines[0].removeprefix('\ufeff').rstrip() != '---':
            return 0

        has_key = False
        allows_sequence = False
        for index, raw in enumerate(lines[1:], start=1):
            line = raw.rstrip()
            if line == '---':
                return index + 1 if has_key else 0
            if not line or line.lstrip().startswith('#'):
                continue
            if line[0].isspace():
                if not has_key:
                    return 0
                continue
            if allows_sequence and line.startswith('- '):
                continue
            entry = FRONT_MATTER_KEY.fullmatch(line)
            if entry is None:
                return 0
            has_key = True
            value = entry.group(1)
            allows_sequence = not value or value.startswith('#')
        return 0

    def _warn(self, code, message, line):
        self.diagnostics.append({'code': code, 'message': message, 'line': line})

    def _check_unsupported_inline(self, text, line):
        # 代码跨度和转义后的标记按普通内容保留。
        text = re.sub(r'(?<!`)(`+)(?!`).+?(?<!`)\1(?!`)', '', text)
        text = ESCAPED_PUNCTUATION.sub('', text)
        text = re.sub(INLINE_LINK, lambda match: match.group(1), text)
        if re.search(r'\[\^[^\]]+\]', text):
            self._warn('unsupported_footnote', '脚注标记已按文本保留。', line)
        if re.search(r'\[[^\]\n]+\]\[[^\]\n]*\]', text):
            self._warn('unsupported_reference_link', '引用式链接或图片已按文本保留。', line)
        if re.search(r'(?<!\w)\$\$[\s\S]+?\$\$|\$(?!\s|\$)[^$\n]+?(?<!\s)\$(?![\d$])', text):
            self._warn('unsupported_math', '公式已按文本保留。', line)
        if re.search(r'<\/?(?:div|table|tr|td|details|summary|iframe|ul|ol|li|img|br|p)\b', text, re.I):
            self._warn('unsupported_html', 'HTML 标签已按文本保留。', line)
    
    def _parse_heading(self, line: str) -> DocumentElement:
        """解析标题"""
        match = re.match(r'^(#{1,6})\s+(.+)$', line.strip())
        if match:
            level = len(match.group(1))
            text = re.sub(r'\s+#+\s*$', '', match.group(2))
            return DocumentElement(type=ElementType.HEADING, content=text, level=level)
        return DocumentElement(type=ElementType.HEADING, content=line.strip(), level=1)
    
    def _parse_code_block(self, lines: List[str], start_index: int) -> tuple:
        """解析代码块"""
        opening = FENCE.match(lines[start_index])
        marker, info = opening.groups()
        language = info.strip().split(maxsplit=1)[0] if info.strip() else ''
        closing = re.compile(r'^ {0,3}' + re.escape(marker[0]) + '{' + str(len(marker)) + r',}[ \t]*$')
        if language.lower() in ('mermaid', 'math', 'latex'):
            self._warn('unsupported_code_rendering', f'{language} 已按代码块保留。', start_index + 1)
        
        code_lines = []
        i = start_index + 1
        while i < len(lines):
            if closing.fullmatch(lines[i]):
                return CodeBlock(language=language, code='\n'.join(code_lines)), i - start_index + 1
            code_lines.append(lines[i])
            i += 1
        
        self._warn('unclosed_code_fence', '代码块未闭合，保留到文件末尾。', start_index + 1)
        return CodeBlock(language=language, code='\n'.join(code_lines)), i - start_index

    @staticmethod
    def _table_cells(line):
        parts = re.split(r'((?:\\\\)*)(?<!\\)\|', line.strip())
        # 将分隔线前的偶数个反斜线留在前一格。
        cells = [parts[i] + (parts[i + 1] if i + 1 < len(parts) else '')
                 for i in range(0, len(parts), 2)]
        if len(cells) > 1 and not cells[0]:
            cells.pop(0)
        if len(cells) > 1 and not cells[-1]:
            cells.pop()
        return [cell.strip() for cell in cells]

    @classmethod
    def _table_start(cls, lines, index):
        if index + 1 >= len(lines) or '|' not in lines[index]:
            return False
        separators = cls._table_cells(lines[index + 1])
        return (len(separators) == len(cls._table_cells(lines[index]))
                and all(re.fullmatch(r':?-{3,}:?', cell) for cell in separators))
    
    def _parse_table(self, lines: List[str], start_index: int) -> tuple:
        """解析已识别的表格。"""
        headers = self._table_cells(lines[start_index])
        separators = self._table_cells(lines[start_index + 1])
        
        alignments = []
        for sep in separators:
            if sep.startswith(':') and sep.endswith(':'):
                alignments.append('center')
            elif sep.endswith(':'):
                alignments.append('right')
            else:
                alignments.append('left')
        
        rows = []
        i = start_index + 2
        while i < len(lines):
            line = lines[i].strip()
            if (not line or '|' not in line or self._table_start(lines, i)
                    or FENCE.match(lines[i]) or re.match(r'^#{1,6}\s', line)
                    or self._is_list_item(line) or line.startswith('>')):
                break
            
            cells = self._table_cells(line)
            for cell in cells:
                self._check_unsupported_inline(cell, i + 1)
            
            if len(cells) == len(headers):
                rows.append(cells)
            elif len(cells) < len(headers):
                # 列数不足，补空列
                rows.append(cells + [''] * (len(headers) - len(cells)))
                self._warn('table_missing_cells', '表格列数少于表头，已补空格。', i + 1)
            else:
                rows.append(cells[:len(headers) - 1] + [r' \| '.join(cells[len(headers) - 1:])])
                self._warn('table_extra_cells', '多出的表格单元格内容已合并到末列。', i + 1)
            i += 1
        
        return TableData(headers=headers, rows=rows, alignments=alignments), i - start_index
    
    def _is_list_item(self, line: str) -> bool:
        """检查是否为列表项"""
        return LIST_MARKER.match(line.strip()) is not None
    
    def _parse_list(self, lines: List[str], start_index: int) -> tuple:
        """解析列表"""
        items = []
        self._list_group_id += 1
        group_id = self._list_group_id
        i = start_index
        indent_stack = [len(lines[start_index]) - len(lines[start_index].lstrip())]
        
        while i < len(lines):
            line = lines[i]
            
            if not line.strip():
                i += 1
                break
            
            current_indent = len(line) - len(line.lstrip())
            stripped = line.strip()
            
            if not self._is_list_item(stripped):
                break
            
            while len(indent_stack) > 1 and current_indent < indent_stack[-1]:
                indent_stack.pop()
            if current_indent > indent_stack[-1]:
                indent_stack.append(current_indent)
            level = min(len(indent_stack), 9)
            if len(indent_stack) > 9:
                self._warn('list_depth', '列表缩进超过 Word 的九级，已使用第九级。', i + 1)
            
            marker, content = LIST_MARKER.match(stripped).groups()
            is_ordered = marker[0].isdigit()
            self._check_unsupported_inline(content, i + 1)
            
            items.append(DocumentElement(
                type=ElementType.LIST_ITEM,
                content=content,
                level=level,
                extra_data={'ordered': is_ordered, 'group_id': group_id,
                            'number': int(marker[:-1]) if is_ordered else None,
                            'delimiter': marker[-1]},
                line=i + 1,
            ))
            
            i += 1
        
        return items, i - start_index
    
    def _is_horizontal_rule(self, line: str) -> bool:
        """检查是否为水平线"""
        stripped = line.strip()
        return (stripped == '---' or 
                stripped == '***' or 
                stripped == '___' or
                re.match(r'^(-|\*|_)\s*\1\s*\1+$', stripped))
    
    def _parse_blockquote(self, lines: List[str], start_index: int) -> tuple:
        """解析引用块"""
        quote_lines = []
        i = start_index
        
        while i < len(lines):
            line = lines[i].strip()
            if line.startswith('>'):
                quote_lines.append(line[1:].strip())
                i += 1
            elif not line:
                break
            else:
                break
        
        return '\n'.join(quote_lines), i - start_index
    
    @staticmethod
    def _standalone_image(line: str):
        return re.fullmatch(INLINE_IMAGE, line) or LEGACY_IMAGE.fullmatch(line)
    
    def _parse_paragraph(self, lines: List[str], start_index: int) -> tuple:
        """解析段落"""
        paragraph_lines = []
        level = 0
        i = start_index
        
        while i < len(lines):
            line = lines[i].strip()
            
            if not line:
                i += 1
                break
            
            underline = SETEXT.fullmatch(lines[i])
            if i > start_index and underline:
                level = 1 if underline.group(1)[0] == '=' else 2
                i += 1
                break

            if (re.match(r'^#{1,6}\s+', line) or
                FENCE.match(lines[i]) or
                self._is_list_item(line) or
                self._is_horizontal_rule(line) or
                line.startswith('>')):
                break
            
            if i > start_index and (self._table_start(lines, i)
                                    or self._standalone_image(line)):
                break
            raw = lines[i]
            hard_break = raw.endswith('  ') or re.search(r'(?<!\\)(?:\\\\)*\\$', raw) is not None
            paragraph_lines.append((line, hard_break, raw.endswith('\\')))
            i += 1
        
        # 智能拼接：中文字符之间不加空格，英文之间加空格
        result, hard_break, slash_break = paragraph_lines[0]
        for line, next_break, next_slash in paragraph_lines[1:]:
            if hard_break:
                if slash_break:
                    result = result[:-1]
                result += '\n' + line
                hard_break = next_break
                slash_break = next_slash
                continue
            if result and line:
                last_char = result[-1]
                first_char = line[0]
                # 检测是否为中文字符（含全角标点）
                def _is_cjk(c):
                    return '\u4e00' <= c <= '\u9fff' or '\u3000' <= c <= '\u303f' or '\uff00' <= c <= '\uffef'
                if _is_cjk(last_char) or _is_cjk(first_char):
                    result += line
                else:
                    result += ' ' + line
            else:
                result += line
            hard_break = next_break
            slash_break = next_slash
        return result, i - start_index, level
