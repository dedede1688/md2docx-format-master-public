"""
Word 文档生成器模块
负责将解析后的 Markdown 元素转换为 Word 文档格式
"""

import os
import re
import tempfile
from copy import deepcopy
from urllib.parse import urlsplit
from typing import List, Optional
from docx import Document

from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.section import Section
from config import ConversionConfig
from markdown_parser import (
    DocumentElement, ElementType, TableData, CodeBlock, 
    ESCAPED_PUNCTUATION, INLINE_IMAGE, INLINE_LINK,
)

# ==================== 字体配置规范 ====================
# 基于调试经验总结的关键发现：
# 1. Run 层级优先级 > 样式层级 - run 对象的字体设置会覆盖段落样式
# 2. 中文字体必须同时设置四个属性才能确保在所有场景下生效：
#    - font.name: 字体族名
#    - w:eastAsia: 东亚字符（中日韩）
#    - w:ascii: ASCII 字符（英文数字）
#    - w:hAnsi: 非 ASCII 字符（阿拉伯文、希伯来文等）
# 3. 优先使用中文名称以确保 Word 界面显示正确
# ======================================================

FONT_CONFIG = {
    'h1': {'zh': '思源宋体 CN Heavy', 'en': 'Source Han Serif CN Heavy'},
    'h2_h6': {'zh': '方正清刻本悦宋简体', 'en': '方正清刻本悦宋简体'},
    'body': {'zh': '仿宋', 'en': 'FangSong'},
    'code': {'zh': 'Consolas', 'en': 'Consolas'}
}

class WordGenerator:
    """Word 文档生成器"""
    
    def __init__(self, template_path: Optional[str] = None, config: Optional[ConversionConfig] = None,
                 template_mode: str = 'append'):
        """
        初始化生成器
        
        Args:
            template_path: 可选的 Word 模板文件路径
            config: 可选的转换配置
        """
        self.config = config or ConversionConfig()
        
        if template_path and os.path.exists(template_path):
            self.doc = Document(template_path)
        else:
            self.doc = Document()
        
        self.base_dir = os.getcwd()
        self.diagnostics = []
        self.source_line = None
        self.downloaded_images = set()
        self._cover_created = False
        self._h1_count = 0  # H1 标题计数器
        self._num_instance_counter = 2  # 编号实例计数器（1=无序列表，2=有序列表，从3开始分配）
        self._group_num_id_map = {}  # 列表组ID -> 编号实例ID 的映射
        
        self._list_group = None
        self._list_counters = {}
        self._preserve_template_styles = bool(template_path and template_mode != 'append')
        self._body_marker = None
        if template_path and template_mode == 'styles':
            self.doc._body.clear_content()
        if template_path and template_mode == 'placeholders':
            markers = [p for p in self.doc.paragraphs if p.text.strip() == '{{正文}}']
            if len(markers) != 1 or markers[0]._p.find('.//' + qn('w:sectPr')) is not None:
                raise ValueError('模板须有且仅有一个单独的 {{正文}} 段落，该段落不能带分节符。')
            self._body_marker = markers[0]._p
        styles = self.doc.styles.element
        original_styles = [deepcopy(item) for item in styles] if self._preserve_template_styles else []
        self._setup_styles()
        for original, adjusted in zip(original_styles, list(styles)):
            styles.replace(adjusted, original)
        if self._body_marker is None:
            self._setup_page_margins()
            self._setup_header()
            self._setup_footer()

    def warn(self, code, message):
        self.diagnostics.append({'code': code, 'message': message, 'line': self.source_line})
    
    def _set_font(self, target, zh_name: str, en_name: str = None):
        """
        统一设置字体的辅助方法 - 基于调试经验总结的标准规范
        
        **关键发现**：
        1. Run 层级优先级 > 样式层级 - run 对象的字体设置会覆盖段落样式
        2. 中文字体必须同时设置四个属性才能确保在所有场景下生效：
           - font.name: 字体族名（通用）
           - w:eastAsia: 东亚字符（中日韩）- 决定 Word 界面显示的字体名称
           - w:ascii: ASCII 字符（英文数字）
           - w:hAnsi: 非 ASCII 字符（阿拉伯文、希伯来文等）
        3. 优先使用中文名称以确保 Word 界面显示正确
        
        Args:
            target: 目标对象（样式对象或 run 对象）
            zh_name: 字体中文名称（如 '思源宋体 CN Heavy'）
            en_name: 字体英文名称（可选，如 'Source Han Serif CN Heavy'）
                     如果为 None，则使用 zh_name
        
        Example:
            # 设置样式层级
            heading_style = self.doc.styles['Heading 1']
            self._set_font(heading_style, '思源宋体 CN Heavy', 'Source Han Serif CN Heavy')
            
            # 设置 run 层级（实际生效）
            run = paragraph.add_run('文本')
            self._set_font(run, '思源宋体 CN Heavy', 'Source Han Serif CN Heavy')
        """
        if en_name is None:
            en_name = zh_name
        
        # 设置字体族名
        target.font.name = zh_name
        
        # 设置四个字体属性 - 这是确保中文字体在所有场景下生效的关键
        # 使用 get_or_add_rPr 确保元素存在，避免 AttributeError 崩溃
        rPr = target._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.insert(0, rFonts)
        rFonts.set(qn('w:eastAsia'), zh_name)
        rFonts.set(qn('w:ascii'), en_name)
        rFonts.set(qn('w:hAnsi'), en_name)
    
    def _setup_styles(self):
        """设置文档样式（v1.2 阅读体验优化版）"""
        # ====== 正文字体与段落样式 ======
        normal_style = self.doc.styles['Normal']
        self._set_font(normal_style, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
        normal_style.font.size = Pt(self.config.body_font_size)           # 正文字号（从配置读取）
        normal_style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)  # 黑色
        normal_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE  # 多倍行距
        normal_style.paragraph_format.line_spacing = self.config.body_line_spacing  # 行距倍数（从配置读取）
        normal_style.paragraph_format.first_line_indent = Cm(0.74)  # 首行缩进2字符
        normal_style.paragraph_format.space_before = Pt(0)   # 段前0磅
        normal_style.paragraph_format.space_after = Pt(0)    # 段后0磅
        
        # ====== 标题字体与层级样式 ======
        # 标题层级配置表：【优化】H2-H6 增加完整的字号/颜色/间距差异化
        heading_config = {
            1: {'size': Pt(28), 'color': RGBColor(0xC0, 0x39, 0x2B), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.CENTER, 'space_before': Pt(18), 'space_after': Pt(14),
                 'line_spacing': Pt(34), 'line_rule': WD_LINE_SPACING.EXACTLY},
            2: {'size': Pt(22), 'color': RGBColor(0x00, 0x20, 0x60), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.LEFT, 'space_before': Pt(18), 'space_after': Pt(18),
                 'line_spacing': Pt(28), 'line_rule': WD_LINE_SPACING.EXACTLY},
            3: {'size': Pt(18), 'color': RGBColor(0x00, 0x20, 0x60), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.LEFT, 'space_before': Pt(18), 'space_after': Pt(18),
                 'line_spacing': Pt(28), 'line_rule': WD_LINE_SPACING.EXACTLY},
            4: {'size': Pt(18), 'color': RGBColor(0x00, 0x20, 0x60), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.LEFT, 'space_before': Pt(12), 'space_after': Pt(12),
                 'line_spacing': Pt(22), 'line_rule': WD_LINE_SPACING.EXACTLY},
            5: {'size': Pt(16), 'color': RGBColor(0x00, 0x20, 0x60), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.LEFT, 'space_before': Pt(12), 'space_after': Pt(12),
                 'line_spacing': Pt(22), 'line_rule': WD_LINE_SPACING.EXACTLY},
            6: {'size': Pt(16), 'color': RGBColor(0x00, 0x20, 0x60), 'bold': True,
                 'align': WD_ALIGN_PARAGRAPH.LEFT, 'space_before': Pt(12), 'space_after': Pt(12),
                 'line_spacing': Pt(22), 'line_rule': WD_LINE_SPACING.EXACTLY},
        }
        
        for i in range(1, 7):
            style_name = f'Heading {i}'
            if style_name in self.doc.styles:
                heading_style = self.doc.styles[style_name]
                cfg = heading_config[i]
                
                if i == 1:
                    self._set_font(heading_style, FONT_CONFIG['h1']['zh'], FONT_CONFIG['h1']['en'])
                else:
                    self._set_font(heading_style, FONT_CONFIG['h2_h6']['zh'], FONT_CONFIG['h2_h6']['en'])
                
                heading_style.font.size = cfg['size']
                heading_style.font.color.rgb = cfg['color']
                heading_style.font.bold = cfg['bold']
                heading_style.paragraph_format.alignment = cfg['align']
                heading_style.paragraph_format.space_before = cfg['space_before']
                heading_style.paragraph_format.space_after = cfg['space_after']
                heading_style.paragraph_format.line_spacing_rule = cfg['line_rule']
                heading_style.paragraph_format.line_spacing = cfg['line_spacing']
                heading_style.paragraph_format.first_line_indent = Cm(0)  # 所有标题首行缩进为0
        
        # ====== 列表样式（优化：字号+行距+项间距）======
        if 'List Paragraph' in self.doc.styles:
            list_style = self.doc.styles['List Paragraph']
            self._set_font(list_style, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
            list_style.font.size = Pt(12)              # 【优化】列表字号与正文一致
            list_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE  # 【优化】列表行距
            list_style.paragraph_format.space_after = Pt(4)  # 【优化】列表项之间增加小间距
        
        # 设置表格样式（宋体）
        if 'Table Grid' in self.doc.styles:
            table_style = self.doc.styles['Table Grid']
            self._set_font(table_style, 'SimSun')
            table_style.font.size = Pt(10.5)
        
        # 如果没有 Table Grid 样式，创建它
        if 'Table Grid' not in self.doc.styles:
            table_style = self.doc.styles.add_style('Table Grid', WD_STYLE_TYPE.TABLE)
            self._set_font(table_style, 'SimSun')
        
        # ====== 代码块样式（优化：背景色+边框效果）======
        if 'Code Block' not in self.doc.styles:
            code_style = self.doc.styles.add_style('Code Block', WD_STYLE_TYPE.PARAGRAPH)
            code_font = code_style.font
            code_font.name = 'Consolas'
            code_font.size = Pt(9.5)               # 【优化】稍大一点更易读
            code_font.bold = False
            code_font.color.rgb = RGBColor(0x33, 0x33, 0x33)  # 【优化】深灰色代码文字
            code_style.paragraph_format.space_before = Pt(8)
            code_style.paragraph_format.space_after = Pt(8)
            code_style.paragraph_format.left_indent = Cm(0.6)
            code_style.paragraph_format.right_indent = Cm(0.3)
        
        # ====== 引用块样式（优化：左边框+背景色+字体颜色）======
        if 'Quote' not in self.doc.styles:
            quote_style = self.doc.styles.add_style('Quote', WD_STYLE_TYPE.PARAGRAPH)
        else:
            quote_style = self.doc.styles['Quote']
        quote_font = quote_style.font
        quote_font.name = 'SimSun'
        quote_font.size = Pt(11)               # 【优化】引用文字比正文略小，但比之前10pt大
        quote_font.color.rgb = RGBColor(0x55, 0x55, 0x55)  # 【优化】深灰色引用文字
        quote_font.italic = False  # 明确关闭斜体！无论样式是否已存在
        # 引用块内的 **加粗**、*斜体* 等由 _add_formatted_text 按Markdown语法解析
        quote_style.paragraph_format.left_indent = Cm(1)
        quote_style.paragraph_format.right_indent = Cm(0.5)   # 【优化】右侧缩进
        quote_style.paragraph_format.space_before = Pt(8)
        quote_style.paragraph_format.space_after = Pt(8)
        # 【优化】添加左边框装饰（深蓝色 3pt）
        try:
            pbdr_quote = OxmlElement('w:pBdr')
            left_bdr = OxmlElement('w:left')
            left_bdr.set(qn('w:val'), 'single')
            left_bdr.set(qn('w:sz'), '18')          # 3pt (18 × 1/8 pt)
            left_bdr.set(qn('w:space'), '6')         # 边框与文字间距 6pt
            left_bdr.set(qn('w:color'), '2980B9')     # 深蓝色
            pbdr_quote.append(left_bdr)
            # 将边框附加到样式的 pPr 中（通过底层 XML 操作）
            quote_style._element.get_or_add_pPr().append(pbdr_quote)
        except Exception:
            pass  # 边框添加失败不影响主要功能

        # 设置列表编号样式
        self._setup_numbering_styles()

    def _setup_numbering_styles(self):
        """设置列表编号样式（创建 AbstractNum 和 Num 定义）"""
        try:
            numbering_part = self.doc.part.numbering_part
            if numbering_part is None:
                return
            
            # 直接创建，不检查是否存在（重复创建会被忽略）
            abstract_num_bullet = self._create_abstract_num('bullet', False)
            self._number_abstract_ids = {
                delimiter: self._create_abstract_num('numbered', True, delimiter)
                for delimiter in ('.', ')')
            }
            
            existing = [int(n.get(qn('w:numId'))) for n in numbering_part._element.findall(qn('w:num'))]
            self._bullet_num_id = max(existing, default=0) + 1
            self._num_instance_counter = self._bullet_num_id + 1
            self._create_num_instance(abstract_num_bullet, self._bullet_num_id)
            
        except Exception as e:
            self.warn('numbering_setup_failed', f'创建 Word 编号样式失败：{e}')
    
    def _create_abstract_num(self, name: str, is_ordered: bool, delimiter: str = '.') -> int:
        """
        创建 AbstractNum 定义（支持九级嵌套列表）
        
        Args:
            name: 样式名称
            is_ordered: True 为有序列表，False 为无序列表
            
        Returns:
            abstractNumId
        """
        numbering_part = self.doc.part.numbering_part
        existing = [int(n.get(qn('w:abstractNumId'))) for n in numbering_part._element.findall(qn('w:abstractNum'))]
        abstract_num_id = max(existing, default=-1) + 1
        
        abstract_num = OxmlElement('w:abstractNum')
        abstract_num.set(qn('w:abstractNumId'), str(abstract_num_id))
        
        name_elem = OxmlElement('w:name')
        name_elem.set(qn('w:val'), name)
        abstract_num.append(name_elem)
        
        # 创建5个层级，支持多级嵌套列表
        for ilvl in range(9):
            lvl = OxmlElement('w:lvl')
            lvl.set(qn('w:ilvl'), str(ilvl))
            
            start = OxmlElement('w:start')
            start.set(qn('w:val'), '1')
            lvl.append(start)
            
            if is_ordered:
                # 有序列表：使用数字编号
                lvl_fmt = OxmlElement('w:numFmt')
                lvl_fmt.set(qn('w:val'), 'decimal')
                lvl.append(lvl_fmt)
                
                lvl_text = OxmlElement('w:lvlText')
                lvl_text.set(qn('w:val'), f'%{ilvl+1}{delimiter}')
                lvl.append(lvl_text)
            else:
                # 无序列表：使用项目符号
                lvl_fmt = OxmlElement('w:numFmt')
                lvl_fmt.set(qn('w:val'), 'bullet')
                lvl.append(lvl_fmt)
                
                lvl_text = OxmlElement('w:lvlText')
                lvl_text.set(qn('w:val'), '•')
                lvl.append(lvl_text)
                
                rPr = OxmlElement('w:rPr')
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:ascii'), 'Consolas')
                rFonts.set(qn('w:hAnsi'), 'Consolas')
                rFonts.set(qn('w:hint'), 'default')
                rPr.append(rFonts)
                lvl.append(rPr)
            
            pPr = OxmlElement('w:pPr')
            ind = OxmlElement('w:ind')
            ind.set(qn('w:left'), str(720 * (ilvl + 1)))  # 每级缩进增加720twips(0.5英寸)
            ind.set(qn('w:hanging'), '360')
            pPr.append(ind)
            lvl.append(pPr)
            
            abstract_num.append(lvl)
        
        numbering_part._element.append(abstract_num)
        
        return abstract_num_id
    
    def _create_num_instance(self, abstract_num_id: int, num_id: int, level=0, start=1):
        """
        创建 Num 实例（引用 AbstractNum）
        
        Args:
            abstract_num_id: AbstractNum ID
            num_id: 要分配的 Num ID
        """
        numbering_part = self.doc.part.numbering_part
        
        num = OxmlElement('w:num')
        num.set(qn('w:numId'), str(num_id))
        
        # abstractNumId 必须放在 lvlOverride 外面
        abstract_num_ref = OxmlElement('w:abstractNumId')
        abstract_num_ref.set(qn('w:val'), str(abstract_num_id))
        num.append(abstract_num_ref)
        
        # lvlOverride：添加 startOverride 确保每个列表组从1开始编号
        num_lvl_override = OxmlElement('w:lvlOverride')
        num_lvl_override.set(qn('w:ilvl'), str(level))
        start_override = OxmlElement('w:startOverride')
        start_override.set(qn('w:val'), str(start))
        num_lvl_override.append(start_override)
        num.append(num_lvl_override)
        
        numbering_part._element.append(num)

    def _setup_page_margins(self):
        """设置页面边距"""
        for section in (list(self.doc.sections)[1:] if self._cover_created else self.doc.sections):
            # 从配置读取边距值（单位：厘米）
            section.top_margin = Cm(self.config.margin_top)
            section.bottom_margin = Cm(self.config.margin_bottom)
            section.left_margin = Cm(self.config.margin_left)
            section.right_margin = Cm(self.config.margin_right)
    
    def _setup_header(self):
        """设置页眉 - 统一应用于所有页面"""
        # 遍历所有节（确保每页都有相同的页眉）
        for section in (list(self.doc.sections)[1:] if self._cover_created else self.doc.sections):
            # 获取页眉
            header = section.header
            
            # 完全清空页眉中的所有段落
            header._element.clear()
            
            # 创建页眉段落（只保留这一个段落）
            header_para = header.add_paragraph()
            header_para.style = None  # 移除默认样式，避免产生额外空行
            
            # 添加页眉文本
            header_text = self.config.header_text
            run = header_para.add_run(header_text)
            
            # 设置字体 - 使用 _set_font 方法确保四个属性都正确设置
            # 中文字体：方正清刻本悦宋简体，英文备用：Times New Roman
            self._set_font(run, self.config.header_font_zh, self.config.header_font_en)
            
            # 设置字号：10 磅
            run.font.size = Pt(self.config.header_font_size)
            
            # 设置颜色：【优化】灰色（#888），不抢眼
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
            
            # 设置段落格式 - 左对齐
            header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT  # 左对齐
            
            # 设置左右缩进为 0
            header_para.paragraph_format.left_indent = Cm(0)
            header_para.paragraph_format.right_indent = Cm(0)
            # 【细节修复】取消页眉首行缩进
            header_para.paragraph_format.first_line_indent = Cm(0)
            
            # 设置段前段后间距为 0，行距为单倍
            header_para.paragraph_format.space_before = Pt(0)
            header_para.paragraph_format.space_after = Pt(0)
            header_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            
            # 添加底部边框（仅底边线）
            pPr = header_para._element.get_or_add_pPr()
            
            # 创建边框元素
            pbdr = OxmlElement('w:pBdr')
            
            # 设置底边边框：【优化】单实线，浅灰色，0.25pt 细线
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')      # 单实线
            bottom.set(qn('w:sz'), '2')             # 0.25 磅 (单位是 1/8 磅)
            bottom.set(qn('w:space'), '2')          # 边框与文本间距
            bottom.set(qn('w:color'), 'BBBBBB')     # 浅灰色
            
            pbdr.append(bottom)
            pPr.append(pbdr)
            
            # 设置页眉顶端距离（1.43 厘米）
            section.header_distance = Cm(1.43)
    
    def _setup_footer(self):
        """设置页脚 - 统一应用于所有页面，包含自动页码"""
        # 遍历所有节（确保每页都有相同的页脚）
        for section in (list(self.doc.sections)[1:] if self._cover_created else self.doc.sections):
            # 获取页脚
            footer = section.footer
            
            # 完全清空页脚中的所有段落
            footer._element.clear()
            
            # 创建页脚段落（只保留这一个段落）
            footer_para = footer.add_paragraph()
            footer_para.style = None  # 移除默认样式，避免产生额外空行
            
            # 添加页码文本 - 使用域代码实现自动计数
            # 格式：第 X 页 / 共 Y 页
            run = footer_para.add_run('第 ')
            # 设置字体
            self._set_font(run, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)  # 【优化】灰色页脚文字
            
            # 添加当前页码域代码 (PAGE)
            page_field = OxmlElement('w:fldSimple')
            page_field.set(qn('w:instr'), ' PAGE ')  # 当前页码
            page_run = OxmlElement('w:r')
            # 为域代码内的 run 设置字体
            page_rpr = OxmlElement('w:rPr')
            page_fonts = OxmlElement('w:rFonts')
            page_fonts.set(qn('w:eastAsia'), FONT_CONFIG['body']['zh'])
            page_fonts.set(qn('w:ascii'), FONT_CONFIG['body']['en'])
            page_fonts.set(qn('w:hAnsi'), FONT_CONFIG['body']['en'])
            page_rpr.append(page_fonts)
            page_size = OxmlElement('w:sz')
            page_size.set(qn('w:val'), '20')  # 10pt = 20 半磅单位
            page_rpr.append(page_size)
            page_color = OxmlElement('w:color')
            page_color.set(qn('w:val'), '888888')   # 【优化】灰色
            page_rpr.append(page_color)
            page_run.append(page_rpr)
            
            page_text = OxmlElement('w:t')
            page_text.text = '1'  # 占位符，实际由 Word 动态计算
            page_run.append(page_text)
            page_field.append(page_run)
            footer_para._element.append(page_field)
            
            run = footer_para.add_run(' 页 / 共 ')
            # 设置字体
            self._set_font(run, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)  # 【优化】灰色页脚文字
            
            # 添加总页数域代码 (NUMPAGES)
            numpages_field = OxmlElement('w:fldSimple')
            numpages_field.set(qn('w:instr'), ' NUMPAGES ')  # 总页数
            numpages_run = OxmlElement('w:r')
            # 为域代码内的 run 设置字体
            numpages_rpr = OxmlElement('w:rPr')
            numpages_fonts = OxmlElement('w:rFonts')
            numpages_fonts.set(qn('w:eastAsia'), FONT_CONFIG['body']['zh'])
            numpages_fonts.set(qn('w:ascii'), FONT_CONFIG['body']['en'])
            numpages_fonts.set(qn('w:hAnsi'), FONT_CONFIG['body']['en'])
            numpages_rpr.append(numpages_fonts)
            numpages_size = OxmlElement('w:sz')
            numpages_size.set(qn('w:val'), '20')  # 10pt = 20 半磅单位
            numpages_rpr.append(numpages_size)
            numpages_color = OxmlElement('w:color')
            numpages_color.set(qn('w:val'), '888888')  # 【优化】灰色
            numpages_rpr.append(numpages_color)
            numpages_run.append(numpages_rpr)
            
            numpages_text = OxmlElement('w:t')
            numpages_text.text = '1'  # 占位符，实际由 Word 动态计算
            numpages_run.append(numpages_text)
            numpages_field.append(numpages_run)
            footer_para._element.append(numpages_field)
            
            run = footer_para.add_run(' 页')
            
            # 设置字体 - 使用 _set_font 方法确保四个属性都正确设置
            # 中文字体：宋体，英文备用：Times New Roman
            self._set_font(run, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
            
            # 设置字号：10 磅
            run.font.size = Pt(10)
            
            # 设置颜色：黑色
            run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)  # 【优化】灰色页脚文字
            
            # 设置段落格式 - 右对齐
            footer_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # 设置左右缩进为 0
            footer_para.paragraph_format.left_indent = Cm(0)
            footer_para.paragraph_format.right_indent = Cm(0)
            
            # 设置段前段后间距为 0，行距为单倍
            footer_para.paragraph_format.space_before = Pt(0)
            footer_para.paragraph_format.space_after = Pt(0)
            footer_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            
            # 添加顶部边框（仅顶边线）
            pPr = footer_para._element.get_or_add_pPr()
            
            # 创建边框元素
            pbdr = OxmlElement('w:pBdr')
            
            # 设置顶部边框：【优化】单实线，浅灰色，0.25pt 细线
            top = OxmlElement('w:top')
            top.set(qn('w:val'), 'single')       # 单实线
            top.set(qn('w:sz'), '2')              # 0.25 磅 (单位是 1/8 磅)
            top.set(qn('w:space'), '2')           # 边框与文本间距
            top.set(qn('w:color'), 'BBBBBB')      # 浅灰色
            
            pbdr.append(top)
            pPr.append(pbdr)
            
            # 设置页脚底端距离（1.40 厘米）
            section.footer_distance = Cm(1.40)
    
    def generate(self, elements: List[DocumentElement], output_path: str) -> str:
        """
        生成 Word 文档
        
        Args:
            elements: 文档元素列表
            output_path: 输出文件路径
            
        Returns:
            生成的文件路径
        """
        body = self.doc._element.body
        original_blocks = set(body)
        if self.config.generate_cover and self._body_marker is None:
            self._add_cover_page(elements)
        
        # 生成目录（如果配置开启）
        if self.config.generate_toc:
            self._add_toc()
        
        for element in elements:
            self.source_line = element.line or None
            self._add_element(element)
        self.source_line = None

        if self._body_marker is not None:
            for block in list(body):
                if block not in original_blocks and block.tag != qn('w:sectPr'):
                    self._body_marker.addprevious(block)
            body.remove(self._body_marker)
        
        self.doc.save(output_path)
        return output_path
    
    def _add_cover_page(self, elements: List[DocumentElement] = None):
        """添加封面页 - 独立节，无页眉页脚
        
        Args:
            elements: 文档元素列表，用于自动提取封面标题（第一个H1）
        """
        # 获取第一个节（封面节）
        cover_section = self.doc.sections[0]
        
        # 封面节不显示页眉页脚
        cover_section.header.is_linked_to_previous = False
        cover_section.footer.is_linked_to_previous = False
        
        # 清空封面节的页眉页脚
        cover_section.header._element.clear()
        cover_section.footer._element.clear()
        
        # 添加空行（将标题往下移到页面中间位置）
        for _ in range(10):
            p = self.doc.add_paragraph()
            p.paragraph_format.line_spacing = Pt(22)
        
        # 提取封面标题：优先使用配置，其次从MD的H1提取，最后用默认值
        cover_title_text = self.config.cover_title
        if not cover_title_text and elements:
            for elem in elements:
                if elem.type == ElementType.HEADING and elem.level == 1:
                    cover_title_text = str(elem.content)
                    break
        if not cover_title_text:
            cover_title_text = '文档'
        
        # 添加封面标题
        cover_title = self.doc.add_paragraph()
        cover_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = cover_title.add_run(cover_title_text)
        self._set_font(run, FONT_CONFIG['h1']['zh'], FONT_CONFIG['h1']['en'])
        run.font.size = Pt(28)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
        cover_title.paragraph_format.space_after = Pt(20)
        cover_title.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        # 添加副标题（靠右对齐，正文样式，12磅字）
        cover_subtitle_text = self.config.cover_subtitle or ''
        cover_subtitle = self.doc.add_paragraph()
        cover_subtitle.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = cover_subtitle.add_run(cover_subtitle_text)
        self._set_font(run, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        cover_subtitle.paragraph_format.space_after = Pt(20)
        
        # 添加空行（将日期往下移到页面底部）
        for _ in range(8):
            p = self.doc.add_paragraph()
            p.paragraph_format.line_spacing = Pt(24)
        
        # 添加日期（右下角）
        from datetime import datetime
        cover_date_text = self.config.cover_date or datetime.now().strftime('%Y年%m月')
        cover_date = self.doc.add_paragraph()
        cover_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = cover_date.add_run(f'日期：{cover_date_text}')
        self._set_font(run, FONT_CONFIG['body']['zh'], FONT_CONFIG['body']['en'])
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x7F, 0x8C, 0x8D)
        
        # 添加分页符，开始新的节（正文节）
        self.doc.add_section()
        
        # 获取正文节
        content_section = self.doc.sections[-1]
        content_section.header.is_linked_to_previous = False
        content_section.footer.is_linked_to_previous = False
        self._cover_created = True
        
        # 设置正文节边距（与封面节一致）
        content_section.top_margin = Cm(self.config.margin_top)
        content_section.bottom_margin = Cm(self.config.margin_bottom)
        content_section.left_margin = Cm(self.config.margin_left)
        content_section.right_margin = Cm(self.config.margin_right)
        
        # 正文节添加页眉页脚
        self._setup_header()
        self._setup_footer()
    
    def _add_toc(self):
        """添加目录（使用 Word 原生 TOC 域）"""
        # 目录标题
        toc_heading = self.doc.add_paragraph()
        toc_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        toc_heading.paragraph_format.first_line_indent = Cm(0)
        run = toc_heading.add_run(self.config.toc_title)
        self._set_font(run, FONT_CONFIG['h1']['zh'], FONT_CONFIG['h1']['en'])
        run.font.size = Pt(22)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)
        toc_heading.paragraph_format.space_after = Pt(12)
        
        # 插入 TOC 域
        para = self.doc.add_paragraph()
        para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        
        fldChar_begin = OxmlElement('w:fldChar')
        fldChar_begin.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = f' TOC \\o "1-{self.config.toc_max_level}" \\h \\z \\u '
        
        fldChar_separate = OxmlElement('w:fldChar')
        fldChar_separate.set(qn('w:fldCharType'), 'separate')
        
        fldChar_end = OxmlElement('w:fldChar')
        fldChar_end.set(qn('w:fldCharType'), 'end')
        
        run = para.add_run()
        r_element = run._r
        r_element.append(fldChar_begin)
        r_element.append(instrText)
        r_element.append(fldChar_separate)
        r_element.append(fldChar_end)
        
        # 目录后分页
        page_break_para = self.doc.add_paragraph()
        page_break_run = page_break_para.add_run()
        page_break_run.add_break(WD_BREAK.PAGE)
    
    def _add_element(self, element: DocumentElement):
        """添加文档元素"""
        if element.type == ElementType.HEADING:
            self._add_heading(element)
        elif element.type == ElementType.PARAGRAPH:
            self._add_paragraph(element)
        elif element.type == ElementType.LIST_ITEM:
            self._add_list_item(element)
        elif element.type == ElementType.CODE_BLOCK:
            self._add_code_block(element)
        elif element.type == ElementType.TABLE:
            self._add_table(element)
        elif element.type == ElementType.IMAGE:
            self._add_image(element)
        elif element.type == ElementType.HORIZONTAL_RULE:
            self._add_horizontal_rule()
        elif element.type == ElementType.BLOCKQUOTE:
            self._add_blockquote(element)
    
    def _add_heading(self, element: DocumentElement):
        """添加标题"""
        level = element.level
        content = str(element.content).strip()
        
        style_name = f'Heading {level}'
        
        # H1 标题前分页（跳过第一个H1，从第二个H1开始）
        if level == 1 and self.config.h1_page_break and self._h1_count > 0:
            page_break_para = self.doc.add_paragraph()
            page_break_run = page_break_para.add_run()
            page_break_run.add_break(WD_BREAK.PAGE)
        
        if level == 1:
            self._h1_count += 1
        
        paragraph = self.doc.add_heading(level=level)
        paragraph.style = style_name
        if level == 2 and content.startswith('——') and not self._preserve_template_styles:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        
        # 添加格式化文本（包含粗体等）
        self._add_formatted_text(paragraph, content)
    
    def _add_paragraph(self, element: DocumentElement):
        """添加段落（v1.2 优化：12pt字号+1.5倍行距+首行缩进）"""
        paragraph = self.doc.add_paragraph()
        
        content = str(element.content)
        
        if not self._preserve_template_styles:
            paragraph.paragraph_format.line_spacing = self.config.body_line_spacing
            paragraph.paragraph_format.first_line_indent = Cm(0.74)

        self._add_formatted_text(paragraph, content)
    
    def _add_code_block(self, element: DocumentElement):
        """添加代码块"""
        code_block: CodeBlock = element.content
        
        if code_block.language:
            lang_para = self.doc.add_paragraph()
            lang_run = lang_para.add_run(f"[{code_block.language}]")
            lang_run.font.size = Pt(8)
        
        code_para = self.doc.add_paragraph()
        code_para.style = 'Code Block'
        code_run = code_para.add_run(code_block.code)
        code_run.font.name = 'Consolas'
        
        # 添加背景色
        self._set_paragraph_shading(code_para, 'F5F5F5')

    def _add_list_item(self, element: DocumentElement):
        """生成可编辑编号，并保留每一级列表的起始数字。"""
        data = element.extra_data
        level = element.level - 1
        ordered = data['ordered']
        delimiter = data['delimiter']
        group_id = data['group_id']
        if group_id != self._list_group:
            self._list_group = group_id
            self._group_num_id_map.clear()
            self._list_counters.clear()
        for depth in tuple(self._group_num_id_map):
            if depth > level:
                del self._group_num_id_map[depth]
                self._list_counters.pop(depth, None)
        series = self._group_num_id_map.get(level)
        if series is None or series[0] != delimiter:
            start = data['number'] if ordered else 1
            num_id = None if ordered else getattr(self, '_bullet_num_id', None)
            if ordered and hasattr(self, '_number_abstract_ids'):
                self._num_instance_counter += 1
                num_id = self._num_instance_counter
                self._create_num_instance(self._number_abstract_ids[delimiter], num_id, level, start)
            series = (delimiter, num_id)
            self._group_num_id_map[level] = series
            self._list_counters[level] = start
        _, num_id = series
        number = self._list_counters[level]
        self._list_counters[level] += 1

        paragraph = self.doc.add_paragraph(style='List Paragraph')
        paragraph.paragraph_format.left_indent = Cm(0.8 * (level + 1))
        paragraph.paragraph_format.first_line_indent = Cm(-0.55)
        if not self._preserve_template_styles:
            paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)

        content = re.sub(r'^\[([ xX])\](?=\s|$)',
                         lambda m: '☑' if m.group(1).lower() == 'x' else '☐',
                         str(element.content).strip())
        if num_id is None:
            prefix = f'{number}{delimiter} ' if ordered else '• '
            content = prefix + content
        else:
            num_pr = paragraph._p.get_or_add_pPr().get_or_add_numPr()
            num_pr.get_or_add_ilvl().val = level
            num_pr.get_or_add_numId().val = num_id
        self._add_formatted_text(paragraph, content)

    def _add_table(self, element: DocumentElement):
        """添加表格（v1.3 优化：12磅字号+自动行高+居中对齐+自动调整）"""
        table_data: TableData = element.content
        
        if not table_data.headers:
            return
        
        # 创建表格（使用 Table Grid 样式保留边框）
        table = self.doc.add_table(rows=len(table_data.rows) + 1, cols=len(table_data.headers))
        table.style = 'Table Grid'
        
        # 设置表格在页面中居中对齐
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # 【优化】设置表头（浅蓝背景色 + 加粗 + 居中）
        header_row = table.rows[0]
        repeat = OxmlElement('w:tblHeader')
        header_row._tr.get_or_add_trPr().append(repeat)
        # 表头底色：浅蓝色 #E8F4FD
        header_bg_color = 'E8F4FD'
        for i, header in enumerate(table_data.headers):
            cell = header_row.cells[i]
            # 使用格式化文本写入，支持 **加粗** 等语法
            self._add_formatted_text(cell.paragraphs[0], header)
            # 表头整体加粗 + 11磅字号
            for run in cell.paragraphs[0].runs:
                run.font.bold = True
                run.font.size = Pt(11)   # 11磅字号
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # 段前段后0，首行缩进0，固定行距12磅
            cell.paragraphs[0].paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            cell.paragraphs[0].paragraph_format.line_spacing = Pt(12)
            cell.paragraphs[0].paragraph_format.space_before = Pt(0)
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            cell.paragraphs[0].paragraph_format.first_line_indent = Cm(0)
            # 垂直居中
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            # 设置单元格内边距（减小到1磅）
            self._set_cell_padding(cell, top=20, bottom=20, left=40, right=40)
            # 设置表头背景色
            self._set_cell_shading(cell, header_bg_color)
        
        # 【优化】设置数据行（斑马纹 + 内边距 + 内联格式）
        zebra_colors = ['FFFFFF', 'F5F7FA']  # 白色 / 浅灰交替
        for row_idx, row_data in enumerate(table_data.rows, start=1):
            row = table.rows[row_idx]
            bg_color = zebra_colors[row_idx % 2]  # 斑马纹交替
            for col_idx, cell_data in enumerate(row_data):
                cell = row.cells[col_idx]
                self._add_formatted_text(cell.paragraphs[0], cell_data)
                cell.paragraphs[0].alignment = {'left': WD_ALIGN_PARAGRAPH.LEFT, 'center': WD_ALIGN_PARAGRAPH.CENTER, 'right': WD_ALIGN_PARAGRAPH.RIGHT}.get(table_data.alignments[col_idx] if col_idx < len(table_data.alignments) else 'left', WD_ALIGN_PARAGRAPH.LEFT)
                # 11磅字号
                for run in cell.paragraphs[0].runs:
                    run.font.size = Pt(11)
                # 段前段后0，首行缩进0，固定行距12磅
                cell.paragraphs[0].paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                cell.paragraphs[0].paragraph_format.line_spacing = Pt(12)
                cell.paragraphs[0].paragraph_format.space_before = Pt(0)
                cell.paragraphs[0].paragraph_format.space_after = Pt(0)
                cell.paragraphs[0].paragraph_format.first_line_indent = Cm(0)
                # 垂直居中
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                # 设置单元格内边距（减小到1磅）
                self._set_cell_padding(cell, top=20, bottom=20, left=40, right=40)
                # 设置斑马纹背景色
                self._set_cell_shading(cell, bg_color)
        
        # 设置表格自动适配：根据内容自动调整列宽和行高
        table.autofit = True
        
        # 通过 XML 强制设置表格布局为 "autofit"（根据内容自动调整）
        tbl = table._tbl
        tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
        
        # 设置表格布局为 autofit（根据内容自动调整）
        tblLayout = OxmlElement('w:tblLayout')
        tblLayout.set(qn('w:type'), 'autofit')
        
        # 确保没有固定表格宽度（移除 w:tblWidth）
        existing_tblWidth = tblPr.find(qn('w:tblWidth'))
        if existing_tblWidth is not None:
            tblPr.remove(existing_tblWidth)
        
        # 添加或替换 tblLayout
        existing_tblLayout = tblPr.find(qn('w:tblLayout'))
        if existing_tblLayout is not None:
            tblPr.remove(existing_tblLayout)
        tblPr.append(tblLayout)
        
        # 关键：移除 Table Grid 样式设置的固定列宽（这是autofit不生效的根本原因）
        # 遍历所有列定义，删除固定宽度属性
        for tblGrid in tbl.findall(qn('w:tblGrid')):
            for col in tblGrid.findall(qn('w:gridCol')):
                if qn('w:w') in col.attrib:
                    del col.attrib[qn('w:w')]
        
        # 同时移除所有单元格的固定宽度设置
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.find(qn('w:tcPr'))
                if tcPr is not None:
                    tcWidth = tcPr.find(qn('w:tcW'))
                    if tcWidth is not None:
                        tcPr.remove(tcWidth)
    
    def _set_cell_padding(self, cell, top: int = 50, bottom: int = 50, left: int = 80, right: int = 80):
        """【新增】设置单元格内边距（单位：twips，1pt = 20 twips）"""
        try:
            tc = cell._element.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for margin_name, value in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
                margin_elem = OxmlElement(f'w:{margin_name}')
                margin_elem.set(qn('w:w'), str(value))   # 宽度 (twips)
                margin_elem.set(qn('w:type'), 'dxa')      # 单位类型
                tcMar.append(margin_elem)
            tc.append(tcMar)
        except Exception:
            pass
    
    def _set_cell_shading(self, cell, color_hex: str):
        """设置单元格背景色"""
        try:
            tc = cell._element.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'autoFill')
            shd.set(qn('w:fill'), color_hex)
            tc.append(shd)
        except Exception:
            pass
    
    def _set_paragraph_shading(self, paragraph, color_hex: str):
        """设置段落背景色"""
        try:
            pPr = paragraph._element.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), color_hex)
            pPr.append(shd)
        except Exception:
            pass
    
    def _add_image(self, element: DocumentElement):
        image = element.content
        path = self._get_image_path(image.url)
        if path:
            try:
                self._insert_picture(self.doc.add_paragraph().add_run(), path, image.alt_text)
                self.doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if image.alt_text:
                    caption = self.doc.add_paragraph(image.alt_text)
                    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    for run in caption.runs:
                        run.font.size = Pt(10)
                return
            except Exception as exc:
                self.warn('image_insert_failed', f'图片无法插入：{image.url}（{exc}）')
        self.doc.add_paragraph(f'[图片：{image.alt_text}]({image.url})')

    def _insert_picture(self, run, path, alt):
        shape = run.add_picture(path)
        section = self._image_section()
        usable_width = (section.page_width - section.left_margin - section.right_margin) / 914400
        width_limit = Inches(max(0.1, min(self.config.max_image_width, usable_width)))
        height_limit = max(Inches(0.1), section.page_height - section.top_margin
                           - section.bottom_margin - Inches(0.5))
        scale = min(1, width_limit / shape.width, height_limit / shape.height)
        if scale < 1:
            shape.width = int(shape.width * scale)
            shape.height = int(shape.height * scale)
        if alt:
            shape._inline.docPr.set('descr', alt)

    def _image_section(self):
        if self._body_marker is not None:
            for block in self._body_marker.itersiblings():
                properties = block if block.tag == qn('w:sectPr') else block.find(
                    './' + qn('w:pPr') + '/' + qn('w:sectPr'))
                if properties is not None:
                    return Section(properties, self.doc.part)
        return self.doc.sections[-1]
    
    def _get_image_path(self, url: str) -> Optional[str]:
        """相对输入文档查图，仅跟踪本次下载的临时文件。"""
        from pathlib import Path
        from urllib.parse import unquote
        url = url.strip().strip('<>')
        if url.startswith(('http://', 'https://')):
            if getattr(self, 'offline', False):
                self.warn('image_offline', f'离线模式未下载网络图片：{url}')
                return None
            try:
                import requests
                with requests.get(url, timeout=(5, 15), stream=True) as response:
                    response.raise_for_status()
                    fd, path = tempfile.mkstemp(suffix='.img')
                    self.downloaded_images.add(path)
                    size = 0
                    with os.fdopen(fd, 'wb') as stream:
                        for chunk in response.iter_content(chunk_size=65536):
                            size += len(chunk)
                            if size > 20 * 1024 * 1024:
                                raise ValueError('网络图片超过 20 MiB')
                            stream.write(chunk)
                return path
            except Exception as exc:
                self.warn('image_download_failed', f'网络图片不可用：{url}（{exc}）')
                return None
        path = Path(unquote(url))
        if not path.is_absolute():
            path = Path(self.base_dir) / path
        if path.is_file():
            return str(path.resolve())
        self.warn('image_missing', f'图片不存在：{url}')
        return None
    
    def _add_horizontal_rule(self):
        """添加水平线（v1.2 优化：使用段落边框底线实现原生细线）"""
        paragraph = self.doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(10)
        paragraph.paragraph_format.space_after = Pt(10)
        
        # 【优化】不再使用字符模拟，改用原生段落边框底线
        pPr = paragraph._element.get_or_add_pPr()
        pbdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')     # 单实线
        bottom.set(qn('w:sz'), '4')             # 0.5pt
        bottom.set(qn('w:space'), '4')          # 与上下内容的间距
        bottom.set(qn('w:color'), 'CCCCCC')     # 浅灰色
        pbdr.append(bottom)
        pPr.append(pbdr)
    
    def _add_blockquote(self, element: DocumentElement):
        """添加引用块"""
        paragraph = self.doc.add_paragraph()
        paragraph.style = 'Quote'
        
        content = str(element.content)
        self._add_formatted_text(paragraph, content)
    
    def _add_formatted_text(self, paragraph, text: str):
        """
        应用行内格式化（粗体、斜体、代码等）
        
        Args:
            paragraph: 段落对象
            text: 要格式化的文本
        """
        
        # 分割文本并识别每个部分的格式
        parts = self._split_formatted_text(text)
        hyperlinks = {}
        style = paragraph.style
        style_name = style.name if style else ''
        is_heading = style_name.startswith('Heading ')
        font_key = ('h1' if style_name == 'Heading 1' else 'h2_h6') if is_heading else 'body'
        
        # 为每个部分创建独立的 run
        for part in parts:
            # 如果是图片
            if part.get('is_image'):
                try:
                    image_path = self._get_image_path(part['image_url'])
                    if image_path and os.path.exists(image_path):
                        # 在段落中添加图片
                        run = paragraph.add_run()
                        self._insert_picture(run, image_path, part.get('alt', ''))
                    else:
                        paragraph.add_run(f"[图片：{part.get('alt', '')}]({part['image_url']})")
                except Exception as e:
                    self.warn('image_insert_failed', f"图片无法插入：{part['image_url']}（{e}）")
                    # 图片加载失败，显示文字
                    paragraph.add_run(f"[图片：{part.get('alt', '')}]({part['image_url']})")
                continue
            
            if not part['text']:
                continue
                
            run = paragraph.add_run(part['text'])
            
            if not self._preserve_template_styles:
                font = FONT_CONFIG[font_key]
                self._set_font(run, font['zh'], font['en'])
            
            # 应用格式（注意：必须显式设置 False，否则会继承段落样式的斜体/粗体等属性）
            run.font.bold = True if part.get('bold') else (None if is_heading else False)
            run.font.italic = bool(part.get('italic', False))
            if part.get('code'):
                run.font.name = 'Consolas'
                run.font.size = Pt(9)
            if part.get('strike'):
                run.font.strike = True
            if part.get('superscript'):
                run.font.superscript = True
            if part.get('color'):
                # 移除 # 前缀
                color_code = part['color'].lstrip('#')
                run.font.color.rgb = RGBColor(int(color_code[:2], 16), int(color_code[2:4], 16), int(color_code[4:], 16))
            if part.get('hyperlink'):
                target, identifier = part['hyperlink']
                if identifier not in hyperlinks:
                    link = OxmlElement('w:hyperlink')
                    link.set(qn('r:id'), paragraph.part.relate_to(target, RT.HYPERLINK, is_external=True))
                    run._r.addprevious(link)
                    hyperlinks[identifier] = link
                run.font.underline = True
                if not part.get('color'):
                    run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
                hyperlinks[identifier].append(run._r)
    
    def _split_formatted_text(self, text: str) -> List[dict]:
        """按文本中最早出现的标记解析，防止后方链接吞掉前方格式。"""
        patterns = [
            (ESCAPED_PUNCTUATION, 'escape'),
            (r'(?<!`)(`+)(?!`)(.+?)(?<!`)\1(?!`)', 'code'),
            (INLINE_IMAGE, 'image'),
            (r'<sup>(.+?)</sup>', 'superscript'),
            (r'<font\s+color=[\"\'](#[0-9a-fA-F]{6}|[a-zA-Z]+)[\"\']>(.+?)</font>', 'color'),
            (INLINE_LINK, 'link'),
            (r'\*\*(.+?)\*\*', 'bold'),
            (r'__(.+?)__', 'bold'),
            (r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)', 'italic'),
            (r'(?<!\w)_([^_]+)_(?!\w)', 'italic'),
            (r'~~(.+?)~~', 'strike'),
        ]
        result = []
        position = 0
        while position < len(text):
            candidates = []
            for priority, (pattern, kind) in enumerate(patterns):
                match = re.search(pattern, text[position:])
                if match:
                    candidates.append((match.start(), priority, match, kind))
            if not candidates:
                result.append({'text': text[position:]})
                break
            start, _, match, kind = min(candidates, key=lambda value: value[:2])
            if start:
                result.append({'text': text[position:position + start]})
            if kind == 'escape':
                result.append({'text': match.group(1)})
            elif kind == 'image':
                result.append({'text': '', 'is_image': True,
                               'alt': ESCAPED_PUNCTUATION.sub(r'\1', match.group(1)),
                               'image_url': ESCAPED_PUNCTUATION.sub(r'\1', match.group(2))})
            elif kind == 'code':
                literal = match.group(2).replace('\n', ' ')
                if literal.startswith(' ') and literal.endswith(' ') and literal.strip():
                    literal = literal[1:-1]
                result.append({'text': literal, 'code': True})
            else:
                inner = match.group(2) if kind == 'color' else match.group(1)
                nested = self._split_formatted_text(inner)
                if kind == 'link':
                    target = ESCAPED_PUNCTUATION.sub(r'\1', match.group(2))
                    target = target.removeprefix('<').removesuffix('>')
                    try:
                        url = urlsplit(target)
                        valid = ((url.scheme in ('http', 'https') and url.hostname)
                                 or (url.scheme == 'mailto' and url.path))
                    except ValueError:
                        valid = False
                    if valid:
                        link = (target, object())
                        for part in nested:
                            part['hyperlink'] = link
                    else:
                        self.warn('unsupported_link', f'链接目标已按文本保留：{target}')
                        nested.append({'text': f' ({target})'})
                for part in nested:
                    if kind == 'color':
                        colors = {'red': 'FF0000', 'blue': '0000FF', 'green': '008000',
                                  'black': '000000', 'white': 'FFFFFF', 'gray': '808080'}
                        color = colors.get(match.group(1).lower(), match.group(1).lstrip('#'))
                        if re.fullmatch(r'[0-9a-fA-F]{6}', color):
                            part['color'] = color
                    elif kind != 'link':
                        part[kind] = True
                    result.append(part)
            position += match.end()
        return result
