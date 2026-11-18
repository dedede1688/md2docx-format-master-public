"""
配置管理模块
处理技能配置和参数
"""

from typing import Optional, Dict, Any
from dataclasses import dataclass


# 旧版 JSON 中接受过这些字段，但转换流程未使用它们。
_LEGACY_UNUSED_KEYS = frozenset({
    'output_filename', 'keep_temp_files', 'auto_open', 'auto_install_fonts',
    'header_color_hex', 'footer_format', 'footer_font_size', 'footer_color_hex',
    'custom_styles',
})


@dataclass
class ConversionConfig:
    """转换配置类"""
    
    # Word 模板路径
    template_path: Optional[str] = None
    template_mode: str = 'append'
    
    # 图片最大宽度（英寸）
    max_image_width: float = 6.0
    
    # 页眉设置
    header_text: str = '机密文件·版权所有'
    header_font_zh: str = '方正清刻本悦宋简体'
    header_font_en: str = 'Times New Roman'
    header_font_size: float = 10

    # 目录设置
    generate_toc: bool = False
    toc_title: str = '目  录'
    toc_max_level: int = 3  # 目录显示到几级标题

    # 封面设置
    generate_cover: bool = True  # 是否生成封面
    cover_title: Optional[str] = None        # 封面标题（None=自动从MD的H1提取）
    cover_subtitle: Optional[str] = None     # 封面副标题（None=不显示）
    cover_date: Optional[str] = None         # 封面日期（None=当前日期）

    # 分页设置
    h1_page_break: bool = True  # 每个一级标题前分页

    # 正文字号（磅）
    body_font_size: float = 14
    body_line_spacing: float = 1.5  # 行距倍数
    
    # 编码方式
    encoding: str = 'utf-8'

    # 页面边距设置（单位：厘米）        
    margin_top: float = 2
    margin_bottom: float = 2
    margin_left: float = 2
    margin_right: float = 2
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ConversionConfig':
        return cls(**{key: value for key, value in data.items() if key not in _LEGACY_UNUSED_KEYS})
