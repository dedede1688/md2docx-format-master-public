"""按 OOXML 写入字体部件、关系及混淆后的字体数据。"""
from pathlib import Path
import os
import tempfile
import uuid
import zipfile
from lxml import etree as ET
from font_manager import ROOT, catalog, standalone_font

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
C = 'http://schemas.openxmlformats.org/package/2006/content-types'


def obfuscate(data, key):
    mask = uuid.UUID(key.strip('{}')).bytes[::-1]
    output = bytearray(data)
    for index in range(min(32, len(output))):
        output[index] ^= mask[index % 16]
    return bytes(output)


def xml_bytes(element):
    return ET.tostring(element, encoding='UTF-8', xml_declaration=True, standalone=True)


def embed_fonts(docx_path):
    with zipfile.ZipFile(docx_path) as archive:
        files = {name: archive.read(name) for name in archive.namelist()}
    fonts = catalog()
    requested = set()
    for name, content in files.items():
        if name == 'word/document.xml' or (name.startswith(('word/header', 'word/footer', 'word/footnotes', 'word/endnotes', 'word/comments')) and name.endswith('.xml')):
            root = ET.fromstring(content)
            for element in root.iter('{' + W + '}rFonts'):
                for kind in ('ascii', 'hAnsi', 'eastAsia', 'cs'):
                    value = element.get('{' + W + '}' + kind)
                    if value:
                        requested.add(value)
            if any(True for element in root.iter('{' + W + '}numPr')):
                requested.add('Consolas')
    styles = ET.fromstring(files['word/styles.xml'])
    # 未显式指定字体的段落、代码语言标注继承 Normal。
    for style in styles.findall('{' + W + '}style'):
        if style.get('{' + W + '}styleId') == 'Normal':
            for element in style.iter('{' + W + '}rFonts'):
                requested.update(value for key, value in element.attrib.items() if key.rsplit('}', 1)[-1] in ('ascii', 'hAnsi', 'eastAsia', 'cs'))
    table = ET.fromstring(files['word/fontTable.xml'])
    rel_path = 'word/_rels/fontTable.xml.rels'
    rels = ET.fromstring(files[rel_path]) if rel_path in files else ET.Element('{' + P + '}Relationships', nsmap={None: P})
    types = ET.fromstring(files['[Content_Types].xml'])
    if not any(item.get('Extension') == 'odttf' for item in types):
        ET.SubElement(types, '{' + C + '}Default', Extension='odttf', ContentType='application/vnd.openxmlformats-officedocument.obfuscatedFont')
    next_id = max([int(item.get('Id')[3:]) for item in rels if (item.get('Id') or '').startswith('rId') and item.get('Id')[3:].isdigit()] or [0]) + 1
    embedded = []
    warnings = []
    stored = {}
    for name in sorted(requested):
        matches = [font for font in fonts if name in font['aliases'] or name == font['family']]
        if not matches:
            warnings.append('包内未找到所用字体：' + name)
            continue
        entry = next((item for item in table if item.get('{' + W + '}name') == name), None)
        if entry is None:
            entry = ET.SubElement(table, '{' + W + '}font')
            entry.set('{' + W + '}name', name)
        for font in matches:
            if not font['embedding_allowed']:
                warnings.append('字体嵌入标记不允许嵌入，已跳过：' + font['family'])
                continue
            if font['file'] not in stored:
                key = '{' + str(uuid.uuid4()).upper() + '}'
                filename = f'portable-font-{len(stored) + 1}.odttf'
                rid = 'rId' + str(next_id)
                next_id += 1
                files['word/fonts/' + filename] = obfuscate(standalone_font(ROOT / font['file']), key)
                ET.SubElement(rels, '{' + P + '}Relationship', Id=rid, Type=R + '/font', Target='fonts/' + filename)
                stored[font['file']] = (rid, key)
                embedded.append({'family': font['family'], 'face': font['face'], 'file': font['file']})
            rid, key = stored[font['file']]
            tag = 'embed' + font['face'][0].upper() + font['face'][1:]
            old = entry.find('{' + W + '}' + tag)
            if old is not None:
                entry.remove(old)
            embed = ET.SubElement(entry, '{' + W + '}' + tag)
            embed.set('{' + R + '}id', rid)
            embed.set('{' + W + '}fontKey', key)
            embed.set('{' + W + '}subsetted', 'false')
    files['word/fontTable.xml'] = xml_bytes(table)
    files[rel_path] = xml_bytes(rels)
    files['[Content_Types].xml'] = xml_bytes(types)
    settings = ET.fromstring(files['word/settings.xml'])
    if settings.find('{' + W + '}embedTrueTypeFonts') is None:
        ET.SubElement(settings, '{' + W + '}embedTrueTypeFonts')
    files['word/settings.xml'] = xml_bytes(settings)
    descriptor, temporary = tempfile.mkstemp(suffix='.docx', dir=Path(docx_path).parent)
    os.close(descriptor)
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, content in files.items():
                archive.writestr(name, content)
        os.replace(temporary, docx_path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return embedded, warnings
