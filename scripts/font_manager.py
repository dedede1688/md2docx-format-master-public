"""只读取技能包内字体，不安装或注册系统字体。"""
from pathlib import Path
import hashlib
import json
import struct

ROOT = Path(__file__).resolve().parents[1]


def tables(data):
    offset = struct.unpack_from('>I', data, 12)[0] if data[:4] == b'ttcf' else 0
    count = struct.unpack_from('>H', data, offset + 4)[0]
    result = {}
    for index in range(count):
        tag, checksum, start, length = struct.unpack_from('>4sIII', data, offset + 12 + index * 16)
        if start + length > len(data):
            raise ValueError('字体表超出文件范围')
        result[tag] = (start, length)
    return offset, result


def font_info(path):
    data = Path(path).read_bytes()
    offset, entries = tables(data)
    names = {}
    aliases = set()
    if b'name' in entries:
        start, length = entries[b'name']
        _, count, storage = struct.unpack_from('>HHH', data, start)
        for index in range(count):
            platform, encoding, language, name_id, size, relative = struct.unpack_from('>HHHHHH', data, start + 6 + index * 12)
            if name_id not in (0, 1, 2, 4, 6, 13, 14, 16, 17):
                continue
            raw = data[start + storage + relative:start + storage + relative + size]
            try:
                value = raw.decode('utf-16-be' if platform in (0, 3) else 'mac_roman').strip('\x00')
            except UnicodeError:
                continue
            if name_id in (1, 4, 6):
                aliases.add(value)
            if name_id not in names or language == 0x409:
                names[name_id] = value
    fs_type = struct.unpack_from('>H', data, entries[b'OS/2'][0] + 8)[0] if b'OS/2' in entries else 2
    style = names.get(2, '').lower()
    # fsSelection 提供独立于本地化样式名的粗体、斜体标记。
    flags = struct.unpack_from('>H', data, entries[b'OS/2'][0] + 62)[0] if b'OS/2' in entries and entries[b'OS/2'][1] >= 64 else 0
    bold = bool(flags & 32) or 'bold' in style
    italic = bool(flags & 1) or 'italic' in style or 'oblique' in style
    face = 'boldItalic' if bold and italic else 'bold' if bold else 'italic' if italic else 'regular'
    aliases.add(names.get(1, Path(path).stem))
    family = names.get(1, Path(path).stem)
    return {'family': family, 'aliases': sorted(aliases), 'face': face,
            'fs_type': fs_type, 'embedding_allowed': not bool(fs_type & (2 | 512)),
            'copyright': names.get(0, ''), 'license_text': names.get(13, ''),
            'license_url': names.get(14, ''), 'sha256': hashlib.sha256(data).hexdigest()}


def catalog():
    result = []
    for path in sorted((ROOT / 'fonts').rglob('*')):
        if path.suffix.lower() not in ('.ttf', '.otf', '.ttc'):
            continue
        info = font_info(path)
        info['file'] = path.relative_to(ROOT).as_posix()
        result.append(info)
    return result


def standalone_font(path):
    """从 TTC 重建首个字体的独立 sfnt；非集合保持原始数据。"""
    data = Path(path).read_bytes()
    if data[:4] != b'ttcf':
        return data
    offset, entries = tables(data)
    count = len(entries)
    power = 1 << (count.bit_length() - 1)
    output = bytearray(data[offset:offset + 4] + struct.pack('>HHHH', count, power * 16, power.bit_length() - 1, count * 16 - power * 16))
    output.extend(b'\0' * count * 16)
    head_start = None
    for index, (tag, (start, size)) in enumerate(sorted(entries.items())):
        body = bytearray(data[start:start + size])
        if tag == b'head':
            body[8:12] = b'\0' * 4
            head_start = len(output)
        padded = body + b'\0' * ((-len(body)) % 4)
        checksum = sum(struct.unpack('>' + 'I' * (len(padded) // 4), padded)) & 0xffffffff
        struct.pack_into('>4sIII', output, 12 + index * 16, tag, checksum, len(output), size)
        output.extend(padded)
    if head_start is not None:
        checksum = sum(struct.unpack('>' + 'I' * (len(output) // 4), output)) & 0xffffffff
        struct.pack_into('>I', output, head_start + 8, (0xB1B0AFBA - checksum) & 0xffffffff)
    return bytes(output)


if __name__ == '__main__':
    result = catalog()
    target = ROOT / 'licenses/fonts'
    target.mkdir(parents=True, exist_ok=True)
    (target / 'font-notices.json').write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'fonts': len(result), 'embedding_allowed': sum(f['embedding_allowed'] for f in result)}))
