#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown 批量导出 Word 脚本（ordinote-export 功能 · Word 分册）

功能：
    把 markdown 笔记批量导出为 .docx，尽量忠实还原 markdown 原始效果：
    标题层级、加粗/斜体/删除线/下划线、行内代码与代码块、引用、
    有序/无序列表（含嵌套）、表格、分割线、图片（相对路径 / 远程 / data URI）、
    链接，以及 <mark> / <span style="background-color:..."> 的文字高亮与背景色。

    公式（$...$ / $$...$$）在 markdown 解析之前抽取，避免 LaTeX 被
    markdown 转义规则改写；随后经 latex2mathml 转成 MathML，再借用本机
    Office 附带的 MML2OMML.XSL 转成 Word 原生公式对象（OMML）。
    无法转换时退回 Cambria Math 字体按原 LaTeX 文本呈现，并在汇总中计数。
    mermaid / tikz 等图表代码块无法在 Word 中渲染，保留为代码文本。

    源笔记只读，脚本不修改、不移动任何笔记文件。

依赖：
    python-docx、markdown-it-py（缺失时提示 pip install）。
    公式转换需要 latex2mathml（缺失时首次运行自动安装）与 Office 的
    MML2OMML.XSL（自动查找，可用 --omml-xsl 指定）。
    不需要 MPE 扩展与网络（远程图片除外）。

用法：
    python md2docx.py <文件或目录路径> [更多路径...] [选项]

选项：
    --out DIR         输出目录（默认与源文件同目录）
    --skip-existing   已有同名 .docx 时跳过
    --exclude REGEX   按路径正则排除文件
    --omml-xsl PATH   MML2OMML.XSL 路径（默认自动查找本机 Office）

退出码：
    0  全部成功（含跳过）
    1  用法错误 / 缺少依赖
    2  存在导出失败的文件
"""

import argparse
import base64
import html as html_lib
import importlib
import io
import os
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

from lxml import etree

try:
    from markdown_it import MarkdownIt
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "[错误] 缺少 markdown-it-py，可执行：pip install markdown-it-py\n"
    )
    sys.exit(1)

try:
    import docx  # noqa: F401  (仅用于依赖检查)
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import nsdecls, qn
    from docx.oxml import parse_xml
    from docx.shared import Cm, Emu, Pt, RGBColor
    from docx.opc.constants import RELATIONSHIP_TYPE as RT
    from docx.image.image import Image as DocxImage
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "[错误] 缺少 python-docx，可执行：pip install python-docx\n"
    )
    sys.exit(1)


# ------------------------------------------------------------
# 样式常量（对齐 GitHub / MPE 浅色预览观感）
# ------------------------------------------------------------

FONT_ASCII = "Calibri"
FONT_EAST_ASIA = "微软雅黑"
FONT_MONO = "Consolas"
FONT_MATH = "Cambria Math"

COLOR_TEXT = "1F2328"
COLOR_QUOTE_TEXT = "57606A"
COLOR_QUOTE_BAR = "D0D7DE"
COLOR_LINK = "0969DA"
COLOR_INLINE_CODE_BG = "EFF1F3"
COLOR_CODE_BLOCK_BG = "F6F8FA"
COLOR_CODE_BLOCK_BORDER = "E5E9ED"
COLOR_TABLE_HEADER_BG = "F0F3F6"
COLOR_H_RULE = "D1D9E0"
COLOR_MARK_DEFAULT = "FFFF00"

# 排版尺寸（打印向：小字号、紧行距，减少页数）
BODY_FONT_PT = 8.5
CODE_FONT_PT = 8
BODY_LINE_SPACING = 1.1
BODY_SPACE_AFTER_PT = 3
PAGE_MARGIN_X_CM = 1.8
PAGE_MARGIN_Y_CM = 1.0
HEADING_SIZES = (13, 12, 11, 10.5, 10, 9.5)

# 图片显示尺寸按 markdown 指定宽度的 85% 插入
IMAGE_SCALE = 0.85

# 公式转换
MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MML2OMML_PATTERNS = (
    "Microsoft Office/root/Office*/MML2OMML.XSL",
    "Microsoft Office/Office*/MML2OMML.XSL",
)
MATH_XSL_OVERRIDE = None  # 命令行 --omml-xsl 指定的路径

# CSS 命名颜色（笔记中常见的 lightpink / 灰度色等都可能出现在 style 属性里）
CSS_COLORS = {
    "black": "000000", "silver": "C0C0C0", "gray": "808080", "grey": "808080",
    "white": "FFFFFF", "maroon": "800000", "red": "FF0000", "purple": "800080",
    "fuchsia": "FF00FF", "green": "008000", "lime": "00FF00", "olive": "808000",
    "yellow": "FFFF00", "navy": "000080", "blue": "0000FF", "teal": "008080",
    "aqua": "00FFFF", "orange": "FFA500", "aliceblue": "F0F8FF",
    "antiquewhite": "FAEBD7", "aquamarine": "7FFFD4", "azure": "F0FFFF",
    "beige": "F5F5DC", "bisque": "FFE4C4", "blanchedalmond": "FFEBCD",
    "blueviolet": "8A2BE2", "brown": "A52A2A", "burlywood": "DEB887",
    "cadetblue": "5F9EA0", "chartreuse": "7FFF00", "chocolate": "D2691E",
    "coral": "FF7F50", "cornflowerblue": "6495ED", "cornsilk": "FFF8DC",
    "crimson": "DC143C", "darkblue": "00008B", "darkcyan": "008B8B",
    "darkgoldenrod": "B8860B", "darkgray": "A9A9A9", "darkgrey": "A9A9A9",
    "darkgreen": "006400", "darkkhaki": "BDB76B", "darkmagenta": "8B008B",
    "darkolivegreen": "556B2F", "darkorange": "FF8C00", "darkorchid": "9932CC",
    "darkred": "8B0000", "darksalmon": "E9967A", "darkseagreen": "8FBC8F",
    "darkslateblue": "483D8B", "darkslategray": "2F4F4F", "darkslategrey": "2F4F4F",
    "darkturquoise": "00CED1", "darkviolet": "9400D3", "deeppink": "FF1493",
    "deepskyblue": "00BFFF", "dimgray": "696969", "dimgrey": "696969",
    "dodgerblue": "1E90FF", "firebrick": "B22222", "floralwhite": "FFFAF0",
    "forestgreen": "228B22", "gainsboro": "DCDCDC", "ghostwhite": "F8F8FF",
    "gold": "FFD700", "goldenrod": "DAA520", "greenyellow": "ADFF2F",
    "honeydew": "F0FFF0", "hotpink": "FF69B4", "indianred": "CD5C5C",
    "indigo": "4B0082", "ivory": "FFFFF0", "khaki": "F0E68C",
    "lavender": "E6E6FA", "lavenderblush": "FFF0F5", "lawngreen": "7CFC00",
    "lemonchiffon": "FFFACD", "lightblue": "ADD8E6", "lightcoral": "F08080",
    "lightcyan": "E0FFFF", "lightgoldenrodyellow": "FAFAD2",
    "lightgray": "D3D3D3", "lightgrey": "D3D3D3", "lightgreen": "90EE90",
    "lightpink": "FFB6C1", "lightsalmon": "FFA07A", "lightseagreen": "20B2AA",
    "lightskyblue": "87CEFA", "lightslategray": "778899", "lightslategrey": "778899",
    "lightsteelblue": "B0C4DE", "lightyellow": "FFFFE0", "limegreen": "32CD32",
    "linen": "FAF0E6", "mediumaquamarine": "66CDAA", "mediumblue": "0000CD",
    "mediumorchid": "BA55D3", "mediumpurple": "9370DB", "mediumseagreen": "3CB371",
    "mediumslateblue": "7B68EE", "mediumspringgreen": "00FA9A",
    "mediumturquoise": "48D1CC", "mediumvioletred": "C71585",
    "midnightblue": "191970", "mintcream": "F5FFFA", "mistyrose": "FFE4E1",
    "moccasin": "FFE4B5", "navajowhite": "FFDEAD", "oldlace": "FDF5E6",
    "olivedrab": "6B8E23", "orangered": "FF4500", "orchid": "DA70D6",
    "palegoldenrod": "EEE8AA", "palegreen": "98FB98", "paleturquoise": "AFEEEE",
    "palevioletred": "DB7093", "papayawhip": "FFEFD5", "peachpuff": "FFDAB9",
    "peru": "CD853F", "pink": "FFC0CB", "plum": "DDA0DD", "powderblue": "B0E0E6",
    "rosybrown": "BC8F8F", "royalblue": "4169E1", "saddlebrown": "8B4513",
    "salmon": "FA8072", "sandybrown": "F4A460", "seagreen": "2E8B57",
    "seashell": "FFF5EE", "sienna": "A0522D", "skyblue": "87CEEB",
    "slateblue": "6A5ACD", "slategray": "708090", "slategrey": "708090",
    "snow": "FFFAFA", "springgreen": "00FF7F", "steelblue": "4682B4",
    "tan": "D2B48C", "thistle": "D8BFD8", "tomato": "FF6347", "turquoise": "40E0D0",
    "violet": "EE82EE", "wheat": "F5DEB3", "whitesmoke": "F5F5F5",
    "yellowgreen": "9ACD32", "transparent": None,
}


def parse_color(value):
    """把 CSS 颜色值转换为 6 位 hex；无法识别或 transparent 返回 None。"""
    if not value:
        return None
    v = value.strip().lower()
    if v in ("inherit", "currentcolor"):
        return None
    if v in CSS_COLORS:
        return CSS_COLORS[v]
    if v.startswith("#"):
        h = v[1:]
        if len(h) == 3:
            return "".join(c * 2 for c in h).upper()
        if len(h) == 6:
            return h.upper()
        if len(h) == 8:  # #RRGGBBAA：忽略 alpha
            return h[:6].upper()
        return None
    m = re.match(r"rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)", v)
    if m:
        return "{:02X}{:02X}{:02X}".format(
            *(max(0, min(255, int(round(float(x))))) for x in m.groups())
        )
    return None


def style_attrs(style_text):
    """从 style 属性中提取 background-color / color。"""
    fill = color = None
    m = re.search(r"background-color\s*:\s*([^;]+)", style_text, flags=re.I)
    if m:
        fill = parse_color(m.group(1))
    m = re.search(r"(?<![-])color\s*:\s*([^;]+)", style_text, flags=re.I)
    if m:
        color = parse_color(m.group(1))
    return fill, color


# CSS 长度单位 → EMU 换算（96dpi：1px = 1/96 in；1pt = 1/72 in）
_EMU_PER_INCH = 914400
_LENGTH_UNITS = {
    "px": _EMU_PER_INCH / 96,
    "pt": _EMU_PER_INCH / 72,
    "cm": _EMU_PER_INCH / 2.54,
    "mm": _EMU_PER_INCH / 25.4,
    "in": _EMU_PER_INCH,
}


def parse_dimension(value, reference_emu, font_pt=10):
    """把 CSS 长度值转换为 EMU；百分比按 reference_emu 折算。无法识别返回 None。"""
    if value is None:
        return None
    v = value.strip().lower()
    if not v or v in ("auto", "none"):
        return None
    if v.endswith("%"):
        try:
            ratio = max(0.0, float(v[:-1].strip()) / 100)
        except ValueError:
            return None
        return int(reference_emu * ratio)
    m = re.fullmatch(r"([\d.]+)\s*([a-z]*)", v)
    if not m:
        return None
    try:
        number = float(m.group(1))
    except ValueError:
        return None
    unit = m.group(2)
    if unit in ("", "px"):
        return int(number * _LENGTH_UNITS["px"])
    if unit in ("em", "rem"):
        return int(number * font_pt * _LENGTH_UNITS["pt"])
    if unit in _LENGTH_UNITS:
        return int(number * _LENGTH_UNITS[unit])
    return None


# ------------------------------------------------------------
# markdown 预处理（公式抽取，与 md2pdf.py 同一策略）
# ------------------------------------------------------------

MATH_SCAN_RE = re.compile(
    r"(?P<code>`[^`\n]*`)"
    r"|(?P<display>(?<!\\)\$\$[\s\S]+?(?<!\\)\$\$)"
    r"|(?P<inline>(?<!\\)\$(?![\s$])(?:[^$\n]|\\\$)+?(?<![\\\s])(?<!\\)\$(?!\d))"
)

FENCE_RE = re.compile(r"^\s*(?:```|~~~)")


def _split_fences(text):
    """按围栏代码块切分文本，返回 [(是否为代码块, 片段), ...]。"""
    segments, buffer, in_code = [], [], False
    for line in text.splitlines(keepends=True):
        if FENCE_RE.match(line):
            segments.append((in_code, "".join(buffer)))
            buffer, in_code = [], not in_code
        buffer.append(line)
    segments.append((in_code, "".join(buffer)))
    return segments


def protect_math(text):
    """把公式抽成占位 span，LaTeX 原文放在 data-tex 与元素文本中。"""
    pieces = []
    for in_code, segment in _split_fences(text):
        if in_code:
            pieces.append(segment)
            continue

        def replace(match):
            if match.lastgroup == "code":
                return match.group(0)
            raw = match.group(0)
            display = raw.startswith("$$")
            tex = " ".join((raw[2:-2] if display else raw[1:-1]).split())
            cls = "math math-display" if display else "math"
            return (
                f'<span class="{cls}" data-tex="{html_lib.escape(tex, quote=True)}">'
                f"{html_lib.escape(tex)}</span>"
            )

        pieces.append(MATH_SCAN_RE.sub(replace, segment))
    return "".join(pieces)


def strip_front_matter(text):
    """剥离文件开头的 YAML front-matter。"""
    if text.startswith("---"):
        match = re.match(r"^---\r?\n.*?\r?\n---\r?\n", text, flags=re.S)
        if match:
            return text[match.end():]
    return text


def md_to_html(md_text):
    md = MarkdownIt("commonmark", {"html": True}).enable(["table", "strikethrough"])
    return md.render(protect_math(md_text))


# ------------------------------------------------------------
# 公式转换：LaTeX → MathML → OMML
# ------------------------------------------------------------

_math_converter = None      # latex2mathml.converter 模块；False 表示不可用
_math_xslt = None           # lxml XSLT 对象；False 表示不可用

BARE_AMP_RE = re.compile(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9A-Fa-f]+);)")
UNSUPPORTED_MATH_RE = re.compile(
    r"\\(?:ce|pu|chemfig|tikz)\b|\\begin\{tikzcd\}"
)
MULTILINE_MATH_RE = re.compile(
    r"^\s*\\begin\{(aligned|align|alignat|gathered|gather|multline|split|eqnarray)\*?\}"
    r"(?:\s*\{[^}]*\})?([\s\S]*)\\end\{\1\*?\}\s*$"
)


def _ensure_math_converter():
    """导入 latex2mathml；首次缺失时自动安装一次。"""
    global _math_converter
    if _math_converter is not None:
        return _math_converter or None

    try:
        _math_converter = importlib.import_module("latex2mathml.converter")
        return _math_converter
    except ImportError:
        pass

    print("[提示] 未检测到 latex2mathml，正在自动安装 ...")
    try:
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "latex2mathml"],
            check=True, timeout=300,
        )
    except Exception as exc:
        print(f"[警告] latex2mathml 自动安装失败：{exc}", file=sys.stderr)
        _math_converter = False
        return None

    importlib.invalidate_caches()
    try:
        _math_converter = importlib.import_module("latex2mathml.converter")
    except ImportError as exc:
        print(f"[警告] latex2mathml 安装后仍无法导入：{exc}", file=sys.stderr)
        _math_converter = False
        return None
    return _math_converter


def _find_mml2omml():
    """定位 Office 自带的 MML2OMML.XSL。"""
    if MATH_XSL_OVERRIDE:
        path = Path(MATH_XSL_OVERRIDE).expanduser()
        return path if path.is_file() else None
    for env in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.environ.get(env)
        if not root:
            continue
        for pattern in MML2OMML_PATTERNS:
            hits = sorted(Path(root).glob(pattern))
            for hit in reversed(hits):
                if hit.is_file():
                    return hit
    return None


def _get_math_xslt():
    global _math_xslt
    if _math_xslt is not None:
        return _math_xslt or None
    path = _find_mml2omml()
    if path is None:
        _math_xslt = False
        return None
    try:
        _math_xslt = etree.XSLT(etree.parse(str(path)))
    except Exception as exc:
        print(f"[警告] 加载 {path} 失败：{exc}", file=sys.stderr)
        _math_xslt = False
        return None
    return _math_xslt


def _tex_to_omml(tex, display):
    """单行 LaTeX 转 OMML 元素；失败返回 None。"""
    xslt = _get_math_xslt()
    if xslt is None:
        return None
    converter = _ensure_math_converter()
    if converter is None:
        return None
    try:
        mathml = converter.convert(tex, display="block" if display else "inline")
        mathml = BARE_AMP_RE.sub("&amp;", mathml)
        result = xslt(etree.fromstring(mathml.encode("utf-8")))
        root = result.getroot()
        if root is not None:
            etree.cleanup_namespaces(root)
    except Exception:
        return None
    return root


def _split_tex_rows(body):
    """按顶层 \\\\ 切分多行公式；花括号与环境内的不切。"""
    rows, buf, brace, env, i = [], [], 0, 0, 0
    while i < len(body):
        if body.startswith("\\begin", i):
            end = body.find("}", i)
            if end != -1:
                env += 1
                buf.append(body[i:end + 1])
                i = end + 1
                continue
        if body.startswith("\\end", i):
            end = body.find("}", i)
            if end != -1:
                env = max(0, env - 1)
                buf.append(body[i:end + 1])
                i = end + 1
                continue
        if body.startswith("\\\\", i):
            if brace == 0 and env == 0:
                rows.append("".join(buf))
                buf = []
                i += 2
                continue
            buf.append("\\\\")
            i += 2
            continue
        char = body[i]
        if char == "\\" and i + 1 < len(body):
            buf.append(body[i:i + 2])
            i += 2
            continue
        if char == "{":
            brace += 1
        elif char == "}":
            brace = max(0, brace - 1)
        buf.append(char)
        i += 1
    rows.append("".join(buf))
    return rows


def _strip_align_markers(text):
    """去掉顶层的对齐符 & 与行距参数 [2pt]，保留嵌套环境内的 &。"""
    out, brace, env, i = [], 0, 0, 0
    while i < len(text):
        if text.startswith("\\begin", i):
            end = text.find("}", i)
            if end != -1:
                env += 1
                out.append(text[i:end + 1])
                i = end + 1
                continue
        if text.startswith("\\end", i):
            end = text.find("}", i)
            if end != -1:
                env = max(0, env - 1)
                out.append(text[i:end + 1])
                i = end + 1
                continue
        char = text[i]
        if char == "\\" and i + 1 < len(text):
            out.append(text[i:i + 2])
            i += 2
            continue
        if char == "{":
            brace += 1
        elif char == "}":
            brace = max(0, brace - 1)
        if char == "&" and brace == 0 and env == 0:
            i += 1
            continue
        out.append(char)
        i += 1
    return re.sub(r"^\s*\[[^\]]*\]", "", "".join(out)).strip()


def tex_to_omml_rows(tex):
    """把 LaTeX 转成若干 OMML 元素；多行环境返回多行，失败返回 None。"""
    if UNSUPPORTED_MATH_RE.search(tex):
        return None

    match = MULTILINE_MATH_RE.match(tex)
    if not match:
        omml = _tex_to_omml(tex, display=False)
        return [omml] if omml is not None else None

    body = match.group(2)
    rows = [
        row for row in (_strip_align_markers(item) for item in _split_tex_rows(body))
        if row
    ]
    if not rows:
        return None
    elements = []
    for row in rows:
        omml = _tex_to_omml(row, display=True)
        if omml is None:
            return None
        elements.append(omml)
    return elements


def wrap_display_math(omml):
    """把 m:oMath 包进 m:oMathPara 并居中。"""
    para = etree.Element(qn("m:oMathPara"), nsmap={"m": MATH_NS})
    props = etree.SubElement(para, qn("m:oMathParaPr"))
    jc = etree.SubElement(props, qn("m:jc"))
    jc.set(qn("m:val"), "center")
    para.append(omml)
    return para


# ------------------------------------------------------------
# docx 低层工具
# ------------------------------------------------------------

def _set_rfonts(run, ascii_font=FONT_ASCII, ea_font=FONT_EAST_ASIA):
    run.font.name = ascii_font
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)


def shade_run(run, fill_hex):
    rpr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    rpr.append(shd)


def shade_paragraph(paragraph, fill_hex):
    ppr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    ppr.append(shd)


def paragraph_border(paragraph, position, color, sz=18, space=8):
    """给段落加单边边框（left 用于引用，bottom 用于标题/分割线）。"""
    ppr = paragraph._p.get_or_add_pPr()
    pbdr = ppr.find(qn("w:pBdr"))
    if pbdr is None:
        pbdr = OxmlElement("w:pBdr")
        ppr.append(pbdr)
    border = OxmlElement(f"w:{position}")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), str(sz))
    border.set(qn("w:space"), str(space))
    border.set(qn("w:color"), color)
    pbdr.append(border)


def add_hyperlink(paragraph, url, run):
    """把已经设置好格式的 run 包进外部超链接元素。"""
    r_id = paragraph.part.relate_to(url, RT.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)


def setup_numbering(document):
    """创建项目符号与十进制编号的 abstractNum，返回 numbering 元素与两个 abstractNumId。"""
    numbering = document.part.numbering_part.element

    abstract_ids = [int(e.get(qn("w:abstractNumId"))) for e in numbering.findall(qn("w:abstractNum"))]
    bullet_abs = (max(abstract_ids) + 1) if abstract_ids else 0
    decimal_abs = bullet_abs + 1

    def level_xml(num_id, fmt, text, indent):
        if fmt == "bullet":
            extra = '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>'
        else:
            extra = ""
        return (
            f'<w:lvl w:ilvl="{num_id}" {nsdecls("w")}>'
            f'<w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{indent}" w:hanging="260"/></w:pPr>'
            f"{extra}</w:lvl>"
        )

    bullet_levels = "".join(
        level_xml(i, "bullet", "\uF0B7", 420 * (i + 1)) for i in range(9)
    )
    decimal_levels = "".join(
        level_xml(i, "decimal", f"%{i + 1}.", 420 * (i + 1)) for i in range(9)
    )

    numbering.append(parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{bullet_abs}">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>{bullet_levels}</w:abstractNum>'
    ))
    numbering.append(parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{decimal_abs}">'
        f'<w:multiLevelType w:val="hybridMultilevel"/>{decimal_levels}</w:abstractNum>'
    ))

    num_ids = [int(e.get(qn("w:numId"))) for e in numbering.findall(qn("w:num"))]
    next_num_id = (max(num_ids) + 1) if num_ids else 1

    def new_num(abs_id, restart):
        nonlocal next_num_id
        num_id = next_num_id
        next_num_id += 1
        overrides = ""
        if restart:
            overrides = "".join(
                f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="1"/></w:lvlOverride>'
                for i in range(9)
            )
        numbering.append(parse_xml(
            f'<w:num {nsdecls("w")} w:numId="{num_id}">'
            f'<w:abstractNumId w:val="{abs_id}"/>{overrides}</w:num>'
        ))
        return num_id

    bullet_num = new_num(bullet_abs, restart=False)
    return bullet_num, decimal_abs, new_num


def configure_styles(document):
    """统一文档样式，使其接近 GitHub / MPE 浅色预览。"""
    # A4 纸张；页边距左右 1.8cm，上下 1cm
    for section in document.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(PAGE_MARGIN_X_CM)
        section.right_margin = Cm(PAGE_MARGIN_X_CM)
        section.top_margin = Cm(PAGE_MARGIN_Y_CM)
        section.bottom_margin = Cm(PAGE_MARGIN_Y_CM)

    normal = document.styles["Normal"]
    normal.font.name = FONT_ASCII
    normal.font.size = Pt(BODY_FONT_PT)
    normal.font.color.rgb = RGBColor.from_string(COLOR_TEXT)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_EAST_ASIA)
    pf = normal.paragraph_format
    pf.line_spacing = BODY_LINE_SPACING
    pf.space_before = Pt(0)
    pf.space_after = Pt(BODY_SPACE_AFTER_PT)

    for i in range(1, 7):
        style = document.styles[f"Heading {i}"]
        style.font.name = FONT_ASCII
        style.font.size = Pt(HEADING_SIZES[i - 1])
        style.font.bold = True
        style.font.italic = False
        style.font.color.rgb = RGBColor.from_string(COLOR_TEXT)
        style.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_EAST_ASIA)
        style.paragraph_format.space_before = Pt(10 if i <= 2 else 6)
        style.paragraph_format.space_after = Pt(4)
        style.paragraph_format.keep_with_next = True
        if i in (1, 2):
            ppr = style.element.get_or_add_pPr()
            pbdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "12" if i == 1 else "6")
            bottom.set(qn("w:space"), "4")
            bottom.set(qn("w:color"), COLOR_H_RULE)
            pbdr.append(bottom)
            # pBdr 在 pPr 中位于 spacing 之前
            spacing = ppr.find(qn("w:spacing"))
            if spacing is not None:
                spacing.addprevious(pbdr)
            else:
                ppr.append(pbdr)


# ------------------------------------------------------------
# HTML → docx 转换器
# ------------------------------------------------------------

INLINE_FORMAT_KEYS = ("b", "i", "u", "strike", "code", "fill", "color", "link", "math")


def _base_format(**over):
    fmt = dict.fromkeys(INLINE_FORMAT_KEYS, None)
    fmt.update(over)
    return fmt


class MarkdownDocxBuilder(HTMLParser):
    """把 markdown-it 产出的 HTML 写入 python-docx 文档。

    mode='doc'：写入文档正文；mode='cell'：写入指定表格单元格。
    """

    def __init__(self, document, base_dir, warnings, numbering_info,
                 mode="doc", cell=None, header_cell=False):
        super().__init__(convert_charrefs=True)
        self.doc = document
        self.base_dir = base_dir
        self.warnings = warnings
        self.bullet_num, self.decimal_abs, self.new_num = numbering_info
        self.mode = mode
        self.cell = cell

        self.para = None
        self.list_stack = []      # 每项为 ("ul", num_id) 或 ("ol", num_id)
        self.quote_depth = 0
        self.style_stack = [_base_format(b=True)] if header_cell else [_base_format()]
        self.in_pre = False
        self.pre_lines = []
        self.math_display = False
        self.in_math_span = False
        self._cell_para_reused = False
        self.usable_width = self._usable_width()

    def _usable_width(self):
        section = self.doc.sections[0]
        return section.page_width - section.left_margin - section.right_margin

    # ---- 段落管理 ----

    def finalize_paragraph(self):
        self.para = None

    def begin_code_paragraph(self):
        para = self._new_raw_paragraph()
        # pPr 子元素严格按 schema 顺序：pBdr → shd → spacing → ind
        borders = "".join(
            f'<w:{pos} w:val="single" w:sz="4" w:space="4" w:color="{COLOR_CODE_BLOCK_BORDER}"/>'
            for pos in ("top", "left", "bottom", "right")
        )
        ppr_xml = (
            f'<w:pPr {nsdecls("w")}>'
            f'<w:pBdr>{borders}</w:pBdr>'
            f'<w:shd w:val="clear" w:color="auto" w:fill="{COLOR_CODE_BLOCK_BG}"/>'
            f'<w:spacing w:before="40" w:after="80" w:line="240" w:lineRule="auto"/>'
            f'<w:ind w:left="120"/>'
            f'</w:pPr>'
        )
        old = para._p.find(qn("w:pPr"))
        if old is not None:
            para._p.remove(old)
        para._p.insert(0, parse_xml(ppr_xml))
        self.para = para

    def _new_cell_paragraph(self):
        """复用单元格自带的第一个空段落，避免每个单元格多出一行空段。"""
        if not self._cell_para_reused:
            self._cell_para_reused = True
            return self.cell.paragraphs[0]
        return self.cell.add_paragraph()

    def _new_raw_paragraph(self):
        if self.mode == "cell":
            return self._new_cell_paragraph()
        return self.doc.add_paragraph()

    def ensure_paragraph(self):
        if self.para is not None:
            return self.para

        if self.mode == "cell":
            para = self._new_cell_paragraph()
            para.paragraph_format.space_after = Pt(BODY_SPACE_AFTER_PT)
            self.para = para
            return para

        para = self.doc.add_paragraph()
        if self.list_stack:
            level = len(self.list_stack) - 1
            entry = self.list_stack[-1]
            if not entry[2]:
                # 列表项的第一个段落：挂编号 / 项目符号
                entry[2] = True
                ppr = para._p.get_or_add_pPr()
                num_pr = OxmlElement("w:numPr")
                ilvl = OxmlElement("w:ilvl")
                ilvl.set(qn("w:val"), str(level))
                numid = OxmlElement("w:numId")
                numid.set(qn("w:val"), str(entry[1]))
                num_pr.append(ilvl)
                num_pr.append(numid)
                ppr.insert(0, num_pr)
            else:
                # 同一项内的后续段落：缩进而不重复编号
                para.paragraph_format.left_indent = Pt(21 * (level + 1))
            para.paragraph_format.space_after = Pt(2)
        if self.quote_depth:
            self._apply_quote_format(para)
        self.para = para
        return para

    def _apply_quote_format(self, para):
        # pPr 中 numPr（若有）必须位于 pBdr 之前，这里保留已插入的 numPr
        old = para._p.find(qn("w:pPr"))
        num_pr = old.find(qn("w:numPr")) if old is not None else None
        ppr_xml = (
            f'<w:pPr {nsdecls("w")}>'
            f'<w:pBdr><w:left w:val="single" w:sz="18" w:space="10" '
            f'w:color="{COLOR_QUOTE_BAR}"/></w:pBdr>'
            f'<w:spacing w:after="80"/>'
            f'<w:ind w:left="360"/>'
            f'</w:pPr>'
        )
        new_ppr = parse_xml(ppr_xml)
        if num_pr is not None:
            new_ppr.insert(0, num_pr)
        if old is not None:
            para._p.remove(old)
        para._p.insert(0, new_ppr)

    # ---- 行内写入 ----

    @property
    def fmt(self):
        return self.style_stack[-1]

    def push_format(self, **over):
        fmt = dict(self.fmt)
        fmt.update({k: v for k, v in over.items() if v is not None})
        self.style_stack.append(fmt)

    def pop_format(self):
        self.style_stack.pop()

    def emit_text(self, text):
        if not text:
            return
        if self.in_math_span:
            # 公式已转成 OMML，span 里的 LaTeX 原文不再写入
            return
        if self.in_pre:
            self.pre_lines.append(text)
            return
        text = re.sub(r"\s+", " ", text)
        if self.para is None and not text.strip():
            return
        para = self.ensure_paragraph()
        run = para.add_run(text)
        self._format_run(run, para, text)

    def _format_run(self, run, para, text):
        fmt = self.fmt
        if fmt["math"]:
            _set_rfonts(run, FONT_MATH, FONT_MATH)
            return
        if fmt["code"]:
            _set_rfonts(run, FONT_MONO, FONT_EAST_ASIA)
            run.font.size = Pt(CODE_FONT_PT)
        else:
            _set_rfonts(run)
        run.font.bold = bool(fmt["b"])
        run.font.italic = bool(fmt["i"])
        run.font.underline = bool(fmt["u"])
        run.font.strike = bool(fmt["strike"])
        color = fmt["color"]
        if color is None and fmt["link"]:
            color = COLOR_LINK
        if color is None and self.quote_depth:
            color = COLOR_QUOTE_TEXT
        if color:
            run.font.color.rgb = RGBColor.from_string(color)
        # shd 在 rPr 中位于 color / sz 之后，最后添加
        if fmt["code"]:
            shade_run(run, COLOR_INLINE_CODE_BG)
        if fmt["fill"]:
            shade_run(run, fmt["fill"])
        if fmt["link"]:
            run.font.underline = True
            add_hyperlink(para, fmt["link"], run)

    def emit_image(self, src, alt, width_attr=None, height_attr=None, style_text=None):
        if not src:
            return
        para = self.ensure_paragraph()
        run = para.add_run()

        # 可用版心：列表内的图片还要扣掉当前层级的缩进
        list_indent = int(Pt(21 * len(self.list_stack))) if self.list_stack else 0
        available = int(self.usable_width) - list_indent

        # style 中的 width/height 优先于标签属性（CSS 表现优先级）
        style_w = style_h = None
        if style_text:
            m = re.search(r"width\s*:\s*([^;]+)", style_text, flags=re.I)
            if m:
                style_w = m.group(1).strip()
            m = re.search(r"height\s*:\s*([^;]+)", style_text, flags=re.I)
            if m:
                style_h = m.group(1).strip()
        want_w = parse_dimension(style_w if style_w is not None else width_attr,
                                 available)
        want_h = parse_dimension(style_h if style_h is not None else height_attr,
                                 available)

        try:
            blob = self._load_image(src)
            image = DocxImage.from_blob(blob)
            native_width, native_height = image.width, image.height
        except Exception as exc:
            placeholder = run
            placeholder.text = f"[图片加载失败：{src}]"
            placeholder.font.color.rgb = RGBColor.from_string("CC0000")
            self.warnings.append(f"图片 {src} 加载失败：{exc}")
            return
        try:
            kwargs = {}
            if want_w is not None or want_h is not None:
                # 指定尺寸超出可用宽度时等比缩小；只给一边时由原生比例推导另一边
                scale = 1.0
                if want_w is not None and want_w > available:
                    scale = available / want_w
                width = (want_w if want_w is not None
                         else native_width * want_h / native_height) * scale
                kwargs["width"] = Emu(int(width * IMAGE_SCALE))
                if want_h is not None:
                    kwargs["height"] = Emu(int(want_h * scale * IMAGE_SCALE))
            else:
                kwargs["width"] = Emu(int(min(native_width, available) * IMAGE_SCALE))
            run.add_picture(io.BytesIO(blob), **kwargs)
        except Exception as exc:
            run.text = f"[图片无法识别：{src}]"
            run.font.color.rgb = RGBColor.from_string("CC0000")
            self.warnings.append(f"图片 {src} 无法识别：{exc}")

    def _load_image(self, src):
        """返回图片二进制；data URI / 远程 URL / 相对与绝对本地路径。"""
        if src.startswith("data:"):
            m = re.match(r"data:image/[^;]+;base64,(.*)", src, flags=re.S)
            if not m:
                raise ValueError("不支持的 data URI 形式")
            return base64.b64decode(m.group(1))
        if re.match(r"https?://", src, flags=re.I):
            req = urllib.request.Request(src, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as resp:
                return resp.read()
        # markdown-it 会把中文、空格等百分号编码；本地路径先反编码
        src = urllib.parse.unquote(src)
        path = Path(src)
        if not path.is_absolute():
            path = self.base_dir / src
        path = path.resolve()
        if not path.is_file():
            raise FileNotFoundError(str(path))
        return path.read_bytes()

    # ---- HTMLParser 事件 ----

    def handle_starttag(self, tag, attrs):
        attr = dict(attrs)
        handler = self._start_handlers.get(tag)
        if handler:
            handler(self, attr)

    def handle_endtag(self, tag):
        handler = self._end_handlers.get(tag)
        if handler:
            handler(self)

    def handle_data(self, data):
        self.emit_text(data)

    # ---- 块级开始 ----

    def _start_heading(self, attr):
        self.finalize_paragraph()
        level = int(self._heading_level)
        if self.mode == "cell":
            self.para = self.cell.add_paragraph()
        else:
            self.para = self.doc.add_paragraph(style=f"Heading {level}")

    def _start_p(self, attr):
        self.finalize_paragraph()

    def _start_blockquote(self, attr):
        self.finalize_paragraph()
        self.quote_depth += 1

    def _start_ul(self, attr):
        self.finalize_paragraph()
        self.list_stack.append(["ul", self.bullet_num, False])

    def _start_ol(self, attr):
        self.finalize_paragraph()
        num_id = self.new_num(self.decimal_abs, restart=True)
        self.list_stack.append(["ol", num_id, False])

    def _start_li(self, attr):
        self.finalize_paragraph()
        if self.list_stack:
            self.list_stack[-1][2] = False

    def _start_pre(self, attr):
        self.finalize_paragraph()
        self.in_pre = True
        self.pre_lines = []
        self.begin_code_paragraph()

    def _start_hr(self, attr):
        self.finalize_paragraph()
        if self.mode != "cell":
            para = self.doc.add_paragraph()
            paragraph_border(para, "bottom", COLOR_H_RULE, sz=6, space=1)
            para.paragraph_format.space_before = Pt(2)
            para.paragraph_format.space_after = Pt(8)

    def _start_table(self, attr):
        # 表格在主流程中预先拆出，cell 模式不应再遇到
        self.finalize_paragraph()

    # ---- 行内开始 ----

    def _start_strong(self, attr):
        self.push_format(b=True)

    def _start_em(self, attr):
        self.push_format(i=True)

    def _start_u(self, attr):
        self.push_format(u=True)

    def _start_strike(self, attr):
        self.push_format(strike=True)

    def _start_inline_code(self, attr):
        if not self.in_pre:
            self.push_format(code=True)

    def _start_a(self, attr):
        href = attr.get("href", "")
        if re.match(r"https?://", href, flags=re.I) or href.startswith("#"):
            self.push_format(link=href if href.startswith("http") else None)
        else:
            self.push_format()

    def _start_mark(self, attr):
        fill, color = style_attrs(attr.get("style", ""))
        self.push_format(fill=fill or COLOR_MARK_DEFAULT, color=color)

    def _emit_math(self, tex, display):
        """把 LaTeX 写成 Word 公式对象；失败返回 False 由调用方降级。"""
        rows = tex_to_omml_rows(tex)
        if rows is None:
            self.warnings.append(f"公式未转换（保留 LaTeX 原文）：{tex[:40]}")
            return False

        if display:
            self.finalize_paragraph()
            for index, omml in enumerate(rows):
                para = self.ensure_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.space_after = Pt(0)
                para._p.append(wrap_display_math(omml))
                if index < len(rows) - 1:
                    self.finalize_paragraph()
            self.finalize_paragraph()
        else:
            para = self.ensure_paragraph()
            para._p.append(rows[0])
        return True

    def _start_span(self, attr):
        classes = attr.get("class", "").split()
        if "math" in classes:
            tex = html_lib.unescape(attr.get("data-tex") or "")
            display = "math-display" in classes
            if tex and self._emit_math(tex, display):
                self.in_math_span = True
                return
            # 降级：以 Cambria Math 按 LaTeX 原文呈现
            self.math_display = display
            if display:
                self.finalize_paragraph()
                para = self.ensure_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            self.push_format(math=True)
            return
        fill, color = (None, None)
        if attr.get("style"):
            fill, color = style_attrs(attr["style"])
        if fill or color:
            self.push_format(fill=fill, color=color)
        else:
            self.push_format()

    def _start_img(self, attr):
        self.emit_image(
            attr.get("src", ""),
            attr.get("alt", ""),
            width_attr=attr.get("width"),
            height_attr=attr.get("height"),
            style_text=attr.get("style"),
        )

    def _start_br(self, attr):
        if self.in_pre:
            self.pre_lines.append("\n")
        elif self.para is not None:
            self.para.add_run().add_break()

    # ---- 结束标签 ----

    def _end_pre(self):
        # commonmark 输出形如 <pre><code>...\n</code></pre>，去掉首尾换行
        text = "".join(self.pre_lines).strip("\r\n")
        lines = text.split("\n")
        for index, line in enumerate(lines):
            if index:
                self.para.add_run().add_break()
            run = self.para.add_run(line)
            _set_rfonts(run, FONT_MONO, FONT_EAST_ASIA)
            run.font.size = Pt(CODE_FONT_PT)
        self.in_pre = False
        self.finalize_paragraph()

    def _end_block_with_para(self):
        self.finalize_paragraph()

    def _end_blockquote(self):
        self.quote_depth = max(0, self.quote_depth - 1)
        self.finalize_paragraph()

    def _end_list(self):
        self.list_stack.pop()
        self.finalize_paragraph()

    def _end_span(self):
        if self.in_math_span:
            self.in_math_span = False
            return
        if self.fmt.get("math"):
            display = self.math_display
            self.pop_format()
            self.math_display = False
            if display:
                self.finalize_paragraph()
        else:
            self.pop_format()

    def _start_th(self, attr):
        # cell 模式下不会遇到 th；主流程表格单独解析
        self.finalize_paragraph()

    def _start_td(self, attr):
        self.finalize_paragraph()

    _heading_level = 1

    def _end_simple(self):
        pass

    _start_handlers = {}
    _end_handlers = {}


# 分发表（在类定义后填充，避免方法在类体中顺序依赖）
MarkdownDocxBuilder._start_handlers = {
    "h1": MarkdownDocxBuilder._start_heading,
    "h2": MarkdownDocxBuilder._start_heading,
    "h3": MarkdownDocxBuilder._start_heading,
    "h4": MarkdownDocxBuilder._start_heading,
    "h5": MarkdownDocxBuilder._start_heading,
    "h6": MarkdownDocxBuilder._start_heading,
    "p": MarkdownDocxBuilder._start_p,
    "blockquote": MarkdownDocxBuilder._start_blockquote,
    "ul": MarkdownDocxBuilder._start_ul,
    "ol": MarkdownDocxBuilder._start_ol,
    "li": MarkdownDocxBuilder._start_li,
    "pre": MarkdownDocxBuilder._start_pre,
    "hr": MarkdownDocxBuilder._start_hr,
    "table": MarkdownDocxBuilder._start_table,
    "th": MarkdownDocxBuilder._start_th,
    "td": MarkdownDocxBuilder._start_td,
    "strong": MarkdownDocxBuilder._start_strong,
    "b": MarkdownDocxBuilder._start_strong,
    "em": MarkdownDocxBuilder._start_em,
    "i": MarkdownDocxBuilder._start_em,
    "u": MarkdownDocxBuilder._start_u,
    "del": MarkdownDocxBuilder._start_strike,
    "s": MarkdownDocxBuilder._start_strike,
    "strike": MarkdownDocxBuilder._start_strike,
    "code": MarkdownDocxBuilder._start_inline_code,
    "a": MarkdownDocxBuilder._start_a,
    "mark": MarkdownDocxBuilder._start_mark,
    "span": MarkdownDocxBuilder._start_span,
    "img": MarkdownDocxBuilder._start_img,
    "br": MarkdownDocxBuilder._start_br,
}
MarkdownDocxBuilder._end_handlers = {
    "h1": MarkdownDocxBuilder._end_block_with_para,
    "h2": MarkdownDocxBuilder._end_block_with_para,
    "h3": MarkdownDocxBuilder._end_block_with_para,
    "h4": MarkdownDocxBuilder._end_block_with_para,
    "h5": MarkdownDocxBuilder._end_block_with_para,
    "h6": MarkdownDocxBuilder._end_block_with_para,
    "p": MarkdownDocxBuilder._end_block_with_para,
    "blockquote": MarkdownDocxBuilder._end_blockquote,
    "ul": MarkdownDocxBuilder._end_list,
    "ol": MarkdownDocxBuilder._end_list,
    "li": MarkdownDocxBuilder._end_block_with_para,
    "pre": MarkdownDocxBuilder._end_pre,
    "table": MarkdownDocxBuilder._end_block_with_para,
    "th": MarkdownDocxBuilder._end_block_with_para,
    "td": MarkdownDocxBuilder._end_block_with_para,
    "strong": MarkdownDocxBuilder.pop_format,
    "b": MarkdownDocxBuilder.pop_format,
    "em": MarkdownDocxBuilder.pop_format,
    "i": MarkdownDocxBuilder.pop_format,
    "u": MarkdownDocxBuilder.pop_format,
    "del": MarkdownDocxBuilder.pop_format,
    "s": MarkdownDocxBuilder.pop_format,
    "strike": MarkdownDocxBuilder.pop_format,
    "code": lambda self: None,  # pre 内未压栈；行内 code 统一在下面处理
}


def _end_code(self):
    # in_pre 状态由 </pre> 收尾；行内 code 在此弹栈
    if not self.in_pre:
        self.pop_format()


MarkdownDocxBuilder._end_handlers["code"] = _end_code
for _tag in ("a", "mark"):
    MarkdownDocxBuilder._end_handlers[_tag] = MarkdownDocxBuilder.pop_format
MarkdownDocxBuilder._end_handlers["span"] = MarkdownDocxBuilder._end_span

# 标题级别通过闭包传入
for _i in range(1, 7):
    def _make_h_start(level):
        def _handler(self, attr):
            self._heading_level = level
            self._start_heading(attr)
        return _handler
    MarkdownDocxBuilder._start_handlers[f"h{_i}"] = _make_h_start(_i)


# ------------------------------------------------------------
# 表格处理
# ------------------------------------------------------------

TABLE_RE = re.compile(r"<table>.*?</table>", flags=re.S)
ROW_RE = re.compile(r"<tr>(.*?)</tr>", flags=re.S)
CELL_RE = re.compile(r"<(th|td)>(.*?)</\1>", flags=re.S)


def render_tables(document, html, base_dir, warnings, numbering_info):
    """把 HTML 按表格切块，非表格片段交给主构建器，表格单独建表并解析单元格。"""
    builder = MarkdownDocxBuilder(document, base_dir, warnings, numbering_info)
    pos = 0
    for match in TABLE_RE.finditer(html):
        builder.feed(html[pos:match.start()])
        builder.finalize_paragraph()

        rows_html = ROW_RE.findall(match.group(0))
        rows = [CELL_RE.findall(row) for row in rows_html]
        rows = [r for r in rows if r]
        if not rows:
            pos = match.end()
            continue
        col_count = max(len(r) for r in rows)

        table = document.add_table(rows=len(rows), cols=col_count)
        try:
            table.style = "Table Grid"
        except KeyError:
            pass
        table.autofit = True

        for r_idx, row in enumerate(rows):
            for c_idx, (tag, inner_html) in enumerate(row):
                cell = table.cell(r_idx, c_idx)
                if r_idx == 0:
                    shade_paragraph(cell.paragraphs[0], COLOR_TABLE_HEADER_BG)
                cell_builder = MarkdownDocxBuilder(
                    document, base_dir, warnings, numbering_info,
                    mode="cell", cell=cell, header_cell=(r_idx == 0),
                )
                cell_builder.usable_width = builder.usable_width / max(col_count, 1)
                cell_builder.feed(inner_html)
                cell_builder.finalize_paragraph()

        document.add_paragraph().paragraph_format.space_after = Pt(2)
        pos = match.end()
    builder.feed(html[pos:])
    builder.finalize_paragraph()


# ------------------------------------------------------------
# 单文件与批处理
# ------------------------------------------------------------

def read_text(path):
    return path.read_text(encoding="utf-8-sig")


def convert_one(md_path, docx_path):
    """完成单文件转换，返回警告列表。"""
    md_text = strip_front_matter(read_text(md_path))
    html = md_to_html(md_text)

    document = Document()
    configure_styles(document)
    numbering_info = setup_numbering(document)
    warnings = []
    render_tables(document, html, md_path.parent.resolve(), warnings, numbering_info)
    document.save(str(docx_path))
    return warnings


def collect_targets(paths, exclude_re):
    """展开目录为 markdown 文件列表，去重并应用排除规则。"""
    targets, seen = [], set()
    for raw in paths:
        path = Path(raw).expanduser()
        if path.is_dir():
            found = sorted(path.rglob("*.md"))
        elif path.is_file():
            found = [path]
        else:
            raise FileNotFoundError(f"路径不存在：{raw}")
        for item in found:
            resolved = item.resolve()
            if resolved in seen:
                continue
            seen.add(resolved)
            if exclude_re and exclude_re.search(str(item)):
                continue
            targets.append(item)
    return targets


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="把 markdown 批量导出为 Word（.docx），忠实还原标题、高亮、图片等格式",
    )
    parser.add_argument("paths", nargs="+", help="markdown 文件或目录，可多个")
    parser.add_argument("--out", help="输出目录，默认与源文件同目录")
    parser.add_argument("--skip-existing", action="store_true", help="已有同名 .docx 时跳过")
    parser.add_argument("--exclude", help="按路径正则排除文件")
    parser.add_argument("--omml-xsl", help="MML2OMML.XSL 路径（默认自动查找本机 Office）")
    return parser.parse_args(argv)


def main(argv=None):
    opts = parse_args(argv)

    global MATH_XSL_OVERRIDE
    MATH_XSL_OVERRIDE = opts.omml_xsl

    try:
        targets = collect_targets(opts.paths, re.compile(opts.exclude) if opts.exclude else None)
    except (FileNotFoundError, re.error) as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1

    if not targets:
        print("[提示] 没有匹配的 markdown 文件")
        return 0

    counts = {"ok": 0, "skip": 0, "fail": 0}
    failures = []
    for md_path in targets:
        out_dir = Path(opts.out).expanduser() if opts.out else md_path.parent
        docx_path = out_dir / (md_path.stem + ".docx")

        if opts.skip_existing and docx_path.exists():
            status, message = "skip", f"已存在同名 Word：{docx_path}"
        else:
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
                warnings = convert_one(md_path, docx_path)
                size = docx_path.stat().st_size
                message = f"{docx_path}（{size / 1024:.0f} KB）"
                if warnings:
                    images = sum(1 for item in warnings if item.startswith("图片"))
                    formulas = sum(1 for item in warnings if item.startswith("公式"))
                    if images:
                        message += f"；{images} 张图片未嵌入"
                    if formulas:
                        message += f"；{formulas} 处公式未转换"
                status = "ok"
            except Exception as exc:  # 单个文件失败不中断整批
                status, message = "fail", f"{type(exc).__name__}: {exc}"

        counts[status] += 1
        if status == "fail":
            failures.append((md_path, message))
        print(f"[{'成功' if status == 'ok' else '跳过' if status == 'skip' else '失败'}] "
              f"{md_path} -> {message}")

    print(f"\n共 {len(targets)} 个文件：成功 {counts['ok']}，跳过 {counts['skip']}，"
          f"失败 {counts['fail']}")
    for md_path, message in failures:
        print(f"  失败原因：{md_path} -> {message}")
    return 2 if counts["fail"] else 0


if __name__ == "__main__":
    sys.exit(main())
