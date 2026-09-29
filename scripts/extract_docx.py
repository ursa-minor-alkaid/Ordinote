#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Word（.docx）→ Markdown 转换脚本（ordinote-extract 功能 · Word 分册）

功能：
    将 .docx 文件转换为 markdown 文件，尽可能保留 Word 的格式信息。

    格式保留能力
    - 文档主标题：首段为“无缩进 / No Spacing / 标题 / Title”样式时输出为 #，
      其余标题样式顺延一级（Heading 1 → ##；--no-title-shift 可关闭）
    - 列表：项目符号与编号列表（decimal 等编号格式），支持多级嵌套与计数；
      编号定义可来自段落直接格式或段落样式（List Bullet / List Number 及其基样式）
    - 行内：加粗 **…**、斜体 *…*、粗斜 ***…***、删除线 ~~…~~、
      下划线 <u>…</u>、荧光笔高亮 <mark>…</mark>
    - 字体区分：
      * 批注字体（默认 楷体/KaiTi/SimKai）→ 淡灰色 mark 包裹
        （本项目约定：霞鹜文楷为 AI 生成正文，楷体为手写批注）
      * Cambria Math 等公式字体中的斜体/加粗视为数学排版样式，不转为强调
    - 表格 → pipe 表格；合并单元格自动去重；全空表格跳过并提示
    - 超链接 → [文本](网址)（网址含空格时用 <网址> 形式）
    - 软回车（w:br / w:cr）→ <br>
    - 内嵌图片 → 导出到输出目录 assets/ 子文件夹，正文以相对路径引用
    - 引用样式（Quote / Intense Quote / 引用）→ > 引用块（相邻引用段合并）
    - 生成信息行（第二段，本项目约定）→ 整行斜体；--no-meta-italic 可关闭

    不保留 / 不提取
    - markdown 无法表达的格式（字号、颜色、对齐、缩进等）
    - 页眉页脚、脚注、尾注、文本框、批注（审阅批注）
    - 修订删除的内容（w:del / w:moveFrom，视作已删除）

用法：
    python extract_docx.py <文件或目录> [更多...] [选项]

    输出默认与源文件同目录、同名（{原名}.md）。纵向笔记场景常用组合：
    python extract_docx.py "archives\纵向笔记（临时）" --strip-version --out "纵向笔记"

选项：
    --out DIR              输出目录（默认：源文件所在目录）
    --suffix SUFFIX        输出文件名后缀：{名}-{SUFFIX}.md（默认：无）
    --strip-version        去除文件名末尾 _v / _v0.5 之类版本号
    --annotation-fonts F   批注字体，逗号分隔（默认：楷体,KaiTi,SimKai）
    --no-annotation        不做批注字体标记
    --no-title-shift       首段不作为文档主标题，标题保持原层级
    --no-meta-italic       第二段不做整行斜体处理
    --no-images            不导出内嵌图片（以 HTML 注释占位）
    --no-auto-install      缺少 python-docx 时不自动安装
    --install-deps         只安装依赖后退出

依赖：
    python-docx（>= 1.1；缺库时可自动安装；本脚本只读源文件，不修改）

退出码：
    0  全部成功
    1  用法错误 / 文件不存在 / 不支持的类型（含旧版 .doc）
    2  转换失败或结果为空
    3  缺少依赖库 python-docx，且自动安装被拒绝或失败

注意：
    - 文本中的 markdown 特殊字符（* _ ~ 等）不做转义，以保持中文笔记输出干净；
      若正文中确有此类字符，请人工核对渲染效果
"""

import argparse
import glob
import re
import sys
from pathlib import Path

# python-docx 延迟加载：缺库时先自动安装再导入（见 load_docx_api）
try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.text.run import Run
except ImportError:
    Document = qn = Run = None

# 是否允许缺依赖时自动安装（--no-auto-install 可关闭）
AUTO_INSTALL = True

# 常量
GRAY_MARK = '<mark style="background-color:#e6e6e6;color:inherit">{}</mark>'
HL_MARK = "<mark>{}</mark>"
U_TAG = "<u>{}</u>"
DEFAULT_ANNOTATION_FONTS = "楷体,KaiTi,SimKai"

HEADING_STYLE_RE = re.compile(r"^(?:Heading|标题)\s*(\d+)$")
TITLE_STYLES = {"No Spacing", "无缩进", "Title", "标题"}
QUOTE_STYLES = {"Quote", "Intense Quote", "引用", "明显引用"}
LIST_BLOCK_RE = re.compile(r"^\s*(?:[-*+]|\d+\.)\s")
META_LINE_MAX = 120  # 生成信息行的最大字符数，超过视为普通段落

# 表格合并单元格去重用的临时标记属性（仅写入内存中的 XML 树，不落盘）
VISITED_ATTR = "{urn:local:extract-docx}visited"


class ExtractError(Exception):
    """单个文件转换失败时抛出，携带退出码与给用户看的说明。

    多文件批量转换时，单个文件失败只记录错误、不中断其他文件。
    """

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


# ------------------------------------------------------------
# 依赖自动安装（与 extract_pdf_pptx.py 相同机制）
# ------------------------------------------------------------

def _confirm(prompt):
    """交互确认，直接回车默认为“是”；非交互终端（管道/无键盘）自动选“是”。"""
    if not sys.stdin or not sys.stdin.isatty():
        print(prompt + "（非交互终端，自动选择 Y）")
        return True
    try:
        return input(prompt).strip().lower() not in ("n", "no")
    except (EOFError, KeyboardInterrupt):
        print()
        return False


def _pip_install():
    import subprocess
    cmd = [sys.executable, "-m", "pip", "install", "python-docx"]
    print(f"[安装] 执行：{' '.join(cmd)}")
    try:
        result = subprocess.run(cmd)
    except OSError as e:
        raise ExtractError(3, f"调用 pip 失败：{e}；请手动执行：pip install python-docx")
    if result.returncode != 0:
        raise ExtractError(
            3, f"python-docx 安装失败（pip 退出码 {result.returncode}），"
               f"请检查网络/权限后手动执行：pip install python-docx")


def ensure_docx(ask=True):
    """确保 python-docx 可导入；缺失时（经确认后）自动安装并重新导入。"""
    import importlib
    try:
        importlib.import_module("docx")
        return
    except ImportError:
        pass
    if not AUTO_INSTALL:
        raise ExtractError(3, "缺少依赖库 python-docx，请先安装：pip install python-docx")
    if ask:
        print("[依赖] 未检测到 python-docx（用于 Word 解析）。")
        if not _confirm("       是否现在自动安装？[Y/n] "):
            raise ExtractError(3, "已取消安装 python-docx，请手动执行：pip install python-docx")
    _pip_install()
    importlib.invalidate_caches()
    try:
        importlib.import_module("docx")
    except ImportError:
        raise ExtractError(
            3, "python-docx 已安装但仍无法导入，"
               f"可能装到了其他 Python 环境（当前解释器：{sys.executable}）")


def load_docx_api():
    """把 docx 的 Document / qn / Run 绑定到模块全局（首次调用时执行）。"""
    global Document, qn, Run
    if Document is not None:
        return
    ensure_docx()
    import importlib
    docx_mod = importlib.import_module("docx")
    Document = docx_mod.Document
    from docx.oxml.ns import qn as _qn
    from docx.text.run import Run as _Run
    qn = _qn
    Run = _Run


# ------------------------------------------------------------
# 转换器
# ------------------------------------------------------------

class DocxConverter:
    """单个 docx 文档的转换器（每个文件实例化一次）。"""

    def __init__(self, opts, doc, stem, out_dir):
        self.opts = opts
        self.doc = doc
        self.stem = stem                      # 输出 md 主名（图片文件名前缀）
        self.out_dir = Path(out_dir)
        self.img_dir = self.out_dir / "assets"
        self.img_seq = 0
        self.has_title = False
        self.numbering = self._load_numbering()
        self.num_counters = {}                # numId -> {ilvl: 当前计数}
        if opts.no_annotation:
            self.annotation_fonts = []
        else:
            fonts = opts.annotation_fonts or DEFAULT_ANNOTATION_FONTS
            self.annotation_fonts = [f.strip().lower()
                                    for f in fonts.split(",") if f.strip()]
        self.stats = {"blocks": 0, "annotation_runs": 0, "highlight_runs": 0,
                      "images": 0, "hyperlinks": 0, "empty_list_items": 0,
                      "empty_tables": 0, "empty_headings": 0, "quotes": 0}

    # ---------- 编号定义 ----------

    def _load_numbering(self):
        """读取编号定义：numId -> {ilvl: {"fmt": 格式, "start": 起始值}}。

        访问的是 python-docx 的半私有 API；不同版本可能变动，失败时按
        全部项目符号处理（仅影响编号列表的序号，不影响内容）。
        """
        info = {}
        try:
            numbering = self.doc.part.numbering_part.numbering_definitions._numbering
        except Exception:
            return info
        abstracts = {}
        for a in numbering.findall(qn("w:abstractNum")):
            levels = {}
            for lvl in a.findall(qn("w:lvl")):
                try:
                    ilvl = int(lvl.get(qn("w:ilvl")))
                except (TypeError, ValueError):
                    continue
                fmt = lvl.find(qn("w:numFmt"))
                start = lvl.find(qn("w:start"))
                levels[ilvl] = {
                    "fmt": fmt.get(qn("w:val")) if fmt is not None else "bullet",
                    "start": int(start.get(qn("w:val"))) if start is not None else 1,
                }
            abstracts[a.get(qn("w:abstractNumId"))] = levels
        for num in numbering.findall(qn("w:num")):
            nid = num.get(qn("w:numId"))
            a = num.find(qn("w:abstractNumId"))
            if nid is None or a is None:
                continue
            levels = abstracts.get(a.get(qn("w:val")))
            if levels:
                info[nid] = levels
        return info

    @staticmethod
    def _numpr_of(ppr):
        """从 pPr 中读取 numPr。返回 (ilvl, numId)、"cancel" 或 None。"""
        numpr = ppr.find(qn("w:numPr"))
        if numpr is None:
            return None
        numid = numpr.find(qn("w:numId"))
        if numid is None or numid.get(qn("w:val")) is None:
            return None
        nid = numid.get(qn("w:val"))
        if nid == "0":  # 显式取消编号
            return "cancel"
        ilvl_el = numpr.find(qn("w:ilvl"))
        try:
            ilvl = int(ilvl_el.get(qn("w:val"))) if ilvl_el is not None else 0
        except (TypeError, ValueError):
            ilvl = 0
        return max(ilvl, 0), nid

    def list_level(self, paragraph):
        """返回 (ilvl, numId)；非列表段落返回 (None, None)。

        编号可定义在段落直接格式上，也可定义在段落样式（含基样式）上，
        如 List Bullet / List Number。直接 numId=0 表示显式取消，优先级最高。
        """
        ppr = paragraph._p.find(qn("w:pPr"))
        if ppr is not None:
            got = self._numpr_of(ppr)
            if got == "cancel":
                return None, None
            if got is not None:
                return got
        style = paragraph.style
        seen = set()
        while style is not None:
            sid = getattr(style, "style_id", None)
            if sid is None or sid in seen:
                break
            seen.add(sid)
            el = getattr(style, "element", None)
            if el is not None:
                spPr = el.find(qn("w:pPr"))
                if spPr is not None:
                    got = self._numpr_of(spPr)
                    if got == "cancel":
                        return None, None
                    if got is not None:
                        return got
            style = getattr(style, "base_style", None)
        return None, None

    def list_marker(self, nid, ilvl):
        """列表项标记：项目符号返回 "-"，编号列表维护计数并返回 "N."。"""
        levels = self.numbering.get(nid) or {}
        spec = levels.get(ilvl)
        fmt = spec["fmt"] if spec else "bullet"
        if fmt in ("bullet", "none"):
            return "-"
        counters = self.num_counters.setdefault(nid, {})
        for k in [k for k in counters if k > ilvl]:  # 上层推进后深层重新计数
            del counters[k]
        start = spec["start"] if spec else 1
        counters[ilvl] = counters.get(ilvl, start - 1) + 1
        return f"{counters[ilvl]}."

    def list_indent(self, nid, ilvl):
        """嵌套缩进：项目符号层级 2 空格，编号层级 3 空格（CommonMark 安全值）。"""
        levels = self.numbering.get(nid) or {}
        indent = ""
        for i in range(ilvl):
            spec = levels.get(i)
            indent += "  " if (spec is None or spec["fmt"] == "bullet") else "   "
        return indent

    # ---------- run 层 ----------

    @staticmethod
    def _run_fonts(run):
        rpr = run._element.find(qn("w:rPr"))
        if rpr is None:
            return "", ""
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            return "", ""
        return (rf.get(qn("w:eastAsia")) or ""), (rf.get(qn("w:ascii")) or "")

    def is_annotation(self, run):
        """批注字体判定（精确匹配“楷体”，避免误伤“霞鹜文楷”之类的 AI 正文字体）。"""
        if not self.annotation_fonts:
            return False
        ea, asc = self._run_fonts(run)
        ea, asc = ea.lower(), asc.lower()
        return any(f in ea or f in asc for f in self.annotation_fonts)

    def is_math(self, run):
        """Cambria Math（Word 公式区域）判定。"""
        _, asc = self._run_fonts(run)
        return "cambria math" in asc.lower()

    @staticmethod
    def _resolved(run, attr, direct):
        """直接格式缺失时回退到字符样式（仅一层）。"""
        if direct is not None:
            return bool(direct)
        try:
            v = getattr(run.style.font, attr, None)
            return bool(v) if v is not None else False
        except Exception:
            return False

    def run_state(self, run):
        """run 的格式状态六元组。"""
        if self.is_math(run):
            bold = italic = False  # 公式字体的斜/粗是数学排版样式
        else:
            bold = self._resolved(run, "bold", run.bold)
            italic = self._resolved(run, "italic", run.italic)
        underline = self._resolved(run, "underline", run.underline)
        strike = self._resolved(run, "strike", run.font.strike)
        highlight = run.font.highlight_color is not None
        annotation = self.is_annotation(run)
        if annotation:
            self.stats["annotation_runs"] += 1
        if highlight:
            self.stats["highlight_runs"] += 1
        return (bold, italic, strike, underline, annotation, highlight)

    def run_images(self, run):
        """该 run 内的内嵌图片 → markdown 片段列表（同时保存图片文件）。"""
        out = []
        blips = run._element.findall(".//" + qn("a:blip"))
        if not blips:
            return out
        if self.opts.no_images:
            self.stats["images"] += len(blips)
            out.append("<!-- 图片未提取 -->")
            return out
        for blip in blips:
            rid = blip.get(qn("r:embed"))
            link = blip.get(qn("r:link"))
            if not rid and not link:
                continue
            try:
                rel = run.part.rels[rid or link]
            except Exception:
                continue
            if rel.is_external:
                self.stats["images"] += 1
                out.append(f"![图]({rel.target_ref})")
                continue
            if not rid:
                continue
            try:
                part = rel.target_part
                blob = part.blob
                ext = Path(str(part.partname)).suffix or ".png"
            except Exception:
                continue
            self.img_seq += 1
            name = f"{self.stem}-img{self.img_seq}{ext}"
            self.img_dir.mkdir(parents=True, exist_ok=True)
            (self.img_dir / name).write_bytes(blob)
            rel_path = f"assets/{name}"
            if " " in rel_path:
                rel_path = f"<{rel_path}>"
            self.stats["images"] += 1
            out.append(f"![图{self.img_seq}]({rel_path})")
        return out

    @staticmethod
    def _wrap(state, text):
        """按状态给文本包 markdown 标记；state=None 表示无需包裹（图片等片段）。"""
        if state is None:
            return text
        bold, italic, strike, underline, annotation, highlight = state
        if annotation:
            text = GRAY_MARK.format(text)
        elif highlight:
            text = HL_MARK.format(text)
        if bold and italic:
            text = f"***{text}***"
        elif bold:
            text = f"**{text}**"
        elif italic:
            text = f"*{text}*"
        if strike:
            text = f"~~{text}~~"
        if underline:
            text = U_TAG.format(text)
        return text

    def runs_to_md(self, runs):
        """合并相邻同格式 run，输出行内 markdown。"""
        pieces = []  # (state, text)；state=None 表示图片等不可合并片段
        for r in runs:
            for img in self.run_images(r):
                pieces.append((None, img))
            if not r.text:
                continue
            pieces.append((self.run_state(r), r.text))
        merged = []
        for state, text in pieces:
            if state is not None and merged and merged[-1][0] == state:
                merged[-1][1] += text
            else:
                merged.append([state, text])
        return "".join(self._wrap(state, text) for state, text in merged)

    # ---------- 段落层 ----------

    def hyperlink_md(self, element, paragraph):
        """w:hyperlink → [文本](网址)。"""
        runs = [Run(r, paragraph) for r in element.findall(qn("w:r"))]
        text = self.runs_to_md(runs)
        rid = element.get(qn("r:id"))
        anchor = element.get(qn("w:anchor"))
        url = None
        if rid:
            try:
                rel = paragraph.part.rels[rid]
                url = rel.target_ref if rel.is_external else None
            except Exception:
                url = None
        elif anchor:
            url = "#" + anchor
        if url:
            self.stats["hyperlinks"] += 1
            if " " in url:
                url = f"<{url}>"
            return f"[{text or url}]({url})"
        return text

    def para_content(self, paragraph):
        """按文档顺序处理段落中的 run / 超链接 / 修订插入，输出行内 markdown。"""
        parts = []
        batch = []

        def flush():
            if batch:
                parts.append(self.runs_to_md(batch))
                batch.clear()

        for child in paragraph._p.iterchildren():
            tag = child.tag
            if tag == qn("w:r"):
                batch.append(Run(child, paragraph))
            elif tag == qn("w:hyperlink"):
                flush()
                parts.append(self.hyperlink_md(child, paragraph))
            elif tag in (qn("w:ins"), qn("w:moveTo")):  # 修订插入/移入的内容保留
                for sub in child.iterchildren():
                    if sub.tag == qn("w:r"):
                        batch.append(Run(sub, paragraph))
                    elif sub.tag == qn("w:hyperlink"):
                        flush()
                        parts.append(self.hyperlink_md(sub, paragraph))
            # 其余（w:del、批注锚点等）跳过
        flush()
        # 软回车（python-docx 把 w:br 读作 \n）→ <br>
        return "".join(parts).replace("\n", "<br>")

    def para_to_blocks(self, paragraph, index):
        """单个段落 → 块字符串列表（通常 0 或 1 个块）。"""
        style = paragraph.style.name if paragraph.style is not None else ""
        raw = "".join(t.text or "" for t in paragraph._p.iter(qn("w:t")))
        text = raw.strip()

        # 1) 标题样式
        m = HEADING_STYLE_RE.match(style)
        if m:
            if not text:
                self.stats["empty_headings"] += 1
                return []
            level = min(int(m.group(1)) + (1 if self.has_title else 0), 6)
            return ["#" * level + " " + text]

        # 2) 文档主标题（首段约定）
        if (index == 0 and text and not self.opts.no_title_shift
                and style in TITLE_STYLES):
            self.has_title = True
            return ["# " + text]

        # 3) 生成信息行（第二段约定；仅在有主标题时生效）
        if (index == 1 and text and self.has_title
                and not self.opts.no_meta_italic and len(text) <= META_LINE_MAX):
            if self.list_level(paragraph)[0] is None:
                return ["*" + text + "*"]

        # 4) 列表项
        lvl, nid = self.list_level(paragraph)
        if lvl is not None:
            inner = self.para_content(paragraph)
            if not inner.strip():
                self.stats["empty_list_items"] += 1
                return []
            return [self.list_indent(nid, lvl) + self.list_marker(nid, lvl) + " " + inner]

        # 5) 引用样式
        if style in QUOTE_STYLES:
            inner = self.para_content(paragraph)
            if not inner.strip():
                return []
            self.stats["quotes"] += 1
            return ["\n".join("> " + ln for ln in inner.split("\n"))]

        # 6) 普通段落（含纯图片段落）
        if not text:
            inner = self.para_content(paragraph)
            return [inner] if inner.strip() else []
        return [self.para_content(paragraph)]

    # ---------- 表格 ----------

    def cell_md(self, cell):
        """单元格 → 单行 markdown（多段落以 <br> 连接）。"""
        parts = []
        for p in cell.paragraphs:
            inner = self.para_content(p)
            if inner.strip():
                parts.append(inner)
        return "<br>".join(parts).replace("|", "\\|")

    def table_to_block(self, table):
        """表格 → markdown pipe 表格；合并单元格去重；全空返回 None。

        去重通过在内存中的 tc 元素上打自定义命名空间标记实现：
        lxml 的元素代理可能被回收后重建，id() 不稳定，不能作去重键；
        属性写在 XML 树本身上，稳定可靠，且仅存在于内存副本中，不会落盘。
        """
        rows = []
        for row in table.rows:
            cells = []
            for cell in row.cells:
                tc = cell._tc
                if tc.get(VISITED_ATTR) == "1":  # 水平/垂直合并的重复单元格
                    continue
                tc.set(VISITED_ATTR, "1")
                cells.append(self.cell_md(cell))
            if cells:
                rows.append(cells)
        if not rows or not any(any(c.strip() for c in r) for r in rows):
            self.stats["empty_tables"] += 1
            return None
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        lines = ["| " + " | ".join(rows[0]) + " |",
                 "| " + " | ".join(["---"] * ncol) + " |"]
        for r in rows[1:]:
            lines.append("| " + " | ".join(r) + " |")
        return "\n".join(lines)


# ------------------------------------------------------------
# 拼装与写出
# ------------------------------------------------------------

def assemble(blocks):
    """块间排版：连续列表项紧凑排列，相邻引用块合并，其余块之间空一行。"""
    out = []
    for b in blocks:
        if not out:
            out.append(b)
            continue
        prev = out[-1]
        prev_is_list = LIST_BLOCK_RE.match(prev.split("\n", 1)[0])
        cur_is_list = LIST_BLOCK_RE.match(b.split("\n", 1)[0])
        if prev_is_list and cur_is_list:
            out.append(b)
        elif prev.startswith("> ") and b.startswith("> "):
            out[-1] = prev + "\n" + b
        else:
            out.append("")
            out.append(b)
    return "\n".join(out).strip() + "\n"


def output_paths(src, opts):
    """返回 (输出主名, 输出目录, 输出 md 路径)。"""
    stem = src.stem
    if opts.strip_version:
        stem = re.sub(r"[-_]v[0-9.]*$", "", stem)
    if opts.suffix:
        stem = f"{stem}-{opts.suffix}"
    out_dir = Path(opts.out) if opts.out else src.parent
    return stem, out_dir, out_dir / f"{stem}.md"


def convert_file(src, opts):
    """处理单个文件：判定类型 → 转换 → 写出 {名}.md。"""
    src = Path(src)
    suffix = src.suffix.lower()
    if suffix == ".doc":
        raise ExtractError(1, f"不支持旧版 .doc（{src.name}），请先用 Word 另存为 .docx")
    if suffix != ".docx":
        raise ExtractError(1, f"不支持的文件类型：{src.name}（仅支持 .docx）")

    load_docx_api()
    doc = Document(str(src))
    stem, out_dir, out_md = output_paths(src, opts)
    conv = DocxConverter(opts, doc, stem, out_dir)

    # 清理上次运行可能残留的同前缀图片文件，避免重复运行时累积
    if not opts.no_images and conv.img_dir.is_dir():
        for old in conv.img_dir.glob(glob.escape(stem) + "-img*"):
            try:
                old.unlink()
            except OSError:
                pass

    blocks = []
    para_index = 0
    for item in doc.iter_inner_content():
        kind = item.__class__.__name__
        if kind == "Paragraph":
            blocks.extend(conv.para_to_blocks(item, para_index))
            para_index += 1
        elif kind == "Table":
            block = conv.table_to_block(item)
            if block is not None:
                blocks.append(block)

    conv.stats["blocks"] = len(blocks)
    content = assemble(blocks)
    if not content.strip():
        raise ExtractError(2, f"转换结果为空：{src.name}")

    out_dir.mkdir(parents=True, exist_ok=True)
    out_md.write_text(content, encoding="utf-8")

    # 汇报统计信息，供调用方（模型/用户）核对转换质量
    print(f"[成功] {src.name} → {out_md}")
    s = conv.stats
    notes = []
    if s["annotation_runs"]:
        notes.append(f"批注字体 {s['annotation_runs']} 处")
    if s["highlight_runs"]:
        notes.append(f"荧光高亮 {s['highlight_runs']} 处")
    if s["images"]:
        notes.append(f"图片 {s['images']} 张")
    if s["hyperlinks"]:
        notes.append(f"超链接 {s['hyperlinks']} 个")
    if s["quotes"]:
        notes.append(f"引用段 {s['quotes']} 个")
    if s["empty_list_items"]:
        notes.append(f"空列表项 {s['empty_list_items']} 个已跳过")
    if s["empty_tables"]:
        notes.append(f"空表格 {s['empty_tables']} 个已跳过")
    if s["empty_headings"]:
        notes.append(f"空标题 {s['empty_headings']} 个已跳过")
    if notes:
        print("       " + "，".join(notes))


# ------------------------------------------------------------
# 主入口
# ------------------------------------------------------------

def parse_args(argv):
    ap = argparse.ArgumentParser(
        description="Word（.docx）→ Markdown 转换（ordinote-extract）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python extract_docx.py 笔记.docx\n"
            "  python extract_docx.py 笔记.docx --suffix 提取 --out 提取稿目录\n"
            "  python extract_docx.py \"archives/纵向笔记（临时）\" --strip-version --out \"纵向笔记\"\n"
        ))
    ap.add_argument("inputs", nargs="+", metavar="文件或目录",
                    help=".docx 文件路径；目录则处理其下所有 .docx（不递归）")
    ap.add_argument("--out", help="输出目录（默认：源文件所在目录）")
    ap.add_argument("--suffix", help="输出文件名后缀：{名}-{后缀}.md")
    ap.add_argument("--strip-version", action="store_true",
                    help="去除文件名末尾 _v / _v0.5 之类版本号")
    ap.add_argument("--annotation-fonts", default=DEFAULT_ANNOTATION_FONTS,
                    help=f"批注字体，逗号分隔（默认：{DEFAULT_ANNOTATION_FONTS}）")
    ap.add_argument("--no-annotation", action="store_true",
                    help="不做批注字体标记")
    ap.add_argument("--no-title-shift", action="store_true",
                    help="首段不作为文档主标题，标题保持原层级")
    ap.add_argument("--no-meta-italic", action="store_true",
                    help="第二段不做整行斜体处理")
    ap.add_argument("--no-images", action="store_true",
                    help="不导出内嵌图片（以 HTML 注释占位）")
    ap.add_argument("--no-auto-install", action="store_true",
                    help="缺少 python-docx 时不自动安装")
    ap.add_argument("--install-deps", action="store_true",
                    help="只安装依赖后退出")
    return ap.parse_args(argv)


def main(argv=None):
    global AUTO_INSTALL

    # Windows 控制台默认可能是 GBK，强制 UTF-8 避免中文输出乱码/报错
    for stream in (sys.stdout, sys.stderr):
        if stream and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    opts = parse_args(argv)

    if opts.no_auto_install:
        AUTO_INSTALL = False
    if opts.install_deps:
        try:
            ensure_docx(ask=False)
        except ExtractError as e:
            print(f"[失败] {e.message}")
            return e.code
        print("[完成] 依赖 python-docx 已就绪。")
        return 0

    # 收集文件（目录自动展开其中的 .docx）
    files = []
    worst = 0
    for arg in opts.inputs:
        p = Path(arg)
        if p.is_dir():
            found = sorted(p.glob("*.docx"))
            if not found:
                print(f"[提醒] 目录中没有 .docx 文件：{p}")
            files.extend(found)
        elif p.is_file():
            files.append(p)
        else:
            print(f"[失败] 文件不存在：{p}")
            worst = max(worst, 1)
    if not files:
        print("[失败] 没有可处理的 .docx 文件")
        return worst or 1

    for src in files:
        try:
            convert_file(src, opts)
        except ExtractError as e:
            print(f"[失败] {e.message}")
            worst = max(worst, e.code)
        except Exception as e:  # 文件损坏等未预期错误
            print(f"[失败] 转换出错：{src.name}（{e}）")
            worst = max(worst, 2)
    return worst


if __name__ == "__main__":
    sys.exit(main())
