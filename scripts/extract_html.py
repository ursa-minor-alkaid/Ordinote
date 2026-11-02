#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""本地 HTML 幻灯片文字提取脚本（ordinote-extract 功能）

功能：
    从自包含 HTML 幻灯片中逐页提取文字，生成 markdown 提取文件。
    提取文件是"原始素材"，后续按 references/ordinote-wash.md 的
    规则做整理（修错别字、去空格换行等）。

用法：
    python extract_html.py <文件路径> [更多文件路径...]

输出：
    在源文件同目录下生成 {原文件名}-提取.md
    - 各张幻灯片之间以 --- 分隔线分开（保留源结构线索）
    - 无文字幻灯片以 HTML 注释占位，便于核对缺页
    - 文件头以 HTML 注释记录来源与提取时间
    - 同名旧提取文件会被直接覆盖

识别规则：
    class 中含有独立单词 slide 的 section / div / article 视为一张幻灯片；
    style、script、noscript、template 的内容丢弃；
    块级元素换行，td/th 以 " | " 分隔，img 记为 [IMG] 或 [IMG:alt]；
    找不到幻灯片容器时，退化为整份 body 作为单页输出。

只依赖标准库。

退出码：
    0  全部提取成功
    1  用法错误 / 文件不存在 / 不支持的文件类型
    2  提取失败或结果为空
"""

import re
import sys
from datetime import datetime
from html.parser import HTMLParser
from pathlib import Path

SLIDE_TAGS = {"section", "div", "article"}
BLOCK_TAGS = {
    "address", "article", "aside", "blockquote", "dd", "div", "dl", "dt",
    "fieldset", "figcaption", "figure", "footer", "form", "h1", "h2", "h3",
    "h4", "h5", "h6", "header", "hr", "li", "main", "nav", "ol", "p", "pre",
    "section", "table", "tbody", "thead", "tfoot", "tr", "ul",
}
SKIP_TAGS = {"style", "script", "noscript", "template"}
VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
    "meta", "param", "source", "track", "wbr",
}

# 各张幻灯片之间的分隔线；后续整理时会保留它作为结构线索
PAGE_SEPARATOR = "\n\n---\n\n"

# 文件头注释：记录来源信息。HTML 注释不会被渲染，仅供追溯
HEADER_TEMPLATE = (
    "<!-- 本文件由 extract_html.py 自动提取生成 -->\n"
    "<!-- 源文件：{name} | 页数：{count} | 提取时间：{time} -->\n"
)

# 是否保留图片的 alt 文本（--no-images 可关闭）
KEEP_IMG_ALT = True


class ExtractError(Exception):
    """单个文件提取失败时抛出，携带退出码与给用户看的说明。

    多文件批量提取时，单个文件失败只记录错误、不中断其他文件。
    """

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


class SlideParser(HTMLParser):
    """按容器切分幻灯片并收集每页文本。"""

    def __init__(self, body_present=True, force_single=False, keep_img_alt=True):
        super().__init__(convert_charrefs=True)
        self.body_present = body_present
        self.force_single = force_single
        self.keep_img_alt = keep_img_alt
        self.stack = []
        self.slides = []
        self.cur_tag = None
        self.cur_depth = 0
        self.cur_cls = ""
        self.buf = []
        self.skip_tag = None
        self.in_body = not body_present
        self.row_cells = 0

    # ---- 标签处理 ----

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        cls = a.get("class", "")

        if tag in SKIP_TAGS:
            self.skip_tag = tag

        if tag == "body":
            self.in_body = True
            if self.force_single:
                self._start_slide("body", "body")

        active = self.in_body and self.skip_tag is None

        if active and self.cur_tag is None and not self.force_single:
            if tag in SLIDE_TAGS and "slide" in cls.split():
                self._start_slide(tag, cls)

        if tag == "br":
            self._emit("\n")
        elif tag == "img":
            if self.cur_tag is not None:
                alt = a.get("alt", "").strip()
                self._emit(f"[IMG:{alt}]" if (alt and self.keep_img_alt) else "[IMG]")
        elif tag == "tr":
            self.row_cells = 0
        elif tag in ("td", "th"):
            if self.cur_tag is not None:
                if self.row_cells:
                    self._emit(" | ")
                self.row_cells += 1
        elif tag in BLOCK_TAGS:
            self._emit("\n")

        if tag not in VOID_TAGS:
            self.stack.append(tag)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID_TAGS:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if tag in VOID_TAGS:
            return
        if self.skip_tag == tag:
            self.skip_tag = None

        if tag in self.stack:
            while self.stack:
                if self.stack.pop() == tag:
                    break

        if self.cur_tag is not None and tag == self.cur_tag and len(self.stack) == self.cur_depth:
            self._finish_slide()

        if tag == "body" and self.force_single and self.cur_tag is not None:
            self._finish_slide()

        if tag in BLOCK_TAGS:
            self._emit("\n")
        if tag == "tr":
            self.row_cells = 0

    def handle_data(self, data):
        if self.skip_tag is None and self.cur_tag is not None:
            self.buf.append(data)

    # ---- 状态维护 ----

    def _start_slide(self, tag, cls):
        self.cur_tag = tag
        self.cur_cls = cls
        self.cur_depth = len(self.stack)
        self.buf = []

    def _finish_slide(self):
        text = "".join(self.buf)
        text = text.replace("\r", "")
        text = re.sub(r"[ \t\u00a0]+", " ", text)
        lines = (line.strip() for line in text.split("\n"))
        text = "\n".join(line for line in lines if line)
        self.slides.append((self.cur_cls, text))
        self.cur_tag = None
        self.cur_cls = ""
        self.buf = []

    def _emit(self, chunk):
        if self.cur_tag is not None and self.skip_tag is None:
            self.buf.append(chunk)


def extract(html, keep_img_alt=True):
    """返回 [(class, text), ...]；无幻灯片容器时按单页整文输出。"""
    body_present = "<body" in html.lower()
    parser = SlideParser(body_present=body_present, keep_img_alt=keep_img_alt)
    parser.feed(html)
    parser.close()

    if not parser.slides and body_present:
        fallback = SlideParser(body_present=body_present, force_single=True,
                               keep_img_alt=keep_img_alt)
        fallback.feed(html)
        fallback.close()
        parser.slides = fallback.slides

    return parser.slides


def assemble(path, slides):
    """把逐页文本拼成最终的 markdown 文本。"""
    header = HEADER_TEMPLATE.format(
        name=path.name,
        count=len(slides),
        time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    # 空页用注释占位：不影响渲染，又能核对哪一页缺了内容
    parts = []
    for i, (_, text) in enumerate(slides, 1):
        text = text.strip()
        parts.append(text if text else f"<!-- 第 {i} 张幻灯片未提取到文字 -->")
    return header + "\n" + PAGE_SEPARATOR.join(parts) + "\n"


def process_one(arg):
    """处理单个文件：读入 → 提取 → 写出 {原文件名}-提取.md。"""
    path = Path(arg)
    if not path.is_file():
        raise ExtractError(1, f"文件不存在：{path}")

    if path.suffix.lower() not in (".html", ".htm"):
        raise ExtractError(1, f"不支持的文件类型：{path.name}（仅支持 .html / .htm）")

    try:
        html = path.read_text(encoding="utf-8-sig", errors="replace")
    except OSError as e:
        raise ExtractError(2, f"读取失败：{path.name}（{e}）")

    slides = extract(html, keep_img_alt=KEEP_IMG_ALT)
    if not any(text.strip() for _, text in slides):
        raise ExtractError(
            2, f"HTML 未提取到文字：{path.name}；"
               "若内容由脚本动态渲染，请改用浏览器打开后另存，或改用 MinerU API"
        )

    content = assemble(path, slides)
    out = path.with_name(f"{path.stem}-提取.md")  # 输出到源文件同目录
    try:
        out.write_text(content, encoding="utf-8")
    except OSError as e:
        raise ExtractError(2, f"写出失败：{out.name}（{e}）")

    # 汇报统计信息，供调用方（模型/用户）核对提取质量
    chars = sum(len(text) for _, text in slides)
    print(f"[成功] {path.name} → {out.name}（{len(slides)} 页，共 {chars} 字）")
    empty = [i for i, (_, text) in enumerate(slides, 1) if not text.strip()]
    if empty:
        print(f"[提醒] 以下页面未提取到文字，可能为纯图片页："
              f"{'、'.join(map(str, empty))}")


def main():
    global KEEP_IMG_ALT

    # Windows 控制台默认可能是 GBK，强制 UTF-8 避免中文输出乱码/报错
    for stream in (sys.stdout, sys.stderr):
        if stream and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    args = sys.argv[1:]

    if "-h" in args or "--help" in args:
        print(__doc__)
        sys.exit(0)
    if "--no-images" in args:
        KEEP_IMG_ALT = False
        args = [a for a in args if a != "--no-images"]

    if not args:
        print(__doc__)  # 打印用法说明
        print("[失败] 请提供至少一个 HTML 文件路径")
        sys.exit(1)

    worst = 0  # 记录最严重的退出码；全部成功时保持 0
    for arg in args:
        try:
            process_one(arg)
        except ExtractError as e:
            print(f"[失败] {e.message}")
            worst = max(worst, e.code)
    sys.exit(worst)


if __name__ == "__main__":
    main()
