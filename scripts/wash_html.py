#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""通用 HTML 幻灯片清洗脚本（ordinote-organizer 配套工具）。

定位：
    技能内 scripts/extract_html.py 负责通用的逐页纯文本提取；本脚本在其之后承担
    “清洗”这一步，按 HTML 语义标记重建笔记结构，并把页内资源落盘。对类名规范的
    幻灯片（如本课程模板）可直接替代“提取 + 人工清洗”两步；对结构松散的页面
    仍可用，只是层级还原依赖标签本身。

处理内容：
    - 分页：class 含独立单词 slide 的 section / div / article 视为一页，否则依次
      退回多 section 结构、整篇一页；页间以 --- 分隔
    - 去噪：页脚、导航、termbar 等装饰不写入正文；页码只在标题编号中保留
    - 结构：标题层级、导读句、要点列表、表格（跨列行转为表后注释）、代码块、
      图注、定义列表、引用、折叠块按语义还原
    - 资源：data URI 图片解码落盘，内联 SVG 存为独立文件并恢复驼峰属性，
      相对路径图片按需复制；正文用相对路径引用
    - 代码块：优先用 class="language-xxx"，否则按首行与命令词判定 bash / text

输入：
    一个或多个 .html 文件，或一个目录（处理其中的 *.html）。

输出：
    默认写在源文件同目录的 {原文件名}-整理.md；--outdir 可改输出目录；--out 可
    指定单个输出文件。图片写入 {assets}/images/{前缀}-figs/，SVG 写入
    {assets}/svgs/{前缀}-svgs/，前缀默认取文件名开头的编号（如 L2），否则 deck。

用法：
    python wash_html.py "deck.html"
    python wash_html.py "resources/slides/L2-xxx.html" --outdir .
    python wash_html.py "resources/slides" --outdir . --skip-existing
"""

import argparse
import base64
import re
import shutil
import sys
from html.parser import HTMLParser
from pathlib import Path

VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input", "link",
    "meta", "param", "source", "track", "wbr",
}
SKIP_TAGS = {"style", "script", "noscript", "template", "nav"}
SKIP_CLASSES = {"footer", "termbar"}
TITLE_CLASSES = {"darktitle", "t", "title"}
KICK_CLASSES = {"kick", "stkick"}
TAKEAWAY_CLASSES = {"takeaway", "darksub", "stsub", "sttext"}
PART_CLASSES = {"dark", "part", "divider", "chapter", "section-break"}
COMMANDS = {
    "spades.py", "blastn", "makeblastdb", "conda", "grep", "head", "awk",
    "python3", "nproc", "free", "cd", "ls", "samtools", "bwa", "git",
    "pip", "python", "docker", "curl", "wget", "Rscript", "seqkit", "fastqc",
}
CAMEL_ATTRS = {
    "viewbox": "viewBox",
    "preserveaspectratio": "preserveAspectRatio",
    "gradientunits": "gradientUnits",
    "gradienttransform": "gradientTransform",
    "patternunits": "patternUnits",
    "patterncontentunits": "patternContentUnits",
    "clippathunits": "clipPathUnits",
    "maskunits": "maskUnits",
    "markerunits": "markerUnits",
    "refx": "refX",
    "refy": "refY",
    "startoffset": "startOffset",
    "textlength": "textLength",
    "lengthadjust": "lengthAdjust",
    "basefrequency": "baseFrequency",
    "numoctaves": "numOctaves",
    "filterunits": "filterUnits",
    "primitiveunits": "primitiveUnits",
    "kernelmatrix": "kernelMatrix",
    "tablevalues": "tableValues",
    "pathlength": "pathLength",
}
STOP_WORDS = {"a", "an", "the", "of", "to", "for", "in", "and", "or", "on",
              "with", "vs", "at", "by", "from"}


class Node:
    def __init__(self, tag, attrs=None, parent=None):
        self.tag = tag
        self.attrs = dict(attrs or {})
        self.children = []
        self.parent = parent

    @property
    def classes(self):
        return set(self.attrs.get("class", "").split())


class TreeBuilder(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.root = Node("#document")
        self.stack = [self.root]

    def handle_starttag(self, tag, attrs):
        node = Node(tag, attrs, self.stack[-1])
        self.stack[-1].children.append(node)
        if tag not in VOID_TAGS:
            self.stack.append(node)

    def handle_startendtag(self, tag, attrs):
        node = Node(tag, attrs, self.stack[-1])
        self.stack[-1].children.append(node)

    def handle_endtag(self, tag):
        for i in range(len(self.stack) - 1, 0, -1):
            if self.stack[i].tag == tag:
                del self.stack[i:]
                break

    def handle_data(self, data):
        self.stack[-1].children.append(data)


def walk(node):
    for child in node.children:
        if isinstance(child, Node):
            yield child
            yield from walk(child)


def plain(node):
    if isinstance(node, str):
        return node
    if node.tag in SKIP_TAGS:
        return ""
    return "".join(plain(c) for c in node.children)


def squeeze(text):
    return re.sub(r"[ \t\u00a0]+", " ", text).replace("\n", " ").strip()


def titlecase(text):
    words = text.split()
    out = []
    for i, word in enumerate(words):
        low = word.lower()
        if i and low in STOP_WORDS:
            out.append(low)
        else:
            out.append(low[:1].upper() + low[1:])
    return " ".join(out)


class Renderer:
    def __init__(self, html_path, out_path, assets_root, prefix):
        self.html_path = html_path
        self.out_path = out_path
        self.assets_root = assets_root
        self.prefix = prefix
        self.img_dir = assets_root / "images" / f"{prefix}-figs"
        self.svg_dir = assets_root / "svgs" / f"{prefix}-svgs"
        self.counter = 0
        self.saved = []
        self.missing = []
        self.skip_nodes = set()
        self.base_level = 3

    # ---------- 资源落盘 ----------
    def _next_name(self, ext):
        self.counter += 1
        return f"{self.prefix}-{self.counter}.{ext}"

    def rel_to_out(self, path):
        try:
            return Path(path).resolve().relative_to(self.out_path.parent.resolve()).as_posix()
        except ValueError:
            import os
            return Path(os.path.relpath(path, self.out_path.parent)).as_posix()

    def save_img(self, node):
        src = (node.attrs.get("src") or "").strip()
        if not src:
            return None
        if src.startswith("data:image/"):
            m = re.match(r"data:image/(\w+);base64,(.*)$", src, re.S)
            if not m:
                return None
            kind, payload = m.group(1).lower(), m.group(2)
            ext = "jpg" if kind in ("jpeg", "jpg") else kind
            self.img_dir.mkdir(parents=True, exist_ok=True)
            name = self._next_name(ext)
            (self.img_dir / name).write_bytes(base64.b64decode(payload))
            rel = self.rel_to_out(self.img_dir / name)
            self.saved.append(rel)
            return rel
        if re.match(r"(https?|data):", src):
            return src
        source = (self.html_path.parent / src).resolve()
        if not source.is_file():
            self.missing.append(src)
            return None
        self.img_dir.mkdir(parents=True, exist_ok=True)
        name = self._next_name(source.suffix.lstrip(".").lower() or "png")
        shutil.copy2(source, self.img_dir / name)
        rel = self.rel_to_out(self.img_dir / name)
        self.saved.append(rel)
        return rel

    def save_svg(self, node):
        raw = self.raw_html(node)
        for low, camel in CAMEL_ATTRS.items():
            raw = re.sub(rf" {low}=", f" {camel}=", raw)
        if "xmlns=" not in raw:
            raw = raw.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"', 1)
        self.svg_dir.mkdir(parents=True, exist_ok=True)
        name = self._next_name("svg")
        (self.svg_dir / name).write_text(raw + "\n", encoding="utf-8")
        rel = self.rel_to_out(self.svg_dir / name)
        self.saved.append(rel)
        return rel

    def raw_html(self, node):
        attrs = "".join(f' {k}="{v}"' for k, v in node.attrs.items())
        body = "".join(self.raw_html(c) if isinstance(c, Node) else c
                       for c in node.children)
        return f"<{node.tag}{attrs}>{body}</{node.tag}>"

    # ---------- 行内 ----------
    def inline(self, node):
        out = []
        for child in node.children if isinstance(node, Node) else []:
            if isinstance(child, str):
                out.append(re.sub(r"\s+", " ", child))
                continue
            if child.tag in SKIP_TAGS:
                continue
            cls = child.classes
            if child.tag in ("b", "strong"):
                out.append(f"**{self.inline(child).strip()}**")
            elif child.tag in ("i", "em"):
                out.append(f"*{self.inline(child).strip()}*")
            elif "mono" in cls:
                out.append(f"`{squeeze(plain(child))}`")
            elif child.tag == "br":
                out.append(" ")
            elif child.tag == "a":
                label = self.inline(child).strip()
                href = child.attrs.get("href", "").strip()
                if label and href and not href.startswith("#"):
                    out.append(f"[{label}]({href})")
                else:
                    out.append(label)
            else:
                out.append(self.inline(child))
        text = "".join(out)
        text = re.sub(r"[ \t\u00a0]+", " ", text)
        text = re.sub(r" (?=[.,;:!?])", "", text)
        return text.strip()

    def text(self, node):
        return squeeze(plain(node))

    # ---------- 块级 ----------
    def block(self, child):
        cls = child.classes
        if id(child) in self.skip_nodes:
            return ""
        if child.tag in SKIP_TAGS or child.tag == "footer" or cls & SKIP_CLASSES:
            return ""
        if "figcap" in cls or child.tag == "figcaption":
            return ""
        if "figbox" in cls or child.tag == "figure":
            return self.figure(child)
        if child.tag == "svg":
            return self.standalone_svg(child)
        if "codewrap" in cls or child.tag == "pre":
            return self.code(child)
        if "kpirow" in cls:
            return self.kpis(child)
        if "callout" in cls:
            return self.callout(child)
        if child.tag == "table" or "tb" in cls:
            return self.table(child)
        if child.tag in ("ul", "ol"):
            return self.listing(child, 0)
        if child.tag == "dl":
            return self.deflist(child)
        if "steprow" in cls:
            return self.step(child)
        if "tcell" in cls:
            return self.term_cell(child)
        if "dcell" in cls:
            return self.def_cell(child)
        if "tocrow" in cls or "toc" in cls:
            return ""
        if "kick" in cls:
            return f"**{self.text(child)}**"
        if child.tag == "blockquote":
            body = "\n\n".join(b for b in self.blocks(child) if b)
            return "\n".join("> " + line if line else ">" for line in body.split("\n"))
        if child.tag == "details":
            summary = ""
            for sub in child.children:
                if isinstance(sub, Node) and sub.tag == "summary":
                    summary = self.inline(sub)
            body = "\n\n".join(b for b in self.blocks(child)
                               if b and not b.startswith(summary))
            return f"**{summary}**\n\n{body}".strip() if summary else body
        if child.tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
            label = self.text(child)
            if not label:
                return ""
            if child.tag in ("h4", "h5", "h6"):
                return f"**{label}**"
            level = min(6, self.base_level + 1)
            return f"{'#' * level} {label}"
        if child.tag == "p":
            return self.inline(child)
        if child.tag == "img":
            rel = self.save_img(child)
            return f"![]({rel})" if rel else ""
        if child.tag == "hr":
            return ""
        chunks = [b for b in self.blocks(child) if b]
        if chunks and all(c.startswith("- ") for c in chunks):
            return "\n".join(chunks)
        if chunks and all(re.match(r"^\d+\. ", c) for c in chunks):
            return "\n".join(chunks)
        return "\n\n".join(chunks)

    def blocks(self, node):
        result = []
        nodes = node.children
        skip = set()
        for i, child in enumerate(nodes):
            if isinstance(child, str) or i in skip:
                continue
            chunk = self.block(child)
            if not chunk:
                continue
            if child.tag in ("svg", "img"):
                for j in range(i + 1, len(nodes)):
                    nxt = nodes[j]
                    if isinstance(nxt, str):
                        if nxt.strip():
                            break
                        continue
                    if "figcap" in nxt.classes or nxt.tag == "figcaption":
                        cap = self.inline(nxt).lstrip("· ").strip()
                        if cap:
                            chunk = f"{chunk}\n\n*{cap}*"
                        skip.add(j)
                        break
                    break
            result.append(chunk)
        return result

    def figure(self, node):
        parts = []
        caption = ""
        for child in node.children:
            if isinstance(child, Node) and ("figcap" in child.classes
                                            or child.tag == "figcaption"):
                caption = self.inline(child)
        if not caption and node.tag != "figure":
            for child in walk(node):
                if "figcap" in child.classes or child.tag == "figcaption":
                    caption = self.inline(child)
                    break
        for child in walk(node):
            if child.tag == "img":
                rel = self.save_img(child)
                if rel:
                    parts.append(f"![]({rel})")
            elif child.tag == "svg":
                parts.append(self.standalone_svg(child))
        if caption:
            caption = caption.lstrip("· ").strip()
            parts.append(f"*{caption}*")
        return "\n\n".join(p for p in parts if p)

    def standalone_svg(self, node):
        return f"![]({self.save_svg(node)})"

    def code(self, node):
        lines = []
        buf = []

        def collect(n):
            for c in n.children:
                if isinstance(c, str):
                    buf.append(c)
                elif c.tag in SKIP_TAGS or "termbar" in c.classes:
                    continue
                elif c.tag == "br":
                    buf.append("\n")
                else:
                    collect(c)

        collect(node)
        for line in "".join(buf).split("\n"):
            line = line.replace("\u00a0", " ").rstrip()
            if line.strip():
                lines.append(line.rstrip())
        body = "\n".join(lines)
        return f"```{self.code_lang(node, lines)}\n{body}\n```"

    def code_lang(self, node, lines):
        for child in [node] + list(walk(node)):
            for token in child.classes:
                m = re.match(r"language-(\w+)", token) or re.match(r"lang-(\w+)", token)
                if m:
                    return m.group(1)
        if not lines:
            return "text"
        if any(line.lstrip().startswith((">", "@")) for line in lines):
            return "text"
        for line in lines:
            head = line.strip()
            if head.startswith("#") or head.endswith((".", ",")):
                continue
            if head.split()[0] in COMMANDS:
                return "bash"
        return "text"

    def kpis(self, node):
        out = []
        for child in node.children:
            if isinstance(child, Node) and "kpi" in child.classes:
                value = label = ""
                for sub in child.children:
                    if isinstance(sub, Node) and "v" in sub.classes:
                        value = self.inline(sub)
                    elif isinstance(sub, Node) and "k" in sub.classes:
                        label = self.inline(sub)
                if value and label:
                    out.append(f"- **{value}**: {label}")
                elif value:
                    out.append(f"- **{value}**")
        return "\n".join(out)

    def callout(self, node):
        tag = ""
        body_node = Node("span")
        for child in node.children:
            if isinstance(child, Node) and "tag" in child.classes:
                tag = self.inline(child)
            else:
                body_node.children.append(child)
        body = self.inline(body_node)
        if tag and body:
            return f"**{tag}.** {body}"
        return body or tag

    def table(self, node):
        rows = []
        notes = []
        for tr in walk(node):
            if tr.tag != "tr":
                continue
            cells = []
            spanned = False
            for cell in tr.children:
                if isinstance(cell, Node) and cell.tag in ("td", "th"):
                    cells.append(self.inline(cell).replace("|", "\\|"))
                    try:
                        if int(cell.attrs.get("colspan", "1") or 1) > 1:
                            spanned = True
                    except ValueError:
                        pass
            if not cells:
                continue
            if spanned and len(cells) == 2:
                notes.append(f"**{cells[0]}.** {cells[1]}")
                continue
            if spanned:
                note = " ".join(c for c in cells if c).strip()
                if note:
                    notes.append(note)
                continue
            rows.append(cells)
        if not rows:
            return "\n\n".join(notes)
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        lines = ["| " + " | ".join(rows[0]) + " |",
                 "| " + " | ".join(["---"] * width) + " |"]
        lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        if notes:
            return "\n\n".join(["\n".join(lines)] + notes)
        return "\n".join(lines)

    def listing(self, node, indent):
        out = []
        for child in node.children:
            if not isinstance(child, Node) or child.tag != "li":
                continue
            own = Node("span")
            nested = []
            for sub in child.children:
                if isinstance(sub, Node) and sub.tag in ("ul", "ol"):
                    nested.append(sub)
                else:
                    own.children.append(sub)
            line = self.inline(own)
            if line:
                out.append(" " * indent + f"- {line}")
            for sub in nested:
                out.append(self.listing(sub, indent + 2))
        return "\n".join(o for o in out if o)

    def deflist(self, node):
        out = []
        term = ""
        for child in node.children:
            if not isinstance(child, Node):
                continue
            if child.tag == "dt":
                term = self.inline(child)
            elif child.tag == "dd":
                desc = self.inline(child)
                if term and desc:
                    out.append(f"- **{term}**: {desc}")
                elif desc:
                    out.append(f"- {desc}")
        return "\n".join(out)

    def step(self, node):
        num = ""
        body = ""
        for child in node.children:
            if isinstance(child, Node) and "stepchip" in child.classes:
                num = self.text(child)
            elif isinstance(child, Node) and child.tag == "p":
                body = self.inline(child)
        return f"{num}. {body}".strip()

    def term_cell(self, node):
        term = ""
        desc = ""
        for child in node.children:
            if isinstance(child, Node) and child.tag == "h5":
                term = self.inline(child)
            elif isinstance(child, Node) and child.tag == "p":
                desc = self.inline(child)
        if term and desc:
            return f"- **{term}**: {desc}"
        return f"- **{term}**" if term else desc

    def def_cell(self, node):
        title = ""
        body = ""
        for child in node.children:
            if isinstance(child, Node) and child.tag in ("h5", "h4"):
                title = self.inline(child)
            elif isinstance(child, Node) and child.tag == "p":
                body = self.inline(child)
        if title and body:
            return f"**{title}.** {body}"
        return title or body

    def toc(self, pages):
        rows = []
        for node in pages:
            for child in node.children:
                if isinstance(child, Node) and "tocrow" in child.classes:
                    num = name = desc = page_no = ""
                    for sub in child.children:
                        if not isinstance(sub, Node):
                            continue
                        if "tno" in sub.classes:
                            num = self.text(sub)
                        elif "tname" in sub.classes:
                            name = self.text(sub)
                        elif "sub" in sub.classes:
                            desc = self.text(sub)
                        elif "pg" in sub.classes:
                            page_no = self.text(sub)
                    rows.append((num, name, desc, page_no))
        if not rows:
            return ""
        lines = ["| No. | Part | Content | Page |", "| --- | --- | --- | --- |"]
        lines += [f"| {n} | {name} | {d} | {p} |" for n, name, d, p in rows]
        return "\n".join(lines)

    def footer_id(self, node):
        for child in walk(node):
            if "footer" in child.classes or child.tag == "footer":
                leaves = [c for c in child.children if isinstance(c, Node) and c.tag == "span"]
                values = [self.text(s) for s in leaves] or [self.text(child)]
                for value in values:
                    m = re.match(r"^(\d+\.\d+)\s*·", value)
                    if m:
                        return m.group(1)
                    m = re.match(r"^[A-Z]+\s*·\s*(\d+\.\d+)", value)
                    if m:
                        return m.group(1)
                    if value.upper().startswith("APP"):
                        return "APP"
        return ""

    def title_of(self, node):
        for child in walk(node):
            if "footer" in child.classes or child.tag == "footer":
                continue
            if child.tag in ("h1", "h2", "h3"):
                return self.inline(child), child
        for child in walk(node):
            if child.classes & TITLE_CLASSES:
                return self.inline(child), child
        return "", None

    def kick_of(self, node):
        for child in walk(node):
            if child.classes & KICK_CLASSES:
                return self.text(child), child
        return "", None

    def takeaway_of(self, node):
        for child in walk(node):
            if child.classes & TAKEAWAY_CLASSES:
                return self.inline(child), child
        return "", None

    def kind_of(self, node, index):
        if index == 0:
            return "cover"
        cls = " ".join(node.classes).lower()
        if any(k in cls for k in PART_CLASSES):
            return "part"
        return "section"

    # ---------- 单页 ----------
    def slide(self, node, index):
        kind = self.kind_of(node, index)
        fid = self.footer_id(node)
        kick, kick_node = self.kick_of(node)
        title, title_node = self.title_of(node)
        takeaway, takeaway_node = self.takeaway_of(node)
        self.skip_nodes = {id(n) for n in (kick_node, title_node, takeaway_node)
                           if n is not None}

        if kind == "cover":
            meta = ""
            for child in walk(node):
                if "darkmeta" in child.classes:
                    meta = " · ".join(self.text(s) for s in child.children
                                      if isinstance(s, Node) and s.tag == "span")
                    break
            parts = [f"# {title}" if title else "", kick, takeaway, meta]
            body = self.body_blocks(node)
            parts.extend(body)
            return "\n\n".join(p for p in parts if p)

        if kind == "part":
            label = kick.strip() if kick else ""
            label = re.sub(r"^PART\s*", "Part ", label, flags=re.I)
            head = f"## {label} · {title}".strip(" ·") if label else (f"## {title}" if title else "")
            body = self.body_blocks(node)
            return "\n\n".join(p for p in [head, takeaway] + body if p)

        if "st" in node.classes:
            label = titlecase(kick) if kick else ""
            head = f"### {fid} · {label}" if fid and label else ""
            if not head:
                head = f"### {title}" if title else ""
            body = self.body_blocks(node)
            return "\n\n".join(p for p in [head, takeaway] + body if p)

        if "CONTENTS" in kick.upper():
            return "\n\n".join(p for p in ["## Contents", takeaway, self.toc([node])] if p)

        head = title
        if kick:
            pieces = [p.strip() for p in kick.split("·")]
            tag = " · ".join(titlecase(p) for p in pieces[2:] if p)
            num = fid if fid and not fid.startswith("APP") else ""
            head = " · ".join(x for x in (num, tag, title) if x)
        head = f"### {head}" if head else ""
        body = self.body_blocks(node)
        return "\n\n".join(p for p in [head, takeaway] + body if p)

    def body_blocks(self, node):
        result = []
        for child in node.children:
            if not isinstance(child, Node):
                continue
            cls = child.classes
            if "footer" in cls or child.tag == "footer" or child.tag in SKIP_TAGS:
                continue
            if cls & {"kick", "stkick"}:
                continue
            if child.classes & TITLE_CLASSES:
                continue
            if child.classes & TAKEAWAY_CLASSES:
                continue
            if "darkmeta" in cls or "sttext" in cls:
                continue
            if "tframe" in cls or "stframe" in cls:
                continue
            chunk = self.block(child)
            if chunk:
                result.append(chunk)
        return result

    def run(self):
        parser = TreeBuilder()
        parser.feed(self.html_path.read_text(encoding="utf-8"))
        pages = self.collect_pages(parser.root)
        chunks = []
        for i, node in enumerate(pages):
            self.base_level = 2 if self.kind_of(node, i) == "part" else 3
            chunk = self.slide(node, i)
            if chunk:
                chunks.append(chunk)
        note = "\n\n---\n\n".join(chunks) + "\n"
        self.out_path.parent.mkdir(parents=True, exist_ok=True)
        self.out_path.write_text(note, encoding="utf-8")
        return len(pages), len(note)

    @staticmethod
    def collect_pages(root):
        candidates = [n for n in walk(root)
                      if n.tag in ("section", "div", "article")
                      and "slide" in n.classes]
        pages = []
        for node in candidates:
            if any(a in pages for a in ancestors(node)):
                continue
            pages.append(node)
        if len(pages) >= 2:
            return pages
        sections = [n for n in walk(root) if n.tag == "section"]
        outer = [n for n in sections if not any(a.tag == "section" for a in ancestors(n))]
        if len(outer) >= 3:
            return outer
        body = next((n for n in walk(root) if n.tag == "body"), root)
        return [body]


def ancestors(node):
    current = node.parent
    while current is not None:
        yield current
        current = current.parent


def detect_prefix(path):
    stem = path.stem
    m = re.match(r"^([A-Za-z]{1,3}\d+)", stem)
    return m.group(1) if m else "deck"


def output_for(src, outdir):
    return (outdir / f"{src.stem}-整理.md") if outdir else src.with_name(f"{src.stem}-整理.md")


def process(src, args):
    out = Path(args.out) if args.out else output_for(src, Path(args.outdir) if args.outdir else None)
    if args.skip_existing and out.exists():
        print(f"[跳过] {out.name} 已存在")
        return
    prefix = args.prefix or detect_prefix(src)
    assets = Path(args.assets) if args.assets else out.parent / "assets"
    renderer = Renderer(src, out, assets, prefix)
    pages, chars = renderer.run()
    imgs = sum(1 for s in renderer.saved if s.lower().endswith((".jpg", ".png", ".gif", ".webp")))
    svgs = sum(1 for s in renderer.saved if s.endswith(".svg"))
    extra = f"，缺失图片 {len(renderer.missing)}" if renderer.missing else ""
    print(f"[完成] {src.name} → {out.name}（{pages} 页，{chars} 字符，图片 {imgs} / SVG {svgs}{extra}）")


def main():
    ap = argparse.ArgumentParser(description="通用 HTML 幻灯片清洗：结构还原 + 资源落盘")
    ap.add_argument("paths", nargs="+", help="一个或多个 .html 文件，或目录")
    ap.add_argument("--out", help="指定单个输出文件（仅单输入时可用）")
    ap.add_argument("--outdir", help="输出目录，默认与源文件同目录")
    ap.add_argument("--assets", help="图片与 SVG 的根目录，默认输出目录下的 assets")
    ap.add_argument("--prefix", help="资源命名前缀，默认取文件名编号（如 L2）")
    ap.add_argument("--skip-existing", action="store_true", help="输出已存在时跳过")
    args = ap.parse_args()

    files = []
    for raw in args.paths:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(p for p in path.glob("*.html") if p.is_file()))
        elif path.is_file():
            files.append(path)
        else:
            sys.exit(f"找不到文件或目录：{path}")
    if not files:
        sys.exit("没有可处理的 .html 文件")
    if args.out and len(files) > 1:
        sys.exit("--out 只能用于单个输入文件")

    for src in files:
        process(src, args)


if __name__ == "__main__":
    main()
