#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Markdown 批量导出 PDF 脚本（ordinote-export 功能）

功能：
    把 markdown 笔记批量导出为 PDF，版式复用本机 Markdown Preview Enhanced
    （MPE）扩展自带的样式与图表资源，观感与 MPE 预览一致。
    不需要下载任何东西：解析用已安装的 markdown-it-py，主题样式与 mermaid
    取自 MPE 扩展目录，公式渲染引擎取自本机已装的其它扩展，打印用本机
    Chrome / Edge 的无头模式。

    公式（$...$ / $$...$$）在 markdown 解析之前抽取，交页面里的 KaTeX 渲染，
    因此 LaTeX 内容不会被 markdown 的转义、强调与段落规则改写。
    tikz 代码块不渲染：MPE 的 tikz 由扩展宿主内置的 wasm TeX 引擎生成，
    该引擎无法被外部脚本调用，导出时 tikz 以代码文本呈现。

    源笔记只读，脚本不修改、不移动任何笔记文件。

用法：
    python md2pdf.py <文件或目录路径> [更多路径...] [选项]

选项：
    --out DIR         输出目录（默认与源文件同目录）
    --theme NAME      预览主题，默认 github-light（对应 MPE 的 preview_theme）
    --prism NAME      代码块配色主题，默认 github（对应 MPE 的 prism_theme）
    --paper SIZE      纸张尺寸，默认 A4
    --margin VALUE    页边距，默认 "1.4cm 1.3cm"
    --wait MS         页面渲染等待时间（毫秒），默认 3000
    --chrome PATH     指定浏览器可执行文件
    --mpe DIR         指定 MPE 扩展目录或其 crossnote 子目录
    --skip-existing   已有同名 PDF 时跳过
    --exclude REGEX   按路径正则排除文件

输出：
    与源文件同目录同名 .pdf；指定 --out 时输出到该目录

退出码：
    0  全部成功（含跳过）
    1  用法错误 / 缺少必需资源（MPE、浏览器）
    2  存在导出失败的文件
"""

import argparse
import base64
import html as html_lib
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    from markdown_it import MarkdownIt
except ImportError:  # pragma: no cover
    sys.stderr.write(
        "[错误] 缺少 markdown-it-py。该库通常随其它 Python 包一同安装；\n"
        "       若确实缺失，可执行：pip install markdown-it-py\n"
    )
    sys.exit(1)


# ------------------------------------------------------------
# 常量：资源候选位置
# ------------------------------------------------------------

# 各类 VS Code 系客户端的扩展目录
EXTENSION_ROOTS = [
    Path.home() / ".trae-cn" / "extensions",
    Path.home() / ".trae" / "extensions",
    Path.home() / ".vscode" / "extensions",
    Path.home() / ".vscode-insiders" / "extensions",
    Path.home() / ".cursor" / "extensions",
    Path.home() / ".windsurf" / "extensions",
]

# MPE 扩展内的资源目录，以及缺一不可的文件
MPE_DIR_PATTERN = "shd101wyy.markdown-preview-enhanced-*/crossnote"
MPE_REQUIRED = [
    "styles/style-template.css",
    "styles/preview_theme",
    "dependencies/mermaid/mermaid.min.js",
]

# 公式渲染引擎（KaTeX）的候选位置与必需文件
KATEX_PATTERNS = [
    "*/node_modules/@marp-team/marp-core/node_modules/katex/dist",
    "*/node_modules/katex/dist",
]
KATEX_REQUIRED = ["katex.min.js", "katex.min.css", "fonts"]

# 浏览器可执行文件候选位置
BROWSER_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
    r"C:\Program Files\Chromium\Application\chrome.exe",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]

DEFAULT_THEME = "github-light"
DEFAULT_PRISM = "github"
DEFAULT_PAPER = "A4"
DEFAULT_MARGIN = "1.4cm 1.3cm"
DEFAULT_WAIT_MS = 3000

# 追加在 MPE 样式之后的打印规则。MPE 自身不写 @page（页边距由它传给
# puppeteer 的参数控制），纸张与边距需要在导出的 HTML 里补齐。
PRINT_CSS_TEMPLATE = """
@page {{ size: {paper}; margin: {margin}; }}
html, body {{ background: #fff !important; }}
html body {{
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}}
.markdown-preview {{ height: auto !important; padding: 0 !important; }}
h1, h2, h3, h4, h5, h6 {{ break-after: avoid; page-break-after: avoid; }}
table, pre, blockquote, img, .mermaid {{
  break-inside: avoid;
  page-break-inside: avoid;
}}
tr, li {{ break-inside: avoid; page-break-inside: avoid; }}
.mermaid {{ text-align: center; }}
"""

HTML_TEMPLATE = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
<base href="{base}">
<style>
{css}
</style>
</head>
<body for="html-export">
<div class="markdown-preview" data-for="preview">
{body}
</div>
{scripts}
</body>
</html>
"""

MATH_SCRIPT_TEMPLATE = """<script src="{katex_js}"></script>
<script>
document.querySelectorAll("span.math").forEach(function (el) {{
  var tex = el.getAttribute("data-tex");
  try {{
    katex.render(tex, el, {{
      displayMode: el.classList.contains("math-display"),
      throwOnError: false
    }});
  }} catch (err) {{
    el.textContent = tex;
  }}
}});
</script>"""

MERMAID_SCRIPT_TEMPLATE = """<script src="{mermaid_js}"></script>
<script>mermaid.initialize({{ startOnLoad: true }});</script>"""


class ResourceMissing(Exception):
    """必需资源缺失时抛出，携带给用户看的说明。"""


# ------------------------------------------------------------
# 资源定位
# ------------------------------------------------------------

def _pick_latest(candidates):
    """多个版本并存时取版本号最大的一个（目录名倒序）。"""
    return sorted(candidates, key=lambda p: p.name)[-1] if candidates else None


def _require(directory, required, label):
    missing = [name for name in required if not (directory / name).exists()]
    if missing:
        raise ResourceMissing(f"{label}不完整，缺少：{'、'.join(missing)}（路径：{directory}）")
    return directory


def find_mpe_crossnote(explicit=None):
    """定位 MPE 扩展自带的资源目录（styles/ 与 dependencies/ 的上级）。"""
    if explicit:
        path = Path(explicit).expanduser()
        if (path / "crossnote").is_dir():
            path = path / "crossnote"
        return _require(path, MPE_REQUIRED, "MPE 资源目录")

    env_path = os.environ.get("MPE_HOME")
    if env_path:
        return find_mpe_crossnote(env_path)

    hits = []
    for root in EXTENSION_ROOTS:
        if root.is_dir():
            hits.extend(root.glob(MPE_DIR_PATTERN))
    latest = _pick_latest(hits)
    if latest is None:
        raise ResourceMissing(
            "未找到 Markdown Preview Enhanced 扩展。请在本机的类 VS Code 客户端"
            "（Trae / VS Code / Cursor 等）中安装该扩展，或用 --mpe 指定其 crossnote 目录。"
        )
    return _require(latest, MPE_REQUIRED, "MPE 资源目录")


def find_katex_dist(explicit=None):
    """定位 KaTeX 的 dist 目录（含 JS、CSS 与字体）；找不到时返回 None。"""
    if explicit:
        return _require(Path(explicit).expanduser(), KATEX_REQUIRED, "KaTeX 目录")

    for root in EXTENSION_ROOTS:
        if not root.is_dir():
            continue
        for pattern in KATEX_PATTERNS:
            hits = [p for p in root.glob(pattern) if (p / "katex.min.js").is_file()]
            latest = _pick_latest(hits)
            if latest is not None:
                return _require(latest, KATEX_REQUIRED, "KaTeX 目录")
    return None


def find_browser(explicit=None):
    """定位浏览器可执行文件，顺序为参数、环境变量、常见安装路径、PATH。"""
    if explicit:
        path = Path(explicit).expanduser()
        if path.is_file():
            return path
        raise ResourceMissing(f"指定的浏览器不存在：{path}")

    env_path = os.environ.get("CHROME_PATH")
    if env_path and Path(env_path).is_file():
        return Path(env_path)

    for candidate in BROWSER_CANDIDATES:
        if Path(candidate).is_file():
            return Path(candidate)

    for name in ("chrome", "google-chrome", "chromium", "msedge"):
        found = shutil.which(name)
        if found:
            return Path(found)

    raise ResourceMissing(
        "未找到 Chrome / Edge / Chromium。请安装其中之一，或用 --chrome 指定可执行文件路径。"
    )


# ------------------------------------------------------------
# 转换
# ------------------------------------------------------------

def read_text(path):
    # utf-8-sig：兼容带 BOM 的文件（PowerShell 写出的文本常带 BOM）
    return path.read_text(encoding="utf-8-sig")


def strip_front_matter(text):
    """剥离文件开头的 YAML front-matter，避免出现在 PDF 正文中。"""
    if text.startswith("---"):
        match = re.match(r"^---\r?\n.*?\r?\n---\r?\n", text, flags=re.S)
        if match:
            return text[match.end():]
    return text


def inline_fonts(css, base_dir):
    """把 KaTeX 的 woff2 字体以 data URI 内嵌。

    file:// 页面加载跨目录字体时会因 CORS 被拦截，内嵌可绕开该限制，
    同时只保留 woff2 一种格式（Chrome / Edge 均支持）。
    base_dir 为 CSS 文件所在目录，CSS 里的 url(fonts/...) 相对它解析。
    """
    css = re.sub(r',\s*url\([^)]*?\.(?:woff|ttf)\)\s*(?:format\([^)]*\))?', "", css)

    def replace(match):
        font_file = base_dir / match.group(2)
        if not font_file.is_file():
            return match.group(0)
        encoded = base64.b64encode(font_file.read_bytes()).decode("ascii")
        return f'url("data:font/woff2;base64,{encoded}")'

    return re.sub(r'url\(\s*(["\']?)(.*?)\1\s*\)', replace, css)


# 公式扫描。必须在 markdown 解析之前完成：markdown 的转义规则会把 LaTeX 里的
# \\ \_ \{ \} \* 改写成别的字符，块级公式也会被段落切分打断。
# code 分支保证围栏之外的行内代码里的 $ 不被当作公式。
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
    """把公式抽成占位元素，交给页面里的 KaTeX 渲染。

    抽取后 LaTeX 原文只存在于占位元素的属性中，不经过 markdown 的转义、
    强调与段落切分规则；元素内的文本是原文，供 KaTeX 缺失时降级显示。
    """
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
            tex = " ".join(raw[2:-2 if display else -1].split())
            cls = "math math-display" if display else "math"
            return (
                f'<span class="{cls}" data-tex="{html_lib.escape(tex, quote=True)}">'
                f"{html_lib.escape(tex)}</span>"
            )

        pieces.append(MATH_SCAN_RE.sub(replace, segment))
    return "".join(pieces)


def md_to_html_body(md_text, md=None):
    """markdown 转 HTML；公式与 mermaid 代码块单独处理。"""
    md = md or MarkdownIt("commonmark", {"html": True}).enable(["table", "strikethrough"])
    body = md.render(protect_math(md_text))
    body = re.sub(
        r'<pre><code class="language-mermaid">(.*?)</code></pre>',
        r'<div class="mermaid">\1</div>',
        body,
        flags=re.S,
    )
    return body


def _title_of(body, fallback):
    match = re.search(r"<h1[^>]*>(.*?)</h1>", body, flags=re.S)
    if match:
        text = re.sub(r"<[^>]+>", "", match.group(1)).strip()
        if text:
            return text
    return fallback


def build_html(body, styles, base_dir, title, browser_assets, paper, margin):
    """组装完整页面：MPE 骨架 + 内联样式 + 本地脚本。"""
    css_parts = list(styles)
    css_parts.append(PRINT_CSS_TEMPLATE.format(paper=paper, margin=margin))

    scripts = ""
    if "<div class=\"mermaid\">" in body:
        scripts += MERMAID_SCRIPT_TEMPLATE.format(mermaid_js=browser_assets["mermaid"].as_uri())
    if '<span class="math' in body and browser_assets.get("katex"):
        scripts += MATH_SCRIPT_TEMPLATE.format(
            katex_js=(browser_assets["katex"] / "katex.min.js").as_uri(),
        )

    return HTML_TEMPLATE.format(
        title=html_lib.escape(title),
        base=base_dir.as_uri().rstrip("/") + "/",
        css="\n".join(css_parts),
        body=body,
        scripts=scripts,
    )


# 浏览器 stderr 中与导出无关的噪声（首次运行时的组件安装、输入法日志等）
BROWSER_NOISE = (
    "externally_managed_app_manager",
    "external_registry_loader",
    "install source",
    "SogouPY",
)


def _browser_error(result):
    """从浏览器输出中挑出有用的错误行，跳过无关噪声。"""
    lines = [line.strip() for line in (result.stderr or "").splitlines() if line.strip()]
    meaningful = [line for line in lines if not any(noise in line for noise in BROWSER_NOISE)]
    return meaningful[-1] if meaningful else (lines[-1] if lines else "无输出")


def render_pdf(html_text, pdf_path, browser, wait_ms):
    """无头浏览器打印。临时 HTML 与浏览器 profile 都放在临时目录，用后即删。"""
    with tempfile.TemporaryDirectory(prefix="ordinote-pdf-") as tmp:
        tmp_dir = Path(tmp)
        html_path = tmp_dir / "page.html"
        html_path.write_text(html_text, encoding="utf-8")

        common = [
            "--disable-gpu",
            "--no-pdf-header-footer",
            "--no-first-run",
            "--no-default-browser-check",
            f"--virtual-time-budget={wait_ms}",
            f"--user-data-dir={tmp_dir / 'profile'}",
            f"--print-to-pdf={pdf_path}",
            html_path.as_uri(),
        ]
        timeout = wait_ms / 1000 + 90 if wait_ms else 120
        errors = []
        for mode in ("--headless=new", "--headless"):
            if pdf_path.exists():
                break
            result = subprocess.run(
                [str(browser), mode, *common],
                capture_output=True, text=True, errors="replace", timeout=timeout,
            )
            if not pdf_path.exists():
                errors.append(f"{mode}：退出码 {result.returncode}，{_browser_error(result)}")

        if not pdf_path.exists():
            raise RuntimeError("浏览器未生成 PDF（" + "；".join(errors) + "）")

        size = pdf_path.stat().st_size
        if size == 0:
            raise RuntimeError("生成的 PDF 为空文件")
    return size


# ------------------------------------------------------------
# 批处理
# ------------------------------------------------------------

def collect_targets(paths, exclude_re, out_dir):
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


def process_one(md_path, opts, styles, mpe_dir, katex_dir, browser):
    """单个文件的完整闭环，返回 (状态, 说明)。"""
    out_dir = Path(opts.out).expanduser() if opts.out else md_path.parent
    # 必须交给浏览器绝对路径：相对路径会按浏览器自身的工作目录解析，导致写入失败
    pdf_path = (out_dir / (md_path.stem + ".pdf")).absolute()

    if opts.skip_existing and pdf_path.exists():
        return "skip", f"已存在同名 PDF：{pdf_path}"

    out_dir.mkdir(parents=True, exist_ok=True)
    body = md_to_html_body(strip_front_matter(read_text(md_path)))
    page = build_html(
        body=body,
        styles=styles,
        base_dir=md_path.parent.resolve(),
        title=_title_of(body, md_path.stem),
        browser_assets={"mermaid": mpe_dir / "dependencies" / "mermaid" / "mermaid.min.js",
                        "katex": katex_dir},
        paper=opts.paper,
        margin=opts.margin,
    )
    size = render_pdf(page, pdf_path, browser, opts.wait)
    return "ok", f"{pdf_path}（{size / 1024:.0f} KB）"


def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="把 markdown 批量导出为 PDF，版式对齐 Markdown Preview Enhanced 预览",
    )
    parser.add_argument("paths", nargs="+", help="markdown 文件或目录，可多个")
    parser.add_argument("--out", help="输出目录，默认与源文件同目录")
    parser.add_argument("--theme", default=DEFAULT_THEME, help=f"预览主题，默认 {DEFAULT_THEME}")
    parser.add_argument("--prism", default=DEFAULT_PRISM, help=f"代码块配色，默认 {DEFAULT_PRISM}")
    parser.add_argument("--paper", default=DEFAULT_PAPER, help=f"纸张尺寸，默认 {DEFAULT_PAPER}")
    parser.add_argument("--margin", default=DEFAULT_MARGIN, help=f"页边距，默认 {DEFAULT_MARGIN}")
    parser.add_argument("--wait", type=int, default=DEFAULT_WAIT_MS, help="渲染等待毫秒数")
    parser.add_argument("--chrome", help="浏览器可执行文件路径")
    parser.add_argument("--mpe", help="MPE 扩展目录或其 crossnote 子目录")
    parser.add_argument("--skip-existing", action="store_true", help="已有同名 PDF 时跳过")
    parser.add_argument("--exclude", help="按路径正则排除文件")
    return parser.parse_args(argv)


def load_styles(mpe_dir, theme, prism, katex_dir):
    """读取样式：MPE 版式基础 + 预览主题 + 代码块配色 + KaTeX（字体内嵌）。"""
    styles, notes = [], []
    styles.append(read_text(mpe_dir / "styles" / "style-template.css"))

    theme_file = mpe_dir / "styles" / "preview_theme" / f"{theme}.css"
    if theme_file.is_file():
        styles.append(read_text(theme_file))
    else:
        notes.append(f"未找到主题 {theme}.css，仅使用基础样式")

    prism_file = mpe_dir / "styles" / "prism_theme" / f"{prism}.css"
    if prism_file.is_file():
        styles.append(read_text(prism_file))

    if katex_dir is not None:
        styles.append(inline_fonts(read_text(katex_dir / "katex.min.css"), katex_dir))
    return styles, notes


def main(argv=None):
    opts = parse_args(argv)

    try:
        mpe_dir = find_mpe_crossnote(opts.mpe)
        browser = find_browser(opts.chrome)
    except ResourceMissing as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1

    katex_dir = find_katex_dist()
    styles, notes = load_styles(mpe_dir, opts.theme, opts.prism, katex_dir)
    if katex_dir is None:
        notes.append("未找到 KaTeX 渲染引擎，公式将以原始 LaTeX 文本呈现")
    if notes:
        for note in notes:
            print(f"[提示] {note}")

    try:
        targets = collect_targets(opts.paths, re.compile(opts.exclude) if opts.exclude else None,
                                  opts.out)
    except (FileNotFoundError, re.error) as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 1

    if not targets:
        print("[提示] 没有匹配的 markdown 文件")
        return 0

    counts = {"ok": 0, "skip": 0, "fail": 0}
    failures = []
    for md_path in targets:
        try:
            status, message = process_one(md_path, opts, styles, mpe_dir, katex_dir, browser)
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