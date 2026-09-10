#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地 PDF/PPT 文字提取脚本（ordinote-input 功能 · 阶段一）

功能：
    从 PDF / PPTX 文件中提取全部文字，生成 markdown 提取文件。
    提取文件是"原始素材"，后续按 references/ordinote-input.md 的
    阶段二规则做轻量整理（修错别字、去空格换行等）。

用法：
    python extract_local.py <文件路径> [更多文件路径...]

输出：
    在源文件同目录下生成 {原文件名}-提取.md
    - 各页 / 各张幻灯片之间以 --- 分隔线分开（保留源结构线索）
    - PPT 的演讲者备注以 "> 备注：" 形式附在对应幻灯片之后
    - 空白页 / 无文字幻灯片以 HTML 注释占位，便于核对缺页
    - 同名旧提取文件会被直接覆盖

退出码：
    0  全部提取成功
    1  用法错误 / 文件不存在 / 不支持的文件类型
    2  提取失败或结果为空（典型：扫描版 PDF）
    3  缺少依赖库（PyMuPDF / python-pptx），且自动安装被拒绝或失败
"""

import sys
import pathlib
import subprocess
from datetime import datetime

# ------------------------------------------------------------
# 常量
# ------------------------------------------------------------

# 各页（幻灯片）之间的分隔线；阶段二整理时会保留它作为结构线索
PAGE_SEPARATOR = "\n\n---\n\n"

# 文件头注释：记录来源信息。HTML 注释不会被渲染，仅供追溯
HEADER_TEMPLATE = (
    "<!-- 本文件由 extract_local.py 自动提取生成 -->\n"
    "<!-- 源文件：{name} | {kind}数：{count} | 提取时间：{time} -->\n"
)


class ExtractError(Exception):
    """单个文件提取失败时抛出，携带退出码与给用户看的说明。

    多文件批量提取时，单个文件失败只记录错误、不中断其他文件。
    """

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


# ------------------------------------------------------------
# 依赖自动安装
# ------------------------------------------------------------

# 依赖表：(可导入模块名列表, pip 包名, 用途说明)
# pymupdf 旧版只注册 fitz 模块名，故两个名字都尝试
DEPENDENCIES = [
    (["pymupdf", "fitz"], "PyMuPDF", "PDF 文字提取"),
    (["pptx"], "python-pptx", "PPT 文字提取"),
]

# 是否允许缺依赖时自动安装（--no-auto-install 可关闭）
AUTO_INSTALL = True


def _confirm(prompt):
    """交互确认，直接回车默认为"是"；非交互终端（管道/无键盘）自动选"是"。"""
    if not sys.stdin or not sys.stdin.isatty():
        print(prompt + "（非交互终端，自动选择 Y）")
        return True
    try:
        return input(prompt).strip().lower() not in ("n", "no")
    except (EOFError, KeyboardInterrupt):
        print()
        return False


def _pip_install(pip_name):
    """用当前解释器调用 pip 安装，失败时抛 ExtractError。

    用 sys.executable 保证装进"正在运行本脚本的 Python"，
    避免电脑上有多个 Python 时装错环境的常见坑。
    """
    cmd = [sys.executable, "-m", "pip", "install", pip_name]
    print(f"[安装] 执行：{' '.join(cmd)}")
    try:
        result = subprocess.run(cmd)
    except OSError as e:
        raise ExtractError(3, f"调用 pip 失败：{e}；请手动执行：pip install {pip_name}")
    if result.returncode != 0:
        raise ExtractError(
            3, f"{pip_name} 安装失败（pip 退出码 {result.returncode}），"
               f"请检查网络/权限后手动执行：pip install {pip_name}")


def ensure_module(import_names, pip_name, purpose, ask=True):
    """导入第三方模块；未安装时（经确认后）自动 pip 安装再重试。

    已安装 → 直接返回模块，不调用 pip、不产生任何额外输出。
    """
    import importlib

    def _try_import():
        for name in import_names:
            try:
                return importlib.import_module(name)
            except ImportError:
                continue
        return None

    mod = _try_import()
    if mod:
        return mod

    if not AUTO_INSTALL and ask:
        raise ExtractError(3, f"缺少依赖库 {pip_name}（用于{purpose}），"
                              f"请先安装：pip install {pip_name}")
    if ask:
        print(f"[依赖] 未检测到 {pip_name}（用于{purpose}）。")
        if not _confirm(f"       是否现在自动安装？[Y/n] "):
            raise ExtractError(3, f"已取消安装 {pip_name}，"
                                  f"需要时请手动执行：pip install {pip_name}")

    _pip_install(pip_name)
    importlib.invalidate_caches()

    mod = _try_import()
    if mod is None:
        raise ExtractError(3, f"{pip_name} 已安装但仍无法导入，"
                              f"可能装到了其他 Python 环境（当前解释器：{sys.executable}）")
    print(f"[安装] {pip_name} 安装成功。")
    return mod


# ------------------------------------------------------------
# PDF 提取（依赖：PyMuPDF）
# ------------------------------------------------------------

def extract_pdf(path):
    """逐页提取 PDF 文字，返回"每页文字"组成的列表（空页为空字符串）。"""
    # 优先新模块名 pymupdf；旧版只有 fitz，ensure_module 会依次尝试
    fitz = ensure_module(["pymupdf", "fitz"], "PyMuPDF", "PDF 文字提取")

    doc = fitz.open(path)
    try:
        blocks = []
        for page in doc:
            # "text" 模式按阅读顺序输出文字，是大多数场景的最佳选择
            blocks.append(page.get_text("text").strip())
    finally:
        doc.close()

    # 整本都没有文字 → 大概率是扫描版 / 图片型 PDF
    if not any(blocks):
        raise ExtractError(
            2, "PDF 所有页面均未提取到文字，可能是扫描版/图片型 PDF，"
               "请改用 MinerU API（scripts/mineru_api.py）处理"
        )
    return blocks


# ------------------------------------------------------------
# PPT 提取（依赖：python-pptx）
# ------------------------------------------------------------

def walk_shapes(shapes):
    """深度优先遍历幻灯片上的形状，逐个产出其中的文字。

    - 组合形状：递归进入其子形状
    - 表格：每行拼成制表符分隔的文本（尽量保留表格结构）
    - 普通文本框：直接取文字内容
    """
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk_shapes(shape.shapes)  # 组合形状 → 递归子形状
        elif getattr(shape, "has_table", False):  # 仅表格形状有该属性
            rows = ("\t".join(cell.text.strip() for cell in row.cells)
                    for row in shape.table.rows)
            yield "\n".join(rows)
        elif shape.has_text_frame:
            yield shape.text_frame.text


def extract_pptx(path):
    """逐张幻灯片提取文字（含演讲者备注），返回"每张文字"组成的列表。"""
    pptx = ensure_module(["pptx"], "python-pptx", "PPT 文字提取")
    Presentation = pptx.Presentation

    prs = Presentation(str(path))  # python-pptx 要求字符串路径
    blocks = []

    for slide in prs.slides:
        # 收集这张幻灯片上所有形状的文字（已去除空段落）
        parts = [t.strip() for t in walk_shapes(slide.shapes) if t.strip()]

        # 演讲者备注里常有口头补充的内容，一并保留
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                parts.append(f"> 备注：{notes}")

        blocks.append("\n\n".join(parts))

    if not any(blocks):
        raise ExtractError(2, "PPT 所有幻灯片均未提取到文字，文件可能损坏或为纯图片版式")
    return blocks


# ------------------------------------------------------------
# 拼装与写出
# ------------------------------------------------------------

def assemble(path, kind, blocks):
    """把逐页文字列表拼成最终的 markdown 文本。

    kind 为"页"或"幻灯片"，用于头部统计和空页占位注释的措辞。
    """
    header = HEADER_TEMPLATE.format(
        name=path.name,
        kind=kind,
        count=len(blocks),
        time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    # 空页用注释占位：不影响渲染，又能核对哪一页缺了内容
    filled = [
        text if text.strip() else f"<!-- 第 {i} {kind}未提取到文字 -->"
        for i, text in enumerate(blocks, start=1)
    ]
    return header + "\n" + PAGE_SEPARATOR.join(filled) + "\n"


def process_one(arg):
    """处理单个文件：判定类型 → 提取 → 写出 {原文件名}-提取.md。"""
    path = pathlib.Path(arg)
    if not path.is_file():
        raise ExtractError(1, f"文件不存在：{path}")

    suffix = path.suffix.lower()
    if suffix == ".pdf":
        blocks, kind = extract_pdf(path), "页"
    elif suffix == ".pptx":
        blocks, kind = extract_pptx(path), "幻灯片"
    elif suffix == ".ppt":
        raise ExtractError(1, f"不支持旧版 .ppt（{path.name}），请先用 PowerPoint 另存为 .pptx")
    else:
        raise ExtractError(1, f"不支持的文件类型：{path.name}（仅支持 .pdf / .pptx）")

    content = assemble(path, kind, blocks)
    out = path.with_name(f"{path.stem}-提取.md")  # 输出到源文件同目录
    out.write_text(content, encoding="utf-8")

    # 汇报统计信息，供调用方（模型/用户）核对提取质量
    chars = sum(len(t) for t in blocks)
    print(f"[成功] {path.name} → {out.name}（{len(blocks)} {kind}，共 {chars} 字）")
    empty = [i for i, t in enumerate(blocks, 1) if not t.strip()]
    if empty:
        print(f"[提醒] 以下{kind}未提取到文字，可能为图片页："
              f"{'、'.join(map(str, empty))}")


# ------------------------------------------------------------
# 主入口
# ------------------------------------------------------------

def install_all_deps():
    """--install-deps：把全部依赖一次装齐（已安装的自动跳过），供交接前预热。"""
    for import_names, pip_name, purpose in DEPENDENCIES:
        ensure_module(import_names, pip_name, purpose)
    print("[完成] 全部依赖已就绪。")


def main():
    global AUTO_INSTALL

    # Windows 控制台默认可能是 GBK，强制 UTF-8 避免中文输出乱码/报错
    for stream in (sys.stdout, sys.stderr):
        if stream and hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")

    args = sys.argv[1:]

    # 解析开关参数（与文件路径可混写）
    if "--no-auto-install" in args:
        AUTO_INSTALL = False
        args = [a for a in args if a != "--no-auto-install"]
    if "--install-deps" in args:
        try:
            install_all_deps()
        except ExtractError as e:
            print(f"[失败] {e.message}")
            sys.exit(e.code)
        sys.exit(0)

    if not args:
        print(__doc__)  # 打印用法说明
        print("[失败] 请提供至少一个 PDF/PPTX 文件路径")
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
