#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MinerU API 文字提取脚本（ordinote-input 功能 · 阶段一 · 云端备选）

功能：
    调用 MinerU 精准解析 API（/api/v4），从 PDF/图片/DOC/PPT 等文件中提取
    文字，生成 markdown 提取文件。适用于扫描件、复杂版式——即本地脚本
    extract_local.py 提取失败或效果较差时的备选方案。

    支持的文件类型：pdf / doc / docx / ppt / pptx / 图片（png / jpg /
    jpeg / jp2 / webp / gif / bmp）。支持直接上传图片提取文字（按其内容
    做 OCR；OCR 默认开启，可用 --no-ocr 关闭）。

    支持一次传入多个文件批量提取：多个文件路径（PDF、图片可混用）会自动
    分批提交（每批 ≤ 50 个，官方单批上限 200 个），每个文件各自生成
    {原文件名}-提取.md。

    提取文件是"原始素材"，后续按 references/ordinote-input.md 的阶段二
    规则做轻量整理（修错别字、去空格换行等）。

用法：
    python mineru_api.py <文件路径> [更多文件路径...] [选项]
    （文件路径可为 PDF/PPT/DOC/图片；可一次传多个文件批量提取）

选项：
    --token <token>      MinerU Token（优先级最高）
    --token-file <路径>  从本地文件读取 Token；不传则按顺序查找默认文件
    --model <name>       模型版本：pipeline / vlm（默认，推荐）/ MinerU-HTML
    --pages <范围>       只解析指定页码，如 "1-50"、"2,4-6"（超长/超页文档分批用）
    --no-ocr             关闭 OCR（默认开启；扫描件建议保持开启）
    --timeout <秒>       等待解析结果的超时时间，默认 900 秒
    --no-auto-install    缺少依赖库时不自动安装
    --install-deps       只安装依赖后退出（供交接前预热）

Token 配置（优先级从高到低，任选其一即可）：
    1) 命令行参数  --token <token>
    2) 本地文件    --token-file <路径>
    3) 环境变量    MINERU_API_TOKEN
    4) 默认文件    scripts/mineru_token.txt  或  ~/.mineru_token
       文件为纯 Token 一行即可；也支持 .env 风格 "MINERU_API_TOKEN=xxx"
    Token 需在 https://mineru.net/apiManage/docs 的"API 管理"页自行创建

输出：
    在源文件同目录下生成 {原文件名}-提取.md
    - 内容取自 MinerU 结果压缩包中的 full.md（已经是干净 Markdown）
    - 与 extract_local.py 不同：MinerU 不返回逐页文字，故正文中无 --- 分页线

退出码：
    0  全部提取成功
    1  用法错误 / 文件不存在 / 不支持的文件类型
    2  提取失败或结果为空（解析任务 failed、压缩包内无 markdown 等）
    3  缺少依赖库（requests）或未配置 Token
    4  网络/接口错误（上传失败、任务超时、接口返回异常等）
"""

import http.client
import io
import json
import os
import pathlib
import subprocess
import sys
import time
import uuid
import zipfile
from datetime import datetime
from urllib.parse import urljoin, urlsplit

# ------------------------------------------------------------
# 常量
# ------------------------------------------------------------

# MinerU 接口地址（精准解析 API）
BASE_URL = "https://mineru.net"
BATCH_UPLOAD_URL = f"{BASE_URL}/api/v4/file-urls/batch"          # 申请上传链接并创建解析任务
BATCH_RESULT_URL = f"{BASE_URL}/api/v4/extract-results/batch/{{batch_id}}"  # 查询批量结果

# Token 环境变量名
TOKEN_ENV = "MINERU_API_TOKEN"

# 默认 Token 文件：脚本同目录的 mineru_token.txt，以及用户主目录的 ~/.mineru_token
SCRIPT_DIR = pathlib.Path(__file__).resolve().parent
DEFAULT_TOKEN_FILES = [SCRIPT_DIR / "mineru_token.txt", pathlib.Path.home() / ".mineru_token"]

# 模型版本：pipeline（默认）/ vlm（推荐）/ MinerU-HTML（仅 html 文件）
DEFAULT_MODEL = "vlm"
VALID_MODELS = ("pipeline", "vlm", "MinerU-HTML")

# 单次申请链接数上限（官方文档：单次不超过 200 个；这里保守取 50，多批更稳）
BATCH_SIZE = 50

# 单个文件大小上限（官方文档：200 MB；页数上限 600 页，超页可用 --pages 分批）
MAX_FILE_SIZE = 200 * 1024 * 1024

# 轮询间隔与默认超时
POLL_INTERVAL = 5.0
DEFAULT_TIMEOUT = 900.0

# 支持的文件类型（对应 MinerU 可解析的格式；本 Skill 主要用 PDF/PPT）
SUPPORTED_EXTS = {
    ".pdf", ".doc", ".docx", ".ppt", ".pptx",
    ".png", ".jpg", ".jpeg", ".jp2", ".webp", ".gif", ".bmp",
}

# 文件头注释：记录来源信息。HTML 注释不会被渲染，仅供追溯
HEADER_TEMPLATE = (
    "<!-- 本文件由 mineru_api.py 调用 MinerU API 自动提取生成 -->\n"
    "<!-- 源文件：{name} | 模型：{model} | 提取时间：{time} -->\n"
)


class ExtractError(Exception):
    """单个文件（或单批任务）失败时抛出，携带退出码与给用户看的说明。"""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


# ------------------------------------------------------------
# 依赖自动安装（与 extract_local.py 保持一致的做法）
# ------------------------------------------------------------

DEPENDENCIES = [
    (["requests"], "requests", "调用 MinerU API"),
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
    """导入第三方模块；未安装时（经确认后）自动 pip 安装再重试。"""
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
# 命令行选项解析
# ------------------------------------------------------------

def pop_option(args, name):
    """取出 `--name value` 形式的选项并返回其值；未出现则返回 None。"""
    if name not in args:
        return None
    i = args.index(name)
    if i + 1 >= len(args):
        raise ExtractError(1, f"选项 {name} 缺少取值")
    value = args[i + 1]
    del args[i:i + 2]
    return value


def read_token_file(path):
    """从本地文件读取 Token。

    兼容两种写法：
      1) 纯 Token 一行：            abcdef123456
      2) .env 风格：                MINERU_API_TOKEN=abcdef123456
    忽略空行与 # 开头的注释行。
    """
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except OSError as e:
        raise ExtractError(3, f"读取 Token 文件失败：{path}（{e}）")
    for line in lines:
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line:
            key, _, value = line.partition("=")
            if key.strip() != TOKEN_ENV:
                continue
            line = value.strip()
        if line:
            return line
    raise ExtractError(3, f"Token 文件内容为空或格式不正确：{path}")


def resolve_token(argv_token, token_file):
    """确定 Token，优先级：--token > --token-file > 环境变量 > 默认文件。"""
    if argv_token:
        return argv_token.strip()

    if token_file:
        return read_token_file(pathlib.Path(token_file).expanduser())

    env_token = os.environ.get(TOKEN_ENV, "").strip()
    if env_token:
        return env_token

    for candidate in DEFAULT_TOKEN_FILES:
        if candidate.is_file():
            return read_token_file(candidate)

    default_hint = "、".join(str(p) for p in DEFAULT_TOKEN_FILES)
    raise ExtractError(
        3, f"未配置 MinerU API Token。请在 https://mineru.net/apiManage/docs "
           f"的\"API 管理\"页创建 Token，然后用以下任一方式配置："
           f"1) 运行参数 --token <token>；"
           f"2) 运行参数 --token-file <路径>；"
           f"3) 设置环境变量 {TOKEN_ENV}；"
           f"4) 把 Token 写入默认文件：{default_hint}")


def collect_items(args):
    """把文件路径参数校验并整理成待处理项列表。"""
    items = []
    seen_names = {}
    for arg in args:
        path = pathlib.Path(arg)
        if not path.is_file():
            raise ExtractError(1, f"文件不存在：{path}")
        if path.suffix.lower() not in SUPPORTED_EXTS:
            raise ExtractError(1, f"不支持的文件类型：{path.name}"
                                  f"（支持 {'/'.join(sorted(SUPPORTED_EXTS))}）")
        size = path.stat().st_size
        if size > MAX_FILE_SIZE:
            raise ExtractError(1, f"文件超过 200MB 上限：{path.name}"
                                  f"（{size / 1024 / 1024:.1f} MB）")
        if path.name in seen_names:
            print(f"[提醒] 存在同名文件 {path.name}，结果可能互相覆盖，建议改名后再提取")
        seen_names[path.name] = True
        items.append({"path": path, "data_id": uuid.uuid4().hex})
    return items


# ------------------------------------------------------------
# MinerU 接口调用
# ------------------------------------------------------------

def request_json(requests, method, url, token, **kwargs):
    """统一发起请求并校验 MinerU 的返回结构，失败时抛 ExtractError。"""
    headers = {"Authorization": f"Bearer {token}", "Accept": "*/*"}
    if "json" in kwargs:
        headers["Content-Type"] = "application/json"
    try:
        resp = requests.request(method, url, headers=headers, timeout=60, **kwargs)
    except requests.RequestException as e:
        raise ExtractError(4, f"网络请求失败：{e}")
    if resp.status_code != 200:
        raise ExtractError(4, f"MinerU 接口返回 HTTP {resp.status_code}：{resp.text[:200]}")
    try:
        payload = resp.json()
    except ValueError:
        raise ExtractError(4, f"MinerU 返回内容不是合法 JSON：{resp.text[:200]}")
    if payload.get("code") != 0:
        raise ExtractError(4, f"MinerU 接口返回错误（code={payload.get('code')}）："
                              f"{payload.get('msg')}")
    return payload.get("data") or {}


def apply_upload_urls(requests, token, items, model, ocr, page_ranges):
    """申请上传链接（同时创建解析任务），返回 batch_id。

    批量解析不支持直接上传文件，需先申请带签名的上传链接再 PUT 文件。
    page_ranges 非空时只解析指定页码（官方字段为每个 file 内的 page_ranges），
    用于超长/超页文档分批处理。
    """
    files = []
    for it in items:
        entry = {"name": it["path"].name, "data_id": it["data_id"], "is_ocr": ocr}
        if page_ranges:
            entry["page_ranges"] = page_ranges
        files.append(entry)
    body = {
        "files": files,
        "model_version": model,
        "enable_formula": True,
        "enable_table": True,
    }
    data = request_json(requests, "POST", BATCH_UPLOAD_URL, token, json=body)
    batch_id = data.get("batch_id")
    urls = data.get("file_urls") or []
    if not batch_id or len(urls) != len(items):
        raise ExtractError(4, f"申请上传链接返回异常：batch_id={batch_id}，"
                              f"链接数={len(urls)}（应为 {len(items)}）")
    for it, url in zip(items, urls):
        it["upload_url"] = url
    return batch_id


def upload_files(requests, items):
    """把本地文件逐个 PUT 到申请到的上传链接。

    上传完成后 MinerU 会自动扫描并提交解析任务，无需再调用提交接口。
    """
    for it in items:
        try:
            with open(it["path"], "rb") as f:
                resp = requests.put(it["upload_url"], data=f, timeout=600)
        except requests.RequestException as e:
            raise ExtractError(4, f"上传 {it['path'].name} 失败：{e}")
        except OSError as e:
            raise ExtractError(1, f"读取文件 {it['path'].name} 失败：{e}")
        if resp.status_code != 200:
            raise ExtractError(4, f"上传 {it['path'].name} 失败：HTTP {resp.status_code}")


def poll_batch(requests, token, batch_id, items, timeout):
    """轮询批量解析结果，直到所有文件都进入终态（done/failed）或超时。

    返回 {data_id: 结果条目} 的映射。
    """
    url = BATCH_RESULT_URL.format(batch_id=batch_id)
    expected = len(items)
    deadline = time.monotonic() + timeout
    results = {}

    while True:
        data = request_json(requests, "GET", url, token)
        for entry in data.get("extract_result") or []:
            key = entry.get("data_id") or entry.get("file_name")
            if key:
                results[key] = entry

        terminal = [e for e in results.values() if e.get("state") in ("done", "failed")]
        done = sum(1 for e in terminal if e.get("state") == "done")
        failed = sum(1 for e in terminal if e.get("state") == "failed")
        print(f"[进度] 已返回 {len(results)}/{expected} 个任务"
              f"（完成 {done}，失败 {failed}）")

        # 只有"返回数量够了"且"全部进入终态"才算结束
        if len(terminal) >= expected and len(results) >= expected:
            return results
        if time.monotonic() > deadline:
            raise ExtractError(4, f"等待解析结果超时（{timeout:.0f} 秒），"
                                  f"仅获取到 {len(results)}/{expected} 个任务")
        time.sleep(POLL_INTERVAL)


# ------------------------------------------------------------
# 结果下载与写出
# ------------------------------------------------------------

def find_entry(names, predicate):
    """在压缩包条目名中查找符合条件的第一个条目。"""
    for name in names:
        if predicate(name):
            return name
    return None


def extract_markdown(zip_bytes):
    """从结果压缩包中取出 full.md 正文，并尽量统计页数。"""
    try:
        zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
    except zipfile.BadZipFile:
        raise ExtractError(2, "解析结果压缩包无法读取（可能下载不完整）")

    with zf:
        names = zf.namelist()
        # full.md 可能在压缩包根目录，也可能在某层子目录下
        md_name = find_entry(names, lambda n: n == "full.md") or \
            find_entry(names, lambda n: n.endswith("/full.md"))
        if md_name is None:
            # 兜底：退而求其次取任意 .md（按路径长度取最短，通常是主文件）
            candidates = [n for n in names if n.lower().endswith(".md")]
            md_name = min(candidates, key=len) if candidates else None
        if md_name is None:
            raise ExtractError(2, "解析结果压缩包中未找到 Markdown 文件")

        text = zf.read(md_name).decode("utf-8", errors="replace")
        if not text.strip():
            raise ExtractError(2, "解析结果 Markdown 内容为空")

        return text, count_pages(zf, names)


def count_pages(zf, names):
    """从 content_list.json 推断页数（取最大 page_idx + 1）；失败返回 None。"""
    json_name = find_entry(names, lambda n: n.endswith("content_list.json"))
    if json_name is None:
        return None
    try:
        items = json.loads(zf.read(json_name).decode("utf-8", errors="replace"))
        pages = [int(it["page_idx"]) for it in items
                 if isinstance(it, dict) and "page_idx" in it]
        return max(pages) + 1 if pages else None
    except (ValueError, TypeError, KeyError):
        return None


def _download_via_http_client(zip_url, redirects=5):
    """用标准库 http.client 下载并跟随重定向。

    部分 CDN 在 requests/urllib3 下会在读取响应体时触发 TLS 层
    SSLEOFError，而标准库直连稳定，故下载统一走这条路径。
    """
    url = zip_url
    for _ in range(redirects + 1):
        parts = urlsplit(url)
        if parts.scheme == "https":
            conn = http.client.HTTPSConnection(parts.netloc, timeout=600)
        elif parts.scheme == "http":
            conn = http.client.HTTPConnection(parts.netloc, timeout=600)
        else:
            raise ExtractError(4, f"下载解析结果失败：不支持的地址 {parts.scheme}")

        target = parts.path or "/"
        if parts.query:
            target = f"{target}?{parts.query}"

        try:
            conn.request("GET", target)
            resp = conn.getresponse()
            if resp.status in (301, 302, 303, 307, 308):
                location = resp.getheader("Location")
                resp.read()
                if not location:
                    raise ExtractError(4, "下载解析结果失败：重定向缺少 Location")
                url = urljoin(url, location)
                continue
            if resp.status != 200:
                raise ExtractError(4, f"下载解析结果失败：HTTP {resp.status}")
            return resp.read()
        except ExtractError:
            raise
        except OSError as e:
            raise ExtractError(4, f"下载解析结果失败：{e}")
        finally:
            conn.close()

    raise ExtractError(4, "下载解析结果失败：重定向次数过多")


def download_result(requests, zip_url):
    """下载结果压缩包字节流（优先标准库，失败再退回 requests）。"""
    try:
        return _download_via_http_client(zip_url)
    except ExtractError:
        try:
            resp = requests.get(zip_url, timeout=600)
        except requests.RequestException as e:
            raise ExtractError(4, f"下载解析结果失败：{e}")
        if resp.status_code != 200:
            raise ExtractError(4, f"下载解析结果失败：HTTP {resp.status_code}")
        return resp.content


def write_extraction(path, model, text):
    """写出 {原文件名}-提取.md，并在文件头写入来源信息。"""
    header = HEADER_TEMPLATE.format(
        name=path.name,
        model=model,
        time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    out = path.with_name(f"{path.stem}-提取.md")
    out.write_text(header + "\n" + text.rstrip() + "\n", encoding="utf-8")
    return out


# ------------------------------------------------------------
# 单批处理
# ------------------------------------------------------------

def process_batch(requests, token, items, model, ocr, timeout, page_ranges):
    """处理一批文件：申请链接 → 上传 → 轮询 → 逐文件下载写出。

    返回各文件的退出码列表与成功计数（单个文件失败不中断其他文件）。
    """
    codes = []
    batch_id = apply_upload_urls(requests, token, items, model, ocr, page_ranges)
    print(f"[任务] 已创建解析任务 batch_id={batch_id}（{len(items)} 个文件），开始上传…")
    upload_files(requests, items)
    print("[任务] 上传完成，等待 MinerU 解析…")

    results = poll_batch(requests, token, batch_id, items, timeout)

    for it in items:
        path = it["path"]
        entry = results.get(it["data_id"]) or \
            next((e for e in results.values() if e.get("file_name") == path.name), None)

        if entry is None:
            print(f"[失败] {path.name}：未在解析结果中找到对应任务")
            codes.append(2)
            continue

        state = entry.get("state")
        if state != "done":
            reason = entry.get("err_msg") or "解析失败"
            print(f"[失败] {path.name}：{reason}")
            codes.append(2)
            continue

        zip_url = entry.get("full_zip_url")
        if not zip_url:
            print(f"[失败] {path.name}：解析完成但未返回结果下载地址")
            codes.append(2)
            continue

        try:
            text, pages = extract_markdown(download_result(requests, zip_url))
        except ExtractError as e:
            print(f"[失败] {e.message}")
            codes.append(e.code)
            continue

        out = write_extraction(path, model, text)
        page_info = f"，约 {pages} 页" if pages else ""
        print(f"[成功] {path.name} → {out.name}（{len(text)} 字符{page_info}）")
        codes.append(0)

    return codes


# ------------------------------------------------------------
# 主入口
# ------------------------------------------------------------

def install_all_deps():
    """--install-deps：把全部依赖一次装齐（已安装的自动跳过）。"""
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

    ocr = "--no-ocr" not in args
    args = [a for a in args if a != "--no-ocr"]

    try:
        token_arg = pop_option(args, "--token")
        token_file = pop_option(args, "--token-file")
        model = pop_option(args, "--model") or DEFAULT_MODEL
        timeout_arg = pop_option(args, "--timeout")
        page_ranges = pop_option(args, "--pages")
    except ExtractError as e:
        print(f"[失败] {e.message}")
        sys.exit(e.code)

    if model not in VALID_MODELS:
        print(f"[失败] --model 只能是 {' / '.join(VALID_MODELS)}，收到：{model}")
        sys.exit(1)

    try:
        timeout = float(timeout_arg) if timeout_arg else DEFAULT_TIMEOUT
    except ValueError:
        print(f"[失败] --timeout 需要是数字（秒），收到：{timeout_arg}")
        sys.exit(1)

    if not args:
        print(__doc__)  # 打印用法说明
        print("[失败] 请提供至少一个文件路径")
        sys.exit(1)

    try:
        token = resolve_token(token_arg, token_file)
        requests = ensure_module(["requests"], "requests", "调用 MinerU API")
        items = collect_items(args)
    except ExtractError as e:
        print(f"[失败] {e.message}")
        sys.exit(e.code)

    if page_ranges:
        print(f"[说明] 仅解析页码范围：{page_ranges}")

    worst = 0  # 记录最严重的退出码；全部成功时保持 0
    for start in range(0, len(items), BATCH_SIZE):
        batch = items[start:start + BATCH_SIZE]
        try:
            codes = process_batch(requests, token, batch, model, ocr, timeout, page_ranges)
        except ExtractError as e:
            print(f"[失败] {e.message}")
            codes = [e.code] * len(batch)
        worst = max([worst] + codes)
    sys.exit(worst)


if __name__ == "__main__":
    main()
