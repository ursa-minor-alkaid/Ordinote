# PDF/PPT/Word/图片 内容提取

## 功能定义

将 PDF/PPT/Word/图片 文件中的文字内容提取出来，生成提取稿（中间文档）。本功能**只负责提取**：提取稿**不是最终笔记**，后续通过 `ordinote-wash`（初始化整理，含修错别字、去噪、公式转写等整理规则）或 `ordinote-add`（添加内容）功能将其转化为笔记。Word（.docx）由 `extract_docx.py` 本地转换并保留格式，输出可直接作为正式笔记。

---

## 调用脚本提取（文字提取）

1. **选择提取方式**：`scripts/` 目录下提供三种提取脚本：
   - `scripts/extract_pdf_pptx.py`：本地 Python 脚本直接提取（适用于**文字版** PDF/PPT；详见下文「PDF/PPT 提取」一节；已实现）
   - `scripts/extract_docx.py`：本地 Python 脚本把 **Word（.docx）** 转换为 markdown（保留格式信息；详见下文「Word 提取」一节；已实现）
   - `scripts/mineru_api.py`：调用 MinerU 精准解析 API 提取（适用于扫描件、复杂版式；支持 PDF/PPT/DOC 及**图片**，且支持**一次上传多个文件批量提取**；已实现，需先配置 Token）
   - 默认先尝试本地脚本；失败或效果差时再改用 MinerU API（旧版 `.doc` 本地脚本不支持：请先用 Word 另存为 `.docx`，或改用 MinerU API）
2. **运行脚本提取文字**，生成提取文件：
   - 输出位置：与源文件同一目录
   - 命名规则：`{原文件名}-提取.md`（或 .txt）
3. **失败处理（必须遵守）**：
   - 脚本无法运行、无法提取、或提取结果为空 → **必须如实告知用户**，说明具体原因
   - 如实告知用户，不要试图用其他方式替代提取
   - 不得编造、猜测或凭记忆补写文件内容
   - 可向用户建议排查方向：文件是否为扫描件、加密、损坏，路径或格式是否有误

---

## PDF/PPT 提取（本地 · scripts/extract_pdf_pptx.py）

从 PDF / PPTX 文件中提取全部文字，生成 markdown 提取稿。适用于**文字版** PDF 与 PPT；扫描版 / 图片型 PDF 提取不到文字时（退出码 2）改用 MinerU API。提取稿是「原始素材」，后续按 `ordinote-wash` 规则整理。脚本只读源文件，不改动源文件。

### 运行

```bash
python scripts/extract_pdf_pptx.py <文件路径> [更多文件路径...] [选项]
```

- 可一次传入多个文件逐个提取；单个文件失败只记录错误，不中断其他文件
- 仅支持 `.pdf` / `.pptx`；旧版 `.ppt` 不支持（退出码 1，提示先用 PowerPoint 另存为 `.pptx`）

### 输出内容

- 输出到源文件同目录，命名为 `{原文件名}-提取.md`（同名旧提取文件会被直接覆盖）
- 各页 / 各张幻灯片之间以 `---` 分隔线分开（保留源结构线索，整理时保留）
- PPT 的演讲者备注以 `> 备注：` 形式附在对应幻灯片之后；表格按制表符分隔保留结构，组合形状内的文字递归提取
- 空白页 / 无文字幻灯片以 HTML 注释占位，便于核对缺页；文件头以 HTML 注释记录来源与提取时间
- 提取完成后打印字符数统计与空白页提醒，供核对提取质量

### 常用选项

| 选项 | 说明 |
|---|---|
| `--no-auto-install` | 缺少依赖库（PyMuPDF / python-pptx）时不自动安装 |
| `--install-deps` | 只安装依赖后退出（供交接前预热） |

退出码：`0` 全部成功；`1` 用法错误 / 文件不存在 / 不支持的类型（含旧版 `.ppt`）；`2` 提取失败或结果为空（典型：扫描版 PDF）；`3` 缺少依赖库 PyMuPDF / python-pptx，且自动安装被拒绝或失败

---

## Word 提取（本地 · scripts/extract_docx.py）

将 Word（.docx）转换为 markdown，尽可能保留 Word 的格式信息。适用于**AI 生成正文、楷体手写批注**的 Word 笔记（如各版本纵向笔记）。脚本只读源文件，不修改 Word 文档。

### 运行

```bash
python scripts/extract_docx.py <文件或目录> [更多...] [选项]
```

- 目录输入会处理其下所有 `.docx`（不递归）；单个文件失败只记录错误，不中断其他文件
- 旧版 `.doc` 不支持（退出码 1，提示先用 Word 另存为 `.docx`，或改用 MinerU API）

纵向笔记场景常用组合（批量转换存档目录、去除版本号、输出到正式笔记目录）：

```bash
python scripts/extract_docx.py "archives\纵向笔记（临时）" --strip-version --out "纵向笔记"
```

### 格式保留与项目约定

- 标题：首段为「无缩进 / No Spacing / Title / 标题」样式时作为文档主标题输出 `#`，其余标题样式顺延一级（`Heading 1` → `##`）
- 生成信息行：第二段整行斜体（本项目约定）
- 行内格式：加粗、斜体、粗斜、删除线、下划线 `<u>`、荧光高亮 `<mark>`
- 批注字体（默认 楷体/KaiTi/SimKai）→ 淡灰色 `<mark>` 包裹；本项目约定**霞鹜文楷为 AI 生成正文、楷体为手写批注**，精确匹配不误伤霞鹜文楷
- Cambria Math 公式字体中的粗体/斜体视为数学排版样式，不转为强调
- 列表（多级嵌套与编号计数）、表格（pipe 表格，合并单元格去重）、超链接、引用样式（相邻引用段合并）、软回车 `<br>`
- 内嵌图片导出到输出目录 `images/` 子文件夹，正文以相对路径引用
- 不保留 / 不提取：字号、颜色、对齐等 markdown 无法表达的格式；页眉页脚、脚注尾注、文本框、审阅批注；修订删除的内容

### 常用选项

| 选项 | 说明 |
|---|---|
| `--out <目录>` | 输出目录（默认：源文件所在目录） |
| `--suffix <后缀>` | 输出文件名后缀：`{名}-{后缀}.md`（默认：无） |
| `--strip-version` | 去除文件名末尾 `_v` / `_v0.5` 之类版本号 |
| `--annotation-fonts <字体>` | 批注字体，逗号分隔（默认：楷体,KaiTi,SimKai） |
| `--no-annotation` | 不做批注字体标记 |
| `--no-title-shift` | 首段不作为文档主标题，标题保持原层级 |
| `--no-meta-italic` | 第二段不做整行斜体处理 |
| `--no-images` | 不导出内嵌图片（以 HTML 注释占位） |
| `--no-auto-install` | 缺少 python-docx 时不自动安装 |

退出码：`0` 全部成功；`1` 用法错误 / 不支持的类型（含旧版 `.doc`）；`2` 转换失败或结果为空；`3` 缺少依赖 python-docx，且自动安装被拒绝或失败

---

## MinerU API 提取（云端备选 · scripts/mineru_api.py）

当本地脚本无法提取（典型：扫描版 / 图片型 PDF）或提取效果差时，改用 MinerU 精准解析 API。

### Token 配置（首次使用必做）

1. 登录 <https://mineru.net/apiManage/docs>，进入「API 管理」页，自行创建一个 Token
2. 用 **PowerShell 命令**把 Token 配置为环境变量（脚本只从环境变量读 Token，不在磁盘上留明文 Token 文件）：

   **永久配置（推荐，一次性）**——写入用户环境变量，重开终端 / Trae 后生效：
   ```powershell
   [Environment]::SetEnvironmentVariable("MINERU_API_TOKEN", "你的token", "User")
   ```

   仅当前会话有效（临时测试用）：
   ```powershell
   $env:MINERU_API_TOKEN = "你的token"
   ```

   验证是否配置成功：
   ```powershell
   [Environment]::GetEnvironmentVariable("MINERU_API_TOKEN", "User")
   ```

   macOS / Linux（写入 shell 配置，永久生效）：
   ```bash
   echo 'export MINERU_API_TOKEN="你的token"' >> ~/.zshrc   # 或 ~/.bashrc
   ```

- 也支持临时参数 `--token 你的token` 覆盖环境变量（不推荐长期使用：会遗留在命令历史里）
- 优先级：`--token` > 环境变量 `MINERU_API_TOKEN`

- 如果上述任务无法完成，则将对应步骤告知用户

### 运行

```bash
python mineru_api.py <文件路径> [更多文件路径...]
```

常用选项：

| 选项 | 说明 |
|---|---|
| `--token <token>` | 临时指定 Token，覆盖环境变量（不推荐长期使用，会留在命令历史里） |
| `--model <name>` | 模型版本：`vlm`（脚本默认，推荐）/ `pipeline` / `MinerU-HTML` |
| `--pages <范围>` | 只解析指定页码，如 `1-200`、`2,4-6`；用于超长/超页文档分批解析 |
| `--no-ocr` | 关闭 OCR（默认开启；扫描件建议保持开启） |
| `--timeout <秒>` | 等待解析结果的超时，默认 900 秒 |
| `--no-auto-install` | 缺依赖库（requests）时不自动安装 |
| `--install-deps` | 只安装依赖后退出 |

### 支持的文件类型与批量上传

**一、支持图片（可直接上传图片提取文字）**

MinerU 不仅能解析 PDF/PPT，还能**直接上传图片**提取文字（等价于对图片做 OCR）。支持的输入格式：

- 文档：`pdf` / `doc` / `docx` / `ppt` / `pptx`
- 图片：`png` / `jpg` / `jpeg` / `jp2` / `webp` / `gif` / `bmp`

把图片路径当成普通文件传给脚本即可（OCR 默认开启，正适合图片/扫描件；如需关闭用 `--no-ocr`）：

```bash
python mineru_api.py 扫描页1.png 扫描页2.jpg
```

**二、支持一次上传多个文件（批量）**

脚本可接收多个文件路径，**一次性批量提交解析**，PDF 与图片可混在一起；每个文件各自生成 `{原文件名}-提取.md`。文件较多时会自动分批（脚本每批 ≤ 50 个，官方单批上限 200 个）：

```bash
python mineru_api.py 第一章.pdf 第二章.pdf 架构图.png
```

注意：

- 批量中的**每个文件都各自遵守 200 MB / 600 页上限**（不是合计上限）
- 批量上传与图片解析不影响 Token 配置和输出命名规则
- 若同一批里存在同名文件，结果可能互相覆盖，脚本会给出提醒

### 与本地脚本的差异

- 输出同样为 `{原文件名}-提取.md`，但**正文中没有 `---` 分页线**：MinerU 只返回整篇 Markdown（取自结果压缩包的 `full.md`），不返回逐页文字
- Word（`.docx`）提取优先用本地 `scripts/extract_docx.py`；旧版 `.doc`、扫描件或复杂版式的 Word 文档可改用 MinerU API
- 内容为 MinerU 解析后的干净 Markdown（含公式 LaTeX、表格等）；后续用 `ordinote-wash` 整理时仍要检查有无缺页
- 单文件大小上限 200 MB、页数上限 600 页（精准解析 API 官方限制）；多个文件会自动分批（每批 ≤ 50 个）上传解析
- 超过 600 页的超长文档：用 `--pages` 按页范围多次提取（如 `--pages 1-600`、`--pages 601-1200`），再人工/后续功能合并
- 注意：网上常见的"10 MB / 20 页"限制属于 **Agent 轻量解析 API**（免登录那个），与我们使用的精准解析 API 无关

---

## 注意事项

1. **不删除原始提取文件**（`{原文件名}-提取.md`）：它是核对提取质量的依据；用户要求删除时才删除
2. **汇报义务**：完成后向用户汇报——提取方式、生成的文件、提取质量（是否完整、有无乱码或缺页）

