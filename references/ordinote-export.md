# 笔记批量导出 PDF / Word

## 功能定义

把 markdown 笔记导出为 PDF 或 Word（.docx），支持单个文件，也支持整个目录批量导出。两种格式使用不同脚本与渲染方案：

- **PDF**（`scripts/md2pdf.py`）：版式复用本机 Markdown Preview Enhanced（MPE）扩展自带的样式与图表资源，观感与 MPE 预览一致。
- **Word**（`scripts/md2docx.py`）：用 python-docx 直接生成 .docx，尽量忠实还原 markdown 原始效果，包括标题等级、文字高亮与背景色、图片插入、表格、列表等。Word 导出不依赖 MPE 扩展与浏览器。

本功能**只负责导出**：源笔记只读，导出过程不修改、不移动、不重命名任何笔记文件。导出物属于对外输出物，不参与 `ordinote-wash` / `ordinote-standardize` 等整理链路。

用户说"导出笔记"但未指定格式时，主动询问需要 PDF 还是 Word（可两者都要），然后查阅对应章节执行。

---

## 一、PDF 导出（md2pdf.py）

### 前置条件（首次使用时向用户说明）

1. **本机类 VS Code 客户端需安装 Markdown Preview Enhanced 扩展**（Trae / VS Code / Cursor 等任选其一）。脚本复用该扩展自带的主题样式（`crossnote/styles/`）与 mermaid 渲染库（`crossnote/dependencies/mermaid/`）。**这一步需要用户自行安装，脚本不会自动安装扩展**；未安装时脚本会直接报错并给出提示。
2. **本机需装有 Chrome / Edge / Chromium**，用于无头打印。
3. 解析环节依赖 Python 库 markdown-it-py（多数环境随其它 Python 包自带，缺失时脚本会提示 `pip install markdown-it-py`）。
4. 公式渲染引擎取自本机已装扩展内的 KaTeX；找不到时公式以原始 LaTeX 文本呈现，不影响其余内容导出。
5. 除上述已有资源外**不下载任何东西**：不需要额外 `pip install`，不需要 `npm install`，不上 CDN。

### 运行

```bash
python scripts/md2pdf.py <文件或目录路径> [更多路径...]
```

路径可以是单个 `.md`、多个 `.md`，或目录（递归收集目录下的 `*.md`）。

常用选项：

| 选项 | 说明 |
|---|---|
| `--out <目录>` | 输出目录，默认与源文件同目录 |
| `--theme <名称>` | 预览主题，默认 `github-light`，可取 `crossnote/styles/preview_theme/` 下的任意主题名 |
| `--prism <名称>` | 代码块配色主题，默认 `github` |
| `--paper <尺寸>` | 纸张，默认 `A4` |
| `--margin <边距>` | 页边距，默认 `1.4cm 1.3cm` |
| `--wait <毫秒>` | 页面渲染等待时间，默认 3000；图表较多时可调大 |
| `--skip-existing` | 已有同名 PDF 时跳过 |
| `--exclude <正则>` | 按路径正则排除文件 |
| `--chrome <路径>` | 手动指定浏览器可执行文件 |
| `--mpe <目录>` | 手动指定 MPE 扩展目录或其 `crossnote` 子目录 |

示例：

```bash
python scripts/md2pdf.py 笔记/生物 --out 导出 --skip-existing
python scripts/md2pdf.py 第一章.md 第二章.md --theme vue --paper A3
```

### PDF 渲染情况与降级

| 内容 | 处理方式 |
|---|---|
| 标题、段落、列表、表格、引用 | 由 markdown-it 解析，套 MPE 主题样式 |
| mermaid 代码块 | 用 MPE 自带的 mermaid 库渲染为矢量图形，与预览一致 |
| 公式（`$...$` / `$$...$$`） | 在 markdown 解析前抽取，交由 KaTeX 渲染，字体内嵌进 PDF；支持 `array`、`aligned`、`pmatrix`、`cases` 等环境与 `\\` 换行，跨行的块级公式不受影响 |
| `tikz` 代码块 | **不渲染**，在 PDF 中以代码文本呈现（原因见下） |
| 代码块 | 使用 MPE 的等宽字体与配色；**无语法高亮**（分词着色由 MPE 扩展内部完成，本机没有可独立引用的分词脚本） |
| 相对路径图片 | 正常加载（按源文件所在目录解析，兼容中文文件名） |
| 远程图片（如 shields.io 徽标） | 需要联网，属于笔记自带外链 |
| MPE 专有语法（admonition、`@import`、TOC 指令、代码块属性等） | 不渲染，按普通文本处理 |

关于 tikz：MPE 的 tikz 由扩展宿主一侧内置的 wasm TeX 引擎渲染（TikZJax 系，资源位于扩展的 `out/tex/` 目录），该引擎打包在扩展代码内部，外部脚本无法调用。因此在导出的 PDF 里 tikz 图不会出现，只保留源码文本。若笔记包含 tikz 图且必须出现在 PDF 中，改用 MPE 预览手动导出该图，或将图另存为图片再嵌入笔记。

降级行为：

- 未找到 MPE 扩展 → 直接报错退出，并提示用户安装扩展或指定路径
- 未找到 KaTeX → 打印提示，公式以原始 LaTeX 文本呈现，其余内容照常导出
- 未找到浏览器 → 直接报错退出，提示安装或用 `--chrome` 指定

---

## 二、Word 导出（md2docx.py）

### 前置条件（首次使用时向用户说明）

1. 基础依赖两个 Python 库：**python-docx** 与 **markdown-it-py**。缺失时脚本直接报错并给出安装提示：

    ```bash
    pip install python-docx markdown-it-py
    ```

2. 公式转换需要 **latex2mathml**（首次运行时脚本自动 `pip install`）与本机 Office 附带的 **MML2OMML.XSL**（脚本自动查找 `Microsoft Office/root/Office*/`，也可用 `--omml-xsl` 指定）。缺少任一条件时公式退回 LaTeX 原文，不影响其余内容。

3. 不需要 MPE 扩展；除远程图片加载与首次安装 latex2mathml 外不需要联网。

### 运行

```bash
python scripts/md2docx.py <文件或目录路径> [更多路径...]
```

路径可以是单个 `.md`、多个 `.md`，或目录（递归收集目录下的 `*.md`）。

常用选项：

| 选项 | 说明 |
|---|---|
| `--out <目录>` | 输出目录，默认与源文件同目录 |
| `--skip-existing` | 已有同名 .docx 时跳过 |
| `--exclude <正则>` | 按路径正则排除文件 |
| `--omml-xsl <路径>` | 指定 MML2OMML.XSL；默认自动查找本机 Office |

示例：

```bash
python scripts/md2docx.py 笔记/生物 --out 导出 --skip-existing
python scripts/md2docx.py 第一章.md 第二章.md
```

### Word 还原情况

| markdown 内容 | Word 中的还原方式 |
|---|---|
| 标题 `#` ~ `######` | Word 内置 Heading 1-6 样式，字号 13 / 12 / 11 / 10.5 / 10 / 9.5，H1/H2 带下边框，可直接用于生成导航与目录。全文第一个一级标题视为文件标题：居中、18 号、加粗，其余一级标题保持常规样式 |
| 正文、段落 | 8.5 号字，Calibri + 微软雅黑，1.1 倍行距，段后 3pt |
| 纸张与页边距 | A4；左右 1.8cm，上下 1cm |
| 加粗、斜体、粗斜、删除线、下划线 | 对应 Word 字符格式（`<u>` 标签同样识别） |
| 行内代码 | Consolas 等宽字体 + 浅灰底纹 |
| 代码块（围栏） | 整段浅灰底纹、边框与等宽字体，保留换行与缩进；语言标注不保留 |
| 有序 / 无序列表 | Word 原生编号与项目符号，支持任意层级嵌套；同一列表项的多段落作为续行缩进，不重复编号；每个有序列表自动从 1 重新编号 |
| 引用块 `>` | 左侧灰色竖线、缩进、灰色文字 |
| 表格 | Word 原生表格（Table Grid 边框），首行表头加粗加底纹，单元格内保留行内格式 |
| 分割线 `---` | 段落下边框横线 |
| 链接 | Word 超链接，蓝色下划线，可点击跳转（页内锚点与非网址链接按普通文本处理） |
| 图片 `![](...)` | 内嵌插入，相对路径按源文件所在目录解析（兼容中文文件名与百分号编码）；支持 http(s) 远程图片与 `data:` URI。**尊重 markdown 中指定的尺寸**：HTML 写法 `<img src=... width=400px height=...>` 的 `width` / `height` 属性，以及 `style="width:...;height:..."`（style 优先）均生效；支持 px（无单位数字按 px）、`%`（相对版心宽度）、pt、cm、mm、in、em 单位；只指定一边时另一边按原生比例自动推导；指定宽度超出版心或位于缩进列表内时，分别按版心/缩进后可用宽度等比缩小。**实际插入宽度为该尺寸的 85%**（未指定尺寸时按原生大小、超版心缩放后再取 85%），用于减小图片占幅 |
| `<mark>` 高亮 | 文字底纹（w:shd）：无样式属性时为黄色，`style="background-color: ..."` 时按指定颜色还原（如 lightpink 淡粉、`#e6e6e6` 淡灰） |
| `<span style="background-color:...;color:...">` | 分别还原为文字底纹与字体颜色；支持命名色、`#rgb` / `#rrggbb` / `#rrggbbaa`、`rgb(...)` |
| 公式 `$...$` / `$$...$$` | 在 markdown 解析前抽取，经 latex2mathml 转 MathML，再由 Office 的 MML2OMML.XSL 转成 **Word 原生公式对象**：行内公式嵌在段落中，块级公式独占段落并居中；`aligned` / `gather` 等多行环境按行拆分。转换失败（含 `\ce` 等 mhchem 命令）时退回 Cambria Math 按 LaTeX 原文呈现并计数 |
| mermaid / tikz 等图表代码块 | **不渲染**，与普通代码块一样保留源码文本 |
| YAML front-matter | 不写入文档 |
| MPE 专有语法（admonition、`@import`、TOC 指令等） | 不识别，按普通文本处理 |

关于公式：转换链路为 latex2mathml → MathML → Office 的 MML2OMML.XSL → OMML。`\begin{aligned}` 等多行环境先按顶层 `\\` 拆行、去掉对齐符 `&`，每行生成一个居中公式段落；`\ce` / `\pu` 等 mhchem 命令、latex2mathml 不支持的宏，以及找不到 XSLT 或 latex2mathml 时，该公式退回 LaTeX 原文并在汇总中计数，不中断导出。

图片相关降级：

- 本地图片文件不存在或无法识别 → 在对应位置插入红色提示文字 `[图片加载失败：...]` / `[图片无法识别：...]`，运行汇总中提示未嵌入的图片数量，其余内容照常导出
- 远程图片加载失败（无网络、链接失效等）→ 同上处理
- 一篇笔记中可以包含多张图片，单张失败不影响其余图片与整篇导出

---

## 三、输出与退出码（两种格式通用）

- 默认输出到源文件同目录，文件名与源文件同名、扩展名分别为 `.pdf` / `.docx`；同名文件默认覆盖（`--skip-existing` 时跳过）
- 每次运行结束时输出汇总：成功 / 跳过 / 失败的数量，失败项附原因
- 退出码：`0` 全部成功，`1` 缺少必需资源、依赖或参数错误，`2` 存在导出失败的文件

---

## 四、失败处理（必须遵守）

1. 脚本报错或某个文件导出失败时，**必须如实告知用户**，说明具体原因（未找到扩展、缺少 Python 库、渲染超时、输出为空、图片缺失等），不得声称已完成导出
2. Word 导出若提示缺少 python-docx / markdown-it-py，告知安装命令 `pip install python-docx markdown-it-py`，由用户自行安装后重试，不要伪造导出结果；latex2mathml 会自动安装，若安装失败或缺少 MML2OMML.XSL，公式退回 LaTeX 原文并在汇总中计数，需如实说明
3. 笔记含 `tikz` 代码块时，**必须主动说明该图在导出物中未渲染**（保留为代码文本），并提供替代做法（mermaid 图仅 PDF 可渲染，Word 中同样只保留文本；需要图表入 Word 时先将图另存为图片再嵌入笔记）
4. Word 导出后若提示有图片未嵌入或公式未转换，**必须告知具体是哪篇笔记、第几处**，并说明可能原因（相对路径变化、文件缺失、远程链接需联网；公式可能含 mhchem 命令或 latex2mathml 不支持的宏）
5. 不得改用其它方式替代（例如改用其它工具的默认模板导出），也不得编造导出结果
6. 单个文件失败不会中断整批，脚本会继续处理其余文件，最后统一汇总
7. PDF 排查方向可提示用户：是否已安装 MPE 扩展、是否装有 Chrome、笔记是否引用了需要联网的远程图片、`--wait` 是否需要调大

---

## 五、注意事项

1. **不修改源笔记**：PDF 导出的临时 HTML 与浏览器 profile 放在系统临时目录用后即删；Word 导出只新建 .docx 文件
2. PDF 需要修改版式时，优先用 `--theme`、`--paper`、`--margin` 三个选项，不要改脚本
3. 导出属于工作流末端动作，通常在笔记整理完成后进行，不打断 `wash` / `standardize` / `add` 等链路
4. **汇报义务**：完成后向用户汇报导出格式与范围（哪些文件、输出到哪个目录）、成功与失败数量，以及失败原因与未渲染内容（tikz / mermaid / 缺失图片）
