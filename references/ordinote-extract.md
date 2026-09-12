# PDF/PPT/图片 内容提取

## 功能定义

将 PDF/PPT/图片 文件中的文字内容提取出来，生成提取稿（中间文档）。本功能**只负责提取**：提取稿**不是最终笔记**，后续通过 `ordinote-wash`（初始化整理，含修错别字、去噪、公式转写等整理规则）或 `ordinote-add`（添加内容）功能将其转化为笔记。

---

## 调用脚本提取（文字提取）

1. **选择提取方式**：`scripts/` 目录下提供两种提取脚本：
   - `scripts/extract_local.py`：本地 Python 脚本直接提取（适用于**文字版** PDF/PPT；已实现）
   - `scripts/mineru_api.py`：调用 MinerU 精准解析 API 提取（适用于扫描件、复杂版式；支持 PDF/PPT/DOC 及**图片**，且支持**一次上传多个文件批量提取**；已实现，需先配置 Token）
   - 默认先尝试本地脚本；失败或效果差时再改用 MinerU API
2. **运行脚本提取文字**，生成提取文件：
   - 输出位置：与源文件同一目录
   - 命名规则：`{原文件名}-提取.md`（或 .txt）
3. **失败处理（必须遵守）**：
   - 脚本无法运行、无法提取、或提取结果为空 → **必须如实告知用户**，说明具体原因
   - 如实告知用户，不要试图用其他方式替代提取
   - 不得编造、猜测或凭记忆补写文件内容
   - 可向用户建议排查方向：文件是否为扫描件、加密、损坏，路径或格式是否有误

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
- 内容为 MinerU 解析后的干净 Markdown（含公式 LaTeX、表格等）；后续用 `ordinote-wash` 整理时仍要检查有无缺页
- 单文件大小上限 200 MB、页数上限 600 页（精准解析 API 官方限制）；多个文件会自动分批（每批 ≤ 50 个）上传解析
- 超过 600 页的超长文档：用 `--pages` 按页范围多次提取（如 `--pages 1-600`、`--pages 601-1200`），再人工/后续功能合并
- 注意：网上常见的"10 MB / 20 页"限制属于 **Agent 轻量解析 API**（免登录那个），与我们使用的精准解析 API 无关

---

## 注意事项

1. **不删除原始提取文件**（`{原文件名}-提取.md`）：它是核对提取质量的依据；用户要求删除时才删除
2. **汇报义务**：完成后向用户汇报——提取方式、生成的文件、提取质量（是否完整、有无乱码或缺页）

