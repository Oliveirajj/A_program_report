# 能效报告自动化工具

该项目在 Windows 环境下自动完成“数据 → 图表/表格 → 模板报告 → AI 分析文本 → `能耗报告_自动生成.docx`”的整套流程。输入数据由 `报告数据_标准化.xlsx` 与 `chart_generation_logic.jsonl` 描述，输出结果由脚本自动写入 Word 文档。

---

## 环境与依赖

- Windows 10/11 + Python 3.10 及以上版本（脚本依赖系统字体与 COM 字体路径）
- 默认字体为微软雅黑，请确认 `C:\Windows\Fonts\msyh.ttc` 存在
- 主要依赖：`python-docx`、`pandas`、`numpy`、`matplotlib`、`openpyxl`、`dashscope`、`python-dotenv`

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

> `text_generator.py` 会从 `.env` 中读取 `QWEN_API_KEY`，请在根目录创建 `.env` 并写入：
>
> ```
> QWEN_API_KEY=your_dashscope_api_key
> ```

---

## 关键输入文件

- `chart_generation_logic.jsonl`：定义 CH 编号、中文标题、章节归属等图表元数据
- `报告数据_标准化.xlsx`：生成所有图表/表格所需的指标数据
- `附件6_报告模板_2.docx`（已抽取样式，用于 `demo_bulid.py` 中的字体尺寸配置）

若逻辑或数据更新，请同步维护以上文件。

---

## 一键生成 `能耗报告_自动生成.docx`

1. **生成 CH 图表与表格**
   ```powershell
   python auto_chart_generator.py
   ```
   - 输出：`charts/generated/CHxxx_*.png` 与对应的 `.xlsx`
   - 若需要自定义数据源或输出目录，可在脚本内调整参数

2. **按章节组装底稿 `demo_report.docx`**
   ```powershell
   python demo_bulid.py --logic-path chart_generation_logic.jsonl --chart-dir charts\generated --output demo_report.docx
   ```
   - 依据 JSONL 中的 `chapter_title` 分组，每个 CH 插入最匹配的 PNG
   - 若缺图，会在文档中写入 `Image not found.` 便于核对

3. **AI 生成章节分析并插回图表**
   ```powershell
   python text_gen.py `
     -i demo_report.docx `
     -o 能耗报告_自动生成.docx `
     --chart-logic chart_generation_logic.jsonl `
     --chart-dir charts\generated
   ```
   - `text_gen.py` 会逐章解析 `demo_report.docx`，调用 DashScope（qwen-vl-plus）生成正文
   - 章节文本中的 `(CHxxx)` 标记会自动被替换成对应的图表图片；无需准备 `ref/` 参考图
   - 若章节正文为空，会回退到模板提示词（`text_generator.get_template_guidance`）

完成后，`能耗报告_自动生成.docx` 即包含：
- 章节级别的生成式分析文字
- 自动清理的旧图表/标题
- 依据 `(CHxxx)` 标记重新插入的最新图表

---

## 常见问题

- **无法连接 DashScope**：确认 `.env` 中 `QWEN_API_KEY` 正确，并可通过 `pip show dashscope` 验证 SDK 安装
- **图表缺失或编号对不上**：检查 `charts/generated/` 是否已有对应 `CHxxx` PNG；若文件名不规范，请参考 `demo_bulid.py` 中的 `find_chart_image`
- **章节未生成文本**：确认 `demo_report.docx` 中该章节有可读段落，或在 `template_sections.json` 中补充模板提示
- **字体/中文错乱**：安装/修复微软雅黑，重跑 `auto_chart_generator.py` 与 `demo_bulid.py`

---

## 项目结构速览

- `auto_chart_generator.py`：读取 `报告数据_标准化.xlsx` 并输出所有 CH 图表/表格
- `demo_bulid.py`：按照章节将图表插入 DOCX 底稿
- `text_gen.py`：生成章节文本并插入图表，产出最终报告
- `text_generator.py`：封装 DashScope 接口、模板提示、章节上下文抽取
- `chart_generation_logic.jsonl`：图表与章节映射
- `charts/generated/`：所有 CH PNG + 数据表
- 其他 `process_*.py` / `analyze_*.py`：为特殊场景准备的数据清洗与校验脚本（可按需参考）

如需扩展模板或替换 AI 模型，可在 `text_generator.py` 中调整提示词或模型 ID，并在 README 中补充新的运行说明。

