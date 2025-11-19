# 能效报告自动化工具

该项目用于批量生成能效报告所需的图表/表格，并组装一个轻量级的 DOCX 示例报告。图表逻辑定义在 `chart_generation_logic.jsonl` 中，标准化数据存放在 `报告数据_标准化.xlsx`，根目录下的脚本会把这些输入转换为 PNG 图像、Excel 表格以及最终的 Word 文档。

## 环境要求

- Windows 平台上的 Python 3.10 及以上版本（脚本依赖 Windows 字体路径）
- 建议安装的核心第三方库：
  - `python-docx`
  - `pandas`、`numpy`、`matplotlib`
  - `openpyxl`
- 请确保 `C:\Windows\Fonts` 中存在微软雅黑 (`msyh.ttc` 或 `msyh.ttf`)，以便 Matplotlib 与 python-docx 正确渲染中文。

如需创建虚拟环境，可执行：

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

若暂时没有 `requirements.txt`，请手动安装上述依赖。

## 生成图表与表格

1. 检查 `chart_generation_logic.jsonl` 是否覆盖所需图表（目前为 CH001–CH019，含 CH013A、CH016A、CH016B 等扩展编号，新加入的 CH016B 会输出“照明分区回路数据表”）。
2. 确认 `报告数据_标准化.xlsx` 中的数据已经更新。
3. 运行图表生成脚本：

```
python auto_chart_generator.py
```

所有 PNG/XLSX 产物将输出到 `charts/generated/`。脚本已预设使用微软雅黑，确保坐标轴与标题可读。

## 构建示例报告

`demo_bulid.py` 会读取 JSONL 配置，按 `chapter_title` 聚合图表，并为每个图表挑选最合适的 PNG 插入到 Word 文件中。文档样式与 `附件6_报告模板_2.docx` 对齐：标题 15 磅、章节标题 14 磅、图表标题与正文 12 磅，均使用微软雅黑。

```
python demo_bulid.py --output demo_report.docx
```

可选参数：

- `--logic-path PATH`：指定其他 JSONL 配置文件。
- `--chart-dir PATH`：改为读取指定目录下的 PNG。
- `--output PATH`：自定义输出 DOCX 路径。

若图像缺失，脚本会在文档中标注 “Image not found.” 以便排查。

## 项目结构

- `auto_chart_generator.py`：负责调度数据、导出图表/表格。
- `demo_bulid.py`：按照模板样式组装 DOCX 示例报告。
- `charts/generated/`：存放按图表编号命名的 PNG/XLSX 结果。
- `chart_generation_logic.jsonl`：驱动上述脚本的图表元数据。
- `报告数据_标准化.xlsx`：已清洗的输入数据源。

## 故障排查

- **Word 中文字体异常**：确认已安装微软雅黑并重新运行 `demo_bulid.py`。
- **Matplotlib 找不到字体**：安装微软雅黑或调整 `auto_chart_generator.py` 中的 `FONT_CANDIDATES`。
- **报告中缺少图像**：检查 `charts/generated/` 是否存在形如 `CH0XX_*.png` 的文件。

如需更复杂的分析或模板导出，可查看根目录下的辅助脚本（如 `generate_main.py`、`extract_docx_images.py` 等）。

