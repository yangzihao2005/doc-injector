# PROGRESS.md

## 当前状态
- [x] v0.1.0 已完成并通过质量门禁

## 进展记录

### 2026-10-23
- 完成 `src/schema.py`：Pydantic 契约（表/图/内容块/payload，行宽与图片比例校验）。
- 完成 `src/scanner.py`：标题优先 + 标点容错正则的锚点定位、版心宽度计算、锚点后就近插入。
- 完成 `src/style_cloner.py`：Run 级样式克隆，显式写入 `ascii/hAnsi/cs/eastAsia`，规避中文字体降级。
- 完成 `src/table_engine.py`：三线表 / 全框网格表、表题在上、`cantSplit`、`tblHeader`。
- 完成 `src/image_engine.py`：按版心比例等比缩放、图片居中、图题在下。
- 完成 `src/injector.py`：顺序注入、原子保存、缺锚点策略、只读/占用/损坏模板的结构化文本降级。
- 完成 `src/cli.py` + `src/demo_assets.py`：`inspect` / `inject` / `demo` / `schema`，生成样例模板与波形图。
- 生成交付样本：`templates/sample_dsp.docx`、`examples/waveform.png`、`examples/sample_payload.json`。
- 质量门禁：`ruff check .`、`mypy .`、`pytest`（20 passed）。
- 验收：版心/页眉页脚/未动段落保持一致；表上图下；两种边框 XML 正确；图片宽度 ≤ 版心 × ratio。

### 2026-10-23 · Patch: 表格内嵌大纲
- `iter_paragraphs()` 改为按文档顺序深度优先遍历正文、表格单元格、嵌套表格与内容控件。
- `heading_level()` 增加中文序号正则（`一、` / `一.`），无 Heading 样式的章节可被 `inspect` 识别。
- `insert_table_after()` 在单元格内插入表格时自动补尾随段落，满足 Word 对嵌套表格的约束。
- 新增 `tests/test_scanner.py`（嵌套锚点、中文大纲、单元格内插入、样式回退）。
- 验收：`inspect templates/sample_ca.docx` 输出四条大纲；`inject ... -o output_test.docx` 成功，Word 打开并导出 PDF 正常。

### 下一步（可选）
- 支持合并单元格（`gridSpan` / `vMerge`）与表注。
- 支持 Markdown / CSV 直接转 payload。
- Word COM / LibreOffice 渲染快照对比测试。
