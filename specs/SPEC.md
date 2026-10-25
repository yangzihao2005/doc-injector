# SPEC.md

## 目标
实现 `doc-injector`：对 Word (`.docx`) 实验报告模板进行无损、锚点驱动的段落 / 学术表格 / 插图注入。

## 范围
- 契约：Pydantic 校验 payload（`src/schema.py`）。
- 定位：标题优先、正文次之、标点容错正则兜底；支持表格单元格与嵌套表格内的锚点（`src/scanner.py`）。
- 排版：表上图下、三线表 / 全框网格表、防跨页断行、表头重复（`src/table_engine.py`）。
- 图像：按版心比例等比缩放、居中、图题在下（`src/image_engine.py`）。
- 文本：Run 级样式克隆 + 中文字体槽位修复（`src/style_cloner.py`）。
- 调度：原子保存、只读/占用/损坏模板降级为结构化文本 + 剪贴板（`src/injector.py`）。
- 交互：`inspect` / `inject` / `demo` / `schema`（`src/cli.py`）。

完整规格见 [specs/doc-injector-spec.md](specs/doc-injector-spec.md)。

## 验收标准
- [x] 注入前后页边距、页眉页脚、未动段落与原文档一致。
- [x] 表题在表格上方；三线表顶底 1.5 磅、栏目线 0.75 磅、无竖线；网格表 0.5 磅全边框。
- [x] 所有表格行 `cantSplit`，表头行 `tblHeader`。
- [x] 图片等比缩放且不超过版心宽度 × `max_width_ratio`，图题位于图片下方。
- [x] 中文字体显式写入 `w:eastAsia`，不依赖机器默认字体。
- [x] 缺锚点可报错或跳过；保存失败自动降级导出。
- [x] 中文序号章节（`一、实验目的`）无需 Heading 样式即可进入大纲。
- [x] 表格单元格 / 嵌套表格内的锚点可定位并安全注入。
- [x] `ruff check .`、`mypy .`、`pytest` 全部通过（20 tests）。
