# Specification Patch: Table-Nested Outline & Content Injection Support
File Target: `specs/doc-injector-patch-spec.md`
Affected Modules: `src/scanner.py`, `tests/test_scanner.py`

---

## 1. 缺陷背景与动机 (Problem Statement)
- **痛点**：高校工科实验报告模板（如计组、电路）通常使用外层边框大表格包裹全文。原实现中 `iter_paragraphs()` 显式跳过了表格内容（`tables are intentionally skipped`），导致：
  1. 位于单元格内部的大纲标题（如“一、实验目的”）无法被识别为锚点，触发 `Error: 未找到锚点`。
  2. 模板大纲扫描命令 `inspect` 输出为空。
- **目标**：在不破坏非表格模板兼容性的前提下，实现对正文段落与嵌套表格单元格段落的全覆盖扫描与模糊识别。

---

## 2. 详细变更需求 (Detailed Requirements)

### 2.1 遍历范围扩展 (`src/scanner.py`)
- 重构 `iter_paragraphs(doc: DocumentObject) -> Iterator[Paragraph]`：
  1. 保持正文顶层 `doc.paragraphs` 的生成顺序。
  2. 新增对 `doc.tables` 内部所有单元格段落（`cell.paragraphs`）的深度优先遍历。
  3. 保留对每个段落原生所属容器（`Paragraph._p`）的引用，以确保后续 `insert_paragraph_after()` 与 `insert_table_after()` 的底层 XML `addnext` 操作在单元格内仍能正常工作。

### 2.2 大纲识别兼容性增强 (`scan_outline`)
- 很多实验报告未采用标准 `Heading` 样式（通常标记为 `Normal` 或“正文”，但带有中文数字序号）。
- 增加中文层级标题正则判定规则：
  - 正则表达式：`r"^[一二三四五六七八九十]+[、.．]"`
  - 当段落不具备标准 `Heading` 样式但命中该正则时，自动将其 `level` 标记为 `1`。

### 2.3 单元格内插入元素容错机制 (`insert_table_after`)
- 需确保当锚点段落位于表格单元格（`_Cell`）内时，`addnext` 能正确将其附加在当前段落后，避免触发 OpenXML 架构校验断言崩溃。

---

## 3. 验收标准与测试用例 (Acceptance Criteria & Tests)

### 3.1 单元测试 (`tests/test_scanner.py`)
1. **测试用例 1：`test_scan_nested_table_paragraphs`**
   - 构造一个包含表格的 `.docx`，并在单元格内放置 `一、实验目的`。
   - 断言 `find_anchor(doc, "实验目的")` 能够成功返回该段落对象，不再为 `None`。
2. **测试用例 2：`test_inspect_outline_detects_cn_numbers`**
   - 验证无大纲样式的中文序号段落（如 `四、实验结果与分析`）能被 `scan_outline` 正确识别为 `level=1` 的大纲条目。

### 3.2 端到端验收
- 重新执行命令：
  ```bash
  python -m src.cli inspect templates/sample_ca.docx
  ```
  预期控制台能够成功打印出包含 `一、实验目的`、`二、实验内容`、`三、实验步骤`、`四、实验结果与分析` 的大纲列表。
- 重新执行注入：
  ```bash
  python -m src.cli inject templates/sample_ca.docx data.json -o output_test.docx
  ```
  预期成功生成 `output_test.docx`，不再报找不到锚点错误。