# Specification: doc-injector (Technical & Academic Report Surgical Injector)

## 1. 目标与范围 (Goal & Scope)
针对大学工科实验报告（DSP、计算机组成原理、电路等）及学术论文规范，构建轻量、无损的 Word (`.docx`) 外科手术式注入工具：
- **零破坏性（Surgical Injection）**：100% 保持原始模板已有的版心边距、页眉页脚、字体层级与全局大纲结构。
- **学术规范对齐（Academic Standards）**：严格遵循国标规范——**“表上图下”**（表题在上、图题在下），同时支持**学术三线表**与**工科全边框网格表**。
- **多模态要素支持**：支持富文本段落（Run 级样式克隆与覆盖）、等比例居中自适应缩放插图、图表交叉编号。
- **降级容错（Fallback）**：当模板结构极度恶化或存在只读保护时，自动降级为结构化剪贴板导出（Clip-Helper）。

---

## 2. 系统架构与模块划分 (Architecture)

依托成熟开源生态，拒绝重复造轮子：
- **底层 OpenXML 处理**：`python-docx` + `docxtpl`
- **图片尺寸探测与预处理**：`Pillow` (PIL)
- **数据结构强契约校验**：`pydantic`
- **命令行工具链**：`click` / `typer`

```
doc-injector/
├── specs/
│   └── doc-injector-spec.md         # 规格说明书
├── templates/
│   └── sample_dsp.docx             # 测试样本模板
├── src/
│   ├── __init__.py
│   ├── schema.py                   # Pydantic 契约定义 (Block, Table, Image, Report)
│   ├── scanner.py                  # 模板大纲与语义锚点识别器
│   ├── table_engine.py             # 表格渲染器 (三线表 / 全边框网格表 + 表题在上)
│   ├── image_engine.py             # 图像缩放与居中渲染器 (图题在下)
│   ├── style_cloner.py             # Run-level 样式克隆与段落覆盖
│   ├── injector.py                 # 注入调度总控
│   └── cli.py                      # 命令行交互入口
└── tests/
    ├── test_table_engine.py
    └── test_injector.py
```

---

## 3. 核心学术图表规范与数据契约 (Data Contracts)

### 3.1 契约模型设计 (`src/schema.py`)

```python
from enum import Enum
from typing import List, Optional, Union
from pydantic import BaseModel, Field

class TableType(str, Enum):
    THREE_LINE = "three_line"  # 标准学术三线表 (顶底粗线、栏目细线、无竖线)
    GRID = "grid"              # 传统工科全边框网格表

class TableModel(BaseModel):
    title: str = Field(..., description="表题（必须放置在表格正上方，如：表1-1 采样测试数据）")
    table_type: TableType = TableType.THREE_LINE
    headers: List[str]
    rows: List[List[Union[str, float, int]]]
    alignment: str = "center"

class ImageModel(BaseModel):
    title: str = Field(..., description="图题（必须放置在图片正下方，如：图2-1 时域波形图）")
    path: str
    max_width_ratio: float = Field(default=0.8, description="占版心最大宽度比例")

class ContentBlock(BaseModel):
    anchor_keyword: str = Field(..., description="目标锚点关键词 (如: '实验目的', '波形分析')")
    paragraphs: Optional[List[str]] = None
    custom_font: Optional[str] = None
    custom_size_pt: Optional[float] = None
    table: Optional[TableModel] = None
    images: Optional[List[ImageModel]] = None

class ReportInjectionPayload(BaseModel):
    blocks: List[ContentBlock]
```

---

## 4. 关键注入算法与学术排版规则

### 4.1 表格规范 (Table Standard)
- **表题位置**：在目标锚点后插入居中段落，注入 `table.title`，设置段后间距为 3pt，段前间距为 6pt，样式设为小五号/五号（与模板正文协调）。
- **三线表渲染机制**：
  - 遍历 XML `w:tcBorders`，清除所有单元格竖线及中间水平线。
  - 顶线与底线：设置 `w:val="single" w:sz="12"` (1.5 磅)。
  - 表头下栏目线：设置 `w:val="single" w:sz="6"` (0.75 磅)。
- **全边框表格渲染机制**：
  - 全单元格设置 `w:val="single" w:sz="4"` (0.5 磅)。
- **防跨页断裂**：自动给所有表格行设置 `cantSplit = True`，表头行设置 `tblHeader = True`。

### 4.2 插图规范 (Image Standard)
- **版心约束自适应**：
  - 计算版心有效宽度：$\text{有效宽度} = \text{页面宽度} - (\text{左边距} + \text{右边距})$。
  - 读取原图分辨率，当宽度超过 $\text{有效宽度} \times \text{max\_width\_ratio}$ 时，自动等比缩小。
- **图题位置**：
  - 图片插入为独立居中段落。
  - 在图片下方紧随插入居中段落，写入 `image.title`，段前间距设为 3pt，段后间距设为 6pt。

### 4.3 文本段落 Run-Level 样式克隆
- 查找锚点段落下方的首个默认正文段落。
- 提取其首个 Run 的核心属性字典：`font.name`, `font.size`, `font.color.rgb`, `font.bold`, `paragraph.paragraph_format.line_spacing`。
- 清除原段落旧 Run，按提取的属性写入新文本；中文字符强制设置 `rPr.rFonts.ascii` 与 `hAnsi/eastAsia`，防止在不同机器上宋体降级为等线体。

---

## 5. 开发分工：Learn 模式 vs Fast 模式

### 5.1 @fast 模式交付物 (模板代码 / 胶水层)
1. `src/schema.py`：Pydantic 实体契约。
2. `src/scanner.py`：正则模糊匹配大纲标题与段落定位。
3. `src/cli.py`：基于 Click 的命令行封装。

### 5.2 @learn 模式手写交付物 (核心排版命门)
1. **`src/table_engine.py`**：
   - 手写 OpenXML 边框修改函数（构建三线表与全框网格表的底层 XML 节点）。
   - 实现表题（在上）与表格的绑定。
2. **`src/style_cloner.py`**：
   - 手写 Run 级样式复制函数，亲自体会 Word 字体中英文继承陷阱。

---

## 6. 验收标准 (Definition of Done)
1. **样式保真**：输入一份空白带标题的测试文档，注入后生成的新文档与原文档的页边距、页眉、未动段落 100% 一致。
2. **图表规范符合国标**：
   - 插入的表格表题在表格上方，三线表与全框表边框无误且不垮页。
   - 插入的波形图图题在图片正下方，无撑爆跳页。
3. **单元测试全绿**：`pytest tests/` 自动化断言 XML 节点生成无误。