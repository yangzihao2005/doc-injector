# doc-injector

> **针对高校工科实验报告与学术制式文档的外科手术式无损注入引擎。**  
> 告别全局格式被破坏、字体降级错位与反复手动微调，严格践行国标学术图表规范（表上图下、标准三线表/网格表）。

---

## 核心特性 (Features)

* **外科手术式注入（Zero Destruction）**：完全保留原模板的封面、页眉页脚、版心边距、正文字体（宋体/Times New Roman 等）与行间距，绝不触发全量重写。
* **原生支持表格嵌套模板**：深度支持高校常见的“全文档由外层大表格包裹”模板，精准识别嵌套在单元格内的各级大纲锚点。
* **学术规范严格对齐**：
  * **表上图下**：表题居中置于表格正上方；图题居中置于插图正下方。
  * **双表格引擎**：支持顶底粗线的标准学术**三线表**与工程实测**全框网格表**，内置跨页断裂保护（`cantSplit` / `tblHeader`）。
* **版心智能自适应插图**：自动根据页面正文有效宽度进行等比例安全缩放，拒绝溢出与跨页撑爆。
* **强契约数据驱动**：基于 Pydantic 校验，输入数据即文档，杜绝多字段漏字段与语法漂移。

---

## 环境准备与安装 (Installation)

### 1. 克隆工程并进入目录
```bash
git clone <your-repo-url> doc-injector
cd doc-injector
```

### 2. 创建并激活虚拟环境（推荐）
```bash
# Windows PowerShell
python -m venv .venv
.venv\Scripts\Activate.ps1

# Linux / macOS
python3 -m venv .venv
source .venv/bin/activate
```

### 3. 安装依赖
```bash
pip install click python-docx docxtpl pydantic pillow
```

---

## 快速使用指南 (Quick Start)

整个使用工作流包含 **三步**：探查锚点 $\to$ 准备数据 $\to$ 一键注入。

### 第一步：探查模板大纲 (Inspect)

在注入前，运行探查命令确认程序能正确识别模板中的章节锚点：

```bash
python -m src.cli inspect templates/sample_ca.docx
```

输出示例：
```text
[0] 一、实验目的 (Heading 1)
[1] 二、实验内容 (Heading 1)
[2] 三、实验步骤 (Heading 1)
[3] 四、实验结果与分析 (Heading 1)
```

---

### 第二步：准备注入数据 (`data.json`)

编写或让大模型生成符合契约的 `data.json`。支持注入纯段落、三线表/网格表以及自适应插图：

```json
{
  "blocks": [
    {
      "anchor_keyword": "实验目的",
      "paragraphs": [
        "1. 掌握算术逻辑单元 (ALU) 的基本构成与控制信号逻辑。",
        "2. 验证加减法运算溢出判别机制及状态标志位生成规律。"
      ]
    },
    {
      "anchor_keyword": "实验步骤",
      "paragraphs": [
        "1. 设计硬件线路：将开关与运算器数据输入端连接。",
        "2. 编写微指令并单步运行调试。"
      ],
      "images": [
        {
          "title": "图3-1 实验硬件连接示意图",
          "path": "assets/circuit.png",
          "max_width_ratio": 0.8
        }
      ]
    },
    {
      "anchor_keyword": "实验结果与分析",
      "paragraphs": [
        "通过单步运行微程序并观察波形，实测输出与理论真值表完全吻合。"
      ],
      "table": {
        "title": "表4-1 运算器测试数据对照表",
        "table_type": "three_line",
        "headers": ["控制信号", "输入A", "输入B", "输出F", "OF标志"],
        "rows": [
          ["ADD", "0x01", "0x02", "0x03", "0"],
          ["SUB", "0x05", "0x01", "0x04", "0"]
        ]
      },
      "images": [
        {
          "title": "图4-1 仿真输出时序波形图",
          "path": "assets/waveform.png",
          "max_width_ratio": 0.85
        }
      ]
    }
  ]
}
```

> **提示**：如果需要全框网格表，只需将 `"table_type"` 改为 `"grid"`。

---

### 第三步：执行注入 (Inject)

在终端运行 `inject` 命令：

```bash
python -m src.cli inject templates/sample_ca.docx data.json -o output_final.docx
```

命令成功完成后，打开 `output_final.docx` 即可直接提交！

---

## 命令行参数一览 (CLI Reference)

```bash
# 查看全局帮助
python -m src.cli --help

# 1. 探查模板大纲与候选锚点
python -m src.cli inspect <TEMPLATE_PATH>

# 2. 执行数据注入
python -m src.cli inject <TEMPLATE_PATH> <PAYLOAD_JSON> -o <OUTPUT_PATH>

# 3. 查看 JSON Schema 定义 (用于编辑器自动补全或 Prompt 约束)
python -m src.cli schema
```

---

## 目录结构说明 (Project Structure)

```text
doc-injector/
├── specs/
│   ├── doc-injector-spec.md            # 项目主设计规范
│   └── doc-injector-patch-spec.md      # 表格嵌套与大纲增强补丁规范
├── templates/
│   └── sample_ca.docx                  # 原始实验报告模板
├── assets/                             # 仿真波形、截图等图片素材
├── src/
│   ├── __init__.py
│   ├── cli.py                          # Click 命令行交互入口
│   ├── schema.py                       # Pydantic 数据模型与校验
│   ├── scanner.py                      # 模板大纲与嵌套表格段落扫描器
│   ├── injector.py                     # 外科手术式注入主调度
│   ├── style_cloner.py                 # Run 级文字样式与中文字体克隆
│   ├── table_engine.py                 # 三线表与全框网格表渲染引擎
│   └── image_engine.py                 # 图片等比缩放与居中图题渲染
├── tests/
│   ├── test_scanner.py
│   └── test_injector.py
├── data.json                           # 待注入的实验结构化数据
└── README.md
```

---

## 运行测试 (Testing)

运行完整的自动化测试套件：

```bash
pytest tests/
```