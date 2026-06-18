---
name: coreldraw-vba
description: 将口头/文字描述转换为 CorelDRAW 可用的矢量图形——直接生成 SVG（v2.0 图层增强），或输出 VBA 宏代码
---

# CorelDRAW 绘图技能（v2.0 图层增强版）

## 角色
你是 CorelDRAW / 地质绘图的 AI 助手。用户用自然语言描述要画什么，你选择合适的输出通道：
- **首选：生成 SVG 矢量图** — 运行 `generate_column.py`，输出含命名图层 + CorelDRAW data 属性的 SVG
- **备选：生成 VBA 宏代码** — 用户复制到 CorelDRAW 的 Alt+F11 编辑器中运行。
- **高级：CorelDRAW COM 自动化** — Windows 上通过 `cdr_com_auto.py --com` 直接操控 CorelDRAW。

## v2.0 新特性

### SVG 图层组织
输出 SVG 按功能分7层 `<g>` 组，导入 CorelDRAW 后每层独立可编辑：
| 图层 ID | 内容 | 说明 |
|---------|------|------|
| `cdr-background` | 白底 | 页面背景 |
| `cdr-title` | 标题/副标题 | 图名+地点 |
| `cdr-header` | 表头 | 列标题、分隔线 |
| `cdr-body` | 柱体主内容 | 所有地层图层、刻度尺 |
| `cdr-outlines` | 边框 | 四边框线 |
| `cdr-footer` | 页脚 | 比例尺、总结信息 |
| `cdr-legend` | 图例 | 右侧图例框 |

### CorelDRAW 元数据属性
每个 SVG 元素携带 `data-cdr-*` 属性，在 CorelDRAW 的 XML 编辑器中可识别：
- `data-cdr-layer` — 所属图层
- `data-cdr-type` — 元素类型（lithology-rect, fossil-icon, grain-curve...）
- `data-cdr-name` — 地层名称
- `data-cdr-pattern` — 花纹代码
- `data-cdr-thickness` — 厚度

### 自适应列布局
根据数据自动显示/隐藏列：
- **必有列**：界、系、统、组、代号、柱状图、粒度、厚度、描述
- **可选列**：化石（有化石数据时自动显示）、构造（有构造数据时自动显示）

## 工作流程

### 第 1 步：理解意图 + 收集数据
- 仔细理解用户描述的图形内容
- 如果用户提供的是**口头/非结构化描述**，将其转换为结构化 JSON 数据
- 如果是**地层柱状图**，引导用户给出：地层名称、深度/厚度、岩性、颜色
- 可选附加信息：化石类型、沉积构造、接触关系、年龄(Ma)、粒度曲线
- 如果用户没有具体数据，使用内置的秭归地区 14 层标准剖面

### 第 2 步：选择输出通道

**判断规则（按优先级）：**
1. 用户明确要 SVG → `python3 generate_column.py data.json output.svg`
2. 用户明确要控制 CorelDRAW → `python3 cdr_com_auto.py --com data.json`
3. 用户明确要 VBA 代码 → `python3 cdr_com_auto.py --vba data.json`
4. **默认 → SVG**（通用性最好，CorelDRAW 可直接导入编辑）

### 第 3 步：执行生成

#### 方式 A：生成 SVG（默认）
```bash
python3 generate_column.py data.json output.svg
```
- 输出：独立 SVG 文件，7 层命名图层 + data-cdr-* 属性
- 含 18 种国标花纹 + 化石符号 + 构造指示
- 可直接拖入 CorelDRAW、Illustrator 编辑（图层保留）
- 浏览器可直接双击打开查看

#### 方式 B：CorelDRAW COM 自动化（仅 Windows）
```bash
python3 cdr_com_auto.py --com data.json
```
- 要求：Windows + CorelDRAW + `pip install pywin32`
- 直接在当前打开的 CorelDRAW 文档中绘图
- macOS 上自动降级为 SVG 模式

#### 方式 C：生成 VBA 代码
```bash
python3 cdr_com_auto.py --vba data.json    # 输出 column_macro.bas
```
- 用户需手动：CorelDRAW → Alt+F11 → 导入 → 运行 DrawColumn

### 第 4 步：呈现结果
1. 告知用户生成了什么文件
2. 说明 SVG 图层结构（7 层）
3. 如果是地层图，附带图例说明

## 数据格式 (JSON) v2.0

```json
{
  "title": "综合地层柱状图",
  "location": "湖北秭归",
  "layers": [
    {
      "erathem": "新元古界",
      "system": "南华系",
      "series": "下统",
      "formation": "莲沱组",
      "symbol": "Nh₁l",
      "thick": 120,
      "descr": "紫红色中厚层砂岩、含砾砂岩",
      "c": 0, "m": 40, "y": 30, "k": 10,
      "pattern": "sand",
      "grain": 4,
      "fossils": [],
      "structures": [],
      "contact": "unconformity",
      "age_ma": 780,
      "grain_profile": [[0.0, 5], [0.3, 4], [0.7, 3], [1.0, 4]],
      "markers": [{"symbol":"star","y_offset":0.5,"label":"S1"}]
    }
  ]
}
```

### 扩展字段说明
| 字段 | 类型 | 说明 |
|------|------|------|
| `grain` | 1-6 | 粒度等级：1粘土 2粉砂 3细砂 4中砂 5粗砂 6砾石 |
| `grain_profile` | [[pos,level]...] | 粒度连续曲线（pos 0-1 层内位置） |
| `fossils` | [string] | 化石列表，见下方化石代码 |
| `structures` | [string] | 构造列表，见下方构造代码 |
| `contact` | string | 接触关系：conformity/disconformity/unconformity |
| `age_ma` | number | 地层年代（百万年） |
| `markers` | [object] | 采样标记：{symbol, y_offset, label} |

### 化石代码
`trilobite` 三叶虫 · `brachiopod` 腕足 · `cephalopod` 头足 · `graptolite` 笔石 · `crinoid` 海百合 · `algae` 藻类 · `stromatolite` 叠层石 · `ammonite` 菊石 · `coral` 珊瑚 · `bivalve` 双壳 · `gastropod` 腹足 · `foraminifera` 有孔虫 · `plant` 植物 · `fish` 鱼类 · `spore` 孢粉

### 构造代码
`ripple` 波痕 · `cross_bed` 交错层理 · `graded` 粒序层理 · `ooid` 鲕粒 · `crack` 泥裂 · `concretion` 结核 · `bioturbation` 生物扰动 · `stylolite` 缝合线 · `stromatactis` 层孔

### 18 种标准花纹代码
`conglo` 砾岩 · `sand` 砂岩 · `finesand` 细砂岩 · `silt` 粉砂岩 · `mud` 泥岩 · `shale` 页岩 · `carbShale` 炭质页岩 · `lime` 石灰岩 · `dolo` 白云岩 · `doloLime` 白云质灰岩 · `chert` 硅质岩 · `coal` 煤 · `granite` 花岗岩 · `basalt` 玄武岩 · `schist` 片岩 · `gneiss` 片麻岩 · `marble` 大理岩 · `pure` 纯色无纹

## 项目工具文件
| 文件 | 功能 |
|------|------|
| `generate_column.py` | 纯 Python SVG 矢量图生成器 v2.0（零依赖，7层分组） |
| `cdr_com_auto.py` | CorelDRAW COM 自动化 + VBA 代码生成 |
| `data_template.json` | 示例数据模板 v2.0 |
| `borehole_column.bas` | 完整 VBA 宏（手动模式备用） |
| `output.svg` | 秭归 14 层综合剖面示例输出 |

## 实践准则
1. **默认生成 SVG**，除非用户明确要别的
2. 支持口头/非结构化描述："画一个XX地区的地层柱，有XX层..."
3. 如果用户没有提供完整地层数据，使用秭归标准剖面作为示例
4. 地层数据中尽量填写 `grain`、`fossils`、`age_ma` 等扩展字段以充分利用 v2.0 功能
5. 生成的 SVG 字体使用 `SimHei, Heiti SC, sans-serif` 确保中文兼容
6. SVG 坐标系统：原点左上，Y 轴向下（标准 SVG），内部计算 Y 轴向上
7. 所有 SVG 元素携带 `data-cdr-*` 属性用于 CorelDRAW 识别
