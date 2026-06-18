# CorelDRAWer-Skill 🎨

> AI 驱动的地质图件生成 — 用自然语言描述，得到可直接用于 CorelDRAW 的 SVG 矢量图

[English Docs](README.md) | [技能文档](.reasonix/skills/coreldraw-vba/SKILL.md)

## 这是什么

一个 **Reasonix 技能**，能将自然语言描述直接输出为：
- **地层柱状图** — 11 列自适应布局、18 种国标花纹、15 种化石符号
- **地质剖面图** — 多钻孔地下横剖面、含断层标注
- **SVG 矢量图**（默认）—— v2.1 含命名图层分组 + CorelDRAW 元数据属性
- **CorelDRAW COM 自动化**（Windows）—— 直接操控 CorelDRAW 画图
- **VBA 宏代码**（传统）—— 复制到 CorelDRAW Alt+F11 编辑器中运行

## 快速开始

### 🥇 自然语言 → SVG（推荐）

在 Reasonix 中直接描述：

> *"画一张秭归地区综合地层柱状图，14 层，从南华系到奥陶系"*
> *"生成一个钻孔柱状图：0-3m 填土、3-8m 粘土、8-20m 砂岩..."*

AI 会自动：
1. 理解需求，生成结构化 JSON 数据
2. 运行 `generate_column.py` 输出 SVG
3. 交付可直接使用的矢量图文件

**SVG 可以：**
- 拖入 **CorelDRAW** 编辑（7 层分组保留，`data-cdr-*` 属性可识别）
- 拖入 **Illustrator / Inkscape** 编辑
- 浏览器直接查看
- 嵌入 Word / PPT 报告

### 🥈 命令行

```bash
# 生成 SVG
python3 generate_column.py data.json output.svg

# 生成 VBA 宏
python3 cdr_com_auto.py --vba data.json

# Windows: 直接画到 CorelDRAW 中
python3 cdr_com_auto.py --com data.json
```

### 🥉 VBA 代码

CorelDRAW → Alt+F11 → 导入 → 运行 `DrawColumn`。

## 地质剖面图 (v1.0)

从多个钻孔数据生成地下横剖面：

```bash
python3 generate_cross_section.py data.json cross_section.svg
```

功能：多钻孔地表线 + 地层对比连线 + 断层标注 + 双比例尺 + 18 种国标花纹填充。

## v2.0 新特性

### SVG 图层分组
输出 SVG 按功能分 7 层，导入 CorelDRAW 后每层独立可编辑：

| 图层 ID | 内容 |
|---------|------|
| `cdr-background` | 白色背景 |
| `cdr-title` | 图名 + 地点 |
| `cdr-header` | 表头 + 分隔线 |
| `cdr-body` | 所有地层图层 + 刻度尺 |
| `cdr-outlines` | 四边框线 |
| `cdr-footer` | 比例尺 + 汇总信息 |
| `cdr-legend` | 右侧图例框 |

### CorelDRAW 元数据属性
每个 SVG 元素携带 `data-cdr-*` 属性，可在 CorelDRAW XML 编辑器中识别：
- `data-cdr-layer` — 所属图层
- `data-cdr-type` — 元素类型（`lithology-rect`、`fossil-icon`、`grain-curve`...）
- `data-cdr-name` — 地层名称
- `data-cdr-pattern` — 花纹代码
- `data-cdr-thickness` — 地层厚度

### 自适应列布局（11 列）
根据数据自动显示/隐藏：

| # | 列名 | 显示条件 | 说明 |
|---|------|----------|------|
| 1 | 界 | 始终 | 合并单元格 |
| 2 | 系 | 始终 | 合并单元格 |
| 3 | 统 | 始终 | 合并单元格 |
| 4 | 组 | 始终 | |
| 5 | 代号 | 始终 | |
| 6 | 化石 | ⚡ 有数据时 | 15 种化石符号 |
| 7 | 岩性柱 | 始终 | 18 种国标花纹 |
| 8 | 粒度 | 始终 | 三角 + 曲线 + 中文标签 |
| 9 | 构造 | ⚡ 有数据时 | 9 种构造符号 |
| 10 | 厚度(m) | 始终 | |
| 11 | 岩性描述 | 始终 | |

## 18 种国标花纹 (GB/T 958)

依据中国地质大学（武汉）秭归产学研基地标准：

| 代码 | 岩性 | 花纹 | 代码 | 岩性 | 花纹 |
|------|------|------|------|------|------|
| `conglo` | 砾岩 | 大圆+散点 | `dolo` | 白云岩 | 菱格交叉线 |
| `sand` | 砂岩 | 密点 | `doloLime` | 白云质灰岩 | 砖格+菱格 |
| `finesand` | 细砂岩 | 细密点 | `chert` | 硅质岩 | 粗交叉线 |
| `silt` | 粉砂岩 | 横线+点 | `coal` | 煤 | 全黑+白线 |
| `mud` | 泥岩 | 密横线 | `granite` | 花岗岩 | 叉+点 |
| `shale` | 页岩 | 横线+短竖 | `basalt` | 玄武岩 | V 形斜排 |
| `carbShale` | 炭质页岩 | 黑底白线 | `schist` | 片岩 | 波浪折线 |
| `lime` | 石灰岩 | 错缝砖格 | `gneiss` | 片麻岩 | 粗细条带 |
| `pure` | 纯色 | 无花纹 | `marble` | 大理岩 | 细网格 |

## 化石符号（15 种）

`trilobite` 三叶虫 · `brachiopod` 腕足 · `cephalopod` 头足 · `graptolite` 笔石 · `crinoid` 海百合 · `algae` 藻类 · `stromatolite` 叠层石 · `ammonite` 菊石 · `coral` 珊瑚 · `bivalve` 双壳 · `gastropod` 腹足 · `foraminifera` 有孔虫 · `plant` 植物 · `fish` 鱼类 · `spore` 孢粉

## 构造符号（9 种）

`ripple` 波痕 · `cross_bed` 交错层理 · `graded` 粒序层理 · `ooid` 鲕粒 · `crack` 泥裂 · `concretion` 结核 · `bioturbation` 生物扰动 · `stylolite` 缝合线 · `stromatactis` 层孔

## JSON 数据格式 v2.0

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
      "descr": "紫红色中厚层砂岩、含砾砂岩，底部为砾岩",
      "c": 0, "m": 40, "y": 30, "k": 10,
      "pattern": "sand",
      "grain": 4,
      "fossils": [],
      "structures": [],
      "contact": "unconformity",
      "age_ma": 780,
      "grain_profile": [[0.0, 5], [0.3, 4], [0.7, 3], [1.0, 4]],
      "markers": [{"symbol": "star", "y_offset": 0.5, "label": "S1"}]
    }
  ]
}
```

### 扩展字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| `grain` | 1–6 | 粒度：1粘土 2粉砂 3细砂 4中砂 5粗砂 6砾石 |
| `grain_profile` | [[位置,级别]...] | 连续粒度曲线（位置 0–1 层内比例） |
| `fossils` | [字符串] | 化石代码列表（见上） |
| `structures` | [字符串] | 构造代码列表（见上） |
| `contact` | 字符串 | `"conformity"` / `"disconformity"` / `"unconformity"` |
| `age_ma` | 数字 | 地层年代（百万年） |
| `markers` | [对象] | 采样标记：`{symbol, y_offset, label}` |

## 示例输出

```bash
python3 generate_column.py
# 输出: output.svg（14 层，总厚 ~1,780m，比例尺自动计算）
```

## 文件说明

| 文件 | 用途 |
|------|------|
| `generate_column.py` | SVG 矢量图生成器 — 地层柱状图（零依赖） |
| `generate_cross_section.py` | SVG 矢量图生成器 — 地质剖面图（零依赖） |
| `cdr_com_auto.py` | CorelDRAW COM 自动化 / VBA 代码生成 v2.0 |
| `data_template.json` | 地层柱状图数据模板 v2.0 |
| `borehole_column.bas` | 完整 VBA 宏（旧版参考，688 行） |
| `output.svg` | 秭归 14 层标准剖面示例 SVG |
| `cross_section_demo.svg` | 3 钻孔剖面图示例（含断层） |
| `.reasonix/skills/coreldraw-vba/SKILL.md` | 技能定义文件 |
| `README.md` | 英文文档 |

## 运行要求

- **SVG 生成**：Python 3.6+（零依赖）
- **CorelDRAW COM**：Windows + CorelDRAW X4+ + `pip install pywin32`
- **VBA**：CorelDRAW（任意平台），Alt+F11 编辑器

## License

MIT

## 引用

使用 CorelDRAWer-Skill 请引用：

```bibtex
@software{coreldrawer_skill,
  author = {Reasonite-lab},
  title = {CorelDRAWer-Skill: AI-powered geological diagram generation},
  year = {2025},
  url = {https://github.com/Reasonite-lab/CorelDRAWer-Skill},
  version = {0.2.1}
}
```

或使用 [CITATION.cff](CITATION.cff) 导入 Zotero / Endnote。
