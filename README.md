# CorelDRAWer-Skill 🎨

> 用自然语言生成 CorelDRAW 地质图 — AI 驱动的矢量绘图技能

## 这是什么

一个 **Reasonix 技能**，能将口头/文字描述直接输出为：
- **SVG 矢量图**（默认）—— 拖入 CorelDRAW 即可编辑，浏览器也能看
- **CorelDRAW COM 自动化**（Windows）—— 直接操控 CorelDRAW 画图
- **VBA 宏代码**（传统）—— 复制到 CorelDRAW 编辑器中运行

## 三种使用方式

### 🥇 方式一：自然对话 → SVG（推荐）

直接在 Reasonix 中说：

> "画一张秭归地区综合地层柱状图"
> "画一个钻孔柱，包含 6 层：0-3m 填土、3-8m 粘土、8-20m 砂岩..."

AI 会自动：
1. 理解你的需求，生成结构化数据
2. 运行 `generate_column.py` 生成 SVG
3. 输出可直接打开的 SVG 文件

**SVG 文件可以：**
- 拖入 **CorelDRAW** 编辑（完美导入，所有矢量元素保留）
- 拖入 **Illustrator / Inkscape** 编辑
- 双击在**浏览器**中查看
- 嵌入 Word / PPT 报告

### 🥈 方式二：命令行调用

```bash
# 生成 SVG
python3 generate_column.py data.json output.svg

# 生成 VBA 代码
python3 cdr_com_auto.py --vba data.json

# Windows 上直接画到 CorelDRAW
python3 cdr_com_auto.py --com data.json
```

### 🥉 方式三：Copy VBA 代码

AI 也可以生成 VBA 代码（老方式），用户在 CorelDRAW 中 Alt+F11 → 粘贴 → 运行。

## 文件说明

| 文件 | 用途 |
|------|------|
| `generate_column.py` | SVG 矢量图生成器（零依赖，纯 Python 标准库）|
| `cdr_com_auto.py` | CorelDRAW COM 自动控制 / VBA 代码生成 |
| `data_template.json` | 地层数据模板，改这个就能换地区 |
| `borehole_column.bas` | 完整 VBA 宏（备用，688 行）|
| `output.svg` | 生成的示例 SVG（秭归地区标准剖面）|
| `.reasonix/skills/coreldraw-vba/SKILL.md` | 技能定义文件 |

## 18 种国家标准花纹（GB/T 958）

依据中国地质大学（武汉）秭归产学研基地《常用地层图例花纹和符号》：

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

## 示例效果

运行以下命令立即生成秭归地区标准剖面：

```bash
python3 generate_column.py
# 输出: output.svg（14 层，总厚 ~1780m，比例尺自动计算）
```

## JSON 数据格式

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
      "descr": "紫红色中厚层砂岩",
      "c": 0, "m": 40, "y": 30, "k": 10,
      "pattern": "sand"
    }
  ]
}
```
