# CorelDRAWer-Skill 🎨

> AI-powered geological diagram generation — describe what you want, get production-ready SVG vector graphics

[中文文档](README_zh.md) | [Skill Docs](.reasonix/skills/coreldraw-vba/SKILL.md)

## What is this?

A **Reasonix skill** that turns natural-language descriptions into:
- **SVG vector graphics** (default) — drag into CorelDRAW, Illustrator, or view in browser
- **CorelDRAW COM automation** (Windows) — directly control CorelDRAW to draw
- **VBA macro code** (legacy) — copy-paste into CorelDRAW's Alt+F11 editor

## Quick Start

### 🥇 Natural Language → SVG (Recommended)

Just describe what you need in the Reasonix chat:

> *"Draw a stratigraphic column for the Zigui area, Hubei — 14 layers from Nanhua to Ordovician"*  
> *"Generate a borehole log with 6 layers: 0-3m fill, 3-8m clay, 8-20m sandstone..."*

The AI will:
1. Parse your description into structured JSON data
2. Run `generate_column.py` to produce the SVG
3. Deliver a ready-to-use SVG file

**SVG files can be:**
- Dragged into **CorelDRAW** for editing (all vector elements preserved)
- Opened in **Illustrator / Inkscape**
- Viewed in any **browser**
- Embedded in Word / PowerPoint reports

### 🥈 Command Line

```bash
# Generate SVG
python3 generate_column.py data.json output.svg

# Generate VBA macro
python3 cdr_com_auto.py --vba data.json

# Windows: draw directly into CorelDRAW
python3 cdr_com_auto.py --com data.json
```

### 🥉 VBA Code (Legacy)

The AI can also generate VBA code. In CorelDRAW: Alt+F11 → Import → Run `DrawColumn`.

## v2.0 Features

### SVG Layer Groups
Output SVG is organized into 7 named layers — each editable independently in CorelDRAW:

| Layer ID | Content |
|----------|---------|
| `cdr-background` | White background |
| `cdr-title` | Title & subtitle |
| `cdr-header` | Column headers & dividers |
| `cdr-body` | All stratigraphic layers, scale ticks |
| `cdr-outlines` | Table borders |
| `cdr-footer` | Scale bar, summary info |
| `cdr-legend` | Lithology legend (right side) |

### CorelDRAW Metadata
Every SVG element carries `data-cdr-*` attributes for identification in CorelDRAW's XML editor:

- `data-cdr-layer` — parent layer group
- `data-cdr-type` — element type (`lithology-rect`, `fossil-icon`, `grain-curve`...)
- `data-cdr-name` — formation name
- `data-cdr-pattern` — pattern code
- `data-cdr-thickness` — layer thickness

### Adaptive Column Layout (11 columns)
Columns automatically show/hide based on available data:

| # | Column | Always? | Notes |
|---|--------|---------|-------|
| 1 | Erathem (界) | ✅ | Merged cells |
| 2 | System (系) | ✅ | Merged cells |
| 3 | Series (统) | ✅ | Merged cells |
| 4 | Formation (组) | ✅ | |
| 5 | Symbol (代号) | ✅ | |
| 6 | Fossils | ⚡ | Shows when fossil data present |
| 7 | Lithology Column | ✅ | 18 standard patterns |
| 8 | Grain Size | ✅ | Triangle + curve + label |
| 9 | Structures | ⚡ | Shows when structure data present |
| 10 | Thickness (m) | ✅ | |
| 11 | Description | ✅ | |

## 18 Standard Lithology Patterns (GB/T 958)

Based on China University of Geosciences (Wuhan) Zigui field base standards:

| Code | Rock Type | Pattern | Code | Rock Type | Pattern |
|------|-----------|---------|------|-----------|---------|
| `conglo` | Conglomerate | Circles + dots | `dolo` | Dolomite | Rhombic cross-hatch |
| `sand` | Sandstone | Dense dots | `doloLime` | Dolomitic limestone | Brick + rhombic |
| `finesand` | Fine sandstone | Fine dots | `chert` | Chert | Bold cross-hatch |
| `silt` | Siltstone | Lines + dots | `coal` | Coal | Black + white lines |
| `mud` | Mudstone | Dense lines | `granite` | Granite | Crosses + dots |
| `shale` | Shale | Lines + ticks | `basalt` | Basalt | V-pattern diagonal |
| `carbShale` | Carbonaceous shale | White dashes | `schist` | Schist | Wavy lines |
| `lime` | Limestone | Brick grid | `gneiss` | Gneiss | Thick/thin bands |
| `pure` | Solid color | No pattern | `marble` | Marble | Fine grid |

## Fossil Icons (15 types)

`trilobite` · `brachiopod` · `cephalopod` · `graptolite` · `crinoid` · `algae` · `stromatolite` · `ammonite` · `coral` · `bivalve` · `gastropod` · `foraminifera` · `plant` · `fish` · `spore`

## Structure Symbols (9 types)

`ripple` · `cross_bed` · `graded` · `ooid` · `crack` · `concretion` · `bioturbation` · `stylolite` · `stromatactis`

## JSON Data Format v2.0

```json
{
  "title": "Composite Stratigraphic Column",
  "location": "Zigui, Hubei",
  "layers": [
    {
      "erathem": "Neoproterozoic",
      "system": "Nanhua",
      "series": "Lower",
      "formation": "Liantuo Fm",
      "symbol": "Nh₁l",
      "thick": 120,
      "descr": "Purplish-red medium-bedded sandstone with basal conglomerate",
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

### Extended Fields

| Field | Type | Description |
|-------|------|-------------|
| `grain` | 1–6 | Grain size: 1=clay 2=silt 3=fine sand 4=medium sand 5=coarse sand 6=gravel |
| `grain_profile` | [[pos,level]...] | Continuous grain-size curve (pos 0–1 fraction within layer) |
| `fossils` | [string] | Fossil codes (see list above) |
| `structures` | [string] | Structure codes (see list above) |
| `contact` | string | `"conformity"` / `"disconformity"` / `"unconformity"` |
| `age_ma` | number | Stratigraphic age in millions of years |
| `markers` | [object] | Sample markers: `{symbol, y_offset, label}` |

## Example Output

Run immediately to generate the Zigui standard section:

```bash
python3 generate_column.py
# Output: output.svg (14 layers, ~1,780m total, auto-calculated scale)
```

## File Reference

| File | Purpose |
|------|---------|
| `generate_column.py` | SVG vector generator (zero dependencies, pure Python stdlib) |
| `cdr_com_auto.py` | CorelDRAW COM automation / VBA code generator |
| `data_template.json` | Stratigraphic data template — edit to switch regions |
| `borehole_column.bas` | Full VBA macro (legacy, 688 lines) |
| `output.svg` | Example SVG output (Zigui 14-layer standard section) |
| `.reasonix/skills/coreldraw-vba/SKILL.md` | Skill definition document |
| `README_zh.md` | Chinese documentation |

## Requirements

- **SVG generation**: Python 3.6+ (no dependencies)
- **CorelDRAW COM**: Windows + CorelDRAW X4+ + `pip install pywin32`
- **VBA**: CorelDRAW (any platform), Alt+F11 editor

## License

MIT
