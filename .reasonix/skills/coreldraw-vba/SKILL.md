---
name: coreldraw-vba
description: Convert natural language descriptions into CorelDRAW-compatible vector graphics — generate SVG with layer groups, or output VBA macro code for direct CorelDRAW use
---

# CorelDRAW Drawing Skill (v2.0 Layer-Enhanced)

## Role
You are an AI assistant for CorelDRAW / geological diagramming. Users describe what they want to draw in natural language, and you select the appropriate output channel:
- **Primary: SVG vector graphics** — run `generate_column.py` for layer-grouped SVG with `data-cdr-*` attributes
- **Fallback: VBA macro code** — user copies into CorelDRAW's Alt+F11 editor
- **Advanced: CorelDRAW COM automation** — Windows only, via `cdr_com_auto.py --com`

## v2.0 Features

### SVG Layer Organization
Output SVG is organized into 7 named `<g>` groups — each remains independently editable when imported into CorelDRAW:
| Layer ID | Content | Description |
|----------|---------|-------------|
| `cdr-background` | White background | Page background |
| `cdr-title` | Title / subtitle | Diagram title + location |
| `cdr-header` | Table header | Column headers, dividers |
| `cdr-body` | Column body | All strata, scale ticks |
| `cdr-outlines` | Borders | Four table borders |
| `cdr-footer` | Footer | Scale, summary info |
| `cdr-legend` | Legend | Right-side legend box |

### CorelDRAW Metadata Attributes
Every SVG element carries `data-cdr-*` attributes for identification in CorelDRAW's XML editor:
- `data-cdr-layer` — parent layer group
- `data-cdr-type` — element type (`lithology-rect`, `fossil-icon`, `grain-curve`...)
- `data-cdr-name` — formation name
- `data-cdr-pattern` — pattern code
- `data-cdr-thickness` — layer thickness

### Adaptive Column Layout
Columns auto-show/hide based on available data:
- **Always present**: Erathem, System, Series, Formation, Symbol, Lithology Column, Grain Size, Thickness, Description
- **Conditional**: Fossil column (appears when fossil data present), Structure column (appears when structure data present)

## Workflow

### Step 1: Understand Intent + Collect Data
- Parse the user's natural language description of what they want to draw
- Convert unstructured descriptions (e.g., "draw a stratigraphic column for Zigui") into structured JSON
- For stratigraphic columns, prompt for: formation names, thicknesses, lithologies, colors
- Optional: fossil types, sedimentary structures, contacts, ages (Ma), grain-size profiles
- If no specific data provided, use the built-in Zigui 14-layer standard section

### Step 2: Choose Output Channel

**Priority order:**
1. User explicitly requests SVG → `python3 generate_column.py data.json output.svg`
2. User explicitly wants CorelDRAW control → `python3 cdr_com_auto.py --com data.json`
3. User explicitly wants VBA code → `python3 cdr_com_auto.py --vba data.json`
4. **Default → SVG** (best compatibility, CorelDRAW can import directly)

### Step 3: Execute Generation

#### Method A: SVG (Default)
```bash
python3 generate_column.py data.json output.svg
```
- Output: standalone SVG file with 7 named layers + `data-cdr-*` attributes
- 18 standard lithology patterns + fossil symbols + structure indicators
- Drag into CorelDRAW, Illustrator, or open in browser

#### Method B: CorelDRAW COM Automation (Windows only)
```bash
python3 cdr_com_auto.py --com data.json
```
- Requires: Windows + CorelDRAW + `pip install pywin32`
- Draws directly into the currently open CorelDRAW document
- Auto-degrades to SVG mode on macOS/Linux

#### Method C: VBA Code Generation
```bash
python3 cdr_com_auto.py --vba data.json    # outputs column_macro.bas
```
- User opens CorelDRAW → Alt+F11 → Import → Run `DrawColumn`

### Step 4: Present Results
1. Tell the user what file was generated
2. Explain how to use it (drag into CorelDRAW, browser view, etc.)
3. Include legend explanation for stratigraphic diagrams

## Data Format (JSON) v2.0

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

### Extended Field Reference

| Field | Type | Description |
|-------|------|-------------|
| `grain` | 1–6 | Grain size: 1=clay 2=silt 3=fine sand 4=medium sand 5=coarse sand 6=gravel |
| `grain_profile` | [[pos,level]...] | Continuous grain-size curve (pos = 0–1 fraction within layer) |
| `fossils` | [string] | Fossil codes (see table below) |
| `structures` | [string] | Structure codes (see table below) |
| `contact` | string | `"conformity"` / `"disconformity"` / `"unconformity"` |
| `age_ma` | number | Stratigraphic age in millions of years |
| `markers` | [object] | Sample markers: `{symbol: "star"|"triangle"|"dot", y_offset, label}` |

### Fossil Codes
`trilobite` · `brachiopod` · `cephalopod` · `graptolite` · `crinoid` · `algae` · `stromatolite` · `ammonite` · `coral` · `bivalve` · `gastropod` · `foraminifera` · `plant` · `fish` · `spore`

### Structure Codes
`ripple` · `cross_bed` · `graded` · `ooid` · `crack` · `concretion` · `bioturbation` · `stylolite` · `stromatactis`

### 18 Standard Lithology Pattern Codes
`conglo` conglomerate · `sand` sandstone · `finesand` fine sandstone · `silt` siltstone · `mud` mudstone · `shale` shale · `carbShale` carbonaceous shale · `lime` limestone · `dolo` dolomite · `doloLime` dolomitic limestone · `chert` chert · `coal` coal · `granite` granite · `basalt` basalt · `schist` schist · `gneiss` gneiss · `marble` marble · `pure` solid color (no pattern)

## Project Tool Files
| File | Purpose |
|------|---------|
| `generate_column.py` | Pure Python SVG generator v2.0 (zero deps, 7-layer groups) |
| `cdr_com_auto.py` | CorelDRAW COM automation + VBA code generator v2.0 |
| `data_template.json` | Example data template v2.0 |
| `borehole_column.bas` | Legacy full VBA macro (reference only) |
| `output.svg` | Zigui 14-layer composite section example output |

## Practice Guidelines
1. **Default to SVG** unless user explicitly requests otherwise
2. Support natural language / unstructured descriptions: "draw a stratigraphic column for XX area with N layers..."
3. If user doesn't provide complete stratigraphic data, use the Zigui standard section as example
4. Fill in `grain`, `fossils`, `age_ma` and other extended fields when you can deduce them, to leverage v2.0 features
5. SVG font stack: `SimHei, Heiti SC, sans-serif` for CJK compatibility
6. SVG coordinate system: standard (origin top-left, Y-down); internal geology coordinates are Y-up
7. All SVG elements carry `data-cdr-*` attributes for CorelDRAW identification
