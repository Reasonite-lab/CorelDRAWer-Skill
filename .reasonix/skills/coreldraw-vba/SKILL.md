---
name: coreldraw-vba
description: Convert natural language descriptions into CorelDRAW-compatible vector graphics — stratigraphic columns, geological cross-sections, SVG/VBA/COM output
---

# CorelDRAWer-Skill v2.1

## Quick Decision Tree

```
User says...                           → You do...
─────────────────────────────────────────────────────────
"draw a stratigraphic column"          → generate_column.py (11-col)
"draw a borehole log"                  → generate_column.py (11-col)
"make a cross-section / profile"       → generate_cross_section.py
"multiple boreholes / wells"           → generate_cross_section.py
"generate VBA code"                    → cdr_com_auto.py --vba
"control CorelDRAW directly"           → cdr_com_auto.py --com
"just show me"                         → generate SVG (default)
no specific data given                 → use built-in DEMO data
```

## Role

You are an AI assistant for geological diagramming. Users describe diagrams in natural language. You:
1. Classify the diagram type (column vs cross-section)
2. Build structured JSON data from their description
3. Run the appropriate Python generator
4. Present the result (SVG file)

## When to Use Each Generator

### `generate_column.py` — Stratigraphic Columns & Borehole Logs

Use for: vertical stratigraphic columns, borehole logs, single-well lithology.

**Natural language triggers:**
- "Draw a stratigraphic column for..."
- "Generate a borehole log with layers..."
- "Make a lithology column showing..."

**Data structure** (write to a temp `.json` file, then run):
```json
{
  "title": "Composite Stratigraphic Column",
  "location": "Zigui, Hubei",
  "layers": [
    {
      "erathem": "Neoproterozoic", "system": "Nanhua", "series": "Lower",
      "formation": "Liantuo Fm", "symbol": "Nh₁l", "thick": 120,
      "descr": "Purplish-red medium-bedded sandstone",
      "c": 0, "m": 40, "y": 30, "k": 10,
      "pattern": "sand", "grain": 4,
      "fossils": ["trilobite"], "structures": ["cross_bed"],
      "contact": "unconformity", "age_ma": 780
    }
  ]
}
```

**Run:** `python3 generate_column.py data.json output.svg`

### `generate_cross_section.py` — Geological Cross-Sections

Use for: subsurface profiles across multiple boreholes, structural sections.

**Natural language triggers:**
- "Draw a cross-section through these boreholes..."
- "Show the subsurface along line A-A'..."
- "Connect these wells in a profile..."

**Data structure:**
```json
{
  "title": "Cross-Section A-A'",
  "orientation": "NW → SE",
  "vertical_exaggeration": 5,
  "boreholes": [
    {
      "id": "ZK001", "x": 0, "elevation": 520, "depth": 180,
      "layers": [
        {"formation": "Alluvium", "thick": 8, "c": 0, "m": 10, "y": 25, "k": 10, "pattern": "conglo"},
        {"formation": "Liantuo Fm", "thick": 55, "c": 0, "m": 40, "y": 30, "k": 10, "pattern": "sand"}
      ]
    }
  ],
  "faults": [
    {"x": 200, "dip": 72, "direction": "NE", "type": "normal", "throw": 40}
  ]
}
```

**Run:** `python3 generate_cross_section.py data.json output.svg`

## Data Reference

### CMYK Color Quick-Guide
| Rock Type | C | M | Y | K | Hex approx |
|--------------|---|---|---|---|------------|
| Sandstone (red) | 0 | 40 | 30 | 10 | `#e589a0` |
| Limestone (gray) | 0 | 0 | 0 | 30 | `#b3b3b3` |
| Shale (dark) | 0 | 0 | 0 | 55 | `#737373` |
| Dolomite (light) | 0 | 0 | 5 | 8 | `#ebebe0` |
| Coal | 0 | 0 | 0 | 90 | `#1a1a1a` |
| Mudstone | 5 | 0 | 15 | 20 | `#c2ccad` |
| Conglomerate | 15 | 0 | 20 | 20 | `#adcca3` |
| Granite (pink) | 0 | 30 | 20 | 5 | `#f2aac4` |

### Lithology → Pattern Mapping (infer from descriptions)
| Keywords in description | → `pattern` | → `grain` |
|-------------------------|-------------|-----------|
| sandstone, 砂岩 | `sand` | 3-4 |
| conglomerate, 砾岩 | `conglo` | 5-6 |
| limestone, 灰岩 | `lime` | 2-3 |
| dolomite, 白云岩 | `dolo` | 2 |
| shale, 页岩 | `shale` | 1 |
| mudstone, 泥岩 | `mud` | 1 |
| siltstone, 粉砂岩 | `silt` | 2 |
| coal, 煤 | `coal` | 1 |
| granite, 花岗岩 | `granite` | 4-5 |
| basalt, 玄武岩 | `basalt` | 3 |
| marble, 大理岩 | `marble` | 2-3 |
| chert, 硅质岩 | `chert` | 2 |
| no pattern / 纯色 | `pure` | — |

### Fossil Codes (15 types)
`trilobite` `brachiopod` `cephalopod` `graptolite` `crinoid` `algae` `stromatolite` `ammonite` `coral` `bivalve` `gastropod` `foraminifera` `plant` `fish` `spore`

### Structure Codes (9 types)
`ripple` `cross_bed` `graded` `ooid` `crack` `concretion` `bioturbation` `stylolite` `stromatactis`

### Contact Types
`"conformity"` (normal, no marker) · `"disconformity"` (wavy line) · `"unconformity"` (sawtooth line)

### Grain Size Scale
`1`=clay · `2`=silt · `3`=fine sand · `4`=medium sand · `5`=coarse sand · `6`=gravel

## Workflow

### Step 1: Parse Intent
- Listen for keywords: "column"/"柱状图"/"log" → column generator; "cross-section"/"profile"/"剖面" → cross-section generator
- Extract: formation names, thicknesses, lithology descriptions, location
- If user says "like Zigui" or gives no data → use built-in demo (just run without args)

### Step 2: Build JSON
- Write a temp JSON file with the extracted data
- Infer `pattern` from lithology keywords (see mapping above)
- Estimate CMYK from rock color descriptions
- Set `grain` from grain-size clues in description
- Add `fossils`/`structures` if mentioned

### Step 3: Run Generator
```bash
# Unified CLI (recommended)
python3 coreldrawer.py column /tmp/data.json output.svg
python3 coreldrawer.py xsection /tmp/data.json output.svg
python3 coreldrawer.py vba /tmp/data.json
python3 coreldrawer.py com /tmp/data.json    # Windows only

# Or directly
python3 generate_column.py /tmp/data.json output.svg
python3 generate_cross_section.py /tmp/data.json output.svg
```

### Step 4: Present
1. Tell user the output filename
2. Mention key stats (N layers, total thickness, patterns used)
3. Remind: "Drag into CorelDRAW to edit — layers are preserved"

## Rules

1. **Always default to SVG** — most universal
2. **Use temp JSON files** — write to `/tmp/` or project dir, run generator on it
3. **Demo data when stuck** — run script without args to show an example
4. **Error recovery** — if generator fails, check JSON is valid; if still fails, offer VBA fallback
5. **Single borehole with layers** → use column generator, NOT cross-section
6. **Multiple boreholes with X coordinates** → use cross-section generator
7. **Fill in ALL possible fields** — the more data, the richer the output

## Output Channels

| Channel | Command | When to Use |
|---------|---------|-------------|
| **SVG** | `python3 generate_column.py` or `generate_cross_section.py` | Default, always |
| **VBA** | `python3 coreldrawer.py vba data.json` | User asks for VBA / Windows unavailable |
| **COM** | `python3 coreldrawer.py com data.json` | Windows + CorelDRAW installed |

## SVG Layer Reference

Output SVGs use named `<g>` groups with `data-cdr-*` attributes:

| Group ID | Contains |
|----------|----------|
| `cdr-background` | White page background |
| `cdr-title` | Title, subtitle, orientation |
| `cdr-header` | Column headers, divider lines |
| `cdr-body` | All geological content (layers, faults, scale ticks) |
| `cdr-outlines` | Table/frame borders |
| `cdr-footer` | Scale info, summary statistics |
| `cdr-legend` | Lithology/pattern legend |

## Cross-Section Specifics

### Fault Data
```json
{"x": 200, "dip": 72, "direction": "NE", "type": "normal", "throw": 40}
```
- `type`: `"normal"` or `"reverse"`
- `throw`: vertical offset in meters (positive number)
- `dip`: degrees from horizontal

### Vertical Exaggeration
- Default: 5×
- Flatter terrain → use higher VE (5–10×)
- Mountainous terrain → use lower VE (2–3×)
