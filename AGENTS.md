# CorelDRAWer-Skill Project

A Reasonix skill project for generating CorelDRAW-compatible geological diagrams. Describe what you want in natural language and get production-ready SVG vector graphics — or VBA macros / COM automation for direct CorelDRAW integration.

## Project

- **Purpose**: Hosts the `coreldraw-vba` Reasonix skill for AI-assisted geological diagram generation
- **Core Skill**: `coreldraw-vba` — converts natural language descriptions into stratigraphic column SVGs, VBA macros, or CorelDRAW COM drawings
- **Skill Location**: `.reasonix/skills/coreldraw-vba/SKILL.md`

## Commands

No build/test commands. The only operation is:

- `/coreldraw-vba` — invoke the drawing skill (or describe your diagram naturally in conversation)

## Output Channels

| Channel | Command | Platform |
|---------|---------|----------|
| SVG (default) | `python3 generate_column.py data.json output.svg` | Any |
| VBA Macro | `python3 cdr_com_auto.py --vba data.json` | Any |
| COM Automation | `python3 cdr_com_auto.py --com data.json` | Windows only |

## Architecture

```
.reasonix/skills/coreldraw-vba/SKILL.md   ← Skill definition (v2.0)
generate_column.py                         ← SVG generator (zero deps)
cdr_com_auto.py                            ← VBA/COM generator
data_template.json                         ← Data format template
borehole_column.bas                        ← Legacy VBA macro (reference)
AGENTS.md                                  ← This file
README.md                                  ← English documentation
README_zh.md                               ← Chinese documentation
```

## Conventions

- Default output: SVG with 7 named layer groups and `data-cdr-*` attributes for CorelDRAW
- All generated code includes undo grouping (`BeginCommandGroup`/`EndCommandGroup`) and error handling
- Default units: millimeters (mm)
- Default page: A4 landscape
- Font: SimHei with Heiti SC / sans-serif fallback (Chinese-compatible)
- Coordinate system: SVG standard (origin top-left, Y-down), internally Y-up for geology
- 18 standard lithology patterns per GB/T 958
- Zero external dependencies for SVG generation

## Notes

- The `borehole_column.bas` file is a legacy 688-line VBA macro kept for reference; the `cdr_com_auto.py` VBA generator is the recommended replacement (v2.0, 11 columns)
- Fossil and structure columns are adaptive — they only appear when the input data includes them
- On macOS/Linux, COM mode automatically falls back to VBA code generation
