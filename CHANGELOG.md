# Changelog

## [v0.2.1] — 2025-06-19

### Added
- **Nature Figure Hunter** (`nature_figure_hunter.py`): search → download PDF → extract figures
  - 82 figures from 20 Communications Earth & Environment papers
  - JPEG (DCTDecode) and FlateDecode (PNG raw) extraction
- **11 new lithology patterns** (29 total): breccia, tuff, andesite, rhyolite, gabbro, serpentinite, slate, quartzite, evaporite, ophiolite, phosphorite
- **Nature style preset** (`--style nature`): Arial font, thinner lines, muted colors
- **Curve track support** (`curves[]`): right-side geophysical/geochemical log panels
- **CITATION.cff**: Zotero/Endnote citation file
- Citation notice in SVG footer

All notable changes to CorelDRAWer-Skill will be documented in this file.

## [v2.1] — 2025-06-18

### Added
- **Geological cross-section generator** (`generate_cross_section.py`)
  - Multi-borehole subsurface profiles with surface interpolation
  - Automatic layer correlation by stratigraphic position
  - Fault rendering: dip line, throw arrows, type label (normal/reverse)
  - Dual scale bars (horizontal + vertical) with exaggeration note
  - 5-layer SVG grouping with `data-cdr-*` attributes
- Cross-section demo data (3 boreholes, 4 layers, 1 fault)
- Cross-section documentation in SKILL.md, README.md, README_zh.md
- MIT License

## [v2.0] — 2025-06-18

### Added
- **7-layer SVG grouping**: `cdr-background`, `cdr-title`, `cdr-header`, `cdr-body`, `cdr-outlines`, `cdr-footer`, `cdr-legend`
- **CorelDRAW metadata**: `data-cdr-*` attributes on all SVG elements
- **Fossil symbol column**: 15 fossil types with Unicode icons, auto-show/hide
- **Sedimentary structure column**: 9 structure types, auto-show/hide
- **Enhanced grain size column**: continuous curve (`grain_profile`) + triangle indicator + Chinese labels
- **Age annotations**: Ma values displayed per layer
- **Adaptive column layout**: 11 columns, optional columns auto-show/hide
- **Inkscape compatibility**: `inkscape:groupmode="layer"` attributes
- VBA generator v2.0: 11-column layout with helper subs
- COM automation v2.0: grain triangles, contact lines, fossil/structures

### Changed
- Rewrote `generate_column.py` with modular structure
- Updated `data_template.json` to v2.0 format with extended fields
- Rewrote `cdr_com_auto.py` VBA generator with 5 helper procedures

### Removed
- All `nature-*` skills moved to `.reasonix_backup/`

## [v1.0] — 2025-06-18

### Added
- Initial stratigraphic column SVG generator (`generate_column.py`)
- 18 standard lithology patterns (GB/T 958)
- CorelDRAW COM automation (`cdr_com_auto.py`)
- VBA macro code generation
- Data template and documentation
- Zigui 14-layer standard section demo
- Junggar Basin ZJ14 well data
