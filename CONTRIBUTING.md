# Contributing to CorelDRAWer-Skill

Thanks for your interest in contributing! This skill generates geological diagrams from natural language descriptions.

## Getting Started

1. **Clone the repo**:
   ```bash
   git clone https://github.com/Reasonite-lab/CorelDRAWer-Skill.git
   cd CorelDRAWer-Skill
   ```

2. **No dependencies needed** — the SVG generators use pure Python stdlib.

3. **Run the demos**:
   ```bash
   python3 generate_column.py
   python3 generate_cross_section.py
   ```

## Project Structure

| Directory/File | Purpose |
|----------------|---------|
| `generate_column.py` | Stratigraphic column SVG generator |
| `generate_cross_section.py` | Geological cross-section SVG generator |
| `cdr_com_auto.py` | CorelDRAW COM / VBA code generator |
| `.reasonix/skills/coreldraw-vba/SKILL.md` | Skill definition for AI agent |
| `data_template.json` | Column data format reference |

## How to Contribute

### Adding a New Lithology Pattern
1. Add the pattern function to `generate_column.py` (and `generate_cross_section.py` if applicable)
2. Register it in the `PATTERNS` dictionary
3. Add documentation in `SKILL.md` and `README.md`

### Adding a New Diagram Type
1. Create a new generator script (e.g., `generate_well_log.py`)
2. Follow the existing conventions:
   - Zero external dependencies
   - SVG output with `cdr-*` layer groups and `data-cdr-*` attributes
   - CLI with `--help` support
   - Built-in demo data
3. Update `SKILL.md` with workflow and data format docs
4. Update `README.md` and `README_zh.md`
5. Add example output to the repo

### Code Style
- Python 3.6+ compatible
- Function docstrings in English
- SVG elements must carry `data-cdr-*` attributes
- Layer grouping: use named `<g>` elements with `id="cdr-*"`
- Keep pattern functions for cross-sections as standalone (no shared imports)

## Pull Request Checklist
- [ ] Code runs with `python3 <script>.py` (zero deps)
- [ ] SVG output validates at https://validator.w3.org/
- [ ] Documentation updated (SKILL.md + README.md)
- [ ] Built-in demo data works
- [ ] `--help` flag works (if applicable)

## License
MIT — see [LICENSE](LICENSE) file.
