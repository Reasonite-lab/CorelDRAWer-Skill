#!/usr/bin/env python3
"""
CorelDRAWer — Unified CLI for geological diagram generation
═════════════════════════════════════════════════════════
Usage:
  python3 coreldrawer.py column [data.json] [output.svg]
  python3 coreldrawer.py xsection [data.json] [output.svg]
  python3 coreldrawer.py vba [data.json]
  python3 coreldrawer.py com [data.json]       # Windows only

Examples:
  python3 coreldrawer.py column                          # built-in demo
  python3 coreldrawer.py column data.json output.svg     # from JSON
  python3 coreldrawer.py xsection boreholes.json xsec.svg
  python3 coreldrawer.py vba data.json                   # → column_macro.bas
  python3 coreldrawer.py com data.json                   # → CorelDRAW directly
═════════════════════════════════════════════════════════
"""

import sys
import os

def print_help():
    print(__doc__)

def main():
    if len(sys.argv) < 2 or sys.argv[1] in ('--help', '-h', 'help'):
        print_help()
        return

    cmd = sys.argv[1]
    args = sys.argv[2:]

    json_file = None
    output_file = None

    for a in args:
        if a.endswith('.json'):
            json_file = a
        elif a.endswith('.svg') or a.endswith('.bas'):
            output_file = a

    if cmd == 'column':
        from generate_column import generate_svg, DEFAULT_DATA
        data = DEFAULT_DATA
        if json_file:
            import json
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        out = output_file or 'output.svg'
        generate_svg(data, out)
        print(f"✅ Stratigraphic column saved: {out}")

    elif cmd == 'xsection':
        from generate_cross_section import generate_cross_section, DEMO_DATA
        data = DEMO_DATA
        if json_file:
            import json
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        out = output_file or 'cross_section.svg'
        generate_cross_section(data, out)
        print(f"✅ Cross-section saved: {out}")

    elif cmd == 'vba':
        from cdr_com_auto import generate_vba_code
        data = None
        if json_file:
            import json
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        if data is None:
            # Use built-in column demo
            from generate_column import DEFAULT_DATA
            data = DEFAULT_DATA
        out = output_file or 'column_macro.bas'
        generate_vba_code(data, out)
        print(f"✅ VBA macro saved: {out}")
        print(f"   Open CorelDRAW → Alt+F11 → Import → Run DrawColumn")

    elif cmd == 'com':
        from cdr_com_auto import com_draw_column
        data = None
        if json_file:
            import json
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
        if data is None:
            from generate_column import DEFAULT_DATA
            data = DEFAULT_DATA
        com_draw_column(data)

    else:
        print(f"Unknown command: {cmd}")
        print_help()


if __name__ == '__main__':
    main()
