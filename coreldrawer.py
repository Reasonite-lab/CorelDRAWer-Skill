#!/usr/bin/env python3
"""
CorelDRAWer — Unified CLI for geological diagram generation.

Usage:
  coreldrawer column [data.json] [output.svg]
  coreldrawer xsection [data.json] [output.svg]
  coreldrawer vba [data.json] [output.bas]
  coreldrawer com [data.json]               # Windows only
"""

import argparse
import json
import sys


def cmd_column(args):
    from generate_column import generate_svg, DEFAULT_DATA
    data = DEFAULT_DATA
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    fmt = getattr(args, 'format', 'svg')
    style = getattr(args, 'style', 'default')

    if fmt == 'dxf':
        from dxf_export import generate_dxf
        out = args.output or data.get('title', 'column').replace(' ', '_') + '.dxf'
        generate_dxf(data, out)
        print(f"✅ DXF saved: {out}  (import into CorelDRAW)")
        return

    out = args.output or 'output.svg'
    try:
        generate_svg(data, out, style=style)
        n = len(data['layers'])
        total = sum(l['thick'] for l in data['layers'])
        print(f"✅ Column saved: {out}  ({n} layers, {total:,.0f}m, style: {style})")
    except (ValueError, KeyError) as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)


def cmd_xsection(args):
    from generate_cross_section import generate_cross_section, DEMO_DATA
    data = DEMO_DATA
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    out = args.output or 'cross_section.svg'
    try:
        generate_cross_section(data, out)
        n = len(data['boreholes'])
        f = len(data.get('faults', []))
        print(f"✅ Cross-section saved: {out}  ({n} boreholes, {f} faults)")
    except (ValueError, KeyError) as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)


def cmd_vba(args):
    from cdr_com_auto import generate_vba_code
    from generate_column import DEFAULT_DATA
    data = DEFAULT_DATA
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    out = args.output or 'column_macro.bas'
    try:
        generate_vba_code(data, out)
        print(f"✅ VBA saved: {out}  (Alt+F11 → Import → Run DrawColumn)")
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)


def cmd_com(args):
    from cdr_com_auto import com_draw_column
    from generate_column import DEFAULT_DATA
    data = DEFAULT_DATA
    if args.input:
        with open(args.input, 'r', encoding='utf-8') as f:
            data = json.load(f)
    try:
        if not com_draw_column(data):
            print("⚠️  COM unavailable — use 'vba' for macro fallback")
            sys.exit(1)
    except Exception as e:
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)


def main():
    parser = argparse.ArgumentParser(
        description="CorelDRAWer — Geological Diagram Generator",
        epilog="Examples:\n  coreldrawer column\n  coreldrawer column data.json out.svg\n  coreldrawer xsection bh.json xsec.svg")
    sub = parser.add_subparsers(dest='command')

    p = sub.add_parser('column', help='Stratigraphic column')
    p.add_argument('input', nargs='?', help='JSON input')
    p.add_argument('output', nargs='?', help='SVG/DXF output')
    p.add_argument('--style', choices=['default', 'nature'], default='default',
                   help='Visual style (default or nature)')
    p.add_argument('--format', choices=['svg', 'dxf'], default='svg',
                   help='Output format (svg or dxf)')

    p = sub.add_parser('xsection', help='Cross-section')
    p.add_argument('input', nargs='?')
    p.add_argument('output', nargs='?')

    p = sub.add_parser('vba', help='VBA macro')
    p.add_argument('input', nargs='?')
    p.add_argument('output', nargs='?')

    p = sub.add_parser('com', help='CorelDRAW COM (Windows)')
    p.add_argument('input', nargs='?')

    args = parser.parse_args()
    {'column': cmd_column, 'xsection': cmd_xsection,
     'vba': cmd_vba, 'com': cmd_com}.get(args.command, lambda _: parser.print_help())(args)


if __name__ == '__main__':
    main()
