#!/usr/bin/env python3
"""
DXF Exporter for CorelDRAWer-Skill
Pure Python, zero dependencies — generates DXF R12 (max compatibility with CorelDRAW).

Usage:
  python3 dxf_export.py data.json output.dxf
"""

import json
import sys
import os

def cmyk_to_dxf_color(c, m, y, k):
    """Convert CMYK to DXF color index (1-255). Simplified mapping."""
    r = int(255 * (1 - c/100) * (1 - k/100))
    g = int(255 * (1 - m/100) * (1 - k/100))
    b = int(255 * (1 - y/100) * (1 - k/100))
    # Map to closest DXF ACI color
    # Simplified: use grayscale for neutral, hue for colored
    if abs(r-g) < 20 and abs(g-b) < 20:
        gray = int(r / 25.5)
        return max(1, min(9, gray))  # DXF gray colors 250-254
    # Crude mapping to primary colors
    if r > g and r > b: return 1   # red
    if g > r and g > b: return 3   # green
    if b > r and b > g: return 5   # blue
    if r > 150 and g > 100: return 2  # yellow
    return 7  # white/black default

def dxf_header(entities):
    """Generate DXF header."""
    return [
        "0", "SECTION", "2", "ENTITIES"
    ]

def dxf_footer():
    return ["0", "ENDSEC", "0", "EOF"]

def dxf_line(layer, x1, y1, x2, y2, color=7):
    return [
        "0", "LINE", "8", layer,
        "62", str(color),
        "10", f"{x1:.3f}", "20", f"{y1:.3f}",
        "11", f"{x2:.3f}", "21", f"{y2:.3f}"
    ]

def dxf_text(layer, x, y, text, height=2.5, color=7, rotation=0):
    # Strip non-ASCII for DXF compatibility
    safe_text = text.encode('ascii', errors='replace').decode('ascii')
    lines = [
        "0", "TEXT", "8", layer,
        "62", str(color),
        "10", f"{x:.3f}", "20", f"{y:.3f}",
        "40", f"{height:.3f}",
        "1", safe_text
    ]
    if rotation != 0:
        lines.extend(["50", f"{rotation:.3f}"])
    return lines

def dxf_rect(layer, x, y, w, h, color=7):
    """Rectangle as polyline."""
    return [
        "0", "POLYLINE", "8", layer,
        "62", str(color), "66", "1", "70", "1",
        "0", "VERTEX", "8", layer, "10", f"{x:.3f}", "20", f"{y:.3f}",
        "0", "VERTEX", "8", layer, "10", f"{x+w:.3f}", "20", f"{y:.3f}",
        "0", "VERTEX", "8", layer, "10", f"{x+w:.3f}", "20", f"{y+h:.3f}",
        "0", "VERTEX", "8", layer, "10", f"{x:.3f}", "20", f"{y+h:.3f}",
        "0", "SEQEND"
    ]

def dxf_circle(layer, cx, cy, r, color=7):
    return [
        "0", "CIRCLE", "8", layer,
        "62", str(color),
        "10", f"{cx:.3f}", "20", f"{cy:.3f}",
        "40", f"{r:.3f}"
    ]

def dxf_filled_polygon(layer, points, color=7):
    """Filled polygon via SOLID or HATCH."""
    # Use multiple SOLID entities for simple shapes
    # For complex polygons, use HATCH
    if len(points) < 3:
        return []
    result = []
    # Use 3-point SOLID for triangles, HATCH for others
    result.extend([
        "0", "HATCH", "8", layer,
        "62", str(color),
        "70", "1",  # solid fill
        "71", "1",  # non-associative
        "0", "BOUNDARY",
        "92", "1",  # one boundary
        "72", "0",  # polyline
        "73", "1",  # closed
        "93", str(len(points))
    ])
    for px, py in points:
        result.extend(["10", f"{px:.3f}", "20", f"{py:.3f}"])
    result.extend([
        "97", "0",  # no more source objects
        "0", "ENDSEC"  # end hatch
    ])
    return result

def generate_dxf(data, output_path):
    """Generate a DXF stratigraphic column from JSON data."""
    layers = data['layers']
    title = data.get('title', 'Stratigraphic Column')
    total_thick = sum(l['thick'] for l in layers)

    # Layout (mm units, same as SVG generator)
    MARGIN_LEFT = 8
    TABLE_TOP = 290
    TABLE_BOTTOM = 22
    HEADER_H = 14
    TABLE_RIGHT = 205
    draw_h = TABLE_TOP - TABLE_BOTTOM - HEADER_H
    scale_f = draw_h / total_thick if total_thick > 0 else 1

    col_x = [0, 8, 21, 34, 47, 65, 77, 84, 124, 139, 146, 156]
    col_w = [0, 13, 13, 13, 18, 12, 7, 40, 15, 7, 10, 65]

    entities = []

    # DXF layers
    layers_used = set()

    def emit(*lines):
        entities.extend(lines)

    # ── Title ──
    emit(*dxf_text("TITLE", (MARGIN_LEFT + TABLE_RIGHT)/2, TABLE_TOP + 12,
                   title.encode('ascii','replace').decode('ascii'), height=4.5, color=7))
    layers_used.add("TITLE")

    # ── Header ──
    emit(*dxf_rect("HEADER", MARGIN_LEFT, TABLE_TOP - HEADER_H,
                   TABLE_RIGHT - MARGIN_LEFT, HEADER_H, color=8))
    layers_used.add("HEADER")

    header_labels = [
        ("Era", 1), ("Sys", 2), ("Ser", 3), ("Fm", 4), ("Sym", 5),
        ("Fossils", 6), ("Lithology", 7), ("Grain", 8),
        ("Struc", 9), ("Thick(m)", 10), ("Description", 11)
    ]
    for label, col in header_labels:
        cx = col_x[col] + col_w[col] / 2
        emit(*dxf_text("HEADER", cx, TABLE_TOP - HEADER_H/2, label, height=2.8, color=7))

    # ── Column dividers ──
    for ci in range(2, 12):
        emit(*dxf_line("GRID", col_x[ci], TABLE_BOTTOM, col_x[ci], TABLE_TOP, color=8))
    layers_used.add("GRID")

    # ── Outer border ──
    emit(*dxf_line("BORDER", MARGIN_LEFT, TABLE_BOTTOM, MARGIN_LEFT, TABLE_TOP, color=7))
    emit(*dxf_line("BORDER", TABLE_RIGHT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_TOP, color=7))
    emit(*dxf_line("BORDER", MARGIN_LEFT, TABLE_TOP, TABLE_RIGHT, TABLE_TOP, color=7))
    emit(*dxf_line("BORDER", MARGIN_LEFT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_BOTTOM, color=7))
    layers_used.add("BORDER")

    # ── Layers ──
    currentY = TABLE_TOP - HEADER_H
    for i, layer in enumerate(layers):
        layH = max(layer['thick'] * scale_f, 3.5)
        layB = currentY - layH
        midY = (currentY + layB) / 2
        c, m, y, k = layer['c'], layer['m'], layer['y'], layer['k']
        color_idx = cmyk_to_dxf_color(c, m, y, k)

        # Lithology rect
        lx = col_x[7] + 0.3
        lw = col_w[7] - 2.5
        emit(*dxf_rect("LITHOLOGY", lx, layB, lw, layH, color=color_idx))

        # Pattern dots (simplified — sparse dots for sandstone-like)
        pname = layer.get('pattern', 'pure')
        if pname in ('sand', 'finesand', 'conglo'):
            import random
            rng = random.Random(i * 100)
            for _ in range(int(lw * layH / 8)):
                dx = lx + 1 + rng.random() * (lw - 2)
                dy = layB + 1 + rng.random() * (layH - 2)
                emit(*dxf_circle("PATTERN", dx, dy, 0.3, color=8))

        # Formation text
        fm = layer.get('formation', '')
        if fm:
            emit(*dxf_text("TEXT", col_x[4] + 1, midY, fm, height=2.2, color=7))

        # Symbol
        sym = layer.get('symbol', '')
        if sym:
            emit(*dxf_text("TEXT", col_x[5] + 1, midY, sym, height=2.2, color=7))

        # Thickness
        emit(*dxf_text("TEXT", col_x[10] + 1, midY, f"{layer['thick']:.1f}", height=2.2, color=7))

        # Description
        descr = layer.get('descr', '')
        if descr:
            emit(*dxf_text("TEXT", col_x[11] + 1, midY, descr, height=2.0, color=7))

        # Row separator
        emit(*dxf_line("GRID", MARGIN_LEFT, currentY, TABLE_RIGHT, currentY, color=8))

        currentY = layB

    layers_used.add("LITHOLOGY")
    layers_used.add("PATTERN")
    layers_used.add("TEXT")

    # ── Footer ──
    ratio = int(total_thick * 1000 / draw_h) if draw_h > 0 else 1
    footer = f"{title} | {total_thick:,.0f}m | {len(layers)} layers | Scale 1:{ratio:,}"
    emit(*dxf_text("FOOTER", MARGIN_LEFT, TABLE_BOTTOM - 5, footer, height=2.0, color=8))
    emit(*dxf_text("FOOTER", MARGIN_LEFT, TABLE_BOTTOM - 8,
                   "CorelDRAWer-Skill (github.com/Reasonite-lab/CorelDRAWer-Skill)",
                   height=1.6, color=9))
    layers_used.add("FOOTER")

    # ── Assemble DXF ──
    lines = []
    lines.extend(dxf_header(entities))
    lines.extend(entities)
    lines.extend(dxf_footer())

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    print(f"✅ DXF saved: {output_path}")
    print(f"   Layers: {', '.join(sorted(layers_used))}")
    print(f"   Import into CorelDRAW: File → Import → {output_path}")
    return output_path


# ── CLI ──
if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 dxf_export.py data.json output.dxf")
        sys.exit(1)

    json_file = None
    output = None
    for a in sys.argv[1:]:
        if a.endswith('.json'):
            json_file = a
        elif a.endswith('.dxf'):
            output = a

    if json_file:
        with open(json_file, 'r') as f:
            data = json.load(f)
    else:
        from generate_column import DEFAULT_DATA
        data = DEFAULT_DATA
        json_file = '(built-in demo)'

    if not output:
        output = data.get('title', 'column').replace(' ', '_') + '.dxf'

    generate_dxf(data, output)
