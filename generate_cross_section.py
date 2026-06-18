#!/usr/bin/env python3
"""
Geological Cross-Section SVG Generator v1.0
══════════════════════════════════════════════════════════
Generates geological cross-sections from borehole data.
Input: multiple boreholes with X, elevation, layer data + optional faults
Output: SVG vector graphics (CorelDRAW / Illustrator / browser ready)

Usage:
  python3 generate_cross_section.py data.json output.svg
  python3 generate_cross_section.py                 # use built-in demo

Dependencies: none (pure Python stdlib)
══════════════════════════════════════════════════════════
"""

import json
import sys
import os
import math
import xml.etree.ElementTree as ET
from xml.dom import minidom

# ============================================================================
# DATA
# ============================================================================

DEMO_DATA = {
    "title": "Geological Cross-Section A-A'",
    "orientation": "NW → SE",
    "vertical_exaggeration": 5,
    "boreholes": [
        {
            "id": "ZK001",
            "x": 0,
            "elevation": 520.0,
            "depth": 180,
            "layers": [
                {"formation": "Quaternary alluvium", "thick": 8, "c": 0, "m": 10, "y": 25, "k": 10, "pattern": "conglo"},
                {"formation": "Liantuo Fm (sandstone)", "thick": 55, "c": 0, "m": 40, "y": 30, "k": 10, "pattern": "sand"},
                {"formation": "Nantuo Fm (conglomerate)", "thick": 32, "c": 15, "m": 0, "y": 20, "k": 20, "pattern": "conglo"},
                {"formation": "Doushantuo Fm (dolomite)", "thick": 85, "c": 5, "m": 0, "y": 10, "k": 30, "pattern": "dolo"},
            ]
        },
        {
            "id": "ZK002",
            "x": 350,
            "elevation": 485.0,
            "depth": 210,
            "layers": [
                {"formation": "Quaternary alluvium", "thick": 12, "c": 0, "m": 10, "y": 25, "k": 10, "pattern": "conglo"},
                {"formation": "Liantuo Fm (sandstone)", "thick": 70, "c": 0, "m": 40, "y": 30, "k": 10, "pattern": "sand"},
                {"formation": "Nantuo Fm (conglomerate)", "thick": 28, "c": 15, "m": 0, "y": 20, "k": 20, "pattern": "conglo"},
                {"formation": "Doushantuo Fm (dolomite)", "thick": 100, "c": 5, "m": 0, "y": 10, "k": 30, "pattern": "dolo"},
            ]
        },
        {
            "id": "ZK003",
            "x": 700,
            "elevation": 510.0,
            "depth": 195,
            "layers": [
                {"formation": "Quaternary alluvium", "thick": 5, "c": 0, "m": 10, "y": 25, "k": 10, "pattern": "conglo"},
                {"formation": "Liantuo Fm (sandstone)", "thick": 60, "c": 0, "m": 40, "y": 30, "k": 10, "pattern": "sand"},
                {"formation": "Nantuo Fm (conglomerate)", "thick": 35, "c": 15, "m": 0, "y": 20, "k": 20, "pattern": "conglo"},
                {"formation": "Doushantuo Fm (dolomite)", "thick": 95, "c": 5, "m": 0, "y": 10, "k": 30, "pattern": "dolo"},
            ]
        }
    ],
    "faults": [
        {"x": 200, "dip": 72, "direction": "NE", "type": "normal", "throw": 40}
    ]
}


# ============================================================================
# PATTERN GENERATORS (shared with generate_column.py)
# ============================================================================

PAT_SPACING = 6.0  # larger spacing for cross-section (wider bands)

def _append(parent, tag, attrs):
    e = ET.SubElement(parent, tag)
    for k, v in attrs.items():
        e.set(k.replace('_', '-'), str(v))
    return e

def add_rect(parent, x, y, w, h, fill=None, stroke='#000', sw=0.2, **extra):
    attrs = {'x': x, 'y': y, 'width': w, 'height': h, 'fill': fill or 'none',
             'stroke': stroke, 'stroke-width': str(sw)}
    attrs.update(extra)
    return _append(parent, 'rect', attrs)

def add_line(parent, x1, y1, x2, y2, stroke='#000', sw=0.2, **extra):
    a = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2, 'stroke': stroke, 'stroke-width': str(sw)}
    a.update(extra)
    return _append(parent, 'line', a)

def add_text(parent, x, y, text, size=3, bold=False, anchor='start', fill='#000', **extra):
    t = ET.SubElement(parent, 'text')
    t.set('x', str(x))
    t.set('y', str(y))
    t.set('font-family', 'SimHei, Heiti SC, sans-serif')
    t.set('font-size', str(size))
    t.set('font-weight', 'bold' if bold else 'normal')
    t.set('text-anchor', anchor)
    t.set('fill', fill)
    for k, v in extra.items():
        t.set(k.replace('_', '-'), str(v))
    t.text = str(text)
    return t

def add_circle(parent, cx, cy, r, fill='#000', stroke=None, sw=0):
    a = {'cx': cx, 'cy': cy, 'r': r, 'fill': fill}
    if stroke:
        a['stroke'] = stroke
        a['stroke-width'] = str(sw)
    return _append(parent, 'circle', a)

def cmyk_fill(c, m, y, k):
    r = int(255 * (1 - c/100) * (1 - k/100))
    g = int(255 * (1 - m/100) * (1 - k/100))
    b = int(255 * (1 - y/100) * (1 - k/100))
    return f"#{r:02x}{g:02x}{b:02x}"

# Simplified pattern set for cross-sections
def pat_sand(g, lx, by, w, h):
    sp = PAT_SPACING * 0.7
    x = lx + sp/2
    while x < lx + w:
        y = by + sp/2
        while y < by + h:
            add_circle(g, x, y, 0.3, fill='#555')
            y += sp
        x += sp

def pat_conglo(g, lx, by, w, h):
    sp = PAT_SPACING * 1.8
    x = lx + sp*0.7
    while x < lx + w - sp*0.3:
        y = by + sp*0.7
        while y < by + h - sp*0.3:
            add_circle(g, x, y, 1.2, fill='none', stroke='#000', sw=0.2)
            y += sp
        x += sp
    # dots
    x = lx + sp*1.6
    while x < lx + w:
        y = by + sp*1.6
        while y < by + h:
            add_circle(g, x, y, 0.3, fill='#555')
            y += sp
        x += sp

def pat_dolo(g, lx, by, w, h):
    sp = PAT_SPACING * 1.2
    dx = lx - (w + h)
    while dx < lx + w + h:
        add_line(g, dx, by, dx + h, by + h, stroke='#999', sw=0.12)
        dx += sp
    dx = lx - (w + h)
    while dx < lx + w + h:
        add_line(g, dx, by + h, dx + h, by, stroke='#999', sw=0.12)
        dx += sp

def pat_lime(g, lx, by, w, h):
    sp = PAT_SPACING
    bw, bh = sp*1.5, sp*0.6
    row, y = 0, by + 0.3
    while y < by + h:
        x = lx + 0.3 + (bw/2 if row % 2 else 0)
        while x < lx + w:
            add_rect(g, x, y, bw-0.4, bh, fill='none', stroke='#666', sw=0.14)
            x += bw
        y += bh + 0.25
        row += 1

def pat_shale(g, lx, by, w, h):
    sp = PAT_SPACING * 0.7
    y = by + sp*0.3
    cnt = 0
    while y < by + h:
        add_line(g, lx + 0.3, y, lx + w - 0.3, y, stroke='#555', sw=0.2)
        cnt += 1
        if cnt % 3 != 0:
            svx = lx + sp*0.8
            while svx < lx + w:
                add_line(g, svx, y - sp*0.3, svx, y + sp*0.3, stroke='#999', sw=0.12)
                svx += sp*1.5
        y += sp

def pat_silt(g, lx, by, w, h):
    sp = PAT_SPACING
    y = by + sp*0.5
    while y < by + h:
        add_line(g, lx + 1, y, lx + w - 1, y, stroke='#aaa', sw=0.18)
        y += sp
    x = lx + sp
    while x < lx + w:
        y = by + sp
        while y < by + h:
            add_circle(g, x, y, 0.18, fill='#888')
            y += sp*2
        x += sp*2

def pat_carbShale(g, lx, by, w, h):
    sp = PAT_SPACING * 0.7
    y = by + sp*0.3
    while y < by + h:
        add_line(g, lx + 0.3, y, lx + w - 0.3, y, stroke='#eee', sw=0.15)
        y += sp

def pat_mud(g, lx, by, w, h):
    sp = PAT_SPACING * 0.6
    y = by + sp*0.3
    while y < by + h:
        add_line(g, lx + 0.3, y, lx + w - 0.3, y, stroke='#555', sw=0.2)
        y += sp

def pat_coal(g, lx, by, w, h):
    """Coal — solid black with white bedding lines"""
    add_line(g, lx, by + h/3, lx + w, by + h/3, stroke='#fff', sw=0.4)
    add_line(g, lx, by + h*2/3, lx + w, by + h*2/3, stroke='#fff', sw=0.4)

def pat_granite(g, lx, by, w, h):
    """Granite — crosses + dense dots"""
    import random
    rng = random.Random(42)
    sp = PAT_SPACING * 0.8
    cnt = int(w * h / (sp * sp) * 0.4)
    for _ in range(cnt):
        x = lx + rng.random() * w
        y = by + rng.random() * h
        gs = sp * 0.3
        add_line(g, x - gs, y - gs, x + gs, y + gs, stroke='#555', sw=0.15)
        add_line(g, x - gs, y + gs, x + gs, y - gs, stroke='#555', sw=0.15)
    x = lx + sp * 0.5
    while x < lx + w:
        y = by + sp * 0.5
        while y < by + h:
            add_circle(g, x, y, 0.25, fill='#555')
            y += sp
        x += sp

def pat_basalt(g, lx, by, w, h):
    """Basalt — V-pattern diagonal"""
    sp = PAT_SPACING * 0.9
    y = by + sp * 0.5
    while y < by + h:
        x = lx + sp * 0.3
        while x < lx + w:
            vh = sp * 0.35
            add_line(g, x, y - vh, x + vh, y, stroke='#555', sw=0.18)
            add_line(g, x + vh, y, x + vh * 2, y - vh, stroke='#555', sw=0.18)
            x += sp * 1.5
        y += sp

def pat_schist(g, lx, by, w, h):
    """Schist — wavy folds"""
    sp = PAT_SPACING * 0.9
    dx = lx - w
    while dx < lx + w * 2:
        mx = dx + h * 0.5
        my = by + h * 0.5
        add_line(g, dx, by, mx, my, stroke='#666', sw=0.2)
        add_line(g, mx, my, dx + h, by + h, stroke='#666', sw=0.2)
        dx += sp

def pat_gneiss(g, lx, by, w, h):
    """Gneiss — alternating thick/thin bands"""
    sp = PAT_SPACING * 0.8
    y = by + sp * 0.2
    band = True
    while y < by + h:
        sw_val = 0.3 if band else 0.1
        col = '#333' if band else '#bbb'
        add_line(g, lx + 0.3, y, lx + w - 0.3, y, stroke=col, sw=sw_val)
        band = not band
        y += sp * (0.6 if band else 0.3)

def pat_marble(g, lx, by, w, h):
    """Marble — fine grid"""
    sp = PAT_SPACING * 1.2
    x = lx + sp * 0.5
    while x < lx + w:
        add_line(g, x, by, x, by + h, stroke='#ddd', sw=0.07)
        x += sp
    y = by + sp * 0.5
    while y < by + h:
        add_line(g, lx, y, lx + w, y, stroke='#ddd', sw=0.07)
        y += sp

def pat_chert(g, lx, by, w, h):
    """Chert — bold cross-hatch"""
    sp = PAT_SPACING * 1.5
    dx = lx - (w + h)
    while dx < lx + w + h:
        add_line(g, dx, by, dx + h, by + h, stroke='#555', sw=0.22)
        dx += sp
    dx = lx - (w + h)
    while dx < lx + w + h:
        add_line(g, dx, by + h, dx + h, by, stroke='#555', sw=0.22)
        dx += sp

def pat_doloLime(g, lx, by, w, h):
    """Dolomitic limestone — brick grid + rhombic"""
    sp = PAT_SPACING * 1.5
    bw, bh = sp * 2, sp * 0.7
    row, y = 0, by + 0.3
    while y < by + h:
        x = lx + 0.3 + (bw/2 if row % 2 else 0)
        while x < lx + w:
            add_rect(g, x, y, bw-0.4, bh, fill='none', stroke='#bbb', sw=0.1)
            x += bw
        y += bh + 0.25
        row += 1
    dx = lx - (w + h)
    while dx < lx + w + h:
        add_line(g, dx, by, dx + h, by + h, stroke='#ccc', sw=0.08)
        dx += sp * 2

PATTERNS = {
    'sand': pat_sand, 'conglo': pat_conglo, 'dolo': pat_dolo,
    'lime': pat_lime, 'shale': pat_shale, 'silt': pat_silt,
    'carbShale': pat_carbShale, 'mud': pat_mud,
    'finesand': pat_sand, 'doloLime': pat_doloLime,
    'chert': pat_chert, 'coal': pat_coal,
    'granite': pat_granite, 'basalt': pat_basalt,
    'schist': pat_schist, 'gneiss': pat_gneiss,
    'marble': pat_marble,
    'pure': lambda g,x,y,w,h: None,
}


# ============================================================================
# MAIN GENERATOR
# ============================================================================

def generate_cross_section(data, output_path=None):
    """
    Generate a geological cross-section SVG from borehole data.

    data format:
    {
        "title": "Cross-Section A-A'",
        "orientation": "NW → SE",
        "vertical_exaggeration": 5,
        "boreholes": [
            {"id": "ZK001", "x": 0, "elevation": 520, "depth": 180,
             "layers": [{"formation": "...", "thick": N, "c":0, "m":0, "y":0, "k":0, "pattern":"..."}, ...]},
            ...
        ],
        "faults": [
            {"x": 200, "dip": 72, "direction": "NE", "type": "normal", "throw": 40}
        ]
    }
    """
    # Input validation
    if not data.get('boreholes'):
        raise ValueError("No boreholes data provided")
    
    boreholes = data['boreholes']
    if len(boreholes) < 2:
        # Single borehole → draw a column-like view
        pass  # still works, just no correlation lines
    
    faults = data.get('faults', [])
    title = data.get('title', "Geological Cross-Section")
    orientation = data.get('orientation', '')
    vx = min(max(data.get('vertical_exaggeration', 5), 1), 20)  # clamp 1-20

    # ============================
    # Coordinate calculations
    # ============================
    # Surface Y in SVG = f(elevation)
    # Subsurface Y = f(elevation - cumulative_thickness)

    max_elev = max(bh['elevation'] for bh in boreholes)
    min_bottom = min(bh['elevation'] - sum(l['thick'] for l in bh['layers']) for bh in boreholes)

    min_x = min(bh['x'] for bh in boreholes)
    max_x = max(bh['x'] for bh in boreholes)

    # Horizontal scale: fit x range into available width
    MARGIN = 50
    BOTTOM_MARGIN = 40
    TOP_MARGIN = 50

    section_width = 600  # drawing width for the section body
    page_w = section_width + MARGIN * 2 + 180  # room for legend

    h_scale = section_width / (max_x - min_x) if max_x > min_x else 1
    # Vertical scale with exaggeration
    elev_range = max_elev - min_bottom
    section_height = 350
    v_scale = section_height / elev_range * vx if elev_range > 0 else 1

    page_h = section_height + TOP_MARGIN + BOTTOM_MARGIN + 60

    def to_svg_x(geo_x):
        return MARGIN + (geo_x - min_x) * h_scale

    def to_svg_y(elev):
        # SVG Y goes down; higher elevation = lower Y value
        return TOP_MARGIN + (max_elev - elev) * v_scale / vx

    # ============================
    # Build SVG
    # ============================
    svg = ET.Element('svg', {
        'xmlns': 'http://www.w3.org/2000/svg',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'viewBox': f'0 0 {page_w} {page_h}',
        'width': f'{page_w}mm',
        'height': f'{page_h}mm',
        'version': '1.1'
    })
    defs = ET.SubElement(svg, 'defs')

    # Background
    bg = ET.SubElement(svg, 'g', {'id': 'cdr-background', 'data-cdr-layer': 'background'})
    add_rect(bg, 0, 0, page_w, page_h, fill='#fff', stroke='none', sw=0)

    # Title
    title_g = ET.SubElement(svg, 'g', {'id': 'cdr-title', 'data-cdr-layer': 'title'})
    cx = MARGIN + section_width / 2
    add_text(title_g, cx, 20, title, size=5, bold=True, anchor='middle', fill='#000',
             data_cdr_type='title')
    if orientation:
        add_text(title_g, cx, 28, orientation, size=3, anchor='middle', fill='#666',
                 data_cdr_type='orientation')
    ve_text = f"Vertical exaggeration: {vx}×"
    add_text(title_g, cx, 35, ve_text, size=2.5, anchor='middle', fill='#999',
             data_cdr_type='ve-note')

    # Body layer
    body = ET.SubElement(svg, 'g', {'id': 'cdr-body', 'data-cdr-layer': 'body'})

    # ============================
    # Surface profile
    # ============================
    surface_points = []
    for bh in boreholes:
        sx = to_svg_x(bh['x'])
        sy = to_svg_y(bh['elevation'])
        surface_points.append((sx, sy))

    # Draw surface line (connect all boreholes)
    d_path = f'M {surface_points[0][0]},{surface_points[0][1]}'
    for sx, sy in surface_points[1:]:
        d_path += f' L {sx},{sy}'
    ET.SubElement(body, 'path', {
        'd': d_path, 'fill': 'none', 'stroke': '#2e7d32', 'stroke-width': '1.2',
        'data-cdr-type': 'surface-line'
    })

    # Surface fill (green-tinted area above surface)
    d_fill = d_path
    # Close the path at top of page
    d_fill += f' L {surface_points[-1][0]},{TOP_MARGIN - 30}'
    d_fill += f' L {surface_points[0][0]},{TOP_MARGIN - 30} Z'
    ET.SubElement(body, 'path', {
        'd': d_fill, 'fill': '#e8f5e9', 'stroke': 'none',
        'data-cdr-type': 'surface-fill'
    })

    # ============================
    # Borehole markers
    # ============================
    for bh in boreholes:
        sx = to_svg_x(bh['x'])
        sy = to_svg_y(bh['elevation'])
        bot_y = to_svg_y(bh['elevation'] - sum(l['thick'] for l in bh['layers']))

        # Borehole line (thin, dashed)
        add_line(body, sx, sy, sx, bot_y, stroke='#333', sw=0.4,
                 data_cdr_type='borehole-line')
        # Borehole collar marker
        add_circle(body, sx, sy, 2.5, fill='#c0392b', stroke='#fff', sw=0.5)
        add_text(body, sx, sy - 5, bh['id'], size=3, bold=True, anchor='middle',
                 fill='#c0392b', data_cdr_type='borehole-label')

    # ============================
    # Layer correlation & drawing
    # ============================
    # Strategy: match layers by formation name across boreholes
    # If names match, draw a continuous band between boreholes

    # Build layer boundaries for each borehole
    bh_boundaries = []  # list of (sx, top_y, bottom_y, layer)
    for bh in boreholes:
        sx = to_svg_x(bh['x'])
        cum_thick = 0
        bh_layers = []
        for layer in bh['layers']:
            top_y = to_svg_y(bh['elevation'] - cum_thick)
            cum_thick += layer['thick']
            bottom_y = to_svg_y(bh['elevation'] - cum_thick)
            bh_layers.append((sx, top_y, bottom_y, layer))
        bh_boundaries.append(bh_layers)

    # Correlate layers across boreholes by formation name
    # Use the first borehole's layer order as reference
    ref_layers = bh_boundaries[0]
    all_formation_names = [l[3]['formation'] for l in ref_layers]

    # Draw each correlated layer as a polygon band between boreholes
    for li, fm_name in enumerate(all_formation_names):
        # Collect top/bottom points across all boreholes for this formation
        top_points = []
        bot_points = []
        for bh_idx, bh_layers in enumerate(bh_boundaries):
            # Find matching layer by index position (fallback if name doesn't match)
            if li < len(bh_layers):
                sx, top_y, bottom_y, layer = bh_layers[li]
                top_points.append((sx, top_y))
                bot_points.append((sx, bottom_y))

        if len(top_points) >= 2:
            # Build polygon: top line L->R, then bottom line R->L
            pts = []
            for sx, sy in top_points:
                pts.append(f'{sx:.1f},{sy:.1f}')
            for sx, sy in reversed(bot_points):
                pts.append(f'{sx:.1f},{sy:.1f}')

            layer = ref_layers[li][3]
            c, m, y, k = layer['c'], layer['m'], layer['y'], layer['k']
            fill_color = cmyk_fill(c, m, y, k)
            pattern_name = layer.get('pattern', 'pure')

            layer_g = ET.SubElement(body, 'g', {
                'data-cdr-type': 'layer-band',
                'data-cdr-name': layer['formation'],
                'data-cdr-pattern': pattern_name
            })

            # Filled polygon
            pts_str = ' '.join(pts)
            ET.SubElement(layer_g, 'polygon', {
                'points': pts_str,
                'fill': fill_color,
                'stroke': '#444',
                'stroke-width': '0.3',
                'data-cdr-type': 'layer-fill'
            })

            # Pattern overlay (simplified — pattern the middle region)
            if pattern_name != 'pure' and pattern_name in PATTERNS:
                # Create clip path for this layer band
                clip_id = f'clip_xsec_{li}'
                clip = ET.SubElement(defs, 'clipPath', {'id': clip_id})
                ET.SubElement(clip, 'polygon', {'points': pts_str})
                pat_g = ET.SubElement(layer_g, 'g', {'clip-path': f'url(#{clip_id})'})
                min_x_svg = min(p[0] for p in top_points)
                max_x_svg = max(p[0] for p in top_points)
                min_y_svg = min(p[1] for p in top_points)
                max_y_svg = max(p[1] for p in bot_points)
                band_w = max_x_svg - min_x_svg
                band_h = max_y_svg - min_y_svg
                if band_w > 0 and band_h > 0:
                    # Adaptive spacing: wider bands get sparser patterns
                    global PAT_SPACING
                    saved_sp = PAT_SPACING
                    scale_factor = max(1.0, band_w / 200)
                    PAT_SPACING = saved_sp * scale_factor
                    PATTERNS[pattern_name](pat_g, min_x_svg, min_y_svg, band_w, band_h)
                    PAT_SPACING = saved_sp

            # Formation label at midpoint
            mid_x = sum(p[0] for p in top_points) / len(top_points)
            mid_y = (sum(p[1] for p in top_points) + sum(p[1] for p in bot_points)) / (2 * len(top_points))
            fm_label = layer['formation']
            add_text(layer_g, mid_x, mid_y, fm_label, size=2.5, anchor='middle',
                     fill='#222', data_cdr_type='formation-label')

    # ============================
    # Faults
    # ============================
    for fault in faults:
        fx = to_svg_x(fault['x'])
        f_top = to_svg_y(max_elev)
        f_bot = to_svg_y(min_bottom)
        dip_angle = fault.get('dip', 70)
        throw = fault.get('throw', 0) * v_scale / vx  # convert to SVG units
        ftype = fault.get('type', 'normal')

        # Fault line (dipping)
        # dip angle from horizontal: steeper = more vertical
        dip_rad = math.radians(90 - dip_angle)
        fx_bot = fx + (f_bot - f_top) * math.tan(dip_rad) * 0.3  # slight offset

        add_line(body, fx, f_top, fx_bot, f_bot,
                 stroke='#c0392b', sw=1.0, data_cdr_type='fault-line')

        # Arrows showing throw direction
        if throw != 0:
            arr_y = f_top + (f_bot - f_top) * 0.5
            arr_dir = 1 if ftype == 'normal' else -1
            # Down-arrow on hanging wall side
            add_line(body, fx + 8, arr_y, fx + 8, arr_y + throw * arr_dir,
                     stroke='#c0392b', sw=0.6, data_cdr_type='fault-throw')
            # Arrowhead
            ah_y = arr_y + throw * arr_dir
            add_line(body, fx + 5, ah_y - 3 * arr_dir, fx + 8, ah_y,
                     stroke='#c0392b', sw=0.5)
            add_line(body, fx + 11, ah_y - 3 * arr_dir, fx + 8, ah_y,
                     stroke='#c0392b', sw=0.5)

        # Fault label
        f_label = f"{ftype} fault\ndip {fault['dip']}° {fault.get('direction','')}"
        lines = f_label.split('\n')
        for lii, line_text in enumerate(lines):
            add_text(body, fx + 12, f_top + 15 + lii * 6, line_text,
                     size=2.2, fill='#c0392b', data_cdr_type='fault-label')

    # ============================
    # Scale bar
    # ============================
    scale_bar_x = MARGIN
    scale_bar_y = page_h - BOTTOM_MARGIN + 15
    scale_len_m = 100  # 100m horizontal scale bar
    scale_len_svg = scale_len_m * h_scale

    add_line(body, scale_bar_x, scale_bar_y, scale_bar_x + scale_len_svg, scale_bar_y,
             stroke='#000', sw=0.8, data_cdr_type='scale-bar')
    add_line(body, scale_bar_x, scale_bar_y - 4, scale_bar_x, scale_bar_y + 4,
             stroke='#000', sw=0.5)
    add_line(body, scale_bar_x + scale_len_svg, scale_bar_y - 4, scale_bar_x + scale_len_svg, scale_bar_y + 4,
             stroke='#000', sw=0.5)
    add_text(body, scale_bar_x + scale_len_svg / 2, scale_bar_y - 5,
             f"{scale_len_m}m", size=2.8, anchor='middle', fill='#000',
             data_cdr_type='scale-label')

    # Vertical scale
    v_scale_bar_x = scale_bar_x + scale_len_svg + 30
    v_scale_len_m = 50  # 50m vertical
    v_scale_len_svg = v_scale_len_m * v_scale / vx
    add_line(body, v_scale_bar_x, scale_bar_y, v_scale_bar_x, scale_bar_y - v_scale_len_svg,
             stroke='#000', sw=0.8, data_cdr_type='vscale-bar')
    add_line(body, v_scale_bar_x - 4, scale_bar_y, v_scale_bar_x + 4, scale_bar_y,
             stroke='#000', sw=0.5)
    add_line(body, v_scale_bar_x - 4, scale_bar_y - v_scale_len_svg, v_scale_bar_x + 4, scale_bar_y - v_scale_len_svg,
             stroke='#000', sw=0.5)
    add_text(body, v_scale_bar_x + 8, scale_bar_y - v_scale_len_svg / 2,
             f"{v_scale_len_m}m", size=2.8, anchor='start', fill='#000',
             data_cdr_type='vscale-label')

    # ============================
    # Legend
    # ============================
    legend_g = ET.SubElement(svg, 'g', {'id': 'cdr-legend', 'data-cdr-layer': 'legend'})
    # Collect unique formations from all boreholes
    unique_fms = {}
    for bh in boreholes:
        for layer in bh['layers']:
            fm = layer['formation']
            if fm not in unique_fms:
                unique_fms[fm] = layer

    if unique_fms:
        lgX = page_w - 150
        lgY = TOP_MARGIN + 10
        lgW = 130
        itemH = 12
        nFm = len(unique_fms)
        lgH = nFm * itemH + 20

        add_rect(legend_g, lgX, lgY, lgW, lgH,
                 fill='#fff', stroke='#666', sw=0.3, data_cdr_type='legend-box')
        add_text(legend_g, lgX + lgW/2, lgY + 8, 'Legend',
                 size=3.5, bold=True, anchor='middle', fill='#000',
                 data_cdr_type='legend-title')
        add_line(legend_g, lgX + 5, lgY + 14, lgX + lgW - 5, lgY + 14,
                 stroke='#ccc', sw=0.15)

        for j, (fm, ldata) in enumerate(unique_fms.items()):
            iy = lgY + 20 + j * itemH
            c, m, y, k = ldata['c'], ldata['m'], ldata['y'], ldata['k']
            fc = cmyk_fill(c, m, y, k)
            swatchX, swatchY = lgX + 8, iy - 4.5
            add_rect(legend_g, swatchX, swatchY, 9, 9, fill=fc, stroke='#888', sw=0.15,
                     data_cdr_type='legend-swatch')
            add_text(legend_g, lgX + 20, iy, fm, size=2.5, fill='#333',
                     data_cdr_type='legend-label')

    # Footer
    footer_g = ET.SubElement(svg, 'g', {'id': 'cdr-footer', 'data-cdr-layer': 'footer'})
    n_bh = len(boreholes)
    footer_text = f"{title}  |  {n_bh} boreholes  |  V.E. = {vx}×"
    add_text(footer_g, MARGIN, page_h - 8, footer_text, size=2.2, fill='#888',
             data_cdr_type='footer-text')

    # Output
    rough = ET.tostring(svg, encoding='unicode')
    try:
        dom = minidom.parseString(rough)
        pretty = dom.toprettyxml(indent='  ', encoding='UTF-8')
        result = pretty.decode('utf-8')
    except:
        result = rough

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(result)
        return output_path
    return result


# ============================================================================
# CLI
# ============================================================================

def main():
    if '--help' in sys.argv or '-h' in sys.argv:
        print("""
generate_cross_section.py — Geological Cross-Section SVG Generator v1.0

Usage:
  python3 generate_cross_section.py                          use built-in demo
  python3 generate_cross_section.py data.json output.svg     generate from JSON
  python3 generate_cross_section.py --json '{"title":"...","boreholes":[...]}' out.svg

Output: SVG vector graphic with 5 layer groups, data-cdr-* attributes,
        surface profile, layer correlation bands, fault rendering,
        dual scale bars, 18 lithology patterns, legend.
""")
        return

    data = DEMO_DATA
    output_path = None

    args = sys.argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a == '--json' and i + 1 < len(args):
            data = json.loads(args[i + 1])
            i += 2
        elif a.endswith('.json'):
            with open(a, 'r', encoding='utf-8') as f:
                data = json.load(f)
            i += 1
        elif a.endswith('.svg'):
            output_path = a
            i += 1
        else:
            i += 1

    if output_path:
        generate_cross_section(data, output_path)
        print(f"✅ Cross-section saved: {output_path}")
    else:
        result = generate_cross_section(data)
        print(result)


if __name__ == '__main__':
    main()
