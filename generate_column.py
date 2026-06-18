#!/usr/bin/env python3
"""
地层柱状图 SVG 矢量图生成器
══════════════════════════════════════════════════════════
依据：中国地质大学(武汉)秭归基地《常用地层图例花纹和符号》
标准：GB/T 958 区域地质图图例
输出：SVG 矢量图（可直接导入 CorelDRAW、Illustrator、浏览器查看）

用法：
  python3 generate_column.py data.json output.svg
  python3 generate_column.py                    # 使用内置示例数据
  python3 generate_column.py --json '{...}'     # 直接传 JSON

依赖：无（纯 Python 标准库）
══════════════════════════════════════════════════════════
"""

import json
import sys
import os
import xml.etree.ElementTree as ET
from xml.dom import minidom

# ============================================================================
# CONFIGURATION
# ============================================================================

DEFAULT_DATA = {
    "title": "秭归地区综合地层柱状图",
    "location": "湖北秭归",
    "layers": [
        {"erathem":"新元古界","system":"南华系","series":"下统","formation":"莲沱组","symbol":"Nh₁l","thick":120,"descr":"紫红色中厚层砂岩、含砾砂岩，底部为砾岩","c":0,"m":40,"y":30,"k":10,"pattern":"sand"},
        {"erathem":"新元古界","system":"南华系","series":"下统","formation":"南沱组","symbol":"Nh₁n","thick":45,"descr":"灰绿色冰碛砾岩、含砾砂质泥岩","c":15,"m":0,"y":20,"k":20,"pattern":"conglo"},
        {"erathem":"新元古界","system":"震旦系","series":"下统","formation":"陡山沱组","symbol":"Z₁d","thick":180,"descr":"灰黑色薄层泥质白云岩夹炭质页岩","c":5,"m":0,"y":10,"k":30,"pattern":"dolo"},
        {"erathem":"新元古界","system":"震旦系","series":"上统","formation":"灯影组","symbol":"Z₂dy","thick":350,"descr":"灰白色厚层白云岩，含燧石条带","c":0,"m":0,"y":5,"k":10,"pattern":"dolo"},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"水井沱组","symbol":"∈₁s","thick":95,"descr":"黑色炭质页岩夹薄层硅质岩","c":0,"m":5,"y":10,"k":65,"pattern":"carbShale"},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"石牌组","symbol":"∈₁sp","thick":150,"descr":"黄绿色粉砂质页岩、泥质粉砂岩","c":5,"m":0,"y":30,"k":15,"pattern":"silt"},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"天河板组","symbol":"∈₁t","thick":80,"descr":"深灰色中厚层鲕状灰岩","c":0,"m":0,"y":0,"k":35,"pattern":"lime"},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"石龙洞组","symbol":"∈₁sl","thick":110,"descr":"灰白色厚层白云岩","c":0,"m":0,"y":5,"k":8,"pattern":"dolo"},
        {"erathem":"古生界","system":"寒武系","series":"中上统","formation":"覃家庙组","symbol":"∈₂₋₃q","thick":210,"descr":"浅灰色薄-中厚层白云岩、泥质白云岩","c":0,"m":5,"y":15,"k":12,"pattern":"dolo"},
        {"erathem":"古生界","system":"寒武系","series":"上统","formation":"三游洞组","symbol":"∈₃s","thick":280,"descr":"灰白色厚层-块状白云岩","c":0,"m":0,"y":3,"k":8,"pattern":"dolo"},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"南津关组","symbol":"O₁n","thick":60,"descr":"深灰色生物碎屑灰岩","c":0,"m":0,"y":0,"k":30,"pattern":"lime"},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"分乡组","symbol":"O₁f","thick":45,"descr":"灰绿色页岩夹薄层灰岩","c":0,"m":0,"y":0,"k":25,"pattern":"shale"},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"红花园组","symbol":"O₁h","thick":35,"descr":"灰色厚层生物碎屑灰岩","c":0,"m":0,"y":0,"k":32,"pattern":"lime"},
        {"erathem":"古生界","system":"奥陶系","series":"中统","formation":"宝塔组","symbol":"O₂b","thick":20,"descr":"紫红色中厚层龟裂纹灰岩","c":0,"m":30,"y":20,"k":10,"pattern":"lime"}
    ]
}

# Layout (mm units → SVG user units)
MARGIN_LEFT   = 8
MARGIN_TOP    = 35    # top margin (was 30)
TABLE_TOP     = 290
TABLE_BOTTOM  = 22
HEADER_H      = 14
TABLE_RIGHT   = 203  # total table width (COL_X[8] + COL_W[8])

COL_X = [0,
    MARGIN_LEFT,                          # 1: 界
    MARGIN_LEFT + 14,                     # 2: 系
    MARGIN_LEFT + 28,                     # 3: 统
    MARGIN_LEFT + 42,                     # 4: 组
    MARGIN_LEFT + 62,                     # 5: 代号
    MARGIN_LEFT + 75,                     # 6: 柱状图
    MARGIN_LEFT + 110,                    # 7: 厚度
    MARGIN_LEFT + 122,                    # 8: 岩性描述
]
COL_W = [0, 14, 14, 14, 20, 13, 35, 12, 73]

PAT_SPACING = 3.2


# ============================================================================
# SVG HELPERS
# ============================================================================

def cmyk_fill(c, m, y, k):
    """Convert CMYK to RGB hex (rough approximation)."""
    r = int(255 * (1 - c/100) * (1 - k/100))
    g = int(255 * (1 - m/100) * (1 - k/100))
    b = int(255 * (1 - y/100) * (1 - k/100))
    return f"#{r:02x}{g:02x}{b:02x}"

def cmyk_rgb(c, m, y, k):
    """Return (r,g,b) tuple."""
    r = int(255 * (1 - c/100) * (1 - k/100))
    g = int(255 * (1 - m/100) * (1 - k/100))
    b = int(255 * (1 - y/100) * (1 - k/100))
    return (r, g, b)

def el(name, **attrs):
    """Create an XML element."""
    e = ET.Element(name)
    for k, v in attrs.items():
        e.set(k.replace('_', '-'), str(v))
    return e

def add_text(parent, x, y, text, size=2.8, bold=False, anchor='start', fill='#000'):
    """Add a <text> element with correct SVG attributes."""
    t = ET.SubElement(parent, 'text')
    t.set('x', str(x))
    t.set('y', str(y))
    t.set('font-family', 'SimHei, Heiti SC, sans-serif')
    t.set('font-size', str(size))
    t.set('font-weight', 'bold' if bold else 'normal')
    t.set('text-anchor', anchor)
    t.set('fill', fill)
    t.text = str(text)
    return t

def add_rect(parent, x, y, w, h, fill=None, stroke='#000', sw=0.2, **extra):
    """Add a <rect> element."""
    attrs = {'x': x, 'y': y, 'width': w, 'height': h}
    if fill:
        attrs['fill'] = fill
    else:
        attrs['fill'] = 'none'
    attrs['stroke'] = stroke
    attrs['stroke-width'] = str(sw)
    attrs.update(extra)
    return _append(parent, 'rect', attrs)

def add_line(parent, x1, y1, x2, y2, stroke='#000', sw=0.2):
    """Add a <line> element."""
    return _append(parent, 'line', {
        'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2,
        'stroke': stroke, 'stroke-width': str(sw)
    })

def add_circle(parent, cx, cy, r, fill='#000', stroke=None, sw=0):
    """Add a <circle> element."""
    a = {'cx': cx, 'cy': cy, 'r': r, 'fill': fill}
    if stroke:
        a['stroke'] = stroke
        a['stroke-width'] = str(sw)
    return _append(parent, 'circle', a)

def _append(parent, tag, attrs):
    e = ET.SubElement(parent, tag)
    for k, v in attrs.items():
        e.set(k.replace('_', '-'), str(v))
    return e


# ============================================================================
# PATTERN GENERATORS
# ============================================================================

def pattern_conglo(g, lx, by, cw, ch):
    """砾岩：大圆 + 散点"""
    sp = PAT_SPACING * 1.8
    # 大圆圈
    x = lx + sp * 0.7
    while x < lx + cw - sp * 0.3:
        y = by + sp * 0.7
        while y < by + ch - sp * 0.3:
            add_circle(g, x, y, 1.5, fill='none', stroke='#000', sw=0.2)
            y += sp
        x += sp
    # 散点
    sp2 = PAT_SPACING * 1.8
    x = lx + sp2 * 1.6
    while x < lx + cw:
        y = by + sp2 * 1.6
        while y < by + ch:
            add_circle(g, x, y, 0.4, fill='#555')
            y += sp2
        x += sp2

def pattern_sand(g, lx, by, cw, ch):
    """砂岩：密点"""
    sp = PAT_SPACING * 0.7
    x = lx + sp / 2
    while x < lx + cw:
        y = by + sp / 2
        while y < by + ch:
            add_circle(g, x, y, 0.35, fill='#555')
            y += sp
        x += sp

def pattern_finesand(g, lx, by, cw, ch):
    """细砂岩"""
    sp = PAT_SPACING * 0.55
    x = lx + sp / 2
    while x < lx + cw:
        y = by + sp / 2
        while y < by + ch:
            add_circle(g, x, y, 0.25, fill='#666')
            y += sp
        x += sp

def pattern_silt(g, lx, by, cw, ch):
    """粉砂岩：横线 + 稀疏点"""
    sp = PAT_SPACING
    y = by + sp * 0.5
    while y < by + ch:
        add_line(g, lx + 1, y, lx + cw - 1, y, stroke='#aaa', sw=0.2)
        y += sp
    x = lx + sp
    while x < lx + cw:
        y = by + sp
        while y < by + ch:
            add_circle(g, x, y, 0.2, fill='#888')
            y += sp * 2
        x += sp * 2

def pattern_mud(g, lx, by, cw, ch):
    """泥岩：密横线"""
    sp = PAT_SPACING * 0.6
    y = by + sp * 0.3
    while y < by + ch:
        add_line(g, lx + 0.5, y, lx + cw - 0.5, y, stroke='#555', sw=0.22)
        y += sp

def pattern_shale(g, lx, by, cw, ch):
    """页岩：横线 + 短竖线"""
    sp = PAT_SPACING * 0.7
    y = by + sp * 0.3
    cnt = 0
    while y < by + ch:
        add_line(g, lx + 0.3, y, lx + cw - 0.3, y, stroke='#555', sw=0.22)
        cnt += 1
        if cnt % 3 != 0:
            svx = lx + sp * 0.8
            while svx < lx + cw:
                svLen = sp * 0.3
                add_line(g, svx, y - svLen, svx, y + svLen, stroke='#999', sw=0.15)
                svx += sp * 1.5
        y += sp

def pattern_carbShale(g, lx, by, cw, ch):
    """炭质页岩：白虚线（配合深色背景）"""
    sp = PAT_SPACING * 0.7
    y = by + sp * 0.3
    while y < by + ch:
        add_line(g, lx + 0.3, y, lx + cw - 0.3, y, stroke='#eee', sw=0.18)
        y += sp

def pattern_lime(g, lx, by, cw, ch):
    """石灰岩：砖格"""
    sp = PAT_SPACING
    bw = sp * 1.5
    bh = sp * 0.6
    row = 0
    y = by + 0.3
    while y < by + ch:
        x = lx + 0.3 + (bw / 2 if row % 2 == 1 else 0)
        while x < lx + cw:
            add_rect(g, x, y, bw - 0.4, bh, fill='none', stroke='#666', sw=0.16)
            x += bw
        y += bh + 0.25
        row += 1

def pattern_dolo(g, lx, by, cw, ch):
    """白云岩：菱格交叉线"""
    sp = PAT_SPACING * 1.2
    dx = lx - (cw + ch)
    while dx < lx + cw + ch:
        add_line(g, dx, by, dx + ch, by + ch, stroke='#999', sw=0.15)
        dx += sp
    dx = lx - (cw + ch)
    while dx < lx + cw + ch:
        add_line(g, dx, by + ch, dx + ch, by, stroke='#999', sw=0.15)
        dx += sp

def pattern_doloLime(g, lx, by, cw, ch):
    """白云质灰岩：砖格+菱格混合"""
    sp = PAT_SPACING
    bw = sp * 2
    bh = sp * 0.7
    row = 0
    y = by + 0.3
    while y < by + ch:
        x = lx + 0.3 + (bw / 2 if row % 2 == 1 else 0)
        while x < lx + cw:
            add_rect(g, x, y, bw - 0.4, bh, fill='none', stroke='#bbb', sw=0.12)
            x += bw
        y += bh + 0.25
        row += 1
    dx = lx - (cw + ch)
    while dx < lx + cw + ch:
        add_line(g, dx, by, dx + ch, by + ch, stroke='#ccc', sw=0.1)
        dx += sp * 2

def pattern_chert(g, lx, by, cw, ch):
    """硅质岩：粗交叉线"""
    sp = PAT_SPACING
    dx = lx - (cw + ch)
    while dx < lx + cw + ch:
        add_line(g, dx, by, dx + ch, by + ch, stroke='#555', sw=0.25)
        dx += sp
    dx = lx - (cw + ch)
    while dx < lx + cw + ch:
        add_line(g, dx, by + ch, dx + ch, by, stroke='#555', sw=0.25)
        dx += sp

def pattern_coal(g, lx, by, cw, ch):
    """煤：全黑 + 白色层理线"""
    # Black fill is already in the rect
    add_line(g, lx + 1, by + ch / 3, lx + cw - 1, by + ch / 3, stroke='#fff', sw=0.3)
    add_line(g, lx + 1, by + ch * 2 / 3, lx + cw - 1, by + ch * 2 / 3, stroke='#fff', sw=0.3)

def pattern_granite(g, lx, by, cw, ch):
    """花岗岩：叉号 + 密点"""
    import random
    rng = random.Random(123)
    sp = PAT_SPACING
    gn = int(cw * ch / (sp * sp) * 0.8)
    for _ in range(gn):
        x = lx + rng.random() * cw
        y = by + rng.random() * ch
        gs = sp * 0.35
        add_line(g, x - gs, y - gs, x + gs, y + gs, stroke='#555', sw=0.18)
        add_line(g, x - gs, y + gs, x + gs, y - gs, stroke='#555', sw=0.18)
    x = lx + sp * 0.5
    while x < lx + cw:
        y = by + sp * 0.5
        while y < by + ch:
            add_circle(g, x, y, 0.3, fill='#555')
            y += sp
        x += sp

def pattern_basalt(g, lx, by, cw, ch):
    """玄武岩：V 字形斜排"""
    sp = PAT_SPACING * 0.9
    y = by + sp * 0.5
    while y < by + ch:
        x = lx + sp * 0.3
        while x < lx + cw:
            vh = sp * 0.4
            add_line(g, x, y - vh, x + vh, y, stroke='#555', sw=0.2)
            add_line(g, x + vh, y, x + vh * 2, y - vh, stroke='#555', sw=0.2)
            x += sp * 1.5
        y += sp

def pattern_schist(g, lx, by, cw, ch):
    """片岩：波浪折线"""
    sp = PAT_SPACING * 0.9
    dx = lx - cw
    while dx < lx + cw * 2:
        mx = dx + ch * 0.5
        my = by + ch * 0.5
        add_line(g, dx, by, mx, my, stroke='#666', sw=0.22)
        add_line(g, mx, my, dx + ch, by + ch, stroke='#666', sw=0.22)
        dx += sp

def pattern_gneiss(g, lx, by, cw, ch):
    """片麻岩：粗细条带交替"""
    sp = PAT_SPACING * 0.8
    y = by + sp * 0.2
    band = True
    while y < by + ch:
        sw_val = 0.35 if band else 0.12
        col = '#333' if band else '#bbb'
        add_line(g, lx + 0.3, y, lx + cw - 0.3, y, stroke=col, sw=sw_val)
        band = not band
        y += sp * (0.6 if band else 0.3)

def pattern_marble(g, lx, by, cw, ch):
    """大理岩：细网格"""
    sp = PAT_SPACING
    x = lx + sp * 0.5
    while x < lx + cw:
        add_line(g, x, by, x, by + ch, stroke='#ddd', sw=0.08)
        x += sp
    y = by + sp * 0.5
    while y < by + ch:
        add_line(g, lx, y, lx + cw, y, stroke='#ddd', sw=0.08)
        y += sp

# Pattern dispatch table
PATTERNS = {
    'conglo': pattern_conglo,
    'sand': pattern_sand,
    'finesand': pattern_finesand,
    'silt': pattern_silt,
    'mud': pattern_mud,
    'shale': pattern_shale,
    'carbShale': pattern_carbShale,
    'lime': pattern_lime,
    'dolo': pattern_dolo,
    'doloLime': pattern_doloLime,
    'chert': pattern_chert,
    'coal': pattern_coal,
    'granite': pattern_granite,
    'basalt': pattern_basalt,
    'schist': pattern_schist,
    'gneiss': pattern_gneiss,
    'marble': pattern_marble,
}


# ============================================================================
# MAIN SVG GENERATOR
# ============================================================================

def generate_svg(data, output_path=None):
    """
    Generate a stratigraphic column SVG from the given data dict.
    
    data format:
    {
        "title": "标题",
        "location": "地点",
        "layers": [
            {
                "erathem": "界", "system": "系", "series": "统",
                "formation": "组", "symbol": "代号",
                "thick": 厚度(m), "descr": "岩性描述",
                "c": 0, "m": 0, "y": 0, "k": 0,
                "pattern": "sand|dolo|lime|..."
            },
            ...
        ]
    }
    """
    layers = data['layers']
    title = data.get('title', '综合地层柱状图')
    location = data.get('location', '')
    
    total_thick = sum(l['thick'] for l in layers)
    n_layers = len(layers)
    
    draw_h = TABLE_TOP - TABLE_BOTTOM - HEADER_H
    scale_f = draw_h / total_thick if total_thick > 0 else 1
    
    # Page dimensions
    page_w = TABLE_RIGHT + 48   # 203 + 48 = 251, room for legend
    page_h = TABLE_TOP + MARGIN_TOP + 8   # 290 + 35 + 8 = 333
    
    # ===== Build SVG =====
    svg = ET.Element('svg', {
        'xmlns': 'http://www.w3.org/2000/svg',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'viewBox': f'0 0 {page_w} {page_h}',
        'width': f'{page_w}mm',
        'height': f'{page_h}mm',
        'version': '1.1'
    })
    
    # White background (CorelDRAW friendly)
    add_rect(svg, 0, 0, page_w, page_h, fill='#fff', stroke='none', sw=0)
    
    # ===== Defs (for clip paths) =====
    defs = ET.SubElement(svg, 'defs')
    
    # ===== Title =====
    title_x = (MARGIN_LEFT + TABLE_RIGHT) / 2
    add_text(svg, title_x, TABLE_TOP + 12, title,
             size=4.5, bold=True, anchor='middle', fill='#000')
    
    if location:
        add_text(svg, title_x, TABLE_TOP + 6, location,
                 size=2.8, bold=False, anchor='middle', fill='#555')
    
    # ===== Table Header =====
    headerT = TABLE_TOP
    headerB = TABLE_TOP - HEADER_H
    
    # Header background
    add_rect(svg, MARGIN_LEFT, headerB, TABLE_RIGHT - MARGIN_LEFT, HEADER_H,
             fill='#f5f5f5', stroke='#000', sw=0.4)
    
    header_labels = [
        ('界', 1), ('系', 2), ('统', 3), ('组', 4), ('代号', 5),
        ('柱 状 图', 6), ('厚度\n(m)', 7), ('岩 性 描 述', 8)
    ]
    for label, col in header_labels:
        cx = COL_X[col] + COL_W[col] / 2
        midY = (headerT + headerB) / 2
        lines = label.split('\n')
        for li, line_text in enumerate(lines):
            ly = midY - (len(lines) - 1) * 3.5 + li * 7
            add_text(svg, cx, ly, line_text, size=2.8 if col != 7 else 2.6,
                     bold=True, anchor='middle', fill='#000')
    
    # Vertical column lines
    for ci in range(2, 9):
        add_line(svg, COL_X[ci], TABLE_BOTTOM, COL_X[ci], TABLE_TOP,
                 stroke='#999', sw=0.2)
    
    # Header bottom line (thick)
    add_line(svg, MARGIN_LEFT, headerB, TABLE_RIGHT, headerB,
             stroke='#000', sw=0.4)
    
    # ===== Depth scale ticks =====
    tick_interval = 100
    if total_thick <= 50: tick_interval = 10
    elif total_thick <= 200: tick_interval = 50
    
    tickX = COL_X[6]
    tickRight = COL_X[6] + COL_W[6] - 2
    
    tick_val = 0
    while tick_val <= total_thick:
        tickY = headerB - tick_val * scale_f
        if tickY >= TABLE_BOTTOM:
            add_line(svg, tickX, tickY, tickRight, tickY, stroke='#aaa', sw=0.15)
            if tick_val % (tick_interval * 2) == 0:
                add_text(svg, tickRight - 1, tickY - 1.5, str(tick_val),
                         size=2.2, anchor='end', fill='#888')
        tick_val += tick_interval
    
    # Depth ruler vertical line
    add_line(svg, tickX, TABLE_BOTTOM, tickX, headerB, stroke='#ccc', sw=0.1)
    
    # ===== Pre-calculate layer Y positions =====
    layer_tops = []    # list of (topY, bottomY, height)
    currentY = headerB
    for layer in layers:
        layH = max(layer['thick'] * scale_f, 3.5)
        layBottom = currentY - layH
        layer_tops.append((currentY, layBottom, layH))
        currentY = layBottom
    
    # ===== Identify merge groups for columns 1-3 =====
    # Each group: (start_idx, end_idx, topY, bottomY, value)
    def find_merge_groups(col_key):
        groups = []
        if not layers: return groups
        start = 0
        for i in range(1, len(layers)):
            if layers[i].get(col_key, '') != layers[start].get(col_key, ''):
                groups.append((start, i-1,
                    layer_tops[start][0], layer_tops[i-1][1],
                    layers[start].get(col_key, '')))
                start = i
        groups.append((start, len(layers)-1,
            layer_tops[start][0], layer_tops[-1][1],
            layers[start].get(col_key, '')))
        return groups
    
    erathem_groups = find_merge_groups('erathem')
    system_groups = find_merge_groups('system')
    series_groups = find_merge_groups('series')
    
    # ===== Layer rows =====
    colLeft = COL_X[6] + 0.3
    colRight = COL_X[6] + COL_W[6] - 2.5
    colW_actual = colRight - colLeft
    
    # Draw merged cell backgrounds (界 / 系 / 统)
    for groups, col_idx, fill_hex in [
        (erathem_groups, 1, '#fafaf8'),
        (system_groups, 2, '#fafaf8'),
        (series_groups, 3, '#fafaf8')
    ]:
        for start, end, topY, bottomY, val in groups:
            if val:
                add_rect(svg, COL_X[col_idx], bottomY,
                         COL_W[col_idx], topY - bottomY,
                         fill=fill_hex, stroke='none', sw=0)
                midY = (topY + bottomY) / 2
                add_text(svg, COL_X[col_idx] + COL_W[col_idx]/2, midY, val,
                         size=2.4, bold=False, anchor='middle', fill='#222')
    
    # Draw each layer (组/代号/柱状图/厚度/描述)
    for i, layer in enumerate(layers):
        topY, bottomY, layH = layer_tops[i]
        
        c, m, y, k = layer['c'], layer['m'], layer['y'], layer['k']
        fill_color = cmyk_fill(c, m, y, k)
        pattern_name = layer.get('pattern', 'pure')
        
        # === Column rectangle (col 6) with pattern ===
        col_g = ET.SubElement(svg, 'g')
        col_g.set('id', f'layer_{i}')
        
        add_rect(col_g, colLeft, bottomY, colW_actual, layH,
                 fill=fill_color, stroke='#555', sw=0.15)
        
        if pattern_name != 'pure' and pattern_name in PATTERNS:
            clip_id = f'clip_{i}'
            clip = ET.SubElement(defs, 'clipPath', {'id': clip_id})
            ET.SubElement(clip, 'rect', {
                'x': str(colLeft), 'y': str(bottomY),
                'width': str(colW_actual), 'height': str(layH)
            })
            pat_g = ET.SubElement(col_g, 'g', {'clip-path': f'url(#{clip_id})'})
            PATTERNS[pattern_name](pat_g, colLeft, bottomY, colW_actual, layH)
        
        # === Text columns 4-5, 7-8 (per-layer, no merging) ===
        midY_ = (topY + bottomY) / 2
        
        # 组名
        fm = layer.get('formation', '')
        if fm:
            add_text(svg, COL_X[4] + 1, midY_, fm, size=2.2, fill='#222')
        # 代号
        sym = layer.get('symbol', '')
        if sym:
            add_text(svg, COL_X[5] + 1, midY_, sym, size=2.2, bold=True, fill='#000')
        # 厚度
        add_text(svg, COL_X[7] + 1, midY_, f"{layer['thick']:.1f}", size=2.2, fill='#222')
        # 描述
        d = layer.get('descr', '')
        if d:
            add_text(svg, COL_X[8] + 1, midY_, d, size=2, fill='#333')
        
        # Row separator (thin, all columns)
        add_line(svg, MARGIN_LEFT, topY, TABLE_RIGHT, topY,
                 stroke='#bbb', sw=0.12)
    
    # ===== Major boundary lines (界 level — thick) =====
    for start, end, topY, bottomY, val in erathem_groups:
        add_line(svg, MARGIN_LEFT, topY, TABLE_RIGHT, topY,
                 stroke='#000', sw=0.35)
    # 系 level — medium
    for start, end, topY, bottomY, val in system_groups:
        add_line(svg, MARGIN_LEFT + COL_W[1], topY, TABLE_RIGHT, topY,
                 stroke='#555', sw=0.22)
    # 统 level — thin (already drawn per-layer above)
    
    # ===== Bottom border =====
    add_line(svg, MARGIN_LEFT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_BOTTOM,
             stroke='#000', sw=0.4)
    
    # ===== Outer border =====
    add_line(svg, MARGIN_LEFT, TABLE_BOTTOM, MARGIN_LEFT, TABLE_TOP,
             stroke='#000', sw=0.4)
    add_line(svg, TABLE_RIGHT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_TOP,
             stroke='#000', sw=0.4)
    add_line(svg, MARGIN_LEFT, TABLE_TOP, TABLE_RIGHT, TABLE_TOP,
             stroke='#000', sw=0.4)
    
    # ===== Footer =====
    ratio = int(total_thick * 1000 / draw_h) if draw_h > 0 else 1
    footer_text = (f"{title}  |  总厚 {total_thick:,.0f}m  |  {n_layers} 层  "
                   f"|  比例尺 1:{ratio:,}")
    add_text(svg, MARGIN_LEFT, TABLE_BOTTOM - 5, footer_text,
             size=2, fill='#666')
    
    # ===== Legend (right side) =====
    # Collect unique patterns in stratigraphic order
    unique = {}
    for layer in layers:
        pn = layer.get('pattern', 'pure')
        if pn not in unique:
            unique[pn] = layer
    
    if len(unique) > 1:
        lgX = TABLE_RIGHT + 4          # 20
        lgY = TABLE_TOP - HEADER_H     # 276
        lgW = 38
        itemH = 7.0
        nUnique = len(unique)
        lgH = nUnique * itemH + 10     # top margin for title
        
        # Legend box
        add_rect(svg, lgX, lgY - lgH, lgW, lgH,
                 fill='#fff', stroke='#666', sw=0.25)
        
        # Legend title (inside box, at top)
        add_text(svg, lgX + lgW/2, lgY - lgH + 5.5, '图  例',
                 size=2.5, bold=True, anchor='middle', fill='#000')
        
        # Separator line under title
        add_line(svg, lgX + 3, lgY - lgH + 8, lgX + lgW - 3, lgY - lgH + 8,
                 stroke='#ccc', sw=0.15)
        
        # Legend items (top to bottom, in stratigraphic order)
        for j, (pn, ldata) in enumerate(unique.items()):
            iy = lgY - lgH + 9.5 + j * itemH  # first item center Y
            swatchY = iy - 3  # swatch top
            
            # Clip path for this swatch (prevent pattern overflow)
            clip_id = f'lgclip_{j}'
            clip = ET.SubElement(defs, 'clipPath', {'id': clip_id})
            ET.SubElement(clip, 'rect', {
                'x': str(lgX + 4), 'y': str(swatchY),
                'width': '6', 'height': '6'
            })
            
            # Color swatch (6x6mm)
            c, m, y_, k = ldata['c'], ldata['m'], ldata['y'], ldata['k']
            fc = cmyk_fill(c, m, y_, k)
            add_rect(svg, lgX + 4, swatchY, 6, 6, fill=fc, stroke='#888', sw=0.12)
            
            # Mini pattern (clipped!)
            if pn != 'pure' and pn in PATTERNS:
                pat_g = ET.SubElement(svg, 'g', {'clip-path': f'url(#{clip_id})'})
                PATTERNS[pn](pat_g, lgX + 4, swatchY, 6, 6)
            
            # Label text
            add_text(svg, lgX + 13, iy, ldata.get('formation', ''),
                     size=2, fill='#333')
    
    # ===== Pretty-print and output =====
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
    else:
        return result


# ============================================================================
# CLI
# ============================================================================

def main():
    data = DEFAULT_DATA
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
        generate_svg(data, output_path)
        print(f"✅ SVG saved: {output_path}")
    else:
        svg_content = generate_svg(data)
        print(svg_content)


if __name__ == '__main__':
    main()
