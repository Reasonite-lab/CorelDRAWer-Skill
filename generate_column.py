#!/usr/bin/env python3
"""
地层柱状图 SVG 矢量图生成器 v2.0
══════════════════════════════════════════════════════════
依据：中国地质大学(武汉)秭归基地《常用地层图例花纹和符号》
标准：GB/T 958 区域地质图图例
输出：SVG 矢量图（可直接导入 CorelDRAW、Illustrator、浏览器查看）

新增 v2.0:
- SVG 图层分组（cdr-background/header/body/scale/legend/footer）
- CorelDRAW data-cdr-* 属性
- 连续粒度曲线 + 渐变填充
- 化石符号列
- 沉积构造指示列
- 可选年龄(Ma)列

用法：
  python3 generate_column.py data.json output.svg
  python3 generate_column.py                    # 使用内置示例数据

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
        {"erathem":"新元古界","system":"南华系","series":"下统","formation":"莲沱组","symbol":"Nh₁l","thick":120,"descr":"紫红色中厚层砂岩、含砾砂岩，底部为砾岩","c":0,"m":40,"y":30,"k":10,"pattern":"sand","grain":4,"contact":"unconformity","age_ma":780},
        {"erathem":"新元古界","system":"南华系","series":"下统","formation":"南沱组","symbol":"Nh₁n","thick":45,"descr":"灰绿色冰碛砾岩、含砾砂质泥岩","c":15,"m":0,"y":20,"k":20,"pattern":"conglo","grain":6,"age_ma":720},
        {"erathem":"新元古界","system":"震旦系","series":"下统","formation":"陡山沱组","symbol":"Z₁d","thick":180,"descr":"灰黑色薄层泥质白云岩夹炭质页岩","c":5,"m":0,"y":10,"k":30,"pattern":"dolo","grain":2,"fossils":["algae"],"age_ma":635},
        {"erathem":"新元古界","system":"震旦系","series":"上统","formation":"灯影组","symbol":"Z₂dy","thick":350,"descr":"灰白色厚层白云岩，含燧石条带","c":0,"m":0,"y":5,"k":10,"pattern":"dolo","grain":2,"fossils":["stromatolite"],"age_ma":551},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"水井沱组","symbol":"∈₁s","thick":95,"descr":"黑色炭质页岩夹薄层硅质岩","c":0,"m":5,"y":10,"k":65,"pattern":"carbShale","grain":1,"fossils":["trilobite","brachiopod"],"age_ma":541},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"石牌组","symbol":"∈₁sp","thick":150,"descr":"黄绿色粉砂质页岩、泥质粉砂岩","c":5,"m":0,"y":30,"k":15,"pattern":"silt","grain":3,"fossils":["trilobite"],"structures":["ripple"],"age_ma":521},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"天河板组","symbol":"∈₁t","thick":80,"descr":"深灰色中厚层鲕状灰岩","c":0,"m":0,"y":0,"k":35,"pattern":"lime","grain":3,"fossils":["trilobite","brachiopod"],"structures":["ooid"],"age_ma":509},
        {"erathem":"古生界","system":"寒武系","series":"下统","formation":"石龙洞组","symbol":"∈₁sl","thick":110,"descr":"灰白色厚层白云岩","c":0,"m":0,"y":5,"k":8,"pattern":"dolo","grain":2,"age_ma":500},
        {"erathem":"古生界","system":"寒武系","series":"中上统","formation":"覃家庙组","symbol":"∈₂₋₃q","thick":210,"descr":"浅灰色薄-中厚层白云岩、泥质白云岩","c":0,"m":5,"y":15,"k":12,"pattern":"dolo","grain":2,"age_ma":497},
        {"erathem":"古生界","system":"寒武系","series":"上统","formation":"三游洞组","symbol":"∈₃s","thick":280,"descr":"灰白色厚层-块状白云岩","c":0,"m":0,"y":3,"k":8,"pattern":"dolo","grain":2,"age_ma":485},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"南津关组","symbol":"O₁n","thick":60,"descr":"深灰色生物碎屑灰岩","c":0,"m":0,"y":0,"k":30,"pattern":"lime","grain":2,"fossils":["brachiopod","crinoid"],"age_ma":478},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"分乡组","symbol":"O₁f","thick":45,"descr":"灰绿色页岩夹薄层灰岩","c":0,"m":0,"y":0,"k":25,"pattern":"shale","grain":1,"fossils":["graptolite"],"age_ma":472},
        {"erathem":"古生界","system":"奥陶系","series":"下统","formation":"红花园组","symbol":"O₁h","thick":35,"descr":"灰色厚层生物碎屑灰岩","c":0,"m":0,"y":0,"k":32,"pattern":"lime","grain":2,"fossils":["cephalopod"],"age_ma":468},
        {"erathem":"古生界","system":"奥陶系","series":"中统","formation":"宝塔组","symbol":"O₂b","thick":20,"descr":"紫红色中厚层龟裂纹灰岩","c":0,"m":30,"y":20,"k":10,"pattern":"lime","grain":2,"fossils":["cephalopod","trilobite"],"structures":["crack"],"age_ma":458}
    ]
}

# ============================================================================
# LAYOUT — extended 11-column format
# ============================================================================
# Columns: 1界 2系 3统 4组 5代号 6化石 7柱状图 8粒度 9构造 10厚度 11描述
#                          flexible widths based on available data

MARGIN_LEFT   = 8
MARGIN_TOP    = 35
TABLE_TOP     = 290
TABLE_BOTTOM  = 22
HEADER_H      = 14

# Default column layout — may be adjusted dynamically
COL_X = [0,
    MARGIN_LEFT,                          # 1: 界
    MARGIN_LEFT + 13,                     # 2: 系
    MARGIN_LEFT + 26,                     # 3: 统
    MARGIN_LEFT + 39,                     # 4: 组
    MARGIN_LEFT + 57,                     # 5: 代号
    MARGIN_LEFT + 69,                     # 6: 化石
    MARGIN_LEFT + 76,                     # 7: 柱状图
    MARGIN_LEFT + 116,                    # 8: 粒度
    MARGIN_LEFT + 124,                    # 9: 构造
    MARGIN_LEFT + 131,                    # 10: 厚度
    MARGIN_LEFT + 141,                    # 11: 描述
]
COL_W = [0,   13,   13,   13,   18,   12,    7,   40,    8,    7,   10,   65]

# Will be recalculated based on data
TABLE_RIGHT = MARGIN_LEFT + 141 + 65

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

def add_text(parent, x, y, text, size=2.8, bold=False, anchor='start', fill='#000', **extra):
    """Add a <text> element with correct SVG attributes."""
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

def add_line(parent, x1, y1, x2, y2, stroke='#000', sw=0.2, **extra):
    """Add a <line> element."""
    a = {'x1': x1, 'y1': y1, 'x2': x2, 'y2': y2, 'stroke': stroke, 'stroke-width': str(sw)}
    a.update(extra)
    return _append(parent, 'line', a)

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
# PATTERN GENERATORS (18 种国标花纹)
# ============================================================================

def pattern_conglo(g, lx, by, cw, ch):
    """砾岩：大圆 + 散点"""
    sp = PAT_SPACING * 1.8
    x = lx + sp * 0.7
    while x < lx + cw - sp * 0.3:
        y = by + sp * 0.7
        while y < by + ch - sp * 0.3:
            add_circle(g, x, y, 1.5, fill='none', stroke='#000', sw=0.2)
            y += sp
        x += sp
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
# FOSSIL / STRUCTURE SYMBOLS
# ============================================================================

FOSSIL_ICONS = {
    'trilobite':     '🦞',   # trilobite
    'brachiopod':    '🦪',   # brachiopod
    'cephalopod':    '🐙',   # cephalopod
    'graptolite':    '🪶',   # graptolite
    'crinoid':       '🌸',   # crinoid
    'algae':         '🌿',   # algae
    'stromatolite':  '🫧',   # stromatolite
    'ammonite':      '🐚',   # ammonite
    'coral':         '🪸',   # coral
    'bivalve':       '🦪',   # bivalve
    'gastropod':     '🐌',   # gastropod
    'foraminifera':  '⭕',   # foraminifera
    'plant':         '🌱',   # plant
    'fish':          '🐟',   # fish
    'spore':         '🔸',   # spore/pollen
}

STRUCTURE_ICONS = {
    'ripple':      '≈',   # ripple marks
    'cross_bed':   '⧵',   # cross bedding
    'graded':      '▽',   # graded bedding
    'ooid':        '◎',   # ooid
    'crack':       '⚡',   # mud crack
    'concretion':  '◉',   # concretion
    'bioturbation':'〰',   # bioturbation
    'stylolite':   '⏚',   # stylolite
    'stromatactis':'◈',   # stromatactis
}

def has_any_fossil(layers):
    for l in layers:
        if l.get('fossils'):
            return True
    return False

def has_any_structure(layers):
    for l in layers:
        if l.get('structures'):
            return True
    return False

def has_any_age(layers):
    for l in layers:
        if l.get('age_ma') is not None:
            return True
    return False


# ============================================================================
# MAIN SVG GENERATOR
# ============================================================================

def generate_svg(data, output_path=None):
    """
    Generate a stratigraphic column SVG from the given data dict.

    Extended fields per layer:
    - grain: 1-6 (clay to gravel) or grain_profile: [(pos, level), ...]
    - fossils: ["trilobite", "brachiopod", ...]
    - structures: ["ripple", "cross_bed", ...]
    - contact: "disconformity" | "unconformity" | "conformity"
    - markers: [{"symbol":"star","y_offset":0.5,"label":"S1"}]
    - age_ma: numeric age in millions of years
    """
    layers = data['layers']
    title = data.get('title', '综合地层柱状图')
    location = data.get('location', '')

    total_thick = sum(l['thick'] for l in layers)
    n_layers = len(layers)

    draw_h = TABLE_TOP - TABLE_BOTTOM - HEADER_H
    scale_f = draw_h / total_thick if total_thick > 0 else 1

    # ============================================
    # ADAPTIVE LAYOUT — show/hide columns dynamically
    # ============================================
    show_fossils = has_any_fossil(layers)
    show_structures = has_any_structure(layers)
    show_age = has_any_age(layers)

    # Dynamic column X/W calculations
    # Base columns: 界/系/统/组/代号 — always present
    # Optional: 化石/年龄 — added if data present
    # Fixed: 柱状图/粒度/构造/厚度/描述

    cur_x = MARGIN_LEFT
    col_defs = []  # list of (label, width, optional?)

    # Column 1: 界
    col_defs.append(('界', 13, False))
    # Column 2: 系
    col_defs.append(('系', 13, False))
    # Column 3: 统
    col_defs.append(('统', 13, False))
    # Column 4: 组
    col_defs.append(('组', 18, False))
    # Column 5: 代号
    col_defs.append(('代号', 12, False))
    # Column 6: 化石 (optional)
    col_defs.append(('化石', 7 if show_fossils else 0, not show_fossils))
    # Column 7: 柱状图
    col_defs.append(('岩 性 柱', 40, False))
    # Column 8: 粒度
    col_defs.append(('粒度', 15, False))  # wider for curve
    # Column 9: 构造 (optional)
    col_defs.append(('构造', 7 if show_structures else 0, not show_structures))
    # Column 10: 厚度
    col_defs.append(('厚度\n(m)', 10, False))
    # Column 11: 描述
    col_defs.append(('岩 性 描 述', 65, False))

    # Recalculate COL_X and COL_W from col_defs
    global COL_X, COL_W, TABLE_RIGHT
    COL_X = [0]
    COL_W = [0]
    for label, w, opt in col_defs:
        if w > 0 or not opt:
            COL_X.append(cur_x)
            w_actual = max(w, 0)
            COL_W.append(w_actual)
            cur_x += w_actual

    TABLE_RIGHT = cur_x

    # Page dimensions
    page_w = TABLE_RIGHT + 48
    page_h = TABLE_TOP + MARGIN_TOP + 8

    # ============================================
    # SVG ROOT
    # ============================================
    svg = ET.Element('svg', {
        'xmlns': 'http://www.w3.org/2000/svg',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'viewBox': f'0 0 {page_w} {page_h}',
        'width': f'{page_w}mm',
        'height': f'{page_h}mm',
        'version': '1.1'
    })

    defs = ET.SubElement(svg, 'defs')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-background
    # ═══════════════════════════════════════════
    bg_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-background',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Background',
        'data-cdr-layer': 'background'
    })
    add_rect(bg_grp, 0, 0, page_w, page_h, fill='#fff', stroke='none', sw=0)

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-title
    # ═══════════════════════════════════════════
    title_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-title',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Title',
        'data-cdr-layer': 'title'
    })
    title_x = (MARGIN_LEFT + TABLE_RIGHT) / 2
    add_text(title_grp, title_x, TABLE_TOP + 12, title,
             size=4.5, bold=True, anchor='middle', fill='#000',
             data_cdr_type='title')
    if location:
        add_text(title_grp, title_x, TABLE_TOP + 6, location,
                 size=2.8, bold=False, anchor='middle', fill='#555',
                 data_cdr_type='subtitle')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-header
    # ═══════════════════════════════════════════
    header_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-header',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Header',
        'data-cdr-layer': 'header'
    })

    headerT = TABLE_TOP
    headerB = TABLE_TOP - HEADER_H

    add_rect(header_grp, MARGIN_LEFT, headerB, TABLE_RIGHT - MARGIN_LEFT, HEADER_H,
             fill='#f5f5f5', stroke='#000', sw=0.4, data_cdr_type='header-bg')

    # Map col_defs to actual column indices
    # The actual col index in COL_X depends on which optional cols are hidden
    def get_col_n(label_map):
        """Find the actual 1-based column index for a given label string."""
        # col_defs may have gaps (w=0 for hidden cols)
        # COL_X has entries for all non-hidden cols
        # We need to map original positions to actual COL_X indices
        pass

    # Actually, let me build the header labels from col_defs directly
    real_col_idx = 0
    for label, w, opt in col_defs:
        if opt and w == 0:
            continue  # hidden optional column
        real_col_idx += 1
        cx = COL_X[real_col_idx] + COL_W[real_col_idx] / 2
        midY = (headerT + headerB) / 2
        lines = label.split('\n')
        for li, line_text in enumerate(lines):
            ly = midY - (len(lines) - 1) * 3.5 + li * 7
            size = 2.8
            if '粒度' in label:
                size = 2.4
            add_text(header_grp, cx, ly, line_text, size=size,
                     bold=True, anchor='middle', fill='#000',
                     data_cdr_type='header-text')

    # Vertical column lines
    for ci in range(2, len(COL_X)):
        if ci >= len(COL_X):
            break
        add_line(header_grp, COL_X[ci], TABLE_BOTTOM, COL_X[ci], TABLE_TOP,
                 stroke='#999', sw=0.2, data_cdr_type='col-divider')

    # Header bottom line
    add_line(header_grp, MARGIN_LEFT, headerB, TABLE_RIGHT, headerB,
             stroke='#000', sw=0.4, data_cdr_type='header-line')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-body (main column content)
    # ═══════════════════════════════════════════
    body_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-body',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Column Body',
        'data-cdr-layer': 'body'
    })

    # ===== Depth scale ticks =====
    tick_interval = 100
    if total_thick <= 50: tick_interval = 10
    elif total_thick <= 200: tick_interval = 50

    # The lithology column is where scale goes
    # Find col index for lithology column (岩性柱)
    lith_col_idx = None
    ri = 0
    for label, w, opt in col_defs:
        if not (opt and w == 0):
            ri += 1
        if label == '岩 性 柱':
            lith_col_idx = ri
            break

    tickX = COL_X[lith_col_idx]
    tickRight = COL_X[lith_col_idx] + COL_W[lith_col_idx] - 2

    tick_val = 0
    while tick_val <= total_thick:
        tickY = headerB - tick_val * scale_f
        if tickY >= TABLE_BOTTOM:
            add_line(body_grp, tickX, tickY, tickRight, tickY,
                     stroke='#aaa', sw=0.15, data_cdr_type='scale-tick')
            if tick_val % (tick_interval * 2) == 0:
                add_text(body_grp, tickRight - 1, tickY - 1.5, str(tick_val),
                         size=2.2, anchor='end', fill='#888', data_cdr_type='scale-label')
        tick_val += tick_interval

    add_line(body_grp, tickX, TABLE_BOTTOM, tickX, headerB,
             stroke='#ccc', sw=0.1, data_cdr_type='scale-ruler')

    # ===== Pre-calculate layer Y positions =====
    layer_tops = []
    currentY = headerB
    for layer in layers:
        layH = max(layer['thick'] * scale_f, 3.5)
        layBottom = currentY - layH
        layer_tops.append((currentY, layBottom, layH))
        currentY = layBottom

    # ===== Merge groups for columns 1-3 =====
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

    # Draw merged cell backgrounds
    for groups, col_key, fill_hex in [
        (erathem_groups, 'erathem', '#fafaf8'),
        (system_groups, 'system', '#fafaf8'),
        (series_groups, 'series', '#fafaf8')
    ]:
        # Find actual col_idx
        col_idx = {'erathem': 1, 'system': 2, 'series': 3}[col_key]
        for start, end, topY, bottomY, val in groups:
            if val:
                add_rect(body_grp, COL_X[col_idx], bottomY,
                         COL_W[col_idx], topY - bottomY,
                         fill=fill_hex, stroke='none', sw=0,
                         data_cdr_type='merge-cell',
                         data_cdr_field=col_key,
                         data_cdr_value=val)
                midY = (topY + bottomY) / 2
                add_text(body_grp, COL_X[col_idx] + COL_W[col_idx]/2, midY, val,
                         size=2.4, bold=False, anchor='middle', fill='#222',
                         data_cdr_type='merge-label',
                         data_cdr_field=col_key)

    # ===== Draw each layer =====
    # Find actual col indices for each column
    def find_col_idx(label_match):
        ri = 0
        for lbl, w, opt in col_defs:
            if not (opt and w == 0):
                ri += 1
            if label_match in lbl:
                return ri
        return None

    col_lith = find_col_idx('岩 性 柱')
    col_grain = find_col_idx('粒度')
    col_form = find_col_idx('组')
    col_sym = find_col_idx('代号')
    col_fossil = find_col_idx('化石')
    col_struct = find_col_idx('构造')
    col_thick = find_col_idx('厚度')
    col_descr = find_col_idx('描 述')

    colLeft = COL_X[col_lith] + 0.3
    colRight = COL_X[col_lith] + COL_W[col_lith] - 2.5
    colW_actual = colRight - colLeft

    # Grain size levels mapping
    grain_levels = {1: 'clay', 2: 'silt', 3: 'fine_sand',
                    4: 'medium_sand', 5: 'coarse_sand', 6: 'gravel'}
    grain_labels = {1: '粘土', 2: '粉砂', 3: '细砂',
                    4: '中砂', 5: '粗砂', 6: '砾石'}

    for i, layer in enumerate(layers):
        topY, bottomY, layH = layer_tops[i]

        c, m, y, k = layer['c'], layer['m'], layer['y'], layer['k']
        fill_color = cmyk_fill(c, m, y, k)
        pattern_name = layer.get('pattern', 'pure')

        # === Layer group with CorelDRAW metadata ===
        layer_g = ET.SubElement(body_grp, 'g', {
            'id': f'layer_{i}',
            'data-cdr-type': 'layer',
            'data-cdr-name': layer.get('formation', ''),
            'data-cdr-pattern': pattern_name,
            'data-cdr-thickness': str(layer['thick']),
            'data-cdr-color': fill_color,
            'data-cdr-index': str(i)
        })

        # === Lithology rectangle with pattern ===
        add_rect(layer_g, colLeft, bottomY, colW_actual, layH,
                 fill=fill_color, stroke='#555', sw=0.15,
                 data_cdr_type='lithology-rect')

        if pattern_name != 'pure' and pattern_name in PATTERNS:
            clip_id = f'clip_{i}'
            clip = ET.SubElement(defs, 'clipPath', {'id': clip_id})
            ET.SubElement(clip, 'rect', {
                'x': str(colLeft), 'y': str(bottomY),
                'width': str(colW_actual), 'height': str(layH)
            })
            pat_g = ET.SubElement(layer_g, 'g', {'clip-path': f'url(#{clip_id})'})
            PATTERNS[pattern_name](pat_g, colLeft, bottomY, colW_actual, layH)

        midY_ = (topY + bottomY) / 2

        # === Formation name (col 组) ===
        fm = layer.get('formation', '')
        if fm:
            add_text(layer_g, COL_X[col_form] + 1, midY_, fm,
                     size=2.2, fill='#222', data_cdr_type='formation-text')

        # === Symbol (col 代号) ===
        sym = layer.get('symbol', '')
        if sym:
            add_text(layer_g, COL_X[col_sym] + 1, midY_, sym,
                     size=2.2, bold=True, fill='#000', data_cdr_type='symbol-text')

        # === Fossil column ===
        if col_fossil and show_fossils:
            fossils = layer.get('fossils', [])
            if fossils:
                fcx = COL_X[col_fossil] + COL_W[col_fossil] / 2
                # Show up to 2 fossil icons stacked
                for fi, fname in enumerate(fossils[:3]):
                    if fname in FOSSIL_ICONS:
                        fy = midY_ + (fi - min(len(fossils[:3]), 2) / 2 + 0.5) * 3.5
                        add_text(layer_g, fcx, fy, FOSSIL_ICONS[fname],
                                 size=3.5, anchor='middle', fill='#333',
                                 data_cdr_type='fossil-icon',
                                 data_cdr_fossil=fname)

        # === Grain size column (col 粒度) — enhanced curve ===
        if col_grain:
            grain_profile = layer.get('grain_profile')
            grain_val = layer.get('grain')

            if grain_profile and len(grain_profile) >= 2:
                # Draw continuous curve from grain_profile
                gsx_base = COL_X[col_grain] + 1
                gsw = COL_W[col_grain] - 2
                lay_top_Y = topY
                lay_bot_Y = bottomY

                # Build polygon points for filled curve
                points_left = []
                points_right = []
                for (pos, level) in grain_profile:
                    gx = gsx_base + gsw * (level - 1) / 5
                    gy = lay_top_Y - pos * layH  # pos is 0-1 fraction from top
                    points_left.append(f'{gx:.1f},{gy:.1f}')
                    points_right.insert(0, f'{gsx_base:.1f},{gy:.1f}')

                all_points = points_left + points_right
                if all_points:
                    pts_str = ' '.join(all_points)
                    ET.SubElement(layer_g, 'polygon', {
                        'points': pts_str,
                        'fill': '#c0c0c0',
                        'opacity': '0.4',
                        'stroke': '#888',
                        'stroke-width': '0.2',
                        'data-cdr-type': 'grain-curve'
                    })
            elif grain_val:
                # Simple triangular indicator (existing behavior)
                gsx = COL_X[col_grain] + 1.5
                gsw = COL_W[col_grain] - 3
                grain_w = gsw * grain_val / 6
                pts = f'{gsx},{bottomY} {gsx+grain_w:.1f},{bottomY} {gsx+grain_w:.1f},{topY} {gsx},{topY}'
                ET.SubElement(layer_g, 'polygon', {
                    'points': pts,
                    'fill': '#aaa',
                    'opacity': '0.5',
                    'data-cdr-type': 'grain-triangle'
                })

            # Grain size label
            if grain_val and grain_val in grain_labels:
                gcx = COL_X[col_grain] + COL_W[col_grain] / 2
                add_text(layer_g, gcx, midY_, grain_labels[grain_val],
                         size=1.8, anchor='middle', fill='#999',
                         data_cdr_type='grain-label')

        # === Structure column ===
        if col_struct and show_structures:
            structures = layer.get('structures', [])
            if structures:
                scx = COL_X[col_struct] + COL_W[col_struct] / 2
                for si, sname in enumerate(structures[:2]):
                    if sname in STRUCTURE_ICONS:
                        sy = midY_ + (si - 0.25) * 4
                        add_text(layer_g, scx, sy, STRUCTURE_ICONS[sname],
                                 size=3.5, anchor='middle', fill='#c0392b',
                                 data_cdr_type='structure-icon',
                                 data_cdr_structure=sname)

        # === Thickness ===
        add_text(layer_g, COL_X[col_thick] + 1, midY_,
                 f"{layer['thick']:.1f}", size=2.2, fill='#222',
                 data_cdr_type='thickness-text')

        # === Description ===
        descr = layer.get('descr', '')
        if descr:
            add_text(layer_g, COL_X[col_descr] + 1, midY_, descr,
                     size=2, fill='#333', data_cdr_type='description-text')

        # === Contact symbol (at bottom of layer) ===
        contact = layer.get('contact', '')
        if contact and i < n_layers - 1:
            cx = colLeft + colW_actual / 2
            cy = bottomY
            if contact == 'disconformity':
                d_path = f'M {cx-4},{cy} Q {cx-2},{cy-1} {cx},{cy} Q {cx+2},{cy+1} {cx+4},{cy}'
                ET.SubElement(layer_g, 'path', {
                    'd': d_path, 'fill': 'none',
                    'stroke': '#c0392b', 'stroke-width': '0.3',
                    'data-cdr-type': 'contact-disconformity'
                })
            elif contact == 'unconformity':
                saw = f'M {cx-4},{cy-1.5} L {cx-2},{cy+1} L {cx},{cy-1.5} L {cx+2},{cy+1} L {cx+4},{cy-1.5}'
                ET.SubElement(layer_g, 'path', {
                    'd': saw, 'fill': 'none',
                    'stroke': '#c0392b', 'stroke-width': '0.35',
                    'data-cdr-type': 'contact-unconformity'
                })

        # === Sample markers ===
        markers = layer.get('markers', [])
        for mk in markers:
            mk_y = bottomY + mk.get('y_offset', 0.5) * layH
            mk_sym = mk.get('symbol', 'dot')
            mk_label = mk.get('label', '')
            mk_x = colLeft - 3
            if mk_sym == 'star':
                add_text(layer_g, mk_x, mk_y, '★', size=3, fill='#c0392b',
                         anchor='middle', bold=True, data_cdr_type='sample-marker')
            elif mk_sym == 'triangle':
                add_text(layer_g, mk_x, mk_y, '▲', size=3, fill='#c0392b',
                         anchor='middle', bold=True, data_cdr_type='sample-marker')
            else:
                add_circle(layer_g, mk_x, mk_y, 1, fill='#c0392b')
            if mk_label:
                add_text(layer_g, mk_x - 4, mk_y, mk_label, size=1.8,
                         fill='#c0392b', anchor='end', data_cdr_type='sample-label')

        # === Age (Ma) annotation on the right of lithology column ===
        if show_age and layer.get('age_ma'):
            age_text = f"{layer['age_ma']}"
            add_text(layer_g, colRight + 1.5, midY_, age_text,
                     size=1.8, fill='#e67e22', anchor='start',
                     data_cdr_type='age-ma')

        # Row separator (thin)
        add_line(body_grp, MARGIN_LEFT, topY, TABLE_RIGHT, topY,
                 stroke='#bbb', sw=0.12, data_cdr_type='row-separator')

    # ===== Major boundary lines =====
    for start, end, topY, bottomY, val in erathem_groups:
        add_line(body_grp, MARGIN_LEFT, topY, TABLE_RIGHT, topY,
                 stroke='#000', sw=0.35, data_cdr_type='erathem-boundary')
    for start, end, topY, bottomY, val in system_groups:
        add_line(body_grp, MARGIN_LEFT + COL_W[1], topY, TABLE_RIGHT, topY,
                 stroke='#555', sw=0.22, data_cdr_type='system-boundary')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-outlines
    # ═══════════════════════════════════════════
    outline_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-outlines',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Outlines',
        'data-cdr-layer': 'outlines'
    })

    # Bottom border
    add_line(outline_grp, MARGIN_LEFT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_BOTTOM,
             stroke='#000', sw=0.4, data_cdr_type='table-bottom')
    # Outer borders
    add_line(outline_grp, MARGIN_LEFT, TABLE_BOTTOM, MARGIN_LEFT, TABLE_TOP,
             stroke='#000', sw=0.4, data_cdr_type='table-left')
    add_line(outline_grp, TABLE_RIGHT, TABLE_BOTTOM, TABLE_RIGHT, TABLE_TOP,
             stroke='#000', sw=0.4, data_cdr_type='table-right')
    add_line(outline_grp, MARGIN_LEFT, TABLE_TOP, TABLE_RIGHT, TABLE_TOP,
             stroke='#000', sw=0.4, data_cdr_type='table-top')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-footer
    # ═══════════════════════════════════════════
    footer_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-footer',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Footer',
        'data-cdr-layer': 'footer'
    })

    ratio = int(total_thick * 1000 / draw_h) if draw_h > 0 else 1
    footer_text = (f"{title}  |  总厚 {total_thick:,.0f}m  |  {n_layers} 层  "
                   f"|  比例尺 1:{ratio:,}")
    add_text(footer_grp, MARGIN_LEFT, TABLE_BOTTOM - 5, footer_text,
             size=2, fill='#666', data_cdr_type='footer-text')

    # ═══════════════════════════════════════════
    # LAYER GROUP: cdr-legend
    # ═══════════════════════════════════════════
    legend_grp = ET.SubElement(svg, 'g', {
        'id': 'cdr-legend',
        'inkscape:groupmode': 'layer',
        'inkscape:label': 'Legend',
        'data-cdr-layer': 'legend'
    })

    # Collect unique patterns in stratigraphic order
    unique = {}
    for layer in layers:
        pn = layer.get('pattern', 'pure')
        if pn not in unique:
            unique[pn] = layer

    if len(unique) > 1:
        lgX = TABLE_RIGHT + 4
        lgY = TABLE_TOP - HEADER_H
        lgW = 38
        itemH = 7.0
        nUnique = len(unique)
        lgH = nUnique * itemH + 10

        add_rect(legend_grp, lgX, lgY - lgH, lgW, lgH,
                 fill='#fff', stroke='#666', sw=0.25,
                 data_cdr_type='legend-box')

        add_text(legend_grp, lgX + lgW/2, lgY - lgH + 5.5, '图  例',
                 size=2.5, bold=True, anchor='middle', fill='#000',
                 data_cdr_type='legend-title')

        add_line(legend_grp, lgX + 3, lgY - lgH + 8, lgX + lgW - 3, lgY - lgH + 8,
                 stroke='#ccc', sw=0.15, data_cdr_type='legend-separator')

        for j, (pn, ldata) in enumerate(unique.items()):
            iy = lgY - lgH + 9.5 + j * itemH
            swatchY = iy - 3

            clip_id = f'lgclip_{j}'
            clip = ET.SubElement(defs, 'clipPath', {'id': clip_id})
            ET.SubElement(clip, 'rect', {
                'x': str(lgX + 4), 'y': str(swatchY),
                'width': '6', 'height': '6'
            })

            c, m, y_, k = ldata['c'], ldata['m'], ldata['y'], ldata['k']
            fc = cmyk_fill(c, m, y_, k)
            add_rect(legend_grp, lgX + 4, swatchY, 6, 6,
                     fill=fc, stroke='#888', sw=0.12,
                     data_cdr_type='legend-swatch')

            if pn != 'pure' and pn in PATTERNS:
                pat_g = ET.SubElement(legend_grp, 'g', {'clip-path': f'url(#{clip_id})'})
                PATTERNS[pn](pat_g, lgX + 4, swatchY, 6, 6)

            add_text(legend_grp, lgX + 13, iy, ldata.get('formation', ''),
                     size=2, fill='#333', data_cdr_type='legend-label')

    # ═══════════════════════════════════════════
    # OUTPUT
    # ═══════════════════════════════════════════
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
