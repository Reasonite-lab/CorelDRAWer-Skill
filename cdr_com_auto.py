#!/usr/bin/env python3
"""
CorelDRAW COM Automation v2.0 — directly control CorelDRAW to draw stratigraphic columns
══════════════════════════════════════════════════════════
Requires: Windows + CorelDRAW X4+ + Python + pywin32
Install: pip install pywin32
Usage:
  python3 cdr_com_auto.py data.json        # draw in current document
  python3 cdr_com_auto.py --com data.json  # force COM mode
  python3 cdr_com_auto.py --vba data.json  # generate VBA code
  python3 cdr_com_auto.py --svg data.json  # generate SVG (delegates to generate_column.py)
  python3 cdr_com_auto.py                  # use built-in demo data
══════════════════════════════════════════════════════════
"""

import json
import sys
import os
import time

# ============================================================================
# TRY TO IMPORT WIN32COM
# ============================================================================
try:
    import win32com.client
    from win32com.client import Dispatch, constants
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ============================================================================
# VBA CODE GENERATOR v2.0
# ============================================================================

def generate_vba_code(data, output_path=None):
    """
    Generate a complete VBA macro that draws the stratigraphic column in CorelDRAW.
    v2.0: 11-column layout with fossil, structure, grain, contact, age support.
    """
    layers = data['layers']
    title = data.get('title', '综合地层柱状图')
    location = data.get('location', '')

    total_thick = sum(l['thick'] for l in layers)
    n_layers = len(layers)

    has_fossils = any(l.get('fossils') for l in layers)
    has_structures = any(l.get('structures') for l in layers)

    # 11-column layout matching SVG v2.0
    # 1界 2系 3统 4组 5代号 6化石 7柱状图 8粒度 9构造 10厚度 11描述
    col_x = [0, 8, 21, 34, 47, 65, 77, 84, 124, 139, 146, 156]
    col_w = [0, 13, 13, 13, 18, 12, 7, 40, 15, 7, 10, 65]
    n_cols = 11

    lines = []
    lines.append("' ============================================================")
    lines.append(f"' CorelDRAW Stratigraphic Column VBA v2.0  —  {title}")
    lines.append(f"' Total thickness: {total_thick}m  ·  {n_layers} layers")
    lines.append(f"' Fossils: {has_fossils}  ·  Structures: {has_structures}")
    lines.append("' ============================================================")
    lines.append("")
    lines.append("Option Explicit")
    lines.append("")

    # ── Helper Subs ──
    lines.append("' ════════════════════════════════")
    lines.append("' Helper Procedures")
    lines.append("' ════════════════════════════════")
    lines.append("")
    lines.append("Private Sub AddHeaderText(layer As Layer, cx As Double, cy As Double, txt As String)")
    lines.append("    Dim s As Shape")
    lines.append("    Set s = layer.CreateArtisticText(cx, cy, txt)")
    lines.append("    s.Text.Story.Font = \"SimHei\"")
    lines.append("    s.Text.Story.Size = 7")
    lines.append("    s.Text.Story.Bold = True")
    lines.append("    s.Text.Story.Alignment = cdrCenterAlignment")
    lines.append("    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100")
    lines.append("End Sub")
    lines.append("")
    lines.append("Private Sub AddCellText(layer As Layer, cx As Double, topY As Double, botY As Double, cw As Double, txt As String)")
    lines.append("    If txt = \"\" Then Exit Sub")
    lines.append("    Dim s As Shape")
    lines.append("    Dim midY As Double: midY = (topY + botY) / 2")
    lines.append("    Set s = layer.CreateArtisticText(cx + 1, midY, txt)")
    lines.append("    s.Text.Story.Font = \"SimHei\"")
    lines.append("    s.Text.Story.Size = 5.5")
    lines.append("    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100")
    lines.append("End Sub")
    lines.append("")
    lines.append("Private Sub AddSmallText(layer As Layer, cx As Double, cy As Double, txt As String, sz As Double)")
    lines.append("    If txt = \"\" Then Exit Sub")
    lines.append("    Dim s As Shape")
    lines.append("    Set s = layer.CreateArtisticText(cx, cy, txt)")
    lines.append("    s.Text.Story.Font = \"SimHei\"")
    lines.append("    s.Text.Story.Size = sz")
    lines.append("    s.Text.Story.Alignment = cdrCenterAlignment")
    lines.append("    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 60")
    lines.append("End Sub")
    lines.append("")
    lines.append("Private Sub DrawGrainTriangle(layer As Layer, gsx As Double, topY As Double, botY As Double, gsw As Double, gv As Long)")
    lines.append("    ' Draw right triangle for grain-size indicator (1-6)")
    lines.append("    Dim w As Double: w = gsw * gv / 6")
    lines.append("    If w < 0.5 Then w = 0.5")
    lines.append("    Dim crv As Curve")
    lines.append("    Set crv = Application.CreateCurve()")
    lines.append("    crv.CreateSubPath gsx, botY")
    lines.append("    crv.AppendLineSegment gsx + w, botY")
    lines.append("    crv.AppendLineSegment gsx + w, topY")
    lines.append("    crv.AppendLineSegment gsx, topY")
    lines.append("    crv.Closed = True")
    lines.append("    Dim s As Shape: Set s = layer.CreateCurve(crv)")
    lines.append("    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 40")
    lines.append("    s.Outline.SetNoOutline")
    lines.append("End Sub")
    lines.append("")
    lines.append("Private Sub DrawContactLine(layer As Layer, cx As Double, cy As Double, contactType As String)")
    lines.append("    Dim s As Shape")
    lines.append("    If contactType = \"unconformity\" Then")
    lines.append("        Set s = layer.CreateArtisticText(cx, cy, \"~~~ 不整合 ~~~\")")
    lines.append("        s.Text.Story.Font = \"SimHei\": s.Text.Story.Size = 4")
    lines.append("        s.Text.Story.Alignment = cdrCenterAlignment")
    lines.append("        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0")
    lines.append("    ElseIf contactType = \"disconformity\" Then")
    lines.append("        Set s = layer.CreateArtisticText(cx, cy, \"--- 假整合 ---\")")
    lines.append("        s.Text.Story.Font = \"SimHei\": s.Text.Story.Size = 4")
    lines.append("        s.Text.Story.Alignment = cdrCenterAlignment")
    lines.append("        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0")
    lines.append("    End If")
    lines.append("End Sub")
    lines.append("")

    # ── Main Sub ──
    lines.append("' ════════════════════════════════")
    lines.append("' Main Entry Point")
    lines.append("' ════════════════════════════════")
    lines.append("")
    lines.append("Public Sub DrawColumn()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    If ActiveDocument Is Nothing Then")
    lines.append("        MsgBox \"Please open a document first\"")
    lines.append("        Exit Sub")
    lines.append("    End If")
    lines.append("")
    lines.append("    ActiveDocument.BeginCommandGroup \"Draw Stratigraphic Column\"")
    lines.append("    ActiveDocument.Unit = cdrMillimeter")
    lines.append("    ActiveDocument.ReferencePoint = cdrBottomLeft")
    lines.append("")

    # Column constants
    lines.append("    ' === Column layout (11 cols) ===")
    for ci in range(1, n_cols + 1):
        lines.append(f"    Const CX{ci} = {col_x[ci]}")
        lines.append(f"    Const CW{ci} = {col_w[ci]}")

    lines.append("    Const TABLE_TOP = 290")
    lines.append("    Const TABLE_BOT = 22")
    lines.append("    Const HEADER_H = 14")
    lines.append("")

    draw_h = 290 - 22 - 14
    scale_f = draw_h / total_thick if total_thick > 0 else 1

    lines.append(f"    Dim scaleF As Double: scaleF = {scale_f}")
    lines.append(f"    Dim totalThick As Double: totalThick = {total_thick}")
    lines.append("")

    # Title
    lines.append("    ' === Title ===")
    lines.append("    Dim t As Shape")
    title_safe = title.replace('"', '""')
    lines.append(f"    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 8, \"{title_safe}\")")
    lines.append("    With t.Text.Story: .Font = \"SimHei\": .Size = 12: .Bold = True: .Alignment = cdrCenterAlignment: End With")
    if location:
        loc_safe = location.replace('"', '""')
        lines.append(f"    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 2, \"{loc_safe}\")")
        lines.append("    With t.Text.Story: .Font = \"SimHei\": .Size = 7: .Alignment = cdrCenterAlignment: End With")
    lines.append("")

    # Header
    lines.append("    ' === Header ===")
    lines.append("    Dim hdr As Shape")
    lines.append("    Set hdr = ActiveLayer.CreateRectangle(CX1, TABLE_TOP - HEADER_H, CX11 + CW11, TABLE_TOP)")
    lines.append("    hdr.Fill.UniformColor.CMYKAssign 0, 0, 0, 5")
    lines.append("    hdr.Outline.Color.CMYKAssign 0, 0, 0, 100")
    lines.append("    hdr.Outline.Width = 0.35")
    lines.append("")

    header_labels = [
        "AddHeaderText ActiveLayer, CX1 + CW1/2, TABLE_TOP - HEADER_H/2, \"界\"",
        "AddHeaderText ActiveLayer, CX2 + CW2/2, TABLE_TOP - HEADER_H/2, \"系\"",
        "AddHeaderText ActiveLayer, CX3 + CW3/2, TABLE_TOP - HEADER_H/2, \"统\"",
        "AddHeaderText ActiveLayer, CX4 + CW4/2, TABLE_TOP - HEADER_H/2, \"组\"",
        "AddHeaderText ActiveLayer, CX5 + CW5/2, TABLE_TOP - HEADER_H/2, \"代号\"",
        "AddHeaderText ActiveLayer, CX6 + CW6/2, TABLE_TOP - HEADER_H/2, \"化石\"",
        "AddHeaderText ActiveLayer, CX7 + CW7/2, TABLE_TOP - HEADER_H/2, \"岩性柱\"",
        "AddHeaderText ActiveLayer, CX8 + CW8/2, TABLE_TOP - HEADER_H/2, \"粒度\"",
        "AddHeaderText ActiveLayer, CX9 + CW9/2, TABLE_TOP - HEADER_H/2, \"构造\"",
        "AddHeaderText ActiveLayer, CX10 + CW10/2, TABLE_TOP - HEADER_H/2, \"厚度(m)\"",
        "AddHeaderText ActiveLayer, CX11 + CW11/2, TABLE_TOP - HEADER_H/2, \"岩性描述\"",
    ]
    for hl in header_labels:
        lines.append(f"    {hl}")

    # Vertical dividers
    lines.append("")
    lines.append("    ' Vertical dividers")
    for ci in range(2, n_cols + 1):
        lines.append(f"    ActiveLayer.CreateLine CX{ci}, TABLE_TOP, CX{ci}, TABLE_BOT")
    lines.append("")

    # Data array
    lines.append(f"    ' === Layer data ({n_layers} layers) ===")
    lines.append(f"    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H")
    lines.append(f"    Dim data({n_layers - 1}, 15) As Variant")
    lines.append("")

    for i, lay in enumerate(layers):
        fm = lay.get('formation', '')
        lines.append(f"    ' Layer {i+1}: {fm}")

        def esc(s):
            return s.replace('"', '""')

        lines.append(f"    data({i}, 0) = \"{esc(lay.get('erathem', ''))}\"")
        lines.append(f"    data({i}, 1) = \"{esc(lay.get('system', ''))}\"")
        lines.append(f"    data({i}, 2) = \"{esc(lay.get('series', ''))}\"")
        lines.append(f"    data({i}, 3) = \"{esc(lay.get('formation', ''))}\"")
        lines.append(f"    data({i}, 4) = \"{esc(lay.get('symbol', ''))}\"")

        fossils = lay.get('fossils', [])
        foss_str = ' '.join(fossils[:3]) if fossils else ''
        lines.append(f"    data({i}, 5) = \"{esc(foss_str)}\"")

        lines.append(f"    data({i}, 6) = {lay['thick']}")
        lines.append(f"    data({i}, 7) = \"{esc(lay.get('descr', ''))}\"")
        lines.append(f"    data({i}, 8) = \"{esc(lay.get('pattern', 'pure'))}\"")

        lines.append(f"    data({i}, 9) = {lay['c']}: data({i}, 10) = {lay['m']}")
        lines.append(f"    data({i}, 11) = {lay['y']}: data({i}, 12) = {lay['k']}")

        lines.append(f"    data({i}, 13) = {lay.get('grain', 0)}")

        structures = lay.get('structures', [])
        struct_str = ' '.join(structures[:2]) if structures else ''
        lines.append(f"    data({i}, 14) = \"{esc(struct_str)}\"")

        contact = lay.get('contact', '')
        age = lay.get('age_ma', '')
        lines.append(f"    data({i}, 15) = \"{esc(contact)}\"")
        lines.append("")

    # Main loop
    lines.append("    ' === Draw each layer ===")
    lines.append("    Dim i As Long")
    lines.append("    For i = 0 To UBound(data, 1)")
    lines.append("        Dim layH As Double: layH = data(i, 6) * scaleF")
    lines.append("        If layH < 3.5 Then layH = 3.5")
    lines.append("        Dim layB As Double: layB = currentY - layH")
    lines.append("        Dim midY As Double: midY = (currentY + layB) / 2")
    lines.append("")
    lines.append("        ' Lithology rect (col 7)")
    lines.append("        Dim rect As Shape")
    lines.append("        Set rect = ActiveLayer.CreateRectangle(CX7 + 0.3, layB, CX7 + CW7 - 2.5, currentY)")
    lines.append("        rect.Fill.UniformColor.CMYKAssign data(i, 9), data(i, 10), data(i, 11), data(i, 12)")
    lines.append("        rect.Outline.Color.CMYKAssign 0, 0, 0, 35")
    lines.append("        rect.Outline.Width = 0.15")
    lines.append("")
    lines.append("        ' Text columns")
    lines.append("        AddCellText ActiveLayer, CX1, currentY, layB, CW1, data(i, 0)")
    lines.append("        AddCellText ActiveLayer, CX2, currentY, layB, CW2, data(i, 1)")
    lines.append("        AddCellText ActiveLayer, CX3, currentY, layB, CW3, data(i, 2)")
    lines.append("        AddCellText ActiveLayer, CX4, currentY, layB, CW4, data(i, 3)")
    lines.append("        AddCellText ActiveLayer, CX5, currentY, layB, CW5, data(i, 4)")

    if has_fossils:
        lines.append("")
        lines.append("        ' Fossil column (col 6)")
        lines.append("        If data(i, 5) <> \"\" Then")
        lines.append("            AddSmallText ActiveLayer, CX6 + CW6/2, midY, data(i, 5), 4.5")
        lines.append("        End If")

    lines.append("")
    lines.append("        ' Thickness (col 10)")
    lines.append("        AddCellText ActiveLayer, CX10, currentY, layB, CW10, Format(data(i, 6), \"0.0\")")
    lines.append("")
    lines.append("        ' Grain triangle (col 8)")
    lines.append("        If data(i, 13) > 0 Then")
    lines.append("            DrawGrainTriangle ActiveLayer, CX8 + 1.5, currentY, layB, CW8 - 3, data(i, 13)")
    lines.append("        End If")

    if has_structures:
        lines.append("")
        lines.append("        ' Structure column (col 9)")
        lines.append("        If data(i, 14) <> \"\" Then")
        lines.append("            AddSmallText ActiveLayer, CX9 + CW9/2, midY, data(i, 14), 4")
        lines.append("        End If")

    lines.append("")
    lines.append("        ' Description (col 11)")
    lines.append("        AddCellText ActiveLayer, CX11, currentY, layB, CW11, data(i, 7)")
    lines.append("")
    lines.append("        ' Contact line")
    lines.append("        If data(i, 15) <> \"\" Then")
    lines.append("            DrawContactLine ActiveLayer, CX7 + CW7/2, layB, data(i, 15)")
    lines.append("        End If")
    lines.append("")
    lines.append("        currentY = layB")
    lines.append("    Next i")
    lines.append("")

    # Borders
    lines.append("    ' === Borders ===")
    lines.append("    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX1, TABLE_TOP): .Outline.Width = 0.4: End With")
    lines.append("    With ActiveLayer.CreateLine(CX11 + CW11, TABLE_BOT, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With")
    lines.append("    With ActiveLayer.CreateLine(CX1, TABLE_TOP, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With")
    lines.append("    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX11 + CW11, TABLE_BOT): .Outline.Width = 0.4: End With")
    lines.append("")

    # Footer
    lines.append("    ' === Footer ===")
    ratio = int(total_thick * 1000 / draw_h) if draw_h > 0 else 1
    lines.append("    Dim ft As Shape")
    lines.append(f"    Set ft = ActiveLayer.CreateArtisticText(CX1, TABLE_BOT - 5, \"{title_safe} | Total {total_thick:,.0f}m | {n_layers} layers | Scale 1:{ratio:,}\")")
    lines.append("    ft.Text.Story.Font = \"SimHei\": ft.Text.Story.Size = 5")
    lines.append("    ft.Fill.UniformColor.CMYKAssign 0, 0, 0, 50")
    lines.append("")

    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"Drawing completed!\"")
    lines.append("    Exit Sub")
    lines.append("")
    lines.append("ErrHandler:")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    vba_code = '\n'.join(lines)

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(vba_code)

    return vba_code


# ============================================================================
# CORELDRAW COM AUTOMATION (Windows only)
# ============================================================================

def com_draw_column(data):
    """
    Connect to CorelDRAW via COM and draw the stratigraphic column.
    Runs on Windows with CorelDRAW installed.
    """
    if not HAS_WIN32:
        print("❌ pywin32 not installed. On Windows, run: pip install pywin32")
        print("   Generating VBA code as fallback...")
        return generate_vba_code(data, 'column_fallback.bas')

    try:
        corel = Dispatch("CorelDRAW.Application")
        corel.Visible = True
        time.sleep(0.5)
    except Exception as e:
        print(f"❌ Cannot connect to CorelDRAW: {e}")
        print("   Make sure CorelDRAW is installed and running.")
        return generate_vba_code(data, 'column_fallback.bas')

    layers_data = data['layers']
    title = data.get('title', '综合地层柱状图')
    location = data.get('location', '')
    total_thick = sum(l['thick'] for l in layers_data)

    print(f"✅ Connected to CorelDRAW")
    print(f"   Drawing: {title}")
    print(f"   {len(layers_data)} layers, {total_thick}m total")

    try:
        doc = corel.ActiveDocument
        if doc is None:
            doc = corel.CreateDocument()
        doc.Unit = 6  # cdrMillimeter
        doc.ReferencePoint = 7  # cdrBottomLeft

        doc.BeginCommandGroup("Draw Stratigraphic Column")

        activeLayer = doc.ActivePage.ActiveLayer

        # 11-column layout
        col_x = [0, 8, 21, 34, 47, 65, 77, 84, 124, 139, 146, 156]
        col_w = [0, 13, 13, 13, 18, 12, 7, 40, 15, 7, 10, 65]
        TABLE_TOP, TABLE_BOT, HEADER_H = 290, 22, 14
        draw_h = TABLE_TOP - TABLE_BOT - HEADER_H
        scale_f = draw_h / total_thick

        cdrCenterAlignment = 2

        # Title
        title_shape = activeLayer.CreateArtisticText(
            (col_x[1] + col_x[11] + col_w[11]) / 2, TABLE_TOP + 8, title)
        title_shape.Text.Story.Font = "SimHei"
        title_shape.Text.Story.Size = 12
        title_shape.Text.Story.Bold = True
        title_shape.Text.Story.Alignment = cdrCenterAlignment

        if location:
            loc_shape = activeLayer.CreateArtisticText(
                (col_x[1] + col_x[11] + col_w[11]) / 2, TABLE_TOP + 2, location)
            loc_shape.Text.Story.Font = "SimHei"
            loc_shape.Text.Story.Size = 7
            loc_shape.Text.Story.Alignment = cdrCenterAlignment

        # Header background
        hdr = activeLayer.CreateRectangle(col_x[1], TABLE_TOP - HEADER_H,
                                          col_x[11] + col_w[11], TABLE_TOP)
        hdr.Fill.UniformColor.CMYKAssign(0, 0, 0, 5)
        hdr.Outline.Color.CMYKAssign(0, 0, 0, 100)
        hdr.Outline.Width = 0.35

        # Header labels
        header_labels = [
            (1, "界"), (2, "系"), (3, "统"), (4, "组"), (5, "代号"),
            (6, "化石"), (7, "岩性柱"), (8, "粒度"), (9, "构造"),
            (10, "厚度\n(m)"), (11, "岩性描述")
        ]
        for hcol, htxt in header_labels:
            s = activeLayer.CreateArtisticText(
                col_x[hcol] + col_w[hcol] // 2, TABLE_TOP - HEADER_H // 2, htxt)
            s.Text.Story.Font = "SimHei"
            s.Text.Story.Size = 7
            s.Text.Story.Bold = True
            s.Text.Story.Alignment = cdrCenterAlignment
            s.Fill.UniformColor.CMYKAssign(0, 0, 0, 100)

        # Draw each layer
        currentY = TABLE_TOP - HEADER_H
        for i, lay in enumerate(layers_data):
            layH = max(lay['thick'] * scale_f, 3.5)
            layB = currentY - layH
            midY = (currentY + layB) / 2
            c, m, y, k = lay['c'], lay['m'], lay['y'], lay['k']

            # Lithology rect (col 7)
            rect = activeLayer.CreateRectangle(
                col_x[7] + 0.3, layB, col_x[7] + col_w[7] - 2.5, currentY)
            rect.Fill.UniformColor.CMYKAssign(c, m, y, k)
            rect.Outline.Color.CMYKAssign(0, 0, 0, 35)
            rect.Outline.Width = 0.15

            # Text columns
            for tcol, field in [(1, 'erathem'), (2, 'system'), (3, 'series'),
                                (4, 'formation'), (5, 'symbol'), (10, 'thick'),
                                (11, 'descr')]:
                txt = lay.get(field, '')
                if field == 'thick':
                    txt = f"{lay['thick']:.1f}"
                if txt:
                    s = activeLayer.CreateArtisticText(col_x[tcol] + 1, midY, str(txt))
                    s.Text.Story.Font = "SimHei"
                    s.Text.Story.Size = 5.5
                    s.Fill.UniformColor.CMYKAssign(0, 0, 0, 100)

            # Fossils (col 6)
            fossils = lay.get('fossils', [])
            if fossils:
                foss_str = ' '.join(fossils[:3])
                s = activeLayer.CreateArtisticText(col_x[6] + col_w[6] / 2, midY, foss_str)
                s.Text.Story.Font = "SimHei"
                s.Text.Story.Size = 4.5
                s.Text.Story.Alignment = cdrCenterAlignment
                s.Fill.UniformColor.CMYKAssign(0, 0, 0, 60)

            # Grain triangle (col 8)
            grain = lay.get('grain', 0)
            if grain > 0:
                gsx = col_x[8] + 1.5
                gsw = col_w[8] - 3
                gw = gsw * grain / 6
                if gw < 0.5:
                    gw = 0.5
                # Draw a filled polygon using CreateCurve
                crv = corel.CreateCurve()
                crv.CreateSubPath(gsx, layB)
                crv.AppendLineSegment(gsx + gw, layB)
                crv.AppendLineSegment(gsx + gw, currentY)
                crv.AppendLineSegment(gsx, currentY)
                crv.Closed = True
                tri = activeLayer.CreateCurve(crv)
                tri.Fill.UniformColor.CMYKAssign(0, 0, 0, 40)
                tri.Outline.SetNoOutline()

            # Structures (col 9)
            structures = lay.get('structures', [])
            if structures:
                struct_str = ' '.join(structures[:2])
                s = activeLayer.CreateArtisticText(col_x[9] + col_w[9] / 2, midY, struct_str)
                s.Text.Story.Font = "SimHei"
                s.Text.Story.Size = 4
                s.Text.Story.Alignment = cdrCenterAlignment
                s.Fill.UniformColor.CMYKAssign(0, 0, 0, 60)

            # Contact
            contact = lay.get('contact', '')
            if contact:
                cx = col_x[7] + col_w[7] / 2
                if contact == 'unconformity':
                    cs = activeLayer.CreateArtisticText(cx, layB, "~~~ 不整合 ~~~")
                    cs.Text.Story.Font = "SimHei"
                    cs.Text.Story.Size = 4
                    cs.Text.Story.Alignment = cdrCenterAlignment
                    cs.Fill.UniformColor.CMYKAssign(0, 100, 80, 0)
                elif contact == 'disconformity':
                    cs = activeLayer.CreateArtisticText(cx, layB, "--- 假整合 ---")
                    cs.Text.Story.Font = "SimHei"
                    cs.Text.Story.Size = 4
                    cs.Text.Story.Alignment = cdrCenterAlignment
                    cs.Fill.UniformColor.CMYKAssign(0, 100, 80, 0)

            currentY = layB

        doc.EndCommandGroup()
        print(f"✅ Drawing completed in CorelDRAW!")
        return True

    except Exception as e:
        try:
            doc.EndCommandGroup()
        except:
            pass
        print(f"❌ Error during drawing: {e}")
        raise


# ============================================================================
# MAIN
# ============================================================================

def main():
    data = {
        "title": "秭归地区综合地层柱状图",
        "location": "湖北秭归",
        "layers": [
            {"erathem":"新元古界","system":"南华系","series":"下统","formation":"莲沱组","symbol":"Nh₁l","thick":120,"descr":"紫红色中厚层砂岩","c":0,"m":40,"y":30,"k":10,"pattern":"sand","grain":4,"contact":"unconformity","age_ma":780},
            {"erathem":"新元古界","system":"南华系","series":"下统","formation":"南沱组","symbol":"Nh₁n","thick":45,"descr":"灰绿色冰碛砾岩","c":15,"m":0,"y":20,"k":20,"pattern":"conglo","grain":6,"age_ma":720},
            {"erathem":"新元古界","system":"震旦系","series":"下统","formation":"陡山沱组","symbol":"Z₁d","thick":180,"descr":"灰黑色泥质白云岩","c":5,"m":0,"y":10,"k":30,"pattern":"dolo","grain":2,"fossils":["algae"],"age_ma":635},
            {"erathem":"新元古界","system":"震旦系","series":"上统","formation":"灯影组","symbol":"Z₂dy","thick":350,"descr":"灰白色厚层白云岩","c":0,"m":0,"y":5,"k":10,"pattern":"dolo","grain":2,"fossils":["stromatolite"],"age_ma":551},
            {"erathem":"古生界","system":"寒武系","series":"下统","formation":"水井沱组","symbol":"∈₁s","thick":95,"descr":"黑色炭质页岩","c":0,"m":5,"y":10,"k":65,"pattern":"carbShale","grain":1,"fossils":["trilobite","brachiopod"],"age_ma":541},
        ]
    }

    args = sys.argv[1:]
    mode = 'auto'  # auto, com, vba, svg

    for a in args:
        if a == '--com': mode = 'com'
        elif a == '--vba': mode = 'vba'
        elif a == '--svg': mode = 'svg'
        elif a == '--new': pass
        elif a.endswith('.json'):
            with open(a, 'r', encoding='utf-8') as f:
                data = json.load(f)

    print(f"📊 Stratigraphic Column Generator v2.0")
    print(f"   Data: {len(data['layers'])} layers, total {sum(l['thick'] for l in data['layers']):.0f}m")
    print()

    if mode == 'svg':
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from generate_column import generate_svg
        output = data.get('title', 'column').replace(' ', '_') + '.svg'
        generate_svg(data, output)
        print(f"✅ SVG saved: {output}")
        return

    if mode == 'com' or (mode == 'auto' and HAS_WIN32):
        result = com_draw_column(data)
        if result:
            return
        else:
            print("⚠️  COM drawing failed. Falling back to VBA...")
            print()

    if mode == 'vba' or mode == 'auto':
        vba = generate_vba_code(data, 'column_macro.bas')
        print(f"✅ VBA macro saved: column_macro.bas")
        print(f"   Open CorelDRAW → Alt+F11 → Import → Run DrawColumn")
        return


if __name__ == '__main__':
    main()
