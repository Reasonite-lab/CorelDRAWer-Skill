#!/usr/bin/env python3
"""
CDR Template Generator — generates VBA that creates structured CorelDRAW documents.
Usage: python3 cdr_template_gen.py data.json [output.bas]
"""

import json, sys

def generate_cdr_template(data, output_path='cdr_template.bas'):
    """Generate VBA that creates a fully-structured CorelDRAW document with:
    - Named layers (Background, Title, Header, Body, Grid, Legend, Footer)
    - Object styles
    - Proper grouping
    """
    layers = data['layers']
    title = data.get('title', 'Stratigraphic Column')
    total_thick = sum(l['thick'] for l in layers)

    lines = []
    lines.append("' ═══════════════════════════════════════")
    lines.append(f"' CorelDRAWer CDR Template — {title}")
    lines.append(f"' Layers: {len(layers)}, Total: {total_thick}m")
    lines.append("' ═══════════════════════════════════════")
    lines.append("Option Explicit")
    lines.append("")

    # ── Layer management ──
    lines.append("' ═══ Layer Setup ═══")
    lines.append("Private Function GetOrCreateLayer(doc As Document, layerName As String) As Layer")
    lines.append("    On Error Resume Next")
    lines.append("    Set GetOrCreateLayer = doc.ActivePage.Layers(layerName)")
    lines.append("    If Err.Number <> 0 Then")
    lines.append("        Set GetOrCreateLayer = doc.ActivePage.CreateLayer(layerName)")
    lines.append("    End If")
    lines.append("    On Error GoTo 0")
    lines.append("End Function")
    lines.append("")

    # ── Style presets ──
    lines.append("' ═══ Style Constants (Nature journal specs) ═══")
    lines.append("Private Const FONT_FAMILY = \"Arial\"")
    lines.append("Private Const TITLE_SIZE = 12#")
    lines.append("Private Const HEADER_SIZE = 7#")
    lines.append("Private Const BODY_SIZE = 5.5#")
    lines.append("Private Const SMALL_SIZE = 4#")
    lines.append("Private Const LINE_THIN = 0.08")
    lines.append("Private Const LINE_MED = 0.15")
    lines.append("Private Const LINE_THICK = 0.25")
    lines.append("Private Const COLOR_TEXT = \"CMYK,0,0,0,100\"")
    lines.append("Private Const COLOR_MUTED = \"CMYK,0,0,0,50\"")
    lines.append("Private Const COLOR_GRID = \"CMYK,0,0,0,8\"")
    lines.append("")

    # ── Helper: styled text ──
    lines.append("' ═══ Styled Text Helper ═══")
    lines.append("Private Function AddStyledText(layer As Layer, x As Double, y As Double, _")
    lines.append("    txt As String, sz As Double, Optional bold As Boolean = False, _")
    lines.append("    Optional align As Long = cdrCenterAlignment) As Shape")
    lines.append("    Set AddStyledText = layer.CreateArtisticText(x, y, txt)")
    lines.append("    With AddStyledText.Text.Story")
    lines.append("        .Font = FONT_FAMILY: .Size = sz: .Bold = bold")
    lines.append("        .Alignment = align")
    lines.append("    End With")
    lines.append("    AddStyledText.Fill.UniformColor.CMYKAssign 0,0,0,100")
    lines.append("End Function")
    lines.append("")

    # ── Layout constants ──
    lines.append("' ═══ Layout ═══")
    lines.append("Private Const MARGIN = 8#")
    lines.append("Private Const TABLE_TOP = 290#")
    lines.append("Private Const TABLE_BOT = 22#")
    lines.append("Private Const HEADER_H = 14#")
    lines.append("")

    col_x = [0, 8, 21, 34, 47, 65, 77, 84, 124, 139, 146, 156]
    col_w = [0, 13, 13, 13, 18, 12, 7, 40, 15, 7, 10, 65]
    draw_h = 290 - 22 - 14
    scale_f = draw_h / total_thick if total_thick > 0 else 1
    TABLE_TOP = 290
    TABLE_BOT = 22
    HEADER_H = 14

    lines.append("' Column positions (11 columns)")
    for ci in range(1, 12):
        lines.append(f"Private Const CX{ci} = {col_x[ci]}#")
        lines.append(f"Private Const CW{ci} = {col_w[ci]}#")
    lines.append(f"Private Const SCALE_F = {scale_f}")
    lines.append(f"Private Const TOTAL_THICK = {total_thick}")
    lines.append("")

    # ── Main Draw Sub ──
    lines.append("' ═══════════════════════════════════════")
    lines.append("' MAIN: DrawCompleteColumn")
    lines.append("' ═══════════════════════════════════════")
    lines.append("Public Sub DrawCompleteColumn()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("")
    lines.append("    Dim doc As Document: Set doc = ActiveDocument")
    lines.append("    doc.BeginCommandGroup \"CorelDRAWer Stratigraphic Column\"")
    lines.append("    doc.Unit = cdrMillimeter")
    lines.append("    doc.ReferencePoint = cdrBottomLeft")
    lines.append("")

    # Create layers
    lines.append("    ' --- Create structured layers ---")
    for ln in ['Background', 'Title', 'Header', 'Body', 'Grid', 'Legend', 'Footer']:
        lines.append(f"    Dim lyr{ln} As Layer: Set lyr{ln} = GetOrCreateLayer(doc, \"CDR_{ln}\")")
    lines.append("")

    # Background
    lines.append("    ' --- Background ---")
    lines.append("    lyrBackground.Activate")
    lines.append(f"    lyrBackground.CreateRectangle 0, 0, {col_x[11]+col_w[11]+50}, {290+50}")
    lines.append("")

    # Title
    lines.append("    ' --- Title ---")
    lines.append("    lyrTitle.Activate")
    title_esc = title.replace('"', '""')
    lines.append(f"    AddStyledText lyrTitle, {(col_x[1]+col_x[11]+col_w[11])/2}, TABLE_TOP+8, \"{title_esc}\", TITLE_SIZE, True")

    loc = data.get('location', '')
    if loc:
        lines.append(f"    AddStyledText lyrTitle, {(col_x[1]+col_x[11]+col_w[11])/2}, TABLE_TOP+2, \"{loc}\", HEADER_SIZE")
    lines.append("")

    # Header
    lines.append("    ' --- Header ---")
    lines.append("    lyrHeader.Activate")
    lines.append(f"    Dim hdr As Shape: Set hdr = lyrHeader.CreateRectangle(CX1, TABLE_TOP-HEADER_H, CX11+CW11, TABLE_TOP)")
    lines.append("    hdr.Fill.UniformColor.CMYKAssign 0,0,0,3")
    lines.append("    hdr.Outline.SetNoOutline")

    header_labels = ["界", "系", "统", "组", "代号", "化石", "岩性柱", "粒度", "构造", "厚度(m)", "岩性描述"]
    for i, label in enumerate(header_labels):
        ci = i + 1
        lines.append(f"    AddStyledText lyrHeader, CX{ci}+CW{ci}/2, TABLE_TOP-HEADER_H/2, \"{label}\", HEADER_SIZE, True")
    lines.append("")

    # Grid
    lines.append("    ' --- Grid lines ---")
    lines.append("    lyrGrid.Activate")
    for ci in range(2, 12):
        lines.append(f"    With lyrGrid.CreateLine(CX{ci}, TABLE_TOP, CX{ci}, TABLE_BOT): .Outline.Width = LINE_THIN: End With")
    lines.append("")

    # Body — draw layers
    lines.append("    ' --- Body: Draw strata ---")
    lines.append("    lyrBody.Activate")
    lines.append("    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H")
    lines.append("    Dim i As Long, layH As Double, layB As Double, midY As Double")
    lines.append("")

    for i, lay in enumerate(layers):
        fm = lay.get('formation', '')
        th = lay['thick']
        c, m, y, k = lay['c'], lay['m'], lay['y'], lay['k']

        lines.append(f"    ' --- Layer {i+1}: {fm} ({th}m) ---")
        lines.append(f"    layH = {th} * SCALE_F: If layH < 3.5 Then layH = 3.5")
        lines.append("    layB = currentY - layH: midY = (currentY + layB) / 2")
        lines.append("")

        # Lithology rect
        lines.append("    ' Lithology rect")
        lines.append("    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)")
        lines.append(f"    r.Fill.UniformColor.CMYKAssign {c},{m},{y},{k}")
        lines.append("    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED")
        lines.append("")

        # Text columns
        for tcol, field in [(1,'erathem'),(2,'system'),(3,'series'),(4,'formation'),(5,'symbol')]:
            val = lay.get(field, '')
            if val:
                val_esc = val.replace('"', '""')
                lines.append(f"    AddStyledText lyrBody, CX{tcol}+1, midY, \"{val_esc}\", BODY_SIZE, False, cdrLeftAlignment")
        lines.append(f"    AddStyledText lyrBody, CX10+1, midY, \"{th:.1f}\", BODY_SIZE, False, cdrLeftAlignment")
        desc = lay.get('descr', '')
        if desc:
            desc_esc = desc.replace('"', '""')
            lines.append(f"    AddStyledText lyrBody, CX11+1, midY, \"{desc_esc}\", SMALL_SIZE, False, cdrLeftAlignment")
        lines.append("")

        # Grain triangle if present
        grain = lay.get('grain', 0)
        if grain > 0:
            lines.append("    ' Grain indicator")
            lines.append(f"    Dim gW As Double: gW = (CW8-3) * {grain} / 6: If gW < 0.5 Then gW = 0.5")
            lines.append("    Dim crv As Curve: Set crv = Application.CreateCurve()")
            lines.append("    crv.CreateSubPath CX8+1.5, layB")
            lines.append("    crv.AppendLineSegment CX8+1.5+gW, layB")
            lines.append("    crv.AppendLineSegment CX8+1.5+gW, currentY")
            lines.append("    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True")
            lines.append("    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)")
            lines.append("    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline")
            lines.append("")

        # Contact
        contact = lay.get('contact', '')
        if contact:
            contact_text = "~~~ Unconformity ~~~" if contact == 'unconformity' else "--- Disconformity ---"
            lines.append(f"    AddStyledText lyrBody, CX7+CW7/2, layB, \"{contact_text}\", SMALL_SIZE, False")
            lines.append("")

        lines.append("    currentY = layB")
        lines.append("")

    # Legend
    lines.append("    ' --- Legend ---")
    lines.append("    lyrLegend.Activate")
    unique = {}
    for lay in layers:
        pn = lay.get('pattern', 'pure')
        if pn not in unique:
            unique[pn] = lay
    leg_x = col_x[11] + col_w[11] + 5
    leg_y = TABLE_TOP - HEADER_H
    lines.append(f"    Dim lgY As Double: lgY = {leg_y}")
    lines.append(f"    With lyrLegend.CreateRectangle({leg_x}, lgY-{len(unique)*7+10}, 35, {len(unique)*7+10})")
    lines.append("        .Fill.UniformColor.CMYKAssign 0,0,0,0: .Outline.Color.CMYKAssign 0,0,0,30")
    lines.append("    End With")
    lines.append(f"    AddStyledText lyrLegend, {leg_x}+17.5, lgY-{len(unique)*7+10}+5, \"Legend\", 6, True")

    for j, (pn, ldata) in enumerate(unique.items()):
        fm = ldata.get('formation', pn)
        fm_esc = fm.replace('"', '""')
        sy = leg_y - (len(unique)*7+10) + 10 + j*7
        lines.append(f"    With lyrLegend.CreateRectangle({leg_x}+3, {sy}-3, 5, 5)")
        lines.append(f"        .Fill.UniformColor.CMYKAssign {ldata['c']},{ldata['m']},{ldata['y']},{ldata['k']}")
        lines.append("        .Outline.Color.CMYKAssign 0,0,0,20")
        lines.append("    End With")
        lines.append(f"    AddStyledText lyrLegend, {leg_x}+10, {sy}, \"{fm_esc}\", SMALL_SIZE, False, cdrLeftAlignment")

    lines.append("")

    # Footer
    lines.append("    ' --- Footer ---")
    lines.append("    lyrFooter.Activate")
    ratio = int(total_thick * 1000 / draw_h) if draw_h > 0 else 1
    lines.append(f"    AddStyledText lyrFooter, MARGIN, TABLE_BOT-5, \"{title_esc} | {total_thick:,.0f}m | {len(layers)} layers | Scale 1:{ratio:,} | CorelDRAWer-Skill\", SMALL_SIZE, False, cdrLeftAlignment")
    lines.append("")

    lines.append("    doc.EndCommandGroup")
    lines.append("    MsgBox \"CDR Template complete!\"")
    lines.append("    Exit Sub")
    lines.append("")
    lines.append("ErrHandler:")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    vba = '\n'.join(lines)
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(vba)
    return vba

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 cdr_template_gen.py data.json [output.bas]")
        sys.exit(1)

    with open(sys.argv[1], 'r') as f:
        data = json.load(f)
    out = sys.argv[2] if len(sys.argv) > 2 else 'cdr_template.bas'
    generate_cdr_template(data, out)
    print(f'✅ CDR template VBA saved: {out}')
