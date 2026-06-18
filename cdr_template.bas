' ═══════════════════════════════════════
' CorelDRAWer CDR Template — São Luís–Grajaú Basin — Aptian Stratigraphy
' Layers: 6, Total: 700m
' ═══════════════════════════════════════
Option Explicit

' ═══ Layer Setup ═══
Private Function GetOrCreateLayer(doc As Document, layerName As String) As Layer
    On Error Resume Next
    Set GetOrCreateLayer = doc.ActivePage.Layers(layerName)
    If Err.Number <> 0 Then
        Set GetOrCreateLayer = doc.ActivePage.CreateLayer(layerName)
    End If
    On Error GoTo 0
End Function

' ═══ Style Constants (Nature journal specs) ═══
Private Const FONT_FAMILY = "Arial"
Private Const TITLE_SIZE = 12#
Private Const HEADER_SIZE = 7#
Private Const BODY_SIZE = 5.5#
Private Const SMALL_SIZE = 4#
Private Const LINE_THIN = 0.08
Private Const LINE_MED = 0.15
Private Const LINE_THICK = 0.25
Private Const COLOR_TEXT = "CMYK,0,0,0,100"
Private Const COLOR_MUTED = "CMYK,0,0,0,50"
Private Const COLOR_GRID = "CMYK,0,0,0,8"

' ═══ Styled Text Helper ═══
Private Function AddStyledText(layer As Layer, x As Double, y As Double, _
    txt As String, sz As Double, Optional bold As Boolean = False, _
    Optional align As Long = cdrCenterAlignment) As Shape
    Set AddStyledText = layer.CreateArtisticText(x, y, txt)
    With AddStyledText.Text.Story
        .Font = FONT_FAMILY: .Size = sz: .Bold = bold
        .Alignment = align
    End With
    AddStyledText.Fill.UniformColor.CMYKAssign 0,0,0,100
End Function

' ═══ Layout ═══
Private Const MARGIN = 8#
Private Const TABLE_TOP = 290#
Private Const TABLE_BOT = 22#
Private Const HEADER_H = 14#

' Column positions (11 columns)
Private Const CX1 = 8#
Private Const CW1 = 13#
Private Const CX2 = 21#
Private Const CW2 = 13#
Private Const CX3 = 34#
Private Const CW3 = 13#
Private Const CX4 = 47#
Private Const CW4 = 18#
Private Const CX5 = 65#
Private Const CW5 = 12#
Private Const CX6 = 77#
Private Const CW6 = 7#
Private Const CX7 = 84#
Private Const CW7 = 40#
Private Const CX8 = 124#
Private Const CW8 = 15#
Private Const CX9 = 139#
Private Const CW9 = 7#
Private Const CX10 = 146#
Private Const CW10 = 10#
Private Const CX11 = 156#
Private Const CW11 = 65#
Private Const SCALE_F = 0.3628571428571429
Private Const TOTAL_THICK = 700

' ═══════════════════════════════════════
' MAIN: DrawCompleteColumn
' ═══════════════════════════════════════
Public Sub DrawCompleteColumn()
    On Error GoTo ErrHandler

    Dim doc As Document: Set doc = ActiveDocument
    doc.BeginCommandGroup "CorelDRAWer Stratigraphic Column"
    doc.Unit = cdrMillimeter
    doc.ReferencePoint = cdrBottomLeft

    ' --- Create structured layers ---
    Dim lyrBackground As Layer: Set lyrBackground = GetOrCreateLayer(doc, "CDR_Background")
    Dim lyrTitle As Layer: Set lyrTitle = GetOrCreateLayer(doc, "CDR_Title")
    Dim lyrHeader As Layer: Set lyrHeader = GetOrCreateLayer(doc, "CDR_Header")
    Dim lyrBody As Layer: Set lyrBody = GetOrCreateLayer(doc, "CDR_Body")
    Dim lyrGrid As Layer: Set lyrGrid = GetOrCreateLayer(doc, "CDR_Grid")
    Dim lyrLegend As Layer: Set lyrLegend = GetOrCreateLayer(doc, "CDR_Legend")
    Dim lyrFooter As Layer: Set lyrFooter = GetOrCreateLayer(doc, "CDR_Footer")

    ' --- Background ---
    lyrBackground.Activate
    lyrBackground.CreateRectangle 0, 0, 271, 340

    ' --- Title ---
    lyrTitle.Activate
    AddStyledText lyrTitle, 114.5, TABLE_TOP+8, "São Luís–Grajaú Basin — Aptian Stratigraphy", TITLE_SIZE, True
    AddStyledText lyrTitle, 114.5, TABLE_TOP+2, "Brazilian Equatorial Margin", HEADER_SIZE

    ' --- Header ---
    lyrHeader.Activate
    Dim hdr As Shape: Set hdr = lyrHeader.CreateRectangle(CX1, TABLE_TOP-HEADER_H, CX11+CW11, TABLE_TOP)
    hdr.Fill.UniformColor.CMYKAssign 0,0,0,3
    hdr.Outline.SetNoOutline
    AddStyledText lyrHeader, CX1+CW1/2, TABLE_TOP-HEADER_H/2, "界", HEADER_SIZE, True
    AddStyledText lyrHeader, CX2+CW2/2, TABLE_TOP-HEADER_H/2, "系", HEADER_SIZE, True
    AddStyledText lyrHeader, CX3+CW3/2, TABLE_TOP-HEADER_H/2, "统", HEADER_SIZE, True
    AddStyledText lyrHeader, CX4+CW4/2, TABLE_TOP-HEADER_H/2, "组", HEADER_SIZE, True
    AddStyledText lyrHeader, CX5+CW5/2, TABLE_TOP-HEADER_H/2, "代号", HEADER_SIZE, True
    AddStyledText lyrHeader, CX6+CW6/2, TABLE_TOP-HEADER_H/2, "化石", HEADER_SIZE, True
    AddStyledText lyrHeader, CX7+CW7/2, TABLE_TOP-HEADER_H/2, "岩性柱", HEADER_SIZE, True
    AddStyledText lyrHeader, CX8+CW8/2, TABLE_TOP-HEADER_H/2, "粒度", HEADER_SIZE, True
    AddStyledText lyrHeader, CX9+CW9/2, TABLE_TOP-HEADER_H/2, "构造", HEADER_SIZE, True
    AddStyledText lyrHeader, CX10+CW10/2, TABLE_TOP-HEADER_H/2, "厚度(m)", HEADER_SIZE, True
    AddStyledText lyrHeader, CX11+CW11/2, TABLE_TOP-HEADER_H/2, "岩性描述", HEADER_SIZE, True

    ' --- Grid lines ---
    lyrGrid.Activate
    With lyrGrid.CreateLine(CX2, TABLE_TOP, CX2, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX3, TABLE_TOP, CX3, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX4, TABLE_TOP, CX4, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX5, TABLE_TOP, CX5, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX6, TABLE_TOP, CX6, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX7, TABLE_TOP, CX7, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX8, TABLE_TOP, CX8, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX9, TABLE_TOP, CX9, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX10, TABLE_TOP, CX10, TABLE_BOT): .Outline.Width = LINE_THIN: End With
    With lyrGrid.CreateLine(CX11, TABLE_TOP, CX11, TABLE_BOT): .Outline.Width = LINE_THIN: End With

    ' --- Body: Draw strata ---
    lyrBody.Activate
    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H
    Dim i As Long, layH As Double, layB As Double, midY As Double

    ' --- Layer 1: Upper Carbonate (85m) ---
    layH = 85 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 0,0,0,25
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Mesozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Cretaceous", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "Upper Aptian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Upper Carbonate", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "K₁u", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "85.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "Grey limestone with dolomite interbeds", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 2 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    currentY = layB

    ' --- Layer 2: Marine Shale (120m) ---
    layH = 120 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 0,5,15,55
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Mesozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Cretaceous", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "Upper Aptian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Marine Shale", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "K₁m", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "120.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "Dark grey marine shale with TOC-rich intervals", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 1 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    currentY = layB

    ' --- Layer 3: Evaporite-Anhydrite (45m) ---
    layH = 45 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 0,0,5,5
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Mesozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Cretaceous", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "Mid Aptian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Evaporite-Anhydrite", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "K₁e", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "45.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "White-grey anhydrite with thin halite layers", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 2 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    AddStyledText lyrBody, CX7+CW7/2, layB, "--- Disconformity ---", SMALL_SIZE, False

    currentY = layB

    ' --- Layer 4: Pre-Salt Siliciclastic (160m) ---
    layH = 160 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 0,30,40,15
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Mesozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Cretaceous", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "Lower Aptian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Pre-Salt Siliciclastic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "K₁p", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "160.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "Red-brown fluvial-lacustrine sandstone with conglomerate lenses", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 5 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    currentY = layB

    ' --- Layer 5: Rift Basin Fill (210m) ---
    layH = 210 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 20,10,30,25
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Mesozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Cretaceous", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "Barremian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Rift Basin Fill", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "K₁r", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "210.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "Variegated conglomerate and coarse sandstone — syn-rift deposits", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 6 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    AddStyledText lyrBody, CX7+CW7/2, layB, "~~~ Unconformity ~~~", SMALL_SIZE, False

    currentY = layB

    ' --- Layer 6: Crystalline Basement (80m) ---
    layH = 80 * SCALE_F: If layH < 3.5 Then layH = 3.5
    layB = currentY - layH: midY = (currentY + layB) / 2

    ' Lithology rect
    Dim r As Shape: Set r = lyrBody.CreateRectangle(CX7+0.3, layB, CX7+CW7-2.5, currentY)
    r.Fill.UniformColor.CMYKAssign 0,20,10,35
    r.Outline.Color.CMYKAssign 0,0,0,35: r.Outline.Width = LINE_MED

    AddStyledText lyrBody, CX1+1, midY, "Precambrian", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX2+1, midY, "Proterozoic", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX3+1, midY, "—", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX4+1, midY, "Crystalline Basement", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX5+1, midY, "Pt", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX10+1, midY, "80.0", BODY_SIZE, False, cdrLeftAlignment
    AddStyledText lyrBody, CX11+1, midY, "Granite-gneiss basement complex", SMALL_SIZE, False, cdrLeftAlignment

    ' Grain indicator
    Dim gW As Double: gW = (CW8-3) * 6 / 6: If gW < 0.5 Then gW = 0.5
    Dim crv As Curve: Set crv = Application.CreateCurve()
    crv.CreateSubPath CX8+1.5, layB
    crv.AppendLineSegment CX8+1.5+gW, layB
    crv.AppendLineSegment CX8+1.5+gW, currentY
    crv.AppendLineSegment CX8+1.5, currentY: crv.Closed = True
    Dim tri As Shape: Set tri = lyrBody.CreateCurve(crv)
    tri.Fill.UniformColor.CMYKAssign 0,0,0,40: tri.Outline.SetNoOutline

    currentY = layB

    ' --- Legend ---
    lyrLegend.Activate
    Dim lgY As Double: lgY = 276
    With lyrLegend.CreateRectangle(226, lgY-52, 35, 52)
        .Fill.UniformColor.CMYKAssign 0,0,0,0: .Outline.Color.CMYKAssign 0,0,0,30
    End With
    AddStyledText lyrLegend, 226+17.5, lgY-52+5, "Legend", 6, True
    With lyrLegend.CreateRectangle(226+3, 234-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 0,0,0,25
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 234, "Upper Carbonate", SMALL_SIZE, False, cdrLeftAlignment
    With lyrLegend.CreateRectangle(226+3, 241-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 0,5,15,55
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 241, "Marine Shale", SMALL_SIZE, False, cdrLeftAlignment
    With lyrLegend.CreateRectangle(226+3, 248-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 0,0,5,5
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 248, "Evaporite-Anhydrite", SMALL_SIZE, False, cdrLeftAlignment
    With lyrLegend.CreateRectangle(226+3, 255-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 0,30,40,15
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 255, "Pre-Salt Siliciclastic", SMALL_SIZE, False, cdrLeftAlignment
    With lyrLegend.CreateRectangle(226+3, 262-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 20,10,30,25
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 262, "Rift Basin Fill", SMALL_SIZE, False, cdrLeftAlignment
    With lyrLegend.CreateRectangle(226+3, 269-3, 5, 5)
        .Fill.UniformColor.CMYKAssign 0,20,10,35
        .Outline.Color.CMYKAssign 0,0,0,20
    End With
    AddStyledText lyrLegend, 226+10, 269, "Crystalline Basement", SMALL_SIZE, False, cdrLeftAlignment

    ' --- Footer ---
    lyrFooter.Activate
    AddStyledText lyrFooter, MARGIN, TABLE_BOT-5, "São Luís–Grajaú Basin — Aptian Stratigraphy | 700m | 6 layers | Scale 1:2,755 | CorelDRAWer-Skill", SMALL_SIZE, False, cdrLeftAlignment

    doc.EndCommandGroup
    MsgBox "CDR Template complete!"
    Exit Sub

ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub