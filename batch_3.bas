' ============================================================
' CorelDRAWer Batch VBA — 3 sections
' ============================================================
Option Explicit


' ════════════════════════════════
' Helper Procedures
' ════════════════════════════════

Private Sub AddHeaderText(layer As Layer, cx As Double, cy As Double, txt As String)
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 7
    s.Text.Story.Bold = True
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddCellText(layer As Layer, cx As Double, topY As Double, botY As Double, cw As Double, txt As String)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Dim midY As Double: midY = (topY + botY) / 2
    Set s = layer.CreateArtisticText(cx + 1, midY, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 5.5
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddSmallText(layer As Layer, cx As Double, cy As Double, txt As String, sz As Double)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = sz
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 60
End Sub

Private Sub DrawGrainTriangle(layer As Layer, gsx As Double, topY As Double, botY As Double, gsw As Double, gv As Long)
    ' Draw right triangle for grain-size indicator (1-6)
    Dim w As Double: w = gsw * gv / 6
    If w < 0.5 Then w = 0.5
    Dim crv As Curve
    Set crv = Application.CreateCurve()
    crv.CreateSubPath gsx, botY
    crv.AppendLineSegment gsx + w, botY
    crv.AppendLineSegment gsx + w, topY
    crv.AppendLineSegment gsx, topY
    crv.Closed = True
    Dim s As Shape: Set s = layer.CreateCurve(crv)
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 40
    s.Outline.SetNoOutline
End Sub

Private Sub DrawContactLine(layer As Layer, cx As Double, cy As Double, contactType As String)
    Dim s As Shape
    If contactType = "unconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "~~~ 不整合 ~~~")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    ElseIf contactType = "disconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "--- 假整合 ---")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    End If
End Sub

' ════════════════════════════════
' Main Entry Point
' ════════════════════════════════

Public Sub DrawSection1()
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Draw Stratigraphic Column"
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft

    ' === Column layout (11 cols) ===
    Const CX1 = 8
    Const CW1 = 13
    Const CX2 = 21
    Const CW2 = 13
    Const CX3 = 34
    Const CW3 = 13
    Const CX4 = 47
    Const CW4 = 18
    Const CX5 = 65
    Const CW5 = 12
    Const CX6 = 77
    Const CW6 = 7
    Const CX7 = 84
    Const CW7 = 40
    Const CX8 = 124
    Const CW8 = 15
    Const CX9 = 139
    Const CW9 = 7
    Const CX10 = 146
    Const CW10 = 10
    Const CX11 = 156
    Const CW11 = 65
    Const TABLE_TOP = 290
    Const TABLE_BOT = 22
    Const HEADER_H = 14

    Dim scaleF As Double: scaleF = 2.9707602339181287
    Dim totalThick As Double: totalThick = 85.5

    ' === Title ===
    Dim t As Shape
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 8, "Meishan Section — Permian-Triassic Boundary GSSP")
    With t.Text.Story: .Font = "SimHei": .Size = 12: .Bold = True: .Alignment = cdrCenterAlignment: End With
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 2, "Meishan, Zhejiang, South China")
    With t.Text.Story: .Font = "SimHei": .Size = 7: .Alignment = cdrCenterAlignment: End With

    ' === Header ===
    Dim hdr As Shape
    Set hdr = ActiveLayer.CreateRectangle(CX1, TABLE_TOP - HEADER_H, CX11 + CW11, TABLE_TOP)
    hdr.Fill.UniformColor.CMYKAssign 0, 0, 0, 5
    hdr.Outline.Color.CMYKAssign 0, 0, 0, 100
    hdr.Outline.Width = 0.35

    AddHeaderText ActiveLayer, CX1 + CW1/2, TABLE_TOP - HEADER_H/2, "界"
    AddHeaderText ActiveLayer, CX2 + CW2/2, TABLE_TOP - HEADER_H/2, "系"
    AddHeaderText ActiveLayer, CX3 + CW3/2, TABLE_TOP - HEADER_H/2, "统"
    AddHeaderText ActiveLayer, CX4 + CW4/2, TABLE_TOP - HEADER_H/2, "组"
    AddHeaderText ActiveLayer, CX5 + CW5/2, TABLE_TOP - HEADER_H/2, "代号"
    AddHeaderText ActiveLayer, CX6 + CW6/2, TABLE_TOP - HEADER_H/2, "化石"
    AddHeaderText ActiveLayer, CX7 + CW7/2, TABLE_TOP - HEADER_H/2, "岩性柱"
    AddHeaderText ActiveLayer, CX8 + CW8/2, TABLE_TOP - HEADER_H/2, "粒度"
    AddHeaderText ActiveLayer, CX9 + CW9/2, TABLE_TOP - HEADER_H/2, "构造"
    AddHeaderText ActiveLayer, CX10 + CW10/2, TABLE_TOP - HEADER_H/2, "厚度(m)"
    AddHeaderText ActiveLayer, CX11 + CW11/2, TABLE_TOP - HEADER_H/2, "岩性描述"

    ' Vertical dividers
    ActiveLayer.CreateLine CX2, TABLE_TOP, CX2, TABLE_BOT
    ActiveLayer.CreateLine CX3, TABLE_TOP, CX3, TABLE_BOT
    ActiveLayer.CreateLine CX4, TABLE_TOP, CX4, TABLE_BOT
    ActiveLayer.CreateLine CX5, TABLE_TOP, CX5, TABLE_BOT
    ActiveLayer.CreateLine CX6, TABLE_TOP, CX6, TABLE_BOT
    ActiveLayer.CreateLine CX7, TABLE_TOP, CX7, TABLE_BOT
    ActiveLayer.CreateLine CX8, TABLE_TOP, CX8, TABLE_BOT
    ActiveLayer.CreateLine CX9, TABLE_TOP, CX9, TABLE_BOT
    ActiveLayer.CreateLine CX10, TABLE_TOP, CX10, TABLE_BOT
    ActiveLayer.CreateLine CX11, TABLE_TOP, CX11, TABLE_BOT

    ' === Layer data (7 layers) ===
    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H
    Dim data(6, 15) As Variant

    ' Layer 1: Yinkeng Fm
    data(0, 0) = "Mesozoic"
    data(0, 1) = "Triassic"
    data(0, 2) = "Lower"
    data(0, 3) = "Yinkeng Fm"
    data(0, 4) = "T₁y"
    data(0, 5) = "ammonite bivalve"
    data(0, 6) = 3.5
    data(0, 7) = "Pale yellow calcareous mudstone with ammonoids (H. parvus FAD)"
    data(0, 8) = "mud"
    data(0, 9) = 0: data(0, 10) = 0
    data(0, 11) = 10: data(0, 12) = 10
    data(0, 13) = 1
    data(0, 14) = ""
    data(0, 15) = "conformity"

    ' Layer 2: Yinkeng Fm
    data(1, 0) = "Mesozoic"
    data(1, 1) = "Triassic"
    data(1, 2) = "Lower"
    data(1, 3) = "Yinkeng Fm"
    data(1, 4) = "T₁y"
    data(1, 5) = "ammonite conodont bivalve"
    data(1, 6) = 7
    data(1, 7) = "Grey medium-bedded argillaceous limestone"
    data(1, 8) = "lime"
    data(1, 9) = 0: data(1, 10) = 0
    data(1, 11) = 0: data(1, 12) = 25
    data(1, 13) = 2
    data(1, 14) = ""
    data(1, 15) = ""

    ' Layer 3: Changxing Fm
    data(2, 0) = "Paleozoic"
    data(2, 1) = "Permian"
    data(2, 2) = "Upper"
    data(2, 3) = "Changxing Fm"
    data(2, 4) = "P₃c"
    data(2, 5) = "conodont brachiopod foraminifera"
    data(2, 6) = 12
    data(2, 7) = "Dark grey medium-bedded bioclastic limestone"
    data(2, 8) = "lime"
    data(2, 9) = 0: data(2, 10) = 0
    data(2, 11) = 0: data(2, 12) = 35
    data(2, 13) = 2
    data(2, 14) = ""
    data(2, 15) = "conformity"

    ' Layer 4: Changxing Fm
    data(3, 0) = "Paleozoic"
    data(3, 1) = "Permian"
    data(3, 2) = "Upper"
    data(3, 3) = "Changxing Fm"
    data(3, 4) = "P₃c"
    data(3, 5) = "conodont brachiopod"
    data(3, 6) = 8
    data(3, 7) = "Dark grey thin-bedded cherty limestone"
    data(3, 8) = "chert"
    data(3, 9) = 0: data(3, 10) = 0
    data(3, 11) = 0: data(3, 12) = 40
    data(3, 13) = 1
    data(3, 14) = ""
    data(3, 15) = ""

    ' Layer 5: Changxing Fm
    data(4, 0) = "Paleozoic"
    data(4, 1) = "Permian"
    data(4, 2) = "Upper"
    data(4, 3) = "Changxing Fm"
    data(4, 4) = "P₃c"
    data(4, 5) = "brachiopod conodont"
    data(4, 6) = 15
    data(4, 7) = "Black calcareous shale with limestone interbeds"
    data(4, 8) = "shale"
    data(4, 9) = 0: data(4, 10) = 5
    data(4, 11) = 15: data(4, 12) = 55
    data(4, 13) = 1
    data(4, 14) = ""
    data(4, 15) = ""

    ' Layer 6: Longtan Fm
    data(5, 0) = "Paleozoic"
    data(5, 1) = "Permian"
    data(5, 2) = "Upper"
    data(5, 3) = "Longtan Fm"
    data(5, 4) = "P₃l"
    data(5, 5) = ""
    data(5, 6) = 22
    data(5, 7) = "Grey sandstone interbedded with mudstone and coal seams"
    data(5, 8) = "sand"
    data(5, 9) = 0: data(5, 10) = 15
    data(5, 11) = 30: data(5, 12) = 20
    data(5, 13) = 4
    data(5, 14) = "cross_bed"
    data(5, 15) = "disconformity"

    ' Layer 7: Gufeng Fm
    data(6, 0) = "Paleozoic"
    data(6, 1) = "Permian"
    data(6, 2) = "Middle"
    data(6, 3) = "Gufeng Fm"
    data(6, 4) = "P₂g"
    data(6, 5) = "ammonite conodont"
    data(6, 6) = 18
    data(6, 7) = "Black siliceous shale with phosphate nodules"
    data(6, 8) = "carbShale"
    data(6, 9) = 0: data(6, 10) = 0
    data(6, 11) = 10: data(6, 12) = 60
    data(6, 13) = 1
    data(6, 14) = "concretion"
    data(6, 15) = ""

    ' === Draw each layer ===
    Dim i As Long
    For i = 0 To UBound(data, 1)
        Dim layH As Double: layH = data(i, 6) * scaleF
        If layH < 3.5 Then layH = 3.5
        Dim layB As Double: layB = currentY - layH
        Dim midY As Double: midY = (currentY + layB) / 2

        ' Lithology rect (col 7)
        Dim rect As Shape
        Set rect = ActiveLayer.CreateRectangle(CX7 + 0.3, layB, CX7 + CW7 - 2.5, currentY)
        rect.Fill.UniformColor.CMYKAssign data(i, 9), data(i, 10), data(i, 11), data(i, 12)
        rect.Outline.Color.CMYKAssign 0, 0, 0, 35
        rect.Outline.Width = 0.15

        ' Text columns
        AddCellText ActiveLayer, CX1, currentY, layB, CW1, data(i, 0)
        AddCellText ActiveLayer, CX2, currentY, layB, CW2, data(i, 1)
        AddCellText ActiveLayer, CX3, currentY, layB, CW3, data(i, 2)
        AddCellText ActiveLayer, CX4, currentY, layB, CW4, data(i, 3)
        AddCellText ActiveLayer, CX5, currentY, layB, CW5, data(i, 4)

        ' Fossil column (col 6)
        If data(i, 5) <> "" Then
            AddSmallText ActiveLayer, CX6 + CW6/2, midY, data(i, 5), 4.5
        End If

        ' Thickness (col 10)
        AddCellText ActiveLayer, CX10, currentY, layB, CW10, Format(data(i, 6), "0.0")

        ' Grain triangle (col 8)
        If data(i, 13) > 0 Then
            DrawGrainTriangle ActiveLayer, CX8 + 1.5, currentY, layB, CW8 - 3, data(i, 13)
        End If

        ' Structure column (col 9)
        If data(i, 14) <> "" Then
            AddSmallText ActiveLayer, CX9 + CW9/2, midY, data(i, 14), 4
        End If

        ' Description (col 11)
        AddCellText ActiveLayer, CX11, currentY, layB, CW11, data(i, 7)

        ' Contact line
        If data(i, 15) <> "" Then
            DrawContactLine ActiveLayer, CX7 + CW7/2, layB, data(i, 15)
        End If

        currentY = layB
    Next i

    ' === Borders ===
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX1, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX11 + CW11, TABLE_BOT, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_TOP, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX11 + CW11, TABLE_BOT): .Outline.Width = 0.4: End With

    ' === Footer ===
    Dim ft As Shape
    Set ft = ActiveLayer.CreateArtisticText(CX1, TABLE_BOT - 5, "Meishan Section — Permian-Triassic Boundary GSSP | Total 86m | 7 layers | Scale 1:336")
    ft.Text.Story.Font = "SimHei": ft.Text.Story.Size = 5
    ft.Fill.UniformColor.CMYKAssign 0, 0, 0, 50

    ActiveDocument.EndCommandGroup
    MsgBox "Drawing completed!"
    Exit Sub

ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub


' ════════════════════════════════
' Helper Procedures
' ════════════════════════════════

Private Sub AddHeaderText(layer As Layer, cx As Double, cy As Double, txt As String)
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 7
    s.Text.Story.Bold = True
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddCellText(layer As Layer, cx As Double, topY As Double, botY As Double, cw As Double, txt As String)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Dim midY As Double: midY = (topY + botY) / 2
    Set s = layer.CreateArtisticText(cx + 1, midY, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 5.5
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddSmallText(layer As Layer, cx As Double, cy As Double, txt As String, sz As Double)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = sz
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 60
End Sub

Private Sub DrawGrainTriangle(layer As Layer, gsx As Double, topY As Double, botY As Double, gsw As Double, gv As Long)
    ' Draw right triangle for grain-size indicator (1-6)
    Dim w As Double: w = gsw * gv / 6
    If w < 0.5 Then w = 0.5
    Dim crv As Curve
    Set crv = Application.CreateCurve()
    crv.CreateSubPath gsx, botY
    crv.AppendLineSegment gsx + w, botY
    crv.AppendLineSegment gsx + w, topY
    crv.AppendLineSegment gsx, topY
    crv.Closed = True
    Dim s As Shape: Set s = layer.CreateCurve(crv)
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 40
    s.Outline.SetNoOutline
End Sub

Private Sub DrawContactLine(layer As Layer, cx As Double, cy As Double, contactType As String)
    Dim s As Shape
    If contactType = "unconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "~~~ 不整合 ~~~")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    ElseIf contactType = "disconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "--- 假整合 ---")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    End If
End Sub

' ════════════════════════════════
' Main Entry Point
' ════════════════════════════════

Public Sub DrawSection2()
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Draw Stratigraphic Column"
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft

    ' === Column layout (11 cols) ===
    Const CX1 = 8
    Const CW1 = 13
    Const CX2 = 21
    Const CW2 = 13
    Const CX3 = 34
    Const CW3 = 13
    Const CX4 = 47
    Const CW4 = 18
    Const CX5 = 65
    Const CW5 = 12
    Const CX6 = 77
    Const CW6 = 7
    Const CX7 = 84
    Const CW7 = 40
    Const CX8 = 124
    Const CW8 = 15
    Const CX9 = 139
    Const CW9 = 7
    Const CX10 = 146
    Const CW10 = 10
    Const CX11 = 156
    Const CW11 = 65
    Const TABLE_TOP = 290
    Const TABLE_BOT = 22
    Const HEADER_H = 14

    Dim scaleF As Double: scaleF = 0.18814814814814815
    Dim totalThick As Double: totalThick = 1350

    ' === Title ===
    Dim t As Shape
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 8, "Integrated Ediacaran Stratigraphy — Eastern Gondwana Margin")
    With t.Text.Story: .Font = "SimHei": .Size = 12: .Bold = True: .Alignment = cdrCenterAlignment: End With
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 2, "Adelaide Rift Complex, South Australia")
    With t.Text.Story: .Font = "SimHei": .Size = 7: .Alignment = cdrCenterAlignment: End With

    ' === Header ===
    Dim hdr As Shape
    Set hdr = ActiveLayer.CreateRectangle(CX1, TABLE_TOP - HEADER_H, CX11 + CW11, TABLE_TOP)
    hdr.Fill.UniformColor.CMYKAssign 0, 0, 0, 5
    hdr.Outline.Color.CMYKAssign 0, 0, 0, 100
    hdr.Outline.Width = 0.35

    AddHeaderText ActiveLayer, CX1 + CW1/2, TABLE_TOP - HEADER_H/2, "界"
    AddHeaderText ActiveLayer, CX2 + CW2/2, TABLE_TOP - HEADER_H/2, "系"
    AddHeaderText ActiveLayer, CX3 + CW3/2, TABLE_TOP - HEADER_H/2, "统"
    AddHeaderText ActiveLayer, CX4 + CW4/2, TABLE_TOP - HEADER_H/2, "组"
    AddHeaderText ActiveLayer, CX5 + CW5/2, TABLE_TOP - HEADER_H/2, "代号"
    AddHeaderText ActiveLayer, CX6 + CW6/2, TABLE_TOP - HEADER_H/2, "化石"
    AddHeaderText ActiveLayer, CX7 + CW7/2, TABLE_TOP - HEADER_H/2, "岩性柱"
    AddHeaderText ActiveLayer, CX8 + CW8/2, TABLE_TOP - HEADER_H/2, "粒度"
    AddHeaderText ActiveLayer, CX9 + CW9/2, TABLE_TOP - HEADER_H/2, "构造"
    AddHeaderText ActiveLayer, CX10 + CW10/2, TABLE_TOP - HEADER_H/2, "厚度(m)"
    AddHeaderText ActiveLayer, CX11 + CW11/2, TABLE_TOP - HEADER_H/2, "岩性描述"

    ' Vertical dividers
    ActiveLayer.CreateLine CX2, TABLE_TOP, CX2, TABLE_BOT
    ActiveLayer.CreateLine CX3, TABLE_TOP, CX3, TABLE_BOT
    ActiveLayer.CreateLine CX4, TABLE_TOP, CX4, TABLE_BOT
    ActiveLayer.CreateLine CX5, TABLE_TOP, CX5, TABLE_BOT
    ActiveLayer.CreateLine CX6, TABLE_TOP, CX6, TABLE_BOT
    ActiveLayer.CreateLine CX7, TABLE_TOP, CX7, TABLE_BOT
    ActiveLayer.CreateLine CX8, TABLE_TOP, CX8, TABLE_BOT
    ActiveLayer.CreateLine CX9, TABLE_TOP, CX9, TABLE_BOT
    ActiveLayer.CreateLine CX10, TABLE_TOP, CX10, TABLE_BOT
    ActiveLayer.CreateLine CX11, TABLE_TOP, CX11, TABLE_BOT

    ' === Layer data (8 layers) ===
    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H
    Dim data(7, 15) As Variant

    ' Layer 1: Uratanna Fm
    data(0, 0) = "Paleozoic"
    data(0, 1) = "Cambrian"
    data(0, 2) = "Terreneuvian"
    data(0, 3) = "Uratanna Fm"
    data(0, 4) = "∈₁u"
    data(0, 5) = "trilobite"
    data(0, 6) = 120
    data(0, 7) = "Feldspathic sandstone with Skolithos"
    data(0, 8) = "sand"
    data(0, 9) = 0: data(0, 10) = 20
    data(0, 11) = 30: data(0, 12) = 10
    data(0, 13) = 3
    data(0, 14) = ""
    data(0, 15) = "disconformity"

    ' Layer 2: Pound Ss
    data(1, 0) = "Neoproterozoic"
    data(1, 1) = "Ediacaran"
    data(1, 2) = "Upper"
    data(1, 3) = "Pound Ss"
    data(1, 4) = "Z₂p"
    data(1, 5) = "algae"
    data(1, 6) = 180
    data(1, 7) = "Quartzite to arkosic sandstone with Ediacara Biota"
    data(1, 8) = "sand"
    data(1, 9) = 0: data(1, 10) = 10
    data(1, 11) = 25: data(1, 12) = 8
    data(1, 13) = 3
    data(1, 14) = ""
    data(1, 15) = ""

    ' Layer 3: Bonney Ss
    data(2, 0) = "Neoproterozoic"
    data(2, 1) = "Ediacaran"
    data(2, 2) = "Upper"
    data(2, 3) = "Bonney Ss"
    data(2, 4) = "Z₂b"
    data(2, 5) = ""
    data(2, 6) = 95
    data(2, 7) = "Fine-grained sandstone with ripple marks"
    data(2, 8) = "finesand"
    data(2, 9) = 0: data(2, 10) = 5
    data(2, 11) = 20: data(2, 12) = 8
    data(2, 13) = 3
    data(2, 14) = "ripple"
    data(2, 15) = ""

    ' Layer 4: Wonoka Fm
    data(3, 0) = "Neoproterozoic"
    data(3, 1) = "Ediacaran"
    data(3, 2) = "Middle"
    data(3, 3) = "Wonoka Fm"
    data(3, 4) = "Z₁w"
    data(3, 5) = ""
    data(3, 6) = 350
    data(3, 7) = "Grey limestone with δ¹³C excursion (~−8‰ Shuram)"
    data(3, 8) = "lime"
    data(3, 9) = 0: data(3, 10) = 0
    data(3, 11) = 0: data(3, 12) = 30
    data(3, 13) = 2
    data(3, 14) = ""
    data(3, 15) = "conformity"

    ' Layer 5: Bunyeroo Fm
    data(4, 0) = "Neoproterozoic"
    data(4, 1) = "Ediacaran"
    data(4, 2) = "Middle"
    data(4, 3) = "Bunyeroo Fm"
    data(4, 4) = "Z₁b"
    data(4, 5) = ""
    data(4, 6) = 140
    data(4, 7) = "Red-brown siltstone with carbonate concretions"
    data(4, 8) = "silt"
    data(4, 9) = 0: data(4, 10) = 15
    data(4, 11) = 30: data(4, 12) = 15
    data(4, 13) = 2
    data(4, 14) = "concretion"
    data(4, 15) = ""

    ' Layer 6: ABC Range Quartzite
    data(5, 0) = "Neoproterozoic"
    data(5, 1) = "Ediacaran"
    data(5, 2) = "Lower"
    data(5, 3) = "ABC Range Quartzite"
    data(5, 4) = "Z₁a"
    data(5, 5) = ""
    data(5, 6) = 280
    data(5, 7) = "Massive cross-bedded quartzite"
    data(5, 8) = "sand"
    data(5, 9) = 0: data(5, 10) = 0
    data(5, 11) = 10: data(5, 12) = 10
    data(5, 13) = 4
    data(5, 14) = "cross_bed"
    data(5, 15) = ""

    ' Layer 7: Elatina Fm
    data(6, 0) = "Neoproterozoic"
    data(6, 1) = "Ediacaran"
    data(6, 2) = "Lower"
    data(6, 3) = "Elatina Fm"
    data(6, 4) = "Z₁e"
    data(6, 5) = ""
    data(6, 6) = 65
    data(6, 7) = "Pebbly diamictite — Marinoan glacial"
    data(6, 8) = "conglo"
    data(6, 9) = 50: data(6, 10) = 20
    data(6, 11) = 50: data(6, 12) = 30
    data(6, 13) = 6
    data(6, 14) = ""
    data(6, 15) = "unconformity"

    ' Layer 8: Yaltipena Fm
    data(7, 0) = "Neoproterozoic"
    data(7, 1) = "Cryogenian"
    data(7, 2) = "Upper"
    data(7, 3) = "Yaltipena Fm"
    data(7, 4) = "Nh₂y"
    data(7, 5) = ""
    data(7, 6) = 120
    data(7, 7) = "Grey-green massive diamictite"
    data(7, 8) = "conglo"
    data(7, 9) = 45: data(7, 10) = 15
    data(7, 11) = 40: data(7, 12) = 25
    data(7, 13) = 6
    data(7, 14) = ""
    data(7, 15) = ""

    ' === Draw each layer ===
    Dim i As Long
    For i = 0 To UBound(data, 1)
        Dim layH As Double: layH = data(i, 6) * scaleF
        If layH < 3.5 Then layH = 3.5
        Dim layB As Double: layB = currentY - layH
        Dim midY As Double: midY = (currentY + layB) / 2

        ' Lithology rect (col 7)
        Dim rect As Shape
        Set rect = ActiveLayer.CreateRectangle(CX7 + 0.3, layB, CX7 + CW7 - 2.5, currentY)
        rect.Fill.UniformColor.CMYKAssign data(i, 9), data(i, 10), data(i, 11), data(i, 12)
        rect.Outline.Color.CMYKAssign 0, 0, 0, 35
        rect.Outline.Width = 0.15

        ' Text columns
        AddCellText ActiveLayer, CX1, currentY, layB, CW1, data(i, 0)
        AddCellText ActiveLayer, CX2, currentY, layB, CW2, data(i, 1)
        AddCellText ActiveLayer, CX3, currentY, layB, CW3, data(i, 2)
        AddCellText ActiveLayer, CX4, currentY, layB, CW4, data(i, 3)
        AddCellText ActiveLayer, CX5, currentY, layB, CW5, data(i, 4)

        ' Fossil column (col 6)
        If data(i, 5) <> "" Then
            AddSmallText ActiveLayer, CX6 + CW6/2, midY, data(i, 5), 4.5
        End If

        ' Thickness (col 10)
        AddCellText ActiveLayer, CX10, currentY, layB, CW10, Format(data(i, 6), "0.0")

        ' Grain triangle (col 8)
        If data(i, 13) > 0 Then
            DrawGrainTriangle ActiveLayer, CX8 + 1.5, currentY, layB, CW8 - 3, data(i, 13)
        End If

        ' Structure column (col 9)
        If data(i, 14) <> "" Then
            AddSmallText ActiveLayer, CX9 + CW9/2, midY, data(i, 14), 4
        End If

        ' Description (col 11)
        AddCellText ActiveLayer, CX11, currentY, layB, CW11, data(i, 7)

        ' Contact line
        If data(i, 15) <> "" Then
            DrawContactLine ActiveLayer, CX7 + CW7/2, layB, data(i, 15)
        End If

        currentY = layB
    Next i

    ' === Borders ===
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX1, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX11 + CW11, TABLE_BOT, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_TOP, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX11 + CW11, TABLE_BOT): .Outline.Width = 0.4: End With

    ' === Footer ===
    Dim ft As Shape
    Set ft = ActiveLayer.CreateArtisticText(CX1, TABLE_BOT - 5, "Integrated Ediacaran Stratigraphy — Eastern Gondwana Margin | Total 1,350m | 8 layers | Scale 1:5,314")
    ft.Text.Story.Font = "SimHei": ft.Text.Story.Size = 5
    ft.Fill.UniformColor.CMYKAssign 0, 0, 0, 50

    ActiveDocument.EndCommandGroup
    MsgBox "Drawing completed!"
    Exit Sub

ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub


' ════════════════════════════════
' Helper Procedures
' ════════════════════════════════

Private Sub AddHeaderText(layer As Layer, cx As Double, cy As Double, txt As String)
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 7
    s.Text.Story.Bold = True
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddCellText(layer As Layer, cx As Double, topY As Double, botY As Double, cw As Double, txt As String)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Dim midY As Double: midY = (topY + botY) / 2
    Set s = layer.CreateArtisticText(cx + 1, midY, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = 5.5
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 100
End Sub

Private Sub AddSmallText(layer As Layer, cx As Double, cy As Double, txt As String, sz As Double)
    If txt = "" Then Exit Sub
    Dim s As Shape
    Set s = layer.CreateArtisticText(cx, cy, txt)
    s.Text.Story.Font = "SimHei"
    s.Text.Story.Size = sz
    s.Text.Story.Alignment = cdrCenterAlignment
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 60
End Sub

Private Sub DrawGrainTriangle(layer As Layer, gsx As Double, topY As Double, botY As Double, gsw As Double, gv As Long)
    ' Draw right triangle for grain-size indicator (1-6)
    Dim w As Double: w = gsw * gv / 6
    If w < 0.5 Then w = 0.5
    Dim crv As Curve
    Set crv = Application.CreateCurve()
    crv.CreateSubPath gsx, botY
    crv.AppendLineSegment gsx + w, botY
    crv.AppendLineSegment gsx + w, topY
    crv.AppendLineSegment gsx, topY
    crv.Closed = True
    Dim s As Shape: Set s = layer.CreateCurve(crv)
    s.Fill.UniformColor.CMYKAssign 0, 0, 0, 40
    s.Outline.SetNoOutline
End Sub

Private Sub DrawContactLine(layer As Layer, cx As Double, cy As Double, contactType As String)
    Dim s As Shape
    If contactType = "unconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "~~~ 不整合 ~~~")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    ElseIf contactType = "disconformity" Then
        Set s = layer.CreateArtisticText(cx, cy, "--- 假整合 ---")
        s.Text.Story.Font = "SimHei": s.Text.Story.Size = 4
        s.Text.Story.Alignment = cdrCenterAlignment
        s.Fill.UniformColor.CMYKAssign 0, 100, 80, 0
    End If
End Sub

' ════════════════════════════════
' Main Entry Point
' ════════════════════════════════

Public Sub DrawSection3()
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        MsgBox "Please open a document first"
        Exit Sub
    End If

    ActiveDocument.BeginCommandGroup "Draw Stratigraphic Column"
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft

    ' === Column layout (11 cols) ===
    Const CX1 = 8
    Const CW1 = 13
    Const CX2 = 21
    Const CW2 = 13
    Const CX3 = 34
    Const CW3 = 13
    Const CX4 = 47
    Const CW4 = 18
    Const CX5 = 65
    Const CW5 = 12
    Const CX6 = 77
    Const CW6 = 7
    Const CX7 = 84
    Const CW7 = 40
    Const CX8 = 124
    Const CW8 = 15
    Const CX9 = 139
    Const CW9 = 7
    Const CX10 = 146
    Const CW10 = 10
    Const CX11 = 156
    Const CW11 = 65
    Const TABLE_TOP = 290
    Const TABLE_BOT = 22
    Const HEADER_H = 14

    Dim scaleF As Double: scaleF = 0.49901768172888017
    Dim totalThick As Double: totalThick = 509

    ' === Title ===
    Dim t As Shape
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 8, "Nantuo-Doushantuo Transition — Snowball Earth Cap Carbonate")
    With t.Text.Story: .Font = "SimHei": .Size = 12: .Bold = True: .Alignment = cdrCenterAlignment: End With
    Set t = ActiveLayer.CreateArtisticText((CX1 + CX11 + CW11)/2, TABLE_TOP + 2, "Jiulongwan, Three Gorges, South China")
    With t.Text.Story: .Font = "SimHei": .Size = 7: .Alignment = cdrCenterAlignment: End With

    ' === Header ===
    Dim hdr As Shape
    Set hdr = ActiveLayer.CreateRectangle(CX1, TABLE_TOP - HEADER_H, CX11 + CW11, TABLE_TOP)
    hdr.Fill.UniformColor.CMYKAssign 0, 0, 0, 5
    hdr.Outline.Color.CMYKAssign 0, 0, 0, 100
    hdr.Outline.Width = 0.35

    AddHeaderText ActiveLayer, CX1 + CW1/2, TABLE_TOP - HEADER_H/2, "界"
    AddHeaderText ActiveLayer, CX2 + CW2/2, TABLE_TOP - HEADER_H/2, "系"
    AddHeaderText ActiveLayer, CX3 + CW3/2, TABLE_TOP - HEADER_H/2, "统"
    AddHeaderText ActiveLayer, CX4 + CW4/2, TABLE_TOP - HEADER_H/2, "组"
    AddHeaderText ActiveLayer, CX5 + CW5/2, TABLE_TOP - HEADER_H/2, "代号"
    AddHeaderText ActiveLayer, CX6 + CW6/2, TABLE_TOP - HEADER_H/2, "化石"
    AddHeaderText ActiveLayer, CX7 + CW7/2, TABLE_TOP - HEADER_H/2, "岩性柱"
    AddHeaderText ActiveLayer, CX8 + CW8/2, TABLE_TOP - HEADER_H/2, "粒度"
    AddHeaderText ActiveLayer, CX9 + CW9/2, TABLE_TOP - HEADER_H/2, "构造"
    AddHeaderText ActiveLayer, CX10 + CW10/2, TABLE_TOP - HEADER_H/2, "厚度(m)"
    AddHeaderText ActiveLayer, CX11 + CW11/2, TABLE_TOP - HEADER_H/2, "岩性描述"

    ' Vertical dividers
    ActiveLayer.CreateLine CX2, TABLE_TOP, CX2, TABLE_BOT
    ActiveLayer.CreateLine CX3, TABLE_TOP, CX3, TABLE_BOT
    ActiveLayer.CreateLine CX4, TABLE_TOP, CX4, TABLE_BOT
    ActiveLayer.CreateLine CX5, TABLE_TOP, CX5, TABLE_BOT
    ActiveLayer.CreateLine CX6, TABLE_TOP, CX6, TABLE_BOT
    ActiveLayer.CreateLine CX7, TABLE_TOP, CX7, TABLE_BOT
    ActiveLayer.CreateLine CX8, TABLE_TOP, CX8, TABLE_BOT
    ActiveLayer.CreateLine CX9, TABLE_TOP, CX9, TABLE_BOT
    ActiveLayer.CreateLine CX10, TABLE_TOP, CX10, TABLE_BOT
    ActiveLayer.CreateLine CX11, TABLE_TOP, CX11, TABLE_BOT

    ' === Layer data (8 layers) ===
    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H
    Dim data(7, 15) As Variant

    ' Layer 1: Yanjiahe Fm
    data(0, 0) = "Paleozoic"
    data(0, 1) = "Cambrian"
    data(0, 2) = "Terreneuvian"
    data(0, 3) = "Yanjiahe Fm"
    data(0, 4) = "∈₁y"
    data(0, 5) = "brachiopod"
    data(0, 6) = 35
    data(0, 7) = "Siliceous shale with SSF"
    data(0, 8) = "shale"
    data(0, 9) = 0: data(0, 10) = 0
    data(0, 11) = 5: data(0, 12) = 50
    data(0, 13) = 1
    data(0, 14) = ""
    data(0, 15) = ""

    ' Layer 2: Dengying Fm
    data(1, 0) = "Neoproterozoic"
    data(1, 1) = "Ediacaran"
    data(1, 2) = "Upper"
    data(1, 3) = "Dengying Fm"
    data(1, 4) = "Z₂dy"
    data(1, 5) = ""
    data(1, 6) = 280
    data(1, 7) = "Thick-bedded peritidal dolomite"
    data(1, 8) = "dolo"
    data(1, 9) = 0: data(1, 10) = 0
    data(1, 11) = 3: data(1, 12) = 8
    data(1, 13) = 2
    data(1, 14) = ""
    data(1, 15) = "disconformity"

    ' Layer 3: Doushantuo Fm
    data(2, 0) = "Neoproterozoic"
    data(2, 1) = "Ediacaran"
    data(2, 2) = "Lower"
    data(2, 3) = "Doushantuo Fm"
    data(2, 4) = "Z₁d"
    data(2, 5) = "algae spore"
    data(2, 6) = 70
    data(2, 7) = "Black shale with large acanthomorphic acritarchs"
    data(2, 8) = "shale"
    data(2, 9) = 0: data(2, 10) = 5
    data(2, 11) = 10: data(2, 12) = 50
    data(2, 13) = 1
    data(2, 14) = ""
    data(2, 15) = ""

    ' Layer 4: Doushantuo Fm
    data(3, 0) = "Neoproterozoic"
    data(3, 1) = "Ediacaran"
    data(3, 2) = "Lower"
    data(3, 3) = "Doushantuo Fm"
    data(3, 4) = "Z₁d"
    data(3, 5) = ""
    data(3, 6) = 25
    data(3, 7) = "Grey dolomitic limestone (upper cap)"
    data(3, 8) = "doloLime"
    data(3, 9) = 0: data(3, 10) = 0
    data(3, 11) = 5: data(3, 12) = 20
    data(3, 13) = 2
    data(3, 14) = ""
    data(3, 15) = ""

    ' Layer 5: Doushantuo Fm
    data(4, 0) = "Neoproterozoic"
    data(4, 1) = "Ediacaran"
    data(4, 2) = "Lower"
    data(4, 3) = "Doushantuo Fm"
    data(4, 4) = "Z₁d"
    data(4, 5) = ""
    data(4, 6) = 8
    data(4, 7) = "Pink crystalline dolomite (middle cap)"
    data(4, 8) = "dolo"
    data(4, 9) = 0: data(4, 10) = 0
    data(4, 11) = 10: data(4, 12) = 8
    data(4, 13) = 2
    data(4, 14) = ""
    data(4, 15) = ""

    ' Layer 6: Doushantuo Fm
    data(5, 0) = "Neoproterozoic"
    data(5, 1) = "Ediacaran"
    data(5, 2) = "Lower"
    data(5, 3) = "Doushantuo Fm"
    data(5, 4) = "Z₁d"
    data(5, 5) = ""
    data(5, 6) = 6
    data(5, 7) = "Pale grey microcrystalline dolomite with tepee structures (lower cap)"
    data(5, 8) = "dolo"
    data(5, 9) = 0: data(5, 10) = 0
    data(5, 11) = 5: data(5, 12) = 10
    data(5, 13) = 2
    data(5, 14) = "crack"
    data(5, 15) = ""

    ' Layer 7: Nantuo Fm
    data(6, 0) = "Neoproterozoic"
    data(6, 1) = "Cryogenian"
    data(6, 2) = "Upper"
    data(6, 3) = "Nantuo Fm"
    data(6, 4) = "Nh₂n"
    data(6, 5) = ""
    data(6, 6) = 55
    data(6, 7) = "Green-grey massive diamictite with striated clasts"
    data(6, 8) = "conglo"
    data(6, 9) = 50: data(6, 10) = 20
    data(6, 11) = 50: data(6, 12) = 30
    data(6, 13) = 6
    data(6, 14) = ""
    data(6, 15) = "unconformity"

    ' Layer 8: Nantuo Fm
    data(7, 0) = "Neoproterozoic"
    data(7, 1) = "Cryogenian"
    data(7, 2) = "Upper"
    data(7, 3) = "Nantuo Fm"
    data(7, 4) = "Nh₂n"
    data(7, 5) = ""
    data(7, 6) = 30
    data(7, 7) = "Dark grey sandy diamictite with dropstones"
    data(7, 8) = "conglo"
    data(7, 9) = 30: data(7, 10) = 10
    data(7, 11) = 25: data(7, 12) = 25
    data(7, 13) = 5
    data(7, 14) = ""
    data(7, 15) = ""

    ' === Draw each layer ===
    Dim i As Long
    For i = 0 To UBound(data, 1)
        Dim layH As Double: layH = data(i, 6) * scaleF
        If layH < 3.5 Then layH = 3.5
        Dim layB As Double: layB = currentY - layH
        Dim midY As Double: midY = (currentY + layB) / 2

        ' Lithology rect (col 7)
        Dim rect As Shape
        Set rect = ActiveLayer.CreateRectangle(CX7 + 0.3, layB, CX7 + CW7 - 2.5, currentY)
        rect.Fill.UniformColor.CMYKAssign data(i, 9), data(i, 10), data(i, 11), data(i, 12)
        rect.Outline.Color.CMYKAssign 0, 0, 0, 35
        rect.Outline.Width = 0.15

        ' Text columns
        AddCellText ActiveLayer, CX1, currentY, layB, CW1, data(i, 0)
        AddCellText ActiveLayer, CX2, currentY, layB, CW2, data(i, 1)
        AddCellText ActiveLayer, CX3, currentY, layB, CW3, data(i, 2)
        AddCellText ActiveLayer, CX4, currentY, layB, CW4, data(i, 3)
        AddCellText ActiveLayer, CX5, currentY, layB, CW5, data(i, 4)

        ' Fossil column (col 6)
        If data(i, 5) <> "" Then
            AddSmallText ActiveLayer, CX6 + CW6/2, midY, data(i, 5), 4.5
        End If

        ' Thickness (col 10)
        AddCellText ActiveLayer, CX10, currentY, layB, CW10, Format(data(i, 6), "0.0")

        ' Grain triangle (col 8)
        If data(i, 13) > 0 Then
            DrawGrainTriangle ActiveLayer, CX8 + 1.5, currentY, layB, CW8 - 3, data(i, 13)
        End If

        ' Structure column (col 9)
        If data(i, 14) <> "" Then
            AddSmallText ActiveLayer, CX9 + CW9/2, midY, data(i, 14), 4
        End If

        ' Description (col 11)
        AddCellText ActiveLayer, CX11, currentY, layB, CW11, data(i, 7)

        ' Contact line
        If data(i, 15) <> "" Then
            DrawContactLine ActiveLayer, CX7 + CW7/2, layB, data(i, 15)
        End If

        currentY = layB
    Next i

    ' === Borders ===
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX1, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX11 + CW11, TABLE_BOT, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_TOP, CX11 + CW11, TABLE_TOP): .Outline.Width = 0.4: End With
    With ActiveLayer.CreateLine(CX1, TABLE_BOT, CX11 + CW11, TABLE_BOT): .Outline.Width = 0.4: End With

    ' === Footer ===
    Dim ft As Shape
    Set ft = ActiveLayer.CreateArtisticText(CX1, TABLE_BOT - 5, "Nantuo-Doushantuo Transition — Snowball Earth Cap Carbonate | Total 509m | 8 layers | Scale 1:2,003")
    ft.Text.Story.Font = "SimHei": ft.Text.Story.Size = 5
    ft.Fill.UniformColor.CMYKAssign 0, 0, 0, 50

    ActiveDocument.EndCommandGroup
    MsgBox "Drawing completed!"
    Exit Sub

ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub

' Master sub — draws all sections
Public Sub DrawAll()
    On Error GoTo ErrHandler
    ActiveDocument.BeginCommandGroup "Batch Stratigraphic Columns"
    ActiveDocument.Unit = cdrMillimeter
    DrawSection1
    DrawSection2
    DrawSection3
    ActiveDocument.EndCommandGroup
    MsgBox "Batch complete!"
    Exit Sub
ErrHandler:
    ActiveDocument.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub