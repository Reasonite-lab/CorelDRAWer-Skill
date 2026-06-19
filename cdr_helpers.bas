
' ══════════════════════════════════════════════════════
' CorelDRAWer VBA Helper Library
' ── Common operations for geological diagram editing
' ══════════════════════════════════════════════════════

Option Explicit

' ── Page Setup ──

Public Sub SetupGeologyPage()
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    ActivePage.SetSize 297, 420  ' A3 portrait
    ActivePage.Orientation = cdrPortrait
End Sub

' ── Import SVG ──

Public Sub ImportGeoSVG(filePath As String)
    ActiveDocument.Import filePath
    ' Ungroup imported SVG to access layers
    Dim s As Shape
    For Each s In ActivePage.Shapes
        If s.Type = cdrGroupShape Then s.UngroupEx
    Next
End Sub

' ── Layer Management ──

Public Sub ShowOnlyLayer(layerName As String)
    Dim l As Layer
    For Each l In ActivePage.Layers
        l.Visible = (l.Name = layerName)
    Next
End Sub

Public Sub ShowAllLayers()
    Dim l As Layer
    For Each l In ActivePage.Layers
        l.Visible = True
    Next
End Sub

' ── Color Utilities ──

Public Sub ApplyNaturePalette()
    ' Set document default CMYK palette to Nature-style muted colors
    Dim pal As Color
    ' Apply to selected shapes
    Dim sr As ShapeRange
    Set sr = ActiveSelectionRange
    If sr.Count = 0 Then Exit Sub
    sr.ApplyUniformFill.CreateCMYKColor 0, 0, 0, 15  ' light gray
End Sub

' ── Lithology Pattern Helper ──

Public Sub DrawLithoRect(x As Double, y As Double, w As Double, h As Double, _
                         c As Long, m As Long, y2 As Long, k As Long)
    Dim s As Shape
    Set s = ActiveLayer.CreateRectangle(x, y, x + w, y + h)
    s.Fill.UniformColor.CMYKAssign c, m, y2, k
    s.Outline.Color.CMYKAssign 0, 0, 0, 35
    s.Outline.Width = 0.15
End Sub

' ── Batch Export ──

Public Sub ExportAllPagesAsSVG(outputDir As String)
    Dim i As Long
    For i = 1 To ActiveDocument.Pages.Count
        ActiveDocument.Pages(i).Activate
        Dim fname As String
        fname = outputDir & "\page_" & Format(i, "000") & ".svg"
        ActiveDocument.Export fname, cdrSVG
    Next
End Sub

' ── Grid Generator ──

Public Sub DrawGrid(cols As Long, rows As Long, cellW As Double, cellH As Double)
    Dim i As Long, j As Long
    For i = 0 To cols
        ActiveLayer.CreateLine i * cellW, 0, i * cellW, rows * cellH
    Next
    For j = 0 To rows
        ActiveLayer.CreateLine 0, j * cellH, cols * cellW, j * cellH
    Next
End Sub
