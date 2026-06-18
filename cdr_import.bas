' ═════════════════════════════════
' CDR Commander — SVG Batch Import (20 files)
' ═════════════════════════════════
Option Explicit

Public Sub ImportAllSVGs()
    On Error GoTo ErrHandler
    Dim doc As Document: Set doc = ActiveDocument
    doc.BeginCommandGroup "SVG Batch Import"
    doc.Unit = cdrMillimeter

    Dim i As Long
    For i = 1 To 20
        Dim layerName As String
        If i = 1 Then
            layerName = "snowball_earth"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/snowball_earth.svg"
        End If
        If i = 2 Then
            layerName = "sao_luis_basin"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/sao_luis_basin.svg"
        End If
        If i = 3 Then
            layerName = "sao_luis_nature"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/sao_luis_nature.svg"
        End If
        If i = 4 Then
            layerName = "adelaide_ediacaran"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/adelaide_ediacaran.svg"
        End If
        If i = 5 Then
            layerName = "sao_luis_nature_final"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/sao_luis_nature_final.svg"
        End If
        If i = 6 Then
            layerName = "yanjiahe_ediacaran"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/yanjiahe_ediacaran.svg"
        End If
        If i = 7 Then
            layerName = "three_gorges_transect"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/three_gorges_transect.svg"
        End If
        If i = 8 Then
            layerName = "sao_luis_default"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/sao_luis_default.svg"
        End If
        If i = 9 Then
            layerName = "秭归地层柱状图"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/秭归地层柱状图.svg"
        End If
        If i = 10 Then
            layerName = "cross_section_demo"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/cross_section_demo.svg"
        End If
        If i = 11 Then
            layerName = "sample_from_csv"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/sample_from_csv.svg"
        End If
        If i = 12 Then
            layerName = "icon"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/icon.svg"
        End If
        If i = 13 Then
            layerName = "准噶尔_ZJ14井"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/准噶尔_ZJ14井.svg"
        End If
        If i = 14 Then
            layerName = "meishan_gssp"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/meishan_gssp.svg"
        End If
        If i = 15 Then
            layerName = "adelaide_xsection"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/adelaide_xsection.svg"
        End If
        If i = 16 Then
            layerName = "cap_carbonate_transect"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/cap_carbonate_transect.svg"
        End If
        If i = 17 Then
            layerName = "output"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/output.svg"
        End If
        If i = 18 Then
            layerName = "logo"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/logo.svg"
        End If
        If i = 19 Then
            layerName = "multi_demo"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/multi_demo.svg"
        End If
        If i = 20 Then
            layerName = "秭归地区综合地层柱状图"
            doc.ActivePage.CreateLayer layerName
            doc.ActivePage.Layers(layerName).Activate
            doc.ActiveLayer.Import "/Users/xubk/Desktop/Corealdrew/秭归地区综合地层柱状图.svg"
        End If
    Next i

    ' Activate first layer
    doc.ActivePage.Layers(1).Activate
    doc.EndCommandGroup
    MsgBox "Imported " & i-1 & " SVGs"
    Exit Sub
ErrHandler:
    doc.EndCommandGroup
    MsgBox "Error: " & Err.Description
End Sub