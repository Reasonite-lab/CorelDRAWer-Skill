' ═══════════════════════════════════════════════════
' CorelDRAWer Export Macro
' Run in CorelDRAW VBA: Alt+F11 → Import → Run ExportToJSON
' Exports document layers and shapes to a JSON file
' ═══════════════════════════════════════════════════

Public Sub ExportToJSON()
    On Error GoTo ErrHandler
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    Dim fso As Object, outFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outPath As String
    outPath = doc.FilePath & "\" & Replace(doc.FileName, ".cdr", "_export.json")
    
    Set outFile = fso.CreateTextFile(outPath, True)
    
    outFile.WriteLine "{"
    outFile.WriteLine "  ""file"": """ & Replace(doc.FileName, Chr(34), "\" & Chr(34)) & ""","
    outFile.WriteLine "  ""pages"": " & doc.Pages.Count & ","
    outFile.WriteLine "  ""layers"": ["
    
    Dim pg As Page, lr As Layer, sh As Shape
    Dim layerIdx As Long, shapeIdx As Long
    layerIdx = 0
    
    For Each pg In doc.Pages
        For Each lr In pg.Layers
            If layerIdx > 0 Then outFile.WriteLine ","
            outFile.WriteLine "    {"
            outFile.WriteLine "      ""name"": """ & lr.Name & ""","
            outFile.WriteLine "      ""shapes"": ["
            
            shapeIdx = 0
            For Each sh In lr.Shapes
                If shapeIdx > 0 Then outFile.WriteLine ","
                outFile.Write "        {""type"": " & sh.Type
                outFile.Write ", ""x"": " & sh.LeftX
                outFile.Write ", ""y"": " & sh.BottomY
                outFile.Write ", ""w"": " & sh.SizeWidth
                outFile.Write ", ""h"": " & sh.SizeHeight
                
                If sh.Type = 6 Then  ' cdrTextShape
                    Dim txt As String
                    txt = Replace(sh.Text.Story.Text, Chr(34), "'")
                    txt = Replace(txt, vbCrLf, " ")
                    outFile.Write ", ""text"": """ & txt & """"
                End If
                
                If sh.Fill.Type = 1 Then  ' cdrUniformFill
                    Dim c As Long, m As Long, y As Long, k As Long
                    sh.Fill.UniformColor.GetCMYK c, m, y, k
                    outFile.Write ", ""cmyk"": [" & c & "," & m & "," & y & "," & k & "]"
                End If
                
                outFile.Write "}"
                shapeIdx = shapeIdx + 1
            Next
            
            outFile.WriteLine ""
            outFile.Write "      ]}"
            layerIdx = layerIdx + 1
        Next
    Next
    
    outFile.WriteLine ""
    outFile.WriteLine "  ]"
    outFile.WriteLine "}"
    outFile.Close
    
    MsgBox "Exported: " & outPath, vbInformation, "CorelDRAWer Export"
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
