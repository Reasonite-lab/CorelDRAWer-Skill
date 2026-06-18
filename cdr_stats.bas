' ═════════════════════════════════
' CDR Commander — Document Stats Exporter
' ═════════════════════════════════
Option Explicit

Public Sub ExportStats()
    On Error GoTo ErrHandler
    Dim doc As Document: Set doc = ActiveDocument
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outFile As Object
    Dim json As String

    json = "{" & vbCrLf
    json = json & "  ""title"": """ & doc.FileName & """," & vbCrLf
    json = json & "  ""pages"": " & doc.Pages.Count & "," & vbCrLf
    json = json & "  ""layers"": " & doc.ActivePage.Layers.Count & "," & vbCrLf
    json = json & "  ""shapes"": " & doc.ActivePage.Shapes.Count & "," & vbCrLf
    json = json & "  ""width_mm"": " & doc.ActivePage.SizeWidth & "," & vbCrLf
    json = json & "  ""height_mm"": " & doc.ActivePage.SizeHeight & """ & vbCrLf
    json = json & "}"

    ' Save to desktop
    Dim path As String
    path = Environ("USERPROFILE") & "\Desktop\cdr_stats.json"
    Set outFile = fso.CreateTextFile(path, True)
    outFile.Write json
    outFile.Close

    MsgBox "Stats exported to " & path
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description
End Sub