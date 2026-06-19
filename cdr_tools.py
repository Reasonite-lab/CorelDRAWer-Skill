#!/usr/bin/env python3
"""
CDR File Tools — parse, inspect, and extract data from CorelDRAW files.
CDR X4+ files are ZIP archives containing RIFF-structured XML data.

Usage:
  python3 cdr_tools.py info file.cdr         # show file structure
  python3 cdr_tools.py extract file.cdr      # extract text/shapes to JSON  
  python3 cdr_tools.py vba-export            # generate VBA macro for CDR→JSON export
"""

import zipfile
import json
import sys
import os
import re
import xml.etree.ElementTree as ET
from collections import defaultdict

def is_cdr(filepath):
    """Check if file appears to be a CorelDRAW document."""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(8)
        # CDR files start with RIFF header
        return header[:4] == b'RIFF' or header[:4] == b'PK\x03\x04'
    except:
        return False

def cdr_info(filepath):
    """Print CDR file structure."""
    if not is_cdr(filepath):
        print(f"❌ Not a CDR file: {filepath}")
        return

    size = os.path.getsize(filepath)
    print(f"📄 {filepath} ({size:,} bytes)")

    if zipfile.is_zipfile(filepath):
        print("   Format: ZIP container (CDR X4+)")
        with zipfile.ZipFile(filepath, 'r') as zf:
            for name in zf.namelist():
                info = zf.getinfo(name)
                print(f"   {'📁' if info.is_dir() else '📄'} {name} ({info.file_size:,} bytes)")
    else:
        print("   Format: RIFF binary (CDR X3 or earlier)")
        # Read RIFF structure
        with open(filepath, 'rb') as f:
            data = f.read(256)
        riff_type = data[8:12]
        print(f"   RIFF type: {riff_type}")

def cdr_extract_text(filepath):
    """Extract text content from CDR file."""
    results = []
    if zipfile.is_zipfile(filepath):
        with zipfile.ZipFile(filepath, 'r') as zf:
            for name in zf.namelist():
                if name.endswith('.xml') or 'content' in name.lower():
                    try:
                        xml_data = zf.read(name)
                        # Try to parse XML and find text nodes
                        text_nodes = re.findall(rb'>([^<>{}\[\]]{2,})<', xml_data)
                        for tn in text_nodes:
                            try:
                                text = tn.decode('utf-8', errors='ignore').strip()
                                if text and len(text) > 1:
                                    results.append(text)
                            except:
                                pass
                    except:
                        pass
    return results

def cdr_extract_to_json(filepath, output_path=None):
    """Extract CDR content to structured JSON."""
    info = {
        'file': os.path.basename(filepath),
        'size': os.path.getsize(filepath),
        'format': 'CDR' + (' X4+ (ZIP)' if zipfile.is_zipfile(filepath) else ' X3 (RIFF)'),
        'pages': [],
        'text_content': []
    }

    if zipfile.is_zipfile(filepath):
        with zipfile.ZipFile(filepath, 'r') as zf:
            for name in zf.namelist():
                entry = {'name': name, 'size': zf.getinfo(name).file_size}
                info['pages'].append(entry)

    info['text_content'] = cdr_extract_text(filepath)

    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(info, f, indent=2, ensure_ascii=False)

    return info

def generate_vba_export_macro():
    """Generate a VBA macro that exports CDR document data to JSON."""
    vba = r"""
' ═══════════════════════════════════════════════════
' CorelDRAWer Export Macro — run in CorelDRAW VBA
' Exports document structure to a JSON-like text file
' ═══════════════════════════════════════════════════

Public Sub ExportToJSON()
    On Error GoTo ErrHandler
    
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Choose output file
    Dim fso As Object, outFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outPath As String
    outPath = doc.FilePath & "\" & Replace(doc.FileName, ".cdr", "_export.json")
    
    Set outFile = fso.CreateTextFile(outPath, True)
    
    ' Write header
    outFile.WriteLine "{"
    outFile.WriteLine "  ""file"": """ & Replace(doc.FileName, """", "\""") & ""","
    outFile.WriteLine "  ""pages"": " & doc.Pages.Count & ","
    outFile.WriteLine "  ""width"": " & doc.Pages(1).SizeWidth & ","
    outFile.WriteLine "  ""height"": " & doc.Pages(1).SizeHeight & ","
    outFile.WriteLine "  ""layers"": ["
    
    ' Iterate layers
    Dim pg As Page, lr As Layer, sh As Shape
    Dim layerIdx As Long, shapeIdx As Long
    layerIdx = 0
    
    For Each pg In doc.Pages
        For Each lr In pg.Layers
            If layerIdx > 0 Then outFile.WriteLine ","
            outFile.WriteLine "    {"
            outFile.WriteLine "      ""name"": """ & Replace(lr.Name, """", "\""") & ""","
            outFile.WriteLine "      ""shapes"": ["
            
            shapeIdx = 0
            For Each sh In lr.Shapes
                If shapeIdx > 0 Then outFile.WriteLine ","
                outFile.Write "        {""type"":""" & sh.Type & """"
                
                ' Text content
                If sh.Type = cdrTextShape Then
                    outFile.Write ", ""text"":""" & Replace(sh.Text.Story.Text, """", "\""") & """"
                End If
                
                ' Position and size
                outFile.Write ", ""x"":" & sh.LeftX
                outFile.Write ", ""y"":" & sh.BottomY
                outFile.Write ", ""w"":" & sh.SizeWidth
                outFile.Write ", ""h"":" & sh.SizeHeight
                
                ' Fill color
                If sh.Fill.Type = cdrUniformFill Then
                    Dim c As Long, m As Long, y As Long, k As Long
                    sh.Fill.UniformColor.GetCMYK c, m, y, k
                    outFile.Write ", ""cmyk"":[" & c & "," & m & "," & y & "," & k & "]"
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
    
    MsgBox "Exported to: " & outPath, vbInformation, "CorelDRAWer Export"
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

' ═══════════════════════════════════════════════════
' Import JSON data and create stratified column
' ═══════════════════════════════════════════════════

Public Sub ImportColumnFromJSON()
    On Error GoTo ErrHandler
    
    ' Select JSON file
    Dim fd As FileDialog
    Set fd = Application.FileDialog(1) ' cdrFileOpen
    fd.Filters.Add "JSON Files", "*.json"
    If Not fd.Show Then Exit Sub
    
    Dim jsonPath As String
    jsonPath = fd.SelectedFiles(1)
    
    ' Read file
    Dim fso As Object, inFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set inFile = fso.OpenTextFile(jsonPath, 1)
    Dim content As String
    content = inFile.ReadAll()
    inFile.Close
    
    ' Simple JSON parse (extract layers array)
    ' This is a simplified parser — for production use, 
    ' use the Python generate_column.py instead
    MsgBox "JSON loaded: " & Len(content) & " chars." & vbCrLf & _
           "For full column generation, use:" & vbCrLf & _
           "python3 generate_column.py " & jsonPath & " output.svg", _
           vbInformation, "CorelDRAWer Import"
    Exit Sub
    
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
"""
    return vba

def main():
    if len(sys.argv) < 2:
        print("CorelDRAWer CDR Tools")
        print("  python3 cdr_tools.py info file.cdr       Show CDR structure")
        print("  python3 cdr_tools.py extract file.cdr    Extract text to JSON")
        print("  python3 cdr_tools.py vba-export          Generate VBA export macro")
        sys.exit(0)

    cmd = sys.argv[1]

    if cmd == 'vba-export':
        vba = generate_vba_export_macro()
        path = 'cdr_export.bas'
        with open(path, 'w', encoding='utf-8') as f:
            f.write(vba)
        print(f"✅ VBA macro saved: {path}")
        print(f"   Open CorelDRAW → Alt+F11 → Import → Run ExportToJSON")

    elif cmd == 'info' and len(sys.argv) > 2:
        cdr_info(sys.argv[2])

    elif cmd == 'extract' and len(sys.argv) > 2:
        info = cdr_extract_to_json(sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
        if not sys.argv[3]:
            print(json.dumps(info, indent=2, ensure_ascii=False))

    else:
        print(f"Unknown command: {cmd}")

if __name__ == '__main__':
    main()
