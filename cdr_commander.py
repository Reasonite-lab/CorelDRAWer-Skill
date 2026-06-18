#!/usr/bin/env python3
"""
CDR Commander — VBA script generator for batch CorelDRAW document operations.
Generates ready-to-run VBA that:
- Batch imports SVG files as layers
- Aligns and distributes objects
- Applies consistent styles across documents
- Exports document stats to JSON
- Merges multiple .cdr documents

Usage: python3 cdr_commander.py <command> [args]
  import-svg folder/          → VBA to import all SVGs as layers
  style-apply style.json      → VBA to apply styles to current doc
  export-stats                → VBA to export document stats to JSON
  merge-docs file1.cdr ...    → VBA to merge CDR files
"""

import json, sys, os, glob

def gen_import_svgs(folder_path, output='cdr_import.bas'):
    """Generate VBA that batch imports SVG files into separate layers."""
    svg_files = glob.glob(os.path.join(folder_path, '*.svg'))
    if not svg_files:
        svg_files = glob.glob(os.path.join(folder_path, '**', '*.svg'), recursive=True)

    lines = []
    lines.append("' ═════════════════════════════════")
    lines.append(f"' CDR Commander — SVG Batch Import ({len(svg_files)} files)")
    lines.append("' ═════════════════════════════════")
    lines.append("Option Explicit")
    lines.append("")
    lines.append("Public Sub ImportAllSVGs()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    Dim doc As Document: Set doc = ActiveDocument")
    lines.append("    doc.BeginCommandGroup \"SVG Batch Import\"")
    lines.append("    doc.Unit = cdrMillimeter")
    lines.append("")
    lines.append("    Dim i As Long")
    lines.append(f"    For i = 1 To {len(svg_files)}")
    lines.append("        Dim layerName As String")

    for i, svg in enumerate(svg_files):
        name = os.path.basename(svg).replace('.svg', '')
        abs_path = os.path.abspath(svg).replace('\\', '\\\\')
        lines.append(f"        If i = {i+1} Then")
        lines.append(f"            layerName = \"{name}\"")
        lines.append(f"            doc.ActivePage.CreateLayer layerName")
        lines.append(f"            doc.ActivePage.Layers(layerName).Activate")
        lines.append(f"            doc.ActiveLayer.Import \"{abs_path}\"")
        lines.append(f"        End If")

    lines.append("    Next i")
    lines.append("")
    lines.append("    ' Activate first layer")
    lines.append("    doc.ActivePage.Layers(1).Activate")
    lines.append("    doc.EndCommandGroup")
    lines.append("    MsgBox \"Imported \" & i-1 & \" SVGs\"")
    lines.append("    Exit Sub")
    lines.append("ErrHandler:")
    lines.append("    doc.EndCommandGroup")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    vba = '\n'.join(lines)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(vba)
    print(f'✅ CDR SVG import VBA saved: {output}')
    print(f'   Open CorelDRAW → Alt+F11 → Import → Run ImportAllSVGs')
    return vba

def gen_style_apply(style_config, output='cdr_style.bas'):
    """Generate VBA that applies consistent styles to all objects in a document."""
    lines = []
    lines.append("' ═════════════════════════════════")
    lines.append(f"' CDR Commander — Style Applicator")
    lines.append("' ═════════════════════════════════")
    lines.append("Option Explicit")
    lines.append("")
    lines.append("Public Sub ApplyStyles()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    Dim doc As Document: Set doc = ActiveDocument")
    lines.append("    doc.BeginCommandGroup \"Apply Styles\"")
    lines.append("")
    lines.append("    ' --- Page Setup ---")
    page = style_config.get('page', {})
    if page:
        w = page.get('width', 210)
        h = page.get('height', 297)
        lines.append(f"    doc.ActivePage.SetSize {w}, {h}")
        lines.append("    doc.Unit = cdrMillimeter")
    lines.append("")

    lines.append("    ' --- Text Styles ---")
    text_styles = style_config.get('text_styles', {})
    for style_name, props in text_styles.items():
        font = props.get('font', 'Arial')
        size = props.get('size', 10)
        bold = props.get('bold', False)
        color = props.get('color', 'CMYK,0,0,0,100')
        lines.append(f"    ' Style: {style_name}")
        lines.append(f"    Dim sr{style_name} As ShapeRange")
        lines.append(f"    Set sr{style_name} = doc.ActivePage.FindShapes(Type:=cdrArtisticTextType)")
        lines.append(f"    If sr{style_name}.Count > 0 Then")
        lines.append(f"        sr{style_name}.Text.Story.Font = \"{font}\"")
        lines.append(f"        sr{style_name}.Text.Story.Size = {size}")
        lines.append(f"        If {str(bold).lower()} Then sr{style_name}.Text.Story.Bold = True")
        lines.append(f"        sr{style_name}.Fill.UniformColor.StringAssign \"{color}\"")
        lines.append(f"    End If")
        lines.append("")
    lines.append("    doc.EndCommandGroup")
    lines.append("    MsgBox \"Styles applied!\"")
    lines.append("    Exit Sub")
    lines.append("ErrHandler:")
    lines.append("    doc.EndCommandGroup")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    vba = '\n'.join(lines)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(vba)
    print(f'✅ CDR style VBA saved: {output}')
    return vba

def gen_export_stats(output='cdr_stats.bas'):
    """Generate VBA that exports document statistics to a JSON file."""
    lines = []
    lines.append("' ═════════════════════════════════")
    lines.append("' CDR Commander — Document Stats Exporter")
    lines.append("' ═════════════════════════════════")
    lines.append("Option Explicit")
    lines.append("")
    lines.append("Public Sub ExportStats()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    Dim doc As Document: Set doc = ActiveDocument")
    lines.append("    Dim fso As Object: Set fso = CreateObject(\"Scripting.FileSystemObject\")")
    lines.append("    Dim outFile As Object")
    lines.append("    Dim json As String")
    lines.append("")
    lines.append("    json = \"{\" & vbCrLf")
    lines.append("    json = json & \"  \"\"title\"\": \"\"\" & doc.FileName & \"\"\",\" & vbCrLf")
    lines.append("    json = json & \"  \"\"pages\"\": \" & doc.Pages.Count & \",\" & vbCrLf")
    lines.append("    json = json & \"  \"\"layers\"\": \" & doc.ActivePage.Layers.Count & \",\" & vbCrLf")
    lines.append("    json = json & \"  \"\"shapes\"\": \" & doc.ActivePage.Shapes.Count & \",\" & vbCrLf")
    lines.append("    json = json & \"  \"\"width_mm\"\": \" & doc.ActivePage.SizeWidth & \",\" & vbCrLf")
    lines.append("    json = json & \"  \"\"height_mm\"\": \" & doc.ActivePage.SizeHeight & \"\"\" & vbCrLf")
    lines.append("    json = json & \"}\"")
    lines.append("")
    lines.append("    ' Save to desktop")
    lines.append("    Dim path As String")
    lines.append("    path = Environ(\"USERPROFILE\") & \"\\Desktop\\cdr_stats.json\"")
    lines.append("    Set outFile = fso.CreateTextFile(path, True)")
    lines.append("    outFile.Write json")
    lines.append("    outFile.Close")
    lines.append("")
    lines.append("    MsgBox \"Stats exported to \" & path")
    lines.append("    Exit Sub")
    lines.append("ErrHandler:")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    vba = '\n'.join(lines)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(vba)
    print(f'✅ CDR stats VBA saved: {output}')
    return vba

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    if cmd == 'import-svg':
        folder = sys.argv[2] if len(sys.argv) > 2 else '.'
        gen_import_svgs(folder)
    elif cmd == 'style-apply':
        config_path = sys.argv[2] if len(sys.argv) > 2 else None
        config = {}
        if config_path:
            with open(config_path) as f:
                config = json.load(f)
        gen_style_apply(config)
    elif cmd == 'export-stats':
        gen_export_stats()
    else:
        print(f'Unknown: {cmd}')
        print(__doc__)
