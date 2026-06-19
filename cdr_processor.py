#!/usr/bin/env python3
"""
CorelDRAW File Processor v1.0
═══════════════════════════════════════════════════
CDR file handling: inspect, convert, batch process.

Usage:
  python3 cdr_processor.py inspect file.cdr
  python3 cdr_processor.py svg2cdr input.svg          # Windows COM
  python3 cdr_processor.py batch *.svg                 # batch SVG→CDR
  python3 cdr_processor.py template output.cdr         # create template
  python3 cdr_processor.py macro                       # generate VBA helpers

Requirements (for COM features): Windows + CorelDRAW + pywin32
SVG operations: platform-independent
═══════════════════════════════════════════════════
"""

import sys
import os
import struct
import zipfile
import xml.etree.ElementTree as ET
import json

# ── CDR File Inspector ──

def inspect_cdr(filepath):
    """Inspect a CorelDRAW (.cdr) file and extract metadata."""
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return

    size = os.path.getsize(filepath)
    print(f"📄 {filepath}")
    print(f"   Size: {size:,} bytes")

    with open(filepath, 'rb') as f:
        header = f.read(16)

    # CDR format detection
    if header[:4] == b'RIFF':
        print("   Format: RIFF/CDR container")
        # Extract version info from RIFF chunk
        riff_size = struct.unpack('<I', header[4:8])[0]
        cdr_type = header[8:12]
        version_id = struct.unpack('<I', header[12:16])[0]
        versions = {
            1: 'CDR 3', 2: 'CDR 4', 3: 'CDR 5', 4: 'CDR 6', 
            5: 'CDR 7', 6: 'CDR 8', 7: 'CDR 9', 8: 'CDR 10',
            9: 'CDR 11/X3', 10: 'CDR 12/X4', 11: 'CDR 13/X5',
            12: 'CDR 14/X6', 13: 'CDR 15/X7', 14: 'CDR 16/2017',
            15: 'CDR 17/2018', 16: 'CDR 18/2019', 17: 'CDR 19/2020',
            18: 'CDR 20/2021', 19: 'CDR 21/2022', 20: 'CDR 22/2023',
            21: 'CDR 23/2024', 22: 'CDR 24/2025',
        }
        ver_name = versions.get(version_id, f'CDR (v{version_id})')
        print(f"   Type: {cdr_type.decode('latin-1', errors='ignore')}")
        print(f"   Version: {ver_name}")

    elif header[:2] == b'PK':
        print("   Format: ZIP/CDRX (CorelDRAW X6+)")
        with zipfile.ZipFile(filepath, 'r') as zf:
            names = zf.namelist()
            print(f"   Entries: {len(names)}")
            for n in names[:10]:
                info = zf.getinfo(n)
                print(f"     {n} ({info.file_size:,} bytes)")
            if len(names) > 10:
                print(f"     ... and {len(names)-10} more")
            # Try reading metadata
            if 'metadata.xml' in names or 'META-INF/metadata.xml' in names:
                meta_file = 'metadata.xml' if 'metadata.xml' in names else 'META-INF/metadata.xml'
                meta = zf.read(meta_file).decode('utf-8', errors='ignore')
                print(f"   Metadata: {meta[:200]}...")

    else:
        print(f"   Format: Unknown (header: {header[:8].hex()})")


# ── SVG → CorelDRAW via COM ──

def svg_to_cdr_com(svg_path, cdr_output=None):
    """Import SVG into CorelDRAW via COM automation (Windows only)."""
    try:
        import win32com.client
        from win32com.client import Dispatch
    except ImportError:
        print("❌ pywin32 required. Run: pip install pywin32")
        return False

    if not os.path.exists(svg_path):
        print(f"❌ File not found: {svg_path}")
        return False

    try:
        corel = Dispatch("CorelDRAW.Application")
        corel.Visible = True
        import time
        time.sleep(0.5)
    except Exception as e:
        print(f"❌ Cannot connect to CorelDRAW: {e}")
        return False

    try:
        doc = corel.ActiveDocument
        if doc is None:
            doc = corel.CreateDocument()

        # Import SVG
        svg_abs = os.path.abspath(svg_path)
        doc.Import(svg_abs)

        if cdr_output:
            doc.SaveAs(os.path.abspath(cdr_output))
            print(f"✅ Saved: {cdr_output}")

        print(f"✅ Imported: {svg_path} → CorelDRAW")
        return True

    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def batch_svg_to_cdr(pattern='*.svg', output_dir='cdr_output'):
    """Batch convert SVG files to CDR via COM."""
    import glob
    files = glob.glob(pattern)
    if not files:
        print(f"No files matching: {pattern}")
        return

    os.makedirs(output_dir, exist_ok=True)
    print(f"📦 Processing {len(files)} files...")

    for i, f in enumerate(files):
        name = os.path.splitext(os.path.basename(f))[0]
        out = f'{output_dir}/{name}.cdr'
        print(f"  [{i+1}/{len(files)}] {f}")
        svg_to_cdr_com(f, out)


# ── CorelDRAW VBA Macro Generator ──

def generate_cdr_macros():
    """Generate a VBA helper library for CorelDRAW."""
    macros = r"""
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
"""
    with open('cdr_helpers.bas', 'w') as f:
        f.write(macros)
    print("✅ VBA helper library saved: cdr_helpers.bas")
    print("   In CorelDRAW: Alt+F11 → Import → cdr_helpers.bas")


# ── SVG Optimizer for CorelDRAW ──

def optimize_svg_for_corel(svg_path, output_path=None):
    """Optimize SVG for best CorelDRAW import compatibility."""
    if output_path is None:
        output_path = svg_path.replace('.svg', '_cdr.svg')

    tree = ET.parse(svg_path)
    root = tree.getroot()

    # Ensure namespaces
    nsmap = {
        'xmlns': 'http://www.w3.org/2000/svg',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'xmlns:inkscape': 'http://www.inkscape.org/namespaces/inkscape',
        'xmlns:cdr': 'http://www.corel.com/coreldraw/2024'
    }
    for k, v in nsmap.items():
        if k not in root.attrib:
            root.set(k, v)

    # Add CorelDRAW-friendly metadata
    desc = ET.SubElement(root, 'desc')
    desc.text = 'Generated by CorelDRAWer-Skill for CorelDRAW import'

    # Convert text to paths? No — keep editable
    # Ensure all fonts are Arial/Helvetica
    for text_elem in root.iter('{http://www.w3.org/2000/svg}text'):
        ff = text_elem.get('font-family', '')
        if 'SimHei' in ff:
            text_elem.set('font-family', 'Arial, Helvetica, SimHei, sans-serif')

    # Add CDR layer names from cdr-layer data attributes
    for g in root.iter('{http://www.w3.org/2000/svg}g'):
        layer_name = g.get('data-cdr-layer')
        if layer_name:
            g.set('id', f'cdr-{layer_name}')

    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    print(f"✅ Optimized: {output_path}")
    return output_path


# ── CLI ──

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]

    if cmd == 'inspect' and len(sys.argv) > 2:
        inspect_cdr(sys.argv[2])

    elif cmd == 'svg2cdr' and len(sys.argv) > 2:
        out = sys.argv[3] if len(sys.argv) > 3 else None
        svg_to_cdr_com(sys.argv[2], out)

    elif cmd == 'batch':
        pattern = sys.argv[2] if len(sys.argv) > 2 else '*.svg'
        batch_svg_to_cdr(pattern)

    elif cmd == 'template':
        out = sys.argv[2] if len(sys.argv) > 2 else 'template.cdr'
        print(f"Creating CorelDRAW template at: {out}")
        print("(Use COM mode on Windows to create directly)")

    elif cmd == 'macro':
        generate_cdr_macros()

    elif cmd == 'optimize' and len(sys.argv) > 2:
        out = sys.argv[3] if len(sys.argv) > 3 else None
        optimize_svg_for_corel(sys.argv[2], out)

    else:
        print(f"Unknown command: {cmd}")
        print(__doc__)


if __name__ == '__main__':
    main()
