#!/usr/bin/env python3
"""
CorelDRAW COM 自动化 —— 直接驱动 CorelDRAW 绘制地层柱状图
══════════════════════════════════════════════════════════════
要求：Windows + CorelDRAW X4+ + Python + pywin32
安装：pip install pywin32
用法：
  python3 cdr_com_auto.py data.json        # 在当前文档中绘制
  python3 cdr_com_auto.py --new data.json  # 新建文档绘制
  python3 cdr_com_auto.py                  # 使用内置示例数据
══════════════════════════════════════════════════════════════
"""

import json
import sys
import os
import time

# ============================================================================
# TRY TO IMPORT WIN32COM
# ============================================================================
try:
    import win32com.client
    from win32com.client import Dispatch, constants
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ============================================================================
# FALLBACK — VBA CODE GENERATOR (for manual copy-paste when not on Windows)
# ============================================================================

def generate_vba_code(data, output_path=None):
    """
    Generate a complete VBA macro that draws the column in CorelDRAW.
    This is the same as the previous borehole_column.bas but integrated here.
    Fallback when COM automation is not available.
    """
    layers = data['layers']
    title = data.get('title', '综合地层柱状图')
    
    total_thick = sum(l['thick'] for l in layers)
    
    # Layout (mm)
    lines = []
    lines.append("' ============================================================")
    lines.append(f"' CorelDRAW 地层柱状图 VBA  —  {title}")
    lines.append(f"' 总厚度: {total_thick}m  ·  {len(layers)} 层")
    lines.append("' ============================================================")
    lines.append("")
    lines.append("Public Sub DrawColumn()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    If ActiveDocument Is Nothing Then")
    lines.append("        MsgBox \"请先打开文档\"")
    lines.append("        Exit Sub")
    lines.append("    End If")
    lines.append("")
    lines.append("    ActiveDocument.BeginCommandGroup \"绘制地层柱状图\"")
    lines.append("    ActiveDocument.Unit = cdrMillimeter")
    lines.append("    ActiveDocument.ReferencePoint = cdrBottomLeft")
    lines.append("")
    
    # Column layout constants
    col_x = [0, 8, 22, 36, 50, 70, 83, 118, 130]
    col_w = [0, 14, 14, 14, 20, 13, 35, 12, 65]
    lines.append("    ' 列坐标")
    for ci in range(1, 9):
        lines.append(f"    Const CX{ci} = {col_x[ci]}")
        lines.append(f"    Const CW{ci} = {col_w[ci]}")
    
    lines.append("    Const TABLE_TOP = 290")
    lines.append("    Const TABLE_BOT = 22")
    lines.append("    Const HEADER_H = 14")
    lines.append("")
    
    draw_h = 290 - 22 - 14
    scale_f = draw_h / total_thick if total_thick > 0 else 1
    
    lines.append(f"    Dim scaleF As Double: scaleF = {scale_f}")
    lines.append(f"    Dim totalThick As Double: totalThick = {total_thick}")
    lines.append("")
    
    # Title
    lines.append("    ' === 标题 ===")
    lines.append(f"    Dim t As Shape")
    lines.append(f"    Set t = ActiveLayer.CreateArtisticText(90, TABLE_TOP + 8, \"{title}\")")
    lines.append("    With t.Text.Story: .Font = \"黑体\": .Size = 12: .Bold = True: End With")
    lines.append("")
    
    # Table header and frame
    lines.append("    ' === 表头 ===")
    headers = [("界",1), ("系",2), ("统",3), ("组",4), ("代号",5), ("柱状图",6), ("厚度(m)",7), ("岩性描述",8)]
    for htxt, hcol in headers:
        cx_val = col_x[hcol] + col_w[hcol] // 2
        lines.append(f"    AddHeaderText ActiveLayer, CX{hcol} + CW{hcol}/2, TABLE_TOP - HEADER_H/2, \"{htxt}\"")
    
    # Layer loop
    lines.append("")
    lines.append(f"    ' === 逐层绘制 ===")
    lines.append(f"    Dim currentY As Double: currentY = TABLE_TOP - HEADER_H")
    lines.append(f"    Dim data({len(layers)-1}, 9) As Variant")
    
    for i, lay in enumerate(layers):
        lines.append(f"    ' Layer {i+1}: {lay['formation']}")
        lines.append(f"    data({i}, 0) = \"{lay.get('erathem', '')}\"")
        lines.append(f"    data({i}, 1) = \"{lay.get('system', '')}\"")
        lines.append(f"    data({i}, 2) = \"{lay.get('series', '')}\"")
        lines.append(f"    data({i}, 3) = \"{lay.get('formation', '')}\"")
        lines.append(f"    data({i}, 4) = \"{lay.get('symbol', '')}\"")
        lines.append(f"    data({i}, 5) = {lay['thick']}")
        lines.append(f"    data({i}, 6) = \"{lay.get('descr', '')}\"")
        lines.append(f"    data({i}, 7) = \"{lay.get('pattern', 'pure')}\"")
        lines.append(f"    data({i}, 8) = {lay['c']}: data({i}, 9) = {lay['m']}")
        lines.append(f"    data({i}, 10) = {lay['y']}: data({i}, 11) = {lay['k']}")
    
    lines.append("")
    lines.append("    Dim i As Long")
    lines.append("    For i = 0 To UBound(data, 1)")
    lines.append("        Dim layH As Double: layH = data(i, 5) * scaleF")
    lines.append("        If layH < 3.5 Then layH = 3.5")
    lines.append("        Dim layB As Double: layB = currentY - layH")
    lines.append("")
    lines.append("        ' 柱状图矩形")
    lines.append("        Dim rect As Shape")
    lines.append("        Set rect = ActiveLayer.CreateRectangle(CX6 + 0.3, layB, CX6 + CW6 - 2.5, currentY)")
    lines.append("        rect.Fill.UniformColor.CMYKAssign data(i, 8), data(i, 9), data(i, 10), data(i, 11)")
    lines.append("        rect.Outline.Color.CMYKAssign 0, 0, 0, 35")
    lines.append("        rect.Outline.Width = 0.15")
    lines.append("")
    lines.append("        ' 文字列")
    lines.append("        AddCellText ActiveLayer, CX1, currentY, layB, CW1, data(i, 0)")
    lines.append("        AddCellText ActiveLayer, CX2, currentY, layB, CW2, data(i, 1)")
    lines.append("        AddCellText ActiveLayer, CX3, currentY, layB, CW3, data(i, 2)")
    lines.append("        AddCellText ActiveLayer, CX4, currentY, layB, CW4, data(i, 3)")
    lines.append("        AddCellText ActiveLayer, CX5, currentY, layB, CW5, data(i, 4)")
    lines.append("        AddCellText ActiveLayer, CX7, currentY, layB, CW7, Format(data(i, 5), \"0.0\")")
    lines.append("        AddCellText ActiveLayer, CX8, currentY, layB, CW8, data(i, 6)")
    lines.append("")
    lines.append("        currentY = layB")
    lines.append("    Next i")
    lines.append("")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"绘制完成！\"")
    lines.append("    Exit Sub")
    lines.append("ErrHandler:")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"错误: \" & Err.Description")
    lines.append("End Sub")
    
    vba_code = '\n'.join(lines)
    
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(vba_code)
    
    return vba_code


# ============================================================================
# CORELDRAW COM AUTOMATION (Windows only)
# ============================================================================

def com_draw_column(data):
    """
    Connect to CorelDRAW via COM and draw the stratigraphic column.
    Runs on Windows with CorelDRAW installed.
    """
    if not HAS_WIN32:
        print("❌ pywin32 not installed. On Windows, run: pip install pywin32")
        print("   Generating VBA code as fallback...")
        return generate_vba_code(data, 'column_fallback.bas')
    
    try:
        corel = Dispatch("CorelDRAW.Application")
        corel.Visible = True
        time.sleep(0.5)
    except Exception as e:
        print(f"❌ Cannot connect to CorelDRAW: {e}")
        print("   Make sure CorelDRAW is installed and running.")
        return generate_vba_code(data, 'column_fallback.bas')
    
    layers_data = data['layers']
    title = data.get('title', '综合地层柱状图')
    total_thick = sum(l['thick'] for l in layers_data)
    
    print(f"✅ Connected to CorelDRAW")
    print(f"   Drawing: {title}")
    print(f"   {len(layers_data)} layers, {total_thick}m total")
    
    try:
        doc = corel.ActiveDocument
    except:
        # Create new document
        try:
            from win32com.client import Dispatch
            createOpt = Dispatch("CorelDRAW.StructCreateOptions")
            createOpt.Units = 6  # cdrMillimeter
            createOpt.PageWidth = 210
            createOpt.PageHeight = 297
            doc = corel.CreateDocumentEx(createOpt)
        except:
            doc = corel.CreateDocument()
    
    doc.Unit = 6  # cdrMillimeter
    doc.ReferencePoint = 7  # cdrBottomLeft
    
    # Begin undo group
    doc.BeginCommandGroup("地层柱状图")
    
    try:
        activeLayer = doc.ActivePage.ActiveLayer
        
        # Layout constants
        col_x = [0, 8, 22, 36, 50, 70, 83, 118, 130]
        col_w = [0, 14, 14, 14, 20, 13, 35, 12, 65]
        TABLE_TOP = 290
        TABLE_BOT = 22
        HEADER_H = 14
        draw_h = TABLE_TOP - TABLE_BOT - HEADER_H
        scale_f = draw_h / total_thick
        
        # Title
        title_shape = activeLayer.CreateArtisticText(90, TABLE_TOP + 8, title)
        title_shape.Text.Story.Font = "黑体"
        title_shape.Text.Story.Size = 12
        title_shape.Text.Story.Bold = True
        
        # Header
        header_labels = [("界",1), ("系",2), ("统",3), ("组",4), 
                        ("代号",5), ("柱状图",6), ("厚度(m)",7), ("岩性描述",8)]
        for htxt, hcol in header_labels:
            s = activeLayer.CreateArtisticText(
                col_x[hcol] + col_w[hcol]//2, TABLE_TOP - HEADER_H/2, htxt)
            s.Text.Story.Font = "黑体"
            s.Text.Story.Size = 7
            s.Text.Story.Bold = True
        
        # Draw each layer
        currentY = TABLE_TOP - HEADER_H
        colLeft = col_x[6] + 0.3
        colRight = col_x[6] + col_w[6] - 2.5
        
        for i, lay in enumerate(layers_data):
            layH = max(lay['thick'] * scale_f, 3.5)
            layB = currentY - layH
            
            c, m, y, k = lay['c'], lay['m'], lay['y'], lay['k']
            
            # Column rectangle
            rect = activeLayer.CreateRectangle(colLeft, layB, colRight, currentY)
            rect.Fill.UniformColor.CMYKAssign(c, m, y, k)
            rect.Outline.Color.CMYKAssign(0, 0, 0, 35)
            rect.Outline.Width = 0.15
            
            # Text columns
            txt_data = [
                (lay.get('erathem',''), 1), (lay.get('system',''), 2),
                (lay.get('series',''), 3), (lay.get('formation',''), 4),
                (lay.get('symbol',''), 5), (f"{lay['thick']:.1f}", 7),
                (lay.get('descr',''), 8)
            ]
            for txt, tcol in txt_data:
                if txt:
                    s = activeLayer.CreateArtisticText(
                        col_x[tcol] + 1, (currentY + layB) / 2, txt)
                    s.Text.Story.Font = "黑体"
                    s.Text.Story.Size = 5.5
                    s.Fill.UniformColor.CMYKAssign(0, 0, 0, 100)
            
            currentY = layB
        
        doc.EndCommandGroup()
        print(f"✅ Drawing completed in CorelDRAW!")
        return True
        
    except Exception as e:
        doc.EndCommandGroup()
        print(f"❌ Error during drawing: {e}")
        raise


# ============================================================================
# MAIN
# ============================================================================

def main():
    data = {
        "title": "秭归地区综合地层柱状图",
        "location": "湖北秭归",
        "layers": [
            {"erathem":"新元古界","system":"南华系","series":"下统","formation":"莲沱组","symbol":"Nh₁l","thick":120,"descr":"紫红色中厚层砂岩","c":0,"m":40,"y":30,"k":10,"pattern":"sand"},
            {"erathem":"新元古界","system":"南华系","series":"下统","formation":"南沱组","symbol":"Nh₁n","thick":45,"descr":"灰绿色冰碛砾岩","c":15,"m":0,"y":20,"k":20,"pattern":"conglo"},
            {"erathem":"新元古界","system":"震旦系","series":"下统","formation":"陡山沱组","symbol":"Z₁d","thick":180,"descr":"灰黑色泥质白云岩","c":5,"m":0,"y":10,"k":30,"pattern":"dolo"},
            {"erathem":"新元古界","system":"震旦系","series":"上统","formation":"灯影组","symbol":"Z₂dy","thick":350,"descr":"灰白色厚层白云岩","c":0,"m":0,"y":5,"k":10,"pattern":"dolo"},
            {"erathem":"古生界","system":"寒武系","series":"下统","formation":"水井沱组","symbol":"∈₁s","thick":95,"descr":"黑色炭质页岩","c":0,"m":5,"y":10,"k":65,"pattern":"carbShale"},
        ]
    }
    
    args = sys.argv[1:]
    mode = 'auto'  # auto, com, vba, svg
    
    for a in args:
        if a == '--com': mode = 'com'
        elif a == '--vba': mode = 'vba'
        elif a == '--svg': mode = 'svg'
        elif a == '--new': pass  # handled below
        elif a.endswith('.json'):
            with open(a, 'r', encoding='utf-8') as f:
                data = json.load(f)
    
    print(f"📊 地层柱状图生成器")
    print(f"   数据: {len(data['layers'])} 层, 总厚 {sum(l['thick'] for l in data['layers']):.0f}m")
    print()
    
    if mode == 'svg':
        # Use the SVG generator
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from generate_column import generate_svg
        output = data.get('title', 'column').replace(' ', '_') + '.svg'
        generate_svg(data, output)
        print(f"✅ SVG saved: {output}")
        return
    
    if mode == 'com' or mode == 'auto':
        if HAS_WIN32:
            result = com_draw_column(data)
            if result:
                return
        else:
            print("⚠️  Not on Windows / pywin32 not available.")
            print("   Switching to SVG mode...")
            print()
    
    if mode == 'vba' or mode == 'auto':
        vba = generate_vba_code(data, 'column_macro.bas')
        print(f"✅ VBA code saved: column_macro.bas")
        print(f"   Open CorelDRAW → Alt+F11 → Import → Run DrawColumn")
        return


if __name__ == '__main__':
    main()
