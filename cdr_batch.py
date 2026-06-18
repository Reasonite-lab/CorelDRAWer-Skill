#!/usr/bin/env python3
"""
CorelDRAWer Batch & CDR Toolkit
══════════════════════════════════════════════════
Usage:
  python3 cdr_batch.py from-json data.json          → SVG (default)
  python3 cdr_batch.py from-csv boreholes.csv       → JSON + SVG
  python3 cdr_batch.py to-vba data.json             → VBA macro  
  python3 cdr_batch.py to-com data.json              → COM automation (Windows)
  python3 cdr_batch.py multi-section sections.json  → multi-column SVG

Dependencies: none (pure Python stdlib)
══════════════════════════════════════════════════
"""

import json, csv, sys, os, io

# ── CSV → JSON converter ──
def csv_to_column_json(csv_path, title=None, location=None):
    """Convert a CSV with columns: formation,thick,descr,pattern,c,m,y,k to JSON."""
    layers = []
    with open(csv_path, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            layers.append({
                "erathem": row.get('erathem', ''),
                "system": row.get('system', ''),
                "series": row.get('series', ''),
                "formation": row.get('formation', 'Layer'),
                "symbol": row.get('symbol', ''),
                "thick": float(row.get('thick', 10)),
                "descr": row.get('descr', ''),
                "c": int(row.get('c', 0)), "m": int(row.get('m', 0)),
                "y": int(row.get('y', 0)), "k": int(row.get('k', 0)),
                "pattern": row.get('pattern', 'pure'),
                "grain": int(row.get('grain', 0)) if row.get('grain') else 0,
                "fossils": row.get('fossils', '').split(',') if row.get('fossils') else [],
                "structures": row.get('structures', '').split(',') if row.get('structures') else [],
                "contact": row.get('contact', ''),
                "age_ma": float(row.get('age_ma', 0)) if row.get('age_ma') else None,
            })
    return {"title": title or os.path.basename(csv_path).replace('.csv',''),
            "location": location or "", "layers": layers}

# ── Multi-section generator ──
def generate_multi_section(data_list, output_path, style='default'):
    """Generate a side-by-side multi-section SVG from multiple column data dicts."""
    from generate_column import generate_svg, DEFAULT_DATA
    import xml.etree.ElementTree as ET
    from xml.dom import minidom

    if not data_list:
        print("No data provided")
        return

    # Generate each column SVG
    svgs = []
    for i, data in enumerate(data_list):
        svg_str = generate_svg(data, style=style)
        svgs.append(ET.fromstring(svg_str))

    # Combine into one wide SVG
    total_w = sum(int(s.get('width','200').replace('mm','')) + 10 for s in svgs)
    max_h = max(int(s.get('height','333').replace('mm','')) for s in svgs)

    root = ET.Element('svg', {
        'xmlns': 'http://www.w3.org/2000/svg',
        'xmlns:xlink': 'http://www.w3.org/1999/xlink',
        'xmlns:inkscape': 'http://www.inkscape.org/namespaces/inkscape',
        'viewBox': f'0 0 {total_w} {max_h}',
        'width': f'{total_w}mm', 'height': f'{max_h}mm', 'version': '1.1'
    })

    # White background
    bg = ET.SubElement(root, 'rect', {'x':'0','y':'0','width':str(total_w),'height':str(max_h),
                                       'fill':'#fff','stroke':'none','stroke-width':'0'})

    x_offset = 0
    for i, svg in enumerate(svgs):
        w = int(svg.get('width','200').replace('mm',''))
        g = ET.SubElement(root, 'g', {'id': f'section_{i}', 'transform': f'translate({x_offset},0)'})
        for child in list(svg):
            g.append(child)
        x_offset += w + 10

    rough = ET.tostring(root, encoding='unicode')
    try:
        dom = minidom.parseString(rough)
        result = dom.toprettyxml(indent='  ', encoding='UTF-8').decode('utf-8')
    except:
        result = rough

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(result)
    print(f'✅ Multi-section saved: {output_path} ({len(data_list)} columns)')
    return result

# ── CDR layer-optimized SVG ──
def generate_cdr_optimized_svg(data, output_path):
    """Generate SVG specifically optimized for CorelDRAW import with named layers."""
    from generate_column import generate_svg
    svg = generate_svg(data, style='nature')
    # The nature style already uses proper layer groups + inkscape namespaces
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(svg)
    print(f'✅ CDR-optimized SVG saved: {output_path}')
    return svg

# ── Batch VBA generator ──
def batch_vba(data_list, output_path='batch_macro.bas'):
    """Generate VBA that creates multiple columns in one CorelDRAW document.
    data_list can be a list of file paths (str) or loaded data dicts."""
    from cdr_com_auto import generate_vba_code

    lines = ["' ============================================================",
             f"' CorelDRAWer Batch VBA — {len(data_list)} sections",
             "' ============================================================",
             "Option Explicit", ""]

    for i, item in enumerate(data_list):
        # Accept both file paths and loaded dicts
        if isinstance(item, str):
            with open(item, 'r') as f:
                data = json.load(f)
        else:
            data = item
        vba = generate_vba_code(data)
        # Rename the main sub
        vba = vba.replace('Public Sub DrawColumn()', f'Public Sub DrawSection{i+1}()')
        # Remove the leading comment block (already in batch header)
        vba = vba.split("Option Explicit\n", 1)[-1] if "Option Explicit" in vba else vba
        vba = vba.replace("Option Explicit\n", f"' --- Section {i+1}: {data.get('title','')} ---\n")
        lines.append(vba)
        lines.append("")

    # Add a master sub
    lines.append("' Master sub — draws all sections")
    lines.append("Public Sub DrawAll()")
    lines.append("    On Error GoTo ErrHandler")
    lines.append("    ActiveDocument.BeginCommandGroup \"Batch Stratigraphic Columns\"")
    lines.append("    ActiveDocument.Unit = cdrMillimeter")
    for i in range(len(data_list)):
        lines.append(f"    DrawSection{i+1}")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"Batch complete!\"")
    lines.append("    Exit Sub")
    lines.append("ErrHandler:")
    lines.append("    ActiveDocument.EndCommandGroup")
    lines.append("    MsgBox \"Error: \" & Err.Description")
    lines.append("End Sub")

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f'✅ Batch VBA saved: {output_path} ({len(data_list)} sections)')

# ── CLI ──
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]
    args = sys.argv[2:]

    if cmd == 'from-csv':
        csv_path = args[0] if args else None
        if not csv_path:
            print("Usage: cdr_batch.py from-csv data.csv [output.svg]")
            return
        output = args[1] if len(args) > 1 else csv_path.replace('.csv','.svg')
        data = csv_to_column_json(csv_path)
        from generate_column import generate_svg
        generate_svg(data, output)
        print(f'✅ Generated from CSV: {output}')

    elif cmd == 'to-vba':
        json_path = args[0] if args else None
        if not json_path:
            print("Usage: cdr_batch.py to-vba data.json [output.bas]")
            return
        with open(json_path, 'r') as f:
            data = json.load(f)
        output = args[1] if len(args) > 1 else 'macro.bas'
        from cdr_com_auto import generate_vba_code
        generate_vba_code(data, output)

    elif cmd == 'cdr-svg':
        json_path = args[0] if args else None
        if not json_path:
            print("Usage: cdr_batch.py cdr-svg data.json [output.svg]")
            return
        with open(json_path, 'r') as f:
            data = json.load(f)
        output = args[1] if len(args) > 1 else json_path.replace('.json','_cdr.svg')
        generate_cdr_optimized_svg(data, output)

    elif cmd == 'multi-section':
        if len(args) < 2:
            print("Usage: cdr_batch.py multi-section file1.json file2.json ... output.svg [style]")
            return
        *json_files, output = args
        style = 'default'
        if output in ('default', 'nature'):
            style = output
            output = json_files.pop()

        sections = []
        for jf in json_files:
            with open(jf, 'r') as f:
                sections.append(json.load(f))
        generate_multi_section(sections, output, style=style)

    elif cmd == 'batch-vba':
        if len(args) < 2:
            print("Usage: cdr_batch.py batch-vba file1.json file2.json ... output.bas")
            return
        *json_files, output = args
        sections = []
        for jf in json_files:
            with open(jf, 'r') as f:
                sections.append(json.load(f))
        batch_vba(sections, output)

    else:
        print(f'Unknown command: {cmd}')
        print(__doc__)

if __name__ == '__main__':
    main()
