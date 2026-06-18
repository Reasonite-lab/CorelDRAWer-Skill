#!/usr/bin/env python3
"""
CDR Layer Enhancer — post-process SVG for optimal CorelDRAW import.
Adds: CDR-compatible layer names, object styles, guides, page setup.

Usage: python3 cdr_enhance.py input.svg output.svg
"""

import sys
import xml.etree.ElementTree as ET

def enhance_for_coreldraw(input_path, output_path):
    """Add CorelDRAW-specific metadata to an existing SVG."""
    ET.register_namespace('', 'http://www.w3.org/2000/svg')
    ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')
    ET.register_namespace('inkscape', 'http://www.inkscape.org/namespaces/inkscape')
    ET.register_namespace('cdr', 'http://www.corel.com/coreldraw/2025')

    tree = ET.parse(input_path)
    root = tree.getroot()

    # Add CorelDRAW namespace
    root.set('xmlns:cdr', 'http://www.corel.com/coreldraw/2025')

    # Add CDR page setup metadata
    page_w = root.get('width', '210mm')
    page_h = root.get('height', '297mm')

    metadata = ET.SubElement(root, 'cdr:metadata')
    ET.SubElement(metadata, 'cdr:page', {
        'width': page_w, 'height': page_h, 'units': 'mm'
    })
    ET.SubElement(metadata, 'cdr:generator', {'name': 'CorelDRAWer-Skill', 'version': '0.2.1'})

    # Enhance layer groups with CDR-specific attributes
    for g in root.iter('{http://www.w3.org/2000/svg}g'):
        gid = g.get('id', '')
        if gid.startswith('cdr-'):
            # Add CDR layer properties
            g.set('cdr:layer', gid.replace('cdr-', ''))
            g.set('cdr:editable', 'true')
            g.set('cdr:visible', 'true')
            g.set('cdr:printable', 'true')

            # Add proper layer locking for reference layers
            if gid in ('cdr-background', 'cdr-outlines', 'cdr-grid'):
                g.set('cdr:locked', 'true')

    # Add guides for alignment
    defs = root.find('{http://www.w3.org/2000/svg}defs')
    if defs is None:
        defs = ET.SubElement(root, 'defs')

    # Horizontal guide at table top
    ET.SubElement(defs, 'cdr:guide', {
        'orientation': 'horizontal', 'position': '290', 'label': 'Table Top'
    })
    ET.SubElement(defs, 'cdr:guide', {
        'orientation': 'horizontal', 'position': '22', 'label': 'Table Bottom'
    })

    # Write
    tree.write(output_path, encoding='UTF-8', xml_declaration=True)
    print(f'✅ CDR-enhanced SVG saved: {output_path}')

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: python3 cdr_enhance.py input.svg output.svg")
        sys.exit(1)
    enhance_for_coreldraw(sys.argv[1], sys.argv[2])
