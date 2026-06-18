#!/usr/bin/env python3
"""
Nature Figure Extractor — download Nature PDF and extract figures as PNG/JPEG.
Usage: python3 extract_nature_figure.py <doi-or-url> [output_dir]
"""

import urllib.request
import ssl
import re
import os
import sys

def download_pdf(url):
    """Download a Nature article PDF."""
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    resp = urllib.request.urlopen(req, context=ctx, timeout=60)
    return resp.read()

def extract_images_from_pdf(pdf_bytes, output_dir='nature_figures'):
    """Extract all JPEG/PNG images from a PDF binary."""
    os.makedirs(output_dir, exist_ok=True)
    
    images = []
    # Find JPEG streams (DCTDecode) — these are directly extractable
    pos = 0
    img_idx = 0
    while True:
        img_start = pdf_bytes.find(b'/Subtype/Image', pos)
        if img_start == -1:
            break
        
        # Look around for metadata
        chunk = pdf_bytes[max(0, img_start-500):img_start+1000]
        w_match = re.search(rb'/Width\s+(\d+)', chunk)
        h_match = re.search(rb'/Height\s+(\d+)', chunk)
        len_match = re.search(rb'/Length\s+(\d+)', chunk)
        is_jpg = b'/DCTDecode' in chunk
        
        width = int(w_match.group(1)) if w_match else 0
        height = int(h_match.group(1)) if h_match else 0
        data_len = int(len_match.group(1)) if len_match else 0
        
        # Find stream
        stream_start = pdf_bytes.find(b'stream', img_start)
        if stream_start == -1:
            pos = img_start + 1
            continue
        # Skip past \r\n or \n after 'stream'
        data_start = stream_start + 6
        if pdf_bytes[data_start:data_start+1] == b'\n':
            data_start += 1
        elif pdf_bytes[data_start:data_start+2] == b'\r\n':
            data_start += 2
        
        if data_len > 0:
            data = pdf_bytes[data_start:data_start + data_len]
        else:
            endstream = pdf_bytes.find(b'endstream', data_start)
            data = pdf_bytes[data_start:endstream]
        
        if is_jpg and width >= 200 and height >= 200:
            img_idx += 1
            filename = f'{output_dir}/figure_{img_idx:02d}_{width}x{height}.jpg'
            with open(filename, 'wb') as f:
                f.write(data)
            images.append((filename, width, height))
            print(f'  ✅ {filename} ({width}×{height}, {len(data):,} bytes)')
        
        pos = img_start + 1
    
    return images

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 extract_nature_figure.py <doi-or-url> [output_dir]")
        print("Example: python3 extract_nature_figure.py 10.1038/s43247-026-03288-3")
        sys.exit(1)
    
    doi_or_url = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else 'nature_figures'
    
    # Build PDF URL
    if doi_or_url.startswith('http'):
        pdf_url = doi_or_url.replace('/figures/1', '.pdf').rstrip('?proof=t%2525C2%2525A0')
        if not pdf_url.endswith('.pdf'):
            pdf_url = pdf_url.rstrip('/') + '.pdf'
    else:
        pdf_url = f'https://www.nature.com/articles/{doi_or_url}.pdf'
    
    print(f'📥 Downloading: {pdf_url}')
    pdf = download_pdf(pdf_url)
    print(f'   {len(pdf):,} bytes downloaded')
    
    print(f'🔍 Extracting figures to: {output_dir}/')
    images = extract_images_from_pdf(pdf, output_dir)
    
    print(f'\n✅ Extracted {len(images)} figure-quality images')
    if images:
        print(f'   Largest: {max(images, key=lambda x: x[1]*x[2])[0]}')

if __name__ == '__main__':
    main()
