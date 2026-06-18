#!/usr/bin/env python3
"""
Nature Figure Hunter — search Nature papers, download PDFs, extract figures.
Usage:
  python3 nature_figure_hunter.py search "stratigraphic column" --max 5
  python3 nature_figure_hunter.py fetch 10.1038/s43247-026-03288-3
  python3 nature_figure_hunter.py search "Ediacaran Cambrian boundary" --extract
"""

import urllib.request
import urllib.parse
import ssl
import re
import os
import sys
import json

# ── SSL workaround ──
CTX = ssl.create_default_context()
CTX.check_hostname = False
CTX.verify_mode = ssl.CERT_NONE
UA = 'Mozilla/5.0 (compatible; CorelDRAWer/0.2; research-bot)'

def api_get(url, timeout=30):
    """GET JSON from API."""
    req = urllib.request.Request(url, headers={'User-Agent': UA})
    resp = urllib.request.urlopen(req, context=CTX, timeout=timeout)
    return json.loads(resp.read())

def download_pdf(doi_or_url, output_dir='papers'):
    """Download a Nature article PDF."""
    os.makedirs(output_dir, exist_ok=True)

    if doi_or_url.startswith('http'):
        pdf_url = doi_or_url
        if not pdf_url.endswith('.pdf'):
            pdf_url = pdf_url.rstrip('/') + '.pdf'
        # Extract DOI from URL
        doi_match = re.search(r'10\.\d{4,}/[\w.-]+', doi_or_url)
        doi = doi_match.group(0) if doi_match else 'unknown'
    else:
        doi = doi_or_url
        pdf_url = f'https://www.nature.com/articles/{doi}.pdf'

    safe_name = re.sub(r'[^a-zA-Z0-9_-]', '_', doi)
    pdf_path = f'{output_dir}/{safe_name}.pdf'

    if os.path.exists(pdf_path):
        print(f'  📄 Already cached: {pdf_path}')
        with open(pdf_path, 'rb') as f:
            return f.read(), safe_name

    print(f'  📥 Downloading: {pdf_url}')
    req = urllib.request.Request(pdf_url, headers={'User-Agent': UA})
    resp = urllib.request.urlopen(req, context=CTX, timeout=60)
    pdf = resp.read()

    with open(pdf_path, 'wb') as f:
        f.write(pdf)
    print(f'     {len(pdf):,} bytes → {pdf_path}')
    return pdf, safe_name

def extract_images(pdf_bytes, prefix='fig', output_dir='figures'):
    """Extract JPEG images from PDF."""
    os.makedirs(output_dir, exist_ok=True)

    pos = 0
    count = 0
    result = []

    while True:
        img_start = pdf_bytes.find(b'/Subtype/Image', pos)
        if img_start == -1:
            break

        chunk = pdf_bytes[max(0, img_start-100):img_start+600]
        is_jpg = b'/DCTDecode' in chunk
        w_m = re.search(rb'/Width\s+(\d+)', chunk)
        h_m = re.search(rb'/Height\s+(\d+)', chunk)
        l_m = re.search(rb'/Length\s+(\d+)', chunk)
        width = int(w_m.group(1)) if w_m else 0
        height = int(h_m.group(1)) if h_m else 0
        data_len = int(l_m.group(1)) if l_m else 0

        stream_start = pdf_bytes.find(b'stream', img_start, img_start+600)
        if stream_start >= 0 and is_jpg and width >= 300:
            data_start = stream_start + 6
            if pdf_bytes[data_start:data_start+1] == b'\n':
                data_start += 1
            elif pdf_bytes[data_start:data_start+2] == b'\r\n':
                data_start += 2

            data = pdf_bytes[data_start:data_start+data_len]
            count += 1
            fname = f'{output_dir}/{prefix}_{count:02d}_{width}x{height}.jpg'
            with open(fname, 'wb') as f:
                f.write(data)
            result.append((fname, width, height, len(data)))
            print(f'  🖼️  {fname} ({width}×{height})')

        pos = stream_start + 10 if stream_start >= 0 else img_start + 1

    return result

def search_openalex(query, max_results=5):
    """Search OpenAlex for Nature-family papers."""
    encoded = urllib.parse.quote(query, safe='')
    url = (f'https://api.openalex.org/works?search={encoded}'
           f'&sort=cited_by_count:desc&per_page={max_results * 3}')  # fetch more, filter later

    print(f'🔍 Searching: "{query}"')
    data = api_get(url)

    results = []
    for work in data.get('results', []):
        doi = work.get('doi', '').replace('https://doi.org/', '')
        title = work.get('title', 'Unknown')
        year = work.get('publication_year', '?')
        cited = work.get('cited_by_count', 0)
        source = work.get('primary_location', {}).get('source', {}).get('display_name', '?')
        is_oa = work.get('open_access', {}).get('is_oa', False)

        if doi and doi.startswith('10.1038/'):
            results.append({
                'doi': doi, 'title': title, 'year': year,
                'cited': cited, 'source': source, 'is_oa': is_oa
            })
        if len(results) >= max_results:
            break

    return results

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]
    args = sys.argv[2:]

    # Parse flags
    max_results = 5
    do_extract = False
    query_parts = []
    for a in args:
        if a == '--max' or a == '-n':
            continue  # handled below
        elif a == '--extract':
            do_extract = True
        else:
            query_parts.append(a)

    # Handle --max N
    for i, a in enumerate(args):
        if a in ('--max', '-n') and i + 1 < len(args):
            max_results = int(args[i + 1])
            query_parts = [x for x in query_parts if x != args[i+1]]

    query = ' '.join(query_parts)

    if cmd == 'search':
        if not query:
            print("Usage: python3 nature_figure_hunter.py search \"your query\" [--max N] [--extract]")
            sys.exit(1)

        results = search_openalex(query, max_results)

        if not results:
            print('  No Nature-family papers found.')
            sys.exit(0)

        print(f'\n📚 Found {len(results)} papers:\n')
        for i, r in enumerate(results):
            oa_tag = '🔓OA' if r['is_oa'] else '🔒'
            print(f'  [{i+1}] {oa_tag} {r["title"][:100]}')
            print(f'      DOI: {r["doi"]}  |  {r["source"]} ({r["year"]})  |  cited {r["cited"]}×')
            print()

        if do_extract:
            print('─' * 60)
            for i, r in enumerate(results):
                if r['is_oa']:
                    print(f'\n📄 [{i+1}/{len(results)}] {r["doi"]}')
                    try:
                        pdf, name = download_pdf(r['doi'])
                        imgs = extract_images(pdf, prefix=name[:30])
                        print(f'     ✅ {len(imgs)} figures extracted')
                    except Exception as e:
                        print(f'     ❌ Failed: {e}')
                else:
                    print(f'\n📄 [{i+1}/{len(results)}] {r["doi"]} — skipped (not OA)')

    elif cmd == 'fetch':
        doi_or_url = query if query else sys.argv[2] if len(sys.argv) > 2 else None
        if not doi_or_url:
            print("Usage: python3 nature_figure_hunter.py fetch <doi-or-url>")
            sys.exit(1)
        pdf, name = download_pdf(doi_or_url)
        imgs = extract_images(pdf, prefix=name[:30])
        print(f'\n✅ {len(imgs)} figures extracted')

    else:
        print(f'Unknown command: {cmd}')
        print(__doc__)

if __name__ == '__main__':
    main()
