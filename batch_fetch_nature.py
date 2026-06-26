#!/usr/bin/env python3
"""
Batch fetch Nature papers: download PDFs via curl, extract figures.
Usage:  python3 batch_fetch_nature.py < doi_list.txt
"""
import subprocess, sys, os, re, zlib, time, json

PDF_DIR = "papers"
FIG_DIR = "nature_figures"
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"

os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(FIG_DIR, exist_ok=True)

def doi_to_slug(doi):
    suffix = doi.replace("10.1038/", "")
    return re.sub(r'[^a-zA-Z0-9_-]', '_', suffix)

def download_pdf(doi):
    slug = doi_to_slug(doi)
    pdf_path = f"{PDF_DIR}/{slug}.pdf"
    if os.path.exists(pdf_path):
        print(f"  📄 Already cached: {pdf_path}")
        return pdf_path

    url = f"https://www.nature.com/articles/{doi.replace('10.1038/', '')}.pdf"
    print(f"  📥 Downloading: {doi}")
    result = subprocess.run(
        ["curl", "-L", "-s", "--max-time", "90", "-o", pdf_path,
         "-H", f"User-Agent: {UA}",
         "-H", "Referer: https://www.nature.com/",
         "-w", "%{http_code}",
         url],
        capture_output=True, text=True, timeout=120
    )
    http_code = result.stdout.strip()
    size = os.path.getsize(pdf_path) if os.path.exists(pdf_path) else 0
    if http_code == "200" and size > 10000:
        print(f"     ✅ {size:,} bytes → {pdf_path}")
        return pdf_path
    else:
        print(f"     ❌ HTTP {http_code}, size={size}")
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        return None

def extract_images(pdf_bytes, prefix="fig"):
    """Extract JPEG and PNG images from PDF bytes."""
    count = 0
    results = []
    pos = 0

    while True:
        img_start = pdf_bytes.find(b'/Subtype/Image', pos)
        if img_start == -1:
            break

        chunk = pdf_bytes[max(0, img_start-100):img_start+800]
        is_jpg = b'/DCTDecode' in chunk

        w_m = re.search(rb'/Width\s+(\d+)', chunk)
        h_m = re.search(rb'/Height\s+(\d+)', chunk)
        l_m = re.search(rb'/Length\s+(\d+)', chunk)
        width = int(w_m.group(1)) if w_m else 0
        height = int(h_m.group(1)) if h_m else 0
        data_len = int(l_m.group(1)) if l_m else 0

        stream_start = pdf_bytes.find(b'stream', img_start, img_start+800)
        if stream_start >= 0 and width >= 300:
            data_start = stream_start + 6
            if pdf_bytes[data_start:data_start+1] == b'\n':
                data_start += 1
            elif pdf_bytes[data_start:data_start+2] == b'\r\n':
                data_start += 2

            data = pdf_bytes[data_start:data_start+data_len]

            if is_jpg:
                count += 1
                fname = f'{FIG_DIR}/{prefix}_{count:02d}_{width}x{height}.jpg'
                with open(fname, 'wb') as f:
                    f.write(data)
                results.append((fname, width, height, len(data)))
                print(f"     🖼️  {os.path.basename(fname)} ({width}×{height})")

            elif width >= 400:
                # FlateDecode (PNG raw)
                try:
                    raw = zlib.decompress(data)
                    count += 1
                    fname = f'{FIG_DIR}/{prefix}_{count:02d}_{width}x{height}.raw'
                    with open(fname, 'wb') as f:
                        f.write(raw)
                    results.append((fname, width, height, len(raw)))
                    print(f"     📦 {os.path.basename(fname)} ({width}x{height}, FlateDecode)")
                except:
                    pass

        pos = stream_start + 10 if stream_start >= 0 else img_start + 1

    return results

def process_doi(doi):
    print(f"\n{'='*60}")
    print(f"📄 Processing: {doi}")
    print(f"{'='*60}")

    pdf_path = download_pdf(doi)
    if not pdf_path:
        return 0

    with open(pdf_path, 'rb') as f:
        pdf_bytes = f.read()

    slug = doi_to_slug(doi)
    imgs = extract_images(pdf_bytes, prefix=slug[:30])
    print(f"     ✅ {len(imgs)} figures extracted")
    return len(imgs)

def main():
    if len(sys.argv) > 1:
        # Read DOIs from file
        with open(sys.argv[1]) as f:
            dois = [line.strip() for line in f if line.strip() and not line.startswith('#')]
    else:
        # Read from stdin
        dois = [line.strip() for line in sys.stdin if line.strip() and not line.startswith('#')]

    total_figures = 0
    for i, doi in enumerate(dois):
        print(f"\n📑 [{i+1}/{len(dois)}]")
        n = process_doi(doi)
        total_figures += n
        if i < len(dois) - 1:
            delay = 5
            print(f"     ⏳ Waiting {delay}s before next...")
            time.sleep(delay)

    print(f"\n{'='*60}")
    print(f"🎉 Done! Processed {len(dois)} papers, extracted {total_figures} figures total.")
    print(f"{'='*60}")

if __name__ == '__main__':
    main()
