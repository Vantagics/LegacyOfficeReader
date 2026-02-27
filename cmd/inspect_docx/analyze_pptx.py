#!/usr/bin/env python3
"""Analyze PPTX output quality."""
import zipfile
import xml.etree.ElementTree as ET
import sys
from collections import Counter

ns = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}

def analyze(path):
    z = zipfile.ZipFile(path)
    slides = 0
    images = 0
    sp = 0; pic = 0; cxn = 0
    geoms = Counter()
    fonts = Counter()
    runs_total = 0; runs_no_fmt = 0
    runs_font = 0; runs_size = 0; runs_bold = 0; runs_color = 0
    paras_total = 0; paras_align = 0; paras_bullet = 0; paras_margin = 0
    paras_spacing = 0; paras_indent = 0
    multi_run_paras = 0
    solid_fill = 0; no_fill = 0
    rotated = 0; line_width = 0
    neg_dims = 0; zero_dims = 0

    for f in z.namelist():
        if f.startswith('ppt/slides/slide') and f.endswith('.xml'):
            slides += 1
            root = ET.parse(z.open(f)).getroot()
            for sp_el in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}sp'):
                sp += 1
                analyze_shape(sp_el, geoms, fonts, 
                    [runs_total, runs_no_fmt, runs_font, runs_size, runs_bold, runs_color],
                    [paras_total, paras_align, paras_bullet, paras_margin, paras_spacing, paras_indent, multi_run_paras])
            for pic_el in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}pic'):
                pic += 1
            for cxn_el in root.iter('{http://schemas.openxmlformats.org/presentationml/2006/main}cxnSp'):
                cxn += 1
            
            content = z.read(f).decode('utf-8', errors='replace')
            solid_fill += content.count('<a:solidFill>')
            no_fill += content.count('<a:noFill/>')
            rotated += content.count(' rot="')
            line_width += content.count('<a:ln w="')
            
            # Count runs and formatting
            for txBody in root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}p'):
                paras_total += 1
                pPr = txBody.find('a:pPr', ns)
                if pPr is not None:
                    if pPr.get('algn'): paras_align += 1
                    if pPr.find('a:buChar', ns) is not None: paras_bullet += 1
                    if pPr.get('marL'): paras_margin += 1
                    if pPr.get('indent'): paras_indent += 1
                    if pPr.find('a:lnSpc', ns) is not None or pPr.find('a:spcBef', ns) is not None:
                        paras_spacing += 1
                
                run_els = txBody.findall('a:r', ns)
                if len(run_els) > 1:
                    multi_run_paras += 1
                for r in run_els:
                    runs_total += 1
                    rPr = r.find('a:rPr', ns)
                    if rPr is None:
                        runs_no_fmt += 1
                        continue
                    has_any = False
                    if rPr.find('a:latin', ns) is not None:
                        tf = rPr.find('a:latin', ns).get('typeface', '')
                        if tf:
                            runs_font += 1
                            fonts[tf] += 1
                            has_any = True
                    if rPr.get('sz') and rPr.get('sz') != '0':
                        runs_size += 1
                        has_any = True
                    if rPr.get('b') == '1':
                        runs_bold += 1
                        has_any = True
                    if rPr.find('a:solidFill', ns) is not None:
                        runs_color += 1
                        has_any = True
                    if not has_any:
                        runs_no_fmt += 1
                    
        if f.startswith('ppt/media/'):
            images += 1

    total = sp + pic + cxn
    print(f"=== PPTX Analysis ===")
    print(f"Slides: {slides}")
    print(f"Images: {images}")
    print(f"Total shapes: {total} (sp={sp}, pic={pic}, cxnSp={cxn})")
    print(f"\n--- Shape Properties ---")
    print(f"SolidFill: {solid_fill}")
    print(f"NoFill: {no_fill}")
    print(f"Rotated: {rotated}")
    print(f"Line width set: {line_width}")
    print(f"\n--- Text Formatting ---")
    print(f"Total paragraphs: {paras_total}")
    print(f"  With alignment: {paras_align}")
    print(f"  With bullet: {paras_bullet}")
    print(f"  With margin: {paras_margin}")
    print(f"  With indent: {paras_indent}")
    print(f"  With spacing: {paras_spacing}")
    print(f"  Multi-run paragraphs: {multi_run_paras}")
    print(f"Total text runs: {runs_total}")
    print(f"  With font: {runs_font}")
    print(f"  With fontSize: {runs_size}")
    print(f"  With bold: {runs_bold}")
    print(f"  With color: {runs_color}")
    pct = runs_no_fmt * 100 / max(runs_total, 1)
    print(f"  No formatting: {runs_no_fmt} ({pct:.1f}%)")
    print(f"\n--- Fonts Used ---")
    for font, count in fonts.most_common():
        print(f"  {font}: {count}")

def analyze_shape(el, geoms, fonts, run_stats, para_stats):
    pass  # handled in main loop

if __name__ == '__main__':
    path = sys.argv[1] if len(sys.argv) > 1 else 'testfie/test.pptx'
    analyze(path)
