#!/usr/bin/env python3
"""Comprehensive DOCX inspection: check formatting, structure, page breaks, etc."""
import zipfile
import xml.etree.ElementTree as ET
import sys

ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

def tag(el):
    return el.tag.split('}')[1] if '}' in el.tag else el.tag

def get_text(el):
    """Get all text from element recursively."""
    texts = []
    for t in el.iter():
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

wns = '{' + ns['w'] + '}'
rns_attr = '{' + ns['r'] + '}'

def check_para(p, idx):
    """Analyze a single paragraph."""
    info = {'idx': idx, 'text': '', 'runs': [], 'heading': 0, 'list': False,
            'alignment': '', 'page_break': False, 'section_break': False,
            'in_table': False, 'has_image': False}
    
    pPr = p.find('w:pPr', ns)
    if pPr is not None:
        pStyle = pPr.find('w:pStyle', ns)
        if pStyle is not None:
            val = pStyle.get(wns + 'val', '')
            if val.startswith('Heading'):
                try:
                    info['heading'] = int(val.replace('Heading', ''))
                except:
                    pass
        
        numPr = pPr.find('w:numPr', ns)
        if numPr is not None:
            info['list'] = True
        
        jc = pPr.find('w:jc', ns)
        if jc is not None:
            info['alignment'] = jc.get(wns + 'val', '')
        
        pbBefore = pPr.find('w:pageBreakBefore', ns)
        if pbBefore is not None:
            info['page_break'] = True
        
        sectPr = pPr.find('w:sectPr', ns)
        if sectPr is not None:
            info['section_break'] = True
    
    for r in p.findall('w:r', ns):
        br = r.find('w:br', ns)
        if br is not None and br.get(wns + 'type', '') == 'page':
            info['page_break'] = True
        
        drawing = r.find('w:drawing', ns)
        if drawing is not None:
            info['has_image'] = True
        
        t = r.find('w:t', ns)
        if t is not None and t.text:
            rPr = r.find('w:rPr', ns)
            run_info = {'text': t.text, 'bold': False, 'italic': False, 
                       'font': '', 'size': 0, 'color': '', 'underline': False}
            if rPr is not None:
                if rPr.find('w:b', ns) is not None:
                    run_info['bold'] = True
                if rPr.find('w:i', ns) is not None:
                    run_info['italic'] = True
                sz = rPr.find('w:sz', ns)
                if sz is not None:
                    try:
                        run_info['size'] = int(sz.get(wns + 'val', '0'))
                    except:
                        pass
                color = rPr.find('w:color', ns)
                if color is not None:
                    run_info['color'] = color.get(wns + 'val', '')
                rFonts = rPr.find('w:rFonts', ns)
                if rFonts is not None:
                    run_info['font'] = (rFonts.get(wns + 'eastAsia', '') or 
                                       rFonts.get(wns + 'ascii', ''))
                u = rPr.find('w:u', ns)
                if u is not None:
                    run_info['underline'] = True
            info['runs'].append(run_info)
            info['text'] += t.text
    
    return info

def main():
    path = sys.argv[1] if len(sys.argv) > 1 else 'testfie/test.docx'
    z = zipfile.ZipFile(path)
    
    print(f'=== DOCX Structure: {path} ===')
    print('Files in archive:')
    for name in sorted(z.namelist()):
        print(f'  {name} ({z.getinfo(name).file_size} bytes)')
    
    doc = z.read('word/document.xml').decode('utf-8')
    root = ET.fromstring(doc)
    body = root.find('w:body', ns)
    
    print(f'\n=== Body Children ===')
    children = list(body)
    child_tags = {}
    for c in children:
        t = tag(c)
        child_tags[t] = child_tags.get(t, 0) + 1
    for t, cnt in sorted(child_tags.items()):
        print(f'  {t}: {cnt}')
    
    # Analyze paragraphs
    print(f'\n=== Paragraph Analysis ===')
    paras = []
    para_idx = 0
    page_breaks = []
    headings = []
    lists = []
    images = []
    tables = body.findall('w:tbl', ns)
    
    for child in children:
        t = tag(child)
        if t == 'p':
            info = check_para(child, para_idx)
            paras.append(info)
            if info['page_break']:
                page_breaks.append(para_idx)
            if info['heading'] > 0:
                headings.append((para_idx, info['heading'], info['text'][:60]))
            if info['list']:
                lists.append(para_idx)
            if info['has_image']:
                images.append(para_idx)
            para_idx += 1
        elif t == 'tbl':
            rows = child.findall('w:tr', ns)
            for row in rows:
                cells = row.findall('w:tc', ns)
                for cell in cells:
                    for cp in cell.findall('w:p', ns):
                        info = check_para(cp, para_idx)
                        info['in_table'] = True
                        paras.append(info)
                        para_idx += 1
    
    print(f'Total paragraphs: {len(paras)}')
    print(f'Page breaks: {len(page_breaks)} at indices {page_breaks}')
    print(f'Headings: {len(headings)}')
    for idx, level, text in headings:
        print(f'  H{level} @{idx}: {text}')
    print(f'List items: {len(lists)}')
    print(f'Images: {len(images)} at indices {images}')
    print(f'Tables: {len(tables)}')
    
    for tbl in tables:
        rows = tbl.findall('w:tr', ns)
        print(f'  Table: {len(rows)} rows')
        for i, row in enumerate(rows[:5]):
            cells = row.findall('w:tc', ns)
            cell_texts = []
            for cell in cells:
                ct = get_text(cell)[:30]
                cell_texts.append(ct)
            print(f'    Row {i}: {len(cells)} cells: {cell_texts}')
    
    # Show first 30 paragraphs with formatting
    print(f'\n=== First 30 Paragraphs Detail ===')
    for info in paras[:30]:
        text = info['text'][:80]
        flags = []
        if info['heading']:
            flags.append(f'H{info["heading"]}')
        if info['list']:
            flags.append('LIST')
        if info['page_break']:
            flags.append('PB')
        if info['section_break']:
            flags.append('SB')
        if info['has_image']:
            flags.append('IMG')
        if info['alignment']:
            flags.append(f'align={info["alignment"]}')
        if info['in_table']:
            flags.append('TABLE')
        flag_str = ' '.join(flags)
        print(f'  P{info["idx"]}: [{flag_str}] {repr(text)}')
        for r in info['runs'][:3]:
            rf = []
            if r['bold']: rf.append('B')
            if r['italic']: rf.append('I')
            if r['underline']: rf.append('U')
            if r['size']: rf.append(f's={r["size"]}')
            if r['font']: rf.append(f'f={r["font"]}')
            if r['color']: rf.append(f'c={r["color"]}')
            print(f'    R: {" ".join(rf)} {repr(r["text"][:50])}')
    
    # Check styles.xml
    print('\n=== Styles ===')
    try:
        styles_xml = z.read('word/styles.xml').decode('utf-8')
        styles_root = ET.fromstring(styles_xml)
        for style in styles_root.findall('w:style', ns):
            sid = style.get(wns + 'styleId', '')
            stype = style.get(wns + 'type', '')
            name_el = style.find('w:name', ns)
            sname = name_el.get(wns + 'val', '') if name_el is not None else ''
            print(f'  {sid} ({stype}): {sname}')
    except:
        print('  Could not read styles.xml')
    
    # Check headers/footers
    print(f'\n=== Headers/Footers ===')
    for name in z.namelist():
        if 'header' in name or 'footer' in name:
            content = z.read(name).decode('utf-8')
            hf_root = ET.fromstring(content)
            text = get_text(hf_root)
            print(f'  {name}: {repr(text[:100])}')
    
    # Check numbering.xml
    print(f'\n=== Numbering ===')
    try:
        num_xml = z.read('word/numbering.xml').decode('utf-8')
        num_root = ET.fromstring(num_xml)
        abstracts = num_root.findall('w:abstractNum', ns)
        nums = num_root.findall('w:num', ns)
        print(f'  abstractNum: {len(abstracts)}, num: {len(nums)}')
    except:
        print('  No numbering.xml')
    
    # Check sectPr (final section properties)
    print('\n=== Final Section Properties ===')
    sectPr = body.find('w:sectPr', ns)
    wns = '{' + ns['w'] + '}'
    rns = '{' + ns['r'] + '}'
    if sectPr is not None:
        pgSz = sectPr.find('w:pgSz', ns)
        if pgSz is not None:
            print(f'  Page size: w={pgSz.get(wns+"w")}, h={pgSz.get(wns+"h")}')
        pgMar = sectPr.find('w:pgMar', ns)
        if pgMar is not None:
            print(f'  Margins: top={pgMar.get(wns+"top")}, bottom={pgMar.get(wns+"bottom")}, left={pgMar.get(wns+"left")}, right={pgMar.get(wns+"right")}')
        for hdr in sectPr.findall('w:headerReference', ns):
            print(f'  Header ref: {hdr.get(rns+"id")}')
        for ftr in sectPr.findall('w:footerReference', ns):
            print(f'  Footer ref: {ftr.get(rns+"id")}')
    
    # Check rels
    print(f'\n=== Document Relationships ===')
    try:
        rels_xml = z.read('word/_rels/document.xml.rels').decode('utf-8')
        rels_root = ET.fromstring(rels_xml)
        for rel in rels_root:
            rid = rel.get('Id', '')
            rtype = rel.get('Type', '').split('/')[-1]
            target = rel.get('Target', '')
            print(f'  {rid}: {rtype} -> {target}')
    except:
        print('  Could not read rels')

if __name__ == '__main__':
    main()
