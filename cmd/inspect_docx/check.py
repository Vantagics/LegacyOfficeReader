import zipfile
import xml.etree.ElementTree as ET

z = zipfile.ZipFile('testfie/test.docx')
doc = z.read('word/document.xml').decode('utf-8')

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
root = ET.fromstring(doc)
body = root.find('w:body', ns)

# Check body direct children
print('Body direct children tags:')
tags = {}
for child in body:
    tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
    tags[tag] = tags.get(tag, 0) + 1
for tag, count in sorted(tags.items()):
    print(f'  {tag}: {count}')

# Check for nested <w:p> inside <w:p> (but not inside <w:tc>)
print('\nChecking for invalid nested <w:p>...')
found_nested = False
for p in body.findall('w:p', ns):
    # Direct child paragraphs of body should not contain nested <w:p>
    nested = p.findall('.//w:p', ns)
    if nested:
        print(f'  NESTED <w:p> found!')
        found_nested = True
        break

if not found_nested:
    print('  No invalid nested <w:p> found')

# Check table structure
print('\nTable structure:')
for tbl in body.findall('w:tbl', ns):
    rows = tbl.findall('w:tr', ns)
    print(f'  Table with {len(rows)} rows')
    for i, row in enumerate(rows):
        cells = row.findall('w:tc', ns)
        print(f'    Row {i}: {len(cells)} cells')
        if i > 5:
            print(f'    ... ({len(rows) - i - 1} more rows)')
            break

# Check for any element with text containing control chars
print('\nChecking all text content for control chars...')
problem_count = 0
for elem in root.iter():
    if elem.text:
        for ch in elem.text:
            code = ord(ch)
            if code < 0x09 or (code > 0x0D and code < 0x20):
                problem_count += 1
                if problem_count <= 3:
                    tag = elem.tag.split('}')[1] if '}' in elem.tag else elem.tag
                    print(f'  U+{code:04X} in <{tag}>: {repr(elem.text[:80])}')
                break
    if elem.tail:
        for ch in elem.tail:
            code = ord(ch)
            if code < 0x09 or (code > 0x0D and code < 0x20):
                problem_count += 1
                if problem_count <= 3:
                    tag = elem.tag.split('}')[1] if '}' in elem.tag else elem.tag
                    print(f'  U+{code:04X} in tail of <{tag}>: {repr(elem.tail[:80])}')
                break
print(f'Total elements with control chars: {problem_count}')
