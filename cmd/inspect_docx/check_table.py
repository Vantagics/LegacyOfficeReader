import zipfile
z = zipfile.ZipFile('testfie/test.docx')
doc = z.read('word/document.xml').decode('utf-8')

# Find the table section
tbl_start = doc.find('<w:tbl>')
tbl_end = doc.find('</w:tbl>') + len('</w:tbl>')
if tbl_start >= 0:
    tbl = doc[tbl_start:tbl_end]
    print(f'Table XML length: {len(tbl)}')
    print(f'First 1000 chars:')
    print(tbl[:1000])
    print(f'...')
    print(f'Last 500 chars:')
    print(tbl[-500:])
    
    print(f'\nRow count: {tbl.count("<w:tr>")}')
    print(f'Cell count: {tbl.count("<w:tc>")}')
else:
    print('No table found')
