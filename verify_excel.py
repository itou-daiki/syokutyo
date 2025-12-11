
import zipfile
import xml.etree.ElementTree as ET
import sys
import os

def parse_excel(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found {file_path}")
            return

        with zipfile.ZipFile(file_path, 'r') as z:
            # 1. Parse Shared Strings
            shared_strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    for si in root.findall('main:si', ns):
                        t = si.find('main:t', ns)
                        if t is not None and t.text:
                            shared_strings.append(t.text)
                        else:
                            texts = [node.text for node in si.findall('.//main:t', ns) if node.text]
                            shared_strings.append("".join(texts))

            # 2. Parse Workbook to get Sheet Names and IDs
            sheet_mapping = {} 
            with z.open('xl/workbook.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                sheets = root.find('main:sheets', ns)
                for sheet in sheets.findall('main:sheet', ns):
                    name = sheet.attrib.get('name')
                    r_id = sheet.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    sheet_mapping[r_id] = name

            # 3. Map Workbook Relationships
            wb_rels = {}
            with z.open('xl/_rels/workbook.xml.rels') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                for rel in root.findall('rel:Relationship', ns):
                    wb_rels[rel.attrib.get('Id')] = rel.attrib.get('Target')

            # 4. Iterate sheets
            print(f"File: {os.path.basename(file_path)}")
            print("-" * 40)
            
            for r_id, sheet_name in sheet_mapping.items():
                target = wb_rels.get(r_id)
                if not target: continue
                target_path = f"xl/{target}" if not target.startswith('/') else target[1:]
                
                if target_path not in z.namelist():
                    target_path = target 
                    if target_path not in z.namelist(): continue

                with z.open(target_path) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    
                    sheet_data = root.find('main:sheetData', ns)
                    rows = sheet_data.findall('main:row', ns)
                    
                    headers = []
                    row2 = []
                    
                    def get_val(c):
                        v = c.find('main:v', ns)
                        t = c.attrib.get('t')
                        if v is not None and v.text:
                            if t == 's':
                                idx = int(v.text)
                                return shared_strings[idx] if idx < len(shared_strings) else f"<Err {idx}>"
                            return v.text
                        return ""

                    if len(rows) > 0:
                        for c in rows[0].findall('main:c', ns): headers.append(get_val(c))
                    
                    if len(rows) > 1:
                        for c in rows[1].findall('main:c', ns): row2.append(get_val(c))

                    print(f"Sheet: {sheet_name}")
                    print(f"Headers: {headers}")
                    print(f"Row 2:   {row2}")
                    print("")

    except Exception as e:
        print(f"Error parsing Excel: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        parse_excel(sys.argv[1])
