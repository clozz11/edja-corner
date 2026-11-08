#!/usr/bin/env python3
"""
Script untuk membuat file Excel Data Penjualan Kedai Kopi Semat Space
VERSI 2 - Dengan data tambahan Oktober
Berisi 2 Sheet:
  1. Data Penjualan (Master Table)
  2. Analisis - Total pendapatan per menu per bulan + TOTAL OMZET
"""

import zipfile
from xml.etree.ElementTree import Element, SubElement, tostring

# ============================================================
# DATA PENJUALAN DARI SCREENSHOT APLIKASI KASIR
# ============================================================

# AGUSTUS (Total: Rp5.846.000)
agustus_data = [
    ("Americano", None, 49, 735000),
    ("Americano", "Cold / Buhaira (Arabica)", 23, 345000),
    ("Americano", "Cold / Sahara (70:30)", 24, 360000),
    ("Americano", "Hot / Buhaira (Arabica)", 1, 15000),
    ("Americano", "Hot / Sahara (70:30)", 1, 15000),
    ("Kopi Susu Saring", None, 49, 686000),
    ("Kopi Susu Saring", "Cold", 26, 364000),
    ("Kopi Susu Saring", "Hot", 23, 322000),
    ("Manuhara", None, 27, 540000),
    ("Manuhara", "Ice Cream", 10, 200000),
    ("Manuhara", "Original", 13, 260000),
    ("Manuhara", "Whip Cream", 4, 80000),
    ("Chocolate", None, 27, 486000),
    ("Chocolate", "Cold", 21, 378000),
    ("Chocolate", "Hot", 6, 108000),
    ("Wedges Fries", None, 27, 405000),
    ("Senoaji", None, 18, 360000),
    ("Matcha Bomb", None, 15, 300000),
    ("Banana Chocolate Smoothie", None, 14, 280000),
    ("V60", None, 14, 252000),
    ("V60", "Cold", 13, 234000),
    ("V60", "Hot", 1, 18000),
    ("Chikin Baito", None, 14, 230000),
    ("Chikin Baito", "Cabe Garam", 10, 170000),
    ("Chikin Baito", "Original", 4, 60000),
    ("Messy Sand", None, 10, 200000),
    ("Cappuccino", None, 10, 180000),
    ("Cappuccino", "Buhaira (Arabica)", 4, 72000),
    ("Cappuccino", "Sahara (70:30)", 6, 108000),
    ("Indomie Goreng", None, 12, 174000),
    ("Indomie Goreng", "Blackpepper / Double", 10, 150000),
    ("Indomie Goreng", "Blackpepper / Single", 2, 24000),
    ("Sembadra", None, 24, 480000),
    ("Sembadra", "Ice Cream", 17, 340000),
    ("Sembadra", "Original", 5, 100000),
    ("Sembadra", "Whip Cream", 2, 40000),
    ("Tubruk Robusta", None, 15, 150000),
    ("Hok Lo Pan", None, 7, 140000),
    ("Banana Peanut Butter Smoothie", None, 7, 140000),
    ("Nasi Mentega Cabe Garam", None, 3, 66000),
    ("Indomie Rebus", None, 3, 39000),
    ("Indomie Rebus", "Blackpepper / Double", 1, 15000),
    ("Indomie Rebus", "Blackpepper / Single", 2, 24000),
    ("Extra Shot/Guest Bean", None, 1, 3000),
]

# SEPTEMBER (Total: Rp5.912.000)
september_data = [
    ("Americano", None, 47, 705000),
    ("Americano", "Cold / Buhaira (Arabica)", 30, 450000),
    ("Americano", "Cold / Sahara (70:30)", 10, 150000),
    ("Americano", "Hot / Buhaira (Arabica)", 6, 90000),
    ("Americano", "Hot / Sahara (70:30)", 1, 15000),
    ("Sembadra", None, 35, 700000),
    ("Sembadra", "Ice Cream", 22, 440000),
    ("Sembadra", "Original", 6, 120000),
    ("Sembadra", "Whip Cream", 7, 140000),
    ("Chocolate", None, 29, 522000),
    ("Chocolate", "Cold", 23, 414000),
    ("Chocolate", "Hot", 6, 108000),
    ("Kopi Susu Saring", None, 37, 518000),
    ("Kopi Susu Saring", "Cold", 20, 280000),
    ("Kopi Susu Saring", "Hot", 17, 238000),
    ("Manuhara", None, 25, 500000),
    ("Manuhara", "Ice Cream", 12, 240000),
    ("Manuhara", "Original", 10, 200000),
    ("Manuhara", "Whip Cream", 3, 60000),
    ("Nasi Mentega Cabe Garam", None, 18, 396000),
    ("V60", None, 19, 342000),
    ("V60", "Cold", 15, 270000),
    ("V60", "Hot", 4, 72000),
    ("Messy Sand", None, 16, 320000),
    ("Matcha Bomb", None, 15, 300000),
    ("Wedges Fries", None, 20, 300000),
    ("Banana Chocolate Smoothie", None, 13, 260000),
    ("Chikin Baito", None, 13, 211000),
    ("Chikin Baito", "Cabe Garam", 8, 136000),
    ("Chikin Baito", "Original", 5, 75000),
    ("Senoaji", None, 9, 180000),
    ("Banana Peanut Butter Smoothie", None, 9, 180000),
    ("Indomie Goreng", None, 12, 168000),
    ("Indomie Goreng", "Blackpepper / Double", 8, 120000),
    ("Indomie Goreng", "Blackpepper / Single", 4, 48000),
    ("Hok Lo Pan", None, 5, 100000),
    ("Cappuccino", None, 5, 90000),
    ("Cappuccino", "Buhaira (Arabica)", 1, 18000),
    ("Cappuccino", "Sahara (70:30)", 4, 72000),
    ("Tubruk Robusta", None, 9, 90000),
    ("Indomie Rebus", None, 2, 30000),
    ("Indomie Rebus", "Blackpepper / Double", 2, 30000),
]

# OKTOBER - UPDATED dengan data tambahan
# Total Oktober sebelumnya terlihat terpotong, sekarang ditambahkan:
# Mie Svarna, Donat, Rona Jingga, Bakso Goreng, Seruni, Rona Kelabu, Espresso
# Extra Shot/Guest Bean diupdate dari x1 ke x3
oktober_data = [
    ("Sembadra", None, 45, 900000),
    ("Sembadra", "Ice Cream", 26, 520000),
    ("Sembadra", "Original", 10, 200000),
    ("Sembadra", "Whip Cream", 9, 180000),
    ("Manuhara", None, 31, 620000),
    ("Manuhara", "Ice Cream", 15, 300000),
    ("Manuhara", "Original", 14, 280000),
    ("Manuhara", "Whip Cream", 2, 40000),
    ("Americano", None, 37, 563000),
    ("Americano", "Cold / Buhaira (Arabica)", 22, 335000),
    ("Americano", "Cold / Sahara (70:30)", 12, 182000),
    ("Americano", "Hot / Buhaira (Arabica)", 0, 0),
    ("Americano", "Hot / Sahara (70:30)", 3, 46000),
    ("Kopi Susu Saring", None, 37, 523000),
    ("Kopi Susu Saring", "Cold", 15, 213000),
    ("Kopi Susu Saring", "Hot", 22, 310000),
    ("Matcha Bomb", None, 24, 488000),
    ("V60", None, 23, 420000),
    ("V60", "Cold", 18, 330000),
    ("V60", "Hot", 5, 90000),
    ("Senoaji", None, 20, 400000),
    ("Chocolate", None, 22, 396000),
    ("Chocolate", "Cold", 20, 360000),
    ("Chocolate", "Hot", 2, 36000),
    ("Messy Sand", None, 19, 380000),
    ("Nasi Mentega Cabe Garam", None, 17, 374000),
    ("Banana Chocolate Smoothie", None, 16, 320000),
    ("Wedges Fries", None, 20, 303000),
    ("Chikin Baito", None, 16, 261000),
    ("Chikin Baito", "Cabe Garam", 10, 170000),
    ("Chikin Baito", "Original", 6, 91000),
    ("Banana Peanut Butter Smoothie", None, 11, 220000),
    ("Indomie Goreng", None, 14, 201000),
    ("Indomie Goreng", "Blackpepper / Double", 11, 165000),
    ("Indomie Goreng", "Blackpepper / Single", 3, 36000),
    ("Tubruk Robusta", None, 17, 177000),
    ("Indomie Rebus", None, 12, 159000),
    ("Indomie Rebus", "Blackpepper / Double", 5, 75000),
    ("Indomie Rebus", "Blackpepper / Single", 7, 84000),
    ("Cappuccino", None, 8, 148000),
    ("Cappuccino", "Buhaira (Arabica)", 1, 18000),
    ("Cappuccino", "Sahara (70:30)", 7, 130000),
    ("Mie Svarna", None, 7, 140000),
    ("Hok Lo Pan", None, 6, 120000),
    ("Donat", None, 6, 96000),
    ("Rona Jingga", None, 5, 80000),
    ("Bakso Goreng", None, 4, 68000),
    ("Seruni", None, 2, 30000),
    ("Seruni", "Hot", 1, 15000),
    ("Seruni", "Iced", 1, 15000),
    ("Rona Kelabu", None, 1, 17000),
    ("Rona Kelabu", "Hot", 1, 17000),
    ("Espresso", None, 1, 10000),
    ("Extra Shot/Guest Bean", None, 3, 9000),
]

# ============================================================
# Helper functions
# ============================================================
def get_main_items(data):
    return [(name, qty, revenue) for name, variant, qty, revenue in data if variant is None]

# Kumpulkan semua nama menu unik
all_menus_set = set()
for d in [agustus_data, september_data, oktober_data]:
    for name, variant, qty, revenue in d:
        if variant is None:
            all_menus_set.add(name)

def total_revenue(menu_name):
    total = 0
    for d in [agustus_data, september_data, oktober_data]:
        for name, variant, qty, revenue in d:
            if variant is None and name == menu_name:
                total += revenue
    return total

all_menus = sorted(all_menus_set, key=lambda m: total_revenue(m), reverse=True)

def make_lookup(data):
    lookup = {}
    for name, variant, qty, revenue in data:
        if variant is None:
            lookup[name] = (qty, revenue)
    return lookup

agustus_lookup = make_lookup(agustus_data)
september_lookup = make_lookup(september_data)
oktober_lookup = make_lookup(oktober_data)

# ============================================================
# Kategori menu
# ============================================================
kategori_map = {
    "Americano": "Kopi",
    "Kopi Susu Saring": "Kopi",
    "Sembadra": "Signature",
    "Manuhara": "Signature",
    "Senoaji": "Signature",
    "Chocolate": "Non-Kopi",
    "Matcha Bomb": "Non-Kopi",
    "V60": "Kopi",
    "Cappuccino": "Kopi",
    "Tubruk Robusta": "Kopi",
    "Espresso": "Kopi",
    "Messy Sand": "Non-Kopi",
    "Banana Chocolate Smoothie": "Non-Kopi",
    "Banana Peanut Butter Smoothie": "Non-Kopi",
    "Seruni": "Non-Kopi",
    "Rona Jingga": "Non-Kopi",
    "Rona Kelabu": "Non-Kopi",
    "Wedges Fries": "Makanan",
    "Chikin Baito": "Makanan",
    "Indomie Goreng": "Makanan",
    "Indomie Rebus": "Makanan",
    "Nasi Mentega Cabe Garam": "Makanan",
    "Hok Lo Pan": "Makanan",
    "Mie Svarna": "Makanan",
    "Donat": "Camilan",
    "Bakso Goreng": "Camilan",
    "Extra Shot/Guest Bean": "Lainnya",
}

# ============================================================
# MEMBUAT FILE XLSX
# ============================================================
NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

def col_letter(idx):
    result = ""
    idx += 1
    while idx > 0:
        idx -= 1
        result = chr(65 + idx % 26) + result
        idx //= 26
    return result

def make_xlsx(filename):
    shared_strings = []
    ss_index = {}

    def get_ss(text):
        text = str(text)
        if text not in ss_index:
            ss_index[text] = len(shared_strings)
            shared_strings.append(text)
        return ss_index[text]

    # Pre-populate shared strings
    headers1 = ["No", "Nama Menu", "Kategori",
                 "Harga Agst (Rp)", "Terjual Agst (Pcs)",
                 "Harga Sept (Rp)", "Terjual Sept (Pcs)",
                 "Harga Okt (Rp)", "Terjual Okt (Pcs)",
                 "Pendapatan Agst (Rp)", "Pendapatan Sept (Rp)", "Pendapatan Okt (Rp)"]
    for h in headers1:
        get_ss(h)

    headers2 = ["Nama Menu", "Total Pendapatan Agustus", "Total Pendapatan September",
                 "Total Pendapatan Oktober", "Total 3 Bulan"]
    for h in headers2:
        get_ss(h)
    get_ss("TOTAL OMZET")

    for menu in all_menus:
        get_ss(menu)
        get_ss(kategori_map.get(menu, "Lainnya"))

    # ---- Build Sheet 1: Data Penjualan ----
    sheet1_rows = []
    sheet1_rows.append(("header", headers1))

    for i, menu in enumerate(all_menus):
        aq, ar = agustus_lookup.get(menu, (0, 0))
        sq, sr = september_lookup.get(menu, (0, 0))
        oq, orev = oktober_lookup.get(menu, (0, 0))

        harga_agst = round(ar / aq) if aq > 0 else 0
        harga_sept = round(sr / sq) if sq > 0 else 0
        harga_okt = round(orev / oq) if oq > 0 else 0

        kategori = kategori_map.get(menu, "Lainnya")
        row = [i+1, menu, kategori, harga_agst, aq, harga_sept, sq, harga_okt, oq, ar, sr, orev]
        sheet1_rows.append(("data", row))

    # ---- Build Sheet 2: Analisis ----
    sheet2_rows = []
    sheet2_rows.append(("header", headers2))

    total_agst = 0
    total_sept = 0
    total_okt = 0

    for menu in all_menus:
        aq, ar = agustus_lookup.get(menu, (0, 0))
        sq, sr = september_lookup.get(menu, (0, 0))
        oq, orev = oktober_lookup.get(menu, (0, 0))
        total_3 = ar + sr + orev

        total_agst += ar
        total_sept += sr
        total_okt += orev

        sheet2_rows.append(("data", [menu, ar, sr, orev, total_3]))

    grand_total = total_agst + total_sept + total_okt
    sheet2_rows.append(("total", ["TOTAL OMZET", total_agst, total_sept, total_okt, grand_total]))

    # ---- XML generators ----
    def build_shared_strings_xml():
        root = Element("sst")
        root.set("xmlns", NAMESPACE)
        root.set("count", str(len(shared_strings)))
        root.set("uniqueCount", str(len(shared_strings)))
        for s in shared_strings:
            si = SubElement(root, "si")
            t = SubElement(si, "t")
            t.set("xml:space", "preserve")
            t.text = s
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + tostring(root, encoding="unicode")

    def build_styles_xml():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="#,##0"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDCE6F1"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color auto="1"/></left>
      <right style="thin"><color auto="1"/></right>
      <top style="thin"><color auto="1"/></top>
      <bottom style="thin"><color auto="1"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>'''

    def build_sheet_xml(rows, col_widths):
        lines = []
        lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        lines.append(f'<worksheet xmlns="{NAMESPACE}">')
        lines.append('<cols>')
        for i, w in enumerate(col_widths):
            lines.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
        lines.append('</cols>')
        lines.append('<sheetData>')

        for row_idx, (row_type, row_data) in enumerate(rows):
            row_num = row_idx + 1
            lines.append(f'<row r="{row_num}">')
            for col_idx, val in enumerate(row_data):
                cell_ref = f"{col_letter(col_idx)}{row_num}"
                if row_type == "header":
                    ss_idx = get_ss(str(val))
                    lines.append(f'<c r="{cell_ref}" s="1" t="s"><v>{ss_idx}</v></c>')
                elif row_type == "total":
                    if isinstance(val, str):
                        ss_idx = get_ss(val)
                        lines.append(f'<c r="{cell_ref}" s="5" t="s"><v>{ss_idx}</v></c>')
                    else:
                        lines.append(f'<c r="{cell_ref}" s="4"><v>{val}</v></c>')
                else:
                    if isinstance(val, str):
                        ss_idx = get_ss(val)
                        lines.append(f'<c r="{cell_ref}" s="2" t="s"><v>{ss_idx}</v></c>')
                    elif isinstance(val, (int, float)):
                        lines.append(f'<c r="{cell_ref}" s="3"><v>{val}</v></c>')
                    else:
                        lines.append(f'<c r="{cell_ref}" s="2"><v>{val}</v></c>')
            lines.append('</row>')

        lines.append('</sheetData>')
        lines.append('</worksheet>')
        return '\n'.join(lines)

    def build_workbook_xml():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data Penjualan" sheetId="1" r:id="rId1"/>
    <sheet name="Analisis" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>'''

    def build_content_types():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>'''

    def build_rels():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    def build_workbook_rels():
        return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

    col_widths_1 = [5, 35, 12, 18, 18, 18, 18, 18, 18, 22, 22, 22]
    col_widths_2 = [35, 28, 28, 28, 22]

    sheet1_xml = build_sheet_xml(sheet1_rows, col_widths_1)
    sheet2_xml = build_sheet_xml(sheet2_rows, col_widths_2)

    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', build_content_types())
        zf.writestr('_rels/.rels', build_rels())
        zf.writestr('xl/workbook.xml', build_workbook_xml())
        zf.writestr('xl/_rels/workbook.xml.rels', build_workbook_rels())
        zf.writestr('xl/styles.xml', build_styles_xml())
        zf.writestr('xl/sharedStrings.xml', build_shared_strings_xml())
        zf.writestr('xl/worksheets/sheet1.xml', sheet1_xml)
        zf.writestr('xl/worksheets/sheet2.xml', sheet2_xml)

    print(f"✅ File '{filename}' berhasil dibuat!")

    # Print summary
    print(f"\n{'='*60}")
    print("RINGKASAN DATA PENJUALAN SEMAT SPACE")
    print(f"{'='*60}")
    print(f"\nTotal Menu: {len(all_menus)} item")
    print(f"\nTotal Omzet Agustus   : Rp {total_agst:>12,}".replace(",", "."))
    print(f"Total Omzet September : Rp {total_sept:>12,}".replace(",", "."))
    print(f"Total Omzet Oktober   : Rp {total_okt:>12,}".replace(",", "."))
    print(f"{'─'*45}")
    print(f"GRAND TOTAL 3 BULAN   : Rp {grand_total:>12,}".replace(",", "."))

    print(f"\n{'='*60}")
    print("SHEET 1: Data Penjualan (Master Table)")
    print(f"{'='*60}")
    fmt = f"{'No':<4} {'Nama Menu':<32} {'Kategori':<12} {'Agst Qty':>9} {'Agst Rev':>12} {'Sept Qty':>9} {'Sept Rev':>12} {'Okt Qty':>9} {'Okt Rev':>12}"
    print(fmt)
    print("─" * 120)
    for i, menu in enumerate(all_menus):
        aq, ar = agustus_lookup.get(menu, (0, 0))
        sq, sr = september_lookup.get(menu, (0, 0))
        oq, orev = oktober_lookup.get(menu, (0, 0))
        kat = kategori_map.get(menu, "Lainnya")
        print(f"{i+1:<4} {menu:<32} {kat:<12} {aq:>9} {ar:>12,} {sq:>9} {sr:>12,} {oq:>9} {orev:>12,}".replace(",", "."))

if __name__ == "__main__":
    make_xlsx("/projects/sandbox/Data_Penjualan_Semat_Space.xlsx")
