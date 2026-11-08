#!/usr/bin/env python3
"""
V6 - Excel Data Penjualan Semat Space
NO embedded charts (user will create manually).
3 sheets: Data Penjualan (formulas), Analisis (formulas), Grafik (chart source data).
All data uses shared strings - no inlineStr.
"""
import zipfile

# ============================================================
# DATA
# ============================================================
agustus = {
    "Americano": (49, 735000), "Kopi Susu Saring": (49, 686000),
    "Manuhara": (27, 540000), "Chocolate": (27, 486000),
    "Sembadra": (24, 480000), "Wedges Fries": (27, 405000),
    "Senoaji": (18, 360000), "Matcha Bomb": (15, 300000),
    "Banana Chocolate Smoothie": (14, 280000), "V60": (14, 252000),
    "Chikin Baito": (14, 230000), "Messy Sand": (10, 200000),
    "Cappuccino": (10, 180000), "Indomie Goreng": (12, 174000),
    "Tubruk Robusta": (15, 150000), "Hok Lo Pan": (7, 140000),
    "Banana Peanut Butter Smoothie": (7, 140000),
    "Nasi Mentega Cabe Garam": (3, 66000), "Indomie Rebus": (3, 39000),
    "Extra Shot/Guest Bean": (1, 3000),
}
september = {
    "Americano": (47, 705000), "Sembadra": (35, 700000),
    "Chocolate": (29, 522000), "Kopi Susu Saring": (37, 518000),
    "Manuhara": (25, 500000), "Nasi Mentega Cabe Garam": (18, 396000),
    "V60": (19, 342000), "Messy Sand": (16, 320000),
    "Matcha Bomb": (15, 300000), "Wedges Fries": (20, 300000),
    "Banana Chocolate Smoothie": (13, 260000), "Chikin Baito": (13, 211000),
    "Senoaji": (9, 180000), "Banana Peanut Butter Smoothie": (9, 180000),
    "Indomie Goreng": (12, 168000), "Hok Lo Pan": (5, 100000),
    "Cappuccino": (5, 90000), "Tubruk Robusta": (9, 90000),
    "Indomie Rebus": (2, 30000),
}
oktober = {
    "Sembadra": (45, 900000), "Manuhara": (31, 620000),
    "Americano": (37, 563000), "Kopi Susu Saring": (37, 523000),
    "Matcha Bomb": (24, 488000), "V60": (23, 420000),
    "Senoaji": (20, 400000), "Chocolate": (22, 396000),
    "Messy Sand": (19, 380000), "Nasi Mentega Cabe Garam": (17, 374000),
    "Banana Chocolate Smoothie": (16, 320000), "Wedges Fries": (20, 303000),
    "Chikin Baito": (16, 261000), "Banana Peanut Butter Smoothie": (11, 220000),
    "Indomie Goreng": (14, 201000), "Tubruk Robusta": (17, 177000),
    "Indomie Rebus": (12, 159000), "Cappuccino": (8, 148000),
    "Mie Svarna": (7, 140000), "Hok Lo Pan": (6, 120000),
    "Donat": (6, 96000), "Rona Jingga": (5, 80000),
    "Bakso Goreng": (4, 68000), "Seruni": (2, 30000),
    "Rona Kelabu": (1, 17000), "Espresso": (1, 10000),
    "Extra Shot/Guest Bean": (3, 9000),
}

all_menus_set = set(agustus) | set(september) | set(oktober)
def total_rev(m):
    return agustus.get(m,(0,0))[1] + september.get(m,(0,0))[1] + oktober.get(m,(0,0))[1]
all_menus = sorted(all_menus_set, key=total_rev, reverse=True)
N = len(all_menus)

kategori_map = {
    "Americano": "Kopi", "Kopi Susu Saring": "Kopi", "V60": "Kopi",
    "Cappuccino": "Kopi", "Tubruk Robusta": "Kopi", "Espresso": "Kopi",
    "Sembadra": "Signature", "Manuhara": "Signature", "Senoaji": "Signature",
    "Chocolate": "Non-Kopi", "Matcha Bomb": "Non-Kopi", "Messy Sand": "Non-Kopi",
    "Banana Chocolate Smoothie": "Non-Kopi", "Banana Peanut Butter Smoothie": "Non-Kopi",
    "Seruni": "Non-Kopi", "Rona Jingga": "Non-Kopi", "Rona Kelabu": "Non-Kopi",
    "Wedges Fries": "Makanan", "Chikin Baito": "Makanan", "Indomie Goreng": "Makanan",
    "Indomie Rebus": "Makanan", "Nasi Mentega Cabe Garam": "Makanan",
    "Hok Lo Pan": "Makanan", "Mie Svarna": "Makanan",
    "Donat": "Camilan", "Bakso Goreng": "Camilan",
    "Extra Shot/Guest Bean": "Lainnya",
}

# Chart source data
top5 = sorted(oktober.items(), key=lambda x: x[1][1], reverse=True)[:5]
bar_names = [t[0] for t in top5]
bar_vals = [t[1][1] for t in top5]

line_menus = ["Sembadra", "Americano", "Kopi Susu Saring"]
line_vals = {m: [agustus.get(m,(0,0))[1], september.get(m,(0,0))[1], oktober.get(m,(0,0))[1]] for m in line_menus}

kat_okt = {}
for name, (q, r) in oktober.items():
    k = kategori_map.get(name, "Lainnya")
    kat_okt[k] = kat_okt.get(k, 0) + r
pie_order = ["Kopi", "Signature", "Non-Kopi", "Makanan", "Camilan", "Lainnya"]
pie_names = [k for k in pie_order if k in kat_okt]
pie_vals = [kat_okt[k] for k in pie_names]

# Shared strings
_ss = []; _si = {}
def ss(t):
    t = str(t)
    if t not in _si: _si[t] = len(_ss); _ss.append(t)
    return _si[t]

def CL(c):
    r = ""; c += 1
    while c > 0: c -= 1; r = chr(65+c%26) + r; c //= 26
    return r

def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("'","&apos;")

# ============================================================
# SHEET 1: Data Penjualan
# ============================================================
def sheet1():
    h = ["No","Nama Menu","Kategori","Harga Agst (Rp)","Terjual Agst (Pcs)",
         "Harga Sept (Rp)","Terjual Sept (Pcs)","Harga Okt (Rp)","Terjual Okt (Pcs)",
         "Pendapatan Agst (Rp)","Pendapatan Sept (Rp)","Pendapatan Okt (Rp)"]
    L = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
         '<cols>']
    for i,w in enumerate([5,35,12,18,18,18,18,18,18,22,22,22]):
        L.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    L.append('</cols><sheetData><row r="1">')
    for ci,hd in enumerate(h):
        L.append(f'<c r="{CL(ci)}1" s="1" t="s"><v>{ss(hd)}</v></c>')
    L.append('</row>')
    for ri,m in enumerate(all_menus):
        r = ri+2
        aq,ar = agustus.get(m,(0,0)); sq,sr = september.get(m,(0,0)); oq,ov = oktober.get(m,(0,0))
        ha = round(ar/aq) if aq else 0; hs = round(sr/sq) if sq else 0; ho = round(ov/oq) if oq else 0
        k = kategori_map.get(m,"Lainnya")
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="3"><v>{ri+1}</v></c>')
        L.append(f'<c r="B{r}" s="2" t="s"><v>{ss(m)}</v></c>')
        L.append(f'<c r="C{r}" s="2" t="s"><v>{ss(k)}</v></c>')
        L.append(f'<c r="D{r}" s="3"><v>{ha}</v></c>')
        L.append(f'<c r="E{r}" s="3"><v>{aq}</v></c>')
        L.append(f'<c r="F{r}" s="3"><v>{hs}</v></c>')
        L.append(f'<c r="G{r}" s="3"><v>{sq}</v></c>')
        L.append(f'<c r="H{r}" s="3"><v>{ho}</v></c>')
        L.append(f'<c r="I{r}" s="3"><v>{oq}</v></c>')
        L.append(f'<c r="J{r}" s="3"><f>D{r}*E{r}</f><v>{ha*aq}</v></c>')
        L.append(f'<c r="K{r}" s="3"><f>F{r}*G{r}</f><v>{hs*sq}</v></c>')
        L.append(f'<c r="L{r}" s="3"><f>H{r}*I{r}</f><v>{ho*oq}</v></c>')
        L.append('</row>')
    L.append('</sheetData></worksheet>')
    return '\n'.join(L)

# ============================================================
# SHEET 2: Analisis
# ============================================================
def sheet2():
    h = ["Nama Menu","Total Pendapatan Agustus","Total Pendapatan September",
         "Total Pendapatan Oktober","Total 3 Bulan"]
    L = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
         '<cols>']
    for i,w in enumerate([35,28,28,28,22]):
        L.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    L.append('</cols><sheetData><row r="1">')
    for ci,hd in enumerate(h):
        L.append(f'<c r="{CL(ci)}1" s="1" t="s"><v>{ss(hd)}</v></c>')
    L.append('</row>')
    ta=ts=to_=0
    for ri,m in enumerate(all_menus):
        r = ri+2
        ar = agustus.get(m,(0,0))[1]; sr = september.get(m,(0,0))[1]; ov = oktober.get(m,(0,0))[1]
        ta+=ar; ts+=sr; to_+=ov
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="2" t="s"><v>{ss(m)}</v></c>')
        L.append(f'<c r="B{r}" s="3"><f>\'Data Penjualan\'!J{r}</f><v>{ar}</v></c>')
        L.append(f'<c r="C{r}" s="3"><f>\'Data Penjualan\'!K{r}</f><v>{sr}</v></c>')
        L.append(f'<c r="D{r}" s="3"><f>\'Data Penjualan\'!L{r}</f><v>{ov}</v></c>')
        L.append(f'<c r="E{r}" s="3"><f>B{r}+C{r}+D{r}</f><v>{ar+sr+ov}</v></c>')
        L.append('</row>')
    tr = N+2; lr = N+1; gt = ta+ts+to_
    L.append(f'<row r="{tr}">')
    L.append(f'<c r="A{tr}" s="5" t="s"><v>{ss("TOTAL OMZET")}</v></c>')
    L.append(f'<c r="B{tr}" s="4"><f>SUM(B2:B{lr})</f><v>{ta}</v></c>')
    L.append(f'<c r="C{tr}" s="4"><f>SUM(C2:C{lr})</f><v>{ts}</v></c>')
    L.append(f'<c r="D{tr}" s="4"><f>SUM(D2:D{lr})</f><v>{to_}</v></c>')
    L.append(f'<c r="E{tr}" s="4"><f>SUM(E2:E{lr})</f><v>{gt}</v></c>')
    L.append('</row>')
    L.append('</sheetData></worksheet>')
    return '\n'.join(L)

# ============================================================
# SHEET 3: Grafik (DATA SOURCE for manual chart creation)
# NO drawing/chart XML — user creates charts manually
#
# Layout:
#   A1:B1  = Header "Menu" | "Pendapatan Oktober"
#   A2:B6  = Top 5 bar data
#
#   A8:D8  = Header "Bulan" | "Sembadra" | "Americano" | "Kopi Susu Saring"
#   A9:D11 = Line data (Agustus, September, Oktober)
#
#   A13:B13 = Header "Kategori" | "Pendapatan Oktober"
#   A14:B19 = Pie data
# ============================================================
def sheet3():
    L = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
         '<cols>']
    for i,w in enumerate([30, 25, 25, 25]):
        L.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    L.append('</cols><sheetData>')

    # === SECTION 1: Bar Chart Data (rows 1-6) ===
    # Header
    L.append(f'<row r="1">')
    L.append(f'<c r="A1" s="1" t="s"><v>{ss("Menu")}</v></c>')
    L.append(f'<c r="B1" s="1" t="s"><v>{ss("Pendapatan Oktober")}</v></c>')
    L.append(f'</row>')
    # Data
    for i, (name, val) in enumerate(zip(bar_names, bar_vals)):
        r = i + 2
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="2" t="s"><v>{ss(name)}</v></c>')
        L.append(f'<c r="B{r}" s="3"><v>{val}</v></c>')
        L.append(f'</row>')

    # === SECTION 2: Line Chart Data (rows 8-11) ===
    # Header
    L.append(f'<row r="8">')
    L.append(f'<c r="A8" s="1" t="s"><v>{ss("Bulan")}</v></c>')
    L.append(f'<c r="B8" s="1" t="s"><v>{ss("Sembadra")}</v></c>')
    L.append(f'<c r="C8" s="1" t="s"><v>{ss("Americano")}</v></c>')
    L.append(f'<c r="D8" s="1" t="s"><v>{ss("Kopi Susu Saring")}</v></c>')
    L.append(f'</row>')
    # Data rows
    months = ["Agustus", "September", "Oktober"]
    for i, month in enumerate(months):
        r = 9 + i
        vals = [line_vals[m][i] for m in line_menus]
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="2" t="s"><v>{ss(month)}</v></c>')
        L.append(f'<c r="B{r}" s="3"><v>{vals[0]}</v></c>')
        L.append(f'<c r="C{r}" s="3"><v>{vals[1]}</v></c>')
        L.append(f'<c r="D{r}" s="3"><v>{vals[2]}</v></c>')
        L.append(f'</row>')

    # === SECTION 3: Pie Chart Data (rows 13-19) ===
    # Header
    L.append(f'<row r="13">')
    L.append(f'<c r="A13" s="1" t="s"><v>{ss("Kategori")}</v></c>')
    L.append(f'<c r="B13" s="1" t="s"><v>{ss("Pendapatan Oktober")}</v></c>')
    L.append(f'</row>')
    # Data
    for i, (name, val) in enumerate(zip(pie_names, pie_vals)):
        r = 14 + i
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="2" t="s"><v>{ss(name)}</v></c>')
        L.append(f'<c r="B{r}" s="3"><v>{val}</v></c>')
        L.append(f'</row>')

    L.append('</sheetData></worksheet>')
    return '\n'.join(L)

# ============================================================
# BUILD — clean file, NO charts, NO drawings
# ============================================================
def build():
    fname = "/projects/sandbox/Data_Penjualan_Semat_Space.xlsx"
    s1 = sheet1(); s2 = sheet2(); s3 = sheet3()

    ss_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(_ss)}" uniqueCount="{len(_ss)}">\n'
    for s in _ss:
        ss_xml += f'<si><t xml:space="preserve">{esc(s)}</t></si>\n'
    ss_xml += '</sst>'

    with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as z:
        # Content Types — NO chart or drawing overrides
        z.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>''')
        z.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>''')
        z.writestr('xl/workbook.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data Penjualan" sheetId="1" r:id="rId1"/>
    <sheet name="Analisis" sheetId="2" r:id="rId2"/>
    <sheet name="Grafik" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>''')
        z.writestr('xl/_rels/workbook.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>''')
        z.writestr('xl/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0"/></numFmts>
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
    <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>''')
        z.writestr('xl/sharedStrings.xml', ss_xml)
        z.writestr('xl/worksheets/sheet1.xml', s1)
        z.writestr('xl/worksheets/sheet2.xml', s2)
        z.writestr('xl/worksheets/sheet3.xml', s3)

    print(f"✅ {fname}")
    print(f"   Total menu: {N}")
    print(f"   Sheet 1: Data Penjualan (with formulas)")
    print(f"   Sheet 2: Analisis (with SUM & cross-sheet formulas)")
    print(f"   Sheet 3: Grafik (chart source data - NO embedded charts)")
    print()
    print("📊 Sheet Grafik - Data siap untuk buat chart manual:")
    print(f"   Bar Chart  : Select A1:B6  → Insert → Column")
    print(f"   Line Chart : Select A8:D11 → Insert → Line")
    print(f"   Pie Chart  : Select A13:B{13+len(pie_names)} → Insert → Pie")

if __name__ == "__main__":
    build()
