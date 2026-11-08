#!/usr/bin/env python3
"""
FINAL VERSION - Excel Data Penjualan Semat Space
- Sheet 1: Data Penjualan (Master Table) with formulas for Pendapatan columns
- Sheet 2: Analisis with SUM formulas + Total 3 Bulan formula
- Sheet 3: Grafik with 3 embedded charts (Bar, Line, Pie)

Charts use inline data caches so they render correctly in Excel.
"""

import zipfile

# ============================================================
# DATA
# ============================================================

# Each entry: (menu, qty, revenue) — main items only (no variants)
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

# All menus sorted by total 3-month revenue
all_menus_set = set(agustus.keys()) | set(september.keys()) | set(oktober.keys())
def total_rev(m):
    return agustus.get(m,(0,0))[1] + september.get(m,(0,0))[1] + oktober.get(m,(0,0))[1]
all_menus = sorted(all_menus_set, key=total_rev, reverse=True)

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

N = len(all_menus)  # number of menu rows

# ============================================================
# CHART DATA
# ============================================================

# Bar: Top 5 Oktober
top5_okt = sorted(oktober.items(), key=lambda x: x[1][1], reverse=True)[:5]
bar_names = [t[0] for t in top5_okt]
bar_vals  = [t[1][1] for t in top5_okt]

# Line: 3 menu tren
line_menus = ["Sembadra", "Americano", "Kopi Susu Saring"]
line_months = ["Agustus", "September", "Oktober"]
line_vals = {}
for m in line_menus:
    line_vals[m] = [agustus.get(m,(0,0))[1], september.get(m,(0,0))[1], oktober.get(m,(0,0))[1]]

# Pie: kategori Oktober
kat_okt = {}
for name, (q, r) in oktober.items():
    k = kategori_map.get(name, "Lainnya")
    kat_okt[k] = kat_okt.get(k, 0) + r
pie_names = list(kat_okt.keys())
pie_vals  = list(kat_okt.values())

COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9B57A0"]

# ============================================================
# SHARED STRINGS
# ============================================================
_ss_list = []
_ss_map = {}

def ss(text):
    text = str(text)
    if text not in _ss_map:
        _ss_map[text] = len(_ss_list)
        _ss_list.append(text)
    return _ss_map[text]

def col_letter(c):
    r = ""
    c += 1
    while c > 0:
        c -= 1
        r = chr(65 + c % 26) + r
        c //= 26
    return r

# ============================================================
# SHEET 1: Data Penjualan — with formulas for Pendapatan
# ============================================================

def build_sheet1():
    """
    Columns (0-indexed):
    A=No, B=Nama Menu, C=Kategori,
    D=Harga Agst, E=Terjual Agst,
    F=Harga Sept, G=Terjual Sept,
    H=Harga Okt, I=Terjual Okt,
    J=Pendapatan Agst (=D*E), K=Pendapatan Sept (=F*G), L=Pendapatan Okt (=H*I)
    """
    hdrs = ["No", "Nama Menu", "Kategori",
            "Harga Agst (Rp)", "Terjual Agst (Pcs)",
            "Harga Sept (Rp)", "Terjual Sept (Pcs)",
            "Harga Okt (Rp)", "Terjual Okt (Pcs)",
            "Pendapatan Agst (Rp)", "Pendapatan Sept (Rp)", "Pendapatan Okt (Rp)"]

    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    lines.append('<cols>')
    widths = [5, 35, 12, 18, 18, 18, 18, 18, 18, 22, 22, 22]
    for i, w in enumerate(widths):
        lines.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    lines.append('</cols>')
    lines.append('<sheetData>')

    # Header row
    lines.append('<row r="1">')
    for ci, h in enumerate(hdrs):
        ref = f"{col_letter(ci)}1"
        idx = ss(h)
        lines.append(f'<c r="{ref}" s="1" t="s"><v>{idx}</v></c>')
    lines.append('</row>')

    # Data rows
    for ri, menu in enumerate(all_menus):
        row = ri + 2
        aq, ar = agustus.get(menu, (0, 0))
        sq, sr = september.get(menu, (0, 0))
        oq, ov = oktober.get(menu, (0, 0))
        ha = round(ar / aq) if aq > 0 else 0
        hs = round(sr / sq) if sq > 0 else 0
        ho = round(ov / oq) if oq > 0 else 0
        kat = kategori_map.get(menu, "Lainnya")

        lines.append(f'<row r="{row}">')
        # A: No
        lines.append(f'<c r="A{row}" s="3"><v>{ri+1}</v></c>')
        # B: Nama Menu
        idx = ss(menu)
        lines.append(f'<c r="B{row}" s="2" t="s"><v>{idx}</v></c>')
        # C: Kategori
        idx = ss(kat)
        lines.append(f'<c r="C{row}" s="2" t="s"><v>{idx}</v></c>')
        # D: Harga Agst
        lines.append(f'<c r="D{row}" s="3"><v>{ha}</v></c>')
        # E: Terjual Agst
        lines.append(f'<c r="E{row}" s="3"><v>{aq}</v></c>')
        # F: Harga Sept
        lines.append(f'<c r="F{row}" s="3"><v>{hs}</v></c>')
        # G: Terjual Sept
        lines.append(f'<c r="G{row}" s="3"><v>{sq}</v></c>')
        # H: Harga Okt
        lines.append(f'<c r="H{row}" s="3"><v>{ho}</v></c>')
        # I: Terjual Okt
        lines.append(f'<c r="I{row}" s="3"><v>{oq}</v></c>')
        # J: Pendapatan Agst = D*E  (FORMULA)
        lines.append(f'<c r="J{row}" s="3"><f>D{row}*E{row}</f><v>{ha*aq}</v></c>')
        # K: Pendapatan Sept = F*G  (FORMULA)
        lines.append(f'<c r="K{row}" s="3"><f>F{row}*G{row}</f><v>{hs*sq}</v></c>')
        # L: Pendapatan Okt = H*I  (FORMULA)
        lines.append(f'<c r="L{row}" s="3"><f>H{row}*I{row}</f><v>{ho*oq}</v></c>')
        lines.append('</row>')

    lines.append('</sheetData>')
    lines.append('</worksheet>')
    return '\n'.join(lines)

# ============================================================
# SHEET 2: Analisis — with SUM formulas
# ============================================================

def build_sheet2():
    """
    Columns:
    A=Nama Menu,
    B=Total Pendapatan Agustus (= 'Data Penjualan'!J row),
    C=Total Pendapatan September (= 'Data Penjualan'!K row),
    D=Total Pendapatan Oktober (= 'Data Penjualan'!L row),
    E=Total 3 Bulan (=B+C+D)
    Row N+2 = TOTAL OMZET with SUM formulas
    """
    hdrs = ["Nama Menu", "Total Pendapatan Agustus", "Total Pendapatan September",
            "Total Pendapatan Oktober", "Total 3 Bulan"]

    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    lines.append('<cols>')
    for i, w in enumerate([35, 28, 28, 28, 22]):
        lines.append(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>')
    lines.append('</cols>')
    lines.append('<sheetData>')

    # Header
    lines.append('<row r="1">')
    for ci, h in enumerate(hdrs):
        ref = f"{col_letter(ci)}1"
        idx = ss(h)
        lines.append(f'<c r="{ref}" s="1" t="s"><v>{idx}</v></c>')
    lines.append('</row>')

    # Data rows with formulas referencing Sheet 1
    for ri, menu in enumerate(all_menus):
        row = ri + 2
        s1row = ri + 2  # corresponding row in Data Penjualan
        ar = agustus.get(menu, (0,0))[1]
        sr = september.get(menu, (0,0))[1]
        ov = oktober.get(menu, (0,0))[1]

        lines.append(f'<row r="{row}">')
        # A: Nama Menu
        idx = ss(menu)
        lines.append(f'<c r="A{row}" s="2" t="s"><v>{idx}</v></c>')
        # B: Pendapatan Agst = reference to 'Data Penjualan'!J
        lines.append(f'<c r="B{row}" s="3"><f>\'Data Penjualan\'!J{s1row}</f><v>{ar}</v></c>')
        # C: Pendapatan Sept = reference to 'Data Penjualan'!K
        lines.append(f'<c r="C{row}" s="3"><f>\'Data Penjualan\'!K{s1row}</f><v>{sr}</v></c>')
        # D: Pendapatan Okt = reference to 'Data Penjualan'!L
        lines.append(f'<c r="D{row}" s="3"><f>\'Data Penjualan\'!L{s1row}</f><v>{ov}</v></c>')
        # E: Total 3 Bulan = B+C+D
        lines.append(f'<c r="E{row}" s="3"><f>B{row}+C{row}+D{row}</f><v>{ar+sr+ov}</v></c>')
        lines.append('</row>')

    # TOTAL OMZET row
    total_row = N + 2
    last_data = N + 1
    ta = sum(agustus.get(m,(0,0))[1] for m in all_menus)
    ts = sum(september.get(m,(0,0))[1] for m in all_menus)
    to_ = sum(oktober.get(m,(0,0))[1] for m in all_menus)
    gt = ta + ts + to_

    lines.append(f'<row r="{total_row}">')
    idx = ss("TOTAL OMZET")
    lines.append(f'<c r="A{total_row}" s="5" t="s"><v>{idx}</v></c>')
    lines.append(f'<c r="B{total_row}" s="4"><f>SUM(B2:B{last_data})</f><v>{ta}</v></c>')
    lines.append(f'<c r="C{total_row}" s="4"><f>SUM(C2:C{last_data})</f><v>{ts}</v></c>')
    lines.append(f'<c r="D{total_row}" s="4"><f>SUM(D2:D{last_data})</f><v>{to_}</v></c>')
    lines.append(f'<c r="E{total_row}" s="4"><f>SUM(E2:E{last_data})</f><v>{gt}</v></c>')
    lines.append('</row>')

    lines.append('</sheetData>')
    lines.append('</worksheet>')
    return '\n'.join(lines)

# ============================================================
# SHEET 3: Grafik (empty sheet with drawing reference)
# ============================================================

def build_sheet3():
    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                 ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    lines.append('<sheetData/>')
    lines.append('<drawing r:id="rId1"/>')
    lines.append('</worksheet>')
    return '\n'.join(lines)

# ============================================================
# CHART XMLs — Using proper numLit/strLit for inline data
# ============================================================

C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def chart1_bar():
    """Bar Chart: Top 5 Menu Oktober"""
    x = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}">
<c:chart>
  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="id-ID" sz="1200" b="1"/><a:t>Top 5 Menu Oktober - Pendapatan (Rp)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/>
  <c:plotArea>
    <c:layout/>
    <c:barChart>
      <c:barDir val="col"/>
      <c:grouping val="clustered"/>
      <c:varyColors val="1"/>
      <c:ser>
        <c:idx val="0"/>
        <c:order val="0"/>
        <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan Oktober</c:v></c:pt></c:strLit></c:tx>'''
    # Color each bar
    for i in range(len(bar_names)):
        x += f'''
        <c:dPt><c:idx val="{i}"/><c:spPr><a:solidFill><a:srgbClr val="{COLORS[i]}"/></a:solidFill></c:spPr></c:dPt>'''
    # Data labels
    x += '''
        <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>'''
    # Category
    x += f'''
        <c:cat><c:strLit><c:ptCount val="{len(bar_names)}"/>'''
    for i, n in enumerate(bar_names):
        x += f'<c:pt idx="{i}"><c:v>{n}</c:v></c:pt>'
    x += '</c:strLit></c:cat>'
    # Values
    x += f'''
        <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="{len(bar_vals)}"/>'''
    for i, v in enumerate(bar_vals):
        x += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    x += '</c:numLit></c:val>'
    x += '''
      </c:ser>
      <c:axId val="111111111"/>
      <c:axId val="222222222"/>
    </c:barChart>
    <c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222222222"/></c:catAx>
    <c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="111111111"/></c:valAx>
  </c:plotArea>
  <c:plotVisOnly val="1"/>
</c:chart>
</c:chartSpace>'''
    return x

def chart2_line():
    """Line Chart: Tren 3 Bulan"""
    x = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}">
<c:chart>
  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="id-ID" sz="1200" b="1"/><a:t>Tren Pendapatan 3 Bulan - Menu Terlaris</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/>
  <c:plotArea>
    <c:layout/>
    <c:lineChart>
      <c:grouping val="standard"/>
      <c:varyColors val="0"/>'''

    for si, menu in enumerate(line_menus):
        vals = line_vals[menu]
        color = COLORS[si]
        x += f'''
      <c:ser>
        <c:idx val="{si}"/>
        <c:order val="{si}"/>
        <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{menu}</c:v></c:pt></c:strLit></c:tx>
        <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>
        <c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr></c:marker>
        <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/></c:dLbls>
        <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>Agustus</c:v></c:pt><c:pt idx="1"><c:v>September</c:v></c:pt><c:pt idx="2"><c:v>Oktober</c:v></c:pt></c:strLit></c:cat>
        <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="3"/>'''
        for i, v in enumerate(vals):
            x += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
        x += '''</c:numLit></c:val>
        <c:smooth val="0"/>
      </c:ser>'''

    x += '''
      <c:marker val="1"/>
      <c:axId val="333333333"/>
      <c:axId val="444444444"/>
    </c:lineChart>
    <c:catAx><c:axId val="333333333"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="444444444"/></c:catAx>
    <c:valAx><c:axId val="444444444"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:crossAx val="333333333"/></c:valAx>
  </c:plotArea>
  <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
  <c:plotVisOnly val="1"/>
</c:chart>
</c:chartSpace>'''
    return x

def chart3_pie():
    """Pie Chart: Kontribusi Kategori Oktober"""
    x = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C}" xmlns:a="{A}" xmlns:r="{R}">
<c:chart>
  <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="1"/></a:pPr><a:r><a:rPr lang="id-ID" sz="1200" b="1"/><a:t>Kontribusi Kategori - Pendapatan Oktober</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/>
  <c:plotArea>
    <c:layout/>
    <c:pieChart>
      <c:varyColors val="1"/>
      <c:ser>
        <c:idx val="0"/>
        <c:order val="0"/>
        <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan Oktober</c:v></c:pt></c:strLit></c:tx>'''
    for i in range(len(pie_names)):
        x += f'''
        <c:dPt><c:idx val="{i}"/><c:spPr><a:solidFill><a:srgbClr val="{COLORS[i]}"/></a:solidFill></c:spPr></c:dPt>'''
    x += '''
        <c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/></c:dLbls>'''
    x += f'''
        <c:cat><c:strLit><c:ptCount val="{len(pie_names)}"/>'''
    for i, n in enumerate(pie_names):
        x += f'<c:pt idx="{i}"><c:v>{n}</c:v></c:pt>'
    x += '</c:strLit></c:cat>'
    x += f'''
        <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="{len(pie_vals)}"/>'''
    for i, v in enumerate(pie_vals):
        x += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    x += '''</c:numLit></c:val>
      </c:ser>
    </c:pieChart>
  </c:plotArea>
  <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
  <c:plotVisOnly val="1"/>
</c:chart>
</c:chartSpace>'''
    return x

# ============================================================
# DRAWING XML
# ============================================================

DR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

def build_drawing():
    x = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{DR}" xmlns:a="{A}" xmlns:r="{R}">'''
    
    chart_positions = [
        (0, 0, 10, 18),    # Chart 1: row 0-18
        (0, 20, 10, 38),   # Chart 2: row 20-38
        (0, 40, 10, 58),   # Chart 3: row 40-58
    ]
    
    for i, (c1, r1, c2, r2) in enumerate(chart_positions):
        x += f'''
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="{i+2}" name="Chart {i+1}"/>
        <xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="{C}" r:id="rId{i+1}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>'''
    
    x += '\n</xdr:wsDr>'
    return x

# ============================================================
# PACKAGE FILES
# ============================================================

def build_shared_strings():
    parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    parts.append(f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(_ss_list)}" uniqueCount="{len(_ss_list)}">')
    for s in _ss_list:
        # Escape XML
        s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("'", "&apos;")
        parts.append(f'<si><t xml:space="preserve">{s}</t></si>')
    parts.append('</sst>')
    return '\n'.join(parts)

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
</styleSheet>'''

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

WORKBOOK = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data Penjualan" sheetId="1" r:id="rId1"/>
    <sheet name="Analisis" sheetId="2" r:id="rId2"/>
    <sheet name="Grafik" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>'''

WB_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

SHEET3_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>'''

DRAWING_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>'''

# ============================================================
# ASSEMBLE
# ============================================================

def main():
    fname = "/projects/sandbox/Data_Penjualan_Semat_Space.xlsx"
    
    # Build sheets first (populates shared strings)
    s1 = build_sheet1()
    s2 = build_sheet2()
    s3 = build_sheet3()
    ss_xml = build_shared_strings()
    
    with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', ROOT_RELS)
        zf.writestr('xl/workbook.xml', WORKBOOK)
        zf.writestr('xl/_rels/workbook.xml.rels', WB_RELS)
        zf.writestr('xl/styles.xml', STYLES)
        zf.writestr('xl/sharedStrings.xml', ss_xml)
        zf.writestr('xl/worksheets/sheet1.xml', s1)
        zf.writestr('xl/worksheets/sheet2.xml', s2)
        zf.writestr('xl/worksheets/sheet3.xml', s3)
        zf.writestr('xl/worksheets/_rels/sheet3.xml.rels', SHEET3_RELS)
        zf.writestr('xl/drawings/drawing1.xml', build_drawing())
        zf.writestr('xl/drawings/_rels/drawing1.xml.rels', DRAWING_RELS)
        zf.writestr('xl/charts/chart1.xml', chart1_bar())
        zf.writestr('xl/charts/chart2.xml', chart2_line())
        zf.writestr('xl/charts/chart3.xml', chart3_pie())
    
    print(f"✅ File '{fname}' berhasil dibuat!")
    print()
    print("📋 SHEET 1 - Data Penjualan:")
    print("   • Kolom J: =D*E (Pendapatan Agustus = Harga × Qty)")
    print("   • Kolom K: =F*G (Pendapatan September = Harga × Qty)")
    print("   • Kolom L: =H*I (Pendapatan Oktober = Harga × Qty)")
    print()
    print("📋 SHEET 2 - Analisis:")
    print("   • Kolom B: ='Data Penjualan'!J (referensi ke Sheet 1)")
    print("   • Kolom C: ='Data Penjualan'!K")
    print("   • Kolom D: ='Data Penjualan'!L")
    print("   • Kolom E: =B+C+D (Total 3 Bulan)")
    print("   • Baris TOTAL OMZET: =SUM(...)")
    print()
    print("📊 SHEET 3 - Grafik (3 Charts):")
    print(f"   1. Bar Chart  : Top 5 Menu Oktober")
    for n, v in zip(bar_names, bar_vals):
        print(f"      - {n}: Rp{v:,}".replace(",","."))
    print(f"   2. Line Chart : Tren 3 Bulan")
    for m in line_menus:
        vv = line_vals[m]
        print(f"      - {m}: {' → '.join(f'Rp{v:,}'.replace(',','.') for v in vv)}")
    print(f"   3. Pie Chart  : Kontribusi Kategori Oktober")
    tot = sum(pie_vals)
    for n, v in zip(pie_names, pie_vals):
        print(f"      - {n}: Rp{v:,} ({v/tot*100:.1f}%)".replace(",","."))

if __name__ == "__main__":
    main()
