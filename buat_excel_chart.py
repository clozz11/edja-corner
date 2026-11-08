#!/usr/bin/env python3
"""
Script untuk membuat file Excel Data Penjualan Kedai Kopi Semat Space
VERSI 3 - Dengan Charts (Bar, Line, Pie)

Charts dibuat menggunakan OOXML DrawingML Chart format.
"""

import zipfile
from xml.etree.ElementTree import Element, SubElement, tostring

# ============================================================
# DATA
# ============================================================

agustus_data = [
    ("Americano", None, 49, 735000),
    ("Kopi Susu Saring", None, 49, 686000),
    ("Manuhara", None, 27, 540000),
    ("Chocolate", None, 27, 486000),
    ("Sembadra", None, 24, 480000),
    ("Wedges Fries", None, 27, 405000),
    ("Senoaji", None, 18, 360000),
    ("Matcha Bomb", None, 15, 300000),
    ("Banana Chocolate Smoothie", None, 14, 280000),
    ("V60", None, 14, 252000),
    ("Chikin Baito", None, 14, 230000),
    ("Messy Sand", None, 10, 200000),
    ("Cappuccino", None, 10, 180000),
    ("Indomie Goreng", None, 12, 174000),
    ("Tubruk Robusta", None, 15, 150000),
    ("Hok Lo Pan", None, 7, 140000),
    ("Banana Peanut Butter Smoothie", None, 7, 140000),
    ("Nasi Mentega Cabe Garam", None, 3, 66000),
    ("Indomie Rebus", None, 3, 39000),
    ("Extra Shot/Guest Bean", None, 1, 3000),
]

september_data = [
    ("Americano", None, 47, 705000),
    ("Sembadra", None, 35, 700000),
    ("Chocolate", None, 29, 522000),
    ("Kopi Susu Saring", None, 37, 518000),
    ("Manuhara", None, 25, 500000),
    ("Nasi Mentega Cabe Garam", None, 18, 396000),
    ("V60", None, 19, 342000),
    ("Messy Sand", None, 16, 320000),
    ("Matcha Bomb", None, 15, 300000),
    ("Wedges Fries", None, 20, 300000),
    ("Banana Chocolate Smoothie", None, 13, 260000),
    ("Chikin Baito", None, 13, 211000),
    ("Senoaji", None, 9, 180000),
    ("Banana Peanut Butter Smoothie", None, 9, 180000),
    ("Indomie Goreng", None, 12, 168000),
    ("Hok Lo Pan", None, 5, 100000),
    ("Cappuccino", None, 5, 90000),
    ("Tubruk Robusta", None, 9, 90000),
    ("Indomie Rebus", None, 2, 30000),
]

oktober_data = [
    ("Sembadra", None, 45, 900000),
    ("Manuhara", None, 31, 620000),
    ("Americano", None, 37, 563000),
    ("Kopi Susu Saring", None, 37, 523000),
    ("Matcha Bomb", None, 24, 488000),
    ("V60", None, 23, 420000),
    ("Senoaji", None, 20, 400000),
    ("Chocolate", None, 22, 396000),
    ("Messy Sand", None, 19, 380000),
    ("Nasi Mentega Cabe Garam", None, 17, 374000),
    ("Banana Chocolate Smoothie", None, 16, 320000),
    ("Wedges Fries", None, 20, 303000),
    ("Chikin Baito", None, 16, 261000),
    ("Banana Peanut Butter Smoothie", None, 11, 220000),
    ("Indomie Goreng", None, 14, 201000),
    ("Tubruk Robusta", None, 17, 177000),
    ("Indomie Rebus", None, 12, 159000),
    ("Cappuccino", None, 8, 148000),
    ("Mie Svarna", None, 7, 140000),
    ("Hok Lo Pan", None, 6, 120000),
    ("Donat", None, 6, 96000),
    ("Rona Jingga", None, 5, 80000),
    ("Bakso Goreng", None, 4, 68000),
    ("Seruni", None, 2, 30000),
    ("Rona Kelabu", None, 1, 17000),
    ("Espresso", None, 1, 10000),
    ("Extra Shot/Guest Bean", None, 3, 9000),
]

def make_lookup(data):
    lookup = {}
    for name, variant, qty, revenue in data:
        if variant is None:
            lookup[name] = (qty, revenue)
    return lookup

agustus_lookup = make_lookup(agustus_data)
september_lookup = make_lookup(september_data)
oktober_lookup = make_lookup(oktober_data)

all_menus_set = set()
for d in [agustus_data, september_data, oktober_data]:
    for name, variant, qty, revenue in d:
        if variant is None:
            all_menus_set.add(name)

def total_revenue(menu_name):
    total = 0
    for lookup in [agustus_lookup, september_lookup, oktober_lookup]:
        if menu_name in lookup:
            total += lookup[menu_name][1]
    return total

all_menus = sorted(all_menus_set, key=lambda m: total_revenue(m), reverse=True)

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

# ============================================================
# CHART DATA PREPARATION
# ============================================================

# Chart 1: Bar Chart - Top 5 Menu Oktober (by revenue)
top5_okt = sorted(oktober_lookup.items(), key=lambda x: x[1][1], reverse=True)[:5]
bar_labels = [item[0] for item in top5_okt]
bar_values = [item[1][1] for item in top5_okt]

# Chart 2: Line Chart - Tren 3 bulan untuk Sembadra, Americano, Kopi Susu Saring
line_menus = ["Sembadra", "Americano", "Kopi Susu Saring"]
line_months = ["Agustus", "September", "Oktober"]
line_data = {}
for menu in line_menus:
    line_data[menu] = [
        agustus_lookup.get(menu, (0,0))[1],
        september_lookup.get(menu, (0,0))[1],
        oktober_lookup.get(menu, (0,0))[1],
    ]

# Chart 3: Pie Chart - Kontribusi Kategori Oktober
kategori_okt = {}
for name, (qty, rev) in oktober_lookup.items():
    kat = kategori_map.get(name, "Lainnya")
    kategori_okt[kat] = kategori_okt.get(kat, 0) + rev

pie_labels = list(kategori_okt.keys())
pie_values = list(kategori_okt.values())

# ============================================================
# OOXML CHART XML BUILDERS
# ============================================================

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478", "9B57A0"]

def build_bar_chart_xml():
    """Bar Chart - Top 5 Menu Oktober"""
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
            <a:r><a:rPr lang="id-ID" sz="1400" b="1"/><a:t>Top 5 Menu Oktober - Pendapatan (Rp)</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
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
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan Oktober</c:v></c:pt></c:strCache></c:strRef></c:tx>'''
    
    # Data points with colors
    for i, (label, value) in enumerate(zip(bar_labels, bar_values)):
        color = COLORS[i % len(COLORS)]
        xml += f'''
          <c:dPt><c:idx val="{i}"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr></c:dPt>'''
    
    # Category (labels)
    xml += f'''
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${len(bar_labels)+1}</c:f>
              <c:strCache>
                <c:ptCount val="{len(bar_labels)}"/>'''
    for i, label in enumerate(bar_labels):
        xml += f'''
                <c:pt idx="{i}"><c:v>{label}</c:v></c:pt>'''
    xml += '''
              </c:strCache>
            </c:strRef>
          </c:cat>'''
    
    # Values
    xml += f'''
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${len(bar_values)+1}</c:f>
              <c:numCache>
                <c:formatCode>#,##0</c:formatCode>
                <c:ptCount val="{len(bar_values)}"/>'''
    for i, value in enumerate(bar_values):
        xml += f'''
                <c:pt idx="{i}"><c:v>{value}</c:v></c:pt>'''
    xml += '''
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>'''
    return xml


def build_line_chart_xml():
    """Line Chart - Tren Pendapatan 3 Bulan"""
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
            <a:r><a:rPr lang="id-ID" sz="1400" b="1"/><a:t>Tren Pendapatan 3 Bulan - Menu Terlaris</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>'''
    
    for idx, menu in enumerate(line_menus):
        color = COLORS[idx % len(COLORS)]
        values = line_data[menu]
        xml += f'''
        <c:ser>
          <c:idx val="{idx}"/>
          <c:order val="{idx}"/>
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{menu}</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:spPr>
            <a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln>
          </c:spPr>
          <c:marker>
            <c:symbol val="circle"/>
            <c:size val="5"/>
            <c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>
          </c:marker>
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A$4</c:f>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>Agustus</c:v></c:pt>
                <c:pt idx="1"><c:v>September</c:v></c:pt>
                <c:pt idx="2"><c:v>Oktober</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B$4</c:f>
              <c:numCache>
                <c:formatCode>#,##0</c:formatCode>
                <c:ptCount val="3"/>'''
        for i, v in enumerate(values):
            xml += f'''
                <c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'''
        xml += '''
              </c:numCache>
            </c:numRef>
          </c:val>
          <c:smooth val="0"/>
        </c:ser>'''
    
    xml += '''
        <c:marker val="1"/>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="b"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>'''
    return xml


def build_pie_chart_xml():
    """Pie Chart - Kontribusi Kategori Oktober"""
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="{C_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">
  <c:chart>
    <c:title>
      <c:tx>
        <c:rich>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
            <a:r><a:rPr lang="id-ID" sz="1400" b="1"/><a:t>Kontribusi Kategori - Pendapatan Oktober</a:t></a:r>
          </a:p>
        </c:rich>
      </c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:pieChart>
        <c:varyColors val="1"/>
        <c:ser>
          <c:idx val="0"/>
          <c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan Oktober</c:v></c:pt></c:strCache></c:strRef></c:tx>'''
    
    for i in range(len(pie_labels)):
        color = COLORS[i % len(COLORS)]
        xml += f'''
          <c:dPt><c:idx val="{i}"/><c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr></c:dPt>'''
    
    # Data labels showing percentage
    xml += '''
          <c:dLbls>
            <c:showLegendKey val="0"/>
            <c:showVal val="0"/>
            <c:showCatName val="1"/>
            <c:showSerName val="0"/>
            <c:showPercent val="1"/>
          </c:dLbls>'''
    
    xml += f'''
          <c:cat>
            <c:strRef>
              <c:f>Sheet1!$A$2:$A${len(pie_labels)+1}</c:f>
              <c:strCache>
                <c:ptCount val="{len(pie_labels)}"/>'''
    for i, label in enumerate(pie_labels):
        xml += f'''
                <c:pt idx="{i}"><c:v>{label}</c:v></c:pt>'''
    xml += '''
              </c:strCache>
            </c:strRef>
          </c:cat>'''
    
    xml += f'''
          <c:val>
            <c:numRef>
              <c:f>Sheet1!$B$2:$B${len(pie_values)+1}</c:f>
              <c:numCache>
                <c:formatCode>#,##0</c:formatCode>
                <c:ptCount val="{len(pie_values)}"/>'''
    for i, value in enumerate(pie_values):
        xml += f'''
                <c:pt idx="{i}"><c:v>{value}</c:v></c:pt>'''
    xml += '''
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    <c:legend>
      <c:legendPos val="r"/>
      <c:overlay val="0"/>
    </c:legend>
    <c:plotVisOnly val="1"/>
  </c:chart>
</c:chartSpace>'''
    return xml


# ============================================================
# DRAWING XML (connects chart to sheet)
# ============================================================

DR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

def build_drawing_xml(chart_count):
    """Create drawing XML with anchors for each chart."""
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="{DR_NS}" xmlns:a="{A_NS}" xmlns:r="{R_NS}">'''
    
    for i in range(chart_count):
        # Each chart is ~15 columns wide, ~15 rows tall
        # Stacked vertically: chart0 at row 0, chart1 at row 18, chart2 at row 36
        from_row = i * 20
        to_row = from_row + 18
        
        xml += f'''
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>{from_row}</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>{to_row}</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="{i+2}" name="Chart {i+1}"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="{C_NS}" r:id="rId{i+1}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>'''
    
    xml += '''
</xdr:wsDr>'''
    return xml


def build_drawing_rels_xml(chart_count):
    xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'''
    for i in range(chart_count):
        xml += f'''
  <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{i+1}.xml"/>'''
    xml += '''
</Relationships>'''
    return xml


# ============================================================
# XLSX BUILDER (with charts)
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

    # Sheet 1 headers
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

    # ---- Build Sheet 1 rows ----
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

    # ---- Build Sheet 2 rows ----
    sheet2_rows = []
    sheet2_rows.append(("header", headers2))
    total_agst = total_sept = total_okt = 0
    for menu in all_menus:
        ar = agustus_lookup.get(menu, (0, 0))[1]
        sr = september_lookup.get(menu, (0, 0))[1]
        orev = oktober_lookup.get(menu, (0, 0))[1]
        total_3 = ar + sr + orev
        total_agst += ar; total_sept += sr; total_okt += orev
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

    def build_sheet_xml(rows, col_widths, drawing_rid=None):
        lines = []
        lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        lines.append(f'<worksheet xmlns="{NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
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
                    else:
                        lines.append(f'<c r="{cell_ref}" s="3"><v>{val}</v></c>')
            lines.append('</row>')
        lines.append('</sheetData>')
        if drawing_rid:
            lines.append(f'<drawing r:id="{drawing_rid}"/>')
        lines.append('</worksheet>')
        return '\n'.join(lines)

    col_widths_1 = [5, 35, 12, 18, 18, 18, 18, 18, 18, 22, 22, 22]
    col_widths_2 = [35, 28, 28, 28, 22]

    # Sheet 3 is the Charts sheet (empty data, just drawing)
    sheet3_xml = build_sheet_xml([], [15]*10, drawing_rid="rId1")
    sheet1_xml = build_sheet_xml(sheet1_rows, col_widths_1)
    sheet2_xml = build_sheet_xml(sheet2_rows, col_widths_2)

    # ---- Build all XML parts ----
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Data Penjualan" sheetId="1" r:id="rId1"/>
    <sheet name="Analisis" sheetId="2" r:id="rId2"/>
    <sheet name="Grafik" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>'''

    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>'''

    sheet3_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>'''

    drawing_xml = build_drawing_xml(3)
    drawing_rels = build_drawing_rels_xml(3)

    # Write xlsx
    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('xl/workbook.xml', workbook_xml)
        zf.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        zf.writestr('xl/styles.xml', build_styles_xml())
        zf.writestr('xl/sharedStrings.xml', build_shared_strings_xml())
        zf.writestr('xl/worksheets/sheet1.xml', sheet1_xml)
        zf.writestr('xl/worksheets/sheet2.xml', sheet2_xml)
        zf.writestr('xl/worksheets/sheet3.xml', sheet3_xml)
        zf.writestr('xl/worksheets/_rels/sheet3.xml.rels', sheet3_rels)
        zf.writestr('xl/drawings/drawing1.xml', drawing_xml)
        zf.writestr('xl/drawings/_rels/drawing1.xml.rels', drawing_rels)
        zf.writestr('xl/charts/chart1.xml', build_bar_chart_xml())
        zf.writestr('xl/charts/chart2.xml', build_line_chart_xml())
        zf.writestr('xl/charts/chart3.xml', build_pie_chart_xml())

    print(f"✅ File '{filename}' berhasil dibuat dengan 3 CHARTS!")
    print(f"\n📊 CHART 1 - Bar Chart: Top 5 Menu Oktober")
    for label, val in zip(bar_labels, bar_values):
        print(f"   {label}: Rp {val:,}".replace(",","."))
    print(f"\n📈 CHART 2 - Line Chart: Tren 3 Bulan")
    for menu in line_menus:
        vals = line_data[menu]
        print(f"   {menu}: {' → '.join(f'Rp{v:,}'.replace(',','.') for v in vals)}")
    print(f"\n🥧 CHART 3 - Pie Chart: Kategori Oktober")
    total_okt_all = sum(pie_values)
    for label, val in zip(pie_labels, pie_values):
        pct = val / total_okt_all * 100
        print(f"   {label}: Rp {val:,} ({pct:.1f}%)".replace(",","."))
    
    print(f"\n💰 Total Omzet: Agst Rp{total_agst:,} | Sept Rp{total_sept:,} | Okt Rp{total_okt:,} | TOTAL Rp{grand_total:,}".replace(",","."))

if __name__ == "__main__":
    make_xlsx("/projects/sandbox/Data_Penjualan_Semat_Space.xlsx")
