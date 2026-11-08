#!/usr/bin/env python3
"""
V4 - Excel Data Penjualan Semat Space
Charts embedded in Sheet "Grafik" using strict OOXML compliance.
Fixed: proper chart namespace declarations, valid axis IDs, 
correct drawing anchors with EMU offsets.
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

# Chart data
top5 = sorted(oktober.items(), key=lambda x: x[1][1], reverse=True)[:5]
bar_names = [t[0] for t in top5]
bar_vals  = [t[1][1] for t in top5]

line_menus = ["Sembadra", "Americano", "Kopi Susu Saring"]
line_vals = {m: [agustus.get(m,(0,0))[1], september.get(m,(0,0))[1], oktober.get(m,(0,0))[1]] for m in line_menus}

kat_okt = {}
for name, (q, r) in oktober.items():
    k = kategori_map.get(name, "Lainnya")
    kat_okt[k] = kat_okt.get(k, 0) + r
pie_names = list(kat_okt.keys())
pie_vals  = list(kat_okt.values())

COLORS = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]

# Shared strings
_ss = []
_si = {}
def ss(t):
    t = str(t)
    if t not in _si:
        _si[t] = len(_ss)
        _ss.append(t)
    return _si[t]

def CL(c):
    r = ""; c += 1
    while c > 0: c -= 1; r = chr(65 + c%26) + r; c //= 26
    return r

def esc(s):
    return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("'","&apos;").replace('"',"&quot;")

# ============================================================
# SHEET BUILDERS
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
    ta=ts=to=0
    for ri,m in enumerate(all_menus):
        r = ri+2; s1r = ri+2
        ar = agustus.get(m,(0,0))[1]; sr = september.get(m,(0,0))[1]; ov = oktober.get(m,(0,0))[1]
        ta+=ar; ts+=sr; to+=ov
        L.append(f'<row r="{r}">')
        L.append(f'<c r="A{r}" s="2" t="s"><v>{ss(m)}</v></c>')
        L.append(f'<c r="B{r}" s="3"><f>\'Data Penjualan\'!J{s1r}</f><v>{ar}</v></c>')
        L.append(f'<c r="C{r}" s="3"><f>\'Data Penjualan\'!K{s1r}</f><v>{sr}</v></c>')
        L.append(f'<c r="D{r}" s="3"><f>\'Data Penjualan\'!L{s1r}</f><v>{ov}</v></c>')
        L.append(f'<c r="E{r}" s="3"><f>B{r}+C{r}+D{r}</f><v>{ar+sr+ov}</v></c>')
        L.append('</row>')
    tr = N+2; lr = N+1; gt = ta+ts+to
    L.append(f'<row r="{tr}">')
    L.append(f'<c r="A{tr}" s="5" t="s"><v>{ss("TOTAL OMZET")}</v></c>')
    L.append(f'<c r="B{tr}" s="4"><f>SUM(B2:B{lr})</f><v>{ta}</v></c>')
    L.append(f'<c r="C{tr}" s="4"><f>SUM(C2:C{lr})</f><v>{ts}</v></c>')
    L.append(f'<c r="D{tr}" s="4"><f>SUM(D2:D{lr})</f><v>{to}</v></c>')
    L.append(f'<c r="E{tr}" s="4"><f>SUM(E2:E{lr})</f><v>{gt}</v></c>')
    L.append('</row>')
    L.append('</sheetData></worksheet>')
    return '\n'.join(L)

def sheet3():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData/>
<drawing r:id="rId1"/>
</worksheet>'''

# ============================================================
# CHARTS using strLit / numLit (inline data, no cell references)
# ============================================================

def mk_bar():
    o = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>
        <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0"/><a:t>Top 5 Menu Oktober - Pendapatan (Rp)</a:t></a:r>
      </a:p></c:rich></c:tx>
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
          <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan</c:v></c:pt></c:strLit></c:tx>
'''
    for i in range(5):
        o += f'          <c:dPt><c:idx val="{i}"/><c:invertIfNegative val="0"/><c:spPr><a:solidFill><a:srgbClr val="{COLORS[i]}"/></a:solidFill></c:spPr></c:dPt>\n'
    o += '          <c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>\n'
    o += f'          <c:cat><c:strLit><c:ptCount val="{len(bar_names)}"/>\n'
    for i,n in enumerate(bar_names):
        o += f'            <c:pt idx="{i}"><c:v>{esc(n)}</c:v></c:pt>\n'
    o += '          </c:strLit></c:cat>\n'
    o += f'          <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="{len(bar_vals)}"/>\n'
    for i,v in enumerate(bar_vals):
        o += f'            <c:pt idx="{i}"><c:v>{v}</c:v></c:pt>\n'
    o += '''          </c:numLit></c:val>
        </c:ser>
        <c:axId val="100"/>
        <c:axId val="200"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="100"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="200"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:crossAx val="100"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>'''
    return o

def mk_line():
    o = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>
        <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0"/><a:t>Tren Pendapatan 3 Bulan</a:t></a:r>
      </a:p></c:rich></c:tx>
      <c:overlay val="0"/>
    </c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:layout/>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:varyColors val="0"/>
'''
    for si, menu in enumerate(line_menus):
        vals = line_vals[menu]
        color = COLORS[si]
        o += f'''        <c:ser>
          <c:idx val="{si}"/>
          <c:order val="{si}"/>
          <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>{esc(menu)}</c:v></c:pt></c:strLit></c:tx>
          <c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>
          <c:marker>
            <c:symbol val="circle"/>
            <c:size val="7"/>
            <c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>
          </c:marker>
          <c:cat><c:strLit><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>Agustus</c:v></c:pt>
            <c:pt idx="1"><c:v>September</c:v></c:pt>
            <c:pt idx="2"><c:v>Oktober</c:v></c:pt>
          </c:strLit></c:cat>
          <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="3"/>
            <c:pt idx="0"><c:v>{vals[0]}</c:v></c:pt>
            <c:pt idx="1"><c:v>{vals[1]}</c:v></c:pt>
            <c:pt idx="2"><c:v>{vals[2]}</c:v></c:pt>
          </c:numLit></c:val>
          <c:smooth val="0"/>
        </c:ser>
'''
    o += '''        <c:marker val="1"/>
        <c:axId val="300"/>
        <c:axId val="400"/>
      </c:lineChart>
      <c:catAx>
        <c:axId val="300"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:crossAx val="400"/>
        <c:crosses val="autoZero"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="400"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="#,##0" sourceLinked="0"/>
        <c:crossAx val="300"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>'''
    return o

def mk_pie():
    o = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>
        <a:pPr><a:defRPr sz="1400" b="1"/></a:pPr>
        <a:r><a:rPr lang="en-US" sz="1400" b="1" dirty="0"/><a:t>Kontribusi Kategori - Oktober</a:t></a:r>
      </a:p></c:rich></c:tx>
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
          <c:tx><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Pendapatan</c:v></c:pt></c:strLit></c:tx>
'''
    for i in range(len(pie_names)):
        o += f'          <c:dPt><c:idx val="{i}"/><c:spPr><a:solidFill><a:srgbClr val="{COLORS[i]}"/></a:solidFill></c:spPr></c:dPt>\n'
    o += '          <c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="1"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/></c:dLbls>\n'
    o += f'          <c:cat><c:strLit><c:ptCount val="{len(pie_names)}"/>\n'
    for i,n in enumerate(pie_names):
        o += f'            <c:pt idx="{i}"><c:v>{esc(n)}</c:v></c:pt>\n'
    o += '          </c:strLit></c:cat>\n'
    o += f'          <c:val><c:numLit><c:formatCode>#,##0</c:formatCode><c:ptCount val="{len(pie_vals)}"/>\n'
    for i,v in enumerate(pie_vals):
        o += f'            <c:pt idx="{i}"><c:v>{v}</c:v></c:pt>\n'
    o += '''          </c:numLit></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
  </c:chart>
</c:chartSpace>'''
    return o

# ============================================================
# DRAWING - using oneCellAnchor with explicit EMU sizes
# ============================================================

def drawing():
    # Each chart ~ 9144000 EMU wide (24cm), 5486400 EMU tall (14.4cm)
    W = 9144000
    H = 5486400
    GAP = 457200  # ~1.2cm gap between charts
    
    o = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
'''
    for i in range(3):
        y_off = i * (H + GAP)
        o += f'''  <xdr:absoluteAnchor>
    <xdr:pos x="0" y="{y_off}"/>
    <xdr:ext cx="{W}" cy="{H}"/>
    <xdr:graphicFrame macro="">
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="{i+2}" name="Chart {i+1}"/>
        <xdr:cNvGraphicFramePr>
          <a:graphicFrameLocks noGrp="1"/>
        </xdr:cNvGraphicFramePr>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="{W}" cy="{H}"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId{i+1}"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:absoluteAnchor>
'''
    o += '</xdr:wsDr>'
    return o

# ============================================================
# PACKAGE
# ============================================================

def build():
    fname = "/projects/sandbox/Data_Penjualan_Semat_Space.xlsx"
    
    s1 = sheet1()
    s2 = sheet2()
    s3 = sheet3()
    
    ss_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    ss_xml += f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(_ss)}" uniqueCount="{len(_ss)}">\n'
    for s in _ss:
        ss_xml += f'<si><t xml:space="preserve">{esc(s)}</t></si>\n'
    ss_xml += '</sst>'

    with zipfile.ZipFile(fname, 'w', zipfile.ZIP_DEFLATED) as z:
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
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/charts/chart3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
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
        
        z.writestr('xl/worksheets/_rels/sheet3.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>''')
        
        z.writestr('xl/drawings/drawing1.xml', drawing())
        
        z.writestr('xl/drawings/_rels/drawing1.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
</Relationships>''')
        
        z.writestr('xl/charts/chart1.xml', mk_bar())
        z.writestr('xl/charts/chart2.xml', mk_line())
        z.writestr('xl/charts/chart3.xml', mk_pie())
    
    print(f"✅ {fname}")
    print("   Sheet 1: Data Penjualan (with formulas D*E, F*G, H*I)")
    print("   Sheet 2: Analisis (with SUM & cross-sheet formulas)")
    print("   Sheet 3: Grafik (Bar + Line + Pie charts)")

if __name__ == "__main__":
    build()
