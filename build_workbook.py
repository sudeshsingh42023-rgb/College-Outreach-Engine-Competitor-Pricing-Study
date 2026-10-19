"""Builds the outreach + competitor workbook. Run: python build_workbook.py"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

F = "Arial"
HDR = PatternFill("solid", fgColor="1F3864"); INP = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF"); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
hfont = Font(name=F, bold=True, color="FFFFFF", size=10)
base = Font(name=F, size=10); blue = Font(name=F, size=10, color="0000FF"); bold = Font(name=F, size=10, bold=True)

def header(ws, row, cols):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c); x.font, x.fill, x.border = hfont, HDR, BOX
        x.alignment = Alignment(wrap_text=True, vertical="center")

def widths(ws, w):
    for i, v in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = v

wb = Workbook()

# ---------------- Legend ----------------
ws = wb.active; ws.title = "Legend"
rows = [
 ("College Outreach Engine & Competitor Pricing Study", True),
 ("", False),
 ("Yellow cells with blue text = inputs you edit. Black cells = formulas (do not overwrite).", False),
 ("Row 5 of Competitor_Benchmark (Asyvyas) comes from asyvyas.in; fill every other row from each company's own public site and cite the URL.", False),
 ("Example rows in College_Pipeline and Outreach_Log are placeholders - delete or overwrite them.", False),
 ("Tier-scoring weights are on the 'Assumptions' sheet. Change them there and every priority score updates.", False),
 ("Only report response rates from real outreach. Do not present the sample CSV in data/ as real results.", False),
]
for i, (t, b) in enumerate(rows, 1):
    c = ws.cell(row=i, column=1, value=t); c.font = Font(name=F, size=14 if i == 1 else 10, bold=b)
ws.column_dimensions["A"].width = 130

# ---------------- Assumptions ----------------
a = wb.create_sheet("Assumptions")
header(a, 1, ["Parameter", "Value", "Note"])
params = [
 ("Tier-1 weight", 3, "Higher = more attractive prospect (placement-cell reach, brand)"),
 ("Tier-2 weight", 2, ""),
 ("Tier-3 weight", 1, ""),
 ("Strength weight per 1,000 students", 1, "Larger colleges = more potential enrolments"),
 ("Warm-contact bonus", 2, "Added when we already know someone at the college"),
 ("Course price incl. GST (INR)", 2949, "Source: asyvyas.in pricing (INR 2,499 + 18% GST)"),
 ("Assumed enrolment rate per college pitch (%)", 5, "ASSUMPTION - edit once you have real data"),
 ("Assumed students enrolled per converted college", 20, "ASSUMPTION - edit once you have real data"),
]
for r, (p, v, n) in enumerate(params, 2):
    a.cell(row=r, column=1, value=p).font = base
    c = a.cell(row=r, column=2, value=v); c.font, c.fill = blue, INP
    a.cell(row=r, column=3, value=n).font = base
widths(a, [48, 14, 70])

# ---------------- Competitor_Benchmark ----------------
cb = wb.create_sheet("Competitor_Benchmark")
cols = ["Company", "Primary audience", "Example course", "Price (INR, incl. GST)", "Delivery format", "Certification",
        "Placement claim (as stated)", "Claim verifiable? (Y/N)", "Free trial / demo?", "Source URL", "Date checked"]
header(cb, 1, cols)
cb.cell(row=2, column=1, value="Fill one row per competitor from its own public pages. Leave unknown cells blank, never guess.").font = Font(name=F, italic=True, size=9)
cb.row_dimensions[1].height = 32
names = ["Asyvyas Solutions", "NxtWave", "Scaler", "upGrad", "Coursera", "Simplilearn", "PW Skills", "Internshala Trainings", "LearnVern / other local"]
for i, n in enumerate(names):
    r = 4 + i
    cb.cell(row=r, column=1, value=n).font = bold
    for c in range(2, 12):
        x = cb.cell(row=r, column=c); x.font, x.fill, x.border = blue, INP, BOX
# Asyvyas row from public website
asy = {2: "College students / freshers", 3: "AI & ML, Full Stack, Data Analytics, Cloud, VLSI, Core Engineering (16 tracks listed)",
       4: 2949, 5: "100% live online classes", 6: "ISO-certified (as stated on site)",
       7: "303,154+ placed students; 4,870+ partner companies (site claim)", 8: "N", 9: "", 10: "https://asyvyas.in/", 11: "19-Oct-2026"}
for c, v in asy.items(): cb.cell(row=4, column=c, value=v)
cb["D4"].number_format = '#,##0'
dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True); cb.add_data_validation(dv); dv.add("H4:H12")
# summary formulas
cb["A15"] = "Median price of filled competitors (excl. Asyvyas)"; cb["A15"].font = bold
cb["D15"] = '=IF(COUNT(D5:D12)=0,"",MEDIAN(D5:D12))'
cb["A16"] = "Asyvyas price vs competitor median"; cb["A16"].font = bold
cb["D16"] = '=IF(OR(D15="",D4=""),"",D4/D15-1)'; cb["D16"].number_format = "0.0%"
cb["A17"] = "Competitors with unverifiable placement claims"; cb["A17"].font = bold
cb["D17"] = '=COUNTIF(H5:H12,"N")'
for r in (15, 16, 17): cb.cell(row=r, column=4).font = base
widths(cb, [26, 24, 34, 16, 22, 22, 34, 14, 14, 30, 13])
for r in range(4, 13):
    for c in (3, 7): cb.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------- College_Pipeline ----------------
cp = wb.create_sheet("College_Pipeline")
cols = ["#", "College / institute", "City", "State", "Tier (1/2/3)", "Student strength", "Warm contact? (Y/N)",
        "T&P contact name", "Email / LinkedIn", "Stage", "Last touch date", "Priority score", "Priority rank", "Notes"]
header(cp, 1, cols); cp.row_dimensions[1].height = 32
example = [1, "EXAMPLE - replace with a real college", "Bengaluru", "Karnataka", 2, 3000, "N", "Name of T&P officer", "name@college.edu", "Not contacted", None, None, None, "Example row"]
N = 100
for r in range(2, N + 2):
    cp.cell(row=r, column=1, value=r - 1).font = base
    for c in range(2, 15):
        if c in (12, 13): continue
        x = cp.cell(row=r, column=c); x.font, x.fill, x.border = blue, INP, BOX
    cp.cell(row=r, column=12, value=(
        f'=IF(OR(B{r}="",E{r}=""),"",'
        f'INDEX(Assumptions!$B$2:$B$4,E{r})+Assumptions!$B$5*F{r}/1000+IF(G{r}="Y",Assumptions!$B$6,0))')).font = base
    cp.cell(row=r, column=12).number_format = "0.0"
    cp.cell(row=r, column=13, value=f'=IF(L{r}="","",RANK(L{r},$L$2:$L${N+1}))').font = base
for c, v in enumerate(example, 1):
    if c in (12, 13): continue
    if v is not None: cp.cell(row=2, column=c, value=v)
cp["K2"] = None
dv1 = DataValidation(type="list", formula1='"1,2,3"', allow_blank=True); dv2 = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
dv3 = DataValidation(type="list", formula1='"Not contacted,Touch 1 sent,Touch 2 sent,Touch 3 sent,Replied,Meeting booked,Pitched,Pilot / drive agreed,Closed - lost"', allow_blank=True)
for d, rng in ((dv1, f"E2:E{N+1}"), (dv2, f"G2:G{N+1}"), (dv3, f"J2:J{N+1}")): cp.add_data_validation(d); d.add(rng)
cp.freeze_panes = "C2"
widths(cp, [5, 36, 14, 16, 11, 12, 12, 22, 28, 20, 14, 12, 12, 30])

# pipeline summary
ps = wb.create_sheet("Pipeline_Summary")
header(ps, 1, ["Stage", "Colleges", "% of contacted"])
stages = ["Not contacted", "Touch 1 sent", "Touch 2 sent", "Touch 3 sent", "Replied", "Meeting booked", "Pitched", "Pilot / drive agreed", "Closed - lost"]
for i, s in enumerate(stages, 2):
    ps.cell(row=i, column=1, value=s).font = base
    ps.cell(row=i, column=2, value=f'=COUNTIF(College_Pipeline!$J$2:$J${N+1},A{i})').font = base
    ps.cell(row=i, column=3, value=f'=IF(SUM($B$3:$B${len(stages)+1})=0,"",B{i}/SUM($B$3:$B${len(stages)+1}))' if s != "Not contacted" else None).font = base
    ps.cell(row=i, column=3).number_format = "0.0%"
r = len(stages) + 3
ps.cell(row=r, column=1, value="Total colleges listed").font = bold
ps.cell(row=r, column=2, value=f'=COUNTA(College_Pipeline!$B$2:$B${N+1})').font = base
ps.cell(row=r+1, column=1, value="Reply rate (replied or beyond / contacted)").font = bold
ps.cell(row=r+1, column=2, value=f'=IF(SUM(B3:B10)=0,"",SUM(B6:B9)/SUM(B3:B10))').font = base; ps.cell(row=r+1, column=2).number_format = "0.0%"
ps.cell(row=r+3, column=1, value="Illustrative revenue potential (uses Assumptions sheet - ASSUMPTIONS, not forecasts)").font = bold
ps.cell(row=r+4, column=1, value="Expected converted colleges").font = base
ps.cell(row=r+4, column=2, value=f'=B{r}*Assumptions!B8/100').font = base; ps.cell(row=r+4, column=2).number_format = "0.0"
ps.cell(row=r+5, column=1, value="Expected enrolments").font = base
ps.cell(row=r+5, column=2, value=f'=B{r+4}*Assumptions!B9').font = base; ps.cell(row=r+5, column=2).number_format = "#,##0"
ps.cell(row=r+6, column=1, value="Expected revenue (INR)").font = base
ps.cell(row=r+6, column=2, value=f'=B{r+5}*Assumptions!B7').font = base; ps.cell(row=r+6, column=2).number_format = "#,##0"
widths(ps, [58, 14, 16])

# ---------------- Outreach_Log ----------------
ol = wb.create_sheet("Outreach_Log")
header(ol, 1, ["Send #", "College", "Contact", "Channel (Email/LinkedIn/WhatsApp)", "Variant (A/B)", "Date sent", "Replied (Y/N)", "Meeting booked (Y/N)", "Notes"])
ol.row_dimensions[1].height = 32
for r in range(2, 302):
    ol.cell(row=r, column=1, value=r - 1).font = base
    for c in range(2, 10):
        x = ol.cell(row=r, column=c); x.font, x.fill, x.border = blue, INP, BOX
ex = ["EXAMPLE - replace with a real college", "Name of T&P officer", "Email", "A", None, "N", "N", "Example row"]
for c, v in enumerate(ex, 2):
    if v is not None: ol.cell(row=2, column=c, value=v)
dvv = DataValidation(type="list", formula1='"A,B"', allow_blank=True); dvy = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
for d, rg in ((dvv, "E2:E301"), (dvy, "G2:H301")): ol.add_data_validation(d); d.add(rg)
ol.freeze_panes = "A2"; widths(ol, [8, 34, 24, 22, 12, 13, 12, 14, 30])

# ---------------- AB_Summary ----------------
ab = wb.create_sheet("AB_Summary")
header(ab, 1, ["Metric", "Variant A", "Variant B"])
L = "Outreach_Log!"
ab["A2"] = "Messages sent"
ab["B2"] = f'=COUNTIFS({L}$E$2:$E$301,"A",{L}$B$2:$B$301,"<>")'; ab["C2"] = f'=COUNTIFS({L}$E$2:$E$301,"B",{L}$B$2:$B$301,"<>")'
ab["A3"] = "Replies"
ab["B3"] = f'=COUNTIFS({L}$E$2:$E$301,"A",{L}$G$2:$G$301,"Y")'; ab["C3"] = f'=COUNTIFS({L}$E$2:$E$301,"B",{L}$G$2:$G$301,"Y")'
ab["A4"] = "Reply rate"; ab["B4"] = '=IF(B2=0,"",B3/B2)'; ab["C4"] = '=IF(C2=0,"",C3/C2)'
ab["A5"] = "Meetings booked"
ab["B5"] = f'=COUNTIFS({L}$E$2:$E$301,"A",{L}$H$2:$H$301,"Y")'; ab["C5"] = f'=COUNTIFS({L}$E$2:$E$301,"B",{L}$H$2:$H$301,"Y")'
ab["A6"] = "Meeting rate"; ab["B6"] = '=IF(B2=0,"",B5/B2)'; ab["C6"] = '=IF(C2=0,"",C5/C2)'
ab["A8"] = "Two-proportion z-test on reply rate (A vs B)"; ab["A8"].font = bold
ab["A9"] = "Pooled reply rate"; ab["B9"] = '=IF(B2+C2=0,"",(B3+C3)/(B2+C2))'
ab["A10"] = "z statistic"; ab["B10"] = '=IF(OR(B2=0,C2=0,B9=""),"",IF(OR(B9=0,B9=1),"",(B4-C4)/SQRT(B9*(1-B9)*(1/B2+1/C2))))'
ab["A11"] = "p-value (two-sided)"; ab["B11"] = '=IF(B10="","",2*(1-NORMSDIST(ABS(B10))))'
ab["A12"] = "Read-out"; ab["B12"] = '=IF(B11="","Need sends in both variants",IF(B11<0.05,"Statistically significant at 5%","Not significant - more sends needed"))'
for r in range(2, 13):
    for c in (1, 2, 3):
        cell = ab.cell(row=r, column=c)
        if cell.font != hfont: cell.font = bold if (c == 1 and r in (8,)) else base
for ref in ("B4", "C4", "B6", "C6", "B9"): ab[ref].number_format = "0.0%"
ab["B10"].number_format = "0.00"; ab["B11"].number_format = "0.000"
ab["A14"] = "Note: with fewer than ~30 sends per variant the test has very little power; report this honestly."
ab["A14"].font = Font(name=F, italic=True, size=9)
widths(ab, [48, 30, 14])

wb.save("asyvyas_outreach_workbook.xlsx"); print("saved")
