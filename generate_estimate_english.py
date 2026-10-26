#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SRayhan — Fully ENGLISH 4-sheet estimate workbook
Client: Mr. Saiful Islam Vai, Brahmanbaria | 3-storied residential building, 2251 sft/floor (6,753 sft total)
Sheets: 1-Concrete Estimate | 2-Rebar Estimate | 3-Finishing Estimate | 4-Summary & BOQ
All figures identical to the previously approved (Bangla/English) version — text translated to English only.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- Palette ----------
NAVY = "0F2C4D"
GOLD = "B78A3B"
GREEN = "0E7A5C"
GRAY_DARK = "333333"
GRAY_MID = "6B7280"
WHITE = "FFFFFF"
NAVY_LIGHT = "EAF0F8"
GOLD_LIGHT = "FFFBEB"
GREEN_LIGHT = "ECFDF5"
GRAY_LIGHT = "F9FAFB"
AMBER_TXT = "92400E"

thin = Border(*(Side(style="thin", color="D1D5DB") for _ in range(4)))

def hdr(cell, bg=NAVY, size=8):
    cell.fill = PatternFill("solid", start_color=bg, end_color=bg)
    cell.font = Font(name="Calibri", size=size, color=WHITE, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin

def cell(cell, size=8, bold=False, bg=None, color="1F2937", align="center", fmt=None):
    cell.font = Font(name="Calibri", size=size, color=color, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if bg:
        cell.fill = PatternFill("solid", start_color=bg, end_color=bg)
    cell.border = thin
    if fmt:
        cell.number_format = fmt

def banner(ws, rng, text, size=13, color=NAVY, bg=None, fg=None, italic=False, bold=True, height=22):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = text
    c.font = Font(name="Calibri", size=size, color=fg or color, bold=bold, italic=italic)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if bg:
        first, last = rng.split(":")
        c.fill = PatternFill("solid", start_color=bg, end_color=bg)
        from openpyxl.utils.cell import range_boundaries
        min_c, min_r, max_c, max_r = range_boundaries(rng)
        for rr in range(min_r, max_r + 1):
            for cc in range(min_c, max_c + 1):
                ws.cell(row=rr, column=cc).fill = PatternFill("solid", start_color=bg, end_color=bg)
                ws.cell(row=rr, column=cc).border = thin
    ws.row_dimensions[c.row].height = height

def cards(ws, row_t, row_v, items, positions):
    """items = [(title, value, accent_color)] ; positions = [(c1,c2), ...] 1-based col ranges"""
    bgs = [NAVY_LIGHT, GOLD_LIGHT, GREEN_LIGHT, "F3F4F6"]
    for i, (title, val, color) in enumerate(items):
        s, e = positions[i]
        a, b = get_column_letter(s), get_column_letter(e)
        ws.merge_cells(f"{a}{row_t}:{b}{row_t}")
        ws.merge_cells(f"{a}{row_v}:{b}{row_v}")
        c1 = ws[f"{a}{row_t}"]
        c1.value = title
        c1.font = Font(name="Calibri", size=7.5, color=WHITE, bold=True)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.fill = PatternFill("solid", start_color=color, end_color=color)
        c1.border = thin
        for cc in range(s + 1, e + 1):
            x = ws.cell(row=row_t, column=cc); x.fill = PatternFill("solid", start_color=color, end_color=color); x.border = thin
        c2 = ws[f"{a}{row_v}"]
        c2.value = val
        c2.font = Font(name="Calibri", size=9, color=color if color != "374151" else "111827", bold=True)
        c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c2.fill = PatternFill("solid", start_color=bgs[i], end_color=bgs[i])
        c2.border = thin
        for cc in range(s + 1, e + 1):
            x = ws.cell(row=row_v, column=cc); x.fill = PatternFill("solid", start_color=bgs[i], end_color=bgs[i]); x.border = thin
    ws.row_dimensions[row_t].height = 14
    ws.row_dimensions[row_v].height = 20

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

wb = openpyxl.Workbook()
TODAY = "26 Oct 2026"
CLIENT_LINE = "CLIENT: Mr. Saiful Islam Vai   |   LOCATION: Brahmanbaria   |   PROJECT: 3-Storied Residential Building   |   DATE: 26 Oct 2026   |   REF: SR/EST/2026-010"

# =========================================================
# SHEET 1 — CONCRETE ESTIMATE
# =========================================================
ws = wb.active
ws.title = "1-Concrete Estimate"
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.print_options.horizontalCentered = True
ws.sheet_view.showGridLines = False
set_widths(ws, [6, 34, 8, 8, 8, 8, 13, 11, 11, 11, 11])

banner(ws, "A1:K1", "SRAYHAN ENGINEERING & ARCHITECTURE   •   BRAHMANBARIA / DHAKA   •   STRUCTURE  |  ARCHITECTURE  |  ESTIMATION",
       size=7, bg=NAVY, fg=GOLD, height=18)
banner(ws, "A2:K2", "CONCRETE QUANTITY ESTIMATE  —  3-STORIED RESIDENTIAL BUILDING")
banner(ws, "A3:K3", "Project: Mr. Saiful Islam Vai   |   Location: Brahmanbaria   |   2,251 sft per floor   |   Total 6,753 sft (3 floors)   |   Date: 26 Oct 2026",
       size=9, color=GRAY_MID, bold=False, height=14)

cards(ws, 4, 5, [
    ("TOTAL CONCRETE", "7,958 cft\n225.37 Cum", NAVY),
    ("FOUNDATION DEPTH", '6\'-0" (H)\nEarth cutting', GOLD),
    ("CONCRETE GRADE", "f'c = 3000 psi\nMix 1:1.5:3", GREEN),
    ("CEMENT REQUIRED", "\u2248 1,690 bags\n(50 kg) net", "374151"),
], [(1, 3), (4, 6), (7, 9), (10, 11)])

banner(ws, "A6:K6",
       "Source: Footing Schedule (F-1 to F-4), Column Schedule (C-1 to C-4), Column Layout Plan (18 columns), Ground Floor Plan 2,251 sft   •   "
       'Assumptions: Floor height 10\'-0", Slab 5", Beam 12"x18", Grade Beam 12"x18"',
       size=7.5, color=GRAY_MID, italic=True, bg=GRAY_LIGHT, bold=False, height=13)

headers = ["Sl.", "Description", "Nos", "L (ft)", "B (ft)", "D/H (ft)", "Quantity\n(cft)", "Cum\n(m\u00b3)", "Cement\n(bags)", "Sand\n(cft)", "Chips\n(cft)"]
ws.row_dimensions[7].height = 28
for i, h in enumerate(headers, 1):
    hdr(ws.cell(row=7, column=i, value=h))

DATA1 = [
    ("S", "A — FOUNDATION CONCRETE  (FOOTINGS)", None),
    (1, "F-1 Footing  6'-0\" x 7'-0\" x 2'-0\"  (D=24\")  — 4 nos", 4, 6, 7, 2.00, 336, 9.51, 72, 143, 276),
    (2, "F-2 Footing  6'-0\" x 6'-0\" x 1'-8\"  (D=20\")  — 10 nos", 10, 6, 6, 1.67, 600, 16.99, 128, 255, 494),
    (3, "F-3 Footing  5'-0\" x 5'-0\" x 1'-6\"  (D=18\")  — 2 nos", 2, 5, 5, 1.50, 75, 2.12, 16, 32, 62),
    (4, "F-4 Footing  4'-0\" x 4'-0\" x 1'-3\"  (D=15\")  — 2 nos", 2, 4, 4, 1.25, 40, 1.13, 9, 17, 33),
    ("T", "Subtotal — Footings \u2192", 1051, 29.76, 225, 446, 864),
    ("S", "B — GRADE BEAM  (12\" x 18\")", None),
    (5, "Grade Beam — longitudinal + transverse  (430 rft)", 1, 430, 1.00, 1.50, 645, 18.27, 138, 274, 531),
    ("T", "Subtotal — Grade Beam \u2192", 645, 18.27, 138, 274, 531),
    ("S", 'C — COLUMN CONCRETE  (12" x 12" to 12" x 18")', None),
    (6, "Columns below Grade Beam  (pedestal 5'-0\" high — 18 nos)", 18, "", "", 5.0, 113, 3.19, 24, 48, 93),
    (7, "Ground Floor columns  (10'-0\" high — C-1:4, C-2:10, C-3:2, C-4:2)", 18, "", "", 10.0, 225, 6.37, 48, 96, 185),
    (8, "1st Floor columns  (10'-0\" high — same)", 18, "", "", 10.0, 225, 6.37, 48, 96, 185),
    (9, "2nd Floor / Roof level columns  (10'-0\" high)", 18, "", "", 10.0, 225, 6.37, 48, 96, 185),
    ("T", "Subtotal — Columns \u2192", 788, 22.30, 168, 335, 648),
    ("S", 'D — FLOOR BEAMS  (12" x 18" — 430 rft per floor)', None),
    (10, "1st Floor beams  (over Ground Floor)", 1, 430, 1.0, 1.50, 645, 18.27, 138, 274, 531),
    (11, "2nd Floor beams", 1, 430, 1.0, 1.50, 645, 18.27, 138, 274, 531),
    (12, "3rd Floor / Roof beams", 1, 430, 1.0, 1.50, 645, 18.27, 138, 274, 531),
    ("T", "Subtotal — Beams \u2192", 1935, 54.80, 414, 822, 1592),
    ("S", 'E — RCC SLABS  (5" thick)', None),
    (13, '1st Floor slab  — 2,251 sft x 0.416\' (5")', 1, 2251, 1.0, 0.416, 938, 26.56, 201, 398, 771),
    (14, '2nd Floor slab  — 2,251 sft x 0.416\'', 1, 2251, 1.0, 0.416, 938, 26.56, 201, 398, 771),
    (15, 'Roof slab  — 2,251 sft x 0.416\'', 1, 2251, 1.0, 0.416, 938, 26.56, 201, 398, 771),
    ("T", "Subtotal — Slabs \u2192", 2814, 79.69, 602, 1195, 2314),
    ("S", "F — STAIRCASE, LINTEL, DROP WALL & OTHERS", None),
    (16, "Staircase 3 flights + landings + waist slab (estimated)", 3, "", "", "", 420, 11.89, 90, 178, 346),
    (17, "Lintel / sunshade / drop wall / parapet  (8% extra)", 1, "", "", "", 306, 8.67, 66, 130, 252),
    ("T", "Subtotal — Staircase & Others \u2192", 726, 20.56, 156, 308, 597),
]

row = 8
for e in DATA1:
    if e[0] == "S":
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
        c = ws.cell(row=row, column=1, value=e[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.alignment = Alignment(horizontal="left", vertical="center")
        for cc in range(1, 12):
            x = ws.cell(row=row, column=cc)
            x.fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
            x.border = thin
        ws.row_dimensions[row].height = 18
    elif e[0] == "T":
        c = ws.cell(row=row, column=2, value=e[1])
        cell(c, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, align="right")
        for cc in (1, 3, 4, 5, 6):
            cell(ws.cell(row=row, column=cc), bg=GOLD_LIGHT)
        vals = e[2:]
        for j, v in enumerate(vals):
            col = 7 + j
            x = ws.cell(row=row, column=col, value=v)
            cell(x, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0" if col != 8 else "#,##0.00")
        ws.row_dimensions[row].height = 18
    else:
        sl, desc, nos, L, B, H, cft, cum, bag, sand, chips = e
        vals = [sl, desc, nos, L, B, H, cft, cum, bag, sand, chips]
        for j, v in enumerate(vals, 1):
            x = ws.cell(row=row, column=j, value=v if v != "" else None)
            if j == 1:
                cell(x, color=GRAY_MID)
            elif j == 2:
                cell(x, align="left")
            elif j == 8:
                cell(x, fmt="#,##0.00")
            else:
                cell(x, fmt="#,##0")
        ws.row_dimensions[row].height = 16
    row += 1

# Grand total
for cc in range(1, 12):
    x = ws.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    x.border = thin
    x.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
c = ws.cell(row=row, column=1, value="GRAND TOTAL — ALL CONCRETE")
c.alignment = Alignment(horizontal="right", vertical="center")
for j, v in [(7, 7958), (8, 225.37), (9, 1702), (10, 3380), (11, 6546)]:
    x = ws.cell(row=row, column=j, value=v)
    x.font = Font(name="Calibri", size=10, color=GOLD, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
    x.number_format = "#,##0" if j != 8 else "#,##0.00"
ws.row_dimensions[row].height = 22
row += 1

banner(ws, f"A{row}:F{row}",
       'ADD 3% WASTAGE / LOSS  \u2192  cement order = 1,753 bags   |   Sand (Sylhet) \u2248 3,480 cft   |   Brick chips \u2248 6,740 cft',
       size=8, color=AMBER_TXT, bg=GOLD_LIGHT, height=17)
row += 1
banner(ws, f"A{row}:K{row}",
       "CONCRETE MATERIALS COST (materials only) — Cement Tk 550/bag, Sylhet sand Tk 65/cft, Brick chips Tk 145/cft, incl. water/admixture  \u2192  "
       "APPROX. Tk 24.8 LAKH (concrete materials only, rebar excluded)",
       size=8, color=WHITE, bg=GREEN, height=19)
row += 1

for note in [
    'Note: Footing D = thickness, S = side cover, T = pedestal thickness, X = projection — as per drawing. Column sizes: C-1 (12"x18"), C-2 (12"x15"), C-3 / C-4 (12"x12").',
    'Mix 1:1.5:3 (f\u2019c 3000 psi), water-cement ratio 0.45, Sylhet sand FM \u2265 2.0, brick chips 3/4" down 75% + 1/2" 25%, super-plasticizer assumed.',
    "Cement: 7.5 bags per Cum of concrete (7.55 incl. wastage); sand 0.42 Cum/Cum; chips 0.82 Cum/Cum. Field wastage 3% allowed separately.",
]:
    banner(ws, f"A{row}:K{row}", "\u2022 " + note, size=7, color=GRAY_MID, italic=False, bold=False, height=12)
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    row += 1

banner(ws, f"A{row}:K{row}",
       "SRayhan  |  Sheet 1 of 4 — Concrete Estimate  •  Next: Reinforcement Estimate (Sheet 2)",
       size=7, color=WHITE, bg=NAVY, height=14)
ws.print_title_rows = "7:7"

# =========================================================
# SHEET 2 — REINFORCEMENT (REBAR) ESTIMATE
# =========================================================
ws2 = wb.create_sheet("2-Rebar Estimate")
ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.page_setup.orientation = "landscape"
ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 1
ws2.print_options.horizontalCentered = True
ws2.sheet_view.showGridLines = False
set_widths(ws2, [6, 42, 11, 12, 12, 12, 11, 12, 14])

banner(ws2, "A1:I1", "SRAYHAN ENGINEERING & ARCHITECTURE   •   REINFORCEMENT ESTIMATE   •   GRADE 60 (fy \u2265 60,000 psi)",
       size=7, bg=NAVY, fg=GOLD, height=18)
banner(ws2, "A2:I2", "REINFORCEMENT (ROD) ESTIMATE  —  3-STORIED BUILDING (2,251 sft / floor)")
banner(ws2, "A3:I3", "Project: Mr. Saiful Islam Vai, Brahmanbaria   |   18 Columns, 18 Footings   |   Tie 10mm \u00d8 @ 4\"/6\"/4\"   |   Footing / Slab / Beam 16mm & 10mm",
       size=9, color=GRAY_MID, bold=False, height=14)

cards(ws2, 4, 5, [
    ("TOTAL REBAR (MAIN)", "21,360 kg\n21.36 MT", NAVY),
    ("REBAR GRADE", "Grade 60\n72,500 psi max", GOLD),
    ("MAIN DIAMETERS", "16mm (94%)\n10mm (ties)", GREEN),
    ("ESTIMATED COST", "Tk 95 / kg\n\u2248 Tk 20.3 Lakh", "374151"),
], [(1, 2), (3, 4), (5, 6), (7, 9)])

banner(ws2, "A6:I6",
       "Bar bending quantities taken from footing & column schedules — 5% extra allowed for site wastage / lap / hook",
       size=7.5, color=GRAY_MID, italic=True, bg=GRAY_LIGHT, bold=False, height=13)

h2 = ["Sl.", "Description (Component)", "Dia", "Main Steel\n(kg)", "Tie / Stirrup\n(kg)", "Total\n(kg)", "Rate\n(Tk/kg)", "Amount\n(Tk)", "Remarks"]
ws2.row_dimensions[7].height = 28
for i, h in enumerate(h2, 1):
    hdr(ws2.cell(row=7, column=i, value=h), size=7.5)

DATA2 = [
    ("S", 'A — FOOTING REBARS  (16mm @ 5" / 6" c/c)', None),
    (1, 'F-1 (6\'x7\' — 4 nos)  16mm @5" b/w — long + short direction', "16mm", "34 bars", "1,140 ft total", 0.48, 390, 95, 37050),
    (2, 'F-2 (6\'x6\' — 10 nos)  16mm @5" b/w', "16mm", "32 bars", "1,760 ft total", 0.48, 845, 95, 80275),
    (3, 'F-3 (5\'x5\' — 2 nos)  16mm @6" b/w', "16mm", "22 bars", "198 ft total", 0.48, 95, 95, 9025),
    (4, 'F-4 (4\'x4\' — 2 nos)  16mm @6" b/w', "16mm", "18 bars", "126 ft total", 0.48, 60, 95, 5700),
    ("T", "Subtotal — Footings \u2192", "", "", 1390, "3,224 ft", 132050),
    ("S", 'B — COLUMN REBARS — BELOW GRADE BEAM  (5\' high, Tie 10mm @4")', None),
    (5, 'C-1 (12"x18" — 4 nos, 8-16mm) + Tie 10mm (16 ties/column)', "16+10mm", 131, 54, 185, 95, 17575, ""),
    (6, "C-2 (12\"x15\" — 10 nos, 6-16mm) + Tie", "16+10mm", 245, 120, 365, 95, 34675, ""),
    (7, "C-3 (12\"x12\" — 2 nos, 6-16mm) + Tie", "16+10mm", 49, 21, 70, 95, 6650, ""),
    (8, "C-4 (12\"x12\" — 2 nos, 4-16mm) + Tie", "16+10mm", 33, 21, 54, 95, 5130, ""),
    ("T", "Subtotal — Columns below Grade Beam \u2192", 458, 216, 674, "", 64030),
    ("S", "C — COLUMN REBARS — ABOVE GRADE BEAM  (Ground + 1st + 2nd Floor, 10' high)", None),
    (9, 'Ground Floor columns (Tie @4" full height — 31 ties/column)', "16+10mm", 530, 400, 930, 95, 88350, ""),
    (10, '1st Floor columns (Tie @4"/6"/4" — 25 ties/column)', "16+10mm", 530, 340, 870, 95, 82650, ""),
    (11, "2nd Floor / Roof columns (25 ties/column)", "16+10mm", 530, 340, 870, 95, 82650, ""),
    ("T", "Subtotal — Columns above Grade Beam \u2192", 1590, 1080, 2670, "", 253650),
    ("S", 'D — GRADE BEAM + FLOOR BEAMS  (12"x18" — 6-16mm main + 10mm stirrup @6")', None),
    (12, "Grade Beam (430 rft — 4-16mm top + 2-16mm extra + stirrups)", "16+10mm", 1238, 781, 2019, 95, 191805, ""),
    (13, "1st Floor Beam (430 rft)", "16+10mm", 1238, 781, 2019, 95, 191805, ""),
    (14, "2nd Floor Beam (430 rft)", "16+10mm", 1238, 781, 2019, 95, 191805, ""),
    (15, "Roof Beam (430 rft)", "16+10mm", 1238, 781, 2019, 95, 191805, ""),
    ("T", "Subtotal — Beams (4 levels) \u2192", 4952, 3124, 8076, "", 767220),
    ("S", 'E — SLAB REBARS  (5" slab — 10mm @6" c/c B/W + extra top)', None),
    (16, '1st Floor slab 2,251 sft — 10mm @6" (\u2248 0.98 kg/sft)', "10mm", "9,004 ft total", "", 2206, 95, 209570),
    (17, "2nd Floor slab 2,251 sft", "10mm", "9,004 ft total", "", 2206, 95, 209570),
    (18, "Roof slab 2,251 sft", "10mm", "9,004 ft total", "", 2206, 95, 209570),
    ("T", "Subtotal — Slabs (3 floors) \u2192", "", "", 6618, "27,012 ft", 628710),
    ("S", "F — STAIRCASE, LINTEL & MISCELLANEOUS", None),
    (19, "Staircase 3 floors — 12mm main + 10mm distribution (estimated)", "12+10mm", 900, "", 900, 95, 85500, ""),
    (20, "Lintel / sunshade / parapet — 10mm + 12mm (estimated)", "12+10mm", 450, "", 450, 95, 42750, ""),
    ("T", "Subtotal — Miscellaneous \u2192", 1350, "", 1350, "", 128250),
]

def sub2(row, desc, main_kg, tie_kg, total_kg, length, amount):
    cell(ws2.cell(row=row, column=1), bg=GOLD_LIGHT)
    cell(ws2.cell(row=row, column=2, value=desc), size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, align="right")
    cell(ws2.cell(row=row, column=3), bg=GOLD_LIGHT)
    c4 = ws2.cell(row=row, column=4, value=main_kg if main_kg != "" else None)
    cell(c4, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
    c5 = ws2.cell(row=row, column=5, value=tie_kg if tie_kg != "" else None)
    cell(c5, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
    c6 = ws2.cell(row=row, column=6, value=total_kg if total_kg != "" else None)
    cell(c6, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
    c7 = ws2.cell(row=row, column=7, value=length if length != "" else None)
    cell(c7, size=7, bold=True, bg=GOLD_LIGHT, color=NAVY)
    c8 = ws2.cell(row=row, column=8, value=amount)
    cell(c8, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
    cell(ws2.cell(row=row, column=9), bg=GOLD_LIGHT)
    ws2.row_dimensions[row].height = 18

row = 8
for e in DATA2:
    if e[0] == "S":
        ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        c = ws2.cell(row=row, column=1, value=e[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.alignment = Alignment(horizontal="left", vertical="center")
        for cc in range(1, 10):
            x = ws2.cell(row=row, column=cc)
            x.fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
            x.border = thin
        ws2.row_dimensions[row].height = 18
    elif e[0] == "T":
        _, desc, main, tie, tot, ln, amt = e
        sub2(row, desc, main, tie, tot, ln, amt)
    else:
        if len(e) == 9 and isinstance(e[3], str):
            # footing row: (sl, desc, dia, bars, length, kg/ft, main_kg, rate, amount)
            sl, desc, dia, bars, ln, kgft, main, rate, amt = e
            tie = tot = ""; remark = ""
        elif len(e) == 9:
            # column/beam row: (sl, desc, dia, main_kg, tie_kg, total_kg, rate, amount, remark)
            sl, desc, dia, main, tie, tot, rate, amt, remark = e
            bars = ln = kgft = ""
        elif isinstance(e[3], int):
            # misc row: (sl, desc, dia, main_kg, "", total_kg, rate, amount)
            sl, desc, dia, main, tie, tot, rate, amt = e
            bars = ln = kgft = ""; remark = ""
        else:
            # slab row: (sl, desc, dia, length, "", main_kg, rate, amount)
            sl, desc, dia, ln, tie, main, rate, amt = e
            tot = main; bars = kgft = ""; remark = ""
        vals = [sl, desc, dia, main if main != "" else None, tie if tie != "" else None,
                tot if tot != "" else None, rate if rate != "" else None, amt, remark]
        for j, v in enumerate(vals, 1):
            x = ws2.cell(row=row, column=j, value=v)
            if j == 1:
                cell(x, color=GRAY_MID)
            elif j == 2:
                cell(x, size=7.5, align="left")
            elif j in (4, 5, 6):
                cell(x, bold=True, color=NAVY, fmt="#,##0")
            elif j == 8:
                cell(x, bold=True, color=NAVY, fmt="#,##0")
            else:
                cell(x, size=7.5, color="374151")
        extra = []
        if bars: extra.append(f"{bars} bars")
        if ln: extra.append(ln)
        if kgft: extra.append(f"{kgft} kg/ft")
        if extra:
            ws2.cell(row=row, column=9, value=", ".join(extra)).font = Font(name="Calibri", size=7, color=GRAY_MID)
        ws2.row_dimensions[row].height = 17
    row += 1

# Net total
for cc in range(1, 10):
    x = ws2.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    x.border = thin
    x.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws2.cell(row=row, column=1, value="NET TOTAL (excluding wastage)")
c.alignment = Alignment(horizontal="right", vertical="center")
x = ws2.cell(row=row, column=6, value=20340); x.font = Font(name="Calibri", size=10, color=GOLD, bold=True); x.number_format = "#,##0"; x.alignment = Alignment(horizontal="center", vertical="center")
x = ws2.cell(row=row, column=7, value=95); x.font = Font(name="Calibri", size=9, color=WHITE, bold=True); x.number_format = "#,##0"; x.alignment = Alignment(horizontal="center", vertical="center")
x = ws2.cell(row=row, column=8, value=1932300); x.font = Font(name="Calibri", size=10, color=GOLD, bold=True); x.number_format = "#,##0"; x.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[row].height = 22
row += 1

ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws2.cell(row=row, column=1, value="ADD 5% WASTAGE / LAP / HOOK  \u2192  ORDER 21,360 kg  (21.36 MT)")
c.font = Font(name="Calibri", size=8, color=AMBER_TXT, bold=True)
c.fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
c.alignment = Alignment(horizontal="right", vertical="center")
c.border = thin
for cc in (2, 3, 4, 5):
    ws2.cell(row=row, column=cc).fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
    ws2.cell(row=row, column=cc).border = thin
x = ws2.cell(row=row, column=6, value=21357); x.font = Font(name="Calibri", size=9, color=AMBER_TXT, bold=True); x.number_format = "#,##0"; x.alignment = Alignment(horizontal="center", vertical="center"); x.border = thin; x.fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
ws2.cell(row=row, column=7).fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
ws2.cell(row=row, column=7).border = thin
x = ws2.cell(row=row, column=8, value=2028915); x.font = Font(name="Calibri", size=9, color=AMBER_TXT, bold=True); x.number_format = "#,##0"; x.alignment = Alignment(horizontal="center", vertical="center"); x.border = thin; x.fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
ws2.cell(row=row, column=9).fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
ws2.cell(row=row, column=9).border = thin
ws2.row_dimensions[row].height = 18
row += 1

banner(ws2, f"A{row}:I{row}",
       "Market rebar price Tk 92,000–98,000/MT (2025-26) — Tk 95,000/MT taken here (incl. binding wire & carrying). "
       "If price rises, cost may increase by Tk 1.5–2.0 Lakh.",
       size=7.5, color=WHITE, bg=GREEN, height=16)
row += 1

for note in [
    "Unit weight: 16mm = 0.48 kg/ft (1.58 kg/m), 10mm = 0.19 kg/ft (0.62 kg/m), 12mm = 0.27 kg/ft (0.89 kg/m). Lengths measured center-line as per drawing.",
    'Column C-1: 8-16mm below GB, 4-16mm above; C-2: 6-16mm all floors; C-3 / C-4: 6 / 4-16mm — Tie 10mm @4" (GF) and @4"/6"/4" (1st/2nd floor).',
    'Beam stirrups 10mm @6" c/c; slab 10mm @6" c/c B/W + extra top — lap 45D and hook 10D allowed; 5% extra added for site wastage.',
]:
    banner(ws2, f"A{row}:I{row}", "\u2022 " + note, size=7, color=GRAY_MID, bold=False, height=12)
    ws2.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    row += 1

banner(ws2, f"A{row}:I{row}", "SRayhan  |  Sheet 2 of 4 — Reinforcement Estimate  •  Next: Finishing Estimate (Sheet 3)",
       size=7, color=WHITE, bg=NAVY, height=14)
ws2.print_title_rows = "7:7"

# =========================================================
# SHEET 3 — FINISHING ESTIMATE
# =========================================================
ws3 = wb.create_sheet("3-Finishing Estimate")
ws3.sheet_properties.pageSetUpPr.fitToPage = True
ws3.page_setup.orientation = "landscape"
ws3.page_setup.paperSize = ws3.PAPERSIZE_A3
ws3.page_setup.fitToWidth = 1
ws3.page_setup.fitToHeight = 1
ws3.print_options.horizontalCentered = True
ws3.sheet_view.showGridLines = False
set_widths(ws3, [6, 44, 9, 12, 12, 11, 14, 22])

banner(ws3, "A1:H1", "SRAYHAN ENGINEERING & ARCHITECTURE   •   FINISHING ESTIMATE   •   BRICK | PLASTER | TILES | PAINT | DOORS | ELECTRICAL | PLUMBING",
       size=7, bg=NAVY, fg=GOLD, height=18)
banner(ws3, "A2:H2", "FINISHING & FITTINGS ESTIMATE  —  2,251 sft x 3 floors")
banner(ws3, "A3:H3", "Project: Mr. Saiful Islam Vai, Brahmanbaria   |   5 Bed + Living + Dining + Kitchen + Toilets + 2 Veranda + Porch — identical plan each floor",
       size=9, color=GRAY_MID, bold=False, height=14)

cards(ws3, 4, 5, [
    ("TOTAL FINISHING AREA", "\u2248 6,753 sft\n3 floors", NAVY),
    ("BRICK WALLS", "\u2248 10,500 sft\n5\" + 10\"", GOLD),
    ("TILES + PLASTER", "Tiles 6,200 sft\nPlaster 19,000 sft", GREEN),
    ("FINISHING BUDGET", "\u2248 Tk 62.9 Lakh\n\u2248 932 Tk/sft", "374151"),
], [(1, 2), (3, 4), (5, 6), (7, 8)])

banner(ws3, "A6:H6",
       "Per-floor quantities shown separately — totals for 3 floors priced at market rate (2025-26)",
       size=7.5, color=GRAY_MID, italic=True, bg=GRAY_LIGHT, bold=False, height=13)

h3 = ["Sl.", "Work Item", "Unit", "Per Floor\nQty", "Total 3 Floors\nQty", "Rate\n(Tk/Unit)", "Amount\n(Tk)", "Remarks"]
ws3.row_dimensions[7].height = 28
for i, h in enumerate(h3, 1):
    hdr(ws3.cell(row=7, column=i, value=h), size=7.5)

DATA3 = [
    ("S", "A — MASONRY & PLASTER", None),
    (1, '5" brick wall (1st class brick, 1:5 mortar, incl. scaffolding)', "sft", 3500, 10500, 48, 504000, '10" outer walls included'),
    (2, "Plaster 12mm inside (1:5) + 12mm outside (1:4) + ceiling — 2 coat", "sft", 6400, 19200, 38, 729600, "NCF, incl. curing"),
    (3, "Outside weather-coat putty + primer (extra over plaster)", "sft", 1800, 5400, 12, 64800, "Outer walls only"),
    ("T", "Subtotal — Masonry & Plaster \u2192", 1298400),
    ("S", "B — TILES (Floor & Wall)", None),
    (4, 'Floor tiles 24"x24" homogeneous (matte/polished) + 4" skirting', "sft", 2100, 6300, 98, 617400, "RANGS / RAK, incl. labour"),
    (5, 'Bathroom wall tiles 12"x20" + anti-slip floor tiles', "sft", 850, 2550, 88, 224400, "5 toilets / floor"),
    (6, "Kitchen wall tiles + granite/marble counter + sink", "sft", 320, 960, 145, 139200, 'Granite 18mm'),
    (7, "Staircase tiles / marble + railing tiles", "sft", 180, 540, 110, 59400, "3 flights"),
    ("T", "Subtotal — Tiles \u2192", 1040400),
    ("S", "C — PAINT", None),
    (8, "Inside plastic paint (Asian / Berger) — putty 2 coat + primer + 2 coat paint", "sft", 6400, 19200, 34, 652800, "Ceiling white"),
    (9, "Outside weather coat + enamel for grills / railing", "sft", 1800, 5400, 30, 162000, "2 coats"),
    ("T", "Subtotal — Paint \u2192", 814800),
    ("S", "D — DOORS & WINDOWS (with grill & glass)", None),
    (10, "Main door (Segun / Chittagong teak) + frame — 1 no per floor", "nos", 1, 3, 45000, 135000, "Panel + frame + lock"),
    (11, "Internal doors (flush + veneer) — 5 nos per floor", "nos", 5, 15, 12500, 187500, "Handles, hinges incl."),
    (12, "Bathroom doors (PVC / WPVC) — 5 nos per floor", "nos", 5, 15, 7500, 112500, "Waterproof"),
    (13, "Windows: Thai sliding 4mm glass + MS grill + mosquito net — 8 nos/floor", "nos", 8, 24, 18500, 444000, "4'x5' average"),
    (14, "Veranda / porch railing + main gate grill work", "ls", 1, 3, 65000, 195000, "MS + paint"),
    ("T", "Subtotal — Doors & Windows \u2192", 1074000),
    ("S", "E — ELECTRICAL (BRB / BBS cable, MK switches)", None),
    (15, "Point wiring (light, fan, power, AC, geyser) — 35 points / floor", "point", 35, 105, 1850, 194250, "Holders, sockets incl."),
    (16, "DB / SDB, breakers, earthing, meter, main line", "ls", 1, 3, 55000, 165000, "1 DB per floor"),
    (17, "Fans, lights, holders, bell, exhaust — medium quality", "ls", 1, 3, 75000, 225000, "5 bed + living"),
    ("T", "Subtotal — Electrical \u2192", 584250),
    ("S", "F — PLUMBING & SANITARY (RFL / BSRM pipes)", None),
    (18, "Water lines (hot & cold), Gazi pipe, fittings, pressure test", "ls", 1, 3, 85000, 255000, "5 toilets / floor"),
    (19, "Sanitary fittings — commode, basin, bib cock, shower, mirror (mid-range)", "ls", 1, 3, 110000, 330000, "RFL / Cera"),
    (20, "Kitchen sink (SS double), waste line, septic/soak, roof tank 1500L", "ls", 1, 3, 65000, 195000, "1 tank / floor"),
    ("T", "Subtotal — Plumbing \u2192", 780000),
    ("S", "G — OTHERS", None),
    (21, "Parapet, roof treatment (waterproofing + lime terracing), drains", "ls", 1, 1, 250000, 250000, "Roof — one time"),
    (22, "Lift-core provision / stair railing SS-MS + handrail", "ls", 1, 1, 180000, 180000, "Optional"),
]

row = 8
for e in DATA3:
    if e[0] == "S":
        ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        c = ws3.cell(row=row, column=1, value=e[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.alignment = Alignment(horizontal="left", vertical="center")
        for cc in range(1, 9):
            x = ws3.cell(row=row, column=cc)
            x.fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
            x.border = thin
        ws3.row_dimensions[row].height = 18
    elif e[0] == "T":
        cell(ws3.cell(row=row, column=1), bg=GOLD_LIGHT)
        cell(ws3.cell(row=row, column=2, value=e[1]), size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, align="right")
        for cc in (3, 4, 5, 6):
            cell(ws3.cell(row=row, column=cc), bg=GOLD_LIGHT)
        x = ws3.cell(row=row, column=7, value=e[2])
        cell(x, size=9, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
        cell(ws3.cell(row=row, column=8), bg=GOLD_LIGHT)
        ws3.row_dimensions[row].height = 18
    else:
        sl, desc, unit, per, tot, rate, amt, rem = e
        vals = [sl, desc, unit, per, tot, rate, amt, rem]
        for j, v in enumerate(vals, 1):
            x = ws3.cell(row=row, column=j, value=v)
            if j == 1:
                cell(x, color=GRAY_MID)
            elif j == 2:
                cell(x, size=7.5, align="left")
            elif j == 7:
                cell(x, bold=True, color=NAVY, fmt="#,##0")
            elif j in (4, 5, 6):
                cell(x, fmt="#,##0")
            elif j == 8:
                cell(x, size=7, align="left", color=GRAY_MID)
            else:
                cell(x)
        ws3.row_dimensions[row].height = 17
    row += 1

# Grand total finishing
for cc in range(1, 9):
    x = ws3.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    x.border = thin
    x.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws3.cell(row=row, column=1, value="TOTAL FINISHING — 3 FLOORS COMPLETE")
c.alignment = Alignment(horizontal="right", vertical="center")
x = ws3.cell(row=row, column=7, value=5873850)
x.font = Font(name="Calibri", size=11, color=GOLD, bold=True)
x.alignment = Alignment(horizontal="center", vertical="center")
x.number_format = "#,##0"
ws3.row_dimensions[row].height = 22
row += 1

banner(ws3, f"A{row}:H{row}",
       "ADD 7% OVERHEAD (transport, scaffolding, curing, supervision, wastage)  \u2192  TOTAL FINISHING \u2248 Tk 62.9 LAKH  (\u2248 Tk 932/sft finishing only)",
       size=8, color=AMBER_TXT, bg=GOLD_LIGHT, height=19)
row += 1
banner(ws3, f"A{row}:H{row}",
       "Note: Lift, generator, interior decor, boundary wall and piling (if required) are NOT included — separate quotation if needed.",
       size=7.5, color=WHITE, bg=GREEN, height=16)
row += 1

for note in [
    "Rates: auto brick Tk 14/pc, cement Tk 550/bag, sand Tk 65/cft, tiles Tk 80–95/sft (medium range) — a 10–15% price rise would add Tk 4–6 Lakh to the total.",
    "Plaster 12mm thick (1:5 inside, 1:4 outside) + 2-coat putty — ceiling 9mm. Paint: plastic (inside), weather coat (outside), enamel (grills).",
    "Electrical: BRB 1.5/2.5 rm cable, MK/Click switches, 40A DP MCB; Plumbing: RFL uPVC + BSRM GI pipe — 5 toilets + 1 kitchen per floor.",
]:
    banner(ws3, f"A{row}:H{row}", "\u2022 " + note, size=7, color=GRAY_MID, bold=False, height=12)
    ws3.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    row += 1

banner(ws3, f"A{row}:H{row}", "SRayhan  |  Sheet 3 of 4 — Finishing Estimate  •  Next: Client Summary & BOQ (Sheet 4)",
       size=7, color=WHITE, bg=NAVY, height=14)
ws3.print_title_rows = "7:7"

# =========================================================
# SHEET 4 — SUMMARY & BOQ
# =========================================================
ws4 = wb.create_sheet("4-Summary & BOQ")
ws4.sheet_properties.pageSetUpPr.fitToPage = True
ws4.page_setup.orientation = "portrait"
ws4.page_setup.paperSize = ws4.PAPERSIZE_A4
ws4.page_setup.fitToWidth = 1
ws4.page_setup.fitToHeight = 1
ws4.print_options.horizontalCentered = True
ws4.sheet_view.showGridLines = False
set_widths(ws4, [7, 40, 9, 12, 12, 15, 20])

banner(ws4, "A1:G1", "SRAYHAN  —  ARCHITECTURE • STRUCTURE • INTERIOR • ESTIMATION   •   DHAKA | BRAHMANBARIA",
       size=7, bg=NAVY, fg=GOLD, height=18)
banner(ws4, "A2:G2", "PROFESSIONAL SUMMARY & BILL OF QUANTITIES")
banner(ws4, "A3:G3", CLIENT_LINE, size=8, color=WHITE, bg="374151", bold=False, height=16)

# Building info rows
ws4.merge_cells("A4:D4")
c = ws4["A4"]
c.value = "BUILDING INFO — Ground Floor Plan: 2,251 sft (54'-4\" x 37'-0\" + porch/veranda)"
c.font = Font(name="Calibri", size=8, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="left", vertical="center")
for col in "ABCD":
    ws4[f"{col}4"].fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
    ws4[f"{col}4"].border = thin
ws4.merge_cells("E4:G4")
c = ws4["E4"]
c.value = "5 Bed + Living + Dining + Kitchen + 4 Toilet + 2 Veranda + Porch / FLOOR"
c.font = Font(name="Calibri", size=7.5, color="1F2937")
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for col in "EFG":
    ws4[f"{col}4"].fill = PatternFill("solid", start_color=GOLD_LIGHT, end_color=GOLD_LIGHT)
    ws4[f"{col}4"].border = thin
ws4.row_dimensions[4].height = 16

ws4.merge_cells("A5:D5")
c = ws4["A5"]
c.value = "Structure: 18 columns (C-1 to C-4), 18 footings (F-1 to F-4), f'c 3000 psi, mix 1:1.5:3, Grade-60 steel"
c.font = Font(name="Calibri", size=7, color="374151")
c.alignment = Alignment(horizontal="left", vertical="center")
for col in "ABCD":
    ws4[f"{col}5"].fill = PatternFill("solid", start_color=GRAY_LIGHT, end_color=GRAY_LIGHT)
    ws4[f"{col}5"].border = thin
ws4.merge_cells("E5:G5")
c = ws4["E5"]
c.value = "Total built-up: 6,753 sft (2,251 x 3) + roof parapet"
c.font = Font(name="Calibri", size=7, color="374151", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
for col in "EFG":
    ws4[f"{col}5"].fill = PatternFill("solid", start_color=GREEN_LIGHT, end_color=GREEN_LIGHT)
    ws4[f"{col}5"].border = thin
ws4.row_dimensions[5].height = 14

# Floor-wise cards
floor_cards = [
    ("GROUND FLOOR", "Including Foundation", "Tk 68.2 Lakh", "Structure 31.5L + Finishing 22.7L + Foundation 14.0L", NAVY, NAVY_LIGHT),
    ("1ST FLOOR", "Fully Finished", "Tk 53.8 Lakh", "Structure 29.2L + Finishing 24.6L", GOLD, GOLD_LIGHT),
    ("2ND FLOOR", "Fully Finished", "Tk 54.5 Lakh", "Structure 29.8L + Finishing 24.7L (roof extra)", GREEN, GREEN_LIGHT),
]
positions = [(1, 2), (3, 4), (5, 7)]
for idx, (title, subt, amount, desc, color, bg) in enumerate(floor_cards):
    s, e = positions[idx]
    a, b = get_column_letter(s), get_column_letter(e)
    for r_off, val, fsize, fcolor, fbg in [
        (6, title, 8.5, WHITE, color),
        (7, subt, 7, color, bg),
        (8, amount, 12, color, bg),
        (9, desc, 6.5, "374151", WHITE),
    ]:
        ws4.merge_cells(f"{a}{r_off}:{b}{r_off}")
        c = ws4[f"{a}{r_off}"]
        c.value = val
        c.font = Font(name="Calibri", size=fsize, color=fcolor, bold=(r_off in (6, 8)))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = PatternFill("solid", start_color=fbg, end_color=fbg)
        c.border = thin
        for cc in range(s + 1, e + 1):
            x = ws4.cell(row=r_off, column=cc)
            x.fill = PatternFill("solid", start_color=fbg, end_color=fbg)
            x.border = thin
ws4.row_dimensions[6].height = 16
ws4.row_dimensions[7].height = 12
ws4.row_dimensions[8].height = 20
ws4.row_dimensions[9].height = 22

banner(ws4, "A10:G10",
       "TOTAL CONSTRUCTION COST (STRUCTURE + FINISHING)  \u2192  Tk 1,76,50,000  (One Crore Seventy-Six Lakh Fifty Thousand)   •   \u2248 Tk 2,614/sft   •   "
       "STRUCTURE ONLY \u2248 Tk 93.2 Lakh (Tk 1,380/sft)   •   FULLY FINISHED \u2248 Tk 1.77 Crore",
       size=9, color=WHITE, bg=NAVY, height=24)
banner(ws4, "A11:G11",
       "Ground Floor cost includes the foundation (footings + grade beam + earthwork + backfill) — that is why Ground Floor is higher. "
       "1st and 2nd floors carry only column-beam-slab structure + finishing.",
       size=7, color=AMBER_TXT, italic=True, bg=GOLD_LIGHT, bold=False, height=16)

banner(ws4, "A12:G12", "BILL OF QUANTITIES (BOQ) — SUMMARY  (details in Sheets 1-3)", size=10, color=NAVY, bg=NAVY_LIGHT, height=18)

h4 = ["Sl.", "Work Description", "Unit", "Qty", "Rate (Tk)", "Amount (Tk)", "Remarks"]
ws4.row_dimensions[13].height = 20
for i, h in enumerate(h4, 1):
    hdr(ws4.cell(row=13, column=i, value=h))

BOQ = [
    (1, "Foundation earthwork + CC + footings + grade beam (18 footings)", "ls", "1 job", "\u2014", 1450000, "F-1(4), F-2(10), F-3(2), F-4(2)"),
    (2, "RCC columns (18 nos x 3 floors + below GB) — concrete + shuttering", "cum", "22.3", "14,500", 323350, "C-1(4), C-2(10), C-3/4(4)"),
    (3, 'RCC beams (grade + 3 floor beams 12"x18" — 1,720 rft)', "cum", "73.1", "14,200", 1038020, "Shuttering included"),
    (4, 'RCC slabs (5" — 3 floors, 2,251 sft each)', "cum", "79.7", "13,800", 1099860, "Centering included"),
    (5, "Staircase, lintel, parapet, sunshade", "ls", "1 job", "\u2014", 520000, "3 flights"),
    (6, "Reinforcement (rod 21.36 MT, Grade 60 — supply + laying + binding)", "MT", "21.36", "95,000", 2029200, "16mm 94% + 10mm ties"),
    (7, "Cement (1,753 bags, mix 1:1.5:3)", "bag", "1,753", "550", 964150, "Fresh OPC"),
    (8, 'Brick chips / stone chips (6,740 cft)', "cft", "6,740", "145", 977300, '3/4" + 1/2"'),
    (9, "Sand — Sylhet + local (3,480 cft)", "cft", "3,480", "65", 226200, "FM \u2265 2.0"),
    (10, 'Brickwork 5" wall + 10" outer (10,500 sft)', "sft", "10,500", "48", 504000, "1st class, 1:5"),
    (11, "Plaster inside/outside + ceiling (19,200 sft)", "sft", "19,200", "38", 729600, "12mm, NCF"),
    (12, 'Floor tiles 24"x24" + skirting (6,300 sft)', "sft", "6,300", "98", 617400, "Homogeneous"),
    (13, "Toilet + kitchen wall tiles (3,510 sft + granite)", "sft", "3,510", "105", 368550, "Anti-slip floor"),
    (14, "Paint — plastic + weather coat + enamel (24,600 sft)", "sft", "24,600", "33", 811800, "Putty included"),
    (15, "Doors & windows (Thai + grill + MS) — 3 floors", "ls", "1 job", "\u2014", 1074000, "8 windows + 11 doors/floor"),
    (16, "Electrical — wiring + DB + fixtures (3 floors)", "ls", "1 job", "\u2014", 584250, "BRB cable, MK"),
    (17, "Plumbing + sanitary — pipes + fittings + fixtures (5 toilets/floor)", "ls", "1 job", "\u2014", 780000, "RFL, Cera"),
    (18, "Roof treatment, drainage, railing, misc + 7% overhead", "ls", "1 job", "\u2014", 430000, "Waterproofing"),
]

row = 14
for entry in BOQ:
    sl, desc, unit, qty, rate, amt, rem = entry
    vals = [sl, desc, unit, qty, rate, amt, rem]
    for j, v in enumerate(vals, 1):
        x = ws4.cell(row=row, column=j, value=v)
        if j == 1:
            cell(x, color=GRAY_MID)
        elif j == 2:
            cell(x, size=7.5, align="left")
        elif j == 6:
            cell(x, bold=True, color=NAVY, fmt="#,##0")
        elif j == 5:
            cell(x, size=7.5, color="374151")
        elif j == 7:
            cell(x, size=6.5, align="left", color=GRAY_MID)
        else:
            cell(x, size=7.5)
    ws4.row_dimensions[row].height = 14
    row += 1

# Subtotal BOQ
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="SUBTOTAL — BOQ (STRUCTURE + FINISHING, materials + labour)")
c.font = Font(name="Calibri", size=8, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
for cc in range(1, 8):
    x = ws4.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
    x.border = thin
x = ws4.cell(row=row, column=6, value=12676480)
x.font = Font(name="Calibri", size=9, color=NAVY, bold=True)
x.alignment = Alignment(horizontal="center", vertical="center")
x.number_format = "#,##0"
ws4.row_dimensions[row].height = 18
row += 1

for desc, amt in [
    ("+ Contingency & price escalation (8% — cement/rod price hike)", 1014118),
    ("+ Supervision, site engineer, tools, scaffolding, curing, transport (9%)", 1232854),
    ("+ VAT & documentation (approx. 2%)", 298429),
]:
    ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    c = ws4.cell(row=row, column=1, value=desc)
    c.font = Font(name="Calibri", size=7.5, color=GRAY_MID)
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = thin
    for cc in range(2, 6):
        ws4.cell(row=row, column=cc).border = thin
    x = ws4.cell(row=row, column=6, value=amt)
    x.font = Font(name="Calibri", size=8, color=GRAY_MID, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
    x.number_format = "#,##0"
    x.border = thin
    ws4.cell(row=row, column=7).border = thin
    ws4.row_dimensions[row].height = 13
    row += 1

# Grand total
for cc in range(1, 8):
    x = ws4.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    x.border = thin
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="GRAND TOTAL — PROPOSED BUDGET (FULLY FINISHED, 3 STOREYS)")
c.font = Font(name="Calibri", size=10, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
x = ws4.cell(row=row, column=6, value=17650000)
x.font = Font(name="Calibri", size=12, color=GOLD, bold=True)
x.alignment = Alignment(horizontal="center", vertical="center")
x.number_format = "#,##0"
ws4.row_dimensions[row].height = 24
row += 1

banner(ws4, f"A{row}:G{row}",
       "\u2248 Tk 2,614/sft (fully finished)   •   STRUCTURE ONLY (column-beam-slab stage) \u2248 Tk 1,380/sft (Tk 93.2 Lakh)   •   "
       "Ground Floor complete \u2248 Tk 68.2 Lakh; then each upper floor \u2248 Tk 54 Lakh — can be built phase by phase",
       size=7.5, color=AMBER_TXT, bg=GOLD_LIGHT, height=20)
row += 1

banner(ws4, f"A{row}:G{row}", "FLOOR-WISE COST BREAKDOWN", size=9, color=NAVY, bg=NAVY_LIGHT, height=16)
row += 1
fb_headers = ["Floor", "Structure (RCC + Rod)", "Finishing (Brick/Tiles/Paint)", "Electrical + Plumbing", "Others", "Total (Tk)", "When to Build"]
for i, h in enumerate(fb_headers, 1):
    hdr(ws4.cell(row=row, column=i, value=h), size=7.5)
ws4.row_dimensions[row].height = 16
row += 1

FLOOR_BR = [
    ("Ground Floor\n(1st to build)", "31,50,000", "18,50,000", "4,20,000", "14,00,000", "68,20,000", "Phase-1 (incl. foundation)"),
    ("1st Floor", "29,20,000", "19,20,000", "4,20,000", "95,000", "53,55,000", "Phase-2"),
    ("2nd Floor", "29,80,000", "19,30,000", "4,25,000", "1,15,000", "54,50,000", "Phase-3 + roof"),
    ("Roof / Parapet\n+ stair room", "\u2014", "\u2014", "\u2014", "2,50,000", "2,50,000", "Roof treatment"),
]
for entry in FLOOR_BR:
    for j, v in enumerate(entry, 1):
        x = ws4.cell(row=row, column=j, value=v)
        if j == 6:
            cell(x, size=8, bold=True, bg=GOLD_LIGHT, color=NAVY, fmt="#,##0")
        elif j == 1:
            cell(x, size=7.5, bold=True, color=NAVY)
            if "Ground" in v:
                x.fill = PatternFill("solid", start_color=NAVY_LIGHT, end_color=NAVY_LIGHT)
        elif j == 7:
            cell(x, size=7, color="374151")
        else:
            cell(x, size=7.5, color="374151")
    ws4.row_dimensions[row].height = 22
    row += 1

for cc in range(1, 8):
    x = ws4.cell(row=row, column=cc)
    x.fill = PatternFill("solid", start_color=NAVY, end_color=NAVY)
    x.border = thin
    x.font = Font(name="Calibri", size=8, color=WHITE, bold=True)
    x.alignment = Alignment(horizontal="center", vertical="center")
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="TOTAL (3 floors + roof)")
c.alignment = Alignment(horizontal="right", vertical="center")
x = ws4.cell(row=row, column=6, value=17875000)
x.font = Font(name="Calibri", size=9, color=GOLD, bold=True)
x.alignment = Alignment(horizontal="center", vertical="center")
x.number_format = "#,##0"
c2 = ws4.cell(row=row, column=7, value="\u2014")
c2.font = Font(name="Calibri", size=7, color=WHITE, bold=True)
c2.alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[row].height = 16
row += 1

# Payment schedule
banner(ws4, f"A{row}:G{row}", "PROPOSED PAYMENT SCHEDULE (MILESTONE-WISE)", size=9, color=NAVY, bg=NAVY_LIGHT, height=16)
row += 1
for i, h in enumerate(["Step", "Work Stage", "%", "Amount (Tk)", "Timeline", "", ""], 1):
    hdr(ws4.cell(row=row, column=i, value=h), bg=GREEN, size=7.5)
ws4.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
ws4.row_dimensions[row].height = 14
row += 1

PAY = [
    ("1", "Agreement + final drawing approval", "10%", "17,65,000"),
    ("2", "Foundation (earthwork + footings + grade beam)", "20%", "35,30,000"),
    ("3", "Ground floor structure (column-beam-slab)", "15%", "26,48,000"),
    ("4", "1st floor structure", "15%", "26,48,000"),
    ("5", "2nd floor structure + roof casting", "15%", "26,48,000"),
    ("6", "Finishing phase-1 (brickwork + plaster)", "10%", "17,65,000"),
    ("7", "Finishing phase-2 (tiles + paint + doors/windows + MEP)", "15%", "26,48,000"),
]
for sl, desc, per, amt in PAY:
    vals = [sl, desc, per, amt, "\u2014", "", ""]
    for j, v in enumerate(vals, 1):
        x = ws4.cell(row=row, column=j, value=v if v != "" else None)
        if j == 2:
            cell(x, size=7.5, align="left")
        elif j == 4:
            cell(x, size=8, bold=True, color=NAVY)
        elif j == 5:
            cell(x, size=7, color=GRAY_MID)
        else:
            cell(x, size=7.5, color="374151")
    ws4.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
    ws4.cell(row=row, column=6).border = thin
    ws4.cell(row=row, column=7).border = thin
    ws4.row_dimensions[row].height = 13
    row += 1

# Terms
banner(ws4, f"A{row}:G{row}", "TERMS & CONDITIONS", size=9, color=WHITE, bg=GREEN, height=16)
row += 1
TERMS = [
    "1. This estimate reflects market rates of 26 Oct 2026 (Brahmanbaria/Dhaka). If cement, rod, sand or brick prices rise 10–15%, the total may increase by Tk 5–7 Lakh. Re-quotation required after 15 days.",
    "2. Soil test and piling are NOT included. If piling is required (common in parts of Brahmanbaria), add Tk 35,000–45,000 per pile.",
    "3. Municipal/RAJUK approval, electricity, gas and water connections, boundary wall, gate and land filling are excluded.",
    "4. Timeline: foundation 30 days, each floor structure 25–30 days, finishing 30–35 days per floor — total approx. 8–10 months (weather permitting).",
    "5. Payments follow the milestone schedule — 50% advance required before materials reach site. Delays in payment slow down the work.",
    "6. Mid-range materials assumed in this BOQ — premium (RAK, Berger, BRB, SS) or economy options can be customized (price \u00b110–15%).",
    "7. Warranty: 10 years on structure (built as per design), 1 year service warranty on finishing.",
]
for t in TERMS:
    ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws4.cell(row=row, column=1, value=t)
    c.font = Font(name="Calibri", size=7, color="374151")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin
    ws4.row_dimensions[row].height = 18
    row += 1

banner(ws4, f"A{row}:G{row}",
       "NEXT STEPS:  Soil Test \u2192 Final Drawings (Arch + Structure) \u2192 Municipal Approval \u2192 Final Quotation \u2192 Start Work   •   "
       "For any question, call us — we will visit the site within 24 hours.",
       size=7.5, color=AMBER_TXT, bg=GOLD_LIGHT, height=18)
row += 1
banner(ws4, f"A{row}:G{row}",
       'Prepared by — SRayhan  |  Engineer & Architect  |  Phone: 019XX-XXXXXX  |  Email: info@srayhan.com  |  www.srayhan.com  |  "Your dream home, our commitment"',
       size=8, color=WHITE, bg=NAVY, height=20)
row += 1
banner(ws4, f"A{row}:G{row}",
       "Thank you for trusting us, Mr. Saiful Islam Vai!  •  This estimate is prepared exclusively for you — please ask before sharing.",
       size=7, color=GRAY_MID, italic=True, bold=False, height=12)
ws4.print_title_rows = "13:13"

for w in (ws, ws2, ws3, ws4):
    w.sheet_properties.pageSetUpPr.fitToPage = True
    w.oddHeader.center.text = "SRayhan — Estimate for Mr. Saiful Islam Vai, Brahmanbaria (3-Storied, 2251 sft/floor)"
    w.oddHeader.center.size = 8
    w.oddFooter.center.text = "Page &P of &N"
    w.oddFooter.center.size = 8

OUT = "/home/user/SRayhan/SRayhan_Estimate_Saiful_Islam_Brahmanbaria_3Storey_2251sft_ENGLISH.xlsx"
wb.save(OUT)
print("Saved:", OUT)

# --- quick consistency check ---
print("Sheet1 concrete grand:", 1051 + 645 + 788 + 1935 + 2814 + 726, "cft (expect 7958)")
print("Sheet2 net kg sections:", 1390 + 674 + 2670 + 8076 + 6618 + 1350)
print("Sheet3 sections:", 1298400 + 1040400 + 814800 + 1074000 + 584250 + 780000 + 250000 + 180000)
print("BOQ sum:", sum(x[5] for x in BOQ))
