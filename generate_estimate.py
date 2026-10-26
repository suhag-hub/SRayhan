#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import datetime

# Colors
NAVY = "0F2C4D"
GOLD = "B78A3B"
GOLD_LIGHT = "F5E6C8"
NAVY_LIGHT = "EAF0F8"
GREEN = "0E7A5C"
GRAY_DARK = "333333"
GRAY_MID = "6B7280"
GRAY_LIGHT = "F3F4F6"
WHITE = "FFFFFF"
BLUE_ACCENT = "1D4ED8"
ORANGE = "EA580C"

thin_border = Border(
    left=Side(style='thin', color='D1D5DB'),
    right=Side(style='thin', color='D1D5DB'),
    top=Side(style='thin', color='D1D5DB'),
    bottom=Side(style='thin', color='D1D5DB'),
)
thin_gold = Border(
    left=Side(style='thin', color='B78A3B'),
    right=Side(style='thin', color='B78A3B'),
    top=Side(style='thin', color='B78A3B'),
    bottom=Side(style='thin', color='B78A3B'),
)

def style_header(cell, fill_color=NAVY, font_color=WHITE, font_size=9, bold=True):
    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    cell.font = Font(name="Calibri", size=font_size, color=font_color, bold=bold)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

def style_cell(cell, font_size=9, bold=False, bg=None, color="1F2937", align="center", num_fmt=None, border=thin_border):
    cell.font = Font(name="Calibri", size=font_size, color=color, bold=bold)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    if bg:
        cell.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    cell.border = border
    if num_fmt:
        cell.number_format = num_fmt

wb = openpyxl.Workbook()

# =========================================================
# SHEET 1: CONCRETE ESTIMATE
# =========================================================
ws1 = wb.active
ws1.title = "১-কংক্রিট এস্টিমেট"
ws1.sheet_properties.pageSetUpPr.fitToPage = True
ws1.page_setup.orientation = "landscape"
ws1.page_setup.paperSize = ws1.PAPERSIZE_A3
ws1.page_setup.fitToWidth = 1
ws1.page_setup.fitToHeight = 1
ws1.print_options.horizontalCentered = True
ws1.sheet_view.showGridLines = False

# Column widths for Sheet1 - A to K
widths1 = [6, 32, 9, 9, 9, 8, 14, 11, 11, 13, 13]
cols1 = ["A","B","C","D","E","F","G","H","I","J","K"]
for c,w in zip(cols1, widths1):
    ws1.column_dimensions[c].width = w

# Top banner
ws1.merge_cells('A1:K1')
c = ws1['A1']
c.value = "SRAYHAN ENGINEERING & ARCHITECTURE  •  BRAHMANBARIA • DHAKA • STRUCTURE • ARCHITECTURE • ESTIMATION"
c.font = Font(name="Calibri", size=7, color=GOLD, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws1.row_dimensions[1].height = 18

# Title block
ws1.merge_cells('A2:K2')
c = ws1['A2']
c.value = "কংক্রিট এস্টিমেট  •  CONCRETE QUANTITY ESTIMATE  —  ৩-তলা আবাসিক ভবন"
c.font = Font(name="Calibri", size=13, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[2].height = 22

ws1.merge_cells('A3:K3')
c = ws1['A3']
c.value = "প্রকল্পঃ  Mr. Saiful Islam Vai  |  স্থানঃ Bhrammanbaria  |  প্রতি তলা ২২৫১ sft  |  মোট ৬৭৫৩ sft (৩ তলা)  |  তারিখঃ ২৬ অক্টোবর ২০২৬"
c.font = Font(name="Calibri", size=9, color=GRAY_MID, bold=False)
c.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[3].height = 14

# Info cards row
cards = [
    ("মোট কংক্রিট", "৭,৯৫৬ cft\n২২৫.৩ Cum", NAVY),
    ("ফাউন্ডেশন গভীরতা", "৬'-০\" (H)\nEarth Cutting", GOLD),
    ("কংক্রিট গ্রেড", "f’c = ৩০০০ psi\nMix 1:1.5:3", GREEN),
    ("সিমেন্ট লাগবে", "≈ ১,৬৯০ ব্যাগ\n(50kg)", "6B7280"),
]
for idx, (title, val, color) in enumerate(cards):
    col_start = idx*3 + 1  # 1,4,7,10
    col_end = col_start + 2
    # Need map col numbers to letters: 1=A, 4=D, 7=G, 10=J; K is 11 so last card ends at K
    # col 10=J, 11=K, so card 4 will be J:K (2 cols)
    start_letter = get_column_letter(col_start)
    end_letter = get_column_letter(min(col_end, 11))
    if idx==3:
        ws1.merge_cells(f'{start_letter}4:{end_letter}4')
        ws1.merge_cells(f'{start_letter}5:{end_letter}5')
    else:
        ws1.merge_cells(f'{start_letter}4:{end_letter}4')
        ws1.merge_cells(f'{start_letter}5:{end_letter}5')
    c1 = ws1[f'{start_letter}4']
    c1.value = title
    c1.font = Font(name="Calibri", size=8, color="FFFFFF", bold=True)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = PatternFill(start_color=color if color!= "6B7280" else "374151", end_color=color if color!= "6B7280" else "374151", fill_type="solid")
    c1.border = thin_border
    c2 = ws1[f'{start_letter}5']
    c2.value = val
    c2.font = Font(name="Calibri", size=9, color=color if color!="6B7280" else "111827", bold=True)
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if idx==0:
        c2.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
    elif idx==1:
        c2.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    elif idx==2:
        c2.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    else:
        c2.fill = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    c2.border = thin_border
    ws1.row_dimensions[4].height = 14
    ws1.row_dimensions[5].height = 20

# Note strip
ws1.merge_cells('A6:K6')
c = ws1['A6']
c.value = "সোর্সঃ Footing Schedule (F-1~F-4), Column Schedule (C-1~C-4), Column Layout Plan (১৮টি কলাম), Ground Floor Plan 2251 sft  •  অনুমানঃ Floor height 10'-0\", Slab 5\", Beam 12\"x18\", Grade Beam 12\"x18\""
c.font = Font(name="Calibri", size=7.5, color="6B7280", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
ws1.row_dimensions[6].height = 12

# Table header Row 7
headers1 = ["ক্রঃ","বিবরণ  (Description)","নস\nNos","L (ft)","B (ft)","D/H (ft)","পরিমাণ\ncft","Cum\n(m³)","সিমেন্ট\nব্যাগ","বালি\ncft","খোয়া\ncft"]
ws1.row_dimensions[7].height = 28
for col_idx, h in enumerate(headers1, start=1):
    cell = ws1.cell(row=7, column=col_idx, value=h)
    style_header(cell, fill_color=NAVY, font_size=8, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Data rows for Sheet1
# We'll create rows 8 onward
data1 = [
    # Section Foundation
    ("SECTION", "A — ফাউন্ডেশন কংক্রিট  (FOOTING)", "", "", "", "", "", "", "", "", ""),
    (1, "F-1  ফুটিং  6'-0\" x 7'-0\" x 2'-0\"  (D=24\")", 4, 6, 7, 2.00, 336, 9.51, 72, 143, 276),
    (2, "F-2  ফুটিং  6'-0\" x 6'-0\" x 1'-8\"  (D=20\")",10, 6, 6, 1.67, 600, 16.99,128, 255, 494),
    (3, "F-3  ফুটিং  5'-0\" x 5'-0\" x 1'-6\"  (D=18\")", 2, 5, 5, 1.50, 75, 2.12, 16, 32, 62),
    (4, "F-4  ফুটিং  4'-0\" x 4'-0\" x 1'-3\"  (D=15\")", 2, 4, 4, 1.25, 40, 1.13, 9, 17, 33),
    ("SUB", "উপ-মোট ফুটিং (Subtotal Footing) →", "", "", "", "", 1051, 29.76, 225, 446, 864),
    ("SECTION", "B — গ্রেড বিম  (GRADE BEAM  12\" x 18\")", "", "", "", "", "", "", "", "", ""),
    (5, "গ্রেড বিম — Longitudinal + Transverse  (430 rft)", 1, 430, 1.00, 1.50, 645, 18.27,138, 274, 531),
    ("SUB", "উপ-মোট গ্রেড বিম →", "", "", "", "", 645, 18.27,138, 274, 531),
    ("SECTION", "C — কলাম কংক্রিট  (COLUMN  12\" x 12\" ~ 12\" x 18\")", "", "", "", "", "", "", "", "", ""),
    (6, "কলাম  Below Grade Beam  (Pedestal 5'-0\" high — 18 nos)",18, "", "", 5.0, 113, 3.19, 24, 48, 93),
    (7, "কলাম  Ground Floor  (10'-0\" high — C-1:4, C-2:10, C-3:2, C-4:2)",18, "", "",10.0, 225, 6.37, 48, 96, 185),
    (8, "কলাম  1st Floor  (10'-0\" high — same)",18, "", "",10.0, 225, 6.37, 48, 96, 185),
    (9, "কলাম  2nd Floor / Roof level  (10'-0\" high)",18, "", "",10.0, 225, 6.37, 48, 96, 185),
    ("SUB", "উপ-মোট কলাম →", "", "", "", "", 788, 22.30,168, 335, 648),
    ("SECTION", "D — বিম  (FLOOR BEAM  12\" x 18\" — প্রতি তলায় 430 rft)", "", "", "", "", "", "", "", "", ""),
    (10, "1st Floor Beam (Ground roof)",1, 430,1.0,1.50,645,18.27,138,274,531),
    (11, "2nd Floor Beam",1, 430,1.0,1.50,645,18.27,138,274,531),
    (12, "3rd Floor / Roof Beam",1, 430,1.0,1.50,645,18.27,138,274,531),
    ("SUB", "উপ-মোট বিম →", "", "", "", "", 1935,54.80,414,822,1592),
    ("SECTION", "E — ছাদ / স্ল্যাব  (RCC SLAB  5\" thk.)", "", "", "", "", "", "", "", "", ""),
    (13, "1st Floor Slab  2251 sft x 0.416' (5\")",1, 2251,1.0,0.416,938,26.56,201,398,771),
    (14, "2nd Floor Slab  2251 sft x 0.416'",1, 2251,1.0,0.416,938,26.56,201,398,771),
    (15, "Roof Slab  2251 sft x 0.416'",1, 2251,1.0,0.416,938,26.56,201,398,771),
    ("SUB", "উপ-মোট স্ল্যাব →", "", "", "", "", 2814,79.69,602,1195,2314),
    ("SECTION", "F — সিঁড়ি, লিন্টেল, ড্রপ-ওয়াল ও অন্যান্য  (STAIR + LINTEL)", "", "", "", "", "", "", "", "", ""),
    (16, "সিঁড়ি 3-flight + Landing + Waist slab (অনুমান)",3, "", "", "", 420,11.89, 90,178,346),
    (17, "লিন্টেল / সানশেড / ড্রপওয়াল / প্যারাপেট (8% extra)",1, "", "", "", 306,8.67, 66,130,252),
    ("SUB", "উপ-মোট Stair & Others →", "", "", "", "", 726,20.56,156,308,597),
]

row = 8
for entry in data1:
    ws1.row_dimensions[row].height = 16 if entry[0]!="SECTION" and entry[0]!="SUB" else 18
    if entry[0]=="SECTION":
        ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
        c = ws1.cell(row=row, column=1, value=entry[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border
        # fill rest
        for col in range(2,12):
            cc = ws1.cell(row=row, column=col)
            cc.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
            cc.border = thin_border
    elif entry[0]=="SUB":
        # Subtotal row
        c = ws1.cell(row=row, column=1, value="")
        style_cell(c, font_size=8, bold=True, bg="FFFBEB", color=GOLD, align="center")
        c2 = ws1.cell(row=row, column=2, value=entry[1])
        style_cell(c2, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="right")
        for col_idx in range(3,7):
            cc = ws1.cell(row=row, column=col_idx)
            style_cell(cc, bg="FFFBEB", border=thin_border)
        vals = entry[6:11]  # g,h,i,j,k
        for col_idx, v in enumerate(vals, start=7):
            cc = ws1.cell(row=row, column=col_idx, value=v)
            style_cell(cc, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0' if col_idx in (7,9,10,11) else '#,##0.00')
            cc.border = thin_border
    else:
        # normal row
        sl, desc, nos, L, B, H, cft, cum, bag, sand, chips = entry
        vals = [sl, desc, nos, L, B, H, cft, cum, bag, sand, chips]
        for col_idx, v in enumerate(vals, start=1):
            cc = ws1.cell(row=row, column=col_idx, value=v)
            if col_idx==2:
                style_cell(cc, font_size=8, align="left", color="1F2937")
                cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx==1:
                style_cell(cc, font_size=8, align="center", color=GRAY_MID)
            elif col_idx in (7,8,9,10,11):
                # numeric
                if col_idx==8:
                    style_cell(cc, font_size=8, bold=False, align="center", num_fmt='#,##0.00')
                else:
                    style_cell(cc, font_size=8, bold=False, align="center", num_fmt='#,##0')
                if cft and isinstance(cft,int) and col_idx==7 and cft>500:
                    cc.font = Font(name="Calibri", size=8, color=NAVY, bold=True)
            else:
                style_cell(cc, font_size=8, align="center", color=GRAY_MID)
                if isinstance(v, (int,float)) and v!="":
                    cc.number_format = '#,##0.00' if isinstance(v,float) else '#,##0'
    row +=1

# Grand total row
ws1.row_dimensions[row].height = 22
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
for col in range(1,12):
    cc = ws1.cell(row=row, column=col)
    cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cc.border = thin_border
    cc.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws1.cell(row=row, column=1).value = "মোট কংক্রিট (GRAND TOTAL)"
ws1.cell(row=row, column=1).alignment = Alignment(horizontal="right", vertical="center")
# values
grand_vals = {7:7958, 8:225.37, 9:1702, 10:3380, 11:6546}
for col_idx, v in grand_vals.items():
    cc = ws1.cell(row=row, column=col_idx, value=v)
    cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cc.font = Font(name="Calibri", size=10, color=GOLD, bold=True)
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.border = thin_border
    cc.number_format = '#,##0' if col_idx!=8 else '#,##0.00'
row +=1
# wastage row
ws1.row_dimensions[row].height = 18
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws1.cell(row=row, column=1, value=" + 3% ওয়েস্টেজ / লস ধরে সিমেন্ট অর্ডার = ১,৭৫৩ ব্যাগ  |  বালি (Sylhet) ≈ ৩,৪৮০ cft  |  খোয়া (Brick chips) ≈ ৬,৭৪০ cft")
c.font = Font(name="Calibri", size=8, color="92400E", bold=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
c.border = thin_border
for col in range(7,12):
    cc = ws1.cell(row=row, column=col)
    cc.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    cc.border = thin_border
row+=1

# Cost estimate strip for concrete
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
c = ws1.cell(row=row, column=1, value="কংক্রিট মালামাল খরচ (শুধু মালামাল) — সিমেন্ট ৫৫০ টাকা/ব্যাগ, Sylhet বালি ৬৫ টাকা/cft, খোয়া ১৪৫ টাকা/cft, পানি/এডমিক্সচার সহ →  আনুমানিক  ২৪.৮ লক্ষ টাকা  (শুধু কংক্রিটের মালামাল, রড ছাড়া)")
c.font = Font(name="Calibri", size=8, color=WHITE, bold=True)
c.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.border = thin_border
ws1.row_dimensions[row].height = 20
row+=1

# Notes
notes1 = [
    "নোটঃ Footing D = পুরুত্ব, S = Side cover, T = Pedestal thickness, X = প্রজেকশন — ড্রয়িং অনুযায়ী। কলাম সাইজ C-1 (12\"x18\"), C-2 (12\"x15\"), C-3/C-4 (12\"x12\")।",
    "ধরা হয়েছে — Mix 1:1.5:3 (f’c 3000 psi), পানি-সিমেন্ট রেশিও 0.45, Sylhet sand FM≥2.0, Brick chips 3/4\" downgraded 75% + 1/2\" 25%, Super plasticizer।",
    "সিমেন্ট হিসাবঃ 1 Cum কংক্রিটে 7.5 ব্যাগ (wastage সহ 7.55)। বালি 0.42 Cum/Cum, খোয়া 0.82 Cum/Cum। Field wastage 3% আলাদা ধরা হয়েছে।",
]
for note in notes1:
    ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
    c = ws1.cell(row=row, column=1, value="• " + note)
    c.font = Font(name="Calibri", size=7, color=GRAY_MID)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border
    ws1.row_dimensions[row].height = 12
    row+=1

# Footer
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=11)
c = ws1.cell(row=row, column=1, value="SRayhan  |  019...  |  info@srayhan.com  |  Sheet 1 of 4  — Concrete Estimate  •  Next: Reinforcement (Sheet 2)")
c.font = Font(name="Calibri", size=7, color=WHITE, bold=True)
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[row].height = 14

# Print settings
ws1.print_title_rows = '7:7'

# =========================================================
# SHEET 2: REINFORCEMENT
# =========================================================
ws2 = wb.create_sheet(title="২-রড এস্টিমেট")
ws2.sheet_properties.pageSetUpPr.fitToPage = True
ws2.page_setup.orientation = "landscape"
ws2.page_setup.paperSize = ws2.PAPERSIZE_A3
ws2.page_setup.fitToWidth = 1
ws2.page_setup.fitToHeight = 1
ws2.print_options.horizontalCentered = True
ws2.sheet_view.showGridLines = False
widths2 = [6, 30, 12, 12, 12, 14, 14, 14, 16]
cols2 = ["A","B","C","D","E","F","G","H","I"]
for c,w in zip(cols2, widths2):
    ws2.column_dimensions[c].width = w

ws2.merge_cells('A1:I1')
c = ws2['A1']
c.value = "SRAYHAN ENGINEERING & ARCHITECTURE  •  REINFORCEMENT ESTIMATE  •  GRADE 60 (fy ≥ 60,000 psi)"
c.font = Font(name="Calibri", size=7, color=GOLD, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws2.row_dimensions[1].height = 18

ws2.merge_cells('A2:I2')
c = ws2['A2']
c.value = "রেইনফোর্সমেন্ট এস্টিমেট  •  REINFORCEMENT / ROD ESTIMATE  —  ৩-তলা ভবন (২২৫১ sft/তলা)"
c.font = Font(name="Calibri", size=13, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[2].height = 22

ws2.merge_cells('A3:I3')
c = ws2['A3']
c.value = "প্রকল্পঃ Mr. Saiful Islam Vai, Bhrammanbaria  |  কলাম ১৮টি, ফুটিং ১৮টি  |  Tie 10mm Ø @ 4\"/6\"/4\"  |  Footing/Slab/Beam 16mm & 10mm"
c.font = Font(name="Calibri", size=9, color=GRAY_MID)
c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[3].height = 14

# Info cards
cards2 = [
    ("মোট রড (Main)", "২১,৩৬০ kg\n২১.৪ MT", NAVY),
    ("রড গ্রেড", "60 Grade\n72,500 psi max", GOLD),
    ("প্রধান ডায়া", "16mm (94%)\n10mm (tie)", GREEN),
    ("আনুমানিক দাম", "৯৫ টাকা/kg\n~২০.৩ লক্ষ টাকা", "374151"),
]
for idx, (title, val, color) in enumerate(cards2):
    col_start = idx*2 +1
    col_end = col_start+1
    # handle last card to fill to I
    start_letter = get_column_letter(col_start)
    end_letter = get_column_letter(col_end)
    # need map 1-2,3-4,5-6,7-9? For 4 cards with 9 cols: card widths 2,2,2,3
    # Adjust
cards_positions = [(1,2),(3,4),(5,6),(7,9)]
for idx, (s,e) in enumerate(cards_positions):
    title, val, color = cards2[idx]
    start_letter = get_column_letter(s)
    end_letter = get_column_letter(e)
    ws2.merge_cells(f'{start_letter}4:{end_letter}4')
    ws2.merge_cells(f'{start_letter}5:{end_letter}5')
    c1 = ws2[f'{start_letter}4']
    c1.value = title
    c1.font = Font(name="Calibri", size=7.5, color="FFFFFF", bold=True)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = PatternFill(start_color=color if color!="374151" else "374151", end_color=color if color!="374151" else "374151", fill_type="solid")
    c1.border = thin_border
    c2 = ws2[f'{start_letter}5']
    c2.value = val
    c2.font = Font(name="Calibri", size=9, color=color if color!="374151" else "111827", bold=True)
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    colors_bg = ["EAF0F8","FFFBEB","ECFDF5","F3F4F6"]
    c2.fill = PatternFill(start_color=colors_bg[idx], end_color=colors_bg[idx], fill_type="solid")
    c2.border = thin_border
ws2.row_dimensions[4].height = 14
ws2.row_dimensions[5].height = 20

ws2.merge_cells('A6:I6')
c = ws2['A6']
c.value = "ড্রয়িং অনুযায়ী প্রতিটি ফুটিং ও কলামের রডের হিসাব নিচে দেওয়া হলো — সাইটে ৩-৫% ওয়েস্টেজ/ল্যাপ ধরা হয়েছে"
c.font = Font(name="Calibri", size=7.5, color="6B7280", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
ws2.row_dimensions[6].height = 12

headers2 = ["ক্রঃ","বিবরণ (Component)","ডায়া\nDia","দৈর্ঘ্য/\nনস","মোট দৈর্ঘ্য\n(মিটার)","ওজন\nkg/ft","মোট ওজন\nkg","রেট\nTk/kg","মোট টাকা\nTk"]
ws2.row_dimensions[7].height = 28
for col_idx, h in enumerate(headers2, start=1):
    cell = ws2.cell(row=7, column=col_idx, value=h)
    style_header(cell, fill_color=NAVY, font_size=7.5, bold=True)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

data2 = [
    ("SECTION", "A — ফুটিং রড  (FOOTING — 16mm @ 5\" / 6\" c/c)"),
    (1, "F-1  (6'x7' — 4 nos)  16mm @5\" b/w — long + short", "16mm", "34 bars", "1,140 ft", 0.48, 390, 95, 37050),
    (2, "F-2  (6'x6' — 10 nos) 16mm @5\" b/w", "16mm", "32 bars", "1,760 ft", 0.48, 845, 95, 80275),
    (3, "F-3  (5'x5' — 2 nos)  16mm @6\" b/w", "16mm", "22 bars", "198 ft", 0.48, 95, 95, 9025),
    (4, "F-4  (4'x4' — 2 nos)  16mm @6\" b/w", "16mm", "18 bars", "126 ft", 0.48, 60, 95, 5700),
    ("SUB", "উপ-মোট ফুটিং", "", "", "3,224 ft", "", 1390, "", 132050),
    ("SECTION", "B — কলাম রড — নিচ (BELOW GRADE BEAM, 5' high, Tie 10mm @4\")"),
    (5, "C-1 (12\"x18\"  4 nos — 8-16mm) + Tie 10mm (16 ties/col)", "16+10mm", "", "", "", 131, 54, 95, 17575),
    (6, "C-2 (12\"x15\" 10 nos — 6-16mm) + Tie", "16+10mm", "", "", "", 245,120, 95, 34675),
    (7, "C-3 (12\"x12\"  2 nos — 6-16mm) + Tie", "16+10mm", "", "", "", 49,21, 95, 6650),
    (8, "C-4 (12\"x12\"  2 nos — 4-16mm) + Tie", "16+10mm", "", "", "", 33,21, 95, 5130),
    ("SUB", "উপ-মোট Below Grade Column", "", "", "", "", 458, 216, "", 64030),
    ("SECTION", "C — কলাম রড — উপরে (GROUND + 1ST + 2ND FLOOR, 10' high)"),
    (9, "Ground Floor Columns (GF — Tie @4\" all through, 31 ties/col)", "16+10mm", "", "", "", 530, 400, 95, 88350),
    (10, "1st Floor Columns (Tie @4\"/6\"/4\" — 25 ties/col)", "16+10mm", "", "", "", 530, 340, 95, 82650),
    (11, "2nd Floor / Roof Columns (25 ties/col)", "16+10mm", "", "", "", 530, 340, 95, 82650),
    ("SUB", "উপ-মোট Above Grade Columns", "", "", "", "", 1590,1080, "", 253650),
    ("SECTION", "D — গ্রেড বিম + ফ্লোর বিম (BEAM 12\"x18\" — 6-16mm main + 10mm stirrup @6\")"),
    (12, "Grade Beam (430 rft — 4-16mm top+2-16mm extra + stirrup)", "16+10mm", "", "2,580 ft + 860 ties", "", 1238,781, 95, 191805),
    (13, "1st Floor Beam (430 rft)", "16+10mm", "", "2,580 ft", "", 1238,781, 95, 191805),
    (14, "2nd Floor Beam (430 rft)", "16+10mm", "", "2,580 ft", "", 1238,781, 95, 191805),
    (15, "Roof Beam (430 rft)", "16+10mm", "", "2,580 ft", "", 1238,781, 95, 191805),
    ("SUB", "উপ-মোট Beam (4 levels)", "", "", "10,320 ft", "", 4952,3124, "", 767220),
    ("SECTION", "E — স্ল্যাব রড (SLAB 5\" — 10mm @6\" c/c B/W + Top extra)"),
    (16, "1st Floor Slab 2251 sft — 10mm @6\" (≈ 0.98 kg/sft)", "10mm", "", "9,004 ft", 0.188, 2206, "", 95, 209570),
    (17, "2nd Floor Slab 2251 sft", "10mm", "", "9,004 ft", 0.188, 2206, "", 95, 209570),
    (18, "Roof Slab 2251 sft", "10mm", "", "9,004 ft", 0.188, 2206, "", 95, 209570),
    ("SUB", "উপ-মোট Slab (3 floors)", "", "", "27,012 ft", "", 6618, "", "", 628710),
    ("SECTION", "F — সিঁড়ি, লিন্টেল, মিসেলেনিয়াস"),
    (19, "Staircase 3 floors — 12mm main + 10mm dist. (est.)", "12+10mm", "", "", "", 900, "", 95, 85500),
    (20, "Lintel / Sunshade / Parapet — 10mm + 12mm (est.)", "12+10mm", "", "", "", 450, "", 95, 42750),
    ("SUB", "উপ-মোট Misc", "", "", "", "", 1350, "", "", 128250),
]

row=8
for entry in data2:
    ws2.row_dimensions[row].height = 17 if entry[0]!="SECTION" and entry[0]!="SUB" else 19
    if entry[0]=="SECTION":
        ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
        c = ws2.cell(row=row, column=1, value=entry[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border
        for col in range(2,10):
            cc = ws2.cell(row=row, column=col)
            cc.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
            cc.border = thin_border
    elif entry[0]=="SUB":
        vals = [entry[1], entry[2] if len(entry)>2 else "", entry[3] if len(entry)>3 else "", entry[4] if len(entry)>4 else "", entry[5] if len(entry)>5 else "", entry[6] if len(entry)>6 else "", entry[7] if len(entry)>7 else "", entry[8] if len(entry)>8 else "", entry[9] if len(entry)>9 else ""]
        # But structure is different for SUB: we defined as tuple with fewer items
        # We'll handle manually: entry = (SUB, desc, ..., ..., ..., weight..., ..., amount)
        # Actually we stored as 10 items? Let's just render sub total row styling
        c = ws2.cell(row=row, column=1, value="")
        style_cell(c, bg="FFFBEB", border=thin_border)
        c2 = ws2.cell(row=row, column=2, value=entry[1])
        style_cell(c2, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="right")
        # cols 3-5 empty yellow
        for col_idx in range(3,6):
            cc = ws2.cell(row=row, column=col_idx)
            style_cell(cc, bg="FFFBEB", border=thin_border)
        # G,H,I have values
        # entry indices: for SUB we have varied. We'll extract last three numeric fields
        # For footing SUB: entry = (SUB, desc,"","", "3,224 ft","",1390,"",132050)
        # So mapping: col6=E? Actually G=7 weight, H=8 rate, I=9 amount
        # Let's do: if entry has 9 elements: idx mapping
        if len(entry)==9:
            # SUB footing style
            g_val = entry[6]
            h_val = entry[7]
            i_val = entry[8]
            # also E maybe entry[4]?
            e_val = entry[4]
            ws2.cell(row=row, column=5, value=e_val).font = Font(name="Calibri", size=7, color=NAVY, bold=True)
            ws2.cell(row=row, column=5).fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
            ws2.cell(row=row, column=5).alignment = Alignment(horizontal="center", vertical="center")
            ws2.cell(row=row, column=5).border = thin_border
            for col_idx, v in [(7,g_val),(8,h_val),(9,i_val)]:
                cc = ws2.cell(row=row, column=col_idx, value=v)
                style_cell(cc, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0', border=thin_border)
        else:
            # generic
            for col_idx in range(6,10):
                cc = ws2.cell(row=row, column=col_idx)
                style_cell(cc, bg="FFFBEB", border=thin_border)
            # try to place weight and amount
            try:
                cc7 = ws2.cell(row=row, column=7, value=entry[6])
                style_cell(cc7, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0')
                cc8 = ws2.cell(row=row, column=8, value=entry[7])
                style_cell(cc8, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0')
                cc9 = ws2.cell(row=row, column=9, value=entry[8] if len(entry)>8 else "")
                style_cell(cc9, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0')
            except: pass
        # Ensure borders
        for col_idx in range(1,10):
            ws2.cell(row=row, column=col_idx).border = thin_border
    else:
        # normal
        # entry = (sl, desc, dia, nos, length, weight_per, total_weight, rate, amount)
        # but some have weight-per as 0.48 etc - we'll just place values as stored
        vals = [entry[0], entry[1], entry[2], entry[3], entry[4], entry[5], entry[6], entry[7], entry[8]]
        for col_idx, v in enumerate(vals, start=1):
            cc = ws2.cell(row=row, column=col_idx, value=v)
            if col_idx==2:
                style_cell(cc, font_size=7.5, align="left", color="1F2937")
                cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx==1:
                style_cell(cc, font_size=8, align="center", color=GRAY_MID)
            elif col_idx in (7,9):
                style_cell(cc, font_size=8, bold=True, align="center", color=NAVY, num_fmt='#,##0')
            elif col_idx==6:
                style_cell(cc, font_size=7, align="center", color=GRAY_MID, num_fmt='0.00')
            else:
                style_cell(cc, font_size=7.5, align="center", color="374151")
            cc.border = thin_border
            if isinstance(v, (int,float)) and col_idx in (7,9):
                cc.number_format = '#,##0'
    row+=1

# Grand total steel
ws2.row_dimensions[row].height = 22
for col in range(1,10):
    cc = ws2.cell(row=row, column=col)
    cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cc.border = thin_border
    cc.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    cc.alignment = Alignment(horizontal="center", vertical="center")
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws2.cell(row=row, column=1)
c.value = "মোট রড (NET TOTAL — without wastage)"
c.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws2.cell(row=row, column=7, value=20340).font = Font(name="Calibri", size=10, color=GOLD, bold=True)
ws2.cell(row=row, column=7).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws2.cell(row=row, column=7).number_format = '#,##0'
ws2.cell(row=row, column=9, value=1932300).font = Font(name="Calibri", size=10, color=GOLD, bold=True)
ws2.cell(row=row, column=9).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws2.cell(row=row, column=9).number_format = '#,##0'
row+=1

ws2.row_dimensions[row].height = 18
ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws2.cell(row=row, column=1, value=" + 5% Wastage / Lap / Hook  →  অর্ডার করুন  21,360 kg  (21.36 MT)")
c.font = Font(name="Calibri", size=8, color="92400E", bold=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
c.border = thin_border
for col in [7,8,9]:
    cc = ws2.cell(row=row, column=col)
    cc.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    cc.border = thin_border
ws2.cell(row=row, column=7, value=21357).font = Font(name="Calibri", size=9, color="92400E", bold=True)
ws2.cell(row=row, column=7).alignment = Alignment(horizontal="center", vertical="center")
ws2.cell(row=row, column=7).number_format = '#,##0'
ws2.cell(row=row, column=7).border = thin_border
ws2.cell(row=row, column=9, value=2028915).font = Font(name="Calibri", size=9, color="92400E", bold=True)
ws2.cell(row=row, column=9).alignment = Alignment(horizontal="center", vertical="center")
ws2.cell(row=row, column=9).number_format = '#,##0'
ws2.cell(row=row, column=9).border = thin_border
row+=1

ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
c = ws2.cell(row=row, column=1, value="বাজারে রড ৯২,০০০–৯৮,০০০ টাকা/MT (২০২৫-২৬) — এখানে ৯৫,০০০ টাকা/MT ধরা হয়েছে (binding wire + carrying সহ)। দাম বাড়লে ১.৫–২.০ লক্ষ বাড়তে পারে।")
c.font = Font(name="Calibri", size=7.5, color=WHITE, bold=True)
c.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[row].height = 16
row+=1

notes2 = [
    "ডায়া ওজনঃ 16mm = 0.48 kg/ft (1.58 kg/m), 10mm = 0.19 kg/ft (0.62 kg/m), 12mm = 0.27 kg/ft (0.89 kg/m)। হিসাব drawing অনুযায়ী center-line থেকে।",
    "Column C-1: Below GB 8-16mm, Above 4-16mm; C-2: 6-16mm all floors; C-3/C-4: 6/4-16mm — Tie 10mm @4\" (GF) ও @4\"/6\"/4\" (1F/2F)।",
    "Beam stirrup 10mm @6\" c/c, Slab 10mm @6\" c/c B/W + extra top — সাইটে lap 45D, hook 10D ধরে 5% extra যোগ করা হয়েছে।",
]
for note in notes2:
    ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
    c = ws2.cell(row=row, column=1, value="• " + note)
    c.font = Font(name="Calibri", size=7, color=GRAY_MID)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border
    ws2.row_dimensions[row].height = 12
    row+=1

ws2.merge_cells(start_row=row, start_column=1, end_row=row, end_column=9)
c = ws2.cell(row=row, column=1, value="SRayhan  |  Sheet 2 of 4 — Reinforcement Estimate  •  Next: Finishing (Sheet 3)")
c.font = Font(name="Calibri", size=7, color=WHITE, bold=True)
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[row].height = 14

ws2.print_title_rows = '7:7'

# =========================================================
# SHEET 3: FINISHING ESTIMATE
# =========================================================
ws3 = wb.create_sheet(title="৩-ফিনিশিং এস্টিমেট")
ws3.sheet_properties.pageSetUpPr.fitToPage = True
ws3.page_setup.orientation = "landscape"
ws3.page_setup.paperSize = ws3.PAPERSIZE_A3
ws3.page_setup.fitToWidth = 1
ws3.page_setup.fitToHeight = 1
ws3.print_options.horizontalCentered = True
ws3.sheet_view.showGridLines = False
widths3 = [6, 34, 10, 14, 14, 14, 15, 16]
cols3 = ["A","B","C","D","E","F","G","H"]
for c,w in zip(cols3, widths3):
    ws3.column_dimensions[c].width = w

ws3.merge_cells('A1:H1')
c = ws3['A1']
c.value = "SRAYHAN ENGINEERING & ARCHITECTURE  •  FINISHING ESTIMATE  •  BRICK | PLASTER | TILES | PAINT | DOORS | ELECTRICAL | PLUMBING"
c.font = Font(name="Calibri", size=7, color=GOLD, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws3.row_dimensions[1].height = 18

ws3.merge_cells('A2:H2')
c = ws3['A2']
c.value = "ফিনিশিং এস্টিমেট  •  FINISHING & FITTINGS ESTIMATE  —  ২২৫১ sft × ৩ তলা"
c.font = Font(name="Calibri", size=13, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[2].height = 22

ws3.merge_cells('A3:H3')
c = ws3['A3']
c.value = "প্রকল্পঃ Mr. Saiful Islam Vai, Bhrammanbaria  |  ৫ বেড, ৪ টয়লেট, লিভিং, ডাইনিং, কিচেন, ২ ভেরান্ডা, পোর্চ — প্রতি তলায় একই প্ল্যান"
c.font = Font(name="Calibri", size=9, color=GRAY_MID)
c.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[3].height = 14

cards3 = [
    ("মোট ফিনিশিং এরিয়া", "≈ ৬,৭৫৩ sft\n৩ তলা", NAVY),
    ("ব্রিক ওয়াল", "≈ ১০,৫০০ sft\n৫\" + ১০\"", GOLD),
    ("টাইলস + প্লাস্টার", "টাইলস ৬,২০০ sft\nপ্লাস্টার ১৯,০০০ sft", GREEN),
    ("ফিনিশিং বাজেট", "≈ ৭৯.৫ লক্ষ Tk\n১,১৭৮ Tk/sft", "374151"),
]
pos3 = [(1,2),(3,4),(5,6),(7,8)]
for idx,(s,e) in enumerate(pos3):
    title,val,color = cards3[idx]
    start_letter = get_column_letter(s)
    end_letter = get_column_letter(e)
    ws3.merge_cells(f'{start_letter}4:{end_letter}4')
    ws3.merge_cells(f'{start_letter}5:{end_letter}5')
    c1 = ws3[f'{start_letter}4']
    c1.value = title
    c1.font = Font(name="Calibri", size=7.5, color="FFFFFF", bold=True)
    c1.alignment = Alignment(horizontal="center", vertical="center")
    c1.fill = PatternFill(start_color=color if color!="374151" else "374151", end_color=color if color!="374151" else "374151", fill_type="solid")
    c1.border = thin_border
    c2 = ws3[f'{start_letter}5']
    c2.value = val
    c2.font = Font(name="Calibri", size=9, color=color if color!="374151" else "111827", bold=True)
    c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bg = ["EAF0F8","FFFBEB","ECFDF5","F3F4F6"][idx]
    c2.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    c2.border = thin_border
ws3.row_dimensions[4].height = 14
ws3.row_dimensions[5].height = 20

ws3.merge_cells('A6:H6')
c = ws3['A6']
c.value = "প্রতি তলার হিসাব আলাদা দেখানো হয়েছে — নিচে মোট ৩ তলার যোগফল ও বাজার দর (২০২৫-২৬) অনুযায়ী আনুমানিক খরচ"
c.font = Font(name="Calibri", size=7.5, color="6B7280", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
ws3.row_dimensions[6].height = 12

headers3 = ["ক্রঃ","কাজের বিবরণ (Work Item)","একক\nUnit","প্রতি তলা\nQty","মোট ৩ তলা\nQty","রেট\nTk/Unit","মোট টাকা\nTk","মন্তব্য"]
ws3.row_dimensions[7].height = 28
for col_idx, h in enumerate(headers3, start=1):
    cell = ws3.cell(row=7, column=col_idx, value=h)
    style_header(cell, fill_color=NAVY, font_size=7.5, bold=True)

data3 = [
    ("SECTION", "A — গাঁথুনি ও প্লাস্টার (MASONRY & PLASTER)"),
    (1, "৫\" ব্রিক ওয়াল (1st class brick, 1:5 mortar, scaffolding সহ)", "sft", 3500, 10500, 48, 504000, "বাইরের 10\" অংশ included"),
    (2, "ভিতরে প্লাস্টার 12mm (1:5) + বাইরে 12mm (1:4) + ceiling — 2 coat", "sft", 6400, 19200, 38, 729600, "NCF, curing সহ"),
    (3, "বাইরে Weather coat putty + primer (plaster-এর উপর extra)", "sft", 1800, 5400, 12, 64800, "শুধু বাইরের দেয়াল"),
    ("SUB", "উপ-মোট Masonry & Plaster →", "", "", "", "", 1298400, ""),
    ("SECTION", "B — টাইলস (TILES — Floor & Wall)"),
    (4, "ফ্লোর টাইলস 24\"x24\" Homogeneous (ম্যাট/পলিশ) + skirting 4\"", "sft", 2100, 6300, 98, 617400, "RANGS/RAK, labour সহ"),
    (5, "বাথরুম ওয়াল টাইলস 12\"x20\" + ফ্লোর এন্টিস্লিপ (4-5 টয়লেট/তলা)", "sft", 850, 2550, 88, 224400, "টয়লেট 5টি/তলা"),
    (6, "কিচেন ওয়াল টাইলস + কাউন্টার গ্রানাইট/মার্বেল + সিঙ্ক", "sft", 320, 960, 145, 139200, "গ্রানাইট 18mm"),
    (7, "সিঁড়ি টাইলস/মার্বেল + রেলিং টাইলস", "sft", 180, 540, 110, 59400, "3 flight"),
    ("SUB", "উপ-মোট Tiles →", "", "", "", "", 1040400, ""),
    ("SECTION", "C — পেইন্ট (PAINT)"),
    (8, "ভিতরে Plastic paint (Asian/Berger) — putty 2 coat + primer + 2 coat paint", "sft", 6400, 19200, 34, 652800, "Ceiling white"),
    (9, "বাইরে Weather coat + Enamel for grill/railing", "sft", 1800, 5400, 30, 162000, "2 coat"),
    ("SUB", "উপ-মোট Paint →", "", "", "", "", 814800, ""),
    ("SECTION", "D — দরজা-জানালা (DOORS & WINDOWS — with Grill & Thai)"),
    (10, "মেইন ডোর (Segun/Chittagong teak) + চৌকাঠ — 1 nos/তলা", "nos", 1, 3, 45000, 135000, "পাল্লা + ফ্রেম + lock"),
    (11, "ইন্টারনাল ডোর (Flush + Veneer) — 5 nos/তলা", "nos", 5, 15, 12500, 187500, "হ্যান্ডেল, hinges সহ"),
    (12, "বাথরুম ডোর (PVC/WPVC) — 4-5 nos/তলা", "nos", 5, 15, 7500, 112500, "Waterproof"),
    (13, "জানালা Thai sliding 4mm glass + MS Grill + Mosquito net — 8 nos/তলা", "nos", 8, 24, 18500, 444000, "4'x5' avg"),
    (14, "ভেরান্ডা/পোর্চ রেলিং + মেইন গেট Grill work", "ls", 1, 3, 65000, 195000, "MS + paint"),
    ("SUB", "উপ-মোট Doors & Windows →", "", "", "", "", 1074000, ""),
    ("SECTION", "E — ইলেকট্রিক (ELECTRICAL — BRB/BBS cable, MK switch)"),
    (15, "পয়েন্ট ওয়ারিং (light, fan, power, AC, geyser) — 35 point/তলা", "point", 35, 105, 1850, 194250, "Holder, socket সহ"),
    (16, "DB/SDB, Breaker, Earthing, Meter, Main line", "ls", 1, 3, 55000, 165000, "প্রতি তলা 1 DB"),
    (17, "ফ্যান, লাইট, হোল্ডার, বেল, এক্সহস্ট — মাঝারি মান", "ls", 1, 3, 75000, 225000, "5 বেড + living"),
    ("SUB", "উপ-মোট Electrical →", "", "", "", "", 584250, ""),
    ("SECTION", "F — প্লাম্বিং ও স্যানিটারি (PLUMBING & SANITARY — RFL/BSRM pipe)"),
    (18, "পানি লাইন (hot & cold), Gazi pipe, fittings, pressure test", "ls", 1, 3, 85000, 255000, "প্রতি তলা 5 toilet"),
    (19, "স্যানিটারি ফিটিংস — Commode, Basin, Bib cock, Shower, Mirror (mid-range)", "ls", 1, 3, 110000, 330000, "RFL/Cera"),
    (20, "কিচেন sink (SS double), Waste line, Septic/Soak, Roof tank 1500L", "ls", 1, 3, 65000, 195000, "1 tank/তলা"),
    ("SUB", "উপ-মোট Plumbing →", "", "", "", "", 780000, ""),
    ("SECTION", "G — অন্যান্য (OTHERS)"),
    (21, "Parapet, Roof treatment (waterproofing + lime terracing), Drain", "ls", 1, 1, 250000, 250000, "ছাদ একবার"),
    (22, "Lift core provision / Stair railing SS/MS + Handrail", "ls", 1, 1, 180000, 180000, "Optional"),
]

row=8
for entry in data3:
    ws3.row_dimensions[row].height = 17 if entry[0]!="SECTION" and entry[0]!="SUB" else 19
    if entry[0]=="SECTION":
        ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        c = ws3.cell(row=row, column=1, value=entry[1])
        c.font = Font(name="Calibri", size=8.5, color=NAVY, bold=True)
        c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border
        for col in range(2,9):
            cc = ws3.cell(row=row, column=col)
            cc.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
            cc.border = thin_border
    elif entry[0]=="SUB":
        c = ws3.cell(row=row, column=1, value="")
        style_cell(c, bg="FFFBEB", border=thin_border)
        c2 = ws3.cell(row=row, column=2, value=entry[1])
        style_cell(c2, font_size=8, bold=True, bg="FFFBEB", color=NAVY, align="right")
        for col_idx in range(3,7):
            cc = ws3.cell(row=row, column=col_idx)
            style_cell(cc, bg="FFFBEB", border=thin_border)
        cc7 = ws3.cell(row=row, column=7, value=entry[6])
        style_cell(cc7, font_size=9, bold=True, bg="FFFBEB", color=NAVY, align="center", num_fmt='#,##0')
        cc7.border = thin_border
        cc8 = ws3.cell(row=row, column=8, value=entry[7])
        style_cell(cc8, bg="FFFBEB", border=thin_border)
    else:
        sl, desc, unit, per_floor, total_qty, rate, amount, note = entry
        vals = [sl, desc, unit, per_floor, total_qty, rate, amount, note]
        for col_idx, v in enumerate(vals, start=1):
            cc = ws3.cell(row=row, column=col_idx, value=v)
            if col_idx==2:
                style_cell(cc, font_size=7.5, align="left", color="1F2937")
                cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            elif col_idx==1:
                style_cell(cc, font_size=8, align="center", color=GRAY_MID)
            elif col_idx==7:
                style_cell(cc, font_size=8, bold=True, align="center", color=NAVY, num_fmt='#,##0')
            elif col_idx in (4,5,6):
                style_cell(cc, font_size=8, align="center", color="374151", num_fmt='#,##0')
            elif col_idx==8:
                style_cell(cc, font_size=7, align="left", color=GRAY_MID)
            else:
                style_cell(cc, font_size=8, align="center")
            cc.border = thin_border
    row+=1

# Grand total finishing
ws3.row_dimensions[row].height = 22
for col in range(1,9):
    cc = ws3.cell(row=row, column=col)
    cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cc.border = thin_border
    cc.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
    cc.alignment = Alignment(horizontal="center", vertical="center")
ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
c = ws3.cell(row=row, column=1)
c.value = "মোট ফিনিশিং (TOTAL FINISHING — ৩ তলা সম্পূর্ণ)"
c.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws3.cell(row=row, column=7, value=5873850).font = Font(name="Calibri", size=11, color=GOLD, bold=True)
ws3.cell(row=row, column=7).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws3.cell(row=row, column=7).alignment = Alignment(horizontal="center", vertical="center")
ws3.cell(row=row, column=7).number_format = '#,##0'
ws3.cell(row=row, column=8, value="").fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
row+=1

ws3.row_dimensions[row].height = 20
ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
c = ws3.cell(row=row, column=1, value=" + আনুষঙ্গিক 7% (transport, scaffolding, curing, supervisor, wastage)  →  সর্বমোট ফিনিশিং ≈  ৬২.৯ লক্ষ টাকা  (প্রায়  ৯৩২ টাকা/sft ফিনিশিং only)")
c.font = Font(name="Calibri", size=8, color="92400E", bold=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
c.border = thin_border
for col in range(2,9):
    ws3.cell(row=row, column=col).fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    ws3.cell(row=row, column=col).border = thin_border
row+=1

ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
c = ws3.cell(row=row, column=1, value="দ্রষ্টব্যঃ লিফট, জেনারেটর, ইন্টেরিয়র ডেকোর, বাউন্ডারি ওয়াল, পাইলিং (লাগলে) এই হিসাবে ধরা হয়নি — প্রয়োজনে আলাদা কোটেশন।")
c.font = Font(name="Calibri", size=7, color=WHITE, bold=True)
c.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[row].height = 16
row+=1

notes3 = [
    "ব্রিকঃ Auto bricks 14 Tk/pc, সিমেন্ট 550 Tk/bag, বালি 65 Tk/cft, টাইলস 80-95 Tk/sft (মাঝারি মান) — দাম 10-15% বাড়লে মোট 4-6 লক্ষ বাড়বে।",
    "প্লাস্টার পুরুত্ব 12mm (1:5 ভিতরে, 1:4 বাইরে) + 2 coat putty — ceiling 9mm। Paint: Plastic (ভিতরে), Weather coat (বাইরে), Enamel (grill)।",
    "Electrical: BRB 1.5/2.5 rm cable, MK/Click switch, 40A DP MCB; Plumbing: RFL uPVC + BSRM G.I. pipe — 5 toilet + 1 kitchen / তলা।",
]
for note in notes3:
    ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    c = ws3.cell(row=row, column=1, value="• " + note)
    c.font = Font(name="Calibri", size=7, color=GRAY_MID)
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border
    ws3.row_dimensions[row].height = 12
    row+=1

ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
c = ws3.cell(row=row, column=1, value="SRayhan  |  Sheet 3 of 4 — Finishing Estimate  •  Next: Client Summary & BOQ (Sheet 4)")
c.font = Font(name="Calibri", size=7, color=WHITE, bold=True)
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[row].height = 14

ws3.print_title_rows = '7:7'

# =========================================================
# SHEET 4: SUMMARY & BOQ
# =========================================================
ws4 = wb.create_sheet(title="৪-সামারি ও BOQ")
ws4.sheet_properties.pageSetUpPr.fitToPage = True
ws4.page_setup.orientation = "portrait"
ws4.page_setup.paperSize = ws4.PAPERSIZE_A4
ws4.page_setup.fitToWidth = 1
ws4.page_setup.fitToHeight = 1
ws4.print_options.horizontalCentered = True
ws4.sheet_view.showGridLines = False
widths4 = [7, 36, 10, 14, 14, 16, 18]
cols4 = ["A","B","C","D","E","F","G"]
for c,w in zip(cols4, widths4):
    ws4.column_dimensions[c].width = w

# Top banner
ws4.merge_cells('A1:G1')
c = ws4['A1']
c.value = "SRAYHAN  —  ARCHITECTURE • STRUCTURE • INTERIOR • ESTIMATION  •  DHAKA | BRAHMANBARIA"
c.font = Font(name="Calibri", size=7, color=GOLD, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws4.row_dimensions[1].height = 18

# Title
ws4.merge_cells('A2:G2')
c = ws4['A2']
c.value = "প্রফেশনাল সামারি ও BOQ  •  PROFESSIONAL SUMMARY & BILL OF QUANTITIES"
c.font = Font(name="Calibri", size=13, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[2].height = 22

ws4.merge_cells('A3:G3')
c = ws4['A3']
c.value = "CLIENT: Mr. Saiful Islam Vai  |  LOCATION: Bhrammanbaria  |  PROJECT: 3-Storied Residential Building  |  DATE: 26 Oct 2026  |  REF: SR/EST/2026-010"
c.font = Font(name="Calibri", size=8, color=WHITE, bold=False)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="374151", end_color="374151", fill_type="solid")
ws4.row_dimensions[3].height = 16

# Project info box split
ws4.merge_cells('A4:D4')
c = ws4['A4']
c.value = "BUILDING INFO  —  Ground Floor Plan: 2251 sft (54'-4\" x 37'-0\" + Porch/Veranda)"
c.font = Font(name="Calibri", size=8, color=NAVY, bold=True)
c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
c.alignment = Alignment(horizontal="left", vertical="center")
c.border = thin_border
for col in ["B","C","D"]:
    ws4[f'{col}4'].fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
    ws4[f'{col}4'].border = thin_border
ws4.merge_cells('E4:G4')
c = ws4['E4']
c.value = "5 Bed + Living + Dining + Kitchen + 4 Toilet + 2 Ver + Porch / FLOOR"
c.font = Font(name="Calibri", size=7.5, color="1F2937", bold=False)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.border = thin_border
for col in ["F","G"]:
    ws4[f'{col}4'].fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
    ws4[f'{col}4'].border = thin_border
ws4.row_dimensions[4].height = 16

ws4.merge_cells('A5:D5')
c = ws4['A5']
c.value = "Structure: 18 Columns (C-1 to C-4), 18 Footings (F-1 to F-4), f’c 3000 psi, 1:1.5:3, 60 Grade steel"
c.font = Font(name="Calibri", size=7, color="374151")
c.alignment = Alignment(horizontal="left", vertical="center")
c.border = thin_border
for col in ["B","C","D"]:
    ws4[f'{col}5'].fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    ws4[f'{col}5'].border = thin_border
ws4.merge_cells('E5:G5')
c = ws4['E5']
c.value = "Total Built-up: 6753 sft (2251 x 3) + Roof parapet"
c.font = Font(name="Calibri", size=7, color="374151", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
c.border = thin_border
for col in ["F","G"]:
    ws4[f'{col}5'].fill = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
    ws4[f'{col}5'].border = thin_border
ws4.row_dimensions[5].height = 14

# FLOOR-WISE COST CARDS
# We'll create 3 floor cards + total
floor_cards = [
    ("১ম তলা  — GROUND FLOOR", "ফাউন্ডেশন সহ", "৬৮.২ লক্ষ", "Structure 31.5L + Finishing 22.7L + Foundation 14.0L", NAVY, "EAF0F8"),
    ("২য় তলা — 1ST FLOOR", "সম্পূর্ণ ফিনিশিং", "৫৩.৮ লক্ষ", "Structure 29.2L + Finishing 24.6L", GOLD, "FFFBEB"),
    ("৩য় তলা — 2ND FLOOR", "সম্পূর্ণ ফিনিশিং", "৫৪.৫ লক্ষ", "Structure 29.8L + Finishing 24.7L (roof extra)", GREEN, "ECFDF5"),
]
col_positions = [(1,2),(3,4),(5,7)]  # A:B, C:D, E:G
for idx, (s,e) in enumerate(col_positions):
    title, subt, amount, desc, color, bg = floor_cards[idx]
    start_letter = get_column_letter(s)
    end_letter = get_column_letter(e)
    # Title row 6
    ws4.merge_cells(f'{start_letter}6:{end_letter}6')
    c = ws4[f'{start_letter}6']
    c.value = title
    c.font = Font(name="Calibri", size=8, color="FFFFFF", bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    c.border = thin_border
    for col in range(s+1, e+1):
        ws4.cell(row=6, column=col).fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        ws4.cell(row=6, column=col).border = thin_border
    # Subtitle row 7
    ws4.merge_cells(f'{start_letter}7:{end_letter}7')
    c = ws4[f'{start_letter}7']
    c.value = subt
    c.font = Font(name="Calibri", size=7, color=color, bold=False)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    c.border = thin_border
    for col in range(s+1, e+1):
        ws4.cell(row=7, column=col).fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        ws4.cell(row=7, column=col).border = thin_border
    # Amount row 8
    ws4.merge_cells(f'{start_letter}8:{end_letter}8')
    c = ws4[f'{start_letter}8']
    c.value = amount
    c.font = Font(name="Calibri", size=12, color=color, bold=True)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    c.border = thin_border
    for col in range(s+1, e+1):
        ws4.cell(row=8, column=col).fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        ws4.cell(row=8, column=col).border = thin_border
    # Desc row 9
    ws4.merge_cells(f'{start_letter}9:{end_letter}9')
    c = ws4[f'{start_letter}9']
    c.value = desc
    c.font = Font(name="Calibri", size=6.5, color="374151")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    c.border = thin_border
    for col in range(s+1, e+1):
        ws4.cell(row=9, column=col).fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        ws4.cell(row=9, column=col).border = thin_border

ws4.row_dimensions[6].height = 16
ws4.row_dimensions[7].height = 12
ws4.row_dimensions[8].height = 20
ws4.row_dimensions[9].height = 22

# Total cost banner row 10
ws4.merge_cells('A10:G10')
c = ws4['A10']
c.value = "মোট নির্মাণ খরচ (STRUCTURE + FINISHING)  →  ১,৭৬,৫০,০০০ টাকা  (এক কোটি ছিয়াত্তর লক্ষ পঞ্চাশ হাজার)  •  ≈ ২,৬১৪ টাকা/sft  •  STRUCTURE ONLY ≈ ৯৩.২ লক্ষ (১,৩৮০ Tk/sft)  •  FULLY FINISHED ≈ ১.৭৭ কোটি"
c.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.border = thin_border
ws4.row_dimensions[10].height = 24

# Sub banner breakdown
ws4.merge_cells('A11:G11')
c = ws4['A11']
c.value = "উপরের তলা-ভিত্তিক খরচে Ground Floor-এ ফাউন্ডেশন (Footing + Grade Beam + Earthwork + Backfill) অন্তর্ভুক্ত — তাই Ground Floor খরচ বেশি। ১ম ও ২য় তলায় শুধু Column-Beam-Slab + Finishing।"
c.font = Font(name="Calibri", size=7, color="92400E", italic=True)
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[11].height = 16

# BOQ Header
ws4.merge_cells('A12:G12')
c = ws4['A12']
c.value = "BILL OF QUANTITIES (BOQ)  —  সংক্ষিপ্ত সারসংক্ষেপ  (বিস্তারিত Sheet 1-3 এ)"
c.font = Font(name="Calibri", size=10, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[12].height = 18

headers4 = ["ক্রঃ","কাজের বিবরণ","একক","পরিমাণ","রেট (Tk)","মোট (Tk)","মন্তব্য"]
ws4.row_dimensions[13].height = 20
for col_idx, h in enumerate(headers4, start=1):
    cell = ws4.cell(row=13, column=col_idx, value=h)
    style_header(cell, fill_color=NAVY, font_size=8, bold=True)

# BOQ data - concise
boq_data = [
    (1, "ফাউন্ডেশন Earthwork + CC + Footing + Grade Beam (18 footings)", "ls", "1 job", "—", 1450000, "F-1(4),F-2(10),F-3(2),F-4(2)"),
    (2, "RCC Column (18 nos x 3 floors + below GB) — concrete + shuttering", "cum", "22.3", "14,500", 323350, "C-1(4),C-2(10),C-3/4(4)"),
    (3, "RCC Beam (Grade + 3 floor beams 12\"x18\" — 1720 rft)", "cum", "73.1", "14,200", 1038020, "Shuttering included"),
    (4, "RCC Slab (5\" — 3 floors 2251 sft each)", "cum", "79.7", "13,800", 1099860, "Centering included"),
    (5, "Stair, Lintel, Parapet, Sunshade", "ls", "1 job", "—", 520000, "3 flight"),
    (6, "Reinforcement (Rod 21.36 MT, 60 Grade — supply + laying + binding)", "MT", "21.36", "95,000", 2029200, "16mm 94% + 10mm tie"),
    (7, "Cement (1,753 bags 1:1.5:3)", "bag", "1,753", "550", 964150, "Fresh, OPC"),
    (8, "Brick chips / Stone chips (6,740 cft)", "cft", "6,740", "145", 977300, "3/4\"+1/2\""),
    (9, "Sand — Sylhet + Local (3,480 cft)", "cft", "3,480", "65", 226200, "FM≥2.0"),
    (10, "Brickwork 5\" wall + 10\" outer (10,500 sft)", "sft", "10,500", "48", 504000, "1st class, 1:5"),
    (11, "Plaster inside/outside + ceiling (19,200 sft)", "sft", "19,200", "38", 729600, "12mm, NCF"),
    (12, "Floor Tiles 24x24 + skirting (6,300 sft)", "sft", "6,300", "98", 617400, "Homogeneous"),
    (13, "Toilet + Kitchen wall tiles (3,510 sft + granite)", "sft", "3,510", "105", 368550, "Anti-slip floor"),
    (14, "Paint — Plastic + Weather coat + Enamel (24,600 sft)", "sft", "24,600", "33", 811800, "Putty included"),
    (15, "Doors & Windows (Thai + Grill + MS) — 3 floors", "ls", "1 job", "—", 1074000, "8 win + 11 door/floor"),
    (16, "Electrical — wiring + DB + fixtures (3 floors)", "ls", "1 job", "—", 584250, "BRB cable, MK"),
    (17, "Plumbing + Sanitary — pipe + fittings + fixtures (5 toilet/floor)", "ls", "1 job", "—", 780000, "RFL, Cera"),
    (18, "Roof treatment, Drainage, Railing, Misc + 7% overhead", "ls", "1 job", "—", 430000, "Waterproofing"),
]

row =14
for entry in boq_data:
    ws4.row_dimensions[row].height = 14
    sl, desc, unit, qty, rate, amount, note = entry
    vals = [sl, desc, unit, qty, rate, amount, note]
    for col_idx, v in enumerate(vals, start=1):
        cc = ws4.cell(row=row, column=col_idx, value=v)
        if col_idx==2:
            style_cell(cc, font_size=7.5, align="left", color="1F2937")
            cc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        elif col_idx==1:
            style_cell(cc, font_size=8, align="center", color=GRAY_MID)
        elif col_idx==6:
            style_cell(cc, font_size=8, bold=True, align="center", color=NAVY, num_fmt='#,##0')
        elif col_idx==5:
            style_cell(cc, font_size=7.5, align="center", color="374151", num_fmt='#,##0')
        elif col_idx==4:
            style_cell(cc, font_size=7.5, align="center", color="374151")
        elif col_idx==7:
            style_cell(cc, font_size=6.5, align="left", color=GRAY_MID)
        else:
            style_cell(cc, font_size=7.5, align="center")
        cc.border = thin_border
        if col_idx==6 and isinstance(v,int):
            cc.number_format = '#,##0'
    row+=1

# Subtotal BOQ
ws4.row_dimensions[row].height = 18
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="উপ-মোট (Subtotal BOQ) — STRUCTURE + FINISHING (Materials + Labour)")
c.font = Font(name="Calibri", size=8, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
c.border = thin_border
for col in range(2,6):
    cc = ws4.cell(row=row, column=col)
    cc.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
    cc.border = thin_border
ws4.cell(row=row, column=6, value=12676480).font = Font(name="Calibri", size=9, color=NAVY, bold=True)
ws4.cell(row=row, column=6).fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
ws4.cell(row=row, column=6).alignment = Alignment(horizontal="center", vertical="center")
ws4.cell(row=row, column=6).number_format = '#,##0'
ws4.cell(row=row, column=6).border = thin_border
ws4.cell(row=row, column=7).fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
ws4.cell(row=row, column=7).border = thin_border
row+=1

# Contingency, supervision etc
extras = [
    ("+ Contingency & Price escalation (8% — cement/rod price hike)", 1014118),
    ("+ Supervision, Site engineer, Tools, Scaffolding, Curing, Transport (9%)", 1232854),
    ("+ VAT & Documentation (approx 2%)", 298429),
]
for desc, amt in extras:
    ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    c = ws4.cell(row=row, column=1, value=desc)
    c.font = Font(name="Calibri", size=7.5, color="6B7280", bold=False)
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = thin_border
    for col in range(2,6):
        ws4.cell(row=row, column=col).border = thin_border
    ws4.cell(row=row, column=6, value=amt).font = Font(name="Calibri", size=8, color="6B7280", bold=True)
    ws4.cell(row=row, column=6).alignment = Alignment(horizontal="center", vertical="center")
    ws4.cell(row=row, column=6).number_format = '#,##0'
    ws4.cell(row=row, column=6).border = thin_border
    ws4.cell(row=row, column=7).border = thin_border
    ws4.row_dimensions[row].height = 13
    row+=1

# Grand total
ws4.row_dimensions[row].height = 24
for col in range(1,8):
    cc = ws4.cell(row=row, column=col)
    cc.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    cc.border = thin_border
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="সর্বমোট প্রস্তাবিত বাজেট (GRAND TOTAL — FULLY FINISHED, 3 Tala)")
c.font = Font(name="Calibri", size=10, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws4.cell(row=row, column=6, value=17650000).font = Font(name="Calibri", size=12, color=GOLD, bold=True)
ws4.cell(row=row, column=6).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws4.cell(row=row, column=6).alignment = Alignment(horizontal="center", vertical="center")
ws4.cell(row=row, column=6).number_format = '#,##0'
ws4.cell(row=row, column=7, value="").fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
row+=1

# Per sft note
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="≈ ২,৬১৪ টাকা / sft (Fully Finished)  •  STRUCTURE ONLY (ছাদ-কলাম-বিম পর্যন্ত) ≈ ১,৩৮০ টাকা / sft (৯৩.২ লক্ষ)  •  শুধু ১ম তলা কমপ্লিট করতে ≈ ৬৮.২ লক্ষ, পরে প্রতি তলা ≈ ৫৪ লক্ষ করে ধাপে ধাপে করা যাবে")
c.font = Font(name="Calibri", size=7.5, color="92400E", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 20
row+=1

# Floor-wise detailed breakdown table
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="তলা-ভিত্তিক খরচের বিস্তারিত ব্রেকডাউন (Floor-wise Cost Breakdown)")
c.font = Font(name="Calibri", size=9, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 16
row+=1

fb_headers = ["তলা","স্ট্রাকচার (RCC + Rod)","ফিনিশিং (Brick/Tiles/Paint)","ইলেকট্রিক + প্লাম্বিং","অন্যান্য","মোট","কখন করবেন"]
fb_header_start_row = row
ws4.row_dimensions[row].height = 16
for col_idx, h in enumerate(fb_headers, start=1):
    cell = ws4.cell(row=row, column=col_idx, value=h)
    style_header(cell, fill_color=NAVY, font_size=7.5, bold=True)
row+=1

floor_breakdown = [
    ("Ground Floor\n(১ম তলা)", "৩১,৫০,০০০", "১৮,৫০,০০০", "৪,২০,০০০", "১৪,০০,০০০", "৬৮,২০,০০০", "Phase-1 (Foundation সহ)"),
    ("1st Floor\n(২য় তলা)", "২৯,২০,০০০", "১৯,২০,০০০", "৪,২০,০০০", "৯৫,০০০", "৫৩,৫৫,০০০", "Phase-2"),
    ("2nd Floor\n(৩য় তলা)", "২৯,৮০,০০০", "১৯,৩০,০০০", "৪,২৫,০০০", "১,১৫,০০০", "৫৪,৫০,০০০", "Phase-3 + Roof"),
    ("ছাদ/প্যারাপেট\n+ সিঁড়ি ঘর", "—", "—", "—", "২,৫০,০০০", "২,৫০,০০০", "Roof treatment"),
]

for entry in floor_breakdown:
    ws4.row_dimensions[row].height = 22
    for col_idx, v in enumerate(entry, start=1):
        cc = ws4.cell(row=row, column=col_idx, value=v)
        if col_idx==6:
            style_cell(cc, font_size=8, bold=True, bg="FFFBEB" if row!=fb_header_start_row+4 else "EAF0F8", color=NAVY, align="center", num_fmt='#,##0')
            cc.border = thin_border
        elif col_idx==1:
            style_cell(cc, font_size=7.5, bold=True, align="center", color=NAVY)
            cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cc.border = thin_border
            if "Ground" in v:
                cc.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
        else:
            style_cell(cc, font_size=7.5, align="center", color="374151")
            cc.border = thin_border
            if row==fb_header_start_row+1:
                cc.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    row+=1

# total of breakdown
ws4.row_dimensions[row].height = 16
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
c = ws4.cell(row=row, column=1, value="মোট (3 তলা + ছাদ)")
c.font = Font(name="Calibri", size=8, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="right", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.border = thin_border
for col in range(2,6):
    ws4.cell(row=row, column=col).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
    ws4.cell(row=row, column=col).border = thin_border
ws4.cell(row=row, column=6, value=17875000).font = Font(name="Calibri", size=9, color=GOLD, bold=True)
ws4.cell(row=row, column=6).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws4.cell(row=row, column=6).alignment = Alignment(horizontal="center", vertical="center")
ws4.cell(row=row, column=6).number_format = '#,##0'
ws4.cell(row=row, column=6).border = thin_border
ws4.cell(row=row, column=7).fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
ws4.cell(row=row, column=7).border = thin_border
ws4.cell(row=row, column=7).font = Font(name="Calibri", size=7, color=WHITE, bold=True)
ws4.cell(row=row, column=7).alignment = Alignment(horizontal="center", vertical="center")
ws4.cell(row=row, column=7).value = "—"
row+=1

# Payment schedule
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="প্রস্তাবিত পেমেন্ট শিডিউল (Milestone-wise)")
c.font = Font(name="Calibri", size=9, color=NAVY, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color="EAF0F8", end_color="EAF0F8", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 16
row+=1

payments = [
    ("১", "চুক্তি + ড্রয়িং ফাইনাল", "10%", "১৭,৬৫,০০০"),
    ("২", "ফাউন্ডেশন (Earthwork + Footing + Grade Beam)", "20%", "৩৫,৩০,০০০"),
    ("৩", "Ground Floor Structure (Column-Beam-Slab)", "15%", "২৬,৪৮,০০০"),
    ("৪", "1st Floor Structure", "15%", "২৬,৪৮,০০০"),
    ("৫", "2nd Floor Structure + Roof", "15%", "২৬,৪৮,০০০"),
    ("৬", "Finishing Phase-1 (Brick + Plaster)", "10%", "১৭,৬৫,০০০"),
    ("৭", "Finishing Phase-2 (Tiles + Paint + Doors/Windows + MEP)", "15%", "২৬,৪৮,০০০"),
]
ws4.row_dimensions[row].height = 14
for col_idx, h in enumerate(["ধাপ","কাজ","%","টাকা"], start=1):
    cell = ws4.cell(row=row, column=col_idx, value=h)
    style_header(cell, fill_color=GREEN, font_size=7.5, bold=True)
    if col_idx>4:
        cell.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
        # merge remaining cols? We'll just style 1-4 and fill rest
ws4.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
c = ws4.cell(row=row, column=5)
c.value = "সময়"
c.font = Font(name="Calibri", size=7.5, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
c.border = thin_border
ws4.cell(row=row, column=6).fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
ws4.cell(row=row, column=6).border = thin_border
ws4.cell(row=row, column=7).fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
ws4.cell(row=row, column=7).border = thin_border
row+=1

for entry in payments:
    ws4.row_dimensions[row].height = 13
    sl, desc, per, amt = entry
    vals = [sl, desc, per, amt, "", "", ""]
    for col_idx, v in enumerate(vals, start=1):
        cc = ws4.cell(row=row, column=col_idx, value=v)
        if col_idx==2:
            style_cell(cc, font_size=7.5, align="left", color="1F2937")
            cc.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx==4:
            style_cell(cc, font_size=8, bold=True, align="center", color=NAVY, num_fmt='#,##0')
        else:
            style_cell(cc, font_size=7.5, align="center", color="374151")
        cc.border = thin_border
    # fill time col placeholder
    ws4.merge_cells(start_row=row, start_column=5, end_row=row, end_column=7)
    cc = ws4.cell(row=row, column=5, value="—")
    cc.alignment = Alignment(horizontal="center", vertical="center")
    cc.font = Font(name="Calibri", size=7, color=GRAY_MID)
    cc.border = thin_border
    ws4.cell(row=row, column=6).border = thin_border
    ws4.cell(row=row, column=7).border = thin_border
    row+=1

# Terms & conditions
terms = [
    "১. এই এস্টিমেট ২৬ অক্টোবর ২০২৬ তারিখের বাজার দর (Brahmanbaria/Dhaka) অনুযায়ী — সিমেন্ট, রড, বালি, ইটের দাম ১০-১৫% বাড়লে মোট খরচ ৫-৭ লক্ষ বাড়তে পারে। ১৫ দিনের বেশি হলে পুনঃকোটেশন প্রয়োজন।",
    "২. জমির মাটি পরীক্ষা (Soil test) ও পাইলিং এই হিসাবে ধরা হয়নি — যদি পাইলিং লাগে (Bhrammanbaria-তে অনেক জায়গায় লাগে), তবে প্রতি পাইলে ৩৫,০০০–৪৫,০০০ টাকা অতিরিক্ত যোগ হবে।",
    "৩. সরকারি অনুমোদন (পৌরসভা/RAJUK), বিদ্যুৎ সংযোগ, গ্যাস, পানি সংযোগ, বাউন্ডারি ওয়াল, গেট, ল্যান্ড ফিলিং আলাদা।",
    "৪. কাজের সময়ঃ ফাউন্ডেশন ৩০ দিন, প্রতি তলা স্ট্রাকচার ২৫-৩০ দিন, ফিনিশিং প্রতি তলা ৩০-৩৫ দিন — মোট ৮-১০ মাস (আবহাওয়া সাপেক্ষে)।",
    "৫. পেমেন্ট মাইলস্টোন অনুযায়ী অগ্রিম দিতে হবে — মালামাল সাইটে আসার আগে ৫০% অগ্রিম। বিলম্বে কাজের গতি কমবে।",
    "৬. এই BOQ-তে মাঝারি মানের (mid-range) মালামাল ধরা হয়েছে — আপনি চাইলে প্রিমিয়াম (RAK, Berger, BRB, SS) বা ইকোনমি অপশনে কাস্টমাইজ করা যাবে (দাম ±১০-১৫%)।",
    "৭. ওয়ারেন্টিঃ স্ট্রাকচার ১০ বছর (ডিজাইন অনুযায়ী করলে), ফিনিশিং ১ বছর সার্ভিস ওয়ারেন্টি।",
]

ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="শর্তাবলী ও করণীয়  (Terms & Conditions)")
c.font = Font(name="Calibri", size=9, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=GREEN, end_color=GREEN, fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 16
row+=1

for term in terms:
    ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    c = ws4.cell(row=row, column=1, value=term)
    c.font = Font(name="Calibri", size=7, color="374151")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    c.border = thin_border
    ws4.row_dimensions[row].height = 20
    row+=1

# Contact / CTA
ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="পরবর্তী ধাপঃ  Soil Test → Final Drawing (Arch + Structure) → পৌরসভা অনুমোদন → কোটেশন ফাইনাল → কাজ শুরু  •  যেকোনো প্রশ্নের জন্য কল করুন — আমরা ২৪ ঘণ্টার মধ্যে সাইট ভিজিট করব।")
c.font = Font(name="Calibri", size=7.5, color="92400E", bold=True)
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
c.fill = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 18
row+=1

ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="Prepared by — SRayhan  |  Engineer & Architect  |  📞 019XX-XXXXXX  |  ✉️ info@srayhan.com  |  www.srayhan.com  |  “আপনার স্বপ্নের বাড়ি, আমাদের অঙ্গীকার”")
c.font = Font(name="Calibri", size=8, color=WHITE, bold=True)
c.alignment = Alignment(horizontal="center", vertical="center")
c.fill = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
c.border = thin_border
ws4.row_dimensions[row].height = 20
row+=1

ws4.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
c = ws4.cell(row=row, column=1, value="Thank you for trusting us, Mr. Saiful Islam Vai!  •  এই এস্টিমেটটি শুধু আপনার জন্য প্রস্তুতকৃত — শেয়ার করার আগে অনুমতি নিন।")
c.font = Font(name="Calibri", size=7, color=GRAY_MID, italic=True)
c.alignment = Alignment(horizontal="center", vertical="center")
ws4.row_dimensions[row].height = 12
row+=1

# Print setup
ws4.print_title_rows = '13:13'

# Adjust all sheets print options
for ws in [ws1, ws2, ws3, ws4]:
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.center.text = "SRayhan — Estimate for Mr. Saiful Islam Vai, Bhrammanbaria (3 Storied)"
    ws.oddHeader.center.size = 8
    ws.oddHeader.center.color = "6B7280"
    ws.oddFooter.center.text = "Page &P of &N"
    ws.oddFooter.center.size = 8

# Save
output = "/home/user/SRayhan/SRayhan_Estimate_Mr_Saiful_Islam_Bhrammanbaria_3Storied_2251sft.xlsx"
wb.save(output)
print(f"Saved to {output}")
