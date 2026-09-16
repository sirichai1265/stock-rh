# -*- coding: utf-8 -*-
"""Stock vs Booking report in the 'TOTAL RH' style from REEFER STOCK-FORM.xls,
for BKK27 (BC2) and LCH27 (HAST) only - all on ONE sheet ('TOTAL RH'):
  1) weekly Stock/Booking/AV-Balance block (20'RE / 40'RH only) per depot
  2) a mini Stock/Booking/Balance (22RE|45RE) roll-up table
  3) a "YEAR BUILT 5 YEARS FOR DURIAN SEASON" mini table (45RE by brand) per depot
  4) the full Built-Year x Brand x Size stock matrix per depot
"""
import datetime as _dt
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter

STOCK = r"C:\Users\HAL-USER\Desktop\9-14-STAYING.xls"
BKG = r"C:\Users\HAL-USER\Desktop\9-14-BKG+PD.xls"
OUT = r"C:\Users\HAL-USER\AppData\Local\Temp\claude\C--Users-HAL-USER-Desktop-STOCK-RH\86e51dd9-7ff4-4213-8bd8-252bb8ffc230\scratchpad\Stock_Daily_Reefer_9-14.xlsx"
REPORT_DATE = _dt.date(2026, 9, 14)  # filenames say 9-14; system clock is behind, so pin it

LOCS = {"BKK27": "BKK27 / BC2", "LCH27": "LCH27 / HAST"}
DEPOT_TITLE = {"BKK27": "BKK / BC2 (BKK27)", "LCH27": "LCH / HAST (LCH27)"}
LAST = 8000

TYPES = [("20'RE", "22RE", "RE22"), ("40'RH", "45RE", "RE45")]
BRANDS = ["CARRIER", "DAIKIN", "THERMO KING"]
SIZES = ["22RE", "45RE"]


def norm_brand(v):
    s = str(v).upper().strip()
    if "CARR" in s:
        return "CARRIER"
    if "DAIKIN" in s:
        return "DAIKIN"
    if "THERMO" in s or s == "TK":
        return "THERMO KING"
    return None


# ---------------- stock ----------------
st = pd.read_excel(STOCK, header=0)
st.columns = [str(c).strip() for c in st.columns]
st = st[st["Location"].isin(LOCS) & st["Size/Type"].notna()]

stock = {loc: {code: int((st[(st["Location"] == loc) & (st["Size/Type"] == code)]).shape[0])
               for _, code, _ in TYPES} for loc in LOCS}
print("STOCK", stock)

# ---------------- booking ----------------
bk = pd.read_excel(BKG, header=0)
bk.columns = [str(c).strip() for c in bk.columns]
if not {"Pickup", "TRAN DT"} <= set(bk.columns):
    bk = pd.read_excel(BKG, header=1)
    bk.columns = [str(c).strip() for c in bk.columns]
bk = bk[bk["Pickup"].isin(LOCS)].copy()
bk["date"] = pd.to_datetime(
    bk["TRAN DT"].astype(str).str.replace(".0", "", regex=False),
    format="%Y%m%d", errors="coerce")
for _, _, bcode in TYPES:
    bk[bcode] = pd.to_numeric(bk.get(bcode, 0), errors="coerce").fillna(0).astype(int)
bk = bk[bk["date"].notna()]
rows = bk[["Pickup", "date"] + [b for _, _, b in TYPES]].values.tolist()
print("booking rows", len(rows))

# ---------------- year x brand x size (reefer only) ----------------
stR = st[st["Size/Type"].isin(SIZES)].copy()
stR["BrandN"] = stR["RF Brand"].map(norm_brand)
stR["Year"] = pd.to_numeric(stR["Built Year"], errors="coerce")
YEARS_FIXED = list(range(2026, 2009, -1))  # fixed "YEAR BUILT" range 2026-2010, per user spec

year_tables = {}
for loc in LOCS:
    sub = stR[stR["Location"] == loc]
    years = YEARS_FIXED
    tbl = []
    for y in years:
        row = {"Year": y}
        for b in BRANDS:
            for sz in SIZES:
                row[(b, sz)] = int(((sub["Year"] == y) & (sub["BrandN"] == b) & (sub["Size/Type"] == sz)).sum())
        tbl.append(row)
    gtot = {"Year": "GTTL"}
    for b in BRANDS:
        for sz in SIZES:
            gtot[(b, sz)] = sum(r[(b, sz)] for r in tbl)
    tbl.append(gtot)
    year_tables[loc] = tbl

# ================= workbook =================
wb = openpyxl.Workbook()

ctl = wb.active
ctl.title = "Control"
ctl["A1"] = "Report date"; ctl["B1"] = REPORT_DATE; ctl["B1"].number_format = "yyyy-mm-dd"
ctl["A3"] = "WK#"; ctl["B3"] = "Monday"; ctl["C3"] = "Saturday"
ctl["A4"] = "=WEEKNUM(B4,21)"; ctl["B4"] = "=$B$1-(WEEKDAY($B$1,2)-1)"; ctl["C4"] = "=B4+5"
ctl["A5"] = "=WEEKNUM(B5,21)"; ctl["B5"] = "=B4+7"; ctl["C5"] = "=B5+5"
ctl["A6"] = "=WEEKNUM(B6,21)"; ctl["B6"] = "=B5+7"; ctl["C6"] = "=B6+5"
ctl["A7"] = "=WEEKNUM(B7,21)"; ctl["B7"] = "=B6+7"; ctl["C7"] = "=B7+5"
ctl["A9"] = "Tomorrow"; ctl["B9"] = "=B1+1"
for c in ["B4", "C4", "B5", "C5", "B6", "C6", "B7", "C7", "B9"]:
    ctl[c].number_format = "yyyy-mm-dd"

br = wb.create_sheet("BookingRaw")
br.append(["Location", "Pickup Date"] + [t[0] for t in TYPES])
for r in rows:
    loc, d = r[0], r[1]
    qtys = r[2:]
    br.append([loc, d.to_pydatetime()] + [int(q) for q in qtys])
for r in range(2, br.max_row + 1):
    br.cell(r, 2).number_format = "yyyy-mm-dd"

type_cols = [chr(ord("C") + i) for i in range(len(TYPES))]  # BookingRaw cols C, D


def mon(d):
    return ('CHOOSE(MONTH(%s),"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG",'
            '"SEP","OCT","NOV","DEC")') % d


def sif(col, loc, crit):
    return ('=SUMIFS(BookingRaw!$%s$2:$%s$%d,BookingRaw!$A$2:$A$%d,"%s",%s)'
            % (col, col, LAST, LAST, loc, crit))


def dc(op, ref):
    return 'BookingRaw!$B$2:$B$%d,"%s"&%s' % (LAST, op, ref)


def rng_label(b, c):
    return ('="Booking on "&IF(%s=%s,DAY(%s)&" "&%s,'
            'IF(MONTH(%s)=MONTH(%s),DAY(%s)&"-"&DAY(%s)&" "&%s,'
            'DAY(%s)&" "&%s&"-"&DAY(%s)&" "&%s))'
            % (b, c, b, mon(b), b, c, b, c, mon(c), b, mon(b), c, mon(c)))


def av_label(c):
    return '="AV Balance till "&DAY(%s)&" "&%s' % (c, mon(c))


def today_label():
    return '="Booking on "&DAY(Control!$B$1)&" "&' + mon("Control!$B$1")


# --------------------------------------------------------------------- #
# Day-dependent week-bucketing rule (user-specified, confirmed 2026-09-13):
#   Mon-Thu ("rule 1"): Count1 = today alone; Count2 = tomorrow..Sat this
#     week; then 2 forward weeks labelled Week1 (=calendar wk+1), Week2 (=wk+2)
#   Fri ("rule 2"): Count1 = today..Sat this week merged (no separate "today"
#     row, since only Fri+Sat remain); then 3 forward weeks Week1/2/3 (=wk+1..+3)
# Sat/Sun aren't specified by the user - fall back to rule 2 (merge whatever
# is left of this week into Count1) as the closer analogue; flagged below.
# --------------------------------------------------------------------- #
_weekday = REPORT_DATE.weekday()  # Mon=0 .. Sun=6
if _weekday <= 3:
    RULE = 1
elif _weekday == 4:
    RULE = 2
else:
    RULE = 2
    print("NOTE: report date is a weekend day - user's rule only covers Mon-Thu/Fri; "
          "falling back to the Friday rule (merged Count1, 3 forward weeks).")
print("weekday=%s -> RULE %d" % (REPORT_DATE.strftime("%A"), RULE))

sm = wb.create_sheet("TOTAL RH")
sm.sheet_view.showGridLines = False
N = len(TYPES)
START_COL = {"BKK27": 2, "LCH27": 8}  # label col B.., H.. ; WK col sits one to the left

sm["A1"] = ('="Stock vs Booking Report - Reefer (20\'RE / 40\'RH)  as of "&DAY(Control!$B$1)&" "'
            '&' + mon("Control!$B$1") + '&" "&YEAR(Control!$B$1)')
_last_col = max(START_COL.values()) + N + 1
sm.merge_cells(start_row=1, start_column=1, end_row=1, end_column=_last_col)

block_bottom = {}
booking_rows = {}   # loc -> list of row numbers holding a booking SUMIFS (not AV, not pending)
pending_row = {}
last_av_row = {}

for loc, label in LOCS.items():
    base = START_COL[loc]
    wk_col_idx = base - 1
    sm.cell(2, base, label)
    sm.merge_cells(start_row=2, start_column=base, end_row=2, end_column=base + N)
    for i, (disp, _, _) in enumerate(TYPES):
        sm.cell(3, base + 1 + i, disp)
    sm.cell(3, wk_col_idx, "WK")

    r = 4
    sm.cell(r, base, "Stock empty in yard")
    for i, (_, scode, _) in enumerate(TYPES):
        sm.cell(r, base + 1 + i, stock[loc][scode])

    r = 5
    sm.cell(r, wk_col_idx, "=Control!$A$4")
    sm.cell(r, base, "Booking Pending pick up")
    for i, colL in enumerate(type_cols):
        sm.cell(r, base + 1 + i, sif(colL, loc, dc("<", "Control!$B$1")))
    pending_row[loc] = r
    bkrows = []
    stock_row = 4

    if RULE == 1:
        r = 6  # Count1 = today only
        sm.cell(r, wk_col_idx, "=Control!$A$4")
        sm.cell(r, base, today_label())
        for i, colL in enumerate(type_cols):
            sm.cell(r, base + 1 + i, sif(colL, loc, dc("=", "Control!$B$1")))
        bkrows.append(r)

        r = 7  # Count2 = tomorrow .. Sat wk1
        sm.cell(r, wk_col_idx, "=Control!$A$4")
        sm.cell(r, base, rng_label("Control!$B$9", "Control!$C$4"))
        for i, colL in enumerate(type_cols):
            sm.cell(r, base + 1 + i, sif(colL, loc, dc(">=", "Control!$B$9") + "," + dc("<=", "Control!$C$4")))
        bkrows.append(r)

        r = 8  # AV till Sat wk1
        sm.cell(r, base, av_label("Control!$C$4"))
        for i in range(N):
            cL = col_letter(base + 1 + i)
            sm.cell(r, base + 1 + i, "=%s%d-%s%d-%s%d-%s%d" % (cL, stock_row, cL, pending_row[loc], cL, 6, cL, 7))
        av_prev = r

        fwd_ctrl = [("Control!$B5", "Control!$C5", "Control!$A$5"), ("Control!$B6", "Control!$C6", "Control!$A$6")]
        r = 9
        for mon_ref, sat_ref, wk_ref in fwd_ctrl:
            sm.cell(r, wk_col_idx, "=%s" % wk_ref)
            sm.cell(r, base, rng_label(mon_ref, sat_ref))
            for i, colL in enumerate(type_cols):
                sm.cell(r, base + 1 + i, sif(colL, loc, dc(">=", mon_ref) + "," + dc("<=", sat_ref)))
            bkrows.append(r)
            r += 1
            sm.cell(r, base, av_label(sat_ref))
            for i in range(N):
                cL = col_letter(base + 1 + i)
                sm.cell(r, base + 1 + i, "=%s%d-%s%d" % (cL, av_prev, cL, r - 1))
            av_prev = r
            r += 1
        bottom = av_prev

    else:  # RULE == 2
        r = 6  # Count1 = today..Sat wk1 merged
        sm.cell(r, wk_col_idx, "=Control!$A$4")
        sm.cell(r, base, rng_label("Control!$B$1", "Control!$C$4"))
        for i, colL in enumerate(type_cols):
            sm.cell(r, base + 1 + i, sif(colL, loc, dc(">=", "Control!$B$1") + "," + dc("<=", "Control!$C$4")))
        bkrows.append(r)

        r = 7  # AV till Sat wk1
        sm.cell(r, base, av_label("Control!$C$4"))
        for i in range(N):
            cL = col_letter(base + 1 + i)
            sm.cell(r, base + 1 + i, "=%s%d-%s%d-%s%d" % (cL, stock_row, cL, pending_row[loc], cL, 6))
        av_prev = r

        fwd_ctrl = [("Control!$B5", "Control!$C5", "Control!$A$5"),
                    ("Control!$B6", "Control!$C6", "Control!$A$6"),
                    ("Control!$B7", "Control!$C7", "Control!$A$7")]
        r = 8
        for mon_ref, sat_ref, wk_ref in fwd_ctrl:
            sm.cell(r, wk_col_idx, "=%s" % wk_ref)
            sm.cell(r, base, rng_label(mon_ref, sat_ref))
            for i, colL in enumerate(type_cols):
                sm.cell(r, base + 1 + i, sif(colL, loc, dc(">=", mon_ref) + "," + dc("<=", sat_ref)))
            bkrows.append(r)
            r += 1
            sm.cell(r, base, av_label(sat_ref))
            for i in range(N):
                cL = col_letter(base + 1 + i)
                sm.cell(r, base + 1 + i, "=%s%d-%s%d" % (cL, av_prev, cL, r - 1))
            av_prev = r
            r += 1
        bottom = av_prev

    booking_rows[loc] = bkrows
    last_av_row[loc] = av_prev
    block_bottom[loc] = bottom

    sm.column_dimensions[col_letter(base)].width = 20   # label column
    for i in range(1, N + 1):
        sm.column_dimensions[col_letter(base + i)].width = 8  # value columns
    sm.column_dimensions[col_letter(wk_col_idx)].width = 4

BLOCK_BOTTOM = max(block_bottom.values())

# ---------------- styling: weekly blocks ----------------
WHITE = Font(color="FFFFFF", bold=True)
BOLD = Font(bold=True)
banner = PatternFill("solid", fgColor="0D3B12")
green = PatternFill("solid", fgColor="2E7D32")
grey = PatternFill("solid", fgColor="F2F2F2")
yellow = PatternFill("solid", fgColor="FFF59D")
ctr = Alignment(horizontal="center")
lft = Alignment(horizontal="left")

sm["A1"].fill = banner; sm["A1"].font = WHITE; sm["A1"].alignment = ctr

for loc, label in LOCS.items():
    base = START_COL[loc]
    wk_col_idx = base - 1
    cell = sm.cell(2, base); cell.fill = green; cell.font = WHITE; cell.alignment = ctr
    for c in range(base, base + N + 1):
        cell = sm.cell(3, c); cell.fill = green; cell.font = WHITE; cell.alignment = ctr
    sm.cell(3, wk_col_idx).fill = green; sm.cell(3, wk_col_idx).font = WHITE; sm.cell(3, wk_col_idx).alignment = ctr
    for c in range(base, base + N + 1):
        cell = sm.cell(4, c); cell.fill = grey; cell.font = BOLD; cell.alignment = ctr if c > base else lft
    # AV-Balance rows are every row from Count/Week's AV onward - identify them as
    # "the row after each booking row" among this block's rows
    av_rows_this_block = sorted(set(range(4, block_bottom[loc] + 1)) - set(booking_rows[loc])
                                 - {4, pending_row[loc]})
    for rr in av_rows_this_block:
        for c in range(base, base + N + 1):
            cell = sm.cell(rr, c); cell.fill = yellow; cell.font = BOLD; cell.alignment = ctr if c > base else lft

wb.calculation.fullCalcOnLoad = True

# ---------------- mini Stock/Booking/Balance (22RE|45RE) roll-up ----------------
MINI_COL = max(START_COL.values()) + N + 3  # a couple columns after the LCH27 block
DEPOT_ROWLABEL = {"BKK27": "BC", "LCH27": "HAST"}
mini_titles = [("STOCK", 4), ("BOOKING", None), ("BALANCE", None)]
r0 = 2
for title, _ in mini_titles:
    sm.cell(r0, MINI_COL, title).font = BOLD
    sm.cell(r0, MINI_COL).fill = grey
    sm.cell(r0, MINI_COL + 1, "22RE").font = BOLD; sm.cell(r0, MINI_COL + 1).alignment = ctr
    sm.cell(r0, MINI_COL + 2, "45RE").font = BOLD; sm.cell(r0, MINI_COL + 2).alignment = ctr
    r0 += 1
    for loc in LOCS:
        sm.cell(r0, MINI_COL, DEPOT_ROWLABEL[loc])
        base = START_COL[loc]
        re_col, rh_col = col_letter(base + 1), col_letter(base + 2)
        if title == "STOCK":
            sm.cell(r0, MINI_COL + 1, "=%s4" % re_col)
            sm.cell(r0, MINI_COL + 2, "=%s4" % rh_col)
        elif title == "BOOKING":
            book_terms_re = "+".join("%s%d" % (re_col, rr) for rr in [pending_row[loc]] + booking_rows[loc])
            book_terms_rh = "+".join("%s%d" % (rh_col, rr) for rr in [pending_row[loc]] + booking_rows[loc])
            sm.cell(r0, MINI_COL + 1, "=%s" % book_terms_re)
            sm.cell(r0, MINI_COL + 2, "=%s" % book_terms_rh)
        else:
            sm.cell(r0, MINI_COL + 1, "=%s%d" % (re_col, last_av_row[loc]))
            sm.cell(r0, MINI_COL + 2, "=%s%d" % (rh_col, last_av_row[loc]))
        sm.cell(r0, MINI_COL + 1).alignment = ctr
        sm.cell(r0, MINI_COL + 2).alignment = ctr
        r0 += 1
    lastloc = list(LOCS)[-1]
    firstrow = r0 - len(LOCS)
    sm.cell(r0, MINI_COL, "TOTAL").font = BOLD
    sm.cell(r0, MINI_COL + 1, "=SUM(%s%d:%s%d)" % (col_letter(MINI_COL + 1), firstrow, col_letter(MINI_COL + 1), r0 - 1)).font = BOLD
    sm.cell(r0, MINI_COL + 2, "=SUM(%s%d:%s%d)" % (col_letter(MINI_COL + 2), firstrow, col_letter(MINI_COL + 2), r0 - 1)).font = BOLD
    sm.cell(r0, MINI_COL + 1).alignment = ctr; sm.cell(r0, MINI_COL + 2).alignment = ctr
    r0 += 2

for cc in (MINI_COL, MINI_COL + 1, MINI_COL + 2):
    sm.column_dimensions[col_letter(cc)].width = 9

# ---------------- "5 years for durian season" mini table (45RE by brand) ----------------
DUR_ROW0 = BLOCK_BOTTOM + 3
for loc, title in DEPOT_TITLE.items():
    base = START_COL[loc]
    last_col = base + 3  # Year + 3 brands = 4 cols
    sm.merge_cells(start_row=DUR_ROW0, start_column=base, end_row=DUR_ROW0, end_column=last_col)
    t = sm.cell(DUR_ROW0, base, title)
    t.fill = PatternFill("solid", fgColor="F4B183"); t.font = BOLD; t.alignment = ctr

    hdr_r = DUR_ROW0 + 1
    sm.merge_cells(start_row=hdr_r, start_column=base, end_row=hdr_r + 1, end_column=base)
    h0 = sm.cell(hdr_r, base, "YEAR BUILT 5 YEARS\nFOR DURIAN SEASON")
    h0.fill = green; h0.font = WHITE; h0.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, b in enumerate(BRANDS):
        hc = sm.cell(hdr_r, base + 1 + i, b)
        hc.fill = green; hc.font = WHITE; hc.alignment = ctr
        sc = sm.cell(hdr_r + 1, base + 1 + i, "45RE")
        sc.fill = green; sc.font = WHITE; sc.alignment = ctr

    years5 = list(range(REPORT_DATE.year, REPORT_DATE.year - 6, -1))  # this yr + previous 5
    year_lookup = {row["Year"]: row for row in year_tables[loc] if row["Year"] != "GTTL"}
    r = hdr_r + 2
    sums = [0, 0, 0]
    for y in years5:
        sm.cell(r, base, y).alignment = ctr
        row = year_lookup.get(y)
        for i, b in enumerate(BRANDS):
            v = row[(b, "45RE")] if row else 0
            sm.cell(r, base + 1 + i, v).alignment = ctr
            sums[i] += v
        r += 1
    sm.cell(r, base, "TOTAL").font = BOLD; sm.cell(r, base).alignment = ctr
    for i in range(3):
        c = sm.cell(r, base + 1 + i, sums[i]); c.font = BOLD; c.alignment = ctr
    over5 = stock[loc]["45RE"] - sum(sums)
    r += 1
    sm.cell(r, base + 3, over5).fill = yellow
    sm.cell(r, base + 2, "OVER 5 Y").alignment = Alignment(horizontal="right")
    DUR_BOTTOM = r

# ---------------- full Built-Year x Brand x Size stock matrix ----------------
YBS_ROW0 = DUR_BOTTOM + 3
TITLE_FILL = banner
HEAD_FILL = green
SUBHEAD_FILL = PatternFill("solid", fgColor="66BB6A")
GTTL_FILL = yellow
STRIPE_FILL = PatternFill("solid", fgColor="F7F7F7")

# Built-Year age bands (per user spec)
BAND_GREEN = PatternFill("solid", fgColor="A9D18E")   # 2026-2021
BAND_YELLOW = PatternFill("solid", fgColor="FFE699")  # 2020-2016
BAND_BROWN = PatternFill("solid", fgColor="C4A484")   # 2015-2011
BAND_RED = PatternFill("solid", fgColor="F4A7A7")     # 2010


def year_band_fill(y):
    if not isinstance(y, int):
        return None
    if 2021 <= y <= 2026:
        return BAND_GREEN
    if 2016 <= y <= 2020:
        return BAND_YELLOW
    if 2011 <= y <= 2015:
        return BAND_BROWN
    if y <= 2010:
        return BAND_RED
    return None
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

N_COLS = 1 + len(BRANDS) * len(SIZES)
col0 = 2
for loc, label in LOCS.items():
    last_col = col0 + N_COLS - 1

    sm.merge_cells(start_row=YBS_ROW0, start_column=col0, end_row=YBS_ROW0, end_column=last_col)
    t = sm.cell(YBS_ROW0, col0, "%s  \u2014  Stock by Built Year x Brand x Size" % label)
    t.fill = TITLE_FILL; t.font = WHITE; t.alignment = ctr

    hr = YBS_ROW0 + 1
    sm.merge_cells(start_row=hr, start_column=col0, end_row=hr + 1, end_column=col0)
    c0 = sm.cell(hr, col0, "YEAR\nBUILT")
    c0.fill = HEAD_FILL; c0.font = WHITE; c0.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    c = col0 + 1
    for b in BRANDS:
        sm.merge_cells(start_row=hr, start_column=c, end_row=hr, end_column=c + 1)
        hc = sm.cell(hr, c, b)
        hc.fill = HEAD_FILL; hc.font = WHITE; hc.alignment = ctr
        sm.cell(hr, c + 1).fill = HEAD_FILL
        for j, sz in enumerate(("20'RE", "40'RH")):
            sc = sm.cell(hr + 1, c + j, sz)
            sc.fill = SUBHEAD_FILL; sc.font = WHITE; sc.alignment = ctr
        c += 2

    r = hr + 2
    for i, row in enumerate(year_tables[loc]):
        is_total = row["Year"] == "GTTL"
        yc = sm.cell(r, col0, row["Year"]); yc.alignment = ctr
        yc.font = BOLD if is_total else Font()
        c = col0 + 1
        for b in BRANDS:
            for sz in SIZES:
                v = row[(b, sz)]
                cell = sm.cell(r, c, v if v else None)
                cell.alignment = ctr
                if is_total:
                    cell.font = BOLD
                c += 1
        fill = GTTL_FILL if is_total else year_band_fill(row["Year"])
        for cc in range(col0, last_col + 1):
            cell = sm.cell(r, cc)
            cell.border = BORDER
            if fill:
                cell.fill = fill
        r += 1

    _cur = sm.column_dimensions[col_letter(col0)].width
    if not _cur or _cur < 9:
        sm.column_dimensions[col_letter(col0)].width = 9
    for cc in range(col0 + 1, last_col + 1):
        sm.column_dimensions[col_letter(cc)].width = 8
    sm.row_dimensions[hr].height = 26

    col0 = last_col + 2

sm.freeze_panes = "B4"

# ---------------- normalize font to Calibri 11 everywhere, row height 13 ----------------
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            f = cell.font
            cell.font = Font(name="Calibri", size=11, bold=f.bold, italic=f.italic,
                              color=f.color, underline=f.underline)
    for r in range(1, ws.max_row + 1):
        if ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = 13

# open the workbook showing TOTAL RH, not Control
_active_idx = wb.sheetnames.index("TOTAL RH")
wb.active = _active_idx
for i, ws in enumerate(wb.worksheets):
    ws.sheet_view.tabSelected = (i == _active_idx)

wb.save(OUT)
print("saved", OUT)
