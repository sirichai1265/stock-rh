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

STOCK = r"C:\Users\HAL-USER\Desktop\9-21-STAYING-RH.xls"
BKG = r"C:\Users\HAL-USER\Desktop\9-21-PD+BKG-3WK-RH.xls"
OUT = r"C:\Users\HAL-USER\AppData\Local\Temp\claude\C--Users-HAL-USER-Desktop-STOCK-RH\86e51dd9-7ff4-4213-8bd8-252bb8ffc230\scratchpad\Stock_Daily_Reefer_9-21.xlsx"
REPORT_DATE = _dt.date(2026, 9, 21)
MERGE_END_DATE = _dt.date(2026, 10, 18)  # fold WK41 (12-18 Oct) stray bookings into the last displayed week

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
# The pivot export's header row shifts around (extra blank/"Data" rows above
# it vary by day), so scan the first block of rows for the one that actually
# contains "Pickup" and "TRAN DT" instead of assuming a fixed offset.
_bk_raw = pd.read_excel(BKG, header=None, nrows=10)
_hdr_row = None
for _i in range(len(_bk_raw)):
    _vals = set(str(v).strip() for v in _bk_raw.iloc[_i].tolist())
    if {"Pickup", "TRAN DT"} <= _vals:
        _hdr_row = _i
        break
if _hdr_row is None:
    raise ValueError("could not find booking header row (Pickup/TRAN DT) in %s" % BKG)
bk = pd.read_excel(BKG, header=_hdr_row)
bk.columns = [str(c).strip() for c in bk.columns]
if "Sum of RE22" in bk.columns:
    # weekly pivot export: Pickup/WEEK are only filled on each group's first
    # row (subtotal/"Total" rows interleaved) and the sheet repeats the same
    # combined table again per-location in side-by-side column blocks with
    # duplicate header names (deduped by pandas to "Pickup.1"/"Pickup.2" etc.)
    # - use only the first, combined block and forward-fill the group labels.
    bk = bk[["Pickup", "TRAN DT", "Sum of RE22", "Sum of RE45"]].copy()
    bk["Pickup"] = bk["Pickup"].ffill()
    bk = bk.rename(columns={"Sum of RE22": "RE22", "Sum of RE45": "RE45"})
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

# combined 45RE-by-year total across both depots, for the "peak build year" stat
_year_45re_combined = {}
for y in YEARS_FIXED:
    total = 0
    for loc in LOCS:
        row = next(r for r in year_tables[loc] if r["Year"] == y)
        total += sum(row[(b, "45RE")] for b in BRANDS)
    _year_45re_combined[y] = total
PEAK_YEAR, PEAK_YEAR_UNITS = max(_year_45re_combined.items(), key=lambda kv: kv[1])
_peak_ties = [y for y, v in _year_45re_combined.items() if v == PEAK_YEAR_UNITS]
PEAK_YEAR_LABEL = " / ".join(str(y) for y in sorted(_peak_ties, reverse=True)) if len(_peak_ties) > 1 else str(PEAK_YEAR)

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

# ---------------- StockRaw: per-container Move Code, for the AV/DMG remarks ----------------
sr = wb.create_sheet("StockRaw")
sr.append(["Location", "Size/Type", "Move Code"])
st_raw = st[st["Size/Type"].isin(SIZES)]
for _, row in st_raw.iterrows():
    sr.append([row["Location"], row["Size/Type"], str(row["Move Code"])])
SR_LAST = max(sr.max_row, 2)

AV_CODES = ["IED", "IEP", "IER"]
DMG_CODE = "OER"


def av_formula(loc, sz):
    parts = [
        'COUNTIFS(StockRaw!$A$2:$A$%d,"%s",StockRaw!$B$2:$B$%d,"%s",StockRaw!$C$2:$C$%d,"%s")'
        % (SR_LAST, loc, SR_LAST, sz, SR_LAST, code) for code in AV_CODES
    ]
    return "=" + "+".join(parts)


def dmg_formula(loc, sz):
    return '=COUNTIFS(StockRaw!$A$2:$A$%d,"%s",StockRaw!$B$2:$B$%d,"%s",StockRaw!$C$2:$C$%d,"%s")' % (
        SR_LAST, loc, SR_LAST, sz, SR_LAST, DMG_CODE)


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

        merge_date_formula = "DATE(%d,%d,%d)" % (MERGE_END_DATE.year, MERGE_END_DATE.month, MERGE_END_DATE.day)
        fwd_ctrl = [
            ("Control!$B5", "Control!$C5", "Control!$A$5", "Control!$C5"),
            # last displayed week: label still reads "till Control!$C6" (e.g. 3 OCT) but the
            # SUMIFS upper bound is widened to MERGE_END_DATE, folding later bookings
            # (e.g. 5-11 Oct) into this row per user request rather than dropping them.
            ("Control!$B6", "Control!$C6", "Control!$A$6", merge_date_formula),
        ]
        r = 9
        for mon_ref, sat_label_ref, wk_ref, sat_value_ref in fwd_ctrl:
            sm.cell(r, wk_col_idx, "=%s" % wk_ref)
            sm.cell(r, base, rng_label(mon_ref, sat_label_ref))
            for i, colL in enumerate(type_cols):
                sm.cell(r, base + 1 + i, sif(colL, loc, dc(">=", mon_ref) + "," + dc("<=", sat_value_ref)))
            bkrows.append(r)
            r += 1
            sm.cell(r, base, av_label(sat_label_ref))
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

# ---------------- Remarks: AV (move code IED/IEP/IER) and DMG (move code OER) ----------------
REM_ROW0 = BLOCK_BOTTOM + 2
sm.cell(REM_ROW0, 1, "Remarks:").font = Font(bold=True, italic=True)
sm.cell(REM_ROW0, 1).alignment = Alignment(horizontal="left")
note = sm.cell(REM_ROW0, 2, "AV = stock with move code IED / IEP / IER (available)   |   DMG = move code OER (damaged)")
note.font = Font(italic=True, size=9)
REM_HDR = REM_ROW0 + 1
for loc, label in LOCS.items():
    base = START_COL[loc]
    for i, (disp, scode, _) in enumerate(TYPES):
        c = sm.cell(REM_HDR, base + 1 + i, disp)
        c.font = BOLD; c.alignment = ctr
    sm.cell(REM_HDR + 1, base, "AV").alignment = lft
    sm.cell(REM_HDR + 2, base, "DMG").alignment = lft
    for i, (disp, scode, _) in enumerate(TYPES):
        cc = base + 1 + i
        sm.cell(REM_HDR + 1, cc, av_formula(loc, scode)).alignment = ctr
        sm.cell(REM_HDR + 2, cc, dmg_formula(loc, scode)).alignment = ctr
REM_BOTTOM = REM_HDR + 2

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
DUR_ROW0 = REM_BOTTOM + 3
DUR_TOTAL_ROW = {}
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
    DUR_TOTAL_ROW[loc] = r
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
YBS_COL0 = {}
YBS_HR = {}
for loc, label in LOCS.items():
    last_col = col0 + N_COLS - 1
    YBS_COL0[loc] = col0

    sm.merge_cells(start_row=YBS_ROW0, start_column=col0, end_row=YBS_ROW0, end_column=last_col)
    t = sm.cell(YBS_ROW0, col0, "%s  \u2014  Stock by Built Year x Brand x Size" % label)
    t.fill = TITLE_FILL; t.font = WHITE; t.alignment = ctr

    hr = YBS_ROW0 + 1
    YBS_HR[loc] = hr
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

# ---------------- grid lines: thin border on every populated/filled cell ----------------
# Native gridlines are off (showGridLines=False), so give every section its own
# ruled table instead of leaving blank-gap rows/columns bordered too.
for _row in sm.iter_rows(min_row=1, max_row=sm.max_row, min_col=1, max_col=sm.max_column):
    for _cell in _row:
        if _cell.value is not None or _cell.fill.fill_type is not None:
            _cell.border = BORDER

sm.freeze_panes = "B4"

# ==================== Dashboard sheet (KPI-card style) ====================
db = wb.create_sheet("Dashboard")
db.sheet_view.showGridLines = False

DB_LIGHT = Font(size=9, color="808080")
DB_LABEL = Font(bold=True, size=9, color="808080")
thin_card = Side(style="thin", color="D9D9D9")
CARD_BORDER = Border(left=thin_card, right=thin_card, top=thin_card, bottom=thin_card)


def db_merge_row(ws, row, c0, c1):
    if c1 > c0:
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)


def draw_card_border(ws, r0, r1, c0, c1):
    for r in range(r0, r1 + 1):
        for c in range(c0, c1 + 1):
            ws.cell(r, c).border = CARD_BORDER


def stat_card(ws, row0, c0, c1, label, value_formula, desc, color="1A1A1A"):
    # cells are left unmerged (single starting cell, natural text overflow) -
    # merging KPI cards made them fussy to resize/reflow, so keep it simple
    lc = ws.cell(row0, c0, label)
    lc.font = DB_LABEL; lc.alignment = Alignment(horizontal="left")
    vc = ws.cell(row0 + 1, c0, value_formula)
    vc.font = Font(bold=True, size=20, color=color); vc.alignment = Alignment(horizontal="left")
    dc = ws.cell(row0 + 2, c0, desc)
    dc.font = DB_LIGHT; dc.alignment = Alignment(horizontal="left")
    draw_card_border(ws, row0, row0 + 2, c0, c1)


TR = "'TOTAL RH'!"  # sheet-qualified reference prefix

# ---- title banner ----
db.merge_cells(start_row=1, start_column=2, end_row=2, end_column=16)
title_cell = db.cell(1, 2, "=" + TR + "A1")
title_cell.font = Font(bold=True, size=16, color="FFFFFF")
title_cell.alignment = Alignment(horizontal="left", vertical="center")
for cc in range(2, 17):
    for rr in (1, 2):
        db.cell(rr, cc).fill = banner
try:
    from openpyxl.drawing.image import Image as XLImage
    logo_candidates = [
        r"C:\Users\HAL-USER\Desktop\STOCK RH\logo ha2.png",
        r"C:\Users\HAL-USER\Desktop\logo ha2.png",
    ]
    logo_path = next((p for p in logo_candidates if __import__("os").path.exists(p)), None)
    if logo_path:
        img = XLImage(logo_path)
        img.height = 42
        img.width = 170
        db.add_image(img, "R1")
except Exception as _e:
    print("logo embed skipped:", _e)

# ---- KPI stat cards ----
KPI_ROW = 4
kpi_cols = [(2, 4), (6, 8), (10, 12), (14, 16)]
last_row_bkk = last_av_row["BKK27"]
last_row_lch = last_av_row["LCH27"]
stat_card(db, KPI_ROW, *kpi_cols[0], "SHORTFALL ALERT",
          "=MIN(%sD%d,%sJ%d)" % (TR, last_row_bkk, TR, last_row_lch),
          "=\"BKK27/LCH27 40'RH, worse \"&%sB%d" % (TR, last_row_bkk), color="C62828")
stat_card(db, KPI_ROW, *kpi_cols[1], "COMBINED 40'RH STOCK",
          "=%sO5" % TR, "BKK27 + LCH27")
stat_card(db, KPI_ROW, *kpi_cols[2], "COMBINED 40'RH BOOKING",
          "=%sO10" % TR, "pending + all weeks")
stat_card(db, KPI_ROW, *kpi_cols[3], "COMBINED 40'RH BALANCE",
          "=%sO15" % TR,
          "=\"pending + all weeks, both depots, \"&%sB%d" % (TR, last_row_bkk), color="C62828")

# ---- Detail by depot (two card-tables) ----
# Row structure (stock/pending/booking/AV-balance rows and how many of them)
# depends on RULE (Mon-Thu vs Fri), so mirror TOTAL RH's actual row range for
# each depot live - both depots share the same RULE-driven row count, only
# the values differ - instead of a hardcoded Rule-1-shaped row list.
DET_ROW0 = KPI_ROW + 4
db.cell(DET_ROW0, 2, "DETAIL BY DEPOT").font = Font(bold=True, size=11, color="0D3B12")
det_depot_cols = {"BKK27": 2, "LCH27": 8}
det_depot_ref_col = {"BKK27": ("C", "D"), "LCH27": ("I", "J")}
det_row_range = list(range(4, block_bottom["BKK27"] + 1))  # canonical row sequence, same for both depots
for loc in LOCS:
    c0 = det_depot_cols[loc]
    lbl_col = "B" if loc == "BKK27" else "H"
    hdr = db.cell(DET_ROW0 + 1, c0, "=%s%s2" % (TR, lbl_col))
    db_merge_row(db, DET_ROW0 + 1, c0, c0 + 2)
    hdr.font = Font(bold=True, size=12, color=("2A78D6" if loc == "BKK27" else "EB6834"))
    db.cell(DET_ROW0 + 2, c0 + 1, "20'RE").font = BOLD
    db.cell(DET_ROW0 + 2, c0 + 2, "40'RH").font = BOLD
    vcol_re, vcol_rh = det_depot_ref_col[loc]
    for i, src_r in enumerate(det_row_range):
        rr = DET_ROW0 + 3 + i
        is_stock = (src_r == 4)
        is_av = (not is_stock) and (src_r != pending_row[loc]) and (src_r not in booking_rows[loc])
        lc = db.cell(rr, c0, "=%s%s%d" % (TR, lbl_col, src_r))
        lc.font = BOLD if (is_stock or is_av) else Font()
        ve = db.cell(rr, c0 + 1, "=%s%s%d" % (TR, vcol_re, src_r))
        vh = db.cell(rr, c0 + 2, "=%s%s%d" % (TR, vcol_rh, src_r))
        for cell in (ve, vh):
            cell.alignment = ctr
            if is_stock or is_av:
                cell.font = BOLD
        if is_stock:
            for cc in range(c0, c0 + 3):
                db.cell(rr, cc).fill = grey
        elif is_av:
            for cc in range(c0, c0 + 3):
                db.cell(rr, cc).fill = yellow
    draw_card_border(db, DET_ROW0 + 1, DET_ROW0 + 2 + len(det_row_range), c0, c0 + 2)

# ---- Remarks (AV/DMG) ----
REM_DB_ROW0 = DET_ROW0 + 3 + len(det_row_range) + 2
db.cell(REM_DB_ROW0, 2, "REMARKS  (AV = move code IED/IEP/IER, DMG = move code OER)").font = Font(bold=True, size=11, color="0D3B12")
for loc in LOCS:
    c0 = det_depot_cols[loc]
    hdr = db.cell(REM_DB_ROW0 + 1, c0, "=%s%s2" % (TR, "B" if loc == "BKK27" else "H"))
    db_merge_row(db, REM_DB_ROW0 + 1, c0, c0 + 2)
    hdr.font = Font(bold=True, size=11, color=("2A78D6" if loc == "BKK27" else "EB6834"))
    db.cell(REM_DB_ROW0 + 2, c0 + 1, "20'RE").font = BOLD
    db.cell(REM_DB_ROW0 + 2, c0 + 2, "40'RH").font = BOLD
    vcol_re, vcol_rh = det_depot_ref_col[loc]
    for i, rlabel in enumerate(("AV", "DMG")):
        rr = REM_DB_ROW0 + 3 + i
        src_r = REM_HDR + 1 + i  # REM_HDR+1 = AV row, REM_HDR+2 = DMG row on TOTAL RH
        db.cell(rr, c0, rlabel)
        db.cell(rr, c0 + 1, "=%s%s%d" % (TR, vcol_re, src_r)).alignment = ctr
        db.cell(rr, c0 + 2, "=%s%s%d" % (TR, vcol_rh, src_r)).alignment = ctr
    draw_card_border(db, REM_DB_ROW0 + 1, REM_DB_ROW0 + 4, c0, c0 + 2)

# ---- RF Seasonal stat cards + current-mix summary ----
RFS_ROW0 = REM_DB_ROW0 + 7
db.cell(RFS_ROW0, 2, "RF SEASONAL \u2014 build year").font = Font(bold=True, size=11, color="0D3B12")
rfs_kpi_row = RFS_ROW0 + 1


def _sum2026_2020_45re(loc):
    # YEARS_FIXED starts at 2026 descending, so rows hr+2 .. hr+8 are years 2026..2020
    ycol = YBS_COL0[loc]
    r0, r1 = YBS_HR[loc] + 2, YBS_HR[loc] + 8
    cols_45re = [col_letter(ycol + 2 + 2 * i) for i in range(len(BRANDS))]  # Carrier/Daikin/Thermo 45RE cols
    terms = ["SUM(%s%s%d:%s%d)" % (TR, c, r0, c, r1) for c in cols_45re]
    return "=" + "+".join(terms)


stat_card(db, rfs_kpi_row, 2, 4, "BKK27 2020-26 BUILT", _sum2026_2020_45re("BKK27"),
          "of 40'RH on hand", color="2A78D6")
stat_card(db, rfs_kpi_row, 6, 8, "LCH27 2020-26 BUILT", _sum2026_2020_45re("LCH27"),
          "of 40'RH on hand", color="EB6834")
stat_card(db, rfs_kpi_row, 10, 12, "PEAK BUILD YEAR", "=\"%s\"" % PEAK_YEAR_LABEL,
          "%d units across both depots" % PEAK_YEAR_UNITS)

# current brand mix (GTTL row), one compact table per depot
MIX_ROW0 = rfs_kpi_row + 4
for loc in LOCS:
    c0 = det_depot_cols[loc]
    ycol = YBS_COL0[loc]
    grow = YBS_HR[loc] + 2 + len(YEARS_FIXED)
    hdr = db.cell(MIX_ROW0, c0, "%s current mix" % loc)
    db_merge_row(db, MIX_ROW0, c0, c0 + 2)
    hdr.font = Font(bold=True, size=10, color=("2A78D6" if loc == "BKK27" else "EB6834"))
    for i, b in enumerate(BRANDS):
        db.cell(MIX_ROW0 + 1, c0 + 1, "22RE").font = BOLD
        db.cell(MIX_ROW0 + 1, c0 + 2, "45RE").font = BOLD
        rr = MIX_ROW0 + 2 + i
        db.cell(rr, c0, b)
        cre = col_letter(ycol + 1 + 2 * i)
        crh = col_letter(ycol + 2 + 2 * i)
        db.cell(rr, c0 + 1, "=%s%s%d" % (TR, cre, grow)).alignment = ctr
        db.cell(rr, c0 + 2, "=%s%s%d" % (TR, crh, grow)).alignment = ctr
    draw_card_border(db, MIX_ROW0, MIX_ROW0 + 1 + len(BRANDS), c0, c0 + 2)
db.cell(MIX_ROW0 + 2 + len(BRANDS) + 1, 2,
        "Full 2010-2026 Year x Brand x Size breakdown is on the TOTAL RH sheet.").font = DB_LIGHT

db.column_dimensions["A"].width = 3
for loc in LOCS:
    c0 = det_depot_cols[loc]
    db.column_dimensions[col_letter(c0)].width = 22
    db.column_dimensions[col_letter(c0 + 1)].width = 9
    db.column_dimensions[col_letter(c0 + 2)].width = 9
# narrow spacer columns everywhere else on the sheet, instead of leaving
# them at Excel's default (much wider) auto-width
_used_cols = set()
for c0 in det_depot_cols.values():
    _used_cols.update([c0, c0 + 1, c0 + 2])
for c0, c1 in kpi_cols:
    _used_cols.update(range(c0, c1 + 1))
for cc in range(2, 17):
    if cc not in _used_cols:
        db.column_dimensions[col_letter(cc)].width = 2.5
db.freeze_panes = "B3"

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
_active_idx = wb.sheetnames.index("Dashboard")
wb.active = _active_idx
for i, ws in enumerate(wb.worksheets):
    ws.sheet_view.tabSelected = (i == _active_idx)

wb.save(OUT)
print("saved", OUT)
