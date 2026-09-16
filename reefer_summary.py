# -*- coding: utf-8 -*-
"""
สรุปยอดตู้ Reefer (RH) จาก folder "STOCK RH"
=================================================

ขั้นตอน:
  1) อ่านไฟล์ในโฟลเดอร์ STOCK RH ด้วย "ไลบรารีอ่านไฟล์" (ไม่ใช้ AI)
        - .pdf  -> pdfplumber  (ดึง text + table ของฟอร์ม RH FORM)
        - .xls / .xlsx / .csv -> pandas (+ xlrd สำหรับ .xls เก่า)
  2) รวม text ที่ดึงได้ทั้งหมดเก็บลงไฟล์  _extracted_text.txt   (เอาไว้ให้คน / AI อ่าน)
  3) สร้าง "สรุปยอดตู้ reefer" ตามโครงฟอร์ม RH FORM
        (แยกตาม Depot x ยี่ห้อเครื่องทำความเย็น x ขนาด x ปีสร้าง)
     เขียนออกเป็นไฟล์ Excel  REEFER_SUMMARY_<วันที่>.xlsx

รันเอง:   python "reefer_summary.py"
(ติดตั้งครั้งแรก:  pip install pdfplumber pandas xlrd openpyxl)
"""

from __future__ import annotations

import datetime as _dt
import re
import sys
from pathlib import Path

import pandas as pd

try:
    import pdfplumber
except ImportError:
    sys.exit("ยังไม่ได้ติดตั้ง pdfplumber  ->  pip install pdfplumber")

# --------------------------------------------------------------------------- #
# ตั้งค่า
# --------------------------------------------------------------------------- #
FOLDER = Path(__file__).resolve().parent          # โฟลเดอร์ STOCK RH (ที่สคริปต์วางอยู่)
TODAY = _dt.date.today().strftime("%Y-%m-%d")

TEXT_DUMP = FOLDER / "_extracted_text.txt"
OUT_XLSX = FOLDER / f"REEFER_SUMMARY_{TODAY}.xlsx"

# ยี่ห้อเครื่องทำความเย็นที่ฟอร์มรองรับ + ขนาดตู้
BRANDS = ["CARRIER", "DAIKIN", "THERMO KING"]
SIZES = ["22RE", "45RE"]

DEPOT_LABEL = {
    "BKK27": "BKK / BC2 (BKK27)",
    "LCH27": "LCH / HAST (LCH27)",
}

# --------------------------------------------------------------------------- #
# STEP 1 — อ่านไฟล์ด้วยไลบรารี (ไม่ใช้ AI)
# --------------------------------------------------------------------------- #
def extract_pdf_text(path: Path) -> str:
    """ดึง text + ตารางจาก PDF ด้วย pdfplumber"""
    chunks = [f"########## PDF: {path.name} ##########"]
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            chunks.append(f"\n----- page {pno} : text -----")
            chunks.append(page.extract_text() or "(ไม่มี text)")
            for tno, table in enumerate(page.extract_tables(), start=1):
                chunks.append(f"\n----- page {pno} : table {tno} -----")
                for row in table:
                    chunks.append(" | ".join("" if c is None else str(c).replace("\n", " ")
                                             for c in row))
    return "\n".join(chunks)


def read_table_file(path: Path) -> dict[str, pd.DataFrame]:
    """อ่าน excel/csv ทุก sheet เป็น DataFrame"""
    suf = path.suffix.lower()
    if suf in {".xls"}:
        xls = pd.ExcelFile(path, engine="xlrd")
    elif suf in {".xlsx", ".xlsm", ".xltx"}:
        xls = pd.ExcelFile(path, engine="openpyxl")
    elif suf in {".csv", ".tsv"}:
        sep = "\t" if suf == ".tsv" else ","
        return {path.stem: pd.read_csv(path, sep=sep)}
    else:
        return {}
    return {name: xls.parse(name) for name in xls.sheet_names}


def dataframe_to_text(name: str, df: pd.DataFrame) -> str:
    return (f"########## TABLE: {name}  (rows={len(df)}, cols={len(df.columns)}) ##########\n"
            + df.to_csv(index=False))


# --------------------------------------------------------------------------- #
# STEP 3 — สร้างสรุปยอด
# --------------------------------------------------------------------------- #
def norm_brand(v) -> str:
    s = str(v).upper().strip()
    if "CARR" in s:
        return "CARRIER"
    if "DAIKIN" in s:
        return "DAIKIN"
    if "THERMO" in s or "T/K" in s or s == "TK":
        return "THERMO KING"
    return s or "UNKNOWN"


def norm_size(v) -> str:
    s = str(v).upper().strip()
    if s.startswith("22"):
        return "22RE"
    if s.startswith("45") or s.startswith("40"):
        return "45RE"
    return s


def find_staying_df(tables: dict[str, pd.DataFrame]) -> pd.DataFrame:
    """หา sheet ที่เป็นรายการตู้ (มีคอลัมน์ Container No / Location / Size)"""
    for df in tables.values():
        cols = {str(c).strip().lower() for c in df.columns}
        if {"container no", "location"} <= cols or {"container no", "size/type"} <= cols:
            return df
    raise SystemExit("ไม่พบ sheet รายการตู้ (ต้องมีคอลัมน์ Container No / Location / Size/Type)")


def build_summary(df: pd.DataFrame, years: list[int]):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # ตัดแถวท้ายที่เป็นยอดรวม / แถวว่าง
    df = df[df["Location"].notna() & df["Container No"].notna()]
    df = df[df["Size/Type"].notna()]

    df["Depot"] = df["Location"].astype(str).str.strip().str.upper()
    df["Brand"] = df["RF Brand"].map(norm_brand)
    df["Size"] = df["Size/Type"].map(norm_size)
    df["Year"] = pd.to_numeric(df["Built Year"], errors="coerce").astype("Int64")

    # ทิศทางการเคลื่อนย้าย (I* = เข้า, O* = ออก)
    mc = df["Move Code"].astype(str).str.upper()
    df["Dir"] = mc.str[0].map({"I": "IN (รับเข้า)", "O": "OUT (จ่ายออก)"}).fillna("อื่นๆ")

    depots = [d for d in DEPOT_LABEL if d in set(df["Depot"])]
    depots += sorted(set(df["Depot"]) - set(DEPOT_LABEL))

    all_years = sorted(set(years) | set(int(y) for y in df["Year"].dropna().unique()), reverse=True)
    col_index = pd.MultiIndex.from_product([BRANDS, SIZES], names=["ยี่ห้อ", "ขนาด"])

    matrices: dict[str, pd.DataFrame] = {}
    for dp in depots:
        sub = df[df["Depot"] == dp]
        mat = pd.DataFrame(0, index=all_years, columns=col_index, dtype=int)
        g = sub.groupby(["Year", "Brand", "Size"]).size()
        for (yr, br, sz), n in g.items():
            if pd.isna(yr) or br not in BRANDS or sz not in SIZES:
                # ยี่ห้อ/ขนาดนอกฟอร์ม -> เก็บใน bucket "อื่นๆ"
                continue
            mat.at[int(yr), (br, sz)] += int(n)
        mat.index.name = "ปีสร้าง"
        mat["รวม"] = mat.sum(axis=1)
        mat.loc["รวมทั้งหมด"] = mat.sum(axis=0)
        matrices[dp] = mat

    # ตารางสรุปภาพรวม
    pivot_depot_size = pd.crosstab(df["Depot"], df["Size"], margins=True, margins_name="รวม")
    pivot_brand = pd.crosstab([df["Depot"]], [df["Brand"], df["Size"]], margins=True, margins_name="รวม")
    pivot_lease = pd.crosstab(df["Depot"], df["Lease Type"], margins=True, margins_name="รวม")
    pivot_dir = pd.crosstab(df["Depot"], df["Dir"], margins=True, margins_name="รวม")

    # ตู้ค้างนาน (aging)
    aging = df[["Depot", "Container No", "Size", "Brand", "Year", "Days", "Lessor", "Lease Type"]].copy()
    aging = aging.sort_values("Days", ascending=False)

    total = len(df)
    return {
        "depots": depots,
        "matrices": matrices,
        "pivot_depot_size": pivot_depot_size,
        "pivot_brand": pivot_brand,
        "pivot_lease": pivot_lease,
        "pivot_dir": pivot_dir,
        "aging": aging,
        "total": total,
        "clean_df": df,
    }


def years_from_pdf_text(text: str) -> list[int]:
    ys = {int(m) for m in re.findall(r"\b(20[0-2]\d)\b", text)}
    return sorted(ys, reverse=True) or list(range(2026, 2009, -1))


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main() -> None:
    pdf_files = sorted(FOLDER.glob("*.pdf"))
    xl_files = sorted(p for p in FOLDER.glob("*")
                      if p.suffix.lower() in {".xls", ".xlsx", ".xlsm", ".csv", ".tsv"}
                      and not p.name.startswith("REEFER_SUMMARY"))

    print(f"โฟลเดอร์      : {FOLDER}")
    print(f"PDF ที่พบ     : {[p.name for p in pdf_files]}")
    print(f"Excel/CSV ที่พบ: {[p.name for p in xl_files]}")

    # ---- STEP 1 + 2 : ดึง text ----
    text_blocks: list[str] = []
    pdf_text_all = ""
    for p in pdf_files:
        t = extract_pdf_text(p)
        pdf_text_all += "\n" + t
        text_blocks.append(t)

    all_tables: dict[str, pd.DataFrame] = {}
    for p in xl_files:
        tabs = read_table_file(p)
        for sh, d in tabs.items():
            key = f"{p.name} :: {sh}"
            all_tables[key] = d
            text_blocks.append(dataframe_to_text(key, d))

    TEXT_DUMP.write_text("\n\n".join(text_blocks), encoding="utf-8")
    print(f"\n[OK] เขียน text ที่ดึงได้ -> {TEXT_DUMP.name}")

    # ---- STEP 3 : สรุปยอด ----
    years = years_from_pdf_text(pdf_text_all)
    staying = find_staying_df(all_tables)
    R = build_summary(staying, years)

    # ---- เขียน Excel ----
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        info = pd.DataFrame({
            "หัวข้อ": ["วันที่สรุป", "แหล่งข้อมูล (Excel)", "แหล่งข้อมูล (ฟอร์ม PDF)",
                     "จำนวนตู้ reefer ค้างทั้งหมด"],
            "ค่า": [TODAY,
                   ", ".join(p.name for p in xl_files),
                   ", ".join(p.name for p in pdf_files),
                   R["total"]],
        })
        info.to_excel(xw, sheet_name="ภาพรวม", index=False, startrow=0)
        R["pivot_depot_size"].to_excel(xw, sheet_name="ภาพรวม", startrow=7)
        R["pivot_lease"].to_excel(xw, sheet_name="ภาพรวม", startrow=7 + len(R["pivot_depot_size"]) + 3)
        R["pivot_dir"].to_excel(xw, sheet_name="ภาพรวม",
                                startrow=7 + len(R["pivot_depot_size"]) + len(R["pivot_lease"]) + 6)

        R["pivot_brand"].to_excel(xw, sheet_name="แยกยี่ห้อ-ขนาด")

        for dp, mat in R["matrices"].items():
            sheet = re.sub(r"[\\/*?:\[\]]", "-", DEPOT_LABEL.get(dp, dp))[:31]
            mat.to_excel(xw, sheet_name=sheet)

        R["aging"].to_excel(xw, sheet_name="ตู้ค้างเรียงตามวัน", index=False)

    print(f"[OK] เขียนสรุป -> {OUT_XLSX.name}")

    # ---- พิมพ์สรุปย่อทางจอ ----
    print("\n" + "=" * 60)
    print(f"สรุปยอดตู้ REEFER  ณ วันที่ {TODAY}")
    print("=" * 60)
    print(f"ตู้ reefer ค้างทั้งหมด : {R['total']} TEU-unit\n")
    print(R["pivot_depot_size"].to_string())
    print("\nแยกตามประเภทสัญญาเช่า (ONE=ตู้สายเรือ, L/T=เช่าระยะยาว, OWN=ตู้บริษัท):")
    print(R["pivot_lease"].to_string())
    print("\nทิศทางการเคลื่อนย้าย:")
    print(R["pivot_dir"].to_string())
    for dp, mat in R["matrices"].items():
        print(f"\n----- {DEPOT_LABEL.get(dp, dp)} -----")
        print(mat.to_string())
    top = R["aging"].head(5)
    print("\nตู้ค้างนานสุด 5 อันดับ (Days):")
    print(top.to_string(index=False))


if __name__ == "__main__":
    main()
