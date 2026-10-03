# -*- coding: utf-8 -*-
"""
ระบบคำนวณค่า Storage (BKK / UTCT / LCH)
=======================================
ขั้นตอน
  1. วางไฟล์จากระบบของเที่ยวเรือ (LOAD / BKG / GATE .xls) หรือไฟล์ TEMPLATE ที่กรอกแล้ว ลงในโฟลเดอร์ INPUT
  2. รัน  python storage_calc.py   (หรือดับเบิลคลิก run_storage.bat)
  3. ได้ผลลัพธ์แยกตามเที่ยวเรือ ตั้งชื่อ  ETD_SVC_VSL_VOY_POL_LWharf
       VOYAGES/<ชื่อ>/<ชื่อ>.xlsx            รายงาน (Tahoma 8 มีเส้นตาราง พร้อมพิมพ์/ส่งต่อ)
       VOYAGES/<ชื่อ>/<ชื่อ>_DASHBOARD.html  dashboard ของเที่ยวนั้น
       VOYAGES/<ชื่อ>/SOURCE/                ไฟล์ต้นฉบับที่ใช้คำนวณ (ย้ายออกจาก INPUT)
       DATABASE/STORAGE_DATABASE.xlsx        ฐานข้อมูลรวมทุกเที่ยว (สร้างใหม่ทุกครั้งที่รัน)
       DASHBOARD.html                        dashboard รวมทุกเที่ยว
  รันเที่ยวเดิมซ้ำได้ ผลจะถูกเขียนทับในโฟลเดอร์เดิม
"""
import glob
import html
import json
import os
import shutil
import sys
from collections import defaultdict
from datetime import date, datetime

import openpyxl
import xlrd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE, "INPUT")
VOY_DIR = os.path.join(BASE, "VOYAGES")
DB_DIR = os.path.join(BASE, "DATABASE")
DB_FILE = os.path.join(DB_DIR, "STORAGE_DATABASE.xlsx")
MASTER_DASH = os.path.join(BASE, "DASHBOARD.html")

# ---------------------------------------------------------------- Tariff (ชีท Charge)
# ขั้นบันได: (วันเริ่ม tier, ยอดสะสม ณ วันเริ่ม, อัตรา/วัน) -> ค่า = สะสม + (วัน - วันเริ่ม) * อัตรา
TIERS = {
    ("G1", "GP", "20'"): [(0, 0, 125), (3, 375, 250), (10, 2275, 400)],
    ("G1", "GP", "40'"): [(0, 0, 250), (3, 750, 500), (10, 4550, 800)],
    ("G1", "DG", "20'"): [(0, 0, 375), (3, 1125, 750), (10, 6825, 1200)],
    ("G1", "DG", "40'"): [(0, 0, 750), (3, 2250, 1500), (10, 13650, 2400)],
    ("G2", "GP", "20'"): [(0, 0, 125), (7, 875, 250), (14, 2625, 400)],
    ("G2", "GP", "40'"): [(0, 0, 250), (7, 1750, 500), (14, 5250, 800)],
    ("G2", "DG", "20'"): [(0, 0, 375), (7, 2625, 750), (14, 7875, 1200)],
    ("G2", "DG", "40'"): [(0, 0, 750), (7, 5250, 1500), (14, 15750, 2400)],
    ("G3", "GP", "20'"): [(0, 0, 160), (7, 1120, 275), (14, 3045, 390)],
    ("G3", "GP", "40'"): [(0, 0, 320), (7, 2240, 550), (14, 6090, 780)],
    ("G3", "DG", "20'"): [(0, 0, 480), (7, 3360, 825), (14, 9135, 1170)],
    ("G3", "DG", "40'"): [(0, 0, 960), (7, 6720, 1650), (14, 18270, 2340)],
    ("G4", "GP", "20'"): [(0, 0, 50)],
    ("G4", "GP", "40'"): [(0, 0, 100)],
    ("G4", "DG", "20'"): [(0, 0, 150)],
    ("G4", "DG", "40'"): [(0, 0, 300)],
    # ลูกค้า (คืนตู้ก่อนวันที่ให้คืน) — ค่าตามชีท Charge!K31:W34
    ("CUST", "GP", "20'"): [(0, 0, 375), (5, 1200, 750), (9, 4800, 1500)],
    ("CUST", "GP", "40'"): [(0, 0, 750), (7, 2400, 1500), (14, 9600, 3000)],
    ("CUST", "DG", "20'"): [(0, 0, 900), (5, 3360, 1800), (9, 9135, 3600)],
    ("CUST", "DG", "40'"): [(0, 0, 1800), (9, 6720, 3600), (9, 18270, 7200)],
}
GROUP_OF_LOC = {"LCH04": "G1", "BKK01": "G3", "BKK04": "G3", "BKK02": "G4"}  # อื่น ๆ = G2
TERMINAL_FT = {"LCH01": 7, "LCH02": 7, "LCH03": 7, "LCH04": 7, "LCH05": 7, "LCH06": 7, "LCH08": 7, "LCH09": 7,
               "LCH10": 7, "BKK02": 5}  # อื่น ๆ (BKK01/04) 3, DG 1
TERMINAL_NAME = {"BKK01": "PAT / Terminal 2", "BKK04": "PAT / Terminal 1", "BKK02": "Unithai (UTCT)",
                 "LCH01": "ESCO B3", "LCH02": "Hutchison A2", "LCH03": "TIPS B4", "LCH04": "LCMT A0/B1",
                 "LCH05": "LCIT B5", "LCH06": "Hutchison A3", "LCH08": "Hutchison C1C2", "LCH09": "Hutchison D1",
                 "LCH10": "LCIT C3"}

CS_COLS = ["Container No.", "Group", "Location", "Terminal", "Booking No.", "Shipper", "Org Shipper", "RCVCD", "POL",
           "POD", "TS PORT", "Type Size", "CGO TERM", "DG", "RF", "Vessel", "Voyage", "SVC", "Size",
           "Original Schedule", "ETD", "Vssl Delay", "1st Date allow to return", "Date In", "Total Stay",
           "Normal F/T Terminal", "Total Storage Days", "Total Storage Charge", "Special F/T",
           "HQ Ref No. / TML / APP BY", "Free Time Customer", "คืนก่อนเรือออก", "Over F/T Customer",
           "Amount due to Customer", "Collection", "Amount due to HAL", "Remark", "SALESNM", "CUST ETD",
           "PICK UP", "COMMON REMARKS", "ROLL"]
DATE_COLS = {"Original Schedule", "ETD", "1st Date allow to return", "Date In", "CUST ETD"}
MONEY_COLS = {"Total Storage Charge", "Amount due to Customer", "Collection", "Amount due to HAL"}
SUM_COLS = ["Total Storage Days"] + sorted(MONEY_COLS)


CFS_ZERO_FT_LOCS = {"BKK01", "BKK04"}  # PAT: ตู้ CGO TERM CFS/CY (stuffing ใน terminal) ไม่มี free time


def ft_terminal(loc, is_dg, cgo_term=None):
    if loc in CFS_ZERO_FT_LOCS and str(cgo_term or "").replace(" ", "").upper() == "CFS/CY":
        return 0
    return 1 if is_dg else TERMINAL_FT.get(loc, 3)


def area_of(loc):
    loc = loc or ""
    if loc == "BKK02":
        return "UTCT"
    if loc.startswith("BKK"):
        return "BKK"
    if loc.startswith("LCH"):
        return "LCH"
    return "OTHER"


def tier_charge(key, days):
    if not days or days <= 0:
        return 0
    row = None
    for t in TIERS[key]:
        if t[0] <= days:
            row = t
    return row[1] + (days - row[0]) * row[2]


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        return date.fromordinal(date(1899, 12, 30).toordinal() + int(v))
    s = str(v).strip()[:10]
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def num(v):
    try:
        return float(v) if v not in (None, "") else 0
    except (TypeError, ValueError):
        return 0


def pad(row, n=110):
    row = list(row)
    return row + [None] * (n - len(row))


# ---------------------------------------------------------------- read input files
def read_book(path):
    """คืน {sheet: rows} และ {sheet: ws(formula)} สำหรับไฟล์ xlsx"""
    if path.lower().endswith(".xls"):
        bk = xlrd.open_workbook(path)
        out = {}
        for sh in bk.sheets():
            rows = []
            for i in range(sh.nrows):
                row = []
                for c in sh.row(i):
                    if c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) or c.value == "":
                        v = None
                    elif c.ctype == xlrd.XL_CELL_DATE:
                        v = xlrd.xldate.xldate_as_datetime(c.value, bk.datemode)
                    elif c.ctype == xlrd.XL_CELL_NUMBER:
                        v = int(c.value) if float(c.value).is_integer() else c.value
                    else:
                        v = c.value
                    row.append(v)
                rows.append(row)
            out[sh.name] = rows
        return out, {}
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb}
    formulas = {}
    if "CS" in wb.sheetnames:
        wf = openpyxl.load_workbook(path, read_only=True)
        formulas["CS"] = [list(r) for r in wf["CS"].iter_rows(max_col=25, values_only=True)]
        wf.close()
    return out, formulas


def clean(rows):
    return [pad(r) for r in rows if r and r[0] is not None and str(r[0]).strip() != ""]


def classify(path):
    name = os.path.basename(path)
    src = {"path": path, "name": name, "kind": None, "loading": [], "booking": [], "gate": [], "manual": {},
           "heads": {}}
    try:
        sheets, formulas = read_book(path)
    except Exception as e:  # noqa: BLE001
        src["error"] = f"อ่านไฟล์ไม่ได้ ({e})"
        return src
    if all(s in sheets for s in ("1.Loading", "2.Bookinglist", "3.GateMove")):
        src["kind"] = "TEMPLATE"
        src["loading"] = clean(sheets["1.Loading"][1:])
        src["booking"] = clean(sheets["2.Bookinglist"][1:])
        src["gate"] = clean(sheets["3.GateMove"][1:])
        src["heads"] = {"LOAD": sheets["1.Loading"][0], "BKG": [h for h in sheets["2.Bookinglist"][0] if h],
                        "GATE": sheets["3.GateMove"][0]}
        cs = sheets.get("CS", [])
        yf = formulas.get("CS", [])
        for i in range(2, len(cs)):
            r = pad(cs[i])
            if not r[0]:
                continue
            y = yf[i][24] if i < len(yf) and len(yf[i]) > 24 else None
            src["manual"][r[0]] = {"orig": r[18], "allow": r[21], "spft": num(r[27]), "tml": r[28],
                                   "collect": num(r[33]), "remark": r[35],
                                   "ft_override": y if isinstance(y, (int, float)) else None}
    else:
        for rows in sheets.values():
            for h in range(min(len(rows), 6)):
                hdr = [str(v).strip() if v is not None else "" for v in rows[h]]
                kind = None
                if hdr and hdr[0] == "B/K No":
                    kind = "BKG"
                elif hdr and hdr[0] == "Container No." and "EQ Date" in hdr:
                    kind = "LOAD"
                elif hdr and hdr[0] == "Container No" and "GateDate" in hdr:
                    kind = "GATE"
                if kind:
                    src["kind"] = kind
                    data = clean(rows[h + 1:])
                    src[{"LOAD": "loading", "BKG": "booking", "GATE": "gate"}[kind]] = data
                    src["heads"][kind] = rows[h]
                    break
            if src["kind"]:
                break
    if not src["kind"]:
        src["error"] = "ไม่รู้จักรูปแบบไฟล์ (ต้องเป็น LOAD / BKG / GATE หรือ TEMPLATE)"
    return src


# ---------------------------------------------------------------- calculation
def compute(sources):
    bk, gates = {}, defaultdict(list)
    for si, s in enumerate(sources):
        for r in s["booking"]:
            bk.setdefault(r[0], (r, si))
        for r in s["gate"]:
            gates[r[0]].append((r, si))
    heads = {}
    for s in sources:
        for k, v in s["heads"].items():
            heads.setdefault(k, v)

    rows, warnings = [], []
    for si, s in enumerate(sources):
        for r in s["loading"]:
            if r[10] != "F":
                continue
            cntr, loc, bkno, tpsz = r[0], r[13], r[16], str(r[4] or "")
            b, bsi = bk.get(bkno, (pad([]), None))
            if bsi is None:
                warnings.append(f"{cntr}: ไม่พบ Booking {bkno}")
            gl = gates.get(cntr, [])
            g, gsi = next(((x, i) for x, i in gl if x[12] == bkno), gl[0] if gl else (None, None))
            if g is None:
                warnings.append(f"{cntr}: ไม่พบวันตู้เข้าในไฟล์ Gate move")
            m = s["manual"].get(cntr, {})

            group = GROUP_OF_LOC.get(loc, "G2")
            is_dg = b[42] not in (None, "", 0, "0")
            is_rf = tpsz.endswith("RE")
            size = "20'" if tpsz.startswith("22") else "40'"
            svc = b[11] or r[19]
            etd = to_date(r[1])
            din = to_date(g[3]) if g else None
            stay = (etd - din).days + 1 if etd and din else None
            ft = ft_terminal(loc, is_dg, b[84])
            note = "CFS/CY ที่ PAT: F/T 0 วัน" if (ft == 0 and loc in CFS_ZERO_FT_LOCS) else None
            if m.get("ft_override") is not None and m["ft_override"] != ft:  # เตือนเฉพาะค่าที่ต่างจากกฎ
                note = f"F/T {m['ft_override']} วัน ตามที่กรอกในชีท CS (กฎให้ {ft} วัน)"
                warnings.append(f"{cntr}: ใช้ F/T Terminal = {m['ft_override']} วัน ตามที่กรอกเองในชีท CS (กฎให้ {ft} วัน)")
                ft = m["ft_override"]
            sp_ft = m.get("spft") or 0
            tml = m.get("tml") or ""
            ft_used = sp_ft if (str(tml).strip().upper() == "TML" and sp_ft > 0) else ft
            days = max(stay - ft_used, 0) if stay is not None else 0
            charge = tier_charge((group, "DG" if is_dg else "GP", size), days)
            cft = (3 if (is_dg or is_rf) else 5) + (1 if ((loc or "").startswith("BKK") and is_rf and svc == "KTS")
                                                    else 0) + (1 if svc == "KHS" else 0)
            allow, orig = to_date(m.get("allow")), to_date(m.get("orig"))
            early = max((allow - din).days, 0) if (allow and din) else 0
            over_c = max(early - sp_ft, 0)
            cust = tier_charge(("CUST", "DG" if is_dg else "GP", size), over_c)
            collect = m.get("collect") or 0

            cs = {
                "Container No.": cntr, "Group": group, "Location": loc, "Terminal": TERMINAL_NAME.get(loc, ""),
                "Booking No.": bkno, "Shipper": b[26] or r[22], "Org Shipper": b[27], "RCVCD": b[4], "POL": b[14],
                "POD": b[18] or r[8], "TS PORT": b[25], "Type Size": tpsz, "CGO TERM": b[84],
                "DG": "DG" if is_dg else None, "RF": "RF" if is_rf else None, "Vessel": r[2], "Voyage": r[3],
                "SVC": svc, "Size": size, "Original Schedule": orig, "ETD": etd,
                "Vssl Delay": (etd - orig).days if (etd and orig) else None, "1st Date allow to return": allow,
                "Date In": din, "Total Stay": stay, "Normal F/T Terminal": ft, "Total Storage Days": days,
                "Total Storage Charge": charge, "Special F/T": sp_ft or None, "HQ Ref No. / TML / APP BY": tml or None,
                "Free Time Customer": cft, "คืนก่อนเรือออก": (etd - allow).days + 1 if (etd and allow) else None,
                "Over F/T Customer": over_c, "Amount due to Customer": cust, "Collection": collect or None,
                "Amount due to HAL": charge - collect, "Remark": m.get("remark") or note, "SALESNM": b[39],
                "CUST ETD": to_date(b[100]), "PICK UP": b[62], "COMMON REMARKS": b[51], "ROLL": b[102],
            }
            key = (etd, svc, b[12] or r[2], b[13] or r[3], b[14] or r[7], b[15] or r[27] or loc)
            rows.append({"cs": cs, "key": key, "area": area_of(loc), "load": r, "bkg": b if bsi is not None else None,
                         "gate": g, "srcs": {si, bsi, gsi} - {None}})
    return rows, warnings, heads


def voyage_name(key):
    etd, *rest = key
    parts = [etd.isoformat() if etd else "NOETD"] + [str(x or "-").strip() for x in rest]
    return "_".join("".join("-" if ch in '\\/:*?"<>| ' else ch for ch in p) for p in parts)


def summarize(rows):
    cs = [r["cs"] if "cs" in r else r for r in rows]
    return {
        "n": len(cs), "c20": sum(c["Size"] == "20'" for c in cs), "c40": sum(c["Size"] == "40'" for c in cs),
        "dg": sum(bool(c["DG"]) for c in cs), "rf": sum(bool(c["RF"]) for c in cs),
        "over": sum((c["Total Storage Days"] or 0) > 0 for c in cs),
        "days": sum(c["Total Storage Days"] or 0 for c in cs),
        "charge": sum(c["Total Storage Charge"] or 0 for c in cs),
        "cust": sum(c["Amount due to Customer"] or 0 for c in cs),
        "collect": sum(c["Collection"] or 0 for c in cs),
        "hal": sum(c["Amount due to HAL"] or 0 for c in cs),
    }


# ---------------------------------------------------------------- Excel styling (Tahoma 8 + grid)
F = Font(name="Tahoma", size=8)
FB = Font(name="Tahoma", size=8, bold=True)
FH = Font(name="Tahoma", size=8, bold=True, color="FFFFFF")
FT_TITLE = Font(name="Tahoma", size=8, bold=True, color="FFFFFF")
FILL_H = PatternFill("solid", fgColor="1F3A4D")
FILL_SUB = PatternFill("solid", fgColor="E3EBF1")
FILL_TOT = PatternFill("solid", fgColor="FFF1D6")
FILL_PAID = PatternFill("solid", fgColor="FBE2C8")
FILL_LABEL = PatternFill("solid", fgColor="F3F5F7")
SIDE = Side(style="thin", color="7F8A93")
GRID = Border(left=SIDE, right=SIDE, top=SIDE, bottom=SIDE)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
VCENTER = Alignment(vertical="center")
FMT_MONEY = "#,##0;-#,##0;\"-\""
FMT_DATE = "dd/mm/yyyy"


def cell(ws, r, c, v=None, font=F, fill=None, fmt=None, align=VCENTER, border=GRID):
    x = ws.cell(r, c)
    if v is not None:
        x.value = v
    x.font = font
    if fill:
        x.fill = fill
    if fmt:
        x.number_format = fmt
    if align:
        x.alignment = align
    if border:
        x.border = border
    return x


def print_setup(ws, title_rows=None, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "&8หน้า &P / &N"
    ws.oddFooter.right.text = "&8&A"
    if title_rows:
        ws.print_title_rows = title_rows


def write_table(ws, top, cols, rows, widths=None, total=None, paid_col=None):
    """ตารางมาตรฐาน: หัวตารางสีเข้ม, Tahoma 8, เส้นตารางทุกช่อง"""
    for j, h in enumerate(cols, 1):
        cell(ws, top, j, h, FH, FILL_H, align=CENTER)
    ws.row_dimensions[top].height = 24
    for i, row in enumerate(rows, top + 1):
        paid = paid_col is not None and (row.get(paid_col) or 0) > 0
        for j, h in enumerate(cols, 1):
            v = row.get(h)
            fmt = FMT_DATE if h in DATE_COLS else FMT_MONEY if h in MONEY_COLS else None
            fill = FILL_PAID if (paid and h in ("Total Storage Charge", "Amount due to HAL", "Total Storage Days")) else None
            cell(ws, i, j, v, F, fill, fmt)
    last = top + len(rows)
    if total:
        last += 1
        for j, h in enumerate(cols, 1):
            v = total.get(h)
            cell(ws, last, j, v, FB, FILL_TOT, FMT_MONEY if (h in MONEY_COLS or h in SUM_COLS) else None)
    for j, h in enumerate(cols, 1):
        if widths and h in widths:
            w = widths[h]
        else:
            longest = max([len(str(h))] + [len(str(r.get(h))) for r in rows[:300] if r.get(h) is not None])
            w = 11 if h in DATE_COLS else min(max(longest * 0.95 + 2, 6), 38)
        ws.column_dimensions[get_column_letter(j)].width = w
    return last


def sheet_raw(wb, title, head, data):
    ws = wb.create_sheet(title)
    head = [h for h in head]
    while head and head[-1] in (None, ""):
        head.pop()
    cols = [str(h) if h is not None else f"Col{i+1}" for i, h in enumerate(head)]
    seen, uniq = {}, []
    for c in cols:  # หัวคอลัมน์ซ้ำ (เช่น Shipper, Pickup) ให้แยกกันได้
        seen[c] = seen.get(c, 0) + 1
        uniq.append(c if seen[c] == 1 else f"{c} ({seen[c]})")
    rows = [{uniq[j]: (r[j] if j < len(r) else None) for j in range(len(uniq))} for r in data]
    write_table(ws, 1, uniq, rows)
    for j, c in enumerate(cols, 1):
        ws.cell(1, j).value = c
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(uniq))}{len(rows) + 1}"
    print_setup(ws, "1:1")
    return ws


def write_voyage_xlsx(path, name, meta, rows, heads):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    s = summarize(rows)
    ws.column_dimensions["A"].width = 3
    for col, w in zip("BCDEFGHIJ", (20, 22, 3, 22, 16, 3, 13, 13, 13)):
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:K2")
    cell(ws, 2, 2, "STORAGE CHARGE REPORT  ·  รายงานค่าฝากตู้ที่ Terminal", FT_TITLE, FILL_H, align=VCENTER)
    for c in range(3, 12):
        cell(ws, 2, c, None, FT_TITLE, FILL_H)
    ws.row_dimensions[2].height = 20
    ws.merge_cells("B3:K3")
    cell(ws, 3, 2, name, FB, FILL_SUB)
    for c in range(3, 12):
        cell(ws, 3, c, None, FB, FILL_SUB)

    info = [("ETD", to_date(meta["etd"])), ("SVC", meta["svc"]), ("VSL", meta["vsl"]), ("VOY", meta["voy"]),
            ("POL", meta["pol"]), ("LWharf", meta["lwharf"]), ("Terminal", meta["terminal"]),
            ("Area / Group", f"{meta['area']} / {meta['group']}")]
    kpi = [("จำนวนตู้ทั้งหมด", s["n"], None), ("20' / 40'", f"{s['c20']} / {s['c40']}", None),
           ("DG / RF", f"{s['dg']} / {s['rf']}", None), ("ตู้ที่เกิน Free time", s["over"], None),
           ("รวมวันคิดเงิน (วัน)", s["days"], None), ("Total Storage Charge", s["charge"], FMT_MONEY),
           ("Amount due to Customer", s["cust"], FMT_MONEY), ("Amount due to HAL", s["hal"], FMT_MONEY)]
    cell(ws, 5, 2, "ข้อมูลเที่ยวเรือ", FH, FILL_H)
    cell(ws, 5, 3, None, FH, FILL_H)
    cell(ws, 5, 5, "สรุปผล", FH, FILL_H)
    cell(ws, 5, 6, None, FH, FILL_H)
    for i, (k, v) in enumerate(info, 6):
        cell(ws, i, 2, k, FB, FILL_LABEL)
        cell(ws, i, 3, v, F, fmt=FMT_DATE if k == "ETD" else None, align=Alignment(horizontal="left"))
    for i, (k, v, fmt) in enumerate(kpi, 6):
        hl = k == "Amount due to HAL"
        cell(ws, i, 5, k, FB, FILL_TOT if hl else FILL_LABEL)
        cell(ws, i, 6, v, FB if hl else F, FILL_TOT if hl else None, fmt, Alignment(horizontal="right"))

    # breakdown by size x cargo
    top = 16
    cell(ws, top - 1, 2, "แยกตามขนาดตู้และประเภทสินค้า", FB, border=None)
    cols = ["ขนาด", "ประเภท", "ตู้", "เกิน F/T", "วันคิดเงิน", "Storage", "Customer", "HAL"]
    colpos = [2, 3, 5, 6, 8, 9, 10, 11]
    ws.column_dimensions["K"].width = 13
    for h, c in zip(cols, colpos):
        cell(ws, top, c, h, FH, FILL_H, align=CENTER)
    r = top
    for size in ("20'", "40'"):
        for cg in ("GP", "RF", "DG"):
            sub = [x for x in rows if x["cs"]["Size"] == size and
                   (("DG" if x["cs"]["DG"] else "RF" if x["cs"]["RF"] else "GP") == cg)]
            if not sub:
                continue
            t = summarize(sub)
            r += 1
            for v, c, fmt in zip([size, cg, t["n"], t["over"], t["days"], t["charge"], t["cust"], t["hal"]],
                                 colpos, [None] * 5 + [FMT_MONEY] * 3):
                cell(ws, r, c, v, F, fmt=fmt)
    r += 1
    for v, c, fmt in zip(["รวม", "", s["n"], s["over"], s["days"], s["charge"], s["cust"], s["hal"]],
                         colpos, [None] * 5 + [FMT_MONEY] * 3):
        cell(ws, r, c, v, FB, FILL_TOT, fmt)
    for rr in range(top, r + 1):  # fill gap columns D, G so grid looks continuous
        style = (FH, FILL_H) if rr == top else (FB, FILL_TOT) if rr == r else (F, None)
        for c in (4, 7):
            cell(ws, rr, c, None, *style)

    # notes
    r += 2
    cell(ws, r, 2, "หมายเหตุ", FB, border=None)
    notes = [f"Free time terminal: {', '.join(sorted({str(x['cs']['Normal F/T Terminal']) for x in rows}))} วัน · "
             f"Storage Days = (ETD − Date In + 1) − Free time",
             f"ไฟล์ต้นฉบับ: {', '.join(meta['sources'])}",
             f"คำนวณเมื่อ: {meta['generated']}"] + [f"ตรวจสอบ: {w}" for w in meta["warnings"]]
    for n in notes:
        r += 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=11)
        cell(ws, r, 2, n, F, border=None, align=Alignment(wrap_text=True, vertical="top"))

    # sign-off block for handing to the next department
    r += 2
    for c, label in zip((2, 5, 8), ("จัดทำโดย (Prepared by)", "ตรวจสอบโดย (Checked by)", "อนุมัติโดย (Approved by)")):
        end = c + 1 if c != 8 else c + 3
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=end)
        ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 3, end_column=end)
        ws.merge_cells(start_row=r + 4, start_column=c, end_row=r + 4, end_column=end)
        for rr in range(r, r + 5):
            for cc in range(c, end + 1):
                cell(ws, rr, cc, None, F, FILL_LABEL if rr == r else None)
        cell(ws, r, c, label, FB, FILL_LABEL, align=CENTER)
        cell(ws, r + 4, c, "วันที่ ......../......../........", F, align=CENTER)
    ws.sheet_view.showGridLines = False
    print_setup(ws, landscape=False)

    # CS detail
    wcs = wb.create_sheet("CS")
    data = sorted((x["cs"] for x in rows), key=lambda c: (-(c["Total Storage Charge"] or 0), c["Container No."]))
    total = {"Container No.": "TOTAL", "Group": f"{len(data)} ตู้"}
    for h in SUM_COLS:
        total[h] = sum(c[h] or 0 for c in data)
    widths = {"Shipper": 30, "Org Shipper": 30, "COMMON REMARKS": 40, "ROLL": 40, "Remark": 24,
              "HQ Ref No. / TML / APP BY": 14, "Terminal": 15}
    write_table(wcs, 1, CS_COLS, data, widths, total, paid_col="Total Storage Charge")
    wcs.freeze_panes = "B2"
    wcs.auto_filter.ref = f"A1:{get_column_letter(len(CS_COLS))}{len(data) + 1}"
    print_setup(wcs, "1:1")

    # raw data used for this voyage
    def uniq(seq):
        out, seen = [], set()
        for x in seq:
            if x is not None and id(x) not in seen:
                seen.add(id(x))
                out.append(x)
        return out
    sheet_raw(wb, "1.Loading", heads.get("LOAD", []), uniq(x["load"] for x in rows))
    if heads.get("BKG"):
        sheet_raw(wb, "2.Bookinglist", heads["BKG"], uniq(x["bkg"] for x in rows))
    if heads.get("GATE"):
        sheet_raw(wb, "3.GateMove", heads["GATE"], uniq(x["gate"] for x in rows))
    wb.save(path)


def write_database(voyages):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Voyages"
    vcols = ["Voyage Key", "ETD", "SVC", "VSL", "VOY", "POL", "LWharf", "Area", "Group", "Terminal", "ตู้", "20'", "40'",
             "DG", "RF", "เกิน F/T", "วันคิดเงิน", "Total Storage Charge", "Amount due to Customer", "Collection",
             "Amount due to HAL", "Processed"]
    vrows, crow = [], []
    for v in voyages:
        m, s = v["meta"], summarize(v["rows"])
        vrows.append({"Voyage Key": m["name"], "ETD": to_date(m["etd"]), "SVC": m["svc"], "VSL": m["vsl"], "VOY": m["voy"],
                      "POL": m["pol"], "LWharf": m["lwharf"], "Area": m["area"], "Group": m["group"],
                      "Terminal": m["terminal"], "ตู้": s["n"], "20'": s["c20"], "40'": s["c40"], "DG": s["dg"],
                      "RF": s["rf"], "เกิน F/T": s["over"], "วันคิดเงิน": s["days"], "Total Storage Charge": s["charge"],
                      "Amount due to Customer": s["cust"], "Collection": s["collect"], "Amount due to HAL": s["hal"],
                      "Processed": m["generated"]})
        for c in v["rows"]:
            crow.append({**c, "Voyage Key": m["name"], "Area": m["area"]})
    tot = {"Voyage Key": f"TOTAL {len(vrows)} เที่ยว"}
    for h in ("ตู้", "20'", "40'", "DG", "RF", "เกิน F/T", "วันคิดเงิน"):
        tot[h] = sum(r[h] for r in vrows)
    for h in MONEY_COLS:
        tot[h] = sum(r[h] for r in vrows)
    write_table(ws, 1, vcols, vrows, {"Voyage Key": 40, "Processed": 16, "Terminal": 15}, tot)
    for i in range(2, len(vrows) + 3):
        for j, h in enumerate(vcols, 1):
            if h in ("ตู้", "20'", "40'", "DG", "RF", "เกิน F/T", "วันคิดเงิน"):
                ws.cell(i, j).number_format = "#,##0"
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(vcols))}{len(vrows) + 1}"
    print_setup(ws, "1:1")

    wc = wb.create_sheet("Containers")
    ccols = ["Voyage Key", "Area"] + CS_COLS
    write_table(wc, 1, ccols, crow, {"Voyage Key": 36, "Shipper": 30, "Org Shipper": 30, "COMMON REMARKS": 40,
                                     "ROLL": 40, "Remark": 24, "Terminal": 15}, paid_col="Total Storage Charge")
    wc.freeze_panes = "C2"
    wc.auto_filter.ref = f"A1:{get_column_letter(len(ccols))}{len(crow) + 1}"
    print_setup(wc, "1:1")
    os.makedirs(DB_DIR, exist_ok=True)
    wb.save(DB_FILE)


# ---------------------------------------------------------------- dashboards (static HTML)
DASH_CSS = """
:root{--bg:#EEF1F3;--surface:#FFFFFF;--sunken:#F5F7F8;--ink:#15222B;--muted:#5A6873;--line:#D5DCE1;
--head:#1F3A4D;--head-ink:#EAF1F5;--head-muted:#A3B6C3;--accent:#D8812A;--accent-ink:#8F4E0E;--free:#B9C8D3;
--good:#2E7A4E;--warn:#B4231A;--warn-soft:#FBE4E1;--rf:#1F6FA8;color-scheme:light}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--bg:#0F161B;--surface:#172229;--sunken:#121B21;
--ink:#E3EAEE;--muted:#93A3AE;--line:#2A3943;--head:#0B1318;--head-ink:#E3EAEE;--head-muted:#7F95A3;--accent:#E99B4C;
--accent-ink:#F2B57A;--free:#3B4F5D;--good:#5CC48A;--warn:#F07B70;--warn-soft:#3A1C19;--rf:#6DB4E8;color-scheme:dark}}
:root[data-theme="dark"]{--bg:#0F161B;--surface:#172229;--sunken:#121B21;--ink:#E3EAEE;--muted:#93A3AE;--line:#2A3943;
--head:#0B1318;--head-ink:#E3EAEE;--head-muted:#7F95A3;--accent:#E99B4C;--accent-ink:#F2B57A;--free:#3B4F5D;
--good:#5CC48A;--warn:#F07B70;--warn-soft:#3A1C19;--rf:#6DB4E8;color-scheme:dark}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:13px/1.5 Tahoma,"Segoe UI",sans-serif}
.num{font-variant-numeric:tabular-nums}
header{background:var(--head);color:var(--head-ink);padding:18px 16px 16px}
.wrap{max-width:1180px;margin:0 auto}
header h1{margin:0;font-size:20px;letter-spacing:.01em}
header h1 b{color:var(--accent)}
header p{margin:4px 0 0;color:var(--head-muted)}
.meta{display:flex;flex-wrap:wrap;gap:6px 22px;margin-top:12px}
.meta div{display:grid}.meta span{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--head-muted)}
.meta strong{font-size:14px}
main{padding:20px 16px 40px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}
.kpi{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:12px 14px;display:grid;gap:2px}
.kpi span{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:var(--muted)}
.kpi strong{font-size:22px;font-variant-numeric:tabular-nums}
.kpi.hl{background:var(--head);border-color:var(--head);color:var(--head-ink)}.kpi.hl span{color:var(--head-muted)}
.kpi small{color:var(--muted);font-size:11.5px}.kpi.hl small{color:var(--head-muted)}
h2{font-size:15px;margin:26px 0 4px}
.sub{color:var(--muted);margin:0 0 10px;font-size:12.5px}
.panel{background:var(--surface);border:1px solid var(--line);border-radius:8px;padding:14px 16px}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:12px;color:var(--muted);margin-bottom:10px}
.legend i{display:inline-block;width:12px;height:10px;border-radius:2px;margin-right:6px;vertical-align:-1px}
.hbars{display:grid;gap:4px}
.hb{display:grid;grid-template-columns:minmax(96px,150px) minmax(0,1fr) 120px;gap:10px;align-items:center;font-size:12px;padding:1px 4px;border-radius:4px}
.hb:hover{background:var(--sunken)}
.hb .lab{font-family:Consolas,monospace;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.hb .track{position:relative;height:12px}
.hb .seg{position:absolute;top:0;bottom:0;border-radius:0}
.hb .seg.free{background:var(--free);border-radius:3px 0 0 3px}
.hb .seg.over{background:var(--accent);border-radius:0 3px 3px 0}
.hb .seg.free.only{border-radius:3px}
.hb .seg.bar{background:var(--accent);border-radius:0 3px 3px 0}
.hb .ftm{position:absolute;top:-3px;bottom:-3px;width:0;border-left:2px dashed var(--ink);opacity:.55}
.hb .val{text-align:right;font-variant-numeric:tabular-nums;color:var(--muted)}
.hb .val b{color:var(--ink)}
.hbars.wide .hb{grid-template-columns:minmax(200px,340px) minmax(0,1fr) 130px}
.hbars.wide .hb .lab{white-space:normal;overflow:visible;line-height:1.3}
.cols{display:flex;align-items:flex-end;gap:6px;height:170px;padding-top:18px;border-bottom:1px solid var(--line)}
.col{flex:1;min-width:20px;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;position:relative}
.col .b{width:70%;max-width:44px;background:var(--rf);border-radius:4px 4px 0 0}
.col .b.etd{background:var(--accent)}
.col .c{font-size:11.5px;font-variant-numeric:tabular-nums;margin-bottom:3px}
.colx{display:flex;gap:6px}.colx div{flex:1;min-width:20px;text-align:center;font-size:11px;color:var(--muted);padding-top:4px}
.scroll{overflow-x:auto;border:1px solid var(--line);border-radius:8px;background:var(--surface)}
table{border-collapse:collapse;width:100%}
th{font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);text-align:left;padding:8px 10px;background:var(--sunken);border-bottom:1px solid var(--line);white-space:nowrap}
td{padding:6px 10px;border-bottom:1px solid var(--line);white-space:nowrap}
tr:last-child td{border-bottom:0}
td.r,th.r{text-align:right}td.r{font-variant-numeric:tabular-nums}
tr.paid td.amt{color:var(--accent-ink);font-weight:bold}
tfoot td{background:var(--sunken);font-weight:bold;border-top:1px solid var(--line)}
.chip{display:inline-block;font-size:11px;font-weight:bold;padding:0 7px;border-radius:999px}
.chip.dg{background:var(--warn-soft);color:var(--warn)}.chip.rf{color:var(--rf);border:1px solid var(--rf)}
.warns{background:var(--warn-soft);color:var(--warn);border-radius:8px;padding:10px 14px;margin-top:14px;display:grid;gap:2px;font-size:12.5px}
a{color:var(--rf)}
.two{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
@media (max-width:820px){.two{grid-template-columns:minmax(0,1fr)}.hb{grid-template-columns:96px minmax(0,1fr) 92px}.hbars.wide .hb{grid-template-columns:minmax(0,1fr) 110px;row-gap:3px}.hbars.wide .hb .lab{grid-column:1/-1}}
footer{color:var(--muted);font-size:12px;margin-top:24px}
"""


CHIP_DG = '<span class="chip dg">DG</span>'
CHIP_RF = '<span class="chip rf">RF</span>'


def e(v):
    return html.escape("" if v is None else str(v))


def money(v):
    return f"{v:,.0f}"


def dmy(d, short=False):
    d = to_date(d)
    if not d:
        return "–"
    return d.strftime("%d/%m") if short else d.strftime("%d/%m/%Y")


def page(title, header_html, body):
    return (f'<!doctype html><html lang="th"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>{e(title)}</title><style>{DASH_CSS}</style></head><body>'
            f'<header><div class="wrap">{header_html}</div></header><main><div class="wrap">{body}'
            f'<footer>สร้างโดย storage_calc.py · {datetime.now():%d/%m/%Y %H:%M}</footer></div></main></body></html>')


def voyage_dashboard(meta, rows):
    s = summarize(rows)
    cs = sorted(rows, key=lambda c: (-(c["Total Stay"] or 0), c["Container No."]))
    hdr = (f'<h1>Storage Dashboard · <b>{e(meta["vsl"])} {e(meta["voy"])}</b></h1><p>{e(meta["name"])}</p>'
           '<div class="meta">' + "".join(
               f'<div><span>{k}</span><strong>{e(v)}</strong></div>' for k, v in
               [("ETD", dmy(meta["etd"])), ("SVC", meta["svc"]), ("POL", meta["pol"]), ("LWharf", meta["lwharf"]),
                ("Terminal", meta["terminal"]), ("Area", meta["area"])]) + "</div>")
    kpis = (f'<div class="kpis">'
            f'<div class="kpi"><span>ตู้ทั้งหมด</span><strong>{s["n"]}</strong><small>20\' {s["c20"]} · 40\' {s["c40"]}</small></div>'
            f'<div class="kpi"><span>เกิน free time</span><strong>{s["over"]}</strong><small>DG {s["dg"]} · RF {s["rf"]}</small></div>'
            f'<div class="kpi"><span>วันคิดเงินรวม</span><strong>{s["days"]}</strong><small>วัน</small></div>'
            f'<div class="kpi"><span>ค่า Storage</span><strong>{money(s["charge"])}</strong><small>THB</small></div>'
            f'<div class="kpi"><span>เก็บลูกค้า</span><strong>{money(s["cust"])}</strong><small>THB</small></div>'
            f'<div class="kpi hl"><span>HAL จ่าย</span><strong>{money(s["hal"])}</strong><small>THB</small></div></div>')

    # stay vs free time per container
    mx = max([c["Total Stay"] or 0 for c in cs] + [max(c["Normal F/T Terminal"] for c in cs) if cs else 1, 1])
    bars = []
    for c in cs:
        st, ft = c["Total Stay"] or 0, c["Normal F/T Terminal"]
        free, over = min(st, ft), max(st - ft, 0)
        tip = (f'{c["Container No."]} · {c["Type Size"]} · เข้า {dmy(c["Date In"])} · อยู่ {st} วัน · '
               f'F/T {ft} วัน · คิดเงิน {over} วัน · {money(c["Total Storage Charge"])} บาท')
        segs = (f'<span class="seg free{" only" if not over else ""}" style="left:0;width:{free / mx * 100:.2f}%"></span>'
                + (f'<span class="seg over" style="left:{free / mx * 100:.2f}%;width:{over / mx * 100:.2f}%"></span>'
                   if over else "")
                + f'<span class="ftm" style="left:{ft / mx * 100:.2f}%"></span>')
        val = (f'<b>{money(c["Total Storage Charge"])}</b> ฿' if c["Total Storage Charge"] else f'{st} วัน')
        bars.append(f'<div class="hb" title="{e(tip)}"><span class="lab">{e(c["Container No."])}</span>'
                    f'<span class="track">{segs}</span><span class="val">{val}</span></div>')
    stay_chart = (f'<h2>วันอยู่ใน terminal เทียบ free time</h2><p class="sub">ต่อตู้ เรียงจากอยู่นานสุด · '
                  f'สเกล 0–{mx} วัน · ชี้ที่แถวเพื่อดูรายละเอียด</p><div class="panel">'
                  '<div class="legend"><span><i style="background:var(--free)"></i>ภายใน free time</span>'
                  '<span><i style="background:var(--accent)"></i>เกิน free time (คิดเงิน)</span>'
                  '<span><i style="border-left:2px dashed var(--ink);width:2px;opacity:.55"></i>เส้น free time</span>'
                  f'</div><div class="hbars">{"".join(bars)}</div></div>')

    # gate-in per day
    cnt = defaultdict(int)
    for c in cs:
        if c["Date In"]:
            cnt[to_date(c["Date In"])] += 1
    etd = to_date(meta["etd"])
    days_html = ""
    if cnt:
        d0 = min(cnt)
        d1 = max(max(cnt), etd or d0)
        span = [date.fromordinal(o) for o in range(d0.toordinal(), d1.toordinal() + 1)]
        cmax = max(cnt.values())
        colhtml = "".join(
            f'<div class="col" title="{dmy(d)}: {cnt.get(d, 0)} ตู้"><span class="c">{cnt.get(d) or ""}</span>'
            f'<span class="b{" etd" if d == etd else ""}" style="height:{(cnt.get(d, 0) / cmax * 100):.1f}%"></span></div>'
            for d in span)
        xhtml = "".join(f'<div>{dmy(d, True)}{"<br>ETD" if d == etd else ""}</div>' for d in span)
        days_html = (f'<h2>จำนวนตู้เข้า terminal ต่อวัน</h2><p class="sub">นับจากวัน Gate in · '
                     f'ETD {dmy(etd)}</p><div class="panel"><div class="cols">{colhtml}</div>'
                     f'<div class="colx">{xhtml}</div></div>')

    trs = "".join(
        f'<tr class="{"paid" if c["Total Storage Charge"] else ""}"><td>{e(c["Container No."])}</td><td>{e(c["Booking No."])}</td>'
        f'<td>{e(c["Shipper"])}</td><td>{e(c["Type Size"])}</td>'
        f'<td>{CHIP_DG if c["DG"] else ""}{CHIP_RF if c["RF"] else ""}</td>'
        f'<td>{dmy(c["Date In"])}</td><td class="r">{c["Total Stay"] if c["Total Stay"] is not None else "–"}</td>'
        f'<td class="r">{c["Normal F/T Terminal"]}</td><td class="r">{c["Total Storage Days"]}</td>'
        f'<td class="r amt">{money(c["Total Storage Charge"])}</td></tr>'
        for c in sorted(rows, key=lambda c: (-(c["Total Storage Charge"] or 0), c["Container No."])))
    table = (f'<h2>รายตู้</h2><div class="scroll"><table><thead><tr><th>Container</th><th>Booking</th><th>Shipper</th>'
             f'<th>Type</th><th></th><th>Date in</th><th class="r">อยู่</th><th class="r">F/T</th><th class="r">วันคิดเงิน</th>'
             f'<th class="r">ค่า Storage</th></tr></thead><tbody>{trs}</tbody><tfoot><tr><td colspan="8">{s["n"]} ตู้</td>'
             f'<td class="r">{s["days"]}</td><td class="r">{money(s["charge"])}</td></tr></tfoot></table></div>')
    warns = ("<div class='warns'><strong>ข้อควรตรวจสอบ</strong>" + "".join(f"<span>{e(w)}</span>" for w in meta["warnings"])
             + "</div>") if meta["warnings"] else ""
    links = (f'<p class="sub" style="margin-top:14px">รายงาน Excel: <a href="{e(meta["name"])}.xlsx">{e(meta["name"])}.xlsx</a>'
             f' · <a href="../../DASHBOARD.html">กลับไป Dashboard รวม</a></p>')
    return page(f"Storage {meta['vsl']} {meta['voy']}", hdr, kpis + warns + links + stay_chart + days_html + table)


def master_dashboard(voyages):
    allc = [c for v in voyages for c in v["rows"]]
    s = summarize(allc)
    hdr = (f'<h1>HAL <b>Storage</b> Dashboard</h1><p>รวมทุกเที่ยวเรือที่คำนวณแล้ว · {len(voyages)} เที่ยว</p>')
    tiles = '<div class="kpis">'
    for a in ("BKK", "UTCT", "LCH"):
        t = summarize([c for v in voyages if v["meta"]["area"] == a for c in v["rows"]])
        tiles += (f'<div class="kpi"><span>{a}</span><strong>{money(t["charge"])}</strong>'
                  f'<small>{t["n"]} ตู้ · เกิน F/T {t["over"]} ตู้</small></div>')
    tiles += (f'<div class="kpi"><span>เก็บลูกค้า</span><strong>{money(s["cust"])}</strong><small>THB</small></div>'
              f'<div class="kpi hl"><span>รวม HAL จ่าย</span><strong>{money(s["hal"])}</strong>'
              f'<small>{s["n"]} ตู้ · {len(voyages)} เที่ยว</small></div></div>')

    vs = sorted(voyages, key=lambda v: (v["meta"]["etd"] or "", v["meta"]["name"]), reverse=True)
    mx = max([summarize(v["rows"])["charge"] for v in vs] + [1])
    bars = "".join(
        f'<div class="hb" title="{e(v["meta"]["name"])}: {money(summarize(v["rows"])["charge"])} บาท">'
        f'<span class="lab">{dmy(v["meta"]["etd"], True)} {e(v["meta"]["vsl"])} {e(v["meta"]["voy"])}</span>'
        f'<span class="track"><span class="seg bar" style="left:0;width:{summarize(v["rows"])["charge"] / mx * 100:.2f}%;'
        f'border-radius:0 3px 3px 0"></span></span><span class="val"><b>{money(summarize(v["rows"])["charge"])}</b> ฿</span></div>'
        for v in vs)
    chart = (f'<h2>ค่า Storage ต่อเที่ยวเรือ</h2><p class="sub">THB · เรียงตาม ETD ล่าสุดก่อน</p>'
             f'<div class="panel"><div class="hbars">{bars}</div></div>')

    trs = ""
    for v in vs:
        m, t = v["meta"], summarize(v["rows"])
        link = f'VOYAGES/{m["name"]}/{m["name"]}_DASHBOARD.html'
        trs += (f'<tr class="{"paid" if t["charge"] else ""}"><td><a href="{e(link)}">{e(m["name"])}</a></td><td>{dmy(m["etd"])}</td>'
                f'<td>{e(m["area"])}</td><td>{e(m["terminal"])}</td><td class="r">{t["n"]}</td><td class="r">{t["over"]}</td>'
                f'<td class="r">{t["days"]}</td><td class="r amt">{money(t["charge"])}</td><td class="r">{money(t["cust"])}</td>'
                f'<td class="r">{money(t["hal"])}</td></tr>')
    table = (f'<h2>เที่ยวเรือทั้งหมด</h2><div class="scroll"><table><thead><tr><th>ETD_SVC_VSL_VOY_POL_LWharf</th>'
             f'<th>ETD</th><th>Area</th><th>Terminal</th><th class="r">ตู้</th><th class="r">เกิน F/T</th>'
             f'<th class="r">วันคิดเงิน</th><th class="r">ค่า Storage</th><th class="r">เก็บลูกค้า</th>'
             f'<th class="r">HAL จ่าย</th></tr></thead><tbody>{trs}</tbody><tfoot><tr><td colspan="4">รวม</td>'
             f'<td class="r">{s["n"]}</td><td class="r">{s["over"]}</td><td class="r">{s["days"]}</td>'
             f'<td class="r">{money(s["charge"])}</td><td class="r">{money(s["cust"])}</td><td class="r">{money(s["hal"])}</td>'
             f'</tr></tfoot></table></div>')

    shp = defaultdict(lambda: [0, 0])
    for c in allc:
        if c["Total Storage Charge"]:
            shp[c["Shipper"] or "-"][0] += c["Total Storage Charge"]
            shp[c["Shipper"] or "-"][1] += 1
    top = sorted(shp.items(), key=lambda x: -x[1][0])[:10]
    shp_html = ""
    if top:
        smx = top[0][1][0]
        shp_html = (f'<h2>Shipper ที่มีค่า Storage สูงสุด</h2><p class="sub">10 อันดับแรก · THB</p><div class="panel"><div class="hbars wide">'
                    + "".join(f'<div class="hb" title="{e(k)}: {v[1]} ตู้ · {money(v[0])} บาท"><span class="lab" style="font-family:inherit">{e(k)}</span>'
                              f'<span class="track"><span class="seg bar" style="left:0;width:{v[0] / smx * 100:.2f}%"></span></span>'
                              f'<span class="val"><b>{money(v[0])}</b> · {v[1]} ตู้</span></div>' for k, v in top)
                    + "</div></div>")
    links = '<p class="sub" style="margin-top:14px">ฐานข้อมูล Excel: <a href="DATABASE/STORAGE_DATABASE.xlsx">DATABASE/STORAGE_DATABASE.xlsx</a></p>'
    body = tiles + links + (chart if voyages else '<p class="sub">ยังไม่มีข้อมูล · วางไฟล์ใน INPUT แล้วรัน run_storage.bat</p>') \
        + (table if voyages else "") + shp_html
    return page("HAL Storage Dashboard", hdr, body)


# ---------------------------------------------------------------- persistence
def jsonable(v):
    if isinstance(v, (date, datetime)):
        return v.isoformat()[:10]
    return v


def load_voyages():
    out = []
    for f in sorted(glob.glob(os.path.join(VOY_DIR, "*", "data.json"))):
        with open(f, encoding="utf-8") as fh:
            v = json.load(fh)
        for c in v["rows"]:
            for k in DATE_COLS:
                c[k] = to_date(c.get(k))
        out.append(v)
    return out


def move_sources(paths, dests):
    """คัดลอกไฟล์ต้นฉบับเข้า SOURCE ของทุกเที่ยวที่ใช้ไฟล์นั้น แล้วลบออกจาก INPUT"""
    for p in paths:
        for d in dests[p]:
            os.makedirs(d, exist_ok=True)
            shutil.copy2(p, os.path.join(d, os.path.basename(p)))
        os.remove(p)


# ---------------------------------------------------------------- main
def main():
    for d in (INPUT_DIR, VOY_DIR, DB_DIR):
        os.makedirs(d, exist_ok=True)
    files = sorted(f for f in glob.glob(os.path.join(INPUT_DIR, "*"))
                   if f.lower().endswith((".xls", ".xlsx", ".xlsm")) and not os.path.basename(f).startswith("~$"))
    print(f"INPUT: {len(files)} ไฟล์")
    sources = [classify(f) for f in files]
    for s in sources:
        print(f"  [{s['kind'] or 'X':8s}] {s['name']}" + (f"  !! {s['error']}" if s.get("error") else ""))
    sources = [s for s in sources if not s.get("error")]

    kinds = {s["kind"] for s in sources}
    have = {k: (k in kinds or "TEMPLATE" in kinds) for k in ("LOAD", "BKG", "GATE")}
    processed = 0
    if sources and not all(have.values()):
        print("\nยังไม่คำนวณ: ขาดไฟล์ " + ", ".join(k for k, v in have.items() if not v)
              + " · ไฟล์ทั้งหมดยังอยู่ใน INPUT")
    elif sources:
        rows, warnings, heads = compute(sources)
        sets = defaultdict(list)
        for r in rows:
            sets[r["key"]].append(r)
        dests = defaultdict(list)
        for key, srows in sorted(sets.items(), key=lambda kv: (kv[0][0] or date.min, str(kv[0]))):
            name = voyage_name(key)
            folder = os.path.join(VOY_DIR, name)
            os.makedirs(folder, exist_ok=True)
            first = srows[0]["cs"]
            locs = sorted({r["cs"]["Location"] for r in srows if r["cs"]["Location"]})
            src_idx = set().union(*(r["srcs"] for r in srows))
            cset = {r["cs"]["Container No."] for r in srows}
            meta = {"name": name, "etd": jsonable(key[0]), "svc": key[1], "vsl": key[2], "voy": key[3],
                    "pol": key[4], "lwharf": key[5], "area": srows[0]["area"], "group": first["Group"],
                    "terminal": " / ".join(TERMINAL_NAME.get(l, l) for l in locs),
                    "generated": f"{datetime.now():%d/%m/%Y %H:%M}",
                    "sources": sorted(sources[i]["name"] for i in src_idx),
                    "warnings": [w for w in warnings if w.split(":")[0] in cset]}
            write_voyage_xlsx(os.path.join(folder, f"{name}.xlsx"), name, meta, srows, heads)
            data = {"meta": meta, "rows": [{k: jsonable(v) for k, v in r["cs"].items()} for r in srows]}
            with open(os.path.join(folder, "data.json"), "w", encoding="utf-8") as fh:
                json.dump(data, fh, ensure_ascii=False, indent=1)
            with open(os.path.join(folder, f"{name}_DASHBOARD.html"), "w", encoding="utf-8") as fh:
                rows_d = [dict(r["cs"]) for r in srows]
                fh.write(voyage_dashboard(meta, rows_d))
            for i in src_idx:
                dests[sources[i]["path"]].append(os.path.join(folder, "SOURCE"))
            s = summarize(srows)
            print(f"\n  ✓ {name}\n    {s['n']} ตู้ · เกิน F/T {s['over']} · Storage {money(s['charge'])} · HAL {money(s['hal'])}")
            for w in meta["warnings"]:
                print("    ⚠", w)
            processed += 1
        used = [p for p in dests]
        move_sources(used, dests)
        left = [s["name"] for s in sources if s["path"] not in dests]
        if left:
            print("\nไฟล์ที่ไม่ได้ใช้ (ยังอยู่ใน INPUT):", ", ".join(left))

    voyages = load_voyages()
    write_database(voyages)
    with open(MASTER_DASH, "w", encoding="utf-8") as fh:
        fh.write(master_dashboard(voyages))
    print(f"\nคำนวณ {processed} เที่ยว · ฐานข้อมูลมีทั้งหมด {len(voyages)} เที่ยว")
    print("  ฐานข้อมูล :", DB_FILE)
    print("  Dashboard :", MASTER_DASH)
    return 0


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
