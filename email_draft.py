# -*- coding: utf-8 -*-
"""สร้าง/อัปเดตดราฟอีเมล์ Outlook ของเที่ยวเรือ (ไม่ส่ง) — แนบ _TEMPLATE.xlsx + แปะ Full Schedule จาก e-heunga.com (ETA/ETB/ETD)
ใช้:  python email_draft.py <ชื่อเที่ยว>   เช่น  python email_draft.py 2026-10-22_KST_SWAL_2609N_THBKK_BKK02
"""
import glob
import html
import json
import os
import subprocess
import sys
import tempfile
import urllib.request

import storage_calc as sc

BASE = os.path.dirname(os.path.abspath(__file__))
API = "https://ebizapi.sinokor.co.kr"
TO = "logistics@heungaline.co.th; csoutbound@heungaline.co.th; sales@heungaline.co.th; OPRLMH@heungaline.co.th; document_outbound@heungaline.co.th"
ST = "font-family:Tahoma;font-size:11px;border:1px solid #888;padding:3px 8px"
TH = 'style="font-family:Tahoma;font-size:11px;border:1px solid #888;background:#1F3A4D;color:#fff;padding:3px 8px"'


def post(path, data):
    req = urllib.request.Request(API + path, json.dumps(data).encode(), {"Content-Type": "application/json"})
    return json.load(urllib.request.urlopen(req, timeout=30))


def fmt(s):
    d, t = s.split(" ")
    y, m, dd = d.split("-")
    return f"{dd}/{m}/{y} {t}"


def full_schedule(vsl, voy, pol, ports):
    """แถวตารางเรือจาก POL ของเที่ยวนี้ ไปจนถึงท่าสุดท้ายที่มี booking (POD/TS) ก่อนเรือวนกลับ POL"""
    rows, seen = [], set()
    y, m = None, None
    first = None
    for mo in range(0, 3):
        pass
    import datetime
    t = datetime.date.today().replace(day=1)
    months = []
    for k in range(-1, 3):
        mm = t.month + k
        yy = t.year + (mm - 1) // 12
        mm = (mm - 1) % 12 + 1
        months.append(f"{yy}{mm:02d}")
    for mo in months:
        for r in post("/Schedule/Vessel", {"compcd": "HALK", "nacd": "", "token": "", "vsl": vsl, "month": mo}):
            k = (r["VYG"], r["PORT"], r["ETA"])
            if k not in seen:
                seen.add(k)
                rows.append(r)
    rows.sort(key=lambda r: r["ETA"])
    start = next((i for i, r in enumerate(rows) if r["VYG"] == voy and r["PORT"] == pol), None)
    if start is None:
        return None
    seq = []
    for r in rows[start:]:
        if seq and r["PORT"] == pol:
            break
        seq.append(r)
    firsts = {}
    for i, r in enumerate(seq):
        firsts.setdefault(r["PORT"], i)
    last = max((i for p, i in firsts.items() if p in ports), default=0)   # ท่าแรกที่เรือแวะของแต่ละ POD/TS (ไม่นับรอบที่ 2)
    return seq[: last + 1]


def main(name):
    meta = json.load(open(os.path.join(sc.VOY_DIR, name, "data.json"), encoding="utf-8"))["meta"]
    vsl, voy, pol, svc = meta["vsl"], meta["voy"], meta["pol"], meta["svc"]
    bkg = glob.glob(os.path.join(sc.VOY_DIR, name, "SOURCE", "*BKG*.xls*"))
    ports = set()
    for f in bkg:
        for r in sc.classify(f)["booking"]:
            if str(r[0]).startswith("HASLS") and (r[12], r[13], r[14]) == (vsl, voy, pol):
                ports.add(r[18])
                if r[25]:
                    ports.add(r[25])
                svc = r[11] or svc
    seq = full_schedule(vsl, voy, pol, ports)
    if not seq:
        raise SystemExit("ไม่พบ schedule ของ %s %s %s บน e-heunga.com" % (vsl, voy, pol))
    body = ""
    for r in seq:
        b = ' style="background:#FFF2CC"' if r["PORT"] == pol and r["VYG"] == voy else ""
        body += (f'<tr{b}><td style="{ST}">{r["VYG"]}</td><td style="{ST}">{r["PORT"]}</td><td style="{ST}">{html.escape(r["WHARFNM"])}</td>'
                 f'<td style="{ST}">{fmt(r["ETA"])}</td><td style="{ST}">{fmt(r["ETB"])}</td><td style="{ST}">{fmt(r["ETD"])}</td></tr>')
    h = ('<div style="font-family:Tahoma;font-size:11px"><p>Dear all,</p>'
         '<p>Please find attached the full storage calculation for the vessel below.</p>'
         f'<p><b>Schedule</b> &nbsp; {svc} / {seq[0]["VSLNM"]} ({vsl}) {voy}</p>'
         f'<table style="border-collapse:collapse"><tr><th {TH}>Voyage</th><th {TH}>Port</th><th {TH}>Terminal</th>'
         f'<th {TH}>ETA</th><th {TH}>ETB</th><th {TH}>ETD</th></tr>{body}</table></div>')
    subject = f"{svc} / FULL STORAGE / {vsl} {voy} / {pol}"
    tmp = os.path.join(tempfile.gettempdir(), "storage_mail.html")
    open(tmp, "w", encoding="utf-8").write(h)
    att = os.path.join(sc.VOY_DIR, name, f"{name}_TEMPLATE.xlsx")
    ps = os.path.join(BASE, "draft_outlook.ps1")
    out = subprocess.run(["powershell", "-NoProfile", "-File", ps, "-Html", tmp, "-Attach", att, "-Sub", subject, "-To", TO],
                         capture_output=True, text=True, encoding="utf-8", errors="replace")
    print(out.stdout.strip() or out.stderr.strip())
    for r in seq:
        print(" ", r["VYG"], r["PORT"], fmt(r["ETA"]), fmt(r["ETB"]), fmt(r["ETD"]))


if __name__ == "__main__":
    main(sys.argv[1])
