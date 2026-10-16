# -*- coding: utf-8 -*-
"""
ซิงก์ HAL Storage Dashboard บน GitHub อัตโนมัติ เมื่อบันทึกไฟล์ STORAGE TEMPLATE จากแอพ HAL Storage Calculator
-------------------------------------------------------------------------------------------------
- ดูโฟลเดอร์ Desktop และ Downloads ทุก 10 วินาที หาไฟล์ชื่อแบบเที่ยวเรือ  ETD_SVC_VSL_VOY_POL_LWharf.xlsx
  (เช่น 2026-10-15_VTS_SWSR_2611N_THLCH_LCH04.xlsx หรือ "... (1).xlsx")
- ถ้าข้อมูลในไฟล์ต่างจากที่อยู่ใน VOYAGES (หรือเป็นเที่ยวใหม่) → คัดลอกเข้า INPUT, รัน storage_calc.py,
  git commit + push → GitHub Pages อัปเดตใน ~1 นาที · ข้อมูลเหมือนเดิม = ไม่ทำอะไร
- ไม่ใช้ Claude · บันทึกการทำงานที่ watch.log

รัน:  pythonw storage_watch.py   (เริ่มอัตโนมัติตอนเปิดเครื่องผ่าน Startup: HAL Storage Sync.vbs)
ทดสอบ 1 รอบ:  python storage_watch.py --once
"""
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import storage_calc as sc  # noqa: E402

HOME = os.path.expanduser("~")
WATCH_DIRS = [d for d in os.environ.get("STORAGE_WATCH_DIRS", "").split(";") if d] or \
    [os.path.join(HOME, "Desktop"), os.path.join(HOME, "Downloads")]
NAME_RE = re.compile(r"^(\d{4}-\d{2}-\d{2}_[A-Z0-9]+_[A-Z0-9]+_[A-Z0-9]+_[A-Z]{5}_[A-Z0-9]{3,6})(?: ?\(\d+\))?\.xlsx$", re.I)
STATE = os.path.join(BASE, ".watch_state.json")
LOG = os.path.join(BASE, "watch.log")
INTERVAL = 10


def log(msg):
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line, flush=True)
    with open(LOG, "a", encoding="utf-8") as fh:
        fh.write(line + "\n")


def load_state():
    try:
        with open(STATE, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:  # noqa: BLE001
        return {}


def save_state(st):
    with open(STATE, "w", encoding="utf-8") as fh:
        json.dump(st, fh, ensure_ascii=False, indent=1)


def sig_rows(pairs):
    return sorted(pairs)


def file_signature(src):
    """(ชื่อเที่ยว, ลายเซ็นผลคำนวณ) จากไฟล์ — ใช้ตรรกะเดียวกับ storage_calc.py"""
    rows, _, _ = sc.compute([src])
    if not rows:
        return None, None
    names = {sc.voyage_name(r["key"]) for r in rows}
    if len(names) != 1:
        return None, None
    name = names.pop()
    pairs = [(r["cs"]["Container No."], r["cs"]["Location"], r["cs"]["Total Stay"], r["cs"]["Normal F/T Terminal"],
              r["cs"]["Total Storage Days"], r["cs"]["Total Storage Charge"]) for r in rows]
    return name, sig_rows(pairs)


def stored_signature(name):
    p = os.path.join(sc.VOY_DIR, name, "data.json")
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        v = json.load(fh)
    return sig_rows([(r["Container No."], r["Location"], r["Total Stay"], r["Normal F/T Terminal"],
                      r["Total Storage Days"], r["Total Storage Charge"]) for r in v["rows"]])


def git(*args):
    r = subprocess.run(["git", *args], cwd=BASE, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)}: {r.stderr.strip() or r.stdout.strip()}")
    return r.stdout.strip()


def sync_file(path, name):
    if [f for f in os.listdir(sc.INPUT_DIR) if not f.startswith(".")]:
        log(f"ข้าม {name}: โฟลเดอร์ INPUT มีไฟล์ค้างอยู่ (จะลองใหม่รอบถัดไป)")
        return False
    shutil.copy2(path, os.path.join(sc.INPUT_DIR, f"{name}_FROM_APP.xlsx"))
    r = subprocess.run([sys.executable.replace("pythonw.exe", "python.exe"), "storage_calc.py"], cwd=BASE,
                       capture_output=True, text=True, encoding="utf-8", errors="replace", stdin=subprocess.DEVNULL)
    if r.returncode != 0 or not os.path.exists(os.path.join(sc.VOY_DIR, name, "data.json")):
        log(f"คำนวณ {name} ไม่สำเร็จ: {(r.stderr or r.stdout)[-300:]}")
        return False
    with open(os.path.join(sc.VOY_DIR, name, "data.json"), encoding="utf-8") as fh:
        v = json.load(fh)
    total = sum(x["Total Storage Charge"] for x in v["rows"])
    git("add", "-A", "VOYAGES", "DATABASE", "DASHBOARD.html")
    if not git("diff", "--cached", "--name-only"):
        log(f"{name}: ไม่มีอะไรเปลี่ยน")
        return True
    git("commit", "-m", f"Sync from app: {name} ({len(v['rows'])} containers, storage {total:,.0f})\n\n"
                        "Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>")
    git("push", "origin", "main")
    log(f"✓ ซิงก์ {name} ขึ้น GitHub แล้ว · {len(v['rows'])} ตู้ · Storage {total:,.0f} · commit {git('rev-parse', '--short', 'HEAD')}")
    return True


def scan(state):
    for d in WATCH_DIRS:
        if not os.path.isdir(d):
            continue
        for f in os.listdir(d):
            m = NAME_RE.match(f)
            if not m:
                continue
            path = os.path.join(d, f)
            try:
                st = os.stat(path)
            except OSError:
                continue
            key = f"{path}|{int(st.st_mtime)}|{st.st_size}"
            if state.get(path) == key or time.time() - st.st_mtime < 3:      # ยังเขียนไฟล์อยู่ / เคยดูแล้ว
                continue
            try:
                src = sc.classify(path)
                if src.get("error") or src["kind"] != "TEMPLATE" or not src["loading"]:
                    log(f"ข้าม {f}: ไม่ใช่ไฟล์ STORAGE TEMPLATE ที่มีข้อมูลเที่ยวเรือ ({src.get('error') or src['kind']})")
                    state[path] = key
                    continue
                name, sig = file_signature(src)
            except PermissionError:
                continue                                                        # เปิดอยู่ใน Excel: ลองใหม่รอบหน้า
            except Exception as e:  # noqa: BLE001
                log(f"อ่าน {f} ไม่ได้: {e}")
                state[path] = key
                continue
            if not name:
                log(f"ข้าม {f}: ไฟล์มีข้อมูลหลายเที่ยวปนกัน (แอพรุ่นเก่า) หรือไม่พบตู้ Full — บันทึกใหม่จากแอพรุ่นล่าสุด "
                    "หรือให้ Claude นำเข้าจากฐานข้อมูลแอพ (--import-app)")
                state[path] = key
                continue
            if sig == stored_signature(name):
                state[path] = key                                               # ข้อมูลเหมือนเดิม: ไม่ซิงก์
                continue
            log(f"พบข้อมูลใหม่: {f} → {name}")
            try:
                if sync_file(path, name):
                    state[path] = key
            except Exception as e:  # noqa: BLE001
                log(f"ซิงก์ {name} ไม่สำเร็จ: {e}")
            save_state(state)
    save_state(state)


def single_instance():
    """กันรันซ้ำ 2 ตัว: จองพอร์ตในเครื่อง ถ้าจองไม่ได้แปลว่ามีตัวที่รันอยู่แล้ว"""
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", 47812))
    except OSError:
        return None
    return s


def main():
    once = "--once" in sys.argv
    lock = None if once else single_instance()
    if not once and lock is None:
        return                                   # มีตัวเฝ้าดูรันอยู่แล้ว
    state = load_state()
    if not once:
        log("เริ่มเฝ้าดู Desktop / Downloads")
    while True:
        try:
            scan(state)
        except Exception as e:  # noqa: BLE001
            log(f"ผิดพลาด: {e}")
        if once:
            return
        time.sleep(INTERVAL)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure") and sys.stdout:
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass
    main()
