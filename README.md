# HAL Storage Calculator

คำนวณค่าฝากตู้ (Storage) ที่ terminal สำหรับตู้ส่งออก แยกพื้นที่ **BKK · UTCT · LCH**
จากไฟล์ระบบ LOAD / BKG / GATE ของแต่ละเที่ยวเรือ

มี 2 ส่วน

| ส่วน | ไฟล์ | ใช้ทำอะไร |
|---|---|---|
| โปรแกรมคำนวณ + เก็บข้อมูล | `storage_calc.py`, `run_storage.bat` | อ่านไฟล์ใน `INPUT` → สร้างรายงาน Excel, dashboard และฐานข้อมูล แยกโฟลเดอร์ตามเที่ยวเรือ |
| แอพบนเว็บ | `storage_app.html` | คำนวณด่วน, คำนวณจากไฟล์, Export Excel, Dashboard รวม (ฐานข้อมูลของแอพ) |

แอพที่เผยแพร่แล้ว: https://claude.ai/artifact/X1NwK1c6Vf26JNXF4vTHkj

## ติดตั้ง

```bash
pip install -r requirements.txt
```

## วิธีใช้ (โปรแกรมคำนวณ)

1. วางไฟล์ของเที่ยวเรือ เช่น `9-30-KMGY-2610N-LCH-LOAD.xls`, `...-BKG.xls`, `...-GATE.xls` ลงใน `INPUT/`
   (ใช้ไฟล์ TEMPLATE ที่กรอกชีท 1.Loading / 2.Bookinglist / 3.GateMove แล้วก็ได้)
2. ดับเบิลคลิก `run_storage.bat` หรือรัน `python storage_calc.py`

ผลลัพธ์ (ชื่อเที่ยว = `ETD_SVC_VSL_VOY_POL_LWharf`)

```
INPUT/                         ไฟล์รอคำนวณ (ว่างหลังรันสำเร็จ)
VOYAGES/<ชื่อเที่ยว>/
  <ชื่อเที่ยว>.xlsx             รายงาน Tahoma 8 มีเส้นตาราง: Summary, CS, 1.Loading, 2.Bookinglist, 3.GateMove
  <ชื่อเที่ยว>_DASHBOARD.html   dashboard ของเที่ยว
  SOURCE/                      ไฟล์ต้นฉบับที่ใช้คำนวณ
  data.json                    ข้อมูลสำหรับสร้างฐานข้อมูล
DATABASE/STORAGE_DATABASE.xlsx ฐานข้อมูลรวม (ชีท Voyages, Containers)
DASHBOARD.html                 dashboard รวมทุกเที่ยว
```

- ไฟล์ไม่ครบ 3 ชนิด → ไม่คำนวณ ไฟล์ยังอยู่ใน `INPUT`
- รันเที่ยวเดิมซ้ำ → เขียนทับโฟลเดอร์เดิม ฐานข้อมูลไม่ซ้ำ
- นำ `STORAGE_DATABASE.xlsx` เข้า Dashboard ของแอพได้ (แท็บ Dashboard → นำเข้า)

## หลักการคำนวณ

- วันอยู่ใน terminal = ETD − วันตู้เข้า (Gate in) + 1
- วันคิดเงิน = วันอยู่ − Free time
- ค่า Storage = ยอดสะสมของขั้น + (วันคิดเงิน − วันเริ่มขั้น) × อัตราต่อวัน

| Group | Terminal | Location |
|---|---|---|
| G1 | LCMT | LCH04 |
| G2 | ESCO · TIPS · LCIT · Hutchison | LCH01 02 03 05 06 08 09 10 |
| G3 | PAT | BKK01 (Terminal 2) · BKK04 (Terminal 1) |
| G4 | Unithai (UTCT) — อัตราเดียว | BKK02 |

Free time terminal (ลำดับการเลือก)

1. ค่าที่กรอกเองในชีท CS ของ TEMPLATE / Special F/T (TML)
2. BKK01 / BKK04 ที่ CGO TERM = CFS/CY → 0 วัน
3. DG → 1 วัน
4. LCH ทุก terminal → 7 วัน · BKK02 → 5 วัน · BKK01/04 → 3 วัน

อัตราอยู่ที่ `TIERS` ใน `storage_calc.py` และ `storage_app.html` (ต้องแก้ทั้งสองที่ให้ตรงกัน)

## ไฟล์ข้อมูลใน repo

| ไฟล์ / โฟลเดอร์ | คืออะไร |
|---|---|
| `TEMPLATE.XLSX` | แม่แบบ Excel เดิม (ชีท 1.Loading, 2.Bookinglist, 3.GateMove, CS, Charge) |
| `messageImage_*.jpg` | ตาราง Tariff ต้นฉบับ |
| `archive/*.xlsx` | ไฟล์เที่ยวเรือตัวอย่างแบบ TEMPLATE (UTCT, BKK, LCH) ใช้ตรวจผลเทียบ Excel |
| `archive/backup_*` | ผลคำนวณเก่าที่ย้ายออกตอนล้างข้อมูล |
| `INPUT/`, `VOYAGES/`, `DATABASE/`, `DASHBOARD.html` | โฟลเดอร์ทำงาน (โปรแกรมสร้างให้เองถ้าไม่มี) |
