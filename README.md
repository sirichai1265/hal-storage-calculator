# HAL Storage Calculator

คำนวณค่าฝากตู้ (Storage) ที่ terminal สำหรับตู้ส่งออก แยกพื้นที่ **BKK · UTCT · LCH**
จากไฟล์ระบบ LOAD / BKG / GATE ของแต่ละเที่ยวเรือ

มี 2 ส่วน

| ส่วน | ไฟล์ | ใช้ทำอะไร |
|---|---|---|
| โปรแกรมคำนวณ + เก็บข้อมูล | `storage_calc.py`, `run_storage.bat` | อ่านไฟล์ใน `INPUT` → สร้างรายงาน Excel, dashboard และฐานข้อมูล แยกโฟลเดอร์ตามเที่ยวเรือ |
| แอพบนเว็บ | `storage_app.html` | คำนวณด่วน, คำนวณจากไฟล์, Export Excel, Dashboard รวม (ฐานข้อมูลของแอพ) |

ลิงก์

- หน้าแรก: https://sirichai1265.github.io/hal-storage-calculator/
- Dashboard รวม: https://sirichai1265.github.io/hal-storage-calculator/DASHBOARD.html
- แอพ (GitHub Pages): https://sirichai1265.github.io/hal-storage-calculator/storage_app.html
- แอพ (claude.ai มี Dashboard ฐานข้อมูลร่วม): https://claude.ai/artifact/X1NwK1c6Vf26JNXF4vTHkj

อัปเดต Dashboard ออนไลน์: วางไฟล์ใน `INPUT/` แล้วดับเบิลคลิก `update_dashboard.bat` (คำนวณ + push ขึ้น GitHub)

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
  <ชื่อเที่ยว>_TEMPLATE.xlsx     ไฟล์ STORAGE TEMPLATE ที่เติมข้อมูลครบทุกชีท (เปิดที่ชีท CS)
  SOURCE/                      ไฟล์ต้นฉบับที่ใช้คำนวณ
  data.json                    ข้อมูลสำหรับสร้างฐานข้อมูล
DATABASE/STORAGE_DATABASE.xlsx ฐานข้อมูลรวม (ชีท Voyages, Containers)
DASHBOARD.html                 dashboard รวมทุกเที่ยว
```

- ไฟล์ไม่ครบ 3 ชนิด → ไม่คำนวณ ไฟล์ยังอยู่ใน `INPUT`
- รันเที่ยวเดิมซ้ำ → เขียนทับโฟลเดอร์เดิม ฐานข้อมูลไม่ซ้ำ
- นำ `STORAGE_DATABASE.xlsx` เข้า Dashboard ของแอพได้ (แท็บ Dashboard → นำเข้า)

## แก้ข้อมูลรายตู้ (CORRECTIONS.csv)

ถ้าข้อมูลจากระบบผิด (เช่น Location ใน LOAD ไม่ตรงกับที่ตู้เข้าจริง) ให้เพิ่มแถวใน `CORRECTIONS.csv` แล้วรันเที่ยวนั้นใหม่
(คัดลอกไฟล์จาก `VOYAGES/<ชื่อเที่ยว>/SOURCE/` กลับไปที่ `INPUT/`)

| Container No. | Booking No. | Location | Remark |
|---|---|---|---|
| SKLU1952893 | HASLS22260800402 | LCH04 | ตู้เข้า LCMT จริงตาม GATE และ BKG LWharf |

- ใช้คู่ Container No. + Booking No. จึงไม่กระทบตู้เดียวกันในเที่ยวอื่น
- ไฟล์ต้นฉบับไม่ถูกแก้ · รายงานบันทึกการแก้ไว้ในคอลัมน์ Remark

## Export จากแท็บ Dashboard ในแอพ

- **Export** ท้ายแถวเที่ยวเรือ → ไฟล์ STORAGE TEMPLATE ของเที่ยวนั้น (เหมือนไฟล์ `_TEMPLATE.xlsx` ใน VOYAGES ทุกช่อง)
- **Export STORAGE TEMPLATE (.zip)** → ทุกเที่ยวตามตัวกรองพื้นที่/เดือน เที่ยวละ 1 ไฟล์ รวมเป็น .zip
- **ไฟล์สรุป BKK/LCH** → ไฟล์สรุปหลายเที่ยว (ชีท BKK = BKK + UTCT, ชีท LCH)
- ข้อมูลดิบต่อเที่ยวเก็บใน collection `voyage_raw` ของแอพ (`<ชื่อเที่ยว>__loading/__booking/__gate`)
  สร้างจาก `python storage_calc.py --app-raw <โฟลเดอร์> [ชื่อเที่ยว ...]` หรือบันทึกจากแท็บ "คำนวณจากไฟล์เรือ"

ข้อมูลสำหรับฐานข้อมูลแอพ: `python storage_calc.py --app-docs <โฟลเดอร์> [ชื่อเที่ยว ...]` → ไฟล์ JSON ต่อเที่ยว
สร้างไฟล์ `_TEMPLATE.xlsx` ใหม่จาก SOURCE ของเที่ยวเดิม: `python storage_calc.py --template-files [ชื่อเที่ยว ...]`

## ไฟล์ BKG รูปแบบใหม่ (104 คอลัมน์)

ตั้งแต่ 14/10/2026 ไฟล์ BKG บางชุดมีคอลัมน์ BN "Roll Over" เพิ่มเข้ามา และย้าย "Roll Over Remark" มาไว้ที่ BP
โปรแกรมและแอพปรับให้อัตโนมัติ: **ลบคอลัมน์ BN และย้าย BP ไปต่อท้ายหลัง CZ** ได้รูปแบบเดิม 103 คอลัมน์
(ไฟล์ต้นฉบับไม่ถูกแก้ · ใช้กับไฟล์ TEMPLATE ที่วางข้อมูลรูปแบบใหม่ในชีท 2.Bookinglist ด้วย)

## บันทึกเข้า Dashboard จาก HAL Storage Calculator

1. ใส่ไฟล์ LOAD / BKG / GATE ในแท็บ "คำนวณจากไฟล์เรือ" (ไฟล์เดียวมีหลายลำได้) → แอพคำนวณและแสดงตารางแยกตามเที่ยวเรือ
   **ยังไม่บันทึกอะไรเข้า Dashboard** จนกว่าจะกดปุ่ม
2. ติ๊กเลือกลำที่ต้องการในตาราง (ค่าเริ่มต้น = ทุกลำที่ข้อมูลครบ · ลำที่ข้อมูลไม่ครบเลือกไม่ได้) คอลัมน์ "สถานะ Dashboard" บอกว่า ใหม่ / ต่างจาก Dashboard / ตรงกับ Dashboard
3. กด **บันทึกเข้า Dashboard (n เที่ยว)** → บันทึกเฉพาะลำที่เลือก
   - ถ้าติ๊ก "ส่งขึ้น GitHub ด้วย (บันทึกไฟล์)" แอพจะเปิดหน้าต่างบันทึกไฟล์ STORAGE TEMPLATE ทีละลำ → กด Save → `storage_watch.py` ในเครื่องซิงก์ขึ้น GitHub Dashboard ใน ~1 นาที
   - `storage_watch.py` เริ่มเองตอนเปิดเครื่อง (`HAL Storage Sync.vbs` ในโฟลเดอร์ Startup) · ซิงก์เฉพาะเมื่อข้อมูลเปลี่ยน · ดู `watch.log`
   - แก้ `storage_calc.py` แล้วให้รีสตาร์ต: ปิด pythonw ที่รัน storage_watch.py แล้วดับเบิลคลิก `HAL Storage Sync.vbs`
- ข้อมูลดิบ (LOAD/BKG/GATE) ที่เก็บต่อลำ: ไฟล์ที่มีแต่ข้อมูลลำเดียว → ทั้งไฟล์ · ไฟล์หลายลำ → เฉพาะแถวของลำนั้น
  (BKG ตาม VSL+VOY+POL · GATE ตามเลขตู้+booking เพราะตู้ที่เคยเข้า booking เก่า สูตร VLOOKUP ใน template จะหยิบแถวเก่า)
- CORRECTIONS ใช้ในแอพด้วย (collection `corrections` ในฐานข้อมูลแอพ ต้องตรงกับ `CORRECTIONS.csv`)
- นำเที่ยวจากฐานข้อมูลแอพเข้าระบบด้วยมือ: ดาวน์โหลด collection `voyage_raw` แล้ว `python storage_calc.py --import-app <โฟลเดอร์>`

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
