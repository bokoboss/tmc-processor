# TMC Processor

TMC Processor เป็นเครื่องมือบน Streamlit สำหรับประมวลผลข้อมูล **Turning Movement Count (TMC)** จากไฟล์ Excel ของงานสำรวจจราจร ให้เป็นตารางสรุป กราฟ ตรวจสอบช่วงเร่งด่วน และรายงาน Excel ที่พร้อมนำไปใช้ต่อในงานรายงานด้านวิศวกรรมจราจร

โปรแกรมนี้ออกแบบมาเพื่อช่วยลดงานซ้ำ เช่น การจัดทิศทางจากไฟล์สำรวจ การรวม movement จากหลาย source stream การเลือกช่วง Peak และการสร้างรายงานตามเทมเพลต Excel

## ความสามารถหลัก

- อ่านไฟล์ TMC Excel ที่มี sheet ทิศทาง เช่น `ทิศ 1`, `ทิศ 2`, หรือ `ทิศ 2+3`
- ตรวจพบ sheet ทิศทางและแสดงตัวอย่างก่อนประมวลผล
- กำหนด Mapping จากทิศทาง/stream ในไฟล์สำรวจไปยัง movement มาตรฐาน เช่น `NS`, `NE`, `WU`, `EU`
- รองรับการรวมหลาย source stream ไปยัง movement เดียว เช่น ทางหลักตรง + ทางคู่ขนานตรง → `NS`
- เก็บรายละเอียด trace กลับได้ผ่าน `Movement_Aggregation_Audit`
- คำนวณ PCU จากค่า PCE factor ที่กำหนดไว้
- สร้างตารางสรุปรายชั่วโมง แยก movement แยกประเภทยานพาหนะ สัดส่วนยานพาหนะ และ PHF
- แสดง Dashboard สำหรับตรวจสอบกราฟปริมาณจราจรรวมรายชั่วโมง
- ให้ผู้ใช้ยืนยันหรือ override ช่วง AM/PM Peak ก่อนสร้างรายงาน
- ใช้ช่วง Peak ที่ยืนยันแล้วเป็น source of truth สำหรับทุก output ที่เกี่ยวข้อง
- ส่งออก Excel report ตามเทมเพลต 4-leg TMC
- รองรับ **Excel Template Mode** ผ่าน Microsoft Excel COM เพื่อรักษา Native Chart, สูตร และรูปแบบของเทมเพลต
- มี **Safe PNG Export Mode** เป็นโหมดสำรองเมื่อ Excel COM ใช้งานไม่ได้
- Save / Load Project Session ผ่านไฟล์ `.tmcproj.json`
- UI ภาษาไทย สำหรับ workflow การใช้งานของทีมงานด้านจราจร

## Workflow การใช้งาน

```text
อัปโหลดไฟล์ → ตั้งค่างาน → กำหนดทิศทาง → ประมวลผล → ตรวจสอบช่วงเร่งด่วน → ส่งออกไฟล์
```

ขั้นตอนทั่วไป:

1. อัปโหลดไฟล์ TMC Excel
2. กรอกข้อมูลรายงาน ป้ายปลายทาง ชื่อถนน และข้อมูล setup
3. ตรวจสอบและแก้ไข Mapping
4. ประมวลผลข้อมูล
5. ตรวจสอบ Dashboard และยืนยันช่วง AM/PM Peak
6. สร้างรายงาน Excel
7. ดาวน์โหลดรายงาน หรือบันทึก Project Session ไว้ใช้ต่อ

## วิธีติดตั้งแบบง่ายสำหรับ Windows

ส่วนนี้เขียนสำหรับผู้ใช้ทั่วไปที่ยังไม่คุ้นกับ Python หรือ command line มาก่อน

### สิ่งที่ต้องมี

1. **Windows**
2. **Python** แนะนำเวอร์ชัน 3.11 หรือ 3.12
3. **Microsoft Excel** ถ้าต้องการใช้โหมดรายงานตามเทมเพลต Excel และ Native Chart

> ตอนติดตั้ง Python ให้เลือกตัวเลือก **Add python.exe to PATH** ด้วย เพื่อให้เรียกคำสั่ง `python` ได้จาก PowerShell

### ขั้นตอนที่ 1: ดาวน์โหลดโปรแกรม

วิธีง่ายที่สุดคือดาวน์โหลดจาก GitHub:

1. เปิดหน้า repo นี้ใน GitHub
2. กดปุ่ม **Code**
3. เลือก **Download ZIP**
4. แตกไฟล์ ZIP ไปไว้ในโฟลเดอร์ที่ต้องการ เช่น

```text
C:\MyRD\tmc-processor
```

ถ้าใช้ Git เป็นอยู่แล้ว สามารถใช้คำสั่งนี้แทนได้:

```powershell
git clone https://github.com/bokoboss/tmc-processor.git
cd tmc-processor
```

### ขั้นตอนที่ 2: เปิด PowerShell ในโฟลเดอร์โปรแกรม

เปิด PowerShell แล้วเข้าไปที่โฟลเดอร์โปรแกรม เช่น:

```powershell
cd "C:\MyRD\tmc-processor"
```

ถ้าแตก ZIP ไว้คนละที่ ให้เปลี่ยน path ให้ตรงกับเครื่องของคุณ

### ขั้นตอนที่ 3: ตรวจว่า Python ใช้งานได้

```powershell
python --version
```

ถ้าเห็นเวอร์ชัน เช่น `Python 3.11.x` หรือ `Python 3.12.x` แปลว่าใช้ได้

ถ้าขึ้นว่าไม่รู้จักคำสั่ง `python` ให้ติดตั้ง Python ใหม่ และอย่าลืมเลือก **Add python.exe to PATH**

### ขั้นตอนที่ 4: สร้างพื้นที่ติดตั้งของโปรแกรม

รันคำสั่งนี้ครั้งแรกครั้งเดียว:

```powershell
python -m venv .venv
```

จากนั้นเปิดใช้งานพื้นที่ติดตั้ง:

```powershell
.\.venv\Scripts\Activate.ps1
```

ถ้า PowerShell ไม่ยอมให้ activate และขึ้นข้อความเกี่ยวกับ execution policy ให้รันคำสั่งนี้ก่อน แล้วลอง activate อีกครั้ง:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

เมื่อ activate สำเร็จ จะเห็น `(.venv)` อยู่หน้าบรรทัดคำสั่ง

### ขั้นตอนที่ 5: ติดตั้ง package ที่โปรแกรมต้องใช้

```powershell
python -m pip install -U pip
python -m pip install -e ".[dev]"
```

ขั้นตอนนี้อาจใช้เวลาสักครู่

### ขั้นตอนที่ 6: ติดตั้ง Excel COM สำหรับโหมดรายงานหลัก

ถ้าต้องการใช้ **Excel Template Mode** ซึ่งเป็นโหมดที่แนะนำ ให้ติดตั้ง `pywin32` เพิ่ม:

```powershell
python -m pip install pywin32
```

จากนั้นทดสอบว่า Python เรียก Microsoft Excel ได้หรือไม่:

```powershell
python scripts\export_with_excel_com_smoke_test.py
```

ถ้าเห็นข้อความประมาณนี้ แปลว่าใช้งานได้:

```text
COM_AVAILABLE: Excel version 16.0
EXPORT_OK
```

ถ้าใช้ไม่ได้ โปรแกรมยังสามารถใช้ **Safe PNG Export Mode** ได้ แต่กราฟในรายงานจะเป็น PNG แบบคงที่แทน Native Chart ของ Excel

### ขั้นตอนที่ 7: เปิดโปรแกรม

```powershell
python -m streamlit run app.py
```

หลังรันคำสั่งนี้ โปรแกรมจะเปิดในเว็บเบราว์เซอร์ ถ้าไม่เปิดเอง ให้ดู URL ใน PowerShell แล้วเปิดตาม เช่น:

```text
http://localhost:8501
```

## วิธีเปิดโปรแกรมครั้งถัดไป

หลังจากติดตั้งครั้งแรกแล้ว ครั้งต่อไปทำแค่ 3 ขั้นตอนนี้:

```powershell
cd "C:\MyRD\tmc-processor"
.\.venv\Scripts\Activate.ps1
python -m streamlit run app.py
```

ถ้าโฟลเดอร์โปรแกรมอยู่คนละที่ ให้เปลี่ยน path ให้ตรงกับเครื่องของคุณ

## สรุปปัญหาที่พบบ่อยตอนติดตั้ง

### 1. PowerShell บอกว่า `python` ไม่รู้จัก

สาเหตุที่พบบ่อยคือยังไม่ได้ติดตั้ง Python หรือไม่ได้เลือก **Add python.exe to PATH** ตอนติดตั้ง

แนวทางแก้:

- ติดตั้ง Python ใหม่
- เลือก **Add python.exe to PATH**
- ปิด PowerShell แล้วเปิดใหม่
- ลอง `python --version` อีกครั้ง

### 2. Activate `.venv` ไม่ได้

ลองรัน:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

คำสั่งนี้มีผลเฉพาะหน้าต่าง PowerShell ที่เปิดอยู่เท่านั้น

### 3. ใช้ Excel Template Mode ไม่ได้

ตรวจว่าเครื่องมี Microsoft Excel แบบ desktop app และติดตั้ง `pywin32` แล้ว:

```powershell
python -m pip install pywin32
python scripts\export_with_excel_com_smoke_test.py
```

ถ้ายังใช้ไม่ได้ ให้ใช้ Safe PNG Export Mode ไปก่อน

## Excel Template Mode และ Excel COM

โหมดที่แนะนำสำหรับ final report คือ **Excel Template Mode** เพราะโปรแกรมจะใช้ Microsoft Excel เปิดสำเนาของเทมเพลต แล้วเติมข้อมูลลงใน cell/range ที่กำหนดไว้ เพื่อรักษา:

- Native Chart ใน Excel
- สูตรในเทมเพลต
- layout และ formatting ของ workbook

ข้อกำหนดสำหรับ Excel Template Mode:

- Windows
- Microsoft Excel แบบ desktop app
- `pywin32`

หาก Excel COM ใช้งานไม่ได้ โปรแกรมจะ fallback ไปใช้ **Safe PNG Export Mode** ซึ่งใช้ openpyxl และกราฟ PNG แบบคงที่แทน

## Project Session

โปรแกรมรองรับการ Save / Load Project Session ผ่านไฟล์ `.tmcproj.json`

Project Session ใช้เก็บการตั้งค่า เช่น:

- ข้อมูลรายงาน
- ป้ายปลายทางและชื่อถนน
- Mapping
- source stream / movement aggregation
- ช่วง Peak ที่ยืนยันแล้ว
- export settings

ไฟล์ session **ไม่เก็บ raw Excel input file** เพื่อลดความเสี่ยงเรื่องข้อมูลจริงและขนาดไฟล์ ผู้ใช้ต้องอัปโหลด raw file ใหม่เมื่อต้องการเปิด session กลับมาใช้งาน

## Many-to-one Movement Aggregation

โปรแกรมรองรับกรณีที่หลาย source stream ต้องรวมเป็น movement เดียวในรายงาน เช่น:

| Source stream | Raw movement | Output movement |
|---|---|---|
| mainline | ตรงทางหลัก | NS |
| frontage | ตรงทางคู่ขนาน | NS |

ผลลัพธ์ในตารางสรุปจะมี `NS` เพียงคอลัมน์เดียว โดยรวมค่าจาก source stream ทั้งหมด และสามารถตรวจสอบย้อนหลังได้ใน sheet `Movement_Aggregation_Audit`

ค่า `source_stream` ที่รองรับใน Mapping editor ได้แก่:

- `mainline`
- `frontage`
- `service_road`
- `ramp`
- `other`

## โครงสร้าง output หลัก

Workbook ที่ export จะมี sheet สำหรับตรวจสอบและใช้งานต่อ เช่น:

- `Setup`
- `Mapping`
- `Normalized_Data`
- `QC_Check`
- `Hourly_Movement_PCU`
- `Hourly_Vehicle_Class`
- `Vehicle_Composition_Report`
- `Vehicle_Group_PCE`
- `PHF_15min`
- `Peak_PHF`
- `Diagram_Data`
- `Movement_Aggregation_Audit`
- `Report_Text`
- `Charts`
- sheet รายงานตามเทมเพลต

ชื่อ sheet อาจเปลี่ยนได้ตามเวอร์ชันของโปรแกรมและการตั้งค่า export

## Validation และ QA

สร้างหรือปรับปรุง baseline expected results:

```powershell
python scripts\create_expected_results.py
```

ตรวจสอบผลกับ baseline:

```powershell
python scripts\validate_expected_results.py
```

Audit template:

```powershell
python scripts\audit_template.py
```

Smoke test สำหรับ Excel COM:

```powershell
python scripts\export_with_excel_com_smoke_test.py
```

## การทดสอบสำหรับผู้พัฒนา

```powershell
python -m pytest
```

## ข้อจำกัดปัจจุบัน

- โปรแกรมยังไม่ infer ทิศทางจราจรจาก geometry อัตโนมัติ ผู้ใช้ต้องตรวจและกำหนด Mapping เอง
- โหมด Excel Template Mode ใช้งานได้เฉพาะ Windows ที่มี Microsoft Excel และ pywin32
- Safe PNG Export Mode ใช้กราฟ PNG แบบคงที่ จึงไม่รักษา Native Chart ของ Excel template
- กรณีสะพาน/อุโมงค์/ทางยกระดับพิเศษยังควรจัดการด้วย tag/audit และ Mapping ก่อน ยังไม่ได้เป็น geometry engine เต็มรูปแบบ
- Project Session ไม่เก็บ raw Excel file ต้องอัปโหลดไฟล์สำรวจใหม่เมื่อนำ session กลับมาใช้

## ความเป็นส่วนตัวและข้อมูลจริง

Repository นี้เป็น public repository ดังนั้นไม่ควร commit ไฟล์ที่มีข้อมูลจริง เช่น:

- raw survey Excel files
- ไฟล์รายงาน/output ที่สร้างจากข้อมูลจริง
- Project Session ที่มีชื่อโครงการจริงหรือข้อมูลเฉพาะงาน
- mapping files ที่อ้างอิงโครงการจริง
- ไฟล์จากลูกค้า หรือไฟล์ที่มีข้อมูลที่ไม่ควรเผยแพร่

แนะนำให้เก็บข้อมูลจริงไว้ในเครื่องหรือ private storage เท่านั้น และใช้ synthetic/demo data สำหรับ public repository

## License

ยังไม่ได้กำหนด license อย่างเป็นทางการ หากต้องการเผยแพร่เป็น open source ควรเลือก license ให้ชัดเจนก่อนใช้งานหรือแจกจ่ายในวงกว้าง
