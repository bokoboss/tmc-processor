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

## การติดตั้ง

แนะนำให้ใช้ virtual environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -U pip
python -m pip install -e ".[dev]"
```

โปรเจกต์กำหนด `requires-python = ">=3.10"` และใช้ dependencies หลัก ได้แก่ Streamlit, pandas, openpyxl และ matplotlib

## การรันโปรแกรม

```powershell
python -m streamlit run app.py
```

หรือถ้า `streamlit` อยู่ใน PATH แล้ว:

```powershell
streamlit run app.py
```

## Excel Template Mode และ Excel COM

โหมดที่แนะนำสำหรับ final report คือ **Excel Template Mode** เพราะโปรแกรมจะใช้ Microsoft Excel เปิดสำเนาของเทมเพลต แล้วเติมข้อมูลลงใน cell/range ที่กำหนดไว้ เพื่อรักษา:

- Native Chart ใน Excel
- สูตรในเทมเพลต
- layout และ formatting ของ workbook

ข้อกำหนดสำหรับ Excel Template Mode:

- Windows
- Microsoft Excel แบบ desktop app
- `pywin32`

ติดตั้ง `pywin32` ได้ด้วย:

```powershell
python -m pip install pywin32
```

ทดสอบ Excel COM:

```powershell
python scripts\export_with_excel_com_smoke_test.py
```

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

## การทดสอบ

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
