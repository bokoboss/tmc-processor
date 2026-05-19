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
- Basic Batch Processing v1 สำหรับประมวลผลไฟล์สำรวจหลายวันของจุดเดียวกันด้วย Mapping Preset และค่า PCE ร่วมกัน
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

## ทดลองใช้งานด้วยไฟล์ Demo

ไฟล์ใน `samples/demo/` เป็นข้อมูลสังเคราะห์ทั้งหมด ใช้สำหรับทดลองโปรแกรมโดยไม่ต้องมีไฟล์สำรวจจริง

1. Run `start_tmc_processor.bat`
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`
3. Load `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`
4. Process
5. Review Peak
6. Generate Excel Report

Batch demo:

1. Open the `ประมวลผลหลายไฟล์` tab.
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx` and `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`.
3. Load `samples/demo/DEMO_TMC1_FourLeg.mapping.json`.
4. Click Process Batch.
5. Download the Batch ZIP.

## วิธีเปิดแบบง่ายที่สุดสำหรับ Windows

ส่วนนี้เขียนสำหรับผู้ใช้ทั่วไปที่ยังไม่คุ้นกับ Python หรือ command line มาก่อน

### สิ่งที่ต้องมี

1. **Windows**
2. **Python** แนะนำเวอร์ชัน 3.11 หรือ 3.12
3. **Microsoft Excel** ถ้าต้องการใช้โหมดรายงานตามเทมเพลต Excel และ Native Chart

> ตอนติดตั้ง Python ให้เลือกตัวเลือก **Add python.exe to PATH** ด้วย

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

### ขั้นตอนที่ 2: เปิดโปรแกรม

ดับเบิลคลิก `start_tmc_processor.bat`

ครั้งแรกอาจใช้เวลาหลายนาที เพราะไฟล์นี้จะสร้าง `.venv` ติดตั้งแพ็กเกจที่จำเป็น ติดตั้ง `pywin32` และทดสอบ Excel COM แบบไม่บังคับให้โดยอัตโนมัติ ครั้งต่อไปจะเปิดเร็วขึ้น เพราะจะข้ามขั้นตอนติดตั้งเมื่อพบ `.venv` แล้ว

เมื่อโปรแกรมเปิดแล้ว ให้ใช้งานผ่านหน้าต่างเว็บเบราว์เซอร์ที่เปิดขึ้นมาอัตโนมัติ หากต้องการปิดโปรแกรม ให้กลับมาที่หน้าต่างสีดำแล้วกด `Ctrl+C`

ถ้าต้องการแยกขั้นตอนติดตั้งและเปิดโปรแกรมเอง ยังสามารถใช้ `setup_windows.bat` เพื่อติดตั้งก่อน แล้วใช้ `run_app.bat` เพื่อเปิดโปรแกรมในครั้งต่อไปได้

อย่านำไฟล์ข้อมูลดิบของโครงการจริง ไฟล์ผลลัพธ์ Excel ไฟล์ `.tmcproj.json` หรือข้อมูลส่วนตัวขึ้น GitHub ให้เก็บไฟล์งานจริงไว้ในเครื่องของผู้ใช้เท่านั้น

## วิธีเปิดโปรแกรมครั้งถัดไป

หลังจากติดตั้งครั้งแรกแล้ว ครั้งต่อไปให้ดับเบิลคลิก `start_tmc_processor.bat` ได้เลย โปรแกรมจะข้ามขั้นตอนติดตั้งและเปิดเร็วขึ้น

ถ้าใช้วิธีแยกขั้นตอน สามารถดับเบิลคลิก `run_app.bat` หลังจากเคยรัน `setup_windows.bat` แล้ว

## สรุปปัญหาที่พบบ่อยตอนติดตั้ง

### 1. PowerShell บอกว่า `python` ไม่รู้จัก

สาเหตุที่พบบ่อยคือยังไม่ได้ติดตั้ง Python หรือไม่ได้เลือก **Add python.exe to PATH** ตอนติดตั้ง

แนวทางแก้:

- ติดตั้ง Python ใหม่
- เลือก **Add python.exe to PATH**
- ปิด PowerShell แล้วเปิดใหม่
- ลอง `python --version` อีกครั้ง

### 2. ใช้ Excel Template Mode ไม่ได้

ตรวจว่าเครื่องมี Microsoft Excel แบบ desktop app แล้วเปิด `start_tmc_processor.bat` อีกครั้ง โปรแกรมจะติดตั้ง `pywin32` และทดสอบ Excel COM ให้แบบไม่บังคับในครั้งแรก

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

## Basic Batch Processing v1

Basic Batch v1 is intended for the same survey point or intersection surveyed
across multiple days. Upload multiple raw TMC Excel workbooks, use one shared
Mapping Preset and shared PCE factors, then download one Batch ZIP.

The batch ZIP contains `batch_summary.xlsx` plus one sanitized folder per
successful input file. Each folder contains `report.xlsx`, `export_summary.txt`,
`session.tmcproj.json`, `mapping_preset.mapping.json`, and chart PNGs. Raw input
Excel files and local file paths are not included. If one workbook fails,
processing continues and the failure is recorded in `batch_summary.xlsx`.

## Mapping Preset vs Project Session

- Mapping Preset (`.mapping.json`) stores only the mapping table: raw sheet,
  source stream, movement label, output movement code, include flags, and
  aggregation fields. Use it to reuse the same intersection mapping across
  survey dates, raw workbooks, or team members.
- Project Session (`.tmcproj.json`) stores the broader job setup, including
  metadata, mapping, PCE factors, peak settings, and export settings. Use it to
  reopen a full job configuration.

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

## การตรวจสอบสำหรับผู้พัฒนา

ชุดทดสอบและสคริปต์ QA ภายในไม่ได้รวมอยู่ใน public release นี้ หากเพิ่มชุดทดสอบใน branch สำหรับพัฒนา สามารถติดตั้ง dependency เสริมได้ด้วย:

```powershell
python -m pip install -e ".[dev]"
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

This project is released under the MIT License. See LICENSE for details.
