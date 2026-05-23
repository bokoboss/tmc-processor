# TMC Processor

TMC Processor เป็นโปรแกรมบน Streamlit สำหรับประมวลผลข้อมูล Turning Movement Count (TMC) จากไฟล์ Excel ให้เป็นตารางสรุป กราฟ ข้อมูล PCU/PCE ช่วงเร่งด่วน และ Excel Report ที่พร้อมนำไปใช้ต่อในงานรายงานจราจร

เวอร์ชันปัจจุบัน: `0.2.0` public beta

## โปรแกรมนี้ใช้ทำอะไร

โปรแกรมนี้ช่วยลดงานซ้ำในการจัดการไฟล์สำรวจ TMC โดยทำงานหลัก ๆ ดังนี้

- แปลงข้อมูล TMC จากไฟล์ Excel สำรวจจราจร
- กำหนดทิศทางและ Mapping จากข้อมูลดิบไปเป็น movement ที่ใช้ในรายงาน
- คำนวณ PCU ด้วยค่า PCE
- ตรวจและยืนยันช่วงเร่งด่วน AM/PM Peak ก่อนส่งออก
- ส่งออก Excel Report
- รองรับการทำหลายไฟล์แบบ Batch สำหรับจุดสำรวจเดียวกันหรือทางแยกเดียวกันหลายวัน

## เหมาะกับใคร

- traffic engineer ที่ต้องเตรียมรายงานปริมาณจราจร
- transport planner ที่ต้องตรวจข้อมูล TMC หลายวัน
- survey/review team ที่ต้องตรวจไฟล์สำรวจและช่วง Peak
- ผู้ใช้ที่ต้องเตรียม Excel Report จากข้อมูลสำรวจจราจร แต่ไม่อยากจัดตารางซ้ำด้วยมือทุกครั้ง

## ความสามารถหลัก

- Single-file workflow สำหรับประมวลผลไฟล์ TMC ทีละไฟล์
- Batch workflow สำหรับประมวลผลหลายวันของจุดสำรวจเดียวกัน
- Mapping Preset (`.mapping.json`) สำหรับใช้ Mapping เดิมซ้ำกับหลายไฟล์
- Project Session (`.tmcproj.json`) สำหรับบันทึกและโหลดการตั้งค่างาน
- Editable PCE factors สำหรับปรับค่า PCE ก่อนคำนวณ PCU
- Peak Review สำหรับตรวจและยืนยัน AM/PM Peak
- Excel Template Mode สำหรับส่งออกด้วย Microsoft Excel และ Excel COM
- Safe PNG Export Mode สำหรับใช้เป็นทางเลือกเมื่อ Excel COM ใช้งานไม่ได้
- Export Package ZIP สำหรับรวมผลลัพธ์ของงานไฟล์เดียว
- Batch Summary / Batch QC ใน `batch_summary.xlsx`
- Demo files สำหรับทดลองโดยไม่ต้องมีไฟล์สำรวจจริง
- Windows launcher ผ่าน `start_tmc_processor.bat`

## วิธีติดตั้งและเปิดใช้งานแบบง่ายที่สุดบน Windows

วิธีนี้เหมาะสำหรับผู้ใช้ทั่วไปที่ไม่คุ้นกับ Python หรือ GitHub มาก่อน

1. ดาวน์โหลด repository จาก GitHub เป็น ZIP หรือ clone ด้วย Git
2. ถ้าดาวน์โหลดเป็น ZIP ให้แตกไฟล์ไปไว้ในโฟลเดอร์ที่ต้องการ เช่น `C:\MyRD\tmc-processor`
3. ดับเบิลคลิกไฟล์ `start_tmc_processor.bat`
4. ครั้งแรกอาจใช้เวลาหลายนาที เพราะโปรแกรมจะสร้าง `.venv` และติดตั้ง package ที่จำเป็น
5. เมื่อพร้อมแล้ว โปรแกรมจะเปิดใน browser
6. ครั้งถัดไปจะเปิดเร็วขึ้น เพราะไม่ต้องติดตั้งใหม่ทั้งหมด

ถ้าเปิดจาก PowerShell หรือ Command Prompt สามารถใช้คำสั่งนี้ได้

```powershell
start_tmc_processor.bat
```

สิ่งที่ควรมีในเครื่อง:

- Windows
- Python 3.10 หรือใหม่กว่า และควรเลือก `Add python.exe to PATH` ตอนติดตั้ง Python
- Microsoft Excel desktop app เฉพาะกรณีที่ต้องการใช้ Excel Template Mode

ถ้า Excel COM ใช้งานไม่ได้ ยังสามารถใช้ Safe PNG Export Mode ได้

## Workflow แบบไฟล์เดียว

ใช้เมื่อต้องการประมวลผลไฟล์ TMC Excel หนึ่งไฟล์

1. เปิดโปรแกรมด้วย `start_tmc_processor.bat`
2. เลือกโหมดทำงานแบบไฟล์เดียว
3. Upload ไฟล์ TMC Excel หนึ่งไฟล์
4. กรอกข้อมูลงาน เช่น ชื่อโครงการ จุดสำรวจ วันที่สำรวจ ชื่อถนน และข้อมูลประกอบรายงาน
5. สร้างหรือโหลด Mapping
   - โหลด Mapping Preset (`.mapping.json`)
   - หรือโหลดไฟล์ Mapping Excel ที่เคยบันทึกไว้
   - หรือแก้ไข Mapping ในตารางของโปรแกรม
6. ตรวจหรือปรับค่า PCE factors ถ้าจำเป็น
7. ประมวลผลไฟล์
8. ตรวจ Dashboard และยืนยัน AM/PM Peak
9. สร้าง Excel Report หรือ Export Package ZIP

Export Package ZIP จะรวมไฟล์ผลลัพธ์ที่ประมวลผลแล้ว เช่น report, chart, summary, Mapping และ Project Session แต่โดยค่าเริ่มต้นจะไม่รวม raw input Excel file

## Workflow แบบ Batch

Batch workflow เหมาะสำหรับกรณีที่มีข้อมูลหลายวันของจุดสำรวจเดียวกันหรือทางแยกเดียวกัน และใช้ Mapping Preset เดียวกัน

1. เปิดโปรแกรมด้วย `start_tmc_processor.bat`
2. เลือก `ประมวลผลหลายไฟล์`
3. Upload ไฟล์ TMC Excel หลายไฟล์ เช่น ไฟล์ของแต่ละวัน
4. โหลด Mapping Preset หนึ่งไฟล์ เช่น `samples/demo/DEMO_TMC1_FourLeg.mapping.json`
5. ตั้งค่า survey date และ output stem ของแต่ละไฟล์
6. กดวิเคราะห์ Batch
7. ตรวจและยืนยัน Peak ของแต่ละไฟล์
8. สร้าง Batch ZIP

Batch ZIP จะมี `batch_summary.xlsx` ซึ่งรวม `Batch_QC` และโฟลเดอร์ผลลัพธ์ของแต่ละไฟล์ที่ประมวลผลสำเร็จ โดยค่าเริ่มต้นจะไม่รวม raw input Excel file และไม่ควรมี local raw file paths อยู่ใน package

ข้อจำกัดสำคัญของ Batch v1:

- เหมาะกับจุดสำรวจเดียวกันหรือทางแยกเดียวกันหลายวัน
- ใช้ Mapping Preset ร่วมกันหนึ่งไฟล์
- ยังไม่มีการเลือก Mapping แยกเป็นรายไฟล์

## ทดลองใช้ด้วย Demo files

ไฟล์ตัวอย่างอยู่ใน `samples/demo/`

- `DEMO_TMC1_FourLeg.xlsx`
- `DEMO_TMC1_FourLeg_Day2.xlsx`
- `DEMO_TMC1_FourLeg.mapping.json`
- `DEMO_TMC1_FourLeg_mapping.xlsx`
- `DEMO_TMC1_FourLeg_session.tmcproj.json`

ไฟล์ทั้งหมดใน `samples/demo/` เป็นข้อมูลสังเคราะห์ ไม่มีข้อมูลสำรวจจริง ไม่มีชื่อโครงการจริง ไม่มีข้อมูลลูกค้า และไม่มีข้อมูลส่วนตัว

### ทดลองแบบไฟล์เดียว

1. เปิดโปรแกรม
2. Upload `samples/demo/DEMO_TMC1_FourLeg.xlsx`
3. โหลด `samples/demo/DEMO_TMC1_FourLeg.mapping.json` หรือ `samples/demo/DEMO_TMC1_FourLeg_mapping.xlsx`
4. ประมวลผลไฟล์
5. ตรวจและยืนยัน Peak
6. สร้าง Excel Report หรือ Export Package ZIP

### ทดลองแบบ Batch

1. เลือก `ประมวลผลหลายไฟล์`
2. Upload ไฟล์ต่อไปนี้
   - `samples/demo/DEMO_TMC1_FourLeg.xlsx`
   - `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
3. โหลด `samples/demo/DEMO_TMC1_FourLeg.mapping.json`
4. ตั้งค่า survey date และ output stem ถ้าต้องการ
5. วิเคราะห์ Batch
6. ตรวจ Peak ของแต่ละไฟล์
7. สร้าง Batch ZIP

## Mapping Preset และ Project Session ต่างกันอย่างไร

Mapping Preset (`.mapping.json`) เก็บเฉพาะข้อมูล Mapping เช่น raw sheet, source stream, movement label, output movement code, include flags และ aggregation fields เหมาะสำหรับใช้ Mapping เดิมซ้ำกับหลายวันหรือหลายไฟล์ของทางแยกเดียวกัน

Project Session (`.tmcproj.json`) เก็บการตั้งค่างานที่กว้างกว่า เช่น metadata, Mapping, PCE factors, peak settings และ export settings ใช้สำหรับกลับมาเปิดงานเดิมต่อภายหลัง

Project Session ไม่ได้ฝัง raw input Excel file ไว้ในไฟล์ ผู้ใช้ต้อง Upload ไฟล์ Excel ต้นทางใหม่เมื่อเปิดงานกลับมาใช้อีกครั้ง

## Excel Template Mode และ Safe PNG Export Mode

Excel Template Mode ใช้ Microsoft Excel desktop app ผ่าน Excel COM บน Windows เพื่อเติมข้อมูลลงใน template และช่วยรักษา native charts, formulas, layout และ formatting ของไฟล์ Excel

Safe PNG Export Mode ใช้ openpyxl และ chart image แบบ PNG เป็นทางเลือกเมื่อ Excel COM ใช้งานไม่ได้ เหมาะสำหรับเครื่องที่ไม่มี Microsoft Excel หรือใช้งาน Excel COM ไม่สำเร็จ

## ความเป็นส่วนตัวและความปลอดภัยของข้อมูล

repository นี้เป็น public repository จึงไม่ควร commit หรืออัปโหลดข้อมูลจริงขึ้น GitHub

ห้าม commit ไฟล์เหล่านี้ถ้าเป็นข้อมูลงานจริง:

- raw survey Excel files
- ไฟล์รายงานหรือ output ที่สร้างจากข้อมูลลูกค้าหรือโครงการจริง
- Project Session ที่มีชื่อโครงการจริงหรือข้อมูลเฉพาะงาน
- Mapping files ของโครงการจริง
- ไฟล์จากลูกค้า หรือไฟล์ที่มีข้อมูลส่วนตัว

พื้นที่เหล่านี้ถูกตั้งใจให้เก็บไฟล์งานจริงในเครื่องและถูก ignore ไว้:

- `samples/raw/`
- `outputs/`
- generated `.tmcproj.json` files นอก `samples/demo/`
- generated ZIP files

โดยค่าเริ่มต้น Export Package ZIP จะไม่รวม raw input Excel files แต่ผู้ใช้ยังควรตรวจ package ก่อนส่งต่อทุกครั้ง

## การตรวจสอบสำหรับผู้พัฒนา

คำสั่งตรวจสอบพื้นฐานก่อน release:

```powershell
python -m py_compile app.py
python scripts/smoke_demo.py
python -m pytest
```

`scripts/smoke_demo.py` ใช้ไฟล์สังเคราะห์จาก `samples/demo/`

## License

โครงการนี้เผยแพร่ภายใต้ MIT License ดูรายละเอียดได้ที่ [LICENSE](LICENSE)
