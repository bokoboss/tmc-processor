# TMC Processor

TMC Processor เป็นโปรแกรมบน Streamlit สำหรับประมวลผลข้อมูล Turning Movement Count (TMC) จากไฟล์ Excel ให้เป็นตารางสรุป ข้อมูล PCU/PCE ช่วงเร่งด่วน และ Excel Report ที่พร้อมนำไปใช้ต่อในงานรายงานจราจร

รุ่นปัจจุบัน: **`v1.0.0` stable**

โค้ดรุ่นปัจจุบันอยู่บน branch `main`

## Legacy version

รุ่น Public Beta เดิมที่เผยแพร่เมื่อ 24 พฤษภาคม 2026 ถูกเก็บไว้เป็น historical snapshot และจะไม่ถูกแก้ไขตามการพัฒนาของ `main`

- Source snapshot: [`legacy/v0.2.0`](https://github.com/bokoboss/tmc-processor/tree/legacy/v0.2.0)
- Original tag: [`v0.2.0`](https://github.com/bokoboss/tmc-processor/tree/v0.2.0)
- Original release: [TMC Processor v0.2.0 Public Beta](https://github.com/bokoboss/tmc-processor/releases/tag/v0.2.0)

หากต้องการดูหรือทดลองรุ่นที่เผยแพร่เดิม ให้ใช้ `v0.2.0` หรือ branch `legacy/v0.2.0` แทน `main`

## โปรแกรมนี้ใช้ทำอะไร

โปรแกรมนี้ช่วยลดงานซ้ำในการจัดการไฟล์สำรวจ TMC โดยทำงานหลัก ๆ ดังนี้

- แปลงข้อมูล TMC จากไฟล์ Excel สำรวจจราจร
- กำหนดทิศทางและ Mapping จากข้อมูลดิบไปเป็น movement ที่ใช้ในรายงาน
- คำนวณ PCU ด้วยค่า PCE
- วิเคราะห์และตรวจ AM/PM Peak
- ให้ผู้ใช้ยืนยัน Peak อย่างชัดเจนก่อนส่งออก
- ส่งออก Excel Report
- รองรับการทำหลายไฟล์แบบ Batch สำหรับจุดสำรวจเดียวกันหรือทางแยกเดียวกันหลายวัน

Canonical workflow ของโปรแกรมคือ:

`Data → Mapping → Analyze → Review → Export`

## เหมาะกับใคร

- traffic engineer ที่ต้องเตรียมรายงานปริมาณจราจร
- transport planner ที่ต้องตรวจข้อมูล TMC หลายวัน
- survey/review team ที่ต้องตรวจไฟล์สำรวจและช่วง Peak
- ผู้ใช้ที่ต้องเตรียม Excel Report จากข้อมูลสำรวจจราจร แต่ไม่อยากจัดตารางซ้ำด้วยมือทุกครั้ง

## ความสามารถหลัก

- Single-file workflow สำหรับประมวลผลไฟล์ TMC ทีละไฟล์
- Batch workflow สำหรับประมวลผลหลายวันของจุดสำรวจเดียวกัน
- Mapping Preset (`.mapping.json`) สำหรับใช้ Mapping เดิมซ้ำกับหลายไฟล์
- Project Session (`.tmcproj.json`) สำหรับบันทึกและโหลดการตั้งค่างานใน Single workflow
- Editable PCE factors สำหรับปรับค่า PCE ก่อนคำนวณ PCU
- Peak Review ที่แยก suggested / draft / confirmed Peak และต้องยืนยันก่อน Export
- Standard Excel report แบบ native-template preserving OOXML เมื่อ template รองรับ Peak ที่ยืนยัน
- Safe PNG Export Mode เป็น explicit alternative และ fallback
- Excel COM คงไว้เป็น optional legacy/diagnostic capability ไม่ใช่ข้อกำหนดของ Standard report
- Export Package ZIP สำหรับรวมผลลัพธ์ของงานไฟล์เดียว
- Batch Summary / Batch QC ใน `batch_summary.xlsx`
- Demo files สำหรับทดลองโดยไม่ต้องมีไฟล์สำรวจจริง
- Windows launcher ผ่าน `start_tmc_processor.bat`

## คำแนะนำการใช้งาน

แนะนำให้ใช้ Single workflow เป็น workflow หลักสำหรับการจัดทำรายงาน เนื่องจากตรวจสอบ Mapping, Peak, QC และรายงานได้ละเอียดที่สุด

Batch workflow เหมาะสำหรับกรณีที่ต้องประมวลผลหลายไฟล์หรือหลายวันของจุดสำรวจเดียวกันโดยใช้ Mapping Preset เดียวกัน และควรใช้เมื่อได้ตรวจสอบ Mapping จาก Single workflow แล้ว

Standard report จะใช้ native-template preserving OOXML เมื่อ Peak ที่ยืนยันสามารถแทนได้โดย template ปัจจุบัน หาก Peak ที่ถูกต้องไม่สามารถแทนใน template ได้ โปรแกรมจะไม่ปัดหรือเปลี่ยน Peak แต่จะ fallback ไป Safe PNG แทน

ข้อจำกัดที่ติดตามอยู่ใน [Issue #23](https://github.com/bokoboss/tmc-processor/issues/23) คือ native-template support สำหรับ rolling 60-minute Peak ที่ไม่ตรงชั่วโมง เช่น `08:15–09:15`

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

Standard report ไม่จำเป็นต้องใช้ Microsoft Excel desktop app หรือ Excel COM ในกรณีที่ native OOXML path รองรับงานนั้นอยู่แล้ว ส่วน Excel COM เป็น optional capability สำหรับ legacy/diagnostic path

## Workflow แบบไฟล์เดียว

ใช้เมื่อต้องการประมวลผลไฟล์ TMC Excel หนึ่งไฟล์

1. เปิดโปรแกรมด้วย `start_tmc_processor.bat`
2. เลือก Single workflow
3. Upload ไฟล์ TMC Excel หนึ่งไฟล์ใน Data
4. กรอกข้อมูลงานและช่วงเวลาสำรวจ
5. สร้างหรือโหลด Mapping
   - โหลด Mapping Preset (`.mapping.json`)
   - หรือโหลดไฟล์ Mapping Excel ที่เคยบันทึกไว้
   - หรือแก้ไข Mapping ในตารางของโปรแกรม
6. ตั้งค่า Peak search / PCE ตามต้องการและกด Analyze
7. ตรวจ QC และ Suggested Peak ใน Review
8. ยืนยัน AM/PM Peak อย่างชัดเจน
9. สร้าง Standard report หรือเลือก Safe PNG ใน Export

Export Package ZIP จะรวมไฟล์ผลลัพธ์ที่ประมวลผลแล้วตาม workflow โดยค่าเริ่มต้นจะไม่รวม raw input Excel file

## Workflow แบบ Batch

Batch workflow เหมาะสำหรับกรณีที่มีข้อมูลหลายวันของจุดสำรวจเดียวกันหรือทางแยกเดียวกัน และใช้ Mapping Preset เดียวกัน

1. เปิดโปรแกรมด้วย `start_tmc_processor.bat`
2. เลือก Batch workflow
3. Upload ไฟล์ TMC Excel หลายไฟล์ใน Data
4. โหลด Mapping Preset หนึ่งไฟล์ เช่น `samples/demo/DEMO_TMC1_FourLeg.mapping.json`
5. ตั้งค่า survey date และ output stem ของแต่ละไฟล์ถ้าจำเป็น
6. กด Analyze
7. ตรวจ QC และยืนยัน Peak ของแต่ละไฟล์ใน Review
8. สร้าง Batch ZIP ใน Export

Batch ZIP จะมี `batch_summary.xlsx` ซึ่งรวม `Batch_QC` และผลลัพธ์รายไฟล์ที่ประมวลผลสำเร็จ โดยค่าเริ่มต้นจะไม่รวม raw input Excel file และไม่ควรมี local raw file paths อยู่ใน package

ข้อจำกัดสำคัญของ Batch:

- เหมาะกับจุดสำรวจเดียวกันหรือทางแยกเดียวกันหลายวัน
- ใช้ Mapping Preset ร่วมกันหนึ่งไฟล์
- ยังไม่มีการเลือก Mapping แยกเป็นรายไฟล์
- Project Session ปัจจุบันรองรับ Single workflow เท่านั้น จึงปิด Save/Open controls ใน Batch เพื่อไม่ให้เกิด partial restore ที่ทำให้เข้าใจผิด

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
4. Analyze
5. ตรวจและยืนยัน Peak
6. สร้าง Standard report หรือ Safe PNG report

### ทดลองแบบ Batch

1. เลือก Batch workflow
2. Upload ไฟล์ต่อไปนี้
   - `samples/demo/DEMO_TMC1_FourLeg.xlsx`
   - `samples/demo/DEMO_TMC1_FourLeg_Day2.xlsx`
3. โหลด `samples/demo/DEMO_TMC1_FourLeg.mapping.json`
4. ตั้งค่า survey date และ output stem ถ้าต้องการ
5. Analyze
6. ตรวจและยืนยัน Peak ของแต่ละไฟล์
7. สร้าง Batch ZIP

## Mapping Preset และ Project Session ต่างกันอย่างไร

Mapping Preset (`.mapping.json`) เก็บข้อมูล Mapping เช่น raw sheet, source stream, movement label, output movement code, include flags และ aggregation fields เหมาะสำหรับใช้ Mapping เดิมซ้ำกับหลายวันหรือหลายไฟล์ของทางแยกเดียวกัน

Project Session (`.tmcproj.json`) เก็บการตั้งค่างานที่กว้างกว่า เช่น metadata, Mapping, PCE factors, Peak settings และ export settings เพื่อกลับมาเปิด Single workflow เดิมต่อภายหลัง

Project Session ไม่ได้ฝัง raw input Excel file ไว้ในไฟล์ ผู้ใช้ต้อง Upload ไฟล์ Excel ต้นทางใหม่เมื่อเปิดงานกลับมาใช้อีกครั้ง

Batch Project Session ยังไม่รองรับการ round-trip Batch uploads และ Batch Mapping Preset จึงไม่เปิด Save/Open ใน Batch รุ่นปัจจุบัน

## Standard report และ Safe PNG

Standard report ใช้ package-preserving OOXML กับ authoritative Excel template เมื่อ confirmed/effective Peak สามารถแทนใน template ได้ โดยรักษา native charts, drawings, formulas, styles และ relationships ของ workbook

Safe PNG ใช้ generated workbook และ chart image เป็น explicit alternative/fallback และยังคง Peak ที่ผู้ใช้ยืนยันไว้โดยไม่ round หรือ snap เพื่อให้เข้ากับ template

ถ้าเป็น rolling 60-minute Peak ที่ valid แต่ไม่ตรงแถวเวลาที่ native template ปัจจุบันรองรับ เช่น `08:15–09:15` Standard จะ fallback ไป Safe PNG จนกว่า Issue #23 จะได้รับการแก้ไข

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
python -m pytest
python -m compileall -q app.py src
git diff --check
```

`scripts/smoke_demo.py` ใช้ไฟล์สังเคราะห์จาก `samples/demo/` สำหรับ smoke ที่เกี่ยวข้อง

## CI and testing

GitHub Actions runs the automated suite on Windows for Python 3.10 and 3.12 on pushes to `main` and pull requests targeting `main`. CI uses committed synthetic/demo fixtures and does not require Microsoft Excel or files under `samples/raw/`.

Real-workbook and Excel/native-template qualification เป็น local/manual release gate สำหรับการเปลี่ยนแปลงที่กระทบ workflow, Mapping, Peak, state หรือ export โดยไฟล์จริงใน `samples/raw/` ไม่อยู่ใน public Git history

## Maintenance status

หลัง `v1.0.0` โครงการเข้าสู่ maintenance mode: แก้บั๊กหรือทำ enhancement แบบมี issue/acceptance แยกเป็นงาน ๆ โดยไม่เปลี่ยน validated behavior โดยไม่จำเป็น

Known enhancement backlog ปัจจุบัน: [Issue #23 — native Excel template support for non-hour-aligned rolling Peaks](https://github.com/bokoboss/tmc-processor/issues/23)

## License

โครงการนี้เผยแพร่ภายใต้ MIT License ดูรายละเอียดได้ที่ [LICENSE](LICENSE)
