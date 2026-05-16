# Auto Export System
<img width="539" height="339" alt="image" src="https://github.com/user-attachments/assets/dd89b373-64fc-4a72-b366-203b7b7830e0" />

Python application สำหรับ export ข้อมูลจาก MySQL Database อัตโนมัติแบบรายชั่วโมง พร้อมระบบ Manual Export, Auto Schedule, System Tray, Log Management และ Excel Report Generation

รองรับการทำงานแบบ background service และสามารถ build เป็น `.exe` สำหรับใช้งานจริงใน production environment ได้

---

# Features

- Auto export ข้อมูลทุกต้นชั่วโมง
- Manual export แบบกำหนดช่วงเวลาเอง
- Export เป็น Excel (`.xlsx`)
- Multi-sheet Excel report
- Auto reset hourly files
- Daily backup system
- System tray support
- Desktop notification
- Real-time GUI log viewer
- Threaded background worker
- Config file support
- Daily log rotation
- Cancel manual export ได้ระหว่างทำงาน
- รองรับ build เป็น Windows EXE

---

# Application Overview

โปรแกรมนี้ถูกออกแบบมาเพื่อ export ข้อมูล sorting system จากฐานข้อมูล MySQL ออกเป็นไฟล์ Excel รายชั่วโมง โดยแบ่งข้อมูลออกเป็น:

- Throughput Report
- Abnormal Report

เหมาะสำหรับ:

- Warehouse Automation
- Logistics Hub
- Parcel Sorting Center
- WCS Monitoring
- Sorting Performance Report

---

# Tech Stack

- Python
- Tkinter
- Pandas
- PyMySQL
- OpenPyXL
- PyStray
- Plyer
- tkcalendar

---

# Project Structure

```text
project/
│
├── AutoExport.py
├── config.ini
│
├── logs/
│   └── YYYY-MM-DD.log
│
└── export/
    ├── 00.xlsx
    ├── 01.xlsx
    ├── 02.xlsx
    └── YYYY-MM-DD/
        ├── 00.xlsx
        ├── 01.xlsx
        └── ...
```

---

# Export Reports

โปรแกรมจะสร้าง Excel จำนวน 2 Sheet

## Sheet 1 — Throughput

ใช้สำหรับสรุป performance ของ feeder

Columns:

| Column | Description |
|---|---|
| Pipeline | Pipeline Name |
| Feeder No | Feeder Number |
| Loading Num | Total Loading |
| Loading Valid Num | Valid Loading |
| Loading Valid Rate | Success Rate |
| Loading Efficiency(per hour) | Hourly Efficiency |
| Duration of Feeder Online | Online Duration |
| Remark | Additional Note |

---

## Sheet 2 — Abnormal

ใช้สำหรับเก็บข้อมูล package ผิดปกติ

Columns:

| Column | Description |
|---|---|
| Billcode | Tracking Number |
| Pipeline | Pipeline |
| Feeder No | Feeder Number |
| Tray Code | Tray Code |
| Sort Time | Sorting Time |
| Sorting Port | Sorting Port |
| Weight | Weight |
| Expect Chute Code | Expected Chute |
| Sort Source | Error Source |

---

# Database Tables

โปรแกรมใช้ข้อมูลจาก table:

```sql
alreadysortinfo
```

---

# Auto Export Logic

ระบบจะทำงานอัตโนมัติทุกต้นชั่วโมง

ตัวอย่าง:

```text
13:00 → Export ข้อมูล 12:00 - 12:59
14:00 → Export ข้อมูล 13:00 - 13:59
15:00 → Export ข้อมูล 14:00 - 14:59
```

---

# Manual Export

รองรับ manual export แบบกำหนด:

- Start Date
- Start Hour
- End Date
- End Hour

ระบบจะ export แบบ hour-by-hour อัตโนมัติ

ตัวอย่าง:

```text
2026-05-01 08:00
→
2026-05-01 18:00
```

จะสร้างไฟล์:

```text
09.xlsx
10.xlsx
11.xlsx
...
18.xlsx
```

---

# Export Limitation

ระบบจำกัดย้อนหลังสูงสุด:

```python
MAX_DAYS_BACK = 7
```

เพื่อป้องกัน query database ปริมาณมากเกินไป

---

# Hourly File Reset System

ทุกวันเวลา:

```text
12:00 PM
```

ระบบจะ:

1. Backup ไฟล์เดิมเข้า folder วันที่
2. Clear hourly files
3. สร้างไฟล์ใหม่ 00-23.xlsx

ตัวอย่าง:

```text
export/
├── 2026-05-17/
│   ├── 00.xlsx
│   ├── 01.xlsx
│   └── ...
│
├── 00.xlsx
├── 01.xlsx
└── ...
```

---

# Notification System

โปรแกรมรองรับ desktop notification

ตัวอย่าง:

```text
Export Success
Export Failed
Reset Complete
```

---

# Logging System

ระบบจะสร้าง log แยกตามวันอัตโนมัติ

```text
logs/YYYY-MM-DD.log
```

ตัวอย่าง:

```text
2026-05-17 10:00:00 - INFO - Export OK
2026-05-17 11:00:00 - INFO - Reset complete
2026-05-17 12:00:00 - ERROR - Database timeout
```

---

# System Tray

โปรแกรมรองรับการทำงานแบบ background

Tray Menu:

```text
Show
Exit
```

เมื่อปิดหน้าต่าง โปรแกรมจะ minimize เข้า tray อัตโนมัติ

---

# Threading Design

โปรแกรมใช้:

```python
threading.Thread()
```

เพื่อแยก:

- GUI Thread
- Auto Export Worker
- Tray Worker
- Manual Export Worker

ข้อดี:

- GUI ไม่ค้าง
- Export พร้อมกันได้
- รองรับ long-running process

---

# Configuration

## config.ini

```ini
[DEFAULT]
output_path=C:/EXPORT
host=127.0.0.1
user=wcs
password=wcs
database=db_wcs
port=3306
```

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/yourname/auto-export-system.git
```

---

## 2. Install Dependencies

```bash
pip install pandas pymysql openpyxl tkcalendar pystray pillow plyer
```

---

## 3. Run Application

```bash
python AutoExport.py
```

---

# Build EXE

ใช้ PyInstaller:

```bash
pyinstaller --onefile --windowed AutoExport.py
```

หรือ:

```bash
pyinstaller --onefile --noconsole AutoExport.py
```

---

# GUI Features

- Output folder browser
- Manual export range selector
- Start / Stop auto export
- Real-time log viewer
- Status monitoring
- Tray minimize
- Cancel running export

---

# Error Handling

ระบบรองรับ:

- Database connection error
- Invalid time range
- File overwrite handling
- Cancel running task
- Log rotation
- Export failure notification

---

# Windows EXE Support

โปรแกรมรองรับ:

- Portable EXE
- Background startup
- Tray execution
- Standalone deployment

มีการ fix path สำหรับ PyInstaller:

```python
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
```

---

# Future Improvements

- Multi database profile
- Email report system
- CSV export
- FTP upload
- Auto archive compression
- Dashboard monitoring
- Web interface
- SQLite local cache

---

# License

MIT License

---

# Author

Developed for warehouse sorting automation and hourly export workflow optimization.
