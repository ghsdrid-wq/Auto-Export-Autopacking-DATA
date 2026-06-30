# STRUCTURE — Auto-Export-Autopacking-DATA

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เปลี่ยน SQL/คอลัมน์, เปลี่ยน config key, เปลี่ยน flow scheduler/backup/reset, หรือเพิ่มไฟล์ — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
แอป GUI (**Tkinter**) ชื่อ *"Auto Export"* — ดึงข้อมูล **Autopacking/Sorting (WCS)** จาก MySQL ออกเป็นไฟล์ Excel **รายชั่วโมง** อัตโนมัติ (และ Manual range ย้อนหลังได้สูงสุด 7 วัน) รันอยู่ใน system tray

## วิธีรัน / Entry point
- รัน: `python AutoExport.py` → คลาส `App` (Tkinter) เริ่มแบบ tray, auto-loop ใน thread
- รองรับ frozen exe (`os.chdir` ไป dir ของ exe)

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `AutoExport.py` | ทั้งโปรแกรม — UI, config, logger รายวัน, tray, `export_data()`, `reset_hourly_files()`, `auto_loop()` |

## Logic สำคัญ
- `export_data(config, start, end, output)` — query 2 ชุดจากตาราง `alreadysortinfo`:
  - `df1` (sheet **Throughput**): group ตาม `pipeline, inductNo` นับ billCode
  - `df2` (sheet **Abnormal**): รายการ `sortSource IN ('拒绝件','无分拣计划','最大循环件')`
  - เซฟไฟล์ชื่อ `<ชั่วโมง>.xlsx` (เช่น `13.xlsx`)
- `auto_loop()` — ทุกต้นชั่วโมง (minute==0) export ชั่วโมงที่ผ่านมา; เวลา 12:00 ทำ `reset_hourly_files()`
- `reset_hourly_files()` — ย้ายไฟล์ 00–23.xlsx ของวันก่อนเข้าโฟลเดอร์ backup (วันที่) แล้วสร้างไฟล์เปล่าใหม่
- `manual_export_range()` — export ทีละชั่วโมงตามช่วงที่เลือก (จำกัด `MAX_DAYS_BACK = 7`)

## Config (`config.ini`, `[DEFAULT]`)
- `output_path`, `host`, `user`, `password`, `database` (default `db_wcs`), `port` (3306)

## Dependencies
- `pymysql`, `pandas`, `openpyxl`, `tkcalendar`, `pystray`, `Pillow`, `plyer` (notification), `tkinter`

## ข้อควรระวัง
- SQL hardcode ตาราง/คอลัมน์ของ WCS (`alreadysortinfo`) — แก้ที่ `export_data()` ถ้า schema เปลี่ยน
- `sorttime` เก็บเป็น epoch milliseconds (มี `FROM_UNIXTIME(sorttime/1000)`)
- log รายวันที่ `logs/<วันที่>.log`
