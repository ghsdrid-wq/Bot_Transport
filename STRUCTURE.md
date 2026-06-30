# STRUCTURE — Bot_Transport (Feishu Auto Report — JMS + Chat)

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic ฟังก์ชัน/คลาส, เปลี่ยน config key, เปลี่ยน JMS payload/columns, หรือเปลี่ยน flow — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
แอป GUI (**CustomTkinter**) ทำงาน 2 อย่างตามรอบเวลา/Run Now:
1. **Export JMS** — สร้าง report งานขนส่ง (transport) จาก JMS J&T → ดาวน์โหลดเป็น Excel
2. **Feishu Chat** — เอา Excel มาแคปเป็น PNG แล้วส่งเข้า **Feishu Chat ตาม Chat ID** (ผ่าน OpenAPI ไม่ใช่ webhook)

ถ้าเลือกทั้งสอง → รัน Export ก่อน แล้วค่อยส่ง Feishu

## วิธีรัน / Entry point
- รัน: `python Bot_Fei_Main.py` → คลาส `App(ctk.CTk)` (แท็บ Home / Setting) — **ตัวหลักที่ใช้งาน**
- `Bot_T.py` = เวอร์ชันเก่า (Tkinter ธรรมดา) ที่ทำเฉพาะ Export JMS — เก็บไว้อ้างอิง

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `Bot_Fei_Main.py` | **ตัวหลัก** — UI + scheduler + โมดูล Feishu (token/upload/send by chat_id) + `export_jms_excel()` (JMS transport export) |
| `createpng.py` | Excel → PNG ด้วย win32com (ย้าย Excel ออกนอกจอ, ซ่อน N/A, retry หลายโหมด CopyPicture, เช็ครูปขาว) — `run_create()` |
| `sendfeishu.py` | ฟังก์ชันส่ง Feishu แบบ **webhook + HMAC sign** (ทางเลือก/เก่า; ตัวหลักใช้ส่งแบบ chat_id ใน Bot_Fei_Main) |
| `Bot_T.py` | เวอร์ชันเก่า Export-only (Tkinter) |
| `requirements.txt` | dependency list |

## ฟังก์ชันสำคัญใน Bot_Fei_Main.py
- `export_jms_excel(auth_token, save_folder, filename, start_time, end_time, ...)` — POST สร้าง job ที่ `transportation/tmsExportTransportReport/reportExport` (ปลายทาง `arriveNetworkCodeList=["999004"]`, `columnList`=`JMS_COLUMNS` ~70 คอลัมน์) → poll `export/selectTask` → ดาวน์โหลดจาก `yl-file.jtexpress.co.th`
- `run_feishu_chat()` — `createpng.run_create()` → `get_tenant_access_token` → `upload_feishu_image` → `send_feishu_image_by_chat_id`
- `scheduler_loop()` — รันทุก interval (ชั่วโมง+นาที), `run_selected_jobs()` ตาม checkbox

## Config (`config.ini`, section `[SETTING]`)
- Scheduler: `run_hour_interval`, `run_minute_interval`, `start_date/end_date`, `start_hour/end_hour`, `run_export_jms`, `run_feishu_chat`
- JMS: `jms_auth_token`, `jms_save_path`, `jms_filename`
- Feishu/PNG: `excel_file`, `excel_sheet_index`, `excel_range` (default `B2:V110`), `png_output_folder`, `png_filename`, `app_id`, `app_secret`, `chat_id`

## Dependencies / บริการภายนอก
- `customtkinter`, `tkcalendar`, `pywin32` (ต้องมี Microsoft Excel), `requests`
- JMS J&T (`jmsgw.jtexpress.co.th/transportation`, ไฟล์ที่ `yl-file.jtexpress.co.th`), Feishu OpenAPI

## ข้อควรระวัง
- ต้องรันบน Windows + Excel (createpng ใช้ COM)
- `jms_auth_token` หมดอายุได้
- ถ้า JMS เพิ่ม/ลดคอลัมน์ ต้องแก้ `JMS_COLUMNS` ใน `Bot_Fei_Main.py` (และ `Bot_T.py` ถ้ายังใช้)
