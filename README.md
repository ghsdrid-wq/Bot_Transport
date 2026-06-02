# Feishu Auto Report - JMS + Chat

## Overview

Python desktop application สำหรับ:

- Export รายงานจาก JMS อัตโนมัติ
- สร้าง PNG จาก Excel Dashboard
- ส่งรูปเข้า Feishu Chat ผ่าน Chat ID
- Auto Scheduler
- Manual Run
- Runtime UI Lock
- Stop / Cancel Process
- Retry & Recovery System

โปรแกรมรวม JMS Export และ Feishu Chat Sender ไว้ในระบบเดียว

---

## Features

- JMS Report Export
- Feishu Chat Integration
- Chat ID Message Sender
- Excel Dashboard to PNG
- Auto Scheduler
- Run Now
- Stop Running Task
- Retry Request System
- Real-time Log Viewer
- Config Persistence
- Runtime UI Lock
- Background Thread Processing
- Windows Desktop GUI
- Build EXE Support

---

## Tech Stack

- Python
- CustomTkinter
- Requests
- tkcalendar
- Win32COM
- PythonCOM
- Threading
- ConfigParser

---

## Project Structure

```text
project/
│
├── Bot_Fei_Main.py
├── createpng.py
├── sendfeishu.py
├── config.ini
├── report.png
│
└── bot_fei_main.log
```

---

## System Workflow

```text
Export JMS Report
        ↓
Create PNG From Excel
        ↓
Get Feishu Token
        ↓
Upload Image
        ↓
Send To Chat ID
        ↓
Complete
```

---

## JMS Export Module

รองรับ:

- Export Transportation Report
- Auto Download Excel
- Export Time Range
- Save Path Selection
- Custom Filename

Workflow:

```text
Create Export Job
        ↓
Poll Export Status
        ↓
Get Download URL
        ↓
Download XLSX
        ↓
Save File
```

---

## Feishu Integration

ใช้ OpenAPI โดยตรง

รองรับ:

- Tenant Access Token
- Image Upload API
- Chat ID Message API
- Retry Request

Workflow:

```text
Get Token
      ↓
Upload Image
      ↓
Receive image_key
      ↓
Send Image To Chat ID
```

---

## Dashboard Generator

ไฟล์:

```text
createpng.py
```

ใช้:

```python
win32.DispatchEx("Excel.Application")
```

สำหรับ:

- เปิด Excel จริง
- Refresh Workbook
- Calculate Workbook
- Export Dashboard PNG

---

## Excel Processing

Workflow:

```text
Open Workbook
      ↓
RefreshAll()
      ↓
CalculateFull()
      ↓
Hide #N/A Rows
      ↓
CopyPicture()
      ↓
Chart.Export()
      ↓
PNG Output
```

---

## Scheduler System

รองรับ:

- Run Every X Hours
- Run Every X Minutes
- Manual Run
- Auto Run
- Stop Scheduler

ตัวอย่าง:

```text
Run every 1 hour 5 minutes
```

---

## GUI Components

### Home

- Run Now
- Start Auto
- Stop Process
- Date Range
- Time Range
- Export JMS
- Feishu Chat
- Live Log

### Setting

- JMS Auth Token
- Save Path
- Excel File
- Sheet Index
- Capture Range
- PNG Output Folder
- App ID
- App Secret
- Chat ID

---

## Runtime Protection

ระหว่างรัน:

- Lock Setting UI
- Lock Scheduler Setting
- Prevent Concurrent Run
- Safe Stop Request

---

## Retry System

ใช้:

```python
request_with_retry()
```

รองรับ:

- Token Request
- Export Request
- Download Request
- Upload Image
- Send Message

---

## Config Example

```ini
run_hour_interval=1
run_minute_interval=5

excel_sheet_index=2
excel_range=B2:V110

png_filename=report.png
```

---

## Installation

### Install Dependencies

```bash
pip install customtkinter requests tkcalendar pywin32
```

### Run

```bash
python Bot_Fei_Main.py
```

---

## Build EXE

```bash
pyinstaller --onefile --windowed Bot_Fei_Main.py
```

---

## Error Handling

รองรับ:

- Missing Token
- Invalid Chat ID
- Excel Not Found
- PNG Export Failure
- Upload Failure
- Download Failure
- Export Timeout
- User Cancel Request

---

## License

MIT License
