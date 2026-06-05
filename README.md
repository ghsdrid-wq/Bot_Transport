# Feishu Auto Report Enterprise Edition

## Overview

Feishu Auto Report Enterprise Edition เป็นระบบ Desktop Automation สำหรับงาน Operation Reporting ที่รวมการ Export ข้อมูลจาก JMS, การสร้าง Dashboard PNG จาก Excel และการส่งรายงานเข้า Feishu Chat ไว้ใน Workflow เดียว

โปรแกรมถูกออกแบบมาเพื่อลดงาน Manual และลดความผิดพลาดจากการทำงานซ้ำ ๆ โดยสามารถทำงานได้ทั้งแบบ Manual และ Scheduled Execution

---

# Table of Contents

- Overview
- Key Features
- System Architecture
- Project Structure
- Core Modules
- JMS Export Engine
- Excel Rendering Engine
- Blank Image Protection
- Feishu Integration
- Scheduler Engine
- Runtime Lock System
- Stop Request System
- Configuration Reference
- UI Overview
- Logging System
- Threading Architecture
- Error Handling
- Installation
- Build EXE
- Future Roadmap

---

# Key Features

## JMS Export Automation

- Create Export Job อัตโนมัติ
- Poll Export Status
- Download XLSX อัตโนมัติ
- Retry Request
- Export Timeout Protection
- User Cancel Support

## Dashboard PNG Generator

- Excel COM Automation
- Refresh Workbook
- Refresh Query
- Calculate Workbook
- Dynamic Date Update
- Hide #N/A Rows
- PNG Export

## PNG Validation Engine

- Blank PNG Detection
- White Image Detection
- Minimum File Size Validation
- Automatic Retry

## Feishu OpenAPI Integration

- Tenant Access Token
- Image Upload API
- Chat ID Messaging
- Retry Request Framework

## Scheduler

- Run Now
- Auto Scheduler
- Hour Based Schedule
- Minute Based Schedule
- Stop Scheduler

## Enterprise UI

- Modern CustomTkinter Interface
- Dark Theme
- Real-Time Log Viewer
- Status Badge
- Runtime Lock
- Configuration Manager

---

# System Architecture

```text
                Scheduler Engine
                        │
                        ▼
                  Task Manager
                        │
         ┌──────────────┼──────────────┐
         ▼                             ▼
    JMS Export                  Dashboard Engine
         │                             │
         ▼                             ▼
 Download XLSX                 Refresh Workbook
         │                             │
         ▼                             ▼
 Save Report                  Generate PNG
         └──────────────┬──────────────┘
                        ▼
                  Feishu Sender
                        ▼
                   Chat Group
```

---

# Project Structure

```text
project/
│
├── Bot_Fei_Main.py
├── createpng.py
├── sendfeishu.py
├── config.ini
│
├── report.xlsx
├── report.png
│
└── bot_fei_main.log
```

---

# Core Modules

## Bot_Fei_Main.py

Main Application Controller

Responsibilities

- Main GUI
- Scheduler Engine
- JMS Export Workflow
- Feishu Workflow
- Runtime Lock
- Stop Request Control
- Config Management
- Logging

---

## createpng.py

Dashboard Rendering Engine

Responsibilities

- Open Excel
- Refresh Workbook
- Calculate Workbook
- Update Report Date
- Hide Invalid Rows
- Generate PNG
- Validate PNG

---

## sendfeishu.py

Feishu Communication Layer

Responsibilities

- Get Token
- Upload Image
- Send Message
- Retry Requests

---

# JMS Export Engine

Workflow

```text
Create Export Job
        ↓
Wait Export Complete
        ↓
Get Download URL
        ↓
Download XLSX
        ↓
Save Report
```

Features

- Automatic Export Job Creation
- Export Status Polling
- Retry Download
- Timeout Protection
- Stop Request Support

---

# Excel Rendering Engine

ใช้ Microsoft Excel COM Automation

```python
win32.DispatchEx("Excel.Application")
```

Capabilities

```python
RefreshAll()
CalculateFull()
CalculateUntilAsyncQueriesDone()
CopyPicture()
Chart.Export()
```

Purpose

- Refresh Dashboard
- Update Data
- Render Dashboard
- Export PNG

---

# Offscreen Rendering Technique

โปรแกรมใช้วิธีการย้ายหน้าต่าง Excel ออกนอกหน้าจอ

```text
Left = -32000
Top  = -32000
```

แทนการ Minimize หรือ Hide

Benefits

- GDI Render ทำงานปกติ
- ลดปัญหา PNG ขาว
- เพิ่มเสถียรภาพการ Export

---

# Blank Image Protection

ระบบตรวจสอบภาพทุกครั้งก่อนใช้งาน

Validation

```text
File Exists
      ↓
File Size Check
      ↓
White Image Detection
      ↓
Variance Detection
      ↓
Export Success
```

Retry

```text
Maximum Retry = 6
```

Copy Modes

```text
screen-picture
printer-picture
screen-bitmap
printer-bitmap
```

---

# Dashboard Workflow

```text
Open Workbook
      ↓
Set Date B2
      ↓
Refresh Workbook
      ↓
Calculate Workbook
      ↓
Hide #N/A Rows
      ↓
Copy Picture
      ↓
Generate PNG
      ↓
Validate PNG
```

---

# Feishu Integration

Authentication

```text
Tenant Access Token
```

Workflow

```text
Get Token
      ↓
Upload Image
      ↓
Receive image_key
      ↓
Send Message
```

Supported APIs

- Auth API
- Image Upload API
- Message API

---

# Retry Framework

Used By

```text
Get Token
Upload Image
Send Message
Create Export Job
Download XLSX
```

Features

- Retry Counter
- Progressive Delay
- Exception Handling

---

# Scheduler Engine

Supported Modes

- Manual Run
- Auto Run

Configuration

```text
Run Every X Hours
At Minute Y
```

Example

```text
Interval = 1 Hour
Minute   = 5
```

Execution

```text
10:05
11:05
12:05
13:05
```

---

# Runtime Lock System

เมื่อ Job เริ่มทำงาน

ระบบจะ Lock

```text
Date Controls
Time Controls
Scheduler Controls
Settings Page
Config Fields
```

Purpose

- Prevent Runtime Modification
- Prevent Invalid State
- Prevent Data Corruption

---

# Stop Request System

รองรับการหยุดงานระหว่างทำงาน

Affected Components

```text
JMS Export
PNG Generator
Feishu Upload
Scheduler
```

Workflow

```text
User Press Stop
        ↓
stop_requested = True
        ↓
Current Task Stops
        ↓
Unlock UI
```

---

# Configuration Reference

## Scheduler

```ini
run_hour_interval=1
run_minute_interval=5
start_hour=13:00
end_hour=23:00
```

## JMS

```ini
jms_auth_token=
jms_save_path=
jms_filename=report.xlsx
```

## Dashboard

```ini
excel_file=
excel_sheet_index=2
excel_range=B2:V110
```

## PNG

```ini
png_output_folder=
png_filename=report.png
```

## Feishu

```ini
app_id=
app_secret=
chat_id=
```

---

# UI Overview

## Home Tab

Contains

- Interval Controls
- Date Range Controls
- Export JMS Checkbox
- Feishu Checkbox
- Start Auto
- Run Now
- Stop Button
- Status Badge
- Live Log

---

## Setting Tab

### JMS Section

- Auth Token
- Save Path
- Output Filename

### Dashboard Section

- Excel File
- Sheet Index
- Capture Range
- PNG Folder

### Feishu Section

- App ID
- App Secret
- Chat ID

---

# Logging System

Example

```text
[10:00:00] Create JMS Export Job
[10:00:15] Download XLSX
[10:01:00] Refresh Workbook
[10:01:25] PNG Ready
[10:01:40] Upload Image
[10:01:45] Send Message
[10:01:50] Done
```

Log Output

```text
bot_fei_main.log
```

---

# Threading Architecture

```text
Main UI Thread
      │
      ├── Scheduler Thread
      │
      └── Worker Thread
              │
              ├── JMS Export
              ├── PNG Generator
              └── Feishu Sender
```

Benefits

- GUI ไม่ค้าง
- Background Processing
- Responsive Interface

---

# Error Handling

Supported

- Invalid Token
- Export Timeout
- Download Failure
- Missing Excel File
- Blank PNG
- Upload Failure
- Invalid Chat ID
- User Cancel Request

---

# Installation

## Dependencies

```bash
pip install customtkinter
pip install requests
pip install pywin32
pip install tkcalendar
pip install pillow
```

---

## Run

```bash
python Bot_Fei_Main.py
```

---

# Build EXE

## OneDir

```bash
pyinstaller --onedir --windowed Bot_Fei_Main.py
```

## OneFile

```bash
pyinstaller --onefile --windowed Bot_Fei_Main.py
```

---

# Deployment Guide

1. Install Microsoft Excel
2. Install Python Dependencies
3. Configure config.ini
4. Configure JMS Token
5. Configure Feishu App
6. Test Run Now
7. Configure Scheduler
8. Build EXE
9. Deploy To Production

---

# Security Notes

- Do not commit App Secret
- Do not expose JMS Token
- Restrict access to config.ini
- Use dedicated Feishu Application

---

# Future Roadmap

- Multiple Dashboard Support
- PDF Export
- Dashboard History
- Multi Image Broadcast
- Telegram Integration
- Slack Integration
- Auto Archive
- Report Storage Database

---

# License

MIT License

---

# Author

Developed for Warehouse Operations, JMS Reporting Automation and Feishu Dashboard Broadcasting.
