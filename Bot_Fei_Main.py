# -*- coding: utf-8 -*-
"""
Bot_Fei_Main.py - Combined JMS Export + Feishu Chat Sender
รวมจาก Bot_T.py + Bot_Fei_Main.py

สิ่งที่รวมแล้ว:
1) CustomTkinter UI: Home / Setting
2) Home มี DateEntry + เวลา + Run minute + checkbox Export JMS / Feishu Chat
3) Log แสดงผลผ่าน Scrollbar เดียวกัน
4) Config ทั้งหมดอยู่ใน Tab Setting และบันทึกลง config.ini
5) Feishu Chat เปลี่ยนจาก WebHook เป็น Chat ID ผ่าน OpenAPI
6) ถ้าเลือกรันทั้ง 2 งาน จะรัน Export JMS ก่อน แล้วค่อย Feishu Chat
"""

import configparser
import json
import logging
import os
import re
import sys
import threading
import time
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox


import customtkinter as ctk
import requests
from tkcalendar import DateEntry

try:
    from createpng import run_create
except Exception as exc:  # pragma: no cover
    run_create = None
    CREATEPNG_IMPORT_ERROR = exc
else:
    CREATEPNG_IMPORT_ERROR = None


APP_TITLE = "Feishu Auto Report - JMS + Chat"
LOG_FILE = "bot_fei_main.log"

MINUTES = [str(i) for i in range(60)]
RUN_HOURS = [str(i) for i in range(1, 25)]

HOURS = [f"{i:02d}:00" for i in range(24)]


# =========================
# A) PATH / CONFIG MODULE
# =========================
def resource_path(filename: str) -> str:
    """รองรับทั้งตอนรัน .py และตอน Pack เป็น .exe ด้วย PyInstaller"""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


CONFIG_FILE = resource_path("config.ini")


DEFAULT_CONFIG = {
    # Home / Scheduler
    "run_hour_interval": "1",
    "run_minute_interval": "5",
    "start_date": "",
    "end_date": "",
    "start_hour": "13:00",
    "end_hour": "23:00",
    "run_export_jms": "1",
    "run_feishu_chat": "1",

    # JMS Export
    "jms_auth_token": "",
    "jms_save_path": "",
    "jms_filename": "report.xlsx",

    # Feishu Chat / Excel to PNG
    "excel_file": "",
    "excel_sheet_index": "2",
    "excel_range": "B2:V110",
    "png_output_folder": "",
    "png_filename": "report.png",
    "app_id": "",
    "app_secret": "",
    "chat_id": "",
}


def load_config() -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    if "SETTING" not in cfg:
        cfg["SETTING"] = {}
    for key, value in DEFAULT_CONFIG.items():
        cfg["SETTING"].setdefault(key, value)
    return cfg


def save_config(cfg: configparser.ConfigParser) -> None:
    folder = os.path.dirname(CONFIG_FILE)
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as file:
        cfg.write(file)


# =========================
# B) COMMON / FEISHU MODULE
# =========================
def clean_text(text: str) -> str:
    """ลบช่องว่าง/ขึ้นบรรทัดที่มักติดมากับ token, app secret, chat id"""
    return re.sub(r"\s+", "", str(text or "").replace("\n", "").replace("\r", "")).strip()


def request_with_retry(func, retries: int = 3, delay: int = 2, log=None, name: str = "request"):
    for attempt in range(1, retries + 1):
        try:
            return func()
        except Exception as exc:
            if log:
                log(f"{name} failed ({attempt}/{retries}): {exc}")
            if attempt >= retries:
                raise
            time.sleep(delay * attempt)


def get_tenant_access_token(app_id: str, app_secret: str, log=None) -> str:
    app_id = clean_text(app_id)
    app_secret = clean_text(app_secret)
    if not app_id or not app_secret:
        raise ValueError("กรุณาใส่ App ID และ App Secret ในหน้า Setting")

    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"

    def do_request():
        res = requests.post(url, json={"app_id": app_id, "app_secret": app_secret}, timeout=20)
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Get Feishu token")
    token = data.get("tenant_access_token")
    if not token:
        raise RuntimeError(f"Get Feishu token failed: {data}")
    if log:
        log("Feishu token OK")
    return token


def upload_feishu_image(token: str, image_path: str, log=None) -> str:
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"ไม่พบไฟล์รูปภาพ: {image_path}")

    url = "https://open.feishu.cn/open-apis/im/v1/images"
    headers = {"Authorization": f"Bearer {token}"}

    def do_request():
        with open(image_path, "rb") as file:
            res = requests.post(
                url,
                headers=headers,
                files={"image": file},
                data={"image_type": "message"},
                timeout=30,
            )
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Upload image")
    if data.get("code") != 0:
        raise RuntimeError(f"Upload image failed: {data}")
    image_key = data.get("data", {}).get("image_key")
    if not image_key:
        raise RuntimeError(f"Upload image missing image_key: {data}")
    return image_key


def send_feishu_image_by_chat_id(token: str, chat_id: str, image_key: str, log=None) -> None:
    chat_id = clean_text(chat_id)
    if not chat_id:
        raise ValueError("กรุณาใส่ Feishu Chat ID ในหน้า Setting")

    url = "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json; charset=utf-8",
    }
    payload = {
        "receive_id": chat_id,
        "msg_type": "image",
        "content": json.dumps({"image_key": image_key}, ensure_ascii=False),
    }

    def do_request():
        res = requests.post(url, headers=headers, json=payload, timeout=30)
        res.raise_for_status()
        return res.json()

    data = request_with_retry(do_request, log=log, name="Send image")
    if data.get("code") != 0:
        raise RuntimeError(f"Send image failed: {data}")
    if log:
        log("ส่งรูปเข้า Feishu Chat สำเร็จ ✅")


# =========================
# C) JMS EXPORT MODULE
# =========================
JMS_COLUMNS = [
    "shipmentNo", "shipmentState", "shipmentName", "gxType", "businessAttribute",
    "shifts", "operationModel", "billingWay", "shipmentType", "plateNumber",
    "plateNumberProvince", "trailerNumber", "vehicleBelongName", "vehicleOrigin",
    "driverName", "sendNetworkCode", "sendNetworkName", "loadingScanStartTime",
    "loadingScanEndTime", "loadCount", "loadingScanTotalTimeShow", "scanTime",
    "standByTime", "plannedDepartureTime", "actualDepartureTime", "delayTimeShow",
    "departureLate", "trackOutTime", "stopTimeShow", "actualStopTimeShow",
    "delayStopTimeShow", "arriveNetworkCode", "arriveNetworkName", "plannedArrivalTime",
    "predictArriveTime", "actualArrivalTime", "tardyTimeShow", "arrivelLate",
    "trackInTime", "carrierName", "carrierShortName", "carrierType", "useTimeShow",
    "actualUseTimeShow", "useWayTimeShow", "runningLate", "unScanTime",
    "unLoadLineTimeShow", "unLoadingScanStartTime", "unLoadingScanEndTime",
    "unLoadCount", "unLoadingScanTotalTimeShow", "unLoadTime", "vehiclelineCode",
    "vehiclelineName", "isAssistLine", "vehicleTypegroup", "vehicletypeName",
    "loadWeight", "loadCapacity", "vehicleDoorCnt", "mileage", "overtimeType",
    "overtimeReasons", "quotationModel", "freightCode", "arriveProvince",
    "oriRegShiftCarrierName", "auditStatus", "auditRemark", "auditer",
]


def export_jms_excel(auth_token: str, save_folder: str, filename: str, start_time: str, end_time: str, stop_checker, log) -> str:
    auth_token = clean_text(auth_token)
    if not auth_token:
        raise ValueError("กรุณาใส่ JMS Auth Token ในหน้า Setting")
    if not save_folder:
        raise ValueError("กรุณาเลือก JMS Save Path ในหน้า Setting")

    os.makedirs(save_folder, exist_ok=True)
    filename = filename.strip() or "report.xlsx"
    if not filename.lower().endswith(".xlsx"):
        filename += ".xlsx"

    base = "https://jmsgw.jtexpress.co.th/transportation"
    headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": auth_token}
    payload = {
        "current": 1,
        "size": 100,
        "timeType": 2,
        "newTimeType": 1,
        "startTime": start_time,
        "endTime": end_time,
        "countryId": "1",
        "arriveNetworkCodeList": ["999004"],
        "sendNetworkCodeList": [],
        "columnList": JMS_COLUMNS,
    }

    log(f"สร้างงาน Export JMS: {start_time} → {end_time}")

    def create_job():
        res = requests.post(f"{base}/tmsExportTransportReport/reportExport", json=payload, headers=headers, timeout=30)
        res.raise_for_status()
        return res

    request_with_retry(create_job, log=log, name="Create JMS export job")
    time.sleep(5)

    task = None
    for _ in range(30):
        if stop_checker():
            log("หยุด Export JMS ตามคำสั่งผู้ใช้")
            return ""
        time.sleep(1)

        def get_tasks():
            res = requests.post(f"{base}/export/selectTask", json={"current": 1, "size": 20}, headers=headers, timeout=30)
            res.raise_for_status()
            return res.json()

        data = request_with_retry(get_tasks, log=log, name="Check JMS export task")
        for row in data.get("data", {}).get("records", []):
            if row.get("state") == 2 and row.get("ossUrl"):
                task = row
                break
        if task:
            break

    if not task:
        raise RuntimeError("ไม่พบไฟล์ JMS ที่ Export เสร็จภายในเวลาที่กำหนด")

    download_url = f"https://yl-file.jtexpress.co.th/{task.get('ossUrl')}"
    output_path = os.path.join(save_folder, filename)

    log("ดาวน์โหลดไฟล์ JMS...")

    def download_file():
        res = requests.get(download_url, timeout=120)
        res.raise_for_status()
        return res.content

    content = request_with_retry(download_file, log=log, name="Download JMS file")
    with open(output_path, "wb") as file:
        file.write(content)

    log(f"Export JMS สำเร็จ: {output_path}")
    return output_path

# =========================
# D) UI / MAIN APP MODULE
# =========================
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.title(APP_TITLE)
        self.geometry("980x680")
        self.resizable(False, False)

        self.cfg = load_config()
        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False
        self.last_auto_key = ""
        self.active_mode = None

        logging.basicConfig(filename=resource_path(LOG_FILE), level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", encoding="utf-8")

        self.widgets = {}
        self._build_ui()
        self._load_values_to_ui()

    # ---------- UI helpers ----------
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(
            self,
            corner_radius=14,
            anchor="w"      # เพิ่ม
        )
        self.tabview.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=14,
            pady=14
        )
        self.tabview.add("Home")
        self.tabview.add("Setting")

        self._build_home(self.tabview.tab("Home"))
        self._build_setting(self.tabview.tab("Setting"))

    def _build_home(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        title = ctk.CTkLabel(parent, text="JMS Export + Feishu Chat", font=ctk.CTkFont(size=24, weight="bold"))
        title.grid(row=0, column=0, sticky="w", padx=18, pady=(18, 6))

        control = ctk.CTkFrame(parent, corner_radius=14)
        control.grid(row=1, column=0, sticky="ew", padx=18, pady=8)
        for col in range(10):
            control.grid_columnconfigure(col, weight=0)
        control.grid_columnconfigure(9, weight=1)

        ctk.CTkLabel(
            control,
            text="Run Time"
        ).grid(row=0, column=0, padx=(14, 6), pady=12)

        self.run_hour_interval = ctk.CTkComboBox(
            control,
            values=RUN_HOURS,
            width=60,
            state="readonly"
        )
        self.run_hour_interval.grid(row=0, column=1, padx=(6,2), pady=12)

        self.run_minute_interval = ctk.CTkComboBox(
            control,
            values=MINUTES,
            width=60,
            state="readonly"
        )
        self.run_minute_interval.grid(row=0, column=2, padx=(2,12), pady=12)

        self.var_export = ctk.BooleanVar(value=True)
        self.var_feishu = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(control, text="Export JMS", variable=self.var_export).grid(row=0, column=3, padx=12, pady=12)
        ctk.CTkCheckBox(control, text="Feishu Chat", variable=self.var_feishu).grid(row=0, column=4, padx=12, pady=12)

        self.btn_start = ctk.CTkButton(control, text="▶ Start Auto", command=self.start_scheduler, width=120)
        self.btn_start.grid(row=0, column=5, padx=(14, 6), pady=12)
        self.btn_run = ctk.CTkButton(control, text="⚡ Run Now", command=self.run_now, width=110)
        self.btn_run.grid(row=0, column=6, padx=6, pady=12)

        date_frame = ctk.CTkFrame(parent, corner_radius=14)
        date_frame.grid(row=2, column=0, sticky="ew", padx=18, pady=8)
        for col in range(8):
            date_frame.grid_columnconfigure(col, weight=0)
        date_frame.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(date_frame, text="Start Date").grid(row=0, column=0, padx=(14, 6), pady=12)
        self.start_date = DateEntry(date_frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.start_date.grid(row=0, column=1, padx=6, pady=12)
        self.start_hour = ctk.CTkComboBox(date_frame, values=HOURS, width=92, state="readonly")
        self.start_hour.grid(row=0, column=2, padx=6, pady=12)

        ctk.CTkLabel(date_frame, text="→ End Date").grid(row=0, column=3, padx=(18, 6), pady=12)
        self.end_date = DateEntry(date_frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.end_date.grid(row=0, column=4, padx=6, pady=12)
        self.end_hour = ctk.CTkComboBox(date_frame, values=HOURS, width=92, state="readonly")
        self.end_hour.grid(row=0, column=5, padx=6, pady=12)

        self.status = ctk.CTkLabel(date_frame, text="Status: Idle", anchor="w")
        self.status.grid(row=0, column=6, columnspan=2, sticky="ew", padx=(20, 14), pady=12)

        log_frame = ctk.CTkFrame(parent, corner_radius=14)
        log_frame.grid(row=3, column=0, sticky="nsew", padx=18, pady=(8, 18))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_frame, text="Live Log", font=ctk.CTkFont(size=16, weight="bold")).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 4))
        self.log_box = ctk.CTkTextbox(log_frame, wrap="none")
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(4, 14))

    def _setting_entry(self, parent, row, label, key, browse=None, show=None):
        ctk.CTkLabel(parent, text=label, width=150, anchor="w").grid(row=row, column=0, padx=(14, 8), pady=7, sticky="w")
        entry = ctk.CTkEntry(parent, show=show)
        entry.grid(row=row, column=1, padx=8, pady=7, sticky="ew")
        self.widgets[key] = entry

        def paste_clean(event=None):
            try:
                text = clean_text(self.clipboard_get())
                entry.delete(0, "end")
                entry.insert(0, text)
            except Exception:
                pass
            return "break"

        entry.bind("<Control-v>", paste_clean)
        entry.bind("<Return>", lambda _e: self.save_from_ui())
        entry.bind("<FocusOut>", lambda _e: self.save_from_ui(silent=True))

        if browse:
            ctk.CTkButton(parent, text="Browse", width=90, command=lambda: self._browse_to_entry(key, browse)).grid(row=row, column=2, padx=(8, 14), pady=7)
        return entry

    def _build_setting(self, parent):
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(parent, corner_radius=14)
        scroll.grid(row=0, column=0, sticky="nsew", padx=18, pady=18)
        scroll.grid_columnconfigure(0, weight=1)

        # JMS
        jms = ctk.CTkFrame(scroll, corner_radius=12)
        jms.grid(row=0, column=0, sticky="ew", padx=4, pady=8)
        jms.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(jms, text="JMS Export Setting", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(14, 8))
        self._setting_entry(jms, 1, "Auth Token", "jms_auth_token", show="*")
        self._setting_entry(jms, 2, "Save Path", "jms_save_path", browse="folder")
        self._setting_entry(jms, 3, "Excel File Name", "jms_filename")

        # Feishu
        feishu = ctk.CTkFrame(scroll, corner_radius=12)
        feishu.grid(row=1, column=0, sticky="ew", padx=4, pady=8)
        feishu.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(feishu, text="Feishu Chat Setting", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(14, 8))
        self._setting_entry(feishu, 1, "Source Excel", "excel_file", browse="file")
        self._setting_entry(feishu, 2, "Sheet Index", "excel_sheet_index")
        self._setting_entry(feishu, 3, "Capture Range", "excel_range")
        self._setting_entry(feishu, 4, "PNG Output Folder", "png_output_folder", browse="folder")
        self._setting_entry(feishu, 5, "PNG File Name", "png_filename")
        self._setting_entry(feishu, 6, "App ID", "app_id")
        self._setting_entry(feishu, 7, "App Secret", "app_secret", show="*")
        self._setting_entry(feishu, 8, "Chat ID", "chat_id")

        save_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        save_frame.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=8,
            pady=(10, 20)
        )

        self.btn_save_setting = ctk.CTkButton(
            save_frame,
            text="💾 Save",
            width=140,
            command=self.save_from_ui
        )
        self.btn_save_setting.pack(anchor="center")

    def _browse_to_entry(self, key: str, mode: str):
        if mode == "folder":
            path = filedialog.askdirectory()
        else:
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")])
        if path:
            entry = self.widgets[key]
            entry.delete(0, "end")
            entry.insert(0, path)
            self.save_from_ui(silent=True)

    # ---------- Config ----------
    def _load_values_to_ui(self):
        s = self.cfg["SETTING"]
        today = datetime.now().date()
        self.run_hour_interval.set(
            s.get("run_hour_interval", "1")
        )

        self.run_minute_interval.set(
            s.get("run_minute_interval", "5")
        )
        self.start_hour.set(s.get("start_hour", "12:00"))
        self.end_hour.set(s.get("end_hour", "12:00"))
        self.var_export.set(s.get("run_export_jms", "1") == "1")
        self.var_feishu.set(s.get("run_feishu_chat", "1") == "1")

        for widget, key, default_date in [
            (self.start_date, "start_date", today),
            (self.end_date, "end_date", today + timedelta(days=1)),
        ]:
            value = s.get(key, "")
            try:
                widget.set_date(datetime.strptime(value, "%Y-%m-%d").date() if value else default_date)
            except Exception:
                widget.set_date(default_date)

        for key, entry in self.widgets.items():
            entry.delete(0, "end")
            entry.insert(0, s.get(key, DEFAULT_CONFIG.get(key, "")))

    def save_from_ui(self, silent: bool = False):
        s = self.cfg["SETTING"]
        s["run_hour_interval"] = self.run_hour_interval.get()
        s["run_minute_interval"] = self.run_minute_interval.get()
        s["start_date"] = str(self.start_date.get_date())
        s["end_date"] = str(self.end_date.get_date())
        s["start_hour"] = self.start_hour.get()
        s["end_hour"] = self.end_hour.get()
        s["run_export_jms"] = "1" if self.var_export.get() else "0"
        s["run_feishu_chat"] = "1" if self.var_feishu.get() else "0"
        for key, entry in self.widgets.items():
            s[key] = entry.get().strip()
        save_config(self.cfg)
        if not silent:
            self.write_log("บันทึก Config แล้ว")

    def get_setting(self, key: str) -> str:
        return self.cfg["SETTING"].get(key, DEFAULT_CONFIG.get(key, "")).strip()

    def get_datetime_range(self):
        start = f"{self.start_date.get_date()} {self.start_hour.get()}:00"
        end = f"{self.end_date.get_date()} {self.end_hour.get()}:59"
        return start, end

    # ---------- Log / State ----------
    def write_log(self, message: str):
        text = f"[{datetime.now().strftime('%H:%M:%S')}] {message}"
        logging.info(text)
        self.after(0, lambda: self._append_log(text))

    def _append_log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def set_status(self, text: str):
        self.after(0, lambda: self.status.configure(text=text))

    def stop_checker(self) -> bool:
        return self.stop_requested

    # ---------- Jobs ----------
    def run_now(self):
        if self.job_running:
            self.write_log("มีงานกำลังทำงานอยู่")
            return
        self.save_from_ui(silent=True)
        self.active_mode = "run"
        self.lock_ui()

        self.btn_start.configure(
            state="disabled"
        )

        self.btn_run.configure(
            state="normal"
        )
        threading.Thread(target=self.run_selected_jobs, daemon=True).start()

    def run_selected_jobs(self):
        if self.job_running:
            return
        self.job_running = True
        self.after(0, self.toggle_run_button)
        self.stop_requested = False
        self.current_mode = None
        self.set_status("Status: Running...")
        self.write_log("เริ่มทำงาน")

        try:
            if not self.var_export.get() and not self.var_feishu.get():
                raise ValueError("กรุณาเลือกอย่างน้อย 1 งาน: Export JMS หรือ Feishu Chat")

            if self.var_export.get():
                start, end = self.get_datetime_range()
                export_jms_excel(
                    auth_token=self.get_setting("jms_auth_token"),
                    save_folder=self.get_setting("jms_save_path"),
                    filename=self.get_setting("jms_filename"),
                    start_time=start,
                    end_time=end,
                    stop_checker=self.stop_checker,
                    log=self.write_log,
                )

            if self.stop_requested:
                self.write_log("หยุดก่อนส่ง Feishu Chat")
                self.set_status("Status: Stopped")
                return

            if self.var_feishu.get():
                self.run_feishu_chat()

            if self.stop_requested:
                self.set_status("Status: Stopped")
                return

            self.set_status("Status: Success")
            self.write_log("งานทั้งหมดเสร็จสมบูรณ์ ✅")
        except Exception as exc:
            self.set_status("Status: Error")
            self.write_log(f"ERROR: {exc}")
            messagebox.showerror("Error", str(exc))
        finally:
            self.job_running = False
            self.after(0, self.toggle_run_button)
            self.unlock_ui()

    def run_feishu_chat(self):
        if run_create is None:
            raise RuntimeError(f"ไม่สามารถ import createpng.py ได้: {CREATEPNG_IMPORT_ERROR}")

        excel_file = self.get_setting("excel_file")
        if not excel_file or not os.path.exists(excel_file):
            raise FileNotFoundError("ไม่พบ Source Excel สำหรับสร้างรูป PNG")

        output_folder = self.get_setting("png_output_folder") or os.path.dirname(excel_file)
        os.makedirs(output_folder, exist_ok=True)

        png_filename = self.get_setting("png_filename") or "report.png"
        if not png_filename.lower().endswith(".png"):
            png_filename += ".png"
        image_path = os.path.join(output_folder, png_filename)

        sheet_index = int(self.get_setting("excel_sheet_index") or "2")
        cell_range = self.get_setting("excel_range") or "B2:V110"
        self.write_log("สร้าง PNG จาก Excel...")

        if self.stop_requested:
            self.write_log("ยกเลิกก่อนสร้าง PNG")
            return

        run_create(
            excel_file,
            str(sheet_index),
            cell_range,
            image_path,
            report_date=str(self.start_date.get_date()),
            log=self.write_log,
            stop_checker=lambda: self.stop_requested
        )

        if self.stop_requested:
            self.write_log("ยกเลิกก่อน Upload Feishu")
            return

        self.write_log("อัปโหลดรูปไป Feishu...")
        token = get_tenant_access_token(self.get_setting("app_id"), self.get_setting("app_secret"), log=self.write_log)
        image_key = upload_feishu_image(token, image_path, log=self.write_log)

        if self.stop_requested:
            self.write_log("ยกเลิกก่อนส่ง Feishu")
            return

        self.write_log("ส่งรูปเข้า Feishu Chat ID...")
        send_feishu_image_by_chat_id(token, self.get_setting("chat_id"), image_key, log=self.write_log)

    def lock_ui(self):
        widgets = [
            self.run_hour_interval,
            self.run_minute_interval,
            self.start_hour,
            self.end_hour,
            self.start_date,
            self.end_date,
        ]

        for widget in widgets:
            try:
                widget.configure(state="disabled")
            except:
                pass

        self.tabview.set("Home")

        for entry in self.widgets.values():
            try:
                entry.configure(state="disabled")
            except:
                pass

        try:
            self.tabview._segmented_button.configure(state="disabled")
        except:
            pass

    def unlock_ui(self):
        widgets = [
            self.run_hour_interval,
            self.run_minute_interval,
            self.start_hour,
            self.end_hour,
            self.start_date,
            self.end_date,
        ]

        for widget in widgets:
            try:
                widget.configure(state="readonly")
            except:
                pass

        self.tabview.set("Home")

        for entry in self.widgets.values():
            try:
                entry.configure(state="normal")
            except:
                pass

        try:
            self.tabview._segmented_button.configure(state="normal")
        except:
            pass

        self.btn_start.configure(state="normal")
        self.btn_run.configure(state="normal")
        self.active_mode = None

    # ---------- Scheduler ----------
    def start_scheduler(self):
        if self.scheduler_running:
            return
        self.save_from_ui(silent=True)
        self.active_mode = "auto"
        self.lock_ui()

        self.btn_run.configure(
            state="disabled"
        )

        self.btn_start.configure(
            state="normal"
        )
        self.stop_requested = False
        self.scheduler_running = True
        self.after(0, self.toggle_auto_button)
        self.stop_requested = False
        self.write_log("เริ่ม Auto Scheduler")
        threading.Thread(target=self.scheduler_loop, daemon=True).start()

    def toggle_auto_button(self):

        if self.scheduler_running:

            self.btn_run.configure(
                state="disabled"
            )

            self.btn_start.configure(
                text="■ Stop",
                fg_color="#8a2d2d",
                hover_color="#a83232",
                command=self.stop_all
            )

        else:

            self.btn_run.configure(
                state="normal"
            )

            self.btn_start.configure(
                text="▶ Start Auto",
                fg_color="#1F6AA5",
                hover_color="#144870",
                command=self.start_scheduler
            )

    def toggle_run_button(self):

        if self.job_running:

            self.btn_start.configure(
                state="disabled"
            )

            self.btn_run.configure(
                text="■ Stop",
                fg_color="#8a2d2d",
                hover_color="#a83232",
                command=self.stop_all
            )

        else:

            self.btn_start.configure(
                state="normal"
            )

            self.btn_run.configure(
                text="⚡ Run Now",
                fg_color="#1F6AA5",
                hover_color="#144870",
                command=self.run_now
            )
    def stop_all(self):

        self.scheduler_running = False
        self.stop_requested = True

        self.set_status("Status: Stopped")
        self.write_log("รับคำสั่งหยุด")

        
    def scheduler_loop(self):

        while self.scheduler_running:

            try:
                hours = int(self.run_hour_interval.get() or 1)
                minutes = int(self.run_minute_interval.get() or 0)
            except ValueError:
                hours = 1
                minutes = 0

            now = datetime.now()

            # ตั้งนาทีตาม Combobox
            next_run = now.replace(
                minute=minutes,
                second=0,
                microsecond=0
            )

            # ถ้าเวลาปัจจุบันเลยจุดนั้นแล้ว
            # ให้ขยับทีละ X ชั่วโมง
            while next_run <= now:
                next_run += timedelta(hours=hours)

            while self.scheduler_running and datetime.now() < next_run:

                remain = int(
                    (next_run - datetime.now()).total_seconds()
                )

                m, s = divmod(max(remain, 0), 60)

                self.set_status(
                    f"Next run: {next_run.strftime('%H:%M:%S')} ({m:02d}:{s:02d})"
                )

                time.sleep(1)

            if not self.scheduler_running:
                break

            self.write_log(
                f"Auto run at {datetime.now().strftime('%H:%M:%S')}"
            )

            self.run_selected_jobs()

        self.scheduler_running = False

        self.after(
            0,
            self.toggle_auto_button
        )

        self.after(
            0,
            self.unlock_ui
        )

        self.set_status(
            "Status: Stopped"
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()
