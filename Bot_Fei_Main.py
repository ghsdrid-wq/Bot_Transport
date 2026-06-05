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
            corner_radius=0,
            anchor="w",
            fg_color="#0a0f1e",
            segmented_button_fg_color="#0a0f1e",
            segmented_button_selected_color="#1a2744",
            segmented_button_selected_hover_color="#1e2f55",
            segmented_button_unselected_color="#0a0f1e",
            segmented_button_unselected_hover_color="#111827",
        )
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        self.tabview.add("  Home  ")
        self.tabview.add("  Setting  ")

        self._build_home(self.tabview.tab("  Home  "))
        self._build_setting(self.tabview.tab("  Setting  "))

    # ── colour tokens ────────────────────────────────────────────────
    C = {
        "bg"       : "#0a0f1e",
        "surface"  : "#0f1629",
        "card"     : "#131d35",
        "border"   : "#1e2d4a",
        "accent"   : "#3b82f6",
        "accent2"  : "#06b6d4",
        "success"  : "#10b981",
        "danger"   : "#ef4444",
        "warn"     : "#f59e0b",
        "text"     : "#e2e8f0",
        "muted"    : "#64748b",
        "log_bg"   : "#060c1a",
    }

    def _card(self, parent, **kw):
        return ctk.CTkFrame(
            parent,
            fg_color=self.C["card"],
            corner_radius=16,
            border_width=1,
            border_color=self.C["border"],
            **kw
        )

    def _label(self, parent, text, size=13, weight="normal", color=None, **kw):
        return ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=size, weight=weight),
            text_color=color or self.C["text"],
            **kw
        )

    def _btn(self, parent, text, command, color=None, hover=None, width=120, height=38):
        c = color or self.C["accent"]
        h = hover or "#2563eb"
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=width,
            height=height,
            fg_color=c,
            hover_color=h,
            corner_radius=10,
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
        )

    def _build_home(self, parent):
        parent.configure(fg_color=self.C["bg"])
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(3, weight=1)

        # ── Header bar ─────────────────────────────────────────────
        header = ctk.CTkFrame(parent, fg_color=self.C["surface"], corner_radius=0, height=64)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        dot_frame = ctk.CTkFrame(header, fg_color="transparent")
        dot_frame.grid(row=0, column=0, padx=20, pady=0, sticky="w")
        for i, col in enumerate(["#ef4444", "#f59e0b", "#10b981"]):
            ctk.CTkFrame(dot_frame, width=12, height=12, corner_radius=6, fg_color=col).grid(row=0, column=i, padx=3, pady=26)

        self._label(
            header,
            "BOT  JMSKKN",
            size=15, weight="bold",
            color=self.C["text"],
        ).grid(row=0, column=1, padx=0, pady=20, sticky="w")

        self.status_badge = ctk.CTkLabel(
            header,
            text="● IDLE",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=self.C["success"],
            fg_color="#0d2318",
            corner_radius=20,
            width=100,
            height=28,
        )
        self.status_badge.grid(row=0, column=2, padx=20, pady=18, sticky="e")

        # ── Control card — single row ──────────────────────────────
        # Interval | divider | [Export JMS] [Feishu Chat] | spacer | [Start Auto] [Run Now]
        ctrl = self._card(parent)
        ctrl.grid(row=1, column=0, sticky="ew", padx=20, pady=(16, 8))
        ctrl.grid_columnconfigure(8, weight=1)   # spacer pushes buttons to right

        # Interval
        self._label(ctrl, "Interval", size=12, color=self.C["muted"]).grid(
            row=0, column=0, padx=(18, 6), pady=16, sticky="w")
        self.run_hour_interval = ctk.CTkComboBox(
            ctrl, values=RUN_HOURS, width=66, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.run_hour_interval.grid(row=0, column=1, padx=3, pady=16)
        self._label(ctrl, "hr", size=12, color=self.C["muted"]).grid(
            row=0, column=2, padx=(2, 4), pady=16)
        self.run_minute_interval = ctk.CTkComboBox(
            ctrl, values=MINUTES, width=66, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.run_minute_interval.grid(row=0, column=3, padx=3, pady=16)
        self._label(ctrl, "min", size=12, color=self.C["muted"]).grid(
            row=0, column=4, padx=(2, 16), pady=16)

        # Divider
        ctk.CTkFrame(ctrl, width=1, height=28, fg_color=self.C["border"]).grid(
            row=0, column=5, padx=4, pady=16)

        # Checkboxes (ก่อน buttons)
        self.var_export = ctk.BooleanVar(value=True)
        self.var_feishu = ctk.BooleanVar(value=True)
        chk_style = dict(
            font=ctk.CTkFont(size=13),
            checkbox_width=18, checkbox_height=18,
            corner_radius=5,
            fg_color=self.C["accent"],
            border_color=self.C["border"],
            hover_color="#2563eb",
            text_color=self.C["text"],
        )
        ctk.CTkCheckBox(ctrl, text="Export JMS",  variable=self.var_export, **chk_style).grid(
            row=0, column=6, padx=(16, 10), pady=16)
        ctk.CTkCheckBox(ctrl, text="Feishu Chat", variable=self.var_feishu, **chk_style).grid(
            row=0, column=7, padx=(0, 8), pady=16)

        # column=8 = spacer (weight=1)

        # Buttons (ขวาสุด)
        self.btn_start = self._btn(ctrl, "▶  Start Auto", self.start_scheduler, width=140)
        self.btn_start.grid(row=0, column=9, padx=(0, 8), pady=16)
        self.btn_run = self._btn(ctrl, "⚡  Run Now", self.run_now,
                                 color=self.C["success"], hover="#059669", width=130)
        self.btn_run.grid(row=0, column=10, padx=(0, 18), pady=16)

        # ── Date range card ────────────────────────────────────────
        date_card = self._card(parent)
        date_card.grid(row=2, column=0, sticky="ew", padx=20, pady=8)
        date_card.grid_columnconfigure(7, weight=1)

        self._label(date_card, "DATE RANGE", size=10, weight="bold", color=self.C["muted"]).grid(
            row=0, column=0, columnspan=8, sticky="w", padx=18, pady=(14, 2))

        # Start
        self._label(date_card, "Start", size=12, color=self.C["muted"]).grid(row=1, column=0, padx=(18, 6), pady=(4, 14), sticky="w")
        self.start_date = DateEntry(
            date_card, width=13, date_pattern="yyyy-mm-dd", state="readonly",
            background="#1a2744", foreground="#e2e8f0",
            selectbackground="#3b82f6", selectforeground="#ffffff",
            font=("Segoe UI", 12),
        )
        self.start_date.grid(row=1, column=1, padx=4, pady=(4, 14))
        self.start_hour = ctk.CTkComboBox(
            date_card, values=HOURS, width=96, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.start_hour.grid(row=1, column=2, padx=4, pady=(4, 14))

        # Arrow
        self._label(date_card, "→", size=16, color=self.C["accent2"]).grid(row=1, column=3, padx=12, pady=(4, 14))

        # End
        self._label(date_card, "End", size=12, color=self.C["muted"]).grid(row=1, column=4, padx=(0, 6), pady=(4, 14), sticky="w")
        self.end_date = DateEntry(
            date_card, width=13, date_pattern="yyyy-mm-dd", state="readonly",
            background="#1a2744", foreground="#e2e8f0",
            selectbackground="#3b82f6", selectforeground="#ffffff",
            font=("Segoe UI", 12),
        )
        self.end_date.grid(row=1, column=5, padx=4, pady=(4, 14))
        self.end_hour = ctk.CTkComboBox(
            date_card, values=HOURS, width=96, state="readonly",
            fg_color=self.C["surface"], border_color=self.C["border"],
            button_color=self.C["accent"], dropdown_fg_color=self.C["card"],
            font=ctk.CTkFont(size=13),
        )
        self.end_hour.grid(row=1, column=6, padx=4, pady=(4, 14))

        # Status pill
        self.status = ctk.CTkLabel(
            date_card,
            text="Status: Idle",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=self.C["muted"],
            anchor="w",
        )
        self.status.grid(row=1, column=7, sticky="ew", padx=(16, 18), pady=(4, 14))

        # ── Live Log card ──────────────────────────────────────────
        log_card = self._card(parent)
        log_card.grid(row=3, column=0, sticky="nsew", padx=20, pady=(8, 20))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        log_header = ctk.CTkFrame(log_card, fg_color="transparent")
        log_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(14, 4))
        log_header.grid_columnconfigure(0, weight=1)

        self._label(log_header, "LIVE LOG", size=10, weight="bold", color=self.C["muted"]).grid(
            row=0, column=0, sticky="w")

        ctk.CTkButton(
            log_header,
            text="Clear",
            width=60, height=26,
            fg_color=self.C["surface"],
            hover_color=self.C["border"],
            text_color=self.C["muted"],
            corner_radius=8,
            font=ctk.CTkFont(size=11),
            command=lambda: self.log_box.delete("1.0", "end"),
        ).grid(row=0, column=1, sticky="e")

        self.log_box = ctk.CTkTextbox(
            log_card,
            wrap="none",
            font=("Consolas", 12),
            fg_color=self.C["log_bg"],
            text_color="#93c5fd",
            corner_radius=10,
            scrollbar_button_color=self.C["border"],
            scrollbar_button_hover_color=self.C["accent"],
        )
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))

        # tag colours
        try:
            self.log_box.tag_config("INFO",    foreground="#93c5fd")
            self.log_box.tag_config("SUCCESS", foreground="#86efac")
            self.log_box.tag_config("WARN",    foreground="#fbbf24")
            self.log_box.tag_config("ERROR",   foreground="#fca5a5")
            self.log_box.tag_config("START",   foreground="#67e8f9")
        except Exception:
            pass

    def _setting_entry(self, parent, row, label, key, browse=None, show=None):
        self._label(parent, label, size=12, color=self.C["muted"], width=160, anchor="w").grid(
            row=row, column=0, padx=(18, 8), pady=7, sticky="w")
        entry = ctk.CTkEntry(
            parent,
            show=show,
            fg_color=self.C["surface"],
            border_color=self.C["border"],
            text_color=self.C["text"],
            placeholder_text_color=self.C["muted"],
            corner_radius=8,
            height=36,
            font=ctk.CTkFont(family="Segoe UI", size=13),
        )
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
            ctk.CTkButton(
                parent, text="Browse", width=88, height=36,
                fg_color=self.C["surface"],
                hover_color=self.C["border"],
                text_color=self.C["muted"],
                border_width=1,
                border_color=self.C["border"],
                corner_radius=8,
                font=ctk.CTkFont(size=12),
                command=lambda: self._browse_to_entry(key, browse),
            ).grid(row=row, column=2, padx=(8, 18), pady=7)
        return entry

    def _section_title(self, parent, row, icon, title):
        f = ctk.CTkFrame(parent, fg_color="transparent")
        f.grid(row=row, column=0, columnspan=3, sticky="ew", padx=14, pady=(18, 4))
        self._label(f, icon, size=16, color=self.C["accent"]).grid(row=0, column=0, padx=(4, 8))
        self._label(f, title, size=14, weight="bold", color=self.C["text"]).grid(row=0, column=1, sticky="w")
        ctk.CTkFrame(f, height=1, fg_color=self.C["border"]).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0))

    def _build_setting(self, parent):
        parent.configure(fg_color=self.C["bg"])
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        scroll = ctk.CTkScrollableFrame(
            parent,
            corner_radius=0,
            fg_color=self.C["bg"],
            scrollbar_button_color=self.C["border"],
            scrollbar_button_hover_color=self.C["accent"],
        )
        scroll.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        scroll.grid_columnconfigure(0, weight=1)

        # ── JMS card ────────────────────────────────────────────────
        jms = self._card(scroll)
        jms.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 8))
        jms.grid_columnconfigure(1, weight=1)
        self._section_title(jms, 0, "⬇", "JMS Export")
        self._setting_entry(jms, 1, "Auth Token", "jms_auth_token", show="*")
        self._setting_entry(jms, 2, "Save Path", "jms_save_path", browse="folder")
        self._setting_entry(jms, 3, "Excel File Name", "jms_filename")

        # ── Feishu card ─────────────────────────────────────────────
        feishu = self._card(scroll)
        feishu.grid(row=1, column=0, sticky="ew", padx=20, pady=8)
        feishu.grid_columnconfigure(1, weight=1)
        self._section_title(feishu, 0, "🚀", "Feishu Chat")
        self._setting_entry(feishu, 1, "Source Excel",   "excel_file",         browse="file")
        self._setting_entry(feishu, 2, "Sheet Index",    "excel_sheet_index")
        self._setting_entry(feishu, 3, "Capture Range",  "excel_range")
        self._setting_entry(feishu, 4, "PNG Output Folder", "png_output_folder", browse="folder")
        self._setting_entry(feishu, 5, "PNG File Name",  "png_filename")
        self._setting_entry(feishu, 6, "App ID",         "app_id")
        self._setting_entry(feishu, 7, "App Secret",     "app_secret",         show="*")
        self._setting_entry(feishu, 8, "Chat ID",        "chat_id")

        # ── Save button ─────────────────────────────────────────────
        btn_row = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_row.grid(row=2, column=0, sticky="ew", padx=20, pady=(8, 28))
        self._btn(btn_row, "💾  Save Settings", self.save_from_ui, width=160).pack(side="left", padx=4)

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
    # ── Log level → tag map ──────────────────────────────────────
    _LOG_TAGS = {
        "INFO"    : ("INFO",    "◉"),
        "OK"      : ("SUCCESS", "✔"),
        "WARN"    : ("WARN",    "⚠"),
        "ERROR"   : ("ERROR",   "✖"),
        "START"   : ("START",   "▶"),
        "SECTION" : ("INFO",    "─"),
    }

    def write_log(self, message: str, level: str = "INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        tag, icon = self._LOG_TAGS.get(level.upper(), ("INFO", "◉"))
        text = f"[{ts}]  {icon}  {message}"
        logging.info(text)
        self.after(0, lambda t=text, g=tag: self._append_log(t, g))

    def _append_log(self, text: str, tag: str = "INFO"):
        try:
            self.log_box.tag_config("INFO",    foreground="#93c5fd")
            self.log_box.tag_config("SUCCESS", foreground="#86efac")
            self.log_box.tag_config("WARN",    foreground="#fbbf24")
            self.log_box.tag_config("ERROR",   foreground="#fca5a5")
            self.log_box.tag_config("START",   foreground="#67e8f9")
        except Exception:
            pass
        self.log_box.insert("end", text + "\n", tag)
        self.log_box.see("end")

    # badge config map: keyword → (dot_color, bg, text_color, label)
    _BADGE = {
        "Running"  : ("#f59e0b", "#2d1f0a", "#fbbf24", "● RUNNING"),
        "Success"  : ("#10b981", "#0d2318", "#86efac", "● DONE"),
        "Error"    : ("#ef4444", "#2d0f0f", "#fca5a5", "● ERROR"),
        "Stopped"  : ("#64748b", "#111827", "#94a3b8", "● STOPPED"),
        "Next run" : ("#3b82f6", "#0e1f3d", "#93c5fd", "● AUTO"),
        "Idle"     : ("#10b981", "#0d2318", "#86efac", "● IDLE"),
    }

    def set_status(self, text: str):
        def _update():
            self.status.configure(text=text)
            for kw, (_, bg, fg, badge_text) in self._BADGE.items():
                if kw.lower() in text.lower():
                    self.status_badge.configure(
                        text=badge_text,
                        text_color=fg,
                        fg_color=bg,
                    )
                    return
            # default
            self.status_badge.configure(text="● IDLE", text_color=self.C["success"], fg_color="#0d2318")
        self.after(0, _update)

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

        self.tabview.set("  Home  ")

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

        self.tabview.set("  Home  ")

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
