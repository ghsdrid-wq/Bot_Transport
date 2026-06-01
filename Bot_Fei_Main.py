import configparser
import os
import re
import sys
import threading
import time
from datetime import datetime, timedelta
from tkinter import END, filedialog

import customtkinter as ctk
import requests

from createpng import run_create
from sendfeishu import run_send


def resource_path(file_name):
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), file_name)
    return os.path.join(os.path.dirname(__file__), file_name)


CONFIG_FILE = resource_path("config.ini")
HOURS = [f"{hour:02d}:00" for hour in range(24)]
MINUTES = [str(minute) for minute in range(60)]
JMS_COLUMNS = [
    "shipmentNo", "shipmentState", "shipmentName", "gxType",
    "businessAttribute", "shifts", "operationModel", "billingWay",
    "shipmentType", "plateNumber", "plateNumberProvince", "trailerNumber",
    "vehicleBelongName", "vehicleOrigin", "driverName", "sendNetworkCode",
    "sendNetworkName", "loadingScanStartTime", "loadingScanEndTime", "loadCount",
    "loadingScanTotalTimeShow", "scanTime", "standByTime", "plannedDepartureTime",
    "actualDepartureTime", "delayTimeShow", "departureLate", "trackOutTime",
    "stopTimeShow", "actualStopTimeShow", "delayStopTimeShow", "arriveNetworkCode",
    "arriveNetworkName", "plannedArrivalTime", "predictArriveTime", "actualArrivalTime",
    "tardyTimeShow", "arrivelLate", "trackInTime", "carrierName", "carrierShortName",
    "carrierType", "useTimeShow", "actualUseTimeShow", "useWayTimeShow", "runningLate",
    "unScanTime", "unLoadLineTimeShow", "unLoadingScanStartTime", "unLoadingScanEndTime",
    "unLoadCount", "unLoadingScanTotalTimeShow", "unLoadTime", "vehiclelineCode",
    "vehiclelineName", "isAssistLine", "vehicleTypegroup", "vehicletypeName",
    "loadWeight", "loadCapacity", "vehicleDoorCnt", "mileage", "overtimeType",
    "overtimeReasons", "quotationModel", "freightCode", "arriveProvince",
    "oriRegShiftCarrierName", "auditStatus", "auditRemark", "auditer",
]


def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding="utf-8")
    if "SETTING" not in config:
        config["SETTING"] = {}
    return config


def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as config_file:
        config.write(config_file)


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Transport Automation Center")
        self.geometry("1040x760")
        self.minsize(900, 680)

        self.config_data = load_config()
        self.entries = {}
        self.scheduler_running = False
        self.job_running = False
        self.stop_requested = False
        self.last_run = None

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.build_ui()
        self.load_home_preferences()
        self.protocol("WM_DELETE_WINDOW", self.close_app)

    def build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0, fg_color=("#0F4C81", "#102A43"))
        header.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            header, text="TRANSPORT AUTOMATION CENTER", font=ctk.CTkFont(size=22, weight="bold")
        ).pack(anchor="w", padx=24, pady=(17, 2))
        ctk.CTkLabel(
            header, text="Export JMS reports and deliver Feishu chat updates from one workspace",
            text_color=("#DCEEFF", "#B7D7F0"),
        ).pack(anchor="w", padx=24, pady=(0, 17))

        self.tabs = ctk.CTkTabview(self, corner_radius=12)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=18, pady=18)
        self.home_tab = self.tabs.add("Home")
        self.setting_tab = self.tabs.add("Setting")
        self.build_home()
        self.build_setting()

    def build_home(self):
        self.home_tab.grid_columnconfigure(0, weight=1)
        self.home_tab.grid_rowconfigure(3, weight=1)

        summary = ctk.CTkFrame(self.home_tab, corner_radius=12)
        summary.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 6))
        summary.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(summary, text="Status", font=ctk.CTkFont(size=14, weight="bold")).grid(
            row=0, column=0, padx=(18, 10), pady=16, sticky="w"
        )
        self.status = ctk.CTkLabel(summary, text="Idle", text_color="#63D297")
        self.status.grid(row=0, column=1, padx=10, pady=16, sticky="w")

        workflow = ctk.CTkFrame(self.home_tab, corner_radius=12)
        workflow.grid(row=1, column=0, sticky="ew", padx=8, pady=6)
        ctk.CTkLabel(workflow, text="Select workflow", font=ctk.CTkFont(size=16, weight="bold")).pack(
            anchor="w", padx=18, pady=(14, 4)
        )
        ctk.CTkLabel(
            workflow, text="When both are selected, Export JMS always runs before Feishu Chat.",
            text_color=("gray40", "gray70"),
        ).pack(anchor="w", padx=18, pady=(0, 10))
        checks = ctk.CTkFrame(workflow, fg_color="transparent")
        checks.pack(fill="x", padx=18, pady=(0, 14))
        self.export_enabled = ctk.BooleanVar(value=True)
        self.feishu_enabled = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(checks, text="Export JMS", variable=self.export_enabled).pack(side="left", padx=(0, 24))
        ctk.CTkCheckBox(checks, text="Feishu Chat", variable=self.feishu_enabled).pack(side="left")

        actions = ctk.CTkFrame(self.home_tab, corner_radius=12)
        actions.grid(row=2, column=0, sticky="ew", padx=8, pady=6)
        self.start_button = ctk.CTkButton(actions, text="Start schedule", command=self.start_scheduler)
        self.start_button.pack(side="left", padx=(18, 8), pady=14)
        self.stop_button = ctk.CTkButton(
            actions, text="Stop", command=self.stop, fg_color="#C0392B", hover_color="#922B21"
        )
        self.stop_button.pack(side="left", padx=8, pady=14)
        self.run_button = ctk.CTkButton(
            actions, text="Run now", command=self.run_once, fg_color="#16856A", hover_color="#116B55"
        )
        self.run_button.pack(side="left", padx=8, pady=14)

        log_card = ctk.CTkFrame(self.home_tab, corner_radius=12)
        log_card.grid(row=3, column=0, sticky="nsew", padx=8, pady=(6, 8))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_card, text="Activity log", font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=18, pady=(14, 5)
        )
        self.log_box = ctk.CTkTextbox(log_card, wrap="none", font=("Consolas", 12))
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=18, pady=(5, 18))
        self.log_box.configure(state="disabled")

    def build_setting(self):
        self.setting_tab.grid_columnconfigure(0, weight=1)
        self.setting_tab.grid_rowconfigure(0, weight=1)
        scroll = ctk.CTkScrollableFrame(self.setting_tab, corner_radius=12)
        scroll.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        scroll.grid_columnconfigure(0, weight=1)

        self.add_section(scroll, "Schedule", [
            ("Run minute of every hour", "minute", "5", "combo", MINUTES),
        ])
        self.add_section(scroll, "Export JMS", [
            ("Auth token", "jms_token", "", "secret", None),
            ("Save folder", "jms_path", "", "folder", None),
            ("Excel file name", "jms_filename", "report.xlsx", "entry", None),
            ("Start date (YYYY-MM-DD)", "jms_start_date", str(datetime.now().date()), "entry", None),
            ("Start hour", "jms_start_hour", "13:00", "combo", HOURS),
            ("End date (YYYY-MM-DD)", "jms_end_date", str((datetime.now() + timedelta(days=1)).date()), "entry", None),
            ("End hour", "jms_end_hour", "23:00", "combo", HOURS),
        ])
        self.add_section(scroll, "Feishu Chat", [
            ("Excel source file", "excel_file", "", "file", None),
            ("PNG output folder", "output", "", "folder", None),
            ("Excel range", "range", "B2:V110", "entry", None),
            ("PNG file name", "file", "report.png", "entry", None),
            ("Chat ID", "chat_id", "", "entry", None),
            ("App ID", "app_id", "", "entry", None),
            ("App secret", "app_secret", "", "secret", None),
        ])
        ctk.CTkButton(scroll, text="Save settings", command=self.save_settings).grid(
            row=len(scroll.winfo_children()), column=0, sticky="e", padx=12, pady=(2, 16)
        )

    def add_section(self, parent, title, fields):
        row = len(parent.winfo_children())
        card = ctk.CTkFrame(parent, corner_radius=10)
        card.grid(row=row, column=0, sticky="ew", padx=10, pady=8)
        card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=16, weight="bold")).grid(
            row=0, column=0, columnspan=3, padx=14, pady=(12, 7), sticky="w"
        )
        settings = self.config_data["SETTING"]
        for index, (label, key, default, kind, values) in enumerate(fields, start=1):
            ctk.CTkLabel(card, text=label).grid(row=index, column=0, padx=14, pady=7, sticky="w")
            if kind == "combo":
                widget = ctk.CTkComboBox(card, values=values, state="readonly")
                widget.set(settings.get(key, default))
            else:
                widget = ctk.CTkEntry(card, show="•" if kind == "secret" else "")
                widget.insert(0, settings.get(key, default))
            widget.grid(row=index, column=1, padx=8, pady=7, sticky="ew")
            self.entries[key] = widget
            if kind in {"file", "folder"}:
                ctk.CTkButton(
                    card, text="Browse", width=78,
                    command=lambda current=widget, folder=kind == "folder": self.browse(current, folder),
                ).grid(row=index, column=2, padx=(0, 14), pady=7)

    def browse(self, entry, is_folder):
        path = filedialog.askdirectory() if is_folder else filedialog.askopenfilename()
        if path:
            entry.delete(0, END)
            entry.insert(0, path)

    def load_home_preferences(self):
        settings = self.config_data["SETTING"]
        self.export_enabled.set(settings.getboolean("run_export_jms", fallback=True))
        self.feishu_enabled.set(settings.getboolean("run_feishu_chat", fallback=True))

    def save_settings(self):
        settings = self.config_data["SETTING"]
        for key, widget in self.entries.items():
            settings[key] = widget.get().strip()
        settings["run_export_jms"] = str(self.export_enabled.get())
        settings["run_feishu_chat"] = str(self.feishu_enabled.get())
        save_config(self.config_data)
        self.write_log("Settings saved")

    def setting(self, key):
        return self.config_data["SETTING"].get(key, "").strip()

    def write_log(self, message):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.after(0, lambda: self._append_log(f"{stamp} | {message}\n"))

    def _append_log(self, text):
        self.log_box.configure(state="normal")
        self.log_box.insert(END, text)
        self.log_box.see(END)
        self.log_box.configure(state="disabled")

    def update_status(self, text, color="#63D297"):
        self.after(0, lambda: self.status.configure(text=text, text_color=color))

    def validate_date(self, key):
        value = self.setting(key)
        datetime.strptime(value, "%Y-%m-%d")
        return value

    def export_jms(self):
        token = re.sub(r"\s+", "", self.setting("jms_token"))
        folder = self.setting("jms_path")
        if not token:
            raise ValueError("JMS auth token is not set")
        if not folder:
            raise ValueError("JMS save folder is not set")
        os.makedirs(folder, exist_ok=True)
        start = f"{self.validate_date('jms_start_date')} {self.setting('jms_start_hour')}:00"
        end = f"{self.validate_date('jms_end_date')} {self.setting('jms_end_hour')}:59"
        headers = {"Content-Type": "application/json;charset=UTF-8", "authtoken": token}
        payload = {
            "current": 1, "size": 100, "timeType": 2, "newTimeType": 1,
            "startTime": start, "endTime": end, "countryId": "1",
            "arriveNetworkCodeList": ["999004"], "sendNetworkCodeList": [],
            "columnList": JMS_COLUMNS,
        }
        base = "https://jmsgw.jtexpress.co.th/transportation"
        response = requests.post(f"{base}/tmsExportTransportReport/reportExport", json=payload, headers=headers, timeout=30)
        response.raise_for_status()
        self.write_log("Export JMS: export job created")
        time.sleep(5)
        task = None
        for _ in range(30):
            if self.stop_requested:
                self.write_log("Export JMS: stopped by user")
                return False
            time.sleep(1)
            response = requests.post(f"{base}/export/selectTask", json={"current": 1, "size": 20}, headers=headers, timeout=30)
            response.raise_for_status()
            for record in response.json().get("data", {}).get("records", []):
                if record.get("state") == 2 and record.get("ossUrl"):
                    task = record
                    break
            if task:
                break
        if not task:
            raise RuntimeError("Export JMS: completed file was not found")
        download = requests.get(f"https://yl-file.jtexpress.co.th/{task['ossUrl']}", timeout=60)
        download.raise_for_status()
        file_name = self.setting("jms_filename") or "report.xlsx"
        output_path = os.path.join(folder, file_name)
        with open(output_path, "wb") as output_file:
            output_file.write(download.content)
        self.write_log(f"Export JMS: downloaded {output_path}")
        return True

    def send_feishu_chat(self):
        excel_path = self.setting("excel_file")
        output_dir = self.setting("output")
        if not excel_path or not os.path.exists(excel_path):
            raise ValueError("Feishu Chat: Excel source file was not found")
        if not output_dir:
            raise ValueError("Feishu Chat: PNG output folder is not set")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, self.setting("file") or "report.png")
        self.write_log("Feishu Chat: creating PNG")
        run_create(excel_path, "Sheet2", self.setting("range") or "B2:V110", output_file, log=self.write_log)
        if self.stop_requested:
            self.write_log("Feishu Chat: stopped before sending")
            return False
        self.write_log("Feishu Chat: sending image")
        run_send(output_file, self.setting("chat_id"), self.setting("app_id"), self.setting("app_secret"), log=self.write_log)
        return True

    def run_once(self):
        if self.job_running:
            self.write_log("A workflow is already running")
            return
        if not self.export_enabled.get() and not self.feishu_enabled.get():
            self.write_log("Select Export JMS, Feishu Chat, or both before running")
            return
        self.save_settings()
        selections = (self.export_enabled.get(), self.feishu_enabled.get())
        threading.Thread(target=self.run_job, args=selections, daemon=True).start()

    def run_job(self, run_export=None, run_feishu=None):
        if self.job_running:
            return
        if run_export is None or run_feishu is None:
            settings = self.config_data["SETTING"]
            run_export = settings.getboolean("run_export_jms", fallback=True)
            run_feishu = settings.getboolean("run_feishu_chat", fallback=True)
        self.job_running = True
        self.stop_requested = False
        self.update_status("Running...", "#F4D35E")
        try:
            if run_export and not self.export_jms():
                return
            if run_feishu and not self.stop_requested:
                self.send_feishu_chat()
            if self.stop_requested:
                self.update_status("Stopped", "#FF9F43")
            else:
                self.write_log("Workflow completed successfully")
                self.update_status("Completed")
        except Exception as error:
            self.write_log(f"ERROR: {error}")
            self.update_status("Error", "#FF6B6B")
        finally:
            self.job_running = False

    def start_scheduler(self):
        if self.scheduler_running:
            self.write_log("Schedule is already active")
            return
        self.save_settings()
        self.scheduler_running = True
        self.stop_requested = False
        self.write_log(f"Schedule started: run at minute {int(self.setting('minute')):02d} of every hour")
        threading.Thread(target=self.scheduler_loop, daemon=True).start()

    def scheduler_loop(self):
        while self.scheduler_running:
            now = datetime.now()
            if now.minute == int(self.setting("minute")):
                key = now.strftime("%Y-%m-%d %H:%M")
                if self.last_run != key and not self.job_running:
                    self.last_run = key
                    self.write_log(f"Scheduled run triggered at {now.strftime('%H:%M')}")
                    threading.Thread(target=self.run_job, daemon=True).start()
            time.sleep(1)

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True
        self.write_log("Stop requested")
        self.update_status("Stopped", "#FF9F43")

    def close_app(self):
        self.scheduler_running = False
        self.stop_requested = True
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
