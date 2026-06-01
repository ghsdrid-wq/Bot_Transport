# ===== IMPORT =====
import tkinter as tk
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
import threading, requests, time, os, configparser, logging, re
from datetime import datetime, timedelta, date
from tkinter.scrolledtext import ScrolledText
import pandas as pd

CONFIG_FILE = "Config.ini"
LOG_FILE = "logJms.log"

minutes = [str(i) for i in range(60)]
hours = [f"{i:02d}:00" for i in range(24)]
hou = hours

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto JMS Bot")
        self.geometry("620x520")

        self.stop_requested = False
        # 🔥 แยก state
        self.scheduler_running = False
        self.job_running = False

        self.setup_log()
        self.setup_ui()
        self.load_config()

    # ===== LOG =====
    def setup_log(self):
        logging.basicConfig(filename=LOG_FILE, level=logging.INFO,
                            format="%(asctime)s | %(levelname)s | %(message)s")

    def log(self, msg):
        now = datetime.now().strftime("%H:%M:%S")
        text = f"[{now}] {msg}"
        logging.info(text)
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")

    # ===== LOCK UI =====
    def set_ui_state(self, running: bool):
        state = "disabled" if running else "normal"

        # START
        if self.scheduler_running:
            self.btn_start.grid_remove()
            self.btn_stop_start.grid()
        else:
            self.btn_stop_start.grid_remove()
            self.btn_start.grid()

        # RUN NOW
        if self.job_running and not self.scheduler_running:
            self.btn_run.grid_remove()
            self.btn_stop_run.grid()
        else:
            self.btn_stop_run.grid_remove()
            self.btn_run.grid()

        self.start_date.config(state=state)
        self.end_date.config(state=state)

        self.btn_start.config(state="disabled" if running else "normal")
        self.btn_run.config(state=state)
        self.btn_browse.config(state=state)

        for w in [self.token, self.path, self.filename]:
            w.config(state=state)

        self.delay.config(state=state)
        self.start_hour.config(state=state)
        self.end_hour.config(state=state)

    # ===== CLEAN =====
    def clean_text(self, text):
        return re.sub(r"\s+", "", text.replace("\n", "").replace("\r", "")).strip()

    def add_menu(self, widget):
        def paste(event=None):
            try:
                txt = self.clean_text(self.clipboard_get())
                widget.delete(0, tk.END)
                widget.insert(0, txt)
            except: pass
            return "break"

        widget.bind("<Control-v>", paste)

    # ===== CONFIG =====
    def save_config(self):
        cfg = configparser.ConfigParser()
        cfg["DEFAULT"] = {
            "token": self.token.get(),
            "path": self.path.get(),
            "delay": self.delay.get(),
            "start_date": str(self.start_date.get_date()),
            "end_date": str(self.end_date.get_date()),
            "start_hour": self.start_hour.get(),
            "end_hour": self.end_hour.get(),
            "filename": self.filename.get()
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            cfg.write(f)

    def load_config(self):
        if not os.path.exists(CONFIG_FILE): return
        cfg = configparser.ConfigParser()
        cfg.read(CONFIG_FILE, encoding="utf-8")
        d = cfg["DEFAULT"]

        self.token.insert(0, d.get("token", ""))
        self.path.insert(0, d.get("path", ""))

        self.delay.set(d.get("delay", "5"))

        # 🔥 วันที่
        if d.get("start_date"):
            try:
                dt = datetime.strptime(d.get("start_date"), "%Y-%m-%d").date()
                self.start_date.set_date(dt)
            except:
                pass

        if d.get("end_date"):
            try:
                dt = datetime.strptime(d.get("end_date"), "%Y-%m-%d").date()
                self.end_date.set_date(dt)
            except:
                pass

        # 🔥 เวลา
        self.start_hour.set(d.get("start_hour", "13:00"))
        self.end_hour.set(d.get("end_hour", "23:00"))

        # 🔥 filename
        self.filename.delete(0, tk.END)
        self.filename.insert(0, d.get("filename", "report.xlsx"))

    def get_datetime_range(self):
        start = f"{self.start_date.get_date()} {self.start_hour.get()}:00"
        end = f"{self.end_date.get_date()} {self.end_hour.get()}:59"
        return start, end


    # ===== UI =====
    def setup_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.tab_home = ttk.Frame(notebook)
        self.tab_setting = ttk.Frame(notebook)

        notebook.add(self.tab_home, text="Home")
        notebook.add(self.tab_setting, text="Setting")

        self.build_home()
        self.build_setting()

    # ===== UI =====
    def build_home(self):
        f = self.tab_home
        # ===== ROW 1 : STATUS =====
        self.status = ttk.Label(f, text="Status: Idle")
        self.status.pack(anchor="w", padx=10, pady=5)

        # ===== FRAME =====
        frame = ttk.Frame(f)
        frame.pack(fill="x", padx=10, pady=5)


    
        # ===== ROW 2 : RUN MINUTE =====
        ttk.Label(frame, text="Run minute").grid(row=0, column=0, sticky="w")

        self.delay = ttk.Combobox(frame, values=minutes, width=2, state="readonly")
        self.delay.grid(row=0, column=1, padx=2, sticky="w")
        self.delay.set("5")

        # ===== ROW : START → END + BUTTON =====
        ttk.Label(frame, text="Start").grid(row=1, column=0, sticky="w")

        self.start_date = DateEntry(frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.start_date.grid(row=1, column=1, sticky="w")

        self.start_hour = ttk.Combobox(frame, values=hou, width=5, state="readonly")
        self.start_hour.grid(row=1, column=2, sticky="w")
        self.start_hour.set("13:00")

        ttk.Label(frame, text="→").grid(row=1, column=3, sticky="w")

        ttk.Label(frame, text="End").grid(row=1, column=4, sticky="w")

        self.end_date = DateEntry(frame, width=12, date_pattern="yyyy-mm-dd", state="readonly")
        self.end_date.grid(row=1, column=5, sticky="w")

        # default date
        today = datetime.now().date()
        self.start_date.set_date(today)
        self.end_date.set_date(today + timedelta(days=1))
        
        self.end_hour = ttk.Combobox(frame, values=hours, width=5, state="readonly")
        self.end_hour.grid(row=1, column=6, sticky="w")
        self.end_hour.set("23:00")

        # ===== BUTTON =====
        # Start
        self.btn_start = ttk.Button(frame, text="Start", command=self.start)
        self.btn_stop_start = ttk.Button(frame, text="Stop", command=self.stop)

        self.btn_start.grid(row=1, column=7, padx=5, sticky="w")
        self.btn_stop_start.grid(row=1, column=7, padx=5, sticky="w")
        self.btn_stop_start.grid_remove()

        # Run
        self.btn_run = ttk.Button(frame, text="Run Now", command=self.run_thread)
        self.btn_stop_run = ttk.Button(frame, text="Stop", command=self.stop)

        self.btn_run.grid(row=1, column=8, sticky="w")
        self.btn_stop_run.grid(row=1, column=8, sticky="w")
        self.btn_stop_run.grid_remove()

        # ===== LOG FRAME =====
        log_frame = ttk.Frame(f)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Scrollbar
        self.log_text = ScrolledText(log_frame, wrap="none")
        self.log_text.pack(fill="both", expand=True)

    def build_setting(self):
        f = self.tab_setting

        # ================= SETTING =================
        # ===== Token FRAME =====
        token_frame = ttk.LabelFrame(f, text="Token", padding=10)
        token_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        token_frame.columnconfigure(1, weight=1)

        ttk.Label(token_frame, text="Auth Token").grid(row=0, column=0)
        self.token = ttk.Entry(token_frame, width=60)
        self.token.grid(row=0, column=1, padx=5, pady=5)

        # ===== Output FRAME =====
        output_frame = ttk.LabelFrame(f, text="JMS", padding=10)
        output_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")
        output_frame.columnconfigure(1, weight=1)

        ttk.Label(output_frame, text="Save Path").grid(row=0, column=0)
        self.path = ttk.Entry(output_frame, width=50)
        self.path.grid(row=0, column=1)

        self.btn_browse = ttk.Button(output_frame, text="Browse", command=self.browse)
        self.btn_browse.grid(row=0, column=2, padx=5)

        # ===== PATH FRAME =====
        name_frame = ttk.LabelFrame(f, text="File name", padding=10)
        name_frame.grid(row=2, column=0, padx=10, pady=10, sticky="ew")
        name_frame.columnconfigure(1, weight=1)

        self.filename = ttk.Entry(name_frame)
        self.filename.insert(0, "report.xlsx")
        self.filename.pack(fill="x", padx=5, pady=5)

        
    def browse(self):
        p = filedialog.askdirectory()
        if p:
            self.path.delete(0, tk.END)
            self.path.insert(0, p)

    # ===== EXPORT =====
    def export_excel_v2(self):
        base = "https://jmsgw.jtexpress.co.th/transportation"

        headers = {
            "Content-Type": "application/json;charset=UTF-8",
            "authtoken": self.clean_text(self.token.get())
        }
        start, end = self.get_datetime_range()
        payload = {
            "current": 1,
            "size": 100,
            "timeType": 2,
            "newTimeType": 1,
            "startTime": start,
            "endTime": end,
            "countryId": "1",

            "arriveNetworkCodeList": ["999004"],
            "sendNetworkCodeList": [],

            "columnList": [
                "shipmentNo",
                "shipmentState",
                "shipmentName",
                "gxType",
                "businessAttribute",
                "shifts",
                "operationModel",
                "billingWay",
                "shipmentType",
                "plateNumber",
                "plateNumberProvince",
                "trailerNumber",
                "vehicleBelongName",
                "vehicleOrigin",
                "driverName",
                "sendNetworkCode",
                "sendNetworkName",
                "loadingScanStartTime",
                "loadingScanEndTime",
                "loadCount",
                "loadingScanTotalTimeShow",
                "scanTime",
                "standByTime",
                "plannedDepartureTime",
                "actualDepartureTime",
                "delayTimeShow",
                "departureLate",
                "trackOutTime",
                "stopTimeShow",
                "actualStopTimeShow",
                "delayStopTimeShow",
                "arriveNetworkCode",
                "arriveNetworkName",
                "plannedArrivalTime",
                "predictArriveTime",
                "actualArrivalTime",
                "tardyTimeShow",
                "arrivelLate",
                "trackInTime",
                "carrierName",
                "carrierShortName",
                "carrierType",
                "useTimeShow",
                "actualUseTimeShow",
                "useWayTimeShow",
                "runningLate",
                "unScanTime",
                "unLoadLineTimeShow",
                "unLoadingScanStartTime",
                "unLoadingScanEndTime",
                "unLoadCount",
                "unLoadingScanTotalTimeShow",
                "unLoadTime",
                "vehiclelineCode",
                "vehiclelineName",
                "isAssistLine",
                "vehicleTypegroup",
                "vehicletypeName",
                "loadWeight",
                "loadCapacity",
                "vehicleDoorCnt",
                "mileage",
                "overtimeType",
                "overtimeReasons",

                # JMS ตัวล่าสุดมีเพิ่ม
                "quotationModel",
                "freightCode",
                "arriveProvince",
                "oriRegShiftCarrierName",
                "auditStatus",
                "auditRemark",
                "auditer"
            ]
        }

        # 1. create job
        requests.post(f"{base}/tmsExportTransportReport/reportExport",
                    json=payload, headers=headers)

        self.log("Create export job...")

        time.sleep(5)

        # 2. find task
        task = None
        for _ in range(30):
            if self.stop_requested:
                self.log("⛔ Stopped by user")
                return
            time.sleep(1)
            res = requests.post(f"{base}/export/selectTask",
                                json={"current": 1, "size": 20},
                                headers=headers).json()

            for r in res.get("data", {}).get("records", []):
                if r.get("state") == 2 and r.get("ossUrl"):
                    task = r
                    break

            if task:
                break

        if not task:
            raise Exception("❌ ไม่เจอไฟล์ที่เสร็จ")

        # 3. build url
        oss_path = task.get("ossUrl")
        download_url = f"https://yl-file.jtexpress.co.th/{oss_path}"

        # 4. download
        file = requests.get(download_url)

        filename = self.filename.get().strip()
        if not filename:
            filename = "report.xlsx"

        path = os.path.join(self.path.get(), filename)
        with open(path, "wb") as f:
            f.write(file.content)

        self.log(f"✅ Downloaded: {filename}")

        
    # ===== RUN =====
    def run_thread(self):
        if self.job_running:
            return
        self.save_config()
        self.set_ui_state(True)
        threading.Thread(target=self.run_job, daemon=True).start()

    def run_job(self):
        self.stop_requested = False
        self.job_running = True
        self.after(0, lambda: self.set_ui_state(True))
        self.after(0, lambda: self.status.config(text="Running..."))

        try:
            self.export_excel_v2()
            self.after(0, lambda: self.status.config(text="Success"))
            
        except Exception as e:
            self.log(str(e))
            self.after(0, lambda: self.status.config(text="Error"))

        finally:
            self.job_running = False
            if not self.scheduler_running:
                self.after(0, lambda: self.set_ui_state(False))

    # ===== SCHEDULER =====
    def scheduler(self):
        while self.scheduler_running:
            now = datetime.now()
            delay = int(self.delay.get() or 0)

            base = now.replace(minute=0, second=0, microsecond=0)
            next_run = base + timedelta(hours=1, minutes=delay)

            if next_run <= now:
                next_run += timedelta(hours=1)

            while self.scheduler_running:
                remain = int((next_run - datetime.now()).total_seconds())
                if remain <= 0: break

                m, s = divmod(remain, 60)
                self.after(0, lambda m=m, s=s:
                    self.status.config(
                        text=f"Next run: {next_run.strftime('%H:%M:%S')} ({m:02d}:{s:02d})"
                    )
                )
                time.sleep(1)

            if self.scheduler_running:
                self.run_thread()

    def start(self):
        if self.scheduler_running:
            return
        self.scheduler_running = True
        self.save_config()
        self.set_ui_state(True)
        threading.Thread(target=self.scheduler, daemon=True).start()

    def stop(self):
        self.scheduler_running = False
        self.stop_requested = True
        if not self.job_running:
            self.set_ui_state(False)
        self.status.config(text="Stopped")

if __name__ == "__main__":
    app = App()
    app.mainloop()