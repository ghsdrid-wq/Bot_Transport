import tkinter as tk
from tkinter import ttk, filedialog
import threading, time, sys, os
from datetime import datetime
import configparser

from createpng import run_create
from sendfeishu import run_send

def resource_path(file):
    if getattr(sys, 'frozen', False):
        return os.path.join(os.path.dirname(sys.executable), file)
    return os.path.join(os.path.dirname(__file__), file)

CONFIG_FILE = resource_path("config.ini")

def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_FILE, encoding="utf-8")

    if "SETTING" not in config:
        config["SETTING"] = {}

    return config

def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Feishu Auto Report")
        self.root.geometry("700x500")

        self.running = False
        self.last_run = None
        self.working = False

        self.config = load_config()

        self.build_ui()

    def build_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)

        self.home_tab = ttk.Frame(notebook)
        self.setting_tab = ttk.Frame(notebook)

        notebook.add(self.home_tab, text="Home")
        notebook.add(self.setting_tab, text="Setting")

        self.build_home()
        self.build_setting()

    def build_home(self):
        frame = ttk.Frame(self.home_tab, padding=10)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Run minute").grid(row=0, column=0)

        self.minute = ttk.Combobox(
            frame,
            values=[str(i) for i in range(60)],
            width=5,
            state="readonly"
        )
        self.minute.set(self.config["SETTING"].get("minute", "5"))
        self.minute.grid(row=0, column=1)

        ttk.Button(frame, text="▶ Start", command=self.start).grid(row=0, column=2)
        ttk.Button(frame, text="⛔ Stop", command=self.stop).grid(row=0, column=3)
        ttk.Button(frame, text="⚡ Run Now", command=self.run_once).grid(row=0, column=4)

        self.log = tk.Text(frame, height=18, bg="#1e1e1e", fg="#00ff9c")
        self.log.grid(row=1, column=0, columnspan=5, sticky="nsew", pady=10)

    def build_setting(self):
        frame = ttk.Frame(self.setting_tab, padding=10)
        frame.pack(fill="both", expand=True)

        def add_row(label, key, row, browse=False, is_folder=False):
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w")

            entry = ttk.Entry(frame, width=50)
            entry.grid(row=row, column=1, padx=5)

            entry.insert(0, self.config["SETTING"].get(key, ""))

            def save(*_):
                self.config["SETTING"][key] = entry.get()
                save_config(self.config)

            entry.bind("<FocusOut>", save)
            entry.bind("<Return>", save)

            if browse:
                def browse_path():
                    if is_folder:
                        path = filedialog.askdirectory()
                    else:
                        path = filedialog.askopenfilename()

                    if path:
                        entry.delete(0, tk.END)
                        entry.insert(0, path)
                        save()

                ttk.Button(frame, text="Browse", command=browse_path).grid(row=row, column=2)

        add_row("Excel File", "excel_file", 0, True)
        add_row("Output Folder", "output", 1, True, True)
        add_row("Range", "range", 2)
        add_row("File Name", "file", 3)

        add_row("Webhook", "webhook", 4)
        add_row("Secret", "secret", 5)
        add_row("App ID", "app_id", 6)
        add_row("App Secret", "app_secret", 7)

    def write_log(self, msg):
        self.log.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n")
        self.log.see(tk.END)

    def run_job(self):
        if self.working:
            return

        self.working = True

        try:
            cfg = self.config["SETTING"]
            
            if not cfg.get("webhook"):
                raise Exception("Webhook not set")

            if not cfg.get("app_id") or not cfg.get("app_secret"):
                raise Exception("App ID / Secret not set")

            excel_path = cfg.get("excel_file")
            if not excel_path or not os.path.exists(excel_path):
                raise Exception("Excel file not found")

            output_dir = cfg.get("output")
            if not output_dir:
                raise Exception("Output folder not set")

            os.makedirs(output_dir, exist_ok=True)

            file_name = cfg.get("file") or "report.png"
            output_file = os.path.join(output_dir, file_name)

            self.write_log("Start Create PNG")

            run_create(
                excel_path,
                "Sheet2",
                cfg.get("range", "B2:V110"),
                output_file,
                log=self.write_log
            )

            self.write_log("Send to Feishu")

            run_send(
                output_dir,
                cfg.get("webhook"),
                cfg.get("secret"),
                cfg.get("app_id"),
                cfg.get("app_secret"),
                log=self.write_log
            )

            self.write_log("Done ✅")

        except Exception as e:
            self.write_log(f"ERROR: {e}")

        finally:
            self.working = False

    def run_once(self):
        threading.Thread(target=self.run_job, daemon=True).start()

    def loop(self):
        while self.running:
            now = datetime.now()

            if now.minute == int(self.minute.get()):
                key = now.strftime("%Y-%m-%d %H:%M")

                if self.last_run != key:
                    self.last_run = key
                    self.write_log(f"Auto run at {now.strftime('%H:%M')}")
                    threading.Thread(target=self.run_job, daemon=True).start()

            time.sleep(1)

    def start(self):
        self.running = True
        self.config["SETTING"]["minute"] = self.minute.get()
        save_config(self.config)

        self.write_log("Auto started")
        threading.Thread(target=self.loop, daemon=True).start()

    def stop(self):
        self.running = False
        self.write_log("Stopped")

root = tk.Tk()
app = App(root)
root.mainloop()