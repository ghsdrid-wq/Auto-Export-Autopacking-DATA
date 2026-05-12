import tkinter as tk
from tkinter import filedialog, ttk
from tkcalendar import DateEntry
import threading
import time
import pymysql
import pandas as pd
import configparser
from datetime import datetime, timedelta
import os
import sys
import logging
import shutil

# ===== FIX EXE PATH =====
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))

# tray + notify
import pystray
from PIL import Image, ImageDraw
from plyer import notification

CONFIG_FILE = "config.ini"
MAX_DAYS_BACK = 7  # จำกัดย้อนหลัง
# ================= CONFIG =================
def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists(CONFIG_FILE):
        config["DEFAULT"] = {
            "output_path": os.getcwd(),
            "host": "127.0.0.1",
            "user": "wcs",
            "password": "wcs",
            "database": "db_wcs",
            "port": "3306",
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)

    config.read(CONFIG_FILE, encoding="utf-8")
    return config

def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

# ================= LOG =================
if not os.path.exists("logs"):
    os.makedirs("logs")

def setup_logger():
    if not os.path.exists("logs"):
        os.makedirs("logs")

    today = datetime.now().strftime("%Y-%m-%d")
    log_file = f"logs/{today}.log"

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    # FILE
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # CONSOLE (debug)
    console = logging.StreamHandler()
    console.setFormatter(formatter)
    logger.addHandler(console)
    

# ================= NOTIFY =================
def notify(title, msg):
    try:
        notification.notify(title=title, message=msg, timeout=5)
    except:
        pass

# ================= EXPORT =================
def export_data(config, start_time, end_time, output_path):
    try:
        conn = pymysql.connect(
            host=config["host"],
            user=config["user"],
            password=config["password"],
            database=config["database"],
            port=int(config["port"]),
            charset="utf8"
        )
        print(f"RUN SQL: {start_time} -> {end_time}")
        q1 = """
        SELECT
            pipeline AS `Pipeline`,
            inductNo AS `Feeder No`,
            COUNT(billCode) AS `Loading Num`,
            COUNT(billCode) AS `Loading Valid Num`,
            '100.00%%' AS `Loading Valid Rate`,
            ROUND(COUNT(billCode), 2) AS `Loading Efficiency(per hour)`,
            '59分钟' AS `Duration of Feeder Online`,
            '无关机时间' AS `Remark`
        FROM alreadysortinfo
        WHERE FROM_UNIXTIME(sorttime/1000)
            BETWEEN %s AND %s
        GROUP BY pipeline, inductNo
        ORDER BY pipeline, CAST(inductNo AS UNSIGNED);
        """

        df1 = pd.read_sql(q1, conn, params=[start_time, end_time])

        q2 = """
        SELECT
            billCode AS `Billcode`,
            pipeline AS `Pipeline`,
            inductNo AS `Feeder No`,
            trayCode AS `Tray Code`,
            FROM_UNIXTIME(sorttime/1000) AS `Sort Time`,
            sortPortCode AS `Sorting Port`,
            weight AS `Weight`,
            destcode AS `Expect Chute Code`,
            sortSource AS `Sort Source`
        FROM alreadysortinfo
        WHERE sortSource IN ('拒绝件','无分拣计划','最大循环件')
        AND sorttime BETWEEN 
            UNIX_TIMESTAMP(%s) * 1000
            AND UNIX_TIMESTAMP(%s) * 1000
        ORDER BY sorttime;
        """

        df2 = pd.read_sql(q2, conn, params=[start_time, end_time])

        dt = datetime.strptime(start_time, "%Y-%m-%d %H:%M:%S")
        name = (dt + timedelta(hours=1)).strftime("%H") + ".xlsx"

        os.makedirs(output_path, exist_ok=True)
        path = os.path.join(output_path, name)

        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            df1.to_excel(writer, sheet_name="Throughput", index=False)
            df2.to_excel(writer, sheet_name="Abnormal", index=False)

        logging.info(f"Export OK: {name}")
        notify("Export Success", name)

    except Exception as e:
        logging.error(str(e))
        notify("Export Failed", str(e))

    finally:
        if 'conn' in locals():
            conn.close()

def reset_hourly_files(output_path):
    try:
        logging.info("🔄 Reset hourly files 00-23")

        # ===============================
        # สร้าง folder backup ตามวันที่
        # ===============================
        backup_date = datetime.now().strftime("%Y-%m-%d")
        backup_folder = os.path.join(output_path, backup_date)

        os.makedirs(backup_folder, exist_ok=True)

        # ===============================
        # ย้ายไฟล์เดิมเข้า backup
        # ===============================
        for h in range(24):
            name = f"{h:02d}.xlsx"

            old_path = os.path.join(output_path, name)
            new_path = os.path.join(backup_folder, name)

            if os.path.exists(old_path):

                # ถ้ามีไฟล์ชื่อซ้ำให้ลบทิ้งก่อน
                if os.path.exists(new_path):
                    os.remove(new_path)

                shutil.move(old_path, new_path)

        logging.info(f"📦 Backup to: {backup_folder}")

        # ===============================
        # สร้างไฟล์ใหม่
        # ===============================
        df1 = pd.DataFrame(columns=[
            "Pipeline",
            "Feeder No",
            "Loading Num",
            "Loading Valid Num",
            "Loading Valid Rate",
            "Loading Efficiency(per hour)",
            "Duration of Feeder Online",
            "Remark"
        ])

        df2 = pd.DataFrame(columns=[
            "Billcode",
            "Pipeline",
            "Feeder No",
            "Tray Code",
            "Sort Time",
            "Sorting Port",
            "Weight",
            "Expect Chute Code",
            "Sort Source"
        ])

        for h in range(24):
            name = f"{h:02d}.xlsx"
            path = os.path.join(output_path, name)

            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df1.to_excel(writer, sheet_name="Throughput", index=False)
                df2.to_excel(writer, sheet_name="Abnormal", index=False)

        notify("Reset", "Hourly files cleared")
        logging.info("✅ Reset complete (00-23)")

    except Exception as e:
        logging.error(f"Reset error: {e}")

# ================= APP =================
class App:
    def __init__(self, root):
        self.root = root
        self.config = load_config()

        self.running = True
        self.last_run_hour = None
        self.manual_running = False
        self.cancel_flag = False
        self.last_reset_date = None
        setup_logger()

        self.path_var = tk.StringVar(value=self.config["DEFAULT"]["output_path"])

        self.build_ui()
        self.update_ui_state()
        if self.running and not self.manual_running:
            self.btn.config(text="Stop", bg="#e74c3c")
            self.status.config(text="Status: Running")
            self.log("Auto Start")

        threading.Thread(target=self.auto_loop, daemon=True).start()
        self.create_tray()

    # ================= UI =================
    def build_ui(self):
        self.root.title("Auto Export")
        self.root.geometry("600x350")
        self.root.resizable(False, False)
        self.root.configure(bg="#eef2f7")

        main = tk.Frame(self.root, bg="#eef2f7", padx=10, pady=10)
        main.grid(row=0, column=0, sticky="nsew")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # ================= OUTPUT =================
        tk.Label(main, text="Output Folder", bg="#eef2f7").grid(row=0, column=0, sticky="w")

        self.entry_path = tk.Entry(main, textvariable=self.path_var, width=55)
        self.entry_path.grid(row=1, column=0, sticky="ew", padx=(0,5))

        self.btn_browse = tk.Button(main, text="Browse", command=self.browse)
        self.btn_browse.grid(row=1, column=1, sticky="e")

        # ================= RANGE + BUTTON =================
        range_frame = tk.Frame(main, bg="white", bd=1, relief="solid", padx=10, pady=8)
        range_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)

        range_frame.grid_columnconfigure(0, weight=1)

        tk.Label(range_frame, text="Manual Export", bg="white").grid(row=0, column=0, sticky="w")

        # ===== INPUT (LEFT) =====
        input_frame = tk.Frame(range_frame, bg="white")
        input_frame.grid(row=1, column=0, sticky="w")

        self.start_date = DateEntry(input_frame, width=12, date_pattern="yyyy-mm-dd")
        self.start_date.grid(row=0, column=0)

        self.start_hour = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(24)], width=3, state="readonly")
        self.start_hour.grid(row=0, column=1, padx=(2,10))

        tk.Label(input_frame, text="→", bg="white").grid(row=0, column=2)

        self.end_date = DateEntry(input_frame, width=12, date_pattern="yyyy-mm-dd")
        self.end_date.grid(row=0, column=3, padx=(10,0))

        self.end_hour = ttk.Combobox(input_frame, values=[f"{i:02d}" for i in range(24)], width=3, state="readonly")
        self.end_hour.grid(row=0, column=4, padx=(2,0))

        # ===== BUTTON (RIGHT) =====
        btn_frame = tk.Frame(range_frame, bg="white")
        btn_frame.grid(row=1, column=1, sticky="e")

        self.btn_manual = tk.Button(
            btn_frame,
            text="Export",
            bg="#e67e22",
            fg="white",
            width=10,
            command=self.manual_export_range
        )
        self.btn_manual.grid(row=0, column=0, padx=5)

        self.btn = tk.Button(
            btn_frame,
            text="Start Auto",
            bg="#2ecc71",
            fg="white",
            width=10,
            command=self.toggle
        )
        self.btn.grid(row=0, column=1)

        # ===== LOG FRAME =====
        log_frame = tk.Frame(main)
        log_frame.grid(row=3, column=0, columnspan=2, sticky="nsew")

        # ให้ขยายได้
        main.grid_rowconfigure(3, weight=1)
        main.grid_columnconfigure(0, weight=1)

        log_frame.grid_rowconfigure(0, weight=1)
        log_frame.grid_columnconfigure(0, weight=1)

        # TEXT
        self.log_text = tk.Text(log_frame, bg="black", fg="#00ff9c")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # SCROLLBAR
        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")

        # LINK กัน
        self.log_text.config(yscrollcommand=scrollbar.set)

        # ================= STATUS =================
        self.status = tk.Label(self.root, text="Status: Idle", anchor="w")
        self.status.grid(row=1, column=0, sticky="ew")

        # ================= GRID WEIGHT =================
        main.grid_rowconfigure(3, weight=1)
        main.grid_columnconfigure(0, weight=1)

        # ================= DEFAULT TIME =================
        now = datetime.now()
        self.start_date.set_date(now)
        self.end_date.set_date(now)

        self.start_hour.set(f"{(now.hour - 1) % 24:02d}")
        self.end_hour.set(f"{now.hour:02d}")

    def log(self, msg):
        self.log_text.insert("end", f"{datetime.now()} - {msg}\n")
        self.log_text.see("end")

    # ================= ACTION =================
    def browse(self):
        path = filedialog.askdirectory()
        if path:
            self.path_var.set(path)
            self.config["DEFAULT"]["output_path"] = path
            save_config(self.config)

    def toggle(self):
        if self.manual_running:
            self.log("❌ Manual กำลังทำงานอยู่")
            return

        self.running = not self.running
        self.update_ui_state()

        if self.running and not self.manual_running:
            self.btn.config(text="Stop", bg="#e74c3c")
            self.status.config(text="Status: Running")
            self.log("Auto Start")
        else:
            self.btn.config(text="Start Auto", bg="#2ecc71")
            self.status.config(text="Status: Stopped")
            self.log("Auto Stop")
        
    def update_ui_state(self):
        # ===== MANUAL RUNNING =====
        if self.manual_running:
            self.btn_manual.config(text="Cancel", bg="#e74c3c")
            self.btn.config(state="disabled")

            self.entry_path.config(state="disabled")
            self.btn_browse.config(state="disabled")

            self.start_date.config(state="disabled")
            self.end_date.config(state="disabled")
            self.start_hour.config(state="disabled")
            self.end_hour.config(state="disabled")

            self.status.config(text="Status: Manual Running")
            return

        # ===== AUTO RUNNING =====
        if self.running:
            self.btn.config(
                text="Stop",
                bg="#e74c3c",
                state="normal"
            )

            self.btn_manual.config(
                text="Export",
                bg="#e67e22",
                state="disabled"
            )

            self.entry_path.config(state="disabled")
            self.btn_browse.config(state="disabled")

            self.start_date.config(state="disabled")
            self.end_date.config(state="disabled")
            self.start_hour.config(state="disabled")
            self.end_hour.config(state="disabled")

            self.status.config(text="Status: Running")

        else:
            self.btn.config(
                text="Start Auto",
                bg="#2ecc71",
                state="normal"
            )

            self.btn_manual.config(
                text="Export",
                bg="#e67e22",
                state="normal"
            )

            self.entry_path.config(state="normal")
            self.btn_browse.config(state="normal")

            self.start_date.config(state="normal")
            self.end_date.config(state="normal")
            self.start_hour.config(state="readonly")
            self.end_hour.config(state="readonly")

            self.status.config(text="Status: Stopped")

    def manual_export_range(self):

        # ===== ถ้ากำลัง run → กด = cancel =====
        if self.manual_running:
            self.cancel_flag = True
            self.log("⛔ Cancel requested...")
            return

        # ===== เริ่ม export =====
        if self.running and not self.manual_running:
            self.log("❌ Stop Auto ก่อน")
            return

        self.manual_running = True
        self.cancel_flag = False
        self.update_ui_state()

        def task():
            try:
                start = datetime.strptime(f"{self.start_date.get()} {self.start_hour.get()}", "%Y-%m-%d %H")
                end = datetime.strptime(f"{self.end_date.get()} {self.end_hour.get()}", "%Y-%m-%d %H")

                if start >= end:
                        self.log("❌ Invalid time range")
                        return
                    
                if (datetime.now() - start).days > MAX_DAYS_BACK:
                    self.log("❌ เกินช่วงย้อนหลังที่กำหนด")
                    return
                
                self.log(f"Start Export Range: {start} -> {end}")

                current = start
                while current < end:

                    # ===== เช็ค cancel =====
                    if self.cancel_flag:
                        self.log("⛔ Export cancelled")
                        return

                    next_hour = current + timedelta(hours=1)

                    self.log(f"Exporting {current}")

                    export_data(
                        self.config["DEFAULT"],
                        current.strftime("%Y-%m-%d %H:%M:%S"),
                        (next_hour - timedelta(seconds=1)).strftime("%Y-%m-%d %H:%M:%S"),
                        self.path_var.get()
                    )

                    current = next_hour

                self.log("✅ Export Done")

            except Exception as e:
                self.log(f"❌ ERROR: {e}")

            finally:
                self.manual_running = False
                self.cancel_flag = False

                self.running = False

                self.update_ui_state()
                

        threading.Thread(target=task, daemon=True).start()
            
    def auto_loop(self):
        self.current_log_date = datetime.now().strftime("%Y-%m-%d")
        while True:
            today = datetime.now().strftime("%Y-%m-%d")
            if today != self.current_log_date:
                setup_logger()
                self.current_log_date = today
                self.log("🔄 New log file created")

            if self.running and not self.manual_running:
                now = datetime.now()
                today = datetime.now().date()
                if (now.hour == 12 and now.minute == 0 and now.second < 5
                    and self.last_reset_date != today):

                    reset_hourly_files(self.path_var.get())
                    self.last_reset_date = today
                    time.sleep(60)
                    continue
                if now.minute == 0 and now.second < 5:
                    start = now.replace(minute=0, second=0) - timedelta(hours=1)
                    end = now.replace(minute=0, second=0) - timedelta(seconds=1)

                    self.log(f"Auto Export: {start} -> {end}")
                    
                    export_data(
                        self.config["DEFAULT"],
                        start.strftime("%Y-%m-%d %H:%M:%S"),
                        end.strftime("%Y-%m-%d %H:%M:%S"),
                        self.path_var.get()
                    )

                    time.sleep(60)

            time.sleep(1)

    # ================= TRAY =================
    def create_tray(self):
        image = Image.new("RGB", (64, 64), "blue")
        draw = ImageDraw.Draw(image)
        draw.text((10, 20), "EX", fill="white")

        menu = pystray.Menu(
            pystray.MenuItem("Show", self.show),
            pystray.MenuItem("Exit", self.exit_app)
        )

        self.tray = pystray.Icon("app", image, "Auto Export", menu)
        threading.Thread(target=self.tray.run, daemon=True).start()

        self.root.protocol("WM_DELETE_WINDOW", self.hide)

    def hide(self):
        self.root.withdraw()

    def show(self):
        self.root.deiconify()

    def exit_app(self):
        self.tray.stop()
        self.root.quit()
        os._exit(0)


# ================= MAIN =================
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()