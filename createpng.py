import win32com.client as win32
import pythoncom
import time, os

def run_create(excel_path, sheet_name, cell_range, output_path, log=None):
    def write(msg):
        if log:
            log(msg)
        else:
            print(msg)

    pythoncom.CoInitialize()

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        if not excel_path or not os.path.exists(excel_path):
            raise Exception("Excel file not found")

        wb = excel.Workbooks.Open(excel_path)

        # 🔥 ใช้ sheet ลำดับที่ 2 กันชื่อเพี้ยน
        ws = wb.Worksheets(2)

        write("Refreshing Excel...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

        time.sleep(2)

        write("Copying range...")
        # 🔥 บังคับให้ Excel render ก่อน
        ws.Activate()

        excel.ScreenUpdating = True
        excel.CalculateFull()

        time.sleep(1)

        # 🔥 บังคับ render โดยไม่ต้อง visible
        for _ in range(5):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

        # 🔥 retry copy กันขาว
        for i in range(3):
            try:
                ws.Range(cell_range).CopyPicture(Appearance=1, Format=2)
                time.sleep(1)
                break
            except:
                time.sleep(1)

        # 🔥 zoom + force redraw
        excel.ActiveWindow.Zoom = 100

        time.sleep(2)

        # 🔥 pump message (สำคัญมาก)
        for _ in range(5):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

        # 🔥 retry copy (กัน fail)
        for i in range(3):
            try:
                ws.Range(cell_range).CopyPicture(Appearance=1, Format=2)
                time.sleep(2)
                break
            except:
                time.sleep(2)

        chart = ws.ChartObjects().Add(
            0, 0,
            ws.Range(cell_range).Width,
            ws.Range(cell_range).Height
        )

        chart.Chart.Paste()

        time.sleep(1)

        for _ in range(3):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

        chart.Chart.Export(output_path)

        chart.Delete()

        write(f"[OK] Exported: {output_path}")

        wb.Close(False)

    finally:
        try:
            excel.Quit()
            del wb
            del excel
        except:
            pass
        pythoncom.CoUninitialize()