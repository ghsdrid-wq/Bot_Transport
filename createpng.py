import win32com.client as win32
import pythoncom
import time, os

def run_create(
    excel_path,
    sheet_name,
    cell_range,
    output_path,
    report_date=None,
    log=None,
    stop_checker=None
):
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
        ws = wb.Worksheets(int(sheet_name))
        if report_date:
            ws.Range("B2").Value = report_date

        write("Refreshing Excel...")
        wb.RefreshAll()

        excel.CalculateFull()
        excel.CalculateUntilAsyncQueriesDone()

        if stop_checker and stop_checker():
            write("ยกเลิกงาน")
            return

        time.sleep(3)

        hidden_rows = []

        last_row = ws.Cells(
            ws.Rows.Count,
            3
        ).End(-4162).Row

        write(f"last_row={last_row}")

        HEADER_ROWS = 4

        for row in range(HEADER_ROWS + 1, last_row + 1):

            value = str(ws.Cells(row, 3).Text).strip()

            if value == "#N/A":
                ws.Rows(row).Hidden = True
                hidden_rows.append(row)

        write(f"hidden_rows={len(hidden_rows)}")

        if stop_checker and stop_checker():
            write("ยกเลิกงาน")
            return

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

            if stop_checker and stop_checker():
                write("ยกเลิกงาน")
                return

            try:
                ws.Range(cell_range).CopyPicture(
                    Appearance=1,
                    Format=2
                )
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

        if stop_checker and stop_checker():
            write("ยกเลิกงาน")
            return

        time.sleep(1)

        for _ in range(3):
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

        if stop_checker and stop_checker():
            write("ยกเลิกงาน")
            return

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