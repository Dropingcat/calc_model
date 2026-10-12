import sys, os
import win32com.client

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
DST = r"C:\Temp\v17_recalc\EN13132_v17_RECALC.xlsx"

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Open(SRC, UpdateLinks=0, ReadOnly=False)
    try:
        excel.CalculateFull()
    except Exception as e:
        print("CalculateFull warning:", e)
    wb.SaveCopyAs(DST)
    print("Saved copy:", DST)
    print("CalcMode now:", excel.Calculation)
    wb.Close(SaveChanges=False)
finally:
    excel.Quit()
print("Excel done")