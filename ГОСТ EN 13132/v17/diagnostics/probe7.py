import win32com.client, io, os

OUT = io.open(r"C:\Temp\v17_recalc\report6.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
try:
    wb = excel.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Range("B20").Value = 3.966567307590398
    ws.Range("B23").Value = 86.26602361007134
    ws.Range("B24").Value = 1.987934206239018
    ws.Range("B26").Value = 0.31778061774717337

    cases = {
        "t1": '="abc"',
        "t2": '="w(MTBE) = "&TEXT(ROUND(B20,1),"0.0")&" % мас."',
        "t3": '="±"',
        "t4": '="ν_eff="&TEXT(ROUND(B23,0),"0")',
        "t5": '="w(MTBE) = "&TEXT(ROUND(B20,1),"0.0")&" ± "&TEXT(ROUND(B26,1),"0.0")&" % мас. (P=0,95, k="&TEXT(ROUND(B24,2),"0.00")&", ν_eff="&TEXT(ROUND(B23,0),"0")&")"',
    }
    r = 1
    for name, f in cases.items():
        ws.Range("A"+str(r)).Formula = f
        try:
            val = ws.Range("A"+str(r)).Value
        except Exception as e:
            val = "ERR:"+str(e)
        log(name, repr(f))
        log("   ->", repr(val))
        r += 1

    wb.Close(SaveChanges=False)
finally:
    excel.Quit()
OUT.close()
print("done")