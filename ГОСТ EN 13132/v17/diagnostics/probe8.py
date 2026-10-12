import win32com.client, io

OUT = io.open(r"C:\Temp\v17_recalc\report7.txt", "w", encoding="utf-8")
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
        "a1": '="x "&TEXT(ROUND(B20,1),"0.0")',
        "a2": '=TEXT(B20,"0.0")',
        "a3": '=TEXT(B20,"0,0")',          # русская маска
        "a4": '="x"&" ± "&"y"',
        "a5": '="w = "&TEXT(B20,"0.0")',
        "a6": '=ROUND(B26,1)',
        "a7": '=TEXT(ROUND(B26,1),"0.0")',
        "a8": '="k="&TEXT(ROUND(B24,2),"0.00")',
        "a9": '="ν_eff="&TEXT(ROUND(B23,0),"0")',
        "a10": '="x±y"',
        "a11": '="("&TEXT(ROUND(B24,2),"0.00")&")"',
        "a12": '=A1&""',
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