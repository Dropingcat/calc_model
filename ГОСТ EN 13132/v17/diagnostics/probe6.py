import openpyxl, zipfile, re, io

SRC = r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v17\EN13132_v17_LIVE.xlsx"
OUT = io.open(r"C:\Temp\v17_recalc\report5.txt", "w", encoding="utf-8")
def log(*a):
    OUT.write(" ".join(str(x) for x in a) + "\n")

z = zipfile.ZipFile(SRC)
xml = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
# Найдём формулу A52 как есть в XML (с ROW()-52 и т.п.)
m = re.search(r'<c r="A52"[^>]*>.*?</c>', xml, re.S)
log("A52 cell xml:", m.group(0) if m else "not found")
# Соберём уникальные значения для формул вида ROW()-N
matches = re.findall(r'ROW\(\)-(\d+)\)\*\(MAX', xml)
uniq = sorted(set(int(n) for n in matches))
log("\nуникальные ROW()-N в sheet2:", uniq)
log("их число:", len(uniq), "диапазон строк 52..151:", list(range(52,152))[:10], "...")

# Посмотрим sheet3 (03_ANOVA?) структуру по именам в rels
rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
log("\nrels snippet:", rels)

OUT.close()
print("done")