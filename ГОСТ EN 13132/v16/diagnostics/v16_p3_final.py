# -*- coding: utf-8 -*-
"""P3 анти-хардкод: ищем в генераторе/ядре литералы из «ложных отчётов» и константы входов."""
import io, sys

targets = {
    "gen": r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16\build_en13132_v16_calc.py",
    "kernel": r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16\anova_gum.py",
}
literals = [
    "0.086197", "0.086196", "0.1557", "0.155674", "119.04", "119.038",
    "1.98", "1.9800", "8.65", "8.6488", "0.9966249", "0.996625", "3.966567",
    "0.343062", "0.3430", "0.039114", "4.0", "0.1041", "0.1406", "0.1713",
    "0.1562", "M_PROBY", "M_IS_PROBA", "Y0_PROBA", "P_IS =", "P_A =", "0.104124",
]
out = []
for tag, path in targets.items():
    src = io.open(path, encoding="utf-8").read()
    lines = src.splitlines()
    out.append("== %s == (%d строк)" % (tag, len(lines)))
    hits = 0
    for i, ln in enumerate(lines, 1):
        for lit in literals:
            if lit in ln:
                # пропускаем чисто документационные строки (docstring) если это не код? 
                # помечаем все
                out.append("  L%d: '%s' <- %s" % (i, lit, ln.strip()[:120]))
                hits += 1
                break
    if hits == 0:
        out.append("  НЕТ хардкод-литералов результата")
out.append("")
out.append("== Проверка: все входы читаются из 01_Ввод ==")
gen = io.open(targets["gen"], encoding="utf-8").read()
g = gen.splitlines()
for i, ln in enumerate(g, 1):
    if "read_inputs" in ln and "def " in ln:
        out.append("  def read_inputs L%d" % i)
    if "inputs = read_inputs" in ln or "inputs = default_inputs" in ln:
        out.append("  L%d: %s" % (i, ln.strip()[:110]))
    if "K_G2PCT =" in ln and "def " not in ln:
        out.append("  L%d: %s" % (i, ln.strip()[:110]))

with io.open(r"C:\Temp\v16_p3_final.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(out))
print("OK")