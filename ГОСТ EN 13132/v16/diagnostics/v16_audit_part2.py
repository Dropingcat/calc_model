# -*- coding: utf-8 -*-
"""Аудит v16 part2: независимый пересчёт u_repr(smooth) и импорт Word-структуры."""
import io, json, math, sys, warnings
sys.path.insert(0, r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16")
import anova_gum as ag
import numpy as np

OUT = {}
OUT["table1"] = ag.TABLE1

# независимый пересчёт по узлам табл.1 (R05: на границе — более высокий диапазон)
def manual_smooth_u_repr(w, TABLE, mode="mean2"):
    ranges = TABLE["ranges"]
    bounds = sorted({float(lo) for lo, hi, *_ in ranges} | {float(hi) for lo, hi, *_ in ranges})
    nodes = []
    for b in bounds:
        for lo, hi, r, R in ranges:
            if b < hi:
                nodes.append((b, r, R))
                break
        else:
            r, R = ranges[-1][2], ranges[-1][3]
            nodes.append((b, r, R))
    # кусочно-линейная интерполяция
    if w <= nodes[0][0]:
        r0, R0 = nodes[0][1], nodes[0][2]
    elif w >= nodes[-1][0]:
        r0, R0 = nodes[-1][1], nodes[-1][2]
    else:
        for i in range(len(nodes)-1):
            if nodes[i][0] <= w <= nodes[i+1][0]:
                x0, r0, R0 = nodes[i]
                x1, r1, R1 = nodes[i+1]
                t = (w - x0) / (x1 - x0) if x1 != x0 else 0.0
                r0 = r0 + t*(r1 - r0)
                R0 = R0 + t*(R1 - R0)
                break
    from math import sqrt
    u_repr = sqrt(R0*R0 - r0*r0/2.0) / 2.8
    return u_repr

w0 = 3.966567307590398
OUT["w0"] = w0
OUT["u_repr_smooth_worker"] = 0.155674
OUT["u_repr_smooth_manual"] = manual_smooth_u_repr(w0, ag.TABLE1)
OUT["u_repr_smooth_kernel"] = ag.u_repr_from_tables(w0, ag.TABLE1, "mean2", "smooth")["u_repr"]

# smooth в точке внутри (3;5] — что даёт ядро vs ступенька
for w in (0.9, 1.0, 2.0, 3.0, 4.0, 5.0):
    sm = ag.u_repr_from_tables(w, ag.TABLE1, "mean2", "smooth")["u_repr"]
    st = ag.u_repr_from_tables(w, ag.TABLE1, "mean2", "step")["u_repr"]
    OUT[f"w{w}"] = {"smooth": sm, "step": st}

with io.open(r"C:\Temp\v16_audit_part2.json", "w", encoding="utf-8") as f:
    json.dump(OUT, f, ensure_ascii=False, indent=1, default=str)
print("OK")
