# -*- coding: utf-8 -*-
"""Финальная адверсариальная сверка F-1: независимый пересчёт + проверка.
Выход — audit_f1.json (машинно-читаемо, UTF-8)."""
import io, json, math, sys
sys.path.insert(0, r"C:\Users\Arhys\Desktop\каллибровка\ГОСТ EN 13132\v16")
import anova_gum as ag

R = {}
# то, что утверждает воркер (F-1): smooth на середине (3;5] ~0.1562,
# а НЕ «между 0.1041 и 0.1406»
mid = ag.u_repr_from_tables(4.0, ag.TABLE1, "mean2", "smooth")["u_repr"]
left_node = ag.u_repr_from_tables(3.0, ag.TABLE1, "mean2", "smooth")["u_repr"]
right_node = ag.u_repr_from_tables(5.0, ag.TABLE1, "mean2", "smooth")["u_repr"]
R["smooth_4p0"] = float(mid)
R["smooth_node_3p0"] = float(left_node)
R["smooth_node_5p0"] = float(right_node)
R["worker_claim_ok"] = bool(left_node < mid < right_node)  # 0.1406<0.1562<0.1713
R["original_tz_bound_impossible"] = not (0.1041 < mid < 0.1406)

# то, что в книге (04_Профиль): max шаг, D4, U_итог
R["kernel_repr_at_w0"] = float(ag.u_repr_from_tables(3.966567307590398, ag.TABLE1, "mean2", "smooth")["u_repr"])

with io.open(r"C:\Temp\v16_f1_final.json", "w", encoding="utf-8") as f:
    json.dump(R, f, ensure_ascii=False, indent=1)
print("OK, smooth(4.0)=%.6f node3=%.6f node5=%.6f" % (mid, left_node, right_node))