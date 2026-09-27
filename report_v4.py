#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""report_v4.py — финальный отчёт по калибровке в Word (REPORT-01).

Читает Metrology_Core_EURACHEM_v4.xlsx (cached values после сохранения Excel'ем)
и result.json (если есть), строит отчёт «Отчёт_по_калибровке.docx»:

  титул → метаданные → результат (x_pred ± U) → уравнения → таблицы ОМНК/WLS/бюджета
  → графики (калибровка+полоса, профиль U(x), Pareto, контрольная карта b1)
  → валидация (20 проверок) → вердикт → подписи.

Запуск:
  python report_v4.py [файл.xlsx] [файл_result.json]
"""
import json
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm, RGBColor

plt.rcParams["font.family"] = "Arial"
plt.rcParams["axes.unicode_minus"] = False

FILE = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else "Metrology_Core_EURACHEM_v4.xlsx")
JSON = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.join(os.path.dirname(FILE), "result.json")
OUT = os.path.join(os.path.dirname(FILE), "Отчёт_по_калибровке.docx")
TMPDIR = os.path.dirname(FILE)


def cell_ok(v):
    """Число или None (не текст-ошибка)."""
    return v if isinstance(v, (int, float)) else None


def read():
    wb = openpyxl.load_workbook(FILE, data_only=True)
    d = {}
    inp, reg, wls, dg, unc, pts, bud, val, itg, std = (
        wb[s] for s in ("Ввод", "Регрессия", "Взвешенная регрессия", "Диагностика",
                        "Неопределенность", "Неопределенность по точкам", "Бюджет",
                        "Валидация", "Итог", "Стандарты")
    )
    d["мета"] = dict(
        оператор=inp["Q4"].value, прибор=inp["Q5"].value,
        партия=inp["Q6"].value, дата=inp["Q7"].value,
        темп=inp["Q8"].value, срок=inp["Q9"].value, id=inp["Q3"].value,
    )
    d["режим"] = inp["B6"].value
    d["u_rel"] = inp["B5"].value
    d["p"] = inp["B4"].value
    d["y_obs"] = inp["B3"].value
    d["b1"] = cell_ok(reg["B7"].value)
    d["b0"] = cell_ok(reg["B8"].value)
    d["s2"] = cell_ok(reg["B6"].value)
    d["r2"] = cell_ok(reg["B12"].value)
    d["LOD"] = cell_ok(reg["B19"].value)
    d["LOQ"] = cell_ok(reg["B20"].value)
    d["b1w"] = cell_ok(wls["B122"].value)
    d["b0w"] = cell_ok(wls["B123"].value)
    d["phi"] = cell_ok(wls["B125"].value)
    d["D"] = cell_ok(wls["B133"].value)
    d["x_pred"] = cell_ok(unc["B1"].value)
    d["u_c"] = cell_ok(unc["B5"].value)
    d["nu_eff"] = cell_ok(unc["B11"].value)
    d["k"] = cell_ok(unc["B12"].value)
    d["U"] = cell_ok(unc["B13"].value)
    d["U_round"] = cell_ok(itg["B12"].value)
    d["x_round"] = cell_ok(itg["B13"].value)
    d["budget"] = [(bud.cell(r, 1).value, bud.cell(r, 4).value) for r in range(15, 22)]
    d["checks"] = [(val.cell(r, 1).value, val.cell(r, 2).value, val.cell(r, 3).value)
                   for r in range(3, 23) if val.cell(r, 1).value]
    d["verdict"] = val.cell(24, 2).value if val.cell(24, 2).value else "—"
    # калибровочные точки
    xs, ys = [], []
    for r in range(9, 109):
        x = inp.cell(r, 2).value
        y = inp.cell(r, 6).value
        if x is not None and y is not None:
            xs.append(x); ys.append(y)
    d["xs"] = np.array(xs, float)
    d["ys"] = np.array(ys, float)
    # профиль U(x)
    px, pu = [], []
    for r in range(4, 204):
        x = prof_p if False else None
    prof_s = wb["Профиль U(x)"]
    for r in range(4, 204):
        x = prof_s.cell(r, 1).value
        u = prof_s.cell(r, 6).value
        if x is not None and u is not None and isinstance(x, (int, float)) and isinstance(u, (int, float)):
            px.append(x); pu.append(u)
    d["px"] = np.array(px, float) if px else np.array([])
    d["pu"] = np.array(pu, float) if pu else np.array([])
    # MC из JSON
    d["mc"] = {}
    if os.path.exists(JSON):
        try:
            with open(JSON, encoding="utf-8") as f:
                d["mc"] = json.load(f)
        except Exception:
            d["mc"] = {}
    return d


def make_charts(d):
    """Строит PNG-графики во временной папке, возвращает список путей."""
    paths = []
    f1 = os.path.join(TMPDIR, "_c1.png")
    plt.figure(figsize=(6.5, 4))
    plt.scatter(d["xs"], d["ys"], s=40, label="точки калибровки", zorder=3)
    xx = np.linspace(d["xs"].min(), d["xs"].max(), 50)
    plt.plot(xx, d["b1"] * xx + d["b0"], "-", color="#4F81BD", label="ОМНК")
    if d["b1w"]:
        plt.plot(xx, d["b1w"] * xx + d["b0w"], "--", color="#C00000", label="WLS")
    plt.xlabel("x (концентрация)")
    plt.ylabel("y (отклик)")
    plt.title("Калибровочная прямая (ОМНК / WLS)")
    plt.legend()
    plt.grid(alpha=0.3)
    plt.tight_layout(); plt.savefig(f1, dpi=150); plt.close()
    paths.append(f1)

    f2 = os.path.join(TMPDIR, "_c2.png")
    plt.figure(figsize=(6.5, 4))
    if len(d["px"]):
        plt.plot(d["px"], d["pu"], color="#C00000", lw=2)
        plt.axvline(d["x_pred"], color="gray", ls=":", label="x_pred")
    plt.xlabel("x")
    plt.ylabel("U(x)")
    plt.title("Профиль расширенной неопределённости U(x)")
    plt.grid(alpha=0.3)
    plt.tight_layout(); plt.savefig(f2, dpi=150); plt.close()
    paths.append(f2)

    f3 = os.path.join(TMPDIR, "_c3.png")
    plt.figure(figsize=(6.5, 4))
    labels = [b[0] for b in d["budget"] if b[0] and b[1] is not None]
    vals = [abs(b[1]) ** 2 for b in d["budget"] if b[0] and b[1] is not None]
    if vals:
        plt.barh(labels, vals, color="#4F81BD")
        plt.xlabel("квадрат вклада (отн. единицы)")
        plt.title("Pareto: вклады в неопределённость (квадраты)")
        plt.gca().invert_yaxis()
        plt.tight_layout(); plt.savefig(f3, dpi=150); plt.close()
        paths.append(f3)
    return paths


def build_report(d):
    doc = Document()
    # стиль по умолчанию
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(11)

    # Титул
    h = doc.add_heading("ОТЧЁТ ПО КАЛИБРОВКЕ", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Metrology Core EURACHEM — расчёт неопределённости обратного предсказания")

    # Метаданные
    doc.add_heading("1. Метаданные калибровки", level=1)
    tbl = doc.add_table(rows=0, cols=2)
    tbl.style = "Table Grid"
    for k, v in d["мета"].items():
        cells = tbl.add_row().cells
        cells[0].text = k
        cells[1].text = str(v) if v is not None else "—"

    # Результат
    doc.add_heading("2. Результат измерения", level=1)
    if d["x_pred"] is not None:
        r = doc.add_paragraph()
        r.add_run(f"x_pred = {d['x_round']} ± {d['U_round']}  ").bold = True
        r.add_run(f"(k={d['k']:.3f}, ν_eff={d['nu_eff']:.1f}, P=0.95)")
        p = doc.add_paragraph(
            f"U = {d['U']:.6g} (округлено до {d['U_round']} по GUM 7.2.6); "
            f"u_c = {d['u_c']:.6g}; LOD = {d['LOD']:.4g}; LOQ = {d['LOQ']:.4g}"
        )
    else:
        doc.add_paragraph("Результат недоступен (ошибка расчёта).")

    # Уравнения
    doc.add_heading("3. Модель", level=1)
    doc.add_paragraph("Прямая модель:  y = b1·x + b0")
    doc.add_paragraph(f"  b1 = {d['b1']:.9g}, b0 = {d['b0']:.9g}, R² = {d['r2']:.6f}")
    if d["b1w"]:
        doc.add_paragraph(f"WLS: b1_w = {d['b1w']:.9g}, b0_w = {d['b0w']:.9g}, φ_w = {d['phi']:.3e}")
    doc.add_paragraph("Обратная модель:  x_pred = (y_obs − b0)/b1")

    # Бюджет
    doc.add_heading("4. Бюджет неопределённости", level=1)
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    tbl.rows[0].cells[0].text = "Источник"
    tbl.rows[0].cells[1].text = "Вклад (отн.)"
    for name, v in d["budget"]:
        if name:
            cells = tbl.add_row().cells
            cells[0].text = str(name)
            cells[1].text = f"{v:.4g}" if v is not None else "—"

    # Графики
    doc.add_heading("5. Графики", level=1)
    imgs = make_charts(d)
    for p in imgs:
        doc.add_picture(p, width=Cm(15))
        doc.add_paragraph("")
        os.remove(p)

    # Валидация
    doc.add_heading("6. Автоматическая валидация", level=1)
    tbl = doc.add_table(rows=1, cols=3)
    tbl.style = "Table Grid"
    hdr = tbl.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "Проверка", "Формула", "Статус"
    for name, formula, status in d["checks"]:
        cells = tbl.add_row().cells
        cells[0].text = str(name)
        cells[1].text = str(formula)[:60]
        cells[2].text = str(status)

    # Вердикт
    doc.add_heading("7. Вердикт", level=1)
    doc.add_paragraph(f"КАЛИБРОВКА: {d['verdict']}")

    # MC
    if d["mc"]:
        doc.add_heading("8. Monte Carlo (JCGM 101)", level=1)
        mc = d["mc"]
        doc.add_paragraph(
            f"x_pred(MC) = {mc.get('x_pred', '—')}, u_MC = {mc.get('u_c', '—')}, "
            f"k = {mc.get('k', '—')}, U(MC) = {mc.get('U', '—')}"
        )
        doc.add_paragraph(f"Статус: {mc.get('status', '—')}")

    # Подписи
    doc.add_heading("9. Утверждение", level=1)
    tbl = doc.add_table(rows=3, cols=3)
    tbl.style = "Table Grid"
    for i, role in enumerate(("Выполнил", "Проверил", "Утвердил")):
        cells = tbl.rows[i].cells
        cells[0].text = role
        cells[1].text = ""
        cells[2].text = "Дата: ______"

    doc.save(OUT)
    print(f"Отчёт сохранён: {OUT}")
    return OUT


def main():
    d = read()
    print(f"файл: {FILE}, вердикт: {d['verdict']}, x_pred={d['x_pred']}")
    build_report(d)
    return 0


if __name__ == "__main__":
    sys.exit(main())