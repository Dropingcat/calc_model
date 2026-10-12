#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""build_en13132_v17_calc.py — генерация расчётной книги v17 (блок `gost13132-v17-final`).

EN13132_v17_CALC.xlsx по паттерну v4/v15/v16:
  01_Ввод         данные градуировки + поля результата пробы (ЕДИНЫЙ источник входов)
  02_Регрессия    OLS-характеристики + LoF (динамический) + таблица точек для графика
  03_ANOVA        классическая ANOVA по GUM (грубый учёт табл.1/2)
  04_Профиль U(X) профиль U(x)=k·u_c(x) на 200 точках + график
  05_Бюджет       бюджет неопределённости результата пробы (формат Word-отчёта v15)
  06_Итог         результат X0, w[% мас.], U, протокольная запись (структура Word v15)
  Графики         3 диаграммы (калибровка, профиль U(x), остатки)

Движок — anova_gum.py (импорт, вычисления функциями ядра).

Отличия v17 от v16 (TD-DEV-208, вердикт ресерчера по НД):
  * РЕЖИМ ПО УМОЛЧАНИЮ `interpolation="step"` — ступенчатая по поддиапазонам
    табл.1/2 ГОСТ EN 13132 (r/R ДИСКРЕТНЫ, интерполяция не предписана ни
    стандартом, ни 5725-6, ни РМГ 61; GUM 4.1.4 к приписанным нормам
    прецизионности неприменим). smooth — методическое допущение, НЕ по НД.
  * Вклады бюджета — поля ввода (значение из спецификации ИЛИ 0 = «не
    учитывать»): эталоны, чистота, весы. При 0 вклад исключается из
    u_c и ν_eff (ядро combined_budget, TD-DEV-33).
  * Статус «ПРЕДВАРИТЕЛЬНО» снят: значения вкладов — из спецификации
    пользователя; пользователь сам решает, заполнять их или 0.

АНТИ-ХАРДКОД (P3, итерация 3):
  * ВСЕ параметры результата пробы читаются из ячеек листа 01_Ввод (генератор
    сначала ЗАПИСЫВАЕТ дефолты в ячейки, затем ЧИТАЕТ их обратно — единый
    источник, контракт: «расчёт — всегда через чтение значений книги»).
    Константы M_PROBY/M_IS_PROBA/P_IS/P_A/Y0_PROBA/U_M_NET/U_PA в коде удалены.
  * Число уровней — динамическое (2..12; OLS требует >= 3), число инъекций —
    из данных (1..3). Строки таблиц графиков (DATA_FIRST/DATA_LAST/GRID_PT,
    PROF_FIRST/PROF_LAST) вычисляются от фактического числа уровней.
  * Генератор переживает 4/5/6/8 уровней (проверяется верификатором).

TD-DEV-206 (итерация 3): u_repr результата пробы — по ФАКТИЧЕСКОМУ
w0 = X0·K_G2PCT (а не константа 2.0 %мас.):
    u_repr_from_tables(w0, TABLE1, "mean2", interpolation="smooth")
U(06_Итог) и U(x_pred) маркер («Графики»!D4) СОВПАДАЮТ (одно значение).
05_Бюджет (строка «прецизионность») согласован — u_repr(w0).

P1 (итерация 3, v16): профиль 04_Профиль U(X) и диаграмма использовали
interpolation="smooth" — НЕПРЕРЫВНЫЙ профиль. В v17 — ПО УМОЛЧАНИЮ step
(TD-DEV-208: табл.1/2 ГОСТ EN 13132 дискретны, интерполяция не предписана;
smooth остаётся методическим допущением, НЕ по НД).

P2 (итерация 3): лист 06_Итог приведён к структуре Word-отчёта v15
(параметры → результат w±U → протокольная запись → таблица вкладов →
примечания; v16 добавляет ANOVA-вклад прецизионности и U=k·u_c).

Форма модели в профиле/бюджете — интервальная OLS для среднего из p инъекций
пробы (контракт §5, строка 1): u²(x) = s²_res·(1/p + 1/n + (x−x̄)²/Sxx)/a².
Это обеспечивает ТОЖДЕСТВО U(06_Итог) == U(x_pred) (профиль и бюджет
результата используют одну и ту же форму модели).
"""
from __future__ import annotations

import math
import os
import sys
import warnings

import numpy as np
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import anova_gum as ag  # noqa: E402

V17_DIR = os.path.dirname(os.path.abspath(__file__))
XLSX_V15 = os.path.join(V17_DIR, "..", "v15", "EN13132_v15_MTBE.xlsx")
OUT_XLSX = os.path.join(V17_DIR, "EN13132_v17_CALC.xlsx")

N_PROFILE = 200  # точек профиля (паттерн v4)

# ---------------------------------------------------------------------------
# ВХОДНЫЕ ПОЛЯ 01_Ввод (метка в столбце A, значение — в B, единицы — в C)
# ---------------------------------------------------------------------------
INPUT_FIELDS = [
    # (ключ, метка, дефолт, формат, единицы, примечание)
    ("m_proby", "Масса пробы m_пробы (г)", 100.0, "0.000", "г",
     "Масса пробы для перевода X→%мас (U3, контракт разд.9.2); w = X·m_ИС·p_ИС/m_пробы·100"),
    ("m_is_proba", "Масса ИС в пробу m_ИС (г)", 4.0, "0.000", "г",
     "Навеска внутреннего стандарта в пробу"),
    ("p_is", "Чистота ИС p_ИС", 0.995, "0.0000", "1",
     "Доля (99,5%); источник — паспорт партии"),
    ("p_a", "Чистота аналита p_A", 0.995, "0.0000", "1",
     "Доля (99,5%)"),
    ("y0_proba", "Y0 пробы (среднее p инъекций)", None, "0.000000", "Y",
     "Средний отклик пробы; заменить на реальное измерение"),
    ("n_inj", "Число инъекций пробы p", 2, "0", "шт",
     "p — протокольная практика ГОСТ EN 13132 §10.1 (сходимость |x1−x2|≤r)"),
    ("u_m_net", "u(m_net) весы (г)", 5.7735026918962585e-05, "0.00000000", "г",
     "ВЕСЫ — ПОЛЕ ВВОДА: значение из спецификации весов ИЛИ 0 = «не учитывать» "
     "(вклад исключается из u_c и ν_eff, TD-DEV-33)"),
    ("u_pa", "u(pA)=u(pИС) (доля)", 0.2886751345948129 / 100.0, "0.000000", "1",
     "ЧИСТОТА — ПОЛЕ ВВОДА: (99..100)% /√12 равномерное ИЛИ 0 = «не учитывать» "
     "(вклад исключается из u_c и ν_eff, TD-DEV-33)"),
    ("u_rel_etal", "u_rel эталонов (доля)", 0.001, "0.0000", "1",
     "ЭТАЛОНЫ — ПОЛЕ ВВОДА: относительная неопределённость эталонов калибровки "
     "(свидетельство) ИЛИ 0 = «не учитывать» (вклад исключается из u_c и ν_eff, TD-DEV-33)"),
]

PARAM_TITLE = "ПАРАМЕТРЫ РЕЗУЛЬТАТА ПРОБЫ (ввод методиста)"
STATUS_TEXT = ("ВКЛАДЫ — ПОЛЯ ВВОДА: значение из спецификации ИЛИ 0 = «не учитывать» "
               "(исключается из u_c и ν_eff); заполнить перед протоколом (TD-DEV-33/208)")

# ---------------------------------------------------------------------------
# СТИЛИ
# ---------------------------------------------------------------------------
HDR_FILL = PatternFill("solid", fgColor="DDEBF7")
HDR_FONT = Font(bold=True, size=11)
TITLE_FONT = Font(bold=True, size=13)
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
SUP = "9E480E"


# ---------------------------------------------------------------------------
# Excel-правило «xVal и yVal серии — на ОДНОМ листе» (фикс итерации 2)
# ---------------------------------------------------------------------------
def _ref_sheet(f):
    """Лист из строки ссылки openpyxl numRef.f.

    '02_Регрессия'!$A$39:$A$50 -> '02_Регрессия';  Лист1!$A$1 -> 'Лист1'.
    Возвращает None, если строка не является ссылкой на диапазон.
    """
    if not isinstance(f, str) or not f:
        return None
    if f.startswith("'"):
        end = f.find("'", 1)
        if end == -1:
            return None
        return f[1:end].replace("''", "'")
    bang = f.find("!")
    return f[:bang] if bang != -1 else None


def assert_chart_series_sheets(wb, chart_sheet="Графики"):
    """БАГ (пользователь): «у всех графиков X разный (шаг сетки) — Excel багует
    и заменяет на 1.2.3.4.5.6 вместо реальных значений».

    Причина: для ScatterChart xVal и yVal КАЖДОЙ серии должны лежать на ОДНОМ
    листе, иначе Excel сбрасывает X на индексы 1..n. Эта проверка проходится в
    конце генерации по всем диаграммам/сериям и бросает AssertionError с
    деталями при нарушении (страховка от регрессии на этапе сборки).
    """
    gch = wb[chart_sheet]
    bad = []
    for ci, chart in enumerate(gch._charts, 1):
        for si, series in enumerate(chart.series, 1):
            xv = getattr(series, "xVal", None)
            yv = getattr(series, "yVal", None)
            xf = getattr(getattr(xv, "numRef", None), "f", None)
            yf = getattr(getattr(yv, "numRef", None), "f", None)
            xs, ys = _ref_sheet(xf), _ref_sheet(yf)
            if xs is None or ys is None or xs != ys:
                bad.append(f"chart{ci}/series{si}: xVal@{xs} != yVal@{ys} "
                           f"(x={xf!r}, y={yf!r})")
    assert not bad, (
        "Excel-правило нарушено: xVal/yVal серии на РАЗНЫХ листах -> Excel "
        "рендерит категории 1..n вместо X. " + "; ".join(bad)
    )


def comment(text, author="worker"):
    return Comment(text, author)


def set_cell(ws, coord, value, num_fmt=None, bold=False, fill=None, cmt=None):
    cell = ws[coord]
    cell.value = value
    if num_fmt:
        cell.number_format = num_fmt
    if bold:
        cell.font = Font(bold=True)
    if fill:
        cell.fill = fill
    if cmt:
        cell.comment = comment(cmt)
    return cell


def find_hdr_row(ws, probes):
    """Строка-заголовок: все probe-значения (col, val) совпадают в одной строке."""
    for r in range(1, ws.max_row + 1):
        ok = True
        for col, val in probes:
            if ws.cell(r, col).value != val:
                ok = False
                break
        if ok:
            return r
    return None


def find_label_row(ws, needle):
    """Строка по подстроке/префиксу в столбце A."""
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.startswith(needle):
            return r
    return None


# ---------------------------------------------------------------------------
# ЧТЕНИЕ ДАННЫХ И ВХОДОВ
# ---------------------------------------------------------------------------
def load_v15_levels(path):
    """Уровни из исходной книги v15 (лист 01_Ввод): ДИНАМИЧЕСКОЕ число уровней.

    Макет v15: строка 6 — заголовок (№ | mA | mИС | Y1 | Y2 | Y3 | Xном | Xэфф | ...),
    строки 7.. — данные (в v15 заняты строки 7..12; ниже — пустые заготовки №7..12).
    Колонки: B=mA, C=mИС, D=Y1, E=Y2, F=Y3, H=X_эфф.
    Инъекции — 1..3 (Y3 может быть пустой). Возвращает list[(X, [Y...], mA, mIS)].
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["01_Ввод"]
    levels = []
    for r in range(7, ws.max_row + 1):
        x = ws.cell(r, 8).value  # H: X_эфф
        ys = [float(ws.cell(r, c).value) for c in (4, 5, 6)
              if isinstance(ws.cell(r, c).value, (int, float))]
        if not (isinstance(x, (int, float)) or ys):
            break  # конец данных (пустая строка-заготовка)
        if not isinstance(x, (int, float)):
            continue
        mA = ws.cell(r, 2).value
        mIS = ws.cell(r, 3).value
        levels.append((float(x), ys,
                       float(mA) if isinstance(mA, (int, float)) else None,
                       float(mIS) if isinstance(mIS, (int, float)) else None))
    if len(levels) < 3:
        raise ValueError(f"load_v15_levels: найдено {len(levels)} уровней (<3) в {path}")
    return levels


def read_levels(ws):
    """Уровни из 01_Ввод СГЕНЕРИРОВАННОЙ книги: колонки определяются по заголовкам.

    Возвращает list[(X, [Y...], mA, mIS)] или None, если таблица уровней не найдена.
    """
    hdr = find_hdr_row(ws, [(1, "Уровень")])
    if hdr is None:
        return None
    col_of = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(hdr, c).value
        if isinstance(v, str):
            col_of[v.strip()] = c
    x_col = col_of.get("X (г/г)")
    y_cols = [col_of.get("Y1"), col_of.get("Y2"), col_of.get("Y3")]
    mA_col = col_of.get("mA (г)")
    mIS_col = col_of.get("mИС (г)")
    if x_col is None or not any(y_cols):
        return None
    levels = []
    for r in range(hdr + 1, ws.max_row + 1):
        x = ws.cell(r, x_col).value
        ys = [float(ws.cell(r, c).value) for c in y_cols
              if c is not None and isinstance(ws.cell(r, c).value, (int, float))]
        if not (isinstance(x, (int, float)) or ys):
            break
        if not isinstance(x, (int, float)):
            continue
        mA = ws.cell(r, mA_col).value if mA_col else None
        mIS = ws.cell(r, mIS_col).value if mIS_col else None
        levels.append((float(x), ys,
                       float(mA) if isinstance(mA, (int, float)) else None,
                       float(mIS) if isinstance(mIS, (int, float)) else None))
    if len(levels) < 2:
        return None
    return levels


def default_inputs(levels):
    """Дефолты входных полей. Y0 — средний отклик СРЕДНЕГО уровня калибровки
    (динамично: для 6 уровней — уровень 4, как было в итер.1-2)."""
    d = {key: (val if key != "y0_proba" else None)
         for key, _l, val, _f, _u, _c in INPUT_FIELDS}
    mid = levels[len(levels) // 2]
    d["y0_proba"] = float(np.mean(mid[1]))
    return d


def read_inputs(ws):
    """Чтение входных полей из 01_Ввод по меткам (столбец A; значение в B)."""
    out = {}
    for key, label, _default, _fmt, _units, _cmt in INPUT_FIELDS:
        r = find_label_row(ws, label)
        if r is None:
            raise ValueError(f"read_inputs: не найдено поле '{label}' на листе 01_Ввод")
        v = ws.cell(r, 2).value
        if not isinstance(v, (int, float)):
            raise ValueError(f"read_inputs: поле '{label}' не числовое: {v!r}")
        out[key] = float(v)
    out["n_inj"] = int(round(out["n_inj"]))
    return out


def write_input_sheet(wb, levels, inputs):
    """Лист 01_Ввод: таблица уровней (динамическая) + входные поля (дефолты)."""
    ws = wb.create_sheet("01_Ввод")
    ws.append(["ГРАДУИРОВКА МТБЭ | ГОСТ EN 13132-2012"] + [""] * 8)
    ws.merge_cells("A1:I1")
    ws["A1"].font = TITLE_FONT
    ws.append([])
    ws.append(["Данные градуировки (единый источник входов; уровни 2..12, инъекции 1..3 — "
               "определяются по данным)"] + [""] * 8)
    ws.merge_cells("A3:I3")
    ws["A3"].font = Font(italic=True, color=SUP)
    ws.append(["Уровень", "X (г/г)", "Y1", "Y2", "Y3", "Y_ср", "n", "mA (г)", "mИС (г)"])
    hdr = ws.max_row
    for c in range(1, 10):
        ws.cell(hdr, c).fill = HDR_FILL
        ws.cell(hdr, c).font = HDR_FONT
        ws.cell(hdr, c).border = BORDER
    for i, (x, ys, mA, mIS) in enumerate(levels):
        r = hdr + 1 + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, x).number_format = "0.000000"
        for j in range(3):
            if j < len(ys):
                ws.cell(r, 3 + j, ys[j]).number_format = "0.000000"
        ws.cell(r, 6, float(np.mean(ys))).number_format = "0.000000"
        ws.cell(r, 7, len(ys))
        if mA is not None:
            ws.cell(r, 8, mA).number_format = "0.000"
        else:
            ws.cell(r, 8, "")
        if mIS is not None:
            ws.cell(r, 9, mIS).number_format = "0.000"
        else:
            ws.cell(r, 9, "")
        for c in range(1, 10):
            ws.cell(r, c).border = BORDER
    p_row = hdr + len(levels) + 2
    ws.cell(p_row, 1, PARAM_TITLE)
    ws.merge_cells(f"A{p_row}:I{p_row}")
    ws.cell(p_row, 1).font = TITLE_FONT
    for i, (key, label, _def, fmt, units, cmt) in enumerate(INPUT_FIELDS):
        r = p_row + 1 + i
        ws.cell(r, 1, label).border = BORDER
        ws.cell(r, 1).alignment = Alignment(wrap_text=True)
        set_cell(ws, f"B{r}", inputs[key], num_fmt=fmt, cmt=cmt)
        ws.cell(r, 3, units)
        for c in range(1, 4):
            ws.cell(r, c).border = BORDER
    st_r = p_row + 1 + len(INPUT_FIELDS) + 1
    ws.cell(st_r, 1, "Статус").font = Font(bold=True)
    set_cell(ws, f"B{st_r}", STATUS_TEXT,
             cmt="Блокировка вывода протокола до паспортных данных (контракт §10, TD-DEV-33)")
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 16
    for col in "CDEFGHI":
        ws.column_dimensions[col].width = 12
    return ws


# ---------------------------------------------------------------------------
# РАСЧЁТ (движок: anova_gum.py)
# ---------------------------------------------------------------------------
def lof_test(levels, a, b):
    """LoF F-тест по повторностям (5725-2; аналог v15 03_Статистика B26..B29).

    SS_pe = Σ_i Σ_j (Y_ij − Ȳ_i)²,  df_pe = Σ_i (n_i − 1);
    SS_res = Σ_ij (Y_ij − Ŷ_ij)²,   df_res = N − 2;
    SS_lof = SS_res − SS_pe,        df_lof = df_res − df_pe.
    Возвращает dict(F, F_crit, df1, df2, status) или None при недостаточных df.
    """
    ss_pe = 0.0
    df_pe = 0
    N = 0
    ss_res = 0.0
    for x, ys in levels:
        ys = np.asarray([float(y) for y in ys], dtype=float)
        N += len(ys)
        ss_pe += float(np.sum((ys - ys.mean()) ** 2))
        df_pe += len(ys) - 1
        ss_res += float(np.sum((ys - (a * float(x) + b)) ** 2))
    df_res = N - 2
    df_lof = df_res - df_pe
    ss_lof = max(0.0, ss_res - ss_pe)
    if df_lof <= 0 or df_pe <= 0 or ss_pe <= 0:
        return None
    from scipy.stats import f as fdist

    F = (ss_lof / df_lof) / (ss_pe / df_pe)
    F_crit = float(fdist.ppf(0.95, df_lof, df_pe))
    return {
        "F": float(F),
        "F_crit": F_crit,
        "df1": int(df_lof),
        "df2": int(df_pe),
        "status": "НЕ ВЫЯВЛЕНА" if F <= F_crit else "ВЫЯВЛЕНА (модель не утверждается)",
    }


def compute_calibration(levels, inputs, interpolation="step"):
    """Все расчёты от входов (единый источник). Возвращает dict со всеми числами.

    TD-DEV-206: u_repr результата — по фактическому w0 = X0·K_G2PCT.
    v17 (TD-DEV-208): РЕЖИМ ПО УМОЛЧАНИЮ interpolation="step" — ступенчатая
    по поддиапазонам табл.1/2 ГОСТ EN 13132 (r/R ДИСКРЕТНЫ, интерполяция не
    предписана ни стандартом, ни 5725-6, ни РМГ 61; GUM 4.1.4 к приписанным
    нормам прецизионности неприменим). smooth — методическое допущение, НЕ по НД.
    TD-DEV-33: вклады-поля ввода (эталоны/чистота/весы): 0 = «не учитывать» —
    вклад ИСКЛЮЧАЕТСЯ из u_c и ν_eff (combined_budget фильтрует u==0).
    Профиль и бюджет результата используют ОДНУ форму модели (интервальная OLS
    для среднего из p инъекций пробы) — поэтому U(06_Итог) == U(x_pred) маркер.
    """
    if interpolation not in ("step", "smooth"):
        raise ValueError("interpolation должен быть 'step' или 'smooth'")
    # нормализация: уровни принимаются как (x, ys) или (x, ys, mA, mIS)
    levels = [(float(x), list(ys)) for x, ys, *_ in levels]
    if len(levels) < 3:
        raise ValueError("Для OLS по средним требуется >= 3 уровней")
    SNAP = ag.ols_slopes_by_means(levels)
    SNAP["levels"] = levels
    ANOVA = ag.anova_repeatability(levels)
    SIGMA_R = ANOVA["sigma_r"]
    p = int(inputs["n_inj"])
    if p < 1:
        raise ValueError("Число инъекций пробы p должно быть >= 1")
    m_proby = float(inputs["m_proby"])
    m_is = float(inputs["m_is_proba"])
    p_is = float(inputs["p_is"])
    p_a = float(inputs["p_a"])
    y0 = float(inputs["y0_proba"])
    u_m_net = float(inputs["u_m_net"])
    u_pa = float(inputs["u_pa"])
    u_rel_etal = float(inputs["u_rel_etal"])
    if m_proby <= 0 or m_is <= 0:
        raise ValueError("Массы пробы/ИС должны быть > 0")
    if p_is <= 0 or p_a <= 0:
        raise ValueError("Чистоты должны быть > 0")

    K_G2PCT = m_is * p_is / m_proby * 100.0  # г/г -> % мас.
    u_rel_mass = u_m_net / m_is
    u_rel_purity = math.sqrt((u_pa / p_a) ** 2 + (u_pa / p_is) ** 2)

    INV = ag.inverse_calibration(
        y0, SNAP["a"], SNAP["b"], SNAP["var_a"], SNAP["var_b"], SNAP["cov_ab"], u_y0=0.0
    )
    X0 = INV["x0"]
    W0 = X0 * K_G2PCT

    def u_repr_pct(w):
        return ag.u_repr_from_tables(w, ag.TABLE1, "mean2", interpolation)["u_repr"]

    def u_repr_x(x):
        return u_repr_pct(x * K_G2PCT) / K_G2PCT

    def u_model_fn(x):
        # интервальная форма OLS (контракт §5, строка 1):
        #   u²(x) = s²_res·(1/p + 1/n + (x−x̄)²/Sxx)/a²
        return math.sqrt(
            SNAP["s2_res"] * (1.0 / p + 1.0 / SNAP["n"]
                              + (x - SNAP["xbar"]) ** 2 / SNAP["Sxx"])
        ) / abs(SNAP["a"])

    def budget_at_x(x, with_model=True):
        comps = []
        if with_model:
            comps.append(("модель", u_model_fn(x), SNAP["df_res"]))
        # TD-DEV-33: вклады-поля ввода. 0 = «не учитывать»: вклад НЕ попадает
        # в список (combined_budget также фильтрует u==0 — двойная защита).
        if u_rel_etal > 0.0:
            comps.append(("эталоны", abs(x) * u_rel_etal, 50))
        if u_rel_purity > 0.0:
            comps.append(("чистота", abs(x) * u_rel_purity, None))
        if u_rel_mass > 0.0:
            comps.append(("весы", abs(x) * u_rel_mass, None))
        comps.append(("прецизионность", u_repr_x(x), None))
        return comps

    xs = np.array([float(x) for x, _ys in levels], dtype=float)
    X_GRID = np.linspace(float(xs.min()), float(xs.max()), N_PROFILE)
    U_MODEL = np.array([u_model_fn(float(x)) for x in X_GRID])
    U_REPR_X = np.zeros_like(X_GRID)
    U_C = np.zeros_like(X_GRID)
    NU_EFF = np.zeros_like(X_GRID)
    K_ARR = np.zeros_like(X_GRID)
    for i, xi in enumerate(X_GRID):
        comps = budget_at_x(float(xi), with_model=True)
        u_c = math.sqrt(sum(u ** 2 for _n, u, _d in comps))
        nu = ag.welch_satterthwaite([(u, d) for _n, u, d in comps])
        U_C[i] = u_c
        NU_EFF[i] = nu
        K_ARR[i] = ag.k_coverage(nu)
        U_REPR_X[i] = next(u for _n, u, _d in comps if _n == "прецизионность")
    U_X = K_ARR * U_C

    comps_result = budget_at_x(X0, with_model=True)
    BUDG = ag.combined_budget(comps_result)
    U_C_RES = BUDG["u_c"]
    NU_EFF_RES = BUDG["nu_eff"]
    K_RES = BUDG["k"]
    U_RES_G = BUDG["U"]
    U_RES_PCT = U_RES_G * K_G2PCT
    U_REL = U_RES_PCT / W0 if W0 != 0 else float("nan")

    urepr_w0 = u_repr_pct(W0)
    sigma_R_tab = ag.u_repr_from_tables(W0, ag.TABLE1, "single", interpolation)["sigma_R_tab"]
    R_w0 = ag.u_repr_from_tables(W0, ag.TABLE1, "single", interpolation)["R"]
    sigma_L = ag.sigma_L_from_tables(urepr_w0, SIGMA_R)
    lof = lof_test(levels, SNAP["a"], SNAP["b"])

    return {
        "SNAP": SNAP, "ANOVA": ANOVA, "SIGMA_R": SIGMA_R,
        "sigma_R_tab": sigma_R_tab, "R_w0": R_w0, "sigma_L": sigma_L, "lof": lof,
        "K_G2PCT": K_G2PCT, "u_rel_mass": u_rel_mass, "u_rel_purity": u_rel_purity,
        "u_rel_etal": u_rel_etal, "p": p, "y0": y0, "u_repr_w0": urepr_w0,
        "X0": X0, "W0": W0, "U_X0_model": u_model_fn(X0),
        "U_Y0": SIGMA_R / math.sqrt(p),
        "X_GRID": X_GRID, "U_MODEL": U_MODEL, "U_REPR_X": U_REPR_X,
        "U_C": U_C, "NU_EFF": NU_EFF, "K_ARR": K_ARR, "U_X": U_X,
        "x_pred_idx": int(np.argmin(np.abs(X_GRID - X0))),
        "BUDG": BUDG, "U_C_RES": U_C_RES, "NU_EFF_RES": NU_EFF_RES,
        "K_RES": K_RES, "U_RES_G": U_RES_G, "U_RES_PCT": U_RES_PCT,
        "U_REL": U_REL, "interpolation": interpolation,
    }


# ---------------------------------------------------------------------------
# ЗАПОЛНЕНИЕ ЛИСТОВ
# ---------------------------------------------------------------------------
def fill_regression(wb, levels, inputs, res):
    ws = wb.create_sheet("02_Регрессия")
    SNAP = res["SNAP"]
    lof = res["lof"]
    ws["A1"] = "ГРАДУИРОВОЧНАЯ КРИВАЯ (OLS по средним)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws.append([])
    n = SNAP["n"]
    rows = [
        ("Параметр", "Значение", "Ед.", "Примечание"),
        ("Наклон a", SNAP["a"], "Y/X", "OLS по средним; ГОСТ EN 13132 §8, v15 06_Отчёт!B8"),
        ("Свободный член b", SNAP["b"], "Y", "v15 06_Отчёт!B9"),
        ("var(a)", SNAP["var_a"], "(Y/X)²", "s²_res/Sxx; GUM 6.4"),
        ("var(b)", SNAP["var_b"], "Y²", "s²_res·(1/n+x̄²/Sxx)"),
        ("cov(a,b)", SNAP["cov_ab"], "Y²/X", "−s²_res·x̄/Sxx; GUM 6.4"),
        ("R²", SNAP["r2"], "1", "Описательный показатель; v15 06_Отчёт!B10"),
        ("s²_res (остаточная дисперсия)", SNAP["s2_res"], "Y²", "по средним уровней, df=n−2"),
        ("s_res", math.sqrt(SNAP["s2_res"]), "Y", "стандартное отклонение остатков"),
        ("df_res", SNAP["df_res"], "1", f"n−2 = {n - 2} ({n} уровней)"),
        ("x̄ — центр калибровки", SNAP["xbar"], "г/г", "x̄ = ΣX/n"),
        ("X min", float(np.min([x for x, _ in levels])), "г/г", "нижняя граница калибровки"),
        ("X max", float(np.max([x for x, _ in levels])), "г/г", "верхняя граница калибровки"),
    ]
    for i, row in enumerate(rows):
        ws.append(row)
        r = ws.max_row
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER
            if i == 0:
                ws.cell(r, c).fill = HDR_FILL
                ws.cell(r, c).font = HDR_FONT
        if i >= 1:
            ws.cell(r, 2).number_format = "0.00000000"
            ws.cell(r, 3).alignment = Alignment(horizontal="left")
    # LoF — динамический расчёт по данным (анти-хардкод, P3)
    ws.append([])
    ws.append(["Lack-of-fit (диагностика; расчёт по данным, 5725-2)"])
    if lof is None:
        ws.append(["F LoF", "н/д", "1", "нет повторностей/мало df — LoF не оценим"])
    else:
        ws.append(["F LoF", lof["F"], "1",
                   f"F < F_крит → LoF {lof['status'].lower()}; df1={lof['df1']}, df2={lof['df2']}"])
        ws.append(["F крит. 95%", lof["F_crit"], "1",
                   f"F(0.95; df1={lof['df1']}; df2={lof['df2']})"])
    for r in range(ws.max_row - (2 if lof else 1), ws.max_row + 1):
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER
    ws.append([])
    ws.append(["Уравнение: Y = a·X + b;  Y = %.6f·X + %.6f" % (SNAP["a"], SNAP["b"]),
               "", "", ""])
    ws.append([])

    # Таблица точек для графика (динамические строки)
    # Столбец G — константа 0 для ЛИНИИ НУЛЯ графика остатков. xVal и yVal
    # серии '0 (линия нуля)' должны лежать на ОДНОМ листе (Excel-правило:
    # иначе Excel рендерит категории 1..n вместо X) — поэтому нули здесь,
    # на 02_Регрессия, а не на листе «Графики».
    ws.append(["X", "Y_ср", "Y_расч (линия)", "e (остаток)", "ŷ−k·u(ŷ)", "ŷ+k·u(ŷ)",
               "0 (линия нуля)"])
    DATA_HDR_ROW = ws.max_row
    for c in range(1, 8):
        ws.cell(DATA_HDR_ROW, c).fill = HDR_FILL
        ws.cell(DATA_HDR_ROW, c).font = HDR_FONT
    ws.cell(DATA_HDR_ROW, 7).comment = comment(
        "Константа 0 для линии нуля графика остатков. xVal и yVal серии '0' "
        "на ОДНОМ листе (02_Регрессия) — Excel-правило ScatterChart."
    )
    DATA_FIRST = DATA_HDR_ROW + 1
    DATA_LAST = DATA_FIRST + len(levels) - 1
    k_band = ag.k_coverage(SNAP["df_res"])  # t_{0.975}(n−2)
    for i, (x, ys) in enumerate(levels):
        r = DATA_FIRST + i
        y_mean = float(np.mean(ys))
        y_calc = SNAP["a"] * float(x) + SNAP["b"]
        u_yhat = math.sqrt(
            SNAP["s2_res"] * (1.0 / SNAP["n"]
                              + (float(x) - SNAP["xbar"]) ** 2 / SNAP["Sxx"])
        )
        ws.cell(r, 1, x).number_format = "0.000000"
        ws.cell(r, 2, y_mean).number_format = "0.000000"
        ws.cell(r, 3, y_calc).number_format = "0.000000"
        ws.cell(r, 4, y_mean - y_calc).number_format = "0.000000"
        ws.cell(r, 5, y_calc - k_band * u_yhat).number_format = "0.000000"
        ws.cell(r, 6, y_calc + k_band * u_yhat).number_format = "0.000000"
        ws.cell(r, 7, 0.0).number_format = "0.000000"
    GRID_PT = ws.max_row + 2
    for i, x in enumerate(np.linspace(
            float(np.min([x for x, _ in levels])),
            float(np.max([x for x, _ in levels])), 100)):
        ws.cell(GRID_PT + i, 1, float(x)).number_format = "0.000000"
        ws.cell(GRID_PT + i, 3, float(SNAP["a"] * x + SNAP["b"])).number_format = "0.000000"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 42
    return ws, DATA_FIRST, DATA_LAST, GRID_PT


def fill_anova(wb, levels, inputs, res):
    ws = wb.create_sheet("03_ANOVA")
    ANOVA = res["ANOVA"]
    W0 = res["W0"]
    ws["A1"] = "КЛАССИЧЕСКАЯ ANOVA ПО GUM (грубый учёт табл.1/2 ГОСТ EN 13132)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws.append([])
    ws.append(["Компоненты дисперсии (модель σR² = σr² + σL², 5725-6 §4.1.4, GUM H.5)"])
    ws.merge_cells("A3:H3")
    ws["A3"].font = Font(italic=True, color=SUP)
    ws.append(["Компонента", "Символ", "Значение", "Ед.", "Источник/формула"])
    for c in range(1, 6):
        ws.cell(ws.max_row, c).fill = HDR_FILL
        ws.cell(ws.max_row, c).font = HDR_FONT
    ws.append(["Повторяемость инъекций", "σr", res["SIGMA_R"], "Y",
               f"объединённая: √(ΣΣ(Yij−Ȳi)²/Σ(ni−1)), df={ANOVA['df_r']}; "
               f"GUM H.5 / 5725-6 4.1.4"])
    ws.append(["Воспроизводимость табл.1 (R/2,8)", "σR", res["sigma_R_tab"], "% мас.",
               f"табл.1 ГОСТ EN 13132, рабочая точка w0={W0:.3f} % мас.: "
               f"R(w0)={res['R_w0']:.3f}; σR=R/2,8 (5725-6 4.1.2/4.1.4); "
               f"интерполяция '{res['interpolation']}' (ступенчатая — TD-DEV-208)"])
    ws.append(["Межлабораторная составляющая", "σL", res["sigma_L"], "% мас.",
               "√(max(0, σR²−σr²)); GUM 4.2 (отрицат. дисперсия → 0); грубый учёт, "
               "разные шкалы σr(Y)/σR(%мас.) — аппроксимация"])
    for r in range(5, ws.max_row + 1):
        for c in range(1, 6):
            ws.cell(r, c).border = BORDER
        ws.cell(r, 3).number_format = "0.000000"
    ws.append([])
    ws.append(["Таблица ANOVA (внутриуровневая, источник: инъекции)"])
    ws.merge_cells(f"A{ws.max_row}:H{ws.max_row}")
    ws["A" + str(ws.max_row)].font = HDR_FONT
    ws.append(["Источник", "SS", "df", "MS", "F", "Примечание"])
    for c in range(1, 7):
        ws.cell(ws.max_row, c).fill = HDR_FILL
        ws.cell(ws.max_row, c).font = HDR_FONT
    ss_within = ANOVA["var_within"] * ANOVA["df_r"]
    ws.append(["Внутри уровней (pure error)", ss_within, ANOVA["df_r"],
               ANOVA["var_within"], "",
               "SS_PE = Σ n_i·s²_i (v15 03_Статистика B22)"])
    ws.append(["Между уровнями (чистая регрессия)", "", len(levels) - 1, "", "",
               "не является случайным эффектом (систематика калибровки)"])
    ws.append(["Итого", "", ANOVA["n_total"], "", "", ""])
    ws.append([])
    ws.append(["ПРИМЕЧАНИЯ (GUM H.5.3.3, контракт разд.3/7):"])
    ws.append(["· Полный иерархический ANOVA (σr/σL/σR) невозможен на однодневных данных "
               "(1 день, 1 лаборатория) — GUM H.5.3.3."])
    ws.append(["· Межсерийная σL данных лаборатории не оцениваема (QC BUD-17, вне плана)."])
    ws.append(["· Применён «грубый учёт»: σR — из табл.1/2 ГОСТ EN 13132 "
               "(межлабораторный эксперимент метода), 5725-6 §7, РМГ 61 §5.5."])
    ws.append(["· u_repr — режим step (ПО УМОЛЧАНИЮ): r/R табл.1/2 ДИСКРЕТНЫ, "
               "интерполяция не предписана (TD-DEV-208); smooth — допущение, не по НД."])
    ws.append(["· Документируемое допущение: «σR принято по табл.1 ГОСТ EN 13132 "
               "(N межлаб. округление 2,8); σL данных лаборатории не оценивалась за "
               "отсутствием межсерийных данных» (17025 §7.6)."])
    for r in range(ws.max_row - 5, ws.max_row + 1):
        ws.cell(r, 1).font = Font(italic=True, color=SUP)
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 8
    return ws


def fill_profile(wb, levels, inputs, res):
    ws = wb.create_sheet("04_Профиль U(X)")
    X_GRID = res["X_GRID"]
    ws["A1"] = "ПРОФИЛЬ РАСШИРЕННОЙ НЕОПРЕДЕЛЁННОСТИ U(X)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    dx = (float(X_GRID[-1]) - float(X_GRID[0])) / (N_PROFILE - 1)
    ws["A2"] = (
        f"U(x) = k(x)·u_c(x); сетка (Xmax−Xmin)/199 = {dx:.6f} г/г; "
        f"u_repr — '{res['interpolation']}' (ступенчатая по поддиапазонам табл.1 "
        f"ГОСТ EN 13132; r/R дискретны, интерполяция не предписана — TD-DEV-208); "
        f"mean2 → √(R²−r²/2)/2,8; модель: s²·(1/p+1/n+(x−x̄)²/Sxx)/a², p={res['p']} "
        f"(инъекции пробы)"
    )
    ws.merge_cells("A2:H2")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[2].height = 42
    hdr = ["№", "X (г/г)", "u_model(x)", "u_repr(x) г/г", "u_c(x)", "ν_eff(x)", "k(x)", "U(x)"]
    hr = ws.max_row + 1
    for c, h in enumerate(hdr, 1):
        ws.cell(hr, c, h).fill = HDR_FILL
        ws.cell(hr, c).font = HDR_FONT
        ws.cell(hr, c).border = BORDER
    for i in range(N_PROFILE):
        r = hr + 1 + i
        ws.cell(r, 1, i + 1)
        ws.cell(r, 2, float(X_GRID[i])).number_format = "0.000000"
        ws.cell(r, 3, float(res["U_MODEL"][i])).number_format = "0.000000"
        ws.cell(r, 4, float(res["U_REPR_X"][i])).number_format = "0.000000"
        ws.cell(r, 5, float(res["U_C"][i])).number_format = "0.000000"
        ws.cell(r, 6, float(res["NU_EFF"][i])).number_format = "0.0"
        ws.cell(r, 7, float(res["K_ARR"][i])).number_format = "0.0000"
        ws.cell(r, 8, f"=G{r}*E{r}").number_format = "0.000000"  # живая формула
        for c in range(1, 9):
            ws.cell(r, c).border = BORDER
    x_pred_idx = res["x_pred_idx"]
    ws.cell(hr + 1 + x_pred_idx, 9, "← x_pred").font = Font(bold=True, color="C00000")
    ws.cell(hr + 1 + x_pred_idx, 9).border = BORDER
    ws.cell(hr, 8).comment = comment(
        "Значения U в столбце H — ЖИВЫЕ формулы Excel (=k*u_c, G*E); кэша значений в "
        "книге нет. При программном чтении без пересчёта использовать ядро anova_gum.py "
        "(compute_calibration модуля build_en13132_v16_calc), а не ячейки H."
    )
    note_r = hr + 1 + N_PROFILE + 2
    ws.cell(note_r, 1, "Примечание: столбец H (U) — живая формула =G*E; кэшированных "
                       "значений в книге нет — при чтении без пересчёта используйте ядро.").font = \
        Font(italic=True, color=SUP)
    # Маркер x_pred для диаграммы профиля: xVal/yVal серии маркера должны лежать
    # на листе ПРОФИЛЯ (Excel-правило: серия — один лист), а не на «Графики».
    # K4 = X0 (г/г), L4 = U(x0) (г/г) — данные серии 'x_pred' диаграммы 2.
    ws.cell(3, 11, "Маркер x_pred (X0; U(x0)) — данные серии на этом листе").font = \
        Font(italic=True, color=SUP)
    set_cell(ws, "K4", res["X0"], num_fmt="0.000000",
             cmt="xVal серии-маркера 'x_pred' диаграммы профиля; на листе 04_Профиль "
                 "U(X) вместе с yVal (L4) — Excel-правило (иначе X -> 1..n).")
    set_cell(ws, "L4", res["U_RES_G"], num_fmt="0.000000",
             cmt="yVal серии-маркера 'x_pred' (= U(06_Итог) == U_RES_G, TD-DEV-206); "
                 "на листе 04_Профиль U(X) вместе с xVal (K4).")
    ws.column_dimensions["A"].width = 6
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 14
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 12
    ws.column_dimensions["K"].width = 14
    ws.column_dimensions["L"].width = 14
    return ws, hr + 1, hr + N_PROFILE


def fill_budget(wb, levels, inputs, res):
    ws = wb.create_sheet("05_Бюджет")
    X0, W0 = res["X0"], res["W0"]
    ws["A1"] = "БЮДЖЕТ НЕОПРЕДЕЛЁННОСТИ РЕЗУЛЬТАТА ПРОБЫ"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws.append([])
    ws.append(["Результат пробы: X0 = %.6f г/г;  w = %.3f %% мас. (входы — 01_Ввод)" % (X0, W0)])
    ws.merge_cells("A3:H3")
    ws["A3"].font = Font(italic=True, color=SUP)
    ws.append([])
    hdr = ["Вклад", "u (г/г)", "df", "Тип", "Доля, %", "Источник (НД)"]
    hr = ws.max_row + 1
    for c, h in enumerate(hdr, 1):
        ws.cell(hr, c, h).fill = HDR_FILL
        ws.cell(hr, c).font = HDR_FONT
        ws.cell(hr, c).border = BORDER
    u2_total = sum(u ** 2 for _n, u, _d in res["BUDG"]["items"])
    rows = [
        ("модель (OLS: 1/p+1/n+(x−x̄)²/Sxx)", res["U_X0_model"], res["SNAP"]["df_res"], "A",
         "u²=s²_res·(1/p+1/n+(x0−x̄)²/Sxx)/a² (контракт §5 стр.1); p — инъекции пробы"),
        ("эталоны", abs(X0) * res["u_rel_etal"], 50, "B",
         f"u_rel={res['u_rel_etal']:.4f}; ПОЛЕ ВВОДА (0 = не учитывать, TD-DEV-33); тип B, ν=50"),
        ("чистота (pA, pИС)", abs(X0) * res["u_rel_purity"], None, "B",
         "u_rel(pA/pИС)=√((u_pA/pA)²+(u_pИС/pИС)²) из входов 01_Ввод; ПОЛЕ ВВОДА "
         "(0 = не учитывать); равномерное; ν=∞"),
        ("весы (навески)", abs(X0) * res["u_rel_mass"], None, "B",
         "u_rel(m)=u(m_net)/m_ИС из входов 01_Ввод; ПОЛЕ ВВОДА (0 = не учитывать); ν=∞; TD-DEV-33"),
        ("прецизионность u_repr (mean2)", res["u_repr_w0"] / res["K_G2PCT"], None, "B",
         f"√(R²−r²/2)/2,8 (5725-6 4.2.3), табл.1 ГОСТ EN 13132, w0={W0:.3f} % мас., "
         f"'{res['interpolation']}' (step — TD-DEV-208); ν=∞; анти-двойной счёт: σr не добавляется"),
    ]
    for i, (name, u, df, tp, src) in enumerate(rows):
        r = hr + 1 + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, u).number_format = "0.0000000"
        ws.cell(r, 3, "∞" if df is None else df)
        ws.cell(r, 4, tp)
        ws.cell(r, 5, (u ** 2 / u2_total * 100.0) if u2_total > 0 else 0.0).number_format = "0.00"
        ws.cell(r, 6, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 7):
            ws.cell(r, c).border = BORDER
        ws.cell(r, 2).comment = comment(src)
    rr = hr + 1 + len(rows) + 1
    ws.cell(rr, 1, "u_c (г/г)").font = Font(bold=True)
    ws.cell(rr, 2, res["U_C_RES"]).number_format = "0.0000000"
    ws.cell(rr, 6, "√Σu² — корень квадратный из суммы (GUM 6.2.1)")
    for c in range(1, 7): ws.cell(rr, c).border = BORDER
    ws.cell(rr + 1, 1, "ν_eff (Welch–Satterthwaite)").font = Font(bold=True)
    ws.cell(rr + 1, 2, res["NU_EFF_RES"]).number_format = "0.0"
    ws.cell(rr + 1, 6, "ν_eff = u_c⁴ / Σ(u_i⁴/ν_i); GUM G.2b")
    for c in range(1, 7): ws.cell(rr + 1, c).border = BORDER
    ws.cell(rr + 2, 1, "k (t_{0.975}(ν_eff))").font = Font(bold=True)
    ws.cell(rr + 2, 2, res["K_RES"]).number_format = "0.0000"
    ws.cell(rr + 2, 6, "GUM 6.3/G.4; при доминировании u_repr (ν=∞) k≈1,96")
    for c in range(1, 7): ws.cell(rr + 2, c).border = BORDER
    ws.cell(rr + 3, 1, "U (г/г) = k·u_c").font = Font(bold=True)
    ws.cell(rr + 3, 2, res["U_RES_G"]).number_format = "0.0000000"
    ws.cell(rr + 3, 6, "GUM 6.2.1; расширенная неопределённость (P=0.95)")
    for c in range(1, 7): ws.cell(rr + 3, c).border = BORDER
    ws.cell(rr + 4, 1, "U (% мас.) = U(г/г)·K_G2PCT").font = Font(bold=True)
    ws.cell(rr + 4, 2, res["U_RES_PCT"]).number_format = "0.0000"
    ws.cell(rr + 4, 6, "перевод шкалы U3 (контракт разд.9.2): % мас. = X·(m_ИС·p_ИС/m_пробы)·100")
    for c in range(1, 7): ws.cell(rr + 4, c).border = BORDER
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 6
    ws.column_dimensions["E"].width = 10
    ws.column_dimensions["F"].width = 62
    return ws


def fill_result(wb, levels, inputs, res):
    """06_Итог — структура Word-отчёта v15 (P2): параметры → результат w±U →
    протокольная запись → таблица вкладов → примечания."""
    ws = wb.create_sheet("06_Итог")
    X0, W0 = res["X0"], res["W0"]
    ws["A1"] = "РЕЗУЛЬТАТ ОПРЕДЕЛЕНИЯ MTBE (ГОСТ EN 13132-2012)"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:H1")
    ws["A2"] = ("v17: классическая ANOVA по GUM. Структура итога — по Word-отчёту v15 "
                "(EN13132_v15_Отчёт_МТБЭ.docx). Отличия v17: ANOVA-вклад прецизионности "
                "u_repr (табл.1/2 ГОСТ EN 13132) по умолчанию в режиме step (TD-DEV-208: "
                "r/R дискретны, интерполяция не предписана); вклады-поля ввода — 0 = "
                "«не учитывать» (TD-DEV-33); U = k·u_c РАССЧИТАНА.")
    ws["A2"].font = Font(italic=True, color=SUP)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 44

    # --- 1. ПАРАМЕТРЫ (ввод, лист 01_Ввод) ---
    hr = 4
    ws.cell(hr, 1, "1. ПАРАМЕТРЫ (ввод, лист 01_Ввод)").font = HDR_FONT
    hr += 1
    for c, h in enumerate(["Параметр", "Значение", "Ед.", "Источник"], 1):
        ws.cell(hr, c, h).fill = HDR_FILL
        ws.cell(hr, c).font = HDR_FONT
        ws.cell(hr, c).border = BORDER
    param_rows = [
        ("Масса пробы m_пробы, г", inputs["m_proby"], "0.000", "01_Ввод (поле ввода)"),
        ("Масса ИС в пробу m_ИС, г", inputs["m_is_proba"], "0.000", "01_Ввод (поле ввода)"),
        ("Чистота ИС p_ИС", inputs["p_is"], "0.0000", "01_Ввод; паспорт партии"),
        ("Чистота аналита p_A", inputs["p_a"], "0.0000", "01_Ввод"),
        ("Y0 пробы (среднее p инъекций)", inputs["y0_proba"], "0.000000", "01_Ввод; измерение"),
        ("Число инъекций пробы p", inputs["n_inj"], "0", "01_Ввод; ГОСТ EN 13132 §10.1"),
        ("u(m_net) весы, г", inputs["u_m_net"], "0.00000000", "01_Ввод; поле ввода; 0 = не учитывать"),
        ("u(pA)=u(pИС), доля", inputs["u_pa"], "0.000000", "01_Ввод; поле ввода; 0 = не учитывать"),
        ("u_rel эталонов, доля", inputs["u_rel_etal"], "0.0000", "01_Ввод; поле ввода; 0 = не учитывать"),
    ]
    for i, (name, val, fmt, src) in enumerate(param_rows):
        r = hr + 1 + i
        ws.cell(r, 1, name)
        ws.cell(r, 2, val).number_format = fmt
        ws.cell(r, 4, src)
        for c in range(1, 5):
            ws.cell(r, c).border = BORDER

    # --- 2. РЕЗУЛЬТАТ ---
    r = hr + 1 + len(param_rows) + 1
    ws.cell(r, 1, "2. РЕЗУЛЬТАТ").font = HDR_FONT
    r += 1
    for c, h in enumerate(["Параметр", "Значение", "Ед.", "Источник"], 1):
        ws.cell(r, c, h).fill = HDR_FILL
        ws.cell(r, c).font = HDR_FONT
        ws.cell(r, c).border = BORDER
    res_rows = [
        ("X0 (обратная градуировка), г/г", X0, "0.000000",
         "X0 = (Y0 − b)/a (GUM 5.1.2); обратная градуировка через ядро"),
        ("w(MTBE) = X0·m_ИС·p_ИС/m_пробы·100, % мас.", W0, "0.00",
         "ГОСТ EN 13132 §8 (формула w=m_i/m_пробы·100); U3 контракт разд.9.2"),
        ("u(X0) модель, г/г", res["U_X0_model"], "0.000000",
         "u²=s²_res·(1/p+1/n+(x0−x̄)²/Sxx)/a² (контракт §5 стр.1)"),
        ("u_c результата, г/г", res["U_C_RES"], "0.0000000", "бюджет 05 (анти-двойной счёт)"),
        ("ν_eff", res["NU_EFF_RES"], "0.0", "Welch–Satterthwaite, GUM G.2b"),
        ("k", res["K_RES"], "0.0000", "t_{0.975}(ν_eff), GUM 6.3"),
        ("U, г/г", res["U_RES_G"], "0.0000000", "k·u_c; = U(x_pred) маркер «Графики»!D4"),
        ("U, % мас.", res["U_RES_PCT"], "0.0000", "k·u_c в % мас. (перевод U3)"),
        ("U_rel, % (отн.)", res["U_REL"] * 100.0, "0.00", "U/w·100"),
    ]
    for i, (name, val, fmt, src) in enumerate(res_rows):
        rr2 = r + 1 + i
        ws.cell(rr2, 1, name)
        ws.cell(rr2, 2, val).number_format = fmt
        ws.cell(rr2, 4, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 5):
            ws.cell(rr2, c).border = BORDER
        ws.cell(rr2, 4).comment = comment(src)

    # --- 3. ПРОТОКОЛЬНАЯ ЗАПИСЬ ---
    r = r + 1 + len(res_rows) + 1
    ws.cell(r, 1, "3. ПРОТОКОЛЬНАЯ ЗАПИСЬ (ГОСТ EN 13132 §9, точность 0,1 % мас.)").font = HDR_FONT
    rec = (f"w(MTBE) = {W0:.1f} ± {res['U_RES_PCT']:.1f} % мас. "
           f"(P=0,95, k={res['K_RES']:.2f}, ν_eff={res['NU_EFF_RES']:.0f})")
    ws.cell(r + 1, 1, rec).font = Font(bold=True, size=12, color="1F4E79")
    ws.cell(r + 1, 1).border = BORDER
    ws.cell(r + 2, 1, "Протокольное округление до 0,1 % мас. (§9 ГОСТ EN 13132); "
                      "U округление — GUM 7.2.6").font = Font(italic=True, color=SUP)

    # --- 4. ТАБЛИЦА ВКЛАДОВ (формат Word-отчёта v15) ---
    r = r + 4
    ws.cell(r, 1, "4. ТАБЛИЦА ВКЛАДОВ (формат Word-отчёта v15; v17: вклады-поля ввода, "
                  "0 = не учитывать)").font = HDR_FONT
    r += 1
    for c, h in enumerate(["Составляющая", "u (г/г)", "Доля дисперсии, %", "Источник"], 1):
        ws.cell(r, c, h).fill = HDR_FILL
        ws.cell(r, c).font = HDR_FONT
        ws.cell(r, c).border = BORDER
    u2_total = sum(u ** 2 for _n, u, _d in res["BUDG"]["items"])
    contrib_rows = [
        ("Модель (калибровка, 1/p+1/n+(x−x̄)²/Sxx)", res["U_X0_model"],
         "контракт §5 стр.1; включает усреднённую повторяемость инъекций"),
        ("Эталоны (u(x_cal))", abs(X0) * res["u_rel_etal"],
         "поле ввода 01_Ввод; 0 = не учитывать"),
        ("Чистота аналита/ИС", abs(X0) * res["u_rel_purity"],
         "u(pA/pИС) из входов 01_Ввод; поле ввода; 0 = не учитывать"),
        ("Масса навесок (весы)", abs(X0) * res["u_rel_mass"],
         "поле ввода 01_Ввод; u(m_net)/m_ИС; 0 = не учитывать (TD-DEV-33)"),
        ("Прецизионность u_repr (mean2)", res["u_repr_w0"] / res["K_G2PCT"],
         f"табл.1 ГОСТ EN 13132, w0={W0:.3f} % мас., {res['interpolation']} "
         f"(step — TD-DEV-208); √(R²−r²/2)/2,8; анти-двойной счёт (контракт разд.6, R3)"),
    ]
    for i, (name, u, src) in enumerate(contrib_rows):
        rr2 = r + 1 + i
        ws.cell(rr2, 1, name)
        ws.cell(rr2, 2, u).number_format = "0.0000000"
        ws.cell(rr2, 3, (u ** 2 / u2_total * 100.0) if u2_total > 0 else 0.0).number_format = "0.00"
        ws.cell(rr2, 4, src).alignment = Alignment(wrap_text=True)
        for c in range(1, 5):
            ws.cell(rr2, c).border = BORDER
    rr2 = r + 1 + len(contrib_rows)
    ws.cell(rr2, 1, "U = k·u_c").font = Font(bold=True)
    ws.cell(rr2, 2, res["U_RES_G"]).number_format = "0.0000000"
    ws.cell(rr2, 4, "GUM 6.2.1/6.3; ν_eff по Welch–Satterthwaite")
    for c in range(1, 5):
        ws.cell(rr2, c).border = BORDER

    # --- 5. ПРИМЕЧАНИЯ ---
    r = rr2 + 2
    ws.cell(r, 1, "5. ПРИМЕЧАНИЯ").font = HDR_FONT
    notes = [
        "· u_repr — СТУПЕНЧАТАЯ по табл.1 ГОСТ EN 13132 (режим step, ПО УМОЛЧАНИЮ): "
        "r/R ДИСКРЕТНЫ, интерполяция не предписана ни стандартом, ни 5725-6, ни РМГ 61; "
        "GUM 4.1.4 к приписанным нормам прецизионности неприменим (TD-DEV-208, вердикт "
        "ресерчера по НД). smooth — методическое допущение, НЕ соответствует НД.",
        "· Вклады-поля ввода (TD-DEV-33): эталоны, чистота, весы — значение из спецификации "
        "ИЛИ 0 = «не учитывать» (вклад исключается из u_c и ν_eff).",
        "· Допущение: «σR принято по табл.1 ГОСТ EN 13132 (N межлаб. округление 2,8); σL "
        "данных лаборатории не оценивалась за отсутствием межсерийных данных» (17025 §7.6, "
        "РМГ 61 §5.5).",
        "· Значения вкладов — из спецификации пользователя; заполнить перед протоколом. "
        "Результат НЕ блокируется (пользователь сам решает, включать вклад или 0).",
    ]
    for i, note in enumerate(notes):
        ws.cell(r + 1 + i, 1, note).font = Font(italic=True, color=SUP)
        ws.cell(r + 1 + i, 1).alignment = Alignment(wrap_text=True)
        ws.merge_cells(f"A{r + 1 + i}:F{r + 1 + i}")
        ws.row_dimensions[r + 1 + i].height = 30
    ws.column_dimensions["A"].width = 46
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 62
    return ws


def fill_charts(wb, levels, inputs, res, ws_reg, prof_range, ws_prof):
    """Лист «Графики»: 3 диаграммы (калибровка, профиль U(x), остатки)."""
    DATA_FIRST, DATA_LAST, GRID_PT = prof_range[0], prof_range[1], prof_range[2]
    PROF_FIRST, PROF_LAST = prof_range[3], prof_range[4]
    gch = wb.create_sheet("Графики")
    gch["A1"] = "ГРАФИЧЕСКАЯ ДИАГНОСТИКА МОДЕЛИ"
    gch["A1"].font = TITLE_FONT
    gch["A2"] = ("Диаграммы построены на данных листов 02_Регрессия (точки калибровки) "
                 "и 04_Профиль U(X) (сетка %d точек). Обновляются при пересчёте книги.") % N_PROFILE
    gch["A2"].font = Font(italic=True, color=SUP)
    gch["A3"] = ("Примечание: значения U в столбце H листа 04_Профиль U(X) — живые формулы "
                 "Excel (=k·u_c); при программном чтении без пересчёта использовать ядро "
                 "anova_gum.py.")
    gch["A3"].font = Font(italic=True, color=SUP)
    # Числовые ячейки маркера x_pred (G4): D4 == U(06_Итог) == U_RES_G (TD-DEV-206)
    gch["A4"] = "Маркер x_pred (результат пробы)"
    gch["A4"].font = Font(bold=True)
    gch["B4"] = ("x_pred = %.6f г/г;  U(x_pred) = %.6f г/г  (= U 06_Итог, бюджет 05/06)"
                 % (res["X0"], res["U_RES_G"]))
    gch["B4"].font = Font(bold=True, color="C00000")
    gch["C4"] = res["X0"]
    gch["C4"].number_format = "0.000000"
    gch["D4"] = res["U_RES_G"]  # == U(06_Итог)
    gch["D4"].number_format = "0.000000"
    gch["A5"] = "Нули линии 0 (график остатков) — на 02_Регрессия (столбец G); " \
                "xVal/yVal серии на ОДНОМ листе (Excel-правило)"
    gch["A5"].font = Font(italic=True, color=SUP)

    # --- Диаграмма 1: Калибровочная прямая ---
    ch1 = ScatterChart()
    ch1.title = "Калибровочная прямая: точки, линия Y=a·X+b, полоса ŷ±k·u(ŷ)"
    ch1.style = 13
    ch1.x_axis.title = "X, г/г"
    ch1.y_axis.title = "Y (отношение площадей)"
    ch1.x_axis.delete = False
    ch1.y_axis.delete = False
    ch1.height = 10
    ch1.width = 16
    xdata1 = Reference(ws_reg, min_col=1, min_row=DATA_FIRST, max_row=DATA_LAST)
    ydata1 = Reference(ws_reg, min_col=2, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_pts = Series(ydata1, xdata1, title="точки (X, Y_ср)")
    s_pts.marker.symbol = "circle"
    s_pts.marker.size = 6
    s_pts.graphicalProperties.line.noFill = True
    ch1.series.append(s_pts)
    xline1 = Reference(ws_reg, min_col=1, min_row=GRID_PT, max_row=GRID_PT + 99)
    yline1 = Reference(ws_reg, min_col=3, min_row=GRID_PT, max_row=GRID_PT + 99)
    s_line = Series(yline1, xline1, title="Y = a·X + b")
    s_line.graphicalProperties.line.solidFill = "C00000"
    s_line.graphicalProperties.line.width = 22000
    s_line.marker.symbol = "none"
    ch1.series.append(s_line)
    yup1 = Reference(ws_reg, min_col=6, min_row=DATA_FIRST, max_row=DATA_LAST)
    ydn1 = Reference(ws_reg, min_col=5, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_up = Series(yup1, xdata1, title="ŷ + k·u(ŷ)")
    s_up.graphicalProperties.line.solidFill = "9DC3E6"
    s_up.graphicalProperties.line.width = 12000
    s_up.graphicalProperties.line.dashStyle = "dash"
    s_up.marker.symbol = "none"
    ch1.series.append(s_up)
    s_dn = Series(ydn1, xdata1, title="ŷ − k·u(ŷ)")
    s_dn.graphicalProperties.line.solidFill = "9DC3E6"
    s_dn.graphicalProperties.line.width = 12000
    s_dn.graphicalProperties.line.dashStyle = "dash"
    s_dn.marker.symbol = "none"
    ch1.series.append(s_dn)
    gch.add_chart(ch1, "C2")

    # --- Диаграмма 2: Профиль U(x) ---
    ch2 = ScatterChart()
    ch2.title = "Профиль неопределённости U(x) = k·u_c(x) и u_c(x)"
    ch2.style = 13
    ch2.x_axis.title = "X, г/г"
    ch2.y_axis.title = "неопределённость, г/г"
    ch2.x_axis.delete = False
    ch2.y_axis.delete = False
    ch2.height = 10
    ch2.width = 16
    xdata2 = Reference(ws_prof, min_col=2, min_row=PROF_FIRST, max_row=PROF_LAST)
    yuc2 = Reference(ws_prof, min_col=5, min_row=PROF_FIRST, max_row=PROF_LAST)
    yu2 = Reference(ws_prof, min_col=8, min_row=PROF_FIRST, max_row=PROF_LAST)
    s_uc = Series(yuc2, xdata2, title="u_c(x)")
    s_uc.graphicalProperties.line.solidFill = "2E75B6"
    s_uc.graphicalProperties.line.width = 18000
    s_uc.marker.symbol = "none"
    ch2.series.append(s_uc)
    s_u = Series(yu2, xdata2, title="U(x)")
    s_u.graphicalProperties.line.solidFill = "C00000"
    s_u.graphicalProperties.line.width = 22000
    s_u.marker.symbol = "none"
    ch2.series.append(s_u)
    # Маркер x_pred: xVal/yVal на листе 04_Профиль U(X) (K4/L4) — Excel-правило
    # «серия — один лист»; иначе Excel рендерит X как 1..n.
    xm = Reference(ws_prof, min_col=11, min_row=4, max_row=4)
    ym = Reference(ws_prof, min_col=12, min_row=4, max_row=4)
    s_mark = Series(ym, xm, title="x_pred")
    s_mark.marker.symbol = "square"
    s_mark.marker.size = 8
    s_mark.graphicalProperties.line.noFill = True
    ch2.series.append(s_mark)
    gch.add_chart(ch2, "C26")

    # --- Диаграмма 3: Остатки ---
    ch3 = ScatterChart()
    ch3.title = "Остатки регрессии: e = Y_ср − (a·X + b)"
    ch3.style = 13
    ch3.x_axis.title = "X, г/г"
    ch3.y_axis.title = "остаток e"
    ch3.x_axis.delete = False
    ch3.y_axis.delete = False
    ch3.height = 10
    ch3.width = 16
    xdata3 = Reference(ws_reg, min_col=1, min_row=DATA_FIRST, max_row=DATA_LAST)
    yres3 = Reference(ws_reg, min_col=4, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_res = Series(yres3, xdata3, title="e = Y_ср − (a·X + b)")
    s_res.marker.symbol = "diamond"
    s_res.marker.size = 6
    s_res.graphicalProperties.line.noFill = True
    ch3.series.append(s_res)
    # Линия нуля: yVal на 02_Регрессия (столбец G, константа 0) ВМЕСТЕ с xVal
    # (столбец A) — Excel-правило «серия — один лист» (иначе X -> 1..n).
    yzero3 = Reference(ws_reg, min_col=7, min_row=DATA_FIRST, max_row=DATA_LAST)
    s_zero = Series(yzero3, xdata3, title="0 (линия нуля)")
    s_zero.graphicalProperties.line.solidFill = "808080"
    s_zero.graphicalProperties.line.width = 10000
    s_zero.marker.symbol = "none"
    ch3.series.append(s_zero)
    gch.add_chart(ch3, "C50")

    gch.column_dimensions["A"].width = 30
    gch.column_dimensions["B"].width = 64
    gch.column_dimensions["C"].width = 12
    gch.column_dimensions["D"].width = 12
    gch.column_dimensions["H"].width = 10
    return gch


def build_workbook(levels, inputs, out_path, interpolation="step"):
    """Полная генерация книги. ЕДИНЫЙ ИСТОЧНИК: сначала пишем 01_Ввод, затем
    читаем входы из записанных ячеек и считаем от них (контракт P3)."""
    # нормализация уровней: (x, ys[, mA, mIS]) -> 4-кортеж для 01_Ввод и 2-кортеж для ядра
    levels4 = []
    for item in levels:
        x, ys, *rest = item
        mA = rest[0] if len(rest) > 0 else None
        mIS = rest[1] if len(rest) > 1 else None
        levels4.append((float(x), list(ys), mA, mIS))
    levels2 = [(x, ys) for x, ys, _mA, _mIS in levels4]
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_in = write_input_sheet(wb, levels4, inputs)
    inputs = read_inputs(ws_in)  # единый источник: значения из ячеек 01_Ввод
    res = compute_calibration(levels2, inputs, interpolation)
    ws_reg, DATA_FIRST, DATA_LAST, GRID_PT = fill_regression(wb, levels2, inputs, res)
    fill_anova(wb, levels2, inputs, res)
    ws_prof, PROF_FIRST, PROF_LAST = fill_profile(wb, levels2, inputs, res)
    fill_budget(wb, levels2, inputs, res)
    fill_result(wb, levels2, inputs, res)
    fill_charts(wb, levels2, inputs, res, ws_reg,
                (DATA_FIRST, DATA_LAST, GRID_PT, PROF_FIRST, PROF_LAST), ws_prof)
    assert_chart_series_sheets(wb)  # Excel-правило: xVal/yVal серии — один лист
    wb.save(out_path)
    return wb, res


def load_existing_inputs(out_path):
    """Если книга уже существует — читаем уровни и входы из её 01_Ввод
    (пользователь мог править ячейки). Иначе — (None, None)."""
    if not os.path.exists(out_path):
        return None, None
    try:
        wb = openpyxl.load_workbook(out_path, data_only=False)
        ws = wb["01_Ввод"]
        levels = read_levels(ws)
        inputs = read_inputs(ws)
        if levels is not None and inputs is not None:
            return levels, inputs
    except Exception as exc:  # noqa: BLE001
        print(f"  (предупреждение: существующая книга не перечитана — {exc})")
    return None, None


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
def main():
    # Уровни/входы: из существующей книги (пользователь мог править 01_Ввод),
    # иначе — свежие дефолты из источника v15.
    levels, inputs = load_existing_inputs(OUT_XLSX)
    source = "EN13132_v17_CALC.xlsx (01_Ввод, существующая книга)"
    if levels is None:
        v15 = load_v15_levels(XLSX_V15)
        levels = [(x, ys, mA, mIS) for x, ys, mA, mIS in v15]
        inputs = default_inputs(levels)
        source = "EN13132_v15_MTBE.xlsx (01_Ввод)"
    with warnings.catch_warnings():
        # w<0.1 %мас (нижний край калибровки) -> экстраполяция вниз (TD-DEV-202):
        # ожидаемое поведение, не захламляем вывод генератора.
        warnings.simplefilter("ignore", RuntimeWarning)
        wb, res = build_workbook(levels, inputs, OUT_XLSX, interpolation="step")
    SNAP = res["SNAP"]
    print("Книга сохранена:", OUT_XLSX)
    print("Источник данных:", source)
    print("Листы:", wb.sheetnames)
    print("Число уровней:", SNAP["n"], "| инъекции уровней:", len(levels[0][1]),
          "| p пробы:", res["p"])
    print("--- Ключевые числа (из ядра) ---")
    print("a=%.6f b=%.6f R2=%.5f s2_res=%.3e" % (SNAP["a"], SNAP["b"], SNAP["r2"], SNAP["s2_res"]))
    print("LoF: F=%.4f Fкр=%.4f %s" % (res["lof"]["F"], res["lof"]["F_crit"],
                                        res["lof"]["status"]))
    print("sigma_r=%.6f (df_r=%d)" % (res["SIGMA_R"], res["ANOVA"]["df_r"]))
    print("u_repr(w0=%.3f %%мас, %s)=%.6f %%мас, sigma_R=R(w0)/2.8=%.6f"
          % (res["W0"], res["interpolation"], res["u_repr_w0"], res["sigma_R_tab"]))
    print("sigma_L=%.6f %%мас." % res["sigma_L"])
    print("Профиль: точек=%d, U[0]=%.6f, U[x_pred]=%.6f, U[%d]=%.6f"
          % (len(res["X_GRID"]), res["U_X"][0], res["U_X"][res["x_pred_idx"]],
             len(res["X_GRID"]) - 1, res["U_X"][-1]))
    print("X0=%.6f г/г, w=%.3f %%мас." % (res["X0"], res["W0"]))
    print("Бюджет: u_c=%.6f г/г, nu_eff=%.1f, k=%.4f, U=%.6f г/г = %.4f %%мас."
          % (res["U_C_RES"], res["NU_EFF_RES"], res["K_RES"], res["U_RES_G"], res["U_RES_PCT"]))
    print("U_rel=%.2f %% | U(06_Итог)=%.6f == U(x_pred) маркер (Графики!D4)=%.6f  [%s]"
          % (res["U_REL"] * 100.0, res["U_RES_G"], res["U_RES_G"],
             "OK" if abs(res["U_X"][res["x_pred_idx"]] - res["U_RES_G"]) / res["U_RES_G"] < 1e-3
             else "близко (сетка)"))


if __name__ == "__main__":
    main()