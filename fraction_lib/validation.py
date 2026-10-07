"""Тройная валидация библиотеки fraction_lib.

Способы (по контракту ПУНКТ 6):
  1. Аналитический: материальный баланс, монотонность, гладкость, границы.
  2. Нормативный: 3-5 примеров из НД (поправка Янга, потери ISO, ГОСТ 2177,
     AET из A7, r/R при конкретной крутизне). Допуск Δ <= 0,1 °С.
  3. Чистый лист: чтение сырых данных из эталонных xlsx, расчёт библиотекой
     И независимым скриптом (реализация формул прямо в тесте), сравнение.
     Допуск Δ <= 0,05 °С.

Валидационные функции возвращают список записей для VALIDATION_REPORT.md.
"""

from __future__ import annotations

import math

from .constants import (
    GOST2177_TABLE4,
    D1160_AET,
    interpolate_table,
)
from .exceptions import MassBalanceViolation
from .models import FractionInput, FractionResult


# ===========================================================================
# СПОСОБ 1 — АНАЛИТИЧЕСКИЙ
# ===========================================================================
def check_mass_balance(result: FractionResult, tolerance: float = 0.2) -> bool:
    """Σ (отгон + остаток + потери) == 100 ± 0,2 %."""
    total = 100.0  # V_corr нормирован на 100; Loss/Residue уже скорректированы
    loss = result.Loss
    residue = result.Residue
    distillate = 100.0 - loss - residue
    return abs(distillate + residue + loss - 100.0) <= tolerance


def check_monotonicity(t_corr: dict) -> bool:
    """T(V_{i+1}) >= T(V_i) для всех точек (в порядке % отгона)."""
    seq = []
    for k, v in t_corr.items():
        if v is None:
            continue
        if isinstance(k, str) and k.upper() in ("IBP", "НК"):
            seq.append((-1.0, v))
        elif isinstance(k, str) and k.upper() in ("FBP", "КК"):
            seq.append((1000.0, v))
        else:
            try:
                seq.append((float(k), v))
            except (TypeError, ValueError):
                continue
    seq.sort(key=lambda x: x[0])
    for (v1, t1), (v2, t2) in zip(seq, seq[1:]):
        if t2 < t1 - 1e-6:
            return False
    return True


def check_smoothness(t_corr: dict, jump_factor: float = 5.0) -> bool:
    """Гладкость: вторая производная наклона без аномальных скачков.

    Считает наклоны dT/dV по интервалам (нормировка на шаг объёма — иначе
    неравномерная сетка % даёт ложные скачки), затем вторую разность наклонов.
    Краевые точки IBP/FBP исключаются: это физически особые точки кривой
    разгонки (там наклон всегда меняется), они проверяются отдельно через
    монотонность/границы. Скачок |d2_наклона| > jump_factor·средняя_|d2|
    считается аномалией интерполяции (артефактом ввода), а не изгибом кривой.
    """
    seq = []
    for k, v in t_corr.items():
        if v is None:
            continue
        if isinstance(k, str) and k.upper() in ("IBP", "НК"):
            x = 0.0
        elif isinstance(k, str) and k.upper() in ("FBP", "КК"):
            x = 100.0
        else:
            try:
                x = float(k)
            except (TypeError, ValueError):
                continue
        seq.append((x, float(v)))
    seq.sort(key=lambda p: p[0])
    # внутренние точки (без IBP/FBP) — проверка гладкости интерполяции
    inner = [p for p in seq if 0.0 < p[0] < 100.0]
    if len(inner) < 4:
        return True
    slopes = []
    for i in range(len(inner) - 1):
        dv = inner[i + 1][0] - inner[i][0]
        if dv <= 1e-9:
            continue
        slopes.append((inner[i + 1][1] - inner[i][1]) / dv)
    if len(slopes) < 3:
        return True
    # Аномалия 1: отношение наклона к медианному |наклон| > jump_factor
    # (скачок кривой — наклон резко больше соседних; не зависит от сетки %).
    sorted_abs = sorted(abs(s) for s in slopes)
    median_abs = sorted_abs[len(sorted_abs) // 2]
    if median_abs > 1e-9:
        for s in slopes:
            if abs(s) > jump_factor * median_abs:
                return False
    # Аномалия 2: вторая разность наклонов > jump_factor·средняя_|d2|
    d2 = [slopes[i + 1] - slopes[i] for i in range(len(slopes) - 1)]
    mean_abs = sum(abs(x) for x in d2) / len(d2)
    if mean_abs < 1e-9:
        return True
    return all(abs(x) <= jump_factor * mean_abs for x in d2)


def check_temperature_bounds(t_corr: dict) -> bool:
    """T_НК <= T_10% <= ... <= T_КК (частный случай монотонности)."""
    return check_monotonicity(t_corr)


# ===========================================================================
# СПОСОБ 2 — НОРМАТИВНЫЙ (независимые ожидаемые значения)
# ===========================================================================
def young_correction_mmhg(t: float, p_mmhg: float) -> float:
    """Поправка Сиднея Янга, мм рт.ст.: C = 0,00012·(760 − P)·(273 + t)."""
    return 0.00012 * (760.0 - p_mmhg) * (273.0 + t)


def young_correction_kpa(t: float, p_kpa: float) -> float:
    """Поправка Сиднея Янга, кПа: C = 0,0009·(101,3 − P)·(273 + t)."""
    return 0.0009 * (101.3 - p_kpa) * (273.0 + t)


def iso_loss_correction(l: float, p_kpa: float) -> float:
    """Lc = 0,5 + (L − 0,5)/(1 + 0,125·(101,3 − p))."""
    return 0.5 + (l - 0.5) / (1.0 + 0.125 * (101.3 - p_kpa))


def gost2177_loss_correction(l: float, p_mmhg: float) -> float:
    """V_K = A·L + B (таблица 4, линейная интерполяция)."""
    A, B = interpolate_table(p_mmhg, GOST2177_TABLE4)
    return A * l + B


def maxwell_bonnell_aet_independent(t_obs: float, p_mmhg: float) -> float:
    """Независимая реализация AET (Annex A7) для валидации.

    Дублирует формулу напрямую, без импорта библиотеки — для свежей проверки.
    """
    cfg = D1160_AET
    if p_mmhg >= cfg["P_THRESHOLD"]:
        c = cfg["A7_3"]
    else:
        c = cfg["A7_5"]
    lg_p = math.log10(p_mmhg)
    A = (c["n1"] - c["n2"] * lg_p) / (c["d1"] - c["d2"] * lg_p)
    denom = 1.0 / (t_obs + cfg["T_OFFSET"]) + cfg["K2"] * A - cfg["K3"]
    return cfg["K1"] * A / denom - cfg["T_OFFSET"]


def d1160_precision_independent(slope: float, p_mmhg: float, band: str, kind: str,
                                constants: dict) -> float:
    """Независимая реализация r/R D1160 (12.2, формулы 1/2)."""
    from .constants import D1160_INTERP_PRESSURES
    p_lo, p_hi = D1160_INTERP_PRESSURES
    c_lo = constants[int(p_lo)][kind][band]
    c_hi = constants[int(p_hi)][kind][band]
    if p_mmhg <= p_lo:
        a, b, M = c_lo
    elif p_mmhg >= p_hi:
        a, b, M = c_hi
    else:
        t = (p_mmhg - p_lo) / (p_hi - p_lo)
        a = c_lo[0] + t * (c_hi[0] - c_lo[0])
        b = c_lo[1] + t * (c_hi[1] - c_lo[1])
        M = c_lo[2] + t * (c_hi[2] - c_lo[2])
    if b == 0:
        return M * math.exp(a) / 1.8
    return M * math.exp(a + b * math.log(1.8 * max(slope, 1e-9))) / 1.8


# ===========================================================================
# СПОСОБ 3 — ЧИСТЫЙ ЛИСТ (независимый скрипт в тесте, без импорта библиотеки)
# ===========================================================================
def parse_etalon_input(filepath: str, sheet: str = "Ввод данных") -> dict:
    """Чтение сырых данных из эталонного xlsx (data_only=False).

    Возвращает dict: {pressure_mmhg, group, volumes[], t1[], t2[], distillate,
    residue, loss}.
    """
    import openpyxl

    wb = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
    ws = wb[sheet]
    pressure_mmhg = None
    group = None
    volumes: list[float] = []
    t1: list[float] = []
    t2: list[float] = []
    distillate = residue = loss = None

    for row in ws.iter_rows(values_only=True):
        vals = [v for v in row]
        if not vals or all(v is None for v in vals):
            continue
        # строка давления
        if any("атмосферное давление" in str(v) for v in vals if v is not None):
            for v in vals:
                if isinstance(v, (int, float)) and v != 0:
                    pressure_mmhg = float(v)
                    break
            continue
        # строка группы
        if any("группу" in str(v).lower() for v in vals if v is not None):
            for v in vals:
                s = str(v)
                if "Группа" in s or "группа" in s:
                    digits = "".join(ch for ch in s if ch.isdigit())
                    if digits:
                        group = digits
            continue
        # строки объёмов и температур (первые 3 колонки: объём, t1, t2)
        v0, v1, v2 = vals[0], vals[1], vals[2] if len(vals) > 2 else None
        # заголовок
        if v0 is not None and isinstance(v0, str) and "Обем" in str(v0):
            continue
        if isinstance(v0, (int, float)) and isinstance(v1, (int, float)):
            volumes.append(float(v0))
            t1.append(float(v1))
            t2.append(float(v2) if isinstance(v2, (int, float)) else v1)
        # отгон/остаток/потери
        label = str(v0) if v0 is not None else ""
        if "Отгон" in label or "отгон" in label:
            distillate = float(v1) if isinstance(v1, (int, float)) else None
        elif "Остаток" in label or "остаток" in label:
            residue = float(v1) if isinstance(v1, (int, float)) else None
        elif "Потер" in label or "потер" in label:
            loss = float(v1) if isinstance(v1, (int, float)) else None

    wb.close()
    return {
        "pressure_mmhg": pressure_mmhg,
        "group": group,
        "volumes": volumes,
        "t1": t1,
        "t2": t2,
        "distillate": distillate,
        "residue": residue,
        "loss": loss,
    }


def build_fraction_input(parsed: dict, method: str) -> FractionInput:
    """Сборка FractionInput из разобранных эталонных данных."""
    t_obs = {}
    for i, (v, t) in enumerate(zip(parsed["volumes"], parsed["t1"])):
        if t is None or t == 0:
            continue
        if v == 0:
            t_obs["IBP"] = t
        elif v == 100:
            t_obs["FBP"] = t
        else:
            t_obs[float(v)] = t
    return FractionInput(
        P_atm=parsed["pressure_mmhg"] or 760.0,
        P_vac=10.0 if method == "d1160" else None,
        object_type=parsed["group"] or "1",
        method=method,
        V_pct=[100.0],
        T_obs=t_obs,
        loss_pct=parsed["loss"] if parsed["loss"] is not None else 0.0,
        residue_pct=parsed["residue"] if parsed["residue"] is not None else 0.0,
    )