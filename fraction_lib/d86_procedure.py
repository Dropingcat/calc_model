"""Процедурный контроль ASTM D86-23 Table 5 — TD-D86-002.

Table 5 задаёт условия испытания по группам продукта (1–4): размер пробы,
аппаратуру (колба/холодильник), скорость конденсации, времена нагрева.
Excel-книга v3.12 эти условия НЕ проверяла — это и был долг: расчёт мог быть
выполнен по данным, полученным вне нормативных условий, молча.

Модуль — нормативный «шлюз» между журналом наблюдения и calculate(), по той
же архитектуре, что procedure.py для ГОСТ 2177 (TD-2177-006). Математическое
ядро D86Calculator не затронут; флаги policy-нарушений совместимы с флагами
базового пайплайна (INVALID_TEST / WARNING:*).

Классы несоответствий (разделение — само по себе метрологически значимо):
  * INVALID — нарушение прямого требования Table 5: испытание недействительно,
              T(V) не подлежит репортингу без повторного определения;
  * WARNING — отклонение, не аннулирующее испытание по тексту НД: решение
              метролога фиксируется в протоколе.

НОРМАТИВНАЯ ОГОВОРКА (важно для аудита): числа ниже взяты из типовой сводки
Table 5 (D86-23): группы 1–4 = бензины / дизельные топки и лёгкие дистилляты /
газойли и средние дистилляты / мазуты и тяжёлые продукты. Лаборатория обязана
сверить значения с ОФИЦИАЛЬНОЙ лицензионной копией ASTM D86-23 до ввода листа
в эксплуатацию (реестр НД + hash — TD-D86-007; пункт закрытия TD-D86-002).
Константы вынесены в constants.py (D86_TABLE5_*), чтобы сверка правкой одной
таблицы не требовала изменения логики контроля.
"""

from __future__ import annotations

from dataclasses import dataclass

from .constants import (
    D86_TABLE5_SAMPLE_ML,
    D86_TABLE5_CONDENSE_RATE,
    D86_TABLE5_IBP_TO_5_MAX_MIN,
    D86_TABLE5_PREHEAT_MAX_MIN,
    D86_TABLE5_GROUPS,
)


@dataclass
class D86TestConditions:
    """Фактические условия проведения испытания (протокол оператора).

    Поля со значением None считаются незаполненными → MISSING (то же
    нормативное допущение, что в procedure.py: нельзя принимать решение по
    недокументированным условиям).
    """

    group: int | None = None            # группа продукта 1..4 (Table 1)
    sample_volume_ml: float | None = None   # залитый объём пробы, мл
    condense_rate_ml_min: float | None = None  # скорость конденсации, мл/мин
    time_to_ibp_min: float | None = None     # время нагрева до IBP, мин
    ibp_to_5_min: float | None = None        # время IBP→5 %, мин
    bath_preheat_ok: bool | None = None      # баня/горелка прогрета по Table 5
    apparatus_group_ok: bool | None = None   # колба/холодильник соответствуют группе
    note: str = ""

    def as_dict(self) -> dict:
        return {
            "group": self.group,
            "sample_volume_ml": self.sample_volume_ml,
            "condense_rate_ml_min": self.condense_rate_ml_min,
            "time_to_ibp_min": self.time_to_ibp_min,
            "ibp_to_5_min": self.ibp_to_5_min,
            "bath_preheat_ok": self.bath_preheat_ok,
            "apparatus_group_ok": self.apparatus_group_ok,
        }


# ---------------------------------------------------------------------------
def check_conditions(c: D86TestConditions) -> dict:
    """Проверка условий испытания по Table 5.

    Возвращает {"verdict": PASS|CONDITIONAL|FAIL,
                "violations": [ {id, severity, requirement, actual} ],
                "flags": [ 'TABLE5_INVALID:<id>', 'WARNING:TABLE5:<id>' ]}

    Флаги совместимы с res.Flags базового калькулятора: INVALID-нарушения
    начинаются с TABLE5_INVALID:, предупреждения — с WARNING:TABLE5:.
    """
    violations: list[dict] = []
    flags: list[str] = []

    if c.group is None or int(c.group) not in D86_TABLE5_GROUPS:
        violations.append({
            "id": "GROUP", "severity": "INVALID",
            "requirement": "Группа продукта должна быть указана (1..4, Table 1/Table 5)",
            "actual": repr(c.group),
        })
        flags.append("TABLE5_INVALID:GROUP")
        # дальше проверять бессмысленно — нет эталона группы
        return {"verdict": "FAIL", "violations": violations, "flags": flags}

    g = int(c.group)
    name = D86_TABLE5_GROUPS[g]

    def add(cid: str, sev: str, req: str, actual) -> None:
        violations.append({"id": cid, "severity": sev,
                           "requirement": req, "actual": actual})
        if sev == "INVALID":
            flags.append(f"TABLE5_INVALID:{cid}")
        else:
            flags.append(f"WARNING:TABLE5:{cid}")

    lo, hi = D86_TABLE5_SAMPLE_ML[g]
    if c.sample_volume_ml is None:
        add("SAMPLE_VOL", "MISSING", f"Объём пробы для гр.{g} ({name}): {lo}–{hi} мл", None)
    elif not (lo - 1e-9 <= c.sample_volume_ml <= hi + 1e-9):
        add("SAMPLE_VOL", "INVALID",
            f"Объём пробы для гр.{g} ({name}): {lo}–{hi} мл", c.sample_volume_ml)

    r_lo, r_hi = D86_TABLE5_CONDENSE_RATE[g]
    if c.condense_rate_ml_min is None:
        add("CONDENSE_RATE", "MISSING", f"Скорость конденсации гр.{g}: {r_lo}–{r_hi} мл/мин", None)
    elif not (r_lo - 1e-9 <= c.condense_rate_ml_min <= r_hi + 1e-9):
        add("CONDENSE_RATE", "INVALID",
            f"Скорость конденсации гр.{g}: {r_lo}–{r_hi} мл/мин", c.condense_rate_ml_min)

    ibp_lim = D86_TABLE5_IBP_TO_5_MAX_MIN[g]
    if ibp_lim is None:          # группа вне жёсткого лимита НД — не проверяем
        pass
    elif c.ibp_to_5_min is None:
        add("IBP_TO_5", "MISSING", f"IBP→5 % ≤ {ibp_lim} мин (гр.{g})", None)
    elif c.ibp_to_5_min > ibp_lim + 1e-9:
        add("IBP_TO_5", "INVALID", f"IBP→5 % ≤ {ibp_lim} мин (гр.{g})", c.ibp_to_5_min)

    t_lim = D86_TABLE5_PREHEAT_MAX_MIN[g]
    if t_lim is None:
        pass
    elif c.time_to_ibp_min is None:
        add("TIME_TO_IBP", "WARNING", f"Время до IBP ориентир ≤ {t_lim} мин (гр.{g})", None)
    elif c.time_to_ibp_min > t_lim + 1e-9:
        add("TIME_TO_IBP", "WARNING",
            f"Время до IBP ориентир ≤ {t_lim} мин (гр.{g})", c.time_to_ibp_min)

    if c.bath_preheat_ok is None:
        add("BATH_PREHEAT", "MISSING", "Баня/нагрев прогреты по Table 5", None)
    elif not c.bath_preheat_ok:
        add("BATH_PREHEAT", "INVALID", "Баня/нагрев прогреты по Table 5", False)

    if c.apparatus_group_ok is None:
        add("APPARATUS", "MISSING", f"Аппаратура соответствует гр.{g}", None)
    elif not c.apparatus_group_ok:
        add("APPARATUS", "INVALID", f"Аппаратура соответствует гр.{g}", False)

    invalid = [v for v in violations if v["severity"] == "INVALID"]
    missing = [v for v in violations if v["severity"] == "MISSING"]
    warning = [v for v in violations if v["severity"] == "WARNING"]
    if invalid or missing:
        verdict = "FAIL"          # MISSING = FAIL по той же политике, что procedure.py
    elif warning:
        verdict = "CONDITIONAL"
    else:
        verdict = "PASS"
    return {"verdict": verdict, "violations": violations, "flags": flags}


def table5_reference_rows() -> list[list]:
    """Строки справочника Table 5 для генерируемого Excel-листа."""
    rows = []
    for g in sorted(D86_TABLE5_GROUPS):
        lo, hi = D86_TABLE5_SAMPLE_ML[g]
        r_lo, r_hi = D86_TABLE5_CONDENSE_RATE[g]
        ibp = D86_TABLE5_IBP_TO_5_MAX_MIN[g]
        pre = D86_TABLE5_PREHEAT_MAX_MIN[g]
        rows.append([g, D86_TABLE5_GROUPS[g], f"{lo}–{hi}",
                     f"{r_lo}–{r_hi}",
                     "—" if ibp is None else f"≤ {ibp}",
                     "—" if pre is None else f"≤ {pre}"])
    return rows


__all__ = ["D86TestConditions", "check_conditions", "table5_reference_rows"]
