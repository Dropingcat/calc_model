"""Процедурный checklist ГОСТ 2177-99, раздел 5 (пп. 5.1–5.5) — TD-2177-006.

Разделяет два класса несоответствий (это и был долг):
  * INVALID  — нарушение прямого требования НД: испытание недействительно,
               результат не подлежит использованию (например, скорость
               конденсации вне 4–5 мл/мин при отборе 5%...95%, п.5.4.6);
  * WARNING  — отклонение от рекомендуемых условий или запись без
               подтверждения: результат условно годен, решение метролога.

Checklist заполняется оператором по факту испытания (протокол) и проверяется
перед расчётом. Математический слой (GOST2177Calculator) остаётся независимым;
этот модуль — нормативный «шлюз» между журналом наблюдения и calculate().

Нормативные пункты (ГОСТ 2177-99 с Изм. №1):
  5.1  отбор и подготовка пробы (без потери лёгких фракций, температура);
  5.2  аппаратура: колба 100 мл, холодильник по группе, термометр с поверкой,
       барометр с погрешностью <= 0,1 кПа (<= 1 мм рт. ст.), местное давление;
  5.3  сборка: термометр по метке колбы, уплотнения, экран/баня;
  5.4  ход испытания: время нагрева до IBP, IBP->5% <= 10 мин (метод А),
       скорость конденсации 4–5 мл/мин, равномерность, учёт потерь/остатка;
  5.5  измерения и записи: считывание по мениску, повторность определений.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class CheckItem:
    id: str                      # напр. "5.4.6"
    requirement: str             # формулировка требования (кратко, с ссылкой)
    answer: bool | None = None   # True/False/None(не заполнено)
    value: float | None = None   # фактическое числовое значение (если есть)
    note: str = ""
    severity: str = "INVALID"    # INVALID | WARNING

    @property
    def status(self) -> str:
        if self.answer is None:
            return "MISSING"     # незаполненный пункт = тоже несоответствие
        return "OK" if self.answer else self.severity


# Дефолтный набор пунктов (заполняется оператором). Границы значений —
# константы метода А ГОСТ 2177 (для метода Б часть пунктов ужесточается/
# ослабляется по редакции; см. requirement-текст каждого пункта).
DEFAULT_CLAUSE_VALUES = {
    "barometer_uncertainty_kpa_max": 0.1,   # п.5.2 (погрешность барометра)
    "time_to_ibp_min_max": 10.0,            # п.5.4 (IBP)
    "ibp_to_5pct_min_max": 10.0,            # п.5.4 (начало отбора 5%)
    "condense_rate_ml_min": (4.0, 5.0),     # п.5.4.6 скорость конденсации
    "read_by_meniscus": True,               # п.5.5
}


def build_checklist(method: str = "A") -> list[CheckItem]:
    """Свежий незаполненный чек-лист пп. 5.1–5.5 для протокола."""
    m = method.upper()
    items = [
        CheckItem("5.1", "Проба отобрана и подготовлена по разделу 4; потеря лёгких фракций исключена"),
        CheckItem("5.2a", "Колба 100 мл, комплектность по аппаратуре (группа продукта соответствует)"),
        CheckItem("5.2b", "Термометр с действующей поверкой/калибровкой (журнал приборов)", severity="WARNING"),
        CheckItem("5.2c", "Барометр: погрешность <= 0,1 кПа, измерено МЕСТНОЕ давление (не sea-level)",
                  value=None),
        CheckItem("5.3", "Сборка: положение резервуара термометра против метки нижнего отвода колбы"),
        CheckItem("5.4a", "Нагрев регулируемый; время до IBP в норме", severity="WARNING"),
        CheckItem("5.4b", "Отбор первых 5% — не более установленного времени"),
        CheckItem("5.4c", "Скорость конденсации 4–5 мл/мин на всём протяжении отбора"),
        CheckItem("5.4d", "Охлаждение холодильника непрерывно, среда нужной температуры"),
        CheckItem("5.5a", "Считывание T по нижнему мениску (для непрозрачных продуктов — по верхнему), метод А"),
        CheckItem("5.5b", "Объёмы отсчитаны по делениям, потери и остаток записаны"),
        CheckItem("5.5c", "Два определения выполнены (пары X1/X2 переданы в расчёт)", severity="WARNING"),
    ]
    if m == "B":
        for it in items:
            if it.id == "5.4b":
                it.requirement += " (метод Б: уточнить по редакции)"
    return items


def evaluate_checklist(items: list[CheckItem]) -> dict:
    """Вердикт по чек-листу: список несоответствий + общий статус.

    Возвращает {"verdict": "PASS"|"CONDITIONAL"|"FAIL",
                "missing": [...], "invalid": [...], "warning": [...]}
    MISSING трактуется как FAIL (нельзя считать по недокументированным условиям).
    """
    missing = [i.id for i in items if i.status == "MISSING"]
    invalid = [i.id for i in items if i.status == "INVALID"]
    warning = [i.id for i in items if i.status == "WARNING"]
    if missing or invalid:
        verdict = "FAIL"
    elif warning:
        verdict = "CONDITIONAL"
    else:
        verdict = "PASS"
    return {"verdict": verdict, "missing": missing,
            "invalid": invalid, "warning": warning}


def check_condensation_rate(rate_ml_min: float,
                            bounds=DEFAULT_CLAUSE_VALUES["condense_rate_ml_min"],
                            items: list[CheckItem] | None = None) -> bool:
    """Числовой контроль п.5.4.6; при переданном items обновляет пункт 5.4c."""
    ok = bounds[0] - 1e-9 <= rate_ml_min <= bounds[1] + 1e-9
    if items:
        for it in items:
            if it.id == "5.4c":
                it.value = rate_ml_min
                it.answer = ok
    return ok
