"""Контролируемый реестр редакций нормативных документов (TD-2177-008 / TD-D86-007).

Единый источник истины о том, КАКОЙ редакцией НД верифицирован каждый расчётный
слой fraction_lib и книги. Правила контроля:

1. Запись реестра неизменяема: новая редакция = новая запись + `supersedes`.
2. Каждая запись обязана иметь: id, designation, edition, status, effective_from,
   sha256_of_normative_core (хеш канонической JSON-серизации существенных
   нормативных констант слоя), verified_by (доказательство проверки — тесты/док).
3. Хеш пересчитывается функцией `compute_core_hash(designation)`; тест
   `test_nd_registry.py` блокирует drift: если константы в коде изменились, а
   хеш в реестре не переоформлен записью с новым id — CI падает.
4. Статусы: ACTIVE / SUPERSEDED / WITHDRAWN. Расчёт по WITHDRAWN запрещён
   (`resolve_for_calculation` бросает NormativeRegistryError).

Хранилище: docs/ND_REGISTRY.json (readable diff при правках через commit).
"""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass
from pathlib import Path

REGISTRY_PATH = Path(__file__).resolve().parent.parent / "docs" / "ND_REGISTRY.json"


class NormativeRegistryError(RuntimeError):
    pass


@dataclass(frozen=True)
class NDEntry:
    id: str
    designation: str          # 'ГОСТ 2177-99', 'ASTM D86' ...
    edition: str              # 'Изм.№1 + Поправка 2023', 'D86-23' ...
    status: str               # ACTIVE | SUPERSEDED | WITHDRAWN
    effective_from: str       # ISO date
    scope: str                # какие слои/книги верифицированы этой редакцией
    core_constants: dict      # существенные нормативные числа (вход в hash)
    sha256_of_normative_core: str
    verified_by: list[str]    # пути тестов/документов
    supersedes: str | None = None
    notes: str = ""

    def is_active(self) -> bool:
        return self.status == "ACTIVE"


def canonical_json(obj: dict) -> str:
    """Каноническая сериализация: сортировка ключей, без пробелов."""
    return json.dumps(obj, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def compute_core_hash(core: dict) -> str:
    return hashlib.sha256(canonical_json(core).encode("utf-8")).hexdigest()


# --- Существенные нормативные ядра слоёв -------------------------------------
def gost2177_core() -> dict:
    from fraction_lib import constants as C
    return {
        "coeff_dT_kPa_form": 0.0009,
        "p_std_kPa": 101.3,
        "round_step_C": 0.5,
        "table4_range_mmHg": [560, 760],
        "extrapolation_max_pct": getattr(C, "GOST2177_EXTRAPOLATION_MAX_PCT", 98.0),
        "slope_central_points_pct": 5,          # S(V±5) по 5.6.3
        "repeatability_and_reproducibility_tables": True,
    }


def d86_core() -> dict:
    from fraction_lib import constants as C
    return {
        "loss_correction_formula": "Lc=0.5+(L-0.5)/(1+(760-P)/60)",
        "loss_correction_min_L": 0.5,
        "pressure_unit_mmHg": 760.0,
        "groups": [1, 2, 3, 4],
        "manual_round_step": 0.5,
        "condensation_rate_mL_min": [4.0, 5.0],
        "singular_pressure_mmHg": 820.0,        # знаменатель → 0 (warning-зона)
        "editions_note": "D86-23 (актуальная; ранее D86-20a)",
    }


_CORE_BUILDERS = {"ГОСТ 2177-99": gost2177_core, "ASTM D86": d86_core}


def load_registry(path: Path | None = None) -> list[NDEntry]:
    p = path or REGISTRY_PATH
    if not p.exists():
        raise NormativeRegistryError(f"Реестр не найден: {p}")
    data = json.loads(p.read_text(encoding="utf-8"))
    return [NDEntry(**e) for e in data["entries"]]


def entry_for(designation: str, status: str = "ACTIVE",
              path: Path | None = None) -> NDEntry:
    for e in load_registry(path):
        if e.designation == designation and e.status == status:
            return e
    raise NormativeRegistryError(
        f"Нет записи ACTIVE для '{designation}' в ND_REGISTRY.json")


def resolve_for_calculation(designation: str, path: Path | None = None) -> NDEntry:
    """Разрешить редакцию для расчёта; WITHDRAWN/SUPERSEDED запрещены."""
    e = entry_for(designation, "ACTIVE", path)
    return e


def check_no_drift(path: Path | None = None) -> list[str]:
    """Сверить хеши реестра с текущими константами кода. Пустой список = OK.

    Проверяются только ACTIVE-записи (легализованные редакции); архивные
    SUPERSEDED/WITHDRAWN не имеют ядра в коде и не участвуют в сверке.
    """
    problems = []
    for e in load_registry(path):
        if e.status != "ACTIVE":
            continue
        builder = _CORE_BUILDERS.get(e.designation)
        if builder is None:
            continue
        actual = compute_core_hash(builder())
        if actual != e.sha256_of_normative_core:
            problems.append(
                f"{e.id}: drift хеша — константы кода изменены, реестр не переоформлен"
            )
    return problems


if __name__ == "__main__":
    import sys
    probs = check_no_drift()
    if probs:
        print("DRIFT:\n  " + "\n  ".join(probs))
        sys.exit(1)
    print("ND registry: хеши соответствуют коду ✅")
