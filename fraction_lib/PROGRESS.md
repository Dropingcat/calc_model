# PROGRESS — fraction_lib (контракт fraction_lib_v2.md)

Задача: рефакторинг расчёта фракционного состава в Python-библиотеку `fraction_lib`
по канону «Агенты и вайб-кодинг» (спецификация = договор о проверке; один пункт за раз;
тесты до кода; проверку пишет не тот, кого проверяют; нельзя быть себе судьёй;
файл прогресса; чекпоинты; критерий завершения — только код/тесты).

Канбан task_id: `fraction-lib-20261002` (секция code-factory).
Контекст: TD-172 (аудит фракционки, 22 техдолга ФР-01..ФР-22), ФР-23+ (рефакторинг).
Эталоны: `snapshot_etalon_20261001/` (5 xlsx, только чтение).
Нормативки: `нормативки/` (PDF + КЛЮЧЕВЫЕ_*.md + СВОДКА).

---

## РИТУАЛ СТАРТА
- [x] project_context.py прочитан (kanban: TD-172-фракционка DONE research; 171 TD открыто; factory DONE/attempt=1).
- [x] Капсула orchestration-thread-process.md прочитана.
- [x] «Где я» проговорено: WS — новый пункт fraction_lib; TD-172/ФР-23+ контекст; канбан fraction-lib-20261002; связь: Excel → Python-библиотека → тройная валидация.
- [x] Осмотр остатков предыдущей попытки: models/exceptions/constants/base/gost2177/iso3405/d86/d1160 — код есть; НЕТ __init__.py, validation.py, tests/, PROGRESS.md, README.md, VALIDATION_REPORT.md.
- [x] Среда: Python 3.11.15, openpyxl 3.1.5, pytest 9.1.1; эталоны и нормативки на месте.

## ПУНКТЫ
- [x] пункт 1 — Каркас (models/exceptions/constants + __init__.py): 24 теста PASS (test_models.py).
- [x] пункт 2 — gost2177.py: 17 тестов PASS (test_gost2177.py). Исправлен interpolate_table (границы 560/760 включаются).
- [x] пункт 3 — iso3405.py: 21 тест PASS (test_iso3405.py). Исправлены: поправка Янга в кПа (ISO), round_half (half away from zero), _compute_slopes (ключи 95/последняя точка), _to_pct_key нормализация, k_criterion защита от domain error, base.calculate пропуск None.
- [x] пункт 4 — d86.py: 14 тестов PASS (test_d86.py).
- [x] пункт 5 — d1160.py: 18 тестов PASS (test_d1160.py). Исправлены критические: формула A7 (приоритет операций: A = (n1−n2·lgP)/(d1−d2·lgP)), d1160_precision (убран лишний e·), OverlapViolation (сравнение AET, нормализация ключей).
- [ ] пункт 6 — тройная валидация (validation.py + test_analytic/test_normative/test_cleanroom + VALIDATION_REPORT.md).
- [ ] пункт 4 — d86.py: код есть, тесты отсутствуют → пишу tests/test_d86.py.
- [ ] пункт 5 — d1160.py: код есть, тесты отсутствуют → пишу tests/test_d1160.py.
- [ ] пункт 6 — тройная валидация (validation.py + test_analytic/test_normative/test_cleanroom + VALIDATION_REPORT.md).
- [ ] пункт 7 — README.md + финальный прогон pytest + все пункты [x].
- [ ] РИТУАЛ ЗАКРЫТИЯ: kanban_report.py + контроль остатка нити.