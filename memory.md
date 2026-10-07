# MEMORY — памятка ассистенту по сессии (calc_model)

_Обновлено: 2026-10-08_

## Окружение и git
- Рабочий каталог: `/workspace` (git-репозиторий).
- **Единственная рабочая ветка: `main`.** Feature-ветки (`feature/wls-v3`) исторически отстают и НЕ мержить в main без крайней необходимости.
- Внешний репо: `origin → https://github.com/Dropingcat/calc_model.git`.
- Токен GitHub:
  - доступен как переменная окружения `token_git` (продублирован в `GITHUB_TOKEN`);
  - прописан в `~/.bashrc` (`export token_git=...`) — после `source ~/.bashrc` виден в shell;
  - хранится в `/workspace/.env` (права 600, добавлен в `.gitignore`);
  - для git-аутентификации настроен `credential.helper store` → `/root/.git-credentials` (права 600).
  - **В URL remote токен НЕ вшивать** — credentials подставляются автоматически.
- Проверка доступа: `git ls-remote origin` работает без интерактивного ввода.
- Push/pull работают «из коробки»: `git push -u origin main` — без запроса учётных данных.

## Состояние репозитория (снимок 2026-10-08: ФАЙЛЫ ИЗ РЕПО)
- HEAD `main` = `db629c9` **«v6.0: GUI с вкладками»** — проект ушёл далеко вперёд относительно локальной копии.
- Эволюция в истории main: v4 (WLS+LoF+Welch-Satterthwaite) → v5/v5.1/v5.2 (back-calc, LOD/LOQ, бюджет неопределённости, лист «Итог», защита листов) → v5.4 (отчёт Word) → v5.5 (портативный CLI + exe onefile, Git LFS) → v6.0 (GUI tkinter).
- Состав main (23 файла, плоская структура, папок нет):
  - ядро расчёта: `cal_v4.py` (~93 КБ), `calc_engine.py` (автономный движок без Excel);
  - интерфейс: `gui.py`, `cal_app.py` (CLI);
  - валидация/MC: `validate_v4.py`, `validate_auto.py`, `lib_validate.py`, `mc_v4.py`;
  - отчётность: `report_v4.py` (Word);
  - тесты: `e2e_test_v4.py`, `test_gui.py`, `test_td07.py`;
  - старые версии моделей: `cal_v2.py`, `cal_v3.py`; артефакты: `Metrology_Core_EURACHEM_v2.xlsx`, `_v3.xlsx`;
  - служебное: `README.md`, `memory.md`, `tracker.md`, `techdebt.md`, `verify_repo.py`, `.gitignore`, `.gitattributes`.
- Вне git: `*.xlsx` (генерируются скриптами), `dist/` (PyInstaller), `.env`.

## Правила на сессию
1. **Перед началом работы — всегда `git pull --ff-only origin main`** либо сверка `python verify_repo.py main`. Локальная копия быстро устаревает.
2. Важные паттерны, решения и находки записывать в `tracker.md`.
3. Задачи, долг и «надо доделать» — в `techdebt.md`.
4. Секреты не коммитить; `.env` в gitignore.
5. Перед пушем проверять `git status` и историю (`git log --oneline -3`).
6. Работать сразу в `main`; содержательный текст из ответа = файл → commit → push → verify.

## Быстрые команды
```bash
source ~/.bashrc                 # подтянуть token_git/GITHUB_TOKEN
git push -u origin main          # запустить в репо
cat /workspace/.env              # значение токена (не показывать наружу без нужды)
```

## Инцидент «не вижу обновление» (2026-09-26)
- Причина: реестр TD-01…TD-06 был составлен в ответе чата, но НЕ был записан в файл/коммит — репо обновлённым не выглядел.
- Исправлено: techdebt.md дополнен полным реестром + тройной валидацией + рекомендациями; коммит 888c4ac запушен в feature/wls-v3; verify_repo.py подтвердил совпадение sha всех файлов.
- Правило на будущее: любой содержательный текст из ответа = сразу в файл → commit → push → verify. Не считать работу сделанной без пуша.
- Примечание: PR #1 уже смержен в main (210b92c), но мерж был ДО этого коммита — для попадания в main нужен новый PR или merge feature/wls-v3 → main.

## Урок «локальная копия устарела, файлы были НЕ из репо» (2026-10-08)
- Симптом: пользователь просил «вытащить файлы из репо», но `/workspace` содержал только состояние ~v4; `verify_repo.py` сверял feature/wls-v3 и показывал «всё сходится», хотя main давно ушёл вперёд.
- Причина: работа велась в feature-ветке, список EXPECTED был захардкожен под неё; ветка `main` в remote не отслеживалась.
- Исправление: синхронизирована локальная ветка `main` с `origin/main` (db629c9, v6.0 GUI) — теперь в `/workspace` лежат **фактические файлы из репозитория** (gui.py, cal_app.py, calc_engine.py, validate_*.py, mc_v4.py, report_v4.py, тесты); EXPECTED в verify_repo.py приведён к 23 файлам main; правило: работать в main, перед стартом pull/сверка.
- Ключевое: скрипт сверки обязан проверять ту ветку, которая является источником истины (main), а не историческую feature-ветку.
