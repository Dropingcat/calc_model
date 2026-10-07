# MEMORY — памятка ассистенту по сессии (calc_model)

_Обновлено: 2026-10-08_

## Окружение и git
- Рабочий каталог: `/workspace` (git-репозиторий, активная ветка `feature/wls-v3`).
- Внешний репо: `origin → https://github.com/Dropingcat/calc_model.git`.
- Токен GitHub:
  - доступен как переменная окружения `token_git` (продублирован в `GITHUB_TOKEN`);
  - прописан в `~/.bashrc` (`export token_git=...`) — после `source ~/.bashrc` виден в shell;
  - хранится в `/workspace/.env` (права 600, добавлен в `.gitignore`);
  - для git-аутентификации настроен `credential.helper store` → `/root/.git-credentials` (права 600).
  - **В URL remote токен НЕ вшивать** — credentials подставляются автоматически.
- Проверка доступа: `git ls-remote origin` работает без интерактивного ввода.
- Push/pull работают «из коробки»: `git push -u origin main` — без запроса учётных данных.

## Состояние репозитория (снимок 2026-10-08)
- HEAD: `e5c3e60` «v4: cal_v4.py (WLS+LoF+Welch-Satterthwaite), .gitignore» — синхронизирован
  с `origin/feature/wls-v3` И `origin/main` (ветки совпадают).
- Отслеживаемые файлы (git ls-files):
  `README.md`, `cal_v2.py`, `cal_v3.py`, `cal_v4.py`, `memory.md`, `tracker.md`,
  `techdebt.md`, `verify_repo.py`, `.gitignore`.
- Локальные артефакты (НЕ в git, генерируются скриптами, *.xlsx в .gitignore):
  `Metrology_Core_EURACHEM_v2.xlsx`, `_v3.xlsx`, `_v4.xlsx` — пересоздаются `python cal_vN.py`.
- Новых папок нет — плоская структура, один корень `/workspace`.
- Домен: метрология / калибровка по EURACHEM (OLS/WLS, неопределённости, LoF) — см. tracker.md.

## Правила на сессию
1. Важные паттерны, решения и находки записывать в `tracker.md`.
2. Задачи, долг и «надо доделать» — в `techdebt.md`.
3. Секреты не коммитить; `.env` в gitignore.
4. Перед пушем проверять `git status` и историю (`git log --oneline -3`).

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

## Урок «устаревший EXPECTED в verify_repo.py» (2026-10-08)
- Симптом: скрипт сверки показывал расхождение — требовал в remote `*.xlsx` и не знал про `cal_v4.py`.
- Причина: список EXPECTED обслуживался вручную и не обновлялся при изменении состава репо.
- Исправление: EXPECTED приведён к `git ls-files`; `.xlsx` удалены из индекса (`git rm --cached`) — они генерируемые артефакты и уже были в .gitignore (дублирование правила `.env` убрано).
- Правило на будущее: после ЛЮБОГО изменения состава файлов (новый модуль/артефакт) — сразу синхронизировать EXPECTED в verify_repo.py и снимок состояния в memory.md, затем commit → push → verify.
