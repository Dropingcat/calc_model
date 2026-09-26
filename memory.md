# MEMORY — памятка ассистенту по сессии (calc_model)

_Обновлено: 2026-09-25_

## Окружение и git
- Рабочий каталог: `/workspace` (git-репозиторий, ветка `main`).
- Внешний репо: `origin → https://github.com/Dropingcat/calc_model.git`.
- Токен GitHub:
  - доступен как переменная окружения `token_git` (продублирован в `GITHUB_TOKEN`);
  - прописан в `~/.bashrc` (`export token_git=...`) — после `source ~/.bashrc` виден в shell;
  - хранится в `/workspace/.env` (права 600, добавлен в `.gitignore`);
  - для git-аутентификации настроен `credential.helper store` → `/root/.git-credentials` (права 600).
  - **В URL remote токен НЕ вшивать** — credentials подставляются автоматически.
- Проверка доступа: `git ls-remote origin` работает без интерактивного ввода.
- Push/pull работают «из коробки»: `git push -u origin main` — без запроса учётных данных.

## Состояние репозитория
- HEAD: `528f60f "first commit"` (README с `# calc_model`), синхронизирован с `origin/main`.
- Файлы проекта: `cal_v2.py`, `Metrology_Core_EURACHEM_v2.xlsx`, `README.md`.
- Домен файлов: метрология / расчёты по EURACHEM (см. tracker.md).

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
