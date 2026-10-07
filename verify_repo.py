#!/usr/bin/env python3
"""Сверка локального рабочего дерева с внешним репозиторием GitHub.

Проверяет по API GitHub (без учёта локального кэша):
  1) список веток origin;
  2) наличие и контрольные суммы ожидаемых файлов в HEAD указанной ветки;
  3) сравнение sha1 blob'ов remote vs локальных файлов.

Использование:
  python verify_repo.py [branch]      # по умолчанию feature/wls-v3
Токен берётся из окружения (token_git / GITHUB_TOKEN) либо из git credential store.
"""
import hashlib
import json
import os
import subprocess
import sys
import urllib.request

OWNER = "Dropingcat"
REPO = "calc_model"
EXPECTED = [
    ".gitattributes",
    ".gitignore",
    "README.md",
    "cal_app.py",
    "cal_v2.py",
    "cal_v3.py",
    "cal_v4.py",
    "calc_engine.py",
    "e2e_test_v4.py",
    "gui.py",
    "lib_validate.py",
    "mc_v4.py",
    "memory.md",
    "report_v4.py",
    "techdebt.md",
    "test_gui.py",
    "test_td07.py",
    "tracker.md",
    "validate_auto.py",
    "validate_v4.py",
    "verify_repo.py",
]
FORBIDDEN = [".env"]  # секреты в репо быть не должно


def get_token() -> str:
    for var in ("token_git", "GITHUB_TOKEN"):
        t = os.environ.get(var, "").strip()
        if t:
            return t
    # fallback: git credential store
    p = subprocess.run(
        ["git", "credential-fill"],
        input="protocol=https\nhost=github.com\n\n",
        capture_output=True, text=True,
    )
    for line in p.stdout.splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1].strip()
    return ""


def api(url: str, token: str) -> dict:
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}" if token else "",
        "Accept": "application/vnd.github+json",
        "User-Agent": "verify-repo-script",
    })
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)


def local_blob_sha(path: str) -> str:
    """git hash-object — канонический sha1 blob'а."""
    out = subprocess.run(
        ["git", "hash-object", "--", path],
        capture_output=True, text=True, cwd=os.path.dirname(os.path.abspath(__file__)),
    )
    return out.stdout.strip()


def main() -> int:
    branch = sys.argv[1] if len(sys.argv) > 1 else "main"
    token = get_token()
    ok = True

    print(f"== Сверка с https://github.com/{OWNER}/{REPO} (ветка: {branch}) ==")
    print(f"Токен: {'найден (' + token[:4] + '...)' if token else 'НЕ найден — запрос без авторизации'}")

    # 1. Ветки
    branches = {b["name"]: b["commit"]["sha"] for b in
                api(f"https://api.github.com/repos/{OWNER}/{REPO}/branches", token)}
    print("\n[1] Ветки в репо:")
    for name, sha in sorted(branches.items()):
        print(f"    {name:<25} {sha[:8]}")
    if branch not in branches:
        print(f"    !! Ветка '{branch}' ОТСУТСТВУЕТ"); ok = False
    else:
        print(f"    OK: ветка '{branch}' присутствует")

    # 2-3. Файлы в HEAD ветки vs локальные
    tree = api(f"https://api.github.com/repos/{OWNER}/{REPO}/git/trees/{branches.get(branch, 'main')}?recursive=1", token)
    remote_files = {t["path"]: t["sha"] for t in tree.get("tree", []) if t["type"] == "blob"}

    # 2a. Явный список (корневые файлы) + 2b. полный обход git-индекса
    tracked = subprocess.run(["git", "ls-files", "-z"], capture_output=True, text=True).stdout.split("\0")
    tracked = [t for t in tracked if t and t not in EXPECTED]
    all_files = EXPECTED + sorted(tracked)

    print(f"\n[2] Файлы git-индекса vs HEAD ветки (всего {len(all_files)}):")
    mismatches, missing = [], []
    for f in all_files:
        here_r, here_l = f in remote_files, os.path.exists(f)
        if here_r and here_l:
            lsha = local_blob_sha(f)
            if lsha != remote_files[f]:
                mismatches.append(f)
        elif not here_r:
            missing.append((f, "нет в remote"))
        elif not here_l:
            missing.append((f, "нет локально"))
    for f in mismatches:
        print(f"    !! {f}: sha РАЗЛИЧАЕТСЯ (нужен pull/push)")
    for f, why in missing:
        print(f"    !! {f}: {why}")
        ok = False
    if mismatches:
        ok = False
    if not (mismatches or missing):
        print(f"    OK: все {len(all_files)} файлов присутствуют с обеих сторон, sha1 совпадают")

    print("\n[3] Секреты (не должны быть в репо):")
    for f in FORBIDDEN:
        leaked = f in remote_files
        print(f"    {'!! УТЕЧКА: ' if leaked else 'OK:   '}{f} в remote HEAD {'найден!' if leaked else 'отсутствует'}")
        if leaked:
            ok = False

    print("\n== ИТОГ:", "ВСЁ СХОДИТСЯ ✅" if ok else "ЕСТЬ РАСХОЖДЕНИЯ ❌", "==")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
