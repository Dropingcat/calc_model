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
    "README.md",
    "cal_v2.py",
    "cal_v3.py",
    "Metrology_Core_EURACHEM_v2.xlsx",
    "Metrology_Core_EURACHEM_v3.xlsx",
    "memory.md",
    "tracker.md",
    "techdebt.md",
    ".gitignore",
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
    branch = sys.argv[1] if len(sys.argv) > 1 else "feature/wls-v3"
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

    print("\n[2] Ожидаемые файлы в HEAD ветки:")
    for f in EXPECTED:
        here_r, here_l = f in remote_files, os.path.exists(f)
        status = "OK " if (here_r and here_l) else "!! "
        sizes = ""
        if here_r and here_l:
            lsha = local_blob_sha(f)
            match = "sha совпадает" if lsha == remote_files[f] else "sha РАЗЛИЧАЕТСЯ (нужен pull/push)"
            if lsha != remote_files[f]:
                ok = False
            sizes = f"  [{match}, {os.path.getsize(f)} б]"
        print(f"    {status}{f:<40} remote:{'есть' if here_r else 'НЕТ'}  local:{'есть' if here_l else 'НЕТ'}{sizes}")
        if not (here_r and here_l):
            ok = False

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
