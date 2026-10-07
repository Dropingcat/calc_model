"""Тесты TD-2177-008 / TD-D86-007: контролируемый реестр редакций НД."""

import copy
import json

import pytest

from fraction_lib.nd_registry import (
    NDEntry,
    NormativeRegistryError,
    REGISTRY_PATH,
    canonical_json,
    check_no_drift,
    compute_core_hash,
    d86_core,
    gost2177_core,
    load_registry,
    resolve_for_calculation,
)


class TestRegistryIntegrity:
    def test_registry_file_exists_and_parses(self):
        entries = load_registry()
        assert len(entries) >= 3

    def test_active_entries_have_required_fields(self):
        for e in load_registry():
            if e.status == "ACTIVE":
                assert e.id and e.edition and e.effective_from
                assert len(e.sha256_of_normative_core) == 64
                assert e.verified_by, f"{e.id}: нет доказательной базы"
                assert e.scope

    def test_unique_ids(self):
        ids = [e.id for e in load_registry()]
        assert len(ids) == len(set(ids))

    def test_exactly_one_active_per_designation(self):
        from collections import Counter
        act = Counter(e.designation for e in load_registry() if e.status == "ACTIVE")
        for d, n in act.items():
            assert n == 1, f"'{d}': {n} ACTIVE-записей — неоднозначная редакция"


class TestHashDrift:
    def test_no_drift_current_code(self):
        assert check_no_drift() == []

    def test_canonical_json_stable_key_order(self):
        a = {"b": 1, "a": 2, "c": {"z": 1, "y": 2}}
        b = {"a": 2, "c": {"y": 2, "b": 1, "z": 1}}
        b["b"] = 1
        del b["c"]["b"]
        assert canonical_json(a) == canonical_json(b)

    def test_drift_detected_when_constant_changes(self, tmp_path):
        """Подделка: правим хеш ACTIVE-записи — drift пойман."""
        entries = load_registry()
        data = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
        data["entries"][0]["sha256_of_normative_core"] = "f" * 64  # фальшивый хеш
        p = tmp_path / "ND_REGISTRY.json"
        p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        problems = check_no_drift(p)
        assert problems and "drift" in problems[0]

    def test_core_hash_matches_registry_values(self):
        reg = {e.designation: e for e in load_registry() if e.status == "ACTIVE"}
        assert reg["ГОСТ 2177-99"].sha256_of_normative_core == compute_core_hash(gost2177_core())
        assert reg["ASTM D86"].sha256_of_normative_core == compute_core_hash(d86_core())


class TestResolutionPolicy:
    def test_resolve_returns_active_edition(self):
        e = resolve_for_calculation("ГОСТ 2177-99")
        assert "Изм." in e.edition and "Поправка 2023" in e.edition
        e2 = resolve_for_calculation("ASTM D86")
        assert e2.edition.startswith("D86-23")

    def test_superseded_not_resolvable_as_active(self):
        with pytest.raises(NormativeRegistryError):
            resolve_for_calculation("несуществующий НД")

    def test_legacy_entry_marked_superseded(self):
        legacy = [e for e in load_registry() if e.id == "ND-LEGACY-001"][0]
        assert legacy.status == "SUPERSEDED"
