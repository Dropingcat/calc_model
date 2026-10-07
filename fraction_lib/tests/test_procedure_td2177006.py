"""TD-2177-006: процедурный checklist ГОСТ 2177 пп. 5.1–5.5."""
import pytest

from fraction_lib.procedure import (
    CheckItem,
    build_checklist,
    check_condensation_rate,
    evaluate_checklist,
)


def _fill_all(items, answer=True):
    for it in items:
        if it.id != "5.4c":
            it.answer = answer
    return items


class TestChecklistTD2177006:
    def test_fresh_checklist_is_missing(self):
        items = build_checklist("A")
        verdict = evaluate_checklist(items)
        assert verdict["verdict"] == "FAIL"
        assert len(verdict["missing"]) == len(items)

    def test_all_ok_passes(self):
        items = _fill_all(build_checklist("A"), True)
        check_condensation_rate(4.5, items=items)
        verdict = evaluate_checklist(items)
        assert verdict["verdict"] == "PASS"

    def test_warning_only_is_conditional(self):
        items = _fill_all(build_checklist("A"), True)
        check_condensation_rate(4.2, items=items)
        # пункт с severity=WARNING отвечаем False
        [i for i in items if i.id == "5.2b"][0].answer = False
        verdict = evaluate_checklist(items)
        assert verdict["verdict"] == "CONDITIONAL"
        assert "5.2b" in verdict["warning"]

    def test_invalid_clause_blocks(self):
        items = _fill_all(build_checklist("A"), True)
        check_condensation_rate(4.5, items=items)
        [i for i in items if i.id == "5.4b"][0].answer = False  # INVALID-класс
        verdict = evaluate_checklist(items)
        assert verdict["verdict"] == "FAIL"
        assert "5.4b" in verdict["invalid"]

    def test_condensation_rate_boundaries(self):
        assert check_condensation_rate(4.0) is True
        assert check_condensation_rate(5.0) is True
        assert check_condensation_rate(3.9) is False
        assert check_condensation_rate(5.1) is False

    def test_method_b_variant_annotated(self):
        items = build_checklist("B")
        it = [i for i in items if i.id == "5.4b"][0]
        assert "метод Б" in it.requirement

    def test_status_property_semantics(self):
        it = CheckItem("x", "требование")          # None -> MISSING
        assert it.status == "MISSING"
        it.answer = True
        assert it.status == "OK"
        it.answer = False
        assert it.status == "INVALID"
        it.severity = "WARNING"
        assert it.status == "WARNING"
