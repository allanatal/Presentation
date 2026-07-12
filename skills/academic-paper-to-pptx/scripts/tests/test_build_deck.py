import os, sys, json, copy
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import build_deck as bd

CLAIMS = {
    "deck": {"title": "T", "source": "p.pdf", "nct": "NCT00000000", "date": "2026-07-12"},
    "claims": [
        {"id": "eff-01", "category": "efficacy", "label": "Median OS IP",
         "value": "19.4 months (95% CI 17.1-22.9)", "slide": None, "source_hint": "Fig 2A"},
        {"id": "eff-03", "category": "efficacy", "label": "OS HR",
         "value": "HR 0.67 (95% CI 0.50-0.90); P = .01", "slide": None, "source_hint": "Fig 2A"},
        {"id": "dem-01", "category": "demographics", "label": "Age",
         "value": "IP 60 (24-70); PS 56 (23-74)", "slide": None, "source_hint": "Table 1"},
    ],
}


# ═══════════ Task 1: Claims core ═══════════

def test_value_lookup():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    assert c.value("eff-01") == "19.4 months (95% CI 17.1-22.9)"


def test_unknown_claim_raises():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    with pytest.raises(KeyError):
        c.value("nope-99")


def test_substitute_expands_and_assigns():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    out = c.substitute("Median OS {{eff-01}}; {{eff-03}}", slide_no=10)
    assert out == "Median OS 19.4 months (95% CI 17.1-22.9); HR 0.67 (95% CI 0.50-0.90); P = .01"
    assert c.referenced["eff-01"] == 10
    assert c.referenced["eff-03"] == 10


def test_substitute_unknown_token_raises():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    with pytest.raises(KeyError):
        c.substitute("bad {{nope-99}}", slide_no=1)


def test_explicit_assign():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("dem-01", 8)
    assert c.referenced["dem-01"] == 8


def test_first_assignment_wins():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("eff-01", 5)
    c.assign("eff-01", 9)
    assert c.referenced["eff-01"] == 5


def test_resolved_json_fills_slide_and_preserves_original():
    data = copy.deepcopy(CLAIMS)
    c = bd.Claims(data)
    c.assign("eff-01", 10)
    c.assign("dem-01", 8)
    resolved = c.resolved_json()
    by = {x["id"]: x for x in resolved["claims"]}
    assert by["eff-01"]["slide"] == 10
    assert by["dem-01"]["slide"] == 8
    assert by["eff-03"]["slide"] is None
    assert data["claims"][0]["slide"] is None  # original untouched


def test_unreferenced_lists_unused_ids():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    c.assign("eff-01", 10)
    assert set(c.unreferenced()) == {"eff-03", "dem-01"}


# ═══════════ Task 2: recursive spec substitution ═══════════

def test_substitute_in_place_walks_nested_structures():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    sdata = {
        "type": "figure",
        "key_message": "OS {{eff-01}}",
        "caption": ["deaths line", "HR is {{eff-03}}"],
        "rows": [{"label": "Age", "vals": ["{{dem-01}}"]}],
    }
    bd.substitute_in_place(sdata, c, slide_no=10)
    assert sdata["key_message"] == "OS 19.4 months (95% CI 17.1-22.9)"
    assert sdata["caption"][1] == "HR is HR 0.67 (95% CI 0.50-0.90); P = .01"
    assert sdata["rows"][0]["vals"][0] == "IP 60 (24-70); PS 56 (23-74)"
    assert c.referenced == {"eff-01": 10, "eff-03": 10, "dem-01": 10}


def test_slide_level_claims_assigned():
    c = bd.Claims(copy.deepcopy(CLAIMS))
    sdata = {"type": "table", "claims": ["dem-01"], "rows": []}
    bd.assign_slide_claims(sdata, c, slide_no=8)
    assert c.referenced["dem-01"] == 8
