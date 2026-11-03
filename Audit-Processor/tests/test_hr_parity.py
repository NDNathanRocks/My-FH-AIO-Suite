"""The merged HR toolbox engine must match the original v1.3 engine."""
import importlib.util
import io
import sys
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from audit_suite.hr_toolbox import engine as new

LEGACY_PATH = Path(__file__).resolve().parents[2] / "HR Data Toolbox Web v1.3" / "engine.py"


def _legacy():
    spec = importlib.util.spec_from_file_location("hr_legacy_engine", LEGACY_PATH)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["hr_legacy_engine"] = mod
    spec.loader.exec_module(mod)
    return mod


legacy = pytest.importorskip("pandas") and (_legacy() if LEGACY_PATH.exists() else None)

WORKING = pd.DataFrame({
    "EEID": ["1", "2", "3", "4", None, "6", "7.0", " 8 "],
    "Location": ["PC RMH", "pc  rmh", "SMH", "XYZ", "SMH", "012345", "ERH", "ERH"],
    "Dept": ["A", "B", "C", "D", "E", "F", "G", "H"],
})
LOOKUP1 = pd.DataFrame({"Employee": ["1", "2", "2", "3", "7", "8"], "Name": ["Ann", "Bo", "Bob", None, "Gi", "Hal"],
                        "Dept": ["Z", "Y", "Y", "X", "W", "V"]})
LOOKUP2 = pd.DataFrame({"Site": ["PC RMH", "SMH", "012345 - PC RMH", "012345 - other", "ERH"],
                        "Region": ["Fraser North", "Fraser South", "Fraser North", "Fraser East", None]})


def _specs(mode2="soft", conflict="separate"):
    return [
        {"name": "people.xlsx", "data": LOOKUP1, "input_keys": ["EEID"], "keys": ["Employee"], "match_mode": "standardized",
         "fields": [{"source": "Name", "output": "Name", "conflict": "separate"}, {"source": "Dept", "output": "Dept", "conflict": conflict}]},
        {"name": "sites.csv", "data": LOOKUP2, "input_keys": ["Location"], "keys": ["Site"], "match_mode": mode2,
         "fields": [{"source": "Region", "output": "Region", "conflict": "separate"}]},
    ]


def _sheet_values(payload):
    wb = openpyxl.load_workbook(io.BytesIO(payload))
    return {ws.title: [[c for c in r] for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


@pytest.mark.skipif(legacy is None, reason="legacy HR toolbox removed")
@pytest.mark.parametrize("policy", ["list", "first", "last"])
@pytest.mark.parametrize("mode2", ["exact", "standardized", "soft"])
@pytest.mark.parametrize("conflict", ["separate", "first_nonblank", "combine"])
@pytest.mark.parametrize("unique_only", [True, False])
def test_engine_parity(policy, mode2, conflict, unique_only):
    kw = dict(duplicate_policy=policy, delimiter="; ", unique_only=unique_only, include_reports=True)
    a = legacy.process(WORKING, _specs(mode2, conflict), **kw)
    b = new.process(WORKING, _specs(mode2, conflict), **kw)
    pd.testing.assert_frame_equal(a.enriched, b.enriched)
    pd.testing.assert_frame_equal(a.unmatched, b.unmatched)
    pd.testing.assert_frame_equal(a.file_results, b.file_results)
    pd.testing.assert_frame_equal(a.summary, b.summary)
    assert (a.matched, a.unmatched_count, a.duplicate_keys) == (b.matched, b.unmatched_count, b.duplicate_keys)
    assert _sheet_values(a.output) == _sheet_values(b.output)


@pytest.mark.skipif(legacy is None, reason="legacy HR toolbox removed")
def test_composite_and_case_sensitive_parity():
    lk = pd.DataFrame({"E": ["1", "1", "2"], "D": ["2026-01-05", "2026/1/6", "2026-01-05"], "V": ["a", "b", "c"]})
    w = pd.DataFrame({"EEID": ["1", "1", "2", "2"], "Day": [pd.Timestamp("2026-01-05"), "2026-01-06", "2026-01-05", None]})
    spec = [{"name": "k", "data": lk, "input_keys": ["EEID", "Day"], "keys": ["E", "D"], "match_mode": "exact",
             "fields": [{"source": "V", "output": "V"}]}]
    for cs in (True, False):
        a = legacy.process(w, spec, case_sensitive=cs)
        b = new.process(w, spec, case_sensitive=cs)
        pd.testing.assert_frame_equal(a.enriched, b.enriched)


@pytest.mark.skipif(legacy is None, reason="legacy HR toolbox removed")
def test_csv_and_xlsx_load_parity(tmp_path):
    csv = io.BytesIO("EEID,Name\n1,Ann\n2,Bé\n".encode("utf-8")); csv.name = "w.csv"
    x = io.BytesIO(); WORKING.to_excel(x, index=False); x.seek(0); x.name = "w.xlsx"
    pd.testing.assert_frame_equal(legacy.load(csv), new.load(csv))
    assert legacy.sheets(x) == new.sheets(x)
    pd.testing.assert_frame_equal(legacy.load(x, "Sheet1"), new.load(x, "Sheet1"))
