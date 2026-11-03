import io
import time

import numpy as np
import openpyxl
import pandas as pd
import pytest

from audit_suite.common.files import InputFileError
from audit_suite.hr_toolbox import engine


def spec(data, wk="EEID", lk="EEID", fields=("Name",), mode="standardized"):
    return [{"name": "lk", "data": data, "input_keys": [wk], "keys": [lk], "match_mode": mode,
             "fields": [{"source": f, "output": f} for f in fields]}]


def test_empty_working_file():
    lk = pd.DataFrame({"EEID": ["1"], "Name": ["A"]})
    res = engine.process(pd.DataFrame({"EEID": pd.Series([], dtype=object)}), spec(lk))
    assert len(res.enriched) == 0 and res.matched == 0
    assert res.summary.loc[3, "Value"] == "0.0%"


def test_blank_ids_unmatched_and_rows_preserved():
    w = pd.DataFrame({"EEID": ["1", None, "", "  ", "2"]})
    lk = pd.DataFrame({"EEID": ["1", None, "2"], "Name": ["A", "ghost", "B"]})
    res = engine.process(w, spec(lk))
    assert len(res.enriched) == 5
    assert list(res.enriched["Name"].fillna("-")) == ["A", "-", "-", "-", "B"]
    assert res.unmatched_count == 3 and res.blank_key_rows == 1


def test_duplicate_policy_error():
    lk = pd.DataFrame({"EEID": ["1", "1"], "Name": ["A", "B"]})
    with pytest.raises(ValueError, match="duplicate identifier"):
        engine.process(pd.DataFrame({"EEID": ["1"]}), spec(lk), duplicate_policy="error")


def test_missing_field_and_mapping_errors():
    lk = pd.DataFrame({"EEID": ["1"], "Name": ["A"]})
    with pytest.raises(ValueError, match="not found"):
        engine.process(pd.DataFrame({"EEID": ["1"]}), spec(lk, fields=("Nope",)))
    bad = spec(lk); bad[0]["keys"] = []
    with pytest.raises(ValueError, match="map every"):
        engine.process(pd.DataFrame({"EEID": ["1"]}), bad)


def test_duplicate_headings_and_corrupt_files():
    f = io.BytesIO(b"A,A \n1,2\n"); f.name = "d.csv"
    with pytest.raises(InputFileError, match="Duplicate column headings"):
        engine.load(f)
    f = io.BytesIO(b"notazip"); f.name = "d.xlsx"
    with pytest.raises(InputFileError):
        engine.sheets(f)
    f = io.BytesIO(b""); f.name = "e.csv"
    with pytest.raises(InputFileError, match="empty"):
        engine.load(f)


def test_soft_match_flags_multiple_keys():
    lk = pd.DataFrame({"Site": ["012345 - PC RMH", "012345 - Other"], "Region": ["N", "E"]})
    w = pd.DataFrame({"Loc": ["012345"]})
    res = engine.process(w, spec(lk, "Loc", "Site", ("Region",), "soft"))
    assert res.enriched.loc[0, "Match Status - lk"] == "Soft match (2 keys)"
    assert res.enriched.loc[0, "Region"] == "N, E" and res.soft_matches == 1


def test_unicode_values_round_trip():
    lk = pd.DataFrame({"EEID": ["1"], "Name": ["Zoë Łukasz 漢字"]})
    res = engine.process(pd.DataFrame({"EEID": ["1"]}), spec(lk))
    ws = openpyxl.load_workbook(io.BytesIO(res.output))["Enriched Data"]
    assert ws["C2"].value == "Zoë Łukasz 漢字"


def test_large_file_performance():
    n = 50_000
    rng = np.random.default_rng(1)
    w = pd.DataFrame({"EEID": rng.integers(0, 20_000, n).astype(str)})
    lk = pd.DataFrame({"EEID": np.arange(0, 20_000).astype(str).repeat(2), "Name": [f"N{i}" for i in range(40_000)]})
    start = time.perf_counter()
    res = engine.process(w, spec(lk))
    elapsed = time.perf_counter() - start
    assert len(res.enriched) == n and res.matched == n
    assert elapsed < 60, elapsed
