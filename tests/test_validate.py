"""
Tests for the data-quality engine (src/validate.py).

Run from the project folder:
    python -m pytest -v
"""
import shutil
from pathlib import Path

import pandas as pd
import pytest

import validate as v

ROOT = Path(__file__).resolve().parents[1]


# ---------------------------------------------------------------- helpers

def make_project(tmp_path, accounts, opps):
    """Create a small project folder with the real rules and the given data."""
    (tmp_path / "config").mkdir()
    (tmp_path / "data").mkdir()
    shutil.copy(ROOT / "config" / "rules.yaml", tmp_path / "config" / "rules.yaml")
    accounts.to_csv(tmp_path / "data" / "accounts.csv", index=False)
    opps.to_csv(tmp_path / "data" / "opportunities.csv", index=False)
    return tmp_path


def clean_accounts():
    return pd.DataFrame({
        "account_id": ["ACC0001", "ACC0002"],
        "account_name": ["Blue Rail Group", "Cedar Water Co"],
        "industry": ["Transportation", "Water"],
        "region": ["West", "Midwest"],
        "owner_email": ["alex.smith@example.com", "sam.lee@example.com"],
        "status": ["Active", "Prospect"],
        "created_date": ["2024-01-10", "2024-02-15"],
    })


def clean_opps():
    return pd.DataFrame({
        "opportunity_id": ["OPP00001", "OPP00002"],
        "account_id": ["ACC0001", "ACC0002"],
        "stage": ["Proposal", "Closed Won"],
        "amount_usd": [50000.0, 120000.0],
        "probability_pct": [50, 100],
        "created_date": ["2024-03-01", "2024-04-01"],
        "close_date": ["2024-05-01", "2024-06-15"],
    })


def run(project):
    cfg = v.load_config(project / "config" / "rules.yaml")
    tables = v.load_tables(cfg, project)
    return v.run_checks(cfg, tables)


# ---------------------------------------------------------------- end-to-end

def test_finds_exactly_the_planted_issues():
    """Every planted issue is found, and nothing else is flagged."""
    cfg = v.load_config(ROOT / "config" / "rules.yaml")
    tables = v.load_tables(cfg, ROOT)
    summary, failures = v.run_checks(cfg, tables)
    expected = pd.read_csv(ROOT / "data" / "expected_issues.csv")

    for check, group in expected.groupby("check"):
        key = cfg["tables"][group["table"].iloc[0]]["key"]
        found = set(failures[check][key]) if check in failures else set()
        assert found == set(group["record_id"]), check

    assert summary["rows_failed"].sum() == len(expected)


def test_clean_data_passes_every_check(tmp_path):
    project = make_project(tmp_path, clean_accounts(), clean_opps())
    summary, failures = run(project)
    assert (summary["status"] == "PASS").all()
    assert failures == {}


# ---------------------------------------------------------------- individual rules

def test_text_na_is_treated_as_a_value_not_as_missing(tmp_path):
    """Regression test: pandas reads "N/A" as missing by default, which hid an invalid status."""
    acc = clean_accounts()
    acc.loc[0, "status"] = "N/A"
    summary, failures = run(make_project(tmp_path, acc, clean_opps()))
    assert list(failures["invalid_status"]["account_id"]) == ["ACC0001"]
    assert "missing_owner_email" not in failures


def test_blank_and_space_only_cells_count_as_missing(tmp_path):
    acc = clean_accounts()
    acc.loc[0, "industry"] = ""
    acc.loc[1, "industry"] = "   "
    _, failures = run(make_project(tmp_path, acc, clean_opps()))
    assert len(failures["missing_industry"]) == 2


def test_zero_amount_fails_but_small_positive_passes(tmp_path):
    opp = clean_opps()
    opp.loc[0, "amount_usd"] = 0
    opp.loc[1, "amount_usd"] = 0.01
    _, failures = run(make_project(tmp_path, clean_accounts(), opp))
    assert list(failures["non_positive_amount"]["opportunity_id"]) == ["OPP00001"]


def test_probability_boundaries(tmp_path):
    opp = pd.concat([clean_opps()] * 2, ignore_index=True)
    opp["opportunity_id"] = ["OPP1", "OPP2", "OPP3", "OPP4"]
    opp["probability_pct"] = [0, 100, -1, 101]
    _, failures = run(make_project(tmp_path, clean_accounts(), opp))
    assert set(failures["probability_out_of_range"]["opportunity_id"]) == {"OPP3", "OPP4"}


def test_duplicate_key_flags_only_the_repeats(tmp_path):
    acc = pd.concat([clean_accounts(), clean_accounts().iloc[[0]]], ignore_index=True)
    _, failures = run(make_project(tmp_path, acc, clean_opps()))
    assert len(failures["duplicate_account_id"]) == 1


def test_same_name_same_id_is_not_a_duplicate_name(tmp_path):
    """An exact copy of a row is a duplicate ID, not the same company under two IDs."""
    acc = pd.concat([clean_accounts(), clean_accounts().iloc[[0]]], ignore_index=True)
    _, failures = run(make_project(tmp_path, acc, clean_opps()))
    assert "duplicate_account_name" not in failures


def test_same_name_different_id_is_flagged_as_warning(tmp_path):
    acc = clean_accounts()
    extra = acc.iloc[[0]].copy()
    extra["account_id"] = "ACC0003"
    acc = pd.concat([acc, extra], ignore_index=True)
    summary, failures = run(make_project(tmp_path, acc, clean_opps()))
    assert list(failures["duplicate_account_name"]["account_id"]) == ["ACC0003"]
    row = summary.set_index("check").loc["duplicate_account_name"]
    assert row["status"] == "WARN"


def test_orphan_account_id(tmp_path):
    opp = clean_opps()
    opp.loc[1, "account_id"] = "ACC9999"
    _, failures = run(make_project(tmp_path, clean_accounts(), opp))
    assert list(failures["orphan_account_id"]["opportunity_id"]) == ["OPP00002"]


def test_close_before_created_ignores_missing_dates(tmp_path):
    opp = clean_opps()
    opp.loc[0, "close_date"] = "2024-02-01"   # before created
    opp.loc[1, "close_date"] = ""             # missing: a different check
    _, failures = run(make_project(tmp_path, clean_accounts(), opp))
    assert list(failures["close_before_created"]["opportunity_id"]) == ["OPP00001"]
    assert list(failures["missing_close_date"]["opportunity_id"]) == ["OPP00002"]


# ---------------------------------------------------------------- error handling

def test_missing_data_file_stops_with_exit_code_2(tmp_path):
    project = make_project(tmp_path, clean_accounts(), clean_opps())
    (project / "data" / "opportunities.csv").unlink()
    assert v.main(["--root", str(project)]) == 2
    assert "Data file not found" in (project / "logs" / "run_log.txt").read_text()


def test_missing_column_stops_with_exit_code_2(tmp_path):
    project = make_project(tmp_path, clean_accounts().drop(columns="status"), clean_opps())
    assert v.main(["--root", str(project)]) == 2


def test_empty_file_stops_with_exit_code_2(tmp_path):
    project = make_project(tmp_path, clean_accounts().iloc[0:0], clean_opps())
    assert v.main(["--root", str(project)]) == 2


# ---------------------------------------------------------------- outputs

def test_each_run_writes_a_report_and_appends_history(tmp_path):
    project = make_project(tmp_path, clean_accounts(), clean_opps())
    assert v.main(["--root", str(project)]) == 0
    assert v.main(["--root", str(project)]) == 0

    history = pd.read_csv(project / "logs" / "dq_history.csv")
    n_checks = 13
    assert len(history) == 2 * n_checks
    assert len(list((project / "reports").glob("dq_report_*.xlsx"))) >= 1
