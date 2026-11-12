"""
Run the data-quality checks defined in config/rules.yaml.

Outputs
  reports/dq_report_<timestamp>.xlsx  Summary sheet + one sheet of failing rows per failed check
  logs/dq_history.csv                 One row per check per run (appended)
  logs/run_log.txt                    Plain-text log of each run

Exit codes
  0  run completed (checks may still have failed; see the report)
  2  run could not complete (missing config, missing data file, missing column)

Run from the project folder:
    python src/validate.py
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parents[1]


class ConfigError(Exception):
    pass


# ---------------------------------------------------------------- checks
# Each check returns a boolean Series: True = row FAILS the check.

def check_not_null(df, c, tables):
    col = c["column"]
    return df[col].isna() | (df[col].astype("string").str.strip() == "")


def check_allowed_values(df, c, tables):
    col = c["column"]
    return df[col].notna() & ~df[col].isin(c["values"])


def check_regex(df, c, tables):
    col = c["column"]
    pattern = re.compile(c["pattern"])
    return df[col].notna() & ~df[col].astype("string").apply(lambda v: bool(pattern.match(str(v))))


def check_range(df, c, tables):
    col = c["column"]
    values = pd.to_numeric(df[col], errors="coerce")
    fail = values.isna() & df[col].notna()  # not a number at all
    if "min" in c:
        fail |= (values < c["min"]) if c.get("min_inclusive", True) else (values <= c["min"])
    if "max" in c:
        fail |= (values > c["max"]) if c.get("max_inclusive", True) else (values >= c["max"])
    return fail.fillna(False)


def check_duplicate_key(df, c, tables):
    # Flags every repeat after the first occurrence of a key.
    return df[c["column"]].duplicated(keep="first")


def check_duplicate_value(df, c, tables, key=None):
    # Same value under more than one distinct key (e.g. one company, two account IDs).
    # Flags the keys after the first one (sorted) for each value.
    col = c["column"]
    pairs = df[[col, key]].dropna().drop_duplicates()
    pairs = pairs.sort_values([col, key])
    extra_keys = set(pairs[pairs.duplicated(subset=[col], keep="first")][key])
    return df[key].isin(extra_keys)


def check_foreign_key(df, c, tables):
    ref = tables[c["ref_table"]]
    valid = set(ref[c["ref_column"]].dropna())
    col = c["column"]
    return df[col].notna() & ~df[col].isin(valid)


def check_date_order(df, c, tables):
    start = pd.to_datetime(df[c["start_column"]], errors="coerce")
    end = pd.to_datetime(df[c["end_column"]], errors="coerce")
    return (start.notna() & end.notna() & (end < start))


CHECKS = {
    "not_null": check_not_null,
    "allowed_values": check_allowed_values,
    "regex": check_regex,
    "range": check_range,
    "duplicate_key": check_duplicate_key,
    "duplicate_value": check_duplicate_value,
    "foreign_key": check_foreign_key,
    "date_order": check_date_order,
}


def columns_needed(c):
    return [c[k] for k in ("column", "start_column", "end_column") if k in c]


# ---------------------------------------------------------------- engine

def load_config(path):
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")
    with open(path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    if not cfg or "tables" not in cfg:
        raise ConfigError("Config has no 'tables' section")
    return cfg


def load_tables(cfg, root):
    tables = {}
    for name, t in cfg["tables"].items():
        path = root / t["file"]
        if not path.exists():
            raise ConfigError(f"Data file not found for table '{name}': {path}")
        # Only truly blank cells count as missing. pandas would otherwise turn text
        # such as "N/A" or "NA" into missing values and hide them from value checks.
        df = pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[""])
        if len(df) == 0:
            raise ConfigError(f"Data file for table '{name}' has no rows: {path}")
        tables[name] = df
    return tables


def run_checks(cfg, tables):
    """Return (summary DataFrame, {check_name: failing rows DataFrame})."""
    summary, failures = [], {}
    for tname, t in cfg["tables"].items():
        df = tables[tname]
        key = t.get("key")
        for c in t["checks"]:
            missing = [col for col in columns_needed(c) if col not in df.columns]
            if missing:
                raise ConfigError(f"{tname}.{c['name']}: column(s) not in file: {missing}")
            if c["type"] not in CHECKS:
                raise ConfigError(f"{tname}.{c['name']}: unknown check type '{c['type']}'")
            if c["type"] == "duplicate_value":
                fail = check_duplicate_value(df, c, tables, key=key)
            else:
                fail = CHECKS[c["type"]](df, c, tables)
            n_fail = int(fail.sum())
            status = "PASS" if n_fail == 0 else ("FAIL" if c.get("severity", "error") == "error" else "WARN")
            summary.append({
                "table": tname,
                "check": c["name"],
                "type": c["type"],
                "column": ", ".join(columns_needed(c)),
                "severity": c.get("severity", "error"),
                "rows_checked": len(df),
                "rows_failed": n_fail,
                "fail_rate_pct": round(100 * n_fail / len(df), 2),
                "status": status,
            })
            if n_fail:
                failures[c["name"]] = df[fail].copy()
    return pd.DataFrame(summary), failures


def write_report(summary, failures, run_id, out_dir):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    out_dir.mkdir(exist_ok=True)
    path = out_dir / f"dq_report_{run_id}.xlsx"
    fills = {"PASS": "C6EFCE", "WARN": "FFEB9C", "FAIL": "FFC7CE"}

    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Summary", index=False)
        for name, rows in failures.items():
            rows.to_excel(xw, sheet_name=name[:31], index=False)

        for ws in xw.book.worksheets:
            for cell in ws[1]:
                cell.font = Font(bold=True)
            for i, col in enumerate(ws.columns, start=1):
                width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 10), 45)
            ws.freeze_panes = "A2"

        ws = xw.book["Summary"]
        status_col = list(summary.columns).index("status") + 1
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=status_col)
            cell.fill = PatternFill("solid", fgColor=fills.get(cell.value, "FFFFFF"))
    return path


def append_history(summary, run_id, run_time, log_dir):
    log_dir.mkdir(exist_ok=True)
    path = log_dir / "dq_history.csv"
    hist = summary[["table", "check", "severity", "rows_checked", "rows_failed", "status"]].copy()
    hist.insert(0, "run_time", run_time)
    hist.insert(0, "run_id", run_id)
    hist.to_csv(path, mode="a", header=not path.exists(), index=False)
    return path


def log(msg, log_dir):
    log_dir.mkdir(exist_ok=True)
    line = f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}"
    print(line)
    with open(log_dir / "run_log.txt", "a", encoding="utf-8") as f:
        f.write(line + "\n")


def main(argv=None):
    p = argparse.ArgumentParser(description="Run data-quality checks.")
    p.add_argument("--config", default="config/rules.yaml")
    p.add_argument("--root", default=str(ROOT), help="Project folder (paths in the config are relative to it)")
    args = p.parse_args(argv)

    root = Path(args.root)
    log_dir, report_dir = root / "logs", root / "reports"
    now = datetime.now()
    run_id = now.strftime("%Y%m%d_%H%M%S")

    try:
        cfg = load_config(root / args.config)
        tables = load_tables(cfg, root)
        summary, failures = run_checks(cfg, tables)
    except ConfigError as e:
        log(f"RUN {run_id} ERROR: {e}", log_dir)
        return 2

    report = write_report(summary, failures, run_id, report_dir)
    append_history(summary, run_id, now.strftime("%Y-%m-%d %H:%M:%S"), log_dir)

    n_fail = int((summary["status"] == "FAIL").sum())
    n_warn = int((summary["status"] == "WARN").sum())
    n_pass = int((summary["status"] == "PASS").sum())
    log(f"RUN {run_id} OK: {len(summary)} checks, {n_pass} pass, {n_warn} warn, {n_fail} fail. Report: {report.name}", log_dir)
    print()
    print(summary[["table", "check", "rows_failed", "status"]].to_string(index=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
