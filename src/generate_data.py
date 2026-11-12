"""
Generate synthetic Accounts and Opportunities data with planted data-quality errors.

The data is entirely synthetic. Every planted error is recorded in
data/expected_issues.csv so the validation engine and tests can be checked
against a known answer.

Run from the project folder:
    python src/generate_data.py
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEED = 42
N_ACCOUNTS = 500
N_OPPS = 1500

INDUSTRIES = ["Transportation", "Buildings", "Environment", "Energy", "Water", "Mining", "Government"]
REGIONS = ["Northeast", "Southeast", "Midwest", "Southwest", "West"]
ACCOUNT_STATUSES = ["Active", "Inactive", "Prospect"]
STAGES = ["Prospecting", "Qualification", "Proposal", "Negotiation", "Closed Won", "Closed Lost"]

PREFIXES = ["North", "Blue", "Summit", "River", "Granite", "Pioneer", "Harbor", "Cedar", "Silver", "Atlas",
            "Prairie", "Coastal", "Union", "Metro", "Lakeside", "Eagle", "Valley", "Frontier", "Keystone", "Liberty"]
CORES = ["Transit", "Water", "Power", "Civil", "Rail", "Ports", "Energy", "Infrastructure", "Environmental",
         "Utilities", "Engineering", "Development", "Mobility", "Resources", "Works"]
SUFFIXES = ["Authority", "Group", "Partners", "Corp", "Agency", "District", "Holdings", "Co"]
FIRST = ["alex", "jordan", "sam", "taylor", "morgan", "casey", "riley", "jamie", "drew", "avery",
         "quinn", "reese", "parker", "hayden", "rowan", "emerson", "logan", "skyler", "dakota", "kendall"]
LAST = ["smith", "patel", "garcia", "nguyen", "brown", "khan", "lee", "martin", "clark", "lopez"]


def make_accounts(rng):
    names = set()
    while len(names) < N_ACCOUNTS:
        names.add(f"{rng.choice(PREFIXES)} {rng.choice(CORES)} {rng.choice(SUFFIXES)}")
    names = sorted(names)
    rng.shuffle(names)

    owners = [f"{f}.{l}@example.com" for f, l in zip(rng.choice(FIRST, 20), rng.choice(LAST, 20))]
    owners = list(dict.fromkeys(owners))  # keep unique, in order

    created = pd.to_datetime("2023-01-01") + pd.to_timedelta(rng.integers(0, 900, N_ACCOUNTS), unit="D")
    return pd.DataFrame({
        "account_id": [f"ACC{i:04d}" for i in range(1, N_ACCOUNTS + 1)],
        "account_name": names,
        "industry": rng.choice(INDUSTRIES, N_ACCOUNTS),
        "region": rng.choice(REGIONS, N_ACCOUNTS),
        "owner_email": rng.choice(owners, N_ACCOUNTS),
        "status": rng.choice(ACCOUNT_STATUSES, N_ACCOUNTS, p=[0.7, 0.15, 0.15]),
        "created_date": created.strftime("%Y-%m-%d"),
    })


def make_opps(rng, accounts):
    acc = accounts.sample(N_OPPS, replace=True, random_state=SEED).reset_index(drop=True)
    acc_created = pd.to_datetime(acc["created_date"])
    created = acc_created + pd.to_timedelta(rng.integers(1, 400, N_OPPS), unit="D")
    close = created + pd.to_timedelta(rng.integers(14, 240, N_OPPS), unit="D")
    stage = rng.choice(STAGES, N_OPPS, p=[0.15, 0.15, 0.15, 0.1, 0.25, 0.2])
    prob_by_stage = {"Prospecting": 10, "Qualification": 25, "Proposal": 50, "Negotiation": 75,
                     "Closed Won": 100, "Closed Lost": 0}
    return pd.DataFrame({
        "opportunity_id": [f"OPP{i:05d}" for i in range(1, N_OPPS + 1)],
        "account_id": acc["account_id"],
        "stage": stage,
        "amount_usd": (rng.lognormal(mean=11.5, sigma=0.9, size=N_OPPS)).round(-2),
        "probability_pct": [prob_by_stage[s] for s in stage],
        "created_date": created.dt.strftime("%Y-%m-%d"),
        "close_date": close.dt.strftime("%Y-%m-%d"),
    })


def plant_errors(rng, accounts, opps):
    """Plant errors on distinct rows and return (accounts, opps, expected_issues)."""
    issues = []

    def take(pool, n):
        chosen = rng.choice(sorted(pool), n, replace=False)
        pool.difference_update(chosen)
        return sorted(chosen)

    # ---------- Accounts ----------
    a_pool = set(range(N_ACCOUNTS))

    for i in take(a_pool, 10):
        accounts.loc[i, "owner_email"] = np.nan
        issues.append(("accounts", "missing_owner_email", accounts.loc[i, "account_id"]))
    for i in take(a_pool, 5):
        accounts.loc[i, "industry"] = np.nan
        issues.append(("accounts", "missing_industry", accounts.loc[i, "account_id"]))
    bad_status = ["Actve", "active", "ACTIVE ", "Closed", "Pending", "N/A", "Inactve"]
    for i, val in zip(take(a_pool, 7), bad_status):
        accounts.loc[i, "status"] = val
        issues.append(("accounts", "invalid_status", accounts.loc[i, "account_id"]))
    bad_email = ["alex.smith@", "jordan.patel.example.com", "sam garcia@example.com", "riley@@example.com", "taylor.nguyen"]
    for i, val in zip(take(a_pool, 5), bad_email):
        accounts.loc[i, "owner_email"] = val
        issues.append(("accounts", "invalid_email_format", accounts.loc[i, "account_id"]))

    # Same company name under a new account ID (a likely duplicate account).
    dup_name_src = take(a_pool, 6)
    new_rows = []
    next_id = N_ACCOUNTS + 1
    for i in dup_name_src:
        row = accounts.loc[i].copy()
        row["account_id"] = f"ACC{next_id:04d}"
        next_id += 1
        new_rows.append(row)
        issues.append(("accounts", "duplicate_account_name", row["account_id"]))

    # Exact copies of existing rows (same ID and all fields).
    exact_src = take(a_pool, 8)
    for i in exact_src:
        new_rows.append(accounts.loc[i].copy())
        issues.append(("accounts", "duplicate_account_id", accounts.loc[i, "account_id"]))

    accounts = pd.concat([accounts, pd.DataFrame(new_rows)], ignore_index=True)

    # ---------- Opportunities ----------
    o_pool = set(range(N_OPPS))
    valid_ids = set(accounts["account_id"])

    for n, i in enumerate(take(o_pool, 12), start=1):
        orphan = f"ACC{9000 + n:04d}"
        assert orphan not in valid_ids
        opps.loc[i, "account_id"] = orphan
        issues.append(("opportunities", "orphan_account_id", opps.loc[i, "opportunity_id"]))
    for i in take(o_pool, 9):
        created = pd.to_datetime(opps.loc[i, "created_date"])
        opps.loc[i, "close_date"] = (created - pd.Timedelta(days=int(rng.integers(5, 60)))).strftime("%Y-%m-%d")
        issues.append(("opportunities", "close_before_created", opps.loc[i, "opportunity_id"]))
    for i, val in zip(take(o_pool, 8), [0, -500, -12000, 0, -1, 0, -250, -3000]):
        opps.loc[i, "amount_usd"] = val
        issues.append(("opportunities", "non_positive_amount", opps.loc[i, "opportunity_id"]))
    for i, val in zip(take(o_pool, 6), ["Closed-Won", "Won", "Negotation", "Lost", "proposal", "TBD"]):
        opps.loc[i, "stage"] = val
        issues.append(("opportunities", "invalid_stage", opps.loc[i, "opportunity_id"]))
    for i, val in zip(take(o_pool, 5), [110, 150, -10, 200, 101]):
        opps.loc[i, "probability_pct"] = val
        issues.append(("opportunities", "probability_out_of_range", opps.loc[i, "opportunity_id"]))
    for i in take(o_pool, 4):
        opps.loc[i, "close_date"] = np.nan
        issues.append(("opportunities", "missing_close_date", opps.loc[i, "opportunity_id"]))

    dup_opp_src = take(o_pool, 5)
    opps = pd.concat([opps, opps.loc[dup_opp_src]], ignore_index=True)
    for i in dup_opp_src:
        issues.append(("opportunities", "duplicate_opportunity_id", opps.loc[i, "opportunity_id"]))

    expected = pd.DataFrame(issues, columns=["table", "check", "record_id"])
    return accounts, opps, expected


def main():
    rng = np.random.default_rng(SEED)
    root = Path(__file__).resolve().parents[1]
    out = root / "data"
    out.mkdir(exist_ok=True)

    accounts = make_accounts(rng)
    opps = make_opps(rng, accounts)
    accounts, opps, expected = plant_errors(rng, accounts, opps)

    # Shuffle so planted rows are not grouped together.
    accounts = accounts.sample(frac=1, random_state=SEED).reset_index(drop=True)
    opps = opps.sample(frac=1, random_state=SEED).reset_index(drop=True)

    accounts.to_csv(out / "accounts.csv", index=False)
    opps.to_csv(out / "opportunities.csv", index=False)
    expected.to_csv(out / "expected_issues.csv", index=False)

    print(f"accounts.csv:      {len(accounts):,} rows")
    print(f"opportunities.csv: {len(opps):,} rows")
    print(f"expected_issues.csv: {len(expected)} planted issues")
    print(expected.groupby(["table", "check"]).size().to_string())


if __name__ == "__main__":
    main()
