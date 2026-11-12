# Test Plan: Business Data Quality Automation

**Scope:** the Python validation engine (`src/validate.py`), its rules (`config/rules.yaml`), the Power Automate Desktop flow (`DQ Daily Check`) and the Windows Task Scheduler job.

**Test data:** synthetic `accounts.csv` (514 rows) and `opportunities.csv` (1,505 rows) from `src/generate_data.py`, with 90 planted issues listed in `data/expected_issues.csv`.

**Environment:** Windows 11, Python 3.13.9 (Anaconda), pandas, pytest 8.4.2, Power Automate Desktop (free), Windows Task Scheduler.

**Status key:** Pass · Fail · Pending (not yet run)

---

## 1. Test cases

### Python engine

| ID | Test | Steps | Expected result | Actual result | Status |
|---|---|---|---|---|---|
| TC01 | Generate test data | `python src\generate_data.py` | 514 accounts, 1,505 opportunities, 90 planted issues across 13 checks | 514 / 1,505 / 90, counts per check as expected (12 Nov 2026, 01:49) | Pass |
| TC02 | Full validation run | `python src\validate.py` | 13 checks run; 12 FAIL, 1 WARN; failed rows total 90; report, history and log written | 12 FAIL, 1 WARN, 90 failed rows; report `dq_report_20261112_020419.xlsx` created (02:04) | Pass |
| TC03 | Results match the answer key | `python -m pytest -v` (test `test_finds_exactly_the_planted_issues`) | Every planted record found, nothing extra flagged | Passed | Pass |
| TC04 | Automated test suite | `python -m pytest -v` | 15 tests pass | 15 passed in 3.46 s (02:15) | Pass |
| TC05 | Missing config file | `python src\validate.py --config config\missing.yaml` then `echo $LASTEXITCODE` | Error logged in `run_log.txt`; exit code 2; no report | "Config file not found" logged; exit code 2 (02:10) | Pass |
| TC06 | Excel report layout | Open the newest report | Summary sheet with 13 rows, status coloured (red FAIL, yellow WARN); one sheet of failing rows per failed check | Summary showed 13 checks: 12 FAIL in red, 1 WARN in yellow; 13 check tabs; invalid_status tab listed 7 rows (12 Nov 2026) | Pass |
| TC07 | Fix one error and re-run (regression) | Correct one invalid status in `data\accounts.csv` (e.g. "Actve" to "Active"), run `python src\validate.py`, then regenerate the data with `python src\generate_data.py` | `invalid_status` drops from 7 to 6; all other counts unchanged | invalid_status dropped from 7 to 6; all other 12 counts unchanged (12 Nov 2026, 04:18); data regenerated afterwards | Pass |

### Power Automate Desktop flow

| ID | Test | Steps | Expected result | Actual result | Status |
|---|---|---|---|---|---|
| TC08 | Flow success path | Run `DQ Daily Check` | `CommandExitCode` = 0; `CommandOutput` contains "RUN … OK" | Exit code 0; output "2026-11-12 02:35:52 RUN … OK" | Pass |
| TC09 | Flow failure path | Rename `opportunities.csv` to `opportunities_x.csv`, run the flow, rename back | Exit code 2; "RUN FAILED (exit 2)" written to `pad_run_log.txt`; flow stops with error message | "RUN FAILED (exit 2)" logged at 02:52:15; flow stopped. Reason text was empty (see D2) | Pass |
| TC10 | Failure reason is logged | Repeat TC09 after the D2 fix | Log line includes "Data file not found…" | Log line at 04:27:07: "RUN FAILED (exit 2): … ERROR: Data file not found for table 'opportunities'" | Pass |
| TC11 | Report filed into dated folder | Run the flow | Folder `reports\YYYY-MM-DD` created; newest report moved into it | `reports\2026-11-12` created; newest report moved (03:00) | Pass |
| TC12 | Second run on the same day | Run the flow twice | Second run does not fail on the existing date folder; both reports filed | Both runs completed (03:20:55, 03:21:28) | Pass |
| TC13 | Excel summary read by the flow | Run the flow; check `pad_run_log.txt` | "RUN OK: 12 checks failed" (WARN not counted as a failure) | "RUN OK: 12 checks failed" on both runs | Pass |

### Scheduled run

| ID | Test | Steps | Expected result | Actual result | Status |
|---|---|---|---|---|---|
| TC14 | Scheduled task runs the check | `Start-ScheduledTask -TaskName "DQ Daily Check"`, wait 40 s, check `run_log.txt` | New "RUN … OK" line; `LastTaskResult` 0; next run 09:00 | First attempts: result 0 but no output (see D3). After fix: run logged at 03:53:53 | Pass |
| TC15 | Scheduled task on battery | Unplug the charger, repeat TC14 | New "RUN … OK" line | On battery: task started 04:34:04, LastTaskResult 0, RUN OK logged 04:34:41 (12 Nov 2026) | Pass |
| TC16 | Daily trigger fires on time | Laptop was off at 09:00, so the trigger was set 5 minutes ahead (10:50:20) with the laptop on; not run manually; then reset to 09:00 | New "RUN … OK" line at the trigger time | Trigger fired at 10:50:20; RUN OK logged 10:50:27; LastTaskResult 0 (12 Nov 2026) | Pass |
| TC17 | Catch-up after a missed run | Laptop shut down at 09:00; start it later and log in | Missed run starts automatically after logon | Laptop was off at 09:00; task started at 10:32:02 after logon, RUN OK logged 10:32:20, LastTaskResult 0 (12 Nov 2026) | Pass |

---

## 2. Defects found during testing

| ID | Found by | Description | Fix | Retest |
|---|---|---|---|---|
| D1 | TC03 (answer-key comparison) | One planted invalid status, the text "N/A", was not flagged. pandas reads "N/A" as a missing value by default, so the allowed-values check skipped it. | Read CSVs with `keep_default_na=False, na_values=[""]` so only blank cells count as missing. Added regression test `test_text_na_is_treated_as_a_value_not_as_missing`. | TC03 and TC04 pass |
| D2 | TC09 | The flow's failure log line had no reason after "RUN FAILED (exit 2):". The script prints errors to standard output, not the error stream. | Log `%CommandOutput%` instead of `%CommandErrorOutput%`. | TC10 pass |
| D3 | TC14 | Scheduled task reported success (result 0) but did not run; no output from either a batch file or Python. Windows' default task settings do not start tasks while a laptop is on battery. | Re-registered the task with `-AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable`, and pointed it directly at `python.exe` with the project as working folder. | TC14 pass; TC15 pass |
| D4 | Flow designer validation | "Increase variable" was saved with the text `FailCount` instead of the variable `%FailCount%`, so the flow would not run. | Re-entered the variable name with `%` signs. | TC13 pass |

---

## 3. Known limitations

- The free Power Automate Desktop cannot schedule desktop flows; the daily schedule runs the Python check through Windows Task Scheduler, and the flow is run on demand.
- The scheduled task runs only while the laptop is on and the user is logged in.
- All data is synthetic. The project is not connected to any real business system.
