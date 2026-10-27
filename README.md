# Lab 5 — Static Code Analysis

A complete submission for the Lab 5 handout: a Python script analysed with
**Pylint**, **Bandit** and **Flake8**, all findings documented, every issue
fixed, and the before/after evidence preserved.

## Result

| Tool | Before | After |
|------|--------|-------|
| Pylint | 34 messages — **5.64 / 10** | 0 messages — **10.00 / 10** |
| Bandit | **9 issues** (3 High, 2 Medium, 4 Low) | **0 issues** |
| Flake8 | **17 violations** | **0 violations** |

## What is in this repository

| Path | What it contains |
|------|------------------|
| `original_script.py` | The buggy 106-line baseline that was analysed. |
| `user_utils.py` | The cleaned and updated script (passes all three tools). |
| `reports/before/` | `pylint_report.txt`, `bandit_report.txt`, `flake8_report.txt` captured before the fixes. |
| `reports/after/` | The same three reports after the fixes; the empty Flake8 file means "no findings". |
| `reports/stages/` | Intermediate reports after the Pylint fix and after the Bandit fix (verification evidence). |
| `Lab5_Static_Code_Analysis_Report.docx` | The main deliverable: issue table, fixes, verification, reflection and appendices. |
| `config.txt`, `users.txt` | Sample data so the program runs end to end. |
| `build_report.py` | Regenerates the DOCX from the fixed data and the report files. |

## The three fix commits

| Commit | Tool | What it fixed |
|--------|------|---------------|
| `fa5586d` | Pylint | Mutable default argument, bare/broad exceptions, redundant global, missing docstrings, unused names, singleton comparisons, file handling. |
| `b2b2b31` | Bandit | MD5 → salted SHA-256, `os.system` → `shutil.copy2`, `shell=True` removed, `eval` → `ast.literal_eval`, `random` → `secrets`, `assert` → `ValueError`, hardcoded secret removed, `/tmp` → `tempfile`. |
| `c3ecc1a` | Flake8 | Blank lines, spacing around comma/operator, 87-character line split, final reports regenerated. |

The baseline commit `424a8b0` contains the original script and the "before"
reports. All fixes were made in `user_utils.py` only; `original_script.py` is a
frozen copy of the baseline so that before/after can always be compared.

## Reproduce

```bash
python -m venv venv
# Windows: venv\Scripts\activate      Linux/macOS: source venv/bin/activate
pip install pylint bandit flake8

# before
pylint original_script.py > reports/before/pylint_report.txt
bandit -r original_script.py > reports/before/bandit_report.txt
flake8 original_script.py > reports/before/flake8_report.txt

# after
pylint user_utils.py > reports/after/pylint_report.txt
bandit -r user_utils.py > reports/after/bandit_report.txt
flake8 user_utils.py > reports/after/flake8_report.txt

python build_report.py     # regenerate the DOCX report
```

Versions used: Python 3.12.14, Pylint 4.1.3, Bandit 1.9.4, Flake8 7.4.1.

## Note on the baseline

The handout refers to a script from an assigned Git repository, which was not
part of the provided materials. `original_script.py` is therefore a
representative user-management utility written to reproduce every issue class
named in the lab (mutable default arguments, overly broad exceptions and
insecure coding patterns) — so the whole exercise is reproducible from this
repository alone.
