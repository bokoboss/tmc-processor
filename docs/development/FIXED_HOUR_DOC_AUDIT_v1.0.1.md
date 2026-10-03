# Fixed-hour documentation remediation and submission audit

Prepared 2026-10-03 on `maintenance/fixed-hour-peak-contract`.
Accepted baseline: `20b166c9d3f9dd9af06830522009f384109bd55d`.

## Remediation

The pre-submission integrity check found two STALE-CONTRADICTION occurrences:
README line 184 called non-hour-aligned rolling Peaks valid with Safe PNG
fallback; line 247 treated native rolling support as active backlog. Both are
corrected. Whole-hour guidance now explicitly requires HH:00 endpoints,
60-minute duration, and four underlying 15-minute intervals. Standard and
Safe PNG accept supported whole-hour Peaks only; legacy incompatible decisions
require fresh Analyze, Review and explicit confirmation, without rounding.
Issue #23 remains open until PR/merge/release acceptance and is no longer an
active native-template enhancement target. USER_GUIDE is consistent.

Historical UX evidence, the baseline rolling audit and v1.0.0 release notes
retain their version-specific account. Their titles, baseline SHAs and
qualification sections identify historical context; they are not current
v1.0.1 operator instructions. Generic Safe PNG alternative/fallback references
do not authorize unsupported rolling Peaks.

## Global occurrence classification

Search covered all Markdown/RST documentation in the repository, excluding
ignored generated artifacts and protected samples. Pattern:
`rolling_60min|rolling Peak|rolling 60.minute|08[:.]15|09[:.]15|Issue #23|#23|Safe PNG.*fallback|non.hour.aligned`.
Every matching line is classified below; multiple terms on one line share its
classification. Documents with no match have no rolling-contract exposure.
This audit's own inventory is evidence rather than operator instructions.

| Document | Line | Classification |
|---|---:|---|
| `CHANGELOG.md` | 35 | HISTORICAL-CONTEXT |
| `docs/UX_UAT_ARCHITECTURE_REVIEW.md` | 94 | HISTORICAL-CONTEXT |
| `USER_GUIDE.md` | 33 | CURRENT-CONTRACT-CORRECT |
| `docs/UX_IMPLEMENTATION_BACKLOG.md` | 233 | HISTORICAL-CONTEXT |
| `docs/releases/v1.0.1.md` | 21 | CURRENT-CONTRACT-CORRECT |
| `docs/releases/v1.0.0.md` | 15 | HISTORICAL-CONTEXT |
| `docs/releases/v1.0.0.md` | 34 | HISTORICAL-CONTEXT |
| `docs/APPROACH_MOVEMENT_TEMPLATE_DESIGN.md` | 214 | HISTORICAL-CONTEXT |
| `RELEASE_NOTES_v0.2.0.md` | 31 | HISTORICAL-CONTEXT |
| `PROJECT_PROFILE.md` | 86 | CURRENT-CONTRACT-CORRECT |
| `PROJECT_PROFILE.md` | 140 | HISTORICAL-CONTEXT |
| `PROJECT_PROFILE.md` | 153 | CURRENT-CONTRACT-CORRECT |
| `PROJECT_PROFILE.md` | 172 | CURRENT-CONTRACT-CORRECT |
| `README.md` | 53 | CURRENT-CONTRACT-CORRECT |
| `README.md` | 70 | CURRENT-CONTRACT-CORRECT |
| `README.md` | 182 | CURRENT-CONTRACT-CORRECT |
| `README.md` | 184 | CURRENT-CONTRACT-CORRECT |
| `README.md` | 247 | CURRENT-CONTRACT-CORRECT |
| `docs/development/EVIDENCE_OOXML_UX7.md` | 66 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX1.md` | 45 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX1.md` | 58 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX1.md` | 65 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX1.md` | 82 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX3.md` | 163 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 97 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 105 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 119 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 172 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 180 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 184 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 186 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 188 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 199 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 200 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 202 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 203 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX7.md` | 212 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX4_UX5.md` | 31 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX4_UX5.md` | 41 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX4_UX5.md` | 88 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX4_UX5.md` | 89 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX4_UX5.md` | 104 | HISTORICAL-CONTEXT |
| `docs/development/EVIDENCE_PACKAGE_UX8_P1.md` | 65 | HISTORICAL-CONTEXT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 54 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 67 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 118 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 123 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 125 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 127 | CURRENT-CONTRACT-CORRECT |
| `docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md` | 138 | CURRENT-CONTRACT-CORRECT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 1 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 15 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 32 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 35 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 38 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 41 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 42 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 48 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 51 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 153 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 154 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 155 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 162 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 167 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 168 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 175 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 176 | HISTORICAL-CONTEXT |
| `docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md` | 177 | HISTORICAL-CONTEXT |

Final inventory: 17 CURRENT-CONTRACT-CORRECT,
51 HISTORICAL-CONTEXT, **0 STALE-CONTRADICTION**.

## Integrity and validation

- All 57 non-template runtime payload files matched the accepted candidate
  manifest before remediation. Portable template files are unchanged in Git.
- All runtime source/version fingerprints remain byte-identical to the
  accepted qualification. 87 of 88 source/test/version fingerprints are
  unchanged; the new contract test differs only by removal of one terminal
  blank line required by staged whitespace validation. Its executable content
  is unchanged. Product remediation changed documentation only.
- Focused contract rerun: **27 passed** in 1.85 seconds.
- `python -m compileall -q app.py src`: exit 0.
- `git diff --check`: exit 0.
- Reused accepted full suite: **504 passed, 14 existing DrawingML warnings**.
  No implementation bytes changed after that result; no full-suite or expensive
  clean-room runtime rerun was necessary for the documentation corrections.
- Fresh isolated Windows runtime, populated Single/Batch Review and
  Standard/Safe PNG qualification remain recorded in
  [the maintenance evidence](FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md).

This audit records the pre-commit submission gate. No merge, release/tag
publication or Issue #23 closure is authorized by this gate.

## Exact submission file scope

All 25 paths below are relative to `C:\MyRD\tmc-processor-public`.

```text
PROJECT_PROFILE.md
README.md
USER_GUIDE.md
docs/development/FIXED_HOUR_DOC_AUDIT_v1.0.1.md
docs/development/FIXED_HOUR_PEAK_v1.0.1_EVIDENCE.md
docs/development/ROLLING_PEAK_AUDIT_v1.0.1.md
docs/releases/v1.0.1.md
pyproject.toml
src/tmc_processor/application/state.py
src/tmc_processor/application/workflow.py
src/tmc_processor/batch.py
src/tmc_processor/constants.py
src/tmc_processor/peaks.py
src/tmc_processor/pipeline.py
src/tmc_processor/session.py
src/tmc_processor/time_utils.py
src/tmc_processor/ui/app_shell.py
src/tmc_processor/ui/workflows/batch.py
src/tmc_processor/ui/workflows/single.py
tests/test_export_package.py
tests/test_fixed_hour_peak_contract.py
tests/test_setup_persistence.py
tests/test_ux7_standard_ooxml_routing.py
tests/test_ux8_p1_ui_hierarchy.py
tests/test_workflow_contract_adapter.py
```
