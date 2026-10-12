# Issue #40: desktop Excel and Cloud export qualification

Date: 2026-10-12, Asia/Bangkok. Issue:
https://github.com/bokoboss/tmc-processor/issues/40.

## Outcome and acceptance boundary

The exact previously rejected Cloud workbook and a newly downloaded public
Cloud V2 export both opened in Microsoft 365 Excel on the user's same machine
without a format/extension error or repair prompt. The user enabled editing;
both then became normal, editable workbooks. Excel's actual calculation values,
supporting PCU data, native chart series and formula-error scan passed.

**The original error was not reproduced. Its root cause remains unproven. No
production hotfix or compatibility workaround was implemented, and no RED/GREEN
defect regression can honestly be claimed.** This evidence is for review of
current usability and investigation status. It does not establish why the
earlier opening failed, close Issue #40, or approve release #36.

## Provenance and changed files

- Accepted and rechecked remote `main`:
  `c23020021978032607072403229132160f977658`.
- Initial clean local HEAD: `2ed777b4016a9657252637ceae9ab0c74daed693`.
- Execution branch: `codex/issue-40-excel-hotfix`, created from accepted main.
- Those initial and accepted commits have identical trees (`git diff 2ed777b
  c230200` is empty).
- Changed tracked files: this evidence package and
  `docs/development/ISSUE40_EXECUTION_CONTRACT.md` only.
- No changes to `app.py`, production source, templates, tests, calculations,
  Mapping, QC, Peak, WorkflowState, diagrams, charts or export routing.
- Routing and success gates: see `ISSUE40_EXECUTION_CONTRACT.md`.
- PR and head/CI identifiers: recorded in the draft PR and its checks.

The evidence directory contained the extracted originals and README; the
original reproduction ZIP was absent. The four supplied binary SHA-256 values
all matched the supplied manifest before use. No workbook was resaved through
openpyxl or another workbook writer.

| Artifact | Bytes | SHA-256 |
| --- | ---: | --- |
| Rejected Cloud v1.1.0 | 272753 | `86638f7d0dd1fdf604db77be992280234ac23d36053b48ffd63a7411caa38805` |
| Earlier Cloud v1.0.1 | 272753 | `5f9af503319df177c0ae99906b9cca96ee5d0e828d81c52bda866fa40074c2bb` |
| Issue #32 Excel-opened control | 264650 | `934ad67a2b3ccbb3c20b59fab301b95476089bbf4e2a123a58c5d5f6a026cae9` |
| Supplied error screenshot | 13980 | `d5d13ca72d219676e85259684f8d2f9417f0058da21e73d7c042989f1b526094` |
| Fresh local native V2 | 275084 | `4ee56b6eadb11fdd17b116879ce204c1365c31ee58de24b3706f21fe3393d827` |
| Fresh public Cloud native V2 | 275214 | `ad029b172dd633dc33abe043cf01cf7547e46710ad38a582e12d6a1deb53ffe2` |

## Microsoft Excel GUI evidence

Excel executable version: **16.0.20430.20146**; live Excel reports version 16.0,
build 20430. The user confirmed the original failure was on this same machine
using M365.

Each file was selected through Excel's actual File > Open dialog. Computer Use
captured the windows. Protected View and normal editable opening were recorded
separately. The user performed Enable Editing, because computer-use rules
prohibit automating security permission actions.

| Workbook | Initial opening | Normal editable result |
| --- | --- | --- |
| Exact rejected Cloud original | Protected View; Summary, diagrams and charts rendered; no error/repair prompt | PASS after user enabled editing; `ReadOnly=False` |
| Issue #32 control | Protected View; no error/repair prompt | PASS after user enabled editing; `ReadOnly=False` |
| Fresh local baseline V2 | Normal opening immediately; no Protected View or repair prompt | PASS; `ReadOnly=False` |
| Fresh actual Cloud download | Protected View; no error/repair prompt | PASS after user enabled editing; `ReadOnly=False` |

The original files carried Internet zone metadata. It was recorded, not removed;
no Trust Center, Windows or Excel security settings were changed. The fresh
Cloud download also retained its Internet zone metadata.

Open and Repair was invoked via File > Open > Browse > Open dropdown on a
SHA-identical **separate copy** of the rejected Cloud workbook. Repair was
selected in Excel's recovery dialog. Excel displayed the workbook, initially in
Protected View, with no recovery-summary dialog or repair-log link. The user
enabled editing; it became editable with all 24 sheets and two native charts.
No repair log was presented or captured. This result is not proof that Excel
performed a repair. Originals were not overwritten or saved.

Local GUI captures include `cloud_protected_view.png`,
`cloud_original_editable_open.png`, `issue32_control_open.png`,
`new_native_v2_baseline_open.png`, `cloud_open_and_repair_result.png`, and
`fresh_cloud_editable_open.png`. Corresponding accessibility text and live
Excel value readings are included in the ignored local evidence archive.

## Cloud and engineering reconciliation

The actual public app at https://tmc-process.streamlit.app/ was exercised with
only the repository's public synthetic demo and its V2 mapping preset:

`Data -> Mapping -> Analyze -> Review -> confirm Peak -> Export -> download XLSX`.

The download was obtained through the browser's real download event, not an
API-generated substitute. Fresh Cloud metadata reports app version 1.1.0,
`Excel Template Mode`, `four_leg_approach_movement_v2`,
`approach_movement`, generated at `2026-10-12T13:44:42Z`.
No compatibility option or silent fallback was used.

The following results were read directly from the already GUI-opened Excel
workbooks through read-only COM calls. No COM opening, saving, or repair was
used as a substitute for the GUI acceptance. The original was also recalculated
using Excel's Ctrl+Alt+F9 before final readings.

| Check | Original / fresh local / fresh Cloud |
| --- | --- |
| Worksheets | 24, including Summary and audit/support sheets |
| `Summary!F30` AM | 2981 |
| `Summary!F31` PM | 3153 |
| `Summary!F32`, `W40`, `AM22` | 23959 each |
| Precise PCU sum in Normalized_Data / Hourly_Totals / Movement_Summary | 23959.041 in each sheet |
| Normalized rows | 2304 |
| Included movement codes | ET, NT, WT |
| NT frontage PCU | 4170.511 |
| NT mainline PCU | 8329.429 |
| Native Excel charts | 2, with 12 category/value points each |
| Excel formula error cells | None across all worksheets (`SpecialCells` formula/error scan) |

V1/V2, PCE factors, physical Mapping, NT stream aggregation, fixed-hour Peak,
Summary rounding and supporting exports are additionally covered by existing
focused/full regressions. The three protected private real workbooks were
parsed locally: 24/15/12 detected sheets and 13824/8640/6912 parsed rows,
respectively. No new real-workbook engineering Mapping was invented and no
private binary was added to Git. This is a parser smoke, not a fresh mapped
real-workbook numerical qualification.

## Systematic OOXML examination

All five compared workbooks passed CRC, XML parsing and internal relationship
target/id checks. Both current outputs retained 24 worksheets and 168 formula
elements. The #32 control has 20 worksheets and 168 formula elements.
The drawing part is byte-identical to the #32 control and retains 33 authored
shape elements and two chart frames. Native chart series are present in actual
Excel and have the expected 12 points.

The rejected and earlier Cloud artifacts have exactly 41 of 42 byte-identical
uncompressed members. Only Export_Metadata (`sheet2.xml`) differs. The #32
comparison also inspected workbook relationships/content types, worksheet
payloads, Summary formulas, chart caches and workbook structure. No specific
OOXML incompatibility was demonstrated; no package-bisect mutations were
justified once both original and fresh files opened normally.

## Validation commands and results

Commands were run from the repository using its existing Python 3.12 environment.

| Gate | Command/method | Result |
| --- | --- | --- |
| V1 native fidelity | `python -m pytest tests/test_ooxml_template_export.py -q --basetemp=dist/issue40-evidence/pytest-baseline` | 31 passed, 2 existing DrawingML warnings |
| Focused data/export | `python -m pytest tests/test_v2_pcu_rounding.py tests/test_pce_factors.py tests/test_physical_mapping.py tests/test_phase_l0_dual_scheme.py tests/test_phase_f_v2_dry_run.py -q --basetemp=dist/issue40-evidence/focused-temp --tb=short` | See final local log / PR |
| Full suite | `python -u -m pytest -q --basetemp=dist/issue40-evidence/pytest-full-native --tb=short` | **551 passed, 30 warnings**, 488.34 seconds; no skips/failures |
| Compile | `python -m compileall -q app.py src` | PASS |
| Whitespace | `git diff --check` | PASS |
| Excel calculations | Local `validate_live_excel.py`, read-only access to named already-open workbooks | PASS; includes fresh downloaded Cloud file |
| Baseline cross-platform CI | CI #130, run `38150593199`, exact source tree equivalent to accepted main | Linux Python 3.12 and Windows Python 3.10/3.12 PASS |
| Draft PR CI | Linux and Windows checks at PR head | See draft PR checks |

The first sandbox full-suite attempts were interrupted after diagnosing a
Windows loopback socket restriction in Streamlit AppTest; they are not counted
as successful runs. The complete suite above ran outside that restriction.
While Excel GUI prompts were active it emitted a handled Windows COM
`0x80010001` diagnostic stack from optional Excel probing; all 551 tests still
passed. Logs retain this diagnostic and the 30 warnings. ZIP/openpyxl successes
are supplemental and are not presented as Excel GUI proof.

## Remaining limitations and reviewer handoff

- Historical format/extension rejection remains unexplained; no proven
  correction or new defect RED/GREEN regression exists.
- Verified GUI qualification covers this installed M365 build, these exact
  artifacts and the public synthetic V2 workflow. It does not certify all Excel
  builds, all workbooks or the historical download/open route.
- V1 is covered by the existing native-export tests/full suite, not a new V1
  GUI acceptance run in this investigation.
- Browser console retains Cloud shell user-details 404s and shell iframe
  warnings; the app workflow completed and downloaded the report. No application
  exception appeared in the exercised flow.
- Workbook binaries, private raw data, desktop screenshots and raw local logs
  remain in ignored local storage. The public PR contains evidence documents.
- Review the evidence before deciding whether #40 can be requalified or needs
  another exact failing artifact/open-route reproduction. No merge, tag,
  release publication or issue closure was performed.
