# Issue #40 Excel interoperability execution contract

Date: 2026-10-12, Asia/Bangkok.

## Baseline and scope

Repository: `bokoboss/tmc-processor`. Authoritative baseline:
`c23020021978032607072403229132160f977658` (fetched `origin/main`).
Initial worktree was clean at `2ed777b4016a9657252637ceae9ab0c74daed693`,
on `codex/issue-38-runtime-hotfix`. Work now uses
`codex/issue-40-excel-hotfix`, created from the authoritative baseline.

Restore usable Microsoft Excel export with correct engineering data. First
verify the supplied evidence hashes, reproduce in desktop Excel, and isolate
the smallest demonstrated defect. Preserve original evidence and private real
workbooks outside Git history. Do not merge, tag, publish, or close release gates.

## Execution routing

Use installed engineering-development-workflow v1.4.1 at
`3547ae260feacf8fc9a102b2abfdb13881e36dab`, including its normative
`MODEL_ROUTING_POLICY.md` and `DEBUGGING_PROTOCOL.md` (read at that revision).
Continue in the existing user-selected chat with one worker: the cause is
unknown, prior static acceptance conflicts with actual Excel failure, and
premium diagnostic judgment is appropriate under the unknown-root-cause route.
No higher tier or additional worker is currently necessary. Once a reliable
reproducer exists, bounded implementation is suitable for Luna High; escalation
requires a diagnosed capability limitation, not a sandbox or dependency problem.

## Protected behavior

Calculation, Mapping, QC, PCE/PCU, Peak selection, NT mainline/frontage,
WorkflowState revision/readiness, V1/V2 data, native Summary, authored drawings
and charts remain protected. Native compatibility fixes require specific proof;
an explicitly labelled compatibility option may be added only if native repair
cannot reasonably succeed. Never silently remove native content or change
engineering semantics.

## Success gates defined before implementation

| Gate | Required evidence |
| --- | --- |
| Provenance | Repository/branch/clean baseline, Issue #40, exact SHA-256 of originals |
| Reproduce | Microsoft Excel version, exact failed/control GUI open, Protected View state |
| Recovery | Open and Repair on a copy, outcome and recovery log if generated |
| Localize | Systematic package, relationships, worksheet, formula, drawing and chart comparison |
| RED | Reproducer for a demonstrated defect fails against unchanged baseline before fix |
| GREEN | Minimal proven fix passes the reproducer, preserving protected content |
| Data | V1/V2, mapping, Peak, PCE/PCU, NT streams, Summary and supporting sheets reconcile |
| Excel | Freshly generated and actual Cloud-downloaded XLSX opens without repair |
| Regression | Focused tests, full pytest, compileall and diff check pass |
| CI | Linux Python 3.12 and Windows Python 3.10/3.12 pass at PR head |
| Delivery | Reviewable PR, changed files, exact commands/results, Excel evidence and limitations |

## Evidence decision rule

ZIP/XML/openpyxl success alone is insufficient for Excel acceptance. Protected
View rendering is recorded separately from normal editable opening. A passing
baseline is not a RED regression and cannot justify speculative writer changes.
If the exact bytes open on the same machine, investigate the original opening
route and editing transition before classifying the failure or claiming a fix.

All source binary evidence, generated workbooks, screenshots, browser snapshots
and raw logs stay under ignored local evidence folders. Public evidence must not
include private workbooks or user-identifying desktop content.
