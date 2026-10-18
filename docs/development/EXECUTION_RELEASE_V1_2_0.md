# v1.2.0 release preparation — execution contract

Date: 2026-10-18 (Asia/Bangkok). Decision: **GO WITH CONDITIONS for preparation and independent review; publication is not authorized.**

## Baseline and scope

- Repository: `C:\MyRD\tmc-processor-public`, origin `https://github.com/bokoboss/tmc-processor.git`.
- Accepted main: `3f5a0f8e92893ac5f8b3174d47081d350a636133`; post-merge [CI 39120306254](https://github.com/bokoboss/tmc-processor/actions/runs/39120306254) passed Windows Python 3.10/3.12 and Linux Python 3.12.
- Initial checkout: `codex/issue42-v1-precision` at `a3df88956d3922afd2ad2b1ea524a42f2dda7026`; its tracked tree equals accepted main. Only untracked `docs/development/ISSUE42_INVESTIGATION.md` is present. Preserve it and all existing ignored outputs.
- Feature branch: `codex/release-v1.2.0-preparation`, created directly from accepted main. Local main remains `1514b142e566f1edf64ac1ba5aaf6b5a1523541b`.
- Published release remains v1.1.0; v1.2.0 is an unreleased source-distribution candidate.
- Changes: package version, existing version-test expectations, README, PROJECT_PROFILE, release notes and this execution/evidence package. Use the existing builder without changing packaging behavior.
- No feature, calculation, mapping, QC, Peak, cache, WorkflowState, Single/Batch, export, template/map, launcher or dependency-policy changes. No private data, local outputs or investigation file added to Git. No merge, tag, release, deploy, branch deletion, or modification of main.

## Governance and execution route

Read AGENTS, PROJECT_PROFILE, installed workflow v1.4.1 reference/manifest/templates, pinned upstream model-routing policy at `3547ae260feacf8fc9a102b2abfdb13881e36dab`, prior release and accepted Phase A/B/#42 evidence. Applied the Ponytail skill and workflow scrutiny guidance. Mode: **STRICT qualification**, because release acceptance includes protected engineering claims; implementation remains mechanical.

Luna High is sufficient for bounded metadata/document edits. Continue this existing GPT-6 Codex session at inherited effort (exact runtime selection is not exposed), one worker, without an additional agent/model switch: the substantive work is cross-tool evidence reconciliation, packaging/runtime/Excel qualification and controlled GitHub handoff. No premium escalation is needed for routine edits. External writes are limited to the explicitly requested feature-branch push and PR. All filesystem writes stay within the repository. A protected behavior difference, failed numerical gate, secret/private payload or contradictory baseline stops release acceptance; diagnose first and report rather than broaden scope. Independent engineering and release acceptance remain outside the executor's own recommendation.

## Pre-implementation scrutiny

Primary false-success risks: inheriting stale pending GUI text despite later acceptance; confusing published v1.1.0 with the candidate; treating Linux CI as production Cloud acceptance; qualifying the checkout rather than the ZIP; treating COM recalculation as GUI/print acceptance; and treating package reproducibility as reproducibility of timestamped Excel exports. Address each explicitly in the evidence. Preserve original print pagination and record A4 readability as a separate inherited limitation.

## Success gates

| Gate | Required evidence |
|---|---|
| G1 Baseline | Exact Git/GitHub main, HEAD, branch, clean tracked tree, tags, latest published release, accepted CI and Phase A/B/#42 records; preserve investigation hash and prior outputs. |
| G2 Scope/version | Package/UI/export version 1.2.0; published v1.1.0 links retained; accurate candidate changes/limitations; no production Cloud acceptance claim. Protected source/template diff empty. |
| G3 Packages | Existing builder creates both named ZIPs twice; byte-identical archives/manifests, CRC and every entry SHA-256 verified; runtime dependencies, QR, both native template/map pairs, allowlists and nested OOXML privacy checks. |
| G4 Runtime/demo | Extracted source bundle, isolated local environment and dependency/import checks; Windows batch launcher startup/health; synthetic Single/Batch and V1/V2 Standard/Safe exports; correct effective Peaks and accepted totals. |
| G5 Regression | Focused release/protected-contract tests, full suite with no hidden exclusions, compileall and git diff --check; log commands, exit codes, counts/warnings and environment limitations. |
| G6 Excel/acceptance | Fresh synthetic Excel read-only open/full recalc/hash preservation if available; distinguish inherited manual GUI acceptance from candidate GUI qualification, original pagination, one-page scaling and readable A4. Explicitly list skipped/unverified gates. |
| G7 Handoff | Evidence with baseline/final implementation SHA, changed-file/diff summary, commands/results, hashes and GO/NO-GO; commit/push only feature branch, PR targeting main, exact-head CI status. No publication. |

If an environment prevents a gate, complete other authorized work and record **UNVERIFIED/NO-GO for that acceptance claim**, without changing protected behavior to make a test pass.
