# Phase B execution contract — Issue #46

Date: 2026-10-13 (Asia/Bangkok). Repository: `C:\MyRD\tmc-processor-public`.
Baseline: `3758330c15e648fa07c0a95cfe35bc2bd8b20d4b`; branch: `codex/excel-import-phase-b`.

Objective: eliminate redundant Single upload parsing while preserving all accepted engineering behavior.
Scope: workbook resource reuse, preview derived from details, one bounded session-local content cache, lifecycle tests, benchmarks and evidence. Batch measured independently; no Batch cache without measured justification.
Protected: calculation/PCU/PCE/Peak/PHF, Phase A QC provenance and severity, mapping, WorkflowState revisions, explicit Peak confirmation, Project Session schema, export formulas/routing and V1/V2 templates. Real sources remain ignored and read-only.

Execution route: existing chat, single worker, inherited session model/effort (exact runtime selection is not exposed to the worker). Luna High is sufficient for a separately routed bounded implementation once the measurement packet is established; the existing worker also handles unknown-root-cause measurement and final protected-behavior adjudication. A higher tier is not required for the bounded implementation. No new model/worker is invoked, no automatic model escalation. Escalate only if verified evidence reveals conflicting contracts or an unavoidable protected-behavior change. This applies the installed v1.4.1 model-routing policy's Luna-first consideration and existing-chat guidance without creating an additional coding-agent execution.

Pre-implementation scrutiny: GO WITH CONDITIONS. Avoid global caches; content identity must use actual bytes, independently of filename identity. Returned mutable details must be isolated. Cache misses/errors/clear must evict old data. Bound both entry count and retained bytes; report oversized-input behavior. No workflow readiness decisions in the cache. Required acceptance evidence below remains conditional until executed.

| Gate | Required success criterion | Method |
|---|---|---|
| Baseline | Exact SHA, clean initial state, GitHub issue and installed governance inspected | Git/status, Issue #46, upstream workflow revision |
| Measurement | At least five runs, identical environment/fixtures, cold/warm medians, open/read/parse counts and memory | `scripts/benchmark_excel_import.py` before/after; separate actual AppTest reruns |
| RED/GREEN | One cold source parse, zero unchanged warm parses, content replacement/clear, filename behavior, mutation/session isolation, bounded retention | Focused pytest with real synthetic Excel payloads and AppTest |
| Regression | Full suite finishes successfully; never count interruption as pass | `.venv/Scripts/python.exe -m pytest` with progress and faulthandler |
| Compatibility | Single/Batch navigation, readiness/Peak invalidation, Project Session round trips | Focused workflow/session tests plus full suite |
| Real parity | Three sources unchanged; exact normalized, QC, summaries, Peak/PHF and export-content parity | Baseline/final local snapshots with fixed metadata |
| Static/scope | Compile, whitespace, protected files/templates unchanged | compileall, git diff --check and diff audit |
| Handoff | Complete evidence and focused local diff; hosted CI explicitly pending | Evidence document; no push/PR/merge/release/deploy |

Checkpoints: establish/measure; RED then implementation/GREEN; full qualification; diff scrutiny and handoff. Stop/replan on scope expansion or protected-contract conflict. Investigate hangs before retrying; preserve failed/interrupted logs. Definition of done for this local request: validated local branch/diff plus evidence for independent acceptance. Hosted CI is unavailable until a separately authorized push; acceptance is not claimed.
