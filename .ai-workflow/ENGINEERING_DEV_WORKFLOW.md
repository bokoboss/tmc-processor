# Engineering Development Workflow v1.0

## Purpose
A reusable coordination protocol for ChatGPT, Codex/local execution, GitHub, and local repositories. It provides the operational benefits of an agent loop without granting ChatGPT remote access to the computer.

## Roles
### ChatGPT — Planner / Reviewer / Acceptance Authority
- recover context and baseline
- define scope and acceptance criteria
- prepare bounded execution instructions
- review evidence
- classify acceptance
- prepare phase handoff

### Codex / Local Executor — Implementation Layer
- inspect the authorized repository
- implement bounded changes
- run required validation
- return reproducible evidence
- never self-accept the work

### GitHub — Remote Source of Truth
- repository history
- pull requests
- CI evidence
- release provenance when applicable

### User — Change Authority
- authorizes destructive or externally visible actions
- controls local machine execution and sensitive data exposure
- resolves business/domain decisions that cannot be inferred from repository evidence

## Operating loop
1. BASELINE — verify repository identity, branch, SHA, dirty state, relevant remote state.
2. PLAN — state objective, constraints, scope, protected behavior, and acceptance criteria.
3. EXECUTE — delegate bounded implementation through the Execution Contract.
4. VERIFY — executor runs project-defined gates and returns Evidence Package.
5. AUDIT — ChatGPT independently checks evidence against requirements.
6. DECIDE — ACCEPT / ACCEPT WITH DOCUMENTED LIMITATIONS / REMEDIATE / BLOCKED.
7. CLOSE — record accepted baseline and next milestone.

## Core rules
- Evidence beats claims.
- Never assume local state.
- Preserve unrelated user work.
- Prefer minimal coherent changes.
- Keep execution within the declared repository root.
- Do not weaken tests to achieve green status.
- Do not push/merge/release/destructively modify Git unless explicitly authorized.
- A commit is not acceptance.
- Passing unit tests alone may not be sufficient when runtime, UI, packaging, real-data, or CI gates apply.

## Evidence schema
Every meaningful implementation should report:
- provenance: repo, path, branch, SHA before/after
- changed files
- commands executed
- test/build/runtime results
- diff check
- CI status when relevant
- limitations, skipped gates, risks, deviations

## Project specialization
This workflow is global. Project-specific facts belong in `PROJECT_PROFILE.md`, including:
- authoritative paths
- standard commands
- architecture/invariants
- protected behaviors
- validation matrix
- release policy
- known risks

## Security model
This workflow does not require RDC, MCP remote shell access, or an always-on local agent. ChatGPT communicates intentions and reviews evidence; execution remains local and explicitly controlled.
