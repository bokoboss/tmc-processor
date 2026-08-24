# Engineering Acceptance Gate

The reviewer uses this gate after receiving the execution evidence package.

## A. Baseline integrity
- [ ] Correct repository and workspace verified.
- [ ] Branch and HEAD provenance are clear.
- [ ] Unrelated user changes were preserved.
- [ ] Scope did not expand without approval.

## B. Requirement coverage
- [ ] Objective is fully implemented.
- [ ] Acceptance criteria are individually satisfied.
- [ ] Required edge cases are addressed.
- [ ] Documentation/user-facing text was updated where necessary.

## C. Regression control
- [ ] Existing relevant tests pass.
- [ ] New/changed behavior has suitable test coverage.
- [ ] No tests/assertions were weakened without justification.
- [ ] Protected behavior/invariants remain intact.

## D. Engineering quality
- [ ] Implementation follows existing architecture/patterns.
- [ ] Change is no broader than necessary.
- [ ] Error handling and boundary conditions are appropriate.
- [ ] `git diff --check` passes.

## E. Runtime / delivery evidence
- [ ] Required build/package checks pass.
- [ ] Required runtime/UI/browser validation passes.
- [ ] Required real-data or domain validation passes.
- [ ] Required CI checks pass.

## F. Risk closure
- [ ] Known limitations are explicitly documented.
- [ ] Skipped/unrun validation is acceptable or resolved.
- [ ] No unresolved blocker remains.

## Decision
Choose one:
- **ACCEPT** — all required gates pass.
- **ACCEPT WITH DOCUMENTED LIMITATIONS** — requirements are met and remaining limitations are explicit and acceptable.
- **REMEDIATE** — implementation is directionally correct but evidence or behavior is incomplete.
- **BLOCKED** — a material dependency, defect, baseline conflict, or missing prerequisite prevents acceptance.

## Accepted baseline record
- Accepted branch:
- Accepted HEAD SHA:
- PR / merge SHA:
- CI run:
- Date:
- Next milestone:
