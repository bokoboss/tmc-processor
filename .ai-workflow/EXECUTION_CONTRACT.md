# Codex / Local Executor Contract

Use this contract as the default boundary for implementation work.

## 1. Objective
State exactly what must change and why.

## 2. Authoritative workspace
- Repository:
- Local path:
- Expected base branch / SHA:

Do not work outside this repository unless explicitly authorized.

## 3. Required preflight
Before changing files:
1. Confirm repository path.
2. Report current branch and HEAD.
3. Report `git status` and preserve unrelated dirty changes.
4. Inspect the relevant implementation and tests before editing.
5. Stop and report if the baseline materially differs from the stated baseline.

## 4. Scope
### In scope
- 

### Out of scope
- 

### Prohibited changes
- Do not rewrite unrelated modules.
- Do not delete or overwrite unrelated user work.
- Do not change external behavior not required by the objective.
- Do not push, merge, release, delete branches, or rewrite Git history unless explicitly requested.

## 5. Implementation principles
- Prefer the smallest coherent change.
- Reuse existing project patterns and abstractions.
- Add or update tests for changed behavior.
- Preserve backwards compatibility unless explicitly waived.
- Do not hide failing tests, weaken assertions, or remove validation merely to make the suite pass.

## 6. Validation
Run the project-specific required gates from `PROJECT_PROFILE.md`, including relevant:
- tests
- lint / static checks
- build / package checks
- runtime smoke checks
- UI/browser checks
- `git diff --check`

If a required gate cannot be run, state why and treat it as unresolved evidence.

## 7. Completion report
Return an evidence package using `EVIDENCE_PACKAGE_TEMPLATE.md`.
Do not report the task as accepted; report only implementation status. Final acceptance belongs to the reviewer.
