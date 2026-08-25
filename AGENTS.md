# Project Agent Instructions

This project follows the Engineering Development Workflow:
https://github.com/bokoboss/engineering-development-workflow

Before changing code:
1. Read `PROJECT_PROFILE.md`.
2. Inspect the actual repository and current Git/GitHub state.
3. Preserve project-specific invariants and protected behavior.
4. Define success gates before implementation.
5. Use the shared workflow's model-routing policy for coding-agent execution.

Project-specific instructions belong below this line and override the shared workflow only where explicitly stated.

## Project-specific instructions

- Preserve validated calculation, mapping, QC, Peak, and export behavior unless the task explicitly scopes and approves a change.
- Treat `WorkflowState` and its revision/invalidation contract in `src/tmc_processor/workflow_state.py` as authoritative for engineering readiness. Legacy Streamlit session-state keys are adapters, not a second source of truth.
- Protect real-workbook regression behavior in `samples/raw/`, including the Kabin Buri, Bo Phloi, and Nong Prue workbooks; do not add those local files to public Git history.
- Do not use destructive Git operations or discard existing user changes without explicit authorization.
- Implementation work must include an evidence package covering scope, baseline, validation commands/results, changed files, and known limitations.
- Keep the project-owned `.ai-workflow/` governance directory intact. Resolve any overlap with this official workflow as a documented follow-up, not as an automatic consolidation.
