# Issue #30 / PR #31 Community Cloud execution contract

Date: 2026-10-08 (Asia/Bangkok).

## Baseline and scope

Repository: `C:\MyRD\tmc-processor-public`, `bokoboss/tmc-processor`.
Existing branch: `codex/issue-30-web-readiness-support`.
Clean local and remote baseline: `46901e719a9b4938b568ad9de3c393dba2ecdc76`.
Accepted release baseline: `main@39c1ba906dcc284b714e5980b9ef9e23dfe47827`.

Correct Support text using Issue #30 comment 6061137433, move Privacy to Data
uploads, add root requirements backed by pyproject, qualify Linux Python 3.12
installation/entrypoint and document Community Cloud deployment.

## Execution route

Pinned governance: engineering-development-workflow v1.4.1, revision
`3547ae260feacf8fc9a102b2abfdb13881e36dab`; MODEL_ROUTING_POLICY reviewed.
Luna High is sufficient for this bounded multi-file packet; no premium worker
or parallel coordination is required. Execute in the existing user-selected
chat with its configured model; no model escalation or additional agents.
Escalate only for a reproducible capability limitation after environment diagnosis.

## Protected behavior and boundaries

Preserve calculations, Mapping, Analyze, Review, Export, WorkflowState, theme,
session contracts and QR bytes. No real samples in Git history. No local
real-workbook/COM rerun: this task does not touch their paths. No Vercel config,
database or permanent upload storage. Keep PR draft; no merge, tag, release or
Issue closure until final evidence review. Package version remains the published
1.0.1 until the separately approved release step.

## Success gates

| Gate | Proof |
| --- | --- |
| Support | Focused tests, exact final wording, state preservation and missing QR fallback |
| Privacy | Absent from header; available beside Single and Batch Data uploads |
| Dependency | Clean Ubuntu Python 3.12 installs root requirements with uv and pip; pip check; no pywin32 |
| Entrypoint | AppTest executes app.py without exceptions; headless Streamlit health smoke |
| Hygiene | compileall app.py/src and git diff --check |
| CI | Updated PR Windows 3.10/3.12 and Linux 3.12 gates succeed |
| Delivery | Bounded commit/push to existing branch; evidence and deployment instructions |

Stop rather than expand scope if Linux readiness requires protected engineering
changes. Community Cloud sign-in, approved main merge, public sample workflow
smoke and release review remain deployment/release gates.
