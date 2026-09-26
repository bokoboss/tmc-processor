"""Architecture and pure application-layer characterization tests."""

from __future__ import annotations

import ast
import importlib
from pathlib import Path

from tmc_processor.application.state import (
    BATCH_PEAK_SOURCE_KEY,
    BATCH_DISPOSITION_KEY,
    SINGLE_CONFIRMED_PEAKS_KEY,
    confirm_batch_peak,
    confirm_single_peaks,
    get_batch_dispositions,
    get_confirmed_peaks,
    set_batch_disposition,
)
from tmc_processor.application.workflow import (
    CANONICAL_WORKFLOW_STAGES,
    synchronize_workflow,
)
from tmc_processor.workflow_state import WorkflowReadiness, WorkflowRevisions, WorkflowState


ROOT = Path(__file__).parents[1]


def test_application_modules_have_no_streamlit_or_ui_imports() -> None:
    application_root = ROOT / "src" / "tmc_processor" / "application"
    for path in application_root.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        imports = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imports.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                imports.append(node.module or "")
        assert not any(name == "streamlit" or name.startswith("streamlit.") for name in imports), path
        assert not any(name == "app" or name.endswith(".app") for name in imports), path
        assert not any("tmc_processor.ui" in name for name in imports), path


def test_plain_mapping_single_confirmation_preserves_draft_and_source() -> None:
    state: dict[str, object] = {}
    assert confirm_single_peaks(state, "07:00-08:00", "17:00-18:00") is True
    assert confirm_single_peaks(state, "07:00-08:00", "17:00-18:00") is False
    assert get_confirmed_peaks(state) == {
        "am_peak_start": "07:00",
        "am_peak_end": "08:00",
        "pm_peak_start": "17:00",
        "pm_peak_end": "18:00",
    }
    assert state["tmc_confirmed_peak_selection_source"] == "user_confirmed"


def test_plain_mapping_batch_confirmation_and_exclusion_are_explicit() -> None:
    state: dict[str, object] = {}
    assert confirm_batch_peak(state, "001_demo", "08:00-09:00", "17:00-18:00") is True
    assert set_batch_disposition(state, "001_demo", "excluded", "duplicate") is True
    assert get_batch_dispositions(state)["001_demo"] == {"disposition": "excluded", "reason": "duplicate"}
    assert state[BATCH_DISPOSITION_KEY]["001_demo"]["disposition"] == "excluded"
    assert state[BATCH_PEAK_SOURCE_KEY]["001_demo"] == "user_confirmed_batch"


def test_workflow_registry_routes_each_canonical_stage_to_an_owner() -> None:
    from tmc_processor.ui.workflows import WORKFLOW_STAGE_RENDERERS, render_workflow_stage

    seen: list[str] = []
    for stage in CANONICAL_WORKFLOW_STAGES:
        original = WORKFLOW_STAGE_RENDERERS[stage]
        WORKFLOW_STAGE_RENDERERS[stage] = lambda *, context, _stage=stage: seen.append(_stage)
        try:
            render_workflow_stage(stage, context={})
        finally:
            WORKFLOW_STAGE_RENDERERS[stage] = original
    assert seen == list(CANONICAL_WORKFLOW_STAGES)


def test_shell_uses_application_services_for_required_orchestration() -> None:
    shell = (ROOT / "src" / "tmc_processor" / "ui" / "app_shell.py").read_text(encoding="utf-8")
    for service_name in (
        "application_analyze_single",
        "application_analyze_single_dry_run",
        "application_analyze_batch",
        "application_export_single_generated",
        "application_export_single_template_com",
        "application_export_batch_reviewed",
    ):
        assert service_name in shell
    for direct_name in (
        "process_tmc(",
        "process_tmc_dry_run_v2(",
        "analyze_batch_files(",
        "generate_batch_zip_from_reviewed_peaks(",
        "export_v2_generated_workbook(",
        "export_v2_template_workbook_com(",
    ):
        assert direct_name not in shell


def test_shell_dispatches_canonical_stages_and_does_not_own_stage_bodies() -> None:
    shell = (ROOT / "src" / "tmc_processor" / "ui" / "app_shell.py").read_text(encoding="utf-8")
    assert shell.count("render_workflow_stage(") >= 2
    for stage in ("Data", "Mapping", "Review", "Export"):
        assert f'active_tab == "{stage}"' not in shell
    assert "render_batch_stage" not in shell
    assert "render_single_stage" not in shell


def test_workflow_modules_own_real_mode_specific_stage_implementations() -> None:
    workflow_root = ROOT / "src" / "tmc_processor" / "ui" / "workflows"
    assert "render_single_stage" in (workflow_root / "single.py").read_text(encoding="utf-8")
    assert "render_batch_stage" in (workflow_root / "batch.py").read_text(encoding="utf-8")
    for name in ("data", "mapping", "analyze", "review", "export"):
        source = (workflow_root / f"{name}.py").read_text(encoding="utf-8")
        assert "render_single_stage(context=context)" in source
        assert "render_batch_stage(context=context)" in source


def test_workflow_transition_matrix_is_pure_and_downgrades_only_downstream_stages() -> None:
    previous = WorkflowState(
        mode="single",
        revisions=WorkflowRevisions(source="source", mapping="mapping", analysis_config="config", analysis_result="result", review_decision="review", export_config="export"),
        readiness=WorkflowReadiness(True, True, True, True, True),
    )
    review_change = WorkflowRevisions(**{**previous.revisions.__dict__, "review_decision": "new-review"})
    review_result = synchronize_workflow(previous, mode="single", revisions=review_change, readiness=previous.readiness)
    assert review_result.transition.review_updated is True
    assert review_result.state.readiness == WorkflowReadiness(True, True, True, True, False)

    analysis_change = WorkflowRevisions(**{**previous.revisions.__dict__, "analysis_config": "new-config"})
    analysis_result = synchronize_workflow(previous, mode="single", revisions=analysis_change, readiness=previous.readiness)
    assert analysis_result.transition.analysis_invalidated is True
    assert analysis_result.state.readiness == WorkflowReadiness(True, True, False, False, False)


def test_application_and_ui_modules_import_cleanly() -> None:
    for module_name in (
        "tmc_processor.application",
        "tmc_processor.application.state",
        "tmc_processor.application.workflow",
        "tmc_processor.application.services",
        "tmc_processor.ui.components",
        "tmc_processor.ui.workflows",
        "tmc_processor.ui.app_shell",
    ):
        assert importlib.import_module(module_name) is not None


def test_ui_components_do_not_import_engineering_calculation_modules() -> None:
    component_root = ROOT / "src" / "tmc_processor" / "ui" / "components"
    prohibited = {"tmc_processor.peaks", "tmc_processor.pcu", "tmc_processor.qc", "tmc_processor.pipeline"}
    for path in component_root.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        imports = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imports.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom):
                imports.append(node.module or "")
        assert not prohibited.intersection(imports), path


def test_top_level_app_is_a_compatibility_facade() -> None:
    tree = ast.parse((ROOT / "app.py").read_text(encoding="utf-8"))
    function_names = {node.name for node in tree.body if isinstance(node, ast.FunctionDef)}
    assert function_names == {"__getattr__", "__dir__"}
    assert "_run_streamlit_app" not in function_names
    assert CANONICAL_WORKFLOW_STAGES == ("Data", "Mapping", "Analyze", "Review", "Export")
