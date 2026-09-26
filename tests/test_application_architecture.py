"""Architecture and pure application-layer characterization tests."""

from __future__ import annotations

import ast
import importlib
from pathlib import Path

from tmc_processor.application.state import (
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
    assert state["tmc_application_peak_source"] == "user_confirmed"


def test_plain_mapping_batch_confirmation_and_exclusion_are_explicit() -> None:
    state: dict[str, object] = {}
    assert confirm_batch_peak(state, "001_demo", "08:00-09:00", "17:00-18:00") is True
    assert set_batch_disposition(state, "001_demo", "excluded", "duplicate") is True
    assert get_batch_dispositions(state)["001_demo"] == {"disposition": "excluded", "reason": "duplicate"}
    assert state[BATCH_DISPOSITION_KEY]["001_demo"]["disposition"] == "excluded"


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
