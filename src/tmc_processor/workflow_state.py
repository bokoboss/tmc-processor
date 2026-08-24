"""Pure application-level workflow state and invalidation contracts.

This module deliberately has no Streamlit dependency.  The UI may keep legacy
session-state keys as an adapter, while these fingerprints and transitions
define which engineering artifacts remain valid after an input change.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass, field, is_dataclass, replace
from datetime import date, datetime, time
import hashlib
import json
import math
from collections.abc import Mapping, Sequence
from typing import Any


_ANALYSIS_INPUT_FIELDS = ("source", "mapping", "analysis_config")


def _canonical_value(value: Any) -> Any:
    """Return a deterministic JSON-compatible representation of ``value``."""

    if value is None or isinstance(value, (bool, int, str)):
        return value
    if isinstance(value, float):
        if math.isnan(value):
            return None
        if math.isinf(value):
            return str(value)
        return value
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {"__bytes_sha256__": hashlib.sha256(bytes(value)).hexdigest()}
    if isinstance(value, (datetime, date, time)):
        return value.isoformat()
    if is_dataclass(value):
        return _canonical_value(asdict(value))
    if isinstance(value, Mapping):
        return {
            str(key): _canonical_value(item)
            for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
        }
    if isinstance(value, (set, frozenset)):
        canonical_items = [_canonical_value(item) for item in value]
        return sorted(canonical_items, key=lambda item: json.dumps(item, sort_keys=True, ensure_ascii=False))
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes, bytearray, memoryview)):
        return [_canonical_value(item) for item in value]

    # DataFrame-like values are accepted without importing pandas into this
    # application contract module.
    to_dict = getattr(value, "to_dict", None)
    if callable(to_dict):
        try:
            return _canonical_value(to_dict(orient="records"))
        except TypeError:
            return _canonical_value(to_dict())

    item = getattr(value, "item", None)
    if callable(item):
        try:
            item_value = item()
        except (TypeError, ValueError):
            item_value = value
        if item_value is not value:
            return _canonical_value(item_value)

    return str(value)


def semantic_fingerprint(value: Any) -> str:
    """Return a stable SHA-256 fingerprint for semantic JSON-like content."""

    payload = json.dumps(
        _canonical_value(value),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
        allow_nan=False,
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def source_fingerprint(content: bytes | bytearray | memoryview | None) -> str | None:
    """Fingerprint workbook bytes, independent of filename or upload metadata."""

    if content is None:
        return None
    return hashlib.sha256(bytes(content)).hexdigest()


def _mapping_records(mapping: Any) -> list[Any]:
    if mapping is None:
        return []
    if isinstance(mapping, Mapping):
        return [mapping]
    if isinstance(mapping, Sequence) and not isinstance(mapping, (str, bytes, bytearray, memoryview)):
        return list(mapping)
    to_dict = getattr(mapping, "to_dict", None)
    if callable(to_dict):
        try:
            records = to_dict(orient="records")
        except TypeError:
            records = to_dict()
        return list(records) if isinstance(records, Sequence) else [records]
    return [mapping]


def mapping_fingerprint(mapping: Any) -> str:
    """Fingerprint mapping rows while ignoring editor/view bookkeeping."""

    return semantic_fingerprint({"rows": _mapping_records(mapping)})


def analysis_config_fingerprint(
    *,
    pce_factors: Mapping[str, Any] | None = None,
    peak_mode: Any = None,
    peak_windows: Mapping[str, Any] | None = None,
    movement_code_scheme: Any = None,
    extra: Mapping[str, Any] | None = None,
) -> str:
    """Fingerprint settings that can change the engineering analysis result."""

    payload: dict[str, Any] = {
        "pce_factors": dict(pce_factors or {}),
        "peak_mode": peak_mode,
        "peak_windows": dict(peak_windows or {}),
        "movement_code_scheme": movement_code_scheme,
    }
    if extra:
        payload["extra"] = dict(extra)
    return semantic_fingerprint(payload)


def analysis_result_fingerprint(result: Any) -> str | None:
    """Fingerprint a deterministic result marker or JSON-compatible result."""

    if result is None:
        return None
    return semantic_fingerprint(result)


def review_decision_fingerprint(decision: Any) -> str | None:
    """Fingerprint effective Peak/review decisions; empty decisions are absent."""

    if decision is None or decision == {} or decision == [] or decision == "":
        return None
    return semantic_fingerprint(decision)


def export_config_fingerprint(config: Any) -> str:
    """Fingerprint metadata and export settings that affect an artifact."""

    return semantic_fingerprint(config)


@dataclass(frozen=True)
class WorkflowRevisions:
    """Semantic revisions for the application dependency chain."""

    source: str | None = None
    mapping: str | None = None
    analysis_config: str | None = None
    analysis_result: str | None = None
    review_decision: str | None = None
    export_config: str | None = None

    def with_updates(self, **changes: str | None) -> "WorkflowRevisions":
        return replace(self, **changes)


@dataclass(frozen=True)
class WorkflowReadiness:
    """Explicit readiness for each engineering stage."""

    source: bool = False
    mapping: bool = False
    analysis: bool = False
    review: bool = False
    export: bool = False


@dataclass(frozen=True)
class WorkflowState:
    """Mode-specific state snapshot suitable for a UI adapter."""

    mode: str
    revisions: WorkflowRevisions = field(default_factory=WorkflowRevisions)
    readiness: WorkflowReadiness = field(default_factory=WorkflowReadiness)


@dataclass(frozen=True)
class WorkflowTransition:
    """Pure decision describing downstream invalidation for a revision change."""

    changed_fields: tuple[str, ...] = ()
    analysis_invalidated: bool = False
    review_invalidated: bool = False
    review_updated: bool = False
    export_invalidated: bool = False

    @property
    def engineering_state_changed(self) -> bool:
        return bool(self.changed_fields)


def transition_workflow(
    previous: WorkflowRevisions | None,
    current: WorkflowRevisions,
) -> WorkflowTransition:
    """Compare semantic revisions and return the required pure transition.

    ``analysis_result`` is an output revision: changing it keeps Analysis valid
    but invalidates any Review/Export artifacts that were based on the prior
    result.  View-only UI state is intentionally not represented here.
    """

    if previous is None:
        return WorkflowTransition()

    changed: list[str] = []
    for field_name in (
        "source",
        "mapping",
        "analysis_config",
        "analysis_result",
        "review_decision",
        "export_config",
    ):
        if getattr(previous, field_name) != getattr(current, field_name):
            changed.append(field_name)

    analysis_inputs_changed = any(field_name in changed for field_name in _ANALYSIS_INPUT_FIELDS)
    analysis_result_changed = "analysis_result" in changed
    review_changed = "review_decision" in changed
    export_config_changed = "export_config" in changed

    return WorkflowTransition(
        changed_fields=tuple(changed),
        analysis_invalidated=analysis_inputs_changed,
        review_invalidated=analysis_inputs_changed or analysis_result_changed,
        review_updated=review_changed,
        export_invalidated=analysis_inputs_changed or analysis_result_changed or review_changed or export_config_changed,
    )
