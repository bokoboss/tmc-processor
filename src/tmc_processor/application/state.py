"""Pure adapters for application/session state.

These helpers intentionally operate on ordinary mutable mappings.  Streamlit
session state is one adapter at the UI boundary; it is not required here and
is not imported by this module.
"""

from __future__ import annotations

from collections.abc import MutableMapping
from typing import Any

SINGLE_CONFIRMED_PEAKS_KEY = "tmc_application_confirmed_peaks"
SINGLE_DRAFT_PEAKS_KEY = "tmc_application_draft_peaks"
SINGLE_PEAK_SOURCE_KEY = "tmc_application_peak_source"
BATCH_CONFIRMED_PEAKS_KEY = "tmc_application_batch_confirmed_peaks"
BATCH_DRAFT_PEAKS_KEY = "tmc_application_batch_draft_peaks"
BATCH_DISPOSITION_KEY = "tmc_application_batch_dispositions"
EXPORT_BACKEND_KEY = "tmc_application_export_backend"

USER_CONFIRMED = "user_confirmed"
USER_CONFIRMED_BATCH = "user_confirmed_batch"


def _mapping(state: MutableMapping[str, object] | None) -> MutableMapping[str, object]:
    """Return a usable mapping without ever reaching into Streamlit state."""

    return state if state is not None else {}


def _complete_peak_payload(am_peak: str, pm_peak: str) -> dict[str, str]:
    return {
        "am_peak_start": str(am_peak).split("-", 1)[0].strip(),
        "am_peak_end": str(am_peak).split("-", 1)[-1].strip(),
        "pm_peak_start": str(pm_peak).split("-", 1)[0].strip(),
        "pm_peak_end": str(pm_peak).split("-", 1)[-1].strip(),
    }


def get_draft_peaks(state: MutableMapping[str, object] | None) -> dict[str, str]:
    value = _mapping(state).get(SINGLE_DRAFT_PEAKS_KEY, {})
    return dict(value) if isinstance(value, dict) else {}


def get_confirmed_peaks(state: MutableMapping[str, object] | None) -> dict[str, str]:
    source = _mapping(state)
    value = source.get(SINGLE_CONFIRMED_PEAKS_KEY, {})
    confirmed = dict(value) if isinstance(value, dict) else {}
    for legacy_key, peak_key in {
        "tmc_confirmed_am_peak_start": "am_peak_start",
        "tmc_confirmed_am_peak_end": "am_peak_end",
        "tmc_confirmed_pm_peak_start": "pm_peak_start",
        "tmc_confirmed_pm_peak_end": "pm_peak_end",
    }.items():
        if source.get(legacy_key):
            confirmed[peak_key] = str(source[legacy_key])
    if source.get("tmc_loaded_confirmed_peaks") and isinstance(source["tmc_loaded_confirmed_peaks"], dict):
        confirmed = {**source["tmc_loaded_confirmed_peaks"], **confirmed}
    return confirmed


def set_draft_peaks(state: MutableMapping[str, object], peaks: dict[str, object] | None) -> dict[str, str]:
    normalized = {str(key): str(value) for key, value in (peaks or {}).items() if value is not None}
    state[SINGLE_DRAFT_PEAKS_KEY] = normalized
    return normalized


def confirm_single_peaks(
    state: MutableMapping[str, object],
    am_peak: str,
    pm_peak: str,
) -> bool:
    """Apply an explicit Single confirmation and report semantic change."""

    confirmed = _complete_peak_payload(am_peak, pm_peak)
    previous = get_confirmed_peaks(state)
    changed = any(previous.get(key, "") != value for key, value in confirmed.items())
    state[SINGLE_CONFIRMED_PEAKS_KEY] = confirmed
    state[SINGLE_DRAFT_PEAKS_KEY] = dict(confirmed)
    state[SINGLE_PEAK_SOURCE_KEY] = USER_CONFIRMED
    for key, value in confirmed.items():
        state[f"tmc_confirmed_{key}"] = value
    state["tmc_confirmed_peak_source"] = USER_CONFIRMED
    state.pop("tmc_loaded_confirmed_peaks", None)
    return changed


def clear_single_review_state(state: MutableMapping[str, object]) -> None:
    for key in (
        SINGLE_CONFIRMED_PEAKS_KEY,
        SINGLE_DRAFT_PEAKS_KEY,
        SINGLE_PEAK_SOURCE_KEY,
        "tmc_confirmed_am_peak_start",
        "tmc_confirmed_am_peak_end",
        "tmc_confirmed_pm_peak_start",
        "tmc_confirmed_pm_peak_end",
        "tmc_confirmed_peak_source",
        "tmc_loaded_confirmed_peaks",
    ):
        state.pop(key, None)


def get_batch_dispositions(state: MutableMapping[str, object] | None) -> dict[str, dict[str, str]]:
    source = _mapping(state)
    value = source.get(BATCH_DISPOSITION_KEY, {})
    if not isinstance(value, dict):
        value = {}
    result = {str(name): dict(details) for name, details in value.items() if isinstance(details, dict)}
    legacy = source.get("tmc_batch_file_dispositions")
    if isinstance(legacy, dict):
        result = {**{str(name): dict(details) for name, details in legacy.items() if isinstance(details, dict)}, **result}
    return result


def get_batch_peaks(state: MutableMapping[str, object] | None, key: str) -> dict[str, dict[str, str]]:
    source = _mapping(state)
    value = source.get(key, {})
    if not isinstance(value, dict):
        value = {}
    result = {str(name): dict(details) for name, details in value.items() if isinstance(details, dict)}
    if key == BATCH_CONFIRMED_PEAKS_KEY:
        legacy = source.get("tmc_batch_confirmed_peaks")
        if isinstance(legacy, dict):
            result = {**{str(name): dict(details) for name, details in legacy.items() if isinstance(details, dict)}, **result}
    return result


def get_batch_confirmed_peaks(state: MutableMapping[str, object] | None) -> dict[str, dict[str, str]]:
    return get_batch_peaks(state, BATCH_CONFIRMED_PEAKS_KEY)


def get_batch_draft_peaks(state: MutableMapping[str, object] | None) -> dict[str, dict[str, str]]:
    return get_batch_peaks(state, BATCH_DRAFT_PEAKS_KEY)


def confirm_batch_peak(state: MutableMapping[str, object], folder_name: str, am_peak: str, pm_peak: str) -> bool:
    confirmed = _complete_peak_payload(am_peak, pm_peak)
    name = str(folder_name)
    previous = get_batch_confirmed_peaks(state).get(name, {})
    confirmed_map = get_batch_confirmed_peaks(state)
    draft_map = get_batch_draft_peaks(state)
    confirmed_map[name] = confirmed
    draft_map[name] = dict(confirmed)
    state[BATCH_CONFIRMED_PEAKS_KEY] = confirmed_map
    state[BATCH_DRAFT_PEAKS_KEY] = draft_map
    state["tmc_batch_confirmed_peaks"] = {**get_batch_peaks(state, "tmc_batch_confirmed_peaks"), name: {"AM": am_peak, "PM": pm_peak}}
    sources = state.setdefault("tmc_batch_confirmed_peak_sources", {})
    if isinstance(sources, dict):
        sources[name] = USER_CONFIRMED_BATCH
    return previous != confirmed


def set_batch_disposition(state: MutableMapping[str, object], folder_name: str, disposition: str, reason: str = "") -> bool:
    name = str(folder_name)
    dispositions = get_batch_dispositions(state)
    previous = dispositions.get(name, {})
    next_value = {"disposition": str(disposition), "reason": str(reason)}
    dispositions[name] = next_value
    state[BATCH_DISPOSITION_KEY] = dispositions
    state["tmc_batch_file_dispositions"] = dict(dispositions)
    return previous != next_value


def clear_batch_review_state(state: MutableMapping[str, object]) -> None:
    for key in (
        BATCH_CONFIRMED_PEAKS_KEY,
        BATCH_DRAFT_PEAKS_KEY,
        BATCH_DISPOSITION_KEY,
        "tmc_batch_confirmed_peaks",
        "tmc_batch_draft_peaks",
        "tmc_batch_confirmed_peak_sources",
        "tmc_batch_file_dispositions",
    ):
        state[key] = {}


def set_export_backend(state: MutableMapping[str, object], backend: str, *, requested: str | None = None, fallback: str | None = None) -> None:
    state[EXPORT_BACKEND_KEY] = {
        "requested": str(requested or backend),
        "used": str(backend),
        "fallback": str(fallback or ""),
    }


def get_export_backend(state: MutableMapping[str, object] | None) -> dict[str, str]:
    value = _mapping(state).get(EXPORT_BACKEND_KEY, {})
    return {str(key): str(item) for key, item in value.items()} if isinstance(value, dict) else {}
