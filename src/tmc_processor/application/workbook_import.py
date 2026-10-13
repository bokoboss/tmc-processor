"""Bounded Single import reuse owned by the caller's session, never global."""
from __future__ import annotations

import hashlib
from io import BytesIO
import pickle
from typing import MutableMapping

from tmc_processor.importer import ParsedSheet, load_detected_sheet_details

SINGLE_WORKBOOK_CACHE_KEY = "tmc_single_workbook_cache"
MAX_SINGLE_WORKBOOK_CACHE_BYTES = 32 * 1024 * 1024
# Bump when parser output/provenance semantics or the serialized ParsedSheet
# contract changes. This invalidates surviving Streamlit sessions after reload.
SINGLE_WORKBOOK_PARSER_REVISION = 1


def single_workbook_details(
    content: bytes | None, state: MutableMapping[str, object],
) -> dict[str, ParsedSheet]:
    """Reuse only the current content; every caller receives independent objects."""
    if content is None:
        state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
        return {}
    digest = hashlib.sha256(content).hexdigest()
    cached = state.get(SINGLE_WORKBOOK_CACHE_KEY)
    if (
        isinstance(cached, tuple)
        and len(cached) == 3
        and type(cached[0]) is int
        and cached[0] == SINGLE_WORKBOOK_PARSER_REVISION
        and cached[1] == digest
        and isinstance(cached[2], bytes)
    ):
        try:
            details = pickle.loads(cached[2])
        except Exception:
            # The cache is disposable; reload from the authoritative upload.
            state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
        else:
            if isinstance(details, dict) and all(
                isinstance(name, str) and isinstance(parsed, ParsedSheet)
                for name, parsed in details.items()
            ):
                return details
            state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
    # Evict before parsing so a failed replacement cannot retain the old source.
    state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
    details = load_detected_sheet_details(BytesIO(content))
    # Serialize only our own parsed result, never uploaded pickle data. Immutable
    # bytes isolate DataFrames and nested debug lists without a shared-user cache.
    payload = pickle.dumps(details, protocol=pickle.HIGHEST_PROTOCOL)
    if len(payload) <= MAX_SINGLE_WORKBOOK_CACHE_BYTES:
        state[SINGLE_WORKBOOK_CACHE_KEY] = (SINGLE_WORKBOOK_PARSER_REVISION, digest, payload)
    return details
