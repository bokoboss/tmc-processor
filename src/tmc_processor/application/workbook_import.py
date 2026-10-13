"""Bounded Single import reuse owned by the caller's session, never global."""
from __future__ import annotations

import hashlib
from io import BytesIO
import pickle
from typing import MutableMapping

from tmc_processor.importer import ParsedSheet, load_detected_sheet_details

SINGLE_WORKBOOK_CACHE_KEY = "tmc_single_workbook_cache"
MAX_SINGLE_WORKBOOK_CACHE_BYTES = 32 * 1024 * 1024


def single_workbook_details(
    content: bytes | None, state: MutableMapping[str, object],
) -> dict[str, ParsedSheet]:
    """Reuse only the current content; every caller receives independent objects."""
    if content is None:
        state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
        return {}
    digest = hashlib.sha256(content).hexdigest()
    cached = state.get(SINGLE_WORKBOOK_CACHE_KEY)
    if cached is not None and cached[0] == digest:
        return pickle.loads(cached[1])
    # Evict before parsing so a failed replacement cannot retain the old source.
    state.pop(SINGLE_WORKBOOK_CACHE_KEY, None)
    details = load_detected_sheet_details(BytesIO(content))
    # Serialize only our own parsed result, never uploaded pickle data. Immutable
    # bytes isolate DataFrames and nested debug lists without a shared-user cache.
    payload = pickle.dumps(details, protocol=pickle.HIGHEST_PROTOCOL)
    if len(payload) <= MAX_SINGLE_WORKBOOK_CACHE_BYTES:
        state[SINGLE_WORKBOOK_CACHE_KEY] = (digest, payload)
    return details
