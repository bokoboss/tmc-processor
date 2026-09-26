"""Application service composition tests."""

from __future__ import annotations

from types import SimpleNamespace

from tmc_processor.application import services


def test_single_analysis_service_delegates_without_reimplementing_engine() -> None:
    calls: list[dict[str, object]] = []

    def fake_process_tmc(**kwargs: object) -> object:
        calls.append(kwargs)
        return "single-result"

    original = services.process_tmc
    services.process_tmc = fake_process_tmc
    try:
        result = services.analyze_single({"Sheet": []}, [], {"movement_code_scheme": "from_to"}, peak_mode="rolling_60min")
    finally:
        services.process_tmc = original

    assert result == "single-result"
    assert calls == [{
        "raw_sheets": {"Sheet": []},
        "mapping": [],
        "setup": {"movement_code_scheme": "from_to"},
        "peak_mode": "rolling_60min",
    }]


def test_batch_analysis_service_passes_domain_items_and_options_through() -> None:
    calls: list[tuple[object, dict[str, object]]] = []

    def fake_analyze(items: list[object], **kwargs: object) -> object:
        calls.append((items, kwargs))
        return SimpleNamespace(items=items)

    original = services.analyze_batch_files
    services.analyze_batch_files = fake_analyze
    try:
        result = services.analyze_batch(["one", "two"], mapping_preset={"mapping_rows": []})
    finally:
        services.analyze_batch_files = original

    assert result.items == ["one", "two"]
    assert calls == [(["one", "two"], {"mapping_preset": {"mapping_rows": []}})]
