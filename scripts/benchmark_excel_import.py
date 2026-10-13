"""Issue #46 import-only and actual Streamlit rerun benchmark (no source writes)."""
from __future__ import annotations

import argparse
from contextlib import contextmanager
import gc
import hashlib
from io import BytesIO
import json
from pathlib import Path
import platform
import statistics
import sys
import time
import tracemalloc
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import openpyxl
import pandas as pd
from pandas.io.excel._openpyxl import OpenpyxlReader
import streamlit
from tmc_processor import importer


@contextmanager
def count_io():
    counts = dict(workbook_opens=0, read_excel=0, excel_parse_calls=0, sheet_reads=0, source_parses=0)
    def counted(key, function):
        def run(*args, **kwargs):
            counts[key] += 1
            return function(*args, **kwargs)
        return run
    with patch("openpyxl.load_workbook", counted("workbook_opens", openpyxl.load_workbook)), \
         patch.object(pd, "read_excel", counted("read_excel", pd.read_excel)), \
         patch.object(pd.ExcelFile, "parse", counted("excel_parse_calls", pd.ExcelFile.parse)), \
         patch.object(OpenpyxlReader, "get_sheet_data", counted("sheet_reads", OpenpyxlReader.get_sheet_data)), \
         patch.object(importer, "parse_raw_count_sheet", counted("source_parses", importer.parse_raw_count_sheet)):
        yield counts


def single(payload, state, optimized):
    if optimized:
        from tmc_processor.application.workbook_import import single_workbook_details
        details = single_workbook_details(payload, state)
        names = list(details)
        summary, previews = importer.preview_from_details(details)
    else:
        buffer = BytesIO(payload)
        names = importer.detect_raw_direction_sheet_names(buffer)
        buffer.seek(0)
        summary, previews = importer.preview_detected_sheets(buffer)
        buffer.seek(0)
        details = importer.load_detected_sheet_details(buffer)
    return names, summary, previews, details


def measure(function, runs):
    samples, counts = [], []
    for _ in range(runs):
        gc.collect()
        with count_io() as observed:
            start = time.perf_counter()
            result = function()
            samples.append(time.perf_counter() - start)
        counts.append(observed)
        del result
    gc.collect()
    tracemalloc.start()
    result = function()
    current, peak = tracemalloc.get_traced_memory()
    del result
    gc.collect()
    retained, _ = tracemalloc.get_traced_memory()
    tracemalloc.stop()
    return dict(seconds=samples, median_seconds=statistics.median(samples),
                min_seconds=min(samples), max_seconds=max(samples), counts=counts,
                memory_bytes=dict(result_live=current, peak=peak, after_result_release=retained))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--optimized", action="store_true")
    parser.add_argument("--runs", type=int, default=5)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--streamlit", action="store_true")
    args = parser.parse_args()
    assert args.runs >= 5
    report = dict(environment=dict(python=sys.version, platform=platform.platform(),
                 pandas=pd.__version__, openpyxl=openpyxl.__version__, streamlit=streamlit.__version__),
                 optimized=args.optimized, runs=args.runs, fixtures=[])
    paths = [ROOT / "samples/demo/DEMO_TMC1_FourLeg.xlsx", *sorted((ROOT / "samples/raw").glob("*.xlsx"))]
    for index, path in enumerate(paths):
        payload = path.read_bytes()
        state = {}
        row = dict(fixture="demo" if index == 0 else f"real_{index}", bytes=len(payload),
                   sha256=hashlib.sha256(payload).hexdigest())
        row["single_cold"] = measure(lambda: single(payload, {}, args.optimized), args.runs)
        single(payload, state, args.optimized)
        row["single_warm"] = measure(lambda: single(payload, state, args.optimized), args.runs)
        row["batch_source_load"] = measure(lambda: importer.load_detected_sheets(BytesIO(payload)), args.runs)
        names, _, _, details = single(payload, state, args.optimized)
        row["source_sheets"] = len(names)
        row["dataframe_deep_bytes"] = sum(int(p.data.memory_usage(deep=True).sum() + p.raw.memory_usage(deep=True).sum()) for p in details.values())
        row["cache_entries"] = len(state)
        row["cache_payload_bytes"] = sum(len(value[1]) for value in state.values()) if state else 0
        report["fixtures"].append(row)
        print(row["fixture"], row["source_sheets"], {key: round(row[key]["median_seconds"], 4) for key in ("single_cold", "single_warm", "batch_source_load")}, flush=True)
    if args.streamlit:
        from streamlit.testing.v1 import AppTest
        from tmc_processor.ui import app_shell
        from tmc_processor.excel_com_export import ExcelComStatus
        payload = paths[0].read_bytes()
        def fresh():
            at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60)
            at.session_state[app_shell.SINGLE_SOURCE_UPLOAD_STATE_KEY] = app_shell._upload_record(app_shell._StoredUpload("demo.xlsx", payload))
            return at
        def run(at):
            at.run()
            assert not at.exception, [e.message for e in at.exception]
            return at
        with patch.object(app_shell, "_probe_excel_com_for_ui", lambda force=False: ExcelComStatus(False, "benchmark", "", "")):
            report["streamlit_demo_cold"] = measure(lambda: run(fresh()), args.runs)
            at = run(fresh())
            report["streamlit_demo_warm"] = measure(lambda: run(at), args.runs)
    args.output.write_text(json.dumps(report, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
