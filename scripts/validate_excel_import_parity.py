"""Local Issue #46 parity snapshots; private sources/results stay in outputs/."""
from __future__ import annotations

from io import BytesIO
import hashlib
import json
from pathlib import Path
import pickle
import sys
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
import pandas as pd
from openpyxl import load_workbook
from tmc_processor.importer import load_detected_sheet_details
from tmc_processor.pipeline import process_tmc
from tmc_processor.batch import BatchItem, analyze_batch_files, generate_batch_zip_from_reviewed_peaks


def workbook_digest(payload):
    book = load_workbook(BytesIO(payload), read_only=True, data_only=False)
    try:
        content = {sheet.title: list(sheet.values) for sheet in book}
    finally:
        book.close()
    return hashlib.sha256(json.dumps(content, default=str, sort_keys=True).encode()).hexdigest()


def main():
    phase = sys.argv[1]
    assert phase in ("baseline", "green")
    files = sorted((ROOT / "samples/raw").glob("*.xlsx"))
    assert len(files) == 3, "Qualification requires all three protected local sources"
    source_hashes = {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in files}
    state = {}
    if phase == "green":
        from tmc_processor.application.workbook_import import single_workbook_details
        details = {path.name: single_workbook_details(path.read_bytes(), state) for path in files}
    else:
        details = {path.name: load_detected_sheet_details(path) for path in files}
    raws = {name: {sheet: parsed.data for sheet, parsed in workbook.items()} for name, workbook in details.items()}
    names = sorted(set().union(*(set(raw) for raw in raws.values())))
    codes = ["NS", "NE", "NW", "SN", "SE", "SW", "EN", "ES", "EW", "WN", "WS", "WE"]
    mapping = pd.DataFrame([
        dict(raw_sheet=name, raw_direction=name, movement_code=codes[i % len(codes)],
             from_leg=codes[i % len(codes)][0], to_leg=codes[i % len(codes)][1],
             turn_type="through", facility_type="at_grade", source_stream="mainline",
             include_in_peak=True, include_in_report=True, aggregation_method="sum")
        for i, name in enumerate(names)
    ])
    setup = dict(project_name="Local software regression qualification", survey_period="07:00-19:00")
    fixed_time = "2026-10-13T00:00:00Z"
    snapshots = dict(source_hashes=source_hashes)
    for path in files:
        raw = raws[path.name]
        auto = process_tmc(raw, mapping, setup, generate_workbook=False)
        confirmed = {row.period: (str(row.peak_start), str(row.peak_end)) for row in auto.peaks.itertuples()}
        result = process_tmc(raw, mapping, setup, confirmed_peak_periods=confirmed, generated_at=fixed_time)
        native = process_tmc(raw, mapping, setup, confirmed_peak_periods=confirmed, generated_at=fixed_time,
                             use_template_report_layout=True, use_ooxml_native_template=True)
        frames = {key: getattr(result, key) for key in ("normalized", "qc", "hourly", "movement", "vehicle", "peaks")}
        frames.update(auto_peaks=auto.peaks, auto_qc=auto.qc)
        snapshots[path.name] = dict(frames=frames, details=details[path.name],
                                   safe_digest=workbook_digest(result.workbook_bytes), native_digest=workbook_digest(native.workbook_bytes))
        print(path.name, len(raw), len(result.normalized), result.normalized["count"].sum(), result.normalized.pcu.sum(), flush=True)
    batch = analyze_batch_files([BatchItem(path.name, path.read_bytes()) for path in files],
                               mapping=mapping, setup=setup, generated_at=fixed_time)
    assert all(item.status == "success" for item in batch.items)
    for item in batch.items:
        item.confirmed_AM_peak, item.confirmed_PM_peak = item.suggested_AM_peak, item.suggested_PM_peak
        result = snapshots[item.file_name]["frames"]
        assert item.total_vehicles == float(result["normalized"]["count"].sum())
        assert item.total_PCU == float(result["normalized"].pcu.sum())
    output = generate_batch_zip_from_reviewed_peaks(batch, setup=setup, use_template_report_layout=False)
    assert all(row.status == "success" for row in output.summary_rows), output.summary_rows
    with ZipFile(BytesIO(output.package_bytes)) as archive:
        snapshots["batch_exports"] = {name: workbook_digest(archive.read(name)) for name in archive.namelist() if name.endswith(".xlsx")}
    snapshots["batch_frames"] = {item.file_name: {key: getattr(item, key) for key in ("qc", "hourly_totals", "hourly_movement_pcu", "movement_summary")} for item in batch.items}
    reference = ROOT / "outputs/phase_b_real_baseline.pkl"
    if phase == "baseline":
        reference.write_bytes(pickle.dumps(snapshots))
    else:
        baseline = pickle.loads(reference.read_bytes())
        assert snapshots["source_hashes"] == baseline["source_hashes"]
        assert snapshots["batch_exports"] == baseline["batch_exports"]
        for name in snapshots["batch_frames"]:
            for key, frame in snapshots["batch_frames"][name].items():
                pd.testing.assert_frame_equal(frame, baseline["batch_frames"][name][key])
        for path in files:
            current, expected = snapshots[path.name], baseline[path.name]
            for key, frame in current["frames"].items():
                pd.testing.assert_frame_equal(frame, expected["frames"][key])
            for name, parsed in current["details"].items():
                pd.testing.assert_frame_equal(parsed.data, expected["details"][name].data)
                pd.testing.assert_frame_equal(parsed.raw, expected["details"][name].raw)
                assert parsed.debug == expected["details"][name].debug
            assert current["safe_digest"] == expected["safe_digest"]
            assert current["native_digest"] == expected["native_digest"]
        print("Exact parsed/provenance/normalized/QC/Peak/PHF/summary/Safe/Native/Batch parity: PASS", flush=True)
    assert source_hashes == {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in files}
    print("Original source SHA-256 unchanged; 3/3 Single and reviewed Batch exports: PASS", flush=True)


if __name__ == "__main__":
    main()
