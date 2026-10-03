# Rolling Peak exposure before v1.0.1

Baseline: `20b166c9d3f9dd9af06830522009f384109bd55d`. Read-only search before implementation.

- `constants.py:106-109`: fixed default, but operator options include fixed and rolling.
- `peaks.py`: candidate generator evaluates every raw interval for rolling; confirmed calculation accepts any positive duration.
- `pipeline.py`: Single and v2 pass the selected mode and stored/explicit Peaks to calculation.
- `batch.py`: three public defaults are rolling; labels are truncated to five characters, Review checks presence only.
- `ui/workflows/single.py` and `batch.py`: both Analyze pages offer the mode selectbox; Batch Review inserts legacy values into options.
- `ui/app_shell.py`: setup adapter restores the mode; Single Review appends prior result intervals; confirmation/readiness check only populated values.
- `session.py`: session load and state application restore rolling mode and old confirmed Peaks without compatibility handling.
- `application/services.py`, `application/state.py`, `application/workflow.py`: services forward mode, state adapters accept arbitrary labels, mode participates in analysis fingerprints.
- `workflow_state.py`: generic revision authority includes mode; transition semantics must be preserved.
- `exporter.py`: recalculates exact chosen Peak; native-binding mismatch routes rolling selections to Safe PNG.
- README describes Issue #23 as an active rolling limitation. USER_GUIDE does not restrict Peak choices.
- The committed demo project session uses fixed_hourly already; no demo migration needed.
- Tests cover mode persistence and rolling export fallback; those explicit expectations require replacement under the clarified contract.

## Complete relevant search results

```text
20b166c9d3f9dd9af06830522009f384109bd55d:samples/demo/DEMO_TMC1_FourLeg_session.tmcproj.json:115:    "peak_mode": "fixed_hourly",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/application/workflow.py:62:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/application/workflow.py:75:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/application/workflow.py:90:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/application/workflow.py:115:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:772:    peak_mode: str,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:805:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:849:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:991:    peak_mode: str,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1029:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1166:    peak_mode: str = "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1219:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1235:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1333:    peak_mode: str = "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1461:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1477:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1545:    peak_mode: str = "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1562:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/batch.py:1572:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/constants.py:107:PEAK_MODE_ROLLING_60MIN = "rolling_60min"
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/constants.py:108:PEAK_MODE_OPTIONS = [PEAK_MODE_FIXED_HOURLY, PEAK_MODE_ROLLING_60MIN]
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/exporter.py:349:            peak_mode=str(setup.get("peak_mode") or DEFAULT_PEAK_MODE),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/exporter.py:384:    peak_mode = setup.get("peak_mode") or _peak_value(am_peak, "peak_mode") or _peak_value(pm_peak, "peak_mode") or ""
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/exporter.py:395:                "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/exporter.py:516:                "peak_mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:12:    PEAK_MODE_OPTIONS,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:13:    PEAK_MODE_ROLLING_60MIN,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:88:def _candidate_starts(interval: pd.DataFrame, window_start: int, window_end: int, peak_mode: str) -> list[int]:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:89:    if peak_mode == PEAK_MODE_FIXED_HOURLY:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:92:    if peak_mode == PEAK_MODE_ROLLING_60MIN:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:95:    raise ValueError(f"Invalid peak calculation mode: {peak_mode}")
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:101:    peak_mode: str = DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:103:    if peak_mode not in PEAK_MODE_OPTIONS:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:104:        raise ValueError(f"Invalid peak calculation mode: {peak_mode}")
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:108:        "peak_mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:129:        for minute in _candidate_starts(interval, window_start, window_end, peak_mode):
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:138:                "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:156:    peak_mode: str = DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:161:    if peak_mode not in PEAK_MODE_OPTIONS:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:162:        raise ValueError(f"Invalid peak calculation mode: {peak_mode}")
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:165:        "peak_mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/peaks.py:194:                "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:71:    peak_mode: str = DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:106:        "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:132:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:137:        peaks = detect_peak_phf(normalized, windows=peak_windows, peak_mode=peak_mode)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:181:    peak_mode: str = DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:204:        "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:222:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/pipeline.py:226:        peaks = detect_peak_phf(normalized, windows=peak_windows, peak_mode=peak_mode)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/session.py:52:    "peak_mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/session.py:314:    if "peak_mode" in peaks:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/session.py:315:        updates["peak_mode_select"] = peaks["peak_mode"] or DEFAULT_PEAK_MODE
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:59:    PEAK_MODE_OPTIONS,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:157:    hourly_interval_options as base_hourly_interval_options,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:260:    "peak_mode": "peak_mode_select",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:274:    "peak_mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:411:    peak_mode: str,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:440:            "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:470:        "peak_mode": DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:500:    if field == "peak_mode":
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:501:        return value if value in PEAK_MODE_OPTIONS else DEFAULT_PEAK_MODE
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:629:        peak_mode=str(values["peak_mode"] or DEFAULT_PEAK_MODE),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:664:def _hourly_interval_options(hourly_movement: pd.DataFrame, peaks: pd.DataFrame) -> list[tuple[str, str, str]]:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:665:    options = base_hourly_interval_options(hourly_movement)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1129:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1142:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1160:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1174:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1517:    peak_mode: str,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1529:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:1539:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:3689:        peak_mode=str(setup.get("peak_mode") or DEFAULT_PEAK_MODE),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:3714:        "peak_mode": setup.get("peak_mode", DEFAULT_PEAK_MODE),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:4458:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:4462:        str(peak_mode or ""),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:4472:    peak_mode: str | None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:4479:        _batch_peak_settings_signature(peak_mode, peak_windows),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:4609:        peak_mode=str(setup.get("peak_mode") or DEFAULT_PEAK_MODE),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5305:        PEAK_MODE_OPTIONS=PEAK_MODE_OPTIONS,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5347:        _hourly_interval_options=_hourly_interval_options,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5686:    peak_mode = st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5713:        peak_mode = st.session_state.get("peak_mode_select", DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5740:    peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5786:                peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/app_shell.py:5854:                peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflow_context.py:42:    PEAK_MODE_OPTIONS: Any = None
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflow_context.py:84:    _hourly_interval_options: Any = None
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflow_context.py:374:    peak_mode: Any = None
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflow_context.py:375:    peak_mode_default: Any = None
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:294:    PEAK_MODE_OPTIONS = context.operations.PEAK_MODE_OPTIONS
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:333:    peak_mode = context.peak_mode
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:334:    peak_mode_default = context.peak_mode_default
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:346:        peak_mode_default = str(st.session_state.get("peak_mode_select") or DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:347:        if peak_mode_default not in PEAK_MODE_OPTIONS:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:348:            peak_mode_default = DEFAULT_PEAK_MODE
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:349:            st.session_state["peak_mode_select"] = peak_mode_default
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:375:        with st.expander(f"วิธีคำนวณ Peak · {peak_mode_default}", expanded=False):
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:376:            peak_mode = st.selectbox(
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:377:                "Peak calculation mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:378:                options=PEAK_MODE_OPTIONS,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:379:                index=PEAK_MODE_OPTIONS.index(peak_mode_default),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:380:                key="peak_mode_select",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:397:    peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:434:                peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:453:            peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:900:    peak_mode = context.peak_mode
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/batch.py:1019:                peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:446:    PEAK_MODE_OPTIONS = context.operations.PEAK_MODE_OPTIONS
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:487:    peak_mode = context.peak_mode
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:488:    peak_mode_default = context.peak_mode_default
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:524:        peak_mode_default = str(st.session_state.get("peak_mode_select") or DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:525:        if peak_mode_default not in PEAK_MODE_OPTIONS:
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:526:            peak_mode_default = DEFAULT_PEAK_MODE
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:527:            st.session_state["peak_mode_select"] = peak_mode_default
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:553:        with st.expander(f"วิธีคำนวณ Peak · {peak_mode_default}", expanded=False):
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:554:            peak_mode = st.selectbox(
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:555:                "Peak calculation mode",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:556:                options=PEAK_MODE_OPTIONS,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:557:                index=PEAK_MODE_OPTIONS.index(peak_mode_default),
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:558:                key="peak_mode_select",
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:574:    peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:604:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:625:                peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:656:    _hourly_interval_options = context.operations._hourly_interval_options
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:751:        interval_options = _hourly_interval_options(hourly_movement, result.peaks)
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:1053:    peak_mode = context.peak_mode
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:1232:                    peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/ui/workflows/single.py:1256:                        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/workflow_state.py:118:    peak_mode: Any = None,
20b166c9d3f9dd9af06830522009f384109bd55d:src/tmc_processor/workflow_state.py:127:        "peak_mode": peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_application_services.py:20:        result = services.analyze_single({"Sheet": []}, [], {"movement_code_scheme": "from_to"}, peak_mode="rolling_60min")
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_application_services.py:29:        "peak_mode": "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ooxml_template_export.py:42:    peak=pd.DataFrame([{"peak_mode":"rolling_60min","am_peak_search_window":"06:00-10:00","pm_peak_search_window":"15:00-19:00","am_peak_start":"07:00","am_peak_end":"08:00","am_peak_pcu":999,"am_phf":0.91,"pm_peak_start":"17:00","pm_peak_end":"18:00","pm_peak_pcu":1111,"pm_phf":0.92,"peak_selection_source":"user_confirmed"}])
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_phase_j_ui_helpers.py:62:        peak_mode=DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_phase_j_ui_helpers.py:117:        peak_mode=DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_phase_j_ui_helpers.py:221:        peak_mode=DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_phase_j_ui_helpers.py:263:        peak_mode=DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_setup_persistence.py:303:    assert _widget_by_key(at.selectbox, app.SETUP_FIELD_WIDGET_KEYS["peak_mode"]).value == expected["peak_mode"]
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_setup_persistence.py:325:        "peak_mode": app.DEFAULT_PEAK_MODE,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_setup_persistence.py:335:        "peak_mode": "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_setup_persistence.py:340:    _widget_by_key(at.selectbox, app.SETUP_FIELD_WIDGET_KEYS["peak_mode"]).set_value(configured["peak_mode"])
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_setup_persistence.py:349:    assert at.session_state[app.SETUP_STATE_KEY]["peak_mode"] == configured["peak_mode"]
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux7_standard_ooxml_routing.py:66:def _run(decision, *, ooxml: bool | None = None, am=("07:00", "08:00"), pm=("16:00", "17:00"), peak_mode="fixed_hourly"):
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux7_standard_ooxml_routing.py:74:        peak_mode=peak_mode,
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux7_standard_ooxml_routing.py:186:        "peak_mode": "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux7_standard_ooxml_routing.py:206:    result = _run(decision, am=("08:15", "09:15"), pm=("17:00", "18:00"), peak_mode="rolling_60min")
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux8_p1_ui_hierarchy.py:127:    assert next(widget for widget in peak.get("selectbox") if widget.key == "peak_mode_select").value == app.DEFAULT_PEAK_MODE
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_ux8_p1_ui_hierarchy.py:153:    assert all(widget.key != "peak_mode_select" for widget in at.selectbox)
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract.py:115:    first = analysis_config_fingerprint(pce_factors={"MC": 1.0}, peak_mode="fixed")
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract.py:116:    second = analysis_config_fingerprint(pce_factors={"MC": 1.2}, peak_mode="fixed")
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract.py:127:        peak_mode="fixed",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract.py:132:        peak_mode="rolling",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract_adapter.py:353:            "peak_mode": "rolling_60min",
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract_adapter.py:366:    assert restored["peak_mode"] == "rolling_60min"
20b166c9d3f9dd9af06830522009f384109bd55d:tests/test_workflow_contract_adapter.py:442:    st.session_state["peak_mode_select"] = "rolling_60min"
```
