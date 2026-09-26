"""Batch workflow stage implementation."""

from __future__ import annotations

from collections.abc import Mapping

import streamlit as st


def render_batch_stage(*, context: Mapping[str, object]) -> None:
    """Render the complete Batch stage surface using shell-provided context."""
    AM_WINDOW = context.get('AM_WINDOW')
    BATCH_CONFIRMED_PEAKS_STATE_KEY = context.get('BATCH_CONFIRMED_PEAKS_STATE_KEY')
    BATCH_CONFIRMED_PEAK_SOURCE_STATE_KEY = context.get('BATCH_CONFIRMED_PEAK_SOURCE_STATE_KEY')
    BATCH_DISPOSITION_STATE_KEY = context.get('BATCH_DISPOSITION_STATE_KEY')
    BATCH_DRAFT_PEAKS_STATE_KEY = context.get('BATCH_DRAFT_PEAKS_STATE_KEY')
    BATCH_EXCEL_TEMPLATE_EXPORT_MODE = context.get('BATCH_EXCEL_TEMPLATE_EXPORT_MODE')
    BATCH_PACKAGE_MIME = context.get('BATCH_PACKAGE_MIME')
    BATCH_SAFE_PNG_EXPORT_LABEL = context.get('BATCH_SAFE_PNG_EXPORT_LABEL')
    BATCH_SAFE_PNG_EXPORT_MODE = context.get('BATCH_SAFE_PNG_EXPORT_MODE')
    BATCH_V2_TEMPLATE_MODE_UNSUPPORTED_TH = context.get('BATCH_V2_TEMPLATE_MODE_UNSUPPORTED_TH')
    BytesIO = context.get('BytesIO')
    DEFAULT_PEAK_MODE = context.get('DEFAULT_PEAK_MODE')
    EXPORT_PREFERENCE_ADVANCED = context.get('EXPORT_PREFERENCE_ADVANCED')
    EXPORT_PREFERENCE_STANDARD = context.get('EXPORT_PREFERENCE_STANDARD')
    PEAK_MODE_OPTIONS = context.get('PEAK_MODE_OPTIONS')
    PM_WINDOW = context.get('PM_WINDOW')
    Path = context.get('Path')
    STANDARD_REPORT_EXPORT_MODE = context.get('STANDARD_REPORT_EXPORT_MODE')
    WORKFLOW_BATCH_MODE = context.get('WORKFLOW_BATCH_MODE')
    WorkflowReadiness = context.get('WorkflowReadiness')
    _apply_standard_batch_export_mode = context.get('_apply_standard_batch_export_mode')
    _batch_analysis_signature = context.get('_batch_analysis_signature')
    _batch_export_mode_options = context.get('_batch_export_mode_options')
    _batch_export_signature = context.get('_batch_export_signature')
    _batch_items_from_uploads = context.get('_batch_items_from_uploads')
    _batch_qc_preview_display_frame = context.get('_batch_qc_preview_display_frame')
    _batch_qc_rows_for_ui = context.get('_batch_qc_rows_for_ui')
    _batch_status_display_frame = context.get('_batch_status_display_frame')
    _batch_status_frame = context.get('_batch_status_frame')
    _batch_summary_counts = context.get('_batch_summary_counts')
    _batch_upload_signature = context.get('_batch_upload_signature')
    _batch_workflow_revisions = context.get('_batch_workflow_revisions')
    _bulk_accept_clean_batch_files = context.get('_bulk_accept_clean_batch_files')
    _coerce_export_mode = context.get('_coerce_export_mode')
    _coerce_setup_time = context.get('_coerce_setup_time')
    _confirm_batch_peak_review = context.get('_confirm_batch_peak_review')
    _default_batch_export_mode = context.get('_default_batch_export_mode')
    _display_status_label = context.get('_display_status_label')
    _exclude_batch_file = context.get('_exclude_batch_file')
    _existing_columns = context.get('_existing_columns')
    _flash_and_rerun = context.get('_flash_and_rerun')
    _format_display_columns = context.get('_format_display_columns')
    _is_v2_scheme = context.get('_is_v2_scheme')
    _mapping_editor_frame = context.get('_mapping_editor_frame')
    _mapping_preset_rows_frame = context.get('_mapping_preset_rows_frame')
    _mark_batch_export_stale_if_inputs_changed = context.get('_mark_batch_export_stale_if_inputs_changed')
    _mark_batch_export_stale_now = context.get('_mark_batch_export_stale_now')
    _mark_batch_stale_if_inputs_changed = context.get('_mark_batch_stale_if_inputs_changed')
    _pce_override_summary = context.get('_pce_override_summary')
    _qc_display_frame = context.get('_qc_display_frame')
    _record_workflow_state = context.get('_record_workflow_state')
    _render_action_hint = context.get('_render_action_hint')
    _render_alert = context.get('_render_alert')
    _render_empty_state = context.get('_render_empty_state')
    _render_hourly_pcu_line_chart = context.get('_render_hourly_pcu_line_chart')
    _render_mapping_scheme_status = context.get('_render_mapping_scheme_status')
    _render_metric_strip = context.get('_render_metric_strip')
    _render_pce_factor_editor = context.get('_render_pce_factor_editor')
    _render_peak_card = context.get('_render_peak_card')
    _render_readiness_checklist = context.get('_render_readiness_checklist')
    _render_section_header = context.get('_render_section_header')
    _render_status_chip = context.get('_render_status_chip')
    _restore_batch_file = context.get('_restore_batch_file')
    _stable_batch_confirmed_peaks = context.get('_stable_batch_confirmed_peaks')
    _standard_report_decision = context.get('_standard_report_decision')
    _sync_batch_analysis_metadata_from_state = context.get('_sync_batch_analysis_metadata_from_state')
    _sync_batch_dispositions_to_analysis = context.get('_sync_batch_dispositions_to_analysis')
    _sync_batch_workflow_from_state = context.get('_sync_batch_workflow_from_state')
    _sync_workflow_after_pce_editor = context.get('_sync_workflow_after_pce_editor')
    _time_from_text = context.get('_time_from_text')
    _use_excel_native_charts_for_export = context.get('_use_excel_native_charts_for_export')
    _use_template_layout_for_export = context.get('_use_template_layout_for_export')
    accepted_names = context.get('accepted_names')
    active_tab = context.get('active_tab')
    am_default = context.get('am_default')
    am_peak_window_end = context.get('am_peak_window_end')
    am_peak_window_start = context.get('am_peak_window_start')
    analyze_batch = context.get('analyze_batch')
    analyzed_batch_revisions = context.get('analyzed_batch_revisions')
    application_analyze_batch = context.get('application_analyze_batch')
    application_export_batch_reviewed = context.get('application_export_batch_reviewed')
    apply_batch_export_mode_change = context.get('apply_batch_export_mode_change')
    apply_mapping_preset_to_detected_sheets = context.get('apply_mapping_preset_to_detected_sheets')
    apply_result = context.get('apply_result')
    batch_analysis = context.get('batch_analysis')
    batch_confirmed_peaks = context.get('batch_confirmed_peaks')
    batch_draft_peaks = context.get('batch_draft_peaks')
    batch_export_left = context.get('batch_export_left')
    batch_export_mode = context.get('batch_export_mode')
    batch_export_options = context.get('batch_export_options')
    batch_export_preference = context.get('batch_export_preference')
    batch_export_right = context.get('batch_export_right')
    batch_export_signature = context.get('batch_export_signature')
    batch_export_stale = context.get('batch_export_stale')
    batch_inputs_ready = context.get('batch_inputs_ready')
    batch_left = context.get('batch_left')
    batch_mapping_scheme = context.get('batch_mapping_scheme')
    batch_metadata_rows = context.get('batch_metadata_rows')
    batch_package_filename = context.get('batch_package_filename')
    batch_preset_name = context.get('batch_preset_name')
    batch_preset_signature = context.get('batch_preset_signature')
    batch_process_block_reason = context.get('batch_process_block_reason')
    batch_processing_block_reason = context.get('batch_processing_block_reason')
    batch_qc = context.get('batch_qc')
    batch_qc_display = context.get('batch_qc_display')
    batch_ready = context.get('batch_ready')
    batch_result = context.get('batch_result')
    batch_review_state = context.get('batch_review_state')
    batch_review_version = context.get('batch_review_version')
    batch_right = context.get('batch_right')
    batch_selected_file_preview = context.get('batch_selected_file_preview')
    batch_signature = context.get('batch_signature')
    batch_stale = context.get('batch_stale')
    batch_standard_decision = context.get('batch_standard_decision')
    batch_uploads = context.get('batch_uploads')
    batch_zip_contents_preview = context.get('batch_zip_contents_preview')
    batch_zip_generation_block_reason = context.get('batch_zip_generation_block_reason')
    block_reason = context.get('block_reason')
    build_setup_for_processing = context.get('build_setup_for_processing')
    cleaned_metadata = context.get('cleaned_metadata')
    column = context.get('column')
    confirm_review = context.get('confirm_review')
    confirmed_count = context.get('confirmed_count')
    counts = context.get('counts')
    detect_raw_direction_sheet_names = context.get('detect_raw_direction_sheet_names')
    detected = context.get('detected')
    direction_cols = context.get('direction_cols')
    download_buffer = context.get('download_buffer')
    draft = context.get('draft')
    draft_am = context.get('draft_am')
    draft_pm = context.get('draft_pm')
    edited_metadata = context.get('edited_metadata')
    eligible_clean_batch_items = context.get('eligible_clean_batch_items')
    eligible_count = context.get('eligible_count')
    exc = context.get('exc')
    excel_com_status = context.get('excel_com_status')
    excluded_count = context.get('excluded_count')
    exclusion_reason = context.get('exclusion_reason')
    export_mode_ready = context.get('export_mode_ready')
    extra_sheets = context.get('extra_sheets')
    failed = context.get('failed')
    failed_columns = context.get('failed_columns')
    failed_display = context.get('failed_display')
    file = context.get('file')
    folder = context.get('folder')
    format_count = context.get('format_count')
    format_pcu = context.get('format_pcu')
    generate_batch = context.get('generate_batch')
    generate_disabled = context.get('generate_disabled')
    has_overrides = context.get('has_overrides')
    item = context.get('item')
    items = context.get('items')
    label = context.get('label')
    label_by_folder = context.get('label_by_folder')
    loaded_batch_preset = context.get('loaded_batch_preset')
    mapping_ready = context.get('mapping_ready')
    metadata_version = context.get('metadata_version')
    missing_sheets = context.get('missing_sheets')
    no_successful_files = context.get('no_successful_files')
    option_labels = context.get('option_labels')
    output_stems_valid = context.get('output_stems_valid')
    override_text = context.get('override_text')
    pce_ready = context.get('pce_ready')
    pd = context.get('pd')
    peak_cols = context.get('peak_cols')
    peak_mode = context.get('peak_mode')
    peak_mode_default = context.get('peak_mode_default')
    peak_windows = context.get('peak_windows')
    peaks_ready = context.get('peaks_ready')
    period_cols = context.get('period_cols')
    pm_default = context.get('pm_default')
    pm_peak_window_end = context.get('pm_peak_window_end')
    pm_peak_window_start = context.get('pm_peak_window_start')
    preset_basic = context.get('preset_basic')
    preset_code_column = context.get('preset_code_column')
    preset_codes = context.get('preset_codes')
    preset_duplicate_count = context.get('preset_duplicate_count')
    preset_include = context.get('preset_include')
    preset_included = context.get('preset_included')
    preset_rows = context.get('preset_rows')
    preview = context.get('preview')
    preview_columns = context.get('preview_columns')
    previous_batch_export_mode = context.get('previous_batch_export_mode')
    qc_summary = context.get('qc_summary')
    qc_summary_columns = context.get('qc_summary_columns')
    queue_columns = context.get('queue_columns')
    queue_frame = context.get('queue_frame')
    review_filter = context.get('review_filter')
    review_labels = context.get('review_labels')
    reviewed_peak_values_complete = context.get('reviewed_peak_values_complete')
    road_cols = context.get('road_cols')
    row = context.get('row')
    safe_output_stem = context.get('safe_output_stem')
    selected_am = context.get('selected_am')
    selected_batch_export_mode = context.get('selected_batch_export_mode')
    selected_folder = context.get('selected_folder')
    selected_item = context.get('selected_item')
    selected_label = context.get('selected_label')
    selected_pce_factors = context.get('selected_pce_factors')
    selected_pm = context.get('selected_pm')
    selected_review_label = context.get('selected_review_label')
    set_active_tab = context.get('set_active_tab')
    setup = context.get('setup')
    setup_cols = context.get('setup_cols')
    standard_batch_mode_changed = context.get('standard_batch_mode_changed')
    status_columns = context.get('status_columns')
    status_display = context.get('status_display')
    status_frame = context.get('status_frame')
    status_items = context.get('status_items')
    status_rows = context.get('status_rows')
    stored = context.get('stored')
    successful_count = context.get('successful_count')
    successful_items = context.get('successful_items')
    survey_period = context.get('survey_period')
    uploaded_ready = context.get('uploaded_ready')
    v2_batch_template_mode_blocked = context.get('v2_batch_template_mode_blocked')
    value = context.get('value')

    batch_export_mode = st.session_state.get("tmc_batch_export_mode", BATCH_SAFE_PNG_EXPORT_LABEL)
    uploaded_ready = bool(batch_uploads)
    mapping_ready = loaded_batch_preset is not None
    pce_ready = bool(selected_pce_factors)
    batch_ready = batch_inputs_ready(
        uploaded_workbook_count=len(batch_uploads or []),
        mapping_available=mapping_ready,
        pce_factors_ready=pce_ready,
        movement_code_scheme=batch_mapping_scheme,
    )
    batch_process_block_reason = batch_processing_block_reason(batch_mapping_scheme)
    batch_signature = _batch_analysis_signature(
        uploads_signature=_batch_upload_signature(batch_uploads),
        preset_signature=batch_preset_signature,
        pce_factors=selected_pce_factors,
        peak_mode=peak_mode,
        peak_windows=peak_windows,
    )
    batch_stale = _mark_batch_stale_if_inputs_changed(batch_signature)
    batch_analysis = st.session_state.get("tmc_batch_analysis_result")
    batch_result = st.session_state.get("tmc_batch_export_result")
    batch_export_signature = _batch_export_signature(
        metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
        shared_setup=setup,
        export_mode=batch_export_mode,
        confirmed_peaks=st.session_state.get(BATCH_CONFIRMED_PEAKS_STATE_KEY) or {},
    )
    batch_export_stale = _mark_batch_export_stale_if_inputs_changed(batch_export_signature)

    if active_tab == "Data":
        _render_section_header("Data Batch", "Set shared metadata and inspect the uploaded workbook inventory.")
        if batch_stale:
            _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")

        batch_left, batch_right = st.columns([1.1, 1])
        with batch_left:
            with st.container(border=True):
                _render_section_header("Data", "Shared project setup and Batch source inventory.")
                _render_readiness_checklist(
                    [
                        ("Mapping Preset", mapping_ready, "Configure it in Mapping"),
                        ("Shared PCE factors", pce_ready, "Configure them in Analyze"),
                        ("Per-file Review", bool(batch_analysis), "Review results after Analyze"),
                        ("ไม่รวม raw Excel ใน ZIP", True, "แพ็กเกจส่งออกมีเฉพาะรายงานและไฟล์ประกอบ"),
                    ]
                )
                setup_cols = st.columns(3)
                setup_cols[0].text_input("ชื่อจุดนับ / TMC title", key="tmc_title_input")
                setup_cols[1].text_input("จุดสำรวจ", key="survey_point_input")
                setup_cols[2].text_input("ช่วงเวลาสำรวจ", key="survey_period_input")
                direction_cols = st.columns(4)
                direction_cols[0].text_input("ป้ายปลายทางด้านเหนือ", key="north_label_input")
                direction_cols[1].text_input("ป้ายปลายทางด้านใต้", key="south_label_input")
                direction_cols[2].text_input("ป้ายปลายทางด้านตะวันออก", key="east_label_input")
                direction_cols[3].text_input("ป้ายปลายทางด้านตะวันตก", key="west_label_input")
                road_cols = st.columns(4)
                road_cols[0].text_input("ชื่อถนนด้านเหนือ", key="north_road_input")
                road_cols[1].text_input("ชื่อถนนด้านใต้", key="south_road_input")
                road_cols[2].text_input("ชื่อถนนด้านตะวันออก", key="east_road_input")
                road_cols[3].text_input("ชื่อถนนด้านตะวันตก", key="west_road_input")

        with batch_right:
            with st.container(border=True):
                _render_section_header("ไฟล์ที่อัปโหลด", "ตรวจจำนวนไฟล์และชื่อไฟล์ก่อนวิเคราะห์ Batch")
                if not batch_metadata_rows:
                    _render_action_hint("Upload Batch workbooks in the Data stage above.")
                else:
                    _render_metric_strip(
                        [
                            ("จำนวนไฟล์", f"{len(batch_metadata_rows):,}", "ไฟล์", "พร้อมตั้งค่ารายไฟล์"),
                        ],
                        columns=1,
                    )
                    st.dataframe(pd.DataFrame({"file_name": [row["file_name"] for row in batch_metadata_rows]}), width="stretch")

            with st.container(border=True):
                _render_section_header("ข้อมูลรายไฟล์", "survey_date_text และ output_stem จะใช้ในรายงานและ ZIP")
                if not batch_metadata_rows:
                    _render_empty_state("ยังไม่มีข้อมูลรายไฟล์", "อัปโหลดไฟล์ Batch เพื่อสร้างตารางตั้งค่า")
                else:
                    metadata_version = int(st.session_state.get("tmc_batch_file_metadata_editor_version", 0) or 0)
                    edited_metadata = st.data_editor(
                        pd.DataFrame(batch_metadata_rows),
                        key=f"tmc_batch_file_metadata_editor_{metadata_version}",
                        hide_index=True,
                        width="stretch",
                        disabled=["file_name"],
                        column_config={
                            "file_name": st.column_config.TextColumn("ชื่อไฟล์ต้นทาง"),
                            "survey_date_text": st.column_config.TextColumn("วันที่สำรวจ"),
                            "output_stem": st.column_config.TextColumn("ชื่อไฟล์ส่งออก"),
                            "notes": st.column_config.TextColumn("หมายเหตุ"),
                        },
                    )
                    cleaned_metadata = []
                    for row in edited_metadata.to_dict("records"):
                        cleaned_metadata.append(
                            {
                                "file_name": Path(str(row.get("file_name", ""))).name,
                                "survey_date_text": str(row.get("survey_date_text", "") or ""),
                                "output_stem": safe_output_stem(str(row.get("output_stem", "") or row.get("file_name", ""))),
                                "notes": str(row.get("notes", "") or ""),
                            }
                        )
                    if cleaned_metadata != st.session_state.get("tmc_batch_file_metadata_table"):
                        st.session_state["tmc_batch_file_metadata_table"] = cleaned_metadata
                        _sync_batch_analysis_metadata_from_state()
                        _mark_batch_export_stale_now()
                        _sync_batch_workflow_from_state(
                            batch_uploads=batch_uploads,
                            mapping_preset=loaded_batch_preset,
                            movement_code_scheme=batch_mapping_scheme,
                            metadata_rows=cleaned_metadata,
                            export_mode=batch_export_mode,
                        )

    if active_tab == "Mapping":
        _render_section_header(
            "Mapping Batch",
            "Apply and validate one Mapping Preset for every uploaded workbook.",
        )
        preset_rows = _mapping_preset_rows_frame(loaded_batch_preset)
        preset_code_column = "output_movement_code" if "output_movement_code" in preset_rows else "movement_code"
        preset_included = 0
        preset_duplicate_count = 0
        if not preset_rows.empty:
            preset_include = preset_rows["include_in_report"].fillna(True).astype(bool) if "include_in_report" in preset_rows else pd.Series(True, index=preset_rows.index)
            preset_codes = preset_rows[preset_code_column].fillna("").astype(str).str.strip() if preset_code_column in preset_rows else pd.Series("", index=preset_rows.index)
            preset_included = int((preset_include & (preset_codes != "")).sum())
            preset_duplicate_count = int((preset_codes[preset_include & (preset_codes != "")].value_counts() > 1).sum())

        _render_metric_strip(
            [
                ("Mapping Preset", "พร้อมใช้งาน" if loaded_batch_preset else "ยังไม่พร้อม", "", batch_preset_name if loaded_batch_preset else "Open the preset above", "พร้อมใช้งาน" if loaded_batch_preset else "ต้องตรวจสอบ"),
                ("แถว Mapping", len(preset_rows), "แถว", "shared preset", "พร้อม" if loaded_batch_preset else "ต้องตรวจสอบ"),
                ("Movement ที่ใช้", preset_included, "แถว", "include_in_report", "พร้อม" if preset_included else "ต้องตรวจสอบ"),
                ("รวมหลาย source", preset_duplicate_count, "movement", "อนุญาตสำหรับ aggregation", "ข้อมูล" if preset_duplicate_count else "พร้อม"),
            ],
            columns=4,
        )
        if loaded_batch_preset:
            _render_mapping_scheme_status(batch_mapping_scheme)

        if not batch_uploads:
            _render_action_hint("Upload Batch workbooks in the Data stage before checking Mapping.")
        if not loaded_batch_preset:
            _render_action_hint("เปิด Mapping Preset เพื่อใช้กับไฟล์ Batch")

        _render_section_header("Mapping readiness", "Resolve Mapping blockers before moving to Analyze.")
        _render_readiness_checklist(
            [
                ("อัปโหลดไฟล์ Batch", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่มีไฟล์"),
                ("Mapping Preset พร้อมใช้งาน", mapping_ready, batch_preset_name if mapping_ready else "ยังไม่เปิด Preset"),
                ("ค่า PCE พร้อมใช้งาน", pce_ready, "พร้อมใช้งาน" if pce_ready else "Configure them in Analyze"),
                ("Metadata รายไฟล์พร้อมใช้งาน", bool(batch_metadata_rows), f"{len(batch_metadata_rows):,} แถว" if batch_metadata_rows else "Configure them in Data"),
            ]
        )
        if mapping_ready:
            if batch_process_block_reason:
                _render_alert(batch_process_block_reason, "warning")
            else:
                _render_action_hint("When Mapping is ready, go to Analyze to run Analyze Batch.")

        with st.expander("สถานะ Sheet matching รายไฟล์", expanded=False):
            _render_action_hint("ตรวจว่า Sheet ในแต่ละไฟล์ตรงกับ Mapping Preset แค่ไหน")
            if not batch_uploads or not loaded_batch_preset:
                _render_empty_state("ยังตรวจ Sheet matching ไม่ได้", "อัปโหลดไฟล์ Batch และเปิด Mapping Preset ก่อน")
            else:
                status_rows = []
                for file in batch_uploads:
                    try:
                        detected = detect_raw_direction_sheet_names(BytesIO(file.getvalue()))
                        apply_result = apply_mapping_preset_to_detected_sheets(loaded_batch_preset, detected)
                        missing_sheets = ", ".join(apply_result.missing_detected_sheets)
                        extra_sheets = ", ".join(apply_result.extra_preset_sheets)
                        status_rows.append(
                            {
                                "file_name": Path(file.name).name,
                                "detected_sheets": len(detected),
                                "matched_sheets": apply_result.matched_sheet_count,
                                "source_coverage": (
                                    f"{apply_result.matched_sheet_count}/{len(detected)}"
                                    if detected
                                    else "0/0"
                                ),
                                "missing_detected_sheets": apply_result.missing_detected_sheet_count,
                                "preset_rows_not_found": apply_result.extra_preset_row_count,
                                "mapping_status": "พร้อม" if apply_result.missing_detected_sheet_count == 0 else "ต้องตรวจสอบ",
                                "notes": missing_sheets or extra_sheets or "",
                            }
                        )
                    except Exception as exc:
                        status_rows.append(
                            {
                                "file_name": Path(file.name).name,
                                "detected_sheets": 0,
                                "matched_sheets": 0,
                                "source_coverage": "0/0",
                                "missing_detected_sheets": 0,
                                "preset_rows_not_found": 0,
                                "mapping_status": "อ่านไฟล์ไม่สำเร็จ",
                                "notes": str(exc),
                            }
                        )
                st.dataframe(pd.DataFrame(status_rows), width="stretch", hide_index=True)

        if loaded_batch_preset and not preset_rows.empty:
            with st.expander("ตัวอย่าง Mapping Preset", expanded=False):
                preset_basic = _mapping_editor_frame(preset_rows, "Basic", batch_mapping_scheme)
                preview_columns = [
                    column
                    for column in [
                        "raw_sheet",
                        "raw_direction",
                        "raw_movement_label",
                        "physical_approach",
                        "physical_movement",
                        "include_in_report",
                        "include_in_peak",
                        "derived_code",
                        "status",
                    ]
                    if column in preset_basic.columns
                ]
                st.dataframe(
                    preset_basic[preview_columns].head(50) if preview_columns else preset_basic.head(50),
                    width="stretch",
                    hide_index=True,
                )

    if active_tab == "Analyze":
        _render_section_header("Analyze Batch", "Configure shared analysis settings and run Analyze Batch.")
        with st.container(border=True):
            _render_section_header("Analysis settings", "These settings apply to every uploaded Batch workbook.")
            survey_period = st.text_input("Survey period", key="survey_period_input")
            peak_mode_default = str(st.session_state.get("peak_mode_select") or DEFAULT_PEAK_MODE)
            if peak_mode_default not in PEAK_MODE_OPTIONS:
                peak_mode_default = DEFAULT_PEAK_MODE
                st.session_state["peak_mode_select"] = peak_mode_default
            peak_mode = st.selectbox(
                "Peak calculation mode",
                options=PEAK_MODE_OPTIONS,
                index=PEAK_MODE_OPTIONS.index(peak_mode_default),
                key="peak_mode_select",
            )
            period_cols = st.columns(4)
            am_peak_window_start = period_cols[0].time_input(
                "AM window start",
                value=_coerce_setup_time(st.session_state["am_peak_window_start_input"], _time_from_text(AM_WINDOW[0])),
                step=900,
                key="am_peak_window_start_input",
            )
            am_peak_window_end = period_cols[1].time_input(
                "AM window end",
                value=_coerce_setup_time(st.session_state["am_peak_window_end_input"], _time_from_text(AM_WINDOW[1])),
                step=900,
                key="am_peak_window_end_input",
            )
            pm_peak_window_start = period_cols[2].time_input(
                "PM window start",
                value=_coerce_setup_time(st.session_state["pm_peak_window_start_input"], _time_from_text(PM_WINDOW[0])),
                step=900,
                key="pm_peak_window_start_input",
            )
            pm_peak_window_end = period_cols[3].time_input(
                "PM window end",
                value=_coerce_setup_time(st.session_state["pm_peak_window_end_input"], _time_from_text(PM_WINDOW[1])),
                step=900,
                key="pm_peak_window_end_input",
            )
            selected_pce_factors = _render_pce_factor_editor()
            _sync_workflow_after_pce_editor(
                is_single_file_mode=False,
                export_mode=str(st.session_state.get("tmc_batch_export_mode") or BATCH_SAFE_PNG_EXPORT_LABEL),
                batch_uploads=batch_uploads,
                mapping_preset=loaded_batch_preset,
                movement_code_scheme=batch_mapping_scheme,
                metadata_rows=batch_metadata_rows,
            )
            has_overrides, override_text = _pce_override_summary(selected_pce_factors)
            _render_status_chip("User-adjusted PCE" if has_overrides else "Default PCE factors", "warning" if has_overrides else "success")
            if has_overrides:
                st.caption(override_text)
        setup = build_setup_for_processing("")
        setup["movement_code_scheme"] = batch_mapping_scheme
        peak_mode = str(setup.get("peak_mode", DEFAULT_PEAK_MODE) or DEFAULT_PEAK_MODE)
        peak_windows = {
            "AM": (setup["am_peak_window_start"], setup["am_peak_window_end"]),
            "PM": (setup["pm_peak_window_start"], setup["pm_peak_window_end"]),
        }
        pce_ready = bool(selected_pce_factors)
        batch_ready = batch_inputs_ready(
            uploaded_workbook_count=len(batch_uploads or []),
            mapping_available=mapping_ready,
            pce_factors_ready=pce_ready,
            movement_code_scheme=batch_mapping_scheme,
        )
        if batch_stale:
            _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
        if not batch_ready:
            _render_readiness_checklist(
                [
                    ("อัปโหลดไฟล์ Batch", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่ได้อัปโหลดไฟล์ TMC Excel"),
                     ("Mapping Preset", mapping_ready, "พร้อมใช้งาน" if mapping_ready else "Configure it in Mapping"),
                     ("Shared PCE factors", pce_ready, "พร้อมใช้งาน" if pce_ready else "Configure them above"),
                ]
            )
            if batch_process_block_reason:
                _render_alert(batch_process_block_reason, "warning")
            else:
                _render_action_hint("เตรียมไฟล์ Mapping Preset และค่า PCE ให้พร้อมก่อนวิเคราะห์ Batch")
        elif not batch_analysis or batch_stale:
            _render_action_hint("วิเคราะห์ Batch เพื่อสร้างรายการตรวจ Peak รายไฟล์")
        analyze_batch = st.button("วิเคราะห์ Batch", type="primary", disabled=not batch_ready, key="analyze_batch_processing")
        if analyze_batch:
            items = _batch_items_from_uploads(batch_uploads)
            with st.spinner("กำลังวิเคราะห์ Batch..."):
                batch_analysis = application_analyze_batch(
                    items,
                    mapping_preset=loaded_batch_preset,
                    setup={**setup, "movement_code_scheme": batch_mapping_scheme},
                    pce_factors=selected_pce_factors,
                    peak_mode=peak_mode,
                    peak_windows=peak_windows,
                    mapping_preset_name=batch_preset_name,
                )
            st.session_state["tmc_batch_analysis_result"] = batch_analysis
            st.session_state["tmc_batch_preset_name"] = batch_preset_name
            st.session_state["tmc_batch_input_signature"] = st.session_state.get("tmc_batch_current_input_signature")
            st.session_state["tmc_batch_stale"] = False
            st.session_state["tmc_batch_export_stale"] = False
            st.session_state.pop("tmc_batch_export_result", None)
            st.session_state["tmc_batch_review_version"] = int(st.session_state.get("tmc_batch_review_version", 0) or 0) + 1
            st.session_state[BATCH_CONFIRMED_PEAKS_STATE_KEY] = {}
            st.session_state[BATCH_DRAFT_PEAKS_STATE_KEY] = {}
            st.session_state[BATCH_CONFIRMED_PEAK_SOURCE_STATE_KEY] = {}
            st.session_state[BATCH_DISPOSITION_STATE_KEY] = {}
            analyzed_batch_revisions = _batch_workflow_revisions(
                uploads=batch_uploads,
                mapping_preset=loaded_batch_preset,
                pce_factors=selected_pce_factors,
                peak_mode=peak_mode,
                peak_windows=peak_windows,
                movement_code_scheme=batch_mapping_scheme,
                metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                setup=setup,
                export_mode=batch_export_mode,
                confirmed_peaks=_stable_batch_confirmed_peaks(
                    st.session_state.get(BATCH_CONFIRMED_PEAKS_STATE_KEY) or {},
                    batch_analysis,
                ),
                analysis_present=True,
            )
            _record_workflow_state(
                WORKFLOW_BATCH_MODE,
                analyzed_batch_revisions,
                WorkflowReadiness(
                    source=bool(batch_uploads),
                    mapping=loaded_batch_preset is not None,
                    analysis=True,
                    review=bool(batch_analysis.successful_items) and reviewed_peak_values_complete(batch_analysis),
                    export=False,
                ),
            )
            set_active_tab("Review")
            _flash_and_rerun("วิเคราะห์ Batch เสร็จแล้ว กรุณาตรวจสอบช่วงเร่งด่วนก่อนสร้าง ZIP")

    if active_tab == "Review":
        if batch_analysis:
            batch_confirmed_peaks = st.session_state.setdefault(BATCH_CONFIRMED_PEAKS_STATE_KEY, {})
            batch_review_version = int(st.session_state.get("tmc_batch_review_version", 0) or 0)
            for item in batch_analysis.successful_items:
                stored = batch_confirmed_peaks.get(item.folder_name, {})
                item.confirmed_AM_peak = stored.get("AM", item.confirmed_AM_peak)
                item.confirmed_PM_peak = stored.get("PM", item.confirmed_PM_peak)
            successful_items = batch_analysis.successful_items
            _sync_batch_dispositions_to_analysis()
            confirmed_count = sum(1 for item in successful_items if batch_review_state(item) == "confirmed")
            excluded_count = sum(1 for item in successful_items if batch_review_state(item) == "excluded")
            successful_count = len(successful_items)
            status_items = [
                ("ไฟล์ทั้งหมด", f"{len(batch_analysis.items):,}", "ไฟล์", ""),
                ("กำหนด Peak", f"{confirmed_count:,}/{successful_count:,}", "ไฟล์", "ไฟล์ที่วิเคราะห์สำเร็จ"),
                ("ไฟล์ไม่สำเร็จ", f"{sum(1 for item in batch_analysis.items if item.status == 'failed'):,}", "ไฟล์", "ไม่ต้องกำหนด Peak"),
            ]
            _render_metric_strip(status_items, columns=3)
            if excluded_count:
                _render_alert(f"{excluded_count} successful file(s) intentionally excluded from Batch export.", "info")
            if successful_count and confirmed_count + excluded_count == successful_count:
                _render_alert("กำหนด Peak ครบแล้ว พร้อมส่งออก Batch", "success")
            elif successful_count:
                _render_alert("ยังมีไฟล์ที่ต้องกำหนด Peak", "warning")
            review_filter = st.radio(
                "Batch review queue",
                options=["Needs review", "All files"],
                index=0,
                horizontal=True,
                key="tmc_batch_review_filter",
                help="Needs review keeps unresolved and failed files visible while confirmed files stay out of the default queue.",
            )
            queue_frame = _batch_status_frame(batch_analysis)
            if review_filter == "Needs review" and not queue_frame.empty:
                queue_frame = queue_frame[queue_frame["review_state"].isin(["needs_review", "failed"])]
            queue_columns = [
                "file_name",
                "survey_date_text",
                "status",
                "review_state",
                "AM suggested",
                "PM suggested",
                "AM confirmed",
                "PM confirmed",
                "QC errors",
                "QC warnings",
                "QC info",
                "disposition_reason",
            ]
            st.dataframe(queue_frame[[column for column in queue_columns if column in queue_frame.columns]], width="stretch", hide_index=True)
            eligible_count = len(eligible_clean_batch_items(batch_analysis))
            if eligible_count:
                st.caption(f"{eligible_count} clean file(s) eligible for explicit bulk acceptance. QC info does not block eligibility.")
            if st.button(
                "Accept suggested Peaks for clean files",
                type="primary",
                disabled=eligible_count == 0,
                key="tmc_batch_bulk_accept_clean",
            ):
                accepted_names = _bulk_accept_clean_batch_files()
                _sync_batch_workflow_from_state(
                    batch_uploads=batch_uploads,
                    mapping_preset=loaded_batch_preset,
                    movement_code_scheme=batch_mapping_scheme,
                    metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                    export_mode=batch_export_mode,
                )
                _flash_and_rerun(f"Accepted suggested Peaks for {len(accepted_names)} clean file(s).")

            if successful_items:
                review_labels = {f"{item.file_name} ({item.survey_date_text or 'no date'})": item.folder_name for item in successful_items}
                selected_folder = st.session_state.get("tmc_batch_selected_review_file") or successful_items[0].folder_name
                label_by_folder = {folder: label for label, folder in review_labels.items()}
                selected_label = label_by_folder.get(selected_folder, next(iter(review_labels)))
                selected_review_label = st.selectbox(
                    "เลือกไฟล์สำหรับตรวจกราฟ",
                    options=list(review_labels),
                    index=list(review_labels).index(selected_label),
                    key="tmc_batch_selected_review_label",
                )
                selected_folder = review_labels[selected_review_label]
                st.session_state["tmc_batch_selected_review_file"] = selected_folder
                selected_item = next(item for item in successful_items if item.folder_name == selected_folder)
                preview = batch_selected_file_preview(selected_item)
                with st.container(border=True):
                    _render_section_header("ไฟล์ที่เลือก", f"{preview['file_name']} · {preview['output_stem']}")
                    _render_metric_strip(
                        [
                            ("วันที่สำรวจ", preview["survey_date_text"] or "-", "", ""),
                            ("สถานะ", preview["status"] or "-", "", "", preview["status"] or "pending"),
                            ("จำนวนรถรวม", format_count(preview["total_vehicles"]), "คัน", ""),
                            ("PCU รวม", format_pcu(preview["total_PCU"]), "PCU", ""),
                        ],
                        columns=4,
                    )
                    _render_metric_strip(
                        [
                            ("QC error", format_count(preview["QC_errors"]), "", ""),
                            ("QC warning", format_count(preview["QC_warnings"]), "", ""),
                            ("QC info", format_count(preview["QC_info"]), "", ""),
                        ],
                        columns=3,
                    )
                _render_section_header("กราฟ PCU รายชั่วโมง", "ใช้ตรวจรูปแบบปริมาณจราจรก่อนกำหนดช่วง Peak")
                if getattr(selected_item, "disposition", "") == "excluded":
                    if st.button("Restore file to Needs review", key=f"restore_batch_file_{selected_item.folder_name}"):
                        _restore_batch_file(selected_item.folder_name)
                        _flash_and_rerun("File restored to Needs review.")
                else:
                    exclusion_reason = st.text_input(
                        "Exclusion reason (required for intentional exclusion)",
                        key=f"batch_exclusion_reason_{selected_item.folder_name}",
                    )
                    if st.button(
                        "Exclude selected file",
                        disabled=not exclusion_reason.strip(),
                        key=f"exclude_batch_file_{selected_item.folder_name}",
                    ):
                        _exclude_batch_file(selected_item.folder_name, exclusion_reason.strip())
                        _flash_and_rerun("File excluded from Batch export.")
                _render_hourly_pcu_line_chart(
                    selected_item.hourly_movement_pcu,
                    preview["confirmed_AM_peak"] or preview["suggested_AM_peak"] or "",
                    preview["confirmed_PM_peak"] or preview["suggested_PM_peak"] or "",
                )
                option_labels = list(dict.fromkeys(selected_item.hourly_period_options or [selected_item.suggested_AM_peak, selected_item.suggested_PM_peak]))
                option_labels = [value for value in option_labels if value]
                _render_section_header("กำหนด Peak ของไฟล์นี้", "ระบบจะใช้ช่วง Peak ที่กำหนดในหน้านี้สำหรับรายงานของไฟล์นี้")
                peak_cols = st.columns(2)
                if option_labels:
                    stored = batch_confirmed_peaks.setdefault(selected_item.folder_name, {"AM": selected_item.confirmed_AM_peak, "PM": selected_item.confirmed_PM_peak})
                    batch_draft_peaks = st.session_state.setdefault(BATCH_DRAFT_PEAKS_STATE_KEY, {})
                    draft = batch_draft_peaks.get(selected_item.folder_name, {})
                    am_default = draft.get("AM") or stored.get("AM") or selected_item.suggested_AM_peak
                    pm_default = draft.get("PM") or stored.get("PM") or selected_item.suggested_PM_peak
                    for value in [am_default, pm_default]:
                        if value and value not in option_labels:
                            option_labels.insert(0, value)
                    option_labels = list(dict.fromkeys(option_labels))
                    with peak_cols[0]:
                        _render_peak_card("AM Peak · ระบบตรวจจับอัตโนมัติ", preview["suggested_AM_peak"] or "", "", "auto_suggested")
                        selected_am = st.selectbox(
                            "ช่วงที่กำหนด AM",
                            options=option_labels,
                            index=option_labels.index(am_default) if am_default in option_labels else 0,
                            key=f"batch_review_am_{batch_review_version}_{selected_item.folder_name}",
                        )
                        draft_am = selected_am
                        selected_am = selected_am if stored.get("AM") == selected_am else ""
                        st.caption("Draft only — confirm this file to apply the selected AM Peak.")
                        _render_status_chip("กำหนดแล้ว" if selected_am else "รอตรวจสอบ", "success" if selected_am else "warning")
                        _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
                    with peak_cols[1]:
                        _render_peak_card("PM Peak · ระบบตรวจจับอัตโนมัติ", preview["suggested_PM_peak"] or "", "", "auto_suggested")
                        selected_pm = st.selectbox(
                            "ช่วงที่กำหนด PM",
                            options=option_labels,
                            index=option_labels.index(pm_default) if pm_default in option_labels else 0,
                            key=f"batch_review_pm_{batch_review_version}_{selected_item.folder_name}",
                        )
                        draft_pm = selected_pm
                        selected_pm = selected_pm if stored.get("PM") == selected_pm else ""
                        st.caption("Draft only — confirm this file to apply the selected PM Peak.")
                        _render_status_chip("กำหนดแล้ว" if selected_pm else "รอตรวจสอบ", "success" if selected_pm else "warning")
                        _render_action_hint("ใช้ช่วงนี้เป็นค่าหลักสำหรับรายงาน")
                    batch_draft_peaks[selected_item.folder_name] = {"AM": draft_am, "PM": draft_pm}
                    confirm_review = st.button(
                        "Confirm Peak Review",
                        type="primary",
                        disabled=not (draft_am and draft_pm),
                        key=f"batch_confirm_peak_review_{batch_review_version}_{selected_item.folder_name}",
                    )
                    if confirm_review:
                        _confirm_batch_peak_review(selected_item.folder_name, draft_am, draft_pm)
                        _sync_batch_workflow_from_state(
                            batch_uploads=batch_uploads,
                            mapping_preset=loaded_batch_preset,
                            movement_code_scheme=batch_mapping_scheme,
                            metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                            export_mode=batch_export_mode,
                        )
                        _flash_and_rerun("Peak Review confirmed for this file.")
                else:
                    _render_alert("ไม่มีช่วงเวลารายชั่วโมงสำหรับกำหนด Peak ของไฟล์นี้", "warning")
            if batch_analysis.has_failures:
                _render_alert("บางไฟล์วิเคราะห์ไม่สำเร็จ ไฟล์เหล่านี้ยังแสดงในตารางและไม่ต้องกำหนด Peak", "warning")
        else:
            _render_empty_state("กรุณาวิเคราะห์ Batch ก่อนตรวจสอบ Peak", "กด วิเคราะห์ Batch เพื่อสร้างตารางตรวจ Peak รายไฟล์")

    if active_tab == "Export":
        _render_section_header("ส่งออก Batch", "สร้าง Batch ZIP พร้อมรายงานรายไฟล์และ batch_summary.xlsx")
        batch_export_options = _batch_export_mode_options(excel_com_status, batch_mapping_scheme)
        previous_batch_export_mode = st.session_state.get("tmc_batch_export_mode", batch_export_mode)
        batch_export_preference = st.radio(
            "Batch report outcome",
            options=[EXPORT_PREFERENCE_STANDARD, EXPORT_PREFERENCE_ADVANCED],
            index=0 if st.session_state.get("tmc_batch_export_preference", EXPORT_PREFERENCE_STANDARD) == EXPORT_PREFERENCE_STANDARD else 1,
            horizontal=True,
            key="tmc_batch_export_preference_control",
            help="Standard report selects native Excel Template when compatible and otherwise uses Safe PNG.",
        )
        st.session_state["tmc_batch_export_preference"] = batch_export_preference
        batch_standard_decision = None
        if batch_export_preference == EXPORT_PREFERENCE_STANDARD:
            batch_standard_decision = _standard_report_decision(
                excel_com_status,
                template_compatible=(not _is_v2_scheme(batch_mapping_scheme)),
            )
            batch_export_mode, standard_batch_mode_changed = _apply_standard_batch_export_mode(
                batch_standard_decision,
                previous_batch_export_mode,
            )
            if standard_batch_mode_changed:
                st.rerun()
            st.info(f"{STANDARD_REPORT_EXPORT_MODE}: {batch_export_mode} selected automatically.")
            if batch_standard_decision.fallback_notice:
                st.warning(batch_standard_decision.fallback_notice)
        else:
            with st.expander("Advanced export options", expanded=True):
                selected_batch_export_mode = st.radio(
                    "Backend",
                    options=batch_export_options,
                    index=batch_export_options.index(
                        _coerce_export_mode(
                            previous_batch_export_mode,
                            batch_export_options,
                            _default_batch_export_mode(excel_com_status, batch_mapping_scheme),
                        )
                    ),
                    key="tmc_batch_export_mode_control",
                    horizontal=True,
                    help="Explicit backend selection is intended for advanced users and diagnostics.",
                )
            if apply_batch_export_mode_change(selected_batch_export_mode, previous_batch_export_mode):
                st.rerun()
            batch_export_mode = selected_batch_export_mode
        batch_analysis = st.session_state.get("tmc_batch_analysis_result")
        batch_result = st.session_state.get("tmc_batch_export_result")
        no_successful_files = not batch_analysis or not batch_analysis.successful_items
        peaks_ready = bool(batch_analysis and batch_analysis.successful_items) and reviewed_peak_values_complete(batch_analysis)
        output_stems_valid = all(str(row.get("output_stem", "")).strip() for row in st.session_state.get("tmc_batch_file_metadata_table") or [])
        v2_batch_template_mode_blocked = (
            _is_v2_scheme(batch_mapping_scheme)
            and batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE)
        )
        export_mode_ready = bool(
            not v2_batch_template_mode_blocked
            and (
                batch_export_mode.startswith(BATCH_SAFE_PNG_EXPORT_MODE)
                or excel_com_status.available
                or not batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE)
            )
        )
        if batch_stale:
            _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
        elif batch_export_stale:
            _render_alert("ข้อมูลส่งออกมีการเปลี่ยนแปลง กรุณาสร้าง Batch ZIP ใหม่", "warning")
        block_reason = batch_zip_generation_block_reason(
            has_successful_files=not no_successful_files,
            peaks_ready=peaks_ready,
            batch_stale=batch_stale,
        )
        generate_disabled = bool(block_reason) or not output_stems_valid or not export_mode_ready
        if not output_stems_valid and not block_reason:
            block_reason = "ตรวจสอบ output_stem ในแท็บตั้งค่า Batch ก่อนสร้าง ZIP"
        if v2_batch_template_mode_blocked and not block_reason:
            block_reason = BATCH_V2_TEMPLATE_MODE_UNSUPPORTED_TH
        if not export_mode_ready and not block_reason:
            block_reason = "โหมดส่งออกยังไม่พร้อม"
        _render_action_hint(block_reason or "พร้อมสร้าง Batch ZIP")
        generate_batch = st.button("Generate Batch ZIP", type="primary", disabled=generate_disabled, key="generate_batch_zip")
        if generate_batch and batch_analysis:
            set_active_tab("Export")
            block_reason = batch_zip_generation_block_reason(
                has_successful_files=bool(batch_analysis.successful_items),
                peaks_ready=bool(batch_analysis.successful_items) and reviewed_peak_values_complete(batch_analysis),
                batch_stale=bool(st.session_state.get("tmc_batch_stale")),
            )
            if block_reason:
                _render_alert(block_reason, "warning")
                st.stop()
            with st.spinner("กำลังสร้าง Batch ZIP..."):
                batch_result = application_export_batch_reviewed(
                    batch_analysis,
                    setup=setup,
                    pce_factors=selected_pce_factors,
                    peak_mode=peak_mode,
                    peak_windows=peak_windows,
                    export_mode=batch_export_mode,
                    use_template_report_layout=_use_template_layout_for_export(batch_export_mode),
                    use_excel_com_native_charts=_use_excel_native_charts_for_export(batch_export_mode, excel_com_status),
            )
            st.session_state["tmc_batch_export_result"] = batch_result
            st.session_state["tmc_batch_export_stale"] = False
            st.session_state["tmc_batch_export_signature"] = _batch_export_signature(
                metadata_rows=st.session_state.get("tmc_batch_file_metadata_table") or [],
                shared_setup=setup,
                export_mode=batch_export_mode,
                confirmed_peaks=st.session_state.get(BATCH_CONFIRMED_PEAKS_STATE_KEY) or {},
            )
            set_active_tab("Export")
            _flash_and_rerun("สร้าง Batch ZIP เสร็จแล้ว")

        batch_export_left, batch_export_right = st.columns([0.95, 1.05])
        with batch_export_left:
            with st.container(border=True):
                _render_section_header("โหมดส่งออก", "สถานะโหมดที่เลือกสำหรับ Batch")
                _render_status_chip(batch_export_mode, "success" if export_mode_ready else "warning")
                if batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE):
                    _render_alert(
                        "Excel Template Mode: เหมาะสำหรับรายงานฉบับใช้งานจริง รักษา Native Chart และรูปแบบ Excel Template เมื่อ Excel COM พร้อมใช้งาน",
                        "info",
                    )
                else:
                    _render_alert(
                        "Safe PNG Export Mode: โหมดสำรอง เหมาะสำหรับตรวจร่างหรือกรณี Excel COM ใช้งานไม่ได้",
                        "info",
                    )
                if batch_export_mode.startswith(BATCH_EXCEL_TEMPLATE_EXPORT_MODE) and len(batch_uploads or []) > 10:
                    _render_alert("มีไฟล์มากกว่า 10 ไฟล์ การสร้างรายงานด้วย Excel Template Mode อาจใช้เวลานานขึ้น", "warning")

        with batch_export_right:
            with st.container(border=True):
                _render_section_header("ความพร้อม Batch", "ตรวจเงื่อนไขก่อนสร้าง Batch ZIP")
                _render_readiness_checklist(
                    [
                        ("อัปโหลดไฟล์แล้ว", uploaded_ready, f"{len(batch_uploads or []):,} ไฟล์" if uploaded_ready else "ยังไม่มีไฟล์"),
                        ("Mapping Preset พร้อม", mapping_ready, "ใช้ Preset เดียวกันทุกไฟล์"),
                        ("Batch Analysis วิเคราะห์แล้ว", bool(batch_analysis and not batch_stale), ""),
                        ("กำหนด Peak ของไฟล์ที่สำเร็จแล้ว", peaks_ready, ""),
                        ("output_stem ใช้งานได้", output_stems_valid, "ใช้เป็นชื่อโฟลเดอร์และชื่อรายงาน"),
                        ("โหมดส่งออกพร้อม", export_mode_ready, ""),
                    ]
                )
        with st.expander("ตัวอย่างไฟล์ใน Batch ZIP", expanded=False):
            st.code(
                "\n".join(
                    [
                        "batch_summary.xlsx",
                        "<output_stem>/report.xlsx",
                        "<output_stem>/export_summary.txt",
                        "<output_stem>/session.tmcproj.json",
                        "<output_stem>/mapping.json",
                        "<output_stem>/charts/",
                    ]
                ),
                language="text",
            )
            st.caption("Batch ZIP ไม่รวม raw input Excel files และไม่รวม local file paths")
        if batch_result:
            _render_section_header("สถานะส่งออกรายไฟล์", "ผลการสร้างรายงานใน Batch ล่าสุด")
            status_display = _batch_status_display_frame(batch_analysis, batch_result)
            status_columns = [
                column
                for column in ["ชื่อไฟล์", "ชื่อส่งออก", "สถานะส่งออก", "โหมดส่งออกที่ใช้", "หมายเหตุ"]
                if column in status_display.columns
            ]
            st.dataframe(status_display[status_columns] if status_columns else status_display, width="stretch", hide_index=True)
            with st.expander("ตัวอย่างไฟล์ใน Batch ZIP ล่าสุด", expanded=False):
                st.code("\n".join(batch_zip_contents_preview(batch_result.summary_rows)), language="text")
                st.caption("Batch ZIP ไม่รวม raw input Excel files และไม่รวม local file paths")
            st.download_button(
                "ดาวน์โหลด Batch ZIP",
                data=download_buffer(batch_result.package_bytes),
                file_name=batch_package_filename(st.session_state.get("tmc_batch_preset_name") or batch_preset_name or "tmc_batch"),
                mime=BATCH_PACKAGE_MIME,
                key="download_batch_zip",
            )
        elif not batch_analysis:
            _render_empty_state("ยังส่งออก Batch ไม่ได้", "วิเคราะห์ Batch และกำหนด Peak ก่อนสร้าง ZIP")
        elif not peaks_ready:
            _render_alert(batch_zip_generation_block_reason(has_successful_files=True, peaks_ready=False, batch_stale=False), "warning")

    if active_tab == "Review":
        _render_section_header(
            "ตรวจสอบข้อมูล Batch",
            "ตรวจสอบสถานะรายไฟล์, QC รวม, และรายละเอียด Batch_QC ก่อนนำผลไปใช้ต่อ",
        )
        batch_analysis = st.session_state.get("tmc_batch_analysis_result")
        batch_result = st.session_state.get("tmc_batch_export_result")
        status_frame = _batch_status_frame(batch_analysis, batch_result)
        if batch_analysis and batch_stale:
            _render_alert("ข้อมูล Batch มีการเปลี่ยนแปลง กรุณาวิเคราะห์ Batch ใหม่", "warning")
        elif batch_analysis and batch_export_stale:
            _render_alert("ข้อมูลส่งออกมีการเปลี่ยนแปลง กรุณาสร้าง Batch ZIP ใหม่", "warning")

        if not batch_analysis:
            _render_empty_state("กรุณาวิเคราะห์ Batch ก่อนตรวจสอบผลรวม", "ผลรวมรายไฟล์และ Batch_QC จะแสดงหลังการวิเคราะห์ Batch")
        else:
            counts = _batch_summary_counts(status_frame)
            _render_metric_strip(
                [
                    ("ไฟล์ทั้งหมด", f"{counts['total_files']:,}", "ไฟล์", ""),
                    ("สำเร็จ", f"{counts['successful_files']:,}", "ไฟล์", "", "success" if counts["successful_files"] else "รอตรวจ"),
                    ("ล้มเหลว", f"{counts['failed_files']:,}", "ไฟล์", "", "failed" if counts["failed_files"] else "พร้อม"),
                    ("ถูกยกเว้น", f"{counts['excluded_files']:,}", "ไฟล์", "", "info" if counts["excluded_files"] else "พร้อม"),
                    ("QC ผิดพลาด", f"{counts['QC_errors']:,}", "รายการ", "", "error" if counts["QC_errors"] else "พร้อม"),
                    ("QC เตือน", f"{counts['QC_warnings']:,}", "รายการ", "", "warning" if counts["QC_warnings"] else "พร้อม"),
                    ("QC ข้อมูล", f"{counts['QC_info']:,}", "รายการ", "", "info" if counts["QC_info"] else "พร้อม"),
                ],
                columns=7,
            )

            _render_section_header("สถานะรายไฟล์", "ตารางตรวจสอบผลวิเคราะห์และความพร้อมส่งออกของแต่ละไฟล์")
            status_display = _batch_status_display_frame(batch_analysis, batch_result)
            status_columns = _existing_columns(
                status_display,
                [
                    "ชื่อไฟล์",
                    "วันที่สำรวจ",
                    "ชื่อส่งออก",
                    "สถานะ",
                    "AM กำหนดแล้ว",
                    "PM กำหนดแล้ว",
                    "PCU รวม",
                    "QC ผิดพลาด",
                    "QC เตือน",
                    "QC ข้อมูล",
                    "สถานะส่งออก",
                    "หมายเหตุ",
                ],
            )
            st.dataframe(status_display[status_columns] if status_columns else status_display, width="stretch", hide_index=True)

            failed = status_frame[status_frame["status"].astype(str) == "failed"]
            if not failed.empty:
                _render_alert("พบไฟล์ที่วิเคราะห์หรือส่งออกไม่สำเร็จ ไฟล์เหล่านี้ไม่ต้องกำหนด Peak", "warning")
                failed_columns = _existing_columns(failed, ["file_name", "output_stem", "status", "notes"])
                failed_display = failed[failed_columns] if failed_columns else failed
                if "status" in failed_display.columns:
                    failed_display = failed_display.copy()
                    failed_display["status"] = failed_display["status"].map(_display_status_label)
                st.dataframe(failed_display, width="stretch", hide_index=True)

            batch_qc = _batch_qc_rows_for_ui(batch_analysis, batch_result)
            if not batch_qc.empty:
                _render_section_header("Batch_QC", "ตัวอย่างรายการ QC รวมสำหรับตรวจสอบก่อนใช้ผลต่อ")
                batch_qc_display = _batch_qc_preview_display_frame(batch_qc)
                st.dataframe(batch_qc_display.head(25), width="stretch", hide_index=True)
                with st.expander("Batch_QC รายละเอียดทั้งหมด", expanded=False):
                    st.dataframe(_qc_display_frame(batch_qc, thai_labels=True), width="stretch", hide_index=True)
            else:
                _render_alert("ยังไม่มีรายการ Batch_QC สำหรับไฟล์ที่วิเคราะห์สำเร็จ", "success")

            if batch_result:
                _render_section_header("ร่องรอย Batch ZIP", "ตรวจสอบองค์ประกอบหลักของแพ็กเกจส่งออกล่าสุด")
                _render_readiness_checklist(
                    [
                        ("สร้าง batch_summary.xlsx แล้ว", True, ""),
                        ("มี Sheet Batch_Summary", True, ""),
                        ("มี Sheet Batch_QC", True, ""),
                        ("ไม่รวม raw input Excel", True, ""),
                    ]
                )

            with st.expander("สรุป QC รายไฟล์", expanded=False):
                qc_summary_columns = _existing_columns(status_frame, ["file_name", "QC errors", "QC warnings", "QC info"])
                qc_summary = status_frame[qc_summary_columns] if qc_summary_columns else status_frame
                st.dataframe(_format_display_columns(qc_summary), width="stretch", hide_index=True)
            with st.expander("รายละเอียดวิเคราะห์ Batch", expanded=False):
                if batch_analysis:
                    st.write(
                        [
                            {
                                "file_name": item.file_name,
                                "detected_sheets": item.detected_sheets,
                                "matched_sheet_count": item.matched_sheet_count,
                                "missing_detected_sheet_count": item.missing_detected_sheet_count,
                                "extra_preset_row_count": item.extra_preset_row_count,
                                "notes": item.notes,
                            }
                            for item in batch_analysis.items
                        ]
                    )
